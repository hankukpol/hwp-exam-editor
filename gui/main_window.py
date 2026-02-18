import os
import csv
import json
import shutil
import subprocess
import sys
import tempfile
import time
import uuid
from pathlib import Path
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLabel, QPushButton, QFileDialog, QProgressBar,
                             QFrame, QAction, QMessageBox, QApplication)
from PyQt5.QtCore import Qt, pyqtSignal, QThread, QTimer
from PyQt5.QtGui import QIcon

from core.config_manager import ConfigManager
from core.exceptions import ProcessingError
from core.models import ExamDocument
from core.service import ExamProcessingService

from .preview_window import PreviewWindow
from .settings_window import SettingsWindow
from .styles import APP_STYLE, apply_shadow


def _document_to_payload(document: ExamDocument) -> dict:
    return {
        "file_type": document.file_type,
        "subject": document.subject,
        "total_count": document.total_count,
        "questions": [
            {
                "number": q.number,
                "question_text": q.question_text,
                "choices": list(q.choices),
                "sub_items": list(q.sub_items),
                "has_table": q.has_table,
                "has_negative": q.has_negative,
                "negative_keyword": q.negative_keyword,
                "answer": q.answer,
                "explanation": q.explanation,
            }
            for q in document.questions
        ],
    }


def _is_windowsapps_python(path: Path) -> bool:
    text = str(path).lower()
    return "windowsapps" in text and path.name.lower().startswith("python")


def _resolve_python_executable() -> str:
    candidates: list[Path] = []
    for raw in (getattr(sys, "executable", ""), getattr(sys, "_base_executable", "")):
        if raw:
            candidates.append(Path(raw))

    try:
        install_root = Path.home() / "AppData" / "Local" / "Programs" / "Python"
        if install_root.exists():
            candidates.extend(sorted(install_root.glob("Python*/python.exe"), reverse=True))
    except Exception:
        pass

    seen: set[str] = set()
    for candidate in candidates:
        key = str(candidate)
        if not key or key in seen:
            continue
        seen.add(key)
        if not candidate.exists():
            continue
        if candidate.name.lower() not in {"python.exe", "pythonw.exe", "python3.exe"}:
            continue
        if _is_windowsapps_python(candidate):
            continue
        return key

    return sys.executable


def _create_worker_temp_dir() -> Path:
    roots = [Path.cwd() / ".runtime_tmp", Path(tempfile.gettempdir())]
    for root in roots:
        try:
            root.mkdir(parents=True, exist_ok=True)
        except Exception:
            continue

        for _ in range(8):
            candidate = root / f"hwp_gen_{uuid.uuid4().hex[:12]}"
            try:
                candidate.mkdir(parents=False, exist_ok=False)
                probe = candidate / ".write_probe"
                probe.write_text("ok", encoding="utf-8")
                probe.unlink(missing_ok=True)
                return candidate
            except Exception:
                shutil.rmtree(candidate, ignore_errors=True)
                continue

    raise RuntimeError("임시 작업 폴더를 만들 수 없습니다. 폴더 권한을 확인해 주세요.")


def _can_terminate_pid(pid: int) -> bool:
    if os.name != "nt":
        return True
    try:
        import ctypes

        PROCESS_TERMINATE = 0x0001
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        access = PROCESS_TERMINATE | PROCESS_QUERY_LIMITED_INFORMATION
        handle = ctypes.windll.kernel32.OpenProcess(access, False, int(pid))
        if not handle:
            return False
        ctypes.windll.kernel32.CloseHandle(handle)
        return True
    except Exception:
        # Do not over-filter on unexpected API errors.
        return True


def _get_hwp_pids(terminable_only: bool = False) -> set[int]:
    """현재 실행 중인 Hwp.exe 프로세스 PID 목록을 반환한다."""
    current_session_id: int | None = None
    if os.name == "nt":
        try:
            import ctypes
            import ctypes.wintypes

            sid = ctypes.wintypes.DWORD()
            if ctypes.windll.kernel32.ProcessIdToSessionId(int(os.getpid()), ctypes.byref(sid)):
                current_session_id = int(sid.value)
        except Exception:
            current_session_id = None

    try:
        creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq Hwp.exe", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=5,
            creationflags=creationflags,
        )
        pids: set[int] = set()
        for parts in csv.reader(result.stdout.strip().splitlines()):
            if len(parts) >= 2:
                try:
                    pid = int(parts[1].strip('"'))
                except ValueError:
                    pass
                else:
                    if current_session_id is not None and len(parts) >= 4:
                        try:
                            session_id = int(parts[3].strip('"'))
                        except ValueError:
                            session_id = None
                        if session_id is not None and session_id != current_session_id:
                            continue
                    if terminable_only and not _can_terminate_pid(pid):
                        continue
                    pids.add(pid)
        return pids
    except Exception:
        return set()


def _kill_hwp_pids(pids: set[int]) -> None:
    """지정된 PID의 Hwp.exe 프로세스를 강제 종료한다."""
    if not pids:
        return
    pids = {pid for pid in pids if _can_terminate_pid(pid)}
    if not pids:
        return
    creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    for pid in pids:
        try:
            subprocess.run(
                ["taskkill", "/F", "/T", "/PID", str(pid)],
                capture_output=True, timeout=5,
                creationflags=creationflags,
            )
        except Exception:
            pass


def _kill_orphaned_hwp(pids_before: set[int]) -> None:
    """서브프로세스 실행 전 PID 목록과 비교하여 새로 생긴 고아 Hwp.exe를 종료한다."""
    pids_after = _get_hwp_pids(terminable_only=True)
    orphaned = pids_after - pids_before
    if not orphaned:
        return
    _append_generation_log(f"killing {len(orphaned)} orphaned Hwp.exe: {orphaned}")
    _kill_hwp_pids(orphaned)


def _kill_all_hwp() -> None:
    """실행 중인 모든 Hwp.exe 프로세스를 강제 종료한다 (최대 3회 시도)."""
    max_attempts = 3

    for attempt in range(max_attempts):
        pids = _get_hwp_pids(terminable_only=True)
        if not pids:
            return
        _append_generation_log(
            f"kill_all_hwp attempt {attempt + 1}/{max_attempts}: {len(pids)} pids={pids}"
        )
        _kill_hwp_pids(pids)
        time.sleep(1)
        remaining = _get_hwp_pids(terminable_only=True)
        if not remaining:
            return
        time.sleep(1)

    final = _get_hwp_pids(terminable_only=True)
    if final:
        _append_generation_log(f"kill_all_hwp: {len(final)} survived all attempts: {final}")


def _cleanup_stale_runtime_tmp() -> None:
    """10분 이상 경과한 hwp_gen_* 임시 디렉터리를 삭제한다."""
    runtime_tmp = Path.cwd() / ".runtime_tmp"
    if not runtime_tmp.exists():
        return
    stale_threshold = 600  # 10분
    now = time.time()
    try:
        for entry in runtime_tmp.iterdir():
            if not entry.is_dir() or not entry.name.startswith("hwp_gen_"):
                continue
            try:
                if (now - entry.stat().st_mtime) > stale_threshold:
                    shutil.rmtree(entry, ignore_errors=True)
            except Exception:
                pass
    except Exception:
        pass


def _append_generation_log(message: str) -> None:
    try:
        log_dir = Path.cwd() / "output"
        log_dir.mkdir(parents=True, exist_ok=True)
        log_path = log_dir / "generation_debug.log"
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        with log_path.open("a", encoding="utf-8") as file:
            file.write(f"[{timestamp}] {message}\n")
    except Exception:
        pass


class DropArea(QFrame):
    fileDropped = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setFrameStyle(QFrame.StyledPanel | QFrame.Sunken)
        self.setObjectName("DropArea")

        layout = QVBoxLayout()
        self.label = QLabel("여기에 HWP 파일을 끌어다 놓으세요\n또는 클릭하여 파일 선택")
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)
        self.setLayout(layout)

        self.setMinimumHeight(200)
        apply_shadow(self)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
            self.setStyleSheet("background-color: #e3f2fd; border: 2px dashed #2196f3;")
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.setStyleSheet("")

    def dropEvent(self, event):
        self.setStyleSheet("")
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        if files:
            self.fileDropped.emit(files[0])

    def mousePressEvent(self, event):
        self.fileDropped.emit("SELECT_FILE")


class GenerationWorker(QThread):
    succeeded = pyqtSignal(list, str)
    failed = pyqtSignal(str)
    progress = pyqtSignal(int, str)

    def __init__(
        self,
        document: ExamDocument,
        source_file: str,
        timeout_sec: int = 60,
        style_required: bool = False,
    ):
        super().__init__()
        self.document = document
        self.source_file = source_file
        self.timeout_sec = timeout_sec
        self.style_required = style_required
        self._proc: subprocess.Popen[str] | None = None
        self._last_progress_raw = ""
        self._last_progress_pct = 0
        self._last_progress_msg = ""
        self._cancel_reason = "출력 생성이 취소되었습니다."

    def cancel(self, reason: str = "출력 생성이 취소되었습니다.") -> None:
        self._cancel_reason = reason
        self.requestInterruption()
        self._terminate_process()

    def _terminate_process(self) -> None:
        proc = self._proc
        if proc is None or proc.poll() is not None:
            return
        try:
            proc.terminate()
        except Exception:
            pass
        try:
            proc.wait(timeout=3)
        except Exception:
            try:
                proc.kill()
            except Exception:
                pass

    def _run_subprocess_attempt(
        self,
        request_path: Path,
        result_path: Path,
        progress_path: Path,
        request: dict,
        timeout_sec: int,
        force_disable_style: bool,
    ) -> dict:
        payload = dict(request)
        if force_disable_style:
            payload["force_disable_style"] = True

        request_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        try:
            result_path.unlink(missing_ok=True)
            progress_path.unlink(missing_ok=True)
        except Exception:
            pass

        if getattr(sys, "frozen", False):
            cmd = [
                sys.executable,
                "--subprocess-generation",
                str(request_path),
                str(result_path),
            ]
        else:
            cmd = [
                _resolve_python_executable(),
                "-m",
                "core.subprocess_generation",
                str(request_path),
                str(result_path),
            ]

        creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        mode = "safe-no-style" if force_disable_style else "style"
        _append_generation_log(
            f"generation start | mode={mode} | timeout={timeout_sec}s | source={self.source_file}"
        )
        _append_generation_log(f"generation cmd | {' '.join(cmd)}")
        self._proc = subprocess.Popen(
            cmd,
            cwd=str(Path.cwd()),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=creationflags,
        )

        started_at = time.monotonic()
        last_progress_change_at = started_at
        last_progress_snapshot = ""
        stall_limit_sec = 120 if (self.style_required and not force_disable_style) else 30

        while True:
            if self.isInterruptionRequested():
                self._terminate_process()
                raise RuntimeError(self._cancel_reason)
            if self._proc.poll() is not None:
                break

            elapsed = time.monotonic() - started_at
            if elapsed >= timeout_sec:
                self._terminate_process()
                raise TimeoutError("generation subprocess timeout")

            self._read_progress(progress_path)

            # 정체 감지: 진행률 변화가 없으면 타임아웃
            current_snapshot = self._last_progress_raw
            if current_snapshot != last_progress_snapshot:
                last_progress_snapshot = current_snapshot
                last_progress_change_at = time.monotonic()
            elif elapsed > 10 and (time.monotonic() - last_progress_change_at) > stall_limit_sec:
                self._terminate_process()
                raise TimeoutError("generation subprocess stalled (no progress)")

            self.msleep(150)

        return_code = self._proc.returncode if self._proc is not None else -1
        self._proc = None
        _append_generation_log(f"generation done | code={return_code} | result_exists={result_path.exists()}")
        if not result_path.exists():
            details = f"exit code: {return_code}"
            _append_generation_log(f"generation missing result | details={details}")
            raise RuntimeError(f"출력 결과 파일을 받지 못했습니다. ({details})")

        return json.loads(result_path.read_text(encoding="utf-8"))

    def run(self):
        _cleanup_stale_runtime_tmp()
        temp_dir = _create_worker_temp_dir()
        request_path = temp_dir / "request.json"
        result_path = temp_dir / "result.json"
        progress_path = temp_dir / "progress.txt"

        # 이전 실행에서 남은 고아 Hwp.exe 정리 (COM 연결 차단 방지)
        # 이전 실행에서 남은 고아 Hwp.exe 정리
        all_pre_pids = _get_hwp_pids()
        pre_pids = _get_hwp_pids(terminable_only=True)
        non_terminable = all_pre_pids - pre_pids
        if non_terminable:
            _append_generation_log(
                f"pre-cleanup: {len(non_terminable)} Hwp.exe skipped (insufficient terminate permission): {non_terminable}"
            )
        force_safe_mode = False
        if pre_pids:
            _append_generation_log(f"pre-cleanup: killing {len(pre_pids)} existing Hwp.exe: {pre_pids}")
            _kill_all_hwp()
            time.sleep(2)
            remaining = _get_hwp_pids(terminable_only=True)
            if remaining:
                _append_generation_log(
                    f"pre-cleanup: {len(remaining)} Hwp.exe still alive (style mode first, safe-mode retry on timeout)"
                )
                # Keep style mode first; fallback is handled by timeout-retry path.
                force_safe_mode = False

        hwp_pids_before = _get_hwp_pids(terminable_only=True)
        try:
            request = {
                "document": _document_to_payload(self.document),
                "source_file": self.source_file,
            }
            result = self._run_subprocess_attempt(
                request_path=request_path,
                result_path=result_path,
                progress_path=progress_path,
                request=request,
                timeout_sec=self.timeout_sec,
                force_disable_style=force_safe_mode,
            )
            if bool(result.get("ok")):
                _append_generation_log(
                    f"generation ok | files={len(result.get('files', []))} | warning={bool(result.get('warning', ''))}"
                )
                warning_text = result.get("warning", "")
                if force_safe_mode:
                    safe_note = "기존 HWP 프로세스가 종료되지 않아 안전모드(스타일 비활성화)로 생성했습니다."
                    warning_text = f"{safe_note}\n{warning_text}".strip() if warning_text else safe_note
                if self.style_required:
                    output_files = list(result.get("files", []))
                    has_non_hwp = any(Path(path).suffix.lower() != ".hwp" for path in output_files)
                    if has_non_hwp or warning_text.strip():
                        detail = warning_text.strip() or "스타일 필수 모드에서 비-HWP 출력이 감지되었습니다."
                        _append_generation_log(f"style-required failed | detail={detail[:500]}")
                        self.failed.emit(
                            "스타일 필수 모드에서는 경고/대체 출력을 허용하지 않습니다.\n"
                            f"{detail}"
                        )
                        return
                self.succeeded.emit(result.get("files", []), warning_text)
            else:
                error_text = str(result.get("error", "알 수 없는 오류로 생성에 실패했습니다."))
                _append_generation_log(f"generation failed | error={error_text[:500]}")
                if self._is_rpc_unavailable_error(error_text):
                    _append_generation_log("generation failed with RPC error -> retry once after HWP cleanup")
                    _kill_all_hwp()
                    time.sleep(2)
                    retry_timeout = max(45, min(90, self.timeout_sec + 30))
                    try:
                        retry_result = self._run_subprocess_attempt(
                            request_path=request_path,
                            result_path=result_path,
                            progress_path=progress_path,
                            request=request,
                            timeout_sec=retry_timeout,
                            force_disable_style=force_safe_mode,
                        )
                        if bool(retry_result.get("ok")):
                            _append_generation_log(
                                "rpc-retry ok | files="
                                f"{len(retry_result.get('files', []))} | warning={bool(retry_result.get('warning', ''))}"
                            )
                            warning_text = retry_result.get("warning", "")
                            if force_safe_mode:
                                safe_note = "safe-no-style mode was used because existing HWP processes could not be cleaned."
                                warning_text = f"{safe_note}\n{warning_text}".strip() if warning_text else safe_note
                            if self.style_required:
                                output_files = list(retry_result.get("files", []))
                                has_non_hwp = any(Path(path).suffix.lower() != ".hwp" for path in output_files)
                                if has_non_hwp or warning_text.strip():
                                    detail = warning_text.strip() or "style-required mode detected non-HWP output."
                                    _append_generation_log(f"style-required failed(after rpc retry) | detail={detail[:500]}")
                                    self.failed.emit(
                                        "style-required mode does not allow warning/fallback output.\n"
                                        f"{detail}"
                                    )
                                    return
                            self.succeeded.emit(retry_result.get("files", []), warning_text)
                            return
                        retry_error_text = str(retry_result.get("error", "rpc retry failed"))
                        _append_generation_log(f"rpc-retry failed | error={retry_error_text[:500]}")
                        error_text = f"{error_text}\n(auto-retry failed) {retry_error_text}"
                    except TimeoutError:
                        _append_generation_log("rpc-retry timeout")
                        error_text = f"{error_text}\n(auto-retry timeout)"
                    except Exception as retry_exc:
                        _append_generation_log(
                            f"rpc-retry exception | {type(retry_exc).__name__}: {retry_exc}"
                        )
                        error_text = f"{error_text}\n(auto-retry exception) {retry_exc}"
                self.failed.emit(error_text)
        except TimeoutError:
            _append_generation_log(
                f"generation timeout | last_progress={self._last_progress_pct}% | {self._last_progress_msg}"
            )
            # 타임아웃 시 모든 HWP 프로세스 정리 (이전 실행 고아 포함)
            _kill_all_hwp()
            time.sleep(1)

            if self.style_required:
                self.failed.emit(
                    f"스타일 필수 모드에서 출력 생성이 시간 초과로 중단되었습니다. "
                    f"(마지막 진행: {self._last_progress_pct}% - {self._last_progress_msg})"
                )
                return

            try:
                retry_timeout = max(30, min(60, self.timeout_sec))
                _append_generation_log(
                    f"retry start | mode=safe-no-style | timeout={retry_timeout}s"
                )
                result = self._run_subprocess_attempt(
                    request_path=request_path,
                    result_path=result_path,
                    progress_path=progress_path,
                    request=request,
                    timeout_sec=retry_timeout,
                    force_disable_style=True,
                )
                if bool(result.get("ok")):
                    warning_text = str(result.get("warning", "") or "").strip()
                    retry_warning = (
                        "스타일 적용 단계에서 응답 지연이 발생해 안전모드(스타일 비활성화)로 재시도했습니다."
                    )
                    if warning_text:
                        warning_text = f"{warning_text}\n{retry_warning}"
                    else:
                        warning_text = retry_warning
                    _append_generation_log(
                        f"retry ok | files={len(result.get('files', []))} | warning=True"
                    )
                    self.succeeded.emit(result.get("files", []), warning_text)
                else:
                    error_text = str(result.get("error", "안전모드 재시도에도 실패했습니다."))
                    _append_generation_log(f"retry failed | error={error_text[:500]}")
                    self.failed.emit(
                        f"출력 생성이 시간 초과되어 안전모드로 재시도했지만 실패했습니다.\n{error_text}"
                    )
            except TimeoutError:
                _append_generation_log("retry timeout")
                self.failed.emit(
                    f"출력 생성이 {self.timeout_sec}초를 초과해 중단되었습니다. "
                    f"(마지막 진행: {self._last_progress_pct}% - {self._last_progress_msg})\n"
                    "한글(HWP) 프로세스 상태를 확인 후 다시 시도해 주세요."
                )
            except Exception as retry_exc:
                _append_generation_log(
                    f"retry exception | {type(retry_exc).__name__}: {retry_exc}"
                )
                self.failed.emit(
                    f"출력 생성이 시간 초과되어 안전모드로 재시도했지만 실패했습니다.\n{retry_exc}"
                )
        except Exception as exc:
            _append_generation_log(f"generation exception | {type(exc).__name__}: {exc}")
            self.failed.emit(str(exc))
        finally:
            self._terminate_process()
            self._proc = None
            _kill_orphaned_hwp(hwp_pids_before)
            try:
                shutil.rmtree(temp_dir, ignore_errors=True)
            except Exception:
                pass

    def _read_progress(self, progress_path: Path) -> None:
        """progress.txt 파일을 읽어 progress 시그널을 발행한다."""
        try:
            if not progress_path.exists():
                return
            text = progress_path.read_text(encoding="utf-8").strip()
            if not text or text == self._last_progress_raw:
                return
            self._last_progress_raw = text
            if "|" not in text:
                return
            pct_str, msg = text.split("|", 1)
            pct = int(pct_str)
            self._last_progress_pct = pct
            self._last_progress_msg = msg
            self.progress.emit(pct, msg)
        except Exception:
            pass
    @staticmethod
    def _is_rpc_unavailable_error(message: str) -> bool:
        text = (message or "").lower()
        return (
            "-2147023174" in text
            or "rpc" in text
            or "서버를 사용할 수 없습니다" in text
            or "rpc server is unavailable" in text
        )


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("HWP 모의고사 자동 편집기 v2.0")
        icon_path = os.path.join("assets", "icon.ico")
        if os.path.isfile(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        self.resize(650, 550)

        self.config_manager = ConfigManager()
        self.service = ExamProcessingService(self.config_manager)
        self.selected_file = ""
        self.current_document: ExamDocument | None = None
        self._generate_worker: GenerationWorker | None = None
        self._generation_guard_timer = QTimer(self)
        self._generation_guard_timer.setSingleShot(True)
        self._generation_guard_timer.timeout.connect(self._on_generation_guard_timeout)

        _cleanup_stale_runtime_tmp()
        self.initUI()
        self.applyStyle()

    def initUI(self):
        menu_bar = self.menuBar()
        settings_action = QAction("설정", self)
        settings_action.triggered.connect(self.showSettings)
        help_action = QAction("안내", self)
        help_action.triggered.connect(self.showHelp)
        menu_bar.addAction(settings_action)
        menu_bar.addAction(help_action)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(30, 30, 30, 30)
        main_layout.setSpacing(20)

        title_label = QLabel("한국경찰학원 모의고사 편집기")
        title_label.setObjectName("MainTitle")
        main_layout.addWidget(title_label)

        subtitle_label = QLabel("선택한 원본 파일을 동원 형식으로 자동 변환합니다.")
        subtitle_label.setObjectName("SubTitle")
        main_layout.addWidget(subtitle_label)

        self.drop_area = DropArea()
        self.drop_area.fileDropped.connect(self.handleFileSelect)
        main_layout.addWidget(self.drop_area)

        self.file_info_label = QLabel("선택한 파일 없음")
        self.file_info_label.setObjectName("FileInfo")
        main_layout.addWidget(self.file_info_label)

        self.status_frame = QFrame()
        self.status_frame.setObjectName("StatusFrame")
        status_layout = QVBoxLayout(self.status_frame)
        self.status_label = QLabel("분석 결과: 대기 중")
        status_layout.addWidget(self.status_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        status_layout.addWidget(self.progress_bar)

        main_layout.addWidget(self.status_frame)

        btn_layout = QHBoxLayout()
        self.preview_btn = QPushButton("미리보기 확인")
        self.preview_btn.setEnabled(False)
        self.preview_btn.clicked.connect(self.showPreview)
        self.process_btn = QPushButton("문제지/해설지 생성")
        self.process_btn.setObjectName("PrimaryBtn")
        self.process_btn.setEnabled(False)
        self.process_btn.clicked.connect(self.startProcessing)

        btn_layout.addWidget(self.preview_btn)
        btn_layout.addWidget(self.process_btn)
        main_layout.addLayout(btn_layout)

        notice_label = QLabel("※ 헤더(동원명) 수정 등은 생성하지 않습니다.\n   1번 문제부터 편집하며 헤더는 직접 수정이 필요합니다.")
        notice_label.setObjectName("NoticeLabel")
        main_layout.addWidget(notice_label)

    def handleFileSelect(self, file_path):
        if file_path == "SELECT_FILE":
            options = QFileDialog.Options()
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "HWP 파일 선택",
                "",
                "HWP Files (*.hwp);;Text Files (*.txt);;All Files (*)",
                options=options,
            )

        if file_path and os.path.isfile(file_path):
            self.selected_file = file_path
            self.file_info_label.setText(f"선택됨: {os.path.basename(file_path)}")
            self.parseSelectedFile()

    def parseSelectedFile(self):
        if not self.selected_file:
            return

        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
        self.status_label.setText("분석 진행 중..")
        QApplication.processEvents()

        try:
            document = self.service.parse_file(self.selected_file)
            self.current_document = document
            self.preview_btn.setEnabled(True)
            self.process_btn.setEnabled(True)

            file_type = "문제+해설 유형" if document.file_type == "TYPE_A" else "문제지만 있는 파일"
            subject = document.subject or "미분류"
            first_line = (
                document.questions[0].question_text.splitlines()[0]
                if document.questions and document.questions[0].question_text
                else "(1번 문항 본문 없음)"
            )
            self.status_label.setText(
                f"분석 완료: {os.path.basename(self.selected_file)} | {document.total_count}문항 | 유형: {file_type} | 과목: {subject}\n"
                f"1번 미리보기: {first_line[:80]}"
            )
        except ProcessingError as exc:
            self.current_document = None
            self.preview_btn.setEnabled(False)
            self.process_btn.setEnabled(False)
            self.status_label.setText("분석 실패")
            QMessageBox.warning(self, "분석 실패", str(exc))
        finally:
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(0)
            self.progress_bar.setVisible(False)

    def showPreview(self):
        if not self.current_document:
            QMessageBox.information(self, "안내", "먼저 파일을 선택해 분석을 완료해주세요.")
            return

        preview = PreviewWindow(self.current_document, self)
        if preview.exec_():
            self.current_document.refresh_total_count()
            self.status_label.setText(
                f"수정 반영됨: {self.current_document.total_count}문항"
            )

    def showSettings(self):
        settings = SettingsWindow(self.config_manager, self)
        if settings.exec_():
            self.service.reload_config()
            QMessageBox.information(self, "설정 저장", "설정이 저장되었습니다.")

    def showHelp(self):
        QMessageBox.information(
            self,
            "안내",
            "1) HWP 파일을 선택합니다.\n"
            "2) 분석 결과를 미리보기에서 확인/수정합니다.\n"
            "3) 생성 버튼으로 결과 파일을 출력합니다.\n\n"
            "참고: HWP 생성 환경이 없으면 텍스트(.txt)로 대체 저장될 수 있습니다.",
        )

    def startProcessing(self):
        if not self.current_document or not self.selected_file:
            QMessageBox.information(self, "안내", "먼저 파일을 선택해 분석을 완료해주세요.")
            return
        if self._generate_worker and self._generate_worker.isRunning():
            QMessageBox.information(self, "안내", "현재 출력 파일 생성이 진행 중입니다.")
            return

        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
        self.status_label.setText("출력 파일 생성 중..")
        QApplication.processEvents()
        self.preview_btn.setEnabled(False)
        self.process_btn.setEnabled(False)

        style_enabled = bool(self.config_manager.get("style.enabled", True))
        timeout_sec = 60
        self._generate_worker = GenerationWorker(
            self.current_document,
            self.selected_file,
            timeout_sec=timeout_sec,
            style_required=style_enabled,
        )
        self._generate_worker.succeeded.connect(self._on_generation_succeeded)
        self._generate_worker.failed.connect(self._on_generation_failed)
        self._generate_worker.finished.connect(self._on_generation_finished)
        self._generate_worker.progress.connect(self._on_generation_progress)
        # 가드 타이머: 첫 시도(timeout_sec) + 재시도(최대 60s) + 정리/대기 여유(30s)
        retry_timeout = max(30, min(60, self._generate_worker.timeout_sec))
        guard_ms = int((self._generate_worker.timeout_sec + retry_timeout + 30) * 1000)
        self._generation_guard_timer.start(guard_ms)
        _append_generation_log(f"ui start generation | guard_ms={guard_ms} | file={self.selected_file}")
        self._generate_worker.start()

    def _on_generation_succeeded(self, output_files: list[str], warning: str):
        self.service.last_warning = warning
        output_dir = str(Path(output_files[0]).parent) if output_files else "(없음)"
        warning_text = (warning or "").strip()
        status_suffix = " | 경고 있음" if warning_text else ""
        self.status_label.setText(
            f"생성 완료: {len(output_files)}개 파일 | 원본: {os.path.basename(self.selected_file)} | 저장 위치: {output_dir}{status_suffix}"
        )
        first_line = (
            self.current_document.questions[0].question_text.splitlines()[0]
            if self.current_document and self.current_document.questions and self.current_document.questions[0].question_text
            else "(1번 문항 본문 없음)"
        )
        QMessageBox.information(
            self,
            "생성 완료",
            f"원본 파일: {self.selected_file}\n"
            f"1번 미리보기: {first_line[:100]}\n\n"
            "다음 파일을 생성하였습니다:\n\n"
            + "\n".join(output_files)
            + ("\n\n서식/스타일 경고가 있어 상세 경고 창을 띄워 표시합니다." if warning_text else ""),
        )
        if warning_text:
            self._show_generation_warning_dialog(warning_text)

    def _on_generation_failed(self, error_message: str):
        self.status_label.setText("출력 생성 실패")
        QMessageBox.warning(self, "생성 실패", error_message)

    def _on_generation_progress(self, pct: int, msg: str):
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(pct)
        self.status_label.setText(f"출력 파일 생성 중.. {pct}% - {msg}")

    def _on_generation_finished(self):
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        if self.current_document:
            self.preview_btn.setEnabled(True)
            self.process_btn.setEnabled(True)
        self._generate_worker = None

    def _on_generation_guard_timeout(self):
        if not self._generate_worker or not self._generate_worker.isRunning():
            return
        self.status_label.setText("출력 생성 제한시간 초과 - 중단 중...")
        _append_generation_log("ui guard timeout -> cancel worker")
        self._generate_worker.cancel(
            "출력 생성 시간이 너무 오래 걸려 중단했습니다. "
            "output/generation_debug.log를 확인해 원인을 점검해 주세요."
        )

    def _show_generation_warning_dialog(self, warning_text: str):
        lines = [line.strip() for line in warning_text.splitlines() if line.strip()]
        details = "\n".join(f"- {line}" for line in lines) if lines else warning_text
        QMessageBox.warning(
            self,
            "서식/스타일 경고",
            "출력 파일 생성은 완료하였습니다.\n아래 경고를 확인하고 템플릿 스타일 설정을 점검해 주세요.\n\n" + details,
        )

    def applyStyle(self):
        self.setStyleSheet(APP_STYLE)
