from __future__ import annotations

from datetime import datetime
from pathlib import Path
import os
import re
import shutil
import unicodedata
from typing import Any, Callable, Optional

from .com_utils import ensure_clean_dispatch as _ensure_clean_dispatch

ProgressCallback = Optional[Callable[[int, str], None]]
from .formatter import HwpFormatter
from .models import ExamDocument, ExamQuestion

try:
    import win32com.client as win32
except ImportError:  # pragma: no cover - depends on environment
    win32 = None

try:
    import winreg
except ImportError:  # pragma: no cover - non-Windows
    winreg = None


class OutputGenerator:
    def __init__(self, config: dict[str, Any]) -> None:
        self.config = config
        self.formatter = HwpFormatter(config)
        self.last_warning = ""
        self.style_config = config.get("style", {})
        self._base_style_enabled = bool(self.style_config.get("enabled", True))
        self._style_required = self._base_style_enabled
        self._template_path_raw = str(self.style_config.get("template_path", "")).strip()
        if not self._template_path_raw:
            self._template_path_raw = str(self.style_config.get("style_map_source", "")).strip()
        self._resolved_template_path: Path | None = None
        self._run_warnings: list[str] = []
        self._file_path_module_name = self._detect_file_path_check_module_name()
        self._module_dll_hint = str(self.style_config.get("module_dll_path", "")).strip()
        self._last_hwp_error: str = ""
        self._on_progress: ProgressCallback = None
        self._progress_pct: int = 0
        self._use_sub_items_table = bool(config.get("format", {}).get("sub_items_table", True))
        # Deprecated compatibility flag. Real-table insertion is always preferred.
        self._sub_items_box_mode = bool(config.get("format", {}).get("sub_items_box_mode", False))

    def _set_progress(self, pct: int, msg: str) -> None:
        self._progress_pct = int(pct)
        if self._on_progress:
            try:
                self._on_progress(self._progress_pct, msg)
            except Exception:
                pass

    def _heartbeat(self, msg: str) -> None:
        if self._on_progress:
            try:
                self._on_progress(self._progress_pct, msg)
            except Exception:
                pass

    def generate(self, document: ExamDocument, output_dir: str, source_stem: str, on_progress: ProgressCallback = None) -> list[str]:
        self._on_progress = on_progress
        self._progress_pct = 0
        style_required = self._base_style_enabled and self._style_required
        # Reset per-run style mode.
        self.formatter.use_styles = self._base_style_enabled
        self.formatter.reset_style_runtime_warnings()
        self._run_warnings = []
        self._last_hwp_error = ""
        self._resolved_template_path = self._resolve_template_path() if self._base_style_enabled else None
        config_warnings = self._collect_style_config_warnings()
        for warning in config_warnings:
            self._warn(warning)
        if style_required and config_warnings:
            raise RuntimeError(
                "스타일 필수 모드에서 설정 검증에 실패했습니다.\n" + "\n".join(config_warnings)
            )
        if self._base_style_enabled and self._resolved_template_path is None:
            if style_required:
                raise RuntimeError("스타일 필수 모드에서 템플릿 경로를 확인할 수 없습니다.")
            self.formatter.use_styles = False
            self._warn("템플릿을 사용할 수 없어 스타일 적용을 비활성화하고 직접 서식으로 출력합니다.")

        output_path = Path(output_dir).expanduser().resolve()
        output_path.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_stem = self._sanitize_filename_component(source_stem)

        if on_progress:
            self._set_progress(2, "HWP 출력 준비 중...")
        hwp_files = self._try_generate_hwp(document, output_path, safe_stem, timestamp, on_progress)
        for warning in self.formatter.style_runtime_warnings:
            self._warn(warning)
        if style_required and self.formatter.style_runtime_warnings:
            if hwp_files:
                self._cleanup_generated_files(hwp_files)
            raise RuntimeError(
                "스타일 필수 모드에서 스타일 적용 오류가 발생했습니다.\n"
                + "\n".join(self.formatter.style_runtime_warnings)
            )
        if hwp_files:
            if style_required and self._run_warnings:
                self._cleanup_generated_files(hwp_files)
                raise RuntimeError(
                    "스타일 필수 모드에서 경고가 발생해 출력을 중단했습니다.\n"
                    + "\n".join(self._run_warnings)
                )
            self.last_warning = "\n".join(self._run_warnings).strip()
            return hwp_files

        if self._last_hwp_error:
            self._warn(f"HWP 생성 실패 원인: {self._last_hwp_error}")
        if style_required:
            details = self._last_hwp_error or "\n".join(self._run_warnings)
            raise RuntimeError(details or "스타일 필수 모드에서 HWP 출력 생성에 실패했습니다.")
        self._warn("\uD55C\uAE00(HWP) \uCD9C\uB825 \uC0DD\uC131\uC744 \uC2E4\uD589\uD558\uC9C0 \uBABB\uD574 .txt \uB300\uCCB4 \uCD9C\uB825\uC73C\uB85C \uC804\uD658\uD588\uC2B5\uB2C8\uB2E4.")
        generated = self._generate_txt_fallback(document, output_path, safe_stem, timestamp)
        self.last_warning = "\n".join(self._run_warnings).strip()
        return generated

    @staticmethod
    def _cleanup_generated_files(paths: list[str]) -> None:
        for path in paths:
            try:
                Path(path).unlink(missing_ok=True)
            except Exception:
                continue

    def _sanitize_filename_component(self, value: str) -> str:
        text = unicodedata.normalize("NFKC", value or "")
        text = "".join(ch for ch in text if unicodedata.category(ch) != "Cs")
        text = "".join(ch for ch in text if ord(ch) >= 32 and not (127 <= ord(ch) <= 159))
        text = re.sub(r'[\\/:*?"<>|]+', "_", text)
        text = re.sub(r"\s+", " ", text).strip().strip(".")
        if not text:
            return "exam"

        reserved = {
            "CON", "PRN", "AUX", "NUL",
            "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
            "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
        }
        if text.upper() in reserved:
            text = f"file_{text}"
        return text[:120]

    def _warn(self, message: str) -> None:
        text = message.strip()
        if text and text not in self._run_warnings:
            self._run_warnings.append(text)

    def _collect_style_config_warnings(self) -> list[str]:
        if not self._base_style_enabled:
            return []

        warnings: list[str] = []
        if not self._template_path_raw:
            warnings.append(
                "\uC2A4\uD0C0\uC77C \uC5F0\uB3D9\uC774 \uCF1C\uC838 \uC788\uC9C0\uB9CC \uD15C\uD50C\uB9BF \uACBD\uB85C\uAC00 \uBE44\uC5B4 \uC788\uC2B5\uB2C8\uB2E4."
            )

        style_targets = [
            ("\uBB38\uC81C", self.formatter.question_style),
            ("\uC9C0\uBB38", self.formatter.passage_style),
            ("\uC120\uC9C0", self.formatter.choice_style),
            ("\uC18C\uBB38\uD56D", self.formatter.sub_items_style),
            ("\uD574\uC124", self.formatter.explanation_style),
        ]
        for label, style_name in style_targets:
            if not style_name.strip():
                warnings.append(f"{label} \uC2A4\uD0C0\uC77C\uBA85\uC774 \uBE44\uC5B4 \uC788\uC5B4 \uAE30\uBCF8 \uC11C\uC2DD\uC73C\uB85C \uB300\uCCB4\uB429\uB2C8\uB2E4.")
                continue
            if not self.formatter.has_style(style_name):
                warnings.append(f"{label} \uC2A4\uD0C0\uC77C\uC744 \uCC3E\uC9C0 \uBABB\uD588\uC2B5\uB2C8\uB2E4: {style_name}")
        return warnings

    def _resolve_template_path(self) -> Path | None:
        if not self._template_path_raw:
            return None
        path = Path(self._template_path_raw).expanduser()
        if not path.is_absolute():
            path = Path.cwd() / path
        path = path.resolve()
        if not path.exists():
            self._warn(
                f"\uC124\uC815\uD55C \uD15C\uD50C\uB9BF \uD30C\uC77C\uC744 \uCC3E\uC9C0 \uBABB\uD588\uC2B5\uB2C8\uB2E4: {path}"
            )
            return None
        return path

    def _detect_file_path_check_module_name(self) -> str | None:
        if winreg is None:
            return None

        key_paths = [
            r"Software\HNC\HwpAutomation\Modules",
            r"Software\WOW6432Node\HNC\HwpAutomation\Modules",
        ]
        roots = [winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE]
        access_modes = [winreg.KEY_READ]
        for wow_flag_name in ("KEY_WOW64_32KEY", "KEY_WOW64_64KEY"):
            wow_flag = getattr(winreg, wow_flag_name, 0)
            if wow_flag:
                access_modes.append(winreg.KEY_READ | wow_flag)

        candidates: list[str] = []
        for key_path in key_paths:
            for root in roots:
                for access in access_modes:
                    try:
                        key = winreg.OpenKey(root, key_path, 0, access)
                    except OSError:
                        continue
                    with key:
                        index = 0
                        while True:
                            try:
                                name, value, _ = winreg.EnumValue(key, index)
                            except OSError:
                                break
                            if isinstance(value, str) and value.strip():
                                cleaned = name.strip()
                                if cleaned and cleaned not in candidates:
                                    candidates.append(cleaned)
                            index += 1
        if "FilePathCheckerModule" in candidates:
            return "FilePathCheckerModule"
        if "SecurityModule" in candidates:
            return "SecurityModule"
        return candidates[0] if candidates else None

    def _candidate_file_path_module_dll_paths(self) -> list[Path]:
        candidates: list[Path] = []
        if self._module_dll_hint:
            candidates.append(Path(self._module_dll_hint).expanduser())

        install_roots: list[Path] = []
        for env_key in ("ProgramFiles(x86)", "ProgramFiles"):
            base = os.environ.get(env_key, "").strip()
            if base:
                install_roots.append(Path(base) / "Hnc")

        for root in install_roots:
            if not root.exists():
                continue
            for pattern in ("HOffice*/Bin/FilePathCheckerModule*.dll", "HOffice*/Bin/SecurityModule*.dll"):
                candidates.extend(sorted(root.glob(pattern)))

        project_root = Path.cwd()
        for pattern in ("**/FilePathCheckerModule*.dll", "**/SecurityModule*.dll"):
            candidates.extend(sorted(project_root.glob(pattern)))

        deduped: list[Path] = []
        seen: set[str] = set()
        for path in candidates:
            try:
                resolved = str(path.resolve())
            except Exception:
                resolved = str(path)
            if resolved in seen:
                continue
            seen.add(resolved)
            deduped.append(path)
        return deduped

    def _ensure_file_path_module_registry(self) -> tuple[str | None, str | None, bool]:
        if winreg is None:
            return None, None, False

        key_paths = [
            r"Software\HNC\HwpAutomation\Modules",
            r"Software\WOW6432Node\HNC\HwpAutomation\Modules",
        ]
        access_modes = [winreg.KEY_READ]
        for wow_flag_name in ("KEY_WOW64_32KEY", "KEY_WOW64_64KEY"):
            wow_flag = getattr(winreg, wow_flag_name, 0)
            if wow_flag:
                access_modes.append(winreg.KEY_READ | wow_flag)

        existing: list[tuple[str, str]] = []
        roots = [winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE]
        for key_path in key_paths:
            for root in roots:
                for access in access_modes:
                    try:
                        key = winreg.OpenKey(root, key_path, 0, access)
                    except OSError:
                        continue
                    with key:
                        index = 0
                        while True:
                            try:
                                name, value, _ = winreg.EnumValue(key, index)
                            except OSError:
                                break
                            index += 1
                            module_name = str(name or "").strip()
                            dll_path = str(value or "").strip()
                            if not module_name or not dll_path:
                                continue
                            dll = Path(dll_path).expanduser()
                            if dll.exists():
                                existing.append((module_name, str(dll)))

        preferred_order = ("FilePathCheckerModule", "SecurityModule")
        for preferred in preferred_order:
            for module_name, dll_path in existing:
                if module_name == preferred:
                    return module_name, dll_path, False
        if existing:
            return existing[0][0], existing[0][1], False

        for dll_path in self._candidate_file_path_module_dll_paths():
            try:
                resolved = str(dll_path.expanduser().resolve())
            except Exception:
                resolved = str(dll_path)
            if not Path(resolved).exists():
                continue
            wrote = False
            for key_path in key_paths:
                try:
                    key = winreg.CreateKeyEx(
                        winreg.HKEY_CURRENT_USER,
                        key_path,
                        0,
                        winreg.KEY_SET_VALUE,
                    )
                    with key:
                        winreg.SetValueEx(key, "FilePathCheckerModule", 0, winreg.REG_SZ, resolved)
                    wrote = True
                except OSError:
                    continue
            if wrote:
                return "FilePathCheckerModule", resolved, True
        return None, None, False

    def _register_file_path_module(self, hwp) -> bool:
        tried: set[str] = set()

        def _try_register(name: str | None) -> bool:
            module_name = (name or "").strip()
            if not module_name or module_name in tried:
                return False
            tried.add(module_name)
            try:
                if bool(hwp.RegisterModule("FilePathCheckDLL", module_name)):
                    self._file_path_module_name = module_name
                    return True
            except Exception:
                return False
            return False

        ordered_candidates = [
            self._file_path_module_name,
            "FilePathCheckerModule",
            "SecurityModule",
        ]
        for name in ordered_candidates:
            if _try_register(name):
                return True

        ensured_name, ensured_path, created = self._ensure_file_path_module_registry()
        if ensured_name and _try_register(ensured_name):
            if created and ensured_path:
                self._warn(
                    f"HwpAutomation Modules 레지스트리가 없어 자동 등록했습니다: "
                    f"{ensured_name} -> {ensured_path}"
                )
            return True
        return False

    def _try_generate_hwp(
        self,
        document: ExamDocument,
        output_path: Path,
        source_stem: str,
        timestamp: str,
        on_progress: ProgressCallback = None,
    ) -> list[str]:
        if win32 is None:
            self._last_hwp_error = "pywin32(win32com) 모듈을 사용할 수 없습니다."
            return []

        is_type_a = document.file_type == "TYPE_A"
        generated: list[str] = []
        question_file = output_path / f"{source_stem}_question_sheet_{timestamp}.hwp"
        if self._write_question_sheet_hwp(question_file, document, on_progress, is_type_a):
            generated.append(str(question_file))
        else:
            return []

        if is_type_a:
            explanation_file = output_path / f"{source_stem}_explanation_sheet_{timestamp}.hwp"
            if self._write_explanation_sheet_hwp(explanation_file, document, on_progress):
                generated.append(str(explanation_file))
            else:
                return []

        if on_progress:
            on_progress(100, "완료")
        return generated

    def _generate_txt_fallback(
        self,
        document: ExamDocument,
        output_path: Path,
        source_stem: str,
        timestamp: str,
    ) -> list[str]:
        safe_stem = self._sanitize_filename_component(source_stem)
        question_file = output_path / f"{safe_stem}_question_sheet_{timestamp}.txt"
        try:
            self._write_question_sheet_txt(question_file, document)
        except OSError:
            question_file = output_path / f"exam_question_sheet_{timestamp}.txt"
            self._write_question_sheet_txt(question_file, document)
            self._warn("출력 파일명에 사용할 수 없는 문자가 있어 기본 파일명(exam)으로 저장했습니다.")
        generated = [str(question_file)]

        if document.file_type == "TYPE_A":
            explanation_file = output_path / f"{safe_stem}_explanation_sheet_{timestamp}.txt"
            try:
                self._write_explanation_sheet_txt(explanation_file, document)
            except OSError:
                explanation_file = output_path / f"exam_explanation_sheet_{timestamp}.txt"
                self._write_explanation_sheet_txt(explanation_file, document)
                self._warn("해설 파일명에 사용할 수 없는 문자가 있어 기본 파일명(exam)으로 저장했습니다.")
            generated.append(str(explanation_file))

        return generated

    def _write_question_sheet_hwp(
        self, path: Path, document: ExamDocument,
        on_progress: ProgressCallback = None, has_explanation: bool = False,
    ) -> bool:
        hwp = None
        success = False
        try:
            if on_progress:
                self._set_progress(5, "문제지 HWP 열는 중...")
            hwp = self._open_hwp_document(target_path=path)
            self.formatter.setup_page(hwp)
            self.formatter.setup_columns(hwp)
            self._prime_body_start(hwp)

            total = len(document.questions)
            for i, question in enumerate(document.questions):
                if on_progress and total > 0:
                    if has_explanation:
                        pct = 8 + int((i / total) * 42)  # 8% ~ 50%
                    else:
                        pct = 8 + int((i / total) * 82)  # 8% ~ 90%
                    self._set_progress(pct, f"문제지 작성 중: {i + 1}/{total}")
                self._insert_question_block(hwp, question)

            # Skip global table walk here.
            # Per-table treat-as-char is applied during insertion, and global
            # control rewrites can destabilize anchors in some HWP builds.
            save_pct = 52 if has_explanation else 92
            if on_progress:
                self._set_progress(save_pct, "문제지 저장 중...")
            self._save_hwp(hwp, path)
            success = True
        except Exception as exc:
            self._last_hwp_error = str(exc)
        finally:
            self._quit_hwp(hwp)
        # COM 종료 후 바이너리 후처리로 style_id 설정
        if success and self.formatter.use_styles:
            try:
                ok = self.formatter.post_process_style_ids(path)
                if not ok:
                    self._warn("문제지 스타일 후처리가 실패했습니다 (post_process_style_ids=False)")
            except Exception as exc:
                self._warn(f"문제지 스타일 후처리 오류: {exc}")
        if success:
            try:
                self.formatter.post_process_question_emphasis_faces(path)
            except Exception as exc:
                self._warn(f"문제지 강조폰트 후처리 오류: {exc}")
        return success

    def _write_explanation_sheet_hwp(
        self, path: Path, document: ExamDocument,
        on_progress: ProgressCallback = None,
    ) -> bool:
        hwp = None
        success = False
        try:
            if on_progress:
                self._set_progress(55, "해설지 HWP 열는 중...")
            hwp = self._open_hwp_document(target_path=path)
            self.formatter.setup_page(hwp)
            self.formatter.setup_columns(hwp)
            self._prime_body_start(hwp)

            total = len(document.questions)
            for i, question in enumerate(document.questions):
                if on_progress and total > 0:
                    pct = 58 + int((i / total) * 35)  # 58% ~ 93%
                    self._set_progress(pct, f"해설지 작성 중: {i + 1}/{total}")
                self._insert_explanation_block(hwp, question)

            if on_progress:
                self._set_progress(95, "해설지 저장 중...")
            self._save_hwp(hwp, path)
            success = True
        except Exception as exc:
            self._last_hwp_error = str(exc)
        finally:
            self._quit_hwp(hwp)
        # COM 종료 후 바이너리 후처리로 style_id 설정
        if success and self.formatter.use_styles:
            try:
                ok = self.formatter.post_process_style_ids(path)
                if not ok:
                    self._warn("해설지 스타일 후처리가 실패했습니다 (post_process_style_ids=False)")
            except Exception as exc:
                self._warn(f"해설지 스타일 후처리 오류: {exc}")
        return success

    def _open_hwp_document(self, target_path: Path | None = None):
        hwp = _ensure_clean_dispatch("HWPFrame.HwpObject")
        hwp.XHwpWindows.Item(0).Visible = False
        module_registered = self._register_file_path_module(hwp)
        self._set_silent_message_boxes(hwp)

        if self._resolved_template_path is not None:
            if module_registered:
                # 템플릿을 출력 경로에 복사한 뒤 열기 (SaveAs 실패 방지)
                open_path = self._resolved_template_path
                if target_path is not None:
                    try:
                        target_path.parent.mkdir(parents=True, exist_ok=True)
                        shutil.copy2(str(self._resolved_template_path), str(target_path))
                        open_path = target_path
                    except Exception:
                        pass  # 복사 실패 시 원본 템플릿으로 열기

                opened = self._open_template_document(hwp, open_path)
                if opened:
                    self._prepare_document_from_template(hwp)
                    return hwp

                if self._style_required and self._base_style_enabled:
                    raise RuntimeError(
                        f"스타일 필수 모드에서 템플릿 열기에 실패했습니다: {self._resolved_template_path}"
                    )
                self.formatter.use_styles = False
                self._warn(
                    f"템플릿 열기에 실패해 빈 문서 기반으로 계속합니다: {self._resolved_template_path}"
                )
                self._warn("이번 출력에서는 스타일 적용을 끄고 직접 서식으로 출력합니다.")
            else:
                if self._style_required and self._base_style_enabled:
                    raise RuntimeError(
                        "스타일 필수 모드에서 보안 모듈 RegisterModule 호출이 실패했습니다. "
                        "HwpAutomation Modules 레지스트리와 DLL 경로를 확인해 주세요."
                    )
                self.formatter.use_styles = False
                self._warn(
                    "보안 모듈 RegisterModule 호출이 실패해 템플릿 열기를 건너뜁니다. "
                    "레지스트리(HKCU/HKLM\\Software\\HNC\\HwpAutomation\\Modules) 값 이름과 DLL 경로를 확인해 주세요."
                )
                self._warn("이번 출력에서는 스타일 적용을 끄고 직접 서식으로 출력합니다.")

        try:
            hwp.XHwpDocuments.Add(1)
        except Exception:
            pass
        return hwp

    def _open_template_document(self, hwp, template_path: Path) -> bool:
        path_text = str(template_path)
        try:
            return bool(hwp.Open(path_text, "HWP", "forceopen:true"))
        except Exception:
            return False

    def _set_silent_message_boxes(self, hwp) -> None:
        # Prevent hidden-window modal prompts (path/security/version warnings).
        try:
            hwp.SetMessageBoxMode(0x000F0000)
        except Exception:
            return

    def _restore_message_boxes(self, hwp) -> None:
        try:
            hwp.SetMessageBoxMode(0x00000000)
        except Exception:
            return

    def _prepare_document_from_template(self, hwp) -> None:
        # Keep style definitions from template and clear only body text.
        try:
            hwp.HAction.Run("SelectAll")
            hwp.HAction.Run("Delete")
            return
        except Exception:
            pass

        try:
            hwp.Run("SelectAll")
            hwp.Run("Delete")
        except Exception:
            self._warn("\uD15C\uD50C\uB9BF\uC744 \uC5F4\uC5C8\uC9C0\uB9CC \uBCF8\uBB38 \uC0AD\uC81C\uC5D0 \uC2E4\uD328\uD588\uC2B5\uB2C8\uB2E4.")

    def _save_hwp(self, hwp, path: Path) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        path_text = str(path)

        def _path_exists() -> bool:
            try:
                return path.exists() and path.stat().st_size > 0
            except Exception:
                return path.exists()

        # 1) 템플릿을 출력 경로에 복사해서 열었으면, 단순 Save로 저장 가능
        def _try_filesave() -> bool:
            try:
                return bool(hwp.HAction.Execute("FileSave", hwp.HParameterSet.HFileOpenSave.HSet))
            except Exception:
                pass
            try:
                hwp.Save()
                return True
            except Exception:
                pass
            return False

        if _path_exists() and _try_filesave():
            # 파일이 이미 해당 경로에 있고 Save 성공 → 완료
            if _path_exists():
                return

        def _try_filesaveas(action_name: str) -> bool:
            try:
                hwp.HAction.GetDefault(action_name, hwp.HParameterSet.HFileOpenSave.HSet)
                fs = hwp.HParameterSet.HFileOpenSave
                try:
                    fs.filename = path_text
                except Exception:
                    pass
                try:
                    fs.FileName = path_text
                except Exception:
                    pass
                try:
                    fs.Format = "HWP"
                except Exception:
                    pass
                try:
                    hset = fs.HSet
                    if hasattr(hset, "SetItem"):
                        hset.SetItem("filename", path_text)
                        hset.SetItem("FileName", path_text)
                        hset.SetItem("Format", "HWP")
                except Exception:
                    pass
                executed = bool(hwp.HAction.Execute(action_name, hwp.HParameterSet.HFileOpenSave.HSet))
                return executed and _path_exists()
            except Exception:
                return False

        if _try_filesaveas("FileSaveAs_S"):
            return
        if _try_filesaveas("FileSaveAs"):
            return

        try:
            hwp.SaveAs(path_text, "HWP", "")
            if _path_exists():
                return
        except Exception:
            pass

        try:
            hwp.SaveAs(path_text)
            if _path_exists():
                return
        except Exception:
            pass

        raise RuntimeError("HWPFrame.HwpObject.SaveAs")

    def _quit_hwp(self, hwp) -> None:
        if hwp is None:
            return
        hwp_pid = self._get_hwp_pid(hwp)
        self._set_silent_message_boxes(hwp)
        try:
            hwp.Clear(3)
        except Exception:
            pass
        try:
            hwp.Quit()
        except Exception:
            pass
        # COM 참조 명시 해제 (GC 지연으로 인한 고아 프로세스 방지)
        try:
            if hasattr(hwp, '_oleobj_'):
                del hwp._oleobj_
        except Exception:
            pass
        del hwp
        # HWP 프로세스가 실제로 종료되었는지 확인, 아니면 강제 종료
        if hwp_pid is not None:
            self._ensure_process_exited(hwp_pid)

    @staticmethod
    def _get_hwp_pid(hwp) -> int | None:
        """COM 객체의 윈도우 핸들로부터 HWP 프로세스 PID를 추출한다."""
        try:
            import ctypes
            import ctypes.wintypes

            hwnd = None
            try:
                hwnd = hwp.XHwpWindows.Item(0).WindowHandle
            except Exception:
                pass
            if hwnd:
                pid = ctypes.wintypes.DWORD()
                ctypes.windll.user32.GetWindowThreadProcessId(
                    int(hwnd), ctypes.byref(pid),
                )
                if pid.value:
                    return pid.value
        except Exception:
            pass
        return None

    @staticmethod
    def _ensure_process_exited(pid: int, timeout_sec: int = 5) -> None:
        """PID 프로세스가 종료되었는지 확인하고, 살아 있으면 강제 종료한다."""
        import subprocess as _sp
        import time

        creationflags = getattr(_sp, "CREATE_NO_WINDOW", 0)
        deadline = time.monotonic() + timeout_sec
        while time.monotonic() < deadline:
            try:
                result = _sp.run(
                    ["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
                    capture_output=True, text=True, timeout=3,
                    creationflags=creationflags,
                )
                if str(pid) not in result.stdout:
                    return
            except Exception:
                return
            time.sleep(0.5)
        try:
            _sp.run(
                ["taskkill", "/F", "/T", "/PID", str(pid)],
                capture_output=True, timeout=5,
                creationflags=creationflags,
            )
        except Exception:
            pass

    def _insert_text(self, hwp, text: str) -> None:
        normalized = text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
        try:
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
            hwp.HParameterSet.HInsertText.Text = normalized
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)
            return
        except Exception:
            pass

        try:
            hwp.InsertText(normalized)
        except Exception:
            return

    def _insert_question_block(self, hwp, question: ExamQuestion) -> None:
        self._heartbeat(f"문제 {question.number}: 시작")
        # Guard against leaked table context from previous question.
        if not self._leave_table_context(hwp):
            self._heartbeat(f"문제 {question.number}: 표 컨텍스트 정리 재시도")
            self._force_table_context_cleanup(hwp)

        number_prefix = f"{question.number}. "
        question_text = question.question_text or ""

        self._heartbeat(f"문제 {question.number}: 문제 스타일")
        self.formatter.apply_question_format(hwp, emphasize=False)
        # Ensure typed question text starts with question font even when the
        # paragraph begins after hidden control chars.
        self.formatter.apply_question_inline_char(hwp, emphasize=False)
        self._heartbeat(f"문제 {question.number}: 문제 본문 입력")
        self._insert_text(hwp, number_prefix)
        self._insert_question_text_with_emphasis(hwp, question_text, question.negative_keyword)
        self._insert_text(hwp, "\r\n")

        if question.sub_items:
            self._heartbeat(f"문제 {question.number}: 소문항 블록")
            self._insert_sub_items_block(
                hwp,
                question.sub_items,
                question.number,
                question.has_table,
                has_following_choices=bool(question.choices),
            )

        if question.choices:
            if not question.sub_items or self.formatter.choice_style != self.formatter.sub_items_style:
                self._heartbeat(f"문제 {question.number}: 선지 스타일")
                self.formatter.apply_choice_format(hwp)
            for choice_text in self._build_choice_lines(question.choices):
                self._insert_text(hwp, f"{choice_text}\r\n")

        self._insert_text(hwp, "\r\n")
        self._heartbeat(f"문제 {question.number}: 완료")

    def _insert_explanation_block(self, hwp, question: ExamQuestion) -> None:
        answer = question.answer or "-"
        self.formatter.apply_question_format(hwp, emphasize=False)
        self._insert_text(hwp, f"{question.number}. 정답 {answer}\r\n")

        if question.explanation:
            self.formatter.apply_explanation_format(hwp)
            self._insert_text(hwp, f"{question.explanation}\r\n")

        self._insert_text(hwp, "\r\n")

    def _insert_question_text_with_emphasis(self, hwp, text: str, keyword: str) -> None:
        if not keyword or keyword not in text:
            self._insert_text(hwp, text)
            return

        before, matched, after = text.partition(keyword)
        self._insert_text(hwp, before)

        # Keep direct formatting only for negative-keyword emphasis.
        self.formatter.apply_question_inline_char(hwp, emphasize=True)
        self._insert_text(hwp, matched)

        self.formatter.apply_question_inline_char(hwp, emphasize=False)
        self._insert_text(hwp, after)

    def _insert_sub_items_block(
        self,
        hwp,
        sub_items: list[str],
        question_number: int | None = None,
        prefer_table: bool = False,
        has_following_choices: bool = False,
    ) -> None:
        if not self._should_use_sub_items_table(sub_items, prefer_table=prefer_table):
            if question_number is not None and self._use_sub_items_table and len(sub_items) >= 2:
                self._heartbeat(f"문제 {question_number}: 소문항 표 생략(길이/복잡도)")
            self.formatter.apply_sub_items_format(hwp)
            for item in sub_items:
                self._insert_text(hwp, f"{item}\r\n")
            return

        try:
            self._heartbeat("소문항: 표 생성")
            created = self._create_single_cell_sub_items_table(hwp)
            if not created:
                raise RuntimeError("table create failed")
            # Keep this conservative: aggressive control-property writes can
            # destabilize 일부 HWP builds.
            self._set_current_table_treat_as_char(hwp)

            self._heartbeat("소문항: 스타일 적용")
            self.formatter.apply_sub_items_format(hwp)
            self._heartbeat("소문항: 텍스트 입력")
            self._insert_text(hwp, "\r\n".join(sub_items))
            self._set_current_table_treat_as_char(hwp)

            # CloseEx is the reliable way to exit table editing in HWP 2014.
            self._heartbeat("소문항: 표 컨텍스트 종료")
            if not self._leave_table_context(hwp):
                self._force_table_context_cleanup(hwp)
            self._move_caret_past_recent_table(hwp)
            if not has_following_choices:
                self._insert_text(hwp, "\r\n")
            self._heartbeat("소문항: 완료")
        except Exception as exc:
            self._heartbeat(f"소문항: 표 실패({type(exc).__name__}), 일반 입력 대체")
            self._leave_table_context(hwp)
            self._force_table_context_cleanup(hwp)
            self.formatter.apply_sub_items_format(hwp)
            for item in sub_items:
                self._insert_text(hwp, f"{item}\r\n")

    def _normalize_inline_choice_spacing(self, choice: str) -> str:
        text = (choice or "").strip()
        if not text:
            return ""
        marker_pattern = re.compile(r"[①②③④⑤]")
        matches = list(marker_pattern.finditer(text))
        if len(matches) < 2:
            return self._strip_choice_noise_suffix(text)

        parts: list[str] = []
        for index, match in enumerate(matches):
            start = match.start()
            end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
            part = text[start:end].strip()
            if not part:
                continue
            part = re.sub(r"\s+", " ", part)
            part = self._strip_choice_noise_suffix(part)
            parts.append(part)
        if len(parts) < 2:
            return self._strip_choice_noise_suffix(text)
        return (" " * 9).join(parts)

    @staticmethod
    def _strip_choice_noise_suffix(text: str) -> str:
        normalized = re.sub(r"\s+", " ", (text or "").strip())
        if not normalized:
            return ""

        # OLE 추출 노이즈가 선지 끝에 1글자로 끼는 케이스를 제거한다.
        # 예: U+3C72, U+2C86, U+4546
        def _is_noise_char(ch: str) -> bool:
            code = ord(ch)
            return (
                0x0370 <= code <= 0x03FF  # Greek
                or 0x2C80 <= code <= 0x2CFF  # Coptic
                or 0x3400 <= code <= 0x4DBF  # CJK Extension A
            )

        if normalized and _is_noise_char(normalized[-1]):
            normalized = normalized[:-1].rstrip()
        return normalized

    def _build_choice_lines(self, choices: list[str]) -> list[str]:
        lines = [
            self._normalize_inline_choice_spacing(choice)
            for choice in choices
            if (choice or "").strip()
        ]
        if self._can_compact_choice_lines(lines):
            normalized = [re.sub(r"\s+", " ", line.strip()) for line in lines]
            return [(" " * 9).join(normalized)]
        return lines

    @staticmethod
    def _can_compact_choice_lines(lines: list[str]) -> bool:
        if len(lines) < 2:
            return False
        marker_line = re.compile(r"^[①②③④⑤]\s*")
        marker_any = re.compile(r"[①②③④⑤]")

        normalized: list[str] = []
        payload_lengths: list[int] = []
        for line in lines:
            text = re.sub(r"\s+", " ", (line or "").strip())
            if not text:
                return False
            if len(marker_any.findall(text)) != 1:
                return False
            if not marker_line.match(text):
                return False
            body = marker_line.sub("", text, count=1).strip()
            if not body:
                return False
            normalized.append(text)
            payload_lengths.append(len(body))

        if not payload_lengths:
            return False
        if max(payload_lengths) > 20:
            return False
        total = sum(len(text) for text in normalized) + (len(normalized) - 1) * 9
        return total <= 110

    def _should_use_sub_items_table(self, sub_items: list[str], prefer_table: bool = False) -> bool:
        if not self._use_sub_items_table:
            return False
        if len(sub_items) < 2:
            return False
        if prefer_table:
            return True

        lengths = [len((line or "").strip()) for line in sub_items]
        if not lengths:
            return False

        total_len = sum(lengths)
        max_len = max(lengths)
        # Very long sub-item blocks tend to destabilize table anchoring in HWP COM.
        if total_len > 380:
            return False
        if max_len > 145:
            return False
        return True

    def _create_single_cell_sub_items_table(self, hwp) -> bool:
        """소문항 박스용 1x1 표를 생성한다. 실패 시 한 번 더 재시도한다."""
        for attempt in range(2):
            # CloseEx/Cancel은 비표 컨텍스트에서도 무해해서 선제적으로 정리한다.
            for _ in range(2):
                try:
                    hwp.HAction.Run("CloseEx")
                except Exception:
                    pass
                try:
                    hwp.HAction.Run("Cancel")
                except Exception:
                    pass
            # Selection/caret cleanup before TableCreate.
            # Keep the caret at current question position; forcing MoveDocEnd
            # can make table anchors drift toward later paragraphs.
            for action_name in ("Cancel",):
                try:
                    hwp.HAction.Run(action_name)
                except Exception:
                    pass

            try:
                if self._execute_table_create(hwp, use_custom_size=False):
                    return True
                if self._execute_table_create(hwp, use_custom_size=True):
                    return True
            except Exception as exc:
                self._heartbeat(f"소문항: 표 생성 예외({type(exc).__name__})")

            if attempt == 0:
                self._heartbeat("소문항: 표 생성 재시도")
                self._advance_to_fresh_region_for_table(hwp)
                for action_name in ("Cancel",):
                    try:
                        hwp.HAction.Run(action_name)
                    except Exception:
                        pass
        return False

    def _execute_table_create(self, hwp, use_custom_size: bool) -> bool:
        width_hwp = int(hwp.MiliToHwpUnit(self._get_sub_items_table_width_mm()))

        # Use legacy HParameterSet route only.
        # TreatAsChar is controlled by HTableCreation.TableProperties.
        hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
        table = hwp.HParameterSet.HTableCreation
        table.Rows = 1
        table.Cols = 1
        try:
            table_props = table.TableProperties
            table_props.TreatAsChar = 1
            table_props.TextWrap = 0
            table_props.TextFlow = 0
            table_props.FlowWithText = 1
            table_props.HorzRelTo = 0
            table_props.VertRelTo = 0
            table_props.WidthRelTo = 0
            table_props.HorzOffset = 0
            table_props.VertOffset = 0
            try:
                table.TableProperties = table_props
            except Exception:
                pass
        except Exception:
            pass
        if use_custom_size:
            table.WidthType = 2
            table.HeightType = 0
            table.WidthValue = width_hwp
            table.HeightValue = hwp.MiliToHwpUnit(0.0)
        return bool(hwp.HAction.Execute("TableCreate", table.HSet))

    def _set_current_table_treat_as_char(self, hwp) -> bool:
        """현재 표를 '글자처럼 취급'으로 강제 설정한다."""
        applied = False

        # Conservative action path first.
        try:
            action = hwp.CreateAction("TablePropertyDialog")
            action_set = action.CreateSet()
            action.GetDefault(action_set)
            action_set.SetItem("TreatAsChar", 1)
            action.Execute(action_set)
            applied = True
        except Exception:
            pass

        # Legacy fallback.
        try:
            hwp.HAction.GetDefault("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
            shape = hwp.HParameterSet.HShapeObject
            setattr(shape, "TreatAsChar", 1)
            if hwp.HAction.Execute("TablePropertyDialog", hwp.HParameterSet.HShapeObject.HSet):
                applied = True
        except Exception:
            pass

        # Direct property fallback (TreatAsChar only).
        matched_ctrl = False
        for ctrl in self._iter_current_table_controls(hwp):
            matched_ctrl = True
            if self._set_control_treat_as_char_only(ctrl):
                applied = True

        # Verify on currently bound table controls when discoverable.
        verified = False
        for ctrl in self._iter_current_table_controls(hwp):
            matched_ctrl = True
            if self._control_property_is_true(ctrl, "TreatAsChar"):
                verified = True
                break

        if matched_ctrl:
            return verified
        return applied

    def _set_control_treat_as_char_only(self, ctrl) -> bool:
        try:
            props = ctrl.Properties
            setter = getattr(props, "SetItem", None)
            if setter is None:
                return False
            setter("TreatAsChar", 1)
            ctrl.Properties = props
            return True
        except Exception:
            return False

    def _try_set_parameter_item(self, param_set, key: str, value: Any) -> bool:
        for target in (param_set, getattr(param_set, "HSet", None)):
            if target is None:
                continue
            setter = getattr(target, "SetItem", None)
            if setter is None:
                continue
            try:
                setter(key, value)
                return True
            except Exception:
                continue
        return False

    def _iter_current_table_controls(self, hwp):
        seen: set[int] = set()
        seeds = []
        for attr in ("CurSelectedCtrl", "ParentCtrl"):
            try:
                ctrl = getattr(hwp, attr, None)
                if ctrl is not None:
                    seeds.append(ctrl)
            except Exception:
                continue

        for seed in seeds:
            ctrl = seed
            for _ in range(8):
                if ctrl is None:
                    break
                key = id(ctrl)
                if key not in seen:
                    seen.add(key)
                    try:
                        ctrl_id = str(getattr(ctrl, "CtrlID", "")).lower()
                    except Exception:
                        ctrl_id = ""
                    if ctrl_id == "tbl":
                        yield ctrl
                try:
                    ctrl = getattr(ctrl, "ParentCtrl", None)
                except Exception:
                    break

    def _apply_table_control_properties(self, hwp, ctrl) -> bool:
        # Preferred: update current properties in-place to avoid resetting
        # unrelated anchor/wrapping fields.
        try:
            props = ctrl.Properties
            setter = getattr(props, "SetItem", None)
            if setter is not None:
                setter("TreatAsChar", 1)
                # Enforce inline-like layout to prevent drifting to page bottom.
                for key, value in (
                    ("TextWrap", 0),
                    ("TextFlow", 0),
                    ("FlowWithText", 1),
                    ("HorzRelTo", 0),
                    ("VertRelTo", 0),
                    ("WidthRelTo", 0),
                    ("HorzOffset", 0),
                    ("VertOffset", 0),
                ):
                    try:
                        setter(key, value)
                    except Exception:
                        continue
                ctrl.Properties = props
                if self._control_property_is_true(ctrl, "TreatAsChar"):
                    return True
        except Exception:
            pass

        for set_name in ("Table", "ShapeObject"):
            try:
                pset = hwp.CreateSet(set_name)
            except Exception:
                continue
            try:
                pset.SetItem("TreatAsChar", 1)
            except Exception:
                pass
            for key, value in (
                ("TextWrap", 0),
                ("TextFlow", 0),
                ("FlowWithText", 1),
                ("HorzRelTo", 0),
                ("VertRelTo", 0),
                ("WidthRelTo", 0),
                ("HorzOffset", 0),
                ("VertOffset", 0),
            ):
                try:
                    pset.SetItem(key, value)
                except Exception:
                    continue
            try:
                ctrl.Properties = pset
            except Exception:
                continue
            if self._control_property_is_true(ctrl, "TreatAsChar"):
                return True
        return False

    def _control_property_is_true(self, ctrl, key: str) -> bool:
        try:
            props = ctrl.Properties
        except Exception:
            return False

        for getter_name in ("Item", "GetItem"):
            getter = getattr(props, getter_name, None)
            if getter is None:
                continue
            try:
                value = getter(key)
                return bool(int(value))
            except Exception:
                continue
        return False

    def _read_table_control_property(self, ctrl, key: str):
        try:
            props = ctrl.Properties
        except Exception:
            return None

        for getter_name in ("Item", "GetItem"):
            getter = getattr(props, getter_name, None)
            if getter is None:
                continue
            try:
                return getter(key)
            except Exception:
                continue
        return None

    def _is_current_table_layout_suspicious(self, hwp) -> bool:
        min_width = int(hwp.MiliToHwpUnit(45.0))
        for ctrl in self._iter_current_table_controls(hwp):
            width = self._read_table_control_property(ctrl, "Width")
            try:
                width_val = int(width)
            except Exception:
                width_val = 0
            if width_val and width_val < min_width:
                return True

            wrap = self._read_table_control_property(ctrl, "TextWrap")
            try:
                if int(wrap) != 0:
                    return True
            except Exception:
                pass
        return False

    def _get_sub_items_table_width_mm(self) -> float:
        """현재 페이지/단 설정을 고려한 소문항 표 너비(mm)를 계산한다."""
        page_width = 210.0  # A4 portrait
        left_margin = float(self.formatter.page_config.get("left_margin", 15.0))
        right_margin = float(self.formatter.page_config.get("right_margin", 15.0))
        columns = max(1, int(self.formatter.columns))
        column_gap = 8.0 if columns > 1 else 0.0

        usable_width = page_width - left_margin - right_margin - (column_gap * max(0, columns - 1))
        per_column_width = usable_width / columns if columns > 0 else usable_width
        # Keep a safety margin for indentation and HWP internal fitting.
        safe_width = per_column_width - 6.0
        return max(50.0, min(82.0, safe_width))

    def _advance_to_fresh_region_for_table(self, hwp) -> None:
        """표 생성 실패 시 커서를 같은 문제 영역에서 가볍게 재정렬한다."""
        try:
            self._insert_text(hwp, "\r\n")
        except Exception:
            return

    def _prime_body_start(self, hwp) -> None:
        # Section/column setup can leave hidden control chars at paragraph start.
        # Start from a fresh paragraph so question #1 gets the same char shape as others.
        try:
            self._insert_text(hwp, "\r\n")
        except Exception:
            return

    def _is_in_table_context(self, hwp) -> bool:
        try:
            inside = bool(hwp.HAction.Run("TableCellBlock"))
            if inside:
                hwp.HAction.Run("Cancel")
            return inside
        except Exception:
            return False

    def _leave_table_context(self, hwp) -> bool:
        # Unconditionally try CloseEx/Cancel first.
        # Table-context probing is not always reliable in COM automation.
        for _ in range(8):
            try:
                hwp.HAction.Run("CloseEx")
            except Exception:
                pass
            try:
                hwp.HAction.Run("Cancel")
            except Exception:
                pass
        for _ in range(4):
            if not self._is_in_table_context(hwp):
                return True
            try:
                hwp.HAction.Run("CloseEx")
            except Exception:
                pass
            try:
                hwp.HAction.Run("Cancel")
            except Exception:
                pass
        return not self._is_in_table_context(hwp)

    def _force_table_context_cleanup(self, hwp) -> None:
        # Fallback cleanup used when context probing is unreliable.
        for _ in range(4):
            try:
                hwp.HAction.Run("CloseEx")
            except Exception:
                pass
            try:
                hwp.HAction.Run("Cancel")
            except Exception:
                pass

    def _move_caret_past_recent_table(self, hwp) -> None:
        # After leaving table edit mode, some builds keep the caret before the
        # table object. Move to paragraph end so subsequent text is inserted
        # after the table, not before it.
        for action_name in ("MoveParaEnd", "MoveRight"):
            try:
                hwp.HAction.Run(action_name)
            except Exception:
                continue
            try:
                hwp.HAction.Run("Cancel")
            except Exception:
                pass

    def _apply_table_box_border(self, hwp) -> None:
        try:
            hwp.HAction.Run("TableCellBlock")
            hwp.HAction.Run("TableCellBlockExtend")
            hwp.HAction.GetDefault("CellBorderFill", hwp.HParameterSet.HCellBorderFill.HSet)
            cb = hwp.HParameterSet.HCellBorderFill
            line_type = hwp.HwpLineType("Solid")
            line_width = hwp.HwpLineWidth("0.12mm")

            cb.ApplyTo = 0
            cb.AllCellsBorderFill = 1
            cb.BorderTypeTop = line_type
            cb.BorderTypeBottom = line_type
            cb.BorderTypeLeft = line_type
            cb.BorderTypeRight = line_type
            cb.BorderWidthTop = line_width
            cb.BorderWidthBottom = line_width
            cb.BorderWidthLeft = line_width
            cb.BorderWidthRight = line_width
            cb.BorderColorTop = 0
            cb.BorderColorBottom = 0
            cb.BorderColorLeft = 0
            cb.BorderColorRight = 0
            hwp.HAction.Execute("CellBorderFill", cb.HSet)
            hwp.HAction.Run("Cancel")
        except Exception:
            return

    def _insert_sub_items_boxed_block(self, hwp, sub_items: list[str]) -> None:
        self.formatter.apply_sub_items_format(hwp)
        border_top = "┌" + ("─" * 66) + "┐"
        border_bottom = "└" + ("─" * 66) + "┘"
        self._insert_text(hwp, f"{border_top}\r\n")
        for item in sub_items:
            self._insert_text(hwp, f"│ {item}\r\n")
        self._insert_text(hwp, f"{border_bottom}\r\n")

    def _normalize_all_tables_treat_as_char(self, hwp, quiet: bool = False) -> None:
        """문서 내 모든 표를 글자처럼 취급으로 보정해 위치 밀림을 줄인다."""
        normalized = 0
        ctrl = getattr(hwp, "HeadCtrl", None)
        while ctrl is not None:
            try:
                ctrl_id = str(getattr(ctrl, "CtrlID", "")).lower()
            except Exception:
                ctrl_id = ""

            if ctrl_id == "tbl":
                try:
                    if self._apply_table_control_properties(hwp, ctrl):
                        normalized += 1
                except Exception:
                    pass

            try:
                ctrl = getattr(ctrl, "Next", None)
            except Exception:
                break

        if normalized > 0 and not quiet:
            self._heartbeat(f"표 글자처럼 보정: {normalized}개")

    def _write_question_sheet_txt(self, path: Path, document: ExamDocument) -> None:
        with path.open("w", encoding="utf-8") as file:
            file.write("[문제지]\n")
            file.write(f"유형: {document.file_type}\n")
            file.write(f"문항 수: {document.total_count}\n")
            file.write("\n")
            for question in document.questions:
                file.write(self._render_question(question))
                file.write("\n")

    def _write_explanation_sheet_txt(self, path: Path, document: ExamDocument) -> None:
        with path.open("w", encoding="utf-8") as file:
            file.write("[해설지]\n")
            file.write(f"유형: {document.file_type}\n")
            file.write(f"문항 수: {document.total_count}\n")
            file.write("\n")
            for question in document.questions:
                answer = question.answer or "-"
                explanation = question.explanation or ""
                file.write(f"{question.number:02d}. 정답 {answer}\n")
                if explanation:
                    file.write(f"{explanation}\n")
                file.write("\n")

    def _render_question(self, question: ExamQuestion) -> str:
        lines = [f"{question.number:02d}. {question.question_text}".rstrip()]
        lines.extend(question.sub_items)
        lines.extend(question.choices)
        return "\n".join(lines).strip() + "\n"
