from __future__ import annotations

import atexit
import csv
import json
import os
import signal
import subprocess
import sys
from pathlib import Path
from typing import Any

from .config_manager import ConfigManager
from .models import ExamDocument, ExamQuestion
from .service import ExamProcessingService

_hwp_pids_before: set[int] = set()


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
        return True


def _get_hwp_pids(terminable_only: bool = False) -> set[int]:
    """현재 실행 중인 Hwp.exe PID 목록."""
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
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
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


def _cleanup_orphaned_hwp() -> None:
    """프로세스 종료 시 이 서브프로세스가 만든 Hwp.exe를 정리한다."""
    orphaned = _get_hwp_pids(terminable_only=True) - _hwp_pids_before
    if not orphaned:
        return
    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    for pid in orphaned:
        try:
            subprocess.run(
                ["taskkill", "/F", "/PID", str(pid)],
                capture_output=True, timeout=5,
                creationflags=creationflags,
            )
        except Exception:
            pass


def _on_sigterm(signum, frame):
    _cleanup_orphaned_hwp()
    sys.exit(1)


def _payload_to_document(payload: dict[str, Any]) -> ExamDocument:
    questions: list[ExamQuestion] = []
    for item in payload.get("questions", []):
        questions.append(
            ExamQuestion(
                number=int(item.get("number", 0)),
                question_text=str(item.get("question_text", "")),
                choices=list(item.get("choices", [])),
                sub_items=list(item.get("sub_items", [])),
                has_table=bool(item.get("has_table", False)),
                has_negative=bool(item.get("has_negative", False)),
                negative_keyword=str(item.get("negative_keyword", "")),
                answer=item.get("answer"),
                explanation=item.get("explanation"),
            )
        )
    return ExamDocument(
        file_type=str(payload.get("file_type", "")),
        subject=str(payload.get("subject", "")),
        questions=questions,
        total_count=int(payload.get("total_count", len(questions))),
    )


def _write_result(path: Path, result: dict[str, Any]) -> None:
    path.write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")


def main(argv: list[str]) -> int:
    global _hwp_pids_before

    if len(argv) != 3:
        print("usage: python -m core.subprocess_generation <request_json> <result_json>")
        return 2

    _hwp_pids_before = _get_hwp_pids(terminable_only=True)
    atexit.register(_cleanup_orphaned_hwp)
    if os.name == "nt":
        signal.signal(signal.SIGBREAK, _on_sigterm)
    signal.signal(signal.SIGTERM, _on_sigterm)

    request_path = Path(argv[1])
    result_path = Path(argv[2])
    pythoncom = None
    initialized = False

    try:
        try:
            import pythoncom as _pythoncom

            pythoncom = _pythoncom
            pythoncom.CoInitialize()
            initialized = True
        except Exception:
            initialized = False

        request = json.loads(request_path.read_text(encoding="utf-8"))
        source_file = str(request.get("source_file", ""))
        document_payload = request.get("document", {})
        document = _payload_to_document(document_payload)
        force_disable_style = bool(request.get("force_disable_style", False))

        progress_path = request_path.parent / "progress.txt"

        def _report_progress(pct: int, msg: str) -> None:
            try:
                progress_path.write_text(f"{pct}|{msg}", encoding="utf-8")
            except Exception:
                pass

        cm = ConfigManager()
        preset_file = str(request.get("active_preset", ""))
        if preset_file:
            cm.load_with_preset(preset_file)
        service = ExamProcessingService(cm)
        if force_disable_style:
            try:
                service.generator._base_style_enabled = False
                service.generator.formatter.use_styles = False
                service.generator._resolved_template_path = None
                service.generator._warn("안전모드 재시도로 스타일 적용을 비활성화했습니다.")
            except Exception:
                pass
        output_files = service.generate_outputs(document, source_file, on_progress=_report_progress)
        _write_result(
            result_path,
            {
                "ok": True,
                "files": output_files,
                "warning": service.last_warning or "",
            },
        )
        return 0
    except Exception as exc:
        try:
            _write_result(result_path, {"ok": False, "error": str(exc)})
        except Exception:
            pass
        return 1
    finally:
        # COM 참조를 GC로 먼저 해제하여 고아 프로세스 방지
        import gc
        gc.collect()

        if initialized and pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
