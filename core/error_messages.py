from __future__ import annotations

from typing import Iterable


def _join_lines(lines: Iterable[str]) -> str:
    return "\n".join(line for line in lines if line.strip())


def _trim_raw_error(text: str, limit: int = 700) -> str:
    raw = (text or "").strip()
    if len(raw) <= limit:
        return raw
    return raw[:limit].rstrip() + " ..."


def build_generation_error_message(raw_message: str) -> str:
    raw = (raw_message or "").strip()
    lower = raw.lower()

    tips: list[str]
    if "-2147023174" in lower or "rpc" in lower:
        tips = [
            "한컴오피스(HWP) 연결이 끊어졌습니다.",
            "열려 있는 Hwp.exe를 모두 종료한 뒤 다시 시도해 주세요.",
            "동시에 한글을 여러 개 실행 중이면 하나만 남기고 닫아 주세요.",
        ]
    elif "class not registered" in lower or "invalid class string" in lower or "hwpframe.hwpobject" in lower:
        tips = [
            "한컴오피스 COM 구성요소를 찾지 못했습니다.",
            "한컴오피스(한글)가 설치되어 있는지 확인해 주세요.",
            "설치되어 있다면 한글을 한 번 실행한 뒤 다시 시도해 주세요.",
        ]
    elif "no module named win32com" in lower or "pywin32" in lower:
        tips = [
            "필수 구성요소(pywin32)가 설치되지 않았거나 손상되었습니다.",
            "프로그램 재설치 또는 pywin32 재설치를 진행해 주세요.",
        ]
    elif "registermodule" in lower or "filepathcheckdll" in lower or "보안 모듈" in raw:
        tips = [
            "한글 보안 모듈 등록에 실패했습니다.",
            "설정에서 보안 모듈 DLL 경로와 레지스트리 등록 상태를 확인해 주세요.",
            "관리자 권한으로 프로그램을 실행한 뒤 다시 시도해 주세요.",
        ]
    elif "timeout" in lower or "stalled" in lower or "시간 초과" in raw:
        tips = [
            "출력 생성 시간이 초과되었습니다.",
            "열려 있는 HWP를 종료한 뒤 다시 시도해 주세요.",
            "같은 파일을 한 번 더 시도해도 실패하면 PC 재부팅 후 재시도해 주세요.",
        ]
    elif "saveas" in lower or "filesave" in lower or "저장" in raw:
        tips = [
            "출력 파일 저장에 실패했습니다.",
            "출력 폴더 쓰기 권한과 파일 잠금(열려 있는지)을 확인해 주세요.",
        ]
    elif "임시 작업 폴더" in raw:
        tips = [
            "임시 폴더를 만들지 못했습니다.",
            "디스크 여유 공간과 폴더 권한을 확인해 주세요.",
        ]
    else:
        tips = [
            "출력 생성 중 오류가 발생했습니다.",
            "한컴오피스 설치/실행 상태를 확인한 뒤 다시 시도해 주세요.",
        ]

    detail = _trim_raw_error(raw)
    return _join_lines(
        [
            *tips,
            "",
            "[원본 오류]",
            detail or "(없음)",
        ]
    )


def build_parse_error_message(raw_message: str) -> str:
    raw = (raw_message or "").strip()
    lower = raw.lower()
    if "hwp 파일(.hwp)만 지원합니다" in raw:
        return raw
    if "암호로 보호" in raw:
        return _join_lines(
            [
                "암호가 설정된 HWP 파일은 자동 분석할 수 없습니다.",
                "비밀번호를 해제한 파일로 다시 시도해 주세요.",
                "",
                "[원본 오류]",
                _trim_raw_error(raw),
            ]
        )
    if "문제 번호를 찾지 못했습니다" in raw:
        return raw
    if "-2147023174" in lower or "rpc" in lower:
        return _join_lines(
            [
                "파일 분석 중 한글 연결(RPC) 오류가 발생했습니다.",
                "열려 있는 HWP를 종료한 뒤 다시 시도해 주세요.",
                "",
                "[원본 오류]",
                _trim_raw_error(raw),
            ]
        )
    return raw
