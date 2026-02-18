from __future__ import annotations

import os
import shutil


def ensure_clean_dispatch(prog_id: str):
    """gen_py 캐시를 보존하면서 COM 객체를 생성한다.

    캐시가 손상되어 dispatch에 실패하는 경우에만 캐시를 삭제 후 재시도한다.
    """
    import win32com
    import win32com.client

    # 1차: 캐시 보존 + DispatchEx (새 인스턴스 생성)
    try:
        return win32com.client.DispatchEx(prog_id)
    except Exception:
        pass

    # 2차: Dispatch fallback (역시 캐시 보존)
    try:
        return win32com.client.Dispatch(prog_id)
    except Exception:
        pass

    # 3차: 둘 다 실패 → 캐시 손상 의심 → 삭제 후 재시도
    gen_path = getattr(win32com, "__gen_path__", "")
    if gen_path and os.path.isdir(gen_path):
        shutil.rmtree(gen_path, ignore_errors=True)

    try:
        return win32com.client.DispatchEx(prog_id)
    except Exception:
        return win32com.client.Dispatch(prog_id)
