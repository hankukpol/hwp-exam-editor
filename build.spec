# -*- mode: python ; coding: utf-8 -*-
# PyInstaller build spec for HWP 모의고사 자동 편집기
# 사용법: pyinstaller build.spec

import os

icon_path = os.path.join("assets", "icon.ico")
if not os.path.exists(icon_path):
    icon_path = None

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=[
        ("config/default_config.json", "config"),
    ],
    hiddenimports=[
        "win32com",
        "win32com.client",
        "win32com.server",
        "pythoncom",
        "pywintypes",
        "olefile",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="HWP모의고사편집기",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_path,
)
