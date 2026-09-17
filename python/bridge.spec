# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_all, collect_submodules

numpy_datas, numpy_binaries, numpy_hiddenimports = collect_all("numpy")
hiddenimports = numpy_hiddenimports + collect_submodules("win32com") + [
    "pythoncom", "pywintypes", "win32com.client", "pypdf",
]

a = Analysis(
    ["bridge.py"],
    pathex=["."],
    binaries=numpy_binaries,
    datas=numpy_datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["tkinter", "tkinterdnd2"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, [], exclude_binaries=True,
    name="word-batch-backend", debug=False,
    bootloader_ignore_signals=False, strip=False, upx=True, console=True,
)
coll = COLLECT(
    exe, a.binaries, a.datas, strip=False, upx=True,
    upx_exclude=[], name="word-batch-backend",
)
