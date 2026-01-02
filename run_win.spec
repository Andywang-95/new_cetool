# -*- mode: python ; coding: utf-8 -*-

import importlib.metadata as importlib_metadata
import ctypes.util
import os

# Collect pythonnet runtime DLL (Python.Runtime.dll) if available on build machine
binaries = []
hiddenimports = ['clr', 'clr_loader']
try:
    # Prefer to locate the file via the distribution object so we get an absolute path
    dist = importlib_metadata.distribution('pythonnet')
    dist_files = dist.files or []
    runtime_dlls = [f for f in dist_files if f.name == 'Python.Runtime.dll']
    if runtime_dlls:
        # locate_file returns an absolute path to the file inside the distribution
        dll_path = str(dist.locate_file(runtime_dlls[0]))
        if os.path.exists(dll_path):
            binaries.append((dll_path, '.'))
    else:
        lib = ctypes.util.find_library('Python.Runtime')
        if lib:
            binaries.append((lib, '.'))
except Exception:
    # best-effort only; PyInstaller hooks may still handle this
    pass


a = Analysis(
    ['ce_tool.py'],
    pathex=['.'],
    binaries=binaries,
    datas=[
        ('app/templates', 'app/templates'),
        ('app/static', 'app/static'),
    ],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ce_tool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='ce_tool',
)