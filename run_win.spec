# -*- mode: python ; coding: utf-8 -*-

import importlib.metadata as importlib_metadata
import ctypes.util
import os
from pathlib import Path

# Collect pythonnet and clr_loader runtime files (Python.Runtime.dll, clr.pyd, etc..)
binaries = []
hiddenimports = ['clr', 'clr_loader', 'pythonnet']
try:
    # Locate and include all pythonnet runtime files
    dist = importlib_metadata.distribution('pythonnet')
    dist_files = dist.files or []
    
    # Find all .dll and .pyd files from pythonnet's runtime directory
    runtime_files = [f for f in dist_files if f.parts[0] == 'pythonnet' and 
                     (str(f).endswith('.dll') or str(f).endswith('.pyd'))]
    
    for runtime_file in runtime_files:
        dll_path = str(dist.locate_file(runtime_file))
        if os.path.exists(dll_path):
            # Put all DLLs/PYDs into the bundled pythonnet directory structure
            binaries.append((dll_path, 'pythonnet'))
            
    # Also try to find Python.Runtime.dll via ctypes as fallback
    if not runtime_files:
        lib = ctypes.util.find_library('Python.Runtime')
        if lib:
            binaries.append((lib, '.'))
except Exception as e:
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