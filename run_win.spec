# -*- mode: python ; coding: utf-8 -*-

import importlib.metadata as importlib_metadata
import os
import shutil
from pathlib import Path

# Collect pythonnet runtime files to ensure DLL availability
binaries = []
hiddenimports = ['clr', 'clr_loader', 'pythonnet']

try:
    # Get pythonnet distribution
    dist = importlib_metadata.distribution('pythonnet')
    
    # Find pythonnet's runtime directory (contains Python.Runtime.dll)
    if hasattr(dist, '_path'):
        pythonnet_root = dist._path
    else:
        # Fallback: try to import and find the path
        import pythonnet
        pythonnet_root = Path(pythonnet.__file__).parent
    
    # Collect all DLLs and PYDs from pythonnet/runtime
    runtime_dir = Path(pythonnet_root) / 'runtime'
    if runtime_dir.exists():
        for dll_file in runtime_dir.glob('*.dll'):
            binaries.append((str(dll_file), 'pythonnet/runtime'))
        for pyd_file in runtime_dir.glob('*.pyd'):
            binaries.append((str(pyd_file), 'pythonnet/runtime'))
    
    # Also collect from pythonnet root
    for dll_file in Path(pythonnet_root).glob('*.dll'):
        binaries.append((str(dll_file), 'pythonnet'))
    for pyd_file in Path(pythonnet_root).glob('*.pyd'):
        binaries.append((str(pyd_file), 'pythonnet'))
        
except Exception as e:
    print(f"Warning: Could not collect pythonnet binaries: {e}")
    # PyInstaller's hooks may still handle this


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