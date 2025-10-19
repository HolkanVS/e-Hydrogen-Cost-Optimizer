# pyinstaller must be installed in the virtual environment in order to import all the libraries
# python -m pip install pyinstaller
# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules, collect_data_files
import os

all_libraries = [
    'tkcalendar',
    'babel',
    'pyomo',
    'highspy',
    'encodings',
]
hidden_imports = [
]
for l in all_libraries:
    hidden_imports += collect_submodules(l)


# copy&paste /modelInputs and /modelOutputs folders to the /_internal folder after compiling
a = Analysis(
    ['hydrogen_optimizer_v_1_0_0.py'],
    pathex = [],
    binaries=[],
    datas=[ 
        ('.\\images','images')
        ],
    hiddenimports=hidden_imports,
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
    name='e-Hydrogen Cost Optimizer v1.0.0',
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
    icon='./images/favicon.ico'
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='e-Hydrogen Cost Optimizer v1.0.0',
)
