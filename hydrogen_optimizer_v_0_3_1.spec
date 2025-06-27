# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

all_libraries = [
    'tkcalendar',
    'babel',
    'pyomo',
    'highspy'
]
hidden_imports = []
for l in all_libraries:
    hidden_imports += collect_submodules(l)

a = Analysis(
    ['hydrogen_optimizer_v_0_3_1.py'],
    pathex=[],
    binaries=[],
    datas=[ ('.\images','images'),
            ('.\CTkScrollableDropdown','CTkScrollableDropdown')
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
    name='e_Hydrogen_Cost_Optimizer_v_0_3_1',
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
    icon='.\images\hydrogen_optimizer_logo_v1.ico'
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='e_Hydrogen_Cost_Optimizer_v_0_3_1',
)
