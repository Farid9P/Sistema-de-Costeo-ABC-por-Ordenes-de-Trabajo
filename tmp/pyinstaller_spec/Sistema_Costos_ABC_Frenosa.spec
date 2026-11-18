# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['C:\\Users\\farid\\OneDrive\\Desktop\\Farid\\Universidad Nacional de Ingenieria\\7 ciclo\\Sistema y Gestión Financiera\\Sistema de Costos ABC\\app\\main.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\farid\\OneDrive\\Desktop\\Farid\\Universidad Nacional de Ingenieria\\7 ciclo\\Sistema y Gestión Financiera\\Sistema de Costos ABC\\app\\schema.sql', 'app'), ('C:\\Users\\farid\\OneDrive\\Desktop\\Farid\\Universidad Nacional de Ingenieria\\7 ciclo\\Sistema y Gestión Financiera\\Sistema de Costos ABC\\images_models', 'images_models')],
    hiddenimports=[],
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
    a.binaries,
    a.datas,
    [],
    name='Sistema_Costos_ABC_Frenosa',
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
)
