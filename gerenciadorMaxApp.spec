# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['gerenciadorMaxApp.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'ttkbootstrap',
        'ttkbootstrap.themes',
        'ttkbootstrap.style',
        'ttkbootstrap.constants',
        'ttkbootstrap.widgets',
        'win32api',
        'pywintypes',
        'app_config',
        'ini_service',
        'sql_service',
        'sevenzip',
        'webdav_client',
        'ui_app',
        'ui_config_window',
        'ui_theme',
    ],
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
    name='gerenciadorMaxApp',
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
