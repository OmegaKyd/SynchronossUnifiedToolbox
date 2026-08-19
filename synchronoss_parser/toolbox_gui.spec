# -*- mode: python ; coding: utf-8 -*-

"""PyInstaller spec file for the toolbox GUI.

Generated to build a single-file, windowed executable. Adjust the
``datas`` list if additional resources need to be bundled with the
application.
"""

block_cipher = None


a = Analysis(
    ['toolbox_gui.py'],
    pathex=['.'],
    binaries=[],
    datas=[],
    hiddenimports=['pandas', 'openpyxl'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='toolbox_gui',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon='app_icon.ico',
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
