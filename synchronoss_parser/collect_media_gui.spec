# -*- mode: python ; coding: utf-8 -*-

"""PyInstaller spec file for the Collect Media GUI.

Produces a single-file, windowed executable. Add any required data files
into the ``datas`` list.
"""

block_cipher = None


a = Analysis(
    ['collect_media_gui.py'],
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
    name='collect_media_gui',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
