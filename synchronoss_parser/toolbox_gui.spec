# -*- mode: python ; coding: utf-8 -*-

"""PyInstaller spec for the Synchronoss Unified Toolbox GUI.

Build with:
  python -m synchronoss_parser.build_exe
or from this directory:
  pyinstaller --noconfirm --clean toolbox_gui.spec

Bundles Omega app icons for the Windows .exe resource and runtime Tk icon.
"""

from pathlib import Path

block_cipher = None

# SPECPATH is set by PyInstaller to the directory containing this .spec
SPEC_DIR = Path(SPECPATH).resolve()

ICON_ICO = SPEC_DIR / "app_icon.ico"
ICON_FILES = []
for name in ("app_icon.ico", "app_icon.png", "app_icon_32.png", "app_icon_64.png"):
    p = SPEC_DIR / name
    if p.is_file():
        ICON_FILES.append((str(p), "."))

a = Analysis(
    [str(SPEC_DIR / "toolbox_gui.py")],
    pathex=[str(SPEC_DIR), str(SPEC_DIR.parent)],
    binaries=[],
    datas=ICON_FILES,
    hiddenimports=["pandas", "openpyxl", "PIL", "PIL.Image", "PIL.ImageTk"],
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
    name="SynchronossUnifiedToolbox",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon=str(ICON_ICO) if ICON_ICO.is_file() else None,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
