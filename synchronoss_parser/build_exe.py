#!/usr/bin/env python3
"""Build a windowed PyInstaller executable for the Synchronoss Unified Toolbox GUI.

Requires:  pip install pyinstaller

Run from anywhere; the script changes into the package directory so the
Omega ``app_icon.ico`` and related assets resolve correctly.
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
SPEC = HERE / "toolbox_gui.spec"
ICON = HERE / "app_icon.ico"


def main() -> None:
    if not SPEC.is_file():
        raise SystemExit(f"Spec not found: {SPEC}")
    if not ICON.is_file():
        print(f"WARNING: {ICON.name} not found – the .exe will use a default icon.")

    # Ensure relative paths inside the .spec resolve next to toolbox_gui.py
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        str(SPEC),
    ]
    print("Running:", " ".join(cmd))
    print("Working directory:", HERE)
    subprocess.run(cmd, check=True, cwd=str(HERE))

    dist = HERE / "dist"
    # Preferred name from the updated .spec
    candidates = [
        dist / "SynchronossUnifiedToolbox.exe",
        dist / "SynchronossUnifiedToolbox",
        dist / "toolbox_gui.exe",
        dist / "toolbox_gui",
    ]
    built = next((p for p in candidates if p.exists()), None)
    if built:
        print(f"\nBuilt: {built}")
        print("The Omega icon is embedded in the .exe and bundled for the window.")
    else:
        print(f"\nBuild finished. Check output under: {dist}")


if __name__ == "__main__":
    main()
