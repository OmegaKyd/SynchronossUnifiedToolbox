#!/usr/bin/env python3
"""Helper script for building PyInstaller executables.

Runs PyInstaller with the provided ``.spec`` files to produce standalone
executables. Ensure PyInstaller is installed and available on the PATH
before executing this script.
"""

from pathlib import Path
import subprocess

HERE = Path(__file__).resolve().parent
SPEC_FILES = [HERE / 'toolbox_gui.spec']


def main() -> None:
    for spec in SPEC_FILES:
        if spec.exists():
            subprocess.run(['pyinstaller', str(spec)], check=True)


if __name__ == '__main__':
    main()
