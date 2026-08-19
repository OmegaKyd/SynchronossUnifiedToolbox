"""Utility helpers for Synchronoss scripts."""

from __future__ import annotations

import hashlib
import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Optional


def normalize_phone_number(value: str) -> str:
    """Return a canonical phone number (prefer 10 digits for US/Canada).

    All non-digit characters are stripped.  A leading US/Canada country code
    ``1`` is removed when the result would otherwise be eleven digits.

    Examples
    --------
    >>> normalize_phone_number("+1 111-222-3333")
    '1112223333'
    >>> normalize_phone_number("(111) 222-3333")
    '1112223333'
    >>> normalize_phone_number("+12223334444")
    '2223334444'
    >>> normalize_phone_number("1112223333")
    '1112223333'
    """
    if not value:
        return ""
    digits = "".join(ch for ch in str(value) if ch.isdigit())
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    return digits


def md5sum(path: Path) -> str:
    """Return MD5 hex digest of a file."""
    h = hashlib.md5()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


def sha256sum(path: Path) -> str:
    """Return SHA-256 hex digest of a file."""
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


def ensure_unique_name(target_dir: Path, filename: str) -> Path:
    """Return a unique path inside *target_dir* that does not overwrite existing files."""
    base = Path(filename).stem
    ext = Path(filename).suffix
    candidate = target_dir / filename
    counter = 0
    while candidate.exists():
        counter += 1
        candidate = target_dir / f"{base}_{counter}{ext}"
    return candidate


def find_gpg() -> Optional[Path]:
    """Locate the ``gpg`` executable.

    Checks ``PATH`` first, then common GPG4Win / GnuPG install locations on Windows.
    Returns ``None`` if not found.
    """
    which = shutil.which("gpg")
    if which:
        return Path(which)

    if sys.platform.startswith("win"):
        candidates = [
            Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "GnuPG" / "bin" / "gpg.exe",
            Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "GnuPG" / "bin" / "gpg.exe",
            Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "Gpg4win" / "bin" / "gpg.exe",
            Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Gpg4win" / "bin" / "gpg.exe",
        ]
        for p in candidates:
            if p.is_file():
                return p
    return None


def gpg_available() -> bool:
    """Return True if a usable ``gpg`` binary is present."""
    return find_gpg() is not None


def decrypt_gpg_file(
    gpg_path: Path,
    output_path: Path,
    passphrase: str,
    *,
    gpg_exe: Optional[Path] = None,
) -> Path:
    """Decrypt a single ``.gpg`` file using the system ``gpg`` binary.

    Uses ``--passphrase-fd 0`` so the passphrase never appears on the command line.
    Raises ``RuntimeError`` on failure.
    """
    exe = gpg_exe or find_gpg()
    if exe is None:
        raise FileNotFoundError(
            "gpg executable not found. Install GPG4Win from https://www.gpg4win.org"
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)

    cmd = [
        str(exe),
        "--batch",
        "--yes",
        "--passphrase-fd",
        "0",
        "--output",
        str(output_path),
        "--decrypt",
        str(gpg_path),
    ]

    proc = subprocess.run(
        cmd,
        input=passphrase + "\n",
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if proc.returncode != 0:
        msg = (proc.stderr or proc.stdout or "").strip()
        raise RuntimeError(f"Failed to decrypt {gpg_path.name}: {msg}")

    if not output_path.is_file():
        raise RuntimeError(f"Decryption reported success but output file missing: {output_path}")

    return output_path
