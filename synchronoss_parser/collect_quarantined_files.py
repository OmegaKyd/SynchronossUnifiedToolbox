#!/usr/bin/env python3
"""Extract and collect quarantined files from a Verizon backup.


Only files whose final extension is in :data:`ALLOWED_MEDIA_EXTENSIONS`
will be copied to the compiled output directory.  Files with other
extensions are ignored after extraction.

"""

from __future__ import annotations

from pathlib import Path
import os
import shutil
import sys
import tempfile
import zipfile
import logging
import re

from .collect_media import ensure_unique_name

logger = logging.getLogger(__name__)

# File extensions that are allowed to be copied after extraction
ALLOWED_MEDIA_EXTENSIONS = {
    ".jpg",
    ".png",
    ".gif",
    ".bmp",
    ".mp4",
    ".mov",
    ".mp3",
    ".wav",
}

# -------------------------------------------------------------
# Helpers for long path handling
# -------------------------------------------------------------

MAX_PATH_WIN = 260


def _win_path(path: Path) -> Path:
    """Return path with Windows long-path prefix if needed."""
    if sys.platform.startswith("win"):
        path_str = str(path)
        if not path_str.startswith("\\\\?\\"):
            return Path("\\\\?\\" + path_str)
    return path


def _shorten_dest(path: Path) -> Path:
    """Shorten final path component if total length exceeds limits."""
    try:
        limit = (
            MAX_PATH_WIN
            if sys.platform.startswith("win")
            else os.pathconf(path.anchor or "/", "PC_PATH_MAX")
        )
    except (OSError, ValueError):
        limit = 4096

    path_str = str(path)
    if len(path_str) <= limit:
        return path

    parent = path.parent
    ext = path.suffix
    stem = path.stem
    available = limit - len(str(parent)) - len(ext) - 1
    if available < 1:
        available = 1
    new_name = stem[:available] + ext
    return parent / new_name


def safe_rglob(root: Path, pattern: str):
    """Yield paths matching pattern with long path handling."""
    root = _win_path(root)
    return root.rglob(pattern)


def safe_rename(src: Path, dest: Path) -> Path:
    """Rename handling long destination paths."""
    src = _win_path(src)
    dest = _shorten_dest(_win_path(dest))
    src.rename(dest)
    return dest


def safe_copy2(src: Path, dest: Path) -> Path:
    """Copy handling long destination paths."""
    src = _win_path(src)
    dest = _shorten_dest(_win_path(dest))
    shutil.copy2(src, dest)
    return dest

# -------------------------------------------------------------
# Default paths used when running as a script
# -------------------------------------------------------------
DEFAULT_ROOT = Path("VZMOBILE")
DEFAULT_COMPILED = DEFAULT_ROOT / "Compiled Quarantine Files"

# -------------------------------------------------------------
# File type detection
# -------------------------------------------------------------

def detect_extension(path: Path) -> str | None:
    """Return file extension based on signature bytes.

    The following signatures are recognized:

    * JPEG ``\xFF\xD8\xFF``
    * PNG ``\x89PNG\r\n\x1a\n``
    * GIF ``GIF87a`` / ``GIF89a``
    * BMP ``BM``
    * PDF ``%PDF``
    * WAV ``RIFF`` + ``WAVE``
    * MP3 ``ID3``
    * MOV ``ftypqt`` (at offset 4)
    * MP4 ``ftyp`` (at offset 4)
    """

    with path.open("rb") as f:
        header = f.read(16)

    signatures = {
        (b"\xFF\xD8\xFF", 0): ".jpg",
        (b"\x89PNG\r\n\x1a\n", 0): ".png",
        (b"GIF87a", 0): ".gif",
        (b"GIF89a", 0): ".gif",
        (b"BM", 0): ".bmp",
        (b"%PDF", 0): ".pdf",
        (b"RIFF", 0): ".wav",
        (b"ID3", 0): ".mp3",
        (b"ftypqt", 4): ".mov",
    }
    for (sig, offset), ext in signatures.items():
        if header[offset : offset + len(sig)] == sig:
            if ext == ".wav" and header[8:12] != b"WAVE":
                continue
            return ext
    if len(header) >= 12 and header[4:8] == b"ftyp":
        return ".mp4"
    return None


def rename_with_extension(path: Path) -> Path:
    """Rename file to have correct extension if detectable."""
    ext = detect_extension(path)
    if ext and path.suffix.lower() != ext:
        new_path = ensure_unique_name(path.parent, path.stem + ext)
        return safe_rename(path, new_path)
    return path

# -------------------------------------------------------------
# Main processing
# -------------------------------------------------------------

def collect_quarantined_files(
    root: Path, compiled_path: Path
) -> tuple[list[Path], list[Path], int]:
    """Scan for quarantined files and copy them to ``compiled_path``.

    Returns a tuple of ``(copied, skipped, total)`` where ``copied`` contains the
    processed files, ``skipped`` contains paths that could not be handled, and
    ``total`` is the number of ``*.zip_file_*`` files discovered.
    """
    compiled_path = _shorten_dest(compiled_path)
    compiled_path.mkdir(parents=True, exist_ok=True)
    copied: list[Path] = []
    skipped: list[Path] = []

    pattern = re.compile(r"(.+)\.zip_file_(\d+)")

    all_paths = [p for p in safe_rglob(root, "*.zip_file_*") if p.is_file()]
    total = len(all_paths)

    groups: dict[Path, list[tuple[int, Path]]] = {}
    for zip_path in all_paths:
        match = pattern.match(zip_path.name)
        if not match:
            skipped.append(zip_path)
            continue

        base = zip_path.parent / match.group(1)
        ext = detect_extension(zip_path)
        if ext:
            if ext.lower() in ALLOWED_MEDIA_EXTENSIONS:
                dest_name = base.name + ext
                dest = ensure_unique_name(compiled_path, dest_name)
                dest = safe_copy2(zip_path, dest)
                copied.append(dest)
            else:
                skipped.append(zip_path)
            continue

        part = int(match.group(2))
        groups.setdefault(base, []).append((part, zip_path))

    for base, parts in groups.items():
        parts.sort(key=lambda x: x[0])
        last_num = parts[-1][0]

        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir = Path(tmpdir)
                for num, src in parts:
                    if num == last_num:
                        name = base.name + ".zip"
                    else:
                        name = f"{base.name}.z{num:02d}"
                    dest = tmpdir / name
                    safe_copy2(src, dest)

                combined = tmpdir / (base.name + "_combined.zip")
                with combined.open("wb") as outf:
                    for i in range(1, last_num):
                        part_file = tmpdir / f"{base.name}.z{i:02d}"
                        if not part_file.exists():
                            continue
                        with part_file.open("rb") as pf:
                            if i == 1:
                                sig = pf.read(4)
                                if sig != b"PK\x07\x08":
                                    outf.write(sig)
                            shutil.copyfileobj(pf, outf)
                    with (tmpdir / (base.name + ".zip")).open("rb") as pf:
                        shutil.copyfileobj(pf, outf)

                extract_dir = tmpdir / "extract"
                extract_dir.mkdir()
                with zipfile.ZipFile(combined) as zf:
                    zf.extractall(extract_dir)

                copied_before = len(copied)
                for extracted in extract_dir.rglob("*"):
                    if extracted.is_dir():
                        continue
                    fixed = rename_with_extension(extracted)
                    if fixed.suffix.lower() not in ALLOWED_MEDIA_EXTENSIONS:
                        continue
                    dest = ensure_unique_name(compiled_path, fixed.name)
                    dest = safe_copy2(fixed, dest)
                    copied.append(dest)
                if len(copied) == copied_before:
                    skipped.append(parts[-1][1])
        except zipfile.BadZipFile:
            skipped.append(parts[-1][1])
            logger.warning("Skipping invalid zip archive %s", parts[-1][1])

    return copied, skipped, total

# -------------------------------------------------------------
# CLI
# -------------------------------------------------------------

def main(root_path: Path = DEFAULT_ROOT, compiled_path: Path = DEFAULT_COMPILED) -> None:
    """CLI entry point using default paths."""
    if not root_path.exists():
        raise SystemExit(f"Root folder '{root_path}' not found.")

    files, skipped, total = collect_quarantined_files(root_path, compiled_path)
    print(
        f"Converted {len(files)} of {total} files from '{root_path}' to '{compiled_path}'."
    )
    if skipped:
        print("Skipped the following files:")
        for path in skipped:
            print(f"  {path}")


if __name__ == "__main__":
    main()
