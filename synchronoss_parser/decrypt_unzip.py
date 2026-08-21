from __future__ import annotations

"""Utilities for decrypting and unpacking Synchronoss archives."""

from pathlib import Path
import shutil
import subprocess
import zipfile


def decrypt_and_unzip(
    archive_path: str | Path,
    password: str,
    output_dir: str | Path | None = None,
    cleanup: bool = True,
) -> list[Path]:
    """Decrypt ``archive_path`` and expand all nested archives.

    Parameters
    ----------
    archive_path:
        Path to the top level ``.zip`` archive. ``str`` paths are accepted.
    password:
        Passphrase used to decrypt any ``.gpg`` files encountered.
    output_dir:
        Optional destination directory.  Defaults to ``archive_path.stem`` in
        the same directory as ``archive_path``. ``str`` paths are accepted.
    cleanup:
        Whether to remove processed archives.  Defaults to ``True``. The
        original ``archive_path`` is always preserved.

    Returns
    -------
    list[Path]
        A list of all extracted file paths.
    """

    archive_path = Path(archive_path).expanduser()
    if output_dir is not None:
        output_dir = Path(output_dir).expanduser()

    if output_dir is None:
        output_dir = archive_path.parent / archive_path.stem

    with zipfile.ZipFile(archive_path) as zf:
        zf.extractall(output_dir)

    processed: set[Path] = set()
    dirs_to_walk: list[Path] = [output_dir]

    while dirs_to_walk:
        current = dirs_to_walk.pop()
        for item in current.iterdir():
            if item.is_dir():
                dirs_to_walk.append(item)
                continue

            if item.suffix == ".zip" and item not in processed:
                dest = item.with_suffix("")
                with zipfile.ZipFile(item) as zf:
                    zf.extractall(dest)
                if cleanup:
                    item.unlink(missing_ok=True)
                processed.add(item)
                dirs_to_walk.append(dest)
                continue

            if item.suffix == ".gpg" and item not in processed:
                if shutil.which("gpg") is None:
                    raise FileNotFoundError("'gpg' executable is required to decrypt files")
                decrypted_path = item.with_suffix("")
                cmd = [
                    "gpg",
                    "--batch",
                    "--yes",
                    "--passphrase",
                    password,
                    "-o",
                    str(decrypted_path),
                    "-d",
                    str(item),
                ]
                proc = subprocess.run(cmd, capture_output=True, text=True)
                if proc.returncode != 0:
                    msg = proc.stderr.strip() or proc.stdout.strip()
                    raise RuntimeError(f"Failed to decrypt {item}: {msg}")
                if zipfile.is_zipfile(decrypted_path):
                    dest = decrypted_path.with_suffix("")
                    with zipfile.ZipFile(decrypted_path) as zf:
                        zf.extractall(dest)
                    if cleanup:
                        decrypted_path.unlink(missing_ok=True)
                        item.unlink(missing_ok=True)
                    processed.update({item, decrypted_path})
                    dirs_to_walk.append(dest)
                else:
                    if cleanup:
                        item.unlink(missing_ok=True)
                    processed.add(item)
                continue

    files = [
        p
        for p in output_dir.rglob("*")
        if p.is_file() and not p.name.endswith(".zip") and not p.name.endswith(".gpg")
    ]
    files.sort()
    return files
