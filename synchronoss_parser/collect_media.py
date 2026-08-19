#!/usr/bin/env python3
"""
Collect media from a Verizon Mobile backup and log metadata to Excel.

Folder structure expected:
VZMOBILE/
 └─ YYYY-MM-DD/
     └─ My Device Name/
         └─ media files

Outputs:
1. Compiled Media/   — all copied media
2. Compiled Media/compiled_media_log/compiled_media_log.xlsx — metadata
   spreadsheet stored in its own folder
"""

from pathlib import Path
import hashlib
import shutil
from fractions import Fraction
from datetime import datetime
import numbers
from PIL import Image, ExifTags
from PIL.TiffImagePlugin import IFDRational
from openpyxl import Workbook

# -------------------------------------------------------------
# Default paths used when running as a script
# -------------------------------------------------------------
DEFAULT_ROOT = Path("VZMOBILE")
DEFAULT_COMPILED = DEFAULT_ROOT / "Compiled Media"
DEFAULT_LOGFILE = DEFAULT_COMPILED / "compiled_media_log" / "compiled_media_log.xlsx"
LOGFILE = DEFAULT_LOGFILE

# Media file extensions to search
MEDIA_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".webp", ".mp4", ".mov"}

# -------------------------------------------------------------
# Helper functions
# -------------------------------------------------------------
def md5sum(path: Path) -> str:
    """Return MD5 hash of a file."""
    h = hashlib.md5()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


def normalize_exif_value(value):
    """Convert EXIF values to Excel-friendly primitive types."""
    if isinstance(value, (IFDRational, Fraction)):
        try:
            return float(value)
        except Exception:
            return str(value)
    if isinstance(value, (list, tuple)):
        return ", ".join(str(normalize_exif_value(v)) for v in value)
    if isinstance(value, bytes):
        try:
            return value.decode("utf-8", errors="replace")
        except Exception:
            return str(value)
    if isinstance(value, datetime):
        return value
    if isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, numbers.Number):
        return float(value)
    return str(value)

def extract_exif(path: Path) -> dict:
    """
    Extract EXIF data from an image (if any).
    Returns dict with human-readable keys; empty dict if none or file not image.
    """
    try:
        with Image.open(path) as img:
            raw = {ExifTags.TAGS.get(k, k): v for k, v in img.getexif().items()}
            return {k: normalize_exif_value(v) for k, v in raw.items()}
    except Exception:
        return {}

def ensure_unique_name(target_dir: Path, filename: str) -> Path:
    """Ensure unique filename inside target_dir to avoid overwrites."""
    base = Path(filename).stem
    ext = Path(filename).suffix
    counter = 0
    candidate = target_dir / filename
    while candidate.exists():
        counter += 1
        candidate = target_dir / f"{base}_{counter}{ext}"
    return candidate

# -------------------------------------------------------------
# Main processing
# -------------------------------------------------------------
def collect_media(root_path: Path, compiled_path: Path):
    """Copy media from ``root_path`` into ``compiled_path`` collecting metadata."""
    compiled_path.mkdir(exist_ok=True)

    records = []
    exif_keys = set()

    for date_dir in sorted(root_path.glob("20??-??-??")):  # match YYYY-MM-DD
        if not date_dir.is_dir():
            continue
        date_str = date_dir.name

        for device_dir in date_dir.iterdir():
            if not device_dir.is_dir():
                continue
            device_name = device_dir.name

            for media_file in device_dir.rglob("*"):
                if media_file.suffix.lower() not in MEDIA_EXTS:
                    continue

                # Copy to compiled folder
                dest = ensure_unique_name(compiled_path, media_file.name)
                shutil.copy2(media_file, dest)

                # Metadata
                exif = extract_exif(media_file)
                exif_keys.update(exif.keys())
                record = {
                    "File Name": dest.name,
                    "Date": date_str,
                    "Device": device_name,
                    "MD5": md5sum(media_file),
                }
                record.update(exif)
                for k, v in list(record.items()):
                    value = normalize_exif_value(v)
                    if not isinstance(value, (str, int, float, bool, datetime)):
                        value = str(value)
                    record[k] = value
                records.append(record)

    return records, sorted(exif_keys)

# -------------------------------------------------------------
# Excel logging
# -------------------------------------------------------------
def write_excel(records, exif_keys, logfile: Path | None = None):
    logfile = logfile or LOGFILE
    logfile.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Media Metadata"

    headers = ["File Name", "Date", "Device", "MD5"] + list(exif_keys)
    ws.append(headers)

    for rec in records:
        row = [rec.get(h, "") for h in headers]
        ws.append(row)

    wb.save(logfile)

def main(
    root_path: Path = DEFAULT_ROOT,
    compiled_path: Path = DEFAULT_COMPILED,
    logfile: Path = DEFAULT_LOGFILE,
) -> None:
    """CLI entry point using default paths."""
    if not root_path.exists():
        raise SystemExit(f"Root folder '{root_path}' not found.")

    records, exif_keys = collect_media(root_path, compiled_path)
    write_excel(records, exif_keys, logfile)
    print(
        f"Copied {len(records)} files from '{root_path}' to '{compiled_path}' and logged metadata to '{logfile}'."
    )


if __name__ == "__main__":
    main()
