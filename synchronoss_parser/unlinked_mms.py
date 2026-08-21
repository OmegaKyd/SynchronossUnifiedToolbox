#!/usr/bin/env python3
"""
Identify MMS attachment files that are present on disk but not referenced by a
real-extension token in any message CSV.

Synchronoss often stores media under names like ``0`` (no extension) and only
references them via SMIL placeholders (e.g. ``null.smi;0;1``). Those tokens are
not reliable message↔file attributions, so the files are surfaced here as
**unlinked** media — dated by the folder (upload date) and not tied to a
specific message. Extensionless files are typed via magic-byte detection.
"""

from __future__ import annotations

import csv
import logging
import os
import re
import shutil
from pathlib import Path
from typing import Dict, List, Optional, Tuple

logger = logging.getLogger("synchronoss.unlinked_mms")

CSV_PATTERN = re.compile(r"\d{8}\.csv$", re.IGNORECASE)
MEDIA_RE = re.compile(r"/mms/(in|out)/([^/]+)/([^/]+)$", re.IGNORECASE)
SMIL_PREFIXES = ("smil", "null", "text0")
SMIL_EXTS = {".smi", ".sml", ".txt"}


def detect_media_type(filepath: Path) -> Optional[str]:
    """
    Detect media type from file header magic bytes (no external libraries).
    Returns an extension string (e.g. ``.jpg``) or None.
    """
    try:
        with open(filepath, "rb") as fh:
            h = fh.read(32)
    except OSError:
        return None

    if len(h) < 4:
        return None

    if h[:3] == b"\xff\xd8\xff":
        return ".jpg"
    if h[:8] == b"\x89PNG\r\n\x1a\n":
        return ".png"
    if h[:6] in (b"GIF87a", b"GIF89a"):
        return ".gif"
    if h[:2] == b"BM":
        return ".bmp"
    if h[:4] == b"RIFF" and h[8:12] == b"WEBP":
        return ".webp"
    if h[4:8] == b"ftyp":
        brand = h[8:12]
        if brand == b"M4A ":
            return ".m4a"
        if brand[:3] in (b"3gp", b"3g2"):
            return ".3gp"
        return ".mp4"
    if h[:5] == b"#!AMR":
        return ".amr"
    if h[:4] == b"\x1aE\xdf\xa3":
        return ".mkv"
    if h[:4] == b"OggS":
        return ".ogg"
    if h[:3] == b"ID3":
        return ".mp3"
    if h[:2] in (b"\xff\xfb", b"\xff\xf3", b"\xff\xf2"):
        return ".mp3"
    if h[:4] == b"fLaC":
        return ".flac"
    return None


def _open_csv_rows(csv_path: Path) -> List[Dict[str, str]]:
    rows: List[Dict[str, str]] = []
    try:
        with open(csv_path, "r", encoding="utf-8-sig", errors="replace", newline="") as fh:
            reader = csv.DictReader(fh)
            for row in reader:
                rows.append({(k or "").strip(): (v or "") for k, v in row.items()})
    except (OSError, csv.Error, UnicodeError) as e:
        logger.warning("Could not read %s: %s", csv_path, e)
    return rows


def find_unlinked_mms(messages_root: Path) -> List[Dict[str, str]]:
    """
    Scan ``messages_root`` for MMS folder media that is never named by a
    real-extension attachment token in any daily CSV.

    Returns a list of dicts sorted by direction, upload date, filename::

        {
          "upload_date": "YYYY-MM-DD",
          "direction": "in" | "out",
          "filename": str,
          "detected_type": str,   # e.g. ".jpg" or ".jpg (by magic bytes)" or "unknown"
          "source_path": str,     # absolute path to the original file
        }
    """
    messages_root = Path(messages_root)
    if not messages_root.is_dir():
        return []

    referenced: set = set()  # real-extension filenames referenced by any MMS row
    media: List[Tuple[str, str, str, Path]] = []  # direction, date_folder, basename, path

    # Collect referenced tokens from CSVs and media files from the tree
    for path in messages_root.rglob("*"):
        if not path.is_file():
            continue
        name = path.name
        norm = str(path).replace("\\", "/")

        if CSV_PATTERN.search(name):
            for row in _open_csv_rows(path):
                if (row.get("Type") or "").lower() != "mms":
                    continue
                for tok in (t.strip() for t in (row.get("Attachments") or "").split(";") if t.strip()):
                    low = tok.lower()
                    ext = os.path.splitext(low)[1]
                    if low.startswith(SMIL_PREFIXES) or ext in SMIL_EXTS:
                        continue
                    if ext:  # real-extension token → actual media file name
                        referenced.add(tok)
            continue

        m = MEDIA_RE.search(norm)
        if m:
            media.append((m.group(1).lower(), m.group(2), m.group(3), path))

    results: List[Dict[str, str]] = []
    for direction, date_folder, basename, fpath in media:
        if basename in referenced:
            continue
        ext = os.path.splitext(basename)[1].lower()
        if ext:
            detected = ext
        else:
            det = detect_media_type(fpath)
            detected = f"{det} (by magic bytes)" if det else "unknown (magic bytes)"
        results.append(
            {
                "upload_date": date_folder,
                "direction": direction,
                "filename": basename,
                "detected_type": detected,
                "source_path": str(fpath.resolve()),
            }
        )

    results.sort(key=lambda r: (r["direction"], r["upload_date"], r["filename"]))
    return results


def materialize_unlinked_mms(
    unlinked: List[Dict[str, str]],
    dest_dir: Path,
    compiled_attachments: Optional[Path] = None,
) -> List[Dict[str, str]]:
    """
    Ensure each unlinked file is reachable under *dest_dir* (or already present
    in *compiled_attachments*) and add a ``rel_path`` key for HTML linking
    (relative to the case ``parsed_output/`` folder).

    Prefer an existing copy in Compiled Attachments (matched by original
    filename). Otherwise copy into ``dest_dir`` (typically ``Unlinked MMS/``).
    """
    dest_dir = Path(dest_dir)
    dest_dir.mkdir(parents=True, exist_ok=True)

    # Map original basename → relative path under compiled attachments
    compiled_map: Dict[str, str] = {}
    if compiled_attachments and Path(compiled_attachments).is_dir():
        for f in Path(compiled_attachments).iterdir():
            if f.is_file():
                # Prefer exact original name; uniquified names still work for display
                compiled_map[f.name] = f"Compiled Attachments/{f.name}"

    enriched: List[Dict[str, str]] = []
    for row in unlinked:
        item = dict(row)
        fname = row["filename"]
        rel = compiled_map.get(fname)
        if not rel:
            # Try stem matches for uniquified copies (name_1.ext)
            stem = Path(fname).stem
            suffix = Path(fname).suffix
            for cname, crel in compiled_map.items():
                if cname == fname or (cname.startswith(stem + "_") and cname.endswith(suffix)):
                    rel = crel
                    break
        if not rel:
            src = Path(row["source_path"])
            if src.is_file():
                # Give extensionless files a detected extension when copying
                out_name = fname
                det = row.get("detected_type", "")
                if not Path(fname).suffix and det.startswith("."):
                    # e.g. ".jpg (by magic bytes)" → ".jpg"
                    magic_ext = det.split()[0]
                    if magic_ext.startswith(".") and len(magic_ext) <= 5:
                        out_name = fname + magic_ext
                dest = dest_dir / out_name
                counter = 1
                while dest.exists():
                    dest = dest_dir / f"{Path(out_name).stem}_{counter}{Path(out_name).suffix}"
                    counter += 1
                try:
                    shutil.copy2(src, dest)
                    rel = f"Unlinked MMS/{dest.name}"
                except OSError as e:
                    logger.warning("Could not copy unlinked media %s: %s", src, e)
                    rel = ""
            else:
                rel = ""
        item["rel_path"] = rel
        enriched.append(item)
    return enriched


def write_unlinked_excel(rows: List[Dict[str, str]], out_path: Path) -> Path:
    """Write Unlinked MMS.xlsx for offline review."""
    try:
        from openpyxl import Workbook
    except ImportError:
        logger.warning("openpyxl not available – skipping Unlinked MMS Excel")
        return out_path

    wb = Workbook()
    ws = wb.active
    ws.title = "Unlinked MMS"
    ws.append(
        [
            "Upload Date",
            "Direction",
            "Filename",
            "Detected Type",
            "Relative Path",
            "Source Path",
        ]
    )
    for r in rows:
        ws.append(
            [
                r.get("upload_date", ""),
                r.get("direction", ""),
                r.get("filename", ""),
                r.get("detected_type", ""),
                r.get("rel_path", ""),
                r.get("source_path", ""),
            ]
        )
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    logger.info("Wrote Unlinked MMS Excel (%d rows) → %s", len(rows), out_path)
    return out_path
