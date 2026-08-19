#!/usr/bin/env python3
"""Collect message attachments and log metadata to Excel.

This script walks a Synchronoss ``messages/attachments`` folder, copies
all attachment files into a single ``Compiled Attachments`` directory,
computes basic metadata (MD5 and EXIF) and attempts to associate each
attachment with the sender and recipients of the message that referenced
it. The results are written to an Excel workbook similar to the
``collect_media`` script.

The CSV files under ``messages/`` are scanned to map attachment filenames
back to their messages. Attachment paths are constructed using utilities
from ``render_transcripts``.
"""

from __future__ import annotations

import csv
import re
import shutil
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Tuple

from openpyxl import Workbook

from .collect_media import md5sum, extract_exif, ensure_unique_name
from .render_transcripts import (
    build_attachment_path,
    build_contact_lookup,
    derive_attachment_day_from_csv_name,
    parse_csv_date,
    split_attachments,
)

# ---------------------------------------------------------------------------
# Default paths
# ---------------------------------------------------------------------------
DEFAULT_ATTACHMENTS_ROOT = Path("messages") / "attachments"
DEFAULT_COMPILED = Path("Compiled Attachments")
DEFAULT_LOGFILE = DEFAULT_COMPILED / "compiled_attachment_log" / "compiled_attachment_log.xlsx"

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def sanitize_filename_component(text: str) -> str:
    """Return ``text`` stripped of common filesystem-unsafe characters."""
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "", text)
    return cleaned.strip().rstrip(".")

# ---------------------------------------------------------------------------
# Helper to map attachment files to message metadata
# ---------------------------------------------------------------------------

def build_metadata_index(
    messages_root: Path, contact_lookup: Callable[[str], str] = lambda x: x
) -> Dict[Path, Dict[str, str]]:
    """Scan message CSV files and map attachment paths to metadata."""
    index: Dict[Path, Dict[str, str]] = {}
    for csv_file in sorted(messages_root.glob("*.csv")):
        day = derive_attachment_day_from_csv_name(csv_file)
        with csv_file.open("r", encoding="utf-8", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                attachments = split_attachments(row.get("Attachments") or "")
                if not attachments:
                    continue
                msg_type = (row.get("Type") or "").strip().lower()
                direction = (row.get("Direction") or "").strip().lower()
                date = (row.get("Date") or "").strip()
                sender = contact_lookup((row.get("Sender") or "").strip())

                raw_recip = (row.get("Recipients") or "")
                recip_parts: List[str] = []
                for part in raw_recip.replace(",", ";").split(";"):
                    p = part.strip()
                    if p:
                        recip_parts.append(contact_lookup(p))
                recipient = "; ".join(recip_parts)

                for fname in attachments:
                    path = build_attachment_path(
                        messages_root, msg_type, direction, day or "", fname
                    )
                    index[path.resolve()] = {
                        "Date": date,
                        "Sender": sender,
                        "Recipient": recipient,
                    }
    return index

# ---------------------------------------------------------------------------
# Main processing
# ---------------------------------------------------------------------------

def collect_attachments(
    attachments_root: Path,
    compiled_path: Path,
    contacts_xlsx: str | Path | None = None,
) -> Tuple[List[Dict[str, str]], List[str]]:
    """Copy attachments from ``attachments_root`` into ``compiled_path``.

    Returns a tuple ``(records, exif_keys)`` where ``records`` is a list of
    metadata dictionaries and ``exif_keys`` is the sorted list of all EXIF
    keys encountered.
    """
    compiled_path.mkdir(exist_ok=True)

    messages_root = attachments_root.parent
    lookup = build_contact_lookup(str(contacts_xlsx) if contacts_xlsx else None)
    metadata_index = build_metadata_index(messages_root, lookup)

    records: List[Dict[str, str]] = []
    exif_keys: set[str] = set()

    for file in attachments_root.rglob("*"):
        if not file.is_file():
            continue

        meta = metadata_index.get(file.resolve(), {})

        # Keep the original filename; only uniquify on collision
        dest = ensure_unique_name(compiled_path, file.name)
        shutil.copy2(file, dest)

        exif = extract_exif(file)
        exif_keys.update(exif.keys())

        record = {
            "File Name": dest.name,
            "Original Name": file.name,
            "Date": meta.get("Date", ""),
            "Sender": meta.get("Sender", ""),
            "Recipient": meta.get("Recipient", ""),
            "MD5": md5sum(file),
        }
        record.update(exif)
        records.append(record)

    return records, sorted(exif_keys)

# ---------------------------------------------------------------------------
# Excel logging
# ---------------------------------------------------------------------------

def write_excel(records: Iterable[Dict[str, str]], exif_keys: Iterable[str], logfile: Path | None = None) -> None:
    logfile = logfile or DEFAULT_LOGFILE
    logfile.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Attachment Metadata"

    headers = ["File Name", "Original Name", "Date", "Sender", "Recipient", "MD5"] + list(exif_keys)
    ws.append(headers)

    for rec in records:
        row = [rec.get(h, "") for h in headers]
        ws.append(row)

    wb.save(logfile)

# ---------------------------------------------------------------------------
# Command line interface
# ---------------------------------------------------------------------------

def main(attachments_root: Path | str = DEFAULT_ATTACHMENTS_ROOT, compiled_path: Path | str = DEFAULT_COMPILED, contacts_xlsx: Path | str | None = None, logfile: Path | None = None) -> None:
    attachments_root = Path(attachments_root)
    compiled_path = Path(compiled_path)
    if not attachments_root.exists():
        raise SystemExit(f"Attachments folder '{attachments_root}' not found.")

    # Keep the Excel log next to the compiled files (not a relative path under CWD)
    if logfile is None:
        logfile = compiled_path / "compiled_attachment_log" / "compiled_attachment_log.xlsx"

    records, exif_keys = collect_attachments(attachments_root, compiled_path, contacts_xlsx)
    write_excel(records, exif_keys, logfile)
    print(
        f"Copied {len(records)} files from '{attachments_root}' to '{compiled_path}' "
        f"and logged metadata to '{logfile}'."
    )


if __name__ == "__main__":
    main()
