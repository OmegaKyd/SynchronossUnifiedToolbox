#!/usr/bin/env python3
"""Generate a log of all attachments across message CSVs.

Scans a messages folder for CSV files and collects every attachment along
with the sender and recipients of the message that referenced it. Results
are written to an Excel workbook and an accompanying HTML table with
thumbnails of image attachments.

Usage:
    attachment-log [--messages DIR] [--out DIR]
    python -m synchronoss_parser.attachment_log [--messages DIR] [--out DIR]

By default it expects a ``messages`` folder in the current working
directory and writes outputs under ``Attachment Log``.
"""

import argparse
import csv
import html
import logging
import os
from pathlib import Path
from typing import List, Tuple

from openpyxl import Workbook
from PIL import Image

from .render_transcripts import (
    Message,
    build_attachment_path,
    derive_attachment_day_from_csv_name,
    split_attachments,
)


AttachmentEntry = Tuple[str, str, str, str, str, str]
# (filename, sender, recipient, msg_type, direction, day)


def collect_attachments(messages_root: Path) -> List[AttachmentEntry]:
    entries: List[AttachmentEntry] = []
    for csv_file in sorted(messages_root.glob("*.csv")):
        day = derive_attachment_day_from_csv_name(csv_file) or ""
        with csv_file.open("r", encoding="utf-8", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                attachments = split_attachments(row.get("Attachments") or "")
                if not attachments:
                    continue
                msg = Message(
                    date_raw=row.get("Date") or "",
                    date_dt=None,
                    msg_type=(row.get("Type") or "").strip().lower(),
                    direction=(row.get("Direction") or "").strip().lower(),
                    attachments=attachments,
                    body=row.get("Body") or "",
                    sender=row.get("Sender") or "",
                    recipients=row.get("Recipients") or "",
                    message_id=row.get("Message ID") or "",
                    attachment_day=day,
                )
                for fname in attachments:
                    entries.append((fname, msg.sender, msg.recipients, msg.msg_type, msg.direction, day))
    return entries


def create_thumbnail(src: Path, dest: Path, size: Tuple[int, int] = (128, 128)) -> bool:
    """Create a thumbnail image.

    Returns ``True`` on success, ``False`` if the source cannot be processed
    as an image or the thumbnail cannot be written.
    """
    try:
        with Image.open(src) as img:
            img.thumbnail(size)
            dest.parent.mkdir(parents=True, exist_ok=True)
            img.save(dest)
        return True
    except Exception:
        return False


def generate_log(messages_root: Path, out_dir: Path) -> None:
    entries = collect_attachments(messages_root)
    out_dir.mkdir(parents=True, exist_ok=True)
    thumb_dir = out_dir / "thumbnails"
    thumb_dir.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.append(["filename", "sender", "recipient"])

    html_rows = []
    for fname, sender, recipient, msg_type, direction, day in entries:
        ws.append([fname, sender, recipient])
        attach_path = build_attachment_path(messages_root, msg_type, direction, day, fname)
        thumb_path = thumb_dir / msg_type / direction / day / fname
        if not create_thumbnail(attach_path, thumb_path):
            logging.warning("Failed to create thumbnail for %s", attach_path)
            thumb_path = None
        html_rows.append((fname, sender, recipient, attach_path, thumb_path))

    wb.save(out_dir / "attachment_log.xlsx")

    html_file = out_dir / "attachment_log.html"
    with html_file.open("w", encoding="utf-8") as f:
        f.write("<table>\n")
        f.write("<tr><th>filename</th><th>sender</th><th>recipient</th><th>thumbnail</th></tr>\n")
        for fname, sender, recipient, attach_path, thumb_path in html_rows:
            link = html.escape(fname)
            rel = os.path.relpath(attach_path, start=out_dir).replace(os.sep, "/")
            f.write("<tr>")
            f.write(f'<td><a href="{rel}">{link}</a></td>')
            f.write(f"<td>{html.escape(sender)}</td>")
            f.write(f"<td>{html.escape(recipient)}</td>")
            if thumb_path and thumb_path.exists():
                rel_thumb = os.path.relpath(thumb_path, start=out_dir).replace(os.sep, "/")
                f.write(f'<td><img src="{rel_thumb}" /></td>')
            else:
                f.write("<td></td>")
            f.write("</tr>\n")
        f.write("</table>\n")


def main() -> None:
    ap = argparse.ArgumentParser(description="Generate attachment log")
    ap.add_argument("--messages", default="messages", help="Folder containing message CSVs")
    ap.add_argument("--out", default="Attachment Log", help="Output folder")
    args = ap.parse_args()
    generate_log(Path(args.messages), Path(args.out))


if __name__ == "__main__":
    main()
