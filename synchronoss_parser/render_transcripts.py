#!/usr/bin/env python3
"""
Chat Transcript Renderer

Reads message CSVs from a folder structure like:

messages/
  20240120.csv
  20241121.csv
  ...
  attachments/
    mms/
      in/
        2024-01-20/
          <files>
      out/
        2024-01-20/
          <files>
    rcs/
      in/
        2024-11-21/
          <files>
      out/
        2024-11-21/
          <files>

Each CSV must have columns: Date, Type, Direction, Attachments, Body, Sender, Recipients, Message ID

This script generates one HTML transcript per CSV (chat-bubble style), plus an index.html.
Attachments are embedded inline when possible (images/audio/video) and otherwise linked.
The attachment lookup path is:
  messages/attachments/{Type}/{Direction}/{YYYY-MM-DD}/{AttachmentFileName}
where YYYY-MM-DD is derived from the CSV filename (e.g., 20241121.csv -> 2024-11-21).

Usage:
  render-transcripts --in messages --out transcripts [--contacts-xlsx contacts.xlsx]

Notes:
- HTML keeps relative links to your existing attachments (no copying).
- Dates are sorted chronologically using best-effort parsing; the raw Date string is also shown.
- "out" messages are right-aligned (sent), "in" are left-aligned (received).
- Safe to run multiple times; outputs are overwritten.
"""

import argparse
import csv
import html
import json
import os
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

import pandas as pd

from .utils import normalize_phone_number


# ------------------------- Config & Utilities -------------------------

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp"}
VIDEO_EXTS = {".mp4", ".webm", ".ogg", ".mov", ".m4v"}
AUDIO_EXTS = {".mp3", ".wav", ".m4a", ".aac", ".flac", ".ogg"}
INLINE_TEXT_EXTS = {".vcard", ".vcf"}  # small text-like files we might show inline

CSV_DATE_FROM_FILENAME_FMT = "%Y%m%d"
ATTACHMENT_FOLDER_DATE_FMT = "%Y-%m-%d"

CSS_STYLES = """
:root {
  --bg: #1e293b;           /* slate-800 */
  --panel: #1f2937;        /* gray-800 */
  --text: #f3f4f6;         /* gray-100 */
  --muted: #cbd5e1;        /* slate-300 */
  /* Stronger green + pure white text for high contrast on outgoing bubbles */
  --sent: #16a34a;         /* green-600 – solid, readable */
  --sent-contrast: #ffffff;
  --sent-meta: #dcfce7;    /* green-100 – soft white for timestamps */
  --recv: #374151;         /* gray-700 */
  --bubble-radius: 16px;
  --max-width: 980px;
}
* { box-sizing: border-box; }
body {
  margin: 0; padding: 0; background: var(--bg); color: var(--text);
  font: 15px/1.5 system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
}
.header {
  position: sticky; top: 0; z-index: 10; background: linear-gradient(180deg, rgba(17,24,39,0.95), rgba(17,24,39,0.7));
  backdrop-filter: blur(6px); border-bottom: 1px solid #1f2937; padding: 10px 16px;
}
.container { max-width: var(--max-width); margin: 0 auto; padding: 12px 16px 80px; }
.thread-meta { color: var(--muted); font-size: 12px; margin-top: 2px; }
.search-bar {
  margin-top: 8px;
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 12px;
}
.search-input {
  width: 100%;
  max-width: 400px;
  padding: 6px 8px;
  border-radius: 8px;
  border: 1px solid #374151;
  background: var(--panel);
  color: var(--text);
}
.att-filter {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  color: var(--muted);
  font-size: 13px;
  cursor: pointer;
  user-select: none;
  white-space: nowrap;
}
.att-filter input {
  accent-color: #16a34a;
  width: 15px;
  height: 15px;
  cursor: pointer;
}

.message {
  display: flex; margin: 8px 0; gap: 10px; align-items: flex-end;
}
.bubble {
  max-width: 74%; padding: 10px 12px; border-radius: var(--bubble-radius);
  word-wrap: break-word; overflow-wrap: anywhere; box-shadow: 0 2px 10px rgba(0,0,0,0.2);
}
.sent { margin-left: auto; justify-content: flex-end; }
.sent .bubble {
  background: var(--sent);
  color: var(--sent-contrast);
}
/* Force readable text on every element inside the green bubble */
.sent .bubble .body-text,
.sent .bubble .sender,
.sent .bubble .missing {
  color: var(--sent-contrast);
}
.sent .bubble .meta {
  color: var(--sent-meta);
  text-align: right;
}
.sent .sender { text-align: right; }
.sent .bubble .attachment a {
  color: #dbeafe;          /* light blue – still readable on green */
}
.received .bubble { background: var(--recv); color: var(--text); }

.sender { font-size: 16px; font-weight: 700; margin-bottom: 4px; }
.meta { font-size: 11px; color: var(--muted); margin-top: 6px; }
.body-text { white-space: pre-wrap; }
.missing { color: var(--muted); font-style: italic; }
.attachments { margin-top: 8px; display: grid; gap: 8px; }
.attachment img { max-width: 100%; border-radius: 12px; display: block; }
.attachment video, .attachment audio { width: 100%; outline: none; }
.attachment a { color: #93c5fd; text-decoration: none; word-break: break-all; }
.attachment a:hover { text-decoration: underline; }

.day-divider { text-align: center; margin: 18px 0; color: var(--muted); font-size: 12px; }
.footer { position: fixed; bottom: 0; left: 0; right: 0; padding: 8px 16px; background: linear-gradient(0deg, rgba(17,24,39,0.95), rgba(17,24,39,0.6)); border-top: 1px solid #1f2937; }
.footer a { color: #93c5fd; text-decoration: none; }
.footer a:hover { text-decoration: underline; }
.code { font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; font-size: 12px; }
.print-btn {
  margin-left: auto;
  padding: 6px 14px;
  border-radius: 8px;
  border: 1px solid #4b5563;
  background: var(--panel);
  color: var(--text);
  font-size: 13px;
  cursor: pointer;
}
.print-btn {
  margin-left: auto;
  padding: 6px 14px;
  border-radius: 8px;
  border: 1px solid #4b5563;
  background: var(--panel);
  color: var(--text);
  font-size: 13px;
  cursor: pointer;
}
.print-btn:hover { background: #374151; border-color: #6b7280; }

/* ----- Print layout -----
   Header becomes a normal block at the top of page 1 (not sticky).
   Footer / search / print controls are hidden.
   Single page footer via @page: chat name on the left, page numbers on the right.
   Turn OFF browser "Headers and footers" so the file:// path does not appear. */
@media print {
  @page {
    margin: 15mm 15mm 20mm 15mm;
    @bottom-right {
      content: "Page " counter(page) " of " counter(pages);
      font-size: 9pt;
      color: #4b5563;
    }
  }
  body {
    background: #ffffff !important;
    color: #111827 !important;
    font-size: 11pt;
  }
  .header {
    position: static !important;
    background: #ffffff !important;
    backdrop-filter: none !important;
    border-bottom: 1px solid #d1d5db !important;
    padding: 0 0 12px 0 !important;
    margin-bottom: 12px;
    color: #111827 !important;
  }
  .header .thread-meta { color: #4b5563 !important; }
  .search-bar,
  .print-btn,
  .footer {
    display: none !important;
  }
  .container {
    max-width: none !important;
    padding: 0 !important;
    margin: 0 !important;
  }
  .message { break-inside: avoid; page-break-inside: avoid; }
  .bubble {
    box-shadow: none !important;
    border: 1px solid #d1d5db;
  }
  .sent .bubble {
    background: #dcfce7 !important;
    color: #14532d !important;
  }
  .sent .bubble .body-text,
  .sent .bubble .sender,
  .sent .bubble .missing,
  .sent .bubble .meta {
    color: #14532d !important;
  }
  .received .bubble {
    background: #f3f4f6 !important;
    color: #111827 !important;
  }
  .day-divider {
    color: #6b7280 !important;
    break-after: avoid;
  }
  .attachment img,
  .attachment video {
    max-width: 3.5in !important;
    max-height: 3.5in !important;
    width: auto !important;
    height: auto !important;
  }
  a { color: #111827 !important; text-decoration: underline; }
}
"""

INDEX_CSS = """
body { background:#1e293b; color:#f3f4f6; font: 15px/1.5 system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; }
.container { max-width: 900px; margin: 0 auto; padding: 24px 16px; }
h1 { margin: 0 0 8px; font-size: 24px; }
.subtitle { color:#cbd5e1; margin-bottom: 16px; font-size: 13px; }
.search-bar { margin-bottom: 16px; }
.search-input { width:100%; max-width:400px; padding:6px 8px; border-radius:8px; border:1px solid #374151; background:#1f2937; color:#f3f4f6; }
.list { display: grid; gap: 10px; }
.item { background:#1f2937; border:1px solid #374151; border-radius: 12px; padding: 12px; }
.item a { color:#93c5fd; text-decoration:none; font-weight:600; }
.item a:hover { text-decoration: underline; }
.meta { color:#cbd5e1; font-size: 12px; margin-top: 4px; }
"""

@dataclass
class Message:
    date_raw: str
    date_dt: Optional[datetime]
    msg_type: str
    direction: str
    attachments: List[str]
    body: str
    sender: str
    recipients: str
    message_id: str
    attachment_day: Optional[str] = None


def parse_csv_date(value: str) -> Optional[datetime]:
    if not value:
        return None
    s = value.strip()
    # Try ISO-8601 with Z
    try:
        if s.endswith("Z"):
            return datetime.fromisoformat(s.replace("Z", "+00:00"))
    except Exception:
        pass
    # Try plain fromisoformat
    try:
        return datetime.fromisoformat(s)
    except Exception:
        pass
    # Try common formats
    fmts = [
        "%Y-%m-%d %H:%M:%S",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %I:%M:%S %p",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%dT%H:%M:%S.%f%z",
        "%Y-%m-%dT%H:%M:%S%z",
    ]
    for f in fmts:
        try:
            return datetime.strptime(s, f)
        except Exception:
            continue
    # Fallback: epoch seconds or ms?
    try:
        n = int(s)
        if n > 10_000_000_000:  # likely ms
            return datetime.fromtimestamp(n/1000, tz=timezone.utc)
        return datetime.fromtimestamp(n, tz=timezone.utc)
    except Exception:
        return None


def split_attachments(field: str) -> List[str]:
    if not field:
        return []
    s = field.strip()
    # Try JSON list first
    if (s.startswith("[") and s.endswith("]")) or (s.startswith("{") and s.endswith("}")):
        try:
            parsed = json.loads(s)
            if isinstance(parsed, list):
                return [str(x).strip() for x in parsed if str(x).strip()]
            if isinstance(parsed, dict) and "files" in parsed:
                return [str(x).strip() for x in parsed["files"] if str(x).strip()]
        except Exception:
            pass
    # Fallback: split on common delimiters (semicolon strongest for your data)
    parts = []
    for chunk in s.replace("|", ";").replace(",", ";").split(";"):
        c = chunk.strip()
        if c:
            parts.append(c)
    return parts


def derive_attachment_day_from_csv_name(csv_path: Path) -> Optional[str]:
    """Convert 20241121.csv -> 2024-11-21"""
    try:
        stem = csv_path.stem  # e.g., 20241121
        dt = datetime.strptime(stem, CSV_DATE_FROM_FILENAME_FMT)
        return dt.strftime(ATTACHMENT_FOLDER_DATE_FMT)
    except Exception:
        return None


def classify_ext(path: Path) -> str:
    ext = path.suffix.lower()
    if ext in IMAGE_EXTS:
        return "image"
    if ext in VIDEO_EXTS:
        return "video"
    if ext in AUDIO_EXTS:
        return "audio"
    if ext in INLINE_TEXT_EXTS:
        return "inline_text"
    return "other"


def safe_text(s: Optional[str]) -> str:
    if s is None:
        return ""
    return html.escape(str(s))


def build_contact_lookup(xlsx_path: Optional[str]) -> Callable[[str], str]:
    """Return a lookup function mapping phone numbers to contact names.

    Returns the contact name when known, otherwise the original input
    (typically a phone number).  For display that should also show the
    number, use :func:`format_contact_label`.
    """
    mapping: Dict[str, str] = {}
    if xlsx_path:
        try:
            df = pd.read_excel(xlsx_path)
            for _, row in df.iterrows():
                first = str(row.get("firstname") or "").strip()
                last = str(row.get("lastname") or "").strip()
                numbers = str(row.get("phone_numbers") or "").split(";")
                name = f"{first} {last}".strip()
                if not name:
                    continue
                for num in numbers:
                    digits = normalize_phone_number(num)
                    if digits:
                        mapping[digits] = name
        except Exception:
            pass

    def lookup(number: str) -> str:
        digits = normalize_phone_number(number)
        return mapping.get(digits, number)

    return lookup


def format_contact_label(raw: str, contact_lookup: Callable[[str], str] = lambda x: x) -> str:
    """Format a party for display: ``Name (digits)`` when a contact is known.

    Examples
    --------
    - known contact  → ``"Tractor Supply (5551234567)"``
    - unknown number → ``"5551234567"``
    - empty          → ``""``
    """
    raw = (raw or "").strip()
    if not raw:
        return ""
    digits = normalize_phone_number(raw)
    name = contact_lookup(raw)
    # lookup returns a real name when it differs from both the raw input
    # and the normalised digits
    if name and digits and name != raw and name != digits:
        return f"{name} ({digits})"
    if digits:
        return digits
    return name or raw


def build_attachment_path(messages_root: Path, msg_type: str, direction: str, day_str: str, filename: str) -> Path:
    return messages_root / "attachments" / msg_type / direction / day_str / filename


def relpath_for_html(from_file: Path, to_target: Path) -> str:
    try:
        return os.path.relpath(to_target, start=from_file.parent).replace(os.sep, "/")
    except Exception:
        return str(to_target).replace(os.sep, "/")


def load_messages_from_csv(csv_file: Path, contact_lookup: Callable[[str], str] = lambda x: x) -> List[Message]:
    """Load messages, keeping raw phone numbers on each Message.

    Contact names are resolved later at display time via
    :func:`format_contact_label` so both name and number can be shown.
    ``contact_lookup`` is accepted for API compatibility but is not applied
    here.
    """
    msgs: List[Message] = []
    day_folder = derive_attachment_day_from_csv_name(csv_file)
    with csv_file.open("r", encoding="utf-8", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            date_raw = (row.get("Date") or "").strip()
            date_dt = parse_csv_date(date_raw)
            msg_type = (row.get("Type") or "").strip().lower()
            direction = (row.get("Direction") or "").strip().lower()
            attachments_field = row.get("Attachments")
            attachments = split_attachments(attachments_field) if attachments_field else []
            body = row.get("Body") or ""
            # Keep the original phone number (normalised when possible) so the
            # number remains available for display alongside the contact name.
            raw_sender = (row.get("Sender") or "").strip()
            sender = normalize_phone_number(raw_sender) or raw_sender
            raw_recip = row.get("Recipients") or ""
            recip_parts = []
            for part in raw_recip.replace(",", ";").split(";"):
                p = part.strip()
                if p:
                    recip_parts.append(normalize_phone_number(p) or p)
            recipients = "; ".join(recip_parts)
            message_id = row.get("Message ID") or ""
            msgs.append(
                Message(
                    date_raw,
                    date_dt,
                    msg_type,
                    direction,
                    attachments,
                    body,
                    sender,
                    recipients,
                    message_id,
                    day_folder,
                )
            )
    # Sort chronologically with stable fallback to raw string
    msgs.sort(key=lambda m: (m.date_dt or datetime.max.replace(tzinfo=timezone.utc), m.date_raw))
    return msgs


# ------------------------- Grouping -------------------------

def sanitize_participants(participants: Tuple[str, ...]) -> str:
    if not participants:
        return "chat"
    cleaned = ["".join(ch for ch in p if ch.isalnum()) for p in participants]
    cleaned = [c or "unknown" for c in cleaned]
    return "-".join(cleaned) or "chat"


def group_messages_by_chat(messages: List[Message], target: str) -> Dict[Tuple[str, ...], List[Message]]:
    groups: Dict[Tuple[str, ...], List[Message]] = {}
    for m in messages:
        if m.msg_type in {"sms", "mms", "rcs"}:
            if m.direction == "out" and not m.sender:
                m.sender = target
            if m.direction == "in" and not m.recipients:
                m.recipients = target
        recips: List[str] = []
        if m.recipients:
            for part in m.recipients.replace(",", ";").split(";"):
                p = part.strip()
                if p:
                    recips.append(p)
        participants = set(recips)
        if m.sender:
            participants.add(m.sender)
        if target:
            participants.add(target)
        key = tuple(sorted(participants))
        groups.setdefault(key, []).append(m)
    for lst in groups.values():
        lst.sort(key=lambda mm: (mm.date_dt or datetime.max.replace(tzinfo=timezone.utc), mm.date_raw))
    return groups


# ------------------------- HTML Rendering -------------------------

def render_thread_html(
    messages_root: Path,
    out_file: Path,
    msgs: List[Message],
    participants: List[str],
    target_number: str,
    contact_lookup: Callable[[str], str] = lambda x: x,
    index_href: str = "index.html",
) -> Tuple[int, int]:
    total = len(msgs)
    with_attachments = 0

    disp_participants = [format_contact_label(p, contact_lookup) for p in participants]
    title = f"Chat – {', '.join(disp_participants)}"

    parts: List[str] = []
    parts.append("<!DOCTYPE html>")
    parts.append("<html lang=\"en\">")
    parts.append("<head>")
    parts.append("<meta charset=\"utf-8\">")
    parts.append("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">")
    parts.append(f"<title>{html.escape(title)}</title>")
    parts.append("<style>" + CSS_STYLES + "</style>")
    # Per-conversation print footer: chat name on the left (paged media)
    css_title = title.replace("\\", "\\\\").replace('"', '\\"')
    parts.append(
        "<style>\n"
        "@media print {\n"
        "  @page {\n"
        "    @bottom-left {\n"
        f'      content: "{css_title}";\n'
        "      font-size: 9pt;\n"
        "      color: #4b5563;\n"
        "    }\n"
        "  }\n"
        "}\n"
        "</style>"
    )
    parts.append("</head>")
    parts.append("<body>")

    parts.append("<div class=\"header\">")
    parts.append("  <div class=\"container\">")
    parts.append(f"    <div><strong>{html.escape(title)}</strong></div>")
    target_disp = format_contact_label(target_number, contact_lookup)
    meta_line = f"Owner: {html.escape(target_disp)}<br>Participants: {html.escape(', '.join(disp_participants))}"
    parts.append(f"    <div class=\"thread-meta\">{meta_line}</div>")
    parts.append(
        "    <div class=\"search-bar\">"
        "<input id=\"search\" class=\"search-input\" placeholder=\"Search messages\">"
        "<label class=\"att-filter\" title=\"Show only messages that include attachments\">"
        "<input type=\"checkbox\" id=\"att-only\"> Attachments only"
        "</label>"
        "<button type=\"button\" class=\"print-btn\" onclick=\"window.print()\" "
        "title=\"Print this conversation\">Print</button>"
        "</div>"
    )
    parts.append("  </div>")
    parts.append("</div>")

    parts.append("<div class=\"container\">")

    current_day: Optional[str] = None

    for m in msgs:
        # Day divider (based on local date of parsed datetime if available, else raw)
        day_label = None
        if m.date_dt:
            day_label = m.date_dt.astimezone().strftime("%A, %B %d, %Y")
        elif m.date_raw:
            # Try to grab just the date part
            day_label = m.date_raw.split("T")[0]
        if day_label and day_label != current_day:
            current_day = day_label
            parts.append(f"<div class=\"day-divider\">{html.escape(current_day)}</div>")

        side_class = "sent" if m.direction == "out" else "received"
        # Mark messages that reference at least one real attachment so the
        # "Attachments only" filter can show/hide them client-side.
        has_att = False
        if m.attachments:
            for _fname in m.attachments:
                if _fname and _fname.lower() not in {"null", "null.txt", "none", "(null)", "aaaa"}:
                    has_att = True
                    break
        att_class = " has-attachment" if has_att else ""
        parts.append(f"<div class=\"message {side_class}{att_class}\">")
        parts.append("  <div class=\"bubble\">")

        sender = safe_text(format_contact_label(m.sender, contact_lookup))
        if sender:
            parts.append(f"    <div class=\"sender\">{sender}</div>")

        body_html = safe_text(m.body)
        if body_html:
            parts.append(f"    <div class=\"body-text\">{body_html}</div>")
        else:
            msg_type = (m.msg_type or "").upper()
            placeholder = (
                f"NO DATA IN CSV FOR THIS {msg_type} MESSAGE - LOG ONLY" if msg_type else "NO DATA IN CSV FOR THIS MESSAGE - LOG ONLY"
            )
            parts.append(f"    <div class=\"body-text missing\">{placeholder}</div>")

        # Attachments
        attachment_snippets: List[str] = []
        if m.attachments:
            for fname in m.attachments:
                if not fname or fname.lower() in {"null", "null.txt", "none", "(null)", "aaaa"}:
                    continue
                if not m.attachment_day:
                    continue
                target = build_attachment_path(
                    messages_root, m.msg_type or "", m.direction or "", m.attachment_day, fname
                )
                if not target.exists():
                    # Some exports stash attachments without the dated subfolder; check fallback path
                    alt = messages_root / "attachments" / (m.msg_type or "") / (m.direction or "") / fname
                    chosen = target if target.exists() else (alt if alt.exists() else None)
                else:
                    chosen = target

                if chosen is None:
                    # Show a small missing-note so you know there *was* an attachment reference
                    attachment_snippets.append(
                        f"<div class=\"attachment\"><em class=\"meta\">(missing attachment: {html.escape(fname)})</em></div>"
                    )
                    continue

                kind = classify_ext(chosen)
                rel = relpath_for_html(out_file, chosen)
                if kind == "image":
                    attachment_snippets.append(
                        f"<div class=\"attachment\"><img loading=\"lazy\" src=\"{rel}\" alt=\"{html.escape(fname)}\"></div>"
                    )
                elif kind == "video":
                    attachment_snippets.append(
                        f"<div class=\"attachment\"><video controls preload=\"metadata\" src=\"{rel}\"></video></div>"
                    )
                elif kind == "audio":
                    attachment_snippets.append(
                        f"<div class=\"attachment\"><audio controls preload=\"metadata\" src=\"{rel}\"></audio></div>"
                    )
                elif kind == "inline_text":
                    try:
                        text = chosen.read_text(encoding="utf-8", errors="replace")
                        text = html.escape(text)
                        attachment_snippets.append(
                            f"<div class=\"attachment\"><pre class=\"code\" style=\"white-space:pre-wrap\">{text}</pre></div>"
                        )
                    except Exception:
                        attachment_snippets.append(
                            f"<div class=\"attachment\"><a href=\"{rel}\" download>{html.escape(fname)}</a></div>"
                        )
                else:
                    attachment_snippets.append(
                        f"<div class=\"attachment\"><a href=\"{rel}\" download>{html.escape(fname)}</a></div>"
                    )
            if attachment_snippets:
                with_attachments += 1
                parts.append("    <div class=\"attachments\">")
                parts.extend(["      " + s for s in attachment_snippets])
                parts.append("    </div>")

        msg_type = (m.msg_type or "unknown").upper()
        if msg_type == "MMS" and not attachment_snippets:
            parts.append(
                f"    <div class=\"body-text missing\">NO {msg_type} ATTACHMENT AVAILABLE - LOG ONLY</div>"
            )

        # Meta line
        if m.date_dt:
            local_str = m.date_dt.astimezone().strftime("%Y-%m-%d %H:%M:%S %Z")
            parts.append(
                f"    <div class=\"meta\">{html.escape(local_str)} · {html.escape(m.direction)} · {html.escape(m.msg_type)}</div>"
            )
        else:
            parts.append(
                f"    <div class=\"meta\">{html.escape(m.date_raw)} · {html.escape(m.direction)} · {html.escape(m.msg_type)}</div>"
            )

        parts.append("  </div>")  # bubble
        parts.append("</div>")    # message

    parts.append("</div>")  # container

    parts.append("<div class=\"footer\">")
    parts.append(
        f"  <div class=\"container\">Return to <a href=\"{html.escape(index_href)}\">index</a></div>"
    )
    parts.append("</div>")
    parts.append("<script>")
    parts.append("""
(function(){
  const search = document.getElementById('search');
  const attOnly = document.getElementById('att-only');
  function applyFilters(){
    const q = (search && search.value || '').toLowerCase();
    const onlyAtt = !!(attOnly && attOnly.checked);
    document.querySelectorAll('.message').forEach(function(m){
      const matchSearch = !q || m.textContent.toLowerCase().includes(q);
      const matchAtt = !onlyAtt || m.classList.contains('has-attachment');
      m.style.display = (matchSearch && matchAtt) ? '' : 'none';
    });
    // Hide day dividers that have no visible messages until the next divider
    document.querySelectorAll('.day-divider').forEach(function(div){
      let el = div.nextElementSibling;
      let anyVisible = false;
      while (el && !el.classList.contains('day-divider')) {
        if (el.classList.contains('message') && el.style.display !== 'none') {
          anyVisible = true;
          break;
        }
        el = el.nextElementSibling;
      }
      div.style.display = anyVisible ? '' : 'none';
    });
  }
  if (search) search.addEventListener('input', applyFilters);
  if (attOnly) attOnly.addEventListener('change', applyFilters);
})();
""".strip())
    parts.append("</script>")
    parts.append("</body></html>")

    out_file.parent.mkdir(parents=True, exist_ok=True)
    out_file.write_text("\n".join(parts), encoding="utf-8")

    return total, with_attachments


# ------------------------- Index Page -------------------------

def build_search_blob(
    msgs: List[Message],
    contact_lookup: Callable[[str], str] = lambda x: x,
) -> str:
    """Concatenate message bodies, party labels, and attachment names for full-text search."""
    parts: List[str] = []
    for m in msgs:
        if m.body:
            parts.append(str(m.body))
        if m.sender:
            parts.append(format_contact_label(m.sender, contact_lookup))
        if m.recipients:
            for p in m.recipients.replace(",", ";").split(";"):
                p = p.strip()
                if p:
                    parts.append(format_contact_label(p, contact_lookup))
        if m.attachments:
            for fname in m.attachments:
                if fname and fname.lower() not in {"null", "null.txt", "none", "(null)", "aaaa"}:
                    parts.append(fname)
        if m.date_raw:
            parts.append(m.date_raw)
    return " ".join(parts).lower()


def write_index(out_dir: Path, entries: List[Tuple]) -> None:
    """Write the master conversation index.

    Each *entry* is a tuple of:
      (title, rel_path, msg_count, with_attach_count[, search_blob])

    The optional fifth element is lower-cased message text used for full-text
    search across all chats from the index page.
    """
    lines: List[str] = []
    lines.append("<!DOCTYPE html>")
    lines.append("<html lang=\"en\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">")
    lines.append("<title>Message Transcripts</title>")
    lines.append(f"<style>{INDEX_CSS}</style>")
    lines.append("</head><body>")
    lines.append("<div class=\"container\">")
    lines.append("  <h1>Message Transcripts</h1>")
    lines.append(
        "  <div class=\"subtitle\">One HTML per chat. Search matches chat titles "
        "<em>and</em> message content. (Times shown in local system timezone inside each transcript.)</div>"
    )
    lines.append(
        "  <div class=\"search-bar\">"
        "<input id=\"search\" class=\"search-input\" "
        "placeholder=\"Search chats &amp; message content…\">"
        "</div>"
    )
    lines.append("  <div class=\"list\">")

    search_blobs: List[str] = []
    for i, entry in enumerate(entries):
        title, rel, c, ca = entry[0], entry[1], entry[2], entry[3]
        blob = entry[4] if len(entry) > 4 else ""
        search_blobs.append(blob or "")
        lines.append(f"    <div class=\"item\" data-i=\"{i}\">")
        lines.append(f"      <a href=\"{html.escape(rel)}\">{html.escape(title)}</a>")
        lines.append(
            f"      <div class=\"meta\">Messages: {c} · Messages with attachments: {ca}</div>"
        )
        lines.append("    </div>")

    lines.append("  </div>")
    lines.append(
        "  <div id=\"search-status\" class=\"subtitle\" style=\"margin-top:12px;display:none\"></div>"
    )
    lines.append("</div>")
    lines.append("<script>")
    # Embed searchable message text as a JSON array (parallel to data-i on each item)
    lines.append(f"const CHAT_SEARCH = {json.dumps(search_blobs, ensure_ascii=False)};")
    lines.append("""
(function(){
  const s = document.getElementById('search');
  const status = document.getElementById('search-status');
  if (!s) return;
  function apply(){
    const q = (s.value || '').toLowerCase().trim();
    let shown = 0;
    document.querySelectorAll('.item').forEach(function(it){
      if (!q) {
        it.style.display = '';
        shown++;
        return;
      }
      const i = parseInt(it.getAttribute('data-i'), 10);
      const titleMeta = (it.textContent || '').toLowerCase();
      const body = (CHAT_SEARCH[i] || '');
      const match = titleMeta.indexOf(q) !== -1 || body.indexOf(q) !== -1;
      it.style.display = match ? '' : 'none';
      if (match) shown++;
    });
    if (status) {
      if (q) {
        status.style.display = '';
        status.textContent = shown + ' chat' + (shown === 1 ? '' : 's') + ' match \"' + s.value + '\"';
      } else {
        status.style.display = 'none';
        status.textContent = '';
      }
    }
  }
  s.addEventListener('input', apply);
})();
""".strip())
    lines.append("</script>")
    lines.append("</body></html>")
    (out_dir / "index.html").write_text("\n".join(lines), encoding="utf-8")


# ------------------------- Main -------------------------

def main():
    ap = argparse.ArgumentParser(description="Render chat transcripts from CSVs into HTML.")
    ap.add_argument("--in", dest="in_dir", required=True, help="Input root folder (expects CSVs inside, plus attachments/...) e.g. messages")
    ap.add_argument("--out", dest="out_dir", required=True, help="Output folder for HTML transcripts, e.g. transcripts")
    ap.add_argument("--target-number", default="", help="Phone number of the target user")
    ap.add_argument(
        "--contacts-xlsx",
        dest="contacts_xlsx",
        default="",
        help="Path to Excel file mapping phone numbers to contacts",
    )
    args = ap.parse_args()

    target = args.target_number
    lookup = build_contact_lookup(args.contacts_xlsx)

    messages_root = Path(args.in_dir).resolve()
    out_root = Path(args.out_dir).resolve()
    out_root.mkdir(parents=True, exist_ok=True)

    csv_files = sorted(messages_root.glob("*.csv"))
    if not csv_files:
        print(f"No CSV files found in {messages_root}")
        return

    all_msgs: List[Message] = []
    call_records: List[Message] = []

    for csv_file in csv_files:
        msgs = load_messages_from_csv(csv_file, lookup)
        for m in msgs:
            if m.msg_type == "call":
                if not m.sender:
                    m.sender = target
                if not m.recipients:
                    m.recipients = target
                call_records.append(m)
            else:
                all_msgs.append(m)

    grouped = group_messages_by_chat(all_msgs, target)

    index_entries: List[Tuple] = []
    for participants, msgs in grouped.items():
        disp = [format_contact_label(p, lookup) for p in participants]
        title = f"Chat – {', '.join(disp)}"
        key = sanitize_participants(participants)
        out_file = out_root / f"chat-{key}.html"
        total, with_attachments = render_thread_html(
            messages_root, out_file, msgs, list(participants), target, lookup
        )
        rel = os.path.relpath(out_file, start=out_root).replace(os.sep, "/")
        search_blob = build_search_blob(msgs, lookup)
        index_entries.append((title, rel, total, with_attachments, search_blob))
        print(f"Rendered chat {', '.join(disp)}: {total} messages ({with_attachments} with attachments)")

    write_index(out_root, index_entries)

    from openpyxl import Workbook

    call_log_path = out_root / "Call Log.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.append(["Date", "Direction", "Sender", "Recipients", "Message ID"])
    for m in call_records:
        if m.date_dt:
            date_str = m.date_dt.astimezone().strftime("%Y-%m-%d %H:%M:%S %Z")
        else:
            date_str = m.date_raw
        sender_disp = format_contact_label(m.sender, lookup) if m.sender else m.sender
        if m.recipients:
            recip_disp = "; ".join(
                format_contact_label(p.strip(), lookup)
                for p in m.recipients.replace(",", ";").split(";")
                if p.strip()
            )
        else:
            recip_disp = m.recipients
        ws.append([date_str, m.direction, sender_disp, recip_disp, m.message_id])
    wb.save(call_log_path)

    print(f"\nDone. Open: {out_root / 'index.html'}")


if __name__ == "__main__":
    main()
