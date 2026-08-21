#!/usr/bin/env python3
"""
Unified Synchronoss end-to-end pipeline.

Accepts a top-level ``selection.zip`` (as delivered by Synchronoss via
leotransfers) plus the GPG decryption passphrase, then automatically:

1. Extracts the outer zip
2. Discovers and decrypts the three standard encrypted payloads
   (main content zip, contacts txt, quarantined zip)
3. Extracts the main content
4. Converts contacts → Excel
5. Collects / normalises quarantined media
6. Collects media + message attachments (with metadata logs)
7. Renders chat transcripts — **one folder per conversation**
8. Produces a named call log and a processing summary

Usage (CLI)
-----------
    python -m synchronoss_parser.full_pipeline \\
        --selection selection.zip \\
        --password "the-gpg-passphrase" \\
        --output ./CaseOutput \\
        --target-name "Jane Doe"

The target phone number is auto-detected from the encrypted filenames.
Optionally supply ``--target-name`` so the target appears by name in
transcripts (same as other contacts).

Or launch the GUI tab via the main toolbox.
"""

from __future__ import annotations

import argparse
import logging
import os
import re
import shutil
import sys
import traceback
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

from .utils import (
    decrypt_gpg_file,
    find_gpg,
    gpg_available,
    normalize_phone_number,
    ensure_unique_name,
)

logger = logging.getLogger("synchronoss.pipeline")


# ---------------------------------------------------------------------------
# Discovery helpers
# ---------------------------------------------------------------------------

PHONE_RE = re.compile(r"(\d{10})")
CONTACTS_RE = re.compile(r"(\d{10})-.*?-contacts\.txt\.gpg$", re.IGNORECASE)
QUARANTINE_RE = re.compile(r"(\d{10})-.*?-quarantined\.zip\.gpg$", re.IGNORECASE)
MAIN_ZIP_RE = re.compile(r"^(\d{10})\.zip\.gpg$", re.IGNORECASE)


def discover_payloads(extracted_root: Path) -> Dict[str, Path]:
    """Locate the three standard encrypted files inside the extracted selection.zip.

    Returns a dict with keys ``main``, ``contacts``, ``quarantine`` (values may be
    missing if a particular payload was not present).
    """
    found: Dict[str, Path] = {}
    for p in extracted_root.rglob("*.gpg"):
        name = p.name
        if MAIN_ZIP_RE.match(name) and "main" not in found:
            found["main"] = p
        elif CONTACTS_RE.search(name) and "contacts" not in found:
            found["contacts"] = p
        elif QUARANTINE_RE.search(name) and "quarantine" not in found:
            found["quarantine"] = p
    return found


def extract_target_phone(payloads: Dict[str, Path], fallback: str = "") -> str:
    """Best-effort extraction of the 10-digit target account number from filenames."""
    for key in ("main", "contacts", "quarantine"):
        p = payloads.get(key)
        if p:
            m = PHONE_RE.search(p.name)
            if m:
                return m.group(1)
    return normalize_phone_number(fallback)


# ---------------------------------------------------------------------------
# Core pipeline stages
# ---------------------------------------------------------------------------

def stage_extract_outer(selection_zip: Path, work_dir: Path) -> Path:
    """Extract the top-level selection.zip into work_dir/selection_extracted."""
    dest = work_dir / "00_selection_extracted"
    dest.mkdir(parents=True, exist_ok=True)
    logger.info("Extracting outer archive %s → %s", selection_zip.name, dest)
    with zipfile.ZipFile(selection_zip) as zf:
        zf.extractall(dest)
    return dest


def stage_decrypt_payloads(
    payloads: Dict[str, Path],
    passphrase: str,
    work_dir: Path,
) -> Dict[str, Path]:
    """Decrypt the discovered .gpg files into work_dir/01_decrypted/."""
    out: Dict[str, Path] = {}
    decrypt_root = work_dir / "01_decrypted"
    decrypt_root.mkdir(parents=True, exist_ok=True)

    for key, gpg_path in payloads.items():
        # Derive a clean output name
        if key == "main":
            out_name = gpg_path.name[:-4]  # strip .gpg → .zip
        elif key == "contacts":
            out_name = "contacts.txt"
        elif key == "quarantine":
            out_name = "quarantined.zip"
        else:
            out_name = gpg_path.name[:-4]

        dest = decrypt_root / out_name
        logger.info("Decrypting %s → %s", gpg_path.name, dest.name)
        decrypt_gpg_file(gpg_path, dest, passphrase)
        out[key] = dest
    return out


def stage_extract_main(main_zip: Path, work_dir: Path) -> Path:
    """Extract the main content zip into work_dir/02_main_content/."""
    dest = work_dir / "02_main_content"
    dest.mkdir(parents=True, exist_ok=True)
    logger.info("Extracting main content archive → %s", dest)
    with zipfile.ZipFile(main_zip) as zf:
        zf.extractall(dest)
    return dest


def stage_contacts(contacts_txt: Path, final_out: Path) -> Optional[Path]:
    """Convert the decrypted contacts.txt into an Excel workbook."""
    try:
        from .contacts_to_excel import convert_contacts
    except ImportError:
        logger.warning("contacts_to_excel not available – skipping contacts conversion")
        return None

    xlsx = final_out / "contacts.xlsx"
    logger.info("Converting contacts → %s", xlsx)
    try:
        rows = convert_contacts(str(contacts_txt), str(xlsx))
        logger.info("Wrote %d contact rows", rows)
        return xlsx
    except Exception as e:
        logger.error("Contacts conversion failed: %s", e)
        return None


def stage_quarantine(
    quarantine_zip: Path, work_dir: Path, final_out: Path
) -> Tuple[Path, List[Dict]]:
    """Extract quarantine zip, recover media, return (compiled_dir, index_rows)."""
    q_root = work_dir / "03_quarantine_raw"
    q_root.mkdir(parents=True, exist_ok=True)
    logger.info("Extracting quarantined archive → %s", q_root)
    with zipfile.ZipFile(quarantine_zip) as zf:
        zf.extractall(q_root)

    compiled = final_out / "Compiled Quarantine Files"
    compiled.mkdir(parents=True, exist_ok=True)
    records: List[Dict] = []

    try:
        from .collect_quarantined_files import collect_quarantined_files
        from .utils import md5sum

        copied, skipped, total = collect_quarantined_files(q_root, compiled)
        logger.info(
            "Quarantine: converted %d of %d files (%d skipped)",
            len(copied), total, len(skipped),
        )
        for dest in copied:
            try:
                digest = md5sum(dest)
            except OSError:
                digest = ""
            records.append(
                {
                    "filename": dest.name,
                    "detected_type": dest.suffix.lower() or "unknown",
                    "md5": digest,
                    "rel_path": f"Compiled Quarantine Files/{dest.name}",
                }
            )
        records.sort(key=lambda r: r.get("filename", "").lower())
    except Exception as e:
        logger.error("Quarantine collection failed: %s", e)
        traceback.print_exc()
    return compiled, records


def stage_media_and_attachments(
    main_content: Path,
    contacts_xlsx: Optional[Path],
    final_out: Path,
) -> Tuple[List[Dict], List[Dict], List[Dict]]:
    """Run collect_media, collect_attachments, and unlinked-MMS discovery.

    Returns ``(vz_records, unlinked_mms_rows, attachment_records)`` for the case index.
    """
    vz_records: List[Dict] = []
    unlinked_rows: List[Dict] = []
    attachment_records: List[Dict] = []

    # Media (VZMOBILE style)
    vz = None
    for candidate in main_content.rglob("VZMOBILE"):
        if candidate.is_dir():
            vz = candidate
            break
    if vz:
        try:
            from .collect_media import collect_media, write_excel
            media_out = final_out / "Compiled Media"
            media_out.mkdir(parents=True, exist_ok=True)
            logfile = media_out / "compiled_media_log" / "compiled_media_log.xlsx"
            records, exif_keys = collect_media(vz, media_out)
            write_excel(records, exif_keys, logfile)
            # Index-friendly subset + relative link under Compiled Media/
            # Plus residual EXIF / metadata for the info popover
            _vz_basic = {"Date", "Device", "File Name", "MD5"}
            for rec in records:
                meta = {
                    str(k): str(v)
                    for k, v in rec.items()
                    if k not in _vz_basic and v not in (None, "")
                }
                vz_records.append(
                    {
                        "upload_date": rec.get("Date", ""),
                        "device": rec.get("Device", ""),
                        "filename": rec.get("File Name", ""),
                        "md5": rec.get("MD5", ""),
                        "rel_path": f"Compiled Media/{rec.get('File Name', '')}",
                        "meta": meta,
                    }
                )
            logger.info("Collected %d media files from VZMOBILE", len(records))
        except Exception as e:
            logger.error("Media collection failed: %s", e)

    # Message attachments
    messages = None
    for candidate in main_content.rglob("messages"):
        if candidate.is_dir() and (candidate / "attachments").exists():
            messages = candidate
            break
    if messages:
        try:
            from .collect_attachments import collect_attachments, write_excel as write_att_excel
            att_out = final_out / "Compiled Attachments"
            att_out.mkdir(parents=True, exist_ok=True)
            att_log = att_out / "compiled_attachment_log" / "compiled_attachment_log.xlsx"
            logger.info(
                "Collecting message attachments from %s → %s",
                messages / "attachments",
                att_out,
            )
            records, exif_keys = collect_attachments(
                attachments_root=messages / "attachments",
                compiled_path=att_out,
                contacts_xlsx=contacts_xlsx,
            )
            write_att_excel(records, exif_keys, att_log)
            _att_basic = {
                "File Name", "Original Name", "Date", "Sender", "Recipient", "MD5"
            }
            for rec in records:
                fname = rec.get("File Name", "") or rec.get("Original Name", "")
                meta = {
                    str(k): str(v)
                    for k, v in rec.items()
                    if k not in _att_basic and v not in (None, "")
                }
                attachment_records.append(
                    {
                        "filename": fname,
                        "original_name": rec.get("Original Name", ""),
                        "date": rec.get("Date", ""),
                        "sender": rec.get("Sender", ""),
                        "recipient": rec.get("Recipient", ""),
                        "md5": rec.get("MD5", ""),
                        "rel_path": f"Compiled Attachments/{fname}" if fname else "",
                        "meta": meta,
                    }
                )
            logger.info(
                "Collected %d attachment files into %s (log: %s)",
                len(records),
                att_out,
                att_log,
            )
        except Exception as e:
            logger.error("Attachment collection failed: %s", e)
            traceback.print_exc()

        # Unlinked MMS: present on disk, not referenced by a real-extension token
        try:
            from .unlinked_mms import (
                find_unlinked_mms,
                materialize_unlinked_mms,
                write_unlinked_excel,
            )
            raw_unlinked = find_unlinked_mms(messages)
            if raw_unlinked:
                unlinked_rows = materialize_unlinked_mms(
                    raw_unlinked,
                    dest_dir=final_out / "Unlinked MMS",
                    compiled_attachments=final_out / "Compiled Attachments",
                )
                write_unlinked_excel(
                    unlinked_rows, final_out / "Unlinked MMS.xlsx"
                )
                logger.info(
                    "Unlinked MMS media: %d file(s) not tied to a message token",
                    len(unlinked_rows),
                )
            else:
                logger.info("Unlinked MMS media: none found")
        except Exception as e:
            logger.error("Unlinked MMS discovery failed: %s", e)
            traceback.print_exc()

    return vz_records, unlinked_rows, attachment_records


def stage_render_transcripts(
    main_content: Path,
    contacts_xlsx: Optional[Path],
    target_number: str,
    final_out: Path,
    target_name: str = "",
    summary: Optional[Dict] = None,
    dv_uploads: Optional[List] = None,
    dv_syncs: Optional[List] = None,
    vz_records: Optional[List] = None,
    unlinked_mms: Optional[List] = None,
    quarantine_records: Optional[List] = None,
    attachment_records: Optional[List] = None,
) -> Path:
    """Render chat transcripts with one folder per conversation.

    The master ``index.html`` is written to *final_out* (case root) so it sits
    alongside contacts, the call log, and other top-level artifacts.
    """
    messages_root = None
    for candidate in main_content.rglob("messages"):
        if candidate.is_dir() and list(candidate.glob("*.csv")):
            messages_root = candidate
            break
    if not messages_root:
        logger.warning("No messages folder with CSVs found – skipping transcript rendering")
        # Still write an index shell so Home / Contacts / DV are available
        try:
            from .render_transcripts import write_index, load_contacts_for_html
            contacts_rows = load_contacts_for_html(
                str(contacts_xlsx) if contacts_xlsx else None
            )
            write_index(
                final_out,
                [],
                summary=summary or {},
                contacts=contacts_rows,
                dv_uploads=dv_uploads or [],
                dv_syncs=dv_syncs or [],
                vz_records=vz_records or [],
                unlinked_mms=unlinked_mms or [],
                quarantine_records=quarantine_records or [],
                attachment_records=attachment_records or [],
            )
        except Exception as e:
            logger.error("Failed to write index shell: %s", e)
        return final_out / "conversations"

    conv_root = final_out / "conversations"
    conv_root.mkdir(parents=True, exist_ok=True)

    try:
        from .render_transcripts import (
            load_messages_from_csv,
            group_messages_by_chat,
            render_thread_html,
            build_contact_lookup,
            format_contact_label,
            build_search_blob,
            write_index,
            load_contacts_for_html,
            Message,
        )
        from openpyxl import Workbook
    except ImportError as e:
        logger.error("Cannot import render_transcripts helpers: %s", e)
        return conv_root

    base_lookup = build_contact_lookup(str(contacts_xlsx) if contacts_xlsx else None)
    target = normalize_phone_number(target_number)
    target_name = (target_name or "").strip()

    # Inject the investigator-supplied target name into the contact lookup so
    # the target appears by name (with number) the same way other contacts do.
    def lookup(number: str) -> str:
        digits = normalize_phone_number(number)
        if target and digits == target and target_name:
            return target_name
        return base_lookup(number)

    all_msgs: List[Message] = []
    call_records: List[Message] = []

    for csv_file in sorted(messages_root.glob("*.csv")):
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
        # Folder names use contact name when known (no number – keeps paths short)
        parts = []
        for p in participants:
            name = lookup(p)
            digits = normalize_phone_number(p)
            # Prefer the contact name alone for the folder; fall back to digits
            label = name if (name and digits and name != digits and name != p) else (digits or p)
            safe = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "", label).strip() or "unknown"
            parts.append(safe)
        folder_name = "-".join(parts)[:120] or "chat"
        conv_dir = conv_root / folder_name
        counter = 1
        while conv_dir.exists():
            conv_dir = conv_root / f"{folder_name}_{counter}"
            counter += 1
        conv_dir.mkdir(parents=True, exist_ok=True)

        out_file = conv_dir / "conversation.html"
        # Relative path from this conversation file back to the Conversations
        # section of the case index (not Home).
        try:
            index_href = os.path.relpath(
                final_out / "index.html", start=out_file.parent
            ).replace("\\", "/") + "#conversations"
        except Exception:
            index_href = "../../index.html#conversations"

        total, with_att = render_thread_html(
            messages_root,
            out_file,
            msgs,
            list(participants),
            target,
            lookup,
            index_href=index_href,
        )
        # Relative path from the case-root index → this conversation
        rel = out_file.relative_to(final_out).as_posix()
        disp = [format_contact_label(p, lookup) for p in participants]
        title = f"Chat – {', '.join(disp)}"
        search_blob = build_search_blob(msgs, lookup)
        index_entries.append((title, rel, total, with_att, search_blob))
        logger.info("Rendered %s → %s (%d messages)", folder_name, out_file, total)

    # Build call-log rows once (Excel + index page)
    call_log_rows: List[Dict] = []
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
        call_log_rows.append(
            {
                "date": date_str,
                "direction": m.direction,
                "sender": sender_disp,
                "recipients": recip_disp,
                "message_id": m.message_id,
            }
        )

    # Master index (Home / Conversations / Contacts / Call Logs / DV)
    contacts_rows = load_contacts_for_html(
        str(contacts_xlsx) if contacts_xlsx else None
    )
    write_index(
        final_out,
        index_entries,
        summary=summary or {},
        contacts=contacts_rows,
        call_logs=call_log_rows,
        dv_uploads=dv_uploads or [],
        dv_syncs=dv_syncs or [],
        vz_records=vz_records or [],
        unlinked_mms=unlinked_mms or [],
        quarantine_records=quarantine_records or [],
        attachment_records=attachment_records or [],
    )
    logger.info("Wrote case index → %s", final_out / "index.html")

    call_log_path = final_out / "Call Log.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Call Log"
    ws.append(["Date", "Direction", "Sender", "Recipients", "Message ID"])
    for row in call_log_rows:
        ws.append(
            [
                row["date"],
                row["direction"],
                row["sender"],
                row["recipients"],
                row["message_id"],
            ]
        )
    wb.save(call_log_path)
    logger.info("Wrote call log → %s", call_log_path)

    return conv_root


def write_summary(
    final_out: Path,
    selection_zip: Path,
    target: str,
    payloads: Dict[str, Path],
    started: datetime,
    target_name: str = "",
    dv_upload_count: int = 0,
    dv_sync_count: int = 0,
    dv_files: Optional[List[Path]] = None,
) -> Tuple[Path, Dict[str, str]]:
    """Write processing_summary.txt and return (path, summary_dict for index.html)."""
    finished = datetime.now()
    summary_path = final_out / "processing_summary.txt"
    payload_names = ", ".join(f"{k}={p.name}" for k, p in payloads.items()) or "(none)"
    dv_names = ", ".join(p.name for p in (dv_files or [])) or "(none)"

    lines = [
        "Synchronoss Unified Pipeline – Processing Summary",
        "=" * 50,
        f"Started          : {started.isoformat(timespec='seconds')}",
        f"Finished         : {finished.isoformat(timespec='seconds')}",
        f"Input archive    : {selection_zip}",
        f"Owner phone      : {target}",
        f"Owner name       : {target_name or '(not supplied)'}",
        f"Output folder    : {final_out}",
        f"Case root        : {final_out.parent}",
        "",
        "Payloads discovered:",
    ]
    for k, p in payloads.items():
        lines.append(f"  {k:12s}: {p.name}")
    lines += [
        "",
        f"DV Access Log files: {dv_names}",
        f"  Upload events  : {dv_upload_count}",
        f"  Sync events    : {dv_sync_count}",
        "",
        "Case root layout:",
        "  original_working/       – intermediate decryption & extraction artifacts",
        "  parsed_output/          – final organized case data (this folder)",
        "",
        "Inside parsed_output/:",
        "  index.html              – case index (Home / Conversations / Contacts / Call Logs /",
        "                            Unlinked MMS / VZMOBILE / DV Access Logs)",
        "  conversations/          – one folder per chat thread",
        "  contacts.xlsx           – converted contact list",
        "  Call Log.xlsx           – voice calls with names",
        "  Unlinked MMS.xlsx       – MMS folder media not tied to a message token",
        "  Unlinked MMS/           – copies of those files (when not already in Compiled Attachments)",
        "  DV Access Logs.xlsx     – upload & sync events (when DV CSVs present)",
        "  Compiled Media/         – media from VZMOBILE (if present)",
        "  Compiled Attachments/   – MMS/RCS attachments (if present)",
        "  Compiled Quarantine Files/ – recovered quarantined media",
        "",
        "Notes:",
        "  • Intermediate decryption artifacts are kept under original_working/",
        "    for chain-of-custody review (same parent folder as parsed_output/).",
        "  • Each conversation folder contains conversation.html plus any",
        "    relative links to original attachments.",
        "  • DV Access Log CSVs are discovered next to selection.zip and inside",
        "    the extracted return (filename pattern: *Dv Access logs*.csv).",
    ]
    summary_path.write_text("\n".join(lines), encoding="utf-8")

    summary_dict: Dict[str, str] = {
        "started": started.isoformat(timespec="seconds"),
        "finished": finished.isoformat(timespec="seconds"),
        "input_archive": str(selection_zip),
        "output_folder": str(final_out),
        "owner_phone": target or "",
        "owner_name": target_name or "",
        "payloads": payload_names,
        "notes": (
            f"DV log files: {dv_names}\n"
            f"DV uploads: {dv_upload_count}  |  DV sync events: {dv_sync_count}"
        ),
    }
    return summary_path, summary_dict


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def run_pipeline(
    selection_zip: Path | str,
    passphrase: str,
    output_dir: Path | str,
    target_number: str = "",
    target_name: str = "",
    keep_work_dir: bool = True,
    progress_callback: Optional[Callable[[str], None]] = None,
) -> Path:
    """Execute the full Synchronoss processing pipeline.

    Parameters
    ----------
    selection_zip
        Path to the outer ``selection.zip`` (or equivalent) delivered by Synchronoss.
    passphrase
        GPG decryption passphrase (the one supplied in the #secure email).
    output_dir
        Case root folder selected by the user.  Inside it the pipeline creates:
          • ``original_working/`` – intermediate decryption & extraction artifacts
          • ``parsed_output/``    – final organized case data (contacts, transcripts, etc.)
    target_number
        Optional override for the target phone number.  Normally left blank –
        the pipeline auto-detects it from the encrypted filenames.
    target_name
        Optional display name for the target account holder (e.g. "Jane Doe").
        When supplied, transcripts and the call log show this name instead of
        the raw phone number, the same way other contacts are labelled.
    keep_work_dir
        If True, intermediate decryption/extraction folders under
        ``original_working/`` are retained (recommended for forensic review).
    progress_callback
        Optional callable that receives human-readable status strings
        (useful for GUI progress updates).

    Returns
    -------
    Path
        The final ``parsed_output/`` directory inside the case root.
    """
    def status(msg: str) -> None:
        logger.info(msg)
        if progress_callback:
            progress_callback(msg)

    selection_zip = Path(selection_zip).expanduser().resolve()
    case_root = Path(output_dir).expanduser().resolve()
    case_root.mkdir(parents=True, exist_ok=True)

    # Layout under the user-selected path:
    #   <case_root>/
    #     original_working/ – intermediate decryption / extraction
    #     parsed_output/    – final case artifacts
    work_dir = case_root / "original_working"
    output_dir = case_root / "parsed_output"
    output_dir.mkdir(parents=True, exist_ok=True)

    if not selection_zip.is_file():
        raise FileNotFoundError(f"Selection archive not found: {selection_zip}")

    if not gpg_available():
        raise FileNotFoundError(
            "gpg executable not found.\n\n"
            "Install GPG4Win (which includes Kleopatra and the gpg command-line tool):\n"
            "  https://www.gpg4win.org\n\n"
            "After installation, restart this application or ensure gpg is on your PATH."
        )

    started = datetime.now()
    if work_dir.exists():
        shutil.rmtree(work_dir, ignore_errors=True)
    work_dir.mkdir(parents=True, exist_ok=True)

    status("1/9  Extracting outer selection.zip …")
    extracted = stage_extract_outer(selection_zip, work_dir)

    status("2/9  Discovering encrypted payloads …")
    payloads = discover_payloads(extracted)
    if not payloads:
        raise RuntimeError(
            "No recognized encrypted payloads found inside selection.zip.\n"
            "Expected files matching:\n"
            "  • <10-digit-phone>.zip.gpg\n"
            "  • <phone>-*-contacts.txt.gpg\n"
            "  • <phone>-*-quarantined.zip.gpg"
        )
    status(f"     Found: {', '.join(payloads.keys())}")

    target = extract_target_phone(payloads, target_number)
    if not target:
        status("WARNING: could not determine owner phone number – contact matching may be incomplete")
    else:
        status(f"     Owner phone auto-detected as {target}")
    if target_name:
        status(f"     Owner name: {target_name}")

    # DV Access Logs live alongside selection.zip and/or inside the extracted return
    status("3/9  Looking for DV Access Log CSVs …")
    dv_uploads: List = []
    dv_syncs: List = []
    dv_files: List[Path] = []
    try:
        from .dv_logs import discover_dv_log_files, parse_all_dv_logs, write_dv_excel

        search_roots = [
            selection_zip.parent,  # files placed next to selection.zip
            extracted,
            case_root,
        ]
        dv_files = discover_dv_log_files(*search_roots)
        if dv_files:
            status(f"     Found {len(dv_files)} DV log file(s): "
                   + ", ".join(p.name for p in dv_files))
            dv_uploads, dv_syncs = parse_all_dv_logs(dv_files)
            write_dv_excel(
                dv_uploads, dv_syncs, output_dir / "DV Access Logs.xlsx"
            )
            status(
                f"     Parsed {len(dv_uploads)} upload events, "
                f"{len(dv_syncs)} sync events"
            )
        else:
            status("     No DV Access Log CSVs found (optional)")
    except Exception as e:
        logger.error("DV Access Log processing failed: %s", e)
        traceback.print_exc()

    status("4/9  Decrypting payloads (this can take a long time for large files) …")
    decrypted = stage_decrypt_payloads(payloads, passphrase, work_dir)

    main_content: Optional[Path] = None
    if "main" in decrypted:
        status("5/9  Extracting main content archive …")
        main_content = stage_extract_main(decrypted["main"], work_dir)
    else:
        status("5/9  No main content zip found – skipping")

    contacts_xlsx: Optional[Path] = None
    if "contacts" in decrypted:
        status("6/9  Converting contacts to Excel …")
        contacts_xlsx = stage_contacts(decrypted["contacts"], output_dir)
    else:
        status("6/9  No contacts file found – skipping")

    quarantine_records: List = []
    if "quarantine" in decrypted:
        status("7/9  Processing quarantined media …")
        _, quarantine_records = stage_quarantine(
            decrypted["quarantine"], work_dir, output_dir
        )
    else:
        status("7/9  No quarantine archive found – skipping")

    # Build summary dict early so the index can show it even if render fails
    summary_path, summary_dict = write_summary(
        output_dir,
        selection_zip,
        target,
        payloads,
        started,
        target_name=target_name,
        dv_upload_count=len(dv_uploads),
        dv_sync_count=len(dv_syncs),
        dv_files=dv_files,
    )

    vz_records: List = []
    unlinked_mms: List = []
    attachment_records: List = []
    if main_content:
        status("8/9  Collecting media, attachments & unlinked MMS …")
        vz_records, unlinked_mms, attachment_records = stage_media_and_attachments(
            main_content, contacts_xlsx, output_dir
        )

        status("9/9  Rendering chat transcripts & case index …")
        stage_render_transcripts(
            main_content,
            contacts_xlsx,
            target,
            output_dir,
            target_name=target_name,
            summary=summary_dict,
            dv_uploads=dv_uploads,
            dv_syncs=dv_syncs,
            vz_records=vz_records,
            unlinked_mms=unlinked_mms,
            quarantine_records=quarantine_records,
            attachment_records=attachment_records,
        )
    else:
        status("8–9/9  Skipped media/transcript stages (no main content)")
        # Still produce an index with Home / Contacts / DV / Quarantine
        try:
            from .render_transcripts import write_index, load_contacts_for_html
            contacts_rows = load_contacts_for_html(
                str(contacts_xlsx) if contacts_xlsx else None
            )
            write_index(
                output_dir,
                [],
                summary=summary_dict,
                contacts=contacts_rows,
                dv_uploads=dv_uploads,
                dv_syncs=dv_syncs,
                quarantine_records=quarantine_records,
            )
            status(f"     Wrote index shell → {output_dir / 'index.html'}")
        except Exception as e:
            logger.error("Failed to write index: %s", e)

    status(f"Done. Summary written to {summary_path}")

    if not keep_work_dir:
        shutil.rmtree(work_dir, ignore_errors=True)
        status("Intermediate work directory removed.")
    else:
        status(f"Intermediate files kept at: {work_dir}")

    return output_dir


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main(argv: Optional[List[str]] = None) -> None:
    parser = argparse.ArgumentParser(
        description="Unified Synchronoss selection.zip → organized case folder pipeline."
    )
    parser.add_argument(
        "--selection",
        required=True,
        help="Path to the outer selection.zip (or equivalent) from leotransfers",
    )
    parser.add_argument(
        "--password",
        required=True,
        help="GPG decryption passphrase supplied in the #secure email",
    )
    parser.add_argument(
        "--output",
        required=True,
        help="Destination folder for the organized case data",
    )
    parser.add_argument(
        "--target-name",
        default="",
        help="Optional display name for the target account holder "
             "(e.g. \"Jane Doe\"). Shown in transcripts and the call log.",
    )
    parser.add_argument(
        "--target-number",
        default="",
        help="Rarely needed override for the target phone number. "
             "Normally auto-detected from the encrypted filenames.",
    )
    parser.add_argument(
        "--clean-work",
        action="store_true",
        help="Delete intermediate decryption/extraction folders after success "
             "(default is to keep them for chain-of-custody).",
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable debug logging",
    )
    args = parser.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s  %(levelname)-7s  %(message)s",
        datefmt="%H:%M:%S",
    )

    try:
        run_pipeline(
            selection_zip=args.selection,
            passphrase=args.password,
            output_dir=args.output,
            target_number=args.target_number,
            target_name=args.target_name,
            keep_work_dir=not args.clean_work,
        )
    except Exception as e:
        logger.error("%s", e)
        if args.verbose:
            traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
