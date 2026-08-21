#!/usr/bin/env python3
"""
Parse Synchronoss / Verizon DV (Device Vault) Access Log CSVs.

These logs are typically delivered alongside selection.zip as files named like:
  "Dv Access logs mdn <LCID> <Month> <Year>.csv"

Two forensically useful views are produced:

* **Uploads** – rows that contain a SHA-256 checksum in the querystring
  (file upload events).  Checksums can be cross-referenced with CyberTip
  or other hash lists.
* **Sync events** – remaining rows (app check-ins / conflict-resolve activity
  without a specific file upload).  Useful for device usage patterns and IP history.

The remoteipaddress field is a comma-separated list; the first entry is the
user's IP and subsequent entries are CDN IPs.
"""

from __future__ import annotations

import csv
import logging
import re
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

logger = logging.getLogger("synchronoss.dv_logs")

# Filename pattern used by Synchronoss legal returns
DV_LOG_NAME_RE = re.compile(r"dv\s+access\s+logs", re.IGNORECASE)
CHECKSUM_RE = re.compile(r"checksum=([a-f0-9]{64})", re.IGNORECASE)


@dataclass
class DvLogRow:
    """One normalised row from a DV access log."""

    server_ts: str = ""
    server_ts_dt: Optional[datetime] = None
    user_ip: str = ""
    cdn_ips: str = ""
    device: str = ""
    operation: str = ""
    checksum: str = ""
    lcid: str = ""
    source_file: str = ""
    raw: Dict[str, str] = field(default_factory=dict)


def _ts_utc(value: str) -> Optional[datetime]:
    """Best-effort parse of server_ts into an aware UTC datetime."""
    if not value or not isinstance(value, str):
        return None
    text = value.strip()
    if text.endswith(" UTC"):
        text = text[:-4].strip()
    if text.endswith("Z"):
        text = text[:-1] + "+00:00"
    try:
        dt = datetime.fromisoformat(text)
    except ValueError:
        # Common alternate forms
        for fmt in (
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%d %H:%M:%S.%f",
            "%m/%d/%Y %H:%M:%S",
            "%Y/%m/%d %H:%M:%S",
        ):
            try:
                dt = datetime.strptime(text, fmt)
                return dt.replace(tzinfo=timezone.utc)
            except ValueError:
                continue
        return None
    if dt.tzinfo is None:
        return dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)


def _extract_checksum(querystring: str) -> str:
    m = CHECKSUM_RE.search(querystring or "")
    return m.group(1) if m else ""


def _extract_user_ip(remoteipaddress: str) -> Tuple[str, str]:
    """Return (user_ip, cdn_ips_string). First entry is the user IP."""
    if not remoteipaddress or remoteipaddress.strip() in ("-", ""):
        return "", ""
    parts = [p.strip() for p in remoteipaddress.split(",") if p.strip()]
    if not parts:
        return "", ""
    user_ip = parts[0]
    cdn_ips = ", ".join(parts[1:]) if len(parts) > 1 else ""
    return user_ip, cdn_ips


def _extract_operation(querystring: str) -> str:
    """Pull a short operation name from the querystring (first key)."""
    qs = (querystring or "").strip()
    if not qs or qs == "-":
        return ""
    qs = qs.lstrip("?")
    return qs.split("=")[0] if qs else ""


def discover_dv_log_files(*search_roots: Path) -> List[Path]:
    """Find DV Access Log CSVs under any of the given roots (non-recursive name match + rglob)."""
    found: List[Path] = []
    seen = set()
    for root in search_roots:
        if root is None:
            continue
        root = Path(root)
        if not root.exists():
            continue
        # Direct children first (common placement "alongside the zip")
        for p in root.iterdir() if root.is_dir() else []:
            if p.is_file() and p.suffix.lower() == ".csv" and DV_LOG_NAME_RE.search(p.name):
                real = p.resolve()
                if real not in seen:
                    seen.add(real)
                    found.append(p)
        # Deeper search
        try:
            for p in root.rglob("*.csv"):
                if DV_LOG_NAME_RE.search(p.name):
                    real = p.resolve()
                    if real not in seen:
                        seen.add(real)
                        found.append(p)
        except OSError as e:
            logger.warning("Error scanning for DV logs under %s: %s", root, e)
    return sorted(found, key=lambda p: p.name.lower())


def parse_dv_log_file(csv_path: Path) -> List[DvLogRow]:
    """Parse one DV access log CSV into normalised DvLogRow objects."""
    rows: List[DvLogRow] = []
    try:
        with open(csv_path, "r", encoding="utf-8-sig", errors="replace") as fh:
            reader = csv.DictReader(fh)
            if not reader.fieldnames:
                return rows
            # Normalise header keys once
            field_map = {h.strip().lower(): h for h in reader.fieldnames if h}

            def get(row: dict, *keys: str) -> str:
                for k in keys:
                    orig = field_map.get(k.lower())
                    if orig is not None:
                        return (row.get(orig) or "").strip()
                return ""

            rel = csv_path.name
            for raw in reader:
                qs = get(raw, "querystring", "query_string", "query")
                remote = get(raw, "remoteipaddress", "remote_ip_address", "remoteip", "ip")
                user_ip, cdn_ips = _extract_user_ip(remote)
                checksum = _extract_checksum(qs)
                ts_raw = get(raw, "server_ts", "serverts", "timestamp", "date", "time")
                row = DvLogRow(
                    server_ts=ts_raw,
                    server_ts_dt=_ts_utc(ts_raw),
                    user_ip=user_ip,
                    cdn_ips=cdn_ips,
                    device=get(raw, "clientidentifier", "client_identifier", "device", "client"),
                    operation=_extract_operation(qs),
                    checksum=checksum,
                    lcid=get(raw, "lcid", "mdn", "account"),
                    source_file=rel,
                    raw={k.strip(): (v or "") for k, v in raw.items() if k},
                )
                rows.append(row)
    except (OSError, csv.Error, UnicodeError) as e:
        logger.error("Failed to read DV log %s: %s", csv_path, e)
    return rows


def parse_all_dv_logs(paths: Iterable[Path]) -> Tuple[List[DvLogRow], List[DvLogRow]]:
    """
    Parse every DV log file and split into (upload_rows, sync_rows).

    Upload rows are those with a non-empty SHA-256 checksum.
    Results are sorted by server_ts ascending.
    """
    all_rows: List[DvLogRow] = []
    for p in paths:
        all_rows.extend(parse_dv_log_file(Path(p)))

    def sort_key(r: DvLogRow):
        if r.server_ts_dt:
            return r.server_ts_dt.isoformat()
        return r.server_ts or ""

    all_rows.sort(key=sort_key)
    uploads = [r for r in all_rows if r.checksum]
    syncs = [r for r in all_rows if not r.checksum]
    return uploads, syncs


def write_dv_excel(
    uploads: List[DvLogRow],
    syncs: List[DvLogRow],
    out_path: Path,
) -> Path:
    """Write a two-sheet Excel workbook (Uploads / Sync Events)."""
    try:
        from openpyxl import Workbook
    except ImportError:
        logger.warning("openpyxl not available – skipping DV Excel export")
        return out_path

    wb = Workbook()
    ws_up = wb.active
    ws_up.title = "Uploads"
    ws_up.append(
        [
            "Timestamp (UTC)",
            "User IP",
            "CDN IPs",
            "Device",
            "File Checksum (SHA-256)",
            "LCID",
            "Source File",
        ]
    )
    for r in uploads:
        if r.server_ts_dt:
            ts = r.server_ts_dt.strftime("%Y-%m-%d %H:%M:%S %Z")
        else:
            ts = r.server_ts
        ws_up.append([ts, r.user_ip, r.cdn_ips, r.device, r.checksum, r.lcid, r.source_file])

    ws_sy = wb.create_sheet("Sync Events")
    ws_sy.append(
        [
            "Timestamp (UTC)",
            "User IP",
            "CDN IPs",
            "Device",
            "Operation",
            "LCID",
            "Source File",
        ]
    )
    for r in syncs:
        if r.server_ts_dt:
            ts = r.server_ts_dt.strftime("%Y-%m-%d %H:%M:%S %Z")
        else:
            ts = r.server_ts
        ws_sy.append([ts, r.user_ip, r.cdn_ips, r.device, r.operation, r.lcid, r.source_file])

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    logger.info(
        "Wrote DV Access Logs Excel (%d uploads, %d sync) → %s",
        len(uploads),
        len(syncs),
        out_path,
    )
    return out_path
