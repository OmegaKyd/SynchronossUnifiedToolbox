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
:root {
  --bg: #0f172a;
  --panel: #1e293b;
  --panel2: #1f2937;
  --border: #334155;
  --text: #f1f5f9;
  --muted: #94a3b8;
  --accent: #38bdf8;
  --accent-dim: #0ea5e9;
  --nav-w: 240px;
  --good: #22c55e;
  --warn: #f59e0b;
}
* { box-sizing: border-box; }
html, body { height: 100%; margin: 0; }
body {
  background: var(--bg); color: var(--text);
  font: 15px/1.5 system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
  display: flex; min-height: 100vh;
}
/* ----- Left navigation ----- */
.nav {
  width: var(--nav-w); flex-shrink: 0;
  background: #020617; border-right: 1px solid var(--border);
  display: flex; flex-direction: column;
  position: sticky; top: 0; height: 100vh; overflow-y: auto;
}
.nav-brand {
  padding: 20px 12px 14px; border-bottom: 1px solid var(--border);
  text-align: center;
}
.nav-brand .title {
  font-weight: 700; font-size: 15px; color: var(--accent);
  letter-spacing: 0.02em; line-height: 1.25;
}
.nav-brand .sub {
  font-weight: 600; font-size: 13px; color: var(--text);
  margin-top: 2px; letter-spacing: 0.01em; line-height: 1.25;
}
.nav-links { padding: 12px 8px; display: flex; flex-direction: column; gap: 4px; flex: 1; }
.nav-links a {
  display: block; padding: 10px 12px; border-radius: 8px;
  color: var(--muted); text-decoration: none; font-weight: 500; font-size: 14px;
}
.nav-links a:hover { background: #1e293b; color: var(--text); }
.nav-links a.active { background: #0c4a6e; color: #e0f2fe; }
.nav-footer {
  padding: 12px 16px; border-top: 1px solid var(--border);
  font-size: 11px; color: var(--muted);
}
/* ----- Main content ----- */
.main { flex: 1; overflow-y: auto; min-width: 0; }
/* Use the full main pane width so tables can spread across the viewport */
.section { display: none; padding: 28px 28px 60px; max-width: none; width: 100%; box-sizing: border-box; }
.section.active { display: block; }
.section h1 { margin: 0 0 6px; font-size: 24px; }
.section .subtitle { color: var(--muted); margin-bottom: 20px; font-size: 13px; max-width: 72rem; }
/* Home cards */
.cards { display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 12px; margin-bottom: 24px; }
.card {
  background: var(--panel2); border: 1px solid var(--border); border-radius: 12px;
  padding: 14px 16px;
}
.card .label { font-size: 11px; text-transform: uppercase; letter-spacing: 0.04em; color: var(--muted); }
.card .value { font-size: 20px; font-weight: 700; margin-top: 4px; color: var(--text); word-break: break-all; }
.card .value.small { font-size: 14px; font-weight: 600; }
.summary-block {
  background: var(--panel2); border: 1px solid var(--border); border-radius: 12px;
  padding: 16px 18px; margin-bottom: 16px;
}
.summary-block h2 { margin: 0 0 10px; font-size: 15px; color: var(--accent); }
.summary-block pre, .summary-block .mono {
  font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
  font-size: 12px; color: #cbd5e1; white-space: pre-wrap; margin: 0;
}
.kv { display: grid; grid-template-columns: 160px 1fr; gap: 6px 12px; font-size: 13px; }
.kv .k { color: var(--muted); }
.kv .v { color: var(--text); word-break: break-word; }
/* Conversations list */
.search-bar { margin-bottom: 16px; }
.search-input {
  width: 100%; max-width: 420px; padding: 8px 12px; border-radius: 8px;
  border: 1px solid var(--border); background: var(--panel2); color: var(--text);
}
.list { display: grid; gap: 10px; }
.item {
  background: var(--panel2); border: 1px solid var(--border); border-radius: 12px; padding: 12px 14px;
}
.item a { color: #7dd3fc; text-decoration: none; font-weight: 600; }
.item a:hover { text-decoration: underline; }
.meta { color: var(--muted); font-size: 12px; margin-top: 4px; }
/* Tables — viewport-bounded so horizontal scrollbar sits at bottom of the pane */
.table-wrap {
  overflow: auto;
  max-height: calc(100vh - 240px);
  border: 1px solid var(--border);
  border-radius: 12px;
  width: 100%;
}
.filter-bar {
  display: flex; flex-wrap: wrap; gap: 10px 14px; align-items: center;
  margin-bottom: 12px;
}
.filter-label {
  display: inline-flex; align-items: center; gap: 6px;
  font-size: 12px; color: var(--muted); font-weight: 500;
}
.filter-select {
  padding: 6px 10px; border-radius: 8px; border: 1px solid var(--border);
  background: var(--panel2); color: var(--text); font-size: 13px;
  max-width: 200px;
}
.filter-count {
  font-size: 12px;
  color: var(--muted);
  margin-left: auto;
  white-space: nowrap;
  font-weight: 500;
}
table.data {
  width: 100%; border-collapse: collapse; font-size: 13px;
  table-layout: fixed;
}
table.data th, table.data td {
  padding: 8px 16px; text-align: left; border-bottom: 1px solid var(--border);
  vertical-align: middle;
  font-size: 13px;
  font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
}
table.data th {
  background: #0f172a; color: var(--muted); font-weight: 600;
  position: sticky; top: 0; z-index: 1;
  user-select: none;
}
table.data tr:hover td { background: rgba(56, 189, 248, 0.06); }
table.data .mono {
  font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
  font-size: 13px;
}
table.data td.col-md5,
table.data th.col-md5 {
  font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
  font-size: 13px;
}
/* Explicit widths (fixed layout respects these) */
/* Preview column: fixed via <col> + cell rules (same on all media tables) */
table.data col.col-preview { width: 72px; }
table.data th.col-preview,
table.data td.col-preview {
  width: 72px;
  min-width: 72px;
  max-width: 72px;
  box-sizing: border-box;
  padding-left: 6px;
  padding-right: 6px;
  text-align: center;
  vertical-align: middle;
  overflow: hidden;
}
table.data col.col-info,
table.data th.col-info,
table.data td.col-info {
  width: 44px;
  min-width: 44px;
  max-width: 44px;
  text-align: center;
  vertical-align: middle;
  padding: 4px;
}
.info-btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 22px;
  height: 22px;
  border-radius: 50%;
  border: 1.5px solid #38bdf8;
  background: #0c4a6e;
  color: #e0f2fe;
  font-size: 12px;
  font-weight: 700;
  font-style: italic;
  font-family: Georgia, "Times New Roman", serif;
  cursor: pointer;
  line-height: 1;
  padding: 0;
  transition: background 0.15s, border-color 0.15s, transform 0.1s;
}
.info-btn:hover,
.info-btn:focus {
  background: #0369a1;
  border-color: #7dd3fc;
  outline: none;
  transform: scale(1.08);
}
.info-btn.active {
  background: #0284c7;
  border-color: #bae6fd;
}
/* Metadata popover */
.meta-popover {
  position: fixed;
  z-index: 9999;
  min-width: 260px;
  max-width: 380px;
  max-height: 70vh;
  overflow: auto;
  background: #0f172a;
  border: 1px solid #334155;
  border-radius: 12px;
  box-shadow: 0 12px 40px rgba(0,0,0,0.55);
  padding: 0;
  display: none;
  font-size: 12px;
  color: var(--text);
}
.meta-popover.open { display: block; }
.meta-popover .meta-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 10px 12px;
  border-bottom: 1px solid #1e293b;
  background: #1e293b;
  border-radius: 12px 12px 0 0;
  font-weight: 600;
  font-size: 13px;
  color: #e0f2fe;
}
.meta-popover .meta-close {
  background: transparent;
  border: none;
  color: #94a3b8;
  font-size: 18px;
  line-height: 1;
  cursor: pointer;
  padding: 0 4px;
}
.meta-popover .meta-close:hover { color: #f1f5f9; }
.meta-popover .meta-body {
  padding: 8px 0;
}
.meta-popover .meta-row {
  display: grid;
  grid-template-columns: 38% 62%;
  gap: 6px 10px;
  padding: 5px 12px;
  border-bottom: 1px solid #1e293b;
}
.meta-popover .meta-row:last-child { border-bottom: none; }
.meta-popover .meta-key {
  color: #94a3b8;
  font-weight: 500;
  word-break: break-word;
}
.meta-popover .meta-val {
  color: #e2e8f0;
  word-break: break-word;
  font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
  font-size: 11.5px;
}
/* Quarantine: fewer columns stretch Preview unless forced narrower (~1/3 prior width) */
#quarantine-table { table-layout: fixed; width: 100%; }
/* Preview stays the same width as other media tables (72px) */
#quarantine-table col.col-preview { width: 72px; }
#quarantine-table th.col-preview,
#quarantine-table td.col-preview {
  width: 72px !important;
  min-width: 72px !important;
  max-width: 72px !important;
  padding-left: 6px !important;
  padding-right: 6px !important;
}
#quarantine-table img.thumb { width: 56px; height: 56px; }
#quarantine-table col.col-filename { width: 52%; }
#quarantine-table col.col-md5 { width: 30%; }
#quarantine-table col.col-type { width: 8%; }
#quarantine-table th.col-filename,
#quarantine-table td.col-filename {
  width: 52%;
  overflow-wrap: anywhere;
  word-break: break-word;
}
#quarantine-table th.col-md5,
#quarantine-table td.col-md5 {
  width: 30%;
  white-space: normal;
  overflow-wrap: anywhere;
  word-break: break-all;
  font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
  font-size: 12px;
}
#quarantine-table th.col-type,
#quarantine-table td.col-type {
  width: 8%;
  min-width: 4em;
  white-space: nowrap;
}
table.data th.col-filename,
table.data td.col-filename {
  width: 36%;
  overflow-wrap: anywhere;
  word-break: break-word;
}
table.data th.col-date,
table.data td.col-date {
  width: 12rem;
  white-space: normal;
  overflow-wrap: anywhere;
  word-break: break-word;
}
table.data th.col-direction,
table.data td.col-direction {
  width: 4.5em;
  min-width: 4em;
  max-width: 5.5em;
  white-space: nowrap;
  text-align: center;
  padding-left: 6px;
  padding-right: 6px;
}
#call-logs-table th.col-date,
#call-logs-table td.col-date {
  width: 16rem;
  white-space: normal;
  overflow-wrap: anywhere;
  word-break: break-word;
}
#call-logs-table th.col-msgid,
#call-logs-table td.col-msgid {
  width: 22%;
  white-space: normal;
  overflow-wrap: anywhere;
  word-break: break-all;
  font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
  font-size: 12px;
}
table.data th.col-device,
table.data td.col-device {
  width: 24%;
  white-space: normal;
  overflow-wrap: anywhere;
  word-break: break-word;
}
table.data th.col-type,
table.data td.col-type {
  width: 12%;
  white-space: normal;
  overflow-wrap: anywhere;
  word-break: break-word;
}
/* Contacts: keep boolean / short-flag columns tight */
#contacts-table th.col-flag,
#contacts-table td.col-flag {
  width: 4.5em;
  min-width: 4em;
  max-width: 5.5em;
  white-space: nowrap;
  text-align: center;
  padding-left: 6px;
  padding-right: 6px;
}
#contacts-table th.col-source,
#contacts-table td.col-source {
  width: 5.5em;
  min-width: 4.5em;
  max-width: 7em;
  white-space: nowrap;
  text-align: center;
  padding-left: 6px;
  padding-right: 6px;
}
#contacts-table th.col-phone,
#contacts-table td.col-phone {
  width: 18%;
  white-space: normal;
  overflow-wrap: anywhere;
  word-break: break-word;
}
/* Hashes / long mono values must wrap so they never bleed into the next column */
table.data th.col-md5,
table.data td.col-md5,
table.data th.col-hash,
table.data td.col-hash {
  width: 18%;
  white-space: normal;
  overflow-wrap: anywhere;
  word-break: break-all;
}
table.data th.col-ip,
table.data td.col-ip {
  width: 9rem;
  white-space: normal;
  overflow-wrap: anywhere;
  word-break: break-all;
}
table.data th.col-lcid,
table.data td.col-lcid {
  width: 12%;
  min-width: 6rem;
  white-space: normal;
  overflow-wrap: anywhere;
  word-break: break-all;
  font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
  font-size: 12px;
}
table.data th.col-source,
table.data td.col-source {
  width: 12%;
  min-width: 5rem;
  white-space: normal;
  overflow-wrap: anywhere;
  word-break: break-word;
}
/* DV tables: keep long hash / source columns from colliding */
#dv-uploads-table th.col-hash,
#dv-uploads-table td.col-hash,
#dv-sync-table th.col-hash,
#dv-sync-table td.col-hash {
  width: 16%;
}
#dv-uploads-table th.col-date,
#dv-uploads-table td.col-date,
#dv-sync-table th.col-date,
#dv-sync-table td.col-date {
  width: 11rem;
  white-space: nowrap;
}
#dv-uploads-table th.col-device,
#dv-uploads-table td.col-device,
#dv-sync-table th.col-device,
#dv-sync-table td.col-device {
  width: 12%;
  white-space: normal;
  overflow-wrap: anywhere;
  word-break: break-word;
}
/* Header tooltips (native title=) */
table.data th[title] {
  cursor: help;
}
/* Sortable headers */
table.data th.sortable {
  cursor: pointer;
}
table.data th.sortable:hover { color: var(--accent); }
table.data th.sortable .sort-ind {
  margin-left: 4px;
  font-size: 10px;
  opacity: 0.7;
}
/* Toolbar buttons (export, etc.) */
.toolbar {
  display: flex; flex-wrap: wrap; gap: 8px; align-items: center;
  margin-bottom: 14px;
}
.btn {
  padding: 7px 12px; border-radius: 8px; border: 1px solid var(--border);
  background: #0c4a6e; color: #e0f2fe; cursor: pointer; font-size: 13px;
  font-weight: 500;
}
.btn:hover { background: #0369a1; }
.btn.secondary {
  background: transparent; color: var(--muted);
}
.btn.secondary:hover { color: var(--text); border-color: var(--muted); }
/* Modal prompt (Yes / No) */
.modal-overlay {
  display: none; position: fixed; inset: 0; z-index: 1000;
  background: rgba(2, 6, 23, 0.72);
  align-items: center; justify-content: center; padding: 20px;
}
.modal-overlay.open { display: flex; }
.modal-dialog {
  background: var(--panel2); border: 1px solid var(--border); border-radius: 14px;
  max-width: 420px; width: 100%; padding: 22px 24px; box-shadow: 0 16px 48px rgba(0,0,0,0.45);
}
.modal-dialog h2 { margin: 0 0 10px; font-size: 17px; color: var(--text); }
.modal-dialog p { margin: 0 0 18px; font-size: 14px; color: var(--muted); line-height: 1.45; }
.modal-actions { display: flex; gap: 10px; justify-content: flex-end; flex-wrap: wrap; }
.modal-actions .btn { min-width: 72px; }
/* Drag handle on the right edge of each header */
table.data th .col-resizer {
  position: absolute;
  top: 0; right: -3px; bottom: 0;
  width: 9px;
  cursor: col-resize;
  user-select: none;
  touch-action: none;
  z-index: 2;
}
table.data th .col-resizer:hover,
table.data th .col-resizer.active {
  background: rgba(56, 189, 248, 0.5);
}
body.col-resizing,
body.col-resizing * { cursor: col-resize !important; user-select: none !important; }
.tabs { display: flex; gap: 6px; margin-bottom: 14px; flex-wrap: wrap; }
.tab-btn {
  padding: 7px 14px; border-radius: 999px; border: 1px solid var(--border);
  background: transparent; color: var(--muted); cursor: pointer; font-size: 13px;
}
.tab-btn.active { background: #0c4a6e; color: #e0f2fe; border-color: #0369a1; }
.badge {
  display: inline-block; padding: 2px 8px; border-radius: 999px;
  font-size: 11px; font-weight: 600; background: #1e3a5f; color: #7dd3fc;
}
.empty { color: var(--muted); padding: 24px; text-align: center; font-size: 14px; }
img.thumb {
  width: 56px; height: 56px; border-radius: 6px;
  border: 1px solid var(--border); object-fit: cover;
  display: block; margin: 0 auto;
}
@media (max-width: 720px) {
  body { flex-direction: column; }
  .nav { width: 100%; height: auto; position: relative; border-right: none; border-bottom: 1px solid var(--border); }
  .nav-links { flex-direction: row; flex-wrap: wrap; }
  .section { padding: 20px 16px 40px; }
}
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
    index_href: str = "index.html#conversations",
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
        f"  <div class=\"container\">Return to "
        f"<a href=\"{html.escape(index_href)}\">Conversations</a></div>"
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


def load_contacts_for_html(xlsx_path: Optional[str]) -> List[Dict[str, str]]:
    """Load contacts from the Excel workbook into a list of simple dicts for HTML."""
    rows: List[Dict[str, str]] = []
    if not xlsx_path:
        return rows
    try:
        df = pd.read_excel(xlsx_path)
    except Exception:
        return rows
    preferred = [
        "firstname", "lastname", "phone_numbers", "phone_types",
        "source", "created", "deleted", "incaseofemergency", "favorite", "itemguid",
    ]
    cols = [c for c in preferred if c in df.columns]
    for _, r in df.iterrows():
        item = {}
        for c in cols:
            v = r.get(c)
            if pd.isna(v):
                item[c] = ""
            else:
                item[c] = str(v).strip()
        rows.append(item)
    return rows


def write_index(
    out_dir: Path,
    entries: List[Tuple],
    *,
    summary: Optional[Dict[str, str]] = None,
    contacts: Optional[List[Dict[str, str]]] = None,
    call_logs: Optional[List[Dict[str, str]]] = None,
    dv_uploads: Optional[List] = None,
    dv_syncs: Optional[List] = None,
    vz_records: Optional[List[Dict[str, str]]] = None,
    unlinked_mms: Optional[List[Dict[str, str]]] = None,
    quarantine_records: Optional[List[Dict[str, str]]] = None,
    attachment_records: Optional[List[Dict[str, str]]] = None,
) -> None:
    """Write the master case index with left navigation."""
    summary = summary or {}
    contacts = contacts or []
    call_logs = call_logs or []
    dv_uploads = dv_uploads or []
    dv_syncs = dv_syncs or []
    vz_records = vz_records or []
    unlinked_mms = unlinked_mms or []
    quarantine_records = quarantine_records or []
    attachment_records = attachment_records or []

    total_msgs = sum(e[2] for e in entries)
    total_att = sum(e[3] for e in entries)

    IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp"}

    def _file_type(fname: str, fallback: str = "") -> str:
        ext = Path(fname or "").suffix.lower()
        if ext:
            return ext
        fb = (fallback or "").strip().lower()
        if fb.startswith("."):
            return fb.split()[0]
        return fb or "(none)"

    def _date_only(raw: str) -> str:
        """Strip time from ISO / datetime strings → YYYY-MM-DD when possible."""
        s = (raw or "").strip()
        if not s:
            return ""
        # Prefer YYYY-MM-DD prefix (handles "2026-07-10 17:32:50 Central Daylight Time")
        if len(s) >= 10 and s[4] == "-" and s[7] == "-" and s[0:4].isdigit() and s[5:7].isdigit() and s[8:10].isdigit():
            return s[:10]
        # ISO-style separator: digit + T + digit (avoid matching the word "Time")
        for i, ch in enumerate(s):
            if ch == "T" and i > 0 and i + 1 < len(s) and s[i - 1].isdigit() and s[i + 1].isdigit():
                return s[:i]
        if " " in s:
            return s.split(" ", 1)[0]
        return s

    def _filter_select(select_id: str, label: str, values: List[str]) -> str:
        opts = ['<option value="">All</option>']
        seen = set()
        for v in sorted((x for x in values if x), key=lambda x: x.lower()):
            key = v.lower()
            if key in seen:
                continue
            seen.add(key)
            opts.append(
                f'<option value="{html.escape(v)}">{html.escape(v)}</option>'
            )
        return (
            f'<label class="filter-label">{html.escape(label)} '
            f'<select id="{html.escape(select_id)}" class="filter-select">'
            f'{"".join(opts)}</select></label>'
        )

    lines: List[str] = []
    lines.append("<!DOCTYPE html>")
    lines.append('<html lang="en"><head>')
    lines.append('<meta charset="utf-8">')
    lines.append('<meta name="viewport" content="width=device-width, initial-scale=1">')
    lines.append("<title>Synchronoss Case Index</title>")
    lines.append(f"<style>{INDEX_CSS}</style>")
    lines.append("</head><body>")

    # ----- Navigation pane -----
    lines.append('<nav class="nav" aria-label="Case navigation">')
    lines.append('  <div class="nav-brand">')
    lines.append('    <div class="title">Synchronoss</div>')
    lines.append('    <div class="sub">Unified Toolbox</div>')
    lines.append("  </div>")
    n_calls = len(call_logs)
    n_contacts = len(contacts)
    n_conv = len(entries)
    n_att = len(attachment_records)
    n_ul = len(unlinked_mms)
    n_dv = len(dv_uploads) + len(dv_syncs)
    n_vz = len(vz_records)
    n_q = len(quarantine_records)

    # Unique user IPs from DV access logs (uploads + syncs) with first/last seen
    dv_ip_map: Dict[str, Dict[str, object]] = {}
    for r in list(dv_uploads) + list(dv_syncs):
        ip = (getattr(r, "user_ip", None) or "").strip()
        if not ip:
            continue
        dt = getattr(r, "server_ts_dt", None)
        if dt is not None:
            ts_display = dt.strftime("%Y-%m-%d %H:%M:%S UTC")
            ts_key = dt.isoformat()
        else:
            ts_raw = (getattr(r, "server_ts", None) or "").strip()
            ts_display = ts_raw
            ts_key = ts_raw
        device = (getattr(r, "device", None) or "").strip()
        entry = dv_ip_map.get(ip)
        if entry is None:
            dv_ip_map[ip] = {
                "ip": ip,
                "first_key": ts_key,
                "last_key": ts_key,
                "first_seen": ts_display,
                "last_seen": ts_display,
                "events": 1,
                "devices": {device} if device else set(),
            }
        else:
            entry["events"] = int(entry["events"]) + 1
            if device:
                entry["devices"].add(device)  # type: ignore[union-attr]
            if ts_key and (not entry["first_key"] or ts_key < str(entry["first_key"])):
                entry["first_key"] = ts_key
                entry["first_seen"] = ts_display
            if ts_key and (not entry["last_key"] or ts_key > str(entry["last_key"])):
                entry["last_key"] = ts_key
                entry["last_seen"] = ts_display
    dv_ip_rows: List[Dict[str, str]] = []
    for ip, entry in sorted(
        dv_ip_map.items(),
        key=lambda kv: str(kv[1].get("last_key") or ""),
        reverse=True,
    ):
        devices = entry.get("devices") or set()
        dv_ip_rows.append(
            {
                "ip": ip,
                "first_seen": str(entry.get("first_seen") or ""),
                "last_seen": str(entry.get("last_seen") or ""),
                "events": str(entry.get("events") or 0),
                "devices": ", ".join(sorted(d for d in devices if d)),
            }
        )
    n_dv_ips = len(dv_ip_rows)

    lines.append('  <div class="nav-links">')
    lines.append('    <a href="#home" class="active" data-section="home">Home</a>')
    lines.append(
        f'    <a href="#call-logs" data-section="call-logs">Call Logs ({n_calls})</a>'
    )
    lines.append(
        f'    <a href="#contacts" data-section="contacts">Contacts ({n_contacts})</a>'
    )
    lines.append(
        f'    <a href="#conversations" data-section="conversations">'
        f'Conversations ({n_conv})</a>'
    )
    lines.append(
        f'    <a href="#attachments" data-section="attachments">'
        f'Compiled Attachments ({n_att})</a>'
    )
    lines.append(
        f'    <a href="#unlinked-mms" data-section="unlinked-mms">'
        f'Unlinked MMS ({n_ul})</a>'
    )
    lines.append(
        f'    <a href="#dv-logs" data-section="dv-logs">DV Access Logs ({n_dv})</a>'
    )
    lines.append(
        f'    <a href="#dv-ips" data-section="dv-ips">DV Unique IPs ({n_dv_ips})</a>'
    )
    lines.append(
        f'    <a href="#vzmobile" data-section="vzmobile">VZMOBILE Media ({n_vz})</a>'
    )
    lines.append(
        f'    <a href="#quarantine" data-section="quarantine">'
        f'Quarantined Files ({n_q})</a>'
    )
    lines.append("  </div>")
    lines.append('  <div class="nav-footer">Offline · file:// safe</div>')
    lines.append("</nav>")

    lines.append('<div class="main">')

    # ===== HOME =====
    lines.append('<section id="home" class="section active">')
    lines.append("  <h1>Case Overview</h1>")
    lines.append(
        '  <div class="subtitle">Processing summary and high-level counts. '
        "Open <strong>Conversations</strong> to review chat transcripts.</div>"
    )
    lines.append('  <div class="cards">')
    owner = summary.get("owner_phone") or "—"
    owner_name = summary.get("owner_name") or ""
    lines.append(
        f'    <div class="card"><div class="label">Owner phone</div>'
        f'<div class="value small">{html.escape(owner)}</div></div>'
    )
    lines.append(
        f'    <div class="card"><div class="label">Owner name</div>'
        f'<div class="value small">{html.escape(owner_name or "(not supplied)")}</div></div>'
    )
    lines.append(
        f'    <div class="card"><div class="label">Conversations</div>'
        f'<div class="value">{len(entries)}</div></div>'
    )
    lines.append(
        f'    <div class="card"><div class="label">Messages</div>'
        f'<div class="value">{total_msgs}</div></div>'
    )
    lines.append(
        f'    <div class="card"><div class="label">Msgs with attachments</div>'
        f'<div class="value">{total_att}</div></div>'
    )
    lines.append(
        f'    <div class="card"><div class="label">Contacts</div>'
        f'<div class="value">{len(contacts)}</div></div>'
    )
    lines.append(
        f'    <div class="card"><div class="label">Call records</div>'
        f'<div class="value">{len(call_logs)}</div></div>'
    )
    lines.append(
        f'    <div class="card"><div class="label">Compiled attachments</div>'
        f'<div class="value">{len(attachment_records)}</div></div>'
    )
    lines.append(
        f'    <div class="card"><div class="label">Unlinked MMS</div>'
        f'<div class="value">{len(unlinked_mms)}</div></div>'
    )
    lines.append(
        f'    <div class="card"><div class="label">VZMOBILE media</div>'
        f'<div class="value">{len(vz_records)}</div></div>'
    )
    lines.append(
        f'    <div class="card"><div class="label">Quarantined files</div>'
        f'<div class="value">{len(quarantine_records)}</div></div>'
    )
    lines.append(
        f'    <div class="card"><div class="label">DV uploads</div>'
        f'<div class="value">{len(dv_uploads)}</div></div>'
    )
    lines.append(
        f'    <div class="card"><div class="label">DV sync events</div>'
        f'<div class="value">{len(dv_syncs)}</div></div>'
    )
    lines.append(
        f'    <div class="card"><div class="label">DV unique IPs</div>'
        f'<div class="value">{n_dv_ips}</div></div>'
    )
    lines.append("  </div>")

    lines.append('  <div class="summary-block">')
    lines.append("    <h2>Processing details</h2>")
    lines.append('    <div class="kv">')
    for key, label in (
        ("started", "Started"),
        ("finished", "Finished"),
        ("input_archive", "Input archive"),
        ("output_folder", "Output folder"),
        ("payloads", "Payloads found"),
    ):
        val = summary.get(key, "")
        if val:
            lines.append(f'      <div class="k">{html.escape(label)}</div>')
            lines.append(f'      <div class="v">{html.escape(str(val))}</div>')
    lines.append("    </div>")
    if summary.get("notes"):
        lines.append(
            f'    <pre style="margin-top:12px">{html.escape(summary["notes"])}</pre>'
        )
    lines.append("  </div>")

    lines.append('  <div class="summary-block">')
    lines.append("    <h2>Output layout</h2>")
    lines.append(
        '    <div class="mono">index.html                 – this page (nav sections below)\n'
        "conversations/             – one folder per chat thread\n"
        "contacts.xlsx              – converted contact list\n"
        "Call Log.xlsx              – voice calls with names\n"
        "Unlinked MMS.xlsx / folder – MMS media not tied to a message token\n"
        "DV Access Logs.xlsx        – upload &amp; sync events (when DV CSVs present)\n"
        "Compiled Media/            – VZMOBILE device backup media\n"
        "Compiled Attachments/      – MMS/RCS attachments (if present)\n"
        "Compiled Quarantine Files/ – recovered quarantined media\n"
        "processing_summary.txt     – plain-text summary</div>"
    )
    lines.append("  </div>")
    lines.append("</section>")

    # ===== CONVERSATIONS =====
    lines.append('<section id="conversations" class="section">')
    lines.append("  <h1>Conversations</h1>")
    lines.append(
        '  <div class="subtitle">One HTML per chat. Search matches chat titles '
        "<em>and</em> message content. Times inside each transcript use the local system timezone.</div>"
    )
    lines.append(
        '  <div class="search-bar">'
        '<input id="search" class="search-input" '
        'placeholder="Search chats &amp; message content…">'
        "</div>"
    )
    lines.append('  <div class="list" id="chat-list">')

    search_blobs: List[str] = []
    for i, entry in enumerate(entries):
        title, rel, c, ca = entry[0], entry[1], entry[2], entry[3]
        blob = entry[4] if len(entry) > 4 else ""
        search_blobs.append(blob or "")
        lines.append(f'    <div class="item" data-i="{i}">')
        lines.append(f'      <a href="{html.escape(rel)}">{html.escape(title)}</a>')
        lines.append(
            f'      <div class="meta">Messages: {c} · Messages with attachments: {ca}</div>'
        )
        lines.append("    </div>")
    if not entries:
        lines.append('    <div class="empty">No conversations were rendered.</div>')
    lines.append("  </div>")
    lines.append(
        '  <div id="search-status" class="subtitle" style="margin-top:12px;display:none"></div>'
    )
    lines.append("</section>")

    # ===== CONTACTS =====
    lines.append('<section id="contacts" class="section">')
    lines.append("  <h1>Contacts</h1>")
    lines.append(
        f'  <div class="subtitle">{len(contacts)} contact'
        f'{"s" if len(contacts) != 1 else ""} from the Synchronoss contacts export. '
        "Also available as <code>contacts.xlsx</code>.</div>"
    )
    lines.append(
        '  <div class="search-bar">'
        '<input id="contact-search" class="search-input" placeholder="Filter contacts…">'
        f'<span class="filter-count" id="contact-filter-count">{len(contacts)} of {len(contacts)} contacts displayed</span>'
        "</div>"
    )
    if contacts:
        lines.append('<div class="table-wrap"><table class="data" id="contacts-table">')
        headers = [
            ("firstname", "First Name"),
            ("lastname", "Last Name"),
            ("phone_numbers", "Phone Number(s)"),
            ("phone_types", "Phone Type(s)"),
            ("created", "Created"),
            ("deleted", "Deleted"),
            ("incaseofemergency", "ICE"),
            ("favorite", "Favorite"),
            ("source", "Source"),  # furthest right
        ]
        present = [h for h in headers if any(c.get(h[0]) for c in contacts)]
        # Always show core columns even if empty
        for core in ("firstname", "lastname", "phone_numbers"):
            if not any(h[0] == core for h in present):
                for h in headers:
                    if h[0] == core:
                        present.insert(0 if core != "phone_numbers" else len(present), h)
                        break

        def _contact_col_class(key: str) -> str:
            if key in ("incaseofemergency", "favorite"):
                return "col-flag"
            if key == "source":
                return "col-source"
            if key in ("phone_numbers", "phone_types"):
                return "col-phone mono" if key == "phone_numbers" else "col-phone"
            if key in ("itemguid",):
                return "mono"
            return ""

        lines.append("  <thead><tr>")
        for key, label in present:
            cls = _contact_col_class(key)
            cls_attr = f' class="{cls}"' if cls else ""
            lines.append(f"    <th{cls_attr}>{html.escape(label)}</th>")
        lines.append("  </tr></thead><tbody>")
        for c in contacts:
            cells = " ".join(
                f'<td class="{_contact_col_class(key)}">'
                f'{html.escape(c.get(key, ""))}</td>'
                for key, _ in present
            )
            blob = " ".join(c.get(k, "") for k, _ in present).lower()
            lines.append(f'  <tr data-search="{html.escape(blob)}">{cells}</tr>')
        lines.append("</tbody></table></div>")
    else:
        lines.append(
            '<div class="empty">No contacts file was found or it could not be parsed.</div>'
        )
    lines.append("</section>")

    # ===== CALL LOGS =====
    lines.append('<section id="call-logs" class="section">')
    lines.append("  <h1>Call Logs</h1>")
    lines.append(
        f'  <div class="subtitle">{len(call_logs)} call record'
        f'{"s" if len(call_logs) != 1 else ""} from message CSVs (Type = call). '
        "Contact names are shown when known. Also available as "
        "<code>Call Log.xlsx</code>.</div>"
    )
    lines.append(
        '  <div class="search-bar">'
        '<input id="call-search" class="search-input" '
        'placeholder="Filter by number, name, direction…">'
        f'<span class="filter-count" id="call-filter-count">{len(call_logs)} of {len(call_logs)} records displayed</span>'
        "</div>"
    )
    if call_logs:
        call_dirs = [r.get("direction", "") for r in call_logs]
        call_dates = [_date_only(r.get("date", "")) for r in call_logs]
        lines.append('  <div class="filter-bar">')
        lines.append(_filter_select("call-filter-direction", "Direction", call_dirs))
        lines.append(_filter_select("call-filter-date", "Date", call_dates))
        lines.append("  </div>")
        lines.append('<div class="table-wrap"><table class="data" id="call-logs-table">')
        lines.append(
            "  <thead><tr>"
            "<th class=\"col-date sortable\" data-col=\"0\">Date</th>"
            "<th class=\"col-direction sortable\" data-col=\"1\">Direction</th>"
            "<th class=\"sortable\" data-col=\"2\">Sender</th>"
            "<th class=\"sortable\" data-col=\"3\">Recipients</th>"
            "<th class=\"col-msgid sortable\" data-col=\"4\">Message ID</th>"
            "</tr></thead><tbody>"
        )
        for row in call_logs:
            date_s = row.get("date", "")
            direction = row.get("direction", "")
            sender = row.get("sender", "")
            recipients = row.get("recipients", "")
            mid = row.get("message_id", "")
            date_key = _date_only(date_s)
            blob = " ".join([date_s, direction, sender, recipients, mid]).lower()
            lines.append(
                f'  <tr data-search="{html.escape(blob)}" '
                f'data-direction="{html.escape(direction)}" '
                f'data-date="{html.escape(date_key)}">'
                f'<td class="col-date mono">{html.escape(date_s)}</td>'
                f'<td class="col-direction">{html.escape(direction)}</td>'
                f'<td>{html.escape(sender)}</td>'
                f'<td>{html.escape(recipients)}</td>'
                f'<td class="col-msgid">{html.escape(mid)}</td>'
                "</tr>"
            )
        lines.append("</tbody></table></div>")
    else:
        lines.append(
            '<div class="empty">No call records were found in the message CSVs.</div>'
        )
    lines.append("</section>")

    # ===== COMPILED ATTACHMENTS =====
    lines.append('<section id="attachments" class="section">')
    lines.append("  <h1>Compiled Attachments</h1>")
    lines.append(
        '  <div class="subtitle">All message attachments collected into '
        "<code>Compiled Attachments/</code> (with metadata log). "
        "Sender/recipient are filled when a message CSV token matched the file. "
        "Click a column header to sort.</div>"
    )
    lines.append(
        '  <div class="search-bar">'
        '<input id="att-search" class="search-input" '
        'placeholder="Search filename, sender, recipient, MD5…">'
        "</div>"
    )
    if attachment_records:
        att_types = [_file_type(r.get("filename", "")) for r in attachment_records]
        att_dates = [_date_only(r.get("date", "")) for r in attachment_records]
        lines.append('  <div class="filter-bar">')
        lines.append(_filter_select("att-filter-type", "File type", att_types))
        lines.append(_filter_select("att-filter-date", "Date", att_dates))
        lines.append(
            f'    <span class="filter-count" id="att-filter-count">'
            f'{n_att} of {n_att} files displayed</span>'
        )
        lines.append("  </div>")
        lines.append('<div class="table-wrap"><table class="data" id="attachments-table">')
        lines.append(
            "  <thead><tr>"
            "<th class=\"col-preview\">Preview</th>"
            "<th class=\"col-filename sortable\" data-col=\"1\">Filename</th>"
            "<th class=\"col-date sortable\" data-col=\"2\">Date</th>"
            "<th class=\"sortable\" data-col=\"3\">Sender</th>"
            "<th class=\"sortable\" data-col=\"4\">Recipient</th>"
            "<th class=\"col-md5 sortable\" data-col=\"5\">MD5</th>"
            "<th class=\"col-type sortable\" data-col=\"6\">File Type</th>"
            "<th class=\"col-info\">Info</th>"
            "</tr></thead><tbody>"
        )
        for row in attachment_records:
            fname = row.get("filename", "")
            date_s = row.get("date", "")
            sender = row.get("sender", "")
            recipient = row.get("recipient", "")
            md5 = row.get("md5", "")
            rel = row.get("rel_path", "")
            meta = row.get("meta") or {}
            ftype = _file_type(fname)
            date_key = _date_only(date_s)
            blob = " ".join([fname, date_s, date_key, sender, recipient, md5, ftype]).lower()
            ext = Path(fname).suffix.lower()
            if rel and ext in IMAGE_EXTS:
                preview = (
                    f'<a href="{html.escape(rel)}" target="_blank" rel="noopener">'
                    f'<img class="thumb" src="{html.escape(rel)}" alt=""></a>'
                )
            elif rel:
                preview = f'<a href="{html.escape(rel)}" target="_blank" rel="noopener">open</a>'
            else:
                preview = "—"
            if meta:
                meta_json = html.escape(json.dumps(meta, ensure_ascii=False), quote=True)
                info_cell = (
                    f'<button type="button" class="info-btn" title="View metadata" '
                    f'data-meta="{meta_json}" aria-label="Metadata">i</button>'
                )
            else:
                info_cell = ""
            lines.append(
                f'  <tr data-search="{html.escape(blob)}" '
                f'data-type="{html.escape(ftype)}" data-date="{html.escape(date_key)}">'
                f'<td class="col-preview" data-sort="">{preview}</td>'
                f'<td class="col-filename" data-sort="{html.escape(fname.lower())}">{html.escape(fname)}</td>'
                f'<td class="col-date" data-sort="{html.escape(date_s)}">{html.escape(date_s)}</td>'
                f'<td data-sort="{html.escape(sender.lower())}">{html.escape(sender)}</td>'
                f'<td data-sort="{html.escape(recipient.lower())}">{html.escape(recipient)}</td>'
                f'<td class="col-md5" data-sort="{html.escape(md5.lower())}">{html.escape(md5)}</td>'
                f'<td class="col-type" data-sort="{html.escape(ftype.lower())}">{html.escape(ftype)}</td>'
                f'<td class="col-info">{info_cell}</td>'
                "</tr>"
            )
        lines.append("</tbody></table></div>")
    else:
        lines.append(
            '<div class="empty">No message attachments were collected '
            "(or no messages/attachments tree was present).</div>"
        )
    lines.append("</section>")

    # ===== UNLINKED MMS =====
    lines.append('<section id="unlinked-mms" class="section">')
    lines.append("  <h1>Unlinked MMS Media</h1>")
    lines.append(
        '  <div class="subtitle">Files present under '
        "<code>messages/attachments/mms/(in|out)/YYYY-MM-DD/</code> that are "
        "<strong>not</strong> referenced by a real-extension attachment token in any "
        "message CSV. Extensionless names (e.g. <code>0</code>) often appear only via "
        "SMIL placeholders and are not attributed to a specific message. "
        "Upload date is the folder name. Also written to "
        "<code>Unlinked MMS.xlsx</code>.</div>"
    )
    lines.append(
        '  <div class="search-bar">'
        '<input id="unlinked-search" class="search-input" '
        'placeholder="Search filename, direction…">'
        "</div>"
    )
    if unlinked_mms:
        ul_types = [
            _file_type(r.get("filename", ""), r.get("detected_type", ""))
            for r in unlinked_mms
        ]
        ul_dates = [r.get("upload_date", "") for r in unlinked_mms]
        lines.append('  <div class="filter-bar">')
        lines.append(_filter_select("ul-filter-type", "File type", ul_types))
        lines.append(_filter_select("ul-filter-date", "Upload Date", ul_dates))
        lines.append(
            f'    <span class="filter-count" id="ul-filter-count">'
            f'{n_ul} of {n_ul} files displayed</span>'
        )
        lines.append("  </div>")
        lines.append('<div class="table-wrap"><table class="data" id="unlinked-mms-table">')
        lines.append(
            "  <thead><tr>"
            "<th class=\"col-preview\">Preview</th>"
            "<th class=\"col-filename sortable\" data-col=\"1\">Filename</th>"
            "<th class=\"col-date sortable\" data-col=\"2\">Upload Date</th>"
            "<th class=\"col-direction sortable\" data-col=\"3\">Direction</th>"
            "<th class=\"col-type sortable\" data-col=\"4\">File Type</th>"
            "</tr></thead><tbody>"
        )
        for row in unlinked_mms:
            ud = row.get("upload_date", "")
            direction = row.get("direction", "")
            fname = row.get("filename", "")
            dtype = row.get("detected_type", "")
            ftype = _file_type(fname, dtype)
            rel = row.get("rel_path", "")
            blob = " ".join([ud, direction, fname, ftype, rel]).lower()
            check_name = rel or fname
            ext = Path(check_name).suffix.lower()
            if not ext and ftype.startswith("."):
                ext = ftype
            if rel and ext in IMAGE_EXTS:
                preview = (
                    f'<a href="{html.escape(rel)}" target="_blank" rel="noopener">'
                    f'<img class="thumb" src="{html.escape(rel)}" alt=""></a>'
                )
            elif rel:
                preview = f'<a href="{html.escape(rel)}" target="_blank" rel="noopener">open</a>'
            else:
                preview = "—"
            lines.append(
                f'  <tr data-search="{html.escape(blob)}" '
                f'data-type="{html.escape(ftype)}" data-date="{html.escape(ud)}">'
                f'<td class="col-preview" data-sort="">{preview}</td>'
                f'<td class="col-filename" data-sort="{html.escape(fname.lower())}">{html.escape(fname)}</td>'
                f'<td class="col-date" data-sort="{html.escape(ud)}">{html.escape(ud)}</td>'
                f'<td class="col-direction" data-sort="{html.escape(direction.lower())}">{html.escape(direction)}</td>'
                f'<td class="col-type" data-sort="{html.escape(ftype.lower())}">{html.escape(ftype)}</td>'
                "</tr>"
            )
        lines.append("</tbody></table></div>")
    else:
        lines.append(
            '<div class="empty">No unlinked MMS media found '
            "(or no messages/attachments/mms tree was present).</div>"
        )
    lines.append("</section>")

    # ===== VZMOBILE =====
    lines.append('<section id="vzmobile" class="section">')
    lines.append("  <h1>VZMOBILE Media</h1>")
    lines.append(
        '  <div class="subtitle">Media from the '
        "<code>VZMOBILE/YYYY-MM-DD/&lt;device&gt;/</code> cloud backup folders. "
        "Upload date is the folder name (per Synchronoss). Files are also under "
        "<code>Compiled Media/</code> with a full metadata log. "
        "Click a column header to sort.</div>"
    )
    lines.append('  <div class="toolbar">')
    lines.append(
        '    <button type="button" class="btn" id="export-md5-txt" '
        'title="Export MD5 hashes as a .txt file (one per line)">'
        "Export MD5 (.txt)</button>"
    )
    lines.append("  </div>")
    lines.append(
        '  <div class="search-bar">'
        '<input id="vz-search" class="search-input" '
        'placeholder="Search filename, MD5…">'
        "</div>"
    )
    # MD5 lists for offline export (VZ-only vs all media)
    def _md5_list(rows: List[Dict[str, str]]) -> List[str]:
        out: List[str] = []
        seen = set()
        for r in rows:
            h = (r.get("md5") or "").strip().lower()
            if h and h not in seen:
                seen.add(h)
                out.append(h)
        return out

    raw_phone = (summary.get("owner_phone") or "").strip()
    # Keep digits (and leading +) for a filesystem-safe default name
    safe_phone = "".join(ch for ch in raw_phone if ch.isdigit() or ch == "+") or "unknown"
    md5_export_payload = {
        "target_phone": safe_phone,
        "vzmobile": _md5_list(vz_records),
        "attachments": _md5_list(attachment_records),
        "quarantine": _md5_list(quarantine_records),
        "unlinked_mms": _md5_list(unlinked_mms),
    }
    lines.append(
        f"<script>const MD5_EXPORT = {json.dumps(md5_export_payload, ensure_ascii=False)};</script>"
    )
    if vz_records:
        vz_types = [_file_type(r.get("filename", "")) for r in vz_records]
        vz_dates = [r.get("upload_date", "") for r in vz_records]
        vz_devices = [r.get("device", "") for r in vz_records]
        lines.append('  <div class="filter-bar">')
        lines.append(_filter_select("vz-filter-type", "File type", vz_types))
        lines.append(_filter_select("vz-filter-date", "Upload Date", vz_dates))
        lines.append(_filter_select("vz-filter-device", "Device", vz_devices))
        lines.append(
            f'    <span class="filter-count" id="vz-filter-count">'
            f'{n_vz} of {n_vz} files displayed</span>'
        )
        lines.append("  </div>")
        lines.append('<div class="table-wrap"><table class="data" id="vzmobile-table">')
        lines.append(
            "  <colgroup>"
            '<col class="col-preview">'
            '<col class="col-filename">'
            '<col class="col-date">'
            '<col class="col-device">'
            '<col class="col-md5">'
            '<col class="col-type">'
            '<col class="col-info">'
            "</colgroup>"
        )
        lines.append(
            "  <thead><tr>"
            "<th class=\"col-preview\">Preview</th>"
            "<th class=\"col-filename sortable\" data-col=\"1\">Filename</th>"
            "<th class=\"col-date sortable\" data-col=\"2\">Upload Date</th>"
            "<th class=\"col-device sortable\" data-col=\"3\">Device</th>"
            "<th class=\"col-md5 sortable\" data-col=\"4\">MD5</th>"
            "<th class=\"col-type sortable\" data-col=\"5\">File Type</th>"
            "<th class=\"col-info\">Info</th>"
            "</tr></thead><tbody>"
        )
        for row in vz_records:
            ud = row.get("upload_date", "")
            device = row.get("device", "")
            fname = row.get("filename", "")
            md5 = row.get("md5", "")
            rel = row.get("rel_path", "")
            meta = row.get("meta") or {}
            ftype = _file_type(fname)
            blob = " ".join([ud, device, fname, md5, ftype]).lower()
            ext = Path(fname).suffix.lower()
            if rel and ext in IMAGE_EXTS:
                preview = (
                    f'<a href="{html.escape(rel)}" target="_blank" rel="noopener">'
                    f'<img class="thumb" src="{html.escape(rel)}" alt=""></a>'
                )
            elif rel:
                preview = f'<a href="{html.escape(rel)}" target="_blank" rel="noopener">open</a>'
            else:
                preview = "—"
            if meta:
                meta_json = html.escape(json.dumps(meta, ensure_ascii=False), quote=True)
                info_cell = (
                    f'<button type="button" class="info-btn" title="View metadata" '
                    f'data-meta="{meta_json}" aria-label="Metadata">i</button>'
                )
            else:
                info_cell = ""
            lines.append(
                f'  <tr data-search="{html.escape(blob)}" '
                f'data-type="{html.escape(ftype)}" data-date="{html.escape(ud)}" '
                f'data-device="{html.escape(device)}">'
                f'<td class="col-preview" data-sort="">{preview}</td>'
                f'<td class="col-filename" data-sort="{html.escape(fname.lower())}">{html.escape(fname)}</td>'
                f'<td class="col-date" data-sort="{html.escape(ud)}">{html.escape(ud)}</td>'
                f'<td class="col-device" data-sort="{html.escape(device.lower())}">{html.escape(device)}</td>'
                f'<td class="col-md5" data-sort="{html.escape(md5.lower())}">{html.escape(md5)}</td>'
                f'<td class="col-type" data-sort="{html.escape(ftype.lower())}">{html.escape(ftype)}</td>'
                f'<td class="col-info">{info_cell}</td>'
                "</tr>"
            )
        lines.append("</tbody></table></div>")
    else:
        lines.append(
            '<div class="empty">No VZMOBILE media was found in this return.</div>'
        )
    lines.append("</section>")

    # ===== QUARANTINE =====
    lines.append('<section id="quarantine" class="section">')
    lines.append("  <h1>Quarantined Files</h1>")
    lines.append(
        '  <div class="subtitle">Media recovered from the Synchronoss quarantined archive '
        "(<code>*-quarantined.zip.gpg</code>). Files are under "
        "<code>Compiled Quarantine Files/</code>. "
        "Click a column header to sort.</div>"
    )
    lines.append(
        '  <div class="search-bar">'
        '<input id="quarantine-search" class="search-input" '
        'placeholder="Search filename, MD5…">'
        "</div>"
    )
    if quarantine_records:
        q_types = [
            _file_type(r.get("filename", ""), r.get("detected_type", ""))
            for r in quarantine_records
        ]
        lines.append('  <div class="filter-bar">')
        lines.append(_filter_select("q-filter-type", "File type", q_types))
        lines.append(
            f'    <span class="filter-count" id="q-filter-count">'
            f'{n_q} of {n_q} files displayed</span>'
        )
        lines.append("  </div>")
        lines.append('<div class="table-wrap"><table class="data" id="quarantine-table">')
        lines.append(
            "  <colgroup>"
            '<col class="col-preview">'
            '<col class="col-filename">'
            '<col class="col-md5">'
            '<col class="col-type">'
            "</colgroup>"
        )
        lines.append(
            "  <thead><tr>"
            "<th class=\"col-preview\">Preview</th>"
            "<th class=\"col-filename sortable\" data-col=\"1\">Filename</th>"
            "<th class=\"col-md5 sortable\" data-col=\"2\">MD5</th>"
            "<th class=\"col-type sortable\" data-col=\"3\">File Type</th>"
            "</tr></thead><tbody>"
        )
        for row in quarantine_records:
            fname = row.get("filename", "")
            dtype = row.get("detected_type", "")
            ftype = _file_type(fname, dtype)
            md5 = row.get("md5", "")
            rel = row.get("rel_path", "")
            blob = " ".join([fname, ftype, md5]).lower()
            ext = Path(fname).suffix.lower() or (
                ftype if ftype.startswith(".") else ""
            )
            if rel and ext in IMAGE_EXTS:
                preview = (
                    f'<a href="{html.escape(rel)}" target="_blank" rel="noopener">'
                    f'<img class="thumb" src="{html.escape(rel)}" alt=""></a>'
                )
            elif rel:
                preview = f'<a href="{html.escape(rel)}" target="_blank" rel="noopener">open</a>'
            else:
                preview = "—"
            lines.append(
                f'  <tr data-search="{html.escape(blob)}" data-type="{html.escape(ftype)}">'
                f'<td class="col-preview" data-sort="">{preview}</td>'
                f'<td class="col-filename" data-sort="{html.escape(fname.lower())}">{html.escape(fname)}</td>'
                f'<td class="col-md5" data-sort="{html.escape(md5.lower())}">{html.escape(md5)}</td>'
                f'<td class="col-type" data-sort="{html.escape(ftype.lower())}">{html.escape(ftype)}</td>'
                "</tr>"
            )
        lines.append("</tbody></table></div>")
    else:
        lines.append(
            '<div class="empty">No quarantined media was recovered '
            "(or no quarantine archive was present in the return).</div>"
        )
    lines.append("</section>")

    # ===== DV ACCESS LOGS =====
    lines.append('<section id="dv-logs" class="section">')
    lines.append("  <h1>DV Access Logs</h1>")
    lines.append(
        '  <div class="subtitle">Device Vault access events. '
        "<strong>Uploads</strong> contain a SHA-256 file checksum (cross-reference with "
        "CyberTip / known-file hashes). <strong>Sync events</strong> show device "
        "check-ins without a specific file upload. Also written to "
        "<code>DV Access Logs.xlsx</code> when present.</div>"
    )
    lines.append('  <div class="tabs">')
    lines.append(
        f'    <button type="button" class="tab-btn active" data-tab="dv-uploads">'
        f'Uploads <span class="badge">{len(dv_uploads)}</span></button>'
    )
    lines.append(
        f'    <button type="button" class="tab-btn" data-tab="dv-sync">'
        f'Sync Events <span class="badge">{len(dv_syncs)}</span></button>'
    )
    lines.append("  </div>")
    lines.append(
        '  <div class="search-bar">'
        '<input id="dv-search" class="search-input" placeholder="Filter by IP, device, checksum, operation…">'
        f'<span class="filter-count" id="dv-filter-count">{n_dv} of {n_dv} events displayed</span>'
        "</div>"
    )

    # Uploads table
    lines.append('<div id="dv-uploads" class="dv-pane">')
    if dv_uploads:
        lines.append('<div class="table-wrap"><table class="data" id="dv-uploads-table">')
        lines.append(
            "  <thead><tr>"
            "<th class=\"col-date\" title=\"When the server recorded this access event (Coordinated Universal Time).\">Timestamp (UTC)</th>"
            "<th class=\"col-ip\" title=\"Public IP address of the device or user that made the request.\">User IP</th>"
            "<th class=\"col-ip\" title=\"Content Delivery Network addresses that forwarded the request. These are shared infrastructure, not the end user’s IP.\">CDN IPs</th>"
            "<th class=\"col-device\" title=\"Client/device identifier reported by the app (for example a phone model string).\">Device</th>"
            "<th class=\"col-hash\" title=\"SHA-256 hash of the file that was uploaded. Useful for matching against known files or CyberTip hash lists.\">File Checksum (SHA-256)</th>"
            "<th class=\"col-lcid\" title=\"Line / account identifier (LCID) associated with this subscriber line in the Synchronoss / Verizon system.\">LCID</th>"
            "<th class=\"col-source\" title=\"Name of the DV Access Log CSV file this row came from.\">Source</th>"
            "</tr></thead><tbody>"
        )
        for r in dv_uploads:
            if getattr(r, "server_ts_dt", None):
                ts = r.server_ts_dt.strftime("%Y-%m-%d %H:%M:%S UTC")
            else:
                ts = getattr(r, "server_ts", "") or ""
            blob = " ".join(
                [
                    ts,
                    getattr(r, "user_ip", ""),
                    getattr(r, "cdn_ips", ""),
                    getattr(r, "device", ""),
                    getattr(r, "checksum", ""),
                    getattr(r, "lcid", ""),
                    getattr(r, "source_file", ""),
                ]
            ).lower()
            lines.append(
                f'  <tr data-search="{html.escape(blob)}">'
                f'<td class="col-date">{html.escape(ts)}</td>'
                f'<td class="col-ip">{html.escape(getattr(r, "user_ip", ""))}</td>'
                f'<td class="col-ip">{html.escape(getattr(r, "cdn_ips", ""))}</td>'
                f'<td class="col-device">{html.escape(getattr(r, "device", ""))}</td>'
                f'<td class="col-hash">{html.escape(getattr(r, "checksum", ""))}</td>'
                f'<td class="col-lcid">{html.escape(getattr(r, "lcid", ""))}</td>'
                f'<td class="col-source">{html.escape(getattr(r, "source_file", ""))}</td>'
                "</tr>"
            )
        lines.append("</tbody></table></div>")
    else:
        lines.append(
            '<div class="empty">No DV upload events found. Place “Dv Access logs … .csv” '
            "files next to selection.zip or inside the extracted return, then re-run.</div>"
        )
    lines.append("</div>")

    # Sync table
    lines.append('<div id="dv-sync" class="dv-pane" style="display:none">')
    if dv_syncs:
        lines.append('<div class="table-wrap"><table class="data" id="dv-sync-table">')
        lines.append(
            "  <thead><tr>"
            "<th class=\"col-date\" title=\"When the server recorded this access event (Coordinated Universal Time).\">Timestamp (UTC)</th>"
            "<th class=\"col-ip\" title=\"Public IP address of the device or user that made the request.\">User IP</th>"
            "<th class=\"col-ip\" title=\"Content Delivery Network addresses that forwarded the request. These are shared infrastructure, not the end user’s IP.\">CDN IPs</th>"
            "<th class=\"col-device\" title=\"Client/device identifier reported by the app (for example a phone model string).\">Device</th>"
            "<th title=\"Type of sync or access activity (no specific file upload). Shows device check-ins and conflict-resolve style events.\">Operation</th>"
            "<th class=\"col-lcid\" title=\"Line / account identifier (LCID) associated with this subscriber line in the Synchronoss / Verizon system.\">LCID</th>"
            "<th class=\"col-source\" title=\"Name of the DV Access Log CSV file this row came from.\">Source</th>"
            "</tr></thead><tbody>"
        )
        for r in dv_syncs:
            if getattr(r, "server_ts_dt", None):
                ts = r.server_ts_dt.strftime("%Y-%m-%d %H:%M:%S UTC")
            else:
                ts = getattr(r, "server_ts", "") or ""
            blob = " ".join(
                [
                    ts,
                    getattr(r, "user_ip", ""),
                    getattr(r, "cdn_ips", ""),
                    getattr(r, "device", ""),
                    getattr(r, "operation", ""),
                    getattr(r, "lcid", ""),
                    getattr(r, "source_file", ""),
                ]
            ).lower()
            lines.append(
                f'  <tr data-search="{html.escape(blob)}">'
                f'<td class="col-date">{html.escape(ts)}</td>'
                f'<td class="col-ip">{html.escape(getattr(r, "user_ip", ""))}</td>'
                f'<td class="col-ip">{html.escape(getattr(r, "cdn_ips", ""))}</td>'
                f'<td class="col-device">{html.escape(getattr(r, "device", ""))}</td>'
                f'<td>{html.escape(getattr(r, "operation", ""))}</td>'
                f'<td class="col-lcid">{html.escape(getattr(r, "lcid", ""))}</td>'
                f'<td class="col-source">{html.escape(getattr(r, "source_file", ""))}</td>'
                "</tr>"
            )
        lines.append("</tbody></table></div>")
    else:
        lines.append(
            '<div class="empty">No DV sync events found (or no DV log files were present).</div>'
        )
    lines.append("</div>")
    lines.append("</section>")

    # ===== DV UNIQUE IPs =====
    lines.append('<section id="dv-ips" class="section">')
    lines.append("  <h1>DV Unique IPs</h1>")
    lines.append(
        f'  <div class="subtitle">{n_dv_ips} unique user IP'
        f'{"s" if n_dv_ips != 1 else ""} from DV Access Log events '
        f"(uploads + sync). First / last seen are derived from event timestamps. "
        "CDN IPs are excluded — only the client (user) IP is listed.</div>"
    )
    lines.append(
        '  <div class="search-bar">'
        '<input id="dv-ip-search" class="search-input" '
        'placeholder="Filter by IP, device…">'
        f'<span class="filter-count" id="dv-ip-filter-count">'
        f'{n_dv_ips} of {n_dv_ips} IPs displayed</span>'
        "</div>"
    )
    if dv_ip_rows:
        lines.append('<div class="table-wrap"><table class="data" id="dv-ips-table">')
        lines.append(
            "  <thead><tr>"
            "<th class=\"col-ip sortable\" data-col=\"0\">User IP</th>"
            "<th class=\"col-date sortable\" data-col=\"1\">First Seen (UTC)</th>"
            "<th class=\"col-date sortable\" data-col=\"2\">Last Seen (UTC)</th>"
            "<th class=\"sortable\" data-col=\"3\">Events</th>"
            "<th class=\"col-device sortable\" data-col=\"4\">Device(s)</th>"
            "</tr></thead><tbody>"
        )
        for row in dv_ip_rows:
            ip = row["ip"]
            first_s = row["first_seen"]
            last_s = row["last_seen"]
            events = row["events"]
            devices = row["devices"]
            blob = " ".join([ip, first_s, last_s, events, devices]).lower()
            lines.append(
                f'  <tr data-search="{html.escape(blob)}">'
                f'<td class="col-ip mono">{html.escape(ip)}</td>'
                f'<td class="col-date">{html.escape(first_s)}</td>'
                f'<td class="col-date">{html.escape(last_s)}</td>'
                f'<td data-sort="{html.escape(events)}">{html.escape(events)}</td>'
                f'<td class="col-device">{html.escape(devices)}</td>'
                "</tr>"
            )
        lines.append("</tbody></table></div>")
    else:
        lines.append(
            '<div class="empty">No user IPs found in DV Access Logs '
            "(or no DV log files were present).</div>"
        )
    lines.append("</section>")

    lines.append("</div>")  # .main

    # Modal for MD5 export scope (Yes / No)
    lines.append(
        '<div class="modal-overlay" id="md5-export-modal" role="dialog" '
        'aria-modal="true" aria-labelledby="md5-export-title">'
    )
    lines.append('  <div class="modal-dialog">')
    lines.append('    <h2 id="md5-export-title">Export MD5 hashes</h2>')
    lines.append(
        "    <p>Would you like to include all media in the data "
        "(such as quarantined files, MMS attachments, etc.)?</p>"
    )
    lines.append('    <div class="modal-actions">')
    lines.append(
        '      <button type="button" class="btn secondary" id="md5-export-no">No</button>'
    )
    lines.append(
        '      <button type="button" class="btn" id="md5-export-yes">Yes</button>'
    )
    lines.append("    </div>")
    lines.append("  </div>")
    lines.append("</div>")

    # Floating metadata popover (shared by all media info buttons)
    lines.append(
        '<div class="meta-popover" id="meta-popover" role="dialog" aria-label="File metadata">'
        '<div class="meta-head"><span>Metadata</span>'
        '<button type="button" class="meta-close" id="meta-popover-close" aria-label="Close">×</button>'
        '</div><div class="meta-body" id="meta-popover-body"></div></div>'
    )

    # ----- Scripts -----
    lines.append("<script>")
    lines.append(f"const CHAT_SEARCH = {json.dumps(search_blobs, ensure_ascii=False)};")
    lines.append("""
(function(){
  // Section switching via nav / hash
  function showSection(id) {
    document.querySelectorAll('.section').forEach(function(s){
      s.classList.toggle('active', s.id === id);
    });
    document.querySelectorAll('.nav-links a').forEach(function(a){
      a.classList.toggle('active', a.getAttribute('data-section') === id);
    });
    try { history.replaceState(null, '', '#' + id); } catch (e) {}
    // Column widths only resolve correctly on visible tables. Hidden
    // sections (display:none) ignore fixed widths until shown — re-lock
    // after layout so Preview / Filename / MD5 are correct on first paint.
    requestAnimationFrame(function(){
      requestAnimationFrame(function(){
        var sec = document.getElementById(id);
        if (!sec || typeof lockTableColumns !== 'function') return;
        sec.querySelectorAll('table.data').forEach(function(t){
          lockTableColumns(t);
        });
      });
    });
  }
  document.querySelectorAll('.nav-links a').forEach(function(a){
    a.addEventListener('click', function(ev){
      ev.preventDefault();
      showSection(a.getAttribute('data-section'));
    });
  });
  var hash = (location.hash || '#home').replace(/^#/, '');
  if (!document.getElementById(hash)) hash = 'home';
  showSection(hash);

  // Conversation search
  var s = document.getElementById('search');
  var status = document.getElementById('search-status');
  if (s) {
    function applyChat(){
      var q = (s.value || '').toLowerCase().trim();
      var shown = 0;
      document.querySelectorAll('#chat-list .item').forEach(function(it){
        if (!q) { it.style.display = ''; shown++; return; }
        var i = parseInt(it.getAttribute('data-i'), 10);
        var titleMeta = (it.textContent || '').toLowerCase();
        var body = (CHAT_SEARCH[i] || '');
        var match = titleMeta.indexOf(q) !== -1 || body.indexOf(q) !== -1;
        it.style.display = match ? '' : 'none';
        if (match) shown++;
      });
      if (status) {
        if (q) {
          status.style.display = '';
          status.textContent = shown + ' chat' + (shown === 1 ? '' : 's') + ' match "' + s.value + '"';
        } else {
          status.style.display = 'none';
          status.textContent = '';
        }
      }
    }
    s.addEventListener('input', applyChat);
  }

  // Contact filter
  var cs = document.getElementById('contact-search');
  if (cs) {
    cs.addEventListener('input', function(){
      var q = (cs.value || '').toLowerCase().trim();
      var rows = document.querySelectorAll('#contacts-table tbody tr');
      var shown = 0;
      rows.forEach(function(tr){
        var blob = tr.getAttribute('data-search') || '';
        var match = !q || blob.indexOf(q) !== -1;
        tr.style.display = match ? '' : 'none';
        if (match) shown++;
      });
      var cnt = document.getElementById('contact-filter-count');
      if (cnt) cnt.textContent = shown + ' of ' + rows.length + ' contacts displayed';
    });
  }

  // Media page filters: text search + cascading column dropdowns
  // countId (optional): element id for "xx of xxx files displayed"
  function wireMediaFilters(tableSel, textId, selects, countId) {
    selects = selects || [];

    function rowMatches(tr, skipSelectId) {
      var q = '';
      var textEl = textId ? document.getElementById(textId) : null;
      if (textEl) q = (textEl.value || '').toLowerCase().trim();
      if (q) {
        var blob = tr.getAttribute('data-search') || '';
        if (blob.indexOf(q) === -1) return false;
      }
      for (var i = 0; i < selects.length; i++) {
        var s = selects[i];
        if (skipSelectId && s.id === skipSelectId) continue;
        var el = document.getElementById(s.id);
        var val = el ? el.value : '';
        if (val) {
          var attr = tr.getAttribute(s.attr) || '';
          if (attr !== val) return false;
        }
      }
      return true;
    }

    function refreshSelectOptions() {
      selects.forEach(function(s) {
        var el = document.getElementById(s.id);
        if (!el) return;
        var current = el.value;
        var seen = {};
        var values = [];
        document.querySelectorAll(tableSel + ' tbody tr').forEach(function(tr) {
          // Options limited by other active filters (and text search), not this select
          if (!rowMatches(tr, s.id)) return;
          var v = tr.getAttribute(s.attr) || '';
          if (!v || seen[v]) return;
          seen[v] = true;
          values.push(v);
        });
        values.sort(function(a, b) {
          return a.toLowerCase().localeCompare(b.toLowerCase());
        });
        el.innerHTML = '';
        var allOpt = document.createElement('option');
        allOpt.value = '';
        allOpt.textContent = 'All';
        el.appendChild(allOpt);
        values.forEach(function(v) {
          var opt = document.createElement('option');
          opt.value = v;
          opt.textContent = v;
          el.appendChild(opt);
        });
        // Keep selection if still valid; otherwise reset to All
        if (current && seen[current]) el.value = current;
        else el.value = '';
      });
    }

    function updateCount() {
      if (!countId) return;
      var el = document.getElementById(countId);
      if (!el) return;
      var rows = document.querySelectorAll(tableSel + ' tbody tr');
      var total = rows.length;
      var shown = 0;
      rows.forEach(function(tr) {
        if (tr.style.display !== 'none') shown++;
      });
      el.textContent = shown + ' of ' + total + ' files displayed';
    }

    function apply() {
      document.querySelectorAll(tableSel + ' tbody tr').forEach(function(tr) {
        tr.style.display = rowMatches(tr, null) ? '' : 'none';
      });
      refreshSelectOptions();
      // Re-apply visibility after option refresh (selection may have cleared)
      document.querySelectorAll(tableSel + ' tbody tr').forEach(function(tr) {
        tr.style.display = rowMatches(tr, null) ? '' : 'none';
      });
      updateCount();
    }

    if (textId) {
      var t = document.getElementById(textId);
      if (t) t.addEventListener('input', apply);
    }
    selects.forEach(function(s) {
      var el = document.getElementById(s.id);
      if (el) el.addEventListener('change', apply);
    });
    // Initial count (in case HTML was static)
    updateCount();
  }
  wireMediaFilters('#call-logs-table', 'call-search', [
    { id: 'call-filter-direction', attr: 'data-direction' },
    { id: 'call-filter-date', attr: 'data-date' }
  ], 'call-filter-count');
  wireMediaFilters('#attachments-table', 'att-search', [
    { id: 'att-filter-type', attr: 'data-type' },
    { id: 'att-filter-date', attr: 'data-date' }
  ], 'att-filter-count');
  wireMediaFilters('#unlinked-mms-table', 'unlinked-search', [
    { id: 'ul-filter-type', attr: 'data-type' },
    { id: 'ul-filter-date', attr: 'data-date' }
  ], 'ul-filter-count');
  wireMediaFilters('#vzmobile-table', 'vz-search', [
    { id: 'vz-filter-type', attr: 'data-type' },
    { id: 'vz-filter-date', attr: 'data-date' },
    { id: 'vz-filter-device', attr: 'data-device' }
  ], 'vz-filter-count');
  wireMediaFilters('#quarantine-table', 'quarantine-search', [
    { id: 'q-filter-type', attr: 'data-type' }
  ], 'q-filter-count');

  // Export MD5 list (.txt, one hash per line)
  function downloadBlob(filename, text, mime) {
    var blob = new Blob([text], { type: mime || 'text/plain' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function(){ URL.revokeObjectURL(url); }, 500);
  }
  function buildMd5Export(includeAll) {
    var data = (typeof MD5_EXPORT !== 'undefined') ? MD5_EXPORT : {};
    var hashes = [];
    var seen = {};
    function addList(arr) {
      (arr || []).forEach(function(h){
        h = (h || '').toString().trim().toLowerCase();
        if (h && !seen[h]) { seen[h] = true; hashes.push(h); }
      });
    }
    addList(data.vzmobile);
    if (includeAll) {
      addList(data.attachments);
      addList(data.quarantine);
      addList(data.unlinked_mms);
    }
    var phone = (data.target_phone || 'unknown').toString();
    var fname = includeAll
      ? ('Synchronoss_' + phone + '_AllMediaMD5.txt')
      : ('Synchronoss_' + phone + '_VZMediaMD5.txt');
    downloadBlob(fname, hashes.join('\\n') + (hashes.length ? '\\n' : ''), 'text/plain');
  }
  var md5Modal = document.getElementById('md5-export-modal');
  function openMd5Modal() {
    if (!md5Modal) return;
    md5Modal.classList.add('open');
  }
  function closeMd5Modal() {
    if (!md5Modal) return;
    md5Modal.classList.remove('open');
  }
  var btnTxt = document.getElementById('export-md5-txt');
  if (btnTxt) {
    btnTxt.addEventListener('click', function(){ openMd5Modal(); });
  }
  var btnYes = document.getElementById('md5-export-yes');
  if (btnYes) {
    btnYes.addEventListener('click', function(){
      closeMd5Modal();
      buildMd5Export(true);
    });
  }
  var btnNo = document.getElementById('md5-export-no');
  if (btnNo) {
    btnNo.addEventListener('click', function(){
      closeMd5Modal();
      buildMd5Export(false);
    });
  }
  if (md5Modal) {
    md5Modal.addEventListener('click', function(ev){
      if (ev.target === md5Modal) closeMd5Modal();
    });
  }

  // Column sort: click sortable headers
  document.querySelectorAll('table.data').forEach(function(table){
    table.querySelectorAll('th.sortable').forEach(function(th){
      th.addEventListener('click', function(ev){
        // Ignore clicks on the resize handle
        if (ev.target && ev.target.classList && ev.target.classList.contains('col-resizer')) return;
        var col = parseInt(th.getAttribute('data-col'), 10);
        if (isNaN(col)) return;
        var tbody = table.tBodies[0];
        if (!tbody) return;
        var rows = Array.prototype.slice.call(tbody.querySelectorAll('tr'));
        var dir = th.getAttribute('data-sort-dir') === 'asc' ? 'desc' : 'asc';
        table.querySelectorAll('th.sortable').forEach(function(h){
          h.removeAttribute('data-sort-dir');
          var ind = h.querySelector('.sort-ind');
          if (ind) ind.textContent = '';
        });
        th.setAttribute('data-sort-dir', dir);
        var ind = th.querySelector('.sort-ind');
        if (!ind) {
          ind = document.createElement('span');
          ind.className = 'sort-ind';
          th.appendChild(ind);
        }
        ind.textContent = dir === 'asc' ? ' ▲' : ' ▼';
        rows.sort(function(a, b){
          var ca = a.children[col];
          var cb = b.children[col];
          var va = (ca && ca.getAttribute('data-sort')) || (ca && ca.textContent) || '';
          var vb = (cb && cb.getAttribute('data-sort')) || (cb && cb.textContent) || '';
          va = String(va).toLowerCase();
          vb = String(vb).toLowerCase();
          if (va < vb) return dir === 'asc' ? -1 : 1;
          if (va > vb) return dir === 'asc' ? 1 : -1;
          return 0;
        });
        rows.forEach(function(r){ tbody.appendChild(r); });
      });
    });
  });

  // DV tabs
  document.querySelectorAll('.tab-btn').forEach(function(btn){
    btn.addEventListener('click', function(){
      document.querySelectorAll('.tab-btn').forEach(function(b){ b.classList.remove('active'); });
      btn.classList.add('active');
      var tab = btn.getAttribute('data-tab');
      document.querySelectorAll('.dv-pane').forEach(function(p){
        p.style.display = (p.id === tab) ? '' : 'none';
      });
    });
  });

  // DV filter (applies to both panes)
  var ds = document.getElementById('dv-search');
  if (ds) {
    ds.addEventListener('input', function(){
      var q = (ds.value || '').toLowerCase().trim();
      var rows = document.querySelectorAll('#dv-uploads-table tbody tr, #dv-sync-table tbody tr');
      var shown = 0;
      rows.forEach(function(tr){
        var blob = tr.getAttribute('data-search') || '';
        var match = !q || blob.indexOf(q) !== -1;
        tr.style.display = match ? '' : 'none';
        if (match) shown++;
      });
      var cnt = document.getElementById('dv-filter-count');
      if (cnt) cnt.textContent = shown + ' of ' + rows.length + ' events displayed';
    });
  }

  // DV unique IP filter
  var dip = document.getElementById('dv-ip-search');
  if (dip) {
    dip.addEventListener('input', function(){
      var q = (dip.value || '').toLowerCase().trim();
      var rows = document.querySelectorAll('#dv-ips-table tbody tr');
      var shown = 0;
      rows.forEach(function(tr){
        var blob = tr.getAttribute('data-search') || '';
        var match = !q || blob.indexOf(q) !== -1;
        tr.style.display = match ? '' : 'none';
        if (match) shown++;
      });
      var cnt = document.getElementById('dv-ip-filter-count');
      if (cnt) cnt.textContent = shown + ' of ' + rows.length + ' IPs displayed';
    });
  }

  // Fixed Preview column width (same on all media tables, including Quarantine).
  function previewWidthFor(table) {
    return 72;
  }
  function applyPreviewWidth(table) {
    if (!table) return;
    var pw = previewWidthFor(table);
    table.style.tableLayout = 'fixed';
    var col = table.querySelector('col.col-preview');
    if (col) {
      col.style.width = pw + 'px';
      col.style.minWidth = pw + 'px';
      col.style.maxWidth = pw + 'px';
    }
    table.querySelectorAll('th.col-preview, td.col-preview').forEach(function(el) {
      el.style.width = pw + 'px';
      el.style.minWidth = pw + 'px';
      el.style.maxWidth = pw + 'px';
      el.style.boxSizing = 'border-box';
    });
  }

  // Lock column widths to pixels while the table is visible.
  // Hidden sections (display:none) cannot resolve % / fixed widths correctly;
  // this is called from showSection after the section is shown.
  function lockTableColumns(table) {
    if (!table) return;
    var total = table.getBoundingClientRect().width || table.offsetWidth || 0;
    if (total < 10) return; // not visible / not laid out yet
    var pw = previewWidthFor(table);
    var cols = table.querySelectorAll('col');
    var ths = table.querySelectorAll('thead th');
    table.style.tableLayout = 'fixed';
    table.style.width = total + 'px';
    applyPreviewWidth(table);

    // Quarantine: explicit pixel split (preview + filename + md5 + type)
    if (table.id === 'quarantine-table' && cols.length >= 4) {
      var rest = Math.max(0, total - pw);
      var wFilename = Math.floor(rest * 0.52);
      var wMd5 = Math.floor(rest * 0.30);
      var wType = Math.max(40, rest - wFilename - wMd5);
      var qWidths = [pw, wFilename, wMd5, wType];
      qWidths.forEach(function(w, i) {
        if (cols[i]) {
          cols[i].style.width = w + 'px';
          cols[i].style.minWidth = (i === 0 ? w : 40) + 'px';
          if (i === 0) cols[i].style.maxWidth = w + 'px';
          else cols[i].style.maxWidth = '';
        }
        if (ths[i]) {
          ths[i].style.width = w + 'px';
          if (i === 0) {
            ths[i].style.minWidth = w + 'px';
            ths[i].style.maxWidth = w + 'px';
          } else {
            ths[i].style.minWidth = '40px';
            ths[i].style.maxWidth = '';
          }
        }
      });
      table.setAttribute('data-cols-locked', '1');
      return;
    }

    // Other media tables: lock Preview to 72px; lock remaining cols from layout
    ths.forEach(function(th, i) {
      if (th.classList.contains('col-preview')) {
        th.style.width = pw + 'px';
        th.style.minWidth = pw + 'px';
        th.style.maxWidth = pw + 'px';
        if (cols[i]) {
          cols[i].style.width = pw + 'px';
          cols[i].style.minWidth = pw + 'px';
          cols[i].style.maxWidth = pw + 'px';
        }
      } else {
        var w = th.getBoundingClientRect().width;
        if (w > 0) {
          th.style.width = w + 'px';
          th.style.minWidth = '40px';
          if (cols[i]) cols[i].style.width = w + 'px';
        }
      }
    });
    table.setAttribute('data-cols-locked', '1');
  }

  // Column resize on non-preview headers (Preview stays fixed)
  function enableColumnResize(table) {
    var ths = Array.prototype.slice.call(table.querySelectorAll('thead th'));
    if (!ths.length) return;
    var pw = previewWidthFor(table);

    // Initial lock only if already visible (e.g. deep-link to this section)
    lockTableColumns(table);

    ths.forEach(function(th, index) {
      if (th.classList.contains('col-preview')) return;
      var handle = document.createElement('div');
      handle.className = 'col-resizer';
      handle.title = 'Drag to resize column';
      th.appendChild(handle);

      handle.addEventListener('mousedown', function(ev) {
        ev.preventDefault();
        ev.stopPropagation();
        // Ensure locked to pixels before drag
        if (!table.getAttribute('data-cols-locked')) lockTableColumns(table);
        var widths = ths.map(function(h) {
          if (h.classList.contains('col-preview')) return pw;
          return h.getBoundingClientRect().width;
        });
        var startX = ev.pageX;
        var startW = widths[index];
        handle.classList.add('active');
        document.body.classList.add('col-resizing');

        function onMove(e) {
          var delta = e.pageX - startX;
          var newW = Math.max(40, startW + delta);
          th.style.width = newW + 'px';
          var col = table.querySelectorAll('col')[index];
          if (col) col.style.width = newW + 'px';
          applyPreviewWidth(table);
        }
        function onUp() {
          handle.classList.remove('active');
          document.body.classList.remove('col-resizing');
          document.removeEventListener('mousemove', onMove);
          document.removeEventListener('mouseup', onUp);
        }
        document.addEventListener('mousemove', onMove);
        document.addEventListener('mouseup', onUp);
      });
    });
  }
  document.querySelectorAll('table.data').forEach(function(table) {
    enableColumnResize(table);
  });

  // Metadata info-button popover
  (function() {
    var pop = document.getElementById('meta-popover');
    var body = document.getElementById('meta-popover-body');
    var closeBtn = document.getElementById('meta-popover-close');
    var activeBtn = null;

    function closePopover() {
      if (!pop) return;
      pop.classList.remove('open');
      if (activeBtn) activeBtn.classList.remove('active');
      activeBtn = null;
    }

    function openPopover(btn) {
      if (!pop || !body) return;
      var raw = btn.getAttribute('data-meta') || '{}';
      var data;
      try { data = JSON.parse(raw); } catch (e) { data = {}; }
      var keys = Object.keys(data);
      if (!keys.length) { closePopover(); return; }
      keys.sort(function(a, b) { return a.toLowerCase().localeCompare(b.toLowerCase()); });
      var html = '';
      keys.forEach(function(k) {
        html += '<div class="meta-row"><div class="meta-key">' + escapeHtml(k) +
                '</div><div class="meta-val">' + escapeHtml(String(data[k])) + '</div></div>';
      });
      body.innerHTML = html;
      if (activeBtn && activeBtn !== btn) activeBtn.classList.remove('active');
      activeBtn = btn;
      btn.classList.add('active');
      pop.classList.add('open');
      // Position near the button, keep on-screen
      var rect = btn.getBoundingClientRect();
      var pw = pop.offsetWidth || 300;
      var ph = pop.offsetHeight || 200;
      var left = rect.right + 8;
      if (left + pw > window.innerWidth - 12) left = rect.left - pw - 8;
      if (left < 8) left = 8;
      var top = rect.top;
      if (top + ph > window.innerHeight - 12) top = Math.max(8, window.innerHeight - ph - 12);
      pop.style.left = left + 'px';
      pop.style.top = top + 'px';
    }

    function escapeHtml(s) {
      return String(s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
    }

    document.addEventListener('click', function(ev) {
      var btn = ev.target.closest ? ev.target.closest('.info-btn') : null;
      if (btn) {
        ev.preventDefault();
        ev.stopPropagation();
        if (activeBtn === btn) closePopover();
        else openPopover(btn);
        return;
      }
      if (pop && pop.classList.contains('open') && !pop.contains(ev.target)) {
        closePopover();
      }
    });
    if (closeBtn) closeBtn.addEventListener('click', function(ev) {
      ev.preventDefault();
      closePopover();
    });
    document.addEventListener('keydown', function(ev) {
      if (ev.key === 'Escape') closePopover();
    });
    window.addEventListener('scroll', closePopover, true);
    window.addEventListener('resize', closePopover);
  })();
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

    call_log_rows: List[Dict[str, str]] = []
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

    write_index(out_root, index_entries, call_logs=call_log_rows)

    from openpyxl import Workbook

    call_log_path = out_root / "Call Log.xlsx"
    wb = Workbook()
    ws = wb.active
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

    print(f"\nDone. Open: {out_root / 'index.html'}")


if __name__ == "__main__":
    main()
