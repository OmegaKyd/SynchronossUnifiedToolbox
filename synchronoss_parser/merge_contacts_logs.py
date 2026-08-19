#!/usr/bin/env python3
"""Annotate call logs with contact names.

Reads a call log CSV (`Call Log/call_log.csv`) and an Excel contacts file, then writes a new CSV
with added ``caller_name`` and ``recipient_name`` columns based on phone
number matches.

Usage:
  merge-contacts-logs --call-log 'Call Log/call_log.csv' --contacts-xlsx contacts.xlsx
"""

import argparse
import sys
from pathlib import Path

# Try imports and give friendly guidance if packages aren't installed
try:  # pragma: no cover - import guidance
    import pandas as pd
except ImportError:  # pragma: no cover - user guidance path
    print("Missing dependency: pandas. Install with:\n  pip install --user pandas openpyxl")
    print("Exiting due to missing dependency.")
    sys.exit(1)

from .render_transcripts import build_contact_lookup


def merge_call_log(call_log_csv: str, contacts_xlsx: str, output_csv: str) -> int:
    """Merge call log with contacts and write a new CSV."""
    lookup = build_contact_lookup(contacts_xlsx)
    df = pd.read_csv(call_log_csv)
    if "caller" in df.columns:
        df["caller_name"] = df["caller"].apply(lookup)
    else:
        df["caller_name"] = ""
    if "recipient" in df.columns:
        df["recipient_name"] = df["recipient"].apply(lookup)
    else:
        df["recipient_name"] = ""
    df.to_csv(output_csv, index=False)
    return len(df)


def main() -> None:  # pragma: no cover - CLI convenience wrapper
    parser = argparse.ArgumentParser(
        description="Merge call log CSV with contacts to annotate names."
    )
    parser.add_argument("--call-log", required=True, help="Path to call_log.csv")
    parser.add_argument(
        "--contacts-xlsx", required=True, help="Path to contacts Excel file"
    )
    parser.add_argument(
        "--output", help="Path for output CSV (default: call_log_named.csv)"
    )
    args = parser.parse_args()

    out_path = (
        Path(args.output)
        if args.output
        else Path(args.call_log).with_name(Path(args.call_log).stem + "_named.csv")
    )

    try:
        rows = merge_call_log(args.call_log, args.contacts_xlsx, str(out_path))
        print(f"Wrote {rows} rows to {out_path}")
    except Exception as e:  # broad but intentional for user guidance
        print(f"Error: {e}")
        sys.exit(1)
    print("Done.")


if __name__ == "__main__":
    main()
