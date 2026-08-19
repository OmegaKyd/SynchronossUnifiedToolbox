#!/usr/bin/env python3
"""Convert Synchronoss contacts exports to Excel files."""

import sys, re, json, argparse
from pathlib import Path

# Try imports and give friendly guidance if packages aren't installed
try:
    import pandas as pd
except ImportError:  # pragma: no cover - user guidance path
    print("Missing dependency: pandas. Install with:\n  pip install --user pandas openpyxl")
    print("Exiting due to missing dependency.")
    sys.exit(1)

# ----- helpers to clean/parse “almost JSON” -----
def quick_clean(txt: str) -> str:
    txt = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', txt)  # remove control chars
    txt = re.sub(r',\s*([}\]])', r'\1', txt)                # trailing commas
    return txt.strip()

def coerce_to_json(txt: str) -> str:
    s = txt.strip()
    if s.startswith('{"contacts"') and not s.endswith('}'):
        s = s.rstrip() + "}}"
    if s.startswith('[') and not s.endswith(']'):
        s = s.rstrip().rstrip(',') + ']'
    return s

def parse_contacts(raw_text: str):
    txt = coerce_to_json(quick_clean(raw_text))
    try:
        data = json.loads(txt)
    except json.JSONDecodeError as e:
        raise ValueError(f"JSON parse error at char {e.pos}: {e.msg}")
    if isinstance(data, dict):
        contacts = (data.get("contacts") or {}).get("contact") or data.get("contact")
    else:
        contacts = data
    if not isinstance(contacts, list):
        raise ValueError("Couldn’t find a contact list in the file.")
    return contacts

def normalize_lists_to_strings(df):
    def list_to_str(v):
        if isinstance(v, list):
            if v and all(isinstance(x, dict) for x in v):
                def one(d): return ", ".join(f"{k}={d.get(k,'')}" for k in d.keys())
                return "; ".join(one(x) for x in v)
            return "; ".join(map(str, v))
        return v
    for col in df.columns:
        df[col] = df[col].apply(list_to_str)
    return df

def extract_phone_columns(df):
    if 'tel' in df.columns:
        numbers, types, prefs = [], [], []
        for val in df['tel'].fillna(""):
            num_list, type_list, pref_list = [], [], []
            for part in str(val).split(';'):
                for kv in part.split(','):
                    k, _, v = kv.partition('=')
                    k, v = k.strip(), v.strip()
                    if k == 'number' and v: num_list.append(v)
                    elif k == 'type' and v: type_list.append(v)
                    elif k == 'preference' and v: pref_list.append(v)
            numbers.append("; ".join(num_list) or None)
            types.append("; ".join(type_list) or None)
            prefs.append("; ".join(pref_list) or None)
        df['phone_numbers'] = numbers
        df['phone_types'] = types
        df['phone_preferences'] = prefs
    return df

def build_dataframe(contacts):
    import pandas as pd
    df = pd.json_normalize(contacts, sep='.')
    df = normalize_lists_to_strings(df)
    df = extract_phone_columns(df)
    preferred = [
        'firstname','lastname','phone_numbers','phone_types','phone_preferences',
        'source','created','deleted','itemguid','incaseofemergency','favorite'
    ]
    cols = [c for c in preferred if c in df.columns] + [c for c in df.columns if c not in preferred]
    return df[cols]

def convert_contacts(input_file: str, output_file: str) -> int:
    """Convert a Synchronoss contacts dump to Excel."""
    in_path = Path(input_file)
    if not in_path.exists():
        raise FileNotFoundError(f"File not found: {in_path}")

    raw = in_path.read_text(encoding="utf-8", errors="ignore")
    contacts = parse_contacts(raw)
    df = build_dataframe(contacts)
    df.to_excel(output_file, index=False)
    return len(df)


def main():  # pragma: no cover - CLI convenience wrapper
    parser = argparse.ArgumentParser(description="Convert Synchronoss contacts dump to Excel.")
    parser.add_argument("--input", required=True, help="Path to contacts.txt")
    parser.add_argument("--output", required=True, help="Path to output .xlsx file")
    args = parser.parse_args()

    try:
        rows = convert_contacts(args.input, args.output)
        print(f"Wrote {rows} rows to {args.output}")
    except Exception as e:  # broad but intentional for user guidance
        print(f"Error: {e}")
        sys.exit(1)
    print("Done.")


def launch_gui():  # pragma: no cover - GUI wrapper
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.title("Contacts to Excel")

    in_var = tk.StringVar()
    out_var = tk.StringVar()
    status = tk.StringVar()

    def browse_in():
        path = filedialog.askopenfilename(
            title="Select contacts.txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        )
        if path:
            in_var.set(path)

    def browse_out():
        path = filedialog.asksaveasfilename(
            title="Save Excel file",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            out_var.set(path)

    def do_convert():
        status.set("Converting...")
        root.update_idletasks()
        try:
            rows = convert_contacts(in_var.get(), out_var.get())
            status.set(f"Wrote {rows} rows to {out_var.get()}")
        except Exception as e:
            status.set(f"Error: {e}")

    tk.Label(root, text="Input file:").grid(row=0, column=0, sticky="e")
    tk.Entry(root, textvariable=in_var, width=50).grid(row=0, column=1, padx=5)
    tk.Button(root, text="Browse", command=browse_in).grid(row=0, column=2)

    tk.Label(root, text="Output file:").grid(row=1, column=0, sticky="e")
    tk.Entry(root, textvariable=out_var, width=50).grid(row=1, column=1, padx=5)
    tk.Button(root, text="Save As", command=browse_out).grid(row=1, column=2)

    tk.Button(root, text="Convert", command=do_convert).grid(
        row=2, column=1, pady=10
    )
    tk.Label(root, textvariable=status, fg="blue").grid(row=3, column=0, columnspan=3)

    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--gui":
        launch_gui()
    else:
        main()
