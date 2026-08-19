# Synchronoss Unified Toolbox

**Version 2.2.0**

Organize Verizon / Synchronoss mobile backup returns into reviewable case data:
contacts, chat transcripts with attachments, media, call logs, and more.

Built on the original Synchronoss Toolbox by [Rempstrom](https://github.com)
and unified into a single automatic pipeline by **Ω OmegaKyd Ω**.

---

## What it does

You provide:

1. The original **`selection.zip`** from Synchronoss  
2. The **GPG passphrase** from the `#secure` email  
3. A **case folder** to write into  
4. Optionally, the account **Owner name**

The toolbox then:

- Extracts `selection.zip`
- Decrypts the standard payloads (`*.zip.gpg`, `*-contacts.txt.gpg`, `*-quarantined.zip.gpg`)
- Converts contacts → Excel
- Recovers quarantined media
- Collects media and message attachments (original filenames preserved)
- Renders **each chat** into its own folder under `conversations/`
- Writes **`index.html`** (searchable across message content)
- Builds a named **Call Log** and a **processing summary**

Owner phone number is **auto-detected** from encrypted filenames.

---

## Requirements (Windows)

| Requirement | Notes |
|-------------|--------|
| **Python 3.8+** | [python.org](https://www.python.org/downloads/) — check “Add Python to PATH” |
| **Packages** | `pip install -r requirements.txt` (Pillow, openpyxl, pandas) |
| **GPG4Win** | [gpg4win.org](https://www.gpg4win.org) or `winget install --id GnuPG.Gpg4win -e --source winget` |

The GUI can offer the winget install and restart the app if `gpg` is missing.

---

## Quick start

1. Put this project folder somewhere short (e.g. `C:\Tools\Synchronoss-Toolbox-Unified`).
2. Install dependencies:

```bat
pip install -r requirements.txt
```

3. Double-click:

```
Run_Toolbox_GUI.bat
```

   (Uses `pyw` / `pythonw` so no command prompt stays open.)

4. In the GUI:

   - **Selection.zip** — original (or copied) file from Synchronoss, *not* unzipped content  
   - **Case folder** — where results should go  
   - **GPG Password** — passphrase from the `#secure` email  
   - **Owner name (optional)** — e.g. `Jane Doe`  
   - Click **Run**

5. When finished, use **Open output folder** (or open `parsed_output\index.html`).

---

## Output layout

```
CaseFolder/                         ← path you select
  original_working/                 ← intermediate decryption / extraction
  parsed_output/                    ← final case data
    index.html                      ← open this first
    conversations/                  ← one folder per chat
      <Contact>/conversation.html
    contacts.xlsx
    Call Log.xlsx
    Compiled Media/                 ← if present in the return
    Compiled Attachments/
    Compiled Quarantine Files/
    processing_summary.txt
```

---

## Transcripts (browser)

- **Index search** — matches chat titles *and* message bodies  
- **Attachments only** — filter messages that have attachments  
- **Print** — chat header on page 1; footer with chat name (left) and page numbers (right). Turn off browser “Headers and footers” to hide the file path.  
- Contact labels show **Name (number)** when known  

---

## Command line (optional)

```bat
python -m synchronoss_parser.full_pipeline ^
  --selection path\to\selection.zip ^
  --output path\to\CaseFolder ^
  --passphrase "YOUR_GPG_PASSPHRASE" ^
  --target-name "Jane Doe"
```

`--output` is the **case root** (creates `original_working/` and `parsed_output/` inside it).

Individual modules remain available for advanced use (`collect_media`, `render_transcripts`, etc.). The **GUI is the supported path** for full case processing.

```bat
python -m synchronoss_parser.toolbox_gui
```

---

## Documentation

| File | Purpose |
|------|---------|
| `User Guide v2.2.0.txt` | Full investigator guide |
| `Disclaimer.txt` | Use and verification notice |
| `LICENSE.txt` | MIT License (© 2026 OmegaKyd) |
| `../CHANGELOG.md` | Version history (project root, outside this folder) |

---

## Building a Windows .exe (optional)

```bat
pip install pyinstaller
python -m synchronoss_parser.build_exe
```

Produces a windowed executable under `dist/` (no console). Requires GPG4Win still installed on the machine that *runs* decryption.

---

## Testing (developers)

```bat
pytest
```

---

## Disclaimer

This software helps organize and review Synchronoss returns. It does **not** replace investigative judgment. **Trust but verify** against the original data.

See `Disclaimer.txt` and `LICENSE.txt`.
