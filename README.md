# Synchronoss Unified Toolbox

**Version 3.0.0**

Organize Verizon / Synchronoss mobile backup returns into reviewable case data:
contacts, chat transcripts with attachments, media, call logs, DV access logs,
unique IPs, and more.

Built using portions of original code by Rempstrom and Alexis Brignoni and unified into a single automatic pipeline by Shane Hardie.

---

## What it does

You provide:

1. The original **`selection.zip`** from Synchronoss  
2. The **GPG passphrase** from the `#secure` email  
3. A **case folder** to write into  
4. Optionally, the account **Owner name**

The toolbox then runs **9 stages**:

1. Extract `selection.zip`
2. Discover encrypted payloads (main / contacts / quarantine)
3. Look for DV Access Log CSVs (optional)
4. Decrypt the `.gpg` payloads
5. Extract the main content archive
6. Convert contacts → Excel
7. Recover quarantined media
8. Collect media, attachments & unlinked MMS
9. Render chat transcripts and write **`index.html`**

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

1. Put this project folder somewhere short (e.g. `C:\Tools\Synchronoss-Unified-Toolbox`).
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
    index.html                      ← case dashboard (nav sections below)
    conversations/                  ← one folder per chat
      <Contact>/conversation.html
    contacts.xlsx
    Call Log.xlsx
    Unlinked MMS.xlsx / Unlinked MMS/
    DV Access Logs.xlsx             ← when DV CSVs are present
    Compiled Media/                 ← if present in the return
    Compiled Attachments/
    Compiled Quarantine Files/
    processing_summary.txt
```

---

## Index (browser)

Open `parsed_output/index.html` offline. Left navigation includes:

- **Home** — summary and counts  
- **Call Logs** — text + Direction/Date filters  
- **Contacts** — filterable table  
- **Conversations** — search titles and message bodies  
- **Compiled Attachments** / **Unlinked MMS** / **VZMOBILE** / **Quarantine** — previews, filters; ⓘ metadata where available  
- **DV Access Logs** — uploads & sync events (header tooltips for field meanings)  
- **DV Unique IPs** — user IPs with first/last seen  

Nav items show record counts. Filter bars show live **xx of xxx displayed**.

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

Individual modules remain available for advanced use. The **GUI is the supported path** for full case processing.

```bat
python -m synchronoss_parser.toolbox_gui
```

---

## Documentation

| File | Purpose |
|------|---------|
| `User Guide v3.0.0.txt` | Full investigator guide |
| `Quick Start Guide.txt` | Short Windows setup |
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
