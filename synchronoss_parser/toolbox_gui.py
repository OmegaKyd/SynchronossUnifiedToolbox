#!/usr/bin/env python3
"""Unified Tkinter GUI exposing multiple Synchronoss tools.

This script bundles the existing utilities into a single window with
tabbed navigation so non-technical users can run them more easily. Long
running operations are executed in background threads while an indeterminate
``ttk.Progressbar`` is shown to keep the interface responsive.

Tabs provided:

* **Collect Media** – wraps ``collect_media.collect_media`` and
  ``collect_media.write_excel``.
* **Contacts to Excel** – wraps ``contacts_to_excel.convert_contacts``.
* **Render Transcripts** – wraps ``render_transcripts.main`` and includes
  an entry for the target phone number.
* **Collect Attachments** – wraps ``collect_attachments.collect_attachments``
  and ``collect_attachments.write_excel``.
* **Collect Quarantined Files** – wraps
  ``collect_quarantined_files.collect_quarantined_files``.

The script can be packaged as a standalone executable with PyInstaller:

```
pyinstaller --onefile synchronoss_parser/toolbox_gui.py
```
"""

from __future__ import annotations

import os
import subprocess
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# ---------------------------------------------------------------------------
# Make this script runnable from any location (double-click, inside the
# package folder, or as ``python -m synchronoss_parser.toolbox_gui``).
# ---------------------------------------------------------------------------
_THIS_FILE = Path(__file__).resolve()
_PKG_DIR = _THIS_FILE.parent                 # .../synchronoss_parser
_PROJECT_ROOT = _PKG_DIR.parent              # parent of the package

# When the user runs the .py file directly (or double-clicks it), the package
# root is not on sys.path.  Insert it so ``import synchronoss_parser`` works.
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

# Local modules (absolute package imports – now safe after the path fix)
from synchronoss_parser import collect_media as cm
from synchronoss_parser import collect_attachments as ca
from synchronoss_parser import collect_quarantined_files
from synchronoss_parser.contacts_to_excel import convert_contacts
from synchronoss_parser import render_transcripts as rt
from synchronoss_parser import decrypt_unzip
from synchronoss_parser import full_pipeline
from synchronoss_parser.utils import normalize_phone_number, gpg_available

APP_VERSION = "2.2.0"


def open_in_file_manager(path: Path | str) -> None:
    """Open *path* in the OS file manager (Explorer / Finder / xdg-open)."""
    target = Path(path).expanduser().resolve()
    if not target.exists():
        messagebox.showerror(
            "Folder not found",
            f"The path does not exist:\n{target}",
        )
        return
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(target))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(target)])
        else:
            subprocess.Popen(["xdg-open", str(target)])
    except Exception as e:
        messagebox.showerror(
            "Could not open folder",
            f"Failed to open:\n{target}\n\n{e}",
        )


def restart_application() -> None:
    """Relaunch this application and exit the current process."""
    try:
        if getattr(sys, "frozen", False):
            # Running as a PyInstaller (or similar) executable
            cmd = [sys.executable, *sys.argv[1:]]
        else:
            # Running as a .py script
            cmd = [sys.executable, str(Path(sys.argv[0]).resolve()), *sys.argv[1:]]
        subprocess.Popen(cmd, cwd=str(Path.cwd()))
    except Exception as e:
        messagebox.showerror(
            "Restart failed",
            f"Could not restart automatically:\n{e}\n\n"
            "Please close and re-open the application manually.",
        )
        return
    # Close the current instance
    try:
        root = tk._default_root  # type: ignore[attr-defined]
        if root is not None:
            root.destroy()
    except Exception:
        pass
    sys.exit(0)


def offer_gpg4win_install(parent: tk.Misc | None = None) -> None:
    """If gpg is missing, ask the user whether to install GPG4Win via winget."""
    install = messagebox.askyesno(
        "GPG not found",
        "gpg was not found on this system.\n\n"
        "GPG4Win is required to decrypt Synchronoss archives.\n\n"
        "Would you like to install it now with winget?\n\n"
        "Command that will be run:\n"
        "  winget install --id GnuPG.Gpg4win -e --source winget\n\n"
        "Administrator privileges may be required.",
        parent=parent,
    )
    if not install:
        return

    # Launch winget in a new console so the user can see progress / UAC prompts
    cmd = "winget install --id GnuPG.Gpg4win -e --source winget"
    try:
        if sys.platform.startswith("win"):
            subprocess.Popen(
                [
                    "powershell.exe",
                    "-NoExit",
                    "-Command",
                    cmd,
                ],
                creationflags=getattr(subprocess, "CREATE_NEW_CONSOLE", 0),
            )
            # Block until the user confirms install is done, then offer restart
            messagebox.showinfo(
                "Installing GPG4Win",
                "A PowerShell window has been opened to install GPG4Win.\n\n"
                "1. Approve any User Account Control (UAC) prompts if shown.\n"
                "2. Wait until the install finishes.\n"
                "3. Click OK below when you are ready to restart this application.",
                parent=parent,
            )
            if messagebox.askyesno(
                "Restart application",
                "Restart this application now so it can find gpg?",
                parent=parent,
            ):
                restart_application()
        else:
            messagebox.showinfo(
                "Not available",
                "Automatic install via winget is only supported on Windows.\n\n"
                "Please install GnuPG manually, then restart this application.",
                parent=parent,
            )
    except FileNotFoundError:
        messagebox.showerror(
            "winget not found",
            "winget could not be started.\n\n"
            "Install GPG4Win manually from:\n"
            "  https://www.gpg4win.org\n\n"
            "Then restart this application.",
            parent=parent,
        )
    except Exception as e:
        messagebox.showerror(
            "Install failed",
            f"Could not start the installer:\n{e}\n\n"
            "Install GPG4Win manually from https://www.gpg4win.org",
            parent=parent,
        )


# ---------------------------------------------------------------------------
# Tab builders
# ---------------------------------------------------------------------------


def build_full_pipeline_ui(parent: tk.Misc) -> None:
    """Build the unified pipeline UI (selection.zip → organized case folder)."""

    frame = ttk.Frame(parent, padding=12)
    frame.pack(fill="both", expand=True)
    # Column 1 (entry fields) expands so all inputs share the same width
    frame.columnconfigure(1, weight=1)

    selection_var = tk.StringVar()
    output_var = tk.StringVar()
    password_var = tk.StringVar()
    target_name_var = tk.StringVar()
    status_var = tk.StringVar()
    last_result_dir: list[Path | None] = [None]  # mutable holder for success path

    # Green progress style for the completed state
    style = ttk.Style()
    try:
        style.configure(
            "Green.Horizontal.TProgressbar",
            troughcolor="#e5e7eb",
            background="#16a34a",
            thickness=18,
        )
    except Exception:
        pass

    # Progress bar + optional centered "✓ Complete" overlay
    progress_frame = ttk.Frame(frame)
    progress = ttk.Progressbar(
        progress_frame, mode="indeterminate", style="Horizontal.TProgressbar"
    )
    progress.pack(fill="x", expand=True)
    done_label = tk.Label(
        progress_frame,
        text="✓  Complete",
        fg="#ffffff",
        bg="#16a34a",
        font=("Segoe UI", 10, "bold"),
        bd=0,
    )

    def reset_progress() -> None:
        """Return the bar to its idle / running indeterminate state."""
        done_label.place_forget()
        progress.stop()
        progress.configure(mode="indeterminate", style="Horizontal.TProgressbar")
        try:
            progress["value"] = 0
        except Exception:
            pass
        open_btn.grid_remove()
        last_result_dir[0] = None

    def show_complete() -> None:
        """Fill the bar green and show a centered checkmark."""
        progress.stop()
        progress.configure(mode="determinate", style="Green.Horizontal.TProgressbar")
        progress["maximum"] = 100
        progress["value"] = 100
        done_label.place(relx=0.5, rely=0.5, anchor="center")

    def browse_selection() -> None:
        path = filedialog.askopenfilename(
            title="Select Selection.zip",
            filetypes=[("Zip archives", "*.zip"), ("All files", "*.*")],
        )
        if path:
            selection_var.set(path)

    def browse_output() -> None:
        path = filedialog.askdirectory(initialdir=output_var.get() or ".")
        if path:
            output_var.set(path)

    def run() -> None:
        if not gpg_available():
            status_var.set(
                "gpg not found. Install GPG4Win, then restart this application."
            )
            offer_gpg4win_install(parent=frame.winfo_toplevel())
            return
        if not selection_var.get() or not password_var.get() or not output_var.get():
            status_var.set("Please fill in Selection.zip, Password and Case folder.")
            return

        reset_progress()
        progress.start()
        status_var.set("Starting full pipeline …")

        def task() -> None:
            def cb(msg: str) -> None:
                frame.after(0, lambda m=msg: status_var.set(m))

            success = False
            result_dir: Path | None = None
            try:
                result_dir = full_pipeline.run_pipeline(
                    selection_zip=selection_var.get(),
                    passphrase=password_var.get(),
                    output_dir=output_var.get(),
                    target_name=target_name_var.get().strip(),
                    keep_work_dir=True,
                    progress_callback=cb,
                )
                msg = (
                    f"Pipeline finished successfully.\n"
                    f"Case root : {output_var.get()}\n"
                    f"  original_working/ – intermediate files\n"
                    f"  parsed_output/    – final case data\n"
                    f"Open: {result_dir / 'index.html'}"
                )
                success = True
            except Exception as e:
                msg = f"Error: {e}"

            def finish() -> None:
                status_var.set(msg)
                if success and result_dir is not None:
                    show_complete()
                    last_result_dir[0] = Path(result_dir)
                    open_btn.grid()
                else:
                    reset_progress()

            frame.after(0, finish)

        threading.Thread(target=task, daemon=True).start()

    # ----- Layout (aligned entry column) -----
    ttk.Label(frame, text="Selection.zip:").grid(row=0, column=0, sticky="e", padx=(5, 8), pady=5)
    ttk.Entry(frame, textvariable=selection_var).grid(row=0, column=1, sticky="ew", padx=5, pady=5)
    ttk.Button(frame, text="Browse", command=browse_selection).grid(row=0, column=2, padx=5, pady=5)
    ttk.Label(
        frame,
        text='Use the original (or copied) "selection.zip" file provided by Synchronoss - not any of the unzipped files',
        foreground="gray",
    ).grid(row=1, column=1, sticky="w", padx=5, pady=(0, 6))

    ttk.Label(frame, text="Case folder:").grid(row=2, column=0, sticky="e", padx=(5, 8), pady=5)
    ttk.Entry(frame, textvariable=output_var).grid(row=2, column=1, sticky="ew", padx=5, pady=5)
    ttk.Button(frame, text="Browse", command=browse_output).grid(row=2, column=2, padx=5, pady=5)

    ttk.Label(frame, text="GPG Password:").grid(row=3, column=0, sticky="e", padx=(5, 8), pady=5)
    ttk.Entry(frame, textvariable=password_var, show="*").grid(
        row=3, column=1, sticky="ew", padx=5, pady=5
    )

    ttk.Label(frame, text="Owner name (optional):").grid(
        row=4, column=0, sticky="e", padx=(5, 8), pady=5
    )
    ttk.Entry(frame, textvariable=target_name_var).grid(
        row=4, column=1, sticky="ew", padx=5, pady=5
    )
    ttk.Label(
        frame,
        text="Account holder’s name – shown in transcripts like other contacts. Phone # is auto-detected.",
        foreground="gray",
    ).grid(row=5, column=1, sticky="w", padx=5, pady=(0, 4))

    ttk.Button(frame, text="Run", command=run).grid(row=6, column=1, pady=12)

    progress_frame.grid(row=7, column=0, columnspan=3, sticky="ew", padx=5, pady=4)
    ttk.Label(frame, textvariable=status_var, wraplength=560, justify="left").grid(
        row=8, column=0, columnspan=3, padx=5, pady=8, sticky="w"
    )

    def open_output_folder() -> None:
        path = last_result_dir[0]
        if path is None:
            # Fall back to the case folder the user selected
            case = output_var.get().strip()
            if case:
                path = Path(case) / "parsed_output"
            else:
                messagebox.showinfo(
                    "No output yet",
                    "Run the Full Pipeline first. The output folder button "
                    "will appear when it finishes successfully.",
                    parent=frame.winfo_toplevel(),
                )
                return
        open_in_file_manager(path)

    open_btn = ttk.Button(
        frame, text="Open output folder", command=open_output_folder
    )
    # Placed but hidden until a successful run
    open_btn.grid(row=9, column=1, pady=(0, 8))
    open_btn.grid_remove()

    # Help text
    help_txt = (
        "Select a case folder. The pipeline creates two subfolders inside it:\n"
        "  original_working/ – intermediate decryption & extraction artifacts\n"
        "  parsed_output/    – final case data (index.html, conversations, contacts, etc.)\n"
        "\n"
        "Workflow:\n"
        "  1. Extract selection.zip\n"
        "  2. Decrypt the three standard .gpg payloads (main / contacts / quarantine)\n"
        "  3. Convert contacts → Excel\n"
        "  4. Recover quarantined media\n"
        "  5. Collect media & attachments\n"
        "  6. Render each conversation under parsed_output/conversations/\n"
        "  7. Write parsed_output/index.html (open this first)"
    )
    ttk.Label(frame, text=help_txt, foreground="gray", justify="left").grid(
        row=10, column=0, columnspan=3, padx=5, pady=5, sticky="w"
    )


def build_collect_media_tab(nb: ttk.Notebook) -> None:
    """Add the Collect Media UI to ``nb``."""

    frame = ttk.Frame(nb)
    nb.add(frame, text="Collect Media")

    in_var = tk.StringVar()
    out_var = tk.StringVar()
    status_var = tk.StringVar()
    contacts_var = tk.StringVar()
    progress = ttk.Progressbar(frame, mode="indeterminate")

    def browse_in() -> None:
        path = filedialog.askdirectory(initialdir=in_var.get() or ".")
        if path:
            in_var.set(path)

    def browse_out() -> None:
        path = filedialog.askdirectory(initialdir=out_var.get() or ".")
        if path:
            out_var.set(path)

    def browse_contacts() -> None:
        path = filedialog.askopenfilename(
            title="Select contacts Excel file",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            contacts_var.set(path)

    def run() -> None:
        progress.start()

        def task() -> None:
            root_path = Path(in_var.get()).expanduser()
            compiled_path = Path(out_var.get()).expanduser()

            if not root_path.exists():
                frame.after(
                    0,
                    lambda: [
                        status_var.set(f"Input folder '{root_path}' does not exist."),
                        progress.stop(),
                    ],
                )
                return

            try:
                compiled_path.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                frame.after(
                    0,
                    lambda: [
                        status_var.set(
                            f"Could not create output folder '{compiled_path}': {e}"
                        ),
                        progress.stop(),
                    ],
                )
                return

            logfile = compiled_path / "compiled_media_log" / "compiled_media_log.xlsx"
            try:
                records, exif_keys = cm.collect_media(root_path, compiled_path)
                cm.write_excel(records, exif_keys, logfile)
                msg = (
                    f"Copied {len(records)} files from '{root_path}' to '{compiled_path}' and "
                    f"logged to '{logfile}'."
                )
            except Exception as e:  # pragma: no cover - user feedback
                msg = f"Error: {e}"

            frame.after(0, lambda: [status_var.set(msg), progress.stop()])

        threading.Thread(target=task, daemon=True).start()

    ttk.Label(frame, text="'VZMOBILE' Folder Path:").grid(
        row=0, column=0, sticky="e", padx=5, pady=5
    )
    ttk.Entry(frame, textvariable=in_var, width=50).grid(
        row=0, column=1, padx=5, pady=5
    )
    ttk.Button(frame, text="Browse", command=browse_in).grid(
        row=0, column=2, padx=5, pady=5
    )

    ttk.Label(frame, text="Output folder:").grid(
        row=1, column=0, sticky="e", padx=5, pady=5
    )
    ttk.Entry(frame, textvariable=out_var, width=50).grid(
        row=1, column=1, padx=5, pady=5
    )
    ttk.Button(frame, text="Browse", command=browse_out).grid(
        row=1, column=2, padx=5, pady=5
    )

    ttk.Label(frame, text="Contacts file:").grid(
        row=2, column=0, sticky="e", padx=5, pady=5
    )
    ttk.Entry(frame, textvariable=contacts_var, width=50).grid(
        row=2, column=1, padx=5, pady=5
    )
    ttk.Button(frame, text="Browse", command=browse_contacts).grid(
        row=2, column=2, padx=5, pady=5
    )

    ttk.Button(frame, text="Run", command=run).grid(row=3, column=1, pady=10)

    progress.grid(row=4, column=0, columnspan=3, sticky="ew", padx=5)

    ttk.Label(frame, textvariable=status_var, wraplength=400, justify="left").grid(
        row=5, column=0, columnspan=3, padx=5, pady=5
    )


def build_contacts_tab(nb: ttk.Notebook) -> None:
    """Add the Contacts to Excel UI to ``nb``."""

    frame = ttk.Frame(nb)
    nb.add(frame, text="Contacts to Excel")

    in_var = tk.StringVar()
    out_var = tk.StringVar()
    status_var = tk.StringVar()
    progress = ttk.Progressbar(frame, mode="indeterminate")

    def browse_in() -> None:
        path = filedialog.askopenfilename(
            title="Select contacts.txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        )
        if path:
            in_var.set(path)

    def browse_out() -> None:
        path = filedialog.asksaveasfilename(
            title="Save Excel file",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            out_var.set(path)

    def convert() -> None:
        progress.start()
        status_var.set("Converting...")

        def task() -> None:
            try:
                rows = convert_contacts(in_var.get(), out_var.get())
                msg = f"Wrote {rows} rows to '{out_var.get()}'"
            except Exception as e:  # pragma: no cover - user feedback
                msg = f"Error: {e}"

            frame.after(0, lambda: [status_var.set(msg), progress.stop()])

        threading.Thread(target=task, daemon=True).start()

    ttk.Label(frame, text="'contacts.txt' File Path:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
    ttk.Entry(frame, textvariable=in_var, width=50).grid(row=0, column=1, padx=5, pady=5)
    ttk.Button(frame, text="Browse", command=browse_in).grid(row=0, column=2, padx=5, pady=5)

    ttk.Label(frame, text="Output file:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
    ttk.Entry(frame, textvariable=out_var, width=50).grid(row=1, column=1, padx=5, pady=5)
    ttk.Button(frame, text="Save As", command=browse_out).grid(row=1, column=2, padx=5, pady=5)

    ttk.Button(frame, text="Convert", command=convert).grid(row=2, column=1, pady=10)

    progress.grid(row=3, column=0, columnspan=3, sticky="ew", padx=5)

    ttk.Label(frame, textvariable=status_var, wraplength=400, justify="left").grid(
        row=4, column=0, columnspan=3, padx=5, pady=5
    )


def build_render_tab(nb: ttk.Notebook) -> None:
    """Add the Render Transcripts UI to ``nb``.

    Allows the user to specify the target phone number whose messages should
    be labeled in the transcript output. The phone number may include common
    formatting characters (``+``, spaces, dashes); these are stripped before
    validation.
    """

    frame = ttk.Frame(nb)
    nb.add(frame, text="Render Transcripts")

    in_var = tk.StringVar()
    out_var = tk.StringVar()
    contacts_var = tk.StringVar()
    target_var = tk.StringVar()
    status_var = tk.StringVar()
    progress = ttk.Progressbar(frame, mode="indeterminate")

    def browse_in() -> None:
        path = filedialog.askdirectory(initialdir=in_var.get() or ".")
        if path:
            in_var.set(path)

    def browse_out() -> None:
        path = filedialog.askdirectory(initialdir=out_var.get() or ".")
        if path:
            out_var.set(path)

    def browse_contacts() -> None:
        path = filedialog.askopenfilename(
            title="Select contacts Excel file",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            contacts_var.set(path)

    def render() -> None:
        raw_target = target_var.get().strip()
        target = normalize_phone_number(raw_target)
        if len(target) != 11:
            status_var.set(
                "Target phone number must be 11 digits after removing formatting."
            )
            return

        contacts_path = Path(contacts_var.get()).expanduser()
        if not (contacts_var.get() and contacts_path.is_file()):
            status_var.set(f"Contacts file '{contacts_path}' does not exist.")
            return

        progress.start()
        status_var.set("Rendering...")

        def task() -> None:
            old_argv = sys.argv[:]
            try:
                sys.argv = [
                    "render-transcripts",
                    "--in",
                    in_var.get(),
                    "--out",
                    out_var.get(),
                    "--target-number",
                    target,
                    "--contacts-xlsx",
                    contacts_path.as_posix(),
                ]
                rt.main()
                msg = f"Rendered transcripts to '{out_var.get()}'"
            except Exception as e:  # pragma: no cover - user feedback
                msg = f"Error: {e}"
            finally:
                sys.argv = old_argv

            frame.after(0, lambda: [status_var.set(msg), progress.stop()])

        threading.Thread(target=task, daemon=True).start()

    ttk.Label(frame, text="Input folder:").grid(
        row=0, column=0, sticky="e", padx=5, pady=5
    )
    ttk.Entry(frame, textvariable=in_var, width=50).grid(
        row=0, column=1, padx=5, pady=5
    )
    ttk.Button(frame, text="Browse", command=browse_in).grid(
        row=0, column=2, padx=5, pady=5
    )

    ttk.Label(frame, text="Output folder:").grid(
        row=1, column=0, sticky="e", padx=5, pady=5
    )
    ttk.Entry(frame, textvariable=out_var, width=50).grid(
        row=1, column=1, padx=5, pady=5
    )
    ttk.Button(frame, text="Browse", command=browse_out).grid(
        row=1, column=2, padx=5, pady=5
    )

    ttk.Label(frame, text="Contacts file:").grid(
        row=2, column=0, sticky="e", padx=5, pady=5
    )
    ttk.Entry(frame, textvariable=contacts_var, width=50).grid(
        row=2, column=1, padx=5, pady=5
    )
    ttk.Button(frame, text="Browse", command=browse_contacts).grid(
        row=2, column=2, padx=5, pady=5
    )

    ttk.Label(
        frame,
        text="Target phone number (11 digits, formatting allowed):",
    ).grid(row=3, column=0, sticky="e", padx=5, pady=5)
    ttk.Entry(frame, textvariable=target_var, width=20).grid(
        row=3, column=1, padx=5, pady=5, sticky="w"
    )

    ttk.Button(frame, text="Render", command=render).grid(row=4, column=1, pady=10)

    progress.grid(row=5, column=0, columnspan=3, sticky="ew", padx=5)

    ttk.Label(frame, textvariable=status_var, wraplength=400, justify="left").grid(
        row=6, column=0, columnspan=3, padx=5, pady=5
    )


# ---------------------------------------------------------------------------
# Collect attachments tab
# ---------------------------------------------------------------------------


def build_collect_attachments_tab(nb: ttk.Notebook) -> None:
    """Add the Collect Attachments UI to ``nb``."""

    frame = ttk.Frame(nb)
    nb.add(frame, text="Collect Attachments")

    attachments_var = tk.StringVar()
    out_var = tk.StringVar()
    contacts_var = tk.StringVar()
    status_var = tk.StringVar()
    progress = ttk.Progressbar(frame, mode="indeterminate")

    def browse_attachments() -> None:
        path = filedialog.askdirectory(initialdir=attachments_var.get() or ".")
        if path:
            attachments_var.set(path)

    def browse_out() -> None:
        path = filedialog.askdirectory(initialdir=out_var.get() or ".")
        if path:
            out_var.set(path)

    def browse_contacts() -> None:
        path = filedialog.askopenfilename(
            title="Select contacts Excel file",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            contacts_var.set(path)

    def run() -> None:
        progress.start()

        def task() -> None:
            attachments_root = Path(attachments_var.get()).expanduser()
            compiled_path = Path(out_var.get()).expanduser()

            if not attachments_root.exists():
                frame.after(
                    0,
                    lambda: [
                        status_var.set(
                            f"Attachments folder '{attachments_root}' does not exist."
                        ),
                        progress.stop(),
                    ],
                )
                return

            try:
                compiled_path.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                frame.after(
                    0,
                    lambda: [
                        status_var.set(
                            f"Could not create output folder '{compiled_path}': {e}"
                        ),
                        progress.stop(),
                    ],
                )
                return

            logfile = (
                compiled_path
                / "compiled_attachment_log"
                / "compiled_attachment_log.xlsx"
            )
            contacts_path = contacts_var.get() or None
            try:
                records, exif_keys = ca.collect_attachments(
                    attachments_root, compiled_path, contacts_path
                )
                ca.write_excel(records, exif_keys, logfile)
                msg = (
                    f"Copied {len(records)} files from '{attachments_root}' to '{compiled_path}' and "
                    f"logged to '{logfile}'."
                )
            except Exception as e:  # pragma: no cover - user feedback
                msg = f"Error: {e}"

            frame.after(0, lambda: [status_var.set(msg), progress.stop()])

        threading.Thread(target=task, daemon=True).start()

    ttk.Label(frame, text="Attachments folder:").grid(
        row=0, column=0, sticky="e", padx=5, pady=5
    )
    ttk.Entry(frame, textvariable=attachments_var, width=50).grid(
        row=0, column=1, padx=5, pady=5
    )
    ttk.Button(frame, text="Browse", command=browse_attachments).grid(
        row=0, column=2, padx=5, pady=5
    )

    ttk.Label(frame, text="Output folder:").grid(
        row=1, column=0, sticky="e", padx=5, pady=5
    )
    ttk.Entry(frame, textvariable=out_var, width=50).grid(
        row=1, column=1, padx=5, pady=5
    )
    ttk.Button(frame, text="Browse", command=browse_out).grid(
        row=1, column=2, padx=5, pady=5
    )

    ttk.Label(frame, text="Contacts file:").grid(
        row=2, column=0, sticky="e", padx=5, pady=5
    )
    ttk.Entry(frame, textvariable=contacts_var, width=50).grid(
        row=2, column=1, padx=5, pady=5
    )
    ttk.Button(frame, text="Browse", command=browse_contacts).grid(
        row=2, column=2, padx=5, pady=5
    )

    ttk.Button(frame, text="Run", command=run).grid(row=3, column=1, pady=10)

    progress.grid(row=4, column=0, columnspan=3, sticky="ew", padx=5)

    ttk.Label(frame, textvariable=status_var, wraplength=400, justify="left").grid(
        row=5, column=0, columnspan=3, padx=5, pady=5
    )


# ---------------------------------------------------------------------------
# Collect quarantined files tab
# ---------------------------------------------------------------------------


def build_collect_quarantine_tab(nb: ttk.Notebook) -> None:
    """Add the Collect Quarantined Files UI to ``nb``."""

    frame = ttk.Frame(nb)
    nb.add(frame, text="Collect Quarantine Files")

    root_var = tk.StringVar()
    out_var = tk.StringVar()
    status_var = tk.StringVar()
    progress = ttk.Progressbar(frame, mode="indeterminate")

    def browse_root() -> None:
        path = filedialog.askdirectory(initialdir=root_var.get() or ".")
        if path:
            root_var.set(path)

    def browse_out() -> None:
        path = filedialog.askdirectory(initialdir=out_var.get() or ".")
        if path:
            out_var.set(path)

    def run() -> None:
        progress.start()

        def task() -> None:
            root_path = Path(root_var.get()).expanduser()
            compiled_path = Path(out_var.get()).expanduser()

            if not root_path.exists():
                frame.after(
                    0,
                    lambda: [
                        status_var.set(f"Root folder '{root_path}' does not exist."),
                        progress.stop(),
                    ],
                )
                return

            try:
                compiled_path.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                frame.after(
                    0,
                    lambda: [
                        status_var.set(
                            f"Could not create output folder '{compiled_path}': {e}"
                        ),
                        progress.stop(),
                    ],
                )
                return

            try:
                copied, skipped, total = (
                    collect_quarantined_files.collect_quarantined_files(
                        root_path, compiled_path
                    )
                )
                msg = (
                    f"Converted {len(copied)} of {total} files from '{root_path}' to '{compiled_path}'."
                )
                if skipped:
                    skipped_str = ", ".join(str(p) for p in skipped)
                    msg += f" Skipped {len(skipped)} files: {skipped_str}."
            except Exception as e:  # pragma: no cover - user feedback
                msg = f"Error: {e}"

            frame.after(0, lambda: [status_var.set(msg), progress.stop()])

        threading.Thread(target=task, daemon=True).start()

    ttk.Label(frame, text="Root folder:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
    ttk.Entry(frame, textvariable=root_var, width=50).grid(row=0, column=1, padx=5, pady=5)
    ttk.Button(frame, text="Browse", command=browse_root).grid(row=0, column=2, padx=5, pady=5)

    ttk.Label(frame, text="Output folder:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
    ttk.Entry(frame, textvariable=out_var, width=50).grid(row=1, column=1, padx=5, pady=5)
    ttk.Button(frame, text="Browse", command=browse_out).grid(row=1, column=2, padx=5, pady=5)

    ttk.Button(frame, text="Run", command=run).grid(row=2, column=1, pady=10)

    progress.grid(row=3, column=0, columnspan=3, sticky="ew", padx=5)

    ttk.Label(frame, textvariable=status_var, wraplength=400, justify="left").grid(
        row=4, column=0, columnspan=3, padx=5, pady=5
    )


# ---------------------------------------------------------------------------
# Decrypt/unzip tab
# ---------------------------------------------------------------------------


def build_decrypt_unzip_tab(nb: ttk.Notebook) -> None:
    """Add the Decrypt & Unzip UI to ``nb``."""

    frame = ttk.Frame(nb)
    nb.add(frame, text="Decrypt & Unzip")

    archive_var = tk.StringVar()
    output_var = tk.StringVar()
    password_var = tk.StringVar()
    status_var = tk.StringVar()
    progress = ttk.Progressbar(frame, mode="indeterminate")

    def browse_archive() -> None:
        path = filedialog.askopenfilename(
            title="Select encrypted archive",
            filetypes=[("Zip files", "*.zip"), ("All files", "*.*")],
        )
        if path:
            archive_var.set(path)

    def browse_output() -> None:
        path = filedialog.askdirectory(initialdir=output_var.get() or ".")
        if path:
            output_var.set(path)

    def run() -> None:
        progress.start()
        status_var.set("Running...")

        def task() -> None:
            archive = Path(archive_var.get()).expanduser()
            out_dir = Path(output_var.get()).expanduser() if output_var.get() else None
            try:

                archive = Path(archive_var.get()).expanduser()
                output_dir = (
                    Path(output_var.get()).expanduser() if output_var.get() else None
                )
                decrypt_unzip.decrypt_and_unzip(
                    archive, password_var.get(), output_dir
                )
                dest = output_dir or archive.parent / archive.stem

                msg = f"Decrypted archive to '{dest}'"
            except Exception as e:  # pragma: no cover - user feedback
                msg = f"Error: {e}"

            frame.after(0, lambda: [status_var.set(msg), progress.stop()])

        threading.Thread(target=task, daemon=True).start()

    ttk.Label(frame, text="Encrypted archive:").grid(
        row=0, column=0, sticky="e", padx=5, pady=5
    )
    ttk.Entry(frame, textvariable=archive_var, width=50).grid(
        row=0, column=1, padx=5, pady=5
    )
    ttk.Button(frame, text="Browse", command=browse_archive).grid(
        row=0, column=2, padx=5, pady=5
    )

    ttk.Label(frame, text="Output folder:").grid(
        row=1, column=0, sticky="e", padx=5, pady=5
    )
    ttk.Entry(frame, textvariable=output_var, width=50).grid(
        row=1, column=1, padx=5, pady=5
    )
    ttk.Button(frame, text="Browse", command=browse_output).grid(
        row=1, column=2, padx=5, pady=5
    )

    ttk.Label(frame, text="Password:").grid(
        row=2, column=0, sticky="e", padx=5, pady=5
    )
    ttk.Entry(frame, textvariable=password_var, show="*", width=20).grid(
        row=2, column=1, padx=5, pady=5, sticky="w"
    )

    ttk.Button(frame, text="Run", command=run).grid(row=3, column=1, pady=10)

    progress.grid(row=4, column=0, columnspan=3, sticky="ew", padx=5)

    ttk.Label(frame, textvariable=status_var, wraplength=400, justify="left").grid(
        row=5, column=0, columnspan=3, padx=5, pady=5
    )


# ---------------------------------------------------------------------------
# Main application
# ---------------------------------------------------------------------------


def _set_window_icon(root: tk.Tk) -> None:
    """Replace the default Tcl feather with the Omega app icon when available."""
    base = Path(__file__).resolve().parent
    ico = base / "app_icon.ico"
    png = base / "app_icon.png"
    try:
        if sys.platform.startswith("win") and ico.is_file():
            root.iconbitmap(default=str(ico))
            return
    except Exception:
        pass
    try:
        if png.is_file():
            # Keep a reference so Tk does not garbage-collect the image
            img = tk.PhotoImage(file=str(png))
            root.iconphoto(True, img)
            root._app_icon_image = img  # type: ignore[attr-defined]
    except Exception:
        pass


def main() -> None:  # pragma: no cover - GUI entry point
    root = tk.Tk()
    root.title(f"Synchronoss Unified Toolbox  v{APP_VERSION}")
    root.minsize(720, 520)
    _set_window_icon(root)

    # Single automatic workflow – selection.zip → complete case folder
    build_full_pipeline_ui(root)

    root.mainloop()


if __name__ == "__main__":  # pragma: no cover - GUI entry point
    main()

