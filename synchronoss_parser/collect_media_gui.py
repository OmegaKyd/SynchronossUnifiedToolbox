#!/usr/bin/env python3
"""Simple Tkinter GUI wrapper around collect_media.py."""

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, ttk

from .collect_media import collect_media, write_excel


def main():
    window = tk.Tk()
    window.title("Collect Media")

    in_var = tk.StringVar()
    out_var = tk.StringVar()
    status_var = tk.StringVar()

    def browse_in():
        path = filedialog.askdirectory(initialdir=in_var.get() or ".")
        if path:
            in_var.set(path)

    def browse_out():
        path = filedialog.askdirectory(initialdir=out_var.get() or ".")
        if path:
            out_var.set(path)

    def run():
        root_path = Path(in_var.get()).expanduser()
        compiled_path = Path(out_var.get()).expanduser()

        if not root_path.exists():
            status_var.set(f"Input folder '{root_path}' does not exist.")
            return

        try:
            compiled_path.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            status_var.set(f"Could not create output folder '{compiled_path}': {e}")
            return

        logfile = compiled_path / "compiled_media_log" / "compiled_media_log.xlsx"

        progress.start()
        window.update_idletasks()
        try:
            records, exif_keys = collect_media(root_path, compiled_path)
            write_excel(records, exif_keys, logfile)
            status_var.set(
                f"Copied {len(records)} files from '{root_path}' to '{compiled_path}' and logged to '{logfile}'."
            )
        except Exception as e:  # pragma: no cover - user feedback
            status_var.set(f"Error: {e}")
        finally:
            progress.stop()

    tk.Label(window, text="'VZMOBILE' Folder Path:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
    tk.Entry(window, textvariable=in_var, width=50).grid(row=0, column=1, padx=5)
    tk.Button(window, text="Browse", command=browse_in).grid(row=0, column=2, padx=5)

    tk.Label(window, text="Output folder:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
    tk.Entry(window, textvariable=out_var, width=50).grid(row=1, column=1, padx=5)
    tk.Button(window, text="Browse", command=browse_out).grid(row=1, column=2, padx=5)

    tk.Button(window, text="Run", command=run).grid(row=2, column=1, pady=10)

    progress = ttk.Progressbar(window, mode="indeterminate")
    progress.grid(row=3, column=0, columnspan=3, sticky="ew", padx=5)

    tk.Label(window, textvariable=status_var, wraplength=400, justify="left").grid(
        row=4, column=0, columnspan=3, padx=5, pady=5
    )

    window.mainloop()


if __name__ == "__main__":  # pragma: no cover - GUI entry point
    main()
