#!/usr/bin/env python3
"""
GUI for PAN Kontaktliste: select Excel file and HTML destination, then generate contact list.
Uses ttk for native-looking widgets on Linux, Windows, and macOS.
"""
from __future__ import annotations

import sys
import tempfile
import webbrowser
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from tkinter import ttk

# Project modules
from excel_reader import load_participants
from render import render_html


def _resource_path(relative: str) -> Path:
    """Path to a file in the project (e.g. data/placeholder.png)."""
    return Path(__file__).resolve().parent / relative


def run_gui() -> None:
    root = tk.Tk()
    root.title("PAN Kontaktliste")

    # Ensure window is large enough for all controls; use minsize so it can't be shrunk too small
    root.minsize(520, 240)
    root.geometry("580x260")
    root.resizable(True, True)

    # Use a frame with padding so content isn't flush against the edges
    main = ttk.Frame(root, padding=12)
    main.pack(fill=tk.BOTH, expand=True)

    xlsx_path_var: tk.StringVar = tk.StringVar(value="")
    html_path_var: tk.StringVar = tk.StringVar(value="")
    open_browser_var: tk.BooleanVar = tk.BooleanVar(value=True)

    def choose_xlsx() -> None:
        path = filedialog.askopenfilename(
            title="Excel-Datei wählen",
            filetypes=[
                ("Excel-Dateien", "*.xlsx"),
                ("Alle Dateien", "*.*"),
            ],
        )
        if path:
            xlsx_path_var.set(path)

    def choose_html() -> None:
        path = filedialog.asksaveasfilename(
            title="HTML-Datei speichern unter",
            defaultextension=".html",
            filetypes=[
                ("HTML-Dateien", "*.html"),
                ("Alle Dateien", "*.*"),
            ],
        )
        if path:
            html_path_var.set(path)

    def create_list() -> None:
        xlsx = xlsx_path_var.get().strip()
        html = html_path_var.get().strip()
        if not xlsx:
            messagebox.showwarning("Eingabe fehlt", "Bitte wählen Sie eine Excel-Datei.")
            return
        if not html:
            messagebox.showwarning("Eingabe fehlt", "Bitte wählen Sie einen Speicherort für die HTML-Datei.")
            return

        placeholder = _resource_path("data/placeholder.png")
        if not placeholder.exists():
            messagebox.showerror(
                "Fehler",
                f"Platzhalterbild fehlt: {placeholder}\nBitte legen Sie data/placeholder.png ab.",
            )
            return

        try:
            with tempfile.TemporaryDirectory(prefix="pan_contact_") as build_dir:
                build_path = Path(build_dir)
                participants = load_participants(xlsx, placeholder, image_output_dir=build_path)
                if not participants:
                    messagebox.showinfo(
                        "Keine Teilnehmer",
                        "In der Excel-Datei sind keine Einträge mit aktivierter Teilnehmyliste.",
                    )
                    return
                render_html(participants, Path(html))
            msg = f"Die Kontaktliste wurde erstellt:\n{html}"
            if open_browser_var.get():
                webbrowser.open(f"file://{Path(html).resolve()}")
                msg += "\n\nDie Liste wurde im Browser geöffnet. Zum Erzeugen einer PDF: Drucken → Als PDF speichern."
            messagebox.showinfo("Fertig", msg)
        except FileNotFoundError as e:
            messagebox.showerror("Datei fehlt", str(e))
        except Exception as e:
            messagebox.showerror("Fehler", str(e))

    row = 0
    ttk.Label(main, text="Excel-Datei:").grid(row=row, column=0, sticky=tk.W, pady=6)
    ttk.Entry(main, textvariable=xlsx_path_var, width=50).grid(row=row, column=1, padx=6, pady=6, sticky=tk.EW)
    ttk.Button(main, text="Durchsuchen ...", command=choose_xlsx).grid(row=row, column=2, pady=6)
    row += 1

    ttk.Label(main, text="HTML-Datei speichern unter:").grid(row=row, column=0, sticky=tk.W, pady=6)
    ttk.Entry(main, textvariable=html_path_var, width=50).grid(row=row, column=1, padx=6, pady=6, sticky=tk.EW)
    ttk.Button(main, text="Durchsuchen ...", command=choose_html).grid(row=row, column=2, pady=6)
    row += 1

    ttk.Checkbutton(
        main,
        text="HTML nach dem Erstellen im Browser öffnen",
        variable=open_browser_var,
    ).grid(row=row, column=1, sticky=tk.W, pady=8)
    row += 1

    create_btn = ttk.Button(main, text="Kontaktliste erstellen", command=create_list)
    create_btn.grid(row=row, column=1, pady=16)

    main.columnconfigure(1, weight=1)

    root.mainloop()


if __name__ == "__main__":
    run_gui()
    sys.exit(0)
