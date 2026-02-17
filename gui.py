#!/usr/bin/env python3
"""
GUI for PAN Kontaktliste: select Excel file and HTML destination, then generate contact list.
Uses wxPython for a native look on Windows, macOS, and Linux.
"""
from __future__ import annotations

import sys
import tempfile
import webbrowser
from pathlib import Path

import wx
import wx.adv

# Project modules
from excel_reader import load_participants
from render import render_html
from version import get_version


def _resource_path(relative: str) -> Path:
    """Path to a file in the project (e.g. data/placeholder.png). Supports PyInstaller frozen exe."""
    if getattr(sys, "frozen", False):
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).resolve().parent
    return base / relative


class MainFrame(wx.Frame):
    def __init__(self) -> None:
        super().__init__(None, title="PAN Kontaktliste", size=(580, 260))
        self.SetMinSize((520, 240))

        self._panel = wx.Panel(self)
        panel = self._panel
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Excel row (label has MinSize so it isn't clipped on Windows)
        row1 = wx.BoxSizer(wx.HORIZONTAL)
        lbl_xlsx = wx.StaticText(panel, label="Excel-Datei:")
        w = lbl_xlsx.GetTextExtent("Excel-Datei:")[0]
        lbl_xlsx.SetMinSize((max(w, 100) + 8, -1))
        row1.Add(lbl_xlsx, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 8)
        self.xlsx_path = wx.TextCtrl(panel, value="", size=(320, -1))
        row1.Add(self.xlsx_path, 1, wx.EXPAND | wx.RIGHT, 6)
        btn_xlsx = wx.Button(panel, label="Durchsuchen ...")
        btn_xlsx.Bind(wx.EVT_BUTTON, self._on_choose_xlsx)
        row1.Add(btn_xlsx, 0)
        sizer.Add(row1, 0, wx.EXPAND | wx.ALL, 6)

        # HTML row (label has MinSize so it isn't clipped on Windows)
        row2 = wx.BoxSizer(wx.HORIZONTAL)
        lbl_html = wx.StaticText(panel, label="HTML-Datei speichern unter:")
        w = lbl_html.GetTextExtent("HTML-Datei speichern unter:")[0]
        lbl_html.SetMinSize((max(w, 220) + 8, -1))
        row2.Add(lbl_html, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 8)
        self.html_path = wx.TextCtrl(panel, value="", size=(320, -1))
        row2.Add(self.html_path, 1, wx.EXPAND | wx.RIGHT, 6)
        btn_html = wx.Button(panel, label="Durchsuchen ...")
        btn_html.Bind(wx.EVT_BUTTON, self._on_choose_html)
        row2.Add(btn_html, 0)
        sizer.Add(row2, 0, wx.EXPAND | wx.ALL, 6)

        # Checkbox
        self.open_browser_cb = wx.CheckBox(
            panel, label="HTML nach dem Erstellen im Browser öffnen"
        )
        self.open_browser_cb.SetValue(True)
        sizer.Add(self.open_browser_cb, 0, wx.LEFT | wx.TOP, 8)

        # Create button
        create_btn = wx.Button(panel, label="Kontaktliste erstellen")
        create_btn.Bind(wx.EVT_BUTTON, self._on_create_list)
        sizer.Add(create_btn, 0, wx.ALL, 16)

        # Menu: Help → About
        menubar = wx.MenuBar()
        help_menu = wx.Menu()
        about_item = help_menu.Append(wx.ID_ABOUT, "Über PAN Kontaktliste...")
        self.Bind(wx.EVT_MENU, self._on_about, about_item)
        menubar.Append(help_menu, "Hilfe")
        self.SetMenuBar(menubar)

        panel.SetSizer(sizer)
        panel.Layout()
        self.Bind(wx.EVT_SHOW, self._on_show)

    def _on_show(self, event: wx.ShowEvent) -> None:
        """Force layout on first show so the window displays correctly on Windows."""
        if event.IsShown():
            wx.CallAfter(self._do_layout)

    def _do_layout(self) -> None:
        self._panel.Layout()
        self.Layout()

    def _on_about(self, _event: wx.CommandEvent) -> None:
        info = wx.adv.AboutDialogInfo()
        info.SetName("PAN Kontaktliste")
        info.SetVersion(get_version())
        info.SetDescription(
            "Erstellt aus einer Excel-Anmeldeliste eine HTML-Kontaktliste für Teilnehmerinnen und Teilnehmer "
            "(einwilligungsbasiert). Lizenz: GPL-3.0-or-later."
        )
        info.SetLicense(
            "Dieses Programm steht unter der GNU General Public License v3.0 (GPLv3).\n"
            "Vollständiger Lizenztext: siehe LICENSE im Projekt oder https://www.gnu.org/licenses/gpl-3.0.html"
        )
        info.SetWebSite("https://github.com/nomike/pan-kontaktliste")
        wx.adv.AboutBox(info)

    def _on_choose_xlsx(self, _event: wx.CommandEvent) -> None:
        with wx.FileDialog(
            self,
            "Excel-Datei wählen",
            wildcard="Excel-Dateien (*.xlsx)|*.xlsx|Alle Dateien (*.*)|*.*",
            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST,
        ) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                self.xlsx_path.SetValue(dlg.GetPath())

    def _on_choose_html(self, _event: wx.CommandEvent) -> None:
        with wx.FileDialog(
            self,
            "HTML-Datei speichern unter",
            defaultFile="",
            wildcard="HTML-Dateien (*.html)|*.html|Alle Dateien (*.*)|*.*",
            style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT,
        ) as dlg:
            if dlg.ShowModal() == wx.ID_OK:
                path = dlg.GetPath()
                if not path.endswith(".html"):
                    path += ".html"
                self.html_path.SetValue(path)

    def _on_create_list(self, _event: wx.CommandEvent) -> None:
        xlsx = self.xlsx_path.GetValue().strip()
        html = self.html_path.GetValue().strip()
        if not xlsx:
            wx.MessageBox(
                "Bitte wählen Sie eine Excel-Datei.",
                "Eingabe fehlt",
                wx.OK | wx.ICON_WARNING,
            )
            return
        if not html:
            wx.MessageBox(
                "Bitte wählen Sie einen Speicherort für die HTML-Datei.",
                "Eingabe fehlt",
                wx.OK | wx.ICON_WARNING,
            )
            return

        placeholder = _resource_path("data/placeholder.png")
        if not placeholder.exists():
            wx.MessageBox(
                f"Platzhalterbild fehlt: {placeholder}\nBitte legen Sie data/placeholder.png ab.",
                "Fehler",
                wx.OK | wx.ICON_ERROR,
            )
            return

        try:
            with tempfile.TemporaryDirectory(prefix="pan_contact_") as build_dir:
                build_path = Path(build_dir)
                participants = load_participants(xlsx, placeholder, image_output_dir=build_path)
                if not participants:
                    wx.MessageBox(
                        "In der Excel-Datei sind keine Einträge mit aktivierter Teilnehmyliste.",
                        "Keine Teilnehmer",
                        wx.OK | wx.ICON_INFORMATION,
                    )
                    return
                render_html(participants, Path(html))
            msg = f"Die Kontaktliste wurde erstellt:\n{html}"
            if self.open_browser_cb.GetValue():
                webbrowser.open(f"file://{Path(html).resolve()}")
                msg += "\n\nDie Liste wurde im Browser geöffnet. Zum Erzeugen einer PDF: Drucken → Als PDF speichern."
            wx.MessageBox(msg, "Fertig", wx.OK | wx.ICON_INFORMATION)
        except FileNotFoundError as e:
            wx.MessageBox(str(e), "Datei fehlt", wx.OK | wx.ICON_ERROR)
        except Exception as e:
            wx.MessageBox(str(e), "Fehler", wx.OK | wx.ICON_ERROR)


def run_gui() -> None:
    app = wx.App()
    frame = MainFrame()
    frame.Show()
    app.MainLoop()


if __name__ == "__main__":
    run_gui()
    sys.exit(0)
