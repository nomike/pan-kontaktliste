#!/usr/bin/env python3
"""
GUI for PAN Kontaktliste: log in to SeaTable, pick a meetup base, generate HTML.
Uses wxPython for a native look on Windows, macOS, and Linux.
Credentials are kept in memory only for the lifetime of the process.
"""
from __future__ import annotations

import sys
import tempfile
import webbrowser
from pathlib import Path

import wx
import wx.adv

try:
    import wx.svg

    _HAS_SVG = True
except ImportError:
    _HAS_SVG = False

from render import render_html
from seatable_reader import (
    DEFAULT_SERVER_URL,
    BaseInfo,
    SeaTableAuthError,
    SeaTableError,
    SeaTableSession,
    list_bases,
    load_participants,
    login,
)
from version import get_version


def _resource_path(relative: str) -> Path:
    """Path to a file in the project (e.g. data/placeholder.png). Supports PyInstaller frozen exe."""
    if getattr(sys, "frozen", False):
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).resolve().parent
    return base / relative


class LoginDialog(wx.Dialog):
    """Ask for SeaTable username and password (not persisted)."""

    def __init__(self, parent: wx.Window | None = None) -> None:
        super().__init__(parent, title="Anmeldung bei SeaTable", size=(460, 240))
        self.SetMinSize((420, 220))

        sizer = wx.BoxSizer(wx.VERTICAL)

        grid = wx.FlexGridSizer(3, 2, 8, 8)
        grid.AddGrowableCol(1, 1)

        grid.Add(wx.StaticText(self, label="Server:"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.server = wx.TextCtrl(self, value=DEFAULT_SERVER_URL)
        grid.Add(self.server, 1, wx.EXPAND)

        grid.Add(
            wx.StaticText(self, label="E-Mail / Benutzername:"),
            0,
            wx.ALIGN_CENTER_VERTICAL,
        )
        self.username = wx.TextCtrl(self)
        grid.Add(self.username, 1, wx.EXPAND)

        grid.Add(wx.StaticText(self, label="Passwort:"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.password = wx.TextCtrl(self, style=wx.TE_PASSWORD)
        grid.Add(self.password, 1, wx.EXPAND)

        sizer.Add(grid, 0, wx.EXPAND | wx.ALL, 12)

        hint = wx.StaticText(
            self,
            label="Zugangsdaten werden nur im Speicher gehalten und nicht gespeichert.",
        )
        hint.Wrap(420)
        sizer.Add(hint, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 12)

        btns = self.CreateStdDialogButtonSizer(wx.OK | wx.CANCEL)
        if btns:
            sizer.Add(btns, 0, wx.EXPAND | wx.ALL, 8)

        self.SetSizer(sizer)
        self.Layout()
        self.username.SetFocus()

    def values(self) -> tuple[str, str, str]:
        return (
            self.username.GetValue().strip(),
            self.password.GetValue(),
            self.server.GetValue().strip() or DEFAULT_SERVER_URL,
        )


class MainFrame(wx.Frame):
    def __init__(self, session: SeaTableSession) -> None:
        super().__init__(None, title="PAN Kontaktliste", size=(640, 520))
        self.SetMinSize((560, 440))
        self._session = session
        self._bases: list[BaseInfo] = []
        self._filtered: list[BaseInfo] = []
        self._app_icon = None
        self._set_icon()

        self._panel = wx.Panel(self)
        panel = self._panel
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Filter + refresh
        row_filter = wx.BoxSizer(wx.HORIZONTAL)
        lbl_filter = wx.StaticText(panel, label="Base filtern:")
        row_filter.Add(lbl_filter, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 8)
        self.filter_ctrl = wx.TextCtrl(panel)
        self.filter_ctrl.SetHint("z. B. Wintertreffen")
        self.filter_ctrl.Bind(wx.EVT_TEXT, self._on_filter_changed)
        row_filter.Add(self.filter_ctrl, 1, wx.EXPAND | wx.RIGHT, 6)
        btn_refresh = wx.Button(panel, label="Aktualisieren")
        btn_refresh.Bind(wx.EVT_BUTTON, self._on_refresh_bases)
        row_filter.Add(btn_refresh, 0)
        sizer.Add(row_filter, 0, wx.EXPAND | wx.ALL, 6)

        sizer.Add(
            wx.StaticText(panel, label="Treffen (SeaTable-Base) wählen:"),
            0,
            wx.LEFT | wx.RIGHT | wx.TOP,
            6,
        )
        self.base_list = wx.ListBox(panel, style=wx.LB_SINGLE)
        self.base_list.Bind(wx.EVT_LISTBOX, self._on_base_selected)
        sizer.Add(self.base_list, 1, wx.EXPAND | wx.ALL, 6)

        # Meetup name row
        row_meetup = wx.BoxSizer(wx.HORIZONTAL)
        lbl_meetup = wx.StaticText(panel, label="Name des Treffens:")
        w = lbl_meetup.GetTextExtent("Name des Treffens:")[0]
        lbl_meetup.SetMinSize((max(w, 120) + 8, -1))
        row_meetup.Add(lbl_meetup, 0, wx.ALIGN_CENTER_VERTICAL | wx.RIGHT, 8)
        self.meetup_name = wx.TextCtrl(panel, value="", size=(320, -1))
        self.meetup_name.SetHint("z. B. PAN Wintertreffen 2026")
        row_meetup.Add(self.meetup_name, 1, wx.EXPAND)
        sizer.Add(row_meetup, 0, wx.EXPAND | wx.ALL, 6)

        # HTML row
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

        self.open_browser_cb = wx.CheckBox(
            panel, label="HTML nach dem Erstellen im Browser öffnen"
        )
        self.open_browser_cb.SetValue(True)
        sizer.Add(self.open_browser_cb, 0, wx.LEFT | wx.TOP, 8)

        create_btn = wx.Button(panel, label="Kontaktliste erstellen")
        create_btn.Bind(wx.EVT_BUTTON, self._on_create_list)
        sizer.Add(create_btn, 0, wx.ALL, 16)

        menubar = wx.MenuBar()
        help_menu = wx.Menu()
        about_item = help_menu.Append(wx.ID_ABOUT, "Über PAN Kontaktliste...")
        self.Bind(wx.EVT_MENU, self._on_about, about_item)
        menubar.Append(help_menu, "Hilfe")
        self.SetMenuBar(menubar)

        panel.SetSizer(sizer)
        panel.Layout()
        self.Bind(wx.EVT_SHOW, self._on_show)
        wx.CallAfter(self._load_bases)

    def _set_icon(self) -> None:
        if not _HAS_SVG:
            return
        icon_path = _resource_path("data/polyamory-logo.svg")
        if not icon_path.exists():
            return
        try:
            svg_img = wx.svg.SVGimage.CreateFromFile(str(icon_path))
            icon_bundle = wx.IconBundle()
            for size in (16, 32, 48, 64, 128, 256):
                bmp = svg_img.ConvertToScaledBitmap(wx.Size(size, size))
                icon = wx.Icon()
                icon.CopyFromBitmap(bmp)
                icon_bundle.AddIcon(icon)
            self.SetIcons(icon_bundle)
            bmp = svg_img.ConvertToScaledBitmap(wx.Size(64, 64))
            self._app_icon = wx.Icon()
            self._app_icon.CopyFromBitmap(bmp)
        except Exception:
            pass

    def _on_show(self, event: wx.ShowEvent) -> None:
        if event.IsShown():
            wx.CallAfter(self._do_layout)

    def _do_layout(self) -> None:
        self._panel.Layout()
        self.Layout()

    def _on_about(self, _event: wx.CommandEvent) -> None:
        info = wx.adv.AboutDialogInfo()
        if self._app_icon:
            info.SetIcon(self._app_icon)
        info.SetName("PAN Kontaktliste")
        info.SetVersion(get_version())
        info.SetDescription(
            "Erstellt aus einer SeaTable-Anmeldeliste eine HTML-Kontaktliste für "
            "Teilnehmerinnen und Teilnehmer (einwilligungsbasiert). Lizenz: GPL-3.0-or-later."
        )
        info.SetLicense(
            "Dieses Programm steht unter der GNU General Public License v3.0 (GPLv3).\n"
            "Vollständiger Lizenztext: siehe LICENSE im Projekt oder https://www.gnu.org/licenses/gpl-3.0.html"
        )
        info.SetWebSite("https://github.com/nomike/pan-kontaktliste")
        wx.adv.AboutBox(info)

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

    def _on_refresh_bases(self, _event: wx.CommandEvent) -> None:
        self._load_bases()

    def _load_bases(self) -> None:
        try:
            wx.BeginBusyCursor()
            self._bases = list_bases(self._session)
        except SeaTableError as e:
            wx.MessageBox(str(e), "Fehler", wx.OK | wx.ICON_ERROR)
            self._bases = []
        except Exception as e:
            wx.MessageBox(str(e), "Fehler", wx.OK | wx.ICON_ERROR)
            self._bases = []
        finally:
            try:
                wx.EndBusyCursor()
            except Exception:
                pass
        self._apply_filter()

    def _on_filter_changed(self, _event: wx.CommandEvent) -> None:
        self._apply_filter()

    def _apply_filter(self) -> None:
        needle = self.filter_ctrl.GetValue().strip().lower()
        if needle:
            self._filtered = [
                b
                for b in self._bases
                if needle in b.name.lower() or needle in b.workspace_name.lower()
            ]
        else:
            self._filtered = list(self._bases)
        self.base_list.Set([b.label for b in self._filtered])

    def _selected_base(self) -> BaseInfo | None:
        idx = self.base_list.GetSelection()
        if idx == wx.NOT_FOUND or idx < 0 or idx >= len(self._filtered):
            return None
        return self._filtered[idx]

    def _on_base_selected(self, _event: wx.CommandEvent) -> None:
        base = self._selected_base()
        if base:
            self.meetup_name.SetValue(base.name)

    def _on_create_list(self, _event: wx.CommandEvent) -> None:
        base = self._selected_base()
        html = self.html_path.GetValue().strip()
        if base is None:
            wx.MessageBox(
                "Bitte wählen Sie ein Treffen (SeaTable-Base).",
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
            wx.BeginBusyCursor()
            with tempfile.TemporaryDirectory(prefix="pan_contact_") as build_dir:
                build_path = Path(build_dir)
                participants = load_participants(
                    self._session,
                    base.workspace_id,
                    base.name,
                    placeholder,
                    image_output_dir=build_path,
                )
                if not participants:
                    wx.MessageBox(
                        "In der Base sind keine Einträge mit aktivierter Teilnehmyliste.",
                        "Keine Teilnehmer",
                        wx.OK | wx.ICON_INFORMATION,
                    )
                    return
                meetup_name = self.meetup_name.GetValue().strip() or base.name
                render_html(participants, Path(html), meetup_name=meetup_name)
            msg = f"Die Kontaktliste wurde erstellt:\n{html}"
            if self.open_browser_cb.GetValue():
                webbrowser.open(f"file://{Path(html).resolve()}")
                msg += (
                    "\n\nDie Liste wurde im Browser geöffnet. "
                    "Zum Erzeugen einer PDF: Drucken → Als PDF speichern."
                )
            wx.MessageBox(msg, "Fertig", wx.OK | wx.ICON_INFORMATION)
        except (SeaTableError, SeaTableAuthError) as e:
            wx.MessageBox(str(e), "SeaTable-Fehler", wx.OK | wx.ICON_ERROR)
        except FileNotFoundError as e:
            wx.MessageBox(str(e), "Datei fehlt", wx.OK | wx.ICON_ERROR)
        except Exception as e:
            wx.MessageBox(str(e), "Fehler", wx.OK | wx.ICON_ERROR)
        finally:
            try:
                wx.EndBusyCursor()
            except Exception:
                pass


def _prompt_login() -> SeaTableSession | None:
    """Show login dialog until success or cancel. Returns None if user cancels."""
    while True:
        dlg = LoginDialog(None)
        result = dlg.ShowModal()
        if result != wx.ID_OK:
            dlg.Destroy()
            return None
        username, password, server = dlg.values()
        dlg.Destroy()
        try:
            wx.BeginBusyCursor()
            return login(username, password, server)
        except SeaTableAuthError as e:
            wx.MessageBox(str(e), "Anmeldung fehlgeschlagen", wx.OK | wx.ICON_ERROR)
        except Exception as e:
            wx.MessageBox(str(e), "Anmeldung fehlgeschlagen", wx.OK | wx.ICON_ERROR)
        finally:
            try:
                wx.EndBusyCursor()
            except Exception:
                pass


def run_gui() -> None:
    app = wx.App()
    session = _prompt_login()
    if session is None:
        return
    frame = MainFrame(session)
    frame.Show()
    app.MainLoop()


if __name__ == "__main__":
    run_gui()
    sys.exit(0)
