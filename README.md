# PAN Kontaktliste

Dieses Programm erstellt aus einer Excel-Anmeldeliste (PAN-Treffen) eine **Kontaktliste** für Teilnehmerinnen und Teilnehmer. Die Ausgabe ist eine **HTML-Datei**, die in jedem Webbrowser geöffnet werden kann. Eine PDF lässt sich direkt im Browser erzeugen (Drucken → Als PDF speichern). Es sind keine zusätzlichen Installationen wie LaTeX nötig – unter Windows, Linux und macOS reicht Python und ein Browser.

Es werden die Einwilligungen aus dem Anmeldeformular berücksichtigt: Nur wer der Teilnehmyliste zugestimmt hat, erscheint in der Liste; E-Mail, Telefon, Nachname, Vorname und Bild werden nur angezeigt, wenn die jeweilige Option gewählt wurde. Fehlt die Einwilligung für ein Bild, wird ein Platzhalterbild verwendet.

## Anforderungen

- **Python 3.10+**
- Ein **Webbrowser** (zum Anzeigen der HTML-Liste und zum Erzeugen einer PDF per Drucken → Als PDF speichern)

## Installation

1. Repository klonen bzw. in den Projektordner wechseln.
2. Virtuelle Umgebung anlegen und aktivieren:

   ```bash
   python3 -m venv .venv
   source .venv/bin/activate   # Linux/macOS
   # oder unter Windows: .venv\Scripts\activate
   ```
3. Abhängigkeiten installieren:

   ```bash
   pip install -r requirements.txt
   ```

## Nutzung

### Grafische Oberfläche (empfohlen)

```bash
python gui.py
```

- **Excel-Datei:** Über „Durchsuchen …“ die Anmeldeliste (`.xlsx`) wählen.
- **HTML-Datei speichern unter:** Zielpfad und Dateiname für die HTML-Datei angeben.
- Optional: „HTML nach dem Erstellen im Browser öffnen“ aktivieren – dann öffnet sich die Liste nach dem Erstellen automatisch.
- **Kontaktliste erstellen** startet die Verarbeitung.

Zum Erzeugen einer PDF: HTML im Browser öffnen → Menü Drucken (oder Strg+P) → „Als PDF speichern“ bzw. „Save as PDF“ wählen.

### Ablauf im Programm

1. Aus der Excel-Datei werden nur Zeilen mit aktivierter **Teilnehmyliste** übernommen.
2. Pro Teilnehmer/in werden immer **Land**, **Rufname/Pseudonym** und **Teilnehmyliste_Couch** in die Liste übernommen.
3. **E-Mail**, **Telefonnummer**, **Nachname**, **Vorname** und **Bild** erscheinen nur, wenn die jeweilige Einwilligung gesetzt ist.
4. Ist keine Einwilligung für ein Bild vorhanden oder kein Bild hinterlegt, wird das Platzhalterbild aus `data/placeholder.png` verwendet.
5. Die Liste wird als eine einzige HTML-Datei mit eingebetteten Bildern (Data-URLs) erzeugt – die Datei kann ohne weitere Ressourcen weitergegeben werden.

## Projektstruktur

- `gui.py` – Einstieg für die grafische Oberfläche (wxPython; Dateiauswahl, Aufruf von Excel-Leser und HTML-Erstellung)
- `excel_reader.py` – Einlesen der Excel-Datei, Filterung nach Einwilligungen, Extraktion von Bildern
- `render.py` – Jinja2-Rendering der HTML-Vorlage (Bilder als Data-URLs)
- `template/contact_list.html.j2` – HTML-Vorlage (Jinja2) für die Kontaktliste
- `data/placeholder.png` – Platzhalterbild, wenn kein Bild oder keine Einwilligung
- `requirements.txt` – Python-Abhängigkeiten

## Lizenz

Dieses Projekt steht unter der **GNU General Public License v3.0** (GPLv3). Siehe die Datei [LICENSE](LICENSE) für den vollständigen Lizenztext.
