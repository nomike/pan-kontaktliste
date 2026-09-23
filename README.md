# PAN Kontaktliste

Dieses Programm erstellt aus einer **SeaTable**-Anmeldeliste (PAN-Treffen) eine **Kontaktliste** für Teilnehmerinnen und Teilnehmer. Die Ausgabe ist eine **HTML-Datei**, die in jedem Webbrowser geöffnet werden kann. Eine PDF lässt sich direkt im Browser erzeugen (Drucken → Als PDF speichern). Es sind keine zusätzlichen Installationen wie LaTeX nötig – unter Windows, Linux und macOS reicht Python und ein Browser.

Es werden die Einwilligungen aus dem Anmeldeformular berücksichtigt: Nur wer der Teilnehmyliste zugestimmt hat, erscheint in der Liste; E-Mail, Telefon, Nachname, Vorname und Bild werden nur angezeigt, wenn die jeweilige Option gewählt wurde. Fehlt die Einwilligung für ein Bild, wird ein Platzhalterbild verwendet.

## Anforderungen

- **Python 3.10+**
- Ein **Webbrowser** (zum Anzeigen der HTML-Liste und zum Erzeugen einer PDF per Drucken → Als PDF speichern)
- Ein **SeaTable-Cloud**-Konto mit Zugriff auf die Anmelde-Bases (z. B. [cloud.seatable.io](https://cloud.seatable.io/))

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

### Grafische Oberfläche

```bash
python gui.py
```

1. Beim Start erscheinen **E-Mail/Benutzername** und **Passwort** für SeaTable. Die Zugangsdaten bleiben nur im Arbeitsspeicher und werden **nicht** gespeichert.
2. Nach der Anmeldung erscheint die Liste aller sichtbaren **Bases** (jedes Treffen ist eine eigene Base). Mit dem Filterfeld lässt sich die Liste eingrenzen.
3. Base auswählen, optional den **Namen des Treffens** anpassen, Zielpfad für die HTML-Datei wählen.
4. **Kontaktliste erstellen** lädt die Zeilen und Bilder per SeaTable-API und erzeugt die HTML-Datei.

Zum Erzeugen einer PDF: HTML im Browser öffnen → Menü Drucken (oder Strg+P) → „Als PDF speichern“ bzw. „Save as PDF“ wählen.

### Ablauf im Programm

1. Aus der gewählten Base werden nur Zeilen mit aktivierter **Teilnehmyliste** übernommen (bevorzugt aus einer View, deren Name „Kontaktliste“ enthält, sonst aus der gesamten Tabelle).
2. Pro Teilnehmer/in werden immer **Land**, **Rufname/Pseudonym** und **Teilnehmyliste_Couch** in die Liste übernommen.
3. **E-Mail**, **Telefonnummer**, **Nachname**, **Vorname** und **Bild** erscheinen nur, wenn die jeweilige Einwilligung gesetzt ist.
4. Bilder werden authentifiziert von SeaTable heruntergeladen; die EXIF-Ausrichtung wird automatisch korrigiert. Fehlt die Einwilligung oder kein Bild, wird `data/placeholder.png` verwendet.
5. Die Liste wird als eine einzige HTML-Datei mit eingebetteten Bildern (Data-URLs) erzeugt.

### Hinweis zum Free-Tarif

SeaTable Cloud Free begrenzt die API-Nutzung (aktuell ca. 3 000 Aufrufe pro Monat und Team). Das Erzeugen einer Kontaktliste verbraucht nur wenige Aufrufe plus einen pro Bild – für gelegentliche Treffen ist das in der Regel ausreichend. Bei HTTP 429 die Monatsquote in SeaTable prüfen.

## Versionierung und Releases

Die Version wird mit [Release Please](https://github.com/googleapis/release-please) verwaltet (Conventional Commits auf `main`). Beim Veröffentlichen eines Releases erstellt eine GitHub Action automatisch **Windows-Builds** (`.exe` und `.zip`) und hängt sie dem Release an.

- **Hilfe → Über PAN Kontaktliste** im Programm zeigt Version, Lizenz und Link zum [GitHub-Projekt](https://github.com/nomike/pan-kontaktliste).

## Entwicklung und Tests

```bash
pip install -r requirements-dev.txt
pytest
```

## Projektstruktur

- `gui.py` – grafische Oberfläche (Login, Base-Auswahl, HTML-Ausgabe)
- `seatable_reader.py` – SeaTable-Login, Base-Liste, Zeilen und Bilder laden
- `render.py` – Jinja2-Rendering der HTML-Vorlage (Bilder als Data-URLs)
- `template/contact_list.html.j2` – HTML-Vorlage (Jinja2) für die Kontaktliste
- `data/placeholder.png` – Platzhalterbild, wenn kein Bild oder keine Einwilligung
- `version.py` – Versionsanzeige (liest aus pyproject.toml)
- `requirements.txt` – Python-Abhängigkeiten
- `tests/` – Unit-Tests (pytest)

## Lizenz

Dieses Projekt steht unter der **GNU General Public License v3.0** (GPLv3). Siehe die Datei [LICENSE](LICENSE) für den vollständigen Lizenztext.
