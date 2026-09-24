# PAN Kontaktliste

Dieses Programm erstellt aus einer **SeaTable**-Anmeldeliste (PAN-Treffen) eine **Kontaktliste** für Teilnehmerinnen und Teilnehmer. Die Ausgabe ist eine **PDF-Datei**. Es sind keine zusätzlichen Installationen wie LaTeX nötig – unter Windows, Linux und macOS reicht Python.

Es werden die Einwilligungen aus dem Anmeldeformular berücksichtigt: Nur wer der Teilnehmyliste zugestimmt hat, erscheint in der Liste; E-Mail, Telefon, Nachname, Vorname und Bild werden nur angezeigt, wenn die jeweilige Option gewählt wurde. Fehlt die Einwilligung für ein Bild, wird ein Platzhalterbild verwendet.

## Anforderungen

- **Python 3.10+**
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
3. Base auswählen, optional den **Namen des Treffens** anpassen, Zielpfad für die PDF-Datei wählen.
4. **Kontaktliste erstellen** lädt die Zeilen und Bilder per SeaTable-API und erzeugt die PDF-Datei.

### Ablauf im Programm

1. Aus der gewählten Base werden nur Zeilen mit aktivierter **Teilnehmyliste** übernommen (bevorzugt aus einer View, deren Name „Kontaktliste“ enthält, sonst aus der gesamten Tabelle).
2. Pro Teilnehmer/in werden immer **Land**, **Rufname/Pseudonym** und **Teilnehmyliste_Couch** in die Liste übernommen.
3. **E-Mail**, **Telefonnummer**, **Nachname**, **Vorname** und **Bild** erscheinen nur, wenn die jeweilige Einwilligung gesetzt ist.
4. Bilder werden authentifiziert von SeaTable heruntergeladen; die EXIF-Ausrichtung wird automatisch korrigiert. Fehlt die Einwilligung oder kein Bild, wird `data/placeholder.png` verwendet.
5. Die Liste wird als PDF mit eingebetteten Bildern erzeugt.

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

- `gui.py` – grafische Oberfläche (Login, Base-Auswahl, PDF-Ausgabe)
- `seatable_reader.py` – SeaTable-Login, Base-Liste, Zeilen und Bilder laden
- `render.py` – Jinja2-Rendering und PDF-Erzeugung via xhtml2pdf (Bilder als Data-URLs)
- `template/contact_list.html.j2` – HTML-Vorlage (Jinja2) für die Kontaktliste
- `data/placeholder.png` – Platzhalterbild, wenn kein Bild oder keine Einwilligung
- `data/app-icon.png` / `data/app-icon.ico` – Programm-Icon (Favicon von [polyamory.de](https://polyamory.de/))
- `version.py` – Versionsanzeige (liest aus pyproject.toml)
- `requirements.txt` – Python-Abhängigkeiten
- `tests/` – Unit-Tests (pytest)

## Lizenz

Dieses Projekt steht unter der **GNU General Public License v3.0** (GPLv3). Siehe die Datei [LICENSE](LICENSE) für den vollständigen Lizenztext.
