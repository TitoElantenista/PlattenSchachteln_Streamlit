# Plattenschachtelung Streamlit

Web-App zur Plattenschachtelung aus einer CSV-Datei. Die App erklärt den benötigten CSV-Aufbau, prüft die hochgeladene Datei, zeigt gefundene Blechstärken und Güten und erzeugt die Ergebnisdateien als ZIP-Download.

Während der Berechnung zeigt die App eine Fortschrittsleiste, den aktuell bearbeiteten Satz und eine grobe Restzeit-Schätzung.

## Start

```bash
python -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/streamlit run app.py
```

Für den PDF-Export muss lokal Chrome oder Chromium installiert sein. Auf Streamlit Community Cloud wird Chromium über `packages.txt` installiert.

## Deployment auf Streamlit Community Cloud

Dieses Repository ist fuer Streamlit Community Cloud vorbereitet.

1. Repository auf GitHub erstellen und diese Dateien hochladen.
2. In Streamlit Community Cloud eine neue App aus dem GitHub-Repository anlegen.
3. Als Main file path `app.py` eintragen.
4. In den Advanced settings Python `3.12` auswählen.
5. Streamlit installiert automatisch `requirements.txt`.

Wichtige Dateien:

- `app.py`: Streamlit-Web-App.
- `SchachtelnScript.py`: Schachtellogik und Ergebnisexport.
- `requirements.txt`: Python-Abhängigkeiten fuer Streamlit Cloud.
- `packages.txt`: Systemabhängigkeiten fuer Streamlit Cloud, aktuell Chromium fuer den PDF-Export.
- `.streamlit/config.toml`: Streamlit-Konfiguration.

CSV-Dateien werden standardmäßig per `.gitignore` ignoriert, damit keine Projekt- oder Produktionsdaten ins Repository committed werden.

## CSV-Aufbau

Die CSV-Datei muss Semikolon als Trennzeichen verwenden. Folgende Spalten werden benötigt:

- `lfd-Nr.`: Positionsnummer oder laufende Nummer.
- `Anzahl`: Stückzahl der Position.
- `Profil`: Blechstärke, zum Beispiel `BL10`.
- `Güte`: Materialgüte, zum Beispiel `S355J2`.
- `Länge`: Länge in Millimeter.
- `Breite`: Breite in Millimeter.

Nur Profile, die mit `BL` beginnen, werden geschachtelt. Die Materialgüte wird unverändert aus der CSV übernommen.

## Eingaben in der App

- Projektname: erscheint in Ergebnisdateien und ZIP-Dateiname.
- Tafeltypen: Name, Breite und Länge in Millimeter.
- Rand: Sicherheitsrand vom Tafelrand bis zur ersten Kontur in Millimeter.
- Abstand: Mindestabstand zwischen zwei Teilen in Millimeter.
- Mindeststärke: BL-Teile unter dieser Stärke werden separat gelistet.
- Stahlgewicht: Rechenwert für Gewichtsangaben in kg/m²/mm.

## Erweiterte Einstellungen

In der Tabelle `Tafeltypen` können beliebig viele verfügbare Tafelmaße gepflegt werden. Nach dem CSV-Upload kann optional der Bereich `Erweiterte Tafelzuordnung` geöffnet werden.

Dort gilt standardmäßig jede Tafel für alle Blechgruppen. Pro Tafeltyp sind alle Gruppen `BL / Güte` zuerst ausgewählt; die Auswahl zeigt zusätzlich Stückzahl und Fläche. Mit `Alle` und `Keine` kann eine Tafel schnell vollständig ein- oder ausgeschlossen werden.

Diese Einstellungen sind optional. Ohne Änderung verwendet jede Gruppe alle definierten Tafeltypen.

## Ergebnis

Der ZIP-Download enthält:

- `Schachtel Ergebnis.md`
- `Schachtel Ergebnis.csv`
- `Platten unter Mindeststaerke.csv`
- `Materialliste Bestellung.xlsx`
- `Schachtel Kontrolle.html`
- `Schachtel Kontrolle.pdf`
- `Schachtel Bilder/`
- `Konfiguration.txt`
