# Expose Processor — Immobilien-Projektierungs-Automation

Liest ein PDF-Exposé (Baugrundstück), prüft Business-Rules,
matched Haustypen und erstellt befüllte `.docx`-Projektierungsdokumente.
Optional: automatischer Google Drive-Upload und E-Mail-Benachrichtigung.

---

## Dateistruktur

```
expose_processor/
├── expose_processor.py          # Hauptscript
├── config.json                  # Konfiguration (API-Keys, Pfade)
├── requirements.txt
├── assets/
│   ├── google_service_account.json   # Google-Credentials (selbst ablegen)
│   └── templates/
│       └── Projektierung_Template.docx   # Word-Template (selbst ablegen)
├── output/
│   └── Projektierungen/         # Generierte DOCX-Dateien
├── system/
│   └── logs/
│       ├── stop_log.csv         # Protokoll gefilterter Exposés
│       └── error_log.txt        # Fehlerprotokoll
└── README.md
```

---

## Installation

### 1. Python-Version

Python **3.10 oder neuer** wird benötigt.

```bash
python --version
```

### 2. Abhängigkeiten installieren

```bash
cd expose_processor
pip install -r requirements.txt
```

---

## Konfiguration

### config.json befüllen

Öffne `config.json` und trage alle Werte ein:

| Feld | Beschreibung |
|------|-------------|
| `anthropic_api_key` | API-Key von [console.anthropic.com](https://console.anthropic.com) |
| `mitarbeiter_nr` | Deine interne Mitarbeiternummer |
| `mitarbeiter_name` | Dein Name (erscheint im Dokument) |
| `email_absender` | Gmail-Adresse des Absenders |
| `email_empfaenger` | Empfänger-Adresse für Benachrichtigungen |
| `email_passwort` | Gmail App-Passwort (nicht das normale Passwort!) |
| `google_drive_hauptordner_id` | Ordner-ID aus dem Drive-URL |
| `google_credentials_pfad` | Pfad zur Service-Account JSON-Datei |

### Gmail App-Passwort erstellen

1. Google-Konto → Sicherheit → 2-Faktor-Authentifizierung aktivieren
2. Sicherheit → App-Passwörter → Neues App-Passwort erstellen
3. Das generierte 16-stellige Passwort in `email_passwort` eintragen

### Google Drive einrichten

1. [Google Cloud Console](https://console.cloud.google.com) → Neues Projekt
2. APIs aktivieren: **Google Drive API**
3. IAM → Dienstkonto erstellen → JSON-Schlüssel herunterladen
4. JSON-Datei als `assets/google_service_account.json` ablegen
5. In Google Drive: den Hauptordner "Projektierungen" mit dem Dienstkonto teilen
6. Ordner-ID aus der URL kopieren (letzter Teil nach `/folders/`) → in `config.json`

### Word-Template anlegen

Lege das Template unter `assets/templates/Projektierung_Template.docx` ab.
Das Template muss folgende Platzhalter enthalten (exakt so geschrieben):

| Platzhalter | Inhalt |
|-------------|--------|
| `{{MITARBEITER_NR}}` | Mitarbeiternummer |
| `{{MITARBEITER_NAME}}` | Mitarbeitername |
| `{{HAUSTYP}}` | Modellname (z. B. "Stadtvilla 136") |
| `{{UEBERSCHRIFT}}` | KI-generierte Portal-Überschrift |
| `{{STRASSE}}` | Straße + Hausnummer |
| `{{REGION}}` | Bundesland/Region |
| `{{PLZ_ORT}}` | PLZ + Ortsname |
| `{{BESCHREIBUNG_ORT}}` | KI-generierter Beschreibungstext |
| `{{GRUNDSTUECK_M2}}` | Grundstücksfläche in m² |
| `{{KAUFPREIS}}` | Kaufpreis formatiert (z. B. "295.000 €") |
| `{{COURTAGE}}` | z. B. "7,14 % inkl. MwSt." oder "courtage-frei" |
| `{{ERSCHLOSSEN}}` | "ja" / "teilerschlossen" / "nein" |
| `{{BESONDERHEITEN}}` | Besondere Hinweise aus dem Exposé |

---

## Ausführung

```bash
python expose_processor.py --pdf pfad/zum/expose.pdf
```

**Beispiel:**

```bash
python expose_processor.py --pdf ~/Downloads/Expose_Berlin_Musterstr12.pdf
```

---

## Ablauf (Übersicht)

```
PDF lesen
    ↓
Claude API: Daten extrahieren (JSON)
    ↓
Business-Rules prüfen (5 Gates):
  Gate 1: Kaufpreis ≤ 350.000 €
  Gate 2: Fläche 400–1.000 m²
  Gate 3: Bebaubare Fläche ≥ 50 m²
  Gate 4: Kein Abriss, keine Gartenlaube
  Gate 5: Mindestens teilerschlossen
    ↓ (bei STOP → CSV-Log + E-Mail)
Haustyp-Matching nach Geschosszahl + bebaubarer Fläche
    ↓ (ggf. zusätzlich Doppelhaus-Matching)
Claude API: Marketing-Text pro Modell
    ↓
Word-Template befüllen → .docx speichern
    ↓
Google Drive Upload
    ↓
E-Mail-Benachrichtigung
```

---

## Haustypen-Katalog

| Modell | Typ | Geschosse | Bebaute Fläche |
|--------|-----|-----------|---------------|
| Winkelbungalow 120 | Bungalow | 1 | 144,35 m² |
| B92 Eco | Bungalow | 1 | 112,72 m² |
| Klassik 112 | Klassikhaus | 1,5 | 73,34 m² |
| Stadtvilla 121 | Stadtvilla | 2 | 78,55 m² |
| Stadtvilla 136 | Stadtvilla | 2 | 85,69 m² |

---

## Logs

- **`system/logs/stop_log.csv`** — jedes gestoppte Exposé mit Datum, Dateiname, Gate und Grund
- **`system/logs/error_log.txt`** — technische Fehler mit Zeitstempel

---

## Fehlerbehebung

| Problem | Lösung |
|---------|--------|
| `anthropic_api_key fehlt` | API-Key in `config.json` eintragen |
| `Template nicht gefunden` | `.docx`-Template unter `assets/templates/` ablegen |
| `E-Mail-Versand schlägt fehl` | Gmail App-Passwort prüfen, kein normales Passwort |
| `Drive Upload schlägt fehl` | Service-Account mit Ordner geteilt? Credentials-Pfad korrekt? |
| `JSON-Parse-Fehler (3x)` | Exposé-Text zu kurz oder unlesbar — PDF manuell prüfen |
