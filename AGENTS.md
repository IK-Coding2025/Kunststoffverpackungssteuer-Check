# Kunststoffverpackungssteuer – Code-Check

Interaktives Tool zur Ermittlung, ob Außenhandelscodes (KN-/WA-Nummern) oder
Produktionscodes (GP-Nummern) vom bislang untersuchten Entwurf des
KunststoffVStG erfasst sind.

## Lokal starten
```
streamlit run app.py
```
(Standard-Port 8501; bei Bedarf `--server.port 8520`)

## Veröffentlichung (analog zu PPWR-Rollen)
- GitHub-Repo anlegen (z. B. `IK-Coding2025/kunststoffverpackungssteuer-check`), Branch `main`
- Diesen Ordner (`Code/`) als Repo-Root pushen
- Auf https://share.streamlit.io „New app" wählen, Repo verbinden, `app.py` als Einstiegspunkt
- Öffentliche URL merken und in `.github/workflows/streamlit_keep_alive.yml` eintragen
  (Platzhalter: `https://kunststoffverpackungssteuer-check.streamlit.app`)

## Daten
- `build_data.py` erzeugt `codes_data.py` aus der Excel-Quelle
  `..\Input\261110_Kunststoffverpackungssteuer_Code-Übersicht und Dokumentation.xlsx`.
  Nach Änderungen an der Excel-Datei: `python build_data.py` erneut ausführen.
- `quellen/` enthält die amtlichen Originaldokumente (WA2026-PDF,
  GP2026-Güterverzeichnis, Gegenüberstellung Warennummern) zum Download in der App.

## Dateien
- `app.py` – komplette App (Suche, Code-Verzeichnisse, Dokumentation)
- `codes_data.py` – generierte Datengrundlage (nicht manuell bearbeiten)
- `logo.jpg` – IK-Logo (zentriert eingebettet)
- `.streamlit/config.toml` – Theme (primaryColor #004996)
- `requirements.txt` – `streamlit`, `pandas`, `openpyxl`
