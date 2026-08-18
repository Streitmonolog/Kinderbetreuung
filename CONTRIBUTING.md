# Mitwirken

Beiträge sind willkommen.

## Entwicklung

1. Repository klonen.
2. `Start_Entwicklung.bat` starten.
3. Änderungen möglichst auf einen Themenbereich begrenzen.
4. Vor einem Pull Request mindestens Start, Speichern/Laden und Export testen.

## Projektstruktur

- `App.ps1` – Einstiegspunkt
- `Modules/Core.ps1` – allgemeine Hilfsfunktionen
- `Modules/Calendar.ps1` – Kalender, Urlaub und Feiertage
- `Modules/Data.ps1` – Datenmodell, Speichern/Laden und Summen
- `Modules/Export.ps1` – XLSX/PDF über Excel oder LibreOffice
- `Modules/UI.ps1` – WinForms-Oberfläche
- `Build/` – EXE-Build

## Pull Requests

Bitte beschreibe:

- was geändert wurde
- warum die Änderung nötig ist
- wie sie getestet wurde
- ob sich Datenformat oder Export verändern

## Steuerliche Aussagen

Bitte keine steuerlichen Regeln als gesicherte Tatsache in Code oder Dokumentation
einbauen, ohne sie anhand einer aktuellen offiziellen Quelle zu prüfen.
