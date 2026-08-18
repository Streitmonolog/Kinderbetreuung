# Changelog

## [1.0.18] - 2026-08-18

### Changed
- In-Process-PowerShell-Runspace explizit auf STA gestellt, damit WinForms-Autocomplete und OLE-basierte UI-Funktionen funktionieren.
- Release-Launcher vollständig überarbeitet.
- Keine selbstentpackende Payload mehr.
- Kein versteckter `powershell.exe`-Unterprozess mehr.
- Kein `ExecutionPolicy Bypass` beim Start der veröffentlichten EXE.
- Windows-PowerShell-Engine wird direkt im EXE-Prozess verwendet.
- Release-Build erzeugt weiterhin EXE und SHA-256-Prüfsumme.

### Fixed
- Autocomplete-Personenprofile übernehmen den zuletzt bekannten Stundenlohn wieder zuverlässig.

### Security
- Build-Architektur geändert, um unnötig verdächtige Verhaltensmuster für AV/EDR-Produkte zu vermeiden.

Alle nennenswerten Änderungen werden hier dokumentiert.

## [1.0.17] - 2026-08-18

### Hinzugefügt
- moderne, am Programmicon orientierte Oberfläche
- modularer PowerShell-Aufbau
- EXE-Launcher mit eingebettetem Programmpaket
- Excel- und LibreOffice-Unterstützung für XLSX/PDF
- frei wählbarer Exportordner
- unentgeltliche und entgeltliche Betreuung pro Termin
- Betreuungsperson, Stunden, Stundenlohn und Fahrtkosten pro Termin
- Personen-Autovervollständigung
- Quartalsübersicht und steuerbezogene Auswertung
- NRW-Feiertage und Urlaubs-/Abwesenheitszeiträume
- Speicherabfrage beim Beenden
- lokale JSON-Datenspeicherung und Sicherungsdatei

### Datenschutz
- Daten werden lokal unter `%LOCALAPPDATA%\Kinderbetreuung` gespeichert
- keine Telemetrie integriert
