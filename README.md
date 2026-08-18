<div align="center">

<img src="Assets/Kinderbetreuung.png" alt="Kinderbetreuung" width="140">

# Kinderbetreuung

**Lokale Windows-Anwendung zur Dokumentation und Auswertung von Kinderbetreuungskosten**

[![Release](https://img.shields.io/github/v/release/Streitmonolog/Kinderbetreuung?display_name=tag&sort=semver)](https://github.com/Streitmonolog/Kinderbetreuung/releases/latest)
[![Build](https://github.com/Streitmonolog/Kinderbetreuung/actions/workflows/build.yml/badge.svg)](https://github.com/Streitmonolog/Kinderbetreuung/actions/workflows/build.yml)
[![Windows](https://img.shields.io/badge/Windows-10%20%7C%2011-0078D4?logo=windows)](https://github.com/Streitmonolog/Kinderbetreuung)
[![License](https://img.shields.io/github/license/Streitmonolog/Kinderbetreuung)](LICENSE)

[**⬇ Neueste Version herunterladen**](https://github.com/Streitmonolog/Kinderbetreuung/releases/latest)

</div>

---

## Was ist Kinderbetreuung?

**Kinderbetreuung** hilft dabei, Betreuungstage, Fahrtkostenerstattungen und entgeltliche Betreuung übersichtlich über ein Kalenderjahr zu dokumentieren.

Die Anwendung arbeitet vollständig lokal, verwaltet Betreuung nach Kalenderwochen und erzeugt auf Wunsch **XLSX- und PDF-Auswertungen** über Microsoft Excel oder LibreOffice.

> [!IMPORTANT]
> Kinderbetreuung ist ein **Dokumentationswerkzeug** und keine Steuer- oder Rechtsberatung. Ob einzelne Kosten steuerlich anerkannt werden, hängt vom jeweiligen Einzelfall ab. Weitere Hinweise: [DISCLAIMER.md](DISCLAIMER.md)

## Highlights

- 📅 Betreuung nach Kalenderwochen verwalten
- ☑️ mehrere Kalenderwochen per `STRG` / `SHIFT` bearbeiten
- 🏖️ Urlaubs- und Abwesenheitszeiträume berücksichtigen
- 🎉 NRW-Feiertage automatisch sperren
- 👨‍👩‍👧 mehrere Betreuungspersonen innerhalb eines Jahres
- 💶 unentgeltliche **und** entgeltliche Betreuung pro Termin
- ⏱️ Stunden und Stundenlohn für entgeltliche Betreuung
- 🚗 Fahrtkostenerstattung pro Termin ein-/ausschalten
- ✍️ Autovervollständigung bereits verwendeter Betreuungspersonen
- 📝 Zusatznotizen und private Dokumentationsfelder
- 📊 Quartalsübersicht Q1–Q4
- 📄 Steuer-Auswertung oder vollständige Dokumentation
- 📗 XLSX + PDF über **Microsoft Excel oder LibreOffice**
- 🔒 lokale Datenspeicherung ohne Cloud-Zwang
- 💾 Rückfrage bei ungespeicherten Änderungen

## Download & Installation

1. Öffne die [**neueste GitHub-Version**](https://github.com/Streitmonolog/Kinderbetreuung/releases/latest).
2. Lade `Kinderbetreuung_<Version>.exe` herunter.
3. Starte die EXE.

Eine Installation ist für die aktuelle portable Version nicht notwendig.

### Windows-Sicherheitshinweis

Die Anwendung ist derzeit noch **nicht kommerziell codesigniert**. Windows kann deshalb beim ersten Start einen SmartScreen-Hinweis anzeigen.

Zur Kontrolle wird zu jedem Release zusätzlich eine **SHA-256-Prüfsumme** veröffentlicht.

## Systemvoraussetzungen

| Komponente | Voraussetzung |
|---|---|
| Betriebssystem | Windows 10 oder Windows 11 |
| Anwendung | Windows PowerShell 5.1 |
| XLSX/PDF-Export | Microsoft Excel **oder** LibreOffice Calc |
| Python | **nicht erforderlich** |
| Cloudkonto | **nicht erforderlich** |

## Datenspeicherung & Datenschutz

Alle Anwendungsdaten werden lokal gespeichert:

```text
%LOCALAPPDATA%\Kinderbetreuung
```

Es gibt derzeit:

- keine Telemetrie,
- kein Tracking,
- keine automatische Cloud-Synchronisation,
- keine Übertragung von Betreuungsdaten an einen Projektserver.

Mehr dazu in [PRIVACY.md](PRIVACY.md).

> [!WARNING]
> Bitte in GitHub-Issues **keine echten Namen, IBANs, Betreuungsdaten oder andere personenbezogene Informationen** veröffentlichen.

## Export

Beim Export kann zwischen zwei Varianten gewählt werden:

**Nur Steuer**  
Enthält die für die steuerliche Dokumentation vorgesehenen Angaben der Anwendung.

**Vollständige Dokumentation**  
Enthält zusätzlich interne Notizen, Ausflüge und weitere Dokumentationsfelder.

Das Programm erkennt automatisch:

1. Microsoft Excel, oder
2. LibreOffice Calc.

Der Zielordner für den Export kann frei gewählt werden.

## Entwicklung

Repository klonen:

```powershell
git clone https://github.com/Streitmonolog/Kinderbetreuung.git
cd Kinderbetreuung
```

Entwicklerversion starten:

```text
Start_Entwicklung.bat
```

Projektstruktur:

```text
App.ps1
Modules/
  Core.ps1
  Calendar.ps1
  Data.ps1
  Export.ps1
  UI.ps1
Assets/
Build/
docs/
```

Technische Details: [ENTWICKLER.md](ENTWICKLER.md)

## Release-Build

Unter Windows:

```text
Build_EXE.bat
```

Der Build erzeugt:

```text
Release/
  Kinderbetreuung_<Version>.exe
  Kinderbetreuung_<Version>.exe.sha256
```

Die Versionsnummer wird zentral über [VERSION](VERSION) verwaltet.

## Automatische Builds

GitHub Actions übernimmt zwei Aufgaben:

| Workflow | Zweck |
|---|---|
| **Build** | Prüft Pushes und Pull Requests und erzeugt einen Windows-Build |
| **Release** | Erstellt bei einem Tag wie `v1.0.17` automatisch ein GitHub Release |

Dadurch wird jedes öffentliche Release reproduzierbar aus dem Repository erzeugt.

## Roadmap

### Als Nächstes

- [ ] anonymisierte Screenshots im README
- [ ] öffentliche Rückmeldungen und Bugreports sammeln
- [ ] zusätzliche automatisierte Tests
- [ ] Code Signing
- [ ] Installer / MSIX-Paket

### Langfristiges Ziel

🏪 **Veröffentlichung im Microsoft Store**

Die technische Vorbereitung dafür wird unter [docs/MICROSOFT_STORE.md](docs/MICROSOFT_STORE.md) dokumentiert.

## Fehler melden & mitwirken

Fehler und Verbesserungsvorschläge sind über [GitHub Issues](https://github.com/Streitmonolog/Kinderbetreuung/issues) willkommen.

Für Beiträge siehe [CONTRIBUTING.md](CONTRIBUTING.md).  
Sicherheitsrelevante Hinweise: [SECURITY.md](SECURITY.md).

## Lizenz

Dieses Projekt steht unter der [MIT-Lizenz](LICENSE).

Copyright © 2026 Lorenz Köcke
