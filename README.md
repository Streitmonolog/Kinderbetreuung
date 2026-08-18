# Kinderbetreuung

![Kinderbetreuung Icon](Assets/Kinderbetreuung.png)

**Kinderbetreuung** ist eine lokale Windows-Anwendung zur Dokumentation von
Betreuungstagen, Fahrtkostenerstattungen und entgeltlicher Kinderbetreuung.
Sie erzeugt übersichtliche XLSX- und PDF-Auswertungen und unterstützt sowohl
Microsoft Excel als auch LibreOffice.

> **Hinweis:** Das Programm ist ein Dokumentationswerkzeug und ersetzt keine
> Steuer- oder Rechtsberatung. Siehe [DISCLAIMER.md](DISCLAIMER.md).

## Funktionen

- Betreuungstage komfortabel nach Kalenderwochen verwalten
- mehrere Kalenderwochen per STRG/SHIFT bearbeiten
- NRW-Feiertage und eigene Urlaubs-/Abwesenheitszeiträume berücksichtigen
- unentgeltliche und entgeltliche Betreuung pro Termin
- unterschiedliche Betreuungspersonen innerhalb eines Jahres
- Stunden und Stundenlohn bei entgeltlicher Betreuung
- Fahrtkostenerstattung pro Termin ein-/ausschalten
- Autovervollständigung bereits verwendeter Betreuungspersonen
- Zusatznotizen und private Dokumentationsfelder
- Quartalsübersicht Q1–Q4
- XLSX- und PDF-Export
- Microsoft Excel **oder** LibreOffice als Export-Backend
- lokale Datenspeicherung ohne Cloud-Zwang
- Speicherabfrage beim Beenden

## Datenschutz

Die Anwendung speichert ihre Daten lokal unter:

```text
%LOCALAPPDATA%\Kinderbetreuung
```

Es sind keine Telemetrie und keine automatische Übertragung der erfassten
Betreuungsdaten an einen Projektserver eingebaut. Weitere Informationen stehen in
[PRIVACY.md](PRIVACY.md).

## Systemvoraussetzungen

- Windows 10 oder Windows 11
- Windows PowerShell 5.1 für die aktuelle Anwendungsgeneration
- für XLSX/PDF-Export entweder:
  - Microsoft Excel oder
  - LibreOffice Calc

Für die normale Nutzung wird **kein Python** benötigt.

## Download

Für Endnutzer sind die fertigen Builds unter **GitHub Releases** vorgesehen.
Die EXE-Datei wird zusammen mit einer SHA-256-Prüfsumme veröffentlicht.

## Entwicklung

Zum lokalen Start aus dem Quellcode:

```text
Start_Entwicklung.bat
```

Die Anwendung ist modular aufgebaut:

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
```

Weitere technische Informationen: [ENTWICKLER.md](ENTWICKLER.md)

## Release-Build erstellen

Unter Windows:

```text
Build_EXE.bat
```

Der Build erzeugt anschließend unter `Release/`:

```text
Kinderbetreuung_<Version>.exe
Kinderbetreuung_<Version>.exe.sha256
```

Der Build wird aus den aktuellen Quelldateien erzeugt. Das eingebettete
`Payload.zip` ist deshalb kein Bestandteil des Repositorys, sondern ein
Build-Artefakt.

## GitHub Actions

Das Repository enthält zwei Workflows:

- **Build** – baut bei Push/Pull Request eine Windows-EXE als Artefakt
- **Release** – baut bei einem Tag wie `v1.0.18` und erstellt daraus ein GitHub Release

Die Versionsnummer liegt zentral in der Datei [VERSION](VERSION).

## Roadmap

### Kurzfristig

- erste öffentliche GitHub-Version
- anonymisierte Screenshots
- Bugreports und Feedback über GitHub Issues
- Release-Prozess stabilisieren

### Mittelfristig

- Code Signing
- Installer bzw. MSIX-Paket
- automatisierte Release-Tests
- Update-Konzept

### Langfristig

- Veröffentlichung im **Microsoft Store**

Details: [docs/MICROSOFT_STORE.md](docs/MICROSOFT_STORE.md)

## Mitwirken

Siehe [CONTRIBUTING.md](CONTRIBUTING.md).

Bei Fehlerberichten bitte **keine echten Namen, IBANs oder andere personenbezogene
Daten** aus der lokalen Datendatei veröffentlichen.

## Lizenz

Dieses Projekt steht unter der [MIT-Lizenz](LICENSE).

Copyright © 2026 Lorenz Köcke


> Build-Fix: `System.Management.Automation.dll` wird automatisch aus der laufenden Windows-PowerShell ermittelt.

> Build-Fix 2: Der eingebettete PowerShell-Runspace läuft explizit im STA-Modus, damit WinForms-Autovervollständigung und Dialoge vollständig funktionieren.
