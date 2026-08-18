# Entwicklerübersicht – Kinderbetreuung 1.0.17

Die Anwendung ist bewusst modular aufgebaut. `App.ps1` enthält keine Fachlogik, sondern initialisiert nur Pfade und Zustand und lädt anschließend die Module.

## Module

**Core.ps1** enthält kleine, allgemeine Hilfsfunktionen. Hier gehören nur Funktionen hinein, die von mehreren Bereichen verwendet werden.

**Calendar.ps1** ist für Kalenderlogik zuständig: Jahresaufbau, NRW-Feiertage, Urlaub und die Prüfung, ob ein Betreuungstag zulässig ist.

**Data.ps1** verwaltet den fachlichen Zustand: Termine, Betreuungspersonen, Personenprofile/Autocomplete, Speichern/Laden sowie Summen.

**Export.ps1** enthält ausschließlich die Ausgabe nach XLSX/PDF und die automatische Auswahl zwischen Microsoft Excel und LibreOffice.

**UI.ps1** baut das WinForms-Fenster auf und verbindet Controls mit den Funktionen der anderen Module.

## Daten

Nutzerdaten liegen nicht im Programmordner, sondern unter `%LOCALAPPDATA%\Kinderbetreuung`. Ein Update der EXE überschreibt deshalb keine Jahresdaten.

## EXE

`Build_EXE.bat` verwendet den auf Windows vorhandenen C#-Compiler aus .NET Framework. `Launcher.cs` und `Payload.zip` werden zu einer einzelnen Windows-GUI-EXE kompiliert. Das Icon wird beim Kompilieren direkt in die EXE eingebettet.
