# Release-Checkliste

## Vor dem Tag

- [ ] `VERSION` aktualisiert
- [ ] `CHANGELOG.md` aktualisiert
- [ ] lokaler Start getestet
- [ ] Speichern/Laden getestet
- [ ] Excel-Export getestet
- [ ] LibreOffice-Export getestet
- [ ] keine personenbezogenen Testdaten im Repository
- [ ] `git status` sauber
- [ ] GitHub-Build auf `main` grün

## Release

```powershell
git tag -a vX.Y.Z -m "Kinderbetreuung vX.Y.Z"
git push origin vX.Y.Z
```

Danach:

- [ ] Release-Workflow grün
- [ ] EXE vorhanden
- [ ] SHA-256-Datei vorhanden
- [ ] EXE vom GitHub-Release herunterladen und testen
- [ ] Release als `Latest` sichtbar

## Nach dem Release

- [ ] GitHub-README prüfen
- [ ] Screenshot bei UI-Änderungen aktualisieren
- [ ] bekannte Probleme dokumentieren
