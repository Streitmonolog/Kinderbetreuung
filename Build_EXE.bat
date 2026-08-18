@echo off
setlocal
cd /d "%~dp0"

echo.
echo ==========================================
echo   Kinderbetreuung - Release EXE erstellen
echo ==========================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build\Build.ps1"

if errorlevel 1 (
    echo.
    echo FEHLER: Der Release-Build ist fehlgeschlagen.
    echo.
    pause
    exit /b 1
)

echo.
echo Build erfolgreich. Dateien liegen unter:
echo %~dp0Release
echo.
pause
