@echo off
setlocal
cd /d "%~dp0"

echo.
echo ==========================================
echo   Kinderbetreuung - Release Build
echo ==========================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build\Build.ps1"
if errorlevel 1 (
    echo.
    echo FEHLER: Die EXE konnte nicht erstellt werden.
    echo.
    pause
    exit /b 1
)

echo.
echo Fertig. Die EXE und SHA256-Pruefsumme liegen im Ordner Release.
echo.
pause
