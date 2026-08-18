@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0App.ps1"
if errorlevel 1 (
  echo.
  echo Startfehler. Details stehen unter:
  echo %%LOCALAPPDATA%%\Kinderbetreuung\startfehler.txt
  echo.
  pause
)
