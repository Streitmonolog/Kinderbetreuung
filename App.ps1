# Kinderbetreuung 1.0.17
# Schlanker Programmeinstieg. Fachlogik liegt unter .\Modules.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"
$AppDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Persistente Nutzerdaten bleiben bei Updates unangetastet.
$UserDataDir = Join-Path $env:LOCALAPPDATA "Kinderbetreuung"
if (-not (Test-Path $UserDataDir)) {
    New-Item -ItemType Directory -Path $UserDataDir -Force | Out-Null
}

$DataFile = Join-Path $UserDataDir "kinderbetreuung-daten.json"
$ErrorFile = Join-Path $UserDataDir "startfehler.txt"

# Daten aus alten portablen Versionen einmalig uebernehmen.
$LegacyDataFile = Join-Path $AppDir "kinderbetreuung-daten.json"
if ((-not (Test-Path $DataFile)) -and (Test-Path $LegacyDataFile)) {
    try { Copy-Item $LegacyDataFile $DataFile -Force } catch {}
}

# Gemeinsamer Laufzeitzustand.
$script:Weeks = @()
$script:CurrentIndex = -1
$script:LoadingDetail = $false
$script:DemoMode = $false
$script:DemoTarget = 0.0
$script:IsDirty = $false
$script:ClosingConfirmed = $false


$modules = @(
    "Core.ps1",
    "Calendar.ps1",
    "Data.ps1",
    "Export.ps1",
    "UI.ps1"
)

foreach ($module in $modules) {
    $path = Join-Path $AppDir ("Modules\" + $module)
    if (-not (Test-Path $path)) {
        throw "Programmkomponente fehlt: $module"
    }
    . $path
}
