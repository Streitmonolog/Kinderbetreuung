param(
    [string]$VersionFile = (Join-Path $PSScriptRoot "..\VERSION")
)

$ErrorActionPreference = "Stop"

$Root = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$Version = (Get-Content $VersionFile -Raw).Trim()

if ([string]::IsNullOrWhiteSpace($Version)) {
    throw "VERSION ist leer."
}

$ReleaseDir = Join-Path $Root "Release"
New-Item -ItemType Directory -Path $ReleaseDir -Force | Out-Null

$GeneratedDir = Join-Path $PSScriptRoot "Generated"
New-Item -ItemType Directory -Path $GeneratedDir -Force | Out-Null

$EmbeddedScript = Join-Path $GeneratedDir "EmbeddedApp.ps1"

# App-Bootstrap ohne externen Modul-Loader.
$app = Get-Content (Join-Path $Root "App.ps1") -Raw -Encoding UTF8

# Der eingebettete Modus wird vom C#-Launcher vor dem Skript gesetzt.
# App.ps1 initialisiert Pfade/Zustand; die Module werden danach inline angehaengt.
$parts = New-Object System.Collections.Generic.List[string]
$parts.Add($app)

foreach ($name in @("Core.ps1","Calendar.ps1","Data.ps1","Export.ps1","UI.ps1")) {
    $path = Join-Path $Root ("Modules\" + $name)
    if (-not (Test-Path $path)) {
        throw "Fehlendes Modul: $path"
    }
    $parts.Add("`r`n# ===== BEGIN MODULE: $name =====`r`n")
    $parts.Add((Get-Content $path -Raw -Encoding UTF8))
    $parts.Add("`r`n# ===== END MODULE: $name =====`r`n")
}

[System.IO.File]::WriteAllText(
    $EmbeddedScript,
    ($parts -join "`r`n"),
    (New-Object System.Text.UTF8Encoding($true))
)

# Windows PowerShell Automation Assembly
# Nicht auf einen festen Dateipfad vertrauen: je nach Windows-Version liegt
# System.Management.Automation.dll im GAC bzw. in einem anderen Framework-Pfad.
try {
    $AutomationDll = [System.Management.Automation.PowerShell].Assembly.Location
}
catch {
    throw "Die Windows-PowerShell-Automation-Assembly konnte nicht ermittelt werden."
}

if ([string]::IsNullOrWhiteSpace($AutomationDll) -or -not (Test-Path $AutomationDll)) {
    throw "System.Management.Automation.dll wurde nicht gefunden. Ermittelter Pfad: $AutomationDll"
}

Write-Host "PowerShell Automation:"
Write-Host "  $AutomationDll"

# Klassischer .NET Framework C# Compiler
$CscCandidates = @(
    "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
    "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe"
)
$Csc = $CscCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $Csc) {
    throw "csc.exe wurde nicht gefunden."
}

$Exe = Join-Path $ReleaseDir ("Kinderbetreuung_{0}.exe" -f $Version)
$Icon = Join-Path $Root "Assets\Kinderbetreuung.ico"
$Launcher = Join-Path $PSScriptRoot "Launcher.cs"

$args = @(
    "/nologo",
    "/target:winexe",
    "/optimize+",
    "/platform:anycpu",
    "/out:$Exe",
    "/win32icon:$Icon",
    "/reference:System.dll",
    "/reference:System.Core.dll",
    "/reference:System.Windows.Forms.dll",
    "/reference:System.Drawing.dll",
    "/reference:$AutomationDll",
    "/resource:$EmbeddedScript,Kinderbetreuung.EmbeddedApp.ps1",
    $Launcher
)

Write-Host ""
Write-Host "Kinderbetreuung $Version - Release Build"
Write-Host "-----------------------------------------"
Write-Host "Launcher: In-Process Windows PowerShell"
Write-Host "Keine Laufzeit-Extraktion, kein powershell.exe-Unterprozess."
Write-Host ""

& $Csc @args
if ($LASTEXITCODE -ne 0) {
    throw "C#-Build fehlgeschlagen (ExitCode $LASTEXITCODE)."
}

if (-not (Test-Path $Exe)) {
    throw "EXE wurde nicht erzeugt."
}

$Hash = (Get-FileHash $Exe -Algorithm SHA256).Hash.ToLowerInvariant()
$HashFile = "$Exe.sha256"
"$Hash  $([System.IO.Path]::GetFileName($Exe))" | Set-Content $HashFile -Encoding ASCII

Write-Host "Erstellt:"
Write-Host "  $Exe"
Write-Host "  $HashFile"
