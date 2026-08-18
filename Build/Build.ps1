param(
    [switch]$CI
)

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $PSScriptRoot
$Version = (Get-Content (Join-Path $Root 'VERSION') -Raw).Trim()

if ($Version -notmatch '^\d+\.\d+\.\d+$') {
    throw "Ungueltige Versionsnummer in VERSION: $Version"
}

$ReleaseDir = Join-Path $Root 'Release'
$Payload = Join-Path $PSScriptRoot 'Payload.zip'
$GeneratedLauncher = Join-Path $PSScriptRoot 'Launcher.generated.cs'
$Exe = Join-Path $ReleaseDir ("Kinderbetreuung_{0}.exe" -f $Version)
$HashFile = $Exe + '.sha256'

New-Item -ItemType Directory -Path $ReleaseDir -Force | Out-Null
Remove-Item $Payload,$GeneratedLauncher,$Exe,$HashFile -Force -ErrorAction SilentlyContinue

# Payload aus den echten Projektquellen erzeugen. Keine generierten oder lokalen Daten einbetten.
$payloadItems = @(
    @{ Source = (Join-Path $Root 'App.ps1'); Target = 'App.ps1' },
    @{ Source = (Join-Path $Root 'Assets\Kinderbetreuung.ico'); Target = 'Assets/Kinderbetreuung.ico' },
    @{ Source = (Join-Path $Root 'Assets\Kinderbetreuung.png'); Target = 'Assets/Kinderbetreuung.png' }
)

Get-ChildItem (Join-Path $Root 'Modules') -Filter '*.ps1' | Sort-Object Name | ForEach-Object {
    $payloadItems += @{ Source = $_.FullName; Target = ('Modules/' + $_.Name) }
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::Open($Payload,[System.IO.Compression.ZipArchiveMode]::Create)
try {
    foreach ($item in $payloadItems) {
        if (-not (Test-Path $item.Source)) { throw "Fehlende Payload-Datei: $($item.Source)" }
        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
            $zip,
            $item.Source,
            $item.Target,
            [System.IO.Compression.CompressionLevel]::Optimal
        ) | Out-Null
    }
}
finally {
    $zip.Dispose()
}

# Launcher aus Vorlage mit zentraler Versionsnummer erzeugen.
$launcher = Get-Content (Join-Path $PSScriptRoot 'Launcher.cs') -Raw
$launcher = $launcher.Replace('__VERSION__',$Version)
Set-Content -Path $GeneratedLauncher -Value $launcher -Encoding UTF8

$candidates = @(
    "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
    "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe"
)
$Csc = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $Csc) { throw 'Microsoft C# Compiler csc.exe wurde nicht gefunden.' }

$args = @(
    '/nologo',
    '/target:winexe',
    '/platform:anycpu',
    '/optimize+',
    ('/win32icon:"{0}"' -f (Join-Path $Root 'Assets\Kinderbetreuung.ico')),
    ('/resource:"{0}",Kinderbetreuung.Payload.zip' -f $Payload),
    '/reference:System.Windows.Forms.dll',
    '/reference:System.IO.Compression.dll',
    '/reference:System.IO.Compression.FileSystem.dll',
    ('/out:"{0}"' -f $Exe),
    ('"{0}"' -f $GeneratedLauncher)
)

$proc = Start-Process -FilePath $Csc -ArgumentList $args -NoNewWindow -Wait -PassThru
if ($proc.ExitCode -ne 0 -or -not (Test-Path $Exe)) {
    throw "EXE-Build fehlgeschlagen (ExitCode $($proc.ExitCode))."
}

$hash = (Get-FileHash $Exe -Algorithm SHA256).Hash.ToLowerInvariant()
Set-Content -Path $HashFile -Value ("{0}  {1}" -f $hash,[IO.Path]::GetFileName($Exe)) -Encoding ASCII

Write-Host ''
Write-Host 'Build erfolgreich:' -ForegroundColor Green
Write-Host $Exe
Write-Host ('SHA256: ' + $hash)

if (-not $CI) {
    Write-Host ''
    Write-Host 'Die Release-Dateien liegen unter .\Release\'
}
