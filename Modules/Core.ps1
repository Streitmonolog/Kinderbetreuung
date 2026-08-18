# Kinderbetreuung 1.0.17 - Core
# Gemeinsame Hilfsfunktionen. Keine Oberflaechenlogik.

function Show-Info([string]$Text) {
    [System.Windows.Forms.MessageBox]::Show($Text, "Kinderbetreuung", "OK", "Information") | Out-Null
}

function Show-Error([string]$Text) {
    [System.Windows.Forms.MessageBox]::Show($Text, "Kinderbetreuung", "OK", "Error") | Out-Null
}

function Write-StartError {
    param([System.Exception]$Exception)
    @"
Zeit: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Fehler: $($Exception.Message)

$($Exception.ToString())
"@ | Set-Content -Path $ErrorFile -Encoding UTF8
}

function Parse-DateDE([string]$Value) {
    $dt = New-Object DateTime
    if ([DateTime]::TryParseExact(
        $Value.Trim(),
        "dd.MM.yyyy",
        [Globalization.CultureInfo]::GetCultureInfo("de-DE"),
        [Globalization.DateTimeStyles]::None,
        [ref]$dt
    )) {
        return $dt
    }
    return $null
}

function Parse-DecimalDE([string]$Value) {
    $v = 0.0
    if ([string]::IsNullOrWhiteSpace($Value)) { return 0.0 }
    [double]::TryParse($Value.Replace(",","."), [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$v) | Out-Null
    return $v
}

