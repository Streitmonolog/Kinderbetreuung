# Kinderbetreuung 1.0.17 - Kalender
# Jahresaufbau, NRW-Feiertage, Urlaub und Datumspruefungen.

function Get-IsoWeek([datetime]$Date) {
    # ISO 8601: Donnerstag bestimmt das ISO-Jahr / die Kalenderwoche.
    $day = [int]$Date.DayOfWeek
    if ($day -eq 0) { $day = 7 }
    $thursday = $Date.AddDays(4 - $day)
    $jan4 = [datetime]::new($thursday.Year,1,4)
    $jan4Day = [int]$jan4.DayOfWeek
    if ($jan4Day -eq 0) { $jan4Day = 7 }
    $firstThursday = $jan4.AddDays(4 - $jan4Day)
    return 1 + [int][math]::Floor(($thursday.Date - $firstThursday.Date).TotalDays / 7)
}

function Get-MondayOfIsoWeek([int]$Year, [int]$Week) {
    $jan4 = [datetime]::new($Year,1,4)
    $day = [int]$jan4.DayOfWeek
    if ($day -eq 0) { $day = 7 }
    $mondayWeek1 = $jan4.AddDays(1 - $day)
    return $mondayWeek1.AddDays(($Week - 1) * 7)
}

function Get-IsoWeeksInYear([int]$Year) {
    # 28. Dezember liegt nach ISO 8601 immer in der letzten KW des Jahres.
    return Get-IsoWeek ([datetime]::new($Year,12,28))
}

$script:Weeks = @()
$script:CurrentIndex = -1
$script:LoadingDetail = $false
$script:DemoMode = $false
$script:DemoTarget = 0.0

function New-YearData {
    $year = [int]$numYear.Value

    $jan1 = [datetime]::new($year,1,1)
    $dow = [int]$jan1.DayOfWeek
    if ($dow -eq 0) { $dow = 7 }
    $firstMonday = $jan1.AddDays(1 - $dow)

    $dec31 = [datetime]::new($year,12,31)
    $dowEnd = [int]$dec31.DayOfWeek
    if ($dowEnd -eq 0) { $dowEnd = 7 }
    $lastMonday = $dec31.AddDays(1 - $dowEnd)

    $script:Weeks = @()
    $displayWeek = 1
    $mon = $firstMonday

    while ($mon -le $lastMonday) {
        $days = @()
        for ($i=0; $i -lt 5; $i++) {
            $d = $mon.AddDays($i)
            $days += [pscustomobject]@{
                Date = $d.ToString("yyyy-MM-dd")
                Checked = $false
                Kosten = "0,00"
                Kategorie = ""
                Notiz = ""
                AusflugZiel = ""
                AusflugKm = "0,00"
                Eintritt = "0,00"
                Person = ""
                Art = "unentgeltlich"
                Stunden = "0,00"
                Stundenlohn = "0,00"
                Fahrtkosten = $true
            }
        }
        $script:Weeks += [pscustomobject]@{
            Week = $displayWeek
            Monday = $mon.ToString("yyyy-MM-dd")
            Days = $days
        }

        $displayWeek++
        $mon = $mon.AddDays(7)
    }

    Refresh-WeekList
    if ($grid.Rows.Count -gt 0) {
        $grid.Rows[0].Selected = $true
        $grid.CurrentCell = $grid.Rows[0].Cells[0]
        Load-WeekDetail 0
    }
    $script:IsDirty = $true
}

function Get-EasterSunday([int]$Year) {
    # Gregorianischer Osteralgorithmus nach Meeus/Jones/Butcher.
    $a = $Year % 19
    $b = [math]::Floor($Year / 100)
    $c = $Year % 100
    $d = [math]::Floor($b / 4)
    $e = $b % 4
    $f = [math]::Floor(($b + 8) / 25)
    $g = [math]::Floor(($b - $f + 1) / 3)
    $h = (19 * $a + $b - $d - $g + 15) % 30
    $i = [math]::Floor($c / 4)
    $k = $c % 4
    $l = (32 + 2 * $e + 2 * $i - $h - $k) % 7
    $m = [math]::Floor(($a + 11 * $h + 22 * $l) / 451)
    $month = [math]::Floor(($h + $l - 7 * $m + 114) / 31)
    $day = (($h + $l - 7 * $m + 114) % 31) + 1
    return [datetime]::new($Year, [int]$month, [int]$day)
}

function Get-NrwHolidays([int]$Year) {
    $easter = Get-EasterSunday $Year
    $map = @{}
    $map[[datetime]::new($Year,1,1).ToString("yyyy-MM-dd")] = "Neujahr"
    $map[$easter.AddDays(-2).ToString("yyyy-MM-dd")] = "Karfreitag"
    $map[$easter.AddDays(1).ToString("yyyy-MM-dd")] = "Ostermontag"
    $map[[datetime]::new($Year,5,1).ToString("yyyy-MM-dd")] = "Tag der Arbeit"
    $map[$easter.AddDays(39).ToString("yyyy-MM-dd")] = "Christi Himmelfahrt"
    $map[$easter.AddDays(50).ToString("yyyy-MM-dd")] = "Pfingstmontag"
    $map[$easter.AddDays(60).ToString("yyyy-MM-dd")] = "Fronleichnam"
    $map[[datetime]::new($Year,10,3).ToString("yyyy-MM-dd")] = "Tag der Deutschen Einheit"
    $map[[datetime]::new($Year,11,1).ToString("yyyy-MM-dd")] = "Allerheiligen"
    $map[[datetime]::new($Year,12,25).ToString("yyyy-MM-dd")] = "1. Weihnachtstag"
    $map[[datetime]::new($Year,12,26).ToString("yyyy-MM-dd")] = "2. Weihnachtstag"
    return $map
}

function Get-VacationRanges {
    $ranges = @()
    foreach ($line in $txtVacation.Lines) {
        $line = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        # Trennung bewusst über " - " bzw. Bindestrich zwischen zwei Datumswerten.
        if ($line -notmatch '^\s*(\d{2}\.\d{2}\.\d{4})\s*-\s*(\d{2}\.\d{2}\.\d{4})\s*$') {
            throw "Ungültiger Urlaubszeitraum: '$line'. Bitte verwenden: TT.MM.JJJJ - TT.MM.JJJJ"
        }

        $start = Parse-DateDE $matches[1]
        $end = Parse-DateDE $matches[2]
        if ($null -eq $start -or $null -eq $end) {
            throw "Ungültiges Datum im Urlaubszeitraum '$line'."
        }
        if ($end.Date -lt $start.Date) {
            throw "Beim Urlaubszeitraum '$line' liegt das Enddatum vor dem Startdatum."
        }

        $ranges += [pscustomobject]@{
            Start = $start.Date
            End = $end.Date
        }
    }
    return $ranges
}

function Get-BlockedReason([datetime]$Date) {
    $selectedYear = [int]$numYear.Value
    if ($Date.Year -ne $selectedYear) {
        return ("Datum gehoert nicht zum ausgewaehlten Steuerjahr {0}." -f $selectedYear)
    }

    if ($chkHolidays.Checked) {
        $holidays = Get-NrwHolidays $Date.Year
        $key = $Date.ToString("yyyy-MM-dd")
        if ($holidays.ContainsKey($key)) {
            return "Feiertag in NRW: $($holidays[$key])"
        }
    }

    foreach ($range in (Get-VacationRanges)) {
        if ($Date.Date -ge $range.Start -and $Date.Date -le $range.End) {
            return ("Urlaub / Abwesenheit: {0:dd.MM.yyyy} - {1:dd.MM.yyyy}" -f $range.Start,$range.End)
        }
    }

    return $null
}

function Apply-DayCheckToSelectedWeeks([int]$DayIndex,[bool]$Checked) {
    if ($script:LoadingDetail) { return }
    if ($DayIndex -lt 0 -or $DayIndex -gt 4) { return }

    $selectedRows = @($grid.SelectedRows)

    # Wenn keine explizite Mehrfachauswahl existiert, immer die aktuell geoeffnete KW verwenden.
    if ($selectedRows.Count -eq 0 -and $script:CurrentIndex -ge 0) {
        $selectedRows = @($grid.Rows[$script:CurrentIndex])
    }

    if ($selectedRows.Count -eq 0) { return }

    foreach ($row in $selectedRows) {
        if ($null -eq $row) { continue }

        $weekIndex = [int]$row.Index
        if ($weekIndex -lt 0 -or $weekIndex -ge $script:Weeks.Count) { continue }

        $day = $script:Weeks[$weekIndex].Days[$DayIndex]
        $dt = [datetime]::ParseExact([string]$day.Date,"yyyy-MM-dd",$null)

        if ($Checked) {
            $blocked = Get-BlockedReason $dt
            if ($blocked) {
                # Nur fuer die aktuell sichtbare KW die Checkbox sofort wieder zuruecknehmen.
                if ($weekIndex -eq $script:CurrentIndex) {
                    $controls = @($chkMo,$chkDi,$chkMi,$chkDo,$chkFr)
                    $script:LoadingDetail = $true
                    $controls[$DayIndex].Checked = $false
                    $script:LoadingDetail = $false
                }

                Show-Error ("Der {0:dd.MM.yyyy} kann nicht als Betreuungstag gesetzt werden.`r`n`r`n{1}" -f $dt,$blocked)
                continue
            }

            # WICHTIG: direkt ins Datenmodell schreiben.
            $day.Checked = $true

            # Standardwerte fuer neue Termine nur dann setzen, wenn noch nichts hinterlegt ist.
            if ([string]::IsNullOrWhiteSpace([string]$day.Person)) {
                $day.Person = [string]$txtCare.Text
            }

            if ([string]::IsNullOrWhiteSpace([string]$day.Art)) {
                $day.Art = if($rbPaid.Checked){"entgeltlich"}else{"unentgeltlich"}
            }

            if ([string]$day.Art -eq "entgeltlich") {
                $day.Fahrtkosten = $false
                if ((Parse-DecimalDE ([string]$day.Stundenlohn)) -le 0 -and [double]$numFee.Value -gt 0) {
                    $day.Stundenlohn = ([double]$numFee.Value).ToString(
                        "0.00",
                        [Globalization.CultureInfo]::GetCultureInfo("de-DE")
                    )
                }
            } else {
                $day.Fahrtkosten = $true
            }
        }
        else {
            # Nur den Betreuungshaken entfernen; Detailwerte bleiben erhalten.
            $day.Checked = $false
        }

        # Diese KW sofort neu zeichnen.
        Refresh-WeekRow $weekIndex
    }

    # Bei der aktuell sichtbaren KW die anderen Detailfelder aus dem Modell
    # nicht neu laden - sonst wuerde der Benutzer beim Tippen gestoert.
    Update-Summary
    $script:IsDirty = $true

    # DataGridView sofort neu zeichnen, ohne auf SelectionChanged zu warten.
    $grid.Refresh()
    $grid.Invalidate()
}