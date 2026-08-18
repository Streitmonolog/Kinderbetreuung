# Kinderbetreuung 1.0.17 - Daten
# Betreuungstermine, Personenprofile, Speicherung, Summen und interner Volltest.

function Get-PersonProfiles {
    $profiles = @{}

    # Global default person first.
    if (-not [string]::IsNullOrWhiteSpace($txtCare.Text)) {
        $name = $txtCare.Text.Trim()
        $profiles[$name.ToLowerInvariant()] = [pscustomobject]@{
            Name = $name
            Rate = 0.0
            Type = "unentgeltlich"
        }
    }

    foreach ($w in $script:Weeks) {
        foreach ($day in $w.Days) {
            $name = [string]$day.Person
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            $name = $name.Trim()
            $key = $name.ToLowerInvariant()

            $rate = Parse-DecimalDE ([string]$day.Stundenlohn)
            $type = [string]$day.Art
            if ($type -notin @("unentgeltlich","entgeltlich")) { $type = "unentgeltlich" }

            if (-not $profiles.ContainsKey($key)) {
                $profiles[$key] = [pscustomobject]@{
                    Name = $name
                    Rate = $rate
                    Type = $type
                }
            } else {
                # Prefer a known paid rate over zero / older unpaid profile.
                if ($rate -gt 0) {
                    $profiles[$key].Rate = $rate
                    $profiles[$key].Type = "entgeltlich"
                    $profiles[$key].Name = $name
                }
            }
        }
    }

    return $profiles
}

function Refresh-PersonSuggestions {
    if ($null -eq $cmbMoPerson) { return }

    $profiles = Get-PersonProfiles
    $source = New-Object System.Windows.Forms.AutoCompleteStringCollection

    foreach ($profile in ($profiles.Values | Sort-Object Name)) {
        [void]$source.Add($profile.Name)
    }

    foreach ($ctrl in @($cmbMoPerson,$cmbDiPerson,$cmbMiPerson,$cmbDoPerson,$cmbFrPerson)) {
        $ctrl.AutoCompleteCustomSource = $source
    }
}

function Apply-PersonProfile([System.Windows.Forms.ComboBox]$Control) {
    if ($script:LoadingDetail) { return }
    if ([string]::IsNullOrWhiteSpace($Control.Text)) { return }

    $index = [int]$Control.Tag
    if ($index -lt 0 -or $index -gt 4) { return }

    $profiles = Get-PersonProfiles
    $key = $Control.Text.Trim().ToLowerInvariant()
    if (-not $profiles.ContainsKey($key)) { return }

    $profile = $profiles[$key]
    $typeControls = @($cmbMoType,$cmbDiType,$cmbMiType,$cmbDoType,$cmbFrType)
    $rateControls = @($txtMoRate,$txtDiRate,$txtMiRate,$txtDoRate,$txtFrRate)
    $travelControls = @($chkMoTravel,$chkDiTravel,$chkMiTravel,$chkDoTravel,$chkFrTravel)

    if ($profile.Rate -gt 0) {
        $typeControls[$index].SelectedItem = "entgeltlich"

        # Only auto-fill if currently empty/zero, so manual changes are respected.
        if ((Parse-DecimalDE $rateControls[$index].Text) -le 0) {
            $rateControls[$index].Text = $profile.Rate.ToString(
                "0.00",
                [Globalization.CultureInfo]::GetCultureInfo("de-DE")
            )
        }

        $travelControls[$index].Checked = $false
    } elseif ($profile.Type -eq "unentgeltlich") {
        if ([string]::IsNullOrWhiteSpace([string]$typeControls[$index].SelectedItem)) {
            $typeControls[$index].SelectedItem = "unentgeltlich"
        }
    }

    Save-CurrentDetail
    Update-Summary
}

function Save-CurrentDetail {
    if ($script:LoadingDetail -or $script:CurrentIndex -lt 0 -or $script:CurrentIndex -ge $script:Weeks.Count) { return }

    $week = $script:Weeks[$script:CurrentIndex]
    $checks = @($chkMo,$chkDi,$chkMi,$chkDo,$chkFr)
    $costs = @($txtMoCost,$txtDiCost,$txtMiCost,$txtDoCost,$txtFrCost)
    $cats = @($txtMoCat,$txtDiCat,$txtMiCat,$txtDoCat,$txtFrCat)
    $notes = @($txtMoNote,$txtDiNote,$txtMiNote,$txtDoNote,$txtFrNote)
    $excursionTargets = @($txtMoTrip,$txtDiTrip,$txtMiTrip,$txtDoTrip,$txtFrTrip)
    $excursionKm = @($txtMoTripKm,$txtDiTripKm,$txtMiTripKm,$txtDoTripKm,$txtFrTripKm)
    $entryCosts = @($txtMoEntry,$txtDiEntry,$txtMiEntry,$txtDoEntry,$txtFrEntry)
    $persons = @($cmbMoPerson,$cmbDiPerson,$cmbMiPerson,$cmbDoPerson,$cmbFrPerson)
    $types = @($cmbMoType,$cmbDiType,$cmbMiType,$cmbDoType,$cmbFrType)
    $hours = @($txtMoHours,$txtDiHours,$txtMiHours,$txtDoHours,$txtFrHours)
    $rates = @($txtMoRate,$txtDiRate,$txtMiRate,$txtDoRate,$txtFrRate)
    $travelChecks = @($chkMoTravel,$chkDiTravel,$chkMiTravel,$chkDoTravel,$chkFrTravel)

    for ($i=0; $i -lt 5; $i++) {
        $week.Days[$i].Checked = $checks[$i].Checked
        $week.Days[$i].Kosten = $costs[$i].Text
        $week.Days[$i].Kategorie = $cats[$i].Text
        $week.Days[$i].Notiz = $notes[$i].Text
        $week.Days[$i].AusflugZiel = $excursionTargets[$i].Text
        $week.Days[$i].AusflugKm = $excursionKm[$i].Text
        $week.Days[$i].Eintritt = $entryCosts[$i].Text
        $week.Days[$i].Person = $persons[$i].Text
        $week.Days[$i].Art = [string]$types[$i].SelectedItem
        $week.Days[$i].Stunden = $hours[$i].Text
        $week.Days[$i].Stundenlohn = $rates[$i].Text
        $week.Days[$i].Fahrtkosten = [bool]$travelChecks[$i].Checked
    }
    Refresh-WeekRow $script:CurrentIndex
    Update-Summary
}

function Load-WeekDetail([int]$Index) {
    if ($Index -lt 0 -or $Index -ge $script:Weeks.Count) { return }
    Save-CurrentDetail

    $script:LoadingDetail = $true
    $script:CurrentIndex = $Index
    $week = $script:Weeks[$Index]
    $mon = [datetime]::ParseExact($week.Monday,"yyyy-MM-dd",$null)
    $lblSelectedWeek.Text = ("KW {0:D2}   |   {1:dd.MM.yyyy} - {2:dd.MM.yyyy}" -f $week.Week,$mon,$mon.AddDays(6))

    $checks = @($chkMo,$chkDi,$chkMi,$chkDo,$chkFr)
    $labels = @($lblMoDate,$lblDiDate,$lblMiDate,$lblDoDate,$lblFrDate)
    $costs = @($txtMoCost,$txtDiCost,$txtMiCost,$txtDoCost,$txtFrCost)
    $cats = @($txtMoCat,$txtDiCat,$txtMiCat,$txtDoCat,$txtFrCat)
    $notes = @($txtMoNote,$txtDiNote,$txtMiNote,$txtDoNote,$txtFrNote)
    $excursionTargets = @($txtMoTrip,$txtDiTrip,$txtMiTrip,$txtDoTrip,$txtFrTrip)
    $excursionKm = @($txtMoTripKm,$txtDiTripKm,$txtMiTripKm,$txtDoTripKm,$txtFrTripKm)
    $entryCosts = @($txtMoEntry,$txtDiEntry,$txtMiEntry,$txtDoEntry,$txtFrEntry)
    $persons = @($cmbMoPerson,$cmbDiPerson,$cmbMiPerson,$cmbDoPerson,$cmbFrPerson)
    $types = @($cmbMoType,$cmbDiType,$cmbMiType,$cmbDoType,$cmbFrType)
    $hours = @($txtMoHours,$txtDiHours,$txtMiHours,$txtDoHours,$txtFrHours)
    $rates = @($txtMoRate,$txtDiRate,$txtMiRate,$txtDoRate,$txtFrRate)
    $travelChecks = @($chkMoTravel,$chkDiTravel,$chkMiTravel,$chkDoTravel,$chkFrTravel)
    $names = @("Montag","Dienstag","Mittwoch","Donnerstag","Freitag")

    for ($i=0; $i -lt 5; $i++) {
        $d = [datetime]::ParseExact($week.Days[$i].Date,"yyyy-MM-dd",$null)
        if ($d.Year -ne [int]$numYear.Value) {
            $labels[$i].Text = $d.ToString("dd.MM.yyyy") + " *"
            $labels[$i].ForeColor = [System.Drawing.Color]::DarkRed
        } else {
            $labels[$i].Text = $d.ToString("dd.MM.yyyy")
            $labels[$i].ForeColor = [System.Drawing.SystemColors]::ControlText
        }
        $checks[$i].Text = $names[$i]
        $checks[$i].Checked = [bool]$week.Days[$i].Checked
        $costs[$i].Text = [string]$week.Days[$i].Kosten
        $cats[$i].Text = [string]$week.Days[$i].Kategorie
        $notes[$i].Text = [string]$week.Days[$i].Notiz
        $excursionTargets[$i].Text = [string]$week.Days[$i].AusflugZiel
        $excursionKm[$i].Text = [string]$week.Days[$i].AusflugKm
        $entryCosts[$i].Text = [string]$week.Days[$i].Eintritt
        $persons[$i].Text = if([string]::IsNullOrWhiteSpace([string]$week.Days[$i].Person)){$txtCare.Text}else{[string]$week.Days[$i].Person}
        $typeVal = [string]$week.Days[$i].Art
        if($typeVal -notin @("unentgeltlich","entgeltlich")){$typeVal="unentgeltlich"}
        $types[$i].SelectedItem = $typeVal
        $hours[$i].Text = [string]$week.Days[$i].Stunden
        $rates[$i].Text = [string]$week.Days[$i].Stundenlohn
        $travelChecks[$i].Checked = [bool]$week.Days[$i].Fahrtkosten
    }
    $script:LoadingDetail = $false
    Refresh-PersonSuggestions
}

function Refresh-WeekList {
    $grid.Rows.Clear()
    foreach ($week in $script:Weeks) {
        $idx = $week.Week - 1
        $mon = [datetime]::ParseExact($week.Monday,"yyyy-MM-dd",$null)
        $days = ($week.Days | Where-Object { $_.Checked }).Count
        $dayNames = @()
        if ($week.Days[0].Checked) { $dayNames += "Mo" }
        if ($week.Days[1].Checked) { $dayNames += "Di" }
        if ($week.Days[2].Checked) { $dayNames += "Mi" }
        if ($week.Days[3].Checked) { $dayNames += "Do" }
        if ($week.Days[4].Checked) { $dayNames += "Fr" }

        [void]$grid.Rows.Add(
            ("KW {0:D2}" -f $week.Week),
            ("{0:dd.MM.} - {1:dd.MM.yyyy}" -f $mon,$mon.AddDays(6)),
            $days,
            ($dayNames -join ", ")
        )
    }
}

function Refresh-WeekRow([int]$Index) {
    if ($Index -lt 0 -or $Index -ge $script:Weeks.Count) { return }
    if ($Index -lt 0 -or $Index -ge $grid.Rows.Count) { return }

    $week = $script:Weeks[$Index]
    $row = $grid.Rows[$Index]

    $dayNames = @("Mo","Di","Mi","Do","Fr")
    $marked = New-Object System.Collections.Generic.List[string]
    $count = 0

    for($i=0; $i -lt $week.Days.Count; $i++){
        if([bool]$week.Days[$i].Checked){
            $count++
            if($i -lt $dayNames.Count){
                [void]$marked.Add($dayNames[$i])
            }
        }
    }

    # Werte IMMER explizit setzen - auch 0 und 1.
    $row.Cells["Anzahl"].Value = [string]$count
    $row.Cells["Tage"].Value = ($marked -join ", ")

    # DataGridView dazu zwingen, die beiden Zellen sofort neu zu zeichnen.
    $grid.InvalidateCell($row.Cells["Anzahl"])
    $grid.InvalidateCell($row.Cells["Tage"])
}
function Apply-PatternToSelected([bool]$Mo,[bool]$Di,[bool]$Mi,[bool]$Do,[bool]$Fr) {
    Save-CurrentDetail
    if ($grid.SelectedRows.Count -eq 0) { Show-Info "Bitte mindestens eine KW auswählen."; return }

    foreach ($row in $grid.SelectedRows) {
        $i = $row.Index
        if ($i -ge 0 -and $i -lt $script:Weeks.Count) {
            $vals = @($Mo,$Di,$Mi,$Do,$Fr)
            for ($d=0; $d -lt 5; $d++) {
                $script:Weeks[$i].Days[$d].Checked = $vals[$d]
            }
            Refresh-WeekRow $i
        }
    }
    if ($script:CurrentIndex -ge 0) { Load-WeekDetail $script:CurrentIndex }
    Update-Summary
    Save-Data
}

function Generate-DemoData {
    $kmOne = [double]$numKm.Value
    $rate = [double]$numRate.Value

    if ($kmOne -le 0) {
        Show-Error "Der Volltest kann nur gestartet werden, wenn eine einfache Strecke in km hinterlegt ist."
        return
    }

    if ($rate -le 0) {
        Show-Error "Der Volltest kann nur gestartet werden, wenn eine gueltige EUR/km-Pauschale hinterlegt ist."
        return
    }

    $answer = [System.Windows.Forms.MessageBox]::Show(
        "Interner Volltest starten?`r`n`r`nDer Test befuellt das aktuell ausgewaehlte Jahr mit vielen synthetischen Betreuungstagen und Zusatzwerten. Die Berechnung wird auf einen zufaelligen Zielwert 2 bis 8 Prozent unter 4.800 EUR abziehbarem Betrag gebracht.`r`n`r`nVorhandene Daten dieses Jahres werden ersetzt.",
        "Interner Volltest",
        "YesNo",
        "Warning"
    )
    if ($answer -ne "Yes") { return }

    # Backup der echten Datendatei
    try {
        if (Test-Path $DataFile) {
            $backup = Join-Path $UserDataDir ("kinderbetreuung-daten_backup_{0}.json" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
            Copy-Item $DataFile $backup -Force
        }
    } catch {}

    $script:DemoMode = $true

    foreach ($w in $script:Weeks) {
        foreach ($d in $w.Days) {
            $d.Checked = $false
            $d.Kosten = "0,00"
            $d.Kategorie = ""
            $d.Notiz = ""
            $d.AusflugZiel = ""
            $d.AusflugKm = "0,00"
            $d.Eintritt = "0,00"
        }
    }

    # Ziel 2-8 % unter maximal abziehbaren 4.800 EUR.
    # Bei 80 % Abzug entspricht das einem Brutto-Kosten-Ziel von Ziel / 0,80.
    $percentBelow = Get-Random -Minimum 2 -Maximum 9
    $deductibleTarget = [math]::Round(4800.0 * (1.0 - ($percentBelow / 100.0)), 2)
    $grossTarget = [math]::Round($deductibleTarget / 0.80, 2)

    # Mo-Fr moegliche Tage sammeln, aber Feiertage, Urlaub und Fremdjahr sperren.
    $eligible = @()
    foreach ($w in $script:Weeks) {
        for ($dayIndex=0; $dayIndex -lt $w.Days.Count; $dayIndex++) {
            $d = $w.Days[$dayIndex]
            $dt = [datetime]::ParseExact([string]$d.Date,"yyyy-MM-dd",$null)
            if (-not (Get-BlockedReason $dt)) {
                $eligible += [pscustomobject]@{
                    Week = $w
                    Day = $d
                    Date = $dt
                    DayIndex = $dayIndex
                }
            }
        }
    }

    if ($eligible.Count -eq 0) {
        $script:DemoMode = $false
        Show-Error "Es wurden keine geeigneten Tage gefunden."
        return
    }

    # Viele Datensaetze erzeugen: ca. 35-55 % der verfuegbaren Werktage,
    # dabei mindestens 40 Tage, sofern verfuegbar.
    $minWanted = [math]::Min(40, $eligible.Count)
    $wanted = [int][math]::Round($eligible.Count * ((Get-Random -Minimum 35 -Maximum 56) / 100.0))
    if ($wanted -lt $minWanted) { $wanted = $minWanted }
    if ($wanted -gt $eligible.Count) { $wanted = $eligible.Count }

    $selected = @($eligible | Sort-Object { Get-Random } | Select-Object -First $wanted)

    # Fahrtkosten je Betreuungstag
    $travelPerDay = [math]::Round(($kmOne * 2.0) * $rate, 2)
    $feePerDay = if($rbPaid.Checked){ [double]$numFee.Value } else { 0.0 }
    $basePerDay = $travelPerDay + $feePerDay
    $travelTotal = [math]::Round($basePerDay * $selected.Count, 2)

    # Falls die Fahrtkosten allein schon ueber dem Ziel liegen, weniger Tage verwenden.
    while ($selected.Count -gt 1 -and $travelTotal -gt $grossTarget) {
        $selected = @($selected | Select-Object -First ($selected.Count - 1))
        $travelTotal = [math]::Round($basePerDay * $selected.Count, 2)
    }

    if ($travelTotal -gt $grossTarget) {
        $script:DemoMode = $false
        Show-Error ("Mit {0:N1} km einfacher Strecke und {1:N2} EUR/km liegt bereits ein einzelner Betreuungstag ueber dem Testziel. Bitte kleinere Testwerte verwenden." -f $kmOne,$rate)
        return
    }

    # Zusatz-/Ausflugsfelder werden nur fuer Speicher- und Voll-Exporttests gefuellt.
    # Sie werden NICHT verwendet, um den steuerlichen Zielbetrag zu erreichen.
    $remaining = [math]::Round($grossTarget - $travelTotal, 2)

    # Falls wir noch deutlich unter dem Ziel liegen, so viele weitere Tage wie moeglich nehmen.
    # Das Ziel kann nur ueber Fahrtkosten + ggf. Betreuungsentgelt angenaehert werden.
    if ($remaining -gt $basePerDay) {
        $remainingCandidates = @($eligible | Where-Object {
            $candidate = $_
            -not ($selected | Where-Object { $_.Date -eq $candidate.Date })
        } | Sort-Object { Get-Random })

        foreach ($candidate in $remainingCandidates) {
            if (($travelTotal + $basePerDay) -gt $grossTarget) { break }
            $selected += $candidate
            $travelTotal = [math]::Round($travelTotal + $basePerDay,2)
        }
    }

    foreach ($item in ($selected | Sort-Object { Get-Random } | Select-Object -First ([math]::Max(1,[int]($selected.Count*0.20))))) {
        $item.Day.Kosten = ((Get-Random -Minimum 300 -Maximum 1801) / 100.0).ToString("0.00",[Globalization.CultureInfo]::GetCultureInfo("de-DE"))
        $item.Day.Kategorie = "TEST"
        $item.Day.Notiz = "Synthetischer Volltest - nicht steuerrelevant"
    }

    # Alle ausgewaehlten Tage aktivieren; einige bekommen zusaetzlich
    # Freizeit-/Ausflugswerte fuer den vollstaendigen Exporttest.
    foreach ($item in $selected) {
        $item.Day.Checked = $true
        if ([string]$item.Day.Art -eq "entgeltlich") {
            $item.Day.Fahrtkosten = $false
        } else {
            $item.Day.Fahrtkosten = $true
        }

        if ((Get-Random -Minimum 1 -Maximum 101) -le 15) {
            $item.Day.AusflugZiel = "Test-Ausflug"
            $item.Day.AusflugKm = ((Get-Random -Minimum 20 -Maximum 1201) / 10.0).ToString("0.0",[Globalization.CultureInfo]::GetCultureInfo("de-DE"))
            $item.Day.Eintritt = ((Get-Random -Minimum 300 -Maximum 2501) / 100.0).ToString("0.00",[Globalization.CultureInfo]::GetCultureInfo("de-DE"))
            if ([string]::IsNullOrWhiteSpace([string]$item.Day.Notiz)) {
                $item.Day.Notiz = "Synthetischer Volltest"
            }
        }
    }

    Refresh-WeekList
    Update-Summary
    if ($script:CurrentIndex -ge 0) { Load-WeekDetail $script:CurrentIndex }

    # Normaler Speicherweg, aber Datensaetze bleiben eindeutig als Test markiert.
    $script:DemoMode = $false
    Save-Data

    # Tatsachliche Testsumme nachrechnen.
    $grossActual = 0.0
    foreach ($item in $selected) {
        $grossActual += $basePerDay
    }
    $grossActual = [math]::Round($grossActual,2)
    $deductibleActual = [math]::Round([math]::Min($grossActual * 0.80,4800.0),2)

    Show-Info (
        "Volltest abgeschlossen.`r`n`r`n" +
        ("Betreuungstage: {0}`r`n" -f $selected.Count) +
        ("Fahrtkosten je Tag: {0:N2} EUR`r`n" -f $travelPerDay) +
        ("Gesamtkosten Test: {0:N2} EUR`r`n" -f $grossActual) +
        ("Abziehbarer Testbetrag: {0:N2} EUR`r`n" -f $deductibleActual) +
        ("Ziel: {0}% unter 4.800 EUR`r`n`r`n" -f $percentBelow) +
        "Die erzeugten Zusatzwerte sind mit TEST / Synthetischer Volltest gekennzeichnet."
    )
}

function Save-Data {
    Save-CurrentDetail

    if (-not (Test-Path $UserDataDir)) {
        New-Item -ItemType Directory -Path $UserDataDir -Force | Out-Null
    }

    $payload = [pscustomobject]@{
        Jahr=[int]$numYear.Value
        Kind=[string]$txtChild.Text
        Betreuungspersonen=[string]$txtCare.Text
        KilometerEinfach=$numKm.Value.ToString([Globalization.CultureInfo]::InvariantCulture)
        EuroKm=$numRate.Value.ToString([Globalization.CultureInfo]::InvariantCulture)
        Urlaub=[string]$txtVacation.Text
        FeiertageNRW=[bool]$chkHolidays.Checked
        BetreuungsvertragVorhanden=[bool]$chkContract.Checked
        IBAN=[string]$txtIban.Text
        BetreuungMitEntgelt=[bool]$rbPaid.Checked
        BetreuungUnentgeltlich=[bool]$rbUnpaid.Checked
        BetreuungEntgelt=$numFee.Value.ToString([Globalization.CultureInfo]::InvariantCulture)
        Wochen=@($script:Weeks)
        GespeichertAm=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    }

    try {
        # Vorhandene Datei zuerst sichern.
        if (Test-Path $DataFile) {
            $backupFile = Join-Path $UserDataDir "kinderbetreuung-daten.bak"
            Copy-Item $DataFile $backupFile -Force
        }

        $json = $payload | ConvertTo-Json -Depth 12

        # Direkt und synchron schreiben.
        Set-Content -Path $DataFile -Value $json -Encoding UTF8 -Force

        # Sicherstellen, dass die Datei wirklich existiert und Inhalt hat.
        if (-not (Test-Path $DataFile)) {
            throw "Die Datendatei wurde nach dem Schreiben nicht gefunden."
        }

        $fileInfo = Get-Item $DataFile
        if ($fileInfo.Length -le 10) {
            throw "Die gespeicherte Datendatei ist leer oder unvollstaendig."
        }

        # Sofortige Ruecklese-Pruefung.
        $verifyRaw = Get-Content $DataFile -Raw -Encoding UTF8
        $verify = $verifyRaw | ConvertFrom-Json

        if ($null -eq $verify) {
            throw "Die gespeicherte JSON-Datei konnte nicht wieder eingelesen werden."
        }

        if ([string]$verify.Kind -ne [string]$txtChild.Text) {
            throw "Pruefung fehlgeschlagen: Der gespeicherte Kind-Name stimmt nicht mit der Eingabe ueberein."
        }

        if ([string]$verify.Betreuungspersonen -ne [string]$txtCare.Text) {
            throw "Pruefung fehlgeschlagen: Die gespeicherte Betreuungsperson stimmt nicht mit der Eingabe ueberein."
        }

        if ([int]$verify.Jahr -ne [int]$numYear.Value) {
            throw "Pruefung fehlgeschlagen: Das gespeicherte Jahr stimmt nicht mit der Eingabe ueberein."
        }

        if (@($verify.Wochen).Count -ne @($script:Weeks).Count) {
            throw "Pruefung fehlgeschlagen: Die Anzahl der gespeicherten Kalenderwochen stimmt nicht."
        }

        return $true
    }
    catch {
        throw ("Speichern fehlgeschlagen.`r`n`r`nPfad:`r`n{0}`r`n`r`nFehler:`r`n{1}" -f $DataFile,$_.Exception.Message)
    }
}

function Repair-YearWeeks {
    $year = [int]$numYear.Value

    $jan1 = [datetime]::new($year,1,1)
    $dow = [int]$jan1.DayOfWeek
    if ($dow -eq 0) { $dow = 7 }
    $firstMonday = $jan1.AddDays(1 - $dow)

    $dec31 = [datetime]::new($year,12,31)
    $dowEnd = [int]$dec31.DayOfWeek
    if ($dowEnd -eq 0) { $dowEnd = 7 }
    $lastMonday = $dec31.AddDays(1 - $dowEnd)

    $expectedMondays = @()
    $mon = $firstMonday
    while ($mon -le $lastMonday) {
        $expectedMondays += $mon
        $mon = $mon.AddDays(7)
    }

    $existingByMonday = @{}
    foreach ($w in $script:Weeks) {
        if ($w.Monday) {
            $existingByMonday[[string]$w.Monday] = $w
        }
    }

    $newWeeks = @()
    $changed = $false

    for ($i=0; $i -lt $expectedMondays.Count; $i++) {
        $monday = $expectedMondays[$i]
        $key = $monday.ToString("yyyy-MM-dd")
        $displayWeek = $i + 1

        if ($existingByMonday.ContainsKey($key)) {
            $week = $existingByMonday[$key]
            $week.Week = $displayWeek

            foreach ($day in $week.Days) {
                if ($null -eq $day.PSObject.Properties["AusflugZiel"]) {
                    $day | Add-Member -NotePropertyName AusflugZiel -NotePropertyValue ""
                }
                if ($null -eq $day.PSObject.Properties["AusflugKm"]) {
                    $day | Add-Member -NotePropertyName AusflugKm -NotePropertyValue "0,00"
                }
                if ($null -eq $day.PSObject.Properties["Eintritt"]) {
                    $day | Add-Member -NotePropertyName Eintritt -NotePropertyValue "0,00"
                }
                if ($null -eq $day.PSObject.Properties["Fahrtkosten"]) { $day | Add-Member -NotePropertyName Fahrtkosten -NotePropertyValue $true }

            }
            $newWeeks += $week
        } else {
            $days = @()
            for ($d=0; $d -lt 5; $d++) {
                $date = $monday.AddDays($d)
                $days += [pscustomobject]@{
                    Date = $date.ToString("yyyy-MM-dd")
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
            $newWeeks += [pscustomobject]@{
                Week = $displayWeek
                Monday = $key
                Days = $days
            }
            $changed = $true
        }
    }

    if ($script:Weeks.Count -ne $expectedMondays.Count) {
        $changed = $true
    }

    $script:Weeks = @($newWeeks)

    if ($changed) {
        Refresh-WeekList
        Update-Summary
        try { Save-Data } catch {}
    }

    return $changed
}

function Load-Data {
    function Read-DataFile([string]$Path) {
        if (-not (Test-Path $Path)) { return $null }
        $raw = Get-Content $Path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) {
            throw "Datendatei ist leer: $Path"
        }
        return ($raw | ConvertFrom-Json)
    }

    if (-not (Test-Path $DataFile)) {
        New-YearData
        $script:IsDirty = $false
        return
    }

    $o = $null
    $loadedFromBackup = $false

    try {
        $o = Read-DataFile $DataFile
    }
    catch {
        $backupFile = Join-Path $UserDataDir "kinderbetreuung-daten.bak"
        if (Test-Path $backupFile) {
            try {
                $o = Read-DataFile $backupFile
                $loadedFromBackup = $true
            }
            catch {
                throw ("Weder Datendatei noch Sicherung konnten geladen werden.`r`n`r`nDatendatei: {0}`r`nSicherung: {1}" -f $DataFile,$backupFile)
            }
        }
        else {
            throw ("Die Datendatei konnte nicht geladen werden und es existiert keine Sicherung.`r`n`r`n{0}" -f $DataFile)
        }
    }

    if ($null -eq $o) {
        throw "Die Datendatei konnte nicht gelesen werden."
    }

    # Beim Befuellen der Controls keine Aenderungen markieren.
    $script:LoadingDetail = $true
    try {
        if ($o.Jahr) { $numYear.Value=[decimal]$o.Jahr }
        $txtChild.Text=[string]$o.Kind
        $txtCare.Text=[string]$o.Betreuungspersonen

        if ($o.KilometerEinfach) {
            $numKm.Value=[decimal]::Parse([string]$o.KilometerEinfach,[Globalization.CultureInfo]::InvariantCulture)
        }
        if ($o.EuroKm) {
            $numRate.Value=[decimal]::Parse([string]$o.EuroKm,[Globalization.CultureInfo]::InvariantCulture)
        }
        if ($null -ne $o.Urlaub) { $txtVacation.Text=[string]$o.Urlaub }
        if ($null -ne $o.FeiertageNRW) { $chkHolidays.Checked=[bool]$o.FeiertageNRW }
        if ($null -ne $o.BetreuungsvertragVorhanden) { $chkContract.Checked=[bool]$o.BetreuungsvertragVorhanden }
        if ($null -ne $o.IBAN) { $txtIban.Text=[string]$o.IBAN }

        if ($null -ne $o.BetreuungMitEntgelt) { $rbPaid.Checked=[bool]$o.BetreuungMitEntgelt }
        if ($null -ne $o.BetreuungUnentgeltlich) { $rbUnpaid.Checked=[bool]$o.BetreuungUnentgeltlich }
        if (-not $rbPaid.Checked -and -not $rbUnpaid.Checked) { $rbUnpaid.Checked=$true }

        if ($o.BetreuungEntgelt) {
            $numFee.Value=[decimal]::Parse([string]$o.BetreuungEntgelt,[Globalization.CultureInfo]::InvariantCulture)
        }

        $script:Weeks = @($o.Wochen)

        foreach ($w in $script:Weeks) {
            foreach ($d in $w.Days) {
                if ($null -eq $d.PSObject.Properties["Person"]) {
                    $d | Add-Member -NotePropertyName Person -NotePropertyValue ""
                }
                if ($null -eq $d.PSObject.Properties["Art"]) {
                    $d | Add-Member -NotePropertyName Art -NotePropertyValue "unentgeltlich"
                }
                if ($null -eq $d.PSObject.Properties["Stunden"]) {
                    $d | Add-Member -NotePropertyName Stunden -NotePropertyValue "0,00"
                }
                if ($null -eq $d.PSObject.Properties["Stundenlohn"]) {
                    $d | Add-Member -NotePropertyName Stundenlohn -NotePropertyValue "0,00"
                }
                if ($null -eq $d.PSObject.Properties["Fahrtkosten"]) {
                    $d | Add-Member -NotePropertyName Fahrtkosten -NotePropertyValue $true
                }
                if ($null -eq $d.PSObject.Properties["AusflugZiel"]) {
                    $d | Add-Member -NotePropertyName AusflugZiel -NotePropertyValue ""
                }
                if ($null -eq $d.PSObject.Properties["AusflugKm"]) {
                    $d | Add-Member -NotePropertyName AusflugKm -NotePropertyValue "0,00"
                }
                if ($null -eq $d.PSObject.Properties["Eintritt"]) {
                    $d | Add-Member -NotePropertyName Eintritt -NotePropertyValue "0,00"
                }
                if ($null -eq $d.PSObject.Properties["Kosten"]) {
                    $d | Add-Member -NotePropertyName Kosten -NotePropertyValue "0,00"
                }
                if ($null -eq $d.PSObject.Properties["Kategorie"]) {
                    $d | Add-Member -NotePropertyName Kategorie -NotePropertyValue ""
                }
                if ($null -eq $d.PSObject.Properties["Notiz"]) {
                    $d | Add-Member -NotePropertyName Notiz -NotePropertyValue ""
                }
            }
        }

        if ($script:Weeks.Count -eq 0) {
            throw "Die gespeicherte Datendatei enthaelt keine Kalenderwochen."
        }
    }
    finally {
        $script:LoadingDetail = $false
    }

    # Alte Dateistaende ggf. um fehlende Jahresendwoche/Felder erweitern.
    $wasRepaired = Repair-YearWeeks

    Refresh-WeekList
    if ($grid.Rows.Count -gt 0) {
        $script:LoadingDetail = $true
        $grid.ClearSelection()
        $grid.Rows[0].Selected = $true
        $grid.CurrentCell = $grid.Rows[0].Cells[0]
        $script:LoadingDetail = $false
        Load-WeekDetail 0
    }

    Update-Summary
    Refresh-PersonSuggestions
    $script:IsDirty = $false

    if ($loadedFromBackup) {
        Show-Info ("Die eigentliche Datendatei konnte nicht geladen werden.`r`nDie Sicherung wurde erfolgreich geladen:`r`n`r`n" + (Join-Path $UserDataDir "kinderbetreuung-daten.bak"))
    }
    elseif ($wasRepaired) {
        Show-Info "Die gespeicherten Jahresdaten wurden auf den aktuellen Programmstand aktualisiert."
    }
}
function Get-CurrentFinancialSummary {
    $kmOne = [double]$numKm.Value
    $rate = [double]$numRate.Value

    $careDays = 0
    $travelTotal = 0.0
    $feeTotal = 0.0

    foreach ($w in $script:Weeks) {
        foreach ($day in $w.Days) {
            if ([bool]$day.Checked) {
                $careDays++
                if ([bool]$day.Fahrtkosten) {
                    $travelTotal += (($kmOne * 2.0) * $rate)
                }

                if ([string]$day.Art -eq "entgeltlich") {
                    $hours = Parse-DecimalDE ([string]$day.Stunden)
                    $hourly = Parse-DecimalDE ([string]$day.Stundenlohn)
                    $feeTotal += ($hours * $hourly)
                }
            }
        }
    }

    $gross = [math]::Round($travelTotal + $feeTotal,2)
    $deductible = [math]::Round([math]::Min($gross * 0.80,4800.0),2)

    return [pscustomobject]@{
        CareDays = $careDays
        TravelTotal = [math]::Round($travelTotal,2)
        FeeTotal = [math]::Round($feeTotal,2)
        Gross = $gross
        Deductible = $deductible
    }
}

function Update-Summary {
    $weekCount = $script:Weeks.Count
    $sum = Get-CurrentFinancialSummary

    $lblSummary.Text = (
        "Wochenzeilen: {0}   |   Betreuungstage: {1}   |   Betreuungssumme: {2:N2} EUR   |   davon aktuell abziehbar: {3:N2} EUR" -f
        $weekCount,$sum.CareDays,$sum.Gross,$sum.Deductible
    )
}

