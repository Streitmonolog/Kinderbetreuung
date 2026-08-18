# Modernes Farbschema passend zum Kinderbetreuungs-Icon.
# Bewusst dezent: klares Blau/Tuerkis, weiche Flaechen, dunkle Typografie.
$clrWindow        = [System.Drawing.Color]::FromArgb(246,249,251)
$clrSurface       = [System.Drawing.Color]::White
$clrSurfaceSoft   = [System.Drawing.Color]::FromArgb(238,246,249)
$clrPrimary       = [System.Drawing.Color]::FromArgb(43,131,159)
$clrPrimaryDark   = [System.Drawing.Color]::FromArgb(31,103,128)
$clrPrimarySoft   = [System.Drawing.Color]::FromArgb(216,237,244)
$clrPrimaryHover  = [System.Drawing.Color]::FromArgb(197,228,238)
$clrBorder        = [System.Drawing.Color]::FromArgb(188,209,218)
$clrText          = [System.Drawing.Color]::FromArgb(31,47,55)
$clrTextMuted     = [System.Drawing.Color]::FromArgb(89,108,116)
$clrGridAlt       = [System.Drawing.Color]::FromArgb(246,250,251)
$clrDangerSoft    = [System.Drawing.Color]::FromArgb(250,238,238)
$clrWhite         = [System.Drawing.Color]::White

# Kinderbetreuung 1.0.17 - Benutzeroberflaeche
# Aufbau der WinForms-Oberflaeche und Event-Verknuepfungen.
# Fachlogik befindet sich in den anderen Modulen.

try {
    $form=New-Object System.Windows.Forms.Form
    $form.Text="Kinderbetreuung 1.0.17"
    $form.Size=[System.Drawing.Size]::new(1200,820)
    $form.StartPosition="CenterScreen"
    $form.BackColor=[System.Drawing.Color]::FromArgb(245,248,250)
    $form.MinimumSize=[System.Drawing.Size]::new(1000,720)
    $iconFile = Join-Path $AppDir "Assets\Kinderbetreuung.ico"
    if (Test-Path $iconFile) {
        try { $form.Icon = New-Object System.Drawing.Icon($iconFile) } catch {}
    }

    $top=New-Object System.Windows.Forms.Panel
    $top.Dock="Top";$top.Height=465
    $top.BackColor=[System.Drawing.Color]::FromArgb(245,248,250)
    $form.Controls.Add($top)

    function AddL($text,$x,$y,$w=120){
        $l=New-Object System.Windows.Forms.Label
        $l.Text=$text;$l.Location=[System.Drawing.Point]::new($x,$y);$l.Size=[System.Drawing.Size]::new($w,22)
        $top.Controls.Add($l)
        return $l
    }

    AddL "Jahr:" 10 12 40 | Out-Null
    $numYear=New-Object System.Windows.Forms.NumericUpDown
    $numYear.Minimum=2020;$numYear.Maximum=2100;$numYear.Value=(Get-Date).Year
    $numYear.Location=[System.Drawing.Point]::new(50,10);$numYear.Width=80;$top.Controls.Add($numYear)

    AddL "Kind:" 150 12 40 | Out-Null
    $txtChild=New-Object System.Windows.Forms.TextBox
    $txtChild.Location=[System.Drawing.Point]::new(190,10);$txtChild.Width=190;$top.Controls.Add($txtChild)

    AddL "Betreuungsperson(en):" 400 12 140 | Out-Null
    $txtCare=New-Object System.Windows.Forms.TextBox
    $txtCare.Location=[System.Drawing.Point]::new(540,10);$txtCare.Width=240;$top.Controls.Add($txtCare)

    AddL "einfache Strecke km:" 800 12 120 | Out-Null
    $numKm=New-Object System.Windows.Forms.NumericUpDown
    $numKm.DecimalPlaces=1;$numKm.Maximum=1000;$numKm.Location=[System.Drawing.Point]::new(920,10);$numKm.Width=65;$top.Controls.Add($numKm)

    AddL "EUR/km:" 995 12 55 | Out-Null
    $numRate=New-Object System.Windows.Forms.NumericUpDown
    $numRate.DecimalPlaces=2;$numRate.Increment=.01;$numRate.Maximum=10;$numRate.Value=.30
    $numRate.Location=[System.Drawing.Point]::new(1050,10);$numRate.Width=65;$top.Controls.Add($numRate)


    AddL "Urlaub / Abwesenheit:" 10 48 130 | Out-Null
    $txtVacation=New-Object System.Windows.Forms.TextBox
    $txtVacation.Location=[System.Drawing.Point]::new(145,45)
    $txtVacation.Size=[System.Drawing.Size]::new(515,42)
    $txtVacation.Multiline=$true
    $txtVacation.ScrollBars="Vertical"
    $top.Controls.Add($txtVacation)

    $lblVacationHint=New-Object System.Windows.Forms.Label
    $lblVacationHint.Text="Eine Zeile je Zeitraum, z. B. 14.07.2025 - 01.08.2025"
    $lblVacationHint.Location=[System.Drawing.Point]::new(675,48)
    $lblVacationHint.Size=[System.Drawing.Size]::new(330,22)
    $top.Controls.Add($lblVacationHint)

    $chkHolidays=New-Object System.Windows.Forms.CheckBox
    $chkHolidays.Text="Feiertage NRW sperren"
    $chkHolidays.Location=[System.Drawing.Point]::new(1015,46)
    $chkHolidays.Size=[System.Drawing.Size]::new(165,24)
    $chkHolidays.Checked=$true
    $top.Controls.Add($chkHolidays)


    $chkContract=New-Object System.Windows.Forms.CheckBox
    $chkContract.Text="Schriftliche Vereinbarung / Betreuungsvertrag vorhanden"
    $chkContract.Location=[System.Drawing.Point]::new(145,92)
    $chkContract.Size=[System.Drawing.Size]::new(340,24)
    $chkContract.Checked=$false
    $top.Controls.Add($chkContract)

    AddL "IBAN Betreuungsperson:" 500 94 145 | Out-Null
    $txtIban=New-Object System.Windows.Forms.TextBox
    $txtIban.Location=[System.Drawing.Point]::new(645,91)
    $txtIban.Size=[System.Drawing.Size]::new(330,22)
    $top.Controls.Add($txtIban)

    $lblIbanHint=New-Object System.Windows.Forms.Label
    $lblIbanHint.Text="optional - zur Dokumentation der unbaren Zahlung"
    $lblIbanHint.Location=[System.Drawing.Point]::new(985,94)
    $lblIbanHint.Size=[System.Drawing.Size]::new(190,22)
    $top.Controls.Add($lblIbanHint)


    AddL "Betreuungsart:" 145 126 100 | Out-Null

    $rbUnpaid=New-Object System.Windows.Forms.RadioButton
    $rbUnpaid.Text="Betreuung unentgeltlich"
    $rbUnpaid.Location=[System.Drawing.Point]::new(245,123)
    $rbUnpaid.Size=[System.Drawing.Size]::new(175,24)
    $rbUnpaid.Checked=$true
    $top.Controls.Add($rbUnpaid)

    $rbPaid=New-Object System.Windows.Forms.RadioButton
    $rbPaid.Text="Betreuung gegen Entgelt"
    $rbPaid.Location=[System.Drawing.Point]::new(430,123)
    $rbPaid.Size=[System.Drawing.Size]::new(175,24)
    $top.Controls.Add($rbPaid)

    AddL "Entgelt je Betreuungstag EUR:" 620 126 175 | Out-Null
    $numFee=New-Object System.Windows.Forms.NumericUpDown
    $numFee.DecimalPlaces=2
    $numFee.Increment=1
    $numFee.Maximum=10000
    $numFee.Location=[System.Drawing.Point]::new(795,123)
    $numFee.Size=[System.Drawing.Size]::new(95,22)
    $numFee.Enabled=$false
    $top.Controls.Add($numFee)

    $lblFeeHint=New-Object System.Windows.Forms.Label
    $lblFeeHint.Text="nur bei Betreuung gegen Entgelt"
    $lblFeeHint.Location=[System.Drawing.Point]::new(900,126)
    $lblFeeHint.Size=[System.Drawing.Size]::new(210,22)
    $top.Controls.Add($lblFeeHint)

    $lblSelectedWeek=New-Object System.Windows.Forms.Label
    $lblSelectedWeek.Text="KW auswählen"
    $lblSelectedWeek.Location=[System.Drawing.Point]::new(10,170)
    $lblSelectedWeek.Size=[System.Drawing.Size]::new(600,28)
    $lblSelectedWeek.Font=New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Bold)
    $top.Controls.Add($lblSelectedWeek)

    $headers=@("Tag","Datum","Betreuung","Person","Art","Fahrt","Std.","EUR/Std.","Zusatz","Kategorie","Notiz","Ausflug","Ausflug-km","Eintritt")
    $xs=@(10,82,165,248,365,475,535,590,660,735,825,990,1070,1145)
    $ws=@(68,78,78,105,100,55,45,60,65,80,155,75,65,60)
    for($i=0;$i -lt $headers.Count;$i++){ AddL $headers[$i] $xs[$i] 202 $ws[$i] | Out-Null }

    $checks=@()
    $dateLabels=@()
    $personBoxes=@()
    $typeBoxes=@()
    $travelBoxes=@()
    $hourBoxes=@()
    $rateBoxes=@()
    $costBoxes=@()
    $catBoxes=@()
    $noteBoxes=@()
    $tripBoxes=@()
    $tripKmBoxes=@()
    $entryBoxes=@()
    $dayNames=@("Montag","Dienstag","Mittwoch","Donnerstag","Freitag")

    for($i=0;$i -lt 5;$i++){
        $y=226+($i*42)

        $chk=New-Object System.Windows.Forms.CheckBox
        $chk.Text=$dayNames[$i];$chk.Location=[System.Drawing.Point]::new(10,$y);$chk.Size=[System.Drawing.Size]::new(72,22)
        $top.Controls.Add($chk);$checks+=$chk

        $dl=New-Object System.Windows.Forms.Label
        $dl.Text="";$dl.Location=[System.Drawing.Point]::new(82,($y+2));$dl.Size=[System.Drawing.Size]::new(80,22)
        $top.Controls.Add($dl);$dateLabels+=$dl

        $person=New-Object System.Windows.Forms.ComboBox
        $person.DropDownStyle="DropDown"
        $person.AutoCompleteMode="SuggestAppend"
        $person.AutoCompleteSource="CustomSource"
        $person.Location=[System.Drawing.Point]::new(248,$y);$person.Size=[System.Drawing.Size]::new(107,22)
        $person.Tag=$i
        $top.Controls.Add($person);$personBoxes+=$person

        $type=New-Object System.Windows.Forms.ComboBox
        $type.DropDownStyle="DropDownList"
        [void]$type.Items.Add("unentgeltlich")
        [void]$type.Items.Add("entgeltlich")
        $type.SelectedItem="unentgeltlich"
        $type.Location=[System.Drawing.Point]::new(365,$y);$type.Size=[System.Drawing.Size]::new(100,22)
        $top.Controls.Add($type);$typeBoxes+=$type

        $travel=New-Object System.Windows.Forms.CheckBox
        $travel.Checked=$true
        $travel.Location=[System.Drawing.Point]::new(490,$y)
        $travel.Size=[System.Drawing.Size]::new(22,22)
        $top.Controls.Add($travel);$travelBoxes+=$travel

        $hours=New-Object System.Windows.Forms.TextBox
        $hours.Text="0,00";$hours.Location=[System.Drawing.Point]::new(535,$y);$hours.Size=[System.Drawing.Size]::new(45,22)
        $top.Controls.Add($hours);$hourBoxes+=$hours

        $rateBox=New-Object System.Windows.Forms.TextBox
        $rateBox.Text="0,00";$rateBox.Location=[System.Drawing.Point]::new(590,$y);$rateBox.Size=[System.Drawing.Size]::new(60,22)
        $top.Controls.Add($rateBox);$rateBoxes+=$rateBox

        $cb=New-Object System.Windows.Forms.TextBox
        $cb.Text="0,00";$cb.Location=[System.Drawing.Point]::new(660,$y);$cb.Size=[System.Drawing.Size]::new(65,22)
        $top.Controls.Add($cb);$costBoxes+=$cb

        $cat=New-Object System.Windows.Forms.TextBox
        $cat.Location=[System.Drawing.Point]::new(735,$y);$cat.Size=[System.Drawing.Size]::new(80,22)
        $top.Controls.Add($cat);$catBoxes+=$cat

        $note=New-Object System.Windows.Forms.TextBox
        $note.Location=[System.Drawing.Point]::new(825,$y);$note.Size=[System.Drawing.Size]::new(155,22)
        $top.Controls.Add($note);$noteBoxes+=$note

        $trip=New-Object System.Windows.Forms.TextBox
        $trip.Location=[System.Drawing.Point]::new(990,$y);$trip.Size=[System.Drawing.Size]::new(70,22)
        $top.Controls.Add($trip);$tripBoxes+=$trip

        $tripKm=New-Object System.Windows.Forms.TextBox
        $tripKm.Text="0,00";$tripKm.Location=[System.Drawing.Point]::new(1070,$y);$tripKm.Size=[System.Drawing.Size]::new(65,22)
        $top.Controls.Add($tripKm);$tripKmBoxes+=$tripKm

        $entry=New-Object System.Windows.Forms.TextBox
        $entry.Text="0,00";$entry.Location=[System.Drawing.Point]::new(1145,$y);$entry.Size=[System.Drawing.Size]::new(60,22)
        $top.Controls.Add($entry);$entryBoxes+=$entry
    }

    $chkMo,$chkDi,$chkMi,$chkDo,$chkFr=$checks
    $lblMoDate,$lblDiDate,$lblMiDate,$lblDoDate,$lblFrDate=$dateLabels
    $txtMoCost,$txtDiCost,$txtMiCost,$txtDoCost,$txtFrCost=$costBoxes
    $txtMoCat,$txtDiCat,$txtMiCat,$txtDoCat,$txtFrCat=$catBoxes
    $txtMoNote,$txtDiNote,$txtMiNote,$txtDoNote,$txtFrNote=$noteBoxes
    $txtMoTrip,$txtDiTrip,$txtMiTrip,$txtDoTrip,$txtFrTrip=$tripBoxes
    $txtMoTripKm,$txtDiTripKm,$txtMiTripKm,$txtDoTripKm,$txtFrTripKm=$tripKmBoxes
    $txtMoEntry,$txtDiEntry,$txtMiEntry,$txtDoEntry,$txtFrEntry=$entryBoxes
    $cmbMoPerson,$cmbDiPerson,$cmbMiPerson,$cmbDoPerson,$cmbFrPerson=$personBoxes
    $cmbMoType,$cmbDiType,$cmbMiType,$cmbDoType,$cmbFrType=$typeBoxes
    $txtMoHours,$txtDiHours,$txtMiHours,$txtDoHours,$txtFrHours=$hourBoxes
    $txtMoRate,$txtDiRate,$txtMiRate,$txtDoRate,$txtFrRate=$rateBoxes
    $chkMoTravel,$chkDiTravel,$chkMiTravel,$chkDoTravel,$chkFrTravel=$travelBoxes

    $toolbar=New-Object System.Windows.Forms.Panel
    $toolbar.Dock="Top";$toolbar.Height=48
    $toolbar.BackColor=[System.Drawing.Color]::FromArgb(232,240,246)
    $form.Controls.Add($toolbar)

    $btnClear=New-Object System.Windows.Forms.Button
    $btnClear.Text="Auswahl leeren";$btnClear.Location=[System.Drawing.Point]::new(10,8);$btnClear.Size=[System.Drawing.Size]::new(120,30);$toolbar.Controls.Add($btnClear)
    $btnYear=New-Object System.Windows.Forms.Button
    $btnYear.Text="Jahr neu aufbauen";$btnYear.Location=[System.Drawing.Point]::new(145,8);$btnYear.Size=[System.Drawing.Size]::new(130,30);$toolbar.Controls.Add($btnYear)
    $btnSave=New-Object System.Windows.Forms.Button
    $btnSave.Text="Speichern";$btnSave.Location=[System.Drawing.Point]::new(290,8);$btnSave.Size=[System.Drawing.Size]::new(90,30);$toolbar.Controls.Add($btnSave)
    $btnExport=New-Object System.Windows.Forms.Button
    $btnExport.Text="Excel + PDF";$btnExport.Location=[System.Drawing.Point]::new(390,8);$btnExport.Size=[System.Drawing.Size]::new(105,30);$toolbar.Controls.Add($btnExport)

    $btnLastWeek=New-Object System.Windows.Forms.Button
    $btnLastWeek.Text="Letzte KW"
    $btnLastWeek.Location=[System.Drawing.Point]::new(505,8)
    $btnLastWeek.Size=[System.Drawing.Size]::new(90,30)
    $toolbar.Controls.Add($btnLastWeek)

    $lblMultiHint=New-Object System.Windows.Forms.Label
    $lblMultiHint.Text="Mehrere KWs mit STRG/SHIFT markieren - Haken oben gelten dann für alle markierten KWs."
    $lblMultiHint.Location=[System.Drawing.Point]::new(610,12);$lblMultiHint.Size=[System.Drawing.Size]::new(550,22)
    $toolbar.Controls.Add($lblMultiHint)

    $contentLayout=New-Object System.Windows.Forms.TableLayoutPanel
    $contentLayout.Dock="Fill"
    $contentLayout.ColumnCount=1
    $contentLayout.RowCount=2
    [void]$contentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))
    [void]$contentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,100)))
    [void]$contentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute,38)))
    $form.Controls.Add($contentLayout)
    $contentLayout.BringToFront()

    $grid=New-Object System.Windows.Forms.DataGridView
    $grid.Dock="Fill";$grid.AllowUserToAddRows=$false;$grid.AllowUserToDeleteRows=$false
    $grid.ReadOnly=$true;$grid.SelectionMode="FullRowSelect";$grid.MultiSelect=$true;$grid.RowHeadersVisible=$false
    $grid.AutoSizeColumnsMode="Fill";$grid.BackgroundColor=[System.Drawing.SystemColors]::Window
    $grid.EnableHeadersVisualStyles=$false
    $grid.ColumnHeadersDefaultCellStyle.BackColor=[System.Drawing.Color]::FromArgb(220,232,240)
    $grid.ColumnHeadersDefaultCellStyle.ForeColor=[System.Drawing.Color]::FromArgb(40,55,65)
    $grid.AlternatingRowsDefaultCellStyle.BackColor=[System.Drawing.Color]::FromArgb(248,251,253)
    $grid.DefaultCellStyle.SelectionBackColor=[System.Drawing.Color]::FromArgb(184,214,232)
    $grid.DefaultCellStyle.SelectionForeColor=[System.Drawing.Color]::Black
    $grid.Margin=New-Object System.Windows.Forms.Padding(0)
    $contentLayout.Controls.Add($grid,0,0)

    foreach($x in @(
        @("KW","Kalenderwoche",90),
        @("Zeitraum","Zeitraum",160),
        @("Anzahl","Betreuungstage",90),
        @("Tage","Markierte Tage",170)
    )){
        $c=New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $c.Name=$x[0];$c.HeaderText=$x[1];$c.FillWeight=$x[2]
        [void]$grid.Columns.Add($c)
    }

    $bottom=New-Object System.Windows.Forms.Panel
    $bottom.Dock="Fill"
    $bottom.BackColor=[System.Drawing.Color]::FromArgb(232,240,246)
    $bottom.Margin=New-Object System.Windows.Forms.Padding(0)
    $contentLayout.Controls.Add($bottom,0,1)

    $statusLayout=New-Object System.Windows.Forms.TableLayoutPanel
    $statusLayout.Dock="Fill"
    $statusLayout.ColumnCount=2
    $statusLayout.RowCount=1
    $statusLayout.Margin=New-Object System.Windows.Forms.Padding(0)
    [void]$statusLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)))
    [void]$statusLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute,180)))
    [void]$statusLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,100)))
    $bottom.Controls.Add($statusLayout)

    $lblSummary=New-Object System.Windows.Forms.Label
    $lblSummary.Dock="Fill"
    $lblSummary.TextAlign="MiddleLeft"
    $lblSummary.Padding=New-Object System.Windows.Forms.Padding(10,0,0,0)
    $lblSummary.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
    $statusLayout.Controls.Add($lblSummary,0,0)

    $lblSignature=New-Object System.Windows.Forms.Label
    $lblSignature.Text="By Lorenz Köcke"
    $lblSignature.Dock="Fill"
    $lblSignature.TextAlign="MiddleRight"
    $lblSignature.Padding=New-Object System.Windows.Forms.Padding(0,0,12,0)
    $lblSignature.ForeColor=[System.Drawing.Color]::DimGray
    $lblSignature.Cursor=[System.Windows.Forms.Cursors]::Hand
    $statusLayout.Controls.Add($lblSignature,1,0)

    $script:SignatureClicks=0
    $script:LastSignatureClick=[datetime]::MinValue

    $lblSignature.Add_Click({
        $now = Get-Date
        if (($now - $script:LastSignatureClick).TotalMilliseconds -le 700) {
            $script:SignatureClicks++
        } else {
            $script:SignatureClicks=1
        }
        $script:LastSignatureClick=$now

        if ($script:SignatureClicks -ge 3) {
            $script:SignatureClicks=0
            Generate-DemoData
        }
    })

    # Betreuungshaken: jeder Wochentag besitzt bewusst einen eigenen Handler.
    # Der Klick wird sofort in das Datenmodell geschrieben und die KW-Zeile aktualisiert.
    $chkMo.Add_CheckedChanged({
        if(-not $script:LoadingDetail){
            Apply-DayCheckToSelectedWeeks 0 $chkMo.Checked
        }
    })

    $chkDi.Add_CheckedChanged({
        if(-not $script:LoadingDetail){
            Apply-DayCheckToSelectedWeeks 1 $chkDi.Checked
        }
    })

    $chkMi.Add_CheckedChanged({
        if(-not $script:LoadingDetail){
            Apply-DayCheckToSelectedWeeks 2 $chkMi.Checked
        }
    })

    $chkDo.Add_CheckedChanged({
        if(-not $script:LoadingDetail){
            Apply-DayCheckToSelectedWeeks 3 $chkDo.Checked
        }
    })

    $chkFr.Add_CheckedChanged({
        if(-not $script:LoadingDetail){
            Apply-DayCheckToSelectedWeeks 4 $chkFr.Checked
        }
    })

    $txtCare.Add_Leave({
        if(-not $script:LoadingDetail){
            Refresh-PersonSuggestions
        }
    })

    $txtVacation.Add_TextChanged({ if(-not $script:LoadingDetail){ $script:IsDirty=$true } })
    $chkHolidays.Add_CheckedChanged({ if(-not $script:LoadingDetail){ $script:IsDirty=$true } })
    $chkContract.Add_CheckedChanged({ if(-not $script:LoadingDetail){ $script:IsDirty=$true } })
    $txtIban.Add_TextChanged({ if(-not $script:LoadingDetail){ $script:IsDirty=$true } })
    $rbPaid.Add_CheckedChanged({
        $numFee.Enabled=$rbPaid.Checked
        if(-not $script:LoadingDetail){ $script:IsDirty=$true }
    })
    $rbUnpaid.Add_CheckedChanged({
        if($rbUnpaid.Checked){ $numFee.Enabled=$false }
        if(-not $script:LoadingDetail){ $script:IsDirty=$true }
    })
    $numFee.Add_ValueChanged({ if(-not $script:LoadingDetail){ $script:IsDirty=$true } })


    # ------------------------------------------------------------
    # Aenderungsueberwachung
    # ------------------------------------------------------------
    $markDirty = {
        if(-not $script:LoadingDetail){
            $script:IsDirty = $true
        }
    }

    function Register-DirtyTracking {
        param([System.Windows.Forms.Control]$Parent)

        foreach($control in $Parent.Controls){
            if($control -is [System.Windows.Forms.TextBox]){
                $control.Add_TextChanged($markDirty)
            }
            elseif($control -is [System.Windows.Forms.ComboBox]){
                $control.Add_TextChanged($markDirty)
                $control.Add_SelectedIndexChanged($markDirty)
            }
            elseif($control -is [System.Windows.Forms.CheckBox] -or
                   $control -is [System.Windows.Forms.RadioButton]){
                $control.Add_CheckedChanged($markDirty)
            }
            elseif($control -is [System.Windows.Forms.NumericUpDown]){
                $control.Add_ValueChanged($markDirty)
            }

            if($control.HasChildren){
                Register-DirtyTracking -Parent $control
            }
        }
    }

    # Keine fest verdrahteten Array-Namen: alle editierbaren Controls des
    # Fensters werden automatisch erfasst. Dadurch bleibt die Ueberwachung
    # auch bei spaeteren UI-Erweiterungen stabil.
    Register-DirtyTracking -Parent $form

    $grid.Add_SelectionChanged({
        if($grid.CurrentRow -and $grid.CurrentRow.Index -ne $script:CurrentIndex){
            Load-WeekDetail $grid.CurrentRow.Index
        }
    })

    $btnClear.Add_Click({
        if($script:LoadingDetail){ return }

        # Aktuelle Eingaben sichern, bevor die markierten Wochen geloescht werden.
        Save-CurrentDetail

        $selected = @($grid.SelectedRows)
        if($selected.Count -eq 0 -and $script:CurrentIndex -ge 0){
            $selected = @($grid.Rows[$script:CurrentIndex])
        }

        foreach($row in $selected){
            if($null -eq $row -or $row.Index -lt 0 -or $row.Index -ge $script:Weeks.Count){ continue }

            $week = $script:Weeks[$row.Index]
            foreach($day in $week.Days){
                $day.Checked = $false
                $day.Person = ""
                $day.Art = "unentgeltlich"
                $day.Fahrtkosten = $false
                $day.Stunden = "0,00"
                $day.Stundenlohn = "0,00"
                $day.Kosten = "0,00"
                $day.Kategorie = ""
                $day.Notiz = ""
                $day.AusflugZiel = ""
                $day.AusflugKm = "0,00"
                $day.Eintritt = "0,00"
            }

            Refresh-WeekRow $row.Index
        }

        # Die Detailansicht neu laden, OHNE vorher die alten Controls erneut
        # in die gerade geloeschten Daten zurueckzuschreiben.
        if($script:CurrentIndex -ge 0 -and $script:CurrentIndex -lt $script:Weeks.Count){
            $script:LoadingDetail = $true
            Load-WeekDetail $script:CurrentIndex
        }

        $script:IsDirty = $true
        Update-Summary
    })

    $btnYear.Add_Click({
        New-YearData
    })

    $btnSave.Add_Click({
        try {
            Save-CurrentDetail

            if(Save-Data){
                $script:IsDirty = $false
                $info = Get-Item $DataFile
                Show-Info (
                    "Gespeichert.`r`n`r`n" +
                    "Datei:`r`n" + $DataFile + "`r`n`r`n" +
                    ("Groesse: {0:N0} Bytes`r`n" -f $info.Length) +
                    ("Zeit: {0:dd.MM.yyyy HH:mm:ss}" -f $info.LastWriteTime)
                )
            }
        }
        catch {
            Show-Error $_.Exception.Message
        }
    })

    $btnExport.Add_Click({
        Export-ExcelPdf
    })

    $btnLastWeek.Add_Click({
        if($grid.Rows.Count -le 0){ return }

        # Noch offene Eingaben der aktuellen KW sichern.
        Save-CurrentDetail

        $last = $grid.Rows.Count - 1

        # SelectionChanged waehrend des Umschaltens kontrolliert unterdruecken,
        # damit nicht versehentlich die vorherige KW erneut geladen wird.
        $script:LoadingDetail = $true
        $grid.ClearSelection()
        $grid.Rows[$last].Selected = $true
        $grid.CurrentCell = $grid.Rows[$last].Cells[0]
        $grid.FirstDisplayedScrollingRowIndex = [math]::Max(0,$last)
        $script:LoadingDetail = $false

        Load-WeekDetail $last
    })

    $form.Add_FormClosing({
        param($sender,$e)

        if($script:ClosingConfirmed){ return }

        # Noch sichtbare Eingaben in das Datenmodell uebernehmen, ohne automatisch
        # zu speichern. Dadurch erkennt die Rueckfrage auch die letzte Aenderung.
        Save-CurrentDetail

        if(-not $script:IsDirty){
            return
        }

        $answer = [System.Windows.Forms.MessageBox]::Show(
            "Es gibt noch nicht gespeicherte Aenderungen.`r`n`r`nMoechten Sie diese vor dem Beenden speichern?",
            "Kinderbetreuung - Aenderungen speichern?",
            [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button1
        )

        switch($answer){
            ([System.Windows.Forms.DialogResult]::Yes) {
                try {
                    $saved = Save-Data
                    if(-not $saved){
                        throw "Speichervorgang wurde nicht bestaetigt."
                    }

                    $script:IsDirty = $false
                    $script:ClosingConfirmed = $true
                }
                catch {
                    $script:ClosingConfirmed = $false
                    $e.Cancel = $true
                    Show-Error $_.Exception.Message
                }
            }
            ([System.Windows.Forms.DialogResult]::No) {
                $script:ClosingConfirmed = $true
            }
            default {
                # Abbrechen: Fenster bleibt offen.
                $e.Cancel = $true
            }
        }
    })

    # Gespeicherte Daten laden, nachdem alle Controls und Events existieren.

    # ------------------------------------------------------------
    # Modernes Erscheinungsbild
    # ------------------------------------------------------------
    $form.BackColor = $clrWindow
    $form.ForeColor = $clrText
    $form.Font = New-Object System.Drawing.Font("Segoe UI",9)

    # Oberer Bereich als ruhige Flaeche.
    if($null -ne $toolbar){
        $toolbar.BackColor = $clrSurfaceSoft
    }
    if($null -ne $top){
        $top.BackColor = $clrWindow
    }

    function Set-ModernButton {
        param(
            [System.Windows.Forms.Button]$Button,
            [bool]$Primary = $false
        )
        if($null -eq $Button){ return }

        $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $Button.FlatAppearance.BorderSize = 1
        $Button.FlatAppearance.MouseOverBackColor = $clrPrimaryHover
        $Button.FlatAppearance.MouseDownBackColor = $clrPrimarySoft
        $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
        $Button.Font = New-Object System.Drawing.Font("Segoe UI",9)

        if($Primary){
            $Button.BackColor = $clrPrimary
            $Button.ForeColor = $clrWhite
            $Button.FlatAppearance.BorderColor = $clrPrimaryDark
            $Button.FlatAppearance.MouseOverBackColor = $clrPrimaryDark
            $Button.FlatAppearance.MouseDownBackColor = $clrPrimaryDark
            $Button.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
        } else {
            $Button.BackColor = $clrSurface
            $Button.ForeColor = $clrText
            $Button.FlatAppearance.BorderColor = $clrBorder
        }
    }

    # Alle Toolbar-Buttons konsistent stylen.
    Set-ModernButton $btnClear
    Set-ModernButton $btnYear
    Set-ModernButton $btnSave $true
    Set-ModernButton $btnExport
    Set-ModernButton $btnLastWeek

    # Etwas mehr "App"-Gefuehl durch gleichmaessige Buttonhoehe.
    foreach($b in @($btnClear,$btnYear,$btnSave,$btnExport,$btnLastWeek)){
        if($null -ne $b){
            $b.Height = 32
        }
    }

    # Haupttabelle moderner und ruhiger.
    $grid.BackgroundColor = $clrSurface
    $grid.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $grid.GridColor = $clrBorder
    $grid.EnableHeadersVisualStyles = $false
    $grid.ColumnHeadersBorderStyle = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::Single
    $grid.ColumnHeadersDefaultCellStyle.BackColor = $clrSurfaceSoft
    $grid.ColumnHeadersDefaultCellStyle.ForeColor = $clrText
    $grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
    $grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = $clrSurfaceSoft
    $grid.ColumnHeadersHeight = 28

    $grid.DefaultCellStyle.BackColor = $clrSurface
    $grid.DefaultCellStyle.ForeColor = $clrText
    $grid.DefaultCellStyle.SelectionBackColor = $clrPrimarySoft
    $grid.DefaultCellStyle.SelectionForeColor = $clrText
    $grid.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(2,1,2,1)
    $grid.RowTemplate.Height = 24
    $grid.AlternatingRowsDefaultCellStyle.BackColor = $clrGridAlt

    # Statusleiste als eigenstaendige Akzentflaeche.
    if($null -ne $bottom){
        $bottom.BackColor = $clrSurfaceSoft
    }
    if($null -ne $lblSummary){
        $lblSummary.BackColor = $clrSurfaceSoft
        $lblSummary.ForeColor = $clrText
        $lblSummary.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
    }
    if($null -ne $lblSignature){
        $lblSignature.BackColor = $clrSurfaceSoft
        $lblSignature.ForeColor = $clrPrimary
        $lblSignature.Font = New-Object System.Drawing.Font("Segoe UI",8)
    }

    # Wochenueberschrift etwas praesenter.
    if($null -ne $lblSelectedWeek){
        $lblSelectedWeek.ForeColor = $clrPrimaryDark
        $lblSelectedWeek.Font = New-Object System.Drawing.Font("Segoe UI",11,[System.Drawing.FontStyle]::Bold)
    }

    # Hinweistext im Toolbar-Bereich dezenter.
    if($null -ne $lblMultiHint){
        $lblMultiHint.ForeColor = $clrTextMuted
    }

    # Eingabefelder leicht vereinheitlichen.
    foreach($ctrl in @(
        $txtChild,$txtCare,$txtVacation,$txtIban,
        $cmbMoPerson,$cmbDiPerson,$cmbMiPerson,$cmbDoPerson,$cmbFrPerson,
        $cmbMoType,$cmbDiType,$cmbMiType,$cmbDoType,$cmbFrType
    )){
        if($null -ne $ctrl){
            $ctrl.Font = New-Object System.Drawing.Font("Segoe UI",9)
        }
    }

    Load-Data

    # Der erfolgreich geladene Zustand ist die saubere Ausgangsbasis.
    $script:IsDirty = $false
    $script:ClosingConfirmed = $false

    [void]$form.ShowDialog()
    exit 0
}
catch{
    Write-StartError $_.Exception
    [System.Windows.Forms.MessageBox]::Show(
        "Startfehler:`r`n$($_.Exception.Message)`r`n`r`nDetails: $ErrorFile",
        "Kinderbetreuung","OK","Error")|Out-Null
    exit 1
}
