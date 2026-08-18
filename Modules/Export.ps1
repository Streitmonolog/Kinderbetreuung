# Kinderbetreuung 1.0.17 - Export
# XLSX/PDF-Ausgabe ueber Microsoft Excel oder LibreOffice.

function Choose-ExportMode {
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Export auswählen"
    $dialog.Size = [System.Drawing.Size]::new(430,210)
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = "FixedDialog"
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Welche Auswertung möchtest du erstellen?"
    $lbl.Location = [System.Drawing.Point]::new(20,20)
    $lbl.Size = [System.Drawing.Size]::new(370,30)
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
    $dialog.Controls.Add($lbl)

    $btnTax = New-Object System.Windows.Forms.Button
    $btnTax.Text = "Nur Steuer"
    $btnTax.Location = [System.Drawing.Point]::new(20,70)
    $btnTax.Size = [System.Drawing.Size]::new(170,55)
    $dialog.Controls.Add($btnTax)

    $btnFull = New-Object System.Windows.Forms.Button
    $btnFull.Text = "Vollständige Dokumentation"
    $btnFull.Location = [System.Drawing.Point]::new(210,70)
    $btnFull.Size = [System.Drawing.Size]::new(180,55)
    $dialog.Controls.Add($btnFull)

    $hint = New-Object System.Windows.Forms.Label
    $hint.Text = "Nur Steuer enthält ausschließlich steuerlich relevante Betreuungskosten."
    $hint.Location = [System.Drawing.Point]::new(20,138)
    $hint.Size = [System.Drawing.Size]::new(370,35)
    $dialog.Controls.Add($hint)

    $script:ExportMode = $null
    $btnTax.Add_Click({ $script:ExportMode = "tax"; $dialog.Close() })
    $btnFull.Add_Click({ $script:ExportMode = "full"; $dialog.Close() })

    [void]$dialog.ShowDialog($form)
    return $script:ExportMode
}

function Select-ExportFolder {
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Ordner fuer Excel- und PDF-Export auswaehlen"
    $dialog.ShowNewFolderButton = $true

    $defaultExport = Join-Path $AppDir "Export"
    if (-not (Test-Path $defaultExport)) {
        New-Item -ItemType Directory -Path $defaultExport -Force | Out-Null
    }
    $dialog.SelectedPath = $defaultExport

    $result = $dialog.ShowDialog($form)
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        return $null
    }
    return $dialog.SelectedPath
}

function Find-LibreOffice {
    $candidates = @(
        (Join-Path $env:ProgramFiles "LibreOffice\program\soffice.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "LibreOffice\program\soffice.exe"),
        (Join-Path $env:LOCALAPPDATA "Programs\LibreOffice\program\soffice.exe")
    )

    foreach($candidate in $candidates){
        if(-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path $candidate)){
            return $candidate
        }
    }

    try {
        $cmd = Get-Command soffice.exe -ErrorAction Stop
        if($cmd -and $cmd.Source){ return $cmd.Source }
    } catch {}

    return $null
}

function Test-ExcelAvailable {
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Quit()
        [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        return $true
    } catch {
        return $false
    }
}

function Get-OfficeBackend {
    if(Test-ExcelAvailable){
        return [pscustomobject]@{ Type="Excel"; Path=$null }
    }

    $lo = Find-LibreOffice
    if($lo){
        return [pscustomobject]@{ Type="LibreOffice"; Path=$lo }
    }

    return [pscustomobject]@{ Type="None"; Path=$null }
}

function Convert-WithLibreOffice {
    param(
        [string]$SofficePath,
        [string]$InputFile,
        [string]$OutputDir,
        [string]$Format
    )

    if(-not (Test-Path $SofficePath)){
        throw "LibreOffice wurde nicht gefunden."
    }
    if(-not (Test-Path $InputFile)){
        throw "Quelldatei fuer LibreOffice wurde nicht gefunden: $InputFile"
    }

    $args = @(
        "--headless",
        "--nologo",
        "--nodefault",
        "--nofirststartwizard",
        "--convert-to", $Format,
        "--outdir", $OutputDir,
        $InputFile
    )

    $p = Start-Process -FilePath $SofficePath -ArgumentList $args -PassThru -Wait -WindowStyle Hidden
    if($p.ExitCode -ne 0){
        throw "LibreOffice-Konvertierung ist mit Fehlercode $($p.ExitCode) fehlgeschlagen."
    }
}

function Escape-Xml([string]$Value) {
    if($null -eq $Value){ return "" }
    return [System.Security.SecurityElement]::Escape([string]$Value)
}

function Build-SpreadsheetMLReport {
    param(
        [string]$Path,
        [string]$Mode
    )

    Save-CurrentDetail

    $year=[int]$numYear.Value
    $kmOne=[double]$numKm.Value
    $rate=[double]$numRate.Value
    $names=@("Montag","Dienstag","Mittwoch","Donnerstag","Freitag")
    $taxTotal=0.0

    $quarterData=@{
        1=[pscustomobject]@{Days=0;Travel=0.0;Fee=0.0;Total=0.0}
        2=[pscustomobject]@{Days=0;Travel=0.0;Fee=0.0;Total=0.0}
        3=[pscustomobject]@{Days=0;Travel=0.0;Fee=0.0;Total=0.0}
        4=[pscustomobject]@{Days=0;Travel=0.0;Fee=0.0;Total=0.0}
    }

    $rows = New-Object System.Collections.Generic.List[string]

    $title = if($Mode -eq "tax"){"Steuer-Auswertung Kinderbetreuung $year"}else{"Vollstaendige Dokumentation Kinderbetreuung $year"}

    $rows.Add("<Row><Cell ss:MergeAcross=`"10`"><Data ss:Type=`"String`">$(Escape-Xml $title)</Data></Cell></Row>")
    $rows.Add("<Row><Cell><Data ss:Type=`"String`">Kind</Data></Cell><Cell><Data ss:Type=`"String`">$(Escape-Xml $txtChild.Text)</Data></Cell></Row>")
    $rows.Add("<Row><Cell><Data ss:Type=`"String`">Betreuungsperson(en)</Data></Cell><Cell><Data ss:Type=`"String`">$(Escape-Xml $txtCare.Text)</Data></Cell></Row>")
    $rows.Add("<Row><Cell><Data ss:Type=`"String`">Schriftliche Vereinbarung / Betreuungsvertrag</Data></Cell><Cell><Data ss:Type=`"String`">$(if($chkContract.Checked){'Ja'}else{'Nein'})</Data></Cell></Row>")
    if($Mode -eq "full"){
        $rows.Add("<Row><Cell><Data ss:Type=`"String`">IBAN Betreuungsperson</Data></Cell><Cell><Data ss:Type=`"String`">$(Escape-Xml $txtIban.Text)</Data></Cell></Row>")
    }
    $rows.Add("<Row/>")

    if($Mode -eq "tax"){
        $headers=@("Datum","Wochentag","KW","Betreuungsperson","Art","Fahrtkosten","Stunden","EUR/Stunde","km gesamt","Fahrtkosten EUR","Betreuungsentgelt EUR","Steuerlich relevante Kosten EUR")
    } else {
        $headers=@("Datum","Wochentag","KW","Betreuungsperson","Art","Fahrtkosten","Stunden","EUR/Stunde","km gesamt","Fahrtkosten EUR","Betreuungsentgelt EUR","Zusatzkosten/Doku EUR","Kategorie","Notiz","Ausflug/Ziel","Ausflug-km","Eintritt/Freizeit EUR","Hinweis")
    }

    $headerCells = ($headers | ForEach-Object {"<Cell ss:StyleID=`"Header`"><Data ss:Type=`"String`">$(Escape-Xml $_)</Data></Cell>"}) -join ""
    $rows.Add("<Row>$headerCells</Row>")

    foreach($w in $script:Weeks){
        for($d=0;$d -lt $w.Days.Count;$d++){
            $day=$w.Days[$d]
            if(-not [bool]$day.Checked){ continue }

            $dt=[datetime]::ParseExact([string]$day.Date,"yyyy-MM-dd",$null)
            if($dt.Year -ne $year){ continue }

            $person=if([string]::IsNullOrWhiteSpace([string]$day.Person)){$txtCare.Text}else{[string]$day.Person}
            $type=[string]$day.Art
            if($type -notin @("unentgeltlich","entgeltlich")){$type="unentgeltlich"}
            $hours=if($type -eq "entgeltlich"){Parse-DecimalDE ([string]$day.Stunden)}else{0.0}
            $hourly=if($type -eq "entgeltlich"){Parse-DecimalDE ([string]$day.Stundenlohn)}else{0.0}
            $fee=[math]::Round($hours*$hourly,2)
            $kmTotal=if([bool]$day.Fahrtkosten){$kmOne*2}else{0.0}
            $travelCost=if([bool]$day.Fahrtkosten){$kmTotal*$rate}else{0.0}
            $taxRelevant=$travelCost+$fee
            $taxTotal += $taxRelevant

            $q=[int][math]::Floor(($dt.Month-1)/3)+1
            $quarterData[$q].Days++
            $quarterData[$q].Travel += $travelCost
            $quarterData[$q].Fee += $fee
            $quarterData[$q].Total += $taxRelevant

            $cells=@()
            $cells += "<Cell><Data ss:Type=`"String`">$($dt.ToString("dd.MM.yyyy"))</Data></Cell>"
            $cells += "<Cell><Data ss:Type=`"String`">$(Escape-Xml $names[$d])</Data></Cell>"
            $cells += "<Cell><Data ss:Type=`"String`">KW $($w.Week)</Data></Cell>"
            $cells += "<Cell><Data ss:Type=`"String`">$(Escape-Xml $person)</Data></Cell>"
            $cells += "<Cell><Data ss:Type=`"String`">$(Escape-Xml $type)</Data></Cell>"
            $cells += "<Cell><Data ss:Type=`"String`">$(if([bool]$day.Fahrtkosten){'Ja'}else{'Nein'})</Data></Cell>"
            $cells += "<Cell ss:StyleID=`"Number`"><Data ss:Type=`"Number`">$hours</Data></Cell>"
            $cells += "<Cell ss:StyleID=`"Currency`"><Data ss:Type=`"Number`">$hourly</Data></Cell>"
            $cells += "<Cell ss:StyleID=`"Number`"><Data ss:Type=`"Number`">$kmTotal</Data></Cell>"
            $cells += "<Cell ss:StyleID=`"Currency`"><Data ss:Type=`"Number`">$travelCost</Data></Cell>"
            $cells += "<Cell ss:StyleID=`"Currency`"><Data ss:Type=`"Number`">$fee</Data></Cell>"

            if($Mode -eq "tax"){
                $cells += "<Cell ss:StyleID=`"Currency`"><Data ss:Type=`"Number`">$taxRelevant</Data></Cell>"
            } else {
                $extra=Parse-DecimalDE ([string]$day.Kosten)
                $tripKm=Parse-DecimalDE ([string]$day.AusflugKm)
                $entry=Parse-DecimalDE ([string]$day.Eintritt)
                $cells += "<Cell ss:StyleID=`"Currency`"><Data ss:Type=`"Number`">$extra</Data></Cell>"
                $cells += "<Cell><Data ss:Type=`"String`">$(Escape-Xml ([string]$day.Kategorie))</Data></Cell>"
                $cells += "<Cell><Data ss:Type=`"String`">$(Escape-Xml ([string]$day.Notiz))</Data></Cell>"
                $cells += "<Cell><Data ss:Type=`"String`">$(Escape-Xml ([string]$day.AusflugZiel))</Data></Cell>"
                $cells += "<Cell ss:StyleID=`"Number`"><Data ss:Type=`"Number`">$tripKm</Data></Cell>"
                $cells += "<Cell ss:StyleID=`"Currency`"><Data ss:Type=`"Number`">$entry</Data></Cell>"
                $hint=if($tripKm -gt 0 -or $entry -gt 0 -or -not [string]::IsNullOrWhiteSpace([string]$day.AusflugZiel)){"Freizeit-/Ausflugskosten: nur Dokumentation, nicht in Steuer-Summe berücksichtigt"}else{""}
                $cells += "<Cell><Data ss:Type=`"String`">$(Escape-Xml $hint)</Data></Cell>"
            }

            $rows.Add("<Row>$($cells -join '')</Row>")
        }
    }

    $deductible=[math]::Round([math]::Min($taxTotal*0.80,4800.0),2)
    $rows.Add("<Row/>")
    $rows.Add("<Row><Cell ss:StyleID=`"Header`"><Data ss:Type=`"String`">Steuerlich relevante Kosten gesamt</Data></Cell><Cell ss:StyleID=`"Currency`"><Data ss:Type=`"Number`">$taxTotal</Data></Cell></Row>")
    $rows.Add("<Row><Cell ss:StyleID=`"Header`"><Data ss:Type=`"String`">Davon 80 % (max. 4.800 EUR)</Data></Cell><Cell ss:StyleID=`"Currency`"><Data ss:Type=`"Number`">$deductible</Data></Cell></Row>")
    $rows.Add("<Row/>")
    $rows.Add("<Row><Cell ss:StyleID=`"Title`"><Data ss:Type=`"String`">Quartalsuebersicht</Data></Cell></Row>")
    $rows.Add("<Row><Cell ss:StyleID=`"Header`"><Data ss:Type=`"String`">Quartal</Data></Cell><Cell ss:StyleID=`"Header`"><Data ss:Type=`"String`">Betreuungstage</Data></Cell><Cell ss:StyleID=`"Header`"><Data ss:Type=`"String`">Fahrtkosten EUR</Data></Cell><Cell ss:StyleID=`"Header`"><Data ss:Type=`"String`">Betreuungsentgelt EUR</Data></Cell><Cell ss:StyleID=`"Header`"><Data ss:Type=`"String`">Zu zahlen / erstatten EUR</Data></Cell></Row>")

    for($q=1;$q -le 4;$q++){
        $qd=$quarterData[$q]
        $rows.Add("<Row><Cell><Data ss:Type=`"String`">Q$q</Data></Cell><Cell ss:StyleID=`"Number`"><Data ss:Type=`"Number`">$($qd.Days)</Data></Cell><Cell ss:StyleID=`"Currency`"><Data ss:Type=`"Number`">$([math]::Round($qd.Travel,2))</Data></Cell><Cell ss:StyleID=`"Currency`"><Data ss:Type=`"Number`">$([math]::Round($qd.Fee,2))</Data></Cell><Cell ss:StyleID=`"Currency`"><Data ss:Type=`"Number`">$([math]::Round($qd.Total,2))</Data></Cell></Row>")
    }

    $xml = @"
<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Bottom"/>
   <Font ss:FontName="Calibri" ss:Size="11"/>
  </Style>
  <Style ss:ID="Title">
   <Font ss:FontName="Calibri" ss:Size="14" ss:Bold="1"/>
  </Style>
  <Style ss:ID="Header">
   <Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1"/>
   <Interior ss:Color="#DCE8F0" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="Currency">
   <NumberFormat ss:Format="#,##0.00 [$EUR]"/>
  </Style>
  <Style ss:ID="Number">
   <NumberFormat ss:Format="#,##0.00"/>
  </Style>
 </Styles>
 <Worksheet ss:Name="Kinderbetreuung">
  <Table>
   $($rows -join "`r`n")
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Layout x:Orientation="Landscape"/>
   </PageSetup>
   <FitToPage/>
   <Print>
    <FitWidth>1</FitWidth>
    <ValidPrinterInfo/>
    <HorizontalResolution>600</HorizontalResolution>
    <VerticalResolution>600</VerticalResolution>
   </Print>
   <Selected/>
  </WorksheetOptions>
 </Worksheet>
</Workbook>
"@

    Set-Content -Path $Path -Value $xml -Encoding UTF8
}

function Export-WithLibreOffice {
    param(
        [string]$SofficePath,
        [string]$Mode,
        [string]$OutputDir
    )

    $year=[int]$numYear.Value
    $safe=($txtChild.Text -replace '[\\/:*?"<>|]','_')
    if([string]::IsNullOrWhiteSpace($safe)){$safe="Kind"}
    $suffix=if($Mode -eq "tax"){"Steuer"}else{"Vollstaendig"}

    $tempDir=Join-Path $env:TEMP ("Kinderbetreuung_LO_" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

    try {
        $xmlPath=Join-Path $tempDir "Kinderbetreuung.xml"
        Build-SpreadsheetMLReport -Path $xmlPath -Mode $Mode

        # SpreadsheetML -> XLSX
        Convert-WithLibreOffice -SofficePath $SofficePath -InputFile $xmlPath -OutputDir $tempDir -Format 'xlsx:"Calc MS Excel 2007 XML"'
        $generatedXlsx=Join-Path $tempDir "Kinderbetreuung.xlsx"
        if(-not (Test-Path $generatedXlsx)){
            # LibreOffice may keep original basename from XML.
            $candidate=Get-ChildItem $tempDir -Filter *.xlsx | Select-Object -First 1
            if($candidate){$generatedXlsx=$candidate.FullName}
        }
        if(-not (Test-Path $generatedXlsx)){ throw "LibreOffice konnte keine XLSX-Datei erzeugen." }

        $xlsx=Join-Path $OutputDir "Kinderbetreuung_${safe}_${year}_${suffix}.xlsx"
        Copy-Item $generatedXlsx $xlsx -Force

        # XLSX -> PDF
        Convert-WithLibreOffice -SofficePath $SofficePath -InputFile $xlsx -OutputDir $OutputDir -Format 'pdf:calc_pdf_Export'
        $generatedPdf=Join-Path $OutputDir ([System.IO.Path]::GetFileNameWithoutExtension($xlsx)+".pdf")

        if(-not (Test-Path $generatedPdf)){ throw "LibreOffice konnte keine PDF-Datei erzeugen." }

        return [pscustomobject]@{Xlsx=$xlsx;Pdf=$generatedPdf}
    }
    finally {
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Export-ExcelPdf {
    try {
        if(Save-Data){
            $script:IsDirty = $false
        }
        $mode = Choose-ExportMode
        if ([string]::IsNullOrWhiteSpace($mode)) { return }

        $out = Select-ExportFolder
        if ([string]::IsNullOrWhiteSpace($out)) { return }

        if ($mode -eq "tax" -and $rbPaid.Checked -and [double]$numFee.Value -le 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Betreuung gegen Entgelt ist ausgewählt, aber das Entgelt je Betreuungstag ist 0,00 EUR. Im Steuerexport wird daher nur die Fahrtkostenerstattung berücksichtigt.",
                "Hinweis",
                "OK",
                "Warning"
            ) | Out-Null
        }

        if ($mode -eq "tax" -and -not $chkContract.Checked) {
            [System.Windows.Forms.MessageBox]::Show(
                "Hinweis: Bei Kinderbetreuung durch Angehörige sollten klare und eindeutige Vereinbarungen vorliegen und tatsächlich so durchgeführt werden. Für die steuerliche Berücksichtigung sind außerdem geeignete Nachweise sowie eine unbare Zahlung auf das Konto der Betreuungsperson erforderlich.`r`n`r`nDie Steuer-Auswertung wird trotzdem erstellt.",
                "Hinweis zur steuerlichen Dokumentation",
                "OK",
                "Warning"
            ) | Out-Null
        }

        $backend = Get-OfficeBackend
        if($backend.Type -eq "None"){
            throw "Es wurde weder Microsoft Excel noch LibreOffice gefunden. Fuer XLSX/PDF-Export bitte Microsoft Excel oder LibreOffice installieren."
        }

        if($backend.Type -eq "LibreOffice"){
            $result = Export-WithLibreOffice -SofficePath $backend.Path -Mode $mode -OutputDir $out
            Show-Info ("Excel und PDF wurden mit LibreOffice erstellt.`r`n`r`n{0}`r`n{1}" -f $result.Xlsx,$result.Pdf)
            Start-Process explorer.exe $out
            return
        }

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible=$false
        $excel.DisplayAlerts=$false
        $wb=$excel.Workbooks.Add()
        $ws=$wb.Worksheets.Item(1)
        $ws.Name="Kinderbetreuung"
        $year=[int]$numYear.Value

        $title = if ($mode -eq "tax") { "Steuer-Auswertung Kinderbetreuung $year" } else { "Vollstaendige Dokumentation Kinderbetreuung $year" }

        $ws.Cells.Item(1,1)=$title
        $ws.Cells.Item(1,1).Font.Bold=$true
        $ws.Cells.Item(1,1).Font.Size=16
        $ws.Range("A1:S1").Merge()
        $ws.Cells.Item(2,1)="Kind:"; $ws.Cells.Item(2,2)=$txtChild.Text
        $ws.Cells.Item(3,1)="Betreuungsperson(en):"; $ws.Cells.Item(3,2)=$txtCare.Text
        $ws.Cells.Item(4,1)="Schriftliche Vereinbarung / Betreuungsvertrag:"
        $ws.Cells.Item(4,2)=if($chkContract.Checked){"Ja"}else{"Nein"}
        if ($mode -eq "full") {
            $ws.Cells.Item(5,1)="IBAN Betreuungsperson:"
            $ws.Cells.Item(5,2)=$txtIban.Text
        }
        $ws.Cells.Item(6,1)="Betreuungsart:"
        $ws.Cells.Item(6,2)=if($rbPaid.Checked){"Betreuung gegen Entgelt"}else{"Betreuung unentgeltlich"}
        if($rbPaid.Checked){
            $ws.Cells.Item(6,4)="Entgelt je Betreuungstag:"
            $ws.Cells.Item(6,5)=[double]$numFee.Value
        }

        if ($mode -eq "tax") {
            $headers=@("Datum","Wochentag","KW","Betreuungsperson","Art","Fahrtkosten","Stunden","EUR/Stunde","km gesamt","Fahrtkosten EUR","Betreuungsentgelt EUR","Steuerlich relevante Kosten EUR")
        } else {
            $headers=@("Datum","Wochentag","KW","Betreuungsperson","Art","Fahrtkosten","Stunden","EUR/Stunde","km Betreuung","Fahrtkosten EUR","Betreuungsentgelt EUR","Zusatzkosten/Doku EUR","Kategorie","Notiz","Ausflug/Ziel","Ausflug-km","Eintritt/Freizeit EUR","Hinweis")
        }

        for($i=0;$i -lt $headers.Count;$i++){
            $ws.Cells.Item(8,$i+1)=$headers[$i]
            $ws.Cells.Item(8,$i+1).Font.Bold=$true
        }

        $r=9
        $kmOne=[double]$numKm.Value
        $rate=[double]$numRate.Value
        $names=@("Montag","Dienstag","Mittwoch","Donnerstag","Freitag")
        $taxTotal = 0.0

        $quarterData = @{
            1 = [pscustomobject]@{ Days=0; Travel=0.0; Fee=0.0; Total=0.0 }
            2 = [pscustomobject]@{ Days=0; Travel=0.0; Fee=0.0; Total=0.0 }
            3 = [pscustomobject]@{ Days=0; Travel=0.0; Fee=0.0; Total=0.0 }
            4 = [pscustomobject]@{ Days=0; Travel=0.0; Fee=0.0; Total=0.0 }
        }

        foreach($w in $script:Weeks){
            for($d=0;$d -lt 5;$d++){
                $day=$w.Days[$d]
                if(-not [bool]$day.Checked){continue}

                $dt=[datetime]::ParseExact($day.Date,"yyyy-MM-dd",$null)
                $extra=Parse-DecimalDE ([string]$day.Kosten)
                $tripKm=Parse-DecimalDE ([string]$day.AusflugKm)
                $entry=Parse-DecimalDE ([string]$day.Eintritt)
                $kmTotal = if([bool]$day.Fahrtkosten){ $kmOne*2 }else{ 0.0 }
                $travelCost = if([bool]$day.Fahrtkosten){ $kmTotal*$rate }else{ 0.0 }
                $person = if([string]::IsNullOrWhiteSpace([string]$day.Person)){$txtCare.Text}else{[string]$day.Person}
                $type = [string]$day.Art
                if($type -notin @("unentgeltlich","entgeltlich")){$type="unentgeltlich"}
                $hours = if($type -eq "entgeltlich"){ Parse-DecimalDE ([string]$day.Stunden) }else{0.0}
                $hourly = if($type -eq "entgeltlich"){ Parse-DecimalDE ([string]$day.Stundenlohn) }else{0.0}
                $fee = [math]::Round($hours * $hourly,2)

                $taxRelevant = $travelCost + $fee
                $taxTotal += $taxRelevant
                $quarter = [int][math]::Floor(($dt.Month - 1) / 3) + 1
                $quarterData[$quarter].Days++
                $quarterData[$quarter].Travel += $travelCost
                $quarterData[$quarter].Fee += $fee
                $quarterData[$quarter].Total += $taxRelevant

                $ws.Cells.Item($r,1)=$dt.ToString("dd.MM.yyyy")
                $ws.Cells.Item($r,2)=$names[$d]
                $ws.Cells.Item($r,3)="KW $($w.Week)"
                $ws.Cells.Item($r,4)=$person
                $ws.Cells.Item($r,5)=$type
                $ws.Cells.Item($r,6)=if([bool]$day.Fahrtkosten){"Ja"}else{"Nein"}
                $ws.Cells.Item($r,7)=$hours
                $ws.Cells.Item($r,8)=$hourly
                $ws.Cells.Item($r,9)=$kmTotal
                $ws.Cells.Item($r,10)=$travelCost
                $ws.Cells.Item($r,11)=$fee
                $ws.Cells.Item($r,12)=$taxRelevant

                if ($mode -eq "full") {
                    $ws.Cells.Item($r,13)=$extra
                    $ws.Cells.Item($r,14)=[string]$day.Kategorie
                    $ws.Cells.Item($r,15)=[string]$day.Notiz
                    $ws.Cells.Item($r,16)=[string]$day.AusflugZiel
                    $ws.Cells.Item($r,17)=$tripKm
                    $ws.Cells.Item($r,18)=$entry
                    if ($tripKm -gt 0 -or $entry -gt 0 -or -not [string]::IsNullOrWhiteSpace([string]$day.AusflugZiel)) {
                        $ws.Cells.Item($r,19)="Freizeit-/Ausflugskosten: nur Dokumentation, nicht in Steuer-Summe berücksichtigt"
                    }
                }
                $r++
            }
        }

        $sum=$r+1
        $ws.Cells.Item($sum,1)="Summen";$ws.Cells.Item($sum,1).Font.Bold=$true
        if($r -gt 9){
            $ws.Cells.Item($sum,9).Formula="=SUM(I9:I$($r-1))"
            $ws.Cells.Item($sum,10).Formula="=SUM(J9:J$($r-1))"
            $ws.Cells.Item($sum,11).Formula="=SUM(K9:K$($r-1))"
            $ws.Cells.Item($sum,12).Formula="=SUM(L9:L$($r-1))"
            if($mode -eq "full"){
                $ws.Cells.Item($sum,13).Formula="=SUM(M9:M$($r-1))"
                $ws.Cells.Item($sum,17).Formula="=SUM(Q9:Q$($r-1))"
                $ws.Cells.Item($sum,18).Formula="=SUM(R9:R$($r-1))"
            }
        }

        $noteRow = $sum + 2
        $ws.Cells.Item($noteRow,1)="Steuerlich berücksichtigte Kosten gesamt:"
        $ws.Cells.Item($noteRow,2)=$taxTotal
        $ws.Cells.Item($noteRow,1).Font.Bold=$true
        $ws.Cells.Item($noteRow,2).Font.Bold=$true

        $deductible = [math]::Min($taxTotal * 0.80, 4800.0)
        $ws.Cells.Item($noteRow+1,1)="Davon 80 % (max. 4.800 EUR):"
        $ws.Cells.Item($noteRow+1,2)=$deductible
        $ws.Cells.Item($noteRow+1,1).Font.Bold=$true
        $ws.Cells.Item($noteRow+1,2).Font.Bold=$true


        $qStart = $noteRow + 4
        $ws.Cells.Item($qStart,1)="Quartalsuebersicht"
        $ws.Cells.Item($qStart,1).Font.Bold=$true
        $ws.Cells.Item($qStart,1).Font.Size=12

        $ws.Cells.Item($qStart+1,1)="Quartal"
        $ws.Cells.Item($qStart+1,2)="Betreuungstage"
        $ws.Cells.Item($qStart+1,3)="Fahrtkosten EUR"
        $ws.Cells.Item($qStart+1,4)="Betreuungsentgelt EUR"
        $ws.Cells.Item($qStart+1,5)="Zu zahlen / erstatten EUR"
        $ws.Range("A$($qStart+1):E$($qStart+1)").Font.Bold=$true

        for($q=1;$q -le 4;$q++){
            $qr=$qStart+1+$q
            $qd=$quarterData[$q]
            $ws.Cells.Item($qr,1)="Q$q"
            $ws.Cells.Item($qr,2)=$qd.Days
            $ws.Cells.Item($qr,3)=[math]::Round($qd.Travel,2)
            $ws.Cells.Item($qr,4)=[math]::Round($qd.Fee,2)
            $ws.Cells.Item($qr,5)=[math]::Round($qd.Total,2)
        }

        $qTotalRow=$qStart+6
        $ws.Cells.Item($qTotalRow,1)="Jahr gesamt"
        $ws.Cells.Item($qTotalRow,1).Font.Bold=$true
        $ws.Cells.Item($qTotalRow,2)="=SUM(B$($qStart+2):B$($qStart+5))"
        $ws.Cells.Item($qTotalRow,3)="=SUM(C$($qStart+2):C$($qStart+5))"
        $ws.Cells.Item($qTotalRow,4)="=SUM(D$($qStart+2):D$($qStart+5))"
        $ws.Cells.Item($qTotalRow,5)="=SUM(E$($qStart+2):E$($qStart+5))"
        $ws.Range("A$qTotalRow:E$qTotalRow").Font.Bold=$true

        $ws.Range("C$($qStart+2):E$qTotalRow").NumberFormat='#,##0.00 "EUR"'

        if ($mode -eq "full") {
            $ws.Cells.Item($qStart+8,1)="Hinweis:"
            $ws.Cells.Item($qStart+8,2)="Ausflug-km und Eintritt/Freizeitkosten werden dokumentiert, aber nicht in der steuerlichen Summe berücksichtigt."
            $ws.Range("B$($qStart+8):S$($qStart+8)").Merge()
        }

        # Einheitliches deutsches Zahlenformat fuer Excel und den daraus erzeugten PDF-Export.
        if($r -gt 9){
            $ws.Range("H9:H$($r-1)").NumberFormatLocal='0,00'
            $ws.Range("J9:M$($r-1)").NumberFormatLocal='#.##0,00 "EUR"'
            if($mode -eq "full"){
                $ws.Range("R9:R$($r-1)").NumberFormatLocal='#.##0,00 "EUR"'
            }
        }

        $ws.Cells.Item($noteRow,2).NumberFormatLocal='#.##0,00 "EUR"'
        $ws.Cells.Item($noteRow+1,2).NumberFormatLocal='#.##0,00 "EUR"'
        $ws.Range("C$($qStart+2):E$qTotalRow").NumberFormatLocal='#.##0,00 "EUR"'

        $ws.Columns.AutoFit()|Out-Null
        $ws.PageSetup.Orientation=2
        $ws.PageSetup.Zoom=$false
        $ws.PageSetup.FitToPagesWide=1
        $ws.PageSetup.FitToPagesTall=$false

        $safe=($txtChild.Text -replace '[\\/:*?"<>|]','_')
        if([string]::IsNullOrWhiteSpace($safe)){$safe="Kind"}
        $suffix = if ($mode -eq "tax") { "Steuer" } else { "Vollstaendig" }

        $xlsx=Join-Path $out "Kinderbetreuung_${safe}_${year}_${suffix}.xlsx"
        $pdf=Join-Path $out "Kinderbetreuung_${safe}_${year}_${suffix}.pdf"
        $wb.SaveAs($xlsx,51)
        $ws.ExportAsFixedFormat(0,$pdf)
        $wb.Close($false);$excel.Quit()
        [Runtime.InteropServices.Marshal]::ReleaseComObject($ws)|Out-Null
        [Runtime.InteropServices.Marshal]::ReleaseComObject($wb)|Out-Null
        [Runtime.InteropServices.Marshal]::ReleaseComObject($excel)|Out-Null

        Show-Info "Excel und PDF wurden mit Microsoft Excel erstellt."
        Start-Process explorer.exe $out
    } catch {
        try { if($wb){$wb.Close($false)}; if($excel){$excel.Quit()} } catch {}
        Show-Error $_.Exception.Message
    }
}

