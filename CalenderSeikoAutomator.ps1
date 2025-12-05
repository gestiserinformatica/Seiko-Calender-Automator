###############################################################################
# FILE DIALOG TO SELECT EXCEL
###############################################################################

Add-Type -AssemblyName System.Windows.Forms

$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
$dialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*"
$dialog.Title = "Select the Excel file"

if ($dialog.ShowDialog() -ne "OK") {
    Write-Host "No file selected. Exiting..." -ForegroundColor Red
    exit
}

$ExcelPath = $dialog.FileName
Write-Host "Excel selected: $ExcelPath" -ForegroundColor Green


###############################################################################
# FUNCTIONS
###############################################################################

function Get-FestivosBarcelona {
    $url = "https://datos.madrid.es/egob/catalogo/300082-8-calendario_laboral.ics"

    try {
        Write-Host "Downloading Barcelona holidays..." -ForegroundColor Cyan
        
        $ics = Invoke-WebRequest -Uri $url -UseBasicParsing
        $lines = $ics.Content -split "`n"

        $festivos = foreach ($line in $lines) {
            if ($line -match "^DTSTART:(\d{8})") {
                [datetime]::ParseExact($Matches[1], "yyyyMMdd", $null)
            }
        }

        Write-Host "Holidays found: $($festivos.Count)" -ForegroundColor Green
        return $festivos
    }
    catch {
        Write-Host "ERROR downloading holidays: $_" -ForegroundColor Red
        return @()
    }
}

function Get-WorkingDayBefore {
    param(
        [datetime]$Fecha,
        [datetime[]]$Festivos
    )

    $newDate = $Fecha.AddDays(-1)

    while (
        $newDate.DayOfWeek -in ('Saturday','Sunday') -or
        $Festivos.Date -contains $newDate.Date
    ) {
        $newDate = $newDate.AddDays(-1)
    }

    return $newDate
}

###############################################################################
# FIND OR CREATE CALENDARS IN LOCAL OUTLOOK PROFILE
###############################################################################

function Find-OrCreateCalendar {
    param([string]$CalendarName)

    $Outlook = New-Object -ComObject Outlook.Application
    $Namespace = $Outlook.GetNamespace("MAPI")

    $MainCalendar = $Namespace.GetDefaultFolder(
        [Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar
    )

    foreach ($folder in $MainCalendar.Folders) {
        if ($folder.Name -eq $CalendarName) {
            Write-Host "✔ Found calendar '$CalendarName'" -ForegroundColor Green
            return $folder
        }
    }

    Write-Host "⚠ Creating calendar '$CalendarName'..." -ForegroundColor Yellow
    return $MainCalendar.Folders.Add(
        $CalendarName,
        [Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar
    )
}

$CalendarEmbargos = Find-OrCreateCalendar -CalendarName "Embargos Seiko"
$CalendarVenta    = Find-OrCreateCalendar -CalendarName "Venta al público"

###############################################################################
# READ EXCEL
###############################################################################

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel module not found. Installing..." -ForegroundColor Yellow
    Install-Module ImportExcel -Scope CurrentUser -Force
}

Write-Host "Reading Excel with Import-Excel..." -ForegroundColor Cyan
$Data = Import-Excel -Path $ExcelPath
Write-Host "Excel loaded: $($Data.Count) rows" -ForegroundColor Green

###############################################################################
# PROCESS ROWS & CREATE EVENTS
###############################################################################

Write-Host "`nCreating Outlook events..." -ForegroundColor Cyan

$Festivos = Get-FestivosBarcelona

foreach ($row in $Data) {

    ###############################################################################
    # EMBARGO EVENT
    ###############################################################################

    if ($row."Activar Embargo?" -eq "SI") {

        $rawDate = $row."embargo hasta"

        if ($rawDate -is [datetime]) { $FechaEmbargo = $rawDate }
        elseif ($rawDate -is [double]) { $FechaEmbargo = ([datetime]"1899-12-30").AddDays($rawDate) }
        else {
            try { $FechaEmbargo = [datetime]$rawDate }
            catch { Write-Host "❌ Invalid embargo date: $rawDate" -ForegroundColor Red; continue }
        }

        $FechaRecordatorio = Get-WorkingDayBefore -Fecha $FechaEmbargo -Festivos $Festivos

        $Titulo = "FIN EMBARGO - $($row.'Texto breve de material') - $($row.Material) - $($FechaEmbargo.ToString('dd/MM/yyyy'))"

        $Appt = $CalendarEmbargos.Items.Add("IPM.Appointment")
        $Appt.Subject = $Titulo
        $Appt.Start = $FechaRecordatorio.Date
        $Appt.AllDayEvent = $true
        $Appt.Body = "Recordatorio de fin de embargo para material $($row.Material)"
        $Appt.Save()

        Write-Host "✔ EMBARGO: $Titulo → $FechaRecordatorio" -ForegroundColor Green
    }

    ###############################################################################
    # VENTA AL PUBLICO EVENT
    ###############################################################################

    if ($row."Activar Venta al Publico?" -eq "SI") {

        $rawFechaVenta = $row."Venta al público desde"

        if ($rawFechaVenta -is [datetime]) { $FechaVenta = $rawFechaVenta }
        elseif ($rawFechaVenta -is [double]) { $FechaVenta = ([datetime]"1899-12-30").AddDays($rawFechaVenta) }
        else {
            try { $FechaVenta = [datetime]$rawFechaVenta }
            catch { Write-Host "❌ Invalid venta date: $rawFechaVenta" -ForegroundColor Red; continue }
        }

        $TituloVenta = "FIN VENTA AL PÚBLICO - $($row.'Texto breve de material') - $($row.Material) - $($FechaVenta.ToString('dd/MM/yyyy'))"

        $ApptVenta = $CalendarVenta.Items.Add("IPM.Appointment")
        $ApptVenta.Subject = $TituloVenta
        $ApptVenta.Start = $FechaVenta.Date
        $ApptVenta.AllDayEvent = $true
        $ApptVenta.Body = "Inicio de venta al público para material $($row.Material)"
        $ApptVenta.Save()

        Write-Host "✔ VENTA: $TituloVenta → $FechaVenta" -ForegroundColor Green
    }

}

###############################################################################
# FINAL SUMMARY — COUNT EVENTS
###############################################################################

$TotalEmbargos = $CalendarEmbargos.Items.Count
$TotalVentas   = $CalendarVenta.Items.Count

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "   FINAL SUMMARY" -ForegroundColor Cyan
Write-Host "============================================"
Write-Host "Total events in 'Embargos Seiko'     : $TotalEmbargos" -ForegroundColor Green
Write-Host "Total events in 'Venta al público'   : $TotalVentas"   -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "DONE ✔" -ForegroundColor Green
