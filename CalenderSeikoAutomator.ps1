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
# SAFE SUBJECT FOR OUTLOOK RESTRICT
###############################################################################

function Sanitize-ForRestrict {
    param([string]$text)

    if (-not $text) { return "" }

    # Remove characters that break Restrict()
    $text = $text.Replace("'", "")
    $text = $text.Replace('"', "")
    $text = $text.Replace("`n"," ")
    $text = $text.Replace("`r"," ")

    return $text
}

###############################################################################
# OUTLOOK: USE MAILBOX seikoembargos@seiko.es (NO CREATION)
# - Find mailbox root as it appears in Outlook profile
# - Find calendars "Embargos Seiko" and "Venta al público" UNDER THAT MAILBOX
# - If not found -> throw and DO NOT CREATE
###############################################################################

function Normalize-Name {
    param([string]$s)

    if ([string]::IsNullOrWhiteSpace($s)) { return "" }

    $normalized = $s.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $normalized.ToCharArray()) {
        $cat = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch)
        if ($cat -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }

    ($sb.ToString().ToLowerInvariant() -replace '\s+', ' ').Trim()
}

function Get-MailboxRootByName {
    param(
        [Parameter(Mandatory=$true)]$Namespace,
        [Parameter(Mandatory=$true)][string]$MailboxName
    )

    $target = Normalize-Name $MailboxName

    foreach ($f in $Namespace.Folders) {
        try {
            if ((Normalize-Name $f.Name) -eq $target) {
                Write-Host "✔ Found mailbox root: $($f.Name)" -ForegroundColor Green
                return $f
            }
        } catch {}
    }

    foreach ($st in $Namespace.Stores) {
        try {
            if ((Normalize-Name $st.DisplayName) -eq $target) {
                $root = $st.GetRootFolder()
                Write-Host "✔ Found mailbox root (Store): $($st.DisplayName)" -ForegroundColor Green
                return $root
            }
        } catch {}
    }

    throw "Mailbox '$MailboxName' not found in Outlook profile. Make sure it is added."
}

function Find-CalendarFolderByNameUnderRoot {
    param(
        [Parameter(Mandatory=$true)]$RootFolder,
        [Parameter(Mandatory=$true)][string]$CalendarName
    )

    $want = Normalize-Name $CalendarName

    $queue = New-Object System.Collections.Queue
    $queue.Enqueue($RootFolder)

    while ($queue.Count -gt 0) {
        $folder = $queue.Dequeue()

        try {
            # Calendar folders: DefaultItemType == 1 (olAppointmentItem)
            if ($folder.DefaultItemType -eq 1) {
                if ((Normalize-Name $folder.Name) -eq $want) {
                    return $folder
                }
            }
        } catch {}

        try {
            foreach ($sub in $folder.Folders) {
                $queue.Enqueue($sub)
            }
        } catch {}
    }

    return $null
}

# One Outlook session (more reliable)
$Outlook   = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")

$TargetMailbox = "seikoembargos@seiko.es"
$MailboxRoot = Get-MailboxRootByName -Namespace $Namespace -MailboxName $TargetMailbox

$CalendarEmbargos = Find-CalendarFolderByNameUnderRoot -RootFolder $MailboxRoot -CalendarName "Embargos Seiko"
if (-not $CalendarEmbargos) {
    throw "Calendar 'Embargos Seiko' NOT found under mailbox '$TargetMailbox'. Nothing was created."
}
Write-Host "✔ Using calendar 'Embargos Seiko' | $($CalendarEmbargos.FolderPath)" -ForegroundColor Green

# Your tree may show "Venta al Publico" (no accent) or "Venta al público"
$CalendarVenta = Find-CalendarFolderByNameUnderRoot -RootFolder $MailboxRoot -CalendarName "Venta al público"
if (-not $CalendarVenta) {
    $CalendarVenta = Find-CalendarFolderByNameUnderRoot -RootFolder $MailboxRoot -CalendarName "Venta al Publico"
}
if (-not $CalendarVenta) {
    throw "Calendar 'Venta al público' NOT found under mailbox '$TargetMailbox'. Nothing was created."
}
Write-Host "✔ Using calendar 'Venta al público' | $($CalendarVenta.FolderPath)" -ForegroundColor Green

# Improve Items behavior for searching/recurrences
$CalendarEmbargos.Items.IncludeRecurrences = $true
$CalendarEmbargos.Items.Sort("[Start]")
$CalendarVenta.Items.IncludeRecurrences = $true
$CalendarVenta.Items.Sort("[Start]")


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

        # Extract date
        $rawDate = $row."embargo hasta"

        # Detect invalid inputs
        if ($null -eq $rawDate -or ($rawDate.ToString() -match "TBD|-|#N|error")) {
            Write-Host "⚠ Skipped (invalid embargo date): $($row.Material) -> $rawDate" -ForegroundColor Yellow
            continue
        }

        # Convert (keep your original conversion logic)
        if     ($rawDate -is [datetime]) { $FechaEmbargo = $rawDate }
        elseif ($rawDate -is [double])   { $FechaEmbargo = ([datetime]"1899-12-30").AddDays($rawDate) }
        else {
            try { $FechaEmbargo = [datetime]$rawDate }
            catch { Write-Host "⚠ Skipped (invalid embargo date): $rawDate" -ForegroundColor Yellow; continue }
        }

        $FechaRecordatorio = Get-WorkingDayBefore -Fecha $FechaEmbargo -Festivos $Festivos

        $Titulo = "FIN EMBARGO - $($row.'Texto breve de material') - $($row.Material) - $($FechaEmbargo.ToString('dd/MM/yyyy'))"
        $safeTitulo = Sanitize-ForRestrict $Titulo

        # Duplicate check (safe)
        $existing = $CalendarEmbargos.Items.Restrict("[Subject] = '$safeTitulo'")
        if ($existing.Count -gt 0) {
            Write-Host "⚠ Already exists (SKIPPED): $Titulo" -ForegroundColor Yellow
            continue
        }

        # Create event in seikoembargos@seiko.es -> Embargos Seiko
        $Appt = $CalendarEmbargos.Items.Add("IPM.Appointment")
        $Appt.Subject = $Titulo
        $Appt.Start = $FechaRecordatorio.Date
        $Appt.AllDayEvent = $true
        $Appt.Body = "Recordatorio de fin de embargo para material $($row.Material)"
        $Appt.Save()

        Write-Host "✔ EMBARGO CREATED (seikoembargos@seiko.es): $Titulo → $FechaRecordatorio" -ForegroundColor Green
    }

    ###############################################################################
    # VENTA AL PUBLICO EVENT
    ###############################################################################
    if ($row."Activar Venta al Publico?" -eq "SI") {

        $rawFechaVenta = $row."Venta al público desde"

        # Detect invalid inputs
        if ($null -eq $rawFechaVenta -or ($rawFechaVenta.ToString() -match "TBD|-|#N|error")) {
            Write-Host "⚠ Skipped (invalid venta date): $($row.Material) -> $rawFechaVenta" -ForegroundColor Yellow
            continue
        }

        # Convert (keep your original conversion logic)
        if     ($rawFechaVenta -is [datetime]) { $FechaVenta = $rawFechaVenta }
        elseif ($rawFechaVenta -is [double])   { $FechaVenta = ([datetime]"1899-12-30").AddDays($rawFechaVenta) }
        else {
            try { $FechaVenta = [datetime]$rawFechaVenta }
            catch { Write-Host "⚠ Skipped (invalid venta date): $rawFechaVenta" -ForegroundColor Yellow; continue }
        }

        $TituloVenta = "FIN VENTA AL PÚBLICO - $($row.'Texto breve de material') - $($row.Material) - $($FechaVenta.ToString('dd/MM/yyyy'))"
        $safeTituloVenta = Sanitize-ForRestrict $TituloVenta

        # Duplicate check
        $existingVenta = $CalendarVenta.Items.Restrict("[Subject] = '$safeTituloVenta'")
        if ($existingVenta.Count -gt 0) {
            Write-Host "⚠ Already exists (SKIPPED): $TituloVenta" -ForegroundColor Yellow
            continue
        }

        # Create event in seikoembargos@seiko.es -> Venta al público
        $ApptVenta = $CalendarVenta.Items.Add("IPM.Appointment")
        $ApptVenta.Subject = $TituloVenta
        $ApptVenta.Start = $FechaVenta.Date
        $ApptVenta.AllDayEvent = $true
        $ApptVenta.Body = "Inicio de venta al público para material $($row.Material)"
        $ApptVenta.Save()

        Write-Host "✔ VENTA CREATED (seikoembargos@seiko.es): $TituloVenta → $FechaVenta" -ForegroundColor Green
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
Write-Host "Mailbox used: $TargetMailbox" -ForegroundColor Green
Write-Host "Total events in 'Embargos Seiko'     : $TotalEmbargos" -ForegroundColor Green
Write-Host "Total events in 'Venta al público'   : $TotalVentas"   -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "DONE ✔" -ForegroundColor Green
