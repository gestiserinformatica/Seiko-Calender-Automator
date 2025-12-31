###############################################################################
# REMOVE ALL EVENTS FROM SPECIFIC OUTLOOK CALENDARS (seikoembargos@seiko.es)
# - Uses ONLY the mailbox seikoembargos@seiko.es already added in Outlook profile
# - Finds calendars by name under that mailbox
# - DOES NOT create anything
###############################################################################

Write-Host "`n=== BORRANDO EVENTOS DE CALENDARIOS (seikoembargos@seiko.es) ===`n" -ForegroundColor Cyan

# Calendarios a limpiar (exactos)
$CalendarsToClean = @(
    "Embargos Seiko",
    "Venta al público"
)

###############################################################################
# Helper: normalize (case-insensitive + remove accents + trim spaces)
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

###############################################################################
# Outlook session
###############################################################################
$Outlook   = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")

###############################################################################
# Find mailbox root in Outlook profile (no shared default folder)
###############################################################################
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

###############################################################################
# Find calendar folder by name anywhere under mailbox root
###############################################################################
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

###############################################################################
# Target mailbox + resolve calendars
###############################################################################
$TargetMailbox = "seikoembargos@seiko.es"
$MailboxRoot = Get-MailboxRootByName -Namespace $Namespace -MailboxName $TargetMailbox

foreach ($CalName in $CalendarsToClean) {

    Write-Host "`nProcesando calendario: $CalName" -ForegroundColor Yellow

    # Try exact (and also "Venta al Publico" without accent as fallback)
    $Calendar = Find-CalendarFolderByNameUnderRoot -RootFolder $MailboxRoot -CalendarName $CalName
    if (-not $Calendar -and (Normalize-Name $CalName) -eq (Normalize-Name "Venta al público")) {
        $Calendar = Find-CalendarFolderByNameUnderRoot -RootFolder $MailboxRoot -CalendarName "Venta al Publico"
    }

    if (-not $Calendar) {
        Write-Host "❌ No encontrado bajo '$TargetMailbox': $CalName — Saltando" -ForegroundColor Red
        continue
    }

    Write-Host "✔ Calendario encontrado: $($Calendar.Name) | $($Calendar.FolderPath)" -ForegroundColor Green

    # Items config to include recurring events properly
    $Items = $Calendar.Items
    $Items.IncludeRecurrences = $true
    $Items.Sort("[Start]")

    $Total = $Items.Count
    Write-Host "Eventos encontrados (incluye recurrencias): $Total" -ForegroundColor Cyan

    # Copy to array to avoid modifying collection while iterating
    $ToDelete = New-Object System.Collections.Generic.List[object]
    foreach ($item in $Items) {
        $null = $ToDelete.Add($item)
    }

    # Delete all items
    $Deleted = 0
    foreach ($item in $ToDelete) {
        try {
            $item.Delete()
            $Deleted++
        }
        catch {
            Write-Host "⚠ Error al borrar un evento: $_" -ForegroundColor Red
        }
    }

    Write-Host "✔ Borrados correctamente: $Deleted eventos en '$($Calendar.Name)'" -ForegroundColor Green
}

Write-Host "`n=== PROCESO FINALIZADO ===" -ForegroundColor Cyan
