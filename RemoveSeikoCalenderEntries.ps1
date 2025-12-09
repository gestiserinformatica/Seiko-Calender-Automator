###############################################################################
# REMOVE ALL EVENTS FROM SPECIFIC OUTLOOK CALENDARS
###############################################################################

Write-Host "`n=== BORRANDO EVENTOS DE CALENDARIOS ===`n" -ForegroundColor Cyan

# Nombres de los calendarios usados
$CalendarsToClean = @(
    "Embargos Seiko",
    "Venta al público"
)

# Cargar Outlook
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")

# Carpeta Calendario principal
$MainCalendar = $Namespace.GetDefaultFolder(
    [Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderCalendar
)

foreach ($CalName in $CalendarsToClean) {

    Write-Host "`nProcesando calendario: $CalName" -ForegroundColor Yellow

    # Buscar calendario dentro del calendario principal
    $Calendar = $MainCalendar.Folders | Where-Object { $_.Name -eq $CalName }

    if (-not $Calendar) {
        Write-Host "❌ No encontrado: $CalName — Saltando" -ForegroundColor Red
        continue
    }

    $Items = $Calendar.Items
    $Total = $Items.Count

    Write-Host "Eventos encontrados: $Total" -ForegroundColor Cyan

    # Lista temporal para evitar modificar colección durante iteración
    $ToDelete = @()
    foreach ($item in $Items) {
        $ToDelete += $item
    }

    # Eliminar todos los eventos
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

    Write-Host "✔ Borrados correctamente: $Deleted eventos en '$CalName'" -ForegroundColor Green
}

Write-Host "`n=== PROCESO FINALIZADO ===" -ForegroundColor Cyan
