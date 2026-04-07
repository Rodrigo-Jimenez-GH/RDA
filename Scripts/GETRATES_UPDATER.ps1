# =========================================
# GET RATES – FEDEX Updater V1.0
# Author: Rodrigo Jimenez Alcocer
# =========================================

# ------------------------------
# Logging
# ------------------------------
$logPath = "$env:TEMP\getrates_update.log"
Start-Transcript -Path $logPath -Append -ErrorAction SilentlyContinue


# ------------------------------
# Banner ASCII
# ------------------------------
Write-Host "=================================================="
Write-Host " ########\              ##\ " -ForegroundColor magenta  -nonewline; Write-Host "########\            " -ForegroundColor red
Write-Host " ##  _____|             ## |" -ForegroundColor magenta  -nonewline; Write-Host "##  _____|           " -ForegroundColor red
Write-Host " ## |    ######\   ####### |" -ForegroundColor magenta  -nonewline; Write-Host "## |      ##\   ##\  " -ForegroundColor red
Write-Host " #####\ ##  __##\  ####### |" -ForegroundColor magenta  -nonewline; Write-Host "#####\    \##\ ##  | " -ForegroundColor red
Write-Host " ##  __|######## |## /  ## |" -ForegroundColor magenta  -nonewline; Write-Host "##  __|    \####  /  " -ForegroundColor red
Write-Host " ## |   ##   ____|## |  ## |" -ForegroundColor magenta  -nonewline; Write-Host "## |       ##  ##<   " -ForegroundColor red
Write-Host " ## |   \#######\ \####### |" -ForegroundColor magenta  -nonewline; Write-Host "########\ ##  /\##\  " -ForegroundColor red
Write-Host " \__|    \_______| \_______|" -ForegroundColor magenta  -nonewline; Write-Host "\________|\__/  \__| " -ForegroundColor red
Write-Host "=================================================="
Write-Host "     GET RATES – FEDEX CORPORATE TOOL UPDATER     " -ForegroundColor yellow
Write-Host "=================================================="
Write-Host "  Developed by (9791981) Rodrigo Jimenez Alcocer  " -ForegroundColor yellow

winget upgrade --id Microsoft.PowerShell --scope user --accept-package-agreements --accept-source-agreements


# ------------------------------
# Cerrar Excel abierto
# ------------------------------
$excelProcs = Get-Process excel -ErrorAction SilentlyContinue
if ($excelProcs) {
    Write-Host " Cerrando instancias de Excel...             " -ForegroundColor magenta
    $excelProcs | Stop-Process -Force
    Write-Host " ↳ Procesos de Excel cerrados Exitosamente   " -ForegroundColor Green
} else {
    Write-Host " ↳ No hay instancias de Excel abiertas       " -ForegroundColor Green
}

Start-Sleep -Seconds 3

# ------------------------------
# Carpetas del script y proyecto raíz
# ------------------------------
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$target = Join-Path $scriptDir ".." | Resolve-Path -Relative
Write-Host "Target folder (proyecto raíz): $target" -ForegroundColor yellow

# ------------------------------
# Limpiar todo excepto la carpeta del script
# ------------------------------
Write-Host " Limpiando carpeta del proyecto... " -ForegroundColor magenta
Remove-Item "$target\*" -Recurse -Force
Write-Host " ↳ Limpieza de carpeta completa    " -ForegroundColor Green

# ------------------------------
# Descargar ZIP del source code de la última release
# ------------------------------
$repoApi = "https://api.github.com/repos/Rodrigo-Jimenez-GH/RDA/releases/latest"
$release = Invoke-RestMethod -Uri $repoApi
$latestTag = $release.tag_name
Write-Host "Última Version Disponible: " -nonewline; write-host "GetRates" $latestTag -ForegroundColor yellow

$zipUrl = "https://github.com/Rodrigo-Jimenez-GH/RDA/archive/refs/tags/$latestTag.zip"
$tempZip = "$env:TEMP\getrates_release.zip"
$tempExtract = "$env:TEMP\GETRATES_EXTRACT"

# Limpiar temporales anteriores
if (Test-Path $tempExtract) { Remove-Item $tempExtract -Recurse -Force }
if (Test-Path $tempZip) { Remove-Item $tempZip -Force }

Write-Host " Descargando ZIP del source code..." -ForegroundColor magenta
Invoke-WebRequest -Uri $zipUrl -OutFile $tempZip -UseBasicParsing
Write-Host " ↳ Descarga completada             " -ForegroundColor Green

# ------------------------------
# Extraer ZIP en carpeta temporal
# ------------------------------
Expand-Archive -Path $tempZip -DestinationPath $tempExtract -Force
Write-Host " Archivos listos en carpeta temporal" -ForegroundColor magenta

# ------------------------------
# Mover contenido de RDA-<tag> al main folder
# ------------------------------
$extractedFolder = Get-ChildItem $tempExtract | Where-Object { $_.PSIsContainer } | Select-Object -First 1
Write-Host  " Clonando repositorio en Carpeta de la MACRO" -ForegroundColor magenta

Get-ChildItem $extractedFolder.FullName -Recurse | ForEach-Object {
    $relativePath = $_.FullName.Substring($extractedFolder.FullName.Length).TrimStart('\')
    $dest = Join-Path $target $relativePath
    $destDir = Split-Path $dest -Parent
    if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
    Copy-Item $_.FullName $dest -Force
}
Write-Host  " ↳ Repositorio GET RATES $latestTag Clonado con exito" -ForegroundColor green

# ------------------------------
# Limpiar temporales
# ------------------------------
Remove-Item $tempZip -Force
Remove-Item $tempExtract -Recurse -Force
Write-Host " Limpieza de archivos temporales listo" -ForegroundColor green

# ------------------------------
# Habilitar macros en Excel
# ------------------------------
Write-Host "========================================================================" -ForegroundColor CYAN
Write-Host "           INICIANDO DESBLOQUEO DE MACROS EN CARPETA OBJETIVO           " -ForegroundColor White -BackgroundColor DarkBlue
Write-Host "========================================================================" -ForegroundColor CYAN
$macroKey = "Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Security"
if (-not (Test-Path $macroKey)) { New-Item -Path $macroKey -Force | Out-Null }
Set-ItemProperty -Path $macroKey -Name "VBAWarnings" -Value 1
Write-Host "PERMISO ACEPTADO: Uso de macros en carpeta objetivo" -ForegroundColor CYAN

$dataKey = "Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Security\Trusted Locations"
if (-not (Test-Path $dataKey)) { New-Item -Path $dataKey -Force | Out-Null } # Habilitar datos externos
Write-Host "PERMISO ACEPTADO: Datos externos" -ForegroundColor CYAN

$extKey = "Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Security\Trusted Documents"
if (-not (Test-Path $extKey)) { New-Item -Path $extKey -Force | Out-Null } # Opcional: permitir todos los vínculos externos sin advertencia
Write-Host "PERMISO ACEPTADO: vinculos externos" -ForegroundColor CYAN

# Evita advertencias de actualización automática de vínculos
$linkKey = "Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Options"
Set-ItemProperty -Path $linkKey -Name "UpdateLinks" -Value 1  # 1 = actualizar todos los vínculos sin preguntar
Write-Host "PERMISO ACEPTADO: Actualziar Vinculos" -ForegroundColor CYAN

$activexKey = "Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Security\AccessVBOM"
Set-ItemProperty -Path $activexKey -Name "AccessVBOM" -Value 1  # Permitir acceso al proyecto VBA
Write-Host "PERMISO ACEPTADO: Controles ACTIVE X" -ForegroundColor CYAN
# ------------------------------
# Crear shortcut en Desktop
# ------------------------------
$excelFile = (Join-Path $target "MACROS RDA.xlsm" | Resolve-Path -ErrorAction Stop).Path
$desktop = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktop "MACROS RDA.lnk"
$WshShell = New-Object -ComObject WScript.Shell

if (Test-Path $shortcutPath) { Remove-Item $shortcutPath -Force }

$shortcut = $WshShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $excelFile
$shortcut.WorkingDirectory = $target
$shortcut.WindowStyle = 1
$shortcut.Description = "Shortcut to MACROS RDA workbook"
$shortcut.Save()
Write-Host "AGREAGADO: Acceso directo a la macro en $shortcutPath" -ForegroundColor Green

# ------------------------------
# Abrir Excel
# ------------------------------
$excelApp = New-Object -ComObject Excel.Application
$excelApp.Visible = $true
$workbook = $excelApp.Workbooks.Open($excelFile)
Write-Host "Abriendo MACRO $excelFile" -ForegroundColor yellow

# ------------------------------
# Finalización
# ------------------------------
Write-Host "==========================================" -ForegroundColor green
Write-Host "  ✭ Actualización completada con éxito ✭ " -ForegroundColor green
Write-Host "==========================================" -ForegroundColor green
Write-Host "log de instalacion exportado correctamente" -ForegroundColor yellow
Stop-Transcript | Out-Null
exit
