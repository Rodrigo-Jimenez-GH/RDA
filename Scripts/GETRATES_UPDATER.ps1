# =========================================
# GET RATES – FEDEX Updater
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
$banner = @"
========================================================

       .%%%%%%..%%%%%%..%%%%%...%%%%%%..%%..%%.
       .%%......%%......%%..%%..%%.......%%%%..
       .%%%%....%%%%....%%..%%..%%%%......%%...
       .%%......%%......%%..%%..%%.......%%%%..
       .%%......%%%%%%..%%%%%...%%%%%%..%%..%%.
       ........................................

        GET RATES – FEDEX CORPORATE TOOL UPDATER
========================================================
    Developed by: 9791981 Rodrigo Jimenez Alcocer
========================================================
"@
Write-Host $banner

# ------------------------------
# Carpetas del script y proyecto raíz
# ------------------------------
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$target = Join-Path $scriptDir ".." | Resolve-Path -Relative
Write-Host "Target folder (proyecto raíz): $target"

# ------------------------------
# Limpiar todo excepto la carpeta del script
# ------------------------------
Write-Host "Limpiando carpeta del proyecto..."
Get-ChildItem $target -Recurse -Force | Where-Object { $_.FullName -ne $scriptDir } | Remove-Item -Recurse -Force
Write-Host "Carpeta limpia."

# ------------------------------
# Descargar ZIP del source code de la última release
# ------------------------------
$repoApi = "https://api.github.com/repos/Rodrigo-Jimenez-GH/RDA/releases/latest"
$release = Invoke-RestMethod -Uri $repoApi
$latestTag = $release.tag_name
Write-Host "Última release: $latestTag"

$zipUrl = "https://github.com/Rodrigo-Jimenez-GH/RDA/archive/refs/tags/$latestTag.zip"
$tempZip = "$env:TEMP\getrates_release.zip"
$tempExtract = "$env:TEMP\GETRATES_EXTRACT"

# Limpiar temporales anteriores
if (Test-Path $tempExtract) { Remove-Item $tempExtract -Recurse -Force }
if (Test-Path $tempZip) { Remove-Item $tempZip -Force }

Write-Host "Descargando ZIP del source code..."
Invoke-WebRequest -Uri $zipUrl -OutFile $tempZip -UseBasicParsing
Write-Host "Descarga completada: $tempZip"

# ------------------------------
# Extraer ZIP en carpeta temporal
# ------------------------------
Expand-Archive -Path $tempZip -DestinationPath $tempExtract -Force
Write-Host "ZIP extraído en: $tempExtract"

# ------------------------------
# Mover contenido de RDA-<tag> al main folder
# ------------------------------
$extractedFolder = Get-ChildItem $tempExtract | Where-Object { $_.PSIsContainer } | Select-Object -First 1
Write-Host "Moviendo contenido de $($extractedFolder.FullName) a $target"

Get-ChildItem $extractedFolder.FullName -Recurse | ForEach-Object {
    $relativePath = $_.FullName.Substring($extractedFolder.FullName.Length).TrimStart('\')
    $dest = Join-Path $target $relativePath
    $destDir = Split-Path $dest -Parent
    if (-not (Test-Path $destDir)) { New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
    Copy-Item $_.FullName $dest -Force
}

# ------------------------------
# Limpiar temporales
# ------------------------------
Remove-Item $tempZip -Force
Remove-Item $tempExtract -Recurse -Force
Write-Host "Archivos temporales eliminados."

# ------------------------------
# Habilitar macros en Excel
# ------------------------------
Write-Host "Configurando seguridad de macros..."
$macroKey = "Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Security"
if (-not (Test-Path $macroKey)) { New-Item -Path $macroKey -Force | Out-Null }
Set-ItemProperty -Path $macroKey -Name "VBAWarnings" -Value 1
Write-Host "Macros habilitadas."

# ------------------------------
# Cerrar Excel abierto
# ------------------------------
$excelProcs = Get-Process excel -ErrorAction SilentlyContinue
if ($excelProcs) {
    Write-Host "Cerrando instancias de Excel..."
    $excelProcs | Stop-Process -Force
    Write-Host "Excel cerrado."
} else {
    Write-Host "No hay instancias de Excel abiertas."
}

# ------------------------------
# Crear shortcut en Desktop
# ------------------------------
$excelFile = Join-Path $target "MACROS RDA.xlsm" | Resolve-Path -ErrorAction Stop
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
Write-Host "Shortcut creado en Desktop: $shortcutPath"

# ------------------------------
# Abrir Excel
# ------------------------------
$excelApp = New-Object -ComObject Excel.Application
$excelApp.Visible = $true
$workbook = $excelApp.Workbooks.Open($excelFile)
Write-Host "Workbook abierto: $excelFile"

# ------------------------------
# Finalización
# ------------------------------
Write-Host "Actualización completada con éxito. Log completo en: $logPath"
Stop-Transcript | Out-Null
exit