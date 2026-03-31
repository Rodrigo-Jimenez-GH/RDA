# =========================================
# GET RATES – FEDEX Updater
# Author: Rodrigo Jimenez Alcocer
# =========================================

# ------------------------------
# Configuración de logging
# ------------------------------
$logPath = "$env:TEMP\getrates_update.log"
Start-Transcript -Path $logPath -Append -ErrorAction SilentlyContinue

# ------------------------------
# Banner ASCII limpio
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
    Develpoded by: 9791981 Rodrigo Jimenez Alcocer
========================================================

"@
Write-Host $banner

# ------------------------------
# Carpeta del script
# ------------------------------
$target = Split-Path -Parent $MyInvocation.MyCommand.Path
Write-Host "Target folder: $target"

# ------------------------------
# Crear carpeta si no existe
# ------------------------------
if (-not (Test-Path $target)) {
    Write-Host "Carpeta no existe, creando: $target"
    New-Item -ItemType Directory -Path $target -Force | Out-Null
} else {
    Write-Host "Carpeta existe."
}

# ------------------------------
# Habilitar macros en Excel
# ------------------------------
Write-Host "Configurando seguridad de macros..."
New-ItemProperty -Path Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Security `
    -Name "VBAWarnings" -Value 1 -PropertyType DWORD -Force | Out-Null
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
# Descargar repositorio como ZIP
# ------------------------------
$repoUrl = "https://github.com/Rodrigo-Jimenez-GH/RDA/archive/refs/heads/main.zip"
$tempZip = "$env:TEMP\getrates.zip"
$tempExtract = "$env:TEMP\GETRATES_EXTRACT"

Write-Host "Descargando repositorio..."
Invoke-WebRequest -Uri $repoUrl -OutFile $tempZip -UseBasicParsing
Write-Host "Descarga completada: $tempZip"

# ------------------------------
# Extraer ZIP
# ------------------------------
if (Test-Path $tempExtract) { Remove-Item $tempExtract -Recurse -Force }
Write-Host "Extrayendo ZIP..."
Expand-Archive -Path $tempZip -DestinationPath $tempExtract -Force
Write-Host "Extracción completada: $tempExtract"

# ------------------------------
# Copiar todos los archivos al target
# ------------------------------
$extractedRoot = Join-Path $tempExtract "RDA-main"
Write-Host "Reemplazando archivos en la carpeta del proyecto..."
Get-ChildItem $extractedRoot -Recurse | ForEach-Object {
    $dest = Join-Path $target $_.FullName.Substring($extractedRoot.Length + 1)
    $destFolder = Split-Path $dest -Parent
    if (-not (Test-Path $destFolder)) { New-Item -ItemType Directory -Path $destFolder -Force | Out-Null }
    Copy-Item $_.FullName $dest -Force
    Write-Host "Actualizado: $dest"
}

# ------------------------------
# Limpieza temporal
# ------------------------------
Write-Host "Eliminando archivos temporales..."
Remove-Item $tempZip -Force
Remove-Item $tempExtract -Recurse -Force
Write-Host "Limpieza completada."

# ------------------------------
# Finalización
# ------------------------------
Write-Host "Proceso de actualización finalizado. Log completo en: $logPath"
Stop-Transcript | Out-Null

# Path to Excel workbook
$targetFile = Join-Path $target "MACROS RDA.xlsm"

# Desktop path
$desktop = [Environment]::GetFolderPath("Desktop")

# Shortcut path (same name)
$shortcutPath = Join-Path $desktop "MACROS RDA.lnk"

# Create WScript.Shell COM object
$WshShell = New-Object -ComObject WScript.Shell

# Remove existing shortcut if exists
if (Test-Path $shortcutPath) {
    Write-Host "Existing shortcut found. Removing..."
    Remove-Item $shortcutPath -Force
}

# Create new shortcut
$shortcut = $WshShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $targetFile
$shortcut.WorkingDirectory = $target
$shortcut.WindowStyle = 1        # Normal window
$shortcut.Description = "Shortcut to MACROS RDA workbook"
$shortcut.Save()

Write-Host "Shortcut created on desktop: $shortcutPath"

# Path to the Excel file
$excelFile = Join-Path $target "MACROS RDA.xlsm"

# Start Excel
$excelApp = New-Object -ComObject Excel.Application
$excelApp.Visible = $true

# Open the workbook
$workbook = $excelApp.Workbooks.Open($excelFile)

Write-Host "Workbook opened: $excelFile"


exit