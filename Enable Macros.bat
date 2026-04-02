@echo off
echo :: ===================================================
echo :: Instalador de Powershell y Habilitador de Macros
echo :: ===================================================
echo Cheking Powershell Install

:: Buscar PowerShell en usuario
set "PSPATH="
for /r "%LOCALAPPDATA%\Microsoft\PowerShell" %%f in (pwsh.exe) do set "PSPATH=%%f"

if defined PSPATH (
    echo PowerShell ya Instalado
    "%PSPATH%" -NoProfile -Command "Write-Host Hola"
) else (
    echo PowerShell no encontrado. Instalando...
    winget install --id Microsoft.PowerShell --scope user --accept-package-agreements --accept-source-agreements -e
)
echo :: ======================== HABILITAR MACROS ===============================
:: Obtener carpeta donde está el BAT
set "CURRENT_FOLDER=%~dp0"
:: Eliminar barra final
if "%CURRENT_FOLDER:~-1%"=="\" set "CURRENT_FOLDER=%CURRENT_FOLDER:~0,-1%"

echo Carpeta actual del script: %CURRENT_FOLDER%

:: Ruta del registro para trusted locations de Excel
set "REG_PATH=HKCU\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations"

:: Buscar el último LocationX
set NEXT=1
for /f "tokens=*" %%i in ('reg query "%REG_PATH%" 2^>nul ^| findstr /r /c:"Location[0-9]*"') do (
    set "LAST_LOC=%%~nxi"
)

if defined LAST_LOC (
    :: Extraer solo el número de LocationX
    for /f "tokens=2 delims=Location" %%n in ("%LAST_LOC%") do set /a NEXT=%%n+1
)

echo Siguiente Location: %NEXT%

:: Añadir trusted location
reg add "%REG_PATH%\Location%NEXT%" /v Path /t REG_SZ /d "%CURRENT_FOLDER%" /f
reg add "%REG_PATH%\Location%NEXT%" /v AllowSubFolders /t REG_DWORD /d 1 /f

echo Macros activadas en: %CURRENT_FOLDER%
pause
