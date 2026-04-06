@echo off
echo [33m:: ===================================================[0m
echo [33m:: Instalador de Powershell y Habilitador de Macros[0m
echo [33m:: ===================================================[0m
echo [33m:: Desarrollado por Rodrigo Jimenez[0m
echo [36m:: ======================== Revisando dependencias de powershell ===============================[0m

where pwsh >nul 2>&1
if %errorlevel%==0 (
    for /f "delims=" %%i in ('where pwsh') do set "PSPATH=%%i"
) else (
    :: Buscar en rutas comunes
    if exist "%ProgramFiles%\PowerShell\7\pwsh.exe" set "PSPATH=%ProgramFiles%\PowerShell\7\pwsh.exe"
    if exist "%LOCALAPPDATA%\Microsoft\PowerShell\7\pwsh.exe" set "PSPATH=%LOCALAPPDATA%\Microsoft\PowerShell\7\pwsh.exe"
)

if defined PSPATH (
    echo PowerShell encontrado en: %PSPATH%
    pwsh -NoProfile -Command "Write-Host 'PowerShell Funcionando correctamente' -ForegroundColor Green"
) else (
    echo [31mPowerShell no encontrado. Instalando...[0m
    winget install --id Microsoft.PowerShell --scope user --accept-package-agreements --accept-source-agreements -e
)
echo [36m:: ======================== HABILITAR MACROS ===============================[0m
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
echo Añadiendo %CURRENT_FOLDER% a lugares de confianza...
reg add "%REG_PATH%\Location%NEXT%" /v Path /t REG_SZ /d "%CURRENT_FOLDER%" /f
echo Añadiendo Subcarpetas a lugares de confianza...
reg add "%REG_PATH%\Location%NEXT%" /v AllowSubFolders /t REG_DWORD /d 1 /f

echo [32mMacros activadas en: %CURRENT_FOLDER%[0m
echo ya puede cerrar la ventana u oprima ENTER
pause
