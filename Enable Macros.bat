@echo off
:: ===================================================
:: Add current folder as trusted location for Excel
:: ===================================================

winget upgrade --id Microsoft.PowerShell --scope user --accept-package-agreements --accept-source-agreements
set "PATH=%PATH%;%LOCALAPPDATA%\Microsoft\PowerShell\7"
:: Get the folder where the BAT file is located
set "CURRENT_FOLDER=%~dp0"

:: Remove trailing backslash if exists
if "%CURRENT_FOLDER:~-1%"=="\" set "CURRENT_FOLDER=%CURRENT_FOLDER:~0,-1%"

:: Registry path for Excel trusted locations
set "REG_PATH=HKCU\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations"

:: Find the next available LocationX number
for /f "tokens=*" %%i in ('reg query "%REG_PATH%" 2^>nul ^| findstr /r /c:"Location[0-9]*"') do (
    set "LAST_LOC=%%i"
)

if defined LAST_LOC (
    for /f "tokens=1 delims=\" %%n in ("%LAST_LOC%") do set /a NEXT=%%n+1
) else (
    set NEXT=1
)

:: Add the trusted location
reg add "%REG_PATH%\Location%NEXT%" /v Path /t REG_SZ /d "%CURRENT_FOLDER%" /f
reg add "%REG_PATH%\Location%NEXT%" /v AllowSubFolders /t REG_DWORD /d 1 /f

exit
