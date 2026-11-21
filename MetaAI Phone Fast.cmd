@echo off
setlocal
title Meta AI Phone Fast (Mail.cx)

set "SAVE_DIR=%~dp0..\Metaai Credential File"
set "WORKER=%~dp0metaai_factory_fast.ps1"
set "ADB=C:\platform-tools\adb.exe"
set "SERIAL=%ADB_SERIAL%"
set "CRED_NAME="
set "COUNT="
set "ACCOUNT_PASS="

:: Locate ADB platform-tools automatically if default path does not exist
if not exist "%ADB%" (
    for /f "delims=" %%I in ('where adb 2^>nul') do if not defined ADB set "ADB=%%I"
)
if not exist "%ADB%" (
    if exist "%LOCALAPPDATA%\Microsoft\WinGet\Packages\Google.PlatformTools_Microsoft.Winget.Source_8wekyb3d8bbwe\platform-tools\adb.exe" (
        set "ADB=%LOCALAPPDATA%\Microsoft\WinGet\Packages\Google.PlatformTools_Microsoft.Winget.Source_8wekyb3d8bbwe\platform-tools\adb.exe"
    ) else if exist "%LOCALAPPDATA%\Android\Sdk\platform-tools\adb.exe" (
        set "ADB=%LOCALAPPDATA%\Android\Sdk\platform-tools\adb.exe"
    )
)

if not exist "%ADB%" (
    echo.
    echo [ERROR] ADB platform-tools executable not found!
    echo Please make sure platform-tools is installed at C:\platform-tools\adb.exe or available in PATH.
    echo.
    pause
    exit /b 1
)

echo.
echo ====================================
echo    Meta AI Phone Fast (Mail.cx)
echo ====================================
echo.

:ask_count
set /p "COUNT=How many accounts to register: "
if not defined COUNT goto :invalid_count
for /f "delims=0123456789" %%A in ("%COUNT%") do goto :invalid_count
if "%COUNT%"=="0" goto :invalid_count
goto :ask_pass

:invalid_count
echo Please enter a positive whole number.
echo.
goto :ask_count

:ask_pass
set /p "ACCOUNT_PASS=Password for accounts (or leave blank to use METAAI_DEFAULT_PASSWORD): "
if not defined ACCOUNT_PASS set "ACCOUNT_PASS=%METAAI_DEFAULT_PASSWORD%"

:ask_cred
set /p "CRED_NAME=Credential file name (without .xlsx): "
if not defined CRED_NAME set "CRED_NAME=metaai_mailcx_%DATE:/=-%"
set "CRED_FILE=%SAVE_DIR%\%CRED_NAME%.xlsx"
set "PROGRESS=%TEMP%\metaai_mailcx_progress.txt"

if not exist "%SAVE_DIR%" mkdir "%SAVE_DIR%"

echo.
echo ------------------------------------
echo Accounts to create : %COUNT%
echo Password : [hidden]
echo Credential output  : %CRED_FILE%
echo Mail Provider      : Mail.cx (api.mail.cx)
echo ADB binary         : %ADB%
echo ------------------------------------
echo.
echo Starting phone worker...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%WORKER%" -Serial "%SERIAL%" -AdbPath "%ADB%" -Count %COUNT% -Password "%ACCOUNT_PASS%" -CredFile "%CRED_FILE%" -ProgressFile "%PROGRESS%"
set "EXIT_CODE=%ERRORLEVEL%"
echo.
echo Worker finished with exit code %EXIT_CODE%.
echo Credential file: %CRED_FILE%
pause
exit /b %EXIT_CODE%
