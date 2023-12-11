@echo off
setlocal enabledelayedexpansion

where /q git
if ErrorLevel 1 (
    echo Git is not installed, installing now...
    powershell -Command "Invoke-WebRequest https://github.com/git-for-windows/git/releases/download/v2.33.0.windows.2/Git-2.33.0.2-64-bit.exe -OutFile git-installer.exe"
    start /wait git-installer.exe /VERYSILENT
    del git-installer.exe
    echo Git has been installed.
    echo Restarting script after installed git...
) else (
    echo Git is already installed!
)


where /q node
if ErrorLevel 1 (
    echo Node.js is not installed, installing now...
    powershell -Command "Invoke-WebRequest https://nodejs.org/dist/v20.9.0/node-v20.9.0-x64.msi -OutFile node.msi"
    msiexec /i node.msi /passive
    del node.msi
    echo Node.js has been installed.
) else (
    echo Node.js is already installed!
)

@REM Git and Node.js have all prepared. check office_addin_sample_scripts.

where /q office_addin_sample_scripts
if ErrorLevel 1 (
    echo Sample scripts are not prepared, installing now...
    npm install -g office_addin_sample_scripts
    echo Sample scripts has been installed.
    echo Restarting this script after installed sample scripts...
    start "" "%~0"
    exit /b
) else (
    echo Sample scripts is already installed!
)

@REM Check if Word has been installed on the local machine.
:: Define the registry keys for Word to check
set "regKeys=HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\winword.exe"

:: Check each registry key
reg query "%regKeys%" >nul 2>&1
if errorlevel 1 (
    echo Word is not installed. Please install Word before running this script...
    exit /b
) else (
    echo Word is installed.
)

set foldername=Word_hello_world_sample
set /a counter=0

:loop
if exist %foldername% (
    set /a counter=counter + 1
    set foldername=Word_hello_world_sample_%counter%
    goto loop
)

office_addin_sample_scripts launch word_hello_world %foldername%

pause