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


@REM where /q yo
@REM if ErrorLevel 1 (
@REM     echo Yeoman is not installed, installing now...
@REM     npm install -g yo
@REM     echo Yeoman has been installed.
@REM     echo Restarting script after installed Yeoman...
@REM     start "" "%~0"
@REM     exit /b
@REM ) else (
@REM     echo Yeoman is already installed!
@REM )

@REM Now Node.js, git have all prepared. Install Yeoman Office.

@REM echo Git and Node.js prepared. Checking Yeoman Office...
@REM yo --generators | findstr /C:"office"
@REM if ErrorLevel 1 (
@REM     echo Yeoman Office is not installed, installing now...
@REM     npm install -g yo generator-office
@REM     echo Yeoman Office has been installed.
@REM     echo Restarting script after installed Yeoman Office...
@REM     start "" "%~0"
@REM     exit /b
@REM ) else (
@REM     echo Yeoman Office has already been installed.
@REM )

@REM Now Yeoman Office has been installed. Create a sample project.

set foldername=Office_sample
set /a counter=0

:loop
if exist %foldername% (
    set /a counter=counter + 1
    set foldername=Office_sample_%counter%
    goto loop
)

@REM yo office --output %foldername% --projectType excel_sample --no-insight
office_addin_sample_scripts launch excel_mail %foldername%

pause