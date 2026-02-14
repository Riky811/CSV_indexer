@echo off
REM index_folder.bat - wrapper to create an index (CSV) of a folder and optionally convert to XLSX.
REM Usage: index_folder.bat "C:\Path\To\Folder" "C:\Path\To\Output.csv" [/xlsx]

SETLOCAL ENABLEDELAYEDEXPANSION

REM If no folder provided, use current directory
if "%~1"=="" (
  set "TARGET=."
) else (
  set "TARGET=%~1"
)

rem Optional output path (CSV). If omitted, script auto-generates in current folder.
set "OUTPUT=%~2"

rem Optional flag to convert CSV to XLSX (requires Excel installed): /xlsx or -x
set "XFLAG="
if /I "%~3"=="/xlsx" set "XFLAG=-ToXlsx"
if /I "%~3"=="-xlsx" set "XFLAG=-ToXlsx"
if /I "%~3"=="/x" set "XFLAG=-ToXlsx"
if /I "%~3"=="-x" set "XFLAG=-ToXlsx"

rem Call the PowerShell script next to this .bat (same folder)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0index_folder.ps1" -TargetFolder "%TARGET%" -OutputPath "%OUTPUT%" %XFLAG%

if %ERRORLEVEL% EQU 0 (
  echo.
  echo Indexing complete.
) else (
  echo.
  echo PowerShell script exited with code %ERRORLEVEL%.
)

pause
ENDLOCAL
