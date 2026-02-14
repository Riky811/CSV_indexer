@echo off
REM convert_csvs_to_xlsx.bat - wrapper to run convert_csvs_to_xlsx.ps1 with ExecutionPolicy Bypass
REM Usage: convert_csvs_to_xlsx.bat -Root "E:\Indexes" -Recursive -Force

SETLOCAL
SET "SCRIPT_DIR=%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%convert_csvs_to_xlsx.ps1" %*
ENDLOCAL
