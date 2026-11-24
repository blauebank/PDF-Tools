@echo off
REM Debug: Fenster bleibt nach dem Skript geöffnet (-NoExit)
set "SCRIPTDIR=%~dp0"
powershell.exe -NoExit -NoProfile -ExecutionPolicy Bypass -File "%SCRIPTDIR%CombinePDFs.ps1" -Collect -DebugMode "%~1"
