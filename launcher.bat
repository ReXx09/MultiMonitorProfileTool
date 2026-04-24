@echo off
REM MultiMonitorProfileTool Launcher
REM Startet das PowerShell-Skript ohne sichtbares Fenster

set PSScript=%~dp0MultiMonitorProfileTool.ps1
set ConfigPath=%~dp0monitor-profiles.json

REM Fenster versteckt starten mit /b, nicht warten mit /i
start /b powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -STA -File "%PSScript%" -ConfigPath "%ConfigPath%"
