@echo off
setlocal

REM MultiMonitorProfileTool Launcher
REM Robuster Start fuer Doppelklick und Autostart (Run-Key)

set "ScriptDir=%~dp0"
set "PSScript=%ScriptDir%MultiMonitorProfileTool.ps1"
set "ConfigPath=%ScriptDir%monitor-profiles.json"
set "PSExe=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

if not exist "%PSExe%" set "PSExe=powershell.exe"

if not exist "%PSScript%" (
    echo [Launcher] Skript nicht gefunden: "%PSScript%"
    exit /b 1
)

REM Normaler Doppelklick-Start: GUI wird sichtbar gezeigt.
REM Fuer Autostart (Run-Key) diese Zeile verwenden und -StartMinimized erganzen.
start "" "%PSExe%" -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%PSScript%" -ConfigPath "%ConfigPath%"
exit /b 0
