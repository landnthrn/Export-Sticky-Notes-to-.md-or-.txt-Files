@echo off
setlocal

set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "SCRIPT=%~dp0Sticky Notes Export.ps1"

"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"

echo.
pause

endlocal

