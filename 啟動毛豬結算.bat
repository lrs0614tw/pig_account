@echo off
setlocal
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0scripts\launcher.ps1" -Action run
set "PIG_EXIT=%ERRORLEVEL%"
if /i "%~1"=="--no-pause" exit /b %PIG_EXIT%
if not "%PIG_EXIT%"=="0" echo Failed. See the message above and the logs folder.
if not "%PIG_EXIT%"=="0" pause
exit /b %PIG_EXIT%
