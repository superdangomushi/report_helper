@echo off
rem install.bat - Double-click this to set up report_helper on Windows.
rem PowerShell の ExecutionPolicy を一時的にバイパスして install.ps1 を実行します。
setlocal
set "PS1=%~dp0install.ps1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS1%"
echo.
pause
endlocal
