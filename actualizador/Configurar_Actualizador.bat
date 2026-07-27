@echo off
chcp 65001 >nul
set "SCRIPT=%~dp0actualizar_comparativo.ps1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Configurar
set "CODE=%ERRORLEVEL%"
echo.
pause
exit /b %CODE%
