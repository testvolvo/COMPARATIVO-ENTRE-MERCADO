@echo off
chcp 65001 >nul
set "SCRIPT=%~dp0actualizar_comparativo.ps1"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"
set "CODE=%ERRORLEVEL%"
echo.
if not "%CODE%"=="0" (
  echo La actualizacion termino con errores.
) else (
  echo Actualizacion finalizada.
)
pause
exit /b %CODE%
