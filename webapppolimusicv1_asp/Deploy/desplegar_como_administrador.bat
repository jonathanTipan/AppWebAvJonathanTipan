@echo off
REM ===========================================================================
REM  PoliMusic - Despliegue en IIS
REM  Haz clic derecho sobre este archivo y elige "Ejecutar como administrador".
REM ===========================================================================
setlocal
cd /d "%~dp0"

echo.
echo ==========================================================
echo   PoliMusic - Despliegue en IIS
echo ==========================================================
echo.
echo Se copiara la aplicacion en C:\inetpub\wwwroot\PoliMusic
echo y se pedira confirmacion de UAC.
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0instalar_iis.ps1"

if errorlevel 1 (
  echo.
  echo Hubo un problema durante la instalacion. Revisa los mensajes de arriba.
)

endlocal
pause
