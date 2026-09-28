@echo off
chcp 65001 >nul
title Deshacer ultima carga - Migracion Colppy
color 0E
cd /d "%~dp0"

echo.
echo ================================================================
echo    DESHACER LA ULTIMA CARGA
echo ================================================================
echo.
echo   Usa esto cuando el Excel que se genero NO te sirvio
echo   y lo vas a tirar.
echo.
echo   Desmarca esas facturas del registro, asi las podes
echo   volver a procesar. Si no haces esto, quedan marcadas
echo   como ya cargadas y la proxima vez se saltean.
echo.
echo   No borra ningun archivo: el Excel lo borras vos.
echo.
echo ----------------------------------------------------------------
echo.

set "PY="
py -3 --version >nul 2>&1 && set "PY=py -3"
if not defined PY python --version >nul 2>&1 && set "PY=python"

if not defined PY (
    color 0C
    echo   [X] Python no esta instalado.
    echo.
    echo   Ejecuta primero:  "1 - INSTALAR (una sola vez).bat"
    echo.
    pause
    exit /b 1
)

%PY% "procesar_migracion.py" --deshacer

echo.
echo ================================================================
echo.
echo   Presiona una tecla para cerrar...
pause >nul
