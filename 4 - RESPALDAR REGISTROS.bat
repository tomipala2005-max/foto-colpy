@echo off
chcp 65001 >nul
title Respaldar registros - Migracion Colppy
color 0A
cd /d "%~dp0"

echo.
echo ================================================================
echo    COPIA DE SEGURIDAD DE LOS REGISTROS
echo ================================================================
echo.
echo   Copia todo lo que hay en "json arch" a una subcarpeta
echo   con la fecha y hora de hoy.
echo.
echo   Ahi vive el control de que facturas ya cargaste.
echo   Si se pierde, no hay forma de reconstruirlo.
echo.
echo   Igual se hace una copia sola cada vez que procesas
echo   facturas: esto es por si la queres hacer a mano.
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

%PY% "procesar_migracion.py" --respaldar

echo.
echo ================================================================
echo.
echo   Presiona una tecla para cerrar...
pause >nul
