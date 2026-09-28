@echo off
chcp 65001 >nul
title Procesador de Facturas - Migracion Colppy
color 0B
cd /d "%~dp0"
setlocal enabledelayedexpansion

echo.
echo ================================================================
echo    PROCESADOR DE FACTURAS  --^>  MIGRACION COLPPY
echo ================================================================
echo.
echo   Excel de ARCA  --^>  subcarpeta "comprobantes arca"
echo   Facturas PDF   --^>  subcarpeta "facturas pdf"
echo   Resultados     --^>  subcarpeta "migraciones generadas"
echo.
echo   (o arrastra los archivos directamente encima de este .bat)
echo.

REM ---------- Carpetas de trabajo ----------
if not exist "comprobantes arca" md "comprobantes arca"
if not exist "facturas pdf" md "facturas pdf"
if not exist "migraciones generadas" md "migraciones generadas"

REM ---------- Buscar Python ----------
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

REM ---------- Verificar librerias (y instalarlas si faltan) ----------
set "LIBS=openpyxl "xlrd==2.0.1" xlwt xlutils pypdf pdfplumber opencv-python-headless pypdfium2 pyzbar pytesseract"

%PY% -c "import openpyxl, xlrd, xlwt, xlutils, pypdf, pdfplumber, cv2, pypdfium2, pytesseract" >nul 2>&1
if not errorlevel 1 goto :libs_ok

color 0E
echo   Faltan librerias. Las instalo ahora, esperame un momento...
echo   (la primera vez puede tardar varios minutos)
echo.
%PY% -m pip install --disable-pip-version-check %LIBS%
echo.

%PY% -c "import openpyxl, xlrd, xlwt, xlutils, pypdf, pdfplumber, cv2, pypdfium2, pytesseract" >nul 2>&1
if not errorlevel 1 goto :libs_instaladas

echo.
echo   No funciono a nivel sistema. Pruebo solo para tu usuario...
echo.
%PY% -m pip install --user --disable-pip-version-check %LIBS%
echo.

%PY% -c "import openpyxl, xlrd, xlwt, xlutils, pypdf, pdfplumber, cv2, pypdfium2, pytesseract" >nul 2>&1
if not errorlevel 1 goto :libs_instaladas

color 0C
echo.
echo ================================================================
echo   [X] No se pudieron instalar las librerias
echo ================================================================
echo.
echo   Proba esto:
echo.
echo   1. Cerra esta ventana.
echo   2. Clic DERECHO sobre "1 - INSTALAR (una sola vez).bat"
echo   3. Elegi "Ejecutar como administrador"
echo.
echo   Si sigue fallando, sacale una foto a esta pantalla
echo   y mandamela.
echo.
pause
exit /b 1

:libs_instaladas
color 0B
echo   [OK] Librerias instaladas. Sigo con el proceso.
echo.

:libs_ok

REM ---------- Si arrastraste archivos, procesarlos ----------
if not "%~1"=="" (
    echo   Procesando los archivos que arrastraste...
    echo.
    %PY% "procesar_migracion.py" %*
    goto :fin
)

REM ---------- Buscar archivos automaticamente ----------
echo   Buscando archivos para procesar...
echo.

set "N=0"

REM Excel de ARCA en la subcarpeta "comprobantes arca"
for %%f in ("comprobantes arca\*.xlsx") do (
    set /a N+=1
    set "F!N!=%%~ff"
    echo     !N!^) %%~nxf   [Excel de ARCA]
)

REM Facturas PDF en la subcarpeta "facturas pdf"
for %%f in ("facturas pdf\*.pdf") do (
    set /a N+=1
    set "F!N!=%%~ff"
    echo     !N!^) %%~nxf   [PDF]
)

if "!N!"=="0" (
    color 0E
    echo   No encontre ningun archivo.
    echo.
    echo ----------------------------------------------------------------
    echo   COMO USARLO
    echo ----------------------------------------------------------------
    echo.
    echo   Opcion A ^(la mas facil^):
    echo     Arrastra el archivo ^(Excel de ARCA o factura PDF^)
    echo     y soltalo ENCIMA de este mismo .bat
    echo     Podes arrastrar varios PDF juntos: van todos a la
    echo     misma planilla.
    echo.
    echo   Opcion B:
    echo     Copia el Excel de ARCA a la subcarpeta   "comprobantes arca"
    echo     o las facturas PDF a la subcarpeta       "facturas pdf"
    echo     y volve a ejecutar esto.
    echo.
    echo   Lo que se genera queda en la subcarpeta "migraciones generadas".
    echo.
    echo   Carpeta actual:
    echo     %CD%
    echo.
    pause
    exit /b 0
)

echo.
echo ----------------------------------------------------------------
echo.

if "!N!"=="1" (
    echo   Encontre 1 archivo. Lo proceso.
    echo.
    %PY% "procesar_migracion.py" "!F1!"
    goto :fin
)

echo   Encontre !N! archivos.
echo.
echo     Escribi el numero del que queres procesar
echo     o escribi  T  para procesar TODOS
echo.
set /p "OPCION=  Tu eleccion: "
echo.

if /i "!OPCION!"=="T" (
    set "TODOS="
    for /l %%i in (1,1,!N!) do set "TODOS=!TODOS! "!F%%i!""
    %PY% "procesar_migracion.py" !TODOS!
    goto :fin
)

set "ELEGIDO=!F%OPCION%!"
if not defined ELEGIDO (
    color 0C
    echo   Opcion invalida.
    echo.
    pause
    exit /b 1
)

%PY% "procesar_migracion.py" "!ELEGIDO!"

:fin
echo.
echo ================================================================
echo    PROCESO TERMINADO
echo ================================================================
echo.
echo   Los archivos generados estan en:
echo   %CD%\migraciones generadas
echo.
echo   Si Excel se queja de que la ruta es muy larga,
echo   copia el archivo al Escritorio y abrilo desde ahi.
echo.
echo ================================================================
echo.
echo   Presiona una tecla para cerrar...
pause >nul
