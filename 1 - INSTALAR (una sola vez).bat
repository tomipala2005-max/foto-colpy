@echo off
chcp 65001 >nul
title Instalacion - Procesador de Facturas
color 0B
cd /d "%~dp0"

echo.
echo ================================================================
echo    INSTALACION - PROCESADOR DE FACTURAS
echo ================================================================
echo.
echo  Esto se corre UNA SOLA VEZ. Despues usa el archivo:
echo     "2 - PROCESAR FACTURAS.bat"
echo.
echo ----------------------------------------------------------------
echo.

REM ---------- Buscar Python ----------
echo [1/4] Buscando Python...
echo.

set "PY="
py -3 --version >nul 2>&1 && set "PY=py -3"
if not defined PY python --version >nul 2>&1 && set "PY=python"

if not defined PY (
    color 0C
    echo   [X] Python NO esta instalado.
    echo.
    echo ================================================================
    echo    QUE HACER
    echo ================================================================
    echo.
    echo   1. Te voy a abrir la pagina de descarga de Python.
    echo.
    echo   2. Descarga el boton amarillo grande "Download Python".
    echo.
    echo   3. Al ejecutar el instalador, MUY IMPORTANTE:
    echo.
    echo         TILDA LA CASILLA DE ABAJO DE TODO QUE DICE
    echo         "Add python.exe to PATH"
    echo.
    echo      Si no la tildas, esto no va a funcionar.
    echo.
    echo   4. Dale "Install Now" y espera a que termine.
    echo.
    echo   5. Volve a ejecutar este mismo archivo.
    echo.
    echo ================================================================
    echo.
    pause
    start https://www.python.org/downloads/
    exit /b 1
)

for /f "tokens=*" %%v in ('%PY% --version 2^>^&1') do echo   [OK] %%v
echo.

REM ---------- Instalar librerias ----------
echo [2/4] Instalando librerias necesarias...
echo       (la primera vez puede tardar varios minutos, es normal)
echo.

%PY% -m pip install --disable-pip-version-check --upgrade pip >nul 2>&1
%PY% -m pip install --disable-pip-version-check openpyxl "xlrd==2.0.1" xlwt xlutils pypdf pdfplumber opencv-python-headless pypdfium2 pyzbar pytesseract
echo.

%PY% -c "import openpyxl, xlrd, xlwt, xlutils, pypdf, pdfplumber, cv2, pypdfium2, pytesseract" >nul 2>&1
if not errorlevel 1 goto :libs_ok

echo.
echo   No funciono a nivel sistema. Pruebo solo para tu usuario...
echo.
%PY% -m pip install --user --disable-pip-version-check openpyxl "xlrd==2.0.1" xlwt xlutils pypdf pdfplumber opencv-python-headless pypdfium2 pyzbar pytesseract
echo.

%PY% -c "import openpyxl, xlrd, xlwt, xlutils, pypdf, pdfplumber, cv2, pypdfium2, pytesseract" >nul 2>&1
if not errorlevel 1 goto :libs_ok

color 0C
echo ================================================================
echo   [X] Fallo la instalacion de las librerias
echo ================================================================
echo.
echo   Cerra esta ventana, hace clic DERECHO sobre este archivo
echo   y elegi "Ejecutar como administrador".
echo.
echo   Si sigue fallando, sacale una foto a esta pantalla.
echo.
pause
exit /b 1

:libs_ok
echo   [OK] Librerias instaladas.
echo.

REM ---------- Verificar LibreOffice ----------
echo [3/4] Buscando LibreOffice...
echo.

set "LO="
if exist "C:\Program Files\LibreOffice\program\soffice.exe" set "LO=1"
if exist "C:\Program Files (x86)\LibreOffice\program\soffice.exe" set "LO=1"
where soffice >nul 2>&1 && set "LO=1"

if not defined LO (
    color 0E
    echo   [!] LibreOffice NO esta instalado. Hace falta.
    echo.
    echo   Te voy a abrir la pagina de descarga.
    echo   Descargalo, instalalo con las opciones por defecto,
    echo   y volve a ejecutar este archivo.
    echo.
    echo   No hace falta que lo abras ni que lo uses,
    echo   solo tiene que estar instalado.
    echo.
    pause
    start https://www.libreoffice.org/download/download-libreoffice/
    exit /b 1
)
echo   [OK] LibreOffice encontrado.
echo.

REM ---------- Verificar el OCR (opcional, para PDF escaneados) ----------
echo [4/4] Verificando el lector de PDF y el OCR...
echo.
%PY% -c "import leer_pdf; ok, m = leer_pdf.estado_ocr(); print('   OCR para PDF escaneados: ' + ('ACTIVADO  ' if ok else 'NO disponible  ') + m)" 2>nul
if errorlevel 1 echo   [!] No pude verificar el OCR. Revisa que leer_pdf.py este en esta carpeta.
echo.
echo   El OCR es opcional: solo hace falta para PDF que son una foto.
echo   Si dice NO disponible, mira el paso 4 del LEEME.md
echo.

color 0A
echo ================================================================
echo    LISTO - TODO INSTALADO
echo ================================================================
echo.
echo   Ya podes usar "2 - PROCESAR FACTURAS.bat"
echo.
echo   No necesitas volver a correr esta instalacion.
echo.
echo ================================================================
echo.
pause
