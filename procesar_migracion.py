#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
 PROCESADOR DE MIGRACION DE FACTURAS DE COMPRA -> COLPPY
 MEPANO SOCIEDAD ANONIMA (CUIT 30710230184) / METALURGICA PABLO NOGUES SRL
 (CUIT 30554641748)
=============================================================================

QUE HACE
--------
Toma un archivo "Mis Comprobantes Recibidos" descargado de ARCA/AFIP y genera
la planilla de migracion lista para importar en Colppy, aplicando todas las
reglas de negocio acordadas.

USO
---
    python procesar_migracion.py "Mis Comprobantes Recibidos - CUIT 30710230184.xlsx"

    # o varios archivos de una:
    python procesar_migracion.py archivo1.xlsx archivo2.xlsx

    # sin actualizar el registro de facturas (prueba en seco):
    python procesar_migracion.py archivo.xlsx --dry-run

La empresa (MEPANO o MPN) se detecta automaticamente por el CUIT del receptor.

SALIDA
------
    Migracion/migracion factura <EMPRESA> <dd-mm-aaaa>.xls    <- Excel formato original
    Migracion/json arch/_registro_*.json                      <- registros de control

    El CSV NO se genera: se revisa el Excel y despues se guarda a mano la hoja
    PASO 2 como CSV delimitado por comas.

REQUISITOS
----------
    pip install openpyxl xlrd==2.0.1 xlwt xlutils
    LibreOffice instalado (para recalcular las formulas de la plantilla)

=============================================================================
 REGLAS DE NEGOCIO IMPLEMENTADAS  (acumuladas de todas las cargas anteriores)
=============================================================================

 1. NO DUPLICAR FACTURAS
    Se lleva un registro de facturas ya procesadas (CUIT emisor + numero de
    factura). Cualquier comprobante ya cargado se saltea automaticamente.
    Registro exacto: (CUIT, "PPPPP-NNNNNNNN")
    Registro difuso: (CUIT, fecha, total) -> respaldo por si falta el numero.

 2. CONVERSION DE MONEDA (solo en PASO 1, nunca en PASO 2)
    Si Moneda != "$", se multiplican TODAS las columnas monetarias de esa fila
    por el Tipo de Cambio, y se deja Tipo Cambio = 1 y Moneda = "$".
    Resultado: en PASO 1 solo quedan "1" y "$", sin saldos en dolares.

 3. FECHA DE VENCIMIENTO (PASO 2)
    Fecha de Vencimiento = Fecha Factura + 30 dias.

 4. CUENTA DE GASTO (PASO 2)
    Se busca el codigo por CUIT del proveedor en CODIGOS_A_USAR.xls:
        - MEPANO -> hoja "CODIGO PROVEEDORES MSA"
        - MPN    -> hoja "CODIGO PROVEEDORES MPN"
    Si el proveedor no existe en la tabla, se usa el fallback de "Compra de
    Materias Primas" y se reporta al final para que lo revises.

 5. REGLAS DE PERCEPCION DE IIBB  (se evaluan en este orden)

    a) FACTURA SIN DESGLOSE (tipicamente Factura B o C)
       Cuando la plantilla vuelca el total tanto en Neto No Gravado como en
       Percepcion IIBB, no es una percepcion real: es un artefacto de las
       formulas. Se anula la percepcion y queda todo en No Gravado.

    b) ESTACION DE SERVICIO -> Neto No Gravado, sin jurisdiccion
       (OPESA, Operadora de Estaciones de Servicio, etc. Como la percepcion
        se va entera a No Gravado, el IIBB queda en 0 y la jurisdiccion tiene
        que quedar vacia.)

    c) AUTOPISTAS / PEAJES  -> Neto No Gravado, sin jurisdiccion
       (Autopistas del Sol, Autopistas Urbanas, Rutas Sur Atlantico,
        Grupo Concesionario del Oeste, etc.)

    d) CENCOSUD             -> Neto No Gravado

    e) SANCOR SEGUROS       -> Neto No Gravado, sin IIBB (no cobra IIBB)

    f) CEAMSE               -> Neto No Gravado, sin IIBB (no cobra IIBB)
       (Coordinacion Ecologica Area Metropolitana Sociedad del Estado)

    g) PERCEPCION MENOR A $30
       Suele ser un resto de redondeo de la conversion de dolares.
       Se pliega dentro del Neto Gravado: si es negativa resta, si es
       positiva suma. La percepcion queda en 0.

    h) PERCEPCION REAL (resto de los casos)
       Se deja la percepcion y se completa Jurisdiccion = "Buenos Aires"
       por defecto. Se reporta al final como PENDIENTE DE VERIFICAR, porque
       podria corresponder CABA.

 6. LIMPIEZA DE PASO 2
    Todas las filas debajo de la ultima factura quedan completamente vacias:
    sin ceros residuales y sin fechas fantasma.

 7. FORMATO IDENTICO AL ORIGINAL  (critico)
    El .xls de salida conserva exactamente el formato de la plantilla:
    dimensiones, anchos de columna, celdas combinadas, formatos de numero y
    alturas de fila.

    OJO - bug conocido de xlutils: al copiar el libro, xlutils crea registros
    de altura explicitos (255) para filas que en el original NO tenian ninguno
    (usaban la altura por defecto de la hoja). Eso cambia como se ve el archivo.
    La funcion _corregir_alturas_default() lo revierte. NO LA SAQUES.

 8. SOLO EXCEL
    El script genera unicamente el .xls, para que puedas revisarlo antes de
    importar. El CSV lo guardas vos a mano desde Excel:
    Archivo > Guardar como > CSV (delimitado por comas), sobre la hoja PASO 2.

    (Quedo la funcion exportar_csv() por si algun dia lo queres automatico:
     alcanza con volver a llamarla desde procesar().)

 9. CARPETAS
    comprobantes arca/     los .xlsx que bajas de ARCA
    facturas pdf/          las facturas sueltas en PDF
    migraciones generadas/ los .xls que genera el script
    json arch/             los registros de control (duplicados, proveedores)

    Las cuatro se crean solas la primera vez que corre el script.

10. FACTURAS EN PDF
    Ademas del export de ARCA, el script acepta facturas sueltas en PDF
    (ver leer_pdf.py). Los datos de cabecera salen del codigo QR de AFIP, que
    viene firmado y no se puede leer mal; el desglose impositivo sale del texto
    del PDF. Antes de cargar nada se muestra todo en pantalla para revisar y
    corregir. Si el PDF es un escaneo sin texto, se piden los datos a mano.

    Los PDF pasan por exactamente el mismo circuito que los comprobantes de
    ARCA: control de duplicados, conversion de moneda, reglas de IIBB, cuenta
    de gasto y escritura del .xls con el formato original.
"""

import sys
import os
import csv
import json
import pickle
import shutil
import datetime
import re
import subprocess
import tempfile
import argparse

try:
    import openpyxl
    import xlrd
    import xlwt
    from xlutils.copy import copy as xl_copy
except ImportError as e:
    sys.exit("Falta una dependencia: %s\n"
             "Instalala con:  pip install openpyxl xlrd==2.0.1 xlwt xlutils" % e)

# Lector de facturas en PDF (opcional: si no esta, se procesan solo los .xlsx).
# Se atrapa cualquier error, no solo ImportError: si leer_pdf.py llegara a tener
# un problema, el circuito de ARCA (que es el de todos los dias) tiene que
# seguir funcionando igual.
leer_pdf = None
_ERROR_LEER_PDF = None
try:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import leer_pdf
except Exception as _e:
    _ERROR_LEER_PDF = _e


# =============================================================================
# CONFIGURACION
# =============================================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RECURSOS_DIR = os.path.join(BASE_DIR, "_recursos")
JSON_DIR = os.path.join(BASE_DIR, "json arch")

# Carpetas de trabajo (se crean solas si no existen)
ENTRADA_ARCA = os.path.join(BASE_DIR, "comprobantes arca")   # los .xlsx de ARCA
ENTRADA_PDF = os.path.join(BASE_DIR, "facturas pdf")         # las facturas en PDF
SALIDA_DIR = os.path.join(BASE_DIR, "migraciones generadas")  # los .xls que salen

PLANTILLA = os.path.join(RECURSOS_DIR, "plantilla_migracion.xls")
CODIGOS = os.path.join(RECURSOS_DIR, "CODIGOS_A_USAR.xls")

HOJA_P1 = "PASO 1>COPIAR MIS COMPROBANTES"
HOJA_P2 = "PASO 2> GUARDAR COMO CSV"

FILA_INICIO = 3          # primera fila de datos (1-indexada)
COLS_P1 = 30             # columnas A..AD de Mis Comprobantes
COLS_P2 = 22             # columnas A..V de PASO 2

# Indices 0-based dentro de la fila de "Mis Comprobantes"
IDX_FECHA = 0
IDX_TIPO = 1
IDX_PTO_VENTA = 2
IDX_NRO_DESDE = 3
IDX_CUIT_EMISOR = 7
IDX_DENOM_EMISOR = 8
IDX_CUIT_RECEPTOR = 10
IDX_TIPO_CAMBIO = 11
IDX_MONEDA = 12
IDX_TOTAL = 29
COLS_MONETARIAS = list(range(13, 30))   # Neto Grav. IVA 0% .. Imp. Total

# Columnas 1-based de PASO 2
C_TIPO = 1
C_LETRA = 2
C_RAZON_SOCIAL = 3
C_CUIT = 4
C_FECHA_FACTURA = 7
C_FECHA_VENC = 8
C_TOTAL = 10
C_NETO_GRAVADO = 14
C_NETO_NO_GRAVADO = 15
C_PERC_IVA = 16
C_PERC_IIBB = 17
C_JURISDICCION = 18
C_CUENTA_GASTO = 19

DIAS_VENCIMIENTO = 30
UMBRAL_IIBB_MENOR = 30.0     # percepciones menores a esto se pliegan en Neto Gravado

EMPRESAS = {
    "30710230184": {
        "nombre": "MEPANO",
        "razon_social": "MEPANO SOCIEDAD ANONIMA",
        "hoja_codigos": "CODIGO PROVEEDORES MSA",
        "cuenta_fallback": 511200,            # Cpra. Mat. Primas y Mater.
        "registro_exact": "_registro_MEPANO_exact.json",
        "registro_fuzzy": "_registro_MEPANO_fuzzy.json",
    },
    "30554641748": {
        "nombre": "MPN",
        "razon_social": "METALURGICA PABLO NOGUES SRL",
        "hoja_codigos": "CODIGO PROVEEDORES MPN",
        "cuenta_fallback": 51511200,          # Cpra. Mat. Primas y Mater. (plan MPN)
        "registro_exact": "_registro_MPN_exact.json",
        "registro_fuzzy": "_registro_MPN_fuzzy.json",
    },
}

# Proveedores con tratamiento especial de IIBB (se buscan como subcadena, en MAYUSCULAS)
NOMBRES_ESTACION_SERVICIO = ["ESTACION DE SERVICIO", "ESTACIONES DE SERVICIO",
                             "OPESA"]
NOMBRES_AUTOPISTA = ["AUTOPISTA", "RUTAS SUR", "CONCESIONARIO DEL OESTE"]
NOMBRES_CENCOSUD = ["CENCOSUD"]
NOMBRES_SANCOR = ["SANCOR"]
NOMBRES_CEAMSE = ["COORDINACION ECOLOGICA", "CEAMSE"]

# Codigos de cuenta resueltos manualmente que no figuran en CODIGOS_A_USAR.xls
CUENTAS_EXTRA_MSA = {
    30717463907: 511200,   # LUCILA LAURO Y NICOLAS LAURO - sin rubro hallado -> mat. primas
    30711528144: 522102,   # BENETTI DISTRIBUCIONES - mayorista baterias/repuestos
    20222970726: 522201,   # BADINO GERMAN SANTIAGO - reparacion art. electricos
    20118379110: 511200,   # ROSSI GABRIEL HECTOR - ferreteria / mat. electricos
}
CUENTAS_EXTRA_MPN = {}

# Tipos de comprobante que la hoja "Tablas" de la plantilla no traduce (quedan
# sin Tipo ni Letra en PASO 2 y Colppy rechaza la fila). Codigo AFIP -> (Tipo
# Colppy, Letra). El 17 es la Liquidacion de Servicios Publicos de AySA.
TIPOS_SIN_TABLA = {
    17: ("FAC", "A"),    # Liquidacion de Servicios Publicos Clase A
    18: ("FAC", "B"),    # Liquidacion de Servicios Publicos Clase B
}

# Detalle de percepciones de las facturas leidas de PDF, por clave de factura.
# El Excel de ARCA junta todo en "Otros Tributos" y la plantilla lo manda
# entero a Percepcion de IIBB; del PDF se sabe cuanto es percepcion de IVA y
# cuanto son tasas no gravadas, y aplicar_reglas() lo reparte.
DETALLE_PDF = {}

# Registro de nombres de proveedor por CUIT, tal como los escribe AFIP.
# Se completa solo cada vez que se procesa un export de ARCA, y se usa para que
# las facturas cargadas desde un PDF queden con el mismo nombre de siempre.
PROVEEDORES = "_proveedores.json"

# Deja constancia de la ultima carga, para poder deshacerla si el Excel
# generado no sirvio. Sin esto, las facturas quedan marcadas como cargadas
# aunque hayas tirado el archivo a la basura.
ULTIMA_CORRIDA = "_ultima_corrida.json"

# Copias de seguridad de los registros. Se hace una en cada corrida, ANTES de
# tocar nada, y se conservan las ultimas. El registro de duplicados no se puede
# reconstruir si se pierde, asi que conviene que sobren.
RESPALDO_DIR = "respaldo"
RESPALDOS_A_CONSERVAR = 20


# =============================================================================
# UTILIDADES
# =============================================================================

def log(msg=""):
    print(msg, flush=True)


def parse_fecha(v):
    """Convierte la fecha de Mis Comprobantes (dd/mm/aaaa o date) a datetime.date."""
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    d, m, y = str(v).strip().split("/")
    return datetime.date(int(y), int(m), int(d))


def clave_factura(vals):
    """Clave unica de factura: (CUIT emisor, 'PPPPP-NNNNNNNN')."""
    pv = int(vals[IDX_PTO_VENTA])
    nd = int(vals[IDX_NRO_DESDE])
    cuit = str(int(vals[IDX_CUIT_EMISOR]))
    return (cuit, "%05d-%08d" % (pv, nd))


def cargar_json(path, default):
    if not os.path.exists(path):
        return default
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def guardar_json(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def buscar_libreoffice():
    for cmd in ("soffice", "libreoffice"):
        ruta = shutil.which(cmd)
        if ruta:
            return ruta
    # rutas tipicas en Windows
    for ruta in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                 r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if os.path.exists(ruta):
            return ruta
    sys.exit("No encontre LibreOffice. Es necesario para recalcular las formulas "
             "de la plantilla.\nInstalalo desde https://www.libreoffice.org/")


def convertir(soffice, entrada, salida_dir, filtro="xlsx"):
    """Convierte un archivo con LibreOffice (tambien fuerza el recalculo de formulas)."""
    os.makedirs(salida_dir, exist_ok=True)
    subprocess.run([soffice, "--headless", "--convert-to", filtro,
                    "--outdir", salida_dir, entrada],
                   check=True, capture_output=True, timeout=180)
    nombre = os.path.splitext(os.path.basename(entrada))[0] + ".xlsx"
    return os.path.join(salida_dir, nombre)


# =============================================================================
# LECTURA DE DATOS
# =============================================================================

def leer_comprobantes(path):
    """Lee 'Mis Comprobantes Recibidos'. Encabezado en fila 2, datos desde la 3."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    headers = [ws.cell(row=2, column=c).value for c in range(1, COLS_P1 + 1)]
    filas = []
    for r in range(3, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, COLS_P1 + 1)]
        if all(v is None for v in vals):
            continue
        filas.append(vals)
    return headers, filas


def respaldar_registros(avisar=True):
    """
    Copia todos los .json de control a "json arch/respaldo/<fecha hora>/".
    Se llama en cada corrida antes de modificar nada, y tambien a mano desde
    "4 - RESPALDAR REGISTROS.bat".
    """
    if not os.path.isdir(JSON_DIR):
        return None

    archivos = [f for f in os.listdir(JSON_DIR)
                if f.lower().endswith(".json")
                and os.path.isfile(os.path.join(JSON_DIR, f))]
    if not archivos:
        if avisar:
            log("No hay registros para respaldar todavia.")
        return None

    raiz = os.path.join(JSON_DIR, RESPALDO_DIR)
    destino = os.path.join(raiz,
                           datetime.datetime.now().strftime("%Y-%m-%d %H%M%S"))
    try:
        os.makedirs(destino, exist_ok=True)
        for nombre in archivos:
            shutil.copy2(os.path.join(JSON_DIR, nombre),
                         os.path.join(destino, nombre))
    except Exception as e:
        log("!! No pude hacer la copia de seguridad: %s" % e)
        return None

    # se conservan solo las ultimas copias
    try:
        copias = sorted(d for d in os.listdir(raiz)
                        if os.path.isdir(os.path.join(raiz, d)))
        for vieja in copias[:-RESPALDOS_A_CONSERVAR]:
            shutil.rmtree(os.path.join(raiz, vieja), ignore_errors=True)
    except Exception:
        pass

    if avisar:
        log("Copia de seguridad: %d archivo(s) en" % len(archivos))
        log("   %s" % destino)
    return destino


def cargar_proveedores():
    """Mapa CUIT -> nombre del proveedor, como lo escribe AFIP."""
    crudo = cargar_json(os.path.join(JSON_DIR, PROVEEDORES), {})
    mapa = {}
    for k, v in crudo.items():
        try:
            mapa[int(k)] = v
        except (TypeError, ValueError):
            continue
    return mapa


def actualizar_proveedores(filas):
    """Guarda los nombres que trae un export de ARCA, para reutilizarlos luego."""
    path = os.path.join(JSON_DIR, PROVEEDORES)
    mapa = cargar_json(path, {})
    for vals in filas:
        cuit = vals[IDX_CUIT_EMISOR]
        nombre = vals[IDX_DENOM_EMISOR]
        if cuit and nombre:
            mapa[str(int(cuit))] = str(nombre).strip()
    guardar_json(path, mapa)


def leer_pdfs(paths, interactivo=True):
    """
    Lee una o varias facturas en PDF y las devuelve como filas identicas a las
    de 'Mis Comprobantes'. Cada factura se muestra en pantalla para revisar
    antes de aceptarla.
    """
    if leer_pdf is None:
        if _ERROR_LEER_PDF is not None:
            sys.exit("No pude cargar leer_pdf.py: %s\n"
                     "Los Excel de ARCA se siguen procesando igual."
                     % _ERROR_LEER_PDF)
        sys.exit("Falta el archivo leer_pdf.py, que es el que lee los PDF.")
    faltan = leer_pdf.dependencias_faltantes()
    if faltan:
        sys.exit("Para leer PDF faltan librerias: %s\n"
                 "Instalalas con:  pip install pypdf pdfplumber opencv-python"
                 % ", ".join(faltan))

    # el lector de PDF necesita saber cuales son tus empresas, para poder
    # preguntar "MEPANO o MPN?" cuando no logra leer el CUIT del cliente
    leer_pdf.registrar_empresas(EMPRESAS)

    # ...y la lista de proveedores, para completar solo el nombre por CUIT
    nombres = cargar_proveedores()
    leer_pdf.registrar_proveedores(nombres)

    filas = []
    for path in paths:
        log()
        log("-" * 78)
        log("Leyendo PDF: %s" % os.path.basename(path))
        datos = leer_pdf.procesar_pdf(path, nombres, interactivo=interactivo)
        if datos is None:
            log("   Salteada.")
            continue
        # el proveedor confirmado queda guardado: la proxima vez que aparezca
        # ese CUIT, el nombre se completa solo
        cuit = datos.get("cuit_emisor")
        nombre = (datos.get("razon_social") or "").strip()
        if cuit and nombre and int(cuit) not in nombres:
            nombres[int(cuit)] = nombre
            mapa = cargar_json(os.path.join(JSON_DIR, PROVEEDORES), {})
            mapa[str(int(cuit))] = nombre
            guardar_json(os.path.join(JSON_DIR, PROVEEDORES), mapa)
            leer_pdf.registrar_proveedores(nombres)
        fila = leer_pdf.a_fila_arca(datos)
        DETALLE_PDF[clave_factura(fila)] = {
            "percepcion_iva": datos.get("percepcion_iva") or 0.0,
            "tasas_no_gravadas": datos.get("tasas_no_gravadas") or 0.0,
            "total": datos.get("total"),
        }
        filas.append(fila)
    return filas


def detectar_empresa(filas):
    """Detecta la empresa por el CUIT del receptor."""
    cuits = set()
    for vals in filas:
        v = vals[IDX_CUIT_RECEPTOR]
        if v is not None:
            cuits.add(str(int(v)))
    if len(cuits) != 1:
        sys.exit("No pude determinar la empresa. CUIT(s) de receptor hallados: %s" % cuits)
    cuit = cuits.pop()
    if cuit not in EMPRESAS:
        sys.exit("CUIT receptor %s desconocido. Empresas configuradas: %s"
                 % (cuit, list(EMPRESAS)))
    return cuit, EMPRESAS[cuit]


def cargar_codigos(hoja, extras):
    """Mapa CUIT -> codigo de cuenta contable desde CODIGOS_A_USAR.xls."""
    if not os.path.exists(CODIGOS):
        sys.exit("No encuentro %s" % CODIGOS)
    libro = xlrd.open_workbook(CODIGOS)
    sh = libro.sheet_by_name(hoja)
    mapa = {}
    # encabezado en la fila 2 (0-indexada: 1); datos desde la 3
    for r in range(2, sh.nrows):
        try:
            cuit = sh.cell_value(r, 0)
            codigo = sh.cell_value(r, 1)
        except IndexError:
            continue
        if not cuit or not codigo:
            continue
        try:
            cuit = int(float(cuit))
            codigo = int(float(codigo))
        except (ValueError, TypeError):
            continue
        if codigo == 0:
            continue
        mapa[cuit] = codigo
    mapa.update(extras)
    return mapa


# =============================================================================
# PASO 1 - CONVERSION DE MONEDA
# =============================================================================

def convertir_moneda(filas):
    """Pasa a pesos toda fila en moneda extranjera. Devuelve (filas, detalle)."""
    convertidas = []
    detalle = []
    for vals in filas:
        vals = list(vals)
        tc = vals[IDX_TIPO_CAMBIO]
        moneda = vals[IDX_MONEDA]
        es_extranjera = (moneda and str(moneda).strip().upper() != "$"
                         and tc and float(tc) != 1.0)
        if es_extranjera:
            original = vals[IDX_TOTAL]
            for idx in COLS_MONETARIAS:
                if vals[idx] is None:
                    continue
                vals[idx] = round(float(vals[idx]) * float(tc), 2)
            vals[IDX_TIPO_CAMBIO] = 1
            vals[IDX_MONEDA] = "$"
            detalle.append((vals[IDX_DENOM_EMISOR], moneda, tc, original, vals[IDX_TOTAL]))
        vals[IDX_FECHA] = parse_fecha(vals[IDX_FECHA])
        convertidas.append(vals)
    return convertidas, detalle


# =============================================================================
# PASO 2 - REGLAS DE NEGOCIO
# =============================================================================

def aplicar_reglas(ws2, valores, nfilas, mapa_codigos, cuenta_fallback,
                   filas_p1=None, detalles=None):
    """
    Vuelca los resultados calculados en PASO 2 y aplica todas las reglas.
    filas_p1: las filas de PASO 1, para completar el tipo de comprobante.
    detalles: por fila, el detalle de percepciones de las facturas en PDF
              (None para las que vienen de ARCA).
    Devuelve (log_reglas, sin_codigo, iibb_a_verificar).
    """
    reglas = []
    sin_codigo = []
    verificar = []
    cols = list(range(1, COLS_P2 + 1))

    for i in range(nfilas):
        r = FILA_INICIO + i

        # 1) materializar los valores calculados por las formulas
        fila = {c: valores.cell(row=r, column=c).value for c in cols}
        for c in cols:
            v = fila[c]
            if isinstance(v, float):
                v = round(v, 2)
            ws2.cell(row=r, column=c).value = v

        nombre = fila[C_RAZON_SOCIAL] or ""
        nombre_up = str(nombre).upper()
        cuit = fila[C_CUIT]
        neto_gravado = fila[C_NETO_GRAVADO] or 0
        neto_no_gravado = fila[C_NETO_NO_GRAVADO] or 0
        perc_iibb = fila[C_PERC_IIBB] or 0
        if isinstance(perc_iibb, float):
            perc_iibb = round(perc_iibb, 2)
        total = fila[C_TOTAL]
        total = float(total) if total not in (None, "") else 0

        # 1b) tipo y letra que la tabla de la plantilla no conoce
        if filas_p1 is not None and not fila[C_TIPO]:
            m = re.match(r"\s*(\d+)", str(filas_p1[i][IDX_TIPO] or ""))
            extra = TIPOS_SIN_TABLA.get(int(m.group(1))) if m else None
            if extra:
                ws2.cell(row=r, column=C_TIPO).value = extra[0]
                ws2.cell(row=r, column=C_LETRA).value = extra[1]
                reglas.append((r, nombre, "tipo %s -> %s %s"
                               % (m.group(1), extra[0], extra[1]), 0))

        # 1c) facturas en PDF: la percepcion de IVA y las tasas no gravadas
        #     no son percepcion de IIBB. Solo si el detalle sigue cerrando con
        #     lo que quedo en IIBB (si lo corregiste a mano en la revision del
        #     PDF, no se toca).
        det = detalles[i] if detalles else None
        if det and perc_iibb:
            escala = total / det["total"] if det.get("total") else 1.0
            p_iva = round((det.get("percepcion_iva") or 0) * escala, 2)
            tasas = round((det.get("tasas_no_gravadas") or 0) * escala, 2)
            if (p_iva or tasas) and perc_iibb - p_iva - tasas >= -0.05:
                previa = fila[C_PERC_IVA] if isinstance(fila[C_PERC_IVA], (int, float)) else 0
                ws2.cell(row=r, column=C_PERC_IVA).value = round(previa + p_iva, 2)
                neto_no_gravado = round(neto_no_gravado + tasas, 2)
                ws2.cell(row=r, column=C_NETO_NO_GRAVADO).value = neto_no_gravado
                perc_iibb = round(perc_iibb - p_iva - tasas, 2)
                if abs(perc_iibb) < 0.01:
                    perc_iibb = 0
                ws2.cell(row=r, column=C_PERC_IIBB).value = perc_iibb
                reglas.append((r, nombre, "PDF: Perc. IVA %.2f / tasas a No Gravado %.2f"
                               % (p_iva, tasas), p_iva + tasas))

        # 2) fecha de vencimiento = fecha factura + 30 dias
        fecha_fac = fila[C_FECHA_FACTURA]
        if isinstance(fecha_fac, datetime.datetime):
            ws2.cell(row=r, column=C_FECHA_VENC).value = (
                fecha_fac + datetime.timedelta(days=DIAS_VENCIMIENTO))

        # 3) cuenta de gasto por CUIT
        codigo = mapa_codigos.get(int(cuit)) if cuit else None
        if codigo:
            ws2.cell(row=r, column=C_CUENTA_GASTO).value = codigo
        else:
            ws2.cell(row=r, column=C_CUENTA_GASTO).value = cuenta_fallback
            sin_codigo.append((r, nombre, cuit, cuenta_fallback))

        # 4) reglas de IIBB (en orden de prioridad)
        def a_no_gravado(motivo, limpiar_jurisdiccion=True):
            ws2.cell(row=r, column=C_NETO_NO_GRAVADO).value = round(
                neto_no_gravado + perc_iibb, 2)
            ws2.cell(row=r, column=C_PERC_IIBB).value = 0
            if limpiar_jurisdiccion:
                ws2.cell(row=r, column=C_JURISDICCION).value = None
            reglas.append((r, nombre, motivo, perc_iibb))

        sin_desglose = (perc_iibb != 0 and neto_no_gravado != 0
                        and round(perc_iibb, 2) == round(total, 2)
                        and round(neto_no_gravado, 2) == round(total, 2))

        if sin_desglose:
            # Factura B/C sin desglose: la plantilla duplico el total. No es IIBB real.
            ws2.cell(row=r, column=C_PERC_IIBB).value = 0
            reglas.append((r, nombre, "sin desglose (Fact. B/C) -> se anula IIBB duplicado",
                           perc_iibb))

        elif any(x in nombre_up for x in NOMBRES_ESTACION_SERVICIO):
            # La percepcion se va entera a No Gravado, asi que el IIBB queda en
            # 0: poner una jurisdiccion ahi no tiene sentido y Colppy la lee
            # como si hubiera percepcion. Va vacia, igual que en autopistas.
            a_no_gravado("estacion de servicio -> No Gravado, sin jurisdiccion")

        elif any(x in nombre_up for x in NOMBRES_AUTOPISTA):
            a_no_gravado("autopista/peaje -> No Gravado")

        elif any(x in nombre_up for x in NOMBRES_CENCOSUD):
            a_no_gravado("cencosud -> No Gravado")

        elif any(x in nombre_up for x in NOMBRES_SANCOR):
            a_no_gravado("sancor -> No Gravado, sin IIBB")

        elif any(x in nombre_up for x in NOMBRES_CEAMSE):
            a_no_gravado("CEAMSE -> No Gravado, sin IIBB")

        elif perc_iibb != 0 and abs(perc_iibb) < UMBRAL_IIBB_MENOR:
            # resto de redondeo: negativo resta, positivo suma
            ws2.cell(row=r, column=C_NETO_GRAVADO).value = round(neto_gravado + perc_iibb, 2)
            ws2.cell(row=r, column=C_PERC_IIBB).value = 0
            reglas.append((r, nombre, "IIBB < $%d -> plegado en Neto Gravado"
                           % UMBRAL_IIBB_MENOR, perc_iibb))

        elif perc_iibb != 0:
            ws2.cell(row=r, column=C_JURISDICCION).value = "Buenos Aires"
            reglas.append((r, nombre, "IIBB real -> Jurisdiccion Buenos Aires (por defecto)",
                           perc_iibb))
            verificar.append((r, nombre, perc_iibb))

        else:
            if isinstance(perc_iibb, float) and abs(perc_iibb) < 0.01:
                ws2.cell(row=r, column=C_PERC_IIBB).value = 0

    return reglas, sin_codigo, verificar


# =============================================================================
# ESCRITURA DEL .XLS CONSERVANDO EL FORMATO ORIGINAL
# =============================================================================

def _construir_estilos(libro):
    """Replica los estilos del libro original como objetos XFStyle de xlwt."""
    estilos = []
    for rdxf in libro.xf_list:
        xf = xlwt.Style.XFStyle()
        xf.num_format_str = libro.format_map[rdxf.format_key].format_str

        f = xf.font
        rf = libro.font_list[rdxf.font_index]
        f.height = rf.height
        f.italic = rf.italic
        f.struck_out = rf.struck_out
        f.outline = rf.outline
        f.shadow = rf.outline
        f.colour_index = rf.colour_index
        f.bold = rf.bold
        f._weight = rf.weight
        f.escapement = rf.escapement
        f.underline = rf.underline_type
        f.family = rf.family
        f.charset = rf.character_set
        f.name = rf.name

        p = xf.protection
        rp = rdxf.protection
        p.cell_locked = rp.cell_locked
        p.formula_hidden = rp.formula_hidden

        b = xf.borders
        rb_ = rdxf.border
        b.left = rb_.left_line_style
        b.right = rb_.right_line_style
        b.top = rb_.top_line_style
        b.bottom = rb_.bottom_line_style
        b.diag = rb_.diag_line_style
        b.left_colour = rb_.left_colour_index
        b.right_colour = rb_.right_colour_index
        b.top_colour = rb_.top_colour_index
        b.bottom_colour = rb_.bottom_colour_index
        b.diag_colour = rb_.diag_colour_index
        b.need_diag1 = rb_.diag_down
        b.need_diag2 = rb_.diag_up

        pat = xf.pattern
        bg = rdxf.background
        pat.pattern = bg.fill_pattern
        pat.pattern_fore_colour = bg.pattern_colour_index
        pat.pattern_back_colour = bg.background_colour_index

        a = xf.alignment
        ra = rdxf.alignment
        a.horz = ra.hor_align
        a.vert = ra.vert_align
        a.dire = ra.text_direction
        a.rota = ra.rotation
        a.wrap = ra.text_wrapped
        a.shri = ra.shrink_to_fit
        a.inde = ra.indent_level

        estilos.append(xf)
    return estilos


def _corregir_alturas_default(libro_orig, libro_nuevo):
    """
    NO SACAR ESTA FUNCION.

    xlutils.copy crea un registro de altura explicito (255) para filas que en el
    original no tenian ninguno, es decir que usaban la altura por defecto de la
    hoja. Eso hace que el archivo se vea distinto al original. Aca restauramos
    el flag has_default_height en todas esas filas.
    """
    for idx, nombre in enumerate(libro_orig.sheet_names()):
        hoja_orig = libro_orig.sheet_by_name(nombre)
        hoja_nueva = libro_nuevo.get_sheet(idx)
        con_altura = hoja_orig.rowinfo_map
        for rowx, fila in list(hoja_nueva.rows.items()):
            if rowx not in con_altura:
                fila.has_default_height = 1


def escribir_xls(filas_p1, ws2_final, nfilas, salida):
    """Genera el .xls final respetando exactamente el formato de la plantilla."""
    libro = xlrd.open_workbook(PLANTILLA, formatting_info=True)
    nuevo = xl_copy(libro)
    estilos = _construir_estilos(libro)
    _corregir_alturas_default(libro, nuevo)

    p1_r = libro.sheet_by_name(HOJA_P1)
    p2_r = libro.sheet_by_name(HOJA_P2)
    p1_w = nuevo.get_sheet(0)
    p2_w = nuevo.get_sheet(1)

    def escribir(hoja_r, hoja_w, fila1, col1, valor):
        r0, c0 = fila1 - 1, col1 - 1
        estilo = estilos[hoja_r.cell_xf_index(r0, c0)]
        if isinstance(valor, datetime.date) and not isinstance(valor, datetime.datetime):
            valor = datetime.datetime(valor.year, valor.month, valor.day)
        hoja_w.write(r0, c0, valor, estilo)

    # PASO 1
    for i, vals in enumerate(filas_p1):
        for c0, v in enumerate(vals):
            escribir(p1_r, p1_w, FILA_INICIO + i, c0 + 1, v)

    # PASO 2
    for i in range(nfilas):
        r = FILA_INICIO + i
        for c in range(1, COLS_P2 + 1):
            escribir(p2_r, p2_w, r, c, ws2_final.cell(row=r, column=c).value)

    # limpiar todo lo que quede debajo de la ultima factura en PASO 2
    for r0 in range(FILA_INICIO - 1 + nfilas, p2_r.nrows):
        for c0 in range(COLS_P2):
            estilo = estilos[p2_r.cell_xf_index(r0, c0)]
            p2_w.write(r0, c0, "", estilo)

    nuevo.save(salida)
    return salida


def exportar_csv(ws2, nfilas, salida):
    """Exporta PASO 2 a CSV delimitado por comas (coma decimal, fecha dd-mm-aaaa)."""
    encabezados = ["Tipo de Comprobante", "Letra", "Razon Social del Proveedor",
                   "CUIT del proveedor", "Numero de Factura", "Fecha Contable",
                   "Fecha Factura", "Fecha de Vencimiento", "Descripcion de la Compra",
                   "Total Factura", "IVA 10.5%", "IVA 21%", "IVA 27%", "Neto Gravado",
                   "Neto No Gravado", "Percepcion de IVA", "Percepcion de IIBB",
                   "Jurisdiccion de IIBB", "Cuenta Gasto", "Medio de Pago",
                   "Centro de Costo 1", "Centro de Costo 2"]
    idx_fecha = {5, 6, 7}                    # F, G, H (0-based)
    idx_num = {9, 10, 11, 12, 13, 14, 15, 16}  # J..Q

    filas = []
    for i in range(nfilas):
        r = FILA_INICIO + i
        salida_fila = []
        for c0 in range(COLS_P2):
            v = ws2.cell(row=r, column=c0 + 1).value
            if c0 in idx_fecha:
                if isinstance(v, (datetime.date, datetime.datetime)):
                    salida_fila.append(v.strftime("%d-%m-%Y"))
                else:
                    salida_fila.append(v or "")
            elif c0 in idx_num:
                if v is None or v == "":
                    salida_fila.append("")
                else:
                    salida_fila.append(("%.2f" % round(float(v), 2)).replace(".", ","))
            else:
                salida_fila.append("" if v is None else str(v))
        filas.append(salida_fila)

    with open(salida, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=",", quoting=csv.QUOTE_MINIMAL)
        w.writerow(encabezados)
        w.writerows(filas)
    return salida


def verificar_formato(salida):
    """Compara el archivo generado contra la plantilla. Devuelve lista de diferencias."""
    orig = xlrd.open_workbook(PLANTILLA, formatting_info=True)
    nuevo = xlrd.open_workbook(salida, formatting_info=True)
    problemas = []

    def altura_efectiva(hoja, r):
        ri = hoja.rowinfo_map.get(r)
        if ri is None or ri.has_default_height:
            return "DEFAULT"
        return ri.height

    if orig.sheet_names() != nuevo.sheet_names():
        problemas.append("cambiaron los nombres/orden de las hojas")

    for nombre in orig.sheet_names():
        o = orig.sheet_by_name(nombre)
        m = nuevo.sheet_by_name(nombre)
        if (o.nrows, o.ncols) != (m.nrows, m.ncols):
            problemas.append("%s: cambiaron las dimensiones" % nombre)
        if set(o.merged_cells) != set(m.merged_cells):
            problemas.append("%s: cambiaron las celdas combinadas" % nombre)
        oc, mc = o.colinfo_map, m.colinfo_map
        for c in range(max(o.ncols, m.ncols)):
            ow = oc[c].width if c in oc else None
            mw = mc[c].width if c in mc else None
            if ow != mw:
                problemas.append("%s: cambio el ancho de la columna %d" % (nombre, c))
                break
        for r in range(max(o.nrows, m.nrows)):
            if altura_efectiva(o, r) != altura_efectiva(m, r):
                problemas.append("%s: cambio la altura de la fila %d" % (nombre, r + 1))
                break
    return problemas


def verificar_totales(ws2, nfilas):
    """Controla que Total Factura = suma de sus componentes en cada fila."""
    descuadres = []
    for i in range(nfilas):
        r = FILA_INICIO + i
        total = ws2.cell(row=r, column=C_TOTAL).value or 0
        partes = 0
        for c in (11, 12, 13, C_NETO_GRAVADO, C_NETO_NO_GRAVADO, C_PERC_IVA, C_PERC_IIBB):
            v = ws2.cell(row=r, column=c).value
            partes += v if isinstance(v, (int, float)) else 0
        if abs(float(total) - partes) > 0.02:
            descuadres.append((r, ws2.cell(row=r, column=C_RAZON_SOCIAL).value,
                               total, round(partes, 2)))
    return descuadres


# =============================================================================
# PROCESO PRINCIPAL
# =============================================================================

def procesar(filas, etiqueta, dry_run=False):
    """
    Procesa una lista de comprobantes (vengan de un export de ARCA o de PDF)
    y genera la planilla de migracion.
    """
    log("=" * 78)
    log("Origen: %s" % etiqueta)
    log("=" * 78)

    if not filas:
        log("No hay comprobantes para procesar.")
        return None

    soffice = buscar_libreoffice()
    cuit_receptor, empresa = detectar_empresa(filas)
    log("Empresa detectada: %s (%s)" % (empresa["nombre"], cuit_receptor))
    log("Comprobantes a evaluar: %d" % len(filas))

    # ---- deduplicacion -----------------------------------------------------
    path_exact = os.path.join(JSON_DIR, empresa["registro_exact"])
    path_fuzzy = os.path.join(JSON_DIR, empresa["registro_fuzzy"])
    reg_exact = set(tuple(x) for x in cargar_json(path_exact, []))
    reg_fuzzy = set((c, d, round(float(t), 2)) for c, d, t in cargar_json(path_fuzzy, []))

    nuevas = []
    repetidas = 0
    for vals in filas:
        clave = clave_factura(vals)
        fecha = parse_fecha(vals[IDX_FECHA])
        total = round(float(vals[IDX_TOTAL]), 2)
        if clave in reg_exact or (clave[0], str(fecha), total) in reg_fuzzy:
            repetidas += 1
        else:
            nuevas.append(vals)

    log("Ya cargadas (se saltean): %d" % repetidas)
    log("Nuevas a procesar: %d" % len(nuevas))
    if not nuevas:
        log("\nNo hay facturas nuevas. No genero ningun archivo.")
        return None
    log()

    # ---- PASO 1: conversion de moneda --------------------------------------
    convertidas, detalle_moneda = convertir_moneda(nuevas)
    if detalle_moneda:
        log("Conversiones de moneda extranjera a pesos:")
        for nombre, moneda, tc, orig, nuevo_total in detalle_moneda:
            log("   %-45s %s %s x %s = %s" % (str(nombre)[:45], orig, moneda, tc, nuevo_total))
        log()

    tmp = tempfile.mkdtemp(prefix="migracion_")
    try:
        # volcar PASO 1 sobre una copia de la plantilla en formato xlsx
        plantilla_xlsx = convertir(soffice, PLANTILLA, tmp)
        wb = openpyxl.load_workbook(plantilla_xlsx, data_only=False)
        ws1 = wb[HOJA_P1]
        for i, vals in enumerate(convertidas):
            for c0, v in enumerate(vals):
                ws1.cell(row=FILA_INICIO + i, column=c0 + 1).value = v
        paso1_path = os.path.join(tmp, "paso1.xlsx")
        wb.save(paso1_path)

        # recalcular las formulas de PASO 2 con LibreOffice
        recalc_dir = os.path.join(tmp, "recalc")
        recalculado = convertir(soffice, paso1_path, recalc_dir,
                                filtro="xlsx:Calc MS Excel 2007 XML")

        valores = openpyxl.load_workbook(recalculado, data_only=True)[HOJA_P2]
        libro_final = openpyxl.load_workbook(paso1_path, data_only=False)
        ws2 = libro_final[HOJA_P2]

        # ---- PASO 2: reglas de negocio -------------------------------------
        extras = CUENTAS_EXTRA_MSA if empresa["nombre"] == "MEPANO" else CUENTAS_EXTRA_MPN
        mapa = cargar_codigos(empresa["hoja_codigos"], extras)
        detalles = [DETALLE_PDF.get(clave_factura(v)) for v in convertidas]
        reglas, sin_codigo, verificar = aplicar_reglas(
            ws2, valores, len(convertidas), mapa, empresa["cuenta_fallback"],
            filas_p1=convertidas, detalles=detalles)

        # ---- salidas -------------------------------------------------------
        hoy = datetime.date.today().strftime("%d-%m-%Y")
        base = "migracion factura %s %s" % (empresa["nombre"], hoy)
        os.makedirs(SALIDA_DIR, exist_ok=True)
        salida_xls = os.path.join(SALIDA_DIR, base + ".xls")

        n = 1
        while os.path.exists(salida_xls):
            n += 1
            salida_xls = os.path.join(SALIDA_DIR, "%s v%d.xls" % (base, n))

        escribir_xls(convertidas, ws2, len(convertidas), salida_xls)

        # ---- controles -----------------------------------------------------
        descuadres = verificar_totales(ws2, len(convertidas))
        problemas_formato = verificar_formato(salida_xls)

        # ---- informe -------------------------------------------------------
        if reglas:
            log("Reglas aplicadas:")
            for r, nombre, motivo, monto in reglas:
                log("   fila %-3d %-42s %-52s $%s" % (r, str(nombre)[:42], motivo, monto))
            log()

        if sin_codigo:
            log("!! Proveedores SIN codigo en CODIGOS_A_USAR (se uso el fallback):")
            for r, nombre, cuit, usado in sin_codigo:
                log("   fila %-3d %-45s CUIT %s -> %s" % (r, str(nombre)[:45], cuit, usado))
            log("   Revisalos y agregalos a CODIGOS_A_USAR.xls si corresponde otro.")
            log()

        if verificar:
            log("!! Percepciones de IIBB a CONFIRMAR (quedaron en Buenos Aires):")
            for r, nombre, monto in verificar:
                log("   fila %-3d %-45s $%s" % (r, str(nombre)[:45], monto))
            log("   Verifica si corresponde CABA o Buenos Aires.")
            log()

        if descuadres:
            log("!! ATENCION - filas donde el total no cierra:")
            for r, nombre, total, suma in descuadres:
                log("   fila %-3d %-40s total=%s suma=%s" % (r, str(nombre)[:40], total, suma))
            log()
        else:
            log("OK  Todas las filas cierran: Total = IVA + Neto Gravado + No Gravado + Percepciones")

        if problemas_formato:
            log("!! ATENCION - el formato NO coincide con el original:")
            for p in problemas_formato:
                log("   %s" % p)
            log()
        else:
            log("OK  Formato identico al original (dimensiones, anchos, alturas, combinadas)")

        # ---- actualizar registro -------------------------------------------
        if dry_run:
            log("\n(dry-run) No actualizo el registro de facturas.")
        else:
            claves_nuevas = [list(clave_factura(vals)) for vals in convertidas]
            for vals in convertidas:
                reg_exact.add(clave_factura(vals))
            guardar_json(path_exact, sorted(reg_exact))
            guardar_json(os.path.join(JSON_DIR, ULTIMA_CORRIDA), {
                "empresa": empresa["nombre"],
                "cuando": datetime.datetime.now().strftime("%d-%m-%Y %H:%M"),
                "archivo": os.path.basename(salida_xls),
                "registro": empresa["registro_exact"],
                "claves": claves_nuevas,
            })
            log("\nRegistro actualizado: %d facturas acumuladas para %s"
                % (len(reg_exact), empresa["nombre"]))

        log("\nGenerado:")
        log("   %s" % salida_xls)
        log("\n   Revisalo y, cuando este OK, guarda la hoja "
            "'PASO 2' como CSV delimitado por comas.")
        if not dry_run:
            log("\n   Si este archivo NO te sirve y lo vas a tirar, ejecuta")
            log("   \"3 - DESHACER ULTIMA CARGA.bat\" para que estas %d facturas"
                % len(convertidas))
            log("   no queden marcadas como ya cargadas.")
        return salida_xls

    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def deshacer():
    """
    Da de baja del registro las facturas de la ultima carga.
    Sirve cuando el Excel generado no sirvio y lo vas a descartar: si no,
    esas facturas quedan marcadas como cargadas y la proxima vez se saltean.
    """
    path = os.path.join(JSON_DIR, ULTIMA_CORRIDA)
    datos = cargar_json(path, None)

    if not datos or not datos.get("claves"):
        log("No hay ninguna carga para deshacer.")
        if datos and datos.get("deshecha"):
            log("La ultima ya la deshiciste el %s." % datos.get("cuando", "?"))
        return

    log("=" * 78)
    log("DESHACER LA ULTIMA CARGA")
    log("=" * 78)
    log("Empresa:  %s" % datos.get("empresa", "?"))
    log("Cuando:   %s" % datos.get("cuando", "?"))
    log("Archivo:  %s" % datos.get("archivo", "?"))
    log("Facturas: %d" % len(datos["claves"]))
    log()
    log("Se van a desmarcar esas facturas, asi las podes volver a cargar.")
    log("El Excel generado no se toca: borralo vos si no te sirve.")
    log()

    if input("Confirmas? (s/N): ").strip().lower() not in ("s", "si", "sí"):
        log("\nNo hice nada.")
        return

    path_reg = os.path.join(JSON_DIR, datos["registro"])
    registro = set(tuple(x) for x in cargar_json(path_reg, []))
    quitadas = 0
    for clave in datos["claves"]:
        if tuple(clave) in registro:
            registro.discard(tuple(clave))
            quitadas += 1
    guardar_json(path_reg, sorted(registro))

    datos["deshecha"] = True
    datos["claves"] = []
    guardar_json(path, datos)

    log("\nListo: desmarque %d factura(s). Quedan %d en el registro de %s."
        % (quitadas, len(registro), datos.get("empresa", "?")))
    log("Ya las podes volver a procesar cuando quieras.")


def main():
    ap = argparse.ArgumentParser(
        description="Genera la planilla de migracion para Colppy a partir de "
                    "'Mis Comprobantes Recibidos' (ARCA) o de facturas en PDF.")
    ap.add_argument("archivos", nargs="*",
                    help="Archivos .xlsx de ARCA y/o facturas .pdf")
    ap.add_argument("--dry-run", action="store_true",
                    help="Procesa y genera los archivos pero no actualiza el registro")
    ap.add_argument("--sin-preguntar", action="store_true",
                    help="No muestra la pantalla de revision de los PDF")
    ap.add_argument("--deshacer", action="store_true",
                    help="Da de baja del registro las facturas de la ultima carga")
    ap.add_argument("--respaldar", action="store_true",
                    help="Solo hace una copia de seguridad de los registros")
    args = ap.parse_args()

    for carpeta in (JSON_DIR, ENTRADA_ARCA, ENTRADA_PDF, SALIDA_DIR):
        os.makedirs(carpeta, exist_ok=True)

    if args.respaldar:
        respaldar_registros()
        return

    if args.deshacer:
        respaldar_registros(avisar=False)
        deshacer()
        return

    # antes de tocar nada, copia de seguridad de los registros
    respaldar_registros(avisar=False)

    if not args.archivos:
        sys.exit("No me pasaste ningun archivo para procesar.")
    if not os.path.exists(PLANTILLA):
        sys.exit("Falta la plantilla en %s" % PLANTILLA)

    hojas, pdfs = [], []
    for archivo in args.archivos:
        if not os.path.exists(archivo):
            # puede venir solo el nombre: se busca en las carpetas de entrada
            for carpeta in (ENTRADA_ARCA, ENTRADA_PDF, BASE_DIR):
                tentativa = os.path.join(carpeta, archivo)
                if os.path.exists(tentativa):
                    archivo = tentativa
                    break
            else:
                log("No encuentro el archivo: %s" % archivo)
                continue
        if archivo.lower().endswith(".pdf"):
            pdfs.append(archivo)
        else:
            hojas.append(archivo)

    # --- exports de ARCA: un archivo de salida por cada uno -------------------
    for archivo in hojas:
        try:
            _, filas = leer_comprobantes(archivo)
            actualizar_proveedores(filas)
            procesar(filas, os.path.basename(archivo), dry_run=args.dry_run)
        except SystemExit:
            raise
        except Exception as e:
            log("\nERROR procesando %s:\n   %s" % (archivo, e))
            import traceback
            traceback.print_exc()
        log()

    # --- PDF: todos juntos en un mismo archivo de salida ---------------------
    if pdfs:
        try:
            filas = leer_pdfs(pdfs, interactivo=not args.sin_preguntar)
            if filas:
                etiqueta = ("%d factura(s) en PDF" % len(filas)) if len(filas) > 1 \
                    else os.path.basename(pdfs[0])
                procesar(filas, etiqueta, dry_run=args.dry_run)
            else:
                log("No quedo ninguna factura para cargar.")
        except SystemExit:
            raise
        except Exception as e:
            log("\nERROR procesando los PDF:\n   %s" % e)
            import traceback
            traceback.print_exc()
        log()


if __name__ == "__main__":
    main()
