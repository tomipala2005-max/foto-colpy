# -*- coding: utf-8 -*-
"""
leer_pdf.py - Lector de facturas en PDF para el circuito de migracion a Colppy.

Saca los datos de una factura electronica argentina en PDF y arma la misma fila
que traeria el export "Mis Comprobantes Recibidos" de ARCA, para que el resto
del circuito (duplicados, moneda, reglas de IIBB, cuenta de gasto, escritura del
.xls) funcione igual que siempre.

DE DONDE SALE CADA DATO
-----------------------
1) CODIGO QR DE AFIP  (fuente principal, es la mas confiable)
   Todo comprobante electronico lleva un QR con los datos firmados:
   fecha, CUIT del emisor, punto de venta, tipo y numero de comprobante,
   importe total, moneda, cotizacion, CUIT del receptor y el CAE.
   Eso se lee y se usa tal cual: no hay margen de error de lectura.

2) TEXTO DEL PDF  (para lo que el QR no trae)
   El desglose impositivo (neto gravado, IVA por alicuota, no gravado,
   exento, percepciones) y la razon social del proveedor.

3) REGISTRO DE PROVEEDORES  (json arch/_proveedores.json)
   Si el CUIT ya aparecio en algun export de ARCA, se usa el nombre exacto
   que usa AFIP, para que quede igual que en las cargas anteriores.

CONTROL
-------
Siempre se verifica que
    neto gravado + IVA + no gravado + exento + otros tributos = total del QR
Si no cuadra, se avisa y se pide revisar a mano. Nunca se inventa un numero.

Si el PDF es un escaneo (una foto, sin texto adentro), se avisa y se piden los
datos a mano: no se adivina nada.
"""

import io
import os
import re
import sys
import json
import base64
import datetime
import time
import unicodedata


# =============================================================================
# DEPENDENCIAS (todas opcionales salvo pypdf/pdfplumber)
# =============================================================================

try:
    from pypdf import PdfReader
except ImportError:
    try:
        from PyPDF2 import PdfReader
    except ImportError:
        PdfReader = None

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

try:
    import cv2
    import numpy as np
except ImportError:
    cv2 = None
    np = None

try:
    import pypdfium2
except ImportError:
    pypdfium2 = None

# pyzbar lee codigos QR mucho mejor que opencv cuando la imagen es una foto
# o un escaneo torcido. Si no esta, se usa solo opencv.
try:
    from pyzbar import pyzbar
except Exception:
    pyzbar = None

# OCR opcional, para PDF que son una foto sin texto adentro.
# Necesita ademas el programa Tesseract instalado en la maquina.
try:
    import pytesseract
except Exception:
    pytesseract = None


def dependencias_faltantes():
    faltan = []
    if PdfReader is None:
        faltan.append("pypdf")
    if pdfplumber is None:
        faltan.append("pdfplumber")
    return faltan


# Donde suele quedar Tesseract en Windows. El instalador NO lo agrega al PATH
# salvo que tildes la casilla, asi que hay que ir a buscarlo.
RUTAS_TESSERACT = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    os.path.join(os.environ.get("LOCALAPPDATA", ""),
                 "Tesseract-OCR", "tesseract.exe"),
    os.path.join(os.environ.get("LOCALAPPDATA", ""),
                 "Programs", "Tesseract-OCR", "tesseract.exe"),
    os.path.join(os.environ.get("PROGRAMFILES", ""),
                 "Tesseract-OCR", "tesseract.exe"),
    "/usr/bin/tesseract",
    "/usr/local/bin/tesseract",
]

_ESTADO_OCR = None      # se calcula una sola vez por corrida
IDIOMA_OCR = "spa"


def _buscar_tesseract():
    """Busca el programa Tesseract: primero en el PATH, despues a mano."""
    import shutil
    ruta = shutil.which("tesseract")
    if ruta:
        return ruta
    for candidata in RUTAS_TESSERACT:
        if candidata and os.path.exists(candidata):
            return candidata
    return None


def estado_ocr():
    """
    Devuelve (disponible, motivo). El motivo explica en castellano que falta,
    asi no se queda callado cuando no puede hacer OCR.
    """
    global _ESTADO_OCR, IDIOMA_OCR
    if _ESTADO_OCR is not None:
        return _ESTADO_OCR

    if pytesseract is None:
        _ESTADO_OCR = (False, "falta la libreria pytesseract "
                              "(pip install pytesseract)")
        return _ESTADO_OCR
    if pypdfium2 is None:
        _ESTADO_OCR = (False, "falta la libreria pypdfium2 "
                              "(pip install pypdfium2)")
        return _ESTADO_OCR

    ruta = _buscar_tesseract()
    if not ruta:
        _ESTADO_OCR = (False, "no encuentro tesseract.exe. Instalalo desde "
                              "https://github.com/UB-Mannheim/tesseract/wiki "
                              "o agrega su carpeta a RUTAS_TESSERACT en "
                              "leer_pdf.py")
        return _ESTADO_OCR

    try:
        pytesseract.pytesseract.tesseract_cmd = ruta
        pytesseract.get_tesseract_version()
    except Exception as e:
        _ESTADO_OCR = (False, "encontre tesseract en %s pero no arranca (%s)"
                              % (ruta, e))
        return _ESTADO_OCR

    # si no tiene el paquete de espanol, se usa el idioma que haya
    try:
        idiomas = pytesseract.get_languages(config="")
        if IDIOMA_OCR not in idiomas:
            IDIOMA_OCR = "eng" if "eng" in idiomas else (idiomas[0] if idiomas
                                                         else "eng")
    except Exception:
        pass

    _ESTADO_OCR = (True, "OK (%s, idioma %s)" % (ruta, IDIOMA_OCR))
    return _ESTADO_OCR


def hay_ocr():
    """True si se puede hacer OCR (pytesseract + el programa Tesseract)."""
    return estado_ocr()[0]


# =============================================================================
# TABLAS
# =============================================================================

# Codigo de comprobante AFIP -> etiqueta igual a la que usa "Mis Comprobantes"
TIPOS_COMPROBANTE = {
    1: "Factura A",            2: "Nota de Débito A",     3: "Nota de Crédito A",
    4: "Recibo A",             6: "Factura B",            7: "Nota de Débito B",
    8: "Nota de Crédito B",    9: "Recibo B",            11: "Factura C",
    12: "Nota de Débito C",   13: "Nota de Crédito C",   15: "Recibo C",
    19: "Factura E",          20: "Nota de Débito E",    21: "Nota de Crédito E",
    51: "Factura M",          52: "Nota de Débito M",    53: "Nota de Crédito M",
    54: "Recibo M",
    201: "Factura de Crédito electrónica MiPyMEs (FCE) A",
    202: "Nota de Débito electrónica MiPyMEs (FCE) A",
    203: "Nota de Crédito electrónica MiPyMEs (FCE) A",
    206: "Factura de Crédito electrónica MiPyMEs (FCE) B",
    207: "Nota de Débito electrónica MiPyMEs (FCE) B",
    208: "Nota de Crédito electrónica MiPyMEs (FCE) B",
    211: "Factura de Crédito electrónica MiPyMEs (FCE) C",
    212: "Nota de Débito electrónica MiPyMEs (FCE) C",
    213: "Nota de Crédito electrónica MiPyMEs (FCE) C",
}

# Comprobantes que restan (notas de credito)
TIPOS_CREDITO = {3, 8, 13, 21, 53, 203, 208, 213}

MONEDAS = {"PES": "$", "DOL": "USD", "060": "EUR", "012": "BRL"}


# =============================================================================
# NUMEROS
# =============================================================================

def parse_num(txt):
    """
    Convierte un numero escrito a mano en float, aguantando los dos formatos
    que aparecen en las facturas argentinas:
        1.234.567,89   (formato local)
        1,234,567.89   (formato ingles, muy comun en sistemas de facturacion)
    Devuelve None si no es un numero.
    """
    if txt is None:
        return None
    s = str(txt).strip()
    s = s.replace("$", "").replace(" ", "").replace("\xa0", "")
    if not s:
        return None
    negativo = s.startswith("-") or (s.startswith("(") and s.endswith(")"))
    s = s.strip("()-+")
    if not s or not re.match(r"^[\d.,]+$", s):
        return None

    hay_punto = "." in s
    hay_coma = "," in s

    if hay_punto and hay_coma:
        # el separador que aparece MAS A LA DERECHA es el decimal
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif hay_coma:
        partes = s.split(",")
        if len(partes) == 2 and len(partes[1]) in (1, 2):
            s = s.replace(",", ".")            # 1234,5  ->  decimal
        elif all(len(p) == 3 for p in partes[1:]):
            s = s.replace(",", "")             # 1,234,567  ->  miles
        else:
            s = s.replace(",", ".")
    elif hay_punto:
        partes = s.split(".")
        if len(partes) == 2 and len(partes[1]) in (1, 2):
            pass                               # 1234.5  ->  decimal
        elif all(len(p) == 3 for p in partes[1:]):
            s = s.replace(".", "")             # 1.234.567  ->  miles

    try:
        v = float(s)
    except ValueError:
        return None
    return -v if negativo else v


RE_NUMERO = re.compile(r"-?\(?\$?\s?\d[\d.,]*\)?")


def limpiar_texto_ocr(texto):
    """
    Emparcha los destrozos tipicos del OCR sobre numeros.
    El mas comun: el espacio que mete despues de la coma decimal, que parte
    "16.504,31" en dos numeros sueltos ("16504," y "31").
    """
    if not texto:
        return texto
    t = texto
    t = re.sub(r"([.,])\s+(\d{2})(?!\d)", r"\1\2", t)   # "16504, 31" -> "16504,31"
    t = re.sub(r",\s*\.", ".", t)                        # "103773,.01" -> "103773.01"
    t = re.sub(r"(\d)\s+([.,])(\d{2})(?!\d)", r"\1\2\3", t)  # "1234 ,56"

    # Algunos sistemas imprimen los importes con un espacio entre cada digito:
    #     "2 7 2 , 7 1 0 . 7 4"  ->  "272,710.74"
    # Se reconoce porque son caracteres sueltos separados por un solo espacio.
    # Ojo de no pegar dos importes distintos ("1.515,84 1.515,84"): ahi los
    # pedazos no son de un caracter, asi que no entra.
    t = re.sub(r"(?<![\d.,])((?:[\d.,] ){3,}[\d.,])(?![\d.,])",
               lambda m: m.group(1).replace(" ", ""), t)
    return t


def numeros_de_linea(linea):
    """Todos los numeros de una linea, de izquierda a derecha."""
    out = []
    for m in RE_NUMERO.finditer(linea):
        v = parse_num(m.group())
        if v is not None:
            out.append(v)
    return out


RE_PORCENTAJE = re.compile(r"\d+(?:[.,]\d+)?\s*%")


def sin_porcentajes(linea):
    """La linea sin las alicuotas, que no son importes sino etiquetas."""
    return RE_PORCENTAJE.sub(" ", linea or "")


def importes_de_linea(linea):
    """Los numeros de la linea que pueden ser importes (sin los porcentajes)."""
    return numeros_de_linea(sin_porcentajes(linea))


def ultimo_numero(linea):
    nums = numeros_de_linea(linea)
    return nums[-1] if nums else None


def numero_de_etiqueta(linea, fin_etiqueta):
    """
    El importe de una etiqueta es el que viene JUSTO DESPUES, no el ultimo de
    la linea. Importa cuando el proveedor pone la banda de totales toda junta:

        192.104,61     IVA 21%     40.341,97     232.446,58
                                   ^ este

    Hay que saltear los restos de la propia alicuota. En "ALICUOTA 21,00 %
    15.145,81" la etiqueta termina en el "21", y lo que sigue es ",00" -> el
    "00" no es un importe, es la parte decimal del 21. Se reconoce porque
    viene pegado a un separador decimal, o porque lo sigue un signo %.
    """
    resto = linea[fin_etiqueta:]
    for m in RE_NUMERO.finditer(resto):
        # pegado a una coma o punto: es la continuacion de la alicuota
        anterior = resto[m.start() - 1] if m.start() > 0 else ""
        if anterior in ".," or m.group()[:1] in ".,":
            continue
        # seguido de % : tambien es la alicuota
        if "%" in resto[m.end():m.end() + 3]:
            continue
        valor = parse_num(m.group())
        if valor is not None:
            return valor
    # como respaldo, el ultimo de la linea, pero nunca una alicuota
    return ultimo_numero(sin_porcentajes(linea))


def numero_en_columna(lineas, indice, inicio, fin, max_filas=4):
    """
    Muchas facturas ponen los titulos en una fila y los importes en la de
    abajo, alineados por columna:

        Bruto         Descuento        Iva 21%      Importe Total
        11564.82       3469.45         1700.03         9795.40

    Dado el lugar que ocupa el titulo en su linea, busca el numero que cae
    justo debajo. Solo mira filas que tengan varios numeros, para no agarrar
    cualquier cifra suelta de mas abajo.
    """
    centro = (inicio + fin) / 2.0
    for j in range(indice + 1, min(indice + 1 + max_filas, len(lineas))):
        linea = lineas[j]
        if not linea.strip():
            continue
        posiciones = []
        for m in RE_NUMERO.finditer(linea):
            if "%" in linea[m.end():m.end() + 3]:
                continue                      # es una alicuota, no un importe
            valor = parse_num(m.group())
            if valor is not None:
                posiciones.append((m.start(), m.end(), valor))
        if len(posiciones) < 2:
            continue            # no parece una fila de importes
        mejor = None
        for ini, fin_n, valor in posiciones:
            solapa = not (fin_n < inicio or ini > fin)
            distancia = abs((ini + fin_n) / 2.0 - centro)
            if solapa and (mejor is None or distancia < mejor[0]):
                mejor = (distancia, valor)
        if mejor:
            return mejor[1]
        return None             # habia fila de importes pero ninguno alineado
    return None


def normalizar(txt):
    """Mayusculas y sin acentos, para comparar etiquetas sin sorpresas."""
    if txt is None:
        return ""
    t = unicodedata.normalize("NFD", str(txt))
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return t.upper()


# =============================================================================
# 1) CODIGO QR DE AFIP
# =============================================================================

def _qr_con_pyzbar(im):
    """pyzbar: el mas tolerante con fotos, escaneos torcidos y sombras."""
    if pyzbar is None:
        return None
    try:
        for codigo in pyzbar.decode(im):
            dato = codigo.data.decode("utf-8", "replace")
            if "p=" in dato:
                return dato
    except Exception:
        pass
    return None


def _qr_con_opencv(im):
    """opencv: siempre esta disponible, pero necesita una imagen mas limpia."""
    if cv2 is None:
        return None
    detector = cv2.QRCodeDetector()
    # el QR necesita margen blanco alrededor para que lo reconozca
    im = cv2.copyMakeBorder(im, 40, 40, 40, 40, cv2.BORDER_CONSTANT, value=255)
    for intento in ("plano", "multi", "curvo"):
        try:
            if intento == "plano":
                texto, _, _ = detector.detectAndDecode(im)
            elif intento == "multi":
                ok, textos, _, _ = detector.detectAndDecodeMulti(im)
                texto = next((t for t in (textos or []) if t and "p=" in t), "") \
                    if ok else ""
            else:
                texto, _, _ = detector.detectAndDecodeCurved(im)
        except Exception:
            continue
        if texto and "p=" in texto:
            return texto
    return None


def _variantes(img):
    """Distintas versiones de la misma imagen, a ver cual se deja leer."""
    yield img
    if cv2 is None:
        return
    try:
        _, otsu = cv2.threshold(img, 0, 255,
                                cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        yield otsu
    except Exception:
        pass
    try:
        # sirve cuando el escaneo tiene sombras o esta desparejo
        yield cv2.adaptiveThreshold(img, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                    cv2.THRESH_BINARY, 41, 12)
    except Exception:
        pass
    try:
        nitido = cv2.filter2D(img, -1, np.array([[0, -1, 0],
                                                 [-1, 5, -1],
                                                 [0, -1, 0]]))
        yield nitido
    except Exception:
        pass


# Segundos como mucho buscando el QR. Los tiques de controlador fiscal
# (estaciones de servicio y similares) directamente no tienen QR: sin este
# tope se perdian 50 segundos por comprobante buscando algo que no esta.
LIMITE_QR_SEGUNDOS = 20


def _buscar_en_imagen(img, con_recortes=False, vence=None):
    """
    Busca el QR en una imagen. Primero entera; si no aparece y se pide,
    la va partiendo en pedazos (en una factura escaneada el QR es chiquito
    y se pierde dentro de la hoja entera).
    """
    if img is None or cv2 is None:
        return None

    for version in _variantes(img):
        if vence and time.time() > vence:
            return None
        texto = _qr_con_pyzbar(version) or _qr_con_opencv(version)
        if texto:
            return texto

    if not con_recortes:
        return None

    alto, ancho = img.shape[:2]
    # ventanas de media hoja que se van pisando entre si, asi el QR nunca
    # queda cortado justo en el borde de un recorte
    paso_y, paso_x = max(alto // 4, 1), max(ancho // 4, 1)
    ventana_y, ventana_x = alto // 2, ancho // 2
    for y in range(0, max(alto - ventana_y, 0) + 1, paso_y):
        for x in range(0, max(ancho - ventana_x, 0) + 1, paso_x):
            if vence and time.time() > vence:
                return None
            recorte = img[y:y + ventana_y, x:x + ventana_x]
            if recorte.size == 0:
                continue
            try:
                grande = cv2.resize(recorte, None, fx=2, fy=2,
                                    interpolation=cv2.INTER_CUBIC)
            except Exception:
                grande = recorte
            for version in _variantes(grande):
                texto = _qr_con_pyzbar(version) or _qr_con_opencv(version)
                if texto:
                    return texto
    return None


def _qr_a_dict(texto):
    if not texto or "p=" not in texto:
        return None
    try:
        carga = texto.split("p=", 1)[1].split("&")[0]
        carga += "=" * (-len(carga) % 4)
        return json.loads(base64.b64decode(carga).decode("utf-8", "replace"))
    except Exception:
        return None


def _paginas_dibujadas(path, escala, paginas=3):
    """Dibuja las primeras paginas del PDF como imagen en escala de grises."""
    if pypdfium2 is None or np is None:
        return
    try:
        doc = pypdfium2.PdfDocument(path)
    except Exception:
        return
    try:
        for i in range(min(len(doc), paginas)):
            try:
                pil = doc[i].render(scale=escala).to_pil().convert("L")
                yield np.array(pil)
            except Exception:
                continue
    finally:
        try:
            doc.close()
        except Exception:
            pass


def _imagenes_incrustadas(path):
    """Las imagenes que ya vienen adentro del PDF, sin redibujar nada."""
    if PdfReader is None or cv2 is None or np is None:
        return
    try:
        lector = PdfReader(path)
    except Exception:
        return
    for pagina in lector.pages[:3]:
        try:
            imagenes = list(pagina.images)
        except Exception:
            continue
        for img in imagenes:
            try:
                arr = cv2.imdecode(np.frombuffer(img.data, np.uint8),
                                   cv2.IMREAD_GRAYSCALE)
            except Exception:
                continue
            if arr is not None:
                yield arr


def leer_qr_afip(path):
    """
    Busca el QR de AFIP. Devuelve el dict de datos o None.

    El orden importa muchisimo. pyzbar encuentra el QR en milisegundos cuando
    esta a la vista, asi que se prueba PRIMERO en todos lados. Recien despues
    se pasa a lo caro (los filtros de opencv y el recorte de la hoja en
    pedazos), que es lo que tarda segundos.
    """
    if cv2 is None or np is None:
        return None

    vence = time.time() + LIMITE_QR_SEGUNDOS

    def listo(texto):
        return _qr_a_dict(texto)

    # se dibuja cada pagina UNA sola vez y se reusa
    paginas = []

    def dibujadas(escala):
        for i, (esc, arr) in enumerate(paginas):
            if esc == escala:
                yield arr
        if not any(esc == escala for esc, _ in paginas):
            for arr in _paginas_dibujadas(path, escala, paginas=2):
                paginas.append((escala, arr))
                yield arr

    # --- 1) pyzbar, que es rapido, sobre todo lo que haya ------------------
    for arr in _imagenes_incrustadas(path):
        datos = listo(_qr_con_pyzbar(arr))
        if datos:
            return datos
    for escala in (3, 5):
        if time.time() > vence:
            return None
        for arr in dibujadas(escala):
            datos = listo(_qr_con_pyzbar(arr))
            if datos:
                return datos

    # --- 2) opencv con filtros, mas lento ----------------------------------
    for arr in _imagenes_incrustadas(path):
        if time.time() > vence:
            return None
        datos = listo(_buscar_en_imagen(arr, vence=vence))
        if datos:
            return datos
    for escala in (3, 5):
        if time.time() > vence:
            return None
        for arr in dibujadas(escala):
            datos = listo(_buscar_en_imagen(arr, vence=vence))
            if datos:
                return datos

    # --- 3) recortando la hoja en pedazos, el ultimo recurso ---------------
    for escala in (3, 5):
        if time.time() > vence:
            return None
        for arr in dibujadas(escala):
            datos = listo(_buscar_en_imagen(arr, con_recortes=True, vence=vence))
            if datos:
                return datos
    return None


# =============================================================================
# 2) TEXTO DEL PDF
# =============================================================================

def leer_texto(path):
    """Texto del PDF respetando la disposicion de las columnas."""
    if pdfplumber is None:
        return ""
    partes = []
    try:
        with pdfplumber.open(path) as pdf:
            for pagina in pdf.pages:
                try:
                    t = pagina.extract_text(layout=True)
                except Exception:
                    t = pagina.extract_text()
                if t:
                    partes.append(t)
    except Exception:
        return ""
    return "\n".join(partes)


def es_escaneo(texto):
    """True si el PDF no tiene texto adentro (es una foto)."""
    return len(re.sub(r"\s", "", texto or "")) < 80


def _imagen_mas_grande(path):
    """
    La mejor imagen disponible de la primera pagina.
    Importante: la imagen incrustada en el PDF suele tener MAS resolucion que
    la pagina dibujada. Hacer OCR sobre una version reducida pierde datos que
    estaban ahi (los CUIT, sobre todo).
    """
    candidatas = []
    if PdfReader is not None and cv2 is not None and np is not None:
        try:
            for pagina in PdfReader(path).pages[:1]:
                for img in list(pagina.images):
                    arr = cv2.imdecode(np.frombuffer(img.data, np.uint8),
                                       cv2.IMREAD_GRAYSCALE)
                    if arr is not None:
                        candidatas.append(arr)
        except Exception:
            pass
    for arr in _paginas_dibujadas(path, 3, paginas=1):
        candidatas.append(arr)
    if not candidatas:
        return None
    return max(candidatas, key=lambda a: a.shape[0] * a.shape[1])


def _variantes_ocr(img):
    """
    Versiones de la imagen para pasarle al OCR, de la que mas suele acertar a
    la que menos. Se corta apenas una da un resultado que cuadra.
    """
    if cv2 is None:
        return
    # ~300 dpi es donde mejor lee Tesseract
    escala = 2 if min(img.shape[:2]) < 2200 else 1
    if escala > 1:
        img = cv2.resize(img, None, fx=escala, fy=escala,
                         interpolation=cv2.INTER_CUBIC)

    _, otsu = cv2.threshold(img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    yield "otsu", otsu, ""
    yield "plano", img, ""
    yield "otsu psm6", otsu, "--psm 6"
    try:
        adap = cv2.adaptiveThreshold(img, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                     cv2.THRESH_BINARY, 51, 15)
        yield "adaptativo", adap, ""
        yield "adaptativo psm6", adap, "--psm 6"
    except Exception:
        pass
    yield "plano psm6", img, "--psm 6"


def _ocr(img, cfg):
    try:
        return pytesseract.image_to_string(img, lang=IDIOMA_OCR, config=cfg)
    except Exception:
        try:
            return pytesseract.image_to_string(img, config=cfg)
        except Exception:
            return ""


def leer_texto_ocr(path):
    """
    OCR para PDF escaneados, con varias pasadas.

    El truco: de cada pasada se sacan los importes y se comprueba si suman el
    total del comprobante. La que cuadra es, casi con certeza, la que leyo bien.
    Los CUIT se juntan de TODAS las pasadas y se filtran por digito verificador,
    asi que alcanza con que una sola los haya leido bien.

    Devuelve (texto, cuits_candidatos).
    """
    disponible, motivo = estado_ocr()
    if not disponible:
        print("  [!] No puedo hacer OCR: %s" % motivo)
        return "", []

    img = _imagen_mas_grande(path)
    if img is None:
        return "", []

    print("  Pasando el escaneo por OCR, esto tarda un rato...")
    mejor, mejor_puntaje = "", -1
    cuits = []
    for nombre, version, cfg in _variantes_ocr(img):
        texto = _ocr(version, cfg)
        if not texto:
            continue

        normalizado = normalizar(texto)
        for c in cuits_del_texto(normalizado):
            if c not in cuits:
                cuits.append(c)
        # el CUIT propio, aunque el OCR le haya errado un digito
        propio = _empresa_mas_parecida(normalizado)
        if propio and propio not in cuits:
            cuits.append(propio)

        imp = extraer_importes(texto)
        total = imp.get("total")
        imp, _ = reparar_por_cuadre(imp, total)
        cuadra = total is not None and abs(_suma(imp) - total) <= 0.05

        # se prefiere la pasada que cuadra; si ninguna, la que mas texto saco
        puntaje = (1000000 + len(texto)) if cuadra else len(texto)
        if puntaje > mejor_puntaje:
            mejor, mejor_puntaje = texto, puntaje

        if cuadra and cuits:
            print("  OCR: los importes cuadran (pasada '%s')." % nombre)
            break

    return mejor, cuits


# =============================================================================
# 3) DESGLOSE IMPOSITIVO DESDE EL TEXTO
# =============================================================================

# Se evaluan EN ESTE ORDEN. El primero que matchea se queda con la linea.
# Ojo: "NO GRAVADO" tiene que ir antes que "GRAVADO".
REGLAS_IMPORTES = [
    ("no_gravado", [r"\bNO\s+GRAVADO", r"IMPORTES?\s+NO\s+GRAVADOS?",
                    r"CONCEPTOS?\s+NO\s+GRAVADOS?", r"\bNO\s+GRAVADAS?\b"]),
    ("exento",     [r"\bEXENTO", r"\bEXENTAS?\b", r"OP\.?\s*EXENTAS"]),
    # "ALICUOTA 21,00%" es como lo escriben los controladores fiscales
    # (tiques de estacion de servicio y similares), sin la palabra IVA.
    ("iva_105",    [r"I\.?V\.?A\.?.{0,12}\b10[.,]5",
                    r"AL[IÍ1]CUOTA.{0,10}\b10[.,]5"]),
    ("iva_27",     [r"I\.?V\.?A\.?.{0,12}\b27\b",
                    r"AL[IÍ1]CUOTA.{0,10}\b27\b"]),
    ("iva_21",     [r"I\.?V\.?A\.?.{0,12}\b21\b",
                    r"AL[IÍ1]CUOTA.{0,10}\b21\b"]),
    ("iva_total",  [r"^\s*I\.?V\.?A\.?\b", r"TOTAL\s+I\.?V\.?A", r"SUMA\s+I\.?V\.?A",
                    r"TOTAL\s+IMPUESTOS?"]),
    ("percepcion_iibb", [r"PERCEP.{0,25}(IIBB|ING.{0,3}BRUTO)",
                         r"(IIBB|ING.{0,3}BRUTO).{0,25}PERCEP",
                         r"PERC\.?\s*(IIBB|I\.?B\.?)\b"]),
    ("percepcion_iva",  [r"PERCEP.{0,20}I\.?V\.?A"]),
    ("otros_tributos",  [r"OTROS\s+TRIBUTOS", r"IMP\.?\s*INTERNOS",
                         r"IMPUESTOS?\s+INTERNOS", r"\bRETENC",
                         r"\bPERCEP", r"IMP\.?\s*MUNICIPAL"]),
    # [BG6] porque el OCR confunde la G de GRAVADO con B o 6
    ("neto_gravado",    [r"IMPORTES?\s+[BG6]RAVADOS?", r"NETO\s+[BG6]RAVADO",
                         r"IMPORTE\s+NETO", r"SUBTOTAL\s+[BG6]RAVADO",
                         r"SUBTOT.{0,25}[BG6]RAVADO",
                         r"\b[BG6]RAVADO\b", r"^\s*NETO\b"]),
    ("subtotal",   [r"SUBTOTAL"]),
    # el (?!...) evita confundir "Total Impuestos" o "Total IVA" con el total
    ("total",      [r"^\s*TOTAL\b(?!\s*(IMPUESTOS?|I\.?V\.?A|DISCRIMINADO|BRUTO))",
                    r"TOTAL\s+CON\s+IMPUESTOS",
                    r"TOTAL\s+(A\s+PAGAR|FACTURA|GENERAL|NETO)",
                    r"IMPORTE\s+TOTAL", r"TOTAL\s*\$"]),
]


def extraer_importes(texto):
    """Recorre el texto y junta los importes por etiqueta."""
    hallados = {}
    lineas = limpiar_texto_ocr(texto or "").splitlines()
    for i, linea in enumerate(lineas):
        limpia = normalizar(linea)
        if not limpia.strip():
            continue
        # Una linea SIN numeros es una fila de titulos: ahi conviven varias
        # etiquetas ("Bruto  Descuento  Iva 21%  Importe Total") y hay que
        # buscarle a cada una su numero en la fila de abajo. Una linea CON
        # numeros, en cambio, es de una sola etiqueta: se corta en la primera
        # que coincide, para que "IMPORTE TOTAL OTROS TRIBUTOS 16.504,31" no
        # se lea tambien como el total de la factura.
        fila_de_titulos = not importes_de_linea(linea)

        for clave, patrones in REGLAS_IMPORTES:
            if clave in hallados:
                continue
            encontrado = None
            for patron in patrones:
                encontrado = re.search(patron, limpia)
                if encontrado:
                    break
            if not encontrado:
                continue

            if fila_de_titulos:
                valor = numero_en_columna(lineas, i,
                                          encontrado.start(), encontrado.end())
                if valor is not None:
                    hallados[clave] = valor
                continue

            valor = numero_de_etiqueta(linea, encontrado.end())
            if valor is not None:
                hallados[clave] = valor
            break
    return hallados


COMPONENTES = ("neto_gravado", "iva_105", "iva_21", "iva_27",
               "no_gravado", "exento", "otros_tributos")


def _suma(imp):
    return round(sum(imp.get(k) or 0 for k in COMPONENTES), 2)


def reparar_por_cuadre(imp, total, tolerancia=0.05):
    """
    Cuando el OCR se come la coma decimal, un importe de 72.122,89 se lee como
    7212289. Se prueba a dividir por 100 los importes sospechosos (enteros
    grandes) y se acepta la combinacion que hace que las partes sumen EXACTO
    el total del comprobante.

    La verificacion es el propio total: si no da exacto, no se toca nada.
    Devuelve (importes, se_reparo).
    """
    if not total:
        return imp, False
    if abs(_suma(imp) - total) <= tolerancia:
        return imp, False

    # candidatos: enteros de 4 cifras o mas, que es como queda un importe
    # al que el OCR le comio la coma
    sospechosos = [k for k in COMPONENTES
                   if imp.get(k) and float(imp[k]).is_integer()
                   and abs(imp[k]) >= 1000]
    if not sospechosos or len(sospechosos) > 8:
        return imp, False

    mejor = None
    for mascara in range(1, 2 ** len(sospechosos)):
        prueba = dict(imp)
        for i, clave in enumerate(sospechosos):
            if mascara >> i & 1:
                prueba[clave] = round(prueba[clave] / 100.0, 2)
        diferencia = abs(_suma(prueba) - total)
        if diferencia <= tolerancia:
            # ante empate, la que toca menos importes
            peso = bin(mascara).count("1")
            if mejor is None or peso < mejor[0]:
                mejor = (peso, prueba)
    if mejor:
        return mejor[1], True
    return imp, False


def descartar_imposibles(imp, total, tolerancia=0.05):
    """
    Saca los importes que no pueden ser: ninguna parte puede ser mayor que el
    total del comprobante. Cuando una lo es, lo que se leyo no era ese importe
    sino otro numero de la hoja. Mejor sacarlo y deducirlo despues que dejar
    un dato inventado.
    """
    if not total:
        return imp, []
    limite = abs(total) + tolerancia
    fuera = []
    nuevo = dict(imp)
    for clave in COMPONENTES + ("percepcion_iibb", "percepcion_iva"):
        valor = nuevo.get(clave)
        if valor and abs(valor) > limite:
            fuera.append(clave)
            nuevo[clave] = 0.0
    return nuevo, fuera


def descartar_migajas(imp, total, tolerancia=0.05):
    """
    Saca las migajas: importes ridiculamente chicos frente al total, que en
    layouts enredados salen de agarrar un numero cualquiera de la hoja.
    Solo se hace cuando las partes NO cuadran (si cuadran, no se toca nada).
    Es preferible un dato faltante, que se pregunta, a uno inventado.
    """
    if not total or abs(_suma(imp) - total) <= tolerancia:
        return imp, []
    umbral = abs(total) * 0.001
    fuera, nuevo = [], dict(imp)
    for clave in COMPONENTES + ("iva_total", "percepcion_iibb", "percepcion_iva"):
        valor = nuevo.get(clave)
        if valor and abs(valor) < umbral:
            fuera.append(clave)
            nuevo[clave] = 0.0
    return nuevo, fuera


def _total_sensato(imp):
    """
    El total leido del texto, si tiene sentido.

    Se descarta cuando es menor que alguna de sus propias partes: ahi lo que
    se leyo no era el total sino cualquier otro numero suelto de la hoja (un
    anio, un numero de cliente, un codigo). Si no hay total creible pero si
    partes, se usa la suma.
    """
    total = imp.get("total")
    partes = [abs(imp.get(k) or 0) for k in COMPONENTES]
    mayor = max(partes) if partes else 0
    suma = _suma(imp)

    if total is not None and abs(total) + 0.05 >= mayor:
        return total
    # El total leido no puede ser menor que una de sus partes: si pasa eso,
    # lo que se leyo no era el total sino cualquier otro numero de la hoja.
    # Se usa la suma solo si de verdad se entendio el desglose (dos partes o
    # mas); si no, es mas honesto quedarse sin total y preguntar.
    if suma and sum(1 for k in COMPONENTES if imp.get(k)) >= 2:
        return suma
    return None


# Debajo de esto, un importe suelto no se toma en serio: en una factura de
# miles de pesos, un "no gravado" de $40 casi siempre es un numero cualquiera
# que quedo pegado, no un concepto real. Subilo o bajalo si te queda corto.
IMPORTE_MINIMO_CREIBLE = 50.0

CAMPOS_REVISABLES = COMPONENTES + ("percepcion_iibb", "percepcion_iva")

NOMBRE_CAMPO = {
    "neto_gravado": "neto gravado", "iva_105": "IVA 10,5%",
    "iva_21": "IVA 21%", "iva_27": "IVA 27%", "no_gravado": "no gravado",
    "exento": "exento", "otros_tributos": "otros tributos",
    "percepcion_iibb": "percepcion de IIBB", "percepcion_iva": "percepcion de IVA",
}


def revisar_coherencia(imp, total, es_credito=False):
    """
    Ultimo control antes de dar por buena la lectura.

    Un importe puede ser consistente y aun asi no tener sentido. Antes de
    escribir un numero raro, se lo saca y se marca la factura para que la mires:
    es preferible completar un campo a mano que corregir despues una carga.

    Se marcan:
      - importes ridiculos para el tamanio de la factura (los $40 de no gravado)
      - importes negativos en una factura comun (solo valen en notas de credito)
      - un IVA que no guarda relacion con ninguna alicuota real

    Devuelve (importes, lista de avisos).
    """
    problemas = []
    nuevo = dict(imp)
    if not total:
        return nuevo, problemas

    piso = max(IMPORTE_MINIMO_CREIBLE, abs(total) * 0.001)

    for clave in CAMPOS_REVISABLES:
        valor = nuevo.get(clave)
        if not valor:
            continue
        etiqueta = NOMBRE_CAMPO.get(clave, clave)

        if abs(valor) < piso and abs(total) > piso * 4:
            problemas.append("saque %s de $%s: es demasiado chico para esta "
                             "factura y no parece un importe real. Cargalo vos "
                             "si corresponde." % (etiqueta, _plata(valor)))
            nuevo[clave] = 0.0
            continue

        if valor < 0 and not es_credito:
            problemas.append("%s daba negativo ($%s) en una factura que no es "
                             "nota de credito. Lo deje en cero: revisalo."
                             % (etiqueta, _plata(valor)))
            nuevo[clave] = 0.0

    # el IVA tiene que guardar relacion con el neto
    neto = nuevo.get("neto_gravado") or 0
    for clave, tasa in ALICUOTAS.items():
        iva = nuevo.get(clave) or 0
        if not iva or not neto:
            continue
        esperado = neto * tasa
        if abs(iva - esperado) > max(2.0, abs(esperado) * 0.05):
            problemas.append("el %s ($%s) no da con el neto ($%s): deberia "
                             "andar por $%s. Revisalo."
                             % (NOMBRE_CAMPO[clave], _plata(iva), _plata(neto),
                                _plata(round(esperado, 2))))
    return nuevo, problemas


def _plata(v):
    return "{:,.2f}".format(v or 0).replace(",", "@").replace(".", ",").replace("@", ".")


ALICUOTAS = {"iva_105": 0.105, "iva_21": 0.21, "iva_27": 0.27}


def _alicuota_en_juego(imp, texto=""):
    """
    Cual es la alicuota de IVA de este comprobante.
    Si hay mas de una, devuelve None: ahi no se puede deducir nada sin inventar.
    """
    con_valor = [k for k in ALICUOTAS if imp.get(k)]
    if len(con_valor) == 1:
        return con_valor[0]
    if len(con_valor) > 1:
        return None

    # ninguna alicuota con importe: se busca mencionada en el texto
    n = normalizar(texto)
    hallazgos = []
    if re.search(r"(I\.?V\.?A\.?|AL[IÍ1]CUOTA)[^\d%]{0,12}10[.,]5", n):
        hallazgos.append("iva_105")
    if re.search(r"(I\.?V\.?A\.?|AL[IÍ1]CUOTA)[^\d%]{0,12}27([.,]0+)?\s*%", n):
        hallazgos.append("iva_27")
    if re.search(r"(I\.?V\.?A\.?|AL[IÍ1]CUOTA)[^\d%]{0,12}21([.,]0+)?\s*%?", n):
        hallazgos.append("iva_21")
    return hallazgos[0] if len(hallazgos) == 1 else None


def completar_por_alicuota(imp, total, texto="", tolerancia=0.05):
    """
    Deduce lo que falta a partir del total y la alicuota de IVA.

    En una factura vale que  neto + IVA + no gravado + exento + otros = total
    y que  IVA = neto x alicuota.  Con el total y la alicuota, el neto y el IVA
    salen solos:
        neto = (total - lo no gravado) / (1 + alicuota)
        IVA  = neto x alicuota

    Reglas para no inventar:
      - Solo se intenta si las partes NO cuadran.
      - Solo si hay UNA alicuota en juego.
      - Si un importe ya se habia leido y la cuenta da algo muy distinto,
        no se toca nada: es mejor avisar que tapar el problema.
      - El resultado se acepta solo si despues cuadra exacto contra el total.

    Devuelve (importes, lista de lo deducido).
    """
    if not total:
        return imp, []
    if abs(_suma(imp) - total) <= tolerancia:
        return imp, []

    clave = _alicuota_en_juego(imp, texto)
    if clave is None:
        return imp, []
    tasa = ALICUOTAS[clave]

    # Freno importante: si la factura NOMBRA conceptos no gravados o exentos
    # pero no se pudo leer cuanto son, deducir mandaria todo a gravado, que es
    # justo lo que no hay que hacer (seguros, servicios publicos). Mejor avisar.
    n = normalizar(texto)
    if not imp.get("no_gravado") and re.search(r"\bNO\s+GRAVAD", n):
        return imp, []
    if not imp.get("exento") and re.search(r"\bEXENT", n):
        return imp, []

    resto = round((imp.get("no_gravado") or 0) + (imp.get("exento") or 0)
                  + (imp.get("otros_tributos") or 0), 2)
    gravado = round(total - resto, 2)
    if gravado <= 0:
        return imp, []

    neto = round(gravado / (1 + tasa), 2)
    iva = round(gravado - neto, 2)

    # si ya habia un valor leido, la cuenta tiene que dar parecido
    for campo, calculado in (("neto_gravado", neto), (clave, iva)):
        leido = imp.get(campo)
        if leido and abs(leido - calculado) > max(1.0, abs(calculado) * 0.02):
            return imp, []

    nuevo = dict(imp)
    deducidos = []
    etiqueta = {"iva_105": "IVA 10,5%", "iva_21": "IVA 21%",
                "iva_27": "IVA 27%"}[clave]

    if abs((nuevo.get("neto_gravado") or 0) - neto) > 0.005:
        nuevo["neto_gravado"] = neto
        deducidos.append("neto gravado")
    if abs((nuevo.get(clave) or 0) - iva) > 0.005:
        nuevo[clave] = iva
        deducidos.append(etiqueta)
    for otra in ALICUOTAS:
        if otra != clave:
            nuevo[otra] = nuevo.get(otra) or 0.0

    if deducidos and abs(_suma(nuevo) - total) <= tolerancia:
        return nuevo, deducidos
    return imp, []


def extraer_razon_social(texto, cuit_emisor):
    """
    Busca el nombre del proveedor. Suele estar en las primeras lineas,
    cerca del CUIT del emisor.
    """
    lineas = [l.strip() for l in (texto or "").splitlines()]
    cuit = str(cuit_emisor or "")
    cuit_fmt = ("%s-%s-%s" % (cuit[:2], cuit[2:10], cuit[10:])) if len(cuit) == 11 else None

    # candidatas: lineas con forma juridica explicita
    forma = re.compile(
        r"\b(S\.?\s?A\.?\s?S\.?|S\.?\s?R\.?\s?L\.?|S\.?\s?A\.?|S\.?\s?C\.?\s?A\.?"
        r"|SOCIEDAD\s+ANONIMA|COOPERATIVA|ASOCIACION|FUNDACION|S\.?\s?H\.?)\s*$",
        re.IGNORECASE)

    ruido = ("FACTURA", "ORIGINAL", "DUPLICADO", "TRIPLICADO", "CODIGO",
             "C.U.I.T", "CUIT", "FECHA", "SENORES", "INGRESOS BRUTOS",
             "PUNTO DE VENTA", "RESPONSABLE INSCRIPTO", "INICIO ACTIVIDADES",
             "DOMICILIO", "CONDICION", "REMITO", "PRESUPUESTO", "TELEFONO",
             "WWW.", "EMAIL", "@")

    def trozos(linea):
        """El texto viene en columnas; se parten por los espacios grandes."""
        return [p.strip() for p in re.split(r"\s{3,}", linea) if p.strip()]

    def limpiar(cand):
        return re.sub(r"\s+", " ", cand).strip(" -:,")

    def sirve(cand):
        if not (4 <= len(cand) <= 70):
            return False
        n = normalizar(cand)
        if any(x in n for x in ruido):
            return False
        if re.search(r"\d{5,}", cand):
            return False
        return True

    tope = min(len(lineas), 25)
    for i in range(tope):
        for trozo in trozos(lineas[i]):
            cand = limpiar(trozo)
            if not sirve(cand):
                continue
            if forma.search(cand) or re.search(
                    r"\bS\.?\s?A\.?\b|\bS\.?\s?R\.?\s?L\.?\b", cand):
                return cand.upper()

    # si no hubo suerte, la linea no vacia justo arriba del CUIT del emisor
    for i, l in enumerate(lineas):
        if (cuit_fmt and cuit_fmt in l) or (cuit and cuit in re.sub(r"[.\-\s]", "", l)):
            for j in range(i - 1, max(-1, i - 6), -1):
                for trozo in trozos(lineas[j]):
                    cand = limpiar(trozo)
                    if sirve(cand):
                        return cand.upper()
            break

    # ultimo recurso: la primera linea que parezca un nombre
    for i in range(tope):
        for trozo in trozos(lineas[i]):
            cand = limpiar(trozo)
            if sirve(cand) and " " in cand:
                return cand.upper()
    return None


# =============================================================================
# RESPALDO: datos de cabecera desde el texto (si no se pudo leer el QR)
# =============================================================================

RE_CUIT = re.compile(r"(?<!\d)(\d{2})[\s.\-]?(\d{8})[\s.\-]?(\d)(?!\d)")


def cuit_valido(cuit):
    """
    Verifica el digito verificador del CUIT (modulo 11).
    Sirve para no confundir un CUIT con cualquier otro numero de 11 cifras.
    """
    s = re.sub(r"\D", "", str(cuit))
    if len(s) != 11:
        return False
    pesos = (5, 4, 3, 2, 7, 6, 5, 4, 3, 2)
    suma = sum(int(d) * p for d, p in zip(s[:10], pesos))
    resto = suma % 11
    verificador = 11 - resto
    if verificador == 11:
        verificador = 0
    elif verificador == 10:
        verificador = 9
    return verificador == int(s[10])


def cuits_del_texto(texto):
    """Todos los CUIT que aparecen en el texto, en orden y sin repetir."""
    vistos, encontrados = set(), []
    for m in RE_CUIT.finditer(texto or ""):
        cuit = int(m.group(1) + m.group(2) + m.group(3))
        if cuit in vistos or not cuit_valido(cuit):
            continue
        vistos.add(cuit)
        encontrados.append(cuit)
    return encontrados


# Tipo de comprobante buscado por texto. Se evalua en orden.
PATRONES_TIPO = [
    (3,  r"(NOTA\s+DE\s+)?CREDITO\s*[\"']?\s*A\b|\bNC\s*A\b"),
    (8,  r"(NOTA\s+DE\s+)?CREDITO\s*[\"']?\s*B\b|\bNC\s*B\b"),
    (13, r"(NOTA\s+DE\s+)?CREDITO\s*[\"']?\s*C\b|\bNC\s*C\b"),
    (2,  r"(NOTA\s+DE\s+)?DEBITO\s*[\"']?\s*A\b|\bND\s*A\b"),
    (7,  r"(NOTA\s+DE\s+)?DEBITO\s*[\"']?\s*B\b|\bND\s*B\b"),
    (12, r"(NOTA\s+DE\s+)?DEBITO\s*[\"']?\s*C\b|\bND\s*C\b"),
    (1,  r"FACTURA\s*[\"']?\s*A\b|COD(IGO)?\.?\s*N?\s*[º°]?\s*0?1\b"),
    (6,  r"FACTURA\s*[\"']?\s*B\b|COD(IGO)?\.?\s*N?\s*[º°]?\s*0?6\b"),
    (11, r"FACTURA\s*[\"']?\s*C\b|COD(IGO)?\.?\s*N?\s*[º°]?\s*11\b"),
    (19, r"FACTURA\s*[\"']?\s*E\b"),
    (51, r"FACTURA\s*[\"']?\s*M\b"),
]


def _empresa_mas_parecida(texto_normalizado, maximo_errores=2):
    """
    Busca en el texto un numero de 11 cifras que se parezca al CUIT de alguna
    de nuestras empresas. Sirve cuando el OCR erro un digito y por eso el CUIT
    no pasa el digito verificador.
    """
    candidatos = set()
    for m in re.finditer(r"(?<!\d)(\d[\d\s.\-]{9,16}\d)(?!\d)",
                         texto_normalizado):
        limpio = re.sub(r"\D", "", m.group(1))
        if len(limpio) == 11:
            candidatos.add(limpio)

    mejor, mejor_errores = None, maximo_errores + 1
    for propio in EMPRESAS_CLIENTE:
        texto_propio = str(propio)
        for cand in candidatos:
            errores = sum(1 for a, b in zip(cand, texto_propio) if a != b)
            if errores < mejor_errores:
                mejor, mejor_errores = propio, errores
    return mejor if mejor_errores <= maximo_errores else None


# Cosas que la gente guarda junto con las facturas pero NO son facturas.
NO_ES_FACTURA = [
    (r"CUPON\s+DE\s+PAGO|CUPON\s+DE\s+COBRO", "un cupon de pago"),
    (r"AVISO\s+DE\s+VENCIMIENTO", "un aviso de vencimiento"),
    (r"\bPRESUPUESTO\b", "un presupuesto"),
    (r"\bREMITO\b(?!.*FACTURA)", "un remito"),
    (r"ORDEN\s+DE\s+COMPRA", "una orden de compra"),
    (r"N[º°.\s]{0,4}DE\s+RESUMEN|RESUMEN\s+DE\s+CUENTA|ESTADO\s+DE\s+CUENTA",
     "un resumen de cuenta"),
    (r"COMPROBANTE\s+DE\s+PAGO", "un comprobante de pago"),
]


def que_parece_ser(texto):
    """
    Si el PDF no parece una factura, decirlo. Es mejor avisar que cargar
    cualquier cosa: un cupon de pago o un resumen de cuenta no van al libro
    de compras.
    """
    n = normalizar(texto)
    for patron, etiqueta in NO_ES_FACTURA:
        if re.search(patron, n):
            return etiqueta
    return None


def cabecera_desde_texto(texto, cuits_extra=None):
    """
    Saca los datos de cabecera leyendo el texto de la factura.
    Es el plan B cuando no aparece el codigo QR de AFIP.

    cuits_extra: CUIT hallados en otras pasadas de OCR. Alcanza con que una
    sola pasada los haya leido bien.
    """
    d = {}
    n = normalizar(texto)

    # --- CUIT: el propio es el receptor, el otro es el proveedor ------------
    cuits = cuits_del_texto(n)
    for c in (cuits_extra or []):
        if c not in cuits:
            cuits.append(c)

    # El CUIT del cliente somos siempre nosotros, y son solo dos posibles.
    # Si el OCR erro uno o dos digitos, el numero no pasa el verificador y se
    # descarta; aca se rescata comparandolo con los nuestros.
    if not any(c in EMPRESAS_CLIENTE for c in cuits):
        propio = _empresa_mas_parecida(n)
        if propio:
            cuits.append(propio)
    propios = [c for c in cuits if c in EMPRESAS_CLIENTE]
    ajenos = [c for c in cuits if c not in EMPRESAS_CLIENTE]
    if propios:
        d["nroDocRec"] = propios[0]
    if ajenos:
        d["cuit"] = ajenos[0]
    # Si el UNICO CUIT del texto es el nuestro, entonces es el del cliente y el
    # del proveedor no esta (suele quedar dentro del logo, que es una imagen).
    # Antes se usaba igual como proveedor: eso cargaba la factura a nombre de
    # tu propia empresa. Mejor dejarlo vacio y que lo completes vos.

    # --- numero de comprobante ---------------------------------------------
    patrones_nro = [
        r"(?:COMP\.?|COMPROBANTE|FACTURA|FACT\.?|NRO\.?|N[º°]|N\s*RO)"
        r"[^\d]{0,10}(\d{4,5})\s*-\s*(\d{6,8})",
        r"(?<!\d)(\d{4,5})\s*-\s*(\d{8})(?!\d)",
        r"(?<!\d)(\d{4,5})\s*-\s*(\d{7})(?!\d)",
    ]
    for patron in patrones_nro:
        m = re.search(patron, n)
        if m:
            d["ptoVta"] = int(m.group(1))
            d["nroCmp"] = int(m.group(2))
            break
    else:
        # facturas que traen el punto de venta y el numero en campos separados
        pv = re.search(r"P(?:UNTO|TO)\.?\s*DE\s*VENTA[^\d]{0,10}(\d{1,5})", n)
        nc = re.search(r"COMP\.?\s*(?:NRO|NUMERO|N[º°]?)\.?[^\d]{0,10}(\d{1,8})", n)
        if pv and nc:
            d["ptoVta"] = int(pv.group(1))
            d["nroCmp"] = int(nc.group(1))

    # --- CAE ----------------------------------------------------------------
    m = re.search(r"\bCAE[A-Z]?\b[^\d]{0,15}(\d{14})", n)
    if m:
        d["codAut"] = int(m.group(1))

    # --- fecha: la de emision, nunca la de vencimiento ----------------------
    for linea in n.splitlines():
        if re.search(r"VENCIMIENTO|\bVTO\b|\bCAE\b|DESDE|HASTA|PERIODO", linea):
            continue
        m = re.search(r"FECHA[^\d]{0,15}(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})", linea)
        if m:
            d["fecha_txt"] = m.group(1)
            break
    if "fecha_txt" not in d:
        m = re.search(r"FECHA(?!\s*(?:DE\s+)?VENC)(?:\s+DE)?\s*(?:EMISION)?"
                      r"[^\d]{0,15}(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})", n)
        if m:
            d["fecha_txt"] = m.group(1)
    if "fecha_txt" not in d:
        # Los tiques de controlador fiscal muchas veces no dicen "Fecha", o el
        # OCR se come la palabra. Se toma la fecha mas reciente que sea
        # plausible: ni futura ni de hace anios (eso descarta cosas como
        # "Inicio de Actividades: 01/01/2013").
        hoy = datetime.date.today()
        candidatas = []
        for txt in re.findall(r"\b(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b", n):
            for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y"):
                try:
                    f = datetime.datetime.strptime(txt, fmt).date()
                except ValueError:
                    continue
                if 0 <= (hoy - f).days <= 730:
                    candidatas.append((f, txt))
                break
        if candidatas:
            d["fecha_txt"] = max(candidatas)[1]

    # --- tipo de comprobante ------------------------------------------------
    for codigo, patron in PATRONES_TIPO:
        if re.search(patron, n):
            d["tipoCmp"] = codigo
            break

    # --- moneda extranjera --------------------------------------------------
    if re.search(r"\bDOLAR|\bU\$S|\bUSD\b|\bDOL\b", n):
        d["moneda"] = "DOL"
        m = re.search(r"(?:TIPO\s+DE\s+CAMBIO|COTIZACION|T\.?C\.?)"
                      r"[^\d]{0,12}([\d.,]+)", n)
        if m:
            ctz = parse_num(m.group(1))
            if ctz and ctz > 1:
                d["ctz"] = ctz
    return d


# =============================================================================
# LECTURA COMPLETA DE UNA FACTURA
# =============================================================================

def _fecha_desde(qr, texto_cab):
    if qr and qr.get("fecha"):
        try:
            return datetime.datetime.strptime(qr["fecha"], "%Y-%m-%d").date()
        except ValueError:
            pass
    txt = texto_cab.get("fecha_txt")
    if txt:
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y"):
            try:
                return datetime.datetime.strptime(txt, fmt).date()
            except ValueError:
                continue
    return None


def leer_factura_pdf(path, nombres_por_cuit=None):
    """
    Lee una factura en PDF.

    Devuelve un dict con:
        ok           - si se pudo leer lo suficiente como para seguir
        escaneo      - True si el PDF no tiene texto (es una foto)
        avisos       - lista de cosas para mirar con lupa
        origen       - 'QR' o 'texto', de donde salio la cabecera
        cuadra       - si los importes suman el total
        ... y todos los campos de la factura
    """
    nombres_por_cuit = nombres_por_cuit or {}
    avisos = []

    faltan = dependencias_faltantes()
    if faltan:
        return {"ok": False, "escaneo": False, "archivo": path,
                "avisos": ["Faltan librerias: %s" % ", ".join(faltan)]}

    texto = leer_texto(path)
    escaneado = es_escaneo(texto)

    qr = leer_qr_afip(path)
    if qr is None and cv2 is None:
        avisos.append("No esta instalado opencv, asi que no pude leer el QR de AFIP. "
                      "Los datos salen del texto del PDF: revisalos bien.")
    elif qr is None:
        avisos.append("No encontre el codigo QR de AFIP. Los datos salen del texto "
                      "del PDF: revisalos bien.")

    # PDF escaneado: se intenta con OCR antes de darse por vencido
    ocr_usado = False
    cuits_ocr = []
    if escaneado:
        texto_ocr, cuits_ocr = leer_texto_ocr(path)
        if not es_escaneo(texto_ocr):
            texto = texto_ocr
            escaneado = False
            ocr_usado = True
            avisos.append("El PDF es un escaneo: los importes los saque con OCR. "
                          "REVISALOS UNO POR UNO, el OCR se equivoca con los "
                          "numeros.")

    if escaneado and not qr:
        disponible, motivo = estado_ocr()
        pista = "" if disponible else " No pude usar OCR porque %s." % motivo
        return {"ok": False, "escaneo": True, "archivo": path,
                "avisos": ["El PDF es un escaneo: no tiene texto adentro y "
                           "tampoco pude leer el QR.%s" % pista]}

    if escaneado and qr:
        avisos.append("El PDF es un escaneo: del QR saque proveedor, fecha, "
                      "numero y total, pero el desglose (neto, IVA) lo tenes "
                      "que cargar vos.")

    cab_txt = cabecera_desde_texto(texto, cuits_extra=cuits_ocr)
    origen = "OCR" if ocr_usado else ("QR" if qr else "texto")

    # --- cabecera (el QR manda; lo que falte se busca en el texto) ---
    cuit_emisor = (qr or {}).get("cuit") or cab_txt.get("cuit")
    cuit_receptor = (qr or {}).get("nroDocRec") or cab_txt.get("nroDocRec")
    pto_venta = (qr or {}).get("ptoVta") or cab_txt.get("ptoVta")
    nro_cmp = (qr or {}).get("nroCmp") or cab_txt.get("nroCmp")
    tipo_cmp = (qr or {}).get("tipoCmp") or cab_txt.get("tipoCmp")
    cae = (qr or {}).get("codAut") or cab_txt.get("codAut")
    fecha = _fecha_desde(qr, cab_txt)
    total = (qr or {}).get("importe")
    moneda_iso = (qr or {}).get("moneda") or cab_txt.get("moneda") or "PES"
    cotizacion = (qr or {}).get("ctz") or cab_txt.get("ctz") or 1

    # --- desglose ---
    imp = extraer_importes(texto)
    if total is None:
        total = _total_sensato(imp)
    # importes imposibles (mayores que el total): no eran ese importe
    imp, descartados = descartar_imposibles(imp, total)
    if descartados:
        avisos.append("Descarte %s porque daba mas que el total de la factura."
                      % " y ".join(descartados))

    imp, migajas = descartar_migajas(imp, total)
    if migajas:
        avisos.append("Ignore %s: eran numeros sueltos, demasiado chicos para "
                      "ser ese importe." % " y ".join(migajas))

    # si al OCR se le comio alguna coma decimal, se arregla contra el total
    imp, reparado = reparar_por_cuadre(imp, total)
    if reparado:
        avisos.append("Algun importe venia sin la coma decimal; lo corregi "
                      "para que las partes sumen el total. Verificalo.")

    # si todavia no cuadra, se deduce lo que falte a partir del total
    # y la alicuota de IVA
    imp, deducidos = completar_por_alicuota(imp, total, texto)
    if deducidos:
        avisos.append("No pude leer %s, asi que lo calcule a partir del total "
                      "y la alicuota de IVA. Da exacto, pero verificalo."
                      % " y ".join(deducidos))

    # control final: antes de escribir un numero raro, sacarlo y avisar
    imp, incoherencias = revisar_coherencia(
        imp, total, es_credito=(tipo_cmp in TIPOS_CREDITO))
    avisos.extend(incoherencias)

    neto = imp.get("neto_gravado")
    if neto is None:
        neto = imp.get("subtotal")      # el subtotal suele ser el neto
    no_grav = imp.get("no_gravado") or 0.0
    exento = imp.get("exento") or 0.0
    iva_105 = imp.get("iva_105") or 0.0
    iva_21 = imp.get("iva_21") or 0.0
    iva_27 = imp.get("iva_27") or 0.0
    iva_total = imp.get("iva_total")
    perc_iibb = imp.get("percepcion_iibb") or 0.0
    perc_iva = imp.get("percepcion_iva") or 0.0
    otros = imp.get("otros_tributos") or 0.0
    otros_tributos = round(perc_iibb + perc_iva + otros, 2)

    # Las alicuotas mandan sobre el "IVA total" leido suelto: muchas facturas
    # tienen una linea que dice solo "IVA" (un encabezado de columna, por
    # ejemplo) de la que se saca un 0 que despues pisaba el IVA de verdad.
    suma_alicuotas = round(iva_105 + iva_21 + iva_27, 2)

    # Si hay un IVA total de verdad y las alicuotas no lo suman, las alicuotas
    # estan mal leidas (suelen venir de una linea de forma de pago con las
    # columnas pegadas). Manda el IVA total.
    if iva_total and suma_alicuotas and abs(suma_alicuotas - iva_total) > 0.05:
        # Una de las dos lecturas esta mal. En vez de elegir a ciegas, se
        # prueban las dos contra el total del comprobante y gana la que cierra.
        resto = (no_grav or 0) + (exento or 0) + otros_tributos
        base = neto or 0
        if total:
            con_alicuotas = abs(base + suma_alicuotas + resto - total)
            con_iva_total = abs(base + iva_total + resto - total)
        else:
            con_alicuotas, con_iva_total = 0.0, 1.0
        if con_iva_total + 0.01 < con_alicuotas:
            iva_105 = iva_21 = iva_27 = 0.0
            suma_alicuotas = 0.0
            avisos.append("Las alicuotas de IVA no cerraban contra el total; "
                          "me quede con el IVA total. Verificalo.")
        else:
            iva_total = suma_alicuotas

    if suma_alicuotas != 0:
        iva_total = suma_alicuotas
    elif iva_total is None:
        iva_total = suma_alicuotas
    elif suma_alicuotas == 0 and iva_total:
        # se conoce el IVA total pero no la alicuota: se deduce del neto
        if neto and abs(iva_total - neto * 0.21) < max(1.0, neto * 0.005):
            iva_21 = iva_total
        elif neto and abs(iva_total - neto * 0.105) < max(1.0, neto * 0.005):
            iva_105 = iva_total
        elif neto and abs(iva_total - neto * 0.27) < max(1.0, neto * 0.005):
            iva_27 = iva_total
        else:
            iva_21 = iva_total
            avisos.append("No pude identificar la alicuota de IVA. La puse en 21%%: "
                          "confirmala.")

    if neto is None and iva_21:
        neto = round(iva_21 / 0.21, 2)
        avisos.append("El neto gravado no figuraba: lo calcule a partir del IVA.")
    if neto is None:
        neto = 0.0

    # comprobante tipo B o C: AFIP no discrimina el IVA
    letra = ""
    if tipo_cmp in TIPOS_COMPROBANTE:
        m = re.search(r"\b([ABCEM])\b", TIPOS_COMPROBANTE[tipo_cmp])
        letra = m.group(1) if m else ""
    if letra in ("B", "C") and iva_total == 0 and total:
        no_grav = no_grav or 0.0
        if neto == 0 and no_grav == 0:
            no_grav = total
            avisos.append("Comprobante %s sin IVA discriminado: el total va "
                          "entero a No Gravado." % letra)

    # --- control de cuadre ---
    suma = round(neto + iva_total + no_grav + exento + otros_tributos, 2)
    cuadra = total is not None and abs(suma - total) <= 0.05

    if not cuadra and total is not None:
        dif = round(total - suma, 2)
        if abs(dif) > 0.05:
            avisos.append("Los importes no cuadran con el total del comprobante: "
                          "faltan %.2f. Revisa el desglose antes de aceptar." % dif)

    # si el CUIT ya es conocido, se usa el nombre de siempre; si no, el del PDF
    razon = nombre_de_cuit(cuit_emisor, nombres_por_cuit)
    if not razon:
        razon = extraer_razon_social(texto, cuit_emisor)
        if not razon:
            avisos.append("No pude leer la razon social del proveedor.")

    parece = que_parece_ser(texto)
    if parece and not qr:
        avisos.append("OJO: esto parece %s, no una factura. Si no va al libro "
                      "de compras, salteala con la S." % parece)

    faltantes = [n for n, v in (("fecha", fecha), ("CUIT del emisor", cuit_emisor),
                                ("punto de venta", pto_venta),
                                ("numero de comprobante", nro_cmp),
                                ("tipo de comprobante", tipo_cmp),
                                ("total", total)) if not v]
    if faltantes:
        avisos.append("No pude leer: %s." % ", ".join(faltantes))

    return {
        "ok": not faltantes and not incoherencias,
        "incoherencias": incoherencias,
        "escaneo": escaneado,
        "archivo": path,
        "origen": origen,
        "avisos": avisos,
        "cuadra": cuadra,
        "fecha": fecha,
        "tipo_cmp": tipo_cmp,
        "letra": letra,
        "pto_venta": pto_venta,
        "nro_cmp": nro_cmp,
        "cae": cae,
        "cuit_emisor": cuit_emisor,
        "razon_social": razon,
        "cuit_receptor": cuit_receptor,
        "moneda": MONEDAS.get(moneda_iso, moneda_iso),
        "cotizacion": cotizacion or 1,
        "neto_gravado": round(neto, 2),
        "iva_105": round(iva_105, 2),
        "iva_21": round(iva_21, 2),
        "iva_27": round(iva_27, 2),
        "iva_total": round(iva_total, 2),
        "no_gravado": round(no_grav, 2),
        "exento": round(exento, 2),
        "otros_tributos": otros_tributos,
        "total": round(total, 2) if total is not None else None,
    }


# =============================================================================
# CONVERSION A FILA "MIS COMPROBANTES"
# =============================================================================

def a_fila_arca(d):
    """
    Arma la fila de 30 columnas identica a la del export de ARCA, para que el
    resto del script no note la diferencia entre un PDF y un archivo de AFIP.
    """
    tipo = TIPOS_COMPROBANTE.get(d.get("tipo_cmp"), "Factura A")
    etiqueta = "%s - %s" % (d.get("tipo_cmp") or 1, tipo)
    fecha = d.get("fecha")
    signo = -1 if d.get("tipo_cmp") in TIPOS_CREDITO else 1

    def m(v):
        return round((v or 0.0) * signo, 2)

    fila = [None] * 30
    fila[0] = fecha.strftime("%d/%m/%Y") if fecha else None
    fila[1] = etiqueta
    fila[2] = d.get("pto_venta")
    fila[3] = d.get("nro_cmp")
    fila[4] = d.get("nro_cmp")
    fila[5] = d.get("cae")
    fila[6] = "CUIT"
    fila[7] = d.get("cuit_emisor")
    fila[8] = d.get("razon_social") or ""
    fila[9] = "CUIT"
    fila[10] = d.get("cuit_receptor")
    fila[11] = d.get("cotizacion") or 1
    fila[12] = d.get("moneda") or "$"
    fila[13] = 0.0                       # Neto Grav. IVA 0%
    fila[14] = 0.0                       # IVA 2,5%
    fila[15] = 0.0                       # Neto Grav. IVA 2,5%
    fila[16] = 0.0                       # IVA 5%
    fila[17] = 0.0                       # Neto Grav. IVA 5%
    fila[18] = m(d.get("iva_105"))
    fila[19] = m(round((d.get("iva_105") or 0) / 0.105, 2)) if d.get("iva_105") else 0.0
    fila[20] = m(d.get("iva_21"))
    fila[21] = m(round((d.get("iva_21") or 0) / 0.21, 2)) if d.get("iva_21") else 0.0
    fila[22] = m(d.get("iva_27"))
    fila[23] = m(round((d.get("iva_27") or 0) / 0.27, 2)) if d.get("iva_27") else 0.0
    fila[24] = m(d.get("neto_gravado"))
    fila[25] = m(d.get("no_gravado"))
    fila[26] = m(d.get("exento"))
    fila[27] = m(d.get("otros_tributos"))
    fila[28] = m(d.get("iva_total"))
    fila[29] = m(d.get("total"))

    # el neto por alicuota se recalcula para que sume exactamente el neto total
    suma_alic = round(fila[19] + fila[21] + fila[23], 2)
    if suma_alic and abs(suma_alic - fila[24]) > 0.02:
        # la unica alicuota presente se lleva todo el neto
        presentes = [i for i in (19, 21, 23) if fila[i]]
        if len(presentes) == 1:
            fila[presentes[0]] = fila[24]
    return fila


# =============================================================================
# PANTALLA DE REVISION
# =============================================================================

# Empresas del grupo, para elegir el cliente sin tener que escribir el CUIT.
# procesar_migracion.py la sobreescribe con su propia tabla EMPRESAS al arrancar.
EMPRESAS_CLIENTE = {
    30710230184: ("MEPANO", "MEPANO SOCIEDAD ANONIMA"),
    30554641748: ("MPN", "METALURGICA PABLO NOGUES SRL"),
}


def registrar_empresas(empresas):
    """
    Carga la lista de empresas propias desde procesar_migracion.py.
    Espera {"30710230184": {"nombre": "MEPANO", "razon_social": "..."}, ...}
    """
    global EMPRESAS_CLIENTE
    nueva = {}
    for cuit, datos in (empresas or {}).items():
        try:
            nueva[int(cuit)] = (datos.get("nombre", ""), datos.get("razon_social", ""))
        except (TypeError, ValueError):
            continue
    if nueva:
        EMPRESAS_CLIENTE = nueva


# Nombres de proveedor por CUIT (json arch/_proveedores.json).
# procesar_migracion.py la carga al arrancar.
NOMBRES_PROVEEDOR = {}


def registrar_proveedores(mapa):
    """Carga el mapa CUIT -> nombre de proveedor para autocompletar."""
    global NOMBRES_PROVEEDOR
    limpio = {}
    for cuit, nombre in (mapa or {}).items():
        try:
            limpio[int(cuit)] = nombre
        except (TypeError, ValueError):
            continue
    NOMBRES_PROVEEDOR = limpio


def nombre_de_cuit(cuit, extra=None):
    """Busca el nombre de un proveedor por su CUIT. None si no lo conoce."""
    try:
        clave = int(cuit)
    except (TypeError, ValueError):
        return None
    if extra and clave in extra:
        return extra[clave]
    return NOMBRES_PROVEEDOR.get(clave)


def autocompletar_nombre(d, extra=None, avisar=True):
    """
    Al cargar el CUIT, completa solo el nombre del proveedor si ya lo conocemos,
    para que quede escrito igual que en las cargas anteriores.
    """
    nombre = nombre_de_cuit(d.get("cuit_emisor"), extra)
    if nombre and nombre != d.get("razon_social"):
        d["razon_social"] = nombre
        if avisar:
            print("      -> Proveedor conocido: %s" % nombre)
    return d


# (nombre interno, etiqueta que se muestra, tipo)
CAMPOS_REVISION = [
    ("razon_social", "Proveedor",              "texto"),
    ("cuit_emisor",  "CUIT del proveedor",     "cuit"),
    ("tipo_cmp",     "Tipo de comprobante",    "tipo"),
    ("pto_venta",    "Punto de venta",         "entero"),
    ("nro_cmp",      "Numero de comprobante",  "entero"),
    ("fecha",        "Fecha de la factura",    "fecha"),
    ("cuit_receptor", "Cliente (cual de las dos empresas)", "empresa"),
    ("moneda",       "Moneda",                 "texto"),
    ("cotizacion",   "Tipo de cambio",         "importe"),
    ("neto_gravado", "Neto gravado",           "importe"),
    ("iva_105",      "IVA 10,5%",              "importe"),
    ("iva_21",       "IVA 21%",                "importe"),
    ("iva_27",       "IVA 27%",                "importe"),
    ("no_gravado",   "No gravado",             "importe"),
    ("exento",       "Exento",                 "importe"),
    ("otros_tributos", "Percepciones / otros tributos", "importe"),
    ("total",        "TOTAL",                  "importe"),
]

# En la carga manual estos no se preguntan: quedan en pesos y tipo de cambio 1,
# que es lo que pasa el 99% de las veces. Igual se pueden corregir despues,
# desde la pantalla de revision, si la factura vino en dolares.
NO_SE_PREGUNTAN_A_MANO = ("moneda", "cotizacion")


def _mostrar(v, tipo):
    if v is None:
        return "(vacio)"
    if tipo == "fecha":
        return v.strftime("%d/%m/%Y")
    if tipo == "importe":
        return "{:,.2f}".format(v).replace(",", "@").replace(".", ",").replace("@", ".")
    if tipo == "tipo":
        return "%s - %s" % (v, TIPOS_COMPROBANTE.get(v, "?"))
    if tipo == "empresa":
        try:
            nombre, razon = EMPRESAS_CLIENTE[int(v)]
            return "%s  (%s)" % (razon, nombre)
        except (KeyError, TypeError, ValueError):
            return "%s  (NO es ninguna de las tuyas)" % v
    return str(v)


def _elegir_empresa(actual):
    """Menu para elegir el cliente sin escribir el CUIT."""
    opciones = sorted(EMPRESAS_CLIENTE.items(), key=lambda x: x[1][0])
    print()
    print("      De que empresa es esta factura?")
    for i, (cuit, (nombre, razon)) in enumerate(opciones, 1):
        marca = "  <- actual" if actual and int(actual) == cuit else ""
        print("        %d) %-12s %s%s" % (i, nombre, razon, marca))
    pista = " (ENTER deja como esta)" if actual else ""
    while True:
        crudo = input("      Elegi una opcion (1-%d)%s: "
                      % (len(opciones), pista)).strip()
        if crudo == "":
            if actual:
                return actual
            print("      Tenes que elegir una de las dos.")
            continue
        if crudo.isdigit() and 1 <= int(crudo) <= len(opciones):
            elegido = opciones[int(crudo) - 1]
            print("      -> %s" % elegido[1][1])
            return elegido[0]
        # tambien se acepta el nombre corto o el CUIT escrito entero
        limpio = re.sub(r"\D", "", crudo)
        if limpio and int(limpio) in EMPRESAS_CLIENTE:
            return int(limpio)
        for cuit, (nombre, _razon) in opciones:
            if crudo.upper() == nombre.upper():
                return cuit
        print("      No entendi. Escribi el numero de la opcion.")


def _pedir(etiqueta, tipo, actual):
    while True:
        if tipo == "empresa":
            return _elegir_empresa(actual)
        crudo = input("   %s [%s]: " % (etiqueta, _mostrar(actual, tipo))).strip()
        if crudo == "":
            return actual
        if tipo == "fecha":
            for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y"):
                try:
                    return datetime.datetime.strptime(crudo, fmt).date()
                except ValueError:
                    continue
            print("      No entendi la fecha. Escribila asi: 20/08/2026")
            continue
        if tipo in ("entero", "cuit", "tipo"):
            limpio = re.sub(r"\D", "", crudo)
            if limpio:
                return int(limpio)
            print("      Tiene que ser un numero.")
            continue
        if tipo == "importe":
            v = parse_num(crudo)
            if v is not None:
                return v
            print("      No entendi el importe. Ejemplo: 217602,90")
            continue
        return crudo


def _recalcular(d):
    d["iva_total"] = round((d.get("iva_105") or 0) + (d.get("iva_21") or 0)
                           + (d.get("iva_27") or 0), 2)
    suma = round((d.get("neto_gravado") or 0) + d["iva_total"]
                 + (d.get("no_gravado") or 0) + (d.get("exento") or 0)
                 + (d.get("otros_tributos") or 0), 2)
    total = d.get("total")
    d["suma_componentes"] = suma
    d["cuadra"] = total is not None and abs(suma - total) <= 0.05
    if d.get("tipo_cmp") in TIPOS_COMPROBANTE:
        m = re.search(r"\b([ABCEM])\b", TIPOS_COMPROBANTE[d["tipo_cmp"]])
        d["letra"] = m.group(1) if m else ""
    return d


def revisar(d, interactivo=True):
    """
    Muestra lo que se leyo y deja corregirlo.
    Devuelve el dict corregido, o None si se decide saltear la factura.
    """
    _recalcular(d)
    while True:
        print()
        print("=" * 68)
        print(" REVISA LOS DATOS DE LA FACTURA")
        print(" Archivo: %s" % os.path.basename(d.get("archivo", "")))
        origenes = {
            "QR": "el codigo QR de AFIP (confiable)",
            "texto": "el texto del PDF",
            "OCR": "OCR de un escaneo (REVISA CADA NUMERO)",
            "manual": "lo que cargaste a mano",
        }
        if d.get("origen"):
            print(" Datos tomados de: %s"
                  % origenes.get(d["origen"], d["origen"]))
        print("=" * 68)
        for i, (clave, etiqueta, tipo) in enumerate(CAMPOS_REVISION, 1):
            print("  %2d) %-32s %s" % (i, etiqueta, _mostrar(d.get(clave), tipo)))
        print("-" * 68)

        print("       Neto + IVA + No grav. + Exento + Percep. = %s" %
              _mostrar(d.get("suma_componentes"), "importe"))
        print("       Total de la factura                      = %s" %
              _mostrar(d.get("total"), "importe"))
        if d.get("cuadra"):
            print("  [OK] Cuadra: la suma da exactamente el total.")
        else:
            print("       Diferencia                               = %s" %
                  _mostrar(round((d.get("total") or 0) - (d.get("suma_componentes") or 0), 2),
                           "importe"))
            print("  [!!] LOS IMPORTES NO CUADRAN.")

        for aviso in d.get("avisos", []):
            print("  [!]  %s" % aviso)

        print("-" * 68)
        if not interactivo:
            return d
        print("  ENTER = esta bien, cargala")
        print("  numero = corregir ese campo      S = saltear esta factura")
        opcion = input("  > ").strip()

        if opcion == "":
            if not d.get("cuadra"):
                conf = input("  Los importes no cuadran. La cargo igual? (s/N): ")
                if conf.strip().lower() not in ("s", "si", "sí"):
                    continue
            return d
        if opcion.lower() in ("s", "saltear", "n"):
            return None
        if opcion.isdigit() and 1 <= int(opcion) <= len(CAMPOS_REVISION):
            clave, etiqueta, tipo = CAMPOS_REVISION[int(opcion) - 1]
            d[clave] = _pedir(etiqueta, tipo, d.get(clave))
            if clave == "cuit_emisor":
                # si cambiaste el CUIT, el nombre se rellena solo
                autocompletar_nombre(d)
            d["avisos"] = []
            _recalcular(d)
            continue
        print("  No entendi. Probá de nuevo.")


def cargar_a_mano(path, nombres_por_cuit=None, cuit_receptor=None):
    """Para PDF escaneados: se piden los datos uno por uno."""
    print()
    print("=" * 68)
    print(" CARGA MANUAL")
    print(" Archivo: %s" % os.path.basename(path))
    print("=" * 68)
    print(" No pude leer este PDF solo. Copiame los datos de la factura.")
    print(" (ENTER deja el valor que esta entre corchetes)")
    print()
    print(" Moneda y tipo de cambio no se preguntan: quedan en $ y 1.")
    print(" Si la factura vino en dolares, corregilos al final,")
    print(" en la pantalla de revision (campos 8 y 9).")
    print()

    d = {
        "archivo": path, "ok": True, "escaneo": True, "origen": "manual",
        "avisos": ["Cargada a mano."], "razon_social": None, "cuit_emisor": None,
        "tipo_cmp": 1, "pto_venta": None, "nro_cmp": None, "cae": None,
        "fecha": None, "cuit_receptor": cuit_receptor, "moneda": "$",
        "cotizacion": 1, "neto_gravado": 0.0, "iva_105": 0.0, "iva_21": 0.0,
        "iva_27": 0.0, "no_gravado": 0.0, "exento": 0.0, "otros_tributos": 0.0,
        "total": None,
    }
    # los campos de plata se van sumando a la vista, para comparar con la factura
    campos_suma = ("neto_gravado", "iva_105", "iva_21", "iva_27",
                   "no_gravado", "exento", "otros_tributos")

    # se pide primero el CUIT: si el proveedor ya es conocido, el nombre
    # se completa solo y queda escrito igual que en las cargas anteriores
    orden = sorted(CAMPOS_REVISION,
                   key=lambda c: {"cuit_emisor": 0, "razon_social": 1}.get(c[0], 2))

    for clave, etiqueta, tipo in orden:
        if clave in NO_SE_PREGUNTAN_A_MANO:
            continue      # moneda y tipo de cambio quedan como estan

        if clave == "total":
            # el total viene precargado con la suma de todo lo anterior:
            # si coincide con la factura, alcanza con dar ENTER
            _recalcular(d)
            d["total"] = d["suma_componentes"]
            print()
            print("   Suma de todo lo que cargaste: %s"
                  % _mostrar(d["suma_componentes"], "importe"))
            print("   Si el total de la factura es ese, dale ENTER.")

        d[clave] = _pedir(etiqueta, tipo, d.get(clave))

        if clave == "cuit_emisor":
            # el nombre se completa solo apenas cargas el CUIT
            autocompletar_nombre(d, nombres_por_cuit)

        if clave in campos_suma:
            _recalcular(d)
            print("      (suma hasta aca: %s)"
                  % _mostrar(d["suma_componentes"], "importe"))

    return revisar(_recalcular(d))


def _asegurar_cliente(d, interactivo=True):
    """
    Si no se pudo leer a nombre de cual de las empresas esta la factura,
    se pregunta en vez de cortar el proceso.
    """
    cuit = d.get("cuit_receptor")
    try:
        if cuit is not None and int(cuit) in EMPRESAS_CLIENTE:
            return d
    except (TypeError, ValueError):
        pass

    if cuit:
        print("  [!] El CUIT del cliente que leí (%s) no es ninguna de tus "
              "empresas." % cuit)
    else:
        print("  [!] No pude leer a nombre de cual de tus empresas esta la factura.")

    if not interactivo:
        return d
    d["cuit_receptor"] = _elegir_empresa(cuit)
    return d


def procesar_pdf(path, nombres_por_cuit=None, interactivo=True):
    """
    Lee un PDF y devuelve el dict ya revisado, o None si se saltea.
    Es el punto de entrada que usa procesar_migracion.py.
    """
    d = leer_factura_pdf(path, nombres_por_cuit)
    if not d.get("ok"):
        for aviso in d.get("avisos", []):
            print("  [!] %s" % aviso)
        if not interactivo:
            return None
        if d.get("escaneo"):
            return cargar_a_mano(path, nombres_por_cuit)
        if d.get("incoherencias") and d.get("total"):
            print("  Hay importes que no cierran. Te la muestro para revisar.")
        else:
            print("  No pude leer algunos datos. Los completas a mano?")
        if input("  (S/n): ").strip().lower() in ("", "s", "si", "sí"):
            base = leer_factura_pdf(path, nombres_por_cuit)
            base.setdefault("moneda", "$")
            _asegurar_cliente(base)
            return revisar(_recalcular(base))
        return None
    _asegurar_cliente(d, interactivo=interactivo)
    return revisar(d, interactivo=interactivo)


if __name__ == "__main__":
    faltan = dependencias_faltantes()
    if faltan:
        sys.exit("Faltan librerias: %s\n"
                 "Instalalas con:  pip install pypdf pdfplumber opencv-python"
                 % ", ".join(faltan))
    if len(sys.argv) < 2:
        sys.exit("Uso: python leer_pdf.py factura.pdf")
    for arg in sys.argv[1:]:
        datos = leer_factura_pdf(arg)
        print(json.dumps(
            {k: (v.isoformat() if isinstance(v, datetime.date) else v)
             for k, v in datos.items()},
            indent=2, ensure_ascii=False))
