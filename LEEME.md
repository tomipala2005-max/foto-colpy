# Procesador de Migración de Facturas → Colppy

Automatiza la carga a la planilla de migración de Colppy, para **MEPANO S.A.** y
**Metalúrgica Pablo Nogués SRL**, a partir de:

- el Excel **"Mis Comprobantes Recibidos"** de ARCA/AFIP, o
- **facturas sueltas en PDF**.

Los dos caminos pasan por el mismo circuito: control de duplicados, conversión de
moneda, reglas de IIBB, cuenta de gasto y el `.xls` con el formato original.

---

## Instalación (una sola vez)

1. **Python 3** — https://www.python.org/downloads/
   Al instalar, tildá **"Add Python to PATH"**.

2. **Librerías** — abrí la consola (tecla Windows → escribí `cmd` → Enter) y pegá:

   ```
   pip install openpyxl xlrd==2.0.1 xlwt xlutils pypdf pdfplumber opencv-python-headless pypdfium2 pyzbar
   ```

   > La versión de `xlrd` tiene que ser **2.0.1** o anterior. Las más nuevas
   > no leen archivos `.xls`.
   >
   > De `pypdf` en adelante son las que leen los PDF. Si solo vas a usar el
   > Excel de ARCA, podés no instalarlas.

4. **OCR (opcional)** — solo si tenés facturas escaneadas, que son una foto sin
   texto adentro. Instalá **Tesseract** desde
   https://github.com/UB-Mannheim/tesseract/wiki, tildando el idioma
   **Spanish** durante la instalación.

   La librería `pytesseract` ya la instalan los `.bat`. **No hace falta que
   agregues Tesseract al PATH**: el script lo busca solo en las carpetas
   habituales de Windows.

   Para saber si quedó activo, ejecutá `1 - INSTALAR (una sola vez).bat`: el
   paso 4 te dice `OCR para PDF escaneados: ACTIVADO` o el motivo exacto por el
   que no.

   > Si lo instalaste en una carpeta rara, agregá la ruta a `RUTAS_TESSERACT`,
   > al principio de `leer_pdf.py`.
   >
   > Sin esto igual funciona: si el PDF es un escaneo, te pide los datos a mano.

3. **LibreOffice** — https://www.libreoffice.org/
   Se usa para recalcular las fórmulas de la plantilla. No hace falta abrirlo.

---

## Uso

Abrí la consola en esta carpeta y ejecutá:

```
python procesar_migracion.py "Mis Comprobantes Recibidos - CUIT 30710230184.xlsx"
```

Varios archivos de una vez:

```
python procesar_migracion.py archivo1.xlsx archivo2.xlsx
```

Facturas en PDF (una o varias; todas van a la misma planilla):

```
python procesar_migracion.py factura1.pdf factura2.pdf
```

Prueba en seco (genera los archivos pero **no** marca las facturas como cargadas):

```
python procesar_migracion.py archivo.xlsx --dry-run
```

**No hace falta aclarar la empresa**: se detecta sola por el CUIT del receptor.

### Las carpetas

| Carpeta | Para qué |
|---|---|
| `comprobantes arca/` | Dejá acá los `.xlsx` que bajás de ARCA |
| `facturas pdf/` | Dejá acá las facturas sueltas en PDF |
| `migraciones generadas/` | Acá salen los `.xls` listos para revisar |

Ejecutás `2 - PROCESAR FACTURAS.bat` sin arrastrar nada y encuentra solo todo lo
que haya en las dos carpetas de entrada. Ya no busca en Descargas.

Las tres se crean solas si no existen.

### Lo más fácil: arrastrar
También podés agarrar el archivo (Excel de ARCA o PDF) y soltarlo **encima** de
`2 - PROCESAR FACTURAS.bat`, esté donde esté. Podés arrastrar varios PDF juntos.

### Atajo para abrir la consola acá
Abrí esta carpeta en el Explorador, hacé clic en la barra de dirección, escribí
`cmd` y Enter.

---

## Qué genera

| Archivo | Contenido |
|---|---|
| `migraciones generadas/migracion factura <EMPRESA> <dd-mm-aaaa>.xls` | Excel con formato **idéntico** al original |
| `json arch/_registro_*.json` | Control de facturas ya cargadas |
| `json arch/_proveedores.json` | Nombres de proveedor por CUIT |

Si corrés dos veces el mismo día, el segundo archivo sale como `... v2.xls`.

### El CSV lo generás vos
El script deja **solo el Excel**, para que puedas revisarlo antes de importar.
Cuando esté OK: abrí el archivo, parate en la hoja **PASO 2**, y
`Archivo → Guardar como → CSV (delimitado por comas)`.

> **Ojo con la ruta larga**: si Excel dice *"la ruta de acceso es mayor que 259
> caracteres"*, copiá el archivo a una carpeta más corta (Escritorio o `C:\`) y
> abrilo desde ahí.

---

## Facturas en PDF

### De dónde saca los datos

| Dato | De dónde sale |
|---|---|
| Fecha, CUIT, punto de venta, número, tipo, total, moneda, CAE | **Código QR de AFIP** |
| Neto gravado, IVA por alícuota, no gravado, exento, percepciones | Texto del PDF |
| Razón social del proveedor | Registro de proveedores por CUIT, o el texto del PDF |

El QR que lleva impreso todo comprobante electrónico trae los datos **firmados
por AFIP**, así que esa parte no se puede leer mal. El desglose impositivo, en
cambio, se lee del texto, y ahí cada proveedor arma el PDF a su manera — por eso
siempre se revisa antes de cargar.

### Si no aparece el QR

No se rinde: saca la cabecera del texto de la factura.

- **Los CUIT**: busca todos los que haya y verifica el dígito verificador, así no
  confunde un CUIT con cualquier número de 11 cifras. El que coincide con MEPANO
  o MPN es el cliente; el otro es el proveedor. Con eso ya sabe de qué empresa es
  la factura sin preguntarte.
- **El número**: entiende `0003-00108864`, `Comp. Nro 00108864`, y el formato con
  punto de venta y número en campos separados.
- **La fecha**: la de emisión. Descarta a propósito las líneas de vencimiento,
  período y Vto. del CAE.
- **El tipo**: Factura / Nota de Crédito / Nota de Débito, letra A, B, C, E o M,
  y también el código numérico (`Cód. 01`, `Cód. 06`...).
- **La moneda**: si la factura está en dólares, lo detecta y busca la cotización.

Te avisa que los datos salieron del texto y no del QR, para que los mires con más
cuidado.

### Pantalla de revisión

Antes de escribir nada, muestra la factura completa:

```
   1) Proveedor                        MARMAQ S A
   2) CUIT del proveedor               30677768769
   ...
  17) TOTAL                            280.999,51
--------------------------------------------------------------------
  [OK] Los importes suman exactamente el total.
--------------------------------------------------------------------
  ENTER = esta bien, cargala
  numero = corregir ese campo      S = saltear esta factura
```

- **ENTER** → la carga.
- **Un número** (ej. `10`) → corrige ese campo y vuelve a mostrar todo.
- **S** → saltea esa factura y sigue con la siguiente.

Abajo de la lista siempre muestra las dos cifras enfrentadas:

```
       Neto + IVA + No grav. + Exento + Percep. = 280.999,51
       Total de la factura                      = 280.999,51
  [OK] Cuadra: la suma da exactamente el total.
```

Si no coinciden, muestra la diferencia, avisa con `[!!]` y pide confirmación
extra antes de cargarla. **Nunca inventa un número**: si no lo pudo leer, lo deja
en cero y te lo marca.

### Si el PDF es un escaneo o una foto
Prueba, en este orden:

1. **Busca el QR igual**, dibujando la hoja en alta resolución y recortándola en
   pedazos. Tiene un tope de 12 segundos: los **tiques de controlador fiscal**
   (estaciones de servicio y similares) no llevan QR, y no tiene sentido
   buscarlo eternamente.
2. **OCR**, si tenés Tesseract instalado (punto 4 de la instalación).
3. **Te pide los datos a mano.** No adivina nada.

#### Cómo hace el OCR para no equivocarse

El OCR sobre la foto de un tique arrugado se equivoca seguido. Tres defensas:

- **Usa la imagen original del PDF**, no una versión reducida. Suele tener
  bastante más resolución, y ahí está la diferencia entre leer un CUIT y no
  leerlo.
- **Prueba varias pasadas** (distintos filtros y modos) y se queda con la que
  **cuadra**: aquella donde neto + IVA + no gravado + percepciones da exacto el
  total. Si una pasada da exacto, casi con seguridad leyó bien.
- **Junta los CUIT de todas las pasadas** y los filtra por dígito verificador.
  Alcanza con que una sola pasada haya leído bien el CUIT.

Corrige solo dos errores clásicos del OCR:

- La coma decimal que se come (`72.122,89` leído como `7212289`): lo detecta
  porque las partes no suman el total, prueba dividir por 100 y se queda con la
  combinación que da exacto.
- El espacio que mete después de la coma (`16.504, 31` partido en dos números).

Y entiende las etiquetas de los controladores fiscales, que no dicen "IVA" sino
**`ALICUOTA 21,00%`**.

#### Si aun así falta un importe, lo deduce

Si el total se leyó bien y hay **una sola alícuota** en juego (21%, 10,5% o 27%),
el neto y el IVA salen por cuenta propia:

```
neto = (total - lo no gravado) / (1 + alícuota)
IVA  = neto x alícuota
```

Así que alcanza con que haya leído el total y reconocido la alícuota para
completar la fila. Te avisa siempre que dedujo algo.

**Cuándo NO lo hace**, a propósito, para no tapar un error:

- Si las partes ya cuadran, no toca nada.
- Si hay **más de una alícuota**, no se puede deducir sin inventar.
- Si la factura **nombra** conceptos no gravados o exentos pero no se pudo leer
  cuánto son. Deducir ahí mandaría todo a gravado, que es justo lo que no hay
  que hacer en seguros y servicios públicos.
- Si un importe ya se había leído y la cuenta da algo **muy distinto** (más de
  2% de diferencia), prefiere avisarte antes que pisarlo.
- Si el resultado no cuadra exacto contra el total, lo descarta.

En cambio sí corrige diferencias chicas: si leyó `21.000,60` donde iba
`21.000,00`, la cuenta contra el total lo endereza.

#### Qué tan bien anda (medido sobre tus 20 facturas)

| | |
|---|---|
| Comprobante + CUIT del proveedor + total | **18 de 18** |
| Además cuadra todo el desglose (neto, IVA, percepciones) | **15 de 18** |
| Entra sin pedirte nada a mano | **17 de 18** |

Las tres que no cuadran el desglose son un servicio público (Edenor), un seguro
(Sancor) y un escaneo. En las tres el proveedor, el comprobante y el total salen
bien: lo que te pide completar es el neto y el IVA.

**Verificación aparte**: se apartaron 10 facturas que no se usaron para ajustar
ninguna regla, y se comprobó contra ellas. Resultado: 9 de 10 cuadran, los
totales, CUIT y números coinciden exactamente con el QR firmado por AFIP, y en
todas el neto y el IVA están **escritos tal cual en el PDF** — no son cuentas
que haya hecho el script. La décima es un escaneo: trae bien el proveedor, el
comprobante y el total, y te pide el neto y el IVA.

**Entiende tablas de importes con los títulos arriba y los números alineados
debajo**, que es como las arman varios proveedores:

```
    Bruto         Descuento        Iva 21%      Importe Total
   11564.82        3469.45         1700.03         9795.40
```

Y también los importes escritos con un espacio entre cada dígito
(`2 7 2 , 7 1 0 . 7 4`), que algunos sistemas imprimen así.

#### Avisa cuando el PDF no es una factura
Si detecta un cupón de pago, un resumen de cuenta, un presupuesto, un remito o
una orden de compra, te lo dice:

```
[!] OJO: esto parece un cupon de pago, no una factura.
    Si no va al libro de compras, salteala con la S.
```

#### Antes de escribir un número raro, te lo deja para revisar

Un importe puede cerrar la cuenta y aun así no tener sentido. Un "no gravado" de
$40 en una factura de $120.000 casi nunca es un concepto real: es un número que
quedó pegado de otra parte de la hoja. Antes de mandarlo a Colppy, lo saca y te
lo marca.

Se marcan tres cosas:

- **Importes ridículos** para el tamaño de la factura (menos de $50, o menos del
  0,1% del total).
- **Importes negativos** en una factura que no es nota de crédito. En una nota de
  crédito sí valen, y no los toca.
- **Un IVA que no guarda relación con el neto**: si el neto es $100.000 y el IVA
  21% da $5.000, te avisa que debería andar por $21.000.

En los tres casos la factura queda marcada para que la mires, con el campo en
cero para que lo completes vos.

> Si te queda corto el piso de $50, cambiá `IMPORTE_MINIMO_CREIBLE` al principio
> de `leer_pdf.py`.

#### Qué esperar de un tique fotografiado
Probado sobre un tique de estación de servicio arrugado: saca bien el
**comprobante, el CUIT del proveedor, la fecha, el total y todo el desglose**.
Lo que suele salir mal es el **nombre del proveedor** (queda con letras
cambiadas). Lo corregís en la pantalla de revisión y **queda guardado**: la
próxima factura de ese CUIT ya viene con el nombre bien.

A medida que cargás los importes te va mostrando el subtotal acumulado
(`suma hasta aca: ...`), así lo comparás contra la factura mientras escribís.
Cuando llega al **TOTAL**, te lo ofrece ya calculado con todo lo que cargaste:
si es el mismo que el de la factura, alcanza con dar ENTER.

**Moneda y tipo de cambio no te los pregunta**: quedan en `$` y `1`, que es lo
que pasa casi siempre. Si la factura vino en dólares, los corregís al final desde
la pantalla de revisión (campos 8 y 9) y el script hace la conversión a pesos
como con cualquier otra factura.

### Servicios públicos (AySA)
Las liquidaciones de AySA ("LSP - Liquidación de Servicios Públicos A17") no
traen el QR de AFIP ni el formato de factura común. El script las reconoce igual:

- **Número**: `0106A11795675` → punto de venta 0106, número 11795675.
- **Desglose**: sale de la tabla *Tasas e Impuestos*. El monto base del IVA es el
  neto gravado (ya con el descuento ERAS), el IVA va al 27%, la **Perc. IVA** va
  a *Percepción de IVA* y las tasas **Financiamiento ERAS/APLA** a *No Gravado*.
  Nada de eso va a Percepción de IIBB.
- **Total**: el *Total a debitar* de la liquidación. La deuda anterior y el
  estado de cuenta de las hojas siguientes **no** se suman.
- **Tipo en Colppy**: `FAC` letra `A` (la hoja *Tablas* de la plantilla no trae
  el tipo 17; el script lo completa solo).

Tarda unos 20 segundos por PDF porque primero busca el QR de AFIP, que AySA no trae.

### Elegir la empresa
Si no puede leer a nombre de cuál de las dos empresas está la factura (o lee un
CUIT que no es ninguna de las tuyas), no corta el proceso: te lo pregunta.

```
      De que empresa es esta factura?
        1) MEPANO       MEPANO SOCIEDAD ANONIMA
        2) MPN          METALURGICA PABLO NOGUES SRL
```

Elegís `1` o `2` y sigue. También acepta que escribas `MPN` o el CUIT entero.

### El nombre se completa solo con el CUIT
Apenas cargás el CUIT del proveedor, el nombre aparece solo:

```
   CUIT del proveedor [(vacio)]: 30-67776876-9
      -> Proveedor conocido: MARMAQ S A
   Proveedor [MARMAQ S A]: _
```

Le das ENTER y queda. Sale del registro `json arch/_proveedores.json`, así que
queda escrito **igual que en las cargas anteriores**. Funciona en los dos lados:
en la carga manual (que por eso pide el CUIT antes que el nombre) y también si
corregís el CUIT desde la pantalla de revisión.

Si el CUIT no está en el registro, deja el nombre que pudo leer del PDF y lo
podés corregir a mano.

### Registro de proveedores
Cada vez que procesás un Excel de ARCA, guarda los nombres de proveedor por CUIT
en `json arch/_proveedores.json`. Así, cuando cargás un PDF de un proveedor
conocido, el nombre queda escrito **igual que en las cargas anteriores**.

---

## Reglas que aplica

### 1. No duplica facturas
Lleva registro de todo lo ya cargado (CUIT + número de factura). Si volvés a
pasar un archivo con facturas viejas, las saltea y procesa solo las nuevas.

> **Si el Excel generado no te sirve y lo vas a tirar**, ejecutá
> `3 - DESHACER ULTIMA CARGA.bat`. Las facturas quedan marcadas como cargadas
> apenas se genera el archivo, así que sin este paso la próxima corrida las
> saltea y las perdés. No borra nada: solo las desmarca para que las puedas
> volver a procesar.
>
> Borrar el `.xls` después de importarlo a Colppy está perfecto — el registro
> vive en `json arch`, no en el Excel.

### 2. Conversión de moneda (solo PASO 1)
Toda fila en dólares se multiplica por el tipo de cambio y queda en pesos, con
`Tipo Cambio = 1` y `Moneda = $`. En PASO 1 no quedan saldos en dólares.
**En PASO 2 no se multiplica nada.**

### 3. Fecha de vencimiento
`Fecha de Vencimiento = Fecha Factura + 30 días`

### 4. Cuenta de gasto
Se busca por CUIT en `_recursos/CODIGOS_A_USAR.xls`
(hoja **MSA** para MEPANO, hoja **MPN** para Pablo Nogués).
Si el proveedor no está, usa **Compra de Materias Primas** y te lo avisa al final
para que lo revises.

### 5. Percepción de IIBB — en este orden

| Caso | Qué hace |
|---|---|
| **Factura B/C sin desglose** | La plantilla duplica el total en No Gravado y en IIBB. No es percepción real → se anula. |
| **Estación de servicio** | → Neto No Gravado, **sin jurisdicción** (OPESA y cía.) |
| **Autopistas y peajes** | → Neto No Gravado (Autopistas del Sol, Autopistas Urbanas, Rutas Sur, Concesionario del Oeste) |
| **Cencosud** | → Neto No Gravado |
| **Sancor Seguros** | → Neto No Gravado, sin IIBB (no cobra IIBB) |
| **CEAMSE** | → Neto No Gravado, sin IIBB (no cobra IIBB) |
| **Menor a $30** | Resto de redondeo del dólar → se pliega en Neto Gravado (negativo resta, positivo suma) |
| **Percepción real** | Queda con Jurisdicción "Buenos Aires" y **te avisa para que confirmes** si va CABA |

### 6. Limpieza de PASO 2
Las filas debajo de la última factura quedan totalmente vacías: sin ceros ni
fechas residuales.

### 7. Formato idéntico al original
Se conservan dimensiones, anchos de columna, celdas combinadas, formatos de
número y alturas de fila. El script lo **verifica solo** en cada corrida y avisa
si algo no coincide.

---

## Controles automáticos

Al final de cada corrida te muestra:

- **Conversiones de moneda** — qué facturas se pasaron de dólares y a qué cambio
- **Reglas de IIBB aplicadas** — fila por fila, qué se hizo y por qué
- **Proveedores sin código** — para agregar a `CODIGOS_A_USAR.xls`
- **IIBB a confirmar** — percepciones reales donde hay que verificar la jurisdicción
- **Cuadre de totales** — que `Total = IVA + Neto Gravado + No Gravado + Percepciones`
- **Verificación de formato** — que el archivo salga igual al original

---

## Mantenimiento

### Agregar un proveedor nuevo con código
Lo mejor es cargarlo en `_recursos/CODIGOS_A_USAR.xls` (hoja MSA o MPN).
Si preferís, también podés agregarlo dentro del script, en `CUENTAS_EXTRA_MSA`
o `CUENTAS_EXTRA_MPN`:

```python
CUENTAS_EXTRA_MSA = {
    30717463907: 511200,
    12345678901: 522102,   # <- nuevo
}
```

### Agregar un proveedor que no cobra IIBB
Buscá la lista que corresponda cerca del principio del script y sumá una parte
del nombre, en MAYÚSCULAS:

```python
NOMBRES_AUTOPISTA = ["AUTOPISTA", "RUTAS SUR", "CONCESIONARIO DEL OESTE"]
```

Si es un caso nuevo (ni autopista, ni Sancor, ni CEAMSE), copiá el bloque de
CEAMSE dentro de `aplicar_reglas()` y cambiale el nombre.

### Un proveedor cuyo PDF no cuadra siempre
Si un proveedor usa una etiqueta rara para sus importes (por ejemplo
"Base Imponible" en vez de "Neto Gravado"), agregala en `leer_pdf.py`, en
`REGLAS_IMPORTES`, sobre la clave que corresponda:

```python
("neto_gravado", [r"IMPORTES?\s+GRAVADOS?", r"NETO\s+GRAVADO",
                  r"BASE\s+IMPONIBLE"]),   # <- nueva
```

> El orden importa: `no_gravado` tiene que quedar **antes** que `neto_gravado`,
> si no "NO GRAVADO" matchea con la regla de "GRAVADO".

### Cambiar los días de vencimiento

```python
DIAS_VENCIMIENTO = 30
```

### Actualizar la plantilla o los códigos
Reemplazá los archivos dentro de `_recursos/` manteniendo el mismo nombre.

---

## Estructura de la carpeta

```
Migracion/
├── 1 - INSTALAR (una sola vez).bat
├── 2 - PROCESAR FACTURAS.bat   <- el de todos los días
├── 3 - DESHACER ULTIMA CARGA.bat
├── 4 - RESPALDAR REGISTROS.bat
├── procesar_migracion.py       <- el script principal
├── leer_pdf.py                 <- el lector de facturas en PDF
├── LEEME.md                    <- este archivo
├── _recursos/
│   ├── plantilla_migracion.xls
│   └── CODIGOS_A_USAR.xls
├── comprobantes arca/          <- ENTRADA: los .xlsx de ARCA
├── facturas pdf/               <- ENTRADA: las facturas en PDF
├── migraciones generadas/      <- SALIDA: los .xls para revisar
└── json arch/                  <- registros de control
```

> **No borres la carpeta `json arch`.** Ahí está el control de qué facturas ya
> se cargaron. Si se pierde, el script no puede detectar duplicados y vas a
> tener que revisarlos a mano.

### Copias de seguridad
Cada vez que procesás facturas, **antes de tocar nada**, se copia todo lo que hay
en `json arch` a `json arch/respaldo/<fecha hora>/`. Se conservan las últimas 20.

Para hacer una copia a mano en cualquier momento:
`4 - RESPALDAR REGISTROS.bat`.

Si alguna vez se arruina un registro, cerrá todo, entrá a `json arch/respaldo/`,
elegí la carpeta de la fecha que te sirva y copiá esos `.json` de vuelta a
`json arch/`, pisando los que están.

> Si podés, copiá `json arch` de vez en cuando a Dropbox o a un pendrive. El
> respaldo automático te cubre de un error del script, no de que se rompa el
> disco.

---

## Problemas comunes

**"No encontré LibreOffice"** — Instalalo. Si ya está, abrí el script y poné la
ruta completa a `soffice.exe` en la función `buscar_libreoffice()`.

**"Falta una dependencia"** — Corré el `pip install` del punto 2.

**"CUIT receptor desconocido"** — El archivo es de otra empresa. Agregala al
diccionario `EMPRESAS` del script.

**"el formato NO coincide con el original"** — No uses el archivo y avisá. Puede
ser que la plantilla de `_recursos/` haya cambiado.

**Excel no abre el archivo por la ruta larga** — Copialo a una carpeta más corta.

**"No encontré el código QR de AFIP"** — Los datos salen del texto del PDF, que
es menos confiable. Revisá con más cuidado la pantalla antes de dar ENTER.

**Instalé Tesseract pero no lo usa** — Faltaba `pytesseract`, la librería de
Python que lo maneja (Tesseract solo no alcanza). Corré
`1 - INSTALAR (una sola vez).bat`: la instala y en el paso 4 te dice si el OCR
quedó activo o qué falta. Si dice que no encuentra `tesseract.exe`, fijate dónde
lo instalaste y agregá esa ruta a `RUTAS_TESSERACT` en `leer_pdf.py`.

**"El PDF es un escaneo"** — No tiene texto adentro y tampoco encontró el QR.
Opciones, de mejor a peor: pedile al proveedor el PDF original (no el escaneo);
volvé a escanear más derecho y a 300 dpi, cuidando que el QR salga entero y
nítido; instalá Tesseract para que lo intente con OCR; o cargalo a mano.

**Los importes no cuadran** — El proveedor usa etiquetas que el script no
reconoce. Corregí los campos con el número que corresponda; si ese proveedor se
repite seguido, avisame y le agrego la etiqueta a `REGLAS_IMPORTES` en
`leer_pdf.py`.
