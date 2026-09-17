"""
Validación Factura Global — Consolidación
==========================================

Toma los CSV descargados de SAP ('Gallo' FAGLL03 y 'Monivoi' ZLMXCOM) y arma
UN solo Excel con seis pestañas:

    1) 'Gallo'            → el CSV de Gallo + 3 columnas nuevas (K, L, M)
    2) 'Monivoi'          → el CSV de Monivoi tal cual
    3) 'Validación - Doc.'→ catálogo construido a partir de la Gallo
    4) 'TD'               → matrices (Tabla Dinámica) de suma de importe por
                            cedis × clase de documento (Facturas y Notas de Crédito)
    5) 'Venta DSD'        → matriz fecha × suma de importe, clases C1/C2/C5/C6/X0
                            y clase en blanco
    6) 'Venta EIAP'       → matriz fecha × suma de importe, clase RV

Pasos que replica (los mismos que se harían a mano en Excel):

    2. Col. C 'Referencia'  → primeros 2 caracteres → nueva col. K 'Referencia 2'
    3. Valores únicos de 'Referencia 2' → col. A de 'Validación - Doc.'
    4. Col. B 'Tipo': 'Global' si contiene número (o es PP/RG/SG),
       'Individual' si son puras letras.
    5-6. Filtrar Gallo col. D 'Asignación' == 'MX Commercial' y copiar
       col. D y E a las col. D y E de 'Validación - Doc.'
    7. Col. L 'Referencia 3'    = XLOOKUP(Referencia 2 → Tipo)      [calculado]
    8. Col. M 'Comercializadora'= XLOOKUP(Nº documento → Asignación) [calculado]
    9. Filtrar Gallo: Comercializadora=='MX Commercial' y Referencia 3=='Global'
       → copiar col. B 'Centro' y E 'Nº documento' a col. G y H del catálogo.
       IMPORTANTE: este bloque G-H se fija con la clasificación ORIGINAL, ANTES
       del override SG/RG de abajo (petición del stakeholder).
    Override SG/RG (POSTERIOR al Paso 9): para las filas con
       Comercializadora=='MX Commercial' y Referencia 2 ∈ {SG, RG} cuya
       Referencia 3 venía como 'Global', se reescribe a 'Individual'. Esto SÍ
       cambia la columna 'Referencia 3' de la hoja Gallo y la matriz 'TD' (que
       filtra Referencia 3=='Global'), pero NO el bloque G-H del Paso 9.

Matriz 'TD' (Tabla Dinámica):
    - Subconjunto base (Facturas y Notas): Referencia 3=='Global' y
      Asignación != 'MX Commercial'.
    - Filas: CADA matriz trae su PROPIA lista de cedis, derivada de SUS propios
      filtros (incluida su clase de documento) y ordenada. Sin recorte de rango:
      entra todo cedis que cumpla, no sólo 80AB..80VW. Por eso los tres bloques
      pueden mostrar distintos cedis y distinto número de filas.
    - Facturas: encabezados B5='C1', C5='C5'; cedis desde A6 hacia abajo.
    - Notas de Crédito: encabezados M5='C2', N5='C6'; cedis desde L6 hacia abajo.
    - Cada celda del cuerpo = suma de 'Importe en moneda local' de los documentos
      de ese cedis y esa clase.
    - Columna 'Grand Total' a la derecha de cada bloque (D en Facturas, O en
      Notas) = suma de las dos clases por cedis. Fila 'Grand Total' al final =
      suma por columna; la celda de cruce es el gran total del bloque.
    - Etiquetas de análisis para el stakeholder: 'Monitor', 'Contabilidad' y
      'Diferencias' en D1:D3 (Facturas) y O1:O3 (Notas). Sólo rótulos.
    - Tercera matriz 'PEP VOUCHERS' (columnas W/X): filtro Referencia 3=='Global',
      Referencia 2=='PP' y Clase de documento=='C2'. Cedis en W (desde W6),
      encabezado 'C2' en X5, con su fila 'Grand Total'. NO excluye 'MX Commercial'.

Pestañas 'Venta DSD' y 'Venta EIAP' (dos columnas: fecha e importe):
    - DOS filtros, ambos obligatorios: 'Clase de documento' Y 'Cuenta'. A
      diferencia de 'TD', NO exigen Referencia 3=='Global' ni excluyen
      'MX Commercial'; suman toda la hoja Gallo que cumpla ambos filtros.
          Venta DSD  → clase C1, C2, C5, C6, X0 o en blanco
                       Y cuenta 4000010 / 4000005
          Venta EIAP → clase RV
                       Y cuenta 4000010 / 4000020 / 4000802 / 4300013
    - Col. A: la fecha. Una fila por día, agrupando por la columna
      'Fe.contabilización' de la Gallo (el export trae 'dd/mm/aaaa hh:mm': la
      hora se recorta, sólo se agrupa por día). Las fechas que el usuario
      seleccionó en la GUI SIEMPRE aparecen, aunque ese día no traiga
      importes (0.00).
    - Col. B: suma de 'Importe en moneda local' de ese día.
    - Si hay filas con la fecha vacía, se muestran en un renglón 'Sin fecha'
      para que el 'Grand Total' siempre cuadre con el filtro (no se pierde
      dinero por el camino).
    - Si el CSV trae fechas FUERA de las seleccionadas, igual se muestran (son
      datos reales) y se reportan en el resumen para que el stakeholder lo vea.

NOTA: los cruces de los pasos 7 y 8 se calculan en Python y se pegan como
VALORES (no fórmulas vivas), para que el archivo abra rápido y sin #N/A. Las
matrices 'TD', 'Venta DSD' y 'Venta EIAP' también se pegan como valores
numéricos ya sumados.

Este módulo NO depende de SAP ni de tkinter: se puede probar solo con
`python Consolidacion.py` teniendo dos CSV a la mano.
"""

import os
import re
import unicodedata
from datetime import datetime

import pandas as pd
from openpyxl.styles import Font

# ----------------------------------------------------------------------
# Constantes del proceso
# ----------------------------------------------------------------------
VALOR_ASIGNACION = "MX Commercial"          # filtro de los pasos 5, 8.5 y 9
PREFIJOS_LETRAS_GLOBAL = {"PP", "RG", "SG"}  # letras que también son 'Global'
PREFIJOS_RECLASIFICAR = {"SG", "RG"}         # paso 8.5: SG/RG 'MX Commercial' → Individual

HOJA_GALLO      = "Gallo"
HOJA_MONIVOI    = "Monivoi"
HOJA_VALIDACION = "Validación - Doc."
HOJA_TD         = "TD"

# --- Layout de la matriz 'TD' ---
# Row 5 = headers, data starts at row 6. Facturas live on the left block
# (row labels in column A), Notas de Crédito on the right block (column L).
TD_FILA_ENCABEZADO = 5
TD_FILA_INICIO     = 6
# (clase de documento, índice de columna 1-based donde va su encabezado/suma)
# Cada bloque lleva su columna 'Grand Total' (D en Facturas, O en Notas).
TD_FACTURAS = {
    "titulo": "FACTURAS",
    "col_cedis": 1,                      # A
    "clases": [("C1", 2), ("C5", 3)],    # B / C
    "col_grand_total": 4,                # D
}
TD_NOTAS = {
    "titulo": "NOTAS DE CRÉDITO",
    "col_cedis": 12,                     # L
    "clases": [("C2", 13), ("C6", 14)],  # M / N
    "col_grand_total": 15,               # O
}
# Etiquetas de análisis para el stakeholder: van en la columna Grand Total,
# filas 1-3 (D1:D3 en Facturas, O1:O3 en Notas). Sólo rótulos.
TD_LABELS_ANALISIS = [(1, "Monitor"), (2, "Contabilidad"), (3, "Diferencias")]

# Tercera matriz: PEP vouchers. Filtro: Referencia 3=='Global',
# Referencia 2=='PP' y Clase de documento=='C2'. Una sola columna de valor (X),
# con las filas de cedis en W. No excluye 'MX Commercial' (sólo esos 3 filtros).
TD_PEP = {
    "titulo": "PEP VOUCHERS",
    "col_cedis": 23,           # W
    "col_valor": 24,           # X
    "encabezado_valor": "C2",  # clase del filtro (va en X5)
}
TD_PEP_REFERENCIA2 = "PP"
TD_PEP_CLASE = "C2"

TD_FORMATO_NUMERO = "#,##0.00"

# --- Pestañas 'Venta DSD' y 'Venta EIAP' ---
# Matriz de dos columnas: fecha (A) e importe sumado (B). Se filtra por DOS
# criterios a la vez (AND): clase de documento Y cuenta. La cadena vacía ""
# entre las clases representa el "(blanks)" de Excel, es decir la clase vacía
# o con puros espacios.
HOJA_VENTA_DSD  = "Venta DSD"
HOJA_VENTA_EIAP = "Venta EIAP"
VENTA_CLASES_DSD  = frozenset({"C1", "C2", "C5", "C6", "X0", ""})
VENTA_CLASES_EIAP = frozenset({"RV"})
# Cuentas de mayor. Se comparan NORMALIZADAS (sin ceros a la izquierda y sin
# el '.0' que a veces deja la conversión xlsx→csv), así que da igual si SAP
# las exporta como '4000010', '0004000010' o '4000010.0'.
VENTA_CUENTAS_DSD  = frozenset({"4000010", "4000005"})
VENTA_CUENTAS_EIAP = frozenset({"4000010", "4000020", "4000802", "4300013"})

VENTA_COL_FECHA_CANDIDATAS = (
    "Fe.contabilización",
    "Fecha contabilización",
    "Fecha de contabilización",
    "Fecha contab.",
)
VENTA_COL_CUENTA_CANDIDATAS = (
    "Cuentas",
    "Cuenta",
    "Cuenta de mayor",
    "Cuenta contable",
    "Cta.mayor",
    "Nº cuenta",
    "No. cuenta",
)
VENTA_ENCABEZADO_FECHA   = "Fecha"
VENTA_ENCABEZADO_IMPORTE = "Importe en moneda local"
VENTA_ETIQUETA_SIN_FECHA = "Sin fecha"
VENTA_ETIQUETA_TOTAL     = "Grand Total"
VENTA_FORMATO_FECHA_EXCEL = "DD.MM.YYYY"

# Formatos de fecha que puede traer el export de SAP. Se prueban primero los de
# año de 4 dígitos; los de 2 dígitos son el último recurso. La ambigüedad real
# (p. ej. 01/09/2026 = 1-sep o 9-ene) se resuelve contra las fechas que el
# usuario seleccionó en la GUI; si no se puede resolver, truena con un mensaje
# claro en vez de adivinar.
VENTA_FORMATOS_FECHA_4D = (
    "%d.%m.%Y", "%m.%d.%Y",
    "%d/%m/%Y", "%m/%d/%Y",
    "%Y-%m-%d", "%Y/%m/%d",
    "%d-%m-%Y", "%m-%d-%Y",
)
VENTA_FORMATOS_FECHA_2D = ("%d.%m.%y", "%m.%d.%y", "%d/%m/%y", "%m/%d/%y")
VENTA_ANIO_MIN = 1990
VENTA_ANIO_MAX = 2100

# El export de Gallo trae la fecha como 'dd/mm/aaaa hh:mm'. Este patrón se queda
# sólo con la parte de FECHA (grupo 1) y descarta lo que venga después de un
# espacio o de una 'T'. No intenta interpretar el orden día/mes: de eso se
# encarga _detectar_formato_fecha.
_RE_SOLO_FECHA = re.compile(r"^(\d{1,4}[./\-]\d{1,2}[./\-]\d{1,4})(?:[ T].*)?$")


# ----------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------
def carpeta_output_desde_input(ruta_input):
    """
    Devuelve la carpeta 'Output' hermana de la carpeta de descarga.

    Si la carpeta elegida se llama 'Input' (el layout estándar del proyecto:
    src/Input), el Output queda JUNTO a ella (src/Output). Si el stakeholder
    eligió cualquier otra carpeta (p. ej. una de SharePoint), el Output se
    crea DENTRO de esa carpeta para no regar archivos fuera de su elección.
    """
    ruta = os.path.normpath(ruta_input)
    if os.path.basename(ruta).lower() == "input":
        return os.path.join(os.path.dirname(ruta), "Output")
    return os.path.join(ruta, "Output")


def _leer_csv(ruta):
    """Lee un CSV descargado por la app: todo como texto, acentos incluidos."""
    if not os.path.exists(ruta):
        raise FileNotFoundError(f"Archivo no encontrado: {ruta}")
    # keep_default_na=False → las celdas vacías quedan como "" (no NaN),
    # así los .strip() y comparaciones de texto no truenan nunca.
    return pd.read_csv(ruta, dtype=str, encoding="utf-8-sig", keep_default_na=False)


def _es_sin_datos(df):
    """True si el CSV es el placeholder 'Sin datos' que genera la descarga."""
    return df.empty or (len(df.columns) == 1 and "Sin datos" in str(df.columns[0]))


def _col(df, nombre, indice, letra):
    """
    Devuelve el nombre REAL de una columna: primero la busca por nombre
    (p. ej. 'Referencia'); si no existe, cae a la posición esperada
    (p. ej. columna C = índice 2). Así aguantamos pequeños cambios de layout.
    """
    if nombre in df.columns:
        return nombre
    if indice < len(df.columns):
        return df.columns[indice]
    raise ValueError(
        f"No encontré la columna '{nombre}' (esperada en la columna {letra}) "
        f"en el archivo de Gallo. Columnas disponibles: {list(df.columns)}"
    )


def _col_por_nombre(df, nombre):
    """
    Resolve a column strictly by name (no positional fallback).

    Used for columns whose position in the SAP export is not fixed/known
    ('Clase de documento', 'Importe en moneda local'). If the header is not
    present we fail loud with a clear, actionable message instead of guessing a
    column index and silently summing the wrong data.
    """
    if nombre in df.columns:
        return nombre
    raise ValueError(
        f"No encontré la columna '{nombre}' en el archivo de Gallo. "
        f"Columnas disponibles: {list(df.columns)}"
    )


def _normalizar_encabezado(texto):
    """
    Normaliza un encabezado para compararlo sin acentos, puntos, espacios ni
    mayúsculas: 'Fe.Contabilización' y 'fecha contabilizacion' se vuelven
    comparables. Sólo se usa para reconocer VARIANTES DE ESCRITURA del mismo
    campo, nunca para adivinar una columna distinta.
    """
    sin_acentos = unicodedata.normalize("NFKD", str(texto))
    sin_acentos = "".join(c for c in sin_acentos if not unicodedata.combining(c))
    return "".join(c for c in sin_acentos.lower() if c.isalnum())


def _col_por_candidatos(df, candidatos, descripcion):
    """
    Resuelve una columna aceptando varias formas de escribir el MISMO campo
    ('Fe.contabilización' / 'Fecha contabilización', 'Cuentas' / 'Cuenta de
    mayor'). La comparación ignora acentos, puntos, espacios y mayúsculas.

    Si no está ninguna, truena con un mensaje accionable que lista las columnas
    reales: preferimos parar a filtrar/agrupar por una columna equivocada.
    """
    normalizadas = {_normalizar_encabezado(c): c for c in df.columns}
    for candidata in candidatos:
        real = normalizadas.get(_normalizar_encabezado(candidata))
        if real is not None:
            return real
    raise ValueError(
        f"No encontré la columna de {descripcion} en el archivo de Gallo "
        f"(busqué: {', '.join(candidatos)}). "
        f"Columnas disponibles: {list(df.columns)}"
    )


def _col_fecha_contabilizacion(df):
    """Columna de fecha de contabilización de la Gallo ('Fe.contabilización')."""
    return _col_por_candidatos(
        df, VENTA_COL_FECHA_CANDIDATAS, "fecha de contabilización"
    )


def _col_cuenta(df):
    """Columna de cuenta de mayor de la Gallo ('Cuentas')."""
    return _col_por_candidatos(df, VENTA_COL_CUENTA_CANDIDATAS, "cuenta")


def _normalizar_cuenta(valor):
    """
    Normaliza una cuenta de mayor para compararla sin sorpresas de formato:

        ' 0004000010 ' -> '4000010'
        '4000010.0'    -> '4000010'   (la conversión xlsx→csv la volvió float)
        '4000010'      -> '4000010'

    Sólo se quitan ceros a la IZQUIERDA y el '.0' decimal; nunca dígitos
    significativos. Un valor no numérico se devuelve tal cual (en mayúsculas),
    para que una cuenta con letras siga siendo comparable.
    """
    texto = str(valor).strip().upper().replace(" ", "")
    if not texto:
        return ""
    # '4000010.0' / '4000010.00' → '4000010' (sólo si los decimales son ceros)
    if re.fullmatch(r"\d+\.0*", texto):
        texto = texto.split(".", 1)[0]
    if texto.isdigit():
        return texto.lstrip("0") or "0"
    return texto


def _solo_fecha(valor):
    """
    Recorta la hora de un valor de fecha: el export de Gallo trae
    'dd/mm/aaaa hh:mm' y para agrupar por día la hora estorba.

        '01/09/2026 14:35' -> '01/09/2026'
        '2026-09-01T00:00' -> '2026-09-01'
        '01.09.2026'       -> '01.09.2026'

    Si el valor no empieza con algo que parezca fecha, se devuelve tal cual
    (ya recortado) para que el detector de formato truene mostrando el valor
    REAL en vez de uno mutilado por este helper.
    """
    texto = str(valor).strip()
    coincidencia = _RE_SOLO_FECHA.match(texto)
    return coincidencia.group(1) if coincidencia else texto


def _clasificar_tipo(referencia2):
    """
    Paso 4: 'Global' si contiene al menos un número, o si es PP/RG/SG.
    'Individual' para el resto (puras letras).
    """
    ref = str(referencia2).strip()
    if any(caracter.isdigit() for caracter in ref):
        return "Global"
    if ref.upper() in PREFIJOS_LETRAS_GLOBAL:
        return "Global"
    return "Individual"


def _parse_importe_us(valor):
    """
    Convert a US-formatted SAP amount string into a float.

    Handles the shapes SAP actually exports (US locale):
        - thousands separator ','  and decimal point '.'   -> '1,234.56'
        - SAP trailing minus (sign after the number)        -> '1,234.56-'
        - leading minus / stray leading '+'                 -> '-1,234.56'
        - accounting parentheses                            -> '(1,234.56)'
        - blank / whitespace / lone '-'                     -> 0.0

    A non-empty token that is still not numeric after cleanup raises ValueError,
    so corrupt amounts stop the run instead of quietly understating a sum.
    """
    if valor is None:
        return 0.0

    s = str(valor).strip()
    if s in ("", "-", "+"):
        return 0.0

    negativo = False

    # Accounting parentheses: (1,234.56) == -1,234.56
    if s.startswith("(") and s.endswith(")"):
        negativo = True
        s = s[1:-1].strip()

    # Sign can be leading or trailing (SAP puts it at the end).
    if s.endswith("-"):
        negativo = True
        s = s[:-1].strip()
    elif s.endswith("+"):
        s = s[:-1].strip()
    if s.startswith("-"):
        negativo = True
        s = s[1:].strip()
    elif s.startswith("+"):
        s = s[1:].strip()

    # US format: ',' is only a thousands separator, so drop it. Drop spaces too.
    s = s.replace(",", "").replace(" ", "")
    if s == "":
        return 0.0

    try:
        numero = float(s)
    except ValueError:
        raise ValueError(
            f"No pude convertir el importe '{valor}' a número en la columna "
            f"'Importe en moneda local'. ¿Formato inesperado? "
            f"(esperaba formato US, p. ej. 1,234.56 o 1,234.56-)."
        )

    return -numero if negativo else numero


# ----------------------------------------------------------------------
# Fechas de las pestañas 'Venta ...'
# ----------------------------------------------------------------------
def _parsear_fechas_seleccionadas(fechas):
    """
    Valida y normaliza las fechas que el usuario eligió en la GUI
    ('DD.MM.YYYY') a objetos date, sin duplicados y en orden cronológico.

    Devuelve [] si no se recibió ninguna (el módulo sigue siendo usable solo).
    """
    if not fechas:
        return []
    if isinstance(fechas, str):
        fechas = [fechas]

    convertidas = []
    for texto in fechas:
        try:
            convertidas.append(datetime.strptime(str(texto).strip(), "%d.%m.%Y").date())
        except ValueError:
            raise ValueError(
                f"La fecha seleccionada '{texto}' no tiene el formato "
                f"esperado 'DD.MM.YYYY'."
            )
    return sorted(set(convertidas))


def _intentar_formato(valores, formato):
    """
    Devuelve la lista de dates si TODOS los valores parsean con ese formato y
    caen en un año razonable; None si alguno falla.

    El chequeo de año no es cosmético: strptime acepta '26' para %Y y devolvería
    el año 26 d.C., lo que haría pasar por bueno un formato equivocado.
    """
    convertidas = []
    for valor in valores:
        try:
            fecha = datetime.strptime(valor, formato).date()
        except ValueError:
            return None
        if not (VENTA_ANIO_MIN <= fecha.year <= VENTA_ANIO_MAX):
            return None
        convertidas.append(fecha)
    return convertidas


def _detectar_formato_fecha(valores, fechas_referencia=()):
    """
    Detecta con qué formato viene la columna de fecha del CSV de Gallo.

    Regla: se prueban los formatos de año de 4 dígitos y, sólo si ninguno
    sirve, los de 2 dígitos. Si varios formatos parsean TODO:
      - si todos dan el MISMO resultado, la ambigüedad da igual → se usa el 1º;
      - si dan resultados distintos, se desempata con las fechas que el usuario
        seleccionó (gana el formato que deja más fechas dentro de la selección);
      - si sigue habiendo empate, se lanza ValueError. Nunca se adivina: un
        01/09/2026 mal interpretado movería el importe de día sin avisar.

    Devuelve el formato (str) o None si no hay ningún valor que parsear.
    """
    valores = [v for v in valores if v]
    if not valores:
        return None

    referencia = set(fechas_referencia or ())

    for tanda in (VENTA_FORMATOS_FECHA_4D, VENTA_FORMATOS_FECHA_2D):
        viables = []
        for formato in tanda:
            convertidas = _intentar_formato(valores, formato)
            if convertidas is not None:
                viables.append((formato, convertidas))

        if not viables:
            continue
        if len(viables) == 1:
            return viables[0][0]

        primera = viables[0][1]
        if all(convertidas == primera for _, convertidas in viables[1:]):
            return viables[0][0]   # mismo resultado: la ambigüedad es inocua

        if referencia:
            puntajes = [
                (sum(1 for f in convertidas if f in referencia), formato)
                for formato, convertidas in viables
            ]
            mejor = max(puntaje for puntaje, _ in puntajes)
            ganadores = [formato for puntaje, formato in puntajes if puntaje == mejor]
            if mejor > 0 and len(ganadores) == 1:
                return ganadores[0]

        raise ValueError(
            "No pude determinar sin ambigüedad el formato de la columna de "
            "fecha de contabilización de Gallo "
            f"(ejemplo: '{valores[0]}'; candidatos: "
            f"{', '.join(formato for formato, _ in viables)}). "
            "Revisa el formato de fecha del export de SAP."
        )

    raise ValueError(
        "No pude interpretar la columna de fecha de contabilización de Gallo "
        f"(ejemplo: '{valores[0]}'). Formatos soportados: "
        f"{', '.join(VENTA_FORMATOS_FECHA_4D + VENTA_FORMATOS_FECHA_2D)}."
    )


# ----------------------------------------------------------------------
# Proceso principal
# ----------------------------------------------------------------------
def consolidar_gallo_monivoi(
    ruta_gallo,
    ruta_monivoi,
    ruta_output,
    nombre_salida=None,
    callback_status=None,
    fechas_seleccionadas=None,
):
    """
    Ejecuta la consolidación completa (pasos 1 a 9 + reclasificación + matriz TD
    + pestañas 'Venta DSD' y 'Venta EIAP').

    Args:
        ruta_gallo   (str): Ruta al CSV descargado de Gallo (FAGLL03).
        ruta_monivoi (str): Ruta al CSV descargado de Monivoi (ZLMXCOM).
        ruta_output  (str): Carpeta donde se guardará el Excel consolidado.
        nombre_salida (str|None): Nombre del .xlsx final. Si None →
            'Consolidado_<nombre del CSV de Gallo>.xlsx'.
        callback_status (function|None): Para reportar avance a la GUI.
        fechas_seleccionadas (list[str]|None): Fechas 'DD.MM.YYYY' que el
            usuario eligió en la GUI (una en modo día, todas las del rango en
            modo rango). Garantizan que esos días aparezcan en 'Venta DSD' y
            'Venta EIAP' aunque no traigan importes, y desempatan el formato de
            fecha del CSV. Si es None, las filas salen sólo de los datos.

    Returns:
        dict: resumen del proceso →
            ruta, filas_gallo, filas_monivoi, refs_unicas,
            filas_catalogo, filas_globales, monivoi_sin_datos,
            filas_reclasificadas, cedis_facturas, cedis_notas, cedis_pep,
            total_facturas, total_notas, total_pep, total_importe_gallo,
            total_venta_dsd, total_venta_eiap, venta_avisos
    """

    def update_status(mensaje):
        if callback_status:
            callback_status(mensaje)

    # Se valida ANTES de leer nada: si la GUI manda una fecha con formato raro,
    # es mejor enterarse en el primer segundo que después de procesar 1M filas.
    fechas_sel = _parsear_fechas_seleccionadas(fechas_seleccionadas)

    # === Paso 1: leer ambos archivos ==================================
    update_status("📄 Leyendo archivo de Gallo...")
    gallo = _leer_csv(ruta_gallo)
    if _es_sin_datos(gallo):
        raise ValueError(
            "El archivo de Gallo no tiene datos (descarga 'sin partidas'). "
            "No hay nada que consolidar."
        )

    update_status("📄 Leyendo archivo de Monivoi...")
    monivoi = _leer_csv(ruta_monivoi)
    monivoi_sin_datos = _es_sin_datos(monivoi)

    # Columnas clave de la Gallo (por nombre, con respaldo por posición)
    col_centro = _col(gallo, "Centro",       1, "B")
    col_ref    = _col(gallo, "Referencia",   2, "C")
    col_asig   = _col(gallo, "Asignación",   3, "D")
    col_doc    = _col(gallo, "Nº documento", 4, "E")

    # Columnas que necesitan las matrices. Se resuelven ESTRICTAMENTE por
    # nombre (su posición en el export no es fija) y fallan pronto con un
    # error claro si no están, antes de procesar nada.
    col_clase   = _col_por_nombre(gallo, "Clase de documento")
    col_importe = _col_por_nombre(gallo, "Importe en moneda local")
    col_fecha   = _col_fecha_contabilizacion(gallo)
    col_cuenta  = _col_cuenta(gallo)

    # === Paso 2: 'Referencia 2' = primeros 2 caracteres de 'Referencia'
    # Se agrega AL FINAL de la tabla; con el layout estándar (A-J) eso la
    # deja exactamente en la columna K.
    update_status("🔧 Creando 'Referencia 2'...")
    gallo["Referencia 2"] = gallo[col_ref].astype(str).str.strip().str[:2]

    # === Paso 3: valores únicos de 'Referencia 2' (en orden de aparición)
    refs_unicas = list(dict.fromkeys(
        ref for ref in gallo["Referencia 2"] if str(ref).strip() != ""
    ))

    # === Paso 4: clasificar cada referencia como Global / Individual ===
    tipos = [_clasificar_tipo(ref) for ref in refs_unicas]
    mapa_tipo = dict(zip(refs_unicas, tipos))

    # === Pasos 5 y 6: catálogo de documentos 'MX Commercial' ===========
    update_status("🔎 Filtrando 'MX Commercial'...")
    mascara_mx = gallo[col_asig].astype(str).str.strip() == VALOR_ASIGNACION
    catalogo_docs = gallo.loc[mascara_mx, [col_asig, col_doc]].copy()

    # === Paso 7: 'Referencia 3' (Referencia 2 → Tipo), como VALORES ====
    update_status("🔁 Cruzando 'Referencia 3' y 'Comercializadora'...")
    gallo["Referencia 3"] = (
        gallo["Referencia 2"].map(mapa_tipo).fillna("")
    )

    # === Paso 8: 'Comercializadora' (Nº documento → Asignación) ========
    # Como XLOOKUP: se queda con la PRIMERA coincidencia del catálogo.
    mapa_comercializadora = {}
    for doc, asig in zip(
        catalogo_docs[col_doc].astype(str).str.strip(),
        catalogo_docs[col_asig].astype(str).str.strip(),
    ):
        mapa_comercializadora.setdefault(doc, asig)

    gallo["Comercializadora"] = (
        gallo[col_doc].astype(str).str.strip().map(mapa_comercializadora).fillna("")
    )

    # === Paso 9: filtro final → 'Centro' y 'Nº documento' globales =====
    # Este bloque (que alimenta G-H de 'Validación - Doc.') se fija con la
    # clasificación ORIGINAL, ANTES del override SG/RG de abajo. Es a propósito
    # (petición del stakeholder): los globales del Paso 9 conservan las filas
    # SG/RG 'MX Commercial' que venían como 'Global', aunque esas mismas filas
    # queden luego como 'Individual' en la hoja Gallo.
    update_status("🧮 Aplicando filtro final (Global + MX Commercial)...")
    mascara_final = (
        (gallo["Comercializadora"] == VALOR_ASIGNACION)
        & (gallo["Referencia 3"] == "Global")
    )
    globales = gallo.loc[mascara_final, [col_centro, col_doc]].copy()

    # === Override SG/RG → 'Individual' (POSTERIOR al Paso 9) ============
    # Business override: for MX Commercial rows whose Referencia 2 is SG or RG,
    # the amount is settled individually, so Referencia 3 becomes 'Individual'
    # instead of the default 'Global' that SG/RG get in _clasificar_tipo.
    # Applied AFTER the Paso 9 snapshot, so it does NOT touch the G-H block; it
    # DOES change 'Referencia 3' in the Gallo sheet and therefore the 'TD' matrix.
    update_status("🏷️ Reclasificando SG/RG de 'MX Commercial' a 'Individual'...")
    ref2_norm = gallo["Referencia 2"].astype(str).str.strip().str.upper()
    mascara_reclasificar = (
        (gallo["Comercializadora"] == VALOR_ASIGNACION)
        & (ref2_norm.isin(PREFIJOS_RECLASIFICAR))
        & (gallo["Referencia 3"] == "Global")
    )
    filas_reclasificadas = int(mascara_reclasificar.sum())
    gallo.loc[mascara_reclasificar, "Referencia 3"] = "Individual"

    # === Importe: UNA sola pasada de parseo para todas las matrices =====
    # Son ~1.15M filas en un rango de 3 días: parsear dos veces la misma
    # columna sería tirar tiempo, y además garantiza que 'TD' y las pestañas
    # 'Venta ...' vean exactamente los mismos números.
    update_status("🔢 Convirtiendo 'Importe en moneda local' a número...")
    importe_num = gallo[col_importe].map(_parse_importe_us)

    # === Matriz 'TD': suma de importe por cedis × clase de documento ===
    update_status("📊 Construyendo matriz 'TD'...")
    td_resumen = _preparar_matriz_td(gallo, col_centro, col_asig, col_clase, importe_num)
    total_importe_gallo = td_resumen["total_importe_gallo"]

    # === Matrices 'Venta DSD' y 'Venta EIAP': suma por fecha ===========
    update_status("📆 Construyendo 'Venta DSD' y 'Venta EIAP'...")
    ventas = _preparar_ventas(
        gallo, col_clase, col_cuenta, col_fecha, importe_num, fechas_sel
    )

    # === Exportar el Excel consolidado =================================
    update_status("💾 Guardando Excel consolidado...")
    os.makedirs(ruta_output, exist_ok=True)

    if not nombre_salida:
        base_gallo = os.path.splitext(os.path.basename(ruta_gallo))[0]
        nombre_salida = f"Consolidado_{base_gallo}.xlsx"
    ruta_final = os.path.join(ruta_output, nombre_salida)

    with pd.ExcelWriter(ruta_final, engine="openpyxl") as writer:
        gallo.to_excel(writer, sheet_name=HOJA_GALLO, index=False)
        monivoi.to_excel(writer, sheet_name=HOJA_MONIVOI, index=False)

        # --- Pestaña 'Validación - Doc.' (armada celda por celda porque
        #     tiene bloques independientes: A-B, D-E y G-H) ---
        ws = writer.book.create_sheet(HOJA_VALIDACION)

        encabezados = {
            "A1": "Referencia 2", "B1": "Tipo",
            "D1": "Asignación",   "E1": "Nº documento",
            "G1": "Centro",       "H1": "Nº documento",
        }
        for celda, texto in encabezados.items():
            ws[celda] = texto
            ws[celda].font = Font(bold=True)

        # Bloque A-B: catálogo de referencias y su tipo (pasos 3 y 4)
        for fila, (ref, tipo) in enumerate(zip(refs_unicas, tipos), start=2):
            ws.cell(row=fila, column=1, value=ref)
            ws.cell(row=fila, column=2, value=tipo)

        # Bloque D-E: documentos con Asignación 'MX Commercial' (paso 6)
        for fila, (asig, doc) in enumerate(
            zip(catalogo_docs[col_asig], catalogo_docs[col_doc]), start=2
        ):
            ws.cell(row=fila, column=4, value=asig)
            ws.cell(row=fila, column=5, value=doc)

        # Bloque G-H: Centro y Nº documento de los globales (paso 9)
        for fila, (centro, doc) in enumerate(
            zip(globales[col_centro], globales[col_doc]), start=2
        ):
            ws.cell(row=fila, column=7, value=centro)
            ws.cell(row=fila, column=8, value=doc)

        # Anchos cómodos para leer el catálogo
        for letra, ancho in {"A": 14, "B": 12, "D": 16, "E": 16,
                             "G": 10, "H": 16}.items():
            ws.column_dimensions[letra].width = ancho

        # --- Pestaña 'TD': matrices de Facturas y Notas de Crédito ---
        ws_td = writer.book.create_sheet(HOJA_TD)
        _escribir_matriz_td(ws_td, td_resumen)

        # --- Pestañas 'Venta DSD' y 'Venta EIAP' ---
        for nombre_hoja, clave in ((HOJA_VENTA_DSD, "dsd"), (HOJA_VENTA_EIAP, "eiap")):
            _escribir_hoja_venta(writer.book.create_sheet(nombre_hoja), ventas[clave])

    update_status("✅ Consolidación completada")

    return {
        "ruta":                 ruta_final,
        "filas_gallo":          len(gallo),
        "filas_monivoi":        0 if monivoi_sin_datos else len(monivoi),
        "refs_unicas":          len(refs_unicas),
        "filas_catalogo":       len(catalogo_docs),
        "filas_globales":       len(globales),
        "monivoi_sin_datos":    monivoi_sin_datos,
        "filas_reclasificadas": filas_reclasificadas,
        "cedis_facturas":       len(td_resumen["facturas"]["cedis"]),
        "cedis_notas":          len(td_resumen["notas"]["cedis"]),
        "cedis_pep":            len(td_resumen["pep"]["cedis"]),
        "total_facturas":       td_resumen["total_facturas"],
        "total_notas":          td_resumen["total_notas"],
        "total_pep":            td_resumen["total_pep"],
        "total_importe_gallo":  total_importe_gallo,
        "total_venta_dsd":      ventas["dsd"]["total"],
        "total_venta_eiap":     ventas["eiap"]["total"],
        "venta_avisos":         ventas["avisos"],
    }


# ----------------------------------------------------------------------
# Matriz 'TD' — cálculo y escritura
# ----------------------------------------------------------------------
def _preparar_matriz_td(gallo, col_centro, col_asig, col_clase, importe_num):
    """
    Build the aggregated data for the 'TD' sheet.

    IMPORTANT: every matrix derives its OWN cedis list from its OWN filters
    (document class included), so the three blocks can list different centros
    and have different row counts. There is no 80AB..80VW range clipping: any
    centro that satisfies a matrix's filters becomes a row of that matrix.

        Facturas     → Referencia 3=='Global' & Asignación != 'MX Commercial'
                       & Clase de documento in {C1, C5}
        Notas        → same base, Clase de documento in {C2, C6}
        PEP vouchers → Referencia 3=='Global' & Referencia 2=='PP'
                       & Clase de documento=='C2'   (MX Commercial NOT excluded)

    Args:
        importe_num (pd.Series): 'Importe en moneda local' ALREADY parsed to
            float by the caller (single pass shared with the 'Venta ...' sheets).

    Returns a dict with one entry per block ('facturas', 'notas', 'pep'), each
    holding its own 'cedis' list and 'suma' lookup, plus the control totals and
    the grand total of the whole 'Importe en moneda local' column.
    """
    # Normalize once.
    centro_norm = gallo[col_centro].astype(str).str.strip()
    clase_norm  = gallo[col_clase].astype(str).str.strip().str.upper()
    ref2_norm   = gallo["Referencia 2"].astype(str).str.strip().str.upper()
    asig_norm   = gallo[col_asig].astype(str).str.strip()
    es_global   = gallo["Referencia 3"] == "Global"
    con_centro  = centro_norm != ""

    # Base subset shared by Facturas and Notas (each adds its own class filter).
    base = es_global & (asig_norm != VALOR_ASIGNACION) & con_centro

    def _bloque(mascara, clases):
        """Cedis list + {(centro, clase): importe} for one two-class block."""
        m = mascara & clase_norm.isin({clase for clase, _ in clases})
        cedis = sorted(centro_norm[m].unique())
        if not m.any():
            return {"cedis": cedis, "suma": {}}
        datos = pd.DataFrame({
            "_centro":  centro_norm[m],
            "_clase":   clase_norm[m],
            "_importe": importe_num[m],
        })
        serie = datos.groupby(["_centro", "_clase"])["_importe"].sum()
        # Round to cents so no float noise (e.g. -265.44000000000005) leaks
        # into the cells or the control totals.
        return {"cedis": cedis,
                "suma": {k: round(float(v), 2) for k, v in serie.items()}}

    facturas = _bloque(base, TD_FACTURAS["clases"])
    notas    = _bloque(base, TD_NOTAS["clases"])

    # PEP vouchers: its own filter, and it does NOT exclude 'MX Commercial'
    # (el stakeholder pidió sólo esos tres filtros).
    mascara_pep = (
        es_global
        & con_centro
        & (ref2_norm == TD_PEP_REFERENCIA2)
        & (clase_norm == TD_PEP_CLASE)
    )
    cedis_pep = sorted(centro_norm[mascara_pep].unique())
    if mascara_pep.any():
        datos_pep = pd.DataFrame({
            "_centro":  centro_norm[mascara_pep],
            "_importe": importe_num[mascara_pep],
        })
        serie_pep = datos_pep.groupby("_centro")["_importe"].sum()
        suma_pep  = {c: round(float(v), 2) for c, v in serie_pep.items()}
    else:
        suma_pep = {}

    return {
        "facturas":            facturas,
        "notas":               notas,
        "pep":                 {"cedis": cedis_pep, "suma": suma_pep},
        "total_facturas":      round(float(sum(facturas["suma"].values())), 2),
        "total_notas":         round(float(sum(notas["suma"].values())), 2),
        "total_pep":           round(float(sum(suma_pep.values())), 2),
        "total_importe_gallo": round(float(importe_num.sum()), 2),
    }


def _escribir_matriz_td(ws_td, td_resumen):
    """
    Write the three matrices onto the 'TD' worksheet.

    Layout (rows 1-3 = stakeholder labels, row 4 = block title, row 5 = headers,
    row 6+ = data, and a 'Grand Total' row right below the last cedis):

        Facturas         → A: cedis, B: sum(C1), C: sum(C5), D: Grand Total (B+C)
        Notas de Crédito → L: cedis, M: sum(C2), N: sum(C6), O: Grand Total (M+N)
        PEP vouchers     → W: cedis, X: sum(C2 con Referencia 2 == 'PP')

    Extras:
        - 'Grand Total' column per block (D / O): per-row sum of its two classes.
        - 'Grand Total' row at the bottom: per-column sum; the corner cell is the
          block's overall total.
        - Analysis labels 'Monitor' / 'Contabilidad' / 'Diferencias' in D1:D3 and
          O1:O3 for the stakeholder's later work (labels only, no values).

    Each block owns its cedis list (derived from its OWN filters), so the blocks
    can differ in which centros they show and in how many rows they have — every
    block therefore computes its own 'Grand Total' row position. Body cells are
    real numbers (values, not formulas) with a thousands/2-decimals format.
    """
    for bloque, datos in ((TD_FACTURAS, td_resumen["facturas"]),
                          (TD_NOTAS,    td_resumen["notas"])):
        cedis      = datos["cedis"]
        suma       = datos["suma"]
        n_cedis    = len(cedis)
        fila_total = TD_FILA_INICIO + n_cedis   # row right below the last cedis
        col_cedis  = bloque["col_cedis"]
        col_gt     = bloque["col_grand_total"]
        clases     = bloque["clases"]

        # Analysis labels for the stakeholder (rows 1-3, in the Grand Total col).
        for fila, texto in TD_LABELS_ANALISIS:
            ws_td.cell(row=fila, column=col_gt, value=texto).font = Font(bold=True)

        # Block title (row 4).
        ws_td.cell(row=4, column=col_cedis, value=bloque["titulo"]).font = Font(bold=True)

        # Header row (row 5): 'Centro' + class headers + 'Grand Total'.
        ws_td.cell(row=TD_FILA_ENCABEZADO, column=col_cedis, value="Centro").font = Font(bold=True)
        for clase, col_idx in clases:
            ws_td.cell(row=TD_FILA_ENCABEZADO, column=col_idx, value=clase).font = Font(bold=True)
        ws_td.cell(row=TD_FILA_ENCABEZADO, column=col_gt, value="Grand Total").font = Font(bold=True)

        # Body rows + per-column accumulators for the bottom 'Grand Total' row.
        totales_clase = {clase: 0.0 for clase, _ in clases}
        for i, centro in enumerate(cedis):
            fila = TD_FILA_INICIO + i
            ws_td.cell(row=fila, column=col_cedis, value=centro)
            suma_fila = 0.0
            for clase, col_idx in clases:
                valor = round(suma.get((centro, clase), 0.0), 2)
                ws_td.cell(row=fila, column=col_idx, value=valor).number_format = TD_FORMATO_NUMERO
                totales_clase[clase] += valor
                suma_fila += valor
            # 'Grand Total' column = sum of this row's classes.
            ws_td.cell(row=fila, column=col_gt, value=round(suma_fila, 2)).number_format = TD_FORMATO_NUMERO

        # Bottom 'Grand Total' row = per-column totals (only when there are cedis).
        if n_cedis > 0:
            ws_td.cell(row=fila_total, column=col_cedis, value="Grand Total").font = Font(bold=True)
            suma_total = 0.0
            for clase, col_idx in clases:
                tot = round(totales_clase[clase], 2)
                celda = ws_td.cell(row=fila_total, column=col_idx, value=tot)
                celda.number_format = TD_FORMATO_NUMERO
                celda.font = Font(bold=True)
                suma_total += tot
            celda_gt = ws_td.cell(row=fila_total, column=col_gt, value=round(suma_total, 2))
            celda_gt.number_format = TD_FORMATO_NUMERO
            celda_gt.font = Font(bold=True)

    # --- Third matrix: PEP vouchers (its own cedis, single value column) ---
    cedis_pep      = td_resumen["pep"]["cedis"]
    suma_pep       = td_resumen["pep"]["suma"]
    col_cedis_pep  = TD_PEP["col_cedis"]
    col_valor_pep  = TD_PEP["col_valor"]
    fila_total_pep = TD_FILA_INICIO + len(cedis_pep)

    ws_td.cell(row=4, column=col_cedis_pep, value=TD_PEP["titulo"]).font = Font(bold=True)
    ws_td.cell(row=TD_FILA_ENCABEZADO, column=col_cedis_pep, value="Centro").font = Font(bold=True)
    ws_td.cell(row=TD_FILA_ENCABEZADO, column=col_valor_pep,
               value=TD_PEP["encabezado_valor"]).font = Font(bold=True)

    total_pep = 0.0
    for i, centro in enumerate(cedis_pep):
        fila = TD_FILA_INICIO + i
        ws_td.cell(row=fila, column=col_cedis_pep, value=centro)
        valor = round(suma_pep.get(centro, 0.0), 2)
        ws_td.cell(row=fila, column=col_valor_pep, value=valor).number_format = TD_FORMATO_NUMERO
        total_pep += valor
    if cedis_pep:
        ws_td.cell(row=fila_total_pep, column=col_cedis_pep, value="Grand Total").font = Font(bold=True)
        celda_pep = ws_td.cell(row=fila_total_pep, column=col_valor_pep, value=round(total_pep, 2))
        celda_pep.number_format = TD_FORMATO_NUMERO
        celda_pep.font = Font(bold=True)

    # Comfortable widths (incl. the Grand Total columns D/O and the PEP W/X).
    anchos = {"A": 12, "B": 16, "C": 16, "D": 16,
              "L": 12, "M": 16, "N": 16, "O": 16,
              "W": 12, "X": 16}
    for letra, ancho in anchos.items():
        ws_td.column_dimensions[letra].width = ancho


# ----------------------------------------------------------------------
# Pestañas 'Venta DSD' / 'Venta EIAP' — cálculo y escritura
# ----------------------------------------------------------------------
def _preparar_ventas(gallo, col_clase, col_cuenta, col_fecha, importe_num, fechas_sel):
    """
    Build the two date × amount matrices ('Venta DSD' and 'Venta EIAP').

    TWO filters, both required (AND): 'Clase de documento' and 'Cuenta'
    (stakeholder's call). Unlike the 'TD' sheet, these blocks do NOT require
    Referencia 3=='Global' and do NOT exclude 'MX Commercial'.

        Venta DSD  → clase in {C1, C2, C5, C6, X0} or blank,
                     AND cuenta in {4000010, 4000005}
        Venta EIAP → clase == RV,
                     AND cuenta in {4000010, 4000020, 4000802, 4300013}

    Accounts are compared normalized (no leading zeros, no trailing '.0'), so
    the export's number formatting cannot silently empty the sheet.

    Rows are the union of (a) the dates the user selected in the GUI — always
    present, 0.00 when the day has no matching rows — and (b) the dates actually
    found in 'Fe.contabilización', so nothing that exists in the data is
    silently dropped. That column ships as 'dd/mm/aaaa hh:mm': the time is
    trimmed, so all the movements of one day land on one row. Rows with an empty
    date are kept apart under a 'Sin fecha' label so the block's 'Grand Total'
    always reconciles with the filters.

    Args:
        importe_num (pd.Series): amounts already parsed to float.
        fechas_sel (list[date]): GUI-selected dates, already validated.

    Returns:
        dict: {'dsd': bloque, 'eiap': bloque, 'avisos': [str]} where each bloque
        is {'filas': [(date, float)], 'sin_fecha_n': int,
            'sin_fecha_importe': float, 'total': float, 'fuera_rango': [date]}.
    """
    clase_norm  = gallo[col_clase].astype(str).str.strip().str.upper()
    cuenta_norm = gallo[col_cuenta].map(_normalizar_cuenta)
    # Trim the time BEFORE detecting the format: '01/09/2026 14:35' is a day.
    fecha_raw   = gallo[col_fecha].map(_solo_fecha)

    # Detect the date format ONCE, over the distinct values of the whole column.
    distintas = sorted({valor for valor in fecha_raw.unique() if valor})
    formato = _detectar_formato_fecha(distintas, fechas_sel)
    mapa_fechas = (
        {valor: datetime.strptime(valor, formato).date() for valor in distintas}
        if formato else {}
    )
    # Blank dates map to None so they can be reported instead of vanishing.
    fecha_norm = fecha_raw.map(lambda valor: mapa_fechas.get(valor))

    cuentas_presentes = set(cuenta_norm.unique())
    seleccionadas = set(fechas_sel)
    avisos = []
    bloques = {}

    for clave, (etiqueta, clases, cuentas) in (
        ("dsd",  ("Venta DSD",  VENTA_CLASES_DSD,  VENTA_CUENTAS_DSD)),
        ("eiap", ("Venta EIAP", VENTA_CLASES_EIAP, VENTA_CUENTAS_EIAP)),
    ):
        cuentas_buscadas = {_normalizar_cuenta(c) for c in cuentas}
        mascara = clase_norm.isin(clases) & cuenta_norm.isin(cuentas_buscadas)

        # Una cuenta pedida que NO existe en el archivo casi siempre significa
        # que ese día no tuvo movimientos... o que el layout de SAP cambió.
        # Vale más avisarlo que entregar una matriz vacía sin explicación.
        faltantes = sorted(c for c in cuentas_buscadas if c not in cuentas_presentes)
        if faltantes:
            avisos.append(
                f"{etiqueta}: la(s) cuenta(s) {', '.join(faltantes)} no "
                f"aparece(n) en el archivo de Gallo."
            )

        datos = pd.DataFrame({
            "_fecha":   fecha_norm[mascara],
            "_importe": importe_num[mascara],
        })

        sin_fecha = datos["_fecha"].isna()
        sin_fecha_n = int(sin_fecha.sum())
        sin_fecha_importe = round(float(datos.loc[sin_fecha, "_importe"].sum()), 2)

        con_fecha = datos.loc[~sin_fecha]
        if con_fecha.empty:
            suma = {}
        else:
            serie = con_fecha.groupby("_fecha")["_importe"].sum()
            suma = {fecha: round(float(valor), 2) for fecha, valor in serie.items()}

        # Selected days always show up (0.00 if empty); data-only days too.
        filas = [(fecha, suma.get(fecha, 0.0))
                 for fecha in sorted(set(suma) | seleccionadas)]
        fuera_rango = sorted(f for f in suma if seleccionadas and f not in seleccionadas)

        total = round(sum(valor for _, valor in filas) + sin_fecha_importe, 2)

        if sin_fecha_n:
            avisos.append(
                f"{etiqueta}: {sin_fecha_n:,} fila(s) sin fecha de contabilización "
                f"(se muestran como '{VENTA_ETIQUETA_SIN_FECHA}')."
            )
        if fuera_rango:
            listado = ", ".join(f.strftime("%d.%m.%Y") for f in fuera_rango[:5])
            extra = "..." if len(fuera_rango) > 5 else ""
            avisos.append(
                f"{etiqueta}: hay fechas fuera de las seleccionadas "
                f"({listado}{extra}); se incluyeron de todas formas."
            )

        bloques[clave] = {
            "filas":             filas,
            "sin_fecha_n":       sin_fecha_n,
            "sin_fecha_importe": sin_fecha_importe,
            "total":             total,
            "fuera_rango":       fuera_rango,
        }

    bloques["avisos"] = avisos
    return bloques


def _escribir_hoja_venta(ws, bloque):
    """
    Write one 'Venta ...' sheet: column A = date, column B = summed amount.

    Row 1 = headers, row 2+ = one row per day, then (only if it exists) a
    'Sin fecha' row, and finally a bold 'Grand Total' row. Dates are written as
    REAL dates formatted DD.MM.YYYY, so the stakeholder can sort and filter them
    as dates; amounts are values (not formulas) with 2 decimals.
    """
    ws.cell(row=1, column=1, value=VENTA_ENCABEZADO_FECHA).font = Font(bold=True)
    ws.cell(row=1, column=2, value=VENTA_ENCABEZADO_IMPORTE).font = Font(bold=True)

    fila = 2
    for fecha, valor in bloque["filas"]:
        celda_fecha = ws.cell(row=fila, column=1, value=fecha)
        celda_fecha.number_format = VENTA_FORMATO_FECHA_EXCEL
        ws.cell(row=fila, column=2, value=valor).number_format = TD_FORMATO_NUMERO
        fila += 1

    # Rows whose date was blank: shown apart so the total still reconciles.
    if bloque["sin_fecha_n"]:
        ws.cell(row=fila, column=1, value=VENTA_ETIQUETA_SIN_FECHA).font = Font(italic=True)
        celda = ws.cell(row=fila, column=2, value=bloque["sin_fecha_importe"])
        celda.number_format = TD_FORMATO_NUMERO
        celda.font = Font(italic=True)
        fila += 1

    if fila > 2:   # there is at least one data row
        ws.cell(row=fila, column=1, value=VENTA_ETIQUETA_TOTAL).font = Font(bold=True)
        celda_total = ws.cell(row=fila, column=2, value=bloque["total"])
        celda_total.number_format = TD_FORMATO_NUMERO
        celda_total.font = Font(bold=True)

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 24


# ----------------------------------------------------------------------
# Ejecución directa (prueba rápida sin GUI ni controller)
# ----------------------------------------------------------------------
if __name__ == "__main__":
    print("=== Prueba de consolidación (standalone) ===")
    ruta_gallo   = input("Ruta del CSV de Gallo   : ").strip().strip('"')
    ruta_monivoi = input("Ruta del CSV de Monivoi : ").strip().strip('"')
    fechas_txt   = input(
        "Fechas seleccionadas 'DD.MM.YYYY' separadas por coma (Enter = ninguna): "
    ).strip()
    fechas = [f.strip() for f in fechas_txt.split(",") if f.strip()] or None

    salida = carpeta_output_desde_input(os.path.dirname(ruta_gallo))

    try:
        resumen = consolidar_gallo_monivoi(
            ruta_gallo, ruta_monivoi, salida,
            callback_status=print,
            fechas_seleccionadas=fechas,
        )
        print("\n🎉 CONSOLIDACIÓN COMPLETADA 🎉")
        for clave, valor in resumen.items():
            print(f"  {clave:<20}: {valor}")
    except Exception as e:
        print(f"❌ Error durante la consolidación: {e}")
        raise
