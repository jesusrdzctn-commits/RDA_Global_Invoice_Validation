"""
Validación Factura Global — Consolidation: shared helpers
=========================================================

Building blocks used by BOTH sides of the consolidation (Gallo and Monivoi).
Nothing in here knows about a specific sheet or business rule:

    leer_csv / es_sin_datos  → read a CSV downloaded by the app
    col_por_candidatos       → find a column, accepting spelling variants
    parse_importe_us         → turn a US-formatted SAP amount into a float
    escribir_filtros         → write the filter legend above a matrix
    ETIQUETA_GRAND_TOTAL, TD_FORMATO_NUMERO, VENTA_ENCABEZADO_IMPORTE
                             → labels/formats shared by several sheets

Dependency rule: this module imports NOTHING from the other Consolidacion_*
modules, so it can never be part of an import cycle.
"""

import os
import re
import unicodedata

import pandas as pd
from openpyxl.styles import Font

# ----------------------------------------------------------------------
# Labels and formats shared by several sheets
# ----------------------------------------------------------------------
# Label of the totals row in 'TD' and in 'MTD'. It MUST be the same on both
# sheets: the TD × MTD cross looks it up with XLOOKUP to compare total against
# total. That is why it lives in ONE single constant (here, because both the
# Gallo side and the Monivoi side need it).
ETIQUETA_GRAND_TOTAL = "Grand Total"

# Number format of every amount cell (sheets 'TD', 'MTD', 'Venta ...' and
# the TD × MTD cross). It is named after 'TD', where it was first used.
TD_FORMATO_NUMERO = "#,##0.00"

# Header of Gallo's amount column. It is also the column-B header of the
# 'Venta ...' sheets and the default column label in parse_importe_us error
# messages; that last use is why it lives here and not in Consolidacion_Gallo.
VENTA_ENCABEZADO_IMPORTE = "Importe en moneda local"

# Look of the filter legend written above every matrix ('TD' and 'MTD'):
# italic and gray, so it reads as a note and never as data.
FILTROS_COLOR = "FF595959"

# A US amount with a thousands separator ALWAYS groups digits in threes
# ('1,234.56'). Anything else with a comma (the European '1.234,56', or a
# '12,5') is NOT US: blindly dropping the comma would turn 1.234,56 into
# 1.23456 and silently understate the sum. That is why it is validated before
# cleaning.
_RE_IMPORTE_US_AGRUPADO = re.compile(r"\d{1,3}(?:,\d{3})+(?:\.\d+)?")


# ----------------------------------------------------------------------
# Reading the downloaded CSVs
# ----------------------------------------------------------------------
def leer_csv(ruta):
    """Read a CSV downloaded by the app: everything as text, accents included."""
    if not os.path.exists(ruta):
        raise FileNotFoundError(f"Archivo no encontrado: {ruta}")
    # keep_default_na=False → empty cells stay as "" (not NaN), so .strip()
    # and text comparisons never blow up.
    return pd.read_csv(ruta, dtype=str, encoding="utf-8-sig", keep_default_na=False)


def es_sin_datos(df):
    """True if the CSV is the 'Sin datos' placeholder written by the download."""
    return df.empty or (len(df.columns) == 1 and "Sin datos" in str(df.columns[0]))


# ----------------------------------------------------------------------
# Column lookup
# ----------------------------------------------------------------------
def _normalizar_encabezado(texto):
    """
    Normalize a header so it can be compared ignoring accents, dots, spaces and
    case: 'Fe.Contabilización' and 'fecha contabilizacion' become comparable.
    Only used to recognize SPELLING VARIANTS of the same field, never to guess
    a different column.
    """
    sin_acentos = unicodedata.normalize("NFKD", str(texto))
    sin_acentos = "".join(c for c in sin_acentos if not unicodedata.combining(c))
    return "".join(c for c in sin_acentos.lower() if c.isalnum())


def col_por_candidatos(df, candidatos, descripcion, origen="Gallo"):
    """
    Resolve a column accepting several spellings of the SAME field
    ('Fe.contabilización' / 'Fecha contabilización', 'Cuentas' / 'Cuenta de
    mayor'). The comparison ignores accents, dots, spaces and case.

    `origen` is only used in the error message ('Gallo' or 'Monivoi'), so the
    stakeholder knows WHICH file to check.

    If none of them is present, it fails with an actionable message that lists
    the real columns: stopping is better than filtering/grouping by the wrong
    column.
    """
    normalizadas = {_normalizar_encabezado(c): c for c in df.columns}
    for candidata in candidatos:
        real = normalizadas.get(_normalizar_encabezado(candidata))
        if real is not None:
            return real
    raise ValueError(
        f"No encontré la columna de {descripcion} en el archivo de {origen} "
        f"(busqué: {', '.join(candidatos)}). "
        f"Columnas disponibles: {list(df.columns)}"
    )


# ----------------------------------------------------------------------
# Amounts
# ----------------------------------------------------------------------
def parse_importe_us(valor, columna=VENTA_ENCABEZADO_IMPORTE):
    """
    Convert a US-formatted SAP amount string into a float.

    Handles the shapes SAP actually exports (US locale):
        - thousands separator ','  and decimal point '.'   -> '1,234.56'
        - SAP trailing minus (sign after the number)        -> '1,234.56-'
        - leading minus / stray leading '+'                 -> '-1,234.56'
        - accounting parentheses                            -> '(1,234.56)'
        - blank / whitespace / lone '-'                     -> 0.0

    A non-empty token that is still not numeric after cleanup raises ValueError,
    so corrupt amounts stop the run instead of quietly understating a sum. Same
    for a comma that is NOT a US thousands separator (e.g. the European
    '1.234,56'): dropping it would turn the amount into 1.23456.

    `columna` is only used in the error message, so it is clear whether the
    odd amount came from 'Importe en moneda local' (Gallo) or from 'Valor neto'
    (Monivoi).
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

    s = s.replace(" ", "")
    if s == "":
        return 0.0

    # US format: ',' is ONLY a thousands separator, and always groups 3 digits.
    # Validate BEFORE stripping it, so a European-formatted amount fails loud
    # instead of silently becoming a different (smaller) number.
    if "," in s:
        if not _RE_IMPORTE_US_AGRUPADO.fullmatch(s):
            raise ValueError(
                f"El importe '{valor}' de la columna '{columna}' no tiene "
                f"formato US (esperaba p. ej. 1,234.56). Si SAP lo exportó en "
                f"formato europeo (1.234,56), hay que corregir el formato del "
                f"export antes de consolidar."
            )
        s = s.replace(",", "")

    try:
        numero = float(s)
    except ValueError:
        raise ValueError(
            f"No pude convertir el importe '{valor}' a número en la columna "
            f"'{columna}'. ¿Formato inesperado? "
            f"(esperaba formato US, p. ej. 1,234.56 o 1,234.56-)."
        )

    return -numero if negativo else numero


# ----------------------------------------------------------------------
# Filter legend shown above each matrix
# ----------------------------------------------------------------------
def escribir_filtros(ws, fila_inicial, columna, filtros):
    """
    Write the filter legend of a matrix, one line per row, starting at
    (fila_inicial, columna) and going down.

    It is what lets the stakeholder reproduce any figure by hand: it names the
    source sheet and the exact conditions the rows of that matrix meet. Each
    caller builds its lines FROM ITS OWN filter constants, so the legend can
    never drift from what the code really sums.

    The text is italic and gray so nobody mistakes it for data, and it is
    written in the block's first column, on rows that the layout keeps free.
    """
    for i, texto in enumerate(filtros):
        celda = ws.cell(row=fila_inicial + i, column=columna, value=texto)
        celda.font = Font(italic=True, color=FILTROS_COLOR)
