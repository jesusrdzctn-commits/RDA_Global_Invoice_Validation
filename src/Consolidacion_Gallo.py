"""
Validación Factura Global — Consolidation: Gallo (FAGLL03)
==========================================================

Everything that is computed FROM the Gallo CSV:

    1) Steps 2-9 + SG/RG override → helper columns K, L, M of the 'Gallo'
       sheet ('Referencia 2', 'Referencia 3', 'Comercializadora')
                                            → preparar_gallo
    2) 'Validación - Doc.' sheet            → escribir_hoja_validacion
    3) 'TD' matrices (Facturas, Notas de Crédito and PEP vouchers)
                                            → preparar_matriz_td / escribir_matriz_td
    4) 'Venta DSD' and 'Venta EIAP' sheets  → preparar_ventas / escribir_hoja_venta

It also validates the dates selected in the GUI (parsear_fechas_seleccionadas),
because only the 'Venta ...' sheets use them.

The cross of 'TD' against the monitor ('MTD') lives in Consolidacion_Monivoi
and reads the TD layout constants defined here (TD_FACTURAS, TD_NOTAS, TD_PEP,
TD_FILA_*): a change to that layout must keep both modules in sync.

Dependencies: Consolidacion_Comun only (this module never imports Monivoi).
"""

import re
from datetime import datetime
from functools import partial

import pandas as pd
from openpyxl.styles import Font

from Consolidacion_Comun import (
    ETIQUETA_GRAND_TOTAL,
    TD_FORMATO_NUMERO,
    VENTA_ENCABEZADO_IMPORTE,
    col_por_candidatos,
    escribir_filtros,
    parse_importe_us,
)

# ----------------------------------------------------------------------
# Process constants
# ----------------------------------------------------------------------
VALOR_ASIGNACION = "MX Commercial"          # filter of steps 5, 8.5 and 9
PREFIJOS_LETRAS_GLOBAL = {"PP", "RG", "SG"}  # letter-only prefixes that are also 'Global'
PREFIJOS_RECLASIFICAR = {"SG", "RG"}         # step 8.5: SG/RG 'MX Commercial' → Individual

# Sheet names (tabs) built from the Gallo.
HOJA_GALLO      = "Gallo"
HOJA_VALIDACION = "Validación - Doc."
HOJA_TD         = "TD"

# --- 'TD' matrix layout ---
# Rows 1-3 = filter legend, row 4 = block title, row 5 = headers, data starts
# at row 6. Facturas live on the left block (row labels in column A), Notas de
# Crédito in the middle one (column L) and PEP vouchers on the right (column W).
# The legend goes in rows 1-3 of each block's FIRST column, which the layout
# keeps free (the 'Monitor'/'Contabilidad'/'Diferencias' labels sit in the
# Grand Total column, see TD_LABELS_ANALISIS).
TD_FILA_FILTROS    = 1
TD_FILA_ENCABEZADO = 5
TD_FILA_INICIO     = 6
# 'clases' = (document class, 1-based column index of its header/sum).
# Each block has its own 'Grand Total' column (D in Facturas, O in Notas).
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
# Analysis labels for the stakeholder: they go in the Grand Total column,
# rows 1-3 (D1:D3 in Facturas, O1:O3 in Notas). Labels only.
TD_LABELS_ANALISIS = [(1, "Monitor"), (2, "Contabilidad"), (3, "Diferencias")]

# Third matrix: PEP vouchers. Filter: Referencia 3=='Global',
# Referencia 2=='PP' and Clase de documento=='C2' (plus the non-blank Centro
# that every TD block needs). A single value column (X), with the cedis rows
# in W. It does NOT exclude 'MX Commercial'.
TD_PEP = {
    "titulo": "PEP VOUCHERS",
    "col_cedis": 23,           # W
    "col_valor": 24,           # X
    "encabezado_valor": "C2",  # class of the filter (goes in X5)
}
TD_PEP_REFERENCIA2 = "PP"
TD_PEP_CLASE = "C2"

# --- 'Venta DSD' and 'Venta EIAP' sheets ---
# Two-column matrix: date (A) and summed amount (B). Filtered by TWO criteria
# at once (AND): document class AND account. The empty string "" among the
# classes stands for Excel's "(blanks)", i.e. a class that is empty or only
# spaces.
HOJA_VENTA_DSD  = "Venta DSD"
HOJA_VENTA_EIAP = "Venta EIAP"
VENTA_CLASES_DSD  = frozenset({"C1", "C2", "C5", "C6", "X0", ""})
VENTA_CLASES_EIAP = frozenset({"RV"})
# G/L accounts. They are compared NORMALIZED (no leading zeros and no '.0'
# left behind by the xlsx→csv conversion), so it does not matter whether SAP
# exports them as '4000010', '0004000010' or '4000010.0'.
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
# (VENTA_ENCABEZADO_IMPORTE lives in Consolidacion_Comun: parse_importe_us uses it too.)
VENTA_ETIQUETA_SIN_FECHA = "Sin fecha"
VENTA_ETIQUETA_TOTAL     = "Grand Total"
VENTA_FORMATO_FECHA_EXCEL = "DD.MM.YYYY"

# Layout: rows 1-4 = filter legend, row 5 = headers, row 6 onwards = data.
# The legend needs four lines here because this matrix filters by TWO criteria
# at once (document class AND account). Data therefore starts on row 6, the
# same as in 'TD' and 'MTD'. No formula points at these sheets, so moving the
# rows down affects nothing else in the workbook.
VENTA_FILA_FILTROS    = 1
VENTA_FILA_ENCABEZADO = 5
VENTA_FILA_INICIO     = 6
# How the empty document class (Excel's "(blanks)") reads in the legend: it IS
# one of the classes that DSD sums, and '' on screen would look like a typo.
VENTA_ETIQUETA_EN_BLANCO = "en blanco"

# Date formats the SAP export may use. The 4-digit-year formats are tried
# first; the 2-digit ones are the last resort. Real ambiguity (e.g. 01/09/2026
# = Sep 1 or Jan 9) is resolved against the dates the user selected in the
# GUI; if it cannot be resolved, it fails with a clear message instead of
# guessing.
VENTA_FORMATOS_FECHA_4D = (
    "%d.%m.%Y", "%m.%d.%Y",
    "%d/%m/%Y", "%m/%d/%Y",
    "%Y-%m-%d", "%Y/%m/%d",
    "%d-%m-%Y", "%m-%d-%Y",
)
VENTA_FORMATOS_FECHA_2D = ("%d.%m.%y", "%m.%d.%y", "%d/%m/%y", "%m/%d/%y")
VENTA_ANIO_MIN = 1990
VENTA_ANIO_MAX = 2100

# The Gallo export brings the date as 'dd/mm/yyyy hh:mm'. This pattern keeps
# only the DATE part (group 1) and drops whatever comes after a space or a 'T'.
# It does not try to interpret the day/month order: _detectar_formato_fecha
# takes care of that.
_RE_SOLO_FECHA = re.compile(r"^(\d{1,4}[./\-]\d{1,2}[./\-]\d{1,4})(?:[ T].*)?$")


# ----------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------
def _col(df, nombre, indice, letra):
    """
    Return the REAL name of a column: look it up by name first
    (e.g. 'Referencia'); if it does not exist, fall back to the expected
    position (e.g. column C = index 2). This tolerates small layout changes.
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


def _col_fecha_contabilizacion(df):
    """Gallo's posting-date column ('Fe.contabilización')."""
    return col_por_candidatos(
        df, VENTA_COL_FECHA_CANDIDATAS, "fecha de contabilización"
    )


def _col_cuenta(df):
    """Gallo's G/L account column ('Cuentas')."""
    return col_por_candidatos(df, VENTA_COL_CUENTA_CANDIDATAS, "cuenta")


def _normalizar_cuenta(valor):
    """
    Normalize a G/L account so it compares without format surprises:

        ' 0004000010 ' -> '4000010'
        '4000010.0'    -> '4000010'   (the xlsx→csv conversion made it a float)
        '4000010'      -> '4000010'

    Only LEADING zeros and an all-zero decimal part are removed; never
    significant digits. A non-numeric value is returned as is (upper-cased and
    without spaces), so an account with letters stays comparable.
    """
    texto = str(valor).strip().upper().replace(" ", "")
    if not texto:
        return ""
    # '4000010.0' / '4000010.00' → '4000010' (only when the decimals are zeros)
    if re.fullmatch(r"\d+\.0*", texto):
        texto = texto.split(".", 1)[0]
    if texto.isdigit():
        return texto.lstrip("0") or "0"
    return texto


def _solo_fecha(valor):
    """
    Cut the time off a date value: the Gallo export brings 'dd/mm/yyyy hh:mm'
    and the time gets in the way when grouping by day.

        '01/09/2026 14:35' -> '01/09/2026'
        '2026-09-01T00:00' -> '2026-09-01'
        '01.09.2026'       -> '01.09.2026'

    If the value does not start with something that looks like a date, it is
    returned as is (just stripped), so the format detector fails showing the
    REAL value instead of one mangled by this helper.
    """
    texto = str(valor).strip()
    coincidencia = _RE_SOLO_FECHA.match(texto)
    return coincidencia.group(1) if coincidencia else texto


def _clasificar_tipo(referencia2):
    """
    Step 4: 'Global' if it contains at least one digit, or if it is PP/RG/SG.
    'Individual' for the rest (e.g. letter-only prefixes).
    """
    ref = str(referencia2).strip()
    if any(caracter.isdigit() for caracter in ref):
        return "Global"
    if ref.upper() in PREFIJOS_LETRAS_GLOBAL:
        return "Global"
    return "Individual"


# ----------------------------------------------------------------------
# Dates for the 'Venta ...' sheets
# ----------------------------------------------------------------------
def parsear_fechas_seleccionadas(fechas):
    """
    Validate and normalize the dates the user picked in the GUI ('DD.MM.YYYY')
    into date objects, without duplicates and in chronological order.

    Returns [] if none were received (the module stays usable on its own).
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
    Return the list of dates if ALL values parse with that format and fall in
    a reasonable year; None if any of them fails.

    The year check is not cosmetic: strptime accepts '26' for %Y and would
    return the year 26 AD, which would let a wrong format pass as valid.
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
    Detect the format of the date column of the Gallo CSV.

    Rule: the 4-digit-year formats are tried first and, only if none of them
    works, the 2-digit ones. If several formats parse EVERYTHING:
      - if they all give the SAME result, the ambiguity is harmless → the 1st
        one is used;
      - if they give different results, the tie is broken with the dates the
        user selected (the format that puts more dates inside the selection
        wins);
      - if it is still a tie, ValueError is raised. It never guesses: a
        misread 01/09/2026 would silently move the amount to another day.

    Returns the format (str), or None if there is no value to parse.
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
            return viables[0][0]   # same result: the ambiguity is harmless

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
# Steps 2-9 — helper columns of the 'Gallo' sheet + 'Validación - Doc.'
# ----------------------------------------------------------------------
def preparar_gallo(gallo, update_status):
    """
    Steps 2-9 + the SG/RG override + ONE amount-parsing pass, over the Gallo
    DataFrame that the orchestrator already read.

    It MUTATES `gallo` in place: it appends the helper columns 'Referencia 2',
    'Referencia 3' and 'Comercializadora' (K, L and M with the standard A-J
    layout), which is exactly what the 'Gallo' sheet shows.

    The step order is part of the business logic: step 9 takes its snapshot of
    the globals BEFORE the SG/RG override (called "step 8.5" in the constants)
    reclassifies rows. Do not reorder.

    Args:
        gallo (pd.DataFrame): the Gallo CSV as read by leer_csv (all text).
        update_status (callable): receives each progress message (str); pass
            `lambda _mensaje: None` to run it silently.

    Returns:
        dict: what the next steps need →
            col_centro, col_asig, col_doc, col_clase, col_fecha, col_cuenta
                (the REAL column names found in this file),
            refs_unicas, tipos (steps 3-4), catalogo_docs (step 6),
            globales (step 9), filas_reclasificadas (SG/RG override),
            importe_num ('Importe en moneda local' as float, same index as
                `gallo`).

    Raises:
        ValueError: a required column is missing or an amount cannot be
            interpreted.
    """
    # Gallo key columns (by name, with a positional fallback)
    col_centro = _col(gallo, "Centro",       1, "B")
    col_ref    = _col(gallo, "Referencia",   2, "C")
    col_asig   = _col(gallo, "Asignación",   3, "D")
    col_doc    = _col(gallo, "Nº documento", 4, "E")

    # Columns the matrices need. They are resolved STRICTLY by name (their
    # position in the export is not fixed) and fail early with a clear error
    # if missing, before anything is processed.
    col_clase   = _col_por_nombre(gallo, "Clase de documento")
    col_importe = _col_por_nombre(gallo, "Importe en moneda local")
    col_fecha   = _col_fecha_contabilizacion(gallo)
    col_cuenta  = _col_cuenta(gallo)

    # === Step 2: 'Referencia 2' = first 2 characters of 'Referencia'
    # It is appended AT THE END of the table; with the standard layout (A-J)
    # that puts it exactly in column K.
    update_status("🔧 Creando 'Referencia 2'...")
    gallo["Referencia 2"] = gallo[col_ref].astype(str).str.strip().str[:2]

    # === Step 3: unique values of 'Referencia 2' (in order of appearance)
    refs_unicas = list(dict.fromkeys(
        ref for ref in gallo["Referencia 2"] if str(ref).strip() != ""
    ))

    # === Step 4: classify each reference as Global / Individual =======
    tipos = [_clasificar_tipo(ref) for ref in refs_unicas]
    mapa_tipo = dict(zip(refs_unicas, tipos))

    # === Steps 5 and 6: catalog of 'MX Commercial' documents ==========
    update_status("🔎 Filtrando 'MX Commercial'...")
    mascara_mx = gallo[col_asig].astype(str).str.strip() == VALOR_ASIGNACION
    catalogo_docs = gallo.loc[mascara_mx, [col_asig, col_doc]].copy()

    # === Step 7: 'Referencia 3' (Referencia 2 → Tipo), as VALUES =======
    update_status("🔁 Cruzando 'Referencia 3' y 'Comercializadora'...")
    gallo["Referencia 3"] = (
        gallo["Referencia 2"].map(mapa_tipo).fillna("")
    )

    # === Step 8: 'Comercializadora' (Nº documento → Asignación) ========
    # Like XLOOKUP: keeps the FIRST match of the catalog.
    mapa_comercializadora = {}
    for doc, asig in zip(
        catalogo_docs[col_doc].astype(str).str.strip(),
        catalogo_docs[col_asig].astype(str).str.strip(),
    ):
        mapa_comercializadora.setdefault(doc, asig)

    gallo["Comercializadora"] = (
        gallo[col_doc].astype(str).str.strip().map(mapa_comercializadora).fillna("")
    )

    # === Step 9: final filter → global 'Centro' and 'Nº documento' =====
    # This block (which feeds G-H of 'Validación - Doc.') is frozen with the
    # ORIGINAL classification, BEFORE the SG/RG override below. That is on
    # purpose (stakeholder request): the Step 9 globals keep the SG/RG
    # 'MX Commercial' rows that came in as 'Global', even though those same
    # rows end up as 'Individual' in the Gallo sheet.
    update_status("🧮 Aplicando filtro final (Global + MX Commercial)...")
    mascara_final = (
        (gallo["Comercializadora"] == VALOR_ASIGNACION)
        & (gallo["Referencia 3"] == "Global")
    )
    globales = gallo.loc[mascara_final, [col_centro, col_doc]].copy()

    # === SG/RG override → 'Individual' ("step 8.5", runs AFTER Step 9) ===
    # Business override: for MX Commercial rows whose Referencia 2 is SG or RG,
    # the amount is settled individually, so Referencia 3 becomes 'Individual'
    # instead of the default 'Global' that SG/RG get in _clasificar_tipo.
    # Applied AFTER the Step 9 snapshot, so it does NOT touch the G-H block; it
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

    # === Amount: ONE single parsing pass for all the matrices ==========
    # A 3-day range is ~1.15M rows: parsing the same column twice would waste
    # time, and a single pass also guarantees that 'TD' and the 'Venta ...'
    # sheets see exactly the same numbers.
    update_status("🔢 Convirtiendo 'Importe en moneda local' a número...")
    importe_num = gallo[col_importe].map(partial(parse_importe_us, columna=col_importe))

    return {
        "col_centro":           col_centro,
        "col_asig":             col_asig,
        "col_doc":              col_doc,
        "col_clase":            col_clase,
        "col_fecha":            col_fecha,
        "col_cuenta":           col_cuenta,
        "refs_unicas":          refs_unicas,
        "tipos":                tipos,
        "catalogo_docs":        catalogo_docs,
        "globales":             globales,
        "filas_reclasificadas": filas_reclasificadas,
        "importe_num":          importe_num,
    }


def escribir_hoja_validacion(ws, pasos_gallo):
    """
    Write the 'Validación - Doc.' sheet. It is built cell by cell because it
    holds three independent blocks side by side:

        A-B → unique 'Referencia 2' values and their type   (steps 3 and 4)
        D-E → documents whose Asignación is 'MX Commercial' (step 6)
        G-H → Centro and Nº documento of the globals        (step 9)

    Args:
        ws: the new (empty) openpyxl worksheet to fill.
        pasos_gallo (dict): what `preparar_gallo` returned.
    """
    refs_unicas   = pasos_gallo["refs_unicas"]
    tipos         = pasos_gallo["tipos"]
    catalogo_docs = pasos_gallo["catalogo_docs"]
    globales      = pasos_gallo["globales"]
    col_asig      = pasos_gallo["col_asig"]
    col_doc       = pasos_gallo["col_doc"]
    col_centro    = pasos_gallo["col_centro"]

    encabezados = {
        "A1": "Referencia 2", "B1": "Tipo",
        "D1": "Asignación",   "E1": "Nº documento",
        "G1": "Centro",       "H1": "Nº documento",
    }
    for celda, texto in encabezados.items():
        ws[celda] = texto
        ws[celda].font = Font(bold=True)

    # Block A-B: catalog of references and their type (steps 3 and 4)
    for fila, (ref, tipo) in enumerate(zip(refs_unicas, tipos), start=2):
        ws.cell(row=fila, column=1, value=ref)
        ws.cell(row=fila, column=2, value=tipo)

    # Block D-E: documents whose Asignación is 'MX Commercial' (step 6)
    for fila, (asig, doc) in enumerate(
        zip(catalogo_docs[col_asig], catalogo_docs[col_doc]), start=2
    ):
        ws.cell(row=fila, column=4, value=asig)
        ws.cell(row=fila, column=5, value=doc)

    # Block G-H: Centro and Nº documento of the globals (step 9)
    for fila, (centro, doc) in enumerate(
        zip(globales[col_centro], globales[col_doc]), start=2
    ):
        ws.cell(row=fila, column=7, value=centro)
        ws.cell(row=fila, column=8, value=doc)

    # Comfortable widths to read the catalog
    for letra, ancho in {"A": 14, "B": 12, "D": 16, "E": 16,
                         "G": 10, "H": 16}.items():
        ws.column_dimensions[letra].width = ancho


# ----------------------------------------------------------------------
# 'TD' matrix — computation and writing
# ----------------------------------------------------------------------
def preparar_matriz_td(gallo, col_centro, col_asig, col_clase, importe_num):
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

    All three blocks also require a non-blank Centro: a row without one has no
    place on the matrix, so it stays out of 'TD' (unlike 'MTD', which shows
    them as 'Sin centro'). Those rows are NOT dropped silently: they are
    counted, their amount is added up and reported in 'avisos', so a
    'Grand Total' that falls short can always be explained.

    Args:
        importe_num (pd.Series): 'Importe en moneda local' ALREADY parsed to
            float by the caller (single pass shared with the 'Venta ...' sheets).

    Returns a dict with one entry per block ('facturas', 'notas', 'pep'), each
    holding its own 'cedis' list and 'suma' lookup, plus the control totals,
    the grand total of the whole 'Importe en moneda local' column and 'avisos'
    (list[str]) with the rows left out for having no Centro.
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
    # The same filters for the rows WITHOUT a centro: they cannot be placed on
    # the matrix, so they end up in 'avisos' instead of vanishing.
    base_sin_centro = es_global & (asig_norm != VALOR_ASIGNACION) & ~con_centro

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
    # (the stakeholder asked for those three filters only).
    filtro_pep = (
        es_global
        & (ref2_norm == TD_PEP_REFERENCIA2)
        & (clase_norm == TD_PEP_CLASE)
    )
    mascara_pep = filtro_pep & con_centro
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

    # Rows that pass a block's filters but come WITHOUT a Centro. They are
    # reported per block (count + amount) so the stakeholder knows exactly how
    # much is missing from that block's 'Grand Total'.
    clases_facturas = {clase for clase, _ in TD_FACTURAS["clases"]}
    clases_notas    = {clase for clase, _ in TD_NOTAS["clases"]}
    avisos = []
    for titulo, mascara in (
        (TD_FACTURAS["titulo"], base_sin_centro & clase_norm.isin(clases_facturas)),
        (TD_NOTAS["titulo"],    base_sin_centro & clase_norm.isin(clases_notas)),
        (TD_PEP["titulo"],      filtro_pep & ~con_centro),
    ):
        n_filas = int(mascara.sum())
        if n_filas:
            importe = round(float(importe_num[mascara].sum()), 2)
            avisos.append(
                f"TD ({titulo}): {n_filas:,} fila(s) sin Centro (importe "
                f"{importe:,.2f}) no caben en la matriz; el "
                f"'{ETIQUETA_GRAND_TOTAL}' de ese bloque no las incluye."
            )

    return {
        "facturas":            facturas,
        "notas":               notas,
        "pep":                 {"cedis": cedis_pep, "suma": suma_pep},
        "total_facturas":      round(float(sum(facturas["suma"].values())), 2),
        "total_notas":         round(float(sum(notas["suma"].values())), 2),
        "total_pep":           round(float(sum(suma_pep.values())), 2),
        "total_importe_gallo": round(float(importe_num.sum()), 2),
        "avisos":              avisos,
        # Names the columns REALLY have in this export (they accept spelling
        # variants: 'Centro'/'Cebe', 'Clase de documento'/…). The filter legend
        # cites them so the stakeholder reads the same header the sheet shows.
        "columnas": {
            "centro":      col_centro,
            "asignacion":  col_asig,
            "clase":       col_clase,
        },
    }


def _leyenda_filtros_td(bloque, columnas):
    """
    Filter legend of a two-class 'TD' block (Facturas / Notas de Crédito),
    built FROM that block's own constants: change TD_FACTURAS["clases"] or
    VALOR_ASIGNACION and the legend follows, so what the sheet claims and what
    the code sums cannot drift apart.

    `columnas` carries the resolved column names of this export (see
    preparar_matriz_td). 'Referencia 3' is NOT resolved: that column does not
    come from SAP, this script creates it in `preparar_gallo` with exactly
    that name.
    """
    clases = " o ".join(f"'{clase}'" for clase, _ in bloque["clases"])
    return [
        f"Filtros — hoja '{HOJA_GALLO}':",
        f"Referencia 3 = 'Global'; {columnas['centro']} no vacío",
        f"{columnas['asignacion']} ≠ '{VALOR_ASIGNACION}'; "
        f"{columnas['clase']} = {clases}",
    ]


def _leyenda_filtros_td_pep(columnas):
    """
    Filter legend of the PEP vouchers block. Same idea as `_leyenda_filtros_td`,
    with the two differences the block really has: it filters by
    'Referencia 2' (created by this script, hence not resolved) instead of the
    Asignación, and it does NOT exclude VALOR_ASIGNACION — which is spelled out
    so nobody assumes the exclusion is there.
    """
    return [
        f"Filtros — hoja '{HOJA_GALLO}':",
        f"Referencia 3 = 'Global'; {columnas['centro']} no vacío",
        f"Referencia 2 = '{TD_PEP_REFERENCIA2}'; "
        f"{columnas['clase']} = '{TD_PEP_CLASE}' "
        f"(sí incluye '{VALOR_ASIGNACION}')",
    ]


def escribir_matriz_td(ws_td, td_resumen):
    """
    Write the three matrices onto the 'TD' worksheet.

    Layout (rows 1-3 = filter legend + stakeholder labels, row 4 = block title,
    row 5 = headers, row 6+ = data, and a 'Grand Total' row right below the
    last cedis):

        Facturas         → A: cedis, B: sum(C1), C: sum(C5), D: Grand Total (B+C)
        Notas de Crédito → L: cedis, M: sum(C2), N: sum(C6), O: Grand Total (M+N)
        PEP vouchers     → W: cedis, X: sum(C2 with Referencia 2 == 'PP')

    Extras:
        - Filter legend above each matrix (A1:A3, L1:L3, W1:W3): the exact
          conditions its rows meet, so any figure can be reproduced by hand
          from the 'Gallo' sheet.
        - 'Grand Total' column per block (D / O): per-row sum of its two classes.
        - 'Grand Total' row at the bottom: per-column sum; the corner cell is the
          block's overall total.
        - Analysis labels 'Monitor' / 'Contabilidad' / 'Diferencias' in D1:D3 and
          O1:O3 for the stakeholder's later work (labels only, no values).

    Each block owns its cedis list (derived from its OWN filters), so the blocks
    can differ in which centros they show and in how many rows they have — every
    block therefore computes its own 'Grand Total' row position. Body cells are
    real numbers (values, not formulas) with a thousands/2-decimals format.

    The cross against 'MTD' (live formulas in E:G, P:R and Y:AA) is written
    afterwards by `escribir_cruce_td_mtd` (Consolidacion_Monivoi). The
    bottom-row label MUST stay ETIQUETA_GRAND_TOTAL: that cross looks it up in
    'MTD' to compare totals.
    """
    # Column names of THIS export, for the filter legends.
    columnas = td_resumen["columnas"]

    for bloque, datos in ((TD_FACTURAS, td_resumen["facturas"]),
                          (TD_NOTAS,    td_resumen["notas"])):
        cedis      = datos["cedis"]
        suma       = datos["suma"]
        n_cedis    = len(cedis)
        fila_total = TD_FILA_INICIO + n_cedis   # row right below the last cedis
        col_cedis  = bloque["col_cedis"]
        col_gt     = bloque["col_grand_total"]
        clases     = bloque["clases"]

        # Filter legend (rows 1-3, in the cedis column) and analysis labels for
        # the stakeholder (same rows, in the Grand Total column).
        escribir_filtros(ws_td, TD_FILA_FILTROS, col_cedis,
                         _leyenda_filtros_td(bloque, columnas))
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
            ws_td.cell(row=fila_total, column=col_cedis,
                       value=ETIQUETA_GRAND_TOTAL).font = Font(bold=True)
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

    escribir_filtros(ws_td, TD_FILA_FILTROS, col_cedis_pep,
                     _leyenda_filtros_td_pep(columnas))
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
        ws_td.cell(row=fila_total_pep, column=col_cedis_pep,
                   value=ETIQUETA_GRAND_TOTAL).font = Font(bold=True)
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
# 'Venta DSD' / 'Venta EIAP' sheets — computation and writing
# ----------------------------------------------------------------------
def preparar_ventas(gallo, col_clase, col_cuenta, col_fecha, importe_num, fechas_sel):
    """
    Build the two date × amount matrices ('Venta DSD' and 'Venta EIAP').

    Returns:
        dict: {'dsd': block, 'eiap': block, 'avisos': list[str]}, where each
        block holds filas [(date, amount)], sin_fecha_n, sin_fecha_importe,
        total and fuera_rango.
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

        # A requested account that does NOT exist in the file almost always
        # means that day had no postings... or that the SAP layout changed.
        # Better to warn than to deliver an empty matrix with no explanation.
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
            # Ingredients of the filter legend written above the matrix: this
            # block's OWN filters and the column names this export really uses
            # (they accept spelling variants). The text itself is built by
            # escribir_hoja_venta, so the legend cannot describe filters other
            # than the ones applied right here.
            "clases":            clases,
            "cuentas":           cuentas,
            "columnas": {
                "clase":  col_clase,
                "cuenta": col_cuenta,
                "fecha":  col_fecha,
            },
        }

    bloques["avisos"] = avisos
    return bloques


def _leyenda_filtros_venta(clases, cuentas, columnas):
    """
    Filter legend of a 'Venta ...' matrix, built FROM that block's own filter
    sets (VENTA_CLASES_* and VENTA_CUENTAS_*) and from the column names this
    export really uses, so the text cannot drift from what the code sums.

    Two things are spelled out because they are exactly what makes a manual
    check disagree with the sheet:
        - the EMPTY document class counts for DSD (Excel shows it as
          "(blanks)"), and
        - the account is compared without leading zeros or a trailing '.0', so
          '0004000010' and '4000010.0' are the same account as '4000010'
          (filtering by the literal text in Excel would leave those rows out).
    """
    lista_clases = " o ".join(
        [f"'{clase}'" for clase in sorted(clase for clase in clases if clase)]
        + ([VENTA_ETIQUETA_EN_BLANCO] if "" in clases else [])
    )
    lista_cuentas = " o ".join(f"'{cuenta}'" for cuenta in sorted(cuentas))
    return [
        f"Filtros — hoja '{HOJA_GALLO}':",
        f"{columnas['clase']} = {lista_clases}",
        f"{columnas['cuenta']} = {lista_cuentas} (sin ceros a la izquierda)",
        f"{columnas['fecha']} vacía → renglón '{VENTA_ETIQUETA_SIN_FECHA}'",
    ]


def escribir_hoja_venta(ws, bloque):
    """
    Write one 'Venta ...' sheet: column A = date, column B = summed amount.

    Rows 1-4 = filter legend, row 5 = headers, row 6+ = one row per day, then
    (only if it exists) a 'Sin fecha' row, and finally a bold 'Grand Total'
    row. Dates are written as REAL dates formatted DD.MM.YYYY, so the
    stakeholder can sort and filter them as dates; amounts are values (not
    formulas) with 2 decimals.

    The legend states the TWO filters this matrix applies at the same time
    (document class AND account), which is the pair nobody remembers by heart.
    """
    escribir_filtros(ws, VENTA_FILA_FILTROS, 1,
                     _leyenda_filtros_venta(bloque["clases"], bloque["cuentas"],
                                            bloque["columnas"]))

    ws.cell(row=VENTA_FILA_ENCABEZADO, column=1,
            value=VENTA_ENCABEZADO_FECHA).font = Font(bold=True)
    ws.cell(row=VENTA_FILA_ENCABEZADO, column=2,
            value=VENTA_ENCABEZADO_IMPORTE).font = Font(bold=True)

    fila = VENTA_FILA_INICIO
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

    if fila > VENTA_FILA_INICIO:   # there is at least one data row
        ws.cell(row=fila, column=1, value=VENTA_ETIQUETA_TOTAL).font = Font(bold=True)
        celda_total = ws.cell(row=fila, column=2, value=bloque["total"])
        celda_total.number_format = TD_FORMATO_NUMERO
        celda_total.font = Font(bold=True)

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 24
