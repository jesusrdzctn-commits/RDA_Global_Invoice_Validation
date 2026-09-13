"""
Validación Factura Global — Consolidación
==========================================

Toma los CSV descargados de SAP ('Gallo' FAGLL03 y 'Monivoi' ZLMXCOM) y arma
UN solo Excel con cuatro pestañas:

    1) 'Gallo'            → el CSV de Gallo + 3 columnas nuevas (K, L, M)
    2) 'Monivoi'          → el CSV de Monivoi tal cual
    3) 'Validación - Doc.'→ catálogo construido a partir de la Gallo
    4) 'TD'               → matrices (Tabla Dinámica) de suma de importe por
                            cedis × clase de documento (Facturas y Notas de Crédito)

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
    - Subconjunto base: Referencia 3=='Global' y Asignación != 'MX Commercial'.
    - Filas: TODOS los Centros ('cedis') únicos de la Gallo (no sólo los del
      subconjunto base), ordenados; un cedis sin actividad calificada sale en 0.
    - Facturas: encabezados B5='C1', C5='C5'; cedis desde A6 hacia abajo.
    - Notas de Crédito: encabezados M5='C2', N5='C6'; cedis desde L6 hacia abajo.
    - Cada celda del cuerpo = suma de 'Importe en moneda local' de los documentos
      de ese cedis y esa clase (dentro del subconjunto base).
    - Columna 'Grand Total' a la derecha de cada bloque (D en Facturas, O en
      Notas) = suma de las dos clases por cedis. Fila 'Grand Total' al final =
      suma por columna; la celda de cruce es el gran total del bloque.
    - Etiquetas de análisis para el stakeholder: 'Monitor', 'Contabilidad' y
      'Diferencias' en D1:D3 (Facturas) y O1:O3 (Notas). Sólo rótulos.
    - Tercera matriz 'PEP VOUCHERS' (columnas V/W): filtro Referencia 3=='Global',
      Referencia 2=='PP' y Clase de documento=='C2'. Suma de importe por cedis,
      encabezado 'C2' en W5, con su fila 'Grand Total'. NO excluye 'MX Commercial'.

NOTA: los cruces de los pasos 7 y 8 se calculan en Python y se pegan como
VALORES (no fórmulas vivas), para que el archivo abra rápido y sin #N/A. La
matriz 'TD' también se pega como valores numéricos ya sumados.

Este módulo NO depende de SAP ni de tkinter: se puede probar solo con
`python Consolidacion.py` teniendo dos CSV a la mano.
"""

import os

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
# Referencia 2=='PP' y Clase de documento=='C2'. Una sola columna de valor (W),
# con las filas de cedis en V. No excluye 'MX Commercial' (sólo esos 3 filtros).
TD_PEP = {
    "titulo": "PEP VOUCHERS",
    "col_cedis": 22,           # V
    "col_valor": 23,           # W
    "encabezado_valor": "C2",  # clase del filtro (va en W5)
}
TD_PEP_REFERENCIA2 = "PP"
TD_PEP_CLASE = "C2"

TD_FORMATO_NUMERO = "#,##0.00"


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
# Proceso principal
# ----------------------------------------------------------------------
def consolidar_gallo_monivoi(
    ruta_gallo,
    ruta_monivoi,
    ruta_output,
    nombre_salida=None,
    callback_status=None,
):
    """
    Ejecuta la consolidación completa (pasos 1 a 9 + reclasificación + matriz TD).

    Args:
        ruta_gallo   (str): Ruta al CSV descargado de Gallo (FAGLL03).
        ruta_monivoi (str): Ruta al CSV descargado de Monivoi (ZLMXCOM).
        ruta_output  (str): Carpeta donde se guardará el Excel consolidado.
        nombre_salida (str|None): Nombre del .xlsx final. Si None →
            'Consolidado_<nombre del CSV de Gallo>.xlsx'.
        callback_status (function|None): Para reportar avance a la GUI.

    Returns:
        dict: resumen del proceso →
            ruta, filas_gallo, filas_monivoi, refs_unicas,
            filas_catalogo, filas_globales, monivoi_sin_datos,
            filas_reclasificadas, cedis_matriz, total_facturas, total_notas,
            total_pep, total_importe_gallo
    """

    def update_status(mensaje):
        if callback_status:
            callback_status(mensaje)

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

    # Columnas que necesita la matriz 'TD'. Se resuelven ESTRICTAMENTE por
    # nombre (su posición en el export no es fija) y fallan pronto con un
    # error claro si no están, antes de procesar nada.
    col_clase   = _col_por_nombre(gallo, "Clase de documento")
    col_importe = _col_por_nombre(gallo, "Importe en moneda local")

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

    # === Matriz 'TD': suma de importe por cedis × clase de documento ===
    update_status("📊 Construyendo matriz 'TD'...")
    td_resumen = _preparar_matriz_td(gallo, col_centro, col_asig, col_clase, col_importe)

    # Suma total de TODA la columna 'Importe en moneda local' (para el popup).
    total_importe_gallo = round(float(gallo[col_importe].map(_parse_importe_us).sum()), 2)

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
        "cedis_matriz":         len(td_resumen["cedis"]),
        "total_facturas":       td_resumen["total_facturas"],
        "total_notas":          td_resumen["total_notas"],
        "total_pep":            td_resumen["total_pep"],
        "total_importe_gallo":  total_importe_gallo,
    }


# ----------------------------------------------------------------------
# Matriz 'TD' — cálculo y escritura
# ----------------------------------------------------------------------
def _preparar_matriz_td(gallo, col_centro, col_asig, col_clase, col_importe):
    """
    Build the aggregated data for the 'TD' sheet.

    Base subset (step II of the spec): Referencia 3 == 'Global' AND
    Asignación != 'MX Commercial'. Within that subset we sum
    'Importe en moneda local' grouped by (Centro, Clase de documento).

    Returns a dict:
        - 'cedis'  : sorted list of unique non-empty Centro values in the subset
                     (shared row labels for BOTH matrices)
        - 'suma'   : dict {(centro, clase_upper): importe_sumado}
        - 'total_facturas' / 'total_notas': control totals for the summary
    """
    # Row labels: ALL unique non-empty centros in the WHOLE Gallo (not only the
    # filtered subset), sorted. A cedis with no qualifying activity still shows
    # up as a row and simply displays zeros. Shared by the three matrices.
    todos_centros = gallo[col_centro].astype(str).str.strip()
    cedis = sorted(c for c in todos_centros.unique() if c != "")

    # --- Facturas / Notas subset: Referencia 3=='Global' & Asignación != MX ---
    mascara_td = (
        (gallo["Referencia 3"] == "Global")
        & (gallo[col_asig].astype(str).str.strip() != VALOR_ASIGNACION)
    )
    td = gallo.loc[mascara_td, [col_centro, col_clase, col_importe]].copy()
    td["_centro"]  = td[col_centro].astype(str).str.strip()
    td["_clase"]   = td[col_clase].astype(str).str.strip().str.upper()
    td["_importe"] = td[col_importe].map(_parse_importe_us)
    td = td[td["_centro"] != ""]

    # (centro, clase) -> summed amount. Round to cents so no float noise
    # (e.g. -265.44000000000005) leaks into the cells or the control totals.
    if td.empty:
        suma = {}
    else:
        serie = td.groupby(["_centro", "_clase"])["_importe"].sum()
        suma = {clave: round(float(valor), 2) for clave, valor in serie.items()}

    clases_facturas = {clase for clase, _ in TD_FACTURAS["clases"]}
    clases_notas    = {clase for clase, _ in TD_NOTAS["clases"]}
    total_facturas  = sum(v for (_, clase), v in suma.items() if clase in clases_facturas)
    total_notas     = sum(v for (_, clase), v in suma.items() if clase in clases_notas)

    # --- PEP vouchers subset: Referencia 3=='Global' & Referencia 2=='PP' &
    #     Clase de documento=='C2'. (No excluye 'MX Commercial': el stakeholder
    #     pidió sólo esos tres filtros.) Suma de importe por cedis. ---
    ref2       = gallo["Referencia 2"].astype(str).str.strip().str.upper()
    clase_norm = gallo[col_clase].astype(str).str.strip().str.upper()
    mascara_pep = (
        (gallo["Referencia 3"] == "Global")
        & (ref2 == TD_PEP_REFERENCIA2)
        & (clase_norm == TD_PEP_CLASE)
    )
    pep = gallo.loc[mascara_pep, [col_centro, col_importe]].copy()
    pep["_centro"]  = pep[col_centro].astype(str).str.strip()
    pep["_importe"] = pep[col_importe].map(_parse_importe_us)
    pep = pep[pep["_centro"] != ""]
    if pep.empty:
        suma_pep = {}
    else:
        serie_pep = pep.groupby("_centro")["_importe"].sum()
        suma_pep = {c: round(float(v), 2) for c, v in serie_pep.items()}

    return {
        "cedis":          cedis,
        "suma":           suma,
        "suma_pep":       suma_pep,
        "total_facturas": round(float(total_facturas), 2),
        "total_notas":    round(float(total_notas), 2),
        "total_pep":      round(float(sum(suma_pep.values())), 2),
    }


def _escribir_matriz_td(ws_td, td_resumen):
    """
    Write the two matrices onto the 'TD' worksheet.

    Layout (rows 1-3 = stakeholder labels, row 4 = block title, row 5 = headers,
    row 6+ = data, and a 'Grand Total' row right below the last cedis):

        Facturas         → A: cedis, B: sum(C1), C: sum(C5), D: Grand Total (B+C)
        Notas de Crédito → L: cedis, M: sum(C2), N: sum(C6), O: Grand Total (M+N)

    Extras:
        - 'Grand Total' column per block (D / O): per-row sum of its two classes.
        - 'Grand Total' row at the bottom: per-column sum; the corner cell is the
          block's overall total.
        - Analysis labels 'Monitor' / 'Contabilidad' / 'Diferencias' in D1:D3 and
          O1:O3 for the stakeholder's later work (labels only, no values).

    Both blocks share the SAME sorted cedis list, so a given row is the same
    centro on the left and the right. Body cells are real numbers (values, not
    formulas) with a thousands/2-decimals number format.
    """
    cedis = td_resumen["cedis"]
    suma  = td_resumen["suma"]
    n_cedis = len(cedis)
    fila_total = TD_FILA_INICIO + n_cedis   # row right below the last cedis

    for bloque in (TD_FACTURAS, TD_NOTAS):
        col_cedis = bloque["col_cedis"]
        col_gt    = bloque["col_grand_total"]
        clases    = bloque["clases"]

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

    # --- Third matrix: PEP vouchers (single value column, keyed by centro) ---
    col_cedis_pep = TD_PEP["col_cedis"]
    col_valor_pep = TD_PEP["col_valor"]
    suma_pep = td_resumen["suma_pep"]

    ws_td.cell(row=4, column=col_cedis_pep, value=TD_PEP["titulo"]).font = Font(bold=True)
    ws_td.cell(row=TD_FILA_ENCABEZADO, column=col_cedis_pep, value="Centro").font = Font(bold=True)
    ws_td.cell(row=TD_FILA_ENCABEZADO, column=col_valor_pep,
               value=TD_PEP["encabezado_valor"]).font = Font(bold=True)

    total_pep = 0.0
    for i, centro in enumerate(cedis):
        fila = TD_FILA_INICIO + i
        ws_td.cell(row=fila, column=col_cedis_pep, value=centro)
        valor = round(suma_pep.get(centro, 0.0), 2)
        ws_td.cell(row=fila, column=col_valor_pep, value=valor).number_format = TD_FORMATO_NUMERO
        total_pep += valor
    if n_cedis > 0:
        ws_td.cell(row=fila_total, column=col_cedis_pep, value="Grand Total").font = Font(bold=True)
        celda_pep = ws_td.cell(row=fila_total, column=col_valor_pep, value=round(total_pep, 2))
        celda_pep.number_format = TD_FORMATO_NUMERO
        celda_pep.font = Font(bold=True)

    # Comfortable widths (incl. the Grand Total columns D/O and the PEP V/W).
    anchos = {"A": 12, "B": 16, "C": 16, "D": 16,
              "L": 12, "M": 16, "N": 16, "O": 16,
              "V": 12, "W": 16}
    for letra, ancho in anchos.items():
        ws_td.column_dimensions[letra].width = ancho


# ----------------------------------------------------------------------
# Ejecución directa (prueba rápida sin GUI ni controller)
# ----------------------------------------------------------------------
if __name__ == "__main__":
    print("=== Prueba de consolidación (standalone) ===")
    ruta_gallo   = input("Ruta del CSV de Gallo   : ").strip().strip('"')
    ruta_monivoi = input("Ruta del CSV de Monivoi : ").strip().strip('"')

    salida = carpeta_output_desde_input(os.path.dirname(ruta_gallo))

    try:
        resumen = consolidar_gallo_monivoi(
            ruta_gallo, ruta_monivoi, salida, callback_status=print
        )
        print("\n🎉 CONSOLIDACIÓN COMPLETADA 🎉")
        for clave, valor in resumen.items():
            print(f"  {clave:<20}: {valor}")
    except Exception as e:
        print(f"❌ Error durante la consolidación: {e}")
        raise
