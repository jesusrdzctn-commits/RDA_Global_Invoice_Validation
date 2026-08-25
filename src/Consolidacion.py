"""
Validación Factura Global — Consolidación
==========================================

Toma los CSV descargados de SAP ('Gallo' FAGLL03 y 'Monivoi' ZLMXCOM) y arma
UN solo Excel con tres pestañas:

    1) 'Gallo'            → el CSV de Gallo + 3 columnas nuevas (K, L, M)
    2) 'Monivoi'          → el CSV de Monivoi tal cual
    3) 'Validación - Doc.'→ catálogo construido a partir de la Gallo

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

NOTA: los cruces de los pasos 7 y 8 se calculan en Python y se pegan como
VALORES (no fórmulas vivas), para que el archivo abra rápido y sin #N/A.

Este módulo NO depende de SAP ni de tkinter: se puede probar solo con
`python Consolidacion.py` teniendo dos CSV a la mano.
"""

import os

import pandas as pd
from openpyxl.styles import Font

# ----------------------------------------------------------------------
# Constantes del proceso
# ----------------------------------------------------------------------
VALOR_ASIGNACION = "MX Commercial"          # filtro de los pasos 5 y 9
PREFIJOS_LETRAS_GLOBAL = {"PP", "RG", "SG"}  # letras que también son 'Global'

HOJA_GALLO      = "Gallo"
HOJA_MONIVOI    = "Monivoi"
HOJA_VALIDACION = "Validación - Doc."


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
    Ejecuta la consolidación completa (pasos 1 a 9).

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
            filas_catalogo, filas_globales, monivoi_sin_datos
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
    update_status("🧮 Aplicando filtro final (Global + MX Commercial)...")
    mascara_final = (
        (gallo["Comercializadora"] == VALOR_ASIGNACION)
        & (gallo["Referencia 3"] == "Global")
    )
    globales = gallo.loc[mascara_final, [col_centro, col_doc]].copy()

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

    update_status("✅ Consolidación completada")

    return {
        "ruta":              ruta_final,
        "filas_gallo":       len(gallo),
        "filas_monivoi":     0 if monivoi_sin_datos else len(monivoi),
        "refs_unicas":       len(refs_unicas),
        "filas_catalogo":    len(catalogo_docs),
        "filas_globales":    len(globales),
        "monivoi_sin_datos": monivoi_sin_datos,
    }


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
            print(f"  {clave:<18}: {valor}")
    except Exception as e:
        print(f"❌ Error durante la consolidación: {e}")
        raise
