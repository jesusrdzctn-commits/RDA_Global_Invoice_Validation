"""
Validación Factura Global — Consolidation (entry point)
=======================================================

Takes the CSVs downloaded from SAP ('Gallo' FAGLL03 and 'Monivoi' ZLMXCOM) and
builds ONE Excel workbook with seven sheets:

    1) 'Gallo'             → the Gallo CSV + 3 new columns (K, L, M)
    2) 'Monivoi'           → the Monivoi CSV + 'Referencia 2' as column A
    3) 'Validación - Doc.' → catalog built from the Gallo
    4) 'TD'                → pivot-style matrices ("Tabla Dinámica") with the
                             sum of amount by cedis × document class (Facturas,
                             Notas de Crédito and PEP vouchers) + cross against
                             'MTD' with LIVE Excel formulas (Factura Monitor /
                             Diferencia / Observación)
    5) 'MTD'               → four matrices of summed 'Valor neto' by cedis,
                             taken from the MONIVOI (SG, RG, PP and TODOS)
    6) 'Venta DSD'         → date × summed amount, classes C1/C2/C5/C6/X0 and
                             blank class
    7) 'Venta EIAP'        → date × summed amount, class RV

How the code is organized
-------------------------
This file only ORCHESTRATES: it reads both CSVs, calls every step in order
and writes the workbook, so the whole end-to-end flow fits in one function
(consolidar_gallo_monivoi). The actual work lives in:

    Consolidacion_Comun.py    → shared helpers: read CSV, find columns, parse
                                amounts, labels/formats used by several sheets
    Consolidacion_Gallo.py    → steps 2-9 + SG/RG override, 'Validación - Doc.',
                                'TD', 'Venta DSD' / 'Venta EIAP'
    Consolidacion_Monivoi.py  → 'MTD' (step 16), 'Referencia 2' column of the
                                Monivoi sheet, TD × MTD cross

Imports only go one way (Comun ← Gallo ← Monivoi ← this file), so there are
no import cycles.

This module does NOT depend on SAP or tkinter: it can be tested on its own
with `python Consolidacion.py` given two CSVs.
"""

import os

import pandas as pd

from Consolidacion_Comun import es_sin_datos, leer_csv
from Consolidacion_Gallo import (
    HOJA_GALLO,
    HOJA_TD,
    HOJA_VALIDACION,
    HOJA_VENTA_DSD,
    HOJA_VENTA_EIAP,
    escribir_hoja_validacion,
    escribir_hoja_venta,
    escribir_matriz_td,
    parsear_fechas_seleccionadas,
    preparar_gallo,
    preparar_matriz_td,
    preparar_ventas,
)
from Consolidacion_Monivoi import (
    HOJA_MONIVOI,
    HOJA_MTD,
    avisos_cruce_td_mtd,
    escribir_cruce_td_mtd,
    escribir_hoja_mtd,
    insertar_referencia2,
    preparar_matrices_mtd,
)

# ----------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------
def carpeta_output_desde_input(ruta_input):
    """
    Return the 'Output' folder that goes with the download folder.

    If the chosen folder is called 'Input' (the project's standard layout:
    src/Input), Output sits NEXT to it (src/Output). If the stakeholder picked
    any other folder (e.g. a SharePoint one), Output is created INSIDE that
    folder so no files get scattered outside their choice.
    """
    ruta = os.path.normpath(ruta_input)
    if os.path.basename(ruta).lower() == "input":
        return os.path.join(os.path.dirname(ruta), "Output")
    return os.path.join(ruta, "Output")


# ----------------------------------------------------------------------
# Main process (orchestrator)
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
    Run the full consolidation (steps 1-9 + SG/RG override + 'TD' matrix +
    'MTD' sheet + TD × MTD cross + 'Venta DSD' and 'Venta EIAP' sheets).

    Args:
        ruta_gallo   (str): Path to the Gallo CSV downloaded from SAP (FAGLL03).
        ruta_monivoi (str): Path to the Monivoi CSV downloaded from SAP (ZLMXCOM).
        ruta_output  (str): Folder where the consolidated workbook is saved.
        nombre_salida (str|None): File name of the final .xlsx. If None →
            'Consolidado_<Gallo CSV name>.xlsx'.
        callback_status (function|None): Reports progress to the GUI.
        fechas_seleccionadas (list[str]|None): 'DD.MM.YYYY' dates the user
            picked in the GUI (one in single-day mode, every day of the range
            in range mode). They guarantee those days show up in 'Venta DSD'
            and 'Venta EIAP' even without amounts, and they break ties when
            detecting the CSV date format. If None, rows come from the data
            only.

    Returns:
        dict: process summary →
            ruta, filas_gallo, filas_monivoi, refs_unicas,
            filas_catalogo, filas_globales, monivoi_sin_datos,
            filas_reclasificadas, cedis_facturas, cedis_notas, cedis_pep,
            total_facturas, total_notas, total_pep, total_importe_gallo,
            td_avisos, total_venta_dsd, total_venta_eiap, venta_avisos,
            centros_mtd_todos, mtd_filas_sin_uuid, mtd_avisos, cruce_avisos

    Raises:
        FileNotFoundError: one of the CSVs does not exist.
        ValueError: input the user can fix (Gallo without data, a missing
            column, an amount or date that cannot be interpreted, a selected
            date that is not 'DD.MM.YYYY').
    """

    def update_status(mensaje):
        if callback_status:
            callback_status(mensaje)

    # Validated BEFORE reading anything: if the GUI sends a malformed date, it
    # is better to find out in the first second than after processing 1M rows.
    fechas_sel = parsear_fechas_seleccionadas(fechas_seleccionadas)

    # === Step 1: read both files ======================================
    update_status("📄 Leyendo archivo de Gallo...")
    gallo = leer_csv(ruta_gallo)
    if es_sin_datos(gallo):
        raise ValueError(
            "El archivo de Gallo no tiene datos (descarga 'sin partidas'). "
            "No hay nada que consolidar."
        )

    update_status("📄 Leyendo archivo de Monivoi...")
    monivoi = leer_csv(ruta_monivoi)
    monivoi_sin_datos = es_sin_datos(monivoi)

    # === Gallo: steps 2-9 + SG/RG override + amount parsing ===========
    # Adds columns K, L, M to `gallo` and returns what the next steps need.
    pasos_gallo = preparar_gallo(gallo, update_status)

    # === Gallo: 'TD' matrix — sum of amount by cedis × document class ===
    update_status("📊 Construyendo matriz 'TD'...")
    td_resumen = preparar_matriz_td(
        gallo,
        pasos_gallo["col_centro"],
        pasos_gallo["col_asig"],
        pasos_gallo["col_clase"],
        pasos_gallo["importe_num"],
    )
    total_importe_gallo = td_resumen["total_importe_gallo"]

    # === Monivoi: step 16 — 'MTD' matrices ============================
    update_status("🧾 Construyendo matrices 'MTD' (SG / RG / PP / TODOS)...")
    mtd = preparar_matrices_mtd(monivoi, monivoi_sin_datos)
    # Helper column 'Referencia 2' as column A of the Monivoi sheet. It goes
    # AFTER the matrices on purpose: they use their own series, not the sheet.
    insertar_referencia2(monivoi, mtd)

    # === TD × MTD cross: what the formulas cannot show on their own ===
    # (centros stamped in the monitor that have no row in TD).
    cruce_avisos = avisos_cruce_td_mtd(td_resumen, mtd)

    # === Gallo: 'Venta DSD' and 'Venta EIAP' — sum by date ============
    update_status("📆 Construyendo 'Venta DSD' y 'Venta EIAP'...")
    ventas = preparar_ventas(
        gallo,
        pasos_gallo["col_clase"],
        pasos_gallo["col_cuenta"],
        pasos_gallo["col_fecha"],
        pasos_gallo["importe_num"],
        fechas_sel,
    )

    # === Export the consolidated workbook ==============================
    update_status("💾 Guardando Excel consolidado...")
    os.makedirs(ruta_output, exist_ok=True)

    if not nombre_salida:
        base_gallo = os.path.splitext(os.path.basename(ruta_gallo))[0]
        nombre_salida = f"Consolidado_{base_gallo}.xlsx"
    ruta_final = os.path.join(ruta_output, nombre_salida)

    # The tab order of the workbook is the order of the writes below.
    with pd.ExcelWriter(ruta_final, engine="openpyxl") as writer:
        gallo.to_excel(writer, sheet_name=HOJA_GALLO, index=False)
        monivoi.to_excel(writer, sheet_name=HOJA_MONIVOI, index=False)

        # --- 'Validación - Doc.' sheet (Gallo catalog: blocks A-B, D-E, G-H) ---
        escribir_hoja_validacion(writer.book.create_sheet(HOJA_VALIDACION), pasos_gallo)

        # --- 'TD' sheet: Facturas, Notas de Crédito and PEP matrices ---
        ws_td = writer.book.create_sheet(HOJA_TD)
        escribir_matriz_td(ws_td, td_resumen)
        # Cross against 'MTD': live formulas in E:G, P:R and Y:AA.
        escribir_cruce_td_mtd(ws_td, td_resumen)
        # openpyxl saves formulas WITHOUT their result. This forces Excel to
        # calculate all of them when the file is opened, instead of showing
        # empty cells.
        writer.book.calculation.fullCalcOnLoad = True

        # --- 'MTD' sheet: four Monivoi matrices (step 16) ---
        ws_mtd = writer.book.create_sheet(HOJA_MTD)
        escribir_hoja_mtd(ws_mtd, mtd)

        # --- 'Venta DSD' and 'Venta EIAP' sheets ---
        for nombre_hoja, clave in ((HOJA_VENTA_DSD, "dsd"), (HOJA_VENTA_EIAP, "eiap")):
            escribir_hoja_venta(writer.book.create_sheet(nombre_hoja), ventas[clave])

    update_status("✅ Consolidación completada")

    return {
        "ruta":                 ruta_final,
        "filas_gallo":          len(gallo),
        "filas_monivoi":        0 if monivoi_sin_datos else len(monivoi),
        "refs_unicas":          len(pasos_gallo["refs_unicas"]),
        "filas_catalogo":       len(pasos_gallo["catalogo_docs"]),
        "filas_globales":       len(pasos_gallo["globales"]),
        "monivoi_sin_datos":    monivoi_sin_datos,
        "filas_reclasificadas": pasos_gallo["filas_reclasificadas"],
        "cedis_facturas":       len(td_resumen["facturas"]["cedis"]),
        "cedis_notas":          len(td_resumen["notas"]["cedis"]),
        "cedis_pep":            len(td_resumen["pep"]["cedis"]),
        "total_facturas":       td_resumen["total_facturas"],
        "total_notas":          td_resumen["total_notas"],
        "total_pep":            td_resumen["total_pep"],
        "total_importe_gallo":  total_importe_gallo,
        "td_avisos":            td_resumen["avisos"],
        "total_venta_dsd":      ventas["dsd"]["total"],
        "total_venta_eiap":     ventas["eiap"]["total"],
        "venta_avisos":         ventas["avisos"],
        "centros_mtd_todos":    len(mtd["bloques"]["todos"]["filas"]),
        "mtd_filas_sin_uuid":   mtd["filas_sin_uuid"],
        "mtd_avisos":           mtd["avisos"],
        "cruce_avisos":         cruce_avisos,
    }


# ----------------------------------------------------------------------
# Direct run (quick test without GUI or controller)
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
