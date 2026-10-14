"""
Validación Factura Global — Consolidation
=========================================

Takes the CSVs downloaded from SAP ('Gallo' FAGLL03, 'Monivoi' ZLMXCOM and,
optionally, 'VBRL' ZSE16) and builds ONE Excel workbook. Tab order:

    1) 'Gallo'                      → the Gallo CSV + 3 new columns (K, L, M)
    2) 'Monivoi'                    → the Monivoi CSV + 'Referencia 2' as column A
    3) 'VBRL'                       → the VBRL CSV, columns D and E turned into
                                      numbers                       [only with VBRL]
    4) 'Validación - Doc.'          → catalog built from the Gallo
    5) 'TD'                         → pivot-style matrices ("Tabla Dinámica") with
                                      the sum of amount by cedis × document class
                                      (Facturas, Notas de Crédito and PEP vouchers)
                                      + cross against 'MTD' with LIVE Excel formulas
                                      (Factura Monitor / Diferencia / Observación)
    6) 'TD.Diferencias_Referencias' → the same Facturas population as 'TD', but down
                                      to Centro × Nº documento × Referencia de pago,
                                      crossed against 'VBRL' with LIVE Excel
                                      formulas (Factura Monitor / Diferencia)
                                                                    [only with VBRL]
    7) 'MTD'                        → four matrices of summed 'Valor neto' by cedis,
                                      taken from the MONIVOI (SG, RG, PP and TODOS)
    8) 'Venta DSD'                  → date × summed amount, classes C1/C2/C5/C6/X0
                                      and blank class

How the code is organized
-------------------------
ONE file, ONE CLASS PER STAGE of the consolidation (the same idea as the four
buttons of the P2P tool: every stage owns its constants, its computation and
the sheet it writes). Every stage class has the same shape:

    Stage(inputs).preparar()   → computes everything and returns its result dict
    Stage.escribir(ws, ...)    → writes that result onto a worksheet

    Comun               shared helpers and labels (read CSV, find columns, parse
                        amounts, filter legend, formula helpers)
    EtapaGallo          steps 2-9 + SG/RG override → helper columns of 'Gallo'
                        and the 'Validación - Doc.' sheet
    EtapaTD             'TD' matrices (Facturas, Notas de Crédito, PEP vouchers)
    EtapaVentas         'Venta DSD' (+ the date handling it needs)
    EtapaMonivoi        'MTD' (step 16) + 'Referencia 2' column of the Monivoi
                        sheet + the invoice list that feeds the VBRL download
    EtapaCruceTDMTD     live formulas that cross 'TD' against 'MTD'
    EtapaVBRL           'VBRL' sheet: what the monitor invoiced, per invoice
    EtapaTDReferencias  'TD.Diferencias_Referencias': accounting vs VBRL, document
                        by document
    Consolidador        the ORCHESTRATOR: reads the files, runs the stages in
                        order and writes the workbook

`consolidar_gallo_monivoi(...)` (bottom of the file) is the function the
controller calls; it just builds a `Consolidador` and runs it.

The classes are defined from the most basic one to the most dependent one
(Comun ← Gallo ← TD ← ... ← Consolidador), so a class only ever uses classes
that are above it.

This module does NOT depend on SAP or tkinter: it can be tested on its own with
`python Consolidacion.py` given two (or three) CSVs.
"""

import os
import re
import unicodedata
from datetime import datetime
from functools import partial

import pandas as pd
from openpyxl.comments import Comment
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ======================================================================
# Shared helpers
# ======================================================================
class Comun:
    """
    Building blocks used by EVERY stage. Nothing in here knows about a specific
    sheet or business rule.

        leer_csv / es_sin_datos   → read a CSV downloaded by the app
        col_por_candidatos        → find a column, accepting spelling variants
        parse_importe_us          → turn a US-formatted SAP amount into a float
        escribir_filtros          → write the filter legend above a matrix
        ref_hoja / texto_formula  → build safe pieces of an Excel formula
        ETIQUETA_GRAND_TOTAL, FORMATO_NUMERO, ENCABEZADO_IMPORTE, TEXTO_SIN_TIMBRE
                                  → labels/formats shared by several sheets
    """

    # Label of the totals row in 'TD' and in 'MTD'. It MUST be the same on both
    # sheets: the TD × MTD cross looks it up with XLOOKUP to compare total
    # against total. That is why it lives in ONE single constant.
    ETIQUETA_GRAND_TOTAL = "Grand Total"

    # Number format of every amount cell ('TD', 'MTD', 'Venta DSD' and the
    # crosses).
    FORMATO_NUMERO = "#,##0.00"

    # Header of Gallo's amount column. It is also the column-B header of the
    # 'Venta DSD' sheet and the default column label in parse_importe_us error
    # messages.
    ENCABEZADO_IMPORTE = "Importe en moneda local"

    # Look of the filter legend written above every matrix: italic and gray, so
    # it reads as a note and never as data.
    FILTROS_COLOR = "FF595959"

    # What a cross shows when the monitor has nothing to compare against
    # (used by the TD × MTD cross and by 'TD.Diferencias_Referencias').
    TEXTO_SIN_TIMBRE = "Sin timbre"

    # A US amount with a thousands separator ALWAYS groups digits in threes
    # ('1,234.56'). Anything else with a comma (the European '1.234,56', or a
    # '12,5') is NOT US: blindly dropping the comma would turn 1.234,56 into
    # 1.23456 and silently understate the sum. That is why it is validated
    # before cleaning.
    _RE_IMPORTE_US_AGRUPADO = re.compile(r"\d{1,3}(?:,\d{3})+(?:\.\d+)?")

    # A sheet name made ONLY of letters (and '_') can go unquoted in a formula
    # ('MTD!A:A'); any other ('Venta DSD', 'Hoja1') needs single quotes. So do
    # R, C, TRUE and FALSE, because Excel mistakes them for something else.
    _RE_HOJA_SIN_COMILLAS = re.compile(r"[A-Za-z_]+")
    _HOJAS_RESERVADAS = frozenset({"R", "C", "TRUE", "FALSE"})

    # ------------------------------------------------------------------
    # Reading the downloaded CSVs
    # ------------------------------------------------------------------
    @staticmethod
    def leer_csv(ruta):
        """Read a CSV downloaded by the app: everything as text, accents included."""
        if not os.path.exists(ruta):
            raise FileNotFoundError(f"Archivo no encontrado: {ruta}")
        # keep_default_na=False → empty cells stay as "" (not NaN), so .strip()
        # and text comparisons never blow up.
        return pd.read_csv(ruta, dtype=str, encoding="utf-8-sig", keep_default_na=False)

    @staticmethod
    def es_sin_datos(df):
        """True if the CSV is the 'Sin datos' placeholder written by the download."""
        return df.empty or (len(df.columns) == 1 and "Sin datos" in str(df.columns[0]))

    # ------------------------------------------------------------------
    # Column lookup
    # ------------------------------------------------------------------
    @staticmethod
    def normalizar_encabezado(texto):
        """
        Normalize a header so it can be compared ignoring accents, dots, spaces
        and case: 'Fe.Contabilización' and 'fecha contabilizacion' become
        comparable. Only used to recognize SPELLING VARIANTS of the same field,
        never to guess a different column.
        """
        sin_acentos = unicodedata.normalize("NFKD", str(texto))
        sin_acentos = "".join(c for c in sin_acentos if not unicodedata.combining(c))
        return "".join(c for c in sin_acentos.lower() if c.isalnum())

    @staticmethod
    def col_por_candidatos(df, candidatos, descripcion, origen="Gallo"):
        """
        Resolve a column accepting several spellings of the SAME field
        ('Fe.contabilización' / 'Fecha contabilización', 'Cuentas' / 'Cuenta de
        mayor'). The comparison ignores accents, dots, spaces and case.

        `origen` is only used in the error message ('Gallo', 'Monivoi' or
        'VBRL'), so the stakeholder knows WHICH file to check.

        If none of them is present, it fails with an actionable message that
        lists the real columns: stopping is better than filtering/grouping by
        the wrong column.
        """
        normalizadas = {Comun.normalizar_encabezado(c): c for c in df.columns}
        for candidata in candidatos:
            real = normalizadas.get(Comun.normalizar_encabezado(candidata))
            if real is not None:
                return real
        raise ValueError(
            f"No encontré la columna de {descripcion} en el archivo de {origen} "
            f"(busqué: {', '.join(candidatos)}). "
            f"Columnas disponibles: {list(df.columns)}"
        )

    # ------------------------------------------------------------------
    # Amounts
    # ------------------------------------------------------------------
    @staticmethod
    def parse_importe_us(valor, columna=ENCABEZADO_IMPORTE):
        """
        Convert a US-formatted SAP amount string into a float.

        Handles the shapes SAP actually exports (US locale):
            - thousands separator ','  and decimal point '.'   -> '1,234.56'
            - SAP trailing minus (sign after the number)        -> '1,234.56-'
            - leading minus / stray leading '+'                 -> '-1,234.56'
            - accounting parentheses                            -> '(1,234.56)'
            - blank / whitespace / lone '-'                     -> 0.0

        A non-empty token that is still not numeric after cleanup raises
        ValueError, so corrupt amounts stop the run instead of quietly
        understating a sum. Same for a comma that is NOT a US thousands
        separator (e.g. the European '1.234,56'): dropping it would turn the
        amount into 1.23456.

        `columna` is only used in the error message, so it is clear whether the
        odd amount came from 'Importe en moneda local' (Gallo), 'Valor neto'
        (Monivoi) or the amount column of the VBRL.
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

        # US format: ',' is ONLY a thousands separator, and always groups 3
        # digits. Validate BEFORE stripping it, so a European-formatted amount
        # fails loud instead of silently becoming a different (smaller) number.
        if "," in s:
            if not Comun._RE_IMPORTE_US_AGRUPADO.fullmatch(s):
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

    # ------------------------------------------------------------------
    # Filter legend shown above each matrix
    # ------------------------------------------------------------------
    @staticmethod
    def escribir_filtros(ws, fila_inicial, columna, filtros):
        """
        Write the filter legend of a matrix, one line per row, starting at
        (fila_inicial, columna) and going down.

        It is what lets the stakeholder reproduce any figure by hand: it names
        the source sheet and the exact conditions the rows of that matrix meet.
        Each caller builds its lines FROM ITS OWN filter constants, so the
        legend can never drift from what the code really sums.

        The text is italic and gray so nobody mistakes it for data, and it is
        written in the block's first column, on rows that the layout keeps free.
        """
        for i, texto in enumerate(filtros):
            celda = ws.cell(row=fila_inicial + i, column=columna, value=texto)
            celda.font = Font(italic=True, color=Comun.FILTROS_COLOR)

    # ------------------------------------------------------------------
    # Pieces of Excel formulas
    # ------------------------------------------------------------------
    @staticmethod
    def ref_hoja(nombre):
        """Sheet name as a formula reference: bare when allowed, quoted otherwise."""
        if (Comun._RE_HOJA_SIN_COMILLAS.fullmatch(nombre)
                and nombre.upper() not in Comun._HOJAS_RESERVADAS):
            return nombre
        return "'" + nombre.replace("'", "''") + "'"

    @staticmethod
    def texto_formula(texto):
        """Text literal for a formula: 'Sin timbre' -> \"Sin timbre\"."""
        return '"' + str(texto).replace('"', '""') + '"'


# ======================================================================
# Stage: Gallo (FAGLL03) — steps 2-9 + SG/RG override + 'Validación - Doc.'
# ======================================================================
class EtapaGallo:
    """
    Everything that is computed FROM the Gallo CSV before any matrix is built:

        1) Steps 2-9 + SG/RG override → helper columns K, L, M of the 'Gallo'
           sheet ('Referencia 2', 'Referencia 3', 'Comercializadora')
                                                          → preparar
        2) 'Validación - Doc.' sheet                      → escribir_validacion

    It also resolves, ONCE, the names of every Gallo column the later stages
    need (they are returned by `preparar`), so a missing column stops the run
    in the first second instead of after processing 1M rows.
    """

    # Sheet names (tabs) built from the Gallo.
    HOJA_GALLO      = "Gallo"
    HOJA_VALIDACION = "Validación - Doc."

    # Process constants.
    VALOR_ASIGNACION       = "MX Commercial"             # filter of steps 5, 8.5 and 9
    PREFIJOS_LETRAS_GLOBAL = frozenset({"PP", "RG", "SG"})  # letter-only prefixes that are also 'Global'
    PREFIJOS_RECLASIFICAR  = frozenset({"SG", "RG"})        # step 8.5: SG/RG 'MX Commercial' → Individual

    # Gallo columns whose header may be spelled in several ways. (Their
    # position in the export is not fixed, so they are resolved by NAME.)
    COL_FECHA_CANDIDATAS = (
        "Fe.contabilización",
        "Fecha contabilización",
        "Fecha de contabilización",
        "Fecha contab.",
    )
    COL_CUENTA_CANDIDATAS = (
        "Cuentas",
        "Cuenta",
        "Cuenta de mayor",
        "Cuenta contable",
        "Cta.mayor",
        "Nº cuenta",
        "No. cuenta",
    )
    COL_REFPAGO_CANDIDATAS = (
        "Referencia de pago",
        "Referencia pago",
        "Ref. de pago",
        "Ref.pago",
    )

    def __init__(self, gallo, update_status=None, con_ref_pago=False):
        """
        Args:
            gallo (pd.DataFrame): the Gallo CSV as read by Comun.leer_csv (all
                text). It is MUTATED in place by `preparar` (helper columns).
            update_status (callable|None): receives each progress message (str).
            con_ref_pago (bool): also resolve the 'Referencia de pago' column
                (only needed when 'TD.Diferencias_Referencias' will be built).
                It is NOT required otherwise, so a Gallo export without that
                column keeps consolidating exactly as before.
        """
        self.gallo = gallo
        self._update_status = update_status or (lambda _mensaje: None)
        self.con_ref_pago = con_ref_pago
        self.resultado = None

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    @staticmethod
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

    @staticmethod
    def _col_por_nombre(df, nombre):
        """
        Resolve a column strictly by name (no positional fallback).

        Used for columns whose position in the SAP export is not fixed/known
        ('Clase de documento', 'Importe en moneda local'). If the header is not
        present we fail loud with a clear, actionable message instead of
        guessing a column index and silently summing the wrong data.
        """
        if nombre in df.columns:
            return nombre
        raise ValueError(
            f"No encontré la columna '{nombre}' en el archivo de Gallo. "
            f"Columnas disponibles: {list(df.columns)}"
        )

    @staticmethod
    def clasificar_tipo(referencia2):
        """
        Step 4: 'Global' if it contains at least one digit, or if it is PP/RG/SG.
        'Individual' for the rest (e.g. letter-only prefixes).
        """
        ref = str(referencia2).strip()
        if any(caracter.isdigit() for caracter in ref):
            return "Global"
        if ref.upper() in EtapaGallo.PREFIJOS_LETRAS_GLOBAL:
            return "Global"
        return "Individual"

    # ------------------------------------------------------------------
    # Steps 2-9 — helper columns of the 'Gallo' sheet
    # ------------------------------------------------------------------
    def preparar(self):
        """
        Steps 2-9 + the SG/RG override + ONE amount-parsing pass, over the Gallo
        DataFrame given to the constructor.

        It MUTATES `self.gallo` in place: it appends the helper columns
        'Referencia 2', 'Referencia 3' and 'Comercializadora' (K, L and M with
        the standard A-J layout), which is exactly what the 'Gallo' sheet shows.

        The step order is part of the business logic: step 9 takes its snapshot
        of the globals BEFORE the SG/RG override (called "step 8.5" in the
        constants) reclassifies rows. Do not reorder.

        Returns:
            dict: what the next stages need (also kept in `self.resultado`) →
                col_centro, col_asig, col_doc, col_clase, col_fecha, col_cuenta,
                col_ref_pago (None unless `con_ref_pago`)
                    (the REAL column names found in this file),
                refs_unicas, tipos (steps 3-4), catalogo_docs (step 6),
                globales (step 9), filas_reclasificadas (SG/RG override),
                importe_num ('Importe en moneda local' as float, same index as
                    the Gallo DataFrame).

        Raises:
            ValueError: a required column is missing or an amount cannot be
                interpreted.
        """
        gallo = self.gallo
        update_status = self._update_status

        # Gallo key columns (by name, with a positional fallback)
        col_centro = self._col(gallo, "Centro",       1, "B")
        col_ref    = self._col(gallo, "Referencia",   2, "C")
        col_asig   = self._col(gallo, "Asignación",   3, "D")
        col_doc    = self._col(gallo, "Nº documento", 4, "E")

        # Columns the matrices need. They are resolved STRICTLY by name (their
        # position in the export is not fixed) and fail early with a clear error
        # if missing, before anything is processed.
        col_clase   = self._col_por_nombre(gallo, "Clase de documento")
        col_importe = self._col_por_nombre(gallo, "Importe en moneda local")
        col_fecha   = Comun.col_por_candidatos(
            gallo, self.COL_FECHA_CANDIDATAS, "fecha de contabilización")
        col_cuenta  = Comun.col_por_candidatos(
            gallo, self.COL_CUENTA_CANDIDATAS, "cuenta")
        col_ref_pago = (
            Comun.col_por_candidatos(
                gallo, self.COL_REFPAGO_CANDIDATAS, "referencia de pago")
            if self.con_ref_pago else None
        )

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
        tipos = [self.clasificar_tipo(ref) for ref in refs_unicas]
        mapa_tipo = dict(zip(refs_unicas, tipos))

        # === Steps 5 and 6: catalog of 'MX Commercial' documents ==========
        update_status("🔎 Filtrando 'MX Commercial'...")
        mascara_mx = gallo[col_asig].astype(str).str.strip() == self.VALOR_ASIGNACION
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
            (gallo["Comercializadora"] == self.VALOR_ASIGNACION)
            & (gallo["Referencia 3"] == "Global")
        )
        globales = gallo.loc[mascara_final, [col_centro, col_doc]].copy()

        # === SG/RG override → 'Individual' ("step 8.5", runs AFTER Step 9) ===
        # Business override: for MX Commercial rows whose Referencia 2 is SG or
        # RG, the amount is settled individually, so Referencia 3 becomes
        # 'Individual' instead of the default 'Global' that SG/RG get in
        # `clasificar_tipo`. Applied AFTER the Step 9 snapshot, so it does NOT
        # touch the G-H block; it DOES change 'Referencia 3' in the Gallo sheet
        # and therefore the 'TD' matrix.
        update_status("🏷️ Reclasificando SG/RG de 'MX Commercial' a 'Individual'...")
        ref2_norm = gallo["Referencia 2"].astype(str).str.strip().str.upper()
        mascara_reclasificar = (
            (gallo["Comercializadora"] == self.VALOR_ASIGNACION)
            & (ref2_norm.isin(self.PREFIJOS_RECLASIFICAR))
            & (gallo["Referencia 3"] == "Global")
        )
        filas_reclasificadas = int(mascara_reclasificar.sum())
        gallo.loc[mascara_reclasificar, "Referencia 3"] = "Individual"

        # === Amount: ONE single parsing pass for all the matrices ==========
        # A 3-day range is ~1.15M rows: parsing the same column twice would
        # waste time, and a single pass also guarantees that 'TD',
        # 'TD.Diferencias_Referencias' and the 'Venta DSD' sheet see exactly
        # the same numbers.
        update_status("🔢 Convirtiendo 'Importe en moneda local' a número...")
        importe_num = gallo[col_importe].map(
            partial(Comun.parse_importe_us, columna=col_importe))

        self.resultado = {
            "col_centro":           col_centro,
            "col_asig":             col_asig,
            "col_doc":              col_doc,
            "col_clase":            col_clase,
            "col_fecha":            col_fecha,
            "col_cuenta":           col_cuenta,
            "col_ref_pago":         col_ref_pago,
            "refs_unicas":          refs_unicas,
            "tipos":                tipos,
            "catalogo_docs":        catalogo_docs,
            "globales":             globales,
            "filas_reclasificadas": filas_reclasificadas,
            "importe_num":          importe_num,
        }
        return self.resultado

    # ------------------------------------------------------------------
    # 'Validación - Doc.' sheet
    # ------------------------------------------------------------------
    def escribir_validacion(self, ws):
        """
        Write the 'Validación - Doc.' sheet. It is built cell by cell because it
        holds three independent blocks side by side:

            A-B → unique 'Referencia 2' values and their type   (steps 3 and 4)
            D-E → documents whose Asignación is 'MX Commercial' (step 6)
            G-H → Centro and Nº documento of the globals        (step 9)

        Args:
            ws: the new (empty) openpyxl worksheet to fill.
        """
        pasos         = self.resultado
        refs_unicas   = pasos["refs_unicas"]
        tipos         = pasos["tipos"]
        catalogo_docs = pasos["catalogo_docs"]
        globales      = pasos["globales"]
        col_asig      = pasos["col_asig"]
        col_doc       = pasos["col_doc"]
        col_centro    = pasos["col_centro"]

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


# ======================================================================
# Stage: 'TD' — pivot-style matrices (Facturas, Notas de Crédito, PEP)
# ======================================================================
class EtapaTD:
    """
    The macro 'TD' sheet ("Tabla Dinámica"): sum of 'Importe en moneda local'
    by cedis × document class, in three blocks side by side:

        Facturas         → A: cedis, B: C1, C: C5, D: Grand Total
        Notas de Crédito → L: cedis, M: C2, N: C6, O: Grand Total
        PEP vouchers     → W: cedis, X: C2 (Referencia 2 == 'PP')

    The cross of these blocks against the monitor ('MTD') is written afterwards
    by `EtapaCruceTDMTD`, and the micro version of the Facturas block (down to
    document and payment reference) by `EtapaTDReferencias`. Both read the
    layout constants defined here, so a layout change happens in ONE place.

    The stage also keeps the normalized Gallo series it builds (`centro_norm`,
    `clase_norm`, ...) so that `mascara_bloque` can hand the EXACT SAME
    population of rows to the micro stage: if the filters of a block change,
    the macro and the micro tables change together and can never drift.
    """

    HOJA_TD = "TD"

    # --- 'TD' matrix layout ---
    # Rows 1-3 = filter legend, row 4 = block title, row 5 = headers, data
    # starts at row 6. Facturas live on the left block (row labels in column A),
    # Notas de Crédito in the middle one (column L) and PEP vouchers on the
    # right (column W). The legend goes in rows 1-3 of each block's FIRST
    # column, which the layout keeps free (the 'Monitor'/'Contabilidad'/
    # 'Diferencias' labels sit in the Grand Total column, see LABELS_ANALISIS).
    FILA_FILTROS    = 1
    FILA_ENCABEZADO = 5
    FILA_INICIO     = 6
    # 'clases' = (document class, 1-based column index of its header/sum).
    # Each block has its own 'Grand Total' column (D in Facturas, O in Notas).
    FACTURAS = {
        "titulo": "FACTURAS",
        "col_cedis": 1,                      # A
        "clases": [("C1", 2), ("C5", 3)],    # B / C
        "col_grand_total": 4,                # D
    }
    NOTAS = {
        "titulo": "NOTAS DE CRÉDITO",
        "col_cedis": 12,                     # L
        "clases": [("C2", 13), ("C6", 14)],  # M / N
        "col_grand_total": 15,               # O
    }
    # Analysis labels for the stakeholder: they go in the Grand Total column,
    # rows 1-3 (D1:D3 in Facturas, O1:O3 in Notas). Labels only.
    LABELS_ANALISIS = [(1, "Monitor"), (2, "Contabilidad"), (3, "Diferencias")]

    # Third matrix: PEP vouchers. Filter: Referencia 3=='Global',
    # Referencia 2=='PP' and Clase de documento=='C2' (plus the non-blank Centro
    # that every TD block needs). A single value column (X), with the cedis
    # rows in W. It does NOT exclude 'MX Commercial'.
    PEP = {
        "titulo": "PEP VOUCHERS",
        "col_cedis": 23,           # W
        "col_valor": 24,           # X
        "encabezado_valor": "C2",  # class of the filter (goes in X5)
    }
    PEP_REFERENCIA2 = "PP"
    PEP_CLASE = "C2"

    def __init__(self, gallo, pasos_gallo):
        """
        Args:
            gallo (pd.DataFrame): the Gallo DataFrame AFTER `EtapaGallo.preparar`
                (it needs the helper columns 'Referencia 2' and 'Referencia 3').
            pasos_gallo (dict): what `EtapaGallo.preparar` returned.
        """
        self.gallo = gallo
        self.pasos = pasos_gallo
        self.resultado = None
        # Normalized series, filled by `preparar` (see `mascara_bloque`).
        self.centro_norm = None
        self.clase_norm = None
        self._base = None

    # ------------------------------------------------------------------
    # Computation
    # ------------------------------------------------------------------
    def preparar(self):
        """
        Build the aggregated data for the 'TD' sheet.

        IMPORTANT: every matrix derives its OWN cedis list from its OWN filters
        (document class included), so the three blocks can list different
        centros and have different row counts. There is no 80AB..80VW range
        clipping: any centro that satisfies a matrix's filters becomes a row of
        that matrix.

            Facturas     → Referencia 3=='Global' & Asignación != 'MX Commercial'
                           & Clase de documento in {C1, C5}
            Notas        → same base, Clase de documento in {C2, C6}
            PEP vouchers → Referencia 3=='Global' & Referencia 2=='PP'
                           & Clase de documento=='C2'  (MX Commercial NOT excluded)

        All three blocks also require a non-blank Centro: a row without one has
        no place on the matrix, so it stays out of 'TD' (unlike 'MTD', which
        shows them as 'Sin centro'). Those rows are NOT dropped silently: they
        are counted, their amount is added up and reported in 'avisos', so a
        'Grand Total' that falls short can always be explained.

        The amount comes ALREADY parsed to float in `pasos_gallo["importe_num"]`
        (single pass shared with the 'Venta DSD' sheet).

        Returns a dict with one entry per block ('facturas', 'notas', 'pep'),
        each holding its own 'cedis' list and 'suma' lookup, plus the control
        totals, the grand total of the whole 'Importe en moneda local' column
        and 'avisos' (list[str]) with the rows left out for having no Centro.
        """
        gallo = self.gallo
        pasos = self.pasos
        col_centro  = pasos["col_centro"]
        col_asig    = pasos["col_asig"]
        col_clase   = pasos["col_clase"]
        importe_num = pasos["importe_num"]
        valor_asignacion = EtapaGallo.VALOR_ASIGNACION

        # Normalize once.
        centro_norm = gallo[col_centro].astype(str).str.strip()
        clase_norm  = gallo[col_clase].astype(str).str.strip().str.upper()
        ref2_norm   = gallo["Referencia 2"].astype(str).str.strip().str.upper()
        asig_norm   = gallo[col_asig].astype(str).str.strip()
        es_global   = gallo["Referencia 3"] == "Global"
        con_centro  = centro_norm != ""

        # Base subset shared by Facturas and Notas (each adds its own class
        # filter).
        base = es_global & (asig_norm != valor_asignacion) & con_centro
        # The same filters for the rows WITHOUT a centro: they cannot be placed
        # on the matrix, so they end up in 'avisos' instead of vanishing.
        base_sin_centro = es_global & (asig_norm != valor_asignacion) & ~con_centro

        # Kept for `mascara_bloque` (the micro stage reuses the same population).
        self.centro_norm = centro_norm
        self.clase_norm = clase_norm
        self._base = base

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

        facturas = _bloque(base, self.FACTURAS["clases"])
        notas    = _bloque(base, self.NOTAS["clases"])

        # PEP vouchers: its own filter, and it does NOT exclude 'MX Commercial'
        # (the stakeholder asked for those three filters only).
        filtro_pep = (
            es_global
            & (ref2_norm == self.PEP_REFERENCIA2)
            & (clase_norm == self.PEP_CLASE)
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
        # reported per block (count + amount) so the stakeholder knows exactly
        # how much is missing from that block's 'Grand Total'.
        clases_facturas = {clase for clase, _ in self.FACTURAS["clases"]}
        clases_notas    = {clase for clase, _ in self.NOTAS["clases"]}
        avisos = []
        for titulo, mascara in (
            (self.FACTURAS["titulo"], base_sin_centro & clase_norm.isin(clases_facturas)),
            (self.NOTAS["titulo"],    base_sin_centro & clase_norm.isin(clases_notas)),
            (self.PEP["titulo"],      filtro_pep & ~con_centro),
        ):
            n_filas = int(mascara.sum())
            if n_filas:
                importe = round(float(importe_num[mascara].sum()), 2)
                avisos.append(
                    f"TD ({titulo}): {n_filas:,} fila(s) sin Centro (importe "
                    f"{importe:,.2f}) no caben en la matriz; el "
                    f"'{Comun.ETIQUETA_GRAND_TOTAL}' de ese bloque no las incluye."
                )

        self.resultado = {
            "facturas":            facturas,
            "notas":               notas,
            "pep":                 {"cedis": cedis_pep, "suma": suma_pep},
            "total_facturas":      round(float(sum(facturas["suma"].values())), 2),
            "total_notas":         round(float(sum(notas["suma"].values())), 2),
            "total_pep":           round(float(sum(suma_pep.values())), 2),
            "total_importe_gallo": round(float(importe_num.sum()), 2),
            "avisos":              avisos,
            # Names the columns REALLY have in this export (they accept
            # spelling variants: 'Centro'/'Cebe', 'Clase de documento'/…). The
            # filter legend cites them so the stakeholder reads the same header
            # the sheet shows.
            "columnas": {
                "centro":      col_centro,
                "asignacion":  col_asig,
                "clase":       col_clase,
            },
        }
        return self.resultado

    def mascara_bloque(self, clases):
        """
        Boolean mask (over the Gallo rows) of the population of a two-class
        block — the SAME rows `preparar` summed for it: Referencia 3 == 'Global',
        Asignación != 'MX Commercial', non-blank Centro and Clase de documento
        in `clases` (a list of (class, column) pairs like FACTURAS["clases"]).

        Must be called after `preparar`. `EtapaTDReferencias` uses it so the
        micro table is built from exactly the rows of the macro table.
        """
        if self._base is None:
            raise RuntimeError("EtapaTD.mascara_bloque requiere llamar antes a preparar().")
        return self._base & self.clase_norm.isin({clase for clase, _ in clases})

    # ------------------------------------------------------------------
    # Filter legends
    # ------------------------------------------------------------------
    @staticmethod
    def leyenda_filtros(bloque, columnas):
        """
        Filter legend of a two-class 'TD' block (Facturas / Notas de Crédito),
        built FROM that block's own constants: change FACTURAS["clases"] or
        EtapaGallo.VALOR_ASIGNACION and the legend follows, so what the sheet
        claims and what the code sums cannot drift apart.

        `columnas` carries the resolved column names of this export (see
        `preparar`). 'Referencia 3' is NOT resolved: that column does not come
        from SAP, `EtapaGallo` creates it with exactly that name.
        """
        clases = " o ".join(f"'{clase}'" for clase, _ in bloque["clases"])
        return [
            f"Filtros — hoja '{EtapaGallo.HOJA_GALLO}':",
            f"Referencia 3 = 'Global'; {columnas['centro']} no vacío",
            f"{columnas['asignacion']} ≠ '{EtapaGallo.VALOR_ASIGNACION}'; "
            f"{columnas['clase']} = {clases}",
        ]

    @staticmethod
    def leyenda_filtros_pep(columnas):
        """
        Filter legend of the PEP vouchers block. Same idea as `leyenda_filtros`,
        with the two differences the block really has: it filters by
        'Referencia 2' (created by `EtapaGallo`, hence not resolved) instead of
        the Asignación, and it does NOT exclude VALOR_ASIGNACION — which is
        spelled out so nobody assumes the exclusion is there.
        """
        return [
            f"Filtros — hoja '{EtapaGallo.HOJA_GALLO}':",
            f"Referencia 3 = 'Global'; {columnas['centro']} no vacío",
            f"Referencia 2 = '{EtapaTD.PEP_REFERENCIA2}'; "
            f"{columnas['clase']} = '{EtapaTD.PEP_CLASE}' "
            f"(sí incluye '{EtapaGallo.VALOR_ASIGNACION}')",
        ]

    # ------------------------------------------------------------------
    # Writing
    # ------------------------------------------------------------------
    def escribir(self, ws_td):
        """
        Write the three matrices onto the 'TD' worksheet.

        Layout (rows 1-3 = filter legend + stakeholder labels, row 4 = block
        title, row 5 = headers, row 6+ = data, and a 'Grand Total' row right
        below the last cedis):

            Facturas         → A: cedis, B: sum(C1), C: sum(C5), D: Grand Total (B+C)
            Notas de Crédito → L: cedis, M: sum(C2), N: sum(C6), O: Grand Total (M+N)
            PEP vouchers     → W: cedis, X: sum(C2 with Referencia 2 == 'PP')

        Extras:
            - Filter legend above each matrix (A1:A3, L1:L3, W1:W3): the exact
              conditions its rows meet, so any figure can be reproduced by hand
              from the 'Gallo' sheet.
            - 'Grand Total' column per block (D / O): per-row sum of its two
              classes.
            - 'Grand Total' row at the bottom: per-column sum; the corner cell
              is the block's overall total.
            - Analysis labels 'Monitor' / 'Contabilidad' / 'Diferencias' in
              D1:D3 and O1:O3 for the stakeholder's later work (labels only, no
              values).

        Each block owns its cedis list (derived from its OWN filters), so the
        blocks can differ in which centros they show and in how many rows they
        have — every block therefore computes its own 'Grand Total' row
        position. Body cells are real numbers (values, not formulas) with a
        thousands/2-decimals format.

        The cross against 'MTD' (live formulas in E:G, P:R and Y:AA) is written
        afterwards by `EtapaCruceTDMTD.escribir`. The bottom-row label MUST stay
        Comun.ETIQUETA_GRAND_TOTAL: that cross looks it up in 'MTD' to compare
        totals.

        Args:
            ws_td: the new (empty) openpyxl worksheet to fill.
        """
        td_resumen = self.resultado
        fila_inicio = self.FILA_INICIO
        fila_encabezado = self.FILA_ENCABEZADO

        # Column names of THIS export, for the filter legends.
        columnas = td_resumen["columnas"]

        for bloque, datos in ((self.FACTURAS, td_resumen["facturas"]),
                              (self.NOTAS,    td_resumen["notas"])):
            cedis      = datos["cedis"]
            suma       = datos["suma"]
            n_cedis    = len(cedis)
            fila_total = fila_inicio + n_cedis   # row right below the last cedis
            col_cedis  = bloque["col_cedis"]
            col_gt     = bloque["col_grand_total"]
            clases     = bloque["clases"]

            # Filter legend (rows 1-3, in the cedis column) and analysis labels
            # for the stakeholder (same rows, in the Grand Total column).
            Comun.escribir_filtros(ws_td, self.FILA_FILTROS, col_cedis,
                                   self.leyenda_filtros(bloque, columnas))
            for fila, texto in self.LABELS_ANALISIS:
                ws_td.cell(row=fila, column=col_gt, value=texto).font = Font(bold=True)

            # Block title (row 4).
            ws_td.cell(row=4, column=col_cedis, value=bloque["titulo"]).font = Font(bold=True)

            # Header row (row 5): 'Centro' + class headers + 'Grand Total'.
            ws_td.cell(row=fila_encabezado, column=col_cedis, value="Centro").font = Font(bold=True)
            for clase, col_idx in clases:
                ws_td.cell(row=fila_encabezado, column=col_idx, value=clase).font = Font(bold=True)
            ws_td.cell(row=fila_encabezado, column=col_gt, value="Grand Total").font = Font(bold=True)

            # Body rows + per-column accumulators for the bottom 'Grand Total'
            # row.
            totales_clase = {clase: 0.0 for clase, _ in clases}
            for i, centro in enumerate(cedis):
                fila = fila_inicio + i
                ws_td.cell(row=fila, column=col_cedis, value=centro)
                suma_fila = 0.0
                for clase, col_idx in clases:
                    valor = round(suma.get((centro, clase), 0.0), 2)
                    ws_td.cell(row=fila, column=col_idx, value=valor).number_format = Comun.FORMATO_NUMERO
                    totales_clase[clase] += valor
                    suma_fila += valor
                # 'Grand Total' column = sum of this row's classes.
                ws_td.cell(row=fila, column=col_gt, value=round(suma_fila, 2)).number_format = Comun.FORMATO_NUMERO

            # Bottom 'Grand Total' row = per-column totals (only when there are
            # cedis).
            if n_cedis > 0:
                ws_td.cell(row=fila_total, column=col_cedis,
                           value=Comun.ETIQUETA_GRAND_TOTAL).font = Font(bold=True)
                suma_total = 0.0
                for clase, col_idx in clases:
                    tot = round(totales_clase[clase], 2)
                    celda = ws_td.cell(row=fila_total, column=col_idx, value=tot)
                    celda.number_format = Comun.FORMATO_NUMERO
                    celda.font = Font(bold=True)
                    suma_total += tot
                celda_gt = ws_td.cell(row=fila_total, column=col_gt, value=round(suma_total, 2))
                celda_gt.number_format = Comun.FORMATO_NUMERO
                celda_gt.font = Font(bold=True)

        # --- Third matrix: PEP vouchers (its own cedis, single value column) ---
        cedis_pep      = td_resumen["pep"]["cedis"]
        suma_pep       = td_resumen["pep"]["suma"]
        col_cedis_pep  = self.PEP["col_cedis"]
        col_valor_pep  = self.PEP["col_valor"]
        fila_total_pep = fila_inicio + len(cedis_pep)

        Comun.escribir_filtros(ws_td, self.FILA_FILTROS, col_cedis_pep,
                               self.leyenda_filtros_pep(columnas))
        ws_td.cell(row=4, column=col_cedis_pep, value=self.PEP["titulo"]).font = Font(bold=True)
        ws_td.cell(row=fila_encabezado, column=col_cedis_pep, value="Centro").font = Font(bold=True)
        ws_td.cell(row=fila_encabezado, column=col_valor_pep,
                   value=self.PEP["encabezado_valor"]).font = Font(bold=True)

        total_pep = 0.0
        for i, centro in enumerate(cedis_pep):
            fila = fila_inicio + i
            ws_td.cell(row=fila, column=col_cedis_pep, value=centro)
            valor = round(suma_pep.get(centro, 0.0), 2)
            ws_td.cell(row=fila, column=col_valor_pep, value=valor).number_format = Comun.FORMATO_NUMERO
            total_pep += valor
        if cedis_pep:
            ws_td.cell(row=fila_total_pep, column=col_cedis_pep,
                       value=Comun.ETIQUETA_GRAND_TOTAL).font = Font(bold=True)
            celda_pep = ws_td.cell(row=fila_total_pep, column=col_valor_pep, value=round(total_pep, 2))
            celda_pep.number_format = Comun.FORMATO_NUMERO
            celda_pep.font = Font(bold=True)

        # Comfortable widths (incl. the Grand Total columns D/O and the PEP W/X).
        anchos = {"A": 12, "B": 16, "C": 16, "D": 16,
                  "L": 12, "M": 16, "N": 16, "O": 16,
                  "W": 12, "X": 16}
        for letra, ancho in anchos.items():
            ws_td.column_dimensions[letra].width = ancho


# ======================================================================
# Stage: 'Venta DSD' — amount by date
# ======================================================================
class EtapaVentas:
    """
    The 'Venta DSD' date × amount sheet, plus the date handling it needs:

        'Venta DSD'  → classes C1/C2/C5/C6/X0 and blank class, accounts 4000010/4000005

    (There used to be a second sheet, 'Venta EIAP'. It was removed from the
    consolidation at the stakeholder's request: that data is now downloaded on
    its own with the 'Solo Venta EIAP' button, see Descargas_SAP.VentaEIAP_FAGLL03.)

    Two-column matrix: date (A) and summed amount (B). It is filtered by TWO
    criteria at once (AND): document class AND account. The empty string ""
    among the classes stands for Excel's "(blanks)", i.e. a class that is empty
    or only spaces.

    Also validates the dates the user selected in the GUI
    (`parsear_fechas_seleccionadas`), because only this sheet uses them.
    """

    HOJA_DSD  = "Venta DSD"
    # (key in the result dict, sheet name), in the order they are written.
    HOJAS = (("dsd", HOJA_DSD),)

    CLASES_DSD  = frozenset({"C1", "C2", "C5", "C6", "X0", ""})
    # G/L accounts. They are compared NORMALIZED (no leading zeros and no '.0'
    # left behind by the xlsx→csv conversion), so it does not matter whether
    # SAP exports them as '4000010', '0004000010' or '4000010.0'.
    CUENTAS_DSD  = frozenset({"4000010", "4000005"})

    ENCABEZADO_FECHA   = "Fecha"
    # (the amount header is Comun.ENCABEZADO_IMPORTE.)
    ETIQUETA_SIN_FECHA = "Sin fecha"
    ETIQUETA_TOTAL     = Comun.ETIQUETA_GRAND_TOTAL
    FORMATO_FECHA_EXCEL = "DD.MM.YYYY"

    # Layout: rows 1-4 = filter legend, row 5 = headers, row 6 onwards = data.
    # The legend needs four lines here because this matrix filters by TWO
    # criteria at once (document class AND account). Data therefore starts on
    # row 6, the same as in 'TD' and 'MTD'. No formula points at these sheets,
    # so moving the rows down affects nothing else in the workbook.
    FILA_FILTROS    = 1
    FILA_ENCABEZADO = 5
    FILA_INICIO     = 6
    # How the empty document class (Excel's "(blanks)") reads in the legend: it
    # IS one of the classes that DSD sums, and '' on screen would look like a
    # typo.
    ETIQUETA_EN_BLANCO = "en blanco"

    # Date formats the SAP export may use. The 4-digit-year formats are tried
    # first; the 2-digit ones are the last resort. Real ambiguity (e.g.
    # 01/09/2026 = Sep 1 or Jan 9) is resolved against the dates the user
    # selected in the GUI; if it cannot be resolved, it fails with a clear
    # message instead of guessing.
    FORMATOS_FECHA_4D = (
        "%d.%m.%Y", "%m.%d.%Y",
        "%d/%m/%Y", "%m/%d/%Y",
        "%Y-%m-%d", "%Y/%m/%d",
        "%d-%m-%Y", "%m-%d-%Y",
    )
    FORMATOS_FECHA_2D = ("%d.%m.%y", "%m.%d.%y", "%d/%m/%y", "%m/%d/%y")
    ANIO_MIN = 1990
    ANIO_MAX = 2100

    # The Gallo export brings the date as 'dd/mm/yyyy hh:mm'. This pattern keeps
    # only the DATE part (group 1) and drops whatever comes after a space or a
    # 'T'. It does not try to interpret the day/month order: `detectar_formato`
    # takes care of that.
    _RE_SOLO_FECHA = re.compile(r"^(\d{1,4}[./\-]\d{1,2}[./\-]\d{1,4})(?:[ T].*)?$")

    def __init__(self, gallo, pasos_gallo, fechas_sel=()):
        """
        Args:
            gallo (pd.DataFrame): the Gallo DataFrame after `EtapaGallo.preparar`.
            pasos_gallo (dict): what `EtapaGallo.preparar` returned.
            fechas_sel (list[date]): what `parsear_fechas_seleccionadas` returned.
        """
        self.gallo = gallo
        self.pasos = pasos_gallo
        self.fechas_sel = fechas_sel
        self.resultado = None

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    @staticmethod
    def normalizar_cuenta(valor):
        """
        Normalize a G/L account so it compares without format surprises:

            ' 0004000010 ' -> '4000010'
            '4000010.0'    -> '4000010'   (the xlsx→csv conversion made it a float)
            '4000010'      -> '4000010'

        Only LEADING zeros and an all-zero decimal part are removed; never
        significant digits. A non-numeric value is returned as is (upper-cased
        and without spaces), so an account with letters stays comparable.
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

    @staticmethod
    def solo_fecha(valor):
        """
        Cut the time off a date value: the Gallo export brings 'dd/mm/yyyy
        hh:mm' and the time gets in the way when grouping by day.

            '01/09/2026 14:35' -> '01/09/2026'
            '2026-09-01T00:00' -> '2026-09-01'
            '01.09.2026'       -> '01.09.2026'

        If the value does not start with something that looks like a date, it is
        returned as is (just stripped), so the format detector fails showing the
        REAL value instead of one mangled by this helper.
        """
        texto = str(valor).strip()
        coincidencia = EtapaVentas._RE_SOLO_FECHA.match(texto)
        return coincidencia.group(1) if coincidencia else texto

    @staticmethod
    def parsear_fechas_seleccionadas(fechas):
        """
        Validate and normalize the dates the user picked in the GUI
        ('DD.MM.YYYY') into date objects, without duplicates and in
        chronological order.

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

    @staticmethod
    def _intentar_formato(valores, formato):
        """
        Return the list of dates if ALL values parse with that format and fall
        in a reasonable year; None if any of them fails.

        The year check is not cosmetic: strptime accepts '26' for %Y and would
        return the year 26 AD, which would let a wrong format pass as valid.
        """
        convertidas = []
        for valor in valores:
            try:
                fecha = datetime.strptime(valor, formato).date()
            except ValueError:
                return None
            if not (EtapaVentas.ANIO_MIN <= fecha.year <= EtapaVentas.ANIO_MAX):
                return None
            convertidas.append(fecha)
        return convertidas

    @staticmethod
    def detectar_formato(valores, fechas_referencia=()):
        """
        Detect the format of the date column of the Gallo CSV.

        Rule: the 4-digit-year formats are tried first and, only if none of them
        works, the 2-digit ones. If several formats parse EVERYTHING:
          - if they all give the SAME result, the ambiguity is harmless → the
            1st one is used;
          - if they give different results, the tie is broken with the dates
            the user selected (the format that puts more dates inside the
            selection wins);
          - if it is still a tie, ValueError is raised. It never guesses: a
            misread 01/09/2026 would silently move the amount to another day.

        Returns the format (str), or None if there is no value to parse.
        """
        valores = [v for v in valores if v]
        if not valores:
            return None

        referencia = set(fechas_referencia or ())

        for tanda in (EtapaVentas.FORMATOS_FECHA_4D, EtapaVentas.FORMATOS_FECHA_2D):
            viables = []
            for formato in tanda:
                convertidas = EtapaVentas._intentar_formato(valores, formato)
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
            f"{', '.join(EtapaVentas.FORMATOS_FECHA_4D + EtapaVentas.FORMATOS_FECHA_2D)}."
        )

    # ------------------------------------------------------------------
    # Computation
    # ------------------------------------------------------------------
    def preparar(self):
        """
        Build the 'Venta DSD' date × amount matrix.

        Returns:
            dict: {'dsd': block, 'avisos': list[str]}, where the block holds
            filas [(date, amount)], sin_fecha_n, sin_fecha_importe, total and
            fuera_rango.
        """
        gallo = self.gallo
        pasos = self.pasos
        col_clase   = pasos["col_clase"]
        col_cuenta  = pasos["col_cuenta"]
        col_fecha   = pasos["col_fecha"]
        importe_num = pasos["importe_num"]
        fechas_sel  = self.fechas_sel

        clase_norm  = gallo[col_clase].astype(str).str.strip().str.upper()
        cuenta_norm = gallo[col_cuenta].map(self.normalizar_cuenta)
        # Trim the time BEFORE detecting the format: '01/09/2026 14:35' is a day.
        fecha_raw   = gallo[col_fecha].map(self.solo_fecha)

        # Detect the date format ONCE, over the distinct values of the whole
        # column.
        distintas = sorted({valor for valor in fecha_raw.unique() if valor})
        formato = self.detectar_formato(distintas, fechas_sel)
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
            ("dsd", (self.HOJA_DSD, self.CLASES_DSD, self.CUENTAS_DSD)),
        ):
            cuentas_buscadas = {self.normalizar_cuenta(c) for c in cuentas}
            mascara = clase_norm.isin(clases) & cuenta_norm.isin(cuentas_buscadas)

            # A requested account that does NOT exist in the file almost always
            # means that day had no postings... or that the SAP layout changed.
            # Better to warn than to deliver an empty matrix with no
            # explanation.
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
                    f"(se muestran como '{self.ETIQUETA_SIN_FECHA}')."
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
                # Ingredients of the filter legend written above the matrix:
                # this block's OWN filters and the column names this export
                # really uses (they accept spelling variants). The text itself
                # is built by `escribir`, so the legend cannot describe filters
                # other than the ones applied right here.
                "clases":            clases,
                "cuentas":           cuentas,
                "columnas": {
                    "clase":  col_clase,
                    "cuenta": col_cuenta,
                    "fecha":  col_fecha,
                },
            }

        bloques["avisos"] = avisos
        self.resultado = bloques
        return bloques

    # ------------------------------------------------------------------
    # Writing
    # ------------------------------------------------------------------
    @staticmethod
    def leyenda_filtros(clases, cuentas, columnas):
        """
        Filter legend of the 'Venta DSD' matrix, built FROM that block's own
        filter sets (CLASES_* and CUENTAS_*) and from the column names this
        export really uses, so the text cannot drift from what the code sums.

        Two things are spelled out because they are exactly what makes a manual
        check disagree with the sheet:
            - the EMPTY document class counts for DSD (Excel shows it as
              "(blanks)"), and
            - the account is compared without leading zeros or a trailing '.0',
              so '0004000010' and '4000010.0' are the same account as '4000010'
              (filtering by the literal text in Excel would leave those rows
              out).
        """
        lista_clases = " o ".join(
            [f"'{clase}'" for clase in sorted(clase for clase in clases if clase)]
            + ([EtapaVentas.ETIQUETA_EN_BLANCO] if "" in clases else [])
        )
        lista_cuentas = " o ".join(f"'{cuenta}'" for cuenta in sorted(cuentas))
        return [
            f"Filtros — hoja '{EtapaGallo.HOJA_GALLO}':",
            f"{columnas['clase']} = {lista_clases}",
            f"{columnas['cuenta']} = {lista_cuentas}",
            f"{columnas['fecha']} vacía → renglón '{EtapaVentas.ETIQUETA_SIN_FECHA}'",
        ]

    def escribir(self, ws, clave):
        """
        Write the 'Venta DSD' sheet: column A = date, column B = summed amount.

        Rows 1-4 = filter legend, row 5 = headers, row 6+ = one row per day,
        then (only if it exists) a 'Sin fecha' row, and finally a bold 'Grand
        Total' row. Dates are written as REAL dates formatted DD.MM.YYYY, so the
        stakeholder can sort and filter them as dates; amounts are values (not
        formulas) with 2 decimals.

        The legend states the TWO filters this matrix applies at the same time
        (document class AND account), which is the pair nobody remembers by
        heart.

        Args:
            ws: the new (empty) openpyxl worksheet to fill.
            clave (str): 'dsd'.
        """
        bloque = self.resultado[clave]
        Comun.escribir_filtros(ws, self.FILA_FILTROS, 1,
                               self.leyenda_filtros(bloque["clases"], bloque["cuentas"],
                                                    bloque["columnas"]))

        ws.cell(row=self.FILA_ENCABEZADO, column=1,
                value=self.ENCABEZADO_FECHA).font = Font(bold=True)
        ws.cell(row=self.FILA_ENCABEZADO, column=2,
                value=Comun.ENCABEZADO_IMPORTE).font = Font(bold=True)

        fila = self.FILA_INICIO
        for fecha, valor in bloque["filas"]:
            celda_fecha = ws.cell(row=fila, column=1, value=fecha)
            celda_fecha.number_format = self.FORMATO_FECHA_EXCEL
            ws.cell(row=fila, column=2, value=valor).number_format = Comun.FORMATO_NUMERO
            fila += 1

        # Rows whose date was blank: shown apart so the total still reconciles.
        if bloque["sin_fecha_n"]:
            ws.cell(row=fila, column=1, value=self.ETIQUETA_SIN_FECHA).font = Font(italic=True)
            celda = ws.cell(row=fila, column=2, value=bloque["sin_fecha_importe"])
            celda.number_format = Comun.FORMATO_NUMERO
            celda.font = Font(italic=True)
            fila += 1

        if fila > self.FILA_INICIO:   # there is at least one data row
            ws.cell(row=fila, column=1, value=self.ETIQUETA_TOTAL).font = Font(bold=True)
            celda_total = ws.cell(row=fila, column=2, value=bloque["total"])
            celda_total.number_format = Comun.FORMATO_NUMERO
            celda_total.font = Font(bold=True)

        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 24


# ======================================================================
# Stage: Monivoi (ZLMXCOM_TRN_MONINVOI) — 'MTD' (step 16)
# ======================================================================
class EtapaMonivoi:
    """
    Everything that is computed FROM the Monivoi (the invoicing monitor):

        1) 'MTD' sheet (step 16): four 'Centro' × 'Valor neto' matrices (SG, RG,
           PP and TODOS)                      → preparar / escribir_mtd
        2) Helper column 'Referencia 2' as column A of the 'Monivoi' sheet
                                              → insertar_referencia2
        3) The list of invoices ('Factura' column) that the VBRL download pastes
           into SAP                           → facturas_unicas / leer_facturas

    The cross of 'MTD' against 'TD' lives in `EtapaCruceTDMTD`.
    """

    # Sheet names (tabs) built from the Monivoi.
    HOJA_MONIVOI = "Monivoi"
    HOJA_MTD     = "MTD"

    # --- 'MTD' sheet (step 16) ---
    # Four 'Centro' × sum of 'Valor neto' matrices, all taken from the MONIVOI.
    # Layout: rows 1-3 = filter legend, row 4 = title, row 5 = headers,
    # row 6 onwards = data. The three top rows exist ONLY for the legend, so the
    # stakeholder reads under which conditions the matrix was built before
    # reading the numbers. The TD × MTD cross is NOT affected by these rows: it
    # looks the centros up with XLOOKUP over WHOLE columns (MTD!A:A, MTD!B:B),
    # never over a fixed range.
    FILA_FILTROS    = 1
    FILA_TITULO     = 4
    FILA_ENCABEZADO = 5
    FILA_INICIO     = 6

    ENCABEZADO_CENTRO   = "Centro"
    ENCABEZADO_VALOR    = "Valor neto"
    ETIQUETA_TOTAL      = Comun.ETIQUETA_GRAND_TOTAL   # same as 'TD' (see the cross)
    ETIQUETA_SIN_CENTRO = "Sin centro"

    PREFIJO_SG = "SG"
    PREFIJO_RG = "RG"
    PREFIJO_PP = "PP"
    # The three prefixes that go into the fourth matrix ('TODOS'). They also
    # define the universe of rows that get parsed: nothing outside these three
    # is touched.
    PREFIJOS_TODOS = frozenset({PREFIJO_SG, PREFIJO_RG, PREFIJO_PP})

    # (internal key, title, prefixes it filters, 'Centro' column, value column)
    BLOQUES = (
        {"clave": "sg",    "titulo": "SG",
         "prefijos": frozenset({PREFIJO_SG}),
         "col_centro": 1,  "col_valor": 2},     # A / B
        {"clave": "rg",    "titulo": "RG",
         "prefijos": frozenset({PREFIJO_RG}),
         "col_centro": 4,  "col_valor": 5},     # D / E
        {"clave": "pp",    "titulo": "PP",
         "prefijos": frozenset({PREFIJO_PP}),
         "col_centro": 11, "col_valor": 12},    # K / L
        {"clave": "todos", "titulo": "TODOS (SG + RG + PP)",
         "prefijos": PREFIJOS_TODOS,
         "col_centro": 17, "col_valor": 18},    # Q / R
    )

    # Name of the helper column added AS THE FIRST COLUMN (A) of the 'Monivoi'
    # sheet, so the matrices can be audited by filtering by hand.
    COL_REFERENCIA2 = "Referencia 2"

    # Monivoi columns that 'MTD' needs. Spelling variants are accepted (the
    # comparison ignores accents, dots, spaces and case).
    COL_REFERENCIA_CANDIDATAS = ("Referencia", "Referencia de factura")
    COL_UUID_CANDIDATAS       = ("UUID", "UUID CFDI", "Folio fiscal", "UUID Fiscal")
    COL_CENTRO_CANDIDATAS     = ("Centro", "Centro de distribución",
                                 "Puesto de expedición")
    COL_VALOR_CANDIDATAS      = ("Valor neto", "Valor neto factura", "Importe neto")
    # The invoice number that is pasted into the VBRL selection.
    COL_FACTURA_CANDIDATAS    = ("Factura", "Nº factura", "No. factura", "N° factura")

    # Texts that count as an EMPTY UUID (besides a blank or whitespace-only
    # cell). They are compared upper-cased and stripped. Note: the SHAPE of the
    # UUID is NOT validated here on purpose — if SAP exported it without dashes
    # it would still be a real UUID, and we do not want to drop good invoices
    # over a format change. Only the "there is nothing here" sentinels are
    # discarded.
    UUID_VACIOS = frozenset({
        "", "NULL", "NONE", "NAN", "NA", "N/A", "#N/A", "#N/D", "-", "--", "#",
    })

    # Tolerance of the 'TODOS == SG + RG + PP' check. One cent covers the
    # rounding accumulated by adding values already rounded to 2 decimals.
    TOLERANCIA_CUADRE = 0.01

    # Direct access to each 'MTD' block by its key ('sg', 'rg', 'pp', 'todos').
    POR_CLAVE = {bloque["clave"]: bloque for bloque in BLOQUES}

    def __init__(self, monivoi, sin_datos):
        """
        Args:
            monivoi (pd.DataFrame): the Monivoi CSV as read by Comun.leer_csv.
                `insertar_referencia2` MUTATES it (new first column).
            sin_datos (bool): the CSV is the 'Sin datos' placeholder.
        """
        self.monivoi = monivoi
        self.sin_datos = sin_datos
        self.resultado = None

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    @staticmethod
    def nombre_columna_libre(df, base):
        """
        Return `base` if that name does NOT exist yet in the DataFrame;
        otherwise return 'base (2)', 'base (3)', ...

        Used to add the helper column 'Referencia 2' to the Monivoi sheet
        without overwriting a same-named column the export may bring
        (overwriting it would silently wipe the stakeholder's original data).
        """
        if base not in df.columns:
            return base
        sufijo = 2
        while f"{base} ({sufijo})" in df.columns:
            sufijo += 1
        return f"{base} ({sufijo})"

    @staticmethod
    def uuid_vacio(valor):
        """
        True if the row's UUID counts as EMPTY or NULL (step 16).

        Covers the blank cell, the whitespace-only cell and the text sentinels
        the export sometimes leaves ('NULL', 'N/A', '-', ...). It does NOT
        validate the UUID's shape: a real UUID written differently (no dashes,
        lower case) still counts as present.
        """
        return str(valor).strip().upper() in EtapaMonivoi.UUID_VACIOS

    @staticmethod
    def _bloques_vacios():
        """Empty structure (one per block) for when the Monivoi has no data."""
        return {
            bloque["clave"]: {
                "filas":              [],
                "sin_centro_n":       0,
                "sin_centro_importe": 0.0,
                "total":              0.0,
            }
            for bloque in EtapaMonivoi.BLOQUES
        }

    # ------------------------------------------------------------------
    # 'MTD' sheet (step 16) — computation
    # ------------------------------------------------------------------
    def preparar(self):
        """
        Step 16 — build the FOUR matrices of the 'MTD' sheet, all of them over
        the MONIVOI sheet (not over the Gallo).

        Grouping key: the first 2 characters of the Monivoi 'Referencia'
        column. Filter shared by all four: the row's UUID can NOT be empty or
        null (see `uuid_vacio`); those rows were never invoiced and including
        them would inflate the total against the monitor.

            Matrix 1 (A/B) → Referencia 2 == 'SG'
            Matrix 2 (D/E) → Referencia 2 == 'RG'
            Matrix 3 (K/L) → Referencia 2 == 'PP'
            Matrix 4 (Q/R) → Referencia 2 ∈ {'SG', 'RG', 'PP'}  (TODOS)

        Each matrix derives its OWN list of centros from ITS own filters, same
        as in 'TD', so the four blocks can have different centros and a
        different number of rows.

        'Valor neto' is only parsed for the rows that go into some matrix (the
        union of the three prefixes), not for the whole sheet.

        Rows that pass the filter but come WITHOUT a centro are not dropped:
        they are shown apart in a 'Sin centro' row, so the 'Grand Total' of
        each block always reconciles with its filter.

        Returns:
            dict: bloques ({'sg'|'rg'|'pp'|'todos': {filas, sin_centro_n,
            sin_centro_importe, total}}), referencia2 (Series for the helper
            column, None when there is no data), filas_sin_uuid and avisos.
        """
        monivoi = self.monivoi
        if self.sin_datos or monivoi.empty:
            self.resultado = {
                "bloques":        self._bloques_vacios(),
                "referencia2":    None,
                "filas_sin_uuid": 0,
                "avisos": [
                    "MTD: el archivo de Monivoi vino sin datos; las cuatro "
                    "matrices quedaron vacías."
                ],
                # No columns were resolved (we never got to read the file), so
                # the legend falls back to the canonical names. With no data
                # there is nothing to audit anyway.
                "columnas": {
                    "uuid":   self.COL_UUID_CANDIDATAS[0],
                    "centro": self.COL_CENTRO_CANDIDATAS[0],
                },
            }
            return self.resultado

        # Monivoi columns. They are resolved by name (accepting spelling
        # variants) and fail with a clear message if missing: stopping is
        # better than summing the wrong column.
        col_ref    = Comun.col_por_candidatos(monivoi, self.COL_REFERENCIA_CANDIDATAS,
                                              "referencia", origen="Monivoi")
        col_uuid   = Comun.col_por_candidatos(monivoi, self.COL_UUID_CANDIDATAS,
                                              "UUID", origen="Monivoi")
        col_centro = Comun.col_por_candidatos(monivoi, self.COL_CENTRO_CANDIDATAS,
                                              "centro de distribución", origen="Monivoi")
        col_valor  = Comun.col_por_candidatos(monivoi, self.COL_VALOR_CANDIDATAS,
                                              "valor neto", origen="Monivoi")

        avisos = []

        # Grouping key: first 2 characters of 'Referencia'. It is kept as is
        # (for column A of the sheet) and compared upper-cased.
        referencia2 = monivoi[col_ref].astype(str).str.strip().str[:2]
        ref2_norm   = referencia2.str.upper()
        centro_norm = monivoi[col_centro].astype(str).str.strip()

        con_uuid = ~monivoi[col_uuid].map(self.uuid_vacio)
        filas_sin_uuid = int((~con_uuid).sum())
        if filas_sin_uuid:
            avisos.append(
                f"MTD: {filas_sin_uuid:,} fila(s) de Monivoi se excluyeron por "
                f"traer el UUID vacío o nulo."
            )

        # Universe of rows that go into SOME matrix. The amount is parsed only
        # here (the union of the three prefixes == the 'TODOS' matrix).
        mascara_union = con_uuid & ref2_norm.isin(self.PREFIJOS_TODOS)
        # astype(float) is NOT cosmetic: if the mask lets no row through,
        # .map() returns an empty text-typed Series (the Monivoi is read as
        # text: 'object' in pandas 2, 'str' in pandas 3), and with pandas 3 its
        # .sum() gives '' (the sum of zero strings) instead of 0.0, breaking
        # the rounding. Pinning the dtype makes the empty case 0.0 on any
        # pandas version.
        valores = monivoi.loc[mascara_union, col_valor].map(
            partial(Comun.parse_importe_us, columna=col_valor)
        ).astype("float64")
        base = pd.DataFrame({
            "_ref2":   ref2_norm[mascara_union],
            "_centro": centro_norm[mascara_union],
            "_valor":  valores,
        })

        bloques = {}
        vacios  = []
        for bloque in self.BLOQUES:
            sub = base[base["_ref2"].isin(bloque["prefijos"])]

            sin_centro         = sub["_centro"] == ""
            sin_centro_n       = int(sin_centro.sum())
            sin_centro_importe = round(float(sub.loc[sin_centro, "_valor"].sum()), 2)

            con_centro = sub.loc[~sin_centro]
            if con_centro.empty:
                filas = []
            else:
                # groupby sorts the keys: centros come out alphabetically, as
                # in 'TD'.
                serie = con_centro.groupby("_centro")["_valor"].sum()
                filas = [(centro, round(float(valor), 2)) for centro, valor in serie.items()]

            # The total adds the ALREADY ROUNDED values, so the column adds up
            # to exactly what is shown on screen.
            total = round(sum(valor for _, valor in filas) + sin_centro_importe, 2)

            if not filas and not sin_centro_n:
                vacios.append(bloque["titulo"])
            if sin_centro_n:
                avisos.append(
                    f"MTD ({bloque['titulo']}): {sin_centro_n:,} fila(s) sin centro "
                    f"(se muestran como '{self.ETIQUETA_SIN_CENTRO}')."
                )

            bloques[bloque["clave"]] = {
                "filas":              filas,
                "sin_centro_n":       sin_centro_n,
                "sin_centro_importe": sin_centro_importe,
                "total":              total,
            }

        if vacios:
            avisos.append(
                f"MTD: sin filas para {', '.join(vacios)} "
                f"(no hubo Referencia con ese prefijo y UUID)."
            )

        # Reconciliation check: 'TODOS' is the union of the other three, so its
        # total MUST be the sum of the three. If it does not match, warn instead
        # of silently delivering the file (the tolerance covers rounding to
        # cents).
        suma_partes = round(
            bloques["sg"]["total"] + bloques["rg"]["total"] + bloques["pp"]["total"], 2
        )
        diferencia = round(bloques["todos"]["total"] - suma_partes, 2)
        if abs(diferencia) > self.TOLERANCIA_CUADRE:
            avisos.append(
                f"MTD: el total de 'TODOS' ({bloques['todos']['total']:,.2f}) no "
                f"cuadra con SG+RG+PP ({suma_partes:,.2f}); diferencia "
                f"{diferencia:,.2f}. Revisa el archivo antes de usar la matriz."
            )

        self.resultado = {
            "bloques":        bloques,
            "referencia2":    referencia2,
            "filas_sin_uuid": filas_sin_uuid,
            "avisos":         avisos,
            # Names the columns REALLY have in this export (they accept
            # spelling variants: 'Centro'/'Centro de distribución',
            # 'UUID'/'UUID CFDI'...). The filter legend cites them, so the
            # stakeholder reads the same header the 'Monivoi' sheet shows.
            "columnas": {
                "uuid":   col_uuid,
                "centro": col_centro,
            },
        }
        return self.resultado

    def insertar_referencia2(self):
        """
        Insert the helper column 'Referencia 2' AS THE FIRST COLUMN (A) of the
        Monivoi sheet, so the stakeholder can audit the 'MTD' matrices by
        filtering by hand.

        Call it AFTER `preparar` (the matrices use their own series, not the
        sheet). It uses a free column name, so it never overwrites a same-named
        column the export already had; in that case a warning is appended to
        resultado["avisos"].

        It MUTATES `self.monivoi` (new first column) and, in that case,
        resultado["avisos"]. It also records the FINAL name of the column in
        resultado["nombre_ref2"], which is what the 'MTD' filter legends cite.
        It does nothing when the Monivoi had no data (resultado["referencia2"]
        is None).
        """
        mtd = self.resultado
        if mtd["referencia2"] is not None:
            nombre_ref2 = self.nombre_columna_libre(self.monivoi, self.COL_REFERENCIA2)
            if nombre_ref2 != self.COL_REFERENCIA2:
                mtd["avisos"].append(
                    f"MTD: la Monivoi ya traía una columna '{self.COL_REFERENCIA2}'; "
                    f"la auxiliar se agregó como '{nombre_ref2}' para no pisarla."
                )
            self.monivoi.insert(0, nombre_ref2, mtd["referencia2"])
            # The filter legend of each matrix must point at the column the
            # stakeholder will REALLY find on the 'Monivoi' sheet. When the
            # export already brought a 'Referencia 2', the helper went in under
            # another name, and a legend still saying 'Referencia 2' would send
            # the audit to the wrong (pre-existing) column.
            mtd["nombre_ref2"] = nombre_ref2

    # ------------------------------------------------------------------
    # 'MTD' sheet — writing
    # ------------------------------------------------------------------
    @staticmethod
    def leyenda_filtros(bloque, nombre_ref2, columnas):
        """
        Filter legend of an 'MTD' block, built FROM THAT BLOCK'S OWN constants
        (its prefixes), so the text can never drift from what the matrix really
        sums: adding a prefix to BLOQUES changes the legend too.

        It states the three conditions `preparar` applies — source sheet,
        'Referencia 2' (first 2 characters of 'Referencia', upper-cased) and a
        non-empty UUID — plus what happens to the rows that pass the filter
        without a centro, which are the ones a manual check usually misses.

        `nombre_ref2` is the REAL name the helper column ended up with on the
        'Monivoi' sheet (see `insertar_referencia2`) and `columnas` carries the
        resolved names of the UUID and Centro columns of this export, so every
        column the legend names is one the stakeholder can actually filter.

        Each line is kept short on purpose: the first block sits in column A
        and only has columns B and C to overflow into before running into the
        next block's legend.
        """
        prefijos = " o ".join(f"'{p}'" for p in sorted(bloque["prefijos"]))
        return [
            f"Filtros — hoja '{EtapaMonivoi.HOJA_MONIVOI}':",
            f"{nombre_ref2} = {prefijos}; {columnas['uuid']} con dato",
            f"{columnas['centro']} vacío → renglón '{EtapaMonivoi.ETIQUETA_SIN_CENTRO}'",
        ]

    def escribir_mtd(self, ws):
        """
        Write the four step-16 matrices onto the 'MTD' sheet.

        Above each matrix (rows 1-3 of its own column) goes the legend of the
        filters that built it, so the stakeholder can reproduce any figure by
        hand from the 'Monivoi' sheet. Call it AFTER `insertar_referencia2`, so
        the legend can cite the helper column by its final name.

        Args:
            ws: the new (empty) openpyxl worksheet to fill.
        """
        mtd = self.resultado
        # Name the helper column really ended up with (it changes only when the
        # export already had one with that name). Without data there is no
        # helper column and the matrices are empty, so the default name is fine.
        nombre_ref2 = mtd.get("nombre_ref2", self.COL_REFERENCIA2)
        columnas = mtd["columnas"]

        for bloque in self.BLOQUES:
            datos      = mtd["bloques"][bloque["clave"]]
            col_centro = bloque["col_centro"]
            col_valor  = bloque["col_valor"]

            # Filter legend (rows 1-3, in the centro column of this block).
            Comun.escribir_filtros(ws, self.FILA_FILTROS, col_centro,
                                   self.leyenda_filtros(bloque, nombre_ref2, columnas))

            # Title (row 4) and headers (row 5).
            ws.cell(row=self.FILA_TITULO, column=col_centro,
                    value=bloque["titulo"]).font = Font(bold=True)
            ws.cell(row=self.FILA_ENCABEZADO, column=col_centro,
                    value=self.ENCABEZADO_CENTRO).font = Font(bold=True)
            ws.cell(row=self.FILA_ENCABEZADO, column=col_valor,
                    value=self.ENCABEZADO_VALOR).font = Font(bold=True)

            # Body: one row per centro.
            fila = self.FILA_INICIO
            for centro, valor in datos["filas"]:
                ws.cell(row=fila, column=col_centro, value=centro)
                ws.cell(row=fila, column=col_valor,
                        value=valor).number_format = Comun.FORMATO_NUMERO
                fila += 1

            # Rows that passed the filter but had no centro: shown apart, so
            # the 'Grand Total' still reconciles with the filter.
            if datos["sin_centro_n"]:
                ws.cell(row=fila, column=col_centro,
                        value=self.ETIQUETA_SIN_CENTRO).font = Font(italic=True)
                celda = ws.cell(row=fila, column=col_valor,
                                value=datos["sin_centro_importe"])
                celda.number_format = Comun.FORMATO_NUMERO
                celda.font = Font(italic=True)
                fila += 1

            # 'Grand Total' only if there was at least one data row.
            if fila > self.FILA_INICIO:
                ws.cell(row=fila, column=col_centro,
                        value=self.ETIQUETA_TOTAL).font = Font(bold=True)
                celda_total = ws.cell(row=fila, column=col_valor, value=datos["total"])
                celda_total.number_format = Comun.FORMATO_NUMERO
                celda_total.font = Font(bold=True)

            # Comfortable widths to read centro and amount.
            ws.column_dimensions[get_column_letter(col_centro)].width = 12
            ws.column_dimensions[get_column_letter(col_valor)].width = 20

    # ------------------------------------------------------------------
    # Invoice list for the VBRL download
    # ------------------------------------------------------------------
    @staticmethod
    def facturas_unicas(monivoi):
        """
        ALL the distinct values of the Monivoi 'Factura' column, in order of
        first appearance: the list that the VBRL download pastes into SAP's
        multiple-selection popup.

        Deliberately NOT filtered by UUID or by 'Referencia': the stakeholder
        asked for every value of the column. Only "there is nothing here" cells
        are discarded (blank, the NULL-like sentinels of UUID_VACIOS and the
        all-zero "initial value" SAP leaves on rows that were never invoiced),
        because pasting those into SAP would select an invoice that does not
        exist. A '.0' left behind by the xlsx→csv conversion ('233982822.0') is
        removed, so SAP receives the real invoice number.

        Returns:
            tuple: (facturas list[str], filas_leidas int, col_factura str).
            Empty list when the Monivoi is the 'Sin datos' placeholder.

        Raises:
            ValueError: the Monivoi has data but no 'Factura' column.
        """
        if Comun.es_sin_datos(monivoi):
            return [], 0, EtapaMonivoi.COL_FACTURA_CANDIDATAS[0]
        col_factura = Comun.col_por_candidatos(
            monivoi, EtapaMonivoi.COL_FACTURA_CANDIDATAS, "factura", origen="Monivoi")

        vistas = {}
        for valor in monivoi[col_factura]:
            texto = str(valor).strip()
            if re.fullmatch(r"[0-9]+\.0+", texto):
                texto = texto.split(".", 1)[0]
            if texto.upper() in EtapaMonivoi.UUID_VACIOS or not texto.strip("0"):
                continue
            vistas.setdefault(texto, None)
        return list(vistas), len(monivoi), col_factura

    @staticmethod
    def leer_facturas(ruta_monivoi):
        """Read the Monivoi CSV and return `facturas_unicas` of it."""
        return EtapaMonivoi.facturas_unicas(Comun.leer_csv(ruta_monivoi))


# ======================================================================
# Stage: cross 'TD' × 'MTD' — live formulas on the 'TD' sheet
# ======================================================================
class EtapaCruceTDMTD:
    """
    The TD × MTD cross: LIVE Excel formulas written on the 'TD' sheet
    (Factura Monitor / Diferencia / Observación), plus the warnings those
    formulas cannot show on their own.

    The cross lives apart from both sides because it is the reconciliation
    itself: its formulas look up MTD values from the TD rows. It reads the TD
    layout (EtapaTD.FACTURAS / NOTAS / PEP) and the MTD layout
    (EtapaMonivoi.BLOQUES), so a layout change keeps working as long as it is
    made in those constants.
    """

    # Three new columns to the right of each 'TD' block, with LIVE Excel
    # FORMULAS (not values), so the stakeholder can audit them and drag them
    # down:
    #
    #     Facturas     (A..D) ↔ MTD 'SG' (A/B) → E 'Factura Monitor', F 'Diferencia', G 'Observación'
    #     Notas        (L..O) ↔ MTD 'RG' (D/E) → P, Q, R
    #     PEP vouchers (W..X) ↔ MTD 'PP' (K/L) → Y, Z, AA
    #
    # The formula goes in EVERY centro row AND in the block's 'Grand Total' row:
    # there the XLOOKUP finds the MTD 'Grand Total' and compares total against
    # total, including centros that only exist in the monitor and the 'Sin
    # centro' row (that is why TD and MTD share Comun.ETIQUETA_GRAND_TOTAL).
    ENCABEZADO_MONITOR     = "Factura Monitor"
    ENCABEZADO_DIFERENCIA  = "Diferencia"
    ENCABEZADO_OBSERVACION = "Observación"
    TEXTO_SIN_TIMBRE       = Comun.TEXTO_SIN_TIMBRE
    TEXTO_SIN_DIFERENCIA   = "Sin diferencia"

    # 'Observación' tolerance: up to one cent counts as 'Sin diferencia'.
    TOLERANCIA = 0.01

    # 'revisar_si' marks the ONLY direction that leaves the Observación blank
    # for review: when accounting exceeds the monitor. The opposite direction
    # shows 'Sin diferencia' ON PURPOSE (confirmed business rule).
    #     '<' → the Diferencia is compared against -tolerance  (ROUND(F6,2)<-0.01)
    #     '>' → the Diferencia is compared against +tolerance  (ROUND(Q6,2)>0.01)
    CRUCES = (
        {   # Facturas ↔ SG. Diferencia = D + E: invoice accounting comes in as
            # negative, so a perfect match gives zero.
            "etiqueta":        "Facturas ↔ SG",
            "td_clave":        "facturas",
            "col_cedis":       EtapaTD.FACTURAS["col_cedis"],        # A
            "col_contab":      EtapaTD.FACTURAS["col_grand_total"],  # D
            "col_monitor":     5,                                    # E
            "col_diferencia":  6,                                    # F
            "col_observacion": 7,                                    # G
            "mtd_clave":       "sg",                                 # MTD A / B
            "operador":        "+",
            "revisar_si":      "<",
        },
        {   # Notas de Crédito ↔ RG. Diferencia = O - P.
            "etiqueta":        "Notas de Crédito ↔ RG",
            "td_clave":        "notas",
            "col_cedis":       EtapaTD.NOTAS["col_cedis"],           # L
            "col_contab":      EtapaTD.NOTAS["col_grand_total"],     # O
            "col_monitor":     16,                                   # P
            "col_diferencia":  17,                                   # Q
            "col_observacion": 18,                                   # R
            "mtd_clave":       "rg",                                 # MTD D / E
            "operador":        "-",
            "revisar_si":      ">",
        },
        {   # PEP vouchers ↔ PP. Diferencia = X - Y.
            "etiqueta":        "PEP vouchers ↔ PP",
            "td_clave":        "pep",
            "col_cedis":       EtapaTD.PEP["col_cedis"],             # W
            "col_contab":      EtapaTD.PEP["col_valor"],             # X
            "col_monitor":     25,                                   # Y
            "col_diferencia":  26,                                   # Z
            "col_observacion": 27,                                   # AA
            "mtd_clave":       "pp",                                 # MTD K / L
            "operador":        "-",
            "revisar_si":      ">",
        },
    )

    def __init__(self, td_resumen, mtd):
        """
        Args:
            td_resumen (dict): what `EtapaTD.preparar` returned.
            mtd (dict): what `EtapaMonivoi.preparar` returned.
        """
        self.td_resumen = td_resumen
        self.mtd = mtd

    # ------------------------------------------------------------------
    # Formulas
    # ------------------------------------------------------------------
    @staticmethod
    def formulas(cruce, fila):
        """
        The three formulas of ONE row of the TD × MTD cross, in this order:
        (Factura Monitor, Diferencia, Observación). Example (Facturas, row 6):

            E6: =IFERROR(XLOOKUP(A6,MTD!A:A,MTD!B:B),"Sin timbre")
            F6: =IFERROR(D6+E6,"Sin timbre")
            G6: =IF(F6="Sin timbre","Sin timbre",IF(ROUND(F6,2)<-0.01,"","Sin diferencia"))

        In the file XLOOKUP is written as `_xlfn.XLOOKUP`: functions newer than
        Excel 2007 need that prefix, otherwise Excel shows #NAME?.
        """
        if cruce["operador"] not in ("+", "-"):
            raise ValueError(
                f"Cruce TD ({cruce['etiqueta']}): operador de diferencia inválido "
                f"'{cruce['operador']}' (se esperaba '+' o '-')."
            )
        if cruce["revisar_si"] not in ("<", ">"):
            raise ValueError(
                f"Cruce TD ({cruce['etiqueta']}): comparación inválida "
                f"'{cruce['revisar_si']}' (se esperaba '<' o '>')."
            )

        letra      = get_column_letter
        bloque_mtd = EtapaMonivoi.POR_CLAVE[cruce["mtd_clave"]]
        hoja_mtd   = Comun.ref_hoja(EtapaMonivoi.HOJA_MTD)
        col_centro = letra(bloque_mtd["col_centro"])
        col_valor  = letra(bloque_mtd["col_valor"])

        cedis   = f"{letra(cruce['col_cedis'])}{fila}"
        contab  = f"{letra(cruce['col_contab'])}{fila}"
        monitor = f"{letra(cruce['col_monitor'])}{fila}"
        dif     = f"{letra(cruce['col_diferencia'])}{fila}"

        sin_timbre     = Comun.texto_formula(EtapaCruceTDMTD.TEXTO_SIN_TIMBRE)
        sin_diferencia = Comun.texto_formula(EtapaCruceTDMTD.TEXTO_SIN_DIFERENCIA)
        umbral = (-EtapaCruceTDMTD.TOLERANCIA if cruce["revisar_si"] == "<"
                  else EtapaCruceTDMTD.TOLERANCIA)

        f_monitor = (
            f"=IFERROR(_xlfn.XLOOKUP({cedis},{hoja_mtd}!{col_centro}:{col_centro},"
            f"{hoja_mtd}!{col_valor}:{col_valor}),{sin_timbre})"
        )
        f_diferencia = f"=IFERROR({contab}{cruce['operador']}{monitor},{sin_timbre})"
        f_observacion = (
            f"=IF({dif}={sin_timbre},{sin_timbre},"
            f"IF(ROUND({dif},2){cruce['revisar_si']}{umbral:g},\"\",{sin_diferencia}))"
        )
        return f_monitor, f_diferencia, f_observacion

    def escribir(self, ws_td):
        """
        Write the cross against 'MTD' onto the 'TD' sheet: three new columns per
        block (see CRUCES), with the header in row 5 and a formula in every
        centro row PLUS the 'Grand Total' row.

        Headers are always written (like the rest of 'TD'); if the block has no
        centros, it has no total row either and no formula is written. Amount
        cells get the thousands/2-decimals format ("Sin timbre" looks the same,
        because it is text).
        """
        td_resumen = self.td_resumen
        negrita = Font(bold=True)
        for cruce in self.CRUCES:
            columnas = (
                (cruce["col_monitor"],     self.ENCABEZADO_MONITOR),
                (cruce["col_diferencia"],  self.ENCABEZADO_DIFERENCIA),
                (cruce["col_observacion"], self.ENCABEZADO_OBSERVACION),
            )
            for col, encabezado in columnas:
                ws_td.cell(row=EtapaTD.FILA_ENCABEZADO, column=col, value=encabezado).font = negrita
                ws_td.column_dimensions[get_column_letter(col)].width = max(16, len(encabezado) + 3)

            n_cedis = len(td_resumen[cruce["td_clave"]]["cedis"])
            if n_cedis == 0:
                continue
            fila_total = EtapaTD.FILA_INICIO + n_cedis   # same formula as EtapaTD.escribir

            for fila in range(EtapaTD.FILA_INICIO, fila_total + 1):
                f_monitor, f_diferencia, f_observacion = self.formulas(cruce, fila)
                celda_monitor = ws_td.cell(row=fila, column=cruce["col_monitor"], value=f_monitor)
                celda_dif     = ws_td.cell(row=fila, column=cruce["col_diferencia"], value=f_diferencia)
                celda_obs     = ws_td.cell(row=fila, column=cruce["col_observacion"], value=f_observacion)
                celda_monitor.number_format = Comun.FORMATO_NUMERO
                celda_dif.number_format     = Comun.FORMATO_NUMERO
                if fila == fila_total:
                    for celda in (celda_monitor, celda_dif, celda_obs):
                        celda.font = negrita

    # ------------------------------------------------------------------
    # Controls the formulas do not show
    # ------------------------------------------------------------------
    def avisos(self):
        """
        Cross checks that the formulas do NOT show on their own.

        The XLOOKUP goes from TD to MTD: a centro that the monitor did stamp
        but that has no row in TD never shows up in a row of its own; it only
        weighs in the 'Grand Total'. Those centros are listed here so the
        stakeholder knows WHAT explains the difference in the total.

        Matching works like XLOOKUP: case-insensitive and keeping the FIRST
        match. That also uncovers the rare case of centros that only differ in
        case ('80AB' / '80ab'): in MTD the second one would stay unmatched, and
        in TD both rows would get the same monitor amount.

        Returns:
            list[str]: warnings ready for the GUI summary.
        """
        td_resumen = self.td_resumen
        mtd = self.mtd
        avisos = []
        for cruce in self.CRUCES:
            etiqueta    = f"Cruce TD ({cruce['etiqueta']})"
            cedis_td    = td_resumen[cruce["td_clave"]]["cedis"]
            centros_mtd = [centro for centro, _ in mtd["bloques"][cruce["mtd_clave"]]["filas"]]

            # MTD row that XLOOKUP would find for each key (the first one).
            primero_mtd = {}
            for centro in centros_mtd:
                primero_mtd.setdefault(centro.casefold(), centro)
            claves_td  = {centro.casefold() for centro in cedis_td}
            alcanzados = {primero_mtd[clave] for clave in claves_td if clave in primero_mtd}
            sin_renglon = [centro for centro in centros_mtd if centro not in alcanzados]

            if sin_renglon:
                listado = ", ".join(sin_renglon[:5]) + ("..." if len(sin_renglon) > 5 else "")
                if cedis_td:
                    avisos.append(
                        f"{etiqueta}: {len(sin_renglon):,} centro(s) del monitor sin "
                        f"renglón en TD ({listado}); su importe sólo pesa en el "
                        f"'{Comun.ETIQUETA_GRAND_TOTAL}'."
                    )
                else:
                    avisos.append(
                        f"{etiqueta}: TD no trae centros en este bloque, pero el "
                        f"monitor sí ({listado}); no hay contra qué cruzar."
                    )

            grupos_td = {}
            for centro in cedis_td:
                grupos_td.setdefault(centro.casefold(), []).append(centro)
            choques = [" / ".join(grupo) for grupo in grupos_td.values() if len(grupo) > 1]
            if choques:
                avisos.append(
                    f"{etiqueta}: en TD hay centros que sólo cambian en mayúsculas/"
                    f"minúsculas ({'; '.join(choques)}); XLOOKUP les asigna el mismo "
                    f"importe del monitor."
                )
        return avisos


# ======================================================================
# Stage: VBRL (ZSE16, table VBRL) — what the monitor invoiced, per invoice
# ======================================================================
class EtapaVBRL:
    """
    The 'VBRL' sheet: the table downloaded with ZSE16 for the invoices listed in
    the Monivoi 'Factura' column. It is the source of the 'Factura Monitor'
    column of 'TD.Diferencias_Referencias'.

    Which columns matter is NOT guessed from names: the cross uses
    `VBRL!D:D` (invoice) and `VBRL!E:E` (amount), exactly the columns of the
    stakeholder's own XLOOKUP, so COL_FACTURA / COL_IMPORTE are positions.
    Their headers are recorded in the result so the cross can cite them.

    The sheet is the downloaded CSV with those two columns converted to REAL
    numbers. That conversion is not cosmetic: XLOOKUP is type-strict, so an
    invoice stored as text in VBRL!D:D never matches the number VALUE() builds
    from the payment reference. (The CSV reaches us as text, because every
    download is read as text.)
    """

    HOJA_VBRL = "VBRL"

    # 1-based positions of the VBRL columns the cross reads (D and E). The
    # download checks the same positions (Descargas_SAP._VBRL_MIN_COLUMNAS and
    # _VBRL_COL_FACTURA_IDX): change both places together.
    COL_FACTURA = 4
    COL_IMPORTE = 5

    # Excel keeps 15 significant digits: a longer invoice number would be
    # silently rounded, so it stays as text (and shows up in the warnings).
    MAX_DIGITOS = 15

    # Two amounts for the same invoice closer than this are the same amount.
    TOLERANCIA_DUPLICADOS = 0.005

    def __init__(self, vbrl, sin_datos):
        """
        Args:
            vbrl (pd.DataFrame): the VBRL CSV as read by Comun.leer_csv.
            sin_datos (bool): the CSV is the 'Sin datos' placeholder.
        """
        self.vbrl = vbrl
        self.sin_datos = sin_datos
        self.resultado = None

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    @staticmethod
    def convertir_factura(valor):
        """
        Value of a VBRL invoice cell as it must be written in VBRL!D:D:

            '233982822'   -> 233982822   (number)
            '0233982822'  -> 233982822   (leading zeros dropped, like VALUE() does)
            '233982822.0' -> 233982822   (the float the xlsx→csv conversion left)
            ''            -> None        (blank cell)
            'ABC123'      -> 'ABC123'    (text: a number lookup can never find it)

        More than MAX_DIGITOS digits also stays as text.
        """
        texto = str(valor).strip()
        if texto == "":
            return None
        if re.fullmatch(r"[0-9]+(?:\.0+)?", texto):
            digitos = texto.split(".", 1)[0].lstrip("0") or "0"
            if len(digitos) <= EtapaVBRL.MAX_DIGITOS:
                return int(digitos)
        return texto

    # ------------------------------------------------------------------
    # Computation
    # ------------------------------------------------------------------
    def preparar(self):
        """
        Convert the invoice and amount columns to numbers and index the sheet
        the way XLOOKUP will read it (first match wins).

        Raises:
            ValueError: the file has data but fewer than 5 columns (the layout
                chosen in the VBRL download is not the expected one), or an
                amount cannot be interpreted (same fail-loud rule as the Gallo
                and Monivoi amounts).

        Returns:
            dict: df (what the 'VBRL' sheet shows), sin_datos, filas, mapa
            ({invoice number: amount}, first row of each invoice), avisos and
            encabezados (headers of columns D and E).
        """
        vbrl = self.vbrl
        letra_factura = get_column_letter(self.COL_FACTURA)
        letra_importe = get_column_letter(self.COL_IMPORTE)

        if self.sin_datos or vbrl.empty:
            self.resultado = {
                "df":          vbrl,
                "sin_datos":   True,
                "filas":       0,
                "mapa":        {},
                "avisos": [
                    "VBRL: el archivo vino sin datos; todos los renglones de "
                    f"TD.Diferencias_Referencias quedarán como "
                    f"'{Comun.TEXTO_SIN_TIMBRE}'."
                ],
                "encabezados": ("", ""),
            }
            return self.resultado

        if vbrl.shape[1] < self.COL_IMPORTE:
            raise ValueError(
                f"El archivo de VBRL trae sólo {vbrl.shape[1]} columna(s) y el cruce "
                f"necesita la {letra_factura} (factura) y la {letra_importe} "
                f"(importe). Columnas: {list(vbrl.columns)}. Revisa el layout que "
                f"se elige en la descarga de VBRL."
            )

        col_factura = vbrl.columns[self.COL_FACTURA - 1]
        col_importe = vbrl.columns[self.COL_IMPORTE - 1]

        facturas = [self.convertir_factura(valor)
                    for valor in vbrl.iloc[:, self.COL_FACTURA - 1]]
        etiqueta_importe = f"{col_importe} (VBRL, columna {letra_importe})"
        importes = [
            None if str(valor).strip() == ""
            else Comun.parse_importe_us(valor, columna=etiqueta_importe)
            for valor in vbrl.iloc[:, self.COL_IMPORTE - 1]
        ]

        # Index the sheet the way XLOOKUP reads it: the FIRST row of each
        # invoice wins. Repeated invoices are counted so they can be reported.
        mapa = {}
        repeticiones = {}
        importes_distintos = set()
        en_blanco = 0
        no_numericas = []
        for factura, importe in zip(facturas, importes):
            if factura is None:
                en_blanco += 1
                continue
            if not isinstance(factura, int):
                no_numericas.append(factura)
                continue
            valor = 0.0 if importe is None else importe
            if factura in mapa:
                repeticiones[factura] += 1
                if abs(mapa[factura] - valor) > self.TOLERANCIA_DUPLICADOS:
                    importes_distintos.add(factura)
            else:
                mapa[factura] = valor
                repeticiones[factura] = 1

        avisos = []
        con_factura = len(facturas) - en_blanco
        if no_numericas:
            ejemplos = ", ".join(dict.fromkeys(no_numericas[:3]))
            avisos.append(
                f"VBRL: {len(no_numericas):,} de {con_factura:,} fila(s) con la "
                f"columna {letra_factura} ('{col_factura}') que no es un número de "
                f"factura (ej.: {ejemplos}); el XLOOKUP no las puede encontrar. "
                f"Si son casi todas, revisa que el layout del VBRL sea el esperado."
            )
        if en_blanco:
            avisos.append(
                f"VBRL: {en_blanco:,} fila(s) sin valor en la columna "
                f"{letra_factura} ('{col_factura}')."
            )
        repetidas = sum(1 for n in repeticiones.values() if n > 1)
        if repetidas:
            avisos.append(
                f"VBRL: {repetidas:,} factura(s) aparecen en más de una fila "
                f"({len(importes_distintos):,} con importes distintos); XLOOKUP "
                f"toma sólo la primera."
            )

        df = vbrl.copy()
        # By position: the export may repeat a header name.
        df.isetitem(self.COL_FACTURA - 1,
                    pd.Series(facturas, index=df.index, dtype=object))
        df.isetitem(self.COL_IMPORTE - 1,
                    pd.Series(importes, index=df.index, dtype="float64"))

        self.resultado = {
            "df":          df,
            "sin_datos":   False,
            "filas":       len(df),
            "mapa":        mapa,
            "avisos":      avisos,
            "encabezados": (str(col_factura), str(col_importe)),
        }
        return self.resultado

    # ------------------------------------------------------------------
    # Writing
    # ------------------------------------------------------------------
    def escribir(self, writer):
        """
        Write the 'VBRL' sheet through the workbook's pandas writer: the
        downloaded table, with the invoice (D) and amount (E) columns as numbers
        and wide enough to show them in full (a 10-digit number in a narrow
        column turns into 7.05E+09).
        """
        df = self.resultado["df"]
        df.to_excel(writer, sheet_name=self.HOJA_VBRL, index=False)
        if not self.resultado["sin_datos"]:
            ws = writer.book[self.HOJA_VBRL]
            ws.column_dimensions[get_column_letter(self.COL_FACTURA)].width = 16
            ws.column_dimensions[get_column_letter(self.COL_IMPORTE)].width = 18


# ======================================================================
# Stage: 'TD.Diferencias_Referencias' — Facturas, document by document
# ======================================================================
class EtapaTDReferencias:
    """
    The micro version of the 'TD' FACTURAS block: the same population of Gallo
    rows, but grouped down to Centro × Nº documento × Referencia de pago, so a
    difference found by centro in 'TD' can be traced to the document that
    causes it.

    Pivot-style table (values, as a PivotTable would show them) with two extra
    columns OUTSIDE the pivot, to its right, that hold LIVE formulas:

        A  Centro                       ┐
        B  Nº documento                 │ pivot rows (flat, labels repeated)
        C  Referencia de pago           │
        D  Importe en moneda local      ┘ sum (the pivot value)
        E  Factura Monitor   XLOOKUP of the payment reference in VBRL!D:D → VBRL!E:E
        F  Diferencia        D + E (invoice accounting is negative, so 0 = match)

    Why not a native PivotTable: openpyxl cannot create one, so — like 'TD' —
    the table is written as values. The filters of the pivot are the SAME as the
    FACTURAS block of 'TD' (see `EtapaTD.mascara_bloque`), which is what makes
    this table's Grand Total tie to that block's Grand Total. The header row has
    an AutoFilter, so the stakeholder can filter by 'Diferencia' as in the
    PivotTable.

    The payment reference carries a leading 'C' ('C233982822') while VBRL holds
    the bare number, and Excel's VALUE("C233982822") is #VALUE!. The formula
    therefore drops one leading 'C' before VALUE(); the references that still do
    not look like an invoice number are reported, not hidden.

    Rows that cannot be matched show 'Sin timbre' in both columns (the same text
    the TD × MTD cross uses), instead of #N/A / #VALUE! noise.
    """

    HOJA_TD_REFERENCIAS = "TD.Diferencias_Referencias"

    # 1-based columns of the table.
    COL_CENTRO     = 1   # A
    COL_DOCUMENTO  = 2   # B
    COL_REFERENCIA = 3   # C
    COL_IMPORTE    = 4   # D
    COL_MONITOR    = 5   # E
    COL_DIFERENCIA = 6   # F

    ENCABEZADO_MONITOR    = "Factura Monitor"
    ENCABEZADO_DIFERENCIA = "Diferencia"

    # The leading letter SAP puts in front of the invoice number in the payment
    # reference ('C233982822').
    PREFIJO_REFERENCIA = "C"

    # Reconciliation tolerance against the 'TD' FACTURAS Grand Total (cents).
    TOLERANCIA_CUADRE = 0.01

    # Excel's hard limit of rows per sheet.
    MAX_FILAS_EXCEL = 1_048_576

    AUTOR_COMENTARIOS = "Validación Factura Global"

    def __init__(self, gallo, pasos_gallo, etapa_td, etapa_vbrl):
        """
        Args:
            gallo (pd.DataFrame): the Gallo DataFrame after `EtapaGallo.preparar`.
            pasos_gallo (dict): what `EtapaGallo.preparar` returned (it must have
                been built with `con_ref_pago=True`).
            etapa_td (EtapaTD): the 'TD' stage, already prepared.
            etapa_vbrl (EtapaVBRL): the 'VBRL' stage, already prepared.
        """
        self.gallo = gallo
        self.pasos = pasos_gallo
        self.etapa_td = etapa_td
        self.etapa_vbrl = etapa_vbrl
        self.resultado = None

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    @staticmethod
    def clave_referencia(referencia):
        """
        The invoice number a payment reference stands for, as the formula
        reads it: one leading 'C' dropped, then digits only.

            'C233982822' -> 233982822
            '233982822'  -> 233982822
            'X233982822' -> None   (the formula gives #VALUE!, shown as 'Sin timbre')
            ''           -> None

        Python-side twin of the VALUE(...) in the formula, used only to measure
        and report how many rows will match. (Excel's VALUE is more lenient —
        it also reads '1E3' or a date — so a handful of exotic references could
        match in Excel and not here; the reverse never happens.)
        """
        texto = str(referencia).strip()
        if texto[:1].upper() == EtapaTDReferencias.PREFIJO_REFERENCIA:
            texto = texto[1:]
        if re.fullmatch(r"[0-9]+", texto):
            digitos = texto.lstrip("0") or "0"
            if len(digitos) <= EtapaVBRL.MAX_DIGITOS:
                return int(digitos)
        return None

    @staticmethod
    def formulas(fila):
        """
        The two formulas of ONE row (Factura Monitor, Diferencia). Example, row 6:

            E6: =IF(C6="","Sin timbre",IFERROR(XLOOKUP(VALUE(IF(UPPER(LEFT(C6,1))="C",
                    MID(C6,2,99),C6)),VBRL!D:D,VBRL!E:E),"Sin timbre"))
            F6: =IFERROR(D6+E6,"Sin timbre")

        XLOOKUP is written as `_xlfn.XLOOKUP` in the file (see
        EtapaCruceTDMTD.formulas). The blank check keeps an empty reference
        from reaching VALUE(), which would read it as 0.
        """
        letra = get_column_letter
        referencia = f"{letra(EtapaTDReferencias.COL_REFERENCIA)}{fila}"
        importe    = f"{letra(EtapaTDReferencias.COL_IMPORTE)}{fila}"
        monitor    = f"{letra(EtapaTDReferencias.COL_MONITOR)}{fila}"
        hoja_vbrl  = Comun.ref_hoja(EtapaVBRL.HOJA_VBRL)
        col_factura = letra(EtapaVBRL.COL_FACTURA)
        col_importe = letra(EtapaVBRL.COL_IMPORTE)
        prefijo    = Comun.texto_formula(EtapaTDReferencias.PREFIJO_REFERENCIA)
        sin_timbre = Comun.texto_formula(Comun.TEXTO_SIN_TIMBRE)

        clave = (
            f"VALUE(IF(UPPER(LEFT({referencia},1))={prefijo},"
            f"MID({referencia},2,99),{referencia}))"
        )
        f_monitor = (
            f'=IF({referencia}="",{sin_timbre},IFERROR(_xlfn.XLOOKUP({clave},'
            f"{hoja_vbrl}!{col_factura}:{col_factura},"
            f"{hoja_vbrl}!{col_importe}:{col_importe}),{sin_timbre}))"
        )
        f_diferencia = f"=IFERROR({importe}+{monitor},{sin_timbre})"
        return f_monitor, f_diferencia

    # ------------------------------------------------------------------
    # Computation
    # ------------------------------------------------------------------
    def preparar(self):
        """
        Group the FACTURAS population of 'TD' by Centro × Nº documento ×
        Referencia de pago and measure how it will cross against VBRL.

        Returns:
            dict: filas ([(centro, documento, referencia, importe)], sorted like
            a PivotTable would), total, avisos (list[str]) and the counters that
            the summary reports (sin_timbre_n, sin_timbre_importe, columnas).
        """
        gallo = self.gallo
        pasos = self.pasos
        td = self.etapa_td
        col_doc = pasos["col_doc"]
        col_ref_pago = pasos["col_ref_pago"]
        if col_ref_pago is None:
            raise ValueError(
                "TD.Diferencias_Referencias necesita la columna 'Referencia de "
                "pago' del Gallo: EtapaGallo se tiene que construir con "
                "con_ref_pago=True."
            )

        # The EXACT rows of the FACTURAS block of 'TD'.
        mascara = td.mascara_bloque(EtapaTD.FACTURAS["clases"])
        datos = pd.DataFrame({
            "_centro":  td.centro_norm[mascara],
            "_doc":     gallo.loc[mascara, col_doc].astype(str).str.strip(),
            "_ref":     gallo.loc[mascara, col_ref_pago].astype(str).str.strip(),
            "_importe": pasos["importe_num"][mascara],
        })
        if datos.empty:
            filas = []
        else:
            # groupby sorts the keys: Centro, then document, then reference —
            # the order a PivotTable shows. Round to cents so no float noise
            # leaks into the cells.
            serie = datos.groupby(["_centro", "_doc", "_ref"])["_importe"].sum()
            filas = [(centro, doc, ref, round(float(valor), 2))
                     for (centro, doc, ref), valor in serie.items()]

        if EtapaTD.FILA_INICIO + len(filas) > self.MAX_FILAS_EXCEL:
            raise ValueError(
                f"TD.Diferencias_Referencias tendría {len(filas):,} renglones y no "
                f"caben en una hoja de Excel ({self.MAX_FILAS_EXCEL:,} filas). "
                f"Consolida un rango de fechas más corto."
            )

        total = round(sum(importe for _, _, _, importe in filas), 2)

        # --- Controls: how will each row cross against VBRL? ---------------
        mapa = self.etapa_vbrl.resultado["mapa"]
        sin_ref_n = sin_ref_importe = 0
        invalidas_n = invalidas_importe = 0
        faltan_n = faltan_importe = 0
        ejemplos_invalidas = []
        conteo_encontradas = {}
        for _, _, ref, importe in filas:
            if ref == "":
                sin_ref_n += 1
                sin_ref_importe += importe
                continue
            clave = self.clave_referencia(ref)
            if clave is None:
                invalidas_n += 1
                invalidas_importe += importe
                if len(ejemplos_invalidas) < 3 and ref not in ejemplos_invalidas:
                    ejemplos_invalidas.append(ref)
            elif clave in mapa:
                conteo_encontradas[clave] = conteo_encontradas.get(clave, 0) + 1
            else:
                faltan_n += 1
                faltan_importe += importe

        sin_timbre_n = sin_ref_n + invalidas_n + faltan_n
        sin_timbre_importe = round(sin_ref_importe + invalidas_importe + faltan_importe, 2)
        repetidas = {clave: n for clave, n in conteo_encontradas.items() if n > 1}
        importe_repetido = round(sum(mapa[clave] * (n - 1) for clave, n in repetidas.items()), 2)

        texto_sin_timbre = Comun.TEXTO_SIN_TIMBRE
        etiqueta = self.HOJA_TD_REFERENCIAS
        avisos = []
        if sin_timbre_n:
            partes = []
            if sin_ref_n:
                partes.append(f"{sin_ref_n:,} sin Referencia de pago")
            if invalidas_n:
                partes.append(
                    f"{invalidas_n:,} con una referencia que no parece factura "
                    f"(ej.: {', '.join(ejemplos_invalidas)})"
                )
            if faltan_n:
                partes.append(f"{faltan_n:,} que no están en VBRL")
            avisos.append(
                f"{etiqueta}: {sin_timbre_n:,} de {len(filas):,} renglón(es) quedan "
                f"'{texto_sin_timbre}' (importe contable {sin_timbre_importe:,.2f}): "
                f"{'; '.join(partes)}."
            )
        if repetidas:
            avisos.append(
                f"{etiqueta}: {len(repetidas):,} factura(s) del monitor aparecen en "
                f"más de un renglón ({sum(repetidas.values()):,} renglones); "
                f"'{self.ENCABEZADO_MONITOR}' repite su importe en cada uno y el "
                f"Grand Total lo cuenta varias veces (importe repetido "
                f"{importe_repetido:,.2f})."
            )

        total_td = self.etapa_td.resultado["total_facturas"]
        if abs(total - total_td) > self.TOLERANCIA_CUADRE:
            avisos.append(
                f"{etiqueta}: el Grand Total contable ({total:,.2f}) no cuadra con "
                f"el de {EtapaTD.HOJA_TD} FACTURAS ({total_td:,.2f}); diferencia "
                f"{round(total - total_td, 2):,.2f}."
            )

        self.resultado = {
            "filas":              filas,
            "total":              total,
            "avisos":             avisos,
            "sin_timbre_n":       sin_timbre_n,
            "sin_timbre_importe": sin_timbre_importe,
            "columnas": {
                "centro":     pasos["col_centro"],
                "documento":  col_doc,
                "referencia": col_ref_pago,
            },
        }
        return self.resultado

    # ------------------------------------------------------------------
    # Writing
    # ------------------------------------------------------------------
    def escribir(self, ws):
        """
        Write the 'TD.Diferencias_Referencias' sheet.

        Layout (same rows as 'TD'): rows 1-3 = filter legend (the FACTURAS one),
        row 4 = title, row 5 = headers (with an AutoFilter), row 6+ = one row per
        Centro × Nº documento × Referencia de pago, and a 'Grand Total' row right
        below. The panes are frozen under the headers because this table is long.

        Columns A:D are values (the pivot); E:F are live formulas (see
        `formulas`). The 'Grand Total' row sums D as a value and E as a formula;
        its 'Diferencia' is D + E, i.e. total against total, so it includes the
        rows that show 'Sin timbre' (summing the visible Diferencia cells would
        leave those out). The header comments say so.

        Args:
            ws: the new (empty) openpyxl worksheet to fill.
        """
        resultado = self.resultado
        filas = resultado["filas"]
        columnas = resultado["columnas"]
        td_columnas = self.etapa_td.resultado["columnas"]
        vbrl = self.etapa_vbrl.resultado
        negrita = Font(bold=True)
        formato = Comun.FORMATO_NUMERO

        # Filter legend (rows 1-3) and block title (row 4).
        Comun.escribir_filtros(ws, EtapaTD.FILA_FILTROS, self.COL_CENTRO,
                               EtapaTD.leyenda_filtros(EtapaTD.FACTURAS, td_columnas))
        ws.cell(row=4, column=self.COL_CENTRO,
                value=f"{EtapaTD.FACTURAS['titulo']} — {columnas['documento']} × "
                      f"{columnas['referencia']}").font = negrita

        # Headers (row 5).
        encabezados = (
            (self.COL_CENTRO,     columnas["centro"]),
            (self.COL_DOCUMENTO,  columnas["documento"]),
            (self.COL_REFERENCIA, columnas["referencia"]),
            (self.COL_IMPORTE,    Comun.ENCABEZADO_IMPORTE),
            (self.COL_MONITOR,    self.ENCABEZADO_MONITOR),
            (self.COL_DIFERENCIA, self.ENCABEZADO_DIFERENCIA),
        )
        for col, texto in encabezados:
            ws.cell(row=EtapaTD.FILA_ENCABEZADO, column=col, value=texto).font = negrita

        # Where the formula looks: VBRL!D:D (invoice) -> VBRL!E:E (amount), citing
        # the real headers of this download when there are any.
        enc_factura, enc_importe = vbrl["encabezados"]
        col_vbrl_factura = get_column_letter(EtapaVBRL.COL_FACTURA)
        col_vbrl_importe = get_column_letter(EtapaVBRL.COL_IMPORTE)
        cita_factura = f" ('{enc_factura}')" if enc_factura else ""
        cita_importe = f" ('{enc_importe}')" if enc_importe else ""
        self._comentar(
            ws.cell(row=EtapaTD.FILA_ENCABEZADO, column=self.COL_MONITOR),
            f"Fórmula en cada renglón: quita la 'C' inicial de la Referencia de pago, "
            f"la convierte en número y la busca en "
            f"{EtapaVBRL.HOJA_VBRL}!{col_vbrl_factura}:{col_vbrl_factura}{cita_factura}; "
            f"regresa {EtapaVBRL.HOJA_VBRL}!{col_vbrl_importe}:{col_vbrl_importe}"
            f"{cita_importe}. Si no la encuentra (o no hay referencia): "
            f"'{Comun.TEXTO_SIN_TIMBRE}'.",
        )
        self._comentar(
            ws.cell(row=EtapaTD.FILA_ENCABEZADO, column=self.COL_DIFERENCIA),
            f"Contabilidad + Monitor. Las facturas vienen en negativo en contabilidad, "
            f"así que 0 = cuadra. En el Grand Total se compara total contra total, "
            f"por lo que incluye los renglones '{Comun.TEXTO_SIN_TIMBRE}'; la suma de "
            f"las celdas visibles de Diferencia no los incluye.",
        )

        # Body: the pivot values (A:D) + the two formula columns (E:F).
        fila = EtapaTD.FILA_INICIO
        for centro, documento, referencia, importe in filas:
            ws.cell(row=fila, column=self.COL_CENTRO, value=centro)
            ws.cell(row=fila, column=self.COL_DOCUMENTO, value=documento)
            # A blank reference stays an empty cell (what a PivotTable shows as
            # "(blank)"); the formula handles it.
            ws.cell(row=fila, column=self.COL_REFERENCIA, value=referencia or None)
            ws.cell(row=fila, column=self.COL_IMPORTE, value=importe).number_format = formato
            f_monitor, f_diferencia = self.formulas(fila)
            ws.cell(row=fila, column=self.COL_MONITOR, value=f_monitor).number_format = formato
            ws.cell(row=fila, column=self.COL_DIFERENCIA, value=f_diferencia).number_format = formato
            fila += 1

        # 'Grand Total' row, only when there is at least one data row.
        if filas:
            ultima = fila - 1
            col_d = get_column_letter(self.COL_IMPORTE)
            col_e = get_column_letter(self.COL_MONITOR)
            ws.cell(row=fila, column=self.COL_CENTRO,
                    value=Comun.ETIQUETA_GRAND_TOTAL).font = negrita
            celdas_total = (
                ws.cell(row=fila, column=self.COL_IMPORTE, value=resultado["total"]),
                ws.cell(row=fila, column=self.COL_MONITOR,
                        value=f"=ROUND(SUM({col_e}{EtapaTD.FILA_INICIO}:{col_e}{ultima}),2)"),
                ws.cell(row=fila, column=self.COL_DIFERENCIA,
                        value=f"=ROUND({col_d}{fila}+{col_e}{fila},2)"),
            )
            for celda in celdas_total:
                celda.number_format = formato
                celda.font = negrita

            # AutoFilter over the header + data rows (NOT the total row), like the
            # funnel the stakeholder uses on 'Diferencia'.
            ws.auto_filter.ref = (
                f"{get_column_letter(self.COL_CENTRO)}{EtapaTD.FILA_ENCABEZADO}:"
                f"{get_column_letter(self.COL_DIFERENCIA)}{ultima}"
            )

        ws.freeze_panes = f"A{EtapaTD.FILA_INICIO}"

        for letra, ancho in {"A": 12, "B": 16, "C": 20, "D": 22, "E": 20, "F": 20}.items():
            ws.column_dimensions[letra].width = ancho

    @classmethod
    def _comentar(cls, celda, texto):
        """Attach a header comment, sized so the whole text is visible."""
        comentario = Comment(texto, cls.AUTOR_COMENTARIOS)
        comentario.width = 420
        comentario.height = 150
        celda.comment = comentario


# ======================================================================
# Helpers of the entry point
# ======================================================================
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


def leer_facturas_monivoi(ruta_monivoi):
    """
    The invoice numbers (distinct values of the Monivoi 'Factura' column) that
    the VBRL download pastes into SAP. See `EtapaMonivoi.facturas_unicas`.

    Returns:
        tuple: (facturas list[str], filas_leidas int, nombre_columna str).
    """
    return EtapaMonivoi.leer_facturas(ruta_monivoi)


# ======================================================================
# Orchestrator
# ======================================================================
class Consolidador:
    """
    Runs the whole consolidation: reads the CSVs, runs every stage in order and
    writes the workbook. The stages themselves hold the logic; this class only
    decides WHICH ones run, in WHAT order, and how their results are put
    together:

        read Gallo / Monivoi / (VBRL)
        EtapaGallo        steps 2-9 + SG/RG override         → 'Gallo', 'Validación - Doc.'
        EtapaTD           'TD' matrices                      → 'TD'
        EtapaMonivoi      'MTD' matrices + 'Referencia 2'    → 'Monivoi', 'MTD'
        EtapaCruceTDMTD   live TD × MTD formulas             → (on 'TD')
        EtapaVBRL         VBRL table as numbers              → 'VBRL'          [only with VBRL]
        EtapaTDReferencias Facturas by document vs VBRL      → 'TD.Diferencias_Referencias' [only with VBRL]
        EtapaVentas       'Venta DSD'                        → that sheet
        write workbook

    The VBRL is OPTIONAL on purpose: a consolidation of an old date, from before
    the VBRL download existed, must keep working. Without it the workbook simply
    has no 'VBRL' and no 'TD.Diferencias_Referencias' sheet.
    """

    def __init__(self, ruta_gallo, ruta_monivoi, ruta_output, nombre_salida=None,
                 callback_status=None, fechas_seleccionadas=None, ruta_vbrl=None):
        """
        Args:
            ruta_gallo   (str): Path to the Gallo CSV downloaded from SAP (FAGLL03).
            ruta_monivoi (str): Path to the Monivoi CSV downloaded from SAP (ZLMXCOM).
            ruta_output  (str): Folder where the consolidated workbook is saved.
            nombre_salida (str|None): File name of the final .xlsx. If None →
                'Consolidado_<Gallo CSV name>.xlsx'.
            callback_status (function|None): Reports progress to the GUI.
            fechas_seleccionadas (list[str]|None): 'DD.MM.YYYY' dates the user
                picked in the GUI (one in single-day mode, every day of the
                range in range mode). They guarantee those days show up in
                'Venta DSD' even without amounts, and they
                break ties when detecting the CSV date format. If None, rows
                come from the data only.
            ruta_vbrl (str|None): Path to the VBRL CSV (ZSE16). If given, the
                workbook also gets the 'VBRL' and 'TD.Diferencias_Referencias'
                sheets.
        """
        self.ruta_gallo = ruta_gallo
        self.ruta_monivoi = ruta_monivoi
        self.ruta_output = ruta_output
        self.nombre_salida = nombre_salida
        self.callback_status = callback_status
        self.fechas_seleccionadas = fechas_seleccionadas
        self.ruta_vbrl = ruta_vbrl or None

    def _status(self, mensaje):
        """Report progress to the caller (the GUI), if it asked for it."""
        if self.callback_status:
            self.callback_status(mensaje)

    def ejecutar(self):
        """
        Run the full consolidation (steps 1-9 + SG/RG override + 'TD' matrix +
        'MTD' sheet + TD × MTD cross + 'Venta DSD' sheet, and
        — with a VBRL file — the 'VBRL' and 'TD.Diferencias_Referencias' sheets).

        Returns:
            dict: process summary →
                ruta, filas_gallo, filas_monivoi, refs_unicas,
                filas_catalogo, filas_globales, monivoi_sin_datos,
                filas_reclasificadas, cedis_facturas, cedis_notas, cedis_pep,
                total_facturas, total_notas, total_pep, total_importe_gallo,
                td_avisos, total_venta_dsd, venta_avisos,
                centros_mtd_todos, mtd_filas_sin_uuid, mtd_avisos, cruce_avisos,
                vbrl_incluido, vbrl_sin_datos, filas_vbrl, vbrl_avisos,
                filas_td_referencias, total_td_referencias,
                td_referencias_sin_timbre, td_referencias_avisos

        Raises:
            FileNotFoundError: one of the CSVs does not exist.
            ValueError: input the user can fix (Gallo without data, a missing
                column, an amount or date that cannot be interpreted, a
                selected date that is not 'DD.MM.YYYY', a VBRL without the
                columns the cross needs).
        """
        update_status = self._status
        con_vbrl = self.ruta_vbrl is not None

        # Validated BEFORE reading anything: if the GUI sends a malformed date,
        # it is better to find out in the first second than after processing 1M
        # rows.
        fechas_sel = EtapaVentas.parsear_fechas_seleccionadas(self.fechas_seleccionadas)

        # === Step 1: read the files =======================================
        update_status("📄 Leyendo archivo de Gallo...")
        gallo = Comun.leer_csv(self.ruta_gallo)
        if Comun.es_sin_datos(gallo):
            raise ValueError(
                "El archivo de Gallo no tiene datos (descarga 'sin partidas'). "
                "No hay nada que consolidar."
            )

        update_status("📄 Leyendo archivo de Monivoi...")
        monivoi = Comun.leer_csv(self.ruta_monivoi)
        monivoi_sin_datos = Comun.es_sin_datos(monivoi)

        # The VBRL is cheap to validate, so it is read AND checked up front: a
        # wrong layout stops the run now, not after processing the Gallo.
        etapa_vbrl = None
        if con_vbrl:
            update_status("📄 Leyendo archivo de VBRL...")
            vbrl = Comun.leer_csv(self.ruta_vbrl)
            etapa_vbrl = EtapaVBRL(vbrl, Comun.es_sin_datos(vbrl))
            etapa_vbrl.preparar()

        # === Gallo: steps 2-9 + SG/RG override + amount parsing ===========
        # Adds columns K, L, M to `gallo` and returns what the next steps need.
        etapa_gallo = EtapaGallo(gallo, update_status, con_ref_pago=con_vbrl)
        pasos_gallo = etapa_gallo.preparar()

        # === Gallo: 'TD' matrix — sum of amount by cedis × document class ===
        update_status("📊 Construyendo matriz 'TD'...")
        etapa_td = EtapaTD(gallo, pasos_gallo)
        td_resumen = etapa_td.preparar()
        total_importe_gallo = td_resumen["total_importe_gallo"]

        # === Monivoi: step 16 — 'MTD' matrices ============================
        update_status("🧾 Construyendo matrices 'MTD' (SG / RG / PP / TODOS)...")
        etapa_monivoi = EtapaMonivoi(monivoi, monivoi_sin_datos)
        mtd = etapa_monivoi.preparar()
        # Helper column 'Referencia 2' as column A of the Monivoi sheet. It goes
        # AFTER the matrices on purpose: they use their own series, not the
        # sheet.
        etapa_monivoi.insertar_referencia2()

        # === TD × MTD cross: what the formulas cannot show on their own ===
        # (centros stamped in the monitor that have no row in TD).
        cruce = EtapaCruceTDMTD(td_resumen, mtd)
        cruce_avisos = cruce.avisos()

        # === TD.Diferencias_Referencias: Facturas by document vs VBRL ======
        etapa_td_ref = None
        if con_vbrl:
            update_status("🔬 Construyendo 'TD.Diferencias_Referencias'...")
            etapa_td_ref = EtapaTDReferencias(gallo, pasos_gallo, etapa_td, etapa_vbrl)
            etapa_td_ref.preparar()

        # === Gallo: 'Venta DSD' — sum by date ============================
        update_status("📆 Construyendo 'Venta DSD'...")
        etapa_ventas = EtapaVentas(gallo, pasos_gallo, fechas_sel)
        ventas = etapa_ventas.preparar()

        # === Export the consolidated workbook ==============================
        update_status("💾 Guardando Excel consolidado...")
        os.makedirs(self.ruta_output, exist_ok=True)

        nombre_salida = self.nombre_salida
        if not nombre_salida:
            base_gallo = os.path.splitext(os.path.basename(self.ruta_gallo))[0]
            nombre_salida = f"Consolidado_{base_gallo}.xlsx"
        ruta_final = os.path.join(self.ruta_output, nombre_salida)

        # The tab order of the workbook is the order of the writes below.
        with pd.ExcelWriter(ruta_final, engine="openpyxl") as writer:
            gallo.to_excel(writer, sheet_name=EtapaGallo.HOJA_GALLO, index=False)
            monivoi.to_excel(writer, sheet_name=EtapaMonivoi.HOJA_MONIVOI, index=False)

            # --- 'VBRL' sheet (only with a VBRL file) ---
            if etapa_vbrl is not None:
                etapa_vbrl.escribir(writer)

            # --- 'Validación - Doc.' sheet (Gallo catalog: blocks A-B, D-E, G-H) ---
            etapa_gallo.escribir_validacion(
                writer.book.create_sheet(EtapaGallo.HOJA_VALIDACION))

            # --- 'TD' sheet: Facturas, Notas de Crédito and PEP matrices ---
            ws_td = writer.book.create_sheet(EtapaTD.HOJA_TD)
            etapa_td.escribir(ws_td)
            # Cross against 'MTD': live formulas in E:G, P:R and Y:AA.
            cruce.escribir(ws_td)
            # openpyxl saves formulas WITHOUT their result. This forces Excel to
            # calculate all of them when the file is opened, instead of showing
            # empty cells.
            writer.book.calculation.fullCalcOnLoad = True

            # --- 'TD.Diferencias_Referencias' sheet (only with a VBRL file) ---
            if etapa_td_ref is not None:
                etapa_td_ref.escribir(
                    writer.book.create_sheet(EtapaTDReferencias.HOJA_TD_REFERENCIAS))

            # --- 'MTD' sheet: four Monivoi matrices (step 16) ---
            etapa_monivoi.escribir_mtd(writer.book.create_sheet(EtapaMonivoi.HOJA_MTD))

            # --- 'Venta DSD' sheet ---
            for clave, nombre_hoja in EtapaVentas.HOJAS:
                etapa_ventas.escribir(writer.book.create_sheet(nombre_hoja), clave)

        update_status("✅ Consolidación completada")

        resumen_vbrl = etapa_vbrl.resultado if etapa_vbrl is not None else None
        resumen_ref = etapa_td_ref.resultado if etapa_td_ref is not None else None
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
            "venta_avisos":         ventas["avisos"],
            "centros_mtd_todos":    len(mtd["bloques"]["todos"]["filas"]),
            "mtd_filas_sin_uuid":   mtd["filas_sin_uuid"],
            "mtd_avisos":           mtd["avisos"],
            "cruce_avisos":         cruce_avisos,
            # --- VBRL and 'TD.Diferencias_Referencias' (neutral when no VBRL) ---
            "vbrl_incluido":        con_vbrl,
            "vbrl_sin_datos":       bool(resumen_vbrl and resumen_vbrl["sin_datos"]),
            "filas_vbrl":           resumen_vbrl["filas"] if resumen_vbrl else 0,
            "vbrl_avisos":          resumen_vbrl["avisos"] if resumen_vbrl else [],
            "filas_td_referencias": len(resumen_ref["filas"]) if resumen_ref else 0,
            "total_td_referencias": resumen_ref["total"] if resumen_ref else 0.0,
            "td_referencias_sin_timbre": resumen_ref["sin_timbre_n"] if resumen_ref else 0,
            "td_referencias_avisos": resumen_ref["avisos"] if resumen_ref else [],
        }


def consolidar_gallo_monivoi(
    ruta_gallo,
    ruta_monivoi,
    ruta_output,
    nombre_salida=None,
    callback_status=None,
    fechas_seleccionadas=None,
    ruta_vbrl=None,
):
    """
    Entry point used by the controller: builds a `Consolidador` and runs it.
    See `Consolidador.__init__` for the arguments and `Consolidador.ejecutar`
    for the returned summary and the exceptions.
    """
    return Consolidador(
        ruta_gallo, ruta_monivoi, ruta_output,
        nombre_salida=nombre_salida,
        callback_status=callback_status,
        fechas_seleccionadas=fechas_seleccionadas,
        ruta_vbrl=ruta_vbrl,
    ).ejecutar()


# ----------------------------------------------------------------------
# Direct run (quick test without GUI or controller)
# ----------------------------------------------------------------------
if __name__ == "__main__":
    print("=== Prueba de consolidación (standalone) ===")
    ruta_gallo   = input("Ruta del CSV de Gallo   : ").strip().strip('"')
    ruta_monivoi = input("Ruta del CSV de Monivoi : ").strip().strip('"')
    ruta_vbrl    = input("Ruta del CSV de VBRL (Enter = sin VBRL): ").strip().strip('"') or None
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
            ruta_vbrl=ruta_vbrl,
        )
        print("\n🎉 CONSOLIDACIÓN COMPLETADA 🎉")
        for clave, valor in resumen.items():
            print(f"  {clave:<20}: {valor}")
    except Exception as e:
        print(f"❌ Error durante la consolidación: {e}")
        raise
