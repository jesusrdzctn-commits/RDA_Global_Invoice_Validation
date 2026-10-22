"""Validación Factura Global — consolidation: the CSVs downloaded from SAP (Gallo FAGLL03,
Monivoi ZLMXCOM and, optionally, VBRL ZSE16) → ONE Excel workbook. Tabs, in order:
    Gallo, Monivoi, VBRL*, FB**, Validación - Doc., TD, TD.Diferencias_Referencias*, MTD, Venta DSD
    (* only with a VBRL file; ** only with a VBRL and an FB03 file, which also add the 'Notas de
    crédito' and 'PEP Vouchers' tables to 'TD.Diferencias_Referencias')

One class per stage, each with its constants, its computation (preparar → result dict) and
the sheet it writes (escribir): Comun (shared helpers) → EtapaGallo → EtapaTD →
EtapaDocumentosFB03 → EtapaVentas → EtapaMonivoi → EtapaCruceTDMTD → EtapaVBRL →
EtapaTDReferencias → EtapaFB → EtapaTDNotasPEP → Consolidador (the orchestrator); a class only
uses the ones above it.
consolidar_gallo_monivoi() is what the controller calls; consolidar_venta_eiap() when ONLY the
'Solo Venta EIAP' file was downloaded (a workbook with the single tab 'Venta EIAP', see
EtapaVentaEIAP). No SAP and no tkinter here.
"""

import os
import re
import unicodedata
from datetime import datetime
from functools import partial

import pandas as pd
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from utils import CUENTAS_EIAP, CLASES_EIAP


# === Shared helpers ===
class Comun:
    """Building blocks of EVERY stage (read CSVs, find columns, parse amounts, filter legends,
    formula pieces, shared labels). No business rule lives here.
    """

    # Label of the totals row in 'TD' AND 'MTD': the TD × MTD cross looks it up with XLOOKUP.
    ETIQUETA_GRAND_TOTAL = "Grand Total"

    FORMATO_NUMERO = "#,##0.00"

    # Gallo's amount header (also column B of 'Venta DSD').
    ENCABEZADO_IMPORTE = "Importe en moneda local"

    # Filter legends in italic gray, so they never read as data.
    FILTROS_COLOR = "FF595959"

    # What a cross shows when the monitor has nothing to compare against.
    TEXTO_SIN_TIMBRE = "Sin timbre"

    # A US amount with thousands separators ALWAYS groups digits in threes ('1,234.56'); any other
    # comma (the European '1.234,56') is NOT US and must fail, not shrink the amount.
    _RE_IMPORTE_US_AGRUPADO = re.compile(r"\d{1,3}(?:,\d{3})+(?:\.\d+)?")

    # Only letters/'_' can go unquoted in a formula ('MTD!A:A'); R, C, TRUE and FALSE need quotes too.
    _RE_HOJA_SIN_COMILLAS = re.compile(r"[A-Za-z_]+")
    _HOJAS_RESERVADAS = frozenset({"R", "C", "TRUE", "FALSE"})

    @staticmethod
    def leer_csv(ruta):
        """Read a CSV downloaded by the app: everything as text, accents included."""
        if not os.path.exists(ruta):
            raise FileNotFoundError(f"Archivo no encontrado: {ruta}")
        # keep_default_na=False: empty cells stay "" (not NaN), so text comparisons never blow up.
        return pd.read_csv(ruta, dtype=str, encoding="utf-8-sig", keep_default_na=False)

    @staticmethod
    def es_sin_datos(df):
        """True if the CSV is the 'Sin datos' placeholder written by the download."""
        return df.empty or (len(df.columns) == 1 and "Sin datos" in str(df.columns[0]))

    # Cell texts that mean "nothing here" (besides a blank), compared upper-cased: the UUID check
    # of 'MTD' and the document lists of the VBRL / FB03 downloads.
    VALORES_VACIOS = frozenset({
        "", "NULL", "NONE", "NAN", "NA", "N/A", "#N/A", "#N/D", "-", "--", "#",
    })

    @staticmethod
    def documentos_distintos(valores):
        """The distinct document numbers in `valores`, in order of first appearance (the list a download
        pastes into SAP): "nothing here" cells and SAP's all-zero initial values are dropped (they
        would select documents that do not exist) and a '.0' left by the xlsx→csv conversion is removed.
        """
        vistos = {}
        for valor in valores:
            texto = str(valor).strip()
            if re.fullmatch(r"[0-9]+\.0+", texto):
                texto = texto.split(".", 1)[0]
            if texto.upper() in Comun.VALORES_VACIOS or not texto.strip("0"):
                continue
            vistos.setdefault(texto, None)
        return list(vistos)

    @staticmethod
    def normalizar_encabezado(texto):
        """Header normalized to recognize SPELLING VARIANTS of the same field (accents, dots, spaces,
        case): 'Fe.Contabilización' == 'fecha contabilizacion'. Never used to guess another column.
        """
        sin_acentos = unicodedata.normalize("NFKD", str(texto))
        sin_acentos = "".join(c for c in sin_acentos if not unicodedata.combining(c))
        return "".join(c for c in sin_acentos.lower() if c.isalnum())

    @staticmethod
    def col_por_candidatos(df, candidatos, descripcion, origen="Gallo"):
        """The real name of a column, accepting spelling variants of the SAME field. None there →
        ValueError
        listing the real columns (`origen` names the file): stopping beats summing the wrong column.
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

    @staticmethod
    def parse_importe_us(valor, columna=ENCABEZADO_IMPORTE):
        """US-formatted SAP amount → float: '1,234.56', trailing minus '1,234.56-', a leading sign,
        accounting parentheses '(1,234.56)'; blank or a lone '-' → 0.0. Anything else non-numeric, or a
        comma that is not a US thousands separator ('1.234,56'), raises ValueError: a corrupt amount
        must stop the run, not shrink a sum. `columna` only goes into the message.
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

        # ',' is ONLY a thousands separator: validate BEFORE stripping it, so a European amount fails loud.
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

    @staticmethod
    def escribir_filtros(ws, fila_inicial, columna, filtros):
        """Write a matrix's filter legend, one line per row from (fila_inicial, columna) down, italic and
        gray: the exact conditions its rows meet, so any figure can be reproduced by hand. Each caller
        builds it from its OWN filter constants, so the legend cannot drift from what is summed.
        """
        for i, texto in enumerate(filtros):
            celda = ws.cell(row=fila_inicial + i, column=columna, value=texto)
            celda.font = Font(italic=True, color=Comun.FILTROS_COLOR)

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


# === Stage: Gallo (FAGLL03) — steps 2-9 + SG/RG override + 'Validación - Doc.' ===
class EtapaGallo:
    """Everything computed FROM the Gallo: the helper columns K, L, M of 'Gallo' (steps 2-9 + SG/RG
    override, `preparar`) and the 'Validación - Doc.' sheet. It resolves every Gallo column the
    later stages need ONCE, so a missing one stops the run in the first second.
    """

    HOJA_GALLO      = "Gallo"
    HOJA_VALIDACION = "Validación - Doc."

    VALOR_ASIGNACION       = "MX Commercial"             # filter of steps 5, 8.5 and 9
    PREFIJOS_LETRAS_GLOBAL = frozenset({"PP", "RG", "SG"})  # letter-only prefixes that are also 'Global'
    PREFIJOS_RECLASIFICAR  = frozenset({"SG", "RG"})        # step 8.5: SG/RG 'MX Commercial' → Individual

    # Their position in the export is not fixed: resolved by NAME (several spellings accepted).
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
        """`gallo` (the CSV as text) is MUTATED by `preparar` (helper columns). `con_ref_pago`: also
        resolve 'Referencia de pago' (only needed for 'TD.Diferencias_Referencias').
        """
        self.gallo = gallo
        self._update_status = update_status or (lambda _mensaje: None)
        self.con_ref_pago = con_ref_pago
        self.resultado = None

    @staticmethod
    def _col(df, nombre, indice, letra):
        """The REAL name of a column: by name, else by its expected position (column C = index 2)."""
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
        """A column strictly by name (its position in the export is not fixed): ValueError with a clear
        message instead of guessing an index and summing the wrong data.
        """
        if nombre in df.columns:
            return nombre
        raise ValueError(
            f"No encontré la columna '{nombre}' en el archivo de Gallo. "
            f"Columnas disponibles: {list(df.columns)}"
        )

    @staticmethod
    def clasificar_tipo(referencia2):
        """Step 4: 'Global' if it has a digit or is PP/RG/SG; 'Individual' otherwise."""
        ref = str(referencia2).strip()
        if any(caracter.isdigit() for caracter in ref):
            return "Global"
        if ref.upper() in EtapaGallo.PREFIJOS_LETRAS_GLOBAL:
            return "Global"
        return "Individual"

    def preparar(self):
        """Steps 2-9 + the SG/RG override + ONE amount-parsing pass. MUTATES `self.gallo`: appends
        'Referencia 2', 'Referencia 3' and 'Comercializadora' (K, L, M with the standard A-J layout).
        The step order is business logic: step 9 snapshots the globals BEFORE the SG/RG override
        ("step 8.5") reclassifies rows; do not reorder. Returns (also self.resultado) the real column
        names, refs_unicas, tipos, catalogo_docs, globales, filas_reclasificadas and importe_num
        (float, same index). ValueError: a missing column or an unreadable amount.
        """
        gallo = self.gallo
        update_status = self._update_status

        col_centro = self._col(gallo, "Centro",       1, "B")
        col_ref    = self._col(gallo, "Referencia",   2, "C")
        col_asig   = self._col(gallo, "Asignación",   3, "D")
        col_doc    = self._col(gallo, "Nº documento", 4, "E")

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

        # Step 2: 'Referencia 2' = first 2 characters of 'Referencia' (appended last: column K).
        update_status("🔧 Creando 'Referencia 2'...")
        gallo["Referencia 2"] = gallo[col_ref].astype(str).str.strip().str[:2]

        # Step 3: unique 'Referencia 2' values, in order of appearance.
        refs_unicas = list(dict.fromkeys(
            ref for ref in gallo["Referencia 2"] if str(ref).strip() != ""
        ))

        # Step 4: Global / Individual.
        tipos = [self.clasificar_tipo(ref) for ref in refs_unicas]
        mapa_tipo = dict(zip(refs_unicas, tipos))

        # Steps 5-6: catalog of 'MX Commercial' documents.
        update_status("🔎 Filtrando 'MX Commercial'...")
        mascara_mx = gallo[col_asig].astype(str).str.strip() == self.VALOR_ASIGNACION
        catalogo_docs = gallo.loc[mascara_mx, [col_asig, col_doc]].copy()

        # Step 7: 'Referencia 3' (Referencia 2 → Tipo), as values.
        update_status("🔁 Cruzando 'Referencia 3' y 'Comercializadora'...")
        gallo["Referencia 3"] = (
            gallo["Referencia 2"].map(mapa_tipo).fillna("")
        )

        # Step 8: 'Comercializadora' (Nº documento → Asignación), first match like XLOOKUP.
        mapa_comercializadora = {}
        for doc, asig in zip(
            catalogo_docs[col_doc].astype(str).str.strip(),
            catalogo_docs[col_asig].astype(str).str.strip(),
        ):
            mapa_comercializadora.setdefault(doc, asig)

        gallo["Comercializadora"] = (
            gallo[col_doc].astype(str).str.strip().map(mapa_comercializadora).fillna("")
        )

        # Step 9: final filter → global Centro / Nº documento (G-H of 'Validación - Doc.'), frozen
        # BEFORE the SG/RG override on purpose (stakeholder request): it keeps those rows as Global.
        update_status("🧮 Aplicando filtro final (Global + MX Commercial)...")
        mascara_final = (
            (gallo["Comercializadora"] == self.VALOR_ASIGNACION)
            & (gallo["Referencia 3"] == "Global")
        )
        globales = gallo.loc[mascara_final, [col_centro, col_doc]].copy()

        # SG/RG override ("step 8.5", AFTER step 9): MX Commercial rows with Referencia 2 SG/RG are
        # settled individually → 'Individual' in 'Referencia 3' (and so in 'TD'), not in block G-H.
        update_status("🏷️ Reclasificando SG/RG de 'MX Commercial' a 'Individual'...")
        ref2_norm = gallo["Referencia 2"].astype(str).str.strip().str.upper()
        mascara_reclasificar = (
            (gallo["Comercializadora"] == self.VALOR_ASIGNACION)
            & (ref2_norm.isin(self.PREFIJOS_RECLASIFICAR))
            & (gallo["Referencia 3"] == "Global")
        )
        filas_reclasificadas = int(mascara_reclasificar.sum())
        gallo.loc[mascara_reclasificar, "Referencia 3"] = "Individual"

        # ONE amount-parsing pass for every matrix: faster (~1.15M rows for 3 days), and 'TD',
        # 'TD.Diferencias_Referencias' and 'Venta DSD' see exactly the same numbers.
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

    def escribir_validacion(self, ws):
        """'Validación - Doc.': three blocks side by side — A-B references and their type (steps 3-4),
        D-E 'MX Commercial' documents (step 6), G-H Centro and Nº documento of the globals (step 9).
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

        for fila, (ref, tipo) in enumerate(zip(refs_unicas, tipos), start=2):
            ws.cell(row=fila, column=1, value=ref)
            ws.cell(row=fila, column=2, value=tipo)

        for fila, (asig, doc) in enumerate(
            zip(catalogo_docs[col_asig], catalogo_docs[col_doc]), start=2
        ):
            ws.cell(row=fila, column=4, value=asig)
            ws.cell(row=fila, column=5, value=doc)

        for fila, (centro, doc) in enumerate(
            zip(globales[col_centro], globales[col_doc]), start=2
        ):
            ws.cell(row=fila, column=7, value=centro)
            ws.cell(row=fila, column=8, value=doc)

        for letra, ancho in {"A": 14, "B": 12, "D": 16, "E": 16,
                             "G": 10, "H": 16}.items():
            ws.column_dimensions[letra].width = ancho


# === Stage: 'TD' — pivot-style matrices (Facturas, Notas de Crédito, PEP) ===
class EtapaTD:
    """The 'TD' sheet: sum of 'Importe en moneda local' by cedis × document class, in three blocks —
    Facturas (A cedis, B C1, C C5, D Grand Total), Notas de Crédito (L, M C2, N C6, O) and PEP
    vouchers (W, X C2 with Referencia 2 == 'PP'). EtapaCruceTDMTD and EtapaTDReferencias read the
    layout constants defined here, and `mascara_bloque` hands the micro stage the EXACT rows of a
    block, so the macro and micro tables can never drift.
    """

    HOJA_TD = "TD"

    # Rows 1-3 = filter legend (in each block's first column; the analysis labels go in its Grand
    # Total column), 4 = title, 5 = headers, 6+ = data.
    FILA_FILTROS    = 1
    FILA_ENCABEZADO = 5
    FILA_INICIO     = 6
    # 'clases' = (document class, 1-based column of its header/sum); one Grand Total column per block.
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
    # Analysis labels for the stakeholder in the Grand Total column, rows 1-3 (labels only).
    LABELS_ANALISIS = [(1, "Monitor"), (2, "Contabilidad"), (3, "Diferencias")]

    # PEP vouchers: Referencia 3 == 'Global', Referencia 2 == 'PP', Clase de documento == 'C2'
    # (plus a non-blank Centro). It does NOT exclude 'MX Commercial'.
    PEP = {
        "titulo": "PEP VOUCHERS",
        "col_cedis": 23,           # W
        "col_valor": 24,           # X
        "encabezado_valor": "C2",  # class of the filter (goes in X5)
    }
    PEP_REFERENCIA2 = "PP"
    PEP_CLASE = "C2"

    def __init__(self, gallo, pasos_gallo):
        self.gallo = gallo
        self.pasos = pasos_gallo
        self.resultado = None
        # Normalized series, kept for `mascara_bloque`.
        self.centro_norm = None
        self.clase_norm = None
        self.ref2_norm = None
        self._base = None
        self._pep = None

    def preparar(self):
        """Aggregate the three 'TD' blocks. Each derives its OWN cedis list from its OWN filters, so
        the blocks may list different centros (no range clipping):
            Facturas  Referencia 3 == 'Global', Asignación != 'MX Commercial', Clase in {C1, C5}
            Notas     the same base, Clase in {C2, C6}
            PEP       Referencia 3 == 'Global', Referencia 2 == 'PP', Clase == 'C2' (MX Commercial too)
        All three need a non-blank Centro: rows without one are counted and reported in 'avisos'
        (never dropped silently), so a short Grand Total can always be explained. The amount comes
        already parsed (`pasos_gallo["importe_num"]`). Returns per block 'cedis' and 'suma', plus
        the control totals and 'avisos'.
        """
        gallo = self.gallo
        pasos = self.pasos
        col_centro  = pasos["col_centro"]
        col_asig    = pasos["col_asig"]
        col_clase   = pasos["col_clase"]
        importe_num = pasos["importe_num"]
        valor_asignacion = EtapaGallo.VALOR_ASIGNACION

        centro_norm = gallo[col_centro].astype(str).str.strip()
        clase_norm  = gallo[col_clase].astype(str).str.strip().str.upper()
        ref2_norm   = gallo["Referencia 2"].astype(str).str.strip().str.upper()
        asig_norm   = gallo[col_asig].astype(str).str.strip()
        es_global   = gallo["Referencia 3"] == "Global"
        con_centro  = centro_norm != ""

        base = es_global & (asig_norm != valor_asignacion) & con_centro
        # The same filters WITHOUT a centro: reported in 'avisos' instead of vanishing.
        base_sin_centro = es_global & (asig_norm != valor_asignacion) & ~con_centro

        self.centro_norm = centro_norm
        self.clase_norm = clase_norm
        self.ref2_norm = ref2_norm
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
            # Round to cents: no float noise (-265.44000000000005) in the cells or the totals.
            return {"cedis": cedis,
                    "suma": {k: round(float(v), 2) for k, v in serie.items()}}

        facturas = _bloque(base, self.FACTURAS["clases"])
        notas    = _bloque(base, self.NOTAS["clases"])

        filtro_pep = (
            es_global
            & (ref2_norm == self.PEP_REFERENCIA2)
            & (clase_norm == self.PEP_CLASE)
        )
        mascara_pep = filtro_pep & con_centro
        self._pep = mascara_pep
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
            # The column names REALLY found in this export, for the filter legends.
            "columnas": {
                "centro":      col_centro,
                "asignacion":  col_asig,
                "clase":       col_clase,
            },
        }
        return self.resultado

    def mascara_bloque(self, clases):
        """Mask of the Gallo rows of a two-class block — the SAME rows `preparar` summed (Global, not
        MX Commercial, non-blank Centro, Clase in `clases`). Call it after `preparar`;
        EtapaTDReferencias builds the micro table from it.
        """
        if self._base is None:
            raise RuntimeError("EtapaTD.mascara_bloque requiere llamar antes a preparar().")
        return self._base & self.clase_norm.isin({clase for clase, _ in clases})

    def mascara_pep(self):
        """Mask of the Gallo rows of the PEP block — the SAME rows `preparar` summed (Global, Referencia
        2 == 'PP', Clase == 'C2', non-blank Centro; 'MX Commercial' included). Call it after `preparar`.
        """
        if self._pep is None:
            raise RuntimeError("EtapaTD.mascara_pep requiere llamar antes a preparar().")
        return self._pep

    @staticmethod
    def leyenda_filtros(bloque, columnas):
        """Filter legend of a two-class block, built FROM its own constants, so it cannot drift from what
        is summed. 'Referencia 3' is not resolved: EtapaGallo creates it with exactly that name.
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
        """Filter legend of the PEP block: it filters by 'Referencia 2' and does NOT exclude 'MX
        Commercial', spelled out so nobody assumes the exclusion is there.
        """
        return [
            f"Filtros — hoja '{EtapaGallo.HOJA_GALLO}':",
            f"Referencia 3 = 'Global'; {columnas['centro']} no vacío",
            f"Referencia 2 = '{EtapaTD.PEP_REFERENCIA2}'; "
            f"{columnas['clase']} = '{EtapaTD.PEP_CLASE}' "
            f"(sí incluye '{EtapaGallo.VALOR_ASIGNACION}')",
        ]

    def escribir(self, ws_td):
        """Write the three blocks on 'TD': rows 1-3 legend + analysis labels, 4 title, 5 headers, 6+
        data, then a 'Grand Total' row (each block computes its own position: row counts differ).
        Values, not formulas, with a thousands/2-decimals format. The TD × MTD cross (E:G, P:R, Y:AA)
        is written afterwards by EtapaCruceTDMTD, which looks the bottom label up in 'MTD': it MUST
        stay Comun.ETIQUETA_GRAND_TOTAL.
        """
        td_resumen = self.resultado
        fila_inicio = self.FILA_INICIO
        fila_encabezado = self.FILA_ENCABEZADO

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

            Comun.escribir_filtros(ws_td, self.FILA_FILTROS, col_cedis,
                                   self.leyenda_filtros(bloque, columnas))
            for fila, texto in self.LABELS_ANALISIS:
                ws_td.cell(row=fila, column=col_gt, value=texto).font = Font(bold=True)

            ws_td.cell(row=4, column=col_cedis, value=bloque["titulo"]).font = Font(bold=True)

            ws_td.cell(row=fila_encabezado, column=col_cedis, value="Centro").font = Font(bold=True)
            for clase, col_idx in clases:
                ws_td.cell(row=fila_encabezado, column=col_idx, value=clase).font = Font(bold=True)
            ws_td.cell(row=fila_encabezado, column=col_gt, value="Grand Total").font = Font(bold=True)

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
                ws_td.cell(row=fila, column=col_gt,
                           value=round(suma_fila, 2)).number_format = Comun.FORMATO_NUMERO

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

        # --- PEP vouchers (its own cedis, one value column) ---
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

        anchos = {"A": 12, "B": 16, "C": 16, "D": 16,
                  "L": 12, "M": 16, "N": 16, "O": 16,
                  "W": 12, "X": 16}
        for letra, ancho in anchos.items():
            ws_td.column_dimensions[letra].width = ancho


# === Stage: the document list of the FB03 download (from the Gallo) ===
class EtapaDocumentosFB03:
    """The documents the FB03 download pastes into SAP, from the GALLO of the same run: 'Nº documento'
    of the rows with Clase de documento in {C2, C6} (the 'Notas de Crédito' classes of 'TD') AND a
    Global 'Referencia' (step 4: a digit in its first 2 characters, or PP / RG / SG). It reuses
    EtapaGallo's pieces and EtapaTD.NOTAS, so it cannot drift from the consolidation.
    Deliberately NOT applied (only those two filters were asked for): the 'MX Commercial'
    exclusion, the non-blank Centro and the SG/RG override — an extra document in FB03 does no
    harm, a missing one would.
    """

    CLASES = frozenset(clase for clase, _ in EtapaTD.NOTAS["clases"])     # {'C2', 'C6'}

    # How the filter is described to the user, built from the same constants it uses.
    _PREFIJOS = tuple(sorted(EtapaGallo.PREFIJOS_LETRAS_GLOBAL))            # ('PP', 'RG', 'SG')
    DESCRIPCION_FILTRO = (
        f"Clase de documento {'/'.join(sorted(CLASES))} y Referencia Global "
        f"(empieza con {', '.join(_PREFIJOS[:-1])} o {_PREFIJOS[-1]}, "
        "o trae un número en sus 2 primeros caracteres)"
    )

    @staticmethod
    def columnas(gallo):
        """(col_referencia, col_documento, col_clase) as EtapaGallo.preparar resolves them (by name, else
        position C / E; 'Clase de documento' strictly by name). ValueError if one is missing.
        """
        return (
            EtapaGallo._col(gallo, "Referencia", 2, "C"),
            EtapaGallo._col(gallo, "Nº documento", 4, "E"),
            EtapaGallo._col_por_nombre(gallo, "Clase de documento"),
        )

    @staticmethod
    def documentos(gallo, columnas=None):
        """The FB03 list of a Gallo DataFrame (not modified) → (documentos, filas_leidas, nombre_columna);
        empty for the 'Sin datos' placeholder. `columnas` may come already resolved (see `leer`).
        """
        if Comun.es_sin_datos(gallo):
            return [], 0, "Nº documento"
        col_ref, col_doc, col_clase = columnas or EtapaDocumentosFB03.columnas(gallo)

        clase = gallo[col_clase].astype(str).str.strip().str.upper()
        # Steps 2-4 of EtapaGallo; a blank 'Referencia 2' is never Global.
        referencia2 = gallo[col_ref].astype(str).str.strip().str[:2]
        tipos = {
            ref: EtapaGallo.clasificar_tipo(ref)
            for ref in referencia2.unique() if str(ref).strip() != ""
        }
        es_global = referencia2.map(tipos).fillna("") == "Global"

        mascara = clase.isin(EtapaDocumentosFB03.CLASES) & es_global
        return Comun.documentos_distintos(gallo.loc[mascara, col_doc]), len(gallo), col_doc

    @staticmethod
    def leer(ruta_gallo):
        """Read the Gallo CSV and return `documentos` of it, loading only the three columns the filter
        needs (resolved on the FULL header, so the positional fallbacks point at the right columns).
        FileNotFoundError if the file is missing, ValueError if a column is.
        """
        if not os.path.exists(ruta_gallo):
            raise FileNotFoundError(f"Archivo no encontrado: {ruta_gallo}")
        encabezado = pd.read_csv(ruta_gallo, dtype=str, encoding="utf-8-sig",
                                 keep_default_na=False, nrows=0)
        nombres = list(encabezado.columns)
        if len(nombres) == 1 and "Sin datos" in str(nombres[0]):
            return [], 0, "Nº documento"
        columnas = EtapaDocumentosFB03.columnas(encabezado)
        posiciones = sorted({nombres.index(c) for c in columnas})
        gallo = pd.read_csv(ruta_gallo, dtype=str, encoding="utf-8-sig",
                            keep_default_na=False, usecols=posiciones)
        # Names from the FULL header, by position (with repeated headers pandas could rename the subset).
        gallo.columns = [nombres[p] for p in posiciones]
        return EtapaDocumentosFB03.documentos(gallo, columnas)


# === Stage: 'Venta DSD' — amount by date ===
class EtapaVentas:
    """The 'Venta DSD' sheet: date (A) × summed amount (B) for Clase de documento C1/C2/C5/C6/X0 or
    blank ("" = Excel's "(blanks)") AND accounts 4000010/4000005. It also validates the dates picked
    in the GUI (only this sheet uses them). 'Venta EIAP' left the consolidation: it is downloaded on
    its own with the 'Solo Venta EIAP' button (see EtapaVentaEIAP).
    """

    HOJA_DSD  = "Venta DSD"
    HOJAS = (("dsd", HOJA_DSD),)

    CLASES_DSD  = frozenset({"C1", "C2", "C5", "C6", "X0", ""})
    # Accounts compared NORMALIZED: '4000010', '0004000010' and '4000010.0' are the same.
    CUENTAS_DSD  = frozenset({"4000010", "4000005"})
    # The matrices `preparar` builds: (key, tab, classes, accounts).
    BLOQUES = (("dsd", HOJA_DSD, CLASES_DSD, CUENTAS_DSD),)
    # Where the rows come from, for the warnings and the filter legend.
    ORIGEN = "Gallo"
    LEYENDA_ORIGEN = "hoja 'Gallo'"

    ENCABEZADO_FECHA   = "Fecha"
    ETIQUETA_SIN_FECHA = "Sin fecha"
    ETIQUETA_TOTAL     = Comun.ETIQUETA_GRAND_TOTAL
    FORMATO_FECHA_EXCEL = "DD.MM.YYYY"

    # Rows 1-4 = legend (two filters: class AND account), 5 = headers, 6+ = data (as in 'TD', 'MTD').
    FILA_FILTROS    = 1
    FILA_ENCABEZADO = 5
    FILA_INICIO     = 6
    # How the empty class reads in the legend ('' on screen would look like a typo).
    ETIQUETA_EN_BLANCO = "en blanco"

    # 4-digit-year formats first, 2-digit ones last; real ambiguity (01/09 = Sep 1 or Jan 9) is
    # resolved with the dates picked in the GUI, or fails — never guessed.
    FORMATOS_FECHA_4D = (
        "%d.%m.%Y", "%m.%d.%Y",
        "%d/%m/%Y", "%m/%d/%Y",
        "%Y-%m-%d", "%Y/%m/%d",
        "%d-%m-%Y", "%m-%d-%Y",
    )
    FORMATOS_FECHA_2D = ("%d.%m.%y", "%m.%d.%y", "%d/%m/%y", "%m/%d/%y")
    ANIO_MIN = 1990
    ANIO_MAX = 2100

    # Keeps the DATE part of 'dd/mm/yyyy hh:mm' (the day/month order is detectar_formato's job).
    _RE_SOLO_FECHA = re.compile(r"^(\d{1,4}[./\-]\d{1,2}[./\-]\d{1,4})(?:[ T].*)?$")

    def __init__(self, gallo, pasos_gallo, fechas_sel=()):
        self.gallo = gallo
        self.pasos = pasos_gallo
        self.fechas_sel = fechas_sel
        self.resultado = None

    @staticmethod
    def normalizar_cuenta(valor):
        """G/L account for comparing: ' 0004000010 ' → '4000010', '4000010.0' → '4000010'. Only leading
        zeros and an all-zero decimal part go; a non-numeric value stays (upper-cased, no spaces).
        """
        texto = str(valor).strip().upper().replace(" ", "")
        if not texto:
            return ""
        if re.fullmatch(r"\d+\.0*", texto):
            texto = texto.split(".", 1)[0]
        if texto.isdigit():
            return texto.lstrip("0") or "0"
        return texto

    @staticmethod
    def solo_fecha(valor):
        """Cut the time off a date ('01/09/2026 14:35' → '01/09/2026', '2026-09-01T00:00' → '2026-09-01');
        anything else is returned stripped, so the format detector shows the REAL value.
        """
        texto = str(valor).strip()
        coincidencia = EtapaVentas._RE_SOLO_FECHA.match(texto)
        return coincidencia.group(1) if coincidencia else texto

    @staticmethod
    def parsear_fechas_seleccionadas(fechas):
        """The GUI dates ('DD.MM.YYYY') as date objects, without duplicates and in order; [] if none."""
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
        """The dates if ALL values parse with `formato` into a reasonable year (strptime takes '26' as the
        year 26 AD for %Y, which would let a wrong format pass), else None.
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
        """Format of the Gallo date column: 4-digit-year formats first, then 2-digit ones. Several
        formats that parse everything: same results → harmless, the first one; different results →
        the one that puts more dates inside the GUI selection; still a tie → ValueError (a misread
        01/09/2026 would move the amount to another day). None if there is nothing to parse.
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

    def preparar(self):
        """'Venta DSD' → {'dsd': {filas, sin_fecha_n, sin_fecha_importe, total, fuera_rango}, 'avisos'}."""
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

        distintas = sorted({valor for valor in fecha_raw.unique() if valor})
        formato = self.detectar_formato(distintas, fechas_sel)
        mapa_fechas = (
            {valor: datetime.strptime(valor, formato).date() for valor in distintas}
            if formato else {}
        )
        # Blank dates → None: reported instead of vanishing.
        fecha_norm = fecha_raw.map(lambda valor: mapa_fechas.get(valor))

        cuentas_presentes = set(cuenta_norm.unique())
        seleccionadas = set(fechas_sel)
        avisos = []
        bloques = {}

        for clave, etiqueta, clases, cuentas in self.BLOQUES:
            cuentas_buscadas = {self.normalizar_cuenta(c) for c in cuentas}
            mascara = clase_norm.isin(clases) & cuenta_norm.isin(cuentas_buscadas)

            # A requested account missing from the file: no postings that day... or a layout change → warn.
            faltantes = sorted(c for c in cuentas_buscadas if c not in cuentas_presentes)
            if faltantes:
                avisos.append(
                    f"{etiqueta}: la(s) cuenta(s) {', '.join(faltantes)} no "
                    f"aparece(n) en el archivo de {self.ORIGEN}."
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

    @staticmethod
    def leyenda_filtros(clases, cuentas, columnas, origen="hoja 'Gallo'"):
        """Filter legend of 'Venta DSD', from its own filter sets and the real column names. It spells out
        what makes a manual check disagree: the EMPTY class counts ("(blanks)" in Excel), and accounts
        compare without leading zeros or a '.0' ('0004000010' and '4000010.0' are '4000010').
        """
        lista_clases = " o ".join(
            [f"'{clase}'" for clase in sorted(clase for clase in clases if clase)]
            + ([EtapaVentas.ETIQUETA_EN_BLANCO] if "" in clases else [])
        )
        lista_cuentas = " o ".join(f"'{cuenta}'" for cuenta in sorted(cuentas))
        return [
            f"Filtros — {origen}:",
            f"{columnas['clase']} = {lista_clases}",
            f"{columnas['cuenta']} = {lista_cuentas}",
            f"{columnas['fecha']} vacía → renglón '{EtapaVentas.ETIQUETA_SIN_FECHA}'",
        ]

    def escribir(self, ws, clave):
        """Write 'Venta DSD': rows 1-4 legend, 5 headers, 6+ one row per day (real dates, DD.MM.YYYY),
        then 'Sin fecha' (if any) and a bold 'Grand Total'. Amounts are values with 2 decimals.
        """
        bloque = self.resultado[clave]
        Comun.escribir_filtros(ws, self.FILA_FILTROS, 1,
                               self.leyenda_filtros(bloque["clases"], bloque["cuentas"],
                                                    bloque["columnas"], self.LEYENDA_ORIGEN))

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


class EtapaVentaEIAP(EtapaVentas):
    """'Venta EIAP' of the EIAP-only consolidation: the matrix of 'Venta DSD' (date × amount, Grand
    Total) over the EIAP download, with ITS filters (classes RV/DA, the EIAP accounts). The sheet is
    the whole workbook: no Gallo / Monivoi involved.
    """

    HOJA_EIAP = "Venta EIAP"
    BLOQUES = (("eiap", HOJA_EIAP, frozenset(CLASES_EIAP), frozenset(CUENTAS_EIAP)),)
    ORIGEN = "Venta EIAP"
    LEYENDA_ORIGEN = "archivo de Venta EIAP"

    @classmethod
    def desde_csv(cls, eiap, fechas_sel=()):
        """The stage for the EIAP CSV (as text): resolves its columns by name and parses the amounts.
        ValueError: a missing column or an unreadable amount.
        """
        def col(candidatos, descripcion):
            return Comun.col_por_candidatos(eiap, candidatos, descripcion, cls.ORIGEN)

        col_importe = col((Comun.ENCABEZADO_IMPORTE,), "importe en moneda local")
        pasos = {
            "col_clase":   col(("Clase de documento",), "clase de documento"),
            "col_cuenta":  col(EtapaGallo.COL_CUENTA_CANDIDATAS, "cuenta"),
            "col_fecha":   col(EtapaGallo.COL_FECHA_CANDIDATAS, "fecha de contabilización"),
            "importe_num": eiap[col_importe].map(partial(Comun.parse_importe_us, columna=col_importe)),
        }
        return cls(eiap, pasos, fechas_sel)

    def preparar(self):
        """Like 'Venta DSD', plus a warning for the rows of other classes (SAP's class filter failed or
        was changed): they are left out of the sum, never added to it.
        """
        bloques = super().preparar()
        clases = self.BLOQUES[0][2]
        clase_norm = self.gallo[self.pasos["col_clase"]].astype(str).str.strip().str.upper()
        ajenas = int((~clase_norm.isin(clases)).sum())
        if ajenas:
            bloques["avisos"].append(
                f"{self.HOJA_EIAP}: {ajenas:,} fila(s) con clase de documento distinta de "
                f"{' / '.join(sorted(clases))}; no se sumaron."
            )
        return bloques


# === Stage: Monivoi (ZLMXCOM_TRN_MONINVOI) — 'MTD' (step 16) ===
class EtapaMonivoi:
    """Everything computed FROM the Monivoi: the 'MTD' sheet (step 16: four Centro × 'Valor neto'
    matrices, SG, RG, PP and TODOS), the helper column 'Referencia 2' (column A of 'Monivoi') and
    the invoice list of the VBRL download ('Factura'). The cross against 'TD' is EtapaCruceTDMTD.
    """

    HOJA_MONIVOI = "Monivoi"
    HOJA_MTD     = "MTD"

    # 'MTD' layout: rows 1-3 = filter legend, 4 = title, 5 = headers, 6+ = data. The TD × MTD cross
    # looks centros up over WHOLE columns (MTD!A:A), so these rows do not affect it.
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
    # The prefixes of 'TODOS': also the only rows whose amount gets parsed.
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

    # Helper column added as column A of 'Monivoi', to audit the matrices by filtering by hand.
    COL_REFERENCIA2 = "Referencia 2"

    COL_REFERENCIA_CANDIDATAS = ("Referencia", "Referencia de factura")
    COL_UUID_CANDIDATAS       = ("UUID", "UUID CFDI", "Folio fiscal", "UUID Fiscal")
    COL_CENTRO_CANDIDATAS     = ("Centro", "Centro de distribución",
                                 "Puesto de expedición")
    COL_VALOR_CANDIDATAS      = ("Valor neto", "Valor neto factura", "Importe neto")
    # The invoice pasted into VBRL: only this exact header ('Factura de reemplazo' and 'Referencia
    # de factura' are OTHER fields).
    COL_FACTURA_CANDIDATAS = ("Factura",)

    # Which Monivoi column feeds which download: only VBRL reads it ('Factura'); the FB03 documents
    # come from the Gallo (EtapaDocumentosFB03).
    COLUMNAS_DESCARGA = {
        "factura":   (COL_FACTURA_CANDIDATAS,       "Factura"),
    }

    # Empty-UUID sentinels; the UUID SHAPE is not validated on purpose (one without dashes counts).
    UUID_VACIOS = Comun.VALORES_VACIOS

    # 'TODOS == SG + RG + PP' tolerance: a cent covers the rounding of values already rounded.
    TOLERANCIA_CUADRE = 0.01

    POR_CLAVE = {bloque["clave"]: bloque for bloque in BLOQUES}

    def __init__(self, monivoi, sin_datos):
        self.monivoi = monivoi
        self.sin_datos = sin_datos
        self.resultado = None

    @staticmethod
    def nombre_columna_libre(df, base):
        """`base`, or 'base (2)', 'base (3)'... if the DataFrame already has it: the helper column must
        never overwrite a same-named column of the export.
        """
        if base not in df.columns:
            return base
        sufijo = 2
        while f"{base} ({sufijo})" in df.columns:
            sufijo += 1
        return f"{base} ({sufijo})"

    @staticmethod
    def uuid_vacio(valor):
        """Step 16: True if the UUID counts as EMPTY (blank, whitespace or a sentinel like 'NULL', 'N/A',
        '-'). Its shape is not validated: a real UUID written differently still counts.
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

    def preparar(self):
        """Step 16 — the FOUR 'MTD' matrices, all over the MONIVOI, grouped by the first 2 characters of
        'Referencia' (SG → A/B, RG → D/E, PP → K/L, the three → Q/R 'TODOS'), only rows with a
        non-empty UUID (the others were never invoiced). Each matrix has its OWN centros; rows without
        a centro go to a 'Sin centro' row, so every Grand Total reconciles. 'Valor neto' is parsed only
        for the rows of some matrix. Returns bloques, referencia2, filas_sin_uuid and avisos.
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
                # Nothing was read: the legend uses the canonical names (there is nothing to audit anyway).
                "columnas": {
                    "uuid":   self.COL_UUID_CANDIDATAS[0],
                    "centro": self.COL_CENTRO_CANDIDATAS[0],
                },
            }
            return self.resultado

        col_ref    = Comun.col_por_candidatos(monivoi, self.COL_REFERENCIA_CANDIDATAS,
                                              "referencia", origen="Monivoi")
        col_uuid   = Comun.col_por_candidatos(monivoi, self.COL_UUID_CANDIDATAS,
                                              "UUID", origen="Monivoi")
        col_centro = Comun.col_por_candidatos(monivoi, self.COL_CENTRO_CANDIDATAS,
                                              "centro de distribución", origen="Monivoi")
        col_valor  = Comun.col_por_candidatos(monivoi, self.COL_VALOR_CANDIDATAS,
                                              "valor neto", origen="Monivoi")

        avisos = []

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

        mascara_union = con_uuid & ref2_norm.isin(self.PREFIJOS_TODOS)
        # astype(float) is NOT cosmetic: with no row, pandas 3 sums an empty text Series to '', not 0.0.
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
                # groupby sorts the keys: centros come out alphabetically, as in 'TD'.
                serie = con_centro.groupby("_centro")["_valor"].sum()
                filas = [(centro, round(float(valor), 2)) for centro, valor in serie.items()]

            # The total adds the ALREADY ROUNDED values: the column adds up to what is shown.
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

        # Reconciliation: TODOS must equal SG + RG + PP; warn instead of delivering silently.
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
            # The column names REALLY found in this export, for the filter legends.
            "columnas": {
                "uuid":   col_uuid,
                "centro": col_centro,
            },
        }
        return self.resultado

    def insertar_referencia2(self):
        """Insert 'Referencia 2' as column A of the Monivoi sheet (under a free name, with a warning, if
        the export already has one) and record its final name in resultado["nombre_ref2"] for the
        legends. Call it AFTER `preparar`; it does nothing without data.
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
            # The legends must cite the column the stakeholder will REALLY find on 'Monivoi'.
            mtd["nombre_ref2"] = nombre_ref2

    @staticmethod
    def leyenda_filtros(bloque, nombre_ref2, columnas):
        """Filter legend of an 'MTD' block, from its own prefixes: source sheet, 'Referencia 2' (by its
        real name `nombre_ref2`), non-empty UUID and what happens to the rows without a centro. Lines
        are short on purpose: the first block only has columns B and C to overflow into.
        """
        prefijos = " o ".join(f"'{p}'" for p in sorted(bloque["prefijos"]))
        return [
            f"Filtros — hoja '{EtapaMonivoi.HOJA_MONIVOI}':",
            f"{nombre_ref2} = {prefijos}; {columnas['uuid']} con dato",
            f"{columnas['centro']} vacío → renglón '{EtapaMonivoi.ETIQUETA_SIN_CENTRO}'",
        ]

    def escribir_mtd(self, ws):
        """Write the four step-16 matrices on 'MTD', each with its filter legend in rows 1-3. Call it
        AFTER `insertar_referencia2`, so the legends cite the helper column by its final name.
        """
        mtd = self.resultado
        # Final name of the helper column (the default only applies without data: no helper column then).
        nombre_ref2 = mtd.get("nombre_ref2", self.COL_REFERENCIA2)
        columnas = mtd["columnas"]

        for bloque in self.BLOQUES:
            datos      = mtd["bloques"][bloque["clave"]]
            col_centro = bloque["col_centro"]
            col_valor  = bloque["col_valor"]

            Comun.escribir_filtros(ws, self.FILA_FILTROS, col_centro,
                                   self.leyenda_filtros(bloque, nombre_ref2, columnas))

            ws.cell(row=self.FILA_TITULO, column=col_centro,
                    value=bloque["titulo"]).font = Font(bold=True)
            ws.cell(row=self.FILA_ENCABEZADO, column=col_centro,
                    value=self.ENCABEZADO_CENTRO).font = Font(bold=True)
            ws.cell(row=self.FILA_ENCABEZADO, column=col_valor,
                    value=self.ENCABEZADO_VALOR).font = Font(bold=True)

            fila = self.FILA_INICIO
            for centro, valor in datos["filas"]:
                ws.cell(row=fila, column=col_centro, value=centro)
                ws.cell(row=fila, column=col_valor,
                        value=valor).number_format = Comun.FORMATO_NUMERO
                fila += 1

            # Rows without a centro: shown apart, so the Grand Total reconciles with the filter.
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

            ws.column_dimensions[get_column_letter(col_centro)].width = 12
            ws.column_dimensions[get_column_letter(col_valor)].width = 20

    @staticmethod
    def documentos_unicos(monivoi, columna):
        """ALL the distinct values of ONE Monivoi column, in order of first appearance: the list a download
        pastes into SAP ('factura' → 'Factura', VBRL). NOT filtered by UUID or 'Referencia' (every value
        was asked for); only "nothing here" cells and the '.0' of the conversion go. `columna` is
        REQUIRED on purpose: a default could paste the wrong column. Returns (valores, filas_leidas,
        nombre_columna); ValueError for an unknown key or a Monivoi without that column.
        """
        if columna not in EtapaMonivoi.COLUMNAS_DESCARGA:
            raise ValueError(
                f"Columna de Monivoi desconocida: {columna!r} "
                f"(opciones: {', '.join(EtapaMonivoi.COLUMNAS_DESCARGA)})."
            )
        candidatas, descripcion = EtapaMonivoi.COLUMNAS_DESCARGA[columna]
        if Comun.es_sin_datos(monivoi):
            return [], 0, candidatas[0]
        col_documento = Comun.col_por_candidatos(
            monivoi, candidatas, descripcion, origen="Monivoi")
        return Comun.documentos_distintos(monivoi[col_documento]), len(monivoi), col_documento

    @staticmethod
    def leer_documentos(ruta_monivoi, columna):
        """Read the Monivoi CSV and return `documentos_unicos` of it for `columna`."""
        return EtapaMonivoi.documentos_unicos(Comun.leer_csv(ruta_monivoi), columna)


# === Stage: cross 'TD' × 'MTD' — live formulas on the 'TD' sheet ===
class EtapaCruceTDMTD:
    """The TD × MTD cross: LIVE Excel formulas on 'TD' (Factura Monitor / Diferencia / Observación)
    and the warnings they cannot show. It reads the layouts of EtapaTD and EtapaMonivoi, so a layout
    change keeps working as long as it is made in those constants.
    """

    # Three columns right of each TD block, with LIVE formulas (auditable, draggable):
    #     Facturas (A..D) ↔ MTD SG (A/B) → E, F, G;  Notas (L..O) ↔ RG (D/E) → P, Q, R;
    #     PEP (W..X) ↔ PP (K/L) → Y, Z, AA.
    # Also in the 'Grand Total' row: there XLOOKUP finds MTD's Grand Total (same label).
    ENCABEZADO_MONITOR     = "Factura Monitor"
    ENCABEZADO_DIFERENCIA  = "Diferencia"
    ENCABEZADO_OBSERVACION = "Observación"
    TEXTO_SIN_TIMBRE       = Comun.TEXTO_SIN_TIMBRE
    TEXTO_SIN_DIFERENCIA   = "Sin diferencia"

    # 'Observación' tolerance: up to one cent counts as 'Sin diferencia'.
    TOLERANCIA = 0.01

    # 'revisar_si': the ONLY direction that leaves Observación blank for review (accounting exceeds
    # the monitor); the other one shows 'Sin diferencia' on purpose (business rule):
    #     '<' → ROUND(F6,2)<-0.01        '>' → ROUND(Q6,2)>0.01
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
        self.td_resumen = td_resumen
        self.mtd = mtd

    @staticmethod
    def formulas(cruce, fila):
        """The three formulas of ONE row (Factura Monitor, Diferencia, Observación); Facturas, row 6:
            E6 =IFERROR(XLOOKUP(A6,MTD!A:A,MTD!B:B),"Sin timbre")
            F6 =IFERROR(D6+E6,"Sin timbre")
            G6 =IF(F6="Sin timbre","Sin timbre",IF(ROUND(F6,2)<-0.01,"","Sin diferencia"))
        XLOOKUP is written `_xlfn.XLOOKUP` (newer than Excel 2007: without it Excel shows #NAME?).
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
        """Write the cross on 'TD': three columns per block (see CRUCES), the header in row 5 and a
        formula in every centro row PLUS the 'Grand Total' row (none when the block has no centros).
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

    def avisos(self):
        """Warnings the formulas cannot show: centros the monitor stamped that have no TD row (they only
        weigh in the Grand Total), and centros that differ only in case (XLOOKUP ignores case and keeps
        the first match). Returns list[str].
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


# === Stage: VBRL (ZSE16) — what the monitor invoiced, per invoice ===
class EtapaVBRL:
    """The 'VBRL' sheet (ZSE16 for the Monivoi invoices), source of 'Factura Monitor' in
    'TD.Diferencias_Referencias' and in its Notas de crédito / PEP Vouchers tables. The lookups
    read VBRL!D:D (invoice) and VBRL!E:E (amount) BY POSITION, like the stakeholder's XLOOKUP.
    SAP delivers the invoice as TEXT with its leading 'C' ('C270043678'), exactly as Gallo's
    'Referencia de pago' and FB03's 'Clave de referencia' carry it, so column D is written as
    trimmed text and the formulas match text against text, with no conversion at all. Only the
    amount (E) becomes a real number.
    """

    HOJA_VBRL = "VBRL"

    # 1-based positions read by the lookups (D, E); the download checks the same ones
    # (Descargas_SAP.VBRL.MIN_COLUMNAS / COL_FACTURA, 0-based there): change both together.
    COL_FACTURA = 4
    COL_IMPORTE = 5

    # What an invoice looks like in SAP's export: an optional 'C' and digits ('C270043678'; a few
    # rows per day come as the bare 10-digit billing document). Anything else is reported.
    _RE_FACTURA = re.compile(r"C?[0-9]+", re.IGNORECASE)

    # Two amounts for the same invoice closer than this are the same amount.
    TOLERANCIA_DUPLICADOS = 0.005

    def __init__(self, vbrl, sin_datos):
        self.vbrl = vbrl
        self.sin_datos = sin_datos
        self.resultado = None

    @staticmethod
    def limpiar_factura(valor):
        """A VBRL invoice cell as written in VBRL!D:D: the same text, trimmed ('C270043678'); '' → None."""
        texto = str(valor).strip()
        return texto or None

    @staticmethod
    def clave_busqueda(valor):
        """The key XLOOKUP compares, as the formulas see it: trimmed and upper-cased (XLOOKUP ignores
        case), so the Python-side controls (`mapa`, the warnings) agree with Excel. '' → None.
        """
        texto = str(valor).strip().upper()
        return texto or None

    def preparar(self):
        """Invoice (text) and amount (number) columns, indexed the way XLOOKUP reads them (first match
        wins). Returns df, sin_datos, filas, mapa ({invoice upper-cased: amount}), avisos and
        encabezados (headers of D and E). ValueError: fewer than 5 columns (wrong layout in the
        download) or an unreadable amount.
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

        facturas = [self.limpiar_factura(valor)
                    for valor in vbrl.iloc[:, self.COL_FACTURA - 1]]
        etiqueta_importe = f"{col_importe} (VBRL, columna {letra_importe})"
        importes = [
            None if str(valor).strip() == ""
            else Comun.parse_importe_us(valor, columna=etiqueta_importe)
            for valor in vbrl.iloc[:, self.COL_IMPORTE - 1]
        ]

        # First row of each invoice wins, as in XLOOKUP; repeats are counted for the warnings.
        mapa = {}
        repeticiones = {}
        importes_distintos = set()
        en_blanco = 0
        no_numericas = []
        for factura, importe in zip(facturas, importes):
            if factura is None:
                en_blanco += 1
                continue
            if not self._RE_FACTURA.fullmatch(factura):
                no_numericas.append(factura)
            # The key is what XLOOKUP compares (case-insensitive); a row that does not look like an
            # invoice is still findable by its exact text, so it stays in the index.
            clave = self.clave_busqueda(factura)
            valor = 0.0 if importe is None else importe
            if clave in mapa:
                repeticiones[clave] += 1
                if abs(mapa[clave] - valor) > self.TOLERANCIA_DUPLICADOS:
                    importes_distintos.add(clave)
            else:
                mapa[clave] = valor
                repeticiones[clave] = 1

        avisos = []
        con_factura = len(facturas) - en_blanco
        if no_numericas:
            ejemplos = ", ".join(dict.fromkeys(no_numericas[:3]))
            avisos.append(
                f"VBRL: {len(no_numericas):,} de {con_factura:,} fila(s) con la "
                f"columna {letra_factura} ('{col_factura}') que no parece una factura "
                f"('C' + dígitos, ej.: {ejemplos}). Si son casi todas, revisa que el "
                f"layout del VBRL sea el esperado."
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

    def escribir(self, writer):
        """Write 'VBRL' with D as text and E as numbers, wide enough to show them."""
        df = self.resultado["df"]
        df.to_excel(writer, sheet_name=self.HOJA_VBRL, index=False)
        if not self.resultado["sin_datos"]:
            ws = writer.book[self.HOJA_VBRL]
            ws.column_dimensions[get_column_letter(self.COL_FACTURA)].width = 16
            ws.column_dimensions[get_column_letter(self.COL_IMPORTE)].width = 18


# === Stage: 'TD.Diferencias_Referencias' — Facturas, document by document ===
class EtapaTDReferencias:
    """The micro version of the 'TD' FACTURAS block: the same rows (EtapaTD.mascara_bloque) grouped by
    Centro × Nº documento × Referencia de pago (A:C) with the summed amount (D), plus two LIVE formulas:
    E 'Factura Monitor' (the reference, as text with its 'C', looked up in VBRL!D:D → VBRL!E:E;
    no match → 'Sin timbre') and F 'Diferencia' (D + E: invoices are negative in accounting, so
    0 = match). Values with an AutoFilter, not a native PivotTable (openpyxl cannot create one), so
    its Grand Total ties to the TD block's.
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

    # Reconciliation tolerance against the 'TD' FACTURAS Grand Total (cents).
    TOLERANCIA_CUADRE = 0.01

    # Excel's hard limit of rows per sheet.
    MAX_FILAS_EXCEL = 1_048_576

    AUTOR_COMENTARIOS = "Validación Factura Global"

    def __init__(self, gallo, pasos_gallo, etapa_td, etapa_vbrl):
        """`pasos_gallo` must come from EtapaGallo(con_ref_pago=True); `etapa_td` and `etapa_vbrl` must
        already be prepared.
        """
        self.gallo = gallo
        self.pasos = pasos_gallo
        self.etapa_td = etapa_td
        self.etapa_vbrl = etapa_vbrl
        self.resultado = None

    @staticmethod
    def formulas(fila):
        """The two formulas of ONE row (Factura Monitor, Diferencia), e.g. row 6:
            E6 =IF(C6="","Sin timbre",IFERROR(XLOOKUP(C6,VBRL!D:D,VBRL!E:E),"Sin timbre"))
            F6 =IFERROR(D6+E6,"Sin timbre")
        The reference and VBRL!D:D are both text with their 'C' ('C259003393'), so the lookup needs no
        conversion. The blank check keeps an empty reference from being looked up as a blank cell.
        """
        letra = get_column_letter
        referencia = f"{letra(EtapaTDReferencias.COL_REFERENCIA)}{fila}"
        importe    = f"{letra(EtapaTDReferencias.COL_IMPORTE)}{fila}"
        monitor    = f"{letra(EtapaTDReferencias.COL_MONITOR)}{fila}"
        hoja_vbrl  = Comun.ref_hoja(EtapaVBRL.HOJA_VBRL)
        col_factura = letra(EtapaVBRL.COL_FACTURA)
        col_importe = letra(EtapaVBRL.COL_IMPORTE)
        sin_timbre = Comun.texto_formula(Comun.TEXTO_SIN_TIMBRE)

        f_monitor = (
            f'=IF({referencia}="",{sin_timbre},IFERROR(_xlfn.XLOOKUP({referencia},'
            f"{hoja_vbrl}!{col_factura}:{col_factura},"
            f"{hoja_vbrl}!{col_importe}:{col_importe}),{sin_timbre}))"
        )
        f_diferencia = f"=IFERROR({importe}+{monitor},{sin_timbre})"
        return f_monitor, f_diferencia

    def preparar(self):
        """Group the FACTURAS rows of 'TD' by Centro × Nº documento × Referencia de pago (PivotTable
        order) and measure how they will cross against VBRL. Returns filas, total, avisos and the
        counters of the summary.
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
            # groupby sorts Centro, document, reference (PivotTable order); round to cents.
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

        # --- Controls: how will each row cross against VBRL? ---
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
            # XLOOKUP compares the trimmed text, ignoring case: same key as VBRL's `mapa`.
            clave = EtapaVBRL.clave_busqueda(ref)
            if clave in mapa:
                conteo_encontradas[clave] = conteo_encontradas.get(clave, 0) + 1
            elif not EtapaVBRL._RE_FACTURA.fullmatch(ref.strip()):
                invalidas_n += 1
                invalidas_importe += importe
                if len(ejemplos_invalidas) < 3 and ref not in ejemplos_invalidas:
                    ejemplos_invalidas.append(ref)
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

    def escribir(self, ws):
        """Write 'TD.Diferencias_Referencias' (rows as in 'TD': 1-3 legend, 4 title, 5 headers with an
        AutoFilter, 6+ data, then 'Grand Total'), panes frozen under the headers. The Grand Total
        'Diferencia' is D + E, total against total, so it includes the 'Sin timbre' rows.
        """
        resultado = self.resultado
        filas = resultado["filas"]
        columnas = resultado["columnas"]
        td_columnas = self.etapa_td.resultado["columnas"]
        vbrl = self.etapa_vbrl.resultado
        negrita = Font(bold=True)
        formato = Comun.FORMATO_NUMERO

        Comun.escribir_filtros(ws, EtapaTD.FILA_FILTROS, self.COL_CENTRO,
                               EtapaTD.leyenda_filtros(EtapaTD.FACTURAS, td_columnas))
        ws.cell(row=4, column=self.COL_CENTRO,
                value=f"{EtapaTD.FACTURAS['titulo']} — {columnas['documento']} × "
                      f"{columnas['referencia']}").font = negrita

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

        # Where the formula looks: VBRL!D:D → VBRL!E:E, citing the real headers when there are any.
        enc_factura, enc_importe = vbrl["encabezados"]
        col_vbrl_factura = get_column_letter(EtapaVBRL.COL_FACTURA)
        col_vbrl_importe = get_column_letter(EtapaVBRL.COL_IMPORTE)
        cita_factura = f" ('{enc_factura}')" if enc_factura else ""
        cita_importe = f" ('{enc_importe}')" if enc_importe else ""
        self._comentar(
            ws.cell(row=EtapaTD.FILA_ENCABEZADO, column=self.COL_MONITOR),
            f"Fórmula en cada renglón: busca la Referencia de pago (texto, con su 'C') en "
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

        fila = EtapaTD.FILA_INICIO
        for centro, documento, referencia, importe in filas:
            ws.cell(row=fila, column=self.COL_CENTRO, value=centro)
            ws.cell(row=fila, column=self.COL_DOCUMENTO, value=documento)
            # A blank reference stays empty (a PivotTable's "(blank)"); the formula handles it.
            ws.cell(row=fila, column=self.COL_REFERENCIA, value=referencia or None)
            ws.cell(row=fila, column=self.COL_IMPORTE, value=importe).number_format = formato
            f_monitor, f_diferencia = self.formulas(fila)
            ws.cell(row=fila, column=self.COL_MONITOR, value=f_monitor).number_format = formato
            ws.cell(row=fila, column=self.COL_DIFERENCIA, value=f_diferencia).number_format = formato
            fila += 1

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

            # AutoFilter over headers + data (not the total row), as the stakeholder filters 'Diferencia'.
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


# === Stage: FB (FB03) — Nº documento → Clave de referencia, for the credit notes and PEP vouchers ===
class EtapaFB:
    """The 'FB' sheet (FB03 download): for every credit note / PEP voucher document, the 'Clave de
    referencia' of the invoice it belongs to ('C' + digits, as VBRL carries it). The lookups of
    EtapaTDNotasPEP read it BY POSITION, like the stakeholder's workbook: FB!C:C (document, a real
    number) → FB!D:D (reference, text). In that workbook A:B hold SAP's list margin; here A:B stay
    EMPTY on purpose and the data starts at C, so the formulas can be copied unchanged.
    """

    HOJA_FB = "FB"

    # 1-based positions read by the lookups (C, D).
    COL_DOCUMENTO = 3
    COL_CLAVE = 4

    COL_DOCUMENTO_CANDIDATAS = ("Nº documento", "Nº doc.", "Número de documento", "Documento")
    COL_CLAVE_CANDIDATAS = ("Clave de referencia", "Clave ref.", "Clave referencia")

    # Excel keeps 15 significant digits: a longer document number would be silently rounded,
    # so it stays as text (and shows up in the warnings).
    MAX_DIGITOS = 15

    def __init__(self, fb, sin_datos):
        self.fb = fb
        self.sin_datos = sin_datos
        self.resultado = None

    @staticmethod
    def convertir_documento(valor):
        """A document cell as the formulas see it: '7075292768', '07075292768' and '7075292768.0' →
        7075292768 (what VALUE() builds from the text of a Gallo document); '' → None; anything that is
        not digits, or has more than MAX_DIGITOS of them, stays text.
        """
        texto = str(valor).strip()
        if texto == "":
            return None
        if re.fullmatch(r"[0-9]+(?:\.0+)?", texto):
            digitos = texto.split(".", 1)[0].lstrip("0") or "0"
            if len(digitos) <= EtapaFB.MAX_DIGITOS:
                return int(digitos)
        return texto

    def preparar(self):
        """Document column to numbers, reference column to trimmed text, indexed the way XLOOKUP reads
        them (first match wins). Returns df, sin_datos, filas, mapa ({document number: reference
        upper-cased, '' if blank}), avisos and encabezados. ValueError: the document or reference
        column is missing from the download.
        """
        fb = self.fb
        if self.sin_datos or fb.empty:
            self.resultado = {
                "df":          fb,
                "sin_datos":   True,
                "filas":       0,
                "mapa":        {},
                "avisos": [
                    "FB: el archivo vino sin datos; las tablas 'Notas de crédito' y "
                    "'PEP Vouchers' no podrán encontrar ninguna referencia (#N/A)."
                ],
                "encabezados": ("", ""),
            }
            return self.resultado

        col_doc = Comun.col_por_candidatos(
            fb, self.COL_DOCUMENTO_CANDIDATAS, "Nº documento", origen="FB03")
        col_clave = Comun.col_por_candidatos(
            fb, self.COL_CLAVE_CANDIDATAS, "Clave de referencia", origen="FB03")
        columnas = list(fb.columns)
        i_doc = columnas.index(col_doc)
        i_clave = columnas.index(col_clave)
        if i_doc == i_clave:
            raise ValueError(
                f"El archivo de FB03 usa la misma columna ('{col_doc}') como documento y "
                f"como clave de referencia. Columnas: {columnas}"
            )

        documentos = [self.convertir_documento(v) for v in fb.iloc[:, i_doc]]
        claves = [EtapaVBRL.limpiar_factura(v) for v in fb.iloc[:, i_clave]]

        # First row of each document wins, as in XLOOKUP.
        mapa = {}
        distintas = {}
        sin_documento = no_numericos = sin_clave = no_parecen = 0
        ejemplos_no_parecen = []
        for documento, clave in zip(documentos, claves):
            if clave is None:
                sin_clave += 1
            elif not EtapaVBRL._RE_FACTURA.fullmatch(clave):
                no_parecen += 1
                if len(ejemplos_no_parecen) < 3 and clave not in ejemplos_no_parecen:
                    ejemplos_no_parecen.append(clave)
            if documento is None:
                sin_documento += 1
                continue
            if not isinstance(documento, int):
                no_numericos += 1
                continue
            texto_clave = EtapaVBRL.clave_busqueda(clave) if clave is not None else ""
            texto_clave = texto_clave or ""
            if documento not in mapa:
                mapa[documento] = texto_clave
                distintas[documento] = {texto_clave}
            else:
                distintas[documento].add(texto_clave)

        con_claves_distintas = sum(1 for valores in distintas.values() if len(valores) > 1)
        avisos = []
        if sin_documento:
            avisos.append(f"FB: {sin_documento:,} fila(s) sin '{col_doc}'.")
        if no_numericos:
            avisos.append(
                f"FB: {no_numericos:,} fila(s) con un '{col_doc}' que no es un número "
                f"(o trae más de {self.MAX_DIGITOS} dígitos); no se pueden buscar."
            )
        if sin_clave:
            avisos.append(f"FB: {sin_clave:,} fila(s) sin '{col_clave}'.")
        if no_parecen:
            avisos.append(
                f"FB: {no_parecen:,} '{col_clave}' que no parecen factura "
                f"('C' + dígitos, ej.: {', '.join(ejemplos_no_parecen)})."
            )
        if con_claves_distintas:
            avisos.append(
                f"FB: {con_claves_distintas:,} documento(s) aparecen con claves de "
                f"referencia distintas; XLOOKUP toma sólo la primera."
            )

        df = fb.iloc[:, [i_doc, i_clave]
                     + [j for j in range(fb.shape[1]) if j not in (i_doc, i_clave)]].copy()
        # By position: the export may repeat a header name.
        df.isetitem(0, pd.Series(documentos, index=df.index, dtype=object))
        df.isetitem(1, pd.Series(claves, index=df.index, dtype=object))

        self.resultado = {
            "df":          df,
            "sin_datos":   False,
            "filas":       len(df),
            "mapa":        mapa,
            "avisos":      avisos,
            "encabezados": (str(col_doc), str(col_clave)),
        }
        return self.resultado

    def escribir(self, writer):
        """Write 'FB' with the data starting at column C (A:B empty), document as number, wide enough to
        show it, and a note on the first header explaining the layout.
        """
        df = self.resultado["df"]
        df.to_excel(writer, sheet_name=self.HOJA_FB, index=False,
                    startcol=self.COL_DOCUMENTO - 1)
        ws = writer.book[self.HOJA_FB]
        ws.column_dimensions[get_column_letter(self.COL_DOCUMENTO)].width = 16
        ws.column_dimensions[get_column_letter(self.COL_CLAVE)].width = 20
        EtapaTDReferencias._comentar(
            ws.cell(row=1, column=self.COL_DOCUMENTO),
            f"Hoja {self.HOJA_FB} (FB03). Las columnas A:B quedan vacías a propósito: los datos "
            f"empiezan en la C porque las fórmulas de 'Notas de crédito' y 'PEP Vouchers' leen "
            f"{self.HOJA_FB}!{get_column_letter(self.COL_DOCUMENTO)}:{get_column_letter(self.COL_DOCUMENTO)} "
            f"(documento, como número) y "
            f"{self.HOJA_FB}!{get_column_letter(self.COL_CLAVE)}:{get_column_letter(self.COL_CLAVE)} "
            f"(clave de referencia), igual que el Excel de referencia.",
        )


# === Stage: 'TD.Diferencias_Referencias' — Notas de crédito (H:M) and PEP Vouchers (O:T) ===
class EtapaTDNotasPEP:
    """Two more tables on 'TD.Diferencias_Referencias', beside the Facturas one (A:F), with the six
    columns of the stakeholder's reference workbook ('TD.DiferenciasNC' / 'TD.DiferenciasPP'):
        Centro | Nº documento | Importe | Referencia de pago | Factura Monitor | Diferencia
    The first three are values (the PivotTable part: Centro × Nº documento); the last three are LIVE
    formulas copied VERBATIM from that workbook (only the column letters move with the block):
        Referencia de pago  =+XLOOKUP(VALUE(<doc>),FB!C:C,FB!D:D)
        Factura Monitor     =+XLOOKUP(<ref>,VBRL!D:D,VBRL!E:E)
        Diferencia          =<importe>-<monitor>
    There is NO IFERROR, as in the reference: a document that is not in FB, or a reference that is
    not in VBRL, shows #N/A (the warnings count them). Each block is an Excel Table (a sheet takes
    only one plain AutoFilter, and Facturas already has it).
        Notas de crédito (H:M)  Referencia 3 = 'Global', Referencia 2 ≠ 'PP', Clase C2 / C6 — the
                                rows of 'TD' NOTAS DE CRÉDITO without its 'PP' ones
        PEP Vouchers (O:T)      Referencia 3 = 'Global', Referencia 2 = 'PP', Clase C2 — the rows of
                                'TD' PEP VOUCHERS
    Both take their rows from EtapaTD, so their totals tie to the 'TD' blocks.
    """

    HOJA = EtapaTDReferencias.HOJA_TD_REFERENCIAS

    # Position of each column inside a block (0-based from its first column).
    OFF_CENTRO     = 0
    OFF_DOCUMENTO  = 1
    OFF_IMPORTE    = 2
    OFF_REFERENCIA = 3
    OFF_MONITOR    = 4
    OFF_DIFERENCIA = 5
    ANCHO_BLOQUE   = 6

    ENCABEZADO_REFERENCIA = "Referencia de pago"
    ENCABEZADO_MONITOR    = EtapaTDReferencias.ENCABEZADO_MONITOR
    ENCABEZADO_DIFERENCIA = EtapaTDReferencias.ENCABEZADO_DIFERENCIA

    NOTAS = {"clave": "notas", "titulo": "Notas de crédito", "col_inicio": 8,    # H
             "tabla": "TblNotasCredito"}
    PEP   = {"clave": "pep",   "titulo": "PEP Vouchers",     "col_inicio": 15,   # O
             "tabla": "TblPEPVouchers"}
    BLOQUES = (NOTAS, PEP)

    # Reconciliation tolerance against the 'TD' blocks (cents).
    TOLERANCIA_CUADRE = 0.01

    # Excel's hard limit of rows per sheet.
    MAX_FILAS_EXCEL = 1_048_576

    def __init__(self, gallo, pasos_gallo, etapa_td, etapa_vbrl, etapa_fb):
        """`etapa_td`, `etapa_vbrl` and `etapa_fb` must already be prepared."""
        self.gallo = gallo
        self.pasos = pasos_gallo
        self.etapa_td = etapa_td
        self.etapa_vbrl = etapa_vbrl
        self.etapa_fb = etapa_fb
        self.resultado = None

    @staticmethod
    def formulas(col_inicio, fila):
        """The three formulas of ONE row (Referencia de pago, Factura Monitor, Diferencia) of the block
        that starts at `col_inicio`, e.g. the block H:M, row 6:
            K6 =+_xlfn.XLOOKUP(VALUE(I6),FB!C:C,FB!D:D)
            L6 =+_xlfn.XLOOKUP(K6,VBRL!D:D,VBRL!E:E)
            M6 =J6-L6
        Same text as the reference workbook's D6:F6 with the letters of the block.
        """
        letra = get_column_letter
        documento  = f"{letra(col_inicio + EtapaTDNotasPEP.OFF_DOCUMENTO)}{fila}"
        importe    = f"{letra(col_inicio + EtapaTDNotasPEP.OFF_IMPORTE)}{fila}"
        referencia = f"{letra(col_inicio + EtapaTDNotasPEP.OFF_REFERENCIA)}{fila}"
        monitor    = f"{letra(col_inicio + EtapaTDNotasPEP.OFF_MONITOR)}{fila}"
        hoja_fb    = Comun.ref_hoja(EtapaFB.HOJA_FB)
        fb_doc     = letra(EtapaFB.COL_DOCUMENTO)
        fb_clave   = letra(EtapaFB.COL_CLAVE)
        hoja_vbrl  = Comun.ref_hoja(EtapaVBRL.HOJA_VBRL)
        vbrl_fact  = letra(EtapaVBRL.COL_FACTURA)
        vbrl_imp   = letra(EtapaVBRL.COL_IMPORTE)

        f_referencia = (
            f"=+_xlfn.XLOOKUP(VALUE({documento}),"
            f"{hoja_fb}!{fb_doc}:{fb_doc},{hoja_fb}!{fb_clave}:{fb_clave})"
        )
        f_monitor = (
            f"=+_xlfn.XLOOKUP({referencia},"
            f"{hoja_vbrl}!{vbrl_fact}:{vbrl_fact},{hoja_vbrl}!{vbrl_imp}:{vbrl_imp})"
        )
        f_diferencia = f"={importe}-{monitor}"
        return f_referencia, f_monitor, f_diferencia

    @staticmethod
    def leyenda_notas(columnas):
        """Filter legend of the Notas de crédito table, from the same constants that build its mask:
        the 'TD' NOTAS legend plus the 'PP' exclusion.
        """
        clases = " o ".join(f"'{clase}'" for clase, _ in EtapaTD.NOTAS["clases"])
        return [
            f"Filtros — hoja '{EtapaGallo.HOJA_GALLO}':",
            f"Referencia 3 = 'Global'; Referencia 2 ≠ '{EtapaTD.PEP_REFERENCIA2}'; "
            f"{columnas['centro']} no vacío",
            f"{columnas['asignacion']} ≠ '{EtapaGallo.VALOR_ASIGNACION}'; "
            f"{columnas['clase']} = {clases}",
        ]

    def _mascara(self, bloque):
        """Mask of the Gallo rows of a block, taken from EtapaTD so it cannot drift from 'TD'."""
        td = self.etapa_td
        if bloque is self.NOTAS:
            sin_pp = td.ref2_norm != EtapaTD.PEP_REFERENCIA2
            return td.mascara_bloque(EtapaTD.NOTAS["clases"]) & sin_pp
        return td.mascara_pep()

    def preparar(self):
        """Group each block's rows by Centro × Nº documento (PivotTable order) and measure how its
        formulas will resolve against FB and VBRL. Returns {'bloques': {clave: {filas, total,
        sin_na ...}}, 'avisos', 'columnas'}.
        """
        gallo = self.gallo
        pasos = self.pasos
        td = self.etapa_td
        col_doc = pasos["col_doc"]
        importe_num = pasos["importe_num"]
        mapa_fb = self.etapa_fb.resultado["mapa"]
        mapa_vbrl = self.etapa_vbrl.resultado["mapa"]
        etiqueta = EtapaTDReferencias.HOJA_TD_REFERENCIAS

        bloques = {}
        avisos = []
        for bloque in self.BLOQUES:
            mascara = self._mascara(bloque)
            datos = pd.DataFrame({
                "_centro":  td.centro_norm[mascara],
                "_doc":     gallo.loc[mascara, col_doc].astype(str).str.strip(),
                "_importe": importe_num[mascara],
            })
            if datos.empty:
                filas = []
            else:
                # groupby sorts Centro, document (PivotTable order); round to cents.
                serie = datos.groupby(["_centro", "_doc"])["_importe"].sum()
                filas = [(centro, doc, round(float(valor), 2))
                         for (centro, doc), valor in serie.items()]

            if EtapaTD.FILA_INICIO + len(filas) > self.MAX_FILAS_EXCEL:
                raise ValueError(
                    f"{etiqueta} ({bloque['titulo']}) tendría {len(filas):,} renglones y "
                    f"no caben en una hoja de Excel ({self.MAX_FILAS_EXCEL:,} filas). "
                    f"Consolida un rango de fechas más corto."
                )
            total = round(sum(importe for _, _, importe in filas), 2)

            # --- Controls: how will each row resolve? (#N/A = no IFERROR, like the reference) ---
            sin_numero_n = sin_fb_n = sin_clave_n = sin_vbrl_n = 0
            sin_numero_i = sin_fb_i = sin_clave_i = sin_vbrl_i = 0.0
            for _, doc, importe in filas:
                numero = EtapaFB.convertir_documento(doc)
                if not isinstance(numero, int):
                    sin_numero_n += 1
                    sin_numero_i += importe
                elif numero not in mapa_fb:
                    sin_fb_n += 1
                    sin_fb_i += importe
                elif mapa_fb[numero] == "":
                    sin_clave_n += 1
                    sin_clave_i += importe
                elif mapa_fb[numero] not in mapa_vbrl:
                    sin_vbrl_n += 1
                    sin_vbrl_i += importe
            na_n = sin_numero_n + sin_fb_n + sin_clave_n + sin_vbrl_n
            na_importe = round(sin_numero_i + sin_fb_i + sin_clave_i + sin_vbrl_i, 2)

            if na_n:
                partes = []
                if sin_numero_n:
                    partes.append(f"{sin_numero_n:,} con un Nº documento que no es número")
                if sin_fb_n:
                    partes.append(f"{sin_fb_n:,} que no están en FB")
                if sin_clave_n:
                    partes.append(f"{sin_clave_n:,} con clave de referencia vacía en FB")
                if sin_vbrl_n:
                    partes.append(f"{sin_vbrl_n:,} cuya referencia no está en VBRL")
                avisos.append(
                    f"{etiqueta} ({bloque['titulo']}): {na_n:,} de {len(filas):,} "
                    f"renglón(es) quedarán en #N/A (importe contable {na_importe:,.2f}): "
                    f"{'; '.join(partes)}."
                )

            bloques[bloque["clave"]] = {
                "filas":      filas,
                "total":      total,
                "na_n":       na_n,
                "na_importe": na_importe,
            }

        # --- Reconciliation with the 'TD' blocks ---
        td_resumen = td.resultado
        mascara_notas = td.mascara_bloque(EtapaTD.NOTAS["clases"])
        con_pp = mascara_notas & (td.ref2_norm == EtapaTD.PEP_REFERENCIA2)
        pp_en_notas = round(float(importe_num[con_pp].sum()), 2)
        esperado_notas = round(td_resumen["total_notas"] - pp_en_notas, 2)
        esperados = (
            (self.NOTAS, esperado_notas,
             f"{EtapaTD.NOTAS['titulo']} de 'TD' sin sus renglones 'PP'"),
            (self.PEP, td_resumen["total_pep"], f"{EtapaTD.PEP['titulo']} de 'TD'"),
        )
        for bloque, esperado, descripcion in esperados:
            total = bloques[bloque["clave"]]["total"]
            if abs(total - esperado) > self.TOLERANCIA_CUADRE:
                avisos.append(
                    f"{etiqueta} ({bloque['titulo']}): el Grand Total ({total:,.2f}) "
                    f"no cuadra con {descripcion} ({esperado:,.2f}); diferencia "
                    f"{round(total - esperado, 2):,.2f}."
                )

        self.resultado = {
            "bloques": bloques,
            "avisos":  avisos,
            "columnas": {
                "centro":    pasos["col_centro"],
                "documento": col_doc,
            },
        }
        return self.resultado

    def escribir(self, ws):
        """Write both blocks on the (already written) 'TD.Diferencias_Referencias': rows 1-3 legend, 4
        title, 5 headers (the Excel Table), 6+ data, then 'Grand Total' (the amount only, as in the
        reference). A block with no rows keeps its title and headers but no Table (Excel does not
        accept a Table without data rows).
        """
        resultado = self.resultado
        td_columnas = self.etapa_td.resultado["columnas"]
        columnas = resultado["columnas"]
        negrita = Font(bold=True)
        formato = Comun.FORMATO_NUMERO
        letra = get_column_letter
        fila_encabezado = EtapaTD.FILA_ENCABEZADO
        fila_inicio = EtapaTD.FILA_INICIO

        for bloque in self.BLOQUES:
            datos = resultado["bloques"][bloque["clave"]]
            filas = datos["filas"]
            inicio = bloque["col_inicio"]

            if bloque is self.NOTAS:
                leyenda = self.leyenda_notas(td_columnas)
            else:
                leyenda = EtapaTD.leyenda_filtros_pep(td_columnas)
            Comun.escribir_filtros(ws, EtapaTD.FILA_FILTROS, inicio, leyenda)
            titulo = ws.cell(row=4, column=inicio, value=bloque["titulo"])
            titulo.font = negrita
            EtapaTDReferencias._comentar(
                titulo,
                "Cada tabla tiene su propio filtro, pero un filtro oculta RENGLONES COMPLETOS "
                "de la hoja: al filtrar una tabla también se ocultan los renglones de las demás "
                "que quedan en esas mismas filas. Quita el filtro antes de revisar otra tabla.",
            )

            encabezados = (
                (self.OFF_CENTRO,     columnas["centro"]),
                (self.OFF_DOCUMENTO,  columnas["documento"]),
                (self.OFF_IMPORTE,    Comun.ENCABEZADO_IMPORTE),
                (self.OFF_REFERENCIA, self.ENCABEZADO_REFERENCIA),
                (self.OFF_MONITOR,    self.ENCABEZADO_MONITOR),
                (self.OFF_DIFERENCIA, self.ENCABEZADO_DIFERENCIA),
            )
            for desplazamiento, texto in encabezados:
                ws.cell(row=fila_encabezado, column=inicio + desplazamiento,
                        value=str(texto)).font = negrita

            letra_fb_doc = letra(EtapaFB.COL_DOCUMENTO)
            letra_fb_clave = letra(EtapaFB.COL_CLAVE)
            letra_vbrl_f = letra(EtapaVBRL.COL_FACTURA)
            letra_vbrl_i = letra(EtapaVBRL.COL_IMPORTE)
            EtapaTDReferencias._comentar(
                ws.cell(row=fila_encabezado, column=inicio + self.OFF_REFERENCIA),
                f"Fórmula en cada renglón: busca el Nº documento (como número) en "
                f"{EtapaFB.HOJA_FB}!{letra_fb_doc}:{letra_fb_doc} y regresa "
                f"{EtapaFB.HOJA_FB}!{letra_fb_clave}:{letra_fb_clave} (clave de referencia). "
                f"Sin coincidencia: #N/A (sin IFERROR, igual que el Excel de referencia).",
            )
            EtapaTDReferencias._comentar(
                ws.cell(row=fila_encabezado, column=inicio + self.OFF_MONITOR),
                f"Fórmula en cada renglón: busca la Referencia de pago (texto, con su 'C') en "
                f"{EtapaVBRL.HOJA_VBRL}!{letra_vbrl_f}:{letra_vbrl_f} y regresa "
                f"{EtapaVBRL.HOJA_VBRL}!{letra_vbrl_i}:{letra_vbrl_i} (importe). "
                f"Sin coincidencia: #N/A.",
            )
            EtapaTDReferencias._comentar(
                ws.cell(row=fila_encabezado, column=inicio + self.OFF_DIFERENCIA),
                "Importe contable − Factura Monitor (igual que el Excel de referencia). "
                "Un renglón con #N/A en Factura Monitor también da #N/A aquí.",
            )

            fila = fila_inicio
            for centro, documento, importe in filas:
                ws.cell(row=fila, column=inicio + self.OFF_CENTRO, value=centro)
                # Text on purpose (as in the reference): the formula turns it into a number.
                ws.cell(row=fila, column=inicio + self.OFF_DOCUMENTO, value=documento)
                ws.cell(row=fila, column=inicio + self.OFF_IMPORTE,
                        value=importe).number_format = formato
                f_ref, f_monitor, f_dif = self.formulas(inicio, fila)
                ws.cell(row=fila, column=inicio + self.OFF_REFERENCIA, value=f_ref)
                ws.cell(row=fila, column=inicio + self.OFF_MONITOR,
                        value=f_monitor).number_format = formato
                ws.cell(row=fila, column=inicio + self.OFF_DIFERENCIA,
                        value=f_dif).number_format = formato
                fila += 1

            if filas:
                ultima = fila - 1
                ws.cell(row=fila, column=inicio + self.OFF_CENTRO,
                        value=Comun.ETIQUETA_GRAND_TOTAL).font = negrita
                celda = ws.cell(row=fila, column=inicio + self.OFF_IMPORTE, value=datos["total"])
                celda.number_format = formato
                celda.font = negrita

                tabla = Table(
                    displayName=bloque["tabla"],
                    ref=(f"{letra(inicio)}{fila_encabezado}:"
                         f"{letra(inicio + self.ANCHO_BLOQUE - 1)}{ultima}"),
                )
                tabla.tableStyleInfo = TableStyleInfo(
                    name=None, showFirstColumn=False, showLastColumn=False,
                    showRowStripes=False, showColumnStripes=False)
                ws.add_table(tabla)

            for desplazamiento, ancho in enumerate((12, 16, 22, 20, 20, 20)):
                ws.column_dimensions[letra(inicio + desplazamiento)].width = ancho

        # Spacer columns between the tables (G, N).
        for bloque in self.BLOQUES:
            ws.column_dimensions[letra(bloque["col_inicio"] - 1)].width = 4


# === Entry point helpers ===
def carpeta_output_desde_input(ruta_input):
    """The 'Output' folder for a download folder: next to it when it is called 'Input' (src/Input →
    src/Output), otherwise INSIDE it, so no files land outside the user's choice.
    """
    ruta = os.path.normpath(ruta_input)
    if os.path.basename(ruta).lower() == "input":
        return os.path.join(os.path.dirname(ruta), "Output")
    return os.path.join(ruta, "Output")


def leer_documentos_monivoi(ruta_monivoi, columna):
    """The distinct values, in order, of the Monivoi column a download pastes into SAP
    (`columna='factura'` → VBRL) → (valores, filas_leidas, nombre_columna).
    """
    return EtapaMonivoi.leer_documentos(ruta_monivoi, columna)


def leer_documentos_gallo_fb03(ruta_gallo):
    """The 'Nº documento' values (distinct, in order) the FB03 download pastes into SAP: Gallo rows with
    Clase C2/C6 and a Global 'Referencia' → (valores, filas_leidas, nombre_columna).
    """
    return EtapaDocumentosFB03.leer(ruta_gallo)


# === Orchestrator ===
class Consolidador:
    """Runs the consolidation: reads the CSVs, runs the stages in order and writes the workbook (the
    stages hold the logic): EtapaGallo → EtapaTD → EtapaMonivoi → EtapaCruceTDMTD → [EtapaVBRL →
    EtapaTDReferencias → EtapaFB → EtapaTDNotasPEP] → EtapaVentas. The VBRL and the FB03 are
    OPTIONAL on purpose: an old date from before those downloads must still consolidate (just
    without their sheets; the Notas de crédito / PEP Vouchers tables need BOTH).
    """

    def __init__(self, ruta_gallo, ruta_monivoi, ruta_output, nombre_salida=None,
                 callback_status=None, fechas_seleccionadas=None, ruta_vbrl=None, ruta_fb03=None):
        """`fechas_seleccionadas` ('DD.MM.YYYY' from the GUI) make those days show in 'Venta DSD' even
        without amounts and break ties in the date format; `ruta_vbrl` adds the 'VBRL' and
        'TD.Diferencias_Referencias' sheets; `ruta_fb03` (together with `ruta_vbrl`) adds the 'FB' sheet
        and the 'Notas de crédito' / 'PEP Vouchers' tables; `nombre_salida` defaults to
        'Consolidado_<Gallo CSV>.xlsx'.
        """
        self.ruta_gallo = ruta_gallo
        self.ruta_monivoi = ruta_monivoi
        self.ruta_output = ruta_output
        self.nombre_salida = nombre_salida
        self.callback_status = callback_status
        self.fechas_seleccionadas = fechas_seleccionadas
        self.ruta_vbrl = ruta_vbrl or None
        self.ruta_fb03 = ruta_fb03 or None

    def _status(self, mensaje):
        """Report progress to the caller (the GUI), if it asked for it."""
        if self.callback_status:
            self.callback_status(mensaje)

    def ejecutar(self):
        """Run the whole consolidation → the summary dict the GUI shows (path, row counts, totals and the
        warnings of every stage). FileNotFoundError: a CSV is missing. ValueError: input the user can
        fix (Gallo without data, a missing column, an unreadable amount or date, a malformed GUI date,
        a VBRL or FB03 without the columns the lookups need).
        """
        update_status = self._status
        con_vbrl = self.ruta_vbrl is not None
        con_fb = self.ruta_fb03 is not None
        # The tables read FB AND VBRL: with only one of them there is nothing to look up.
        con_notas_pep = con_vbrl and con_fb

        # Validated BEFORE reading anything: better in the first second than after 1M rows.
        fechas_sel = EtapaVentas.parsear_fechas_seleccionadas(self.fechas_seleccionadas)

        # Step 1: read the files.
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

        # The VBRL is read AND checked up front: a wrong layout stops the run now.
        etapa_vbrl = None
        if con_vbrl:
            update_status("📄 Leyendo archivo de VBRL...")
            vbrl = Comun.leer_csv(self.ruta_vbrl)
            etapa_vbrl = EtapaVBRL(vbrl, Comun.es_sin_datos(vbrl))
            etapa_vbrl.preparar()

        # Same for the FB03 (only useful together with the VBRL).
        etapa_fb = None
        if con_notas_pep:
            update_status("📄 Leyendo archivo de FB03...")
            fb = Comun.leer_csv(self.ruta_fb03)
            etapa_fb = EtapaFB(fb, Comun.es_sin_datos(fb))
            etapa_fb.preparar()

        # Gallo: steps 2-9 + SG/RG override + amounts (adds K, L, M).
        etapa_gallo = EtapaGallo(gallo, update_status, con_ref_pago=con_vbrl)
        pasos_gallo = etapa_gallo.preparar()

        # 'TD' matrices.
        update_status("📊 Construyendo matriz 'TD'...")
        etapa_td = EtapaTD(gallo, pasos_gallo)
        td_resumen = etapa_td.preparar()
        total_importe_gallo = td_resumen["total_importe_gallo"]

        # Monivoi: step 16, 'MTD'.
        update_status("🧾 Construyendo matrices 'MTD' (SG / RG / PP / TODOS)...")
        etapa_monivoi = EtapaMonivoi(monivoi, monivoi_sin_datos)
        mtd = etapa_monivoi.preparar()
        # 'Referencia 2' AFTER the matrices on purpose: they use their own series.
        etapa_monivoi.insertar_referencia2()

        # TD × MTD cross: what the formulas cannot show.
        cruce = EtapaCruceTDMTD(td_resumen, mtd)
        cruce_avisos = cruce.avisos()

        # 'TD.Diferencias_Referencias' (only with VBRL).
        etapa_td_ref = None
        if con_vbrl:
            update_status("🔬 Construyendo 'TD.Diferencias_Referencias'...")
            etapa_td_ref = EtapaTDReferencias(gallo, pasos_gallo, etapa_td, etapa_vbrl)
            etapa_td_ref.preparar()

        # 'Notas de crédito' and 'PEP Vouchers' (only with VBRL and FB03).
        etapa_notas_pep = None
        if con_notas_pep:
            update_status("🔬 Construyendo 'Notas de crédito' y 'PEP Vouchers'...")
            etapa_notas_pep = EtapaTDNotasPEP(gallo, pasos_gallo, etapa_td, etapa_vbrl, etapa_fb)
            etapa_notas_pep.preparar()

        # 'Venta DSD'.
        update_status("📆 Construyendo 'Venta DSD'...")
        etapa_ventas = EtapaVentas(gallo, pasos_gallo, fechas_sel)
        ventas = etapa_ventas.preparar()

        # Export the workbook.
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

            if etapa_vbrl is not None:
                etapa_vbrl.escribir(writer)

            if etapa_fb is not None:
                etapa_fb.escribir(writer)

            etapa_gallo.escribir_validacion(
                writer.book.create_sheet(EtapaGallo.HOJA_VALIDACION))

            ws_td = writer.book.create_sheet(EtapaTD.HOJA_TD)
            etapa_td.escribir(ws_td)
            cruce.escribir(ws_td)
            # openpyxl saves formulas WITHOUT results: Excel recalculates everything when it opens the file.
            writer.book.calculation.fullCalcOnLoad = True

            if etapa_td_ref is not None:
                ws_referencias = writer.book.create_sheet(EtapaTDReferencias.HOJA_TD_REFERENCIAS)
                etapa_td_ref.escribir(ws_referencias)
                if etapa_notas_pep is not None:
                    etapa_notas_pep.escribir(ws_referencias)

            etapa_monivoi.escribir_mtd(writer.book.create_sheet(EtapaMonivoi.HOJA_MTD))

            for clave, nombre_hoja in EtapaVentas.HOJAS:
                etapa_ventas.escribir(writer.book.create_sheet(nombre_hoja), clave)

        update_status("✅ Consolidación completada")

        resumen_vbrl = etapa_vbrl.resultado if etapa_vbrl is not None else None
        resumen_ref = etapa_td_ref.resultado if etapa_td_ref is not None else None
        resumen_fb = etapa_fb.resultado if etapa_fb is not None else None
        resumen_np = etapa_notas_pep.resultado if etapa_notas_pep is not None else None
        notas_pep_avisos = []
        if con_vbrl != con_fb:
            faltante = "FB03" if con_vbrl else "VBRL"
            notas_pep_avisos.append(
                f"No se generaron las tablas 'Notas de crédito' y 'PEP Vouchers' ni la "
                f"hoja {EtapaFB.HOJA_FB}: sus fórmulas necesitan el VBRL y el FB03 y "
                f"faltó el de {faltante}."
            )
        if resumen_fb:
            notas_pep_avisos.extend(resumen_fb["avisos"])
        if resumen_np:
            notas_pep_avisos.extend(resumen_np["avisos"])
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
            "vbrl_incluido":        con_vbrl,
            "vbrl_sin_datos":       bool(resumen_vbrl and resumen_vbrl["sin_datos"]),
            "filas_vbrl":           resumen_vbrl["filas"] if resumen_vbrl else 0,
            "vbrl_avisos":          resumen_vbrl["avisos"] if resumen_vbrl else [],
            "filas_td_referencias": len(resumen_ref["filas"]) if resumen_ref else 0,
            "total_td_referencias": resumen_ref["total"] if resumen_ref else 0.0,
            "td_referencias_sin_timbre": resumen_ref["sin_timbre_n"] if resumen_ref else 0,
            "td_referencias_avisos": resumen_ref["avisos"] if resumen_ref else [],
            "fb03_incluido":        con_notas_pep,
            "filas_fb":             resumen_fb["filas"] if resumen_fb else 0,
            "filas_notas_ref":      len(resumen_np["bloques"]["notas"]["filas"]) if resumen_np else 0,
            "total_notas_ref":      resumen_np["bloques"]["notas"]["total"] if resumen_np else 0.0,
            "notas_ref_na":         resumen_np["bloques"]["notas"]["na_n"] if resumen_np else 0,
            "filas_pep_ref":        len(resumen_np["bloques"]["pep"]["filas"]) if resumen_np else 0,
            "total_pep_ref":        resumen_np["bloques"]["pep"]["total"] if resumen_np else 0.0,
            "pep_ref_na":           resumen_np["bloques"]["pep"]["na_n"] if resumen_np else 0,
            "notas_pep_avisos":     notas_pep_avisos,
        }


def consolidar_gallo_monivoi(
    ruta_gallo,
    ruta_monivoi,
    ruta_output,
    nombre_salida=None,
    callback_status=None,
    fechas_seleccionadas=None,
    ruta_vbrl=None,
    ruta_fb03=None,
):
    """The controller's entry point: Consolidador(...).ejecutar() (see its docs)."""
    return Consolidador(
        ruta_gallo, ruta_monivoi, ruta_output,
        nombre_salida=nombre_salida,
        callback_status=callback_status,
        fechas_seleccionadas=fechas_seleccionadas,
        ruta_vbrl=ruta_vbrl,
        ruta_fb03=ruta_fb03,
    ).ejecutar()


def consolidar_venta_eiap(ruta_eiap, ruta_output, nombre_salida=None, callback_status=None,
                          fechas_seleccionadas=None):
    """The controller's entry point when ONLY 'Venta EIAP' was downloaded: a workbook with the single
    'Venta EIAP' tab. `fechas_seleccionadas` ('DD.MM.YYYY'): every day of the queried range, which
    show up even without amounts. FileNotFoundError / ValueError as in Consolidador.ejecutar().
    """
    status = callback_status or (lambda _mensaje: None)
    fechas_sel = EtapaVentas.parsear_fechas_seleccionadas(fechas_seleccionadas)

    status("📄 Leyendo archivo de Venta EIAP...")
    eiap = Comun.leer_csv(ruta_eiap)
    if Comun.es_sin_datos(eiap):
        raise ValueError(
            "El archivo de Venta EIAP no tiene datos (descarga 'sin partidas'). "
            "No hay nada que consolidar."
        )

    status("📆 Construyendo 'Venta EIAP'...")
    etapa = EtapaVentaEIAP.desde_csv(eiap, fechas_sel)
    ventas = etapa.preparar()

    status("💾 Guardando Excel consolidado...")
    os.makedirs(ruta_output, exist_ok=True)
    if not nombre_salida:
        nombre_salida = f"Consolidado_{os.path.splitext(os.path.basename(ruta_eiap))[0]}.xlsx"
    ruta_final = os.path.join(ruta_output, nombre_salida)

    libro = Workbook()
    hoja = libro.active
    hoja.title = EtapaVentaEIAP.HOJA_EIAP
    etapa.escribir(hoja, "eiap")
    libro.save(ruta_final)

    status("✅ Consolidación completada")
    return {
        "ruta":             ruta_final,
        "filas_eiap":       len(eiap),
        "total_venta_eiap": ventas["eiap"]["total"],
        "venta_avisos":     ventas["avisos"],
    }


# Quick test without the GUI:  python Consolidacion.py
if __name__ == "__main__":
    print("=== Prueba de consolidación (standalone) ===")
    ruta_gallo   = input("Ruta del CSV de Gallo   : ").strip().strip('"')
    ruta_monivoi = input("Ruta del CSV de Monivoi : ").strip().strip('"')
    ruta_vbrl    = input("Ruta del CSV de VBRL (Enter = sin VBRL): ").strip().strip('"') or None
    ruta_fb03    = input("Ruta del CSV de FB03 (Enter = sin FB03): ").strip().strip('"') or None
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
            ruta_fb03=ruta_fb03,
        )
        print("\n🎉 CONSOLIDACIÓN COMPLETADA 🎉")
        for clave, valor in resumen.items():
            print(f"  {clave:<20}: {valor}")
    except Exception as e:
        print(f"❌ Error durante la consolidación: {e}")
        raise
