"""Validación Factura Global — consolidation: the CSVs downloaded from SAP (Gallo FAGLL03,
Monivoi ZLMXCOM and, optionally, VBRL ZSE16) → ONE Excel workbook. Tabs, in order:
    Gallo, Monivoi, VBRL*, FB**, Validación - Doc., TD, TD.Diferencias_Referencias*, MTD,
    Validación_doc*
    (* only with a VBRL file; ** only with a VBRL and an FB03 file, which also add the 'Notas de
    crédito' and 'PEP Vouchers' tables to 'TD.Diferencias_Referencias' and to 'Validación_doc')

One class per stage, each with its constants, its computation (preparar → result dict) and
the sheet it writes (escribir): Comun (shared helpers) → EtapaGallo → EtapaTD →
EtapaDocumentosFB03 → EtapaVentas → EtapaMonivoi → EtapaCruceTDMTD → EtapaVBRL →
EtapaTDReferencias → EtapaFB → EtapaTDNotasPEP → EtapaValidacionDoc (the #N/A rows of those tables,
side by side) → EtapaListasSegundaCorrida (what the 2nd run of downloads pastes into SAP, from the
same #N/A rows) → Consolidador (the orchestrator); a class only uses the ones above it.
consolidar_gallo_monivoi() is what the controller calls; calcular_listas_segunda_corrida() gives
the '2a corrida' download button the same lists without writing a workbook; consolidar_venta_dsd() and
consolidar_venta_eiap() run when ONLY the file of that sale was downloaded (each its own workbook
with the tabs 'Gallo' + 'Venta DSD' / 'Venta EIAP', see EtapaVentas: they are NOT part of the full
workbook). No SAP and no tkinter here.
"""

import os
import re
import unicodedata
from datetime import datetime
from functools import partial

import pandas as pd
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from utils import CUENTAS_DSD, CLASES_DSD, CUENTAS_EIAP, CLASES_EIAP


# === Shared helpers ===
class Comun:
    """Stage-shared helpers: CSV reading, column lookup, amount parsing, legends, formula pieces."""

    # Label of the totals row in 'TD' AND 'MTD': the TD × MTD cross looks it up with XLOOKUP.
    ETIQUETA_GRAND_TOTAL = "Grand Total"

    FORMATO_NUMERO = "#,##0.00"

    ENCABEZADO_IMPORTE = "Importe en moneda local"

    FILTROS_COLOR = "FF595959"

    TEXTO_SIN_TIMBRE = "Sin timbre"

    # US thousands groups are 3 digits; any other comma ('1.234,56') must fail, not shrink the amount.
    _RE_IMPORTE_US_AGRUPADO = re.compile(r"\d{1,3}(?:,\d{3})+(?:\.\d+)?")

    # Only letters/'_' can go unquoted in a formula ('MTD!A:A'); R, C, TRUE and FALSE need quotes too.
    _RE_HOJA_SIN_COMILLAS = re.compile(r"[A-Za-z_]+")
    _HOJAS_RESERVADAS = frozenset({"R", "C", "TRUE", "FALSE"})

    @staticmethod
    def leer_csv(ruta):
        """Read a CSV downloaded by the app: everything as text, accents included."""
        if not os.path.exists(ruta):
            raise FileNotFoundError(f"Archivo no encontrado: {ruta}")
        return pd.read_csv(ruta, dtype=str, encoding="utf-8-sig", keep_default_na=False)

    @staticmethod
    def es_sin_datos(df):
        """True if the CSV is the 'Sin datos' placeholder written by the download."""
        return df.empty or (len(df.columns) == 1 and "Sin datos" in str(df.columns[0]))

    # Cell texts meaning 'nothing here' (besides blank), compared upper-cased.
    VALORES_VACIOS = frozenset({
        "", "NULL", "NONE", "NAN", "NA", "N/A", "#N/A", "#N/D", "-", "--", "#",
    })

    @staticmethod
    def documentos_distintos(valores):
        """Distinct document numbers of `valores` in order of appearance (the list pasted into SAP); drops
        'nothing here' cells, SAP's all-zero initial values and the '.0' of the xlsx→csv conversion."""
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
        """Real name of a column, accepting spelling variants of the same field; none found → ValueError
        listing the real columns (`origen` = file): stopping beats summing the wrong one."""
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
        """US-formatted SAP amount → float ('1,234.56', '1,234.56-', '(1,234.56)'; blank → 0.0); anything
        else, even '1.234,56', raises ValueError (a corrupt amount must stop the run)."""
        if valor is None:
            return 0.0

        s = str(valor).strip()
        if s in ("", "-", "+"):
            return 0.0

        negativo = False

        if s.startswith("(") and s.endswith(")"):
            negativo = True
            s = s[1:-1].strip()

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

        # ',' is ONLY a thousands separator: validate BEFORE stripping it, so '1.234,56' fails loud.
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
        """Write a matrix's filter legend (italic gray), one line per row from (fila_inicial, columna) down;
        callers build it from their own filter constants so it cannot drift from what is summed."""
        for i, texto in enumerate(filtros):
            celda = ws.cell(row=fila_inicial + i, column=columna, value=texto)
            celda.font = Font(italic=True, color=Comun.FILTROS_COLOR)

    AUTOR_COMENTARIOS = "Validación Factura Global"

    @staticmethod
    def comentar(celda, texto, ancho=420, alto=150):
        """Attach a note to `celda`, sized so the whole text is visible."""
        comentario = Comment(texto, Comun.AUTOR_COMENTARIOS)
        comentario.width = ancho
        comentario.height = alto
        celda.comment = comentario

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


# === Stage: Gallo (FAGLL03) — 'Referencia 2/3' + 'MX Commercial' copy + 'Validación - Doc.' ===
class EtapaGallo:
    """Everything computed FROM the Gallo: K, L ('Referencia 2/3') and 'Validación - Doc.'. Rule
    25-oct-2026: 'MX Commercial' rows → D:E; their Global ones → G:H, then 'Individual'."""

    HOJA_GALLO      = "Gallo"
    HOJA_VALIDACION = "Validación - Doc."

    VALOR_ASIGNACION       = "MX Commercial"
    PREFIJOS_LETRAS_GLOBAL = frozenset({"PP", "RG", "SG"})

    # True: K, L are LIVE template formulas; False: values. Python computes the values either way.
    REF_COMO_FORMULAS = True

    # True: the 'MX Commercial' override covers the whole document; False: only the line itself.
    MX_POR_DOCUMENTO = True

    VAL_COL_REFERENCIA = "A"
    VAL_COL_TIPO       = "B"

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
        """`gallo` (CSV as text) is MUTATED by `preparar` (helper columns). `con_ref_pago`: also resolve
        'Referencia de pago' (for 'TD.Diferencias_Referencias')."""
        self.gallo = gallo
        self._update_status = update_status or (lambda _mensaje: None)
        self.con_ref_pago = con_ref_pago
        self.resultado = None
        self._mascara_reclasificadas = None

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

    # Step 4 reads the FIRST character only, ASCII digits only (str.isdigit accepts '²').
    DIGITOS = frozenset("0123456789")

    @staticmethod
    def clasificar_tipo(referencia2):
        """Step 4: 'Global' if it STARTS with a digit (0-9) or is PP/RG/SG; 'Individual' otherwise
        ('1A' → Global, 'A1' → Individual, '' → Individual).
        """
        ref = str(referencia2).strip()
        if ref[:1] in EtapaGallo.DIGITOS:
            return "Global"
        if ref.upper() in EtapaGallo.PREFIJOS_LETRAS_GLOBAL:
            return "Global"
        return "Individual"

    @staticmethod
    def referencia2(referencias):
        """'Referencia 2' of a Series of 'Referencia': the first two characters AS THEY ARE (the template's
        =LEFT(C2,2), no trimming); shared by 'Gallo' and the FB03 list so both classify alike."""
        return referencias.astype(str).str[:2]

    @staticmethod
    def _letra_columna(df, nombre):
        """Excel letter of column `nombre` as `to_excel(index=False)` will write it. ValueError if the
        name is not there exactly once: a formula could not tell which column it means.
        """
        posiciones = [i for i, columna in enumerate(df.columns) if columna == nombre]
        if len(posiciones) != 1:
            raise ValueError(
                f"La columna '{nombre}' del Gallo aparece {len(posiciones)} veces; las fórmulas de "
                f"'Referencia 2' / 'Referencia 3' necesitan que esté exactamente una vez. "
                f"Columnas: {list(df.columns)}"
            )
        return get_column_letter(posiciones[0] + 1)

    def preparar(self):
        """Ref 2/3 → 'MX Commercial' rows (D:E) → their Global rows (G:H) → 'Individual'; ONE amount pass.
        MUTATES `self.gallo` (K, L); returns (also in resultado) columns, refs_unicas, tipos,
        catalogo_docs, globales, filas_reclasificadas, importe_num, avisos. ValueError on bad data."""
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

        update_status("🔧 Creando 'Referencia 2'...")
        referencia_texto = gallo[col_ref].astype(str)
        gallo["Referencia 2"] = self.referencia2(referencia_texto)

        avisos = []
        con_espacio = referencia_texto.str[:1].str.isspace()
        n_espacio = int(con_espacio.sum())
        if n_espacio:
            ejemplo = repr(str(referencia_texto[con_espacio].iloc[0])[:12])
            avisos.append(
                f"Gallo: {n_espacio:,} fila(s) con '{col_ref}' que EMPIEZA con espacio (ej.: "
                f"{ejemplo}). 'Referencia 2' toma los 2 primeros caracteres tal cual, igual que la "
                f"fórmula =IZQUIERDA(C2,2) (no recorta espacios): ' PP-1' queda como ' P' "
                f"('Individual') y '  1' queda en blanco; ' 12-AB' sigue siendo 'Global'."
            )

        letra_ref  = self._letra_columna(gallo, col_ref)
        letra_ref2 = self._letra_columna(gallo, "Referencia 2")

        refs_unicas = [ref for ref in pd.unique(gallo["Referencia 2"]) if str(ref).strip() != ""]

        update_status("🔁 Cruzando 'Referencia 3'...")
        tipos = [self.clasificar_tipo(ref) for ref in refs_unicas]
        mapa_tipo = dict(zip(refs_unicas, tipos))
        gallo["Referencia 3"] = (
            gallo["Referencia 2"].map(mapa_tipo).fillna("")
        )
        letra_ref3 = self._letra_columna(gallo, "Referencia 3")

        update_status("🔎 Filtrando 'MX Commercial'...")
        mascara_mx = gallo[col_asig].astype(str).str.strip() == self.VALOR_ASIGNACION
        catalogo_docs = gallo.loc[mascara_mx, [col_asig, col_doc]].copy()

        update_status("🧮 Copiando los 'Global' de 'MX Commercial'...")
        mascara_global_mx = mascara_mx & (gallo["Referencia 3"] == "Global")
        n_global_mx = int(mascara_global_mx.sum())
        if self.MX_POR_DOCUMENTO:
            documentos = gallo[col_doc].astype(str).str.strip()
            documentos_mx = set(documentos[mascara_mx]) - {""}
            mascara_global_mx = (
                (mascara_mx | documentos.isin(documentos_mx)) & (gallo["Referencia 3"] == "Global"))
        globales = gallo.loc[mascara_global_mx, [col_centro, col_asig, col_doc]].copy()

        update_status("🏷️ Reclasificando 'MX Commercial' de 'Global' a 'Individual'...")
        filas_reclasificadas = int(mascara_global_mx.sum())
        gallo.loc[mascara_global_mx, "Referencia 3"] = "Individual"
        self._mascara_reclasificadas = mascara_global_mx

        update_status("🔢 Convirtiendo 'Importe en moneda local' a número...")
        importe_num = gallo[col_importe].map(
            partial(Comun.parse_importe_us, columna=col_importe))

        n_extra = filas_reclasificadas - n_global_mx
        if n_extra:
            avisos.append(
                f"Gallo: {n_extra:,} fila(s) Global de documentos con alguna línea '{self.VALOR_ASIGNACION}' "
                f"(importe {round(float(importe_num[mascara_global_mx & ~mascara_mx].sum()), 2):,.2f}) "
                f"también pasaron a 'Individual' (el documento completo se trata como "
                f"'{self.VALOR_ASIGNACION}'); aparecen en 'Validación - Doc.' G:H con las demás."
            )

        self.resultado = {
            "col_centro":           col_centro,
            "col_ref":              col_ref,
            "col_asig":             col_asig,
            "col_doc":              col_doc,
            "col_clase":            col_clase,
            "col_fecha":            col_fecha,
            "col_cuenta":           col_cuenta,
            "col_ref_pago":         col_ref_pago,
            "letra_ref":            letra_ref,
            "letra_ref2":           letra_ref2,
            "letra_ref3":           letra_ref3,
            "refs_unicas":          refs_unicas,
            "tipos":                tipos,
            "catalogo_docs":        catalogo_docs,
            "globales":             globales,
            "filas_reclasificadas": filas_reclasificadas,
            "importe_num":          importe_num,
            "avisos":               avisos,
        }
        return self.resultado

    def para_escribir(self):
        """'Gallo' frame to write. REF_COMO_FORMULAS: K =LEFT(C2,2), L =XLOOKUP(K2,'Validación -
        Doc.'!A:A,'Validación - Doc.'!B:B,"") live (the empty-string fallback avoids #N/A on a blank
        ref); rows turned 'Individual' keep that value in L. The working frame is never touched."""
        if self.resultado is None:
            raise RuntimeError("EtapaGallo.para_escribir requiere llamar antes a preparar().")
        if not self.REF_COMO_FORMULAS:
            return self.gallo

        pasos = self.resultado
        gallo = self.gallo
        n = len(gallo)
        hoja = Comun.ref_hoja(self.HOJA_VALIDACION)
        col_a, col_b = self.VAL_COL_REFERENCIA, self.VAL_COL_TIPO
        letra_ref, letra_ref2 = pasos["letra_ref"], pasos["letra_ref2"]

        f_ref2 = pd.Series(
            [f"=LEFT({letra_ref}{fila},2)" for fila in range(2, n + 2)],
            index=gallo.index, dtype=object)
        f_ref3 = pd.Series(
            [f"=_xlfn.XLOOKUP({letra_ref2}{fila},{hoja}!{col_a}:{col_a},{hoja}!{col_b}:{col_b},\"\")"
             for fila in range(2, n + 2)],
            index=gallo.index, dtype=object)
        f_ref3 = f_ref3.mask(self._mascara_reclasificadas, "Individual")

        salida = gallo.copy(deep=False)
        salida["Referencia 2"] = f_ref2
        salida["Referencia 3"] = f_ref3
        return salida

    def escribir_validacion(self, ws):
        """'Validación - Doc.': A-B every 'Referencia 2' and its type, D-E Asignación / Nº documento of ALL
        'MX Commercial' rows (audit), G-H Centro / Nº documento of those that were 'Global' (before step
        4)."""
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

        for celda, texto in (
            ("D1", f"Filtro {col_asig} = '{self.VALOR_ASIGNACION}' sobre la hoja {self.HOJA_GALLO}: se "
                   f"copian TODOS los renglones del filtro (control para auditoría)."),
            ("G1", f"Sólo los renglones de D:E cuya 'Referencia 3' era 'Global'. Se copian ANTES de "
                   f"pasarlos a 'Individual' en la hoja {self.HOJA_GALLO}."),
        ):
            Comun.comentar(ws[celda], texto, ancho=320, alto=90)

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
    """The 'TD' sheet: sums 'Importe en moneda local' by cedis × document class in three blocks — Facturas
    (A:D), Notas de Crédito (L:O), PEP vouchers (W:X, Referencia 2 'PP'). The cross, the micro table and
    the semáforo read these layouts; `mascara_bloque` hands them each block's exact rows."""

    HOJA_TD = "TD"

    # Rows 1-3 legend (+ analysis labels), 4 title, 5 headers, 6+ data.
    FILA_FILTROS    = 1
    FILA_ENCABEZADO = 5
    FILA_INICIO     = 6
    PEP_REFERENCIA2 = "PP"
    # Block keys: clases=(class, column), excluir_ref2, tabla=Excel Table name, col_grand_total.
    FACTURAS = {
        "titulo": "FACTURAS",
        "tabla": "TblTDFacturas",
        "col_cedis": 1,
        "clases": [("C1", 2), ("C5", 3)],
        "excluir_ref2": None,
        "col_grand_total": 4,
    }
    NOTAS = {
        "titulo": "NOTAS DE CRÉDITO",
        "tabla": "TblTDNotas",
        "col_cedis": 12,
        "clases": [("C2", 13), ("C6", 14)],
        "excluir_ref2": PEP_REFERENCIA2,
        "col_grand_total": 15,
    }
    LABELS_ANALISIS = [(1, "Monitor"), (2, "Contabilidad"), (3, "Diferencias")]

    PEP = {
        "titulo": "PEP VOUCHERS",
        "tabla": "TblTDPEP",
        "col_cedis": 23,
        "col_valor": 24,
        "encabezado_valor": "C2",
    }
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
        """Per block (own cedis and filters): Facturas Global, not 'MX Commercial', C1/C5; Notas same with
        C2/C6, ref2≠'PP'; PEP Global, ref2='PP', C2. Blank-Centro rows go to 'avisos', never dropped.
        Amounts come parsed; returns per block 'cedis', 'suma', totals, 'avisos'."""
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
            return {"cedis": cedis,
                    "suma": {k: round(float(v), 2) for k, v in serie.items()}}

        facturas = _bloque(self._con_exclusion(base, self.FACTURAS, ref2_norm), self.FACTURAS["clases"])
        notas    = _bloque(self._con_exclusion(base, self.NOTAS, ref2_norm), self.NOTAS["clases"])

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
            (self.FACTURAS["titulo"], self._con_exclusion(base_sin_centro, self.FACTURAS, ref2_norm)
                                      & clase_norm.isin(clases_facturas)),
            (self.NOTAS["titulo"],    self._con_exclusion(base_sin_centro, self.NOTAS, ref2_norm)
                                      & clase_norm.isin(clases_notas)),
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
            "columnas": {
                "centro":      col_centro,
                "asignacion":  col_asig,
                "clase":       col_clase,
            },
        }
        return self.resultado

    @staticmethod
    def _con_exclusion(mascara, bloque, ref2_norm):
        """`mascara` without the rows whose 'Referencia 2' the block leaves out (`excluir_ref2`)."""
        excluir = bloque.get("excluir_ref2")
        if excluir is None:
            return mascara
        return mascara & (ref2_norm != str(excluir).strip().upper())

    def mascara_bloque(self, clases, excluir_ref2=None):
        """Mask of the Gallo rows `preparar` summed for a two-class block (Global, not MX Commercial, Centro
        set, Clase in `clases`, minus `excluir_ref2` if given); the micro table uses it. Call after
        `preparar`."""
        if self._base is None:
            raise RuntimeError("EtapaTD.mascara_bloque requiere llamar antes a preparar().")
        mascara = self._base & self.clase_norm.isin({clase for clase, _ in clases})
        return self._con_exclusion(mascara, {"excluir_ref2": excluir_ref2}, self.ref2_norm)

    def mascara_pep(self):
        """Mask of the Gallo rows of the PEP block — the SAME rows `preparar` summed (Global, Referencia
        2 == 'PP', Clase == 'C2', non-blank Centro). Call it after `preparar`.
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
        excluir = bloque.get("excluir_ref2")
        ref2 = f"; Referencia 2 ≠ '{excluir}'" if excluir is not None else ""
        return [
            f"Filtros — hoja '{EtapaGallo.HOJA_GALLO}':",
            f"Referencia 3 = 'Global'{ref2}; {columnas['centro']} no vacío",
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
            f"('{EtapaGallo.VALOR_ASIGNACION}' ya pasó a 'Individual')",
        ]

    @staticmethod
    def formula_total(col, primera, ultima):
        """Grand Total cell of column `col` over rows `primera`..`ultima`: ROUND(SUBTOTAL(9,…),2), so it
        follows the block Table's filters, skips text such as 'Sin timbre' and carries no float noise."""
        letra = get_column_letter(col)
        return f"=ROUND(SUBTOTAL(9,{letra}{primera}:{letra}{ultima}),2)"

    def escribir(self, ws_td):
        """Write the blocks: rows 1-3 legend + labels, 4 title, 5 headers, 6+ centro values, then a SUBTOTAL
        'Grand Total' row per block (label MUST stay Comun.ETIQUETA_GRAND_TOTAL). The cross (E:G, P:R,
        Y:AA) and the Tables are added later by EtapaCruceTDMTD."""
        td_resumen = self.resultado
        fila_inicio = self.FILA_INICIO
        fila_encabezado = self.FILA_ENCABEZADO

        columnas = td_resumen["columnas"]

        for bloque, datos in ((self.FACTURAS, td_resumen["facturas"]),
                              (self.NOTAS,    td_resumen["notas"])):
            cedis      = datos["cedis"]
            suma       = datos["suma"]
            n_cedis    = len(cedis)
            fila_total = fila_inicio + n_cedis
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

            for i, centro in enumerate(cedis):
                fila = fila_inicio + i
                ws_td.cell(row=fila, column=col_cedis, value=centro)
                suma_fila = 0.0
                for clase, col_idx in clases:
                    valor = round(suma.get((centro, clase), 0.0), 2)
                    ws_td.cell(row=fila, column=col_idx, value=valor).number_format = Comun.FORMATO_NUMERO
                    suma_fila += valor
                ws_td.cell(row=fila, column=col_gt,
                           value=round(suma_fila, 2)).number_format = Comun.FORMATO_NUMERO

            if n_cedis > 0:
                ultima = fila_total - 1
                ws_td.cell(row=fila_total, column=col_cedis,
                           value=Comun.ETIQUETA_GRAND_TOTAL).font = Font(bold=True)
                for col_idx in [col for _, col in clases] + [col_gt]:
                    celda = ws_td.cell(row=fila_total, column=col_idx,
                                       value=self.formula_total(col_idx, fila_inicio, ultima))
                    celda.number_format = Comun.FORMATO_NUMERO
                    celda.font = Font(bold=True)

        cedis_pep      = td_resumen["pep"]["cedis"]
        suma_pep       = td_resumen["pep"]["suma"]
        col_cedis_pep  = self.PEP["col_cedis"]
        col_valor_pep  = self.PEP["col_valor"]
        fila_total_pep = fila_inicio + len(cedis_pep)

        Comun.escribir_filtros(ws_td, self.FILA_FILTROS, col_cedis_pep,
                               self.leyenda_filtros_pep(columnas))
        for fila, texto in self.LABELS_ANALISIS:
            ws_td.cell(row=fila, column=col_valor_pep, value=texto).font = Font(bold=True)
        ws_td.cell(row=4, column=col_cedis_pep, value=self.PEP["titulo"]).font = Font(bold=True)
        ws_td.cell(row=fila_encabezado, column=col_cedis_pep, value="Centro").font = Font(bold=True)
        ws_td.cell(row=fila_encabezado, column=col_valor_pep,
                   value=self.PEP["encabezado_valor"]).font = Font(bold=True)

        for i, centro in enumerate(cedis_pep):
            fila = fila_inicio + i
            ws_td.cell(row=fila, column=col_cedis_pep, value=centro)
            valor = round(suma_pep.get(centro, 0.0), 2)
            ws_td.cell(row=fila, column=col_valor_pep, value=valor).number_format = Comun.FORMATO_NUMERO
        if cedis_pep:
            ws_td.cell(row=fila_total_pep, column=col_cedis_pep,
                       value=Comun.ETIQUETA_GRAND_TOTAL).font = Font(bold=True)
            celda_pep = ws_td.cell(
                row=fila_total_pep, column=col_valor_pep,
                value=self.formula_total(col_valor_pep, fila_inicio, fila_total_pep - 1))
            celda_pep.number_format = Comun.FORMATO_NUMERO
            celda_pep.font = Font(bold=True)

        anchos = {"A": 12, "B": 16, "C": 16, "D": 16,
                  "L": 12, "M": 16, "N": 16, "O": 16,
                  "W": 26, "X": 16}
        for letra, ancho in anchos.items():
            ws_td.column_dimensions[letra].width = ancho


# === Stage: the document list of the FB03 download (from the Gallo) ===
class EtapaDocumentosFB03:
    """FB03 document list from the run's GALLO: 'Nº documento' of rows with Clase in {C2, C6} AND a Global
    'Referencia' (digit start or PP/RG/SG; 'PP' stays: the PEP table reads FB too). Reuses EtapaGallo and
    EtapaTD.NOTAS. Deliberately NOT applied: 'MX Commercial' exclusion, non-blank Centro, override (an
    extra FB03 document is harmless, a missing one is not)."""

    CLASES = frozenset(clase for clase, _ in EtapaTD.NOTAS["clases"])

    _PREFIJOS = tuple(sorted(EtapaGallo.PREFIJOS_LETRAS_GLOBAL))
    DESCRIPCION_FILTRO = (
        f"Clase de documento {'/'.join(sorted(CLASES))} y Referencia Global"
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
        """FB03 list of a Gallo DataFrame (not modified) → (documentos, filas_leidas, nombre_columna); empty
        for the 'Sin datos' placeholder. `columnas` may come pre-resolved (see `leer`)."""
        if Comun.es_sin_datos(gallo):
            return [], 0, "Nº documento"
        col_ref, col_doc, col_clase = columnas or EtapaDocumentosFB03.columnas(gallo)

        clase = gallo[col_clase].astype(str).str.strip().str.upper()
        # Step 1 of EtapaGallo (same definition); a blank 'Referencia 2' is never Global.
        referencia2 = EtapaGallo.referencia2(gallo[col_ref])
        tipos = {
            ref: EtapaGallo.clasificar_tipo(ref)
            for ref in referencia2.unique() if str(ref).strip() != ""
        }
        es_global = referencia2.map(tipos).fillna("") == "Global"

        mascara = clase.isin(EtapaDocumentosFB03.CLASES) & es_global
        return Comun.documentos_distintos(gallo.loc[mascara, col_doc]), len(gallo), col_doc

    @staticmethod
    def leer(ruta_gallo):
        """Read the Gallo CSV and return `documentos`, loading only the 3 columns needed (resolved on the
        FULL header). FileNotFoundError: missing file; ValueError: missing column."""
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
        gallo.columns = [nombres[p] for p in posiciones]
        return EtapaDocumentosFB03.documentos(gallo, columnas)


# === Stage: 'Venta DSD' and 'Venta EIAP' — amount by date ===
class EtapaVentas:
    """Sales sheet of a sales button: date (A) × summed amount (B) from that sale's FAGLL03 file; subclasses
    set tab, filters (DSD: C1/C2/C5/C6/X0/blank + accts 4000005/10; EIAP: RV/DA + utils.CUENTAS_EIAP) and
    texts. Rows fitting no filter go to 'avisos', never summed."""

    BLOQUES = ()
    ORIGEN = ""
    LEYENDA_ORIGEN = ""

    MAX_FILAS_EXCEL = 1_048_576

    ENCABEZADO_FECHA   = "Fecha"
    ETIQUETA_SIN_FECHA = "Sin fecha"
    ETIQUETA_TOTAL     = Comun.ETIQUETA_GRAND_TOTAL
    # Escaped slashes: an unescaped "/" in an Excel format is the viewer's regional date separator.
    FORMATO_FECHA_EXCEL = r"DD\/MM\/YYYY"

    FILA_FILTROS    = 1
    FILA_ENCABEZADO = 5
    FILA_INICIO     = 6
    ETIQUETA_EN_BLANCO = "en blanco"

    # 4-digit years first; ambiguity (01/09) is settled with the GUI dates or fails, never guessed.
    FORMATOS_FECHA_4D = (
        "%d.%m.%Y", "%m.%d.%Y",
        "%d/%m/%Y", "%m/%d/%Y",
        "%Y-%m-%d", "%Y/%m/%d",
        "%d-%m-%Y", "%m-%d-%Y",
    )
    FORMATOS_FECHA_2D = ("%d.%m.%y", "%m.%d.%y", "%d/%m/%y", "%m/%d/%y")
    ANIO_MIN = 1990
    ANIO_MAX = 2100

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
        """Format of the Gallo date column: 4-digit years first. Several formats parse all → same results:
        the first; different: the one placing more dates in the GUI selection; tie → ValueError (a
        misread 01/09/2026 would move the amount to another day). None if nothing to parse."""
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
                return viables[0][0]

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

    @classmethod
    def desde_csv(cls, eiap, fechas_sel=()):
        """The stage for the sale's CSV (as text): resolves its columns by name and parses the amounts.
        ValueError: a missing column or an unreadable amount. TypeError: called on the base class.
        """
        if not cls.BLOQUES:
            raise TypeError("EtapaVentas es la base: usa EtapaVentaDSD o EtapaVentaEIAP.")

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
        """Both sheets → {'dsd': {…}, 'eiap': {…}, 'avisos'}. Rows fitting no sheet (e.g. a class SAP's
        filter should have excluded) are reported with their amount, never summed."""
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
        fecha_norm = fecha_raw.map(lambda valor: mapa_fechas.get(valor))

        cuentas_presentes = set(cuenta_norm.unique())
        seleccionadas = set(fechas_sel)
        avisos = []
        bloques = {}
        mascaras = {}

        for clave, etiqueta, clases, cuentas in self.BLOQUES:
            cuentas_buscadas = {self.normalizar_cuenta(c) for c in cuentas}
            mascara = clase_norm.isin(clases) & cuenta_norm.isin(cuentas_buscadas)
            mascaras[clave] = mascara

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

        cubiertas = pd.Series(False, index=gallo.index)
        for mascara in mascaras.values():
            cubiertas = cubiertas | mascara
        ajenas = ~cubiertas
        n_ajenas = int(ajenas.sum())
        if n_ajenas:
            importe_ajeno = round(float(importe_num[ajenas].sum()), 2)
            pares = (
                pd.DataFrame({"_clase": clase_norm[ajenas], "_cuenta": cuenta_norm[ajenas]})
                .drop_duplicates().head(3)
            )
            ejemplos = ", ".join(
                f"{clase or self.ETIQUETA_EN_BLANCO}/{cuenta}"
                for clase, cuenta in zip(pares["_clase"], pares["_cuenta"])
            )
            hojas = " ni ".join(f"'{etiqueta}'" for _, etiqueta, _, _ in self.BLOQUES)
            avisos.append(
                f"{self.ORIGEN}: {n_ajenas:,} fila(s) (importe {importe_ajeno:,.2f}) con una "
                f"combinación de clase de documento y cuenta que no es de {hojas} "
                f"(clase/cuenta, ej.: {ejemplos}); no se sumaron."
            )

        bloques["avisos"] = avisos
        self.resultado = bloques
        return bloques

    @staticmethod
    def leyenda_filtros(clases, cuentas, columnas, origen="hoja 'Gallo'"):
        """Filter legend of 'Venta DSD' from its own filter sets and real column names; spells out what
        breaks a manual check: the EMPTY class counts ('(blanks)') and accounts compare without leading
        zeros or '.0'."""
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
        """Write 'Venta DSD': rows 1-4 legend, 5 headers, one row per day (real DD/MM/YYYY dates), 'Sin
        fecha' (if any) and a bold 'Grand Total'; amounts as 2-decimal values."""
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

        if bloque["sin_fecha_n"]:
            ws.cell(row=fila, column=1, value=self.ETIQUETA_SIN_FECHA).font = Font(italic=True)
            celda = ws.cell(row=fila, column=2, value=bloque["sin_fecha_importe"])
            celda.number_format = Comun.FORMATO_NUMERO
            celda.font = Font(italic=True)
            fila += 1

        if fila > self.FILA_INICIO:
            ws.cell(row=fila, column=1, value=self.ETIQUETA_TOTAL).font = Font(bold=True)
            celda_total = ws.cell(row=fila, column=2, value=bloque["total"])
            celda_total.number_format = Comun.FORMATO_NUMERO
            celda_total.font = Font(bold=True)

        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 24


class EtapaVentaDSD(EtapaVentas):
    """'Venta DSD': classes C1/C2/C5/C6/X0 (or blank) and accounts 4000005/4000010 of ITS download."""

    HOJA_DSD = "Venta DSD"
    # The blank class counts for DSD, but SAP's class filter of the download cannot ask for it: it
    # only matters for a file that already carries blank-class rows.
    CLASES_HOJA = frozenset(CLASES_DSD) | {""}
    CUENTAS_HOJA = frozenset(CUENTAS_DSD)
    BLOQUES = (("dsd", HOJA_DSD, CLASES_HOJA, CUENTAS_HOJA),)
    ORIGEN = "Venta DSD"
    LEYENDA_ORIGEN = "archivo de Venta DSD"


class EtapaVentaEIAP(EtapaVentas):
    """'Venta EIAP': classes RV/DA and the EIAP accounts of ITS download."""

    HOJA_EIAP = "Venta EIAP"
    CLASES_HOJA = frozenset(CLASES_EIAP)
    CUENTAS_HOJA = frozenset(CUENTAS_EIAP)
    BLOQUES = (("eiap", HOJA_EIAP, CLASES_HOJA, CUENTAS_HOJA),)
    ORIGEN = "Venta EIAP"
    LEYENDA_ORIGEN = "archivo de Venta EIAP"


# === Stage: Monivoi (ZLMXCOM_TRN_MONINVOI) — 'MTD' (step 16) ===
class EtapaMonivoi:
    """The 'MTD' sheet (step 16: four Centro × 'Valor neto' matrices SG, RG, PP, TODOS), the helper column
    'Referencia 2' (A of 'Monivoi') and the VBRL invoice list ('Factura'). The cross is EtapaCruceTDMTD."""

    HOJA_MONIVOI = "Monivoi"
    HOJA_MTD     = "MTD"

    # Layout: rows 1-3 legend, 4 title, 5 headers, 6+ data (the cross reads whole columns, MTD!A:A).
    FILA_FILTROS    = 1
    FILA_TITULO     = 4
    FILA_ENCABEZADO = 5
    FILA_INICIO     = 6

    ENCABEZADO_CENTRO   = "Centro"
    ENCABEZADO_VALOR    = "Valor neto"
    ETIQUETA_TOTAL      = Comun.ETIQUETA_GRAND_TOTAL
    ETIQUETA_SIN_CENTRO = "Sin centro"

    PREFIJO_SG = "SG"
    PREFIJO_RG = "RG"
    PREFIJO_PP = "PP"
    PREFIJOS_TODOS = frozenset({PREFIJO_SG, PREFIJO_RG, PREFIJO_PP})

    BLOQUES = (
        {"clave": "sg",    "titulo": "SG",
         "prefijos": frozenset({PREFIJO_SG}),
         "col_centro": 1,  "col_valor": 2},
        {"clave": "rg",    "titulo": "RG",
         "prefijos": frozenset({PREFIJO_RG}),
         "col_centro": 4,  "col_valor": 5},
        {"clave": "pp",    "titulo": "PP",
         "prefijos": frozenset({PREFIJO_PP}),
         "col_centro": 11, "col_valor": 12},
        {"clave": "todos", "titulo": "TODOS (SG + RG + PP)",
         "prefijos": PREFIJOS_TODOS,
         "col_centro": 17, "col_valor": 18},
    )

    COL_REFERENCIA2 = "Referencia 2"

    COL_REFERENCIA_CANDIDATAS = ("Referencia", "Referencia de factura")
    COL_UUID_CANDIDATAS       = ("UUID", "UUID CFDI", "Folio fiscal", "UUID Fiscal")
    COL_CENTRO_CANDIDATAS     = ("Centro", "Centro de distribución",
                                 "Puesto de expedición")
    COL_VALOR_CANDIDATAS      = ("Valor neto", "Valor neto factura", "Importe neto")
    COL_FACTURA_CANDIDATAS = ("Factura",)

    COLUMNAS_DESCARGA = {
        "factura":   (COL_FACTURA_CANDIDATAS,       "Factura"),
    }

    # Empty-UUID sentinels; the UUID SHAPE is not validated on purpose.
    UUID_VACIOS = Comun.VALORES_VACIOS

    # 'TODOS == SG + RG + PP' tolerance: a cent covers rounding of already rounded values.
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
        """Step 16: four 'MTD' matrices (SG, RG, PP, TODOS) over the MONIVOI by the first 2 chars of
        'Referencia', non-empty UUID only, each with own centros + 'Sin centro' row. 'Valor neto' parsed
        only for rows in some matrix. Returns bloques, referencia2, filas_sin_uuid, avisos."""
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
                serie = con_centro.groupby("_centro")["_valor"].sum()
                filas = [(centro, round(float(valor), 2)) for centro, valor in serie.items()]

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
            "columnas": {
                "uuid":   col_uuid,
                "centro": col_centro,
            },
        }
        return self.resultado

    def insertar_referencia2(self):
        """Insert 'Referencia 2' as column A of 'Monivoi' (free name + warning if one exists), recording the
        final name in resultado['nombre_ref2']. Call AFTER `preparar`; no-op without data."""
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
        """Filter legend of an 'MTD' block: source sheet, 'Referencia 2' (by real name `nombre_ref2`),
        non-empty UUID, treatment of rows without centro. Short lines on purpose: block 1 only has B:C to
        overflow into."""
        prefijos = " o ".join(f"'{p}'" for p in sorted(bloque["prefijos"]))
        return [
            f"Filtros — hoja '{EtapaMonivoi.HOJA_MONIVOI}':",
            f"{nombre_ref2} = {prefijos}; {columnas['uuid']} con dato",
            f"{columnas['centro']} vacío → renglón '{EtapaMonivoi.ETIQUETA_SIN_CENTRO}'",
        ]

    @classmethod
    def fila_total(cls, datos):
        """Row of the 'Grand Total' of one MTD block (`datos` = resultado['bloques'][clave]), None if the
        block has no rows (no total is written). Computed here, beside the writer, because the TD
        semáforo points at it."""
        filas = cls.FILA_INICIO + len(datos["filas"]) + (1 if datos["sin_centro_n"] else 0)
        return filas if filas > cls.FILA_INICIO else None

    def escribir_mtd(self, ws):
        """Write the four step-16 matrices on 'MTD', each with its filter legend in rows 1-3. Call AFTER
        `insertar_referencia2` so legends cite the helper column's final name."""
        mtd = self.resultado
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

            if datos["sin_centro_n"]:
                ws.cell(row=fila, column=col_centro,
                        value=self.ETIQUETA_SIN_CENTRO).font = Font(italic=True)
                celda = ws.cell(row=fila, column=col_valor,
                                value=datos["sin_centro_importe"])
                celda.number_format = Comun.FORMATO_NUMERO
                celda.font = Font(italic=True)
                fila += 1

            # 'Grand Total' only if there was at least one data row (fila_total says where).
            if self.fila_total(datos) is not None:
                ws.cell(row=fila, column=col_centro,
                        value=self.ETIQUETA_TOTAL).font = Font(bold=True)
                celda_total = ws.cell(row=fila, column=col_valor, value=datos["total"])
                celda_total.number_format = Comun.FORMATO_NUMERO
                celda_total.font = Font(bold=True)

            ws.column_dimensions[get_column_letter(col_centro)].width = 12
            ws.column_dimensions[get_column_letter(col_valor)].width = 20

    @staticmethod
    def documentos_unicos(monivoi, columna):
        """ALL distinct values of ONE Monivoi column, first-appearance order (e.g. 'factura' → 'Factura' for
        VBRL); unfiltered, only blank-like cells and '.0' go. `columna` is required (no default to
        mis-paste). Returns (valores, filas_leidas, nombre_columna); ValueError: bad key/column."""
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
    """TD × MTD cross: LIVE formulas right of each TD block (Factura Monitor / Diferencia / Observación):
    Facturas↔SG → E:G (Dif = D+E, accounting is negative), Notas↔RG → P:R (O−P), PEP↔PP → Y:AA (X−Y);
    Grand Total = SUBTOTALs. Also each block's Table and the warnings."""

    ENCABEZADO_MONITOR     = "Factura Monitor"
    ENCABEZADO_DIFERENCIA  = "Diferencia"
    ENCABEZADO_OBSERVACION = "Observación"
    TEXTO_SIN_TIMBRE       = Comun.TEXTO_SIN_TIMBRE
    TEXTO_SIN_DIFERENCIA   = "Sin diferencia"

    # 'Observación' tolerance: up to one cent is 'Sin diferencia'.
    TOLERANCIA = 0.01

    # 'revisar_si': only this direction leaves Observación blank (accounting > monitor).
    CRUCES = (
        {
            "etiqueta":        "Facturas ↔ SG",
            "td_clave":        "facturas",
            "tabla":           EtapaTD.FACTURAS["tabla"],
            "col_cedis":       EtapaTD.FACTURAS["col_cedis"],
            "col_contab":      EtapaTD.FACTURAS["col_grand_total"],
            "col_monitor":     5,
            "col_diferencia":  6,
            "col_observacion": 7,
            "mtd_clave":       "sg",
            "operador":        "+",
            "revisar_si":      "<",
        },
        {
            "etiqueta":        "Notas de Crédito ↔ RG",
            "td_clave":        "notas",
            "tabla":           EtapaTD.NOTAS["tabla"],
            "col_cedis":       EtapaTD.NOTAS["col_cedis"],
            "col_contab":      EtapaTD.NOTAS["col_grand_total"],
            "col_monitor":     16,
            "col_diferencia":  17,
            "col_observacion": 18,
            "mtd_clave":       "rg",
            "operador":        "-",
            "revisar_si":      ">",
        },
        {
            "etiqueta":        "PEP vouchers ↔ PP",
            "td_clave":        "pep",
            "tabla":           EtapaTD.PEP["tabla"],
            "col_cedis":       EtapaTD.PEP["col_cedis"],
            "col_contab":      EtapaTD.PEP["col_valor"],
            "col_monitor":     25,
            "col_diferencia":  26,
            "col_observacion": 27,
            "mtd_clave":       "pp",
            "operador":        "-",
            "revisar_si":      ">",
        },
    )

    def __init__(self, td_resumen, mtd):
        self.td_resumen = td_resumen
        self.mtd = mtd

    @staticmethod
    def formulas(cruce, fila):
        """Row formulas: E =XLOOKUP(A,MTD!A:A,MTD!B:B,"Sin timbre"), F =IFERROR(D+E,"Sin timbre"), G
        =IF(F="Sin timbre","Sin timbre",IF(ROUND(F,2)<-0.01,"","Sin diferencia")) — 'Sin timbre' FIRST
        (Excel ranks text above numbers); `_xlfn.` prefix or Excel 2007 shows #NAME?."""
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
            f"=_xlfn.XLOOKUP({cedis},{hoja_mtd}!{col_centro}:{col_centro},"
            f"{hoja_mtd}!{col_valor}:{col_valor},{sin_timbre})"
        )
        f_diferencia = f"=IFERROR({contab}{cruce['operador']}{monitor},{sin_timbre})"
        f_observacion = (
            f"=IF({dif}={sin_timbre},{sin_timbre},"
            f"IF(ROUND({dif},2){cruce['revisar_si']}{umbral:g},\"\",{sin_diferencia}))"
        )
        return f_monitor, f_diferencia, f_observacion

    @classmethod
    def formulas_total(cls, cruce, fila_total, primera, ultima):
        """'Grand Total' cells: Factura Monitor and Diferencia are SUBTOTALs of the rows above (follow the
        Table's filters, skip 'Sin timbre'); Observación judges that Diferencia with the same formula as
        every row."""
        f_monitor = EtapaTD.formula_total(cruce["col_monitor"], primera, ultima)
        f_diferencia = EtapaTD.formula_total(cruce["col_diferencia"], primera, ultima)
        _, _, f_observacion = cls.formulas(cruce, fila_total)
        return f_monitor, f_diferencia, f_observacion

    def escribir(self, ws_td):
        """Write the cross on 'TD': header (row 5), a formula in every centro row, a SUBTOTAL 'Grand Total'
        row (none without centros) and one Excel Table per block over its full width (a sheet has ONE
        plain AutoFilter, but several Tables); panes frozen under the headers."""
        td_resumen = self.td_resumen
        negrita = Font(bold=True)
        letra = get_column_letter
        for cruce in self.CRUCES:
            columnas = (
                (cruce["col_monitor"],     self.ENCABEZADO_MONITOR),
                (cruce["col_diferencia"],  self.ENCABEZADO_DIFERENCIA),
                (cruce["col_observacion"], self.ENCABEZADO_OBSERVACION),
            )
            for col, encabezado in columnas:
                ws_td.cell(row=EtapaTD.FILA_ENCABEZADO, column=col, value=encabezado).font = negrita
                ws_td.column_dimensions[letra(col)].width = max(16, len(encabezado) + 3)

            n_cedis = len(td_resumen[cruce["td_clave"]]["cedis"])
            if n_cedis == 0:
                continue
            primera    = EtapaTD.FILA_INICIO
            fila_total = primera + n_cedis
            ultima     = fila_total - 1

            for fila in range(primera, ultima + 1):
                f_monitor, f_diferencia, f_observacion = self.formulas(cruce, fila)
                ws_td.cell(row=fila, column=cruce["col_monitor"],
                           value=f_monitor).number_format = Comun.FORMATO_NUMERO
                ws_td.cell(row=fila, column=cruce["col_diferencia"],
                           value=f_diferencia).number_format = Comun.FORMATO_NUMERO
                ws_td.cell(row=fila, column=cruce["col_observacion"], value=f_observacion)

            f_monitor, f_diferencia, f_observacion = self.formulas_total(
                cruce, fila_total, primera, ultima)
            celda_monitor = ws_td.cell(row=fila_total, column=cruce["col_monitor"], value=f_monitor)
            celda_dif     = ws_td.cell(row=fila_total, column=cruce["col_diferencia"], value=f_diferencia)
            celda_obs     = ws_td.cell(row=fila_total, column=cruce["col_observacion"], value=f_observacion)
            celda_monitor.number_format = Comun.FORMATO_NUMERO
            celda_dif.number_format     = Comun.FORMATO_NUMERO
            for celda in (celda_monitor, celda_dif, celda_obs):
                celda.font = negrita

            tabla = Table(
                displayName=cruce["tabla"],
                ref=(f"{letra(cruce['col_cedis'])}{EtapaTD.FILA_ENCABEZADO}:"
                     f"{letra(cruce['col_observacion'])}{ultima}"),
            )
            tabla.tableStyleInfo = TableStyleInfo(
                name=None, showFirstColumn=False, showLastColumn=False,
                showRowStripes=False, showColumnStripes=False)
            ws_td.add_table(tabla)

            Comun.comentar(
                ws_td.cell(row=4, column=cruce["col_cedis"]),
                "Cada bloque tiene su propio filtro, pero un filtro oculta RENGLONES COMPLETOS de la "
                "hoja: al filtrar un bloque también se ocultan los renglones de los demás que quedan "
                "en esas mismas filas. Quita el filtro antes de revisar otro bloque.",
            )
            Comun.comentar(
                ws_td.cell(row=EtapaTD.FILA_ENCABEZADO, column=cruce["col_monitor"]),
                f"Fórmula en cada renglón: busca el Centro en {EtapaMonivoi.HOJA_MTD} y regresa su "
                f"importe; sin coincidencia: '{self.TEXTO_SIN_TIMBRE}'. El 'Grand Total' de esta "
                f"columna y de Diferencia es un SUBTOTAL: suma sólo los renglones visibles (sigue "
                f"los filtros) y no cuenta los '{self.TEXTO_SIN_TIMBRE}'.",
            )

        ws_td.freeze_panes = f"A{EtapaTD.FILA_INICIO}"

    def avisos(self):
        """Warnings formulas cannot show: centros the monitor stamped with no TD row (in MTD's Grand Total,
        not TD's), and centros differing only in case (XLOOKUP ignores case, keeps the first match).
        Returns list[str]."""
        td_resumen = self.td_resumen
        mtd = self.mtd
        avisos = []
        encabezado_monitor = self.ENCABEZADO_MONITOR
        for cruce in self.CRUCES:
            etiqueta    = f"Cruce TD ({cruce['etiqueta']})"
            cedis_td    = td_resumen[cruce["td_clave"]]["cedis"]
            centros_mtd = [centro for centro, _ in mtd["bloques"][cruce["mtd_clave"]]["filas"]]

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
                        f"renglón en TD ({listado}); su importe NO entra en el "
                        f"'{Comun.ETIQUETA_GRAND_TOTAL}' de '{encabezado_monitor}' en TD (sí en el de "
                        f"{EtapaMonivoi.HOJA_MTD}): de ahí sale la diferencia entre ambos totales."
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
    """The 'VBRL' sheet (ZSE16 of the Monivoi invoices), source of 'Factura Monitor' in the micro tables.
    Lookups read VBRL!D:D (invoice) and E:E (amount) BY POSITION. SAP gives the invoice as TEXT with its
    'C' ('C270043678'), so D is trimmed text matched against text; only E becomes a number."""

    HOJA_VBRL = "VBRL"

    # 1-based D, E read by the lookups; keep in step with Descargas_SAP.VBRL (0-based).
    COL_FACTURA = 4
    COL_IMPORTE = 5

    # SAP invoice: optional 'C' + digits (some rows are the bare 10-digit document); else reported.
    _RE_FACTURA = re.compile(r"C?[0-9]+", re.IGNORECASE)

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
        """Invoice (text) and amount (number) columns, indexed as XLOOKUP reads them (first match wins).
        Returns df, sin_datos, filas, mapa ({invoice upper-cased: amount}), avisos, encabezados (of D and
        E). ValueError: fewer than 5 columns or an unreadable amount."""
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
                    "VBRL: el archivo vino sin datos; todas las fórmulas de "
                    "TD.Diferencias_Referencias que lo buscan quedarán en #N/A."
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
    """Micro version of 'TD' FACTURAS: same rows (EtapaTD.mascara_bloque) by Centro × Nº documento ×
    Referencia de pago (A:D), plus LIVE E =+XLOOKUP(C,VBRL!D:D,VBRL!E:E) and F =+D+E, no IFERROR (0 =
    match). References are looked up AS TEXT (no VALUE()); an empty one is '(blank)', else XLOOKUP would
    hit VBRL's first empty cell. `preparar` mirrors the #N/A rows (`na`)."""

    HOJA_TD_REFERENCIAS = "TD.Diferencias_Referencias"

    COL_CENTRO     = 1
    COL_DOCUMENTO  = 2
    COL_REFERENCIA = 3
    COL_IMPORTE    = 4
    COL_MONITOR    = 5
    COL_DIFERENCIA = 6

    ENCABEZADO_MONITOR    = "Factura Monitor"
    ENCABEZADO_DIFERENCIA = "Diferencia"

    TEXTO_SIN_REFERENCIA = "(blank)"

    TOLERANCIA_CUADRE = 0.01

    MAX_FILAS_EXCEL = 1_048_576

    AUTOR_COMENTARIOS = Comun.AUTOR_COMENTARIOS

    def __init__(self, gallo, pasos_gallo, etapa_td, etapa_vbrl):
        """`pasos_gallo` must come from EtapaGallo(con_ref_pago=True); `etapa_td` and `etapa_vbrl` must
        already be prepared.
        """
        self.gallo = gallo
        self.pasos = pasos_gallo
        self.etapa_td = etapa_td
        self.etapa_vbrl = etapa_vbrl
        self.resultado = None

    @classmethod
    def texto_referencia(cls, referencia):
        """What column C shows for a reference: the reference itself, or '(blank)' when it is empty."""
        return referencia if referencia != "" else cls.TEXTO_SIN_REFERENCIA

    @staticmethod
    def formulas(fila):
        """The two formulas of ONE row, e.g. row 6: E6 =+XLOOKUP(C6,VBRL!D:D,VBRL!E:E), F6 =+D6+E6. Written
        as `_xlfn.XLOOKUP` (newer than Excel 2007: otherwise #NAME?)."""
        letra = get_column_letter
        referencia = f"{letra(EtapaTDReferencias.COL_REFERENCIA)}{fila}"
        importe    = f"{letra(EtapaTDReferencias.COL_IMPORTE)}{fila}"
        monitor    = f"{letra(EtapaTDReferencias.COL_MONITOR)}{fila}"
        hoja_vbrl  = Comun.ref_hoja(EtapaVBRL.HOJA_VBRL)
        col_factura = letra(EtapaVBRL.COL_FACTURA)
        col_importe = letra(EtapaVBRL.COL_IMPORTE)

        f_monitor = (
            f"=+_xlfn.XLOOKUP({referencia},"
            f"{hoja_vbrl}!{col_factura}:{col_factura},"
            f"{hoja_vbrl}!{col_importe}:{col_importe})"
        )
        f_diferencia = f"=+{importe}+{monitor}"
        return f_monitor, f_diferencia

    @staticmethod
    def formulas_total(fila_total, ultima):
        """Grand Total of E and F: E = AGGREGATE(9,6,…) (SUM skipping errors: a plain SUM over one #N/A is
        #N/A); F = D + E, total against total, so it includes the accounting of the #N/A rows."""
        letra = get_column_letter
        col_d = letra(EtapaTDReferencias.COL_IMPORTE)
        col_e = letra(EtapaTDReferencias.COL_MONITOR)
        f_monitor = (f"=ROUND(_xlfn.AGGREGATE(9,6,{col_e}{EtapaTD.FILA_INICIO}:{col_e}{ultima}),2)")
        f_diferencia = f"=ROUND({col_d}{fila_total}+{col_e}{fila_total},2)"
        return f_monitor, f_diferencia

    def preparar(self):
        """Group the FACTURAS rows of 'TD' by Centro × Nº documento × Referencia de pago (PivotTable order);
        per row, will its formulas end in #N/A (reference not in VBRL!D:D, compared like XLOOKUP:
        trimmed, case-insensitive)? Returns filas, na, total, avisos, counters."""
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

        mapa = self.etapa_vbrl.resultado["mapa"]
        na = []
        sin_ref_n = sin_ref_importe = 0
        invalidas_n = invalidas_importe = 0
        faltan_n = faltan_importe = 0
        ejemplos_invalidas = []
        conteo_encontradas = {}
        for _, _, ref, importe in filas:
            clave = EtapaVBRL.clave_busqueda(self.texto_referencia(ref))
            if clave in mapa:
                na.append(False)
                conteo_encontradas[clave] = conteo_encontradas.get(clave, 0) + 1
                continue
            na.append(True)
            if ref == "":
                sin_ref_n += 1
                sin_ref_importe += importe
            elif not EtapaVBRL._RE_FACTURA.fullmatch(ref):
                invalidas_n += 1
                invalidas_importe += importe
                if len(ejemplos_invalidas) < 3 and ref not in ejemplos_invalidas:
                    ejemplos_invalidas.append(ref)
            else:
                faltan_n += 1
                faltan_importe += importe

        na_n = sin_ref_n + invalidas_n + faltan_n
        na_importe = round(sin_ref_importe + invalidas_importe + faltan_importe, 2)
        repetidas = {clave: n for clave, n in conteo_encontradas.items() if n > 1}
        importe_repetido = round(sum(mapa[clave] * (n - 1) for clave, n in repetidas.items()), 2)

        etiqueta = self.HOJA_TD_REFERENCIAS
        avisos = []
        if na_n:
            partes = []
            if sin_ref_n:
                partes.append(f"{sin_ref_n:,} sin Referencia de pago "
                              f"(se muestran como '{self.TEXTO_SIN_REFERENCIA}')")
            if invalidas_n:
                partes.append(
                    f"{invalidas_n:,} con una referencia que no parece factura "
                    f"(ej.: {', '.join(ejemplos_invalidas)})"
                )
            if faltan_n:
                partes.append(f"{faltan_n:,} que no están en VBRL")
            avisos.append(
                f"{etiqueta} (Facturas): {na_n:,} de {len(filas):,} renglón(es) quedarán en "
                f"#N/A (importe contable {na_importe:,.2f}): {'; '.join(partes)}."
            )
        if repetidas:
            avisos.append(
                f"{etiqueta} (Facturas): {len(repetidas):,} factura(s) del monitor aparecen en "
                f"más de un renglón ({sum(repetidas.values()):,} renglones); "
                f"'{self.ENCABEZADO_MONITOR}' repite su importe en cada uno y el "
                f"Grand Total lo cuenta varias veces (importe repetido "
                f"{importe_repetido:,.2f})."
            )

        total_td = self.etapa_td.resultado["total_facturas"]
        if abs(total - total_td) > self.TOLERANCIA_CUADRE:
            avisos.append(
                f"{etiqueta} (Facturas): el Grand Total contable ({total:,.2f}) no cuadra con "
                f"el de {EtapaTD.HOJA_TD} FACTURAS ({total_td:,.2f}); diferencia "
                f"{round(total - total_td, 2):,.2f}."
            )

        self.resultado = {
            "filas":      filas,
            "na":         na,
            "total":      total,
            "avisos":     avisos,
            "na_n":       na_n,
            "na_importe": na_importe,
            "columnas": {
                "centro":     pasos["col_centro"],
                "documento":  col_doc,
                "referencia": col_ref_pago,
            },
        }
        return self.resultado

    def escribir(self, ws):
        """Write 'TD.Diferencias_Referencias': rows 1-3 legend, 4 title, 5 headers with an AutoFilter, 6+
        data, 'Grand Total'; panes frozen under the headers."""
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

        enc_factura, enc_importe = vbrl["encabezados"]
        col_vbrl_factura = get_column_letter(EtapaVBRL.COL_FACTURA)
        col_vbrl_importe = get_column_letter(EtapaVBRL.COL_IMPORTE)
        cita_factura = f" ('{enc_factura}')" if enc_factura else ""
        cita_importe = f" ('{enc_importe}')" if enc_importe else ""
        self._comentar(
            ws.cell(row=EtapaTD.FILA_ENCABEZADO, column=self.COL_MONITOR),
            f"Fórmula en cada renglón: busca la Referencia de pago (como texto, con su 'C') en "
            f"{EtapaVBRL.HOJA_VBRL}!{col_vbrl_factura}:{col_vbrl_factura}{cita_factura} y "
            f"regresa {EtapaVBRL.HOJA_VBRL}!{col_vbrl_importe}:{col_vbrl_importe}"
            f"{cita_importe}. Sin coincidencia: #N/A (sin IFERROR, igual que en Notas de crédito "
            f"y PEP Vouchers). Una referencia vacía se muestra como "
            f"'{self.TEXTO_SIN_REFERENCIA}', como en una tabla dinámica.",
        )
        self._comentar(
            ws.cell(row=EtapaTD.FILA_ENCABEZADO, column=self.COL_DIFERENCIA),
            "Contabilidad + Monitor. Las facturas vienen en negativo en contabilidad, así que "
            "0 = cuadra; si Factura Monitor es #N/A, aquí también. En el Grand Total, Factura "
            "Monitor suma sólo lo que el monitor encontró (se saltan los #N/A) y Diferencia "
            "compara total contra total, así que incluye la contabilidad de los renglones #N/A.",
        )

        fila = EtapaTD.FILA_INICIO
        for centro, documento, referencia, importe in filas:
            ws.cell(row=fila, column=self.COL_CENTRO, value=centro)
            ws.cell(row=fila, column=self.COL_DOCUMENTO, value=documento)
            ws.cell(row=fila, column=self.COL_REFERENCIA, value=self.texto_referencia(referencia))
            ws.cell(row=fila, column=self.COL_IMPORTE, value=importe).number_format = formato
            f_monitor, f_diferencia = self.formulas(fila)
            ws.cell(row=fila, column=self.COL_MONITOR, value=f_monitor).number_format = formato
            ws.cell(row=fila, column=self.COL_DIFERENCIA, value=f_diferencia).number_format = formato
            fila += 1

        if filas:
            ultima = fila - 1
            f_total_monitor, f_total_diferencia = self.formulas_total(fila, ultima)
            ws.cell(row=fila, column=self.COL_CENTRO,
                    value=Comun.ETIQUETA_GRAND_TOTAL).font = negrita
            celdas_total = (
                ws.cell(row=fila, column=self.COL_IMPORTE, value=resultado["total"]),
                ws.cell(row=fila, column=self.COL_MONITOR, value=f_total_monitor),
                ws.cell(row=fila, column=self.COL_DIFERENCIA, value=f_total_diferencia),
            )
            for celda in celdas_total:
                celda.number_format = formato
                celda.font = negrita

            # AutoFilter over headers + data, not the total row (the stakeholder filters 'Diferencia').
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
    """The 'FB' sheet (FB03): for each credit note / PEP voucher document, the 'Clave de referencia' of its
    invoice ('C' + digits, as VBRL carries it). EtapaTDNotasPEP reads it BY POSITION: FB!C:C (document, a
    number) → FB!D:D (reference, text). A:B stay EMPTY on purpose (SAP's list margin in the stakeholder's
    workbook) so formulas copy unchanged."""

    HOJA_FB = "FB"

    COL_DOCUMENTO = 3
    COL_CLAVE = 4

    COL_DOCUMENTO_CANDIDATAS = ("Nº documento", "Nº doc.", "Número de documento", "Documento")
    COL_CLAVE_CANDIDATAS = ("Clave de referencia", "Clave ref.", "Clave referencia")

    # Excel keeps 15 significant digits: longer document numbers stay text (and are warned).
    MAX_DIGITOS = 15

    def __init__(self, fb, sin_datos):
        self.fb = fb
        self.sin_datos = sin_datos
        self.resultado = None

    @staticmethod
    def convertir_documento(valor):
        """Document cell as the formulas see it: '7075292768', '07075292768', '7075292768.0' → 7075292768
        (what VALUE() builds from a Gallo document); '' → None; non-digits or > MAX_DIGITOS digits stay
        text."""
        texto = str(valor).strip()
        if texto == "":
            return None
        if re.fullmatch(r"[0-9]+(?:\.0+)?", texto):
            digitos = texto.split(".", 1)[0].lstrip("0") or "0"
            if len(digitos) <= EtapaFB.MAX_DIGITOS:
                return int(digitos)
        return texto

    def preparar(self):
        """Document column → numbers, reference → trimmed text, indexed as XLOOKUP reads them (first match
        wins). Returns df, sin_datos, filas, mapa ({document: reference upper-cased, '' if blank}),
        avisos, encabezados. ValueError: document or reference column missing."""
        fb = self.fb
        if self.sin_datos or fb.empty:
            self.resultado = {
                "df":          fb,
                "sin_datos":   True,
                "filas":       0,
                "mapa":        {},
                "mapa_texto":  {},
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

        mapa = {}
        mapa_texto = {}
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
                mapa_texto[documento] = clave or ""
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
        df.isetitem(0, pd.Series(documentos, index=df.index, dtype=object))
        df.isetitem(1, pd.Series(claves, index=df.index, dtype=object))

        self.resultado = {
            "df":          df,
            "sin_datos":   False,
            "filas":       len(df),
            "mapa":        mapa,
            "mapa_texto":  mapa_texto,
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
    """Notas de crédito (H:M) and PEP Vouchers (O:T) tables on 'TD.Diferencias_Referencias': Centro | Nº
    documento | Importe (values, same rows as the 'TD' blocks) + LIVE formulas copied VERBATIM from the
    reference workbook: Referencia de pago =+XLOOKUP(doc,FB!C:C,FB!D:D), Factura Monitor
    =+XLOOKUP(ref,VBRL!D:D,VBRL!E:E), Diferencia =importe-monitor. No IFERROR: unmatched shows #N/A. Each
    block is an Excel Table."""

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

    NOTAS = {"clave": "notas", "titulo": "Notas de crédito", "col_inicio": 8,
             "tabla": "TblNotasCredito", "con_value": False}
    PEP   = {"clave": "pep",   "titulo": "PEP Vouchers",     "col_inicio": 15,
             "tabla": "TblPEPVouchers", "con_value": False}
    BLOQUES = (NOTAS, PEP)

    TOLERANCIA_CUADRE = 0.01

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
    def formulas(col_inicio, fila, con_value=True):
        """The three formulas of ONE row at `col_inicio`; block H:M, row 6, with `con_value`: K6
        =+_xlfn.XLOOKUP(VALUE(I6),FB!C:C,FB!D:D), L6 =+_xlfn.XLOOKUP(K6,VBRL!D:D,VBRL!E:E), M6 =J6-L6.
        Without `con_value` K6 has no VALUE() (I6 must then be a number)."""
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

        buscado = f"VALUE({documento})" if con_value else documento
        f_referencia = (
            f"=+_xlfn.XLOOKUP({buscado},"
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

    OK = "ok"
    VALOR = "valor"
    SIN_FB = "sin_fb"
    SIN_CLAVE = "sin_clave"
    SIN_VBRL = "sin_vbrl"
    ESTADOS_NA = frozenset({SIN_FB, SIN_CLAVE, SIN_VBRL})

    _RE_VALOR_NUMERICO = re.compile(r"[0-9]+(?:\.[0-9]*)?")

    @classmethod
    def resolver(cls, documento, mapa_fb, mapa_vbrl, con_value=True):
        """Outcome of ONE row in Excel → (estado, numero): estado is OK, VALOR (#VALUE!) or an ESTADOS_NA;
        numero the FB!C:C key found, else None. Empty doc → empty cell (VALUE → 0); > 15 digits → SIN_FB.
        Without `con_value` no #VALUE!: the cell holds EtapaFB.convertir_documento(doc), as FB!C:C does."""
        texto = str(documento).strip()
        if not con_value:
            clave = EtapaFB.convertir_documento(texto)
            numero = 0 if clave is None else clave
            if numero not in mapa_fb:
                return cls.SIN_FB, None
            referencia = mapa_fb[numero]
            if referencia == "":
                return cls.SIN_CLAVE, numero
            if referencia not in mapa_vbrl:
                return cls.SIN_VBRL, numero
            return cls.OK, numero
        if texto == "":
            numero = 0
        elif cls._RE_VALOR_NUMERICO.fullmatch(texto):
            entero, _, decimales = texto.partition(".")
            if decimales.strip("0"):
                return cls.SIN_FB, None
            digitos = entero.lstrip("0") or "0"
            if len(digitos) > EtapaFB.MAX_DIGITOS:
                return cls.SIN_FB, None
            numero = int(digitos)
        else:
            return cls.VALOR, None
        if numero not in mapa_fb:
            return cls.SIN_FB, None
        clave = mapa_fb[numero]
        if clave == "":
            return cls.SIN_CLAVE, numero
        if clave not in mapa_vbrl:
            return cls.SIN_VBRL, numero
        return cls.OK, numero

    def _mascara(self, bloque):
        """Mask of the Gallo rows of a block, taken from EtapaTD so it cannot drift from 'TD'."""
        td = self.etapa_td
        if bloque is self.NOTAS:
            return td.mascara_bloque(EtapaTD.NOTAS["clases"], EtapaTD.NOTAS["excluir_ref2"])
        return td.mascara_pep()

    def preparar(self):
        """Group each block's rows by Centro × Nº documento (PivotTable order) and measure how its formulas
        resolve against FB and VBRL. Returns {'bloques': {clave: {filas, total, sin_na…}}, 'avisos',
        'columnas'}."""
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

            estados = []
            conteo = {estado: 0 for estado in (self.VALOR, *self.ESTADOS_NA)}
            importes = {estado: 0.0 for estado in conteo}
            for _, doc, importe in filas:
                estado, _ = self.resolver(doc, mapa_fb, mapa_vbrl, bloque["con_value"])
                estados.append(estado)
                if estado in conteo:
                    conteo[estado] += 1
                    importes[estado] += importe
            na_n = sum(conteo[estado] for estado in self.ESTADOS_NA)
            na_importe = round(sum(importes[estado] for estado in self.ESTADOS_NA), 2)
            valor_n = conteo[self.VALOR]

            if na_n:
                partes = []
                if conteo[self.SIN_FB]:
                    partes.append(f"{conteo[self.SIN_FB]:,} que no están en FB")
                if conteo[self.SIN_CLAVE]:
                    partes.append(f"{conteo[self.SIN_CLAVE]:,} con clave de referencia vacía en FB")
                if conteo[self.SIN_VBRL]:
                    partes.append(f"{conteo[self.SIN_VBRL]:,} cuya referencia no está en VBRL")
                avisos.append(
                    f"{etiqueta} ({bloque['titulo']}): {na_n:,} de {len(filas):,} "
                    f"renglón(es) quedarán en #N/A (importe contable {na_importe:,.2f}): "
                    f"{'; '.join(partes)}."
                )
            if valor_n:
                avisos.append(
                    f"{etiqueta} ({bloque['titulo']}): {valor_n:,} renglón(es) con un Nº documento "
                    f"que no es un número quedarán en #VALUE! (importe contable "
                    f"{round(importes[self.VALOR], 2):,.2f}); no entran en las listas de la 2a corrida."
                )

            bloques[bloque["clave"]] = {
                "filas":      filas,
                "estados":    estados,
                "total":      total,
                "na_n":       na_n,
                "na_importe": na_importe,
                "valor_n":    valor_n,
            }

        td_resumen = td.resultado
        esperados = (
            (self.NOTAS, td_resumen["total_notas"], f"{EtapaTD.NOTAS['titulo']} de 'TD'"),
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
        """Write both blocks on 'TD.Diferencias_Referencias': legend (1-3), title, headers (the Table),
        data, 'Grand Total' (amount only). No rows → title and headers, no Table."""
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
                f"Fórmula en cada renglón: busca el Nº documento ("
                f"{'convertido a número con VALUE' if bloque['con_value'] else 'ya escrito como número, sin VALUE'}) en "
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
                # `con_value`: document as TEXT (as the reference), else a NUMBER (FB!C:C); empty stays empty.
                celda_doc = ws.cell(row=fila, column=inicio + self.OFF_DOCUMENTO,
                                    value=documento or None)
                if not bloque["con_value"] and documento:
                    clave_doc = EtapaFB.convertir_documento(documento)
                    celda_doc.value = clave_doc
                    if isinstance(clave_doc, int):
                        celda_doc.number_format = "0"
                ws.cell(row=fila, column=inicio + self.OFF_IMPORTE,
                        value=importe).number_format = formato
                f_ref, f_monitor, f_dif = self.formulas(inicio, fila, bloque["con_value"])
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

        for bloque in self.BLOQUES:
            ws.column_dimensions[letra(bloque["col_inicio"] - 1)].width = 4


# === Stage: 'Validación_doc' — the #N/A rows of the three tables, side by side (audit control) ===
class EtapaValidacionDoc:
    """Sheet 'Validación_doc' (NOT 'Validación - Doc.'): the three 'TD.Diferencias_Referencias' tables cut
    to their first 4 columns and to rows whose 'Diferencia' ends in #N/A (documents awaiting validation),
    as values. Facturas A:D, Notas F:I, PEP Q:T. Importe sits in D, H, S ON PURPOSE (the semáforo sums
    those columns) and there is NO totals row (a whole-column SUM would count it twice)."""

    HOJA = "Validación_doc"

    BLOQUES = (
        {"clave": "facturas", "titulo": "Facturas", "col_inicio": 1,
         "orden": ("centro", "documento", "referencia", "importe")},
        {"clave": "notas", "titulo": "Notas de crédito", "col_inicio": 6,
         "orden": ("centro", "documento", "importe", "referencia")},
        {"clave": "pep", "titulo": "PEP Vouchers", "col_inicio": 17,
         "orden": ("centro", "documento", "importe", "referencia")},
    )

    TEXTO_NO_DISPONIBLE = "#N/A"

    TOLERANCIA_CUADRE = 0.01

    def __init__(self, etapa_td_ref, etapa_notas_pep=None):
        """`etapa_td_ref` (and `etapa_notas_pep`, when there is one) must already be prepared."""
        self.etapa_td_ref = etapa_td_ref
        self.etapa_notas_pep = etapa_notas_pep
        self.resultado = None

    @classmethod
    def col_importe(cls, bloque):
        """1-based column of the 'Importe' of `bloque` (D, H, S)."""
        return bloque["col_inicio"] + bloque["orden"].index("importe")

    def preparar(self):
        """Pick the #N/A rows of each table → {'bloques': {clave: {filas, total}}, 'avisos', 'columnas'};
        `filas` are (centro, documento, referencia, importe) in source-table order. A cross-check vs the
        tables' own #N/A amounts goes to 'avisos'."""
        ref = self.etapa_td_ref.resultado
        bloques = {}
        avisos = []

        facturas = [
            (centro, documento, EtapaTDReferencias.texto_referencia(referencia), importe)
            for (centro, documento, referencia, importe), es_na in zip(ref["filas"], ref["na"])
            if es_na
        ]
        bloques["facturas"] = {"filas": facturas, "esperado": ref["na_importe"]}

        columnas = dict(ref["columnas"])
        np_ = self.etapa_notas_pep
        if np_ is not None:
            mapa_fb = np_.etapa_fb.resultado["mapa"]
            texto_fb = np_.etapa_fb.resultado["mapa_texto"]
            mapa_vbrl = np_.etapa_vbrl.resultado["mapa"]
            for clave in ("notas", "pep"):
                origen = np_.resultado["bloques"][clave]
                filas = []
                for (centro, documento, importe), estado in zip(origen["filas"], origen["estados"]):
                    if estado not in EtapaTDNotasPEP.ESTADOS_NA:
                        continue
                    if estado == EtapaTDNotasPEP.SIN_FB:
                        referencia = self.TEXTO_NO_DISPONIBLE
                    elif estado == EtapaTDNotasPEP.SIN_CLAVE:
                        referencia = 0
                    else:
                        _, numero = EtapaTDNotasPEP.resolver(documento, mapa_fb, mapa_vbrl)
                        referencia = texto_fb.get(numero) or None
                    filas.append((centro, documento, referencia, importe))
                bloques[clave] = {"filas": filas, "esperado": origen["na_importe"]}

        for clave, datos in bloques.items():
            datos["total"] = round(sum(fila[3] for fila in datos["filas"]), 2)
            if abs(datos["total"] - datos["esperado"]) > self.TOLERANCIA_CUADRE:
                avisos.append(
                    f"{self.HOJA} ({clave}): el importe de los renglones #N/A ({datos['total']:,.2f}) "
                    f"no cuadra con el que reporta su tabla ({datos['esperado']:,.2f})."
                )

        self.resultado = {"bloques": bloques, "avisos": avisos, "columnas": columnas}
        return self.resultado

    @staticmethod
    def _poner_texto(ws, fila, columna, valor):
        """Write a SAP value as DATA, never a formula: openpyxl turns any '='-leading string into a formula
        (e.g. '=HYPERLINK(...)' would run on open). The literal '#N/A' stays an error value, as in the
        source tables."""
        celda = ws.cell(row=fila, column=columna, value=valor)
        if celda.data_type == "f":
            celda.data_type = "s"
        return celda

    def escribir(self, ws):
        """Write the blocks on `ws`: row 1 bold headers, #N/A rows from row 2, panes frozen under the
        headers. A table without #N/A rows keeps its headers; only the blocks `preparar` found (no FB03 →
        just Facturas)."""
        resultado = self.resultado
        columnas = resultado["columnas"]
        negrita = Font(bold=True)
        textos = {
            "centro":     columnas["centro"],
            "documento":  columnas["documento"],
            "importe":    Comun.ENCABEZADO_IMPORTE,
        }
        for bloque in self.BLOQUES:
            datos = resultado["bloques"].get(bloque["clave"])
            if datos is None:
                continue
            inicio = bloque["col_inicio"]
            referencia_encabezado = (columnas["referencia"] if bloque["clave"] == "facturas"
                                     else EtapaTDNotasPEP.ENCABEZADO_REFERENCIA)
            encabezados = dict(textos, referencia=referencia_encabezado)

            for desplazamiento, campo in enumerate(bloque["orden"]):
                celda = ws.cell(row=1, column=inicio + desplazamiento, value=str(encabezados[campo]))
                celda.font = negrita

            Comun.comentar(
                ws.cell(row=1, column=inicio),
                f"{bloque['titulo']}: los renglones de la tabla de '"
                f"{EtapaTDReferencias.HOJA_TD_REFERENCIAS}' cuya Diferencia termina en #N/A "
                f"(sus primeras cuatro columnas, como valores).",
                ancho=320, alto=90,
            )
            Comun.comentar(
                ws.cell(row=1, column=self.col_importe(bloque)),
                "Sin renglón de totales a propósito: el semáforo de TD suma la columna completa "
                "(SUM(Validación_doc!D:D), H:H, S:S) y un total la contaría dos veces.",
                ancho=320, alto=90,
            )

            posiciones = {campo: inicio + i for i, campo in enumerate(bloque["orden"])}
            for fila, (centro, documento, referencia, importe) in enumerate(datos["filas"], start=2):
                self._poner_texto(ws, fila, posiciones["centro"], centro)
                self._poner_texto(ws, fila, posiciones["documento"], documento or None)
                self._poner_texto(ws, fila, posiciones["referencia"], referencia)
                ws.cell(row=fila, column=posiciones["importe"],
                        value=importe).number_format = Comun.FORMATO_NUMERO

            for desplazamiento, ancho in zip(range(4), (12, 16, 22, 22)):
                ws.column_dimensions[get_column_letter(inicio + desplazamiento)].width = ancho

        ws.freeze_panes = "A2"


# === Stage: the 'TD' semáforo — Monitor / Contabilidad / Diferencias, one per block ===
class EtapaSemaforoTD:
    """Traffic lights at the top of 'TD' (rows 1-3, beside 'Monitor' / 'Contabilidad' / 'Diferencias'): per
    block three LIVE ROUND(…,2) formulas that must be 0. Facturas, Grand Total row r: Monitor
    =E{r}-MTD!B{rt}; Contabilidad =D{r}-XLOOKUP('Grand Total',TD.Diferencias_Referencias!A:A,…!D:D);
    Diferencias =F{r}-SUM(Validación_doc!D:D). Notas / PEP: same with their own columns. Green at 0, red
    otherwise. Sources absent from the workbook are not referenced (Excel would ask for a file)."""

    ETIQUETA_SIN_VBRL = "N/D (sin VBRL)"
    ETIQUETA_SIN_FB03 = "N/D (sin FB03)"

    FUENTES = {
        "facturas": (EtapaTDReferencias.COL_CENTRO, EtapaTDReferencias.COL_IMPORTE),
        "notas": (EtapaTDNotasPEP.NOTAS["col_inicio"] + EtapaTDNotasPEP.OFF_CENTRO,
                  EtapaTDNotasPEP.NOTAS["col_inicio"] + EtapaTDNotasPEP.OFF_IMPORTE),
        "pep": (EtapaTDNotasPEP.PEP["col_inicio"] + EtapaTDNotasPEP.OFF_CENTRO,
                EtapaTDNotasPEP.PEP["col_inicio"] + EtapaTDNotasPEP.OFF_IMPORTE),
    }

    VERDE = ("C6EFCE", "006100")
    ROJO = ("FFC7CE", "9C0006")

    def __init__(self, td_resumen, mtd, con_referencias, con_notas_pep):
        """`td_resumen` / `mtd`: `resultado` of EtapaTD / EtapaMonivoi. `con_referencias`: the workbook has
        'TD.Diferencias_Referencias' and 'Validación_doc' (VBRL given); `con_notas_pep`: they also carry
        the Notas / PEP blocks (VBRL and FB03 given)."""
        self.td_resumen = td_resumen
        self.mtd = mtd
        self.con_referencias = bool(con_referencias)
        self.con_notas_pep = bool(con_referencias and con_notas_pep)

    def hay_fuente(self, clave):
        """True if the micro table and the 'Validación_doc' block of `clave` are in the workbook."""
        return self.con_notas_pep if clave != "facturas" else self.con_referencias

    def motivo_sin_fuente(self):
        """Why a block has no source: the text its Contabilidad / Diferencias cells show."""
        return self.ETIQUETA_SIN_FB03 if self.con_referencias else self.ETIQUETA_SIN_VBRL

    def formulas(self, cruce, fila_total):
        """The three formulas of ONE block → {'Monitor', 'Contabilidad', 'Diferencias'}; the last two are
        None when the block has no source."""
        letra = get_column_letter
        clave = cruce["td_clave"]
        r = fila_total

        bloque_mtd = EtapaMonivoi.POR_CLAVE[cruce["mtd_clave"]]
        fila_mtd = EtapaMonivoi.fila_total(self.mtd["bloques"][cruce["mtd_clave"]])
        monitor_td = f"{letra(cruce['col_monitor'])}{r}"
        if fila_mtd is None:
            # No MTD rows at all, so no total to subtract: the monitor's side is 0.
            f_monitor = f"=ROUND({monitor_td},2)"
        else:
            f_monitor = (f"=ROUND({monitor_td}-{Comun.ref_hoja(EtapaMonivoi.HOJA_MTD)}!"
                         f"{letra(bloque_mtd['col_valor'])}{fila_mtd},2)")

        if not self.hay_fuente(clave):
            return {"Monitor": f_monitor, "Contabilidad": None, "Diferencias": None}

        col_etiqueta, col_importe = self.FUENTES[clave]
        hoja_ref = Comun.ref_hoja(EtapaTDReferencias.HOJA_TD_REFERENCIAS)
        etiqueta = letra(col_etiqueta)
        importe = letra(col_importe)
        f_contabilidad = (
            f"=ROUND({letra(cruce['col_contab'])}{r}-(_xlfn.XLOOKUP("
            f"{Comun.texto_formula(Comun.ETIQUETA_GRAND_TOTAL)},"
            f"{hoja_ref}!{etiqueta}:{etiqueta},{hoja_ref}!{importe}:{importe})),2)"
        )
        bloque_validacion = next(b for b in EtapaValidacionDoc.BLOQUES if b["clave"] == clave)
        col_validacion = letra(EtapaValidacionDoc.col_importe(bloque_validacion))
        hoja_val = Comun.ref_hoja(EtapaValidacionDoc.HOJA)
        f_diferencias = (f"=ROUND({letra(cruce['col_diferencia'])}{r}-"
                         f"SUM({hoja_val}!{col_validacion}:{col_validacion}),2)")
        return {"Monitor": f_monitor, "Contabilidad": f_contabilidad, "Diferencias": f_diferencias}

    def _semaforo_visual(self, ws_td, celda):
        """Green at 0, red otherwise (errors included): a conditional format on that ONE cell."""
        direccion = celda.coordinate
        estilo = lambda colores: dict(  # noqa: E731
            fill=PatternFill(start_color=colores[0], end_color=colores[0], fill_type="solid"),
            font=Font(bold=True, color=colores[1]))
        ws_td.conditional_formatting.add(
            direccion, FormulaRule(formula=[f"IFERROR({direccion}=0,FALSE)"], **estilo(self.VERDE)))
        ws_td.conditional_formatting.add(
            direccion, FormulaRule(formula=[f"NOT(IFERROR({direccion}=0,FALSE))"], **estilo(self.ROJO)))

    def escribir(self, ws_td):
        """Fill the result cell right of each block's three labels (written by EtapaTD.escribir): formula,
        number format, traffic-light colours and a note on each label."""
        filas = {texto: fila for fila, texto in EtapaTD.LABELS_ANALISIS}
        notas = {
            "Monitor": "Factura Monitor del Grand Total de esta tabla menos el Grand Total de la matriz de "
                       f"{EtapaMonivoi.HOJA_MTD}. 0 = todos los centros del monitor tienen renglón aquí.",
            "Contabilidad": "Grand Total de esta tabla menos el de la tabla de "
                            f"{EtapaTDReferencias.HOJA_TD_REFERENCIAS}. 0 = esa tabla está completa.",
            "Diferencias": "Diferencia del Grand Total de esta tabla menos la suma del importe de los "
                           f"documentos pendientes en {EtapaValidacionDoc.HOJA}. 0 = lo pendiente "
                           "explica la diferencia.",
        }
        for cruce in EtapaCruceTDMTD.CRUCES:
            clave = cruce["td_clave"]
            n_cedis = len(self.td_resumen[clave]["cedis"])
            if n_cedis == 0:
                continue                         # no Grand Total row to judge
            fila_total = EtapaTD.FILA_INICIO + n_cedis
            col_etiqueta = cruce["col_contab"]
            col_resultado = col_etiqueta + 1
            formulas = self.formulas(cruce, fila_total)

            for texto, fila in filas.items():
                celda = ws_td.cell(row=fila, column=col_resultado)
                formula = formulas[texto]
                if formula is None:
                    celda.value = self.motivo_sin_fuente()
                    celda.font = Font(italic=True, color=Comun.FILTROS_COLOR)
                else:
                    celda.value = formula
                    celda.number_format = Comun.FORMATO_NUMERO
                    celda.font = Font(bold=True)
                    self._semaforo_visual(ws_td, celda)
                Comun.comentar(ws_td.cell(row=fila, column=col_etiqueta), notas[texto],
                               ancho=320, alto=90)


# === Stage: the lists of the 2nd run of downloads (from 'TD.Diferencias_Referencias') ===
class EtapaListasSegundaCorrida:
    """Lists the 2nd run pastes into SAP, from the #N/A rows of the THREE 'TD.Diferencias_Referencias'
    tables as the stages mirror them (never read back): `referencias` = 'Referencia de pago' (Notas /
    PEP: what their formula SHOWS) → VBRP, VBRK, 2nd VBRL; `documentos` = 'Nº documento' → BKPF. 'FP'
    rule (22-oct-2026): a reference starting 'FP' removes the row from BOTH lists. Also dropped: blanks,
    repeats, #VALUE! rows."""

    PREFIJO_EXCLUIDO = "FP"
    ORIGENES = (("facturas", "Facturas"), ("notas", "Notas de crédito"), ("pep", "PEP Vouchers"))

    COL_SOLICITANTE_CANDIDATAS = ("Solic.", "Solicitante")

    def __init__(self, etapa_td_ref, etapa_notas_pep=None):
        """`etapa_td_ref` (and `etapa_notas_pep`, when there is one) must already be prepared."""
        self.etapa_td_ref = etapa_td_ref
        self.etapa_notas_pep = etapa_notas_pep
        self.resultado = None

    @classmethod
    def excluida(cls, valor):
        """True for a 'Referencia de pago' that takes its whole row out of both lists ('FP...')."""
        return str(valor).strip().upper().startswith(cls.PREFIJO_EXCLUIDO)

    def _filas_na(self):
        """(origin key, Referencia de pago or None, Nº documento) of every #N/A row, in sheet order."""
        facturas = self.etapa_td_ref.resultado
        for (_, documento, referencia, _), es_na in zip(facturas["filas"], facturas["na"]):
            if es_na:
                # An empty reference shows '(blank)': there is nothing to copy from it.
                yield "facturas", (referencia or None), documento
        if self.etapa_notas_pep is None:
            return
        np_ = self.etapa_notas_pep
        mapa_fb = np_.etapa_fb.resultado["mapa"]
        texto_fb = np_.etapa_fb.resultado["mapa_texto"]
        mapa_vbrl = np_.etapa_vbrl.resultado["mapa"]
        for clave, _ in self.ORIGENES[1:]:
            bloque = np_.resultado["bloques"][clave]
            for (_, documento, _), estado in zip(bloque["filas"], bloque["estados"]):
                if estado not in EtapaTDNotasPEP.ESTADOS_NA:
                    continue
                referencia = None
                if estado == EtapaTDNotasPEP.SIN_VBRL:
                    _, numero = EtapaTDNotasPEP.resolver(documento, mapa_fb, mapa_vbrl)
                    referencia = texto_fb.get(numero) or None
                yield clave, referencia, documento

    def preparar(self):
        """Build both lists → {referencias, documentos, renglones_na (per table), excluidas_fp {renglones,
        referencias, documentos}, con_notas_pep, avisos}."""
        renglones_na = {clave: 0 for clave, _ in self.ORIGENES}
        referencias, documentos = [], []
        referencias_fp, documentos_todos = [], []
        renglones_fp = 0
        for origen, referencia, documento in self._filas_na():
            renglones_na[origen] += 1
            documentos_todos.append(documento)
            if referencia is not None and self.excluida(referencia):
                renglones_fp += 1
                referencias_fp.append(referencia)
                continue
            if referencia is not None:
                referencias.append(referencia)
            documentos.append(documento)

        listas = {
            "referencias": Comun.documentos_distintos(referencias),
            "documentos":  Comun.documentos_distintos(documentos),
        }
        # A document stays if ANY of its #N/A rows survives.
        sin_renglon = set(Comun.documentos_distintos(documentos_todos)) - set(listas["documentos"])
        excluidas = {
            "renglones":   renglones_fp,
            "referencias": len(Comun.documentos_distintos(referencias_fp)),
            "documentos":  len(sin_renglon),
        }

        avisos = []
        if self.etapa_notas_pep is None:
            avisos.append(
                "2a corrida: las listas salen sólo de la tabla Facturas (sin FB03 no hay "
                "'Notas de crédito' ni 'PEP Vouchers')."
            )
        self.resultado = {
            "referencias":   listas["referencias"],
            "documentos":    listas["documentos"],
            "renglones_na":  renglones_na,
            "excluidas_fp":  excluidas,
            "con_notas_pep": self.etapa_notas_pep is not None,
            "avisos":        avisos,
        }
        return self.resultado

    @classmethod
    def leer_solicitantes(cls, ruta_vbrk):
        """Distinct 'Solic.' values of the VBRK CSV (what FBL5N and KNVP paste) → (valores, filas_leidas,
        nombre_columna); empty for the 'Sin datos' placeholder. FileNotFoundError: missing file;
        ValueError: missing column."""
        vbrk = Comun.leer_csv(ruta_vbrk)
        if Comun.es_sin_datos(vbrk):
            return [], 0, cls.COL_SOLICITANTE_CANDIDATAS[0]
        columna = Comun.col_por_candidatos(
            vbrk, cls.COL_SOLICITANTE_CANDIDATAS, "solicitante", origen="VBRK")
        return Comun.documentos_distintos(vbrk[columna]), len(vbrk), columna


# === Stage: the 2nd consolidation — the six 2nd-run downloads as sheets ===
class EtapaCliente:
    """FBL5N → 'Cliente' sheet. The file is SAP's text list, not a table (dash lines, margins, a header line
    'Nº ident.fis.1' / 'Cliente' / 'Nombre 1' / 'NIF 3'): split on TAB (or '|'), separator lines dropped,
    header found by NAME and columns taken from where THAT line puts them, so margins cannot shift them."""

    # (sheet name, header matcher, required); order = sheet order (C..F).
    CAMPOS = (
        ("Nº ident.fis.1", lambda n: n.endswith("identfis1") or n == "rfc", True),
        ("Cliente",        lambda n: n == "cliente",                       True),
        ("Nombre 1",       lambda n: n in ("nombre1", "nombre"),           True),
        ("NIF 3",          lambda n: n == "nif3",                          False),
    )
    _RE_SEPARADOR = re.compile(r"[\s\-=+|_*]*")
    LINEAS_EN_ERROR = 4

    @staticmethod
    def leer_texto(ruta):
        """The text of SAP's file whatever its encoding (UTF-16 / UTF-8 with or without BOM, ANSI)."""
        if not os.path.exists(ruta):
            raise FileNotFoundError(f"Archivo no encontrado: {ruta}")
        with open(ruta, "rb") as f:
            datos = f.read()
        if datos.startswith((b"\xff\xfe", b"\xfe\xff")):
            return datos.decode("utf-16")
        if datos.startswith(b"\xef\xbb\xbf"):
            return datos[3:].decode("utf-8", errors="replace")
        if datos[:2000].count(b"\x00") > 0:
            return datos.decode("utf-16-le", errors="replace")
        try:
            return datos.decode("utf-8")
        except UnicodeDecodeError:
            return datos.decode("cp1252", errors="replace")

    @classmethod
    def celdas(cls, texto):
        """The text of the list as rows of trimmed cells (separator and blank lines dropped)."""
        lineas = [linea for linea in texto.splitlines() if linea.strip()]
        con_tab = sum("\t" in linea for linea in lineas)
        con_barra = sum("|" in linea for linea in lineas)
        if con_tab == 0 and con_barra == 0:
            return [[linea.strip()] for linea in lineas
                    if not cls._RE_SEPARADOR.fullmatch(linea)]
        delimitador = "\t" if con_tab >= con_barra else "|"
        return [
            [celda.strip() for celda in linea.split(delimitador)]
            for linea in lineas
            if not cls._RE_SEPARADOR.fullmatch(linea)
        ]

    @classmethod
    def leer(cls, ruta):
        """Customers of the FBL5N file → {'filas': [(rfc, cliente, nombre, nif3)], 'sin_datos', 'avisos'};
        `cliente` via EtapaFB.convertir_documento (same conversion as the lookups' other side).
        ValueError: no header line with the required columns."""
        texto = cls.leer_texto(ruta)
        if texto.lstrip("﻿ \t\r\n").lower().startswith("sin datos"):
            return {"filas": [], "sin_datos": True,
                    "avisos": ["Cliente (FBL5N): el archivo vino sin datos; la pestaña quedó vacía."]}
        filas = cls.celdas(texto)

        posiciones = None
        inicio = 0
        for i, fila in enumerate(filas):
            normalizadas = [Comun.normalizar_encabezado(c) for c in fila]
            encontrados = {}
            for nombre, coincide, _ in cls.CAMPOS:
                for j, n in enumerate(normalizadas):
                    if n and coincide(n) and j not in encontrados.values():
                        encontrados[nombre] = j
                        break
            if all(nombre in encontrados for nombre, _, requerido in cls.CAMPOS if requerido):
                posiciones, inicio = encontrados, i + 1
                break
        if posiciones is None:
            vista = " / ".join(" | ".join(f)[:100] for f in filas[:cls.LINEAS_EN_ERROR])
            raise ValueError(
                f"No reconocí el archivo de FBL5N '{os.path.basename(ruta)}': busqué la línea de "
                f"encabezados con 'Nº ident.fis.1', 'Cliente' y 'Nombre 1' y no está. "
                f"Primeras líneas: {vista or '(archivo vacío)'}"
            )

        def celda(fila, nombre):
            j = posiciones.get(nombre)
            return fila[j] if j is not None and j < len(fila) else ""

        encabezado_norm = {Comun.normalizar_encabezado(celda_h) for celda_h in filas[inicio - 1]}
        datos = []
        for fila in filas[inicio:]:
            cliente = celda(fila, "Cliente")
            if cliente == "":
                continue
            if {Comun.normalizar_encabezado(c) for c in fila} == encabezado_norm:
                continue
            datos.append((celda(fila, "Nº ident.fis.1"), EtapaFB.convertir_documento(cliente),
                          celda(fila, "Nombre 1"), celda(fila, "NIF 3")))
        avisos = []
        if not datos:
            avisos.append("Cliente (FBL5N): el archivo trae encabezados pero ningún cliente; "
                          "la pestaña quedó vacía.")
        return {"filas": datos, "sin_datos": False, "avisos": avisos}


class EtapaHojasSegundaCorrida:
    """Six sheets the 'Diferencias' formulas read, laid out like the template: headers in row 1, data from
    COLUMN C (A:B empty, as 'FB'): VBRK, Cliente (FBL5N: C RFC · D Cliente · E Nombre 1 · F NIF 3), VBRP,
    KNVP, VBRL (2), BKPF. Columns found by NAME; the ones formulas read are REQUIRED (missing → stop
    listing the real columns), extras go last. Customer numbers are NUMBERS, the rest trimmed text."""

    # Columns 'Diferencias' formulas read (1-based): ONE definition for the layout and the formulas.
    VBRK_FACTURA, VBRK_SOLICITANTE = 3, 5
    CLIENTE_RFC, CLIENTE_NUMERO, CLIENTE_NOMBRE = 3, 4, 5
    VBRP_FACTURA, VBRP_MOPED, VBRP_TPOS = 3, 4, 8
    KNVP_CLIENTE, KNVP_DESTINATARIO = 4, 10
    VBRL2_LISTA, VBRL2_FACTURA = 4, 6
    BKPF_DOCUMENTO, BKPF_ANULADO = 4, 10

    HOJA_VBRK, HOJA_CLIENTE, HOJA_VBRP = "VBRK", "Cliente", "VBRP"
    HOJA_KNVP, HOJA_VBRL2, HOJA_BKPF = "KNVP", "VBRL (2)", "BKPF"

    CLIENTE = "cliente"

    HOJAS = (
        {"clave": "vbrk", "hoja": HOJA_VBRK, "etiqueta": "VBRK", "columnas": (
            (VBRK_FACTURA,      ("Factura",),                 True,  None),
            (4,                 ("Pagador",),                 False, CLIENTE),
            (VBRK_SOLICITANTE,  ("Solic.", "Solicitante"),    True,  CLIENTE),
        )},
        {"clave": "fbl5n", "hoja": HOJA_CLIENTE, "etiqueta": "FBL5N", "columnas": ()},
        {"clave": "vbrp", "hoja": HOJA_VBRP, "etiqueta": "VBRP", "columnas": (
            (VBRP_FACTURA,      ("Factura",),                 True,  None),
            (VBRP_MOPED,        ("MoPed", "Motivo de pedido"), True, None),
            (5,                 ("Doc.venta",),               False, None),
            (6,                 ("Doc.modelo",),              False, None),
            (7,                 ("Ce.", "Centro"),            False, None),
            (VBRP_TPOS,         ("TPos", "Tipo de posición"), True,  None),
        )},
        {"clave": "knvp", "hoja": HOJA_KNVP, "etiqueta": "KNVP", "columnas": (
            (3,                 ("Md.", "Mandante"),          False, None),
            (KNVP_CLIENTE,      ("Cliente",),                 True,  CLIENTE),
            (5,                 ("OrgVt",),                   False, None),
            (6,                 ("CDis",),                    False, None),
            (7,                 ("Se",),                      False, None),
            (8,                 ("Func.",),                   False, None),
            (9,                 ("CInt",),                    False, None),
            (KNVP_DESTINATARIO, ("Cliente",),                 True,  CLIENTE),
            (11,                ("Acreedor",),                False, None),
            (12,                ("Nº pers.",),                False, None),
            (13,                ("Interl.",),                 False, None),
            (14,                ("DenI",),                    False, None),
            (15,                ("I",),                       False, None),
        )},
        {"clave": "vbrl2", "hoja": HOJA_VBRL2, "etiqueta": "VBRL 2a corrida", "columnas": (
            (3,                 ("Md.", "Mandante"),          False, None),
            (VBRL2_LISTA,       ("List.fact.",),              True,  None),
            (5,                 ("Pos.",),                    False, None),
            (VBRL2_FACTURA,     ("Factura",),                 True,  None),
            (7,                 ("Valor neto",),              False, None),
            (8,                 ("Importe del impuesto",),    False, None),
            (9,                 ("Cond.COM.",),               False, None),
            (10,                ("Impto.",),                  False, None),
            (11,                ("Solic.", "Solicitante"),    False, CLIENTE),
            (12,                ("Prc.",),                    False, None),
        )},
        {"clave": "bkpf", "hoja": HOJA_BKPF, "etiqueta": "BKPF", "columnas": (
            (3,                 ("Soc.", "Sociedad"),         False, None),
            (BKPF_DOCUMENTO,    ("Nº doc.", "Nº documento"),  True,  None),
            (5,                 ("FechaEntr",),               False, None),
            (6,                 ("Hora",),                    False, None),
            (7,                 ("Fe.contab.",),              False, None),
            (8,                 ("Nombre del usuario",),      False, None),
            (9,                 ("Oper.ref.",),               False, None),
            (BKPF_ANULADO,      ("Anul.con",),                True,  None),
            (11,                ("Mon.",),                    False, None),
            (12,                ("Clase doc.",),              False, None),
            (13,                ("Clave ref.",),              False, None),
        )},
    )
    CLAVES = tuple(hoja["clave"] for hoja in HOJAS)

    @staticmethod
    def clave_texto(valor):
        """A reference / document as XLOOKUP compares it (same rule as EtapaVBRL.clave_busqueda)."""
        return EtapaVBRL.clave_busqueda(valor)

    @staticmethod
    def clave_cliente(valor):
        """A customer number as XLOOKUP compares it: a number stays a number, a text is upper-cased."""
        return valor.upper() if isinstance(valor, str) else valor

    def __init__(self, rutas, update_status=None):
        """`rutas`: {clave: path} with the six keys of CLAVES (vbrk, fbl5n, vbrp, knvp, vbrl2, bkpf)."""
        faltan = [clave for clave in self.CLAVES if not rutas.get(clave)]
        if faltan:
            raise ValueError(f"Faltan archivos de la 2a corrida: {', '.join(faltan)}.")
        self.rutas = rutas
        self._update_status = update_status or (lambda _mensaje: None)
        self.resultado = None
        self.mapas = None

    @staticmethod
    def _nombres_base(columnas):
        """Each header without the suffix pandas adds to a REPEATED name ('Cliente.1' → 'Cliente'; only
        when the base name is also a column, so a real 'Importe.1' header stays as it is)."""
        existentes = {str(c) for c in columnas}
        bases = []
        for columna in columnas:
            texto = str(columna)
            coincidencia = re.fullmatch(r"(.*)\.\d+", texto)
            bases.append(coincidencia.group(1) if coincidencia and coincidencia.group(1) in existentes
                         else texto)
        return bases

    @classmethod
    def valor_celda(cls, valor, tipo):
        """What goes in a cell: customer numbers as numbers (EtapaFB.convertir_documento), anything else
        trimmed text; blank → None (an empty cell)."""
        if tipo == cls.CLIENTE:
            return EtapaFB.convertir_documento(valor)
        texto = str(valor).strip()
        return texto or None

    @classmethod
    def acomodar(cls, df, hoja):
        """The frame of a ZSE16 file laid out like the template → {'encabezados': {col: texto}, 'filas':
        [{col: valor}], 'posiciones': [cols]}. ValueError: a column a formula reads is missing."""
        columnas = list(df.columns)
        bases = cls._nombres_base(columnas)
        normalizadas = [Comun.normalizar_encabezado(b) for b in bases]
        usadas = set()
        asignadas = {}
        maxima = max(posicion for posicion, *_ in hoja["columnas"])
        for posicion, candidatos, requerida, tipo in hoja["columnas"]:
            objetivos = {Comun.normalizar_encabezado(c) for c in candidatos}
            indice = next((i for i, n in enumerate(normalizadas)
                           if i not in usadas and n in objetivos), None)
            if indice is None:
                if requerida:
                    letra = get_column_letter(posicion)
                    raise ValueError(
                        f"No encontré la columna '{candidatos[0]}' en el archivo de {hoja['etiqueta']} "
                        f"(las fórmulas de 'Diferencias' la leen de {hoja['hoja']}!{letra}:{letra}). "
                        f"Columnas disponibles: {columnas}"
                    )
                continue
            usadas.add(indice)
            asignadas[posicion] = (indice, tipo)
        # Every other column of the file goes after the template's, in the file's own order.
        for indice in range(len(columnas)):
            if indice not in usadas:
                maxima += 1
                asignadas[maxima] = (indice, None)
        posiciones = sorted(asignadas)
        encabezados = {p: bases[asignadas[p][0]] for p in posiciones}
        valores = [df.iloc[:, asignadas[p][0]].tolist() for p in posiciones]
        filas = [
            {p: cls.valor_celda(v, asignadas[p][1]) for p, v in zip(posiciones, fila)}
            for fila in zip(*valores)
        ] if posiciones else []
        return {"encabezados": encabezados, "filas": filas, "posiciones": posiciones}

    def preparar(self):
        """Read and check the six files (BEFORE the big work: a wrong layout stops the run now) and build
        the sheets' content and the lookup mirrors. Returns {clave: {hoja, encabezados, filas, posiciones,
        sin_datos}, 'avisos': [...]}. FileNotFoundError / ValueError as the readers raise them."""
        resultado = {}
        avisos = []
        for hoja in self.HOJAS:
            clave, ruta = hoja["clave"], self.rutas[hoja["clave"]]
            self._update_status(f"📄 Leyendo {hoja['etiqueta']} (2a corrida)...")
            if clave == "fbl5n":
                leido = EtapaCliente.leer(ruta)
                avisos.extend(leido["avisos"])
                encabezados = {p: nombre for p, (nombre, _, _) in
                               zip((3, 4, 5, 6), EtapaCliente.CAMPOS)}
                filas = [
                    {3: rfc or None, 4: cliente, 5: nombre or None, 6: nif3 or None}
                    for rfc, cliente, nombre, nif3 in leido["filas"]
                ]
                resultado[clave] = {"hoja": hoja["hoja"], "encabezados": encabezados, "filas": filas,
                                    "posiciones": [3, 4, 5, 6], "sin_datos": leido["sin_datos"]}
                continue
            df = Comun.leer_csv(ruta)
            if Comun.es_sin_datos(df):
                avisos.append(f"{hoja['hoja']}: el archivo de {hoja['etiqueta']} vino sin datos; "
                              f"la pestaña quedó vacía y sus búsquedas darán #N/A.")
                resultado[clave] = {"hoja": hoja["hoja"], "encabezados": {}, "filas": [],
                                    "posiciones": [], "sin_datos": True}
                continue
            resultado[clave] = dict(self.acomodar(df, hoja), hoja=hoja["hoja"], sin_datos=False)
        resultado["avisos"] = avisos
        self.resultado = resultado
        self.mapas = self._construir_mapas(resultado)
        return resultado

    def _construir_mapas(self, resultado):
        """{sheet: {key: value}} the way XLOOKUP reads each pair of columns (the FIRST row of a key wins),
        for the Python mirror of 'Diferencias'. A blank result is kept as None (Excel shows it as 0)."""
        def mapa(clave, col_clave, col_valor, normalizar):
            indice = {}
            for fila in resultado[clave]["filas"]:
                llave = normalizar(fila.get(col_clave))
                if llave is None or llave in indice:
                    continue
                indice[llave] = fila.get(col_valor)
            return indice

        texto, cliente = self.clave_texto, self.clave_cliente
        cliente_fila = {}
        for fila in resultado["fbl5n"]["filas"]:
            llave = cliente(fila.get(self.CLIENTE_NUMERO))
            if llave is not None and llave not in cliente_fila:
                cliente_fila[llave] = (fila.get(self.CLIENTE_RFC), fila.get(self.CLIENTE_NOMBRE))
        return {
            "vbrk":    mapa("vbrk", self.VBRK_FACTURA, self.VBRK_SOLICITANTE, texto),
            "cliente": cliente_fila,
            "vbrp":    {
                "moped": mapa("vbrp", self.VBRP_FACTURA, self.VBRP_MOPED, texto),
                "tpos":  mapa("vbrp", self.VBRP_FACTURA, self.VBRP_TPOS, texto),
            },
            "knvp":    mapa("knvp", self.KNVP_CLIENTE, self.KNVP_DESTINATARIO, cliente),
            "vbrl2":   mapa("vbrl2", self.VBRL2_FACTURA, self.VBRL2_LISTA, texto),
            "bkpf":    mapa("bkpf", self.BKPF_DOCUMENTO, self.BKPF_ANULADO, texto),
        }

    @staticmethod
    def poner(ws, fila, columna, valor):
        """Write a value copied from SAP as DATA: openpyxl turns a string that starts with '=' into a
        formula (it would RUN when the stakeholder opens the workbook) and one like '#N/A' into an error
        cell; here both stay text."""
        celda = ws.cell(row=fila, column=columna, value=valor)
        if celda.data_type in ("f", "e"):
            celda.data_type = "s"
        return celda

    def escribir(self, libro):
        """Create the six sheets in the workbook, in the template's order (header row 1, data from row 2,
        A:B empty). Call it AFTER 'Diferencias' (the stakeholder's analysis tab comes first)."""
        negrita = Font(bold=True)
        for hoja in self.HOJAS:
            datos = self.resultado[hoja["clave"]]
            ws = libro.create_sheet(hoja["hoja"])
            for posicion, texto in datos["encabezados"].items():
                ws.cell(row=1, column=posicion, value=str(texto)).font = negrita
            for numero, fila in enumerate(datos["filas"], start=2):
                for posicion, valor in fila.items():
                    if valor is not None:
                        self.poner(ws, numero, posicion, valor)
            if datos["posiciones"]:
                Comun.comentar(
                    ws.cell(row=1, column=datos["posiciones"][0]),
                    f"Hoja {hoja['hoja']} ({hoja['etiqueta']}, 2a corrida). Las columnas A:B quedan "
                    f"vacías a propósito: los datos empiezan en la C porque las fórmulas de "
                    f"'{EtapaDiferencias.HOJA}' leen estas hojas por letra de columna, igual que el "
                    f"Excel de referencia.",
                    ancho=360, alto=100,
                )
            for posicion in datos["posiciones"]:
                ws.column_dimensions[get_column_letter(posicion)].width = 16
            ws.freeze_panes = "A2"


# === Stage: 'Diferencias' — the stakeholder's final analysis tab ===
class EtapaDiferencias:
    """Sheet 'Diferencias': one row per document awaiting validation (same rows as 'Validación_doc'); A:D
    values, F 'Tipo' (Factura / NC / PV), G:Q LIVE template formulas (XLOOKUPs into VBRK, Cliente, VBRP,
    KNVP, 'VBRL (2)', BKPF; no IFERROR). Two user-approved simplifications: J tests the first 2
    characters as TEXT; P/Q read a blank 'Anul.con' (empty or padded) as not reversed. `preparar` mirrors
    every formula."""

    HOJA = "Diferencias"
    FILA_ENCABEZADO = 2
    FILA_INICIO = 3

    TIPOS = (("facturas", "Factura"), ("notas", "NC"), ("pep", "PV"))

    ENCABEZADOS = (
        (1, "CL"), (2, "Nº documento"), (3, "Referencia de pago"), (4, "Sum of Importe en moneda local"),
        (5, "Monto"), (6, "Tipo"), (7, "No cliente"), (8, "Nombre cliente"), (9, "RFC"),
        (10, "Observación"), (11, "Motivo de Pedido"), (12, "Clase de documento"),
        (13, "Destinatario Factura"), (14, "Lista de Factura"), (15, "Estatus Lista de factura"),
        (16, "Documento Anulado"), (17, "Estatus Documento anulado"),
    )
    ANCHOS = {"A": 8, "B": 14, "C": 18, "D": 16, "E": 13, "F": 9, "G": 12, "H": 32, "I": 18,
              "J": 20, "K": 12, "L": 12, "M": 20, "N": 14, "O": 14, "P": 16, "Q": 24}
    ULTIMA_COLUMNA = 17

    RFC_GLOBAL = "XAXX010101000"
    PREFIJOS_VENDEDOR = ("20", "21")
    TEXTO_GLOBAL = "Cliente Global"
    TEXTO_VENDEDOR = "Vendedor"
    TEXTO_INDIVIDUAL = "Cliente Individual"
    TEXTO_SIN_DATO = "N/D"
    NA = "#N/A"

    def __init__(self, etapa_validacion, etapa_hojas):
        """Both stages must already be prepared."""
        self.etapa_validacion = etapa_validacion
        self.etapa_hojas = etapa_hojas
        self.resultado = None

    @staticmethod
    def formulas(fila):
        """{column: formula} of the formula columns of ONE row (see the class docstring)."""
        H = EtapaHojasSegundaCorrida
        letra = get_column_letter
        hoja = Comun.ref_hoja

        def hoja_col(nombre, col):
            l = letra(col)
            return f"{hoja(nombre)}!{l}:{l}"

        vbrk_f, vbrk_s = hoja_col(H.HOJA_VBRK, H.VBRK_FACTURA), hoja_col(H.HOJA_VBRK, H.VBRK_SOLICITANTE)
        cli_n, cli_nom, cli_rfc = (hoja_col(H.HOJA_CLIENTE, H.CLIENTE_NUMERO),
                                   hoja_col(H.HOJA_CLIENTE, H.CLIENTE_NOMBRE),
                                   hoja_col(H.HOJA_CLIENTE, H.CLIENTE_RFC))
        vbrp_f, vbrp_m, vbrp_t = (hoja_col(H.HOJA_VBRP, H.VBRP_FACTURA), hoja_col(H.HOJA_VBRP, H.VBRP_MOPED),
                                  hoja_col(H.HOJA_VBRP, H.VBRP_TPOS))
        knvp_c, knvp_d = hoja_col(H.HOJA_KNVP, H.KNVP_CLIENTE), hoja_col(H.HOJA_KNVP, H.KNVP_DESTINATARIO)
        vbrl_f, vbrl_l = hoja_col(H.HOJA_VBRL2, H.VBRL2_FACTURA), hoja_col(H.HOJA_VBRL2, H.VBRL2_LISTA)
        # $-anchored like the template's BKPF lookups.
        bkpf_d = f"{hoja(H.HOJA_BKPF)}!${letra(H.BKPF_DOCUMENTO)}:${letra(H.BKPF_DOCUMENTO)}"
        bkpf_a = f"{hoja(H.HOJA_BKPF)}!${letra(H.BKPF_ANULADO)}:${letra(H.BKPF_ANULADO)}"

        r = fila
        sin_dato = Comun.texto_formula(EtapaDiferencias.TEXTO_SIN_DATO)
        anulado = f"TRIM(_xlfn.XLOOKUP(B{r},{bkpf_d},{bkpf_a})&\"\")"
        return {
            5:  f"=+ABS(D{r})",
            7:  f"=+_xlfn.XLOOKUP(C{r},{vbrk_f},{vbrk_s})",
            8:  f"=+_xlfn.XLOOKUP(G{r},{cli_n},{cli_nom})",
            9:  f"=+TRIM(_xlfn.XLOOKUP(G{r},{cli_n},{cli_rfc}))",
            10: (f"=IF(I{r}={Comun.texto_formula(EtapaDiferencias.RFC_GLOBAL)},"
                 f"{Comun.texto_formula(EtapaDiferencias.TEXTO_GLOBAL)},"
                 f"IF(OR(LEFT(G{r},2)=\"20\",LEFT(G{r},2)=\"21\"),"
                 f"{Comun.texto_formula(EtapaDiferencias.TEXTO_VENDEDOR)},"
                 f"{Comun.texto_formula(EtapaDiferencias.TEXTO_INDIVIDUAL)}))"),
            11: f"=+_xlfn.XLOOKUP(C{r},{vbrp_f},{vbrp_m})",
            12: f"=+_xlfn.XLOOKUP(C{r},{vbrp_f},{vbrp_t})",
            13: f"=+_xlfn.XLOOKUP(G{r},{knvp_c},{knvp_d})",
            14: f"=+_xlfn.XLOOKUP(C{r},{vbrl_f},{vbrl_l},{sin_dato})",
            15: f"=+IF(N{r}={sin_dato},\"No\",\"Si\")",
            16: f"=IF({anulado}=\"\",{sin_dato},{anulado})",
            17: f"=+IF(P{r}={sin_dato},\"No\",\"Si\")",
        }

    def evaluar(self, referencia, documento):
        """Python mirror: what each formula of one row shows once Excel calculates → {'G'…'Q': value or NA}.
        `referencia`: what column C holds ('#N/A' error cell, number or text); `documento`: column B."""
        H = EtapaHojasSegundaCorrida
        mapas = self.etapa_hojas.mapas
        NA = self.NA
        # An error cell in C poisons every lookup; a number never matches a text key.
        ref_error = referencia == NA
        ref = H.clave_texto(referencia) if isinstance(referencia, str) and not ref_error else None

        out = {}
        if ref_error:
            g = NA
        elif ref in mapas["vbrk"]:
            g = mapas["vbrk"][ref]
            g = 0 if g is None else g
        else:
            g = NA
        out["G"] = g

        persona = mapas["cliente"].get(H.clave_cliente(g)) if g != NA else None
        if persona is None:
            out["H"] = out["I"] = out["J"] = NA
        else:
            rfc, nombre = persona
            out["H"] = 0 if nombre is None else nombre
            out["I"] = (rfc or "").strip()
            if out["I"].upper() == self.RFC_GLOBAL.upper():
                out["J"] = self.TEXTO_GLOBAL
            elif str(g)[:2] in self.PREFIJOS_VENDEDOR:
                out["J"] = self.TEXTO_VENDEDOR
            else:
                out["J"] = self.TEXTO_INDIVIDUAL

        for letra, clave in (("K", "moped"), ("L", "tpos")):
            if ref is not None and ref in mapas["vbrp"][clave]:
                valor = mapas["vbrp"][clave][ref]
                out[letra] = 0 if valor is None else valor
            else:
                out[letra] = NA

        if g != NA and H.clave_cliente(g) in mapas["knvp"]:
            valor = mapas["knvp"][H.clave_cliente(g)]
            out["M"] = 0 if valor is None else valor
        else:
            out["M"] = NA

        if ref_error:
            out["N"] = out["O"] = NA
        else:
            if ref is not None and ref in mapas["vbrl2"]:
                valor = mapas["vbrl2"][ref]
                out["N"] = 0 if valor is None else valor
            else:
                out["N"] = self.TEXTO_SIN_DATO
            out["O"] = "No" if out["N"] == self.TEXTO_SIN_DATO else "Si"

        doc = H.clave_texto(documento)
        if doc is not None and doc in mapas["bkpf"]:
            texto = (mapas["bkpf"][doc] or "").strip()
            out["P"] = texto if texto else self.TEXTO_SIN_DATO
            out["Q"] = "No" if out["P"] == self.TEXTO_SIN_DATO else "Si"
        else:
            out["P"] = out["Q"] = NA
        return out

    def preparar(self):
        """The rows of the sheet and how their formulas will end → {'filas': [(tipo, centro, documento,
        referencia, importe)], 'esperado': [mirror per row], 'conteo': {...}, 'avisos': [...]}."""
        validacion = self.etapa_validacion.resultado["bloques"]
        filas = []
        for clave, tipo in self.TIPOS:
            bloque = validacion.get(clave)
            if bloque is None:
                continue
            for centro, documento, referencia, importe in bloque["filas"]:
                filas.append((tipo, centro, documento, referencia, importe))
        esperado = [self.evaluar(referencia, documento) for _, _, documento, referencia, _ in filas]

        n = len(filas)
        NA = self.NA
        cuenta = lambda condicion: sum(1 for e in esperado if condicion(e))
        sin_vbrk = cuenta(lambda e: e["G"] == NA)
        conteo = {
            "filas":         n,
            "por_tipo":      {tipo: sum(1 for f in filas if f[0] == tipo) for _, tipo in self.TIPOS},
            "sin_vbrk":      sin_vbrk,
            "sin_cliente":   cuenta(lambda e: e["G"] != NA and e["H"] == NA),
            "sin_vbrp":      cuenta(lambda e: e["K"] == NA),
            "sin_knvp":      cuenta(lambda e: e["G"] != NA and e["M"] == NA),
            "sin_lista":     cuenta(lambda e: e["N"] == self.TEXTO_SIN_DATO),
            "sin_bkpf":      cuenta(lambda e: e["P"] == NA),
            "anuladas":      cuenta(lambda e: e["Q"] == "Si"),
            "cliente_global":     cuenta(lambda e: e["J"] == self.TEXTO_GLOBAL),
            "vendedor":           cuenta(lambda e: e["J"] == self.TEXTO_VENDEDOR),
            "cliente_individual": cuenta(lambda e: e["J"] == self.TEXTO_INDIVIDUAL),
            "con_fp":        sum(1 for f in filas
                                 if isinstance(f[3], str) and EtapaListasSegundaCorrida.excluida(f[3])),
            "importe":       round(sum(f[4] for f in filas), 2),
        }

        etiqueta = self.HOJA
        avisos = []
        if n == 0:
            avisos.append(
                f"{etiqueta}: no hay renglones en #N/A en 'TD.Diferencias_Referencias' (o no hay FB03 "
                f"y las Facturas están completas); la pestaña quedó sólo con encabezados."
            )
        else:
            if sin_vbrk:
                extra = ""
                if conteo["con_fp"]:
                    extra = (f" ({conteo['con_fp']:,} empiezan con 'FP': la 2a corrida no los pidió)")
                aviso = (f"{etiqueta}: {sin_vbrk:,} de {n:,} renglón(es) con una Referencia de pago que "
                         f"no está en VBRK{extra}; su No cliente y todo lo que sale de él quedan en #N/A.")
                if sin_vbrk == n:
                    aviso += (" Son TODOS: revisa que la columna 'Factura' de VBRK traiga la misma "
                              "referencia ('C' + dígitos) que la 'Referencia de pago'.")
                avisos.append(aviso)
            if conteo["sin_cliente"]:
                avisos.append(f"{etiqueta}: {conteo['sin_cliente']:,} renglón(es) cuyo cliente no está en "
                              f"'Cliente' (FBL5N); Nombre, RFC y Observación quedan en #N/A.")
            if conteo["sin_knvp"]:
                avisos.append(f"{etiqueta}: {conteo['sin_knvp']:,} renglón(es) cuyo cliente no está en KNVP; "
                              f"Destinatario Factura queda en #N/A.")
            if conteo["sin_vbrp"] and conteo["sin_vbrp"] != sin_vbrk:
                avisos.append(f"{etiqueta}: {conteo['sin_vbrp']:,} renglón(es) con una referencia que no "
                              f"está en VBRP; Motivo de Pedido y Clase de documento quedan en #N/A.")
            if conteo["sin_bkpf"]:
                avisos.append(f"{etiqueta}: {conteo['sin_bkpf']:,} renglón(es) con un Nº documento que no "
                              f"está en BKPF; Documento Anulado queda en #N/A.")

        self.resultado = {"filas": filas, "esperado": esperado, "conteo": conteo, "avisos": avisos}
        return self.resultado

    def escribir(self, ws):
        """Write the sheet like the template: row 2 headers (with an AutoFilter), data from row 3, A:D as
        values and E, G:Q as formulas; panes frozen under the headers."""
        resultado = self.resultado
        negrita = Font(bold=True)
        formato = Comun.FORMATO_NUMERO
        for columna, texto in self.ENCABEZADOS:
            ws.cell(row=self.FILA_ENCABEZADO, column=columna, value=texto).font = negrita

        Comun.comentar(
            ws.cell(row=self.FILA_ENCABEZADO, column=6),
            "Factura = renglón de la tabla Facturas de 'TD.Diferencias_Referencias'; NC = de Notas de "
            "crédito; PV = de PEP Vouchers. Son los renglones cuya Diferencia terminó en #N/A (los "
            "mismos de 'Validación_doc').",
            ancho=360, alto=100,
        )
        Comun.comentar(
            ws.cell(row=self.FILA_ENCABEZADO, column=10),
            "Cliente Global si el RFC es XAXX010101000; Vendedor si el No cliente empieza con 20 o 21; "
            "si no, Cliente Individual. Misma regla del Excel de referencia con una fórmula más corta.",
            ancho=360, alto=90,
        )
        Comun.comentar(
            ws.cell(row=self.FILA_ENCABEZADO, column=16),
            "Documento de anulación (Anul.con) del Nº documento en BKPF; 'N/D' si no tiene (vacío o sólo "
            "espacios). Si el documento no está en BKPF: #N/A.",
            ancho=360, alto=90,
        )

        fila = self.FILA_INICIO
        for tipo, centro, documento, referencia, importe in resultado["filas"]:
            EtapaHojasSegundaCorrida.poner(ws, fila, 1, centro or None)
            EtapaHojasSegundaCorrida.poner(ws, fila, 2, documento or None)
            # '#N/A' (the document is not in FB) stays the error cell the other tables show.
            celda = ws.cell(row=fila, column=3, value=referencia)
            if celda.data_type == "f":
                celda.data_type = "s"
            ws.cell(row=fila, column=4, value=importe).number_format = formato
            ws.cell(row=fila, column=6, value=tipo)
            for columna, formula in self.formulas(fila).items():
                celda = ws.cell(row=fila, column=columna, value=formula)
                if columna == 5:
                    celda.number_format = formato
            fila += 1

        ultima = max(fila - 1, self.FILA_ENCABEZADO)
        ws.auto_filter.ref = (f"A{self.FILA_ENCABEZADO}:"
                              f"{get_column_letter(self.ULTIMA_COLUMNA)}{ultima}")
        ws.freeze_panes = f"A{self.FILA_INICIO}"
        for letra, ancho in self.ANCHOS.items():
            ws.column_dimensions[letra].width = ancho


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


def leer_solicitantes_vbrk(ruta_vbrk):
    """The 'Solic.' values (distinct, in order) of the VBRK download that FBL5N and KNVP paste into SAP
    → (valores, filas_leidas, nombre_columna).
    """
    return EtapaListasSegundaCorrida.leer_solicitantes(ruta_vbrk)


def preparar_tablas_diferencias(gallo, pasos_gallo, etapa_td, etapa_vbrl, etapa_fb=None,
                                update_status=None):
    """The three 'TD.Diferencias_Referencias' tables and the 2nd-run lists from their #N/A rows →
    (etapa_td_ref, etapa_notas_pep or None, listas). Stages must be prepared; `pasos_gallo` from
    EtapaGallo(con_ref_pago=True). ONE path for the consolidation and calcular_listas_segunda_corrida."""
    status = update_status or (lambda _mensaje: None)
    status("🔬 Construyendo 'TD.Diferencias_Referencias'...")
    etapa_td_ref = EtapaTDReferencias(gallo, pasos_gallo, etapa_td, etapa_vbrl)
    etapa_td_ref.preparar()
    etapa_notas_pep = None
    if etapa_fb is not None:
        status("🔬 Construyendo 'Notas de crédito' y 'PEP Vouchers'...")
        etapa_notas_pep = EtapaTDNotasPEP(gallo, pasos_gallo, etapa_td, etapa_vbrl, etapa_fb)
        etapa_notas_pep.preparar()
    listas = EtapaListasSegundaCorrida(etapa_td_ref, etapa_notas_pep).preparar()
    return etapa_td_ref, etapa_notas_pep, listas


# === Orchestrator ===
class Consolidador:
    """Runs the consolidation: reads the CSVs, runs the stages (EtapaGallo → EtapaTD → EtapaMonivoi →
    EtapaCruceTDMTD → [EtapaVBRL → … → EtapaListasSegundaCorrida]) and writes the workbook. VBRL and FB03
    are OPTIONAL (old dates must still consolidate; Notas / PEP tables need BOTH). The 2nd consolidation
    (`rutas_segunda`) is the SAME run plus six sheets and 'Diferencias': one code path."""

    def __init__(self, ruta_gallo, ruta_monivoi, ruta_output, nombre_salida=None,
                 callback_status=None, ruta_vbrl=None, ruta_fb03=None, rutas_segunda=None):
        """`ruta_vbrl` adds 'VBRL', 'TD.Diferencias_Referencias', 'Validación_doc'; `ruta_fb03` (with it)
        adds 'FB' and the Notas / PEP tables; `nombre_salida` defaults to 'Consolidado_<Gallo>.xlsx';
        `rutas_segunda` = 2nd consolidation (needs `ruta_vbrl`)."""
        self.ruta_gallo = ruta_gallo
        self.ruta_monivoi = ruta_monivoi
        self.ruta_output = ruta_output
        self.nombre_salida = nombre_salida
        self.callback_status = callback_status
        self.ruta_vbrl = ruta_vbrl or None
        self.ruta_fb03 = ruta_fb03 or None
        self.rutas_segunda = dict(rutas_segunda) if rutas_segunda else None
        if self.rutas_segunda is not None and self.ruta_vbrl is None:
            raise ValueError(
                "La 2a consolidación necesita el archivo de VBRL de la 1a corrida: los renglones de "
                "'Diferencias' salen de las tablas que ese archivo alimenta."
            )

    def _status(self, mensaje):
        """Report progress to the caller (the GUI), if it asked for it."""
        if self.callback_status:
            self.callback_status(mensaje)

    def _escribir_libro(self, ruta, etapa_gallo, monivoi, etapa_vbrl, etapa_fb, etapa_td, cruce,
                        etapa_td_ref, etapa_notas_pep, etapa_monivoi, etapa_validacion,
                        etapa_diferencias, etapa_hojas):
        """Write every sheet to `ruta`. The tab order of the workbook is the order of the writes."""
        with pd.ExcelWriter(ruta, engine="openpyxl") as writer:

            etapa_gallo.para_escribir().to_excel(
                writer, sheet_name=EtapaGallo.HOJA_GALLO, index=False)
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
            EtapaSemaforoTD(
                etapa_td.resultado, etapa_monivoi.resultado,
                con_referencias=etapa_td_ref is not None,
                con_notas_pep=etapa_notas_pep is not None,
            ).escribir(ws_td)
            # openpyxl saves formulas WITHOUT results: Excel recalculates on open.
            writer.book.calculation.fullCalcOnLoad = True

            if etapa_td_ref is not None:
                ws_referencias = writer.book.create_sheet(EtapaTDReferencias.HOJA_TD_REFERENCIAS)
                etapa_td_ref.escribir(ws_referencias)
                if etapa_notas_pep is not None:
                    etapa_notas_pep.escribir(ws_referencias)

            etapa_monivoi.escribir_mtd(writer.book.create_sheet(EtapaMonivoi.HOJA_MTD))

            if etapa_validacion is not None:
                etapa_validacion.escribir(writer.book.create_sheet(EtapaValidacionDoc.HOJA))

            if etapa_diferencias is not None:
                ws_dif = writer.book.create_sheet(EtapaDiferencias.HOJA)
                etapa_diferencias.escribir(ws_dif)
                etapa_hojas.escribir(writer.book)
                # Opens on 'Diferencias', ONLY that tab selected (several would be edited as a group).
                for ws in writer.book.worksheets:
                    ws.sheet_view.tabSelected = ws is ws_dif
                writer.book.active = writer.book.worksheets.index(ws_dif)

    def ejecutar(self):
        """Run the whole consolidation → GUI summary dict. FileNotFoundError: missing CSV. ValueError:
        fixable input (no data, missing column, bad amount/date/GUI date, VBRL / FB03 lacking columns)."""
        update_status = self._status
        con_vbrl = self.ruta_vbrl is not None
        con_fb = self.ruta_fb03 is not None
        con_notas_pep = con_vbrl and con_fb

        # Steps: read+check files → Gallo → TD → MTD → cross → VBRL/FB → lists → sheets → atomic save.
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

        etapa_vbrl = None
        if con_vbrl:
            update_status("📄 Leyendo archivo de VBRL...")
            vbrl = Comun.leer_csv(self.ruta_vbrl)
            etapa_vbrl = EtapaVBRL(vbrl, Comun.es_sin_datos(vbrl))
            etapa_vbrl.preparar()

        etapa_hojas = None
        if self.rutas_segunda is not None:
            etapa_hojas = EtapaHojasSegundaCorrida(self.rutas_segunda, update_status)
            etapa_hojas.preparar()

        etapa_fb = None
        if con_notas_pep:
            update_status("📄 Leyendo archivo de FB03...")
            fb = Comun.leer_csv(self.ruta_fb03)
            etapa_fb = EtapaFB(fb, Comun.es_sin_datos(fb))
            etapa_fb.preparar()

        etapa_gallo = EtapaGallo(gallo, update_status, con_ref_pago=con_vbrl)
        pasos_gallo = etapa_gallo.preparar()

        update_status("📊 Construyendo matriz 'TD'...")
        etapa_td = EtapaTD(gallo, pasos_gallo)
        td_resumen = etapa_td.preparar()
        total_importe_gallo = td_resumen["total_importe_gallo"]

        update_status("🧾 Construyendo matrices 'MTD' (SG / RG / PP / TODOS)...")
        etapa_monivoi = EtapaMonivoi(monivoi, monivoi_sin_datos)
        mtd = etapa_monivoi.preparar()
        etapa_monivoi.insertar_referencia2()

        cruce = EtapaCruceTDMTD(td_resumen, mtd)
        cruce_avisos = cruce.avisos()

        etapa_td_ref = etapa_notas_pep = listas_2da = etapa_validacion = None
        if con_vbrl:
            etapa_td_ref, etapa_notas_pep, listas_2da = preparar_tablas_diferencias(
                gallo, pasos_gallo, etapa_td, etapa_vbrl, etapa_fb, update_status)
            update_status("📋 Construyendo 'Validación_doc'...")
            etapa_validacion = EtapaValidacionDoc(etapa_td_ref, etapa_notas_pep)
            etapa_validacion.preparar()

        etapa_diferencias = None
        if etapa_hojas is not None:
            update_status("🧮 Construyendo 'Diferencias'...")
            etapa_diferencias = EtapaDiferencias(etapa_validacion, etapa_hojas)
            etapa_diferencias.preparar()

        update_status("💾 Guardando Excel consolidado...")
        os.makedirs(self.ruta_output, exist_ok=True)

        nombre_salida = self.nombre_salida
        if not nombre_salida:
            base_gallo = os.path.splitext(os.path.basename(self.ruta_gallo))[0]
            nombre_salida = f"Consolidado_{base_gallo}.xlsx"
        ruta_final = os.path.join(self.ruta_output, nombre_salida)
        ruta_temporal = os.path.join(
            self.ruta_output, f".{os.path.splitext(nombre_salida)[0]}.{os.getpid()}.tmp.xlsx")

        try:
            self._escribir_libro(
                ruta_temporal, etapa_gallo, monivoi, etapa_vbrl, etapa_fb, etapa_td, cruce,
                etapa_td_ref, etapa_notas_pep, etapa_monivoi, etapa_validacion,
                etapa_diferencias, etapa_hojas)
            os.replace(ruta_temporal, ruta_final)
        finally:
            if os.path.exists(ruta_temporal):
                try:
                    os.remove(ruta_temporal)
                except OSError:
                    pass

        update_status("✅ Consolidación completada")

        resumen_vbrl = etapa_vbrl.resultado if etapa_vbrl is not None else None
        resumen_ref = etapa_td_ref.resultado if etapa_td_ref is not None else None
        resumen_fb = etapa_fb.resultado if etapa_fb is not None else None
        resumen_np = etapa_notas_pep.resultado if etapa_notas_pep is not None else None
        resumen_val = etapa_validacion.resultado if etapa_validacion is not None else None
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
            "td_avisos":            pasos_gallo["avisos"] + td_resumen["avisos"],
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
            "td_referencias_na":    resumen_ref["na_n"] if resumen_ref else 0,
            "td_referencias_avisos": (resumen_ref["avisos"] + resumen_val["avisos"]) if resumen_ref else [],
            "filas_validacion_doc": (
                {clave: len(datos["filas"]) for clave, datos in resumen_val["bloques"].items()}
                if resumen_val else {}),
            "fb03_incluido":        con_notas_pep,
            "filas_fb":             resumen_fb["filas"] if resumen_fb else 0,
            "filas_notas_ref":      len(resumen_np["bloques"]["notas"]["filas"]) if resumen_np else 0,
            "total_notas_ref":      resumen_np["bloques"]["notas"]["total"] if resumen_np else 0.0,
            "notas_ref_na":         resumen_np["bloques"]["notas"]["na_n"] if resumen_np else 0,
            "filas_pep_ref":        len(resumen_np["bloques"]["pep"]["filas"]) if resumen_np else 0,
            "total_pep_ref":        resumen_np["bloques"]["pep"]["total"] if resumen_np else 0.0,
            "pep_ref_na":           resumen_np["bloques"]["pep"]["na_n"] if resumen_np else 0,
            "notas_pep_avisos":     notas_pep_avisos,
            "listas_2da":           listas_2da,
            "segunda_incluida":     etapa_hojas is not None,
            "filas_hojas_segunda":  ({d["hoja"]: len(d["filas"]) for k, d in etapa_hojas.resultado.items()
                                      if k != "avisos"} if etapa_hojas is not None else {}),
            "diferencias":          etapa_diferencias.resultado["conteo"] if etapa_diferencias else None,
            "segunda_avisos":       ((etapa_hojas.resultado["avisos"] + etapa_diferencias.resultado["avisos"])
                                     if etapa_diferencias else []),
        }


def consolidar_gallo_monivoi(
    ruta_gallo,
    ruta_monivoi,
    ruta_output,
    nombre_salida=None,
    callback_status=None,
    ruta_vbrl=None,
    ruta_fb03=None,
    rutas_segunda=None,
):
    """The controller's entry point: Consolidador(...).ejecutar() (see its docs). `rutas_segunda` makes
    it the 2nd consolidation: {vbrk, fbl5n, vbrp, knvp, vbrl2, bkpf: path} (needs `ruta_vbrl`)."""
    return Consolidador(
        ruta_gallo, ruta_monivoi, ruta_output,
        nombre_salida=nombre_salida,
        callback_status=callback_status,
        ruta_vbrl=ruta_vbrl,
        ruta_fb03=ruta_fb03,
        rutas_segunda=rutas_segunda,
    ).ejecutar()


def calcular_listas_segunda_corrida(ruta_gallo, ruta_vbrl, ruta_fb03=None, callback_status=None):
    """2nd-run lists WITHOUT a workbook ('2a corrida' button): same stages as the consolidation up to
    preparar_tablas_diferencias, so lists equal the consolidated sheet's #N/A rows. `ruta_fb03` optional
    (no FB03: only Facturas + warning). Returns EtapaListasSegundaCorrida.resultado."""
    status = callback_status or (lambda _mensaje: None)
    status("📄 Leyendo archivo de VBRL...")
    vbrl = Comun.leer_csv(ruta_vbrl)
    etapa_vbrl = EtapaVBRL(vbrl, Comun.es_sin_datos(vbrl))
    etapa_vbrl.preparar()
    etapa_fb = None
    if ruta_fb03:
        status("📄 Leyendo archivo de FB03...")
        fb = Comun.leer_csv(ruta_fb03)
        etapa_fb = EtapaFB(fb, Comun.es_sin_datos(fb))
        etapa_fb.preparar()

    status("📄 Leyendo archivo de Gallo...")
    gallo = Comun.leer_csv(ruta_gallo)
    if Comun.es_sin_datos(gallo):
        raise ValueError(
            "El archivo de Gallo no tiene datos (descarga 'sin partidas'): no hay renglones "
            "de los que salgan las listas de la 2a corrida."
        )
    pasos_gallo = EtapaGallo(gallo, status, con_ref_pago=True).preparar()
    status("📊 Construyendo matriz 'TD'...")
    etapa_td = EtapaTD(gallo, pasos_gallo)
    etapa_td.preparar()
    _, _, listas = preparar_tablas_diferencias(
        gallo, pasos_gallo, etapa_td, etapa_vbrl, etapa_fb, status)
    return listas


def _consolidar_venta(etapa_cls, ruta_csv, ruta_output, nombre_salida, callback_status,
                      fechas_seleccionadas):
    """Workbook of ONE sale: its CSV as 'Gallo' (as downloaded, must fit one sheet) plus the sale's tab from
    `etapa_cls`. Returns the GUI summary dict: ruta, filas_archivo, total_venta, venta_avisos."""
    status = callback_status or (lambda _mensaje: None)
    fechas_sel = EtapaVentas.parsear_fechas_seleccionadas(fechas_seleccionadas)
    (clave, nombre_hoja, _, _), = etapa_cls.BLOQUES

    status(f"📄 Leyendo archivo de {etapa_cls.ORIGEN}...")
    datos = Comun.leer_csv(ruta_csv)
    if Comun.es_sin_datos(datos):
        raise ValueError(
            f"El archivo de {etapa_cls.ORIGEN} no tiene datos (descarga 'sin partidas'). "
            "No hay nada que consolidar."
        )

    if len(datos) + 1 > EtapaVentas.MAX_FILAS_EXCEL:
        raise ValueError(
            f"El archivo de {etapa_cls.ORIGEN} trae {len(datos):,} filas y no cabe en la pestaña "
            f"'{EtapaGallo.HOJA_GALLO}' de Excel ({EtapaVentas.MAX_FILAS_EXCEL:,} filas con el "
            "encabezado). Descarga un periodo más corto."
        )

    status(f"📆 Construyendo '{nombre_hoja}'...")
    etapa = etapa_cls.desde_csv(datos, fechas_sel)
    ventas = etapa.preparar()

    status("💾 Guardando Excel consolidado...")
    os.makedirs(ruta_output, exist_ok=True)
    if not nombre_salida:
        nombre_salida = f"Consolidado_{os.path.splitext(os.path.basename(ruta_csv))[0]}.xlsx"
    ruta_final = os.path.join(ruta_output, nombre_salida)

    with pd.ExcelWriter(ruta_final, engine="openpyxl") as writer:
        datos.to_excel(writer, sheet_name=EtapaGallo.HOJA_GALLO, index=False)
        etapa.escribir(writer.book.create_sheet(nombre_hoja), clave)

    status("✅ Consolidación completada")
    return {
        "ruta":          ruta_final,
        "filas_archivo": len(datos),
        "total_venta":   ventas[clave]["total"],
        "venta_avisos":  ventas["avisos"],
    }


def consolidar_venta_dsd(ruta_dsd, ruta_output, nombre_salida=None, callback_status=None,
                         fechas_seleccionadas=None):
    """Entry point of the 'Venta DSD' button: workbook with 'Gallo' and 'Venta DSD'. `fechas_seleccionadas`
    ('DD.MM.YYYY'): every day of the queried range, shown even without amounts. FileNotFoundError /
    ValueError as in Consolidador.ejecutar()."""
    return _consolidar_venta(EtapaVentaDSD, ruta_dsd, ruta_output, nombre_salida, callback_status,
                             fechas_seleccionadas)


def consolidar_venta_eiap(ruta_eiap, ruta_output, nombre_salida=None, callback_status=None,
                          fechas_seleccionadas=None):
    """The controller's entry point for the 'Venta EIAP' button: a workbook with the tabs 'Gallo' and
    'Venta EIAP'. Same arguments and errors as consolidar_venta_dsd().
    """
    return _consolidar_venta(EtapaVentaEIAP, ruta_eiap, ruta_output, nombre_salida, callback_status,
                             fechas_seleccionadas)


# Quick test without the GUI:  python Consolidacion.py
if __name__ == "__main__":
    print("=== Prueba de consolidación (standalone) ===")
    ruta_gallo   = input("Ruta del CSV de Gallo   : ").strip().strip('"')
    ruta_monivoi = input("Ruta del CSV de Monivoi : ").strip().strip('"')
    ruta_vbrl    = input("Ruta del CSV de VBRL (Enter = sin VBRL): ").strip().strip('"') or None
    ruta_fb03    = input("Ruta del CSV de FB03 (Enter = sin FB03): ").strip().strip('"') or None

    salida = carpeta_output_desde_input(os.path.dirname(ruta_gallo))

    try:
        resumen = consolidar_gallo_monivoi(
            ruta_gallo, ruta_monivoi, salida,
            callback_status=print,
            ruta_vbrl=ruta_vbrl,
            ruta_fb03=ruta_fb03,
        )
        print("\n🎉 CONSOLIDACIÓN COMPLETADA 🎉")
        for clave, valor in resumen.items():
            print(f"  {clave:<20}: {valor}")
    except Exception as e:
        print(f"❌ Error durante la consolidación: {e}")
        raise
