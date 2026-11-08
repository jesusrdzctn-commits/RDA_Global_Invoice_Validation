"""Validación Factura Global — helpers without SAP or tkinter (the GUI opens and tests on its own).

The ONLY place for the Spanish months, the file-name formats and the date offsets of the SAP
queries: Monivoi +1 day, the 2nd Monivoi that same day ±1 and the sales buttons ('Venta DSD',
'Venta EIAP') −1 day on the end date, plus the accounts and document classes of each sale (asked
to SAP AND summed by its tab). File NAMES always keep the dates the user typed.
"""

import json
import os
import sys
import unicodedata
from datetime import datetime, timedelta

# Date format the user types everywhere in the app.
FORMATO_FECHA = "%d.%m.%Y"

# Company code (BUKRS) the app starts with: the GUI field and the command line default to it.
SOCIEDAD_DEFECTO = "MX21"

# The two sales buttons, 'Venta EIAP' and 'Venta DSD': each has its OWN FAGLL03 query (a merged one
# was too big and the program failed). What SAP is asked for AND what its tab sums come from these
# lists (one source for both). To add an account or a class to a sale, add it to ITS list: the query
# and the tab follow. (4000010 is in both lists; the classes do not overlap.)
CUENTAS_EIAP = ("4000010", "4000020", "4000802", "4300013", "4300014")
CLASES_EIAP = ("RV", "DA")
CUENTAS_DSD = ("4000005", "4000010")
CLASES_DSD = ("C1", "C2", "C5", "C6", "X0")

MESES_ES = {
    1: "ene", 2: "feb", 3: "mar", 4: "abr", 5: "may", 6: "jun",
    7: "jul", 8: "ago", 9: "sep", 10: "oct", 11: "nov", 12: "dic",
}


# --- Text ---
def quitar_acentos(texto):
    """'Validación' -> 'Validacion': the text without accents, everything else as it was."""
    descompuesto = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in descompuesto if not unicodedata.combining(c))


def normalizar_encabezado(texto):
    """A header normalized to recognize SPELLING VARIANTS of the same field (accents, dots, spaces, case):
    'Fe.Contabilización' == 'fecha contabilizacion', 'Solic.' == 'solic'. Never used to guess another column.
    The ONE rule of the downloads (that check the files they write) and of the consolidation (that reads them).
    """
    return "".join(c for c in quitar_acentos(texto).lower() if c.isalnum())


# --- Dates ---
def parse_fecha(fecha_str):
    """'DD.MM.YYYY' -> datetime. Raises ValueError if the format is invalid."""
    return datetime.strptime(fecha_str.strip(), FORMATO_FECHA)


def fecha_efectiva_monivoi(fecha_str):
    """Monivoi works with the day AFTER the typed one (where the stamping the stakeholders use
    lives). Only for the SAP query (SO_ERDAT): '18.08.2026' -> datetime(2026, 8, 19).
    """
    return parse_fecha(fecha_str) + timedelta(days=1)


def fecha_efectiva_monivoi_str(fecha_str):
    """fecha_efectiva_monivoi() as 'DD.MM.YYYY', for SO_ERDAT: '18.08.2026' -> '19.08.2026'."""
    return fecha_efectiva_monivoi(fecha_str).strftime(FORMATO_FECHA)


# The 2nd Monivoi (UUIDs of BAMA / OXXO) widens the 1st one's days by this many days on EACH side:
# SAT stamps are sometimes processed a day early or late.
MARGEN_DIAS_MONIVOI_2 = 1


def rango_monivoi_2(desde_str, hasta_str):
    """SO_ERDAT (LOW, HIGH) of the 2nd Monivoi: the 1st one's days ±MARGEN_DIAS_MONIVOI_2, i.e. typed
    '25.10.2026' -> ('25.10.2026', '27.10.2026'). ValueError for a bad date or a backwards range."""
    if parse_fecha(hasta_str) < parse_fecha(desde_str):
        raise ValueError(f"El periodo es inválido: {desde_str} — {hasta_str}.")
    margen = timedelta(days=MARGEN_DIAS_MONIVOI_2)
    desde = fecha_efectiva_monivoi(desde_str) - margen
    hasta = fecha_efectiva_monivoi(hasta_str) + margen
    return desde.strftime(FORMATO_FECHA), hasta.strftime(FORMATO_FECHA)


def fecha_efectiva_eiap(fecha_str):
    """The sales buttons ('Venta DSD', 'Venta EIAP') query SAP up to the day BEFORE the typed end date
    (only their queries): '10.01.2026' -> datetime(2026, 1, 9); '01.03.2026' -> datetime(2026, 2, 28).
    """
    return parse_fecha(fecha_str) - timedelta(days=1)


def _sufijo_dia(f):
    """datetime -> 'D_mmm_YY' (e.g. 10_ago_26; no leading zero on the day)."""
    return f"{f.day}_{MESES_ES[f.month]}_{f:%y}"


# --- File names (no offsets: the typed dates) ---
def nombre_archivo_dia(fecha_str):
    """'18.08.2026' -> '18_ago_26.csv'."""
    return f"{_sufijo_dia(parse_fecha(fecha_str))}.csv"


def nombre_archivo_rango(desde, hasta):
    """('18.08.2026', '20.08.2026') -> '18_ago_26_a_20_ago_26.csv'."""
    d = _sufijo_dia(parse_fecha(desde))
    h = _sufijo_dia(parse_fecha(hasta))
    return f"{d}_a_{h}.csv"


def nombre_archivo_dia_monivoi(fecha_str):
    """'18.08.2026' -> 'Monitoreo_18_ago_26.csv'  (name = the user's date)."""
    return f"Monitoreo_{_sufijo_dia(parse_fecha(fecha_str))}.csv"


def nombre_archivo_rango_monivoi(desde, hasta):
    """('18.08.2026', '20.08.2026') -> 'Monitoreo_18_ago_26_a_20_ago_26.csv'."""
    d = _sufijo_dia(parse_fecha(desde))
    h = _sufijo_dia(parse_fecha(hasta))
    return f"Monitoreo_{d}_a_{h}.csv"


# VBRL and FB03 are downloaded ONCE per period (with the Monivoi / Gallo file of that period)
# and carry its dates, so the files of a run sit together.
def nombre_archivo_dia_vbrl(fecha_str):
    """'18.08.2026' -> 'VBRL_18_ago_26.csv'."""
    return f"VBRL_{_sufijo_dia(parse_fecha(fecha_str))}.csv"


def nombre_archivo_rango_vbrl(desde, hasta):
    """('18.08.2026', '20.08.2026') -> 'VBRL_18_ago_26_a_20_ago_26.csv'."""
    d = _sufijo_dia(parse_fecha(desde))
    h = _sufijo_dia(parse_fecha(hasta))
    return f"VBRL_{d}_a_{h}.csv"


def nombre_archivo_dia_fb03(fecha_str):
    """'18.08.2026' -> 'FB03_18_ago_26.csv'."""
    return f"FB03_{_sufijo_dia(parse_fecha(fecha_str))}.csv"


def nombre_archivo_rango_fb03(desde, hasta):
    """('18.08.2026', '20.08.2026') -> 'FB03_18_ago_26_a_20_ago_26.csv'."""
    d = _sufijo_dia(parse_fecha(desde))
    h = _sufijo_dia(parse_fecha(hasta))
    return f"FB03_{d}_a_{h}.csv"


# 2nd run (after the consolidation): one file per transaction, named like VBRL / FB03 with the
# dates of the period (all of them .csv: the ZSE16 tables are converted from SAP's .xlsx).
ARCHIVOS_SEGUNDA_CORRIDA = {
    "vbrp":     ("VBRP",        ".csv"),
    "vbrk":     ("VBRK",        ".csv"),
    "vbrl2":    ("VBRL_2",      ".csv"),
    "kna1":     ("KNA1",        ".csv"),
    "knvp":     ("KNVP",        ".csv"),
    "bkpf":     ("BKPF",        ".csv"),
    "monivoi2": ("Monitoreo_2", ".csv"),
}


def nombre_archivo_dia_2da(clave, fecha_str):
    """('vbrp', '18.08.2026') -> 'VBRP_18_ago_26.csv'; KeyError for an unknown key."""
    prefijo, ext = ARCHIVOS_SEGUNDA_CORRIDA[clave]
    return f"{prefijo}_{_sufijo_dia(parse_fecha(fecha_str))}{ext}"


def nombre_archivo_rango_2da(clave, desde, hasta):
    """('kna1', '18.08.2026', '20.08.2026') -> 'KNA1_18_ago_26_a_20_ago_26.csv'."""
    prefijo, ext = ARCHIVOS_SEGUNDA_CORRIDA[clave]
    d = _sufijo_dia(parse_fecha(desde))
    h = _sufijo_dia(parse_fecha(hasta))
    return f"{prefijo}_{d}_a_{h}{ext}"


# The two runs the main download button does in ONE go, the 1st before the 2nd (the lists of the 2nd are
# computed from the files of the 1st): (controller key, name in the GUI's file list). ONE definition for that
# list and for the transactions the controller runs (a test ties them together); the controller runs the 2nd
# run in the order of its PASADAS_2DA, the GUI lists its files in the order below.
CORRIDAS_DESCARGA = {
    1: (("gallo", "Gallo"), ("monivoi", "Monivoi"), ("vbrl", "VBRL"), ("fb03", "FB03")),
    2: (("vbrp", "VBRP"), ("vbrk", "VBRK"), ("vbrl2", "VBRL_2"),
        ("kna1", "KNA1"), ("knvp", "KNVP"), ("bkpf", "BKPF"), ("monivoi2", "Monivoi_2")),
}


# --- Folders and config.json (ONE place for the GUI and the command line) ---
def ruta_config():
    """config.json (last download/consolidation folders): next to the .exe (PyInstaller) or to this .py."""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "config.json")


def carpeta_base_proyecto():
    """Documents\\Validacion Factura Global\\src of the current Windows user (parent of Input, Output, Logs)."""
    perfil = os.environ.get("USERPROFILE") or os.path.expanduser("~")
    return os.path.join(perfil, "Documents", "Validacion Factura Global", "src")


def carpetas_por_defecto():
    """(download folder, consolidation folder) when config.json has none: <base>\\Input and <base>\\Output."""
    base = carpeta_base_proyecto()
    return os.path.join(base, "Input"), os.path.join(base, "Output")


def cargar_config(ruta=None):
    """The whole config.json as a dict; {} if it is missing, unreadable or not a JSON object."""
    try:
        with open(ruta or ruta_config(), "r", encoding="utf-8") as f:
            datos = json.load(f)
    except (OSError, ValueError):          # ValueError covers JSONDecodeError and bad UTF-8
        return {}
    return datos if isinstance(datos, dict) else {}


def carpetas_configuradas(config=None):
    """(download folder, consolidation folder) the app starts with: the ones in `config` (default: the
    config.json), each falling back to its default when missing, blank or not a text."""
    config = cargar_config() if config is None else config
    por_defecto = carpetas_por_defecto()
    elegidas = []
    for clave, defecto in zip(("ruta_descarga", "ruta_consolidacion"), por_defecto):
        valor = config.get(clave)
        valor = valor.strip() if isinstance(valor, str) else ""
        elegidas.append(valor or defecto)
    return tuple(elegidas)


# --- The name of the file of ANY transaction (the GUI previews them, the command line uses this) ---
_NOMBRES_1A_CORRIDA = {
    "gallo":   (nombre_archivo_dia,         nombre_archivo_rango),
    "monivoi": (nombre_archivo_dia_monivoi, nombre_archivo_rango_monivoi),
    "vbrl":    (nombre_archivo_dia_vbrl,    nombre_archivo_rango_vbrl),
    "fb03":    (nombre_archivo_dia_fb03,    nombre_archivo_rango_fb03),
}


def nombre_archivo_transaccion(clave, fecha=None, desde=None, hasta=None):
    """File name of the transaction `clave` ('gallo', 'monivoi', 'vbrl', 'fb03' or any key of
    ARCHIVOS_SEGUNDA_CORRIDA) for ONE day (`fecha`) or a range (`desde` and `hasta`), exactly the one the
    GUI shows: ('vbrk', fecha='03.11.2026') -> 'VBRK_3_nov_26.csv'. ValueError: unknown key, a bad or missing
    date, or both a day and a range."""
    if fecha is not None and (desde is not None or hasta is not None):
        raise ValueError("Indica un día (fecha) o un rango (desde y hasta), no los dos.")
    if fecha is None and (desde is None or hasta is None):
        raise ValueError("Falta la fecha, o el rango completo (desde y hasta).")
    if clave in _NOMBRES_1A_CORRIDA:
        por_dia, por_rango = _NOMBRES_1A_CORRIDA[clave]
        return por_dia(fecha) if fecha is not None else por_rango(desde, hasta)
    if clave in ARCHIVOS_SEGUNDA_CORRIDA:
        return (nombre_archivo_dia_2da(clave, fecha) if fecha is not None
                else nombre_archivo_rango_2da(clave, desde, hasta))
    raise ValueError(f"Transacción desconocida: {clave!r}.")


# The sales buttons only work with a range ("Varios días"): no per-day name.
def nombre_archivo_rango_eiap(desde, hasta):
    """('01.01.2026', '10.01.2026') -> 'Venta_EIAP_1_ene_26_a_10_ene_26.csv'."""
    d = _sufijo_dia(parse_fecha(desde))
    h = _sufijo_dia(parse_fecha(hasta))
    return f"Venta_EIAP_{d}_a_{h}.csv"


def nombre_archivo_rango_dsd(desde, hasta):
    """('01.01.2026', '10.01.2026') -> 'Venta_DSD_1_ene_26_a_10_ene_26.csv'."""
    d = _sufijo_dia(parse_fecha(desde))
    h = _sufijo_dia(parse_fecha(hasta))
    return f"Venta_DSD_{d}_a_{h}.csv"


# Quick test:  python utils.py
if __name__ == "__main__":
    print("Gallo   día   :", nombre_archivo_dia("15.06.2026"))
    print("Gallo   rango :", nombre_archivo_rango("15.06.2026", "17.06.2026"))
    print("Monivoi día   :", nombre_archivo_dia_monivoi("11.08.2026"))
    print("Monivoi rango :", nombre_archivo_rango_monivoi("11.08.2026", "13.08.2026"))
    print("Monivoi a SAP :", fecha_efectiva_monivoi_str("18.08.2026"))
    print("Monivoi 2 SAP :", rango_monivoi_2("18.08.2026", "18.08.2026"))
    print("VBRL    día   :", nombre_archivo_dia_vbrl("11.08.2026"))
    print("VBRL    rango :", nombre_archivo_rango_vbrl("11.08.2026", "13.08.2026"))
    print("FB03    día   :", nombre_archivo_dia_fb03("11.08.2026"))
    print("FB03    rango :", nombre_archivo_rango_fb03("11.08.2026", "13.08.2026"))
    print("EIAP    rango :", nombre_archivo_rango_eiap("01.01.2026", "10.01.2026"))
    print("DSD     rango :", nombre_archivo_rango_dsd("01.01.2026", "10.01.2026"))
    print("EIAP    a SAP :", fecha_efectiva_eiap("10.01.2026").strftime(FORMATO_FECHA))
    for clave in ARCHIVOS_SEGUNDA_CORRIDA:
        print(f"{clave:<7} día   :", nombre_archivo_dia_2da(clave, "11.08.2026"))
