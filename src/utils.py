"""Validación Factura Global — helpers without SAP or tkinter (the GUI opens and tests on its own).

The ONLY place for the Spanish months, the file-name formats and the two date offsets of the
SAP queries: Monivoi +1 day and the sales buttons ('Venta DSD', 'Venta EIAP') −1 day on the end
date, plus the accounts and document classes of each sale (asked to SAP AND summed by its tab).
File NAMES always keep the dates the user typed.
"""

from datetime import datetime, timedelta

# Date format the user types everywhere in the app.
FORMATO_FECHA = "%d.%m.%Y"

# The two sales buttons, 'Venta EIAP' and 'Venta DSD': each has its OWN FAGLL03 query (a merged one
# was too big and the program failed). What SAP is asked for AND what its tab sums come from these
# lists (one source for both). To add an account or a class to a sale, add it to ITS list: the query
# and the tab follow. (4000010 is in both lists; the classes do not overlap.)
CUENTAS_EIAP = ("4000010", "4000020", "4000802", "4300013")
CLASES_EIAP = ("RV", "DA")
CUENTAS_DSD = ("4000005", "4000010")
CLASES_DSD = ("C1", "C2", "C5", "C6", "X0")

MESES_ES = {
    1: "ene", 2: "feb", 3: "mar", 4: "abr", 5: "may", 6: "jun",
    7: "jul", 8: "ago", 9: "sep", 10: "oct", 11: "nov", 12: "dic",
}


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


def fecha_efectiva_eiap(fecha_str):
    """The sales buttons ('Venta DSD', 'Venta EIAP') query SAP up to the day BEFORE the typed end date
    (only their queries): '10.01.2026' -> datetime(2026, 1, 9); '01.03.2026' -> datetime(2026, 2, 28).
    """
    return parse_fecha(fecha_str) - timedelta(days=1)


def fecha_efectiva_eiap_str(fecha_str):
    """fecha_efectiva_eiap() as 'DD.MM.YYYY', for SO_BUDAT-HIGH: '10.01.2026' -> '09.01.2026'."""
    return fecha_efectiva_eiap(fecha_str).strftime(FORMATO_FECHA)


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
# dates of the period. FBL5N is SAP's own text list, kept with the extension of its recording.
ARCHIVOS_SEGUNDA_CORRIDA = {
    "vbrp":  ("VBRP",   ".csv"),
    "vbrk":  ("VBRK",   ".csv"),
    "vbrl2": ("VBRL_2", ".csv"),
    "fbl5n": ("FBL5N",  ".XLS"),
    "knvp":  ("KNVP",   ".csv"),
    "bkpf":  ("BKPF",   ".csv"),
}


def nombre_archivo_dia_2da(clave, fecha_str):
    """('vbrp', '18.08.2026') -> 'VBRP_18_ago_26.csv'; KeyError for an unknown key."""
    prefijo, ext = ARCHIVOS_SEGUNDA_CORRIDA[clave]
    return f"{prefijo}_{_sufijo_dia(parse_fecha(fecha_str))}{ext}"


def nombre_archivo_rango_2da(clave, desde, hasta):
    """('fbl5n', '18.08.2026', '20.08.2026') -> 'FBL5N_18_ago_26_a_20_ago_26.XLS'."""
    prefijo, ext = ARCHIVOS_SEGUNDA_CORRIDA[clave]
    d = _sufijo_dia(parse_fecha(desde))
    h = _sufijo_dia(parse_fecha(hasta))
    return f"{prefijo}_{d}_a_{h}{ext}"


# The two runs of the main download button (GUI radio buttons): (controller key, name in the legend).
# ONE definition for the GUI legends and for what the controller runs (a test ties them together).
CORRIDAS_DESCARGA = {
    1: (("gallo", "Gallo"), ("monivoi", "Monivoi"), ("vbrl", "VBRL"), ("fb03", "FB03")),
    2: (("vbrp", "VBRP"), ("vbrk", "VBRK"), ("vbrl2", "VBRL_2"),
        ("fbl5n", "FBL5N"), ("knvp", "KNVP"), ("bkpf", "BKPF")),
}


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
    print("VBRL    día   :", nombre_archivo_dia_vbrl("11.08.2026"))
    print("VBRL    rango :", nombre_archivo_rango_vbrl("11.08.2026", "13.08.2026"))
    print("FB03    día   :", nombre_archivo_dia_fb03("11.08.2026"))
    print("FB03    rango :", nombre_archivo_rango_fb03("11.08.2026", "13.08.2026"))
    print("EIAP    rango :", nombre_archivo_rango_eiap("01.01.2026", "10.01.2026"))
    print("DSD     rango :", nombre_archivo_rango_dsd("01.01.2026", "10.01.2026"))
    print("EIAP    a SAP :", fecha_efectiva_eiap_str("10.01.2026"))
    for clave in ARCHIVOS_SEGUNDA_CORRIDA:
        print(f"{clave:<7} día   :", nombre_archivo_dia_2da(clave, "11.08.2026"))
