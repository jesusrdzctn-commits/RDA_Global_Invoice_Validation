"""
Validación Factura Global — utils.py
====================================

Shared helpers WITHOUT SAP or tkinter dependencies, so that:
  - the GUI (interfaz_GUI.py) can use them and still open/test on its own, and
  - the SAP layer (Descargas_SAP.py) can use them without duplicating logic.

This is the ONLY source of truth for:
  - the Spanish month mapping,
  - the file-name format (Gallo, Monivoi, VBRL and Venta EIAP),
  - Monivoi's "one day later" offset  (the +1 lives ONLY here!), and
  - the 'Solo Venta EIAP' "one day earlier" offset on the end date (the -1
    lives ONLY here!).

Golden rule: if the file-name format or one of the offsets ever changes, this
file is the only one to touch.
"""

from datetime import datetime, timedelta

# Date format the user types everywhere in the app.
FORMATO_FECHA = "%d.%m.%Y"

# Month in Spanish, 3 letters — used to name the output files.
MESES_ES = {
    1: "ene", 2: "feb", 3: "mar", 4: "abr", 5: "may", 6: "jun",
    7: "jul", 8: "ago", 9: "sep", 10: "oct", 11: "nov", 12: "dic",
}


# ----------------------------------------------------------------------
# Dates
# ----------------------------------------------------------------------
def parse_fecha(fecha_str):
    """'DD.MM.YYYY' -> datetime. Raises ValueError if the format is invalid."""
    return datetime.strptime(fecha_str.strip(), FORMATO_FECHA)


def fecha_efectiva_monivoi(fecha_str):
    """
    Monivoi works with the day AFTER the one the user types, because that is
    where the stamping (timbrado) the stakeholders use as reference lives.
    This is the ONLY place where that +1 is applied, and it ONLY affects the
    SAP QUERY (SO_ERDAT). The file NAME does not go through here: it keeps the
    user's date on purpose, so the stakeholder is not confused.

        '18.08.2026' -> datetime(2026, 8, 19)
    """
    return parse_fecha(fecha_str) + timedelta(days=1)


def fecha_efectiva_monivoi_str(fecha_str):
    """
    Same as fecha_efectiva_monivoi(), but returns the 'DD.MM.YYYY' string,
    ready to be typed into the SAP fields (SO_ERDAT).

        '18.08.2026' -> '19.08.2026'
    """
    return fecha_efectiva_monivoi(fecha_str).strftime(FORMATO_FECHA)


def fecha_efectiva_eiap(fecha_str):
    """
    'Solo Venta EIAP' queries SAP up to the day BEFORE the end date the user
    types (the typed end date is left out). This is the ONLY place where that
    -1 is applied, and it ONLY affects the SAP QUERY of that button: the regular
    Gallo / Monivoi / VBRL downloads never go through here, and the file NAME
    keeps the user's dates on purpose (same rule as Monivoi).

        '10.01.2026' -> datetime(2026, 1, 9)
        '01.03.2026' -> datetime(2026, 2, 28)
    """
    return parse_fecha(fecha_str) - timedelta(days=1)


def fecha_efectiva_eiap_str(fecha_str):
    """
    Same as fecha_efectiva_eiap(), but returns the 'DD.MM.YYYY' string, ready
    to be typed into the SAP field (SO_BUDAT-HIGH).

        '10.01.2026' -> '09.01.2026'
    """
    return fecha_efectiva_eiap(fecha_str).strftime(FORMATO_FECHA)


def _sufijo_dia(f):
    """datetime -> 'D_mmm_YY' (e.g. 10_ago_26; no leading zero on the day). Internal use."""
    return f"{f.day}_{MESES_ES[f.month]}_{f:%y}"


# ----------------------------------------------------------------------
# File names — Gallo (FAGLL03)   [no offset]
# ----------------------------------------------------------------------
def nombre_archivo_dia(fecha_str):
    """'18.08.2026' -> '18_ago_26.csv'."""
    return f"{_sufijo_dia(parse_fecha(fecha_str))}.csv"


def nombre_archivo_rango(desde, hasta):
    """('18.08.2026', '20.08.2026') -> '18_ago_26_a_20_ago_26.csv'."""
    d = _sufijo_dia(parse_fecha(desde))
    h = _sufijo_dia(parse_fecha(hasta))
    return f"{d}_a_{h}.csv"


# ----------------------------------------------------------------------
# File names — Monivoi (ZLMXCOM_TRN_MONINVOI)   [NO offset in the name]
# NOTE: the +1 ONLY applies to the SAP QUERY (SO_ERDAT, via
# fecha_efectiva_monivoi_str). The file NAME carries the SAME date the user
# types, even though the data inside is from the next day. That is on
# purpose: this way the stakeholder is not confused by the file's date.
# ----------------------------------------------------------------------
def nombre_archivo_dia_monivoi(fecha_str):
    """'18.08.2026' -> 'Monitoreo_18_ago_26.csv'  (name = the user's date)."""
    return f"Monitoreo_{_sufijo_dia(parse_fecha(fecha_str))}.csv"


def nombre_archivo_rango_monivoi(desde, hasta):
    """('18.08.2026', '20.08.2026') -> 'Monitoreo_18_ago_26_a_20_ago_26.csv'."""
    d = _sufijo_dia(parse_fecha(desde))
    h = _sufijo_dia(parse_fecha(hasta))
    return f"Monitoreo_{d}_a_{h}.csv"


# ----------------------------------------------------------------------
# File names — VBRL (ZSE16, table VBRL)   [no offset]
# VBRL is NOT downloaded by date: it is downloaded ONCE per period, for the
# invoices listed in the Monivoi file of that same period. Its name therefore
# carries the same dates as the Monivoi file (the user's dates, no offset), so
# the three files of a run sit together and are easy to match by eye.
# ----------------------------------------------------------------------
def nombre_archivo_dia_vbrl(fecha_str):
    """'18.08.2026' -> 'VBRL_18_ago_26.csv'."""
    return f"VBRL_{_sufijo_dia(parse_fecha(fecha_str))}.csv"


def nombre_archivo_rango_vbrl(desde, hasta):
    """('18.08.2026', '20.08.2026') -> 'VBRL_18_ago_26_a_20_ago_26.csv'."""
    d = _sufijo_dia(parse_fecha(desde))
    h = _sufijo_dia(parse_fecha(hasta))
    return f"VBRL_{d}_a_{h}.csv"


# ----------------------------------------------------------------------
# File names — Venta EIAP (FAGLL03 with the EIAP accounts)   [NO offset in the name]
# NOTE: the -1 ONLY applies to the SAP QUERY (see fecha_efectiva_eiap_str).
# The file NAME carries the SAME dates the user types, so the files of one run
# (Gallo, Monivoi, VBRL, Venta EIAP) sit together and are easy to match by eye.
# The button only works with a range ("Varios días"), so there is no per-day
# name.
# ----------------------------------------------------------------------
def nombre_archivo_rango_eiap(desde, hasta):
    """('01.01.2026', '10.01.2026') -> 'Venta_EIAP_1_ene_26_a_10_ene_26.csv'."""
    d = _sufijo_dia(parse_fecha(desde))
    h = _sufijo_dia(parse_fecha(hasta))
    return f"Venta_EIAP_{d}_a_{h}.csv"


# ----------------------------------------------------------------------
# Quick test:  python utils.py
# ----------------------------------------------------------------------
if __name__ == "__main__":
    print("Gallo   día   :", nombre_archivo_dia("15.06.2026"))
    print("Gallo   rango :", nombre_archivo_rango("15.06.2026", "17.06.2026"))
    print("Monivoi día   :", nombre_archivo_dia_monivoi("11.08.2026"))
    print("Monivoi rango :", nombre_archivo_rango_monivoi("11.08.2026", "13.08.2026"))
    print("Monivoi a SAP :", fecha_efectiva_monivoi_str("18.08.2026"))
    print("VBRL    día   :", nombre_archivo_dia_vbrl("11.08.2026"))
    print("VBRL    rango :", nombre_archivo_rango_vbrl("11.08.2026", "13.08.2026"))
    print("EIAP    rango :", nombre_archivo_rango_eiap("01.01.2026", "10.01.2026"))
    print("EIAP    a SAP :", fecha_efectiva_eiap_str("10.01.2026"))
