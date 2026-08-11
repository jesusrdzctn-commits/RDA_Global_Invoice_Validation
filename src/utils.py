"""
Validación Factura Global — utils.py
====================================

Helpers compartidos SIN dependencias de SAP ni de tkinter, para que:
  - la GUI (interfaz_GUI.py) los use y siga siendo abrible/testeable sola, y
  - la capa SAP (Descargas_SAP.py) los use y no dupliquemos lógica.

Aquí vive la ÚNICA fuente de verdad para:
  - el mapeo de meses en español,
  - el formato de los nombres de archivo (Gallo y Monivoi), y
  - el desfase de "un día antes" de Monivoi  (¡el -1 vive SOLO aquí!).

Regla de oro: si algún día cambia el formato del nombre o el desfase de
Monivoi, se toca este archivo y nada más.
"""

from datetime import datetime, timedelta

# Formato de fecha que teclea el usuario en toda la app.
FORMATO_FECHA = "%d.%m.%Y"

# Mes en español, 3 letras — para nombrar los archivos de salida.
MESES_ES = {
    1: "ene", 2: "feb", 3: "mar", 4: "abr", 5: "may", 6: "jun",
    7: "jul", 8: "ago", 9: "sep", 10: "oct", 11: "nov", 12: "dic",
}


# ----------------------------------------------------------------------
# Fechas
# ----------------------------------------------------------------------
def parse_fecha(fecha_str):
    """'DD.MM.YYYY' -> datetime. Lanza ValueError si el formato es inválido."""
    return datetime.strptime(fecha_str.strip(), FORMATO_FECHA)


def fecha_efectiva_monivoi(fecha_str):
    """
    Monivoi trabaja con el día ANTERIOR al que teclea el usuario.
    Este es el ÚNICO lugar donde se aplica ese -1: tanto los nombres de archivo
    como (si quieres) la fecha que se manda a SAP deben pasar por aquí.

        '11.08.2026' -> datetime(2026, 8, 10)
    """
    return parse_fecha(fecha_str) - timedelta(days=1)


def fecha_efectiva_monivoi_str(fecha_str):
    """
    Igual que fecha_efectiva_monivoi() pero devuelve el string 'DD.MM.YYYY'
    listo para escribirlo en los campos de SAP (SO_ERDAT).

        '11.08.2026' -> '10.08.2026'
    """
    return fecha_efectiva_monivoi(fecha_str).strftime(FORMATO_FECHA)


def _sufijo_dia(f):
    """datetime -> 'D_mmm_YY' (ej: 10_ago_26). Uso interno."""
    return f"{f.day}_{MESES_ES[f.month]}_{f:%y}"


# ----------------------------------------------------------------------
# Nombres de archivo — Gallo (FAGLL03)   [sin desfase]
# ----------------------------------------------------------------------
def nombre_archivo_dia(fecha_str):
    """'15.06.2026' -> '15_jun_26.csv'."""
    return f"{_sufijo_dia(parse_fecha(fecha_str))}.csv"


def nombre_archivo_rango(desde, hasta):
    """('15.06.2026', '17.06.2026') -> '15_jun_26_a_17_jun_26.csv'."""
    d = _sufijo_dia(parse_fecha(desde))
    h = _sufijo_dia(parse_fecha(hasta))
    return f"{d}_a_{h}.csv"


# ----------------------------------------------------------------------
# Nombres de archivo — Monivoi (ZLMXCOM_TRN_MONINVOI)   [día -1 incluido]
# El -1 ya viene aplicado vía fecha_efectiva_monivoi(): no lo repitas afuera.
# ----------------------------------------------------------------------
def nombre_archivo_dia_monivoi(fecha_str):
    """'11.08.2026' -> 'Monitoreo_10_ago_26.csv'  (día -1)."""
    return f"Monitoreo_{_sufijo_dia(fecha_efectiva_monivoi(fecha_str))}.csv"


def nombre_archivo_rango_monivoi(desde, hasta):
    """('11.08.2026', '13.08.2026') -> 'Monitoreo_10_ago_26_a_12_ago_26.csv'."""
    d = _sufijo_dia(fecha_efectiva_monivoi(desde))
    h = _sufijo_dia(fecha_efectiva_monivoi(hasta))
    return f"Monitoreo_{d}_a_{h}.csv"


# ----------------------------------------------------------------------
# Prueba rápida:  python utils.py
# ----------------------------------------------------------------------
if __name__ == "__main__":
    print("Gallo   día   :", nombre_archivo_dia("15.06.2026"))
    print("Gallo   rango :", nombre_archivo_rango("15.06.2026", "17.06.2026"))
    print("Monivoi día   :", nombre_archivo_dia_monivoi("11.08.2026"))
    print("Monivoi rango :", nombre_archivo_rango_monivoi("11.08.2026", "13.08.2026"))
    print("Monivoi a SAP :", fecha_efectiva_monivoi_str("11.08.2026"))
