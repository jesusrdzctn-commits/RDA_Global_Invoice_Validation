"""
Validación Factura Global — Descargas SAP
==========================================

Módulo de extracción de transacciones SAP para el proyecto
'Validación Factura Global'.

Primera transacción implementada: 'Gallo' (FAGLL03).
Basado en la grabación VBS original, con tres mejoras:
  1. Las fechas las ingresa el usuario (ya no van fijas a 14.06.2026).
  2. El nombre del archivo se calcula como (fecha_hasta + 1 día),
     en formato DD_mmm_YY.xlsx (mes en español).  Ej: 14.06.2026 -> 15_jun_26.xlsx
  3. La ruta y el nombre del archivo se escriben directo en el diálogo de
     exportación (más robusto que navegar con F4 como hacía el VBS).
"""

import os
import time
import traceback
from datetime import datetime, timedelta

import openpyxl as _oxl          # asegura que PyInstaller incluya el motor de Excel
import win32com.client
import pywintypes


# ----------------------------------------------------------------------
# Constantes de la transacción (tomadas de la grabación VBS original)
# ----------------------------------------------------------------------

# Mes en español, 3 letras — para nombrar el archivo de salida.
_MESES_ES = {
    1: "ene", 2: "feb", 3: "mar", 4: "abr", 5: "may", 6: "jun",
    7: "jul", 8: "ago", 9: "sep", 10: "oct", 11: "nov", 12: "dic",
}

# Cuentas que vienen en el multiple-selection (SD_SAKNR) del VBS.
_CUENTAS_GALLO = [
    "4000005", "4150000", "4200000", "4200004",
    "4200015", "4260000", "7000005",
]

# Sociedad por defecto (BUKRS) que trae el VBS.
_SOCIEDAD_DEFAULT = "MX21"

# Layout / variante de visualización por defecto (PA_VARI).
_VARIANTE_DEFAULT = "BASE/VALID"


# ----------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------
def ruta_input_default():
    """
    Devuelve la carpeta de entrada del proyecto:
        C:\\Users\\<usuario>\\Documents\\Validacion Factura Global\\src\\Input

    Usa el usuario del sistema, así corre en cualquier equipo sin tocar el código.
    """
    user_profile = os.environ.get("USERPROFILE") or os.path.expanduser("~")
    return os.path.join(
        user_profile, "Documents", "Validacion Factura Global", "src", "Input"
    )


def nombre_archivo_gallo(date_high):
    """
    Recibe la fecha HIGH del rango ('DD.MM.YYYY') y devuelve el nombre del
    archivo nombrado con el día SIGUIENTE, en formato DD_mmm_YY.xlsx.

    Ej: '14.06.2026' -> '15_jun_26.xlsx'

    (Si prefieres el día con cero a la izquierda, cambia '{fecha.day}' por
     '{fecha.day:02d}' en el return → daría '05_ene_26' en lugar de '5_ene_26'.)
    """
    fecha = datetime.strptime(date_high, "%d.%m.%Y") + timedelta(days=1)
    return f"{fecha.day}_{_MESES_ES[fecha.month]}_{fecha:%y}.xlsx"


def _verificar_sin_partidas(session, ruta_archivo):
    """
    Detecta el mensaje SAP de 'No se ha seleccionado ninguna partida'.
    Si aparece: cierra el popup, crea un Excel vacío como placeholder y
    devuelve True para que el llamador sepa que no hay datos.
    """
    try:
        sin_partidas = False
        MENSAJES_SIN_DATOS = ("MSITEM033", "MSITEM030", "ninguna partida", "ninguna cuenta")

        try:
            sbar_text = session.findById("wnd[0]/sbar").Text.strip()
        except Exception:
            sbar_text = ""

        if any(m in sbar_text or m in sbar_text.lower() for m in MENSAJES_SIN_DATOS):
            sin_partidas = True

        if not sin_partidas:
            try:
                popup_text = session.findById("wnd[1]/usr/txtMESSTXT1").Text.strip()
                if any(m in popup_text or m in popup_text.lower() for m in MENSAJES_SIN_DATOS):
                    sin_partidas = True
            except Exception:
                pass

        if sin_partidas:
            print(f"[GALLO] Sin partidas → creando archivo vacío: {ruta_archivo}")
            for btn in ("wnd[1]/tbar[0]/btn[0]", "wnd[0]/tbar[0]/btn[3]"):
                try:
                    session.findById(btn).press()
                except Exception:
                    pass
            try:
                carpeta = os.path.dirname(ruta_archivo)
                if carpeta:
                    os.makedirs(carpeta, exist_ok=True)
                wb = _oxl.Workbook()
                wb.active.title = "Sin datos"
                wb.save(ruta_archivo)
                wb.close()
            except Exception as e:
                raise RuntimeError(f"No se pudo crear el archivo vacío: {ruta_archivo}") from e
            return True

        return False

    except Exception as e:
        print(f"[ERROR] _verificar_sin_partidas: {e}")
        return False


# ----------------------------------------------------------------------
# Transacción Gallo (FAGLL03)
# ----------------------------------------------------------------------
def Gallo_FAGLL03(
    DateFrom,
    DateTo,
    FolderPath=None,
    FileName=None,
    sociedad=_SOCIEDAD_DEFAULT,
    cuentas=None,
    variante=_VARIANTE_DEFAULT,
):
    """
    Descarga la transacción FAGLL03 ('Gallo') desde SAP a un archivo .xlsx.

    Args:
        DateFrom (str): Fecha inicial del rango, formato 'DD.MM.YYYY'.
        DateTo   (str): Fecha final del rango,   formato 'DD.MM.YYYY'.
        FolderPath (str|None): Carpeta destino. Si None → ruta_input_default().
        FileName   (str|None): Nombre del archivo. Si None → (DateTo + 1 día).
        sociedad   (str): Sociedad (BUKRS). Default 'MX21'.
        cuentas    (list|None): Cuentas a filtrar. Si None → las del VBS.
        variante   (str|None): Nombre del layout/variante (PA_VARI).
            Default 'BASE/VALID'.
            - Si se proporciona → se escribe directo (ROBUSTO, recomendado).
            - Si es None        → se replica la navegación F4 grabada (FRÁGIL).

    Returns:
        bool: True si descargó con datos, False si no hubo datos o error.
    """
    session = None
    cuentas       = cuentas or _CUENTAS_GALLO
    FolderPath    = FolderPath or ruta_input_default()
    FileName      = FileName or nombre_archivo_gallo(DateTo)
    ruta_completa = os.path.join(FolderPath, FileName)

    try:
        # === Validaciones mínimas ===
        for etiqueta, fecha in (("DateFrom", DateFrom), ("DateTo", DateTo)):
            datetime.strptime(fecha, "%d.%m.%Y")  # lanza ValueError si el formato es inválido

        os.makedirs(FolderPath, exist_ok=True)

        # === Conexión SAP ===
        try:
            SapGuiAuto  = win32com.client.GetObject("SAPGUI")
            application = SapGuiAuto.GetScriptingEngine
            connection  = application.Children(0)
            session     = connection.Children(0)
        except Exception as e:
            raise ConnectionError(
                "No fue posible conectarse a SAP GUI. "
                "Verifica que SAP esté abierto y con una sesión activa."
            ) from e

        # === Abrir transacción FAGLL03 ===
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nFAGLL03"
        session.findById("wnd[0]").sendVKey(0)

        # === Cargar cuentas (multiple selection SD_SAKNR) ===
        session.findById("wnd[0]/usr/btn%_SD_SAKNR_%_APP_%-VALU_PUSH").press()
        for i, cuenta in enumerate(cuentas, start=0):
            cuenta = str(cuenta).strip()
            field_id = (
                f"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/"
                f"ssubSCREEN_HEADER:SAPLALDB:3010/"
                f"tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,{i}]"
            )
            campo = session.findById(field_id)
            campo.text          = cuenta
            campo.setFocus()
            campo.caretPosition = len(cuenta)
        session.findById("wnd[1]/tbar[0]/btn[0]").press()  # copiar/adoptar
        session.findById("wnd[1]/tbar[0]/btn[8]").press()  # ejecutar y cerrar popup

        # === Parámetros de selección ===
        session.findById("wnd[0]/usr/radX_AISEL").select()          # todas las partidas
        session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").text  = sociedad
        session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text  = DateFrom
        session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = DateTo

        # === Layout / variante (PA_VARI) ===
        if variante:
            # Camino ROBUSTO: escribir el nombre del layout directamente.
            campo_vari = session.findById("wnd[0]/usr/ctxtPA_VARI")
            campo_vari.text          = variante
            campo_vari.setFocus()
            campo_vari.caretPosition = len(variante)
        else:
            # Camino GRABADO (frágil): F4 + scroll + clic por coordenada.
            # Depende de la posición exacta del layout en la lista; si SAP
            # cambia el orden o agregan layouts, esto puede fallar.
            session.findById("wnd[0]/usr/ctxtPA_VARI").setFocus()
            session.findById("wnd[0]/usr/ctxtPA_VARI").caretPosition = 9
            session.findById("wnd[0]").sendVKey(4)
            session.findById("wnd[1]/usr").verticalScrollbar.position = 991
            session.findById("wnd[1]/usr/lbl[1,24]").setFocus()
            session.findById("wnd[1]/usr/lbl[1,24]").caretPosition = 6
            session.findById("wnd[1]").sendVKey(2)

        # === Ejecutar reporte ===
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(4)

        # === ¿Sin partidas? ===
        if _verificar_sin_partidas(session, ruta_completa):
            return False  # archivo vacío ya creado

        # === Exportar a Excel (forma robusta: path + filename directos) ===
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text     = FolderPath
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = FileName
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(FileName)
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        # === Regresar a la pantalla inicial ===
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()

        print(f"[GALLO] Archivo descargado correctamente: {ruta_completa}")
        return True

    except pywintypes.com_error as e:
        print(f"[ERROR COM SAP] Gallo_FAGLL03: {e}")
        return False
    except (ValueError, TypeError, ConnectionError, RuntimeError) as e:
        print(f"[ERROR CONTROLADO] Gallo_FAGLL03: {e}")
        return False
    except Exception as e:
        print(f"[ERROR NO CONTROLADO] Gallo_FAGLL03: {e}")
        print(traceback.format_exc())
        return False
    finally:
        if session is not None:
            try:
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
            except Exception:
                pass


# ----------------------------------------------------------------------
# Ejecución directa (prueba rápida sin GUI ni controller)
# ----------------------------------------------------------------------
if __name__ == "__main__":
    print("=== Prueba descarga Gallo (FAGLL03) ===")
    date_from = input("Fecha desde (DD.MM.YYYY): ").strip()
    date_to   = input("Fecha hasta (DD.MM.YYYY): ").strip()

    archivo = nombre_archivo_gallo(date_to)
    print(f"\nEl archivo se nombrará : {archivo}")
    print(f"Se guardará en         : {ruta_input_default()}\n")

    ok = Gallo_FAGLL03(date_from, date_to)
    print("✅ Descarga con datos." if ok else "⚠️ Sin datos o error (revisa la consola).")
