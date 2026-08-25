"""
Validación Factura Global — Descargas SAP
==========================================

Módulo de extracción de transacciones SAP para el proyecto
'Validación Factura Global'.

Transacciones implementadas:
  - 'Gallo'   (FAGLL03)
  - 'Monivoi' (ZLMXCOM_TRN_MONINVOI)

Basado en las grabaciones VBS originales, con estas mejoras:
  1. Las fechas las ingresa el usuario (ya no van fijas).
  2. El nombre del archivo usa la MISMA fecha del usuario, en formato
     DD_mmm_YY (mes en español).
  3. Soporta descarga de un solo día o de un rango (día por día).
  4. La ruta y el nombre del archivo se escriben directo en el diálogo de
     exportación (más robusto que navegar con F4 como hacía el VBS).
"""

import os
import time
import traceback
from datetime import datetime

import openpyxl as _oxl          # asegura que PyInstaller incluya el motor de Excel
import pandas as pd
import win32com.client
import win32clipboard         
import pywintypes

from utils import (
    nombre_archivo_dia,          nombre_archivo_rango,
    nombre_archivo_dia_monivoi,  nombre_archivo_rango_monivoi,
    fecha_efectiva_monivoi_str,
)

# ----------------------------------------------------------------------
# Constantes de la transacción (tomadas de la grabación VBS original)
# ----------------------------------------------------------------------

# Cuentas que vienen en el multiple-selection (SD_SAKNR) del VBS.
_CUENTAS_GALLO = [
    "4000005", "4000010", "4000020", "4000802",
    "4150000", "4200000", "4200004", "4200015",
    "4260000", "4300013", "7000005",
]

# Sociedad por defecto (BUKRS) que trae el VBS.
_SOCIEDAD_DEFAULT = "MX21"

# Layout / variante de visualización por defecto (PA_VARI).
_VARIANTE_DEFAULT = "BASE/VALID"


# ----------------------------------------------------------------------
# Constantes de la transacción Monivoi (ZLMXCOM_TRN_MONINVOI)
# ----------------------------------------------------------------------
_TX_MONIVOI    = "ZLMXCOM_TRN_MONINVOI"   # transacción custom (Z)
_VKORG_DEFAULT = "MX21"                    # organización de ventas
_VTWEG_DEFAULT = "01"                      # canal de distribución
_RFC_DEFAULT   = "XAXX010101000"           # RFC genérico ("público en general")


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


def _xlsx_a_csv(ruta_xlsx, ruta_csv, eliminar_xlsx=True):
    """
    Convierte a CSV (utf-8-sig) el .xlsx que SAP genera al exportar, para que
    TODOS los archivos descargados queden en el mismo formato. Luego borra el
    .xlsx intermedio. utf-8-sig hace que Excel muestre bien los acentos.
    """
    df = pd.read_excel(ruta_xlsx, dtype=str, engine="openpyxl")
    df.to_csv(ruta_csv, index=False, encoding="utf-8-sig")
    if eliminar_xlsx:
        _borrar_xlsx_con_reintentos(ruta_xlsx)


def _borrar_xlsx_con_reintentos(ruta_xlsx, intentos=5, espera=2):
    """
    Borra el .xlsx intermedio. En Windows, si Excel todavía lo tiene abierto,
    el archivo queda BLOQUEADO y os.remove falla. Por eso en cada intento
    primero re-intentamos cerrar el workbook en Excel y luego borrar, dándole
    tiempo a que suelte el archivo.
    """
    for intento in range(1, intentos + 1):
        _cerrar_excel_workbook(ruta_xlsx)        # reintenta cerrarlo en Excel
        try:
            if os.path.exists(ruta_xlsx):
                os.remove(ruta_xlsx)
            print(f"[XLSX] Temporal borrado: {ruta_xlsx}")
            return True
        except OSError as e:
            print(f"[XLSX] Aún bloqueado (intento {intento}/{intentos}): {e}")
            time.sleep(espera)

    print(f"[XLSX] ⚠️ No se pudo borrar (sigue abierto/bloqueado): {ruta_xlsx}")
    return False


def _cerrar_excel_workbook(ruta_archivo, intentos=3, espera=2):
    """
    Cierra (sin guardar) el workbook que SAP/Excel deja abierto tras la
    exportación, para que el archivo no quede abierto ni ocupando memoria.
    Reintenta un par de veces por si Excel aún no terminaba de abrirlo.
    Falla en silencio si Excel no está corriendo.
    """
    ruta_abs = os.path.abspath(ruta_archivo)
    for _ in range(intentos):
        try:
            excel = win32com.client.GetObject(Class="Excel.Application")
        except Exception:
            return  # no hay Excel abierto → nada que cerrar
        cerrado = False
        try:
            excel.DisplayAlerts = False
        except Exception:
            pass
        try:
            for wb in list(excel.Workbooks):
                try:
                    if os.path.abspath(wb.FullName) == ruta_abs:
                        wb.Close(SaveChanges=False)
                        cerrado = True
                        print(f"[EXCEL] Workbook cerrado: {ruta_archivo}")
                        break
                except Exception:
                    continue
        finally:
            try:
                excel.DisplayAlerts = True
            except Exception:
                pass
        if cerrado:
            return
        time.sleep(espera)  # tal vez aún no terminaba de abrir; reintenta


def _verificar_sin_partidas(session, ruta_archivo):
    """
    Detecta el mensaje SAP de 'No se ha seleccionado ninguna partida'.
    Si aparece: cierra el popup, crea un Excel/CSV vacío como placeholder y
    devuelve True para que el llamador sepa que no hay datos.

    NOTA: los códigos de mensaje están afinados para FAGLL03 (Gallo). Para
    Monivoi (Z) el texto de 'sin datos' puede ser distinto; si en tus pruebas
    ves que no lo detecta, agrega el mensaje exacto a MENSAJES_SIN_DATOS.
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
            print(f"[SIN DATOS] Creando archivo vacío: {ruta_archivo}")
            for btn in ("wnd[1]/tbar[0]/btn[0]", "wnd[0]/tbar[0]/btn[3]"):
                try:
                    session.findById(btn).press()
                except Exception:
                    pass
            try:
                carpeta = os.path.dirname(ruta_archivo)
                if carpeta:
                    os.makedirs(carpeta, exist_ok=True)
                if ruta_archivo.lower().endswith(".csv"):
                    pd.DataFrame({"Sin datos": []}).to_csv(
                        ruta_archivo, index=False, encoding="utf-8-sig"
                    )
                else:
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

def _copiar_al_portapapeles(texto, intentos=5, espera=0.3):
    """
    Pone 'texto' en el portapapeles de Windows (Unicode). Reintenta unas veces
    porque a veces otra app tiene el portapapeles ocupado un instante.
    OJO: esto SOBRESCRIBE lo que el usuario tuviera copiado en ese momento.
    """
    ultimo_error = None
    for _ in range(intentos):
        try:
            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardText(texto, win32clipboard.CF_UNICODETEXT)
                return True
            finally:
                win32clipboard.CloseClipboard()
        except Exception as e:
            ultimo_error = e
            time.sleep(espera)
    raise RuntimeError(f"No se pudo escribir en el portapapeles: {ultimo_error}")


def _importar_cuentas_desde_popup(session):
    """
    Pulsa el botón 'Importar desde portapapeles' del popup de selección
    múltiple. Prueba los IDs conocidos por si cambia entre sistemas SAP.
    """
    for btn_id in ("wnd[1]/tbar[0]/btn[24]", "wnd[1]/tbar[0]/btn[25]"):
        try:
            session.findById(btn_id).press()
            return True
        except Exception:
            continue
    raise RuntimeError(
        "No encontré el botón 'Importar desde portapapeles' en el popup. "
        "Confirma su ID con Script Recording (Alt+F12) y agrégalo a la lista."
    )

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
    Descarga la transacción FAGLL03 ('Gallo') desde SAP a un archivo .csv.

    Args:
        DateFrom (str): Fecha inicial del rango, formato 'DD.MM.YYYY'.
        DateTo   (str): Fecha final del rango,   formato 'DD.MM.YYYY'.
        FolderPath (str|None): Carpeta destino. Si None → ruta_input_default().
        FileName   (str|None): Nombre del archivo. Si None → nombre_archivo_dia(DateTo).
        sociedad   (str): Sociedad (BUKRS). Default 'MX21'.
        cuentas    (list|None): Cuentas a filtrar. Si None → las del VBS.
        variante   (str|None): Nombre del layout/variante (PA_VARI).
            Default 'BASE/VALID'. Si llega None o vacío, usa el default.
            Siempre se escribe por nombre (ROBUSTO).

    Returns:
        bool: True si descargó con datos, False si no hubo datos o error.
    """
    session = None
    cuentas       = cuentas or _CUENTAS_GALLO
    FolderPath    = FolderPath or ruta_input_default()
    FileName      = FileName or nombre_archivo_dia(DateTo)
    ruta_completa = os.path.join(FolderPath, FileName)   # destino final (puede ser .csv)

    # SAP sólo exporta en formato hoja de cálculo (.xlsx). Si el archivo final
    # que queremos es .csv, exportamos primero a un .xlsx temporal y al terminar
    # lo convertimos. Así TODOS los archivos quedan en el mismo formato.
    _base, _ext     = os.path.splitext(FileName)
    _exportar_a_csv = _ext.lower() == ".csv"
    _nombre_xlsx    = (_base + ".xlsx") if _exportar_a_csv else FileName
    _ruta_xlsx      = os.path.join(FolderPath, _nombre_xlsx)

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

        # === Cargar cuentas (multiple selection SD_SAKNR) vía PORTAPAPELES ===
        # En vez de escribir fila por fila (y tener que lidiar con el scroll cuando
        # hay muchas cuentas), copiamos TODAS las cuentas al portapapeles —una por
        # línea— y usamos el botón 'Importar desde portapapeles' del popup. Así no
        # importa cuántas cuentas sean: entran de un solo jalón.
        session.findById("wnd[0]/usr/btn%_SD_SAKNR_%_APP_%-VALU_PUSH").press()

        _copiar_al_portapapeles("\r\n".join(str(c).strip() for c in cuentas))
        _importar_cuentas_desde_popup(session)
        time.sleep(1)                                       # deja que SAP pinte las cuentas

        session.findById("wnd[1]/tbar[0]/btn[0]").press()  # copiar/adoptar
        session.findById("wnd[1]/tbar[0]/btn[8]").press()  # ejecutar y cerrar popup

         # === Parámetros de selección ===
        session.findById("wnd[0]/usr/radX_AISEL").select()          # todas las partidas
        session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").text  = sociedad
        session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text  = DateFrom
        session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = DateTo

        # Layout/variante SIEMPRE por nombre (robusto). Si llega vacío o None,
        # usa el default en lugar del camino frágil F4+scroll+coordenada.
        variante = variante or _VARIANTE_DEFAULT
        campo_vari = session.findById("wnd[0]/usr/ctxtPA_VARI")
        campo_vari.text          = variante
        campo_vari.setFocus()
        campo_vari.caretPosition = len(variante)

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
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = _nombre_xlsx
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(_nombre_xlsx)
        session.findById("wnd[1]/tbar[0]/btn[11]").press()

        # === Cerrar el Excel que SAP abre tras exportar ===
        time.sleep(3)
        _cerrar_excel_workbook(_ruta_xlsx)

        # === Convertir el .xlsx temporal a .csv (consistencia de formato) ===
        if _exportar_a_csv:
            _xlsx_a_csv(_ruta_xlsx, ruta_completa)

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
# Transacción Monivoi (ZLMXCOM_TRN_MONINVOI)
# ----------------------------------------------------------------------
def Monivoi_ZLMXCOM(
    DateFrom,
    DateTo,
    FolderPath=None,
    FileName=None,
    sociedad=_SOCIEDAD_DEFAULT,
    vkorg=_VKORG_DEFAULT,
    vtweg=_VTWEG_DEFAULT,
    rfc=_RFC_DEFAULT,
):
    """
    Descarga la transacción ZLMXCOM_TRN_MONINVOI ('Monivoi') desde SAP a .csv.

    Diferencias vs Gallo (FAGLL03):
      - Sociedad por P_BUKRS (no SD_BUKRS-LOW).
      - Filtra por fecha de CREACIÓN (SO_ERDAT), no de contabilización.
      - Lleva org. de ventas (VKORG), canal (VTWEG) y RFC (STCD1).
      - Sin cuentas ni layout/variante.
      - Al exportar, SAP muestra primero un selector de formato (cmbG_LISTBOX).

    Args:
        DateFrom, DateTo (str): rango de fechas 'DD.MM.YYYY' (ERDAT).
        FolderPath (str|None): carpeta destino. Si None → ruta_input_default().
        FileName   (str|None): nombre del archivo. Si None → nombre_archivo_dia_monivoi(DateTo).
        sociedad (str): P_BUKRS. Default 'MX21'.
        vkorg (str): organización de ventas (SO_VKORG). Default 'MX21'.
        vtweg (str): canal de distribución (SO_VTWEG). Default '01'.
        rfc   (str): RFC a filtrar (SO_STCD1). Default 'XAXX010101000'.

    Returns:
        bool: True si descargó con datos, False si no hubo datos o hubo error.
    """
    session = None
    FolderPath    = FolderPath or ruta_input_default()
    FileName      = FileName or nombre_archivo_dia_monivoi(DateTo)
    ruta_completa = os.path.join(FolderPath, FileName)

    # Igual que Gallo: SAP exporta a .xlsx; si el destino es .csv, exportamos a
    # un .xlsx temporal y al terminar lo convertimos.
    _base, _ext     = os.path.splitext(FileName)
    _exportar_a_csv = _ext.lower() == ".csv"
    _nombre_xlsx    = (_base + ".xlsx") if _exportar_a_csv else FileName
    _ruta_xlsx      = os.path.join(FolderPath, _nombre_xlsx)

    try:
        # === Validar fechas ===
        for _etq, fecha in (("DateFrom", DateFrom), ("DateTo", DateTo)):
            datetime.strptime(fecha, "%d.%m.%Y")

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

        # === Abrir la transacción (con /n para arrancar desde cualquier pantalla) ===
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n" + _TX_MONIVOI
        session.findById("wnd[0]").sendVKey(0)

        # === Parámetros de selección ===
        session.findById("wnd[0]/usr/ctxtP_BUKRS").text      = sociedad
        session.findById("wnd[0]/usr/ctxtSO_ERDAT-LOW").text  = fecha_efectiva_monivoi_str(DateFrom)
        session.findById("wnd[0]/usr/ctxtSO_ERDAT-HIGH").text = fecha_efectiva_monivoi_str(DateTo)
        session.findById("wnd[0]/usr/ctxtSO_VKORG-LOW").text  = vkorg
        session.findById("wnd[0]/usr/ctxtSO_VTWEG-LOW").text  = vtweg

        campo_rfc = session.findById("wnd[0]/usr/txtSO_STCD1-LOW")
        campo_rfc.text          = rfc
        campo_rfc.setFocus()
        campo_rfc.caretPosition = len(rfc)

        # === Ejecutar reporte ===
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(4)

        # === ¿Sin datos? (reusa el helper de Gallo; ver nota en la función) ===
        if _verificar_sin_partidas(session, ruta_completa):
            return False

        # === Exportar a hoja de cálculo ===
        # OJO: aquí Monivoi difiere de Gallo. SAP abre primero un diálogo con un
        # selector de formato (cmbG_LISTBOX); tomamos el valor por defecto y OK.
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        try:
            session.findById("wnd[1]/usr/cmbG_LISTBOX").setFocus()
        except Exception:
            pass  # por si en algún equipo no aparece el selector de formato
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # === Diálogo de guardar ===
        # El VBS no fijaba la ruta; nosotros SÍ, para dejar el archivo en Input.
        try:
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = FolderPath
        except Exception:
            pass
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = _nombre_xlsx
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(_nombre_xlsx)
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # === Cerrar el Excel que SAP abre + convertir a CSV ===
        time.sleep(3)
        _cerrar_excel_workbook(_ruta_xlsx)
        if _exportar_a_csv:
            _xlsx_a_csv(_ruta_xlsx, ruta_completa)

        # === Regresar a la pantalla inicial ===
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()

        print(f"[MONIVOI] Archivo descargado correctamente: {ruta_completa}")
        return True

    except pywintypes.com_error as e:
        print(f"[ERROR COM SAP] Monivoi_ZLMXCOM: {e}")
        return False
    except (ValueError, TypeError, ConnectionError, RuntimeError) as e:
        print(f"[ERROR CONTROLADO] Monivoi_ZLMXCOM: {e}")
        return False
    except Exception as e:
        print(f"[ERROR NO CONTROLADO] Monivoi_ZLMXCOM: {e}")
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
    print("=== Prueba de descarga (standalone) ===")
    print("1) Gallo (FAGLL03)")
    print("2) Monivoi (ZLMXCOM_TRN_MONINVOI)")
    opcion = input("Elige transacción [1/2]: ").strip()

    date_from = input("Fecha desde (DD.MM.YYYY): ").strip()
    date_to   = input("Fecha hasta (DD.MM.YYYY): ").strip()

    if opcion == "2":
        archivo = nombre_archivo_dia_monivoi(date_to)
        print(f"\nEl archivo se nombrará : {archivo}")
        print(f"Se guardará en         : {ruta_input_default()}\n")
        ok = Monivoi_ZLMXCOM(date_from, date_to)
    else:
        archivo = nombre_archivo_dia(date_to)
        print(f"\nEl archivo se nombrará : {archivo}")
        print(f"Se guardará en         : {ruta_input_default()}\n")
        ok = Gallo_FAGLL03(date_from, date_to)

    print("✅ Descarga con datos." if ok else "⚠️ Sin datos o error (revisa la consola).")
