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

import ctypes
from ctypes import wintypes

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

def _terminar_proceso(pid):
    """Cierra por la fuerza el proceso 'pid' (Windows). True si lo logró."""
    PROCESS_TERMINATE = 0x0001
    k32 = ctypes.windll.kernel32
    k32.OpenProcess.restype  = wintypes.HANDLE
    k32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    k32.TerminateProcess.argtypes = [wintypes.HANDLE, wintypes.UINT]
    k32.CloseHandle.argtypes      = [wintypes.HANDLE]

    handle = k32.OpenProcess(PROCESS_TERMINATE, False, int(pid))
    if not handle:
        return False
    try:
        return bool(k32.TerminateProcess(handle, 1))
    finally:
        k32.CloseHandle(handle)


def _liberar_archivo_con_restart_manager(ruta_archivo):
    """
    Último recurso cuando el .xlsx queda BLOQUEADO y COM no lo pudo cerrar.

    ¿Por qué? En Gallo, Excel abre el archivo exportado y a veces se queda con
    un DIÁLOGO modal esperando un clic (el ícono parpadea en DORADO: Windows
    diciendo "esta ventana quiere tu atención"). Un Excel en modo modal NO
    responde a COM, pero SÍ mantiene el archivo agarrado → os.remove falla.

    Este helper usa el Windows Restart Manager (la misma API del "este archivo
    está siendo usado por: Excel"). Pregunta QUÉ procesos tienen agarrado ESTE
    archivo en específico y cierra SÓLO los que sean EXCEL.EXE. Jamás toca SAP:
      - si el que lo tiene es SAP (caso 'in place', como Monivoi) → no hace nada;
      - si es un Excel standalone atorado en un diálogo (Gallo) → lo cierra.

    Seguro: los datos YA están en disco (por eso el .csv se genera bien); ese
    Excel solo MUESTRA el archivo, no tiene cambios que valga la pena guardar.
    """
    try:
        rstrtmgr = ctypes.WinDLL("rstrtmgr")

        class RM_UNIQUE_PROCESS(ctypes.Structure):
            _fields_ = [("dwProcessId", wintypes.DWORD),
                        ("ProcessStartTime", wintypes.FILETIME)]

        CCH_APP = 255
        CCH_SVC = 63

        class RM_PROCESS_INFO(ctypes.Structure):
            _fields_ = [
                ("Process", RM_UNIQUE_PROCESS),
                ("strAppName", wintypes.WCHAR * (CCH_APP + 1)),
                ("strServiceShortName", wintypes.WCHAR * (CCH_SVC + 1)),
                ("ApplicationType", ctypes.c_int),
                ("AppStatus", wintypes.DWORD),
                ("TSSessionId", wintypes.DWORD),
                ("bRestartable", wintypes.BOOL),
            ]

        sesion = wintypes.DWORD(0)
        clave  = (ctypes.c_wchar * (32 + 1))()          # CCH_RM_SESSION_KEY = 32
        if rstrtmgr.RmStartSession(ctypes.byref(sesion), 0, clave) != 0:
            print("[RM] No se pudo iniciar la sesión de Restart Manager.")
            return False

        liberado = False
        try:
            archivos = (ctypes.c_wchar_p * 1)(os.path.abspath(ruta_archivo))
            if rstrtmgr.RmRegisterResources(
                sesion, 1, archivos, 0, None, 0, None
            ) != 0:
                print("[RM] No se pudo registrar el archivo.")
                return False

            necesarios = wintypes.UINT(0)
            cuantos    = wintypes.UINT(0)
            razon      = wintypes.DWORD(0)

            # 1er llamado: ¿cuántos procesos? (devuelve ERROR_MORE_DATA)
            rstrtmgr.RmGetList(sesion, ctypes.byref(necesarios),
                               ctypes.byref(cuantos), None, ctypes.byref(razon))

            if necesarios.value == 0:
                print(f"[RM] Ningún proceso tiene agarrado: {ruta_archivo}")
                return False

            cuantos = wintypes.UINT(necesarios.value)
            infos   = (RM_PROCESS_INFO * necesarios.value)()
            if rstrtmgr.RmGetList(sesion, ctypes.byref(necesarios),
                                  ctypes.byref(cuantos), infos,
                                  ctypes.byref(razon)) != 0:
                print("[RM] No se pudo obtener la lista de procesos.")
                return False

            for i in range(cuantos.value):
                nombre = infos[i].strAppName or "<desconocido>"
                pid    = int(infos[i].Process.dwProcessId)
                print(f"[RM] '{os.path.basename(ruta_archivo)}' lo tiene: "
                      f"{nombre} (PID {pid})")
                # SÓLO Excel; nunca SAP ni otros procesos.
                if "excel" in nombre.lower():
                    if _terminar_proceso(pid):
                        print(f"[RM] Excel cerrado (PID {pid}) → archivo liberado.")
                        liberado = True
                    else:
                        print(f"[RM] No se pudo cerrar el PID {pid}.")
        finally:
            rstrtmgr.RmEndSession(sesion)

        return liberado

    except Exception as e:
        print(f"[RM] Falló el Restart Manager: {e}")
        return False

def _borrar_xlsx_con_reintentos(ruta_xlsx, intentos=10, espera=2):
    """
    Borra el .xlsx intermedio. En Windows, si Excel todavía lo tiene abierto,
    el archivo queda BLOQUEADO y os.remove falla. Por eso en cada intento
    primero re-intentamos cerrar el archivo en Excel y luego borrar, dándole
    tiempo a que lo suelte.

    Ventana amplia (intentos=10) a propósito: Gallo es el reporte grande y el
    primero de la corrida, así que arranca Excel "en frío" y puede tardar varios
    segundos en terminar de abrir el archivo; hasta que no lo abre del todo, no
    lo podemos cerrar. Sale en cuanto logra borrar, así que en el caso normal es
    rápido.
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

            # Si COM no lo suelta (típico de Gallo: Excel atorado en un diálogo
            # DORADO que COM no puede cerrar), escalamos al Restart Manager para
            # cerrar SÓLO ese Excel y liberar el archivo. Lo hacemos a la mitad
            # de los intentos, no en el primero, por si se libera solo.
            if intento == max(1, intentos // 2):
                _liberar_archivo_con_restart_manager(ruta_xlsx)

            time.sleep(espera)

    # Si llegamos aquí, no se pudo borrar: volcamos qué tiene Excel abierto.
    print(f"[XLSX] ⚠️ No se pudo borrar (sigue abierto/bloqueado): {ruta_xlsx}")
    _diagnostico_excel(ruta_xlsx)
    return False


def _instancias_excel_abiertas():
    """
    Devuelve TODAS las instancias de Excel abiertas en el equipo, no sólo la
    primera. Esto es CLAVE para el bug de "el .xlsx se queda abierto y no se
    puede borrar": cuando SAP exporta, abre el archivo en Excel, y muchas veces
    lo hace en una instancia DISTINTA a la que el usuario ya tenía abierta.
    win32com.GetObject(Class="Excel.Application") sólo alcanza UNA instancia, así
    que si el libro está en otra, nunca lo cerramos.

    Para verlas todas escaneamos la Running Object Table (ROT) de Windows: ahí
    cada libro abierto aparece registrado por su ruta; de cada libro tomamos su
    .Application y las juntamos sin repetir (usando el Hwnd como identidad).

    Degrada con elegancia: si el escaneo de la ROT falla por lo que sea, al menos
    devuelve la instancia principal vía GetObject.
    """
    instancias = {}

    # 1) Vía ROT: descubre TODAS las instancias a partir de los libros abiertos.
    try:
        import pythoncom
        rot = pythoncom.GetRunningObjectTable()
        ctx = pythoncom.CreateBindCtx(0)
        for moniker in rot.EnumRunning():
            try:
                nombre = moniker.GetDisplayName(ctx, None)
            except Exception:
                continue
            if not nombre:
                continue
            # Los libros abiertos se registran por su ruta de archivo.
            if not nombre.lower().endswith((".xls", ".xlsx", ".xlsm", ".xlsb", ".csv")):
                continue
            try:
                obj = rot.GetObject(moniker)
                wb  = win32com.client.Dispatch(obj)
                app = wb.Application
                instancias[int(app.Hwnd)] = app
            except Exception:
                continue
    except Exception:
        pass

    # 2) Fallback / complemento: la instancia principal registrada.
    try:
        app = win32com.client.GetObject(Class="Excel.Application")
        instancias[int(app.Hwnd)] = app
    except Exception:
        pass

    return list(instancias.values())


def _cerrar_excel_workbook(ruta_archivo, intentos=3, espera=1):
    """
    Cierra (sin guardar) el archivo que SAP/Excel deja abierto tras la
    exportación, para que no quede bloqueado (si sigue abierto, Windows no deja
    borrar el .xlsx temporal).

    Barre TODAS las instancias de Excel (ver _instancias_excel_abiertas), no sólo
    una, porque SAP suele abrir el archivo en una instancia aparte. En cada
    instancia busca el archivo en DOS lugares:
      A) Los libros normales (colección Workbooks), calzando por RUTA COMPLETA y,
         si falla, por NOMBRE de archivo (SAP a veces reporta la ruta con otra
         forma: mayúsculas, ruta corta 8.3, etc.).
      B) Las ventanas en VISTA PROTEGIDA (ProtectedViewWindows). ¡Esto es clave
         para Gallo! Excel abre en "Vista protegida" los archivos recién
         exportados/grandes, y esas ventanas NO aparecen en Workbooks, así que
         antes ni las tocábamos y el archivo quedaba bloqueado.

    Si tras cerrar una instancia queda SIN libros ni vistas protegidas, la cierra
    para no dejar una ventana vacía. NUNCA cierra una instancia que todavía tenga
    otros libros del usuario.

    Reintenta unas veces por si Excel aún no terminaba de abrir el archivo (Gallo,
    al ser el reporte grande y el primero, arranca Excel "en frío" y tarda más).
    Devuelve True si logró cerrar algo que calzaba; False si no encontró nada.
    """
    ruta_abs      = os.path.normcase(os.path.abspath(ruta_archivo))
    objetivo_base = os.path.normcase(os.path.basename(ruta_archivo))

    for _ in range(intentos):
        instancias = _instancias_excel_abiertas()
        if not instancias:
            return False  # no hay Excel abierto → nada que cerrar

        cerrado = False
        for excel in instancias:
            try:
                excel.DisplayAlerts = False
            except Exception:
                pass

            try:
                # === A) Libros normales ===
                for wb in list(excel.Workbooks):
                    try:
                        calza = (
                            os.path.normcase(os.path.abspath(wb.FullName)) == ruta_abs
                            or os.path.normcase(str(wb.Name)) == objetivo_base
                        )
                        if calza:
                            wb.Close(SaveChanges=False)
                            cerrado = True
                            print(f"[EXCEL] Workbook cerrado: {ruta_archivo}")
                    except Exception:
                        continue

                # === B) Ventanas en Vista Protegida (no están en Workbooks) ===
                try:
                    for pvw in list(excel.ProtectedViewWindows):
                        try:
                            info = ""
                            for prop in ("SourceName", "Caption"):
                                try:
                                    info += os.path.normcase(str(getattr(pvw, prop)))
                                except Exception:
                                    pass
                            if objetivo_base in info:
                                pvw.Close()
                                cerrado = True
                                print(f"[EXCEL] Vista protegida cerrada: {ruta_archivo}")
                        except Exception:
                            continue
                except Exception:
                    pass  # la instancia puede no exponer ProtectedViewWindows

                # === Si la instancia quedó totalmente vacía, ciérrala ===
                try:
                    sin_libros = int(excel.Workbooks.Count) == 0
                    sin_vista_prot = True
                    try:
                        sin_vista_prot = int(excel.ProtectedViewWindows.Count) == 0
                    except Exception:
                        pass
                    if sin_libros and sin_vista_prot:
                        excel.Quit()
                        print("[EXCEL] Instancia de Excel cerrada (sin libros abiertos)")
                except Exception:
                    pass
            finally:
                try:
                    excel.DisplayAlerts = True
                except Exception:
                    pass

        if cerrado:
            return True
        time.sleep(espera)  # tal vez aún no terminaba de abrir; reintenta

    return False


def _diagnostico_excel(ruta_archivo):
    """
    Vuelca en consola QUÉ tiene Excel abierto cuando no se pudo cerrar/borrar el
    archivo. Sirve para depurar en el equipo del usuario sin adivinar: dice
    cuántas instancias hay y, en cada una, qué libros y qué vistas protegidas
    tiene. Si algún día Gallo vuelve a no cerrarse, este log es la pista.
    """
    try:
        instancias = _instancias_excel_abiertas()
        print(
            f"[EXCEL][DIAG] No se pudo cerrar '{os.path.basename(ruta_archivo)}'. "
            f"Instancias de Excel detectadas: {len(instancias)}"
        )
        for i, excel in enumerate(instancias, 1):
            try:
                libros = [str(wb.Name) for wb in list(excel.Workbooks)]
            except Exception as e:
                libros = [f"<error al listar: {e}>"]
            vistas = []
            try:
                vistas = [str(p.Caption) for p in list(excel.ProtectedViewWindows)]
            except Exception:
                pass
            print(f"[EXCEL][DIAG]   Instancia {i}: libros={libros} vista_protegida={vistas}")
    except Exception as e:
        print(f"[EXCEL][DIAG] Falló el diagnóstico: {e}")


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
        time.sleep(3)  # deja que SAP termine de ESCRIBIR el archivo

        # === PRIMERO salir de la lista en SAP → suelta el .xlsx ===
        # Si SAP muestra el resultado con "Excel in place", es SAP quien tiene
        # el archivo agarrado (no un Excel que podamos cerrar por COM). Al
        # regresar a la pantalla inicial, SAP suelta ese Excel incrustado y el
        # archivo queda libre para convertir/borrar.
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        time.sleep(2)

        # === LUEGO cerrar cualquier Excel normal que haya quedado abierto ===
        _cerrar_excel_workbook(_ruta_xlsx)

        # === Convertir el .xlsx temporal a .csv (consistencia de formato) ===
        if _exportar_a_csv:
            _xlsx_a_csv(_ruta_xlsx, ruta_completa)

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
