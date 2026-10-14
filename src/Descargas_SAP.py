"""
Validación Factura Global — Descargas SAP
==========================================

SAP transaction extraction module for the 'Validación Factura Global'
project.

Implemented transactions:
  - 'Gallo'   (FAGLL03)
  - 'Monivoi' (ZLMXCOM_TRN_MONINVOI)
  - 'VBRL'    (ZSE16, table VBRL) — for the invoices listed in the Monivoi file
  - 'Venta EIAP' (FAGLL03 with the EIAP accounts; the end date is moved back
    one day, see utils.fecha_efectiva_eiap)

Based on the original VBS recordings, with these improvements:
  1. The dates are entered by the user (no longer hard-coded).
  2. The file name uses the SAME date the user typed, in D_mmm_YY format
     (Spanish month, no leading zero on the day; see utils.py).
  3. Supports downloading a single day or a whole range in ONE SAP query (one
     file for the entire range, not one file per day).
  4. The folder and the file name are typed straight into the export dialog
     (more robust than navigating with F4 like the VBS did).
"""

import os
import re
import time
import traceback
from datetime import datetime

import openpyxl as _oxl          # makes PyInstaller bundle the Excel engine (also used below)
import pandas as pd
import win32com.client
import win32clipboard
import pywintypes

import ctypes
from ctypes import wintypes

from utils import (
    FORMATO_FECHA,
    parse_fecha,
    nombre_archivo_dia,
    nombre_archivo_dia_monivoi,
    nombre_archivo_rango_eiap,
    fecha_efectiva_monivoi_str,
    fecha_efectiva_eiap,
)

# ----------------------------------------------------------------------
# Transaction constants (taken from the original VBS recording)
# ----------------------------------------------------------------------

# Accounts in the VBS multiple selection (SD_SAKNR).
_CUENTAS_GALLO = [
    "4000005", "4150000", "4200000", "4200004", "4200015",
    "4260000", "7000005",
]

# Accounts of the 'Solo Venta EIAP' button (same FAGLL03 selection, SD_SAKNR).
# Public (no underscore) because the controller shows them in the confirmation.
CUENTAS_EIAP = ("4000010", "4000020", "4000802", "4300013")

# Default company code (BUKRS) from the VBS.
_SOCIEDAD_DEFAULT = "MX21"

# Default display layout / variant (PA_VARI).
_VARIANTE_DEFAULT = "BASE/VALID"

# A range is now ONE query and ONE spreadsheet export, so both can take much
# longer than a single day. These are upper bounds (seconds), not fixed pauses:
# the waits return as soon as SAP is free / the file is complete.
_ESPERA_CONSULTA_SAP = 1800     # SAP still running the report
_ESPERA_EXPORTACION  = 900      # SAP still writing the exported .xlsx

# An .xlsx holds at most 1,048,576 rows, header included. SAP's spreadsheet
# export cannot go past that, so a result that reaches it may have been cut and
# must NOT be used (see _xlsx_a_csv).
_EXCEL_MAX_FILAS_DATOS = 1_048_576 - 1


class ExportacionTruncadaError(RuntimeError):
    """
    The SAP export reached Excel's row limit, so the file may be incomplete.

    Gallo_FAGLL03 / Monivoi_ZLMXCOM let this one propagate (instead of turning
    it into the usual `False`), so the controller shows its message in the
    summary dialog: it is not "no data", it is "do not trust this download".
    """


# ----------------------------------------------------------------------
# Monivoi transaction constants (ZLMXCOM_TRN_MONINVOI)
# ----------------------------------------------------------------------
_TX_MONIVOI    = "ZLMXCOM_TRN_MONINVOI"   # custom (Z) transaction
_VKORG_DEFAULT = "MX21"                    # sales organization
_VTWEG_DEFAULT = "01"                      # distribution channel
_RFC_DEFAULT   = "XAXX010101000"           # generic RFC ("público en general" = general public)
# Confirmation popup SAP raises when Back is pressed on the Monivoi list
# ("end the list?"). Option 1 is "Yes"; the recording presses it.
_SPOP_OPCION_1 = "wnd[1]/usr/btnSPOP-OPTION1"


# ----------------------------------------------------------------------
# VBRL transaction constants (ZSE16 → table VBRL), from the VBS recording
# ----------------------------------------------------------------------
_TX_ZSE16      = "ZSE16"                   # table browser (custom copy of SE16)
_TABLA_VBRL    = "VBRL"                    # table typed in GV_TABNAME
_VBRL_ARCHIVO_DEFAULT = "VBRL.csv"
# Multiple-selection button of the FIRST selection field of the table: the one
# the recording fills with the invoice numbers.
_VBRL_BTN_SELECCION_MULTIPLE = "wnd[0]/usr/btn%_I1_%_APP_%-VALU_PUSH"
# Layout popup: the layout is looked up BY NAME among the labels of the popup.
# (The recording double-clicked the label at row 3, which silently picks another
# layout as soon as one is added or removed in SAP.)
_VBRL_LAYOUT_NOMBRE = "VALS"
# List > Export > Spreadsheet, as recorded (it is NOT the same menu path as in
# Gallo/Monivoi).
_VBRL_MENU_EXPORTAR = "wnd[0]/mbar/menu[0]/menu[10]/menu[1]/menu[1]"
# Status-bar texts (lower case) that mean "the table has no entries for that
# selection". They are deliberately specific: a generic fragment could also match
# an unrelated message and turn a real error into a fake "no data". If your SAP
# words it differently, the run STOPS with the exact status-bar text printed (it
# never guesses): copy that text here.
_MENSAJES_SIN_ENTRADAS_SE16 = (
    "no table entries",                 # EN: "No table entries found for specified key"
    "ninguna entrada de tabla",         # ES variants
    "no existen entradas de tabla",
    "no se encontraron entradas de tabla",
    "no se han encontrado entradas de tabla",
)
# Checks on the downloaded file. The consolidation reads VBRL POSITIONALLY
# (column D = invoice, column E = amount, as in the stakeholder's XLOOKUP). The
# layout is picked by name, but someone could still edit the columns of that
# layout in SAP and silently change what gets exported. These checks catch that
# at download time.
# Keep them in sync with Consolidacion.EtapaVBRL.COL_FACTURA / COL_IMPORTE
# (1-based 4 and 5): change both places together.
_VBRL_MIN_COLUMNAS    = 5
_VBRL_COL_FACTURA_IDX = 3        # 0-based index of column D
_VBRL_MIN_FRACCION_NUMERICA = 0.9


# ----------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------
def ruta_input_default():
    """
    Return the project's input folder:
        C:\\Users\\<user>\\Documents\\Validacion Factura Global\\src\\Input

    It uses the system user, so it runs on any machine without touching the
    code.
    """
    user_profile = os.environ.get("USERPROFILE") or os.path.expanduser("~")
    return os.path.join(
        user_profile, "Documents", "Validacion Factura Global", "src", "Input"
    )


def _xlsx_a_csv(ruta_xlsx, ruta_csv, eliminar_xlsx=True):
    """
    Convert the .xlsx that SAP generates on export into CSV (utf-8-sig), so
    that ALL downloaded files end up in the same format. Then delete the
    intermediate .xlsx. utf-8-sig makes Excel display accents correctly.
    """
    df = pd.read_excel(ruta_xlsx, dtype=str, engine="openpyxl")

    # Truncation guard. A whole range is exported as ONE spreadsheet, and an
    # .xlsx cannot hold more than 1,048,576 rows: a result that reaches that
    # size may have been cut off, and a silently incomplete file would
    # understate every sum downstream. Stop instead of writing the CSV.
    if len(df) >= _EXCEL_MAX_FILAS_DATOS:
        filas = len(df)
        del df
        if eliminar_xlsx:
            _borrar_xlsx_con_reintentos(ruta_xlsx)
        raise ExportacionTruncadaError(
            f"El export de SAP llegó al límite de filas de Excel ({filas:,}): el archivo "
            "pudo quedar INCOMPLETO y no se usó. Descarga un periodo más corto (menos "
            "días) y vuelve a intentar."
        )

    df.to_csv(ruta_csv, index=False, encoding="utf-8-sig")
    if eliminar_xlsx:
        _borrar_xlsx_con_reintentos(ruta_xlsx)

def _terminar_proceso(pid):
    """Force-close process 'pid' (Windows). True if it succeeded."""
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
    Last resort when the .xlsx stays LOCKED and COM could not close it.

    Why? In Gallo, Excel opens the exported file and sometimes gets stuck on a
    modal DIALOG waiting for a click (the taskbar icon flashes AMBER: Windows
    saying "this window wants your attention"). An Excel in a modal state does
    NOT answer COM, but it DOES keep holding the file → os.remove fails.

    This helper uses the Windows Restart Manager (the same API behind "this
    file is in use by: Excel"). It asks WHICH processes are holding THIS
    specific file and closes ONLY those whose application name contains
    "Excel". It never touches SAP:
      - if SAP is the one holding it ('in place' case, like Monivoi) → nothing;
      - if it is a standalone Excel stuck on a dialog (Gallo) → it is closed.

    Safe: the data is ALREADY on disk (that is why the .csv comes out fine);
    that Excel only DISPLAYS the file, it has no changes worth saving.
    Caveat: TerminateProcess ends the WHOLE Excel process, so if that same
    instance also had other workbooks open, their unsaved changes are lost.
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

            # 1st call: how many processes? (returns ERROR_MORE_DATA)
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
                # ONLY Excel; never SAP or any other process.
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
    Delete the intermediate .xlsx. On Windows, if Excel still has it open, the
    file stays LOCKED and os.remove fails. That is why every attempt first
    tries to close the file in Excel again and then deletes it, giving Excel
    time to let go of it.

    Wide window (intentos=10) on purpose: Gallo is the big report and the
    first one of the run, so Excel starts "cold" and can take several seconds
    to finish opening the file; until it is fully open, we cannot close it.
    It exits as soon as the delete succeeds, so the normal case is fast.
    """
    for intento in range(1, intentos + 1):
        _cerrar_excel_workbook(ruta_xlsx)        # try closing it in Excel again
        try:
            if os.path.exists(ruta_xlsx):
                os.remove(ruta_xlsx)
            print(f"[XLSX] Temporal borrado: {ruta_xlsx}")
            return True
        except OSError as e:
            print(f"[XLSX] Aún bloqueado (intento {intento}/{intentos}): {e}")

            # If COM does not let go (typical in Gallo: Excel stuck on an AMBER
            # dialog that COM cannot close), escalate to the Restart Manager to
            # close ONLY that Excel and free the file. It happens halfway
            # through the attempts, not on the first one, in case the file
            # frees itself.
            if intento == max(1, intentos // 2):
                _liberar_archivo_con_restart_manager(ruta_xlsx)

            time.sleep(espera)

    # Getting here means it could not be deleted: dump what Excel has open.
    print(f"[XLSX] ⚠️ No se pudo borrar (sigue abierto/bloqueado): {ruta_xlsx}")
    _diagnostico_excel(ruta_xlsx)
    return False


def _instancias_excel_abiertas():
    """
    Return ALL the Excel instances open on the machine, not just the first
    one. This is KEY for the "the .xlsx stays open and cannot be deleted" bug:
    when SAP exports, it opens the file in Excel, and it often does so in an
    instance DIFFERENT from the one the user already had open.
    win32com.GetObject(Class="Excel.Application") only reaches ONE instance,
    so if the workbook is in another one, we would never close it.

    To see them all we scan the Windows Running Object Table (ROT): every open
    workbook is registered there by its path; from each workbook we take its
    .Application and collect them without repeats (using Hwnd as identity).

    Degrades gracefully: if the ROT scan fails for whatever reason, it at
    least returns the main instance via GetObject.
    """
    instancias = {}

    # 1) Via the ROT: find ALL the instances from the open workbooks.
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
            # Open workbooks are registered by their file path.
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

    # 2) Fallback / complement: the registered main instance.
    try:
        app = win32com.client.GetObject(Class="Excel.Application")
        instancias[int(app.Hwnd)] = app
    except Exception:
        pass

    return list(instancias.values())


def _cerrar_excel_workbook(ruta_archivo, intentos=3, espera=1):
    """
    Close (without saving) the file that SAP/Excel leaves open after the
    export, so it does not stay locked (while it is open, Windows does not let
    us delete the temporary .xlsx).

    Sweeps ALL Excel instances (see _instancias_excel_abiertas), not just one,
    because SAP usually opens the file in a separate instance. In each
    instance it looks for the file in TWO places:
      A) Regular workbooks (Workbooks collection), matching by FULL PATH and,
         if that fails, by file NAME (SAP sometimes reports the path in a
         different shape: upper case, 8.3 short path, etc.).
      B) PROTECTED VIEW windows (ProtectedViewWindows). This is key for
         Gallo! Excel opens freshly exported/large files in "Protected View",
         and those windows do NOT show up in Workbooks, so before this we did
         not even touch them and the file stayed locked.

    If, after closing, an instance is left with NO workbooks or protected
    views, it is closed so no empty window is left behind. It NEVER closes an
    instance that still holds other workbooks of the user.

    Retries a few times in case Excel had not finished opening the file yet
    (Gallo, being the big report and the first one, starts Excel "cold" and
    takes longer). Returns True if it closed something that matched; False if
    it found nothing.
    """
    ruta_abs      = os.path.normcase(os.path.abspath(ruta_archivo))
    objetivo_base = os.path.normcase(os.path.basename(ruta_archivo))

    for _ in range(intentos):
        instancias = _instancias_excel_abiertas()
        if not instancias:
            return False  # no Excel open → nothing to close

        cerrado = False
        for excel in instancias:
            try:
                excel.DisplayAlerts = False
            except Exception:
                pass

            try:
                # === A) Regular workbooks ===
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

                # === B) Protected View windows (not in Workbooks) ===
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
                    pass  # the instance may not expose ProtectedViewWindows

                # === If the instance was left completely empty, close it ===
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
        time.sleep(espera)  # maybe it had not finished opening yet; retry

    return False


def _diagnostico_excel(ruta_archivo):
    """
    Dump to the console WHAT Excel has open when the file could not be
    closed/deleted. Useful to debug on the user's machine without guessing:
    it tells how many instances there are and, for each one, which workbooks
    and protected views it has. If Gallo ever stops closing again, this log is
    the clue.
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


def crear_archivo_sin_datos(ruta_archivo):
    """
    Create the 'Sin datos' placeholder that stands for a download with no rows
    (a CSV with a single 'Sin datos' column, or an empty workbook for .xlsx), so
    the rest of the app can tell "downloaded, but empty" from "not downloaded".
    """
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


def _verificar_sin_partidas(session, ruta_archivo, mensajes_extra=()):
    """
    Detect SAP's 'No se ha seleccionado ninguna partida' (no items selected)
    message. If it shows up: close the popup, create an empty Excel/CSV as a
    placeholder and return True so the caller knows there is no data.

    NOTE: the message codes are tuned for FAGLL03 (Gallo). For Monivoi (Z)
    the 'no data' text may be different; if your tests show it is not
    detected, add the exact message to MENSAJES_SIN_DATOS. Transactions with
    their own wording (ZSE16/VBRL) pass it in `mensajes_extra` (lower case).
    """
    try:
        sin_partidas = False
        MENSAJES_SIN_DATOS = (
            ("MSITEM033", "MSITEM030", "ninguna partida", "ninguna cuenta")
            + tuple(mensajes_extra)
        )

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
            crear_archivo_sin_datos(ruta_archivo)
            return True

        return False

    except Exception as e:
        print(f"[ERROR] _verificar_sin_partidas: {e}")
        return False

def _copiar_al_portapapeles(texto, intentos=5, espera=0.3):
    """
    Put 'texto' on the Windows clipboard (Unicode). Retries a few times
    because another app sometimes holds the clipboard for an instant.
    NOTE: this OVERWRITES whatever the user had copied at that moment.
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


def _leer_portapapeles(intentos=5, espera=0.3):
    """
    Return the Unicode text currently on the Windows clipboard ('' if there is
    none). Retries a few times for the same reason as _copiar_al_portapapeles.
    """
    ultimo_error = None
    for _ in range(intentos):
        try:
            win32clipboard.OpenClipboard()
            try:
                if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                    return win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                return ""
            finally:
                win32clipboard.CloseClipboard()
        except Exception as e:
            ultimo_error = e
            time.sleep(espera)
    raise RuntimeError(f"No se pudo leer el portapapeles: {ultimo_error}")


def _copiar_y_verificar_portapapeles(texto, intentos=3):
    """
    Put 'texto' on the clipboard and read it back. If something else changed it
    in between (another app, the user copying), write it again. Raises
    RuntimeError if it cannot be made to stick: pasting the wrong text into
    SAP's selection would query the wrong invoices.
    """
    for _ in range(intentos):
        _copiar_al_portapapeles(texto)
        if _leer_portapapeles().replace("\x00", "") == texto:
            return True
        time.sleep(0.3)
    raise RuntimeError(
        "El portapapeles cambió mientras se preparaba la lista de facturas. "
        "No copies nada mientras corre la descarga e inténtalo de nuevo."
    )


def _importar_desde_portapapeles_popup(session):
    """
    Press the 'Importar desde portapapeles' (import from clipboard) button of
    the multiple-selection popup (used for the Gallo accounts and the VBRL
    invoices). Tries the known IDs in case it changes between SAP systems.
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
# Gallo transaction (FAGLL03)
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
    Download transaction FAGLL03 ('Gallo') from SAP into a .csv file.

    Args:
        DateFrom (str): Start date of the range, 'DD.MM.YYYY' format.
        DateTo   (str): End date of the range,   'DD.MM.YYYY' format.
        FolderPath (str|None): Target folder. If None → ruta_input_default().
        FileName   (str|None): File name. If None → nombre_archivo_dia(DateTo).
        sociedad   (str): Company code (BUKRS). Default 'MX21'.
        cuentas    (list|None): Accounts to filter. If None → the VBS ones.
        variante   (str|None): Name of the layout/variant (PA_VARI).
            Default 'BASE/VALID'. If None or empty, the default is used.
            It is always typed by name (ROBUST).

    Returns:
        bool: True if it downloaded data, False if there was no data or an error.
    """
    session = None
    cuentas       = cuentas or _CUENTAS_GALLO
    FolderPath    = FolderPath or ruta_input_default()
    FileName      = FileName or nombre_archivo_dia(DateTo)
    ruta_completa = os.path.join(FolderPath, FileName)   # final target (may be .csv)

    # SAP only exports in spreadsheet format (.xlsx). If the final file we
    # want is .csv, we first export to a temporary .xlsx and convert it at the
    # end. This way ALL the files end up in the same format.
    _base, _ext     = os.path.splitext(FileName)
    _exportar_a_csv = _ext.lower() == ".csv"
    _nombre_xlsx    = (_base + ".xlsx") if _exportar_a_csv else FileName
    _ruta_xlsx      = os.path.join(FolderPath, _nombre_xlsx)

    try:
        # === Minimal validation ===
        for etiqueta, fecha in (("DateFrom", DateFrom), ("DateTo", DateTo)):
            datetime.strptime(fecha, "%d.%m.%Y")  # raises ValueError if the format is invalid

        os.makedirs(FolderPath, exist_ok=True)

        # === SAP connection ===
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

        # === Open transaction FAGLL03 ===
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/nFAGLL03"
        session.findById("wnd[0]").sendVKey(0)

        # === Load the accounts (multiple selection SD_SAKNR) via the CLIPBOARD ===
        # Instead of typing row by row (and having to deal with scrolling when
        # there are many accounts), we copy ALL the accounts to the clipboard —one
        # per line— and use the popup's 'Importar desde portapapeles' button. This
        # way it does not matter how many accounts there are: they go in at once.
        session.findById("wnd[0]/usr/btn%_SD_SAKNR_%_APP_%-VALU_PUSH").press()

        _copiar_al_portapapeles("\r\n".join(str(c).strip() for c in cuentas))
        _importar_desde_portapapeles_popup(session)
        time.sleep(1)                                       # let SAP paint the accounts

        session.findById("wnd[1]/tbar[0]/btn[0]").press()  # copy/adopt
        session.findById("wnd[1]/tbar[0]/btn[8]").press()  # execute and close the popup

        # === Selection parameters ===
        session.findById("wnd[0]/usr/radX_AISEL").select()          # all items
        session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").text  = sociedad
        session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text  = DateFrom
        session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = DateTo

        # Layout/variant ALWAYS by name (robust). If it is empty or None, use
        # the default instead of the fragile F4 + scroll + coordinate path.
        variante = variante or _VARIANTE_DEFAULT
        campo_vari = session.findById("wnd[0]/usr/ctxtPA_VARI")
        campo_vari.text          = variante
        campo_vari.setFocus()
        campo_vari.caretPosition = len(variante)

        # === Run the report ===
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        _esperar_sap_libre(session, timeout=_ESPERA_CONSULTA_SAP)   # a wide range runs longer
        time.sleep(4)

        # === No items? ===
        if _verificar_sin_partidas(session, ruta_completa):
            return False  # empty file already created

        # === Export to Excel (robust way: folder + file name typed directly) ===
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text     = FolderPath
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = _nombre_xlsx
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(_nombre_xlsx)
        t_exportacion = time.time()
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        _esperar_exportacion(_ruta_xlsx, t_exportacion)   # let SAP finish WRITING the file

        # === FIRST leave the list in SAP → it releases the .xlsx ===
        # If SAP shows the result with "Excel in place", SAP itself is the one
        # holding the file (not an Excel we could close through COM). Going
        # back to the initial screen makes SAP release that embedded Excel and
        # the file is free to convert/delete.
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        time.sleep(2)

        # === THEN close any regular Excel that was left open ===
        _cerrar_excel_workbook(_ruta_xlsx)

        # === Convert the temporary .xlsx to .csv (format consistency) ===
        if _exportar_a_csv:
            _xlsx_a_csv(_ruta_xlsx, ruta_completa)

        print(f"[GALLO] Archivo descargado correctamente: {ruta_completa}")
        return True

    except pywintypes.com_error as e:
        print(f"[ERROR COM SAP] Gallo_FAGLL03: {e}")
        return False
    except ExportacionTruncadaError:
        raise    # shown to the user by the caller (see the class docstring)
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
# SAP navigation helpers (shared by the transactions below)
# ----------------------------------------------------------------------
def _existe_en_sap(session, id_componente):
    """True if the SAP GUI component exists (findById raises when it does not)."""
    try:
        session.findById(id_componente)
        return True
    except Exception:
        return False


def _esperar_componente(session, id_componente, timeout=5.0, intervalo=0.25):
    """
    Poll until the SAP GUI component exists. Returns True as soon as it does and
    False if it still does not after `timeout` seconds. Never raises.
    """
    limite = time.time() + timeout
    while True:
        if _existe_en_sap(session, id_componente):
            return True
        if time.time() >= limite:
            return False
        time.sleep(intervalo)


def _esperar_sap_libre(session, timeout=300):
    """Wait while SAP GUI is busy (long queries). Never raises; gives up after `timeout` s."""
    limite = time.time() + timeout
    try:
        while session.Busy and time.time() < limite:
            time.sleep(0.5)
    except Exception:
        pass


def _cerrar_popups_abiertos(session, maximo=3):
    """
    Best effort: cancel (F12) any popup that a failed run left open, so the
    next transaction does not start on top of it.
    """
    for _ in range(maximo):
        if not _existe_en_sap(session, "wnd[1]"):
            return
        try:
            session.findById("wnd[1]").sendVKey(12)
        except Exception:
            return


def _volver_al_menu_desde_lista(session, etiqueta="MONIVOI"):
    """
    Leave an exported list and land on the SAP main menu. It is the recorded
    sequence:

        Back  →  popup "end the list?"  →  option 1 (Yes)  →  Back

    The popup is awaited (a few seconds) instead of assumed: pressing a second
    Back while it is open does nothing, and the session would stay stuck on it.
    If SAP does not ask for the confirmation, that step is skipped. If a DIFFERENT
    popup shows up it is left alone (it could be asking for something else) and
    the function raises, so the caller can report it.

    Raises RuntimeError if a popup is still open at the end.
    """
    session.findById("wnd[0]/tbar[0]/btn[3]").press()

    if _esperar_componente(session, _SPOP_OPCION_1):
        session.findById(_SPOP_OPCION_1).press()
        _esperar_sap_libre(session)
    elif _existe_en_sap(session, "wnd[1]"):
        raise RuntimeError(
            "SAP mostró un popup distinto al de confirmar el cierre de la lista "
            "y no lo toqué."
        )
    else:
        print(f"[{etiqueta}] SAP no pidió confirmar el cierre de la lista; sigo al menú.")

    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    if _existe_en_sap(session, "wnd[1]"):
        raise RuntimeError("Quedó un popup abierto al regresar al menú principal.")


# ----------------------------------------------------------------------
# Venta EIAP ('Solo Venta EIAP' button): FAGLL03 with the EIAP accounts
# ----------------------------------------------------------------------
def VentaEIAP_FAGLL03(
    DateFrom,
    DateTo,
    FolderPath=None,
    FileName=None,
    sociedad=_SOCIEDAD_DEFAULT,
):
    """
    Download the Gallo transaction (FAGLL03) for the EIAP accounts
    (CUENTAS_EIAP) into a .csv file, as ONE query for the whole range.

    The ONLY difference vs the regular Gallo download, besides the accounts, is
    the end date: SAP is queried up to DateTo MINUS ONE DAY (see
    utils.fecha_efectiva_eiap, the single place where that -1 lives). Typing
    01.01.2026 → 10.01.2026 queries 01.01.2026 → 09.01.2026. The file name keeps
    the dates as typed (see utils.nombre_archivo_rango_eiap).

    Args:
        DateFrom, DateTo (str): 'DD.MM.YYYY' dates as the user typed them.
        FolderPath (str|None): target folder. If None → ruta_input_default().
        FileName   (str|None): file name. If None → nombre_archivo_rango_eiap(...).
        sociedad   (str): Company code (BUKRS). Default 'MX21'.

    Returns:
        bool: True if it downloaded data, False if there was no data or an error
        (same contract as Gallo_FAGLL03).

    Raises:
        ValueError: a date is not 'DD.MM.YYYY', or DateTo minus one day falls
            before DateFrom (e.g. DateFrom == DateTo): the range would be empty
            or backwards, and SAP must never be queried with it.
    """
    desde = parse_fecha(DateFrom)
    hasta = fecha_efectiva_eiap(DateTo)
    if hasta < desde:
        raise ValueError(
            f"Venta EIAP: con Fecha Hasta {DateTo} la consulta terminaría el "
            f"{hasta.strftime(FORMATO_FECHA)}, antes de la Fecha Desde ({DateFrom}). "
            "El periodo debe abarcar al menos 2 días."
        )
    return Gallo_FAGLL03(
        DateFrom=desde.strftime(FORMATO_FECHA),
        DateTo=hasta.strftime(FORMATO_FECHA),
        FolderPath=FolderPath,
        FileName=FileName or nombre_archivo_rango_eiap(DateFrom, DateTo),
        sociedad=sociedad,
        cuentas=list(CUENTAS_EIAP),
    )


# ----------------------------------------------------------------------
# Monivoi transaction (ZLMXCOM_TRN_MONINVOI)
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
    Download transaction ZLMXCOM_TRN_MONINVOI ('Monivoi') from SAP into .csv.

    Differences vs Gallo (FAGLL03):
      - Company code through P_BUKRS (not SD_BUKRS-LOW).
      - Filters by CREATION date (SO_ERDAT), not by posting date.
      - Carries sales org. (VKORG), channel (VTWEG) and RFC (STCD1).
      - No accounts and no layout/variant.
      - On export, SAP first shows a format selector (cmbG_LISTBOX).

    Args:
        DateFrom, DateTo (str): 'DD.MM.YYYY' dates as the user typed them;
            SO_ERDAT gets them +1 day (see utils.fecha_efectiva_monivoi).
        FolderPath (str|None): target folder. If None → ruta_input_default().
        FileName   (str|None): file name. If None → nombre_archivo_dia_monivoi(DateTo).
        sociedad (str): P_BUKRS. Default 'MX21'.
        vkorg (str): sales organization (SO_VKORG). Default 'MX21'.
        vtweg (str): distribution channel (SO_VTWEG). Default '01'.
        rfc   (str): RFC to filter (SO_STCD1). Default 'XAXX010101000'.

    Returns:
        bool: True if it downloaded data, False if there was no data or an error.
    """
    session = None
    en_menu = False      # True once the run has left SAP on the main menu
    FolderPath    = FolderPath or ruta_input_default()
    FileName      = FileName or nombre_archivo_dia_monivoi(DateTo)
    ruta_completa = os.path.join(FolderPath, FileName)

    # Same as Gallo: SAP exports to .xlsx; if the target is .csv, we export to
    # a temporary .xlsx and convert it at the end.
    _base, _ext     = os.path.splitext(FileName)
    _exportar_a_csv = _ext.lower() == ".csv"
    _nombre_xlsx    = (_base + ".xlsx") if _exportar_a_csv else FileName
    _ruta_xlsx      = os.path.join(FolderPath, _nombre_xlsx)

    try:
        # === Validate the dates ===
        for _etq, fecha in (("DateFrom", DateFrom), ("DateTo", DateTo)):
            datetime.strptime(fecha, "%d.%m.%Y")

        os.makedirs(FolderPath, exist_ok=True)

        # === SAP connection ===
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

        # === Open the transaction (with /n so it starts from any screen) ===
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n" + _TX_MONIVOI
        session.findById("wnd[0]").sendVKey(0)

        # === Selection parameters ===
        session.findById("wnd[0]/usr/ctxtP_BUKRS").text      = sociedad
        session.findById("wnd[0]/usr/ctxtSO_ERDAT-LOW").text  = fecha_efectiva_monivoi_str(DateFrom)
        session.findById("wnd[0]/usr/ctxtSO_ERDAT-HIGH").text = fecha_efectiva_monivoi_str(DateTo)
        session.findById("wnd[0]/usr/ctxtSO_VKORG-LOW").text  = vkorg
        session.findById("wnd[0]/usr/ctxtSO_VTWEG-LOW").text  = vtweg

        campo_rfc = session.findById("wnd[0]/usr/txtSO_STCD1-LOW")
        campo_rfc.text          = rfc
        campo_rfc.setFocus()
        campo_rfc.caretPosition = len(rfc)

        # === Run the report ===
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        _esperar_sap_libre(session, timeout=_ESPERA_CONSULTA_SAP)   # a wide range runs longer
        time.sleep(4)

        # === No data? (reuses the Gallo helper; see the note in that function) ===
        if _verificar_sin_partidas(session, ruta_completa):
            return False

        # === Export to spreadsheet ===
        # NOTE: here Monivoi differs from Gallo. SAP first opens a dialog with a
        # format selector (cmbG_LISTBOX); we take the default value and press OK.
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        try:
            session.findById("wnd[1]/usr/cmbG_LISTBOX").setFocus()
        except Exception:
            pass  # in case the format selector does not show up on some machine
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # === Save dialog ===
        # The VBS did not set the folder; we DO, to leave the file in Input.
        try:
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = FolderPath
        except Exception:
            pass
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = _nombre_xlsx
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(_nombre_xlsx)
        t_exportacion = time.time()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # === Wait for the file, close the Excel that SAP opens + convert to CSV ===
        _esperar_exportacion(_ruta_xlsx, t_exportacion)
        _cerrar_excel_workbook(_ruta_xlsx)
        if _exportar_a_csv:
            _xlsx_a_csv(_ruta_xlsx, ruta_completa)

        print(f"[MONIVOI] Archivo descargado correctamente: {ruta_completa}")

        # === Back to the main menu: Back → "end the list?" → Yes → Back ===
        # The file is already safe on disk, so a navigation problem here is
        # reported but does not turn a good download into a failure (the next
        # transaction also starts with /n and clears leftover popups).
        try:
            _volver_al_menu_desde_lista(session, "MONIVOI")
            en_menu = True
        except Exception as e:
            print(
                "[MONIVOI] ⚠️ El archivo ya quedó guardado, pero no pude regresar al "
                f"menú principal de SAP: {e}"
            )
        return True

    except pywintypes.com_error as e:
        print(f"[ERROR COM SAP] Monivoi_ZLMXCOM: {e}")
        return False
    except ExportacionTruncadaError:
        raise    # shown to the user by the caller (see the class docstring)
    except (ValueError, TypeError, ConnectionError, RuntimeError) as e:
        print(f"[ERROR CONTROLADO] Monivoi_ZLMXCOM: {e}")
        return False
    except Exception as e:
        print(f"[ERROR NO CONTROLADO] Monivoi_ZLMXCOM: {e}")
        print(traceback.format_exc())
        return False
    finally:
        # Only when the run did NOT reach the main menu (an extra Back from the
        # menu itself could raise SAP's "log off" prompt).
        if session is not None and not en_menu:
            try:
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
            except Exception:
                pass


# ----------------------------------------------------------------------
# VBRL transaction (ZSE16, table VBRL)
# ----------------------------------------------------------------------
def _normalizar_facturas_vbrl(facturas):
    """
    Clean the invoice list that goes to the clipboard: text, trimmed, without
    blanks or duplicates, first-appearance order kept. A '.0' left behind by a
    float conversion ('233982822.0') is removed.

    Raises ValueError if nothing is left. An EMPTY list is refused on purpose:
    with nothing in the selection field (and the max-hits limit lifted) ZSE16
    would try to read the WHOLE table.
    """
    if facturas is None or isinstance(facturas, (str, bytes)):
        raise ValueError(
            "La lista de facturas debe ser una colección (lista), no un texto suelto."
        )
    vistas = set()
    limpias = []
    for valor in facturas:
        if valor is None:
            continue
        texto = str(valor).strip()
        con_cero_decimal = re.fullmatch(r"([0-9]+)\.0+", texto)
        if con_cero_decimal:
            texto = con_cero_decimal.group(1)
        if not texto:
            continue
        # One invoice per clipboard line: a line break or tab INSIDE a value
        # would split it into several selections.
        if any(c in texto for c in "\r\n\t\x00"):
            raise ValueError(f"Número de factura con caracteres no válidos: {texto!r}")
        if texto not in vistas:
            vistas.add(texto)
            limpias.append(texto)
    if not limpias:
        raise ValueError(
            "No hay facturas que consultar en VBRL (la lista está vacía): "
            "sin facturas SAP intentaría leer la tabla completa."
        )
    return limpias


def _quitar_limite_resultados(session):
    """
    Clear 'Maximum No. of Hits' (MAX_SEL) on the ZSE16 selection screen, as the
    recording does. It is NOT optional: left at its default SAP returns only the
    first rows, and the invoices that fall outside would show up as 'Sin timbre'
    with no error anywhere. That is why the field is read back and the run stops
    if it is not empty.
    """
    campo = session.findById("wnd[0]/usr/txtMAX_SEL")
    campo.text = ""
    campo.setFocus()
    valor = str(campo.text).strip()
    if valor.strip("0"):                      # '' and '0' / '000' mean "no limit"
        raise RuntimeError(
            f"No pude quitar el límite de resultados de ZSE16 (MAX_SEL = '{valor}'): "
            "SAP devolvería sólo una parte de las facturas."
        )


def _etiquetas_del_popup(session):
    """
    Labels of the open popup's user area, as (component, name, text) tuples:
    `name` is the label's own id ('lbl[1,3]' = column 1, row 3) and `text` is
    stripped. Children that cannot be read are skipped.
    """
    area = session.findById("wnd[1]/usr")
    etiquetas = []
    for i in range(int(area.Children.Count)):
        hijo = area.Children(i)
        try:
            if str(hijo.Type) != "GuiLabel":
                continue
            etiquetas.append((hijo, str(hijo.Name), str(hijo.Text).strip()))
        except Exception:
            continue
    return etiquetas


def _elegir_layout_por_nombre(session, nombre=_VBRL_LAYOUT_NOMBRE):
    """
    Pick the display layout called `nombre` ('VALS') in the layouts popup and
    confirm it with F2 (double click), like the recording does.

    The rows of the popup are the labels of its FIRST column ('lbl[1,<row>]').
    A row matches when its text is the name, or starts with the name followed by
    a space (in case SAP puts the description in the same label). The comparison
    ignores case and surrounding blanks, and it is a WHOLE-word match: 'VALS2'
    is not 'VALS'. Description columns are never considered, so a description
    that happens to say 'VALS' cannot be picked by mistake.

    Raises RuntimeError, without touching anything, if no row matches or if more
    than one does. Nothing is guessed: a wrong layout would export other columns.
    """
    objetivo = nombre.strip().upper()
    filas = [
        (hijo, id_corto, texto)
        for hijo, id_corto, texto in _etiquetas_del_popup(session)
        if re.fullmatch(r"lbl\[1,\d+\]", id_corto) and texto
    ]
    coincidencias = [
        fila for fila in filas
        if fila[2].upper() == objetivo or fila[2].upper().split(None, 1)[0] == objetivo
    ]
    visibles = ", ".join(repr(texto) for _, _, texto in filas) or "(ninguno)"

    if not coincidencias:
        raise RuntimeError(
            f"No encontré el layout '{nombre}' en la lista de layouts de ZSE16. "
            f"Layouts que veo: {visibles}. Si lo renombraron, actualiza "
            "_VBRL_LAYOUT_NOMBRE; si la lista es larga, puede estar fuera de la vista."
        )
    if len(coincidencias) > 1:
        raise RuntimeError(
            f"Hay {len(coincidencias)} layouts que coinciden con '{nombre}' "
            f"({', '.join(repr(t) for _, _, t in coincidencias)}); no sé cuál elegir."
        )

    hijo, id_corto, texto = coincidencias[0]
    print(f"[VBRL] Layout elegido por nombre: {texto!r} ({id_corto})")
    hijo.setFocus()
    hijo.caretPosition = min(2, len(texto))
    session.findById("wnd[1]").sendVKey(2)

    # The popup must be gone before the export menu is used.
    limite = time.time() + 3
    while _existe_en_sap(session, "wnd[1]") and time.time() < limite:
        time.sleep(0.25)
    if _existe_en_sap(session, "wnd[1]"):
        raise RuntimeError(f"El popup de layouts sigue abierto después de elegir '{texto}'.")


def _esperar_archivo_nuevo(ruta, desde, timeout=90, estable=1.0):
    """
    Wait until SAP has finished WRITING `ruta`: the file exists, it was modified
    after `desde` (epoch seconds; a stale file from an earlier run must not
    count) and its size has not changed for `estable` seconds.
    Returns False if that does not happen within `timeout` seconds.
    """
    limite = time.time() + timeout
    ultimo_tam = None
    estable_desde = None
    while time.time() < limite:
        try:
            info = os.stat(ruta)
        except OSError:
            info = None
        # 2 s of slack: some file systems keep modification times in 2 s steps.
        if info is not None and info.st_size > 0 and info.st_mtime >= desde - 2:
            if info.st_size == ultimo_tam:
                if time.time() - estable_desde >= estable:
                    return True
            else:
                ultimo_tam, estable_desde = info.st_size, time.time()
        time.sleep(0.5)
    return False


def _esperar_exportacion(ruta_xlsx, desde):
    """
    Wait until SAP has finished WRITING the exported .xlsx of Gallo / Monivoi.

    It replaces the old fixed 3 s pause: a whole range is one big file (tens or
    hundreds of MB), and converting a half-written .xlsx would give a corrupt or
    incomplete CSV. `desde` is the epoch time taken just before pressing the save
    button. The size must stay unchanged for 3 s (a big workbook is written in
    bursts). Raises RuntimeError if the file is not complete within
    _ESPERA_EXPORTACION seconds.
    """
    if not _esperar_archivo_nuevo(ruta_xlsx, desde,
                                  timeout=_ESPERA_EXPORTACION, estable=3.0):
        raise RuntimeError(f"SAP no terminó de escribir el archivo exportado: {ruta_xlsx}")


def _validar_csv_vbrl(ruta_csv, facturas):
    """
    Sanity-check the converted VBRL file BEFORE it replaces the final one.

    The consolidation reads VBRL by POSITION (column D = invoice, column E =
    amount, like the stakeholder's XLOOKUP), and the layout is chosen by screen
    position. If the layout list ever changes, SAP would export other columns
    and the cross would silently fill with 'Sin timbre'. So:

      - fewer than _VBRL_MIN_COLUMNAS columns, no rows, or a column D that is
        not (almost) all digits → RuntimeError (the file is not used);
      - how many of the requested invoices show up in column D is only
        REPORTED: if D carries another folio than the one pasted into SAP, zero
        matches is normal.

    Returns (filas, pedidas, encontradas).
    """
    df = pd.read_csv(ruta_csv, dtype=str, encoding="utf-8-sig", keep_default_na=False)
    if df.shape[1] < _VBRL_MIN_COLUMNAS:
        raise RuntimeError(
            f"El archivo exportado de VBRL trae {df.shape[1]} columna(s) y se esperan al "
            f"menos {_VBRL_MIN_COLUMNAS} (D = factura, E = importe). Revisa el layout elegido."
        )
    if df.empty:
        raise RuntimeError("El archivo exportado de VBRL no trae filas.")

    col = df.iloc[:, _VBRL_COL_FACTURA_IDX].astype(str).str.strip()
    col = col[col != ""]
    if col.empty:
        raise RuntimeError("La columna D del archivo exportado de VBRL está vacía. Revisa el layout elegido.")
    es_numerica = col.str.fullmatch(r"[0-9]+(\.0+)?")
    fraccion = float(es_numerica.mean())
    if fraccion < _VBRL_MIN_FRACCION_NUMERICA:
        raise RuntimeError(
            f"La columna D del archivo exportado de VBRL no parece de facturas "
            f"(sólo {fraccion:.0%} de sus valores son numéricos; primero: {col.iloc[0]!r}). "
            "Revisa el layout elegido."
        )

    pedidas = {int(f) for f in facturas if re.fullmatch(r"[0-9]+", f)}
    halladas = {int(v.split(".", 1)[0]) for v in col[es_numerica]}
    encontradas = len(pedidas & halladas)
    print(
        f"[VBRL] {len(df):,} fila(s); {encontradas:,} de {len(pedidas):,} factura(s) "
        f"pedida(s) aparecen en la columna D."
    )
    if pedidas and not encontradas:
        print(
            "[VBRL] ⚠️ Ninguna factura pedida aparece en la columna D. Si esa columna lleva "
            "otro folio es normal; si no, revisa el layout elegido."
        )
    return len(df), len(pedidas), encontradas


def VBRL_ZSE16(facturas, FolderPath=None, FileName=None):
    """
    Download table VBRL through transaction ZSE16 for a list of invoices and
    save it as a .csv file.

    The invoices are the distinct values of the Monivoi 'Factura' column (the
    caller reads them; see Consolidacion.leer_facturas_monivoi). They travel
    through the CLIPBOARD into the multiple-selection popup of the table's first
    selection field, exactly like the recording does.

    Differences vs Gallo / Monivoi:
      - It is NOT downloaded by date: it is one query for the whole period.
      - 'Maximum No. of Hits' is lifted (and checked), otherwise SAP would cut
        the result.
      - The file is converted to a temporary '.part' and only replaces the final
        one after it passes _validar_csv_vbrl, so a failed run never leaves a
        half-baked or stale-looking VBRL file behind.

    Args:
        facturas (list[str]): invoice numbers to look up. Must not be empty.
        FolderPath (str|None): target folder. If None → ruta_input_default().
        FileName   (str|None): file name (.csv). If None → 'VBRL.csv'.

    Returns:
        bool: True if it downloaded data, False if there was no data or an error.
    """
    session = None
    FolderPath    = FolderPath or ruta_input_default()
    FileName      = FileName or _VBRL_ARCHIVO_DEFAULT
    ruta_completa = os.path.join(FolderPath, FileName)

    # SAP exports .xlsx; like the other downloads, VBRL ends up as .csv, so the
    # export goes to a temporary .xlsx that is converted at the end.
    _base, _ext   = os.path.splitext(FileName)
    _nombre_xlsx  = _base + ".xlsx"
    _ruta_xlsx    = os.path.join(FolderPath, _nombre_xlsx)
    _ruta_parcial = ruta_completa + ".part"

    try:
        # === Validation (before touching SAP) ===
        if os.path.basename(FileName) != FileName:
            raise ValueError(f"El nombre del archivo no debe incluir carpetas: {FileName!r}")
        if _ext.lower() != ".csv":
            raise ValueError(f"VBRL se guarda como .csv (se recibió '{FileName}').")
        lista = _normalizar_facturas_vbrl(facturas)
        os.makedirs(FolderPath, exist_ok=True)

        # Leftovers of an earlier failed run must not be mistaken for this run's
        # output (and an existing .xlsx would make SAP ask to overwrite it).
        for resto in (_ruta_xlsx, _ruta_parcial):
            if os.path.exists(resto):
                try:
                    os.remove(resto)
                except OSError as e:
                    raise RuntimeError(
                        f"No pude borrar un temporal anterior (¿está abierto en Excel?): {resto}"
                    ) from e

        # === SAP connection ===
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

        # A popup left open by the previous download (e.g. an unconfirmed "end the
        # list?") would block the transaction code field.
        _cerrar_popups_abiertos(session)

        # === Open ZSE16 on table VBRL (with /n so it starts from any screen) ===
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "/n" + _TX_ZSE16
        session.findById("wnd[0]").sendVKey(0)
        campo_tabla = session.findById("wnd[0]/usr/ctxtGV_TABNAME")
        campo_tabla.text = _TABLA_VBRL
        campo_tabla.caretPosition = len(_TABLA_VBRL)
        session.findById("wnd[0]").sendVKey(0)

        # === Paste the invoices into the first selection field via the CLIPBOARD ===
        # NOTE: like the Gallo accounts, this OVERWRITES whatever the user had
        # copied. The text is read back and re-checked before importing it.
        print(f"[VBRL] {len(lista):,} factura(s) únicas a consultar.")
        session.findById(_VBRL_BTN_SELECCION_MULTIPLE).press()
        _copiar_y_verificar_portapapeles("\r\n".join(lista))
        _importar_desde_portapapeles_popup(session)
        time.sleep(1 + len(lista) / 5000)          # let SAP load the list (more for long ones)
        session.findById("wnd[1]/tbar[0]/btn[8]").press()      # copy and close the popup

        # === Lift 'Maximum No. of Hits' (checked: see _quitar_limite_resultados) ===
        _quitar_limite_resultados(session)

        # === Run the table browser ===
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        _esperar_sap_libre(session)
        time.sleep(4)

        # === No entries? ===
        if _verificar_sin_partidas(session, ruta_completa, _MENSAJES_SIN_ENTRADAS_SE16):
            return False  # empty file already created

        # Still on the selection screen with a message we do not know: stop
        # instead of guessing (a wrong "no data" would hide a real problem).
        if _existe_en_sap(session, "wnd[0]/usr/txtMAX_SEL"):
            try:
                aviso = session.findById("wnd[0]/sbar").Text.strip()
            except Exception:
                aviso = ""
            raise RuntimeError(
                "ZSE16 no mostró la lista de resultados. SAP dice: "
                f"'{aviso or '(sin mensaje)'}'. Si significa 'sin entradas', agrega ese texto "
                "a _MENSAJES_SIN_ENTRADAS_SE16."
            )

        # === Layout (looked up by name: _VBRL_LAYOUT_NOMBRE) ===
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        _elegir_layout_por_nombre(session)
        _esperar_sap_libre(session)

        # === Export: List > Export > Spreadsheet ===
        session.findById(_VBRL_MENU_EXPORTAR).select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()          # default format
        try:
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = FolderPath
        except Exception:
            pass  # without the folder field SAP saves in its default one; the wait below would time out
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = _nombre_xlsx
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(_nombre_xlsx)
        t_exportacion = time.time()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # Wait for the file instead of a fixed pause: VBRL can be long.
        if not _esperar_archivo_nuevo(_ruta_xlsx, t_exportacion):
            raise RuntimeError(f"SAP no terminó de escribir el archivo de VBRL: {_ruta_xlsx}")

        # === FIRST leave the list in SAP → it releases the .xlsx (3 backs, as recorded) ===
        for _ in range(3):
            try:
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
            except Exception:
                break
        time.sleep(2)

        # === THEN close any regular Excel that was left open ===
        _cerrar_excel_workbook(_ruta_xlsx)

        # === Convert, check, and only then replace the final file ===
        _xlsx_a_csv(_ruta_xlsx, _ruta_parcial)
        _validar_csv_vbrl(_ruta_parcial, lista)
        try:
            os.replace(_ruta_parcial, ruta_completa)
        except OSError as e:
            raise RuntimeError(
                f"No pude guardar {ruta_completa} (¿está abierto en Excel?). "
                "Ciérralo y vuelve a descargar."
            ) from e

        print(f"[VBRL] Archivo descargado correctamente: {ruta_completa}")
        return True

    except pywintypes.com_error as e:
        print(f"[ERROR COM SAP] VBRL_ZSE16: {e}")
        return False
    except (ValueError, TypeError, ConnectionError, RuntimeError) as e:
        print(f"[ERROR CONTROLADO] VBRL_ZSE16: {e}")
        return False
    except Exception as e:
        print(f"[ERROR NO CONTROLADO] VBRL_ZSE16: {e}")
        print(traceback.format_exc())
        return False
    finally:
        if os.path.exists(_ruta_parcial):
            try:
                os.remove(_ruta_parcial)
            except OSError:
                pass
        if session is not None:
            _cerrar_popups_abiertos(session)
            try:
                session.findById("wnd[0]/tbar[0]/btn[3]").press()
            except Exception:
                pass


# ----------------------------------------------------------------------
# Direct run (quick test without GUI or controller)
# ----------------------------------------------------------------------
if __name__ == "__main__":
    print("=== Prueba de descarga (standalone) ===")
    print("1) Gallo (FAGLL03)")
    print("2) Monivoi (ZLMXCOM_TRN_MONINVOI)")
    print("3) VBRL (ZSE16) — con las facturas del archivo de Monivoi del mismo periodo")
    print("4) Venta EIAP (FAGLL03, cuentas EIAP) — la fecha hasta se consulta menos 1 día")
    opcion = input("Elige transacción [1/2/3/4]: ").strip()

    date_from = input("Fecha desde (DD.MM.YYYY): ").strip()
    date_to   = input("Fecha hasta (DD.MM.YYYY): ").strip()

    if opcion == "3":
        # VBRL is not downloaded by date: the dates only NAME the files (same
        # rule as the GUI: one name for a single day, another for a range) and
        # locate the Monivoi file the invoices are read from.
        from Consolidacion import leer_facturas_monivoi
        from utils import (
            nombre_archivo_rango_monivoi, nombre_archivo_dia_vbrl,
            nombre_archivo_rango_vbrl,
        )
        un_dia = date_from == date_to
        archivo_monivoi = (
            nombre_archivo_dia_monivoi(date_to) if un_dia
            else nombre_archivo_rango_monivoi(date_from, date_to)
        )
        archivo = (
            nombre_archivo_dia_vbrl(date_to) if un_dia
            else nombre_archivo_rango_vbrl(date_from, date_to)
        )
        ruta_monivoi = os.path.join(ruta_input_default(), archivo_monivoi)
        print(f"\nFacturas leídas de    : {ruta_monivoi}")
        facturas, _filas, _col = leer_facturas_monivoi(ruta_monivoi)
        print(f"El archivo se nombrará : {archivo}")
        print(f"Se guardará en         : {ruta_input_default()}\n")
        ok = VBRL_ZSE16(facturas, FileName=archivo)
    elif opcion == "4":
        try:
            archivo = nombre_archivo_rango_eiap(date_from, date_to)
            print(f"\nEl archivo se nombrará : {archivo}")
            print(f"Se consultará en SAP   : {date_from} — "
                  f"{fecha_efectiva_eiap(date_to).strftime(FORMATO_FECHA)}")
            print(f"Se guardará en         : {ruta_input_default()}\n")
            ok = VentaEIAP_FAGLL03(date_from, date_to)
        except ValueError as e:
            print(f"❌ {e}")
            ok = False
    elif opcion == "2":
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
