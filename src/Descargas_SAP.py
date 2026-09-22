"""
Validación Factura Global — Descargas SAP
==========================================

SAP transaction extraction module for the 'Validación Factura Global'
project.

Implemented transactions:
  - 'Gallo'   (FAGLL03)
  - 'Monivoi' (ZLMXCOM_TRN_MONINVOI)

Based on the original VBS recordings, with these improvements:
  1. The dates are entered by the user (no longer hard-coded).
  2. The file name uses the SAME date the user typed, in D_mmm_YY format
     (Spanish month, no leading zero on the day; see utils.py).
  3. Supports downloading a single day or a range (day by day).
  4. The folder and the file name are typed straight into the export dialog
     (more robust than navigating with F4 like the VBS did).
"""

import os
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
    nombre_archivo_dia,
    nombre_archivo_dia_monivoi,
    fecha_efectiva_monivoi_str,
)

# ----------------------------------------------------------------------
# Transaction constants (taken from the original VBS recording)
# ----------------------------------------------------------------------

# Accounts in the VBS multiple selection (SD_SAKNR).
_CUENTAS_GALLO = [
    "4000005", "4000010", "4000020", "4000802",
    "4150000", "4200000", "4200004", "4200015",
    "4260000", "4300013", "7000005",
]

# Default company code (BUKRS) from the VBS.
_SOCIEDAD_DEFAULT = "MX21"

# Default display layout / variant (PA_VARI).
_VARIANTE_DEFAULT = "BASE/VALID"


# ----------------------------------------------------------------------
# Monivoi transaction constants (ZLMXCOM_TRN_MONINVOI)
# ----------------------------------------------------------------------
_TX_MONIVOI    = "ZLMXCOM_TRN_MONINVOI"   # custom (Z) transaction
_VKORG_DEFAULT = "MX21"                    # sales organization
_VTWEG_DEFAULT = "01"                      # distribution channel
_RFC_DEFAULT   = "XAXX010101000"           # generic RFC ("público en general" = general public)


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


def _verificar_sin_partidas(session, ruta_archivo):
    """
    Detect SAP's 'No se ha seleccionado ninguna partida' (no items selected)
    message. If it shows up: close the popup, create an empty Excel/CSV as a
    placeholder and return True so the caller knows there is no data.

    NOTE: the message codes are tuned for FAGLL03 (Gallo). For Monivoi (Z)
    the 'no data' text may be different; if your tests show it is not
    detected, add the exact message to MENSAJES_SIN_DATOS.
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


def _importar_cuentas_desde_popup(session):
    """
    Press the 'Importar desde portapapeles' (import from clipboard) button of
    the multiple-selection popup. Tries the known IDs in case it changes
    between SAP systems.
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
        _importar_cuentas_desde_popup(session)
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
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        time.sleep(3)  # let SAP finish WRITING the file

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
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # === Close the Excel that SAP opens + convert to CSV ===
        time.sleep(3)
        _cerrar_excel_workbook(_ruta_xlsx)
        if _exportar_a_csv:
            _xlsx_a_csv(_ruta_xlsx, ruta_completa)

        # === Go back to the initial screen ===
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
# Direct run (quick test without GUI or controller)
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
