"""
Validación Factura Global — SAP downloads (SAP GUI scripting)

Gallo (FAGLL03), Monivoi (ZLMXCOM_TRN_MONINVOI), VBRL (ZSE16), FB03 (document list) and
Venta EIAP (FAGLL03 with the EIAP accounts). Every download follows its VBS recording; the
extra steps are guards that stop the run instead of exporting wrong data.

    SAP, Portapapeles, Excel, Archivos      helpers shared by every download
    Paralelo                                several SAP windows at once (the P&L way)
    Descarga → Gallo, Monivoi; DescargaConLista → VBRL, FB03
    Gallo_FAGLL03(), Monivoi_ZLMXCOM()...   the entry points the controller calls
"""

import contextlib
import ctypes
import gc
import os
import re
import threading
import time
import traceback
from ctypes import wintypes
from datetime import datetime

import openpyxl as _oxl          # also makes PyInstaller bundle the Excel engine
import pandas as pd
import pythoncom
import pywintypes
import win32clipboard
import win32com.client

from utils import (
    FORMATO_FECHA,
    parse_fecha,
    nombre_archivo_dia,
    nombre_archivo_dia_monivoi,
    nombre_archivo_rango_eiap,
    fecha_efectiva_monivoi_str,
    fecha_efectiva_eiap,
    CUENTAS_EIAP,
    CLASES_EIAP,
)

MAX_VENTANAS_SAP = 4                                           # corporate limit of SAP windows per user
_SOCIEDAD = "MX21"                                             # default company code (BUKRS)
_ESPERA_CONSULTA = 1800        # upper bound (s) for a report; the wait returns as soon as SAP is free
_MSG_SIN_SAP = (
    "No fue posible conectarse a SAP GUI. "
    "Verifica que SAP esté abierto y con una sesión activa."
)


class ExportacionTruncadaError(RuntimeError):
    """The export reached Excel's row limit, so it may be incomplete. Propagates: the summary shows it."""


class DescargaPorRevisarError(RuntimeError):
    """SAP exported, but something failed afterwards; the message says what, and which file was kept."""


# ======================================================================
# SAP GUI
# ======================================================================
class SAP:
    """The connection to SAP GUI and the small navigation helpers every download uses."""

    @staticmethod
    def aplicacion():
        """The scripting engine (GuiApplication). ConnectionError if SAP GUI is not there."""
        try:
            return win32com.client.GetObject("SAPGUI").GetScriptingEngine
        except Exception as e:
            raise ConnectionError(_MSG_SIN_SAP) from e

    @staticmethod
    def conexion():
        """The first connection (the logon the user works in)."""
        aplicacion = SAP.aplicacion()
        try:
            return aplicacion.Children(0)
        except Exception as e:
            raise ConnectionError(_MSG_SIN_SAP) from e

    @staticmethod
    def sesion():
        """Its first window: the one the sequential downloads work on."""
        conexion = SAP.conexion()
        try:
            return conexion.Children(0)
        except Exception as e:
            raise ConnectionError(_MSG_SIN_SAP) from e

    @staticmethod
    def existe(session, id_componente):
        """True if the component exists (findById raises when it does not)."""
        try:
            session.findById(id_componente)
            return True
        except Exception:
            return False

    @staticmethod
    def esperar(session, id_componente, timeout=5.0, intervalo=0.25):
        """Wait until the component exists; False if it still does not after `timeout` s."""
        limite = time.time() + timeout
        while True:
            if SAP.existe(session, id_componente):
                return True
            if time.time() >= limite:
                return False
            time.sleep(intervalo)

    @staticmethod
    def esperar_cierre(session, ventana="wnd[1]", timeout=3):
        """Wait until `ventana` is gone; False if it is still open after `timeout` s."""
        limite = time.time() + timeout
        while SAP.existe(session, ventana) and time.time() < limite:
            time.sleep(0.25)
        return not SAP.existe(session, ventana)

    @staticmethod
    def esperar_libre(session, timeout=300):
        """Wait while SAP is busy with this window (long queries). Never raises."""
        limite = time.time() + timeout
        try:
            while session.Busy and time.time() < limite:
                time.sleep(0.5)
        except Exception:
            pass

    @staticmethod
    def cerrar_popups(session, encima_de=0, maximo=5):
        """Cancel (F12) the popups above wnd[encima_de], the topmost first. Best effort."""
        for _ in range(maximo):
            ventana = next(
                (f"wnd[{n}]" for n in range(encima_de + 4, encima_de, -1)
                 if SAP.existe(session, f"wnd[{n}]")), None
            )
            if ventana is None:
                return
            try:
                session.findById(ventana).sendVKey(12)
            except Exception:
                return

    @staticmethod
    def aviso(session):
        """Text of the status bar ('' if it cannot be read)."""
        try:
            return session.findById("wnd[0]/sbar").Text.strip()
        except Exception:
            return ""

    @staticmethod
    def sin_partidas(session, ruta_archivo, mensajes_extra=()):
        """
        SAP said "no items" (FAGLL03 codes, or the texts a transaction passes, lower case):
        close the message, write the 'Sin datos' placeholder and return True. Never raises.
        """
        try:
            mensajes = ("MSITEM033", "MSITEM030", "ninguna partida", "ninguna cuenta") + tuple(mensajes_extra)
            textos = [SAP.aviso(session)]
            try:
                textos.append(session.findById("wnd[1]/usr/txtMESSTXT1").Text.strip())
            except Exception:
                pass
            if not any(m in t or m in t.lower() for t in textos for m in mensajes):
                return False
            print(f"[SIN DATOS] Creando archivo vacío: {ruta_archivo}")
            for boton in ("wnd[1]/tbar[0]/btn[0]", "wnd[0]/tbar[0]/btn[3]"):
                try:
                    session.findById(boton).press()
                except Exception:
                    pass
            Archivos.crear_sin_datos(ruta_archivo)
            return True
        except Exception as e:
            print(f"[ERROR] Revisando si SAP dijo 'sin datos': {e}")
            return False

    @staticmethod
    def elementos_de_lista(session, id_area):
        """Elements of a classic list area: [(component, Type, label text, (column, row) or None)]."""
        area = session.findById(id_area)
        elementos = []
        for i in range(int(area.Children.Count)):
            hijo = area.Children(i)
            tipo = str(hijo.Type)
            texto = str(hijo.Text).strip() if tipo == "GuiLabel" else ""
            m = re.search(r"\[(\d+),(\d+)\]$", str(hijo.Id))           # '.../lbl[4,5]' → (4, 5)
            elementos.append((hijo, tipo, texto, (int(m.group(1)), int(m.group(2))) if m else None))
        return elementos

    @staticmethod
    def renglones_de_lista(elementos):
        """
        Those elements by list row: {row: {"texto": its labels left to right,
        "etiquetas": [(column, label, text)], "casillas": [checkboxes]}}.
        """
        renglones = {}
        for componente, tipo, texto, posicion in elementos:
            if posicion is None:
                continue
            columna, fila = posicion
            renglon = renglones.setdefault(fila, {"etiquetas": [], "casillas": []})
            if tipo == "GuiLabel" and texto:
                renglon["etiquetas"].append((columna, componente, texto))
            elif tipo == "GuiCheckBox":
                renglon["casillas"].append(componente)
        for renglon in renglones.values():
            renglon["etiquetas"].sort(key=lambda etiqueta: etiqueta[0])
            renglon["texto"] = "  ".join(t for _, _, t in renglon["etiquetas"])
        return renglones

    @staticmethod
    def resumen_renglones(renglones, maximo=15):
        """'5: VBELN  Factura; 6: ...' (for error messages)."""
        vistos = [f"{fila}: {renglones[fila]['texto'][:80]}" for fila in sorted(renglones)
                  if renglones[fila]["texto"]]
        return "; ".join(vistos[:maximo]) or "(ninguno)"


# ======================================================================
# Clipboard
# ======================================================================
class Portapapeles:
    """
    The Windows clipboard, ONE for the whole machine. Every list pasted into SAP goes
    write → read back → SAP imports it inside exclusivo(), so another SAP window cannot
    overwrite it halfway. NOTE: it replaces whatever the user had copied.
    """
    _lock = threading.RLock()
    ESPERA = 300                   # s a window waits for another one to release it

    @staticmethod
    def copiar(texto, intentos=5, espera=0.3):
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
            except Exception as e:             # another app may hold it for an instant
                ultimo_error = e
                time.sleep(espera)
        raise RuntimeError(f"No se pudo escribir en el portapapeles: {ultimo_error}")

    @staticmethod
    def leer(intentos=5, espera=0.3):
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

    @staticmethod
    def copiar_y_verificar(texto, intentos=3):
        """Copy and read back (pasting a list changed by someone else would query the wrong data)."""
        for _ in range(intentos):
            Portapapeles.copiar(texto)
            if Portapapeles.leer().replace("\x00", "") == texto:
                return True
            time.sleep(0.3)
        raise RuntimeError(
            "El portapapeles cambió mientras se preparaba la lista para SAP. "
            "No copies nada mientras corre la descarga e inténtalo de nuevo."
        )

    @classmethod
    @contextlib.contextmanager
    def exclusivo(cls):
        """Hold the clipboard from 'write' until SAP imported it. Re-entrant; gives up after ESPERA s."""
        if not cls._lock.acquire(timeout=cls.ESPERA):
            raise RuntimeError(
                "Otra ventana de SAP tuvo ocupado el portapapeles más de "
                f"{cls.ESPERA // 60} minutos: no pude pegar la lista."
            )
        try:
            yield
        finally:
            cls._lock.release()

    @staticmethod
    def importar_en_popup(session):
        """'Import from clipboard' of the multiple-selection popup (its Id varies between systems)."""
        for boton in ("wnd[1]/tbar[0]/btn[24]", "wnd[1]/tbar[0]/btn[25]"):
            try:
                session.findById(boton).press()
                return True
            except Exception:
                continue
        raise RuntimeError(
            "No encontré el botón 'Importar desde portapapeles' en el popup. "
            "Confirma su ID con Script Recording (Alt+F12) y agrégalo a la lista."
        )


# ======================================================================
# Excel (shared by every SAP window)
# ======================================================================
class Excel:
    """
    After an export SAP opens the file in Excel, and while Excel holds it the temporary
    .xlsx cannot be deleted. Excel is SHARED by every SAP window: during a parallel run
    (modo_paralelo) only the file of the own export is closed; quitting empty instances
    and killing a stuck Excel wait until nothing is in flight (limpiar_tras_paralelo).
    """
    _lock = threading.Lock()
    _nivel = 0                     # > 0 while a parallel run is active
    _pendientes = []               # temporary .xlsx that stayed locked during a parallel run

    @staticmethod
    def terminar_proceso(pid):
        """Force-close process `pid` (Windows). True if it succeeded."""
        PROCESS_TERMINATE = 0x0001
        k32 = ctypes.windll.kernel32
        k32.OpenProcess.restype = wintypes.HANDLE
        k32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        k32.TerminateProcess.argtypes = [wintypes.HANDLE, wintypes.UINT]
        k32.CloseHandle.argtypes = [wintypes.HANDLE]
        handle = k32.OpenProcess(PROCESS_TERMINATE, False, int(pid))
        if not handle:
            return False
        try:
            return bool(k32.TerminateProcess(handle, 1))
        finally:
            k32.CloseHandle(handle)

    @staticmethod
    def liberar_archivo(ruta_archivo):
        """
        Last resort for a file still locked: ask the Windows Restart Manager which processes
        hold it and terminate only the ones named Excel, never SAP (typical in Gallo: an Excel
        stuck on a modal dialog ignores COM but keeps the file). The data is already on disk;
        other workbooks open in that same Excel would lose their unsaved changes.
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
            clave = (ctypes.c_wchar * (32 + 1))()          # CCH_RM_SESSION_KEY = 32
            if rstrtmgr.RmStartSession(ctypes.byref(sesion), 0, clave) != 0:
                print("[RM] No se pudo iniciar la sesión de Restart Manager.")
                return False

            liberado = False
            try:
                archivos = (ctypes.c_wchar_p * 1)(os.path.abspath(ruta_archivo))
                if rstrtmgr.RmRegisterResources(sesion, 1, archivos, 0, None, 0, None) != 0:
                    print("[RM] No se pudo registrar el archivo.")
                    return False

                necesarios = wintypes.UINT(0)
                cuantos = wintypes.UINT(0)
                razon = wintypes.DWORD(0)
                # 1st call: how many processes (it returns ERROR_MORE_DATA)
                rstrtmgr.RmGetList(sesion, ctypes.byref(necesarios),
                                   ctypes.byref(cuantos), None, ctypes.byref(razon))
                if necesarios.value == 0:
                    print(f"[RM] Ningún proceso tiene agarrado: {ruta_archivo}")
                    return False

                cuantos = wintypes.UINT(necesarios.value)
                infos = (RM_PROCESS_INFO * necesarios.value)()
                if rstrtmgr.RmGetList(sesion, ctypes.byref(necesarios),
                                      ctypes.byref(cuantos), infos, ctypes.byref(razon)) != 0:
                    print("[RM] No se pudo obtener la lista de procesos.")
                    return False

                for i in range(cuantos.value):
                    nombre = infos[i].strAppName or "<desconocido>"
                    pid = int(infos[i].Process.dwProcessId)
                    print(f"[RM] '{os.path.basename(ruta_archivo)}' lo tiene: {nombre} (PID {pid})")
                    if "excel" in nombre.lower():                 # ONLY Excel
                        if Excel.terminar_proceso(pid):
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

    @staticmethod
    def instancias():
        """
        EVERY Excel instance of the machine, found through the workbooks registered in the
        Running Object Table (SAP often opens its export in an instance of its own, which
        GetObject alone would never reach), plus the main one. Never raises.
        """
        instancias = {}
        try:
            rot = pythoncom.GetRunningObjectTable()
            ctx = pythoncom.CreateBindCtx(0)
            for moniker in rot.EnumRunning():
                try:
                    nombre = moniker.GetDisplayName(ctx, None)
                except Exception:
                    continue
                if not nombre or not nombre.lower().endswith((".xls", ".xlsx", ".xlsm", ".xlsb", ".csv")):
                    continue
                try:
                    app = win32com.client.Dispatch(rot.GetObject(moniker)).Application
                    instancias[int(app.Hwnd)] = app
                except Exception:
                    continue
        except Exception:
            pass
        try:
            app = win32com.client.GetObject(Class="Excel.Application")
            instancias[int(app.Hwnd)] = app
        except Exception:
            pass
        return list(instancias.values())

    @staticmethod
    def cerrar_si_vacia(excel):
        """Quit this instance only if it holds no workbook and no Protected View window."""
        try:
            if int(excel.Workbooks.Count) != 0:
                return False
            try:
                if int(excel.ProtectedViewWindows.Count) != 0:
                    return False
            except Exception:
                pass
            excel.Quit()
            return True
        except Exception:
            return False

    @classmethod
    def cerrar_libro(cls, ruta_archivo, intentos=3, espera=1):
        """
        Close (without saving) the exported file in every Excel instance: regular workbooks (by
        full path, or by name: SAP may report the path in another shape) and Protected View
        windows (Excel opens big exports there, outside Workbooks). An instance left empty is
        quit, except during a parallel run. Retries in case Excel had not finished opening it.
        """
        ruta_abs = os.path.normcase(os.path.abspath(ruta_archivo))
        objetivo = os.path.normcase(os.path.basename(ruta_archivo))
        for _ in range(intentos):
            instancias = cls.instancias()
            if not instancias:
                return False
            cerrado = False
            for excel in instancias:
                try:
                    excel.DisplayAlerts = False
                except Exception:
                    pass
                try:
                    for wb in list(excel.Workbooks):
                        try:
                            if (os.path.normcase(os.path.abspath(wb.FullName)) == ruta_abs
                                    or os.path.normcase(str(wb.Name)) == objetivo):
                                wb.Close(SaveChanges=False)
                                cerrado = True
                                print(f"[EXCEL] Workbook cerrado: {ruta_archivo}")
                        except Exception:
                            continue
                    try:
                        for pvw in list(excel.ProtectedViewWindows):
                            try:
                                info = ""
                                for prop in ("SourceName", "Caption"):
                                    try:
                                        info += os.path.normcase(str(getattr(pvw, prop)))
                                    except Exception:
                                        pass
                                if objetivo in info:
                                    pvw.Close()
                                    cerrado = True
                                    print(f"[EXCEL] Vista protegida cerrada: {ruta_archivo}")
                            except Exception:
                                continue
                    except Exception:
                        pass
                    # The empty Excel may be the one ANOTHER window's export is starting up in.
                    if not cls.activo() and cls.cerrar_si_vacia(excel):
                        print("[EXCEL] Instancia de Excel cerrada (sin libros abiertos)")
                finally:
                    try:
                        excel.DisplayAlerts = True
                    except Exception:
                        pass
            if cerrado:
                return True
            time.sleep(espera)
        return False

    @classmethod
    def diagnostico(cls, ruta_archivo):
        """Print what every Excel instance has open (the clue when a file stays locked)."""
        try:
            instancias = cls.instancias()
            print(f"[EXCEL][DIAG] No se pudo cerrar '{os.path.basename(ruta_archivo)}'. "
                  f"Instancias de Excel detectadas: {len(instancias)}")
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

    @classmethod
    def borrar_xlsx(cls, ruta_xlsx, intentos=10, espera=2):
        """
        Delete the temporary .xlsx, closing it in Excel before every attempt (a cold Excel takes
        a while to open a big file). Halfway through, the Restart Manager may close the Excel
        that holds it, except during a parallel run: then the file waits for limpiar_tras_paralelo.
        """
        for intento in range(1, intentos + 1):
            cls.cerrar_libro(ruta_xlsx)
            try:
                if os.path.exists(ruta_xlsx):
                    os.remove(ruta_xlsx)
                print(f"[XLSX] Temporal borrado: {ruta_xlsx}")
                return True
            except OSError as e:
                print(f"[XLSX] Aún bloqueado (intento {intento}/{intentos}): {e}")
                if intento == max(1, intentos // 2):
                    if cls.activo():
                        print("[XLSX] Corrida en paralelo: no cierro procesos de Excel ahora; "
                              "se reintenta cuando terminen todas las descargas.")
                    else:
                        cls.liberar_archivo(ruta_xlsx)
                time.sleep(espera)
        print(f"[XLSX] ⚠️ No se pudo borrar (sigue abierto/bloqueado): {ruta_xlsx}")
        cls.diagnostico(ruta_xlsx)
        if cls.activo():
            with cls._lock:
                cls._pendientes.append(ruta_xlsx)
        return False

    @classmethod
    def activo(cls):
        """True while a parallel run is active."""
        return cls._nivel > 0

    @classmethod
    def limpiar_tras_paralelo(cls):
        """Once nothing is in flight: delete the .xlsx left locked, quit empty instances. Never raises."""
        try:
            with cls._lock:
                pendientes = list(dict.fromkeys(cls._pendientes))
                del cls._pendientes[:]
            for ruta in pendientes:
                if os.path.exists(ruta):
                    print(f"[XLSX] Limpieza final: {ruta}")
                    cls.borrar_xlsx(ruta)
            for excel in cls.instancias():
                if cls.cerrar_si_vacia(excel):
                    print("[EXCEL] Instancia de Excel cerrada (sin libros abiertos)")
        except Exception as e:
            print(f"[EXCEL] La limpieza final falló: {e}")

    @classmethod
    @contextlib.contextmanager
    def modo_paralelo(cls):
        """Wrap each parallel pass with it; when the OUTERMOST one ends, limpiar_tras_paralelo() runs."""
        with cls._lock:
            cls._nivel += 1
        try:
            yield
        finally:
            with cls._lock:
                cls._nivel -= 1
                ultimo = cls._nivel == 0
            if ultimo:
                cls.limpiar_tras_paralelo()


# ======================================================================
# Files
# ======================================================================
class Archivos:
    """The files of a download: input folder, .xlsx → .csv, waiting for SAP to write, checks."""
    MAX_FILAS = 1_048_576 - 1      # data rows an .xlsx holds: a result that reaches it may have been cut
    ESPERA_EXPORTACION = 900       # upper bound (s) for SAP to finish writing a big export

    @staticmethod
    def ruta_input_default():
        """Documents\\Validacion Factura Global\\src\\Input of the current Windows user."""
        perfil = os.environ.get("USERPROFILE") or os.path.expanduser("~")
        return os.path.join(perfil, "Documents", "Validacion Factura Global", "src", "Input")

    @staticmethod
    def xlsx_a_csv(ruta_xlsx, ruta_csv, eliminar_xlsx=True):
        """SAP only exports .xlsx; every download ends up as a utf-8-sig .csv (Excel shows the accents)."""
        df = pd.read_excel(ruta_xlsx, dtype=str, engine="openpyxl")
        if len(df) >= Archivos.MAX_FILAS:
            filas = len(df)
            del df
            if eliminar_xlsx:
                Excel.borrar_xlsx(ruta_xlsx)
            raise ExportacionTruncadaError(
                f"El export de SAP llegó al límite de filas de Excel ({filas:,}): el archivo "
                "pudo quedar INCOMPLETO y no se usó. Descarga un periodo más corto (menos "
                "días) y vuelve a intentar."
            )
        df.to_csv(ruta_csv, index=False, encoding="utf-8-sig")
        if eliminar_xlsx:
            Excel.borrar_xlsx(ruta_xlsx)

    @staticmethod
    def crear_sin_datos(ruta_archivo):
        """The 'Sin datos' placeholder: 'downloaded, but empty' is not the same as 'not downloaded'."""
        try:
            carpeta = os.path.dirname(ruta_archivo)
            if carpeta:
                os.makedirs(carpeta, exist_ok=True)
            if ruta_archivo.lower().endswith(".csv"):
                pd.DataFrame({"Sin datos": []}).to_csv(ruta_archivo, index=False, encoding="utf-8-sig")
            else:
                wb = _oxl.Workbook()
                wb.active.title = "Sin datos"
                wb.save(ruta_archivo)
                wb.close()
        except Exception as e:
            raise RuntimeError(f"No se pudo crear el archivo vacío: {ruta_archivo}") from e

    @staticmethod
    def esperar_nuevo(ruta, desde, timeout=90, estable=1.0):
        """
        True once SAP finished WRITING `ruta`: it exists, it was modified after `desde` (a file
        of an earlier run must not count) and its size did not change for `estable` s.
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

    @staticmethod
    def esperar_exportacion(ruta_xlsx, desde):
        """Big exports (Gallo, Monivoi) are written in bursts: size unchanged for 3 s; late → RuntimeError."""
        if not Archivos.esperar_nuevo(ruta_xlsx, desde, timeout=Archivos.ESPERA_EXPORTACION, estable=3.0):
            raise RuntimeError(f"SAP no terminó de escribir el archivo exportado: {ruta_xlsx}")

    @staticmethod
    def normalizar_documentos(documentos, etiqueta):
        """
        The document numbers for the clipboard: text, trimmed, no blanks or duplicates (first
        order kept), without the '.0' of a float. ValueError if nothing is left: with an empty
        selection field SAP would read the WHOLE table.
        """
        if documentos is None or isinstance(documentos, (str, bytes)):
            raise ValueError("La lista de documentos debe ser una colección (lista), no un texto suelto.")
        vistos = set()
        limpios = []
        for valor in documentos:
            if valor is None:
                continue
            texto = str(valor).strip()
            con_cero_decimal = re.fullmatch(r"([0-9]+)\.0+", texto)
            if con_cero_decimal:
                texto = con_cero_decimal.group(1)
            if not texto:
                continue
            if any(c in texto for c in "\r\n\t\x00"):         # one document per clipboard line
                raise ValueError(f"Número de documento con caracteres no válidos: {texto!r}")
            if texto not in vistos:
                vistos.add(texto)
                limpios.append(texto)
        if not limpios:
            raise ValueError(
                f"No hay documentos que consultar en {etiqueta} (la lista está vacía): "
                "sin documentos SAP intentaría leer la tabla completa."
            )
        return limpios

    @staticmethod
    def validar_csv(ruta_csv, documentos, etiqueta, unidad=("documento(s)", "pedido(s)"), revisar=None):
        """
        Check a converted download before it replaces the final file: it must have rows and
        pass `revisar(df)` (RuntimeError otherwise). Then report how many of the requested
        documents appear in its best column (only reported: the table may key its rows by
        another folio). Returns (rows, requested, found).
        """
        df = pd.read_csv(ruta_csv, dtype=str, encoding="utf-8-sig", keep_default_na=False)
        if df.shape[1] < 1 or df.empty:
            raise RuntimeError(f"El archivo exportado de {etiqueta} no trae filas.")
        if revisar is not None:
            revisar(df)
        pedidos = {int(d) for d in documentos if re.fullmatch(r"[0-9]+", d)}
        mejor, mejor_columna = 0, None
        for columna in df.columns:
            valores = df[columna].astype(str).str.strip()
            valores = valores[valores.str.fullmatch(r"[0-9]+(\.0+)?")]
            coinciden = len(pedidos & {int(v.split(".", 1)[0]) for v in valores})
            if coinciden > mejor:
                mejor, mejor_columna = coinciden, columna
        donde = f"en la columna '{mejor_columna}'" if mejor_columna is not None else "en ninguna columna"
        print(f"[{etiqueta}] {len(df):,} fila(s); {mejor:,} de {len(pedidos):,} {unidad[0]} "
              f"{unidad[1]} aparecen {donde}.")
        if pedidos and not mejor:
            print(f"[{etiqueta}] ⚠️ Ningún documento pedido aparece en el archivo. Si la tabla usa otro "
                  "folio es normal; si no, revisa el layout elegido.")
        return len(df), len(pedidos), mejor


# ======================================================================
# Several SAP windows at once
# ======================================================================
class Paralelo:
    """
    The P&L way: make sure the first connection has enough windows (the ones already open
    are used as they are; the missing ones are opened with CreateSession and stay open),
    then one thread per window: it initializes COM, takes its window BY INDEX
    (connection.Children(i)) and runs its download. Each thread can label the console
    lines it prints (etiqueta/marca) to show whether the windows really overlapped.
    """
    ESPERA_NUEVA_VENTANA = 20      # s SAP may take to show a window opened with CreateSession
    DESFASE = 0                    # s between the starts of two windows of a pass (P&L used 30)
    _hilo = threading.local()

    @staticmethod
    def asegurar_ventanas(deseadas):
        """
        Open windows until the first connection has `deseadas` (at most MAX_VENTANAS_SAP).
        Returns (windows available, up to `deseadas`; windows opened now).
        """
        deseadas = min(int(deseadas), MAX_VENTANAS_SAP)
        if deseadas < 1:
            raise ValueError("Se necesita al menos una ventana de SAP.")
        conexion = SAP.conexion()
        abiertas = 0
        while int(conexion.Children.Count) < deseadas:
            antes = int(conexion.Children.Count)
            primera = conexion.Children(0)
            try:
                primera.CreateSession()           # a GuiSession method (not of the connection)
            except Exception as e:
                print(f"[SAP] No pude abrir otra ventana: {e}")
                break
            limite = time.time() + Paralelo.ESPERA_NUEVA_VENTANA
            while int(conexion.Children.Count) <= antes and time.time() < limite:
                time.sleep(0.25)
            if int(conexion.Children.Count) <= antes:
                # At the limit SAP opens nothing and shows a popup in the calling window.
                print("[SAP] SAP no abrió otra ventana (probablemente se alcanzó el límite de ventanas).")
                SAP.cerrar_popups(primera)
                break
            abiertas += 1
            time.sleep(0.8)                       # let the new window finish starting up
        return min(int(conexion.Children.Count), deseadas), abiertas

    @staticmethod
    def en_ventana(indice, funcion, *args, **kwargs):
        """
        Run funcion(*args, session=<window `indice`>, **kwargs) in the CURRENT thread, inside
        its own COM apartment (a COM object only works in the thread that got it). Whatever
        `funcion` returns is returned and whatever it raises is raised again.
        """
        pythoncom.CoInitialize()
        error = resultado = None
        try:
            try:
                conexion = SAP.conexion()
                try:
                    sesion = conexion.Children(int(indice))
                except Exception as e:
                    raise ConnectionError(f"No encontré la ventana {int(indice) + 1} de SAP: {e}") from e
                resultado = funcion(*args, session=sesion, **kwargs)
            except BaseException as e:
                error = e
            finally:
                conexion = sesion = None
        finally:
            if error is not None:
                Paralelo._soltar_trazas(error)    # tracebacks keep COM proxies alive: drop them first
            gc.collect()
            pythoncom.CoUninitialize()
        if error is not None:
            raise error
        return resultado

    @staticmethod
    def _soltar_trazas(error):
        """Drop the traceback of `error` and of every exception chained to it (cycle-safe)."""
        vistas, pendientes = set(), [error]
        while pendientes:
            actual = pendientes.pop()
            if actual is None or id(actual) in vistas:
                continue
            vistas.add(id(actual))
            actual.__traceback__ = None
            pendientes.extend((actual.__cause__, actual.__context__))

    @classmethod
    def correr(cls, trabajos, al_esperar=None, intervalo=1.0):
        """
        Run every job — (window index, function, args) — in its own thread at the same time and
        wait for all of them. Returns, in order, (True, result) or (False, exception) per job.
        `al_esperar(positions still running)` is called from THIS thread about every `intervalo`
        s (Tk must only be touched by its own thread); if it raises, the windows are still
        waited for before the error goes up.
        """
        salidas = [None] * len(trabajos)

        def correr_uno(posicion, indice, funcion, args):
            try:
                salidas[posicion] = (True, cls.en_ventana(indice, funcion, *args))
            except BaseException as e:
                salidas[posicion] = (False, e)

        hilos = []
        for posicion, (indice, funcion, args) in enumerate(trabajos):
            if hilos and cls.DESFASE > 0:
                time.sleep(cls.DESFASE)
            hilo = threading.Thread(target=correr_uno, args=(posicion, indice, funcion, args),
                                    name=f"sap-ventana-{int(indice) + 1}", daemon=True)
            hilo.start()
            hilos.append(hilo)

        error = None
        while True:
            vivos = {p for p, hilo in enumerate(hilos) if hilo.is_alive()}
            if not vivos:
                break
            if al_esperar is not None and error is None:
                try:
                    al_esperar(vivos)
                except Exception as e:
                    error = e
            hilos[min(vivos)].join(intervalo)
        if error is not None:
            raise error
        return salidas

    @classmethod
    @contextlib.contextmanager
    def etiqueta(cls, texto):
        """Label the timeline lines (marca) of the CURRENT thread, e.g. 'ventana 2 · Monivoi'."""
        anterior = getattr(cls._hilo, "etiqueta", None)
        cls._hilo.etiqueta = texto
        try:
            yield
        finally:
            cls._hilo.etiqueta = anterior

    @classmethod
    def marca(cls, texto):
        """'[HH:MM:SS] [ventana 2 · Monivoi] texto', only in a labelled thread. Never raises."""
        etiqueta = getattr(cls._hilo, "etiqueta", None)
        if etiqueta:
            try:
                print(f"[{datetime.now():%H:%M:%S}] [{etiqueta}] {texto}")
            except Exception:
                pass


# ======================================================================
# Downloads
# ======================================================================
class Descarga:
    """
    What every download shares: the paths, the SAP window, the error handling and the clean-up.
    A subclass writes its recording in pasos(); ejecutar() returns True (data) / False (no data
    or a handled error, already printed). ExportacionTruncadaError and DescargaPorRevisarError
    go up to the controller, which shows them in the summary.
    """
    NOMBRE = ""                    # name in the console messages ('Gallo_FAGLL03')
    MARCA = ""                     # prefix of its timeline lines ('FAGLL03')

    def __init__(self, carpeta, archivo, session):
        self.carpeta = carpeta or Archivos.ruta_input_default()
        self.archivo = archivo
        self.ruta = os.path.join(self.carpeta, archivo)
        self.base, ext = os.path.splitext(archivo)
        self.ext = ext.lower()
        self.a_csv = self.ext == ".csv"           # SAP exports .xlsx: a .csv target is converted at the end
        self.nombre_xlsx = self.base + ".xlsx" if self.a_csv else archivo
        self.ruta_xlsx = os.path.join(self.carpeta, self.nombre_xlsx)
        self._recibida = session                  # adopted by conectar(), so a run that fails its
        self.session = None                       # checks never touches (presses Back in) a window
        self.en_menu = False                      # True once the run left SAP on the main menu

    def ejecutar(self):
        try:
            return self.pasos()
        except pywintypes.com_error as e:
            print(f"[ERROR COM SAP] {self.NOMBRE}: {e}")
        except (ExportacionTruncadaError, DescargaPorRevisarError):
            raise
        except (ValueError, TypeError, ConnectionError, RuntimeError) as e:
            print(f"[ERROR CONTROLADO] {self.NOMBRE}: {e}")
        except Exception as e:
            print(f"[ERROR NO CONTROLADO] {self.NOMBRE}: {e}")
            print(traceback.format_exc())
        finally:
            self.limpiar()
        return False

    def pasos(self):
        raise NotImplementedError

    def conectar(self):
        """The window a parallel run handed over, or the first one (sequential run)."""
        self.session = self._recibida if self._recibida is not None else SAP.sesion()
        return self.session

    def limpiar(self):
        """One Back, only in a window this run used and did not leave on the main menu
        (a Back from the menu itself raises SAP's 'log off?' prompt)."""
        if self.session is not None and not self.en_menu:
            try:
                self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            except Exception:
                pass

    def consultar(self):
        """Execute (F8) and wait for SAP's answer."""
        Paralelo.marca(f"{self.MARCA}: consulta enviada a SAP")
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        SAP.esperar_libre(self.session, timeout=_ESPERA_CONSULTA)
        Paralelo.marca(f"{self.MARCA}: SAP respondió la consulta")
        time.sleep(4)

    def dialogo_guardar(self, boton, carpeta_obligatoria=False):
        """The 'save as' dialog: folder and file typed in, then `boton`.
        Returns (time just before saving, whether the folder field was there)."""
        s = self.session
        con_carpeta = True
        if carpeta_obligatoria:
            s.findById("wnd[1]/usr/ctxtDY_PATH").text = self.carpeta
        else:
            try:
                s.findById("wnd[1]/usr/ctxtDY_PATH").text = self.carpeta
            except Exception:
                con_carpeta = False               # then SAP saves in its own default folder
        s.findById("wnd[1]/usr/ctxtDY_FILENAME").text = self.nombre_xlsx
        s.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(self.nombre_xlsx)
        desde = time.time()
        Paralelo.marca(f"{self.MARCA}: exportando {self.nombre_xlsx}")
        s.findById(boton).press()
        return desde, con_carpeta


class Gallo(Descarga):
    """FAGLL03 ('Gallo'): GL line items of the Gallo accounts for the period, ONE query → ONE file."""
    NOMBRE = "Gallo_FAGLL03"
    MARCA = "FAGLL03"
    CUENTAS = ["4000005", "4150000", "4200000", "4200004", "4200015", "4260000", "7000005"]
    VARIANTE = "BASE/VALID"

    # 'Dynamic selections' (EIAP only). Tree node keys come from a recording: they are system-specific.
    ARBOL = "wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell"
    NODO_PADRE, NODO_CLASE = " " * 9 + "23", " " * 9 + "30"
    BTN_CLASES = "wnd[0]/usr/btn%_%%DYN001_%_APP_%-VALU_PUSH"
    FILA_CLASE = ("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/"
                  "tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,{}]")

    def __init__(self, desde, hasta, carpeta, archivo, sociedad, cuentas, variante, session, clases=None):
        super().__init__(carpeta, archivo or nombre_archivo_dia(hasta), session)
        self.desde, self.hasta = desde, hasta
        self.sociedad = sociedad
        self.cuentas = cuentas or self.CUENTAS
        self.variante = variante or self.VARIANTE
        self.clases = tuple(clases or ())              # document classes to filter by (empty = all)

    def filtrar_clases(self, s):
        """Dynamic selections > document class = self.clases (port of the EIAP recording)."""
        s.findById("wnd[0]/tbar[1]/btn[25]").press()
        arbol = s.findById(self.ARBOL)
        arbol.expandNode(self.NODO_PADRE)
        arbol.selectNode(self.NODO_CLASE)
        arbol.topNode = self.NODO_PADRE
        arbol.doubleClickNode(self.NODO_CLASE)
        s.findById(self.BTN_CLASES).press()
        for i, clase in enumerate(self.clases):
            fila = s.findById(self.FILA_CLASE.format(i))
            fila.text = clase
            fila.setFocus()
            fila.caretPosition = len(clase)
        s.findById("wnd[1]/tbar[0]/btn[8]").press()                   # copy
        if not SAP.esperar_cierre(s):
            raise RuntimeError("El popup de selección de clases de documento sigue abierto.")
        s.findById("wnd[0]/tbar[0]/btn[11]").press()

    def pasos(self):
        for fecha in (self.desde, self.hasta):
            datetime.strptime(fecha, "%d.%m.%Y")          # ValueError if malformed
        os.makedirs(self.carpeta, exist_ok=True)
        s = self.conectar()
        s.findById("wnd[0]").maximize()
        s.findById("wnd[0]/tbar[0]/okcd").text = "/nFAGLL03"
        s.findById("wnd[0]").sendVKey(0)
        if self.clases:
            self.filtrar_clases(s)

        # Accounts: all at once through the clipboard (no row-by-row typing or scrolling).
        s.findById("wnd[0]/usr/btn%_SD_SAKNR_%_APP_%-VALU_PUSH").press()
        with Portapapeles.exclusivo():
            Portapapeles.copiar_y_verificar("\r\n".join(str(c).strip() for c in self.cuentas))
            Portapapeles.importar_en_popup(s)
            time.sleep(1)
        s.findById("wnd[1]/tbar[0]/btn[0]").press()
        s.findById("wnd[1]/tbar[0]/btn[8]").press()

        s.findById("wnd[0]/usr/radX_AISEL").select()                  # all items
        s.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").text = self.sociedad
        s.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = self.desde
        s.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = self.hasta
        campo = s.findById("wnd[0]/usr/ctxtPA_VARI")                  # layout always by name
        campo.text = self.variante
        campo.setFocus()
        campo.caretPosition = len(self.variante)

        self.consultar()
        if SAP.sin_partidas(s, self.ruta):
            return False

        s.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        s.findById("wnd[1]/tbar[0]/btn[0]").press()
        desde, _ = self.dialogo_guardar("wnd[1]/tbar[0]/btn[11]", carpeta_obligatoria=True)
        Archivos.esperar_exportacion(self.ruta_xlsx, desde)
        Paralelo.marca(f"{self.MARCA}: archivo exportado")
        # Leaving the list FIRST makes SAP release an Excel shown "in place"; then close any other.
        s.findById("wnd[0]/tbar[0]/btn[3]").press()
        s.findById("wnd[0]/tbar[0]/btn[3]").press()
        time.sleep(2)
        Excel.cerrar_libro(self.ruta_xlsx)
        if self.a_csv:
            Archivos.xlsx_a_csv(self.ruta_xlsx, self.ruta)
        print(f"[GALLO] Archivo descargado correctamente: {self.ruta}")
        return True


class Monivoi(Descarga):
    """ZLMXCOM_TRN_MONINVOI ('Monivoi'): invoices by CREATION date (+1 day, see utils), ONE query."""
    NOMBRE = "Monivoi_ZLMXCOM"
    MARCA = "Monivoi"
    TX = "ZLMXCOM_TRN_MONINVOI"
    VKORG = "MX21"                 # sales organization
    VTWEG = "01"                   # distribution channel
    RFC = "XAXX010101000"          # generic RFC ("público en general")

    def __init__(self, desde, hasta, carpeta, archivo, sociedad, vkorg, vtweg, rfc, session):
        super().__init__(carpeta, archivo or nombre_archivo_dia_monivoi(hasta), session)
        self.desde, self.hasta = desde, hasta
        self.sociedad, self.vkorg, self.vtweg, self.rfc = sociedad, vkorg, vtweg, rfc

    def pasos(self):
        for fecha in (self.desde, self.hasta):
            datetime.strptime(fecha, "%d.%m.%Y")
        os.makedirs(self.carpeta, exist_ok=True)
        s = self.conectar()
        s.findById("wnd[0]").maximize()
        s.findById("wnd[0]/tbar[0]/okcd").text = "/n" + self.TX
        s.findById("wnd[0]").sendVKey(0)

        s.findById("wnd[0]/usr/ctxtP_BUKRS").text = self.sociedad
        s.findById("wnd[0]/usr/ctxtSO_ERDAT-LOW").text = fecha_efectiva_monivoi_str(self.desde)
        s.findById("wnd[0]/usr/ctxtSO_ERDAT-HIGH").text = fecha_efectiva_monivoi_str(self.hasta)
        s.findById("wnd[0]/usr/ctxtSO_VKORG-LOW").text = self.vkorg
        s.findById("wnd[0]/usr/ctxtSO_VTWEG-LOW").text = self.vtweg
        campo = s.findById("wnd[0]/usr/txtSO_STCD1-LOW")
        campo.text = self.rfc
        campo.setFocus()
        campo.caretPosition = len(self.rfc)

        self.consultar()
        if SAP.sin_partidas(s, self.ruta):
            return False

        s.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        try:
            s.findById("wnd[1]/usr/cmbG_LISTBOX").setFocus()          # format selector: keep its default
        except Exception:
            pass
        s.findById("wnd[1]/tbar[0]/btn[0]").press()
        desde, _ = self.dialogo_guardar("wnd[1]/tbar[0]/btn[0]")
        Archivos.esperar_exportacion(self.ruta_xlsx, desde)
        Paralelo.marca(f"{self.MARCA}: archivo exportado")
        Excel.cerrar_libro(self.ruta_xlsx)
        if self.a_csv:
            Archivos.xlsx_a_csv(self.ruta_xlsx, self.ruta)
        print(f"[MONIVOI] Archivo descargado correctamente: {self.ruta}")
        self.volver_al_menu()
        return True

    def volver_al_menu(self):
        """Back → 'end the list?' → Yes → Back, as recorded. The file is safe: a problem is only reported."""
        s = self.session
        try:
            s.findById("wnd[0]/tbar[0]/btn[3]").press()
            if SAP.esperar(s, "wnd[1]/usr/btnSPOP-OPTION1"):        # a Back pressed over it does nothing
                s.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                SAP.esperar_libre(s)
            elif SAP.existe(s, "wnd[1]"):
                raise RuntimeError("SAP mostró un popup distinto al de confirmar el cierre de la lista "
                                   "y no lo toqué.")
            else:
                print("[MONIVOI] SAP no pidió confirmar el cierre de la lista; sigo al menú.")
            s.findById("wnd[0]/tbar[0]/btn[3]").press()
            if SAP.existe(s, "wnd[1]"):
                raise RuntimeError("Quedó un popup abierto al regresar al menú principal.")
            self.en_menu = True
        except Exception as e:
            print("[MONIVOI] ⚠️ El archivo ya quedó guardado, pero no pude regresar al "
                  f"menú principal de SAP: {e}")


class ZSE16:
    """
    ZSE16 popups. 'Fields for selection' is remembered PER USER, so the VBRL download sets
    it on every run; the list of layouts comes as an ALV grid or as a classic list.
    """
    MENU_CAMPOS = "wnd[0]/mbar/menu[3]/menu[2]"        # Settings > Fields for selection
    BTN_DESMARCAR = "wnd[1]/tbar[0]/btn[14]"           # Deselect all
    BTN_BUSCAR = "wnd[1]/tbar[0]/btn[71]"              # Find (of a list popup)
    COLUMNA_LAYOUT = "VARIANT"                         # technical name of the grid's 'Layout' column

    @staticmethod
    def patron(campo):
        """The WHOLE field name anywhere in a text: 'VBELN' matches 'Línea 5: VBELN', never 'VBELN_VF'."""
        nombre = re.escape(campo.strip().upper())
        return re.compile(rf"(?<![A-Z0-9_/]){nombre}(?![A-Z0-9_/])")

    @staticmethod
    def renglon_del_campo(session, campo):
        """
        Among the lines of 'Fields for selection' on screen, the one of `campo`: it has a
        CHECKBOX and the whole name in its texts (whether SAP writes it first, after the
        description or in the same label). Returns (row or None, rows); RuntimeError if two
        lines could be it (a wrong field would query other documents).
        """
        patron = ZSE16.patron(campo)
        renglones = SAP.renglones_de_lista(SAP.elementos_de_lista(session, "wnd[1]/usr"))
        candidatos = sorted(f for f, r in renglones.items()
                            if r["casillas"] and patron.search(r["texto"].upper()))
        if len(candidatos) > 1:
            raise RuntimeError(
                f"En 'Campos para selección' hay {len(candidatos)} renglones que podrían ser el "
                f"campo '{campo}' y no sé cuál marcar: "
                + "; ".join(repr(renglones[f]["texto"][:80]) for f in candidatos)
            )
        return (candidatos[0] if candidatos else None), renglones

    @staticmethod
    def buscar_campo(session, campo, intentos=3):
        """
        Bring `campo` into view with the recording's Find (type it → hit list → F2). The hits
        with the whole name are tried in order, those that START with it first; after every F2
        the field line must be on screen, so a wrong hit costs one more Find, never a wrong field.
        Returns its row; RuntimeError if it never shows up.
        """
        patron = ZSE16.patron(campo)
        al_inicio = re.compile(rf"[\W\d_]*{re.escape(campo.strip().upper())}(?![A-Z0-9_/])")
        vistos = "(ninguna)"
        for intento in range(intentos):
            session.findById(ZSE16.BTN_BUSCAR).press()
            if not SAP.esperar(session, "wnd[2]/usr/txtRSYSF-STRING", 10):
                raise RuntimeError("SAP no abrió el cuadro de búsqueda de 'Campos para selección'.")
            session.findById("wnd[2]/usr/txtRSYSF-STRING").text = campo
            session.findById("wnd[2]/usr/txtRSYSF-STRING").caretPosition = len(campo)
            session.findById("wnd[2]/tbar[0]/btn[0]").press()

            if SAP.esperar(session, "wnd[3]/usr", 5):
                renglones = SAP.renglones_de_lista(SAP.elementos_de_lista(session, "wnd[3]/usr"))
                vistos = SAP.resumen_renglones(renglones)
                con_nombre = [f for f in sorted(renglones) if patron.search(renglones[f]["texto"].upper())]
                orden = ([f for f in con_nombre if al_inicio.match(renglones[f]["texto"].upper())]
                         + [f for f in con_nombre if not al_inicio.match(renglones[f]["texto"].upper())])
                if intento >= len(orden):
                    SAP.cerrar_popups(session, encima_de=1)
                    break
                etiquetas = renglones[orden[intento]]["etiquetas"]
                etiqueta, texto = next(((comp, t) for _, comp, t in etiquetas if patron.search(t.upper())),
                                       (etiquetas[0][1], etiquetas[0][2]))
                etiqueta.setFocus()
                etiqueta.caretPosition = min(3, len(texto))
                session.findById("wnd[3]").sendVKey(2)                  # go to the hit
            elif SAP.existe(session, "wnd[2]"):                         # the Find found nothing
                aviso = SAP.aviso(session)
                SAP.cerrar_popups(session, encima_de=1)
                raise RuntimeError(f"SAP no encontró el campo '{campo}' en 'Campos para selección' "
                                   f"(mensaje: '{aviso or 'sin mensaje'}').")
            # (neither wnd[3] nor wnd[2]: SAP jumped straight to its only hit)

            if SAP.existe(session, "wnd[2]"):
                SAP.cerrar_popups(session, encima_de=1)                # that hit was not a line to go to
            if not SAP.esperar(session, "wnd[1]/usr", 5):
                raise RuntimeError("La ventana 'Campos para selección' se cerró al buscar el campo.")
            fila, _ = ZSE16.renglon_del_campo(session, campo)
            if fila is not None:
                return fila
        raise RuntimeError(f"No logré que 'Campos para selección' mostrara el renglón del campo "
                           f"'{campo}'. Líneas de la búsqueda: {vistos}")

    @staticmethod
    def campos_en_pantalla(session, maximo=10):
        """Labels of the selection fields I1, I2... on screen (best effort, for error messages)."""
        textos = []
        for n in range(1, maximo + 1):
            if not SAP.existe(session, f"wnd[0]/usr/btn%_I{n}_%_APP_%-VALU_PUSH"):
                break
            try:
                textos.append(str(session.findById(f"wnd[0]/usr/txt%_I{n}_%_APP_%-TEXT").Text).strip())
            except Exception:
                textos.append(f"I{n}")
        return textos

    @staticmethod
    def dejar_solo_campo(session, campo):
        """
        `campo` as the ONLY selection field, like VBRL.vbs: Fields for selection → Deselect all
        → tick its line → OK. The line is found by its text (the recording ticks a fixed
        chk[2,5]); the Find is only used when it is not on screen. Guards: no other visible
        field left ticked, the tick read back, and exactly ONE field (I1, no I2) at the end.
        """
        session.findById(ZSE16.MENU_CAMPOS).select()
        if not SAP.esperar(session, ZSE16.BTN_DESMARCAR, 10):
            raise RuntimeError("SAP no abrió la ventana 'Campos para selección' de ZSE16.")
        session.findById(ZSE16.BTN_DESMARCAR).press()

        fila, renglones = ZSE16.renglon_del_campo(session, campo)
        if fila is None:
            fila = ZSE16.buscar_campo(session, campo)
            _, renglones = ZSE16.renglon_del_campo(session, campo)

        casillas = renglones[fila]["casillas"]
        if len(casillas) != 1:
            raise RuntimeError(f"En 'Campos para selección' esperaba UNA casilla en el renglón de "
                               f"'{campo}' (renglón {fila}) y hay {len(casillas)}.")
        otras = [renglones[f]["texto"][:40] for f in sorted(renglones) if f != fila
                 for c in renglones[f]["casillas"] if bool(c.selected)]
        if otras:
            raise RuntimeError("El botón 'Desmarcar todo' de 'Campos para selección' dejó campos "
                               "marcados: " + "; ".join(repr(t) for t in otras))
        casilla = casillas[0]
        casilla.selected = True
        casilla.setFocus()
        if not bool(casilla.selected):
            raise RuntimeError(f"No pude marcar el campo '{campo}' en 'Campos para selección'.")
        session.findById("wnd[1]/tbar[0]/btn[0]").press()                 # OK
        SAP.esperar_libre(session)

        if not SAP.esperar_cierre(session):
            raise RuntimeError("La ventana 'Campos para selección' sigue abierta después de OK.")
        if not SAP.existe(session, "wnd[0]/usr/btn%_I1_%_APP_%-VALU_PUSH"):
            raise RuntimeError(f"ZSE16 no mostró el campo de selección '{campo}' (I1).")
        if SAP.existe(session, "wnd[0]/usr/btn%_I2_%_APP_%-VALU_PUSH"):
            campos = ", ".join(ZSE16.campos_en_pantalla(session))
            raise RuntimeError(
                f"ZSE16 quedó con más de un campo de selección ({campos}); "
                f"esperaba sólo '{campo}'. Revisa 'Opciones > Campos para selección'."
            )

    @staticmethod
    def elegir_layout_en_grid(session, id_grid, nombre):
        """
        The layout `nombre` in the ALV grid of layouts, BY NAME (the recording clicks row 0,
        which is whatever layout sorts first): exactly one row must be called `nombre`, and the
        popup must close after the click. Returns the name it saw.
        """
        grid = session.findById(id_grid)
        try:
            nombres = [str(grid.GetCellValue(i, ZSE16.COLUMNA_LAYOUT)).strip()
                       for i in range(int(grid.RowCount))]
        except Exception as e:
            try:
                columnas = [str(c) for c in grid.ColumnOrder]
            except Exception:
                columnas = ["(no pude leerlas)"]
            raise RuntimeError(f"La lista de layouts no tiene la columna '{ZSE16.COLUMNA_LAYOUT}' "
                               f"(columnas: {', '.join(columnas)}): {e}") from e
        elegidas = [i for i, n in enumerate(nombres) if n.upper() == nombre.strip().upper()]
        if len(elegidas) != 1:
            vistas = ", ".join(repr(n) for n in nombres[:15]) or "(ninguno)"
            raise RuntimeError(f"Esperaba exactamente un layout '{nombre}' en la lista de layouts y hay "
                               f"{len(elegidas)}. Layouts que vi: {vistas}")
        fila = elegidas[0]
        grid.currentCellRow = fila
        grid.selectedRows = str(fila)
        grid.clickCurrentCell()
        SAP.esperar_libre(session)
        if not SAP.esperar_cierre(session):
            raise RuntimeError(f"El popup de layouts sigue abierto después de elegir '{nombre}'.")
        return nombres[fila]

    @staticmethod
    def elegir_layout_en_lista(session, nombre):
        """Earlier recording: Find → type the name → pick the hit line of the layout → OK. Its text."""
        session.findById(ZSE16.BTN_BUSCAR).press()
        if not SAP.esperar(session, "wnd[2]/usr/txtRSYSF-STRING", 10):
            raise RuntimeError("SAP no abrió el cuadro de búsqueda de layouts.")
        session.findById("wnd[2]/usr/txtRSYSF-STRING").text = nombre
        session.findById("wnd[2]/usr/txtRSYSF-STRING").caretPosition = len(nombre)
        session.findById("wnd[2]/tbar[0]/btn[0]").press()
        if not SAP.esperar(session, "wnd[3]/usr", 10):
            raise RuntimeError(f"SAP no mostró resultados al buscar el layout '{nombre}' "
                               f"(mensaje: '{SAP.aviso(session) or 'sin mensaje'}').")
        hits = [(h, t) for h, tipo, t, _ in SAP.elementos_de_lista(session, "wnd[3]/usr")
                if tipo == "GuiLabel" and t]
        hit, texto = hits[ZSE16.indice_coincidencia([t for _, t in hits], nombre)]
        print(f"[VBRL] Layout elegido por búsqueda de nombre: {texto!r}")
        hit.setFocus()
        hit.caretPosition = min(2, len(texto))
        session.findById("wnd[3]").sendVKey(2)                       # go to the hit
        session.findById("wnd[1]/tbar[0]/btn[0]").press()            # choose it
        SAP.esperar_libre(session)
        if not SAP.esperar_cierre(session):
            raise RuntimeError(f"El popup de layouts sigue abierto después de elegir '{texto}'.")
        return texto

    @staticmethod
    def indice_coincidencia(textos, nombre, que="el layout"):
        """
        Which hit line is the item `nombre`: it must contain the WHOLE name ('VALS2', '/VALS'
        are other items) and START with it (not a description or a header that mentions it),
        and exactly one line may qualify. RuntimeError listing the lines otherwise.
        """
        objetivo = re.escape(nombre.strip().upper())
        palabra = re.compile(rf"(?<![A-Z0-9_/]){objetivo}(?![A-Z0-9_])")
        al_inicio = re.compile(rf"[\W\d_]*{objetivo}(?![A-Z0-9_])")
        candidatos = [i for i, t in enumerate(textos)
                      if palabra.search(t.upper()) and al_inicio.match(t.upper())]
        if len(candidatos) == 1:
            return candidatos[0]
        if not candidatos:
            vistas = "; ".join(repr(t[:80]) for t in textos[:15]) or "(ninguna)"
            raise RuntimeError(f"La búsqueda en SAP no devolvió ninguna línea que empiece con {que} "
                               f"'{nombre}'. Líneas que vi: {vistas}")
        raise RuntimeError(f"Hay {len(candidatos)} líneas que podrían ser {que} '{nombre}' y no sé cuál "
                           f"elegir: {'; '.join(repr(textos[i][:80]) for i in candidatos)}")


class DescargaConLista(Descarga):
    """
    VBRL and FB03: ONE query for a list of documents pasted through the clipboard, saved as
    .csv. The conversion goes to a '.part' that replaces the final file only once checked;
    if something fails after SAP exported, the CSV is still kept as *_REVISAR.csv (or as the
    .part if even that rename fails; the consolidation reads neither) and
    DescargaPorRevisarError says why.
    """
    ETIQUETA = ""                  # 'VBRL' / 'FB03'
    UNIDAD = ("documento(s)", "pedido(s)")
    MENSAJES_SIN_DATOS = ()        # status-bar texts (lower case) meaning "no entries"

    def __init__(self, documentos, carpeta, archivo, session):
        super().__init__(carpeta, archivo, session)
        self.documentos = documentos
        self.lista = []
        self.ruta_parcial = self.ruta + ".part"
        self.ruta_revisar = os.path.join(self.carpeta, self.base + "_REVISAR.csv")
        self.exportado = False     # SAP finished writing the .xlsx in time
        self.con_carpeta = True    # the save dialog had the folder field
        self.conservar_parcial = False   # the .part is the only copy left: limpiar() keeps it

    def preparar(self):
        """Checks before touching SAP."""
        if os.path.basename(self.archivo) != self.archivo:
            raise ValueError(f"El nombre del archivo no debe incluir carpetas: {self.archivo!r}")
        if self.ext != ".csv":
            raise ValueError(f"{self.ETIQUETA} se guarda como .csv (se recibió '{self.archivo}').")
        self.lista = Archivos.normalizar_documentos(self.documentos, self.ETIQUETA)

    def borrar_restos(self):
        """Files of an earlier run must not pass for this one's (an old .xlsx makes SAP ask to overwrite)."""
        for resto in (self.ruta_xlsx, self.ruta_parcial, self.ruta_revisar):
            if os.path.exists(resto):
                try:
                    os.remove(resto)
                except OSError as e:
                    raise RuntimeError(
                        f"No pude borrar un temporal anterior (¿está abierto en Excel?): {resto}"
                    ) from e

    def pegar_lista(self, boton):
        """Paste the list into the multiple selection that `boton` opens, through the clipboard."""
        s = self.session
        s.findById(boton).press()
        with Portapapeles.exclusivo():
            Portapapeles.copiar_y_verificar("\r\n".join(self.lista))
            s.findById("wnd[1]/tbar[0]/btn[24]").press()            # import from clipboard
            time.sleep(1 + len(self.lista) / 5000)                   # let SAP load the list
        s.findById("wnd[1]/tbar[0]/btn[8]").press()                 # copy and close (no Enter before it)
        SAP.esperar_libre(s)

    def sin_resultados(self, id_seleccion, que):
        """
        After Execute: True if SAP said there are no entries (placeholder written). Still on
        the selection screen with any other message: RuntimeError with SAP's text (no guessing).
        """
        if SAP.sin_partidas(self.session, self.ruta, self.MENSAJES_SIN_DATOS):
            return True
        if SAP.existe(self.session, id_seleccion):
            raise RuntimeError(
                f"{self.ETIQUETA} no mostró {que}. SAP dice: '{SAP.aviso(self.session) or '(sin mensaje)'}'. "
                f"Si significa que no hay datos, agrega ese texto a {type(self).__name__}.MENSAJES_SIN_DATOS."
            )
        return False

    def exportar(self, menu):
        """List > Export > Spreadsheet → default format → save → wait for the file."""
        s = self.session
        s.findById(menu).select()
        s.findById("wnd[1]/tbar[0]/btn[0]").press()                 # default format
        desde, self.con_carpeta = self.dialogo_guardar("wnd[1]/tbar[0]/btn[0]")
        self.exportado = Archivos.esperar_nuevo(self.ruta_xlsx, desde)
        if self.exportado:
            Paralelo.marca(f"{self.MARCA}: archivo exportado")

    def volver_al_menu(self):
        """3 × Back, as recorded. Leaving the list makes SAP release the .xlsx; then close any Excel left."""
        try:
            for _ in range(3):
                self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self.en_menu = True
        except Exception as e:
            print(f"[{self.ETIQUETA}] ⚠️ No pude regresar al menú principal de SAP "
                  f"(el archivo ya se exportó): {e}")
        time.sleep(2)
        Excel.cerrar_libro(self.ruta_xlsx)

    def guardar(self, revisar=None):
        """Convert, check and only then replace the final .csv (see the class docstring)."""
        if not os.path.exists(self.ruta_xlsx):
            sin_carpeta = ("" if self.con_carpeta else " El diálogo de guardar no tenía el campo de carpeta: "
                                                       "SAP lo guardó en su carpeta por defecto.")
            raise DescargaPorRevisarError(
                f"SAP no dejó el archivo {self.nombre_xlsx} en {self.carpeta}.{sin_carpeta}"
            )
        try:
            Archivos.xlsx_a_csv(self.ruta_xlsx, self.ruta_parcial)
        except ExportacionTruncadaError:
            raise
        except Exception as e:
            raise DescargaPorRevisarError(
                f"No pude convertir {self.nombre_xlsx} a CSV (lo dejé en la carpeta para revisarlo): {e}"
            ) from e
        try:
            if not self.exportado:
                raise RuntimeError("SAP no terminó de escribir el archivo a tiempo: puede estar incompleto.")
            Archivos.validar_csv(self.ruta_parcial, self.lista, self.ETIQUETA, self.UNIDAD, revisar)
        except Exception as e:
            guardado = self.ruta_revisar
            try:
                os.replace(self.ruta_parcial, guardado)
            except OSError as e2:
                self.conservar_parcial, guardado = True, self.ruta_parcial
                print(f"[{self.ETIQUETA}] ⚠️ No pude renombrarlo a "
                      f"{os.path.basename(self.ruta_revisar)}: {e2}")
            print(f"[{self.ETIQUETA}] ⚠️ Guardado para revisar: {guardado}")
            raise DescargaPorRevisarError(f"guardé {os.path.basename(guardado)} para revisarlo: {e}") from e
        try:
            os.replace(self.ruta_parcial, self.ruta)
        except OSError as e:
            raise RuntimeError(f"No pude guardar {self.ruta} (¿está abierto en Excel?). "
                               "Ciérralo y vuelve a descargar.") from e
        print(f"[{self.ETIQUETA}] Archivo descargado correctamente: {self.ruta}")
        return True

    def limpiar(self):
        if os.path.exists(self.ruta_parcial) and not self.conservar_parcial:
            try:
                os.remove(self.ruta_parcial)
            except OSError:
                pass
        if self.session is not None and not self.en_menu:
            SAP.cerrar_popups(self.session)
            super().limpiar()


class VBRL(DescargaConLista):
    """ZSE16 on table VBRL for the Monivoi invoices ('Factura' column), as VBRL.vbs (17-oct-2026) does."""
    NOMBRE = "VBRL_ZSE16"
    MARCA = "ZSE16/VBRL"
    ETIQUETA = "VBRL"
    ARCHIVO = "VBRL.csv"
    UNIDAD = ("factura(s)", "pedida(s)")
    CAMPO = "VBELN"                # the ONLY selection field: the invoice of the Monivoi 'Factura' column
    LAYOUT = "VALS"                # always picked BY NAME
    GRID_LAYOUTS = "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell"
    MENU_EXPORTAR_GRID = "wnd[0]/mbar/menu[0]/menu[10]/menu[3]/menu[1]"     # 17-oct-2026 recording
    MENU_EXPORTAR_LISTA = "wnd[0]/mbar/menu[0]/menu[10]/menu[1]/menu[1]"    # earlier recording
    # The consolidation reads VBRL BY POSITION (D = invoice, E = amount, like the stakeholder's
    # XLOOKUP): keep these in sync with Consolidacion.EtapaVBRL.COL_FACTURA / COL_IMPORTE.
    MIN_COLUMNAS = 5
    COL_FACTURA = 3                # 0-based: column D
    # Share of column D that must look like an invoice as SAP exports it: 'C' + digits ('C270043678'),
    # or the bare digits of a billing document (~0.2% of the rows).
    MIN_NUMERICA = 0.9
    PATRON_FACTURA = r"[Cc]?[0-9]+(\.0+)?"
    MENSAJES_SIN_DATOS = (
        "no table entries",
        "ninguna entrada de tabla",
        "no existen entradas de tabla",
        "no se encontraron entradas de tabla",
        "no se han encontrado entradas de tabla",
    )

    def __init__(self, documentos, carpeta, archivo, session):
        super().__init__(documentos, carpeta, archivo or self.ARCHIVO, session)

    def pasos(self):
        self.preparar()
        os.makedirs(self.carpeta, exist_ok=True)
        self.borrar_restos()
        s = self.conectar()
        SAP.cerrar_popups(s)                     # a popup left by an earlier run blocks the command field
        s.findById("wnd[0]").maximize()
        s.findById("wnd[0]/tbar[0]/okcd").text = "/nZSE16"
        s.findById("wnd[0]").sendVKey(0)
        s.findById("wnd[0]/usr/ctxtGV_TABNAME").text = "VBRL"
        s.findById("wnd[0]/usr/ctxtGV_TABNAME").caretPosition = 4
        s.findById("wnd[0]").sendVKey(0)
        ZSE16.dejar_solo_campo(s, self.CAMPO)

        print(f"[VBRL] {len(self.lista):,} factura(s) única(s) a consultar.")
        self.pegar_lista("wnd[0]/usr/btn%_I1_%_APP_%-VALU_PUSH")
        s.findById("wnd[0]/usr/txtMAX_SEL").text = ""
        s.findById("wnd[0]/usr/txtMAX_SEL").setFocus()
        valor_max = str(s.findById("wnd[0]/usr/txtMAX_SEL").text).strip()
        if valor_max.strip("0"):                 # '' / '0' = no limit; anything else cuts the result silently
            raise RuntimeError(f"No pude quitar el límite de resultados de ZSE16 (MAX_SEL = '{valor_max}'): "
                               "SAP devolvería sólo una parte de las facturas.")

        self.consultar()
        if self.sin_resultados("wnd[0]/usr/txtMAX_SEL", "la lista de resultados"):
            return False
        self.exportar(self.elegir_layout())
        self.volver_al_menu()
        return self.guardar(self.revisar_columnas)

    def elegir_layout(self):
        """Layout VALS by name, in the popup SAP opened (grid or list). Returns that popup's export menu."""
        s = self.session
        s.findById("wnd[0]/tbar[1]/btn[33]").press()
        modo, limite = None, time.time() + 10
        while modo is None and time.time() < limite:
            if SAP.existe(s, self.GRID_LAYOUTS):
                modo = "grid"
            elif SAP.existe(s, ZSE16.BTN_BUSCAR):
                modo = "lista"
            else:
                time.sleep(0.25)
        if modo is None:
            raise RuntimeError("SAP no abrió la lista de layouts de ZSE16.")
        if modo == "lista":
            ZSE16.elegir_layout_en_lista(s, self.LAYOUT)
            return self.MENU_EXPORTAR_LISTA
        nombre = ZSE16.elegir_layout_en_grid(s, self.GRID_LAYOUTS, self.LAYOUT)
        print(f"[VBRL] Layout elegido por nombre en la lista de layouts: {nombre!r}")
        return self.MENU_EXPORTAR_GRID

    @classmethod
    def revisar_columnas(cls, df):
        """Column D must be the invoice: someone could edit the layout's columns in SAP."""
        columnas = ", ".join(f"{chr(65 + i)}='{c}'" for i, c in enumerate(df.columns[:8]))
        if df.shape[1] < cls.MIN_COLUMNAS:
            raise RuntimeError(f"el archivo de VBRL trae {df.shape[1]} columna(s) y se esperan al menos "
                               f"{cls.MIN_COLUMNAS} (D = factura, E = importe). Columnas: {columnas}")
        col = df.iloc[:, cls.COL_FACTURA].astype(str).str.strip()
        col = col[col != ""]
        if col.empty:
            raise RuntimeError(f"la columna D del archivo de VBRL está vacía. Columnas: {columnas}")
        fraccion = float(col.str.fullmatch(cls.PATRON_FACTURA).mean())
        if fraccion < cls.MIN_NUMERICA:
            raise RuntimeError(f"la columna D del archivo de VBRL no parece de facturas (sólo {fraccion:.0%} "
                               f"de sus valores son 'C' + dígitos o un número; primero: {col.iloc[0]!r}). "
                               f"Columnas: {columnas}")


class FB03(DescargaConLista):
    """
    FB03 document list for the Gallo documents (Clase C2/C6, Referencia Global; see
    Consolidacion.leer_documentos_gallo_fb03), as FB03.vbs does. The fiscal year goes from
    the year of DateFrom to the year of DateTo.
    """
    NOMBRE = "FB03_Documentos"
    MARCA = "FB03"
    ETIQUETA = "FB03"
    ARCHIVO = "FB03.csv"
    MENSAJES_SIN_DATOS = (
        "no documents",
        "no document was selected",
        "ningún documento",
        "no se encontraron documentos",
        "no se han seleccionado documentos",
        "no hay documentos",
        "no existen documentos",
    )

    def __init__(self, documentos, desde, hasta, carpeta, archivo, sociedad, session):
        super().__init__(documentos, carpeta, archivo or self.ARCHIVO, session)
        self.desde, self.hasta, self.sociedad = desde, hasta, sociedad

    def pasos(self):
        self.preparar()
        sociedad = "" if self.sociedad is None else str(self.sociedad).strip().upper()
        if not re.fullmatch(r"[A-Z0-9]{1,4}", sociedad):
            raise ValueError("La sociedad debe ser un código de 1 a 4 letras o dígitos "
                             f"(se recibió {sociedad!r}).")
        anio_desde = datetime.strptime(self.desde, "%d.%m.%Y").year
        anio_hasta = datetime.strptime(self.hasta, "%d.%m.%Y").year
        if anio_hasta < anio_desde:
            raise ValueError(f"El periodo es inválido: {self.desde} — {self.hasta}.")
        os.makedirs(self.carpeta, exist_ok=True)
        self.borrar_restos()

        s = self.conectar()
        SAP.cerrar_popups(s)
        s.findById("wnd[0]").maximize()
        s.findById("wnd[0]/tbar[0]/okcd").text = "/nFB03"
        s.findById("wnd[0]").sendVKey(0)
        s.findById("wnd[0]/tbar[1]/btn[20]").press()                     # document list
        s.findById("wnd[0]/usr/ctxtBR_BUKRS-LOW").text = sociedad
        s.findById("wnd[0]/usr/txtBR_BELNR-LOW").text = "0"              # placeholder of the recording
        s.findById("wnd[0]/usr/txtBR_GJAHR-LOW").text = str(anio_desde)
        if anio_hasta != anio_desde:
            s.findById("wnd[0]/usr/txtBR_GJAHR-HIGH").text = str(anio_hasta)
        s.findById("wnd[0]/usr/txtBR_GJAHR-LOW").setFocus()
        s.findById("wnd[0]/usr/txtBR_GJAHR-LOW").caretPosition = 4

        anios = str(anio_desde) if anio_hasta == anio_desde else f"{anio_desde}-{anio_hasta}"
        print(f"[FB03] {len(self.lista):,} documento(s) únicos a consultar (ejercicio {anios}).")
        self.pegar_lista("wnd[0]/usr/btn%_BR_BELNR_%_APP_%-VALU_PUSH")
        self.consultar()
        if self.sin_resultados("wnd[0]/usr/btn%_BR_BELNR_%_APP_%-VALU_PUSH", "la lista de documentos"):
            return False
        self.cambiar_layout()
        self.exportar("wnd[0]/mbar/menu[0]/menu[1]/menu[1]")
        self.volver_al_menu()
        return self.guardar()

    def cambiar_layout(self):
        """Open and close the layouts popup, then 'Change layout': row 2 + 'single column' twice."""
        s = self.session
        s.findById("wnd[0]/tbar[1]/btn[33]").press()
        if SAP.esperar(s, "wnd[1]", 5):
            s.findById("wnd[1]").close()
        else:
            print("[FB03] El botón de layouts no abrió popup; sigo con 'Cambiar layout'.")
        s.findById("wnd[0]/tbar[1]/btn[32]").press()
        alv = "wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620"
        if not SAP.esperar(s, alv + "/btnAPP_FL_SING", 10):
            raise RuntimeError("SAP no abrió la ventana de 'Cambiar layout' de FB03.")
        columnas = alv + "/cntlCONTAINER2_LAYO/shellcont/shell"
        s.findById(columnas).currentCellRow = 2
        s.findById(columnas).selectedRows = "2"
        s.findById(alv + "/btnAPP_FL_SING").press()
        s.findById(columnas).selectedRows = "2"
        s.findById(alv + "/btnAPP_FL_SING").press()
        s.findById("wnd[1]/tbar[0]/btn[0]").press()
        SAP.esperar_libre(s)
        if not SAP.esperar_cierre(s):
            raise RuntimeError("La ventana de 'Cambiar layout' de FB03 sigue abierta.")


# ======================================================================
# Entry points (controller.py). True = downloaded with data; False = no data or a
# handled error (printed in the console). `session`: the SAP window of a parallel run.
# ======================================================================
def Gallo_FAGLL03(DateFrom, DateTo, FolderPath=None, FileName=None, sociedad=_SOCIEDAD,
                  cuentas=None, variante=Gallo.VARIANTE, session=None, clases=None):
    """FAGLL03 for 'DD.MM.YYYY' dates → FileName (default: nombre_archivo_dia(DateTo))."""
    return Gallo(DateFrom, DateTo, FolderPath, FileName, sociedad, cuentas, variante, session, clases).ejecutar()


def VentaEIAP_FAGLL03(DateFrom, DateTo, FolderPath=None, FileName=None, sociedad=_SOCIEDAD, session=None):
    """
    Gallo with the EIAP accounts, queried up to DateTo MINUS one day (utils.fecha_efectiva_eiap);
    the file name keeps the typed dates. ValueError if that leaves an empty or backwards range.
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
        session=session,
        clases=CLASES_EIAP,
    )


def Monivoi_ZLMXCOM(DateFrom, DateTo, FolderPath=None, FileName=None, sociedad=_SOCIEDAD,
                    vkorg=Monivoi.VKORG, vtweg=Monivoi.VTWEG, rfc=Monivoi.RFC, session=None):
    """ZLMXCOM_TRN_MONINVOI for 'DD.MM.YYYY' dates → FileName (default nombre_archivo_dia_monivoi)."""
    return Monivoi(DateFrom, DateTo, FolderPath, FileName, sociedad, vkorg, vtweg, rfc, session).ejecutar()


def VBRL_ZSE16(documentos, FolderPath=None, FileName=None, session=None):
    """Table VBRL (ZSE16) for the invoices in `documentos` → FileName (.csv, default 'VBRL.csv')."""
    return VBRL(documentos, FolderPath, FileName, session).ejecutar()


def FB03_Documentos(documentos, DateFrom, DateTo, FolderPath=None, FileName=None,
                    sociedad=_SOCIEDAD, session=None):
    """FB03 document list for `documentos` (fiscal years of the period) → FileName (.csv)."""
    return FB03(documentos, DateFrom, DateTo, FolderPath, FileName, sociedad, session).ejecutar()


# ======================================================================
# Quick test without the GUI:  python Descargas_SAP.py
# ======================================================================
if __name__ == "__main__":
    print("=== Prueba de descarga (standalone) ===")
    print("1) Gallo (FAGLL03)")
    print("2) Monivoi (ZLMXCOM_TRN_MONINVOI)")
    print("3) VBRL (ZSE16) — con las facturas (columna 'Factura') del archivo de Monivoi del mismo periodo")
    print("4) FB03 (lista de documentos) — con los documentos (columna 'Nº documento') del archivo de Gallo "
          "del mismo periodo, sólo Clase de documento C2/C6 y Referencia Global")
    print("5) Venta EIAP (FAGLL03, cuentas EIAP) — la fecha hasta se consulta menos 1 día")
    opcion = input("Elige transacción [1/2/3/4/5]: ").strip()
    date_from = input("Fecha desde (DD.MM.YYYY): ").strip()
    date_to = input("Fecha hasta (DD.MM.YYYY): ").strip()
    carpeta = Archivos.ruta_input_default()

    if opcion in ("3", "4"):
        # The dates only NAME the files (one name for a day, another for a range) and find
        # the file the documents are read from (Monivoi for VBRL, Gallo for FB03).
        from Consolidacion import leer_documentos_monivoi, leer_documentos_gallo_fb03
        from utils import (
            nombre_archivo_rango,
            nombre_archivo_rango_monivoi,
            nombre_archivo_dia_vbrl, nombre_archivo_rango_vbrl,
            nombre_archivo_dia_fb03, nombre_archivo_rango_fb03,
        )
        un_dia = date_from == date_to
        if opcion == "3":
            archivo_origen = (nombre_archivo_dia_monivoi(date_to) if un_dia
                              else nombre_archivo_rango_monivoi(date_from, date_to))
            archivo = (nombre_archivo_dia_vbrl(date_to) if un_dia
                       else nombre_archivo_rango_vbrl(date_from, date_to))
        else:
            archivo_origen = (nombre_archivo_dia(date_to) if un_dia
                              else nombre_archivo_rango(date_from, date_to))
            archivo = (nombre_archivo_dia_fb03(date_to) if un_dia
                       else nombre_archivo_rango_fb03(date_from, date_to))
        ruta_origen = os.path.join(carpeta, archivo_origen)
        print(f"\nDocumentos leídos de   : {ruta_origen}")
        if opcion == "3":
            documentos, _filas, _col = leer_documentos_monivoi(ruta_origen, "factura")
        else:
            documentos, _filas, _col = leer_documentos_gallo_fb03(ruta_origen)
        print(f"Se leyó la columna     : {_col}")
        print(f"El archivo se nombrará : {archivo}")
        print(f"Se guardará en         : {carpeta}\n")
        try:
            if opcion == "3":
                ok = VBRL_ZSE16(documentos, FileName=archivo)
            else:
                ok = FB03_Documentos(documentos, date_from, date_to, FileName=archivo)
        except DescargaPorRevisarError as e:
            print(f"⚠️ {e}")
            ok = False
    elif opcion == "5":
        try:
            archivo = nombre_archivo_rango_eiap(date_from, date_to)
            print(f"\nEl archivo se nombrará : {archivo}")
            print(f"Se consultará en SAP   : {date_from} — "
                  f"{fecha_efectiva_eiap(date_to).strftime(FORMATO_FECHA)}")
            print(f"Se guardará en         : {carpeta}\n")
            ok = VentaEIAP_FAGLL03(date_from, date_to)
        except ValueError as e:
            print(f"❌ {e}")
            ok = False
    elif opcion == "2":
        print(f"\nEl archivo se nombrará : {nombre_archivo_dia_monivoi(date_to)}")
        print(f"Se guardará en         : {carpeta}\n")
        ok = Monivoi_ZLMXCOM(date_from, date_to)
    else:
        print(f"\nEl archivo se nombrará : {nombre_archivo_dia(date_to)}")
        print(f"Se guardará en         : {carpeta}\n")
        ok = Gallo_FAGLL03(date_from, date_to)

    print("✅ Descarga con datos." if ok else "⚠️ Sin datos o error (revisa la consola).")
