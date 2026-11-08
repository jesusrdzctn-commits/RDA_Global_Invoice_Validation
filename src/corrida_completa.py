"""Validación Factura Global — command line: the full run (SAP download + consolidation) with NO window.

    ValidacionFacturaGlobal.exe --corrida-completa --fecha 03.11.2026

It is the 'Corrida completa' button without the GUI: the same engines (controller.correr_corrida_completa),
fed by `GuiSinVentana`, a stand-in that answers what the controller asks of the window from the command
line and from config.json, and that reports to a log instead of to widgets.

Why this module is careful about its surroundings (the .exe is built with PyInstaller --windowed):
  * a windowed process has no console: sys.stdout / sys.stderr are None, so every print() of the engines
    would vanish. `conectar_consola` attaches to the console of the parent (cmd, PowerShell) when there is
    one, and EVERYTHING that is printed is also written to a log file (`_Registro`): a run started by the
    Task Scheduler, which has no console at all, still leaves its story in Logs\\.
  * only one full run at a time (`adquirir_unica_corrida`): two of them would fight over the SAP windows,
    the clipboard (the lists travel through it) and the very same files.
  * the exit code says how it went (SALIDA_*), for whoever schedules it.

Importing this module is cheap and needs neither SAP nor pandas: the controller (and with it everything
heavy) is imported only when a run really starts.
"""

import argparse
import io
import os
import re
import sys
import tempfile
import threading
import time
import traceback
from datetime import datetime

from utils import (
    FORMATO_FECHA, SOCIEDAD_DEFECTO, CORRIDAS_DESCARGA, parse_fecha, nombre_archivo_transaccion,
    carpeta_base_proyecto, carpetas_configuradas,
)

# --- Exit codes ---------------------------------------------------------------------------------
SALIDA_COMPLETA = 0       # downloaded everything, consolidated, no warnings
SALIDA_AVISOS = 1         # the workbook exists but something deserves a look (see the log)
SALIDA_USO = 2            # wrong command line: nothing was run
SALIDA_OCUPADO = 3        # another full run is already in progress: nothing was run
SALIDA_FALLO = 4          # failed or incomplete: there is NO new workbook
SALIDA_INTERRUMPIDA = 130  # Ctrl+C

_SALIDA_POR_ESTADO = {"completa": SALIDA_COMPLETA, "avisos": SALIDA_AVISOS,
                      "incompleta": SALIDA_FALLO, "error": SALIDA_FALLO}

NOMBRE_MUTEX = "Local\\ValidacionFacturaGlobal.CorridaCompleta"
ANIO_MINIMO = 2000          # an earlier year is a typo, not a period
_RE_SOCIEDAD = re.compile(r"[A-Z0-9]{4}")     # an SAP company code: four letters / digits
_RE_CRONOMETRO = re.compile(r"\s*\(\d+(?::\d{1,2}){1,2}\)\s*$")     # the ' (1:05)' clock the parallel runs append


class ErrorDeUso(Exception):
    """The command line is wrong (exit code SALIDA_USO). The message is for the user."""


# ==================================================================================================
# Arguments
# ==================================================================================================
def texto_ayuda():
    """The help text (Spanish, written by hand: argparse would print its own in English)."""
    n_transacciones = len(CORRIDAS_DESCARGA[1]) + len(CORRIDAS_DESCARGA[2])
    exe = "ValidacionFacturaGlobal.exe"
    return (
        "Validación Factura Global — línea de comandos\n"
        "\n"
        "Uso:\n"
        f"  {exe} --corrida-completa --fecha 03.11.2026\n"
        f"  {exe} --corrida-completa --desde 01.11.2026 --hasta 03.11.2026\n"
        "\n"
        "Hace lo mismo que el botón \"Corrida completa\" de la interfaz, pero sin abrirla: descarga las\n"
        f"{n_transacciones} transacciones de SAP y, al terminar, consolida en Excel los archivos de ESTA corrida\n"
        "(nunca los de una anterior). No pide confirmación: SAP GUI debe estar abierto y con sesión iniciada.\n"
        "\n"
        "Opciones:\n"
        "  --corrida-completa    La acción a ejecutar (obligatoria).\n"
        "  --fecha DD.MM.AAAA    Un solo día.\n"
        "  --desde DD.MM.AAAA    Inicio de un periodo (va con --hasta; no se combina con --fecha).\n"
        "  --hasta DD.MM.AAAA    Fin del periodo.\n"
        f"  --sociedad XXXX       Sociedad de SAP (por defecto {SOCIEDAD_DEFECTO}).\n"
        "  --paralelo            Descargas en paralelo, por pasadas, como el switch de la interfaz\n"
        "                        (por defecto una tras otra). Usa varias ventanas de SAP.\n"
        "  -h, --help            Muestra esta ayuda.\n"
        "\n"
        "Carpetas: las que se usaron por última vez en la interfaz (config.json, junto al .exe) o, si no hay,\n"
        "Documents\\Validacion Factura Global\\src\\Input y \\Output. La bitácora de cada corrida queda en\n"
        "...\\src\\Logs\\corrida_completa_AAAAMMDD_HHMMSS.log\n"
        "\n"
        "Código de salida:\n"
        f"  {SALIDA_COMPLETA}    terminó bien\n"
        f"  {SALIDA_AVISOS}    terminó con avisos: el Excel se generó, pero hay que revisar la bitácora\n"
        f"  {SALIDA_USO}    argumentos incorrectos (no se ejecutó nada)\n"
        f"  {SALIDA_OCUPADO}    ya hay otra corrida completa en curso (no se ejecutó nada)\n"
        f"  {SALIDA_FALLO}    falló o quedó incompleta: NO hay Excel nuevo\n"
        f"  {SALIDA_INTERRUMPIDA}  interrumpida con Ctrl+C\n"
        "\n"
        "El .exe es de ventana: cmd no espera a que termine y por eso no verás su código de salida con\n"
        "%ERRORLEVEL% a menos que lo lances así (en el Programador de tareas, con \"Iniciar\" y esperar):\n"
        f"  start /wait \"\" {exe} --corrida-completa --fecha 03.11.2026\n"
        "  echo %ERRORLEVEL%\n"
        "En PowerShell:  (Start-Process -Wait -PassThru .\\" + exe + " -ArgumentList '--corrida-completa','--fecha','03.11.2026').ExitCode\n"
    )


class _Parser(argparse.ArgumentParser):
    """argparse that never exits or prints by itself (a windowed .exe has nowhere to print): its errors are
    raised as ErrorDeUso and the help is `texto_ayuda`."""

    def error(self, message):
        raise ErrorDeUso(message)


def crear_parser():
    """The parser. No abbreviations (`--fec` is an error, not a guess) and no @file arguments."""
    parser = _Parser(prog="ValidacionFacturaGlobal.exe", add_help=False, allow_abbrev=False,
                     fromfile_prefix_chars=None)
    parser.add_argument("-h", "--help", action="store_true", dest="ayuda")
    parser.add_argument("--corrida-completa", action="store_true", dest="corrida_completa")
    parser.add_argument("--paralelo", action="store_true")
    # 'append' so that a repeated option is DETECTED (argparse would silently keep the last one).
    for opcion in ("--fecha", "--desde", "--hasta", "--sociedad"):
        parser.add_argument(opcion, action="append", dest=opcion[2:], metavar="VALOR")
    return parser


def parsear(argv):
    """argv (without the program name) → Namespace. ErrorDeUso for anything argparse rejects."""
    return crear_parser().parse_args(list(argv))


def _unico(valores, opcion):
    """The only value of an option that can come once: None if it was not given. ErrorDeUso if repeated."""
    if not valores:
        return None
    if len(valores) > 1:
        raise ErrorDeUso(f"{opcion} se indicó más de una vez.")
    return valores[0]


def _fecha_valida(texto, opcion, hoy):
    """'3.11.2026' / '03.11.2026' → '03.11.2026' (the form SAP and the file names use). ErrorDeUso if it is
    not a real date DD.MM.AAAA, is before ANIO_MINIMO or is after today (a typo; an unattended run must not
    spend half an hour on it)."""
    try:
        fecha = parse_fecha(texto)
    except (ValueError, AttributeError):
        raise ErrorDeUso(f"{opcion}: '{texto}' no es una fecha válida. Usa el formato DD.MM.AAAA "
                         "(por ejemplo 03.11.2026).") from None
    if fecha.year < ANIO_MINIMO:
        raise ErrorDeUso(f"{opcion}: el año {fecha.year} no es válido (¿un error de captura?).")
    if fecha.date() > hoy.date():
        raise ErrorDeUso(f"{opcion}: {fecha.strftime(FORMATO_FECHA)} es una fecha futura (hoy es "
                         f"{hoy.strftime(FORMATO_FECHA)}).")
    return fecha.strftime(FORMATO_FECHA)


def resolver_periodo(args, hoy=None):
    """The validated Namespace → {"modo": "single"|"range", "fecha" | "desde"+"hasta", "sociedad",
    "paralelo"} with the dates normalised to DD.MM.AAAA. ErrorDeUso for anything the user must fix."""
    hoy = hoy or datetime.now()
    if not args.corrida_completa:
        raise ErrorDeUso("Falta --corrida-completa (es la única acción de la línea de comandos).")
    fecha, desde, hasta = (_unico(getattr(args, o), f"--{o}") for o in ("fecha", "desde", "hasta"))
    sociedad = _unico(args.sociedad, "--sociedad")

    if fecha is not None and (desde is not None or hasta is not None):
        raise ErrorDeUso("--fecha no se combina con --desde / --hasta: usa un día o un periodo.")
    if fecha is None and (desde is None or hasta is None):
        raise ErrorDeUso("Indica el día con --fecha, o el periodo con --desde y --hasta (los dos).")

    if sociedad is None:
        sociedad = SOCIEDAD_DEFECTO
    sociedad = sociedad.strip().upper()
    if not _RE_SOCIEDAD.fullmatch(sociedad):
        raise ErrorDeUso(f"--sociedad: '{sociedad}' no es una sociedad de SAP (cuatro letras o números, "
                         f"como {SOCIEDAD_DEFECTO}).")

    periodo = {"sociedad": sociedad, "paralelo": bool(args.paralelo)}
    if fecha is not None:
        periodo.update(modo="single", fecha=_fecha_valida(fecha, "--fecha", hoy))
        return periodo
    d = _fecha_valida(desde, "--desde", hoy)
    h = _fecha_valida(hasta, "--hasta", hoy)
    if parse_fecha(d) > parse_fecha(h):
        raise ErrorDeUso(f"El periodo es inválido: --desde ({d}) es posterior a --hasta ({h}).")
    periodo.update(modo="range", desde=d, hasta=h)
    return periodo


# ==================================================================================================
# Output: console (when there is one) + log file
# ==================================================================================================
def _escribir_en(flujo, texto):
    """Write and flush; a console that cannot encode a character gets '?' instead of an error."""
    try:
        flujo.write(texto)
    except UnicodeEncodeError:
        codificacion = getattr(flujo, "encoding", None) or "ascii"
        flujo.write(texto.encode(codificacion, "replace").decode(codificacion))
    flujo.flush()


class _Flujo(io.TextIOBase):
    """What sys.stdout / sys.stderr are replaced with while a run goes: forwards to the `_Registro`."""

    def __init__(self, registro, canal):
        super().__init__()
        self._registro = registro
        self._canal = canal

    encoding = "utf-8"
    errors = "replace"

    def write(self, texto):
        if not isinstance(texto, str):
            raise TypeError(f"write() argument must be str, not {type(texto).__name__}")
        self._registro.escribir_parcial(texto, self._canal)
        return len(texto)

    def writable(self):
        return True

    def flush(self):
        self._registro.vaciar_parcial(self._canal)       # print(..., end="", flush=True): the text goes out now

    def close(self):
        pass                # nobody gets to close the process' output from inside


class _Registro:
    """Everything the run says, to the console (if any) AND to the log file. A sink that fails (full disk,
    closed console) is dropped for good: the other one goes on and the run never dies of its own log."""

    def __init__(self, archivo=None, ruta=None, consola_out=None, consola_err=None, ahora=datetime.now):
        self._candado = threading.RLock()        # the parallel runs print from several threads
        self._archivo = archivo
        self.ruta = ruta
        self._consola = {"out": consola_out, "err": consola_err}
        self._ahora = ahora
        self._pendiente = {}                     # (channel, thread) -> text written without its line break yet
        self.salida = _Flujo(self, "out")
        self.errores = _Flujo(self, "err")

    def escribir(self, texto, canal="out"):
        with self._candado:
            if self._archivo is not None:
                try:
                    _escribir_en(self._archivo, texto)
                except (OSError, ValueError):
                    self._archivo = None
            consola = self._consola.get(canal)
            if consola is not None:
                try:
                    _escribir_en(consola, texto)
                except (OSError, ValueError):
                    self._consola[canal] = None

    def escribir_parcial(self, texto, canal="out"):
        """What `print` does: it writes the text and, in ANOTHER call, the line break. Two threads printing
        at once (the parallel runs) would splice their lines between the two calls, so the text of each
        thread waits in a buffer of its own until its line is complete."""
        with self._candado:
            clave = (canal, threading.get_ident())
            completo, salto, resto = (self._pendiente.pop(clave, "") + texto).rpartition("\n")
            if salto:
                self.escribir(completo + salto, canal)
            if resto:
                self._pendiente[clave] = resto

    def vaciar_parcial(self, canal="out", todos=False):
        """Write what this thread (or, with `todos`, every thread) left without its line break."""
        with self._candado:
            for clave in [k for k in self._pendiente if todos or k == (canal, threading.get_ident())]:
                self.escribir(self._pendiente.pop(clave), clave[0])

    def linea(self, texto, canal="out"):
        """One line with the time of day in front."""
        self.escribir(f"[{self._ahora():%H:%M:%S}] {texto}\n", canal)

    def cerrar(self):
        self.vaciar_parcial(todos=True)
        with self._candado:
            if self._archivo is not None:
                try:
                    self._archivo.close()
                except OSError:
                    pass
                self._archivo = None


def abrir_bitacora(carpeta, ahora):
    """Create (never overwrite) corrida_completa_AAAAMMDD_HHMMSS.log in `carpeta`, falling back to the temp
    folder. → (file, path), or (None, None) when no log can be written (the run goes on with the console)."""
    nombre = f"corrida_completa_{ahora:%Y%m%d_%H%M%S}"
    for destino in (carpeta, os.path.join(tempfile.gettempdir(), "ValidacionFacturaGlobal")):
        try:
            os.makedirs(destino, exist_ok=True)
        except OSError:
            continue
        for intento in range(10):
            ruta = os.path.join(destino, f"{nombre}{'' if intento == 0 else f'_{intento}'}.log")
            try:
                # "x": a second run in the same second must not overwrite the first one's log.
                # utf-8-sig writes its BOM only into a NEW file, so Notepad shows the accents right.
                archivo = open(ruta, "x", encoding="utf-8-sig", errors="replace", buffering=1)
            except FileExistsError:
                continue
            except OSError:
                break
            return archivo, ruta
    return None, None


def conectar_consola():
    """→ (stdout, stderr) this process can show text on; None where it has none.

    A --windowed .exe starts with sys.stdout / sys.stderr = None. On Windows we attach to the console of
    the parent process (cmd, PowerShell) if it has one and write to it as UTF-16 (the console of
    `open("CONOUT$")`), so the accents and the emojis come out right whatever the code page. Started by the
    Task Scheduler or a shortcut there is no parent console: both stay None and the log file is all there is."""
    salida, errores = sys.stdout, sys.stderr
    if (salida is None or errores is None) and sys.platform == "win32":
        try:
            import ctypes
            if ctypes.windll.kernel32.AttachConsole(0xFFFFFFFF):        # ATTACH_PARENT_PROCESS
                consola = io.TextIOWrapper(io.open("CONOUT$", "wb", buffering=0), encoding="utf-8",
                                           errors="replace", line_buffering=True, write_through=True)
                salida = consola if salida is None else salida
                errores = consola if errores is None else errores
        except (OSError, AttributeError, ValueError):
            pass
    return salida, errores


# ==================================================================================================
# One full run at a time
# ==================================================================================================
_mutex_abierto = None          # the handle must stay alive as long as the run does


def adquirir_unica_corrida():
    """True if this process is the only full run of the Windows session (a named mutex; it is released
    when the process ends, even if it crashes). False if another one holds it. Where the mutex cannot be
    created (not Windows, or an odd failure of the API) the run goes on: it returns True, since refusing to
    run over a guard that cannot even be built would block every run for nothing."""
    global _mutex_abierto
    if sys.platform != "win32":
        return True
    try:
        import ctypes
        from ctypes import wintypes
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.CreateMutexW.restype = wintypes.HANDLE
        kernel32.CreateMutexW.argtypes = (wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR)
        kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)
        manejador = kernel32.CreateMutexW(None, False, NOMBRE_MUTEX)
        error = ctypes.get_last_error()
    except (OSError, AttributeError, ImportError):
        return True
    if not manejador:
        return True
    if error == 183:                                     # ERROR_ALREADY_EXISTS
        kernel32.CloseHandle(manejador)
        return False
    _mutex_abierto = manejador
    return True


# ==================================================================================================
# The stand-in for the window
# ==================================================================================================
_PASOS = {"descarga": "Descarga de SAP", "consolidacion": "Consolidación en Excel"}
_ESTADOS = {"en_curso": "en curso", "listo": "listo", "parcial": "terminó con avisos", "error": "con error"}


class GuiSinVentana:
    """What the controller asks of the window, answered from the command line and config.json. It touches
    no Tk, no dialog and NEVER writes config.json (the GUI does, on every read: a scheduled run must not
    change the folders the person last chose). Status lines go to the log; the parallel runs repeat theirs
    every second with a new clock, so a line is logged only when it changes (plus a heartbeat)."""

    LATIDO_SEGUNDOS = 120

    def __init__(self, periodo, ruta_descarga, ruta_consolidacion, registro, reloj=time.monotonic):
        self._periodo = periodo
        self._ruta_descarga = ruta_descarga
        self._ruta_consolidacion = ruta_consolidacion
        self._registro = registro
        self._reloj = reloj
        self._ultimo = None
        self._ultimo_tiempo = 0.0
        self.deshabilitado = False

    # --- what the controller reads ---
    def validate_dates(self):
        return True                 # `resolver_periodo` already did, before anything started

    def get_config(self, transaccion="gallo"):
        p = self._periodo
        if p["modo"] == "single":
            nombre = nombre_archivo_transaccion(transaccion, fecha=p["fecha"])
        else:
            nombre = nombre_archivo_transaccion(transaccion, desde=p["desde"], hasta=p["hasta"])
        cfg = {"mode": p["modo"], "sociedad": p["sociedad"], "input_path": self._ruta_descarga,
               "filename": nombre}
        if p["modo"] == "single":
            cfg["fecha"] = p["fecha"]
        else:
            cfg["date_from"], cfg["date_to"] = p["desde"], p["hasta"]
        return cfg

    def get_consolidacion_path(self):
        os.makedirs(self._ruta_consolidacion, exist_ok=True)
        return self._ruta_consolidacion

    def usar_paralelo(self):
        return self._periodo["paralelo"]

    # --- what the controller reports ---
    def set_status(self, mensaje):
        texto = " ".join(str(mensaje).split())
        sin_reloj = _RE_CRONOMETRO.sub("", texto)
        ahora = self._reloj()
        if sin_reloj == self._ultimo and ahora - self._ultimo_tiempo < self.LATIDO_SEGUNDOS:
            return
        self._ultimo, self._ultimo_tiempo = sin_reloj, ahora
        self._registro.linea(texto)

    def marcar_paso(self, paso, estado):
        self._registro.linea(f"▶ {_PASOS.get(paso, paso)}: {_ESTADOS.get(estado, estado)}")

    def disable_buttons(self):
        self.deshabilitado = True

    def enable_buttons(self):
        self.deshabilitado = False


# ==================================================================================================
# Entry point
# ==================================================================================================
def _encabezado(registro, periodo, ruta_descarga, ruta_consolidacion):
    """The first lines of the log: what was asked and where everything goes."""
    p = periodo
    cuando = f"{p['fecha']}" if p["modo"] == "single" else f"{p['desde']} — {p['hasta']}"
    registro.escribir(
        "Validación Factura Global — corrida completa (línea de comandos)\n"
        f"Inicio            : {datetime.now():%Y-%m-%d %H:%M:%S}\n"
        f"Periodo           : {cuando}\n"
        f"Sociedad          : {p['sociedad']}\n"
        f"Descargas         : {'en paralelo, por pasadas' if p['paralelo'] else 'una tras otra'}\n"
        f"Carpeta de descarga     : {ruta_descarga}\n"
        f"Carpeta de consolidación: {ruta_consolidacion}\n"
        f"Bitácora          : {registro.ruta or '(no se pudo crear; sólo consola)'}\n\n")


def _correr(periodo, registro):
    """The run itself → exit code. Imports the controller NOW (pandas, pywin32 and the whole SAP code), so a
    wrong command line never pays for it and an import error is logged like any other."""
    from controller import ValidacionFacturaController

    ruta_descarga, ruta_consolidacion = carpetas_configuradas()
    _encabezado(registro, periodo, ruta_descarga, ruta_consolidacion)
    gui = GuiSinVentana(periodo, ruta_descarga, ruta_consolidacion, registro)
    controlador = ValidacionFacturaController(gui)
    salida = controlador.correr_corrida_completa()

    titulo, cuerpo = controlador.texto_final_corrida(salida)
    codigo = _SALIDA_POR_ESTADO.get(salida.get("estado"), SALIDA_FALLO)
    registro.escribir(f"\n{'=' * 70}\n{titulo}\n{'=' * 70}\n{cuerpo}\n\n")
    registro.linea(f"Código de salida: {codigo}")
    return codigo


def ejecutar_linea_de_comandos(argv, hoy=None):
    """The command line of the .exe: `argv` is sys.argv[1:] → the process exit code. Never raises."""
    pide_ayuda = any(a in ("-h", "--help") for a in argv)
    consola_out, consola_err = conectar_consola()
    archivo, ruta = (None, None) if pide_ayuda else abrir_bitacora(
        os.path.join(carpeta_base_proyecto(), "Logs"), datetime.now())
    registro = _Registro(archivo, ruta, consola_out, consola_err)
    anterior = sys.stdout, sys.stderr
    sys.stdout, sys.stderr = registro.salida, registro.errores
    try:
        try:
            args = parsear(argv)
            if args.ayuda:
                registro.escribir(texto_ayuda())
                return SALIDA_COMPLETA
            periodo = resolver_periodo(args, hoy)
        except ErrorDeUso as e:
            registro.escribir(f"Error: {e}\nEscribe  ValidacionFacturaGlobal.exe --help  para ver cómo se usa.\n",
                              "err")
            return SALIDA_USO
        if not adquirir_unica_corrida():
            registro.escribir("Ya hay otra corrida completa en curso en esta sesión de Windows: esta no se "
                              "ejecutó (dos a la vez se pelearían las ventanas de SAP, el portapapeles y los "
                              "archivos).\n", "err")
            return SALIDA_OCUPADO
        return _correr(periodo, registro)
    except KeyboardInterrupt:
        registro.escribir("\nCorrida interrumpida por el usuario (Ctrl+C).\n", "err")
        return SALIDA_INTERRUMPIDA
    except (Exception, SystemExit):
        registro.escribir(f"\nError inesperado:\n{traceback.format_exc()}\n", "err")
        return SALIDA_FALLO
    finally:
        sys.stdout, sys.stderr = anterior
        registro.cerrar()
