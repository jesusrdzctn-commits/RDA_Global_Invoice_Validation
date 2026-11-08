"""Validación Factura Global — GUI (tkinter).

A slate side rail on the left holds the period of the run and the progress of the WHOLE process
(download of every SAP transaction -> consolidation in Excel), one coloured dot per step. The work
area holds one card per phase and a status strip whose bar fills while a process runs. Colour carries meaning (green done, amber warnings, red error), never decoration.

The controller wires the on_* callbacks and reports the rail with marcar_paso(). `python
interfaz_GUI.py` opens the window on its own. This module imports no SAP dependencies on purpose
(a test enforces it).
"""

import json
import os
import tkinter as tk
import tkinter.font as tkfont
from datetime import datetime
from tkinter import filedialog, messagebox

from utils import (
    nombre_archivo_dia,          nombre_archivo_rango,
    nombre_archivo_dia_monivoi,  nombre_archivo_rango_monivoi,
    nombre_archivo_dia_vbrl,     nombre_archivo_rango_vbrl,
    nombre_archivo_dia_fb03,     nombre_archivo_rango_fb03,
    nombre_archivo_rango_eiap,
    nombre_archivo_rango_dsd,
    ARCHIVOS_SEGUNDA_CORRIDA,    CORRIDAS_DESCARGA,
    ruta_config,                 carpetas_por_defecto,       carpetas_configuradas,
    nombre_archivo_dia_2da,      nombre_archivo_rango_2da,
    parse_fecha,
)

FUENTE = "Segoe UI"
FUENTE_MONO = "Consolas"


# ----------------------------------------------------------------------
# Palette, step states and the steps of the process
# ----------------------------------------------------------------------
class Paleta:
    fondo       = "#EEF3F6"   # work area
    tarjeta     = "#FFFFFF"
    borde       = "#D6E0E6"
    campo       = "#F2F6F8"   # entries and lists
    riel        = "#0E2431"   # side rail (slate)
    riel_suave  = "#1B3B4D"   # lines and idle pills on the rail
    riel_brillo = "#2B5A72"   # light edge of the rail (catches the "light")
    sombra      = "#04121A"   # tint of the shadow the rail casts
    riel_texto  = "#9DB8C6"   # secondary text on the rail
    riel_apagado = "#5F7F91"  # pending dots on the rail
    texto       = "#10232F"
    suave       = "#5B7080"   # secondary text
    acento      = "#0E7C8F"   # primary actions (teal)
    acento_vivo = "#2CC4D6"   # the same accent, bright enough for the dark rail
    pista       = "#DCE6EB"   # empty track of the progress bar
    foco        = "#5EC8D8"   # keyboard focus ring
    inactivo_fondo  = "#E1E8EC"
    inactivo_borde  = "#CBD6DC"
    inactivo_texto  = "#9AAAB5"
    verde       = "#188A4C"
    ambar       = "#E59A0B"
    rojo        = "#D64545"


# state -> (label, text colour, chip background, symbol)
ESTADOS = {
    "pendiente": ("Pendiente",  "#5B7080", "#E6EDF1", "○"),
    "siguiente": ("Siguiente",  "#0B6B7C", "#D6EEF2", "▶"),
    "en_curso":  ("En curso",   "#0B6B7C", "#D6EEF2", "●"),
    "listo":     ("Listo",      "#13733F", "#DDF3E6", "✓"),
    "parcial":   ("Con avisos", "#8F5C00", "#FCEFD2", "!"),
    "error":     ("Con error",  "#B83232", "#FBE0E0", "✕"),
}

# State colours on the dark rail: (ring, fill, symbol, state text). A None fill means "rail colour".
ESTADOS_RIEL = {
    "pendiente": (Paleta.riel_apagado, None,               Paleta.riel_apagado, Paleta.riel_texto),
    "siguiente": (Paleta.acento_vivo,  Paleta.riel_suave,  "#FFFFFF",           Paleta.acento_vivo),
    "en_curso":  (Paleta.acento_vivo,  Paleta.acento_vivo, Paleta.riel,         Paleta.acento_vivo),
    "listo":     (Paleta.verde,        Paleta.verde,       "#FFFFFF",           "#6FD39A"),
    "parcial":   (Paleta.ambar,        Paleta.ambar,       "#FFFFFF",           "#F2B955"),
    "error":     (Paleta.rojo,         Paleta.rojo,        "#FFFFFF",           "#FF8A8A"),
}

# step key -> (number, title on the rail). The order is the order of the process. The keys are also the
# keys of the state chips of the two cards.
PASOS = (
    ("descarga",      "1", "Descarga de SAP"),
    ("consolidacion", "2", "Consolidación en Excel"),
)


def _mezclar(color, otro, t):
    """`color` moved a fraction `t` (0-1) towards `otro`; both '#RRGGBB' (hover / pressed shades)."""
    a = [int(color[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(otro[i:i + 2], 16) for i in (1, 3, 5)]
    return "#%02X%02X%02X" % tuple(round(x + (y - x) * t) for x, y in zip(a, b))


def _degradado_sombra(fondo, tinte, ancho, opacidad):
    """`ancho` hex colours: `tinte` blended over `fondo` at `opacidad` next to the rail, fading to plain
    `fondo`. Tk has no alpha channel, so the blend is computed against the known background."""
    return [_mezclar(fondo, tinte, opacidad * (1 - i / ancho) ** 2) for i in range(ancho)]


def _rect_redondeado(canvas, x1, y1, x2, y2, radio, **opciones):
    """Rounded rectangle as a smoothed polygon (a Canvas has no rounded rectangle item)."""
    r = max(0, min(radio, (x2 - x1) / 2, (y2 - y1) / 2))
    puntos = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
              x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return canvas.create_polygon(puntos, smooth=True, **opciones)


def _ruta_config():
    """config.json (last download/consolidation folders): utils.ruta_config, shared with the command line."""
    return ruta_config()


class _Fijo:
    """Stands in for a tk variable where only `.get()` is read (nothing can trace it or write to it)."""

    def __init__(self, valor):
        self._valor = valor

    def get(self):
        return self._valor


# ----------------------------------------------------------------------
# Custom widgets (a Canvas can round corners; tk.Button / tk.Checkbutton cannot)
#
# They answer the slice of the tk.Button / tk.Radiobutton / tk.Checkbutton API the app and its tests
# use: invoke(), cget("state" | "text"), configure(state=..., text=...). Every one works with the
# keyboard too (Tab to reach it, Enter / Space to press it).
# ----------------------------------------------------------------------
class _WidgetLienzo(tk.Canvas):
    """Common ground of the custom widgets: enabled flag, focus ring state and the tk-like API."""

    def __init__(self, parent, fondo_padre, **opciones):
        super().__init__(parent, bg=fondo_padre, highlightthickness=0, bd=0, takefocus=1,
                         cursor="hand2", **opciones)
        self._texto = ""
        self._habilitado = True
        self._hover = self._presionado = self._foco = False

    # --- tk-like API ---------------------------------------------------
    def cget(self, clave):
        if clave == "state":
            return "normal" if self._habilitado else "disabled"
        if clave == "text":
            return self._texto
        return super().cget(clave)

    __getitem__ = cget

    def configure(self, cnf=None, **opciones):
        """`state=` and `text=` act on the widget; anything else goes to the Canvas."""
        if cnf is None and not opciones:
            return super().configure()                         # the query form
        if isinstance(cnf, dict):
            opciones = {**cnf, **opciones}
        elif cnf is not None:
            return super().configure(cnf)                      # a single option name: query it
        estado = opciones.pop("state", None)
        texto = opciones.pop("text", None)
        if estado is not None:
            self.habilitar(estado != "disabled")
        if texto is not None:
            self.poner_texto(texto)
        if opciones:
            return super().configure(**opciones)
        return None

    config = configure

    def habilitar(self, habilitado):
        self._habilitado = bool(habilitado)
        super().configure(cursor="hand2" if self._habilitado else "arrow",
                          takefocus=1 if self._habilitado else 0)
        if not self._habilitado:
            self._hover = self._presionado = False
        self._dibujar()

    def poner_texto(self, texto):
        self._texto = texto
        self._al_cambiar_texto()
        self._dibujar()

    def invoke(self):                                          # pragma: no cover - overridden
        raise NotImplementedError

    # --- hooks for the subclasses ----------------------------------------
    def _al_cambiar_texto(self):
        pass

    def _dibujar(self):
        """Default: a rounded rectangle with centred text, coloured by `_colores()` (the button and the pills);
        the switch draws itself."""
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w <= 2 or h <= 2:
            return
        relleno, borde, texto = self._colores()
        _rect_redondeado(self, 1, 1, w - 1, h - 1, self._radio, fill=relleno, outline=borde, width=2)
        self.create_text(w / 2, h / 2, text=self._texto, fill=texto, font=self._fuente)

    # --- shared event plumbing -------------------------------------------
    def _enlazar_eventos(self):
        self.bind("<Enter>",           lambda _e: self._cambiar(_hover=True))
        self.bind("<Leave>",           lambda _e: self._cambiar(_hover=False, _presionado=False))
        self.bind("<ButtonPress-1>",   self._al_presionar)
        self.bind("<ButtonRelease-1>", self._al_soltar)
        self.bind("<FocusIn>",         lambda _e: self._cambiar(_foco=True))
        self.bind("<FocusOut>",        lambda _e: self._cambiar(_foco=False))
        self.bind("<Return>",          lambda _e: self.invoke())
        self.bind("<space>",           lambda _e: self.invoke())

    def _cambiar(self, **cambios):
        for clave, valor in cambios.items():
            setattr(self, clave, valor)
        self._dibujar()

    def _al_presionar(self, _evento):
        if self._habilitado:
            self._cambiar(_presionado=True)

    def _al_soltar(self, evento):
        dentro = 0 <= evento.x <= self.winfo_width() and 0 <= evento.y <= self.winfo_height()
        fue_presionado = self._presionado
        self._cambiar(_presionado=False)
        if self._habilitado and dentro and fue_presionado:
            self.invoke()


class BotonRedondeado(_WidgetLienzo):
    """A flat, rounded, hover-aware button.

    estilo: "solido"   filled with `color` (the primary actions)
            "contorno" white with a `color` outline (secondary actions)
            "sutil"    light grey (utility buttons such as 'Examinar...')
    """
    ESTILOS = ("solido", "contorno", "sutil")

    def __init__(self, parent, texto, comando, color, fondo_padre, *, estilo="solido", fg="#FFFFFF",
                 fuente=(FUENTE, 10, "bold"), alto=42, radio=10, ancho=None, margen=32):
        if estilo not in self.ESTILOS:
            raise ValueError(f"Estilo de botón desconocido: {estilo!r}")
        super().__init__(parent, fondo_padre, height=alto)
        self._texto, self._comando = texto, comando
        self.color, self._fg, self._fuente = color, fg, fuente
        self._estilo, self._radio, self._ancho_fijo, self._margen = estilo, radio, ancho, margen
        self._reservar_ancho()
        self._enlazar_eventos()
        self.bind("<Configure>", lambda _e: self._dibujar())

    def _reservar_ancho(self):
        # Minimum width = the text plus margins. A Canvas asks for ~394 px by default, which would force
        # every column that wide; buttons packed with fill="x" stretch to their card anyway.
        ancho = self._ancho_fijo or tkfont.Font(root=self, font=self._fuente).measure(self._texto) + self._margen
        super().configure(width=ancho)

    def _al_cambiar_texto(self):
        self._reservar_ancho()

    def _colores(self):
        """(fill, outline, text colour) for the current state."""
        activo, hover, presionado = self._habilitado, self._hover, self._presionado
        if self._estilo == "solido":
            if not activo:
                return Paleta.inactivo_fondo, Paleta.inactivo_fondo, Paleta.inactivo_texto
            relleno = (_mezclar(self.color, "#000000", 0.24) if presionado
                       else _mezclar(self.color, "#000000", 0.10) if hover else self.color)
            return relleno, (Paleta.foco if self._foco else relleno), self._fg
        if self._estilo == "contorno":
            if not activo:
                return Paleta.tarjeta, Paleta.inactivo_borde, Paleta.inactivo_texto
            relleno = (_mezclar(Paleta.tarjeta, self.color, 0.20) if presionado
                       else _mezclar(Paleta.tarjeta, self.color, 0.09) if hover else Paleta.tarjeta)
            return relleno, (Paleta.foco if self._foco else self.color), self.color
        # "sutil"
        if not activo:
            return Paleta.inactivo_fondo, Paleta.inactivo_fondo, Paleta.inactivo_texto
        relleno = (_mezclar(Paleta.campo, "#000000", 0.12) if presionado
                   else _mezclar(Paleta.campo, "#000000", 0.05) if hover else Paleta.campo)
        return relleno, (Paleta.foco if self._foco else Paleta.borde), self.color

    def invoke(self):
        if self._habilitado and self._comando:
            self._comando()


class Segmento(_WidgetLienzo):
    """One pill of a segmented control: selected when `variable` equals `valor`. invoke() selects it
    (sets the variable, then calls `comando`), like tk.Radiobutton.invoke()."""

    def __init__(self, parent, texto, variable, valor, comando, fondo_padre, *, oscuro=False,
                 alto=28, radio=9, margen=26, fuente=(FUENTE, 9, "bold")):
        super().__init__(parent, fondo_padre, height=alto)
        self._texto, self._variable, self._valor, self._comando = texto, variable, valor, comando
        self._oscuro, self._radio, self._fuente = oscuro, radio, fuente
        super().configure(width=tkfont.Font(root=self, font=fuente).measure(texto) + margen)
        self._enlazar_eventos()
        self.bind("<Configure>", lambda _e: self._dibujar())
        self._traza = variable.trace_add("write", lambda *_: self._redibujar_seguro())
        self.bind("<Destroy>", self._al_destruir, add="+")

    def _al_destruir(self, evento):
        if evento.widget is not self:
            return
        try:
            self._variable.trace_remove("write", self._traza)
        except tk.TclError:
            pass                                               # the interpreter is going away

    def _redibujar_seguro(self):
        try:
            self._dibujar()
        except tk.TclError:
            pass

    @property
    def seleccionado(self):
        return self._variable.get() == self._valor

    def _colores(self):
        """(fill, outline, text colour) for the current state."""
        if self._oscuro:
            sel = (Paleta.tarjeta, Paleta.tarjeta, Paleta.riel)
            base = (Paleta.riel_suave, Paleta.riel_suave, Paleta.riel_texto)
            hover = _mezclar(Paleta.riel_suave, "#FFFFFF", 0.10)
            fondo = Paleta.riel
        else:
            sel = (Paleta.acento, Paleta.acento, "#FFFFFF")
            base = (Paleta.campo, Paleta.borde, Paleta.suave)
            hover = _mezclar(Paleta.campo, "#000000", 0.05)
            fondo = Paleta.tarjeta
        if self.seleccionado:
            relleno, borde, texto = sel
            if not self._habilitado:                           # still looks selected, only muted
                relleno, borde = _mezclar(relleno, fondo, 0.45), _mezclar(borde, fondo, 0.45)
            elif self._hover:
                relleno = borde = _mezclar(relleno, "#000000", 0.08)
        else:
            relleno, borde, texto = base
            if not self._habilitado:
                texto = _mezclar(texto, fondo, 0.5)
            elif self._hover:
                relleno = hover
        if self._foco and self._habilitado:
            borde = Paleta.foco
        return relleno, borde, texto

    def invoke(self):
        if not self._habilitado:
            return
        self._variable.set(self._valor)
        if self._comando:
            self._comando()


class Interruptor(_WidgetLienzo):
    """On / off switch with its label, bound to a tk.BooleanVar (the tk.Checkbutton of the app)."""

    def __init__(self, parent, texto, variable, fondo_padre, *, alto=28, pista=(42, 22),
                 fuente=(FUENTE, 9, "bold"), brecha=10):
        self._pista_w, self._pista_h, self._brecha, self._fuente = pista[0], pista[1], brecha, fuente
        medida = tkfont.Font(root=parent, font=fuente).measure(texto)
        super().__init__(parent, fondo_padre, height=alto, width=pista[0] + brecha + medida + 6)
        self._alto = alto
        self._texto, self._variable = texto, variable
        self._enlazar_eventos()
        self._traza = variable.trace_add("write", lambda *_: self._redibujar_seguro())
        self.bind("<Destroy>", self._al_destruir, add="+")
        self._dibujar()

    _al_destruir = Segmento._al_destruir
    _redibujar_seguro = Segmento._redibujar_seguro

    def _dibujar(self):
        self.delete("all")
        pw, ph, alto = self._pista_w, self._pista_h, self._alto
        y1 = (alto - ph) / 2
        encendido = bool(self._variable.get())
        if not self._habilitado:
            pista, perilla, texto = Paleta.inactivo_fondo, Paleta.tarjeta, Paleta.inactivo_texto
        else:
            pista = Paleta.acento if encendido else "#B4C3CC"
            if self._hover:
                pista = _mezclar(pista, "#000000", 0.10)
            perilla, texto = "#FFFFFF", Paleta.texto
        borde = Paleta.foco if (self._foco and self._habilitado) else pista
        _rect_redondeado(self, 2, y1, 2 + pw, y1 + ph, ph / 2, fill=pista, outline=borde, width=2)
        margen = 3
        d = ph - 2 * margen
        x = 2 + (pw - margen - d if encendido else margen)
        self.create_oval(x, y1 + margen, x + d, y1 + margen + d, fill=perilla, outline=perilla)
        self.create_text(2 + pw + self._brecha, alto / 2, text=self._texto, fill=texto,
                         font=self._fuente, anchor="w")

    def invoke(self):
        if not self._habilitado:
            return
        self._variable.set(not bool(self._variable.get()))
        self._dibujar()


class BarraProgreso(tk.Canvas):
    """A thin rounded 0-100 progress bar. `barra.establecer(40, color)` moves (and recolours) it."""

    def __init__(self, parent, fondo, alto=8):
        super().__init__(parent, height=alto, bg=fondo, highlightthickness=0, bd=0)
        self._valor, self._color = 0.0, Paleta.acento
        self.bind("<Configure>", lambda _e: self._dibujar())

    @property
    def valor(self):
        return self._valor

    @property
    def color(self):
        return self._color

    def establecer(self, valor, color=None):
        self._valor = max(0.0, min(100.0, float(valor)))
        if color is not None:
            self._color = color
        self._dibujar()

    def _dibujar(self):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w <= 2 or h <= 2:
            return
        r = h / 2
        self.create_line(r, r, w - r, r, width=h, capstyle=tk.ROUND, fill=Paleta.pista)
        if self._valor > 0:
            self.create_line(r, r, r + (w - h) * self._valor / 100.0, r, width=h, capstyle=tk.ROUND,
                             fill=self._color)


# ----------------------------------------------------------------------
# The window
# ----------------------------------------------------------------------
class ValidacionFacturaGUI:
    """Side rail (period + progress of the process), cards (SAP downloads, Excel consolidation, folders
    and parameters) and a status strip. Method sections: config.json, interface, rail steps, pickers,
    file names, validation, status, button state, handlers."""

    # Design size = MINIMUM size: _ajustar_al_contenido grows the window when the content needs more room.
    ANCHO_BASE = 1040
    ANCHO_RIEL = 270
    # Dates that give the longest file names (SAP-style 'D_mmm_YY' with a two-digit day): the width of the
    # 'Archivos que se generarán' box is measured with them.
    FECHA_LARGA_DESDE, FECHA_LARGA_HASTA = "28.10.2026", "30.12.2026"

    LEYENDA_CONSOLIDACION = ("Gallo + Monivoi + VBRL + FB03 + las 7 pestañas de la 2a corrida "
                             "→ Excel con Diferencias y Agrupadores")

    def __init__(self, root):
        self.root = root
        self.root.title("Robotic Desktop Automation — Validación Factura Global")
        self.root.configure(bg=Paleta.fondo)
        # Pixel sizes follow the display scaling; point sizes of fonts already do.
        self.escala = max(1.0, self.root.winfo_fpixels("1i") / 96.0)

        # Callbacks the controller wires up
        self.on_download_gallo   = None
        self.on_download_monivoi = None
        self.on_download_dsd     = None
        self.on_download_eiap    = None
        self.on_download_todos   = None
        self.on_consolidar       = None
        self.on_corrida_completa = None

        self.mode_var = tk.StringVar(value="single")

        # Parallel run by passes (opt-in): always starts off (sequential is the safe default).
        self.paralelo_var = tk.BooleanVar(value=False)

        self._config_file = _ruta_config()
        cfg = self._cargar_config()
        ruta_input_default, _ = carpetas_por_defecto()
        ruta_descarga, ruta_consolidacion = carpetas_configuradas(cfg)
        self.path_var = tk.StringVar(value=ruta_descarga)
        self.consolidacion_path_var = tk.StringVar(value=ruta_consolidacion)
        os.makedirs(ruta_input_default, exist_ok=True)

        hoy = datetime.now().strftime("%d.%m.%Y")
        self.date_single_var = tk.StringVar(value=hoy)
        self.date_from_var   = tk.StringVar(value=hoy)
        self.date_to_var     = tk.StringVar(value=hoy)
        self.sociedad_var    = tk.StringVar(value="MX21")
        self.status_var      = tk.StringVar(value="✓ Listo para comenzar")
        self.filename_var    = tk.StringVar()
        self.resumen_periodo_var = tk.StringVar()

        # Progress of the process, as reported by the controller (marcar_paso).
        self._estado_paso = {clave: "pendiente" for clave, *_ in PASOS}
        self._chips_paso = {}          # "descarga" / "consolidacion" -> state chip of the card
        self._riel_paso = {}           # step key -> (circle canvas, state label, number)
        self._conectores = []          # lines between the dots of the rail
        self._resumenes_periodo = []   # the period hint of each date variant (they share one text)
        self._ocupado = False
        self._barra_cerrada = False    # the bar already shows the end state of a step
        self._tono = "ok"              # tone of the last status message
        self._ultimo_mensaje = None    # the last status message (without its clock): what the bar paced on

        self._build_ui()
        self._refrescar_pasos()
        self._actualizar_resumen_periodo()
        # Last: the traces would otherwise fire while the window is still being assembled.
        for variable in (self.date_single_var, self.date_from_var, self.date_to_var):
            variable.trace_add("write", lambda *_: self._al_cambiar_periodo())
        self.mode_var.trace_add("write", lambda *_: self._on_mode_change())
        self._on_mode_change()
        self._ajustar_al_contenido()

    def _px(self, valor):
        """A design size in pixels, scaled for the display."""
        return int(round(valor * self.escala))

    # ------------------------------------------------------------------
    # Window geometry
    # ------------------------------------------------------------------
    def _igualar_columna_izquierda(self):
        """The 'Archivos que se generarán' box and the 'Consolidación en Excel' card get exactly the same
        width: the largest that either needs. The file names change with the period (a range is longer than
        a day; a day number has no leading zero), so the box is measured with the LONGEST names of both
        modes: nothing moves when the mode or the dates change."""
        px = self._px
        necesario = 0
        for modo in ("single", "range"):
            self.filename_var.set(self._texto_archivos(self._lineas_con_fechas_largas(modo)))
            self.root.update_idletasks()
            necesario = max(necesario, self._izquierda_descarga.winfo_reqwidth())
        self._update_filename_preview()                       # back to the names of the dates in the entries
        necesario = max(necesario, self._marco_consolidacion.winfo_reqwidth())
        # Each column also holds the gap that separates it from its neighbour (16 px / 6 px).
        self._cuerpo_descarga.columnconfigure(0, minsize=necesario + px(16))
        self._fila_inferior.columnconfigure(0, minsize=necesario + px(6))

    def _ajustar_al_contenido(self):
        """Fit the window to what the layout really needs (fonts and DPI change it), within the screen."""
        # The date block always reserves the room of its taller variant, so nothing moves with the mode.
        alto_fechas = 0
        for marco in (self.single_frame, self.range_frame):
            self.single_frame.pack_forget()
            self.range_frame.pack_forget()
            marco.pack(fill="x")
            self.root.update_idletasks()
            alto_fechas = max(alto_fechas, marco.winfo_reqheight())
        self.date_container.configure(height=alto_fechas)
        self._on_mode_change()
        self._igualar_columna_izquierda()

        self.root.update_idletasks()
        # The rail is as tall as its content asks for (it cannot propagate its size: its width is fixed).
        self.riel.configure(height=self._riel_contenido.winfo_reqheight())
        self.root.update_idletasks()
        pantalla_w, pantalla_h = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        ancho = min(max(int(self.ANCHO_BASE * self.escala), self.root.winfo_reqwidth()), pantalla_w - 20)
        alto = min(self.root.winfo_reqheight(), pantalla_h - 60)
        x = max(0, (pantalla_w - ancho) // 2)
        y = max(0, (pantalla_h - alto) // 3)
        self.root.geometry(f"{ancho}x{alto}+{x}+{y}")
        self.root.minsize(min(self.root.winfo_reqwidth(), ancho), min(self.root.winfo_reqheight(), alto))

    # ------------------------------------------------------------------
    # config.json
    # ------------------------------------------------------------------
    def _cargar_config(self):
        """Read the whole config.json (dict). If it does not exist or fails → {}."""
        try:
            with open(self._config_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except (FileNotFoundError, json.JSONDecodeError, OSError, ValueError):
            return {}

    def _guardar_config(self):
        """Save both folders (download and consolidation) to config.json."""
        try:
            with open(self._config_file, "w", encoding="utf-8") as f:
                json.dump(
                    {
                        "ruta_descarga":      self.path_var.get().strip(),
                        "ruta_consolidacion": self.consolidacion_path_var.get().strip(),
                    },
                    f, ensure_ascii=False, indent=2,
                )
        except OSError as e:
            print(f"[CONFIG] No se pudo guardar la configuración: {e}")

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
    def _build_ui(self):
        px = self._px
        self.riel = tk.Frame(self.root, bg=Paleta.riel, width=px(self.ANCHO_RIEL))
        self.riel.pack(side=tk.LEFT, fill="y")
        self.riel.pack_propagate(False)
        self._riel_contenido = tk.Frame(self.riel, bg=Paleta.riel)
        self._riel_contenido.pack(fill="both", expand=True)
        self._build_riel(self._riel_contenido)

        principal = tk.Frame(self.root, bg=Paleta.fondo, padx=px(20), pady=px(16))
        principal.pack(side=tk.LEFT, fill="both", expand=True)
        self._dibujar_sombra_riel(self.riel)

        # Status strip first (side=bottom) so it is never pushed out when the window is short.
        self._build_pie(principal)

        self._build_tarjeta_descarga(principal).pack(fill="x")
        abajo = tk.Frame(principal, bg=Paleta.fondo)
        abajo.pack(fill="both", expand=True, pady=(px(12), 0))
        # Column 0 (Consolidación) has a fixed width, set by _igualar_columna_izquierda to the same width
        # as the 'Archivos que se generarán' box; the folders card takes all the rest.
        abajo.columnconfigure(1, weight=1)
        abajo.rowconfigure(0, weight=1)
        self._fila_inferior = abajo
        self._marco_consolidacion = self._build_tarjeta_consolidacion(abajo)
        self._marco_consolidacion.grid(row=0, column=0, sticky="nsew", padx=(0, px(6)))
        self._build_tarjeta_carpetas(abajo).grid(row=0, column=1, sticky="nsew", padx=(px(6), 0))

    # --- side rail ----------------------------------------------------
    def _dibujar_sombra_riel(self, riel):
        """3D look: a soft shadow cast on the work area plus a light 1 px edge on the rail itself."""
        ancho = self._px(18)      # narrower than the 20 px gutter next to the rail, so it never touches a card
        sombra = tk.Canvas(self.root, width=ancho, bg=Paleta.fondo, highlightthickness=0, bd=0)
        alto = self.root.winfo_screenheight()   # taller than any window: the canvas clips the lines
        for x, color in enumerate(_degradado_sombra(Paleta.fondo, Paleta.sombra, ancho, 0.34)):
            sombra.create_line(x, 0, x, alto, fill=color)
        # Overlay glued to the rail's right edge: it takes no layout space.
        sombra.place(in_=riel, relx=1.0, y=0, relheight=1)
        tk.Frame(riel, bg=Paleta.riel_brillo, width=1).place(relx=1.0, x=-1, y=0, relheight=1)

    def _build_riel(self, riel):
        px = self._px
        marca = tk.Frame(riel, bg=Paleta.riel, padx=px(22), pady=px(12))
        marca.pack(fill="x")
        tk.Label(marca, text="RDA", font=(FUENTE, 24, "bold"), fg="#FFFFFF",
                 bg=Paleta.riel).pack(anchor="w")
        tk.Label(marca, text="Validación Factura Global", font=(FUENTE, 11, "bold"), fg="#FFFFFF",
                 bg=Paleta.riel).pack(anchor="w")
        tk.Label(marca, text="Análisis Contabilidad vs Monitoreo\n(SAT) de Facturas en CEDIS", font=(FUENTE, 9),
                 fg=Paleta.riel_texto, bg=Paleta.riel, justify="left").pack(anchor="w", pady=(px(4), 0))

        self._separador_riel(riel)
        self._build_periodo(riel)
        self._separador_riel(riel)

        # Progress of the whole process
        avance = tk.Frame(riel, bg=Paleta.riel, padx=px(22), pady=px(4))
        avance.pack(fill="x")
        self._titulo_riel(avance, "AVANCE DEL PROCESO")
        for posicion, (clave, numero, titulo) in enumerate(PASOS):
            if posicion:
                conector = tk.Frame(avance, bg=Paleta.riel_suave, width=2, height=px(6))
                conector.pack(anchor="w", padx=(px(13), 0))
                self._conectores.append(conector)
            fila = tk.Frame(avance, bg=Paleta.riel)
            fila.pack(fill="x", pady=(px(6) if posicion == 0 else 0, 0))
            circulo = tk.Canvas(fila, width=px(28), height=px(28), bg=Paleta.riel, highlightthickness=0)
            circulo.pack(side=tk.LEFT)
            textos = tk.Frame(fila, bg=Paleta.riel)
            textos.pack(side=tk.LEFT, padx=(px(10), 0))
            tk.Label(textos, text=titulo, font=(FUENTE, 10, "bold"), fg="#FFFFFF",
                     bg=Paleta.riel, anchor="w").pack(anchor="w")
            estado = tk.Label(textos, text="", font=(FUENTE, 8), fg=Paleta.riel_texto,
                              bg=Paleta.riel, anchor="w")
            estado.pack(anchor="w")
            self._riel_paso[clave] = (circulo, estado, numero)

        # Footer of the rail: what to have ready, in plain words.
        tk.Label(
            riel,
            text="Antes de descargar: SAP GUI abierto\ncon sesión iniciada. Cada proceso\n"
                 "pide confirmación antes de empezar.",
            font=(FUENTE, 8), fg=Paleta.riel_texto, bg=Paleta.riel, justify="left",
        ).pack(side=tk.BOTTOM, anchor="w", padx=px(22), pady=px(12))

    def _build_periodo(self, riel):
        """Period of the run: one day or a range. It belongs to the whole process (file names, SAP queries)."""
        px = self._px
        periodo = tk.Frame(riel, bg=Paleta.riel, padx=px(22), pady=px(4))
        periodo.pack(fill="x")
        self._titulo_riel(periodo, "PERIODO A VALIDAR")

        modos = tk.Frame(periodo, bg=Paleta.riel)
        modos.pack(fill="x", pady=(px(6), 0))
        # No command: the mode_var trace (added once the window is built) already runs _on_mode_change.
        self.radios_modo = self._segmentos(
            modos, self.mode_var, (("single", "Un solo día"), ("range", "Varios días")),
            None, fondo=Paleta.riel, oscuro=True,
        )

        # Both variants live in a container as tall as the taller one (see _ajustar_al_contenido).
        self.date_container = tk.Frame(periodo, bg=Paleta.riel, height=1)
        self.date_container.pack(fill="x", pady=(px(6), 0))
        self.date_container.pack_propagate(False)

        self.single_frame = tk.Frame(self.date_container, bg=Paleta.riel)
        self._fila_fecha(self.single_frame, 0, "Fecha", self.date_single_var)
        self._resumen_periodo(self.single_frame, 1)

        self.range_frame = tk.Frame(self.date_container, bg=Paleta.riel)
        self._fila_fecha(self.range_frame, 0, "Desde", self.date_from_var)
        self._fila_fecha(self.range_frame, 1, "Hasta", self.date_to_var)
        self._resumen_periodo(self.range_frame, 2)

    def _resumen_periodo(self, marco, fila):
        """Live hint under the dates of one variant: how many days the run covers, or what is wrong."""
        etiqueta = tk.Label(
            marco, textvariable=self.resumen_periodo_var, font=(FUENTE, 9), fg="#FFFFFF",
            bg=Paleta.riel, justify="left", anchor="nw", height=2,   # always 2 lines: the block never changes size
            wraplength=self._px(226),
        )
        etiqueta.grid(row=fila, column=0, columnspan=2, sticky="w", pady=(self._px(4), 0))
        self._resumenes_periodo.append(etiqueta)

    def _fila_fecha(self, marco, fila, etiqueta, variable):
        px = self._px
        tk.Label(marco, text=etiqueta, font=(FUENTE, 10, "bold"), fg="#FFFFFF", bg=Paleta.riel,
                 width=6, anchor="w").grid(row=fila, column=0, sticky="w", pady=px(3))
        tk.Entry(
            marco, textvariable=variable, width=12, font=(FUENTE, 11, "bold"), bg="#FFFFFF",
            fg=Paleta.riel, relief=tk.FLAT, bd=0, justify="center", insertbackground=Paleta.riel,
            highlightthickness=2, highlightbackground=Paleta.riel, highlightcolor=Paleta.acento_vivo,
        ).grid(row=fila, column=1, sticky="w", ipady=px(3), pady=px(3))

    @staticmethod
    def _titulo_riel(parent, texto):
        tk.Label(parent, text=texto, font=(FUENTE, 8, "bold"), fg=Paleta.riel_texto,
                 bg=Paleta.riel, anchor="w").pack(fill="x")

    def _separador_riel(self, riel):
        tk.Frame(riel, bg=Paleta.riel_suave, height=1).pack(fill="x", padx=self._px(22), pady=self._px(7))

    # --- cards --------------------------------------------------------
    def _tarjeta(self, parent, titulo, *, chip=None, franja=Paleta.acento):
        """A white card with a coloured top edge, a title and (optionally) the state chip of its process."""
        px = self._px
        marco = tk.Frame(parent, bg=Paleta.tarjeta, highlightthickness=1,
                         highlightbackground=Paleta.borde, highlightcolor=Paleta.borde)
        tk.Frame(marco, bg=franja, height=px(3)).pack(fill="x")
        interior = tk.Frame(marco, bg=Paleta.tarjeta, padx=px(16), pady=px(12))
        interior.pack(fill="both", expand=True)
        cabecera = tk.Frame(interior, bg=Paleta.tarjeta)
        cabecera.pack(fill="x")
        tk.Label(cabecera, text=titulo, font=(FUENTE, 11, "bold"), fg=Paleta.texto,
                 bg=Paleta.tarjeta).pack(side=tk.LEFT)
        if chip:
            etiqueta = tk.Label(cabecera, font=(FUENTE, 8, "bold"), padx=px(9), pady=px(2))
            etiqueta.pack(side=tk.RIGHT)
            self._chips_paso[chip] = etiqueta
        return marco, interior, cabecera

    def _segmentos(self, parent, variable, opciones, comando, *, fondo, oscuro=False):
        """A row of pills bound to `variable` → {value: Segmento}."""
        px = self._px
        segmentos = {}
        for valor, texto in opciones:
            segmento = Segmento(parent, texto, variable, valor, comando, fondo, oscuro=oscuro,
                                alto=px(28), radio=px(9), margen=px(26))
            segmento.pack(side=tk.LEFT, padx=(0, px(6)))
            segmentos[valor] = segmento
        return segmentos

    def _boton(self, parent, texto, comando, *, estilo="solido", color=Paleta.acento,
               alto=42, radio=10, ancho=None, margen=32, fuente=(FUENTE, 10, "bold")):
        px = self._px
        return BotonRedondeado(
            parent, texto, comando, color, Paleta.tarjeta, estilo=estilo, fuente=fuente,
            alto=px(alto), radio=px(radio), ancho=px(ancho) if ancho else None, margen=px(margen),
        )

    @staticmethod
    def _leyenda(parent, lineas, ancho_envoltura):
        """Italic grey legend that wraps to its card and keeps `lineas` lines of room, so it never moves things."""
        etiqueta = tk.Label(parent, text="", font=(FUENTE, 9, "italic"), fg=Paleta.suave,
                            bg=Paleta.tarjeta, anchor="nw", justify="left", height=lineas,
                            wraplength=ancho_envoltura)
        etiqueta.bind("<Configure>", lambda e, t=etiqueta: t.configure(wraplength=max(e.width - 4, 120)))
        return etiqueta

    def _build_tarjeta_descarga(self, parent):
        px = self._px
        marco, interior, _ = self._tarjeta(parent, "Descarga de SAP", chip="descarga")

        cuerpo = tk.Frame(interior, bg=Paleta.tarjeta)
        cuerpo.pack(fill="x", pady=(px(10), 0))
        cuerpo.columnconfigure(1, weight=1)      # column 0: fixed width (_igualar_columna_izquierda)
        self._cuerpo_descarga = cuerpo

        # Left: the files this download will leave in the folder (always room for the longer list) and,
        # right under it and as wide as it, the button that downloads all of them.
        izquierda = tk.Frame(cuerpo, bg=Paleta.tarjeta)
        izquierda.grid(row=0, column=0, sticky="nsew", padx=(0, px(16)))
        self._izquierda_descarga = izquierda
        self._titulo_campo(izquierda, "Archivos que se generarán")
        caja = tk.Frame(izquierda, bg=Paleta.campo, highlightthickness=1,
                        highlightbackground=Paleta.borde)
        caja.pack(fill="x")
        self._caja_archivos = caja
        lineas = len(self._lineas_archivos())
        tk.Label(caja, textvariable=self.filename_var, font=(FUENTE_MONO, 9), fg=Paleta.texto,
                 bg=Paleta.campo, justify="left", anchor="nw", height=lineas,
                 padx=px(10), pady=px(6)).pack(fill="x")
        n = sum(len(transacciones) for transacciones in CORRIDAS_DESCARGA.values())
        self.download_btn = self._boton(izquierda, f"📸 Descargar todo ({n} transacciones)",
                                        self._handle_todos)
        self.download_btn.pack(fill="x", pady=(px(12), 0))

        # Right, ONE column: the single-transaction downloads, the parallel switch and, at the bottom
        # (level with the main button), the full run: download + consolidation.
        derecha = tk.Frame(cuerpo, bg=Paleta.tarjeta)
        derecha.grid(row=0, column=1, sticky="nsew")
        self._titulo_campo(derecha, "Descargas sueltas (una transacción)")
        botones = (
            ("download_gallo_btn",   "🐓 Gallo",       self._handle_gallo),
            ("download_monivoi_btn", "📊 Monivoi",     self._handle_monivoi),
            ("download_dsd_btn",     "🛒 Venta DSD",   self._handle_dsd),
            ("download_eiap_btn",    "📈 Venta EIAP",  self._handle_eiap),
        )
        for atributo, texto, comando in botones:
            boton = self._boton(derecha, texto, comando, estilo="contorno", alto=32, radio=8,
                                margen=24, fuente=(FUENTE, 9, "bold"))
            boton.pack(fill="x", pady=(0, px(6)))
            setattr(self, atributo, boton)

        self.paralelo_chk = Interruptor(
            derecha, "⚡ En paralelo, por pasadas", self.paralelo_var, Paleta.tarjeta,
            alto=px(28), pista=(px(42), px(22)), brecha=px(10),
        )
        self.paralelo_chk.pack(anchor="w", pady=(px(10), 0))
        tk.Label(derecha, text="Usa las ventanas de SAP que ya tengas abiertas.", font=(FUENTE, 8),
                 fg=Paleta.suave, bg=Paleta.tarjeta, anchor="w").pack(anchor="w")

        self.corrida_btn = self._boton(derecha, "🚀 Corrida completa", self._handle_corrida_completa,
                                       color=Paleta.riel)
        self.corrida_btn.pack(side=tk.BOTTOM, fill="x")
        return marco

    def _build_tarjeta_consolidacion(self, parent):
        px = self._px
        marco, interior, _ = self._tarjeta(parent, "Consolidación en Excel", chip="consolidacion")
        self.leyenda_consolidacion = self._leyenda(interior, 2, px(360))
        self.leyenda_consolidacion.config(text=self.LEYENDA_CONSOLIDACION)
        self.leyenda_consolidacion.pack(fill="x", pady=(px(10), 0))

        self.consolidar_btn = self._boton(interior, "🧩 Consolidar...!", self._handle_consolidar)
        self.consolidar_btn.pack(side=tk.BOTTOM, fill="x", pady=(px(10), 0))
        return marco

    def _build_tarjeta_carpetas(self, parent):
        px = self._px
        marco, interior, cabecera = self._tarjeta(
            parent, "Carpetas y parámetros", franja=Paleta.riel_brillo)
        # Company code, next to the title.
        entrada = self._entrada(cabecera, self.sociedad_var, width=6, justify="center",
                                fuente=(FUENTE, 10, "bold"))
        entrada.pack(side=tk.RIGHT, ipady=px(2))
        tk.Label(cabecera, text="Sociedad", font=(FUENTE, 9), fg=Paleta.suave,
                 bg=Paleta.tarjeta).pack(side=tk.RIGHT, padx=(0, px(8)))

        self._campo_carpeta(interior, "Carpeta de descarga (donde SAP deja los archivos)",
                            self.path_var, self._elegir_carpeta)
        self._campo_carpeta(interior, "Carpeta de consolidación (donde se guarda el Excel)",
                            self.consolidacion_path_var, self._elegir_carpeta_consolidacion)
        return marco

    def _campo_carpeta(self, parent, etiqueta, variable, comando):
        px = self._px
        self._titulo_campo(parent, etiqueta, pady=(px(10), px(3)))
        fila = tk.Frame(parent, bg=Paleta.tarjeta)
        fila.pack(fill="x")
        self._entrada(fila, variable, width=10, fuente=(FUENTE, 9)).pack(
            side=tk.LEFT, fill="x", expand=True, padx=(0, px(8)), ipady=px(5))
        self._boton(fila, "Examinar...", comando, estilo="sutil", color=Paleta.texto, alto=30,
                    radio=8, ancho=96, fuente=(FUENTE, 9, "bold")).pack(side=tk.LEFT)

    @staticmethod
    def _entrada(parent, variable, *, width, fuente, justify="left"):
        return tk.Entry(
            parent, textvariable=variable, width=width, font=fuente, bg=Paleta.campo,
            fg=Paleta.texto, relief=tk.FLAT, bd=0, justify=justify, insertbackground=Paleta.texto,
            highlightthickness=1, highlightbackground=Paleta.borde, highlightcolor=Paleta.acento,
        )

    def _titulo_campo(self, parent, texto, pady=None):
        """Small caption above a field, so each one says what it is."""
        tk.Label(parent, text=texto, font=(FUENTE, 8, "bold"), fg=Paleta.suave,
                 bg=Paleta.tarjeta, anchor="w").pack(fill="x", pady=pady if pady is not None else (0, self._px(3)))

    # --- status strip -------------------------------------------------
    def _build_pie(self, parent):
        px = self._px
        pie = tk.Frame(parent, bg=Paleta.fondo)
        pie.pack(side=tk.BOTTOM, fill="x", pady=(px(12), 0))
        fila = tk.Frame(pie, bg=Paleta.fondo)
        fila.pack(fill="x")
        self.punto = tk.Canvas(fila, width=px(12), height=px(12), bg=Paleta.fondo, highlightthickness=0)
        self.punto.pack(side=tk.LEFT, padx=(0, px(8)))
        # Two lines of room: a long message wraps instead of pushing the cards.
        self.status_lbl = tk.Label(
            fila, textvariable=self.status_var, font=(FUENTE, 10), fg=Paleta.texto,
            bg=Paleta.fondo, anchor="w", justify="left", height=2, wraplength=px(560),
        )
        self.status_lbl.pack(side=tk.LEFT, fill="x", expand=True)
        fila.bind("<Configure>", lambda e: self.status_lbl.configure(wraplength=max(e.width - px(40), 160)))
        self.progreso = BarraProgreso(pie, Paleta.fondo, alto=px(8))
        self.progreso.pack(fill="x", pady=(px(4), 0))
        self._pintar_punto(Paleta.verde)

    def _pintar_punto(self, color):
        px = self._px
        self.punto.delete("all")
        self.punto.create_oval(1, 1, px(11), px(11), fill=color, outline=color)

    # ------------------------------------------------------------------
    # Steps of the process (the controller reports them)
    # ------------------------------------------------------------------
    def marcar_paso(self, paso, estado):
        """
        Set the state of one step of the rail: paso in ('descarga', 'consolidacion'), estado in
        ('pendiente', 'en_curso', 'listo', 'parcial', 'error'). A step that
        starts resets the bar; one that ends fills it (green done, amber with warnings, red on error).
        Main thread only: it touches Tk.
        """
        if paso not in self._estado_paso:
            raise ValueError(f"Paso desconocido: {paso!r}")
        if estado not in ESTADOS or estado == "siguiente":
            raise ValueError(f"Estado desconocido: {estado!r}")
        self._estado_paso[paso] = estado
        if estado == "en_curso":
            self._barra_cerrada = False
            self._ultimo_mensaje = None
            self.progreso.establecer(4, Paleta.acento)
        elif estado == "listo":
            self._cerrar_barra(100, Paleta.verde)
            self._pintar_punto(Paleta.verde)
        elif estado == "parcial":
            self._cerrar_barra(100, Paleta.ambar)
            self._pintar_punto(Paleta.ambar)
        elif estado == "error":
            self._cerrar_barra(max(self.progreso.valor, 8), Paleta.rojo)
            self._pintar_punto(Paleta.rojo)
        self._refrescar_pasos()
        self.root.update_idletasks()

    def reiniciar_pasos(self):
        """Every step back to 'pendiente' and the bar empty: a new period starts a new process."""
        if self._ocupado:
            return
        for clave in self._estado_paso:
            self._estado_paso[clave] = "pendiente"
        self._barra_cerrada = False
        self.progreso.establecer(0, Paleta.acento)
        self._refrescar_pasos()

    def _cerrar_barra(self, valor, color):
        self._barra_cerrada = True
        self.progreso.establecer(valor, color)

    def _estados_mostrados(self):
        """
        State of each step as shown. 'Siguiente' is not stored: it is the first pending step after the
        last one that was already run (or the first step, if none was), and only while nothing runs and
        the last step run did not fail (a failed step is retried: the error is the cue, nothing "moves on").
        """
        estados = dict(self._estado_paso)
        orden = [clave for clave, *_ in PASOS]
        if "en_curso" in estados.values():
            return estados
        ultimo = max((i for i, c in enumerate(orden)
                      if estados[c] in ("listo", "parcial", "error")), default=-1)
        if ultimo >= 0 and estados[orden[ultimo]] == "error":
            return estados
        for clave in orden[ultimo + 1:]:
            if estados[clave] == "pendiente":
                estados[clave] = "siguiente"
                break
        return estados

    def _refrescar_pasos(self):
        px = self._px
        estados = self._estados_mostrados()
        for clave, estado in estados.items():
            circulo, etiqueta, numero = self._riel_paso[clave]
            anillo, relleno, simbolo_color, texto_color = ESTADOS_RIEL[estado]
            circulo.delete("all")
            circulo.create_oval(2, 2, px(26), px(26), outline=anillo, width=2, fill=relleno or Paleta.riel)
            simbolo = {"listo": "✓", "error": "✕", "parcial": "!"}.get(estado, numero)
            circulo.create_text(px(14), px(14), text=simbolo, fill=simbolo_color, font=(FUENTE, 10, "bold"))
            etiqueta.configure(text=ESTADOS[estado][0], fg=texto_color)
        orden = [clave for clave, *_ in PASOS]
        for posicion, conector in enumerate(self._conectores):
            # The line below a finished step lights up: the process moved on from it.
            conector.configure(bg=ESTADOS_RIEL[estados[orden[posicion]]][0]
                               if estados[orden[posicion]] in ("listo", "parcial") else Paleta.riel_suave)
        self._refrescar_chips(estados)

    def _refrescar_chips(self, estados=None):
        """The chip of each card shows the state of its step."""
        estados = estados or self._estados_mostrados()
        for tarjeta, chip in self._chips_paso.items():
            etiqueta, texto, fondo, simbolo = ESTADOS[estados[tarjeta]]
            chip.configure(text=f"{simbolo}  {etiqueta}", fg=texto, bg=fondo)

    # ------------------------------------------------------------------
    # Folder pickers
    # ------------------------------------------------------------------
    def _elegir_carpeta(self):
        """Open the file browser so the stakeholder can pick the download folder."""
        carpeta = filedialog.askdirectory(
            title="Elige la carpeta donde se descargarán los archivos",
            initialdir=self.path_var.get() or os.path.expanduser("~"),
        )
        if carpeta:
            self.path_var.set(os.path.normpath(carpeta))
            self._guardar_config()

    def _elegir_carpeta_consolidacion(self):
        """File browser to pick where the consolidated workbook will be saved."""
        carpeta = filedialog.askdirectory(
            title="Elige la carpeta donde se guardará el Excel consolidado",
            initialdir=self.consolidacion_path_var.get() or os.path.expanduser("~"),
        )
        if carpeta:
            self.consolidacion_path_var.set(os.path.normpath(carpeta))
            self._guardar_config()

    # ------------------------------------------------------------------
    # File names
    # ------------------------------------------------------------------
    def _preview_nombre(self):
        try:
            if self.mode_var.get() == "single":
                return nombre_archivo_dia(self.date_single_var.get().strip())
            return nombre_archivo_rango(
                self.date_from_var.get().strip(),
                self.date_to_var.get().strip(),
            )
        except ValueError:
            return "(fecha inválida)"

    def _preview_nombre_monivoi(self):
        try:
            if self.mode_var.get() == "single":
                return nombre_archivo_dia_monivoi(self.date_single_var.get().strip())
            return nombre_archivo_rango_monivoi(
                self.date_from_var.get().strip(),
                self.date_to_var.get().strip(),
            )
        except ValueError:
            return "(fecha inválida)"

    def _preview_nombre_vbrl(self):
        try:
            if self.mode_var.get() == "single":
                return nombre_archivo_dia_vbrl(self.date_single_var.get().strip())
            return nombre_archivo_rango_vbrl(
                self.date_from_var.get().strip(),
                self.date_to_var.get().strip(),
            )
        except ValueError:
            return "(fecha inválida)"

    def _preview_nombre_fb03(self):
        try:
            if self.mode_var.get() == "single":
                return nombre_archivo_dia_fb03(self.date_single_var.get().strip())
            return nombre_archivo_rango_fb03(
                self.date_from_var.get().strip(),
                self.date_to_var.get().strip(),
            )
        except ValueError:
            return "(fecha inválida)"

    def _preview_nombre_2da(self, clave):
        """File of a 2nd-run transaction ('vbrp', 'kna1'...), named with the GUI dates."""
        try:
            if self.mode_var.get() == "single":
                return nombre_archivo_dia_2da(clave, self.date_single_var.get().strip())
            return nombre_archivo_rango_2da(
                clave,
                self.date_from_var.get().strip(),
                self.date_to_var.get().strip(),
            )
        except ValueError:
            return "(fecha inválida)"

    def _preview_nombre_venta(self, nombre_archivo):
        # The sales buttons ('Venta DSD', 'Venta EIAP') only work with a range.
        if self.mode_var.get() == "single":
            return "(sólo con 'Varios días')"
        try:
            return nombre_archivo(
                self.date_from_var.get().strip(),
                self.date_to_var.get().strip(),
            )
        except ValueError:
            return "(fecha inválida)"

    def _preview_nombre_dsd(self):
        return self._preview_nombre_venta(nombre_archivo_rango_dsd)

    def _preview_nombre_eiap(self):
        return self._preview_nombre_venta(nombre_archivo_rango_eiap)

    def _lineas_archivos(self):
        """[(name, file)] of what the main download button leaves in the folder: its 1st run and its 2nd
        run (CORRIDAS_DESCARGA: the same lists the controller runs), then 'Venta DSD' and 'Venta EIAP',
        whose buttons sit next to it."""
        nombres_1a = {
            "gallo":   self._preview_nombre,
            "monivoi": self._preview_nombre_monivoi,
            "vbrl":    self._preview_nombre_vbrl,
            "fb03":    self._preview_nombre_fb03,
        }
        return (
            [(nombre, nombres_1a[clave]()) for clave, nombre in CORRIDAS_DESCARGA[1]]
            + [(nombre, self._preview_nombre_2da(clave)) for clave, nombre in CORRIDAS_DESCARGA[2]]
            + [("Venta DSD", self._preview_nombre_dsd()), ("Venta EIAP", self._preview_nombre_eiap())]
        )

    @staticmethod
    def _texto_archivos(lineas):
        """The lines of `_lineas_archivos`, names aligned in a column."""
        ancho = max(len(nombre) for nombre, _ in lineas)
        return "\n".join(f"{nombre:<{ancho}}  {archivo}" for nombre, archivo in lineas)

    def _lineas_con_fechas_largas(self, modo):
        """`_lineas_archivos()` with the LONGEST names of period mode `modo` (two-digit days), whatever the
        entries say now. Only the four period variables are swapped for fixed stand-ins (nothing is set, so no
        trace fires) and they are always put back."""
        reales = self.mode_var, self.date_single_var, self.date_from_var, self.date_to_var
        self.mode_var, self.date_single_var = _Fijo(modo), _Fijo(self.FECHA_LARGA_DESDE)
        self.date_from_var, self.date_to_var = _Fijo(self.FECHA_LARGA_DESDE), _Fijo(self.FECHA_LARGA_HASTA)
        try:
            return self._lineas_archivos()
        finally:
            self.mode_var, self.date_single_var, self.date_from_var, self.date_to_var = reales

    def _update_filename_preview(self):
        """The files of the main download, names aligned in a column."""
        self.filename_var.set(self._texto_archivos(self._lineas_archivos()))

    # ------------------------------------------------------------------
    # Period (rail)
    # ------------------------------------------------------------------
    def _on_mode_change(self):
        if self.mode_var.get() == "single":
            self.range_frame.pack_forget()
            self.single_frame.pack(fill="x")
        else:
            self.single_frame.pack_forget()
            self.range_frame.pack(fill="x")
        self._al_cambiar_periodo()

    def _al_cambiar_periodo(self):
        """The period changed: file names and the hint follow, and the steps of the previous period no
        longer describe this one (their files are other files), so the rail starts over."""
        self._update_filename_preview()
        self._actualizar_resumen_periodo()
        if self._ocupado:
            return
        if any(estado != "pendiente" for estado in self._estado_paso.values()):
            self.reiniciar_pasos()
            self.set_status("✓ Listo para comenzar")

    def _pintar_resumen_periodo(self, color):
        for etiqueta in self._resumenes_periodo:
            etiqueta.configure(fg=color)

    def _actualizar_resumen_periodo(self):
        """Live hint under the dates: how many days the run covers, or what is wrong with them."""
        try:
            if self.mode_var.get() == "single":
                dia = parse_fecha(self.date_single_var.get())
                texto = f"✓ 1 día: {dia.strftime('%d.%m.%Y')}"
            else:
                desde = parse_fecha(self.date_from_var.get())
                hasta = parse_fecha(self.date_to_var.get())
                if desde > hasta:
                    raise ValueError("desde > hasta")
                dias = (hasta - desde).days + 1
                texto = (f"✓ {dias} día(s)\n{desde.strftime('%d.%m.%Y')} → {hasta.strftime('%d.%m.%Y')}")
            self.resumen_periodo_var.set(texto)
            self._pintar_resumen_periodo("#FFFFFF")
        except ValueError:
            self.resumen_periodo_var.set("⚠ Revisa las fechas\nDD.MM.AAAA · Desde ≤ Hasta")
            self._pintar_resumen_periodo("#FFB4B4")

    # ------------------------------------------------------------------
    # Validation
    # ------------------------------------------------------------------
    def validate_dates(self):
        try:
            if self.mode_var.get() == "single":
                datetime.strptime(self.date_single_var.get().strip(), "%d.%m.%Y")
            else:
                d1 = datetime.strptime(self.date_from_var.get().strip(), "%d.%m.%Y")
                d2 = datetime.strptime(self.date_to_var.get().strip(), "%d.%m.%Y")
                if d1 > d2:
                    messagebox.showerror("Error", "La fecha desde no puede ser mayor a la fecha hasta")
                    return False
        except ValueError:
            messagebox.showerror("Error", "Formato de fecha inválido. Use DD.MM.YYYY")
            return False

        ruta = self.path_var.get().strip()
        if not ruta:
            messagebox.showerror("Error", "Indica una carpeta de descarga válida.")
            return False
        try:
            os.makedirs(ruta, exist_ok=True)
        except OSError as e:
            messagebox.showerror(
                "Error", f"No se pudo usar la carpeta de descarga:\n{ruta}\n\n{e}"
            )
            return False

        return True

    # ------------------------------------------------------------------
    # Accessors used by the controller
    # ------------------------------------------------------------------
    def get_config(self, transaccion="gallo"):
        mode = self.mode_var.get()
        # Any other value falls back to the Gallo name.
        nombres = {
            "monivoi": self._preview_nombre_monivoi,
            "vbrl":    self._preview_nombre_vbrl,
            "fb03":    self._preview_nombre_fb03,
            "dsd":     self._preview_nombre_dsd,
            "eiap":    self._preview_nombre_eiap,
        }
        if transaccion in ARCHIVOS_SEGUNDA_CORRIDA:
            filename = self._preview_nombre_2da(transaccion)
        else:
            filename = nombres.get(transaccion, self._preview_nombre)()
        input_path = self.path_var.get().strip()
        os.makedirs(input_path, exist_ok=True)
        self._guardar_config()
        cfg = {
            "mode":       mode,
            "sociedad":   self.sociedad_var.get().strip().upper() or "MX21",
            "input_path": input_path,
            "filename":   filename,
        }
        if mode == "single":
            cfg["fecha"] = self.date_single_var.get().strip()
        else:
            cfg["date_from"] = self.date_from_var.get().strip()
            cfg["date_to"]   = self.date_to_var.get().strip()
        return cfg

    def get_consolidacion_path(self):
        """Folder where the consolidated workbook will be saved (created if missing)."""
        ruta = self.consolidacion_path_var.get().strip()
        os.makedirs(ruta, exist_ok=True)
        self._guardar_config()
        return ruta

    def usar_paralelo(self):
        """True if the 'en paralelo' switch is on (read by the controller)."""
        return bool(self.paralelo_var.get())

    # ------------------------------------------------------------------
    # Status strip
    # ------------------------------------------------------------------
    @staticmethod
    def _tono_de(mensaje):
        """'ok' / 'error' / 'aviso' from the leading symbol of a status message; anything else is work in progress."""
        inicio = mensaje.lstrip()[:1]
        if inicio in ("✅", "✓"):
            return "ok"
        if inicio == "❌":
            return "error"
        if inicio == "⚠":
            return "aviso"
        return "trabajando"

    @staticmethod
    def _sin_cronometro(mensaje):
        """The message without a trailing '(m:ss)' clock: the parallel runs repeat their status every
        second with a new clock, and that alone is not progress."""
        cuerpo, parentesis, cola = mensaje.rstrip().rpartition("(")
        if parentesis and cola.endswith(")") and cola[:-1].replace(":", "").isdigit():
            return cuerpo.rstrip()
        return mensaje

    _COLOR_TONO = {"ok": Paleta.verde, "error": Paleta.rojo, "aviso": Paleta.ambar, "trabajando": Paleta.acento}

    def set_status(self, message):
        """
        Show a progress / result line. The colour of the dot follows the leading symbol of the message
        (✅ green, ❌ red, ⚠ amber, anything else teal). While a step runs every NEW message (a clock that
        only ticks does not count) also advances the bar a fraction of what is left: it never reaches the
        end by itself, since the length of a step is not known and the steps run on the UI thread, where a
        timer-driven animation would freeze.
        """
        self.status_var.set(message)
        self._tono = self._tono_de(message)
        self._pintar_punto(self._COLOR_TONO[self._tono])
        nuevo = self._sin_cronometro(message)
        if self._ocupado and not self._barra_cerrada and nuevo != self._ultimo_mensaje:
            self.progreso.establecer(self.progreso.valor + (92.0 - self.progreso.valor) * 0.08,
                                     Paleta.acento)
        self._ultimo_mensaje = nuevo
        self.root.update_idletasks()

    # ------------------------------------------------------------------
    # Button state
    # ------------------------------------------------------------------
    def _botones_accion(self):
        """(button, enabled color) for EVERY button that starts a process."""
        return tuple(
            (boton, boton.color) for boton in (
                self.download_gallo_btn, self.download_monivoi_btn, self.download_dsd_btn,
                self.download_eiap_btn, self.download_btn, self.consolidar_btn, self.corrida_btn,
            )
        )

    def disable_buttons(self):
        self._ocupado = True
        self._barra_cerrada = False
        self._tono = "trabajando"      # a process that reports nothing must not end with a "done" bar
        self._ultimo_mensaje = None
        for boton, _ in self._botones_accion():
            boton.habilitar(False)
        self.paralelo_chk.habilitar(False)
        for radio in self.radios_modo.values():        # the period cannot change while one is running
            radio.habilitar(False)
        self.progreso.establecer(3, Paleta.acento)
        self.root.configure(cursor="watch")
        self.root.update_idletasks()

    def enable_buttons(self):
        self._ocupado = False
        for boton, _ in self._botones_accion():
            boton.habilitar(True)
        self.paralelo_chk.habilitar(True)
        for radio in self.radios_modo.values():
            radio.habilitar(True)
        if not self._barra_cerrada:
            self._cerrar_barra_por_mensaje()
        self.root.configure(cursor="")
        self.root.update_idletasks()

    def _cerrar_barra_por_mensaje(self):
        """A process that never reported a step (the single downloads) still ends with a full bar when
        its last message was a result; work that just stopped leaves the bar empty."""
        if self._tono == "ok":
            self._cerrar_barra(100, Paleta.verde)
        elif self._tono == "aviso":
            self._cerrar_barra(100, Paleta.ambar)
        elif self._tono == "error":
            self._cerrar_barra(max(self.progreso.valor, 8), Paleta.rojo)
        else:
            self.progreso.establecer(0, Paleta.acento)

    # ------------------------------------------------------------------
    # Handlers
    # ------------------------------------------------------------------
    def _disparar(self, callback, transaccion="gallo"):
        """Validate inputs and call the controller callback (or warn if unwired)."""
        if not self.validate_dates():
            return
        if callback:
            callback()
        else:
            messagebox.showinfo(
                "Info", f"Funcionalidad no conectada\n\n{self.get_config(transaccion)}"
            )

    def _handle_gallo(self):
        self._disparar(self.on_download_gallo, "gallo")

    def _handle_monivoi(self):
        self._disparar(self.on_download_monivoi, "monivoi")

    def _handle_dsd(self):
        self._disparar(self.on_download_dsd, "dsd")

    def _handle_eiap(self):
        self._disparar(self.on_download_eiap, "eiap")

    def _handle_todos(self):
        self._disparar(self.on_download_todos, "gallo")

    def _handle_consolidar(self):
        self._disparar(self.on_consolidar, "gallo")

    def _handle_corrida_completa(self):
        self._disparar(self.on_corrida_completa, "gallo")


def main():
    root = tk.Tk()
    ValidacionFacturaGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
