"""Validación Factura Global — GUI (tkinter): dates, folders and buttons only. The controller
wires the on_* callbacks; `python interfaz_GUI.py` opens the window on its own.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import os
import sys
import json


from utils import (
    nombre_archivo_dia,          nombre_archivo_rango,
    nombre_archivo_dia_monivoi,  nombre_archivo_rango_monivoi,
    nombre_archivo_dia_vbrl,     nombre_archivo_rango_vbrl,
    nombre_archivo_dia_fb03,     nombre_archivo_rango_fb03,
    nombre_archivo_rango_eiap,
    ARCHIVOS_SEGUNDA_CORRIDA,    CORRIDAS_DESCARGA,
    nombre_archivo_dia_2da,      nombre_archivo_rango_2da,
)


def _ruta_config():
    """config.json (last download/consolidation folders) next to the .exe (PyInstaller) or to this .py."""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "config.json")


class ValidacionFacturaGUI:
    # Design size = MINIMUM size: _ajustar_alto_ventana grows it when the content needs more room.
    ANCHO_BASE = 580
    ALTO_BASE  = 690

    # The two runs: radio buttons for the downloads and for the consolidations, each with a legend.
    LEYENDAS_DESCARGA = {
        1: " · ".join(nombre for _, nombre in CORRIDAS_DESCARGA[1]),
        2: " · ".join(nombre for _, nombre in CORRIDAS_DESCARGA[2]),
    }
    LEYENDAS_CONSOLIDACION = {
        1: "Gallo + Monivoi + VBRL + FB03 → Excel consolidado",
        2: "Por definir",
    }
    # The 2nd consolidation is not defined yet: its radio button stays disabled until it is.
    CONSOLIDACION_2_DISPONIBLE = False

    def __init__(self, root):
        self.root = root
        self.root.title("VALIDACIÓN FACTURA GLOBAL")
        self.root.geometry(f"{self.ANCHO_BASE}x{self.ALTO_BASE}")
        self.root.resizable(False, False)

        self.bg_color        = "#FFFFFF"
        self.primary_color   = "#00094F"
        self.secondary_color = "#333333"
        self.light_gray      = "#F5F5F5"
        self.border_color    = "#E0E0E0"
        self.success_color   = "#28A745"

        self.root.configure(bg=self.bg_color)

        # Callbacks the controller wires up
        self.on_download_gallo   = None
        self.on_download_monivoi = None
        self.on_download_eiap    = None
        self.on_download_todos   = None
        self.on_consolidar       = None

        self.mode_var = tk.StringVar(value="single")

        # Parallel run by passes (opt-in): it always starts unchecked (sequential is the safe default).
        self.paralelo_var = tk.BooleanVar(value=False)

        # Which run the main buttons execute (1a / 2a); always the 1st when the window opens.
        self.corrida_descarga_var = tk.IntVar(value=1)
        self.corrida_consolidacion_var = tk.IntVar(value=1)

        # Download folder: the user's Documents path by default; it can be changed (e.g. a synced SharePoint
        # folder) and is remembered in config.json, like the consolidation folder.
        self._config_file = _ruta_config()
        cfg = self._cargar_config()
        user_profile = os.environ.get("USERPROFILE") or os.path.expanduser("~")
        base_proyecto = os.path.join(
            user_profile, "Documents", "Validacion Factura Global", "src"
        )
        ruta_input_default  = os.path.join(base_proyecto, "Input")
        ruta_output_default = os.path.join(base_proyecto, "Output")

        self.path_var = tk.StringVar(
            value=(cfg.get("ruta_descarga") or "").strip() or ruta_input_default
        )
        self.consolidacion_path_var = tk.StringVar(
            value=(cfg.get("ruta_consolidacion") or "").strip() or ruta_output_default
        )
        os.makedirs(ruta_input_default, exist_ok=True)

        self._build_ui()
        self._ajustar_alto_ventana()
        self._on_mode_change()

    # --- config.json (last folders used) ---
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

    # --- Interface: header, Notebook with two tabs, status bar always visible below ---
    def _build_ui(self):
        main = tk.Frame(self.root, bg=self.bg_color, padx=24, pady=18)
        main.pack(fill="both", expand=True)

        # --- Header (fixed) ---
        tk.Label(
            main, text="VALIDACIÓN FACTURA GLOBAL",
            font=("Segoe UI", 16, "bold"),
            bg=self.bg_color, fg=self.primary_color,
        ).pack(pady=(0, 4))
        tk.Label(
            main, text="Automatización de Descarga y Consolidación - Contabilidad vs Monitoreo de Facturas",
            font=("Segoe UI", 10),
            bg=self.bg_color, fg=self.secondary_color,
        ).pack(pady=(0, 14))

        # --- Notebook with the two tabs ---
        style = ttk.Style()
        try:
            style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"),
                            padding=(16, 8))
        except Exception:
            pass

        notebook = ttk.Notebook(main)
        notebook.pack(fill="both", expand=True)

        tab_descarga = tk.Frame(notebook, bg=self.bg_color)
        tab_config   = tk.Frame(notebook, bg=self.bg_color)
        notebook.add(tab_descarga, text="  Descarga y Consolidación  ")
        notebook.add(tab_config,   text="  Configuración  ")

        self._build_tab_descarga(tab_descarga)
        self._build_tab_config(tab_config)

        # --- Status bar (fixed, below the Notebook) ---
        self.status_var = tk.StringVar(value="✓ Listo para comenzar")
        tk.Label(
            main, textvariable=self.status_var, font=("Segoe UI", 10),
            bg=self.bg_color, fg="#666666",
        ).pack(pady=(12, 0))

    # --- Tab 1: Descarga y Consolidación ---
    def _build_tab_descarga(self, parent):
        cont = tk.Frame(parent, bg=self.bg_color, padx=16, pady=16)
        cont.pack(fill="both", expand=True)

        # --- Mode selector ---
        mode_frame = tk.LabelFrame(
            cont, text="  Tipo de intervalo  ",
            font=("Segoe UI", 10, "bold"),
            bg=self.bg_color, fg=self.primary_color,
            bd=1, relief=tk.SOLID, padx=15, pady=6,
        )
        mode_frame.pack(fill="x", pady=(0, 10))

        for text, value in [("Un solo día", "single"), ("Varios días", "range")]:
            tk.Radiobutton(
                mode_frame, text=text, variable=self.mode_var, value=value,
                font=("Segoe UI", 9), bg=self.bg_color, fg=self.primary_color,
                activebackground=self.bg_color, selectcolor=self.light_gray,
                cursor="hand2", command=self._on_mode_change,
            ).pack(side=tk.LEFT, padx=(0, 16))

        # --- Date container (switches between single/range) ---
        self.date_container = tk.Frame(cont, bg=self.bg_color)
        self.date_container.pack(fill="x", pady=(0, 10))

        hoy = datetime.now().strftime("%d.%m.%Y")

        # "Single day" frame
        self.single_frame = tk.LabelFrame(
            self.date_container, text="  Fecha  ",
            font=("Segoe UI", 10, "bold"),
            bg=self.bg_color, fg=self.primary_color,
            bd=1, relief=tk.SOLID, padx=15, pady=6,
        )
        tk.Label(
            self.single_frame, text="Fecha:", font=("Segoe UI", 9),
            bg=self.bg_color, fg=self.secondary_color,
        ).grid(row=0, column=0, sticky="w", pady=8)
        self.date_single_var = tk.StringVar(value=hoy)
        tk.Entry(
            self.single_frame, textvariable=self.date_single_var, width=15,
            font=("Segoe UI", 10), bg=self.light_gray, fg=self.primary_color,
            relief=tk.FLAT, bd=1, highlightthickness=1,
            highlightbackground=self.border_color, highlightcolor=self.primary_color,
        ).grid(row=0, column=1, padx=10, pady=8, sticky="w")
        tk.Label(
            self.single_frame, text="(DD.MM.YYYY)", font=("Segoe UI", 8),
            bg=self.bg_color, fg="#999999",
        ).grid(row=0, column=2, sticky="w")

        # "Several days" frame
        self.range_frame = tk.LabelFrame(
            self.date_container, text="  Intervalo de Tiempo  ",
            font=("Segoe UI", 10, "bold"),
            bg=self.bg_color, fg=self.primary_color,
            bd=1, relief=tk.SOLID, padx=15, pady=6,
        )
        # Both dates on ONE row (Desde ... Hasta ...): the window stays short enough for a laptop.
        for col, (label, attr) in enumerate([
            ("Desde:", "date_from_var"),
            ("Hasta:", "date_to_var"),
        ]):
            tk.Label(
                self.range_frame, text=label, font=("Segoe UI", 9),
                bg=self.bg_color, fg=self.secondary_color,
            ).grid(row=0, column=2 * col, sticky="w", pady=8)
            var = tk.StringVar(value=hoy)
            setattr(self, attr, var)
            tk.Entry(
                self.range_frame, textvariable=var, width=12,
                font=("Segoe UI", 10), bg=self.light_gray, fg=self.primary_color,
                relief=tk.FLAT, bd=1, highlightthickness=1,
                highlightbackground=self.border_color, highlightcolor=self.primary_color,
            ).grid(row=0, column=2 * col + 1, padx=(8, 16), pady=8, sticky="w")
        tk.Label(
            self.range_frame, text="(DD.MM.YYYY)", font=("Segoe UI", 8),
            bg=self.bg_color, fg="#999999",
        ).grid(row=0, column=4, sticky="w")

        for var in (self.date_single_var, self.date_from_var, self.date_to_var):
            var.trace_add("write", lambda *_: self._update_filename_preview())

        # --- Runs: downloads (left) and consolidations (right), each with its legend ---
        corridas = tk.Frame(cont, bg=self.bg_color)
        corridas.pack(fill="x", pady=(0, 10))
        corridas.columnconfigure(0, weight=1, uniform="corridas")
        corridas.columnconfigure(1, weight=1, uniform="corridas")
        self.radios_descarga, self.leyenda_descarga = self._grupo_corridas(
            corridas, 0, "  ⬇️ Descargas  ", self.corrida_descarga_var,
            self._on_corrida_descarga_change, (0, 6),
        )
        self.radios_consolidacion, self.leyenda_consolidacion = self._grupo_corridas(
            corridas, 1, "  🧩 Consolidación  ", self.corrida_consolidacion_var,
            self._on_corrida_consolidacion_change, (6, 0),
        )

        # --- Output file names (the folder is in the Configuración tab) ---
        info_frame = tk.Frame(cont, bg=self.light_gray, bd=1, relief=tk.SOLID)
        info_frame.pack(fill="x", pady=(0, 10))

        self.filename_var = tk.StringVar()
        tk.Label(
            info_frame, textvariable=self.filename_var,
            font=("Segoe UI", 8, "bold"),
            bg=self.light_gray, fg=self.primary_color, anchor="w", justify="left",
            wraplength=500,
        ).pack(fill="x", padx=12, pady=6)

        # --- Action buttons ---
        btn_frame = tk.Frame(cont, bg=self.bg_color)
        btn_frame.pack(fill="x", pady=(4, 0))

        _btn_kwargs = dict(
            font=("Segoe UI", 11, "bold"),
            bg=self.primary_color, fg=self.bg_color,
            relief=tk.RAISED, bd=2, pady=12, cursor="hand2",
            activebackground=self.secondary_color, activeforeground=self.bg_color,
        )

        # Single-transaction buttons: same blue, smaller and in italics (variants of the main one).
        _btn_individual_kwargs = dict(
            _btn_kwargs, font=("Segoe UI", 10, "bold italic"), pady=6,
        )
        fila_individual = tk.Frame(btn_frame, bg=self.bg_color)
        fila_individual.pack(fill="x", pady=(0, 8))
        # uniform= gives the three columns the same width whatever their labels.
        fila_individual.columnconfigure(0, weight=1, uniform="descargas")
        fila_individual.columnconfigure(1, weight=1, uniform="descargas")
        fila_individual.columnconfigure(2, weight=1, uniform="descargas")

        self.download_gallo_btn = tk.Button(
            fila_individual, text="🐓 Gallo",
            command=self._handle_gallo, **_btn_individual_kwargs,
        )
        self.download_gallo_btn.grid(row=0, column=0, sticky="ew", padx=(0, 4))

        self.download_monivoi_btn = tk.Button(
            fila_individual, text="📊 Monivoi",
            command=self._handle_monivoi, **_btn_individual_kwargs,
        )
        self.download_monivoi_btn.grid(row=0, column=1, sticky="ew", padx=4)

        self.download_eiap_btn = tk.Button(
            fila_individual, text="📈 Venta DSD + EIAP",
            command=self._handle_eiap, **_btn_individual_kwargs,
        )
        self.download_eiap_btn.grid(row=0, column=2, sticky="ew", padx=(4, 0))

        # Its text follows the 'Descargas' radio button (_on_corrida_descarga_change).
        self.download_btn = tk.Button(
            btn_frame, text="", command=self._handle_todos, **_btn_kwargs,
        )
        self.download_btn.pack(fill="x")

        self.paralelo_chk = tk.Checkbutton(
            btn_frame, text="⚡ En paralelo, por pasadas",
            variable=self.paralelo_var,
            font=("Segoe UI", 9), bg=self.bg_color, fg=self.primary_color,
            activebackground=self.bg_color, selectcolor=self.light_gray,
            cursor="hand2", anchor="w",
        )
        self.paralelo_chk.pack(fill="x", pady=(6, 0))

        # Green 'process data', as opposed to the blue 'bring data' buttons.
        _btn_consolidar_kwargs = dict(_btn_kwargs, bg=self.success_color)
        self.consolidar_btn = tk.Button(
            btn_frame, text="", command=self._handle_consolidar, **_btn_consolidar_kwargs,
        )
        self.consolidar_btn.pack(fill="x", pady=(10, 0))

        self._on_corrida_descarga_change()
        self._on_corrida_consolidacion_change()

    def _grupo_corridas(self, parent, columna, titulo, variable, al_cambiar, padx):
        """One LabelFrame with the '1a corrida' / '2a corrida' radio buttons and a legend below them
        (italic, gray) → ({1: radio, 2: radio}, legend label)."""
        marco = tk.LabelFrame(
            parent, text=titulo, font=("Segoe UI", 10, "bold"),
            bg=self.bg_color, fg=self.primary_color,
            bd=1, relief=tk.SOLID, padx=10, pady=6,
        )
        marco.grid(row=0, column=columna, sticky="nsew", padx=padx)
        fila = tk.Frame(marco, bg=self.bg_color)
        fila.pack(fill="x")
        radios = {}
        for valor, texto in ((1, "1a corrida"), (2, "2a corrida")):
            radios[valor] = tk.Radiobutton(
                fila, text=texto, variable=variable, value=valor,
                font=("Segoe UI", 9), bg=self.bg_color, fg=self.primary_color,
                activebackground=self.bg_color, selectcolor=self.light_gray,
                cursor="hand2", command=al_cambiar,
            )
            radios[valor].pack(side=tk.LEFT, padx=(0, 10))
        leyenda = tk.Label(
            marco, text="", font=("Segoe UI", 8, "italic"), bg=self.bg_color, fg="#666666",
            anchor="w", justify="left", wraplength=225,
        )
        leyenda.pack(fill="x", pady=(2, 0))
        return radios, leyenda

    def _on_corrida_descarga_change(self):
        """Legend, main button text and file names follow the selected download run."""
        corrida = self.corrida_descarga()
        self.leyenda_descarga.config(text=self.LEYENDAS_DESCARGA[corrida])
        n = len(CORRIDAS_DESCARGA[corrida])
        self.download_btn.config(text=f"⬇️ Descargar {corrida}a corrida ({n} transacciones)")
        self._update_filename_preview()

    def _on_corrida_consolidacion_change(self):
        corrida = self.corrida_consolidacion()
        self.leyenda_consolidacion.config(text=self.LEYENDAS_CONSOLIDACION[corrida])
        self.consolidar_btn.config(text=f"🧩 Consolidar {corrida}a corrida...!")
        self._estado_radio_consolidacion_2()

    def _estado_radio_consolidacion_2(self):
        """The 2nd consolidation radio is only usable once that consolidation exists."""
        if not self.CONSOLIDACION_2_DISPONIBLE:
            self.radios_consolidacion[2].config(state="disabled")

    def corrida_descarga(self):
        """1 or 2: the run the main download button executes (read by the controller)."""
        return 2 if self.corrida_descarga_var.get() == 2 else 1

    def corrida_consolidacion(self):
        """1 or 2: the consolidation the green button executes (read by the controller)."""
        return 2 if self.corrida_consolidacion_var.get() == 2 else 1

    # --- Tab 2: Configuración (company code and folders) ---
    def _build_tab_config(self, parent):
        cont = tk.Frame(parent, bg=self.bg_color, padx=16, pady=16)
        cont.pack(fill="both", expand=True)

        # --- Parameters (company code) ---
        param_frame = tk.LabelFrame(
            cont, text="  Parámetros  ",
            font=("Segoe UI", 10, "bold"),
            bg=self.bg_color, fg=self.primary_color,
            bd=1, relief=tk.SOLID, padx=15, pady=15,
        )
        param_frame.pack(fill="x", pady=(0, 12))

        tk.Label(
            param_frame, text="Sociedad:", font=("Segoe UI", 9),
            bg=self.bg_color, fg=self.secondary_color,
        ).grid(row=0, column=0, sticky="w", pady=6)
        self.sociedad_var = tk.StringVar(value="MX21")
        tk.Entry(
            param_frame, textvariable=self.sociedad_var, width=15,
            font=("Segoe UI", 10), bg=self.light_gray, fg=self.primary_color,
            relief=tk.FLAT, bd=1, highlightthickness=1,
            highlightbackground=self.border_color, highlightcolor=self.primary_color,
        ).grid(row=0, column=1, padx=10, pady=6, sticky="w")

        # --- Download folder (editable by the stakeholder) ---
        path_frame = tk.LabelFrame(
            cont, text="  Carpeta de descarga  ",
            font=("Segoe UI", 10, "bold"),
            bg=self.bg_color, fg=self.primary_color,
            bd=1, relief=tk.SOLID, padx=15, pady=12,
        )
        path_frame.pack(fill="x", pady=(0, 12))

        tk.Entry(
            path_frame, textvariable=self.path_var,
            font=("Segoe UI", 9), bg=self.light_gray, fg=self.primary_color,
            relief=tk.FLAT, bd=1, highlightthickness=1,
            highlightbackground=self.border_color, highlightcolor=self.primary_color,
        ).pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 8), ipady=4)

        tk.Button(
            path_frame, text="Examinar...", font=("Segoe UI", 9, "bold"),
            bg=self.primary_color, fg=self.bg_color, relief=tk.FLAT,
            cursor="hand2", padx=12, command=self._elegir_carpeta,
        ).pack(side=tk.LEFT)

        # --- Consolidation folder (green, like the consolidate button) ---
        cons_path_frame = tk.LabelFrame(
            cont, text="  Carpeta de consolidación  ",
            font=("Segoe UI", 10, "bold"),
            bg=self.bg_color, fg=self.primary_color,
            bd=1, relief=tk.SOLID, padx=15, pady=12,
        )
        cons_path_frame.pack(fill="x", pady=(0, 12))

        tk.Entry(
            cons_path_frame, textvariable=self.consolidacion_path_var,
            font=("Segoe UI", 9), bg=self.light_gray, fg=self.primary_color,
            relief=tk.FLAT, bd=1, highlightthickness=1,
            highlightbackground=self.border_color, highlightcolor=self.primary_color,
        ).pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 8), ipady=4)

        tk.Button(
            cons_path_frame, text="Examinar...", font=("Segoe UI", 9, "bold"),
            bg=self.success_color, fg=self.bg_color, relief=tk.FLAT,
            cursor="hand2", padx=12, command=self._elegir_carpeta_consolidacion,
        ).pack(side=tk.LEFT)

    # --- Folder pickers ---
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

    # --- File names: the typed dates, no offsets (those only apply to the SAP queries) ---
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
        """File of a 2nd-run transaction ('vbrp', 'fbl5n'...), named with the GUI dates."""
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

    def _preview_nombre_eiap(self):
        # 'Venta DSD + EIAP' only works with a range.
        if self.mode_var.get() == "single":
            return "(sólo con 'Varios días')"
        try:
            return nombre_archivo_rango_eiap(
                self.date_from_var.get().strip(),
                self.date_to_var.get().strip(),
            )
        except ValueError:
            return "(fecha inválida)"

    def _update_filename_preview(self):
        """The files of the selected download run (the 1st one also lists 'Venta DSD + EIAP', whose
        button sits next to Gallo and Monivoi)."""
        if not hasattr(self, "corrida_descarga_var"):
            return                                # still building the window
        if self.corrida_descarga() == 2:
            self.filename_var.set("\n".join(
                f"📄 {nombre}: {self._preview_nombre_2da(clave)}"
                for clave, nombre in CORRIDAS_DESCARGA[2]
            ))
            return
        self.filename_var.set(
            f"📄 Gallo:      {self._preview_nombre()}\n"
            f"📄 Monivoi:  {self._preview_nombre_monivoi()}\n"
            f"📄 VBRL:       {self._preview_nombre_vbrl()}\n"
            f"📄 FB03:       {self._preview_nombre_fb03()}\n"
            f"📄 Venta DSD + EIAP: {self._preview_nombre_eiap()}"
        )

    def _on_mode_change(self):
        if self.mode_var.get() == "single":
            self.range_frame.pack_forget()
            self.single_frame.pack(fill="x")
        else:
            self.single_frame.pack_forget()
            self.range_frame.pack(fill="x")
        self._update_filename_preview()

    def _ajustar_alto_ventana(self):
        """Grow the window ONLY when display scaling makes the content taller than the design size
        (measured with the taller 'Varios días' frame AND its file names, the longest preview text:
        measured with an empty preview, a wrapped name line was left out of view).
        """
        modo = self.mode_var.get()
        corrida = self.corrida_descarga_var.get()
        self.mode_var.set("range")
        self.single_frame.pack_forget()
        self.range_frame.pack(fill="x")
        ancho, alto = self.ANCHO_BASE, self.ALTO_BASE
        for opcion in CORRIDAS_DESCARGA:               # every run's legend and file names
            self.corrida_descarga_var.set(opcion)
            self._on_corrida_descarga_change()
            self.root.update_idletasks()
            ancho = max(ancho, self.root.winfo_reqwidth())
            alto = max(alto, self.root.winfo_reqheight())
        self.mode_var.set(modo)
        self.corrida_descarga_var.set(corrida)
        self._on_corrida_descarga_change()
        self.root.geometry(f"{ancho}x{alto}")
        self.range_frame.pack_forget()

    # --- Validation and configuration ---
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

    def get_config(self, transaccion="gallo"):
        mode = self.mode_var.get()
        # Any other value falls back to the Gallo name.
        nombres = {
            "monivoi": self._preview_nombre_monivoi,
            "vbrl":    self._preview_nombre_vbrl,
            "fb03":    self._preview_nombre_fb03,
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

    def set_status(self, message):
        self.status_var.set(message)
        self.root.update_idletasks()

    # --- Button state ---
    def _botones_accion(self):
        """(button, enabled color) for EVERY button that starts a process."""
        return (
            (self.download_gallo_btn,   self.primary_color),
            (self.download_monivoi_btn, self.primary_color),
            (self.download_eiap_btn,    self.primary_color),
            (self.download_btn,         self.primary_color),
            (self.consolidar_btn,       self.success_color),
        )

    def _radios_corrida(self):
        return tuple(self.radios_descarga.values()) + tuple(self.radios_consolidacion.values())

    def disable_buttons(self):
        for btn, _ in self._botones_accion():
            btn.config(state="disabled", bg="#666666")
        self.paralelo_chk.config(state="disabled")
        for radio in self._radios_corrida():           # the run cannot change while one is running
            radio.config(state="disabled")

    def enable_buttons(self):
        for btn, color in self._botones_accion():
            btn.config(state="normal", bg=color)
        self.paralelo_chk.config(state="normal")
        for radio in self._radios_corrida():
            radio.config(state="normal")
        self._estado_radio_consolidacion_2()

    def usar_paralelo(self):
        """True if the 'en paralelo' checkbox is ticked (read by the controller)."""
        return bool(self.paralelo_var.get())

    # --- Handlers ---
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

    def _handle_eiap(self):
        self._disparar(self.on_download_eiap, "eiap")

    def _handle_todos(self):
        self._disparar(self.on_download_todos, "gallo")

    def _handle_consolidar(self):
        self._disparar(self.on_consolidar, "gallo")


def main():
    root = tk.Tk()
    ValidacionFacturaGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
