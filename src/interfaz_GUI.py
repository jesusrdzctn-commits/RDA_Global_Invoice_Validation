"""
Validación Factura Global — GUI (tkinter)
=========================================

Only the window: dates, folders and buttons. It knows nothing about SAP or
pandas; the controller (controller.py) plugs the actions into the on_*
callbacks. That keeps this file importable and testable on its own:
`python interfaz_GUI.py` opens the window without the controller.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import os
import sys
import json


# Shared naming helpers. They live in utils.py (no SAP or tkinter
# dependencies), so the GUI can be opened and tested on its own on any machine.
from utils import (
    nombre_archivo_dia,          nombre_archivo_rango,
    nombre_archivo_dia_monivoi,  nombre_archivo_rango_monivoi,
    nombre_archivo_dia_vbrl,     nombre_archivo_rango_vbrl,
    nombre_archivo_rango_eiap,
)


def _ruta_config():
    """
    Location of the config.json where the last folders used (download and
    consolidation) are remembered. It is saved NEXT TO the executable (when
    packaged with PyInstaller) or next to this .py (when run as a script).
    This way each machine keeps its own folders without touching the code.
    """
    if getattr(sys, "frozen", False):        # running as .exe (PyInstaller)
        base = os.path.dirname(sys.executable)
    else:                                    # running as a .py script
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "config.json")


class ValidacionFacturaGUI:
    # Design size of the window. It is the MINIMUM: if the content needs more
    # room (display scaling, extra rows), _ajustar_alto_ventana enlarges it.
    ANCHO_BASE = 580
    ALTO_BASE  = 660

    def __init__(self, root):
        self.root = root
        self.root.title("Validación Factura Global")
        self.root.geometry(f"{self.ANCHO_BASE}x{self.ALTO_BASE}")
        self.root.resizable(False, False)

        # Color scheme
        self.bg_color        = "#FFFFFF"
        self.primary_color   = "#00094F"
        self.secondary_color = "#333333"
        self.light_gray      = "#F5F5F5"
        self.border_color    = "#E0E0E0"
        self.success_color   = "#28A745"

        self.root.configure(bg=self.bg_color)

        # Callbacks the controller will wire up
        self.on_download_gallo   = None   # downloads ONLY Gallo from SAP
        self.on_download_monivoi = None   # downloads ONLY Monivoi from SAP
        self.on_download_eiap    = None   # downloads ONLY Venta EIAP (Gallo, EIAP accounts) from SAP
        self.on_download_todos   = None   # downloads Gallo + Monivoi + VBRL from SAP
        self.on_consolidar       = None   # consolidates the downloaded CSVs

        # Interval mode: "single" (one day) or "range" (several days)
        self.mode_var = tk.StringVar(value="single")

        # --- Download folder ------------------------------------------------
        # Starts at a default path based on the system user, but the
        # stakeholder can change it (e.g. to a SharePoint folder synced with
        # OneDrive) with the "Examinar..." button. If a folder was used
        # before, it is restored from config.json.
        self._config_file = _ruta_config()
        cfg = self._cargar_config()
        user_profile = os.environ.get("USERPROFILE") or os.path.expanduser("~")
        base_proyecto = os.path.join(
            user_profile, "Documents", "Validacion Factura Global", "src"
        )
        ruta_input_default  = os.path.join(base_proyecto, "Input")
        ruta_output_default = os.path.join(base_proyecto, "Output")

        # Download folder (where SAP leaves the CSVs) and consolidation folder
        # (where the consolidated workbook is saved). Both are remembered in
        # config.json and both can be changed with their "Examinar..." button.
        self.path_var = tk.StringVar(
            value=(cfg.get("ruta_descarga") or "").strip() or ruta_input_default
        )
        self.consolidacion_path_var = tk.StringVar(
            value=(cfg.get("ruta_consolidacion") or "").strip() or ruta_output_default
        )
        os.makedirs(ruta_input_default, exist_ok=True)  # make sure the default one exists

        self._build_ui()
        self._ajustar_alto_ventana()  # keep everything visible (DPI / extra row)
        self._on_mode_change()  # shows the right frame + the first preview

    # ------------------------------------------------------------------
    # Persistence of the last folders used (config.json)
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
            # Not critical: if it cannot be saved, the app keeps working.
            print(f"[CONFIG] No se pudo guardar la configuración: {e}")

    # ------------------------------------------------------------------
    # Building the interface
    #
    # Structure: a fixed header on top, a Notebook with two tabs in the middle
    # ("Descarga y Consolidación" and "Configuración") and the status bar
    # ALWAYS visible at the bottom (outside the Notebook), so the status shows
    # no matter which tab you are on.
    # ------------------------------------------------------------------
    def _build_ui(self):
        main = tk.Frame(self.root, bg=self.bg_color, padx=24, pady=18)
        main.pack(fill="both", expand=True)

        # --- Header (fixed) ---
        tk.Label(
            main, text="Validación Factura Global",
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
            pass  # if the machine's theme ignores the style, no harm done

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

    # ------------------------------------------------------------------
    # Tab 1: "Descarga y Consolidación" (mode, dates, file names, buttons)
    # ------------------------------------------------------------------
    def _build_tab_descarga(self, parent):
        cont = tk.Frame(parent, bg=self.bg_color, padx=16, pady=16)
        cont.pack(fill="both", expand=True)

        # --- Mode selector ---
        mode_frame = tk.LabelFrame(
            cont, text="  Tipo de intervalo  ",
            font=("Segoe UI", 10, "bold"),
            bg=self.bg_color, fg=self.primary_color,
            bd=1, relief=tk.SOLID, padx=15, pady=10,
        )
        mode_frame.pack(fill="x", pady=(0, 12))

        for text, value in [("Un solo día", "single"), ("Varios días", "range")]:
            tk.Radiobutton(
                mode_frame, text=text, variable=self.mode_var, value=value,
                font=("Segoe UI", 9), bg=self.bg_color, fg=self.primary_color,
                activebackground=self.bg_color, selectcolor=self.light_gray,
                cursor="hand2", command=self._on_mode_change,
            ).pack(side=tk.LEFT, padx=(0, 16))

        # --- Date container (switches between single/range) ---
        self.date_container = tk.Frame(cont, bg=self.bg_color)
        self.date_container.pack(fill="x", pady=(0, 12))

        hoy = datetime.now().strftime("%d.%m.%Y")

        # "Single day" frame
        self.single_frame = tk.LabelFrame(
            self.date_container, text="  Fecha  ",
            font=("Segoe UI", 10, "bold"),
            bg=self.bg_color, fg=self.primary_color,
            bd=1, relief=tk.SOLID, padx=15, pady=15,
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
            bd=1, relief=tk.SOLID, padx=15, pady=15,
        )
        for row, (label, attr) in enumerate([
            ("Fecha Desde:", "date_from_var"),
            ("Fecha Hasta:", "date_to_var"),
        ]):
            tk.Label(
                self.range_frame, text=label, font=("Segoe UI", 9),
                bg=self.bg_color, fg=self.secondary_color,
            ).grid(row=row, column=0, sticky="w", pady=8)
            var = tk.StringVar(value=hoy)
            setattr(self, attr, var)
            tk.Entry(
                self.range_frame, textvariable=var, width=15,
                font=("Segoe UI", 10), bg=self.light_gray, fg=self.primary_color,
                relief=tk.FLAT, bd=1, highlightthickness=1,
                highlightbackground=self.border_color, highlightcolor=self.primary_color,
            ).grid(row=row, column=1, padx=10, pady=8, sticky="w")
            tk.Label(
                self.range_frame, text="(DD.MM.YYYY)", font=("Segoe UI", 8),
                bg=self.bg_color, fg="#999999",
            ).grid(row=row, column=2, sticky="w")

        # Recompute the preview whenever any date changes
        for var in (self.date_single_var, self.date_from_var, self.date_to_var):
            var.trace_add("write", lambda *_: self._update_filename_preview())

        # --- Output info (file names ONLY; the folder lives in the "Configuración" tab) ---
        info_frame = tk.Frame(cont, bg=self.light_gray, bd=1, relief=tk.SOLID)
        info_frame.pack(fill="x", pady=(0, 14))

        self.filename_var = tk.StringVar()
        tk.Label(
            info_frame, textvariable=self.filename_var,
            font=("Segoe UI", 9, "bold"),
            bg=self.light_gray, fg=self.primary_color, anchor="w", justify="left",
            wraplength=500,
        ).pack(fill="x", padx=12, pady=10)

        # --- Action buttons: download and consolidate (separate) ---
        btn_frame = tk.Frame(cont, bg=self.bg_color)
        btn_frame.pack(fill="x", pady=(4, 0))

        _btn_kwargs = dict(
            font=("Segoe UI", 11, "bold"),
            bg=self.primary_color, fg=self.bg_color,
            relief=tk.RAISED, bd=2, pady=12, cursor="hand2",
            activebackground=self.secondary_color, activeforeground=self.bg_color,
        )

        # Row of single-transaction downloads (above the combined button).
        # Same blue because they are the same SAP downloads, but narrower,
        # shorter and in italics so they read as "variants" of the main one.
        _btn_individual_kwargs = dict(
            _btn_kwargs, font=("Segoe UI", 10, "bold italic"), pady=6,
        )
        fila_individual = tk.Frame(btn_frame, bg=self.bg_color)
        fila_individual.pack(fill="x", pady=(0, 8))
        # uniform= forces the three columns to the exact same width, even
        # though the button labels have different lengths.
        fila_individual.columnconfigure(0, weight=1, uniform="descargas")
        fila_individual.columnconfigure(1, weight=1, uniform="descargas")
        fila_individual.columnconfigure(2, weight=1, uniform="descargas")

        self.download_gallo_btn = tk.Button(
            fila_individual, text="🐓 Solo Gallo",
            command=self._handle_gallo, **_btn_individual_kwargs,
        )
        self.download_gallo_btn.grid(row=0, column=0, sticky="ew", padx=(0, 4))

        self.download_monivoi_btn = tk.Button(
            fila_individual, text="📊 Solo Monivoi",
            command=self._handle_monivoi, **_btn_individual_kwargs,
        )
        self.download_monivoi_btn.grid(row=0, column=1, sticky="ew", padx=4)

        self.download_eiap_btn = tk.Button(
            fila_individual, text="📈 Solo Venta EIAP",
            command=self._handle_eiap, **_btn_individual_kwargs,
        )
        self.download_eiap_btn.grid(row=0, column=2, sticky="ew", padx=(4, 0))

        self.download_btn = tk.Button(
            btn_frame, text="🐓📊🧾 Descargar Gallo + Monivoi + VBRL",
            command=self._handle_todos, **_btn_kwargs,
        )
        self.download_btn.pack(fill="x")

        # The consolidation button is green to tell "bring data" (blue) apart
        # from "process data" (green).
        _btn_consolidar_kwargs = dict(_btn_kwargs, bg=self.success_color)
        self.consolidar_btn = tk.Button(
            btn_frame, text="🧩 Consolidar...!",
            command=self._handle_consolidar, **_btn_consolidar_kwargs,
        )
        self.consolidar_btn.pack(fill="x", pady=(10, 0))

    # ------------------------------------------------------------------
    # Tab 2: "Configuración" (company code + download/consolidation folders)
    # ------------------------------------------------------------------
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

        # --- Consolidation folder (where the consolidated workbook is saved) ---
        # Same as the download one: it can be typed by hand or picked with
        # "Examinar...". The button is green to tie it visually to the
        # consolidate button.
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

    # ------------------------------------------------------------------
    # Folder selection (download and consolidation)
    # ------------------------------------------------------------------
    def _elegir_carpeta(self):
        """Open the file browser so the stakeholder can pick the download folder."""
        carpeta = filedialog.askdirectory(
            title="Elige la carpeta donde se descargarán los archivos",
            initialdir=self.path_var.get() or os.path.expanduser("~"),
        )
        if carpeta:  # if the user cancels, nothing is touched
            self.path_var.set(os.path.normpath(carpeta))
            self._guardar_config()  # remember the choice for next time

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
    # File-name preview. None of the names carries a date offset: Monivoi's +1
    # day only applies to the SAP query (utils.fecha_efectiva_monivoi). VBRL is
    # not downloaded by date; it carries the same dates as the Monivoi file it
    # is built from.
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

    def _preview_nombre_eiap(self):
        # 'Solo Venta EIAP' only works with a range; the name carries the dates
        # as typed (the -1 day on the end date only applies to the SAP query).
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
        self.filename_var.set(
            f"📄 Gallo:      {self._preview_nombre()}\n"
            f"📄 Monivoi:  {self._preview_nombre_monivoi()}\n"
            f"📄 VBRL:       {self._preview_nombre_vbrl()}\n"
            f"📄 Venta EIAP: {self._preview_nombre_eiap()}"
        )

    # ------------------------------------------------------------------
    # Switch mode
    # ------------------------------------------------------------------
    def _on_mode_change(self):
        if self.mode_var.get() == "single":
            self.range_frame.pack_forget()
            self.single_frame.pack(fill="x")
        else:
            self.single_frame.pack_forget()
            self.range_frame.pack(fill="x")
        self._update_filename_preview()

    def _ajustar_alto_ventana(self):
        """
        The window is fixed-size, but with display scaling (125%/150%) the
        content can require more height than the design size and the bottom
        buttons would be clipped. Measure with the "Varios días" frame (the
        taller of the two) and grow ONLY when needed; never below the base
        size. _on_mode_change() then re-shows the correct frame.
        """
        self.single_frame.pack_forget()
        self.range_frame.pack(fill="x")
        self.root.update_idletasks()
        ancho = max(self.ANCHO_BASE, self.root.winfo_reqwidth())
        alto  = max(self.ALTO_BASE,  self.root.winfo_reqheight())
        self.root.geometry(f"{ancho}x{alto}")
        self.range_frame.pack_forget()

    # ------------------------------------------------------------------
    # Validation and configuration
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

        # --- Validate the download folder ---
        ruta = self.path_var.get().strip()
        if not ruta:
            messagebox.showerror("Error", "Indica una carpeta de descarga válida.")
            return False
        try:
            os.makedirs(ruta, exist_ok=True)  # creates it if it does not exist yet
        except OSError as e:
            messagebox.showerror(
                "Error", f"No se pudo usar la carpeta de descarga:\n{ruta}\n\n{e}"
            )
            return False

        return True

    def get_config(self, transaccion="gallo"):
        mode = self.mode_var.get()
        # Any other value falls back to the Gallo name (the default).
        nombres = {
            "monivoi": self._preview_nombre_monivoi,
            "vbrl":    self._preview_nombre_vbrl,
            "eiap":    self._preview_nombre_eiap,
        }
        filename = nombres.get(transaccion, self._preview_nombre)()
        input_path = self.path_var.get().strip()
        os.makedirs(input_path, exist_ok=True)  # creates the folder if it does not exist yet
        self._guardar_config()  # remember the folder used in this run
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
        self._guardar_config()  # remember the folder used
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

    def disable_buttons(self):
        for btn, _ in self._botones_accion():
            btn.config(state="disabled", bg="#666666")

    def enable_buttons(self):
        for btn, color in self._botones_accion():
            btn.config(state="normal", bg=color)

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
    # Standalone entry point to see the GUI without the controller.
    root = tk.Tk()
    ValidacionFacturaGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
