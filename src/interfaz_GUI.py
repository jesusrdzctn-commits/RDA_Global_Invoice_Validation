import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import os
import sys
import json


# Helpers de nombrado compartidos. Viven en utils.py (sin dependencias de SAP
# ni de tkinter), así la GUI se puede abrir y probar sola en cualquier equipo.
from utils import (
    nombre_archivo_dia,          nombre_archivo_rango,
    nombre_archivo_dia_monivoi,  nombre_archivo_rango_monivoi,
)


def _ruta_config():
    """
    Ubicación del config.json donde recordamos la última carpeta de descarga.
    Se guarda JUNTO al ejecutable (cuando está empaquetado con PyInstaller) o
    junto a este .py (cuando se corre como script). Así cada equipo conserva su
    propia ruta sin tocar el código.
    """
    if getattr(sys, "frozen", False):        # corriendo como .exe (PyInstaller)
        base = os.path.dirname(sys.executable)
    else:                                    # corriendo como script .py
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

        # Esquema de colores
        self.bg_color        = "#FFFFFF"
        self.primary_color   = "#00094F"
        self.secondary_color = "#333333"
        self.light_gray      = "#F5F5F5"
        self.border_color    = "#E0E0E0"
        self.success_color   = "#28A745"

        self.root.configure(bg=self.bg_color)

        # Callbacks que el controller conectará
        self.on_download_gallo   = None   # downloads ONLY Gallo from SAP
        self.on_download_monivoi = None   # downloads ONLY Monivoi from SAP
        self.on_download_ambos   = None   # descarga Gallo + Monivoi desde SAP
        self.on_consolidar       = None   # consolidación de los CSV descargados

        # Modo de intervalo: "single" (un día) o "range" (varios días)
        self.mode_var = tk.StringVar(value="single")

        # --- Ruta de descarga ---------------------------------------------
        # Arranca en una ruta por defecto según el usuario del sistema, pero el
        # stakeholder la puede cambiar (p. ej. a una carpeta de SharePoint
        # sincronizada con OneDrive) con el botón "Examinar...". Si ya se usó
        # una ruta antes, se recupera del config.json.
        self._config_file = _ruta_config()
        cfg = self._cargar_config()
        user_profile = os.environ.get("USERPROFILE") or os.path.expanduser("~")
        base_proyecto = os.path.join(
            user_profile, "Documents", "Validacion Factura Global", "src"
        )
        ruta_input_default  = os.path.join(base_proyecto, "Input")
        ruta_output_default = os.path.join(base_proyecto, "Output")

        # Ruta de descarga (dónde SAP deja los CSV) y ruta de consolidación
        # (dónde se guarda el Excel consolidado). Ambas se recuerdan en
        # config.json y ambas se pueden cambiar con su botón "Examinar...".
        self.path_var = tk.StringVar(
            value=(cfg.get("ruta_descarga") or "").strip() or ruta_input_default
        )
        self.consolidacion_path_var = tk.StringVar(
            value=(cfg.get("ruta_consolidacion") or "").strip() or ruta_output_default
        )
        os.makedirs(ruta_input_default, exist_ok=True)  # asegura que la default exista

        self._build_ui()
        self._ajustar_alto_ventana()  # keep everything visible (DPI / extra row)
        self._on_mode_change()  # muestra el frame correcto + preview inicial

    # ------------------------------------------------------------------
    # Persistencia de la última ruta usada (config.json)
    # ------------------------------------------------------------------
    def _cargar_config(self):
        """Lee el config.json completo (dict). Si no existe o falla → {}."""
        try:
            with open(self._config_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except (FileNotFoundError, json.JSONDecodeError, OSError, ValueError):
            return {}

    def _guardar_config(self):
        """Guarda ambas rutas (descarga y consolidación) en config.json."""
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
            # No es crítico: si no se puede guardar, la app sigue funcionando.
            print(f"[CONFIG] No se pudo guardar la configuración: {e}")

    # ------------------------------------------------------------------
    # Construcción de la interfaz
    #
    # Estructura: un encabezado fijo arriba, un Notebook con dos pestañas en
    # medio ("Descarga y Consolidación" y "Configuración") y la barra de estado
    # SIEMPRE visible abajo (fuera del Notebook), para que el estatus se vea sin
    # importar en qué pestaña estés.
    # ------------------------------------------------------------------
    def _build_ui(self):
        main = tk.Frame(self.root, bg=self.bg_color, padx=24, pady=18)
        main.pack(fill="both", expand=True)

        # --- Encabezado (fijo) ---
        tk.Label(
            main, text="Validación Factura Global",
            font=("Segoe UI", 16, "bold"),
            bg=self.bg_color, fg=self.primary_color,
        ).pack(pady=(0, 4))
        tk.Label(
            main, text="Automatización de descarga y consolidación - Transacciones desde SAP",
            font=("Segoe UI", 10),
            bg=self.bg_color, fg=self.secondary_color,
        ).pack(pady=(0, 14))

        # --- Notebook con las dos pestañas ---
        style = ttk.Style()
        try:
            style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"),
                            padding=(16, 8))
        except Exception:
            pass  # si el tema del equipo ignora el estilo, no pasa nada

        notebook = ttk.Notebook(main)
        notebook.pack(fill="both", expand=True)

        tab_descarga = tk.Frame(notebook, bg=self.bg_color)
        tab_config   = tk.Frame(notebook, bg=self.bg_color)
        notebook.add(tab_descarga, text="  Descarga y Consolidación  ")
        notebook.add(tab_config,   text="  Configuración  ")

        self._build_tab_descarga(tab_descarga)
        self._build_tab_config(tab_config)

        # --- Barra de estado (fija, debajo del Notebook) ---
        self.status_var = tk.StringVar(value="✓ Listo para comenzar")
        tk.Label(
            main, textvariable=self.status_var, font=("Segoe UI", 10),
            bg=self.bg_color, fg="#666666",
        ).pack(pady=(12, 0))

    # ------------------------------------------------------------------
    # Pestaña 1: Descarga y Consolidación (modo, fechas, nombres, botones)
    # ------------------------------------------------------------------
    def _build_tab_descarga(self, parent):
        cont = tk.Frame(parent, bg=self.bg_color, padx=16, pady=16)
        cont.pack(fill="both", expand=True)

        # --- Selector de modo ---
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

        # --- Contenedor de fechas (alterna single/range) ---
        self.date_container = tk.Frame(cont, bg=self.bg_color)
        self.date_container.pack(fill="x", pady=(0, 12))

        hoy = datetime.now().strftime("%d.%m.%Y")

        # Frame "un solo día"
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

        # Frame "varios días"
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

        # Recalcular vista previa cuando cambie cualquier fecha
        for var in (self.date_single_var, self.date_from_var, self.date_to_var):
            var.trace_add("write", lambda *_: self._update_filename_preview())

        # --- Info de salida (SÓLO los nombres; la ruta vive en Configuración) ---
        info_frame = tk.Frame(cont, bg=self.light_gray, bd=1, relief=tk.SOLID)
        info_frame.pack(fill="x", pady=(0, 14))

        self.filename_var = tk.StringVar()
        tk.Label(
            info_frame, textvariable=self.filename_var,
            font=("Segoe UI", 9, "bold"),
            bg=self.light_gray, fg=self.primary_color, anchor="w", justify="left",
            wraplength=500,
        ).pack(fill="x", padx=12, pady=10)

        # --- Botones de acción: descargar y consolidar (separados) ---
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
        # uniform= forces both columns to the exact same width, even though
        # the button labels have different lengths.
        fila_individual.columnconfigure(0, weight=1, uniform="descargas")
        fila_individual.columnconfigure(1, weight=1, uniform="descargas")

        self.download_gallo_btn = tk.Button(
            fila_individual, text="🐓 Solo Gallo",
            command=self._handle_gallo, **_btn_individual_kwargs,
        )
        self.download_gallo_btn.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        self.download_monivoi_btn = tk.Button(
            fila_individual, text="📊 Solo Monivoi",
            command=self._handle_monivoi, **_btn_individual_kwargs,
        )
        self.download_monivoi_btn.grid(row=0, column=1, sticky="ew", padx=(5, 0))

        self.download_btn = tk.Button(
            btn_frame, text="🐓📊 Descargar Gallo + Monivoi",
            command=self._handle_ambos, **_btn_kwargs,
        )
        self.download_btn.pack(fill="x")

        # El botón de consolidación va en verde para distinguir "traer datos"
        # (azul) de "procesar datos" (verde).
        _btn_consolidar_kwargs = dict(_btn_kwargs, bg=self.success_color)
        self.consolidar_btn = tk.Button(
            btn_frame, text="🧩 Consolidar...!",
            command=self._handle_consolidar, **_btn_consolidar_kwargs,
        )
        self.consolidar_btn.pack(fill="x", pady=(10, 0))

    # ------------------------------------------------------------------
    # Pestaña 2: Configuración (sociedad + carpetas de descarga/consolidación)
    # ------------------------------------------------------------------
    def _build_tab_config(self, parent):
        cont = tk.Frame(parent, bg=self.bg_color, padx=16, pady=16)
        cont.pack(fill="both", expand=True)

        # --- Parámetros (sociedad) ---
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

        # --- Carpeta de descarga (editable por el stakeholder) ---
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

        # --- Carpeta de consolidación (dónde se guarda el Excel consolidado) ---
        # Igual que la de descarga: se puede escribir a mano o elegir con
        # "Examinar...". El botón va en verde para amarrarlo visualmente con el
        # botón de consolidar.
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
    # Selección de carpeta de descarga
    # ------------------------------------------------------------------
    def _elegir_carpeta(self):
        """Abre el explorador para que el stakeholder elija la carpeta de descarga."""
        carpeta = filedialog.askdirectory(
            title="Elige la carpeta donde se descargarán los archivos",
            initialdir=self.path_var.get() or os.path.expanduser("~"),
        )
        if carpeta:  # si el usuario cancela, no tocamos nada
            self.path_var.set(os.path.normpath(carpeta))
            self._guardar_config()  # recuerda la elección para la próxima vez

    def _elegir_carpeta_consolidacion(self):
        """Explorador para elegir dónde se guardará el Excel consolidado."""
        carpeta = filedialog.askdirectory(
            title="Elige la carpeta donde se guardará el Excel consolidado",
            initialdir=self.consolidacion_path_var.get() or os.path.expanduser("~"),
        )
        if carpeta:
            self.consolidacion_path_var.set(os.path.normpath(carpeta))
            self._guardar_config()

    # ------------------------------------------------------------------
    # Vista previa de nombres de archivo (sin desfase para Gallo; el -1 de
    # Monivoi ya viene aplicado dentro de utils.py)
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

    def _update_filename_preview(self):
        self.filename_var.set(
            f"📄 Gallo:      {self._preview_nombre()}\n"
            f"📄 Monivoi:  {self._preview_nombre_monivoi()}"
        )

    # ------------------------------------------------------------------
    # Alternar modo
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
    # Validación y configuración
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

        # --- Validar la carpeta de descarga ---
        ruta = self.path_var.get().strip()
        if not ruta:
            messagebox.showerror("Error", "Indica una carpeta de descarga válida.")
            return False
        try:
            os.makedirs(ruta, exist_ok=True)  # la crea si aún no existe
        except OSError as e:
            messagebox.showerror(
                "Error", f"No se pudo usar la carpeta de descarga:\n{ruta}\n\n{e}"
            )
            return False

        return True

    def get_config(self, transaccion="gallo"):
        mode = self.mode_var.get()
        filename = (
            self._preview_nombre_monivoi() if transaccion == "monivoi"
            else self._preview_nombre()
        )
        input_path = self.path_var.get().strip()
        os.makedirs(input_path, exist_ok=True)  # crea la carpeta si aún no existe
        self._guardar_config()  # recuerda la ruta usada en esta corrida
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
        """Ruta donde se guardará el Excel consolidado (la crea si no existe)."""
        ruta = self.consolidacion_path_var.get().strip()
        os.makedirs(ruta, exist_ok=True)
        self._guardar_config()  # recuerda la ruta usada
        return ruta

    def set_status(self, message):
        self.status_var.set(message)
        self.root.update_idletasks()

    # --- Estado de los botones ---
    def _botones_accion(self):
        """(button, enabled color) for EVERY button that starts a process."""
        return (
            (self.download_gallo_btn,   self.primary_color),
            (self.download_monivoi_btn, self.primary_color),
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

    def _handle_ambos(self):
        self._disparar(self.on_download_ambos, "gallo")

    def _handle_consolidar(self):
        self._disparar(self.on_consolidar, "gallo")


def main():
    # Punto de entrada standalone para ver la GUI sin el controller.
    root = tk.Tk()
    ValidacionFacturaGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
