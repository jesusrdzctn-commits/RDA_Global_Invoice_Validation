import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import os


# Mes en español (espejo del helper en Descargas_SAP.py). Se duplica a propósito
# para que la GUI NO importe nada de SAP y se pueda abrir/probar en cualquier
# equipo. Cuando el proyecto crezca lo movemos a un utils.py compartido.
_MESES_ES = {
    1: "ene", 2: "feb", 3: "mar", 4: "abr", 5: "may", 6: "jun",
    7: "jul", 8: "ago", 9: "sep", 10: "oct", 11: "nov", 12: "dic",
}


class ValidacionFacturaGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Validación Factura Global")
        self.root.geometry("580x600")
        self.root.resizable(False, False)

        # Esquema de colores 
        self.bg_color        = "#FFFFFF"
        self.primary_color   = "#00094F"
        self.secondary_color = "#333333"
        self.light_gray      = "#F5F5F5"
        self.border_color    = "#E0E0E0"
        self.success_color   = "#28A745"

        self.root.configure(bg=self.bg_color)

        # Callback que el controller conectará
        self.on_download_gallo = None

        # Modo de intervalo: "single" (un día) o "range" (varios días)
        self.mode_var = tk.StringVar(value="single")

        # Ruta de salida dinámica según el usuario del sistema
        user_profile = os.environ.get("USERPROFILE") or os.path.expanduser("~")
        self.input_path = os.path.join(
            user_profile, "Documents", "Validacion Factura Global", "src", "Input"
        )
        os.makedirs(self.input_path, exist_ok=True)

        self._build_ui()
        self._on_mode_change()  # muestra el frame correcto + preview inicial

    # ------------------------------------------------------------------
    # Construcción de la interfaz
    # ------------------------------------------------------------------
    def _build_ui(self):
        main = tk.Frame(self.root, bg=self.bg_color, padx=24, pady=20)
        main.pack(fill="both", expand=True)

        tk.Label(
            main, text="Validación Factura Global",
            font=("Segoe UI", 16, "bold"),
            bg=self.bg_color, fg=self.primary_color,
        ).pack(pady=(0, 4))
        tk.Label(
            main, text="Descarga de transacción Gallo (FAGLL03)",
            font=("Segoe UI", 10),
            bg=self.bg_color, fg=self.secondary_color,
        ).pack(pady=(0, 16))

        # --- Selector de modo ---
        mode_frame = tk.LabelFrame(
            main, text="  Tipo de intervalo  ",
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
        self.date_container = tk.Frame(main, bg=self.bg_color)
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

        # --- Parámetros (sociedad) ---
        param_frame = tk.LabelFrame(
            main, text="  Parámetros  ",
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

        # --- Info de salida (nombre + ruta) ---
        info_frame = tk.Frame(main, bg=self.light_gray, bd=1, relief=tk.SOLID)
        info_frame.pack(fill="x", pady=(0, 14))

        self.filename_var = tk.StringVar()
        tk.Label(
            info_frame, textvariable=self.filename_var,
            font=("Segoe UI", 9, "bold"),
            bg=self.light_gray, fg=self.primary_color, anchor="w", justify="left",
            wraplength=500,
        ).pack(fill="x", padx=12, pady=(10, 2))
        tk.Label(
            info_frame, text=f"📁 {self.input_path}", font=("Segoe UI", 8),
            bg=self.light_gray, fg="#666666", anchor="w", justify="left",
            wraplength=500,
        ).pack(fill="x", padx=12, pady=(0, 10))

        # --- Botón de descarga ---
        self.download_btn = tk.Button(
            main, text="🐓 Descargar Gallo (FAGLL03)",
            command=self._handle_download,
            font=("Segoe UI", 11, "bold"),
            bg=self.primary_color, fg=self.bg_color,
            relief=tk.RAISED, bd=2, padx=30, pady=12, cursor="hand2",
            activebackground=self.secondary_color, activeforeground=self.bg_color,
        )
        self.download_btn.pack(fill="x", pady=(0, 14))

        # --- Barra de estado ---
        self.status_var = tk.StringVar(value="✓ Listo para comenzar")
        tk.Label(
            main, textvariable=self.status_var, font=("Segoe UI", 10),
            bg=self.bg_color, fg="#666666",
        ).pack()

    # ------------------------------------------------------------------
    # Helpers de nombrado (espejo de Descargas_SAP.py, sin desfase de día)
    # ------------------------------------------------------------------
    def _nombre_dia(self, fecha_str):
        f = datetime.strptime(fecha_str, "%d.%m.%Y")
        return f"{f.day}_{_MESES_ES[f.month]}_{f:%y}.csv"

    def _nombre_rango(self, desde, hasta):
        d = datetime.strptime(desde, "%d.%m.%Y")
        h = datetime.strptime(hasta, "%d.%m.%Y")
        return (
            f"{d.day}_{_MESES_ES[d.month]}_{d:%y}"
            f"_a_{h.day}_{_MESES_ES[h.month]}_{h:%y}.csv"
        )

    def _preview_nombre(self):
        try:
            if self.mode_var.get() == "single":
                return self._nombre_dia(self.date_single_var.get().strip())
            return self._nombre_rango(
                self.date_from_var.get().strip(),
                self.date_to_var.get().strip(),
            )
        except ValueError:
            return "(fecha inválida)"

    def _update_filename_preview(self):
        self.filename_var.set(f"📄 El archivo se llamará:  {self._preview_nombre()}")

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

    # ------------------------------------------------------------------
    # Validación y configuración
    # ------------------------------------------------------------------
    def validate_dates(self):
        try:
            if self.mode_var.get() == "single":
                datetime.strptime(self.date_single_var.get().strip(), "%d.%m.%Y")
                return True
            d1 = datetime.strptime(self.date_from_var.get().strip(), "%d.%m.%Y")
            d2 = datetime.strptime(self.date_to_var.get().strip(), "%d.%m.%Y")
            if d1 > d2:
                messagebox.showerror("Error", "La fecha desde no puede ser mayor a la fecha hasta")
                return False
            return True
        except ValueError:
            messagebox.showerror("Error", "Formato de fecha inválido. Use DD.MM.YYYY")
            return False

    def get_config(self):
        mode = self.mode_var.get()
        cfg = {
            "mode":       mode,
            "sociedad":   self.sociedad_var.get().strip().upper() or "MX21",
            "input_path": self.input_path,
            "filename":   self._preview_nombre(),
        }
        if mode == "single":
            cfg["fecha"] = self.date_single_var.get().strip()
        else:
            cfg["date_from"] = self.date_from_var.get().strip()
            cfg["date_to"]   = self.date_to_var.get().strip()
        return cfg

    def set_status(self, message):
        self.status_var.set(message)
        self.root.update_idletasks()

    def disable_buttons(self):
        self.download_btn.config(state="disabled", bg="#666666")

    def enable_buttons(self):
        self.download_btn.config(state="normal", bg=self.primary_color)

    def _handle_download(self):
        if not self.validate_dates():
            return
        if self.on_download_gallo:
            self.on_download_gallo()
        else:
            messagebox.showinfo("Info", f"Funcionalidad no conectada\n\n{self.get_config()}")


def main():
    # Punto de entrada standalone para ver la GUI sin el controller.
    root = tk.Tk()
    ValidacionFacturaGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
