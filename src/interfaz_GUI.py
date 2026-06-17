import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timedelta
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
        self.root.geometry("560x540")
        self.root.resizable(False, False)

        # Esquema de colores 
        self.bg_color        = "#FFFFFF"
        self.primary_color   = "#0D005F"
        self.secondary_color = "#333333"
        self.light_gray      = "#F5F5F5"
        self.border_color    = "#E0E0E0"
        self.success_color   = "#28A745"

        self.root.configure(bg=self.bg_color)

        # Callback que el controller conectará
        self.on_download_gallo = None

        # Ruta de salida dinámica según el usuario del sistema
        user_profile = os.environ.get("USERPROFILE") or os.path.expanduser("~")
        self.input_path = os.path.join(
            user_profile, "Documents", "Validacion Factura Global", "src", "Input"
        )
        os.makedirs(self.input_path, exist_ok=True)

        self._build_ui()
        self._update_filename_preview()

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
        ).pack(pady=(0, 18))

        # --- Intervalo de tiempo ---
        date_frame = tk.LabelFrame(
            main, text="  Intervalo de Tiempo  ",
            font=("Segoe UI", 10, "bold"),
            bg=self.bg_color, fg=self.primary_color,
            bd=1, relief=tk.SOLID, padx=15, pady=15,
        )
        date_frame.pack(fill="x", pady=(0, 16))

        hoy = datetime.now().strftime("%d.%m.%Y")
        for row, (label, attr, default) in enumerate([
            ("Fecha Desde:", "date_from_var", hoy),
            ("Fecha Hasta:", "date_to_var",   hoy),
        ]):
            tk.Label(
                date_frame, text=label, font=("Segoe UI", 9),
                bg=self.bg_color, fg=self.secondary_color,
            ).grid(row=row, column=0, sticky="w", pady=8)

            var = tk.StringVar(value=default)
            setattr(self, attr, var)
            tk.Entry(
                date_frame, textvariable=var, width=15,
                font=("Segoe UI", 10), bg=self.light_gray, fg=self.primary_color,
                relief=tk.FLAT, bd=1, highlightthickness=1,
                highlightbackground=self.border_color, highlightcolor=self.primary_color,
            ).grid(row=row, column=1, padx=10, pady=8, sticky="w")

            tk.Label(
                date_frame, text="(DD.MM.YYYY)", font=("Segoe UI", 8),
                bg=self.bg_color, fg="#999999",
            ).grid(row=row, column=2, sticky="w")

        # Recalcular la vista previa del nombre cuando cambie la fecha "hasta"
        self.date_to_var.trace_add("write", lambda *_: self._update_filename_preview())

        # --- Parámetros (sociedad) ---
        param_frame = tk.LabelFrame(
            main, text="  Parámetros  ",
            font=("Segoe UI", 10, "bold"),
            bg=self.bg_color, fg=self.primary_color,
            bd=1, relief=tk.SOLID, padx=15, pady=15,
        )
        param_frame.pack(fill="x", pady=(0, 16))

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
        info_frame.pack(fill="x", pady=(0, 16))

        self.filename_var = tk.StringVar()
        tk.Label(
            info_frame, textvariable=self.filename_var,
            font=("Segoe UI", 9, "bold"),
            bg=self.light_gray, fg=self.primary_color, anchor="w", justify="left",
        ).pack(fill="x", padx=12, pady=(10, 2))
        tk.Label(
            info_frame, text=f"📁 {self.input_path}", font=("Segoe UI", 8),
            bg=self.light_gray, fg="#666666", anchor="w", justify="left",
            wraplength=480,
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
        self.download_btn.pack(fill="x", pady=(0, 16))

        # --- Barra de estado ---
        self.status_var = tk.StringVar(value="✓ Listo para comenzar")
        tk.Label(
            main, textvariable=self.status_var, font=("Segoe UI", 10),
            bg=self.bg_color, fg="#666666",
        ).pack()

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _nombre_preview(self, date_high):
        # Calcula el nombre (fecha + 1 día) en formato DD_mmm_YY.xlsx.
        try:
            f = datetime.strptime(date_high, "%d.%m.%Y") + timedelta(days=1)
            return f"{f.day}_{_MESES_ES[f.month]}_{f:%y}.xlsx"
        except ValueError:
            return "(fecha inválida)"

    def _update_filename_preview(self):
        nombre = self._nombre_preview(self.date_to_var.get().strip())
        self.filename_var.set(f"📄 El archivo se llamará:  {nombre}")

    def validate_dates(self):
        # Valida formato DD.MM.YYYY y que Desde <= Hasta.
        try:
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
        # Devuelve todos los parámetros actuales de la GUI.
        return {
            "date_from":  self.date_from_var.get().strip(),
            "date_to":    self.date_to_var.get().strip(),
            "sociedad":   self.sociedad_var.get().strip().upper() or "MX21",
            "input_path": self.input_path,
            "filename":   self._nombre_preview(self.date_to_var.get().strip()),
        }

    def set_status(self, message):
        self.status_var.set(message)
        self.root.update_idletasks()

    def disable_buttons(self):
        self.download_btn.config(state="disabled", bg="#666666")

    def enable_buttons(self):
        self.download_btn.config(state="normal", bg=self.primary_color)

    def _handle_download(self):
        # Valida fechas y delega al callback de descarga.
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
