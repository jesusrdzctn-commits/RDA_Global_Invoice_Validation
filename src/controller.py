import os
from tkinter import messagebox

from Descargas_SAP import Gallo_FAGLL03


class ValidacionFacturaController:
    """Conecta la GUI con la lógica de descarga SAP (transacción Gallo)."""

    def __init__(self, gui):
        self.gui = gui
        # Engancha el botón de la GUI a este método
        self.gui.on_download_gallo = self.execute_download_gallo

    def execute_download_gallo(self):
        config = self.gui.get_config()

        if not self.gui.validate_dates():
            return

        confirm = messagebox.askyesno(
            "Confirmar descarga Gallo",
            f"¿Iniciar la descarga de la transacción Gallo (FAGLL03)?\n\n"
            f"Sociedad : {config['sociedad']}\n"
            f"Periodo  : {config['date_from']} — {config['date_to']}\n"
            f"Archivo  : {config['filename']}\n\n"
            f"Se guardará en:\n{config['input_path']}"
        )
        if not confirm:
            return

        os.makedirs(config['input_path'], exist_ok=True)

        self.gui.disable_buttons()
        self.gui.set_status("⏳ Descargando Gallo... espere (no cierre SAP)")

        try:
            con_datos = Gallo_FAGLL03(
                DateFrom=config['date_from'],
                DateTo=config['date_to'],
                FolderPath=config['input_path'],
                FileName=config['filename'],
                sociedad=config['sociedad'],
                # variante usa el default 'BASE/VALID' definido en Descargas_SAP.py
            )

            if con_datos:
                self.gui.set_status("✅ ¡Descarga completada!")
                messagebox.showinfo(
                    "Éxito",
                    f"La Gallo se descargó correctamente.\n\n"
                    f"Archivo: {config['filename']}\n"
                    f"Ruta: {config['input_path']}"
                )
            else:
                self.gui.set_status("⚠️ Sin movimientos para ese periodo")
                messagebox.showwarning(
                    "Sin datos",
                    "La transacción no devolvió partidas para ese periodo/sociedad.\n"
                    "Se generó un archivo vacío como marcador."
                )
        except Exception as e:
            self.gui.set_status("❌ Error en la descarga")
            messagebox.showerror("Error", f"Ocurrió un error durante la descarga:\n\n{str(e)}")
        finally:
            self.gui.enable_buttons()
            if "completada" not in self.gui.status_var.get():
                self.gui.set_status("✓ Listo para comenzar")
