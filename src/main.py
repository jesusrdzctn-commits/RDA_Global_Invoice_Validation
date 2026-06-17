"""
Validación Factura Global
==========================

Punto de entrada principal de la aplicación.
Conecta la interfaz gráfica (GUI) con la lógica de negocio (Controller).

Fecha: 2026
"""

import tkinter as tk
from interfaz_GUI import ValidacionFacturaGUI
from controller import ValidacionFacturaController


def main():
    root = tk.Tk()
    gui = ValidacionFacturaGUI(root)
    ValidacionFacturaController(gui)   # engancha los callbacks de la GUI
    root.mainloop()


if __name__ == "__main__":
    main()
