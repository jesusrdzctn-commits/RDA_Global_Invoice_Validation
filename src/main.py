"""
Validación Factura Global
==========================

Main entry point of the application.
Connects the graphical interface (GUI) with the business logic (Controller).

Date: 2026
"""

import tkinter as tk
from interfaz_GUI import ValidacionFacturaGUI
from controller import ValidacionFacturaController


def main():
    root = tk.Tk()
    gui = ValidacionFacturaGUI(root)
    ValidacionFacturaController(gui)   # hooks up the GUI callbacks
    root.mainloop()


if __name__ == "__main__":
    main()
