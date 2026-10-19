"""Validación Factura Global — entry point: the GUI wired to the controller."""

import tkinter as tk
from interfaz_GUI import ValidacionFacturaGUI
from controller import ValidacionFacturaController


def main():
    root = tk.Tk()
    gui = ValidacionFacturaGUI(root)
    ValidacionFacturaController(gui)
    root.mainloop()


if __name__ == "__main__":
    main()
