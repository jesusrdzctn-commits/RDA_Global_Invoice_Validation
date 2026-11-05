"""Validación Factura Global — entry point: the GUI wired to the controller."""

import sys
import tkinter as tk
from interfaz_GUI import ValidacionFacturaGUI
from controller import ValidacionFacturaController


def activar_dpi_alto():
    """Windows only: make the process DPI-aware, so Tk draws crisp text and the GUI sees the real display
    scaling (it sizes its pixel measures from it). Must run before the first Tk window exists. Purely
    cosmetic: where it is not available (old Windows, no shcore) the window just keeps the system scaling."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)    # 1 = system DPI aware
        except (AttributeError, OSError):
            ctypes.windll.user32.SetProcessDPIAware()
    except (ImportError, AttributeError, OSError):
        pass


def main():
    activar_dpi_alto()
    root = tk.Tk()
    gui = ValidacionFacturaGUI(root)
    ValidacionFacturaController(gui)
    root.mainloop()


if __name__ == "__main__":
    main()
