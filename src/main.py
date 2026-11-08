"""Validación Factura Global — entry point. Without arguments it opens the GUI wired to the controller; with
arguments it is the command line (corrida_completa.py):  ValidacionFacturaGlobal.exe --corrida-completa --fecha 03.11.2026
"""

import sys


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


def abrir_interfaz():
    """The window. Tk, the GUI and the controller are imported HERE so the command line never loads them."""
    import tkinter as tk
    from interfaz_GUI import ValidacionFacturaGUI
    from controller import ValidacionFacturaController

    activar_dpi_alto()
    root = tk.Tk()
    gui = ValidacionFacturaGUI(root)
    ValidacionFacturaController(gui)
    root.mainloop()
    return 0


def main(argv=None):
    """→ the exit code of the process. Any argument means the command line (a typo must NOT open a window
    that a scheduled task would wait on forever); none means the GUI."""
    argv = sys.argv[1:] if argv is None else list(argv)
    if argv:
        from corrida_completa import ejecutar_linea_de_comandos
        return ejecutar_linea_de_comandos(argv)
    return abrir_interfaz()


if __name__ == "__main__":
    sys.exit(main())
