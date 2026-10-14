"""
Script that builds the Validación Factura Global executable.

Run it from the folder that holds the .py files, WITH THE VIRTUAL ENVIRONMENT
ACTIVE:

    .venv\\Scripts\\activate
    python build_exe.py

PyInstaller bundles what the interpreter that runs it can import, so the
environment decides the size of dist/: from the project venv (pandas, openpyxl,
pywin32, pyinstaller + their dependencies) it only packs those; from the global
Python it would pack everything installed there.
"""

import subprocess
import sys
import os

# --- Build configuration ----------------------------------------------
# APP_NAME is PyInstaller's --name, and together with the build mode it decides
# WHERE the .exe ends up:
#     --onedir  → dist/<APP_NAME>/<APP_NAME>.exe   (folder: .exe + _internal/)
#     --onefile → dist/<APP_NAME>.exe              (single file, slower to start)
# --onedir is the default here because corporate antivirus tends to be less
# nervous about it; the whole folder is what gets zipped and handed over.
APP_NAME = "ValidacionFacturaGlobal"
MODO_BUILD = "--onedir"

# Heavy libraries the app never imports. They only sneak into dist/ when the
# environment used for the build has them installed (pandas declares several of
# them as optional extras), and they add hundreds of MB. Building from a clean
# virtual environment makes this list unnecessary, but it does no harm there.
# If a build ever fails with ModuleNotFoundError for one of these, remove it.
MODULOS_EXCLUIDOS = [
    "matplotlib", "scipy", "IPython", "jedi", "parso", "tornado", "zmq",
    "nbformat", "jsonschema", "prompt_toolkit", "psutil", "PIL",
    "contourpy", "kiwisolver",
]


def ruta_ejecutable():
    """
    Where PyInstaller leaves the .exe for the current build mode (see
    MODO_BUILD). Used to check the build really produced it.
    """
    if MODO_BUILD == "--onefile":
        return os.path.join("dist", f"{APP_NAME}.exe")
    return os.path.join("dist", APP_NAME, f"{APP_NAME}.exe")


def _tamano_mb(ruta):
    """Size in MB of a file, or of a whole folder (--onedir leaves a folder)."""
    if os.path.isfile(ruta):
        total = os.path.getsize(ruta)
    else:
        total = sum(
            os.path.getsize(os.path.join(raiz, nombre))
            for raiz, _subcarpetas, archivos in os.walk(ruta)
            for nombre in archivos
        )
    return total / (1024 * 1024)


def install_pyinstaller():
    """Install PyInstaller if it is not available."""
    try:
        import PyInstaller  # noqa: F401
        print("✅ PyInstaller ya está instalado")
    except ImportError:
        print("📦 Instalando PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✅ PyInstaller instalado correctamente")


def create_executable():
    """
    Build the executable.

    Returns True only if PyInstaller finished AND the .exe really showed up
    where the current build mode leaves it (ruta_ejecutable).
    """
    print("🚀 Creando ejecutable...")

    # PyInstaller is invoked as a MODULE of the interpreter that is running this
    # script (sys.executable), never as the loose 'pyinstaller' command: what
    # ends up inside dist/ is whatever THAT interpreter can import. Calling the
    # command found on PATH could be the global Python's PyInstaller, which
    # would bundle the global site-packages and silently undo the isolation of
    # the virtual environment (that is also why this must be run with the venv
    # active: .venv\Scripts\activate && python build_exe.py).
    cmd = [
        sys.executable, "-m", "PyInstaller",
        MODO_BUILD,              # --onedir / --onefile (see the constants above)
        "--windowed",            # No console window (GUI app)
        f"--name={APP_NAME}",    # Name of the executable
        "--icon=NONE",           # No custom icon (change it if there is a .ico)
        "--clean",               # Clears the cache of previous builds
        "--noconfirm",           # Does not ask before overwriting dist/ and build/

        # Hidden imports that PyInstaller does not detect automatically
        "--hidden-import=openpyxl",
        "--hidden-import=openpyxl.cell._writer",
        "--hidden-import=pandas",
        "--hidden-import=win32com",
        "--hidden-import=win32com.client",
        "--hidden-import=pywintypes",
    ]
    # Heavy libraries the app never imports (see MODULOS_EXCLUIDOS).
    cmd += [f"--exclude-module={modulo}" for modulo in MODULOS_EXCLUIDOS]
    # Entry point
    cmd.append("main.py")

    try:
        subprocess.run(cmd, check=True)
    except subprocess.CalledProcessError as e:
        print(f"❌ Error creando ejecutable: {e}")
        print("\n💡 Sugerencias:")
        print("   - Verifica que estés en la carpeta donde están los archivos .py")
        print("   - Ejecuta: pip install pyinstaller")
        print("   - Si el error menciona 'win32com', ejecuta: pip install pywin32")
        print("   - Si el error menciona 'openpyxl', ejecuta: pip install openpyxl")
        print("   - Si el error menciona 'pandas', ejecuta: pip install pandas")
        print("   - Si es un ModuleNotFoundError de una librería que SÍ se usa,")
        print("     quítala de MODULOS_EXCLUIDOS (hasta arriba de este script)")
        return False

    exe_path = ruta_ejecutable()
    if not os.path.exists(exe_path):
        print("❌ Error: No se encontró el ejecutable después de la compilación")
        print(f"   Lo esperaba en: {exe_path}")
        return False

    carpeta = os.path.dirname(exe_path)
    print("✅ ¡Ejecutable creado exitosamente!")
    print(f"📁 Ubicación: {exe_path}")
    if MODO_BUILD == "--onedir":
        print(f"📊 Tamaño: {_tamano_mb(carpeta):,.1f} MB la carpeta completa "
              f"({_tamano_mb(exe_path):,.1f} MB el .exe)")
    else:
        print(f"📊 Tamaño del ejecutable: {_tamano_mb(exe_path):,.1f} MB")

    print("\n📋 INSTRUCCIONES PARA EL USUARIO:")
    if MODO_BUILD == "--onedir":
        print(f"1. Distribuir la CARPETA completa: {carpeta} (comprimida en .zip)")
        print("2. El usuario la descomprime y ejecuta el .exe que viene adentro")
    else:
        print(f"1. Distribuir el archivo: {exe_path}")
        print("2. El usuario solo necesita ejecutar el .exe")
    print("3. No requiere tener Python instalado")
    print("4. SAP GUI debe estar abierto y conectado antes de usar la descarga")
    print("5. La primera ejecución creará automáticamente las carpetas necesarias en Documentos")
    return True


def check_source_files():
    """Check that every required source file exists."""
    archivos_requeridos = [
        "main.py",
        "interfaz_GUI.py",
        "controller.py",
        "Descargas_SAP.py",
        "Consolidacion.py",
        "utils.py",
    ]

    print("🔍 Verificando archivos fuente...")
    todos_ok = True
    for archivo in archivos_requeridos:
        if os.path.exists(archivo):
            print(f"   ✅ {archivo}")
        else:
            print(f"   ❌ {archivo} — NO ENCONTRADO")
            todos_ok = False

    return todos_ok


def main():
    print("🔨 CONSTRUCCIÓN DE EJECUTABLE - VALIDACIÓN FACTURA GLOBAL")
    print("=" * 60)

    # Check the source files
    if not check_source_files():
        print("\n❌ Faltan archivos fuente. Asegúrate de ejecutar este script")
        print("   desde la carpeta donde están todos los archivos .py del proyecto.")
        return

    print()

    # Install PyInstaller if needed
    install_pyinstaller()

    print()

    # Build the executable
    ok = create_executable()

    print("\n🎉 ¡Proceso completado!" if ok
          else "\n⚠️ El proceso terminó con errores (revisa los mensajes de arriba).")


if __name__ == "__main__":
    main()
