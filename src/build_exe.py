"""
Script para crear el ejecutable de Validación Factura Global
Ejecutar desde la carpeta donde están los .py: python build_exe.py
"""

import subprocess
import sys
import os


def install_pyinstaller():
    """Instalar PyInstaller si no está disponible"""
    try:
        import PyInstaller  # noqa: F401
        print("✅ PyInstaller ya está instalado")
    except ImportError:
        print("📦 Instalando PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✅ PyInstaller instalado correctamente")


def create_executable():
    """Crear el ejecutable"""
    print("🚀 Creando ejecutable...")

    cmd = [
        "pyinstaller",
        "--onedir",             # UN solo .exe (fácil de enviar al stakeholder)
        "--windowed",            # Sin ventana de consola (app con GUI)
        "--name=ValidacionFacturaGlobal",  # Nombre del ejecutable
        "--icon=NONE",           # Sin icono personalizado (cambiar si hay un .ico)
        "--clean",               # Limpia caché de builds anteriores
        "--noconfirm",           # No pregunta al sobrescribir dist/ y build/

        # Módulos ocultos que PyInstaller no detecta automáticamente
        "--hidden-import=openpyxl",
        "--hidden-import=openpyxl.cell._writer",
        "--hidden-import=pandas",
        "--hidden-import=win32com",
        "--hidden-import=win32com.client",
        "--hidden-import=pywintypes",

        # Punto de entrada
        "main.py",
    ]

    try:
        subprocess.run(cmd, check=True)
        print("✅ Ejecutable creado exitosamente!")
        print("📁 Ubicación: dist/ValidacionFacturaGlobal.exe")

        exe_path = os.path.join("dist", "ValidacionFacturaGlobal.exe")
        if os.path.exists(exe_path):
            size_mb = os.path.getsize(exe_path) / (1024 * 1024)
            print(f"📊 Tamaño del ejecutable: {size_mb:.1f} MB")

            print("\n📋 INSTRUCCIONES PARA EL USUARIO:")
            print("1. Distribuir el archivo: dist/ValidacionFacturaGlobal.exe")
            print("2. El usuario solo necesita ejecutar el .exe")
            print("3. No requiere tener Python instalado")
            print("4. SAP GUI debe estar abierto y conectado antes de usar la descarga")
            print("5. La primera ejecución creará automáticamente las carpetas necesarias en Documentos")
        else:
            print("❌ Error: No se encontró el ejecutable después de la compilación")

    except subprocess.CalledProcessError as e:
        print(f"❌ Error creando ejecutable: {e}")
        print("\n💡 Sugerencias:")
        print("   - Verifica que estés en la carpeta donde están los archivos .py")
        print("   - Ejecuta: pip install pyinstaller")
        print("   - Si el error menciona 'win32com', ejecuta: pip install pywin32")
        print("   - Si el error menciona 'openpyxl', ejecuta: pip install openpyxl")
        print("   - Si el error menciona 'pandas', ejecuta: pip install pandas")


def check_source_files():
    """Verificar que todos los archivos fuente necesarios existen"""
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

    # Verificar archivos fuente
    if not check_source_files():
        print("\n❌ Faltan archivos fuente. Asegúrate de ejecutar este script")
        print("   desde la carpeta donde están todos los archivos .py del proyecto.")
        return

    print()

    # Instalar PyInstaller si es necesario
    install_pyinstaller()

    print()

    # Crear ejecutable
    create_executable()

    print("\n🎉 ¡Proceso completado!")


if __name__ == "__main__":
    main()
