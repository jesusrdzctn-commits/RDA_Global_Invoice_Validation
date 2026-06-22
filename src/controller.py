import os
import time
from datetime import datetime, timedelta
from tkinter import messagebox

import pandas as pd

from Descargas_SAP import Gallo_FAGLL03, nombre_archivo_dia, nombre_archivo_rango


class ValidacionFacturaController:
    """Conecta la GUI con la lógica de descarga SAP (transacción Gallo)."""

    def __init__(self, gui):
        self.gui = gui
        # Engancha el botón de la GUI a este método
        self.gui.on_download_gallo = self.execute_download_gallo

    # =================================================================
    # ENTRADA: decide entre un solo día o un rango
    # =================================================================
    def execute_download_gallo(self):
        config = self.gui.get_config()

        if not self.gui.validate_dates():
            return

        if config["mode"] == "single":
            self._confirmar_y_ejecutar_single(config)
        else:
            self._confirmar_y_ejecutar_range(config)

    # =================================================================
    # MODO: UN SOLO DÍA
    # =================================================================
    def _confirmar_y_ejecutar_single(self, config):
        fecha = config["fecha"]

        confirm = messagebox.askyesno(
            "Confirmar descarga Gallo (un día)",
            f"¿Iniciar la descarga de la transacción Gallo (FAGLL03)?\n\n"
            f"Sociedad : {config['sociedad']}\n"
            f"Fecha    : {fecha}\n"
            f"Archivo  : {config['filename']}\n\n"
            f"Se guardará en:\n{config['input_path']}"
        )
        if not confirm:
            return

        os.makedirs(config["input_path"], exist_ok=True)
        self.gui.disable_buttons()
        self.gui.set_status(f"⏳ Descargando Gallo {fecha}... (no cierre SAP)")

        try:
            con_datos = Gallo_FAGLL03(
                DateFrom=fecha,
                DateTo=fecha,
                FolderPath=config["input_path"],
                FileName=config["filename"],
                sociedad=config["sociedad"],
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
                self.gui.set_status("⚠️ Sin movimientos para esa fecha")
                messagebox.showwarning(
                    "Sin datos",
                    "La transacción no devolvió partidas para esa fecha/sociedad.\n"
                    "Se generó un archivo vacío como marcador."
                )
        except Exception as e:
            self.gui.set_status("❌ Error en la descarga")
            messagebox.showerror("Error", f"Ocurrió un error durante la descarga:\n\n{e}")
        finally:
            self.gui.enable_buttons()
            if "completada" not in self.gui.status_var.get():
                self.gui.set_status("✓ Listo para comenzar")

    # =================================================================
    # MODO: VARIOS DÍAS (descarga día por día + apila en un solo Excel)
    # =================================================================
    def _confirmar_y_ejecutar_range(self, config):
        desde = config["date_from"]
        hasta = config["date_to"]
        dias  = self._iterar_dias(desde, hasta)

        confirm = messagebox.askyesno(
            "Confirmar descarga Gallo (varios días)",
            f"¿Iniciar la descarga de la transacción Gallo (FAGLL03)?\n\n"
            f"Sociedad      : {config['sociedad']}\n"
            f"Periodo       : {desde} — {hasta}\n"
            f"Días a bajar  : {len(dias)} (uno por uno)\n"
            f"Consolidado   : {config['filename']}\n\n"
            f"Los archivos por día se guardan en una subcarpeta,\n"
            f"y el consolidado se guarda en:\n{config['input_path']}"
        )
        if not confirm:
            return

        os.makedirs(config["input_path"], exist_ok=True)
        self.gui.disable_buttons()

        try:
            self._run_range_download(config, dias)
            n_dias, total_filas = getattr(self, "_resumen_apilado", (0, 0))
            self.gui.set_status("✅ ¡Descarga de varios días completada!")
            messagebox.showinfo(
                "Éxito",
                f"Se descargaron {len(dias)} día(s) y se apilaron en un solo CSV.\n\n"
                f"Consolidado: {config['filename']}\n"
                f"{total_filas:,} filas en una sola tabla.\n"
                f"Ruta: {config['input_path']}"
            )
        except Exception as e:
            self.gui.set_status("❌ Error en la descarga")
            messagebox.showerror("Error", f"Ocurrió un error durante la descarga:\n\n{e}")
        finally:
            self.gui.enable_buttons()
            if "completada" not in self.gui.status_var.get():
                self.gui.set_status("✓ Listo para comenzar")

    def _run_range_download(self, config, dias):
        input_path = config["input_path"]
        sociedad   = config["sociedad"]

        # Subcarpeta donde se guardan los archivos por día (se conservan)
        nombre_consolidado = config["filename"]                         # ej: 15_jun_26_a_17_jun_26.csv
        nombre_subcarpeta  = "Dias_" + os.path.splitext(nombre_consolidado)[0]
        carpeta_dias = os.path.join(input_path, nombre_subcarpeta)
        os.makedirs(carpeta_dias, exist_ok=True)

        archivos_con_datos = []
        total = len(dias)

        for i, dia in enumerate(dias, 1):
            self.gui.set_status(f"📥 Descargando día {i}/{total}: {dia}... (no cierre SAP)")
            fname = nombre_archivo_dia(dia)
            ruta_dia = os.path.join(carpeta_dias, fname)

            con_datos = Gallo_FAGLL03(
                DateFrom=dia,
                DateTo=dia,
                FolderPath=carpeta_dias,
                FileName=fname,
                sociedad=sociedad,
            )
            time.sleep(3)  # respiro entre descargas para que SAP libere la pantalla

            if con_datos and os.path.exists(ruta_dia):
                archivos_con_datos.append(ruta_dia)
                print(f"[RANGO] {dia}: con datos → {ruta_dia}")
            else:
                print(f"[RANGO] {dia}: sin datos → se omite del consolidado")

        # Apilar todos los días con datos en un solo CSV (una sola tabla)
        self.gui.set_status("📋 Apilando todos los días en un solo CSV...")
        ruta_consolidado = os.path.join(input_path, nombre_consolidado)
        n_dias, total_filas = self._apilar_a_csv(archivos_con_datos, ruta_consolidado)
        self._resumen_apilado = (n_dias, total_filas)
        print(
            f"[RANGO] Consolidado: {n_dias} día(s) con datos, "
            f"{total_filas} filas → {ruta_consolidado}"
        )

    # =================================================================
    # HELPERS
    # =================================================================
    @staticmethod
    def _iterar_dias(desde, hasta):
        """Devuelve la lista de fechas 'DD.MM.YYYY' entre desde y hasta (inclusive)."""
        d = datetime.strptime(desde, "%d.%m.%Y")
        h = datetime.strptime(hasta, "%d.%m.%Y")
        dias = []
        while d <= h:
            dias.append(d.strftime("%d.%m.%Y"))
            d += timedelta(days=1)
        return dias

    @staticmethod
    def _apilar_a_csv(rutas, ruta_final):
        """
        Lee todos los Excel por día, los concatena en UNA SOLA tabla y la guarda
        en formato CSV (sin límite de filas y todo en una sola hoja).

        Se usa utf-8-sig para que Excel muestre bien los acentos al previsualizar.
        Ignora archivos vacíos (placeholder 'Sin datos').
        Devuelve (n_archivos_con_datos, total_filas).
        """
        dfs = []
        for ruta in rutas:
            try:
                df = pd.read_csv(ruta, dtype=str, encoding="utf-8-sig")
                if df.empty or (len(df.columns) == 1 and "Sin datos" in str(df.columns[0])):
                    continue
                dfs.append(df)
            except Exception as e:
                print(f"[APILAR] Aviso al leer {ruta}: {e}")

        if not dfs:
            pd.DataFrame({"Sin datos": []}).to_csv(ruta_final, index=False, encoding="utf-8-sig")
            return 0, 0

        df_total    = pd.concat(dfs, ignore_index=True)
        total_filas = len(df_total)
        df_total.to_csv(ruta_final, index=False, encoding="utf-8-sig")
        return len(dfs), total_filas
