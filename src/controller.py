import os
import time
from datetime import datetime, timedelta
from tkinter import messagebox

import pandas as pd

from Descargas_SAP import (
    Gallo_FAGLL03,   nombre_archivo_dia,          nombre_archivo_rango,
    Monivoi_ZLMXCOM, nombre_archivo_dia_monivoi,  nombre_archivo_rango_monivoi,
)


class ValidacionFacturaController:
    """Conecta la GUI con la lógica de descarga SAP (Gallo y Monivoi)."""

    def __init__(self, gui):
        self.gui = gui

        # --- Perfiles de transacción ---
        # Qué función SAP usar y cómo nombrar los archivos por día. Ambas
        # funciones comparten firma (DateFrom, DateTo, FolderPath, FileName,
        # sociedad); sus extras (cuentas/variante en Gallo, vkorg/vtweg/rfc en
        # Monivoi) viajan como defaults, así que el controller no los toca.
        self._TX_GALLO = {
            "clave":      "gallo",
            "nombre":     "Gallo (FAGLL03)",
            "descarga":   Gallo_FAGLL03,
            "nombre_dia": nombre_archivo_dia,
        }
        self._TX_MONIVOI = {
            "clave":      "monivoi",
            "nombre":     "Monivoi (ZLMXCOM_TRN_MONINVOI)",
            "descarga":   Monivoi_ZLMXCOM,
            "nombre_dia": nombre_archivo_dia_monivoi,
        }

        # Un solo botón dispara AMBAS transacciones en secuencia.
        self.gui.on_download_ambos = self.execute_download_ambos

    # =================================================================
    # ENTRADA ÚNICA: corre Gallo y luego Monivoi, con un solo resumen
    # =================================================================
    def execute_download_ambos(self):
        if not self.gui.validate_dates():
            return

        txs  = [self._TX_GALLO, self._TX_MONIVOI]
        base = self.gui.get_config(self._TX_GALLO["clave"])   # fechas/sociedad/modo comunes
        mode = base["mode"]

        periodo = (
            f"Fecha    : {base['fecha']}" if mode == "single"
            else f"Periodo  : {base['date_from']} — {base['date_to']}"
        )
        confirm = messagebox.askyesno(
            "Confirmar descarga (Gallo + Monivoi)",
            f"¿Iniciar la descarga de AMBAS transacciones?\n\n"
            f"Sociedad : {base['sociedad']}\n"
            f"{periodo}\n\n"
            f"Se ejecutan una tras otra:\n"
            f"  1) {self._TX_GALLO['nombre']}\n"
            f"  2) {self._TX_MONIVOI['nombre']}\n\n"
            f"Se guardará en:\n{base['input_path']}"
        )
        if not confirm:
            return

        os.makedirs(base["input_path"], exist_ok=True)
        self.gui.disable_buttons()

        resumen = []
        try:
            for i, tx in enumerate(txs, 1):
                config = self.gui.get_config(tx["clave"])   # filename correcto por tx
                self.gui.set_status(
                    f"⏳ ({i}/{len(txs)}) Descargando {tx['nombre']}... (no cierre SAP)"
                )
                # Cada transacción va en su propio try: si una truena, la otra
                # igual se ejecuta. Al final, un solo resumen dice qué pasó.
                try:
                    if config["mode"] == "single":
                        con_datos = self._run_single_download(tx, config)
                        resumen.append(
                            f"✅ {tx['nombre']}: {config['filename']}" if con_datos
                            else f"⚠️ {tx['nombre']}: sin datos o error (ver consola)"
                        )
                    else:
                        dias = self._iterar_dias(config["date_from"], config["date_to"])
                        n_dias, total_filas = self._run_range_download(tx, config, dias)
                        resumen.append(
                            f"✅ {tx['nombre']}: {n_dias} día(s), "
                            f"{total_filas:,} filas → {config['filename']}"
                        )
                except Exception as e:
                    resumen.append(f"❌ {tx['nombre']}: error — {e}")

            self.gui.set_status("✅ ¡Descarga de Gallo + Monivoi completada!")
            messagebox.showinfo(
                "Resumen de descarga",
                "Proceso terminado:\n\n" + "\n".join(resumen) +
                f"\n\nRuta: {base['input_path']}"
            )
        except Exception as e:
            self.gui.set_status("❌ Error en la descarga")
            messagebox.showerror("Error", f"Ocurrió un error inesperado:\n\n{e}")
        finally:
            self.gui.enable_buttons()
            if "completada" not in self.gui.status_var.get():
                self.gui.set_status("✓ Listo para comenzar")

    # =================================================================
    # NÚCLEOS DE DESCARGA (sin confirm ni popup: sólo ejecutan)
    # =================================================================
    def _run_single_download(self, tx, config):
        """Descarga UN día de una transacción.
        Devuelve True (con datos) / False (sin datos o error controlado)."""
        fecha = config["fecha"]
        return tx["descarga"](
            DateFrom=fecha,
            DateTo=fecha,
            FolderPath=config["input_path"],
            FileName=config["filename"],
            sociedad=config["sociedad"],
        )

    def _run_range_download(self, tx, config, dias):
        input_path = config["input_path"]
        sociedad   = config["sociedad"]

        # Subcarpeta donde se guardan los archivos por día (se conservan)
        nombre_consolidado = config["filename"]
        nombre_subcarpeta  = "Dias_" + os.path.splitext(nombre_consolidado)[0]
        carpeta_dias = os.path.join(input_path, nombre_subcarpeta)
        os.makedirs(carpeta_dias, exist_ok=True)

        archivos_con_datos = []
        total = len(dias)

        for i, dia in enumerate(dias, 1):
            self.gui.set_status(f"📥 {tx['nombre']} — día {i}/{total}: {dia}... (no cierre SAP)")
            fname = tx["nombre_dia"](dia)
            ruta_dia = os.path.join(carpeta_dias, fname)

            con_datos = tx["descarga"](
                DateFrom=dia,
                DateTo=dia,
                FolderPath=carpeta_dias,
                FileName=fname,
                sociedad=sociedad,
            )
            time.sleep(3)  # respiro entre descargas para que SAP libere la pantalla

            if con_datos and os.path.exists(ruta_dia):
                archivos_con_datos.append(ruta_dia)
                print(f"[RANGO {tx['clave'].upper()}] {dia}: con datos → {ruta_dia}")
            else:
                print(f"[RANGO {tx['clave'].upper()}] {dia}: sin datos → se omite")

        # Apilar todos los días con datos en un solo CSV (una sola tabla)
        self.gui.set_status("📋 Apilando todos los días en un solo CSV...")
        ruta_consolidado = os.path.join(input_path, nombre_consolidado)
        n_dias, total_filas = self._apilar_a_csv(archivos_con_datos, ruta_consolidado)
        print(
            f"[RANGO {tx['clave'].upper()}] Consolidado: {n_dias} día(s), "
            f"{total_filas} filas → {ruta_consolidado}"
        )
        return n_dias, total_filas

    # =================================================================
    # HELPERS (agnósticos de transacción)
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
        Lee todos los CSV por día, los concatena en UNA SOLA tabla y la guarda
        en CSV (utf-8-sig). Ignora placeholders 'Sin datos'.
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
