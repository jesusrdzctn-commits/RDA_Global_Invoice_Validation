import os
import time
from datetime import datetime, timedelta
from tkinter import messagebox, filedialog

import pandas as pd

from Descargas_SAP import (
    Gallo_FAGLL03,   nombre_archivo_dia,          nombre_archivo_rango,
    Monivoi_ZLMXCOM, nombre_archivo_dia_monivoi,  nombre_archivo_rango_monivoi,
)
from Consolidacion import consolidar_gallo_monivoi


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

        # Re-entrancy lock: SAP GUI can only handle one script run at a time.
        self._ocupado = False

        # Three download buttons (Gallo only, Monivoi only, both) share the
        # SAME engine; only the list of transactions to run changes.
        self.gui.on_download_gallo   = self.execute_download_gallo
        self.gui.on_download_monivoi = self.execute_download_monivoi
        self.gui.on_download_ambos   = self.execute_download_ambos
        # Botón aparte para consolidar los CSV ya descargados.
        self.gui.on_consolidar = self.execute_consolidacion

    # =================================================================
    # DOWNLOAD ENTRY POINTS (one per button)
    # =================================================================
    def execute_download_gallo(self):
        self._execute_download([self._TX_GALLO], "Gallo")

    def execute_download_monivoi(self):
        self._execute_download([self._TX_MONIVOI], "Monivoi")

    def execute_download_ambos(self):
        self._execute_download([self._TX_GALLO, self._TX_MONIVOI], "Gallo + Monivoi")

    # =================================================================
    # DOWNLOAD ENGINE: runs the transactions in order, one single summary
    # =================================================================
    def _execute_download(self, txs, etiqueta):
        """
        Run the transactions in `txs` one after another and show ONE summary.
        `etiqueta` is the short label used in dialog titles and the status bar
        ("Gallo", "Monivoi", "Gallo + Monivoi").
        """
        if not txs:
            raise ValueError("No se indicó ninguna transacción a descargar.")
        if self._ocupado:
            return
        if not self.gui.validate_dates():
            return

        base = self.gui.get_config(txs[0]["clave"])   # fechas/sociedad/modo comunes
        mode = base["mode"]

        periodo = (
            f"Fecha    : {base['fecha']}" if mode == "single"
            else f"Periodo  : {base['date_from']} — {base['date_to']}"
        )
        if len(txs) == 1:
            pregunta = f"¿Iniciar la descarga de {txs[0]['nombre']}?"
            detalle  = "Sólo se ejecuta esta transacción.\n\n"
        else:
            pregunta = (
                "¿Iniciar la descarga de AMBAS transacciones?" if len(txs) == 2
                else f"¿Iniciar la descarga de las {len(txs)} transacciones?"
            )
            detalle = (
                "Se ejecutan una tras otra:\n"
                + "".join(f"  {i}) {tx['nombre']}\n" for i, tx in enumerate(txs, 1))
                + "\n"
            )
        confirm = messagebox.askyesno(
            f"Confirmar descarga ({etiqueta})",
            f"{pregunta}\n\n"
            f"Sociedad : {base['sociedad']}\n"
            f"{periodo}\n\n"
            f"{detalle}"
            f"Se guardará en:\n{base['input_path']}"
        )
        if not confirm:
            return

        self._ocupado = True
        resumen = []
        todo_ok = True           # False if any transaction fails or has no data
        estado_final = None      # status bar message to keep once finished
        try:
            self.gui.disable_buttons()
            os.makedirs(base["input_path"], exist_ok=True)
            for i, tx in enumerate(txs, 1):
                prefijo = f"({i}/{len(txs)}) " if len(txs) > 1 else ""
                self.gui.set_status(
                    f"⏳ {prefijo}Descargando {tx['nombre']}... (no cierre SAP)"
                )
                # Cada transacción va en su propio try: si una truena, la otra
                # igual se ejecuta. Al final, un solo resumen dice qué pasó.
                try:
                    config = self.gui.get_config(tx["clave"])   # filename correcto por tx
                    if config["mode"] == "single":
                        con_datos = self._run_single_download(tx, config)
                        if con_datos:
                            resumen.append(f"✅ {tx['nombre']}: {config['filename']}")
                        else:
                            todo_ok = False
                            resumen.append(
                                f"⚠️ {tx['nombre']}: sin datos o error (ver consola)"
                            )
                    else:
                        dias = self._iterar_dias(config["date_from"], config["date_to"])
                        n_dias, total_filas = self._run_range_download(tx, config, dias)
                        if n_dias:
                            resumen.append(
                                f"✅ {tx['nombre']}: {n_dias} de {len(dias)} día(s) "
                                f"con datos, {total_filas:,} filas → {config['filename']}"
                            )
                        else:
                            todo_ok = False
                            resumen.append(
                                f"⚠️ {tx['nombre']}: ningún día del periodo trajo datos "
                                f"→ {config['filename']} (vacío)"
                            )
                except Exception as e:
                    todo_ok = False
                    resumen.append(f"❌ {tx['nombre']}: error — {e}")

            estado_final = (
                f"✅ ¡Descarga de {etiqueta} completada!" if todo_ok
                else f"⚠️ Descarga de {etiqueta} terminada con avisos"
            )
            self.gui.set_status(estado_final)
            mostrar = messagebox.showinfo if todo_ok else messagebox.showwarning
            mostrar(
                f"Resumen de descarga ({etiqueta})",
                "Proceso terminado:\n\n" + "\n".join(resumen) +
                f"\n\nRuta: {base['input_path']}"
            )
        except Exception as e:
            estado_final = None   # unexpected error: status bar goes back to "Listo"
            self.gui.set_status("❌ Error en la descarga")
            messagebox.showerror("Error", f"Ocurrió un error inesperado:\n\n{e}")
        finally:
            self._ocupado = False
            self.gui.enable_buttons()
            if estado_final is None:
                self.gui.set_status("✓ Listo para comenzar")

    # =================================================================
    # CONSOLIDACIÓN: junta los CSV descargados en un solo Excel
    # =================================================================
    def execute_consolidacion(self):
        """
        Toma los CSV de Gallo y Monivoi (según las fechas de la GUI) y genera
        el Excel consolidado con la pestaña 'Validación - Doc.'.
        Si algún CSV no está donde se espera, ofrece elegirlo manualmente.
        """
        if self._ocupado:
            return
        if not self.gui.validate_dates():
            return

        cfg_gallo   = self.gui.get_config("gallo")
        cfg_monivoi = self.gui.get_config("monivoi")
        input_path  = cfg_gallo["input_path"]

        # --- Localizar archivos: automático por fechas, respaldo manual ---
        ruta_gallo = self._localizar_archivo(input_path, cfg_gallo["filename"], "Gallo")
        if not ruta_gallo:
            return
        ruta_monivoi = self._localizar_archivo(input_path, cfg_monivoi["filename"], "Monivoi")
        if not ruta_monivoi:
            return

        # --- Fechas elegidas por el usuario (para 'Venta DSD' y 'Venta EIAP') ---
        # En modo día es una sola; en rango, TODOS los días del periodo. Así esos
        # días siempre aparecen en las matrices de venta aunque no traigan
        # importes, y el módulo puede desempatar el formato de fecha del CSV.
        fechas_seleccionadas = self._fechas_seleccionadas(cfg_gallo)

        # --- Salida: carpeta de consolidación elegida por el usuario en la GUI ---
        try:
            ruta_output = self.gui.get_consolidacion_path()
        except OSError as e:
            messagebox.showerror(
                "Error", f"No se pudo usar la carpeta de consolidación:\n{e}"
            )
            return
        nombre_salida = f"Consolidado_{os.path.splitext(cfg_gallo['filename'])[0]}.xlsx"

        confirm = messagebox.askyesno(
            "Confirmar consolidación",
            f"¿Consolidar estos archivos en un solo Excel?\n\n"
            f"  🐓 Gallo   : {os.path.basename(ruta_gallo)}\n"
            f"  📊 Monivoi : {os.path.basename(ruta_monivoi)}\n\n"
            f"Se generará:\n  {nombre_salida}\n\n"
            f"En la carpeta:\n  {ruta_output}"
        )
        if not confirm:
            return

        self._ocupado = True
        try:
            self.gui.disable_buttons()
            resumen = consolidar_gallo_monivoi(
                ruta_gallo, ruta_monivoi, ruta_output,
                nombre_salida=nombre_salida,
                callback_status=self.gui.set_status,
                fechas_seleccionadas=fechas_seleccionadas,
            )

            nota_monivoi = (
                "\n⚠️ El archivo de Monivoi venía sin datos; su pestaña quedó vacía."
                if resumen["monivoi_sin_datos"] else ""
            )
            # Avisos de las pestañas de venta (filas sin fecha, fechas fuera del
            # periodo elegido). No son errores: son cosas que el stakeholder
            # debe ver antes de comparar contra el monitor.
            avisos = resumen.get("venta_avisos") or []
            nota_ventas = ("\n" + "\n".join(f"⚠️ {a}" for a in avisos)) if avisos else ""

            self.gui.set_status("✅ ¡Consolidación completada!")
            messagebox.showinfo(
                "Consolidación completada",
                f"Archivo generado:\n{resumen['ruta']}\n\n"
                f"  • Filas Gallo              : {resumen['filas_gallo']:,}\n"
                f"  • Filas Monivoi            : {resumen['filas_monivoi']:,}\n"
                f"  • Referencias únicas       : {resumen['refs_unicas']:,}\n"
                f"  • Docs 'MX Commercial'     : {resumen['filas_catalogo']:,}\n"
                f"  • Globales (filtro final)  : {resumen['filas_globales']:,}\n"
                f"  • Σ Importe en moneda local: {resumen['total_importe_gallo']:,.2f}\n"
                f"  • Σ Venta DSD              : {resumen['total_venta_dsd']:,.2f}\n"
                f"  • Σ Venta EIAP             : {resumen['total_venta_eiap']:,.2f}"
                f"{nota_monivoi}{nota_ventas}"
            )
        except (FileNotFoundError, ValueError) as e:
            self.gui.set_status("❌ Error en la consolidación")
            messagebox.showerror("Error en la consolidación", str(e))
        except Exception as e:
            self.gui.set_status("❌ Error en la consolidación")
            messagebox.showerror("Error", f"Ocurrió un error inesperado:\n\n{e}")
        finally:
            self._ocupado = False
            self.gui.enable_buttons()
            if "completada" not in self.gui.status_var.get():
                self.gui.set_status("✓ Listo para comenzar")

    def _fechas_seleccionadas(self, config):
        """
        Lista de fechas 'DD.MM.YYYY' que el usuario eligió en la GUI:
        una sola en modo día, todas las del periodo en modo rango.
        """
        if config["mode"] == "single":
            return [config["fecha"]]
        return self._iterar_dias(config["date_from"], config["date_to"])

    def _localizar_archivo(self, input_path, filename, etiqueta):
        """
        Busca el CSV con el nombre esperado (según fechas de la GUI). Si no
        está, pregunta si se quiere elegir manualmente con el explorador.
        Devuelve la ruta encontrada/elegida, o None para cancelar.
        """
        ruta = os.path.join(input_path, filename)
        if os.path.exists(ruta):
            return ruta

        buscar = messagebox.askyesno(
            f"Archivo de {etiqueta} no encontrado",
            f"No encontré el archivo esperado:\n\n  {filename}\n\n"
            f"en la carpeta:\n  {input_path}\n\n"
            f"¿Quieres seleccionarlo manualmente?"
        )
        if not buscar:
            return None

        ruta_manual = filedialog.askopenfilename(
            title=f"Selecciona el CSV de {etiqueta}",
            initialdir=input_path,
            filetypes=[("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")],
        )
        return ruta_manual or None

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
