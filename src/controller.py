"""
Validación Factura Global — Controller
======================================

Glue between the GUI (interfaz_GUI.py) and the business logic:

    download buttons    → Descargas_SAP.Gallo_FAGLL03 / Monivoi_ZLMXCOM
                          (one day, or a whole range as ONE SAP query → ONE
                          CSV), Descargas_SAP.VentaEIAP_FAGLL03 ('Solo Venta
                          EIAP': Gallo with the EIAP accounts, end date - 1
                          day) and, after Monivoi, Descargas_SAP.VBRL_ZSE16
                          (ONE query for the whole period, with the invoices
                          of the Monivoi file)
    'Consolidar' button → Consolidacion.consolidar_gallo_monivoi

The GUI knows nothing about SAP: it only exposes the on_* callbacks that this
class wires up.
"""

import os
import time
from datetime import datetime, timedelta
from tkinter import messagebox, filedialog

import pandas as pd

from Descargas_SAP import (
    Gallo_FAGLL03,
    Monivoi_ZLMXCOM,
    VentaEIAP_FAGLL03, CUENTAS_EIAP,
    VBRL_ZSE16,        crear_archivo_sin_datos,
)
from Consolidacion import consolidar_gallo_monivoi, leer_facturas_monivoi
from utils import FORMATO_FECHA, parse_fecha, fecha_efectiva_eiap


class ValidacionFacturaController:
    """Connects the GUI with the SAP downloads (Gallo, Monivoi, VBRL) and the consolidation."""

    def __init__(self, gui):
        self.gui = gui

        # --- Transaction profiles ---
        # Which SAP function to use. All of them share the same signature
        # (DateFrom, DateTo, FolderPath, FileName, sociedad); their extras
        # (cuentas/variante in Gallo, vkorg/vtweg/rfc in Monivoi, the EIAP
        # accounts and the -1 day on the end date in Venta EIAP) live inside the
        # functions, so the controller never touches them. A single day or a
        # whole range is always ONE call → ONE file.
        #
        # Optional keys: "validar" (config → error text or None; checked BEFORE
        # the confirmation) and "nota" (config → extra text for the
        # confirmation dialog).
        self._TX_GALLO = {
            "clave":    "gallo",
            "nombre":   "Gallo (FAGLL03)",
            "descarga": Gallo_FAGLL03,
        }
        self._TX_MONIVOI = {
            "clave":    "monivoi",
            "nombre":   "Monivoi (ZLMXCOM_TRN_MONINVOI)",
            "descarga": Monivoi_ZLMXCOM,
        }
        # 'Solo Venta EIAP': Gallo with the four EIAP accounts, only for a range
        # ("Varios días"), querying SAP up to the end date MINUS ONE DAY.
        self._TX_EIAP = {
            "clave":    "eiap",
            "nombre":   "Venta EIAP (FAGLL03, cuentas EIAP)",
            "descarga": VentaEIAP_FAGLL03,
            "validar":  self._validar_eiap,
            "nota":     self._nota_eiap,
        }
        # VBRL is NOT downloaded by date: it is ONE query for the whole period,
        # with the invoices of the Monivoi file of THIS run. So instead of
        # "descarga" it has its own routine ("ejecutar") and says
        # what it depends on ("requiere"): the engine refuses to run it unless
        # Monivoi runs before it.
        self._TX_VBRL = {
            "clave":    "vbrl",
            "nombre":   "VBRL (ZSE16)",
            "ejecutar": self._descargar_vbrl,
            "requiere": "monivoi",
        }

        # Re-entrancy lock: SAP GUI can only handle one script run at a time.
        self._ocupado = False

        # Four download buttons (Gallo only, Monivoi only, Venta EIAP only, all
        # three of the main flow) share the SAME engine; only the list of
        # transactions to run changes.
        self.gui.on_download_gallo   = self.execute_download_gallo
        self.gui.on_download_monivoi = self.execute_download_monivoi
        self.gui.on_download_eiap    = self.execute_download_eiap
        self.gui.on_download_todos   = self.execute_download_todos
        # Separate button to consolidate the CSVs that were already downloaded.
        self.gui.on_consolidar = self.execute_consolidacion

    # =================================================================
    # DOWNLOAD ENTRY POINTS (one per button)
    # =================================================================
    def execute_download_gallo(self):
        self._execute_download([self._TX_GALLO], "Gallo")

    def execute_download_monivoi(self):
        self._execute_download([self._TX_MONIVOI], "Monivoi")

    def execute_download_eiap(self):
        self._execute_download([self._TX_EIAP], "Venta EIAP")

    def execute_download_todos(self):
        self._execute_download(
            [self._TX_GALLO, self._TX_MONIVOI, self._TX_VBRL], "Gallo + Monivoi + VBRL"
        )

    # =================================================================
    # DOWNLOAD ENGINE: runs the transactions in order, one single summary
    # =================================================================
    def _execute_download(self, txs, etiqueta):
        """
        Run the transactions in `txs` one after another and show ONE summary.
        `etiqueta` is the short label used in dialog titles and the status bar
        ("Gallo", "Monivoi", "Gallo + Monivoi + VBRL").
        """
        if not txs:
            raise ValueError("No se indicó ninguna transacción a descargar.")
        # A transaction that depends on another one (VBRL needs the Monivoi file
        # of THIS run) is only valid if that one runs earlier in the same list;
        # otherwise it could pick up a stale file from an earlier run.
        for pos, tx in enumerate(txs):
            requiere = tx.get("requiere")
            if requiere and requiere not in [t["clave"] for t in txs[:pos]]:
                raise ValueError(
                    f"{tx['nombre']} necesita que '{requiere}' se descargue antes, en la misma corrida."
                )
        if self._ocupado:
            return
        if not self.gui.validate_dates():
            return

        base = self.gui.get_config(txs[0]["clave"])   # shared dates/company code/mode
        mode = base["mode"]

        # Transaction-specific checks BEFORE asking for confirmation, so a request
        # that cannot run never reaches SAP (e.g. 'Solo Venta EIAP' needs a range).
        for tx in txs:
            validar = tx.get("validar")
            problema = validar(self.gui.get_config(tx["clave"])) if validar else None
            if problema:
                messagebox.showerror(f"Descarga de {tx['nombre']}", problema)
                return

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
            if any(tx["clave"] == "vbrl" for tx in txs):
                detalle += (
                    "VBRL consulta las facturas del archivo de Monivoi de este periodo.\n"
                    "Mientras corre NO copies nada: el portapapeles se usa para pasarle\n"
                    "la lista de facturas a SAP.\n\n"
                )
        # Transaction-specific notes (e.g. the period SAP will really be queried with).
        for tx in txs:
            nota = tx.get("nota")
            if nota:
                detalle += nota(self.gui.get_config(tx["clave"]))
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
        resultados = {}          # tx key → "datos" | "vacio" | "error" (what VBRL checks on Monivoi)
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
                # Each transaction gets its own try: if one blows up, the other
                # one still runs. At the end, one single summary says what happened.
                try:
                    config = self.gui.get_config(tx["clave"])   # the right filename for each tx
                    inicio_tx = time.time()
                    if "ejecutar" in tx:
                        # Not downloaded by date (VBRL): the transaction has its own routine.
                        estado, linea = tx["ejecutar"](tx, config, resultados)
                        resultados[tx["clave"]] = estado
                        resumen.append(linea)
                        if estado != "datos":
                            todo_ok = False
                    else:
                        # One day or a whole range: ONE SAP query → ONE file.
                        con_datos = self._run_download(tx, config)
                        if con_datos:
                            resultados[tx["clave"]] = "datos"
                            resumen.append(f"✅ {tx['nombre']}: {config['filename']}")
                        else:
                            todo_ok = False
                            resultados[tx["clave"]] = self._clasificar_sin_datos(
                                os.path.join(config["input_path"], config["filename"]), inicio_tx
                            )
                            resumen.append(
                                f"⚠️ {tx['nombre']}: sin datos o error (ver consola)"
                            )
                except Exception as e:
                    todo_ok = False
                    resultados[tx["clave"]] = "error"
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
    # CONSOLIDATION: merges the downloaded CSVs into one Excel workbook
    # =================================================================
    def execute_consolidacion(self):
        """
        Take the Gallo and Monivoi CSVs (named after the GUI dates), plus the
        VBRL one when there is one, and build the consolidated workbook (see
        Consolidacion.py for its sheets). If a Gallo/Monivoi CSV is not where it
        is expected, offer to pick it manually; VBRL is optional.
        """
        if self._ocupado:
            return
        if not self.gui.validate_dates():
            return

        cfg_gallo   = self.gui.get_config("gallo")
        cfg_monivoi = self.gui.get_config("monivoi")
        cfg_vbrl    = self.gui.get_config("vbrl")
        input_path  = cfg_gallo["input_path"]

        # --- Locate the files: automatically by date, manual fallback ---
        ruta_gallo = self._localizar_archivo(input_path, cfg_gallo["filename"], "Gallo")
        if not ruta_gallo:
            return
        ruta_monivoi = self._localizar_archivo(input_path, cfg_monivoi["filename"], "Monivoi")
        if not ruta_monivoi:
            return
        # VBRL is optional: without it the workbook comes out as before.
        ruta_vbrl, cancelar = self._resolver_vbrl(input_path, cfg_vbrl["filename"], ruta_monivoi)
        if cancelar:
            return

        # --- Dates picked by the user (for 'Venta DSD') ---
        # One in single-day mode; EVERY day of the period in range mode. This
        # way those days always show up in the sales matrices even without
        # amounts, and the module can break ties on the CSV date format.
        fechas_seleccionadas = self._fechas_seleccionadas(cfg_gallo)

        # --- Output: consolidation folder chosen by the user in the GUI ---
        try:
            ruta_output = self.gui.get_consolidacion_path()
        except OSError as e:
            messagebox.showerror(
                "Error", f"No se pudo usar la carpeta de consolidación:\n{e}"
            )
            return
        nombre_salida = f"Consolidado_{os.path.splitext(cfg_gallo['filename'])[0]}.xlsx"

        linea_vbrl = (
            f"  🧾 VBRL    : {os.path.basename(ruta_vbrl)}\n" if ruta_vbrl
            else "  🧾 VBRL    : (sin VBRL: no se generan 'VBRL' ni 'TD.Diferencias_Referencias')\n"
        )
        confirm = messagebox.askyesno(
            "Confirmar consolidación",
            f"¿Consolidar estos archivos en un solo Excel?\n\n"
            f"  🐓 Gallo   : {os.path.basename(ruta_gallo)}\n"
            f"  📊 Monivoi : {os.path.basename(ruta_monivoi)}\n"
            f"{linea_vbrl}\n"
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
                ruta_vbrl=ruta_vbrl,
            )

            nota_monivoi = (
                "\n⚠️ El archivo de Monivoi venía sin datos; su pestaña quedó vacía."
                if resumen["monivoi_sin_datos"] else ""
            )
            # Warnings from the sales sheets (rows without a date, dates outside
            # the chosen period), from 'TD' (rows without Centro), from 'MTD',
            # from the TD × MTD cross and, with a VBRL file, from VBRL and from
            # 'TD.Diferencias_Referencias'. They are not errors: they are things
            # the stakeholder must see before comparing against the monitor.
            avisos = (
                (resumen.get("venta_avisos") or [])
                + (resumen.get("td_avisos") or [])
                + (resumen.get("mtd_avisos") or [])
                + (resumen.get("cruce_avisos") or [])
                + (resumen.get("vbrl_avisos") or [])
                + (resumen.get("td_referencias_avisos") or [])
            )
            nota_ventas = ("\n" + "\n".join(f"⚠️ {a}" for a in avisos)) if avisos else ""

            lineas_vbrl = ""
            if resumen.get("vbrl_incluido"):
                lineas_vbrl = (
                    f"\n  • Filas VBRL               : {resumen['filas_vbrl']:,}"
                    f"\n  • Referencias (facturas)   : {resumen['filas_td_referencias']:,}"
                    f"\n      …sin timbre            : {resumen['td_referencias_sin_timbre']:,}"
                )
                if resumen["vbrl_sin_datos"]:
                    lineas_vbrl += "\n⚠️ El archivo de VBRL venía sin datos; su pestaña quedó vacía."

            self.gui.set_status("✅ ¡Consolidación completada!")
            messagebox.showinfo(
                "Consolidación completada",
                f"Archivo generado:\n{resumen['ruta']}\n\n"
                f"  • Filas Gallo              : {resumen['filas_gallo']:,}\n"
                f"  • Filas Monivoi            : {resumen['filas_monivoi']:,}\n"
                f"  • Referencias únicas       : {resumen['refs_unicas']:,}\n"
                f"  • Docs 'MX Commercial'     : {resumen['filas_catalogo']:,}\n"
                f"  • Globales (filtro final)  : {resumen['filas_globales']:,}\n"
                f"  • Σ Importe en moneda local: {resumen['total_importe_gallo']:,.2f}"
                f"{lineas_vbrl}{nota_monivoi}{nota_ventas}"
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

    def _resolver_vbrl(self, input_path, filename, ruta_monivoi):
        """
        Locate the VBRL CSV for the consolidation. VBRL is OPTIONAL: without it
        the workbook comes out as before (no 'VBRL' and no
        'TD.Diferencias_Referencias' sheet).

        Returns (ruta, cancelar): `ruta` None means "consolidate without VBRL";
        `cancelar` True means "do not consolidate at all".
        """
        ruta = os.path.join(input_path, filename)
        if not os.path.exists(ruta):
            decision = messagebox.askyesnocancel(
                "Archivo de VBRL no encontrado",
                f"No encontré el archivo esperado:\n\n  {filename}\n\n"
                f"en la carpeta:\n  {input_path}\n\n"
                "Sin VBRL el Excel sale sin las pestañas 'VBRL' y 'TD.Diferencias_Referencias'.\n\n"
                "Sí = consolidar SIN VBRL\n"
                "No = seleccionar el archivo manualmente\n"
                "Cancelar = no consolidar",
            )
            if decision is None:
                return None, True
            if decision:
                return None, False
            ruta = filedialog.askopenfilename(
                title="Selecciona el CSV de VBRL",
                initialdir=input_path,
                filetypes=[("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")],
            )
            if not ruta:
                return None, True

        # Staleness guard: VBRL is built from the invoices of the Monivoi file,
        # so a VBRL OLDER than that Monivoi (e.g. Monivoi was downloaded again
        # afterwards) may not cover its invoices and would leave extra 'Sin timbre'.
        try:
            vbrl_anterior = os.path.getmtime(ruta) < os.path.getmtime(ruta_monivoi) - 2
        except OSError:
            vbrl_anterior = False
        if vbrl_anterior:
            decision = messagebox.askyesnocancel(
                "VBRL anterior al archivo de Monivoi",
                f"El archivo de VBRL:\n  {os.path.basename(ruta)}\n"
                f"se descargó ANTES que el de Monivoi:\n  {os.path.basename(ruta_monivoi)}\n\n"
                "Probablemente no cubre las facturas actuales y saldrían 'Sin timbre' de más.\n\n"
                "Sí = usarlo de todos modos\n"
                "No = consolidar SIN VBRL\n"
                "Cancelar = no consolidar",
            )
            if decision is None:
                return None, True
            if not decision:
                return None, False
        return ruta, False

    def _fechas_seleccionadas(self, config):
        """
        List of 'DD.MM.YYYY' dates the user picked in the GUI: a single one in
        single-day mode, every day of the period in range mode.
        """
        if config["mode"] == "single":
            return [config["fecha"]]
        return self._iterar_dias(config["date_from"], config["date_to"])

    def _localizar_archivo(self, input_path, filename, etiqueta):
        """
        Look for the CSV with the expected name (from the GUI dates). If it is
        not there, ask whether to pick it manually with the file browser.
        Returns the path found/picked, or None to cancel.
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
    # DOWNLOAD CORES (no confirmation or popup: they just run)
    # =================================================================
    def _descargar_vbrl(self, tx, config, resultados):
        """
        VBRL step of the download engine: ONE query for the whole period, with
        the invoices (distinct 'Factura' values) of the Monivoi file that was
        just downloaded in this same run.

        Returns (estado, summary line); estado is "datos", "vacio" or "error".
        """
        if resultados.get("monivoi") == "error":
            return "error", (
                f"⚠️ {tx['nombre']}: omitido — Monivoi falló y no hay facturas que consultar"
            )

        cfg_monivoi  = self.gui.get_config("monivoi")
        ruta_monivoi = os.path.join(cfg_monivoi["input_path"], cfg_monivoi["filename"])
        ruta_vbrl    = os.path.join(config["input_path"], config["filename"])

        facturas, filas, col_factura = leer_facturas_monivoi(ruta_monivoi)
        if filas == 0:
            # Monivoi of the period is the 'Sin datos' placeholder: there is
            # nothing to ask SAP. Leave VBRL as a placeholder too, so the three
            # files of the period agree and no old VBRL can pass for this one.
            crear_archivo_sin_datos(ruta_vbrl)
            return "vacio", f"⚠️ {tx['nombre']}: Monivoi sin datos → {config['filename']} (vacío)"
        if not facturas:
            return "error", (
                f"⚠️ {tx['nombre']}: Monivoi trae {filas:,} fila(s) pero ninguna factura "
                f"utilizable en la columna '{col_factura}'"
            )

        self.gui.set_status(
            f"⏳ Descargando {tx['nombre']}: {len(facturas):,} factura(s)... (no cierre SAP)"
        )
        inicio = time.time()
        if VBRL_ZSE16(facturas, FolderPath=config["input_path"], FileName=config["filename"]):
            return "datos", f"✅ {tx['nombre']}: {len(facturas):,} factura(s) → {config['filename']}"
        if self._clasificar_sin_datos(ruta_vbrl, inicio) == "vacio":
            return "vacio", (
                f"⚠️ {tx['nombre']}: SAP no devolvió entradas para esas facturas "
                f"→ {config['filename']} (vacío)"
            )
        return "error", f"⚠️ {tx['nombre']}: error (ver consola)"

    @staticmethod
    def _clasificar_sin_datos(ruta, desde):
        """
        A download returned False: was it "SAP had no rows" or an error? The
        download functions create the 'Sin datos' placeholder only in the first
        case, so look for a placeholder written since `desde` (epoch seconds).
        Returns "vacio" or "error".
        """
        try:
            # 2 s of slack: some file systems store modification times in 2 s steps.
            if os.path.getmtime(ruta) < desde - 2:
                return "error"            # an older file: not this run's output
            df = pd.read_csv(ruta, dtype=str, encoding="utf-8-sig", nrows=1)
        except Exception:
            return "error"
        sin_datos = df.empty or (len(df.columns) == 1 and "Sin datos" in str(df.columns[0]))
        return "vacio" if sin_datos else "error"

    def _run_download(self, tx, config):
        """
        Download a transaction for the period in the GUI as ONE SAP query → ONE
        file: a single day (from = to) or the whole range ("Varios días"). There
        are no per-day files any more.
        Returns True (with data) / False (no data or a handled error).
        """
        if config["mode"] == "single":
            desde = hasta = config["fecha"]
        else:
            desde, hasta = config["date_from"], config["date_to"]
        return tx["descarga"](
            DateFrom=desde,
            DateTo=hasta,
            FolderPath=config["input_path"],
            FileName=config["filename"],
            sociedad=config["sociedad"],
        )

    # =================================================================
    # HELPERS (transaction-agnostic)
    # =================================================================
    @staticmethod
    def _iterar_dias(desde, hasta):
        """Return the list of 'DD.MM.YYYY' dates between desde and hasta (inclusive)."""
        d = datetime.strptime(desde, "%d.%m.%Y")
        h = datetime.strptime(hasta, "%d.%m.%Y")
        dias = []
        while d <= h:
            dias.append(d.strftime("%d.%m.%Y"))
            d += timedelta(days=1)
        return dias

    # =================================================================
    # 'SOLO VENTA EIAP' helpers
    # =================================================================
    @staticmethod
    def _validar_eiap(config):
        """
        'Solo Venta EIAP' only works with a range ("Varios días"), and SAP is
        queried up to the end date MINUS ONE DAY (utils.fecha_efectiva_eiap), so
        the range must still make sense after that.

        Returns an error text, or None when the request can run.
        """
        if config["mode"] != "range":
            return (
                "'Solo Venta EIAP' trabaja con el intervalo \"Varios días\".\n\n"
                "Cambia el 'Tipo de intervalo' a \"Varios días\" y captura Fecha Desde y "
                "Fecha Hasta (a la Fecha Hasta se le resta 1 día en la consulta)."
            )
        desde = parse_fecha(config["date_from"])
        hasta = fecha_efectiva_eiap(config["date_to"])
        if hasta < desde:
            return (
                f"Con Fecha Hasta {config['date_to']} la consulta terminaría el "
                f"{hasta.strftime(FORMATO_FECHA)}, ANTES de la Fecha Desde "
                f"({config['date_from']}).\n\n"
                "En 'Solo Venta EIAP' la Fecha Hasta se resta 1 día, así que el periodo "
                "debe abarcar al menos 2 días (ej. 01.01.2026 a 02.01.2026 consulta sólo el 01.01.2026)."
            )
        return None

    @staticmethod
    def _nota_eiap(config):
        """Confirmation text: the period SAP will REALLY be queried with, and the accounts."""
        hasta = fecha_efectiva_eiap(config["date_to"]).strftime(FORMATO_FECHA)
        return (
            f"Consulta en SAP : {config['date_from']} — {hasta}\n"
            "(a la Fecha Hasta se le resta 1 día, sólo en este botón)\n"
            f"Cuentas         : {', '.join(CUENTAS_EIAP)}\n\n"
        )
