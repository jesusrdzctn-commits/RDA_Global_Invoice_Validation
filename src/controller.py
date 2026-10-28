"""Validación Factura Global — controller: glue between the GUI (interfaz_GUI.py), the SAP
downloads (Descargas_SAP.py) and the consolidation (Consolidacion.py).

The main download button runs Gallo, Monivoi, VBRL and FB03 one after another (default) or,
with the GUI checkbox, IN PARALLEL by passes: Gallo ‖ Monivoi, then VBRL ‖ FB03 (VBRL needs
the Monivoi file and FB03 the Gallo file of pass 1).
After a consolidation, the 2nd run (VBRK, VBRP, VBRL 2nd run, FBL5N, KNVP, BKPF) pastes into SAP
what the #N/A rows of 'TD.Diferencias_Referencias' carry: one after another, or, with the same
checkbox, IN PARALLEL as two passes of THREE transactions, each one in its own SAP window
(PASADAS_2DA): VBRK ‖ VBRP ‖ VBRL 2nd run, then FBL5N ‖ KNVP ‖ BKPF (FBL5N and KNVP paste the
'Solic.' column of the VBRK downloaded in pass 1).
"""

import os
import time
from datetime import datetime, timedelta
from tkinter import messagebox, filedialog

import pandas as pd

from Descargas_SAP import (
    Gallo_FAGLL03, Monivoi_ZLMXCOM, VentaEIAP_FAGLL03, VentaDSD_FAGLL03,
    VBRL_ZSE16, FB03_Documentos, DescargaPorRevisarError,
    VBRP_ZSE16, VBRK_ZSE16, VBRL2_ZSE16, KNVP_ZSE16, BKPF_ZSE16, FBL5N_Deudores,
    Archivos, Excel, Paralelo, MAX_VENTANAS_SAP,
)
from Consolidacion import (
    consolidar_gallo_monivoi, consolidar_venta_dsd, consolidar_venta_eiap, leer_documentos_monivoi,
    leer_documentos_gallo_fb03, leer_solicitantes_vbrk, calcular_listas_segunda_corrida,
    EtapaDocumentosFB03,
)
from utils import (
    FORMATO_FECHA, parse_fecha, fecha_efectiva_eiap,
    CUENTAS_EIAP, CLASES_EIAP, CUENTAS_DSD, CLASES_DSD,
    nombre_archivo_dia_2da, nombre_archivo_rango_2da, CORRIDAS_DESCARGA,
)


def _fmt_duracion(segundos):
    """Seconds → '45 s', '3 min 05 s' or '1 h 02 min' (for the summaries)."""
    total = max(0, int(round(segundos)))
    horas, resto = divmod(total, 3600)
    minutos, seg = divmod(resto, 60)
    if horas:
        return f"{horas} h {minutos:02d} min"
    if minutos:
        return f"{minutos} min {seg:02d} s"
    return f"{seg} s"


class ValidacionFacturaController:
    """Connects the GUI with the SAP downloads (Gallo, Monivoi, VBRL, FB03, 2nd run) and the consolidation.
    Sections: download entry points → engines (sequential, parallel) → consolidation → 2nd run → sales
    buttons."""

    # 2nd run in parallel: two PASSES of three; a pass starts when the previous one is over.
    PASADAS_2DA = (("vbrk", "vbrp", "vbrl2"), ("fbl5n", "knvp", "bkpf"))

    def __init__(self, gui):
        self.gui = gui

        # Profiles: "descarga" (DateFrom, DateTo, FolderPath, FileName, sociedad), optional "validar" / "nota".
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
        self._TX_DSD = {
            "clave":    "dsd",
            "nombre":   "Venta DSD (FAGLL03)",
            "descarga": VentaDSD_FAGLL03,
            "validar":  lambda config: self._validar_venta("Venta DSD", config),
            "nota":     lambda config: self._nota_venta(config, CUENTAS_DSD, CLASES_DSD),
        }
        self._TX_EIAP = {
            "clave":    "eiap",
            "nombre":   "Venta EIAP (FAGLL03)",
            "descarga": VentaEIAP_FAGLL03,
            "validar":  lambda config: self._validar_venta("Venta EIAP", config),
            "nota":     lambda config: self._nota_venta(config, CUENTAS_EIAP, CLASES_EIAP),
        }
        # VBRL / FB03: ONE query, list read from a file another transaction downloaded in THIS run ("requiere").
        self._TX_VBRL = {
            "clave":    "vbrl",
            "nombre":   "VBRL (ZSE16)",
            "ejecutar": self._descargar_vbrl,
            "requiere": "monivoi",
            "leer_documentos": lambda ruta: leer_documentos_monivoi(ruta, "factura"),
            "unidad":   "factura(s)",
            "lista":    "las facturas del archivo de Monivoi de este periodo",
        }
        self._TX_FB03 = {
            "clave":    "fb03",
            "nombre":   "FB03 (documentos)",
            "ejecutar": self._descargar_fb03,
            "requiere": "gallo",
            "leer_documentos": leer_documentos_gallo_fb03,
            "unidad":   "documento(s)",
            "lista":    ("los documentos del archivo de Gallo de este periodo, sólo los de " + EtapaDocumentosFB03.DESCRIPCION_FILTRO),
            "sin_documentos_es_vacio": True,
            "filtro":   EtapaDocumentosFB03.DESCRIPCION_FILTRO,
        }
        # 2nd run: lists from the consolidation ("lista_de" → "documentos"); FBL5N / KNVP read VBRK's 'Solic.'.
        self._TX_VBRP = {
            "clave": "vbrp", "nombre": "VBRP (ZSE16)", "ejecutar": self._descargar_lista_2da,
            "descarga": VBRP_ZSE16, "lista_de": "referencias", "unidad": "referencia(s)",
        }
        self._TX_VBRK = {
            "clave": "vbrk", "nombre": "VBRK (ZSE16)", "ejecutar": self._descargar_lista_2da,
            "descarga": VBRK_ZSE16, "lista_de": "referencias", "unidad": "referencia(s)",
        }
        self._TX_VBRL2 = {
            "clave": "vbrl2", "nombre": "VBRL 2a corrida (ZSE16)", "ejecutar": self._descargar_lista_2da,
            "descarga": VBRL2_ZSE16, "lista_de": "referencias", "unidad": "referencia(s)",
        }
        self._TX_BKPF = {
            "clave": "bkpf", "nombre": "BKPF (ZSE16)", "ejecutar": self._descargar_lista_2da,
            "descarga": BKPF_ZSE16, "lista_de": "documentos", "unidad": "documento(s)",
            "con_periodo": True,
        }
        self._TX_FBL5N = {
            "clave": "fbl5n", "nombre": "FBL5N (deudores)", "ejecutar": self._descargar_solicitantes,
            "descarga": FBL5N_Deudores, "requiere": "vbrk", "leer_documentos": leer_solicitantes_vbrk,
            "unidad": "cliente(s)",
            "lista": "los solicitantes (columna 'Solic.') del VBRK de esta corrida",
        }
        self._TX_KNVP = {
            "clave": "knvp", "nombre": "KNVP (ZSE16)", "ejecutar": self._descargar_solicitantes,
            "descarga": KNVP_ZSE16, "requiere": "vbrk", "leer_documentos": leer_solicitantes_vbrk,
            "unidad": "cliente(s)",
            "lista": "los solicitantes (columna 'Solic.') del VBRK de esta corrida",
        }
        self._TX_POR_CLAVE = {
            tx["clave"]: tx
            for tx in (self._TX_GALLO, self._TX_MONIVOI, self._TX_DSD, self._TX_EIAP, self._TX_VBRL,
                       self._TX_FB03,
                       self._TX_VBRP, self._TX_VBRK, self._TX_VBRL2, self._TX_BKPF,
                       self._TX_FBL5N, self._TX_KNVP)
        }

        self._ocupado = False

        self.gui.on_download_gallo   = self.execute_download_gallo
        self.gui.on_download_monivoi = self.execute_download_monivoi
        self.gui.on_download_dsd     = self.execute_download_dsd
        self.gui.on_download_eiap    = self.execute_download_eiap
        self.gui.on_download_todos   = self.execute_download_todos
        self.gui.on_consolidar = self.execute_consolidacion

    def execute_download_gallo(self):
        self._execute_download([self._TX_GALLO], "Gallo")

    def execute_download_monivoi(self):
        self._execute_download([self._TX_MONIVOI], "Monivoi")

    def execute_download_dsd(self):
        self._execute_download([self._TX_DSD], "Venta DSD")

    def execute_download_eiap(self):
        self._execute_download([self._TX_EIAP], "Venta EIAP")

    def execute_download_todos(self):
        """Main download button: the run its GUI radio selects (1a = Gallo, Monivoi, VBRL, FB03; 2a = the
        six downloads of execute_download_segunda)."""
        if self._corrida_gui("corrida_descarga") == 2:
            self.execute_download_segunda()
            return
        # The GUI checkbox decides HOW the four run; single-transaction buttons never go parallel.
        usar_paralelo = getattr(self.gui, "usar_paralelo", None)
        corrida = CORRIDAS_DESCARGA[1]
        self._execute_download(
            [self._TX_POR_CLAVE[clave] for clave, _ in corrida],
            " + ".join(nombre for _, nombre in corrida),
            paralelo=bool(usar_paralelo()) if callable(usar_paralelo) else False,
        )

    def _corrida_gui(self, metodo):
        """1 or 2 from a GUI radio button reader (`metodo`); 1 when the GUI has none (older GUI)."""
        leer = getattr(self.gui, metodo, None)
        return 2 if callable(leer) and leer() == 2 else 1

    def execute_download_segunda(self):
        """'2a corrida' radio + download button: the six downloads, with the 1st consolidation's lists
        recomputed from the 1st run's files (calcular_listas_segunda_corrida)."""
        if self._ocupado:
            return
        if not self.gui.validate_dates():
            return
        cfg = {clave: self.gui.get_config(clave) for clave in ("gallo", "monivoi", "vbrl", "fb03")}
        rutas = {clave: os.path.join(c["input_path"], c["filename"]) for clave, c in cfg.items()}

        faltan = [cfg[clave]["filename"] for clave in ("gallo", "vbrl") if not os.path.exists(rutas[clave])]
        if faltan:
            messagebox.showerror(
                "2a corrida de descargas",
                "La 2a corrida pega lo que quedó en #N/A en la 1a consolidación, y eso sale de los\n"
                "archivos de la 1a corrida. No encontré:\n\n"
                + "".join(f"  • {nombre}\n" for nombre in faltan)
                + f"\nen la carpeta:\n  {cfg['gallo']['input_path']}\n\n"
                "Descarga primero la 1a corrida de este periodo.",
            )
            return
        ruta_fb03 = rutas["fb03"]
        if not os.path.exists(ruta_fb03):
            if not messagebox.askyesno(
                "Archivo de FB03 no encontrado",
                f"No encontré {cfg['fb03']['filename']} en la carpeta de descarga.\n\n"
                "Sin FB03 las listas salen SÓLO de la tabla Facturas (sin 'Notas de crédito'\n"
                "ni 'PEP Vouchers').\n\n¿Continuar así?",
            ):
                return
            ruta_fb03 = None

        self._ocupado = True
        listas = None
        try:
            self.gui.disable_buttons()
            listas = calcular_listas_segunda_corrida(
                rutas["gallo"], rutas["vbrl"], ruta_fb03, callback_status=self.gui.set_status)
        except (FileNotFoundError, ValueError) as e:
            self.gui.set_status("❌ Error al calcular las listas de la 2a corrida")
            messagebox.showerror("2a corrida de descargas", f"No pude calcular las listas:\n\n{e}")
        except Exception as e:
            self.gui.set_status("❌ Error al calcular las listas de la 2a corrida")
            messagebox.showerror("Error", f"Ocurrió un error inesperado al calcular las listas:\n\n{e}")
        finally:
            self._ocupado = False
            self.gui.enable_buttons()
        if listas is None:
            return
        self.gui.set_status("✓ Listo para comenzar")

        origen = [f"Calculadas de: {cfg['gallo']['filename']}, {cfg['vbrl']['filename']}"
                  + (f", {cfg['fb03']['filename']}" if ruta_fb03 else " (sin FB03)")]
        # Same staleness checks as the consolidation: an older child may not cover its parent.
        for hijo, padre, efecto in (("vbrl", "monivoi", "facturas"), ("fb03", "gallo", "documentos")):
            if hijo == "fb03" and ruta_fb03 is None:
                continue
            try:
                anterior = os.path.getmtime(rutas[hijo]) < os.path.getmtime(rutas[padre]) - 2
            except OSError:
                anterior = False
            if anterior:
                origen.append(
                    f"⚠️ {cfg[hijo]['filename']} es ANTERIOR a {cfg[padre]['filename']}: puede no "
                    f"cubrir sus {efecto} (#N/A de más)."
                )
        try:
            self._ofrecer_segunda_corrida(listas, cfg["gallo"], origen="\n".join(origen))
        except Exception as e:
            self.gui.set_status("❌ Error al preparar la 2a corrida")
            messagebox.showerror("2a corrida de descargas", f"No pude preparar la 2a corrida:\n\n{e}")

    def _execute_download(self, txs, etiqueta, paralelo=False):
        """Run `txs` (one after another, or in parallel by passes) and show ONE summary."""
        if not txs:
            raise ValueError("No se indicó ninguna transacción a descargar.")
        # A dependent transaction (VBRL needs THIS run's Monivoi) goes after its parent: no stale files.
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

        base = self.gui.get_config(txs[0]["clave"])
        mode = base["mode"]

        for tx in txs:
            validar = tx.get("validar")
            problema = validar(self.gui.get_config(tx["clave"])) if validar else None
            if problema:
                messagebox.showerror(f"Descarga de {tx['nombre']}", problema)
                return

        n_ventanas = self._ventanas_necesarias(txs) if paralelo else 1
        paralelo = n_ventanas > 1
        configs = None
        if paralelo:
            # Read everything workers need NOW: Tk may only be touched by its own thread.
            configs = {tx["clave"]: self.gui.get_config(tx["clave"]) for tx in txs}
            problema = self._validar_paralelo(txs, configs)
            if problema:
                messagebox.showerror(f"Descarga de {etiqueta}", problema)
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
            if paralelo:
                pregunta = pregunta.replace("¿Iniciar la descarga", "¿Iniciar EN PARALELO la descarga")
                detalle = (
                    "Se ejecutan por pasadas: las transacciones de una pasada corren a la vez, cada una en su propia ventana de SAP, y cada pasada empieza cuando termina la anterior:\n\n"
                    + "".join(
                        f"  Pasada {i}: {' ‖ '.join(tx['nombre'] for tx in pasada)}\n"
                        for i, pasada in enumerate(self._pasadas(txs, n_ventanas), 1)
                    )
                    + "\n"
                    f"Uso las ventanas de SAP que ya estén abiertas...\n"
                    "*IMPORTANTE:* No trabajar en esas ventanas mientras corre y en la medida de lo posible evitar hacer click en ellas.\n\n"
                )
            else:
                detalle = (
                    "Se ejecutan una tras otra:\n"
                    + "".join(f"  {i}) {tx['nombre']}\n" for i, tx in enumerate(txs, 1))
                    + "\n"
                )
            con_lista = [tx for tx in txs if tx.get("lista")]
            if con_lista:
                detalle += (
                    "".join(
                        f"{self._nombre_corto(tx)} consulta {tx['lista']}.\n"
                        for tx in con_lista
                    )
                    + "\nMientras corren NO copiar ni pegar nada!! Puesto que el portapapeles se usa para pasarle las listas a SAP...\n\n"
                )
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
        estado_final = None
        try:
            self.gui.disable_buttons()
            os.makedirs(base["input_path"], exist_ok=True)
            inicio_total = time.time()
            if paralelo:
                lineas, resultados, duraciones, notas = self._descargar_en_paralelo(
                    txs, configs, n_ventanas
                )
            else:
                lineas, resultados, duraciones = self._correr_secuencial(txs)
                notas = []
            total = time.time() - inicio_total

            resumen = [
                f"{lineas[tx['clave']]} ({_fmt_duracion(duraciones[tx['clave']])})" for tx in txs
            ]
            todo_ok = all(resultados.get(tx["clave"]) == "datos" for tx in txs)

            estado_final = (
                f"✅ ¡Descarga de {etiqueta} completada!" if todo_ok
                else f"⚠️ Descarga de {etiqueta} terminada con avisos"
            )
            self.gui.set_status(estado_final)
            mostrar = messagebox.showinfo if todo_ok else messagebox.showwarning
            mostrar(
                f"Resumen de descarga ({etiqueta})",
                "Proceso terminado:\n\n" + "\n".join(resumen)
                + "\n\n" + "".join(f"{n}\n" for n in notas)
                + f"⏱️ Tiempo total: {_fmt_duracion(total)}"
                + f"\nRuta: {base['input_path']}"
            )
        except Exception as e:
            estado_final = None
            self.gui.set_status("❌ Error en la descarga")
            messagebox.showerror("Error", f"Ocurrió un error inesperado:\n\n{e}")
        finally:
            self._ocupado = False
            self.gui.enable_buttons()
            if estado_final is None:
                self.gui.set_status("✓ Listo para comenzar")

    def _ejecutar_tx(self, tx, config, resultados, *, session=None, cfg_padre=None, estado_cb=None):
        """Run ONE transaction → (estado "datos" | "vacio" | "error", summary line). Never raises."""
        inicio_tx = time.time()
        try:
            if "ejecutar" in tx:
                return tx["ejecutar"](
                    tx, config, resultados,
                    session=session, cfg_padre=cfg_padre, estado_cb=estado_cb,
                )
            if self._run_download(tx, config, session=session):
                return "datos", f"✅ {tx['nombre']}: {config['filename']}"
            estado = self._clasificar_sin_datos(
                os.path.join(config["input_path"], config["filename"]), inicio_tx
            )
            return estado, f"⚠️ {tx['nombre']}: sin datos o error (ver consola)"
        except DescargaPorRevisarError as e:
            return "error", f"⚠️ {tx['nombre']}: {e}"
        except Exception as e:
            return "error", f"❌ {tx['nombre']}: error — {e}"

    def _correr_secuencial(self, txs, configs=None):
        """Run `txs` one after another → three dicts by "clave": lines, estados, seconds. `configs`
        ({clave: config}, the 2nd run) replaces reading the GUI, also for the parent's file.
        """
        lineas, resultados, duraciones = {}, {}, {}
        for i, tx in enumerate(txs, 1):
            clave = tx["clave"]
            prefijo = f"({i}/{len(txs)}) " if len(txs) > 1 else ""
            self.gui.set_status(
                f"⏳ {prefijo}Descargando {tx['nombre']}... (no cierre SAP)"
            )
            inicio = time.time()
            try:
                config = configs[clave] if configs is not None else self.gui.get_config(clave)
            except Exception as e:
                estado, linea = "error", f"❌ {tx['nombre']}: error — {e}"
            else:
                requiere = tx.get("requiere")
                cfg_padre = configs.get(requiere) if (configs is not None and requiere) else None
                estado, linea = self._ejecutar_tx(
                    tx, config, resultados, cfg_padre=cfg_padre, estado_cb=self.gui.set_status
                )
            resultados[clave] = estado
            lineas[clave] = linea
            duraciones[clave] = time.time() - inicio
        return lineas, resultados, duraciones

    @staticmethod
    def _niveles(txs):
        """{clave: level}: 0 if it needs no other transaction, else the level of the one it needs + 1,
        which must come EARLIER in `txs` (otherwise the dependent one could read a stale file).
        """
        nivel = {}
        for tx in txs:
            requiere = tx.get("requiere")
            if requiere and requiere not in nivel:
                raise ValueError(
                    f"{tx['nombre']} necesita que '{requiere}' se descargue antes, en la misma corrida."
                )
            nivel[tx["clave"]] = nivel[requiere] + 1 if requiere else 0
        return nivel

    @classmethod
    def _pasadas(cls, txs, n_ventanas):
        """Split `txs` into PASSES: a pass runs its transactions at once, one per window; it starts when the
        previous one is over, after the transactions it needs, with no more transactions than windows.
        E.g. Gallo, Monivoi, VBRL, FB03 on 2 windows → [[Gallo, Monivoi], [VBRL, FB03]]."""
        if int(n_ventanas) < 1:
            raise ValueError("Se necesita al menos una ventana de SAP.")
        nivel = cls._niveles(txs)
        grupos = {}
        for tx in txs:
            grupos.setdefault(nivel[tx["clave"]], []).append(tx)
        pasadas = []
        for n in sorted(grupos):
            grupo = grupos[n]
            pasadas.extend(grupo[i:i + n_ventanas] for i in range(0, len(grupo), n_ventanas))
        return pasadas

    @classmethod
    def _ventanas_necesarias(cls, txs):
        """Windows worth having: the biggest pass, capped at MAX_VENTANAS_SAP."""
        por_nivel = {}
        for n in cls._niveles(txs).values():
            por_nivel[n] = por_nivel.get(n, 0) + 1
        return min(max(por_nivel.values(), default=1), MAX_VENTANAS_SAP)

    @staticmethod
    def _nombre_corto(tx):
        """'Gallo (FAGLL03)' → 'Gallo' (for the status bar and the summary)."""
        return tx["nombre"].split(" (")[0]

    @classmethod
    def _etiqueta_ventana(cls, indice, tx):
        """Label of the console timeline: (1, Monivoi) → 'ventana 2 · Monivoi'."""
        return f"ventana {int(indice) + 1} · {cls._nombre_corto(tx)}"

    @staticmethod
    def _validar_paralelo(txs, configs):
        """Two transactions that would write the SAME file cannot run together. Error text or None."""
        vistos = {}
        for tx in txs:
            cfg = configs[tx["clave"]]
            destino = os.path.normcase(
                os.path.abspath(os.path.join(cfg["input_path"], cfg["filename"]))
            )
            if destino in vistos:
                return (
                    f"{vistos[destino]} y {tx['nombre']} guardarían el MISMO archivo "
                    f"({cfg['filename']}): no pueden correr a la vez."
                )
            vistos[destino] = tx["nombre"]
        return None

    @staticmethod
    def _validar_pasadas(pasadas):
        """Fixed passes (the 2nd run's): no transaction twice, and each one's parent ("requiere")
        in an EARLIER pass — never in its own, where it would not have ended yet.
        """
        vistas = set()
        for pasada in pasadas:
            claves = [tx["clave"] for tx in pasada]
            if len(set(claves)) != len(claves) or vistas.intersection(claves):
                raise ValueError("Una transacción aparece más de una vez en las pasadas.")
            for tx in pasada:
                requiere = tx.get("requiere")
                if requiere and requiere not in vistas:
                    raise ValueError(
                        f"{tx['nombre']} necesita que '{requiere}' se descargue en una pasada anterior."
                    )
            vistas.update(claves)

    @staticmethod
    def _partir_pasadas(pasadas, n_ventanas):
        """Fixed passes cut to at most `n_ventanas` transactions each, keeping the order: with 2 windows
        [[VBRK, VBRP, VBRL2], [FBL5N, KNVP, BKPF]] → [[VBRK, VBRP], [VBRL2], [FBL5N, KNVP], [BKPF]].
        """
        if int(n_ventanas) < 1:
            raise ValueError("Se necesita al menos una ventana de SAP.")
        return [pasada[i:i + n_ventanas] for pasada in pasadas for i in range(0, len(pasada), n_ventanas)]

    @staticmethod
    def _ventanas_de_pasadas(pasadas):
        """Windows worth having for fixed passes: the biggest pass, capped at MAX_VENTANAS_SAP."""
        return min(max((len(pasada) for pasada in pasadas), default=1), MAX_VENTANAS_SAP)

    def _descargar_en_paralelo(self, txs, configs, n_ventanas, pasadas=None):
        """The P&L way: SAP windows 1..n, the ones the USER already has open (this code never opens one),
        then the passes (`pasadas`: fixed ones, else split by dependency). With a single window it falls
        back to the sequential engine."""
        self.gui.set_status("⏳ Preparando las ventanas de SAP... (no cierre SAP)")
        disponibles = Paralelo.ventanas_disponibles(n_ventanas)
        if disponibles < 2:
            lineas, resultados, duraciones = self._correr_secuencial(txs, configs)
            return lineas, resultados, duraciones, [
                "ℹ️ Sólo hubo 1 ventana de SAP disponible: las transacciones corrieron una tras otra."
            ]
        lineas, resultados, duraciones, info = self._correr_paralelo(
            txs, configs, disponibles, pasadas=pasadas)
        notas = [f"⚡ Modo paralelo: {disponibles} ventana(s) de SAP a la vez."]
        if disponibles < n_ventanas:
            notas.append(
                f"ℹ️ Sólo hubo {disponibles} de las {n_ventanas} ventanas de SAP que pide este "
                f"proceso: cada pasada se partió en grupos de {disponibles}."
            )
        notas += self._notas_de_pasadas(txs, info)
        return lineas, resultados, duraciones, notas

    def _notas_de_pasadas(self, txs, pasadas):
        """'⚡ Pasada 1 · Gallo ‖ Monivoi: 3 min 06 s' per pass: its wall-clock time (not the sum of
        its transactions, which would hide a window that waited for another).
        """
        por_clave = {tx["clave"]: tx for tx in txs}
        return [
            f"⚡ Pasada {pasada['numero']} · "
            f"{' ‖ '.join(self._nombre_corto(por_clave[c]) for c in pasada['claves'])}: "
            f"{_fmt_duracion(pasada['segundos'])}"
            for pasada in pasadas
        ]

    def _correr_paralelo(self, txs, configs, n_ventanas, pasadas=None):
        """Run `txs` pass by pass on SAP windows 1..n, each pass inside Excel.modo_paralelo(), the next one
        starting when the previous is over, whatever its result. `pasadas`: fixed, else `_pasadas`. MAIN
        thread (Tk). Returns _correr_secuencial's dicts + per pass {numero, claves, segundos}."""
        lineas, resultados, duraciones = {}, {}, {}
        if pasadas is None:
            pasadas = self._pasadas(txs, n_ventanas)
        else:
            self._validar_pasadas(pasadas)
            pasadas = self._partir_pasadas(pasadas, n_ventanas)
        info = []
        inicio_global = time.time()

        for numero, pasada in enumerate(pasadas, 1):
            inicio_pasada = time.time()
            trabajos = []
            for indice, tx in enumerate(pasada):
                requiere = tx.get("requiere")
                trabajos.append((indice, self._trabajo_en_ventana, (
                    tx, configs[tx["clave"]], dict(resultados),
                    configs[requiere] if requiere else None,
                    self._etiqueta_ventana(indice, tx),
                )))

            def estado(vivos, pasada=pasada, numero=numero):
                self.gui.set_status(self._texto_estado_paralelo(
                    txs, resultados, {pasada[p]["clave"] for p in vivos},
                    time.time() - inicio_global, numero, len(pasadas),
                ))

            with Excel.modo_paralelo():
                salidas = Paralelo.correr(trabajos, al_esperar=estado)
            for tx, (ok, valor) in zip(pasada, salidas):
                if ok:
                    estado_tx, linea, segundos = valor
                else:
                    estado_tx, linea, segundos = (
                        "error", f"❌ {tx['nombre']}: error — {valor}", time.time() - inicio_pasada
                    )
                resultados[tx["clave"]] = estado_tx
                lineas[tx["clave"]] = linea
                duraciones[tx["clave"]] = segundos
            info.append({
                "numero": numero,
                "claves": [tx["clave"] for tx in pasada],
                "segundos": time.time() - inicio_pasada,
            })
        return lineas, resultados, duraciones, info

    def _trabajo_en_ventana(self, tx, config, resultados, cfg_padre, etiqueta, *, session):
        """Worker thread: run `tx` on its SAP window `session` → (estado, line, seconds); no Tk."""
        with Paralelo.etiqueta(etiqueta):
            inicio = time.time()
            Paralelo.marca("inicio")
            estado, linea = self._ejecutar_tx(
                tx, config, resultados, session=session, cfg_padre=cfg_padre
            )
            segundos = time.time() - inicio
            Paralelo.marca(f"fin: {estado} ({_fmt_duracion(segundos)})")
        return estado, linea, segundos

    @classmethod
    def _texto_estado_paralelo(cls, txs, resultados, en_ejecucion, segundos,
                               pasada=None, total_pasadas=None):
        """Status bar: '⚡ Pasada 1/2 · Gallo ⏳  Monivoi ✅  VBRL 💤  FB03 💤  (1:05)'."""
        partes = []
        for tx in txs:
            clave = tx["clave"]
            if clave in resultados:
                marca = "✅" if resultados[clave] == "datos" else "⚠️"
            elif clave in en_ejecucion:
                marca = "⏳"
            else:
                marca = "💤"
            partes.append(f"{cls._nombre_corto(tx)} {marca}")
        minutos, seg = divmod(int(segundos), 60)
        prefijo = f"⚡ Pasada {pasada}/{total_pasadas} · " if pasada else "⚡ "
        return prefijo + "  ".join(partes) + f"  ({minutos}:{seg:02d})"

    def execute_consolidacion(self):
        """Consolidate the Gallo and Monivoi CSVs of the GUI dates (VBRL and FB03 optional), offering to
        pick by hand a file that is not where it is expected."""
        if self._ocupado:
            return
        if self._corrida_gui("corrida_consolidacion") == 2:
            self.execute_consolidacion_segunda()
            return
        if not self.gui.validate_dates():
            return

        cfg_gallo   = self.gui.get_config("gallo")
        cfg_monivoi = self.gui.get_config("monivoi")
        cfg_vbrl    = self.gui.get_config("vbrl")
        input_path  = cfg_gallo["input_path"]

        ventas = self._ventas_solas(cfg_gallo, cfg_monivoi)
        if ventas:
            self._consolidar_solo_ventas(ventas)
            return

        ruta_gallo = self._localizar_archivo(input_path, cfg_gallo["filename"], "Gallo")
        if not ruta_gallo:
            return
        ruta_monivoi = self._localizar_archivo(input_path, cfg_monivoi["filename"], "Monivoi")
        if not ruta_monivoi:
            return
        ruta_vbrl, cancelar = self._resolver_vbrl(input_path, cfg_vbrl["filename"], ruta_monivoi)
        if cancelar:
            return
        # FB03 only matters with a VBRL (the Notas / PEP tables read both): without one it is not asked for.
        ruta_fb03 = None
        if ruta_vbrl:
            cfg_fb03 = self.gui.get_config("fb03")
            ruta_fb03, cancelar = self._resolver_fb03(input_path, cfg_fb03["filename"], ruta_gallo)
            if cancelar:
                return

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
        if ruta_fb03:
            linea_fb03 = f"  📑 FB03    : {os.path.basename(ruta_fb03)}\n"
        elif ruta_vbrl:
            linea_fb03 = ("  📑 FB03    : (sin FB03: no se generan 'FB', 'Notas de crédito' "
                          "ni 'PEP Vouchers')\n")
        else:
            linea_fb03 = ""
        confirm = messagebox.askyesno(
            "Confirmar consolidación",
            f"¿Consolidar estos archivos en un solo Excel?\n\n"
            f"  🐓 Gallo   : {os.path.basename(ruta_gallo)}\n"
            f"  📊 Monivoi : {os.path.basename(ruta_monivoi)}\n"
            f"{linea_vbrl}{linea_fb03}\n"
            f"Se generará:\n  {nombre_salida}\n\n"
            f"En la carpeta:\n  {ruta_output}"
        )
        if not confirm:
            return

        self._ocupado = True
        listas_2da = None
        try:
            self.gui.disable_buttons()
            resumen = consolidar_gallo_monivoi(
                ruta_gallo, ruta_monivoi, ruta_output,
                nombre_salida=nombre_salida,
                callback_status=self.gui.set_status,
                ruta_vbrl=ruta_vbrl,
                ruta_fb03=ruta_fb03,
            )

            nota_monivoi = (
                "\n⚠️ El archivo de Monivoi venía sin datos; su pestaña quedó vacía."
                if resumen["monivoi_sin_datos"] else ""
            )
            avisos = self._avisos_consolidacion(resumen)
            nota_ventas = ("\n" + "\n".join(f"⚠️ {a}" for a in avisos)) if avisos else ""

            lineas_vbrl = ""
            if resumen.get("vbrl_incluido"):
                lineas_vbrl = (
                    f"\n  • Filas VBRL               : {resumen['filas_vbrl']:,}"
                    f"\n  • Referencias (facturas)   : {resumen['filas_td_referencias']:,}"
                    f"\n      …en #N/A               : {resumen['td_referencias_na']:,}"
                )
                if resumen["vbrl_sin_datos"]:
                    lineas_vbrl += "\n⚠️ El archivo de VBRL venía sin datos; su pestaña quedó vacía."
                if resumen.get("fb03_incluido"):
                    lineas_vbrl += (
                        f"\n  • Filas FB03               : {resumen['filas_fb']:,}"
                        f"\n  • Notas de crédito         : {resumen['filas_notas_ref']:,}"
                        f"  (Σ {resumen['total_notas_ref']:,.2f})"
                        f"\n      …en #N/A               : {resumen['notas_ref_na']:,}"
                        f"\n  • PEP Vouchers             : {resumen['filas_pep_ref']:,}"
                        f"  (Σ {resumen['total_pep_ref']:,.2f})"
                        f"\n      …en #N/A               : {resumen['pep_ref_na']:,}\n\n"
                    )
                listas = resumen.get("listas_2da")

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
            listas_2da = resumen.get("listas_2da")
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
        if listas_2da is not None:
            try:
                self._ofrecer_segunda_corrida(listas_2da, cfg_gallo)
            except Exception as e:
                self.gui.set_status("❌ Error al preparar la 2a corrida")
                messagebox.showerror(
                    "2a corrida de descargas",
                    f"El Excel consolidado sí se generó, pero no pude preparar la 2a corrida:\n\n{e}",
                )

    @staticmethod
    def _avisos_consolidacion(resumen):
        """Every warning of a consolidation summary, in the order the stakeholder should read them
        (the 2nd consolidation's own come last)."""
        return (
            (resumen.get("td_avisos") or [])
            + (resumen.get("mtd_avisos") or [])
            + (resumen.get("cruce_avisos") or [])
            + (resumen.get("vbrl_avisos") or [])
            + (resumen.get("td_referencias_avisos") or [])
            + (resumen.get("notas_pep_avisos") or [])
            + ((resumen.get("listas_2da") or {}).get("avisos") or [])
            + (resumen.get("segunda_avisos") or [])
        )

    _ETIQUETAS_SEGUNDA = (
        ("vbrk", "VBRK"), ("fbl5n", "FBL5N"), ("vbrp", "VBRP"),
        ("knvp", "KNVP"), ("vbrl2", "VBRL (2a corrida)"), ("bkpf", "BKPF"),
    )
    # A 2nd-run file older than this many seconds was asked for an older 1st run (lookups may miss).
    _TOLERANCIA_ANTIGUEDAD = 2

    def execute_consolidacion_segunda(self):
        """'2a corrida' radio + consolidation button: the 1st consolidation plus the six sheets of the 2nd
        run and the 'Diferencias' tab (same file name, same code path)."""
        if self._ocupado:
            return
        if not self.gui.validate_dates():
            return

        cfg_gallo   = self.gui.get_config("gallo")
        cfg_monivoi = self.gui.get_config("monivoi")
        input_path  = cfg_gallo["input_path"]

        ruta_gallo = self._localizar_archivo(input_path, cfg_gallo["filename"], "Gallo")
        if not ruta_gallo:
            return
        ruta_monivoi = self._localizar_archivo(input_path, cfg_monivoi["filename"], "Monivoi")
        if not ruta_monivoi:
            return
        # The VBRL is NOT optional here: 'Diferencias' lists the #N/A rows of the tables it feeds.
        ruta_vbrl = self._localizar_archivo(
            input_path, self.gui.get_config("vbrl")["filename"], "VBRL")
        if not ruta_vbrl:
            return
        ruta_fb03, cancelar = self._resolver_fb03(
            input_path, self.gui.get_config("fb03")["filename"], ruta_gallo)
        if cancelar:
            return

        rutas_segunda = {}
        for clave, etiqueta in self._ETIQUETAS_SEGUNDA:
            nombre = self.gui.get_config(clave)["filename"]
            if clave == "fbl5n":
                ruta = self._localizar_archivo(
                    input_path, nombre, etiqueta,
                    tipos=[("Listas de SAP", "*.XLS *.xls *.txt"), ("Todos los archivos", "*.*")])
            else:
                ruta = self._localizar_archivo(input_path, nombre, etiqueta)
            if not ruta:
                return
            rutas_segunda[clave] = ruta

        padres = [ruta for ruta in (ruta_gallo, ruta_monivoi, ruta_vbrl, ruta_fb03) if ruta]
        try:
            ultimo_padre = max(os.path.getmtime(ruta) for ruta in padres)
            anteriores = [os.path.basename(ruta) for ruta in rutas_segunda.values()
                          if os.path.getmtime(ruta) < ultimo_padre - self._TOLERANCIA_ANTIGUEDAD]
        except OSError:
            anteriores = []
        if anteriores:
            if not messagebox.askyesno(
                "Descargas de la 2a corrida anteriores a la 1a",
                "Estos archivos de la 2a corrida se descargaron ANTES que alguno de la 1a corrida:\n\n"
                + "".join(f"  • {nombre}\n" for nombre in anteriores)
                + "\nProbablemente no cubren los documentos actuales y 'Diferencias' mostraría #N/A "
                "de más.\n\n¿Consolidar de todos modos?",
            ):
                return

        try:
            ruta_output = self.gui.get_consolidacion_path()
        except OSError as e:
            messagebox.showerror("Error", f"No se pudo usar la carpeta de consolidación:\n{e}")
            return
        nombre_salida = f"Consolidado_{os.path.splitext(cfg_gallo['filename'])[0]}.xlsx"
        ruta_final = os.path.join(ruta_output, nombre_salida)

        linea_fb03 = (f"  📑 FB03    : {os.path.basename(ruta_fb03)}\n" if ruta_fb03 else
                      "  📑 FB03    : (sin FB03: 'Diferencias' sólo con las Facturas)\n")
        lineas_segunda = "".join(
            f"  • {etiqueta:<17}: {os.path.basename(rutas_segunda[clave])}\n"
            for clave, etiqueta in self._ETIQUETAS_SEGUNDA)
        reemplaza = ("\n⚠️ Ya existe y se REEMPLAZA (el mismo Excel, completado con las pestañas nuevas)."
                     if os.path.exists(ruta_final) else "")
        if not messagebox.askyesno(
            "Confirmar 2a consolidación",
            "¿Completar la consolidación con la 2a corrida?\n\n"
            f"  🐓 Gallo   : {os.path.basename(ruta_gallo)}\n"
            f"  📊 Monivoi : {os.path.basename(ruta_monivoi)}\n"
            f"  🧾 VBRL    : {os.path.basename(ruta_vbrl)}\n"
            f"{linea_fb03}\n"
            f"2a corrida:\n{lineas_segunda}\n"
            f"Se generará:\n  {nombre_salida}{reemplaza}\n\n"
            f"En la carpeta:\n  {ruta_output}"
        ):
            return

        self._ocupado = True
        try:
            self.gui.disable_buttons()
            resumen = consolidar_gallo_monivoi(
                ruta_gallo, ruta_monivoi, ruta_output,
                nombre_salida=nombre_salida,
                callback_status=self.gui.set_status,
                ruta_vbrl=ruta_vbrl,
                ruta_fb03=ruta_fb03,
                rutas_segunda=rutas_segunda,
            )
            avisos = self._avisos_consolidacion(resumen)
            nota_avisos = ("\n" + "\n".join(f"⚠️ {a}" for a in avisos)) if avisos else ""
            d = resumen["diferencias"]
            por_tipo = d["por_tipo"]
            self.gui.set_status("✅ ¡2a consolidación completada!")
            messagebox.showinfo(
                "2a consolidación completada",
                f"Archivo generado:\n{resumen['ruta']}\n\n"
                f"  • Renglones en 'Diferencias': {d['filas']:,}"
                f"  (Facturas {por_tipo.get('Factura', 0):,} · NC {por_tipo.get('NC', 0):,}"
                f" · PV {por_tipo.get('PV', 0):,})\n"
                f"  • Σ Monto                    : {d['importe']:,.2f}\n"
                f"  • Sin No cliente (VBRK)      : {d['sin_vbrk']:,}\n"
                f"  • Sin Lista de factura       : {d['sin_lista']:,}\n"
                f"  • Documentos anulados        : {d['anuladas']:,}\n"
                f"  • Cliente Global / Vendedor / Individual: "
                f"{d['cliente_global']:,} / {d['vendedor']:,} / {d['cliente_individual']:,}"
                f"{nota_avisos}"
            )
        except (FileNotFoundError, ValueError) as e:
            self.gui.set_status("❌ Error en la 2a consolidación")
            messagebox.showerror("Error en la 2a consolidación", str(e))
        except Exception as e:
            self.gui.set_status("❌ Error en la 2a consolidación")
            messagebox.showerror("Error", f"Ocurrió un error inesperado:\n\n{e}")
        finally:
            self._ocupado = False
            self.gui.enable_buttons()
            if "completada" not in self.gui.status_var.get():
                self.gui.set_status("✓ Listo para comenzar")

    _VENTAS = (
        ("dsd",  "Venta DSD",  consolidar_venta_dsd),
        ("eiap", "Venta EIAP", consolidar_venta_eiap),
    )

    def _ventas_solas(self, cfg_gallo, cfg_monivoi):
        """The sales downloaded for the GUI range — [(name, config, consolidation function)] — when the
        folder has NEITHER the Gallo nor the Monivoi file (only the sales buttons were used); []
        otherwise."""
        def existe(cfg):
            return os.path.exists(os.path.join(cfg["input_path"], cfg["filename"]))
        if existe(cfg_gallo) or existe(cfg_monivoi):
            return []
        ventas = []
        for clave, nombre, consolidar in self._VENTAS:
            cfg = self.gui.get_config(clave)
            if cfg["mode"] == "range" and existe(cfg):
                ventas.append((nombre, cfg, consolidar))
        return ventas

    def _consolidar_solo_ventas(self, ventas):
        """Consolidate ONLY the sales downloads found: one workbook per sale, each with its single tab."""
        for nombre, cfg, _ in ventas:
            error = self._validar_venta(nombre, cfg)
            if error:
                messagebox.showerror(nombre, error)
                return
        try:
            ruta_output = self.gui.get_consolidacion_path()
        except OSError as e:
            messagebox.showerror(
                "Error", f"No se pudo usar la carpeta de consolidación:\n{e}"
            )
            return

        nombres = [nombre for nombre, _, _ in ventas]
        trabajos = []         # (name, CSV path, workbook name, days SAP was queried for, function)
        lineas = []
        for nombre, cfg, consolidar in ventas:
            hasta = fecha_efectiva_eiap(cfg["date_to"]).strftime(FORMATO_FECHA)
            # Every day SAP was queried for (Hasta - 1), so they all show even without amounts.
            dias = self._iterar_dias(cfg["date_from"], hasta)
            nombre_salida = f"Consolidado_{os.path.splitext(cfg['filename'])[0]}.xlsx"
            trabajos.append((nombre, os.path.join(cfg["input_path"], cfg["filename"]),
                             nombre_salida, dias, consolidar))
            lineas.append((cfg["filename"], f"{cfg['date_from']} — {hasta}", nombre_salida))

        if len(ventas) == 1:
            archivo, periodo, nombre_salida = lineas[0]
            detalle = (
                f"Sólo se encontró la descarga de {nombres[0]} (no hay Gallo ni Monivoi).\n"
                f"Se generará un Excel con las pestañas 'Gallo' (la base descargada) y '{nombres[0]}'.\n\n"
                f"  🛒 {nombres[0]}: {archivo}\n"
                f"  Periodo       : {periodo}\n\n"
                f"Se generará:\n  {nombre_salida}\n\n"
            )
        else:
            detalle = (
                f"Sólo se encontraron las descargas de {' y '.join(nombres)} (no hay Gallo ni Monivoi).\n"
                "Se generará un Excel por cada una, con las pestañas 'Gallo' (la base descargada) y la de su venta.\n\n"
                + "".join(
                    f"  🛒 {nombre}: {archivo}\n"
                    f"  Periodo       : {periodo}\n"
                    f"  Se generará   : {nombre_salida}\n\n"
                    for nombre, (archivo, periodo, nombre_salida) in zip(nombres, lineas)
                )
            )
        confirm = messagebox.askyesno(
            "Confirmar consolidación",
            f"{detalle}"
            f"En la carpeta:\n  {ruta_output}"
        )
        if not confirm:
            return

        self._ocupado = True
        try:
            self.gui.disable_buttons()
            resumenes = []
            for nombre, ruta_csv, nombre_salida, dias, consolidar in trabajos:
                resumenes.append((nombre, consolidar(
                    ruta_csv, ruta_output,
                    nombre_salida=nombre_salida,
                    callback_status=self.gui.set_status,
                    fechas_seleccionadas=dias,
                )))
            partes = []
            for nombre, resumen in resumenes:
                avisos = resumen.get("venta_avisos") or []
                nota = ("\n" + "\n".join(f"⚠️ {a}" for a in avisos)) if avisos else ""
                partes.append(
                    f"Archivo generado:\n{resumen['ruta']}\n\n"
                    f"  • Filas del archivo        : {resumen['filas_archivo']:,}\n"
                    f"  • Σ {nombre} (moneda local): {resumen['total_venta']:,.2f}"
                    f"{nota}"
                )
            self.gui.set_status("✅ ¡Consolidación completada!")
            messagebox.showinfo("Consolidación completada", "\n\n".join(partes))
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
        """The VBRL CSV (OPTIONAL) → (ruta or None = without VBRL, cancelar = do not consolidate)."""
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

        # A VBRL OLDER than its Monivoi may not cover its invoices (extra #N/A).
        try:
            vbrl_anterior = os.path.getmtime(ruta) < os.path.getmtime(ruta_monivoi) - 2
        except OSError:
            vbrl_anterior = False
        if vbrl_anterior:
            decision = messagebox.askyesnocancel(
                "VBRL anterior al archivo de Monivoi",
                f"El archivo de VBRL:\n  {os.path.basename(ruta)}\n"
                f"se descargó ANTES que el de Monivoi:\n  {os.path.basename(ruta_monivoi)}\n\n"
                "Probablemente no cubre las facturas actuales y saldrían #N/A de más.\n\n"
                "Sí = usarlo de todos modos\n"
                "No = consolidar SIN VBRL\n"
                "Cancelar = no consolidar",
            )
            if decision is None:
                return None, True
            if not decision:
                return None, False
        return ruta, False

    def _resolver_fb03(self, input_path, filename, ruta_gallo):
        """The FB03 CSV (OPTIONAL, only asked for with a VBRL) → (ruta or None = without FB03, cancelar = do
        not consolidate)."""
        ruta = os.path.join(input_path, filename)
        if not os.path.exists(ruta):
            decision = messagebox.askyesnocancel(
                "Archivo de FB03 no encontrado",
                f"No encontré el archivo esperado:\n\n  {filename}\n\n"
                f"en la carpeta:\n  {input_path}\n\n"
                "Sin FB03 el Excel sale sin las tablas 'Notas de crédito' y 'PEP Vouchers' "
                "(ni la pestaña 'FB').\n\n"
                "Sí = consolidar SIN FB03\n"
                "No = seleccionar el archivo manualmente\n"
                "Cancelar = no consolidar",
            )
            if decision is None:
                return None, True
            if decision:
                return None, False
            ruta = filedialog.askopenfilename(
                title="Selecciona el CSV de FB03",
                initialdir=input_path,
                filetypes=[("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")],
            )
            if not ruta:
                return None, True

        # The FB03 list comes from the Gallo: one OLDER than it may not cover its documents (#N/A).
        try:
            fb03_anterior = os.path.getmtime(ruta) < os.path.getmtime(ruta_gallo) - 2
        except OSError:
            fb03_anterior = False
        if fb03_anterior:
            decision = messagebox.askyesnocancel(
                "FB03 anterior al archivo de Gallo",
                f"El archivo de FB03:\n  {os.path.basename(ruta)}\n"
                f"se descargó ANTES que el de Gallo:\n  {os.path.basename(ruta_gallo)}\n\n"
                "Probablemente no cubre los documentos actuales y saldrían #N/A de más.\n\n"
                "Sí = usarlo de todos modos\n"
                "No = consolidar SIN FB03\n"
                "Cancelar = no consolidar",
            )
            if decision is None:
                return None, True
            if not decision:
                return None, False
        return ruta, False

    def _localizar_archivo(self, input_path, filename, etiqueta, tipos=None):
        """The file with the expected name, or one picked by hand if it is not there; None = cancel.
        `tipos`: the dialog's file types (CSV by default; FBL5N is SAP's text list)."""
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
            title=f"Selecciona el archivo de {etiqueta}" if tipos else f"Selecciona el CSV de {etiqueta}",
            initialdir=input_path,
            filetypes=tipos or [("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")],
        )
        return ruta_manual or None

    def _txs_segunda_corrida(self):
        """PASADAS_2DA as transaction profiles: one list per pass (its transactions share a pass,
        one SAP window each)."""
        return [[self._TX_POR_CLAVE[clave] for clave in pasada] for pasada in self.PASADAS_2DA]

    @staticmethod
    def _configs_segunda_corrida(base, listas, txs):
        """One config per transaction from `base`, the config the consolidation read (folder, company code,
        period) — never the GUI again: it may have been edited meanwhile. List transactions get their
        list (a key of `listas`) as "documentos"."""
        configs = {}
        for tx in txs:
            cfg = {campo: base[campo] for campo in
                   ("mode", "sociedad", "input_path", "fecha", "date_from", "date_to") if campo in base}
            if base["mode"] == "single":
                cfg["filename"] = nombre_archivo_dia_2da(tx["clave"], base["fecha"])
            else:
                cfg["filename"] = nombre_archivo_rango_2da(tx["clave"], base["date_from"], base["date_to"])
            if tx.get("lista_de"):
                cfg["documentos"] = list(listas[tx["lista_de"]])
            configs[tx["clave"]] = cfg
        return configs

    def _ofrecer_segunda_corrida(self, listas, base, origen=None):
        """After a consolidation (or the '2a corrida' button): offer the 2nd run with its lists, or say
        there is nothing to ask. `origen`: extra lines for the confirmation (where the lists came from).
        """
        if not listas["referencias"] and not listas["documentos"]:
            messagebox.showinfo(
                "2a corrida de descargas",
                "En 'TD.Diferencias_Referencias' ningún renglón quedó en #N/A (o sólo los que\n"
                "tienen una Referencia de pago que empieza con 'FP'): no hay nada que consultar\n"
                "en VBRP, VBRK, VBRL, BKPF, FBL5N ni KNVP.",
            )
            return
        self._execute_segunda_corrida(listas, base, origen=origen)

    def _execute_segunda_corrida(self, listas, base, origen=None):
        """Confirm and run the 2nd run (two passes of three at once with the GUI checkbox, else one after
        another) and show ONE summary."""
        if self._ocupado:
            return
        pasadas = self._txs_segunda_corrida()
        self._validar_pasadas(pasadas)        # VBRK in an earlier pass than FBL5N / KNVP, in both modes
        txs = [tx for pasada in pasadas for tx in pasada]
        n_ventanas = self._ventanas_de_pasadas(pasadas)
        configs = self._configs_segunda_corrida(base, listas, txs)
        problema = self._validar_paralelo(txs, configs)
        if problema:
            messagebox.showerror("2a corrida de descargas", problema)
            return
        usar_paralelo = getattr(self.gui, "usar_paralelo", None)
        paralelo = bool(usar_paralelo()) if callable(usar_paralelo) else False

        periodo = (
            f"Fecha    : {base['fecha']}" if base["mode"] == "single"
            else f"Periodo  : {base['date_from']} — {base['date_to']}"
        )
        excluidas = listas["excluidas_fp"]
        nota_fp = ""
        if excluidas["renglones"]:
            nota_fp = (f"(fuera {excluidas['renglones']:,} renglón(es) cuya Referencia de pago empieza "
                       "con 'FP':\n ni su referencia ni su Nº documento se consultan)\n")
        nota_origen = f"{origen}\n" if origen else ""
        nota_avisos = "".join(f"⚠️ {aviso}\n" for aviso in listas.get("avisos") or [])
        if paralelo:
            orden = (
                "En paralelo, por pasadas: las transacciones de una pasada corren a la vez, cada una en su propia ventana de SAP, y la siguiente empieza cuando termina:\n\n"
                + "".join(
                    f"  Pasada {i}: {' ‖ '.join(self._nombre_corto(tx) for tx in pasada)}\n"
                    for i, pasada in enumerate(pasadas, 1)
                )
                + "\nFBL5N y KNVP usan la columna 'Solic.' de VBRK (por eso VBRK va en la pasada 1).\n\n"
                f"Uso las ventanas 1 a {n_ventanas} de SAP que ya estén abiertas... No trabajar en ellas mientras corre.\n"
            )
        else:
            orden = ("Se ejecutan una tras otra:\n"
                     + "".join(f"  {i}) {tx['nombre']}\n" for i, tx in enumerate(txs, 1)) + "\n")
        confirm = messagebox.askyesno(
            "2a corrida de descargas",
            "¿Descargar ahora las 6 transacciones de la 2a corrida?\n\n"
            f"Sociedad : {base['sociedad']}\n{periodo}\n\n"
            f"{nota_fp}{nota_origen}{nota_avisos}\n"
            f"{orden}"
            "Mientras corren NO copiar ni pegar nada!! El portapapeles se usa para pasarle las listas a SAP.\n\n"
            f"Se guardará en:\n{base['input_path']}"
        )
        if not confirm:
            return

        self._ocupado = True
        estado_final = None
        try:
            self.gui.disable_buttons()
            os.makedirs(base["input_path"], exist_ok=True)
            inicio_total = time.time()
            if paralelo:
                lineas, resultados, duraciones, notas = self._descargar_en_paralelo(
                    txs, configs, n_ventanas, pasadas=pasadas
                )
            else:
                lineas, resultados, duraciones = self._correr_secuencial(txs, configs)
                notas = []
            total = time.time() - inicio_total
            todo_ok = all(resultados.get(tx["clave"]) == "datos" for tx in txs)
            estado_final = ("✅ ¡2a corrida completada!" if todo_ok
                            else "⚠️ 2a corrida terminada con avisos")
            self.gui.set_status(estado_final)
            mostrar = messagebox.showinfo if todo_ok else messagebox.showwarning
            mostrar(
                "Resumen de la 2a corrida",
                "Proceso terminado:\n\n"
                + "\n".join(f"{lineas[tx['clave']]} ({_fmt_duracion(duraciones[tx['clave']])})"
                            for tx in txs)
                + "\n\n" + "".join(f"{n}\n" for n in notas)
                + f"⏱️ Tiempo total: {_fmt_duracion(total)}"
                + f"\nRuta: {base['input_path']}"
            )
        except Exception as e:
            estado_final = None
            self.gui.set_status("❌ Error en la 2a corrida")
            messagebox.showerror("Error", f"Ocurrió un error inesperado en la 2a corrida:\n\n{e}")
        finally:
            self._ocupado = False
            self.gui.enable_buttons()
            if estado_final is None:
                self.gui.set_status("✓ Listo para comenzar")

    @staticmethod
    def _texto_info(info):
        """What a 2nd-run download chose / checked, for the summary line: ' · layout 'X' · ...'."""
        partes = []
        if info.get("layout"):
            partes.append(f"layout {info['layout']!r}")
        if info.get("campo"):
            partes.append(f"campo {info['campo']!r}")
        if info.get("presentes"):
            encontrados, pedidos = info["presentes"]
            partes.append(f"{encontrados:,} de {pedidos:,} cliente(s) aparecen en el archivo")
        return "".join(f" · {parte}" for parte in partes)

    def _descargar_lista_2da(self, tx, config, resultados, *, session=None, cfg_padre=None, estado_cb=None):
        """VBRP / VBRK / VBRL 2nd run / BKPF: ONE query with the list the consolidation built
        (config["documentos"]). An empty list is a normal empty result: 'Sin datos' placeholder.
        """
        documentos = config.get("documentos") or []
        ruta_salida = os.path.join(config["input_path"], config["filename"])
        if not documentos:
            Archivos.crear_sin_datos(ruta_salida)
            return "vacio", (f"⚠️ {tx['nombre']}: ningún renglón en #N/A trae {tx['unidad']} que consultar "
                             f"→ {config['filename']} (vacío)")
        if estado_cb is not None:
            estado_cb(f"⏳ Descargando {tx['nombre']}: {len(documentos):,} {tx['unidad']}... (no cierre SAP)")
        info = {}
        extra = {} if session is None else {"session": session}
        if tx.get("con_periodo"):
            if config["mode"] == "single":
                extra["DateFrom"] = extra["DateTo"] = config["fecha"]
            else:
                extra["DateFrom"], extra["DateTo"] = config["date_from"], config["date_to"]
            extra["sociedad"] = config["sociedad"]
        inicio = time.time()
        try:
            ok = tx["descarga"](documentos, FolderPath=config["input_path"], FileName=config["filename"],
                                info=info, **extra)
        except DescargaPorRevisarError as e:
            raise DescargaPorRevisarError(f"{e}{self._texto_info(info)}") from e
        detalle = self._texto_info(info)
        if ok:
            return "datos", f"✅ {tx['nombre']}: {len(documentos):,} {tx['unidad']} → {config['filename']}{detalle}"
        if self._clasificar_sin_datos(ruta_salida, inicio) == "vacio":
            return "vacio", (f"⚠️ {tx['nombre']}: SAP no devolvió entradas para la lista consultada "
                             f"→ {config['filename']} (vacío){detalle}")
        return "error", f"⚠️ {tx['nombre']}: error (ver consola){detalle}"

    def _descargar_solicitantes(self, tx, config, resultados, *, session=None, cfg_padre=None, estado_cb=None):
        """FBL5N / KNVP: the 'Solic.' customers of THIS run's VBRK (see `_descargar_con_documentos`)."""
        info = {}
        extra = {} if session is None else {"session": session}

        def descargar(documentos):
            try:
                return tx["descarga"](documentos, FolderPath=config["input_path"],
                                      FileName=config["filename"], info=info, **extra)
            except DescargaPorRevisarError as e:
                raise DescargaPorRevisarError(f"{e}{self._texto_info(info)}") from e

        estado, linea = self._descargar_con_documentos(
            tx, config, resultados, descargar, cfg_padre=cfg_padre, estado_cb=estado_cb)
        return estado, linea + self._texto_info(info)

    def _descargar_vbrl(self, tx, config, resultados, *, session=None, cfg_padre=None, estado_cb=None):
        """VBRL step of the download engine (see `_descargar_con_documentos`)."""
        extra = {} if session is None else {"session": session}
        return self._descargar_con_documentos(
            tx, config, resultados,
            lambda documentos: VBRL_ZSE16(
                documentos, FolderPath=config["input_path"], FileName=config["filename"],
                **extra,
            ),
            cfg_padre=cfg_padre, estado_cb=estado_cb,
        )

    def _descargar_fb03(self, tx, config, resultados, *, session=None, cfg_padre=None, estado_cb=None):
        """FB03 step of the download engine (see `_descargar_con_documentos`)."""
        if config["mode"] == "single":
            desde = hasta = config["fecha"]
        else:
            desde, hasta = config["date_from"], config["date_to"]
        extra = {} if session is None else {"session": session}
        return self._descargar_con_documentos(
            tx, config, resultados,
            lambda documentos: FB03_Documentos(
                documentos, desde, hasta,
                FolderPath=config["input_path"], FileName=config["filename"],
                sociedad=config["sociedad"],
                **extra,
            ),
            cfg_padre=cfg_padre, estado_cb=estado_cb,
        )

    def _descargar_con_documentos(
        self, tx, config, resultados, descargar, *, cfg_padre=None, estado_cb=None
    ):
        """Step of transactions queried ONCE with a list from the file their parent (tx["requiere"])
        downloaded in THIS run: VBRL ← Monivoi 'Factura'; FB03 ← Gallo 'Nº documento' (C2/C6, Global).
        Workers get `cfg_padre`, no `estado_cb` (Tk: main thread only). Returns (estado, summary line)."""
        padre = tx["requiere"]
        nombre_padre = self._nombre_corto(self._TX_POR_CLAVE[padre]) if padre in self._TX_POR_CLAVE else padre
        if padre not in resultados:
            return "error", (
                f"⚠️ {tx['nombre']}: omitido — {nombre_padre} no se descargó en esta corrida"
            )
        if resultados[padre] == "error":
            return "error", (
                f"⚠️ {tx['nombre']}: omitido — {nombre_padre} falló y no hay documentos que consultar"
            )

        if cfg_padre is None:
            cfg_padre = self.gui.get_config(padre)
        ruta_padre  = os.path.join(cfg_padre["input_path"], cfg_padre["filename"])
        ruta_salida = os.path.join(config["input_path"], config["filename"])

        documentos, filas, col_documento = tx["leer_documentos"](ruta_padre)
        if filas == 0:
            Archivos.crear_sin_datos(ruta_salida)
            return "vacio", f"⚠️ {tx['nombre']}: {nombre_padre} sin datos → {config['filename']} (vacío)"
        if not documentos:
            if tx.get("sin_documentos_es_vacio"):
                Archivos.crear_sin_datos(ruta_salida)
                return "vacio", (
                    f"⚠️ {tx['nombre']}: {nombre_padre} trae {filas:,} fila(s) pero ninguna con "
                    f"{tx['filtro']} → {config['filename']} (vacío)"
                )
            return "error", (
                f"⚠️ {tx['nombre']}: {nombre_padre} trae {filas:,} fila(s) pero ningún valor "
                f"utilizable en la columna '{col_documento}'"
            )

        if estado_cb is not None:
            estado_cb(
                f"⏳ Descargando {tx['nombre']}: {len(documentos):,} {tx['unidad']}... (no cierre SAP)"
            )
        inicio = time.time()
        if descargar(documentos):
            return "datos", f"✅ {tx['nombre']}: {len(documentos):,} {tx['unidad']} → {config['filename']}"
        if self._clasificar_sin_datos(ruta_salida, inicio) == "vacio":
            return "vacio", (
                f"⚠️ {tx['nombre']}: SAP no devolvió entradas para la lista consultada "
                f"→ {config['filename']} (vacío)"
            )
        return "error", f"⚠️ {tx['nombre']}: error (ver consola)"

    @staticmethod
    def _clasificar_sin_datos(ruta, desde):
        """A download returned False: "vacio" if it wrote the placeholder since `desde`, else "error"."""
        try:
            # 2 s of slack: some file systems store modification times in 2 s steps.
            if os.path.getmtime(ruta) < desde - 2:
                return "error"
            df = pd.read_csv(ruta, dtype=str, encoding="utf-8-sig", nrows=1)
        except Exception:
            return "error"
        sin_datos = df.empty or (len(df.columns) == 1 and "Sin datos" in str(df.columns[0]))
        return "vacio" if sin_datos else "error"

    def _run_download(self, tx, config, session=None):
        """ONE SAP query → ONE file for the GUI period (a day or a range); True with data."""
        if config["mode"] == "single":
            desde = hasta = config["fecha"]
        else:
            desde, hasta = config["date_from"], config["date_to"]
        extra = {} if session is None else {"session": session}
        return tx["descarga"](
            DateFrom=desde,
            DateTo=hasta,
            FolderPath=config["input_path"],
            FileName=config["filename"],
            sociedad=config["sociedad"],
            **extra,
        )

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

    @staticmethod
    def _validar_venta(nombre, config):
        """A sales button (`nombre`: 'Venta DSD' / 'Venta EIAP') needs a range that still makes sense once
        SAP is queried up to the end date minus one day (utils.fecha_efectiva_eiap). Error text, or None.
        """
        if config["mode"] != "range":
            return (
                f"'{nombre}' trabaja con el intervalo \"Varios días\".\n\n"
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
                f"En '{nombre}' la Fecha Hasta se resta 1 día, así que el periodo "
                "debe abarcar al menos 2 días (ej. 01.01.2026 a 02.01.2026 consulta sólo el 01.01.2026)."
            )
        return None

    @staticmethod
    def _nota_venta(config, cuentas, clases):
        """Confirmation text: the period SAP will REALLY be queried with, and the accounts and classes."""
        hasta = fecha_efectiva_eiap(config["date_to"]).strftime(FORMATO_FECHA)
        return (
            f"Consulta en SAP : {config['date_from']} — {hasta}\n"
            "(a la Fecha Hasta se le resta 1 día)\n"
            f"Cuentas         : {', '.join(cuentas)}\n"
            f"Clases de doc.  : {', '.join(clases)}\n\n"
        )
