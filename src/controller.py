"""Validación Factura Global — controller: glue between the GUI (interfaz_GUI.py), the SAP
downloads (Descargas_SAP.py) and the consolidation (Consolidacion.py).

The main download button runs EVERYTHING in one go. 1st run: Gallo, Monivoi, VBRL and FB03 one after
another (default) or, with the GUI checkbox, IN PARALLEL by passes: Gallo ‖ Monivoi, then VBRL ‖ FB03
(VBRL needs the Monivoi file and FB03 the Gallo file of pass 1). From those files the lists the 2nd run
pastes into SAP (what the #N/A rows of 'TD.Diferencias_Referencias' carry) are computed IN MEMORY and the
2nd run (VBRK, VBRP, VBRL 2nd run, KNA1, KNVP, BKPF, Monivoi 2nd run) follows: one after another, or, with
the same checkbox, IN PARALLEL by passes of up to THREE windows (PASADAS_2DA): VBRK ‖ VBRP ‖ VBRL 2nd run,
then KNA1 ‖ KNVP ‖ BKPF, then Monivoi 2nd run (KNA1 and KNVP paste the 'Solic.' column of the VBRK of
pass 1; Monivoi 2nd run the 'List.fact.' lists of the VBRL 2nd run). The one consolidation button then
builds the whole workbook, 'Diferencias' and 'Agrupadores' included.

The "Corrida completa" button does both steps in a row (download, then consolidation of what THIS run
downloaded) and so does the command line (`--corrida-completa`, see corrida_completa.py). Both call the same
dialog-free engines (`_correr_corrida_completa`); only the buttons' wrappers show message boxes.
"""

import os
import time
import traceback
from datetime import datetime, timedelta
from tkinter import messagebox, filedialog

import pandas as pd

from Descargas_SAP import (
    Gallo_FAGLL03, Monivoi_ZLMXCOM, VentaEIAP_FAGLL03, VentaDSD_FAGLL03,
    VBRL_ZSE16, FB03_Documentos, DescargaPorRevisarError,
    VBRP_ZSE16, VBRK_ZSE16, VBRL2_ZSE16, KNVP_ZSE16, BKPF_ZSE16, KNA1_ZSE16, Monivoi2_ZLMXCOM,
    Archivos, Excel, Paralelo, MAX_VENTANAS_SAP,
)
from Consolidacion import (
    consolidar_gallo_monivoi, consolidar_venta_dsd, consolidar_venta_eiap, leer_documentos_monivoi,
    leer_documentos_gallo_fb03, leer_solicitantes_vbrk, calcular_listas_segunda_corrida,
    leer_listas_vbrl2, EtapaDocumentosFB03,
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

    # 2nd run in parallel: PASSES of up to 3 windows (a 4th stays free); each starts when the previous one ends.
    PASADAS_2DA = (("vbrk", "vbrp", "vbrl2"), ("kna1", "knvp", "bkpf"), ("monivoi2",))

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
        # 2nd run: lists from the consolidation ("lista_de"); KNA1 / KNVP read VBRK's 'Solic.', Monivoi 2 VBRL_2's.
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
        self._TX_KNA1 = {
            "clave": "kna1", "nombre": "KNA1 (ZSE16)", "ejecutar": self._descargar_solicitantes,
            "descarga": KNA1_ZSE16, "requiere": "vbrk", "leer_documentos": leer_solicitantes_vbrk,
            "unidad": "cliente(s)",
            "lista": "los solicitantes (columna 'Solic.') del VBRK de esta corrida",
        }
        self._TX_KNVP = {
            "clave": "knvp", "nombre": "KNVP (ZSE16)", "ejecutar": self._descargar_solicitantes,
            "descarga": KNVP_ZSE16, "requiere": "vbrk", "leer_documentos": leer_solicitantes_vbrk,
            "unidad": "cliente(s)",
            "lista": "los solicitantes (columna 'Solic.') del VBRK de esta corrida",
        }
        self._TX_MONIVOI2 = {
            "clave": "monivoi2", "nombre": "Monivoi 2a corrida (ZLMXCOM_TRN_MONINVOI)",
            "ejecutar": self._descargar_monivoi2, "requiere": "vbrl2",
            "leer_documentos": leer_listas_vbrl2, "unidad": "lista(s) de factura",
            "lista": "las listas de factura ('List.fact.') del VBRL 2a corrida recién descargado",
            "sin_documentos_es_vacio": True, "filtro": "Lista de factura ('List.fact.')",
        }
        self._TX_POR_CLAVE = {
            tx["clave"]: tx
            for tx in (self._TX_GALLO, self._TX_MONIVOI, self._TX_DSD, self._TX_EIAP, self._TX_VBRL,
                       self._TX_FB03,
                       self._TX_VBRP, self._TX_VBRK, self._TX_VBRL2, self._TX_BKPF,
                       self._TX_KNA1, self._TX_KNVP, self._TX_MONIVOI2)
        }

        self._ocupado = False

        self.gui.on_download_gallo   = self.execute_download_gallo
        self.gui.on_download_monivoi = self.execute_download_monivoi
        self.gui.on_download_dsd     = self.execute_download_dsd
        self.gui.on_download_eiap    = self.execute_download_eiap
        self.gui.on_download_todos   = self.execute_download_todos
        self.gui.on_consolidar = self.execute_consolidacion
        self.gui.on_corrida_completa = self.execute_corrida_completa

    def execute_download_gallo(self):
        self._execute_download([self._TX_GALLO], "Gallo")

    def execute_download_monivoi(self):
        self._execute_download([self._TX_MONIVOI], "Monivoi")

    def execute_download_dsd(self):
        self._execute_download([self._TX_DSD], "Venta DSD")

    def execute_download_eiap(self):
        self._execute_download([self._TX_EIAP], "Venta EIAP")

    @staticmethod
    def _texto_periodo(base):
        """'Fecha    : 03.11.2026' or 'Periodo  : 01.11.2026 — 03.11.2026' (confirmations and summaries)."""
        return (
            f"Fecha    : {base['fecha']}" if base["mode"] == "single"
            else f"Periodo  : {base['date_from']} — {base['date_to']}"
        )

    def _planear_descarga_todos(self):
        """What the main download needs, read ONCE: Tk may only be touched by its own thread and the GUI can
        be edited while SAP works, so the run is configured from this plan and never from the GUI again.
        → (plan, problema). `problema` (text for the user: two transactions that would write the same file)
        comes with `plan` None. A wrong dependency between transactions is a programming error: ValueError.
        plan: base (Gallo's config), configs1, txs1, txs2, pasadas2, paralelo, n_ventanas1, n_ventanas2."""
        txs1 = [self._TX_POR_CLAVE[clave] for clave, _ in CORRIDAS_DESCARGA[1]]
        pasadas2 = self._txs_segunda_corrida()
        txs2 = [tx for pasada in pasadas2 for tx in pasada]
        self._validar_dependencias(txs1)
        self._validar_pasadas(pasadas2)        # VBRK in an earlier pass than KNA1 / KNVP, in both modes

        configs1 = {tx["clave"]: self.gui.get_config(tx["clave"]) for tx in txs1}
        base = configs1["gallo"]
        # The 2nd run's file names do not depend on its lists: none of the eleven may repeat.
        nombres2 = self._configs_segunda_corrida(base, {"referencias": [], "documentos": []}, txs2)
        problema = self._validar_paralelo(txs1 + txs2, {**configs1, **nombres2})
        if problema:
            return None, problema

        paralelo = self._paralelo_gui()
        n_ventanas1 = self._ventanas_necesarias(txs1) if paralelo else 1
        n_ventanas2 = self._ventanas_de_pasadas(pasadas2) if paralelo else 1
        return {
            "base": base, "configs1": configs1, "txs1": txs1, "txs2": txs2, "pasadas2": pasadas2,
            "paralelo": n_ventanas1 > 1, "n_ventanas1": n_ventanas1, "n_ventanas2": n_ventanas2,
        }, None

    def _texto_orden_y_consultas(self, plan):
        """The two blocks of the confirmation that describe HOW the download goes: the order of the
        transactions and the lists each one asks SAP for → (orden, consultas)."""
        txs1, txs2, pasadas2 = plan["txs1"], plan["txs2"], plan["pasadas2"]
        if plan["paralelo"]:
            orden = (
                "En paralelo, por pasadas: las transacciones de una pasada corren a la vez, cada una en su propia "
                "ventana de SAP, y cada pasada empieza cuando termina la anterior:\n\n"
                + self._texto_orden("1a corrida", txs1, self._pasadas(txs1, plan["n_ventanas1"]), True)
                + "\n" + self._texto_orden("2a corrida", txs2, pasadas2, True)
                + "\nUso las ventanas de SAP que ya estén abiertas...\n"
                "*IMPORTANTE:* No trabajar en esas ventanas mientras corre y en la medida de lo posible evitar hacer click en ellas.\n\n"
            )
        else:
            orden = (
                "Se ejecutan una tras otra:\n\n"
                + self._texto_orden("1a corrida", txs1, None, False)
                + "\n" + self._texto_orden("2a corrida", txs2, None, False) + "\n"
            )
        consultas = "".join(
            f"{self._nombre_corto(tx)} consulta {tx['lista']}.\n" for tx in txs1 if tx.get("lista")
        )
        return orden, consultas

    def execute_download_todos(self):
        """Main download button: EVERYTHING in one run, with ONE confirmation and ONE summary. The 1st run
        (Gallo, Monivoi, VBRL, FB03) goes first; the lists the 2nd run asks SAP for are then computed IN
        MEMORY from its files (calcular_listas_segunda_corrida) and the 2nd run (PASADAS_2DA) follows with
        them. The GUI checkbox decides HOW each run goes: one after another or in parallel by passes. The
        2nd run is skipped, saying so, when the 1st one did not leave what its lists come from."""
        if self._ocupado:
            return
        if not self.gui.validate_dates():
            return

        plan, problema = self._planear_descarga_todos()
        if problema:
            messagebox.showerror("Descarga de SAP", problema)
            return
        base, txs1, txs2 = plan["base"], plan["txs1"], plan["txs2"]

        orden, consultas = self._texto_orden_y_consultas(plan)
        confirm = messagebox.askyesno(
            "Confirmar descarga (todas las transacciones)",
            f"¿Iniciar la descarga de las {len(txs1) + len(txs2)} transacciones?\n\n"
            f"Sociedad : {base['sociedad']}\n"
            f"{self._texto_periodo(base)}\n\n"
            f"{orden}"
            f"{consultas}"
            "Las listas de la 2a corrida se calculan solas con los archivos de la 1a (lo que quede en #N/A "
            "en 'TD.Diferencias_Referencias'): si la 1a corrida no deja bien el Gallo o el VBRL, la 2a no se ejecuta.\n\n"
            "Mientras corren NO copiar ni pegar nada!! Puesto que el portapapeles se usa para pasarle las listas a SAP...\n\n"
            f"Se guardará en:\n{base['input_path']}"
        )
        if not confirm:
            return

        self._ocupado = True
        estado_final = None
        salida = {"resumen": []}        # lines of what already ran: an unexpected error must not hide them
        try:
            self.gui.disable_buttons()
            self._ejecutar_descarga_todos(plan, salida)
            estado_final = salida["estado_final"]
            mostrar = messagebox.showinfo if salida["todo_ok"] else messagebox.showwarning
            mostrar(
                "Resumen de descarga (todas las transacciones)",
                "Proceso terminado:\n\n" + "\n".join(salida["resumen"])
                + f"\n\n⏱️ Tiempo total: {_fmt_duracion(salida['total'])}"
                + f"\nRuta: {base['input_path']}"
            )
        except Exception as e:
            estado_final = None
            self.gui.set_status("❌ Error en la descarga")
            self._marcar_paso("descarga", "error")
            ya_corrio = ("\n\nLo que ya había terminado:\n" + "\n".join(salida["resumen"])) if salida["resumen"] else ""
            messagebox.showerror("Error", f"Ocurrió un error inesperado:\n\n{e}{ya_corrio}")
        finally:
            self._ocupado = False
            self.gui.enable_buttons()
            if estado_final is None:
                self.gui.set_status("✓ Listo para comenzar")

    def _ejecutar_descarga_todos(self, plan, salida):
        """The main download itself, with NO dialogs (the 'Descargar todo' and 'Corrida completa' buttons and the
        command line share it): 1st run, the lists of the 2nd computed from its files, 2nd run. Marks the rail
        and the status. Fills `salida` AS IT GOES (so a caller that catches an unexpected error still has the
        lines of what ran): resumen (lines), then todo_ok, listas (None: the 2nd run did not run), total
        (seconds), estado_final, configs1, configs2 (None without a 2nd run), resultados1, resultados2."""
        base, txs1, txs2, pasadas2 = plan["base"], plan["txs1"], plan["txs2"], plan["pasadas2"]
        configs1 = plan["configs1"]
        self._marcar_paso("descarga", "en_curso")
        os.makedirs(base["input_path"], exist_ok=True)
        inicio_total = time.time()

        lineas1, resultados1, duraciones1, notas1 = self._correr_etapa(txs1, configs1, plan["n_ventanas1"])
        salida["resumen"] = self._resumen_etapa("1a corrida", txs1, lineas1, duraciones1, notas1)
        salida.update(configs1=configs1, resultados1=resultados1, configs2=None, resultados2={})
        todo_ok = all(resultados1.get(tx["clave"]) == "datos" for tx in txs1)

        listas, avisos = self._listas_segunda(configs1, resultados1)
        if listas is None:
            todo_ok = False
            salida["resumen"] += [""] + avisos
        else:
            configs2 = self._configs_segunda_corrida(base, listas, txs2)
            self.gui.set_status("⏳ Iniciando la 2a corrida... (no cierre SAP)")
            lineas2, resultados2, duraciones2, notas2 = self._correr_etapa(
                txs2, configs2, plan["n_ventanas2"], pasadas=pasadas2)
            todo_ok = (todo_ok and not avisos
                       and all(resultados2.get(tx["clave"]) == "datos" for tx in txs2))
            salida["resumen"] += ([""] + self._texto_listas(listas) + avisos
                                  + self._resumen_etapa("2a corrida", txs2, lineas2, duraciones2, notas2))
            salida.update(configs2=configs2, resultados2=resultados2)

        if listas is None:
            estado_final = "⚠️ Descarga de la 1a corrida terminada; la 2a corrida NO se ejecutó"
        elif todo_ok:
            estado_final = "✅ ¡Descarga completada (1a + 2a corrida)!"
        else:
            estado_final = "⚠️ Descarga terminada con avisos"
        self.gui.set_status(estado_final)
        self._marcar_paso("descarga", "listo" if todo_ok else "parcial")
        salida.update(todo_ok=todo_ok, listas=listas, total=time.time() - inicio_total,
                      estado_final=estado_final)
        return salida

    def _paralelo_gui(self):
        """True if the GUI's 'en paralelo' switch is on (False for a GUI without one)."""
        usar_paralelo = getattr(self.gui, "usar_paralelo", None)
        return bool(usar_paralelo()) if callable(usar_paralelo) else False

    def _marcar_paso(self, paso, estado):
        """Report the state of a step of the process to the GUI's progress rail ('descarga',
        'consolidacion' x 'en_curso', 'listo', 'parcial', 'error'). `paso` None (a single-transaction
        download) and a GUI without `marcar_paso` (older GUI, test doubles) are skipped. Main thread
        only: it touches Tk."""
        marcar = getattr(self.gui, "marcar_paso", None)
        if paso is not None and callable(marcar):
            marcar(paso, estado)

    def _texto_orden(self, titulo, txs, pasadas, paralelo):
        """The confirmation's lines for ONE run: its passes (parallel) or its transactions in order."""
        if paralelo:
            cuerpo = "".join(
                f"  Pasada {i}: {' ‖ '.join(self._nombre_corto(tx) for tx in pasada)}\n"
                for i, pasada in enumerate(pasadas, 1)
            )
        else:
            cuerpo = "".join(f"  {i}) {tx['nombre']}\n" for i, tx in enumerate(txs, 1))
        return f"{titulo}:\n{cuerpo}"

    def _correr_etapa(self, txs, configs, n_ventanas, pasadas=None):
        """Run `txs` with their `configs` (read before the run, never the GUI): in parallel by passes when
        `n_ventanas` > 1 (`pasadas`: fixed passes, else split by dependency), else one after another.
        → (lines, estados, seconds, notes) as _descargar_en_paralelo; no notes when sequential."""
        if n_ventanas > 1:
            return self._descargar_en_paralelo(txs, configs, n_ventanas, pasadas=pasadas)
        lineas, resultados, duraciones = self._correr_secuencial(txs, configs)
        return lineas, resultados, duraciones, []

    @staticmethod
    def _resumen_etapa(titulo, txs, lineas, duraciones, notas=()):
        """['— titulo —', '✅ Gallo...: file (45 s)', ..., notes of the run]: one line per transaction of the run,
        then ITS notes (parallel timings, window count): each run numbers its own passes, so they stay apart."""
        return ([f"— {titulo} —"]
                + [f"{lineas[tx['clave']]} ({_fmt_duracion(duraciones[tx['clave']])})" for tx in txs]
                + list(notas))

    @staticmethod
    def _texto_listas(listas):
        """Summary lines about the lists the 2nd run was asked with (what was queried, what was left out)."""
        lineas = [
            f"ℹ️ Listas de la 2a corrida: {len(listas['referencias']):,} referencia(s) · "
            f"{len(listas['documentos']):,} documento(s) en #N/A"
        ]
        excluidas = listas["excluidas_fp"]["renglones"]
        if excluidas:
            lineas.append(f"ℹ️ Fuera {excluidas:,} renglón(es) cuya Referencia de pago empieza con 'FP': "
                          "ni su referencia ni su Nº documento se consultan.")
        if not listas["referencias"] and not listas["documentos"]:
            lineas.append("ℹ️ Ningún renglón quedó en #N/A: la 2a corrida sólo dejó archivos vacíos "
                          "(para que la consolidación los encuentre).")
        lineas += [f"⚠️ {aviso}" for aviso in listas.get("avisos") or []]
        return lineas

    def _listas_segunda(self, configs, resultados):
        """The lists of the 2nd run, computed IN MEMORY from the files the 1st run just left → (listas, avisos).
        Only files THIS run wrote count (`resultados`): an older file of a failed download is never used. Gallo
        and VBRL are required (`listas` None, `avisos` saying why); without FB03 the lists come from the
        Facturas table only (an aviso). Never raises: the 1st run's summary must survive."""
        rutas = {clave: os.path.join(c["input_path"], c["filename"]) for clave, c in configs.items()}

        def dejado(clave):
            return resultados.get(clave) in ("datos", "vacio") and os.path.exists(rutas[clave])

        faltan = [configs[clave]["filename"] for clave in ("gallo", "vbrl") if not dejado(clave)]
        if faltan:
            return None, [
                "⚠️ La 2a corrida NO se ejecutó: sus listas salen del Gallo y del VBRL y no quedó bien "
                "descargado: " + ", ".join(faltan) + ". Revisa la 1a corrida y vuelve a descargar."
            ]
        ruta_fb03 = rutas["fb03"] if dejado("fb03") else None
        avisos = []
        if ruta_fb03 is None:
            avisos.append(
                f"⚠️ {configs['fb03']['filename']} no quedó bien descargado: las listas de la 2a corrida "
                "salen SÓLO de la tabla Facturas (sin 'Notas de crédito' ni 'PEP Vouchers')."
            )
        try:
            listas = calcular_listas_segunda_corrida(
                rutas["gallo"], rutas["vbrl"], ruta_fb03, callback_status=self.gui.set_status)
        except (FileNotFoundError, ValueError) as e:
            return None, [f"⚠️ La 2a corrida NO se ejecutó: no pude calcular sus listas: {e}"]
        except Exception as e:
            return None, [f"❌ La 2a corrida NO se ejecutó: error inesperado al calcular sus listas: {e}"]
        return listas, avisos

    @staticmethod
    def _validar_dependencias(txs):
        """A dependent transaction (VBRL needs THIS run's Monivoi) goes after its parent: no stale files.
        ValueError otherwise (a programming error, not something the user can fix)."""
        for pos, tx in enumerate(txs):
            requiere = tx.get("requiere")
            if requiere and requiere not in [t["clave"] for t in txs[:pos]]:
                raise ValueError(
                    f"{tx['nombre']} necesita que '{requiere}' se descargue antes, en la misma corrida."
                )

    def _execute_download(self, txs, etiqueta, paralelo=False):
        """Run `txs` (one after another, or in parallel by passes) and show ONE summary."""
        if not txs:
            raise ValueError("No se indicó ninguna transacción a descargar.")
        self._validar_dependencias(txs)
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
                    "Uso las ventanas de SAP que ya estén abiertas...\n"
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
        [[VBRK, VBRP, VBRL2], [KNA1, KNVP, BKPF]] → [[VBRK, VBRP], [VBRL2], [KNA1, KNVP], [BKPF]].
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
        ("vbrk", "VBRK"), ("kna1", "KNA1"), ("vbrp", "VBRP"),
        ("knvp", "KNVP"), ("vbrl2", "VBRL (2a corrida)"), ("bkpf", "BKPF"),
    )
    _ETIQUETA_MONIVOI2 = "Monivoi (2a corrida)"
    # A 2nd-run file older than this many seconds was asked for an older 1st run (lookups may miss).
    _TOLERANCIA_ANTIGUEDAD = 2

    def execute_consolidacion(self):
        """Main consolidation button: ONE workbook with everything — Gallo, Monivoi, VBRL, FB03 and the 2nd
        run's sheets, 'Diferencias' and 'Agrupadores' — offering to pick by hand a file that is not where it
        is expected. With NEITHER Gallo nor Monivoi in the folder (only the sales buttons were used) it
        consolidates the sales alone. The VBRL is NOT optional: 'Diferencias' lists the #N/A rows of the
        tables it feeds. A missing Monitoreo_2 is not asked for (VBRL_2 had no lists either): blank tab."""
        if self._ocupado:
            return
        if not self.gui.validate_dates():
            return

        cfg_gallo   = self.gui.get_config("gallo")
        cfg_monivoi = self.gui.get_config("monivoi")
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
        ruta_vbrl = self._localizar_vbrl(input_path, self.gui.get_config("vbrl")["filename"], ruta_monivoi)
        if not ruta_vbrl:
            return
        ruta_fb03, cancelar = self._resolver_fb03(
            input_path, self.gui.get_config("fb03")["filename"], ruta_gallo)
        if cancelar:
            return

        rutas_segunda = {}
        for clave, etiqueta in self._ETIQUETAS_SEGUNDA:
            ruta = self._localizar_archivo(input_path, self.gui.get_config(clave)["filename"], etiqueta)
            if not ruta:
                return
            rutas_segunda[clave] = ruta
        ruta_monivoi2 = os.path.join(input_path, self.gui.get_config("monivoi2")["filename"])
        if not os.path.exists(ruta_monivoi2):
            ruta_monivoi2 = None
        rutas_segunda["monivoi2"] = ruta_monivoi2

        padres = [ruta for ruta in (ruta_gallo, ruta_monivoi, ruta_vbrl, ruta_fb03) if ruta]
        anteriores = self._anteriores_segunda(padres, rutas_segunda)
        if anteriores:
            if not messagebox.askyesno(
                "Descargas de la 2a corrida anteriores a la 1a",
                "Estos archivos de la 2a corrida se descargaron ANTES que el archivo del que salieron sus "
                "listas:\n\n"
                + "".join(f"  • {linea}\n" for linea in anteriores)
                + "\nProbablemente no cubren los documentos actuales y 'Diferencias' mostraría #N/A "
                "(o 'Sin UUID') de más.\n\n¿Consolidar de todos modos?",
            ):
                return

        try:
            ruta_output = self.gui.get_consolidacion_path()
        except OSError as e:
            messagebox.showerror("Error", f"No se pudo usar la carpeta de consolidación:\n{e}")
            return
        nombre_salida = self._nombre_salida(cfg_gallo)
        ruta_final = os.path.join(ruta_output, nombre_salida)

        linea_fb03 = (f"  📑 FB03    : {os.path.basename(ruta_fb03)}\n" if ruta_fb03 else
                      "  📑 FB03    : (sin FB03: 'Diferencias' sólo con las Facturas)\n")
        lineas_segunda = "".join(
            f"  • {etiqueta:<17}: {os.path.basename(rutas_segunda[clave])}\n"
            for clave, etiqueta in self._ETIQUETAS_SEGUNDA)
        nombre_monivoi2 = os.path.basename(ruta_monivoi2) if ruta_monivoi2 else "(sin archivo: Monitor (2) en blanco)"
        lineas_segunda += f"  • {self._ETIQUETA_MONIVOI2:<17}: {nombre_monivoi2}\n"
        reemplaza = ("\n⚠️ Ya existe y se REEMPLAZA." if os.path.exists(ruta_final) else "")
        if not messagebox.askyesno(
            "Confirmar consolidación",
            "¿Consolidar estos archivos en un solo Excel?\n\n"
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
            _, texto, _ = self._consolidar_archivos(
                ruta_gallo, ruta_monivoi, ruta_vbrl, ruta_fb03, rutas_segunda, ruta_output, nombre_salida)
            messagebox.showinfo("Consolidación completada", texto)
        except (FileNotFoundError, ValueError) as e:
            messagebox.showerror("Error en la consolidación", str(e))
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error inesperado:\n\n{e}")
        finally:
            self._ocupado = False
            self.gui.enable_buttons()
            if "completada" not in self.gui.status_var.get():
                self.gui.set_status("✓ Listo para comenzar")

    @staticmethod
    def _nombre_salida(cfg_gallo):
        """'Consolidado_<Gallo file without extension>.xlsx': ONE name for the button and the full run."""
        return f"Consolidado_{os.path.splitext(cfg_gallo['filename'])[0]}.xlsx"

    def _texto_resumen_consolidacion(self, resumen):
        """The message that closes a consolidation → (text, con_avisos): where the workbook is, the control
        lines of every stage and the warnings the stakeholder should read. `con_avisos` also counts an empty
        Monivoi / VBRL file."""
        avisos = self._avisos_consolidacion(resumen)
        nota_avisos = ("\n" + "\n".join(f"⚠️ {a}" for a in avisos)) if avisos else ""
        sin_datos = []
        if resumen["monivoi_sin_datos"]:
            sin_datos.append("⚠️ El archivo de Monivoi venía sin datos; su pestaña quedó vacía.")
        if resumen.get("vbrl_sin_datos"):
            sin_datos.append("⚠️ El archivo de VBRL venía sin datos; su pestaña quedó vacía.")
        nota_sin_datos = ("\n" + "\n".join(sin_datos)) if sin_datos else ""
        d = resumen["diferencias"]
        por_tipo = d["por_tipo"]
        semaforo = d["semaforo"]
        texto = (
            f"Archivo generado:\n{resumen['ruta']}\n\n"
            f"{self._lineas_primera_consolidacion(resumen)}\n"
            f"  • Renglones en 'Diferencias': {d['filas']:,}"
            f"  (Facturas {por_tipo.get('Factura', 0):,} · NC {por_tipo.get('NC', 0):,}"
            f" · PV {por_tipo.get('PV', 0):,})\n"
            f"  • Σ Monto                    : {d['importe']:,.2f}\n"
            f"  • Sin No cliente (VBRK)      : {d['sin_vbrk']:,}\n"
            f"  • Sin Lista de factura       : {d['sin_lista']:,}\n"
            f"  • Documentos anulados        : {d['anuladas']:,}\n"
            f"  • Cliente Global / Vendedor / Individual: "
            f"{d['cliente_global']:,} / {d['vendedor']:,} / {d['cliente_individual']:,}\n"
            f"  • Estatus 1 🟢 / 🟡 / 🔴     : {semaforo['verde']:,} / "
            f"{semaforo['amarillo']:,} / {semaforo['rojo']:,}"
            f"  (SIN REGLA {d['sin_regla']:,})\n"
            f"  • W.A BAMA / OXXO: UUID {d['wa_uuid']:,} · Sin UUID {d['wa_sin_uuid']:,}"
            f"{nota_sin_datos}{nota_avisos}"
        )
        return texto, bool(avisos or sin_datos)

    def _consolidar_archivos(self, ruta_gallo, ruta_monivoi, ruta_vbrl, ruta_fb03, rutas_segunda,
                             ruta_output, nombre_salida):
        """The consolidation itself, with NO dialogs (the 'Consolidar' button and the full run share it): marks
        the rail, shows the status and → (resumen, texto, con_avisos). Whatever the consolidator raises
        (FileNotFoundError / ValueError: the files are wrong; anything else: a bug) is raised again AFTER the
        step is marked as an error, for the caller to say it."""
        self._marcar_paso("consolidacion", "en_curso")
        try:
            resumen = consolidar_gallo_monivoi(
                ruta_gallo, ruta_monivoi, ruta_output,
                nombre_salida=nombre_salida,
                callback_status=self.gui.set_status,
                ruta_vbrl=ruta_vbrl,
                ruta_fb03=ruta_fb03,
                rutas_segunda=rutas_segunda,
            )
            texto, con_avisos = self._texto_resumen_consolidacion(resumen)
        except Exception:
            self.gui.set_status("❌ Error en la consolidación")
            self._marcar_paso("consolidacion", "error")
            raise
        self.gui.set_status("✅ ¡Consolidación completada!")
        self._marcar_paso("consolidacion", "parcial" if con_avisos else "listo")
        return resumen, texto, con_avisos

    # ------------------------------------------------------------------
    # Full run: the download of everything and, right after it, the consolidation
    # ------------------------------------------------------------------
    # What THIS run must have left for its consolidation to make sense: 9 files, the four of the 1st run (Gallo,
    # Monivoi, VBRL, FB03) and five of the 2nd one (VBRK, VBRP, KNA1, KNVP, BKPF). The other two of the 2nd run,
    # VBRL_2 and Monitoreo_2, are optional (an empty sheet / a blank tab, with a warning), so a failed one only
    # lowers the final state to "avisos".
    _OBLIGATORIOS_CORRIDA = ("gallo", "monivoi", "vbrl", "fb03", "vbrk", "vbrp", "kna1", "knvp", "bkpf")
    _OPCIONALES_CORRIDA = ("vbrl2", "monivoi2")

    def _planear_corrida_completa(self):
        """`_planear_descarga_todos` plus what only the consolidation needs, also read NOW (the GUI can be edited
        while SAP works): the output folder and the workbook name. → (plan, problema) as the download's."""
        plan, problema = self._planear_descarga_todos()
        if problema:
            return None, problema
        try:
            plan["ruta_output"] = self.gui.get_consolidacion_path()
        except OSError as e:
            return None, f"No se pudo usar la carpeta de consolidación:\n{e}"
        plan["nombre_salida"] = self._nombre_salida(plan["base"])
        return plan, None

    def _archivos_de_esta_corrida(self, salida):
        """The files the consolidation of the full run reads, ONLY those THIS run left (the state of each
        download says so: an older file of a failed download is never used) → (rutas, faltan). `rutas`:
        {clave: path, or None for an optional one that is missing}; `faltan`: the names of the required
        transactions that left no usable file."""
        configs = {**salida["configs1"], **(salida["configs2"] or {})}
        resultados = {**salida["resultados1"], **salida["resultados2"]}
        rutas, faltan = {}, []
        for clave in self._OBLIGATORIOS_CORRIDA + self._OPCIONALES_CORRIDA:
            cfg = configs.get(clave)
            ruta = os.path.join(cfg["input_path"], cfg["filename"]) if cfg else None
            usable = bool(ruta) and resultados.get(clave) in ("datos", "vacio") and os.path.isfile(ruta)
            rutas[clave] = ruta if usable else None
            if not usable and clave in self._OBLIGATORIOS_CORRIDA:
                faltan.append(self._nombre_corto(self._TX_POR_CLAVE[clave]))
        return rutas, faltan

    def _consolidar_corrida(self, plan, salida):
        """Second half of the full run: consolidate what the first half left → `salida["consolidacion"]` =
        {hecha, ruta, texto, con_avisos, reemplazo, motivo}. `motivo` says why nothing was consolidated. Never
        raises: a failure here must not hide the download's summary."""
        cons = {"hecha": False, "ruta": None, "texto": "", "con_avisos": False, "reemplazo": False,
                "motivo": None}
        salida["consolidacion"] = cons
        rutas, faltan = self._archivos_de_esta_corrida(salida)
        if faltan:
            cons["motivo"] = (
                "La consolidación NO se ejecutó: estas descargas no dejaron un archivo bueno en esta corrida: "
                + ", ".join(faltan) + ". Revisa la descarga y vuelve a intentar (el Excel necesita los archivos "
                "de ESTA corrida, no los de una anterior)."
            )
            self.gui.set_status("⚠️ Descarga con fallas: la consolidación NO se ejecutó")
            self._marcar_paso("consolidacion", "error")
            return cons
        rutas_segunda = {clave: rutas[clave] for clave in
                         tuple(c for c, _ in self._ETIQUETAS_SEGUNDA) + ("monivoi2",)}
        cons["reemplazo"] = os.path.exists(os.path.join(plan["ruta_output"], plan["nombre_salida"]))
        try:
            resumen, texto, con_avisos = self._consolidar_archivos(
                rutas["gallo"], rutas["monivoi"], rutas["vbrl"], rutas["fb03"], rutas_segunda,
                plan["ruta_output"], plan["nombre_salida"])
        except (FileNotFoundError, ValueError) as e:
            cons["motivo"] = f"Error en la consolidación: {e}"
        except Exception as e:
            print(traceback.format_exc())
            cons["motivo"] = f"Error inesperado en la consolidación: {e}"
        else:
            cons.update(hecha=True, ruta=resumen["ruta"], texto=texto, con_avisos=con_avisos)
        return cons

    def _correr_corrida_completa(self, plan, salida):
        """Both steps of the workflow in a row, with NO dialogs: `_ejecutar_descarga_todos`, then the
        consolidation of the files that download left. Fills `salida` (see the download's) plus: consolidacion,
        total_corrida (seconds), estado ('completa' | 'avisos' | 'incompleta') and estado_final (the status
        line). An unexpected error in the DOWNLOAD propagates (`_corrida_completa_protegida` catches it)."""
        inicio = time.time()
        salida["paso"] = "descarga"
        self._ejecutar_descarga_todos(plan, salida)
        salida["paso"] = "consolidacion"
        cons = self._consolidar_corrida(plan, salida)
        salida["total_corrida"] = time.time() - inicio
        if not cons["hecha"]:
            salida["estado"] = "incompleta"
            salida["estado_final"] = "⚠️ Corrida completa terminada: la consolidación NO se ejecutó"
        elif salida["todo_ok"] and not cons["con_avisos"]:
            salida["estado"] = "completa"
            salida["estado_final"] = "✅ ¡Corrida completa terminada (descarga + consolidación)!"
        else:
            salida["estado"] = "avisos"
            salida["estado_final"] = "⚠️ Corrida completa terminada con avisos"
        self.gui.set_status(salida["estado_final"])
        return salida

    def _corrida_completa_protegida(self, plan, salida):
        """`_correr_corrida_completa` that never raises: an unexpected error becomes estado 'error' (with the
        text in `salida["error"]`), the status line and the rail's red step, and `salida["resumen"]` keeps the
        lines of whatever had already run."""
        try:
            self._correr_corrida_completa(plan, salida)
        except Exception as e:
            print(traceback.format_exc())
            self.gui.set_status("❌ Error en la corrida completa")
            self._marcar_paso(salida.get("paso", "descarga"), "error")
            salida.update(estado="error", estado_final="❌ Error en la corrida completa",
                          error=f"Ocurrió un error inesperado: {e}")
        return salida

    def texto_final_corrida(self, salida):
        """(título, cuerpo) of what closes a full run: the dialog of the button and the log of the command line."""
        estado = salida.get("estado", "error")
        titulo = {
            "completa": "Corrida completa terminada",
            "avisos": "Corrida completa terminada con avisos",
            "incompleta": "Corrida completa: la consolidación no se ejecutó",
        }.get(estado, "Error en la corrida completa")
        lineas = []
        if salida.get("error"):
            lineas += [salida["error"], ""]
        if salida.get("resumen"):
            lineas += ["— Descarga —"] + salida["resumen"]
            if "total" in salida:
                lineas += ["", f"⏱️ Descarga: {_fmt_duracion(salida['total'])}"]
        cons = salida.get("consolidacion")
        if cons is not None:
            lineas += ["", "— Consolidación —"]
            if cons["hecha"]:
                if cons["reemplazo"]:
                    lineas.append("ℹ️ Ya existía un archivo con ese nombre y se reemplazó.")
                lineas.append(cons["texto"])
            else:
                lineas.append(f"⚠️ {cons['motivo']}")
        if "total_corrida" in salida:
            lineas += ["", f"⏱️ Tiempo total: {_fmt_duracion(salida['total_corrida'])}"]
        return titulo, "\n".join(lineas).strip("\n")

    def execute_corrida_completa(self):
        """'Corrida completa' button: the download of EVERYTHING and, right after it, the consolidation of what
        THAT run downloaded (never of an older file), with ONE confirmation and ONE summary. The same engines
        as the two main buttons, so what it leaves is what pressing them one after the other would."""
        if self._ocupado:
            return
        if not self.gui.validate_dates():
            return

        plan, problema = self._planear_corrida_completa()
        if problema:
            messagebox.showerror("Corrida completa", problema)
            return
        base, txs1, txs2 = plan["base"], plan["txs1"], plan["txs2"]
        orden, consultas = self._texto_orden_y_consultas(plan)
        reemplaza = ("\n⚠️ Ya existe y se REEMPLAZA."
                     if os.path.exists(os.path.join(plan["ruta_output"], plan["nombre_salida"])) else "")
        confirm = messagebox.askyesno(
            "Confirmar corrida completa",
            f"¿Iniciar la corrida completa (descarga de las {len(txs1) + len(txs2)} transacciones y, al "
            "terminar, la consolidación en Excel)?\n\n"
            f"Sociedad : {base['sociedad']}\n"
            f"{self._texto_periodo(base)}\n\n"
            f"{orden}"
            f"{consultas}"
            "La consolidación usa SOLO los archivos que esta corrida acaba de descargar. Si alguno de los 9 "
            "indispensables no queda bien, no se ejecuta y se avisa (VBRL 2 y Monivoi 2 son opcionales: si "
            "fallan, se consolida sin ellos y sólo se avisa).\n\n"
            "Mientras corre NO copiar ni pegar nada!! Puesto que el portapapeles se usa para pasarle las listas a SAP...\n\n"
            f"Se descargará en:\n{base['input_path']}\n\n"
            f"Se generará:\n  {plan['nombre_salida']}{reemplaza}\n\n"
            f"En la carpeta:\n  {plan['ruta_output']}"
        )
        if not confirm:
            return

        self._ocupado = True
        salida = {"resumen": []}
        try:
            self.gui.disable_buttons()
            self._corrida_completa_protegida(plan, salida)
            titulo, cuerpo = self.texto_final_corrida(salida)
            mostrar = {"completa": messagebox.showinfo, "avisos": messagebox.showwarning,
                       "incompleta": messagebox.showwarning}.get(salida["estado"], messagebox.showerror)
            mostrar(titulo, cuerpo)
        finally:
            self._ocupado = False
            self.gui.enable_buttons()

    def correr_corrida_completa(self):
        """Command-line entry (corrida_completa.py): the full run with NO dialogs and no confirmation → the
        `salida` dict of `_corrida_completa_protegida` (`estado` 'completa' | 'avisos' | 'incompleta' | 'error').
        Never raises."""
        salida = {"resumen": [], "estado": "error", "error": None}
        if self._ocupado:
            salida["error"] = "Ya hay un proceso en curso."
            return salida
        try:
            plan, problema = self._planear_corrida_completa()
        except Exception as e:
            print(traceback.format_exc())
            plan, problema = None, f"No pude preparar la corrida: {e}"
        if problema:
            salida["error"] = problema
            return salida
        self._ocupado = True
        try:
            self.gui.disable_buttons()
            self._corrida_completa_protegida(plan, salida)
        finally:
            self._ocupado = False
            self.gui.enable_buttons()
        return salida

    @staticmethod
    def _lineas_primera_consolidacion(resumen):
        """Control lines of the summary that the old 1st consolidation showed (rows read, #N/A left in each
        reference table): one line per item, each ending in a line break. The VBRL / FB03 blocks only appear
        when that file was part of the run."""
        texto = (
            f"  • Filas Gallo              : {resumen['filas_gallo']:,}\n"
            f"  • Filas Monivoi            : {resumen['filas_monivoi']:,}\n"
            f"  • Referencias únicas       : {resumen['refs_unicas']:,}\n"
            f"  • Docs 'MX Commercial'     : {resumen['filas_catalogo']:,}\n"
            f"  • Globales (filtro final)  : {resumen['filas_globales']:,}\n"
            f"  • Σ Importe en moneda local: {resumen['total_importe_gallo']:,.2f}\n"
        )
        if resumen.get("vbrl_incluido"):
            texto += (
                f"  • Filas VBRL               : {resumen['filas_vbrl']:,}\n"
                f"  • Referencias (facturas)   : {resumen['filas_td_referencias']:,}\n"
                f"      …en #N/A               : {resumen['td_referencias_na']:,}\n"
            )
            if resumen.get("fb03_incluido"):
                texto += (
                    f"  • Filas FB03               : {resumen['filas_fb']:,}\n"
                    f"  • Notas de crédito         : {resumen['filas_notas_ref']:,}"
                    f"  (Σ {resumen['total_notas_ref']:,.2f})\n"
                    f"      …en #N/A               : {resumen['notas_ref_na']:,}\n"
                    f"  • PEP Vouchers             : {resumen['filas_pep_ref']:,}"
                    f"  (Σ {resumen['total_pep_ref']:,.2f})\n"
                    f"      …en #N/A               : {resumen['pep_ref_na']:,}\n"
                )
        return texto

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

    def _localizar_vbrl(self, input_path, filename, ruta_monivoi):
        """The VBRL CSV (REQUIRED: 'Diferencias' needs it) → its path, or None = do not consolidate. A file
        picked by hand if it is not where it is expected; one OLDER than its Monivoi may not cover its
        invoices (extra #N/A), so that is confirmed."""
        ruta = self._localizar_archivo(input_path, filename, "VBRL")
        if not ruta:
            return None
        try:
            vbrl_anterior = os.path.getmtime(ruta) < os.path.getmtime(ruta_monivoi) - 2
        except OSError:
            vbrl_anterior = False
        if vbrl_anterior and not messagebox.askyesno(
            "VBRL anterior al archivo de Monivoi",
            f"El archivo de VBRL:\n  {os.path.basename(ruta)}\n"
            f"se descargó ANTES que el de Monivoi:\n  {os.path.basename(ruta_monivoi)}\n\n"
            "Probablemente no cubre las facturas actuales y saldrían #N/A de más.\n\n"
            "¿Usarlo de todos modos?",
        ):
            return None
        return ruta

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

    def _anteriores_segunda(self, padres, rutas_segunda):
        """2nd-run files older than the 1st-run files their lists came from (Monitoreo_2: older than its VBRL_2,
        whose lists it pasted) as 'name (reason)' lines; [] when a date cannot be read."""
        tolerancia = self._TOLERANCIA_ANTIGUEDAD
        try:
            ultimo_padre = max(os.path.getmtime(ruta) for ruta in padres)
            viejos = {ruta: "anterior a la 1a corrida" for ruta in rutas_segunda.values()
                      if ruta and os.path.getmtime(ruta) < ultimo_padre - tolerancia}
            monivoi2, vbrl2 = rutas_segunda.get("monivoi2"), rutas_segunda.get("vbrl2")
            if (monivoi2 and vbrl2 and monivoi2 not in viejos
                    and os.path.getmtime(monivoi2) < os.path.getmtime(vbrl2) - tolerancia):
                viejos[monivoi2] = f"anterior a {os.path.basename(vbrl2)}"
        except OSError:
            return []
        return [f"{os.path.basename(ruta)} ({motivo})" for ruta, motivo in viejos.items()]

    def _localizar_archivo(self, input_path, filename, etiqueta):
        """The file with the expected name, or one picked by hand if it is not there; None = cancel."""
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

    @staticmethod
    def _texto_info(info):
        """What a 2nd-run download chose / checked, for the summary line: ' · layout 'X' · ...'."""
        partes = []
        if info.get("layout"):
            partes.append(f"layout {info['layout']!r}")
        if info.get("campo"):
            partes.append(f"campo {info['campo']!r}")
        if info.get("fechas"):
            desde, hasta = info["fechas"]
            partes.append(f"creadas del {desde} al {hasta}")
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
        """KNA1 / KNVP: the 'Solic.' customers of THIS run's VBRK (see `_descargar_con_documentos`)."""
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

    def _descargar_monivoi2(self, tx, config, resultados, *, session=None, cfg_padre=None, estado_cb=None):
        """Monivoi 2nd run: the 'List.fact.' lists of THIS run's VBRL 2nd run (see `_descargar_con_documentos`;
        a VBRL 2nd run without data → 'Sin datos' file, no SAP query), period of utils.rango_monivoi_2."""
        if config["mode"] == "single":
            desde = hasta = config["fecha"]
        else:
            desde, hasta = config["date_from"], config["date_to"]
        info = {}
        extra = {} if session is None else {"session": session}

        def descargar(documentos):
            try:
                return Monivoi2_ZLMXCOM(documentos, desde, hasta, FolderPath=config["input_path"],
                                        FileName=config["filename"], sociedad=config["sociedad"],
                                        info=info, **extra)
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
        """Step of transactions queried ONCE with a list from the file their parent (tx["requiere"]) downloaded
        in THIS run: VBRL ← Monivoi 'Factura'; FB03 ← Gallo 'Nº documento'; Monivoi 2 ← VBRL_2 'List.fact.'.
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
