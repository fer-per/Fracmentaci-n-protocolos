"""
modules/gui/procesador.py - Capa de presentación: hilo de procesamiento.

Delega toda la lógica de negocio al CasoUsoFragmentarPdf a través del
contenedor de dependencias. El mixin solo administra el estado de la GUI
(botones, barra de progreso, log) y el hilo de fondo.
"""
import time
import threading
import queue
from tkinter import messagebox
from pathlib import Path

import configuracion as _cfg
from modules.contenedor import caso_uso_fragmentar_pdf
from modules.gui.estilos import ACENTO, ADVERTENCIA, EXITO


class ProcesadorMixin:

    # ─── Página en blanco al final (función de utilidad pura) ─────────────
    @staticmethod
    def _agregar_pagina_blanco(ruta_pdf: Path):
        try:
            from pypdf import PdfWriter, PdfReader as _PdfReader
            lector_tmp = _PdfReader(str(ruta_pdf))
            escritor_tmp = PdfWriter()
            for pagina in lector_tmp.pages:
                escritor_tmp.add_page(pagina)
            ultima_pg = lector_tmp.pages[-1]
            w = float(ultima_pg.mediabox.width)
            h = float(ultima_pg.mediabox.height)
            escritor_tmp.add_blank_page(width=w, height=h)
            with open(str(ruta_pdf), "wb") as f_salida:
                escritor_tmp.write(f_salida)
        except Exception:
            pass

    # ─── Poll de la cola de logs (hilo principal) ─────────────────────────
    def _poll_log_queue(self):
        while not self.log_queue.empty():
            try:
                msg = self.log_queue.get_nowait()
                tag = "info"
                if "ERROR" in msg or "CRITICO" in msg:
                    tag = "err"
                elif "WARNING" in msg or "OMITIDO" in msg:
                    tag = "warn"
                elif msg.startswith("OK") or "✓" in msg:
                    tag = "ok"
                self._log(msg, tag)
            except queue.Empty:
                break
        self.after(200, self._poll_log_queue)

    # ─── Inicio del proceso ────────────────────────────────────────────────
    def _start_processing(self):
        if self.procesando:
            return

        if not self.ruta_excel:
            messagebox.showwarning(
                "Falta el Excel",
                "No has seleccionado un archivo Excel.\n\n"
                "Ve al Paso ① y haz clic en «Seleccionar…» junto a 'Archivo Excel'.",
            )
            return

        if not self.ruta_pdf:
            messagebox.showwarning(
                "Falta el PDF",
                "No has seleccionado un archivo PDF.\n\n"
                "Ve al Paso ① y haz clic en «Seleccionar…» junto a 'Archivo PDF'.",
            )
            return

        if self.df_inventario is None or len(self.df_inventario) == 0:
            messagebox.showwarning(
                "Excel vacío",
                "El archivo Excel no contiene datos para procesar.\n"
                "Verifica que sea el archivo correcto.",
            )
            return

        self.procesando = True
        self.bandera_cancelacion = False
        # Leer offsets en hilo principal antes de pasar a hilo de fondo
        self._on_offset_change()
        self.btn_procesar.config(state="disabled", text="⏳  Procesando…", bg="#312e81")
        self.btn_cancelar.config(state="normal")
        self.marco_resultados.pack_forget()
        self.progreso["value"] = 0
        self.etiqueta_progreso.config(text="Preparando…")
        self._log("— Proceso iniciado —", "ok")

        hilo = threading.Thread(target=self._ejecutar_proceso, daemon=True)
        hilo.start()

    def _cancel_processing(self):
        if self.procesando:
            self.bandera_cancelacion = True
            self._log("Cancelación solicitada. Esperando que termine la fila actual…", "warn")
            self.btn_cancelar.config(state="disabled", text="Cancelando…")

    # ─── Hilo de procesamiento ─────────────────────────────────────────────
    def _ejecutar_proceso(self):
        tiempo_inicio = time.time()
        try:
            fila_inicio       = int(self.spin_inicio.get())
            fila_fin          = int(self.spin_fin.get())
            verificar_saltos  = not self.var_ignorar_saltos.get()
            directorio_salida = str(self.directorio_salida)
            segmentos         = self._get_segments()
            paginas_ignoradas = self._get_ignored_pages()
            sobreescritura_exacta = getattr(self, "_exact_overrides", {})

            if paginas_ignoradas:
                self._log_thread(
                    f"Páginas PDF ignoradas ({len(paginas_ignoradas)}): "
                    + ", ".join(str(p) for p in sorted(paginas_ignoradas))
                )

            total_filas = len(self.df_inventario)

            # ── Callbacks para el caso de uso ──────────────────────────────
            def al_progresar(actual: int, total: int, msg: str):
                pct = int(actual / total * 100)
                self.after(0, lambda p=pct, m=msg: self._update_progress(p, m))

            def al_loguear(msg: str, nivel: str = "info"):
                self._log_thread(msg, nivel)

            def es_cancelado() -> bool:
                return self.bandera_cancelacion

            # ── Delegar en el Caso de Uso ──────────────────────────────────
            resultado = caso_uso_fragmentar_pdf.ejecutar(
                ruta_excel=self.ruta_excel,
                ruta_pdf=self.ruta_pdf,
                directorio_salida=directorio_salida,
                ruta_csv=str(_cfg.CSV_PENDIENTES),
                filas_omitidas=_cfg.FILAS_A_OMITIR,
                fila_inicio=fila_inicio,
                fila_fin=fila_fin,
                verificar_saltos=verificar_saltos,
                folio_inicio_excel=self._folio_inicio,
                pag_pdf_inicio=self._pdf_page_inicio,
                segmentos=segmentos,
                paginas_pdf_ignoradas=paginas_ignoradas,
                sobreescritura_exacta=sobreescritura_exacta,
                simulacion=False,
                callback_progreso=al_progresar,
                callback_log=al_loguear,
                verificar_cancelacion=es_cancelado,
            )

            tiempo_transcurrido = time.time() - tiempo_inicio
            procesados = resultado.get("procesados", 0)
            omitidos   = resultado.get("omitidos", 0)
            errores    = []

            cancelado = self.bandera_cancelacion
            self.after(0, lambda: self._mostrar_resultados(
                total_filas, procesados, omitidos, tiempo_transcurrido, errores, cancelado
            ))

        except FileNotFoundError as e:
            self.after(0, lambda: messagebox.showerror(
                "Archivo no encontrado",
                f"No se encontró el archivo:\n\n{e}\n\n"
                "Verifica que no se haya movido o eliminado.",
            ))
            self._log_thread(f"ERROR: Archivo no encontrado - {e}", "err")
        except Exception as e:
            self.after(0, lambda: messagebox.showerror(
                "Error inesperado",
                f"Ocurrió un error durante el proceso:\n\n{e}\n\n"
                "Revisa el registro de actividad para más detalles.",
            ))
            self._log_thread(f"ERROR CRÍTICO: {e}", "err")
        finally:
            self.after(0, self._reiniciar_boton)

    # ─── Utilidades del hilo ───────────────────────────────────────────────
    def _log_thread(self, mensaje: str, nivel: str = "info"):
        """Envía un mensaje al log desde un hilo secundario (thread-safe)."""
        tag = nivel
        if tag == "info":
            if mensaje.startswith("OK") or "✓" in mensaje:
                tag = "ok"
            elif "ERROR" in mensaje.upper():
                tag = "err"
            elif "omitida" in mensaje or "cancelad" in mensaje.lower():
                tag = "warn"
        self.after(0, lambda m=mensaje, t=tag: self._log(m, t))

    def _update_progress(self, pct: int, msg: str):
        self.progreso["value"] = pct
        self.etiqueta_progreso.config(text=msg)

    def _reiniciar_boton(self):
        self.procesando = False
        self.bandera_cancelacion = False
        self.btn_procesar.config(state="normal", text="▶  PROCESAR", bg=ACENTO)
        self.btn_cancelar.config(state="disabled", text="✕  Cancelar")

    def _mostrar_resultados(self, total, procesados, omitidos, transcurrido, errores, cancelado=False):
        self.progreso["value"] = 100 if not cancelado else self.progreso["value"]

        if cancelado:
            self.etiqueta_progreso.config(text="Cancelado por el usuario")
            self.etiqueta_titulo_resultado.config(
                text=f"Proceso cancelado  ·  {procesados} PDFs generados antes de cancelar",
                fg=ADVERTENCIA,
            )
        elif omitidos == 0:
            self.etiqueta_progreso.config(text="¡Completado!")
            self.etiqueta_titulo_resultado.config(
                text=f"✅  ¡Todos los {procesados} registros se procesaron correctamente!",
                fg=EXITO,
            )
        else:
            self.etiqueta_progreso.config(text="Completado con omisiones")
            self.etiqueta_titulo_resultado.config(
                text=f"Proceso terminado  ·  {procesados} exitosos, {omitidos} omitidos",
                fg=ADVERTENCIA,
            )

        cuerpo = (
            f"Total de filas:          {total}\n"
            f"PDFs generados:          {procesados}\n"
            f"No procesados:           {omitidos}\n"
            f"Tiempo:                  {transcurrido:.1f} segundos\n"
            f"Carpeta de salida:       {self.directorio_salida}"
        )
        if errores:
            cuerpo += "\n\nDetalles de omisiones (primeros 10):\n"
            for e in errores[:10]:
                cuerpo += f"  · {e}\n"
            if len(errores) > 10:
                cuerpo += f"  … y {len(errores)-10} más"

        self.etiqueta_cuerpo_resultado.config(text=cuerpo)
        self.marco_resultados.pack(fill="x", padx=24, pady=(8, 16))

        self._log("— Proceso finalizado —", "ok")
        self._log(f"  Generados: {procesados} | Omitidos: {omitidos} | Tiempo: {transcurrido:.1f}s")
