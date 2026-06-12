"""
modules/gui/analizadores.py - Capa de presentación: analizadores.

Los métodos de este mixin leen el estado de la GUI (spinboxes, segmentos, rutas)
y delegan la lógica de análisis a los Casos de Uso del contenedor.
No contienen ninguna regla de negocio directa.
"""
import time
from tkinter import messagebox
import configuracion as _cfg
from modules.contenedor import (
    caso_uso_analizar_folios,
    caso_uso_analizar_topica,
    caso_uso_analizar_cronica,
    caso_uso_verificar_cobertura,
)
from modules.gui.estilos import TEXTO_MUTED, EXITO_ATENUADO, ADVERTENCIA, EXITO, ERROR


class AnalizadorMixin:

    # ─── Analizador de sucesión de folios ─────────────────────────────────
    def _analyze_folios(self):
        """Analiza la sucesión de folios en el rango de filas seleccionado."""
        if not self.ruta_excel:
            messagebox.showwarning("Sin datos", "Primero carga un archivo Excel.")
            return
        try:
            fila_inicio = int(self.spin_inicio.get())
            fila_fin   = int(self.spin_fin.get())
        except ValueError:
            fila_inicio, fila_fin = 0, 0

        segmentos = self._get_segments()
        rango_txt = f"filas {fila_inicio}–{fila_fin}" if fila_fin > 0 else f"desde fila {fila_inicio}"

        try:
            resultado = caso_uso_analizar_folios.ejecutar(
                ruta_excel=str(self.ruta_excel),
                filas_omitidas=_cfg.FILAS_A_OMITIR,
                fila_inicio=fila_inicio,
                fila_fin=fila_fin,
                segmentos=segmentos,
            )
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo analizar los folios:\n{exc}")
            return

        cant_repetidos = len(resultado.get("repeated_folios", []))
        if cant_repetidos > 0:
            self._log(
                f"[FOLIOS REPETIDOS] {cant_repetidos} folio(s) repetido(s) en {rango_txt}. "
                f"Ver panel del analizador.", "warn",
            )
        self._mostrar_resultado_analizador(
            f"ANALISIS DE SUCESION DE FOLIOS  ({rango_txt})",
            resultado["resumen"], resultado["ok"]
        )

    # ─── Analizador de DATA TÓPICA ─────────────────────────────────────────
    def _analyze_data_topica(self):
        """Analiza DATA TÓPICA en el rango de filas seleccionado."""
        if not self.ruta_excel:
            messagebox.showwarning("Sin datos", "Primero carga un archivo Excel.")
            return
        try:
            fila_inicio = int(self.spin_inicio.get())
            fila_fin   = int(self.spin_fin.get())
        except ValueError:
            fila_inicio, fila_fin = 0, 0

        rango_txt = f"filas {fila_inicio}–{fila_fin}" if fila_fin > 0 else f"desde fila {fila_inicio}"

        try:
            resultado = caso_uso_analizar_topica.ejecutar(
                ruta_excel=str(self.ruta_excel),
                filas_omitidas=_cfg.FILAS_A_OMITIR,
                fila_inicio=fila_inicio,
                fila_fin=fila_fin,
            )
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo analizar DATA TÓPICA:\n{exc}")
            return

        self._mostrar_resultado_analizador(
            f"ANALISIS DATA TOPICA  ({rango_txt})", resultado["resumen"], resultado["ok"]
        )

    # ─── Analizador de DATA CRÓNICA ────────────────────────────────────────
    def _analyze_data_cronica(self):
        """Analiza DATA CRÓNICA en el rango de filas seleccionado."""
        if not self.ruta_excel:
            messagebox.showwarning("Sin datos", "Primero carga un archivo Excel.")
            return
        try:
            fila_inicio = int(self.spin_inicio.get())
            fila_fin   = int(self.spin_fin.get())
        except ValueError:
            fila_inicio, fila_fin = 0, 0

        rango_txt = f"filas {fila_inicio}–{fila_fin}" if fila_fin > 0 else f"desde fila {fila_inicio}"

        try:
            resultado = caso_uso_analizar_cronica.ejecutar(
                ruta_excel=str(self.ruta_excel),
                filas_omitidas=_cfg.FILAS_A_OMITIR,
                fila_inicio=fila_inicio,
                fila_fin=fila_fin,
            )
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudo analizar DATA CRÓNICA:\n{exc}")
            return

        self._mostrar_resultado_analizador(
            f"ANALISIS DATA CRONICA  ({rango_txt})", resultado["resumen"], resultado["ok"]
        )

    # ─── Verificador de cobertura PDF ──────────────────────────────────────
    def _analyze_pdf_coverage(self):
        """Verifica si la última hoja del PDF coincide exactamente con el
        final del rango configurado (filas manuales + segmentos adicionales)."""
        if not self.ruta_excel:
            self._mostrar_resultado_analizador(
                "VERIFICACION DE COBERTURA PDF",
                "[ERR]  No hay Excel cargado. Carga primero un archivo Excel.", False
            )
            return
        if not self.ruta_pdf:
            self._mostrar_resultado_analizador(
                "VERIFICACION DE COBERTURA PDF",
                "[ERR]  No hay PDF seleccionado. Selecciona primero el archivo PDF.", False
            )
            return

        try:
            fila_inicio = int(self.spin_inicio.get())
            fila_fin   = int(self.spin_fin.get())
        except ValueError:
            self._mostrar_resultado_analizador(
                "VERIFICACION DE COBERTURA PDF",
                "[ERR]  Rango de filas no válido (revisa los spinboxes).", False
            )
            return

        segmentos = self._get_segments()

        try:
            resultado = caso_uso_verificar_cobertura.ejecutar(
                ruta_excel=str(self.ruta_excel),
                ruta_pdf=str(self.ruta_pdf),
                filas_omitidas=_cfg.FILAS_A_OMITIR,
                fila_inicio=fila_inicio,
                fila_fin=fila_fin,
                folio_inicio_excel=self._folio_inicio,
                pag_pdf_inicio=self._pdf_page_inicio,
                segmentos=segmentos,
            )
        except Exception as exc:
            self._mostrar_resultado_analizador(
                "VERIFICACION DE COBERTURA PDF",
                f"[ERR]  Error al abrir los archivos: {exc}", False
            )
            return

        pdf_name = self.ruta_pdf.name if self.ruta_pdf else "—"
        seg_desc = f"{len(segmentos)}  {segmentos}"
        rango_desc = (
            f"Filas Excel {fila_inicio} – {fila_fin}"
            + (" (hasta el final)" if fila_fin == 0 else "")
        )
        lineas_encabezado = [
            f"  PDF seleccionado   : {pdf_name}",
            f"  Rango de filas     : {rango_desc}",
            f"  Segmentos offset   : {seg_desc}",
            "",
        ]
        res_completo = "\n".join(lineas_encabezado) + resultado["resumen"]
        self._mostrar_resultado_analizador("VERIFICACION DE COBERTURA PDF", res_completo, resultado["ok"])

    # ─── Render de resultados de análisis ──────────────────────────────────
    def _mostrar_resultado_analizador(self, titulo: str, resumen: str, ok: bool):
        """Escribe el resultado de un análisis en el panel de analizadores."""
        self.txt_analizador.config(state="normal")
        self.txt_analizador.delete("1.0", "end")
        ts = time.strftime("%H:%M:%S")
        sep = "-" * 60
        encabezado = f"[{ts}] {titulo}\n{sep}\n"
        self.txt_analizador.insert("end", encabezado, "info")

        _en_seccion_rep = False
        for linea in resumen.splitlines():
            if "FOLIOS REPETIDOS" in linea or "─── FOLIOS REPETIDOS" in linea:
                _en_seccion_rep = True
            if _en_seccion_rep and linea.startswith("[OK]"):
                _en_seccion_rep = False

            if linea.startswith("[OK]"):
                tag = "ok"
            elif linea.startswith("[ERR]"):
                tag = "err"
            elif linea.startswith("[!!"):
                tag = "rep" if ("repetida" in linea or "repetido" in linea) else "warn"
            elif linea.strip().startswith("[REP]"):
                tag = "rep"
            elif _en_seccion_rep:
                tag = "rep"
            else:
                tag = "info"
            self.txt_analizador.insert("end", linea + "\n", tag)

        self.txt_analizador.see("end")
        self.txt_analizador.config(state="disabled")
