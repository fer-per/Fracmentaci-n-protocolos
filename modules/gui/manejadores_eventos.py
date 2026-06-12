import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import os
import time
import configuracion as _cfg
from modules.gui.estilos import *
from modules.lector_excel import cargar_excel_dataframe
import modules.dominio.servicios.servicio_folios as servicio_folios

class ManejadorEventosMixin:
    # ─── Seleccionar archivos ─────────────────────────────────────────────
    def _select_excel(self):
        ruta = filedialog.askopenfilename(
            title="Selecciona el archivo Excel con el inventario",
            filetypes=[
                ("Archivos Excel", "*.xlsx *.xls"),
                ("Todos los archivos", "*.*"),
            ],
        )
        if ruta:
            self.ruta_excel = Path(ruta)
            self.etiqueta_excel.config(text=self.ruta_excel.name, fg=TEXTO)
            self.icono_excel.config(text="✅", fg=EXITO)
            self._log("Archivo Excel seleccionado: " + self.ruta_excel.name, "ok")
            self._cargar_vista_previa()
            self._actualizar_etiqueta_listo()

    def _select_pdf(self):
        ruta = filedialog.askopenfilename(
            title="Selecciona el PDF escaneado con los documentos",
            filetypes=[
                ("Archivos PDF", "*.pdf"),
                ("Todos los archivos", "*.*"),
            ],
        )
        if ruta:
            self.ruta_pdf = Path(ruta)
            self.etiqueta_pdf.config(text=self.ruta_pdf.name, fg=TEXTO)
            self.icono_pdf.config(text="✅", fg=EXITO)
            self._log("Archivo PDF seleccionado: " + self.ruta_pdf.name, "ok")
            self._actualizar_etiqueta_listo()

    def _select_output_dir(self):
        ruta = filedialog.askdirectory(
            title="Elige la carpeta donde se guardarán los documentos",
            initialdir=str(self.directorio_salida),
        )
        if ruta:
            self.directorio_salida = Path(ruta)
            self.etiqueta_salida.config(text=str(self.directorio_salida))
            self._log("Carpeta de salida cambiada a: " + str(self.directorio_salida), "info")

    def _clear_excel(self):
        """Quita la selección del archivo Excel."""
        self.ruta_excel = None
        self.df_inventario = None
        self.etiqueta_excel.config(text="No seleccionado", fg=TEXTO_ATENUADO)
        self.icono_excel.config(text="⬜", fg=TEXTO_ATENUADO)
        # Vaciar la tabla de vista previa
        for item in self.tabla_vista_previa.get_children():
            self.tabla_vista_previa.delete(item)
        self.etiqueta_info_tabla.config(
            text="  ↑  Carga un archivo Excel para ver las filas aquí"
        )
        self._log("Selección de Excel quitada.", "warn")
        self._actualizar_etiqueta_listo()

    def _clear_pdf(self):
        """Quita la selección del archivo PDF."""
        self.ruta_pdf = None
        self.etiqueta_pdf.config(text="No seleccionado", fg=TEXTO_ATENUADO)
        self.icono_pdf.config(text="⬜", fg=TEXTO_ATENUADO)
        self._log("Selección de PDF quitada.", "warn")
        self._actualizar_etiqueta_listo()

    def _refresh_preview_range(self):
        """Re-popula la tabla de vista previa filtrando por el rango fila inicio / fila fin actual."""
        if self.df_inventario is None:
            return
        try:
            inicio = int(self.spin_inicio.get())
        except ValueError:
            inicio = self.df_inventario.index.min()
        try:
            fin = int(self.spin_fin.get())
        except ValueError:
            fin = 0
        if fin == 0 or fin > self.df_inventario.index.max():
            fin = self.df_inventario.index.max()
        self._poblar_tabla(fila_inicio=inicio, fila_fin=fin)

    # ─── Cargar vista previa del Excel ────────────────────────────────────
    def _cargar_vista_previa(self):
        try:
            self.df_inventario = cargar_excel_dataframe(self.ruta_excel, filas_omitidas=_cfg.FILAS_A_OMITIR)
        except Exception as e:
            messagebox.showerror(
                "No se pudo leer el Excel",
                f"Ocurrió un error al abrir el archivo:\n\n{e}\n\n"
                "Verifica que sea un archivo .xlsx válido.",
            )
            self._log(f"Error al leer Excel: {e}", "err")
            return

        # Configurar spinboxes con el rango real del Excel
        primer_registro = self.df_inventario.index.min()
        ultimo_registro = self.df_inventario.index.max()
        self.spin_inicio.config(from_=primer_registro, to=ultimo_registro)
        self.spin_fin.config(from_=primer_registro, to=ultimo_registro)
        
        self.spin_inicio.delete(0, "end")
        self.spin_inicio.insert(0, str(primer_registro))
        self.spin_fin.delete(0, "end")
        self.spin_fin.insert(0, str(ultimo_registro))

        # Poblar tabla con todas las filas inicialmente
        self._poblar_tabla(fila_inicio=primer_registro, fila_fin=ultimo_registro)

        total = len(self.df_inventario)
        sin_folio = sum(
            1 for _, r in self.df_inventario.iterrows()
            if not str(r.get(_cfg.COL_FOLIOS, "")).strip()
               or str(r.get(_cfg.COL_FOLIOS, "")).strip().lower() == "nan"
        )
        self.etiqueta_info_tabla.config(
            text=f"  {total} filas encontradas  ·  {sin_folio} sin folios (en amarillo)",
        )
        self._log(f"Excel cargado: {total} filas, {sin_folio} sin folios.", "info")

    def _poblar_tabla(self, fila_inicio=None, fila_fin=None):
        """Llena el Treeview mostrando solo las filas en el rango [fila_inicio, fila_fin]."""
        if self.df_inventario is None:
            return

        primera = self.df_inventario.index.min()
        ultima  = self.df_inventario.index.max()
        if fila_inicio is None or fila_inicio < primera:
            fila_inicio = primera
        if fila_fin is None or fila_fin == 0 or fila_fin > ultima:
            fila_fin = ultima

        # Limpiar tabla
        for item in self.tabla_vista_previa.get_children():
            self.tabla_vista_previa.delete(item)

        visibles = 0
        for fila_excel, row in self.df_inventario.iterrows():
            if fila_excel < fila_inicio or fila_excel > fila_fin:
                continue
            reg       = str(row.get(_cfg.COL_REGISTRO, "")).strip()
            escribano = str(row.get(_cfg.COL_ESCRIBANO, "")).strip()[:25]
            prot      = str(row.get(_cfg.COL_PROTOCOLO, "")).strip()
            folios    = str(row.get(_cfg.COL_FOLIOS, "")).strip()
            titulo    = str(row.get(_cfg.COL_TITULO_EST, "")).strip()[:25]
            int1      = str(row.get(_cfg.COL_INT1, "")).strip()[:25]

            if folios.lower() == "nan":
                folios = ""
            if reg.lower() == "nan":
                reg = ""

            tag = "sin_folio" if not folios else ""
            self.tabla_vista_previa.insert("", "end",
                                         values=(fila_excel, reg, escribano, prot, folios, titulo, int1),
                                         tags=(tag,))
            visibles += 1

        self.tabla_vista_previa.tag_configure("sin_folio", foreground=ADVERTENCIA)

        total = len(self.df_inventario)
        if visibles < total:
            self.etiqueta_info_tabla.config(
                text=(f"  Mostrando {visibles} filas "
                      f"(fila {fila_inicio} → {fila_fin})  ·  Total en Excel: {total}")
            )
        else:
            sin_folio = sum(
                1 for _, r in self.df_inventario.iterrows()
                if not str(r.get(_cfg.COL_FOLIOS, "")).strip()
                   or str(r.get(_cfg.COL_FOLIOS, "")).strip().lower() == "nan"
            )
            self.etiqueta_info_tabla.config(
                text=f"  {total} filas encontradas  ·  {sin_folio} sin folios (en amarillo)"
            )

    def _on_offset_change(self):
        """Actualiza los valores de offset y refresca la etiqueta descriptiva."""
        try:
            fi = int(self.spin_folio_inicio.get())
            pp = int(self.spin_pdf_inicio.get())
            self._folio_inicio = max(1, fi)
            self._pdf_page_inicio = max(1, pp)
        except ValueError:
            return
        portadas = self._pdf_page_inicio - 1
        texto = (
            f"  Folio {self._folio_inicio}r = pagina {self._pdf_page_inicio} del PDF"
            + (f"  ({portadas} pag. introductorias)" if portadas > 0 else "")
        )
        self.etiqueta_info_desplazamiento.config(text=texto, fg=TEXTO_MUTED)

        self._auto_refresh_coverage()

    def _add_segment(self):
        """Agrega un segmento adicional a la tabla usando notacion de folio (401r, 53v, etc.)."""
        texto_folio = self.entrada_seg_folio.get().strip()
        pag_abs = servicio_folios.texto_folio_a_pagina_abs(texto_folio)
        if pag_abs is None:
            messagebox.showwarning(
                "Formato invalido",
                f"'{texto_folio}' no es un folio valido.\n"
                "Usa el formato: numero + r o v  (ej: 401r, 53v, 30r)"
            )
            return
        try:
            pag = int(self.entrada_seg_pag.get().strip())
            if pag < 1:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Valor invalido", "Pagina PDF debe ser un numero entero positivo.")
            return
        self.tabla_segmentos.insert("", "end", values=(texto_folio, pag), tags=(str(pag_abs),))
        self._auto_refresh_coverage()

    def _del_segment(self):
        """Elimina el segmento seleccionado en la tabla."""
        seleccionado = self.tabla_segmentos.selection()
        if not seleccionado:
            messagebox.showinfo("Sin seleccion", "Selecciona un segmento en la tabla para eliminarlo.")
            return
        for item in seleccionado:
            self.tabla_segmentos.delete(item)
        self._auto_refresh_coverage()

    # ─── Páginas a ignorar ────────────────────────────────────────────────
    def _parse_ignore_input(self, texto: str):
        texto = texto.strip()
        import re as _re
        m_rango = _re.match(r'^(\d+)\s*[-–]\s*(\d+)$', texto)
        m_simple = _re.match(r'^(\d+)$', texto)
        if m_rango:
            ini = int(m_rango.group(1))
            fin = int(m_rango.group(2))
            if fin < ini:
                return None, f"Rango incoherente: {ini} > {fin}"
            if ini < 1:
                return None, "El numero de pagina debe ser >= 1"
            return set(range(ini, fin + 1)), f"{ini}-{fin}"
        elif m_simple:
            n = int(m_simple.group(1))
            if n < 1:
                return None, "El numero de pagina debe ser >= 1"
            return {n}, str(n)
        else:
            return None, f"Formato no reconocido: '{texto}'\nUsa un numero (ej: 5) o rango (ej: 10-15)"

    def _add_ignored_page(self):
        """Agrega una pagina o rango de paginas a la lista de ignorados."""
        texto = self.entrada_ignorar_pag.get().strip()
        paginas, etiqueta_o_error = self._parse_ignore_input(texto)
        if paginas is None:
            messagebox.showwarning("Entrada invalida", etiqueta_o_error)
            return
        paginas_ordenadas = sorted(paginas)
        detalle = ", ".join(str(p) for p in paginas_ordenadas[:20])
        if len(paginas_ordenadas) > 20:
            detalle += f" … ({len(paginas_ordenadas)} paginas)"
        self.tabla_ignorados.insert("", "end",
                                  values=(etiqueta_o_error, detalle),
                                  tags=("|".join(str(p) for p in paginas_ordenadas),))
        self._refresh_ignore_label()
        self._auto_refresh_coverage()

    def _del_ignored_page(self):
        """Elimina la entrada de ignorados seleccionada."""
        seleccionado = self.tabla_ignorados.selection()
        if not seleccionado:
            messagebox.showinfo("Sin seleccion",
                                "Selecciona una entrada en la tabla de ignorados para eliminarla.")
            return
        for item in seleccionado:
            self.tabla_ignorados.delete(item)
        self._refresh_ignore_label()
        self._auto_refresh_coverage()

    def _refresh_ignore_label(self):
        """Actualiza el contador de paginas ignoradas."""
        total = len(self._get_ignored_pages())
        if total == 0:
            self.etiqueta_cantidad_ignorados.config(
                text="  Sin páginas ignoradas.", fg=TEXTO_MUTED)
        else:
            self.etiqueta_cantidad_ignorados.config(
                text=f"  {total} página(s) PDF serán ignoradas en la fragmentación.",
                fg=ADVERTENCIA)

    def _get_ignored_pages(self) -> set:
        """Devuelve el set completo de paginas PDF a ignorar, uniendo todas las entradas."""
        resultado = set()
        for item in self.tabla_ignorados.get_children():
            tags = self.tabla_ignorados.item(item, "tags")
            if tags:
                try:
                    for p_str in tags[0].split("|"):
                        if p_str:
                            resultado.add(int(p_str))
                except (ValueError, IndexError):
                    pass
        return resultado

    def _auto_refresh_coverage(self):
        """Re-ejecuta el analisis de cobertura PDF si hay Excel y PDF cargados."""
        if self.df_inventario is not None and self.ruta_pdf is not None:
            self._analyze_pdf_coverage()
        self._auto_refresh_report()

    def _get_segments(self) -> list:
        """
        Devuelve la lista de segmentos como (page_abs_inicio, pdf_page_inicio),
        ordenada de menor a mayor page_abs.
        """
        primer_abs = servicio_folios._folio_a_pagina_abs(self._folio_inicio, 'r')
        segs = [(primer_abs, self._pdf_page_inicio)]
        for item in self.tabla_segmentos.get_children():
            tags = self.tabla_segmentos.item(item, "tags")
            vals = self.tabla_segmentos.item(item, "values")
            try:
                page_abs = int(tags[0]) if tags else servicio_folios.texto_folio_a_pagina_abs(str(vals[0]))
                pdf_pag  = int(vals[1])
                if page_abs and pdf_pag > 0:
                    segs.append((page_abs, pdf_pag))
            except (ValueError, IndexError, TypeError):
                pass
        return sorted(segs, key=lambda x: x[0])

    def _open_output_folder(self):
        salida = self.directorio_salida
        salida.mkdir(parents=True, exist_ok=True)
        os.startfile(str(salida))
