import tkinter as tk
from tkinter import ttk, messagebox
import time as _time
import configuracion as _cfg
from modules.gui.estilos import *
import modules.dominio.servicios.servicio_folios as servicio_folios

class ConstructorUiMixin:
    # ─── Construir la interfaz ────────────────────────────────────────────
    def _construir_ui(self):
        # Lienzo principal desplazable (scrollable)
        self.lienzo = tk.Canvas(self, bg=FONDO, highlightthickness=0)
        estilo_sb = ttk.Style()
        estilo_sb.configure("Thin.Vertical.TScrollbar", troughcolor=FONDO, background=BORDE, width=8)
        self.barra_desplazamiento = ttk.Scrollbar(self, orient="vertical", command=self.lienzo.yview,
                                                style="Thin.Vertical.TScrollbar")
        self.principal = tk.Frame(self.lienzo, bg=FONDO)

        # Actualizar región de scroll
        self.principal.bind("<Configure>", self._al_configurar_contenido)
        self._ventana_lienzo = self.lienzo.create_window((0, 0), window=self.principal, anchor="nw")
        self.lienzo.configure(yscrollcommand=self.barra_desplazamiento.set)

        self.lienzo.pack(side="left", fill="both", expand=True)
        self.barra_desplazamiento.pack(side="right", fill="y")

        # Ancho responsivo
        self.lienzo.bind("<Configure>", self._al_configurar_lienzo)

        # Rueda del mouse
        self.bind_all("<MouseWheel>",
                      lambda e: self.lienzo.yview_scroll(-1*(e.delta//120), "units"))

        # ── Encabezado Estilo Notion ──
        encabezado = tk.Frame(self.principal, bg=FONDO)
        encabezado.pack(fill="x", padx=32, pady=(32, 0))

        tk.Label(encabezado, text="🗂",
                 font=("Segoe UI Emoji", 28), bg=FONDO, fg=TEXTO).pack(anchor="w")
        tk.Label(encabezado, text="Automatización Archivística",
                 font=FUENTE_TITULO, bg=FONDO, fg=TEXTO).pack(anchor="w", pady=(4, 0))
        tk.Label(encabezado,
                 text="Fragmenta PDFs escaneados en documentos individuales según el inventario Excel.",
                 font=("Segoe UI", 12), bg=FONDO, fg=TEXTO_MUTED, anchor="w",
                 ).pack(anchor="w", pady=(4, 0))

        # Línea separadora sutil
        sep = tk.Frame(self.principal, bg=BORDE, height=1)
        sep.pack(fill="x", padx=32, pady=(20, 14))

        # ══════════════════════════════════════════════════════════════════
        # PASO 1: Seleccionar archivos
        # ══════════════════════════════════════════════════════════════════
        self._encabezado_paso("①", "Selecciona tus archivos",
                             "Elige el Excel con el inventario y el PDF escaneado.")

        tarjeta_archivos = self._crear_tarjeta()

        # ── Excel ──
        fila_e = tk.Frame(tarjeta_archivos, bg=FONDO_TARJETA)
        fila_e.pack(fill="x", padx=20, pady=(16, 6))

        self.icono_excel = tk.Label(fila_e, text="📄", font=("Segoe UI Emoji", 16),
                                   bg=FONDO_TARJETA, fg="#2a7a3e", width=2)
        self.icono_excel.pack(side="left")

        marco_lbl_e = tk.Frame(fila_e, bg=FONDO_TARJETA)
        marco_lbl_e.pack(side="left", fill="x", expand=True, padx=(6, 8))
        tk.Label(marco_lbl_e, text="Archivo Excel (.xlsx)",
                 font=FUENTE_NEGRITA, bg=FONDO_TARJETA, fg=TEXTO, anchor="w").pack(anchor="w")
        self.etiqueta_excel = tk.Label(marco_lbl_e, text="No seleccionado",
                                      font=FUENTE_PEQUENA, bg=FONDO_TARJETA, fg=TEXTO_ATENUADO, anchor="w")
        self.etiqueta_excel.pack(anchor="w")

        btn_e = self._crear_boton(fila_e, "Seleccionar…", self._select_excel, width=14)
        btn_e.pack(side="right", padx=(4, 0))
        EtiquetaInformacion(btn_e, "Haz clic para buscar tu archivo Excel\ncon el inventario de documentos.")

        self.btn_quitar_excel = self._crear_boton(
            fila_e, "✕ Quitar", self._clear_excel, width=7, tipo="secundario")
        self.btn_quitar_excel.pack(side="right", padx=(4, 0))
        EtiquetaInformacion(self.btn_quitar_excel, "Quita la selección actual del archivo Excel.")

        # ── PDF ──
        fila_p = tk.Frame(tarjeta_archivos, bg=FONDO_TARJETA)
        fila_p.pack(fill="x", padx=20, pady=(6, 16))

        self.icono_pdf = tk.Label(fila_p, text="📕", font=("Segoe UI Emoji", 16),
                                  bg=FONDO_TARJETA, fg="#c84e1b", width=2)
        self.icono_pdf.pack(side="left")

        marco_lbl_p = tk.Frame(fila_p, bg=FONDO_TARJETA)
        marco_lbl_p.pack(side="left", fill="x", expand=True, padx=(6, 8))
        tk.Label(marco_lbl_p, text="Archivo PDF protocolo escaneado",
                 font=FUENTE_NEGRITA, bg=FONDO_TARJETA, fg=TEXTO, anchor="w").pack(anchor="w")
        self.etiqueta_pdf = tk.Label(marco_lbl_p, text="No seleccionado",
                                    font=FUENTE_PEQUENA, bg=FONDO_TARJETA, fg=TEXTO_ATENUADO, anchor="w")
        self.etiqueta_pdf.pack(anchor="w")

        btn_p = self._crear_boton(fila_p, "Seleccionar…", self._select_pdf, width=14)
        btn_p.pack(side="right", padx=(4, 0))
        EtiquetaInformacion(btn_p, "Haz clic para buscar el PDF escaneado\nque contiene los documentos originales.")

        self.btn_quitar_pdf = self._crear_boton(
            fila_p, "✕ Quitar", self._clear_pdf, width=7, tipo="secundario")
        self.btn_quitar_pdf.pack(side="right", padx=(4, 0))
        EtiquetaInformacion(self.btn_quitar_pdf, "Quita la selección actual del archivo PDF.")

        # ══════════════════════════════════════════════════════════════════
        # PASO 2: Vista previa y rango
        # ══════════════════════════════════════════════════════════════════
        self._encabezado_paso("②", "Revisa los datos y elige el rango",
                             "La tabla muestra las filas del Excel. Puedes procesar todas o solo un rango.")

        tarjeta_vista_previa = self._crear_tarjeta()

        # Tabla de vista previa
        contenedor_tabla = tk.Frame(tarjeta_vista_previa, bg=FONDO_TARJETA)
        contenedor_tabla.pack(fill="x", padx=16, pady=(14, 8))

        cols = ("fila", "reg", "escribano", "prot", "folios", "titulo", "interesado1")
        col_names = {
            "fila": "Fila", "reg": "Reg.", "escribano": "Escribano",
            "prot": "Prot.", "folios": "Folios", "titulo": "Título",
            "interesado1": "Interesado 1",
        }

        estilo = ttk.Style()
        estilo.theme_use("clam")
        estilo.configure("Dark.Treeview",
                         background="#ffffff", foreground=TEXTO,
                         fieldbackground="#ffffff",
                         font=FUENTE_PEQUENA, rowheight=28, borderwidth=0)
        estilo.configure("Dark.Treeview.Heading",
                         background="#f7f7f5", foreground=TEXTO_MUTED,
                         font=("Segoe UI", 9, "bold"),
                         borderwidth=0, relief="flat")
        estilo.map("Dark.Treeview",
                   background=[("selected", "#e2f6ff")],
                   foreground=[("selected", TEXTO)])

        self.tabla_vista_previa = ttk.Treeview(
            contenedor_tabla, columns=cols, show="headings",
            height=7, style="Dark.Treeview",
        )
        for col_id in cols:
            w = 50 if col_id in ("fila", "reg", "prot") else 85 if col_id == "folios" else 150
            self.tabla_vista_previa.heading(col_id, text=col_names[col_id])
            anchor = "center" if col_id in ("fila", "reg", "prot", "folios") else "w"
            self.tabla_vista_previa.column(col_id, width=w, minwidth=40, anchor=anchor)

        scroll_tabla = ttk.Scrollbar(contenedor_tabla, orient="vertical",
                                    command=self.tabla_vista_previa.yview)
        self.tabla_vista_previa.configure(yscrollcommand=scroll_tabla.set)
        self.tabla_vista_previa.pack(side="left", fill="x", expand=True)
        scroll_tabla.pack(side="right", fill="y")

        # Info de la tabla
        self.etiqueta_info_tabla = tk.Label(
            tarjeta_vista_previa,
            text="  Carga un archivo Excel para ver las filas aquí →",
            font=("Segoe UI", 9), bg=FONDO_TARJETA, fg=TEXTO_MUTED, anchor="w",
        )
        self.etiqueta_info_tabla.pack(fill="x", padx=16, pady=(0, 8))

        # ── Rango de filas ──
        marco_rango = tk.Frame(tarjeta_vista_previa, bg=FONDO_TARJETA)
        marco_rango.pack(padx=16, pady=(4, 14))

        tk.Label(marco_rango, text="Inicia desde fila:", font=FUENTE_NEGRITA,
                 bg=FONDO_TARJETA, fg=TEXTO).pack(side="left", padx=(0, 5))
        self.spin_inicio = tk.Spinbox(
            marco_rango, from_=1, to=99999, width=7, font=FUENTE,
            bg=FONDO_ENTRADA, fg=TEXTO, buttonbackground=BORDE,
            insertbackground=TEXTO, relief="flat", justify="center",
            command=self._refresh_preview_range,
        )
        self.spin_inicio.pack(side="left", padx=(0, 20))
        self.spin_inicio.bind("<FocusOut>", lambda e: self._refresh_preview_range())
        self.spin_inicio.bind("<Return>", lambda e: self._refresh_preview_range())
        EtiquetaInformacion(self.spin_inicio, "Primera fila del Excel que quieres procesar.\nNormalmente déjalo en 1.")

        tk.Label(marco_rango, text="Termina en fila:", font=FUENTE_NEGRITA,
                 bg=FONDO_TARJETA, fg=TEXTO).pack(side="left", padx=(0, 5))
        self.spin_fin = tk.Spinbox(
            marco_rango, from_=0, to=99999, width=7, font=FUENTE,
            bg=FONDO_ENTRADA, fg=TEXTO, buttonbackground=BORDE,
            insertbackground=TEXTO, relief="flat", justify="center",
            command=self._refresh_preview_range,
        )
        self.spin_fin.delete(0, "end")
        self.spin_fin.insert(0, "0")
        self.spin_fin.pack(side="left", padx=(0, 10))
        self.spin_fin.bind("<FocusOut>", lambda e: self._refresh_preview_range())
        self.spin_fin.bind("<Return>", lambda e: self._refresh_preview_range())
        EtiquetaInformacion(self.spin_fin, "Última fila que quieres procesar.\nDeja en 0 para procesar todas.")

        # ══════════════════════════════════════════════════════════════════
        # PASO 2b: Configuración de inicio de folios / PDF
        # ══════════════════════════════════════════════════════════════════
        self._encabezado_paso("②b", "Configuración de inicio",
                             "Indica desde qué folio empieza el protocolo y desde qué página del PDF.")

        tarjeta_desplazamiento = self._crear_tarjeta()
        fila_desplazamiento = tk.Frame(tarjeta_desplazamiento, bg=FONDO_TARJETA)
        fila_desplazamiento.pack(padx=16, pady=(14, 10))

        tk.Label(fila_desplazamiento, text="Folio inicio del Excel:",
                 font=FUENTE_NEGRITA, bg=FONDO_TARJETA, fg=TEXTO).pack(side="left", padx=(0, 5))
        self.spin_folio_inicio = tk.Spinbox(
            fila_desplazamiento, from_=1, to=99999, width=7, font=FUENTE,
            bg=FONDO_ENTRADA, fg=TEXTO, buttonbackground=BORDE,
            insertbackground=TEXTO, relief="flat", justify="center",
            command=self._on_offset_change,
        )
        self.spin_folio_inicio.delete(0, "end")
        self.spin_folio_inicio.insert(0, "1")
        self.spin_folio_inicio.pack(side="left", padx=(0, 20))
        self.spin_folio_inicio.bind("<FocusOut>", lambda e: self._on_offset_change())
        self.spin_folio_inicio.bind("<Return>",   lambda e: self._on_offset_change())
        EtiquetaInformacion(self.spin_folio_inicio,
                            "Número de folio con el que EMPIEZA el protocolo en el Excel.\n"
                            "Ej.: si el Excel inicia en '30r', escribe 30.\n"
                            "Si empieza en '1r', déjalo en 1 (valor por defecto).")

        tk.Label(fila_desplazamiento, text="Página PDF de inicio:",
                 font=FUENTE_NEGRITA, bg=FONDO_TARJETA, fg=TEXTO).pack(side="left", padx=(0, 5))
        self.spin_pdf_inicio = tk.Spinbox(
            fila_desplazamiento, from_=1, to=99999, width=7, font=FUENTE,
            bg=FONDO_ENTRADA, fg=TEXTO, buttonbackground=BORDE,
            insertbackground=TEXTO, relief="flat", justify="center",
            command=self._on_offset_change,
        )
        self.spin_pdf_inicio.delete(0, "end")
        self.spin_pdf_inicio.insert(0, "1")
        self.spin_pdf_inicio.pack(side="left", padx=(0, 10))
        self.spin_pdf_inicio.bind("<FocusOut>", lambda e: self._on_offset_change())
        self.spin_pdf_inicio.bind("<Return>",   lambda e: self._on_offset_change())
        EtiquetaInformacion(self.spin_pdf_inicio,
                            "Número de página REAL del PDF donde empieza la primera imagen del protocolo.\n"
                            "Ej.: si hay 2 portadas antes, escribe 3.\n"
                            "Si no hay portadas, déjalo en 1 (valor por defecto).")

        self.etiqueta_info_desplazamiento = tk.Label(
            tarjeta_desplazamiento,
            text="  -> La primera imagen del PDF corresponde al folio 1r del protocolo.",
            font=FUENTE_PEQUENA, bg=FONDO_TARJETA, fg=TEXTO_MUTED, anchor="w",
        )
        self.etiqueta_info_desplazamiento.pack(fill="x", padx=16, pady=(0, 6))

        # ── Segmentos adicionales (para saltos en el PDF) ─────────────────
        lbl_seg = tk.Label(
            tarjeta_desplazamiento,
            text="  Saltos de folios (cuando el PDF no contiene los folios del salto):",
            font=FUENTE_PEQUENA, bg=FONDO_TARJETA, fg=TEXTO_ATENUADO, anchor="w",
        )
        lbl_seg.pack(fill="x", padx=16, pady=(2, 4))

        marco_seg = tk.Frame(tarjeta_desplazamiento, bg=FONDO_TARJETA)
        marco_seg.pack(fill="x", padx=16, pady=(0, 4))

        # Tabla de segmentos
        self.tabla_segmentos = ttk.Treeview(
            marco_seg, columns=("folio", "pdf_pag"), show="headings",
            height=3, style="Dark.Treeview",
        )
        self.tabla_segmentos.heading("folio",   text="Folio inicio Excel")
        self.tabla_segmentos.heading("pdf_pag", text="Pagina PDF inicio")
        self.tabla_segmentos.column("folio",   width=160, anchor="center")
        self.tabla_segmentos.column("pdf_pag", width=160, anchor="center")
        self.tabla_segmentos.pack(side="left", fill="x", expand=True)

        scroll_seg = ttk.Scrollbar(marco_seg, orient="vertical",
                                   command=self.tabla_segmentos.yview)
        self.tabla_segmentos.configure(yscrollcommand=scroll_seg.set)
        scroll_seg.pack(side="right", fill="y")

        # Entradas para nuevo segmento
        fila_ent = tk.Frame(tarjeta_desplazamiento, bg=FONDO_TARJETA)
        fila_ent.pack(fill="x", padx=16, pady=(2, 8))

        tk.Label(fila_ent, text="Folio (ej: 401r, 53v):", font=FUENTE_PEQUENA,
                 bg=FONDO_TARJETA, fg=TEXTO).pack(side="left", padx=(0, 3))
        self.entrada_seg_folio = tk.Entry(
            fila_ent, width=9, font=FUENTE, bg=FONDO_ENTRADA, fg=TEXTO,
            insertbackground=TEXTO, relief="flat", justify="center",
            highlightthickness=1, highlightbackground=BORDE, highlightcolor=ACENTO,
        )
        self.entrada_seg_folio.insert(0, "401r")
        self.entrada_seg_folio.pack(side="left", padx=(0, 10))

        tk.Label(fila_ent, text="Pagina PDF:", font=FUENTE_PEQUENA,
                 bg=FONDO_TARJETA, fg=TEXTO).pack(side="left", padx=(0, 3))
        self.entrada_seg_pag = tk.Entry(
            fila_ent, width=7, font=FUENTE, bg=FONDO_ENTRADA, fg=TEXTO,
            insertbackground=TEXTO, relief="flat", justify="center",
            highlightthickness=1, highlightbackground=BORDE, highlightcolor=ACENTO,
        )
        self.entrada_seg_pag.insert(0, "635")
        self.entrada_seg_pag.pack(side="left", padx=(0, 10))

        btn_add_seg = self._crear_boton(fila_ent, "+ Agregar", self._add_segment, width=10, tipo="success")
        btn_add_seg.pack(side="left", padx=(0, 6))
        EtiquetaInformacion(btn_add_seg,
                            "Agrega un segmento: indica desde que folio del Excel\n"
                            "empieza la siguiente seccion en el PDF.\n"
                            "Util cuando el PDF no contiene los folios del salto.")

        btn_del_seg = self._crear_boton(fila_ent, "− Eliminar", self._del_segment, width=10, tipo="danger")
        btn_del_seg.pack(side="left")
        EtiquetaInformacion(btn_del_seg, "Elimina el segmento seleccionado en la tabla.")

        # ── Páginas a ignorar ────────────────────────────────────────────
        tk.Frame(tarjeta_desplazamiento, bg=BORDE, height=1).pack(fill="x", padx=16, pady=(6, 0))

        lbl_ign = tk.Label(
            tarjeta_desplazamiento,
            text="  Páginas PDF a ignorar (hojas que NO deben incluirse en ningún documento):",
            font=FUENTE_PEQUENA, bg=FONDO_TARJETA, fg=TEXTO_ATENUADO, anchor="w",
        )
        lbl_ign.pack(fill="x", padx=16, pady=(6, 4))

        marco_ign = tk.Frame(tarjeta_desplazamiento, bg=FONDO_TARJETA)
        marco_ign.pack(fill="x", padx=16, pady=(0, 4))

        self.tabla_ignorados = ttk.Treeview(
            marco_ign, columns=("entrada", "paginas"), show="headings",
            height=3, style="Dark.Treeview",
        )
        self.tabla_ignorados.heading("entrada", text="Entrada")
        self.tabla_ignorados.heading("paginas", text="Páginas PDF ignoradas")
        self.tabla_ignorados.column("entrada", width=130, anchor="center")
        self.tabla_ignorados.column("paginas", width=340, anchor="w")
        self.tabla_ignorados.pack(side="left", fill="x", expand=True)

        scroll_ign = ttk.Scrollbar(marco_ign, orient="vertical",
                                   command=self.tabla_ignorados.yview)
        self.tabla_ignorados.configure(yscrollcommand=scroll_ign.set)
        scroll_ign.pack(side="right", fill="y")

        fila_ent_ign = tk.Frame(tarjeta_desplazamiento, bg=FONDO_TARJETA)
        fila_ent_ign.pack(fill="x", padx=16, pady=(2, 10))

        tk.Label(fila_ent_ign, text="Pág. o rango (ej: 5  ó  10-15):",
                 font=FUENTE_PEQUENA, bg=FONDO_TARJETA, fg=TEXTO).pack(side="left", padx=(0, 4))
        self.entrada_ignorar_pag = tk.Entry(
            fila_ent_ign, width=10, font=FUENTE, bg=FONDO_ENTRADA, fg=TEXTO,
            insertbackground=TEXTO, relief="flat", justify="center",
            highlightthickness=1, highlightbackground=BORDE, highlightcolor=ACENTO,
        )
        self.entrada_ignorar_pag.insert(0, "5")
        self.entrada_ignorar_pag.pack(side="left", padx=(0, 10))
        EtiquetaInformacion(self.entrada_ignorar_pag,
                            "Escribe una pagina simple (ej: 5)\n"
                            "o un rango (ej: 10-15) para ignorar varias a la vez.")

        btn_add_ign = self._crear_boton(
            fila_ent_ign, "+ Ignorar", self._add_ignored_page, width=10, tipo="warning")
        btn_add_ign.pack(side="left", padx=(0, 6))
        EtiquetaInformacion(btn_add_ign,
                            "Agrega la pagina o rango indicado a la lista de exclusion.\n"
                            "Esas paginas seran omitidas en todos los documentos generados.")

        btn_del_ign = self._crear_boton(
            fila_ent_ign, "− Quitar", self._del_ignored_page, width=10, tipo="secundario")
        btn_del_ign.pack(side="left")
        EtiquetaInformacion(btn_del_ign, "Quita la entrada seleccionada de la lista de ignorados.")

        self.etiqueta_cantidad_ignorados = tk.Label(
            tarjeta_desplazamiento,
            text="  Sin páginas ignoradas.",
            font=FUENTE_PEQUENA, bg=FONDO_TARJETA, fg=TEXTO_MUTED, anchor="w",
        )
        self.etiqueta_cantidad_ignorados.pack(fill="x", padx=16, pady=(0, 8))


        # PASO 2c: Analizadores
        # ══════════════════════════════════════════════════════════════════
        self._encabezado_paso("②c", "Analizadores",
                             "Verifica la sucesión de folios y la DATA TÓPICA antes de fragmentar.")

        tarjeta_analizador = self._crear_tarjeta()

        # ── Primera fila de botones ──
        fila_btn_an1 = tk.Frame(tarjeta_analizador, bg=FONDO_TARJETA)
        fila_btn_an1.pack(fill="x", padx=16, pady=(14, 4))

        btn_analizar_fol = self._crear_boton(
            fila_btn_an1, "Analizar Folios", self._analyze_folios)
        btn_analizar_fol.pack(side="left", padx=(0, 6))
        EtiquetaInformacion(btn_analizar_fol,
                            "Verifica que la sucesion de folios del Excel sea continua.\n"
                            "Muestra que hojas o caras faltan si hay saltos.")

        btn_analizar_top = self._crear_boton(
            fila_btn_an1, "Analizar TÓPICA", self._analyze_data_topica)
        btn_analizar_top.pack(side="left", padx=(0, 6))
        EtiquetaInformacion(btn_analizar_top,
                            "Verifica que la columna DATA TÓPICA tenga formato correcto.")

        btn_analizar_cro = self._crear_boton(
            fila_btn_an1, "Analizar CRÓNICA", self._analyze_data_cronica)
        btn_analizar_cro.pack(side="left", padx=(0, 6))
        EtiquetaInformacion(btn_analizar_cro,
                            "Verifica el formato de fechas (d/m/yyyy) y que las fechas\n"
                            "vayan en orden cronologico progresivo.")

        # ── Segunda fila de botones ──
        fila_btn_an2 = tk.Frame(tarjeta_analizador, bg=FONDO_TARJETA)
        fila_btn_an2.pack(fill="x", padx=16, pady=(0, 6))

        btn_verificar_cob = self._crear_boton(
            fila_btn_an2, "Verificar Cobertura PDF", self._analyze_pdf_coverage, tipo="success")
        btn_verificar_cob.pack(side="left", padx=(0, 6))
        EtiquetaInformacion(btn_verificar_cob,
                            "Compara la ultima pagina PDF que usaria el rango configurado\n"
                            "con el total real de paginas del PDF.\n"
                            "Indica si el PDF esta bien alineado o si sobran/faltan hojas.\n"
                            "Se actualiza automaticamente al cambiar segmentos o ignorados.")

        btn_reporte = self._crear_boton(
            fila_btn_an2, "📋 Reporte", self._generar_reporte_fragmentacion, tipo="info")
        btn_reporte.pack(side="left")
        EtiquetaInformacion(btn_reporte,
                            "Genera un reporte detallado del mapeo folio \u2192 páginas PDF.\n"
                            "Muestra columnas del Excel y resalta registros afectados\n"
                            "por segmentos, p\u00e1ginas ignoradas o folios ignorados.")

        # Panel de resultados de análisis (texto expandible)
        self.txt_analizador = tk.Text(
            tarjeta_analizador, height=7, bg="#f7f7f5", fg=TEXTO_ATENUADO,
            font=FUENTE_LOG, relief="flat", wrap="word",
            insertbackground=TEXTO, state="disabled",
            padx=12, pady=10,
        )
        scroll_an = ttk.Scrollbar(tarjeta_analizador, orient="vertical",
                                  command=self.txt_analizador.yview)
        self.txt_analizador.configure(yscrollcommand=scroll_an.set)
        self.txt_analizador.pack(side="left", fill="both", expand=True, padx=(16, 0), pady=(6, 14))
        scroll_an.pack(side="right", fill="y", padx=(0, 16), pady=(6, 14))

        self.txt_analizador.tag_configure("ok",   foreground=EXITO)
        self.txt_analizador.tag_configure("warn", foreground="#c2660a")
        self.txt_analizador.tag_configure("err",  foreground=ERROR)
        self.txt_analizador.tag_configure("info", foreground=TEXTO_MUTED)
        self.txt_analizador.tag_configure("rep",  foreground="#df5b24")

        # ══════════════════════════════════════════════════════════════════
        # PASO 3: Carpeta de salida
        # ══════════════════════════════════════════════════════════════════
        self._encabezado_paso("③", "Carpeta de salida",
                             "Elige dónde se guardarán los PDFs generados.")

        tarjeta_salida = self._crear_tarjeta()
        fila_salida = tk.Frame(tarjeta_salida, bg=FONDO_TARJETA)
        fila_salida.pack(fill="x", padx=16, pady=14)

        tk.Label(fila_salida, text="📁", font=("Segoe UI Emoji", 16),
                 bg=FONDO_TARJETA, fg=TEXTO_ATENUADO).pack(side="left")

        self.etiqueta_salida = tk.Label(
            fila_salida, text=str(self.directorio_salida),
            font=FUENTE, bg="#f7f7f5", fg=TEXTO_ATENUADO,
            relief="flat", padx=10, pady=6, anchor="w",
        )
        self.etiqueta_salida.pack(side="left", fill="x", expand=True, padx=(8, 8))

        btn_out = self._crear_boton(fila_salida, "Cambiar…", self._select_output_dir, width=12)
        btn_out.pack(side="right")
        EtiquetaInformacion(btn_out, "Cambia la carpeta donde se guardan\nlos documentos separados.")

        # ══════════════════════════════════════════════════════════════════
        # PASO 4: Opciones y Procesar
        # ══════════════════════════════════════════════════════════════════
        self._encabezado_paso("④", "¡Procesar!",
                             "Revisa que todo esté listo y presiona el botón.")

        tarjeta_accion = self._crear_tarjeta()

        # Opciones
        fila_opts = tk.Frame(tarjeta_accion, bg=FONDO_TARJETA)
        fila_opts.pack(padx=20, pady=(18, 8))

        self.var_ignorar_saltos = tk.BooleanVar(value=True)
        cb = tk.Checkbutton(
            fila_opts, text="  Ignorar saltos de secuencia entre folios",
            variable=self.var_ignorar_saltos, font=FUENTE, bg=FONDO_TARJETA, fg=TEXTO_ATENUADO,
            selectcolor="#e2f6ff", activebackground=FONDO_TARJETA, activeforeground=TEXTO,
            cursor="hand2", relief="flat", bd=0,
        )
        cb.pack(side="left")
        EtiquetaInformacion(cb, "Si está marcado, el sistema NO detendrá un registro\n"
                              "solo porque hay un salto numérico entre folios.\n"
                              "Recomendado: dejarlo marcado.")

        # Estado de "listo para procesar"
        self.etiqueta_listo = tk.Label(
            tarjeta_accion, text="", font=("Segoe UI", 10), bg=FONDO_TARJETA, fg=TEXTO_ATENUADO,
        )
        self.etiqueta_listo.pack(pady=(0, 4))

        # Botones: Procesar y Cancelar
        fila_btn = tk.Frame(tarjeta_accion, bg=FONDO_TARJETA)
        fila_btn.pack(pady=(2, 6))

        self.btn_procesar = tk.Button(
            fila_btn, text="▶  PROCESAR", font=FUENTE_BOTON_G,
            bg=ACENTO, fg="#ffffff",
            activebackground=ACENTO_SELECCION, activeforeground="#ffffff",
            relief="flat", cursor="hand2", padx=40, pady=12,
            command=self._start_processing,
        )
        self.btn_procesar.pack(side="left", padx=(0, 12))
        self.btn_procesar.bind("<Enter>", lambda e: self.btn_procesar.configure(bg=ACENTO_SELECCION))
        self.btn_procesar.bind("<Leave>", lambda e: self.btn_procesar.configure(bg=ACENTO))
        EtiquetaInformacion(self.btn_procesar, "Inicia el proceso de extracción de documentos.\n"
                                             "Asegúrate de haber seleccionado ambos archivos.")

        self.btn_cancelar = self._crear_boton(
            fila_btn, "✕  Cancelar", self._cancel_processing, tipo="danger")
        self.btn_cancelar.configure(pady=12, padx=20, state="disabled")
        self.btn_cancelar.pack(side="left")

        self._actualizar_etiqueta_listo()

        # Barra de progreso
        estilo.configure("Accent.Horizontal.TProgressbar",
                         troughcolor=FONDO_ENTRADA, background=ACENTO, thickness=10)
        self.progreso = ttk.Progressbar(
            tarjeta_accion, mode="determinate", length=500,
            style="Accent.Horizontal.TProgressbar",
        )
        self.progreso.pack(pady=(6, 4))

        self.etiqueta_progreso = tk.Label(
            tarjeta_accion, text="", font=("Segoe UI", 10), bg=FONDO_TARJETA, fg=TEXTO_MUTED,
        )
        self.etiqueta_progreso.pack(pady=(0, 14))

        # ─── Registro de actividad (log en tiempo real) ───────────────────
        cabecera_log = tk.Frame(self.principal, bg=FONDO)
        cabecera_log.pack(fill="x", padx=32, pady=(16, 4))
        tk.Label(cabecera_log, text="📋  Registro de actividad",
                 font=("Segoe UI", 11, "bold"), bg=FONDO, fg=TEXTO_ATENUADO, anchor="w").pack(side="left")

        tarjeta_log = self._crear_tarjeta()
        self.txt_bitacora = tk.Text(
            tarjeta_log, height=8, bg="#f7f7f5", fg=TEXTO_ATENUADO,
            font=FUENTE_LOG, relief="flat", wrap="word",
            insertbackground=TEXTO, state="disabled",
            padx=12, pady=10,
        )
        scroll_log = ttk.Scrollbar(tarjeta_log, orient="vertical",
                                   command=self.txt_bitacora.yview)
        self.txt_bitacora.configure(yscrollcommand=scroll_log.set)
        self.txt_bitacora.pack(side="left", fill="both", expand=True, padx=(16, 0), pady=14)
        scroll_log.pack(side="right", fill="y", padx=(0, 16), pady=14)

        # Tag para colores en el log
        self.txt_bitacora.tag_configure("ok",   foreground=EXITO)
        self.txt_bitacora.tag_configure("warn", foreground="#c2660a")
        self.txt_bitacora.tag_configure("err",  foreground=ERROR)
        self.txt_bitacora.tag_configure("info", foreground=TEXTO_MUTED)

        # ══════════════════════════════════════════════════════════════════
        # Panel de resultados (oculto inicialmente)
        # ══════════════════════════════════════════════════════════════════
        self.marco_resultados = tk.Frame(self.principal, bg="#f0fdf4")

        self.etiqueta_titulo_resultado = tk.Label(
            self.marco_resultados, text="",
            font=("Segoe UI", 15, "bold"), bg="#f0fdf4", fg=EXITO,
        )
        self.etiqueta_titulo_resultado.pack(pady=(16, 6))

        self.etiqueta_cuerpo_resultado = tk.Label(
            self.marco_resultados, text="",
            font=FUENTE, bg="#f0fdf4", fg=TEXTO, justify="left",
        )
        self.etiqueta_cuerpo_resultado.pack(padx=16, pady=(0, 6))

        self.btn_abrir_salida = self._crear_boton(
            self.marco_resultados, "📂  Abrir carpeta de salida",
            self._open_output_folder, tipo="info")
        self.btn_abrir_salida.configure(padx=20, pady=8)
        self.btn_abrir_salida.pack(pady=(4, 16))
        EtiquetaInformacion(self.btn_abrir_salida, "Abre en el Explorador de Windows la carpeta\n"
                                                 "donde se guardaron los PDFs.")

        # Espaciado final
        tk.Frame(self.principal, bg=FONDO, height=32).pack()

    # ─── Scroll & resize stability ─────────────────────────────────────────
    def _al_configurar_contenido(self, evento):
        new_height = self.principal.winfo_reqheight()
        try:
            old_region = self.lienzo.cget("scrollregion")
            if old_region:
                _, _, _, old_height = map(int, old_region.split())
                if abs(new_height - old_height) < 4:
                    return
        except Exception:
            pass
        self.lienzo.configure(scrollregion=(0, 0, evento.width, new_height))

    def _al_configurar_lienzo(self, evento):
        if hasattr(self, "_resize_after_id") and self._resize_after_id:
            self.after_cancel(self._resize_after_id)
        self._resize_after_id = self.after(30, lambda w=evento.width: self._aplicar_ancho_lienzo(w))

    def _aplicar_ancho_lienzo(self, ancho):
        self._resize_after_id = None
        self.lienzo.itemconfig(self._ventana_lienzo, width=ancho)

    def _encabezado_paso(self, numero, titulo, subtitulo):
        frame = tk.Frame(self.principal, bg=FONDO)
        frame.pack(fill="x", padx=32, pady=(24, 8))

        circle_lbl = tk.Label(
            frame, text=numero, font=FUENTE_PASO,
            bg=ACENTO, fg="#ffffff", width=3, height=1,
        )
        circle_lbl.pack(side="left", padx=(0, 10))

        text_frame = tk.Frame(frame, bg=FONDO)
        text_frame.pack(side="left", fill="both", expand=True)

        tk.Label(text_frame, text=titulo, font=FUENTE_NEGRITA,
                 bg=FONDO, fg=TEXTO, anchor="w").pack(anchor="w")
        tk.Label(text_frame, text=subtitulo, font=FUENTE_PEQUENA,
                 bg=FONDO, fg=TEXTO_MUTED, anchor="w").pack(anchor="w")

    def _crear_tarjeta(self):
        tarjeta = tk.Frame(self.principal, bg=FONDO_TARJETA, highlightthickness=1,
                           highlightbackground=BORDE, highlightcolor=BORDE)
        tarjeta.pack(fill="x", padx=32, pady=6)
        return tarjeta

    def _crear_boton(self, padre, texto, comando, width=None, tipo="primary"):
        paleta = {
            "primary":   BOTON_PRINCIPAL,
            "success":   BOTON_EXITO,
            "warning":   BOTON_ADVERTENCIA,
            "danger":    BOTON_PELIGRO,
            "secundario": BOTON_SECUNDARIO,
            "info":      BOTON_INFO,
            "purple":    BOTON_PURPURA,
        }.get(tipo, BOTON_PRINCIPAL)

        btn = tk.Button(
            padre, text=texto, font=FUENTE_NEGRITA, command=comando,
            bg=paleta["bg"], fg=paleta["fg"],
            activebackground=paleta["hover"], activeforeground=paleta["fg"],
            relief="flat", cursor="hand2", borderwidth=0,
            highlightthickness=0,
        )
        if width:
            btn.configure(width=width)
        btn.configure(pady=6, padx=14)

        def _on_enter(e):
            btn.configure(bg=paleta["hover"])
        def _on_leave(e):
            btn.configure(bg=paleta["bg"])

        btn.bind("<Enter>", _on_enter)
        btn.bind("<Leave>", _on_leave)
        return btn

    def _actualizar_etiqueta_listo(self):
        ready = True
        missing = []
        if not self.ruta_excel:
            ready = False
            missing.append("Archivo Excel")
        if not self.ruta_pdf:
            ready = False
            missing.append("Archivo PDF")

        if ready:
            self.etiqueta_listo.config(text="✓ Todo listo para iniciar la fragmentación", fg=EXITO)
            if not self.procesando:
                self.btn_procesar.config(state="normal", bg=ACENTO)
        else:
            self.etiqueta_listo.config(text=f"⚠ Falta seleccionar: {', '.join(missing)}", fg=ADVERTENCIA)
            self.btn_procesar.config(state="disabled", bg=BORDE)

    def _generar_reporte_fragmentacion(self, _rebuild_only=False):
        if self.df_inventario is None:
            messagebox.showwarning("Sin datos", "Primero carga un archivo Excel.")
            return

        if not _rebuild_only:
            if self.ventana_reporte is not None:
                try:
                    if self.ventana_reporte.winfo_exists():
                        self._generar_reporte_fragmentacion(_rebuild_only=True)
                        self.ventana_reporte.lift()
                        self.ventana_reporte.focus_force()
                        return
                except Exception:
                    pass
                self.ventana_reporte = None
                self.txt_reporte = None

        segmentos           = self._get_segments()
        paginas_ignoradas   = self._get_ignored_pages()

        try:
            fila_inicio = int(self.spin_inicio.get())
            fila_fin   = int(self.spin_fin.get())
        except ValueError:
            fila_inicio, fila_fin = 0, 0

        df_slice = self.df_inventario.copy()
        if fila_inicio > 0:
            df_slice = df_slice[df_slice.index >= fila_inicio]
        if fila_fin > 0:
            df_slice = df_slice[df_slice.index <= fila_fin]

        seg_labels = []
        for item in self.tabla_segmentos.get_children():
            vals = self.tabla_segmentos.item(item, "values")
            seg_labels.append(f"Folio {vals[0]} → Pág. PDF {vals[1]}")

        ign_page_labels = []
        for item in self.tabla_ignorados.get_children():
            vals = self.tabla_ignorados.item(item, "values")
            ign_page_labels.append(f"{vals[0]}  →  {vals[1]}")

        SEP  = "═" * 68
        SEP2 = "─" * 68
        lineas = []
        ts = _time.strftime("%Y-%m-%d  %H:%M:%S")

        lineas.append(SEP)
        lineas.append("  REPORTE DE FRAGMENTACIÓN")
        lineas.append(f"  Generado: {ts}")
        if self.ruta_excel:
            lineas.append(f"  Excel   : {self.ruta_excel.name}")
        if self.ruta_pdf:
            lineas.append(f"  PDF     : {self.ruta_pdf.name}")
        lineas.append(SEP)

        # Configuración activa
        lineas.append("")
        lineas.append("  ── CONFIGURACIÓN ACTIVA ──────────────────────────────────────")
        lineas.append("")
        lineas.append(f"  Folio inicio Excel : {self._folio_inicio}r   →  Pág. PDF inicio: {self._pdf_page_inicio}")
        lineas.append("")

        if seg_labels:
            lineas.append("  [SALTO PDF]  Segmentos adicionales:")
            for s in seg_labels:
                lineas.append(f"    • {s}")
        else:
            lineas.append("  [SALTO PDF]  Sin segmentos adicionales.")
        lineas.append("")

        if ign_page_labels:
            lineas.append("  [PAG-IGN]  Páginas PDF ignoradas (excluidas del output):")
            for s in ign_page_labels:
                lineas.append(f"    • {s}")
        else:
            lineas.append("  [PAG-IGN]  Sin páginas PDF ignoradas.")
        lineas.append("")

        # Opciones globales
        lineas.append("  [OPTS]  Opciones globales de proceso:")
        lineas.append(f"    • Ignorar saltos folios  : {'SI' if self.var_ignorar_saltos.get() else 'NO'}")
        lineas.append("")

        # Columnas del Excel
        lineas.append(SEP2)
        lineas.append("  COLUMNAS DEL EXCEL UTILIZADAS")
        lineas.append(SEP2)
        col_map = [
            ("Folios",        _cfg.COL_FOLIOS),
            ("Registro",      _cfg.COL_REGISTRO),
            ("Escribano",     _cfg.COL_ESCRIBANO.replace("\n", " ")),
            ("Protocolo",     _cfg.COL_PROTOCOLO),
            ("Lugar (Tóp.)",  _cfg.COL_LUGAR),
            ("Fecha inicial", _cfg.COL_FECHA_INI),
            ("Título est.",   _cfg.COL_TITULO_EST),
            ("Interesado 1",  _cfg.COL_INT1),
            ("Interesado 2",  _cfg.COL_INT2),
            ("Observaciones", _cfg.COL_OBS),
        ]
        for etiqueta, col in col_map:
            lineas.append(f"  {etiqueta:<15} : {col}")
        lineas.append("")

        # Mapeo
        lineas.append(SEP2)
        lineas.append("  MAPEO  FILA → FOLIO(S) → PÁGINAS PDF")
        lineas.append(SEP2)
        hdr = f"  {'Fila':<6} {'Reg.':<6} {'Folio(s)':<14} {'Págs. PDF (inicio-fin)':<28} {'Banderas'}"
        lineas.append(hdr)
        lineas.append("  " + "-" * 66)

        prev_active_seg = segmentos[0] if segmentos else None
        prev_pdf_last   = None
        prev_fila       = None

        for fila_excel, row in df_slice.iterrows():
            folio_str = str(row.get(_cfg.COL_FOLIOS, "")).strip()
            reg       = str(row.get(_cfg.COL_REGISTRO, "")).strip()

            if not folio_str or folio_str.lower() == "nan":
                lineas.append(f"  {str(fila_excel):<6} {reg:<6} {'(sin folio)':<14} {'\u2014':<28}")
                prev_fila = fila_excel
                continue

            paginas, err = servicio_folios.parsear_rango_folios(
                folio_str,
                folio_inicio_excel=self._folio_inicio,
                pag_pdf_inicio=self._pdf_page_inicio,
                segmentos=segmentos
            )
            if err or not paginas:
                lineas.append(f"  {str(fila_excel):<6} {reg:<6} {folio_str:<14} ERROR: {err}")
                prev_fila = fila_excel
                continue

            # Desplazar páginas nominales saltando las ignoradas
            if paginas_ignoradas and paginas:
                paginas = servicio_folios.desplazar_paginas_ignoradas(paginas, paginas_ignoradas)

            nominal_pages = paginas[:]
            paginas = list(paginas)

            if not paginas:
                lineas.append(f"  {str(fila_excel):<6} {reg:<6} {folio_str:<14} (sin páginas efectivas)")
                prev_fila = fila_excel
                continue

            p_min, p_max = paginas[0], paginas[-1]
            sub_lines    = []

            # Cambio de segmento
            seg_changed = False
            if len(segmentos) > 1:
                first_folio_txt = folio_str.split('-')[0].strip()
                abs_ini = servicio_folios.texto_folio_a_pagina_abs(first_folio_txt)
                if abs_ini is not None:
                    active_seg = servicio_folios._buscar_segmento(abs_ini, segmentos)
                    if active_seg != prev_active_seg:
                        seg_changed = True
                        sep = (
                            f"  {'':6} {'':6}  ↳ CAMBIO DE SEGMENTO"
                            + (f" (fila {prev_fila} terminó en pág. {prev_pdf_last})" if prev_pdf_last else "")
                            + f"  →  nuevo segmento inicia en pág.PDF {active_seg[1]}"
                        )
                        lineas.append(sep)
                    prev_active_seg = active_seg

            # Sucesión con el registro anterior
            if prev_pdf_last is not None and not seg_changed:
                expected = prev_pdf_last + 1
                if p_min > expected:
                    sub_lines.append(
                        f"             ⚠ BRECHA: {p_min - expected} pág(s) sin asignar "
                        f"(págs. {expected}–{p_min-1}) tras fila {prev_fila}"
                    )
                elif p_min < expected:
                    sub_lines.append(
                        f"             ⚠ SOLAPAMIENTO: {expected - p_min} pág(s) "
                        f"compartidas con fila {prev_fila} (desde pág. {p_min})"
                    )

            # Páginas ignoradas en este rango
            pag_ign = sorted(p for p in nominal_pages if p in paginas_ignoradas) if paginas_ignoradas else []
            if pag_ign:
                sub_lines.append(
                    f"             ↳ [PAG-IGN] Excluidas del output PDF: "
                    f"{pag_ign[:10]}{'…' if len(pag_ign)>10 else ''}"
                )

            pag_desc = f"{p_min}–{p_max}  ({len(paginas)} págs.)"

            inline = ""
            if seg_changed:       inline = "[NUEVO SEG.]"
            elif pag_ign:         inline = f"[{len(pag_ign)} pag.ignorada(s)]"
            elif sub_lines:       inline = "[VER DETALLE]"

            lineas.append(f"  {str(fila_excel):<6} {reg:<6} {folio_str:<14} {pag_desc:<34} {inline}")
            lineas.extend(sub_lines)

            prev_pdf_last = p_max
            prev_fila     = fila_excel

        lineas.append("")
        lineas.append(SEP)
        lineas.append("  FIN DEL REPORTE")
        lineas.append(SEP)
        report_text = "\n".join(lineas)
        self._cache_lineas_reporte = lineas
        self._cache_texto_reporte  = report_text

        if _rebuild_only and self.txt_reporte is not None:
            self._insertar_lineas_en_txt_reporte(lineas)
            return

        # Ventana de reporte
        win = tk.Toplevel(self)
        win.title("Reporte de Fragmentacion")
        win.configure(bg=FONDO)
        win.geometry("960x680")
        win.minsize(700, 400)
        self.ventana_reporte = win
        win.protocol("WM_DELETE_WINDOW", self._al_cerrar_reporte)

        top_bar = tk.Frame(win, bg=FONDO)
        top_bar.pack(fill="x", padx=16, pady=(12, 4))
        tk.Label(top_bar, text="📋  Reporte de Fragmentacion",
                 font=FUENTE_PASO, bg=FONDO, fg=TEXTO).pack(side="left")

        btn_refresh = self._crear_boton(
            top_bar, "🔄  Actualizar", self._do_refresh_report, tipo="success")
        btn_refresh.configure(pady=4, padx=12)
        btn_refresh.pack(side="right", padx=(0, 8))

        def _save_report():
            from tkinter import filedialog as _fd
            path = _fd.asksaveasfilename(
                title="Guardar reporte", defaultextension=".txt",
                filetypes=[("Texto", "*.txt"), ("Todos", "*.*")],
                initialfile=f"reporte_fragmentacion_{_time.strftime('%Y%m%d_%H%M%S')}.txt",
            )
            if path:
                try:
                    with open(path, "w", encoding="utf-8") as f:
                        f.write(self._cache_texto_reporte or "")
                    messagebox.showinfo("Guardado", f"Reporte guardado en:\n{path}")
                except Exception as exc:
                    messagebox.showerror("Error", f"No se pudo guardar:\n{exc}")

        btn_save = self._crear_boton(
            top_bar, "💾  Guardar .txt", _save_report, tipo="info")
        btn_save.configure(pady=4, padx=12)
        btn_save.pack(side="right")

        txt = tk.Text(
            win, bg=FONDO_ENTRADA, fg=TEXTO, font=FUENTE_LOG,
            relief="flat", wrap="none", padx=12, pady=10,
            insertbackground=TEXTO,
        )
        sb_y = ttk.Scrollbar(win, orient="vertical",   command=txt.yview)
        sb_x = ttk.Scrollbar(win, orient="horizontal", command=txt.xview)
        txt.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
        sb_y.pack(side="right",  fill="y")
        sb_x.pack(side="bottom", fill="x")
        txt.pack(fill="both", expand=True, padx=(16, 0), pady=(4, 0))

        txt.tag_configure("header",    foreground=ACENTO_CLARO, font=("Consolas", 9, "bold"))
        txt.tag_configure("seg_flag",  foreground="#ffd700")
        txt.tag_configure("pag_flag",  foreground=ERROR)
        txt.tag_configure("fol_flag",  foreground=ADVERTENCIA)
        txt.tag_configure("exact_flag",foreground="#22d3ee")
        txt.tag_configure("ovr_flag",  foreground="#38bdf8")
        txt.tag_configure("nom_flag",  foreground="#c084fc")
        txt.tag_configure("dup_flag",  foreground="#f472b6")
        txt.tag_configure("rep_flag",  foreground="#fb923c")
        txt.tag_configure("opts_flag", foreground="#86efac")
        txt.tag_configure("col_key",   foreground=TEXTO_ATENUADO)
        txt.tag_configure("normal",    foreground=TEXTO)
        self.txt_reporte = txt

        self._insertar_lineas_en_txt_reporte(lineas)

    def _insertar_lineas_en_txt_reporte(self, lineas):
        txt = self.txt_reporte
        if txt is None:
            return
        txt.config(state="normal")
        txt.delete("1.0", "end")
        for line in lineas:
            if line.startswith("=") or line.startswith("-") or line.startswith("\u2550") or line.startswith("\u2500"):
                txt.insert("end", line + "\n", "header")
            elif any(line.strip().startswith(k) for k in
                     ("REPORTE", "FIN DEL", "CONFIGURACI", "COLUMNAS", "MAPEO", "OPTS")):
                txt.insert("end", line + "\n", "header")
            elif "[EXACT-OVR]" in line:
                txt.insert("end", line + "\n", "exact_flag")
            elif "[PAG-OVR]" in line:
                txt.insert("end", line + "\n", "ovr_flag")
            elif "[NOM]" in line or "[NOM:" in line:
                txt.insert("end", line + "\n", "nom_flag")
            elif "[DUP-IGN]" in line or "[OMITIDA" in line:
                txt.insert("end", line + "\n", "dup_flag")
            elif "[OPTS]" in line or "Bypass" in line or "Forzar" in line:
                txt.insert("end", line + "\n", "opts_flag")
            elif "[SALTO" in line:
                txt.insert("end", line + "\n", "seg_flag")
            elif "[PAG-IGN" in line:
                txt.insert("end", line + "\n", "pag_flag")
            elif "[FOL-IGN" in line:
                txt.insert("end", line + "\n", "fol_flag")
            elif "[REP-FOL" in line:
                txt.insert("end", line + "\n", "rep_flag")
            elif " : " in line and line.strip()[:10].replace(" ", "").isalpha():
                txt.insert("end", line + "\n", "col_key")
            else:
                txt.insert("end", line + "\n", "normal")
        txt.config(state="disabled")

    def _al_cerrar_reporte(self):
        try:
            if self.ventana_reporte:
                self.ventana_reporte.destroy()
        except Exception:
            pass
        self.ventana_reporte = None
        self.txt_reporte = None

    def _do_refresh_report(self):
        if self.txt_reporte is None or self.df_inventario is None:
            return
        try:
            self._generar_reporte_fragmentacion(_rebuild_only=True)
        except Exception:
            import traceback
            txt = self.txt_reporte
            txt.config(state="normal")
            txt.delete("1.0", "end")
            txt.insert("end", f"[ERROR AL ACTUALIZAR EL REPORTE]\n\n{traceback.format_exc()}", "pag_flag")
            txt.config(state="disabled")

    def _auto_refresh_report(self):
        if self.ventana_reporte is None or self.txt_reporte is None:
            return
        try:
            if self.ventana_reporte.winfo_exists():
                self._do_refresh_report()
        except Exception:
            pass
