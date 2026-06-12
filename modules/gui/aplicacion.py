import tkinter as tk
import queue
from pathlib import Path
import configuracion as _cfg
from modules.gui.estilos import FONDO
from modules.gui.constructor_ui import ConstructorUiMixin
from modules.gui.manejadores_eventos import ManejadorEventosMixin
from modules.gui.analizadores import AnalizadorMixin
from modules.gui.procesador import ProcesadorMixin

class Aplicacion(tk.Tk, ConstructorUiMixin, ManejadorEventosMixin, AnalizadorMixin, ProcesadorMixin):
    def __init__(self):
        super().__init__()
        self.title("Fragmentación de Archivos — Archivos Históricos")
        self.configure(bg=FONDO)
        self.minsize(960, 780)
        self.geometry("1060x840")

        try:
            self.iconbitmap(default="")
        except Exception:
            pass

        # Estado en español
        self.ruta_excel = None
        self.ruta_pdf = None
        self.directorio_salida = _cfg.DIRECTORIO_SALIDA
        self.df_inventario = None
        self.procesando = False
        self.bandera_cancelacion = False
        
        # Ventana de reporte persistente
        self.ventana_reporte = None
        self.txt_reporte = None
        self._cache_lineas_reporte = []
        self._cache_texto_reporte  = ""

        # Offset de inicio (folio Excel vs. página PDF)
        self._folio_inicio = 1   # número de folio con que inicia el protocolo en el Excel
        self._pdf_page_inicio = 1   # página del PDF donde empieza la primera imagen del protocolo

        # Cola para logs en tiempo real
        self.log_queue = queue.Queue()

        # Anti-flicker
        self._last_canvas_width = 0
        self._last_content_height = 0
        self._resize_after_id = None

        self._construir_ui()
        self._centrar_ventana()
        self._poll_log_queue()

    def _centrar_ventana(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def _log(self, mensaje: str, tag: str = "info"):
        """Escribe un mensaje en la bitácora de actividad en tiempo real."""
        self.txt_bitacora.config(state="normal")
        self.txt_bitacora.insert("end", mensaje + "\n", tag)
        self.txt_bitacora.see("end")
        self.txt_bitacora.config(state="disabled")
