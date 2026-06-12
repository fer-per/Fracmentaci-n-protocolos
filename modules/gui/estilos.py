import tkinter as tk

# ─── Colores y estilos ──────────────────────────────────
FONDO           = "#f7f4f3"
FONDO_TARJETA   = "#ffffff"
FONDO_ENTRADA   = "#efe7e5"
FONDO_SELECCION = "#e3d7d4"

TEXTO           = "#32131c"
TEXTO_ATENUADO  = "#5b3340"
TEXTO_MUTED     = "#83626d"

ACENTO          = "#800020"        # Burgundy
ACENTO_CLARO    = "#b04a63"
ACENTO_SELECCION= "#660019"

EXITO           = "#3c7a57"
EXITO_ATENUADO  = "#63a67e"

ADVERTENCIA     = "#c58a2b"        # Dorado oscuro
ERROR           = "#a1283d"

BORDE           = "#d8c8c4"

FUENTE          = ("Segoe UI", 11)
FUENTE_NEGRITA  = ("Segoe UI", 11, "bold")
FUENTE_TITULO   = ("Segoe UI", 20, "bold")
FUENTE_SUBTITULO= ("Segoe UI", 11)
FUENTE_PEQUENA  = ("Segoe UI", 9)
FUENTE_PASO     = ("Segoe UI", 12, "bold")
FUENTE_LOG      = ("Consolas", 10)
FUENTE_BOTON_G  = ("Segoe UI", 13, "bold")

# ── Paleta de botones por tipo (Estilo Burgundy) ───────────────────────────────
BOTON_PRINCIPAL = {"bg": "#7b1e3a", "hover": "#65182f", "fg": "#ffffff"}  # Vino principal
BOTON_EXITO     = {"bg": "#2e7d5b", "hover": "#25664a", "fg": "#ffffff"}  # Verde elegante
BOTON_ADVERTENCIA= {"bg": "#c76a2a", "hover": "#ad5920", "fg": "#ffffff"}  # Cobre / ámbar
BOTON_PELIGRO   = {"bg": "#b23a48", "hover": "#952f3c", "fg": "#ffffff"}  # Rojo vino
BOTON_SECUNDARIO= {"bg": "#f0e6e8", "hover": "#e2cfd4", "fg": "#4a1f2d"}  # Rosa grisáceo
BOTON_INFO      = {"bg": "#4a1f2d", "hover": "#3a1823", "fg": "#ffffff"}  # Vino oscuro
BOTON_PURPURA   = {"bg": "#6f3aa8", "hover": "#5d2f8d", "fg": "#ffffff"}  # Púrpura sofisticado


# ─── Tooltip helper ──────────────────────────────────────────────────────────
class EtiquetaInformacion:
    """Etiqueta emergente (Tooltip) para widgets. Aparece al pasar el mouse."""
    def __init__(self, widget, texto):
        self.widget = widget
        self.texto = texto
        self.ventana_consejo = None
        widget.bind("<Enter>", self._mostrar)
        widget.bind("<Leave>", self._ocultar)

    def _mostrar(self, evento=None):
        if self.ventana_consejo or not self.texto:
            return
        x, y, cx, cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 27
        y = y + self.widget.winfo_rooty() + 27
        self.ventana_consejo = vc = tk.Toplevel(self.widget)
        vc.wm_overrideredirect(True)
        vc.wm_geometry(f"+{x}+{y}")
        
        # Estilo Notion para el tooltip (muy limpio)
        marco = tk.Frame(vc, bg=FONDO_TARJETA, highlightthickness=1,
                         highlightbackground=BORDE, highlightcolor=BORDE)
        marco.pack()
        etiqueta = tk.Label(marco, text=self.texto, justify="left",
                           background=FONDO_TARJETA, foreground=TEXTO,
                           font=FUENTE_PEQUENA, relief="flat", padx=6, pady=4)
        etiqueta.pack(ipadx=1)

    def _ocultar(self, evento=None):
        vc = self.ventana_consejo
        self.ventana_consejo = None
        if vc:
            vc.destroy()
