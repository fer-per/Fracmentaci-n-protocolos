"""
configuracion.py — Configuración central del sistema.

Todas las rutas, constantes de filas Excel y nombres de columna están aquí.
Los alias al final permiten compatibilidad con código que aún use los
nombres del config.py original mientras se completa la migración.
"""
from pathlib import Path

# ─── Carpetas de salida ───────────────────────────────────────────────────────
DIRECTORIO_SALIDA = Path("output")
DIRECTORIO_LOGS   = Path("logs")

CSV_PENDIENTES = DIRECTORIO_LOGS / "pendientes.csv"

# ─── Parámetros del Excel ─────────────────────────────────────────────────────
# Estructura del Excel:
#   Fila 4:  Seccion: XVI          → Siglo (nivel 2 de la jerarquía)
#   Fila 7:  Código del fondo: N7  → Número de acervo (nivel 1 de la jerarquía)
#   Fila 8:  Encabezados de columnas
#   Fila 9:  Sub-encabezados (FECHA INICIAL / FECHA FINAL)
#   Filas 10+: Marcadores de sección (Protocolo Nº X, Registro Nº X) + datos
FILAS_A_OMITIR = 7   # skiprows para pandas (salta filas 1-7, usa fila 8 como header)

# Filas del Excel donde se leen los metadatos globales del fondo
FILA_META_SIGLO   = 4   # «Seccion: XVI»
FILA_META_ACERVO  = 7   # «Código del fondo: N7»

# Nombres de columna tal como aparecen en la fila 8 del Excel (después de strip).
COL_REGISTRO   = "N° DE REGISTRO"
COL_ESCRIBANO  = "ESCRIBANO/\nNOTARIO"    # la celda tiene salto de línea
COL_PROTOCOLO  = "N° DE PROT."
COL_FOLIOS     = "N° DE FOLIOS"
COL_LUGAR      = "DATA TÓPICA (Lugar)"
COL_TOPICA     = COL_LUGAR
COL_FECHA_INI  = "FECHA INICIAL"          # sub-columna de DATA CRÓNICA
COL_FECHA_FIN  = "FECHA FINAL"            # sub-columna de DATA CRÓNICA
COL_TITULO     = "TÍTULO/\nESCRITURA"     # título libre
COL_TITULO_EST = "Titulo estandar"        # columna I – se usa para la carpeta
COL_INT1       = "INTERESADO 1"
COL_INT2       = "INTERESADO 2"
COL_OBS        = "OBSERVACIONES"

# ─── Aliases de compatibilidad (nombres del config.py original) ───────────────
# Permiten que módulos aún no migrados usen 'import configuracion as config'.
OUTPUT_DIR      = DIRECTORIO_SALIDA
LOGS_DIR        = DIRECTORIO_LOGS
PENDIENTES_CSV  = CSV_PENDIENTES
SKIP_ROWS       = FILAS_A_OMITIR
META_ROW_SIGLO  = FILA_META_SIGLO
META_ROW_ACERVO = FILA_META_ACERVO
