"""
modules/lector_excel.py
------------------------
Adaptador para lectura directa de Excel como DataFrame (para la GUI)
y metadatos de cabecera.
"""
import logging
from pathlib import Path
import pandas as pd
import openpyxl
import re

import configuracion as _cfg
from modules.infraestructura.adaptadores import RepositorioExcelPandas

logger = logging.getLogger(__name__)
_repositorio = RepositorioExcelPandas()


def cargar_metadatos_excel(
    ruta_excel: Path,
    fila_siglo: int = 4,
    fila_acervo: int = 7,
) -> dict:
    """
    Devuelve dict con 'siglo', 'acervo_num', 'raw_siglo', 'raw_acervo'.
    """
    orig_siglo  = _cfg.META_ROW_SIGLO
    orig_acervo = _cfg.META_ROW_ACERVO
    _cfg.META_ROW_SIGLO  = fila_siglo
    _cfg.META_ROW_ACERVO = fila_acervo
    try:
        metadatos = _repositorio.cargar_metadatos(str(ruta_excel))
    finally:
        _cfg.META_ROW_SIGLO  = orig_siglo
        _cfg.META_ROW_ACERVO = orig_acervo

    return {
        "siglo":      metadatos.siglo,
        "acervo_num": metadatos.num_acervo,
        "raw_siglo":  metadatos.siglo_original,
        "raw_acervo": metadatos.acervo_original,
    }


def cargar_excel_dataframe(
    ruta_excel: Path,
    filas_omitidas: int = 7,
    col_fecha_ini: str = "FECHA INICIAL",
) -> pd.DataFrame:
    """
    Devuelve un DataFrame con el índice = fila real del Excel.
    Construye el DataFrame directamente para mantener la compatibilidad con la GUI.
    """
    path = Path(ruta_excel)
    if not path.exists():
        raise FileNotFoundError(f"Archivo Excel no encontrado: {path}")

    logger.info(f"Cargando Excel: {path}")

    df = pd.read_excel(
        path,
        skiprows=filas_omitidas,
        header=0,
        engine="openpyxl",
        dtype=str,
    )

    df.columns = [str(c).strip() for c in df.columns]

    columnas_anonimas = [c for c in df.columns if str(c).startswith("Unnamed:")]
    if columnas_anonimas:
        df.drop(columns=columnas_anonimas, inplace=True)

    mapeo_nombres = {}
    columnas_cronicas = [c for c in df.columns if "DATA CR" in c.upper()]
    if len(columnas_cronicas) >= 2:
        mapeo_nombres[columnas_cronicas[0]] = col_fecha_ini
        mapeo_nombres[columnas_cronicas[1]] = "FECHA FINAL"
    elif len(columnas_cronicas) == 1:
        mapeo_nombres[columnas_cronicas[0]] = col_fecha_ini
    df.rename(columns=mapeo_nombres, inplace=True)

    df.index = df.index + filas_omitidas + 2
    df.dropna(how="all", inplace=True)
    df = df.fillna("")

    primera_columna = df.columns[0]
    _patron_marcador_seccion = re.compile(r'^\s*(protocolo|registro)\b', re.IGNORECASE)

    def _es_fila_de_datos(val: str) -> bool:
        v = str(val).strip()
        if not v or _patron_marcador_seccion.match(v):
            return False
        return True

    antes = len(df)
    df = df[df[primera_columna].apply(_es_fila_de_datos)]
    eliminados = antes - len(df)
    if eliminados:
        logger.info(f"Eliminados {eliminados} marcadores de sección")

    logger.info(f"Excel cargado: {len(df)} registros")
    return df
