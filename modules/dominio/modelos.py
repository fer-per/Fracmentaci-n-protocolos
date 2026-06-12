from dataclasses import dataclass, field
from typing import Optional, List, Set, Tuple

@dataclass
class RegistroDocumento:
    fila_excel: int
    registro_id: str
    escribano: str
    protocolo: str
    folios_origen: str
    titulo: str
    fecha_inicio: str
    fecha_fin: str
    interesado1: str
    interesado2: str
    lugar: str
    observaciones: str = ""

    def a_diccionario(self) -> dict:
        """Helper para mantener compatibilidad con interfaces basadas en diccionarios legados."""
        return {
            "N° DE REGISTRO": self.registro_id,
            "ESCRIBANO/NOTARIO": self.escribano,
            "PROTOCOLO": self.protocolo,
            "N° DE FOLIOS": self.folios_origen,
            "TITULO/ESCRITURA": self.titulo,
            "FECHA INICIAL": self.fecha_inicio,
            "FECHA FINAL": self.fecha_fin,
            "INTERESADO 1": self.interesado1,
            "INTERESADO 2": self.interesado2,
            "LUGAR (DATA TÓPICA)": self.lugar,
            "OBSERVACIONES": self.observaciones,
        }

@dataclass
class MetadatosArchivo:
    siglo: str
    num_acervo: str
    siglo_original: str = ""
    acervo_original: str = ""

@dataclass
class SegmentoDesplazamiento:
    pag_abs_inicio: int
    pag_pdf_inicio: int
    
@dataclass
class ResultadoValidacion:
    es_valido: bool
    mensaje_error: str
