from abc import ABC, abstractmethod
from typing import List, Set, Tuple
from modules.dominio.modelos import RegistroDocumento, MetadatosArchivo

class RepositorioExcel(ABC):
    @abstractmethod
    def cargar_metadatos(self, ruta_excel: str) -> MetadatosArchivo:
        """Lee los metadatos globales del fondo (siglo, num_acervo) desde las cabeceras del Excel."""
        pass

    @abstractmethod
    def cargar_registros(self, ruta_excel: str, filas_omitidas: int) -> List[RegistroDocumento]:
        """Carga y parsea los registros de documentos de la hoja de cálculo."""
        pass


class ServicioPdf(ABC):
    @abstractmethod
    def obtener_cantidad_paginas(self, ruta_pdf: str) -> int:
        """Devuelve el número total de páginas físicas de un archivo PDF."""
        pass

    @abstractmethod
    def extraer_paginas(
        self,
        ruta_pdf: str,
        numeros_paginas: List[int],
        ruta_destino: str,
        simulacion: bool = False,
        paginas_ignoradas: Set[int] = None,
    ) -> bool:
        """Separa páginas específicas de un PDF de origen en un archivo de destino."""
        pass


class ServicioAlmacenamiento(ABC):
    @abstractmethod
    def construir_ruta_destino(
        self,
        directorio_salida: str,
        metadatos: MetadatosArchivo,
        registro: RegistroDocumento,
        simulacion: bool = False,
    ) -> str:
        """Construye y sanitiza la jerarquía de directorios de 11 niveles y el nombre de archivo para un registro."""
        pass

    @abstractmethod
    def registrar_pendiente_csv(
        self,
        ruta_csv: str,
        registro: RegistroDocumento,
        motivo: str,
        simulacion: bool = False,
    ) -> None:
        """Registra un archivo omitido o con errores en el archivo CSV de pendientes."""
        pass
