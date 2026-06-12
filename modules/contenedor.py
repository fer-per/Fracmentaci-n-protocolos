"""
modules/contenedor.py - Contenedor de Dependencias (Dependency Injection Wiring)

Único lugar del sistema donde se instancian y conectan las implementaciones
concretas con los casos de uso. La GUI importa los casos de uso
ya construidos desde aquí.
"""
from modules.infraestructura.adaptadores import (
    RepositorioExcelPandas,
    ServicioPdfPyPdf,
    ServicioAlmacenamientoWindows,
)
from modules.aplicacion.casos_uso import (
    CasoUsoFragmentarPdf,
    CasoUsoAnalizarFolios,
    CasoUsoAnalizarTopica,
    CasoUsoAnalizarCronica,
    CasoUsoVerificarCobertura,
)

# ─── Adaptadores (Singletons simples) ───────────────────────────────────────────
repositorio_excel      = RepositorioExcelPandas()
servicio_pdf           = ServicioPdfPyPdf()
servicio_almacenamiento = ServicioAlmacenamientoWindows()

# ─── Casos de Uso ya cableados (Listos para usar en la GUI) ─────────────────────
caso_uso_fragmentar_pdf = CasoUsoFragmentarPdf(
    repo_excel=repositorio_excel,
    servicio_pdf=servicio_pdf,
    servicio_almacenamiento=servicio_almacenamiento,
)

caso_uso_analizar_folios = CasoUsoAnalizarFolios(
    repo_excel=repositorio_excel,
)

caso_uso_analizar_topica = CasoUsoAnalizarTopica(
    repo_excel=repositorio_excel,
)

caso_uso_analizar_cronica = CasoUsoAnalizarCronica(
    repo_excel=repositorio_excel,
)

caso_uso_verificar_cobertura = CasoUsoVerificarCobertura(
    repo_excel=repositorio_excel,
    servicio_pdf=servicio_pdf,
)
