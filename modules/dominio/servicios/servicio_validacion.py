from typing import Optional, List, Tuple, Set
from modules.dominio.modelos import RegistroDocumento, ResultadoValidacion
import modules.dominio.servicios.servicio_folios as servicio_folios

def validar_registro(
    registro: RegistroDocumento,
    total_paginas_pdf: int,
    prev_ultima_pagina: Optional[int] = None,
    verificar_saltos: bool = True,
    folio_inicio_excel: int = 1,
    pag_pdf_inicio: int = 1,
    segmentos: List[Tuple[int, int]] = None,
    paginas_pdf_ignoradas: Set[int] = None,
) -> ResultadoValidacion:
    """
    Valida un RegistroDocumento antes de procesarlo.
    """
    folios_val = registro.folios_origen

    # 1. Folios vacíos / nulos
    if not folios_val or str(folios_val).strip() == "" or str(folios_val).strip().lower() == "nan":
        return ResultadoValidacion(es_valido=False, mensaje_error="Folios vacíos o nulos")

    # 2. Formato inválido + cálculo de páginas con offset
    paginas, error_parseo = servicio_folios.parsear_rango_folios(
        str(folios_val),
        folio_inicio_excel=folio_inicio_excel,
        pag_pdf_inicio=pag_pdf_inicio,
        segmentos=segmentos,
    )
    if error_parseo:
        return ResultadoValidacion(es_valido=False, mensaje_error=error_parseo)

    # Ajustar páginas saltando las ignoradas
    if paginas_pdf_ignoradas and paginas:
        paginas = servicio_folios.desplazar_paginas_ignoradas(paginas, paginas_pdf_ignoradas)

    # 3. Páginas fuera del rango del PDF
    max_pag = max(paginas)
    min_pag = min(paginas)
    if max_pag > total_paginas_pdf:
        return ResultadoValidacion(
            es_valido=False,
            mensaje_error=(
                f"Folio '{folios_val}' requiere página {max_pag} "
                f"pero el PDF solo tiene {total_paginas_pdf} páginas"
            )
        )
    if min_pag < 1:
        return ResultadoValidacion(
            es_valido=False,
            mensaje_error=(
                f"Folio '{folios_val}' resulta en número de página < 1 "
                f"(revisa el folio de inicio del Excel y la página de inicio del PDF)"
            )
        )

    # 4. Detección de salto de secuencia
    if verificar_saltos and prev_ultima_pagina is not None:
        primer_actual = servicio_folios.primera_pagina_de_rango(
            str(folios_val),
            folio_inicio_excel=folio_inicio_excel,
            pag_pdf_inicio=pag_pdf_inicio,
            segmentos=segmentos,
        )
        if primer_actual is not None:
            proxima_esperada = prev_ultima_pagina + 1
            if primer_actual != proxima_esperada:
                paginas_faltantes = list(range(proxima_esperada, primer_actual))
                return ResultadoValidacion(
                    es_valido=False,
                    mensaje_error=(
                        f"Salto de secuencia detectado: "
                        f"se esperaba página {proxima_esperada} "
                        f"pero se encontró página {primer_actual} "
                        f"(folios '{folios_val}'). "
                        f"Páginas faltantes: {paginas_faltantes}"
                    )
                )

    return ResultadoValidacion(es_valido=True, mensaje_error="")
