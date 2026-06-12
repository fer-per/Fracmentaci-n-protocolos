from typing import List, Set, Tuple, Dict, Callable, Optional
from modules.dominio.modelos import RegistroDocumento, MetadatosArchivo, ResultadoValidacion
import modules.dominio.servicios.servicio_folios as servicio_folios
import modules.dominio.servicios.servicio_validacion as servicio_validacion
import modules.dominio.servicios.servicio_analisis as servicio_analisis
from modules.aplicacion.puertos import RepositorioExcel, ServicioPdf, ServicioAlmacenamiento

class CasoUsoVerificarCobertura:
    def __init__(self, repo_excel: RepositorioExcel, servicio_pdf: ServicioPdf):
        self.repo_excel = repo_excel
        self.servicio_pdf = servicio_pdf

    def ejecutar(
        self,
        ruta_excel: str,
        ruta_pdf: str,
        filas_omitidas: int,
        fila_inicio: int,
        fila_fin: int,
        folio_inicio_excel: int,
        pag_pdf_inicio: int,
        segmentos: List[Tuple[int, int]]
    ) -> dict:
        total_paginas_pdf = self.servicio_pdf.obtener_cantidad_paginas(ruta_pdf)
        registros = self.repo_excel.cargar_registros(ruta_excel, filas_omitidas)
        
        # Filtrar rango de filas
        registros_filtrados = [
            r for r in registros 
            if (fila_inicio <= r.fila_excel <= fila_fin if fila_fin > 0 else r.fila_excel >= fila_inicio)
        ]
        
        if not registros_filtrados:
            return {"ok": False, "resumen": "[ERR] No hay registros en el rango seleccionado."}

        max_pagina_usada = 0
        ultimo_folio_valido = ""
        ultima_fila_valida = None
        cantidad_invalidos = 0

        for registro in registros_filtrados:
            paginas, err = servicio_folios.parsear_rango_folios(
                registro.folios_origen,
                folio_inicio_excel=folio_inicio_excel,
                pag_pdf_inicio=pag_pdf_inicio,
                segmentos=segmentos
            )
            if err or not paginas:
                cantidad_invalidos += 1
                continue
            
            p_max = max(paginas)
            if p_max > max_pagina_usada:
                max_pagina_usada = p_max
                ultimo_folio_valido = registro.folios_origen
                ultima_fila_valida = registro.fila_excel

        lineas = []
        lineas.append("ANÁLISIS DE COBERTURA PDF:")
        lineas.append(f"  Registros en rango : {len(registros_filtrados)}")
        if cantidad_invalidos:
            lineas.append(f"  [!!]  Folios inválidos ignorados: {cantidad_invalidos}")
        lineas.append("")

        if max_pagina_usada == 0:
            lineas.append("[ERR]  No se encontró ningún folio válido en el rango. Verifica el Excel y los segmentos.")
            return {"ok": False, "resumen": "\n".join(lineas)}

        lineas.append(f"  Última página PDF usada por el rango: {max_pagina_usada}")
        lineas.append(f"    (corresponde al folio '{ultimo_folio_valido}', fila Excel {ultima_fila_valida})")
        lineas.append(f"  Última página real del PDF           : {total_paginas_pdf}")
        lineas.append("")

        diferencia = total_paginas_pdf - max_pagina_usada
        ok = False
        if diferencia == 0:
            lineas.append("[OK]  PERFECTO: La última hoja del PDF coincide exactamente con el")
            lineas.append("      final del rango. El PDF se separará sin que sobre ninguna hoja.")
            ok = True
        elif diferencia > 0:
            lineas.append(f"[!!]  ADVERTENCIA: Sobrarían {diferencia} página(s) del PDF sin asignar.")
            lineas.append(f"      El rango solo cubre hasta la página {max_pagina_usada},")
            lineas.append(f"      pero el PDF tiene {total_paginas_pdf} páginas en total.")
            lineas.append("")
            lineas.append("      Posibles causas:")
            lineas.append("        - El rango de filas no llega hasta el último registro del protocolo.")
            lineas.append("        - Faltan registros en el Excel para las últimas hojas del PDF.")
            lineas.append("        - El offset (folio inicio / página PDF) no está bien configurado.")
            lineas.append("        - Faltan segmentos adicionales para cubrir los folios restantes.")
            ok = False
        else:
            lineas.append(f"[ERR]  ERROR: El rango requiere {-diferencia} página(s) MÁS de las que")
            lineas.append(f"      tiene el PDF ({total_paginas_pdf} páginas disponibles,")
            lineas.append(f"      pero el rango alcanza la página {max_pagina_usada}).")
            lineas.append("")
            lineas.append("      Posibles causas:")
            lineas.append("        - El offset (folio inicio / página PDF) es incorrecto.")
            lineas.append("        - El rango de filas excede lo que este PDF contiene.")
            ok = False

        return {"ok": ok, "resumen": "\n".join(lineas)}


class CasoUsoFragmentarPdf:
    def __init__(self, repo_excel: RepositorioExcel, servicio_pdf: ServicioPdf, servicio_almacenamiento: ServicioAlmacenamiento):
        self.repo_excel = repo_excel
        self.servicio_pdf = servicio_pdf
        self.servicio_almacenamiento = servicio_almacenamiento

    def ejecutar(
        self,
        ruta_excel: str,
        ruta_pdf: str,
        directorio_salida: str,
        ruta_csv: str,
        filas_omitidas: int,
        fila_inicio: int,
        fila_fin: int,
        verificar_saltos: bool,
        folio_inicio_excel: int,
        pag_pdf_inicio: int,
        segmentos: List[Tuple[int, int]],
        paginas_pdf_ignoradas: Set[int],
        sobreescritura_exacta: Dict[int, Tuple[int, int]],
        simulacion: bool = False,
        callback_progreso: Callable[[int, int, str], None] = None,
        callback_log: Callable[[str, str], None] = None,
        verificar_cancelacion: Callable[[], bool] = None
    ) -> dict:
        def loguear(mensaje, nivel="info"):
            if callback_log:
                callback_log(mensaje, nivel)

        loguear("INICIO DEL PROCESO ARCHIVÍSTICO ")
        metadatos = self.repo_excel.cargar_metadatos(ruta_excel)
        registros = self.repo_excel.cargar_registros(ruta_excel, filas_omitidas)
        
        # Filtrar rango
        registros_filtrados = [
            r for r in registros 
            if (fila_inicio <= r.fila_excel <= fila_fin if fila_fin > 0 else r.fila_excel >= fila_inicio)
        ]
        total_registros = len(registros_filtrados)
        
        if total_registros == 0:
            loguear(f"No hay filas en el rango {fila_inicio}–{fila_fin}. Abortando.", "error")
            return {"exito": False, "procesados": 0, "omitidos": 0}

        total_paginas_pdf = self.servicio_pdf.obtener_cantidad_paginas(ruta_pdf)
        
        procesados = 0
        omitidos = 0
        prev_ultima_pagina = None

        for idx, registro in enumerate(registros_filtrados):
            if verificar_cancelacion and verificar_cancelacion():
                loguear("Proceso cancelado por el usuario.", "warn")
                break

            reg_id = registro.registro_id or f"fila_{registro.fila_excel}"
            
            # Revisar si hay sobreescritura de página exacta (override)
            override = sobreescritura_exacta.get(registro.fila_excel)
            es_valido = True
            mensaje_error = ""
            paginas = []

            if override:
                p_inicio, p_fin = override
                paginas = list(range(p_inicio, p_fin + 1))
                loguear(f"Fila {registro.fila_excel} / Reg {reg_id}: Usando override exacto de páginas {p_inicio}-{p_fin}")
                if not paginas:
                    es_valido = False
                    mensaje_error = "Override de rango vacío"
                elif min(paginas) < 1 or max(paginas) > total_paginas_pdf:
                    es_valido = False
                    mensaje_error = f"Override fuera de límites del PDF (1-{total_paginas_pdf})"
            else:
                # Validación estándar
                res_val = servicio_validacion.validar_registro(
                    registro=registro,
                    total_paginas_pdf=total_paginas_pdf,
                    prev_ultima_pagina=prev_ultima_pagina,
                    verificar_saltos=verificar_saltos,
                    folio_inicio_excel=folio_inicio_excel,
                    pag_pdf_inicio=pag_pdf_inicio,
                    segmentos=segmentos,
                    paginas_pdf_ignoradas=paginas_pdf_ignoradas
                )
                es_valido = res_val.es_valido
                mensaje_error = res_val.mensaje_error
                
                if es_valido:
                    paginas, _ = servicio_folios.parsear_rango_folios(
                        registro.folios_origen,
                        folio_inicio_excel=folio_inicio_excel,
                        pag_pdf_inicio=pag_pdf_inicio,
                        segmentos=segmentos
                    )
                    if paginas_pdf_ignoradas and paginas:
                        paginas = servicio_folios.desplazar_paginas_ignoradas(paginas, paginas_pdf_ignoradas)

            if not es_valido:
                nivel = "warn" if "Salto" in mensaje_error else "info"
                loguear(f"Fila {registro.fila_excel} / Reg {reg_id}: OMITIDO - {mensaje_error}", nivel)
                self.servicio_almacenamiento.registrar_pendiente_csv(ruta_csv, registro, mensaje_error, simulacion)
                omitidos += 1
                if callback_progreso:
                    callback_progreso(idx + 1, total_registros, f"Fila {registro.fila_excel}: Omitido")
                continue

            # Construir ruta completa
            ruta_destino = self.servicio_almacenamiento.construir_ruta_destino(
                directorio_salida=directorio_salida,
                metadatos=metadatos,
                registro=registro,
                simulacion=simulacion
            )

            # Extraer páginas
            exito = self.servicio_pdf.extraer_paginas(
                ruta_pdf=ruta_pdf,
                numeros_paginas=paginas,
                ruta_destino=ruta_destino,
                simulacion=simulacion,
                paginas_ignoradas=paginas_pdf_ignoradas
            )

            if exito:
                loguear(f"Fila {registro.fila_excel} / Reg {reg_id}: OK - {ruta_destino} ({len(paginas)} pág.)")
                procesados += 1
                prev_ultima_pagina = paginas[-1] if paginas else prev_ultima_pagina
            else:
                loguear(f"Fila {registro.fila_excel} / Reg {reg_id}: ERROR al extraer páginas - {registro.folios_origen}", "error")
                self.servicio_almacenamiento.registrar_pendiente_csv(ruta_csv, registro, "Error al extraer páginas", simulacion)
                omitidos += 1

            if callback_progreso:
                callback_progreso(idx + 1, total_registros, f"Fila {registro.fila_excel}: OK")

        return {"exito": True, "procesados": procesados, "omitidos": omitidos}


class CasoUsoAnalizarFolios:
    def __init__(self, repo_excel: RepositorioExcel):
        self.repo_excel = repo_excel

    def ejecutar(
        self,
        ruta_excel: str,
        filas_omitidas: int,
        fila_inicio: int,
        fila_fin: int,
        segmentos: List[Tuple[int, int]]
    ) -> dict:
        registros = self.repo_excel.cargar_registros(ruta_excel, filas_omitidas)
        registros_filtrados = [
            r for r in registros 
            if (fila_inicio <= r.fila_excel <= fila_fin if fila_fin > 0 else r.fila_excel >= fila_inicio)
        ]
        
        lista_folios = [r.folios_origen for r in registros_filtrados]
        indices_filas = [r.fila_excel for r in registros_filtrados]
        
        return servicio_folios.analizar_secuencia_folios(
            lista_folios=lista_folios,
            indices=indices_filas,
            segmentos=segmentos
        )


class CasoUsoAnalizarTopica:
    def __init__(self, repo_excel: RepositorioExcel):
        self.repo_excel = repo_excel

    def ejecutar(
        self,
        ruta_excel: str,
        filas_omitidas: int,
        fila_inicio: int,
        fila_fin: int
    ) -> dict:
        registros = self.repo_excel.cargar_registros(ruta_excel, filas_omitidas)
        registros_filtrados = [
            r for r in registros 
            if (fila_inicio <= r.fila_excel <= fila_fin if fila_fin > 0 else r.fila_excel >= fila_inicio)
        ]
        return servicio_analisis.analizar_data_topica(registros_filtrados)


class CasoUsoAnalizarCronica:
    def __init__(self, repo_excel: RepositorioExcel):
        self.repo_excel = repo_excel

    def ejecutar(
        self,
        ruta_excel: str,
        filas_omitidas: int,
        fila_inicio: int,
        fila_fin: int
    ) -> dict:
        registros = self.repo_excel.cargar_registros(ruta_excel, filas_omitidas)
        registros_filtrados = [
            r for r in registros 
            if (fila_inicio <= r.fila_excel <= fila_fin if fila_fin > 0 else r.fila_excel >= fila_inicio)
        ]
        return servicio_analisis.analizar_data_cronica(registros_filtrados)
