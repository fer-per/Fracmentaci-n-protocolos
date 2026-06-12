import re
from typing import Optional, List, Tuple, Set

# Expresión regular: captura folio de inicio, cara (r/v), y opcionalmente folio de fin con su cara.
_EXPRESION_FOLIO = re.compile(
    r'^\s*(\d+)(r|v)(?:\s*-\s*(\d+)(r|v))?\s*$',
    re.IGNORECASE
)

def _folio_a_pagina_abs(folio: int, cara: str) -> int:
    cara = cara.lower()
    if cara == 'r':
        return 2 * folio - 1
    else:
        return 2 * folio

def texto_folio_a_pagina_abs(texto: str) -> Optional[int]:
    texto = str(texto).strip()
    m = _EXPRESION_FOLIO.match(texto)
    if not m:
        return None
    f = int(m.group(1))
    c = m.group(2).lower()
    return _folio_a_pagina_abs(f, c)

def _buscar_segmento(pag_abs: int, segmentos: List[Tuple[int, int]]) -> Tuple[int, int]:
    mejor = segmentos[0] if segmentos else (1, 1)
    for seg in sorted(segmentos, key=lambda x: x[0]):
        if pag_abs >= seg[0]:
            mejor = seg
        else:
            break
    return mejor

def _aplicar_desplazamiento(pag_abs: int, folio_inicio_excel: int, pag_pdf_inicio: int) -> int:
    primer_abs = _folio_a_pagina_abs(folio_inicio_excel, 'r')
    diferencia = primer_abs - pag_pdf_inicio
    return pag_abs - diferencia

def folio_a_pagina(folio: int, cara: str,
                   folio_inicio_excel: int = 1,
                   pag_pdf_inicio: int = 1,
                   segmentos: List[Tuple[int, int]] = None) -> int:
    pag_abs = _folio_a_pagina_abs(folio, cara)
    if segmentos:
        seg_pag_abs_inicio, seg_pdf_inicio = _buscar_segmento(pag_abs, segmentos)
        return pag_abs - (seg_pag_abs_inicio - seg_pdf_inicio)
    return _aplicar_desplazamiento(pag_abs, folio_inicio_excel, pag_pdf_inicio)

def parsear_rango_folios(
    texto: str,
    folio_inicio_excel: int = 1,
    pag_pdf_inicio: int = 1,
    segmentos: List[Tuple[int, int]] = None,
) -> Tuple[Optional[List[int]], Optional[str]]:
    if not texto or not str(texto).strip():
        return None, "Campo de folios vacío"

    texto = str(texto).strip()
    m = _EXPRESION_FOLIO.match(texto)
    if not m:
        return None, f"Formato de folios inválido: '{texto}'"

    f_ini = int(m.group(1))
    c_ini = m.group(2).lower()
    f_fin = int(m.group(3)) if m.group(3) else None
    c_fin = m.group(4).lower() if m.group(4) else None

    # Caso de folio simple
    if f_fin is None:
        pag_ini = folio_a_pagina(f_ini, c_ini, folio_inicio_excel, pag_pdf_inicio, segmentos)
        return [pag_ini], None

    abs_ini = _folio_a_pagina_abs(f_ini, c_ini)
    abs_fin = _folio_a_pagina_abs(f_fin, c_fin)

    if abs_fin < abs_ini:
        return None, (
            f"Rango incoherente: '{texto}' "
            f"(folio fin {f_fin}{c_fin} < folio inicio {f_ini}{c_ini})"
        )

    paginas = []
    vistas = set()
    for pag_abs in range(abs_ini, abs_fin + 1):
        if segmentos:
            seg_abs, seg_pdf = _buscar_segmento(pag_abs, segmentos)
            pag_pdf = pag_abs - (seg_abs - seg_pdf)
        else:
            primer_abs = _folio_a_pagina_abs(folio_inicio_excel, 'r')
            pag_pdf  = pag_abs - (primer_abs - pag_pdf_inicio)

        if pag_pdf < 1:
            continue

        if pag_pdf not in vistas:
            paginas.append(pag_pdf)
            vistas.add(pag_pdf)

    if not paginas:
        return None, f"El rango '{texto}' no produce páginas PDF válidas"

    paginas.sort()
    return paginas, None

def desplazar_paginas_ignoradas(paginas: List[int], paginas_pdf_ignoradas: Set[int]) -> List[int]:
    if not paginas_pdf_ignoradas:
        return list(paginas)

    ignoradas_ordenadas = sorted(paginas_pdf_ignoradas)
    resultado = []
    for nominal in paginas:
        fisica = nominal
        for ign in ignoradas_ordenadas:
            if ign <= fisica:
                fisica += 1
            else:
                break
        while fisica in paginas_pdf_ignoradas:
            fisica += 1
        resultado.append(fisica)
    return resultado

def ultima_pagina_de_rango(
    texto: str,
    folio_inicio_excel: int = 1,
    pag_pdf_inicio: int = 1,
    segmentos: List[Tuple[int, int]] = None,
) -> Optional[int]:
    paginas, err = parsear_rango_folios(texto, folio_inicio_excel, pag_pdf_inicio, segmentos)
    if err or not paginas:
        return None
    return paginas[-1]

def primera_pagina_de_rango(
    texto: str,
    folio_inicio_excel: int = 1,
    pag_pdf_inicio: int = 1,
    segmentos: List[Tuple[int, int]] = None,
) -> Optional[int]:
    paginas, err = parsear_rango_folios(texto, folio_inicio_excel, pag_pdf_inicio, segmentos)
    if err or not paginas:
        return None
    return paginas[0]

def analizar_secuencia_folios(
    lista_folios: List[str],
    indices: List[int] = None,
    segmentos: List[Tuple[int, int]] = None
) -> dict:
    saltos             = []
    solapamientos      = []
    entradas_invalidas = []
    folios_repetidos   = []

    prev_ultimo_folio = None
    prev_ultima_cara  = None
    prev_pag_ini_abs  = None
    prev_pag_fin_abs  = None
    prev_texto        = None
    prev_fila_id      = None

    folios_vistos: dict = {}

    for idx, texto in enumerate(lista_folios):
        fila_id = indices[idx] if (indices and idx < len(indices)) else (idx + 1)
        texto_s = str(texto).strip() if texto else ""
        if not texto_s or texto_s.lower() == "nan":
            continue

        m = _EXPRESION_FOLIO.match(texto_s)
        if not m:
            entradas_invalidas.append({
                "indice":  idx,
                "fila_id": fila_id,
                "texto":  texto_s,
                "detalle": f"Fila {fila_id}: Formato inválido -> '{texto_s}'",
            })
            prev_ultimo_folio = None
            prev_ultima_cara  = None
            prev_pag_ini_abs  = None
            prev_pag_fin_abs  = None
            prev_texto        = texto_s
            prev_fila_id      = fila_id
            continue

        f_ini = int(m.group(1))
        c_ini = m.group(2).lower()
        f_fin = int(m.group(3)) if m.group(3) else f_ini
        c_fin = m.group(4).lower() if m.group(4) else c_ini

        pag_ini_abs = _folio_a_pagina_abs(f_ini, c_ini)
        pag_fin_abs = _folio_a_pagina_abs(f_fin, c_fin)

        if pag_fin_abs < pag_ini_abs:
            entradas_invalidas.append({
                "indice":  idx,
                "fila_id": fila_id,
                "texto":  texto_s,
                "detalle": f"Fila {fila_id}: Rango incoherente '{texto_s}' (fin < inicio)",
            })
            prev_ultimo_folio = None
            prev_ultima_cara  = None
            prev_pag_ini_abs  = None
            prev_pag_fin_abs  = None
            prev_texto        = texto_s
            prev_fila_id      = fila_id
            continue

        if pag_ini_abs not in folios_vistos:
            folios_vistos[pag_ini_abs] = []
        folios_vistos[pag_ini_abs].append((fila_id, texto_s, pag_ini_abs, pag_fin_abs))

        if prev_ultimo_folio is not None:
            pag_esperada = _folio_a_pagina_abs(prev_ultimo_folio, prev_ultima_cara) + 1
            pag_actual   = pag_ini_abs

            if pag_actual > pag_esperada:
                saltos.append({
                    "indice":       idx,
                    "fila_id":      fila_id,
                    "prev_fila_id": prev_fila_id,
                    "prev_texto":  prev_texto,
                    "curr_texto":  texto_s,
                    "detalle": f"Fila {fila_id}: Salto de '{prev_texto}' a '{texto_s}'",
                })
            elif pag_actual < pag_esperada:
                es_repeticion = (
                    prev_pag_ini_abs is not None
                    and pag_ini_abs == prev_pag_ini_abs
                )
                if not es_repeticion:
                    solapamientos.append({
                        "indice":       idx,
                        "fila_id":      fila_id,
                        "prev_fila_id": prev_fila_id,
                        "prev_texto":  prev_texto,
                        "curr_texto":  texto_s,
                        "detalle": f"Fila {fila_id}: Solapamiento de '{prev_texto}' con '{texto_s}'",
                    })

        prev_ultimo_folio = f_fin
        prev_ultima_cara  = c_fin
        prev_pag_ini_abs  = pag_ini_abs
        prev_pag_fin_abs  = pag_fin_abs
        prev_texto        = texto_s
        prev_fila_id      = fila_id

    for ini_abs, ocurrencias in folios_vistos.items():
        if len(ocurrencias) > 1:
            folio_ini_lbl = _pagina_abs_a_etiqueta_folio(ini_abs)
            ocurrencias_info = []
            for occ_fila_id, occ_texto, occ_ini_abs, occ_fin_abs in ocurrencias:
                occ_n_paginas = occ_fin_abs - occ_ini_abs + 1
                if segmentos:
                    seg_abs, seg_pdf = _buscar_segmento(occ_ini_abs, segmentos)
                    pdf_ini_est = occ_ini_abs - (seg_abs - seg_pdf)
                    pdf_fin_est = pdf_ini_est + occ_n_paginas - 1
                else:
                    pdf_ini_est = None
                    pdf_fin_est = None
                occ_ini_lbl = _pagina_abs_a_etiqueta_folio(occ_ini_abs)
                occ_fin_lbl = _pagina_abs_a_etiqueta_folio(occ_fin_abs)
                occ_rango   = occ_ini_lbl if occ_ini_lbl == occ_fin_lbl else f"{occ_ini_lbl}-{occ_fin_lbl}"
                ocurrencias_info.append((
                    occ_fila_id, occ_texto,
                    pdf_ini_est, pdf_fin_est,
                    occ_rango, occ_n_paginas
                ))

            folios_repetidos.append({
                "folio_inicio":  folio_ini_lbl,
                "pag_ini_abs":  ini_abs,
                "ocurrencias":   ocurrencias_info,
                "detalle": (
                    f"Folio inicio '{folio_ini_lbl}' aparece "
                    f"{len(ocurrencias)} veces — filas: "
                    + ", ".join(str(r) for r, *_ in ocurrencias_info)
                ),
            })

    es_correcto = not saltos and not solapamientos and not entradas_invalidas

    lineas = []
    if es_correcto and not folios_repetidos:
        lineas.append("[OK]  Sucesión de folios correcta. No se detectaron saltos, solapamientos ni folios repetidos.")
    elif es_correcto and folios_repetidos:
        lineas.append("[OK]  Sucesión de folios correcta (sin saltos ni solapamientos reales).")
        lineas.append(f"[!!]  Se detectaron {len(folios_repetidos)} foliación(es) repetida(s) — ver abajo.")
    else:
        if entradas_invalidas:
            lineas.append(f"[!!]  {len(entradas_invalidas)} entrada(s) con formato inválido:")
            for e in entradas_invalidas:
                lineas.append(f"    . {e['detalle']}")
        if saltos:
            lineas.append(f"[!!]  {len(saltos)} salto(s) detectado(s):")
            for g in saltos:
                lineas.append(f"    . {g['detalle']}")
        if solapamientos:
            lineas.append(f"[!!]  {len(solapamientos)} solapamiento(s) real(es) detectado(s):")
            for o in solapamientos:
                lineas.append(f"    . {o['detalle']}")
        if folios_repetidos:
            lineas.append(f"[!!]  {len(folios_repetidos)} foliación(es) repetida(s) — ver abajo.")

    if folios_repetidos:
        lineas.append("")
        lineas.append("  ─── FOLIOS REPETIDOS (mismo folio inicio en varias filas) ──────")
        lineas.append("  Estas filas comparten el mismo folio de inicio.")
        lineas.append("  El PDF tiene DOS (o más) digitalizaciones del mismo folio físico.")
        lineas.append("  Cada segunda+ ocurrencia REQUIERE un Override Exacto de páginas PDF.")
        lineas.append("")
        for rep in folios_repetidos:
            n_occ = len(rep["ocurrencias"])
            lineas.append(f"  [REP] Folio inicio '{rep['folio_inicio']}'  →  {n_occ} ocurrencias:")
            for i, occ in enumerate(rep["ocurrencias"], 1):
                occ_fila_id, occ_texto, pdf_ini_est, pdf_fin_est, occ_rango, occ_n_paginas = occ
                if pdf_ini_est is not None:
                    est_str = f"  (pag.PDF estimada: {pdf_ini_est}–{pdf_fin_est}, {occ_n_paginas} pags.)"
                else:
                    est_str = f"  ({occ_n_paginas} pags.)"
                marca = "  ← PRIMERA" if i == 1 else f"  ← REPETICION {i-1}: necesita override exacto"
                lineas.append(f"         Fila {occ_fila_id}: '{occ_rango}'{est_str}{marca}")
        lineas.append("")
        lineas.append("  ACCION RECOMENDADA:")
        lineas.append("  1. Localiza en el PDF la página donde empieza la segunda copia física.")
        lineas.append("  2. Ve a 'Override de páginas exacto por fila' (paso 2d).")
        lineas.append("  3. Agrega Fila, Pag. inicio y Pag. fin para cada REPETICION.")
        lineas.append("  4. La página estimada es un punto de partida; verifica en el PDF.")

    return {
        "ok":              es_correcto,
        "saltos":          saltos,
        "solapamientos":   solapamientos,
        "invalidos":       entradas_invalidas,
        "repeated_folios": folios_repetidos,
        "resumen":         "\n".join(lineas),
    }

def _pagina_abs_a_etiqueta_folio(pag_abs: int) -> str:
    folio = (pag_abs + 1) // 2
    cara  = "r" if pag_abs % 2 == 1 else "v"
    return f"{folio}{cara}"
