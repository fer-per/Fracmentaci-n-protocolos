import re
from typing import List, Optional
from modules.dominio.modelos import RegistroDocumento

# --- Patrones de Validación Tópica ---
_VALID_TOPICA_RE = re.compile(
    r'^[A-Za-záéíóúÁÉÍÓÚüÜñÑ][A-Za-záéíóúÁÉÍÓÚüÜñÑ0-9\s,.()\-\']+$'
)
_ONLY_DIGITS_RE = re.compile(r'^\d+$')

def _clasificar_topica(valor: str) -> str:
    v = str(valor).strip() if valor else ""
    if not v or v.lower() == "nan":
        return "vacio"
    if _ONLY_DIGITS_RE.match(v):
        return "invalido"
    if not _VALID_TOPICA_RE.match(v):
        return "invalido"
    return "ok"

def analizar_data_topica(registros: List[RegistroDocumento]) -> dict:
    lista_vacios    = []
    lista_invalidos = []
    cantidad_ok     = 0
    total           = 0

    for idx, registro in enumerate(registros):
        valor = registro.lugar
        reg_id = registro.registro_id or str(idx + 1)
        estado = _clasificar_topica(valor)
        total += 1

        if estado == "vacio":
            lista_vacios.append({
                "indice": idx,
                "reg_id": reg_id,
                "valor": str(valor).strip(),
                "detalle": f"Reg. {reg_id}: DATA TÓPICA vacía o nula",
            })
        elif estado == "invalido":
            lista_invalidos.append({
                "indice": idx,
                "reg_id": reg_id,
                "valor": str(valor).strip(),
                "detalle": f"Reg. {reg_id}: Formato inválido -> '{str(valor).strip()}'",
            })
        else:
            cantidad_ok += 1

    ok = not lista_vacios and not lista_invalidos

    lineas = []
    if ok:
        lineas.append(f"[OK]  DATA TOPICA correcta en los {total} registros.")
    else:
        lineas.append(f"[>>]  Total registros analizados: {total}")
        lineas.append(f"[OK]  Con DATA TOPICA válida:     {cantidad_ok}")
        if lista_vacios:
            lineas.append(f"[!!]  Vacíos ({len(lista_vacios)}):")
            for e in lista_vacios[:20]:
                lineas.append(f"    . {e['detalle']}")
            if len(lista_vacios) > 20:
                lineas.append(f"    ... y {len(lista_vacios) - 20} más")
        if lista_invalidos:
            lineas.append(f"[ERR] Formato inválido ({len(lista_invalidos)}):")
            for e in lista_invalidos[:20]:
                lineas.append(f"    . {e['detalle']}")
            if len(lista_invalidos) > 20:
                lineas.append(f"    ... y {len(lista_invalidos) - 20} más")

    return {
        "ok":         ok,
        "vacios":     lista_vacios,
        "invalidos":  lista_invalidos,
        "resumen":    "\n".join(lineas),
        "total":      total,
        "cantidad_ok": cantidad_ok,
    }


# --- Patrones de Validación Crónica ---
_DATE_RE = re.compile(r'^(\d{1,2})/(\d{1,2})/(\d{4})$')

def _parsear_fecha(valor: str) -> Optional[tuple]:
    v = str(valor).strip() if valor else ""
    if not v or v.lower() == "nan":
        return None
    m = _DATE_RE.match(v)
    if not m:
        return None
    dia, mes, anio = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if not (1 <= mes <= 12 and 1 <= dia <= 31 and anio >= 1000):
        return None
    return (anio, mes, dia)

def _tupla_a_cadena(t: tuple) -> str:
    return f"{t[2]}/{t[1]}/{t[0]}"

def analizar_data_cronica(registros: List[RegistroDocumento]) -> dict:
    formato_invalido   = []
    incoherentes       = []
    sin_progresion     = []
    cantidad_ok        = 0
    total              = 0

    prev_fecha_ini_parsed = None
    prev_reg_id           = None

    for idx, registro in enumerate(registros):
        val_ini = registro.fecha_inicio
        val_fin = registro.fecha_fin
        reg_id  = registro.registro_id or str(idx + 1)
        total  += 1

        parsed_ini = _parsear_fecha(val_ini)
        parsed_fin = _parsear_fecha(val_fin) if val_fin else None

        ini_str = str(val_ini).strip() if val_ini else ""
        fin_str = str(val_fin).strip() if val_fin else ""

        if ini_str and ini_str.lower() != "nan" and parsed_ini is None:
            formato_invalido.append({
                "indice": idx,
                "reg_id": reg_id,
                "campo": "FECHA INICIAL",
                "valor": ini_str,
                "detalle": f"Reg. {reg_id}: FECHA INICIAL con formato inválido -> '{ini_str}' (esperado d/m/yyyy)",
            })

        if fin_str and fin_str.lower() != "nan" and parsed_fin is None:
            formato_invalido.append({
                "indice": idx,
                "reg_id": reg_id,
                "campo": "FECHA FINAL",
                "valor": fin_str,
                "detalle": f"Reg. {reg_id}: FECHA FINAL con formato inválido -> '{fin_str}' (esperado d/m/yyyy)",
            })

        if parsed_ini and parsed_fin and parsed_fin < parsed_ini:
            incoherentes.append({
                "indice": idx,
                "reg_id": reg_id,
                "fecha_ini": ini_str,
                "fecha_fin": fin_str,
                "detalle": (
                    f"Reg. {reg_id}: FECHA FINAL ({fin_str}) es anterior a "
                    f"FECHA INICIAL ({ini_str})"
                ),
            })

        if parsed_ini and prev_fecha_ini_parsed is not None:
            if parsed_ini < prev_fecha_ini_parsed:
                sin_progresion.append({
                    "indice": idx,
                    "reg_id": reg_id,
                    "prev_reg_id": prev_reg_id,
                    "fecha_actual": ini_str,
                    "fecha_anterior": _tupla_a_cadena(prev_fecha_ini_parsed),
                    "detalle": (
                        f"Reg. {reg_id}: FECHA INICIAL ({ini_str}) es anterior a "
                        f"la del reg. {prev_reg_id} ({_tupla_a_cadena(prev_fecha_ini_parsed)}) "
                        f"-> regresión cronológica"
                    ),
                })

        if parsed_ini:
            prev_fecha_ini_parsed = parsed_ini
            prev_reg_id           = reg_id
            cantidad_ok += 1

    ok = not formato_invalido and not incoherentes and not sin_progresion

    lineas = []
    if ok:
        lineas.append(f"[OK]  DATA CRONICA correcta en los {total} registros.")
    else:
        lineas.append(f"[>>]  Total registros analizados: {total}")
        lineas.append(f"[OK]  Con fecha válida y progresiva: {cantidad_ok}")
        if formato_invalido:
            lineas.append(f"[ERR] Formato inválido ({len(formato_invalido)}):")
            for e in formato_invalido[:20]:
                lineas.append(f"    . {e['detalle']}")
            if len(formato_invalido) > 20:
                lineas.append(f"    ... y {len(formato_invalido) - 20} más")
        if incoherentes:
            lineas.append(f"[!!]  Fecha final anterior a fecha inicial ({len(incoherentes)}):")
            for e in incoherentes[:20]:
                lineas.append(f"    . {e['detalle']}")
            if len(incoherentes) > 20:
                lineas.append(f"    ... y {len(incoherentes) - 20} más")
        if sin_progresion:
            lineas.append(f"[!!]  Regresión cronológica ({len(sin_progresion)}):")
            for e in sin_progresion[:20]:
                lineas.append(f"    . {e['detalle']}")
            if len(sin_progresion) > 20:
                lineas.append(f"    ... y {len(sin_progresion) - 20} más")

    return {
        "ok":              ok,
        "formato_invalido": formato_invalido,
        "incoherentes":    incoherentes,
        "sin_progresion":  sin_progresion,
        "resumen":         "\n".join(lineas),
        "total":           total,
        "cantidad_ok":     cantidad_ok,
    }
