"""
Deteccion automatica de layout para la familia de formatos "estandar"
(JJ Production Upfit, JJ Warehouse Upfit, JJ Exterior, NIPRO, NOVO AP2, y
cualquier otro proyecto -- incluyendo los de una sola persona que hoy se
hacen manual -- que comparta esa misma estructura de encabezados).

En vez de escribir un perfil a mano por cada proyecto nuevo (row ranges,
columnas), esto busca las etiquetas de encabezado ("NAME", "SKILL", "MON",
"REG.", etc.) donde sea que esten, y deriva la configuracion sola. Solo
hace falta escribir un perfil a mano cuando el formato es realmente
distinto (multi-codigo, multi-fila por dia, como Project Daisy o HCA).
"""

import re
from dataclasses import dataclass

DAY_LABELS = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]


def _norm(v) -> str:
    if v is None:
        return ""
    return re.sub(r"[^A-Z]", "", str(v).upper())


def _find_all(ws, target: str, max_row=25, max_col=30) -> list[tuple[int, int]]:
    """Todas las celdas cuyo texto normalizado empiece con `target`."""
    hits = []
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            if _norm(ws.cell(r, c).value).startswith(target):
                hits.append((r, c))
    return hits


def _find_first(ws, target: str, **kw):
    hits = _find_all(ws, target, **kw)
    return hits[0] if hits else None


@dataclass
class DetectedLayout:
    name_col: int
    skill_col: int | None
    day_cols: list          # 7 columnas, lun..dom
    reg_col: int
    ot_col: int
    header_last_row: int    # ultima fila de encabezado
    data_start_row: int


def _best_name_column(ws, guess_col: int, data_start_row: int, span: int = 3) -> int:
    """
    La etiqueta de encabezado 'NAME' a veces no cae exactamente sobre la
    columna donde realmente estan escritos los nombres (encabezados
    corridos, columnas de numero/# de por medio). Para no depender de eso,
    se elige -- entre las columnas cercanas al hit de texto -- la que tiene
    mas celdas con texto tipo nombre (varias palabras, sin ser puramente
    numerico) en las filas de datos.
    """
    best_col, best_score = guess_col, -1
    for c in range(max(1, guess_col - span), guess_col + span + 1):
        score = 0
        for r in range(data_start_row, data_start_row + 25):
            v = ws.cell(r, c).value
            if isinstance(v, str) and len(v.strip()) > 3 and re.search(r"[A-Za-z].*[A-Za-z]", v):
                score += 1
        if score > best_score:
            best_col, best_score = c, score
    return best_col


def detect_source_layout(ws) -> DetectedLayout:
    name_hit = _find_first(ws, "NAME")
    if not name_hit:
        raise ValueError("No se encontro la columna NAME en el archivo fuente")
    name_row, name_col = name_hit
    name_col = _best_name_column(ws, name_col, name_row + 1)

    skill_hit = _find_first(ws, "SKILL")
    skill_col = skill_hit[1] if skill_hit else None

    day_cols = []
    day_row = None
    for label in DAY_LABELS:
        hit = _find_first(ws, label)
        if hit:
            day_row = hit[0]
            day_cols.append(hit[1])
    if len(day_cols) < 5:
        raise ValueError(f"Solo se detectaron {len(day_cols)} columnas de dias en la fuente")

    reg_hit = _find_first(ws, "REG")
    ot_hit = _find_first(ws, "OT") or _find_first(ws, "O/T")
    reg_col = reg_hit[1] if reg_hit else day_cols[-1] + 1
    ot_col = ot_hit[1] if ot_hit else reg_col + 1

    header_last_row = max(name_row, day_row or name_row)
    return DetectedLayout(
        name_col=name_col, skill_col=skill_col, day_cols=day_cols,
        reg_col=reg_col, ot_col=ot_col,
        header_last_row=header_last_row, data_start_row=header_last_row + 1,
    )


def detect_dest_layout(ws) -> DetectedLayout:
    name_hit = _find_first(ws, "NAME")
    if not name_hit:
        raise ValueError("No se encontro la columna NAME en la plantilla destino")
    name_row, name_col = name_hit
    skill_col = name_col + 1  # convencion observada en todos los proyectos "PH"

    day_cols = []
    day_row = None
    for label in DAY_LABELS[:3]:  # MON/TUE/WED alcanzan para ubicar la fila; luego se completa por posicion
        hit = _find_first(ws, label)
        if hit:
            day_row = hit[0]
            day_cols.append(hit[1])
    if not day_cols:
        raise ValueError("No se encontraron columnas de dias en la plantilla destino")
    start_day_col = min(day_cols)
    day_cols = list(range(start_day_col, start_day_col + 7))

    total_hit = _find_first(ws, "TOTALHRS")
    if total_hit:
        reg_col = total_hit[1]
    else:
        reg_col = day_cols[-1] + 1
    ot_col = reg_col + 1

    header_last_row = max(name_row, day_row or name_row) + 1  # sub-fila REG/OT debajo del encabezado
    return DetectedLayout(
        name_col=name_col, skill_col=skill_col, day_cols=day_cols,
        reg_col=reg_col, ot_col=ot_col,
        header_last_row=header_last_row, data_start_row=header_last_row + 1,
    )


def detect_roster_range(ws, name_col: int, data_start_row: int) -> tuple[int, int, int | None]:
    """Devuelve (primera_fila, ultima_fila, fila_de_totales) escaneando hacia
    abajo desde data_start_row hasta encontrar una fila sin nombre."""
    r = data_start_row
    last_with_name = None
    while r < ws.max_row + 5:
        name = ws.cell(r, name_col).value
        if name and str(name).strip():
            last_with_name = r
            r += 1
        else:
            break
    if last_with_name is None:
        raise ValueError("No se detecto ningun empleado en el roster destino")

    totals_row = None
    for candidate in range(last_with_name + 1, min(last_with_name + 4, ws.max_row + 1)):
        for c in range(name_col, name_col + 15):
            v = ws.cell(candidate, c).value
            if isinstance(v, str) and v.strip().startswith("=SUM("):
                totals_row = candidate
                break
        if totals_row:
            break

    return data_start_row, last_with_name, totals_row
