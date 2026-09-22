"""
Procesador para proyectos "multi-codigo" tipo HCA Palms West: una fila
fuente por persona+fecha+codigo de costo, y una fila destino por cada
combinacion persona+codigo (fija, definida en el perfil -- ver
profiles/schema_multicode.py).

Reutiliza directamente:
- engine.regot.cumulative_row_order para el reparto de OT cuando una
  persona tiene varios codigos (regla confirmada con el caso Dilman
  Garcia).
- engine.xlsx_writer para el resaltado de empleados nuevos.

Tambien arma la pestana COBRAR (facturacion), agrupada por
codigo+posicion/tarifa, detectando los "slots" de fila igual que se hizo a
mano: busca las filas que ya traen (CODIGO, POSICION) en la plantilla y
llena hacia abajo desde ahi.
"""

import collections

from engine.regot import DayHours, cumulative_row_order


def read_source(ws, profile) -> dict:
    """Devuelve {(nombre_fuente, codigo): {"days":[7], "rt":x, "ot":y}}"""
    date_to_col = {d: i for i, d in enumerate(profile.week_dates)}
    recs = collections.defaultdict(lambda: {"days": [0.0] * 7, "rt": 0.0, "ot": 0.0})
    for r in range(profile.source_data_start_row, ws.max_row + 1):
        name = ws.cell(r, profile.source_name_col).value
        if not name:
            continue
        code = ws.cell(r, profile.source_code_col).value
        date = ws.cell(r, profile.source_date_col).value
        hrs = ws.cell(r, profile.source_hours_col).value or 0
        rt = ws.cell(r, profile.source_rt_col).value or 0
        ot = ws.cell(r, profile.source_ot_col).value or 0
        key = (str(name).strip(), code)
        if date in date_to_col:
            recs[key]["days"][date_to_col[date]] = hrs
        recs[key]["rt"] += rt
        recs[key]["ot"] += ot
    return dict(recs)


def compute_regot(profile, recs: dict) -> dict:
    """
    Devuelve {fila_destino: (reg, ot)} usando la regla confirmada:
    - Si la persona tiene un solo codigo -> se usa el REG/OT exacto que
      trae la fuente para ese codigo (puede tener OT diario propio).
    - Si tiene varios codigos -> se reparte con el umbral de 40h acumulado,
      EN EL ORDEN en que las filas de esa persona aparecen en la plantilla
      TIMESHEET (profile.row_to_source_key, ordenado por numero de fila).
    """
    by_name = collections.defaultdict(list)
    for row, key in sorted(profile.row_to_source_key.items()):
        name = key[0]
        by_name[name].append((row, key))

    result = {}
    for name, rows in by_name.items():
        entries = []
        for row, key in rows:
            d = recs.get(key)
            if d is None:
                continue
            entries.append(DayHours(row_ref=row, days=d["days"], source_rt=d["rt"], source_ot=d["ot"]))
        if not entries:
            continue
        result.update(cumulative_row_order(entries))
    return result


def write_timesheet(ws, profile, recs: dict, regot: dict):
    written = 0
    for row, key in profile.row_to_source_key.items():
        d = recs.get(key)
        if d is None or row not in regot:
            continue
        for i, col in enumerate(profile.ts_day_cols):
            v = d["days"][i]
            ws.cell(row, col).value = v if v else None
        reg, ot = regot[row]
        ws.cell(row, profile.ts_reg_col).value = reg
        ws.cell(row, profile.ts_ot_col).value = ot
        written += 1
    return written


# --------------------------------------------------------------------------
# COBRAR (facturacion): agrupa por (codigo, posicion/tarifa) usando los
# "slots" que ya trae la plantilla (misma tecnica usada a mano con HCA).
# --------------------------------------------------------------------------

def detect_cobrar_slots(cobrar_ws, code_col=6, pos_col=7, header_row=3):
    """Escanea las filas que ya traen (CODIGO, POSICION) en la plantilla
    COBRAR y calcula el rango de filas de datos que le corresponde a cada
    una (desde justo despues de la etiqueta anterior hasta una fila antes
    de esta)."""
    label_rows = []
    for r in range(header_row + 1, cobrar_ws.max_row + 1):
        code = cobrar_ws.cell(r, code_col).value
        pos = cobrar_ws.cell(r, pos_col).value
        if code is None or pos is None:
            continue
        code_s = str(code).strip()
        if code_s.upper() == "TOTAL":
            break
        label_rows.append((r, code_s, str(pos).strip()))

    slots = []
    prev = header_row
    for (r, code, pos) in label_rows:
        slots.append({"label_row": r, "code": code, "pos_label": pos, "start": prev + 1, "end": r - 1})
        prev = r
    return slots


def build_cobrar_groups(ws_timesheet, profile, regot: dict, tier_fn):
    """
    tier_fn(position_text, billing_rate) -> etiqueta de grupo (ej. separa
    HELPER en 2 niveles de tarifa como en HCA). Devuelve
    {(codigo, etiqueta): [(fila_ts, nombre, tarifa, reg, ot), ...]}
    """
    groups = collections.defaultdict(list)
    for row, (name, code) in profile.row_to_source_key.items():
        if row not in regot:
            continue
        pos = ws_timesheet.cell(row, 3).value or ""
        rate = ws_timesheet.cell(row, profile.ts_rate_col).value or 0
        reg, ot = regot[row]
        label = tier_fn(pos, rate)
        groups[(code.strip(), label)].append((row, name, rate, reg, ot))
    return groups


def write_cobrar(cobrar_ws, slots, groups):
    """Llena cada slot con sus empleados y escribe la formula SUM en la fila
    de etiqueta. Lanza ValueError si un grupo no cabe en su slot (mas
    personas de las que la plantilla reservo esa semana -- necesita
    ampliarse a mano con engine.xlsx_writer antes de continuar)."""
    for s in slots:
        key = (s["code"], s["pos_label"])
        emps = groups.get(key, [])
        start, end, label_row = s["start"], s["end"], s["label_row"]
        if len(emps) > (end - start + 1):
            raise ValueError(
                f"El grupo {key} tiene {len(emps)} personas pero la plantilla solo "
                f"reservo {end - start + 1} filas -- hay que insertar filas antes de continuar."
            )
        for i, (row, name, rate, reg, ot) in enumerate(emps):
            r = start + i
            cobrar_ws.cell(r, 2).value = reg
            cobrar_ws.cell(r, 3).value = ot
            cobrar_ws.cell(r, 4).value = round(reg * rate, 2)
            cobrar_ws.cell(r, 5).value = round(ot * rate * 1.5, 2)
        cobrar_ws.cell(label_row, 2).value = f"=SUM(B{start}:B{end})"
        cobrar_ws.cell(label_row, 3).value = f"=SUM(C{start}:C{end})"
        cobrar_ws.cell(label_row, 4).value = f"=SUM(D{start}:D{end})"
        cobrar_ws.cell(label_row, 5).value = f"=SUM(E{start}:E{end})"
