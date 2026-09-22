"""
Procesador para el formato "Project Daisy": una fila por empleado + fecha +
codigo de costo (no una fila por empleado con columnas de dia). El mismo
empleado puede tener varias filas el mismo dia si trabajo mas de un codigo.

Reglas (de /areas/timesheet-automation.md):
- Sumar las horas de cada dia en las columnas diarias del destino.
- REG/OT del destino = suma exacta de RT Hours / OT Hours de la fuente por
  semana (no recalcular con la formula de >40h del destino).
- Codigos de clasificacion terminados en "2" (INTALMECH2, INTALLABOR2,
  INTALSPEC2, ...) son turno de noche y se acumulan aparte de los codigos
  base ("1"), como categoria distinta en el bloque de totales por posicion
  al final de la plantilla (MECHANIC vs MECHANIC - NIGHT SHIFT, etc.) -- no
  como una fila de empleado nueva en el roster principal.
- Las horas DT (doble tiempo) se suman dentro de OT, ya que la plantilla
  destino no tiene una columna propia para DT (asuncion a confirmar la
  primera vez que aparezca un caso real).
"""

from dataclasses import dataclass, field
from collections import defaultdict

DATE_ORDER = None  # se arma dinamicamente a partir de las fechas presentes


def _is_night_shift(classification_code: str) -> bool:
    return bool(classification_code) and str(classification_code).strip().endswith("2")


@dataclass
class DaisyEmployeeWeek:
    name: str
    shift: str            # "DAY" | "NIGHT"
    days: list             # [lun..dom]
    reg: float
    ot: float

    @property
    def total(self):
        return sum(self.days)


def read_daisy_source(ws, name_col=1, code_col=6, date_col=14, rt_col=17, ot_col=18, dt_col=19,
                       data_start_row=12) -> list[DaisyEmployeeWeek]:
    # 1. Recolectar todas las fechas presentes, para mapearlas a lun..dom en orden
    raw_rows = []
    dates_seen = set()
    for r in range(data_start_row, ws.max_row + 1):
        name = ws.cell(r, name_col).value
        if not name or not str(name).strip():
            continue
        date = ws.cell(r, date_col).value
        code = ws.cell(r, code_col).value
        rt = ws.cell(r, rt_col).value or 0
        ot = ws.cell(r, ot_col).value or 0
        dt = ws.cell(r, dt_col).value or 0
        raw_rows.append((str(name).strip(), code, date, float(rt), float(ot), float(dt)))
        if date:
            dates_seen.add(str(date))

    week_dates = sorted(dates_seen)  # asume formato MM/DD/YYYY, ordena bien como string solo si es consistente
    # ordenar de verdad por fecha real:
    import datetime
    def parse(d):
        return datetime.datetime.strptime(d, "%m/%d/%Y")
    week_dates = sorted(dates_seen, key=parse)
    date_to_col = {d: i for i, d in enumerate(week_dates[:7])}

    # 2. Acumular por (nombre, turno)
    acc = defaultdict(lambda: {"days": [0.0] * 7, "reg": 0.0, "ot": 0.0})
    for name, code, date, rt, ot, dt in raw_rows:
        shift = "NIGHT" if _is_night_shift(code) else "DAY"
        key = (name, shift)
        if date and str(date) in date_to_col:
            acc[key]["days"][date_to_col[str(date)]] += rt + ot + dt
        acc[key]["reg"] += rt
        acc[key]["ot"] += ot + dt   # DT se suma dentro de OT (ver docstring)

    return [
        DaisyEmployeeWeek(name=name, shift=shift, days=v["days"], reg=v["reg"], ot=v["ot"])
        for (name, shift), v in acc.items()
    ]


def write_daisy_result(dest_ws, employees: list[DaisyEmployeeWeek], matches: dict,
                        day_cols=range(8, 15), reg_col=15, ot_col=16):
    """
    `matches`: {source_name: dest_row} para los empleados de turno DIA (los
    unicos que este metodo escribe -- turno noche no tiene fila propia en
    el roster principal, va a la categoria "- NIGHT SHIFT" del bloque de
    totales al final de la plantilla, que se deja fuera por ahora al no
    haber un caso real con el que validarlo).
    """
    by_name = {e.name: e for e in employees if e.shift == "DAY"}
    written = 0
    for name, row in matches.items():
        emp = by_name.get(name)
        if not emp:
            continue
        for i, col in enumerate(day_cols):
            v = emp.days[i]
            dest_ws.cell(row, col).value = v if v else None
        dest_ws.cell(row, reg_col).value = emp.reg
        dest_ws.cell(row, ot_col).value = emp.ot
        written += 1
    return written
