"""
Procesador generico para perfiles SimpleWeeklyProfile: una fila por
empleado, columnas de horas por dia, REG/OT en columnas fijas.

Flujo (el mismo que hacemos a mano cada semana):
1. Leer todas las personas y horas del archivo fuente.
2. Emparejar nombres contra el roster de la plantilla destino.
3. Llenar horas diarias + REG/OT (segun la estrategia del perfil) para los
   que matchean.
4. Reportar quien trabajo en la fuente pero NO esta en el roster (con sus
   horas exactas por dia), para que un humano decida si se agregan.
5. (Opcional, en un paso aparte) agregar los faltantes con
   engine.xlsx_writer, buscando su tarifa historica o marcandolos en
   amarillo con tarifa 0 si nunca aparecieron antes.
"""

from dataclasses import dataclass, field

from engine.matching import match_source_to_roster


def _clean_hours(v):
    if v is None or v == "-" or v == "":
        return 0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0


@dataclass
class SourceEmployee:
    name: str
    skill: str | None
    days: list        # [lun..dom]
    source_reg: float
    source_ot: float

    @property
    def total(self):
        return sum(self.days)


@dataclass
class ProcessResult:
    filled_rows: dict = field(default_factory=dict)     # dest_row -> SourceEmployee
    missing_people: list = field(default_factory=list)  # SourceEmployee no en roster
    ambiguous: list = field(default_factory=list)        # (nombre, [filas candidatas])
    position_breakdown: dict = field(default_factory=dict)  # posicion -> {"reg":x,"ot":y}


def read_source(ws, profile) -> list[SourceEmployee]:
    employees = []
    for r in range(profile.source_data_start_row, ws.max_row + 1):
        name = ws.cell(r, profile.source_name_col).value
        if not name or not str(name).strip():
            continue
        skill = ws.cell(r, profile.source_skill_col).value if profile.source_skill_col else None
        days = [_clean_hours(ws.cell(r, c).value) for c in profile.source_day_cols]
        reg = _clean_hours(ws.cell(r, profile.source_reg_col).value) if profile.source_reg_col else None
        ot = _clean_hours(ws.cell(r, profile.source_ot_col).value) if profile.source_ot_col else None
        if reg is None or ot is None:
            # sin REG/OT propio en la fuente -> se calculara con weekly_40_threshold
            total = sum(days)
            reg = min(total, 40)
            ot = max(total - 40, 0)
        employees.append(SourceEmployee(name=str(name).strip(), skill=skill, days=days, source_reg=reg, source_ot=ot))
    return employees


def read_roster(ws, profile) -> dict:
    roster = {}
    for r in range(profile.dest_data_start_row, profile.dest_data_end_row + 1):
        name = ws.cell(r, profile.dest_name_col).value
        if name and str(name).strip():
            roster[r] = str(name).strip()
    return roster


def process(source_ws, dest_ws, profile) -> ProcessResult:
    employees = read_source(source_ws, profile)
    # solo nos importa gente que realmente trabajo horas esa semana; alguien
    # listado en la fuente con 0 horas (dias en blanco o "-") no trabajo y no
    # debe reportarse como "falta agregarlo"
    employees = [e for e in employees if e.total > 0]
    roster = read_roster(dest_ws, profile)
    match = match_source_to_roster([e.name for e in employees], roster)

    result = ProcessResult()
    by_name = {e.name: e for e in employees}

    for name, row in match.matches.items():
        emp = by_name[name]
        result.filled_rows[row] = emp

    for name in match.unmatched_source:
        result.missing_people.append(by_name[name])

    result.ambiguous = match.ambiguous

    # desglose por posicion (roster + faltantes), aplicando merge_positions
    breakdown = {}
    def _bucket(pos):
        pos = (pos or "SIN CLASIFICAR").strip().upper()
        return profile.merge_positions.get(pos, pos)

    for emp in list(by_name.values()):
        if emp not in result.missing_people and emp.name not in match.matches:
            continue
        bucket = _bucket(emp.skill)
        d = breakdown.setdefault(bucket, {"reg": 0.0, "ot": 0.0})
        d["reg"] += emp.source_reg
        d["ot"] += emp.source_ot
    result.position_breakdown = breakdown

    return result


def write_result(dest_ws, profile, result: ProcessResult):
    for row, emp in result.filled_rows.items():
        for i, col in enumerate(profile.dest_day_cols):
            v = emp.days[i]
            dest_ws.cell(row, col).value = v if v else None
        dest_ws.cell(row, profile.dest_reg_col).value = emp.source_reg
        dest_ws.cell(row, profile.dest_ot_col).value = emp.source_ot


def add_missing_employees(dest_ws, profile, missing_people: list,
                           rate_store: dict | None = None, global_store: dict | None = None) -> list:
    """
    Agrega al roster a quienes trabajaron pero no estaban en la plantilla:
    inserta una fila nueva por persona (preservando formulas/formato) y
    llena su nombre, posicion y horas.

    - INTAL_RATE (lo que se le paga a la persona) se busca primero en la
      memoria de ESTE proyecto y, si no esta ahi, en la memoria GLOBAL de
      trabajadores (pudo haber trabajado antes en otro proyecto distinto
      -- su pago es suyo, no depende del proyecto). Pero solo cuenta si
      esta VIGENTE (la vimos hace 90 dias o menos) -- si hace mas de 3
      meses que no aparece, no le ponemos su tarifa vieja directo, pudo
      cambiar.
    - RATE (lo que se le cobra al cliente) SOLO se completa si esta
      persona ya aparecio VIGENTE en este mismo proyecto -- nunca se copia
      el rate de cobro de otro proyecto. Si no esta vigente pero sabemos
      su posicion, se le pone la tarifa TIPICA que hoy se cobra por esa
      posicion en este proyecto (la de la gente activa), no la suya vieja.
    - Se resalta en amarillo cuando no hay un INTAL_RATE vigente -- ya sea
      porque nunca lo supimos, o porque hace mas de 3 meses que no
      aparece.

    Devuelve la lista de (row, SourceEmployee, tiene_intal_rate_vigente).
    """
    from engine.xlsx_writer import insert_rows_preserving_formulas, extend_sum_ranges, highlight_row_yellow, rebase_formula_row
    from engine.rate_store import lookup, is_stale, position_standard_rate

    if not missing_people:
        return []

    count = len(missing_people)
    old_end = profile.dest_data_end_row
    insert_at = old_end + 1
    style_row = old_end if old_end >= profile.dest_data_start_row else None

    # antes de insertar (que borra los valores de la fila de referencia),
    # guardamos las formulas "propias de su fila" que trae esa fila de
    # referencia -- por ejemplo columnas de monto ($) tipo =O14*D14 -- para
    # replicarlas en cada fila nueva, adaptadas a su propio numero de fila.
    template_formulas = {}
    if style_row:
        for c in range(1, dest_ws.max_column + 1):
            v = dest_ws.cell(style_row, c).value
            if isinstance(v, str) and v.startswith("="):
                template_formulas[c] = v

    written_cols = {profile.dest_name_col, profile.dest_intal_rate_reg_col, profile.dest_rate_reg_col,
                     profile.dest_reg_col, profile.dest_ot_col, *profile.dest_day_cols}
    if profile.dest_skill_col:
        written_cols.add(profile.dest_skill_col)

    insert_rows_preserving_formulas(dest_ws, insert_at, count, style_template_row=style_row)

    added = []
    for i, emp in enumerate(missing_people):
        row = insert_at + i
        known_project = lookup(rate_store, emp.name) if rate_store else None
        known_global = lookup(global_store, emp.name) if global_store else None
        project_fresh = bool(known_project) and not is_stale(known_project)
        global_fresh = bool(known_global) and not is_stale(known_global)

        position_hint = (known_project or {}).get("position") or (known_global or {}).get("position")

        if project_fresh or global_fresh:
            # vigente en alguno de los dos lados: se le pone su tarifa real
            intal_rate = (known_project or {}).get("intal_rate") if project_fresh else None
            if not intal_rate:
                intal_rate = (known_global or {}).get("intal_rate") if global_fresh else None
            intal_rate = intal_rate or 0
            rate = (known_project or {}).get("rate", 0) if project_fresh else 0
            has_intal_rate = bool(intal_rate)
        else:
            # no vigente en ningun lado (nunca la vimos, o hace +3 meses):
            # INTAL_RATE en 0 y amarillo, pero si sabemos su posicion se le
            # pone la tarifa tipica que hoy se cobra por esa posicion aqui
            intal_rate = 0
            position_for_rate = emp.skill or position_hint
            rate = (position_standard_rate(rate_store, position_for_rate) or 0) if rate_store else 0
            has_intal_rate = False

        dest_ws.cell(row, profile.dest_name_col).value = emp.name
        if profile.dest_skill_col:
            position = emp.skill or position_hint
            dest_ws.cell(row, profile.dest_skill_col).value = position if position else None
        dest_ws.cell(row, profile.dest_intal_rate_reg_col).value = intal_rate
        dest_ws.cell(row, profile.dest_rate_reg_col).value = rate
        for j, col in enumerate(profile.dest_day_cols):
            v = emp.days[j]
            dest_ws.cell(row, col).value = v if v else None
        dest_ws.cell(row, profile.dest_reg_col).value = emp.source_reg
        dest_ws.cell(row, profile.dest_ot_col).value = emp.source_ot

        for c, formula in template_formulas.items():
            if c in written_cols:
                continue
            dest_ws.cell(row, c).value = rebase_formula_row(formula, style_row, row)

        if not has_intal_rate:
            highlight_row_yellow(dest_ws, row, profile.dest_name_col, profile.dest_ot_col)
        added.append((row, emp, has_intal_rate))

    new_end = insert_at + count - 1
    if profile.dest_totals_row is not None:
        new_totals_row = profile.dest_totals_row + count
        extend_sum_ranges(dest_ws, new_totals_row, old_end, new_end)
        profile.dest_totals_row = new_totals_row

    profile.dest_data_end_row = new_end
    return added


def sync_rate_store(dest_ws, profile, filled_rows: dict, added: list, rate_store: dict, global_store: dict,
                     reference_date: str | None = None):
    """
    Despues de escribir el archivo, guarda lo que quedo en cada fila:
    - en la memoria del PROYECTO (posicion + RATE de cobro + INTAL_RATE de
      respaldo), para gente que ya estaba en el roster o se acaba de
      agregar a este proyecto.
    - en la memoria GLOBAL (solo INTAL_RATE + posicion), para que la
      proxima vez que esta persona aparezca en CUALQUIER proyecto ya se
      sepa cuanto se le paga.

    La gente que YA estaba en el roster (filled_rows) siempre refresca su
    fecha de "vista por ultima vez" -- trabajo esta semana con datos
    reales. La gente que se acaba de AGREGAR solo refresca esa fecha si su
    tarifa vino confirmada (de la memoria, no adivinada por posicion) --
    una tarifa adivinada nunca debe quedar guardada como si fuera vigente.
    """
    from datetime import date as _date
    from engine.rate_store import record, record_global

    ref = reference_date or _date.today().isoformat()

    def _num(v):
        try:
            return float(v)
        except (TypeError, ValueError):
            return None

    def _pos(row):
        v = dest_ws.cell(row, profile.dest_skill_col).value if profile.dest_skill_col else None
        return str(v).strip() if v else None

    for row, emp in filled_rows.items():
        position = _pos(row)
        intal_rate = _num(dest_ws.cell(row, profile.dest_intal_rate_reg_col).value)
        rate = _num(dest_ws.cell(row, profile.dest_rate_reg_col).value)
        record(rate_store, emp.name, position=position, intal_rate=intal_rate, rate=rate, last_seen=ref)
        record_global(global_store, emp.name, position=position, intal_rate=intal_rate, last_seen=ref)

    for row, emp, confirmed in added:
        position = _pos(row)
        if confirmed:
            intal_rate = _num(dest_ws.cell(row, profile.dest_intal_rate_reg_col).value)
            rate = _num(dest_ws.cell(row, profile.dest_rate_reg_col).value)
            record(rate_store, emp.name, position=position, intal_rate=intal_rate, rate=rate, last_seen=ref)
            record_global(global_store, emp.name, position=position, intal_rate=intal_rate, last_seen=ref)
        elif position:
            # tarifa adivinada (0 + rate tipico de la posicion): NUNCA se
            # guarda como si fuera confirmada -- solo la posicion, sin fecha.
            record(rate_store, emp.name, position=position)
            record_global(global_store, emp.name, position=position)
