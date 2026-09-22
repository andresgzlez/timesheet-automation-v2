"""
Perfil para proyectos "multi-codigo" (una fila por empleado + codigo de
costo, ej. HCA Palms West). A diferencia de los formatos "PH" estandar, la
plantilla destino NO trae el codigo de costo escrito en ninguna columna --
solo el nombre repetido una vez por cada codigo que esa persona trabaja.
Por eso el mapeo (fila destino -> (nombre, codigo)) es informacion propia
del proyecto que se define UNA VEZ (parte del "formato nuevo" que necesita
definirse a mano la primera vez, igual que se hizo aqui con HCA), y luego
se reutiliza automaticamente cada semana mientras el roster de
personas+codigos no cambie. Si aparece un codigo o persona que no esta en
el mapeo, se reporta como "falta definir" en vez de adivinar.
"""

from dataclasses import dataclass, field


@dataclass
class MultiCodeProfile:
    project_name: str
    source_sheet: str
    dest_timesheet_sheet: str
    dest_cobrar_sheet: str | None

    # columnas de la fuente (una fila por persona+fecha+codigo)
    source_name_col: int
    source_code_col: int
    source_date_col: int
    source_hours_col: int
    source_rt_col: int
    source_ot_col: int
    source_data_start_row: int
    week_dates: list          # ["MM/DD/YYYY", ...] lun..dom, en orden

    # columnas de la plantilla TIMESHEET
    ts_name_col: int
    ts_day_cols: list          # lun..dom
    ts_reg_col: int
    ts_ot_col: int
    ts_rate_col: int           # tarifa de facturacion (para COBRAR)
    ts_data_start_row: int
    ts_data_end_row: int

    # el mapeo fila->(nombre_fuente, codigo) definido una vez para este
    # proyecto (ver docstring del modulo)
    row_to_source_key: dict = field(default_factory=dict)   # {fila_destino: (nombre_fuente, codigo)}
