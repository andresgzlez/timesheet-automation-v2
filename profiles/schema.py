"""
Esquema de "perfil" de proyecto: la configuracion que le dice al motor como
leer un archivo fuente y una plantilla destino especificos, sin tener que
escribir codigo nuevo por proyecto.

Dos familias de perfiles, segun lo que ya vimos en los proyectos reales:

- SimpleWeeklyProfile: una fila por empleado, columnas de horas por dia,
  REG/OT en columnas fijas. Cubre JJ Production Upfit, JJ Warehouse Upfit,
  JJ Exterior, NIPRO, NOVO AP2 (con columnas desplazadas).

- MultiCodeProfile: una fila por empleado+codigo de costo (HCA Palms West),
  o una fila por empleado+codigo+dia (Project Daisy). Usa
  engine.regot.cumulative_row_order para el reparto de OT.
"""

from dataclasses import dataclass, field


@dataclass
class SimpleWeeklyProfile:
    project_name: str

    # --- Fuente ---
    source_sheet: str
    source_name_col: int
    source_skill_col: int | None
    source_day_cols: list[int]        # [lun..dom] en orden, columnas del origen
    source_reg_col: int | None        # si la fuente ya trae REG/OT propios
    source_ot_col: int | None
    source_data_start_row: int

    # --- Destino ---
    dest_sheet: str
    dest_name_col: int
    dest_skill_col: int | None
    dest_intal_rate_reg_col: int
    dest_rate_reg_col: int
    dest_day_cols: list[int]          # [lun..dom] en orden, columnas del destino
    dest_reg_col: int
    dest_ot_col: int
    dest_data_start_row: int
    dest_data_end_row: int
    dest_totals_row: int | None = None   # fila con las formulas =SUM(...) del roster
    dest_row_num_col: int | None = None  # columna "No" (numero de fila), normalmente name_col - 1

    # --- Reglas ---
    reg_ot_strategy: str = "exact_source"  # "exact_source" | "weekly_40_threshold"
    merge_positions: dict = field(default_factory=dict)  # ej. {"FINISHER": "MECHANIC"}


@dataclass
class MultiCodeProfile:
    project_name: str
    source_sheet: str
    dest_timesheet_sheet: str
    dest_cobrar_sheet: str | None = None
    # el resto de configuracion de este tipo de perfil se define caso a caso
    # dado lo especifico de cada estructura (ver HCA Palms West como referencia)
    notes: str = ""
