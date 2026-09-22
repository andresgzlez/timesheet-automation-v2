import json
from pathlib import Path

from profiles.schema_multicode import MultiCodeProfile

# Mapeo fila TIMESHEET -> (nombre_fuente, codigo de costo), definido una vez
# para este proyecto (ver docstring de engine/process_hca.py). Se genero a
# partir del matching ya confirmado con el cliente para HCA Palms West, y
# queda guardado en disco como parte del perfil (no en un archivo temporal).
_MAPPING_PATH = Path(__file__).parent / "data" / "hca_palms_west_mapping.json"
with open(_MAPPING_PATH) as f:
    _raw = json.load(f)
ROW_TO_SOURCE_KEY = {int(row): tuple(key) for row, key in _raw.items()}

PROFILE = MultiCodeProfile(
    project_name="HCA Palms West Tower Expansion",
    source_sheet="Project Payroll Detailed",
    dest_timesheet_sheet="TIMESHEET",
    dest_cobrar_sheet="COBRAR",
    source_name_col=1,
    source_code_col=6,
    source_date_col=5,
    source_hours_col=8,
    source_rt_col=9,
    source_ot_col=10,
    source_data_start_row=10,
    week_dates=["09/14/2026", "09/15/2026", "09/16/2026", "09/17/2026", "09/18/2026", "09/19/2026", "09/20/2026"],
    ts_name_col=2,
    ts_day_cols=[8, 9, 10, 11, 12, 13, 14],
    ts_reg_col=15,
    ts_ot_col=16,
    ts_rate_col=6,
    ts_data_start_row=11,
    ts_data_end_row=50,
    row_to_source_key=ROW_TO_SOURCE_KEY,
)


def tier_fn(position_text: str, billing_rate: float) -> str:
    """HELPER se separa en 2 niveles de tarifa de facturacion (confirmado
    con el cliente): HELPER (tarifa baja, ~$23.60) vs HELPER #2 (~$29)."""
    pos = (position_text or "").strip()
    if pos == "HELPER":
        return "HELPER" if billing_rate < 25 else "HELPER #2"
    return pos
