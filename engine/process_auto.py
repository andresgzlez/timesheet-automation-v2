"""
Punto de entrada para la familia "estandar" de formatos (JJ Production
Upfit, JJ Warehouse Upfit, JJ Exterior, NIPRO, NOVO AP2, y cualquier otro
proyecto nuevo -- incluidos los de una sola persona -- que comparta esa
misma estructura de encabezados).

No requiere un perfil escrito a mano: detecta columnas y filas solo
(engine.autodetect) y arma un SimpleWeeklyProfile al vuelo. Asi, un
proyecto nuevo de esta familia funciona en la web app sin que nadie tenga
que configurar nada -- solo aplica cuando el formato es el conocido; si
la deteccion falla (formato realmente distinto), avisa en vez de adivinar.
"""

from profiles.schema import SimpleWeeklyProfile
from engine.autodetect import detect_source_layout, detect_dest_layout, detect_roster_range

DEFAULT_MERGE_POSITIONS = {"FINISHER": "MECHANIC"}


def build_profile_auto(project_name: str, source_ws, dest_ws, dest_sheet_name: str,
                        source_sheet_name: str = "Sheet1",
                        merge_positions: dict | None = None) -> SimpleWeeklyProfile:
    s = detect_source_layout(source_ws)
    d = detect_dest_layout(dest_ws)
    d_start, d_end, d_totals = detect_roster_range(dest_ws, d.name_col, d.data_start_row)

    if d_totals is None:
        raise ValueError(
            "No se encontro la fila de totales (formula =SUM(...)) despues del roster -- "
            "revisa si la plantilla tiene el formato esperado antes de continuar."
        )

    return SimpleWeeklyProfile(
        project_name=project_name,
        source_sheet=source_sheet_name,
        source_name_col=s.name_col,
        source_skill_col=s.skill_col,
        source_day_cols=s.day_cols,
        source_reg_col=s.reg_col,
        source_ot_col=s.ot_col,
        source_data_start_row=s.data_start_row,
        dest_sheet=dest_sheet_name,
        dest_name_col=d.name_col,
        dest_skill_col=d.skill_col,
        dest_intal_rate_reg_col=d.name_col + 2,   # convencion: INTAL_RATE 2 columnas despues de NAME
        dest_rate_reg_col=d.name_col + 4,         # convencion: RATE 4 columnas despues de NAME
        dest_day_cols=d.day_cols,
        dest_reg_col=d.reg_col,
        dest_ot_col=d.ot_col,
        dest_data_start_row=d_start,
        dest_data_end_row=d_end,
        dest_totals_row=d_totals,
        reg_ot_strategy="exact_source",
        merge_positions=merge_positions or DEFAULT_MERGE_POSITIONS,
    )
