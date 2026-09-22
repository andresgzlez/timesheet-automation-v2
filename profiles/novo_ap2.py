from profiles.schema import SimpleWeeklyProfile

# NOVO AP2 tiene el layout de columnas corrido una posicion a la izquierda
# respecto a los demas proyectos JJ/NIPRO -- confirmado semana a semana, por
# eso este perfil se define aparte en vez de reusar los offsets estandar.
PROFILE = SimpleWeeklyProfile(
    project_name="NOVO AP2",
    source_sheet="Sheet1",
    source_name_col=2,
    source_skill_col=7,
    source_day_cols=[8, 9, 10, 11, 12, 13, 14],
    source_reg_col=15,
    source_ot_col=16,
    source_data_start_row=8,
    dest_sheet="PH",
    dest_name_col=2,
    dest_skill_col=3,
    dest_intal_rate_reg_col=4,
    dest_rate_reg_col=6,
    dest_day_cols=[8, 9, 10, 11, 12, 13, 14],
    dest_reg_col=15,
    dest_ot_col=16,
    dest_data_start_row=11,
    dest_data_end_row=29,
    dest_totals_row=30,
    reg_ot_strategy="exact_source",
    merge_positions={"FINISHER": "MECHANIC"},
)
