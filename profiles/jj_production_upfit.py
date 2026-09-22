from profiles.schema import SimpleWeeklyProfile

PROFILE = SimpleWeeklyProfile(
    project_name="JJ Production Upfit",
    source_sheet="Sheet1",
    source_name_col=3,          # C
    source_skill_col=8,         # H
    source_day_cols=[9, 10, 11, 12, 13, 14, 15],   # I..O (Mon..Sun)
    source_reg_col=16,          # P
    source_ot_col=17,           # Q
    source_data_start_row=9,
    dest_sheet="PH",
    dest_name_col=2,            # B
    dest_skill_col=3,           # C
    dest_intal_rate_reg_col=4,  # D
    dest_rate_reg_col=6,        # F
    dest_day_cols=[8, 9, 10, 11, 12, 13, 14],       # H..N (Mon..Sun)
    dest_reg_col=15,            # O
    dest_ot_col=16,             # P
    dest_data_start_row=10,
    dest_data_end_row=29,
    dest_totals_row=30,
    reg_ot_strategy="exact_source",
    merge_positions={"FINISHER": "MECHANIC"},
)
