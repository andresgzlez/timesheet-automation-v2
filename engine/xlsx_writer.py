"""
Motor de escritura de Excel: inserta filas nuevas SIN romper formulas ni
formato, preservando estilos, y sabe resaltar en amarillo a un empleado
nuevo sin tarifa historica.

Esta es la tecnica de "insercion de filas preservando formulas" que usamos
manualmente en JJ Production Upfit, JJ Warehouse Upfit, JJ Exterior y
Project Daisy: como openpyxl no reacomoda automaticamente las referencias
de celda ni los rangos de SUMA al insertar filas, hay que:
  1. Guardar una foto de todas las celdas (valor + estilos)
  2. Desmergear todo
  3. Calcular el nuevo numero de fila de cada celda
  4. Reescribir formulas cambiando referencias de fila con regex
  5. Volver a mergear en las posiciones nuevas
  6. Extender manualmente los rangos de SUMA que cubran las filas nuevas
"""

import copy
import re

YELLOW_FILL_RGB = "FFFF00"


def _shift_row(row: int, insert_at: int, count: int) -> int:
    return row + count if row >= insert_at else row


def _rewrite_formula_refs(formula: str, insert_at: int, count: int) -> str:
    """Reescribe referencias de fila (A1, $A$1, etc.) en una formula para
    reflejar las filas insertadas. Tambien extiende el extremo final de
    rangos SUM(...) que cruzan el punto de insercion."""

    def repl(m):
        col_abs, col, row_abs, row = m.groups()
        row_num = int(row)
        new_row = _shift_row(row_num, insert_at, count)
        return f"{col_abs}{col}{row_abs}{new_row}"

    return re.sub(r"(\$?)([A-Z]{1,3})(\$?)(\d+)", repl, formula)


def insert_rows_preserving_formulas(ws, insert_at: int, count: int, style_template_row: int | None = None):
    """
    Inserta `count` filas en blanco justo antes de `insert_at`, preservando
    formulas, formatos y merges de todo el sheet.

    Si `style_template_row` se da, las nuevas filas heredan el estilo
    (fuente, relleno, borde, alineacion, formato numerico) de esa fila.

    Devuelve el rango de filas nuevas: (insert_at, insert_at + count - 1).
    """
    max_row = ws.max_row
    max_col = ws.max_column

    # 1. Foto de todas las celdas
    snapshot = {}
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(r, c)
            snapshot[(r, c)] = {
                "value": cell.value,
                "font": copy.copy(cell.font),
                "fill": copy.copy(cell.fill),
                "border": copy.copy(cell.border),
                "alignment": copy.copy(cell.alignment),
                "number_format": cell.number_format,
            }

    # 2. Guardar merges y desmergear
    merged_ranges = list(ws.merged_cells.ranges)
    for mr in merged_ranges:
        ws.unmerge_cells(str(mr))

    style_row_snapshot = None
    if style_template_row is not None:
        style_row_snapshot = {
            c: snapshot[(style_template_row, c)] for c in range(1, max_col + 1)
        }

    # 3-4. Limpiar hoja y reescribir en las posiciones nuevas
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(r, c).value = None

    for (r, c), data in snapshot.items():
        new_r = _shift_row(r, insert_at, count)
        cell = ws.cell(new_r, c)
        value = data["value"]
        if isinstance(value, str) and value.startswith("="):
            value = "=" + _rewrite_formula_refs(value[1:], insert_at, count)
        cell.value = value
        cell.font = data["font"]
        cell.fill = data["fill"]
        cell.border = data["border"]
        cell.alignment = data["alignment"]
        cell.number_format = data["number_format"]

    # 5. Re-mergear en posiciones nuevas
    for mr in merged_ranges:
        min_r = _shift_row(mr.min_row, insert_at, count)
        max_r = _shift_row(mr.max_row, insert_at, count)
        ws.merge_cells(start_row=min_r, start_column=mr.min_col, end_row=max_r, end_column=mr.max_col)

    # 6. Aplicar estilo de plantilla + limpiar valores en las filas nuevas
    new_start, new_end = insert_at, insert_at + count - 1
    if style_row_snapshot:
        for r in range(new_start, new_end + 1):
            for c in range(1, max_col + 1):
                tmpl = style_row_snapshot[c]
                cell = ws.cell(r, c)
                cell.value = None
                cell.font = copy.copy(tmpl["font"])
                cell.fill = copy.copy(tmpl["fill"])
                cell.border = copy.copy(tmpl["border"])
                cell.alignment = copy.copy(tmpl["alignment"])
                cell.number_format = tmpl["number_format"]

    return new_start, new_end


def extend_sum_ranges(ws, row: int, old_end: int, new_end: int):
    """
    Busca formulas tipo =SUM(X#:Y{old_end}) en una fila (normalmente la fila
    de totales) y extiende el limite final a new_end. Necesario porque el
    paso 4 de insert_rows_preserving_formulas solo desplaza referencias que
    ya estaban despues del punto de insercion, pero no "estira" un rango que
    terminaba justo antes.
    """
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row, c)
        if isinstance(cell.value, str) and cell.value.startswith("="):
            cell.value = re.sub(
                rf":([A-Z]{{1,3}})\$?{old_end}\)",
                rf":\g<1>{new_end})",
                cell.value,
            )


def rebase_formula_row(formula: str, old_row: int, new_row: int) -> str:
    """Adapta una formula "propia de su fila" (ej. =O14*D14, que solo usa
    celdas de esa misma fila) a una fila nueva: cambia unicamente las
    referencias que apuntaban exactamente a `old_row` por `new_row`, y deja
    cualquier otra referencia (a otra fila fija, como un encabezado) tal
    cual."""

    def repl(m):
        col_abs, col, row_abs, row = m.groups()
        if int(row) == old_row:
            return f"{col_abs}{col}{row_abs}{new_row}"
        return m.group(0)

    return re.sub(r"(\$?)([A-Z]{1,3})(\$?)(\d+)", repl, formula)


def highlight_row_yellow(ws, row: int, col_start: int, col_end: int):
    """Marca en amarillo una fila (empleado nuevo sin tarifa historica)."""
    from openpyxl.styles import PatternFill
    fill = PatternFill(start_color=YELLOW_FILL_RGB, end_color=YELLOW_FILL_RGB, fill_type="solid")
    for c in range(col_start, col_end + 1):
        ws.cell(row, c).fill = fill
