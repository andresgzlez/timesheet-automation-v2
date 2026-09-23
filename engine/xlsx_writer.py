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

import bisect
import copy
import re

YELLOW_FILL_RGB = "FFFF00"

# Colores fijos para diferenciar de un vistazo a que posicion pertenece
# cada fila (roster y cuadros de resumen tipo "AVERAGE"/"INVOICE" al pie
# de la hoja) -- confirmados por el cliente sobre una plantilla real:
# HELPER azul, SAFETY verde, FOREMAN gris, SPOTTER durazno, MECHANIC sin
# relleno (blanco). Una posicion que no este en este mapa (o que no
# tengamos como identificar) se deja sin relleno -- mejor blanco que un
# color inventado sin significado.
POSITION_FILL_RGB = {
    "HELPER": "ADD8E6",
    "SAFETY": "C6E0B4",
    "FOREMAN": "D9D9D9",
    "SPOTTER": "FFD9B3",
    "MECHANIC": None,
    "LABOR": "FBE5D6",
    "ELECTRICIAN": "FFE699",
}


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


def insert_rows_preserving_formulas_multi(ws, insert_ats: list[int], style_template_row: int | None = None):
    """
    Igual que insert_rows_preserving_formulas, pero inserta UNA fila en
    cada una de varias posiciones distintas (`insert_ats`, en coordenadas
    de la hoja ANTES de insertar nada) en una sola pasada por toda la hoja.

    Existe porque llamar insert_rows_preserving_formulas una vez POR
    PERSONA (para poder ubicar a cada quien en su lugar alfabetico) es
    O(filas x columnas) CADA VEZ -- en una plantilla con muchas columnas
    (algunas traen cientos, casi todas vacias, pero openpyxl igual las
    cuenta) eso puede tardar mas de un minuto con solo unas pocas personas
    nuevas. Esta version hace ese trabajo pesado una unica vez sin
    importar cuantas filas se esten insertando.

    `insert_ats` debe venir en el orden final deseado (asc; si dos
    entradas comparten la misma posicion, quedan una despues de la otra en
    ese mismo orden). Devuelve la lista de filas fisicas nuevas, en ese
    mismo orden.
    """
    if not insert_ats:
        return []

    sorted_ats = sorted(insert_ats)
    count = len(sorted_ats)

    def shift_for(row: int) -> int:
        # cuantas inserciones caen en o antes de esta fila (original)
        return bisect.bisect_right(sorted_ats, row)

    max_row = ws.max_row
    max_col = ws.max_column

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

    merged_ranges = list(ws.merged_cells.ranges)
    for mr in merged_ranges:
        ws.unmerge_cells(str(mr))

    style_row_snapshot = None
    if style_template_row is not None:
        style_row_snapshot = {
            c: snapshot[(style_template_row, c)] for c in range(1, max_col + 1)
        }

    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(r, c).value = None

    def _rewrite_formula_refs_multi(formula: str) -> str:
        def repl(m):
            col_abs, col, row_abs, row = m.groups()
            row_num = int(row)
            new_row = row_num + shift_for(row_num)
            return f"{col_abs}{col}{row_abs}{new_row}"
        return re.sub(r"(\$?)([A-Z]{1,3})(\$?)(\d+)", repl, formula)

    for (r, c), data in snapshot.items():
        new_r = r + shift_for(r)
        cell = ws.cell(new_r, c)
        value = data["value"]
        if isinstance(value, str) and value.startswith("="):
            value = "=" + _rewrite_formula_refs_multi(value[1:])
        cell.value = value
        cell.font = data["font"]
        cell.fill = data["fill"]
        cell.border = data["border"]
        cell.alignment = data["alignment"]
        cell.number_format = data["number_format"]

    for mr in merged_ranges:
        min_r = mr.min_row + shift_for(mr.min_row)
        max_r = mr.max_row + shift_for(mr.max_row)
        ws.merge_cells(start_row=min_r, start_column=mr.min_col, end_row=max_r, end_column=mr.max_col)

    new_rows = [at + i for i, at in enumerate(sorted_ats)]

    if style_row_snapshot:
        for nr in new_rows:
            for c in range(1, max_col + 1):
                tmpl = style_row_snapshot[c]
                cell = ws.cell(nr, c)
                cell.value = None
                cell.font = copy.copy(tmpl["font"])
                # el relleno/color de fondo de la fila de referencia NO se
                # copia -- si esa fila tenia algun color puesto a mano (por
                # cualquier motivo, sin relacion con la persona nueva), no
                # tiene sentido que TODOS los agregados hereden ese mismo
                # color. Fila nueva = sin relleno; el unico color que se
                # pone a proposito es el amarillo de "sin tarifa
                # confirmada" (ver highlight_row_yellow, mas abajo).
                cell.border = copy.copy(tmpl["border"])
                cell.alignment = copy.copy(tmpl["alignment"])
                cell.number_format = tmpl["number_format"]

    # devolver en el orden en que el llamador paso insert_ats, no el orden
    # ordenado -- como sorted() es estable y new_rows ya sale ascendente
    # segun sorted_ats, alcanza con mapear cada insert_at original a su
    # fila fisica en el mismo orden en que fueron apareciendo.
    return new_rows


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


def fix_sum_range_start(ws, row: int, start_row: int):
    """
    Fuerza el limite INICIAL de cualquier =SUM(COL#:...) en esta fila a
    `start_row`. El limite inicial del roster (la primera fila de datos)
    es un ancla fija que nunca deberia moverse -- pero si alguien nuevo se
    inserta alfabeticamente ANTES del primer empleado (ej. el roster
    empezaba en "Carlos" y ahora "Alicia" entra antes), el desplazamiento
    generico de referencias de fila (pensado para que el limite FINAL
    crezca) tambien mueve por error ese limite inicial, dejando fuera las
    filas nuevas que quedaron arriba. Se llama despues de cualquier
    insercion en el roster, sin importar donde haya caido.
    """
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row, c)
        v = cell.value
        if isinstance(v, str) and v.upper().startswith("=SUM("):
            cell.value = re.sub(
                r"^(=SUM\()([A-Z]{1,3})\$?\d+",
                rf"\g<1>\g<2>{start_row}",
                v,
                count=1,
                flags=re.IGNORECASE,
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


def _resolve_position_bucket(position: str | None, merge_positions: dict | None = None) -> str | None:
    if not position:
        return None
    key = position.strip().upper()
    if merge_positions:
        key = merge_positions.get(key, key)
    return key


def apply_position_fill(ws, row: int, col_start: int, col_end: int, position: str | None,
                         merge_positions: dict | None = None):
    """Colorea una fila segun la posicion de la persona (POSITION_FILL_RGB),
    o la deja sin relleno si la posicion no esta en el mapa conocido."""
    from openpyxl.styles import PatternFill
    bucket = _resolve_position_bucket(position, merge_positions)
    rgb = POSITION_FILL_RGB.get(bucket) if bucket else None
    fill = PatternFill(start_color=rgb, end_color=rgb, fill_type="solid") if rgb else PatternFill(fill_type=None)
    for c in range(col_start, col_end + 1):
        ws.cell(row, c).fill = fill


def clear_row_fill(ws, row: int, col_start: int, col_end: int):
    """Quita cualquier relleno/color de fondo de una fila (columnas
    col_start..col_end), dejandola sin relleno.

    La plantilla que se sube cada semana suele ser una copia de la salida
    de la semana anterior, asi que cualquier color que haya quedado puesto
    (a mano, o por un bug de una version vieja de esta app que coloreaba
    filas sin querer) se arrastra semana tras semana sin que nuestro
    codigo lo haya pedido -- results en un roster donde cada quien
    termina con "su color" al azar, sin ningun significado real. Se llama
    sobre TODO el roster existente antes de escribir los datos de esta
    semana, para que cada corrida salga limpia; el unico color que se
    aplica a proposito de ahi en mas es el amarillo de
    highlight_row_yellow (para alguien sin tarifa confirmada)."""
    from openpyxl.styles import PatternFill
    no_fill = PatternFill(fill_type=None)
    for c in range(col_start, col_end + 1):
        ws.cell(row, c).fill = no_fill
