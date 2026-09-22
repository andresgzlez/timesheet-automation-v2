"""
Motor de calculo de REG (horas regulares) y OT (horas extra).

Encapsula las 3 estrategias que hemos usado en los proyectos reales:

1. `exact_source`: usar el REG/OT que ya trae la fuente, tal cual, sin
   recalcular (la mayoria de proyectos: JJ, NIPRO, NOVO, Project Daisy).
   La fuente puede tener OT diario que no sigue la regla simple de 40h/semana
   y hay que respetarlo igual.

2. `weekly_40_threshold`: REG = min(total, 40), OT = max(total-40, 0).
   Es la formula que ya trae la plantilla destino por defecto; se usa
   cuando la fuente no da su propio split confiable.

3. `cumulative_row_order`: para el caso HCA Palms West, cuando una persona
   tiene VARIAS filas (una por codigo de costo). Las horas OT no se calculan
   por fecha calendario, sino acumulando el total de cada fila EN EL ORDEN
   en que las filas aparecen en la plantilla destino, aplicando el umbral de
   40h de forma secuencial sobre esa acumulacion.
   Regla importante confirmada por el cliente: si la persona solo tiene UNA
   fila (un solo codigo), se usa `exact_source` para esa fila -- el reparto
   acumulado solo aplica cuando hay que repartir entre varias filas.
"""

from dataclasses import dataclass


@dataclass
class DayHours:
    """Horas de una fila fuente (persona + codigo de costo, o persona sola)."""
    row_ref: object          # identificador de la fila (ej. numero de fila destino)
    days: list                # [lun, mar, mie, jue, vie, sab, dom]
    source_rt: float = 0.0    # REG que trae la fuente para esta fila/codigo
    source_ot: float = 0.0    # OT que trae la fuente para esta fila/codigo

    @property
    def total(self) -> float:
        return sum(d or 0 for d in self.days)


def exact_source(entry: DayHours) -> tuple[float, float]:
    """Usa los valores de REG/OT tal como vienen en la fuente."""
    return entry.source_rt, entry.source_ot


def weekly_40_threshold(entry: DayHours) -> tuple[float, float]:
    """REG hasta 40h, el resto OT -- la formula estandar de la plantilla."""
    total = entry.total
    reg = min(total, 40)
    ot = max(total - 40, 0)
    return reg, ot


def cumulative_row_order(entries: list[DayHours]) -> dict:
    """
    `entries`: todas las filas de UNA misma persona, en el orden en que
    aparecen en la plantilla destino (no en orden de fecha).

    Si solo hay una fila -> se respeta el REG/OT exacto de la fuente
    (puede incluir OT diario que la regla de 40h no capturaria).

    Si hay varias filas -> se acumula el total de horas de cada fila en ese
    orden, aplicando el umbral de 40h de forma secuencial. Esto puede
    producir un split de REG/OT distinto al que trae la fuente por fecha,
    y ese es el resultado correcto (confirmado con el caso Dilman Garcia).

    Devuelve {row_ref: (reg, ot)}.
    """
    if len(entries) == 1:
        e = entries[0]
        return {e.row_ref: exact_source(e)}

    result = {}
    cumulative = 0.0
    for e in entries:
        total = e.total
        new_cumulative = cumulative + total
        if new_cumulative <= 40:
            reg, ot = total, 0.0
        elif cumulative >= 40:
            reg, ot = 0.0, total
        else:
            reg = 40 - cumulative
            ot = total - reg
        result[e.row_ref] = (reg, ot)
        cumulative = new_cumulative
    return result
