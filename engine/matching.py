"""
Motor de matching de nombres entre el archivo fuente (timesheet del lider de
proyecto) y la plantilla destino (roster de la empresa).

Extraido y generalizado de la logica usada manualmente en: JJ Production
Upfit, JJ Warehouse Upfit, JJ Exterior, NIPRO, NOVO AP2, Project Daisy y
HCA Palms West.

Reglas que replica:
- Normaliza (quita parentesis, caracteres no alfabeticos, mayusculas)
- Matching por subsecuencia de tokens (permite nombres parciales, orden
  distinto, apellidos de mas o de menos)
- Nombres alternativos entre parentesis, incluyendo varios separados por " - "
  (ej. "JEFREY ALEXANDER LOBO GARCIA (ALEX LOBO)" debe matchear "Alex Lobo")
- Cuando hay ambiguedad (varios candidatos posibles), NO decide sola: la deja
  marcada para revision humana en vez de adivinar (regla de Project Daisy)
"""

import re
from dataclasses import dataclass, field


def normalize(name: str) -> str:
    """Quita parentesis, caracteres no alfabeticos, y pasa a mayusculas."""
    if not name:
        return ""
    base = re.sub(r"\(.*?\)", "", name)
    base = re.sub(r"[^A-Za-z\s]", " ", base)
    return " ".join(base.upper().split())


def name_variants(full_name: str) -> list[str]:
    """
    Devuelve todas las formas normalizadas bajo las que un nombre puede
    aparecer: el nombre base, y cada nombre alterno entre parentesis
    (separados por " - " si hay varios, ej. apodos combinados).
    """
    if not full_name:
        return []
    base = re.sub(r"\(.*?\)", "", full_name)
    variants = [normalize(base)]
    for paren in re.findall(r"\((.*?)\)", full_name):
        for alt in re.split(r"\s*-\s*", paren):
            v = normalize(alt)
            if v and v not in variants:
                variants.append(v)
    return [v for v in variants if v]


def is_subsequence(needle_tokens: list[str], haystack_tokens: list[str]) -> bool:
    """True si todos los tokens de needle aparecen en haystack, en el mismo
    orden relativo (no necesariamente consecutivos). Permite nombres
    parciales o con apellidos de mas."""
    it = iter(haystack_tokens)
    return all(tok in it for tok in needle_tokens)


def names_match(source_name: str, dest_name: str) -> bool:
    """True si alguna variante del nombre fuente matchea alguna variante del
    nombre destino via subsecuencia de tokens (en cualquier direccion)."""
    src_variants = name_variants(source_name)
    dst_variants = name_variants(dest_name)
    for s in src_variants:
        s_tokens = s.split()
        for d in dst_variants:
            d_tokens = d.split()
            if is_subsequence(s_tokens, d_tokens) or is_subsequence(d_tokens, s_tokens):
                return True
    return False


@dataclass
class MatchResult:
    matches: dict = field(default_factory=dict)      # source_name -> dest_row
    unmatched_source: list = field(default_factory=list)   # gente en fuente, no en roster
    ambiguous: list = field(default_factory=list)     # (source_name, [candidatos]) - necesita revision humana


def match_source_to_roster(source_names: list[str], roster: dict[int, str]) -> MatchResult:
    """
    roster: {fila_destino: nombre_destino}
    Devuelve MatchResult. Nombres ambiguos (mas de un candidato posible) NO
    se resuelven solos -> quedan en `ambiguous` para que un humano decida,
    igual que se hizo con los typos de Project Daisy.
    """
    result = MatchResult()
    for sname in source_names:
        candidates = [row for row, dname in roster.items() if names_match(sname, dname)]
        if len(candidates) == 1:
            result.matches[sname] = candidates[0]
        elif len(candidates) == 0:
            result.unmatched_source.append(sname)
        else:
            result.ambiguous.append((sname, candidates))
    return result
