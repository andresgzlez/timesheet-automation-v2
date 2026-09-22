"""
Memoria de tarifas de trabajadores, en dos niveles:

- Global (una sola base, entre TODOS los proyectos): el INTAL_RATE --lo
  que le pagamos a la persona-- es suyo, no cambia segun en que proyecto
  trabaje esa semana. Si alguien falta en Uplift pero ya trabajo antes en
  Cub Spine (otro proyecto), su INTAL_RATE se puede sacar de ahi sin
  problema.

- Por proyecto: el RATE --lo que se le cobra al cliente de ESE proyecto
  por esa posicion-- SI es especifico de cada proyecto y nunca se copia
  de otro. Si alguien nunca trabajo en este proyecto, su RATE se deja en
  0 aunque ya sepamos su INTAL_RATE por el lado global.

Alguien se marca en amarillo (revisar a mano) cuando no tenemos un
INTAL_RATE VIGENTE: o nunca lo supimos, o lo supimos pero hace mas de 3
meses (90 dias) que esa persona no aparece con datos reales -- pudo
cambiar de tarifa en ese tiempo, asi que no se le pone su numero viejo
directo. En ese caso, si sabemos su posicion, se le pone como RATE de
cobro la tarifa tipica que hoy se le cobra a esa posicion en ESE
proyecto (la mediana entre la gente activa de esa posicion), y el
INTAL_RATE se deja en 0 para que alguien lo confirme.

Nota (limite conocido v1): esto se guarda como JSON en disco dentro del
propio contenedor. En Railway, el disco no es garantizado-persistente
entre redeploys (si se reconstruye la app, esta memoria se pierde). Para
produccion real hay que moverlo a una base de datos o a un volumen
persistente -- aqui alcanza para validar que la logica funciona.
"""

import difflib
import json
import re
from datetime import date
from pathlib import Path

from engine.matching import normalize, names_match

RATES_DIR = Path(__file__).resolve().parent.parent / "profiles" / "data" / "rates"
GLOBAL_PATH = RATES_DIR / "_global_workers.json"


def _slug_project(project_name: str) -> str:
    """Convierte el nombre del proyecto en nombre de archivo. A diferencia
    de `normalize` (pensada para comparar nombres de PERSONAS, que ignora
    numeros finales como en "Yaeki 2"), aqui SI hay que conservar los
    numeros -- "Novo AP2" y "Novo AP3" son proyectos distintos, cada uno
    con su propia tarifa de cobro, y no deben caer en el mismo archivo."""
    s = re.sub(r"[^A-Za-z0-9]+", "_", project_name.strip().upper()).strip("_")
    return s or "PROYECTO_SIN_NOMBRE"


def _path(project_name: str) -> Path:
    RATES_DIR.mkdir(parents=True, exist_ok=True)
    return RATES_DIR / f"{_slug_project(project_name)}.json"


def _read_json(p: Path) -> dict:
    if p.exists():
        try:
            return json.loads(p.read_text())
        except Exception:
            return {}
    return {}


def load(project_name: str) -> dict:
    """Memoria de ESTE proyecto: posicion + RATE de cobro visto la ultima vez."""
    return _read_json(_path(project_name))


def save(project_name: str, store: dict):
    _path(project_name).write_text(json.dumps(store, indent=2, ensure_ascii=False))


def load_global() -> dict:
    """Memoria de trabajadores entre TODOS los proyectos: solo INTAL_RATE
    (lo que se le paga a la persona) y su posicion mas reciente, como pista."""
    RATES_DIR.mkdir(parents=True, exist_ok=True)
    return _read_json(GLOBAL_PATH)


def save_global(store: dict):
    RATES_DIR.mkdir(parents=True, exist_ok=True)
    GLOBAL_PATH.write_text(json.dumps(store, indent=2, ensure_ascii=False))


def _tokens_fuzzy_match(a_tokens: list[str], b_tokens: list[str], cutoff: float = 0.8) -> bool:
    """Como is_subsequence, pero cada palabra puede diferir un poco (permite
    errores de tipeo por palabra, ej. "MATUTE" vs "MATURE" -> 0.83), en vez
    de exigir la palabra exacta. Compara la lista mas corta contra la mas
    larga, en el mismo orden relativo."""
    short, long_ = (a_tokens, b_tokens) if len(a_tokens) <= len(b_tokens) else (b_tokens, a_tokens)
    if not short:
        return False
    it = iter(long_)
    for tok in short:
        found = False
        for cand in it:
            if difflib.SequenceMatcher(None, tok, cand).ratio() >= cutoff:
                found = True
                break
        if not found:
            return False
    return True


def lookup(store: dict, name: str) -> dict | None:
    """Busca a la persona en la memoria, de mas a menos estricto:
    1. Nombre normalizado exacto.
    2. Coincidencia de tokens exacta (permite apellidos de mas/menos u
       orden distinto -- la misma logica que emparejar nombres).
    3. Similitud de texto sobre el nombre completo (para nombres cortos
       parecidos).
    4. Similitud de texto POR PALABRA (permite errores de tipeo en un
       apellido, ej. "Matute"/"Mature", aunque el nombre completo tenga
       palabras de mas que bajen la similitud total).
    Nunca decide entre dos personas distintas con nombre parecido -- si
    hay mas de un candidato cercano en un mismo nivel, no arriesga."""
    key = normalize(name)
    if key in store:
        return store[key]

    token_matches = [k for k in store if names_match(name, k)]
    if len(token_matches) == 1:
        return store[token_matches[0]]

    close = difflib.get_close_matches(key, store.keys(), n=2, cutoff=0.84)
    if len(close) == 1:
        return store[close[0]]

    name_tokens = key.split()
    fuzzy_matches = [k for k in store if _tokens_fuzzy_match(name_tokens, k.split())]
    if len(fuzzy_matches) == 1:
        return store[fuzzy_matches[0]]

    return None


def _bump_last_seen(entry: dict, last_seen: str | None, has_data: bool):
    """Actualiza la fecha de "ultima vez vista con datos reales", solo hacia
    adelante (nunca retrocede si ya teniamos una fecha mas reciente)."""
    if not last_seen or not has_data:
        return
    prev = entry.get("last_seen")
    if not prev or last_seen > prev:
        entry["last_seen"] = last_seen


def record(store: dict, name: str, position=None, intal_rate=None, rate=None, last_seen: str | None = None):
    """Guarda en una memoria de PROYECTO: posicion y RATE de cobro (el
    INTAL_RATE tambien se guarda aqui como respaldo, pero la fuente de
    verdad para el INTAL_RATE es la memoria global). `last_seen` (fecha
    ISO, ej. "2026-09-20") marca cuando se vio a esta persona con datos
    reales -- de ahi se calcula si esta vigente o "vencida" (+90 dias)."""
    key = normalize(name)
    entry = store.get(key, {})
    if position:
        entry["position"] = position
    if intal_rate not in (None, 0):
        entry["intal_rate"] = intal_rate
    if rate not in (None, 0):
        entry["rate"] = rate
    has_data = bool(position or intal_rate not in (None, 0) or rate not in (None, 0))
    _bump_last_seen(entry, last_seen, has_data)
    if entry:
        store[key] = entry


def record_global(store: dict, name: str, position=None, intal_rate=None, last_seen: str | None = None):
    """Guarda en la memoria GLOBAL de trabajadores: solo INTAL_RATE (lo que
    se le paga a la persona) y posicion, sin importar en que proyecto."""
    key = normalize(name)
    entry = store.get(key, {})
    if position:
        entry["position"] = position
    if intal_rate not in (None, 0):
        entry["intal_rate"] = intal_rate
    has_data = bool(position or intal_rate not in (None, 0))
    _bump_last_seen(entry, last_seen, has_data)
    if entry:
        store[key] = entry


def is_stale(entry: dict | None, reference_date: str | None = None, max_days: int = 90) -> bool:
    """True si no sabemos cuando fue la ultima vez que esta persona aparecio
    con datos reales, o si fue hace mas de `max_days` dias (3 meses). Sin
    fecha guardada se trata como vencido -- mas seguro que asumir que sigue
    activo con una tarifa que pudo haber cambiado."""
    if not entry:
        return True
    last_seen = entry.get("last_seen")
    if not last_seen:
        return True
    ref = date.fromisoformat(reference_date) if reference_date else date.today()
    try:
        seen = date.fromisoformat(last_seen)
    except ValueError:
        return True
    return (ref - seen).days > max_days


def position_standard_rate(project_store: dict, position: str | None,
                            reference_date: str | None = None, max_days: int = 90) -> float | None:
    """Tarifa de cobro tipica de una posicion en ESTE proyecto: la mediana
    del RATE entre la gente de esa posicion que sigue vigente (vista en los
    ultimos `max_days` dias). Para alguien que vuelve despues de mucho
    tiempo sin aparecer aqui, no se le pone SU tarifa vieja (pudo cambiar):
    se le pone la que hoy se le cobra al cliente por esa posicion."""
    if not position:
        return None
    pos_key = position.strip().upper()
    rates = []
    for entry in project_store.values():
        entry_pos = entry.get("position")
        if not entry_pos or entry_pos.strip().upper() != pos_key:
            continue
        if entry.get("rate") in (None, 0):
            continue
        if is_stale(entry, reference_date, max_days):
            continue
        rates.append(entry["rate"])
    if not rates:
        return None
    rates.sort()
    n = len(rates)
    mid = n // 2
    return rates[mid] if n % 2 else (rates[mid - 1] + rates[mid]) / 2
