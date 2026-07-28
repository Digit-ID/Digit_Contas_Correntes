"""
Motor de conciliacion de cuentas corrientes de proveedores.

Logica (segun especificacion del usuario):
  - Se trabaja sobre la columna de importes con signo llamada 'DEB_CRED'
    (positivos y negativos) y la columna de fecha 'DATA_CERTA'. Ambas se
    localizan por NOMBRE de cabecera, no por posicion (el orden puede variar).
  - Se buscan grupos de lineas cuyo DEB_CRED SUMA CERO: una factura con su
    pago, una factura con varios pagos, incluyendo notas de credito.
  - Desempate cuando un importe puede casar con varias lineas: se elige la de
    fecha DATA_CERTA mas cercana.
  - Tolerancia por defecto: +/- 0,01 (un centimo) por redondeos.
  - Salida: cada linea recibe CONCILIADO (True/False) y una referencia
    correlativa 'Referencia 001', 'Referencia 002'... (misma para todas las
    lineas del grupo). Si no casa con nada -> 'SEM IDENTIFICAR' y False.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from itertools import combinations
import re
import unicodedata

import pandas as pd


# --------------------------------------------------------------------------
# Configuracion
# --------------------------------------------------------------------------

@dataclass
class ReconConfig:
    tolerancia_centimos: int = 1        # +/- 0,01
    max_lineas_por_grupo: int = 8       # tamano max de un grupo (1 ancla + N)
    max_candidatas: int = 40            # limite de pool para busqueda de subconjunto
    prefijo_ref: str = "Referência"     # texto de la referencia
    texto_sin_id: str = "SEM IDENTIFICAR"


# --------------------------------------------------------------------------
# Deteccion de columnas por nombre (tolerante a acentos/espacios/orden)
# --------------------------------------------------------------------------

def _norm_header(value) -> str:
    if value is None:
        return ""
    s = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Za-z0-9]", "", s).upper()


# alias aceptados para cada campo (todos normalizados)
_ALIAS_VALOR = {"DEBCRED", "DEBECRED", "DEBITOCREDITO", "SALDO", "VALOR"}
_ALIAS_FECHA = {"DATACERTA", "FECHACIERTA", "FECHA", "DATA", "DATALANC", "DATALANCAMENTO"}


def detectar_columna(headers: list, alias: set[str], preferido: str) -> str | None:
    """Devuelve el nombre original de la cabecera que coincide con el alias."""
    norm_pref = _norm_header(preferido)
    # 1) coincidencia exacta con el nombre preferido
    for h in headers:
        if _norm_header(h) == norm_pref:
            return h
    # 2) coincidencia con cualquier alias
    for h in headers:
        if _norm_header(h) in alias:
            return h
    # 3) contiene el nombre preferido
    for h in headers:
        if norm_pref and norm_pref in _norm_header(h):
            return h
    return None


def detectar_columnas(headers: list) -> dict:
    return {
        "valor": detectar_columna(headers, _ALIAS_VALOR, "DEB_CRED"),
        "fecha": detectar_columna(headers, _ALIAS_FECHA, "DATA_CERTA"),
    }


# --------------------------------------------------------------------------
# Parsing de importes (formato europeo/portugues: 1.234,56 y -992,63)
# --------------------------------------------------------------------------

def _to_cents(value) -> int | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)):
        return int(round(float(value) * 100))
    s = str(value).strip()
    if s in ("", "-", "—", "–"):
        return None
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg, s = True, s[1:-1]
    s = s.replace("€", "").replace(" ", "").replace(" ", "")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):          # 1.234,56
            s = s.replace(".", "").replace(",", ".")
        else:                                     # 1,234.56
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        c = int(round(float(s) * 100))
    except ValueError:
        return None
    return -c if neg else c


# --------------------------------------------------------------------------
# Busqueda de subconjunto que suma cero con el ancla
# --------------------------------------------------------------------------

def _mejor_subconjunto(cands: list[tuple[int, int, float]], objetivo: int,
                       tol: int, max_n: int, fecha_ancla):
    """
    cands: lista de (idx, valor_cents, dias_dist_al_ancla)
    Devuelve la lista de idx cuyo valor suma ~objetivo, prefiriendo el grupo
    con menor distancia de fechas (desempate por DATA_CERTA mas cercana).
    """
    n = len(cands)
    max_n = min(max_n, n)
    for r in range(1, max_n + 1):
        soluciones = []
        for combo in combinations(range(n), r):
            s = sum(cands[i][1] for i in combo)
            if abs(s - objetivo) <= tol:
                dist = sum(cands[i][2] for i in combo)
                soluciones.append((dist, [cands[i][0] for i in combo]))
        if soluciones:
            soluciones.sort(key=lambda x: x[0])   # menor distancia de fechas gana
            return soluciones[0][1]
    return None


# --------------------------------------------------------------------------
# API principal
# --------------------------------------------------------------------------

@dataclass
class Resultado:
    asignacion: dict          # excel_row -> {"conciliado": bool, "referencia": str, "grupo": int|None}
    grupos: list              # lista de dicts con info de cada grupo
    resumen: dict = field(default_factory=dict)
    col_valor: str | None = None
    col_fecha: str | None = None


def conciliar(filas: list[dict], cfg: ReconConfig | None = None) -> Resultado:
    """
    filas: lista de dicts con claves:
        'row'   -> identificador de fila (p.ej. numero de fila en Excel)
        'valor' -> importe con signo (cents, int) o None si vacia
        'fecha' -> pd.Timestamp o None
    """
    cfg = cfg or ReconConfig()

    # solo participan las lineas con importe distinto de cero
    activas = [f for f in filas if f["valor"] not in (None, 0)]
    pendientes = {f["row"]: f for f in activas}

    grupos = []
    gid = 0

    def _dist_dias(a, b):
        if a is None or b is None or pd.isna(a) or pd.isna(b):
            return 9999
        return abs((a - b).days)

    # ---- Fase 1: 1:1 por importe opuesto + fecha mas cercana ----
    # Se generan todas las parejas posibles (importe opuesto dentro de la
    # tolerancia) y se casan de forma GLOBAL empezando por la de fecha mas
    # cercana. Asi el desempate no depende del orden de las filas: cada linea
    # acaba con su opuesta de DATA_CERTA mas proxima.
    posibles = []
    filas_activas = sorted(pendientes.values(), key=lambda x: x["row"])
    for i in range(len(filas_activas)):
        a = filas_activas[i]
        for j in range(i + 1, len(filas_activas)):
            b = filas_activas[j]
            if abs(a["valor"] + b["valor"]) <= cfg.tolerancia_centimos and \
                    (a["valor"] > 0) != (b["valor"] > 0):
                posibles.append((_dist_dias(a["fecha"], b["fecha"]),
                                 a["row"], b["row"]))
    posibles.sort(key=lambda p: (p[0], p[1], p[2]))  # menor distancia primero
    for dist, ra, rb in posibles:
        if ra in pendientes and rb in pendientes:
            gid += 1
            miembros = [ra, rb]
            grupos.append({"id": gid, "rows": miembros, "tipo": "factura+pago (1:1)"})
            pendientes.pop(ra, None)
            pendientes.pop(rb, None)

    # ---- Fase 2: 1:varias / varias:1 (subconjunto que suma cero) ----
    # ancla = la linea de mayor importe absoluto; se le buscan opuestas
    for row in sorted(pendientes.keys(), key=lambda r: -abs(pendientes[r]["valor"])):
        if row not in pendientes:
            continue
        ancla = pendientes[row]
        objetivo = -ancla["valor"]
        cands = [
            (g["row"], g["valor"], _dist_dias(g["fecha"], ancla["fecha"]))
            for g in pendientes.values()
            if g["row"] != row and (g["valor"] > 0) != (ancla["valor"] > 0)
        ]
        # ordenar por cercania de fecha y acotar el pool
        cands.sort(key=lambda c: c[2])
        cands = cands[: cfg.max_candidatas]
        combo = _mejor_subconjunto(
            cands, objetivo, cfg.tolerancia_centimos,
            cfg.max_lineas_por_grupo - 1, ancla["fecha"],
        )
        if combo is None:
            continue
        gid += 1
        miembros = [row] + list(combo)
        grupos.append({"id": gid, "rows": miembros, "tipo": "pago de varias facturas"})
        for m in miembros:
            pendientes.pop(m, None)

    # ---- Construir asignacion ----
    asignacion = {}
    for f in filas:
        asignacion[f["row"]] = {"conciliado": False,
                                "referencia": cfg.texto_sin_id,
                                "grupo": None}

    # numerar referencias por orden de aparicion (fila mas alta del grupo)
    grupos.sort(key=lambda g: min(g["rows"]))
    for n, g in enumerate(grupos, start=1):
        ref = f"{cfg.prefijo_ref} {n:03d}"
        g["referencia"] = ref
        for m in g["rows"]:
            asignacion[m] = {"conciliado": True, "referencia": ref, "grupo": n}

    total = len(filas)
    casadas = sum(1 for a in asignacion.values() if a["conciliado"])
    saldo = sum(f["valor"] for f in filas
                if f["valor"] and not asignacion[f["row"]]["conciliado"])
    resumen = {
        "lineas_total": total,
        "lineas_casadas": casadas,
        "lineas_sueltas": total - casadas,
        "grupos": len(grupos),
        "saldo_abierto_eur": round((saldo or 0) / 100, 2),
    }
    return Resultado(asignacion, grupos, resumen)
