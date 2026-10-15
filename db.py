"""Persistencia en SQLite (archivo local). Importes en centavos enteros."""
from __future__ import annotations

import os
import sqlite3
from decimal import Decimal
from pathlib import Path

DB_PATH = Path(os.environ.get("AUTORRETENCION_DB", Path(__file__).parent / "data" / "autorretencion.db"))

SCHEMA = """
CREATE TABLE IF NOT EXISTS movimientos(
  key TEXT PRIMARY KEY, id TEXT, empresa TEXT NOT NULL, centrocosto TEXT, fecha TEXT NOT NULL,
  tipodocumento TEXT, nrodocumento TEXT, codigocuenta TEXT NOT NULL, nombrecuenta TEXT, descripcion TEXT,
  factura TEXT, identificacion TEXT, tercero TEXT, valor_c INTEGER NOT NULL, debito_c INTEGER, credito_c INTEGER,
  fila INTEGER, base_c INTEGER DEFAULT 0, porcentaje TEXT DEFAULT '0', tipo_persona TEXT DEFAULT '');
CREATE INDEX IF NOT EXISTS ix_mov ON movimientos(empresa, fecha);
CREATE TABLE IF NOT EXISTS meta(k TEXT PRIMARY KEY, v TEXT);
CREATE TABLE IF NOT EXISTS empresas(
  nombre TEXT PRIMARY KEY, tarifa TEXT, activa INTEGER NOT NULL DEFAULT 1, observaciones TEXT DEFAULT '');
"""
COLS = ["id", "empresa", "centrocosto", "fecha", "tipodocumento", "nrodocumento", "codigocuenta", "nombrecuenta",
        "descripcion", "factura", "identificacion", "tercero", "tipo_persona"]
c100 = lambda d: int((d * 100).to_integral_value())  # noqa: E731


def conn() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    c.executescript(SCHEMA)
    # migración de bases creadas con versiones anteriores
    cols = {r["name"] for r in c.execute("PRAGMA table_info(movimientos)")}
    for name, ddl in (("base_c", "INTEGER DEFAULT 0"), ("porcentaje", "TEXT DEFAULT '0'"), ("tipo_persona", "TEXT DEFAULT ''")):
        if name not in cols:
            c.execute(f"ALTER TABLE movimientos ADD COLUMN {name} {ddl}")
    return c


def _to_mov(r: sqlite3.Row) -> dict:
    m = {k: r[k] for k in ["key", *COLS, "fila"]}
    m["valor"], m["debito"], m["credito"], m["baseimpuesto"] = (Decimal(r[k] or 0) / 100 for k in ("valor_c", "debito_c", "credito_c", "base_c"))
    m["porcentajeimpuesto"] = Decimal(r["porcentaje"] or 0)
    return m


def upsert_movimientos(rows: list[dict]) -> dict:
    """Inserta o actualiza por clave (empresa+id). Reprocesar un archivo NO duplica."""
    added = updated = unchanged = repeated = 0
    seen = set()
    with conn() as c:
        for m in rows:
            if m["key"] in seen:
                repeated += 1
            seen.add(m["key"])
            vals = {**{k: m[k] for k in COLS}, "key": m["key"], "valor_c": c100(m["valor"]),
                    "debito_c": c100(m["debito"]), "credito_c": c100(m["credito"]), "fila": m["fila"],
                    "base_c": c100(m["baseimpuesto"]), "porcentaje": str(m["porcentajeimpuesto"])}
            prev = c.execute("SELECT * FROM movimientos WHERE key=?", (m["key"],)).fetchone()
            if prev is None:
                added += 1
            elif all(prev[k] == vals[k] for k in ("valor_c", "fecha", "codigocuenta", "descripcion", "tercero",
                                                  "nrodocumento", "base_c", "tipo_persona")):
                unchanged += 1
                continue
            else:
                updated += 1
            cols = list(vals)
            c.execute(f"INSERT OR REPLACE INTO movimientos({','.join(cols)}) VALUES({','.join('?' * len(cols))})",
                      [vals[k] for k in cols])
    return dict(added=added, updated=updated, unchanged=unchanged, repeated=repeated)


def movimientos(empresa: str | None = None) -> list[dict]:
    with conn() as c:
        q, a = "SELECT * FROM movimientos", ()
        if empresa:
            q, a = q + " WHERE empresa=?", (empresa,)
        return [_to_mov(r) for r in c.execute(q + " ORDER BY fecha, CAST(id AS INTEGER)", a)]


def inventario() -> list[dict]:
    with conn() as c:
        return [dict(r) for r in c.execute(
            "SELECT empresa, substr(fecha,1,7) mes, COUNT(*) movimientos, "
            "SUM(codigocuenta LIKE '4%') ingresos, "
            "SUM(codigocuenta LIKE '2365%' OR codigocuenta LIKE '2368%') retenciones "
            "FROM movimientos GROUP BY 1,2 ORDER BY 1,2")]


def borrar_movimientos() -> None:
    with conn() as c:
        c.execute("DELETE FROM movimientos")


def total_movimientos() -> int:
    with conn() as c:
        return c.execute("SELECT COUNT(*) FROM movimientos").fetchone()[0]


# ---- maestro de empresas
def empresas() -> list[dict]:
    with conn() as c:
        return [{**dict(r), "activa": bool(r["activa"])} for r in c.execute("SELECT * FROM empresas ORDER BY nombre")]


def guardar_empresas(lista: list[dict]) -> None:
    """Reemplaza el maestro completo con la lista recibida."""
    with conn() as c:
        c.execute("DELETE FROM empresas")
        c.executemany("INSERT INTO empresas VALUES(?,?,?,?)",
                      [(e["nombre"].strip(), None if e.get("tarifa") in (None, "") else str(e["tarifa"]),
                        int(bool(e.get("activa", True))), e.get("observaciones") or "") for e in lista])


# ---- parámetros
def get_meta(k: str, default: str = "") -> str:
    with conn() as c:
        r = c.execute("SELECT v FROM meta WHERE k=?", (k,)).fetchone()
        return r["v"] if r else default


def set_meta(k: str, v: str) -> None:
    with conn() as c:
        c.execute("INSERT OR REPLACE INTO meta VALUES(?,?)", (k, v))
