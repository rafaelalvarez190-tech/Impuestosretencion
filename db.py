"""Persistencia en SQLite (archivo local). Importes en centavos enteros."""
from __future__ import annotations

import hashlib
import hmac
import os
import sqlite3
import uuid
from decimal import Decimal
from pathlib import Path

DB_PATH = Path(os.environ.get("AUTORRETENCION_DB", Path(__file__).parent / "data" / "autorretencion.db"))

SCHEMA = """
CREATE TABLE IF NOT EXISTS movimientos(
  key TEXT PRIMARY KEY, id TEXT, empresa TEXT NOT NULL, centrocosto TEXT, fecha TEXT NOT NULL,
  tipodocumento TEXT, nrodocumento TEXT, codigocuenta TEXT NOT NULL, nombrecuenta TEXT, descripcion TEXT,
  factura TEXT, identificacion TEXT, tercero TEXT, valor_c INTEGER NOT NULL, debito_c INTEGER, credito_c INTEGER,
  fila INTEGER);
CREATE INDEX IF NOT EXISTS ix_mov ON movimientos(empresa, fecha);
CREATE TABLE IF NOT EXISTS tarifas(
  id TEXT PRIMARY KEY, empresa TEXT NOT NULL, porcentaje TEXT NOT NULL, desde TEXT NOT NULL, hasta TEXT,
  activa INTEGER NOT NULL, referencia TEXT, observaciones TEXT);
CREATE TABLE IF NOT EXISTS meta(k TEXT PRIMARY KEY, v TEXT);
"""
COLS = ["id", "empresa", "centrocosto", "fecha", "tipodocumento", "nrodocumento", "codigocuenta", "nombrecuenta",
        "descripcion", "factura", "identificacion", "tercero"]
c100 = lambda d: int((d * 100).to_integral_value())  # noqa: E731


def conn() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    c.executescript(SCHEMA)
    return c


def _to_mov(r: sqlite3.Row) -> dict:
    m = {k: r[k] for k in ["key", *COLS, "fila"]}
    m["valor"], m["debito"], m["credito"] = (Decimal(r[k]) / 100 for k in ("valor_c", "debito_c", "credito_c"))
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
                    "debito_c": c100(m["debito"]), "credito_c": c100(m["credito"]), "fila": m["fila"]}
            prev = c.execute("SELECT * FROM movimientos WHERE key=?", (m["key"],)).fetchone()
            if prev is None:
                added += 1
            elif all(prev[k] == vals[k] for k in ("valor_c", "fecha", "codigocuenta", "descripcion", "tercero", "nrodocumento")):
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
            "SUM(codigocuenta LIKE '4%') ingresos FROM movimientos GROUP BY 1,2 ORDER BY 1,2")]


def borrar_movimientos() -> None:
    with conn() as c:
        c.execute("DELETE FROM movimientos")


def total_movimientos() -> int:
    with conn() as c:
        return c.execute("SELECT COUNT(*) FROM movimientos").fetchone()[0]


# ---- tarifas
def tarifas() -> list[dict]:
    with conn() as c:
        return [{**dict(r), "activa": bool(r["activa"])} for r in c.execute("SELECT * FROM tarifas ORDER BY empresa, desde")]


def guardar_tarifa(t: dict) -> dict:
    t = {**t, "id": t.get("id") or uuid.uuid4().hex, "porcentaje": str(t["porcentaje"]), "hasta": t.get("hasta") or None}
    with conn() as c:
        c.execute("INSERT OR REPLACE INTO tarifas VALUES(:id,:empresa,:porcentaje,:desde,:hasta,:activa,:referencia,:observaciones)",
                  {**t, "activa": int(t["activa"]), "referencia": t.get("referencia", ""), "observaciones": t.get("observaciones", "")})
    return t


def eliminar_tarifa(tid: str) -> None:
    with conn() as c:
        c.execute("DELETE FROM tarifas WHERE id=?", (tid,))


# ---- clave de edición (PBKDF2). Si existe la variable AUTORRETENCION_PIN se usa esa en lugar de la guardada.
def _hash(pin: str, salt: bytes) -> str:
    return hashlib.pbkdf2_hmac("sha256", pin.encode(), salt, 200_000).hex()


def pin_configurado() -> bool:
    return bool(os.environ.get("AUTORRETENCION_PIN")) or _meta("pin") is not None


def _meta(k: str):
    with conn() as c:
        r = c.execute("SELECT v FROM meta WHERE k=?", (k,)).fetchone()
        return r["v"] if r else None


def crear_pin(pin: str) -> None:
    salt = os.urandom(16)
    with conn() as c:
        c.execute("INSERT OR REPLACE INTO meta VALUES('pin',?)", (salt.hex() + ":" + _hash(pin, salt),))


def verificar_pin(pin: str) -> bool:
    env = os.environ.get("AUTORRETENCION_PIN")
    if env:
        return hmac.compare_digest(pin, env)
    v = _meta("pin")
    if not v:
        return False
    salt, h = v.split(":")
    return hmac.compare_digest(_hash(pin, bytes.fromhex(salt)), h)
