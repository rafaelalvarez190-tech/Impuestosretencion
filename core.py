"""Núcleo de cálculo de la autorretención (sin dependencias de Streamlit).

Todo el dinero se maneja con Decimal. La base sale EXCLUSIVAMENTE de la columna `valor`.
"""
from __future__ import annotations

import csv
import hashlib
import io
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from decimal import ROUND_HALF_UP, Decimal
from typing import Any, Iterable

PREFIJO_INGRESOS = "4"
CENT = Decimal("0.01")
REQUIRED = ["empresa", "fecha", "codigocuenta", "nombrecuenta", "valor"]


# ---------------------------------------------------------------- números y fechas
def parse_number(v: Any) -> Decimal:
    """Acepta Decimal/int/float o texto es-CO ('1.234,56', '182024,87', '(100,00)')."""
    if v is None or v == "":
        return Decimal(0)
    if isinstance(v, Decimal):
        return v
    if isinstance(v, (int, float)):
        return Decimal(str(v))
    s = re.sub(r"[\s$]", "", str(v))
    if not s:
        return Decimal(0)
    neg = s.startswith("(") and s.endswith(")")
    if neg:
        s = s[1:-1]
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    d = Decimal(s)  # lanza InvalidOperation si no es número
    return -d if neg else d


def q2(d: Decimal) -> Decimal:
    return d.quantize(CENT, rounding=ROUND_HALF_UP)


def normalize_date(v: Any) -> str:
    if v is None or v == "":
        return ""
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, (int, float)):  # serial de Excel
        return (date(1899, 12, 30) + timedelta(days=int(v))).isoformat()
    s = str(v).strip()
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return f"{m[1]}-{m[2]}-{m[3]}"
    m = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{4})", s)
    if m:
        return f"{m[3]}-{int(m[2]):02d}-{int(m[1]):02d}"
    return ""


# ---------------------------------------------------------------- lectura de archivos
def read_table(data: bytes, filename: str) -> list[list[Any]]:
    if filename.lower().endswith((".xlsx", ".xlsm")):
        from openpyxl import load_workbook

        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        return [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)
                if any(c not in (None, "") for c in r)]
    try:
        text = data.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = data.decode("cp1252")
    first = text.split("\n", 1)[0]
    delim = ";" if first.count(";") > first.count(",") else ","
    return [r for r in csv.reader(io.StringIO(text), delimiter=delim) if any(c.strip() for c in r)]


def _norm_header(h: Any) -> str:
    s = unicodedata.normalize("NFD", str(h).strip().lower())
    return re.sub(r"[\s_]+", "", "".join(c for c in s if unicodedata.category(c) != "Mn"))


def normalize_rows(table: list[list[Any]]) -> tuple[list[dict], list[str]]:
    errors: list[str] = []
    if not table:
        return [], ["El archivo está vacío."]
    headers = [_norm_header(h) for h in table[0]]
    missing = [h for h in REQUIRED if h not in headers]
    if missing:
        return [], [f"Faltan columnas obligatorias: {', '.join(missing)}."]
    ix = {h: i for i, h in enumerate(headers)}

    def get(r, h):
        i = ix.get(h)
        return "" if i is None or i >= len(r) or r[i] is None else r[i]

    rows = []
    for n, r in enumerate(table[1:], start=2):
        codigo, empresa = str(get(r, "codigocuenta")).strip(), str(get(r, "empresa")).strip()
        fecha = normalize_date(get(r, "fecha"))
        if not codigo or not empresa:
            errors.append(f"Fila {n}: sin empresa o código de cuenta.")
            continue
        if not fecha:
            errors.append(f"Fila {n}: fecha inválida ({get(r, 'fecha')}).")
            continue
        try:
            valor = parse_number(get(r, "valor"))
            debito, credito = parse_number(get(r, "debito")), parse_number(get(r, "credito"))
        except Exception:
            errors.append(f"Fila {n}: valor no numérico ({get(r, 'valor')}).")
            continue
        s = lambda h: str(get(r, h)).strip()  # noqa: E731
        m = dict(
            id=s("id"), empresa=empresa, centrocosto=s("centrocosto"), fecha=fecha,
            tipodocumento=s("tipodocumento"), nrodocumento=s("nrodocumento"),
            codigocuenta=codigo, nombrecuenta=s("nombrecuenta"), descripcion=s("descripcion"),
            factura=s("factura"), identificacion=s("identificacion"), tercero=s("tercero"),
            valor=valor, debito=debito, credito=credito, fila=n,  # débito/crédito solo informativos
        )
        m["key"] = f"{empresa}|{m['id']}" if m["id"] else f"{empresa}|h:{fingerprint(m)}"
        rows.append(m)
    return rows, errors


def fingerprint(m: dict) -> str:
    raw = "¦".join(str(m[k]) for k in ("empresa", "fecha", "tipodocumento", "nrodocumento",
                                      "codigocuenta", "identificacion", "descripcion", "factura", "valor"))
    return hashlib.sha1(raw.encode()).hexdigest()[:16]


def es_cuenta_ingreso(codigo: str) -> bool:
    return str(codigo).strip().startswith(PREFIJO_INGRESOS)


# ---------------------------------------------------------------- base
@dataclass
class Cuenta:
    codigo: str
    nombre: str
    raw: Decimal = Decimal(0)       # suma de `valor` tal cual el auxiliar
    valor: Decimal = Decimal(0)     # valor presentado (con o sin inversión de signo)
    movs: list[dict] = field(default_factory=list)

    @property
    def n(self) -> int:
        return len(self.movs)


@dataclass
class Base:
    cuentas: list[Cuenta]
    invertir_signo: bool
    excluidos_no4: int

    @property
    def base(self) -> Decimal:
        return sum((c.valor for c in self.cuentas), Decimal(0))

    @property
    def base_raw(self) -> Decimal:
        return sum((c.raw for c in self.cuentas), Decimal(0))

    @property
    def n_movs(self) -> int:
        return sum(c.n for c in self.cuentas)

    @property
    def negativas(self) -> list[Cuenta]:
        return [c for c in self.cuentas if c.valor < 0]

    @property
    def efecto_negativas(self) -> Decimal:
        return sum((c.valor for c in self.negativas), Decimal(0))

    @property
    def base_sin_negativas(self) -> Decimal:
        return self.base - self.efecto_negativas


def build_base(movs: Iterable[dict], invertir_signo: bool = True) -> Base:
    """Base = suma neta (con signo) de `valor` de las cuentas que empiezan por 4."""
    acc: dict[str, Cuenta] = {}
    excl = 0
    for m in movs:
        if not es_cuenta_ingreso(m["codigocuenta"]):
            excl += 1
            continue
        c = acc.setdefault(m["codigocuenta"], Cuenta(m["codigocuenta"], m["nombrecuenta"]))
        c.raw += m["valor"]
        c.movs.append(m)
    sgn = -1 if invertir_signo else 1
    for c in acc.values():
        c.valor = c.raw * sgn
    return Base(sorted(acc.values(), key=lambda c: c.codigo), invertir_signo, excl)


def filtrar(movs: Iterable[dict], empresa: str, desde: str, hasta: str) -> list[dict]:
    return [m for m in movs if m["empresa"] == empresa and desde <= m["fecha"] <= hasta]


def calc_autorretencion(base: Decimal, porcentaje: Decimal | float | str) -> Decimal:
    return q2(base * Decimal(str(porcentaje)) / Decimal(100))


# ---------------------------------------------------------------- tarifas
def validar_tarifa(t: dict) -> list[str]:
    e = []
    if not str(t.get("empresa", "")).strip():
        e.append("Seleccione la empresa.")
    try:
        p = Decimal(str(t.get("porcentaje")))
        if not (0 <= p <= 100):
            e.append("El porcentaje debe estar entre 0 y 100.")
    except Exception:
        e.append("El porcentaje es obligatorio.")
    if not t.get("desde"):
        e.append("La fecha inicial de vigencia es obligatoria.")
    if t.get("desde") and t.get("hasta") and t["hasta"] < t["desde"]:
        e.append("La fecha final no puede ser anterior a la inicial.")
    return e


def _se_superponen(a: dict, b: dict) -> bool:
    return (not b.get("hasta") or a["desde"] <= b["hasta"]) and (not a.get("hasta") or b["desde"] <= a["hasta"])


def buscar_solapes(tarifas: list[dict], t: dict) -> list[dict]:
    if not t.get("activa", True):
        return []
    return [o for o in tarifas if o["id"] != t.get("id") and o["empresa"] == t["empresa"]
            and o["activa"] and _se_superponen(o, t)]


@dataclass
class ResTarifa:
    status: str  # ok | none | partial | ambiguous
    tarifa: dict | None = None
    candidatas: list[dict] = field(default_factory=list)
    msg: str = ""


def resolver_tarifa(tarifas: list[dict], empresa: str, desde: str, hasta: str) -> ResTarifa:
    """Solo hay cálculo (status 'ok') si UNA tarifa activa cubre todo el período."""
    act = [t for t in tarifas if t["empresa"] == empresa and t["activa"]]
    cubre = [t for t in act if t["desde"] <= desde and (not t.get("hasta") or t["hasta"] >= hasta)]
    if len(cubre) == 1:
        return ResTarifa("ok", cubre[0], cubre)
    if len(cubre) > 1:
        return ResTarifa("ambiguous", None, cubre,
                         "Hay varias tarifas activas con vigencias superpuestas para este período. Corrija la configuración.")
    parcial = [t for t in act if t["desde"] <= hasta and (not t.get("hasta") or t["hasta"] >= desde)]
    if parcial:
        return ResTarifa("partial", None, parcial,
                         "Ninguna tarifa activa cubre todo el período seleccionado (solo parcialmente). "
                         "Complete la configuración o ajuste el período.")
    return ResTarifa("none", None, [],
                     "No hay una tarifa de autorretención definida para esta empresa y período. "
                     "Configúrela para calcular.")


# ---------------------------------------------------------------- controles
def fmt(n: Decimal | float, dec: int = 2) -> str:
    s = f"{Decimal(n):,.{dec}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def ejecutar_controles(movs_periodo: list[dict], base: Base, tr: ResTarifa, total_movs: int, claves_unicas: int) -> list[dict]:
    out = []

    def add(titulo, ok, detalle, nivel="error"):
        out.append(dict(titulo=titulo, ok=ok, detalle=detalle, nivel="ok" if ok else nivel))

    no4 = [c.codigo for c in base.cuentas if not es_cuenta_ingreso(c.codigo)]
    add("Todas las cuentas incluidas comienzan por 4", not no4,
        f"Fuera de regla: {', '.join(no4)}" if no4 else
        f"{len(base.cuentas)} cuenta(s) de ingresos; {base.excluidos_no4} movimiento(s) de otras clases excluidos.")
    add("Base = suma neta de la columna `valor`", sum((m['valor'] for c in base.cuentas for m in c.movs), Decimal(0)) == base.base_raw,
        f"Suma de `valor` (signo del auxiliar): {fmt(base.base_raw)}. No se usan débito, crédito ni saldos.")
    neg = base.negativas
    add("Cuentas con saldo negativo revisadas", not neg,
        f"{len(neg)} cuenta(s) reducen la base en {fmt(abs(base.efecto_negativas))}. Base sin ellas: {fmt(base.base_sin_negativas)}."
        if neg else "Ninguna cuenta de ingresos con saldo negativo.", "warn")
    add("Tarifa vigente aplicada al período", tr.status == "ok",
        f"{tr.tarifa['porcentaje']}% — vigencia {tr.tarifa['desde']} a {tr.tarifa.get('hasta') or 'sin fecha final'}."
        if tr.status == "ok" else tr.msg)
    add("Hay movimientos de la empresa en el período", bool(movs_periodo),
        f"{len(movs_periodo)} movimiento(s) en el rango." if movs_periodo else "No hay movimientos en el rango elegido.")
    add("Sin movimientos duplicados", total_movs == claves_unicas,
        f"{total_movs} movimiento(s) únicos almacenados (clave: empresa + id del auxiliar).")
    add("Base rastreable hasta los movimientos originales", True,
        "Cada cuenta enlaza a sus movimientos (fila del archivo e id del auxiliar).")
    return out
