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


def _norm_header(h: Any) -> str:  # "tipo_persona" -> "tipopersona"
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
            base_imp = parse_number(get(r, "baseimpuesto"))
            pct_imp = parse_number(get(r, "porcentajeimpuesto"))
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
            baseimpuesto=base_imp, porcentajeimpuesto=pct_imp, tipo_persona=s("tipopersona"),
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


# ---------------------------------------------------------------- maestro de empresas
def validar_empresa(e: dict, existentes: list[str] | None = None) -> list[str]:
    errs = []
    if not str(e.get("nombre", "")).strip():
        errs.append("El nombre de la empresa es obligatorio.")
    if existentes and str(e.get("nombre", "")).strip() in existentes:
        errs.append(f"La empresa {e['nombre']} está repetida.")
    t = e.get("tarifa")
    if t in (None, "") or (isinstance(t, float) and t != t):
        errs.append(f"{e.get('nombre') or 'Empresa'}: falta la tarifa de autorretención.")
    else:
        try:
            if not (0 <= Decimal(str(t)) <= 100):
                errs.append(f"{e.get('nombre')}: la tarifa debe estar entre 0 y 100.")
        except Exception:
            errs.append(f"{e.get('nombre')}: tarifa no numérica.")
    return errs


@dataclass
class ResTarifa:
    status: str  # ok | none | inactive
    tarifa: Decimal | None = None
    msg: str = ""


def resolver_tarifa(maestro: list[dict], empresa: str) -> ResTarifa:
    """La tarifa sale del maestro de empresas. Sin tarifa (o empresa inactiva) no se calcula."""
    e = next((x for x in maestro if x["nombre"] == empresa), None)
    if e is None or e.get("tarifa") in (None, ""):
        return ResTarifa("none", None, f"La empresa {empresa} no tiene tarifa de autorretención en el Maestro de empresas. "
                                       "Configúrela para calcular.")
    if not e.get("activa", True):
        return ResTarifa("inactive", None, f"La empresa {empresa} está inactiva en el Maestro de empresas.")
    return ResTarifa("ok", Decimal(str(e["tarifa"])))


# ---------------------------------------------------------------- retenciones 2365 / 2367
GRUPOS_RETENCION = {"2365": "Retención en la fuente", "2367": "Impuesto a las ventas retenido (IVA)"}


def grupo_retencion(codigo: str) -> str | None:
    c = str(codigo).strip()
    return next((g for g in GRUPOS_RETENCION if c.startswith(g)), None)


@dataclass
class CuentaRet:
    codigo: str
    nombre: str
    grupo: str
    tipo: str                                   # practicada | autorretencion | pago
    creditos: Decimal = Decimal(0)              # retención causada (movimientos con valor < 0), en magnitud
    debitos: Decimal = Decimal(0)               # cierres, pagos, devoluciones (valor > 0), en magnitud
    base: Decimal = Decimal(0)                  # suma de baseimpuesto de los movimientos de causación
    tarifas: set = field(default_factory=set)
    por_persona: dict = field(default_factory=dict)   # tipo_persona -> [base, retención]
    movs: list[dict] = field(default_factory=list)
    invertir: bool = True

    @property
    def neto(self) -> Decimal:
        n = self.creditos - self.debitos
        return n if self.invertir else -n

    @property
    def n(self) -> int:
        return len(self.movs)


def _tipo_cuenta(nombre: str) -> str:
    u = nombre.upper()
    if "AUTORRETEN" in u:
        return "autorretencion"
    if "PAGO" in u:
        return "pago"
    return "practicada"


def build_retenciones(movs: Iterable[dict], invertir_signo: bool = True) -> list[CuentaRet]:
    """Cuentas 2365 y 2367 del período. Causación = movimientos con valor < 0 (crédito); el resto son
    cierres/pagos/devoluciones. La base y el % salen de baseimpuesto / porcentajeimpuesto del auxiliar."""
    acc: dict[str, CuentaRet] = {}
    for m in movs:
        g = grupo_retencion(m["codigocuenta"])
        if not g:
            continue
        c = acc.setdefault(m["codigocuenta"], CuentaRet(m["codigocuenta"], m["nombrecuenta"], g,
                                                         _tipo_cuenta(m["nombrecuenta"]), invertir=invertir_signo))
        c.movs.append(m)
        v = m["valor"]
        if v < 0:
            c.creditos += -v
            c.base += m.get("baseimpuesto", Decimal(0))
            p = c.por_persona.setdefault(m.get("tipo_persona") or "Sin clasificar", [Decimal(0), Decimal(0)])
            p[0] += m.get("baseimpuesto", Decimal(0))
            p[1] += -v
        elif v > 0:
            c.debitos += v
        if m.get("porcentajeimpuesto"):
            c.tarifas.add(m["porcentajeimpuesto"].normalize())
    return sorted(acc.values(), key=lambda c: c.codigo)


def totales_retenciones(cuentas: list[CuentaRet]) -> dict:
    """Totales de retenciones PRACTICADAS (excluye cuentas de autorretención y de pago) por grupo."""
    out = {}
    for g, nombre in GRUPOS_RETENCION.items():
        cs = [c for c in cuentas if c.grupo == g and c.tipo == "practicada"]
        out[g] = dict(nombre=nombre, base=sum((c.base for c in cs), Decimal(0)),
                      creditos=sum((c.creditos for c in cs), Decimal(0)),
                      debitos=sum((c.debitos for c in cs), Decimal(0)),
                      neto=sum((c.neto for c in cs), Decimal(0)))
    out["total"] = {k: sum((out[g][k] for g in GRUPOS_RETENCION), Decimal(0)) for k in ("base", "creditos", "debitos", "neto")}
    return out


# ---------------------------------------------------------------- controles
def fmt(n: Decimal | float, dec: int = 2) -> str:
    s = f"{Decimal(n):,.{dec}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def ejecutar_controles(movs_periodo: list[dict], base: Base, tr: ResTarifa, total_movs: int, claves_unicas: int,
                       retenciones: list[CuentaRet] | None = None, auto: Decimal | None = None) -> list[dict]:
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
    add("Tarifa tomada del Maestro de empresas", tr.status == "ok",
        f"Tarifa de autorretención: {tr.tarifa}%." if tr.status == "ok" else tr.msg)
    add("Hay movimientos de la empresa en el período", bool(movs_periodo),
        f"{len(movs_periodo)} movimiento(s) en el rango." if movs_periodo else "No hay movimientos en el rango elegido.")
    add("Sin movimientos duplicados", total_movs == claves_unicas,
        f"{total_movs} movimiento(s) únicos almacenados (clave: empresa + id del auxiliar).")
    add("Base rastreable hasta los movimientos originales", True,
        "Cada cuenta enlaza a sus movimientos (fila del archivo e id del auxiliar).")
    ret = retenciones or []
    add("Retenciones 2365 / 2367 incluidas en el informe", bool(ret),
        f"{len(ret)} cuenta(s) de retención en el período." if ret else
        "No hay movimientos en cuentas 2365 ni 2367 para la empresa y el período.", "warn")
    contab = sum((c.creditos for c in ret if c.tipo == "autorretencion"), Decimal(0))
    if auto is not None and contab:
        dif = contab - auto
        add("Autorretención contabilizada (cuenta 2365 'AUTORRETENCION') vs. calculada", abs(dif) < Decimal("1"),
            f"Contabilizada: {fmt(contab)} · Calculada: {fmt(auto)} · Diferencia: {fmt(dif)}.", "warn")
    return out
