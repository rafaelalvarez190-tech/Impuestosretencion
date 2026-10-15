"""Exportación del informe a Excel (openpyxl)."""
from __future__ import annotations

import io
from decimal import Decimal

from openpyxl import Workbook
from openpyxl.styles import Font

from core import Base, CuentaRet, ResTarifa, calc_autorretencion, totales_retenciones, GRUPOS_RETENCION

NUM = "#,##0.00"
TIPO = {"practicada": "Retención practicada", "autorretencion": "Autorretención contabilizada", "pago": "Pago / cierre"}


def _f(x):
    return float(x) if isinstance(x, Decimal) else x


def _sheet(wb, title, rows, widths, num_cols=()):
    ws = wb.create_sheet(title)
    for r in rows:
        ws.append([_f(x) for x in r])
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w
    for c in ws[1]:
        c.font = Font(bold=True)
    for col in num_cols:
        for cell in ws[col][1:]:
            if isinstance(cell.value, (int, float)):
                cell.number_format = NUM
    return ws


def exportar(empresa: str, desde: str, hasta: str, base: Base, tr: ResTarifa, checks: list[dict],
             retenciones: list[CuentaRet] | None = None) -> bytes:
    ret = retenciones or []
    ok = tr.status == "ok"
    pct = tr.tarifa if ok else None
    auto = calc_autorretencion(base.base, pct) if ok else None
    tot = totales_retenciones(ret)
    wb = Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("Resumen")
    filas = [
        ["LIQUIDACIÓN PRELIMINAR DE AUTORRETENCIÓN — sujeta a revisión"], [],
        ["Empresa", empresa], ["Período de liquidación", f"{desde} a {hasta}"],
        ["Tarifa de autorretención (Maestro de empresas) %", _f(pct) if ok else "No configurada"],
        ["Signo de presentación", "Invertido (ingresos positivos)" if base.invertir_signo else "Tal cual el auxiliar"], [],
        ["Base total de autorretención", _f(base.base)],
        ["Porcentaje aplicado (%)", _f(pct) if ok else None],
        ["Autorretención calculada", _f(auto) if ok else "No calculada"], [],
        ["RETENCIONES PRACTICADAS EN EL PERÍODO (cuentas 2365 y 2367)"],
        ["2365 Retención en la fuente — causada (créditos)", _f(tot["2365"]["creditos"])],
        ["2365 Retención en la fuente — débitos (cierres/devoluciones)", _f(tot["2365"]["debitos"])],
        ["2367 IVA retenido — causada (créditos)", _f(tot["2367"]["creditos"])],
        ["2367 IVA retenido — débitos (cierres/devoluciones)", _f(tot["2367"]["debitos"])],
        ["Total retenciones practicadas (causadas)", _f(tot["total"]["creditos"])], [],
        ["La parametrización de la tarifa no sustituye la validación de las reglas tributarias aplicables a la empresa y al período."],
    ]
    for f in filas:
        ws.append(f)
    ws["A1"].font = Font(bold=True, size=13)
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 58, 46
    for r in (8, 10, 13, 14, 15, 16, 17):
        if isinstance(ws[f"B{r}"].value, float):
            ws[f"B{r}"].number_format = NUM
    for r in (8, 10, 12, 17):
        ws[f"A{r}"].font = Font(bold=True)

    det = [["Código", "Cuenta", "Movimientos", "Suma valor (auxiliar)", "Valor neto", "Autorretención"]]
    det += [[c.codigo, c.nombre, c.n, c.raw, c.valor, calc_autorretencion(c.valor, pct) if ok else None] for c in base.cuentas]
    det.append(["TOTAL", "Base total", base.n_movs, base.base_raw, base.base, auto])
    w = _sheet(wb, "Detalle por cuenta", det, [14, 40, 12, 22, 20, 18], "DEF")
    for c in w[w.max_row]:
        c.font = Font(bold=True)

    # ---- retenciones 2365 / 2367
    rt = [["Grupo", "Código", "Cuenta", "Tipo de cuenta", "Tarifa(s) %", "Base (movimientos de causación)",
           "Retención causada (créditos)", "Débitos (cierres / pagos / devoluciones)", "Neto", "Movimientos"]]
    for g, nombre in GRUPOS_RETENCION.items():
        cs = [c for c in ret if c.grupo == g]
        for c in cs:
            rt.append([f"{g} {nombre}", c.codigo, c.nombre, TIPO[c.tipo], ", ".join(f"{t:f}" for t in sorted(c.tarifas)),
                       c.base, c.creditos, c.debitos, c.neto, c.n])
        t = tot[g]
        rt.append([f"Subtotal {g} (solo retenciones practicadas)", "", "", "", "", t["base"], t["creditos"], t["debitos"], t["neto"], ""])
    t = tot["total"]
    rt.append(["TOTAL RETENCIONES PRACTICADAS", "", "", "", "", t["base"], t["creditos"], t["debitos"], t["neto"], ""])
    w = _sheet(wb, "Retenciones 2365-2367", rt, [44, 14, 40, 26, 12, 24, 24, 26, 20, 12], "FGHI")
    for row in w.iter_rows(min_row=2):
        if str(row[0].value).startswith(("Subtotal", "TOTAL")):
            for c in row:
                c.font = Font(bold=True)

    pp = [["Código", "Cuenta", "Tipo de persona", "Base", "Retención causada"]]
    for c in ret:
        if c.tipo == "practicada":
            for per, (b, r) in sorted(c.por_persona.items()):
                pp.append([c.codigo, c.nombre, per, b, r])
    _sheet(wb, "Retenciones por tipo persona", pp, [14, 44, 18, 22, 22], "DE")

    mr = [["Cuenta", "Nombre cuenta", "Fila archivo", "Id auxiliar", "Fecha", "Tipo doc.", "Nro doc.", "Identificación", "Tercero",
           "Tipo persona", "Descripción", "Base impuesto", "% impuesto", "valor"]]
    for c in ret:
        mr += [[m["codigocuenta"], m["nombrecuenta"], m["fila"], m["id"], m["fecha"], m["tipodocumento"], m["nrodocumento"],
                m["identificacion"], m["tercero"], m.get("tipo_persona", ""), m["descripcion"], m.get("baseimpuesto"),
                m.get("porcentajeimpuesto"), m["valor"]] for m in c.movs]
    _sheet(wb, "Mov. retenciones", mr, [12, 30, 10, 10, 12, 8, 10, 14, 30, 12, 40, 18, 10, 18], "LN")

    mv = [["Cuenta", "Nombre cuenta", "Fila archivo", "Id auxiliar", "Fecha", "Tipo doc.", "Nro doc.", "Identificación",
           "Tercero", "Descripción", "Centro costo", "valor"]]
    for c in base.cuentas:
        mv += [[m["codigocuenta"], m["nombrecuenta"], m["fila"], m["id"], m["fecha"], m["tipodocumento"], m["nrodocumento"],
                m["identificacion"], m["tercero"], m["descripcion"], m["centrocosto"], m["valor"]] for m in c.movs]
    _sheet(wb, "Mov. ingresos", mv, [12, 30, 10, 10, 12, 8, 10, 14, 30, 40, 20, 18], "L")

    ng = [["Código", "Cuenta", "Valor neto"]] + [[c.codigo, c.nombre, c.valor] for c in base.negativas]
    ng += [[], ["Efecto de las negativas en la base", "", base.efecto_negativas], ["Base sin cuentas negativas", "", base.base_sin_negativas]]
    _sheet(wb, "Cuentas negativas", ng, [34, 40, 20], "C")

    vl = [["Control", "Resultado", "Detalle"]] + [[c["titulo"], "OK" if c["ok"] else ("Revisar" if c["nivel"] == "warn" else "Falla"), c["detalle"]] for c in checks]
    _sheet(wb, "Validaciones", vl, [48, 10, 100])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
