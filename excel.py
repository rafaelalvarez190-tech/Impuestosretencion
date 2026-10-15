"""Exportación del informe a Excel (openpyxl)."""
from __future__ import annotations

import io
from decimal import Decimal

from openpyxl import Workbook
from openpyxl.styles import Font

from core import Base, ResTarifa, calc_autorretencion

NUM = "#,##0.00"


def _sheet(wb, title, rows, widths, num_cols=()):
    ws = wb.create_sheet(title)
    for r in rows:
        ws.append([float(x) if isinstance(x, Decimal) else x for x in r])
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(1, i).column_letter].width = w
    for c in ws[1]:
        c.font = Font(bold=True)
    for col in num_cols:
        for cell in ws[col][1:]:
            if isinstance(cell.value, (int, float)):
                cell.number_format = NUM
    return ws


def exportar(empresa: str, desde: str, hasta: str, base: Base, tr: ResTarifa, checks: list[dict]) -> bytes:
    ok = tr.status == "ok"
    pct = Decimal(tr.tarifa["porcentaje"]) if ok else None
    auto = calc_autorretencion(base.base, pct) if ok else None
    wb = Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("Resumen")
    filas = [
        ["LIQUIDACIÓN PRELIMINAR DE AUTORRETENCIÓN — sujeta a revisión"], [],
        ["Empresa", empresa], ["Período de liquidación", f"{desde} a {hasta}"],
        ["Tarifa aplicada (%)", float(pct) if ok else "No configurada"],
        ["Vigencia de la tarifa", f"{tr.tarifa['desde']} a {tr.tarifa.get('hasta') or 'sin fecha final'}" if ok else ""],
        ["Referencia normativa", (tr.tarifa.get("referencia") or "") if ok else ""],
        ["Signo de presentación", "Invertido (ingresos positivos)" if base.invertir_signo else "Tal cual el auxiliar"], [],
        ["Base total de autorretención", float(base.base)],
        ["Porcentaje aplicado (%)", float(pct) if ok else None],
        ["Autorretención calculada", float(auto) if ok else "No calculada"], [],
        ["La parametrización de la tarifa no sustituye la validación de las reglas tributarias aplicables a la empresa y al período."],
    ]
    for f in filas:
        ws.append(f)
    ws["A1"].font = Font(bold=True, size=13)
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 34, 46
    for ref in ("B10", "B12"):
        if isinstance(ws[ref].value, float):
            ws[ref].number_format = NUM
    for ref in ("A10", "A12"):
        ws[ref].font = Font(bold=True)

    det = [["Código", "Cuenta", "Movimientos", "Suma valor (auxiliar)", "Valor neto", "Autorretención"]]
    det += [[c.codigo, c.nombre, c.n, c.raw, c.valor, calc_autorretencion(c.valor, pct) if ok else None] for c in base.cuentas]
    det.append(["TOTAL", "Base total", base.n_movs, base.base_raw, base.base, auto])
    w = _sheet(wb, "Detalle por cuenta", det, [14, 40, 12, 22, 20, 18], "DEF")
    for c in w[w.max_row]:
        c.font = Font(bold=True)

    mv = [["Cuenta", "Nombre cuenta", "Fila archivo", "Id auxiliar", "Fecha", "Tipo doc.", "Nro doc.", "Identificación",
           "Tercero", "Descripción", "Centro costo", "valor"]]
    for c in base.cuentas:
        mv += [[m["codigocuenta"], m["nombrecuenta"], m["fila"], m["id"], m["fecha"], m["tipodocumento"], m["nrodocumento"],
                m["identificacion"], m["tercero"], m["descripcion"], m["centrocosto"], m["valor"]] for m in c.movs]
    _sheet(wb, "Movimientos", mv, [12, 30, 10, 10, 12, 8, 10, 14, 30, 40, 20, 18], "L")

    ng = [["Código", "Cuenta", "Valor neto"]] + [[c.codigo, c.nombre, c.valor] for c in base.negativas]
    ng += [[], ["Efecto de las negativas en la base", "", base.efecto_negativas], ["Base sin cuentas negativas", "", base.base_sin_negativas]]
    _sheet(wb, "Cuentas negativas", ng, [34, 40, 20], "C")

    vl = [["Control", "Resultado", "Detalle"]] + [[c["titulo"], "OK" if c["ok"] else ("Revisar" if c["nivel"] == "warn" else "Falla"), c["detalle"]] for c in checks]
    _sheet(wb, "Validaciones", vl, [48, 10, 100])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
