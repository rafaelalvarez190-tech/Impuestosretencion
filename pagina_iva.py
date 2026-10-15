"""Pestaña «Impuesto sobre las ventas (IVA)»: cuentas 2408 + total de la clase 4. Pide su propio auxiliar."""
from datetime import date, timedelta

import pandas as pd
import streamlit as st

import core
import db
from core import fmt
from excel_iva import exportar_iva

M = "iva"
STEPS = ["1. Cargar auxiliar de IVA", "2. Liquidación de IVA"]
TIPO = {"generado": "IVA generado", "descontable": "IVA descontable", "otro": "Otro"}


def _tarifa_txt(c) -> str:
    return ", ".join(f"{k:f}" if k != core.SIN_TARIFA else k for k in c.por_tarifa)


def render() -> None:
    ss = st.session_state
    if "iva_goto" in ss:
        ss["iva_step"] = ss.pop("iva_goto")
    ss.setdefault("iva_step", STEPS[0])
    step = st.radio("Paso IVA", STEPS, horizontal=True, key="iva_step", label_visibility="collapsed")

    # ============================================================ 1. CARGA
    if step == STEPS[0]:
        st.subheader("Cargar auxiliar contable para IVA")
        st.caption("Este módulo tiene su propio archivo: cargue el auxiliar que contiene las cuentas **2408** y la **clase 4** "
                   "(obligatorias: empresa, fecha, codigocuenta, nombrecuenta, valor). No se mezcla con el archivo de retenciones.")
        files = st.file_uploader("CSV o Excel del auxiliar de IVA", type=["csv", "xlsx", "xlsm"], accept_multiple_files=True,
                                 key="iva_files")
        if files and st.button("Procesar archivos de IVA y continuar", type="primary", key="iva_procesar"):
            flash, cargados, det = [], 0, set()
            for f in files:
                try:
                    rows, errs = core.normalize_rows(core.read_table(f.getvalue(), f.name))
                except Exception as e:  # noqa: BLE001
                    st.error(f"**{f.name}**: no se pudo leer ({e}).")
                    continue
                if not rows:
                    st.error(f"**{f.name}**: {errs[0] if errs else 'sin movimientos válidos'}")
                    continue
                r = db.upsert_movimientos(rows, M)
                cargados += 1
                det |= {m["empresa"] for m in rows}
                n2408 = sum(core.es_cuenta_iva(m["codigocuenta"]) for m in rows)
                n4 = sum(core.es_cuenta_ingreso(m["codigocuenta"]) for m in rows)
                msg = (f"**{f.name}**: {len(rows)} filas leídas ({n4} de clase 4, {n2408} de cuentas 2408) → **{r['added']}** nuevas, "
                       f"{r['updated']} actualizadas, {r['unchanged']} ya existentes (no duplicadas)."
                       + (f" {r['repeated']} repetidas dentro del archivo." if r["repeated"] else "")
                       + (f"\n\n{len(errs)} fila(s) omitidas: " + " | ".join(errs[:5]) if errs else ""))
                flash.append(("warning" if errs else "success", msg))
                if not n2408:
                    flash.append(("warning", f"**{f.name}** no trae ninguna cuenta 2408: verifique que sea el auxiliar de IVA."))
            if cargados:
                flash.append(("info", "Empresas detectadas: **" + ", ".join(sorted(det)) + "**."))
                ss["iva_flash"] = flash
                ss["iva_goto"] = STEPS[1]
                st.rerun()
        st.subheader("Movimientos de IVA almacenados")
        inv = db.inventario(M)
        if inv:
            st.dataframe(pd.DataFrame(inv).rename(columns={"empresa": "Empresa", "mes": "Mes", "movimientos": "Movimientos",
                                                             "ingresos": "De ingresos (4)", "iva": "IVA (2408)", "reteiva": "IVA retenido (135517)"}),
                         hide_index=True, width="stretch")
            if st.button("Continuar a la liquidación de IVA →", key="iva_cont"):
                ss["iva_goto"] = STEPS[1]; st.rerun()
            with st.expander("Borrar los movimientos de IVA"):
                if st.checkbox("Confirmo que quiero borrar todos los movimientos de este módulo", key="iva_conf_borrar") \
                        and st.button("Borrar", key="iva_borrar"):
                    db.borrar_movimientos(M); st.rerun()
        else:
            st.info("Aún no hay movimientos de IVA cargados.")
        return

    # ============================================================ 2. LIQUIDACIÓN
    for kind, msg in ss.pop("iva_flash", []):
        getattr(st, kind)(msg)
    todos = db.movimientos(modulo=M)
    empresas = sorted({m["empresa"] for m in todos})
    if not empresas:
        st.info("Primero cargue el auxiliar de IVA (paso 1).")
        return
    st.subheader("Impuesto sobre las ventas (IVA)")
    c1, c2, c3, c4 = st.columns([2, 1.3, 1.3, 1.3])
    empresa = c1.selectbox("Empresa", empresas, key="iva_empresa")
    meses = sorted({m["fecha"][:7] for m in todos if m["empresa"] == empresa}, reverse=True)
    mes = c2.selectbox("Período", meses + ["Rango personalizado"], key="iva_mes")
    if mes == "Rango personalizado":
        fechas = [m["fecha"] for m in todos if m["empresa"] == empresa]
        d_def, h_def = date.fromisoformat(min(fechas)), date.fromisoformat(max(fechas))
    else:
        y, mm = map(int, mes.split("-"))
        d_def, h_def = date(y, mm, 1), date(y + (mm == 12), mm % 12 + 1, 1) - timedelta(days=1)
    desde = c3.date_input("Desde", d_def, key=f"iva_d{empresa}{mes}")
    hasta = c4.date_input("Hasta", h_def, key=f"iva_h{empresa}{mes}")
    invertir = st.checkbox("Presentar ingresos e IVA generado en positivo (invertir el signo crédito del auxiliar)", value=True,
                           key="iva_invertir")
    nits_txt = st.text_input("NIT excluidos del cálculo (separe con comas; sin dígito de verificación)",
                             value=db.get_meta("nits_excluidos_iva", core.NITS_EXCLUIDOS_DEFECTO), key="iva_nits")
    if ", ".join(core.parse_nits(nits_txt)) != db.get_meta("nits_excluidos_iva", core.NITS_EXCLUIDOS_DEFECTO):
        db.set_meta("nits_excluidos_iva", ", ".join(core.parse_nits(nits_txt)))
    if hasta < desde:
        st.error("La fecha final es anterior a la inicial.")
        return

    nits = core.parse_nits(nits_txt)
    m_todos = core.filtrar(todos, empresa, desde.isoformat(), hasta.isoformat())
    movs, excluidos = core.separar_excluidos(m_todos, nits)          # solo `movs` entra al cálculo
    ingresos = core.build_base(movs, invertir)                        # clase 4: neto por cuenta
    iva = core.build_iva(movs, invertir)                              # cuentas 2408
    t = core.totales_iva(iva)
    reteiva = core.build_reteiva(movs)                                # cuentas 135517 (IVA que le retienen)
    rt = core.totales_reteiva(reteiva)
    liq = core.liquidar_iva(iva, reteiva, invertir)

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Total ingresos (clase 4, neto)", fmt(ingresos.base), f"{len(ingresos.cuentas)} cuentas", delta_color="off")
    k2.metric("IVA generado (2408)", fmt(liq["generado"]), "créditos", delta_color="off")
    k3.metric("Descontable / devoluciones (−)", fmt(liq["descontable"]), "débitos 2408", delta_color="off")
    k4.metric("IVA retenido (135517) (−)", fmt(liq["retenido"]), "retenciones que le aplican", delta_color="off")
    k5.metric("IVA por pagar", fmt(liq["por_pagar"]), "generado − descontable − retenido", delta_color="off")
    st.caption(f"**Empresa:** {empresa} · **Período:** {desde} a {hasta} · Liquidación preliminar sujeta a revisión.")

    # ---- clase 4
    st.subheader("Ingresos (clase 4) — neto por cuenta")
    if ingresos.cuentas:
        df = pd.DataFrame([{"Código": c.codigo, "Cuenta": c.nombre, "Mov.": c.n, "Suma de valor (auxiliar)": float(c.raw),
                            "Valor neto": float(c.valor)} for c in ingresos.cuentas]
                          + [{"Código": "TOTAL", "Cuenta": "Total ingresos (clase 4)", "Mov.": ingresos.n_movs,
                              "Suma de valor (auxiliar)": float(ingresos.base_raw), "Valor neto": float(ingresos.base)}])
        st.dataframe(df.style.format({"Suma de valor (auxiliar)": "{:,.2f}", "Valor neto": "{:,.2f}"})
                     .apply(lambda r: ["font-weight:700; background-color:#eef2f7; color:#1c2430" if r["Código"] == "TOTAL" else "" for _ in r], axis=1),
                     hide_index=True, width="stretch")
        for c in ingresos.cuentas:
            with st.expander(f"{c.codigo} · {c.nombre} — {c.n} movimientos · neto {fmt(c.valor)}"):
                st.dataframe(pd.DataFrame([{"Fila": m["fila"], "Id": m["id"], "Fecha": m["fecha"],
                                            "Doc.": f"{m['tipodocumento']} {m['nrodocumento']}".strip(), "Tercero": m["tercero"],
                                            "Descripción": m["descripcion"], "C. costo": m["centrocosto"], "valor": float(m["valor"])}
                                           for m in c.movs]), hide_index=True, width="stretch")
    else:
        st.info("No hay cuentas de clase 4 en el período.")

    # ---- 2408
    st.subheader("IVA (cuentas 2408) — clasificación")
    st.caption("IVA generado = movimientos crédito (`valor` negativo). Descontable / devoluciones / débitos = movimientos débito (se muestran en negativo y restan). "
               "Base y tarifa vienen de `baseimpuesto` y `porcentajeimpuesto` del auxiliar.")
    if iva:
        z = (-1 if invertir else 1)
        filas = [{"Código": c.codigo, "Cuenta": c.nombre, "Clasificación": TIPO[c.tipo], "Tarifa %": _tarifa_txt(c),
                  "Base (créditos)": float(c.base_cred), "IVA generado": float(c.generado), "Base (débitos)": float(c.base_deb),
                  "Débitos (−)": -float(c.debitos), "Neto": float(c.neto), "Mov.": c.n} for c in iva]
        filas.append({"Código": "TOTAL", "Cuenta": "Total IVA 2408", "Clasificación": "", "Tarifa %": "",
                      "Base (créditos)": float(t["base_cred"]), "IVA generado": float(t["generado"]), "Base (débitos)": float(t["base_deb"]),
                      "Débitos (−)": -float(t["debitos"]), "Neto": float(t["generado"] - t["debitos"]), "Mov.": None})
        fm = {k: "{:,.2f}" for k in ["Base (créditos)", "IVA generado", "Base (débitos)", "Débitos (−)", "Neto"]}
        st.markdown("**Por cuenta**")
        st.dataframe(pd.DataFrame(filas).style.format(fm, na_rep="")
                     .apply(lambda r: ["font-weight:700; background-color:#eef2f7; color:#1c2430" if r["Código"] == "TOTAL" else "" for _ in r], axis=1),
                     hide_index=True, width="stretch")
        st.markdown("**Por tarifa**")
        pt = [{"Tarifa %": f"{r['tarifa']:f}" if r["tarifa"] != core.SIN_TARIFA else r["tarifa"], "Base (créditos)": float(r["base_cred"]),
               "IVA generado": float(r["generado"]), "Base (débitos)": float(r["base_deb"]), "Débitos (−)": -float(r["debitos"]),
               "Neto": float(r["neto"])} for r in core.iva_por_tarifa(iva)]
        pt.append({"Tarifa %": "TOTAL", "Base (créditos)": float(t["base_cred"]), "IVA generado": float(t["generado"]),
                   "Base (débitos)": float(t["base_deb"]), "Débitos (−)": -float(t["debitos"]), "Neto": float(t["generado"] - t["debitos"])})
        st.dataframe(pd.DataFrame(pt).style.format(fm)
                     .apply(lambda r: ["font-weight:700; background-color:#eef2f7; color:#1c2430" if r["Tarifa %"] == "TOTAL" else "" for _ in r], axis=1),
                     hide_index=True, width="stretch")
        for c in iva:
            with st.expander(f"{c.codigo} · {c.nombre} — {c.n} movimientos"):
                st.dataframe(pd.DataFrame([{"Fila": m["fila"], "Id": m["id"], "Fecha": m["fecha"],
                                            "Doc.": f"{m['tipodocumento']} {m['nrodocumento']}".strip(), "Tercero": m["tercero"],
                                            "Descripción": m["descripcion"], "Base": float(m.get("baseimpuesto", 0)),
                                            "%": float(m.get("porcentajeimpuesto", 0)), "valor": float(m["valor"])}
                                           for m in c.movs]), hide_index=True, width="stretch")
    else:
        st.info("No hay movimientos en cuentas 2408 para esta empresa y período.")

    # ---- 135517
    st.subheader("IVA retenido (cuenta 135517) — retenciones que le aplican")
    st.caption("Retención aplicada = movimientos débito (`valor` positivo). Reversiones = movimientos crédito (se muestran en negativo y restan). "
               "El neto de esta cuenta se resta del IVA a pagar.")
    if reteiva:
        filas = [{"Código": c.codigo, "Cuenta": c.nombre, "Tarifa %": ", ".join(f"{k:f}" if k != core.SIN_TARIFA else k for k in c.por_tarifa),
                  "Base (aplicadas)": float(c.base_apl), "Retención aplicada": float(c.aplicada), "Reversiones (−)": -float(c.reversion),
                  "Neto": float(c.neto), "Mov.": c.n} for c in reteiva]
        filas.append({"Código": "TOTAL", "Cuenta": "Total IVA retenido (135517)", "Tarifa %": "", "Base (aplicadas)": float(rt["base_apl"]),
                      "Retención aplicada": float(rt["aplicada"]), "Reversiones (−)": -float(rt["reversion"]), "Neto": float(rt["neto"]), "Mov.": None})
        st.dataframe(pd.DataFrame(filas).style.format({k: "{:,.2f}" for k in ["Base (aplicadas)", "Retención aplicada", "Reversiones (−)", "Neto"]}, na_rep="")
                     .apply(lambda r: ["font-weight:700; background-color:#eef2f7; color:#1c2430" if r["Código"] == "TOTAL" else "" for _ in r], axis=1),
                     hide_index=True, width="stretch")
        for c in reteiva:
            with st.expander(f"{c.codigo} · {c.nombre} — {c.n} movimientos"):
                st.dataframe(pd.DataFrame([{"Fila": m["fila"], "Id": m["id"], "Fecha": m["fecha"],
                                            "Doc.": f"{m['tipodocumento']} {m['nrodocumento']}".strip(), "Tercero": m["tercero"],
                                            "Descripción": m["descripcion"], "Base": float(m.get("baseimpuesto", 0)),
                                            "%": float(m.get("porcentajeimpuesto", 0)), "valor": float(m["valor"])}
                                           for m in c.movs]), hide_index=True, width="stretch")
    else:
        st.info("No hay movimientos en cuentas 135517 para esta empresa y período.")

    # ---- liquidación
    st.subheader("Liquidación del IVA (preliminar)")
    liq_df = pd.DataFrame([
        {"Concepto": "IVA generado (2408, créditos)", "Valor": float(liq["generado"])},
        {"Concepto": "(−) IVA descontable / devoluciones / otros débitos (2408)", "Valor": float(liq["descontable"])},
        {"Concepto": "= IVA neto 2408", "Valor": float(liq["neto_2408"])},
        {"Concepto": "(−) IVA retenido (135517)", "Valor": float(liq["retenido"])},
        {"Concepto": "= IVA POR PAGAR", "Valor": float(liq["por_pagar"])},
    ])
    st.dataframe(liq_df.style.format({"Valor": "{:,.2f}"}).apply(
        lambda r: ["font-weight:700; background-color:#eef2f7; color:#1c2430" if r["Concepto"].startswith("=") else "" for _ in r], axis=1),
        hide_index=True, width="stretch")
    if not invertir:
        st.caption("Con el signo del auxiliar: el IVA generado es crédito (negativo) y el IVA por pagar sale negativo (pasivo).")

    if excluidos:
        with st.expander(f"Movimientos excluidos del cálculo por NIT ({len(excluidos)}) — {', '.join(nits)}"):
            st.dataframe(pd.DataFrame([{"Fila": m["fila"], "Id": m["id"], "Fecha": m["fecha"], "Cuenta": m["codigocuenta"],
                                        "Nombre cuenta": m["nombrecuenta"], "NIT": m["identificacion"],
                                        "Descripción": m["descripcion"], "valor": float(m["valor"])} for m in excluidos]),
                         hide_index=True, width="stretch")

    st.subheader("Controles y validaciones")
    checks = core.ejecutar_controles_iva(movs, ingresos, iva, len(todos), len({m["key"] for m in todos}), excluidos, nits, reteiva)
    for c in checks:
        st.markdown(f"{ {'ok': '✅', 'warn': '⚠️', 'error': '❌'}[c['nivel']] } **{c['titulo']}** — {c['detalle']}")

    st.download_button("Exportar IVA a Excel", exportar_iva(empresa, desde.isoformat(), hasta.isoformat(), ingresos, iva, checks, excluidos, nits, reteiva),
                       file_name=f"iva_{empresa}_{desde}_{hasta}.xlsx", type="primary", key="iva_export",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
