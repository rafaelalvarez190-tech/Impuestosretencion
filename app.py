"""Retencion y/o autorretención sobre ingresos — app Streamlit.  Ejecutar:  streamlit run app.py"""
from datetime import date, timedelta
from decimal import Decimal

import pandas as pd
import streamlit as st

import core
import db
from core import fmt
from excel import exportar

st.set_page_config(page_title="Retencion y/o autorretención sobre ingresos", page_icon="%", layout="wide")

st.title("Autorretención sobre ingresos")
st.caption("Cuentas contables clase 4 · base = suma neta de la columna `valor` del auxiliar")
st.warning("Liquidación **preliminar sujeta a revisión**. La parametrización de la tarifa no sustituye la validación "
           "de las reglas tributarias aplicables a la empresa y al período.")

ss = st.session_state
STEPS = ["1. Cargar auxiliar", "2. Tarifas por empresa", "3. Liquidación"]
if "goto" in ss:                       # navegación programática (debe aplicarse antes de crear el radio)
    ss["step"] = ss.pop("goto")
ss.setdefault("step", STEPS[0])
step = st.radio("Paso", STEPS, horizontal=True, key="step", label_visibility="collapsed")

# =============================================================== 1. AUXILIAR
if step == STEPS[0]:
    st.subheader("Cargar auxiliar contable")
    files = st.file_uploader("CSV o Excel (obligatorias: empresa, fecha, codigocuenta, nombrecuenta, valor)",
                             type=["csv", "xlsx", "xlsm"], accept_multiple_files=True)
    st.caption("Volver a cargar el mismo archivo no duplica movimientos: se identifican por empresa + id del auxiliar. "
               "Al procesar, la app detecta las empresas del archivo y pasa a pedirle la tarifa de cada una.")
    if files and st.button("Procesar archivos y continuar", type="primary"):
        flash, cargados, empresas_det = [], 0, set()
        for f in files:
            try:
                rows, errs = core.normalize_rows(core.read_table(f.getvalue(), f.name))
            except Exception as e:  # noqa: BLE001
                st.error(f"**{f.name}**: no se pudo leer ({e}).")
                continue
            if not rows:
                st.error(f"**{f.name}**: {errs[0] if errs else 'sin movimientos válidos'}")
                continue
            r = db.upsert_movimientos(rows)
            cargados += 1
            empresas_det |= {m["empresa"] for m in rows}
            msg = (f"**{f.name}**: {len(rows)} filas leídas → **{r['added']}** nuevas, {r['updated']} actualizadas, "
                   f"{r['unchanged']} ya existentes (no duplicadas)."
                   + (f" {r['repeated']} repetidas dentro del archivo." if r["repeated"] else "")
                   + (f"\n\n{len(errs)} fila(s) omitidas: " + " | ".join(errs[:5]) if errs else ""))
            flash.append(("warning" if errs else "success", msg))
        if cargados:
            flash.append(("info", "Empresas detectadas en el archivo: **" + ", ".join(sorted(empresas_det)) + "**."))
            ss["flash"] = flash
            ss["goto"] = STEPS[1]
            st.rerun()
    st.subheader("Movimientos almacenados")
    inv = db.inventario()
    if inv:
        st.dataframe(pd.DataFrame(inv).rename(columns={"empresa": "Empresa", "mes": "Mes", "movimientos": "Movimientos",
                                                         "ingresos": "De ingresos (4)", "retenciones": "Retenciones (2365/2368)"}), hide_index=True, width="stretch")
        if st.button("Continuar a las tarifas por empresa →"):
            ss["goto"] = STEPS[1]; st.rerun()
        with st.expander("Borrar todos los movimientos"):
            if st.checkbox("Confirmo que quiero borrar todos los movimientos (el maestro de empresas no se borra)") and st.button("Borrar"):
                db.borrar_movimientos()
                st.rerun()
    else:
        st.info("Aún no hay movimientos cargados.")

# =============================================================== 2. TARIFAS POR EMPRESA (maestro)
if step == STEPS[1]:
    for kind, msg in ss.pop("flash", []):
        getattr(st, kind)(msg)
    st.subheader("Maestro de empresas: tarifa de autorretención")
    en_datos = sorted({m["empresa"] for m in db.movimientos()})
    if not en_datos:
        st.info("Primero cargue un auxiliar (paso 1)."); st.stop()
    st.caption("Estas son las empresas detectadas en el auxiliar. Ingrese la **tarifa de autorretención (%)** de cada una y guarde. "
               "La tarifa no está en el código y queda guardada en el maestro para las próximas cargas.")
    maestro = {e["nombre"]: e for e in db.empresas()}
    nombres = sorted(set(en_datos) | set(maestro))
    df0 = pd.DataFrame([{"nombre": n, "en_auxiliar": n in en_datos,
                         "tarifa": pd.to_numeric(maestro.get(n, {}).get("tarifa"), errors="coerce"),
                         "activa": maestro.get(n, {}).get("activa", True),
                         "observaciones": maestro.get(n, {}).get("observaciones", "")} for n in nombres])
    sin_tarifa = [r.nombre for r in df0.itertuples() if r.en_auxiliar and pd.isna(r.tarifa)]
    if sin_tarifa:
        st.warning("Falta la tarifa de: **" + ", ".join(sin_tarifa) + "**.")
    ed = st.data_editor(
        df0, hide_index=True, width="stretch", key=f"maestro_{len(nombres)}", num_rows="fixed",
        column_config={
            "nombre": st.column_config.TextColumn("Empresa", disabled=True),
            "en_auxiliar": st.column_config.CheckboxColumn("En el auxiliar", disabled=True),
            "tarifa": st.column_config.NumberColumn("Tarifa de autorretención (%)", min_value=0.0, max_value=100.0, step=0.01, format="%.4f"),
            "activa": st.column_config.CheckboxColumn("Activa"),
            "observaciones": st.column_config.TextColumn("Observaciones / referencia normativa"),
        })
    st.markdown("**NIT excluidos del cálculo**")
    nits_txt = st.text_input("Los movimientos cuyo tercero tenga alguno de estos NIT no entran a la base ni a las retenciones "
                             "(separe con comas; sin dígito de verificación).",
                             value=db.get_meta("nits_excluidos", core.NITS_EXCLUIDOS_DEFECTO))
    if st.button("Guardar y continuar a la liquidación →", type="primary"):
        lista = [dict(nombre=str(r["nombre"]), tarifa=None if pd.isna(r["tarifa"]) else r["tarifa"],
                      activa=True if pd.isna(r["activa"]) else bool(r["activa"]),
                      observaciones="" if pd.isna(r["observaciones"]) else r["observaciones"]) for _, r in ed.iterrows()]
        errs = [e for x in lista if x["nombre"] in en_datos for e in core.validar_empresa(x)]
        if errs:
            st.error("\n".join(f"- {e}" for e in errs))
        else:
            db.guardar_empresas(lista)
            db.set_meta("nits_excluidos", ", ".join(core.parse_nits(nits_txt)))
            ss["flash"] = []
            ss["goto"] = STEPS[2]
            st.rerun()

# =============================================================== 3. LIQUIDACIÓN
if step == STEPS[2]:
    todos = db.movimientos()
    empresas = sorted({m["empresa"] for m in todos})
    if not empresas:
        st.info("Cargue un auxiliar en el paso 1.")
        st.stop()
    c1, c2, c3, c4 = st.columns([2, 1.3, 1.3, 1.3])
    empresa = c1.selectbox("Empresa", empresas)
    meses = sorted({m["fecha"][:7] for m in todos if m["empresa"] == empresa}, reverse=True)
    mes = c2.selectbox("Período", meses + ["Rango personalizado"])
    if mes == "Rango personalizado":
        fechas = [m["fecha"] for m in todos if m["empresa"] == empresa]
        d_def, h_def = date.fromisoformat(min(fechas)), date.fromisoformat(max(fechas))
    else:
        y, mm = map(int, mes.split("-"))
        d_def = date(y, mm, 1)
        h_def = (date(y + (mm == 12), mm % 12 + 1, 1) - timedelta(days=1))
    desde = c3.date_input("Desde", d_def, key=f"d{empresa}{mes}")
    hasta = c4.date_input("Hasta", h_def, key=f"h{empresa}{mes}")
    invertir = st.checkbox("Presentar ingresos en positivo (invertir el signo crédito del auxiliar)", value=True)
    if hasta < desde:
        st.error("La fecha final es anterior a la inicial."); st.stop()

    nits = core.parse_nits(db.get_meta("nits_excluidos", core.NITS_EXCLUIDOS_DEFECTO))
    movs_todos = core.filtrar(todos, empresa, desde.isoformat(), hasta.isoformat())
    movs, excluidos = core.separar_excluidos(movs_todos, nits)   # solo `movs` entra al cálculo
    base = core.build_base(movs, invertir)
    tr = core.resolver_tarifa(db.empresas(), empresa)
    retenciones = core.build_retenciones(movs, invertir)
    tot_ret = core.totales_retenciones(retenciones)
    ok = tr.status == "ok"
    pct = tr.tarifa if ok else None
    auto = core.calc_autorretencion(base.base, pct) if ok else None

    if not ok:
        st.error(f"**No se calcula la autorretención.** {tr.msg}" + "\n\nVaya al paso *2. Tarifas por empresa*.")
    k1, k2, k3 = st.columns(3)
    k1.metric("Base de autorretención", fmt(base.base), f"{len(base.cuentas)} cuentas · {base.n_movs} movimientos", delta_color="off")
    k2.metric("Porcentaje aplicado", f"{pct} %" if ok else "—",
              "Del Maestro de empresas" if ok else "Sin tarifa aplicable", delta_color="off")
    k3.metric("Autorretención calculada (preliminar)", fmt(auto) if ok else "No calculada", "Base × % / 100" if ok else None, delta_color="off")

    st.subheader("Informe de autorretención")
    st.caption(f"**Empresa:** {empresa} · **Período de liquidación:** {desde} a {hasta}"
               + (f" · **Tarifa:** {pct} %" if ok else ""))
    tabla = pd.DataFrame([{"Código": c.codigo, "Cuenta": c.nombre, "Mov.": c.n, "Suma de valor (auxiliar)": float(c.raw),
                           "Valor neto": float(c.valor), "Autorretención": float(core.calc_autorretencion(c.valor, pct)) if ok else None}
                          for c in base.cuentas])
    if len(tabla):
        tot = {"Código": "TOTAL", "Cuenta": "Base total de autorretención", "Mov.": base.n_movs, "Suma de valor (auxiliar)": float(base.base_raw),
               "Valor neto": float(base.base), "Autorretención": float(auto) if ok else None}
        tabla = pd.concat([tabla, pd.DataFrame([tot])], ignore_index=True)
        neg_codes = {c.codigo for c in base.negativas}
        st.dataframe(tabla.style.format({k: "{:,.2f}" for k in ["Suma de valor (auxiliar)", "Valor neto", "Autorretención"]}, na_rep="—")
                     .apply(lambda r: ["background-color: #fff3d6; color:#5a3d00" if r["Código"] in neg_codes else "" for _ in r], axis=1),
                     hide_index=True, width="stretch")
    else:
        st.info("No hay cuentas de ingresos en el período.")

    st.markdown("**Rastreo a movimientos originales**")
    for c in base.cuentas:
        with st.expander(f"{c.codigo} · {c.nombre} — {c.n} movimientos · valor neto {fmt(c.valor)}"):
            q = st.text_input("Buscar", key=f"q{c.codigo}", placeholder="tercero, descripción, documento…")
            dfm = pd.DataFrame([{"Fila": m["fila"], "Id": m["id"], "Fecha": m["fecha"], "Doc.": f"{m['tipodocumento']} {m['nrodocumento']}".strip(),
                                 "Tercero": m["tercero"], "Identificación": m["identificacion"], "Descripción": m["descripcion"],
                                 "C. costo": m["centrocosto"], "valor": float(m["valor"])} for m in c.movs])
            if q:
                dfm = dfm[dfm.apply(lambda r: q.lower() in " ".join(map(str, r.values)).lower(), axis=1)]
            st.dataframe(dfm, hide_index=True, width="stretch")

    # ------------------------------------------------ retenciones 2365 / 2368
    st.subheader("Retenciones practicadas (cuentas 2365 y 2368)")
    st.caption("Retención causada = movimientos con `valor` negativo (crédito); Débitos = cierres, pagos y devoluciones. "
               "Base y tarifa vienen de `baseimpuesto` y `porcentajeimpuesto` del auxiliar. Las cuentas de *autorretención* y de *pago* "
               "se muestran pero no suman en el total de retenciones practicadas.")
    if retenciones:
        TIPO = {"practicada": "Retención practicada", "autorretencion": "Autorretención contabilizada", "pago": "Pago / cierre"}
        filas = []
        for g, nombre in core.GRUPOS_RETENCION.items():
            cs = [c for c in retenciones if c.grupo == g]
            if not cs:
                continue
            for c in cs:
                filas.append({"Grupo": g, "Código": c.codigo, "Cuenta": c.nombre, "Tipo": TIPO[c.tipo],
                              "Tarifa %": ", ".join(f"{t:f}" for t in sorted(c.tarifas)), "Base": float(c.base),
                              "Retención causada": float(c.creditos), "Débitos (−)": -float(c.debitos), "Neto": float(c.creditos - c.debitos), "Mov.": c.n})
            t = tot_ret[g]
            filas.append({"Grupo": g, "Código": "SUBTOTAL", "Cuenta": f"{nombre} (practicadas)", "Tipo": "", "Tarifa %": "",
                          "Base": float(t["base"]), "Retención causada": float(t["creditos"]), "Débitos (−)": -float(t["debitos"]),
                          "Neto": float(t["creditos"] - t["debitos"]), "Mov.": None})
        t = tot_ret["total"]
        filas.append({"Grupo": "", "Código": "TOTAL", "Cuenta": "Total retenciones practicadas (neto)", "Tipo": "", "Tarifa %": "",
                      "Base": float(t["base"]), "Retención causada": float(t["creditos"]), "Débitos (−)": -float(t["debitos"]),
                      "Neto": float(t["creditos"] - t["debitos"]), "Mov.": None})
        dfr = pd.DataFrame(filas)
        st.dataframe(dfr.style.format({k: "{:,.2f}" for k in ["Base", "Retención causada", "Débitos (−)", "Neto"]}, na_rep="")
                     .apply(lambda r: ["font-weight:700; background-color:#eef2f7; color:#1c2430" if str(r["Código"]) in ("SUBTOTAL", "TOTAL") else "" for _ in r], axis=1),
                     hide_index=True, width="stretch")
        pers = [{"Código": c.codigo, "Cuenta": c.nombre, "Tipo de persona": per, "Base": float(b), "Retención causada": float(r),
                 "Reversiones (−)": -float(d), "Neto": float(r - d)}
                for c in retenciones if c.tipo == "practicada" for per, (b, r, d) in sorted(c.por_persona.items())]
        if pers:
            with st.expander("Retenciones practicadas por tipo de persona (jurídica / natural)"):
                st.dataframe(pd.DataFrame(pers).style.format({k: "{:,.2f}" for k in ["Base", "Retención causada", "Reversiones (−)", "Neto"]}), hide_index=True, width="stretch")
        for c in retenciones:
            with st.expander(f"{c.codigo} · {c.nombre} — {c.n} movimientos"):
                st.dataframe(pd.DataFrame([{"Fila": m["fila"], "Id": m["id"], "Fecha": m["fecha"],
                                            "Doc.": f"{m['tipodocumento']} {m['nrodocumento']}".strip(), "Tercero": m["tercero"],
                                            "Persona": m.get("tipo_persona", ""), "Descripción": m["descripcion"],
                                            "Base": float(m.get("baseimpuesto", 0)), "%": float(m.get("porcentajeimpuesto", 0)),
                                            "valor": float(m["valor"])} for m in c.movs]), hide_index=True, width="stretch")
    else:
        st.info("No hay movimientos en cuentas 2365 ni 2368 para esta empresa y período.")

    if excluidos:
        with st.expander(f"Movimientos excluidos del cálculo por NIT ({len(excluidos)}) — {', '.join(nits)}"):
            st.caption("Estos movimientos no entran a la base de autorretención ni a las retenciones. Se listan para trazabilidad.")
            st.dataframe(pd.DataFrame([{"Fila": m["fila"], "Id": m["id"], "Fecha": m["fecha"], "Doc.": f"{m['tipodocumento']} {m['nrodocumento']}".strip(),
                                        "Cuenta": m["codigocuenta"], "Nombre cuenta": m["nombrecuenta"], "NIT": m["identificacion"],
                                        "Descripción": m["descripcion"], "valor": float(m["valor"])} for m in excluidos]),
                         hide_index=True, width="stretch")

    st.subheader("Cuentas con saldo negativo")
    if base.negativas:
        st.caption("Simulación informativa — no modifica la liquidación.")
        st.dataframe(pd.DataFrame([{"Código": c.codigo, "Cuenta": c.nombre, "Valor neto": float(c.valor)} for c in base.negativas]), hide_index=True)
        txt = (f"**Base actual:** {fmt(base.base)} · **Efecto de las negativas:** {fmt(base.efecto_negativas)} · "
               f"**Base sin ellas:** {fmt(base.base_sin_negativas)}")
        if ok:
            txt += f"  \n**Autorretención actual:** {fmt(auto)} · **sin cuentas negativas:** {fmt(core.calc_autorretencion(base.base_sin_negativas, pct))}"
        st.markdown(txt)
    else:
        st.caption("Ninguna cuenta de ingresos tiene saldo negativo en el período.")

    st.subheader("Controles y validaciones")
    checks = core.ejecutar_controles(movs, base, tr, len(todos), len({m["key"] for m in todos}), retenciones, auto, excluidos, nits)
    for c in checks:
        icon = {"ok": "✅", "warn": "⚠️", "error": "❌"}[c["nivel"]]
        st.markdown(f"{icon} **{c['titulo']}** — {c['detalle']}")

    st.download_button("Exportar a Excel", exportar(empresa, desde.isoformat(), hasta.isoformat(), base, tr, checks, retenciones, excluidos, nits),
                       file_name=f"autorretencion_{empresa}_{desde}_{hasta}.xlsx", type="primary",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
