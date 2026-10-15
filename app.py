"""Autorretención sobre ingresos — app Streamlit.  Ejecutar:  streamlit run app.py"""
from datetime import date, timedelta
from decimal import Decimal

import pandas as pd
import streamlit as st

import core
import db
from core import fmt
from excel import exportar

st.set_page_config(page_title="Autorretención sobre ingresos", page_icon="%", layout="wide")
ss = st.session_state
ss.setdefault("auth", False)

st.title("Autorretención sobre ingresos")
st.caption("Cuentas contables clase 4 · base = suma neta de la columna `valor` del auxiliar")
st.warning("Liquidación **preliminar sujeta a revisión**. La parametrización de la tarifa no sustituye la validación "
           "de las reglas tributarias aplicables a la empresa y al período.")

tab_datos, tab_cfg, tab_liq = st.tabs(["1. Auxiliar", "2. Configuración", "3. Liquidación"])

# =============================================================== 1. AUXILIAR
with tab_datos:
    st.subheader("Cargar auxiliar contable")
    files = st.file_uploader("CSV o Excel (obligatorias: empresa, fecha, codigocuenta, nombrecuenta, valor)",
                             type=["csv", "xlsx", "xlsm"], accept_multiple_files=True)
    st.caption("Volver a cargar el mismo archivo no duplica movimientos: se identifican por empresa + id del auxiliar.")
    if files and st.button("Procesar archivos", type="primary"):
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
            msg = (f"**{f.name}**: {len(rows)} filas leídas → **{r['added']}** nuevas, {r['updated']} actualizadas, "
                   f"{r['unchanged']} ya existentes (no duplicadas)."
                   + (f" {r['repeated']} repetidas dentro del archivo." if r["repeated"] else ""))
            (st.warning if errs else st.success)(msg + (f"\n\n{len(errs)} fila(s) omitidas: " + " | ".join(errs[:5]) if errs else ""))
    st.subheader("Movimientos almacenados")
    inv = db.inventario()
    if inv:
        st.dataframe(pd.DataFrame(inv).rename(columns={"empresa": "Empresa", "mes": "Mes", "movimientos": "Movimientos",
                                                         "ingresos": "De ingresos (4)"}), hide_index=True, width="stretch")
        with st.expander("Borrar todos los movimientos"):
            if st.checkbox("Confirmo que quiero borrar todos los movimientos (las tarifas no se borran)") and st.button("Borrar"):
                db.borrar_movimientos()
                st.rerun()
    else:
        st.info("Aún no hay movimientos cargados.")

# =============================================================== 2. CONFIGURACIÓN
with tab_cfg:
    st.subheader("Tarifas de autorretención")
    st.caption("El porcentaje no está en el código: lo define el usuario autorizado por empresa y vigencia.")
    if not ss.auth:
        c1, c2 = st.columns([1, 2])
        with c1.form("login"):
            nuevo = not db.pin_configurado()
            pin = st.text_input("Cree una clave de edición (mín. 4 caracteres)" if nuevo else "Clave de edición", type="password")
            if st.form_submit_button("Crear y entrar" if nuevo else "Desbloquear edición"):
                if len(pin) < 4:
                    st.error("Mínimo 4 caracteres.")
                elif nuevo:
                    db.crear_pin(pin); ss.auth = True; st.rerun()
                elif db.verificar_pin(pin):
                    ss.auth = True; st.rerun()
                else:
                    st.error("Clave incorrecta.")
        c2.info("Modo solo lectura. Puede consultar las tarifas; para crearlas o editarlas desbloquee con la clave.")
    else:
        if st.button("Bloquear edición"):
            ss.auth = False; st.rerun()

    tarifas = db.tarifas()
    if ss.auth:
        empresas = sorted({m["empresa"] for m in db.movimientos()} | {t["empresa"] for t in tarifas})
        opciones = {"➕ Nueva tarifa": None, **{f"{t['empresa']} · {t['porcentaje']}% · {t['desde']} → {t['hasta'] or '…'}": t for t in tarifas}}
        sel = opciones[st.selectbox("Crear o editar", list(opciones))]
        t0 = sel or {}
        with st.form("tarifa", clear_on_submit=False):
            c1, c2, c3 = st.columns(3)
            emp_sel = c1.selectbox("Empresa", empresas + ["(otra…)"], index=(empresas.index(t0["empresa"]) if t0 and t0["empresa"] in empresas else 0))
            emp_otra = c1.text_input("Si es otra, escríbala") if emp_sel == "(otra…)" else ""
            pct = c2.number_input("Porcentaje (%)", min_value=0.0, max_value=100.0, step=0.01, format="%.4f",
                                  value=float(t0["porcentaje"]) if t0 else 0.0)
            activa = c3.selectbox("Estado", ["Activa", "Inactiva"], index=0 if not t0 or t0["activa"] else 1)
            d1, d2 = st.columns(2)
            desde = d1.date_input("Vigente desde", value=date.fromisoformat(t0["desde"]) if t0 else date.today().replace(month=1, day=1))
            con_fin = d2.checkbox("Tiene fecha final", value=bool(t0 and t0["hasta"]))
            hasta = d2.date_input("Vigente hasta", value=date.fromisoformat(t0["hasta"]) if t0 and t0["hasta"] else date.today().replace(month=12, day=31))
            ref = st.text_input("Referencia normativa (opcional)", value=(t0 or {}).get("referencia") or "")
            obs = st.text_input("Observaciones (opcional)", value=(t0 or {}).get("observaciones") or "")
            forzar = st.checkbox("Confirmo guardar aunque la vigencia se superponga con otra tarifa activa")
            ok = st.form_submit_button("Guardar tarifa", type="primary")
        if ok:
            t = dict(id=(t0 or {}).get("id"), empresa=(emp_otra if emp_sel == "(otra…)" else emp_sel).strip(), porcentaje=pct,
                     desde=desde.isoformat(), hasta=hasta.isoformat() if con_fin else None, activa=activa == "Activa",
                     referencia=ref.strip(), observaciones=obs.strip())
            errs = core.validar_tarifa(t)
            sol = core.buscar_solapes(tarifas, t) if not errs else []
            if errs:
                st.error(" ".join(errs))
            elif sol and not forzar:
                st.warning("**Advertencia: vigencia superpuesta** con otra tarifa activa de esta empresa: "
                           + "; ".join(f"{o['porcentaje']}% ({o['desde']} a {o['hasta'] or 'sin fecha final'})" for o in sol)
                           + ". Marque la casilla de confirmación para guardar. Mientras se superpongan, esos períodos no se calcularán.")
            else:
                db.guardar_tarifa(t); st.success("Tarifa guardada."); st.rerun()
        if sel and st.button("Eliminar tarifa seleccionada"):
            db.eliminar_tarifa(sel["id"]); st.rerun()
    if tarifas:
        df = pd.DataFrame([{**t, "superpuesta": "Sí" if core.buscar_solapes(tarifas, t) else "", "activa": "Activa" if t["activa"] else "Inactiva"} for t in tarifas])
        st.dataframe(df.drop(columns=["id"]).rename(columns=str.capitalize), hide_index=True, width="stretch")
    else:
        st.info("No hay tarifas configuradas. Sin tarifa, la autorretención no se calcula.")

# =============================================================== 3. LIQUIDACIÓN
with tab_liq:
    todos = db.movimientos()
    empresas = sorted({m["empresa"] for m in todos})
    if not empresas:
        st.info("Cargue un auxiliar en la pestaña 1.")
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

    movs = core.filtrar(todos, empresa, desde.isoformat(), hasta.isoformat())
    base = core.build_base(movs, invertir)
    tr = core.resolver_tarifa(db.tarifas(), empresa, desde.isoformat(), hasta.isoformat())
    ok = tr.status == "ok"
    pct = Decimal(tr.tarifa["porcentaje"]) if ok else None
    auto = core.calc_autorretencion(base.base, pct) if ok else None

    if not ok:
        st.error(f"**No se calcula la autorretención.** {tr.msg}" + "".join(
            f"\n- {t['porcentaje']}% ({t['desde']} a {t['hasta'] or 'sin fecha final'})" for t in tr.candidatas)
            + "\n\nVaya a *2. Configuración*.")
    k1, k2, k3 = st.columns(3)
    k1.metric("Base de autorretención", fmt(base.base), f"{len(base.cuentas)} cuentas · {base.n_movs} movimientos", delta_color="off")
    k2.metric("Porcentaje aplicado", f"{pct} %" if ok else "—",
              f"Vigencia {tr.tarifa['desde']} a {tr.tarifa['hasta'] or 'sin fecha final'}" if ok else "Sin tarifa aplicable", delta_color="off")
    k3.metric("Autorretención calculada (preliminar)", fmt(auto) if ok else "No calculada", "Base × % / 100" if ok else None, delta_color="off")

    st.subheader("Informe de autorretención")
    st.caption(f"**Empresa:** {empresa} · **Período de liquidación:** {desde} a {hasta}"
               + (f" · **Tarifa:** {pct} % (vigencia {tr.tarifa['desde']} a {tr.tarifa['hasta'] or 'sin fecha final'})" if ok else ""))
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
    checks = core.ejecutar_controles(movs, base, tr, len(todos), len({m["key"] for m in todos}))
    for c in checks:
        icon = {"ok": "✅", "warn": "⚠️", "error": "❌"}[c["nivel"]]
        st.markdown(f"{icon} **{c['titulo']}** — {c['detalle']}")

    st.download_button("Exportar a Excel", exportar(empresa, desde.isoformat(), hasta.isoformat(), base, tr, checks),
                       file_name=f"autorretencion_{empresa}_{desde}_{hasta}.xlsx", type="primary",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
