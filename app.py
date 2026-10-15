"""Impuestos — app Streamlit (retención y autorretención + IVA).  Ejecutar:  streamlit run app.py"""
import streamlit as st

import pagina_iva
import pagina_ret

st.set_page_config(page_title="Retenciones, autorretención e IVA", page_icon="%", layout="wide")

st.title("Retenciones, autorretención e IVA")
st.warning("Liquidaciones **preliminares sujetas a revisión**. La parametrización no sustituye la validación "
           "de las reglas tributarias aplicables a la empresa y al período.")

tab_ret, tab_iva = st.tabs(["Retención y autorretenciones", "Impuesto sobre las ventas IVA"])
with tab_ret:
    st.caption("Autorretención sobre ingresos (cuentas 4) y retenciones practicadas (2365 y 2368) · columna `valor` del auxiliar")
    pagina_ret.render()
with tab_iva:
    st.caption("IVA: cuentas 2408 y total de la clase 4 (neto por cuenta) · requiere su propio auxiliar")
    pagina_iva.render()
