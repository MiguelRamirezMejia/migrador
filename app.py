# -*- coding: utf-8 -*-
import io
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from core import esquema as esquema_mod
from core.excel_loader import ExcelLoadError
from core.pipeline import ejecutar_validacion, resumen
from core.sql_generator import generar_sql_completo

st.set_page_config(page_title="Migrador - Validador", page_icon="🔎", layout="wide")

st.markdown(
    """
    <style>
    .block-container {padding-top: 2rem;}
    .metric-card {
        border-radius: 10px; padding: 1rem 1.2rem; text-align: center;
        border: 1px solid rgba(255,255,255,0.08);
    }
    .metric-card h2 {margin: 0; font-size: 2rem;}
    .metric-card p {margin: 0; opacity: 0.75; font-size: 0.85rem; text-transform: uppercase; letter-spacing: .05em;}
    .ok {background: rgba(46, 160, 67, 0.15); border-color: rgba(46,160,67,0.4);}
    .warn {background: rgba(227, 179, 65, 0.15); border-color: rgba(227,179,65,0.4);}
    .err {background: rgba(220, 53, 69, 0.15); border-color: rgba(220,53,69,0.4);}
    .total {background: rgba(90, 130, 230, 0.12); border-color: rgba(90,130,230,0.35);}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("🔎 Migrador — Validador de plantillas de migración")
st.caption("Sube la plantilla Excel, valida su integridad contra las reglas del sistema destino y descarga el reporte y el SQL.")

try:
    ESQUEMA = esquema_mod.cargar()
except esquema_mod.EsquemaError as e:
    st.error(f"No se pudo cargar el esquema de migración: {e}")
    st.stop()
COD = ESQUEMA.etiqueta("codigo_servicio", "código de servicio")

if "resultado" not in st.session_state:
    st.session_state.resultado = None

with st.sidebar:
    st.header("Archivo")
    st.caption(f"Esquema: **{ESQUEMA.nombre}** (`{ESQUEMA.origen.name}`)")
    st.caption("Funciona 100% sin conexión a base de datos — el catálogo de ciudades va empaquetado localmente.")
    archivo = st.file_uploader("Plantilla Excel", type=["xlsx"])

    st.divider()
    usar_bd = st.toggle("Reconciliar contra una base ya poblada", value=False,
                          help="Opcional. Si el destino ya tiene datos (no es una migración desde cero), activa esto "
                               "para que PRODUCTOS y PROYECTOS se comparen con matching tolerante contra el catálogo "
                               "real, en vez de solo contra la propia plantilla.")
    conn_params = None
    if usar_bd:
        host = st.text_input("Host", value="127.0.0.1")
        port = st.text_input("Puerto", value="5432")
        dbname = st.text_input("Base de datos")
        user = st.text_input("Usuario", value="postgres")
        password = st.text_input("Password", type="password")
        conn_params = (host, port, user, password, dbname)

    validar_btn = st.button("Validar plantilla", type="primary", use_container_width=True, disabled=archivo is None)

if validar_btn and archivo is not None:
    conn = None
    if usar_bd:
        from core.db import connect, DbConnectionError
        host, port, user, password, dbname = conn_params
        if not dbname:
            st.error("Activaste la reconciliación contra BD pero no indicaste el nombre de la base de datos.")
            st.stop()
        with st.spinner("Conectando a la base de datos..."):
            try:
                conn = connect(host, port, user, password, dbname)
            except DbConnectionError as e:
                st.error(f"No se pudo conectar a la base de datos: {e}")
                st.stop()

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp) / archivo.name
        tmp_path.write_bytes(archivo.getvalue())
        with st.spinner("Validando plantilla..."):
            try:
                findings, contexto = ejecutar_validacion(str(tmp_path), conn=conn, esquema=ESQUEMA)
            except ExcelLoadError as e:
                st.error(f"Error al leer el Excel: {e}")
                st.stop()
            except KeyError as e:
                st.error(f"Falta una hoja o columna esperada en el Excel: {e}")
                st.stop()
            except esquema_mod.EsquemaError as e:
                st.error(f"El esquema de migración está incompleto: {e}")
                st.stop()

    st.session_state.resultado = (findings, contexto)

resultado = st.session_state.resultado

if resultado is None:
    st.info("Sube una plantilla y presiona **Validar plantilla** en la barra lateral para comenzar.")
    st.stop()

findings, contexto = resultado
res = resumen(findings)

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown(f'<div class="metric-card total"><h2>{res["total"]}</h2><p>Total filas evaluadas</p></div>', unsafe_allow_html=True)
with c2:
    st.markdown(f'<div class="metric-card ok"><h2>{res["OK"]}</h2><p>OK</p></div>', unsafe_allow_html=True)
with c3:
    st.markdown(f'<div class="metric-card warn"><h2>{res["ADVERTENCIA"]}</h2><p>Advertencias</p></div>', unsafe_allow_html=True)
with c4:
    st.markdown(f'<div class="metric-card err"><h2>{res["ERROR"]}</h2><p>Errores bloqueantes</p></div>', unsafe_allow_html=True)

st.caption(f"Convención de join detectada en {ESQUEMA.hoja('productos_contratos')}: **{contexto['convencion_join_productos_contratos']}**")
if contexto.get("reconciliado_con_bd"):
    n_prod = sum(1 for p in contexto["productos_index"].values() if p.get("id_real"))
    n_proy = sum(1 for p in contexto["proyectos_index"].values() if p.get("id_real"))
    st.caption(f"🔗 Reconciliado contra BD real: {n_prod} producto(s) y {n_proy} proyecto(s) se reutilizarán (no se crearán duplicados).")

st.divider()

df = pd.DataFrame([{
    "Hoja": f.hoja, "Fila Excel": f.fila, "Entidad": f.entidad,
    "Severidad": f.severidad, "Referencia": f.referencia, "Mensaje": f.mensaje,
} for f in findings])

colf1, colf2, colf3 = st.columns(3)
with colf1:
    hojas = ["(todas)"] + sorted(df["Hoja"].unique().tolist())
    hoja_sel = st.selectbox("Filtrar por hoja", hojas)
with colf2:
    severidades = ["(todas)", "ERROR", "ADVERTENCIA", "OK"]
    sev_sel = st.selectbox("Filtrar por severidad", severidades, index=1)
with colf3:
    busqueda = st.text_input(f"Buscar ({COD}, documento, nombre...)")

df_filtrado = df.copy()
if hoja_sel != "(todas)":
    df_filtrado = df_filtrado[df_filtrado["Hoja"] == hoja_sel]
if sev_sel != "(todas)":
    df_filtrado = df_filtrado[df_filtrado["Severidad"] == sev_sel]
if busqueda:
    mask = df_filtrado["Referencia"].astype(str).str.contains(busqueda, case=False, na=False) | \
           df_filtrado["Mensaje"].astype(str).str.contains(busqueda, case=False, na=False)
    df_filtrado = df_filtrado[mask]


def _color_severidad(val):
    colores = {"ERROR": "background-color: rgba(220,53,69,0.25)",
               "ADVERTENCIA": "background-color: rgba(227,179,65,0.25)",
               "OK": "background-color: rgba(46,160,67,0.15)"}
    return colores.get(val, "")


st.dataframe(
    df_filtrado.style.map(_color_severidad, subset=["Severidad"]),
    use_container_width=True,
    height=480,
)

st.divider()

col_dl1, col_dl2 = st.columns(2)
with col_dl1:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Reporte")
    st.download_button(
        "⬇️ Descargar reporte completo (Excel)",
        data=buffer.getvalue(),
        file_name="reporte_validacion_migracion.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with col_dl2:
    if res["ERROR"] > 0:
        st.button("⬇️ Generar SQL de migración", disabled=True, use_container_width=True,
                   help="Deshabilitado: hay errores bloqueantes que debes corregir primero.")
    else:
        sql_text = generar_sql_completo(contexto)
        st.download_button(
            "⬇️ Descargar SQL de migración",
            data=sql_text,
            file_name="migracion.sql",
            mime="text/plain",
            use_container_width=True,
        )
