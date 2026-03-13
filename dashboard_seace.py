import streamlit as st
import pandas as pd
import os
from datetime import datetime

st.set_page_config(
    page_title="SEACE Monitor",
    page_icon="📋",
    layout="wide"
)

st.title("📋 SEACE Monitor — Procesos de Bienes")
st.caption("Regiones: Cusco, Puno, Apurímac, Madre de Dios, Arequipa, Ayacucho, Moquegua")

# Buscar el archivo Excel más reciente en el Escritorio
escritorio = os.path.expanduser("~/Desktop")
archivos = [f for f in os.listdir(escritorio) if f.startswith("SEACE_Bienes_") and f.endswith(".xlsx")]

if not archivos:
    st.error("No se encontró ningún archivo SEACE_Bienes en el Escritorio. Ejecuta primero seace_monitor.py")
    st.stop()

archivo_reciente = sorted(archivos)[-1]
ruta = os.path.join(escritorio, archivo_reciente)
st.success(f"Archivo cargado: {archivo_reciente}")

# Cargar datos
df = pd.read_excel(ruta)

# Métricas resumen
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Total procesos", f"{len(df):,}")
with col2:
    valor_total = df["Monto referencial"].sum()
    st.metric("Valor referencial total", f"S/ {valor_total:,.0f}")
with col3:
    regiones = df["Departamento"].nunique()
    st.metric("Regiones", regiones)
with col4:
    entidades = df["Entidad"].nunique()
    st.metric("Entidades convocantes", entidades)

st.divider()

# Filtros
st.subheader("Filtros")
col_f1, col_f2, col_f3 = st.columns(3)

with col_f1:
    regiones_lista = ["Todas"] + sorted(df["Departamento"].dropna().unique().tolist())
    region_sel = st.selectbox("Región", regiones_lista)

with col_f2:
    tipos_lista = ["Todos"] + sorted(df["Tipo de proceso"].dropna().unique().tolist())
    tipo_sel = st.selectbox("Tipo de proceso", tipos_lista)

with col_f3:
    busqueda = st.text_input("Buscar en descripción", placeholder="Ej: madera, uniforme, medicamento...")

# Aplicar filtros
df_filtrado = df.copy()
if region_sel != "Todas":
    df_filtrado = df_filtrado[df_filtrado["Departamento"] == region_sel]
if tipo_sel != "Todos":
    df_filtrado = df_filtrado[df_filtrado["Tipo de proceso"] == tipo_sel]
if busqueda:
    df_filtrado = df_filtrado[
        df_filtrado["Descripción / Objeto"].str.contains(busqueda, case=False, na=False)
    ]

st.caption(f"Mostrando {len(df_filtrado):,} de {len(df):,} procesos")

st.divider()

# Gráficos
col_g1, col_g2 = st.columns(2)

with col_g1:
    st.subheader("Procesos por región")
    chart_region = df_filtrado["Departamento"].value_counts()
    st.bar_chart(chart_region)

with col_g2:
    st.subheader("Procesos por tipo")
    chart_tipo = df_filtrado["Tipo de proceso"].value_counts()
    st.bar_chart(chart_tipo)

st.divider()

# Tabla de datos
st.subheader("Detalle de procesos")
st.dataframe(
    df_filtrado[[
        "N° Convocatoria", "Descripción / Objeto", "Entidad",
        "Tipo de proceso", "Monto referencial", "Departamento",
        "Provincia", "Fecha convocatoria", "Estado"
    ]],
    use_container_width=True,
    height=400
)

# Botón descargar
st.download_button(
    label="Descargar Excel filtrado",
    data=open(ruta, "rb").read(),
    file_name=f"SEACE_Filtrado_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)