import streamlit as st
import pandas as pd
import requests
import openpyxl
from io import BytesIO
from datetime import datetime

st.set_page_config(
    page_title="SEACE Monitor",
    page_icon="📋",
    layout="wide"
)

st.title("📋 SEACE Monitor — Procesos de Bienes")
st.caption("Regiones: Cusco, Puno, Apurímac, Madre de Dios, Arequipa, Ayacucho, Moquegua")

REGIONES = ["CUSCO", "PUNO", "APURIMAC", "MADRE DE DIOS", "AREQUIPA", "AYACUCHO", "MOQUEGUA"]

PALABRAS_CLAVE = [
    "MADERA", "TABLERO", "TRIPLAY", "PARQUET", "MACHIMBRE",
    "PROTECCION PERSONAL", "EPP", "CASCO", "GUANTES", "LENTES DE SEGURIDAD",
    "BOTAS DE SEGURIDAD", "CHALECO", "ARNES", "RESPIRADOR", "MASCARILLA",
    "UNIFORME", "INDUMENTARIA", "ROPA DE TRABAJO", "CALZADO", "POLO",
    "BUZO", "OVEROL", "TERNO", "VESTIMENTA",
    "MEDICAMENTO", "FARMACO", "MEDICINA", "ANTIBIOTICO", "ANALGESICO",
    "INSUMO MEDICO", "MATERIAL MEDICO", "REACTIVO", "VACUNA",
    "EQUIPO DE PROTECCION", "SEGURIDAD INDUSTRIAL", "BOTIQUIN"
]

@st.cache_data(ttl=3600)
def cargar_datos():
    URL = "https://tinyurl.com/conosceconvocatorias2025"
    with st.spinner("Descargando datos del OSCE... (1-2 minutos)"):
        respuesta = requests.get(URL, timeout=120, allow_redirects=True)
        archivo = openpyxl.load_workbook(BytesIO(respuesta.content), read_only=True)
        hoja = archivo.active
        filas = list(hoja.rows)
        encabezados = [str(c.value).strip() if c.value else "" for c in filas[0]]
        procesos = []
        for fila in filas[1:]:
            valores = [c.value for c in fila]
            fila_dict = dict(zip(encabezados, valores))
            region = str(fila_dict.get("departamento_item", "")).upper()
            objeto = str(fila_dict.get("objetocontractual", "")).upper()
            descripcion = str(fila_dict.get("descripcion_item", "")).upper()
            texto = objeto + " " + descripcion
            if any(r in region for r in REGIONES):
                if "BIEN" in objeto and any(kw in texto for kw in PALABRAS_CLAVE):
                    procesos.append({
                        "N° Convocatoria": fila_dict.get("nroconvocatoria", ""),
                        "Descripción / Objeto": fila_dict.get("descripcion_item", ""),
                        "Entidad": fila_dict.get("entidad", ""),
                        "Tipo de proceso": fila_dict.get("tipoprocesoseleccion", ""),
                        "Monto referencial": fila_dict.get("montoreferencial", 0),
                        "Moneda": fila_dict.get("moneda", ""),
                        "Departamento": fila_dict.get("departamento_item", ""),
                        "Provincia": fila_dict.get("provincia_item", ""),
                        "Distrito": fila_dict.get("distrito_item", ""),
                        "Fecha convocatoria": fila_dict.get("fecha_convocatoria", ""),
                        "Estado": fila_dict.get("estadoitem", ""),
                    })
        return pd.DataFrame(procesos)

df = cargar_datos()

# Métricas
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Total procesos", f"{len(df):,}")
with col2:
    valor_total = pd.to_numeric(df["Monto referencial"], errors="coerce").sum()
    st.metric("Valor referencial total", f"S/ {valor_total:,.0f}")
with col3:
    st.metric("Regiones", df["Departamento"].nunique())
with col4:
    st.metric("Entidades convocantes", df["Entidad"].nunique())

st.divider()

# Filtros
st.subheader("Filtros")
col_f1, col_f2, col_f3 = st.columns(3)
with col_f1:
    region_sel = st.selectbox("Región", ["Todas"] + sorted(df["Departamento"].dropna().unique().tolist()))
with col_f2:
    tipo_sel = st.selectbox("Tipo de proceso", ["Todos"] + sorted(df["Tipo de proceso"].dropna().unique().tolist()))
with col_f3:
    busqueda = st.text_input("Buscar en descripción", placeholder="Ej: madera, uniforme, medicamento...")

df_filtrado = df.copy()
if region_sel != "Todas":
    df_filtrado = df_filtrado[df_filtrado["Departamento"] == region_sel]
if tipo_sel != "Todos":
    df_filtrado = df_filtrado[df_filtrado["Tipo de proceso"] == tipo_sel]
if busqueda:
    df_filtrado = df_filtrado[df_filtrado["Descripción / Objeto"].str.contains(busqueda, case=False, na=False)]

st.caption(f"Mostrando {len(df_filtrado):,} de {len(df):,} procesos")
st.divider()

# Gráficos
col_g1, col_g2 = st.columns(2)
with col_g1:
    st.subheader("Procesos por región")
    st.bar_chart(df_filtrado["Departamento"].value_counts())
with col_g2:
    st.subheader("Procesos por tipo")
    st.bar_chart(df_filtrado["Tipo de proceso"].value_counts())

st.divider()

# Tabla
st.subheader("Detalle de procesos")
st.dataframe(df_filtrado, use_container_width=True, height=400)

# Descargar
csv = df_filtrado.to_csv(index=False).encode("utf-8")
st.download_button(
    label="Descargar CSV filtrado",
    data=csv,
    file_name=f"SEACE_Filtrado_{datetime.now().strftime('%Y-%m-%d')}.csv",
    mime="text/csv"
)