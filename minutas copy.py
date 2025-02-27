import pandas as pd
import streamlit as st
import os

# Ruta del archivo CSV
file_path = r"C:\Users\cokeh\Desktop\car index minutas\CAR INDEX.csv"
file_path_con_comentarios = r"C:\Users\cokeh\Desktop\car index minutas\CAR_INDEX_con_comentarios.xlsx"  # Nuevo archivo para guardar comentarios

# Verificar si el archivo existe
if not os.path.exists(file_path):
    st.error("🚨 El archivo no se encontró. Verifica la ruta y el nombre del archivo.")
    st.stop()

# Cargar el archivo CSV con manejo de errores
try:
    df = pd.read_csv(file_path, encoding="utf-8", sep=";", engine="python", on_bad_lines="skip", dtype=str)
    df.columns = df.columns.str.strip()
except UnicodeDecodeError:
    df = pd.read_csv(file_path, encoding="latin1", sep=";", engine="python", on_bad_lines="skip", dtype=str)
    df.columns = df.columns.str.strip()
except Exception as e:
    st.error(f"Error al leer el archivo CSV: {e}")
    st.stop()

# =====================  DISEÑO DEL DASHBOARD  ===================== #
st.set_page_config(page_title="Dashboard CAR INDEX", layout="wide")

st.markdown('<p class="big-title">📊 Dashboard de Indicadores - CAR INDEX</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Monitorea los KPI de los vendedores con análisis visual y comentarios.</p>', unsafe_allow_html=True)
st.divider()

# ===================== Filtros ===================== #
# Verificar que las columnas necesarias existan
def validar_columna(df, columna):
    if columna not in df.columns:
        st.error(f"🚨 La columna '{columna}' no existe en los datos.")
        st.stop()

columnas_necesarias = ["YM", "Week Calendar", "Sucursal", "Jefe de venta s", "Vendedor", 
                       "actividad primer intento", "Conversión de leads", "Cobertura del tubo", 
                       "AvancePonderadoTotal"]
for col in columnas_necesarias:
    validar_columna(df, col)

# Filtros
ym_seleccionado = st.sidebar.selectbox("📅 Selecciona YM", [None] + list(df["YM"].dropna().unique()))
df_filtrado = df if ym_seleccionado is None else df[df["YM"] == ym_seleccionado]
week_calendar_seleccionado = st.sidebar.selectbox("📆 Selecciona Week Calendar", [None] + list(df_filtrado["Week Calendar"].dropna().unique()))
sucursal_seleccionada = st.sidebar.selectbox("🏢 Selecciona una Sucursal", [None] + list(df["Sucursal"].dropna().unique()))

if sucursal_seleccionada:
    df_filtrado = df_filtrado[df_filtrado["Sucursal"] == sucursal_seleccionada]
if week_calendar_seleccionado:
    df_filtrado = df_filtrado[df_filtrado["Week Calendar"] == week_calendar_seleccionado]

jefe_seleccionado = st.sidebar.selectbox("👔 Selecciona un Jefe de Venta", [None] + list(df_filtrado["Jefe de venta s"].dropna().unique()))
if jefe_seleccionado:
    df_filtrado = df_filtrado[df_filtrado["Jefe de venta s"] == jefe_seleccionado]

vendedor_seleccionado = st.sidebar.selectbox("🧑‍💼 Selecciona un Vendedor", [None] + list(df_filtrado["Vendedor"].dropna().unique()))
if vendedor_seleccionado:
    df_filtrado = df_filtrado[df_filtrado["Vendedor"] == vendedor_seleccionado]

# =====================  Mostrar KPIs sin cálculos ===================== #

# Mostrar los valores de "Avance Ponderado Total" y "Cobertura del Tubo"
col1, col2 = st.columns(2)

with col1:
    if "AvancePonderadoTotal" in df_filtrado.columns:
        kpi_avance = df_filtrado["AvancePonderadoTotal"].iloc[0]  # Mostrar el primer valor de la columna filtrada
        st.metric("Avance Ponderado Total", kpi_avance)

with col2:
    if "Cobertura del tubo" in df_filtrado.columns:
        kpi_cobertura = df_filtrado["Cobertura del tubo"].iloc[0]  # Mostrar el primer valor de la columna filtrada
        st.metric("Cobertura del Tubo", kpi_cobertura)

# =====================  Agregar Comentarios a la Tabla ===================== #

# Agregar la columna de Comentarios si no existe
if "Comentarios" not in df_filtrado.columns:
    df_filtrado["Comentarios"] = ""

# Crear un formulario para agregar comentarios
for index, row in df_filtrado.iterrows():
    comentario = st.text_input(f"Comentario para {row['Vendedor']} (Semana: {row['Week Calendar']})", value=row["Comentarios"], key=index)
    df_filtrado.at[index, "Comentarios"] = comentario  # Actualizar el dataframe con el comentario ingresado

# =====================  Mostrar Tabla Completa con Comentarios ===================== #
st.markdown("### 📊 Tabla de Indicadores con Comentarios")
st.dataframe(df_filtrado, use_container_width=True)

# =====================  Guardar Comentarios en un Archivo Excel ===================== #
@st.cache_data
def guardar_comentarios(df, archivo_guardar):
    # Guardar el archivo Excel con los comentarios
    df.to_excel(archivo_guardar, index=False)
    return archivo_guardar

# Botón para guardar los comentarios en un archivo
if st.button("💾 Guardar Comentarios"):
    archivo_guardado = guardar_comentarios(df_filtrado, file_path_con_comentarios)
    st.success(f"Comentarios guardados exitosamente. Puedes descargar el archivo desde aquí:")
    st.markdown(f"[Haz clic aquí para descargar el archivo con comentarios guardados]({archivo_guardado})")
