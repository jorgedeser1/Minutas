import pandas as pd
import streamlit as st
import os

# Ruta del archivo CSV
file_path = r"C:\Users\cokeh\Desktop\car index minutas\CAR INDEX.csv"
file_path_con_comentarios = r"C:\Users\cokeh\Desktop\car index minutas\CAR_INDEX_con_comentarios.xlsx"  # Archivo donde se guardan los comentarios

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
                       "actividad primer intento", "Porcentaje primer intento", "Conversión de leads", 
                       "PorcentajeDeAvanceConversión", "AvancePonderadoTotal", "on con gestion 100%", 
                       "Cobertura del tubo"]
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

# Mostrar los valores de "Avance Ponderado Total", "Cobertura del Tubo" y "Cumplimiento Minutas"
col1, col2, col3 = st.columns(3)

with col1:
    if "AvancePonderadoTotal" in df_filtrado.columns:
        kpi_avance = df_filtrado["AvancePonderadoTotal"].iloc[0]  # Mostrar el primer valor de la columna filtrada
        st.metric("Avance Ponderado Total", kpi_avance)

# Mostrar el KPI de **Cobertura del Tubo**
with col2:
    kpi_cobertura = "N/A"
    if "Cobertura del tubo" in df_filtrado.columns:
        kpi_cobertura = df_filtrado["Cobertura del tubo"].iloc[0]
    st.metric("Cobertura del Tubo", kpi_cobertura)

# Mostrar el KPI de **Cumplimiento Minutas**
with col3:
    kpi_cumplimiento = 0  # Por defecto, se setea a 0
    if os.path.exists(file_path_con_comentarios):
        df_comentarios_guardados = pd.read_excel(file_path_con_comentarios)
        
        # Filtrar las semanas con comentarios
        semanas_completadas = df_comentarios_guardados[df_comentarios_guardados["Comentarios"].notna()]["Week Calendar"].unique()
        
        # Total de semanas posibles
        semanas_totales = df_comentarios_guardados["Week Calendar"].nunique()
        
        if semanas_totales > 0:
            # Calcular el porcentaje de cumplimiento
            kpi_cumplimiento = (len(semanas_completadas) / semanas_totales) * 100
    
    st.metric("KPI Cumplimiento Minutas", f"{kpi_cumplimiento:.2f}%")

# ===================== Agregar Comentarios a la Tabla ===================== #

# Agregar la columna de Comentarios si no existe
if "Comentarios" not in df_filtrado.columns:
    df_filtrado["Comentarios"] = ""

# Mostrar el formulario de comentarios solo si se ha seleccionado un vendedor
if vendedor_seleccionado:
    for index, row in df_filtrado.iterrows():
        comentario = st.text_input(f"Comentario para {row['Vendedor']} (Semana: {row['Week Calendar']})", 
                                  value=row["Comentarios"], key=index)
        # Actualizar el dataframe con el comentario ingresado, manteniendo la fila original
        df_filtrado.at[index, "Comentarios"] = comentario

# ===================== Mostrar Tabla Completa con Comentarios ===================== #
st.markdown("### 📊 Tabla de Indicadores con Comentarios")
st.dataframe(df_filtrado, use_container_width=True)

# ===================== Guardar Comentarios en un Archivo Excel ===================== #

def guardar_comentarios_acumulados(df, archivo_guardar):
    # Verificar si ya existe un archivo con comentarios previos
    if os.path.exists(archivo_guardar):
        df_comentarios_guardados = pd.read_excel(archivo_guardar)  # Leer el archivo con comentarios anteriores
        # Asegurarnos de que las columnas del dataframe original y el de los comentarios coincidan
        df_comentarios_guardados = pd.concat([df_comentarios_guardados, df[["YM", "Week Calendar", "Sucursal", "Jefe de venta s", "Vendedor", 
                                                                          "actividad primer intento", "Porcentaje primer intento", "Conversión de leads", 
                                                                          "PorcentajeDeAvanceConversión", "AvancePonderadoTotal", "on con gestion 100%", 
                                                                          "Cobertura del tubo", "Comentarios"]]], ignore_index=True)
    else:
        # Si no existe, simplemente asignar los datos actuales
        df_comentarios_guardados = df[["YM", "Week Calendar", "Sucursal", "Jefe de venta s", "Vendedor", 
                                       "actividad primer intento", "Porcentaje primer intento", "Conversión de leads", 
                                       "PorcentajeDeAvanceConversión", "AvancePonderadoTotal", "on con gestion 100%", 
                                       "Cobertura del tubo", "Comentarios"]]
    
    # Guardar el archivo Excel con los comentarios acumulados
    df_comentarios_guardados.to_excel(archivo_guardar, index=False)
    
    return archivo_guardar  # Devolver el archivo guardado

# Botón para guardar los comentarios en un archivo (ubicado después de la tabla)
st.markdown("### 💾 Guardar Comentarios")
if st.button("Guardar Comentarios"):
    archivo_guardado = guardar_comentarios_acumulados(df_filtrado, file_path_con_comentarios)
    st.success(f"Comentarios guardados exitosamente. Puedes descargar el archivo desde aquí:")
    st.markdown(f"[Haz clic aquí para descargar el archivo con comentarios guardados]({archivo_guardado})")
