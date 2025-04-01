import pandas as pd
import streamlit as st
import os
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import io
import time
import numpy as np

# ===================== Autenticación con contraseña inicial ===================== #
PASSWORD = "Minutas2025" 

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
    st.session_state.sharepoint_authenticated = False  # Nueva variable de estado

if not st.session_state.authenticated:
    st.title("🔒 Acceso Restringido")
    password_input = st.text_input("Ingrese la contraseña:", type="password")
    
    if st.button("🔑 Ingresar"):
        if password_input == PASSWORD:
            st.session_state.authenticated = True
            st.success("✅ Acceso permitido. Cargando la aplicación...")
            time.sleep(1)
            st.rerun()
        else:
            st.error("❌ Contraseña incorrecta. Intente nuevamente.")
    st.stop()

# ===================== Configuración de la página ===================== #
st.set_page_config(page_title="Dashboard CAR INDEX", layout="wide")

# ===================== Autenticación SharePoint Dinámica ===================== #
st.sidebar.markdown("## 🔑 Credenciales SharePoint")

if not st.session_state.sharepoint_authenticated:
    username = st.sidebar.text_input("Usuario SharePoint (ej: tuemail@salfa.cl)")
    password = st.sidebar.text_input("Contraseña SharePoint", type="password")
    
    if st.sidebar.button("🚀 Conectar a SharePoint"):
        if not username or not password:
            st.sidebar.warning("⚠️ Ingresa usuario y contraseña")
        else:
            try:
                # Credenciales de SharePoint
                site_url = "https://salfa.sharepoint.com/sites/MinutasCar"
                
                context = AuthenticationContext(site_url)
                if not context.acquire_token_for_user(username.strip(), password.strip()):
                    st.sidebar.error("❌ Error de autenticación. Verifica tus credenciales.")
                else:
                    st.session_state.ctx = ClientContext(site_url, context)
                    st.session_state.sharepoint_authenticated = True
                    st.sidebar.success("✅ Conectado a SharePoint!")
                    time.sleep(1)
                    st.rerun()
            except Exception as e:
                st.sidebar.error(f"❌ Error de conexión: {str(e)}")
    st.stop()  # Detiene la ejecución hasta autenticar SharePoint

# ===================== Conexión a SharePoint ===================== #
try:
    ctx = st.session_state.ctx
    status_container = st.empty()
    status_container.success("✅ Conectado a SharePoint. Cargando datos...")
    
    folder_url = "/sites/MinutasCar/Documentos%20compartidos"
    files = ctx.web.get_folder_by_server_relative_url(folder_url).files
    ctx.load(files)
    ctx.execute_query()

    # Buscar archivos necesarios
    car_index_csv = None
    mpp_xlsx = None
    for file in files:
        file_name = file.properties["Name"]
        file_url = file.properties["ServerRelativeUrl"]

        if file_name == "CAR INDEX.csv":
            car_index_csv = file_url
        elif file_name == "MPP.xlsx":
            mpp_xlsx = file_url

    if not car_index_csv:
        st.error("🚨 No se encontró el archivo 'CAR INDEX.csv' en SharePoint.")
        st.stop()

    # Leer CAR INDEX.csv
    try:
        response_csv = File.open_binary(ctx, car_index_csv)
        df_csv = pd.read_csv(io.BytesIO(response_csv.content), encoding="utf-8", sep=";", engine="python", dtype=str)
        df_csv.columns = df_csv.columns.str.strip()
    except Exception as e:
        st.error(f"❌ Error al leer el archivo CSV: {e}")
        st.stop()

    # Leer MPP.xlsx si existe
    df_mpp = None
    if mpp_xlsx:
        try:
            response_mpp = File.open_binary(ctx, mpp_xlsx)
            df_mpp = pd.read_excel(io.BytesIO(response_mpp.content), engine='openpyxl')
            df_mpp.columns = df_mpp.columns.str.strip()
        except Exception as e:
            st.warning(f"⚠️ Error al leer MPP.xlsx: {e}")

except Exception as e:
    st.error(f"❌ Error al conectar con SharePoint: {e}")
    st.stop()
# ===================== DISEÑO DEL DASHBOARD ===================== #
st.markdown("""
    <style>
        .big-title {
            font-size: 48px;
            font-weight: bold;
            text-align: center;
            color: #4A90E2;
            text-shadow: 3px 3px 7px rgba(0, 0, 0, 0.3);
            animation: fadeIn 1.2s ease-in-out;
        }
        
        .sub-title {
            font-size: 36px;
            font-weight: bold;
            text-align: center;
            color: #222;
            text-shadow: 2px 2px 5px rgba(0, 0, 0, 0.2);
            margin-top: -10px;
            animation: fadeIn 1.5s ease-in-out;
        }

        .custom-box {
            border: 2px solid #4A90E2;
            border-radius: 10px;
            padding: 15px;
            margin: 10px 0;
            background-color: #f9f9f9;
            box-shadow: 3px 3px 10px rgba(0, 0, 0, 0.1);
        }

        .section-title {
            font-size: 28px;
            font-weight: bold;
            color: #4A90E2;
            margin-bottom: 10px;
        }
        
        .filter-box {
            background-color: #f0f2f6;
            border-radius: 8px;
            padding: 12px;
            margin-bottom: 15px;
        }

        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(-15px); }
            to { opacity: 1; transform: translateY(0); }
        }
        
        .divider {
            border-top: 3px solid #4A90E2;
            margin: 15px 0;
        }
        
        .metric-value {
            font-size: 24px !important;
        }
    </style>
""", unsafe_allow_html=True)

# Títulos estilizados
st.markdown('<p class="big-title">📊 Feedback Car</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">Compromisos - Acciones - Resultados</p>', unsafe_allow_html=True)
st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

# ===================== Funciones auxiliares mejoradas ===================== #
def convertir_a_float(valor):
    """Convierte un valor a float, manejando casos especiales como NaN, None o strings vacíos"""
    if pd.isna(valor) or valor is None or str(valor).strip() in ['', 'N/A', 'NaN', 'nan']:
        return float('nan')
    try:
        valor = str(valor).replace('%', '').replace(',', '.').strip()
        return float(valor) if valor else float('nan')
    except ValueError:
        return float('nan')

def normalizar_valor(valor, umbral):
    """Normaliza un valor respecto a un umbral, manejando NaN"""
    if pd.isna(valor):
        return 0
    if valor >= umbral:
        return 100
    else:
        return (valor / umbral) * 100

def mostrar_nota(nota):
    """Muestra la nota como pelotitas, manejando valores no válidos"""
    if pd.isna(nota) or nota < 1:
        return "⚪⚪⚪⚪⚪ (s/o)"
    pelotita_llena = "🔵"
    pelotita_vacia = "⚪"
    nota_int = int(nota)
    pelotitas = (pelotita_llena * nota_int) + (pelotita_vacia * (5 - nota_int))
    return pelotitas

# ===================== Filtros CAR INDEX ===================== #
st.markdown("### 🔍 Filtros CAR INDEX")
with st.expander("Mostrar/Ocultar Filtros", expanded=True):
    col_f1, col_f2, col_f3 = st.columns(3)
    
    with col_f1:
        jefes_ventas = [
            None,
            "Alexis Mollo", "Cristhian Veragua", "Cristian Cortes", "Cristian Rojo", 
            "Felipe Carrasco", "Felipe Castro", "Gonzalo Robles", "Iván Vásquez", 
            "Jhonatan Salini", "Jorge Rubilar", "Jose Moreno", "María José Briones", 
            "Meliza Angel", "Oscar Lazo", "Rene Nova", "Ricardo Riquelme", 
            "Roy Luque Bou Daher", "Sergio Tapia"
        ]

        jefe_seleccionado = st.selectbox("👔 Jefe de Venta", jefes_ventas)
    
    with col_f2:
        if jefe_seleccionado:
            ym_options = [None] + list(df_csv[df_csv["Jefe de venta s"] == jefe_seleccionado]["YM"].dropna().unique())
        else:
            ym_options = [None] + list(df_csv["YM"].dropna().unique())
        ym_seleccionado = st.selectbox("📅 YM", ym_options)
    
    with col_f3:
        if jefe_seleccionado and ym_seleccionado:
            week_options = [None] + list(df_csv[(df_csv["Jefe de venta s"] == jefe_seleccionado) & 
                           (df_csv["YM"] == ym_seleccionado)]["Week Calendar"].dropna().unique())
        elif jefe_seleccionado:
            week_options = [None] + list(df_csv[df_csv["Jefe de venta s"] == jefe_seleccionado]["Week Calendar"].dropna().unique())
        else:
            week_options = [None] + list(df_csv["Week Calendar"].dropna().unique())
        week_calendar_seleccionado = st.selectbox("📆 Week Calendar", week_options)

    vendedor_options = [None]
    if jefe_seleccionado:
        if ym_seleccionado and week_calendar_seleccionado:
            vendedor_options += list(df_csv[(df_csv["Jefe de venta s"] == jefe_seleccionado) & 
                                          (df_csv["YM"] == ym_seleccionado) & 
                                          (df_csv["Week Calendar"] == week_calendar_seleccionado)]["Vendedor"].dropna().unique())
        elif ym_seleccionado:
            vendedor_options += list(df_csv[(df_csv["Jefe de venta s"] == jefe_seleccionado) & 
                                          (df_csv["YM"] == ym_seleccionado)]["Vendedor"].dropna().unique())
        else:
            vendedor_options += list(df_csv[df_csv["Jefe de venta s"] == jefe_seleccionado]["Vendedor"].dropna().unique())
    
    vendedor_seleccionado = st.selectbox("👨‍💼 Vendedor", vendedor_options)

# Aplicar filtros
df_filtrado = df_csv.copy()
if jefe_seleccionado:
    df_filtrado = df_filtrado[df_filtrado["Jefe de venta s"] == jefe_seleccionado]
if ym_seleccionado:
    df_filtrado = df_filtrado[df_filtrado["YM"] == ym_seleccionado]
if week_calendar_seleccionado:
    df_filtrado = df_filtrado[df_filtrado["Week Calendar"] == week_calendar_seleccionado]
if vendedor_seleccionado:
    df_filtrado = df_filtrado[df_filtrado["Vendedor"] == vendedor_seleccionado]

# ===================== Secciones en Columnas ===================== #
st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
st.markdown('<p class="section-title">📊 KPIs Principales</p>', unsafe_allow_html=True)

col1, col2, col3 = st.columns(3)

# Sección 1: Tratamiento Leads
with col1:
    st.markdown('<div class="custom-box">', unsafe_allow_html=True)
    st.markdown("#### 📊 Tratamiento Leads")
    
    # KPI: Actividad Primer Intento
    kpi_actividad = df_filtrado["actividad primer intento"].iloc[0] if "actividad primer intento" in df_filtrado.columns else "N/A"
    kpi_actividad_float = convertir_a_float(kpi_actividad)
    
    if pd.isna(kpi_actividad_float):
        valor_normalizado_actividad = 0
        nota_actividad = 1
        actividad_icon = "⚠️"
        valor_mostrado = "s/o"
    else:
        valor_normalizado_actividad = normalizar_valor(kpi_actividad_float, 50)
        nota_actividad = max(1, min(int((valor_normalizado_actividad / 100) * 5), 5))
        actividad_icon = ":green[✔️]" if kpi_actividad_float >= 50 else ":red[❌]"
        valor_mostrado = f"{kpi_actividad_float:.2f}%"
    
    st.markdown(f"**Actividad Primer Intento** {actividad_icon}")
    st.metric("Valor", valor_mostrado, label_visibility="collapsed")
    st.markdown(f"**Nota:** {mostrar_nota(nota_actividad)}")
    st.markdown("---")

    # KPI: Conversión de Leads
    kpi_conversion = df_filtrado["Conversión de leads"].iloc[0] if "Conversión de leads" in df_filtrado.columns else "N/A"
    kpi_conversion_float = convertir_a_float(kpi_conversion)
    
    if pd.isna(kpi_conversion_float):
        valor_normalizado_conversion = 0
        nota_conversion = 1
        conversion_icon = "⚠️"
        valor_mostrado = "s/o"
    else:
        valor_normalizado_conversion = normalizar_valor(kpi_conversion_float, 45)
        nota_conversion = max(1, min(int((valor_normalizado_conversion / 100) * 5), 5))
        conversion_icon = ":green[✔️]" if kpi_conversion_float >= 45 else ":red[❌]"
        valor_mostrado = f"{kpi_conversion_float:.2f}%"
    
    st.markdown(f"**Conversión de Leads** {conversion_icon}")
    st.metric("Valor", valor_mostrado, label_visibility="collapsed")
    st.markdown(f"**Nota:** {mostrar_nota(nota_conversion)}")
    st.markdown('</div>', unsafe_allow_html=True)

# Sección 2: Cobertura del Tubo
with col2:
    st.markdown('<div class="custom-box">', unsafe_allow_html=True)
    st.markdown("#### 📈 Cobertura del Tubo")
    
    # KPI: Cobertura del Tubo
    kpi_cobertura = df_filtrado["Cobertura del tubo"].iloc[0] if "Cobertura del tubo" in df_filtrado.columns else "N/A"
    kpi_cobertura_float = convertir_a_float(kpi_cobertura)
    
    if pd.isna(kpi_cobertura_float):
        valor_normalizado_cobertura = 0
        nota_cobertura = 1
        cobertura_icon = "⚠️"
        valor_mostrado = "s/o"
    else:
        valor_normalizado_cobertura = normalizar_valor(kpi_cobertura_float, 100)
        nota_cobertura = 5 if kpi_cobertura_float == 100 else 3
        cobertura_icon = ":green[✔️]" if kpi_cobertura_float == 100 else ":red[❌]"
        valor_mostrado = f"{kpi_cobertura_float:.2f}%"
    
    st.markdown(f"**Cobertura del Tubo** {cobertura_icon}")
    st.metric("Valor", valor_mostrado, label_visibility="collapsed")
    st.markdown(f"**Nota:** {mostrar_nota(nota_cobertura)}")
    st.markdown('</div>', unsafe_allow_html=True)

# Sección 3: Avance Forecast Semanal
with col3:
    st.markdown('<div class="custom-box">', unsafe_allow_html=True)
    st.markdown("#### 📅 Avance Forecast Semanal")
    st.markdown("🔜 Próximamente...")
    st.markdown('</div>', unsafe_allow_html=True)

# ===================== Tabla CAR INDEX ===================== #


# ===================== Acumulación de Comentarios CAR INDEX ===================== #
if jefe_seleccionado:
    # Si no existe la columna de comentarios, crearla
    if "Comentarios" not in df_filtrado.columns:
        df_filtrado.loc[:, "Comentarios"] = ""

    # Acumulación de comentarios
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown('<p class="section-title">💬 Comentarios CAR INDEX</p>', unsafe_allow_html=True)
    
    for index, row in df_filtrado.iterrows():
        # Crear un campo para el nuevo comentario
        comentario_nuevo = st.text_area(f"Comentario para {row['Vendedor']} (Semana: {row['Week Calendar']})", 
                                       key=f"car_{index}")
        
        # Usar `.loc` para modificar el DataFrame de manera segura
        df_filtrado.loc[index, "Comentarios"] = comentario_nuevo

# ===================== Función para guardar comentarios CAR INDEX ===================== #
# ===================== Función para guardar comentarios CAR INDEX ===================== #
def guardar_comentarios_car_en_sharepoint(df_nuevos, ctx, jefe_seleccionado):
    """Guarda los comentarios de CAR INDEX en la carpeta 'car index' de SharePoint con todas las columnas"""
    # Definir las columnas requeridas fuera del bloque try para que esté disponible en todo el ámbito
    columnas_requeridas = [
        "YM", "Week Calendar", "Sucursal", "Jefe de venta s", "Vendedor",
        "Comentarios", "actividad primer intento", "Porcentaje primer intento",
        "Conversión de leads", "PorcentajeDeAvanceConversión", "AvancePonderadoTotal",
        "on con gestion 100%", "Cobertura del tubo"
    ]
    
    try:
        # Definir la ruta y nombre del archivo
        archivo = f"CAR_INDEX_con_comentarios_{jefe_seleccionado}.xlsx"
        folder_url = "/sites/MinutasCar/Documentos compartidos/car index"
        
        # Obtener referencia a la carpeta
        folder = ctx.web.get_folder_by_server_relative_url(folder_url)
        ctx.load(folder)
        ctx.execute_query()
        
        # Verificar si el archivo ya existe
        file_url = f"{folder_url}/{archivo}"
        file = ctx.web.get_file_by_server_relative_url(file_url)
        
        try:
            # Intentar cargar archivo existente
            file_content = io.BytesIO()
            file.download(file_content).execute_query()
            file_content.seek(0)
            df_existente = pd.read_excel(file_content, engine='openpyxl')
            
            # Asegurar que el archivo existente tenga todas las columnas requeridas
            for col in columnas_requeridas:
                if col not in df_existente.columns:
                    df_existente[col] = "" if col != "Comentarios" else np.nan
        except:
            # Crear nuevo DataFrame con todas las columnas requeridas
            df_existente = pd.DataFrame(columns=columnas_requeridas)
        
        # Asegurar que el nuevo DataFrame tenga todas las columnas requeridas
        for col in columnas_requeridas:
            if col not in df_nuevos.columns:
                df_nuevos[col] = "" if col != "Comentarios" else np.nan
        
        # Combinar con nuevos datos (conservando todas las columnas)
        df_actualizado = pd.concat([df_existente, df_nuevos], ignore_index=True)
        
        # Ordenar columnas según el formato requerido
        df_actualizado = df_actualizado[columnas_requeridas]
        
        # Guardar en memoria
        output = io.BytesIO()
        df_actualizado.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)
        
        # Subir a SharePoint
        folder.upload_file(archivo, output).execute_query()
        
        return True
    except Exception as e:
        st.error(f"Error al guardar comentarios CAR INDEX: {str(e)}")
        return False

# ===================== Botón para guardar comentarios CAR INDEX ===================== #
if jefe_seleccionado and st.button("💾 Guardar Comentarios CAR INDEX"):
    # Definir columnas requeridas
    columnas_requeridas = [
        "YM", "Week Calendar", "Sucursal", "Jefe de venta s", "Vendedor",
        "Comentarios", "actividad primer intento", "Porcentaje primer intento",
        "Conversión de leads", "PorcentajeDeAvanceConversión", "AvancePonderadoTotal",
        "on con gestion 100%", "Cobertura del tubo"
    ]
    
    # Crear DataFrame con todas las columnas
    df_para_guardar = df_filtrado.copy()
    
    # Asegurar que todas las columnas existan
    for col in columnas_requeridas:
        if col not in df_para_guardar.columns:
            df_para_guardar[col] = "" if col != "Comentarios" else np.nan
    
    # Seleccionar solo las columnas requeridas
    df_para_guardar = df_para_guardar[columnas_requeridas]
    
    if guardar_comentarios_car_en_sharepoint(df_para_guardar, ctx, jefe_seleccionado):
        st.success(f"✅ Comentarios CAR INDEX para {jefe_seleccionado} guardados exitosamente en la carpeta 'car index'.")
    else:
        st.error("❌ Error al guardar los comentarios CAR INDEX")

# ===================== Sección MPP ===================== #
# ===================== Sección MPP ===================== #
if df_mpp is not None:
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    st.markdown('<p class="section-title">📌 Gestión integral de ventas: Desempeño de adicionales</p>', unsafe_allow_html=True)
    
    # Filtros MPP
    st.markdown("### 🔍 Filtros")
    with st.expander("Mostrar/Ocultar Filtros", expanded=True):
        col_mf1, col_mf2 = st.columns(2)
        
        with col_mf1:
            jefes_mpp = [None] + list(df_mpp["Jefes de Venta"].dropna().unique())
            jefe_seleccionado_mpp = st.selectbox("👔 Jefe de Venta", jefes_mpp, key="jefe_mpp")
            
            if jefe_seleccionado_mpp:
                sucursales_mpp = [None] + list(df_mpp[df_mpp["Jefes de Venta"] == jefe_seleccionado_mpp]["Sucursal Autycam"].dropna().unique())
            else:
                sucursales_mpp = [None] + list(df_mpp["Sucursal Autycam"].dropna().unique())
            sucursal_seleccionada_mpp = st.selectbox("🏢 Sucursal", sucursales_mpp, key="sucursal_mpp")
        
        with col_mf2:
            años_mpp = [None] + list(df_mpp["Date - Año"].dropna().unique())
            año_seleccionado_mpp = st.selectbox("📅 Año", años_mpp, key="año_mpp")
            
            if año_seleccionado_mpp:
                meses_mpp = [None] + list(df_mpp[df_mpp["Date - Año"] == año_seleccionado_mpp]["Date - Mes"].dropna().unique())
            else:
                meses_mpp = [None] + list(df_mpp["Date - Mes"].dropna().unique())
            mes_seleccionado_mpp = st.selectbox("📆 Mes", meses_mpp, key="mes_mpp")
    
    # Nuevo filtro para Vendedor Generico
    vendedor_options_mpp = [None]
    if jefe_seleccionado_mpp or sucursal_seleccionada_mpp or año_seleccionado_mpp or mes_seleccionado_mpp:
        temp_df = df_mpp.copy()
        if jefe_seleccionado_mpp:
            temp_df = temp_df[temp_df["Jefes de Venta"] == jefe_seleccionado_mpp]
        if sucursal_seleccionada_mpp:
            temp_df = temp_df[temp_df["Sucursal Autycam"] == sucursal_seleccionada_mpp]
        if año_seleccionado_mpp:
            temp_df = temp_df[temp_df["Date - Año"] == año_seleccionado_mpp]
        if mes_seleccionado_mpp:
            temp_df = temp_df[temp_df["Date - Mes"] == mes_seleccionado_mpp]
        
        vendedor_options_mpp += list(temp_df["Vendedor Generico"].dropna().unique())
    
    vendedor_seleccionado_mpp = st.selectbox("👨‍💼 Vendedor Generico", vendedor_options_mpp, key="vendedor_mpp")

    # Aplicar filtros MPP
    df_mpp_filtrado = df_mpp.copy()
    if jefe_seleccionado_mpp:
        df_mpp_filtrado = df_mpp_filtrado[df_mpp_filtrado["Jefes de Venta"] == jefe_seleccionado_mpp]
    if sucursal_seleccionada_mpp:
        df_mpp_filtrado = df_mpp_filtrado[df_mpp_filtrado["Sucursal Autycam"] == sucursal_seleccionada_mpp]
    if año_seleccionado_mpp:
        df_mpp_filtrado = df_mpp_filtrado[df_mpp_filtrado["Date - Año"] == año_seleccionado_mpp]
    if mes_seleccionado_mpp:
        df_mpp_filtrado = df_mpp_filtrado[df_mpp_filtrado["Date - Mes"] == mes_seleccionado_mpp]
    if vendedor_seleccionado_mpp:
        df_mpp_filtrado = df_mpp_filtrado[df_mpp_filtrado["Vendedor Generico"] == vendedor_seleccionado_mpp]
    
    # Mostrar KPIs MPP
    st.markdown("### 📊 KPIs Desempeño de adicionales")
    col_mpp1, col_mpp2, col_mpp3, col_mpp4 = st.columns(4)  # Cambiado a 4 columnas

    with col_mpp1:
        st.markdown('<div class="custom-box">', unsafe_allow_html=True)
        st.markdown("##### 💳 Penetración Créditos 1ra")
        kpi_p1 = df_mpp_filtrado["Penetracion Creditos 1ra"].mean()
        st.metric("Valor", f"{kpi_p1:.2f}%" if not pd.isna(kpi_p1) else "s/o", label_visibility="collapsed")
        st.markdown('</div>', unsafe_allow_html=True)

    with col_mpp2:
        st.markdown('<div class="custom-box">', unsafe_allow_html=True)
        st.markdown("##### 💰 Penetración Créditos 2da")
        kpi_p2 = df_mpp_filtrado["Penetracion Creditos 2da"].mean()
        st.metric("Valor", f"{kpi_p2:.2f}%" if not pd.isna(kpi_p2) else "s/o", label_visibility="collapsed")
        st.markdown('</div>', unsafe_allow_html=True)

    with col_mpp3:
        st.markdown('<div class="custom-box">', unsafe_allow_html=True)
        st.markdown("##### 🛡️ Penetración Seguros")
        kpi_seg = df_mpp_filtrado["Penetracion seguros"].mean()
        st.metric("Valor", f"{kpi_seg:.2f}%" if not pd.isna(kpi_seg) else "s/o", label_visibility="collapsed")
        st.markdown('</div>', unsafe_allow_html=True)

    with col_mpp4:
        st.markdown('<div class="custom-box">', unsafe_allow_html=True)
        st.markdown("##### 📈 Penetración MPP")
        kpi_mpp = df_mpp_filtrado["Penetracion MPP"].mean()
        st.metric("Valor", f"{kpi_mpp:.2f}%" if not pd.isna(kpi_mpp) else "s/o", label_visibility="collapsed")
        st.markdown('</div>', unsafe_allow_html=True)
    


    # ===================== Acumulación de Comentarios MPP ===================== #
    if jefe_seleccionado_mpp:
        # Si no existe la columna de comentarios, crearla
        if "Comentarios" not in df_mpp_filtrado.columns:
            df_mpp_filtrado.loc[:, "Comentarios"] = ""

        # Acumulación de comentarios MPP
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        st.markdown('<p class="section-title">💬 Comentarios MPP</p>', unsafe_allow_html=True)
        
        for index, row in df_mpp_filtrado.iterrows():
            # Crear un campo para el nuevo comentario
            comentario_nuevo = st.text_area(
                f"Comentario MPP para {row['Vendedor Generico']} ({row['Date - Mes']}/{row['Date - Año']})", 
                key=f"mpp_{index}"
            )
            
            # Usar `.loc` para modificar el DataFrame de manera segura
            df_mpp_filtrado.loc[index, "Comentarios"] = comentario_nuevo

# ===================== Función para guardar comentarios MPP ===================== #
def guardar_comentarios_mpp_en_sharepoint(df_nuevos, ctx, jefe_seleccionado_mpp):
    """Guarda los comentarios de MPP en la carpeta 'comentarios' de SharePoint"""
    try:
        # Definir la ruta y nombre del archivo
        archivo = f"MPP_con_comentarios_{jefe_seleccionado_mpp}.xlsx"
        folder_url = "/sites/MinutasCar/Documentos compartidos/comentarios"
        
        # Obtener referencia a la carpeta
        folder = ctx.web.get_folder_by_server_relative_url(folder_url)
        ctx.load(folder)
        ctx.execute_query()
        
        # Verificar si el archivo ya existe
        file_url = f"{folder_url}/{archivo}"
        file = ctx.web.get_file_by_server_relative_url(file_url)
        
        try:
            # Intentar cargar archivo existente
            file_content = io.BytesIO()
            file.download(file_content).execute_query()
            file_content.seek(0)
            df_existente = pd.read_excel(file_content, engine='openpyxl')
        except:
            # Crear nuevo DataFrame si no existe con todas las columnas originales
            columnas = [
                "Jefes de Venta", "Sucursal Autycam", "Vendedor Generico",
                "Créditos marca", "Penetracion Creditos 1ra", 
                "Créditos 2da Op", "Penetracion Creditos 2da",
                "Créditos totales", "Penetracion Creditos Retail",
                "Seguros", "Penetracion seguros",
                "MPP", "Penetracion MPP",
                "Date - Año", "Date - Mes", "Comentarios"
            ]
            df_existente = pd.DataFrame(columns=columnas)
        
        # Combinar con nuevos datos (conservando todas las columnas)
        df_actualizado = pd.concat([df_existente, df_nuevos], ignore_index=True)
        
        # Guardar en memoria
        output = io.BytesIO()
        df_actualizado.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)
        
        # Subir a SharePoint
        folder.upload_file(archivo, output).execute_query()
        
        return True
    except Exception as e:
        st.error(f"Error al guardar comentarios MPP: {str(e)}")
        return False

# ===================== Botón para guardar comentarios MPP ===================== #
if jefe_seleccionado_mpp and st.button("💾 Guardar Comentarios MPP"):
    # Conservar todas las columnas originales más los comentarios
    columnas_guardar = [
        "Jefes de Venta", "Sucursal Autycam", "Vendedor Generico",
        "Créditos marca", "Penetracion Creditos 1ra", 
        "Créditos 2da Op", "Penetracion Creditos 2da",
        "Créditos totales", "Penetracion Creditos Retail",
        "Seguros", "Penetracion seguros",
        "MPP", "Penetracion MPP",
        "Date - Año", "Date - Mes", "Comentarios"
    ]
    
    # Asegurarse de que todas las columnas existan en el DataFrame
    df_para_guardar = df_mpp_filtrado.copy()
    for col in columnas_guardar:
        if col not in df_para_guardar.columns and col != "Comentarios":
            df_para_guardar[col] = ""  # O puedes usar np.nan si prefieres
    
    # Reordenar columnas según el formato deseado
    df_para_guardar = df_para_guardar[columnas_guardar]
    
    if guardar_comentarios_mpp_en_sharepoint(df_para_guardar, ctx, jefe_seleccionado_mpp):
        st.success(f"✅ Comentarios MPP para {jefe_seleccionado_mpp} guardados exitosamente.")
    else:
        st.error("❌ Error al guardar los comentarios MPP")
