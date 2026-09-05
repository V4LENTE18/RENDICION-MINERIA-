import streamlit as st
import pandas as pd
import plotly.express as px
from streamlit_gsheets import GSheetsConnection
import os
import io
from datetime import date

# ==========================================
# CONFIGURACIÓN DE LA PÁGINA (DISEÑO RESPONSIVO)
# ==========================================
st.set_page_config(
    page_title="Rendición Minera",
    page_icon="⛏️",
    layout="wide",
    initial_sidebar_state="collapsed" # Inicia cerrado en móviles para dar más espacio
)

# Estilo CSS adicional para mejorar el aspecto en móviles
st.markdown("""
    <style>
    .stButton>button { border-radius: 8px; font-weight: bold; }
    .stMetric { background-color: #f8f9fa; padding: 10px; border-radius: 10px; border: 1px solid #e9ecef; }
    </style>
""", unsafe_allow_html=True)

CATEGORIAS = [
    "ABARROTES E IMPLEMENTOS DE COCINA",
    "VERDURAS Y FRUTAS",
    "HERRAMIENTAS Y EPPS",
    "COMBUSTIBLE",
    "GASTOS ADMINISTRATIVOS",
    "GASTOS ANÁLISIS",
    "INSUMOS",
    "PAGOS TRANQUERA"
]

NOMBRE_PESTAÑA = "Hoja 1"

if "lista_productos" not in st.session_state:
    st.session_state.lista_productos = []

# ==========================================
# CONEXIÓN Y FUNCIONES DE BASE DE DATOS (MÁS DE 1000 ITEMS)
# ==========================================
conn = st.connection("gsheets", type=GSheetsConnection)

def cargar_datos():
    try:
        df = conn.read(worksheet=NOMBRE_PESTAÑA, ttl=0)
        columnas_requeridas = ["ID", "Fecha", "Categoría", "N° Serie", "Descripción", "Cantidad", "Unidad", "Precio Unitario", "Total"]
        
        if df is None or df.empty:
            return pd.DataFrame(columns=columnas_requeridas)
        
        for col in columnas_requeridas:
            if col not in df.columns:
                df[col] = "-" if col in ["N° Serie", "Unidad", "Descripción", "Categoría"] else 0

        # Forzar conversión limpia
        df['ID'] = pd.to_numeric(df['ID'], errors='coerce').fillna(0).astype(int)
        df['Cantidad'] = pd.to_numeric(df['Cantidad'], errors='coerce').fillna(0.0)
        df['Precio Unitario'] = pd.to_numeric(df['Precio Unitario'], errors='coerce').fillna(0.0)
        df['Total'] = pd.to_numeric(df['Total'], errors='coerce').fillna(0.0)
        
        df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
        df = df.sort_values(by="Fecha", ascending=False).reset_index(drop=True)
        df['Fecha'] = df['Fecha'].dt.date.fillna(date.today())
        
        df = df[df['Descripción'].astype(str).str.strip() != ""]
        return df[columnas_requeridas]
    except Exception as e:
        return pd.DataFrame(columns=["ID", "Fecha", "Categoría", "N° Serie", "Descripción", "Cantidad", "Unidad", "Precio Unitario", "Total"])

def guardar_datos(df):
    try:
        df_a_guardar = df.copy()
        
        # LIMPIEZA PROFUNDA (Clave para evitar bloqueos con >1000 filas en Google Sheets)
        df_a_guardar['Fecha'] = pd.to_datetime(df_a_guardar['Fecha'], errors='coerce').dt.strftime('%Y-%m-%d')
        df_a_guardar['Fecha'] = df_a_guardar['Fecha'].fillna(date.today().strftime('%Y-%m-%d'))
        
        # Convertir tipos nativos de Python para evitar errores de serialización JSON de Pandas
        df_a_guardar['ID'] = df_a_guardar['ID'].astype(int)
        df_a_guardar['Descripción'] = df_a_guardar['Descripción'].astype(str)
        df_a_guardar['Categoría'] = df_a_guardar['Categoría'].astype(str)
        df_a_guardar['N° Serie'] = df_a_guardar['N° Serie'].astype(str)
        df_a_guardar['Unidad'] = df_a_guardar['Unidad'].astype(str)
        
        df_a_guardar['Cantidad'] = df_a_guardar['Cantidad'].astype(float)
        df_a_guardar['Precio Unitario'] = df_a_guardar['Precio Unitario'].astype(float)
        df_a_guardar['Total'] = df_a_guardar['Total'].astype(float)
        
        df_a_guardar = df_a_guardar.fillna("") # Eliminar cualquier Null fantasma
        
        conn.update(worksheet=NOMBRE_PESTAÑA, data=df_a_guardar)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"Error técnico al guardar: {e}")
        return False

# ==========================================
# GENERACIÓN DE EXCEL OFICIAL
# ==========================================
def generar_excel_dinamico(df, periodo, presidente, tesorero, fiscal):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        formato_titulo_amarillo = workbook.add_format({'bold': True, 'font_size': 16, 'bg_color': '#FFFF00', 'color': 'black', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        formato_titulo_naranja = workbook.add_format({'bold': True, 'font_size': 14, 'bg_color': '#FFC000', 'color': 'black', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        formato_directiva = workbook.add_format({'bold': True, 'font_size': 12, 'bg_color': '#FFFFFF', 'color': 'black', 'border': 1, 'align': 'left', 'valign': 'vcenter'})
        formato_encabezado_tabla = workbook.add_format({'bold': True, 'bg_color': '#1E3A8A', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        formato_moneda = workbook.add_format({'num_format': 'S/ #,##0.00'})
        formato_total = workbook.add_format({'bold': True, 'bg_color': '#F3F4F6', 'num_format': 'S/ #,##0.00', 'border': 1})

        def dibujar_encabezado_oficial(worksheet):
            worksheet.set_column('A:H', 15) 
            worksheet.set_row(0, 45) 
            worksheet.set_row(2, 25)
            worksheet.set_row(4, 25)
            worksheet.merge_range('A1:H1', 'SOCIEDAD MINERA REY', formato_titulo_amarillo)
            worksheet.merge_range('A3:H3', 'RENDICIÓN DE CUENTAS', formato_titulo_amarillo)
            worksheet.merge_range('A5:H5', periodo.upper(), formato_titulo_naranja)
            worksheet.merge_range('A7:H7', f'PRESIDENTE: {presidente.upper()}', formato_directiva)
            worksheet.merge_range('A8:H8', f'TESORERO: {tesorero.upper()}', formato_directiva)
            worksheet.merge_range('A9:H9', f'FISCAL: {fiscal.upper()}', formato_directiva)
            
            logo_path = 'logo.png'
            if os.path.exists(logo_path):
                opciones_logo = {'x_scale': 0.6, 'y_scale': 0.6, 'x_offset': 10, 'y_offset': 5, 'object_position': 1}
                worksheet.insert_image('A1', logo_path, opciones_logo)
                worksheet.insert_image('H1', logo_path, opciones_logo)

        worksheet_dash = workbook.add_worksheet('DASHBOARD')
        dibujar_encabezado_oficial(worksheet_dash)
        worksheet_dash.set_column('A:A', 35)
        worksheet_dash.set_column('B:B', 20)
        
        if not df.empty:
            resumen = df.groupby("Categoría")["Total"].sum().reset_index().sort_values(by="Total", ascending=False)
            fila_inicio = 11
            worksheet_dash.write(fila_inicio, 0, 'CATEGORÍA / RUBRO', formato_encabezado_tabla)
            worksheet_dash.write(fila_inicio, 1, 'MONTO TOTAL', formato_encabezado_tabla)
            
            fila_actual = fila_inicio + 1
            for index, row in resumen.iterrows():
                worksheet_dash.write(fila_actual, 0, row['Categoría'])
                worksheet_dash.write(fila_actual, 1, row['Total'], formato_moneda)
                fila_actual += 1
                
            worksheet_dash.write(fila_actual, 0, 'TOTAL GENERAL', formato_encabezado_tabla)
            worksheet_dash.write(fila_actual, 1, df['Total'].sum(), formato_total)
            
            chart_doughnut = workbook.add_chart({'type': 'doughnut'})
            chart_doughnut.add_series({
                'name': 'Distribución Porcentual',
                'categories': ['DASHBOARD', fila_inicio+1, 0, fila_actual-1, 0],
                'values':     ['DASHBOARD', fila_inicio+1, 1, fila_actual-1, 1],
                'data_labels': {'percentage': True, 'leader_lines': True}
            })
            chart_doughnut.set_title({'name': 'Distribución Porcentual de Gastos'})
            chart_doughnut.set_size({'width': 480, 'height': 320})
            worksheet_dash.insert_chart('D11', chart_doughnut)

            chart_bar = workbook.add_chart({'type': 'bar'})
            chart_bar.add_series({
                'name': 'Monto Total S/',
                'categories': ['DASHBOARD', fila_inicio+1, 0, fila_actual-1, 0],
                'values':     ['DASHBOARD', fila_inicio+1, 1, fila_actual-1, 1],
                'fill':       {'color': '#1E3A8A'}
            })
            chart_bar.set_title({'name': 'Ranking de Egresos'})
            chart_bar.set_size({'width': 480, 'height': 320})
            chart_bar.set_legend({'position': 'none'})
            worksheet_dash.insert_chart('D28', chart_bar)

        if not df.empty:
            for cat in df['Categoría'].unique():
                df_cat = df[df['Categoría'] == cat].drop(columns=['ID'])
                nombre_hoja = str(cat)[:31].replace(':', '').replace('/', '').replace('?', '').replace('*', '').replace('[', '').replace(']', '')
                
                worksheet_cat = workbook.add_worksheet(nombre_hoja)
                dibujar_encabezado_oficial(worksheet_cat)
                
                worksheet_cat.set_column('A:A', 15)
                worksheet_cat.set_column('B:B', 30)
                worksheet_cat.set_column('C:C', 18)
                worksheet_cat.set_column('D:D', 45)
                worksheet_cat.set_column('E:F', 12)
                worksheet_cat.set_column('G:G', 15, formato_moneda)
                worksheet_cat.set_column('H:H', 18, formato_total)
                
                fila_tabla = 11
                for col_num, col_name in enumerate(df_cat.columns):
                    worksheet_cat.write(fila_tabla, col_num, col_name, formato_encabezado_tabla)
                
                fila_datos = fila_tabla + 1
                for _, row_data in df_cat.iterrows():
                    worksheet_cat.write(fila_datos, 0, str(row_data['Fecha']))
                    worksheet_cat.write(fila_datos, 1, row_data['Categoría'])
                    worksheet_cat.write(fila_datos, 2, row_data['N° Serie'])
                    worksheet_cat.write(fila_datos, 3, row_data['Descripción'])
                    worksheet_cat.write(fila_datos, 4, row_data['Cantidad'])
                    worksheet_cat.write(fila_datos, 5, row_data['Unidad'])
                    worksheet_cat.write(fila_datos, 6, row_data['Precio Unitario'], formato_moneda)
                    worksheet_cat.write(fila_datos, 7, row_data['Total'], formato_total)
                    fila_datos += 1
                    
                worksheet_cat.write(fila_datos, 6, "TOTAL RUBRO", formato_encabezado_tabla)
                worksheet_cat.write(fila_datos, 7, df_cat['Total'].sum(), formato_total)

    return output.getvalue()

# ==========================================
# INICIO DE LA APLICACIÓN
# ==========================================
df_gastos = cargar_datos()

# --- BARRA LATERAL (Simplificada para no estorbar en móviles) ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2933/2933923.png", width=60)
    st.title("Menú Principal")
    
    if st.button("🔄 Sincronizar / Refrescar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    
    with st.expander("📝 Datos del Reporte Excel", expanded=False):
        input_periodo = st.text_input("Mes/Año", "JUNIO - 2025")
        input_pres = st.text_input("Presidente", "RUDISON CARRASCO SALAZAR")
        input_tes = st.text_input("Tesorero", "CLEVER ALATA VELASQUEZ")
        input_fisc = st.text_input("Fiscal", "HITLER ESPINOZA LOPEZ")

# ==========================================
# PESTAÑAS PRINCIPALES (Adaptables a móviles)
# ==========================================
tab_dashboard, tab_registro, tab_base_datos = st.tabs([
    "📊 Resumen", 
    "🛒 Registrar Boleta", 
    "🗄️ Base de Datos"
])

# ---------------------------------------------------------
# PESTAÑA 1: DASHBOARD
# ---------------------------------------------------------
with tab_dashboard:
    st.header("📊 Panel de Control")
    
    # KPIs Responsivos (En móvil se apilarán automáticamente)
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    total_gastado = df_gastos["Total"].sum() if not df_gastos.empty else 0.0
    total_registros = len(df_gastos)
    promedio_gasto = df_gastos["Total"].mean() if total_registros > 0 else 0.0
    categoria_mayor = df_gastos.groupby("Categoría")["Total"].sum().idxmax() if not df_gastos.empty and total_gastado > 0 else "N/A"

    kpi1.metric("💰 Gasto Total", f"S/ {total_gastado:,.2f}")
    kpi2.metric("🧾 Operaciones", total_registros)
    kpi3.metric("📈 Promedio", f"S/ {promedio_gasto:,.2f}")
    kpi4.metric("🔥 Mayor Rubro", categoria_mayor)
    st.markdown("---")

    if not df_gastos.empty and total_gastado > 0:
        resumen_cat = df_gastos.groupby("Categoría")["Total"].sum().reset_index().sort_values(by="Total", ascending=False)
        col_barras, col_pastel = st.columns([1.5, 1]) # Proporción ajustada
        
        with col_barras:
            fig_bar = px.bar(resumen_cat, x="Total", y="Categoría", orientation='h', text="Total", color="Categoría", color_discrete_sequence=px.colors.qualitative.Bold)
            fig_bar.update_traces(texttemplate='S/ %{text:,.2f}', textposition='outside')
            fig_bar.update_layout(showlegend=False, xaxis_title="Monto (S/)", yaxis_title="", margin=dict(l=0, r=0, t=30, b=0), height=350)
            st.plotly_chart(fig_bar, use_container_width=True)

        with col_pastel:
            fig_pie = px.pie(resumen_cat, values='Total', names='Categoría', hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
            fig_pie.update_traces(textposition='inside', textinfo='percent+label')
            fig_pie.update_layout(showlegend=False, margin=dict(l=0, r=0, t=30, b=0), height=350)
            st.plotly_chart(fig_pie, use_container_width=True)
    else:
        st.info("Sin datos suficientes para gráficos.")


# ---------------------------------------------------------
# PESTAÑA 2: NUEVO REGISTRO MÚLTIPLE (Optimizado Móvil/PC)
# ---------------------------------------------------------
with tab_registro:
    st.header("📝 Ingreso Rápido de Comprobante")
    st.caption("Llena la Fecha y Serie arriba, luego añade todos los productos abajo.")
    
    with st.container():
        c_fecha, c_serie = st.columns(2)
        fecha_ingreso = c_fecha.date_input("🗓️ Fecha Boleta/Factura", date.today())
        serie_ingreso = c_serie.text_input("🧾 N° Documento", placeholder="Ej: F001-00123")
    
    st.markdown("---")
    
    # Formulario rediseñado en bloques de 2 (Perfecto para móvil)
    with st.form("form_item", clear_on_submit=True):
        st.subheader("🛒 Datos del Producto")
        categoria = st.selectbox("Categoría / Rubro", CATEGORIAS)
        descripcion = st.text_input("Descripción del Producto*", placeholder="Ej: Sacos de Arroz")
        
        # Bloque 1
        colA, colB = st.columns(2)
        cantidad = colA.number_input("Cantidad", min_value=0.01, step=1.0, value=1.0)
        unidad = colB.text_input("Unidad", "UND")
        
        # Bloque 2
        colC, colD = st.columns(2)
        precio_unitario = colC.number_input("P. Unitario", min_value=0.0, step=1.0, value=0.0)
        monto_total = colD.number_input("Total (Dejar 0 para Autocalcular)", min_value=0.0, step=1.0, value=0.0)
        
        st.write("") # Espaciador
        btn_agregar = st.form_submit_button("➕ Añadir a la Lista Temporal", type="secondary", use_container_width=True)
        
        if btn_agregar:
            if descripcion.strip() != "":
                total_final = monto_total if monto_total > 0 else round((cantidad * precio_unitario), 2)
                st.session_state.lista_productos.append({
                    "Categoría": categoria,
                    "Descripción": descripcion.upper(),
                    "Cantidad": cantidad,
                    "Unidad": unidad.upper(),
                    "Precio Unitario": precio_unitario,
                    "Total": total_final
                })
                st.success(f"✔️ {descripcion.upper()} añadido.")
            else:
                st.error("⚠️ Falta la descripción.")
    
    # MOSTRAR CARRITO
    if len(st.session_state.lista_productos) > 0:
        st.markdown("### 📋 Resumen de la Boleta")
        df_lista = pd.DataFrame(st.session_state.lista_productos)
        
        st.dataframe(
            df_lista, 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "Precio Unitario": st.column_config.NumberColumn("P. Unit", format="S/ %.2f"),
                "Total": st.column_config.NumberColumn("Total", format="S/ %.2f")
            }
        )
        
        suma_comprobante = df_lista["Total"].sum()
        st.info(f"**💰 Total del Comprobante: S/ {suma_comprobante:,.2f}**")
        
        # Botones de Acción
        col_guardar, col_limpiar = st.columns([2, 1])
        
        if col_guardar.button("💾 GUARDAR BOLETA EN LA NUBE", type="primary", use_container_width=True):
            ids_existentes = pd.to_numeric(df_gastos["ID"], errors='coerce').dropna()
            nuevo_id = int(ids_existentes.max() + 1) if not ids_existentes.empty else 1
            
            nuevos_registros = []
            for item in st.session_state.lista_productos:
                nuevos_registros.append({
                    "ID": nuevo_id,
                    "Fecha": fecha_ingreso,
                    "Categoría": item["Categoría"],
                    "N° Serie": serie_ingreso if serie_ingreso.strip() else "-",
                    "Descripción": item["Descripción"],
                    "Cantidad": item["Cantidad"],
                    "Unidad": item["Unidad"],
                    "Precio Unitario": item["Precio Unitario"],
                    "Total": item["Total"]
                })
                nuevo_id += 1 
            
            df_nuevos = pd.DataFrame(nuevos_registros)
            df_gastos = pd.concat([df_gastos, df_nuevos], ignore_index=True)
            df_gastos['Fecha'] = pd.to_datetime(df_gastos['Fecha'], errors='coerce').dt.date
            df_gastos = df_gastos.sort_values(by="Fecha", ascending=False).reset_index(drop=True)
            
            with st.spinner('Sincronizando...'):
                exito = guardar_datos(df_gastos)
                
            if exito:
                st.session_state.lista_productos = [] 
                st.success(f"✅ Registros guardados.")
                st.rerun()
                
        if col_limpiar.button("🗑️ Cancelar", use_container_width=True):
            st.session_state.lista_productos = []
            st.rerun()


# ---------------------------------------------------------
# PESTAÑA 3: EDITOR Y DESCARGAS (Anti-Lag)
# ---------------------------------------------------------
with tab_base_datos:
    st.header("🗄️ Base de Datos")
    
    if not df_gastos.empty:
        # 1. Buscador Rápido (No congela la app)
        st.subheader("🔍 Buscador Rápido")
        busqueda = st.text_input("🔎 Escribe para buscar (Descripción, Categoría, Fecha...)", placeholder="Ej: Arroz, Combustible, 2025...")
        
        if busqueda:
            # Filtrar dataframe como texto
            df_vista = df_gastos[df_gastos.astype(str).apply(lambda x: x.str.contains(busqueda, case=False)).any(axis=1)]
        else:
            df_vista = df_gastos
            
        st.dataframe(
            df_vista, 
            use_container_width=True, 
            hide_index=True,
            height=300,
            column_config={
                "Fecha": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY"),
                "Precio Unitario": st.column_config.NumberColumn("P. Unit", format="S/ %.2f"),
                "Total": st.column_config.NumberColumn("Total", format="S/ %.2f")
            }
        )
        
        # 2. Modo Edición Oculto (Para evitar Lag en celulares con +1000 items)
        with st.expander("✏️ Habilitar Edición Global (Precaución)"):
            st.warning("⚠️ Edita directamente las celdas de la tabla inferior y presiona Guardar.")
            edited_df = st.data_editor(
                df_gastos,
                use_container_width=True,
                num_rows="dynamic",
                hide_index=False,
                key="editor_datos_nube",
                height=400,
                column_config={
                    "ID": st.column_config.NumberColumn("ID", disabled=True),
                    "Fecha": st.column_config.DateColumn("Fecha", format="YYYY-MM-DD", required=True),
                    "Categoría": st.column_config.SelectboxColumn("Categoría", options=CATEGORIAS, required=True),
                    "Precio Unitario": st.column_config.NumberColumn("P. Unitario", format="S/ %.2f"),
                    "Total": st.column_config.NumberColumn("Total", format="S/ %.2f", required=True)
                }
            )

            if st.button("💾 Confirmar y Guardar Ediciones", type="primary", use_container_width=True):
                edited_df['Fecha'] = pd.to_datetime(edited_df['Fecha'], errors='coerce').dt.date
                edited_df = edited_df.sort_values(by="Fecha", ascending=False).reset_index(drop=True)
                
                with st.spinner('Sobreescribiendo Base de Datos...'):
                    exito = guardar_datos(edited_df)
                    
                if exito:
                    st.success("✅ ¡Modificaciones guardadas!")
                    st.rerun()

        st.markdown("---")
        # 3. Descarga de Excel
        st.subheader("📥 Exportar")
        excel_data = generar_excel_dinamico(df_gastos, input_periodo, input_pres, input_tes, input_fisc)
        
        st.download_button(
            label="📊 Descargar Excel Oficial",
            data=excel_data,
            file_name=f'Rendicion_{input_periodo.replace(" ", "_")}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            type="primary",
            use_container_width=True
        )
    else:
        st.warning("No hay registros actualmente en Google Sheets.")
