import streamlit as st
import pandas as pd
import plotly.express as px
from streamlit_gsheets import GSheetsConnection
import os
import io
from datetime import date

# ==========================================
# CONFIGURACIÓN DE LA PÁGINA
# ==========================================
st.set_page_config(
    page_title="Rendición de Cuentas - Minera",
    page_icon="⛏️",
    layout="wide",
    initial_sidebar_state="expanded"
)

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

# ==========================================
# CONEXIÓN EN TIEMPO REAL CON GOOGLE SHEETS
# ==========================================
conn = st.connection("gsheets", type=GSheetsConnection)

# ⚠️ Pon aquí el nombre exacto de la pestaña de tu Google Sheets
NOMBRE_PESTAÑA = "Hoja 1"

def cargar_datos():
    try:
        df = conn.read(ttl=0)
        if df.empty or "ID" not in df.columns:
            return pd.DataFrame(columns=["ID", "Fecha", "Categoría", "N° Serie", "Descripción", "Cantidad", "Unidad", "Precio Unitario", "Total"])
        
        # Convertir a numéricos
        df['Cantidad'] = pd.to_numeric(df['Cantidad'], errors='coerce').fillna(0)
        df['Precio Unitario'] = pd.to_numeric(df['Precio Unitario'], errors='coerce').fillna(0)
        df['Total'] = pd.to_numeric(df['Total'], errors='coerce').fillna(0)
        
        # Convertir a fechas, ORDENAR (más reciente a más antigua) y devolver a texto seguro
        df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
        df = df.sort_values(by="Fecha", ascending=False).reset_index(drop=True)
        df['Fecha'] = df['Fecha'].dt.strftime('%Y-%m-%d')
        
        return df
    except Exception:
        return pd.DataFrame(columns=["ID", "Fecha", "Categoría", "N° Serie", "Descripción", "Cantidad", "Unidad", "Precio Unitario", "Total"])

def guardar_datos(df):
    conn.update(data=df)

# ==========================================
# GENERACIÓN DE EXCEL OFICIAL 365
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
        worksheet_dash.set_column('A:A', 40)
        worksheet_dash.set_column('B:B', 20)
        
        if not df.empty:
            resumen = df.groupby("Categoría")["Total"].sum().reset_index()
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
            
            chart = workbook.add_chart({'type': 'doughnut'})
            chart.add_series({
                'name': 'Distribución',
                'categories': ['DASHBOARD', fila_inicio+1, 0, fila_actual-1, 0],
                'values':     ['DASHBOARD', fila_inicio+1, 1, fila_actual-1, 1],
                'data_labels': {'percentage': True, 'leader_lines': True}
            })
            chart.set_title({'name': 'Distribución de Gastos'})
            chart.set_size({'width': 550, 'height': 350})
            worksheet_dash.insert_chart('D12', chart)

        if not df.empty:
            categorias_unicas = df['Categoría'].unique()
            for cat in categorias_unicas:
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
                    worksheet_cat.write(fila_datos, 0, row_data['Fecha'])
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

# --- BARRA LATERAL ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/2933/2933923.png", width=80)
    st.title("Registro de Egresos")
    
    if st.button("🔄 Actualizar Datos de la Nube", use_container_width=True):
        st.rerun()
    
    with st.expander("📝 CONFIGURACIÓN DEL REPORTE", expanded=True):
        input_periodo = st.text_input("Período a Rendir", "JUNIO - 2025")
        input_pres = st.text_input("Presidente", "RUDISON CARRASCO SALAZAR")
        input_tes = st.text_input("Tesorero", "CLEVER ALATA VELASQUEZ")
        input_fisc = st.text_input("Fiscal", "HITLER ESPINOZA LOPEZ")

    st.markdown("---")
    
    with st.form("registro_form", clear_on_submit=True):
        st.subheader("Nuevo Comprobante")
        fecha = st.date_input("Fecha", date.today())
        categoria = st.selectbox("Categoría / Rubro", CATEGORIAS)
        serie = st.text_input("N° Documento / Serie", placeholder="Ej: F001-00123")
        descripcion = st.text_input("Descripción del Producto*", placeholder="Ej: Sacos de Arroz")
        
        col1, col2 = st.columns(2)
        with col1:
            cantidad = st.number_input("Cantidad", min_value=0.01, step=1.0, value=1.0)
        with col2:
            unidad = st.text_input("Unidad", placeholder="Ej: UND, GLN, KG")
            
        col3, col4 = st.columns(2)
        with col3:
            precio_unitario = st.number_input("P. Unitario (Opcional)", min_value=0.0, step=1.0, value=0.0)
        with col4:
            monto_total = st.number_input("Monto Total (S/)*", min_value=0.0, step=1.0, value=0.0)
        
        submit_button = st.form_submit_button("Guardar Registro", type="primary", use_container_width=True)
        
        if submit_button:
            if descripcion and monto_total > 0:
                nuevo_id = int(df_gastos["ID"].max() + 1) if not df_gastos.empty and pd.notnull(df_gastos["ID"].max()) else 1
                
                nuevo_registro = pd.DataFrame([{
                    "ID": nuevo_id,
                    "Fecha": fecha.strftime('%Y-%m-%d'),
                    "Categoría": categoria,
                    "N° Serie": serie if serie else "-",
                    "Descripción": descripcion.upper(),
                    "Cantidad": cantidad,
                    "Unidad": unidad.upper() if unidad else "UND",
                    "Precio Unitario": precio_unitario,
                    "Total": monto_total 
                }])
                
                df_gastos = pd.concat([df_gastos, nuevo_registro], ignore_index=True)
                
                # Re-ordenar automáticamente al agregar un dato nuevo
                df_gastos['Fecha'] = pd.to_datetime(df_gastos['Fecha'], errors='coerce')
                df_gastos = df_gastos.sort_values(by="Fecha", ascending=False).reset_index(drop=True)
                df_gastos['Fecha'] = df_gastos['Fecha'].dt.strftime('%Y-%m-%d')
                
                guardar_datos(df_gastos)
                st.success("✅ Registro guardado en la Nube y ordenado exitosamente!")
                st.rerun()
            else:
                st.error("⚠️ Ingrese una descripción y un Monto Total válido.")

# --- ÁREA PRINCIPAL DASHBOARD ---
st.title("📊 Panel de Control - Sociedad Minera Rey")
st.markdown("Visualización analítica y profesional de rendición de cuentas.")

# Tarjetas KPI 
kpi1, kpi2, kpi3, kpi4 = st.columns(4)
total_gastado = df_gastos["Total"].sum() if not df_gastos.empty else 0.0
total_registros = len(df_gastos)
promedio_gasto = df_gastos["Total"].mean() if total_registros > 0 else 0.0
categoria_mayor = df_gastos.groupby("Categoría")["Total"].sum().idxmax() if not df_gastos.empty and total_gastado > 0 else "N/A"

kpi1.metric("💰 Egreso Total General", f"S/ {total_gastado:,.2f}")
kpi2.metric("🧾 Total de Operaciones", total_registros)
kpi3.metric("📈 Promedio por Compra", f"S/ {promedio_gasto:,.2f}")
kpi4.metric("🔥 Mayor Rubro de Gasto", categoria_mayor)

st.markdown("---")

# Gráficos Interactivos
if not df_gastos.empty and total_gastado > 0:
    resumen_cat = df_gastos.groupby("Categoría")["Total"].sum().reset_index()
    resumen_cat_ordenado = resumen_cat.sort_values(by="Total", ascending=False)
    
    col_barras, col_pastel = st.columns([1.2, 1])
    
    with col_barras:
        st.subheader("Ranking de Gastos por Categoría")
        fig_bar = px.bar(
            resumen_cat_ordenado, x="Total", y="Categoría", orientation='h',
            text="Total", color="Categoría",
            color_discrete_sequence=px.colors.qualitative.Bold
        )
        fig_bar.update_traces(texttemplate='S/ %{text:,.2f}', textposition='outside')
        fig_bar.update_layout(showlegend=False, xaxis_title="Monto (S/)", yaxis_title="", margin=dict(l=0, r=0, t=30, b=0))
        st.plotly_chart(fig_bar, use_container_width=True)

    with col_pastel:
        st.subheader("Distribución Porcentual")
        fig_pie = px.pie(
            resumen_cat, values='Total', names='Categoría', hole=0.4, 
            color_discrete_sequence=px.colors.qualitative.Pastel
        )
        fig_pie.update_traces(textposition='inside', textinfo='percent+label')
        fig_pie.update_layout(showlegend=False, margin=dict(l=0, r=0, t=30, b=0))
        st.plotly_chart(fig_pie, use_container_width=True)
else:
    st.info("Agrega datos para visualizar los gráficos interactivos.")

st.markdown("---")

# ==========================================
# SECCIÓN DE EDICIÓN Y ELIMINACIÓN
# ==========================================
st.subheader("🗄️ Base de Datos Interactiva (Ordenada por Fecha)")
st.info("💡 Haz doble clic en cualquier celda para corregirla. Para eliminar, selecciona la fila a la izquierda y presiona el ícono de papelera.")

if not df_gastos.empty:
    edited_df = st.data_editor(
        df_gastos,
        use_container_width=True,
        num_rows="dynamic",
        hide_index=False,
        key="editor_datos_nube",
        column_config={
            "ID": None,
            "Fecha": st.column_config.DateColumn("Fecha", format="YYYY-MM-DD", required=True),
            "Categoría": st.column_config.SelectboxColumn("Categoría", options=CATEGORIAS, required=True),
            "Precio Unitario": st.column_config.NumberColumn("P. Unitario", format="S/ %.2f"),
            "Total": st.column_config.NumberColumn("Total", format="S/ %.2f", required=True)
        }
    )

    if not df_gastos.equals(edited_df):
        # Si editan la tabla, aseguramos que se vuelva a ordenar por fecha antes de guardar
        edited_df['Fecha'] = pd.to_datetime(edited_df['Fecha'], errors='coerce')
        edited_df = edited_df.sort_values(by="Fecha", ascending=False).reset_index(drop=True)
        edited_df['Fecha'] = edited_df['Fecha'].dt.strftime('%Y-%m-%d')
        
        guardar_datos(edited_df)
        st.success("✅ Cambios sincronizados con la nube!")
        st.rerun()

    st.markdown("---")
    excel_data = generar_excel_dinamico(df_gastos, input_periodo, input_pres, input_tes, input_fisc)
    
    st.download_button(
        label="📊 Descargar Reporte Oficial en Excel 365",
        data=excel_data,
        file_name=f'Rendicion_{input_periodo.replace(" ", "_")}.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        type="primary"
    )
else:
    st.warning("No hay registros actualmente en Google Sheets.")