import pandas as pd
import streamlit as st
import plotly.express as px
from io import BytesIO
import openpyxl
from openpyxl.styles import Alignment, Border, Side

# Título de la app
st.title("📊 Resumen de ingresos mensuales de YouTube")

# Subir archivos .csv o .xlsx
uploaded_files = st.file_uploader(
    "Sube tus archivos (.csv o .xlsx) de varios canales",
    type=["csv", "xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    dataframes = []
    
    # Leer y procesar archivos
    for uploaded_file in uploaded_files:
        filename = uploaded_file.name
        if filename.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        elif filename.endswith('.xlsx'):
            df = pd.read_excel(uploaded_file)

        df['Canal'] = filename.rsplit('.', 1)[0]  # Añadir nombre del canal
        df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')  # Convertir a fecha
        df = df.dropna(subset=['Fecha', 'Ingresos estimados (USD)'])  # Eliminar filas con valores nulos
        dataframes.append(df)

    # Concatenar todos los dataframes en uno solo
    df_total = pd.concat(dataframes)

    # Crear columna de mes
    df_total['Mes'] = df_total['Fecha'].dt.to_period('M').dt.to_timestamp()

    # Resumen mensual (todos los canales)
    resumen_mensual = df_total.groupby('Mes')['Ingresos estimados (USD)'].sum().reset_index()

    # Resumen por canal
    resumen_canal = df_total.groupby(['Canal', 'Mes'])['Ingresos estimados (USD)'].sum().reset_index()

    # Resumen anual (por año)
    df_total['Año'] = df_total['Fecha'].dt.year
    resumen_anual = df_total.groupby('Año')['Ingresos estimados (USD)'].sum().reset_index()

    # Mostrar resúmenes en Streamlit
    st.subheader("🗓️ Ingresos por mes (todos los canales)")
    st.dataframe(resumen_mensual)

    st.subheader("📈 Ingresos por canal (por mes)")
    st.dataframe(resumen_canal)

    st.subheader("📅 Ingresos anuales")
    st.dataframe(resumen_anual)

    # Crear gráfico de ingresos por mes (todos los canales)
    fig = px.line(resumen_mensual, x='Mes', y='Ingresos estimados (USD)', 
                  title='Ingresos mensuales totales por todos los canales')
    st.plotly_chart(fig)

    # Crear gráfico de ingresos por canal (mensual)
    fig_canal = px.line(resumen_canal, x='Mes', y='Ingresos estimados (USD)', color='Canal',
                        title='Ingresos mensuales por canal')
    st.plotly_chart(fig_canal)

    # Crear gráfico de ingresos anuales
    fig_anual = px.bar(resumen_anual, x='Año', y='Ingresos estimados (USD)', 
                       title='Ingresos anuales totales por todos los canales')
    st.plotly_chart(fig_anual)

    # Descargar resumen como Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        resumen_mensual.to_excel(writer, index=False, sheet_name="Resumen Mensual")
        resumen_canal.to_excel(writer, index=False, sheet_name="Resumen por Canal")
        resumen_anual.to_excel(writer, index=False, sheet_name="Resumen Anual")

        # Obtener el workbook
        workbook = writer.book
        worksheet1 = workbook["Resumen Mensual"]
        worksheet2 = workbook["Resumen por Canal"]
        worksheet3 = workbook["Resumen Anual"]

        # Formato de las celdas en las hojas
        for worksheet in [worksheet1, worksheet2, worksheet3]:
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter  # Obtener la letra de la columna
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column].width = adjusted_width

            # Formato de celdas (alineación, borde, formato de número)
            for row in worksheet.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = Border(
                        left=Side(border_style="thin"),
                        right=Side(border_style="thin"),
                        top=Side(border_style="thin"),
                        bottom=Side(border_style="thin")
                    )

            # Formato de número de ingresos como dinero (USD)
            if worksheet == worksheet1 or worksheet == worksheet2 or worksheet == worksheet3:
                for row in worksheet.iter_rows(min_row=2, min_col=2, max_col=2):
                    for cell in row:
                        cell.number_format = '"$"#,##0.00'

    output.seek(0)

    st.download_button(
        label="⬇️ Descargar resumen en Excel",
        data=output.getvalue(),
        file_name="resumen_ingresos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

