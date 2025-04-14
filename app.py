import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.io as pio
import openpyxl
from io import BytesIO
from openpyxl.drawing.image import Image
from io import BytesIO

# Título de la app
st.title("📊 Resumen de ingresos mensuales de YouTube")

# Subir archivos
uploaded_files = st.file_uploader(
    "Sube tus archivos (.csv o .xlsx) de varios canales",
    type=["csv", "xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    dataframes = []
    
    for uploaded_file in uploaded_files:
        filename = uploaded_file.name
        if filename.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        elif filename.endswith('.xlsx'):
            df = pd.read_excel(uploaded_file)

        # Añadir el nombre del canal
        df['Canal'] = filename.rsplit('.', 1)[0]
        df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
        df = df.dropna(subset=['Fecha', 'Ingresos estimados (USD)'])
        dataframes.append(df)

    df_total = pd.concat(dataframes)

    # Creación de la columna de mes
    df_total['Mes'] = df_total['Fecha'].dt.to_period('M').dt.to_timestamp()

    # Resumen mensual
    resumen_mensual = df_total.groupby('Mes')['Ingresos estimados (USD)'].sum().reset_index()

    # Crear gráficos
    fig_total = px.line(resumen_mensual, x='Mes', y='Ingresos estimados (USD)', 
                        title="Ingresos mensuales totales (todos los canales)")
    fig_total.update_xaxes(type='category')

    fig_canal = px.line(df_total, x='Mes', y='Ingresos estimados (USD)', color='Canal',
                        title="Ingresos mensuales por canal")
    fig_canal.update_xaxes(type='category')

    fig_anual = px.bar(resumen_mensual, x='Mes', y='Ingresos estimados (USD)', 
                       title="Ingresos anuales totales (todos los canales)")

    # Mostrar los gráficos en Streamlit
    st.plotly_chart(fig_total)
    st.plotly_chart(fig_canal)
    st.plotly_chart(fig_anual)

    # Descargar resumen como Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        resumen_mensual.to_excel(writer, sheet_name='Resumen mensual', index=False)
        df_total.to_excel(writer, sheet_name='Detalles por canal', index=False)

        # Guardar archivo Excel
        writer.save()

    # Convertir gráficos a imagen con Kaleido y agregar a Excel
    img_bytes_total = fig_total.to_image(format="png")
    img_bytes_canal = fig_canal.to_image(format="png")
    img_bytes_anual = fig_anual.to_image(format="png")

    # Escribir la imagen en Excel
    wb = openpyxl.load_workbook(output)
    ws = wb['Resumen mensual']

    # Insertar imagenes en celdas de Excel
    img_total = Image(BytesIO(img_bytes_total))
    img_canal = Image(BytesIO(img_bytes_canal))
    img_anual = Image(BytesIO(img_bytes_anual))

    ws.add_image(img_total, 'F2')
    ws.add_image(img_canal, 'F20')
    ws.add_image(img_anual, 'F40')

    # Ajustar el formato de las columnas
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width

    # Formato de la fecha en "ene-2024"
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1):
        for cell in row:
            cell.number_format = '[$-F400]mmm-yy'

    # Formato de las celdas de ingresos como moneda
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=2, max_col=2):
        for cell in row:
            cell.number_format = '"$"#,##0.00'

    # Generar archivo final
    final_output = BytesIO()
    wb.save(final_output)
    st.download_button(
        label="⬇️ Descargar resumen con gráficos en Excel",
        data=final_output.getvalue(),
        file_name="resumen_ingresos_con_graficos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
