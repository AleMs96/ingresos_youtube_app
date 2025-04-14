import pandas as pd
import plotly.express as px
from io import BytesIO
import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.styles import PatternFill, Alignment, Border, Side, NamedStyle
import streamlit as st

# Cargar los archivos subidos
st.title("📊 Resumen de ingresos mensuales de YouTube")

uploaded_files = st.file_uploader(
    "Subí tus archivos (.csv o .xlsx) de varios canales",
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

        df['Canal'] = filename.rsplit('.', 1)[0]
        df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
        df = df.dropna(subset=['Fecha', 'Ingresos estimados (USD)'])
        dataframes.append(df)

    df_total = pd.concat(dataframes)
    df_total['Mes'] = df_total['Fecha'].dt.to_period('M').dt.to_timestamp()

    # Resumen mensual
    resumen_mensual = df_total.groupby('Mes')['Ingresos estimados (USD)'].sum().reset_index()

    # Resumen por canal
    resumen_por_canal = df_total.groupby(['Canal', 'Mes'])['Ingresos estimados (USD)'].sum().reset_index()

    # Crear gráfico de línea
    fig = px.line(resumen_mensual, x='Mes', y='Ingresos estimados (USD)', title="Ingresos por mes")

    # Guardar gráfico como imagen
    img_bytes = fig.to_image(format="png")

    # Crear el archivo Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        resumen_mensual.to_excel(writer, index=False, sheet_name="Resumen Mensual")
        resumen_por_canal.to_excel(writer, index=False, sheet_name="Resumen por Canal")
        
        # Obtener el libro de trabajo
        workbook = writer.book
        worksheet = workbook["Resumen Mensual"]

        # Insertar el gráfico como imagen
        img = Image(BytesIO(img_bytes))
        img.width = 800  # Tamaño de la imagen (ajustar según lo necesites)
        img.height = 600
        worksheet.add_image(img, "E5")  # Colocar la imagen en la celda E5

        # Estilo para el formato de dinero
        money_style = NamedStyle(name="money_style", number_format='"$"#,##0.00')
        if "Ingresos estimados (USD)" in resumen_mensual.columns:
            # Aplicar el formato a la columna de ingresos
            for row in worksheet.iter_rows(min_row=2, min_col=2, max_col=2):  # Columna de "Ingresos estimados (USD)"
                for cell in row:
                    cell.style = money_style

        # Darle formato a las celdas de la tabla
        header_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")  # Color verde claro
        header_alignment = Alignment(horizontal="center", vertical="center")  # Centrar texto
        border = Border(
            top=Side(style='thin'),
            bottom=Side(style='thin'),
            left=Side(style='thin'),
            right=Side(style='thin')
        )

        # Formato para las cabeceras
        for col in worksheet.columns:
            for cell in col:
                cell.border = border
                if cell.row == 1:  # Es la cabecera
                    cell.fill = header_fill
                    cell.alignment = header_alignment

        # Ajustar ancho de las columnas
        column_widths = [max(len(str(cell.value)) for cell in col) for col in worksheet.columns]
        for i, col in enumerate(worksheet.columns):
            worksheet.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width = column_widths[i] + 2  # Ajustar el ancho

        # Formatear la fecha en "Mes-Año"
        for row in worksheet.iter_rows(min_row=2, min_col=1, max_col=1):  # Columna de "Mes"
            for cell in row:
                cell.number_format = '[$-C0A]mmm-aaaa'  # Formato "ene-2024"

        # Resumen por canal
        worksheet_canal = workbook["Resumen por Canal"]

        # Aplicar formato similar para "Resumen por Canal"
        for col in worksheet_canal.columns:
            for cell in col:
                cell.border = border
                if cell.row == 1:  # Es la cabecera
                    cell.fill = header_fill
                    cell.alignment = header_alignment

        # Ajustar ancho de las columnas en "Resumen por Canal"
        column_widths_canal = [max(len(str(cell.value)) for cell in col) for col in worksheet_canal.columns]
        for i, col in enumerate(worksheet_canal.columns):
            worksheet_canal.column_dimensions[openpyxl.utils.get_column_letter(i+1)].width = column_widths_canal[i] + 2  # Ajustar el ancho

        # Formatear la fecha en "Mes-Año" para "Resumen por Canal
