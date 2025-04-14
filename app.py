import pandas as pd
import streamlit as st
from io import BytesIO
import plotly.express as px
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, numbers
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from PIL import Image
import plotly.io as pio

st.set_page_config(page_title="Ingresos YouTube", layout="centered")
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
    df_total['Año'] = df_total['Fecha'].dt.year

    resumen_mensual = df_total.groupby('Mes')['Ingresos estimados (USD)'].sum().reset_index()
    resumen_anual = df_total.groupby('Año')['Ingresos estimados (USD)'].sum().reset_index()
    resumen_canal = df_total.groupby('Canal')['Ingresos estimados (USD)'].sum().reset_index()

    # Gráfico de línea
    fig = px.line(resumen_mensual, x="Mes", y="Ingresos estimados (USD)", title="📈 Evolución mensual de ingresos")
    st.plotly_chart(fig)

    # Mostrar tablas
    st.subheader("🗓️ Ingresos por mes (todos los canales)")
    st.dataframe(resumen_mensual)

    st.subheader("📅 Ingresos por año")
    st.dataframe(resumen_anual)

    st.subheader("📺 Ingresos por canal")
    st.dataframe(resumen_canal)

    # Generar Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        resumen_mensual.to_excel(writer, index=False, sheet_name="Resumen mensual")
        resumen_anual.to_excel(writer, index=False, sheet_name="Resumen anual")
        resumen_canal.to_excel(writer, index=False, sheet_name="Resumen por canal")

        workbook = writer.book

        def estilizar_hoja(sheet_name, formato_fecha=False, col_dinero=2):
            ws = writer.sheets[sheet_name]
            thin_border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
            for row in ws.iter_rows():
                for cell in row:
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    if cell.row == 1:
                        cell.font = Font(bold=True, color="FFFFFF")
                        cell.fill = PatternFill(start_color="4F81BD", fill_type="solid")

            if formato_fecha:
                for cell in ws["A"][1:]:
                    cell.number_format = "mmm-yyyy"

            for row in ws.iter_rows(min_row=2, min_col=col_dinero, max_col=col_dinero):
                for cell in row:
                    cell.number_format = '"$"#,##0.00'

            for i, column_cells in enumerate(ws.iter_cols(1, ws.max_column)):
                max_length = max(len(str(cell.value)) if cell.value else 0 for cell in column_cells)
                col_letter = get_column_letter(i + 1)
                ws.column_dimensions[col_letter].width = max_length + 4

        estilizar_hoja("Resumen mensual", formato_fecha=True)
        estilizar_hoja("Resumen anual")
        estilizar_hoja("Resumen por canal")

        # Insertar gráfico como imagen (nueva hoja)
        img_bytes = fig.to_image(format="png")
        img_stream = BytesIO(img_bytes)
        img = Image.open(img_stream)

        from openpyxl.drawing.image import Image as XLImage
        chart_sheet = workbook.create_sheet("Gráfico")
        img_path = "/tmp/grafico_ingresos.png"
        img.save(img_path)
        xl_img = XLImage(img_path)
        chart_sheet.add_image(xl_img, "B2")

    st.download_button(
        label="⬇️ Descargar resumen en Excel con gráfico",
        data=output.getvalue(),
        file_name="resumen_ingresos_youtube.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
