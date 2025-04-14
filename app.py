import pandas as pd
import streamlit as st
import plotly.express as px
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

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

    resumen_mensual = df_total.groupby('Mes')['Ingresos estimados (USD)'].sum().reset_index()
    resumen_mensual.columns = ['Mes/Año', 'Ingresos (USD)']

    st.subheader("🗓️ Ingresos por mes (todos los canales)")
    st.dataframe(resumen_mensual)

    # 📈 Gráfico de línea
    st.subheader("📈 Evolución de ingresos mensuales")
    fig = px.line(resumen_mensual, x='Mes/Año', y='Ingresos (USD)', markers=True)
    fig.update_layout(xaxis_title="Mes", yaxis_title="Ingresos (USD)", title="Ingresos por mes")
    st.plotly_chart(fig, use_container_width=True)

    # Crear libro Excel
    wb = Workbook()

    # Función para aplicar formato y agregar total
    def agregar_hoja(df, titulo, ws=None, fecha_col=None):
        if not ws:
            ws = wb.create_sheet(title=titulo)
        else:
            ws.title = titulo

        total_col_idx = df.shape[1]

        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)

                if r_idx == 1:
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal='center')
                elif fecha_col and c_idx == fecha_col:
                    cell.number_format = "mmm-yyyy"
                elif c_idx == total_col_idx:
                    cell.number_format = '#,##0.00'

        # Agregar total al final
        total_row = ws.max_row + 1
        ws.cell(row=total_row, column=total_col_idx - 1, value="TOTAL").font = Font(bold=True)
        suma_formula = f"=SUM({get_column_letter(total_col_idx)}2:{get_column_letter(total_col_idx)}{total_row - 1})"
        ws.cell(row=total_row, column=total_col_idx, value=suma_formula).font = Font(bold=True)

        # Autoajuste columnas
        for col in ws.columns:
            max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max_length + 2

    # Hoja 1: Resumen mensual total
    ws1 = wb.active
    agregar_hoja(resumen_mensual, "Resumen mensual", ws=ws1, fecha_col=1)

    # Hoja 2: Resumen anual
    df_total['Año'] = df_total['Fecha'].dt.year
    resumen_anual = df_total.groupby('Año')['Ingresos estimados (USD)'].sum().reset_index()
    resumen_anual.columns = ['Año', 'Ingresos (USD)']
    agregar_hoja(resumen_anual, "Resumen anual")

    # Hoja 3: Por canal
    resumen_canal = df_total.groupby(['Canal', 'Mes'])['Ingresos estimados (USD)'].sum().reset_index()
    resumen_canal.columns = ['Canal', 'Mes/Año', 'Ingresos (USD)']
    agregar_hoja(resumen_canal, "Por canal", fecha_col=2)

    # Exportar a BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="⬇️ Descargar resumen completo en Excel",
        data=output,
        file_name="resumen_ingresos_completo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
