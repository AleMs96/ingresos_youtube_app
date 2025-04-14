import pandas as pd
import streamlit as st
from io import BytesIO

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

    st.subheader("🗓️ Ingresos por mes (todos los canales)")
    st.dataframe(resumen_mensual)

    # Descargar resumen como Excel
    output = BytesIO()
    resumen_mensual.to_excel(output, index=False)
    st.download_button(
        label="⬇️ Descargar resumen en Excel",
        data=output.getvalue(),
        file_name="resumen_ingresos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
