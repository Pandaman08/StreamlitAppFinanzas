import streamlit as st

st.set_page_config(
    page_title="Finanzas Corporativas",
    page_icon="💹",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("<h1 style='text-align: center; color: #2E5382;'>💹 Panel de Finanzas Corporativas</h1>", unsafe_allow_html=True)

st.info("Usa el menú lateral para navegar entre los módulos:")

st.write("""
- 📊 Consolidador SMV (importa y procesa estados financieros).
- 📈 Dashboard Económico (indicadores y visualizaciones).
- 📑 Reportes (exportaciones y descargas).
""")
