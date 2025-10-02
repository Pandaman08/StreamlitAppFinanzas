import streamlit as st

# ================= CONFIGURACIÓN DE PÁGINA =================
st.set_page_config(
    page_title="Finanzas Corporativas",
    page_icon="💹",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ================= HEADER =================
st.markdown(
    """
    <h1 style='text-align: center; color: #1F3B6C; font-weight: bold;'>
        💹 Panel de Finanzas Corporativas
    </h1>
    <h4 style='text-align: center; color: #2E5C87;'>
        Análisis y consolidación de estados financieros, indicadores y reportes ejecutivos
    </h4>
    """,
    unsafe_allow_html=True
)

# ================= MENSAJE INFORMATIVO =================
st.info(
    """
    Navega por los módulos disponibles en el menú lateral para realizar análisis financieros completos:
    """
)

# ================= LISTA DE MÓDULOS =================
st.markdown(
    """
    <ul style='color: #2E5C87; font-size: 16px; line-height: 1.6;'>
        <li>📊 <b>Consolidador SMV:</b> Importa y procesa estados financieros de la SMV.</li>
        <li>📈 <b>Dashboard Económico:</b> Visualiza indicadores y tendencias económicas clave.</li>
        <li>📑 <b>Reportes:</b> Genera exportaciones y descargas de informes ejecutivos.</li>
    </ul>
    """,
    unsafe_allow_html=True
)

