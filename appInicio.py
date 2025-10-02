import streamlit as st

# --- CONFIGURACIÓN DE LA APP ---
st.set_page_config(
    page_title="Dashboard Económico",
    page_icon="💹",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ENCABEZADO ---
st.markdown(
    """
    <style>
    .main-title {
        font-size:40px !important;
        color:#2E5382;
        text-align:center;
        font-weight: bold;
    }
    .subtitle {
        font-size:20px !important;
        color:#555;
        text-align:center;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.markdown("<p class='main-title'>📊 Dashboard Económico</p>", unsafe_allow_html=True)
st.markdown("<p class='subtitle'>Indicadores financieros y reportes interactivos</p>", unsafe_allow_html=True)

# --- MENÚ LATERAL ---
menu = ["🏠 Inicio", "📈 Indicadores", "📑 Reportes"]
choice = st.sidebar.radio("Navegación", menu)

# --- CONTENIDO ---
if choice == "🏠 Inicio":
    st.subheader("🏠 Página Principal")
    st.info("Bienvenido al panel de control económico.")
    st.write("Aquí puedes mostrar gráficos, tablas y KPIs relevantes de economía.")

    col1, col2, col3 = st.columns(3)
    col1.metric("PIB Anual", "3.2%", "▲ 0.4%")
    col2.metric("Inflación", "6.1%", "▼ 0.2%")
    col3.metric("Tipo de Cambio", "S/ 3.78", "▲ 0.05")

elif choice == "📈 Indicadores":
    st.subheader("📈 Indicadores Financieros")
    st.success("Aquí puedes cargar gráficos de series de tiempo, comparaciones y proyecciones.")
    st.line_chart({"PIB": [2.3, 2.8, 3.1, 3.5, 3.2], "Inflación": [5.8, 6.2, 6.5, 6.3, 6.1]})

elif choice == "📑 Reportes":
    st.subheader("📑 Reportes Económicos")
    st.warning("Aquí puedes mostrar PDFs, tablas interactivas o descargar reportes en Excel.")

