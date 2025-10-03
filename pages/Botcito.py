import streamlit as st
import pandas as pd
import os
import glob
from openai import OpenAI

# ----------------------------
# Configuración inicial
# ----------------------------
st.set_page_config(page_title="FinAI Bot – Financiero Inteligente", page_icon="🤖", layout="wide")
OPENAI_API_KEY = "CLAVE"
client = OpenAI(api_key=OPENAI_API_KEY)

# ----------------------------
# CSS para títulos y resúmenes
# ----------------------------
st.markdown("""
<style>
    .titulo-hoja {
        background: linear-gradient(135deg, #6f42c1 0%, #9d6bff 100%);
        color: white;
        padding: 12px;
        border-radius: 15px;
        font-size: 1.5em;
        font-weight: bold;
        text-align: center;
        margin-top: 15px;
        margin-bottom: 10px;
    }
    .resumen-hoja {
        background: #d0f0fd; /* celeste */
        color: #004080;       /* azul */
        padding:15px;
        border-radius:15px;
        margin:5px;
    }
</style>
""", unsafe_allow_html=True)

# ----------------------------
# Cargar Excel
# ----------------------------
descargas = os.path.expanduser("~/Downloads")
archivos = glob.glob(os.path.join(descargas, "Analisis_Financiero_*.xlsx"))

df_dict = {}
empresa_nombre = "Empresa Analizada"

if archivos:
    archivo_reciente = max(archivos, key=os.path.getctime)
    excel = pd.ExcelFile(archivo_reciente)
    empresa_nombre = os.path.basename(archivo_reciente).split("Analisis_Financiero_")[1].split(".xlsx")[0]
    for sheet in excel.sheet_names:
        df_dict[sheet] = pd.read_excel(excel, sheet_name=sheet)
else:
    archivo = st.file_uploader("📂 Subir archivo Excel", type=["xlsx"])
    if archivo:
        excel = pd.ExcelFile(archivo)
        empresa_nombre = archivo.name.replace(".xlsx", "")
        for sheet in excel.sheet_names:
            df_dict[sheet] = pd.read_excel(excel, sheet_name=sheet)
st.title("🚀 Botcito – Asistente Financiero")
st.markdown(f"### 🏢 {empresa_nombre}", unsafe_allow_html=True)

# ----------------------------
# Inicializar session_state para resúmenes
# ----------------------------
if "resumenes" not in st.session_state:
    st.session_state["resumenes"] = {}


# ----------------------------
# Selector de hojas
# ----------------------------
if df_dict:
    hoja_seleccionada = st.selectbox("Selecciona la hoja que deseas analizar:", list(df_dict.keys()))
    
    if hoja_seleccionada:
        df = df_dict[hoja_seleccionada]

        # Mostrar título difuminado
        st.markdown(f"<div class='titulo-hoja'>{hoja_seleccionada}</div>", unsafe_allow_html=True)

        # Mostrar cuadro de datos
        st.dataframe(df.head(10))

        # Botón para generar análisis si aún no existe
        if hoja_seleccionada not in st.session_state["resumenes"]:
            if st.button(f"Generar análisis para '{hoja_seleccionada}'"):
                with st.spinner("Generando análisis..."):
                    texto = df.head(100).to_string(index=False)
                    prompt = f"""
                    Eres un analista financiero experto. Resume de forma clara y detallada la hoja '{hoja_seleccionada}'.
                    Explica los puntos clave, tendencias, riesgos y oportunidades, de manera entendible para alguien sin conocimientos financieros.
                    Datos de la hoja:
                    {texto}
                    """
                    try:
                        respuesta = client.chat.completions.create(
                            model="gpt-4o-mini",
                            messages=[{"role": "user", "content": prompt}]
                        )
                        resumen = respuesta.choices[0].message.content
                    except Exception as e:
                        resumen = f"Error generando análisis: {e}"

                    st.session_state["resumenes"][hoja_seleccionada] = resumen

        # Mostrar análisis si ya existe
        if hoja_seleccionada in st.session_state["resumenes"]:
            st.markdown(f"<div class='resumen-hoja'>{st.session_state['resumenes'][hoja_seleccionada]}</div>", unsafe_allow_html=True)

# ----------------------------
# Historial de análisis ocultable
# ----------------------------
if st.session_state.get("resumenes"):
    with st.expander("📂 Historial de análisis generados en esta sesión", expanded=False):
        for hoja, resumen in st.session_state["resumenes"].items():
            st.markdown(f"<div class='titulo-hoja'>{hoja}</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='resumen-hoja'>{resumen}</div>", unsafe_allow_html=True)