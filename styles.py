import streamlit as st

def apply_custom_styles():
    """Aplica estilos CSS personalizados a toda la app."""
    st.markdown("""
    <style>
        /* Colores principales */
        :root {
            --primary-color: #2E7D32;       /* Verde oscuro profesional */
            --secondary-color: #4CAF50;     /* Verde medio */
            --accent-color: #81C784;        /* Verde claro */
            --background-color: #f5f9f5;    /* Fondo suave */
            --text-color: #2E7D32;          /* Texto oscuro */
            --border-color: #c8e6c9;        /* Borde suave */
            --card-bg: white;
            --shadow: 0 4px 12px rgba(46, 125, 50, 0.1);
        }

        /* Estilos generales */
        .stApp {
            background-color: var(--background-color);
            color: var(--text-color);
        }

        /* Botones */
        .stButton > button {
            background-color: var(--primary-color);
            color: white;
            border-radius: 8px;
            padding: 10px 20px;
            font-weight: bold;
            transition: all 0.3s ease;
        }
        .stButton > button:hover {
            background-color: var(--secondary-color);
            transform: translateY(-2px);
            box-shadow: var(--shadow);
        }

        /* Tarjetas */
        .stMetric {
            background-color: var(--card-bg);
            border-radius: 12px;
            padding: 15px;
            box-shadow: var(--shadow);
            border-left: 4px solid var(--primary-color);
        }

        /* Tabs */
        .stTabs [data-baseweb="tab-list"] {
            gap: 2px;
        }
        .stTabs [data-baseweb="tab"] {
            height: 40px;
            background-color: var(--card-bg);
            border: 1px solid var(--border-color);
            border-radius: 6px 6px 0 0;
            padding: 0 15px;
            font-weight: bold;
            color: var(--text-color);
        }
        .stTabs [data-baseweb="tab"]:hover {
            background-color: var(--accent-color);
        }
        .stTabs [data-baseweb="tab"][aria-selected="true"] {
            background-color: var(--primary-color);
            color: white;
        }

        /* Encabezados */
        h1, h2, h3 {
            color: var(--primary-color);
            font-weight: 600;
        }

        /* DataFrames */
        .stDataFrame {
            border-radius: 8px;
            overflow: hidden;
            box-shadow: var(--shadow);
        }

        /* Sidebar */
        .css-1d391kg {
            background-color: var(--card-bg);
            border-right: 1px solid var(--border-color);
        }

        /* Progress bar */
        .stProgress > div > div > div {
            background-color: var(--secondary-color);
        }

        /* Alertas */
        .stAlert {
            background-color: var(--accent-color);
            border-left: 4px solid var(--primary-color);
        }

        /* Inputs */
        .stTextInput > div > div > input {
            border: 1px solid var(--border-color);
            border-radius: 6px;
        }
        .stTextInput > div > div > input:focus {
            border-color: var(--primary-color);
            box-shadow: 0 0 0 2px rgba(46, 125, 50, 0.2);
        }

        /* Info boxes */
        .stInfo {
            background-color: var(--accent-color);
            border-left: 4px solid var(--primary-color);
        }

        /* Metricas */
        .stMetric > div > div > div > div {
            font-size: 1.2em;
            font-weight: bold;
        }
    </style>
    """, unsafe_allow_html=True)