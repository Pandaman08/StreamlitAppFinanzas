import streamlit as st
from styles import apply_custom_styles
from processor import procesar_archivos
from analyzer import calcular_analisis_vh, calcular_ratios
from exporter import exportar_a_excel
import plotly.graph_objects as go
import pandas as pd

# ================= CONFIGURACIÓN INICIAL =================
st.set_page_config(
    page_title="Consolidador SMV - Finanzas Corporativas",
    layout="wide",
    page_icon="📊",
)

apply_custom_styles()

# ================= HEADER CON LOGO =================
col_logo, col_title = st.columns([1, 5])
with col_logo:
    st.image("assets/estado-financiero.png", width=100)
with col_title:
    st.title("📊 Consolidador de Estados Financieros - SMV")
    st.markdown("**Análisis Financiero Automatizado** | Sube archivos Excel del SMV (2002-2024) y obtén análisis completo con gráficas.")

# ================= SIDEBAR =================
with st.sidebar:
    st.header("⚙️ Configuración")
    nombre_empresa = st.text_input("Nombre de la Empresa", value="EMPRESA ANALIZADA", help="Aparecerá en el reporte")
    st.markdown("---")
    st.markdown("### 📋 Instrucciones")
    st.info("""
    1. Descarga archivos Excel (.xls) del SMV
    2. Súbelos (pueden ser de cualquier año: 2002-2024)
    3. Espera el procesamiento
    4. Revisa resultados y descarga el consolidado
    """)

# ================= UPLOAD FILES =================
archivos = st.file_uploader(
    "📁 Selecciona archivos Excel (.xls) del SMV",
    type=["xls"],
    accept_multiple_files=True
)
if not archivos:
    st.warning("👆 **Por favor, sube los archivos Excel del SMV para comenzar el análisis.**")
    st.stop()
elif len(archivos) < 5:
    st.error(f"❌ **Se requieren al menos 5 archivos. Has subido solo {len(archivos)}.**")
    st.info("💡 **Consejo**: Mantén presionada la tecla **Ctrl** (Windows) o **Cmd** (Mac) mientras haces clic para seleccionar varios archivos a la vez.")
    st.stop()

# ================= PROCESAR ARCHIVOS =================
with st.spinner("📦 Procesando archivos..."):
    df_balance, df_resultados, df_flujo_efectivo = procesar_archivos(archivos)

# ================= ANÁLISIS VERTICAL Y HORIZONTAL =================
with st.spinner("📈 Calculando análisis vertical y horizontal..."):
    df_vertical_balance, df_horizontal_balance, df_vertical_resultados, df_horizontal_resultados = calcular_analisis_vh(df_balance, df_resultados)

# ================= CÁLCULO DE RATIOS =================
with st.spinner("🧮 Calculando ratios financieros..."):
    df_ratios, debug_info, anios_comunes = calcular_ratios(df_balance, df_resultados)

# ================= SIDEBAR STATUS =================
with st.sidebar:
    st.markdown("---")
    st.success(f"✅ **{len(archivos)}** archivos procesados")
    if anios_comunes:
        st.info(f"📅 **Años:** {', '.join(map(str, anios_comunes))}")
    st.metric("Ratios Calculados", len(df_ratios) if not df_ratios.empty else 0)

# ================= TABS =================
tab1, tab2, tab3, tab4 = st.tabs(["📊 Estados Financieros", "📈 Análisis V/H", "🧮 Ratios y Gráficas", "📥 Descargar"])

with tab1:
    st.subheader("💼 Estado de Situación Financiera")
    if not df_balance.empty:
        st.dataframe(df_balance, use_container_width=True)
    else:
        st.warning("No se encontró data del Balance")
    st.markdown("---")
    st.subheader("💰 Estado de Resultados")
    if not df_resultados.empty:
        st.dataframe(df_resultados, use_container_width=True)
    else:
        st.warning("No se encontró data del Estado de Resultados")
    st.markdown("---")
    st.subheader("💵 Estado de Flujo de Efectivo")
    if not df_flujo_efectivo.empty:
        st.dataframe(df_flujo_efectivo, use_container_width=True)
    else:
        st.warning("No se encontró data del Flujo de Efectivo")

with tab2:
    st.subheader("📊 Análisis Vertical y Horizontal - Estado de Situación Financiera")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Análisis Vertical (%)**")
        if not df_vertical_balance.empty:
            st.dataframe(df_vertical_balance.fillna("N/A"), use_container_width=True)
    with col2:
        st.markdown("**Análisis Horizontal (Variación %)**")
        if not df_horizontal_balance.empty:
            st.dataframe(df_horizontal_balance.fillna("N/A"), use_container_width=True)
    st.markdown("---")
    st.subheader("📊 Análisis Vertical y Horizontal - Estado de Resultados")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Análisis Vertical (%)**")
        if not df_vertical_resultados.empty:
            st.dataframe(df_vertical_resultados.fillna("N/A"), use_container_width=True)
    with col2:
        st.markdown("**Análisis Horizontal (Variación %)**")
        if not df_horizontal_resultados.empty:
            st.dataframe(df_horizontal_resultados.fillna("N/A"), use_container_width=True)

with tab3:
    st.subheader("🧮 Ratios Financieros")
    if not df_ratios.empty:
        ultimo_anio = df_ratios.columns[-1]
        penultimo_anio = df_ratios.columns[-2] if len(df_ratios.columns) > 1 else ultimo_anio
        def format_pct(val):
            return f"{val:.2%}" if isinstance(val, (int, float)) else "N/A"
        def format_num(val, dec=2):
            return f"{val:.{dec}f}" if isinstance(val, (int, float)) else "N/A"
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            val_actual = df_ratios.loc['ROE', ultimo_anio] if 'ROE' in df_ratios.index else "N/A"
            val_anterior = df_ratios.loc['ROE', penultimo_anio] if 'ROE' in df_ratios.index else "N/A"
            delta = val_actual - val_anterior if isinstance(val_actual,(int,float)) and isinstance(val_anterior,(int,float)) else None
            st.metric("ROE", format_pct(val_actual), delta=(format_pct(delta) if delta is not None else None))
        with col2:
            val_actual = df_ratios.loc['ROA', ultimo_anio] if 'ROA' in df_ratios.index else "N/A"
            val_anterior = df_ratios.loc['ROA', penultimo_anio] if 'ROA' in df_ratios.index else "N/A"
            delta = val_actual - val_anterior if isinstance(val_actual,(int,float)) and isinstance(val_anterior,(int,float)) else None
            st.metric("ROA", format_pct(val_actual), delta=(format_pct(delta) if delta is not None else None))
        with col3:
            val_actual = df_ratios.loc['Liquidez Corriente', ultimo_anio] if 'Liquidez Corriente' in df_ratios.index else "N/A"
            val_anterior = df_ratios.loc['Liquidez Corriente', penultimo_anio] if 'Liquidez Corriente' in df_ratios.index else "N/A"
            delta = val_actual - val_anterior if isinstance(val_actual,(int,float)) and isinstance(val_anterior,(int,float)) else None
            st.metric("Liquidez Corriente", format_num(val_actual,2), delta=(format_num(delta,2) if delta is not None else None))
        with col4:
            val_actual = df_ratios.loc['Margen Neto', ultimo_anio] if 'Margen Neto' in df_ratios.index else "N/A"
            val_anterior = df_ratios.loc['Margen Neto', penultimo_anio] if 'Margen Neto' in df_ratios.index else "N/A"
            delta = val_actual - val_anterior if isinstance(val_actual,(int,float)) and isinstance(val_anterior,(int,float)) else None
            st.metric("Margen Neto", format_pct(val_actual), delta=(format_pct(delta) if delta is not None else None))
        st.markdown("---")
        st.markdown("### 📋 Tabla de Ratios")
        st.dataframe(df_ratios, use_container_width=True)
        st.markdown("---")
        st.markdown("### 📈 Gráficas Individuales por Ratio")
        col1, col2 = st.columns(2)
        for idx, ratio in enumerate(df_ratios.index):
            fig = go.Figure()
            yvals = pd.to_numeric(df_ratios.loc[ratio], errors='coerce')
            fig.add_trace(go.Scatter(
                x=df_ratios.columns,
                y=yvals,
                mode='lines+markers',
                name=ratio,
                line=dict(width=3),
                marker=dict(size=8)
            ))
            fig.update_layout(
                title=f"{ratio}",
                xaxis_title="Año",
                yaxis_title="Valor",
                height=350,
                showlegend=False
            )
            if idx % 2 == 0:
                with col1:
                    st.plotly_chart(fig, use_container_width=True)
            else:
                with col2:
                    st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("No se pudieron calcular ratios")

with tab4:
    st.subheader("📥 Descargar Reporte Consolidado")
    st.markdown(f"**Empresa:** {nombre_empresa}")
    st.markdown(f"**Años analizados:** {', '.join(map(str, anios_comunes)) if anios_comunes else 'N/A'}")
    with st.spinner("🎨 Generando Excel con estilos y gráficas..."):
        output_excel = exportar_a_excel(
            df_balance, df_resultados, df_flujo_efectivo,
            df_vertical_balance, df_horizontal_balance,
            df_vertical_resultados, df_horizontal_resultados,
            df_ratios, nombre_empresa, anios_comunes
        )
    st.download_button(
        label="📥 Descargar Excel Consolidado (Con Gráficas)",
        data=output_excel.getvalue(),
        file_name=f"Analisis_Financiero_{nombre_empresa.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_excel_con_graficas"
    )
    st.success("✅ ¡Proceso completado! El archivo incluye estados financieros, análisis V/H, ratios y gráficas.")