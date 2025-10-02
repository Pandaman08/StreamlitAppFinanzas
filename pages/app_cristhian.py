import numpy as np
import streamlit as st
import pandas as pd
import io
from bs4 import BeautifulSoup
import re
import unicodedata

# ------------------- CONFIG -------------------
st.set_page_config(page_title="Consolidador SMV - Finanzas Corporativas", layout="wide")
st.title("📊 Consolidador de Estados Financieros - SMV")
st.markdown("Sube los archivos Excel descargados de la Superintendencia del Mercado de Valores (SMV).")

# ------------------- UPLOADER -------------------
archivos = st.file_uploader(
    "Selecciona archivos Excel (.xls)",
    type=["xls"],
    accept_multiple_files=True
)

if not archivos:
    st.info("👆 Por favor, sube al menos un archivo Excel (.xls) exportado desde la SMV.")
    st.stop()

# ------------------- UTILIDADES -------------------
def normalize_name(s):
    """Normaliza textos: quita tildes, pasa a mayúsculas, compacta espacios y elimina notas tipo (9) al final."""
    if not isinstance(s, str):
        return s
    s2 = unicodedata.normalize('NFKD', s).encode('ASCII', 'ignore').decode('ASCII')
    s2 = re.sub(r'\s+', ' ', s2).strip().upper()
    s2 = re.sub(r'\s*\(\d+\)\s*$', '', s2)
    return s2

def limpiar_valor(valor):
    """Limpia y convierte strings numéricos a float. Maneja paréntesis como negativos."""
    if valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    v = str(valor).strip()
    if v == '' or v == '0':
        return 0.0
    v = v.replace(',', '').replace('\xa0', '').replace(' ', '')
    # Paréntesis -> negativo
    if v.startswith('(') and v.endswith(')'):
        v = '-' + v[1:-1]
    # Eliminar símbolos no numéricos residuales
    v = re.sub(r'[^\d\-\.\+]', '', v)
    try:
        return float(v) if v != '' else 0.0
    except Exception:
        return 0.0

# ------------------- ESTRUCTURAS -------------------
datos_balance = {}      # datos_balance[anio][cuenta_key] = valor
datos_resultados = {}   # datos_resultados[anio][cuenta] = valor

# ------------------- PROCESAR ARCHIVOS -------------------
for archivo in archivos:
    st.write(f"📦 Procesando: {archivo.name}")

    contenido = None
    # El .xls exportado por SMV suele ser HTML dentro de un fichero .xls -> intentamos decodificar
    for cod in ['latin-1', 'cp1252', 'utf-8']:
        try:
            archivo.seek(0)
            raw = archivo.read()
            # raw puede ser bytes o str según streamlit; aseguramos bytes -> decode
            if isinstance(raw, bytes):
                contenido = raw.decode(cod, errors='ignore')
            else:
                contenido = str(raw)
            break
        except Exception:
            contenido = None
            continue

    if not contenido:
        st.error(f"❌ No se pudo leer {archivo.name} como texto HTML")
        continue

    soup = BeautifulSoup(contenido, 'html.parser')

    # ---------- TABLA BALANCE (id=gvReporte) ----------
    tabla_balance = soup.find('table', {'id': 'gvReporte'})
    if tabla_balance:
        filas = []
        for tr in tabla_balance.find_all('tr'):
            celdas = [td.get_text(strip=True) for td in tr.find_all(['td', 'th'])]
            if celdas:
                filas.append(celdas)

        if len(filas) > 1:
            encabezados = filas[0]
            columnas_anios = encabezados[2:]  # asume columnas: cuenta | descripcion | año1 | año2 ...
            anios = []
            for col in columnas_anios:
                m = re.search(r'\b(19|20)\d{2}\b', col)
                if m:
                    anios.append(int(m.group(0)))
                else:
                    anios.append(None)

            current_section = None
            for fila in filas[1:]:
                if len(fila) < 1:
                    continue
                first_cell = fila[0].strip()
                # Detectar encabezado de sección corto (sin montos)
                if re.search(r'^(ACTIVO|PASIVO|PATRIMONIO|TOTAL)\b', first_cell, flags=re.IGNORECASE) and len(fila) <= 2:
                    current_section = normalize_name(first_cell)
                    continue

                cuenta_raw = fila[0]
                cuenta_norm = normalize_name(cuenta_raw)

                for i, valor_str in enumerate(fila[2:]):
                    anio = anios[i] if i < len(anios) else None
                    if anio is None:
                        continue
                    valor = limpiar_valor(valor_str)

                    cuenta_key = f"{current_section}||{cuenta_norm}" if current_section else cuenta_norm

                    if anio not in datos_balance:
                        datos_balance[anio] = {}

                    if cuenta_key in datos_balance[anio]:
                        existing = datos_balance[anio][cuenta_key]
                        # No sobrescribir un existing no-cero con 0
                        if existing == 0 and valor != 0:
                            datos_balance[anio][cuenta_key] = valor
                    else:
                        datos_balance[anio][cuenta_key] = valor

    # ---------- TABLA RESULTADOS (id=gvReporte1) ----------
    tabla_resultados = soup.find('table', {'id': 'gvReporte1'})
    if tabla_resultados:
        filas = []
        for tr in tabla_resultados.find_all('tr'):
            celdas = [td.get_text(strip=True) for td in tr.find_all(['td', 'th'])]
            if celdas:
                filas.append(celdas)

        if len(filas) > 1:
            encabezados = filas[0]
            columnas_anios = encabezados[2:]
            anios = []
            for col in columnas_anios:
                m = re.search(r'\b(19|20)\d{2}\b', col)
                if m:
                    anios.append(int(m.group(0)))
                else:
                    anios.append(None)

            for fila in filas[1:]:
                if len(fila) < 1:
                    continue
                cuenta_raw = fila[0]
                cuenta_norm = normalize_name(cuenta_raw)
                for i, valor_str in enumerate(fila[2:]):
                    anio = anios[i] if i < len(anios) else None
                    if anio is None:
                        continue
                    valor = limpiar_valor(valor_str)
                    if anio not in datos_resultados:
                        datos_resultados[anio] = {}
                    if cuenta_norm in datos_resultados[anio]:
                        existing = datos_resultados[anio][cuenta_norm]
                        if existing == 0 and valor != 0:
                            datos_resultados[anio][cuenta_norm] = valor
                    else:
                        datos_resultados[anio][cuenta_norm] = valor

# ------------------- CREAR DATAFRAMES -------------------
if datos_balance:
    df_balance = pd.DataFrame.from_dict(datos_balance, orient='index').fillna(0.0).T
else:
    df_balance = pd.DataFrame()

if datos_resultados:
    df_resultados = pd.DataFrame.from_dict(datos_resultados, orient='index').fillna(0.0).T
else:
    df_resultados = pd.DataFrame()

# Normalizar índices (nombres de cuenta)
if not df_balance.empty:
    df_balance.index = [normalize_name(i) for i in df_balance.index]
if not df_resultados.empty:
    df_resultados.index = [normalize_name(i) for i in df_resultados.index]

# Reordenar columnas (años)
if not df_balance.empty:
    df_balance = df_balance.reindex(sorted(df_balance.columns), axis=1)
if not df_resultados.empty:
    df_resultados = df_resultados.reindex(sorted(df_resultados.columns), axis=1)

# ------------------- MANEJO INVENTARIOS -------------------
# Buscar variantes
if not df_balance.empty:
    keywords = ['INVENTARIOS', 'INVENTARIO', 'EXISTENCIAS', 'EXISTENCIA']
    inventarios_idx = [idx for idx in df_balance.index if any(k in idx for k in keywords)]

    inventarios_en_pasivo = [i for i in inventarios_idx if 'PASIVO' in i or 'PASIVOS' in i]
    inventarios_en_activo = [i for i in inventarios_idx if ('ACTIVO' in i or 'ACTIVOS' in i or 'CORRIENTE' in i)]
    
    df_balance_display = df_balance.copy()
    # Quedarse solo con columnas numéricas >= 2020
    df_balance_display = df_balance_display.loc[:, [
    col for col in df_balance_display.columns if str(col).isdigit() and int(col) >= 2020
    ]]

    if inventarios_en_pasivo:
        df_balance_display = df_balance_display.drop(index=inventarios_en_pasivo, errors='ignore')

    if inventarios_en_pasivo:
        df_inventarios_pasivo = df_balance.loc[inventarios_en_pasivo].copy()
    else:
        df_inventarios_pasivo = pd.DataFrame()
else:
    df_balance_display = pd.DataFrame()
    inventarios_idx = []
    inventarios_en_pasivo = []
    inventarios_en_activo = []
    df_inventarios_pasivo = pd.DataFrame()

def inventario_activos_valor(anio):
    if df_balance.empty:
        return 0.0
    for idx in inventarios_en_activo:
        if anio in df_balance.columns:
            val = df_balance.loc[idx, anio]
            if val != 0:
                return val
    for idx in inventarios_idx:
        if idx not in inventarios_en_pasivo:
            val = df_balance.loc[idx, anio]
            if val != 0:
                return val
    return 0.0

filas_resaltar = [
    "ACTIVOS",
    "ACTIVOS CORRIENTES",
    "TOTAL ACTIVOS CORRIENTES",
    "ACTIVOS NO CORRIENTES",
    "TOTAL ACTIVOS NO CORRIENTES",
    "TOTAL DE ACTIVOS",
    "PASIVOS Y PATRIMONIO",
    "PASIVOS CORRIENTES",
    "TOTAL PASIVOS CORRIENTES",
    "PASIVOS NO CORRIENTES",
    "TOTAL PASIVOS NO CORRIENTES",
    "TOTAL PASIVOS",
    "PATRIMONIO",
    "TOTAL PATRIMONIO",
    "TOTAL PASIVO Y PATRIMONIO"
]

def resaltar_filas(row):
    if row.name.strip().upper() in [f.upper() for f in filas_resaltar]:
        return ['background-color: #ffe599; font-weight: bold; color: black'] * len(row)
    return [''] * len(row)

# 🔹 Formatear números con 2 decimales y % donde corresponda
styled = (
    df_balance_display.style
    .apply(resaltar_filas, axis=1)
    .format("{:,.2f}")   # Mantener formato con 2 decimales
)

st.dataframe(styled)


# ------------------- ANALISIS VERTICAL -------------------
if not df_balance_display.empty:
    df_vertical = df_balance_display.copy()
    total_activos_row = None

    # Buscar fila "TOTAL ACTIVOS" (no corriente)
    for idx in df_vertical.index:
        if "TOTAL" in idx and ("ACTIVO" in idx or "ACTIVOS" in idx) and "CORRIENTE" not in idx:
            total_activos_row = idx
            break

    if total_activos_row is not None:
        total_activos = df_vertical.loc[total_activos_row]
        for col in df_vertical.columns:
            total_val = total_activos[col]
            if total_val != 0:
                df_vertical[col] = (df_vertical[col] / total_val) * 100
            else:
                df_vertical[col] = 0.0
        df_vertical = df_vertical.round(2)
    else:
        df_vertical = pd.DataFrame()
        st.warning("⚠️ No se encontró la fila 'TOTAL ACTIVOS' para el análisis vertical.")
else:
    df_vertical = pd.DataFrame()

# ------------------- ANALISIS HORIZONTAL -------------------
if not df_balance_display.empty and len(df_balance_display.columns) >= 2:
    df_horizontal_pct = pd.DataFrame(index=df_balance_display.index)
    columnas = df_balance_display.columns.tolist()

    for i in range(len(columnas) - 1):
        anio_ant = columnas[i]
        anio_act = columnas[i + 1]
        col_nombre = f"{anio_ant}-{anio_act}"

        with pd.option_context('mode.use_inf_as_na', True):
            df_horizontal_pct[col_nombre] = (
                (df_balance_display[anio_act] - df_balance_display[anio_ant]) /
                df_balance_display[anio_ant].replace({0: pd.NA})
            ) * 100

    # Redondear y mantener NaN (no None) para que el formato no falle
    df_horizontal_pct = df_horizontal_pct.round(2).replace([pd.NA], np.nan)
else:
    df_horizontal_pct = pd.DataFrame()



# ------------------- RATIOS FINANCIEROS -------------------
ratios_data = {}
anios_comunes = sorted(list(set(df_balance.columns) & set(df_resultados.columns))) if (not df_balance.empty and not df_resultados.empty) else []

def buscar_cuenta_por_keywords(df, keywords):
    """Busca la primera cuenta que contenga todas las palabras clave."""
    for idx in df.index:
        if all(kw.lower() in idx.lower() for kw in keywords):
            return idx
    return None

if anios_comunes:
    for i, anio in enumerate(anios_comunes):
        ratios_data.setdefault(anio, {})

        # Activo Corriente
        act_corr_idx = buscar_cuenta_por_keywords(df_balance, ["TOTAL", "ACTIVO", "CORRIENTE"])
        activo_corriente = df_balance.loc[act_corr_idx, anio] if act_corr_idx in df_balance.index else 0.0

        # Inventarios
        inv_idx = buscar_cuenta_por_keywords(df_balance, ["INVENTARIOS"]) or buscar_cuenta_por_keywords(df_balance, ["EXISTENCIAS"])
        inventarios = df_balance.loc[inv_idx, anio] if inv_idx in df_balance.index else inventario_activos_valor(anio)

        # Pasivo Corriente
        pas_corr_idx = buscar_cuenta_por_keywords(df_balance, ["TOTAL", "PASIVO", "CORRIENTE"])
        pasivo_corriente = df_balance.loc[pas_corr_idx, anio] if pas_corr_idx in df_balance.index else 0.0

        # Cuentas por Cobrar (sumar variantes)
        cxc_keywords_list = [
            ["CUENTAS", "COBRAR", "COMERCIALES"],
            ["CUENTAS", "COBRAR", "VINCULADAS"],
            ["OTRAS", "CUENTAS", "COBRAR"]
        ]
        cxc_val = 0.0
        for kws in cxc_keywords_list:
            idx = buscar_cuenta_por_keywords(df_balance, kws)
            if idx in df_balance.index:
                cxc_val += df_balance.loc[idx, anio]

        # Activos Totales
        act_tot_idx = buscar_cuenta_por_keywords(df_balance, ["TOTAL", "ACTIVO"])
        activos_totales = df_balance.loc[act_tot_idx, anio] if act_tot_idx in df_balance.index else 0.0

        # Pasivo Total
        pas_tot_idx = buscar_cuenta_por_keywords(df_balance, ["TOTAL", "PASIVO"])
        pasivo_total = df_balance.loc[pas_tot_idx, anio] if pas_tot_idx in df_balance.index else 0.0

        # Patrimonio
        patr_idx = (buscar_cuenta_por_keywords(df_balance, ["TOTAL", "PATRIMONIO", "NETO"]) or
                    buscar_cuenta_por_keywords(df_balance, ["PATRIMONIO", "NETO"]) or
                    buscar_cuenta_por_keywords(df_balance, ["CAPITAL", "EMITIDO"]))
        patrimonio = df_balance.loc[patr_idx, anio] if patr_idx in df_balance.index else 0.0

        # Estado de Resultados: Ventas, Costo, Utilidad Neta
        ventas_idx = (buscar_cuenta_por_keywords(df_resultados, ["INGRESOS", "ACTIVIDADES", "ORDINARIAS"]) or
                      buscar_cuenta_por_keywords(df_resultados, ["VENTAS", "NETAS"]) or
                      buscar_cuenta_por_keywords(df_resultados, ["VENTAS"]))
        ventas = df_resultados.loc[ventas_idx, anio] if ventas_idx in df_resultados.index else 0.0

        costo_idx = buscar_cuenta_por_keywords(df_resultados, ["COSTO", "VENTAS"])
        costo_ventas = df_resultados.loc[costo_idx, anio] if costo_idx in df_resultados.index else 0.0

        util_idx = (buscar_cuenta_por_keywords(df_resultados, ["UTILIDAD", "NETA"]) or
                    buscar_cuenta_por_keywords(df_resultados, ["GANANCIA", "NETA"]) or
                    buscar_cuenta_por_keywords(df_resultados, ["UTILIDAD", "PERDIDA", "NETA", "EJERCICIO"]))
        utilidad_neta = df_resultados.loc[util_idx, anio] if util_idx in df_resultados.index else 0.0

        # Promedios con año anterior (si existe)
        def promedio(a, b):
            # si ambos 0 -> 0, si uno 0 -> (a+b)/2 normal
            return (a + b) / 2.0

        cxc_prom = cxc_val
        inv_prom = inventarios
        act_prom = activos_totales
        patr_prom = patrimonio

        if i > 0:
            anio_ant = anios_comunes[i - 1]

            # cxc anterior
            cxc_ant = 0.0
            for kws in cxc_keywords_list:
                idx = buscar_cuenta_por_keywords(df_balance, kws)
                if idx in df_balance.index:
                    cxc_ant += df_balance.loc[idx, anio_ant]

            # inventarios anterior
            inv_ant = df_balance.loc[inv_idx, anio_ant] if (inv_idx in df_balance.index) else inventario_activos_valor(anio_ant)

            # activos totales anterior
            act_ant = df_balance.loc[act_tot_idx, anio_ant] if (act_tot_idx in df_balance.index) else 0.0

            # patrimonio anterior
            patr_ant = df_balance.loc[patr_idx, anio_ant] if (patr_idx in df_balance.index) else 0.0

            # calcular promedios pero cuidando ceros
            cxc_prom = promedio(cxc_val, cxc_ant) if (cxc_val != 0 or cxc_ant != 0) else 0.0
            inv_prom = promedio(inventarios, inv_ant) if (inventarios != 0 or inv_ant != 0) else 0.0
            act_prom = promedio(activos_totales, act_ant) if (activos_totales != 0 or act_ant != 0) else 0.0
            patr_prom = promedio(patrimonio, patr_ant) if (patrimonio != 0 or patr_ant != 0) else 0.0

        # --- calcular ratios con defensas ante division por cero ---
        def safe_div(num, den):
            try:
                return num / den if den not in (0, None) else None
            except Exception:
                return None

        ratios_data[anio]["Liquidez Corriente"] = safe_div(activo_corriente, pasivo_corriente)
        ratios_data[anio]["Prueba Ácida"] = safe_div((activo_corriente - inventarios), pasivo_corriente)
        ratios_data[anio]["Rotación CxC"] = safe_div(ventas, cxc_prom) if cxc_prom not in (0, None) else None
        ratios_data[anio]["Rotación Inventarios"] = safe_div(abs(costo_ventas), inv_prom) if inv_prom not in (0, None) else None
        ratios_data[anio]["Rotación Activos Totales"] = safe_div(ventas, act_prom) if act_prom not in (0, None) else None
        ratios_data[anio]["Razón Deuda Total"] = safe_div(pasivo_total, activos_totales)
        ratios_data[anio]["Razón Deuda/Patrimonio"] = safe_div(pasivo_total, patrimonio)
        ratios_data[anio]["Margen Neto"] = safe_div(utilidad_neta, ventas)
        ratios_data[anio]["ROA"] = safe_div(utilidad_neta, act_prom)
        ratios_data[anio]["ROE"] = safe_div(utilidad_neta, patr_prom)

    # convertir a DataFrame orientado por años (filas=años, columnas=ratios)
    df_ratios = pd.DataFrame.from_dict(ratios_data, orient='index').round(4) if ratios_data else pd.DataFrame()
else:
    df_ratios = pd.DataFrame()

# ------------------- OUTPUT / UI -------------------
st.markdown("---")
if not df_balance_display.empty:
    st.subheader("📊 Estado de Situación Financiera (consolidado)")
    st.dataframe(df_balance_display.style.format("{:,.2f}"))

if not df_resultados.empty:
    st.subheader("📊 Estado de Resultados (consolidado)")
    st.dataframe(df_resultados.style.format("{:,.2f}"))

if not df_vertical.empty:
    st.subheader("📈 Análisis Vertical (Balance)")
    st.dataframe(df_vertical.style.format("{:,.2f}%"))

if not df_horizontal_pct.empty:
    st.subheader("📈 Análisis Horizontal (Balance) - % Variación")
    st.dataframe(df_horizontal_pct.style.format("{:,.2f}%"))

if not df_ratios.empty:
    st.subheader("📈 Ratios Financieros")
    st.dataframe(df_ratios.style.format("{:,.4f}"))

# ------------------- OBSERVACIONES AUTOMÁTICAS (BÁSICAS) -------------------
def generar_observaciones(df_ratios):
    obs = []
    if df_ratios.empty:
        return obs
    # Tomar último año disponible
    ultimo_anio = df_ratios.index.max()
    row = df_ratios.loc[ultimo_anio]
    # Liquidez
    liq = row.get("Liquidez Corriente")
    if liq is not None:
        if liq < 1:
            obs.append(f"🔴 Liquidez corriente ({ultimo_anio}) = {liq:.2f}: posible problema de corto plazo.")
        elif liq < 1.5:
            obs.append(f"🟠 Liquidez corriente ({ultimo_anio}) = {liq:.2f}: aceptable pero mejorar.")
        else:
            obs.append(f"🟢 Liquidez corriente ({ultimo_anio}) = {liq:.2f}: cómodo nivel de liquidez.")
    # Endeudamiento
    deuda_total = row.get("Razón Deuda Total")
    if deuda_total is not None:
        if deuda_total > 0.7:
            obs.append(f"🔴 Alta proporción de deuda sobre activos ({deuda_total:.2f}).")
        elif deuda_total > 0.4:
            obs.append(f"🟠 Moderada deuda sobre activos ({deuda_total:.2f}).")
        else:
            obs.append(f"🟢 Bajo nivel de endeudamiento ({deuda_total:.2f}).")
    # Rentabilidad
    roa = row.get("ROA")
    roe = row.get("ROE")
    if roa is not None:
        obs.append(f"➡️ ROA ({ultimo_anio}) = {roa:.4f}")
    if roe is not None:
        obs.append(f"➡️ ROE ({ultimo_anio}) = {roe:.4f}")
    return obs

obs = generar_observaciones(df_ratios)
if obs:
    st.subheader("📝 Observaciones automáticas")
    for o in obs:
        st.write(o)

# ------------------- EXPORTAR A EXCEL -------------------
output = io.BytesIO()
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    if not df_balance_display.empty:
        df_balance_display.to_excel(writer, sheet_name='ESTADO_SITUACION_FINANCIERA', index_label='Cuenta')
    if not df_resultados.empty:
        df_resultados.to_excel(writer, sheet_name='ESTADO_RESULTADOS', index_label='Cuenta')
    if not df_vertical.empty:
        df_vertical.to_excel(writer, sheet_name='ANALISIS_VERTICAL_BALANCE', index_label='Cuenta')
    if not df_horizontal_pct.empty:
        df_horizontal_pct.to_excel(writer, sheet_name='ANALISIS_HORIZONTAL_BALANCE', index_label='Cuenta')
    if not df_ratios.empty:
        df_ratios.to_excel(writer, sheet_name='RATIOS_FINANCIEROS', index_label='AÑO')
    if not df_inventarios_pasivo.empty:
        df_inventarios_pasivo.to_excel(writer, sheet_name='INVENTARIOS_PASIVO', index_label='Cuenta')

# Preparar descarga
st.download_button(
    label="📥 Descargar Excel Consolidado (Completo)",
    data=output.getvalue(),
    file_name="Consolidado_Estados_Financieros_Completo.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.success("✅ ¡Proceso completado! El archivo incluye estados financieros, análisis V/H y los ratios calculados.")
