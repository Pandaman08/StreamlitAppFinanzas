import streamlit as st
import pandas as pd
import io
from bs4 import BeautifulSoup
import re
import unicodedata

st.set_page_config(page_title="Consolidador SMV - Finanzas Corporativas", layout="wide")
st.title("📊 Consolidador de Estados Financieros - SMV")
st.markdown("Sube los archivos Excel descargados de la Superintendencia del Mercado de Valores (SMV).")

archivos = st.file_uploader(
    "Selecciona archivos Excel (.xls)",
    type=["xls"],
    accept_multiple_files=True
)

if not archivos:
    st.info("👆 Por favor, sube al menos un archivo Excel.")
    st.stop()

# ------------------- Utilidades -------------------

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
    if not valor or valor == '0' or valor == '':
        return 0.0
    valor = str(valor).replace(',', '').replace('\xa0', '').replace(' ', '')
    # Manejar paréntesis (números negativos en formatos contables)
    if valor.startswith('(') and valor.endswith(')'):
        valor = '-' + valor[1:-1]
    try:
        return float(valor)
    except:
        return 0.0

# ------------------- Contenedores -------------------
# Estructura: datos_balance[anio][cuenta] = valor
# Nota: mantendremos esta estructura y luego transformaremos a DataFrame con cuentas como índice.

datos_balance = {}
datos_resultados = {}

# ------------------- Procesar archivos -------------------
for archivo in archivos:
    st.write(f"📦 Procesando: {archivo.name}")

    contenido = None
    # Intentar decodificar como texto (HTML oculto en .xls)
    for cod in ['latin-1', 'cp1252', 'utf-8']:
        try:
            archivo.seek(0)
            contenido = archivo.read().decode(cod)
            break
        except Exception:
            continue

    if not contenido:
        st.error(f"❌ No se pudo leer {archivo.name} como texto HTML")
        continue

    soup = BeautifulSoup(contenido, 'html.parser')

    # ---------- Procesar Balance General (tabla id=gvReporte) ----------
    tabla_balance = soup.find('table', {'id': 'gvReporte'})
    if tabla_balance:
        filas = []
        for tr in tabla_balance.find_all('tr'):
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

            # Mantener contexto de sección (ej. ACTIVOS CORRIENTES / PASIVOS CORRIENTES)
            current_section = None
            for fila in filas[1:]:
                if len(fila) < 2:
                    continue
                first_cell = fila[0].strip()
                # Detectar si la fila es un encabezado de sección
                if re.search(r'ACTIVO|PASIVO|PATRIMONIO|TOTAL', first_cell, flags=re.IGNORECASE) and len(fila) <= 2:
                    # Consideramos los headers cortos (sin montos) como marcadores de sección
                    current_section = normalize_name(first_cell)
                    continue

                cuenta_raw = fila[0]
                for i, valor_str in enumerate(fila[2:]):
                    anio = anios[i]
                    if anio is None:
                        continue
                    valor = limpiar_valor(valor_str)

                    # Normalizar el nombre de la cuenta para evitar duplicados por espacios/tildes/notas
                    cuenta = normalize_name(cuenta_raw)

                    # Construir una llave que preserve sección si existe
                    cuenta_key = f"{current_section}||{cuenta}" if current_section else cuenta

                    if anio not in datos_balance:
                        datos_balance[anio] = {}

                    # Lógica: NO sobrescribir un valor no-cero con un cero (las tablas SMV a veces agregan filas de ceros duplicadas)
                    if cuenta_key in datos_balance[anio]:
                        existing = datos_balance[anio][cuenta_key]
                        if existing == 0 and valor != 0:
                            datos_balance[anio][cuenta_key] = valor
                        # si existing != 0 y valor == 0 -> mantener existing
                        # si ambos != 0 -> mantener existing (evita elección arbitraria)
                    else:
                        datos_balance[anio][cuenta_key] = valor

    # ---------- Procesar Estado de Resultados (tabla id=gvReporte1) ----------
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
                if len(fila) < 2:
                    continue
                cuenta_raw = fila[0]
                for i, valor_str in enumerate(fila[2:]):
                    anio = anios[i]
                    if anio is None:
                        continue
                    valor = limpiar_valor(valor_str)

                    cuenta = normalize_name(cuenta_raw)

                    if anio not in datos_resultados:
                        datos_resultados[anio] = {}

                    if cuenta in datos_resultados[anio]:
                        existing = datos_resultados[anio][cuenta]
                        if existing == 0 and valor != 0:
                            datos_resultados[anio][cuenta] = valor
                    else:
                        datos_resultados[anio][cuenta] = valor

# ------------------- Crear DataFrames (cuentas como índice, años como columnas) -------------------
# Creamos DataFrames con orient='index' (índice = año) y luego los transponemos para tener cuentas como índice
if datos_balance:
    df_balance = pd.DataFrame.from_dict(datos_balance, orient='index').fillna(0.0).T
else:
    df_balance = pd.DataFrame()

if datos_resultados:
    df_resultados = pd.DataFrame.from_dict(datos_resultados, orient='index').fillna(0.0).T
else:
    df_resultados = pd.DataFrame()

# Normalizar índices (nombres de cuenta) - ya normalizamos previamente, pero nos aseguramos
if not df_balance.empty:
    df_balance.index = [normalize_name(i) for i in df_balance.index]
if not df_resultados.empty:
    df_resultados.index = [normalize_name(i) for i in df_resultados.index]

# Reordenar columnas (años) en orden ascendente
if not df_balance.empty:
    df_balance = df_balance.reindex(sorted(df_balance.columns), axis=1)
if not df_resultados.empty:
    df_resultados = df_resultados.reindex(sorted(df_resultados.columns), axis=1)

# ------------------- Manejo especial de INVENTARIOS / EXISTENCIAS -------------------
# Buscamos variantes: 'INVENTARIOS', 'INVENTARIO', 'EXISTENCIAS', 'EXISTENCIA'
if not df_balance.empty:
    keywords = ['INVENTARIOS', 'INVENTARIO', 'EXISTENCIAS', 'EXISTENCIA']
    inventarios_idx = [idx for idx in df_balance.index if any(k in idx for k in keywords)]

    # Clasificar según presencia de palabras clave de sección en el índice
    inventarios_en_pasivo = [i for i in inventarios_idx if 'PASIVO' in i or 'PASIVOS' in i]
    inventarios_en_activo = [i for i in inventarios_idx if ('ACTIVO' in i or 'ACTIVOS' in i or 'CORRIENTE' in i)]

    # Vista para la UI (excluir inventarios que claramente pertenecen a PASIVO)
    df_balance_display = df_balance.copy()
    if inventarios_en_pasivo:
        df_balance_display = df_balance_display.drop(index=inventarios_en_pasivo, errors='ignore')
    else:
        df_balance_display = df_balance.copy()

    # DataFrame con inventarios en pasivo para auditoría
    if inventarios_en_pasivo:
        df_inventarios_pasivo = df_balance.loc[inventarios_en_pasivo].copy()
    else:
        df_inventarios_pasivo = pd.DataFrame()
else:
    df_balance_display = pd.DataFrame()
    df_inventarios_pasivo = pd.DataFrame()

# Función que devuelve el inventario correcto para un año (prioriza activos y variantes EXISTENCIA/EXISTENCIAS)
def inventario_activos_valor(anio):
    if df_balance.empty:
        return 0.0
    # 1) Prioriza inventarios dentro de activos (no-cero)
    for idx in inventarios_en_activo:
        if anio in df_balance.columns:
            val = df_balance.loc[idx, anio]
            if val != 0:
                return val
    # 2) Si no hay candidatos activos, busca cualquier 'INVENTARIOS' no en pasivo (cualquier no-cero)
    for idx in inventarios_idx:
        if idx not in inventarios_en_pasivo:
            val = df_balance.loc[idx, anio]
            if val != 0:
                return val
    # 3) Si no hay nada, devolver 0.0
    return 0.0

# ------------------- ANÁLISIS VERTICAL (Balance) -------------------
if not df_balance_display.empty:
    df_vertical = df_balance_display.copy()
    total_activos_row = None
    for idx in df_vertical.index:
        total_activos_row = None
        for idx in df_vertical.index:
            if "TOTAL" in idx and ("ACTIVOS" in idx or "ACTIVO" in idx) and "CORRIENTE" not in idx:
                total_activos_row = idx
                break

    if total_activos_row is not None:
        total_activos = df_vertical.loc[total_activos_row]
        for col in df_vertical.columns:
            if total_activos[col] != 0:
                df_vertical[col] = (df_vertical[col] / total_activos[col]) * 100
            else:
                df_vertical[col] = 0.0
        df_vertical = df_vertical.round(2)
    else:
        st.warning("⚠️ No se encontró la fila 'TOTAL DE ACTIVOS' para el análisis vertical.")
else:
    df_vertical = pd.DataFrame()

# ------------------- ANÁLISIS HORIZONTAL (Balance) -------------------
if not df_balance_display.empty:
    df_horizontal_pct = df_balance_display.copy()
    columnas = df_horizontal_pct.columns.tolist()
    nuevas_columnas = []
    for i in range(len(columnas) - 1):
        anio_actual = columnas[i + 1]
        anio_anterior = columnas[i]
        col_nombre = f"{anio_anterior}-{anio_actual}"
        # Evitar división por cero en caso el año anterior sea 0 -> NaN
        df_horizontal_pct[col_nombre] = (df_horizontal_pct[anio_actual] - df_horizontal_pct[anio_anterior]) / df_horizontal_pct[anio_anterior] * 100
        nuevas_columnas.append(col_nombre)

    df_horizontal_pct = df_horizontal_pct[nuevas_columnas].round(2)
    df_horizontal_pct = df_horizontal_pct.replace([float('inf'), float('-inf')], pd.NA)
else:
    df_horizontal_pct = pd.DataFrame()

# ------------------- CÁLCULO DE RATIOS FINANCIEROS -------------------
ratios_data = {}
anios_comunes = sorted(list(set(df_balance.columns) & set(df_resultados.columns))) if (not df_balance.empty and not df_resultados.empty) else []

def buscar_cuenta_flexible(df, keywords, fallback_keywords=None):
    """
    Busca una cuenta que contenga AL MENOS UNA de las palabras clave (case-insensitive).
    Si no encuentra nada, intenta con las palabras clave de fallback.
    """
    # Primera pasada: buscar con las palabras clave principales
    for idx in df.index:
        idx_lower = idx.lower()
        if any(kw.lower() in idx_lower for kw in keywords):
            return idx

    # Segunda pasada: si no encontró con las principales, intenta con las de fallback
    if fallback_keywords:
        for idx in df.index:
            idx_lower = idx.lower()
            if any(kw.lower() in idx_lower for kw in fallback_keywords):
                return idx

    # Tercera pasada: si aún no, busca por coincidencia parcial más amplia
    for idx in df.index:
        idx_lower = idx.lower()
        # Verificar si alguna palabra clave principal está contenida en el índice
        for kw in keywords:
            if kw.lower() in idx_lower:
                return idx

    # Cuarta pasada: si todo falla, devuelve None
    return None

if anios_comunes:
    for i, anio in enumerate(anios_comunes):
        ratios_data[anio] = {}

        # --- Balance General ---
        # Activo Corriente
        act_corr_keywords = ["TOTAL", "ACTIVO", "CORRIENTE"]
        act_corr = buscar_cuenta_flexible(df_balance, act_corr_keywords)
        activo_corriente = df_balance.loc[act_corr, anio] if act_corr else 0.0

        # Inventarios / Existencias
        inv_keywords = ["INVENTARIOS", "EXISTENCIAS"]
        inv_fallback = ["INVENTARIO", "EXISTENCIA"]  # Fallback
        inventarios = 0.0
        inv_row = buscar_cuenta_flexible(df_balance, inv_keywords, inv_fallback)
        if inv_row:
            inventarios = df_balance.loc[inv_row, anio]

        # Pasivo Corriente
        pas_corr_keywords = ["TOTAL", "PASIVO", "CORRIENTE"]
        pas_corr = buscar_cuenta_flexible(df_balance, pas_corr_keywords)
        pasivo_corriente = df_balance.loc[pas_corr, anio] if pas_corr else 0.0

        # Cuentas por Cobrar
        cxc_keywords = ["CUENTAS", "COBRAR", "COMERCIALES"]
        cxc_rows = [idx for idx in df_balance.index if any(kw.lower() in idx.lower() for kw in cxc_keywords)]
        cxc_val = sum(df_balance.loc[r, anio] for r in cxc_rows) if cxc_rows else 0.0

        # Activos Totales
        act_tot_keywords = ["TOTAL", "ACTIVO"]
        act_tot = buscar_cuenta_flexible(df_balance, act_tot_keywords)
        activos_totales = df_balance.loc[act_tot, anio] if act_tot else 0.0

        # Pasivo Total
        pas_tot_keywords = ["TOTAL", "PASIVO"]
        pas_tot = buscar_cuenta_flexible(df_balance, pas_tot_keywords)
        pasivo_total = df_balance.loc[pas_tot, anio] if pas_tot else 0.0

        # Patrimonio Neto
        patr_keywords = ["TOTAL", "PATRIMONIO", "NETO"]
        patr_fallback = ["CAPITAL", "EMITIDO"]  # Fallback si no encuentra PATRIMONIO NETO
        patr = buscar_cuenta_flexible(df_balance, patr_keywords, patr_fallback)
        patrimonio = df_balance.loc[patr, anio] if patr else 0.0

        # --- Estado de Resultados ---
        # Ventas
        ventas_keywords = ["VENTAS", "INGRESOS", "ORDINARIAS", "BRUTOS"]
        ventas_fallback = ["INGRESOS", "OPERACIONALES"]  # Fallback
        ventas = 0.0
        ventas_row = buscar_cuenta_flexible(df_resultados, ventas_keywords, ventas_fallback)
        if ventas_row:
            ventas = df_resultados.loc[ventas_row, anio]

        # Costo de Ventas
        costo_ventas_keywords = ["COSTO", "VENTAS", "OPERACIONALES"]
        costo_ventas_fallback = ["COSTO", "VENTAS"]  # Fallback
        costo_ventas = 0.0
        costo_ventas_row = buscar_cuenta_flexible(df_resultados, costo_ventas_keywords, costo_ventas_fallback)
        if costo_ventas_row:
            costo_ventas = df_resultados.loc[costo_ventas_row, anio]

        # Utilidad Neta
        utilidad_neta_keywords = ["UTILIDAD", "PERDIDA", "NETA", "EJERCICIO"]
        utilidad_neta_fallback = ["GANANCIA", "PERDIDA", "NETA"]  # Fallback
        utilidad_neta = 0.0
        utilidad_neta_row = buscar_cuenta_flexible(df_resultados, utilidad_neta_keywords, utilidad_neta_fallback)
        if utilidad_neta_row:
            utilidad_neta = df_resultados.loc[utilidad_neta_row, anio]

        # --- Promedios con año anterior ---
        def promedio(actual, anterior):
            return (actual + anterior) / 2 if (actual + anterior) != 0 else actual

        cxc_prom = cxc_val
        inv_prom = inventarios
        act_prom = activos_totales
        patr_prom = patrimonio

        if i > 0:
            anio_ant = anios_comunes[i-1]
            
            # CxC anterior
            cxc_ant = sum(df_balance.loc[r, anio_ant] for r in cxc_rows) if cxc_rows else 0.0
            
            # Inventarios anterior
            inv_ant = 0.0
            if inv_row:
                inv_ant = df_balance.loc[inv_row, anio_ant]
            
            # Activos y Patrimonio anteriores
            act_ant = df_balance.loc[act_tot, anio_ant] if act_tot else 0.0
            patr_ant = df_balance.loc[patr, anio_ant] if patr else 0.0

            cxc_prom = promedio(cxc_val, cxc_ant)
            inv_prom = promedio(inventarios, inv_ant)
            act_prom = promedio(activos_totales, act_ant)
            patr_prom = promedio(patrimonio, patr_ant)

        # --- Cálculo de ratios ---
        # 1. Liquidez Corriente
        if pasivo_corriente != 0:
            ratios_data[anio]["Liquidez Corriente"] = activo_corriente / pasivo_corriente
        else:
            ratios_data[anio]["Liquidez Corriente"] = None

        # 2. Prueba Ácida
        if pasivo_corriente != 0:
            ratios_data[anio]["Prueba Ácida"] = (activo_corriente - inventarios) / pasivo_corriente
        else:
            ratios_data[anio]["Prueba Ácida"] = None

        # 3. Rotación CxC
        if cxc_prom != 0:
            ratios_data[anio]["Rotación CxC"] = ventas / cxc_prom
        else:
            ratios_data[anio]["Rotación CxC"] = None

        # 4. Rotación Inventarios
        if inv_prom != 0:
            ratios_data[anio]["Rotación Inventarios"] = abs(costo_ventas) / inv_prom
        else:
            ratios_data[anio]["Rotación Inventarios"] = None

        # 5. Rotación Activos Totales
        if act_prom != 0:
            ratios_data[anio]["Rotación Activos Totales"] = ventas / act_prom
        else:
            ratios_data[anio]["Rotación Activos Totales"] = None

        # 6. Razón Deuda Total
        if activos_totales != 0:
            ratios_data[anio]["Razón Deuda Total"] = pasivo_total / activos_totales
        else:
            ratios_data[anio]["Razón Deuda Total"] = None

        # 7. Razón Deuda/Patrimonio
        if patrimonio != 0:
            ratios_data[anio]["Razón Deuda/Patrimonio"] = pasivo_total / patrimonio
        else:
            ratios_data[anio]["Razón Deuda/Patrimonio"] = None

        # 8. Margen Neto
        if ventas != 0:
            ratios_data[anio]["Margen Neto"] = utilidad_neta / ventas
        else:
            ratios_data[anio]["Margen Neto"] = None

        # 9. ROA
        if act_prom != 0:
            ratios_data[anio]["ROA"] = utilidad_neta / act_prom
        else:
            ratios_data[anio]["ROA"] = None

        # 10. ROE
        if patr_prom != 0:
            ratios_data[anio]["ROE"] = utilidad_neta / patr_prom
        else:
            ratios_data[anio]["ROE"] = None

    df_ratios = pd.DataFrame.from_dict(ratios_data, orient='index').round(4).T
else:
    df_ratios = pd.DataFrame()

# ------------------- Mostrar resultados -------------------
if not df_balance_display.empty:
    st.subheader("📊 Estado de Situación Financiera")
    st.dataframe(df_balance_display)

if not df_resultados.empty:
    st.subheader("📊 Estado de Resultados")
    st.dataframe(df_resultados)

if not df_ratios.empty:
    st.subheader("📈 Ratios Financieros")
    st.dataframe(df_ratios)

output = io.BytesIO()
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    if not df_balance_display.empty:
        df_balance_display.to_excel(writer, sheet_name='ESTADO DE SITUACION FINANCIERA', index_label='Cuenta')
    if not df_resultados.empty:
        df_resultados.to_excel(writer, sheet_name='ESTADO DE RESULTADOS', index_label='Cuenta')
    if not df_vertical.empty:
        df_vertical.to_excel(writer, sheet_name='ANALISIS_VERTICAL_BALANCE', index_label='Cuenta')
    if not df_horizontal_pct.empty:
        df_horizontal_pct.to_excel(writer, sheet_name='ANALISIS_HORIZONTAL_BALANCE', index_label='Cuenta')

    # Solo escribir la hoja con los ratios calculados
    if not df_ratios.empty:
        df_ratios.to_excel(writer, sheet_name='RATIOS_FINANCIEROS', index_label='Ratio')

    # Escribir la hoja de inventarios en pasivo para auditoría si existe
    if not df_inventarios_pasivo.empty:
        df_inventarios_pasivo.to_excel(writer, sheet_name='INVENTARIOS_PASIVO', index_label='Cuenta')

st.download_button(
    label="📥 Descargar Excel Consolidado (Completo)",
    data=output.getvalue(),
    file_name="Consolidado_Estados_Financieros_Completo.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.success("✅ ¡Proceso completado! El archivo incluye estados financieros, análisis V/H y los ratios solicitados.")
