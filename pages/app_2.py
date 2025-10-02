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

def buscar_cuenta_exacta(df, keywords):
    """Busca una cuenta que contenga todas las palabras clave (case-insensitive)."""
    for idx in df.index:
        if all(kw.lower() in idx.lower() for kw in keywords):
            return idx
    return None

if anios_comunes:
    for i, anio in enumerate(anios_comunes):
        ratios_data[anio] = {}

        # --- Balance General ---
        # Activo Corriente
        act_corr = buscar_cuenta_exacta(df_balance, ["TOTAL", "ACTIVO", "CORRIENTE"])
        activo_corriente = df_balance.loc[act_corr, anio] if act_corr else 0.0

        # Inventarios / Existencias
        inv1 = buscar_cuenta_exacta(df_balance, ["INVENTARIOS"])
        inv2 = buscar_cuenta_exacta(df_balance, ["EXISTENCIAS"])
        inventarios = df_balance.loc[inv1 or inv2, anio] if (inv1 or inv2) else 0.0

        # Pasivo Corriente
        pas_corr = buscar_cuenta_exacta(df_balance, ["TOTAL", "PASIVO", "CORRIENTE"])
        pasivo_corriente = df_balance.loc[pas_corr, anio] if pas_corr else 0.0

        # Cuentas por Cobrar
        cxc1 = buscar_cuenta_exacta(df_balance, ["CUENTAS", "COBRAR", "COMERCIALES"])
        cxc2 = buscar_cuenta_exacta(df_balance, ["CUENTAS", "COBRAR", "VINCULADAS"])
        cxc3 = buscar_cuenta_exacta(df_balance, ["OTRAS", "CUENTAS", "COBRAR"])
        
        cxc_val = 0.0
        for cxc in [cxc1, cxc2, cxc3]:
            if cxc and cxc in df_balance.index:
                cxc_val += df_balance.loc[cxc, anio]

        # Activos Totales
        act_tot = buscar_cuenta_exacta(df_balance, ["TOTAL", "ACTIVO"])
        activos_totales = df_balance.loc[act_tot, anio] if act_tot else 0.0

        # Pasivo Total
        pas_tot = buscar_cuenta_exacta(df_balance, ["TOTAL", "PASIVO"])
        pasivo_total = df_balance.loc[pas_tot, anio] if pas_tot else 0.0

        # Patrimonio Neto
        patr1 = buscar_cuenta_exacta(df_balance, ["TOTAL", "PATRIMONIO", "NETO"])
        patr2 = buscar_cuenta_exacta(df_balance, ["PATRIMONIO", "NETO"])
        patr3 = buscar_cuenta_exacta(df_balance, ["CAPITAL", "EMITIDO"])  # Si no hay total, usar capital como proxy
        patrimonio = df_balance.loc[patr1 or patr2 or patr3, anio] if (patr1 or patr2 or patr3) else 0.0

        # --- Estado de Resultados ---
        # Ventas
        ventas1 = buscar_cuenta_exacta(df_resultados, ["INGRESOS", "ACTIVIDADES", "ORDINARIAS"])
        ventas2 = buscar_cuenta_exacta(df_resultados, ["VENTAS", "NETAS"])
        ventas = df_resultados.loc[ventas1 or ventas2, anio] if (ventas1 or ventas2) else 0.0

        # Costo de Ventas
        costo1 = buscar_cuenta_exacta(df_resultados, ["COSTO", "VENTAS"])
        costo2 = buscar_cuenta_exacta(df_resultados, ["COSTO", "VENTAS", "OPERACIONALES"])
        costo_ventas = df_resultados.loc[costo1 or costo2, anio] if (costo1 or costo2) else 0.0

        # Utilidad Neta
        util1 = buscar_cuenta_exacta(df_resultados, ["UTILIDAD", "PERDIDA", "NETA", "EJERCICIO"])
        util2 = buscar_cuenta_exacta(df_resultados, ["GANANCIA", "PERDIDA", "NETA", "EJERCICIO"])
        utilidad_neta = df_resultados.loc[util1 or util2, anio] if (util1 or util2) else 0.0

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
            cxc_ant = 0.0
            for cxc in [cxc1, cxc2, cxc3]:
                if cxc and cxc in df_balance.index:
                    cxc_ant += df_balance.loc[cxc, anio_ant]
            
            # Inventarios anterior
            inv_ant = df_balance.loc[inv1 or inv2, anio_ant] if (inv1 or inv2) else 0.0
            
            # Activos y Patrimonio anteriores
            act_ant = df_balance.loc[act_tot, anio_ant] if act_tot else 0.0
            patr_ant = df_balance.loc[patr1 or patr2 or patr3, anio_ant] if (patr1 or patr2 or patr3) else 0.0

            cxc_prom = promedio(cxc_val, cxc_ant)
            inv_prom = promedio(inventarios, inv_ant)
            act_prom = promedio(activos_totales, act_ant)
            patr_prom = promedio(patrimonio, patr_ant)

        # --- Cálculo de ratios ---
        ratios_data[anio]["Liquidez Corriente"] = activo_corriente / pasivo_corriente if pasivo_corriente != 0 else None
        ratios_data[anio]["Prueba Ácida"] = (activo_corriente - inventarios) / pasivo_corriente if pasivo_corriente != 0 else None
        ratios_data[anio]["Rotación CxC"] = ventas / cxc_prom if cxc_prom != 0 else None
        ratios_data[anio]["Rotación Inventarios"] = abs(costo_ventas) / inv_prom if inv_prom != 0 else None
        ratios_data[anio]["Rotación Activos Totales"] = ventas / act_prom if act_prom != 0 else None
        ratios_data[anio]["Razón Deuda Total"] = pasivo_total / activos_totales if activos_totales != 0 else None
        ratios_data[anio]["Razón Deuda/Patrimonio"] = pasivo_total / patrimonio if patrimonio != 0 else None
        ratios_data[anio]["Margen Neto"] = utilidad_neta / ventas if ventas != 0 else None
        ratios_data[anio]["ROA"] = utilidad_neta / act_prom if act_prom != 0 else None
        ratios_data[anio]["ROE"] = utilidad_neta / patr_prom if patr_prom != 0 else None

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
