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


def mapear_cuenta_normalizada(cuenta_original, anio):
    """
    Mapea nombres de cuentas antiguas (pre-2010) a nomenclatura moderna.
    Devuelve el nombre normalizado estándar.
    """
    cuenta = normalize_name(cuenta_original)
    
    # Diccionario de mapeo para cuentas pre-2010 a post-2010
    mapeo_antiguo = {
        # Balance General - Activos
        "CAJA Y BANCOS": "EFECTIVO Y EQUIVALENTES AL EFECTIVO",
        "VALORES NEGOCIABLES": "OTROS ACTIVOS FINANCIEROS CORRIENTES",
        "EXISTENCIAS": "INVENTARIOS",
        "GASTOS PAGADOS POR ANTICIPADO": "ANTICIPOS",
        "INVERSIONES PERMANENTES": "INVERSIONES EN SUBSIDIARIAS NEGOCIOS CONJUNTOS Y ASOCIADAS",
        "INMUEBLES MAQUINARIA Y EQUIPO NETO DE DEPRECIACION ACUMULADA": "PROPIEDADES PLANTA Y EQUIPO",
        "ACTIVO INTANGIBLE NETO DE DEPRECIACION ACUMULADA": "ACTIVOS INTANGIBLES DISTINTOS DE LA PLUSVALIA",
        "OTROS ACTIVOS": "OTROS ACTIVOS NO FINANCIEROS NO CORRIENTES",
        "IMPUESTO A LA RENTA Y PARTICIPACIONES DIFERIDOS ACTIVO": "ACTIVOS POR IMPUESTOS DIFERIDOS",
        
        # Balance General - Pasivos
        "SOBREGIROS Y PAGARES BANCARIOS": "OTROS PASIVOS FINANCIEROS CORRIENTES",
        "PARTE CORRIENTE DE LAS DEUDAS A LARGO PLAZO": "OTROS PASIVOS FINANCIEROS CORRIENTES",
        "DEUDAS A LARGO PLAZO": "OTROS PASIVOS FINANCIEROS NO CORRIENTES",
        "INGRESOS DIFERIDOS": "INGRESOS DIFERIDOS",
        "IMPUESTO A LA RENTA Y PARTICIPACIONES DIFERIDOS PASIVO": "PASIVOS POR IMPUESTOS DIFERIDOS",
        
        # Balance General - Patrimonio
        "CAPITAL": "CAPITAL EMITIDO",
        "CAPITAL ADICIONAL": "PRIMAS DE EMISION",
        "EXCEDENTE DE REVALUACION": "SUPERAVIT DE REVALUACION",
        "RESERVAS LEGALES": "OTRAS RESERVAS DE CAPITAL",
        "OTRAS RESERVAS": "OTRAS RESERVAS DE PATRIMONIO",
        "RESULTADOS ACUMULADOS": "RESULTADOS ACUMULADOS",
        
        # Estado de Resultados
        "VENTAS NETAS INGRESOS OPERACIONALES": "INGRESOS DE ACTIVIDADES ORDINARIAS",
        "OTROS INGRESOS OPERACIONALES": "OTROS INGRESOS OPERATIVOS",
        "TOTAL DE INGRESOS BRUTOS": "INGRESOS DE ACTIVIDADES ORDINARIAS",
        "COSTO DE VENTAS": "COSTO DE VENTAS",
        "UTILIDAD BRUTA": "GANANCIA PERDIDA BRUTA",
        "GASTOS DE ADMINISTRACION": "GASTOS DE ADMINISTRACION",
        "GASTOS DE VENTAS": "GASTOS DE VENTAS Y DISTRIBUCION",
        "UTILIDAD OPERATIVA": "GANANCIA PERDIDA OPERATIVA",
        "INGRESOS FINANCIEROS": "INGRESOS FINANCIEROS",
        "GASTOS FINANCIEROS": "GASTOS FINANCIEROS",
        "OTROS INGRESOS": "OTROS INGRESOS OPERATIVOS",
        "OTROS GASTOS": "OTROS GASTOS OPERATIVOS",
        "RESULTADO POR EXPOSICION A LA INFLACION": "DIFERENCIAS DE CAMBIO NETO",
        "RESULTADOS ANTES DE PARTIDAS EXTRAORDINARIAS PARTICIPACIONES Y DEL IMPUESTO A LA RENTA": "GANANCIA PERDIDA ANTES DE IMPUESTOS",
        "PARTICIPACIONES": "PARTICIPACION EN OTRO RESULTADO INTEGRAL DE SUBSIDIARIAS ASOCIADAS Y NEGOCIOS CONJUNTOS",
        "IMPUESTO A LA RENTA": "INGRESO GASTO POR IMPUESTO",
        "RESULTADO ANTES DE PARTIDAS EXTRAORDINARIAS": "GANANCIA PERDIDA NETA DE OPERACIONES CONTINUADAS",
        "INGRESOS EXTRAORDINARIOS": "OTROS INGRESOS OPERATIVOS",
        "GASTOS EXTRAORDINARIOS": "OTROS GASTOS OPERATIVOS",
        "RESULTADO ANTES DE INTERES MINORITARIO": "GANANCIA PERDIDA NETA DEL EJERCICIO",
        "INTERES MINORITARIO": "PARTICIPACION NO CONTROLADORA",
        "UTILIDAD PERDIDA NETA DEL EJERCICIO": "GANANCIA PERDIDA NETA DEL EJERCICIO",
        "UTILIDAD PERDIDA NETA ATRIBUIBLE A LOS ACCIONISTAS": "GANANCIA PERDIDA NETA DEL EJERCICIO"
    }
    
    # Si es pre-2010 y existe mapeo exacto, usar el mapeo
    if anio < 2010:
        # Primero intentar mapeo exacto
        if cuenta in mapeo_antiguo:
            return mapeo_antiguo[cuenta]
        
        # Si no hay mapeo exacto, buscar coincidencias parciales clave
        for key_antigua, key_moderna in mapeo_antiguo.items():
            if cuenta == key_antigua:
                return key_moderna
    
    # Para cuentas modernas o sin mapeo, devolver normalizado
    return cuenta


# ------------------- Contenedores -------------------
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

    # ---------- Procesar Balance General / Estado de Situación Financiera ----------
    # Buscar tabla por id (gvReporte) o por contenido del span anterior
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

            # Lista de encabezados de sección a ignorar
            encabezados_seccion = [
                "ACTIVOS", "ACTIVO",
                "ACTIVOS CORRIENTES", "ACTIVO CORRIENTE",
                "ACTIVOS NO CORRIENTES", "ACTIVO NO CORRIENTE",
                "PASIVOS", "PASIVO",
                "PASIVOS CORRIENTES", "PASIVO CORRIENTE",
                "PASIVOS NO CORRIENTES", "PASIVO NO CORRIENTE",
                "PATRIMONIO", "PATRIMONIO NETO",
                "PASIVO Y PATRIMONIO", "PASIVOS Y PATRIMONIO",
                "CUENTAS POR COBRAR COMERCIALES Y OTRAS CUENTAS POR COBRAR",
                "CUENTAS POR PAGAR COMERCIALES Y OTRAS CUENTAS POR PAGAR"
            ]

            for fila in filas[1:]:
                if len(fila) < 3:
                    continue
                
                cuenta_raw = fila[0].strip()
                
                # Saltar si está vacío
                if not cuenta_raw:
                    continue
                
                cuenta_normalizada_temp = normalize_name(cuenta_raw)
                
                # Ignorar encabezados de sección
                if cuenta_normalizada_temp in encabezados_seccion:
                    continue
                
                # Verificar si todos los valores son 0 (encabezado sin datos)
                valores_fila = [limpiar_valor(v) for v in fila[2:]]
                if all(v == 0 for v in valores_fila):
                    continue
                
                for i, valor_str in enumerate(fila[2:]):
                    anio = anios[i]
                    if anio is None:
                        continue
                    
                    valor = limpiar_valor(valor_str)
                    
                    # Mapear cuenta según año
                    cuenta_normalizada = mapear_cuenta_normalizada(cuenta_raw, anio)
                    
                    if anio not in datos_balance:
                        datos_balance[anio] = {}
                    
                    # Lógica mejorada: priorizar valores no-cero
                    if cuenta_normalizada not in datos_balance[anio]:
                        datos_balance[anio][cuenta_normalizada] = valor
                    elif datos_balance[anio][cuenta_normalizada] == 0 and valor != 0:
                        datos_balance[anio][cuenta_normalizada] = valor
                    # Si ambos son no-cero, NO sumar (mantener el primero encontrado)
                    # Esto evita duplicados incorrectos

    # ---------- Procesar Estado de Resultados ----------
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
            
            # Para estado de resultados, puede haber múltiples columnas por año (trimestral, acumulado)
            # Priorizar columnas "Acumulado" o las últimas del año
            for col in columnas_anios:
                m = re.search(r'\b(19|20)\d{2}\b', col)
                if m:
                    anios.append(int(m.group(0)))
                else:
                    anios.append(None)

            for fila in filas[1:]:
                if len(fila) < 2:
                    continue
                
                cuenta_raw = fila[0].strip()
                
                # Ignorar filas vacías o de encabezado
                if len(fila) <= 2:
                    continue
                
                for i, valor_str in enumerate(fila[2:]):
                    anio = anios[i]
                    if anio is None:
                        continue
                    
                    valor = limpiar_valor(valor_str)
                    
                    # Mapear cuenta según año
                    cuenta_normalizada = mapear_cuenta_normalizada(cuenta_raw, anio)
                    
                    if anio not in datos_resultados:
                        datos_resultados[anio] = {}
                    
                    # Para estado de resultados: priorizar última columna del año (acumulado)
                    if cuenta_normalizada not in datos_resultados[anio]:
                        datos_resultados[anio][cuenta_normalizada] = valor
                    else:
                        # Sobrescribir con el último valor (asumiendo que es acumulado)
                        datos_resultados[anio][cuenta_normalizada] = valor

# ------------------- Crear DataFrames -------------------
if datos_balance:
    df_balance = pd.DataFrame.from_dict(datos_balance, orient='index').fillna(0.0).T
    df_balance = df_balance.reindex(sorted(df_balance.columns), axis=1)
else:
    df_balance = pd.DataFrame()

if datos_resultados:
    df_resultados = pd.DataFrame.from_dict(datos_resultados, orient='index').fillna(0.0).T
    df_resultados = df_resultados.reindex(sorted(df_resultados.columns), axis=1)
else:
    df_resultados = pd.DataFrame()

# ------------------- ANÁLISIS VERTICAL (Balance) -------------------
df_vertical = pd.DataFrame()
if not df_balance.empty:
    df_vertical = df_balance.copy()
    
    # Buscar "TOTAL DE ACTIVOS" o similar
    total_activos_row = None
    for idx in df_vertical.index:
        if "TOTAL" in idx and ("ACTIVO" in idx) and "CORRIENTE" not in idx and "NO CORRIENTE" not in idx:
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
        st.warning("⚠️ No se encontró 'TOTAL DE ACTIVOS' para análisis vertical.")
        df_vertical = pd.DataFrame()

# ------------------- ANÁLISIS HORIZONTAL (Balance) -------------------
df_horizontal_pct = pd.DataFrame()
if not df_balance.empty and len(df_balance.columns) > 1:
    df_horizontal_pct = df_balance.copy()
    columnas = df_horizontal_pct.columns.tolist()
    nuevas_columnas = []
    
    for i in range(len(columnas) - 1):
        anio_actual = columnas[i + 1]
        anio_anterior = columnas[i]
        col_nombre = f"{anio_anterior}-{anio_actual}"
        
        with pd.option_context('mode.use_inf_as_na', True):
            df_horizontal_pct[col_nombre] = ((df_horizontal_pct[anio_actual] - df_horizontal_pct[anio_anterior]) / 
                                             df_horizontal_pct[anio_anterior] * 100)
        nuevas_columnas.append(col_nombre)
    
    df_horizontal_pct = df_horizontal_pct[nuevas_columnas].round(2)
    df_horizontal_pct = df_horizontal_pct.replace([float('inf'), float('-inf')], pd.NA)

# ------------------- CÁLCULO DE RATIOS FINANCIEROS -------------------
def buscar_cuenta_flexible(df, keywords_list):
    """
    Busca una cuenta que coincida con cualquiera de las listas de keywords.
    Retorna la primera coincidencia encontrada.
    """
    for keywords in keywords_list:
        for idx in df.index:
            if all(kw.upper() in idx.upper() for kw in keywords):
                return idx
    return None

ratios_data = {}
anios_comunes = sorted(list(set(df_balance.columns) & set(df_resultados.columns))) if (not df_balance.empty and not df_resultados.empty) else []

if anios_comunes:
    for i, anio in enumerate(anios_comunes):
        ratios_data[anio] = {}

        # --- Balance General ---
        # Activo Corriente
        act_corr = buscar_cuenta_flexible(df_balance, [
            ["TOTAL", "ACTIVO", "CORRIENTE"],
            ["TOTAL", "ACTIVOS", "CORRIENTES"]
        ])
        activo_corriente = df_balance.loc[act_corr, anio] if act_corr else 0.0

        # Inventarios
        inv = buscar_cuenta_flexible(df_balance, [
            ["INVENTARIOS"],
            ["EXISTENCIAS"]
        ])
        inventarios = df_balance.loc[inv, anio] if inv else 0.0

        # Pasivo Corriente
        pas_corr = buscar_cuenta_flexible(df_balance, [
            ["TOTAL", "PASIVO", "CORRIENTE"],
            ["TOTAL", "PASIVOS", "CORRIENTES"]
        ])
        pasivo_corriente = df_balance.loc[pas_corr, anio] if pas_corr else 0.0

        # Cuentas por Cobrar (suma de todas las variantes)
        cxc_comerciales = buscar_cuenta_flexible(df_balance, [
            ["CUENTAS", "COBRAR", "COMERCIALES"],
            ["CUENTAS", "COBRAR", "COMERCIALES", "OTRAS", "CUENTAS"]
        ])
        cxc_vinculadas = buscar_cuenta_flexible(df_balance, [
            ["CUENTAS", "COBRAR", "ENTIDADES", "RELACIONADAS"],
            ["CUENTAS", "COBRAR", "VINCULADAS"]
        ])
        otras_cxc = buscar_cuenta_flexible(df_balance, [
            ["OTRAS", "CUENTAS", "COBRAR"]
        ])
        
        cxc_val = 0.0
        for cxc_idx in [cxc_comerciales, cxc_vinculadas, otras_cxc]:
            if cxc_idx and cxc_idx in df_balance.index:
                cxc_val += df_balance.loc[cxc_idx, anio]

        # Activos Totales
        act_tot = buscar_cuenta_flexible(df_balance, [
            ["TOTAL", "ACTIVO"],
            ["TOTAL", "ACTIVOS"]
        ])
        activos_totales = df_balance.loc[act_tot, anio] if act_tot else 0.0

        # Pasivo Total
        pas_tot = buscar_cuenta_flexible(df_balance, [
            ["TOTAL", "PASIVO"],
            ["TOTAL", "PASIVOS"]
        ])
        pasivo_total = df_balance.loc[pas_tot, anio] if pas_tot else 0.0

        # Patrimonio Neto
        patr = buscar_cuenta_flexible(df_balance, [
            ["TOTAL", "PATRIMONIO"],
            ["PATRIMONIO", "NETO"]
        ])
        patrimonio = df_balance.loc[patr, anio] if patr else 0.0

        # --- Estado de Resultados ---
        # Ventas
        ventas = buscar_cuenta_flexible(df_resultados, [
            ["INGRESOS", "ACTIVIDADES", "ORDINARIAS"],
            ["VENTAS", "NETAS"]
        ])
        ventas_val = df_resultados.loc[ventas, anio] if ventas else 0.0

        # Costo de Ventas
        costo = buscar_cuenta_flexible(df_resultados, [
            ["COSTO", "VENTAS"]
        ])
        costo_ventas = df_resultados.loc[costo, anio] if costo else 0.0

        # Utilidad Neta
        util = buscar_cuenta_flexible(df_resultados, [
            ["GANANCIA", "PERDIDA", "NETA", "EJERCICIO"],
            ["UTILIDAD", "PERDIDA", "NETA", "EJERCICIO"]
        ])
        utilidad_neta = df_resultados.loc[util, anio] if util else 0.0

        # --- Promedios con año anterior ---
        cxc_prom = cxc_val
        inv_prom = inventarios
        act_prom = activos_totales
        patr_prom = patrimonio

        if i > 0:
            anio_ant = anios_comunes[i-1]
            
            # CxC anterior
            cxc_ant = 0.0
            for cxc_idx in [cxc_comerciales, cxc_vinculadas, otras_cxc]:
                if cxc_idx and cxc_idx in df_balance.index:
                    cxc_ant += df_balance.loc[cxc_idx, anio_ant]
            
            # Inventarios anterior
            inv_ant = df_balance.loc[inv, anio_ant] if inv else 0.0
            
            # Activos y Patrimonio anteriores
            act_ant = df_balance.loc[act_tot, anio_ant] if act_tot else 0.0
            patr_ant = df_balance.loc[patr, anio_ant] if patr else 0.0

            cxc_prom = (cxc_val + cxc_ant) / 2 if (cxc_val + cxc_ant) != 0 else cxc_val
            inv_prom = (inventarios + inv_ant) / 2 if (inventarios + inv_ant) != 0 else inventarios
            act_prom = (activos_totales + act_ant) / 2 if (activos_totales + act_ant) != 0 else activos_totales
            patr_prom = (patrimonio + patr_ant) / 2 if (patrimonio + patr_ant) != 0 else patrimonio

        # --- Cálculo de ratios ---
        ratios_data[anio]["Liquidez Corriente"] = activo_corriente / pasivo_corriente if pasivo_corriente != 0 else None
        ratios_data[anio]["Prueba Ácida"] = (activo_corriente - inventarios) / pasivo_corriente if pasivo_corriente != 0 else None
        ratios_data[anio]["Rotación CxC"] = ventas_val / cxc_prom if cxc_prom != 0 else None
        ratios_data[anio]["Rotación Inventarios"] = abs(costo_ventas) / inv_prom if inv_prom != 0 else None
        ratios_data[anio]["Rotación Activos Totales"] = ventas_val / act_prom if act_prom != 0 else None
        ratios_data[anio]["Razón Deuda Total"] = pasivo_total / activos_totales if activos_totales != 0 else None
        ratios_data[anio]["Razón Deuda/Patrimonio"] = pasivo_total / patrimonio if patrimonio != 0 else None
        ratios_data[anio]["Margen Neto"] = utilidad_neta / ventas_val if ventas_val != 0 else None
        ratios_data[anio]["ROA"] = utilidad_neta / act_prom if act_prom != 0 else None
        ratios_data[anio]["ROE"] = utilidad_neta / patr_prom if patr_prom != 0 else None

    df_ratios = pd.DataFrame.from_dict(ratios_data, orient='index').round(4).T
else:
    df_ratios = pd.DataFrame()

# ------------------- Mostrar resultados -------------------
if not df_balance.empty:
    st.subheader("📊 Estado de Situación Financiera / Balance General")
    st.dataframe(df_balance, use_container_width=True)

if not df_resultados.empty:
    st.subheader("📊 Estado de Resultados")
    st.dataframe(df_resultados, use_container_width=True)

if not df_ratios.empty:
    st.subheader("📈 Ratios Financieros")
    st.dataframe(df_ratios, use_container_width=True)

# ------------------- Exportar a Excel -------------------
output = io.BytesIO()
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    if not df_balance.empty:
        df_balance.to_excel(writer, sheet_name='ESTADO SITUACION FINANCIERA', index_label='Cuenta')
    if not df_resultados.empty:
        df_resultados.to_excel(writer, sheet_name='ESTADO DE RESULTADOS', index_label='Cuenta')
    if not df_vertical.empty:
        df_vertical.to_excel(writer, sheet_name='ANALISIS_VERTICAL_BALANCE', index_label='Cuenta')
    if not df_horizontal_pct.empty:
        df_horizontal_pct.to_excel(writer, sheet_name='ANALISIS_HORIZONTAL_BALANCE', index_label='Cuenta')
    if not df_ratios.empty:
        df_ratios.to_excel(writer, sheet_name='RATIOS_FINANCIEROS', index_label='Ratio')

st.download_button(
    label="📥 Descargar Excel Consolidado (Completo)",
    data=output.getvalue(),
    file_name="Consolidado_Estados_Financieros_Completo.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.success("✅ ¡Proceso completado! El archivo incluye estados financieros, análisis V/H y ratios financieros.")