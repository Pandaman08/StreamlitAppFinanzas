import pandas as pd
from bs4 import BeautifulSoup
import re
from utils import normalize_name, limpiar_valor, mapear_cuenta_normalizada


def _extraer_anios_desde_encabezados(encabezados):
    anios = []
    for col in encabezados[2:]:
        m = re.search(r'\b(19|20)\d{2}\b', str(col))
        if m:
            anios.append(int(m.group(0)))
        else:
            anios.append(None)
    return anios


def _obtener_filas_html(contenido, table_id):
    soup = BeautifulSoup(contenido, 'html.parser')
    tabla = soup.find('table', {'id': table_id})
    if not tabla:
        return []

    filas = []
    for tr in tabla.find_all('tr'):
        celdas = [td.get_text(strip=True) for td in tr.find_all(['td', 'th'])]
        if celdas:
            filas.append(celdas)
    return filas


def _normalizar_filas_excel(df):
    filas = []
    for row in df.itertuples(index=False, name=None):
        celdas = []
        for valor in row:
            if pd.isna(valor):
                celdas.append("")
            else:
                celdas.append(str(valor).strip())

        while celdas and celdas[-1] == "":
            celdas.pop()

        if any(celda != "" for celda in celdas):
            filas.append(celdas)
    return filas


def _detectar_tipo_hoja(nombre_hoja, filas):
    nombre = normalize_name(nombre_hoja or "")
    muestra = " ".join(
        " ".join(fila[:2]) for fila in filas[:20] if fila
    )
    muestra = normalize_name(muestra)

    claves = f"{nombre} {muestra}"

    if any(k in claves for k in ["FLUJO", "EFECTIVO", "CASH FLOW"]):
        return "flujo"
    if any(k in claves for k in ["RESULTADO", "RESULTADOS", "GANANCIA", "PERDIDA"]):
        return "resultados"
    if any(k in claves for k in ["SITUACION FINANCIERA", "BALANCE", "ACTIVO", "PASIVO", "PATRIMONIO"]):
        return "balance"
    return None


def _iterar_cuentas_anuales(filas):
    if len(filas) <= 1:
        return

    anios = _extraer_anios_desde_encabezados(filas[0])
    for fila in filas[1:]:
        if len(fila) <= 2:
            continue

        cuenta_raw = str(fila[0]).strip()
        if not cuenta_raw:
            continue

        for i_col, valor_str in enumerate(fila[2:]):
            if i_col >= len(anios):
                continue
            anio = anios[i_col]
            if anio is None:
                continue
            yield fila, cuenta_raw, anio, valor_str


def _asignar_valor(destino, anio, cuenta, valor, sobrescribir=False):
    if anio not in destino:
        destino[anio] = {}

    if sobrescribir or cuenta not in destino[anio]:
        destino[anio][cuenta] = valor
        return

    if destino[anio][cuenta] == 0 and valor != 0:
        destino[anio][cuenta] = valor


def _fila_balance_valida(fila, encabezados_seccion):
    cuenta_raw = str(fila[0]).strip()
    if not cuenta_raw:
        return False

    if normalize_name(cuenta_raw) in encabezados_seccion:
        return False

    valores_fila = [limpiar_valor(v) for v in fila[2:]]
    return not all(v == 0 for v in valores_fila)


def _procesar_filas_balance(filas, datos_balance):
    encabezados_seccion = [
        "ACTIVOS", "ACTIVO", "ACTIVOS CORRIENTES", "ACTIVO CORRIENTE",
        "ACTIVOS NO CORRIENTES", "ACTIVO NO CORRIENTE",
        "PASIVOS", "PASIVO", "PASIVOS CORRIENTES", "PASIVO CORRIENTE",
        "PASIVOS NO CORRIENTES", "PASIVO NO CORRIENTE",
        "PATRIMONIO", "PATRIMONIO NETO", "PASIVO Y PATRIMONIO", "PASIVOS Y PATRIMONIO",
        "CUENTAS POR COBRAR COMERCIALES Y OTRAS CUENTAS POR COBRAR",
        "CUENTAS POR PAGAR COMERCIALES Y OTRAS CUENTAS POR PAGAR"
    ]

    for fila, cuenta_raw, anio, valor_str in _iterar_cuentas_anuales(filas):
        if not _fila_balance_valida(fila, encabezados_seccion):
            continue

        valor = limpiar_valor(valor_str)
        cuenta_normalizada = mapear_cuenta_normalizada(cuenta_raw, anio)
        _asignar_valor(datos_balance, anio, cuenta_normalizada, valor, sobrescribir=False)


def _procesar_filas_resultados(filas, datos_resultados):
    for _, cuenta_raw, anio, valor_str in _iterar_cuentas_anuales(filas):
        valor = limpiar_valor(valor_str)
        cuenta_normalizada = mapear_cuenta_normalizada(cuenta_raw, anio)
        _asignar_valor(datos_resultados, anio, cuenta_normalizada, valor, sobrescribir=True)


def _procesar_filas_flujo(filas, datos_flujo_efectivo):
    for _, cuenta_raw, anio, valor_str in _iterar_cuentas_anuales(filas):
        valor = limpiar_valor(valor_str)
        cuenta_normalizada = mapear_cuenta_normalizada(cuenta_raw, anio)
        _asignar_valor(datos_flujo_efectivo, anio, cuenta_normalizada, valor, sobrescribir=False)


def _parsear_excel(archivo):
    archivo.seek(0)
    xls = pd.ExcelFile(archivo)

    filas_balance = []
    filas_resultados = []
    filas_flujo = []

    for nombre_hoja in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=nombre_hoja, header=None)
        filas = _normalizar_filas_excel(df)
        if not filas:
            continue

        tipo = _detectar_tipo_hoja(nombre_hoja, filas)
        if tipo == "balance" and not filas_balance:
            filas_balance = filas
        elif tipo == "resultados" and not filas_resultados:
            filas_resultados = filas
        elif tipo == "flujo" and not filas_flujo:
            filas_flujo = filas

    return filas_balance, filas_resultados, filas_flujo


def _parsear_html(archivo):
    contenido = None
    for cod in ['latin-1', 'cp1252', 'utf-8']:
        try:
            archivo.seek(0)
            contenido = archivo.read().decode(cod)
            break
        except Exception:
            continue

    if not contenido:
        return [], [], []

    filas_balance = _obtener_filas_html(contenido, 'gvReporte')
    filas_resultados = _obtener_filas_html(contenido, 'gvReporte1')
    filas_flujo = _obtener_filas_html(contenido, 'gvReporte3')
    return filas_balance, filas_resultados, filas_flujo

def procesar_archivos(archivos):
    """Procesa los archivos subidos y devuelve datos de balance, resultados y flujo de efectivo."""
    datos_balance = {}
    datos_resultados = {}
    datos_flujo_efectivo = {}

    for archivo in archivos:
        nombre = (getattr(archivo, 'name', '') or '').lower()

        filas_balance = []
        filas_resultados = []
        filas_flujo = []

        # .xlsx/.xlsm se procesan como Excel nativo.
        # .xls se intenta como HTML del SMV y, si falla, como Excel binario.
        if nombre.endswith(('.xlsx', '.xlsm', '.xlsb')):
            try:
                filas_balance, filas_resultados, filas_flujo = _parsear_excel(archivo)
            except Exception:
                filas_balance, filas_resultados, filas_flujo = _parsear_html(archivo)
        elif nombre.endswith('.xls'):
            filas_balance, filas_resultados, filas_flujo = _parsear_html(archivo)
            if not filas_balance and not filas_resultados and not filas_flujo:
                try:
                    filas_balance, filas_resultados, filas_flujo = _parsear_excel(archivo)
                except Exception:
                    continue
        else:
            filas_balance, filas_resultados, filas_flujo = _parsear_html(archivo)

        _procesar_filas_balance(filas_balance, datos_balance)
        _procesar_filas_resultados(filas_resultados, datos_resultados)
        _procesar_filas_flujo(filas_flujo, datos_flujo_efectivo)

    df_balance = pd.DataFrame.from_dict(datos_balance, orient='index').fillna(0.0).T if datos_balance else pd.DataFrame()
    df_resultados = pd.DataFrame.from_dict(datos_resultados, orient='index').fillna(0.0).T if datos_resultados else pd.DataFrame()
    df_flujo_efectivo = pd.DataFrame.from_dict(datos_flujo_efectivo, orient='index').fillna(0.0).T if datos_flujo_efectivo else pd.DataFrame()

    # ⭐️ ELIMINAR EL PRIMER AÑO DE TODAS LAS TABLAS (LOGICA DE TU COMPAÑERO)
    if not df_balance.empty:
        df_balance = df_balance.reindex(sorted(df_balance.columns), axis=1)
        if len(df_balance.columns) > 1:
            df_balance = df_balance.iloc[:, 1:]  # ← Elimina la primera columna (primer año)

    if not df_resultados.empty:
        df_resultados = df_resultados.reindex(sorted(df_resultados.columns), axis=1)
        if len(df_resultados.columns) > 1:
            df_resultados = df_resultados.iloc[:, 1:]  # ← Elimina la primera columna (primer año)

    if not df_flujo_efectivo.empty:
        df_flujo_efectivo = df_flujo_efectivo.reindex(sorted(df_flujo_efectivo.columns), axis=1)
        if len(df_flujo_efectivo.columns) > 1:
            df_flujo_efectivo = df_flujo_efectivo.iloc[:, 1:]  # ← Elimina la primera columna (primer año)

    return df_balance, df_resultados, df_flujo_efectivo