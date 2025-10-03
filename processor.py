import pandas as pd
from bs4 import BeautifulSoup
import re
from utils import normalize_name, limpiar_valor, mapear_cuenta_normalizada

def procesar_archivos(archivos):
    """Procesa los archivos subidos y devuelve datos de balance, resultados y flujo de efectivo."""
    datos_balance = {}
    datos_resultados = {}
    datos_flujo_efectivo = {}

    for i, archivo in enumerate(archivos):
        contenido = None
        for cod in ['latin-1', 'cp1252', 'utf-8']:
            try:
                archivo.seek(0)
                contenido = archivo.read().decode(cod)
                break
            except:
                continue
        if not contenido:
            continue

        soup = BeautifulSoup(contenido, 'html.parser')

        # Procesar Balance
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
                encabezados_seccion = [
                    "ACTIVOS", "ACTIVO", "ACTIVOS CORRIENTES", "ACTIVO CORRIENTE",
                    "ACTIVOS NO CORRIENTES", "ACTIVO NO CORRIENTE",
                    "PASIVOS", "PASIVO", "PASIVOS CORRIENTES", "PASIVO CORRIENTE",
                    "PASIVOS NO CORRIENTES", "PASIVO NO CORRIENTE",
                    "PATRIMONIO", "PATRIMONIO NETO", "PASIVO Y PATRIMONIO", "PASIVOS Y PATRIMONIO",
                    "CUENTAS POR COBRAR COMERCIALES Y OTRAS CUENTAS POR COBRAR",
                    "CUENTAS POR PAGAR COMERCIALES Y OTRAS CUENTAS POR PAGAR"
                ]
                for fila in filas[1:]:
                    if len(fila) < 3:
                        continue
                    cuenta_raw = fila[0].strip()
                    if not cuenta_raw:
                        continue
                    cuenta_normalizada_temp = normalize_name(cuenta_raw)
                    if cuenta_normalizada_temp in encabezados_seccion:
                        continue
                    valores_fila = [limpiar_valor(v) for v in fila[2:]]
                    if all(v == 0 for v in valores_fila):
                        continue
                    for i_col, valor_str in enumerate(fila[2:]):
                        anio = anios[i_col]
                        if anio is None:
                            continue
                        valor = limpiar_valor(valor_str)
                        cuenta_normalizada = mapear_cuenta_normalizada(cuenta_raw, anio)
                        if anio not in datos_balance:
                            datos_balance[anio] = {}
                        if cuenta_normalizada not in datos_balance[anio]:
                            datos_balance[anio][cuenta_normalizada] = valor
                        elif datos_balance[anio][cuenta_normalizada] == 0 and valor != 0:
                            datos_balance[anio][cuenta_normalizada] = valor

        # Procesar Estado de Resultados
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
                    cuenta_raw = fila[0].strip()
                    if len(fila) <= 2:
                        continue
                    for i_col, valor_str in enumerate(fila[2:]):
                        anio = anios[i_col]
                        if anio is None:
                            continue
                        valor = limpiar_valor(valor_str)
                        cuenta_normalizada = mapear_cuenta_normalizada(cuenta_raw, anio)
                        if anio not in datos_resultados:
                            datos_resultados[anio] = {}
                        datos_resultados[anio][cuenta_normalizada] = valor

        # Procesar Flujo de Efectivo
        tabla_flujo = soup.find('table', {'id': 'gvReporte3'})
        if tabla_flujo:
            filas = []
            for tr in tabla_flujo.find_all('tr'):
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
                    cuenta_raw = fila[0].strip()
                    for i_col, valor_str in enumerate(fila[2:]):
                        anio = anios[i_col]
                        if anio is None:
                            continue
                        valor = limpiar_valor(valor_str)
                        cuenta_normalizada = mapear_cuenta_normalizada(cuenta_raw, anio)
                        if anio not in datos_flujo_efectivo:
                            datos_flujo_efectivo[anio] = {}
                        if cuenta_normalizada not in datos_flujo_efectivo[anio]:
                            datos_flujo_efectivo[anio][cuenta_normalizada] = valor
                        elif datos_flujo_efectivo[anio][cuenta_normalizada] == 0 and valor != 0:
                            datos_flujo_efectivo[anio][cuenta_normalizada] = valor

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