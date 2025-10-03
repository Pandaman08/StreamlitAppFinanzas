import pandas as pd
from utils import buscar_cuenta_flexible, buscar_cuenta_parcial

def calcular_analisis_vh(df_balance, df_resultados):
    """Calcula análisis vertical y horizontal para balance y resultados."""
    df_vertical_balance = pd.DataFrame()
    df_horizontal_balance = pd.DataFrame()
    if not df_balance.empty:
        df_vertical_balance = df_balance.copy()
        total_activos_row = None
        for idx in df_vertical_balance.index:
            if "TOTAL" in idx and ("ACTIVO" in idx) and "CORRIENTE" not in idx and "NO CORRIENTE" not in idx:
                total_activos_row = idx
                break
        if total_activos_row:
            total_activos = df_vertical_balance.loc[total_activos_row]
            for col in df_vertical_balance.columns:
                if total_activos[col] != 0:
                    df_vertical_balance[col] = (df_vertical_balance[col] / total_activos[col]) * 100
            df_vertical_balance = df_vertical_balance.round(2)
        df_horizontal_balance = df_balance.copy()
        columnas = df_horizontal_balance.columns.tolist()
        nuevas_columnas = []
        for i in range(len(columnas) - 1):
            anio_actual = columnas[i + 1]
            anio_anterior = columnas[i]
            col_nombre = f"{anio_anterior}-{anio_actual}"
            with pd.option_context('mode.use_inf_as_na', True):
                df_horizontal_balance[col_nombre] = ((df_horizontal_balance[anio_actual] - df_horizontal_balance[anio_anterior]) / 
                                                     df_horizontal_balance[anio_anterior] * 100)
            nuevas_columnas.append(col_nombre)
        df_horizontal_balance = df_horizontal_balance[nuevas_columnas].round(2)
        df_horizontal_balance = df_horizontal_balance.replace([float('inf'), float('-inf')], pd.NA)

    df_vertical_resultados = pd.DataFrame()
    df_horizontal_resultados = pd.DataFrame()
    if not df_resultados.empty:
        df_vertical_resultados = df_resultados.copy()
        ventas_row = buscar_cuenta_flexible(df_resultados, [
            ["INGRESOS", "ACTIVIDADES", "ORDINARIAS"],
            ["VENTAS", "NETAS"]
        ])
        if ventas_row:
            ventas = df_vertical_resultados.loc[ventas_row]
            for col in df_vertical_resultados.columns:
                if ventas[col] != 0:
                    df_vertical_resultados[col] = (df_vertical_resultados[col] / ventas[col]) * 100
            df_vertical_resultados = df_vertical_resultados.round(2)
        df_horizontal_resultados = df_resultados.copy()
        columnas = df_horizontal_resultados.columns.tolist()
        nuevas_columnas = []
        for i in range(len(columnas) - 1):
            anio_actual = columnas[i + 1]
            anio_anterior = columnas[i]
            col_nombre = f"{anio_anterior}-{anio_actual}"
            with pd.option_context('mode.use_inf_as_na', True):
                df_horizontal_resultados[col_nombre] = ((df_horizontal_resultados[anio_actual] - df_horizontal_resultados[anio_anterior]) / 
                                                        df_horizontal_resultados[anio_anterior] * 100)
            nuevas_columnas.append(col_nombre)
        df_horizontal_resultados = df_horizontal_resultados[nuevas_columnas].round(2)
        df_horizontal_resultados = df_horizontal_resultados.replace([float('inf'), float('-inf')], pd.NA)

    return df_vertical_balance, df_horizontal_balance, df_vertical_resultados, df_horizontal_resultados

def calcular_ratios(df_balance, df_resultados):
    """Calcula ratios financieros."""
    ratios_data = {}
    debug_info = {}
    anios_comunes = sorted(list(set(df_balance.columns) & set(df_resultados.columns))) if (not df_balance.empty and not df_resultados.empty) else []

    if anios_comunes:
        for i, anio in enumerate(anios_comunes):
            ratios_data[anio] = {}
            debug_info[anio] = {}
            # Activo Corriente
            act_corr = buscar_cuenta_flexible(df_balance, [
                ["TOTAL", "ACTIVO", "CORRIENTE"],
                ["TOTAL", "ACTIVOS", "CORRIENTES"]
            ])
            activo_corriente = df_balance.loc[act_corr, anio] if act_corr in df_balance.index else 0.0
            debug_info[anio]["activo_corriente"] = f"{act_corr} = {activo_corriente}"
            # Inventarios
            inv = buscar_cuenta_flexible(df_balance, [
                ["INVENTARIOS"],
                ["EXISTENCIAS"]
            ])
            if not inv:
                inv = buscar_cuenta_parcial(df_balance, ["INVENTARIO", "EXISTENCIA"]) if not df_balance.empty else None
            inventarios = df_balance.loc[inv, anio] if inv in df_balance.index else 0.0
            debug_info[anio]["inventarios"] = f"{inv} = {inventarios}"
            # Pasivo Corriente
            pas_corr = buscar_cuenta_flexible(df_balance, [
                ["TOTAL", "PASIVO", "CORRIENTE"],
                ["TOTAL", "PASIVOS", "CORRIENTES"]
            ])
            pasivo_corriente = df_balance.loc[pas_corr, anio] if pas_corr in df_balance.index else 0.0
            debug_info[anio]["pasivo_corriente"] = f"{pas_corr} = {pasivo_corriente}"
            # Cuentas por Cobrar
            cxc_comerciales = buscar_cuenta_flexible(df_balance, [
                ["CUENTAS", "COBRAR", "COMERCIALES"]
            ])
            if not cxc_comerciales:
                cxc_comerciales = buscar_cuenta_parcial(df_balance, ["CUENTAS", "COBRAR", "COMERCIAL"]) if not df_balance.empty else None
            cxc_vinculadas = buscar_cuenta_flexible(df_balance, [
                ["CUENTAS", "COBRAR", "ENTIDADES", "RELACIONADAS"],
                ["CUENTAS", "COBRAR", "VINCULADAS"]
            ])
            if not cxc_vinculadas:
                cxc_vinculadas = buscar_cuenta_parcial(df_balance, ["CUENTAS", "COBRAR", "VINCULADA"]) if not df_balance.empty else None
            otras_cxc = buscar_cuenta_flexible(df_balance, [
                ["OTRAS", "CUENTAS", "COBRAR"]
            ])
            cxc_val = 0.0
            for cxc_idx in [cxc_comerciales, cxc_vinculadas, otras_cxc]:
                if cxc_idx and cxc_idx in df_balance.index:
                    cxc_val += df_balance.loc[cxc_idx, anio]
            debug_info[anio]["cxc"] = f"com:{cxc_comerciales}, vinc:{cxc_vinculadas}, otras:{otras_cxc} = {cxc_val}"
            # Activos Totales
            act_tot = buscar_cuenta_flexible(df_balance, [
                ["TOTAL", "ACTIVO"],
                ["TOTAL", "ACTIVOS"]
            ])
            activos_totales = df_balance.loc[act_tot, anio] if act_tot in df_balance.index else 0.0
            debug_info[anio]["activos_totales"] = f"{act_tot} = {activos_totales}"
            # Pasivo Total
            pas_tot = buscar_cuenta_flexible(df_balance, [
                ["TOTAL", "PASIVO"],
                ["TOTAL", "PASIVOS"]
            ])
            pasivo_total = df_balance.loc[pas_tot, anio] if pas_tot in df_balance.index else 0.0
            # Patrimonio
            patr = buscar_cuenta_flexible(df_balance, [
                ["TOTAL", "PATRIMONIO"],
                ["PATRIMONIO", "NETO"],
                ["TOTAL", "PATRIMONIO", "NETO"]
            ])
            if not patr:
                patr = buscar_cuenta_parcial(df_balance, ["PATRIMONIO"]) if not df_balance.empty else None
            patrimonio = df_balance.loc[patr, anio] if patr in df_balance.index else 0.0
            if patrimonio == 0.0 and activos_totales != 0.0:
                patrimonio = activos_totales - pasivo_total
            debug_info[anio]["patrimonio"] = f"{patr} = {patrimonio}"
            # Ventas
            ventas = buscar_cuenta_flexible(df_resultados, [
                ["INGRESOS", "ACTIVIDADES", "ORDINARIAS"]
            ])
            if not ventas:
                ventas = buscar_cuenta_parcial(df_resultados, ["INGRESOS", "ACTIVIDADES"]) if not df_resultados.empty else None
            if not ventas:
                ventas = buscar_cuenta_parcial(df_resultados, ["VENTAS", "NETAS"]) if not df_resultados.empty else None
            if not ventas:
                ventas = buscar_cuenta_parcial(df_resultados, ["INGRESOS", "OPERACIONALES"]) if not df_resultados.empty else None
            ventas_val = df_resultados.loc[ventas, anio] if ventas in df_resultados.index else 0.0
            debug_info[anio]["ventas"] = f"{ventas} = {ventas_val}"
            # Costo de Ventas
            costo = buscar_cuenta_flexible(df_resultados, [
                ["COSTO", "VENTAS"]
            ])
            if not costo:
                costo = buscar_cuenta_parcial(df_resultados, ["COSTO", "VENTA"]) if not df_resultados.empty else None
            costo_ventas = df_resultados.loc[costo, anio] if costo in df_resultados.index else 0.0
            debug_info[anio]["costo_ventas"] = f"{costo} = {costo_ventas}"
            # Utilidad Neta
            util = buscar_cuenta_flexible(df_resultados, [
                ["GANANCIA", "PERDIDA", "NETA", "EJERCICIO"]
            ])
            if not util:
                util = buscar_cuenta_parcial(df_resultados, ["GANANCIA", "NETA", "EJERCICIO"]) if not df_resultados.empty else None
            if not util:
                util = buscar_cuenta_parcial(df_resultados, ["UTILIDAD", "NETA", "EJERCICIO"]) if not df_resultados.empty else None
            if not util and not df_resultados.empty:
                for idx in df_resultados.index:
                    if "UTILIDAD" in idx and "EJERCICIO" in idx and "NETA" in idx:
                        util = idx
                        break
            utilidad_neta = df_resultados.loc[util, anio] if util in df_resultados.index else 0.0
            debug_info[anio]["utilidad_neta"] = f"{util} = {utilidad_neta}"
            # Promedios
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
                # Inventarios anteriores
                inv_ant = df_balance.loc[inv, anio_ant] if inv in df_balance.index else 0.0
                act_ant = df_balance.loc[act_tot, anio_ant] if act_tot in df_balance.index else 0.0
                # Patrimonio anterior
                patr_ant = df_balance.loc[patr, anio_ant] if patr in df_balance.index else 0.0
                if patr_ant == 0.0 and act_ant != 0.0:
                    pas_ant = df_balance.loc[pas_tot, anio_ant] if pas_tot in df_balance.index else 0.0
                    patr_ant = act_ant - pas_ant
                # Promedios con control
                cxc_prom = (cxc_val + cxc_ant) / 2 if (cxc_val + cxc_ant) != 0 else "N/A"
                inv_prom = (inventarios + inv_ant) / 2 if (inventarios + inv_ant) != 0 else "N/A"
                act_prom = (activos_totales + act_ant) / 2 if (activos_totales + act_ant) != 0 else "N/A"
                patr_prom = (patrimonio + patr_ant) / 2 if (patrimonio + patr_ant) != 0 else "N/A"
            # Calcular ratios con "N/A" si no se puede
            def safe_div(num, den):
                if isinstance(den, (int, float)) and den != 0:
                    return num / den
                return "N/A"
            ratios_data[anio]["Liquidez Corriente"] = safe_div(activo_corriente, pasivo_corriente)
            ratios_data[anio]["Prueba Ácida"] = safe_div((activo_corriente - inventarios), pasivo_corriente)
            ratios_data[anio]["Rotación CxC"] = safe_div(ventas_val, cxc_prom) if not isinstance(cxc_prom, str) else "N/A"
            ratios_data[anio]["Rotación Inventarios"] = safe_div(abs(costo_ventas), inv_prom) if not isinstance(inv_prom, str) else "N/A"
            ratios_data[anio]["Rotación Activos Totales"] = safe_div(ventas_val, act_prom) if not isinstance(act_prom, str) else "N/A"
            ratios_data[anio]["Razón Deuda Total"] = safe_div(pasivo_total, activos_totales)
            ratios_data[anio]["Razón Deuda/Patrimonio"] = safe_div(pasivo_total, patrimonio)
            ratios_data[anio]["Margen Neto"] = safe_div(utilidad_neta, ventas_val)
            ratios_data[anio]["ROA"] = safe_div(utilidad_neta, act_prom) if not isinstance(act_prom, str) else "N/A"
            ratios_data[anio]["ROE"] = safe_div(utilidad_neta, patr_prom) if not isinstance(patr_prom, str) else "N/A"

    # Crear DataFrame de ratios y redondear sólo celdas numéricas
    if ratios_data:
        df_ratios = pd.DataFrame.from_dict(ratios_data, orient='index')
        def round_if_num(x):
            return round(x, 4) if isinstance(x, (int, float)) else x
        df_ratios = df_ratios.applymap(round_if_num).T
    else:
        df_ratios = pd.DataFrame()

    return df_ratios, debug_info, anios_comunes