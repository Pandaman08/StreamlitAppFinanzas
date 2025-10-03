import unicodedata
import re

def normalize_name(s):
    """Normaliza textos: quita tildes, mayúsculas, compacta espacios y elimina notas (9)"""
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
    if valor.startswith('(') and valor.endswith(')'):
        valor = '-' + valor[1:-1]
    try:
        return float(valor)
    except:
        return 0.0

def mapear_cuenta_normalizada(cuenta_original, anio):
    """Mapea nombres de cuentas antiguas (pre-2010) a nomenclatura moderna."""
    cuenta = normalize_name(cuenta_original)
    mapeo_antiguo = {
        # Balance - Activos
        "CAJA Y BANCOS": "EFECTIVO Y EQUIVALENTES AL EFECTIVO",
        "VALORES NEGOCIABLES": "OTROS ACTIVOS FINANCIEROS",
        "EXISTENCIAS": "INVENTARIOS",
        "GASTOS PAGADOS POR ANTICIPADO": "ANTICIPOS",
        "INVERSIONES PERMANENTES": "INVERSIONES EN SUBSIDIARIAS NEGOCIOS CONJUNTOS Y ASOCIADAS",
        "INMUEBLES MAQUINARIA Y EQUIPO NETO DE DEPRECIACION ACUMULADA": "PROPIEDADES PLANTA Y EQUIPO",
        "INMUEBLES MAQUINARIA Y EQUIPO": "PROPIEDADES PLANTA Y EQUIPO",
        "ACTIVO INTANGIBLE NETO DE DEPRECIACION ACUMULADA": "ACTIVOS INTANGIBLES DISTINTOS DE LA PLUSVALIA",
        "ACTIVOS INTANGIBLES": "ACTIVOS INTANGIBLES DISTINTOS DE LA PLUSVALIA",
        "OTROS ACTIVOS": "OTROS ACTIVOS NO FINANCIEROS",
        "IMPUESTO A LA RENTA Y PARTICIPACIONES DIFERIDOS ACTIVO": "ACTIVOS POR IMPUESTOS DIFERIDOS",
        # Balance - Pasivos
        "SOBREGIROS Y PAGARES BANCARIOS": "OTROS PASIVOS FINANCIEROS",
        "PARTE CORRIENTE DE LAS DEUDAS A LARGO PLAZO": "OTROS PASIVOS FINANCIEROS",
        "DEUDAS A LARGO PLAZO": "OTROS PASIVOS FINANCIEROS",
        "INGRESOS DIFERIDOS": "INGRESOS DIFERIDOS",
        "IMPUESTO A LA RENTA Y PARTICIPACIONES DIFERIDOS PASIVO": "PASIVOS POR IMPUESTOS DIFERIDOS",
        # Balance - Patrimonio
        "CAPITAL": "CAPITAL EMITIDO",
        "CAPITAL ADICIONAL": "PRIMAS DE EMISION",
        "EXCEDENTE DE REVALUACION": "SUPERAVIT DE REVALUACION",
        "RESERVAS LEGALES": "OTRAS RESERVAS DE CAPITAL",
        "OTRAS RESERVAS": "OTRAS RESERVAS DE PATRIMONIO",
        "RESULTADOS ACUMULADOS": "RESULTADOS ACUMULADOS",
        # Estado de Resultados
        "VENTAS NETAS INGRESOS OPERACIONALES": "INGRESOS DE ACTIVIDADES ORDINARIAS",
        "VENTAS NETAS": "INGRESOS DE ACTIVIDADES ORDINARIAS",
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
        "PARTICIPACIONES": "OTROS INGRESOS GASTOS DE LAS SUBSIDIARIAS ASOCIADAS Y NEGOCIOS CONJUNTOS",
        "IMPUESTO A LA RENTA": "INGRESO GASTO POR IMPUESTO",
        "RESULTADO ANTES DE PARTIDAS EXTRAORDINARIAS": "GANANCIA PERDIDA NETA DE OPERACIONES CONTINUADAS",
        "INGRESOS EXTRAORDINARIOS": "OTROS INGRESOS OPERATIVOS",
        "GASTOS EXTRAORDINARIOS": "OTROS GASTOS OPERATIVOS",
        "RESULTADO ANTES DE INTERES MINORITARIO": "GANANCIA PERDIDA NETA DEL EJERCICIO",
        "INTERES MINORITARIO": "PARTICIPACION NO CONTROLADORA",
        "UTILIDAD PERDIDA NETA DEL EJERCICIO": "GANANCIA PERDIDA NETA DEL EJERCICIO",
        "UTILIDAD NETA DEL EJERCICIO": "GANANCIA PERDIDA NETA DEL EJERCICIO",
        "UTILIDAD PERDIDA NETA ATRIBUIBLE A LOS ACCIONISTAS": "GANANCIA PERDIDA NETA DEL EJERCICIO"
    }
    if anio < 2010:
        if cuenta in mapeo_antiguo:
            return mapeo_antiguo[cuenta]
        # Búsqueda flexible
        if "VENTAS" in cuenta and "NETAS" in cuenta:
            return "INGRESOS DE ACTIVIDADES ORDINARIAS"
        if "UTILIDAD" in cuenta and "NETA" in cuenta and "EJERCICIO" in cuenta:
            return "GANANCIA PERDIDA NETA DEL EJERCICIO"
        if "EXISTENCIAS" in cuenta:
            return "INVENTARIOS"
    return cuenta

def buscar_cuenta_flexible(df, keywords_list):
    """Busca una cuenta que coincida con cualquiera de las listas de keywords."""
    for keywords in keywords_list:
        for idx in df.index:
            if all(kw.upper() in idx.upper() for kw in keywords):
                return idx
    return None

def buscar_cuenta_parcial(df, keywords):
    """Búsqueda con coincidencia parcial (al menos una palabra clave)."""
    for idx in df.index:
        if any(kw.upper() in idx.upper() for kw in keywords):
            return idx
    return None