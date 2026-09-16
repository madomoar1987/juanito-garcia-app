#!/usr/bin/env python3
"""
JUANITO — Serie histórica mensual desde Power BI
Extrae el valor de cada medida confirmada mes a mes desde Ene-2025 hasta el mes actual.
Salida: data/latest/series.json

Estructura:
{
  "generado": "2026-09-02",
  "periodos": ["2025-01", "2025-02", ...],
  "datasets": {
     "cxc": { "% Morosidad": [0.14, 0.15, ...], ... },
     ...
  }
}

Todos los valores provienen exclusivamente de la API REST de Power BI (executeQueries).
"""

import os, json, re, time, datetime, requests
from pathlib import Path

TENANT_ID     = os.environ["AZURE_TENANT_ID"]
CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
USERNAME      = os.environ["PBI_USERNAME"]
PASSWORD      = os.environ["PBI_PASSWORD"]

PBI_SCOPE = "https://analysis.windows.net/powerbi/api/.default"
TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
PBI_BASE  = "https://api.powerbi.com/v1.0/myorg"
WS_ID     = "461932ad-b5ec-4fd6-aa97-f1fc7bdc5169"

# El reporte FILLRATE vive en OTRO workspace (visto en su URL el 2026-09-06),
# no en el que usan los otros diez. Sin esto sus consultas fallan con 404.
WS_FILLRATE = "dba94185-3df0-4310-b6b4-9c3b7bc30104"
FILLRATE_REPORT_ID = "9641c6b7-2524-4774-b3be-b0370687cd3d"

# dataset_id -> workspace, para los que no están en WS_ID
WS_POR_DATASET = {}

OUTPUT_DIR = Path("data/latest")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

DATASET_IDS = {
    "cxc":            "2eec70cd-0820-408f-b938-a2cd547b0c18",
    "cxp":            "45a8ab8d-e162-4398-a260-3a9a5f90829f",
    "margen":         "38076daa-d2cd-4a93-858a-82c0a4cf8cb6",
    "mermas":         "fdd58a2d-654d-41e0-b587-3d43c337ed47",
    "compras":        "06408938-8202-424e-80c0-b42c178dabde",
    "inventario":     "0e27d784-41a4-48f0-9208-60210119f0a7",
    "control_ds":     "0aca7bdd-6b72-49c2-be41-17ae0f6b5848",
    "planificacion":  "30074d92-7ec1-4762-82f2-1cb29c15dcfe",
    "consumo":        "c972c8cb-e5fc-4b60-8f5e-265a78e1e796",
    "fill_rate":      "7f4ebe22-5e90-4e35-973b-4af3c58497e5",
    # 13. Productividad por funcionario. No estaba registrado aquí, así que
    # este reporte nunca tuvo serie mensual pese a tener dataset propio.
    "productividad":  "a042ba6a-c82c-4fc1-bf95-b9e84bd15fc6",
}

# Pares (tabla, columna) de fecha a probar, en orden de probabilidad
DATE_CANDIDATES = [
    ("Calendario", "Date"),
    ("Calendario", "Fecha"),
    ("Calendario", "FECHA"),
    ("Fecha", "Date"),
    ("Fecha", "Fecha"),
    ("Calendar", "Date"),
    ("DimCalendario", "Date"),
    ("Dim Calendario", "Date"),
    ("dCalendario", "Date"),
    ("Tiempo", "Fecha"),
]

INICIO_ANIO = 2025
INICIO_MES  = 1

# Medidas de la sonda genérica que ya tienen equivalente capturado y por tanto
# NO deben publicarse: mostrarían una segunda cifra del mismo concepto,
# calculada sin los filtros del reporte.
#
# 'Venta Total' vs 'Ventas (S/.)': la primera venía sin ninguno de los 12
# filtros del visual y daba ~S/350K más al mes (bonificaciones, chatarra y
# servicios que el reporte excluye). Tener las dos en el dashboard es
# justamente el problema que veníamos corrigiendo.
SUSTITUIDAS_POR_CAPTURA = {
    ("margen", "Venta Total"),
    # Ambas contradicen los KPIs capturados del mismo reporte:
    #   · 'Refinanciamiento' trae variaciones mensuales — con meses negativos —
    #     junto a un KPI de saldo refinanciado de S/6.96M. Son cosas distintas
    #     con el mismo nombre.
    #   · 'Proveedores' va acumulando (1, 4, 10, … 121) mientras el KPI dice
    #     138 proveedores. No es un conteo mensual.
    # Se retiran: dos cifras para lo mismo es peor que una sola.
    ("cxp", "Refinanciamiento"),
    ("cxp", "Proveedores"),
}


def get_token():
    r = requests.post(TOKEN_URL, data={
        "grant_type": "password", "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET, "username": USERNAME,
        "password": PASSWORD, "scope": PBI_SCOPE,
    }, timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]


def espera_429(r):
    """Segundos a esperar cuando Power BI responde 429, según lo que él pide.

    Copiada de fetch_powerbi, donde ya estaba resuelta. Aquí se leía solo la
    cabecera Retry-After y, cuando no venía, se esperaban 12 segundos fijos:
    Power BI suele decirlo únicamente en el texto ("Retry in 32 seconds"), así
    que se reintentaba demasiado pronto, se gastaban los tres intentos y la
    serie se perdía entera.
    """
    cab = r.headers.get("Retry-After")
    if cab and str(cab).strip().isdigit():
        return min(int(cab), 90)
    m = re.search(r"[Rr]etry in (\d+) second", r.text or "")
    return min(int(m.group(1)) + 2, 90) if m else 15


def dax(token, dataset_id, query, label="q", retries=3):
    ws = WS_POR_DATASET.get(dataset_id, WS_ID)
    url = f"{PBI_BASE}/groups/{ws}/datasets/{dataset_id}/executeQueries"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {"queries": [{"query": query}], "serializerSettings": {"includeNulls": True}}
    # Un 429 NO gasta intento: es "espera y vuelve", no un fallo. Contándolo
    # como intento, tres throttles seguidos —que es lo normal cuando la
    # corrida crece— mataban la consulta aunque el dato estuviera ahí.
    esperas = 0
    attempt = 0
    while attempt < retries:
        try:
            r = requests.post(url, json=body, headers=headers, timeout=60)
            if r.status_code == 429:
                if esperas >= 6:
                    print(f"      [{label}] throttled 6 veces — se abandona")
                    DIAG_SERIES.append({
                        "consulta": f"429:{label}", "http": 429,
                        "error": "Power BI limitó las peticiones seis veces "
                                 "seguidas; la corrida pide más consultas de "
                                 "las que el espacio de trabajo admite"})
                    return None
                esperas += 1
                wait = espera_429(r)
                print(f"      [{label}] throttled — espera {wait}s ({esperas}/6)")
                time.sleep(wait)
                continue
            if r.status_code != 200:
                if attempt == retries - 1:
                    print(f"      [{label}] HTTP {r.status_code}: {r.text[:150]}")
                return None
            tables = r.json().get("results", [{}])[0].get("tables", [])
            return tables[0].get("rows", []) if tables else []
        except Exception as e:
            attempt += 1
            if attempt >= retries:
                print(f"      [{label}] error: {e}")
            else:
                time.sleep(3)
    return None


def _sanear_no_finitos(obj):
    """Reemplaza Infinity/-Infinity/NaN por None dentro de listas y dicts.

    Devuelve cuántos reemplazó. Necesario porque json.dumps los escribiría
    literalmente y el JSON dejaría de ser parseable por el navegador.
    """
    import math
    n = 0
    if isinstance(obj, dict):
        it = obj.items()
    elif isinstance(obj, list):
        it = enumerate(obj)
    else:
        return 0
    for k, v in it:
        if isinstance(v, float) and not math.isfinite(v):
            obj[k] = None
            n += 1
        else:
            n += _sanear_no_finitos(v)
    return n


def hoy_lima():
    """Hoy en Lima (UTC-5). El runner de GitHub corre en UTC y desde las 7 de
    la tarde de Lima adelantaba el dia — y con el, el mes, los ultimos dias
    del mes."""
    return (datetime.datetime.now(datetime.timezone.utc)
            - datetime.timedelta(hours=5)).date()


# Mes en curso, en formato "AAAA-MM". Lo usa _ultimo() para no confundir un
# mes a medias con un cierre.
PERIODO_PARCIAL = None


def build_periodos():
    """Lista de (anio, mes) desde Ene-2025 hasta el mes EN CURSO.

    El mes en curso va incompleto por definicion, y antes se excluia por eso:
    en setiembre la app solo mostraba hasta agosto y no habia forma de ver el
    avance del mes. Se incluye igual, marcado como parcial, porque saber que
    llevamos 12 dias de setiembre es informacion — siempre que quede claro
    que no se compara de igual a igual contra un mes cerrado.
    """
    global PERIODO_PARCIAL
    hoy = hoy_lima()
    fin_anio, fin_mes = hoy.year, hoy.month
    PERIODO_PARCIAL = f"{fin_anio:04d}-{fin_mes:02d}"
    out, a, m = [], INICIO_ANIO, INICIO_MES
    while (a, m) <= (fin_anio, fin_mes):
        out.append((a, m))
        m += 1
        if m > 12:
            m, a = 1, a + 1
    return out


def _probar_par(token, dataset_id, medida, tbl, col):
    q = (f'EVALUATE ROW("v", CALCULATE([{medida}], '
         f"FILTER(ALL('{tbl}'), YEAR('{tbl}'[{col}]) = 2025 && MONTH('{tbl}'[{col}]) = 6)))")
    rows = dax(token, dataset_id, q, f"probe {tbl}[{col}]", retries=1)
    return rows is not None and len(rows) > 0


def columnas_fecha_reales(token, dataset_id):
    """Descubre columnas de tipo fecha del modelo vía DMV INFO.COLUMNS()/INFO.TABLES()."""
    tablas = {}
    rows = dax(token, dataset_id, "EVALUATE INFO.TABLES()", "info-tables", retries=1)
    if rows:
        for r in rows:
            tid = r.get("[ID]", r.get("ID"))
            nm  = r.get("[Name]", r.get("Name"))
            if tid is not None and nm:
                tablas[tid] = nm

    cols = dax(token, dataset_id, "EVALUATE INFO.COLUMNS()", "info-columns", retries=1)
    pares = []
    if cols:
        for c in cols:
            nm = c.get("[ExplicitName]", c.get("ExplicitName")) or c.get("[Name]", c.get("Name"))
            dt = str(c.get("[ExplicitDataType]", c.get("ExplicitDataType", "")) or "")
            tid = c.get("[TableID]", c.get("TableID"))
            tname = tablas.get(tid)
            if not nm or not tname:
                continue
            # 9 = DateTime en el modelo tabular; además acepta nombres típicos
            es_fecha = dt == "9" or any(k in nm.lower() for k in ("fecha", "date", "dia", "día"))
            if es_fecha:
                pares.append((tname, nm))
    return pares


def detectar_fecha(token, dataset_id, medida):
    """Encuentra la tabla/columna de fecha que sí filtra la medida."""
    # 1) Nombres frecuentes (rápido)
    for tbl, col in DATE_CANDIDATES:
        if _probar_par(token, dataset_id, medida, tbl, col):
            print(f"    Tabla de fecha: '{tbl}'[{col}]")
            return tbl, col

    # 2) Descubrimiento real del modelo
    print("    Candidatos comunes fallaron — leyendo esquema del modelo...")
    pares = columnas_fecha_reales(token, dataset_id)
    if pares:
        print(f"    Columnas de fecha halladas: {pares[:12]}")
    vistos = set(DATE_CANDIDATES)
    for tbl, col in pares:
        if (tbl, col) in vistos:
            continue
        vistos.add((tbl, col))
        if _probar_par(token, dataset_id, medida, tbl, col):
            print(f"    Tabla de fecha: '{tbl}'[{col}]")
            return tbl, col
    return None, None


def serie_batch(token, dataset_id, medidas, tbl, col, periodos):
    """Un solo query DAX que devuelve todos los meses y todas las medidas."""
    a_ini, m_ini = periodos[0]
    a_fin, m_fin = periodos[-1]
    ult_dia = 31
    cols = []
    for i, med in enumerate(medidas):
        cols.append(
            f'    "m{i}", VAR a = [Anio] VAR m = [Mes] RETURN '
            f"CALCULATE([{med}], FILTER(ALL('{tbl}'), "
            f"YEAR('{tbl}'[{col}]) = a && MONTH('{tbl}'[{col}]) = m))"
        )
    cols_txt = ",\n".join(cols)
    q = f"""EVALUATE
VAR Periodos =
    DISTINCT(
        SELECTCOLUMNS(
            FILTER(
                ALL('{tbl}'),
                '{tbl}'[{col}] >= DATE({a_ini}, {m_ini}, 1)
                && '{tbl}'[{col}] <= DATE({a_fin}, {m_fin}, {ult_dia})
            ),
            "Anio", YEAR('{tbl}'[{col}]),
            "Mes",  MONTH('{tbl}'[{col}])
        )
    )
RETURN
ADDCOLUMNS(
    Periodos,
{cols_txt}
)"""
    rows = dax(token, dataset_id, q, "serie-batch")
    if not rows:
        return None

    # Indexa resultados por (anio, mes)
    por_periodo = {}
    for row in rows:
        a = row.get("[Anio]", row.get("Anio"))
        m = row.get("[Mes]",  row.get("Mes"))
        if a is None or m is None:
            continue
        por_periodo[(int(a), int(m))] = row

    salida = {med: [] for med in medidas}
    for (a, m) in periodos:
        row = por_periodo.get((a, m))
        for i, med in enumerate(medidas):
            v = row.get(f"[m{i}]", row.get(f"m{i}")) if row else None
            salida[med].append(v)
    return salida


def serie_mes_a_mes(token, dataset_id, medidas, tbl, col, periodos):
    """Fallback: un query por mes con todas las medidas en un ROW()."""
    salida = {med: [] for med in medidas}
    for (a, m) in periodos:
        partes = []
        for i, med in enumerate(medidas):
            partes.append(
                f'  "m{i}", CALCULATE([{med}], FILTER(ALL(\'{tbl}\'), '
                f"YEAR('{tbl}'[{col}]) = {a} && MONTH('{tbl}'[{col}]) = {m}))"
            )
        q = "EVALUATE\nROW(\n" + ",\n".join(partes) + "\n)"
        rows = dax(token, dataset_id, q, f"{a}-{m:02d}")
        row = rows[0] if rows else None
        for i, med in enumerate(medidas):
            v = row.get(f"[m{i}]", row.get(f"m{i}")) if row else None
            salida[med].append(v)
        print(f"      {a}-{m:02d} ok")
    return salida


CXC_DATE_TBL = "LocalDateTable_9de71e0e-3dad-40db-b49a-7124064ba537"

MESES_CORTOS = {
    "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
    "jul": 7, "ago": 8, "sep": 9, "set": 9, "oct": 10, "nov": 11, "dic": 12,
}


def serie_cxc_morosidad(token, ds_id, periodos):
    """Serie mensual de % Morosidad — consulta exacta del gráfico "% DE MOROSIDAD".

    Capturada con Copiar consulta el 2026-09-06. Dos hallazgos que hacían
    imposible obtener esta serie con la sonda genérica:

    1. La medida del gráfico es '[% MOROSIDAD x mes]', NO '[% MOROSIDAD]'.
       La segunda es un escalar que ignora la fecha — por eso la tarjeta
       devolvía siempre el mismo número y la serie salía "plana".
    2. El eje de tiempo no es una tabla Calendario normal: el año viene de
       una tabla de fechas automática ('LocalDateTable_9de71e0e-...') y el
       mes de 'Calendario (FACT)'. Ninguna de las dos estaba en
       DATE_CANDIDATES, y la detección automática no podía adivinarlas.

    Se envía la consulta tal cual la genera Power BI (sin reescribirla) y
    después se mapean las filas a los períodos del dashboard.
    """
    q = (
        "DEFINE\n"
        "\tVAR __DS0FilterTable = \n"
        "\t\tFILTER(\n"
        f"\t\t\tKEEPFILTERS(VALUES('{CXC_DATE_TBL}'[Año])),\n"
        f"\t\t\t'{CXC_DATE_TBL}'[Año] > 2024\n"
        "\t\t)\n\n"
        "\tVAR __DS0Core = \n"
        "\t\tSUMMARIZECOLUMNS(\n"
        f"\t\t\t'{CXC_DATE_TBL}'[Año],\n"
        "\t\t\t'Calendario (FACT)'[Mes Corto],\n"
        "\t\t\t'Calendario (FACT)'[Mes Numero],\n"
        "\t\t\t__DS0FilterTable,\n"
        "\t\t\t\"v__MOROSIDAD_x_mes\", 'DATA_FACTURACION'[% MOROSIDAD x mes]\n"
        "\t\t)\n\n"
        "EVALUATE\n\t__DS0Core\n\n"
        "ORDER BY\n"
        f"\t'{CXC_DATE_TBL}'[Año],\n"
        "\t'Calendario (FACT)'[Mes Numero]"
    )
    rows = dax(token, ds_id, q, "cxc-morosidad-mensual")
    if not rows:
        return None

    por_periodo = {}
    for r in rows:
        anio = r.get(f"{CXC_DATE_TBL}[Año]") or r.get("[Año]")
        mes = r.get("Calendario (FACT)[Mes Numero]") or r.get("[Mes Numero]")
        val = r.get("[v__MOROSIDAD_x_mes]") or r.get("v__MOROSIDAD_x_mes")
        if mes is None:
            corto = str(r.get("Calendario (FACT)[Mes Corto]") or "").strip().lower()[:3]
            mes = MESES_CORTOS.get(corto)
        if anio is None or mes is None or val is None:
            continue
        try:
            por_periodo[(int(anio), int(mes))] = float(val)
        except (TypeError, ValueError):
            continue

    if not por_periodo:
        return None
    return [por_periodo.get(p) for p in periodos]


MERMAS_MEDIDAS = [
    ("v__Merma_total",             "% Merma Total"),
    ("SumDesviacion_Bases",        "Desviación Bases"),
    ("CANT_REAL_CONSUMO_INTERNO",  "Consumo Interno"),
    ("CANT_REAL_DESTRUCCION",      "Destrucción"),
]


# Plantas. Cada entrada: (etiqueta, valor de [almacen], medida).
#
# Agrupar por [almacen] con la medida base NO sirve: el mapeo de almacén era
# correcto, pero cada planta usa su propia medida. ATE usa '% Merma total' y
# Pachacamac usa '% Merma total SALSAS' — con la base, Pachacamac daba cifras
# que no coincidían con el reporte. Tercera vez que generalizar produce
# números plausibles y equivocados, así que solo entra lo capturado.
#
# Falta PLANTA TERCEROS: sin su Copiar consulta no se sabe ni su almacén ni
# su medida, y no se van a deducir.
# Plantas. Cada entrada: (etiqueta, [almacenes], medida, lleva_TIPO_DE_BASE).
#
# Nada de esto es deducible — los cuatro visuales de planta de un mismo
# reporte resultaron distintos entre sí:
#
#   · ATE y Pachacamac: 1 almacén, 7 filtros (con TIPO DE BASE), y medidas
#     DISTINTAS entre ellas ('% Merma total' vs '% Merma total SALSAS').
#   · Terceros: 3 almacenes agrupados, 6 filtros (SIN TIPO DE BASE).
#
# Y el nombre del visual no dice nada del almacén: "PLANTA ATE" es
# "Lacteos Producción" y "PLANTA TERCEROS" son tres maquilas.
#
# Nota sobre '% Merma total MAQUILA': se llegó a descartar por no coincidir
# con el gráfico de la unidad de negocio MAQUILA (que usa MMAQUILA, con
# doble M). Sí se usa — es la de Terceros. Dos medidas de nombre casi igual
# para cosas distintas.
PLANTAS_MERMAS = [
    # etiqueta,           almacenes,                    medida,                  tipo_base
    ("Planta ATE",        ["Lacteos Producción"],       "% Merma total",         True),
    ("Planta Pachacamac", ["Salsas Producción"],        "% Merma total SALSAS",  True),
    ("Planta Terceros",   ["Abuela Maquila",
                           "Piamonte Maquila",
                           "Lacteos Dosimetria"],       "% Merma total MAQUILA", False),
]


def _q_mermas_planta(anio, almacenes, medida, alias, tipo_base=True):
    """Merma mensual de UNA planta: filtro de [almacen] + su medida propia.

    Reproduce las consultas capturadas de PLANTA ATE, PLANTA PACHACAMAC y
    PLANTA TERCEROS (Copiar consulta, 2026-09-06). Parte del bloque de
    filtros del segmento (5 ó 6 según `tipo_base`) y le añade el TREATAS de
    almacén como último VAR.
    """
    base = _q_mermas_segmento(anio, "Tabla Mermas", medida, alias, tipo_base)
    n_ultimo = 6 if tipo_base else 5          # cuántos VAR trae ya el bloque
    nuevo = n_ultimo + 1
    # Las llaves del conjunto son obligatorias: TREATAS({"a","b"}, col).
    # Sin ellas DAX interpreta cada valor como un argumento suelto y falla.
    vals = ",\n".join(f'\t\t\t\t"{a}"' for a in almacenes).lstrip("\t")
    filtro = (f"\tVAR __DS0FilterTable{nuevo} = \n"
              f"\t\tTREATAS(\n\t\t\t{{{vals}}},\n"
              f"\t\t\t'Tabla Mermas'[almacen]\n\t\t)\n\n")
    base = base.replace("\tVAR __DS0Core = \n", filtro + "\tVAR __DS0Core = \n")
    return base.replace(
        f"\t\t\t__DS0FilterTable{n_ultimo},\n",
        f"\t\t\t__DS0FilterTable{n_ultimo},\n\t\t\t__DS0FilterTable{nuevo},\n",
    )


def _q_mermas(anio):
    """Consulta del gráfico "PAUNO" de MERMA MENSUAL POR PLANTA, tal cual la
    genera Power BI (Copiar consulta, 2026-09-06), con el año del segmentador
    como parámetro.

    Los 6 filtros del visual se reproducen literalmente. Son los que hacen
    que el número sea comparable con lo que ve el equipo en el reporte:
      1. 'Calendario'[Date] >= 31/07/2025 (ventana móvil del visual)
      2. 'Tabla Mermas'[TIPO DE BASE] no vacío
      3. 'Calendario'[Año] = <anio>  ← único parámetro; equivale a mover el
         segmentador ANO del reporte, que es justamente lo que haría un
         usuario para ver 2025 en vez de 2026.
      4. categoria_producto excluye BONIFICACION Y REBATES, CHATARRA,
         INTERESES, MATERIA PRIMA, SERVICIOS, SUMINISTROS y vacío
      5. producto excluye 2 ítems puntuales
      6. estado distinto de "Cancelado"

    Sin estos filtros el % de merma sale inflado: incluye chatarra, materia
    prima, servicios y comprobantes anulados.
    """
    return (
        "DEFINE\n"
        "\tVAR __DS0FilterTable = \n"
        "\t\tFILTER(\n"
        "\t\t\tKEEPFILTERS(VALUES('Calendario'[Date])),\n"
        "\t\t\t'Calendario'[Date] >= (DATE(2025, 7, 31) + TIME(0, 0, 1))\n"
        "\t\t)\n\n"
        "\tVAR __DS0FilterTable2 = \n"
        "\t\tFILTER(\n"
        "\t\t\tKEEPFILTERS(VALUES('Tabla Mermas'[TIPO DE BASE])),\n"
        "\t\t\tNOT('Tabla Mermas'[TIPO DE BASE] IN {BLANK()})\n"
        "\t\t)\n\n"
        "\tVAR __DS0FilterTable3 = \n"
        f"\t\tTREATAS({{{anio}}}, 'Calendario'[Año])\n\n"
        "\tVAR __DS0FilterTable4 = \n"
        "\t\tFILTER(\n"
        "\t\t\tKEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[categoria_producto])),\n"
        "\t\t\tNOT(\n"
        "\t\t\t\t'Maestra de Facturacion (Total)'[categoria_producto] IN {\"BONIFICACION Y REBATES\",\n"
        "\t\t\t\t\t\"CHATARRA\",\n\t\t\t\t\t\"INTERESES\",\n\t\t\t\t\t\"MATERIA PRIMA\",\n"
        "\t\t\t\t\t\"SERVICIOS\",\n\t\t\t\t\t\"SUMINISTROS\",\n\t\t\t\t\tBLANK()}\n"
        "\t\t\t)\n\t\t)\n\n"
        "\tVAR __DS0FilterTable5 = \n"
        "\t\tFILTER(\n"
        "\t\t\tKEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[producto])),\n"
        "\t\t\tNOT(\n"
        "\t\t\t\t'Maestra de Facturacion (Total)'[producto] IN {\"PAVO C/M C/ASA EP CONG (8 KG)\",\n"
        "\t\t\t\t\t\"ALIMENTACION COMERCIAL\"}\n"
        "\t\t\t)\n\t\t)\n\n"
        "\tVAR __DS0FilterTable6 = \n"
        "\t\tFILTER(\n"
        "\t\t\tKEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[estado])),\n"
        "\t\t\tNOT('Maestra de Facturacion (Total)'[estado] IN {\"Cancelado\"})\n"
        "\t\t)\n\n"
        "\tVAR __DS0Core = \n"
        "\t\tSUMMARIZECOLUMNS(\n"
        "\t\t\t'Calendario'[Año],\n\t\t\t'Calendario'[MES],\n"
        "\t\t\t__DS0FilterTable,\n\t\t\t__DS0FilterTable2,\n\t\t\t__DS0FilterTable3,\n"
        "\t\t\t__DS0FilterTable4,\n\t\t\t__DS0FilterTable5,\n\t\t\t__DS0FilterTable6,\n"
        "\t\t\t\"SumDesviacion_Bases\", IGNORE(CALCULATE(SUM('Tabla Mermas'[Desviacion Bases]))),\n"
        "\t\t\t\"CANT_REAL_CONSUMO_INTERNO\", IGNORE('Maestra de Kardex (Total)'[CANT REAL CONSUMO INTERNO]),\n"
        "\t\t\t\"CANT_REAL_DESTRUCCION\", IGNORE('Maestra de Kardex (Total)'[CANT REAL DESTRUCCION]),\n"
        "\t\t\t\"v__Merma_total\", 'Tabla Mermas'[% Merma total]\n"
        "\t\t)\n\n"
        "EVALUATE\n\t__DS0Core\n\n"
        "ORDER BY\n\t'Calendario'[Año], 'Calendario'[MES]"
    )


# Segmentos de merma. Cada entrada: (etiqueta, medida, lleva_filtro_tipo_base)
#
# El sondeo por nombre se retiró: encontró '% Merma total MAQUILA', que EXISTE
# en el modelo pero NO es la del gráfico — el visual usa
# '% Merma total MMAQUILA' (doble M). Sus valores no coincidían con el reporte.
# Con dos medidas de nombre casi idéntico, adivinar por el título del visual
# no es fiable, así que solo entra lo confirmado.
#
# Tampoco hay un patrón uniforme de filtros: B&D usa 5 (sin TIPO DE BASE) y
# MAQUILA usa 6 (con TIPO DE BASE). Por eso el flag es por segmento.
SEGMENTOS_MERMAS = [
    # etiqueta,   medida,                    tipo_base   origen
    ("B&D",       "% Merma total B&D",       False),   # Copiar consulta 2026-09-06
    ("MAQUILA",   "% Merma total MMAQUILA",  True),    # Copiar consulta 2026-09-06
    ("TIGO",      "% Merma total TIGO",      False),   # sondeo + valores validados
                                                       # contra el gráfico (meses 2-8)
]


def _q_mermas_segmento(anio, medida_tbl, medida, alias, tipo_base=False):
    """Consulta del gráfico "B&D" de MERMA MENSUAL POR UNIDAD DE NEGOCIO,
    tal cual la genera Power BI (Copiar consulta, 2026-09-06).

    El bloque de filtros varía según el visual: B&D usa 5 (sin TIPO DE BASE)
    y MAQUILA usa 6 (con él). `tipo_base` selecciona cuál, y la numeración de
    los VAR se ajusta sola. No es un patrón deducible — cada uno viene de su
    Copiar consulta.

    La medida también cambia por segmento y se pasa explícita: los nombres
    están en SEGMENTOS_MERMAS, ninguno se adivina.
    """
    # Numeración de los VAR: sin TIPO DE BASE van 5 filtros (patrón B&D);
    # con él van 6 y todo corre un número (patrón MAQUILA). Se respeta el
    # orden que genera Power BI en cada caso.
    n = 1
    tb = ""
    if tipo_base:
        tb = ("\tVAR __DS0FilterTable2 = \n"
              "\t\tFILTER(\n"
              "\t\t\tKEEPFILTERS(VALUES('Tabla Mermas'[TIPO DE BASE])),\n"
              "\t\t\tNOT('Tabla Mermas'[TIPO DE BASE] IN {BLANK()})\n"
              "\t\t)\n\n")
        n = 2
    v = lambda i: f"__DS0FilterTable{'' if i == 1 else i}"
    usados = "".join(f"\t\t\t{v(i)},\n" for i in range(1, n + 5))
    return (
        "DEFINE\n"
        "\tVAR __DS0FilterTable = \n"
        "\t\tFILTER(\n"
        "\t\t\tKEEPFILTERS(VALUES('Calendario'[Date])),\n"
        "\t\t\t'Calendario'[Date] >= (DATE(2025, 7, 31) + TIME(0, 0, 1))\n"
        "\t\t)\n\n"
        + tb +
        f"\tVAR {v(n + 1)} = \n"
        f"\t\tTREATAS({{{anio}}}, 'Calendario'[Año])\n\n"
        f"\tVAR {v(n + 2)} = \n"
        "\t\tFILTER(\n"
        "\t\t\tKEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[categoria_producto])),\n"
        "\t\t\tNOT(\n"
        "\t\t\t\t'Maestra de Facturacion (Total)'[categoria_producto] IN {\"BONIFICACION Y REBATES\",\n"
        "\t\t\t\t\t\"CHATARRA\",\n\t\t\t\t\t\"INTERESES\",\n\t\t\t\t\t\"MATERIA PRIMA\",\n"
        "\t\t\t\t\t\"SERVICIOS\",\n\t\t\t\t\t\"SUMINISTROS\",\n\t\t\t\t\tBLANK()}\n"
        "\t\t\t)\n\t\t)\n\n"
        f"\tVAR {v(n + 3)} = \n"
        "\t\tFILTER(\n"
        "\t\t\tKEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[producto])),\n"
        "\t\t\tNOT(\n"
        "\t\t\t\t'Maestra de Facturacion (Total)'[producto] IN {\"PAVO C/M C/ASA EP CONG (8 KG)\",\n"
        "\t\t\t\t\t\"ALIMENTACION COMERCIAL\"}\n"
        "\t\t\t)\n\t\t)\n\n"
        f"\tVAR {v(n + 4)} = \n"
        "\t\tFILTER(\n"
        "\t\t\tKEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[estado])),\n"
        "\t\t\tNOT('Maestra de Facturacion (Total)'[estado] IN {\"Cancelado\"})\n"
        "\t\t)\n\n"
        "\tVAR __DS0Core = \n"
        "\t\tSUMMARIZECOLUMNS(\n"
        "\t\t\t'Calendario'[Año],\n\t\t\t'Calendario'[MES],\n"
        + usados +
        f"\t\t\t\"{alias}\", '{medida_tbl}'[{medida}]\n"
        "\t\t)\n\n"
        "EVALUATE\n\t__DS0Core\n\n"
        "ORDER BY\n\t'Calendario'[Año], 'Calendario'[MES]"
    )


MARGEN_LOCALDATE = "LocalDateTable_17f9bf1b-53ab-4dc5-bca5-92ec1d35bc75"

# Bloque de filtros del visual "MARGEN VARIABLE PAUNO", tal cual lo genera
# Power BI (Copiar consulta, 2026-09-06). Son ONCE filtros sobre CUATRO
# tablas distintas: 'Exl A Maestra de Facturas de Venta', 'Maestra de
# Facturacion (Total)', 'Maestra de Kardex (Total)' y 'OrdenxFactura'.
#
# Ojo: las exclusiones NO son las mismas en cada tabla. Por ejemplo
# 'Exl A...' excluye {CHATARRA, SERVICIOS} y 'Maestra de Facturacion'
# excluye {BONIFICACION Y REBATES, INTERESES, MATERIA PRIMA, SUMINISTROS,
# SERVICIOS} — CHATARRA no está en esa segunda lista. Copiar una sobre otra
# cambiaría el resultado, así que va literal.
#
# __ANIO__ es el único parámetro (el segmentador ANO del reporte).
_MARGEN_FILTROS = """	VAR __DS0FilterTable = 
		TREATAS({"VENTA BRUTA"}, 'Exl A Maestra de Facturas de Venta'[CATEGORIZACION])

	VAR __DS0FilterTable2 = 
		TREATAS({__ANIO__}, 'Calendario'[Año])

	VAR __DS0FilterTable3 = 
		FILTER(
			KEEPFILTERS(VALUES('Exl A Maestra de Facturas de Venta'[categoria_producto])),
			NOT(
				'Exl A Maestra de Facturas de Venta'[categoria_producto] IN {"CHATARRA",
					"SERVICIOS",
					BLANK()}
			)
		)

	VAR __DS0FilterTable4 = 
		FILTER(
			KEEPFILTERS(VALUES('Exl A Maestra de Facturas de Venta'[producto])),
			NOT(
				'Exl A Maestra de Facturas de Venta'[producto] IN {"PAVO C/M C/ASA EP CONG (8 KG)",
					"ALIMENTACION COMERCIAL"}
			)
		)

	VAR __DS0FilterTable5 = 
		FILTER(
			KEEPFILTERS(VALUES('Exl A Maestra de Facturas de Venta'[estado])),
			NOT('Exl A Maestra de Facturas de Venta'[estado] IN {"Cancelado"})
		)

	VAR __DS0FilterTable6 = 
		FILTER(
			KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[categoria_producto])),
			NOT(
				'Maestra de Facturacion (Total)'[categoria_producto] IN {BLANK(),
					"BONIFICACION Y REBATES",
					"INTERESES",
					"MATERIA PRIMA",
					"SUMINISTROS",
					"SERVICIOS"}
			)
		)

	VAR __DS0FilterTable7 = 
		FILTER(
			KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[estado])),
			NOT('Maestra de Facturacion (Total)'[estado] IN {"Cancelado"})
		)

	VAR __DS0FilterTable8 = 
		FILTER(
			KEEPFILTERS(VALUES('Maestra de Kardex (Total)'[CUENTA ORIGEN])),
			NOT('Maestra de Kardex (Total)'[CUENTA ORIGEN] IN {BLANK()})
		)

	VAR __DS0FilterTable9 = 
		FILTER(
			KEEPFILTERS(VALUES('OrdenxFactura'[categoria_producto])),
			NOT(
				'OrdenxFactura'[categoria_producto] IN {BLANK(),
					"BONIFICACION Y REBATES",
					"CHATARRA",
					"SERVICIOS"}
			)
		)

	VAR __DS0FilterTable10 = 
		FILTER(
			KEEPFILTERS(VALUES('OrdenxFactura'[estado])),
			NOT('OrdenxFactura'[estado] IN {"Cancelado"})
		)

	VAR __DS0FilterTable11 = 
		FILTER(
			KEEPFILTERS(VALUES('OrdenxFactura'[producto])),
			NOT(
				'OrdenxFactura'[producto] IN {"PAVO C/M C/ASA EP CONG (8 KG)",
					"ALIMENTACION COMERCIAL"}
			)
		)
"""

_MARGEN_USADOS = "".join(
    "\t\t\t__DS0FilterTable%s,\n" % ("" if i == 1 else i) for i in range(1, 12)
)


def _q_margen(anio, medida_tbl, medida, alias):
    """Consulta del visual "MARGEN VARIABLE PAUNO" con los 11 filtros reales.

    De la consulta original solo se conserva __DS0Core. Lo demás
    (__DS0PrimaryWindowed, __DS0Secondary, NATURALLEFTOUTERJOIN,
    SUBSTITUTEWITHINDEX) es la maquinaria que Power BI usa para pintar el eje
    y la leyenda del gráfico: no cambia los valores, solo los ordena y los
    empareja. Nosotros queremos las cifras, así que evaluamos el core.

    Se mantienen las CINCO columnas de agrupación del original — las tres de
    la tabla de fechas automática más 'Calendario'[Año]/[MES] — porque quitar
    alguna cambiaría el contexto en que se evalúa la medida. Las filas del
    producto cruzado que no corresponden salen en blanco y se descartan al
    leer.
    """
    ld = MARGEN_LOCALDATE
    return (
        "DEFINE\n"
        + _MARGEN_FILTROS.replace("__ANIO__", str(anio))
        + "\n\tVAR __DS0Core = \n"
        "\t\tSUMMARIZECOLUMNS(\n"
        f"\t\t\t'{ld}'[Año],\n"
        f"\t\t\t'{ld}'[Mes],\n"
        f"\t\t\t'{ld}'[NroMes],\n"
        "\t\t\t'Calendario'[MES],\n"
        "\t\t\t'Calendario'[Año],\n"
        + _MARGEN_USADOS
        + f"\t\t\t\"{alias}\", '{medida_tbl}'[{medida}]\n"
        "\t\t)\n\n"
        "EVALUATE\n\t__DS0Core\n\n"
        f"ORDER BY\n\t'{ld}'[Año], '{ld}'[NroMes]"
    )


# Bloque de filtros del visual "PAUNO" de Precio por Kilo Vs Costo por Kilo
# (Copiar consulta, 2026-09-06). Son DOCE, no once: lleva uno que el visual
# de margen no tiene — excluye "BONIFICACION SELL IN" de [DESCRIPCION] — y
# eso corre la numeración de todos los demás (el año pasa de 2 a 3).
#
# Se escribe completo en vez de derivarlo del bloque de margen: en este
# reporte cada visual resultó tener su propio juego de filtros, y armarlo
# por diferencias es justo el tipo de atajo que ya produjo cifras
# equivocadas en Mermas.
_MARGEN_PK_FILTROS = """	VAR __DS0FilterTable = 
		TREATAS({"VENTA BRUTA"}, 'Exl A Maestra de Facturas de Venta'[CATEGORIZACION])

	VAR __DS0FilterTable2 = 
		FILTER(
			KEEPFILTERS(VALUES('Exl A Maestra de Facturas de Venta'[DESCRIPCION])),
			NOT('Exl A Maestra de Facturas de Venta'[DESCRIPCION] IN {"BONIFICACION SELL IN"})
		)

	VAR __DS0FilterTable3 = 
		TREATAS({__ANIO__}, 'Calendario'[Año])

	VAR __DS0FilterTable4 = 
		FILTER(
			KEEPFILTERS(VALUES('Exl A Maestra de Facturas de Venta'[categoria_producto])),
			NOT(
				'Exl A Maestra de Facturas de Venta'[categoria_producto] IN {"CHATARRA",
					"SERVICIOS",
					BLANK()}
			)
		)

	VAR __DS0FilterTable5 = 
		FILTER(
			KEEPFILTERS(VALUES('Exl A Maestra de Facturas de Venta'[producto])),
			NOT(
				'Exl A Maestra de Facturas de Venta'[producto] IN {"PAVO C/M C/ASA EP CONG (8 KG)",
					"ALIMENTACION COMERCIAL"}
			)
		)

	VAR __DS0FilterTable6 = 
		FILTER(
			KEEPFILTERS(VALUES('Exl A Maestra de Facturas de Venta'[estado])),
			NOT('Exl A Maestra de Facturas de Venta'[estado] IN {"Cancelado"})
		)

	VAR __DS0FilterTable7 = 
		FILTER(
			KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[categoria_producto])),
			NOT(
				'Maestra de Facturacion (Total)'[categoria_producto] IN {BLANK(),
					"BONIFICACION Y REBATES",
					"INTERESES",
					"MATERIA PRIMA",
					"SUMINISTROS",
					"SERVICIOS"}
			)
		)

	VAR __DS0FilterTable8 = 
		FILTER(
			KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[estado])),
			NOT('Maestra de Facturacion (Total)'[estado] IN {"Cancelado"})
		)

	VAR __DS0FilterTable9 = 
		FILTER(
			KEEPFILTERS(VALUES('Maestra de Kardex (Total)'[CUENTA ORIGEN])),
			NOT('Maestra de Kardex (Total)'[CUENTA ORIGEN] IN {BLANK()})
		)

	VAR __DS0FilterTable10 = 
		FILTER(
			KEEPFILTERS(VALUES('OrdenxFactura'[categoria_producto])),
			NOT(
				'OrdenxFactura'[categoria_producto] IN {BLANK(),
					"BONIFICACION Y REBATES",
					"CHATARRA",
					"SERVICIOS"}
			)
		)

	VAR __DS0FilterTable11 = 
		FILTER(
			KEEPFILTERS(VALUES('OrdenxFactura'[estado])),
			NOT('OrdenxFactura'[estado] IN {"Cancelado"})
		)

	VAR __DS0FilterTable12 = 
		FILTER(
			KEEPFILTERS(VALUES('OrdenxFactura'[producto])),
			NOT(
				'OrdenxFactura'[producto] IN {"PAVO C/M C/ASA EP CONG (8 KG)",
					"ALIMENTACION COMERCIAL"}
			)
		)
"""

_MARGEN_PK_USADOS = "".join(
    "\t\t\t__DS0FilterTable%s,\n" % ("" if i == 1 else i) for i in range(1, 13)
)


def _q_margen_precio_costo(anio):
    """Precio x Kilo y Costo x Kilo mensuales, con los 12 filtros del visual.

    Ambas medidas salen de la misma consulta, igual que en el gráfico. La app
    venía trayendo 'Costo x Kilo' por la sonda genérica, es decir SIN ninguno
    de estos filtros.
    """
    ld = MARGEN_LOCALDATE
    return (
        "DEFINE\n"
        + _MARGEN_PK_FILTROS.replace("__ANIO__", str(anio))
        + "\n\tVAR __DS0Core = \n"
        "\t\tSUMMARIZECOLUMNS(\n"
        f"\t\t\t'{ld}'[Año],\n"
        "\t\t\t'Calendario'[Mes Corto],\n"
        "\t\t\t'Calendario'[MES],\n"
        + _MARGEN_PK_USADOS
        + "\t\t\t\"Precio_x_Kilo\", 'Medidas Julito'[Precio x Kilo],\n"
        "\t\t\t\"Costo_x_Kilo\", 'Medidas Julito'[Costo x Kilo]\n"
        "\t\t)\n\n"
        "EVALUATE\n\t__DS0Core\n\n"
        "ORDER BY\n\t'Calendario'[MES]"
    )


def _q_margen_ventas(anio):
    """Ventas mensuales (S/.) del visual "VENTAS (S/.) PAUNO".

    Copiar consulta 2026-09-06. Comparte los DOCE filtros del visual de
    precio/costo — verificado por diff carácter a carácter contra la
    consulta original, no asumido. Cambian dos cosas: agrupa además por
    'Calendario'[Año], y la cifra es una suma de columna, no una medida:
    SUM('Exl A Maestra de Facturas de Venta'[Monto_Neto_Factura]).

    Como en el visual de margen, se conserva solo __DS0Core; el resto es la
    maquinaria de eje y leyenda.
    """
    ld = MARGEN_LOCALDATE
    return (
        "DEFINE\n"
        + _MARGEN_PK_FILTROS.replace("__ANIO__", str(anio))
        + "\n\tVAR __DS0Core = \n"
        "\t\tSUMMARIZECOLUMNS(\n"
        f"\t\t\t'{ld}'[Año],\n"
        "\t\t\t'Calendario'[Mes Corto],\n"
        "\t\t\t'Calendario'[MES],\n"
        "\t\t\t'Calendario'[Año],\n"
        + _MARGEN_PK_USADOS
        + "\t\t\t\"SumMonto_Neto_Factura\", "
        "CALCULATE(SUM('Exl A Maestra de Facturas de Venta'[Monto_Neto_Factura]))\n"
        "\t\t)\n\n"
        "EVALUATE\n\t__DS0Core\n\n"
        "ORDER BY\n\t'Calendario'[MES]"
    )


def serie_margen_ventas(token, ds_id, periodos):
    """Serie mensual de Ventas (S/.)."""
    anios = sorted({a for a, _ in periodos})
    mapa = {}
    for anio in anios:
        for r in dax(token, ds_id, _q_margen_ventas(anio), f"margen-ventas-{anio}") or []:
            v = r.get("[SumMonto_Neto_Factura]", r.get("SumMonto_Neto_Factura"))
            if v is None:
                continue
            mes = r.get("Calendario[MES]") or r.get("[MES]")
            if isinstance(mes, str):
                mes = MESES_CORTOS.get(mes.strip().lower()[:3])
            if mes is None:
                continue
            try:
                mapa[(int(anio), int(mes))] = float(v)
            except (TypeError, ValueError):
                pass
    serie = [mapa.get(pp) for pp in periodos]
    if not any(x is not None for x in serie):
        print("    ✗ Ventas (S/.): sin datos")
        return {}
    print(f"    [Ventas (S/.)]: "
          f"{sum(1 for x in serie if x is not None)}/{len(serie)} meses")
    return {"Ventas (S/.)": serie}


def serie_margen_precio_costo(token, ds_id, periodos):
    """Series mensuales de Precio x Kilo y Costo x Kilo."""
    anios = sorted({a for a, _ in periodos})
    pares = {"Precio x Kilo": {}, "Costo x Kilo": {}}
    for anio in anios:
        for r in dax(token, ds_id, _q_margen_precio_costo(anio),
                     f"margen-pk-{anio}") or []:
            mes = r.get("Calendario[MES]") or r.get("[MES]")
            if isinstance(mes, str):
                mes = MESES_CORTOS.get(mes.strip().lower()[:3])
            if mes is None:
                continue
            for alias, etiqueta in (("Precio_x_Kilo", "Precio x Kilo"),
                                    ("Costo_x_Kilo", "Costo x Kilo")):
                v = r.get(f"[{alias}]", r.get(alias))
                if v is None:
                    continue
                try:
                    pares[etiqueta][(int(anio), int(mes))] = float(v)
                except (TypeError, ValueError):
                    pass

    out = {}
    for etiqueta, mapa in pares.items():
        serie = [mapa.get(pp) for pp in periodos]
        if any(x is not None for x in serie):
            out[etiqueta] = serie
            print(f"    [{etiqueta}]: "
                  f"{sum(1 for x in serie if x is not None)}/{len(serie)} meses")
    return out


def dataset_de_reporte(token, ws, report_id):
    """Devuelve el datasetId de un reporte. Necesario para FILLRATE, cuyo
    dataset no conocíamos: solo teníamos el id del reporte, de su URL."""
    r = requests.get(f"{PBI_BASE}/groups/{ws}/reports/{report_id}",
                     headers={"Authorization": f"Bearer {token}"}, timeout=25)
    if not r.ok:
        print(f"    ✗ no se pudo resolver el dataset del reporte: HTTP {r.status_code}")
        return None
    return r.json().get("datasetId")


# La tabla de fechas del reporte de compras. La borré yo por accidente al
# quitar las funciones muertas del fill rate: el corte por "\ndef " se llevó
# también las constantes que vivían entre dos funciones, y con ellas
# _COMPRAS_FILTROS. Las dos consultas del ratio morían con NameError, el
# try/except lo imprimía y seguía, y el KPI "Ratio consumo / compra" salió
# vacío con el reporte lleno en Power BI.
#
# _COMPRAS_FILTROS ya no hace falta: las consultas pasaron a mandarse
# verbatim desde el catálogo, que es como debieron estar desde el principio.
COMPRAS_LOCALDATE = "LocalDateTable_42b93aac-d3d2-4a95-b7b8-66bfc130de2b"


def _q_compras_ratio():
    """Ratio por categoría de insumo — la captura de la tabla, VERBATIM.

    Ver la nota de _q_compras_ratio_total: las dos estaban escritas a mano y
    las dos morían con NameError.
    """
    e = cargar_catalogo().get(
        "Eficiencia de Costo de compra  de materiales  (Operación)#a5e2b0c1992a")
    return e.get("dax") if e else None


def _q_compras_ratio_total():
    """Ratio mensual total — la captura del gráfico, enviada VERBATIM.

    Antes esta consulta estaba transcrita a mano, igual que la de categorías.
    Las dos referenciaban _COMPRAS_FILTROS, que no está definido en ninguna
    parte del archivo: cada llamada moría con NameError, el try/except de
    arriba lo imprimía y seguía, y el dataset 'compras' no llegaba a crearse.
    Por eso "Ratio consumo / compra" llevaba vacío en la app mientras el
    reporte estaba lleno en Power BI.

    Transcribir a mano fue el error; se manda la exportación tal cual.
    """
    e = cargar_catalogo().get(
        "Eficiencia de Costo de compra  de materiales  (Operación)#4db75af70c3b")
    return e.get("dax") if e else None


def serie_compras_ratio_total(token, ds_id, periodos):
    """Serie del ratio total, tal como la dibuja el gráfico del reporte."""
    filas = dax(token, ds_id, _q_compras_ratio_total(), "compras-ratio-total")
    if not filas:
        return None
    ld = COMPRAS_LOCALDATE
    mapa = {}
    for r in filas:
        anio = r.get("Calendario[Año]") or r.get("[Año]")
        mes = r.get(f"{ld}[NroMes]") or r.get("[NroMes]")
        v = r.get("[v_ratio]", r.get("v_ratio"))
        if anio is None or mes is None or v is None:
            continue
        try:
            mapa[(int(anio), int(mes))] = float(v)
        except (TypeError, ValueError):
            pass
    serie = [mapa.get(p) for p in periodos]
    return serie if any(x is not None for x in serie) else None


def serie_compras_ratio(token, ds_id, periodos):
    """Series mensuales de ratio consumo/compra, total y por categoría.

    Ratio > 100% significa que se consumió más de lo comprado (se está
    jalando inventario); < 100% que se compró de más. Por eso se publican
    también las cantidades: el ratio solo dice la dirección, no el tamaño.
    """
    # La captura de la tabla trae DOS EVALUATE: el primero es el eje del
    # visual y el segundo el cuerpo con los datos. dax() se queda con la
    # primera tabla, que aquí es la equivocada y llega sin medidas.
    q = _q_compras_ratio()
    tablas = dax_crudo(token, ds_id, q, "compras-ratio") if q else []
    filas = next((t for t in tablas
                  if t and any("ratio" in k.lower() for k in t[0])), None)
    if not filas:
        # Sin esto el fallo es mudo: el dataset 'compras' no llegaba a
        # crearse, el KPI "Ratio consumo / compra" salía vacío en la app y
        # nada en la corrida decía por qué. Un indicador sin dato tiene que
        # dejar rastro, siempre.
        DIAG_SERIES.append({
            "consulta": "compras:ratio", "http": 200,
            "error": "la consulta del visual 'Eficiencia de Costo de compra' no "
                     "devolvió filas; el dataset 'compras' se queda sin serie y "
                     "el KPI Ratio consumo/compra sin dato"})
        print("    ✗ compras-ratio: sin filas")
        return {}

    ld = COMPRAS_LOCALDATE
    por_cat, tot_cons, tot_comp = {}, {}, {}
    for r in filas:
        cat = (r.get("Maestra de Productos[data.categoria_producto]")
               or r.get("[data.categoria_producto]"))
        anio = r.get("Calendario[Año]") or r.get("[Año]")
        mes = r.get(f"{ld}[NroMes]") or r.get("[NroMes]")
        if cat is None or anio is None or mes is None:
            continue
        try:
            per = (int(anio), int(mes))
        except (TypeError, ValueError):
            continue
        cons = r.get("[Consumo_CANT_10_92_93]", r.get("Consumo_CANT_10_92_93"))
        comp = r.get("[compra_mensual_Cant_2]", r.get("compra_mensual_Cant_2"))
        ratio = r.get("[v_ratio]", r.get("v_ratio"))
        if ratio is not None:
            try:
                por_cat.setdefault(str(cat), {})[per] = float(ratio)
            except (TypeError, ValueError):
                pass
        # La compra viene en negativo en el modelo (salida de caja): se toma
        # el valor absoluto para poder sumarla con el consumo.
        for val, acum in ((cons, tot_cons), (comp, tot_comp)):
            if val is None:
                continue
            try:
                acum[per] = acum.get(per, 0.0) + abs(float(val))
            except (TypeError, ValueError):
                pass

    out = {}
    # El total viene de la consulta del GRÁFICO, no de sumar las categorías de
    # la tabla: sus filtros no coinciden (ver _q_compras_ratio_total).
    total = serie_compras_ratio_total(token, ds_id, periodos)
    if total:
        out["% Ratio Consumo/Compra"] = total
        print(f"    [% Ratio Consumo/Compra]: "
              f"{sum(1 for x in total if x is not None)}/{len(total)} meses")
        # Contraste informativo contra el total derivado de la tabla. No se
        # publica: sirve para detectar si alguna vez dejan de ser comparables.
        deriv = [(tot_cons.get(p) / tot_comp[p]) if tot_comp.get(p) else None
                 for p in periodos]
        difs = [abs(a - b) for a, b in zip(total, deriv)
                if a is not None and b is not None]
        if difs and max(difs) > 0.01:
            print(f"    · nota: gráfico vs tabla difieren hasta "
                  f"{max(difs) * 100:.2f} pp (alcances distintos)")
    for etiqueta, acum in (("Consumo (cant)", tot_cons), ("Compra (cant)", tot_comp)):
        serie = [acum.get(p) for p in periodos]
        if any(x is not None for x in serie):
            out[etiqueta] = serie
    for cat in sorted(por_cat):
        serie = [por_cat[cat].get(p) for p in periodos]
        if sum(1 for x in serie if x is not None) >= 3:
            out[f"% Ratio {cat}"] = serie
            print(f"    [% Ratio {cat}]: "
                  f"{sum(1 for x in serie if x is not None)}/{len(serie)} meses")
    return out


CXP_LOCALDATE = "LocalDateTable_42b93aac-d3d2-4a95-b7b8-66bfc130de2b"


def _q_cxp_rotacion():
    """Rotación de cuentas por pagar: compras acumuladas y días de pago.

    Copiar consulta 2026-09-06 sobre el gráfico "Rotación de cuentas por pagar
    (días)". Se conserva íntegro, incluido __ValueFilterDM0 — el filtro de
    valor que descarta los meses sin compras. Sin él aparecerían meses en cero
    que el gráfico no dibuja.

    'Calendario'[Columna Mostrar] = "MOSTRAR" es una bandera del calendario
    que decide qué meses entran al gráfico; no es un filtro de fecha, así que
    no se puede sustituir por un rango.

    Solo se quita el TOPN(1001), que limita filas sin cambiar valores.
    """
    ld = CXP_LOCALDATE
    return (
        "DEFINE\n"
        "\tVAR __DS0FilterTable = \n"
        "\t\tTREATAS({\"MOSTRAR\"}, 'Calendario'[Columna Mostrar])\n\n"
        "\tVAR __ValueFilterDM0 = \n"
        "\t\tFILTER(\n"
        "\t\t\tKEEPFILTERS(\n"
        "\t\t\t\tSUMMARIZECOLUMNS(\n"
        f"\t\t\t\t\t'{ld}'[Año],\n"
        "\t\t\t\t\t'Calendario'[Columna Mostrar],\n"
        "\t\t\t\t\t'Calendario'[Mes corto],\n"
        "\t\t\t\t\t'Calendario'[Mes Nº],\n"
        "\t\t\t\t\t__DS0FilterTable,\n"
        "\t\t\t\t\t\"TOTAL_ACUMULADO_CTA_60\", 'CUENTAS CONTABLES'[TOTAL ACUMULADO CTA 60],\n"
        "\t\t\t\t\t\"DIAS\", 'CUENTAS CONTABLES'[DIAS]\n"
        "\t\t\t\t)\n"
        "\t\t\t),\n"
        "\t\t\t[TOTAL_ACUMULADO_CTA_60] > 0\n"
        "\t\t)\n\n"
        "EVALUATE\n"
        "\tSUMMARIZECOLUMNS(\n"
        f"\t\t'{ld}'[Año],\n"
        "\t\t'Calendario'[Columna Mostrar],\n"
        "\t\t'Calendario'[Mes corto],\n"
        "\t\t'Calendario'[Mes Nº],\n"
        "\t\t__DS0FilterTable,\n"
        "\t\t__ValueFilterDM0,\n"
        "\t\t\"TOTAL_ACUMULADO_CTA_60\", 'CUENTAS CONTABLES'[TOTAL ACUMULADO CTA 60],\n"
        "\t\t\"DIAS\", 'CUENTAS CONTABLES'[DIAS]\n"
        "\t)\n\n"
        f"ORDER BY\n\t'{ld}'[Año], 'Calendario'[Mes Nº]"
    )


def serie_cxp_rotacion(token, ds_id, periodos):
    """Series mensuales de compras (cta 60) y días de cuentas por pagar."""
    filas = dax(token, ds_id, _q_cxp_rotacion(), "cxp-rotacion")
    if not filas:
        return {}
    ld = CXP_LOCALDATE
    pares = {"Compras (cta 60)": {}, "Días CxP": {}}
    for r in filas:
        anio = r.get(f"{ld}[Año]") or r.get("[Año]")
        mes = r.get("Calendario[Mes Nº]") or r.get("[Mes Nº]")
        if mes is None:
            corto = str(r.get("Calendario[Mes corto]") or "").strip().lower()[:3]
            mes = MESES_CORTOS.get(corto)
        if anio is None or mes is None:
            continue
        try:
            per = (int(anio), int(mes))
        except (TypeError, ValueError):
            continue
        for alias, etiqueta in (("TOTAL_ACUMULADO_CTA_60", "Compras (cta 60)"),
                                ("DIAS", "Días CxP")):
            v = r.get(f"[{alias}]", r.get(alias))
            if v is None:
                continue
            try:
                pares[etiqueta][per] = float(v)
            except (TypeError, ValueError):
                pass

    out = {}
    for etiqueta, mapa in pares.items():
        serie = [mapa.get(p) for p in periodos]
        if any(x is not None for x in serie):
            out[etiqueta] = serie
            print(f"    [{etiqueta}]: "
                  f"{sum(1 for x in serie if x is not None)}/{len(serie)} meses")
    return out


CONSUMO_LOCALDATE = "LocalDateTable_98d99889-e9ac-4e68-a26f-5b9fd4115987"

# Filtros del visual "MIP / TN VENDIDA" del reporte '14. Consumo Materiales
# indirectos de produccion' (Copiar consulta, 2026-09-06).
#
# El tercero es redundante tal como lo genera Power BI —pide TIPO DE OPERACION
# en {"Costo"} Y ADEMÁS en {"Costo","Gasto"}, que equivale a solo "Costo"— pero
# se deja textual: simplificarlo sería reescribir la consulta del reporte, y no
# hay forma de comprobar desde aquí que el modelo lo interprete igual.
#
# __ANIO__ es el único parámetro (el segmentador del reporte).
_CONSUMO_FILTROS = """	VAR __DS0FilterTable =
		FILTER(
			KEEPFILTERS(VALUES('Maestra de Kardex (Total)'[cuenta_analitica])),
			AND(
				NOT('Maestra de Kardex (Total)'[cuenta_analitica] IN {BLANK()}),
				NOT(
					'Maestra de Kardex (Total)'[cuenta_analitica] IN {BLANK(),
						"[941002] CONTROL INTERNO",
						"ALMACEN ATE",
						"ALMACEN PACHACAMAC"}
				)
			)
		)

	VAR __DS0FilterTable2 =
		TREATAS({__ANIO__}, 'Calendario'[Año])

	VAR __DS0FilterTable3 =
		FILTER(
			KEEPFILTERS(VALUES('PLANTA POR CECOS'[TIPO DE OPERACION])),
			AND(
				'PLANTA POR CECOS'[TIPO DE OPERACION] IN {"Costo"},
				'PLANTA POR CECOS'[TIPO DE OPERACION] IN {"Costo",
					"Gasto"}
			)
		)

	VAR __DS0FilterTable4 =
		FILTER(
			KEEPFILTERS(VALUES('Maestra de Kardex (Total)'[categoria_hijo])),
			NOT('Maestra de Kardex (Total)'[categoria_hijo] IN {"ACUERDOS COMERCIALES"})
		)

	VAR __DS0FilterTable5 =
		FILTER(
			KEEPFILTERS(VALUES('Maestra de Kardex (Total)'[CUENTA ORIGEN])),
			NOT('Maestra de Kardex (Total)'[CUENTA ORIGEN] IN {BLANK()})
		)
"""


# Filtros del visual "MIP / TN PRODUCIDA" (Copiar consulta, 2026-09-06).
#
# Son los mismos cinco del visual de TN VENDIDA, pero en OTRO orden y con el
# de cuenta_analitica escrito más simple: aquí es un solo NOT, mientras que
# allá va envuelto en un AND con un NOT(IN {BLANK()}) previo que no cambia
# nada. Equivalen, pero cada bloque se escribe como lo genera su propio
# visual en vez de reutilizar el otro.
_CONSUMO_FILTROS_PROD = """	VAR __DS0FilterTable =
		TREATAS({__ANIO__}, 'Calendario'[Año])

	VAR __DS0FilterTable2 =
		FILTER(
			KEEPFILTERS(VALUES('PLANTA POR CECOS'[TIPO DE OPERACION])),
			AND(
				'PLANTA POR CECOS'[TIPO DE OPERACION] IN {"Costo"},
				'PLANTA POR CECOS'[TIPO DE OPERACION] IN {"Costo",
					"Gasto"}
			)
		)

	VAR __DS0FilterTable3 =
		FILTER(
			KEEPFILTERS(VALUES('Maestra de Kardex (Total)'[categoria_hijo])),
			NOT('Maestra de Kardex (Total)'[categoria_hijo] IN {"ACUERDOS COMERCIALES"})
		)

	VAR __DS0FilterTable4 =
		FILTER(
			KEEPFILTERS(VALUES('Maestra de Kardex (Total)'[CUENTA ORIGEN])),
			NOT('Maestra de Kardex (Total)'[CUENTA ORIGEN] IN {BLANK()})
		)

	VAR __DS0FilterTable5 =
		FILTER(
			KEEPFILTERS(VALUES('Maestra de Kardex (Total)'[cuenta_analitica])),
			NOT(
				'Maestra de Kardex (Total)'[cuenta_analitica] IN {BLANK(),
					"[941002] CONTROL INTERNO",
					"ALMACEN ATE",
					"ALMACEN PACHACAMAC"}
			)
		)
"""


def _q_consumo_mip(anio, medida_tbl, medida, alias, filtros=None):
    """Serie mensual de un ratio de MIP (materiales indirectos de planta).

    Copiar consulta 2026-09-06. Solo se quita el TOPN(1001), que limita filas
    sin cambiar valores. La medida va como parámetro porque el reporte tiene
    varios visuales con los mismos filtros y distinta medida.
    """
    ld = CONSUMO_LOCALDATE
    return (
        "DEFINE\n"
        + (filtros or _CONSUMO_FILTROS).replace("__ANIO__", str(anio))
        + "\nEVALUATE\n"
        "\tSUMMARIZECOLUMNS(\n"
        "\t\t'Calendario'[Mes Corto],\n"
        "\t\t'Calendario'[Mes Numero],\n"
        f"\t\t'{ld}'[Año],\n"
        "\t\t__DS0FilterTable,\n"
        "\t\t__DS0FilterTable2,\n"
        "\t\t__DS0FilterTable3,\n"
        "\t\t__DS0FilterTable4,\n"
        "\t\t__DS0FilterTable5,\n"
        f"\t\t\"{alias}\", '{medida_tbl}'[{medida}]\n"
        "\t)\n\n"
        f"ORDER BY\n\t'{ld}'[Año], 'Calendario'[Mes Numero]"
    )


def serie_consumo(token, ds_id, periodos, medidas):
    """Series mensuales del reporte de consumo de materiales indirectos.

    `medidas` es una lista de (etiqueta, tabla, medida, bloque_de_filtros).
    Cada visual trae su propio bloque: los de TN vendida y TN producida
    llevan los mismos cinco filtros pero escritos distinto.
    """
    anios = sorted({a for a, _ in periodos})
    out = {}
    for etiqueta, tbl, medida, filtros in medidas:
        alias = "v_" + re.sub(r"[^A-Za-z0-9]", "_", etiqueta)
        mapa = {}
        for anio in anios:
            q = _q_consumo_mip(anio, tbl, medida, alias, filtros)
            for r in dax(token, ds_id, q, f"consumo-{etiqueta}-{anio}") or []:
                mes = r.get("Calendario[Mes Numero]") or r.get("[Mes Numero]")
                if isinstance(mes, str):
                    mes = MESES_CORTOS.get(mes.strip().lower()[:3])
                if mes is None:
                    corto = str(r.get("Calendario[Mes Corto]") or "").strip().lower()[:3]
                    mes = MESES_CORTOS.get(corto)
                v = r.get(f"[{alias}]", r.get(alias))
                if mes is None or v is None:
                    continue
                try:
                    mapa[(int(anio), int(mes))] = float(v)
                except (TypeError, ValueError):
                    pass
        serie = [mapa.get(p) for p in periodos]
        if any(x is not None for x in serie):
            out[etiqueta] = serie
            print(f"    [{etiqueta}]: "
                  f"{sum(1 for x in serie if x is not None)}/{len(serie)} meses")
        else:
            print(f"    ✗ {etiqueta}: sin datos ({medida})")
    return out


PRODUCTIVIDAD_LOCALDATE = "LocalDateTable_98d99889-e9ac-4e68-a26f-5b9fd4115987"

# Filtros del visual "PLANILLA (S/.) ENTRE KG PRODUCIDO" del reporte
# '13. Reporte de productividad por funcionario' (Copiar consulta, 2026-09-06).
#
# Ocho filtros. Los cuatro primeros acotan qué gasto cuenta como planilla
# operativa ('Exl Cuenta Contables'), y los tres siguientes son las mismas
# exclusiones de facturación que usa el reporte de Margen. El octavo es la
# ventana móvil desde el 31/07/2025.
#
# 'Calendario'[MesActual] = "Otros" excluye el mes en curso: el reporte no
# grafica el mes incompleto en esta serie.
_PRODUCTIVIDAD_FILTROS = """	VAR __DS0FilterTable =
		TREATAS({"Otros"}, 'Calendario'[MesActual])

	VAR __DS0FilterTable2 =
		TREATAS({"AMBAS",
			"GASTO DE PERSONAL OPERATIVO"}, 'Exl Cuenta Contables'[CLASIFICACION])

	VAR __DS0FilterTable3 =
		TREATAS({"GASTO DE PERSONAL"}, 'Exl Cuenta Contables'[GASTO_PERSONAL])

	VAR __DS0FilterTable4 =
		TREATAS({"Gasto de Personal Operativo"}, 'Exl Cuenta Contables'[SUB_CATEGORIA])

	VAR __DS0FilterTable5 =
		FILTER(
			KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[categoria_producto])),
			NOT(
				'Maestra de Facturacion (Total)'[categoria_producto] IN {"BONIFICACION Y REBATES",
					"CHATARRA",
					"INTERESES",
					"MATERIA PRIMA",
					"SERVICIOS",
					"SUMINISTROS",
					BLANK()}
			)
		)

	VAR __DS0FilterTable6 =
		FILTER(
			KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[producto])),
			NOT(
				'Maestra de Facturacion (Total)'[producto] IN {"PAVO C/M C/ASA EP CONG (8 KG)",
					"ALIMENTACION COMERCIAL"}
			)
		)

	VAR __DS0FilterTable7 =
		FILTER(
			KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[estado])),
			NOT('Maestra de Facturacion (Total)'[estado] IN {"Cancelado"})
		)

	VAR __DS0FilterTable8 =
		FILTER(
			KEEPFILTERS(VALUES('Calendario'[Date])),
			'Calendario'[Date] >= (DATE(2025, 7, 31) + TIME(0, 0, 1))
		)
"""

# Los dos visuales de productividad (KG PRODUCIDO y KG VENDIDO) comparten los
# ocho filtros — verificado por diff contra ambas consultas, no asumido — y se
# diferencian solo en las medidas. Por eso van como dos listas y una sola
# plantilla de consulta.
PRODUCTIVIDAD_MEDIDAS = [
    ("SUM_PROD",       "Producción (kg)",   "IGNORE('Maestra de Kardex (Total)'[SUM PROD])"),
    ("GASTO_TOTAL",    "Planilla (S/)",     "IGNORE('Exl Cuenta Contables'[GASTO TOTAL])"),
    ("PROD_OPERATIVA", "Planilla / kg producido", "'Exl Cuenta Contables'[PROD OPERATIVA]"),
]

PRODUCTIVIDAD_MEDIDAS_VENDIDO = [
    ("SumPeso_total_K", "Venta (kg)",
     "IGNORE(\n\t\t\tCALCULATE(SUM('Maestra de Facturacion (Total)'[Peso total K]))\n\t\t)"),
    ("VENTA_KG_X_SOL",  "Kg vendidos por sol de planilla",
     "'Exl Cuenta Contables'[VENTA KG X SOL]"),
]


def _q_productividad(medidas=None, sin_mes_actual=True):
    """Serie mensual de productividad: planilla, producción y su ratio.

    Copiar consulta 2026-09-06. Solo se quita el TOPN(1001), que limita filas
    sin cambiar valores. No lleva filtro de año — [Año] es columna de
    agrupación — así que una sola consulta cubre todo el histórico.
    """
    ld = PRODUCTIVIDAD_LOCALDATE
    # __DS0FilterTable es el 'Calendario'[MesActual] = "Otros", que excluye el
    # mes en curso porque el reporte no grafica el mes incompleto. Aquí sí lo
    # queremos: sin él, Producción y Planilla por kilo se quedaban sin dato de
    # setiembre y la app mostraba dos indicadores vacíos con Power BI lleno.
    #
    # El mes en curso va incompleto y eso ya lo dice cada tarjeta con su fecha
    # de corte; ocultarlo no lo hace más completo, solo lo esconde.
    primero = 2 if sin_mes_actual else 1
    usados = "".join(f"\t\t__DS0FilterTable{'' if i == 1 else i},\n"
                     for i in range(primero, 9))
    lista = medidas or PRODUCTIVIDAD_MEDIDAS
    medidas = "".join(f'\t\t"{alias}", {expr},\n'
                      for alias, _etq, expr in lista).rstrip(",\n") + "\n"
    return (
        "DEFINE\n"
        + _PRODUCTIVIDAD_FILTROS
        + "\nEVALUATE\n"
        "\tSUMMARIZECOLUMNS(\n"
        f"\t\t'{ld}'[Año],\n"
        "\t\t'Calendario'[Mes Corto],\n"
        "\t\t'Calendario'[Mes Numero],\n"
        + usados + medidas +
        "\t)\n\n"
        f"ORDER BY\n\t'{ld}'[Año], 'Calendario'[Mes Numero]"
    )


def serie_productividad(token, ds_id, periodos):
    """Series mensuales del reporte de productividad por funcionario."""
    out = {}
    for lista, etiqueta in ((PRODUCTIVIDAD_MEDIDAS, "producido"),
                            (PRODUCTIVIDAD_MEDIDAS_VENDIDO, "vendido")):
        out.update(_serie_productividad_una(token, ds_id, periodos, lista, etiqueta))
    return out


def _serie_productividad_una(token, ds_id, periodos, lista, etiqueta):
    filas = dax(token, ds_id, _q_productividad(lista, sin_mes_actual=False),
                f"productividad-{etiqueta}")
    if not filas:
        return {}
    ld = PRODUCTIVIDAD_LOCALDATE
    mapas = {etq: {} for _a, etq, _e in lista}
    for r in filas:
        anio = r.get(f"{ld}[Año]") or r.get("[Año]")
        mes = r.get("Calendario[Mes Numero]") or r.get("[Mes Numero]")
        if isinstance(mes, str):
            mes = MESES_CORTOS.get(mes.strip().lower()[:3])
        if mes is None:
            corto = str(r.get("Calendario[Mes Corto]") or "").strip().lower()[:3]
            mes = MESES_CORTOS.get(corto)
        if anio is None or mes is None:
            continue
        try:
            per = (int(anio), int(mes))
        except (TypeError, ValueError):
            continue
        for alias, etq, _e in lista:
            v = r.get(f"[{alias}]", r.get(alias))
            if v is None:
                continue
            try:
                mapas[etq][per] = float(v)
            except (TypeError, ValueError):
                pass

    out = {}
    for _a, etq, _e in lista:
        serie = [mapas[etq].get(p) for p in periodos]
        if any(x is not None for x in serie):
            out[etq] = serie
            print(f"    [{etq}]: "
                  f"{sum(1 for x in serie if x is not None)}/{len(serie)} meses")
    return out


CATALOGO = Path("capturas/catalogo.json")

# Problemas de las consultas capturadas. Se vuelcan en series.json para poder
# diagnosticarlos sin acceso al log del workflow, igual que en fetch_powerbi.
DIAG_SERIES = []


def cargar_catalogo():
    """Consultas exportadas del Analizador, indexadas por nombre de visual."""
    if not CATALOGO.exists():
        return {}
    try:
        filas = json.loads(CATALOGO.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"  · catálogo ilegible ({e})")
        return {}
    por_visual = {}
    for f in filas:
        # También por "Visual#hash": dos visuales del mismo reporte se llaman
        # igual —el gráfico del ratio total y su tabla por categoría— y
        # quedarse con el de más filas descartaba justo el total.
        if f.get("visual") and f.get("hash") and f.get("dax"):
            por_visual[f"{f['visual']}#{f['hash']}"] = f
        # Si el mismo visual sale en varias pestañas, se queda el de más filas:
        # suele ser el que trae el detalle completo.
        v = f.get("visual")
        if not v or not f.get("dax"):
            continue
        prev = por_visual.get(v)
        if prev is None or (f.get("filas") or 0) > (prev.get("filas") or 0):
            por_visual[v] = f
    return por_visual


def partir_evaluates(dax):
    """Separa un DAX del Analizador en (bloque DEFINE, [consultas EVALUATE]).

    La API REST de Power BI devuelve UNA sola tabla por consulta, aunque el
    DAX traiga varios EVALUATE. Los visuales de matriz traen dos: el primero
    es el eje (los períodos) y el segundo el cuerpo con los datos — así que
    enviándolo entero solo llegaba el eje, y el cuerpo no se veía nunca.

    Cada EVALUATE se devuelve completo con su ORDER BY, listo para anteponerle
    el DEFINE y mandarlo por separado.
    """
    i = dax.find("EVALUATE")
    if i < 0:
        return "", [dax]
    define = dax[:i]
    resto = dax[i:]
    partes, actual = [], []
    for linea in resto.splitlines():
        if linea.lstrip().startswith("EVALUATE") and actual:
            partes.append("\n".join(actual).rstrip())
            actual = [linea]
        else:
            actual.append(linea)
    if actual:
        partes.append("\n".join(actual).rstrip())
    return define, partes


def tablas_de_captura(ejecutor, dax, label):
    """Ejecuta un DAX capturado y devuelve una tabla por cada EVALUATE.

    `ejecutor` recibe (consulta, etiqueta) y devuelve la lista de filas.
    """
    define, partes = partir_evaluates(dax)
    if len(partes) <= 1:
        filas = ejecutor(dax, label)
        return [filas] if filas else []
    tablas = []
    for n, parte in enumerate(partes, 1):
        filas = ejecutor(define + parte, f"{label}#{n}")
        tablas.append(filas or [])
    return tablas

def dax_crudo(token, dataset_id, query, label):
    """Ejecuta DAX y devuelve TODAS las tablas del resultado.

    Las consultas del Analizador traen dos EVALUATE: el primero es el eje
    (los períodos) y el segundo el cuerpo con los datos. dax() se queda con la
    primera tabla, que aquí es la equivocada.
    """
    ws = WS_POR_DATASET.get(dataset_id, WS_ID)
    url = f"{PBI_BASE}/groups/{ws}/datasets/{dataset_id}/executeQueries"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {"queries": [{"query": query}], "serializerSettings": {"includeNulls": True}}
    try:
        r = requests.post(url, json=body, headers=headers, timeout=90)
        if r.status_code != 200:
            print(f"      [{label}] HTTP {r.status_code}: {r.text[:180]}")
            DIAG_SERIES.append({"consulta": label, "http": r.status_code,
                                "error": r.text[:400]})
            return []
        return [t.get("rows", []) for t in r.json()["results"][0].get("tables", [])]
    except Exception as e:
        print(f"      [{label}] error: {e}")
        DIAG_SERIES.append({"consulta": label, "http": 0, "error": repr(e)[:300]})
        return []


# Datasets que ya contestaron, del más reciente al más antiguo.
_DATASETS_QUE_RESPONDEN = []


def dataset_de_captura(token, entrada, candidatos):
    """Averigua contra qué dataset corre una consulta capturada, probándola.

    Las exportaciones no dicen de qué modelo salieron, y deducirlo por las
    tablas que mencionan es adivinar. Se prueba en cada candidato y se queda
    el primero que devuelve filas: si la tabla no existe, Power BI responde
    error y se pasa al siguiente.
    """
    # Se empieza por los datasets que YA respondieron en esta corrida.
    #
    # Cada intento fallido cuesta una petición completa —la consulta entera
    # contra un modelo que no la tiene— y con veinte capturas y seis
    # candidatos eso es más de cien peticiones desperdiciadas. La corrida
    # #119 cruzó el límite del espacio de trabajo por esto: 23 consultas
    # devolvieron 429 y varias series llegaron vacías.
    #
    # Casi todas las capturas del mismo reporte viven en el mismo modelo, así
    # que probar primero los conocidos acierta a la primera la mayoría de las
    # veces y el resto del orden se respeta igual.
    vistos = [d for d in _DATASETS_QUE_RESPONDEN if d in candidatos]
    resto = [d for d in candidatos if d and d not in vistos]
    for ds in vistos + resto:
        if not ds:
            continue
        tablas = dax_crudo(token, ds, entrada["dax"], f"sonda:{ds[:8]}")
        if tablas and tablas[-1]:
            if ds in _DATASETS_QUE_RESPONDEN:
                _DATASETS_QUE_RESPONDEN.remove(ds)
            _DATASETS_QUE_RESPONDEN.insert(0, ds)
            return ds, tablas
    return None, None


def serie_desde_captura(token, ds_id, periodos, visual, col_dim, col_medida,
                        prefijo, catalogo=None):
    """Series mensuales a partir de una consulta EXPORTADA del Analizador.

    Se envía el DAX exactamente como lo generó Power BI — sin reescribirlo.
    Transcribir consultas a mano fue la fuente de varios errores en este
    proyecto (una medida con doble M, un filtro de más, un grupo de menos);
    ejecutando la captura literal ese riesgo desaparece.

    `col_dim` es la columna por la que se abre la serie (Categoria, Canal…) y
    `col_medida` el alias de la medida en el resultado. El período se arma con
    las columnas de la tabla de fechas automática que traen estas consultas.
    """
    cat = catalogo if catalogo is not None else cargar_catalogo()
    entrada = cat.get(visual)
    if not entrada:
        print(f"    · '{visual}': no está en el catálogo — falta exportar esa pestaña")
        return {}

    # Una consulta por EVALUATE: la API devuelve una sola tabla por llamada.
    tablas = tablas_de_captura(
        lambda q, lb: (dax_crudo(token, ds_id, q, lb) or [[]])[0],
        entrada["dax"], f"captura:{visual}")
    if not tablas:
        return {}

    def busca(fila, sufijo):
        for k, v in fila.items():
            if k.endswith(sufijo):
                return v
        return None

    # No siempre el cuerpo es la última tabla: en los visuales de matriz
    # jerárquica la última puede ser el nivel superior, sin la dimensión que
    # interesa. Se elige la que traiga la medida (y la dimensión, si aplica).
    filas, mejor_n = [], 0
    for t in tablas:
        if not t:
            continue
        n = (1 if busca(t[0], f"[{col_medida}]") is not None else 0)
        if col_dim and busca(t[0], f"[{col_dim}]") is not None:
            n += 1
        if n > mejor_n or (n == mejor_n and n and len(t) > len(filas)):
            filas, mejor_n = t, n
    if not mejor_n:
        disponibles = sorted({k for t in tablas for f in t[:5] for k in f})
        print(f"    · '{visual}': ninguna tabla trae [{col_medida}]"
              f"{f' ni [{col_dim}]' if col_dim else ''}")
        DIAG_SERIES.append({
            "consulta": f"captura:{visual}", "http": 200,
            "error": (f"falta [{col_medida}]"
                      + (f" o [{col_dim}]" if col_dim else "")
                      + f". Claves devueltas: {disponibles[:30]}")})
        return {}

    # Los visuales de matriz usan SUBSTITUTEWITHINDEX: cambian las columnas de
    # fecha por un [ColumnIndex] y publican el eje —los períodos en orden— en
    # la PRIMERA tabla del resultado. Sin reconstruir esa correspondencia el
    # cuerpo no tiene fechas y no se puede armar ninguna serie.
    # SUBSTITUTEWITHINDEX cambia una de las dos dimensiones por un
    # [ColumnIndex] y publica sus valores en la PRIMERA tabla. Cuál de las dos
    # depende de cómo esté armado el visual:
    #
    #   · La merma por planta indexa los MESES: cada fila es una planta.
    #   · El margen por unidad de negocio indexa las UNIDADES: cada fila es un
    #     mes y trae sus columnas de fecha completas.
    #
    # Distinguirlas es simple: si la fila ya trae el año, lo indexado no son
    # los meses. Asumir siempre lo primero hacía que el margen por UEN
    # devolviera 29 filas y ninguna serie.
    eje, eje_es_dim = [], False
    if filas and any(k.endswith("[ColumnIndex]") for k in filas[0]):
        eje_es_dim = busca(filas[0], "[Año]") is not None
        for f in (tablas[0] if len(tablas) > 1 else []):
            if eje_es_dim:
                # El eje son los valores de la dimensión (B&D, TIGO, MAQUILA).
                v = busca(f, f"[{col_dim}]") if col_dim else None
                if v is None:
                    v = next((x for k, x in f.items()
                              if not k.endswith("[ColumnIndex]")
                              and isinstance(x, str)), None)
                eje.append(v)
                continue
            a_, m = busca(f, "[Año]"), busca(f, "[NroMes]")
            if m is None:
                corto = busca(f, "[Mes]")
                if isinstance(corto, str):
                    m = MESES_CORTOS.get(corto.strip().lower()[:3])
            try:
                eje.append((int(a_), int(m)))
            except (TypeError, ValueError):
                eje.append(None)
        if not eje:
            print(f"    · '{visual}': usa ColumnIndex pero no vino el eje")
            return {}

    por_dim = {}
    for f in filas:
        # col_dim None: la consulta no abre por ninguna dimensión (una sola
        # línea en el tiempo), así que todas las filas van al mismo grupo.
        dim = busca(f, f"[{col_dim}]") if col_dim else ""
        val = busca(f, f"[{col_medida}]")
        idx = busca(f, "[ColumnIndex]")
        en_rango = isinstance(idx, (int, float)) and 0 <= int(idx) < len(eje)
        if eje and eje_es_dim:
            # La dimensión viene del eje; el período, de la propia fila.
            dim = eje[int(idx)] if en_rango else None
            anio, mes = busca(f, "[Año]"), busca(f, "[NroMes]")
            if mes is None:
                mes = busca(f, "[MES]")
            if mes is None:
                corto = busca(f, "[Mes Corto]") or busca(f, "[Mes]")
                if isinstance(corto, str):
                    mes = MESES_CORTOS.get(corto.strip().lower()[:3])
            per = None
            if anio is not None and mes is not None:
                try:
                    per = (int(anio), int(mes))
                except (TypeError, ValueError):
                    per = None
        elif eje:
            per = eje[int(idx)] if en_rango else None
        else:
            anio, mes = busca(f, "[Año]"), busca(f, "[NroMes]")
            if mes is None:
                corto = busca(f, "[Mes]")
                if isinstance(corto, str):
                    mes = MESES_CORTOS.get(corto.strip().lower()[:3])
            per = None
            if anio is not None and mes is not None:
                try:
                    per = (int(anio), int(mes))
                except (TypeError, ValueError):
                    per = None
        if dim is None or val is None or per is None:
            continue
        try:
            por_dim.setdefault(str(dim), {})[per] = float(val)
        except (TypeError, ValueError):
            continue

    out = {}
    for dim in sorted(por_dim):
        serie = [por_dim[dim].get(p) for p in periodos]
        n = sum(1 for x in serie if x is not None)
        if n >= 3:
            etiqueta = f"{prefijo} {dim}".strip()
            out[etiqueta] = serie
            print(f"    [{etiqueta}]: {n}/{len(serie)} meses")
    if not out:
        muestra = filas[:2]
        print(f"    · '{visual}': sin series utilizables")
        DIAG_SERIES.append({
            "consulta": f"captura:{visual}", "http": 200,
            "error": (f"{len(filas)} filas pero ninguna serie con 3+ meses. "
                      f"eje={'sí' if eje else 'no'}. Muestra: {muestra}")})
    return out


def serie_margen(token, ds_id, periodos, medidas):
    """Series mensuales del reporte de Margen.

    `medidas` es una lista de (etiqueta, tabla, medida, bloque_de_filtros).
    Cada visual trae su propio bloque: los de TN vendida y TN producida
    llevan los mismos cinco filtros pero escritos distinto.
    """
    anios = sorted({a for a, _ in periodos})
    ld = MARGEN_LOCALDATE
    out = {}
    for etiqueta, tbl, medida in medidas:
        alias = "v_" + re.sub(r"[^A-Za-z0-9]", "_", etiqueta)
        por_periodo = {}
        for anio in anios:
            q = _q_margen(anio, tbl, medida, alias)
            for r in dax(token, ds_id, q, f"margen-{etiqueta}-{anio}") or []:
                val = r.get(f"[{alias}]", r.get(alias))
                if val is None:
                    continue          # fila del producto cruzado sin dato
                mes = r.get("Calendario[MES]") or r.get(f"{ld}[NroMes]")
                if isinstance(mes, str):
                    mes = MESES_CORTOS.get(mes.strip().lower()[:3])
                if mes is None:
                    continue
                try:
                    por_periodo[(int(anio), int(mes))] = float(val)
                except (TypeError, ValueError):
                    pass
        serie = [por_periodo.get(pp) for pp in periodos]
        if any(x is not None for x in serie):
            out[etiqueta] = serie
            print(f"    [{etiqueta}]: "
                  f"{sum(1 for x in serie if x is not None)}/{len(serie)} meses")
        else:
            print(f"    ✗ {etiqueta}: sin datos ({medida})")
    return out


def serie_mermas_exacta(token, ds_id, periodos):
    """Series mensuales de Mermas con los filtros reales del visual.

    Devuelve {nombre_medida: [valores por período]} o None.
    Se consulta un año por vez porque el visual filtra por el segmentador ANO.
    """
    por_periodo = {alias: {} for _, alias in MERMAS_MEDIDAS}
    anios = sorted({a for a, _ in periodos})
    for anio in anios:
        rows = dax(token, ds_id, _q_mermas(anio), f"mermas-{anio}")
        for r in rows or []:
            mes = r.get("Calendario[MES]") or r.get("[MES]")
            if mes is None:
                continue
            if isinstance(mes, str):
                mes = MESES_CORTOS.get(mes.strip().lower()[:3])
            if mes is None:
                continue
            for col, alias in MERMAS_MEDIDAS:
                v = r.get(f"[{col}]", r.get(col))
                if v is None:
                    continue
                try:
                    por_periodo[alias][(int(anio), int(mes))] = float(v)
                except (TypeError, ValueError):
                    pass

    out = {}
    for _, alias in MERMAS_MEDIDAS:
        serie = [por_periodo[alias].get(p) for p in periodos]
        if any(v is not None for v in serie):
            out[alias] = serie
    return out or None


def serie_mermas_plantas(token, ds_id, periodos):
    """Series mensuales de % Merma por planta.

    Solo las de PLANTAS_MERMAS: cada una con su almacén y su medida, ambos
    confirmados por Copiar consulta. Ver la nota de esa constante.
    """
    anios = sorted({a for a, _ in periodos})
    out = {}
    for etiqueta, almacenes, medida, tipo_base in PLANTAS_MERMAS:
        alias = "v_" + re.sub(r"[^A-Za-z0-9]", "_", etiqueta)
        por_periodo = {}
        for anio in anios:
            q = _q_mermas_planta(anio, almacenes, medida, alias, tipo_base)
            for r in dax(token, ds_id, q, f"mermas-{etiqueta}-{anio}") or []:
                mes = r.get("Calendario[MES]") or r.get("[MES]")
                if isinstance(mes, str):
                    mes = MESES_CORTOS.get(mes.strip().lower()[:3])
                val = r.get(f"[{alias}]", r.get(alias))
                if mes is None or val is None:
                    continue
                try:
                    por_periodo[(int(anio), int(mes))] = float(val)
                except (TypeError, ValueError):
                    pass
        serie = [por_periodo.get(pp) for pp in periodos]
        # Las tres medidas de planta son porcentajes pero no vienen en la misma
        # escala: '% Merma total MAQUILA' devuelve 1.86 donde sus hermanas
        # devuelven 0.0303. Publicadas juntas, la app mostraba Planta Terceros
        # en 185.6% y la ponía a la cabeza de todos los problemas de merma.
        # Una merma por encima del 50% no existe: lo que pase de ahí está en
        # escala de porcentaje y se normaliza a fracción.
        vistos = [x for x in serie if x is not None]
        if vistos and max(abs(x) for x in vistos) > 0.5:
            serie = [None if x is None else x / 100 for x in serie]
            print(f"    · [% Merma {etiqueta}] venía en escala de porcentaje; "
                  f"normalizada a fracción")
            DIAG_SERIES.append({
                "consulta": f"mermas:{etiqueta}", "http": 200,
                "error": f"la medida '{medida}' devuelve porcentaje (máx "
                         f"{max(abs(x) for x in vistos):.4f}) mientras las otras "
                         f"plantas devuelven fracción; se normalizó dividiendo "
                         f"entre 100"})
        if any(x is not None for x in serie):
            out[f"% Merma {etiqueta}"] = serie
            print(f"    [% Merma {etiqueta}]: "
                  f"{sum(1 for x in serie if x is not None)}/{len(serie)} meses "
                  f"({'+'.join(almacenes)} / {medida})")
        else:
            print(f"    ✗ % Merma {etiqueta}: sin datos")
    return out


def serie_mermas_segmentos(token, ds_id, periodos):
    """Series mensuales de % Merma por unidad de negocio.

    Solo los segmentos de SEGMENTOS_MERMAS: cada uno con la medida y el juego
    de filtros que realmente usa su visual. Ver la nota de esa constante sobre
    por qué se retiró el sondeo por nombre.
    """
    anios = sorted({a for a, _ in periodos})
    out = {}
    for etiqueta, medida, tipo_base in SEGMENTOS_MERMAS:
        alias = "v_" + re.sub(r"[^A-Za-z0-9]", "_", etiqueta)
        por_periodo = {}
        for anio in anios:
            q = _q_mermas_segmento(anio, "Tabla Mermas", medida, alias, tipo_base)
            for r in dax(token, ds_id, q, f"mermas-{etiqueta}-{anio}") or []:
                mes = r.get("Calendario[MES]") or r.get("[MES]")
                if isinstance(mes, str):
                    mes = MESES_CORTOS.get(mes.strip().lower()[:3])
                val = r.get(f"[{alias}]", r.get(alias))
                if mes is None or val is None:
                    continue
                try:
                    por_periodo[(int(anio), int(mes))] = float(val)
                except (TypeError, ValueError):
                    pass
        serie = [por_periodo.get(pp) for pp in periodos]
        # Las tres medidas de planta son porcentajes pero no vienen en la misma
        # escala: '% Merma total MAQUILA' devuelve 1.86 donde sus hermanas
        # devuelven 0.0303. Publicadas juntas, la app mostraba Planta Terceros
        # en 185.6% y la ponía a la cabeza de todos los problemas de merma.
        # Una merma por encima del 50% no existe: lo que pase de ahí está en
        # escala de porcentaje y se normaliza a fracción.
        vistos = [x for x in serie if x is not None]
        if vistos and max(abs(x) for x in vistos) > 0.5:
            serie = [None if x is None else x / 100 for x in serie]
            print(f"    · [% Merma {etiqueta}] venía en escala de porcentaje; "
                  f"normalizada a fracción")
            DIAG_SERIES.append({
                "consulta": f"mermas:{etiqueta}", "http": 200,
                "error": f"la medida '{medida}' devuelve porcentaje (máx "
                         f"{max(abs(x) for x in vistos):.4f}) mientras las otras "
                         f"plantas devuelven fracción; se normalizó dividiendo "
                         f"entre 100"})
        if any(x is not None for x in serie):
            out[f"% Merma {etiqueta}"] = serie
            print(f"    [% Merma {etiqueta}]: "
                  f"{sum(1 for x in serie if x is not None)}/{len(serie)} meses "
                  f"({medida})")
        else:
            print(f"    ✗ % Merma {etiqueta}: sin datos ({medida})")
    return out



# KPIs que deben salir de una serie capturada y no del sondeo genérico.
# (reporte, etiqueta del KPI, dataset, serie, formato, meta)
#
# Nace de una contradicción real: el reporte de Margen mostraba 49.5% de
# margen y S/2.96 de costo por kilo —del sondeo, sin los filtros del visual—
# mientras su propia serie capturada decía 46.48% y S/3.16. Dos cifras para lo
# mismo en la misma pantalla.
RECONCILIAR = [
    ("margen_variable", "Margen variable", "margen", "% Margen Variable", "pct", "≥52%"),
    ("margen_variable", "Costo/kg",        "margen", "Costo x Kilo",      "soles", None),
    ("margen_variable", "Precio/kg",       "margen", "Precio x Kilo",     "soles", None),
    ("margen_variable", "Ventas mes",      "margen", "Ventas (S/.)",      "soles", None),
    ("cuentas_por_pagar", "Días CxP",      "cxp",    "Días CxP",          "dias",  "<90d"),
    # Mermas no tiene tarjeta de KPI —su página son todos gráficos— así que el
    # valor sale del último mes de la serie capturada, que es la misma cifra
    # que dibuja el gráfico del reporte.
    ("mermas", "Merma Total", "mermas", "% Merma Total", "pct", "<2%"),
    # Consumo: el reporte muestra el ratio del mes, no un acumulado.
    ("consumo_materiales", "Costo x TN Vendida",   "consumo", "MIP / TN Vendida",   "soles", None),
    ("consumo_materiales", "Costo x TN Producida", "consumo", "MIP / TN Producida", "soles", None),
    ("productividad", "Planilla S/. / KG Producido", "productividad",
     "Planilla / kg producido", "soles", None),
    ("productividad", "Planilla S/. / KG Vendido", "productividad",
     "Kg vendidos por sol de planilla", "soles", None),
    # Reporte 12: los tres KPIs de Fill Rate que quedaban sin contrastar.
    ("fill_rate", "Facturación",  "fill_rate", "Facturación",    "soles", None),
    ("fill_rate", "Orden de Venta", "fill_rate", "Orden de Venta", "soles", None),
]


def _ultimo(serie, periodos, saltar_parcial=True):
    """Último valor de la serie. Por defecto ignora el mes en curso.

    Desde que las series llegan hasta el mes en curso, tomar "el último"
    significaba tomar un mes a medias: el tablero titulado "Agosto" pasó a
    mostrar el margen de trece días de setiembre (44.0% en vez de 46.5%) y las
    ventas del mes en S/1.62M en vez de S/7.69M. Las tarjetas son del último
    mes CERRADO; el mes en curso tiene su propia pestaña.
    """
    ini = len(serie) - 1
    if saltar_parcial and periodos and ini < len(periodos) and periodos[ini] == PERIODO_PARCIAL:
        ini -= 1
    for i in range(ini, -1, -1):
        if serie[i] is not None:
            return serie[i], (periodos[i] if i < len(periodos) else None)
    return None, None


def _formatear(valor, formato):
    if formato == "pct":
        v = valor * 100 if abs(valor) <= 1.5 else valor
        return f"{v:.1f}%", v
    if formato == "dias":
        return f"{valor:.0f}d", valor
    if abs(valor) >= 1_000_000:
        return f"S/{valor / 1_000_000:.2f}M", valor
    if abs(valor) >= 1_000:
        return f"S/{valor:,.0f}", valor
    return f"S/{valor:.2f}", valor


def reconciliar_kpis(resultado, periodos_str):
    """Hace que los KPIs del dashboard coincidan con las series capturadas.

    fetch_powerbi corre antes que este script y construye los KPIs con el
    sondeo genérico, que consulta las medidas SIN los filtros del visual. Donde
    existe una serie capturada del mismo concepto, esa manda: es la que
    reproduce lo que ve el usuario en Power BI.

    Se anota el mes del dato en `meta`, porque un KPI sin período invita a
    leerlo como "hoy" cuando es el último mes cerrado.
    """
    ruta = OUTPUT_DIR / "summaries.json"
    if not ruta.exists():
        return
    try:
        datos = json.loads(ruta.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"  · no se pudo reconciliar summaries.json: {e}")
        return

    cambios = 0
    for empresa, edatos in (datos.get("empresas") or {}).items():
        reportes = edatos.get("reportes") or {}
        # Cifras que fetch_powerbi leyó ejecutando la tarjeta del reporte. No
        # se tocan: la tarjeta es lo que el usuario ve en Power BI, y la serie
        # mensual mide otra cosa. Costo x TN Producida se publicaba en 151.64
        # —el mes, según la serie— cuando la tarjeta marca 143.46.
        de_tarjeta = {tuple(x) for x in (datos.get("kpis_de_tarjeta") or [])}

        for tipo, etiqueta, ds, serie_nom, formato, meta in RECONCILIAR:
            if (tipo, etiqueta) in de_tarjeta:
                print(f"    = [{tipo}] {etiqueta}: se respeta el valor de la tarjeta")
                continue
            rep = reportes.get(tipo)
            serie = (resultado.get(ds) or {}).get(serie_nom)
            if not rep or not serie:
                continue
            valor, periodo = _ultimo(serie, periodos_str)
            if valor is None:
                continue
            texto, _ = _formatear(valor, formato)

            # Aviso de mes incompleto: si el último valor se sale mucho de lo
            # que venían siendo los meses anteriores, casi siempre es un corte
            # parcial y no una mejora. Pasa con la planilla por kilo, que en el
            # mes en curso cae a S/0.04 desde S/0.50 porque solo hay unos días
            # cargados — leído sin contexto parece una mejora de diez veces.
            previos = [x for x in serie[:-1] if x is not None][-6:]
            aviso = None
            if len(previos) >= 3:
                previos_ord = sorted(previos)
                mediana = previos_ord[len(previos_ord) // 2]
                if mediana and (valor < mediana * 0.3 or valor > mediana * 3):
                    aviso = f"{periodo} parcial — mediana previa "
                    aviso += _formatear(mediana, formato)[0]
            kpis = rep.setdefault("kpis", [])
            actual = next((k for k in kpis if k["label"] == etiqueta), None)
            if actual is None:
                # fuente='reporte': viene de una serie capturada. Sin esto
                # queda marcado como no contrastado, que es justo lo contrario.
                kpis.append({"label": etiqueta, "valor": texto,
                             "meta": aviso or meta or (periodo or ""),
                             "fuente": "reporte",
                             **({"estado": "yellow"} if aviso else {})})
                cambios += 1
                print(f"    + [{tipo}] {etiqueta} = {texto} ({periodo})")
            elif actual.get("valor") != texto:
                print(f"    ~ [{tipo}] {etiqueta}: {actual['valor']} → {texto} "
                      f"({periodo}, de la consulta capturada)")
                # La alerta del reporte es texto libre que incluye el valor
                # viejo ("Merma 4.87% — sobre meta 2%"). Si se corrige el KPI
                # y no la alerta, la misma tarjeta muestra dos cifras
                # distintas del mismo indicador.
                anterior = actual.get("valor")
                if anterior and rep.get("alerta") and anterior in rep["alerta"]:
                    rep["alerta"] = rep["alerta"].replace(anterior, texto)
                    print(f"      · alerta de {tipo} actualizada a {texto}")
                # Y si el extractor habia marcado el KPI como no publicable,
                # esa nota ya no aplica: la serie capturada trae el valor bueno.
                # Sin esto la tarjeta mostraba "2.9%" con un "sin total
                # confiable" debajo — el dato correcto desmentido por su
                # propio pie de foto.
                elif "sin total confiable" in (rep.get("alerta") or ""):
                    rep["alerta"] = None
                    print(f"      · aviso de 'sin total confiable' retirado en {tipo}")
                if actual.get("meta", "").startswith("el total del reporte no cuadra"):
                    actual["meta"] = meta or (periodo or "")
                actual["valor"] = texto
                actual["fuente"] = "reporte"
                if aviso:
                    actual["meta"] = aviso
                    actual["estado"] = "yellow"
                elif meta:
                    actual["meta"] = meta
                cambios += 1
    # Venta Perdida = Orden de Venta − Facturación. El visual mensual del
    # reporte 12 no la trae como medida, pero sí sus dos componentes, y esa
    # resta es su definición. Quedaba el valor viejo del sondeo —S/31.37M,
    # cinco veces mayor que la propia orden de venta— marcado como verificado.
    for empresa, edatos in (datos.get("empresas") or {}).items():
        fr = (edatos.get("reportes") or {}).get("fill_rate")
        if not fr:
            continue
        kp = {k["label"]: k for k in fr.get("kpis", [])}
        ov, fa = kp.get("Orden de Venta"), kp.get("Facturación")
        if not (ov and fa and ov.get("fuente") == "reporte"
                and fa.get("fuente") == "reporte"):
            continue

        def _n(t):
            t = str(t).replace("S/", "").replace(",", "")
            return float(t.rstrip("M")) * (1e6 if t.endswith("M") else 1)

        try:
            perdida = _n(ov["valor"]) - _n(fa["valor"])
        except ValueError:
            continue
        if perdida < 0:
            continue
        texto, _ = _formatear(perdida, "soles")
        vp = kp.get("Venta Perdida")
        if vp is None:
            fr["kpis"].append({"label": "Venta Perdida", "valor": texto,
                               "meta": "orden de venta − facturación",
                               "fuente": "reporte"})
        else:
            if vp.get("valor") != texto:
                print(f"    ~ [fill_rate] Venta Perdida: {vp['valor']} → {texto} "
                      f"(orden de venta − facturación)")
            vp["valor"] = texto
            vp["meta"] = "orden de venta − facturación"
            vp["fuente"] = "reporte"
        cambios += 1

    if cambios:
        ruta.write_text(json.dumps(datos, ensure_ascii=False, indent=2))
        print(f"  ✓ {cambios} KPI(s) reconciliados con las series capturadas")


def main():
    print("=== JUANITO — SERIE HISTÓRICA MENSUAL ===\n")
    token = get_token()
    print("Token OK\n")

    cache_file = OUTPUT_DIR / "measure_cache.json"
    if not cache_file.exists():
        print("ERROR: falta measure_cache.json — corre fetch_powerbi.py primero.")
        return
    cache = json.loads(cache_file.read_text())

    periodos = build_periodos()
    periodos_str = [f"{a}-{m:02d}" for a, m in periodos]
    print(f"Períodos: {periodos_str[0]} → {periodos_str[-1]} ({len(periodos)} meses)\n")

    resultado = {}
    fechas_usadas = {}
    sin_serie = {}

    # ── CxC: consulta exacta del gráfico (la sonda genérica no la alcanza)
    cxc_id = DATASET_IDS.get("cxc")
    if cxc_id:
        print("── cxc (consulta exacta del gráfico '% DE MOROSIDAD')")
        try:
            s = serie_cxc_morosidad(token, cxc_id, periodos)
            if s:
                con_dato = sum(1 for v in s if v is not None)
                resultado.setdefault("cxc", {})["% Morosidad"] = s
                fechas_usadas["cxc"] = f"{CXC_DATE_TBL}[Año] + 'Calendario (FACT)'[Mes Numero]"
                print(f"    [% Morosidad]: {con_dato}/{len(s)} meses con dato — "
                      f"medida '[% MOROSIDAD x mes]'\n")
            else:
                print("    Sin filas — la consulta no devolvió datos\n")
        except Exception as e:
            print(f"    ✗ {e}\n")

    # ── Mermas: consulta exacta del gráfico, con los 6 filtros del visual.
    # La sonda genérica sí devolvía una serie, pero SIN esos filtros: incluía
    # chatarra, materia prima, servicios y comprobantes cancelados, así que
    # el % no coincidía con el que ve el equipo en el reporte.
    mermas_id = DATASET_IDS.get("mermas")
    if mermas_id:
        print("── mermas (consulta exacta del gráfico 'MERMA MENSUAL POR PLANTA')")
        try:
            s = serie_mermas_exacta(token, mermas_id, periodos)
            if s:
                resultado["mermas"] = s
                fechas_usadas["mermas"] = "Calendario[Año] + Calendario[MES] (6 filtros del visual)"
                for med, vals in s.items():
                    print(f"    [{med}]: {sum(1 for v in vals if v is not None)}/{len(vals)} meses")
                # Segmentos (B&D, TIGO, MAQUILA, plantas...): mismo bloque de
                # filtros menos el de TIPO DE BASE, con la medida por segmento.
                resultado["mermas"].update(serie_mermas_segmentos(token, mermas_id, periodos))
                # Plantas: la medida base filtrada por 'Tabla Mermas'[almacen]
                resultado["mermas"].update(serie_mermas_plantas(token, mermas_id, periodos))
                print()
            else:
                print("    Sin filas\n")
        except Exception as e:
            print(f"    ✗ {e}\n")

    # ── Margen: consulta exacta del visual, con sus 11 filtros
    margen_id = DATASET_IDS.get("margen")
    if margen_id:
        print("── margen (consulta exacta de 'MARGEN VARIABLE PAUNO')")
        try:
            s = serie_margen(token, margen_id, periodos, [
                ("% Margen Variable", "Medidas Julito", "% Margen Contribución. %"),
            ])
            if s:
                resultado.setdefault("margen", {}).update(s)
            # Precio y costo por kilo: otro visual, otros 12 filtros.
            # Sustituyen al 'Costo x Kilo' de la sonda genérica, que venía
            # sin ningún filtro de negocio.
            resultado.setdefault("margen", {}).update(
                serie_margen_precio_costo(token, margen_id, periodos))
            # Ventas: mismos 12 filtros (verificado por diff), otra columna
            resultado["margen"].update(
                serie_margen_ventas(token, margen_id, periodos))
            print()
        except Exception as e:
            print(f"    ✗ {e}\n")

    # ── Fill Rate: otro workspace, y el dataset se resuelve desde el reporte
    print("── fill_rate (consulta exacta de 'FILLRATE POR MES')")
    try:
        fr_id = dataset_de_reporte(token, WS_FILLRATE, FILLRATE_REPORT_ID)
        if fr_id:
            WS_POR_DATASET[fr_id] = WS_FILLRATE
            print(f"    · dataset {fr_id} en workspace FILLRATE")
            # La consulta de "FILLRATE POR MES" se dejó de pedir.
            #
            # Agrupa por 'CALENDARIO'[MES] sin el año, y ese calendario solo
            # llega a 2025-04: sus puntos caían en el tramo viejo de un eje
            # que llega a 2026-09. Se suprimían antes de publicar, así que su
            # único efecto era dejar dos avisos idénticos cada día.
            #
            # No hace falta: la captura "FILL RATE (S/) - MENSUAL" agrupa por
            # año y mes de su propia tabla de fechas y ya entrega el fill
            # rate, la facturación y la orden de venta de los 15 meses, más
            # la venta perdida y la amonestación.
            pass
        print()
    except Exception as e:
        print(f"    ✗ {e}\n")

    # ── Compras: ratio consumo/compra con los filtros reales del visual
    compras_id = DATASET_IDS.get("compras")
    if compras_id:
        print("── compras (consulta exacta de 'Eficiencia de Costo de compra')")
        try:
            s = serie_compras_ratio(token, compras_id, periodos)
            if s:
                resultado["compras"] = s
            print()
        except Exception as e:
            print(f"    ✗ {e}\n")

    # ── CxP: rotación (compras cta 60 y días de pago) del gráfico real.
    # Reemplaza la serie 'Cuentas x Pagar' de la sonda genérica, que salía
    # plana porque la medida ignoraba el filtro de fecha.
    cxp_id = DATASET_IDS.get("cxp")
    if cxp_id:
        print("── cxp (consulta exacta de 'Rotación de cuentas por pagar')")
        try:
            s = serie_cxp_rotacion(token, cxp_id, periodos)
            if s:
                resultado.setdefault("cxp", {}).update(s)
            print()
        except Exception as e:
            print(f"    ✗ {e}\n")

    # ── Consumo de materiales indirectos (MIP)
    consumo_id = DATASET_IDS.get("consumo")
    if consumo_id:
        print("── consumo (consulta exacta de 'MIP / TN VENDIDA')")
        try:
            s = serie_consumo(token, consumo_id, periodos, [
                ("MIP / TN Vendida", "Maestra de Kardex (Total)",
                 "Ratio costo / kg", _CONSUMO_FILTROS),
                ("MIP / TN Producida", "Medidas",
                 "Costo x ton producida", _CONSUMO_FILTROS_PROD),
            ])
            if s:
                resultado.setdefault("consumo", {}).update(s)
            print()
        except Exception as e:
            print(f"    ✗ {e}\n")

    # ── Productividad por funcionario
    prod_id = DATASET_IDS.get("productividad")
    if prod_id:
        print("── productividad (consulta exacta de 'PLANILLA ENTRE KG PRODUCIDO')")
        try:
            s = serie_productividad(token, prod_id, periodos)
            if s:
                resultado.setdefault("productividad", {}).update(s)
            print()
        except Exception as e:
            print(f"    ✗ {e}\n")

    # ── Series desde las capturas exportadas del Analizador. La consulta va
    # literal, sin transcribir: ver serie_desde_captura().
    cat_visuales = cargar_catalogo()
    if cat_visuales and DATASET_IDS.get("margen"):
        print(f"── capturas del Analizador ({len(cat_visuales)} visuales en catálogo)")
        # (visual, dimensión, alias de la medida, prefijo, dataset destino,
        #  datasets candidatos donde puede vivir la consulta)
        CAPTURAS = [
            ("COSTO UNITARIO",  "Categoria", "C__Unit", "Costo unitario",
             "margen", ["margen"]),
            ("PRECIO UNITARIO", "Canal", "Precio_unitario", "Precio unitario",
             "margen", ["margen"]),
            # El margen por unidad de negocio, con la medida del propio
            # reporte. Hasta ahora la app lo calculaba dividiendo precio entre
            # costo y daba 65.3% donde el reporte marca 62.30%: el reporte
            # pondera por producto y el cálculo promediaba. Es la pieza que
            # faltaba para poder abrir el margen por B&D, TIGO y MAQUILA.
            ("MARGEN VARIABLE POR UNIDAD DE NEGOCIO", "TIPO DE NEGOCIO N1",
             "v__Margen_Contribución___", "% Margen", "margen", ["margen"]),
            # Le da a S&OP su primera serie mensual: hasta ahora ese reporte
            # tenía KPIs pero ninguna tendencia.
            ("EVOLUCIÓN DE INVENTARIO (S/. MM)", None, "Días_Rotación",
             "Días de rotación", "inventario",
             ["compras", "inventario", "planificacion"]),
            ("EVOLUCIÓN DE INVENTARIO (S/. MM)", None,
             "Consumo_Acumulado_Total_S__saldo", "Consumo acumulado",
             "inventario", ["compras", "inventario", "planificacion"]),
            # Reporte 12 (Cálculo de Provisiones): cierra los tres KPIs de
            # Fill Rate que seguían viniendo del sondeo.
            ("FILL RATE (S/) - MENSUAL", None, "FACTURACION", "Facturación",
             "fill_rate", ["fill_rate", "margen", "compras"]),
            ("FILL RATE (S/) - MENSUAL", None, "ORDEN_DE_VENTA",
             "Orden de Venta", "fill_rate", ["fill_rate", "margen", "compras"]),
            ("FILL RATE (S/) - MENSUAL", None, "v__FillRate",
             "% Fill Rate mensual", "fill_rate",
             ["fill_rate", "margen", "compras"]),
            # La misma captura ya trae estas dos y no cuestan una consulta
            # más. La venta perdida dice en soles lo que el fill rate dice en
            # porcentaje, que es como se prioriza.
            ("FILL RATE (S/) - MENSUAL", None, "SumVENTA_PERDIDA_SKU_PEDIDO",
             "Venta perdida (S/)", "fill_rate",
             ["fill_rate", "margen", "compras"]),
            ("FILL RATE (S/) - MENSUAL", None, "SumAMONESTACION",
             "Amonestación (S/)", "fill_rate",
             ["fill_rate", "margen", "compras"]),
        ]
        resuelto = {}
        for visual, dim, medida, prefijo, destino, candidatos in CAPTURAS:
            entrada = cat_visuales.get(visual)
            if not entrada:
                print(f"    · '{visual}': no está en el catálogo")
                continue
            ds = resuelto.get(visual)
            if ds is None:
                ds, _ = dataset_de_captura(
                    token, entrada, [DATASET_IDS.get(c) for c in candidatos])
                if not ds:
                    print(f"    ✗ '{visual}': no corre en {candidatos}")
                    continue
                resuelto[visual] = ds
            try:
                s = serie_desde_captura(token, ds, periodos, visual, dim,
                                        medida, prefijo, cat_visuales)
                if s:
                    resultado.setdefault(destino, {}).update(s)
            except Exception as e:
                print(f"    ✗ {visual}: {e}")
        print()

    for ds_key, medidas_dict in cache.items():
        ds_id = DATASET_IDS.get(ds_key)
        if not ds_id:
            continue
        if ds_key == "fill_rate" and "fill_rate" in resultado:
            continue
        if ds_key == "compras" and "compras" in resultado:
            continue
        if ds_key == "mermas" and "mermas" in resultado:
            # ya resuelto con la consulta exacta; la genérica solo añadiría
            # una serie sin los filtros del visual (números no comparables)
            continue
        if ds_key == "cxc" and "cxc" in resultado:
            # ya resuelto arriba con la consulta exacta; la sonda genérica
            # solo volvería a marcar '% Morosidad' como plana
            continue
        medidas = list(medidas_dict.keys())
        if not medidas:
            continue

        print(f"── {ds_key} ({len(medidas)} medidas)")
        tbl, col = detectar_fecha(token, ds_id, medidas[0])
        if not tbl:
            # Deja constancia de QUÉ contiene el modelo, para no volver a adivinar
            print("    Sin tabla de fecha que filtre las medidas — diagnosticando el modelo...")
            diag = {"medidas": medidas, "tablas": [], "columnas_fecha": []}
            t_rows = dax(token, ds_id, "EVALUATE INFO.TABLES()", "diag-tables", retries=1)
            if t_rows:
                diag["tablas"] = sorted({(r.get("[Name]") or r.get("Name") or "") for r in t_rows} - {""})
            diag["columnas_fecha"] = [f"{t}[{c}]" for t, c in columnas_fecha_reales(token, ds_id)]
            if not t_rows:
                # DMV bloqueada: cae al esquema REST
                r = requests.get(f"{PBI_BASE}/groups/{WS_ID}/datasets/{ds_id}/tables",
                                 headers={"Authorization": f"Bearer {token}"}, timeout=25)
                if r.ok:
                    diag["tablas"] = [t.get("name") for t in r.json().get("value", [])]
                    diag["esquema_rest"] = {
                        t.get("name"): [c.get("name") for c in t.get("columns", [])]
                        for t in r.json().get("value", [])
                    }
                else:
                    diag["esquema_rest_error"] = f"HTTP {r.status_code}"
            sin_serie[ds_key] = diag
            print(f"    Tablas del modelo: {diag['tablas'][:15]}")
            print(f"    Columnas de fecha: {diag['columnas_fecha'][:15]}")
            print(f"    → estas medidas no admiten corte mensual con el modelo actual\n")
            continue
        fechas_usadas[ds_key] = f"{tbl}[{col}]"

        serie = serie_batch(token, ds_id, medidas, tbl, col, periodos)
        if serie is None:
            print("    Batch falló — usando mes a mes")
            serie = serie_mes_a_mes(token, ds_id, medidas, tbl, col, periodos)

        # Reporta cobertura y detecta series "planas": la medida no responde
        # al filtro de fecha genérico (CALCULATE + FILTER(ALL(Calendario)))
        # y Power BI devuelve el mismo total sin importar el mes — mismo bug
        # confirmado a mano en Control Interno/CxP/Inventario/Avance/Compras.
        # No se puede saber de antemano qué medida tiene este problema sin
        # Copiar consulta, así que se detecta por evidencia: casi todos los
        # valores no nulos son idénticos.
        serie_confiable = {}
        for med, vals in serie.items():
            con_dato = sum(1 for v in vals if v is not None)
            no_nulos = [v for v in vals if v is not None]
            # Misma regla que usa el dashboard en el navegador (3 pruebas):
            media = (sum(no_nulos) / len(no_nulos)) if no_nulos else 0
            # 1) pocos valores distintos (redondeo RELATIVO: el ruido de coma
            #    flotante no debe partir un mismo valor en dos buckets)
            distintos = len({round(v / media * 1e4) for v in no_nulos}) if media else 0
            pocas_variantes = con_dato >= 3 and distintos <= max(2, con_dato // 7)
            # 2) amplitud relativa < 1% => ruido, no tendencia
            amplitud = (max(no_nulos) - min(no_nulos)) / media if (no_nulos and media) else 0
            plana_por_amplitud = con_dato >= 3 and abs(amplitud) < 0.01
            # 3) valor congelado: una misma cifra ocupa más de la mitad de los meses
            moda = 0
            if no_nulos and media:
                conteo = {}
                for v in no_nulos:
                    k = round(v / media * 1e4)
                    conteo[k] = conteo.get(k, 0) + 1
                moda = max(conteo.values())
            congelada = con_dato >= 3 and moda > len(no_nulos) / 2
            es_plana = pocas_variantes or plana_por_amplitud or congelada
            estado = "PLANA — no confiable" if es_plana else "OK"
            print(f"    [{med}]: {con_dato}/{len(vals)} meses con dato, {distintos} valores distintos → {estado}")
            if es_plana:
                sin_serie.setdefault(ds_key, {"medidas": [], "tablas": [tbl], "columnas_fecha": [f"{tbl}[{col}]"]})
                sin_serie[ds_key].setdefault("medidas_planas", []).append({
                    "medida": med, "valores_distintos": distintos, "meses_con_dato": con_dato,
                    "amplitud_relativa_pct": round(abs(amplitud) * 100, 4),
                    "repeticiones_del_valor_mas_comun": moda,
                    "motivo": "La medida no varía por mes con el filtro CALCULATE+FILTER(ALL(Calendario)) — "
                              "posible relación inactiva o la medida ignora el filtro de fecha (confirmado con "
                              "Copiar consulta solo en Control Interno; para las demás falta validar en vivo).",
                })
            elif (ds_key, med) in SUSTITUIDAS_POR_CAPTURA:
                print(f"    [{med}]: se omite — sustituida por la serie capturada")
            else:
                serie_confiable[med] = vals

        if serie_confiable:
            # update, no asignación: 'margen' ya trae la serie de la consulta
            # exacta y asignar la pisaría. Las capturadas mandan, así que las
            # de la sonda genérica solo rellenan lo que falta.
            base = resultado.setdefault(ds_key, {})
            for med, vals in serie_confiable.items():
                base.setdefault(med, vals)
        print()

    out = {
        "generado": hoy_lima().isoformat(),
        # El ultimo periodo es el mes en curso: la app lo marca para que nadie
        # lea una caida del 60% que en realidad son doce dias contra treinta.
        "parcial": f"{hoy_lima():%Y-%m}",
        "fuente": "Power BI REST API executeQueries",
        "periodos": periodos_str,
        "tablas_fecha": fechas_usadas,
        "sin_serie": sin_serie,
        "datasets": resultado,
    }
    path = OUTPUT_DIR / "series.json"
    # json.dumps escribe Infinity/NaN, que NO son JSON válido: el navegador
    # falla en JSON.parse y la app se queda sin NINGUNA serie. Aparecen de
    # verdad — p.ej. '% Merma <almacén>' en meses sin producción divide entre
    # cero — así que se convierten a null (mes sin dato) en vez de ocultarlos.
    # allow_nan=False haría reventar el script; preferimos publicar el resto.
    if DIAG_SERIES:
        out["diagnostico"] = DIAG_SERIES
        print(f"\n⚠ {len(DIAG_SERIES)} consulta(s) con problema — "
              f"detalle en series.json → diagnostico")

    n_inf = _sanear_no_finitos(out)
    if n_inf:
        print(f"⚠ {n_inf} valores no finitos (Infinity/NaN) → null")
    # Series que quedan con menos de 3 meses útiles no sirven para tendencia
    for ds_key in list(out["datasets"]):
        for med in list(out["datasets"][ds_key]):
            vals = out["datasets"][ds_key][med]
            if sum(1 for v in vals if v is not None) < 3:
                del out["datasets"][ds_key][med]
                print(f"⚠ [{ds_key}] {med}: <3 meses con dato — se descarta")
        if not out["datasets"][ds_key]:
            del out["datasets"][ds_key]
    path.write_text(json.dumps(out, ensure_ascii=False, indent=2, allow_nan=False))
    print(f"Guardado: {path}")

    # Los KPIs de summaries.json los arma fetch_powerbi con el sondeo
    # genérico; donde hay serie capturada del mismo concepto, se corrigen.
    reconciliar_kpis(resultado, periodos_str)
    print(f"Datasets con serie: {list(resultado.keys())}")


if __name__ == "__main__":
    main()
