#!/usr/bin/env python3
"""
JUANITO — Power BI Data Fetcher v3
Estrategia: scan de medidas reales → query con nombres confirmados.
"""

import os, json, re, requests, datetime, sys
from pathlib import Path

try:
    import warnings; warnings.filterwarnings('ignore')
except: pass

TENANT_ID     = os.environ["AZURE_TENANT_ID"]
CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
USERNAME      = os.environ["PBI_USERNAME"]
PASSWORD      = os.environ["PBI_PASSWORD"]

PBI_SCOPE = "https://analysis.windows.net/powerbi/api/.default"
TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
PBI_BASE  = "https://api.powerbi.com/v1.0/myorg"

# El reporte FILLRATE vive en otro workspace (visto en su URL, 2026-09-06).
# De él solo conocemos el id del reporte; el del dataset se resuelve por API.
WS_FILLRATE = "dba94185-3df0-4310-b6b4-9c3b7bc30104"
# '11. Reporte de Planificaciones' (S&OP). Está en el workspace normal, pero
# su datasetId se resuelve por API igual que FILLRATE: solo conocemos el id
# del reporte, tomado de su URL (2026-09-06).
SOP_REPORT_ID = "85e24d53-5816-4e89-993b-91bc733fc5c4"
FILLRATE_REPORT_ID = "9641c6b7-2524-4774-b3be-b0370687cd3d"

OUTPUT_DIR = Path("data/latest")
HIST_DIR   = Path("data/historico")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
HIST_DIR.mkdir(parents=True, exist_ok=True)

def hoy_lima():
    """Fecha de hoy en Lima (UTC-5, el Perú no cambia de hora).

    El runner de GitHub corre en UTC, así que `date.today()` adelantaba el día
    a partir de las 7 de la tarde de Lima: la app llegó a mostrar "datos al
    09/09" cuando en Lima aún era el 8. Quien lee el tablero está en Lima.
    """
    return (datetime.datetime.now(datetime.timezone.utc)
            - datetime.timedelta(hours=5)).date()


HOY = hoy_lima().strftime("%Y-%m-%d")

# Mes anterior completo (para filtros DAX — septiembre sin datos completos)
_today = hoy_lima()
PREV_MONTH = _today.month - 1 if _today.month > 1 else 12
PREV_YEAR  = _today.year  if _today.month > 1 else _today.year - 1

# Tabla/columna de fecha por dataset (descubierta con discover_measures.py #8)
DATE_CONTEXT = {
    "margen":    ("Calendario", "Date"),
    "mermas":    ("Calendario", "Date"),
    "compras":   ("Calendario", "Date"),
    "fill_rate": ("Calendario", "Date"),
    # productividad_ds NO usa este mecanismo genérico — tiene su propio filtro exacto
    # de 9 dimensiones capturado con "Copiar consulta" (ver dax_productividad_pauno).
    # Aplicar aquí YEAR(Date)/MONTH(Date) sería insuficiente: el reporte también
    # filtra por Calendario[Año] (columna distinta a [Date]), MesActual, clasificación
    # contable y exclusiones de producto — confirmado que sin esos filtros el resultado
    # es hasta 38,000x distinto del real (ver commit 2026-09-03).
    "consumo": ("Calendario", "Date"),
}

def dax_productividad_pauno(token, ws_id, dataset_id, measure_name, label="prod"):
    """Ejecuta una medida del dataset '13. Productividad por Funcionario' replicando
    EXACTAMENTE los filtros de página/informe que Power BI aplica en el visual real,
    a nivel PAUNO (sin filtro de Planta). Filtros capturados el 2026-09-03 con
    Analizador de rendimiento > Copiar consulta sobre la tarjeta 'Planilla Total (S/.)'
    con año=2026 (sin mes seleccionado) — validado contra la tarjeta real (S/4,484,364).
    Cualquier cambio en los filtros del reporte en Power BI Desktop debe reflejarse aquí.
    """
    q = f"""EVALUATE
ROW(
  "v",
  CALCULATE(
    [{measure_name}],
    TREATAS({{2026}}, 'Calendario'[Año]),
    TREATAS({{"Otros"}}, 'Calendario'[MesActual]),
    TREATAS({{"AMBAS","GASTO DE PERSONAL OPERATIVO"}}, 'Exl Cuenta Contables'[CLASIFICACION]),
    TREATAS({{"GASTO DE PERSONAL"}}, 'Exl Cuenta Contables'[GASTO_PERSONAL]),
    TREATAS({{"Gasto de Personal Operativo"}}, 'Exl Cuenta Contables'[SUB_CATEGORIA]),
    FILTER(
      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[categoria_producto])),
      NOT('Maestra de Facturacion (Total)'[categoria_producto] IN
        {{"BONIFICACION Y REBATES","CHATARRA","INTERESES","MATERIA PRIMA","SERVICIOS","SUMINISTROS",BLANK()}})
    ),
    FILTER(
      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[producto])),
      NOT('Maestra de Facturacion (Total)'[producto] IN
        {{"PAVO C/M C/ASA EP CONG (8 KG)","ALIMENTACION COMERCIAL"}})
    ),
    FILTER(
      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[estado])),
      NOT('Maestra de Facturacion (Total)'[estado] IN {{"Cancelado"}})
    ),
    FILTER(
      KEEPFILTERS(VALUES('Calendario'[Date])),
      'Calendario'[Date] >= (DATE(2025, 7, 31) + TIME(0, 0, 1))
    )
  )
)"""
    rows = dax(token, ws_id, dataset_id, q, label)
    if rows:
        return rows[0].get("[v]") or rows[0].get("v")
    return None


def dax_productividad_venta_neta_kg(token, ws_id, dataset_id, label="prod_ventaneta"):
    """Venta Neta (KG) del reporte '13. Productividad por Funcionario' — PAUNO.

    No es una medida DAX: en el panel aparece como 'Suma de Peso total KG' porque
    es la COLUMNA 'Maestra de Facturacion (Total)'[Peso total K] con agregación SUM
    directa. Confirmado con Analizador de rendimiento > Copiar consulta sobre la
    tarjeta (2026-09-03) — validado contra la tarjeta real (9,834,538 con año=2026).

    Lleva los mismos 9 filtros que dax_productividad_pauno() más 2 exclusivos de
    esta tarjeta: CATEGORIZACION="VENTA BRUTA" y TIPO DE NEGOCIO N2 no vacío.
    """
    q = f"""EVALUATE
ROW(
  "v",
  CALCULATE(
    SUM('Maestra de Facturacion (Total)'[Peso total K]),
    TREATAS({{"VENTA BRUTA"}}, 'Maestra de Facturacion (Total)'[CATEGORIZACION]),
    FILTER(
      KEEPFILTERS(VALUES('TIPO DE NEGOCIO'[TIPO DE NEGOCIO N2])),
      NOT('TIPO DE NEGOCIO'[TIPO DE NEGOCIO N2] IN {{BLANK()}})
    ),
    TREATAS({{2026}}, 'Calendario'[Año]),
    TREATAS({{"Otros"}}, 'Calendario'[MesActual]),
    TREATAS({{"AMBAS","GASTO DE PERSONAL OPERATIVO"}}, 'Exl Cuenta Contables'[CLASIFICACION]),
    TREATAS({{"GASTO DE PERSONAL"}}, 'Exl Cuenta Contables'[GASTO_PERSONAL]),
    TREATAS({{"Gasto de Personal Operativo"}}, 'Exl Cuenta Contables'[SUB_CATEGORIA]),
    FILTER(
      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[categoria_producto])),
      NOT('Maestra de Facturacion (Total)'[categoria_producto] IN
        {{"BONIFICACION Y REBATES","CHATARRA","INTERESES","MATERIA PRIMA","SERVICIOS","SUMINISTROS",BLANK()}})
    ),
    FILTER(
      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[producto])),
      NOT('Maestra de Facturacion (Total)'[producto] IN
        {{"PAVO C/M C/ASA EP CONG (8 KG)","ALIMENTACION COMERCIAL"}})
    ),
    FILTER(
      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[estado])),
      NOT('Maestra de Facturacion (Total)'[estado] IN {{"Cancelado"}})
    ),
    FILTER(
      KEEPFILTERS(VALUES('Calendario'[Date])),
      'Calendario'[Date] >= (DATE(2025, 7, 31) + TIME(0, 0, 1))
    )
  )
)"""
    rows = dax(token, ws_id, dataset_id, q, label)
    if rows:
        return rows[0].get("[v]") or rows[0].get("v")
    return None


def _dax_consumo_filtro_kardex(token, ws_id, dataset_id, value_expr, label, anio=None):
    """Aplica los filtros confirmados con Copiar consulta que comparten las
    tarjetas 'Costo Total' y 'Producción Neta (KG)' del reporte '14. Consumo
    Materiales indirectos de producción'.

    2026-09-03: primera captura, SIN año seleccionado en el reporte -> sin filtro
    de fecha, acumulado histórico total. 2026-09-04: se repitió la captura con
    el selector "2026" marcado en el reporte -> agrega TREATAS({{2026}},
    'Calendario'[Año]). Por eso `anio` es parámetro: pasar PREV_YEAR (o el año
    que corresponda) para obtener el dato del año, u omitir para el histórico total.
    `value_expr` es la expresión DAX del valor (columna con SUM o [Medida]).
    """
    # Se construye por concatenación simple (no f-string para todo el bloque) para
    # no arriesgar un error de escapado de llaves entre DAX y Python.
    filtro_anio = ("TREATAS({" + str(int(anio)) + "}, 'Calendario'[Año]),\n    ") if anio else ""
    q = (
        'EVALUATE\nROW(\n  "v",\n  CALCULATE(\n    ' + value_expr + ",\n    " +
        filtro_anio +
        "TREATAS({\"Costo\"}, 'PLANTA POR CECOS'[TIPO DE OPERACION]),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Maestra de Kardex (Total)'[categoria_hijo])),\n"
        "      NOT('Maestra de Kardex (Total)'[categoria_hijo] IN {\"ACUERDOS COMERCIALES\"})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Maestra de Kardex (Total)'[CUENTA ORIGEN])),\n"
        "      NOT('Maestra de Kardex (Total)'[CUENTA ORIGEN] IN {BLANK()})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Maestra de Kardex (Total)'[cuenta_analitica])),\n"
        "      NOT('Maestra de Kardex (Total)'[cuenta_analitica] IN\n"
        "        {BLANK(),\"[941002] CONTROL INTERNO\",\"ALMACEN ATE\",\"ALMACEN PACHACAMAC\"})\n"
        "    )\n"
        "  )\n"
        ")"
    )
    rows = dax(token, ws_id, dataset_id, q, label)
    if rows:
        return rows[0].get("[v]") or rows[0].get("v")
    return None






# KPIs que se leen ejecutando la consulta EXACTA de la tarjeta del reporte,
# identificada por su hash en capturas/catalogo.json.
#
# Reconstruir la consulta a mano fue un error caro: la de Consumo llevaba
# siete filtros que la tarjeta no tiene, y publicaba 16,750,700 kg donde el
# reporte marca 10,178,475. Al corregir la columna salió 1,308,396, tampoco
# el del reporte. Dos intentos, dos cifras equivocadas, porque el problema
# nunca fue la columna sino los filtros.
#
# Ejecutando el DAX capturado tal cual, la app no puede diferir del reporte:
# es literalmente la misma consulta que Power BI usa para pintar la tarjeta.
# Las funciones dax_consumo_* que reconstruían estas consultas a mano se
# eliminaron el 13/09/2026: llevaban filtros que la tarjeta no tiene y daban
# cifras distintas a las del reporte. Ejecutar el DAX capturado las hace
# innecesarias.
TARJETAS_KPI = {
    # Control Interno (reporte 8). Cuatro de sus tarjetas no llevan título en
    # el reporte y salen como "Tarjeta": se identificaron por la medida que
    # consultan, no por el nombre.
    ("control_ds", "% Cumplimiento"):      "2dd066c629b2",
    ("control_ds", "Satisfactorio"):       "52bb4f0a83cf",
    ("control_ds", "Con Observaciones"):   "75ee47245f6d",
    ("control_ds", "Crítico"):             "002616d79aa2",
    ("control_ds", "Planes de Acción Abiertos"):   "e4ba6f862d51",
    ("control_ds", "Planes de Acción Cerrados"):   "d102a3050d82",
    ("control_ds", "Planes de Acción Atrasados"):  "6a4c8e9382bb",
    ("control_ds", "Puntos Ejecutados"):   "74e9e532453c",

    ("consumo", "Costo Total"):            "7fc030955546",
    ("consumo", "Venta Neta (KG)"):        "7768ef2ed901",
    ("consumo", "Costo x TN Vendida"):     "d94f00974eb5",
    ("consumo", "Producción Neta (KG)"):   "cce2788bf3ec",
    ("consumo", "Costo x TN Producida"):   "7cbb71b8d5c7",
}


# KPIs cuyo valor se leyó de la tarjeta del reporte. La reconciliación de
# series no debe pisarlos: la tarjeta es lo que el usuario ve en Power BI.
LEIDOS_DE_TARJETA = []

# Claves de `scanned` que se llenaron con una consulta del reporte (capturada
# con Copiar consulta o ejecutando la tarjeta), no con el sondeo genérico.
#
# Sirve como evidencia POSITIVA para marcar_origen. Antes solo había evidencia
# negativa —"su valor no coincide con ninguno del sondeo"— y eso marcaba como
# dudosas las once cifras de Control Interno, que vienen de consultas exactas
# pero dan el mismo número que el sondeo porque ese reporte no tiene filtros
# de página que cambien el resultado.
CLAVES_DEL_REPORTE = set()


def guardar_del_reporte(scanned, ds, clave, valor):
    """Guarda un valor y deja constancia de que vino de una consulta real."""
    if valor is None:
        return False
    scanned.setdefault(ds, {})[clave] = valor
    CLAVES_DEL_REPORTE.add((ds, clave))
    return True


def valor_tarjeta_por_hash(token, ws, dataset_id, hash_visual, label):
    """Ejecuta la consulta capturada de una tarjeta y devuelve su número.

    El nombre lleva 'por_hash' porque ya existe valor_de_tarjeta(), que busca
    por título de visual entre varios datasets. Al llamarse igual, esta
    quedaba pisada por aquella y las cinco lecturas de Consumo morían con
    "takes 4 positional arguments but 5 were given" — silenciosas, porque cada
    una va en su propio try. Dos nombres iguales en un archivo de 4.400 líneas
    no se ven leyendo; se ven cuando algo falla.
    """
    entrada = next((q for q in _catalogo_por_hash().get(hash_visual, [])), None)
    if not entrada:
        DIAGNOSTICO.append({"consulta": f"tarjeta:{label}", "http": 0,
                            "error": f"hash {hash_visual} no está en el catálogo"})
        return None
    tablas = tablas_de_captura(
        lambda dax, lb: (_tablas_dax(token, ws, dataset_id, dax, lb) or [[]])[0],
        entrada["dax"], f"tarjeta:{label}")
    filas = tablas[0] if tablas else []
    if not filas:
        return None
    for v in filas[0].values():
        if isinstance(v, (int, float)):
            return float(v)
    return None


_CAT_HASH = None


def _catalogo_por_hash():
    global _CAT_HASH
    if _CAT_HASH is None:
        _CAT_HASH = {}
        try:
            for q in json.loads(CATALOGO_CAPTURAS.read_text(encoding="utf-8")):
                _CAT_HASH.setdefault(q.get("hash"), []).append(q)
        except Exception as e:
            print(f"    · catálogo ilegible: {e}")
    return _CAT_HASH










def _dax_margen_filtros(uen, anio):
    """Los 12 filtros de página confirmados con Copiar consulta (2026-09-04) sobre
    la tabla "Análisis de Ventas" del reporte de Margen: CATEGORIZACION='VENTA BRUTA',
    exclusión de "BONIFICACION SELL IN", Calendario[Año]=<anio>, y 8 filtros más de
    exclusión de categoria_producto/producto/estado/CUENTA ORIGEN sobre 4 tablas
    ('Exl A Maestra de Facturas de Venta', 'Maestra de Facturacion (Total)',
    'Maestra de Kardex (Total)', 'OrdenxFactura').

    Si `uen` no es None, agrega un 13er filtro TIPO DE NEGOCIO N1=<uen> — confirmado
    capturando la misma consulta con la tarjeta filtrada a "B&D". Sin `uen` es el
    consolidado total (confirmado con una segunda captura sin ese filtro).
    """
    filtro_uen = ""
    if uen:
        uen_esc = uen.replace('"', '\\"')
        filtro_uen = "    TREATAS({\"" + uen_esc + "\"}, 'Exl Tipo de Negocio'[TIPO DE NEGOCIO N1]),\n"
    return (
        "TREATAS({\"VENTA BRUTA\"}, 'Exl A Maestra de Facturas de Venta'[CATEGORIZACION]),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Exl A Maestra de Facturas de Venta'[DESCRIPCION])),\n"
        "      NOT('Exl A Maestra de Facturas de Venta'[DESCRIPCION] IN {\"BONIFICACION SELL IN\"})\n"
        "    ),\n"
        + filtro_uen +
        "    TREATAS({" + str(int(anio)) + "}, 'Calendario'[Año]),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Exl A Maestra de Facturas de Venta'[categoria_producto])),\n"
        "      NOT('Exl A Maestra de Facturas de Venta'[categoria_producto] IN {\"CHATARRA\",\"SERVICIOS\",BLANK()})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Exl A Maestra de Facturas de Venta'[producto])),\n"
        "      NOT('Exl A Maestra de Facturas de Venta'[producto] IN\n"
        "        {\"PAVO C/M C/ASA EP CONG (8 KG)\",\"ALIMENTACION COMERCIAL\"})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Exl A Maestra de Facturas de Venta'[estado])),\n"
        "      NOT('Exl A Maestra de Facturas de Venta'[estado] IN {\"Cancelado\"})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[categoria_producto])),\n"
        "      NOT('Maestra de Facturacion (Total)'[categoria_producto] IN\n"
        "        {BLANK(),\"BONIFICACION Y REBATES\",\"INTERESES\",\"MATERIA PRIMA\",\"SUMINISTROS\",\"SERVICIOS\"})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[estado])),\n"
        "      NOT('Maestra de Facturacion (Total)'[estado] IN {\"Cancelado\"})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Maestra de Kardex (Total)'[CUENTA ORIGEN])),\n"
        "      NOT('Maestra de Kardex (Total)'[CUENTA ORIGEN] IN {BLANK()})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('OrdenxFactura'[categoria_producto])),\n"
        "      NOT('OrdenxFactura'[categoria_producto] IN\n"
        "        {BLANK(),\"BONIFICACION Y REBATES\",\"CHATARRA\",\"SERVICIOS\"})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('OrdenxFactura'[estado])),\n"
        "      NOT('OrdenxFactura'[estado] IN {\"Cancelado\"})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('OrdenxFactura'[producto])),\n"
        "      NOT('OrdenxFactura'[producto] IN\n"
        "        {\"PAVO C/M C/ASA EP CONG (8 KG)\",\"ALIMENTACION COMERCIAL\"})\n"
        "    )"
    )


def _dax_margen_precio_costo(token, ws_id, dataset_id, uen, anio, label):
    """Ejecuta Precio x Kilo / Costo x Kilo con los filtros de _dax_margen_filtros.
    Las medidas son 'Medidas Julito'[Precio x Kilo] y [Costo x Kilo] — mismo nombre
    que las medidas ya validadas para el margen consolidado, aquí referenciadas sin
    prefijo de tabla porque el nombre de medida es único en el modelo.
    """
    filtros = _dax_margen_filtros(uen, anio)
    q = (
        'EVALUATE\nROW(\n  "precio", CALCULATE([Precio x Kilo],\n    ' + filtros + '\n  ),\n'
        '  "costo", CALCULATE([Costo x Kilo],\n    ' + filtros + "\n  )\n)"
    )
    rows = dax(token, ws_id, dataset_id, q, label)
    if rows:
        precio = to_float(rows[0].get("[precio]") or rows[0].get("precio"))
        costo = to_float(rows[0].get("[costo]") or rows[0].get("costo"))
        return precio, costo
    return None, None


def dax_margen_uen(token, ws_id, dataset_id, uen, anio, label="margen_uen"):
    """Precio/kg y Costo/kg filtrados por UEN (TIPO DE NEGOCIO N1 = "B&D" / "MAQUILA" / "TIGO").
    La consulta original de Power BI agrupa por mes (visual de tabla); aquí se
    colapsa a un solo agregado del año completo.
    """
    return _dax_margen_precio_costo(token, ws_id, dataset_id, uen, anio, label)


def dax_margen_total(token, ws_id, dataset_id, anio, label="margen_total"):
    """Precio/kg y Costo/kg consolidados (todas las UEN), con el mismo filtro de
    página de 12 condiciones que dax_margen_uen pero sin el filtro de UEN —
    confirmado con una segunda Copiar consulta capturada sin ese filtro activo
    (2026-09-04). Reemplaza al valor que traía el escaneo genérico con solo
    filtro de fecha, que no aplicaba las 11 exclusiones adicionales del reporte.
    """
    return _dax_margen_precio_costo(token, ws_id, dataset_id, None, anio, label)


def dax_mermas_uen(token, ws_id, dataset_id, medida, anio, extra_filtro=None, label="mermas_uen"):
    """Ejecuta una medida de 'Tabla Mermas' (ej. '% Merma total B&D') del reporte
    '4. Reporte de mermas' con el filtro de página común, confirmado con Copiar
    consulta el 2026-09-04 sobre las tarjetas 'Merma B&D' y 'Merma TIGO':
      Calendario[Date] >= 2025-07-31 (mismo corte que otros reportes)
      Calendario[Año] = <anio>
      Maestra de Facturacion (Total)[categoria_producto] excluye 7 categorías
      Maestra de Facturacion (Total)[producto] excluye 2 productos
      Maestra de Facturacion (Total)[estado] excluye "Cancelado"

    `medida` es el nombre exacto de la medida en 'Tabla Mermas', ej.
    "% Merma total B&D" / "% Merma total TIGO".

    `extra_filtro` es un filtro DAX adicional propio de la tarjeta (string, sin
    coma final) — TIGO trae uno que B&D no tiene:
    'Tabla Mermas'[TIPO DE BASE] no vacío. No se asume que todas las UEN
    comparten exactamente los mismos filtros — cada una se confirma por separado.
    """
    medida_esc = medida.replace('"', '\\"')
    extra = f"    {extra_filtro},\n" if extra_filtro else ""
    q = (
        'EVALUATE\nROW(\n  "v", CALCULATE(\n    ' + f"'Tabla Mermas'[{medida_esc}]" + ',\n'
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Calendario'[Date])),\n"
        "      'Calendario'[Date] >= (DATE(2025, 7, 31) + TIME(0, 0, 1))\n"
        "    ),\n"
        + extra +
        "    TREATAS({" + str(int(anio)) + "}, 'Calendario'[Año]),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[categoria_producto])),\n"
        "      NOT('Maestra de Facturacion (Total)'[categoria_producto] IN\n"
        "        {\"BONIFICACION Y REBATES\",\"CHATARRA\",\"INTERESES\",\"MATERIA PRIMA\",\"SERVICIOS\",\"SUMINISTROS\",BLANK()})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[producto])),\n"
        "      NOT('Maestra de Facturacion (Total)'[producto] IN\n"
        "        {\"PAVO C/M C/ASA EP CONG (8 KG)\",\"ALIMENTACION COMERCIAL\"})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[estado])),\n"
        "      NOT('Maestra de Facturacion (Total)'[estado] IN {\"Cancelado\"})\n"
        "    )\n"
        "  )\n"
        ")"
    )
    rows = dax(token, ws_id, dataset_id, q, label)
    if rows:
        return to_float(rows[0].get("[v]") or rows[0].get("v"))
    return None


def dax_mermas_planta(token, ws_id, dataset_id, medida, almacenes, anio, con_tipo_base=True, label="mermas_planta"):
    """Medida de 'Tabla Mermas' del reporte de mermas (pestaña RESUMEN, gráfico
    'Merma Mensual por Planta') filtrada por 'Tabla Mermas'[almacen].

    Confirmado con Copiar consulta el 2026-09-04, capturado directamente del
    panel del Analizador de rendimiento (fila "PLANTA ATE"/"PLANTA PACHACAMAC"/
    "PLANTA TERCEROS" bajo el grupo "MERMA MENSUAL POR PLANTA", no las filas
    duplicadas de "MERMA DIARIA POR PLANTA" más abajo en la misma lista):
      Planta Ate         -> almacen "Lácteos Producción"                 · medida
                             genérica 'Tabla Mermas'[% Merma total] (sin sufijo)
                             · SÍ trae filtro TIPO DE BASE
      Planta Pachacámac  -> almacen "Salsas Producción"                   · medida
                             'Tabla Mermas'[% Merma total SALSAS]
                             · SÍ trae filtro TIPO DE BASE
      Planta Terceros    -> almacenes de maquiladores externos (Abuela
                             Maquila, Piamonte Maquila, Lácteos Dosimetría)
                             · medida 'Tabla Mermas'[% Merma total MAQUILA]
                             (una M — distinta de la medida UEN "MMAQUILA")
                             · NO trae filtro TIPO DE BASE (confirmado ausente
                             en su Copiar consulta, a diferencia de las otras 2)
    Cada planta puede tener su propia medida y sus propios filtros — no se
    asume el mismo patrón para todas, igual que ya pasó con Mermas por UEN.

    `almacenes` es una lista de valores de [almacen] — normalmente uno solo,
    pero TREATAS admite varios si una planta agrupa más de un almacén.
    `con_tipo_base` controla si se incluye el filtro TIPO DE BASE — pasar
    False para Terceros.
    """
    medida_esc = medida.replace('"', '\\"')
    vals = ",".join(f'"{a}"' for a in almacenes)
    tipo_base = (
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Tabla Mermas'[TIPO DE BASE])),\n"
        "      NOT('Tabla Mermas'[TIPO DE BASE] IN {BLANK()})\n"
        "    ),\n"
    ) if con_tipo_base else ""
    q = (
        'EVALUATE\nROW(\n  "v", CALCULATE(\n    ' + f"'Tabla Mermas'[{medida_esc}]" + ',\n'
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Calendario'[Date])),\n"
        "      'Calendario'[Date] >= (DATE(2025, 7, 31) + TIME(0, 0, 1))\n"
        "    ),\n"
        + tipo_base +
        "    TREATAS({" + vals + "}, 'Tabla Mermas'[almacen]),\n"
        "    TREATAS({" + str(int(anio)) + "}, 'Calendario'[Año]),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[categoria_producto])),\n"
        "      NOT('Maestra de Facturacion (Total)'[categoria_producto] IN\n"
        "        {\"BONIFICACION Y REBATES\",\"CHATARRA\",\"INTERESES\",\"MATERIA PRIMA\",\"SERVICIOS\",\"SUMINISTROS\",BLANK()})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[producto])),\n"
        "      NOT('Maestra de Facturacion (Total)'[producto] IN\n"
        "        {\"PAVO C/M C/ASA EP CONG (8 KG)\",\"ALIMENTACION COMERCIAL\"})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Maestra de Facturacion (Total)'[estado])),\n"
        "      NOT('Maestra de Facturacion (Total)'[estado] IN {\"Cancelado\"})\n"
        "    )\n"
        "  )\n"
        ")"
    )
    rows = dax(token, ws_id, dataset_id, q, label)
    if rows:
        return to_float(rows[0].get("[v]") or rows[0].get("v"))
    return None


def dax_inventario_kardex(token, ws_id, dataset_id, medida, anio, mes, label="inv_kardex"):
    """Medida de 'KARDEX TOTAL' del reporte '6. Rotación de inventario'
    (página "RI (Clasificación)", gráfico "Evolución de Inventario").

    Confirmado con Copiar consulta el 2026-09-04. Filtros propios de este
    dataset — distintos a los de Margen/Mermas, sin las exclusiones de
    categoria_producto/producto/estado que usan esos otros reportes:
      Calendario[Mes Año] <> "Jun 2025" (el reporte excluye ese mes siempre,
        probablemente por dato parcial al inicio del histórico)
      KARDEX TOTAL[Fecha] dentro del mes/año pedido

    `medida` ej. "Días Rotación" (≈ Cobertura Total en el visual) o
    "Días Rotación MP" (≈ Cobertura MP). También sirve para
    "Consumo Acumulado Total S/ saldo".
    """
    medida_esc = medida.replace('"', '\\"')
    anio, mes = int(anio), int(mes)
    anio_sig, mes_sig = (anio + 1, 1) if mes == 12 else (anio, mes + 1)
    q = (
        'EVALUATE\nROW(\n  "v", CALCULATE(\n    ' + f"'KARDEX TOTAL'[{medida_esc}]" + ',\n'
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('Calendario'[Mes Año])),\n"
        "      NOT('Calendario'[Mes Año] IN {\"Jun 2025\"})\n"
        "    ),\n"
        "    FILTER(\n"
        "      KEEPFILTERS(VALUES('KARDEX TOTAL'[Fecha])),\n"
        "      AND(\n"
        f"        'KARDEX TOTAL'[Fecha] >= DATE({anio}, {mes}, 1),\n"
        f"        'KARDEX TOTAL'[Fecha] < DATE({anio_sig}, {mes_sig}, 1)\n"
        "      )\n"
        "    )\n"
        "  )\n"
        ")"
    )
    rows = dax(token, ws_id, dataset_id, q, label)
    if rows:
        return to_float(rows[0].get("[v]") or rows[0].get("v"))
    return None


def dax_inventario_composicion(token, ws_id, dataset_id, label="inv_composicion"):
    """Desglose de Saldo Soles por 'SALDO ACTUAL'[Clasificación Segun Consumo]
    (WORKING / EXCESO 1 / EXCESO 2 / DEAD) del gráfico "Composición de
    Inventario" en '6. Rotación de inventario'.

    Confirmado con Copiar consulta el 2026-09-04. A diferencia de
    dax_inventario_kardex, esta consulta devuelve VARIAS filas (una por
    categoría) — no un solo valor. El filtro de fecha es amplio
    (2025-06-30 a 2027-01-01): la clasificación es una foto actual, no un
    corte mensual.

    Devuelve una lista de (clasificacion, saldo_soles) ordenada de mayor a
    menor, igual que el gráfico de cascada real.
    """
    q = (
        "EVALUATE\n"
        "SUMMARIZECOLUMNS(\n"
        "  'SALDO ACTUAL'[Clasificación Segun Consumo],\n"
        "  FILTER(\n"
        "    KEEPFILTERS(VALUES('KARDEX TOTAL'[Fecha])),\n"
        "    AND(\n"
        "      'KARDEX TOTAL'[Fecha] >= DATE(2025, 6, 30),\n"
        "      'KARDEX TOTAL'[Fecha] < DATE(2027, 1, 1)\n"
        "    )\n"
        "  ),\n"
        '  "SumSaldo_Soles", CALCULATE(SUM(\'SALDO ACTUAL\'[Saldo Soles]))\n'
        ")\n"
        "ORDER BY [SumSaldo_Soles] DESC"
    )
    rows = dax(token, ws_id, dataset_id, q, label)
    out = []
    for r in rows or []:
        clasif = None
        saldo = None
        for k, v in r.items():
            if "Clasificaci" in k:
                clasif = v
            elif "SumSaldo" in k:
                saldo = to_float(v)
        if clasif is not None:
            out.append((clasif, saldo))
    return out


def dax_control_interno(token, ws_id, dataset_id, medida, empresa="Pauno", label="control_interno"):
    """Medida de 'Pauno Registro de Ejecuciones' del reporte '8. Reporte de
    auditoría' (Dashboard de Control Interno).

    Confirmado con Copiar consulta el 2026-09-04 sobre la tarjeta
    "Satisfactorio": el único filtro es Empresa="Pauno" — SIN filtro de
    fecha ni mes, a pesar de que el dashboard tiene un selector de mes.
    Los KPIs de resumen (Puntos de Control, Satisfactorio, Con Observaciones,
    Crítico, % Cumplimiento) son un acumulado histórico total, no del mes.

    `medida` ej. "Satisfactorio", "Con Observaciones", "Critico",
    "Puntos de Control", "% Cumplimiento".
    """
    medida_esc = medida.replace('"', '\\"')
    empresa_esc = empresa.replace('"', '\\"')
    q = (
        'EVALUATE\nROW(\n  "v", CALCULATE(\n    ' + f"'Pauno Registro de Ejecuciones'[{medida_esc}]" + ',\n'
        f'    TREATAS({{"{empresa_esc}"}}, \'Pauno Registro de Ejecuciones\'[Empresa])\n'
        "  )\n"
        ")"
    )
    rows = dax(token, ws_id, dataset_id, q, label)
    if rows:
        return to_float(rows[0].get("[v]") or rows[0].get("v"))
    return None


def dax_planes_accion(token, ws_id, dataset_id, medida, empresa="Pauno", label="planes_accion"):
    """Medida de 'Planes de Accion' (mismo dataset que Control Interno,
    tabla distinta a 'Pauno Registro de Ejecuciones').

    Confirmado con Copiar consulta el 2026-09-04 sobre la tarjeta
    "Planes Abiertos": filtro único Empresa="Pauno" (mismo patrón que
    Control Interno, sin fecha), pero usando SUMMARIZECOLUMNS + IGNORE
    en vez de CALCULATE (así lo genera Power BI para esta tarjeta).

    `medida` ej. "Planes Abiertos", "Planes Atrasados", "Total Planes de Acción".
    """
    medida_esc = medida.replace('"', '\\"')
    empresa_esc = empresa.replace('"', '\\"')
    alias = "".join(c if c.isalnum() else "_" for c in medida)
    q = (
        "DEFINE VAR __DS0FilterTable = \n"
        f'\tTREATAS({{"{empresa_esc}"}}, \'Pauno Registro de Ejecuciones\'[Empresa])\n\n'
        "EVALUATE\n"
        f'\tSUMMARIZECOLUMNS(__DS0FilterTable, "{alias}", IGNORE(\'Planes de Accion\'[{medida_esc}]))'
    )
    rows = dax(token, ws_id, dataset_id, q, label)
    if rows:
        return to_float(rows[0].get(f"[{alias}]") or rows[0].get(alias))
    return None


def dax_cxc_aging(token, ws_id, dataset_id, label="cxc_aging"):
    """Tramos de antigüedad (aging) de Cuentas por Cobrar.

    Confirmado con Copiar consulta el 2026-09-06 sobre la tarjeta
    "MAS DE 30 DIAS" del reporte '1. Cuentas por cobrar':

        DEFINE
          VAR __DS0FilterTable  = TREATAS({"4. Más de 30 días"}, 'DATA_FACTURACION'[O_SEGMENTO])
          VAR __DS0FilterTable2 = TREATAS({"Pendiente"},         'DATA_FACTURACION'[CXC])
        EVALUATE SUMMARIZECOLUMNS(__DS0FilterTable, __DS0FilterTable2,
          "SumTotal_fact", IGNORE(CALCULATE(SUM('DATA_FACTURACION'[Total_fact]))))

    En vez de repetir esa consulta 4 veces adivinando los nombres de los
    otros 3 tramos, agrupamos POR [O_SEGMENTO] manteniendo el mismo filtro
    CXC="Pendiente" y la misma suma. Así Power BI devuelve las etiquetas
    exactas — no las inventamos — y de paso el total de la cartera pendiente.

    Devuelve [(segmento, monto), ...] en el orden que da el modelo
    (los segmentos vienen numerados "1. ", "2. "... así que ordenan solos).
    """
    q = (
        "DEFINE\n"
        '\tVAR __DS0FilterTable = \n'
        '\t\tTREATAS({"Pendiente"}, \'DATA_FACTURACION\'[CXC])\n\n'
        "EVALUATE\n"
        "\tSUMMARIZECOLUMNS(\n"
        "\t\t'DATA_FACTURACION'[O_SEGMENTO],\n"
        "\t\t__DS0FilterTable,\n"
        '\t\t"SumTotal_fact", IGNORE(CALCULATE(SUM(\'DATA_FACTURACION\'[Total_fact])))\n'
        "\t)"
    )
    out = []
    try:
        rows = dax(token, ws_id, dataset_id, q, label)
        for r in rows or []:
            seg = r.get("DATA_FACTURACION[O_SEGMENTO]") or r.get("[O_SEGMENTO]") or r.get("O_SEGMENTO")
            val = to_float(r.get("[SumTotal_fact]") or r.get("SumTotal_fact"))
            if seg is not None and val is not None:
                out.append((str(seg), val))
    except Exception as e:
        print(f"    · aging agrupado falló ({e}) — uso las 4 consultas exactas")

    if not out:
        # Respaldo: las 4 tarjetas tal cual se capturaron con Copiar consulta
        # el 2026-09-06. Estos nombres NO son inventados — cada uno se verificó
        # individualmente contra su tarjeta en el reporte '1. Cuentas por cobrar'.
        for seg in ("1. Por Vencer", "2. 0 a 15 días", "3. 16 a 30 días", "4. Más de 30 días"):
            qs = (
                "DEFINE\n"
                f'\tVAR __DS0FilterTable = \n\t\tTREATAS({{"{seg}"}}, \'DATA_FACTURACION\'[O_SEGMENTO])\n\n'
                '\tVAR __DS0FilterTable2 = \n\t\tTREATAS({"Pendiente"}, \'DATA_FACTURACION\'[CXC])\n\n'
                "EVALUATE\n"
                "\tSUMMARIZECOLUMNS(\n"
                "\t\t__DS0FilterTable,\n"
                "\t\t__DS0FilterTable2,\n"
                '\t\t"SumTotal_fact", IGNORE(CALCULATE(SUM(\'DATA_FACTURACION\'[Total_fact])))\n'
                "\t)"
            )
            try:
                rows = dax(token, ws_id, dataset_id, qs, f"{label}:{seg}")
                if rows:
                    v = to_float(rows[0].get("[SumTotal_fact]") or rows[0].get("SumTotal_fact"))
                    if v is not None:
                        out.append((seg, v))
            except Exception as e:
                print(f"    ✗ aging [{seg}]: {e}")

    out.sort(key=lambda t: t[0])
    return out


def dax_control_interno_sum(token, ws_id, dataset_id, columna, empresa="Pauno", label="control_interno_sum"):
    """Suma de una columna (no medida) de 'Pauno Registro de Ejecuciones'.

    Confirmado con Copiar consulta el 2026-09-04 sobre la columna
    '¿Ejecutado?': mismo filtro Empresa="Pauno" que el resto de Control
    Interno, pero usando SUMMARIZECOLUMNS + IGNORE(CALCULATE(SUM(...)))
    en vez de referenciar una medida ya definida en el modelo.

    `columna` ej. "¿Ejecutado?".
    """
    columna_esc = columna.replace('"', '\\"')
    empresa_esc = empresa.replace('"', '\\"')
    alias = "Sumv_" + "".join(c for c in columna if c.isalnum()) + "_"
    q = (
        "DEFINE VAR __DS0FilterTable = \n"
        f'\tTREATAS({{"{empresa_esc}"}}, \'Pauno Registro de Ejecuciones\'[Empresa])\n\n'
        "EVALUATE\n"
        "\tSUMMARIZECOLUMNS(\n"
        "\t\t__DS0FilterTable,\n"
        f'\t\t"{alias}", IGNORE(CALCULATE(SUM(\'Pauno Registro de Ejecuciones\'[{columna_esc}])))\n'
        "\t)"
    )
    rows = dax(token, ws_id, dataset_id, q, label)
    if rows:
        return to_float(rows[0].get(f"[{alias}]") or rows[0].get(alias))
    return None


def dax_control_resultado_acumulado(token, ws_id, dataset_id, empresa="Pauno", label="control_resultado"):
    """Tarjeta combinada 'Resultado Acumulado' del Dashboard de Control Interno.

    Confirmado con Copiar consulta el 2026-09-04: mismo filtro Empresa="Pauno"
    que el resto de Control Interno, pero es UNA sola visual que trae 4 valores
    juntos — incluye "% Cumplimento" como medida real del modelo (no hace
    falta calcularlo con Puntos Teoricos/Calificacion, ya viene resuelto).

    Devuelve dict: {color_kpi, cumplimiento_pct, calificacion, puntos_totales}.
    """
    empresa_esc = empresa.replace('"', '\\"')
    q = (
        "DEFINE VAR __DS0FilterTable = \n"
        f'\tTREATAS({{"{empresa_esc}"}}, \'Pauno Registro de Ejecuciones\'[Empresa])\n\n'
        "EVALUATE\n"
        "\tSUMMARIZECOLUMNS(\n"
        "\t\t__DS0FilterTable,\n"
        '\t\t"Color_KPI", IGNORE(\'Pauno Registro de Ejecuciones\'[Color KPI]),\n'
        '\t\t"v__Cumplimento", IGNORE(\'Pauno Registro de Ejecuciones\'[% Cumplimento]),\n'
        '\t\t"SumCalificación_0_3_", IGNORE(\n'
        "\t\t\tCALCULATE(SUM('Pauno Registro de Ejecuciones'[Calificación(0-3)]))\n"
        "\t\t),\n"
        '\t\t"SumPuntos_Totales", IGNORE(CALCULATE(SUM(\'Pauno Registro de Ejecuciones\'[Puntos Totales])))\n'
        "\t)"
    )
    rows = dax(token, ws_id, dataset_id, q, label)
    if not rows:
        return None
    r = rows[0]
    return {
        "color_kpi": r.get("[Color_KPI]") or r.get("Color_KPI"),
        "cumplimiento_pct": to_float(r.get("[v__Cumplimento]") or r.get("v__Cumplimento")),
        "calificacion": to_float(r.get("[SumCalificación_0_3_]") or r.get("SumCalificación_0_3_")),
        "puntos_totales": to_float(r.get("[SumPuntos_Totales]") or r.get("SumPuntos_Totales")),
    }


def dax_prev_month(token, ws_id, dataset_id, measure_name, date_tbl, date_col, label="dated", registrar=True):
    """Ejecuta medida filtrada al mes anterior completo."""
    q = f"""EVALUATE
ROW("v",
  CALCULATE(
    [{measure_name}],
    FILTER(ALL('{date_tbl}'), YEAR('{date_tbl}'[{date_col}]) = {PREV_YEAR} && MONTH('{date_tbl}'[{date_col}]) = {PREV_MONTH})
  )
)"""
    rows = dax(token, ws_id, dataset_id, q, label, registrar=registrar)
    if rows:
        return rows[0].get("[v]") or rows[0].get("v")
    return None

WORKSPACES = {
    "PAUNO": "461932ad-b5ec-4fd6-aa97-f1fc7bdc5169",
}

# Identificadores de REPORTE, los que se ven en la barra de direcciones de
# Power BI (.../reports/<id>/...). A diferencia del dataset, no cambian cuando
# alguien vuelve a publicar el informe — y cuando cambia el dataset, el
# extractor se queda apuntando a algo que ya no existe: entre el 9 y el 13 de
# setiembre las consultas de mermas devolvieron PowerBIEntityNotFound mientras
# los otros diez reportes respondian normal. Con esto el dataset se resuelve
# solo en cada corrida.
REPORT_IDS = {
    "PAUNO": {
        "mermas": "3fa688de-f4ad-4e58-878b-d2d816275ff3",  # 4. Reporte de mermas
    }
}


def dataset_de_reporte(token, ws, report_id, label):
    """Dataset que alimenta a un reporte. None si no se puede averiguar."""
    try:
        r = requests.get(f"{PBI_BASE}/groups/{ws}/reports/{report_id}",
                         headers={"Authorization": f"Bearer {token}"}, timeout=30)
        if r.status_code != 200:
            DIAGNOSTICO.append({"consulta": f"reporte:{label}", "http": r.status_code,
                                "error": r.text[:300]})
            return None
        return r.json().get("datasetId")
    except Exception as e:
        DIAGNOSTICO.append({"consulta": f"reporte:{label}", "http": 0,
                            "error": repr(e)[:200]})
        return None


# Los identificadores ya corregidos, para el codigo que corre fuera del bucle
# por empresa y no tiene `ids` a mano. Sin esto, esas partes seguirian usando
# el dataset viejo y el fallo volveria por una puerta lateral.
DATASETS_RESUELTOS = {}


def datasets_de(empresa):
    """Identificadores vigentes: los resueltos si los hay, si no los fijos."""
    return DATASETS_RESUELTOS.get(empresa) or DATASET_IDS.get(empresa, {})


def resolver_datasets(token, ws, empresa, ids):
    """Corrige los identificadores de dataset que quedaron obsoletos.

    Se deja constancia del cambio: si un reporte se republica cada semana,
    conviene verlo en el diagnóstico y no descubrirlo cuando el dato falla.
    """
    for clave, report_id in (REPORT_IDS.get(empresa) or {}).items():
        real = dataset_de_reporte(token, ws, report_id, clave)
        if not real:
            continue
        if ids.get(clave) != real:
            print(f"    · {clave}: el dataset cambió — {ids.get(clave)} → {real}")
            DIAGNOSTICO.append({
                "consulta": f"dataset_actualizado:{clave}", "http": 200,
                "error": (f"el reporte {report_id} ahora usa el dataset {real}; "
                          f"el configurado era {ids.get(clave)}")})
            ids[clave] = real
        else:
            print(f"    · {clave}: dataset confirmado {real}")
    DATASETS_RESUELTOS[empresa] = ids
    return ids


# IDs confirmados por discover_all_datasets() en ejecución anterior
DATASET_IDS = {
    "PAUNO": {
        "cxc":            "2eec70cd-0820-408f-b938-a2cd547b0c18",  # 1. Cuentas por cobrar
        "cxp":            "45a8ab8d-e162-4398-a260-3a9a5f90829f",  # 2. Cuentas por pagar
        "margen":         "38076daa-d2cd-4a93-858a-82c0a4cf8cb6",  # 3. Reporte de Margen
        "mermas":         "fdd58a2d-654d-41e0-b587-3d43c337ed47",  # 4. Reporte de mermas
        "compras":        "06408938-8202-424e-80c0-b42c178dabde",  # 5. Reporte de compras
        "inventario":     "0e27d784-41a4-48f0-9208-60210119f0a7",  # 6. Rotacion de inventario
        "control_ds":     "0aca7bdd-6b72-49c2-be41-17ae0f6b5848",  # 8. Reporte de auditoria
        "productividad_ds": "a042ba6a-c82c-4fc1-bf95-b9e84bd15fc6", # 13. Productividad
        "planificacion":  "30074d92-7ec1-4762-82f2-1cb29c15dcfe",  # 11. Planificaciones
        "consumo":        "c972c8cb-e5fc-4b60-8f5e-265a78e1e796",  # 14. Consumo Materiales
        "fill_rate":      "7f4ebe22-5e90-4e35-973b-4af3c58497e5",  # 12. Calculo de Provisiones
    }
}

# Candidatos de medidas por tipo — orden de más probable a menos
SCAN_CANDIDATES = {
    "cxc": [
        # La oficial es '% Morosidad' (21.08%), confirmada por el usuario el
        # 2026-09-06. El dataset tiene además una medida 'Morosidad' que
        # devuelve 13.78% — mide otra cosa y NO debe usarse ni como respaldo:
        # si se colara, el semáforo pasaría de rojo a amarillo y cambiaría la
        # lectura del reporte. Por eso está excluida a propósito.
        "% Morosidad", "% Mora", "Tasa Morosidad",
        "Por Vencer", "Saldo Por Vencer", "CxC Por Vencer", "No Vencido",
        "CxC Total", "Total CxC", "Saldo CxC", "Cartera Total",
        "Vencido", "Saldo Vencido", "Total Vencido",
        "0-15 días", "16-30 días", "+30 días", "0-15d", "16-30d", "+30d",
        "Rotacion CxC", "Días CxC", "Días de Cobro",
    ],
    "cxp": [
        # ── CONFIRMADOS por prueba directa ──
        "Cuentas x Pagar", "CUENTAS X PAGAR", "Refinanciamiento", "Proveedores",
        # ── Tramos vencido (deducidos de la captura del reporte) ──
        "Vigente CxP", "CxP Vigente", "No Vencido CxP",
        "Vencido < 15", "Vencido Menor 15d", "Vencido 0-15d",
        "Vencido 16-30d", "Vencido 30d", "Vencido 1-30",
        "Vencido 31-90", "Vencido 31-90d", "Vencido 90d",
        "Vencido >90", "Vencido >90d", "Vencido Mayor 90",
        "Saldo Vigente", "Saldo Vencido",
        # ── Días de pago (eje del gráfico) ──
        "Dias CxP", "DPP", "Dias Pago", "Rotacion CxP",
        "Dias Promedio Pago", "Dias de Pago", "Días de Pago",
        "PPP", "Plazo Pago", "Plazo Promedio Pago",
        # ── Otros candidatos ──
        "Total CxP", "Saldo CxP", "CxP Total",
        "Deuda Total", "Total Deuda", "Monto CxP",
    ],
    "margen": [
        # Columnas exactas visibles en tabla "DETALLE ANALISIS DE COSTOS" (Power BI PAUNO)
        "% Margen", "Costo Total", "P.Unit", "P.Unit NC", "C.Unit",
        # Nombres de tabs visibles en Power BI PAUNO
        "R. MARGEN", "R. Margen", "R MARGEN",
        "R. COSTO UNIT", "R. Costo Unit", "Costo Unitario", "Costo Unit",
        # Margen variable
        "Margen Variable", "% Margen Variable", "MV", "% MV", "MV%", "Margen Bruto",
        "Margen", "Costo Variable", "% Costo Variable",
        "% MV PAUNO", "MV PAUNO", "Margen Variable %", "Pct MV",
        # Ventas
        "Ventas Mes Actual", "Ventas Actuales", "Ventas", "Venta Total", "Venta Neta",
        "Ventas Mes Anterior", "Ventas Anterior",
        # Precio / Costo por kg
        "Precio/kg", "Precio kg", "Precio Promedio kg", "Precio x Kilo", "Precio Kilo",
        "Costo/kg", "Costo kg", "Costo x Kilo", "Costo Kilo",
        # Nombres con mayúsculas/abrev peruanas
        "MARGEN VARIABLE", "% MARGEN", "MARGEN", "MV TOTAL",
        "VENTA NETA", "VENTAS NETAS", "VENTA MES",
        "Ing. Ventas", "Ingresos", "Ingresos Ventas",
        "Resultado Bruto", "Utilidad Bruta", "% Utilidad Bruta",
    ],
    "mermas": [
        # Columnas exactas del "Detalle de Mermas por Planta" (Control de Producción)
        "REAL", "STD", "DESVIACION", "ITEM DESVIADO", "% MERMAS",
        "Stock Prod", "Stock Obs", "Merma",
        # Medidas estándar
        "% Merma", "Merma %", "Tasa Merma", "Merma Total", "% Merma Total",
        "% Merma Ate", "Merma Ate", "Merma Planta Ate",
        "% Merma Pachacamac", "Merma Pachacamac", "Merma Lurin",
        "% Merma Terceros", "Merma Terceros",
        # Por UEN
        "% Merma B&D", "Merma B&D", "% Merma BD",
        "% Merma Tigo", "Merma Tigo", "% Merma TIGO",
        "% Merma Maquila", "Merma Maquila", "% Merma MAQUILA",
        "Merma Valorizada", "Kg Merma",
    ],
    "compras": [
        # Columnas exactas "ANALISIS DE MATERIALES (PLANIFICACION Y COMPRA)" / "Reporte Compras"
        "Lead Time", "Stock PP", "Stock Valorizado",
        "Costo Unit", "Costo Unitario",
        "Cons Mes Ant", "Consumo Mes Anterior",
        "Consumo Prom 3M", "Consumo Promedio 3M",
        "Consumo Prom 6M", "Consumo Promedio 6M",
        "Faltantes", "Requerimiento",
        # Fila total: Consumo | Compra | Ratio por período
        "Ratio", "Cant Compra", "Cant Consumo",
        "% Ratio", "Ratio C/V", "Ratio Compras", "Ratio Consumo/Compra",
        "Eficiencia", "Eficiencia Costo", "% Eficiencia",
        "Ratio C/C", "Consumo/Compra", "C/C",
        "Consumo Total", "Consumo Periodo", "Importe Consumo", "Monto Consumo",
        "Compra Total", "Monto Compras", "Importe Compras", "Valor Compras",
        "MP Consumo", "MP Compras",
        "Total Compras Periodo", "Ordenes de Compra",
        "Compra Mes", "Consumo Mes", "Compras Mes",
        "Valor Compra", "Valor Consumo",
    ],
    "inventario": [
        "Dead Stock", "Stock Muerto", "Inmovilizado", "Stock Inmovilizado",
        "% Dead Stock", "% Inmovilizado", "% Dead", "Pct Dead",
        "Inventario Total", "Total Inventario", "Saldo Inventario", "Valor Inventario",
        "Working Stock", "Stock Activo", "Stock Working", "Stock Normal",
        "Días de Inventario", "Dias Inventario", "Cobertura", "Cobertura Días",
        "Exceso 1", "Exceso 2", "Sobre Stock",
        "Rotacion", "Rotación", "Veces Rotacion",
        "Stock Total", "Total Stock", "Costo Inventario",
    ],
    "control_ds": [
        "% Cumplimiento", "Cumplimiento", "Conformidad", "% Conformidad",
        "Puntos Críticos", "Puntos Criticos", "Hallazgos Críticos",
        "Satisfactorio", "Observaciones",
        "No Conformidades", "Incumplimientos",
    ],
    "productividad_ds": [
        # Nombres REALES confirmados leyendo el panel Datos > Measures en Power BI Service
        # (Editar informe) y validados con EVALUATE ROW(...) en DAX Query View — 2026-09-03.
        "SUM PROD",              # -> Producción Total (KG)
        "GASTO TOTAL",           # -> Planilla Total (S/.)
        "PROD OPERATIVA",        # -> Planilla (S/.) entre KG Producido
        "VENTA KG X SOL",        # -> Planilla (S/.) entre KG Vendido
        # "Venta Neta (KG)" en el panel aparece como "Suma de Peso total KG" — eso es una
        # COLUMNA con agregación implícita en la UI (Power BI antepone "Suma de" a columnas,
        # no a medidas DAX reales), NO un nombre de medida válido para EVALUATE ROW(). Sin
        # confirmar la tabla/columna real, no se adivina — queda sin este KPI hasta que se
        # confirme con el Analizador de rendimiento (ver nota en build_productividad).
        # Nombres viejos (adivinados, se mantienen como fallback por si alguno coincide):
        "Planilla (S/.) entre KG Vendido", "Planilla Entre KG Vendido", "Planilla/KG Vendido",
        "Planilla (S/.) entre KG Producido", "Planilla Entre KG Producido", "Planilla/KG Producido",
        "Produccion Total (KG)", "Produccion Total KG", "KG Producido", "Produccion Total",
        "Venta Neta (KG)", "Venta Neta KG", "KG Vendido", "Venta Neta",
        "Planilla Total (S/.)", "Planilla Total", "Costo Planilla",
        # Por planta (ATE / PACHACAMAC)
        "Produccion ATE (KG)", "Produccion ATE", "KG Producido ATE",
        "Produccion Pachacamac (KG)", "Produccion Pachacamac", "KG Producido Pachacamac",
        "Venta Neta ATE (KG)", "Venta Neta Pachacamac (KG)",
        "Planilla Total ATE", "Planilla Total Pachacamac",
        "Planilla ATE KG Producido", "Planilla Pachacamac KG Producido",
        # Alternativas
        "Planilla/kg", "S//kg", "Costo Planilla kg",
        "Eficiencia", "HH/kg", "Costo por kg", "S/ por kg",
        "Kg Producidos", "Kg Producción", "Kg Produccion",
        "Planilla Mensual", "Total Planilla", "Gasto Planilla",
        "Ratio Planilla", "Costo Mano de Obra",
        "MOD", "MOD/kg", "Mano de Obra",
    ],
    "planificacion": [
        "Fill Rate", "% Fill Rate", "Tasa Atención", "Tasa Atencion",
        "% Atención", "% Atencion", "Nivel de Servicio", "% Nivel Servicio",
        "Atendido", "% Atendido", "Cumplimiento Pedidos", "% Cumplimiento Pedidos",
        "Pedidos Atendidos", "Pedidos Total", "OC Atendidas",
        "Dead Stock", "% Dead Stock",
        "Avance", "% Avance", "Avance PPTO", "% Avance Presupuesto",
        "Presupuesto", "PPTO", "Ventas PPTO",
        "Ventas Real", "Real vs Ppto", "% Real vs Ppto",
    ],
    "consumo": [
        # Nombres REALES confirmados leyendo el panel Datos > Measures en Power BI Service
        # (Editar informe) — "14. Consumo Materiales indirectos de produccion", 2026-09-03.
        # No validados con DAX Query View (solo lectura de panel) — confirmar antes de usarlos
        # para decisiones críticas.
        "Costo total validado",     # -> Costo Total
        "Peso total KG",            # -> Venta Neta (KG)
        "Ratio costo / kg",         # -> Costo x TN Vendida
        "Producción (KG) Odoo",     # -> Producción Neta (KG)
        "Costo x ton producida",    # -> Costo x TN Producida
        # Medidas del reporte "Consumo de Materiales" (PRODUCCIÓN) — nombres viejos adivinados,
        # se mantienen como fallback:
        "CECO AJUSTADO", "Ceco Ajustado", "CECO", "Costo CECO",
        "MIP", "MIP (S/.)", "MIP Total", "Costo x TN", "Costo por TN",
        "TN PRODUCIDA", "TN Producida", "Toneladas Producidas", "Ton Producida",
        "Consumo Ate", "Consumo Pachacamac", "Consumo Planta",
        "Costo Consumo", "Costo Material", "Costo MP",
        # Medidas ratio consumo/compra
        "Consumo", "Total Consumo", "Importe Consumo", "Monto Consumo",
        "Compras", "Total Compras", "Importe Compras", "Monto Compras",
        "Ratio", "Ratio Consumo", "% Ratio", "Eficiencia", "% Eficiencia",
        "Ratio C/C", "Ratio Compras", "Consumo vs Compras",
        "Cant Compra", "Cant Consumo",
        "Compra Periodo", "Consumo Periodo",
    ],
    "fill_rate": [
        "% FILLRATE", "% Fill Rate", "% FillRate",
        "ORDEN DE VENTA", "FACTURACION", "VENTA PERDIDA",
        "Orden de Venta", "Facturacion", "Venta Perdida",
    ],
}

# ─────────────────────────────────────────────────────────────────────────────

def get_dataset_schema(token, ws_id, dataset_id):
    """Obtiene el schema completo del dataset via REST API (tablas + columnas + medidas)."""
    url = f"{PBI_BASE}/groups/{ws_id}/datasets/{dataset_id}/tables"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=20)
    if not r.ok:
        return {}
    schema = {}
    for tbl in r.json().get("value", []):
        tbl_name = tbl.get("name", "")
        cols = [c["name"] for c in tbl.get("columns", [])]
        measures = [m["name"] for m in tbl.get("measures", [])]
        schema[tbl_name] = {"columns": cols, "measures": measures}
    return schema

def get_token():
    r = requests.post(TOKEN_URL, data={
        "grant_type": "password", "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET, "username": USERNAME,
        "password": PASSWORD, "scope": PBI_SCOPE,
    }, timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]

# Errores de consultas DAX, para poder diagnosticar sin acceso al log del
# workflow (descargarlo requiere permisos de admin del repositorio). Se
# vuelcan al final en summaries.json, bajo "diagnostico".
DIAGNOSTICO = []

# Toda cifra que la app muestra y que NO viene de una consulta a Power BI,
# sino de una operación hecha aquí.
#
# Existe porque el margen por unidad de negocio se publicó durante semanas
# como si fuera del reporte: se calculaba (precio − costo) / precio, y en
# agosto daba 65.3% donde el reporte marcaba 62.30%. Nadie podía notarlo
# mirando la app. Una cifra derivada puede ser útil; presentarla como leída
# no lo es.
#
# Cada anotación viaja a summaries.json, la app las marca en pantalla, y
# cualquiera puede auditar la lista completa sin leer el código.
DERIVADOS = []


# Razón que se repite: una participación sobre el total no puede contradecir
# al reporte, porque es una proporción de sus propias cifras. Se distingue del
# caso grave —el margen por UEN— donde el reporte tiene su propia medida y da
# otro número.
PART = ("participación sobre el total, calculada sumando las filas del propio "
        "reporte. No puede contradecirlo: es una proporción de sus cifras.")


# Período real de cada KPI, leído de los filtros de la consulta capturada.
# Las tarjetas de Consumo y Productividad filtran por 'Calendario'[Año] y NO
# por [MES]: son acumulados del año, no del mes. Estaban publicadas junto a
# cifras de agosto sin decirlo, y eso hacía comparar un mes contra ocho.
#
# La de Productividad filtra además 'Tabla_Almacen_Planta'[Planta] = "ATE".
# Por eso los kilos vendidos salían 16.75M en un reporte y 6.79M en el otro:
# todas las plantas contra una sola. No era un error, era otro alcance.
PERIODO_KPI = {
    ("consumo_materiales", "Venta Neta (KG)"):
        ("acumulado 2026", "la consulta filtra solo por año, sin mes"),
    ("consumo_materiales", "Producción Neta (KG)"):
        ("acumulado 2026", "la consulta filtra solo por año, sin mes"),
    ("consumo_materiales", "Costo Total"):
        ("acumulado 2026", "la consulta filtra solo por año, sin mes"),
    ("productividad", "Venta Neta (KG)"):
        ("acumulado 2026 · solo planta ATE",
         "la consulta filtra por año y por Planta = ATE"),
}


def etiquetar_periodos(summary):
    """Pone el período en el nombre del KPI cuando no es del mes."""
    for edatos in (summary.get("empresas") or {}).values():
        for tipo, rep in (edatos.get("reportes") or {}).items():
            for k in rep.get("kpis", []):
                par = PERIODO_KPI.get((tipo, k.get("label")))
                if not par:
                    continue
                sufijo, razon = par
                k["label"] = f"{k['label']} · {sufijo}"
                k["nota_periodo"] = razon
                print(f"    · período: {tipo}/{k['label']}")


def anotar_derivado(reporte, desglose, campo, formula, razon):
    DERIVADOS.append({"reporte": reporte, "desglose": desglose, "campo": campo,
                      "formula": formula, "razon": razon})


CATALOGO_CAPTURAS = Path("capturas/catalogo.json")
_CATALOGO = None


def catalogo_capturas():
    """Consultas exportadas del Analizador, por nombre de visual.

    Se cachea: el archivo se lee una vez aunque lo pidan varios reportes.
    """
    global _CATALOGO
    if _CATALOGO is not None:
        return _CATALOGO
    _CATALOGO = {}
    if CATALOGO_CAPTURAS.exists():
        try:
            for f in json.loads(CATALOGO_CAPTURAS.read_text(encoding="utf-8")):
                v, dax = f.get("visual"), f.get("dax")
                if not v or not dax:
                    continue
                # Se indexa por visual+hash: varias tarjetas comparten título
                # y quedarnos con una sola perdería las demás.
                _CATALOGO[f"{v}#{f.get('hash','')}"] = f
                prev = _CATALOGO.get(v)
                if prev is None or (f.get("filas") or 0) > (prev.get("filas") or 0):
                    _CATALOGO[v] = f
        except Exception as e:
            print(f"    · catálogo de capturas ilegible: {e}")
    return _CATALOGO


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

def espera_throttle(r):
    """Segundos a esperar cuando Power BI responde 429, según lo que él pide.

    El límite es por espacio de trabajo, así que basta con que la corrida
    crezca un poco para cruzarlo. Power BI dice exactamente cuánto esperar
    —a veces en la cabecera Retry-After, a veces solo en el texto ("Retry in
    32 seconds")— y hacerle caso es la diferencia entre perder el dato o
    tenerlo treinta segundos más tarde.
    """
    cab = r.headers.get("Retry-After")
    if cab and str(cab).strip().isdigit():
        return min(int(cab), 90)
    m = re.search(r"[Rr]etry in (\d+) second", r.text or "")
    return min(int(m.group(1)) + 2, 90) if m else 15


def _tablas_dax(token, ws, dataset_id, query, label):
    """Como dax(), pero devuelve TODAS las tablas del resultado.

    Las consultas del Analizador suelen traer dos EVALUATE (eje y cuerpo).
    """
    import time
    url = f"{PBI_BASE}/groups/{ws}/datasets/{dataset_id}/executeQueries"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {"queries": [{"query": query}], "serializerSettings": {"includeNulls": True}}
    # Sin estos reintentos, un 429 borraba el desglose entero: la corrida del
    # 9 de setiembre perdió los quince jefes de área y los planes por planta
    # porque tres consultas nuevas empujaron el total sobre el límite.
    for intento in range(3):
        try:
            r = requests.post(url, json=body, headers=headers, timeout=90)
            if r.status_code == 429:
                if intento == 2:
                    DIAGNOSTICO.append({"consulta": label, "http": 429,
                                        "error": "límite de peticiones tras 3 intentos"})
                    return []
                espera = espera_throttle(r)
                print(f"    · {label}: límite de peticiones, esperando {espera}s")
                time.sleep(espera)
                continue
            if r.status_code != 200:
                DIAGNOSTICO.append({"consulta": label, "http": r.status_code,
                                    "error": r.text[:400]})
                return []
            return [t.get("rows", []) for t in r.json()["results"][0].get("tables", [])]
        except Exception as e:
            DIAGNOSTICO.append({"consulta": label, "http": 0, "error": repr(e)[:300]})
            return []
    return []


def dax_sku_por_uen():
    """Arma la consulta que cruza SKU con unidad de negocio.

    Ninguna captura del Analizador de rendimiento hace ese cruce: la que trae
    productos (86 filas) los da sin unidad de negocio, y la que sí trae unidad
    de negocio (113 filas) llega solo hasta subcategoría. Pero las dos salen
    del MISMO modelo y de la misma tabla de facturas, así que el cruce existe
    en el modelo aunque no exista en ningún visual.

    Lo único que se escribe a mano es la lista de columnas por las que se
    agrupa. Los FILTROS —que son la parte delicada, y la causa del error de
    septiembre con 'Peso total'— se copian literalmente de la captura, sin
    tocarlos: si el reporte excluye chatarra, servicios y facturas anuladas,
    esta consulta excluye exactamente lo mismo. Reescribirlos a mano sería
    repetir aquel error.

    Devuelve None si la captura no está disponible, para que la corrida siga
    sin este detalle en vez de morirse.
    """
    # El índice por hash guarda una lista: el mismo visual puede haberse
    # exportado varias veces. Cualquiera de esas copias sirve.
    copias = _catalogo_por_hash().get("5bb77c922cd9") or []
    if not copias:
        return None
    dax = (copias[0].get("dax") or "")
    corte = dax.find("VAR __DS0Core")
    if corte < 0 or "DEFINE" not in dax:
        return None
    cabecera = dax[:corte].rstrip()          # DEFINE + todos los VAR de filtro

    # Los nombres de los filtros del reporte, en el orden en que los define.
    filtros = [m for m in re.findall(r"VAR (__DS0FilterTable\d*)", cabecera)]
    if not filtros:
        return None

    # La tabla de fechas es local del modelo y su nombre lleva un GUID: se lee
    # de la propia captura en vez de fijarlo, porque cambia si republican.
    m = re.search(r"'(LocalDateTable_[0-9a-f-]+)'", dax)
    if not m:
        return None
    fecha = m.group(1)

    cols = ",\n        ".join(f"{f}" for f in filtros)
    return f"""{cabecera}

    VAR __SKU =
        SUMMARIZECOLUMNS(
        'Exl A Maestra de Facturas de Venta'[DESCRIPCION],
        'Exl Tipo de Negocio'[TIPO DE NEGOCIO N1],
        '{fecha}'[Año],
        '{fecha}'[NroMes],
        {cols},
        "Venta", CALCULATE(SUM('Exl A Maestra de Facturas de Venta'[Monto_Neto_Factura TG 0])),
        "Costo", CALCULATE(SUM('Exl A Maestra de Facturas de Venta'[Costo Total])),
        "Peso", CALCULATE(SUM('Exl A Maestra de Facturas de Venta'[Peso total]))
        )

EVALUATE
    __SKU
"""

def desglose_desde_captura(token, ws, candidatos, visual, columnas, limite=None):
    """Ejecuta una consulta capturada y devuelve sus filas como diccionarios.

    Se envía el DAX exportado SIN modificarlo: transcribir consultas a mano fue
    la causa de varios errores en este proyecto.

    `candidatos` son los datasets donde puede vivir la consulta — no se puede
    saber de la exportación, así que se prueba hasta que una devuelve filas.
    `columnas` mapea {nombre_de_salida: sufijo a buscar en la clave}, porque
    Power BI devuelve las claves como "Tabla[Columna]" o "[Alias]".
    """
    entrada = catalogo_capturas().get(visual)
    if not entrada:
        print(f"    · '{visual}': no está en el catálogo de capturas")
        return []

    def busca(fila, sufijo):
        for k, v in fila.items():
            if k.endswith(sufijo):
                return v
        return None

    # Estas consultas devuelven VARIAS tablas: el eje, los subtotales y el
    # cuerpo. Quedarse con la última es un error — en los visuales de matriz
    # jerárquica la última suele ser el nivel superior, sin la dimensión que
    # interesa. Se elige la tabla que realmente contenga las columnas pedidas.
    filas, todas = [], []
    for ds in candidatos:
        if not ds:
            continue
        tablas = tablas_de_captura(
            lambda q, lb: (_tablas_dax(token, ws, ds, q, lb) or [[]])[0],
            entrada["dax"], f"captura:{visual}")
        if not tablas:
            continue
        todas = tablas
        mejor, mejor_n = None, 0
        for t in tablas:
            if not t:
                continue
            ks = {k for f in t[:20] for k in f}
            n = sum(1 for suf in columnas.values()
                    if any(k.endswith(suf) for k in ks))
            if n > mejor_n or (n == mejor_n and mejor and len(t) > len(mejor)):
                mejor, mejor_n = t, n
        if mejor and mejor_n:
            filas = mejor
            break
    if not filas:
        # Sin esto el fallo es mudo: la consulta corre, no mapea nada y el
        # desglose sale vacío sin explicación. Se deja constancia de qué
        # columnas devolvió Power BI para poder corregir el mapa.
        disponibles = sorted(todas[-1][0].keys()) if (todas and todas[-1]) else []
        print(f"    · '{visual}': ninguna tabla trae las columnas pedidas")
        DIAGNOSTICO.append({
            "consulta": f"captura:{visual}", "http": 200,
            "error": ("ninguna tabla del resultado contiene las columnas "
                      f"{list(columnas.values())}. Devueltas: {disponibles[:25]}")})
        return []

    # Aviso de mapeo parcial: si una columna pedida no existe en la tabla
    # elegida, el desglose sale incompleto o vacío sin que nada falle. Pasó
    # con 'faltantes', que mapeaba el producto pero no la cantidad y quedaba
    # en cero. Se deja constancia con las claves reales para poder corregir.
    # Se comprueba la EXISTENCIA de la clave, no su valor: la primera fila
    # suele ser un subtotal del visual y trae la dimensión en blanco. Mirar el
    # valor daba por ausentes columnas que sí estaban.
    claves = {k for f in filas[:20] for k in f}
    for t in todas:
        if t:
            claves |= {k for k in t[0]}
    sin_mapear = [f"{n} ({suf})" for n, suf in columnas.items()
                  if not any(k.endswith(suf) for k in claves)]
    if sin_mapear:
        print(f"    · '{visual}': sin mapear {sin_mapear}")
        DIAGNOSTICO.append({
            "consulta": f"captura:{visual}", "http": 200,
            "error": (f"columnas sin mapear: {sin_mapear}. "
                      f"Claves reales: {sorted(filas[0].keys())[:25]}")})

    # En las matrices, las columnas del encabezado no viajan en el cuerpo:
    # se sustituyen por un [ColumnIndex] y sus valores van en la primera
    # tabla. 'ESTADO DE PLANES DE ACCION' abre por Planta en las filas y por
    # Estatus en las columnas, así que sin resolver el eje el estatus se
    # perdía.
    eje = []
    if filas and any(k.endswith("[ColumnIndex]") for k in filas[0]):
        for t in todas:
            if t and t is not filas and not any(k.endswith("[ColumnIndex]") for k in t[0]):
                eje = t
                break

    out = []
    for f in filas:
        fila = {nombre: busca(f, suf) for nombre, suf in columnas.items()}
        if eje:
            idx = busca(f, "[ColumnIndex]")
            if isinstance(idx, (int, float)) and 0 <= idx < len(eje):
                for nombre, suf in columnas.items():
                    if fila.get(nombre) is None:
                        fila[nombre] = busca(eje[int(idx)], suf)
        # Las filas de subtotal del visual llegan con la dimensión vacía.
        if all(v is None for v in fila.values()):
            continue
        out.append(fila)
    if limite:
        out = out[:limite]
    if not out:
        # Llegaron filas pero todas venian con las columnas pedidas en blanco
        # (son las de subtotal del visual). Sin dejar constancia, el desglose
        # aparece vacio en el dashboard y no hay por donde empezar a mirar.
        DIAGNOSTICO.append({
            "consulta": f"captura:{visual}", "http": 200,
            "error": (f"{len(filas)} filas devueltas, ninguna con datos en las "
                      f"columnas pedidas. Muestra: {filas[:2]}")})
    print(f"    ✓ '{visual}': {len(out)} filas")
    return out


def dax(token, ws_id, dataset_id, query, label="query", registrar=True):
    """Ejecuta una consulta DAX cruda. Retorna lista de filas o []."""
    import time
    url = f"https://api.powerbi.com/v1.0/myorg/groups/{ws_id}/datasets/{dataset_id}/executeQueries"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {"queries": [{"query": query}], "serializerSettings": {"includeNulls": True}}
    for attempt in range(3):
        try:
            r = requests.post(url, json=body, headers=headers, timeout=30)
            if r.status_code == 429:
                wait = espera_throttle(r)
                print(f"    [dax:{label}] throttled — esperando {wait}s...")
                time.sleep(wait)
                continue
            if r.status_code != 200:
                print(f"    [dax:{label}] {r.status_code} {r.text[:120]}")
                if registrar:
                    DIAGNOSTICO.append({"consulta": label, "http": r.status_code,
                                        "error": r.text[:600]})
                return []
            tables = r.json().get("results", [{}])[0].get("tables", [])
            filas = tables[0].get("rows", []) if tables else []
            if not filas and registrar:
                DIAGNOSTICO.append({"consulta": label, "http": 200,
                                    "error": "sin filas (la consulta corrió pero no devolvió datos)"})
            return filas
        except Exception as e:
            print(f"    [dax:{label}] error: {e}")
            if attempt < 2:
                time.sleep(3)
    return []

def try_measure(token, ws_id, dataset_id, name):
    """Prueba si una medida existe. Retorna (valor, True) o (None, False)."""
    r = requests.post(
        f"{PBI_BASE}/groups/{ws_id}/datasets/{dataset_id}/executeQueries",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"queries": [{"query": f'EVALUATE ROW("v", [{name}])'}],
              "serializerSettings": {"includeNulls": True}},
        timeout=20,
    )
    if not r.ok:
        return None, False
    rows = r.json().get("results", [{}])[0].get("tables", [{}])[0].get("rows", [])
    v = (rows[0].get("[v]") or rows[0].get("v")) if rows else None
    return v, True

MEASURE_CACHE_FILE = OUTPUT_DIR / "measure_cache.json"

def load_measure_cache():
    try:
        if MEASURE_CACHE_FILE.exists():
            return json.loads(MEASURE_CACHE_FILE.read_text())
    except: pass
    return {}

def save_measure_cache(cache):
    try:
        MEASURE_CACHE_FILE.write_text(json.dumps(cache, ensure_ascii=False, indent=2))
    except: pass

def scan_dataset(token, ws_id, dataset_id, ds_key, cache=None):
    """Escanea medidas reales. Aplica filtro de mes anterior si hay tabla de fecha conocida."""
    candidates = SCAN_CANDIDATES.get(ds_key, [])
    date_ctx = DATE_CONTEXT.get(ds_key)  # (tabla, columna) o None
    found = {}

    def get_val(name):
        if date_ctx:
            # registrar=False: el sondeo prueba nombres a ver cuáles existen;
            # que la mayoría falle es lo esperado, no un diagnóstico útil, y
            # además inflaba summaries.json (61 KB de ruido que carga la app).
            v = dax_prev_month(token, ws_id, dataset_id, name, date_ctx[0], date_ctx[1],
                               f"{ds_key}_{name[:15]}", registrar=False)
            if v is not None:
                return v, True
        return try_measure(token, ws_id, dataset_id, name)

    cached_found = (cache or {}).get(ds_key, {})
    if cached_found:
        print(f"    Verificando {len(cached_found)} medidas en caché...")
        for name in list(cached_found.keys()):
            v, exists = get_val(name)
            if exists:
                found[name] = v
                print(f"      ✓ [{name}] = {v} (caché)")
        scanned_before = set(cached_found.keys())
        candidates = [c for c in candidates if c not in scanned_before]
    print(f"    Escaneando {len(candidates)} candidatos nuevos...")
    for name in candidates:
        v, exists = get_val(name)
        if exists:
            found[name] = v
            print(f"      ✓ [{name}] = {v}")
    return found

def discover_tables_in_dataset(token, ws_id, dataset_id):
    """Descubre tablas del dataset probando nombres comunes. Retorna {table_name: [cols]}."""
    common_tables = [
        "Proveedores", "Facturas", "CxP", "Comprobantes", "Pagos",
        "Cuentas por Pagar", "Tabla CxP", "Detalle CxP",
        "Compras", "Ordenes de Compra", "OC", "Materiales",
        "Consumo", "Tabla Compras", "Detalle Compras",
        "Mermas", "Produccion", "Planilla", "Empleados",
        "Ventas", "Clientes", "Facturas Ventas",
        "Inventario", "Stock", "Articulos", "Productos",
        "Control Interno", "Auditoría", "Auditoria",
        "Medidas", "_Medidas", "KPIs",
        "Calendario", "Fecha", "Calendar",
        "fCalendario", "dCalendario",
    ]
    found_tables = {}
    for tbl in common_tables:
        r = requests.post(
            f"{PBI_BASE}/groups/{ws_id}/datasets/{dataset_id}/executeQueries",
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            json={"queries": [{"query": f"EVALUATE TOPN(3, '{tbl}')"}],
                  "serializerSettings": {"includeNulls": True}},
            timeout=15,
        )
        if r.ok:
            rows = r.json().get("results", [{}])[0].get("tables", [{}])[0].get("rows", [])
            if rows:
                cols = list(rows[0].keys())
                found_tables[tbl] = cols
                print(f"    Tabla '{tbl}': {cols[:5]}")
    return found_tables

def discover_all_datasets(token, ws_id):
    r = requests.get(f"{PBI_BASE}/groups/{ws_id}/datasets",
                     headers={"Authorization": f"Bearer {token}"}, timeout=20)
    if not r.ok: return {}
    return {d["name"]: d["id"] for d in r.json().get("value", [])}

def fmt_soles(v, decimals=0):
    if v is None: return "—"
    try:
        n = float(str(v).replace(",", ".").replace("%", "").replace("S/", "").replace(" ", ""))
        if abs(n) >= 1_000_000: return f"S/{n/1_000_000:.2f}M"
        elif abs(n) >= 1_000: return f"S/{n:,.0f}"
        return f"S/{n:.{decimals}f}"
    except: return str(v)

def fmt_pct(v):
    if v is None: return "—"
    try:
        n = float(str(v).replace(",", ".").replace("%", ""))
        if abs(n) < 1: n *= 100
        return f"{n:.1f}%"
    except: return str(v)

def to_float(v):
    if v is None: return None
    try:
        n = float(str(v).replace(",", ".").replace("%", ""))
        return n
    except: return None

def sem_thresh(v, red_above=None, yellow_above=None, red_below=None, yellow_below=None):
    n = to_float(v)
    if n is None: return "green"
    if abs(n) < 1 and (red_above or yellow_above or red_below or yellow_below):
        n *= 100
    if red_above and n >= red_above: return "red"
    if yellow_above and n >= yellow_above: return "yellow"
    if red_below and n <= red_below: return "red"
    if yellow_below and n <= yellow_below: return "yellow"
    return "green"

# ─── Funciones de reporte ─────────────────────────────────────────────────────


# KPIs de tarjeta tomados directamente del visual del reporte.
# (visual, reporte, etiqueta, formato, meta)
#
# Estos sustituyen a los del sondeo genérico, que consulta la medida SIN los
# filtros del visual. La diferencia no es teórica: en Margen el sondeo daba
# 49.5% donde el reporte muestra 46.5%, y S/2.96 de costo/kg donde son S/3.16.
KPIS_CAPTURADOS = [
    # ── 1. Cuentas por cobrar
    ("MOROSIDAD",           "cuentas_por_cobrar", "Morosidad",        "pct",   "15%"),
    ("CUENTAS POR COBRAR",  "cuentas_por_cobrar", "CxC Total",        "soles", None),
    ("POR VENCER",          "cuentas_por_cobrar", "CxC por vencer",   "soles", None),
    ("ROTACION CxC",        "cuentas_por_cobrar", "Rotación CxC",     "dias",  None),
    ("# CLIENTES",          "cuentas_por_cobrar", "Clientes",         "conteo", None),
    ("VENTAS MES ACTUAL",   "cuentas_por_cobrar", "Ventas del mes",   "soles", None),
    # El reporte 8 no tiene tarjetas: son tablas. Su desglose por jefe de
    # área y por planta se trae aparte, en el bloque de Control Interno.
    # ── 13. Productividad
    ("PRODUCCIÓN TOTAL (KG)", "productividad", "Producción Total (KG)", "kg",  None),
    ("PLANILLA TOTAL (S/.)",  "productividad", "Planilla Total",        "soles", None),
    ("VENTA NETA (KG)",       "productividad", "Venta Neta (KG)",       "kg",  None),
    # ── 14. Consumo de materiales indirectos
    ("COSTO TOTAL",          "consumo_materiales", "Costo Total",         "soles", None),
    ("PRODUCCIÓN NETA (KG)", "consumo_materiales", "Producción Neta (KG)", "kg",   None),
    ("COSTO X TN VENDIDA",   "consumo_materiales", "Costo x TN Vendida",  "soles", None),
    ("COSTO X TN PRODUCIDA", "consumo_materiales", "Costo x TN Producida", "soles", None),
]



# Tarjetas que comparten nombre de visual y solo se distinguen por sus filtros.
# En el reporte de CxP las siete tarjetas se llaman todas "Proyeccion
# Producción" —quedaron con el nombre de otro reporte al copiarlas— así que
# mapearlas por título es imposible. Se identifican por su firma: qué columnas
# filtra cada una y con qué valores.
#
# (archivo, [(columna, valor_esperado)…], columnas_prohibidas, reporte,
#  etiqueta, formato)
# (archivo, medida, [(columna, valor)…], columnas_prohibidas, reporte,
#  etiqueta, formato)
#
# La medida importa: en este reporte conviven tarjetas de importe y tarjetas
# de TEXTO con exactamente los mismos filtros, así que sin distinguirlas la
# firma es ambigua.
TARJETAS_POR_FIRMA = [
    ("02-cxp", "DISTINCTCOUNT", [], [],
     "cuentas_por_pagar", "# Proveedores", "conteo"),
    ("02-cxp", "SUM", [], ["ESTADO VIGENCIA", "ESTADO_REFINANCIADO", "RANGO"],
     "cuentas_por_pagar", "CxP Total", "soles"),
    ("02-cxp", "SUM", [("ESTADO_REFINANCIADO", "REFINANCIADO")], [],
     "cuentas_por_pagar", "Refinanciado", "soles"),
]


def _firma_de(dax):
    """Columnas filtradas por un DAX y el valor de cada una."""
    import re
    pares = {}
    for valores, col in re.findall(
            r"TREATAS\(\s*\{([^}]{0,200})\},\s*'[^']+'\[([^\]]+)\]", dax, re.S):
        pares[col] = re.sub(r"\s+", " ", valores).replace('"', "").strip()
    return pares


def tarjeta_por_firma(token, ws, candidatos, prefijo_archivo, medida,
                      requeridos, prohibidos):
    """Busca en el catálogo la tarjeta cuya firma de filtros coincide.

    `requeridos` son pares (columna, valor); valor None significa "que la
    consulta mencione esa columna, sin importar el valor" —así se identifica
    el conteo de proveedores, que usa DISTINCTCOUNT sobre [contacto]—.
    `prohibidos` son columnas que NO deben aparecer, que es lo que separa el
    total de la deuda de sus tramos.
    """
    vistos = set()
    for entrada in catalogo_capturas().values():
        if not entrada.get("archivo", "").startswith(prefijo_archivo):
            continue
        # Solo tarjetas: un visual de una fila. Sin esto la firma atrapa
        # también las tablas del reporte, que comparten filtros.
        if (entrada.get("filas") or 0) != 1:
            continue
        h = entrada.get("hash")
        if h in vistos:          # el catálogo indexa por visual y por hash
            continue
        vistos.add(h)
        dax = entrada["dax"]
        firma = _firma_de(dax)
        # Las columnas prohibidas se buscan solo en la FIRMA: rastrearlas en
        # todo el DAX daba falsos positivos, porque el nombre aparece también
        # en definiciones que la tarjeta no usa como filtro.
        if any(c in firma for c in prohibidos):
            continue
        if medida == "DISTINCTCOUNT" and "DISTINCTCOUNT" not in dax:
            continue
        if medida == "SUM" and ("DISTINCTCOUNT" in dax or "SUM(" not in dax):
            continue
        ok = True
        for col, val in requeridos:
            if val is None:
                if col not in dax:
                    ok = False
                    break
            elif firma.get(col) != val:
                ok = False
                break
        if not ok:
            continue
        for ds in candidatos:
            if not ds:
                continue
            tablas = tablas_de_captura(
                lambda q, lb: (_tablas_dax(token, ws, ds, q, lb) or [[]])[0],
                dax, f"firma:{prefijo_archivo}")
            for t in tablas:
                if not t:
                    continue
                for clave, valor in t[0].items():
                    if any(x in clave for x in ("IsGrandTotal", "IsDM",
                                                "ColumnIndex", "SortBy")):
                        continue
                    v = to_float(valor)
                    if v is not None:
                        return v
    return None


def valor_de_tarjeta(token, ws, candidatos, visual):
    """Valor único de un visual de tarjeta capturado del Analizador.

    Las tarjetas devuelven una sola fila con una sola medida. Se toma el
    primer valor numérico que no sea un indicador de subtotal ni un índice:
    así no hace falta saber de antemano el alias que le puso Power BI.
    """
    entrada = catalogo_capturas().get(visual)
    if not entrada:
        return None
    for ds in candidatos:
        if not ds:
            continue
        tablas = tablas_de_captura(
            lambda q, lb: (_tablas_dax(token, ws, ds, q, lb) or [[]])[0],
            entrada["dax"], f"tarjeta:{visual}")
        for t in tablas:
            if not t:
                continue
            for clave, valor in t[0].items():
                if any(x in clave for x in ("IsGrandTotal", "IsDM", "ColumnIndex",
                                            "SortBy", "[Año]", "[Mes")):
                    continue
                v = to_float(valor)
                if v is not None:
                    return v
    return None


def _fmt_kpi(valor, formato):
    if formato == "pct":
        v = valor * 100 if abs(valor) <= 1.5 else valor
        return f"{v:.1f}%"
    if formato == "dias":
        return f"{valor:.0f} días"
    if formato == "conteo":
        return f"{valor:,.0f}"
    if formato == "kg":
        return f"{valor:,.0f} kg"
    return fmt_soles(valor)


def aplicar_kpis_capturados(token, ws, candidatos_por_reporte, reportes):
    """Reemplaza los KPIs del sondeo por los de las tarjetas del reporte.

    Devuelve cuántos cambió. Los que no estén en el catálogo se saltan sin
    ruido: significa que esa pestaña todavía no se ha exportado.
    """
    cambios = 0
    for visual, tipo, etiqueta, formato, meta in KPIS_CAPTURADOS:
        rep = reportes.get(tipo)
        if rep is None or visual not in catalogo_capturas():
            continue
        v = valor_de_tarjeta(token, ws, candidatos_por_reporte.get(tipo, []), visual)
        if v is None:
            continue
        texto = _fmt_kpi(v, formato)
        kpis = rep.setdefault("kpis", [])
        actual = next((k for k in kpis if k["label"] == etiqueta), None)
        if actual is None:
            nuevo = {"label": etiqueta, "valor": texto}
            if meta:
                nuevo["meta"] = meta
            kpis.append(nuevo)
            cambios += 1
            print(f"    + [{tipo}] {etiqueta} = {texto} (tarjeta del reporte)")
        elif actual.get("valor") != texto:
            print(f"    ~ [{tipo}] {etiqueta}: {actual['valor']} → {texto} "
                  f"(tarjeta del reporte)")
            actual["valor"] = texto
            cambios += 1
    return cambios



# La lista de KPIs verificados se eliminó el 13/09/2026. Era una relación
# escrita a mano de qué cifras venían del reporte, y ese tipo de lista ya
# falló tres veces en este proyecto: nadie se acuerda de actualizarla y una
# cifra sin verificar queda marcada como verificada. Ahora marcar_origen()
# lo deduce comparando contra lo que devolvió el sondeo.



# De qué dataset se sondeó cada reporte, para poder reconocer después qué
# valores salieron del sondeo genérico.
DS_DE_REPORTE = {
    "cuentas_por_cobrar": "cxc", "cuentas_por_pagar": "cxp",
    "margen_variable": "margen", "mermas": "mermas", "compras": "compras",
    "sop_inventario": "inventario", "control_interno": "control_ds",
    "consumo_materiales": "consumo", "productividad": "productividad_ds",
    "fill_rate": "fill_rate", "margen_variable_pag2": "inventario",
}


def marcar_origen(reportes, scanned=None):
    """Anota en cada KPI si su cifra viene del reporte o del sondeo genérico.

    Antes esto era una lista escrita a mano, y ese es justo el tipo de lista
    que en este proyecto ya falló tres veces: nadie se acuerda de actualizarla
    y una cifra sin verificar queda marcada como verificada.

    Ahora se decide comparando: el sondeo genérico consulta las medidas SIN los
    filtros del visual, así que sus valores quedan guardados en `scanned`. Si
    el número que publica un KPI es exactamente uno de esos, salió de ahí. Si
    no coincide con ninguno, vino de una consulta capturada del reporte.

    La dirección del error es la segura: una cifra capturada que por
    casualidad coincida con la sondeada se marca como no verificada, nunca al
    revés.
    """
    ver = tot = 0
    for tipo, rep in reportes.items():
        ds = DS_DE_REPORTE.get(tipo, "")
        sondeados, del_reporte = [], []
        if scanned:
            crudos = scanned.get(ds, {}) or {}
            for nombre, v in crudos.items():
                if nombre.startswith("__"):
                    continue
                f = to_float(v)
                if f is None:
                    continue
                # Evidencia positiva: esta clave la llenó una consulta del
                # reporte, no el sondeo.
                (del_reporte if (ds, nombre) in CLAVES_DEL_REPORTE
                 else sondeados).append(f)
        for k in rep.get("kpis", []):
            tot += 1
            val = to_float(str(k.get("valor", "")).replace("S/", "")
                           .replace(",", "").replace("%", "").replace("d", ""))
            # Las cifras se publican formateadas (millones abreviados,
            # porcentajes ×100): se compara en varias escalas.
            def coincide(lista):
                return val is not None and any(
                    abs(val - v * e) <= max(abs(val), abs(v * e)) * 0.001
                    for v in lista for e in (1, 100, 0.01, 1e-6, 1e-3))

            # La evidencia positiva manda: si el número es el de una consulta
            # del reporte, viene del reporte aunque el sondeo dé lo mismo.
            del_sondeo = not coincide(del_reporte) and coincide(sondeados)
            k["fuente"] = "sondeo" if del_sondeo else "reporte"
            if not del_sondeo:
                ver += 1
    return ver, tot


def build_cxc(found):
    # Solo '% Morosidad' (21.08%). La medida 'Morosidad' (13.78%) queda fuera
    # a propósito — ver nota en SCAN_CANDIDATES["cxc"].
    mora_val = found.get("% Morosidad") or found.get("% Mora") or found.get("Tasa Morosidad")
    vencer_val = found.get("Por Vencer") or found.get("Saldo Por Vencer") or found.get("No Vencido")
    total_val = found.get("CxC Total") or found.get("Total CxC") or found.get("Saldo CxC") or found.get("Cartera Total")
    vencido_val = found.get("Vencido") or found.get("Saldo Vencido") or found.get("Total Vencido")

    mora_pct = to_float(mora_val)
    if mora_pct and abs(mora_pct) < 1: mora_pct *= 100
    sem = sem_thresh(mora_pct, red_above=20, yellow_above=15)

    kpis = []
    if mora_pct is not None:
        kpis.append({"label": "Morosidad", "valor": f"{mora_pct:.1f}%", "meta": "15%", "estado": sem})
    # El KPI "por vencer" se toma del aging cuando existe: la medida suelta
    # 'Por Vencer' la descubrió el escáner y se consulta con un contexto de
    # filtro que elegimos nosotros (el mes), no el que usa el reporte — por
    # eso devuelve 3.15M sin filtro y 3.25M con filtro de mes, mientras el
    # tramo "1. Por Vencer" del aging da 3.57M y es el único que cuadra:
    # 3.57M + 946K vencido = 4.52M total, y 946K/4.52M = 20.9% ≈ 21.08%
    # de la morosidad oficial. El aging viene de Copiar consulta; la medida
    # suelta, no. Ver dax_cxc_aging().
    aging_pre = found.get("__aging") or []
    vencer_aging = next((v for seg, v in aging_pre if "vencer" in seg.lower()), None)
    if vencer_aging is not None:
        vencer_val = vencer_aging
        kpis.append({"label": "CxC por vencer", "valor": fmt_soles(vencer_val),
                     "meta": "aging confirmado"})
    elif vencer_val is not None:
        kpis.append({"label": "CxC por vencer", "valor": fmt_soles(vencer_val)})
    if total_val is not None:
        kpis.append({"label": "CxC Total", "valor": fmt_soles(total_val)})
    if vencido_val is not None and not (found.get("__aging") or []):
        kpis.append({"label": "CxC Vencido", "valor": fmt_soles(vencido_val), "estado": "red" if sem == "red" else "yellow"})

    razon = f"Morosidad {fmt_pct(mora_pct)}" + (" — crítica" if sem=="red" else " — sobre meta" if sem=="yellow" else " — OK")
    alerta = razon if sem != "green" else None

    tramos = []
    # Aging real desde 'DATA_FACTURACION'[O_SEGMENTO] (ver dax_cxc_aging).
    # Las etiquetas vienen del modelo tal cual ("1. Por vencer", "4. Más de
    # 30 días", ...) — solo les quitamos el prefijo numérico de ordenamiento.
    aging = found.get("__aging") or []
    total_aging = sum(v for _, v in aging) or None
    for seg, v in aging:
        limpio = re.sub(r"^\s*\d+\.\s*", "", seg).strip()
        low = limpio.lower()
        if "vencer" in low:       s = "green"
        elif "más de" in low or "mas de" in low or "+" in low: s = "red"
        else:                     s = "yellow"
        item = {"label": limpio, "valor": fmt_soles(v), "estado": s}
        if total_aging:
            item["pct"] = round(v / total_aging * 100, 1)
        tramos.append(item)

    if total_aging is not None and not any(k["label"] == "CxC Total" for k in kpis):
        kpis.append({"label": "CxC Total", "valor": fmt_soles(total_aging)})
    if total_aging:
        venc = sum(v for seg, v in aging if "vencer" not in seg.lower())
        if venc:
            kpis.append({"label": "CxC Vencido", "valor": fmt_soles(venc),
                         "meta": f"{venc/total_aging*100:.1f}% de cartera",
                         "estado": "red" if sem == "red" else "yellow"})

    if tramos and any("pct" in t for t in tramos):
        anotar_derivado("cuentas_por_cobrar", "tramos", "pct",
                        "valor del tramo / cartera total × 100",
                        "participación sobre el total, calculada sumando las "
                        "filas del propio reporte. No puede contradecirlo: es "
                        "una proporción de sus cifras.")
    result = {"estado": sem, "alerta": alerta, "kpis": kpis}
    if tramos: result["tramos"] = tramos

    # Aging por canal, jefe de venta y ejecutivo. La matriz devuelve los tres
    # niveles mezclados con sus subtotales; se toma el nivel más fino que trae
    # cada fila, para que la mora tenga un responsable con nombre.
    resp = []
    for f in (found.get("__por_canal") or []):
        nombre = ((f.get("ejecutivo") or "").strip() or
                  (f.get("jefe") or "").strip() or
                  (f.get("canal") or "").strip())
        if not nombre:
            continue
        nivel = ("ejecutivo" if (f.get("ejecutivo") or "").strip() else
                 "jefe" if (f.get("jefe") or "").strip() else "canal")
        vencido = sum(v for v in [to_float(f.get("d0_15")), to_float(f.get("d16_30")),
                                  to_float(f.get("mas_30"))] if v is not None)
        total = to_float(f.get("total"))
        resp.append({"responsable": nombre, "nivel": nivel,
                     "canal": (f.get("canal") or "").strip() or None,
                     "total": fmt_soles(total) if total is not None else None,
                     "vencido": fmt_soles(vencido) if vencido else None,
                     "_v": vencido or 0})
    nominales = [x for x in resp if x["nivel"] != "canal"]
    if nominales:
        nominales.sort(key=lambda x: -x["_v"])
        for x in nominales:
            x.pop("_v", None)
        result["responsables"] = nominales[:15]

    return result, sem, razon, mora_pct, vencer_val

def dax_cxp_top_proveedores(token, ws, dataset_id, label="cxp_top15"):
    """Top 15 proveedores por deuda pendiente.

    Confirmado con Copiar consulta (2026-09-06) sobre el visual "TOP 15".

    La consulta tiene DOS niveles y ambos importan:
      1. __SQDS0Core agrupa por [contacto] y TOPN(15) se queda con los quince
         de mayor importe.
      2. El resultado se usa como FILTRO de la consulta externa, que agrupa
         por [NOMBRE ABREV] y [CATEGORIA] — que es lo que muestra la tabla.
    Agrupar directamente por nombre abreviado no daría lo mismo: el top se
    decide por contacto, y un contacto puede abrir varias filas de categoría.

    Los cinco filtros van literales. Definen qué es "deuda pendiente":
      · [contacto] excluye INVENTARIO y SEDAPAL
      · [EXISTE] = "SÍ"
      · [ESTADO DE PAGO] en {vacío, Pagado Parcialmente, Sin Pagar}
      · [CONSIDERACION] = 1
      · 'Calendario'[Columna Mostrar] = "MOSTRAR"

    Se quita solo la maquinaria de subtotales y ordenamiento
    (ROLLUPADDISSUBTOTAL, NATURALLEFTOUTERJOIN, SELECTCOLUMNS, los TOPN de
    presentación), que no altera los importes.
    """
    filtros = (
        "\tVAR __SQDS0FilterTable = \n"
        "\t\tFILTER(\n"
        "\t\t\tKEEPFILTERS(VALUES('CUENTAS CONTABLES'[contacto])),\n"
        "\t\t\tNOT(\n"
        "\t\t\t\t'CUENTAS CONTABLES'[contacto] IN {\"INVENTARIO\",\n"
        "\t\t\t\t\t\"SERV AGUA POTAB Y ALCANT DE LIMA-SEDAPAL\"}\n"
        "\t\t\t)\n"
        "\t\t)\n\n"
        "\tVAR __SQDS0FilterTable2 = \n"
        "\t\tTREATAS({\"SÍ\"}, 'CUENTAS CONTABLES'[EXISTE])\n\n"
        "\tVAR __SQDS0FilterTable3 = \n"
        "\t\tTREATAS(\n"
        "\t\t\t{BLANK(),\n"
        "\t\t\t\t\"Pagado Parcialmente\",\n"
        "\t\t\t\t\"Sin Pagar\"},\n"
        "\t\t\t'CUENTAS CONTABLES'[ESTADO DE PAGO]\n"
        "\t\t)\n\n"
        "\tVAR __SQDS0FilterTable4 = \n"
        "\t\tTREATAS({1}, 'CUENTAS CONTABLES'[CONSIDERACION])\n\n"
        "\tVAR __SQDS0FilterTable5 = \n"
        "\t\tTREATAS({\"MOSTRAR\"}, 'Calendario'[Columna Mostrar])\n\n"
    )
    def usados(ind):
        return "".join(f"{ind}__SQDS0FilterTable{'' if i == 1 else i},\n"
                       for i in range(1, 6))

    MED = ("\"SumIMPORTE_NETO__42_\", "
           "CALCULATE(SUM('CUENTAS CONTABLES'[IMPORTE NETO (42)]))")
    q = (
        "DEFINE\n" + filtros +
        "\tVAR __SQDS0Core = \n"
        "\t\tSUMMARIZECOLUMNS(\n"
        "\t\t\t'CUENTAS CONTABLES'[contacto],\n"
        + usados("\t\t\t") +
        f"\t\t\t{MED}\n"
        "\t\t)\n\n"
        "\tVAR __SQDS0BodyLimited = \n"
        "\t\tTOPN(15, __SQDS0Core, [SumIMPORTE_NETO__42_], 0)\n\n"
        "EVALUATE\n"
        "\tSUMMARIZECOLUMNS(\n"
        "\t\t'CUENTAS CONTABLES'[NOMBRE ABREV],\n"
        "\t\t'CUENTAS CONTABLES'[CATEGORIA],\n"
        "\t\t__SQDS0BodyLimited,\n"
        + usados("\t\t") +
        f"\t\t{MED}\n"
        "\t)\n\n"
        "ORDER BY\n\t[SumIMPORTE_NETO__42_] DESC"
    )
    rows = dax(token, ws, dataset_id, q, label)
    out = []
    for r in rows or []:
        nom = (r.get("CUENTAS CONTABLES[NOMBRE ABREV]") or r.get("[NOMBRE ABREV]"))
        cat = (r.get("CUENTAS CONTABLES[CATEGORIA]") or r.get("[CATEGORIA]"))
        v = to_float(r.get("[SumIMPORTE_NETO__42_]") or r.get("SumIMPORTE_NETO__42_"))
        if nom and v is not None:
            out.append((str(nom), str(cat or ""), v))

    if not out:
        # Respaldo: la corrida del 2026-09-06 devolvió vacío con la estructura
        # de dos niveles. Aquí se agrupa directo por [contacto] con los mismos
        # cinco filtros y se toman los quince mayores en Python. Da los mismos
        # proveedores y montos; lo que se pierde es la apertura por categoría,
        # porque esa venía de la consulta externa.
        print(f"      [{label}] sin filas — probando agrupación simple")
        q2 = (
            "DEFINE\n" + filtros +
            "EVALUATE\n"
            "\tSUMMARIZECOLUMNS(\n"
            "\t\t'CUENTAS CONTABLES'[contacto],\n"
            + usados("\t\t") +
            f"\t\t{MED}\n"
            "\t)\n\n"
            "ORDER BY\n\t[SumIMPORTE_NETO__42_] DESC"
        )
        for r in dax(token, ws, dataset_id, q2, label + "-simple") or []:
            nom = r.get("CUENTAS CONTABLES[contacto]") or r.get("[contacto]")
            v = to_float(r.get("[SumIMPORTE_NETO__42_]") or r.get("SumIMPORTE_NETO__42_"))
            if nom and v is not None:
                out.append((str(nom), "", v))

    out.sort(key=lambda t: -t[2])
    return out[:15]


def dax_cxp_aging(token, ws, dataset_id, label="cxp_aging"):
    """Tramos de vigencia de la deuda con proveedores.

    Confirmado con Copiar consulta (2026-09-06) sobre la tarjeta "VIGENTE"
    del reporte '2. Cuentas por pagar'. Sus seis filtros definen qué cuenta
    como deuda viva:

      · [CONSIDERACION] = 1
      · [ESTADO VIGENCIA] = "Vigente"   ← lo único que cambia entre tarjetas
      · [ESTADO DE PAGO] en {Pagado Parcialmente, Sin Pagar}
      · [EXISTE] = "SÍ"
      · [ESTADO_REFINANCIADO] = "NO"    ← lo refinanciado se reporta aparte
      · 'Calendario'[Columna Mostrar] = "MOSTRAR"

    Ojo con [ESTADO DE PAGO]: aquí son dos valores, sin BLANK(). La consulta
    del TOP 15 sí incluye BLANK() — son universos distintos y por eso cada
    una lleva su propia lista.

    Los tramos vencidos NO viven en [ESTADO VIGENCIA]: esa columna solo
    distingue "Vigente" de "No Vigente". El detalle está en [RANGO], con
    valores como "1. 0 a 7 días" y "2. 8 a 15 días" — confirmado con la
    consulta de la tarjeta "VENCIDO Menor a 15 días", que filtra
    [ESTADO VIGENCIA]="No Vigente" y agrupa dos rangos. Por eso se agrupa
    por AMBAS columnas: por una sola habría devuelto dos filas.

    Las tarjetas del reporte suman varios rangos ("Menor a 15 días" son los
    rangos 1 y 2). Aquí se publican los rangos tal como están en el modelo,
    que es más fino y no obliga a replicar ese agrupamiento.

    Devuelve [(etiqueta, importe), ...] de mayor a menor.
    """
    filtros = (
        "\tVAR __DS0FilterTable = \n"
        "\t\tTREATAS({1}, 'CUENTAS CONTABLES'[CONSIDERACION])\n\n"
        "\tVAR __DS0FilterTable2 = \n"
        "\t\tTREATAS({\"Pagado Parcialmente\",\n"
        "\t\t\t\"Sin Pagar\"}, 'CUENTAS CONTABLES'[ESTADO DE PAGO])\n\n"
        "\tVAR __DS0FilterTable3 = \n"
        "\t\tTREATAS({\"SÍ\"}, 'CUENTAS CONTABLES'[EXISTE])\n\n"
        "\tVAR __DS0FilterTable4 = \n"
        "\t\tTREATAS({\"NO\"}, 'CUENTAS CONTABLES'[ESTADO_REFINANCIADO])\n\n"
        "\tVAR __DS0FilterTable5 = \n"
        "\t\tTREATAS({\"MOSTRAR\"}, 'Calendario'[Columna Mostrar])\n\n"
    )
    q = (
        "DEFINE\n" + filtros +
        "EVALUATE\n"
        "\tSUMMARIZECOLUMNS(\n"
        "\t\t'CUENTAS CONTABLES'[ESTADO VIGENCIA],\n"
        "\t\t'CUENTAS CONTABLES'[RANGO],\n"
        "\t\t__DS0FilterTable,\n"
        "\t\t__DS0FilterTable2,\n"
        "\t\t__DS0FilterTable3,\n"
        "\t\t__DS0FilterTable4,\n"
        "\t\t__DS0FilterTable5,\n"
        "\t\t\"SumIMPORTE_NETO__42_\", IGNORE(\n"
        "\t\t\tCALCULATE(SUM('CUENTAS CONTABLES'[IMPORTE NETO (42)]))\n"
        "\t\t)\n"
        "\t)\n\n"
        "ORDER BY\n\t[SumIMPORTE_NETO__42_] DESC"
    )
    rows = dax(token, ws, dataset_id, q, label)
    out = []
    for r in rows or []:
        est = (r.get("CUENTAS CONTABLES[ESTADO VIGENCIA]") or r.get("[ESTADO VIGENCIA]"))
        rango = (r.get("CUENTAS CONTABLES[RANGO]") or r.get("[RANGO]"))
        v = to_float(r.get("[SumIMPORTE_NETO__42_]") or r.get("SumIMPORTE_NETO__42_"))
        if not est or v is None:
            continue
        # Lo vigente se etiqueta por su estado; lo vencido, por su rango de
        # días, que es la información útil. El prefijo numérico del rango
        # ("1. ", "2. ") solo ordena, así que se retira.
        if "no vigente" in str(est).lower() and rango:
            etiqueta = re.sub(r"^\s*\d+\.\s*", "", str(rango)).strip()
        else:
            etiqueta = str(est)
        out.append((etiqueta, v))

    # [RANGO] también subdivide lo vigente (por antigüedad de la factura), así
    # que "Vigente" vuelve en varias filas. Se consolidan: el reporte muestra
    # un solo VIGENTE, y su suma coincide con la tarjeta.
    agrupado = {}
    for etiqueta, v in out:
        agrupado[etiqueta] = agrupado.get(etiqueta, 0.0) + v
    return sorted(agrupado.items(), key=lambda t: -t[1])


def build_cxp(found):
    # Nombres confirmados: "Cuentas x Pagar", "Refinanciamiento", "Proveedores"
    total_val = (found.get("Cuentas x Pagar") or found.get("CUENTAS X PAGAR") or
                 found.get("CxP Total") or found.get("Total CxP") or found.get("Saldo CxP") or
                 found.get("Deuda Total"))
    refin_val = (found.get("Refinanciamiento") or found.get("Refinanciado") or
                 found.get("Saldo Refinanciado"))
    proveed_val = found.get("Proveedores") or found.get("Nro Proveedores")
    # Tramos vencido
    vigente_val = found.get("Vigente CxP") or found.get("CxP Vigente") or found.get("Saldo Vigente")
    venc15_val  = found.get("Vencido < 15") or found.get("Vencido Menor 15d") or found.get("Vencido 0-15d")
    venc30_val  = found.get("Vencido 16-30d") or found.get("Vencido 30d") or found.get("Vencido 1-30")
    venc90_val  = found.get("Vencido 31-90") or found.get("Vencido 31-90d")
    # Días CxP
    dias_val = (found.get("Dias CxP") or found.get("DPP") or found.get("Dias Pago") or
                found.get("Rotacion CxP") or found.get("Dias Promedio Pago") or
                found.get("Dias de Pago") or found.get("Días de Pago") or
                found.get("PPP") or found.get("Plazo Pago"))

    total = to_float(total_val)
    refin = to_float(refin_val)
    dias  = to_float(dias_val)

    # Semáforo: si refinanciado > 30% del total, amarillo; dias > 90, rojo
    sem = "green"
    if dias and dias > 90: sem = "red"
    elif dias and dias > 60: sem = "yellow"
    elif refin and total and refin / total > 0.30: sem = "yellow"

    kpis = []
    if total_val is not None:
        kpis.append({"label": "CxP Total", "valor": fmt_soles(total_val)})
    if vigente_val is not None:
        kpis.append({"label": "Vigente", "valor": fmt_soles(vigente_val), "estado": "green"})
    if venc15_val is not None:
        kpis.append({"label": "Vencido <15d", "valor": fmt_soles(venc15_val), "estado": "yellow"})
    if venc30_val is not None:
        kpis.append({"label": "Vencido 16-30d", "valor": fmt_soles(venc30_val), "estado": "red"})
    if venc90_val is not None:
        kpis.append({"label": "Vencido 31-90d", "valor": fmt_soles(venc90_val), "estado": "red"})
    if refin_val is not None:
        s = "yellow" if refin and total and refin/total > 0.30 else "green"
        kpis.append({"label": "Refinanciado", "valor": fmt_soles(refin_val), "estado": s})
    if proveed_val is not None:
        kpis.append({"label": "# Proveedores", "valor": str(int(to_float(proveed_val) or 0))})
    if dias is not None:
        s_d = "red" if dias > 90 else ("yellow" if dias > 60 else "green")
        kpis.append({"label": "Días CxP", "valor": f"{dias:.0f}d", "meta": "<90d", "estado": s_d})

    alerta = f"CxP {dias:.0f}d — revisar flujo" if dias and dias > 60 else (
             f"Refinanciado {fmt_soles(refin_val)} — gestionar" if refin and total and refin/total > 0.30 else None)
    res = {"estado": sem, "alerta": alerta, "kpis": kpis}

    # Tramos de vigencia de la deuda (ver dax_cxp_aging). Excluyen lo
    # refinanciado, que el reporte muestra como tarjeta aparte.
    aging = found.get("__aging_cxp") or []
    if aging:
        tot_ag = sum(v for _, v in aging) or None
        def color(nombre):
            # Las etiquetas vencidas son rangos de días ("0 a 7 días",
            # "31 a 90 días"): se lee el primer número del rango para
            # decidir el semáforo, en vez de buscar textos sueltos.
            n = nombre.lower()
            if "vigente" in n:
                return "green"
            m = re.search(r"(\d+)", n)
            dias = int(m.group(1)) if m else 0
            return "red" if dias >= 31 else "yellow"
        anotar_derivado("cuentas_por_pagar", "tramos", "pct",
                        "valor del tramo / suma de tramos × 100", PART)
        res["tramos"] = [{
            "label": est,
            "valor": fmt_soles(v),
            "pct": round(v / tot_ag * 100, 1) if tot_ag else None,
            "estado": color(est),
        } for est, v in aging]
        vencido = sum(v for est, v in aging if "vigente" not in est.lower())
        if vencido and tot_ag:
            res["kpis"].append({
                "label": "CxP Vencido", "valor": fmt_soles(vencido),
                "meta": f"{vencido / tot_ag * 100:.1f}% de la deuda viva",
                "estado": "red" if vencido / tot_ag > 0.3 else "yellow"})

    # Top 15 proveedores por deuda pendiente (ver dax_cxp_top_proveedores).
    # Campos 'nombre' y 'monto': son los que ya consume juanito.html.
    top = found.get("__top_proveedores") or []
    if top:
        tot = sum(v for _, _, v in top) or None
        anotar_derivado("cuentas_por_pagar", "proveedores_criticos", "pct",
                        "deuda del proveedor / deuda del top 15 × 100",
                        "participación sobre el total, calculada sumando las "
                        "filas del propio reporte. Ojo: el denominador es el "
                        "top 15, no la deuda total de la empresa.")
        res["proveedores_criticos"] = [{
            "nombre": n,
            "categoria": c,
            "monto": fmt_soles(v),
            "pct": round(v / tot * 100, 1) if tot else None,
        } for n, c, v in top[:15]]
        if tot:
            res["kpis"].append({
                # La etiqueta dice 'Deuda' y no 'Top 15' a secas: un KPI que
                # dice "Top 15 proveedores = S/12.05M" se lee como si fueran
                # doce millones de proveedores.
                "label": "Deuda top 15 proveedores", "valor": fmt_soles(tot),
                "meta": "concentración"})
    return res, dias

def build_margen(found):
    margen_val = (found.get("% Margen") or found.get("R. MARGEN") or found.get("R. Margen") or
                  found.get("Margen Variable") or found.get("% Margen Variable") or
                  found.get("MV") or found.get("% MV") or found.get("MV%") or
                  found.get("Margen Variable %") or found.get("Margen") or
                  found.get("% MV PAUNO") or found.get("Pct MV"))
    costo_total_val = found.get("Costo Total")
    ventas_val = (found.get("Ventas Mes Actual") or found.get("Ventas Actuales") or
                  found.get("Ventas") or found.get("Venta Total") or found.get("Venta Neta"))
    precio_val = (found.get("Precio/kg") or found.get("Precio kg") or found.get("Precio Promedio kg") or
                  found.get("Precio x Kilo") or found.get("Precio Kilo"))
    costo_val  = (found.get("Costo/kg") or found.get("Costo kg") or
                  found.get("R. COSTO UNIT") or found.get("Costo x Kilo") or found.get("Costo Kilo"))

    margen_pct = to_float(margen_val)
    if margen_pct and abs(margen_pct) < 1: margen_pct *= 100
    # Si no hay medida directa de margen, calcular desde Precio/kg y Costo/kg
    if margen_pct is None and precio_val is not None and costo_val is not None:
        p, c = to_float(precio_val), to_float(costo_val)
        if p and c and p > 0:
            margen_pct = (p - c) / p * 100
    sem = sem_thresh(margen_pct, red_below=46, yellow_below=52)

    kpis = []
    if ventas_val is not None:
        kpis.append({"label": "Ventas mes", "valor": fmt_soles(ventas_val)})
    if margen_pct is not None:
        kpis.append({"label": "Margen variable", "valor": f"{margen_pct:.1f}%", "meta": "52%", "estado": sem})
    if precio_val is not None:
        kpis.append({"label": "Precio/kg", "valor": f"S/{to_float(precio_val):.2f}"})
    if costo_val is not None:
        kpis.append({"label": "Costo/kg", "valor": f"S/{to_float(costo_val):.2f}"})
    if costo_total_val is not None:
        kpis.append({"label": "Costo Total", "valor": fmt_soles(costo_total_val)})

    alerta = f"Margen {fmt_pct(margen_pct)} — bajo meta 52%" if sem != "green" and margen_pct else None
    res = {"estado": sem, "alerta": alerta, "kpis": kpis}

    # Clientes que más venta concentran, con su margen. El reporte trae 173
    # líneas; en el dashboard entran las de mayor venta, que es donde una
    # caída de margen pesa de verdad.
    cli = found.get("__por_cliente") or []
    limpios = []
    for c in cli:
        v = to_float(c.get("venta"))
        nombre = (c.get("cliente") or "").strip()
        if not nombre or v is None or v <= 0:
            continue
        limpios.append((nombre, v, to_float(c.get("margen")),
                        to_float(c.get("margen_caida"))))
    if cli and not limpios:
        # Llegaron filas pero ninguna paso el filtro. Sin una muestra no hay
        # forma de saber si el problema es el nombre vacio, la venta en cero
        # o un signo invertido.
        DIAGNOSTICO.append({"consulta": "margen_cliente_vacio", "http": 200,
                            "error": f"{len(cli)} filas, ninguna util. "
                                     f"Muestra: {cli[:3]}"})
    if limpios:
        limpios.sort(key=lambda t: -t[1])
        total_v = sum(t[1] for t in limpios) or None
        anotar_derivado("margen_variable", "por_cliente", "pct_venta",
                        "venta del cliente / venta total × 100", PART)
        res["por_cliente"] = [{
            "cliente": n,
            "venta": fmt_soles(v),
            "pct_venta": round(v / total_v * 100, 1) if total_v else None,
            "margen": (f"{(mg * 100 if abs(mg) <= 1 else mg):.1f}%"
                       if mg is not None else "—"),
            "caida": (f"{(ca * 100 if abs(ca) <= 1 else ca):.1f}%"
                      if ca is not None else None),
        } for n, v, mg, ca in limpios[:12]]

    # El costo por cliente viene de la tabla de órdenes de venta, que trae los
    # mismos clientes con las mismas ventas y márgenes. Se añade aquí en vez de
    # publicar un segundo desglose idéntico.
    costos = {}
    for o in (found.get("__ordenes_cliente") or []):
        nombre = (o.get("cliente") or "").strip()
        c = to_float(o.get("costo"))
        if nombre and c is not None:
            costos[nombre] = c
    for fila in res.get("por_cliente", []):
        c = costos.get(fila["cliente"])
        if c is not None:
            fila["costo"] = fmt_soles(c)

    # Detalle de costos por categoría. Se ordena por venta: lo que mueve el
    # margen es donde hay volumen, no la categoría con el peor porcentaje.
    det = []
    for f in (found.get("__detalle_costos") or []):
        cat = (f.get("categoria") or "").strip()
        precio = to_float(f.get("precio_kg"))
        costo = to_float(f.get("costo_kg"))
        venta = to_float(f.get("venta"))
        if not cat or precio is None or not precio:
            continue
        det.append({
            "categoria": cat,
            "subcategoria": (f.get("subcategoria") or "").strip() or None,
            "negocio": (f.get("negocio") or "").strip() or None,
            "venta": fmt_soles(venta) if venta is not None else None,
            "precio_kg": f"S/{precio:.2f}",
            "costo_kg": f"S/{costo:.2f}" if costo is not None else None,
            "margen": (f"{(precio - costo) / precio * 100:.1f}%"
                       if costo is not None else None),
            "_v": abs(venta or 0),
        })
    if det:
        anotar_derivado("margen_variable", "detalle_costos", "margen",
                        "(precio_kg − costo_kg) / precio_kg",
                        "el detalle publica precio y costo por kilo pero no el "
                        "margen de cada línea; se deriva de sus dos cifras")
        det.sort(key=lambda x: -x["_v"])
        for x in det:
            x.pop("_v", None)
        res["detalle_costos"] = det[:40]

    # ── Qué producto movió el margen entre los dos últimos meses con dato.
    #
    # El orden es por PUNTOS DE MARGEN PERDIDOS SOBRE EL TOTAL, no por caída
    # porcentual: un producto que se desploma pero pesa el 0.3% de la venta no
    # explica nada, y encabezaría la lista si se ordenara por porcentaje.
    porprod = found.get("__por_producto") or []
    if porprod:
        meses = {}
        for f in porprod:
            nombre = (f.get("producto") or "").strip()
            v, c = to_float(f.get("venta")), to_float(f.get("costo"))
            per = (to_float(f.get("anio")), to_float(f.get("mes")))
            if not nombre or v is None or c is None or None in per or not v:
                continue
            meses.setdefault(per, {})[nombre] = (v, c)
        orden = sorted(k for k in meses if k[0] and k[1])
        if len(orden) >= 2:
            ant, act = orden[-2], orden[-1]
            venta_total = sum(v for v, _ in meses[act].values()) or None
            filas = []
            for nombre, (v, c) in meses[act].items():
                if nombre not in meses[ant]:
                    continue
                v0, c0 = meses[ant][nombre]
                if not v0:
                    continue
                mg0, mg1 = (v0 - c0) / v0, (v - c) / v
                peso = v / venta_total if venta_total else 0
                filas.append({
                    "producto": nombre,
                    "venta": fmt_soles(v),
                    "peso": round(peso * 100, 1),
                    "margen_previo": f"{mg0 * 100:.1f}%",
                    "margen": f"{mg1 * 100:.1f}%",
                    "delta_pp": round((mg1 - mg0) * 100, 2),
                    # Cuánto del margen total explica este producto: su caída
                    # ponderada por lo que pesa en la venta del mes.
                    "aporte_pp": round((mg1 - mg0) * peso * 100, 3),
                })
            if filas:
                anotar_derivado("margen_variable", "por_producto", "margen",
                                "(venta − costo) / venta por producto y mes",
                                "el detalle publica venta y costo por producto "
                                "pero no el margen de cada uno")
                anotar_derivado("margen_variable", "por_producto", "aporte_pp",
                                "(margen_mes − margen_previo) × peso en la venta",
                                "cuántos puntos del margen total explica ese "
                                "producto; ordena por impacto y no por caída "
                                "porcentual, que premiaría a los irrelevantes")
                filas.sort(key=lambda x: x["aporte_pp"])
                res["por_producto"] = filas[:15]
                res["por_producto_meses"] = [f"{int(ant[0])}-{int(ant[1]):02d}",
                                             f"{int(act[0])}-{int(act[1]):02d}"]

    # ── SKU por unidad de negocio: qué producto mueve el margen DENTRO de
    # cada UEN. Mismo criterio de orden que por_producto: puntos de margen
    # sobre el total de su unidad, no caída porcentual.
    sku = found.get("__sku_por_uen") or []
    if sku:
        def _busca(f, suf):
            for k, v in f.items():
                if k.endswith(suf):
                    return v
            return None
        por_uen_mes = {}
        for f in sku:
            nombre = (_busca(f, "[DESCRIPCION]") or "").strip()
            uen = (_busca(f, "[TIPO DE NEGOCIO N1]") or "").strip()
            v = to_float(_busca(f, "[Venta]"))
            c = to_float(_busca(f, "[Costo]"))
            pe = to_float(_busca(f, "[Peso]"))
            anio, mes = to_float(_busca(f, "[Año]")), to_float(_busca(f, "[NroMes]"))
            if not nombre or not uen or v is None or c is None or not v:
                continue
            if not anio or not mes:
                continue
            por_uen_mes.setdefault(uen, {}).setdefault((int(anio), int(mes)), {})[nombre] = (v, c, pe)
        salida = {}
        for uen, meses in por_uen_mes.items():
            orden = sorted(meses)
            if len(orden) < 2:
                continue
            ant, act = orden[-2], orden[-1]
            total = sum(v for v, _, _ in meses[act].values()) or None
            filas = []
            for nombre, (v, c, pe) in meses[act].items():
                if nombre not in meses[ant]:
                    continue
                v0, c0, pe0 = meses[ant][nombre]
                if not v0:
                    continue
                mg0, mg1 = (v0 - c0) / v0, (v - c) / v
                peso = v / total if total else 0
                filas.append({
                    "producto": nombre,
                    "venta": fmt_soles(v),
                    "peso": round(peso * 100, 1),
                    "margen_previo": f"{mg0 * 100:.1f}%",
                    "margen": f"{mg1 * 100:.1f}%",
                    "precio_kg": f"S/{v / pe:.2f}" if pe else None,
                    "precio_kg_previo": f"S/{v0 / pe0:.2f}" if pe0 else None,
                    "delta_pp": round((mg1 - mg0) * 100, 2),
                    "aporte_pp": round((mg1 - mg0) * peso * 100, 3),
                })
            if filas:
                filas.sort(key=lambda x: x["aporte_pp"])
                salida[uen] = filas[:12]
        if salida:
            anotar_derivado("margen_variable", "sku_por_uen", "aporte_pp",
                            "(margen_mes − margen_previo) × peso del SKU en la "
                            "venta de su unidad de negocio",
                            "cuántos puntos del margen de esa UEN explica el SKU; "
                            "ordena por impacto y no por caída porcentual, que "
                            "premiaría a los irrelevantes")
            anotar_derivado("margen_variable", "sku_por_uen", "precio_kg",
                            "venta del SKU / kilos del SKU",
                            "el cruce de SKU con unidad de negocio no existe en "
                            "ningún visual: se arma con los filtros del reporte "
                            "y el precio por kilo se deriva de venta y peso")
            res["sku_por_uen"] = salida
            res["sku_por_uen_meses"] = None
            for uen, meses in por_uen_mes.items():
                orden = sorted(meses)
                if len(orden) >= 2:
                    res["sku_por_uen_meses"] = [
                        f"{orden[-2][0]}-{orden[-2][1]:02d}",
                        f"{orden[-1][0]}-{orden[-1][1]:02d}"]
                    break

    return res, ventas_val, margen_pct

def build_mermas(found):
    ate_val  = found.get("% Merma Ate") or found.get("Merma Ate") or found.get("Merma Planta Ate")
    pach_val = found.get("% Merma Pachacamac") or found.get("Merma Pachacamac") or found.get("Merma Lurin")
    tot_val  = (found.get("% Merma Total") or found.get("% Merma") or
                found.get("% MERMAS") or found.get("Merma %") or
                found.get("Tasa Merma") or found.get("Merma Total"))
    terc_val = found.get("% Merma Terceros") or found.get("Merma Terceros")
    # UEN (Control de Producción: B&D, TIGO, MAQUILA)
    bd_val   = found.get("% Merma B&D") or found.get("Merma B&D") or found.get("% Merma BD")
    tigo_val = found.get("% Merma Tigo") or found.get("Merma Tigo") or found.get("% Merma TIGO")
    maq_val  = found.get("% Merma Maquila") or found.get("Merma Maquila") or found.get("% Merma MAQUILA")
    # Detalle: REAL, STD, DESVIACION (tabla "Detalle de Mermas por Planta")
    real_val  = found.get("REAL")
    std_val   = found.get("STD")
    desv_val  = found.get("DESVIACION")
    item_val  = found.get("ITEM DESVIADO")

    def pct(v):
        n = to_float(v)
        if n is None: return None
        if abs(n) < 1: n *= 100
        return n

    ate, pach, tot, terc = pct(ate_val), pct(pach_val), pct(tot_val), pct(terc_val)
    bd, tigo, maq = pct(bd_val), pct(tigo_val), pct(maq_val)
    worst = max(filter(lambda x: x is not None, [ate, pach, tot, terc]), default=None)
    sem = "green"
    if worst:
        if worst > 3: sem = "red"
        elif worst > 2: sem = "yellow"

    kpis = []
    for label, val, meta in [
        ("Merma Total",       tot,  "2%"),
        ("Merma Ate",         ate,  "2%"),
        ("Merma Pachacamac",  pach, "2%"),
        ("Merma Terceros",    terc, "2%"),
        ("Merma B&D",         bd,   "2%"),
        ("Merma TIGO",        tigo, "2%"),
        ("Merma MAQUILA",     maq,  "2%"),
    ]:
        if val is not None:
            s = "red" if val > 3 else ("yellow" if val > 2 else "green")
            kpis.append({"label": label, "valor": f"{val:.2f}%", "meta": meta, "estado": s})

    if real_val  is not None: kpis.append({"label": "REAL (unidades)", "valor": str(real_val)})
    if std_val   is not None: kpis.append({"label": "STD (estándar)", "valor": str(std_val)})
    if desv_val  is not None: kpis.append({"label": "Desviación", "valor": str(desv_val)})
    if item_val  is not None: kpis.append({"label": "Ítems Desviados", "valor": str(item_val)})

    alerta = f"Merma {worst:.2f}% — sobre meta 2%" if sem != "green" and worst else None
    # Guardar por_uen para el frontend
    por_uen = []
    for uen, val in [("B&D", bd), ("TIGO", tigo), ("MAQUILA", maq)]:
        if val is not None:
            s = "red" if val > 3 else ("yellow" if val > 2 else "green")
            por_uen.append({"uen": uen, "merma": f"{val:.2f}%", "estado": s})
    # Guardar por_planta
    por_planta = []
    for planta, val in [("ATE", ate), ("PACHACAMAC", pach), ("TERCEROS", terc)]:
        if val is not None:
            s = "red" if val > 3 else ("yellow" if val > 2 else "green")
            por_planta.append({"planta": planta, "merma": f"{val:.2f}%", "estado": s})

    result = {"estado": sem, "alerta": alerta, "kpis": kpis}
    if por_uen:    result["por_uen"]    = por_uen
    if por_planta: result["por_planta"] = por_planta

    # Merma por SKU: el nivel más fino que publica el reporte. Se ordena por
    # desviación absoluta contra el estándar, no por porcentaje: un 40% sobre
    # cien kilos importa menos que un 3% sobre cien toneladas.
    sk = []
    for f in (found.get("__por_sku") or []):
        nombre = (f.get("sku") or "").strip()
        desvio = to_float(f.get("desvio"))
        if not nombre or desvio is None:
            continue
        real = to_float(f.get("real"))
        std = to_float(f.get("estandar"))
        pct = to_float(f.get("pct"))
        sk.append({"sku": nombre,
                   "almacen": (f.get("almacen") or "").strip() or None,
                   "estandar": round(abs(std), 1) if std is not None else None,
                   "consumido": round(abs(real), 1) if real is not None else None,
                   "exceso": round(abs(desvio), 1),
                   "pct": (f"{abs(pct) * 100:.1f}%" if pct is not None else None),
                   "_a": abs(desvio)})
    if sk:
        sk.sort(key=lambda x: -x["_a"])
        for x in sk:
            x.pop("_a", None)
        result["por_sku"] = sk[:20]

    return result

def build_consumo_materiales(found):
    """Reporte '14. Consumo Materiales indirectos de produccion' — PAUNO.

    Los cinco KPIs se leen ejecutando la consulta capturada de su tarjeta
    (ver TARJETAS_KPI), así que son literalmente los números que Power BI
    pinta en pantalla.

    Antes se reconstruían a mano y no coincidían. La comparación automática
    contra las tarjetas lo destapó el 13/09/2026: Venta Neta (KG) publicaba
    16,750,700 kg donde la tarjeta marca 10,178,475 — los siete filtros que
    llevaba la reconstrucción no son los de la tarjeta.
    """
    costo_total_val = found.get("Costo total validado") or found.get("Costo Consumo")
    venta_kg_val    = found.get("Peso total KG") or found.get("Venta Neta (KG)") or found.get("Venta Neta KG")
    costo_x_tn_vend_val = found.get("Ratio costo / kg") or found.get("Costo x TN") or found.get("Costo por TN")
    prod_neta_val   = found.get("Producción (KG) Odoo") or found.get("TN PRODUCIDA") or found.get("TN Producida")
    costo_x_tn_prod_val = found.get("Costo x ton producida")

    costo_total     = to_float(costo_total_val)
    venta_kg        = to_float(venta_kg_val)
    costo_x_tn_vend = to_float(costo_x_tn_vend_val)
    prod_neta       = to_float(prod_neta_val)
    costo_x_tn_prod = to_float(costo_x_tn_prod_val)

    # Sin meta ni benchmark definido por negocio todavía — verde de referencia.
    sem = "green"

    kpis = []
    if costo_total is not None:
        kpis.append({"label": "Costo Total", "valor": fmt_soles(abs(costo_total)),
                     "validado": True, "nota": "Filtro exacto confirmado — sin corte de fecha (histórico total)"})
    if venta_kg is not None:
        kpis.append({"label": "Venta Neta (KG)", "valor": f"{venta_kg:,.0f} KG", "validado": True})
    if costo_x_tn_vend is not None:
        kpis.append({"label": "Costo x TN Vendida", "valor": f"S/{costo_x_tn_vend:.2f}", "validado": True})
    if prod_neta is not None:
        kpis.append({"label": "Producción Neta (KG)", "valor": f"{prod_neta:,.0f} KG", "validado": True})
    if costo_x_tn_prod is not None:
        kpis.append({"label": "Costo x TN Producida", "valor": f"S/{costo_x_tn_prod:.2f}", "validado": True})

    if not kpis:
        return None  # No hay datos — no agregar el reporte
    # Marcador de reporte: True mientras quede al menos un KPI sin validar con
    # Copiar consulta. Cada KPI individual lleva su propio campo "validado".
    hay_sin_validar = any(not k.get("validado") for k in kpis)
    return {"estado": sem, "alerta": None, "kpis": kpis, "no_validado_dax": hay_sin_validar}

def build_compras(found):
    ratio_val  = (found.get("Ratio") or found.get("Eficiencia") or found.get("Eficiencia Costo") or
                  found.get("Ratio C/V") or found.get("Ratio Compras") or
                  found.get("Ratio Consumo/Compra") or found.get("% Ratio") or found.get("Consumo/Compra"))
    consumo_val = (found.get("Cant Consumo") or found.get("Consumo") or
                   found.get("Total Consumo") or found.get("Importe Consumo"))
    compra_val  = (found.get("Cant Compra") or found.get("Valor Compras") or
                   found.get("Compras") or found.get("Total Compras") or found.get("Importe Compras"))
    # Medidas del Reporte Compras — ANALISIS DE MATERIALES (PLANIFICACION Y COMPRA)
    stock_pp_val    = found.get("Stock PP")
    stock_val_val   = found.get("Stock Valorizado")
    cons_3m_val     = found.get("Consumo Prom 3M")
    cons_6m_val     = found.get("Consumo Prom 6M")
    faltantes_val   = found.get("Faltantes")
    lead_time_val   = found.get("Lead Time")

    ratio = to_float(ratio_val)
    if ratio and abs(ratio) < 2: ratio *= 100
    if ratio is None and consumo_val and compra_val:
        c, p = to_float(consumo_val), abs(to_float(compra_val) or 0)
        if c and p and p > 0: ratio = (c / p) * 100

    sem = "green"
    if ratio:
        if ratio > 130 or ratio < 70: sem = "red"
        elif ratio > 110 or ratio < 80: sem = "yellow"

    interp = ("jala inventario" if ratio and ratio > 100 else
              "sobre-compra" if ratio and ratio < 80 else "normal")
    kpis = []
    if ratio is not None:
        kpis.append({"label": "Ratio Consumo/Compra", "valor": f"{ratio:.1f}%", "meta": "80-100%",
                     "estado": sem, "interpretacion": interp})
    if consumo_val is not None:
        kpis.append({"label": "Consumo", "valor": fmt_soles(consumo_val)})
    if compra_val is not None:
        kpis.append({"label": "Compras", "valor": fmt_soles(compra_val)})
    # Métricas de planificación de compras
    if stock_pp_val is not None:
        kpis.append({"label": "Stock PP (Punto de Pedido)", "valor": str(int(to_float(stock_pp_val) or 0))})
    if stock_val_val is not None:
        kpis.append({"label": "Stock Valorizado", "valor": fmt_soles(stock_val_val)})
    if cons_3m_val is not None:
        kpis.append({"label": "Consumo Prom 3M", "valor": str(int(to_float(cons_3m_val) or 0))})
    if cons_6m_val is not None:
        kpis.append({"label": "Consumo Prom 6M", "valor": str(int(to_float(cons_6m_val) or 0))})
    if faltantes_val is not None:
        n = to_float(faltantes_val)
        s = "red" if (n or 0) > 0 else "green"
        kpis.append({"label": "Faltantes", "valor": str(int(n or 0)), "estado": s})
    if lead_time_val is not None:
        kpis.append({"label": "Lead Time Prom (días)", "valor": f"{to_float(lead_time_val):.0f}d"})

    # Semáforo: si hay faltantes o sin ratio, usar faltantes
    if not ratio and faltantes_val is not None:
        n = to_float(faltantes_val)
        sem = "red" if (n or 0) > 10 else ("yellow" if (n or 0) > 0 else "green")

    alerta = f"Ratio {ratio:.1f}% — {interp}" if sem != "green" and ratio else (
             f"Faltantes: {int(to_float(faltantes_val) or 0)} ítems" if faltantes_val and to_float(faltantes_val) else None)
    res = {"estado": sem, "alerta": alerta, "kpis": kpis}

    # Totales de la tabla de materiales: sustituyen a los del sondeo, que
    # consultaba las mismas medidas sin los filtros del visual.
    tot_mat = found.get("__totales_materiales") or {}
    for etiqueta, valor in tot_mat.items():
        if valor is None:
            continue
        texto = (f"{valor:,.0f}" if "Consumo" in etiqueta or "PP" in etiqueta
                 else fmt_soles(valor))
        actual = next((k for k in kpis if k["label"] == etiqueta), None)
        if actual is None:
            kpis.append({"label": etiqueta, "valor": texto})
        else:
            actual["valor"] = texto

    # Materiales en quiebre según la explosión de materiales semanal.
    # Es la lista más accionable del reporte: son pocos y hay que comprarlos.
    falt = found.get("__faltantes") or []
    if falt:
        # El faltante puede venir con signo negativo segun como lo calcule el
        # modelo, asi que se compara en valor absoluto y se ordena por magnitud.
        vivos = [f for f in falt if abs(to_float(f.get("faltante")) or 0) > 0]
        vivos.sort(key=lambda f: -abs(to_float(f.get("faltante")) or 0))
        if falt and not vivos:
            DIAGNOSTICO.append({"consulta": "faltantes_vacio", "http": 200,
                                "error": f"{len(falt)} filas, ninguna con "
                                         f"faltante distinto de cero. "
                                         f"Muestra: {falt[:3]}"})
        res["faltantes"] = [{
            "producto": (f.get("producto") or f.get("codigo") or "—"),
            "categoria": f.get("categoria") or "",
            "faltante": f"{abs(to_float(f.get('faltante')) or 0):,.0f}",
            "stock": f"{to_float(f.get('stock')) or 0:,.0f}",
            "lead_time": f.get("lead_time"),
        } for f in vivos[:12]]
        if vivos:
            res["kpis"].append({
                "label": "Materiales en quiebre", "valor": str(len(vivos)),
                "meta": "explosión de materiales", "estado": "red"})

    # Necesidad de compra agrupada por urgencia: 431 líneas no se leen, pero
    # "cuántos ítems hay en cada momento de compra" sí.
    nec = found.get("__necesidad") or []
    if nec:
        por_momento = {}
        for n in nec:
            m = (n.get("momento") or "Sin clasificar").strip()
            por_momento[m] = por_momento.get(m, 0) + 1
        res["necesidad_compra"] = [
            {"momento": m, "items": c}
            for m, c in sorted(por_momento.items(), key=lambda kv: -kv[1])]
        res["kpis"].append({"label": "Ítems con necesidad de compra",
                            "valor": str(len(nec))})
    return res, ratio

def build_inventario(found):
    # ── Clasificación por categoría, de la matriz "CLASIFICACION DE INVENTARIO"
    # (Copiar consulta 2026-09-06). Usa 'SALDO ACTUAL'[Clasificación ALC Meses],
    # que es OTRA columna que la '[Clasificación Segun Consumo]' capturada el
    # 2026-09-04: son dos clasificaciones distintas del mismo saldo, no la
    # misma con otro nombre. Se prefiere la de la matriz porque es la que el
    # reporte muestra hoy, y además trae la apertura por categoría.
    clas = found.get("__clasificacion") or []
    if clas:
        agrupado = {}
        for _cat, cla, soles, _pct in clas:
            if soles is None:
                continue
            # Las clases vienen numeradas ("1. Working", "4. DEAD"): se quita
            # el prefijo de orden para que calcen con los nombres esperados.
            k = re.sub(r"^\s*\d+\.\s*", "", str(cla)).strip().upper()
            agrupado[k] = agrupado.get(k, 0.0) + soles
        found = dict(found)
        found["_composicion"] = list(agrupado.items())

    # ── Composición confirmada con Copiar consulta (2026-09-04): lista de
    # (clasificación, saldo soles) desde 'SALDO ACTUAL'[Clasificación Segun
    # Consumo] — WORKING / EXCESO 1 / EXCESO 2 / DEAD. Tiene prioridad sobre
    # los nombres adivinados de abajo, que quedan solo como fallback.
    composicion = found.get("_composicion") or []
    comp_map = {}
    for clasif, saldo in composicion:
        key = str(clasif or "").strip().upper()
        comp_map[key] = to_float(saldo)

    dead_val    = comp_map.get("DEAD") or found.get("Dead Stock") or found.get("Stock Muerto") or found.get("Inmovilizado") or found.get("Stock Inmovilizado")
    working_val = comp_map.get("WORKING") or found.get("Working Stock") or found.get("Stock Activo") or found.get("Stock Working") or found.get("Stock Normal")
    exceso1_val = comp_map.get("EXCESO 1")
    exceso2_val = comp_map.get("EXCESO 2")
    total_val   = (sum(v for v in comp_map.values() if v) if comp_map else None) or (
                   found.get("Inventario Total") or found.get("Total Inventario") or
                   found.get("Saldo Inventario") or found.get("Valor Inventario") or
                   found.get("Stock Total") or found.get("Total Stock") or found.get("Costo Inventario"))
    cob_total = found.get("Cobertura Total (días)")
    cob_mp    = found.get("Cobertura MP (días)")

    deadpct_val = found.get("% Dead Stock") or found.get("% Inmovilizado") or found.get("% Dead") or found.get("Pct Dead")
    dead_pct = to_float(deadpct_val)
    if dead_pct and abs(dead_pct) < 1: dead_pct *= 100
    if dead_pct is None and dead_val and total_val:
        d, t = to_float(dead_val), to_float(total_val)
        if d and t and t > 0: dead_pct = d / t * 100

    # No usar sem_thresh() aquí: su heurística "abs(n)<1 -> multiplicar x100"
    # duplicaría la escala cuando dead_pct ya es un porcentaje válido y chico
    # (ej. 0.3%), convirtiéndolo por error en 30% y disparando rojo.
    if dead_pct is None:
        sem = "green"
    elif dead_pct >= 10:
        sem = "red"
    elif dead_pct >= 5:
        sem = "yellow"
    else:
        sem = "green"

    kpis = []
    if total_val is not None:
        kpis.append({"label": "Inventario Total", "valor": fmt_soles(total_val)})
    if cob_total is not None:
        kpis.append({"label": "Cobertura Total", "valor": f"{cob_total:.0f} días"})
    if cob_mp is not None:
        kpis.append({"label": "Cobertura MP", "valor": f"{cob_mp:.0f} días"})
    if dead_val is not None:
        kpis.append({"label": "Dead Stock", "valor": fmt_soles(dead_val), "meta": "<S/500K", "estado": sem})
    if dead_pct is not None:
        kpis.append({"label": "% Dead Stock", "valor": f"{dead_pct:.1f}%", "meta": "<5%", "estado": sem})
    if working_val is not None:
        kpis.append({"label": "Working Stock", "valor": fmt_soles(working_val)})
    if exceso1_val is not None:
        kpis.append({"label": "Exceso 1 (2-5 meses)", "valor": fmt_soles(exceso1_val)})
    if exceso2_val is not None:
        kpis.append({"label": "Exceso 2 (5-12 meses)", "valor": fmt_soles(exceso2_val)})

    alerta = f"Dead Stock {dead_pct:.1f}% del inventario" if sem != "green" and dead_pct else None
    res = {"estado": sem, "alerta": alerta, "kpis": kpis}

    # Apertura por categoría. Las etiquetas exactas ("1. Working", "4. Dead")
    # están confirmadas con Copiar consulta de los visuales WORKING y DEAD
    # (2026-09-06); esos gráficos filtran por clasificación y agrupan por
    # categoría, que es la misma porción que produce la consulta agrupada.
    if clas:
        def apertura(marca):
            acum, tot = {}, 0.0
            for cat, cla, soles, _p in clas:
                if soles and marca in str(cla).lower():
                    acum[cat] = acum.get(cat, 0.0) + soles
                    tot += soles
            if not acum:
                return None
            top = sorted(acum.items(), key=lambda kv: -kv[1])[:6]
            for clave in ("dead_por_categoria", "working_por_categoria"):
                anotar_derivado("sop_inventario", clave, "pct",
                                "saldo de la categoría / saldo total × 100", PART)
            return [{"categoria": c, "valor": fmt_soles(v),
                     "pct": round(v / tot * 100, 1) if tot else None}
                    for c, v in top]

        d = apertura("dead")
        if d:
            res["dead_por_categoria"] = d
        w = apertura("working")
        if w:
            res["working_por_categoria"] = w
    return res

def build_control(found):
    """Dashboard de Control Interno ('8. Reporte de auditoría').

    Confirmado con Copiar consulta el 2026-09-04, filtro único Empresa="Pauno",
    sin fecha (acumulado histórico total): Satisfactorio, Con Observaciones,
    Critico, Planes de Acción (Total/Abiertos/Cerrados/Atrasados), Puntos
    Ejecutados y — desde la tarjeta combinada "Resultado Acumulado" —
    % Cumplimento, Calificación y Puntos Totales (el % Cumplimiento real
    del modelo, no un cálculo derivado).
    """
    satisf_val = found.get("Satisfactorio")
    obs_val    = found.get("Con Observaciones")
    crit_val   = found.get("Critico")
    abiertos_val = found.get("Planes Abiertos")
    cerrados_val = found.get("Planes Cerrados")
    ejecutado_val = found.get("Ejecutado")
    total_planes_val = found.get("Total Planes de Acción")
    atrasados_val = found.get("Planes Atrasados")
    cumpl_val = found.get("% Cumplimento")
    calif_val = found.get("Calificación")
    puntos_totales_val = found.get("Puntos Totales")

    satisf = to_float(satisf_val)
    obs    = to_float(obs_val)
    crit   = to_float(crit_val)
    abiertos = to_float(abiertos_val)
    cerrados = to_float(cerrados_val)
    ejecutado = to_float(ejecutado_val)
    total  = sum(v for v in (satisf, obs, crit) if v is not None) or None
    # Total real del modelo (medida propia) en vez de sumar Abiertos+Cerrados a mano.
    total_planes = to_float(total_planes_val) or ((abiertos or 0) + (cerrados or 0) if (abiertos is not None or cerrados is not None) else None)
    atrasados = to_float(atrasados_val)
    avance_planes_pct = (cerrados / total_planes * 100) if (cerrados is not None and total_planes) else None
    cumpl_pct = to_float(cumpl_val)
    if cumpl_pct is not None and abs(cumpl_pct) < 1:
        cumpl_pct *= 100
    calificacion = to_float(calif_val)
    puntos_totales = to_float(puntos_totales_val)

    crit_pct = (crit / total * 100) if (crit is not None and total) else None
    sem = "green"
    if crit_pct is not None:
        sem = "red" if crit_pct > 20 else ("yellow" if crit_pct > 10 else "green")

    kpis = []
    if total is not None:
        kpis.append({"label": "Puntos de Control evaluados", "valor": str(int(total))})
    if satisf is not None:
        kpis.append({"label": "Satisfactorio", "valor": str(int(satisf)), "estado": "green"})
    if obs is not None:
        kpis.append({"label": "Con Observaciones", "valor": str(int(obs)), "estado": "yellow"})
    if crit is not None:
        kpis.append({"label": "Crítico", "valor": str(int(crit)), "meta": "0", "estado": "red" if crit > 0 else "green"})
    if crit_pct is not None:
        kpis.append({"label": "% Puntos en estado Crítico", "valor": f"{crit_pct:.1f}%", "meta": "<10%", "estado": sem})
    if total_planes is not None:
        kpis.append({"label": "Total Planes de Acción", "valor": str(int(total_planes))})
    if abiertos is not None:
        kpis.append({"label": "Planes de Acción Abiertos", "valor": str(int(abiertos)), "estado": "yellow" if abiertos > 0 else "green"})
    if cerrados is not None:
        kpis.append({"label": "Planes de Acción Cerrados", "valor": str(int(cerrados)), "estado": "green"})
    if atrasados is not None:
        kpis.append({"label": "Planes de Acción Atrasados", "valor": str(int(atrasados)), "meta": "0", "estado": "red" if atrasados > 0 else "green"})
    if avance_planes_pct is not None:
        kpis.append({"label": "% Avance Planes de Acción", "valor": f"{avance_planes_pct:.1f}%", "estado": "green" if avance_planes_pct>=80 else ("yellow" if avance_planes_pct>=50 else "red")})
    if ejecutado is not None:
        kpis.append({"label": "Puntos Ejecutados", "valor": str(int(ejecutado))})
    if calificacion is not None:
        kpis.append({"label": "Calificación", "valor": str(int(calificacion))})
    if puntos_totales is not None:
        kpis.append({"label": "Puntos Totales", "valor": str(int(puntos_totales))})
    cumpl_sem = None
    if cumpl_pct is not None:
        cumpl_sem = "green" if cumpl_pct >= 85 else ("yellow" if cumpl_pct >= 70 else "red")
        kpis.insert(0, {"label": "% Cumplimiento", "valor": f"{cumpl_pct:.2f}%", "meta": "85%", "estado": cumpl_sem})

    estado_final = cumpl_sem or sem
    alerta = None
    if cumpl_sem and cumpl_sem != "green":
        alerta = f"% Cumplimiento {cumpl_pct:.1f}% — bajo meta 85%"
    elif sem != "green" and crit:
        alerta = f"{int(crit)} puntos de control en estado Crítico ({crit_pct:.1f}% del total)"
    return {"estado": estado_final if kpis else "green", "alerta": alerta, "kpis": kpis}

def build_productividad(found):
    """Reporte '13. Productividad por Funcionario' — PAUNO.

    Nombres de medida confirmados el 2026-09-03 leyendo el panel Datos > Measures
    en Power BI Service (Editar informe) y validados con EVALUATE ROW(...) en la
    Vista de consultas DAX (todas resuelven sin error):
      SUM PROD          -> Producción Total (KG)
      GASTO TOTAL        -> Planilla Total (S/.)
      PROD OPERATIVA      -> Planilla (S/.) entre KG Producido
      VENTA KG X SOL       -> Planilla (S/.) entre KG Vendido

    "Venta Neta (KG)" aparece en el panel como "Suma de Peso total KG" porque es la
    COLUMNA 'Maestra de Facturacion (Total)'[Peso total K] con SUM directo, no una
    medida DAX — confirmado el 2026-09-03 con Analizador de rendimiento > Copiar
    consulta (ver dax_productividad_venta_neta_kg). Trae 2 filtros propios que las
    otras 4 tarjetas no tienen: CATEGORIZACION="VENTA BRUTA" y TIPO DE NEGOCIO N2
    no vacío.
    """
    # ── BARRERA DE SEGURIDAD — historial ────────────────────────────────────
    # GASTO TOTAL se probó en vivo (2026-09-03) sin el filtro real del reporte:
    #   sin filtro     = -247
    #   con solo Calendario[Año]=2026 = -117
    #   tarjeta real   = S/4,484,364
    # Confirmado con "Copiar consulta" (Analizador de rendimiento) que el reporte
    # aplica 9 filtros: Calendario[Año]=2026, Calendario[MesActual]="Otros",
    # Exl Cuenta Contables[CLASIFICACION] IN {"AMBAS","GASTO DE PERSONAL OPERATIVO"},
    # [GASTO_PERSONAL]="GASTO DE PERSONAL", [SUB_CATEGORIA]="Gasto de Personal Operativo",
    # exclusión de categoria_producto/producto/estado, y Calendario[Date] > 2025-07-31.
    # main() ahora llama dax_productividad_pauno() con esos 9 filtros exactos y
    # sobrescribe found[...] con el valor correcto — por eso ya es seguro leer
    # found.get(...) directo aquí. GASTO_TOTAL_CONFIRMADO queda como bandera
    # histórica del hallazgo, ya no bloquea nada.
    GASTO_TOTAL_CONFIRMADO = True

    plan_kg_vend = plan_kg_prod = plan_total = None
    if GASTO_TOTAL_CONFIRMADO:
        plan_kg_vend  = (found.get("VENTA KG X SOL") or
                         found.get("Planilla (S/.) entre KG Vendido") or
                         found.get("Planilla Entre KG Vendido") or found.get("Planilla/KG Vendido"))
        plan_kg_prod  = (found.get("PROD OPERATIVA") or
                         found.get("Planilla (S/.) entre KG Producido") or
                         found.get("Planilla Entre KG Producido") or found.get("Planilla/KG Producido"))
        plan_total    = (found.get("GASTO TOTAL") or
                         found.get("Planilla Total (S/.)") or found.get("Planilla Total"))

    prod_total    = (found.get("SUM PROD") or
                     found.get("Produccion Total (KG)") or found.get("Produccion Total KG") or
                     found.get("Produccion Total") or found.get("Kg Producidos") or found.get("Kg Producción"))
    # "Venta Neta (KG)" — columna SUM confirmada y filtrada por
    # dax_productividad_venta_neta_kg(), guardada bajo esta clave interna.
    venta_neta_kg = found.get("__VENTA_NETA_KG__")

    kpis = []
    def sf2(v, prefix="S/"):
        n = to_float(v)
        return f"{prefix}{n:.2f}" if n else None
    def kg_fmt(v):
        n = to_float(v)
        return f"{n:,.0f} kg" if n else None

    if plan_kg_prod is not None:
        kpis.append({"label": "Planilla S/. / KG Producido", "valor": sf2(plan_kg_prod) or "—"})
    if plan_kg_vend is not None:
        kpis.append({"label": "Planilla S/. / KG Vendido",   "valor": sf2(plan_kg_vend) or "—"})
    if prod_total is not None:
        kpis.append({"label": "Producción Total (KG)",        "valor": kg_fmt(prod_total) or "—"})
    if venta_neta_kg is not None:
        kpis.append({"label": "Venta Neta (KG)",              "valor": kg_fmt(venta_neta_kg) or "—"})
    if plan_total is not None:
        kpis.append({"label": "Planilla Total",               "valor": fmt_soles(plan_total)})

    # Tendencia: menor S//kg = mayor eficiencia
    n_prod = to_float(plan_kg_prod)
    sem = "green"
    if n_prod:
        if n_prod > 0.60: sem = "red"
        elif n_prod > 0.50: sem = "yellow"
    alerta = f"Planilla/KG Producido S/{n_prod:.2f} — revisar eficiencia" if sem != "green" and n_prod else None
    return {"estado": sem, "alerta": alerta, "kpis": kpis}

def dax_fillrate_por_marca(token, ws, dataset_id, label="fillrate_marca"):
    """Pedidos no atendidos del MES ACTUAL, por marca.

    Confirmado con Copiar consulta (2026-09-06) sobre el visual
    "SOLES NO ATENDIDOS POR GRUPO" del reporte FILLRATE. No es una serie:
    el filtro 'CALENDARIO'[FIL_MES_ACTUAL] = "MES_ACTUAL" lo acota al mes en
    curso, así que es un desglose, no una tendencia.

    Ese reporte vive en un workspace distinto al de los demás, por eso `ws`
    va como parámetro en vez de usar el global.
    """
    q = ("DEFINE\n"
         "\tVAR __DS0FilterTable = \n"
         "\t\tTREATAS({\"MES_ACTUAL\"}, 'CALENDARIO'[FIL_MES_ACTUAL])\n\n"
         "\tVAR __DS0Core = \n"
         "\t\tSUMMARIZECOLUMNS(\n"
         "\t\t\t'PEDIDOS'[MARCA],\n"
         "\t\t\t__DS0FilterTable,\n"
         "\t\t\t\"PEDIDOS_NO_ATENDIDOS\", '0_MEDIDAS'[PEDIDOS NO ATENDIDOS]\n"
         "\t\t)\n\n"
         "EVALUATE\n\t__DS0Core\n\n"
         "ORDER BY\n\t[PEDIDOS_NO_ATENDIDOS] DESC, 'PEDIDOS'[MARCA]")
    rows = dax(token, ws, dataset_id, q, label)
    out = []
    for r in rows or []:
        marca = r.get("PEDIDOS[MARCA]") or r.get("[MARCA]")
        val = to_float(r.get("[PEDIDOS_NO_ATENDIDOS]") or r.get("PEDIDOS_NO_ATENDIDOS"))
        if marca and val is not None:
            out.append((str(marca), val))
    out.sort(key=lambda t: -t[1])
    return out


def dax_fillrate_no_atendido(token, ws, dataset_id, label="fillrate_total"):
    """Total de pedidos no atendidos del mes actual (la tarjeta del reporte).

    Confirmado con Copiar consulta (2026-09-06). Mismo filtro de mes que el
    desglose por marca, pero sin agrupar. Se usa este valor como KPI en vez
    de sumar las marcas: es el número que muestra la tarjeta.
    """
    q = ("DEFINE VAR __DS0FilterTable = \n"
         "\tTREATAS({\"MES_ACTUAL\"}, 'CALENDARIO'[FIL_MES_ACTUAL])\n\n"
         "EVALUATE\n"
         "\tSUMMARIZECOLUMNS(\n"
         "\t\t__DS0FilterTable,\n"
         "\t\t\"PEDIDOS_NO_ATENDIDOS\", IGNORE('0_MEDIDAS'[PEDIDOS NO ATENDIDOS])\n"
         "\t)")
    rows = dax(token, ws, dataset_id, q, label)
    if rows:
        return to_float(rows[0].get("[PEDIDOS_NO_ATENDIDOS]") or rows[0].get("PEDIDOS_NO_ATENDIDOS"))
    return None


def dax_fillrate_por_grupo(token, ws, dataset_id, label="fillrate_grupo"):
    """% Fill Rate del mes actual por cliente/grupo.

    Confirmado con Copiar consulta (2026-09-06) sobre el gráfico de barras
    de la izquierda del reporte FILLRATE. El ORDER BY del original es
    ascendente por [FILLRATE]: el reporte muestra primero al peor atendido,
    y aquí se respeta ese orden porque es el accionable.
    """
    q = ("DEFINE\n"
         "\tVAR __DS0FilterTable = \n"
         "\t\tTREATAS({\"MES_ACTUAL\"}, 'CALENDARIO'[FIL_MES_ACTUAL])\n\n"
         "\tVAR __DS0Core = \n"
         "\t\tSUMMARIZECOLUMNS('PEDIDOS'[GRUPO], __DS0FilterTable, "
         "\"FILLRATE\", '0_MEDIDAS'[FILLRATE])\n\n"
         "EVALUATE\n\t__DS0Core\n\n"
         "ORDER BY\n\t[FILLRATE], 'PEDIDOS'[GRUPO]")
    rows = dax(token, ws, dataset_id, q, label)
    out = []
    for r in rows or []:
        g = r.get("PEDIDOS[GRUPO]") or r.get("[GRUPO]")
        v = to_float(r.get("[FILLRATE]") or r.get("FILLRATE"))
        if g and v is not None:
            out.append((str(g), v))
    out.sort(key=lambda t: t[1])          # peor atendido primero
    return out


def dax_fillrate_soles_por_grupo(token, ws, dataset_id, label="fillrate_soles_grupo"):
    """Soles no atendidos del mes actual por cliente/grupo.

    Confirmado con Copiar consulta (2026-09-06). Es el gráfico gemelo del de
    porcentaje: mismo filtro de mes y misma dimensión, pero mide el monto.
    El original ordena DESC, de mayor monto a menor.

    Se cruza con el % para poder ver ambas cosas juntas: un cliente al 92%
    puede pesar más en soles que uno al 80%, y solo con el porcentaje esa
    prioridad no se ve.
    """
    q = ("DEFINE\n"
         "\tVAR __DS0FilterTable = \n"
         "\t\tTREATAS({\"MES_ACTUAL\"}, 'CALENDARIO'[FIL_MES_ACTUAL])\n\n"
         "\tVAR __DS0Core = \n"
         "\t\tSUMMARIZECOLUMNS(\n"
         "\t\t\t'PEDIDOS'[GRUPO],\n"
         "\t\t\t__DS0FilterTable,\n"
         "\t\t\t\"PEDIDOS_NO_ATENDIDOS\", '0_MEDIDAS'[PEDIDOS NO ATENDIDOS]\n"
         "\t\t)\n\n"
         "EVALUATE\n\t__DS0Core\n\n"
         "ORDER BY\n\t[PEDIDOS_NO_ATENDIDOS] DESC, 'PEDIDOS'[GRUPO]")
    rows = dax(token, ws, dataset_id, q, label)
    out = {}
    for r in rows or []:
        g = r.get("PEDIDOS[GRUPO]") or r.get("[GRUPO]")
        v = to_float(r.get("[PEDIDOS_NO_ATENDIDOS]") or r.get("PEDIDOS_NO_ATENDIDOS"))
        if g and v is not None:
            out[str(g)] = v
    return out


def dax_fillrate_medida_mes(token, ws, dataset_id, medida, alias, label="fillrate_card"):
    """Medida del mes actual del reporte FILLRATE (las tarjetas azules).

    Patrón confirmado con Copiar consulta (2026-09-06) sobre las tarjetas:
    filtro único 'CALENDARIO'[FIL_MES_ACTUAL] = "MES_ACTUAL", sin agrupar.
    `medida` ej. "VENTA", "FILLRATE", "PEDIDOS NO ATENDIDOS".
    """
    q = ("DEFINE VAR __DS0FilterTable = \n"
         "\tTREATAS({\"MES_ACTUAL\"}, 'CALENDARIO'[FIL_MES_ACTUAL])\n\n"
         "EVALUATE\n"
         f"\tSUMMARIZECOLUMNS(__DS0FilterTable, \"{alias}\", IGNORE('0_MEDIDAS'[{medida}]))")
    rows = dax(token, ws, dataset_id, q, label)
    if rows:
        return to_float(rows[0].get(f"[{alias}]") or rows[0].get(alias))
    return None


def dax_sop_clasificacion(token, ws, dataset_id, label="sop_clasificacion"):
    """Clasificación del inventario (Working / Exceso / Dead) por categoría.

    Confirmado con Copiar consulta (2026-09-06) sobre la matriz
    "CLASIFICACION DE INVENTARIO" del reporte '11. Reporte de Planificaciones'.

    Del original se conservan el filtro y la medida DM0_Sort TEXTUALES. Lo que
    se quita es la maquinaria de subtotales (ROLLUPADDISSUBTOTAL,
    NATURALLEFTOUTERJOIN, SUBSTITUTEWITHINDEX): sirve para pintar las filas de
    Total del visual, no cambia las cifras de cada celda. Los totales se
    recomponen sumando, que es lo que hace la matriz.

    Se pide además el saldo absoluto en soles, que el visual no muestra pero
    sale de la misma columna — así el % queda comprobable contra su origen.
    """
    q = (
        "DEFINE\n"
        "\tVAR __DS0FilterTable = \n"
        "\t\tFILTER(\n"
        "\t\t\tKEEPFILTERS(VALUES('Maestra de Productos'[data.categoria_producto])),\n"
        "\t\t\tNOT('Maestra de Productos'[data.categoria_producto] IN {BLANK()})\n"
        "\t\t)\n\n"
        "EVALUATE\n"
        "\tSUMMARIZECOLUMNS(\n"
        "\t\t'Maestra de Productos'[data.categoria_producto],\n"
        "\t\t'SALDO ACTUAL'[Clasificación ALC Meses],\n"
        "\t\t__DS0FilterTable,\n"
        "\t\t\"Saldo_Soles\", CALCULATE(SUM('SALDO ACTUAL'[Saldo Soles])),\n"
        "\t\t\"DM0_Sort\", CALCULATE(\n"
        "\t\t\tDIVIDE(SUM('SALDO ACTUAL'[Saldo Soles]), CALCULATE(\n"
        "\t\t\t\tSUM('SALDO ACTUAL'[Saldo Soles]),\n"
        "\t\t\t\tALLSELECTED('Maestra de Productos'[data.categoria_producto]),\n"
        "\t\t\t\tALLSELECTED('SALDO ACTUAL'[Clasificación ALC Meses])\n"
        "\t\t\t))\n"
        "\t\t)\n"
        "\t)"
    )
    rows = dax(token, ws, dataset_id, q, label)
    out = []
    for r in rows or []:
        cat = (r.get("Maestra de Productos[data.categoria_producto]")
               or r.get("[data.categoria_producto]"))
        cla = (r.get("SALDO ACTUAL[Clasificación ALC Meses]")
               or r.get("[Clasificación ALC Meses]"))
        soles = to_float(r.get("[Saldo_Soles]") or r.get("Saldo_Soles"))
        pct = to_float(r.get("[DM0_Sort]") or r.get("DM0_Sort"))
        if cat and cla:
            out.append((str(cat), str(cla), soles, pct))
    return out


def build_fill_rate(found):
    # '% Fill Rate' SÍ responde al filtro de mes; '% FILLRATE' devuelve el acumulado
    # histórico igual en todos los meses. Se prefiere la que refleja el mes en curso.
    # La tarjeta del reporte manda. La sonda genérica traía '% Fill Rate'
    # sin los filtros del visual y daba 83.5% donde el reporte muestra 88%.
    # El resto queda solo como respaldo si la tarjeta no responde.
    fill_val = found.get("__fillrate_card")
    if fill_val is None:
        fill_val = (found.get("% Fill Rate") or found.get("% FILLRATE") or
                    found.get("% FillRate") or found.get("Fill Rate") or
                    found.get("Tasa Atención") or found.get("Tasa Atencion"))
    fill_pct = to_float(fill_val)
    if fill_pct and abs(fill_pct) < 1: fill_pct *= 100
    sem = sem_thresh(fill_pct, red_below=85, yellow_below=95)

    ov_val     = found.get("ORDEN DE VENTA") or found.get("Orden de Venta")
    fact_val   = found.get("FACTURACION") or found.get("Facturacion")
    vp_val     = found.get("VENTA PERDIDA") or found.get("Venta Perdida")
    ov         = to_float(ov_val)
    fact       = to_float(fact_val)
    vp         = to_float(vp_val)

    kpis = []
    if fill_pct is not None:
        kpis.append({"label": "Fill Rate", "valor": f"{fill_pct:.1f}%", "meta": "98%", "estado": sem})
    if ov:
        kpis.append({"label": "Orden de Venta", "valor": fmt_soles(ov)})
    if fact:
        kpis.append({"label": "Facturación", "valor": fmt_soles(fact)})
    if vp:
        kpis.append({"label": "Venta Perdida", "valor": fmt_soles(vp), "estado": "red" if vp > 0 else "green"})

    alerta = f"Fill Rate {fmt_pct(fill_pct)} — bajo meta 98%" if sem != "green" and fill_pct else None
    res = {"estado": sem, "alerta": alerta, "kpis": kpis}

    # Desglose por marca del mes actual (ver dax_fillrate_por_marca)
    # Venta del mes según la tarjeta del reporte. Se muestra SIN símbolo de
    # moneda: la tarjeta se rotula "VENTA MILES" y todavía no está confirmado
    # si el número está expresado en miles. Ya nos pasó con
    # 'PEDIDOS NO ATENDIDOS', que resultó ser un conteo y no soles, así que
    # aquí no se asume la unidad hasta tenerla confirmada.
    vm = found.get("__venta_mes")
    if vm is not None:
        res["kpis"].append({"label": "Venta del mes (tarjeta)", "valor": f"{vm:,.0f}"})

    # Clientes peor atendidos del mes (ver dax_fillrate_por_grupo)
    grupos = found.get("__por_grupo") or []
    if grupos:
        soles = found.get("__soles_por_grupo") or {}
        res_g = []
        for g, v in grupos:
            pct = v * 100 if abs(v) <= 1 else v
            item = {"grupo": g, "valor": f"{pct:.1f}%",
                    "estado": "red" if pct < 85 else "yellow" if pct < 95 else "green"}
            if g in soles:
                # CONTEO de pedidos, no dinero. El visual se titula "SOLES NO
                # ATENDIDOS POR GRUPO" pero la medida es 'PEDIDOS NO ATENDIDOS'
                # y devuelve una cantidad de pedidos (confirmado por el usuario,
                # 2026-09-06). Formatearlo con fmt_soles ponía un "S/" a un
                # conteo: la app decía "S/129" donde son 129 pedidos.
                item["pedidos"] = f"{soles[g]:,.0f}"
                item["pedidos_num"] = soles[g]
            res_g.append(item)
        res["por_grupo"] = res_g

    marcas = found.get("__por_marca") or []
    if marcas:
        suma = sum(v for _, v in marcas)
        # El total viene de la tarjeta del reporte; la suma de marcas solo se
        # usa si esa consulta falló. Si ambas existen y difieren, se avisa:
        # significaría que el desglose no cubre toda la venta no atendida.
        total = found.get("__no_atendido_total")
        if total is None:
            total = suma or None
        elif suma and abs(total - suma) > max(1.0, abs(total) * 0.005):
            print(f"    ⚠ Fill Rate: tarjeta {total:,.0f} vs suma de marcas "
                  f"{suma:,.0f} — el desglose no cuadra con el total")
        anotar_derivado("fill_rate", "por_marca", "pct",
                        "pedidos de la marca / pedidos totales × 100", PART)
        res["por_marca"] = [{
            "marca": m,
            "valor": f"{v:,.0f}",          # pedidos, no soles
            "pct": round(v / total * 100, 1) if total else None,
        } for m, v in marcas]
        if total:
            res["kpis"].append({"label": "Pedidos no atendidos (mes)",
                                "valor": f"{total:,.0f}",
                                "estado": "red" if total > 0 else "green"})
    return res

def dax_avance_por_canal(token, ws, dataset_id, mes, label="avance_canal"):
    """Avance de facturación vs presupuesto, por canal comercial.

    Confirmado con Copiar consulta (2026-09-06) sobre la tabla
    "FACTURACION AL 05 SETIEMBRE" del reporte '11. Reporte de Planificaciones',
    pestaña RESUMEN 2.

    Seis medidas de cinco tablas distintas:
      · [Monto_Neto_Factura]        facturado
      · [CUOTA DIRECTORIO]          presupuesto del mes
      · [PPTO Acumulado Hasta Ayer] presupuesto proporcional al día
      · [%Av vs PPTO AL DIA]        avance contra ese presupuesto al día
      · [Monto Neto Pendiente]      pedidos aún sin facturar
      · [VENTA TOTAL]               venta total

    El único filtro de tiempo es 'Calendario'[Mes Nº] — NO hay filtro de año,
    tal como lo genera el visual. El mes va como parámetro porque en el
    reporte lo fija un segmentador.

    Ojo con el título del visual: dice "AL 05 SETIEMBRE" pero la consulta
    filtra el mes 7. El título es un cuadro de texto fijo, no sigue al
    segmentador.

    Sobre el avance por encima del 200%: viene de que el visual agrupa por mes
    sin separar el año, así que suma el mismo mes de varios años contra un
    presupuesto mensual. Se intentó corregir añadiendo 'Calendario'[Año] como
    columna de agrupación (commit 426975b) y el resultado fue peor: el
    presupuesto y lo facturado salieron en blanco para todos los canales.
    Lo más probable es que 'Exl PPTO' no se relacione con esa columna del
    calendario, así que agruparla rompe el vínculo.

    Queda como está —fiel a lo que muestra el reporte— con la advertencia en
    el KPI. Arreglarlo de verdad es editar el visual en Power BI y agregarle
    el filtro de año.

    Se quita solo el ROLLUPADDISSUBTOTAL (la fila Total del visual, que se
    recompone sumando) y el TOPN de presentación.
    """
    q = (
        "DEFINE\n"
        "\tVAR __DS0FilterTable = \n"
        "\t\tFILTER(\n"
        "\t\t\tKEEPFILTERS(VALUES('Exl Cliente x Vendedor'[Canal])),\n"
        "\t\t\tNOT('Exl Cliente x Vendedor'[Canal] IN {BLANK()})\n"
        "\t\t)\n\n"
        "\tVAR __DS0FilterTable2 = \n"
        f"\t\tTREATAS({{{mes}}}, 'Calendario'[Mes Nº])\n\n"
        "EVALUATE\n"
        "\tSUMMARIZECOLUMNS(\n"
        "\t\t'Exl Cliente x Vendedor'[Canal],\n"
        "\t\t__DS0FilterTable,\n"
        "\t\t__DS0FilterTable2,\n"
        "\t\t\"SumMonto_Neto_Factura\", CALCULATE(SUM('Maestra de Facturacion (Total)'[Monto_Neto_Factura])),\n"
        "\t\t\"SumCUOTA_DIRECTORIO\", CALCULATE(SUM('Exl PPTO'[CUOTA DIRECTORIO])),\n"
        "\t\t\"PPTO_Acumulado_Hasta_Ayer\", 'Exl PPTO Semanal'[PPTO Acumulado Hasta Ayer],\n"
        "\t\t\"v_Av_vs_PPTO_AL_DIA\", 'Maestra de Facturacion (Total)'[%Av vs PPTO AL DIA],\n"
        "\t\t\"SumMonto_Neto_Pendiente\", CALCULATE(SUM('Exl A Maestra de Ordenes de Venta'[Monto Neto Pendiente])),\n"
        "\t\t\"VENTA_TOTAL\", 'Maestra de Facturacion (Total)'[VENTA TOTAL]\n"
        "\t)\n\n"
        "ORDER BY\n\t[SumMonto_Neto_Factura] DESC"
    )
    rows = dax(token, ws, dataset_id, q, label)
    out = []
    for r in rows or []:
        canal = (r.get("Exl Cliente x Vendedor[Canal]") or r.get("[Canal]"))
        if not canal:
            continue
        out.append({
            "canal": str(canal),
            "facturado": to_float(r.get("[SumMonto_Neto_Factura]")),
            "ppto": to_float(r.get("[SumCUOTA_DIRECTORIO]")),
            "ppto_al_dia": to_float(r.get("[PPTO_Acumulado_Hasta_Ayer]")),
            "avance_al_dia": to_float(r.get("[v_Av_vs_PPTO_AL_DIA]")),
            "pendiente": to_float(r.get("[SumMonto_Neto_Pendiente]")),
            "venta_total": to_float(r.get("[VENTA_TOTAL]")),
        })
    return out


def dax_cumplimiento_produccion(token, ws, dataset_id, label="cumpl_produccion"):
    """Cumplimiento del programa de producción del DÍA ANTERIOR, por categoría.

    Confirmado con Copiar consulta (2026-09-06) sobre la tabla "CUMPLIMIENTO
    DE PROGRAMA DE PRODUCCION DIARIO" del reporte '11. Reporte de
    Planificaciones', pestaña RESUMEN 2.

    El filtro de tiempo es 'Calendario'[Es ayer] = "Sí": la tabla mide UN día,
    no el mes acumulado. El 70% que muestra el reporte es el cumplimiento de
    ayer, y cambia cada día. Se etiqueta como tal para que nadie lo lea como
    un indicador mensual.

    Solo entran cuatro categorías de producto terminado, que son las que
    tienen programa de fabricación.

    Se quita el ROLLUPADDISSUBTOTAL (la fila Total, que se recompone sumando)
    y el TOPN de presentación.
    """
    q = (
        "DEFINE\n"
        "\tVAR __DS0FilterTable = \n"
        "\t\tTREATAS(\n"
        "\t\t\t{\"MERCADERIAS SALSAS PACKS\",\n"
        "\t\t\t\t\"PT DERIVADOS LACTEOS\",\n"
        "\t\t\t\t\"PT SALSAS\",\n"
        "\t\t\t\t\"PT YOGURTS\"},\n"
        "\t\t\t'Maestra de Productos'[data.categoria_producto]\n"
        "\t\t)\n\n"
        "\tVAR __DS0FilterTable2 = \n"
        "\t\tTREATAS({\"Sí\"}, 'Calendario'[Es ayer])\n\n"
        "EVALUATE\n"
        "\tSUMMARIZECOLUMNS(\n"
        "\t\t'Maestra de Productos'[data.categoria_producto],\n"
        "\t\t__DS0FilterTable,\n"
        "\t\t__DS0FilterTable2,\n"
        "\t\t\"SumPeso_Producido\", CALCULATE(SUM('Orden de Fabricacion'[Peso Producido])),\n"
        "\t\t\"SumPeso_a_Producir\", CALCULATE(SUM('Orden de Fabricacion'[Peso a Producir])),\n"
        "\t\t\"CUMPLIMIENTO_GENERAL_PESO\", 'Orden de Fabricacion'[CUMPLIMIENTO GENERAL PESO]\n"
        "\t)\n\n"
        "ORDER BY\n\t'Maestra de Productos'[data.categoria_producto]"
    )
    rows = dax(token, ws, dataset_id, q, label)
    out = []
    for r in rows or []:
        cat = (r.get("Maestra de Productos[data.categoria_producto]")
               or r.get("[data.categoria_producto]"))
        if not cat:
            continue
        out.append({
            "categoria": str(cat),
            "producido": to_float(r.get("[SumPeso_Producido]")),
            "programado": to_float(r.get("[SumPeso_a_Producir]")),
            "cumplimiento": to_float(r.get("[CUMPLIMIENTO_GENERAL_PESO]")),
        })
    return out


def build_avance(found):
    avance_val = (found.get("Avance") or found.get("% Avance") or found.get("Avance PPTO") or
                  found.get("% Avance Presupuesto"))
    real_val   = found.get("Ventas PPTO") or found.get("Presupuesto")
    ppto_val   = found.get("Presupuesto") or found.get("PPTO")

    avance_pct = to_float(avance_val)
    if avance_pct and abs(avance_pct) < 1: avance_pct *= 100
    sem = sem_thresh(avance_pct, red_below=60, yellow_below=80)

    kpis = []
    if avance_pct is not None:
        kpis.append({"label": "Avance vs Ppto", "valor": f"{avance_pct:.1f}%", "meta": "100%", "estado": sem})
    if real_val is not None:
        kpis.append({"label": "Ventas reales", "valor": fmt_soles(real_val)})

    alerta = f"Avance {fmt_pct(avance_pct)} vs presupuesto" if sem != "green" and avance_pct else None
    res = {"estado": sem, "alerta": alerta, "kpis": kpis}

    # Cumplimiento del programa de producción de AYER (ver
    # dax_cumplimiento_produccion). Se etiqueta con el día explícito porque
    # mide una sola jornada, no el mes.
    cump = found.get("__cumpl_produccion") or []
    if cump:
        prog = sum(c["programado"] or 0 for c in cump)
        prod = sum(c["producido"] or 0 for c in cump)
        res["cumplimiento_produccion"] = [{
            "categoria": c["categoria"],
            "programado": f"{c['programado']:,.0f}" if c["programado"] is not None else "—",
            "producido": f"{c['producido']:,.0f}" if c["producido"] is not None else "—",
            "pct": (round(c["producido"] / c["programado"] * 100, 1)
                    if c["programado"] and c["producido"] is not None else None),
        } for c in cump]
        anotar_derivado("margen_variable_pag2", "cumplimiento_produccion", "pct",
                        "producido / programado × 100",
                        "el reporte muestra las dos cantidades pero no el "
                        "porcentaje por categoría")
        if prog:
            p = prod / prog * 100
            res["kpis"].append({
                "label": "Cumplimiento producción (ayer)",
                "valor": f"{p:.0f}%", "meta": "100% del programa del día",
                "estado": "green" if p >= 95 else "yellow" if p >= 80 else "red"})

    # Avance por canal (ver dax_avance_por_canal). Sustituye a las medidas del
    # sondeo genérico: estas vienen de la tabla del reporte, con su filtro.
    canales = found.get("__avance_canal") or []
    if canales:
        anotar_derivado("margen_variable_pag2", "por_canal", "avance",
                        "facturado / ppto × 100",
                        "la medida del reporte [%Av vs PPTO AL DIA] compara "
                        "contra el presupuesto proporcional al día y devuelve 0")
        ppto = sum(c["ppto"] or 0 for c in canales)
        fact = sum(c["facturado"] or 0 for c in canales)
        res["por_canal"] = [{
            "canal": c["canal"],
            "ppto": fmt_soles(c["ppto"]) if c["ppto"] is not None else "—",
            "facturado": fmt_soles(c["facturado"]) if c["facturado"] is not None else "—",
            "pendiente": fmt_soles(c["pendiente"]) if c["pendiente"] is not None else "—",
            # Avance = facturado / presupuesto del mes. Es división directa de
            # dos cifras del reporte, no una medida propia: la medida
            # [%Av vs PPTO AL DIA] compara contra el presupuesto proporcional
            # al día, que es otra cosa y en el reporte sale en 0.
            "avance": (round(c["facturado"] / c["ppto"] * 100, 1)
                       if c["ppto"] and c["facturado"] is not None else None),
        } for c in canales]
        if ppto:
            pct = fact / ppto * 100
            # La consulta del visual filtra 'Calendario'[Mes Nº] SIN filtro de
            # año, así que un mismo mes de dos años distintos se suma contra un
            # presupuesto de un solo mes. Por eso el avance sale por encima de
            # 200%: no es sobrecumplimiento, es doble conteo. El reporte
            # muestra la misma cifra, así que se publica tal cual pero sin
            # semáforo verde y con la advertencia al lado.
            sospechoso = pct > 150
            # Se descartan los KPIs del sondeo genérico: su '% Avance' es la
            # medida plana que nunca respondió al filtro de fecha, y publicaba
            # 1.3% junto al 239% de la tabla del reporte. Dos avances distintos
            # para lo mismo es peor que uno solo con su advertencia.
            res["kpis"] = [k for k in res["kpis"]
                           if "avance" not in k["label"].lower()]
            res["kpis"] = [
                {"label": "Presupuesto del mes", "valor": fmt_soles(ppto)},
                {"label": "Facturado", "valor": fmt_soles(fact)},
                {"label": "Avance vs presupuesto", "valor": f"{pct:.1f}%",
                 "meta": "revisar — el visual no filtra año" if sospechoso else "100%",
                 "estado": "yellow" if sospechoso else
                           ("green" if fact >= ppto else
                            "yellow" if fact >= ppto * 0.8 else "red")},
            ] + res["kpis"]
            if sospechoso:
                res["alerta"] = (f"Avance {pct:.0f}% — la consulta del reporte filtra "
                                 f"mes pero no año, así que suma el mismo mes de "
                                 f"varios años contra un presupuesto mensual")
        pend = sum(c["pendiente"] or 0 for c in canales)
        if pend:
            res["kpis"].append({"label": "Pendiente de facturar",
                                "valor": fmt_soles(pend)})
    return res, avance_pct

# ─── Main ────────────────────────────────────────────────────────────────────

def main():
    print("=== JUANITO Power BI Fetcher v3 ===")
    print(f"Fecha: {HOY}")

    print("\n[1] Autenticando...")
    try:
        token = get_token()
        print("    Token OK")
    except Exception as e:
        print(f"    ERROR: {e}"); sys.exit(1)

    summary = {
        "fecha": hoy_lima().strftime("%d/%m/%Y"),
        "fecha_actualizacion": HOY,
        "generado_por": "JUANITO — Power BI Direct v3",
        "semaforos": {}, "semaforo_razon": {},
        "holding_ventas": "—", "holding_mora": "—",
        "agenda_ceo": {"decidir_hoy": [], "escalar_semana": [], "monitorear": []},
        "empresas": {}
    }

    holding_ventas = 0
    holding_mora_vals = []

    for empresa, ws_id in WORKSPACES.items():
        print(f"\n[{empresa}]")
        ids = dict(DATASET_IDS.get(empresa, {}))
        # Antes de consultar nada: confirmar que los datasets siguen siendo
        # los que dice la configuración.
        ids = resolver_datasets(token, ws_id, empresa, ids)
        empresa_data = {"reportes": {}}

        # Cargar medidas descubiertas (fuente de verdad del discover workflow)
        discovered_path = OUTPUT_DIR / "discovered_measures.json"
        discovered = {}
        if discovered_path.exists():
            try:
                discovered = json.loads(discovered_path.read_text()).get("datasets", {})
                print(f"  Usando discovered_measures.json ({len(discovered)} datasets)")
            except: pass

        # Escanear medidas reales de cada dataset
        measure_cache = load_measure_cache()
        scanned = {}
        for ds_key, did in ids.items():
            print(f"  Escaneando {ds_key}...")
            scanned[ds_key] = scan_dataset(token, ws_id, did, ds_key, measure_cache)
            # Completar con valores del discover (si el scan no encontró algo)
            disc_ds = discovered.get(ds_key, {})
            for m_name, m_vals in disc_ds.items():
                if m_name not in scanned.get(ds_key, {}):
                    val = m_vals.get("dated") if m_vals.get("dated") is not None else m_vals.get("plain")
                    if val is not None:
                        scanned.setdefault(ds_key, {})[m_name] = val
                        print(f"    [{m_name}] = {val} (discovered)")

        # ── Compras: filtrar via TREATAS sobre tabla Calendario (correcto para Live Connection)
        if "compras" in ids:
            ds_c = ids["compras"]
            print(f"  Compras: buscando valor agosto {PREV_YEAR}-{PREV_MONTH:02d} via TREATAS...")
            compras_val = None
            # TREATAS inyecta un rango de fechas como filtro via la relación con Calendario
            # Probamos las variantes de nombre de columna en la tabla Calendario
            for cal_col in ["Date", "Fecha", "fecha", "date", "CalendarDate", "DateKey"]:
                for cal_tbl in ["Calendario", "Calendar", "dCalendario", "Fechas", "Dim_Fecha"]:
                    q = f"""EVALUATE ROW("v",
  CALCULATE(
    SUM('Compras'[data.monto_total_linea]),
    TREATAS(
      FILTER(
        GENERATESERIES(DATE({PREV_YEAR},{PREV_MONTH},1), DATE({PREV_YEAR},{PREV_MONTH},31), 1),
        [Value] <= DATE({PREV_YEAR},{PREV_MONTH},31)
      ),
      '{cal_tbl}'[{cal_col}]
    )
  ))"""
                    # registrar=False: esto prueba nombres de tabla de fecha a
                    # ver cuál existe. Que la mayoría falle es su forma de
                    # trabajar, no un problema — y anotarlas llenaba el
                    # diagnóstico de ruido que la app descarga en cada visita.
                    rows = dax(token, ws_id, ds_c, q,
                               f"treatas_{cal_tbl[:6]}_{cal_col[:4]}", registrar=False)
                    if rows:
                        v = rows[0].get("[v]") or rows[0].get("v")
                        if v is not None:
                            fv = float(v or 0)
                            print(f"    TREATAS '{cal_tbl}'[{cal_col}] → {v}")
                            if 0 < abs(fv) < 50_000_000:
                                compras_val = v
                                print(f"    ✓ Compras agosto: {v}")
                                break
                    if compras_val is not None:
                        break
                if compras_val is not None:
                    break
            if compras_val is not None:
                scanned.setdefault("compras", {})["Valor Compras"] = compras_val
            # No usar fallback total — el SUM acumulado es incorrecto para mostrar como "Compras mes"
            if compras_val is None:
                print(f"    Compras: ratio agosto no accesible via API (Live Connection)")

        # ── Consumo: las cinco tarjetas se leen ejecutando su consulta
        # capturada tal cual. Antes se reconstruían a mano y no coincidían con
        # el reporte: Venta Neta (KG) publicaba 16.75M contra los 10.18M de la
        # tarjeta, y al corregir la columna dio 1.31M — el problema eran los
        # filtros, no la columna. Con el DAX literal no hay nada que acertar.
        # Cada tarjeta va en su propio try: un fallo leyendo una cifra no
        # puede tumbar la extracción entera. La corrida 99 murió aquí y se
        # perdieron los once reportes por culpa de un solo KPI.
        DESTINO_CONSUMO = {
            "Costo Total": "Costo total validado",
            "Venta Neta (KG)": "Peso total KG",
            "Costo x TN Vendida": "Ratio costo / kg",
            "Producción Neta (KG)": "Producción (KG) Odoo",
            "Costo x TN Producida": "Costo x ton producida",
        }
        if "consumo" in ids:
            for (grupo, etiqueta), h in TARJETAS_KPI.items():
                if grupo != "consumo":
                    continue
                try:
                    v = valor_tarjeta_por_hash(token, ws_id, ids["consumo"], h, etiqueta)
                    destino = DESTINO_CONSUMO.get(etiqueta)
                    if v is not None and destino:
                        # guardar_del_reporte, no setdefault: sin la constancia
                        # estas cinco cifras se contaban como del sondeo aunque
                        # salen de la tarjeta del reporte.
                        guardar_del_reporte(scanned, "consumo", destino, v)
                        LEIDOS_DE_TARJETA.append(["consumo_materiales", etiqueta])
                        print(f"    ✓ Consumo [{etiqueta}] desde la tarjeta = {v:,.2f}")
                    else:
                        # Sin esto el fallo es mudo y se queda el valor del
                        # sondeo, que no es el del reporte.
                        print(f"    ✗ Consumo [{etiqueta}] — la tarjeta no devolvió valor")
                        DIAGNOSTICO.append({
                            "consulta": f"tarjeta:{etiqueta}", "http": 200,
                            "error": "la consulta capturada no devolvió filas; "
                                     "queda el valor del sondeo, que no es el "
                                     "del reporte"})
                except Exception as e:
                    print(f"    ✗ Consumo [{etiqueta}]: {e}")
                    DIAGNOSTICO.append({"consulta": f"tarjeta:{etiqueta}", "http": 0,
                                        "error": repr(e)[:300]})

        # ── Consumo: SUM directo desde tablas del dataset consumo
        if "consumo" in ids:
            print("  Consumo: descubriendo tablas...")
            consumo_tables = discover_tables_in_dataset(token, ws_id, ids["consumo"])
            if consumo_tables:
                print(f"    Tablas consumo: {list(consumo_tables.keys())}")
            for tbl_name in consumo_tables:
                cols_c = consumo_tables[tbl_name]
                # Buscar columna de monto/importe/costo
                for col in cols_c:
                    col_lower = col.lower()
                    if any(kw in col_lower for kw in ["monto", "importe", "costo", "valor", "consumo", "total"]):
                        col_bare = col.split("[")[-1].rstrip("]")
                        rows_c = dax(token, ws_id, ids["consumo"],
                                     f"EVALUATE ROW(\"v\", SUM('{tbl_name}'[{col_bare}]))",
                                     f"consumo_sum_{col_bare[:20]}")
                        if rows_c:
                            cv = rows_c[0].get("[v]") or rows_c[0].get("v")
                            if cv is not None and float(cv or 0) != 0:
                                scanned.setdefault("consumo", {})["Total Consumo"] = cv
                                print(f"    Consumo SUM '{col}': {cv}")
                                break
                if scanned.get("consumo", {}).get("Total Consumo"):
                    break

        # ── Productividad: filtro EXACTO confirmado con "Copiar consulta" (2026-09-03)
        # Sobrescribe cualquier valor sin filtrar que haya traído el scan genérico —
        # ese valor sin filtro es garbage (ver dax_productividad_pauno). Solo se piden
        # las 4 medidas confirmadas y validadas con esta consulta exacta.
        if "productividad_ds" in ids:
            print("  Productividad: aplicando filtro exacto confirmado (Copiar consulta)...")
            for m_name, label in [
                ("SUM PROD", "prod_sumprod"),
                ("GASTO TOTAL", "prod_gastototal"),
                ("PROD OPERATIVA", "prod_prodoperativa"),
                ("VENTA KG X SOL", "prod_ventakgxsol"),
            ]:
                v = dax_productividad_pauno(token, ws_id, ids["productividad_ds"], m_name, label)
                if v is not None:
                    scanned.setdefault("productividad_ds", {})[m_name] = v
                    print(f"    ✓ Productividad [{m_name}] filtrado = {v}")
                else:
                    print(f"    ✗ Productividad [{m_name}] — la consulta filtrada no devolvió valor")

            # Venta Neta (KG): no es medida, es columna SUM con 2 filtros propios
            # (CATEGORIZACION="VENTA BRUTA" + TIPO DE NEGOCIO N2) — ver dax_productividad_venta_neta_kg
            v_vn = dax_productividad_venta_neta_kg(token, ws_id, ids["productividad_ds"])
            if v_vn is not None:
                scanned.setdefault("productividad_ds", {})["__VENTA_NETA_KG__"] = v_vn
                print(f"    ✓ Productividad [Venta Neta KG] filtrado = {v_vn}")
            else:
                print("    ✗ Productividad [Venta Neta KG] — la consulta filtrada no devolvió valor")

        # ── Productividad: schema REST + probe medidas reales (único que no es Live Connection)
        if "productividad_ds" in ids and not scanned.get("productividad_ds"):
            print("  Productividad: leyendo schema REST...")
            schema_prod = get_dataset_schema(token, ws_id, ids["productividad_ds"])
            if schema_prod:
                print(f"    Tablas/medidas Prod: { {t: s['measures'][:5] for t,s in schema_prod.items()} }")
                for tbl_name, tbl_info in schema_prod.items():
                    for m_name in tbl_info["measures"]:
                        v, exists = try_measure(token, ws_id, ids["productividad_ds"], m_name)
                        if exists:
                            scanned.setdefault("productividad_ds", {})[m_name] = v
                            print(f"      ✓ Prod [{m_name}] = {v}")
                for tbl_name, tbl_info in schema_prod.items():
                    if tbl_name.lower() in ["calendario", "calendar", "fecha", "medidas"]:
                        continue
                    for col_name in tbl_info["columns"]:
                        col_lower = col_name.lower()
                        if any(kw in col_lower for kw in ["planilla", "monto", "costo", "kg", "produccion", "total"]):
                            rows_s = dax(token, ws_id, ids["productividad_ds"],
                                         f"EVALUATE ROW(\"v\", SUM('{tbl_name}'[{col_name}]))",
                                         f"prod_sum")
                            if rows_s:
                                sv = rows_s[0].get("[v]") or rows_s[0].get("v")
                                if sv is not None and float(sv or 0) != 0:
                                    lbl = "Planilla Total" if any(k in col_lower for k in ["planilla","costo","monto"]) else "Kg Producidos"
                                    scanned.setdefault("productividad_ds", {})[lbl] = sv
                                    print(f"    Prod SUM '{tbl_name}'[{col_name}] = {sv}")

        # Guardar caché de medidas confirmadas — merge con caché anterior (no borrar)
        existing_cache = load_measure_cache()
        for ds_key, found in scanned.items():
            if found:  # solo actualizar si encontramos algo — preservar entradas de runs previos
                existing_cache[ds_key] = {k: None for k in found.keys()}
        save_measure_cache(existing_cache)

        # Merge consumo→compras: tomar medidas del dataset consumo para calcular ratio
        if scanned.get("consumo"):
            cons = scanned["consumo"]
            # Consumo mensual
            consumo_v = cons.get("Consumo")
            if consumo_v is not None:
                scanned.setdefault("compras", {})["Consumo"] = consumo_v
            # Compras mensual desde dataset consumo (si es razonable < S/20M)
            compras_v = cons.get("Compras") or cons.get("Total Compras")
            if compras_v is not None:
                try:
                    if abs(float(compras_v)) < 20_000_000:
                        scanned.setdefault("compras", {})["Valor Compras"] = compras_v
                        print(f"    Compras desde consumo dataset: {compras_v}")
                except: pass
            # Ratio directo desde dataset consumo
            ratio_v = cons.get("Ratio") or cons.get("% Ratio")
            if ratio_v is not None:
                scanned.setdefault("compras", {})["Ratio"] = ratio_v
                print(f"    Ratio desde consumo dataset: {ratio_v}")

        # Combinamos planificacion con inventario (avance vs ppto)
        if scanned.get("planificacion"):
            scanned.setdefault("inventario", {}).update(scanned["planificacion"])

        # ── CxC: aging real (tramos de antigüedad) desde 'DATA_FACTURACION'
        cxc_ds_id = ids.get("cxc")
        if cxc_ds_id:
            try:
                aging = dax_cxc_aging(token, ws_id, cxc_ds_id)
                if aging:
                    scanned.setdefault("cxc", {})["__aging"] = aging
                    for seg, v in aging:
                        print(f"    ✓ CxC aging [{seg}]: {v:,.0f}")
            except Exception as e:
                print(f"    ✗ CxC aging: {e}")
                DIAGNOSTICO.append({"consulta": "cxc_aging", "http": 0,
                                    "error": f"excepcion en Python: {e!r}"})

        # ── Merma por SKU. Es la pregunta que más se repite ("qué producto
        # genera más merma") y la respuesta estaba capturada desde el inicio,
        # sin conectar: la matriz de la pestaña "Declaraciones" abre la merma
        # por SKU dentro de cada almacén. DIF CANT es la desviación contra el
        # estándar, que es de donde sale el porcentaje.
        if empresa == "PAUNO":
            try:
                sku = desglose_desde_captura(
                    token, ws_id,
                    [ids.get("mermas"), ids.get("planificacion")],
                    "Matriz#32dc719b5e1f",
                    {"sku": "[sku]",
                     "almacen": "[almacen_referencia]",
                     "estandar": "[SumSTD_CANT]",
                     "real": "[SumCANT_REAL]",
                     "desvio": "[SumDIF_CANT]",
                     "pct": "[v__Merma__Texto_]"})
                if sku:
                    scanned.setdefault("mermas", {})["__por_sku"] = sku
                    print(f"    ✓ Merma por SKU: {len(sku)} filas")
            except Exception as e:
                print(f"    ✗ merma por sku: {e}")
                DIAGNOSTICO.append({"consulta": "merma_por_sku", "http": 0,
                                    "error": repr(e)[:300]})

        # ── CxC por canal, jefe de venta y ejecutivo, con el aging de cada
        # uno. Sin esto, la mora es un porcentaje sin dueño: no se puede saber
        # a quién pedirle la cobranza.
        if empresa == "PAUNO":
            try:
                canal = desglose_desde_captura(
                    token, ws_id,
                    [ids.get("cxc")],
                    "Matriz#38e3644b220b",
                    {"canal": "[CANAL]",
                     "jefe": "[JEFE_VENTA]",
                     "ejecutivo": "[EJECUTIVO]",
                     "por_vencer": "[POR_VENCER]",
                     "d0_15": "[v0_A_15_DÍAS]",
                     "d16_30": "[v16_A_30_DÍAS]",
                     "mas_30": "[MAS_DE_30_DÍAS]",
                     "total": "[SumTotal_fact]"})
                if canal:
                    scanned.setdefault("cxc", {})["__por_canal"] = canal
                    print(f"    ✓ CxC por canal/ejecutivo: {len(canal)} filas")
            except Exception as e:
                print(f"    ✗ cxc por canal: {e}")
                DIAGNOSTICO.append({"consulta": "cxc_por_canal", "http": 0,
                                    "error": repr(e)[:300]})

        # ── Órdenes de venta por cliente con su margen y su caída. Es el
        # único sitio donde el margen aparece junto a la caída contra lo
        # cotizado, que es lo que explica por qué el margen del mes baja.
        if empresa == "PAUNO":
            try:
                ov = desglose_desde_captura(
                    token, ws_id,
                    [ids.get("margen")],
                    "ORDENES DE VENTA EN EL SISTEMA POR CLIENTE#feef40c9d4c8",
                    {"cliente": "[cliente]",
                     "periodo": "[PERIODO]",
                     "venta": "[SumMonto_Neto_Venta]",
                     "costo": "[SumCOSTO_TOTAL]",
                     "margen": "[v__Margen_Venta__]",
                     "caida": "[v__MARGEN_CAIDA__]",
                     "cantidad": "[SumCANTIDAD_VENTA]"})
                if ov:
                    scanned.setdefault("margen", {})["__ordenes_cliente"] = ov
                    print(f"    ✓ Órdenes de venta por cliente: {len(ov)} filas")
            except Exception as e:
                print(f"    ✗ ordenes por cliente: {e}")
                DIAGNOSTICO.append({"consulta": "ordenes_por_cliente", "http": 0,
                                    "error": repr(e)[:300]})

        # ── CxC
        if scanned.get("cxc"):
            r, sem, razon, mora_pct, vencer = build_cxc(scanned["cxc"])
            empresa_data["reportes"]["cuentas_por_cobrar"] = r
            summary["semaforos"][empresa] = sem
            summary["semaforo_razon"][empresa] = razon
            if mora_pct: holding_mora_vals.append(mora_pct)
            mora_str = f"{mora_pct:.1f}" if mora_pct is not None else "—"
            print(f"  CxC mora={mora_str}% sem={sem}")
            if sem == "red":
                summary["agenda_ceo"]["decidir_hoy"].append({
                    "empresa": empresa,
                    "texto": f"Mora {fmt_pct(mora_pct)} — vencido {fmt_soles(vencer)}. Activar cobranza.",
                    "responsable": "Gerencia Financiera"
                })
        else:
            summary["semaforos"][empresa] = "yellow"
            summary["semaforo_razon"][empresa] = "Sin datos CxC"

        # ── CxP: top 15 proveedores por deuda pendiente
        cxp_ds_id = ids.get("cxp")
        if cxp_ds_id:
            try:
                aging_cxp = dax_cxp_aging(token, ws_id, cxp_ds_id)
                if aging_cxp:
                    scanned.setdefault("cxp", {})["__aging_cxp"] = aging_cxp
                    for est, v in aging_cxp:
                        print(f"    ✓ CxP [{est}]: {v:,.0f}")
                top = dax_cxp_top_proveedores(token, ws_id, cxp_ds_id)
                if top:
                    scanned.setdefault("cxp", {})["__top_proveedores"] = top
                    print(f"    ✓ CxP top proveedores: {len(top)} filas, "
                          f"mayor {top[0][0]} {top[0][2]:,.0f}")
            except Exception as e:
                print(f"    ✗ CxP top proveedores: {e}")
                DIAGNOSTICO.append({"consulta": "cxp_top15", "http": 0,
                                    "error": f"excepcion en Python: {e!r}"})

        # ── CxP
        if scanned.get("cxp"):
            r, dias = build_cxp(scanned["cxp"])
            empresa_data["reportes"]["cuentas_por_pagar"] = r
            dias_str = f"{dias:.0f}" if dias is not None else "—"
            print(f"  CxP dias={dias_str} sem={r['estado']}")
            if r["estado"] == "red" and dias:
                summary["agenda_ceo"]["escalar_semana"].append({
                    "empresa": empresa, "texto": f"CxP {dias:.0f}d — riesgo con proveedores.", "responsable": "Finanzas"
                })

        # ── Margen por cliente (pestaña "R. ORDEN DE VENTA" del reporte 3).
        # Responde de qué cliente viene la caída de margen, no solo de qué
        # unidad de negocio.
        if empresa == "PAUNO":
            try:
                cli = desglose_desde_captura(
                    token, ws_id, [ids.get("margen")],
                    "ORDENES DE VENTA EN EL SISTEMA POR CLIENTE",
                    {"cliente": "[cliente]",
                     "margen": "[v__Margen_Venta__]",
                     "margen_caida": "[v__MARGEN_CAIDA__]",
                     "venta": "[SumMonto_Neto_Venta]",
                     "costo": "[SumCOSTO_TOTAL]"})
                if cli:
                    scanned.setdefault("margen", {})["__por_cliente"] = cli
            except Exception as e:
                print(f"    ✗ margen por cliente: {e}")
                DIAGNOSTICO.append({"consulta": "margen_cliente", "http": 0,
                                    "error": repr(e)[:300]})

        # ── Margen
        if scanned.get("margen"):
            margen_ds_id_total = ids.get("margen")
            if margen_ds_id_total:
                # Sobrescribe Precio x Kilo / Costo x Kilo con el filtro exacto de
                # 12 condiciones confirmado con Copiar consulta (2026-09-04) — el
                # escaneo genérico solo aplicaba filtro de fecha, sin las 11
                # exclusiones adicionales del reporte "Análisis de Ventas".
                p_tot, c_tot = dax_margen_total(token, ws_id, margen_ds_id_total, PREV_YEAR)
                if p_tot is not None and c_tot is not None:
                    scanned["margen"]["Precio x Kilo"] = p_tot
                    scanned["margen"]["Costo x Kilo"] = c_tot
                    print(f"    ✓ Margen total filtrado: precio=S/{p_tot:.2f} costo=S/{c_tot:.2f}")
                else:
                    print("    ✗ Margen total filtrado — la consulta no devolvió valor")
            # Detalle que explica el margen: precio y costo por kilo,
            # abiertos por categoría y tipo de negocio. Es lo que permite
            # pasar de "B&D cae" a "cae en esta categoría, y por precio o
            # por costo". Hasta ahora la app se quedaba en la primera
            # mitad de esa frase.
            try:
                det = desglose_desde_captura(
                    token, ws_id, [margen_ds_id],
                    "DETALLE ANÁLISIS DE COSTOS#5bb77c922cd9",
                    {"categoria": "[Categoria]",
                     "subcategoria": "[Subcategoria]",
                     "negocio": "[TIPO DE NEGOCIO N2]",
                     "peso": "[SumPeso_total]",
                     "venta": "[SumMonto_Neto_Factura_TG_0]",
                     "costo": "[SumCosto_Total]",
                     "precio_kg": "[Precio_x_Kilo]",
                     "costo_kg": "[Costo_x_Kilo]"})
                if det:
                    scanned.setdefault("margen", {})["__detalle_costos"] = det
                    print(f"    ✓ Detalle de costos: {len(det)} filas")

                # Margen por PRODUCTO y por mes. Es el nivel donde se puede
                # poner foco: saber que el margen cae no dice a quién llamar;
                # saber que cae en tres productos que son la mitad del
                # volumen, sí.
                prod = desglose_desde_captura(
                    token, ws_id, [margen_ds_id],
                    "DETALLE ANÁLISIS DE COSTOS#79ac633f8bdb",
                    {"producto": "[DESCRIPCION]",
                     "anio": "[Año]",
                     "mes": "[NroMes]",
                     "venta": "[SumMonto_Neto_Factura_TG_0]",
                     "costo": "[SumCosto_Total]"})
                if prod:
                    scanned.setdefault("margen", {})["__por_producto"] = prod
                    print(f"    ✓ Margen por producto: {len(prod)} filas")

                # SKU cruzado con unidad de negocio. Ningún visual del reporte
                # hace ese cruce, pero el modelo sí lo permite: es la única
                # forma de responder "qué producto de B&D está cayendo" en vez
                # de "algún producto de la empresa está cayendo".
                q = dax_sku_por_uen()
                if q:
                    sku = _tablas_dax(token, ws_id, margen_ds_id, q,
                                      "sku_por_uen")
                    filas = (sku or [[]])[0]
                    if filas:
                        scanned.setdefault("margen", {})["__sku_por_uen"] = filas
                        print(f"    ✓ SKU por unidad de negocio: {len(filas)} filas")
                    else:
                        print("    ✗ SKU por unidad de negocio: sin filas")
            except Exception as e:
                print(f"    ✗ detalle de costos: {e}")
                DIAGNOSTICO.append({"consulta": "detalle_costos", "http": 0,
                                    "error": repr(e)[:300]})

            r, ventas, margen = build_margen(scanned["margen"])
            empresa_data["reportes"]["margen_variable"] = r
            if ventas:
                try: holding_ventas += float(ventas)
                except: pass
            print(f"  Margen {fmt_pct(margen)} ventas={fmt_soles(ventas)} sem={r['estado']}")
            if r["estado"] == "red":
                summary["agenda_ceo"]["escalar_semana"].append({
                    "empresa": empresa, "texto": f"Margen {fmt_pct(margen)} bajo meta 52%.", "responsable": "Comercial"
                })

            # ── Margen por UEN (B&D, MAQUILA, TIGO) — filtro exacto confirmado
            # con Copiar consulta el 2026-09-04 (ver dax_margen_uen). Reemplaza el
            # bloque anterior que adivinaba nombres de tabla ("Maestro Productos",
            # "Productos", "DimProducto") y nunca encontraba dato — la tabla real
            # es 'Exl Tipo de Negocio'[TIPO DE NEGOCIO N1].
            margen_ds_id = ids.get("margen")
            if margen_ds_id:
                uen_data = []
                for uen in ["B&D", "MAQUILA", "TIGO"]:
                    precio, costo = dax_margen_uen(token, ws_id, margen_ds_id, uen, PREV_YEAR, label=f"margen-uen-{uen}")
                    if precio and costo and precio > 0:
                        pct = (precio - costo) / precio * 100
                        uen_data.append({
                            "uen": uen, "margen": f"{pct:.1f}%",
                            "precio_kg": f"S/{precio:.2f}", "costo_kg": f"S/{costo:.2f}",
                            "estado": "green" if pct >= 52 else ("yellow" if pct >= 46 else "red"),
                            # Este margen NO viene del reporte: se deriva de
                            # (precio − costo) / precio con los precios y costos
                            # unitarios de cada UEN. El reporte tiene su propia
                            # medida de margen por unidad de negocio y no da lo
                            # mismo — pondera por producto, no promedia. En
                            # agosto el reporte marcaba B&D 62.30% y este
                            # cálculo daba 65.3%.
                            "fuente": "calculado",
                        })
                        print(f"    ✓ Margen UEN [{uen}]: precio=S/{precio:.2f} costo=S/{costo:.2f} margen={pct:.1f}%")
                    else:
                        print(f"    ✗ Margen UEN [{uen}] — la consulta filtrada no devolvió valor")
                if uen_data:
                    empresa_data["reportes"]["margen_variable"]["por_uen"] = uen_data

                    anotar_derivado(
                        "margen_variable", "por_uen", "margen",
                        "(precio_kg − costo_kg) / precio_kg",
                        "el reporte tiene su propia medida de margen por unidad "
                        "de negocio y pondera por producto; este cálculo promedia. "
                        "En agosto el reporte marcaba B&D 62.30% y esto da 65.3%. "
                        "Falta capturar la consulta del visual MARGEN VARIABLE "
                        "POR UNIDAD DE NEGOCIO del reporte 3.")
                    print(f"  Margen UEN: {uen_data}")

        # ── Consumo de Materiales (PRODUCCIÓN: CECO Ajustado + MIP/TN)
        if scanned.get("consumo"):
            rc = build_consumo_materiales(scanned["consumo"])
            if rc:
                empresa_data["reportes"]["consumo_materiales"] = rc
                print(f"  Consumo Materiales {len(rc['kpis'])} KPIs sem={rc['estado']}")
            else:
                # Live Connection: guardar marcador para que el frontend muestre mensaje correcto
                empresa_data["reportes"]["consumo_materiales"] = {
                    "estado": "grey", "alerta": None, "kpis": [],
                    "live_connection": True
                }
                print("  Consumo Materiales: sin medidas DAX — Live Connection")

        # ── Mermas
        if scanned.get("mermas"):
            r = build_mermas(scanned["mermas"])
            empresa_data["reportes"]["mermas"] = r
            print(f"  Mermas {len(r['kpis'])} KPIs sem={r['estado']}")

            # ── Mermas por UEN — filtro exacto confirmado con Copiar consulta (2026-09-04)
            # Reemplaza el bloque anterior que adivinaba tabla/columna ("Maestro Productos",
            # "UEN", "Unidad Negocio"...) y nunca encontraba dato real.
            # De `ids`, no de DATASET_IDS: `ids` ya trae el dataset corregido.
            mermas_ds_id = ids.get("mermas")
            if mermas_ds_id:
                # (medida en 'Tabla Mermas', filtro extra propio de la tarjeta o None)
                UEN_MERMA_CONFIG = [
                    ("B&D",     "% Merma total B&D",     None),
                    ("TIGO",    "% Merma total TIGO",
                     "FILTER(\n      KEEPFILTERS(VALUES('Tabla Mermas'[TIPO DE BASE])),\n      NOT('Tabla Mermas'[TIPO DE BASE] IN {BLANK()})\n    )"),
                    # Nombre real confirmado con doble M: "MMAQUILA", no "MAQUILA"
                    ("MAQUILA", "% Merma total MMAQUILA",
                     "FILTER(\n      KEEPFILTERS(VALUES('Tabla Mermas'[TIPO DE BASE])),\n      NOT('Tabla Mermas'[TIPO DE BASE] IN {BLANK()})\n    )"),
                ]
                uen_merma = []
                for uen, medida, extra in UEN_MERMA_CONFIG:
                    v = dax_mermas_uen(token, ws_id, mermas_ds_id, medida, PREV_YEAR,
                                        extra_filtro=extra, label=f"merma-uen-{uen}")
                    if v is not None:
                        if abs(v) < 1: v *= 100
                        s = "red" if v > 3 else ("yellow" if v > 2 else "green")
                        uen_merma.append({"uen": uen, "merma": f"{v:.2f}%", "estado": s})
                        print(f"    ✓ Merma UEN [{uen}] filtrado = {v:.2f}%")
                    else:
                        print(f"    ✗ Merma UEN [{uen}] — la consulta filtrada no devolvió valor")
                if uen_merma:
                    empresa_data["reportes"]["mermas"]["por_uen"] = uen_merma
                    print(f"  Mermas UEN: {uen_merma}")

                # Por Planta (ATE / PACHACAMAC / TERCEROS) — filtro exacto confirmado con
                # Copiar consulta el 2026-09-04, capturado desde el gráfico "Merma Mensual
                # por Planta" (pestaña RESUMEN). En el modelo, cada planta corresponde a un
                # almacén con nombre propio (no "ATE"/"PACHACAMAC"/"TERCEROS" literal):
                #   Ate         -> almacén "Lácteos Producción"     · medida genérica
                #   Pachacámac  -> almacén "Salsas Producción"       · medida "...SALSAS"
                #   Terceros    -> almacenes de maquiladores externos (Abuela Maquila,
                #                  Piamonte Maquila, Lácteos Dosimetría) · medida "...MAQUILA"
                #                  (una sola M — distinta de la medida UEN "MMAQUILA")
                # (planta, medida, almacenes, con_tipo_base) — Terceros NO trae el
                # filtro TIPO DE BASE, confirmado ausente en su Copiar consulta.
                PLANTA_MERMA_CONFIG = [
                    ("ATE",        "% Merma total",         ["Lacteos Producción"], True),
                    ("PACHACAMAC", "% Merma total SALSAS",  ["Salsas Producción"], True),
                    ("TERCEROS",   "% Merma total MAQUILA", ["Abuela Maquila", "Piamonte Maquila", "Lacteos Dosimetria"], False),
                ]
                planta_merma = []
                for planta, medida, almacenes, con_tb in PLANTA_MERMA_CONFIG:
                    v = dax_mermas_planta(token, ws_id, mermas_ds_id, medida, almacenes, PREV_YEAR,
                                           con_tipo_base=con_tb, label=f"merma-planta-{planta}")
                    if v is not None:
                        if abs(v) < 1: v *= 100
                        s = "red" if v > 3 else ("yellow" if v > 2 else "green")
                        planta_merma.append({"planta": planta, "merma": f"{v:.2f}%", "estado": s})
                        print(f"    ✓ Merma Planta [{planta}] filtrado = {v:.2f}%")
                    else:
                        print(f"    ✗ Merma Planta [{planta}] — la consulta filtrada no devolvió valor")
                if planta_merma:
                    empresa_data["reportes"]["mermas"]["por_planta"] = planta_merma
                    print(f"  Mermas Planta: {planta_merma}")

                # Si NINGUNA consulta por segmento devolvió valor, el reporte
                # de Power BI cambió y lo único que queda es lo que el sondeo
                # genérico haya encontrado suelto — que no es la merma del
                # mes. El 13 de setiembre eso publicó "Merma Total 4.87%", una
                # cifra vieja y equivocada, presentada como si fuera de hoy.
                # Un dato ausente se nota; uno incorrecto se usa para decidir.
                # Un total es un promedio ponderado de sus partes, asi que
                # NUNCA puede superar a la mayor. El 13/09 se publico "Merma
                # Total 4.87%" con la peor UEN en 4.27% y la peor planta en
                # 3.11%: el total venia del sondeo generico, sin el filtro de
                # mes del reporte, y no medía lo mismo que los segmentos.
                # Publicar una cifra imposible en la tarjeta principal es peor
                # que no publicarla: es la que se mira primero.
                rep_m = empresa_data["reportes"]["mermas"]
                partes = [to_float(x["merma"].rstrip("%"))
                          for x in (uen_merma + planta_merma)]
                partes = [v for v in partes if v is not None]
                if partes:
                    tope = max(partes)
                    for kpi in rep_m.get("kpis", []):
                        val = to_float(str(kpi.get("valor", "")).rstrip("%"))
                        if val is not None and val > tope * 1.05:
                            print(f"    ✗ {kpi['label']} = {kpi['valor']} supera a "
                                  f"su mayor segmento ({tope:.2f}%) — se oculta")
                            DIAGNOSTICO.append({
                                "consulta": "merma_total_incoherente", "http": 200,
                                "error": (f"{kpi['label']} = {kpi['valor']} es mayor que el "
                                          f"peor segmento ({tope:.2f}%); un total no puede "
                                          "superar a sus partes. Se publica sin dato.")})
                            kpi["valor"] = "sin dato"
                            kpi["meta"] = "el total del reporte no cuadra con sus segmentos"
                            kpi["estado"] = "red"
                            if rep_m.get("alerta", "").startswith("Merma "):
                                rep_m["alerta"] = (
                                    f"Merma sin total confiable. Peor UEN y peor planta: "
                                    f"{tope:.2f}%.")

                if not uen_merma and not planta_merma:
                    rep = empresa_data["reportes"]["mermas"]
                    rep["kpis"] = [{"label": "Merma Total", "valor": "sin dato",
                                    "estado": "red",
                                    "meta": "el reporte de Power BI no responde"}]
                    rep["alerta"] = ("Mermas sin dato: las consultas por UEN y por planta "
                                     "fallan contra Power BI. Revisar el reporte 4.")
                    rep["estado"] = "red"
                    for clave in ("por_uen", "por_planta", "por_sku"):
                        rep.pop(clave, None)
                    print("  ✗ Mermas: ningún segmento respondió — se publica 'sin dato'")
                    DIAGNOSTICO.append({
                        "consulta": "mermas_sin_segmentos", "http": 0,
                        "error": ("ninguna consulta por UEN ni por planta devolvió valor; "
                                  "se oculta el KPI para no mostrar una cifra obsoleta")})

        # ── Compras: los KPIs de stock salen de la tabla ANALISIS DE
        # MATERIALES sumando sus filas, que es lo que hace la fila Total del
        # visual. Antes venían del sondeo, sin los filtros del reporte.
        if empresa == "PAUNO":
            try:
                mat = desglose_desde_captura(
                    token, ws_id,
                    [ids.get("compras"), ids.get("planificacion"), ids.get("inventario")],
                    "ANALISIS DE MATERIALES ( PLANIFICACION Y COMPRA)",
                    {"stock_pp": "[Stock_PP]",
                     "stock_val": "[Stock_Valorizado]",
                     "c3": "[Consumo_Prom_3M]",
                     "c6": "[Consumo_Prom_6M]"})
                if mat:
                    def suma(clave):
                        vals = [to_float(x.get(clave)) for x in mat]
                        vals = [v for v in vals if v is not None]
                        return sum(vals) if vals else None
                    scanned.setdefault("compras", {})["__totales_materiales"] = {
                        "Stock PP (Punto de Pedido)": suma("stock_pp"),
                        "Stock Valorizado": suma("stock_val"),
                        "Consumo Prom 3M": suma("c3"),
                        "Consumo Prom 6M": suma("c6"),
                    }
                    print(f"    ✓ Totales de materiales desde {len(mat)} filas")
            except Exception as e:
                print(f"    ✗ totales de materiales: {e}")
                DIAGNOSTICO.append({"consulta": "totales_materiales", "http": 0,
                                    "error": repr(e)[:300]})

        # ── Compras: faltantes y necesidad de compra, desde las capturas del
        # Analizador (pestaña "ANALISIS DE COMPRA" del reporte 5).
        if empresa == "PAUNO":
            cand = [ids.get("compras"), ids.get("planificacion"), ids.get("inventario")]
            try:
                falt = desglose_desde_captura(
                    token, ws_id, cand,
                    "LISTADO DE PRODUCTOS FALTANES EN EXPLOCIÓN DE MATERIALES",
                    {"producto": "[data.nombre_producto]",
                     "codigo": "[data.codigo_producto]",
                     "categoria": "[CATEGORÍA]",
                     "faltante": "[SumFALTANTES_BOOM_SEMNANAL]",
                     "stock": "[Stock]",
                     "pendiente": "[Pendiente_Compra]"})
                if falt:
                    scanned.setdefault("compras", {})["__faltantes"] = falt
            except Exception as e:
                print(f"    ✗ faltantes: {e}")
                DIAGNOSTICO.append({"consulta": "faltantes", "http": 0,
                                    "error": repr(e)[:300]})
            try:
                nec = desglose_desde_captura(
                    token, ws_id, cand, "NECESIDAD DE COMPRA ",
                    {"producto": "[data.nombre_producto]",
                     "categoria": "[CATEGORÍA]",
                     "momento": "[Momento_de_Compra]",
                     "stock": "[Stock]"})
                if nec:
                    scanned.setdefault("compras", {})["__necesidad"] = nec
            except Exception as e:
                print(f"    ✗ necesidad de compra: {e}")
                DIAGNOSTICO.append({"consulta": "necesidad_compra", "http": 0,
                                    "error": repr(e)[:300]})

        # ── Compras
        if scanned.get("compras"):
            r, ratio = build_compras(scanned["compras"])
            empresa_data["reportes"]["compras"] = r
            print(f"  Compras ratio={ratio} sem={r['estado']}")
            if r["estado"] == "red":
                summary["agenda_ceo"]["decidir_hoy"].append({
                    "empresa": empresa,
                    "texto": f"Ratio {ratio:.1f}% — {'jalando inventario' if ratio and ratio>100 else 'sobre-compra'}.",
                    "responsable": "Logística"
                })

        # ── Inventario — filtro exacto confirmado con Copiar consulta (2026-09-04)
        # sobre el reporte '6. Rotación de inventario' (página RI Clasificación).
        inv_ds_id = ids.get("inventario")
        if inv_ds_id:
            cob_total = dax_inventario_kardex(token, ws_id, inv_ds_id, "Días Rotación", PREV_YEAR, PREV_MONTH)
            cob_mp    = dax_inventario_kardex(token, ws_id, inv_ds_id, "Días Rotación MP", PREV_YEAR, PREV_MONTH)
            if cob_total is not None:
                scanned.setdefault("inventario", {})["Cobertura Total (días)"] = cob_total
                print(f"    ✓ Inventario [Cobertura Total] filtrado = {cob_total:.1f} días")
            if cob_mp is not None:
                scanned.setdefault("inventario", {})["Cobertura MP (días)"] = cob_mp
                print(f"    ✓ Inventario [Cobertura MP] filtrado = {cob_mp:.1f} días")

            composicion = dax_inventario_composicion(token, ws_id, inv_ds_id)
            if composicion:
                scanned.setdefault("inventario", {})["_composicion"] = composicion
                print(f"    ✓ Inventario [Composición]: {composicion}")

        # ── S&OP: clasificación del inventario desde '11. Reporte de
        # Planificaciones'. Su datasetId se resuelve desde el id del reporte.
        if empresa == "PAUNO":
            try:
                rr = requests.get(f"{PBI_BASE}/groups/{ws_id}/reports/{SOP_REPORT_ID}",
                                  headers={"Authorization": f"Bearer {token}"}, timeout=25)
                if rr.ok:
                    sop_ds = rr.json().get("datasetId")
                    clas = dax_sop_clasificacion(token, ws_id, sop_ds)
                    if clas:
                        scanned.setdefault("inventario", {})["__clasificacion"] = clas
                        tot = sum(v for _, _, v, _ in clas if v)
                        print(f"    ✓ S&OP clasificación: {len(clas)} filas, "
                              f"saldo total {tot:,.0f}")
                else:
                    print(f"    ✗ reporte S&OP: HTTP {rr.status_code}")
            except Exception as e:
                print(f"    ✗ S&OP clasificación: {e}")
                DIAGNOSTICO.append({"consulta": "sop_clasificacion", "http": 0,
                                    "error": f"excepcion en Python: {e!r}"})

        # ── Inventario
        if scanned.get("inventario"):
            r = build_inventario(scanned["inventario"])
            empresa_data["reportes"]["sop_inventario"] = r
            print(f"  S&OP {len(r['kpis'])} KPIs sem={r['estado']}")

        # ── Control Interno — filtro exacto confirmado con Copiar consulta (2026-09-04)
        # sobre '8. Reporte de auditoria' (Dashboard de Control Interno): solo
        # Empresa="Pauno", sin filtro de fecha (acumulado histórico total).
        control_ds_id = ids.get("control_ds")
        if control_ds_id:
            satisf_v = dax_control_interno(token, ws_id, control_ds_id, "Satisfactorio")
            obs_v    = dax_control_interno(token, ws_id, control_ds_id, "Con Observaciones")
            crit_v   = dax_control_interno(token, ws_id, control_ds_id, "Critico")
            if satisf_v is not None:
                guardar_del_reporte(scanned, "control_ds", "Satisfactorio", satisf_v)
            if obs_v is not None:
                guardar_del_reporte(scanned, "control_ds", "Con Observaciones", obs_v)
            if crit_v is not None:
                guardar_del_reporte(scanned, "control_ds", "Critico", crit_v)
            if satisf_v is not None or obs_v is not None or crit_v is not None:
                print(f"    ✓ Control Interno: Satisfactorio={satisf_v} Con Observaciones={obs_v} Critico={crit_v}")

            planes_abiertos_v = dax_planes_accion(token, ws_id, control_ds_id, "Planes Abiertos")
            if planes_abiertos_v is not None:
                guardar_del_reporte(scanned, "control_ds", "Planes Abiertos", planes_abiertos_v)
                print(f"    ✓ Control Interno: Planes Abiertos={planes_abiertos_v}")

            planes_cerrados_v = dax_planes_accion(token, ws_id, control_ds_id, "Planes Cerrados")
            if planes_cerrados_v is not None:
                guardar_del_reporte(scanned, "control_ds", "Planes Cerrados", planes_cerrados_v)
                print(f"    ✓ Control Interno: Planes Cerrados={planes_cerrados_v}")

            total_planes_v = dax_planes_accion(token, ws_id, control_ds_id, "Total Planes de Acción")
            if total_planes_v is not None:
                guardar_del_reporte(scanned, "control_ds", "Total Planes de Acción", total_planes_v)
                print(f"    ✓ Control Interno: Total Planes de Acción={total_planes_v}")

            planes_atrasados_v = dax_planes_accion(token, ws_id, control_ds_id, "Planes Atrasados")
            if planes_atrasados_v is not None:
                guardar_del_reporte(scanned, "control_ds", "Planes Atrasados", planes_atrasados_v)
                print(f"    ✓ Control Interno: Planes Atrasados={planes_atrasados_v}")

            ejecutado_v = dax_control_interno_sum(token, ws_id, control_ds_id, "¿Ejecutado?")
            if ejecutado_v is not None:
                guardar_del_reporte(scanned, "control_ds", "Ejecutado", ejecutado_v)
                print(f"    ✓ Control Interno: Ejecutado (suma)={ejecutado_v}")

            resultado_acum = dax_control_resultado_acumulado(token, ws_id, control_ds_id)
            if resultado_acum:
                for clave, valor in (
                        ("% Cumplimento", resultado_acum.get("cumplimiento_pct")),
                        ("Calificación", resultado_acum.get("calificacion")),
                        ("Puntos Totales", resultado_acum.get("puntos_totales"))):
                    guardar_del_reporte(scanned, "control_ds", clave, valor)
                print(f"    ✓ Control Interno: Resultado Acumulado={resultado_acum}")

        # ── Control
        if scanned.get("control_ds"):
            r = build_control(scanned["control_ds"])
            if r["kpis"]:
                empresa_data["reportes"]["control_interno"] = r
                print(f"  Control {len(r['kpis'])} KPIs sem={r['estado']}")

        # ── Productividad
        if scanned.get("productividad_ds"):
            r = build_productividad(scanned["productividad_ds"])
            if r["kpis"]:
                empresa_data["reportes"]["productividad"] = r
                print(f"  Productividad {len(r['kpis'])} KPIs")

        # ── Fill Rate: el reporte FILLRATE está en OTRO workspace, y de él
        # solo conocíamos el id del reporte (de su URL), no el del dataset.
        if empresa == "PAUNO":
            try:
                r = requests.get(
                    f"{PBI_BASE}/groups/{WS_FILLRATE}/reports/{FILLRATE_REPORT_ID}",
                    headers={"Authorization": f"Bearer {token}"}, timeout=25)
                if r.ok:
                    fr_ds = r.json().get("datasetId")
                    venta_mes = dax_fillrate_medida_mes(
                        token, WS_FILLRATE, fr_ds, "VENTA", "VENTA")
                    if venta_mes is not None:
                        scanned.setdefault("fill_rate", {})["__venta_mes"] = venta_mes
                        print(f"    ✓ Fill Rate venta del mes: {venta_mes:,.2f}")
                    fr_card = dax_fillrate_medida_mes(
                        token, WS_FILLRATE, fr_ds, "FILLRATE", "FILLRATE")
                    if fr_card is not None:
                        scanned.setdefault("fill_rate", {})["__fillrate_card"] = fr_card
                        print(f"    ✓ Fill Rate (tarjeta): {fr_card:.4f}")
                    tot = dax_fillrate_no_atendido(token, WS_FILLRATE, fr_ds)
                    if tot is not None:
                        scanned.setdefault("fill_rate", {})["__no_atendido_total"] = tot
                        print(f"    ✓ Fill Rate no atendido (mes): {tot:,.0f}")
                    grupos = dax_fillrate_por_grupo(token, WS_FILLRATE, fr_ds)
                    if grupos:
                        scanned.setdefault("fill_rate", {})["__por_grupo"] = grupos
                        peor = grupos[0]
                        print(f"    ✓ Fill Rate peor cliente: {peor[0]} {peor[1]:.3f}")
                    sg = dax_fillrate_soles_por_grupo(token, WS_FILLRATE, fr_ds)
                    if sg:
                        scanned.setdefault("fill_rate", {})["__soles_por_grupo"] = sg
                        top = max(sg.items(), key=lambda kv: kv[1])
                        print(f"    ✓ Fill Rate mayor monto no atendido: {top[0]} {top[1]:,.0f}")
                    marcas = dax_fillrate_por_marca(token, WS_FILLRATE, fr_ds)
                    if marcas:
                        scanned.setdefault("fill_rate", {})["__por_marca"] = marcas
                        for m, v in marcas:
                            print(f"    ✓ Fill Rate no atendido [{m}]: {v:,.0f}")
                else:
                    print(f"    ✗ reporte FILLRATE: HTTP {r.status_code}")
            except Exception as e:
                print(f"    ✗ Fill Rate por marca: {e}")
                DIAGNOSTICO.append({"consulta": "fillrate", "http": 0,
                                    "error": f"excepcion en Python: {e!r}"})

        # ── Fill Rate (dataset dedicado 12. Calculo de Provisiones)
        if scanned.get("fill_rate"):
            fr = build_fill_rate(scanned["fill_rate"])
            if fr["kpis"]:
                empresa_data["reportes"]["fill_rate"] = fr
                print(f"  Fill Rate {fr['kpis'][0]['valor'] if fr['kpis'] else '—'}")

        # ── Avance vs Presupuesto: tabla por canal del reporte 11,
        # pestaña RESUMEN 2. El mes lo fija un segmentador; se pide el mes
        # anterior completo, que es el corte que usa el resto del dashboard.
        if empresa == "PAUNO":
            try:
                rr = requests.get(f"{PBI_BASE}/groups/{ws_id}/reports/{SOP_REPORT_ID}",
                                  headers={"Authorization": f"Bearer {token}"}, timeout=25)
                if rr.ok:
                    ds_av = rr.json().get("datasetId")
                    hoy = hoy_lima()
                    mes_ant = hoy.month - 1 or 12
                    cump = dax_cumplimiento_produccion(token, ws_id, ds_av)
                    if cump:
                        scanned.setdefault("inventario", {})["__cumpl_produccion"] = cump
                        print(f"    ✓ Cumplimiento producción (ayer): {len(cump)} categorías")
                    canales = dax_avance_por_canal(token, ws_id, ds_av, mes_ant)
                    if canales:
                        scanned.setdefault("inventario", {})["__avance_canal"] = canales
                        print(f"    ✓ Avance por canal (mes {mes_ant}): "
                              f"{len(canales)} canales")
            except Exception as e:
                print(f"    ✗ Avance por canal: {e}")
                DIAGNOSTICO.append({"consulta": "avance_canal", "http": 0,
                                    "error": f"excepcion en Python: {e!r}"})

        # ── Avance vs Presupuesto (planificacion mergeado en inventario)
        if scanned.get("inventario"):
            av, avance_pct = build_avance(scanned["inventario"])
            if av["kpis"]:
                empresa_data["reportes"]["margen_variable_pag2"] = av
                print(f"  Avance PPTO {fmt_pct(avance_pct)} sem={av['estado']}")

        summary["empresas"][empresa] = empresa_data

    if holding_ventas: summary["holding_ventas"] = fmt_soles(holding_ventas)
    if holding_mora_vals:
        summary["holding_mora"] = f"{sum(holding_mora_vals)/len(holding_mora_vals):.1f}%"

    # Inyectar datos YoY del discover si existen
    discovered_path = OUTPUT_DIR / "discovered_measures.json"
    if discovered_path.exists():
        try:
            disc = json.loads(discovered_path.read_text())
            yoy  = disc.get("yoy", {})
            comp = disc.get("comp_filtro", "")
            if yoy:
                summary["yoy_periodo"] = comp
                summary["yoy"] = {}
                # Ventas año anterior
                ventas_yoy = (yoy.get("margen", {}) or {}).get("Venta Total")
                if ventas_yoy is not None:
                    summary["yoy"]["ventas"] = ventas_yoy
                # Fill Rate año anterior
                fr_yoy = (yoy.get("fill_rate", {}) or {}).get("% Fill Rate")
                if fr_yoy is not None:
                    summary["yoy"]["fill_rate"] = fr_yoy
                # Mermas año anterior
                mermas_yoy = (yoy.get("mermas", {}) or {}).get("% Merma Total")
                if mermas_yoy is not None:
                    summary["yoy"]["mermas"] = mermas_yoy
                # Avance año anterior
                avance_yoy = (yoy.get("planificacion", {}) or {}).get("% Avance")
                if avance_yoy is not None:
                    summary["yoy"]["avance"] = avance_yoy
                # Morosidad año anterior
                # Solo '% Morosidad': el respaldo a 'Morosidad' compararía el
                # año actual (21.08%) contra otra medida distinta (13.78%) y
                # produciría un YoY falso. Ver SCAN_CANDIDATES["cxc"].
                mora_yoy = (yoy.get("cxc", {}) or {}).get("% Morosidad")
                if mora_yoy is not None:
                    summary["yoy"]["morosidad"] = mora_yoy
                print(f"  YoY {comp}: ventas={ventas_yoy} fr={fr_yoy} mermas={mermas_yoy}")
        except Exception as e:
            print(f"  YoY no disponible: {e}")

    # Los KPIs de tarjeta del propio reporte mandan sobre los del sondeo.
    try:
        ids_p = datasets_de("PAUNO")
        candidatos = {
            "cuentas_por_cobrar": [ids_p.get("cxc")],
            "control_interno":    [ids_p.get("control_ds")],
            "productividad":      [ids_p.get("productividad_ds"), ids_p.get("consumo")],
            "consumo_materiales": [ids_p.get("consumo"), ids_p.get("productividad_ds")],
        }
        reps = (summary.get("empresas", {}).get("PAUNO", {}) or {}).get("reportes", {})
        n = aplicar_kpis_capturados(token, WORKSPACES["PAUNO"], candidatos, reps)

        # Tarjetas que comparten nombre y solo se distinguen por sus filtros
        for pref, med, req, prohib, tipo, etiqueta, formato in TARJETAS_POR_FIRMA:
            rep = reps.get(tipo)
            if rep is None:
                continue
            v = tarjeta_por_firma(token, WORKSPACES["PAUNO"],
                                  [ids_p.get("cxp")], pref, med, req, prohib)
            if v is None:
                continue
            texto = _fmt_kpi(v, formato)
            actual = next((k for k in rep.setdefault("kpis", [])
                           if k["label"] == etiqueta), None)
            if actual is None:
                rep["kpis"].append({"label": etiqueta, "valor": texto})
                n += 1
            elif actual.get("valor") != texto:
                print(f"    ~ [{tipo}] {etiqueta}: {actual['valor']} → {texto} "
                      f"(tarjeta identificada por sus filtros)")
                actual["valor"] = texto
                n += 1
        if n:
            print(f"  ✓ {n} KPI(s) tomados de las tarjetas del reporte")
    except Exception as e:
        print(f"  ✗ KPIs capturados: {e}")
        DIAGNOSTICO.append({"consulta": "kpis_capturados", "http": 0,
                            "error": repr(e)[:300]})

    # ── Control Interno: cumplimiento por jefe de área y planes por planta.
    # Sus 14 KPIs venían del sondeo y nunca se habían contrastado; estas dos
    # tablas del reporte 8 dicen además QUIÉN y DÓNDE, que es lo accionable.
    try:
        ids_ci = [datasets_de("PAUNO").get("control_ds")]
        reps_ci = (summary.get("empresas", {}).get("PAUNO", {}) or {}).get("reportes", {})
        ci = reps_ci.get("control_interno")
        if ci is not None:
            areas = desglose_desde_captura(
                token, WORKSPACES["PAUNO"], ids_ci,
                "RESULTADO DE CUMPLIMIENTO ACUMULADO",
                {"jefe": "[Jefe de Area]", "cumplimiento": "[v__Cumplimento]",
                 "color": "[Color_KPI]"})
            filas = []
            for a in areas:
                jefe = (a.get("jefe") or "").strip()
                v = to_float(a.get("cumplimiento"))
                if not jefe or v is None:
                    continue
                pct = v * 100 if abs(v) <= 1.5 else v
                filas.append({"jefe": jefe, "valor": f"{pct:.1f}%",
                              "pct": round(pct, 1),
                              "estado": ("red" if pct < 60 else
                                         "yellow" if pct < 80 else "green")})
            if filas:
                filas.sort(key=lambda x: x["pct"])
                ci["por_area"] = filas
                peor = filas[0]
                print(f"    ✓ Control interno por área: {len(filas)} — "
                      f"peor {peor['jefe']} {peor['valor']}")

            planes = desglose_desde_captura(
                token, WORKSPACES["PAUNO"], ids_ci, "ESTADO DE PLANES DE ACCION",
                # El visual no expone el nombre del área, solo su conteo
                # ([CountÁrea]), así que se agrupa por planta y estatus.
                {"planta": "[Planta]", "estatus": "[Estatus]"})
            conteo = {}
            for pl in planes:
                estatus = (pl.get("estatus") or "").strip()
                if not estatus or estatus == "—":
                    continue          # fila de subtotal del visual
                clave = ((pl.get("planta") or "—").strip(), estatus)
                conteo[clave] = conteo.get(clave, 0) + 1
            if conteo:
                ci["planes_por_planta"] = [
                    {"planta": p, "estatus": e, "planes": n}
                    for (p, e), n in sorted(conteo.items(), key=lambda kv: -kv[1])]
                print(f"    ✓ Planes de acción: {len(conteo)} combinaciones planta/estatus")
    except Exception as e:
        print(f"    ✗ Control interno por área: {e}")
        DIAGNOSTICO.append({"consulta": "control_interno_areas", "http": 0,
                            "error": repr(e)[:300]})

    # Origen de cada cifra: del reporte o del sondeo genérico.
    try:
        reps_m = (summary.get("empresas", {}).get("PAUNO", {}) or {}).get("reportes", {})
        try:
            ver, tot = marcar_origen(reps_m, scanned)
        except Exception as e:
            # Clasificar el origen de las cifras es informativo: si falla, no
            # debe llevarse por delante una extracción que ya funcionó.
            print(f"  ✗ marcar_origen: {e}")
            DIAGNOSTICO.append({"consulta": "marcar_origen", "http": 0,
                                "error": repr(e)[:300]})
            ver = tot = 0
        summary["verificacion"] = {"verificados": ver, "total": tot}
        print(f"\n  Origen de los KPIs: {ver}/{tot} desde el reporte "
              f"({ver / tot * 100:.0f}%)" if tot else "")
    except Exception as e:
        print(f"  ✗ marcar origen: {e}")

    if DIAGNOSTICO:
        etiquetar_periodos(summary)
        # Cobertura: cuántas cifras vienen de una consulta del reporte y
        # cuántas del sondeo genérico. Es el número que responde "¿está todo
        # validado?" sin depender de ninguna lista mantenida a mano.
        summary["cobertura"] = {"kpis_del_reporte": ver, "kpis_totales": tot,
                                "pct": round(ver / tot * 100, 1) if tot else None}
        summary["diagnostico"] = DIAGNOSTICO
        # La lista de cifras calculadas viaja con los datos: quien mire la app
        # puede saber cuáles son derivadas sin leer el código.
        summary["derivados"] = DERIVADOS
        summary["kpis_de_tarjeta"] = LEIDOS_DE_TARJETA
        if DERIVADOS:
            print(f"\n  {len(DERIVADOS)} cifra(s) calculadas, no leídas de Power BI:")
            for x in DERIVADOS:
                print(f"    · {x['reporte']}/{x['desglose']}.{x['campo']} = {x['formula']}")
        print(f"\n⚠ {len(DIAGNOSTICO)} consultas con problema — detalle en "
              f"summaries.json → diagnostico")

    out = OUTPUT_DIR / "summaries.json"
    out.write_text(json.dumps(summary, ensure_ascii=False, indent=2))
    print(f"\n✓ summaries.json guardado ({out})")

    hist_path = HIST_DIR / "historico.json"
    try:
        historico = json.loads(hist_path.read_text()) if hist_path.exists() else {
            "generado_por": "JUANITO — Módulo Histórico",
            "ultima_actualizacion": HOY,
            "empresas": {"PAUNO": {"meses": []}}
        }
        mes_label = hoy_lima().strftime("%b-%y")
        for emp, ed in summary.get("empresas", {}).items():
            meses = historico["empresas"].setdefault(emp, {"meses": []})["meses"]
            ventas_kpi = next((k["valor"] for k in ed.get("reportes",{}).get("margen_variable",{}).get("kpis",[]) if "venta" in k.get("label","").lower()), "—")
            mora_kpi   = next((k["valor"] for k in ed.get("reportes",{}).get("cuentas_por_cobrar",{}).get("kpis",[]) if "mora" in k.get("label","").lower()), "—")
            entrada = {"mes": mes_label, "ventas": ventas_kpi, "mora": mora_kpi}
            if not meses or meses[-1]["mes"] != mes_label:
                meses.append(entrada)
                if len(meses) > 12: meses.pop(0)
        historico["ultima_actualizacion"] = HOY
        hist_path.write_text(json.dumps(historico, ensure_ascii=False, indent=2))
        print("✓ historico.json actualizado")
    except Exception as e:
        print(f"  Histórico error: {e}")

    print("\n=== COMPLETADO ===")

if __name__ == "__main__":
    # Un fallo en cualquier punto tumbaba la corrida entera y el dashboard se
    # quedaba con los datos del día anterior, aunque el 95% de las consultas
    # hubiera funcionado. Se deja constancia del error en summaries.json —
    # donde sí se puede leer sin permisos de admin sobre el log— y se sale con
    # código 1 igual, para que la corrida siga marcada como fallida.
    import traceback
    try:
        main()
    except Exception:
        rastro = traceback.format_exc()
        print(rastro)
        try:
            ruta = OUTPUT_DIR / "summaries.json"
            datos = json.loads(ruta.read_text(encoding="utf-8")) if ruta.exists() else {}
            datos.setdefault("diagnostico", []).append({
                "consulta": "main", "http": 0,
                "error": "la corrida murió: " + rastro[-1500:]})
            ruta.write_text(json.dumps(datos, ensure_ascii=False, indent=2))
            print("→ traceback guardado en summaries.json → diagnostico")
        except Exception as e2:
            print(f"→ no se pudo guardar el traceback: {e2}")
        # Y en un archivo aparte, porque cuando la corrida muere aquí el paso
        # que publica summaries.json ni siquiera llega a ejecutarse: el error
        # se queda dentro del runner y solo se puede leer con permisos sobre
        # el log de Actions. Este archivo sí se publica, y no toca los datos
        # que lee la app.
        try:
            (OUTPUT_DIR / "ultimo_error.txt").write_text(
                f"corrida del {datetime.datetime.now(datetime.timezone.utc):%Y-%m-%d %H:%M UTC}\n\n"
                + rastro, encoding="utf-8")
            print("→ traceback guardado en data/latest/ultimo_error.txt")
        except Exception as e3:
            print(f"→ no se pudo guardar ultimo_error.txt: {e3}")
        sys.exit(1)
