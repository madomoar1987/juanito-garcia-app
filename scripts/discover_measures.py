#!/usr/bin/env python3
"""
Descubre TODAS las medidas reales de cada dataset via schema REST + executeQueries.
Guarda resultados en data/latest/discovered_measures.json para que fetch_powerbi.py
use los nombres exactos sin adivinar.
Incluye comparativo YoY: mes actual vs mismo mes del año anterior.
"""
import os, json, requests, datetime
from pathlib import Path

TENANT_ID     = os.environ["AZURE_TENANT_ID"]
CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
USERNAME      = os.environ["PBI_USERNAME"]
PASSWORD      = os.environ["PBI_PASSWORD"]
WS_ID         = "461932ad-b5ec-4fd6-aa97-f1fc7bdc5169"

TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
PBI_BASE  = "https://api.powerbi.com/v1.0/myorg"

r = requests.post(TOKEN_URL, data={
    "grant_type": "password", "client_id": CLIENT_ID,
    "client_secret": CLIENT_SECRET, "username": USERNAME,
    "password": PASSWORD, "scope": "https://analysis.windows.net/powerbi/api/.default",
})
token = r.json()["access_token"]
H = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

today = datetime.date.today()
PREV_MONTH = today.month - 1 if today.month > 1 else 12
PREV_YEAR  = today.year  if today.month > 1 else today.year - 1

# Año anterior mismo mes (comparativo YoY)
COMP_YEAR  = PREV_YEAR - 1
COMP_MONTH = PREV_MONTH

DATASETS = {
    "margen":        "38076daa-d2cd-4a93-858a-82c0a4cf8cb6",
    "compras":       "06408938-8202-424e-80c0-b42c178dabde",
    "inventario":    "0e27d784-41a4-48f0-9208-60210119f0a7",
    "mermas":        "fdd58a2d-654d-41e0-b587-3d43c337ed47",
    "consumo":       "c972c8cb-e5fc-4b60-8f5e-265a78e1e796",
    "planificacion": "30074d92-7ec1-4762-82f2-1cb29c15dcfe",
    "cxc":           "2eec70cd-0820-408f-b938-a2cd547b0c18",
    "cxp":           "45a8ab8d-e162-4398-a260-3a9a5f90829f",
    "fill_rate":     "7f4ebe22-5e90-4e35-973b-4af3c58497e5",
}

DATE_CTX = {
    "margen":        ("Calendario", "Date"),
    "mermas":        ("Calendario", "Date"),
    "compras":       ("Calendario", "Date"),
    "inventario":    ("Calendario", "Date"),
    "consumo":       ("Calendario", "Date"),
    "planificacion": ("Calendario", "Date"),
    "fill_rate":     ("Calendario", "Date"),
}

CANDIDATOS = {
    "margen":   ["Venta Total","Ventas","Margen Variable","% Margen Variable","MV","% MV","MV%",
                 "R. MARGEN","R. Margen","R MARGEN","R. COSTO UNIT","R. Costo Unit",
                 "Margen Variable %","Pct MV","% MV PAUNO","MV PAUNO",
                 "MARGEN VARIABLE","Venta Neta","Precio/kg","Precio x Kilo","Precio Kilo",
                 "Costo/kg","Costo x Kilo","Costo Kilo","Margen Bruto","% Margen"],
    "compras":  ["Ratio","Cant Compra","Cant Consumo","Eficiencia","Eficiencia Costo",
                 "Ratio C/V","% Ratio","Consumo/Compra","Valor Compras","Total Compras",
                 "Monto Compras","Compras","Importe Compras","Compra Mes","Consumo Mes",
                 "Monto Compra","Monto Consumo","Ratio Operativo","% Eficiencia"],
    "inventario":["Stock Total","Dead Stock","% Dead Stock","Inventario Total",
                  "Working Stock","Rotacion","Dias Inventario","Valor Stock",
                  "Stock Valorizado","Inmovilizado","% Inmovilizado"],
    "mermas":   ["% Merma Total", "% MERMAS", "% Merma", "Merma Total", "Merma %",
                 "% Merma Ate", "% Merma Pachacamac", "% Merma Terceros",
                 "% Merma B&D", "% Merma BD", "% Merma Tigo", "% Merma TIGO",
                 "% Merma Maquila", "% Merma MAQUILA",
                 "REAL", "STD", "DESVIACION", "ITEM DESVIADO",
                 "Merma Kg", "Merma Soles", "Merma Valorizada", "Kg Merma"],
    "consumo":  ["Consumo","Compras","Total Consumo","Total Compras","Ratio","% Ratio",
                 "Eficiencia","Cant Compra","Cant Consumo","Monto Consumo","Monto Compras",
                 "Consumo Periodo","Compra Periodo","Ratio C/C","Importe Consumo",
                 "Importe Compras","% Eficiencia","Ratio Operativo"],
    "planificacion":["% Avance","Avance","Ventas Real","Fill Rate","% Fill Rate",
                     "Presupuesto","PPTO","Avance vs Ppto","% Cumplimiento Ppto"],
    "cxc":      ["% Morosidad","Morosidad","Por Vencer","CxC Total","Vencido",
                 "Saldo CxC","Total CxC","% Vencido","Cartera Vencida"],
    "cxp":      ["Cuentas x Pagar","CUENTAS X PAGAR","Refinanciamiento","Proveedores",
                 "CxP Total","Total CxP","Saldo CxP","Deuda Total"],
    "fill_rate":["% FILLRATE","% Fill Rate","ORDEN DE VENTA","FACTURACION","VENTA PERDIDA",
                 "Fill Rate","Orden de Venta","Facturacion","Venta Perdida",
                 "% Cumplimiento","Pedidos","Atendidos","No Atendidos"],
}

# Medidas clave a consultar para YoY (solo las confirmadas que responden a filtro fecha)
KPI_YOY = {
    "margen":        ["Venta Total", "R. MARGEN", "Margen Variable", "% Margen Variable", "MV", "% MV"],
    "mermas":        ["% Merma Total", "% Merma Ate", "% Merma Pachacamac", "% Merma Terceros",
                     "% Merma B&D", "% Merma Tigo", "% Merma TIGO", "% Merma Maquila"],
    "planificacion": ["% Avance"],
    "fill_rate":     ["% Fill Rate", "ORDEN DE VENTA", "FACTURACION", "VENTA PERDIDA"],
    "cxc":           ["% Morosidad", "Morosidad"],
    "inventario":    ["Stock Total"],
}

def dax(ds_id, query):
    import time
    url = f"{PBI_BASE}/groups/{WS_ID}/datasets/{ds_id}/executeQueries"
    for attempt in range(2):
        resp = requests.post(url, headers=H,
                             json={"queries": [{"query": query}],
                                   "serializerSettings": {"includeNulls": True}},
                             timeout=25)
        if resp.status_code == 429:
            time.sleep(int(resp.headers.get("Retry-After", 15)))
            continue
        if not resp.ok:
            return None
        rows = resp.json().get("results",[{}])[0].get("tables",[{}])[0].get("rows",[])
        return rows
    return None

def get_schema(ds_id):
    url = f"{PBI_BASE}/groups/{WS_ID}/datasets/{ds_id}/tables"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=20)
    if not r.ok:
        return {}
    schema = {}
    for tbl in r.json().get("value", []):
        name = tbl.get("name","")
        schema[name] = {
            "columns":  [c["name"] for c in tbl.get("columns",[])],
            "measures": [m["name"] for m in tbl.get("measures",[])],
        }
    return schema

def get_all_measures_dmv(ds_id):
    rows = dax(ds_id, "EVALUATE INFO.MEASURES()")
    if not rows:
        return []
    measures = []
    for row in rows:
        name = (row.get("[Name]") or row.get("[MEASURE_NAME]") or
                row.get("Name") or row.get("MEASURE_NAME") or "")
        if name:
            measures.append(name)
    return measures

def get_all_tables_dmv(ds_id):
    rows = dax(ds_id, "EVALUATE INFO.TABLES()")
    if not rows:
        return []
    tables = []
    for row in rows:
        name = row.get("[Name]") or row.get("Name") or ""
        if name:
            tables.append(name)
    return tables

def try_plain(ds_id, m):
    rows = dax(ds_id, f'EVALUATE ROW("v", [{m}])')
    if rows and rows[0]:
        return list(rows[0].values())[0]
    return None

def try_dated(ds_id, m, tbl, col, year=None, month=None):
    y  = year  or PREV_YEAR
    mo = month or PREV_MONTH
    q = (f'EVALUATE ROW("v", CALCULATE([{m}],'
         f'FILTER(ALL(\'{tbl}\'), YEAR(\'{tbl}\'[{col}]) = {y} && MONTH(\'{tbl}\'[{col}]) = {mo})))')
    rows = dax(ds_id, q)
    if rows and rows[0]:
        return list(rows[0].values())[0]
    return None

print(f"=== DISCOVERY COMPLETO === {PREV_YEAR}-{PREV_MONTH:02d} vs YoY {COMP_YEAR}-{COMP_MONTH:02d}\n")

results     = {}  # {ds_key: {measure: {plain, dated}}}
results_yoy = {}  # {ds_key: {measure: valor_año_anterior}}

for ds_key, ds_id in DATASETS.items():
    print(f"\n{'='*60}")
    print(f"DATASET: {ds_key} ({ds_id[:8]}...)")
    date_ctx = DATE_CTX.get(ds_key)
    results[ds_key] = {}

    # 1. Descubrir medidas via INFO.MEASURES()
    all_measures = get_all_measures_dmv(ds_id)
    if all_measures:
        print(f"  INFO.MEASURES() -> {len(all_measures)} medidas: {all_measures[:20]}")
        if len(all_measures) > 20:
            print(f"    ... y {len(all_measures)-20} mas")
    else:
        schema = get_schema(ds_id)
        if schema:
            for tbl_name, tbl_info in schema.items():
                meds = tbl_info.get("measures", [])
                if meds:
                    print(f"  Schema '{tbl_name}': {meds}")
                    all_measures.extend(meds)
        if not all_measures:
            print("  Sin medidas por DMV ni schema — usando candidatos")
            all_measures = CANDIDATOS.get(ds_key, [])

    tables = get_all_tables_dmv(ds_id)
    if tables:
        print(f"  INFO.TABLES() -> {tables}")

    # 2. Probar cada medida (mes actual)
    print(f"  Probando {len(all_measures)} medidas (mes actual {PREV_YEAR}-{PREV_MONTH:02d})...")
    for m in all_measures:
        v_plain = try_plain(ds_id, m)
        if v_plain is None:
            continue
        entry = {"plain": v_plain}
        if date_ctx:
            v_dated = try_dated(ds_id, m, date_ctx[0], date_ctx[1])
            entry["dated"] = v_dated
            print(f"  OK [{m}]  sin_filtro={v_plain}  |  {PREV_YEAR}-{PREV_MONTH:02d}={v_dated}")
        else:
            print(f"  OK [{m}] = {v_plain}")
        results[ds_key][m] = entry

    # 3. YoY: mismo mes año anterior para medidas clave
    yoy_measures = KPI_YOY.get(ds_key, [])
    if yoy_measures and date_ctx:
        results_yoy[ds_key] = {}
        print(f"  YoY {COMP_YEAR}-{COMP_MONTH:02d}: consultando {yoy_measures}...")
        for m in yoy_measures:
            v_yoy = try_dated(ds_id, m, date_ctx[0], date_ctx[1],
                              year=COMP_YEAR, month=COMP_MONTH)
            if v_yoy is not None:
                results_yoy[ds_key][m] = v_yoy
                print(f"  OK YoY [{m}] {COMP_YEAR}-{COMP_MONTH:02d} = {v_yoy}")
            else:
                print(f"  NO YoY [{m}] sin datos para {COMP_YEAR}-{COMP_MONTH:02d}")

out_path = Path("data/latest/discovered_measures.json")
out_path.parent.mkdir(parents=True, exist_ok=True)
out_path.write_text(json.dumps({
    "generado":    str(today),
    "mes_filtro":  f"{PREV_YEAR}-{PREV_MONTH:02d}",
    "comp_filtro": f"{COMP_YEAR}-{COMP_MONTH:02d}",
    "datasets":    results,
    "yoy":         results_yoy,
}, ensure_ascii=False, indent=2))
print(f"\nGuardado en {out_path}")
print(f"YoY {COMP_YEAR}-{COMP_MONTH:02d} incluido para: {list(results_yoy.keys())}")
print("\n=== FIN ===")
