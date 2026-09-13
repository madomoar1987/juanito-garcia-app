#!/usr/bin/env python3
"""Ejecuta las tarjetas del reporte y las compara con lo que publica la app.

Es la validación que pedía el usuario sin tener que mandar capturas de
pantalla cada día. Las exportaciones del Analizador de rendimiento contienen
la consulta exacta de cada tarjeta —la misma que Power BI ejecuta para pintar
el número en pantalla—, así que ejecutarlas y comparar equivale a mirar el
reporte, pero automático.

No reemplaza al validador cruzado (validar_datos.py): aquel comprueba que las
cifras sean coherentes entre sí; este comprueba que sean las del reporte.

Corre dentro del workflow, donde están las credenciales. Anota el resultado en
summaries.json y nunca corta la corrida: publicar con una advertencia sirve
más que no publicar.

Uso:  python scripts/validar_tarjetas.py [--anotar]
"""

import json
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).parent))
import fetch_powerbi as F                                   # noqa: E402

RUTA = pathlib.Path("data/latest/summaries.json")

# Tarjeta del reporte → KPI que publica la app.
#   (nombre del visual, archivo de captura) : (reporte, etiqueta del KPI)
#
# El nombre del visual se compara en mayúsculas y sin acentos. Las tarjetas
# que no están aquí se cuentan aparte: no se inventan correspondencias, porque
# emparejar mal es peor que no emparejar.
MAPA = {
    ("MOROSIDAD", "01-cxc"):            ("cuentas_por_cobrar", "Morosidad"),
    ("CUENTAS POR COBRAR", "01-cxc"):   ("cuentas_por_cobrar", "CxC Total"),
    ("POR VENCER", "01-cxc"):           ("cuentas_por_cobrar", "CxC por vencer"),
    ("# CLIENTES", "01-cxc"):           ("cuentas_por_cobrar", "Clientes"),
    ("VENTAS MES ACTUAL", "01-cxc"):    ("cuentas_por_cobrar", "Ventas del mes"),
    ("ROTACION CXC", "01-cxc"):         ("cuentas_por_cobrar", "Rotación CxC"),
    ("COSTO TOTAL", "14-consumo"):      ("consumo_materiales", "Costo Total"),
    ("VENTA NETA (KG)", "14-consumo"):  ("consumo_materiales", "Venta Neta (KG)"),
    ("COSTO X TN VENDIDA", "14-consumo"):   ("consumo_materiales", "Costo x TN Vendida"),
    ("COSTO X TN PRODUCIDA", "14-consumo"): ("consumo_materiales", "Costo x TN Producida"),
    ("PRODUCCION NETA (KG)", "14-consumo"): ("consumo_materiales", "Producción Neta (KG)"),
    ("PLANILLA TOTAL (S/.)", "13-productividad"): ("productividad", "Planilla Total"),
    ("PRODUCCION TOTAL (KG)", "13-productividad"): ("productividad", "Producción Total (KG)"),
    ("VENTA NETA (KG)", "13-productividad"):   ("productividad", "Venta Neta (KG)"),
}

# Dataset donde vive cada archivo de captura.
DATASET = {
    "01-cxc": "cxc", "02-cxp": "cxp", "03-margen": "margen",
    "04-mermas": "mermas", "05-compras": "compras",
    "08-auditoria": "control_ds", "11-planificaciones": "planificacion",
    "12-provisiones": "fill_rate", "13-productividad": "productividad_ds",
    "14-consumo": "consumo",
}

TOL = 0.02


def sin_acentos(t):
    for a, b in [("Á", "A"), ("É", "E"), ("Í", "I"), ("Ó", "O"), ("Ú", "U"), ("Ñ", "N")]:
        t = t.replace(a, b)
    return t


def clave_visual(q):
    return (sin_acentos((q.get("visual") or "").upper().strip()),
            (q.get("archivo") or "").split("__")[0])


def valor_de(filas):
    """Primer número de la primera fila. Una tarjeta devuelve una sola cifra."""
    if not filas:
        return None
    for v in filas[0].values():
        if isinstance(v, (int, float)):
            return float(v)
    return None


def main():
    anotar = "--anotar" in sys.argv
    if not RUTA.exists():
        sys.exit(f"no existe {RUTA}")
    datos = json.loads(RUTA.read_text(encoding="utf-8"))
    reportes = (datos.get("empresas", {}).get("PAUNO", {}) or {}).get("reportes", {})

    catalogo = json.loads(F.CATALOGO_CAPTURAS.read_text(encoding="utf-8"))
    tarjetas = [q for q in catalogo if (q.get("filas") or 0) == 1]

    token = F.get_token()
    ws = F.WORKSPACES["PAUNO"]
    ids = F.datasets_de("PAUNO")

    print(f"Validando {len(tarjetas)} tarjetas del reporte contra la app\n")
    iguales, distintas, sin_mapa, sin_dato = [], [], 0, 0

    for q in tarjetas:
        clave = clave_visual(q)
        destino = MAPA.get(clave)
        if not destino:
            sin_mapa += 1
            continue
        rep_k, etiqueta = destino
        ds_id = ids.get(DATASET.get(clave[1], ""))
        if not ds_id:
            sin_dato += 1
            continue

        tablas = F.tablas_de_captura(
            lambda dax, lb: (F._tablas_dax(token, ws, ds_id, dax, lb) or [[]])[0],
            q["dax"], f"tarjeta:{clave[0]}")
        v_reporte = valor_de(tablas[0] if tablas else [])

        k = next((x for x in (reportes.get(rep_k) or {}).get("kpis", [])
                  if (x.get("label") or "").split(" · ")[0] == etiqueta), None)
        v_app = None
        if k:
            try:
                v_app = float(str(k["valor"]).replace("S/", "").replace(",", "")
                              .replace("%", "").replace("kg", "").replace("KG", "")
                              .replace("d", "").replace("M", "e6").replace("K", "e3")
                              .strip())
            except ValueError:
                v_app = None

        if v_reporte is None or v_app is None:
            sin_dato += 1
            print(f"  ?      {etiqueta}: sin comparación posible")
            continue

        # La app formatea (porcentajes ×100, millones abreviados): se compara
        # en proporción, no en valor absoluto.
        escala = [1, 100, 0.01, 1e6, 1e-6]
        dif = min(abs(v_app - v_reporte * e) / max(abs(v_app), abs(v_reporte * e) or 1)
                  for e in escala)
        if dif <= TOL:
            iguales.append(etiqueta)
            print(f"  OK     {etiqueta}")
        else:
            distintas.append({"kpi": etiqueta, "reporte": rep_k,
                              "en_powerbi": round(v_reporte, 4),
                              "en_juanito": round(v_app, 4),
                              "diferencia_pct": round(dif * 100, 1)})
            print(f"  FALLA  {etiqueta}: Power BI {v_reporte:,.2f} · "
                  f"Juanito {v_app:,.2f}")

    print(f"\n{len(iguales)} iguales · {len(distintas)} distintas · "
          f"{sin_mapa} tarjetas sin correspondencia · {sin_dato} sin datos")

    if anotar:
        datos["validacion_tarjetas"] = {
            "iguales": len(iguales),
            "distintas": distintas,
            "sin_correspondencia": sin_mapa,
        }
        RUTA.write_text(json.dumps(datos, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"→ anotado en {RUTA}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
