#!/usr/bin/env python3
"""Cruza cifras de Power BI entre sí y avisa cuando no cuadran.

No compara contra nada externo: usa solo lo que publican los reportes. La idea
es que muchas cifras aparecen dos veces por caminos distintos —el total y sus
partes, un ratio y sus dos operandos, la misma magnitud en dos reportes— y por
aritmética tienen que coincidir. Cuando no coinciden, una de las dos está mal
o miden cosas distintas con el mismo nombre, y en ambos casos hay que saberlo.

Esto no valida que Power BI tenga razón. Valida que lo que publicamos sea
internamente consistente, que es lo máximo que se puede comprobar sin abrir el
reporte a mano.

Uso:  python scripts/validar_datos.py [ruta/summaries.json]
"""

import json
import pathlib
import re
import sys

RUTA = pathlib.Path("data/latest/summaries.json")
TOL = 0.02          # 2% de diferencia se acepta: redondeos de presentación


def num(v):
    """Convierte 'S/4.27M', '17.8%', '16,750,700 KG' a número."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip().replace("−", "-")
    neg = t.startswith("-")
    # Primero se quitan las unidades escritas ("KG", "TN", "d", "%"), si no
    # "9,621,234 kg" se leía como 9.6 billones: la k de kg pasaba por "miles".
    t = re.sub(r"(?i)\s*(kg|tn|ton|un|d|días|dias)\b.*$", "", t)
    # El sufijo de escala va pegado al número: S/4.27M, 700K.
    mult = 1.0
    m = re.search(r"(?i)([0-9])\s*([MK])\s*$", t)
    if m:
        mult = 1e6 if m.group(2).upper() == "M" else 1e3
        t = t[:m.end(1)]
    t = re.sub(r"[^0-9.]", "", t)
    try:
        x = float(t) * mult
    except ValueError:
        return None
    return -x if neg else x


class Informe:
    def __init__(self):
        self.ok = 0
        self.fallos = []
        self.saltados = []

    def comparar(self, titulo, a, b, detalle, tol=TOL):
        if a is None or b is None:
            self.saltados.append(f"{titulo}: falta un dato")
            return
        base = max(abs(a), abs(b)) or 1
        dif = abs(a - b) / base
        if dif <= tol:
            self.ok += 1
            print(f"  OK     {titulo}  ({detalle})")
        else:
            self.fallos.append((titulo, a, b, dif, detalle))
            print(f"  FALLA  {titulo}")
            print(f"         {detalle}")
            print(f"         difieren en {dif*100:.1f}%")


def kpi(rep, etiqueta):
    for k in (rep or {}).get("kpis", []):
        if k.get("label") == etiqueta:
            return num(k.get("valor"))
    return None


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    anotar = "--anotar" in sys.argv
    ruta = pathlib.Path(args[0]) if args else RUTA
    d = json.loads(ruta.read_text(encoding="utf-8"))
    rp = (d.get("empresas", {}).get("PAUNO", {}) or {}).get("reportes", {})
    inf = Informe()

    print(f"Validación cruzada · datos al {d.get('fecha')}\n")

    # ── Totales contra la suma de sus partes ────────────────────────────
    print("Un total tiene que ser la suma de sus partes")
    cxc = rp.get("cuentas_por_cobrar", {})
    tramos = cxc.get("tramos") or []
    if tramos:
        inf.comparar("CxC: total vs suma de tramos",
                     kpi(cxc, "CxC Total"),
                     sum(num(t.get("valor")) or 0 for t in tramos),
                     f"{len(tramos)} tramos del aging")

    cxp = rp.get("cuentas_por_pagar", {})
    tcxp = cxp.get("tramos") or []
    if tcxp:
        inf.comparar("CxP: total vs suma de tramos",
                     kpi(cxp, "CxP Total"),
                     sum(num(t.get("valor")) or 0 for t in tcxp),
                     f"{len(tcxp)} tramos")

    inv = rp.get("sop_inventario", {})
    partes = [kpi(inv, x) for x in ("Dead Stock", "Working Stock",
                                    "Exceso 1 (2-5 meses)", "Exceso 2 (5-12 meses)")]
    if all(p is not None for p in partes):
        inf.comparar("Inventario: total vs dead + working + excesos",
                     kpi(inv, "Inventario Total"), sum(partes),
                     "las cuatro clasificaciones")

    # ── Un ratio contra sus dos operandos ───────────────────────────────
    print("\nUn ratio tiene que dar lo que dan sus operandos")
    if kpi(cxc, "CxC Total"):
        inf.comparar("CxC: morosidad vs vencido / total",
                     kpi(cxc, "Morosidad"),
                     kpi(cxc, "CxC Vencido") / kpi(cxc, "CxC Total") * 100,
                     f"{kpi(cxc, 'CxC Vencido'):,.0f} / {kpi(cxc, 'CxC Total'):,.0f}")

    if kpi(inv, "Inventario Total"):
        inf.comparar("Inventario: % dead stock vs dead / total",
                     kpi(inv, "% Dead Stock"),
                     kpi(inv, "Dead Stock") / kpi(inv, "Inventario Total") * 100,
                     "declarado contra calculado")

    mg = rp.get("margen_variable", {})
    p, c = kpi(mg, "Precio/kg"), kpi(mg, "Costo/kg")
    if p:
        inf.comparar("Margen: % vs (precio − costo) / precio",
                     kpi(mg, "Margen variable"), (p - c) / p * 100,
                     f"(S/{p:.2f} − S/{c:.2f}) / S/{p:.2f}")

    fr = rp.get("fill_rate", {})
    ov = kpi(fr, "Orden de Venta")
    if ov:
        inf.comparar("Fill rate: % vs facturación / orden de venta",
                     kpi(fr, "Fill Rate"),
                     kpi(fr, "Facturación") / ov * 100,
                     f"{kpi(fr, 'Facturación'):,.0f} / {ov:,.0f}")
        inf.comparar("Fill rate: venta perdida vs orden − facturación",
                     kpi(fr, "Venta Perdida"), ov - kpi(fr, "Facturación"),
                     "la diferencia entre lo pedido y lo facturado")

    pr = rp.get("productividad", {})
    prod_kg = kpi(pr, "Producción Total (KG)")
    if prod_kg:
        inf.comparar("Productividad: planilla por kg producido",
                     kpi(pr, "Planilla S/. / KG Producido"),
                     kpi(pr, "Planilla Total") / prod_kg,
                     f"{kpi(pr, 'Planilla Total'):,.0f} / {prod_kg:,.0f} kg")
    vend_kg = kpi(pr, "Venta Neta (KG)")
    if vend_kg:
        inf.comparar("Productividad: planilla por kg vendido",
                     kpi(pr, "Planilla S/. / KG Vendido"),
                     kpi(pr, "Planilla Total") / vend_kg,
                     f"{kpi(pr, 'Planilla Total'):,.0f} / {vend_kg:,.0f} kg")

    co = rp.get("consumo_materiales", {})
    vn = kpi(co, "Venta Neta (KG)")
    if vn:
        inf.comparar("Consumo: costo por tonelada vendida",
                     kpi(co, "Costo x TN Vendida"),
                     abs(kpi(co, "Costo Total")) / (vn / 1000),
                     f"{abs(kpi(co, 'Costo Total')):,.0f} / {vn/1000:,.0f} TN")

    # ── La misma magnitud medida en dos reportes ────────────────────────
    print("\nLa misma magnitud, medida en dos reportes distintos")
    inf.comparar("Kilos vendidos: Consumo vs Productividad",
                 kpi(co, "Venta Neta (KG)"), kpi(pr, "Venta Neta (KG)"),
                 "los dos reportes dicen 'Venta Neta (KG)'")
    inf.comparar("Kilos producidos: Consumo vs Productividad",
                 kpi(co, "Producción Neta (KG)"), kpi(pr, "Producción Total (KG)"),
                 "producción del mismo mes")
    inf.comparar("Ventas del mes: Margen vs Fill Rate",
                 kpi(mg, "Ventas mes"), kpi(fr, "Facturación"),
                 "facturación del mismo mes por dos caminos")

    # ── Un total no puede superar a la mayor de sus partes ──────────────
    print("\nUn promedio no puede salirse del rango de sus partes")
    for tipo, clave, campo, etiqueta in [
        ("mermas", "por_uen", "merma", "Merma Total"),
        ("mermas", "por_planta", "merma", "Merma Total"),
    ]:
        filas = (rp.get(tipo, {}) or {}).get(clave) or []
        vals = [num(f.get(campo)) for f in filas]
        vals = [v for v in vals if v is not None]
        tot = kpi(rp.get(tipo, {}), etiqueta)
        if vals and tot is not None:
            lo, hi = min(vals), max(vals)
            dentro = lo * 0.95 <= tot <= hi * 1.05
            if dentro:
                inf.ok += 1
                print(f"  OK     {etiqueta} dentro del rango de {clave} "
                      f"({lo:.2f}% a {hi:.2f}%)")
            else:
                inf.fallos.append((f"{etiqueta} fuera del rango de {clave}",
                                   tot, hi, 0, ""))
                print(f"  FALLA  {etiqueta} = {tot:.2f}% fuera del rango de "
                      f"{clave} ({lo:.2f}% a {hi:.2f}%)")

    # ── Resumen ────────────────────────────────────────────────────────
    print(f"\n{'─'*66}")
    print(f"{inf.ok} comprobaciones cuadran · {len(inf.fallos)} no cuadran"
          + (f" · {len(inf.saltados)} sin datos" if inf.saltados else ""))
    if inf.fallos:
        print("\nLo que no cuadra:")
        for t, a, b, dif, det in inf.fallos:
            print(f"  · {t}")
            if dif:
                print(f"      publicado {a:,.2f}  ·  se esperaba {b:,.2f}"
                      f"  ({dif*100:.0f}% de diferencia)")

    # Con --anotar, el resultado viaja dentro de los datos. Así la app puede
    # avisar de una contradicción en la misma pantalla donde está la cifra, en
    # vez de dejarla escondida en el registro de una corrida que nadie abre.
    if anotar:
        d["validacion"] = {
            "cuadran": inf.ok,
            "no_cuadran": [
                {"prueba": t, "publicado": round(a, 4), "esperado": round(b, 4),
                 "diferencia_pct": round(dif * 100, 1), "detalle": det}
                for t, a, b, dif, det in inf.fallos if dif],
            "sin_datos": inf.saltados,
        }
        ruta.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"\n→ resultado anotado en {ruta}")
        # En modo anotar no se corta la corrida: publicar el dato con la
        # contradicción marcada es mejor que no publicar nada. Congelar el
        # tablero por un descuadre ya nos costó cuatro días en setiembre.
        return 0
    return 1 if inf.fallos else 0


if __name__ == "__main__":
    sys.exit(main())
