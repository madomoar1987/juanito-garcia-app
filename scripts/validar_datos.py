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
    # El signo va después del símbolo de moneda: "S/-331,795". Buscarlo solo
    # al principio convertía un tramo negativo en positivo y desplazaba la
    # suma del aging en el doble de ese tramo.
    primer_digito = next((i for i, ch in enumerate(t) if ch.isdigit()), len(t))
    neg = "-" in t[:primer_digito]
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
    """Valor de un KPI. Ignora el sufijo de período que se le añade al nombre
    ("Planilla Total · acumulado 2026"), para que las comprobaciones sigan
    encontrándolo después de corregir la etiqueta."""
    for k in (rep or {}).get("kpis", []):
        lbl = k.get("label") or ""
        if lbl == etiqueta or lbl.split(" · ")[0] == etiqueta:
            return num(k.get("valor"))
    return None


# Un KPI con serie mensual equivalente: si el número no es el del mes, se
# averigua de qué sí es y se corrige el nombre.
EQUIVALENCIAS = [
    ("productividad", "Producción Total (KG)", "productividad", "Producción (kg)", "kg"),
    ("productividad", "Planilla Total", "productividad", "Planilla (S/)", "soles"),
    ("productividad", "Venta Neta (KG)", "productividad", "Venta (kg)", "kg"),
    ("margen_variable", "Ventas mes", "margen", "Ventas (S/.)", "soles"),
    ("fill_rate", "Facturación", "fill_rate", "Facturación", "soles"),
    ("fill_rate", "Orden de Venta", "fill_rate", "Orden de Venta", "soles"),
]


# Cifras de saldo: valen "hoy", no "el mes". Una cartera vencida o una deuda
# son una foto del momento; el gráfico mensual de lo mismo es el cierre de
# cada mes. Las dos son correctas y distintas, y publicarlas con el mismo
# nombre junto a cifras del mes cerrado hace que se lean como si fueran del
# mes. Se marca cuál es cuál en vez de pedirle a nadie que elija.
SALDOS = [
    ("cuentas_por_cobrar", "Morosidad", "cxc", "% Morosidad", True),
]


def marcar_saldos(rp, ser, cerrado):
    per = ser.get("periodos") or []
    if not per or not cerrado:
        return
    i = per.index(cerrado)
    for rep_k, etiqueta, ds, sk, es_pct in SALDOS:
        arr = (ser.get("datasets", {}).get(ds) or {}).get(sk)
        rep = rp.get(rep_k) or {}
        k = next((x for x in rep.get("kpis", []) if x.get("label") == etiqueta), None)
        if not arr or not k or i >= len(arr) or arr[i] is None:
            continue
        pub = num(k.get("valor"))
        cierre = arr[i] * 100 if es_pct else abs(arr[i])
        if pub is None or abs(pub - cierre) / max(abs(pub), abs(cierre)) <= 0.03:
            continue
        k["label"] = f"{etiqueta} · hoy"
        k["nota_periodo"] = (f"saldo del día. El cierre de {cerrado} fue "
                             f"{cierre:.1f}{'%' if es_pct else ''}.")
        print(f"  MARCA  {etiqueta} → {k['label']}  "
              f"(hoy {pub:.1f} · cierre de {cerrado} {cierre:.1f})")


def corregir_etiquetas(rp, ser, cerrado, inf):
    """Renombra los KPIs que no son del mes, diciendo de qué período son.

    "Producción Total (KG)" publicaba 9,621,234 kg junto a cifras de agosto, y
    resulta ser el acumulado de 2026 — coincide al 0.0%. Un acumulado del año
    presentado al lado del mes hace comparar peras con sandías, y el que mira
    no tiene cómo saberlo.

    La corrección es automática y no adivina: solo renombra cuando el número
    coincide con una de las agregaciones posibles. Si no coincide con ninguna,
    se deja como está y se anota, porque inventarle un período sería peor.
    """
    print("\nDe qué período es realmente cada cifra")
    per = ser.get("periodos") or []
    if not per or not cerrado:
        return
    hasta = per.index(cerrado)
    anio = cerrado[:4]
    for rep_k, etiqueta, ds, sk, _u in EQUIVALENCIAS:
        arr = (ser.get("datasets", {}).get(ds) or {}).get(sk)
        rep = rp.get(rep_k) or {}
        k = next((x for x in rep.get("kpis", []) if x.get("label") == etiqueta), None)
        if not arr or not k:
            continue
        pub = num(k.get("valor"))
        if pub is None:
            continue
        vals = [(per[i], arr[i]) for i in range(min(len(per), len(arr)))
                if arr[i] is not None and i <= hasta]
        if not vals:
            continue
        mes = abs(vals[-1][1])
        ytd = sum(abs(v) for p, v in vals if p.startswith(anio))
        todo = sum(abs(v) for _, v in vals)

        def cerca(x):
            return x and abs(pub - x) / max(abs(pub), abs(x)) <= 0.02

        if cerca(mes):
            print(f"  OK     {etiqueta}: es del mes, como dice")
            inf.ok += 1
        elif cerca(ytd):
            nuevo = f"{etiqueta} · acumulado {anio}"
            print(f"  CORRIGE  {etiqueta} → {nuevo}  (coincide con el acumulado del año)")
            k["label"] = nuevo
            k["nota_periodo"] = (f"no es del mes: es la suma de {anio} hasta "
                                 f"{cerrado}. El mes fue {mes:,.0f}.")
        elif cerca(todo):
            nuevo = f"{etiqueta} · acumulado histórico"
            print(f"  CORRIGE  {etiqueta} → {nuevo}  (coincide con toda la serie)")
            k["label"] = nuevo
            k["nota_periodo"] = (f"no es del mes ni del año: es toda la serie "
                                 f"disponible. El mes fue {mes:,.0f}.")
        else:
            print(f"  ?      {etiqueta}: no coincide con mes, año ni total")
            inf.fallos.append((f"{etiqueta}: período desconocido", pub, mes,
                               abs(pub - mes) / max(pub, mes),
                               f"mes {mes:,.0f} · año {ytd:,.0f} · total {todo:,.0f}"))


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
        # El aging NO incluye la deuda refinanciada: son los vencimientos
        # corrientes. Total = tramos + refinanciado, y cuadra al peso
        # (9,775,627 + 6,450,000 = 16,225,627 = S/16.23M).
        inf.comparar("CxP: total vs tramos + refinanciado",
                     kpi(cxp, "CxP Total"),
                     sum(num(t.get("valor")) or 0 for t in tcxp)
                     + (kpi(cxp, "Refinanciado") or 0),
                     f"{len(tcxp)} tramos del aging más la deuda refinanciada")

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
    # "Planilla por kg vendido" es una medida del reporte con sus propios
    # filtros. Dividir la planilla del año entre unos kilos que solo cuentan
    # la planta ATE no reproduce esa medida: comparaba peras con manzanas y
    # marcaba un 34% de descuadre que no existía.
    print("  (planilla por kg vendido: medida propia del reporte, "
          "no se reconstruye dividiendo dos tarjetas)")

    co = rp.get("consumo_materiales", {})
    # Igual que arriba: [Ratio costo / kg] es una medida del reporte con siete
    # filtros propios. La división cruda de dos tarjetas da otra cosa.
    print("  (costo por tonelada: medida propia del reporte, igual que arriba)")

    # ── La misma magnitud medida en dos reportes ────────────────────────
    print("\nLa misma magnitud, medida en dos reportes distintos")
    # Los kilos de Consumo y los de Productividad NO son comparables, y ya se
    # sabe por qué: las consultas capturadas muestran que la de Productividad
    # filtra Planta = "ATE" y la de Consumo no. Todas las plantas contra una.
    # Compararlos producía un descuadre del 59% que no era un error sino otro
    # alcance, y un aviso que no se puede resolver deja de mirarse.
    print("  (kilos de Consumo vs Productividad: no se comparan — "
          "Productividad filtra solo planta ATE)")
    # Tampoco se comparan las ventas de Margen contra la facturación de Fill
    # Rate. Difieren 21% y durante semanas se avisó como posible descuadre;
    # confirmado con gerencia el 14/09/2026: la venta del mes es la del
    # reporte de Margen (agosto: S/7.69M) y la facturación de Fill Rate mide
    # otra cosa — el cumplimiento del pedido, no la venta.
    print("  (ventas de Margen vs facturación de Fill Rate: no se comparan — "
          "la venta del mes es la de Margen; Fill Rate mide cumplimiento)")


    # ── La tarjeta contra la serie mensual ──────────────────────────────
    # Son dos consultas distintas al mismo reporte: la tarjeta del mes y el
    # gráfico mensual. Tienen que dar lo mismo para el mes cerrado. Si no, una
    # de las dos trae otro filtro.
    print("\nLa tarjeta del mes contra el gráfico mensual (dos consultas)")
    ser = {}
    rutas = ruta.parent / "series.json"
    if rutas.exists():
        try:
            ser = json.loads(rutas.read_text(encoding="utf-8"))
        except Exception:
            ser = {}
    per = ser.get("periodos") or []
    parcial = ser.get("parcial")
    cerrado = next((x for x in reversed(per) if x != parcial), None)
    if cerrado:
        i = per.index(cerrado)
        PARES = [
            ("mermas", "Merma Total", ("mermas", "% Merma Total"), True),
            ("margen_variable", "Margen variable", ("margen", "% Margen Variable"), True),
            ("margen_variable", "Ventas mes", ("margen", "Ventas (S/.)"), False),
            ("margen_variable", "Precio/kg", ("margen", "Precio x Kilo"), False),
            ("margen_variable", "Costo/kg", ("margen", "Costo x Kilo"), False),
            ("cuentas_por_cobrar", "Morosidad", ("cxc", "% Morosidad"), True),
            ("cuentas_por_pagar", "Días CxP", ("cxp", "Días CxP"), False),
            ("fill_rate", "Facturación", ("fill_rate", "Facturación"), False),
            ("fill_rate", "Orden de Venta", ("fill_rate", "Orden de Venta"), False),
            ("consumo_materiales", "Costo x TN Vendida", ("consumo", "MIP / TN Vendida"), False),
            ("productividad", "Planilla S/. / KG Producido",
             ("productividad", "Planilla / kg producido"), False),
        ]
        for rep_k, etiqueta, (ds, sk), es_pct in PARES:
            arr = (ser.get("datasets", {}).get(ds) or {}).get(sk)
            if not arr or i >= len(arr) or arr[i] is None:
                inf.saltados.append(f"{etiqueta}: sin serie para {cerrado}")
                continue
            v_serie = arr[i] * 100 if es_pct else abs(arr[i])
            inf.comparar(f"{etiqueta}: tarjeta vs serie de {cerrado}",
                         kpi(rp.get(rep_k, {}), etiqueta), v_serie,
                         f"{rep_k} · {ds}/{sk}", tol=0.03)
    else:
        inf.saltados.append("no hay series.json para cruzar")

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

    if anotar:
        corregir_etiquetas(rp, ser, cerrado, inf)
        marcar_saldos(rp, ser, cerrado)
        # Un descuadre que acaba de quedar explicado ya no es un descuadre.
        # Si no se retira, el correo avisaría cada día de algo resuelto, y un
        # aviso que siempre suena deja de mirarse.
        explicados = {k["label"].split(" · ")[0]
                      for r in rp.values() for k in r.get("kpis", [])
                      if k.get("nota_periodo")}
        antes = len(inf.fallos)
        inf.fallos = [f for f in inf.fallos
                      if not any(e in f[0] for e in explicados)]
        if antes != len(inf.fallos):
            print(f"\n  {antes - len(inf.fallos)} descuadre(s) quedaron "
                  f"explicados al corregir el período")

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
