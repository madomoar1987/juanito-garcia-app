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

    def afirmar(self, ok, titulo, detalle=""):
        """Una comprobación que no compara dos números sino que verifica una
        condición: la serie está alineada, el campo existe."""
        if ok:
            self.ok += 1
            print(f"  OK     {titulo}" + (f"  ({detalle})" if detalle else ""))
        else:
            self.fallos.append((titulo, None, None, None, detalle))
            print(f"  FALLA  {titulo}")
            if detalle:
                print(f"         {detalle}")

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


# Qué serie mensual le corresponde a cada reporte. Un reporte sin serie no se
# puede sellar: su tarjeta se queda sin período confirmado, que es correcto —
# control interno y avance vs presupuesto son acumulados sin corte mensual.
DATASET_DE_REPORTE = {
    "mermas": "mermas",
    "margen_variable": "margen",
    "cuentas_por_cobrar": "cxc",
    "cuentas_por_pagar": "cxp",
    "fill_rate": "fill_rate",
    "consumo_materiales": "consumo",
    "productividad": "productividad",
    "compras": "compras",
    "sop_inventario": "inventario",
}


def escalas_mezcladas(ser):
    """Medidas de la misma familia publicadas en escalas distintas.

    '% Merma Planta Terceros' llegaba en 1.86 mientras '% Merma Planta ATE'
    llegaba en 0.0303: el mismo porcentaje, dos escalas. La app mostraba la
    primera como 185.6% y la ponía a la cabeza de todos los problemas de
    merma. No hay forma de notarlo mirando una sola serie — solo comparando a
    las hermanas entre sí.

    Se agrupan las series por su prefijo (lo que va antes del último espacio)
    y se comprueba que todas estén en el mismo orden de magnitud.
    """
    malas = []
    for ds, series in (ser.get("datasets") or {}).items():
        familias = {}
        for sk, arr in (series or {}).items():
            if not sk.startswith("%") or not isinstance(arr, list):
                continue
            vals = [abs(x) for x in arr if x is not None]
            if not vals:
                continue
            # fracción (<= 1) o porcentaje (> 1): dos mundos distintos.
            familias.setdefault(ds, []).append((sk, max(vals)))
        for fam, items in familias.items():
            frac = [k for k, v in items if v <= 1]
            pct = [k for k, v in items if v > 1]
            if frac and pct:
                malas.append((fam, frac, pct))
    return malas



def series_desalineadas(ser):
    """Series cuyos datos están al principio del eje y no al final.

    Una serie mensual termina en el mes en curso. Si sus únicos puntos están en
    los primeros meses del eje —2025-01 a 2025-04 en un eje que llega a
    2026-09— no es que falten datos recientes: es que los datos se escribieron
    en las posiciones equivocadas. Se detectó porque la tarjeta del reporte,
    que es del mes actual, coincidía con el punto de 2025-04.

    No se corrige sola: mover los puntos sería inventarles un mes. Se avisa
    para que se vuelva a capturar el visual, y mientras tanto esa serie no se
    usa para sellar períodos.
    """
    per = ser.get("periodos") or []
    if len(per) < 8:
        return []
    malas = []
    for ds, series in (ser.get("datasets") or {}).items():
        for sk, arr in (series or {}).items():
            if not isinstance(arr, list) or len(arr) != len(per):
                continue
            idx = [i for i, x in enumerate(arr) if x is not None]
            if not idx or len(idx) > len(per) // 2:
                continue
            # Todos los puntos en el primer tercio del eje y ninguno cerca del
            # final: el eje va al revés de lo que debería.
            if idx[-1] < len(per) // 3:
                malas.append((f"{ds}/{sk}", len(idx), per[idx[0]], per[idx[-1]]))
    return malas



DESALINEADAS = set()


HOY_TXT = ""


def sellar_periodos(rp, ser, parcial=None, cerrado=None):
    """Marca cada KPI con el mes al que pertenece, deducido de su propio valor.

    La pantalla de reportes mostraba once tarjetas bajo el título "este mes" y
    no todas eran del mes: ventas y margen eran de agosto, la mora era el saldo
    de hoy, consumo de materiales era el acumulado del año y el fill rate venía
    de otra medida. Todas correctas, todas con el mismo rótulo.

    En vez de mantener a mano una lista de qué es cada tarjeta —que envejece y
    miente—, se busca el valor en las series y se mira en qué mes aparece. Si
    aparece en uno solo, ese es su período y se sella. Si aparece en varios o
    en ninguno, no se inventa: se deja sin sellar y la app dice que no está
    confirmado, que es la verdad.
    """
    per = ser.get("periodos") or []
    datasets = ser.get("datasets") or {}
    if not per or not datasets:
        return 0, 0

    sellados = sin_sellar = 0
    for rep_k, rep in rp.items():
        if not isinstance(rep, dict):
            continue
        # Solo se busca dentro de la serie del PROPIO reporte. Buscar en todas
        # producía coincidencias de casualidad: el 69.32% de control interno
        # "aparecía" en un mes de 2025 de otra medida cualquiera, y la tarjeta
        # quedaba sellada con un mes que no tiene nada que ver.
        ds_k = DATASET_DE_REPORTE.get(rep_k)
        series_rep = datasets.get(ds_k) if ds_k else None
        if not series_rep:
            # Sin serie mensual del propio reporte no hay contra qué sellar.
            # Tampoco ahí se deja el hueco: se dice que no hay corte mensual
            # y a qué fecha está el dato, en vez de que la app muestre "sin
            # confirmar" y el lector lo lea como si fuera del mes en curso.
            for k in rep.get("kpis", []) or []:
                if not k.get("periodo") and not k.get("nota_periodo"):
                    k["nota_periodo"] = (
                        "este reporte no publica serie mensual, así que no hay "
                        f"con qué confirmar el mes. Dato al {HOY_TXT}.")
            sin_sellar += len(rep.get("kpis", []) or [])
            continue
        for k in rep.get("kpis", []) or []:
            v = num(k.get("valor"))
            if v is None or v == 0:
                continue
            # Se pregunta primero por los dos meses que una tarjeta puede ser:
            # el que corre o el último cerrado. Buscar en los veintiún meses
            # del eje devolvía "coincide con varios" y dejaba sin sellar
            # tarjetas que eran claramente del mes cerrado: un 46.5% de margen
            # se repite en más de un mes del año.
            # La tarjeta viene redondeada para mostrarse, así que la
            # tolerancia sale de cuántos decimales enseña: "2.9%" puede ser
            # cualquier cosa entre 2.85 y 2.95, y "S/0.53" entre 0.525 y
            # 0.535. Con una tolerancia fija en porcentaje, la merma —que se
            # muestra con un decimal sobre un número chico— se quedaba fuera
            # por dos centésimas.
            txt = str(k.get("valor"))
            dec = len(txt.split(".")[1].rstrip("%KGkg ")) if "." in txt else 0
            tol_abs = max(0.5 * (10 ** -dec), abs(v) * 0.002)
            candidatos = [p for p in (parcial, cerrado) if p in per]
            meses, fuentes = set(), set()
            for orden in (candidatos, range(len(per))):
                idxs = ([per.index(p) for p in orden] if orden is candidatos
                        else list(orden))
                for sk, arr in series_rep.items():
                    if not isinstance(arr, list) or f"{ds_k}/{sk}" in DESALINEADAS:
                        continue
                    for i in idxs:
                        x = arr[i] if i < len(arr) else None
                        if x is None or i >= len(per):
                            continue
                        # Los porcentajes viajan como fracción en las series y
                        # como número en las tarjetas: se prueban las dos.
                        for cand in (x, x * 100):
                            if cand and abs(abs(cand) - abs(v)) <= tol_abs:
                                meses.add(per[i])
                                fuentes.add(f"{ds_k}/{sk}")
                                break
                if len(meses) == 1:
                    break            # ya se resolvió con los meses probables
                meses, fuentes = set(), set()
            if len(meses) == 1:
                k["periodo"] = meses.pop()
                if len(fuentes) == 1:
                    k["periodo_fuente"] = fuentes.pop()
                sellados += 1
            else:
                # No coincide con ningún mes del eje de su propio reporte.
                # Eso no es "no se sabe": es que la cifra no es mensual — un
                # saldo vivo, un acumulado del año, un conteo. Dejarlo en
                # blanco obligaba a la app a decir "sin confirmar" y al lector
                # a suponer que era del mes en curso, que es justo lo que no
                # se quiere. Se dice lo que sí se sabe.
                if not k.get("nota_periodo"):
                    k["nota_periodo"] = (
                        "no corresponde a ningún mes del eje: es un saldo al "
                        f"corte o un acumulado. Dato al {HOY_TXT}."
                        if len(meses) == 0 else
                        "el valor aparece en más de un mes del eje, así que no "
                        f"se puede sellar sin suponer. Dato al {HOY_TXT}.")
                sin_sellar += 1
    return sellados, sin_sellar



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
        saldos = {(r, e) for r, e, _, _, _ in SALDOS}
        for rep_k, etiqueta, (ds, sk), es_pct in PARES:
            arr = (ser.get("datasets", {}).get(ds) or {}).get(sk)
            # Una cifra de saldo vale hoy, no el mes cerrado: la tarjeta de
            # morosidad es la cartera de este momento y su punto en la serie
            # es el del mes en curso. Compararla contra el cierre anterior
            # marcaba un descuadre del 21% que no existía — son dos fechas.
            es_saldo = (rep_k, etiqueta) in saldos
            mes = parcial if (es_saldo and parcial in per) else cerrado
            j = per.index(mes)
            if not arr or j >= len(arr) or arr[j] is None:
                inf.saltados.append(f"{etiqueta}: sin serie para {mes}")
                continue
            v_serie = arr[j] * 100 if es_pct else abs(arr[j])
            inf.comparar(f"{etiqueta}: tarjeta vs serie de {mes}" +
                         (" (saldo del día)" if es_saldo else ""),
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
        # Solo tiene sentido si ambos miden el mismo período. Los segmentos de
        # merma vienen acumulados del año y la tarjeta es de un mes: un mal
        # mes puede superar a cualquier promedio anual sin que nada esté mal.
        # La prueba pasaba por casualidad, no por estar bien planteada.
        per_seg = (rp.get(tipo, {}) or {}).get("por_uen_periodo") or ""
        per_kpi = ""
        for k in (rp.get(tipo, {}) or {}).get("kpis", []):
            if k.get("label") == etiqueta:
                per_kpi = k.get("periodo") or ""
        if vals and tot is not None and per_seg and per_kpi != per_seg:
            print(f"  ·      {etiqueta} no se compara con {clave}: "
                  f"la tarjeta es de {per_kpi or 'período sin declarar'} y los "
                  f"segmentos son {per_seg}")
        elif vals and tot is not None:
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

    # Un porcentaje solo significa algo si su base es positiva. Una venta
    # neta negativa —más devoluciones que ventas— es un dato correcto; el
    # margen calculado sobre ella no lo es: cambia de signo y se lee al revés.
    # La nota de crédito llegó a publicarse con "53.5% de margen" sobre
    # -S/158,758, y tres SKU de B&D y MAQUILA con margen positivo sobre venta
    # negativa. Ninguna prueba lo miraba porque cada cifra, por separado,
    # parecía razonable.
    print("\nNingún porcentaje calculado sobre una base negativa")
    def _plata(x):
        if isinstance(x, (int, float)): return float(x)
        if isinstance(x, str):
            try: return float(x.replace("S/", "").replace(",", "").strip())
            except ValueError: return None
        return None

    malos = []
    def _mirar(o, ruta=""):
        if isinstance(o, dict):
            v = _plata(o.get("venta")) if "venta" in o else None
            if v is not None and v <= 0:
                for campo in ("margen", "margen_previo", "margen_enero"):
                    if o.get(campo):
                        malos.append((ruta, o.get("producto") or o.get("cliente")
                                      or "?", campo, o.get("venta"), o[campo]))
            for k, val in o.items():
                _mirar(val, f"{ruta}/{k}")
        elif isinstance(o, list):
            for i, val in enumerate(o):
                _mirar(val, f"{ruta}[{i}]")
    _mirar(rp)
    if malos:
        for ruta, quien, campo, venta, val in malos[:6]:
            inf.afirmar(False, f"{quien}: {campo} sobre una venta negativa",
                        f"venta {venta} y {campo} {val} — un porcentaje sobre "
                        f"una base negativa se lee al revés; publicar la "
                        f"devolución sin margen es lo correcto")
        if len(malos) > 6:
            print(f"         …y {len(malos) - 6} más")
    else:
        inf.ok += 1
        print("  OK     ninguna fila publica margen sobre venta menor o igual a cero")

    # Un precio por kilo no se multiplica por ocho en un mes. Cuando pasa, es
    # el peso mal registrado, y como el margen se deriva del precio, la fila
    # se va arriba del ranking de mejoras con una cifra que no existe.
    print("\nNingún precio por kilo con un salto imposible")
    saltos = []
    def _precios(o):
        if isinstance(o, dict):
            a, b = num(o.get("precio_kg_previo")), num(o.get("precio_kg"))
            if a and b and a > 0 and b > 0:
                x = max(b / a, a / b)
                if x > 3:
                    saltos.append((o.get("producto") or "?", a, b, x))
            for v in o.values(): _precios(v)
        elif isinstance(o, list):
            for v in o: _precios(v)
    _precios(rp)
    if saltos:
        for quien, a, b, x in saltos[:5]:
            inf.afirmar(False, f"{quien}: precio por kilo x{x:.1f} en un mes",
                        f"de S/{a:.2f} a S/{b:.2f} — un precio no se mueve así; "
                        f"lo más probable es que el peso esté mal registrado, y "
                        f"el margen derivado de ese precio tampoco vale")
    else:
        inf.ok += 1
        print("  OK     ningún precio por kilo se multiplica o divide por más de tres")

    # Los metadatos del resumen —cobertura, derivados, cifras leídas de la
    # tarjeta— se escribían dentro de "si hubo consultas con problema". Con la
    # corrida limpia del 16/09 no se escribió ninguno, y sin kpis_de_tarjeta
    # la serie mensual pisó el valor de la tarjeta: se publicó Costo x TN
    # Producida en 151.64 cuando Power BI muestra 143.28. Cuanto mejor salía
    # la corrida, peor el dato. Que falten es señal de que volvió a pasar.
    print("\nEl resumen trae sus metadatos")
    faltan = [k for k in ("cobertura", "derivados", "kpis_de_tarjeta")
              if not (d.get(k))]
    if faltan:
        inf.afirmar(False, "faltan metadatos en summaries.json",
                    f"no vienen {', '.join(faltan)} — sin ellos la app no sabe "
                    f"qué cifra es derivada ni cuál viene de la tarjeta, y la "
                    f"serie mensual puede pisar el valor que muestra Power BI")
    else:
        inf.ok += 1
        print("  OK     cobertura, derivados y cifras de tarjeta presentes")

    # Un desglose tiene que sumar lo que dice su propia tarjeta, o declarar
    # por qué no. La tabla de clientes sumaba S/4.09M bajo un KPI de S/7.69M
    # y nada lo advertía: son órdenes de venta contra facturación, y setiembre
    # contra agosto. Dos diferencias legítimas, invisibles las dos.
    #
    # Las excepciones se declaran aquí con su motivo. Lo que no está en la
    # lista tiene que cuadrar al 3%.
    print("\nCada desglose suma lo que dice su tarjeta")
    CUADRAN = [
        ("cuentas_por_cobrar", "tramos", "valor", "CxC Total", None),
        ("cuentas_por_cobrar", "responsables", "total", "CxC Total", None),
        ("sop_inventario", "dead_por_categoria", "valor", "Dead Stock", None),
        ("sop_inventario", "working_por_categoria", "valor", "Working Stock", None),
        ("margen_variable_pag2", "por_canal", "facturado", "Facturado", None),
        ("cuentas_por_pagar", "tramos", "valor", "CxP Total",
         "el aging son los vencimientos corrientes; el total incluye además "
         "la deuda refinanciada, que no vence por tramos"),
        ("margen_variable", "por_cliente", "venta", "Ventas mes",
         "son ORDENES DE VENTA del mes en curso y el KPI es FACTURACION del "
         "mes cerrado: dos bases y dos periodos distintos"),
        ("margen_variable", "por_documento", "venta", "Ventas mes",
         "es la facturacion del mes EN CURSO a la fecha y el KPI es el mes "
         "cerrado completo"),
    ]

    def _kpi_val(rep, etiqueta):
        for k in rep.get("kpis", []) or []:
            if (k.get("label") or "").lower().startswith(etiqueta.lower()):
                return num(k.get("valor"))
        return None

    for rep_k, desg, campo, etiqueta, motivo in CUADRAN:
        rep = rp.get(rep_k) or {}
        filas = rep.get(desg)
        if not isinstance(filas, list) or not filas:
            continue
        vals = [num(f.get(campo)) for f in filas if isinstance(f, dict)]
        vals = [v for v in vals if v is not None]
        tot = _kpi_val(rep, etiqueta)
        if not vals or not tot:
            continue
        suma = sum(vals)
        ratio = suma / tot
        cuadra = 0.97 <= ratio <= 1.03
        if motivo:
            # Declarado: lo que se vigila es que SIGA sin cuadrar. Si un día
            # cuadra, la excepción sobra y hay que quitarla.
            if cuadra:
                inf.afirmar(False, f"{rep_k}/{desg}: la excepción ya no aplica",
                            f"suma {suma:,.0f} y {etiqueta} {tot:,.0f} ahora "
                            f"coinciden; sobra la excepción «{motivo}»")
            else:
                inf.ok += 1
                # El motivo viaja DENTRO del dato: si solo vive aquí, la app
                # muestra una tabla que no suma su tarjeta y nadie sabe por qué.
                rep.setdefault("_no_cuadran", {})[desg] = {
                    "kpi": etiqueta, "suma": round(suma), "kpi_valor": round(tot),
                    "motivo": motivo}
                print(f"  OK     {rep_k}/{desg} no cuadra con {etiqueta} "
                      f"({ratio:.2f}) y está explicado")
        else:
            inf.afirmar(cuadra, f"{rep_k}/{desg} vs {etiqueta}",
                        f"suma {suma:,.0f} contra {tot:,.0f} ({ratio:.2f}): un "
                        f"desglose que no suma lo que dice su tarjeta se lee "
                        f"como si le faltaran filas")

    print("\nPorcentajes de la misma familia en la misma escala")
    mezcla = escalas_mezcladas(ser)
    if mezcla:
        for fam, frac, pct in mezcla:
            inf.afirmar(False, f"{fam}: porcentajes en dos escalas",
                        f"en fracción {frac} · en porcentaje {pct} — publicadas "
                        f"juntas, la app muestra unas cien veces más grandes que "
                        f"las otras")
    else:
        inf.afirmar(True, "todas las familias de porcentajes en una sola escala")

    malas = series_desalineadas(ser)
    print("\nSeries escritas en el tramo equivocado del eje")
    if malas:
        DESALINEADAS.update(m[0] for m in malas)
        for nombre, n, ini, fin in malas:
            inf.afirmar(False, f"{nombre}: eje al revés",
                        f"sus {n} puntos van de {ini} a {fin}, pero el eje "
                        f"llega a {per[-1]} — hay que volver a capturar "
                        f"ese visual")
    else:
        inf.afirmar(True, "ninguna serie desalineada")

    if anotar:
        corregir_etiquetas(rp, ser, cerrado, inf)
        marcar_saldos(rp, ser, cerrado)
        globals()["HOY_TXT"] = str(d.get("fecha") or "")
        a, b = sellar_periodos(rp, ser, parcial, cerrado)
        print(f"\nPeríodo de cada tarjeta: {a} selladas con su mes, "
              f"{b} sin confirmar")
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
            if dif is None and det:
                print(f"      {det}")
            if dif:
                print(f"      publicado {a:,.2f}  ·  se esperaba {b:,.2f}"
                      f"  ({dif*100:.0f}% de diferencia)")

    # Con --anotar, el resultado viaja dentro de los datos. Así la app puede
    # avisar de una contradicción en la misma pantalla donde está la cifra, en
    # vez de dejarla escondida en el registro de una corrida que nadie abre.
    if anotar:
        # Las diferencias ya investigadas se marcan aquí, en un solo sitio, y
        # viajan marcadas dentro del dato. Antes solo las conocía el paso de
        # verificación del workflow, así que la app pintaba en ámbar una
        # diferencia explicada —el peso mal registrado de un SKU— sin forma de
        # dejar de pintarla. Una alarma que suena siempre deja de mirarse.
        conocidos = set()
        try:
            ruta_con = pathlib.Path("data/descuadres_conocidos.json")
            if ruta_con.exists():
                conocidos = {x["prueba"] for x in json.loads(
                    ruta_con.read_text(encoding="utf-8"))["conocidos"]}
        except Exception as e:
            print(f"  · no se pudo leer descuadres_conocidos.json: {e}")

        def _llano(x):
            # Los nombres de producto traen espacios dobles del maestro; sin
            # normalizarlos el emparejamiento fallaba en silencio.
            return re.sub(r"\s+", " ", str(x)).strip().lower()

        conocidos_llanos = {_llano(c) for c in conocidos}

        def _conocido(titulo):
            # Coincide por prefijo: el título lleva el nombre del SKU y la
            # entrada conocida describe el caso, no la corrida.
            t = _llano(titulo)
            return any(t.startswith(c) or c.startswith(t)
                       for c in conocidos_llanos)

        d["validacion"] = {
            "cuadran": inf.ok,
            # OJO con el "if dif" que había aquí.
            #
            # Descartaba toda comprobación que no compare dos números —las de
            # afirmar(): serie desalineada, margen sobre base negativa, precio
            # imposible, metadatos ausentes—. Esas se imprimían en el log del
            # workflow y NO llegaban nunca a los datos, así que la app no las
            # mostraba y el paso de verificación posterior leía "0 no cuadran"
            # con descuadres reales encima de la mesa. Una comprobación que no
            # puede fallar a la vista no es una comprobación.
            "no_cuadran": [
                ({"prueba": t, "publicado": round(a, 4), "esperado": round(b, 4),
                  "diferencia_pct": round(dif * 100, 1), "detalle": det}
                 if dif else
                 {"prueba": t, "publicado": None, "esperado": None,
                  "diferencia_pct": None, "detalle": det})
                | {"conocido": _conocido(t)}
                for t, a, b, dif, det in inf.fallos],
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
