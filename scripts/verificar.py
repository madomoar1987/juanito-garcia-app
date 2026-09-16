#!/usr/bin/env python3
"""
JUANITO — Verificación estática, sin tocar Power BI.

Nació porque tres corridas seguidas del workflow devolvieron datos vacíos por
un error que era detectable sin salir de la máquina: un bloque de código que
leía `found["__top_proveedores"]` había quedado dentro de build_compras en vez
de build_cxp, así que nadie leía lo que la consulta sí estaba trayendo. No
había error que registrar, solo un dato que se perdía en el camino.

Correr esto ANTES de pedir una corrida del workflow. Verifica:

  1. Que cada clave interna "__algo" que alguien escribe, alguien la lea.
  2. Que cada dax_* y serie_* definida se llame desde algún sitio.
  3. Que las funciones de construcción produzcan los KPIs esperados.
  4. Que los nombres de campo que emite el script sean los que lee juanito.html.

Uso:  python3 scripts/verificar.py
"""

import importlib.util
import os
import ast
import builtins
import json
from collections import Counter
import collections
import re
import sys

RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FALLOS = []


def revisar(ok, mensaje):
    print(f"  {'OK ' if ok else 'FALLA'}  {mensaje}")
    if not ok:
        FALLOS.append(mensaje)


def cargar_fetch_powerbi():
    """Importa el script con credenciales falsas: solo se usan en get_token()."""
    for var in ("AZURE_TENANT_ID", "AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET",
                "PBI_USERNAME", "PBI_PASSWORD"):
        os.environ.setdefault(var, "verificacion")
    ruta = os.path.join(RAIZ, "scripts", "fetch_powerbi.py")
    spec = importlib.util.spec_from_file_location("fp", ruta)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def leer(*partes):
    with open(os.path.join(RAIZ, *partes), encoding="utf-8") as fh:
        return fh.read()


def funcion_contenedora(lineas, idx):
    for j in range(idx, -1, -1):
        if lineas[j].startswith("def "):
            return lineas[j][4:lineas[j].index("(")]
    return "?"


def test_claves_internas(src):
    """Toda clave '__x' escrita debe leerse, y viceversa."""
    print("\n1. Claves internas escritas vs leídas")
    lineas = src.split("\n")
    escrituras, lecturas = {}, {}
    for i, linea in enumerate(lineas):
        for m in re.finditer(r'\["(__[a-z_]+)"\]\s*=', linea):
            escrituras.setdefault(m.group(1), []).append(funcion_contenedora(lineas, i))
        # Escritura indirecta: un ayudante que recibe la clave como argumento
        # y la guarda él. `extraer_dimensiones` trae seis cortes con la misma
        # mecánica y repetir seis veces la asignación literal solo para que
        # este test la viera sería peor código.
        for m in re.finditer(r'^\s*"(__[a-z_]+)",\s', linea):
            escrituras.setdefault(m.group(1), []).append(funcion_contenedora(lineas, i))
        for m in re.finditer(r'found\.get\("(__[a-z_]+)"\)', linea):
            lecturas.setdefault(m.group(1), []).append(funcion_contenedora(lineas, i))
        # Lectura indirecta: el ayudante recibe la clave y hace el found.get.
        for m in re.finditer(r'_diaria\("(__[a-z_]+)"', linea):
            lecturas.setdefault(m.group(1), []).append(funcion_contenedora(lineas, i))
    for clave in sorted(set(escrituras) | set(lecturas)):
        e, l = escrituras.get(clave, []), lecturas.get(clave, [])
        revisar(bool(e) and bool(l),
                f"{clave}: escribe {e or 'NADIE'} — lee {l or 'NADIE'}")


def test_orden_carga_build(src):
    """Cada bloque que trae datos debe correr ANTES de construir su reporte.

    El margen por cliente salió vacío tres corridas seguidas por esto: la
    consulta funcionaba y guardaba el resultado en `scanned`, pero
    build_margen() ya se había ejecutado veinte líneas antes. No hay error ni
    diagnóstico posible — el dato llega tarde y nadie lo lee.

    Los pares se deducen del código en vez de mantenerse a mano: la lista
    escrita a mano dejó pasar exactamente el mismo fallo con tres desgloses
    nuevos, porque nadie se acordó de añadirlos.
    """
    print("\n5. Orden: la carga de datos antes de construir el reporte")
    lineas = src.split("\n")

    # Dónde se guarda cada dato: scanned.setdefault("cxc", {})["__aging"] = …
    cargas = {}
    for i, l in enumerate(lineas):
        m = re.search(r'scanned(?:\.setdefault\(|\[)"([a-z_]+)".*?\["(__[a-z_]+)"\]\s*=', l)
        if m:
            cargas.setdefault(m.group(1), []).append((m.group(2), i + 1))

    # Qué constructor lee cada clave. Un mismo grupo de `scanned` puede
    # alimentar a dos constructores distintos (inventario alimenta a
    # build_inventario y a build_avance), así que la comparación tiene que ser
    # contra el que de verdad lee esa clave, no contra el primero que aparezca.
    lector = {}
    for i, l in enumerate(lineas):
        m = re.search(r'found\.get\("(__[a-z_]+)"\)', l)
        if m:
            lector[m.group(1)] = funcion_contenedora(lineas, i)

    # Dónde se llama cada constructor: build_x(scanned["cxc"])
    llamada = {}
    for i, l in enumerate(lineas):
        m = re.search(r'(build_[a-z_]+)\(scanned\["([a-z_]+)"\]', l)
        if m and m.group(1) not in llamada:
            llamada[m.group(1)] = i + 1

    if not cargas:
        revisar(False, "no se detectó ninguna carga en `scanned`")
        return
    for grupo, items in sorted(cargas.items()):
        for clave, la in items:
            fn = lector.get(clave)
            if not fn:
                revisar(False, f"{clave}: se carga en L{la} y ningún build lo lee")
                continue
            lb = llamada.get(fn)
            if lb is None:
                revisar(False, f"{clave}: lo lee {fn}(), que nadie llama con scanned[...]")
                continue
            revisar(la < lb, f"{clave} (L{la}) → {fn}() en L{lb}"
                             + ("" if la < lb else "  ← LLEGA TARDE"))


def test_funciones_usadas(src, prefijo, archivo):
    print(f"\n2. Funciones {prefijo}* de {archivo} que nadie llama")
    for nombre in re.findall(rf"^def ({prefijo}[a-z0-9_]+)\(", src, re.M):
        llamadas = len(re.findall(rf"\b{nombre}\(", src)) - 1
        revisar(llamadas > 0, f"{nombre} ({llamadas} llamada/s)")


def test_construccion(fp):
    """Cada build_* con datos sintéticos debe producir sus KPIs y estructuras."""
    print("\n3. Funciones de construcción con datos sintéticos")

    casos = [
        ("build_cxc",
         lambda: fp.build_cxc({
             "% Morosidad": 0.2108,
             "__aging": [("1. Por Vencer", 3570000.0), ("2. 0 a 15 días", 226121.0),
                         ("3. 16 a 30 días", 277775.0), ("4. Más de 30 días", 442248.0)]}),
         ["Morosidad", "CxC por vencer", "CxC Total", "CxC Vencido"], ["tramos"]),
        ("build_cxp",
         lambda: fp.build_cxp({
             "Cuentas x Pagar": 16810000, "Refinanciamiento": 6960000, "Proveedores": 139,
             "__top_proveedores": [("E & M S.R.L.", "LECHE", 5175254.0)],
             "__aging_cxp": [("Vigente", 6467432.0), ("0 a 7 días", 600000.0),
                             ("8 a 15 días", 454715.0), ("31 a 90 días", 1300005.0)]}),
         ["CxP Total", "Refinanciado", "# Proveedores", "Deuda top 15 proveedores",
          "CxP Vencido"],
         ["proveedores_criticos", "tramos"]),
        ("build_fill_rate",
         lambda: fp.build_fill_rate({
             "__fillrate_card": 0.884, "__venta_mes": 1720.0, "__no_atendido_total": 199.0,
             "__por_grupo": [("SPSA", 0.925)], "__soles_por_grupo": {"SPSA": 129.0},
             "__por_marca": [("MAQUILA", 118.0)]}),
         ["Fill Rate", "Venta del mes (tarjeta)", "Pedidos no atendidos (mes)"],
         ["por_grupo", "por_marca"]),
        ("build_inventario",
         lambda: fp.build_inventario({"__clasificacion": [
             ("ENVASES Y EMBALAJES", "1. Working", 1900000.0, None),
             ("ENVASES Y EMBALAJES", "4. Dead", 917801.0, None),
             ("PT SALSAS", "2. Exceso 1", 714044.0, None)]}),
         ["Inventario Total", "Dead Stock", "% Dead Stock", "Working Stock"],
         ["dead_por_categoria", "working_por_categoria"]),
    ]
    salidas = {}
    for nombre, fn, kpis_esperados, estructuras in casos:
        res = fn()
        if isinstance(res, tuple):
            res = res[0]
        salidas.update({k: v for k, v in res.items() if k in estructuras})
        etiquetas = [k["label"] for k in res.get("kpis", [])]
        faltan = [k for k in kpis_esperados if k not in etiquetas]
        revisar(not faltan, f"{nombre}: KPIs {etiquetas}"
                            + (f" — FALTAN {faltan}" if faltan else ""))
        for est in estructuras:
            revisar(bool(res.get(est)), f"{nombre}: produce '{est}'")
    return salidas


def test_campos(salidas, html):
    """Los campos emitidos deben ser los que juanito.html lee."""
    print("\n4. Campos emitidos vs campos que lee la app")
    variable = {"tramos": "t", "proveedores_criticos": "p", "dead_por_categoria": "c",
                "working_por_categoria": "c", "por_grupo": "g", "por_marca": "m"}
    ignorar = {"map", "join", "length", "filter", "slice", "forEach"}
    for estructura, var in variable.items():
        filas = salidas.get(estructura)
        if not filas:
            revisar(False, f"{estructura}: no se produjo, no se puede comparar")
            continue
        emitidos = set(filas[0].keys())
        pos = html.find(f"?.{estructura} || []")
        if pos < 0:
            pos = html.find(f".{estructura} || []")
        bloque = html[pos:pos + 1500] if pos >= 0 else ""
        leidos = {u for u in re.findall(rf"\b{var}\.([a-zA-Z_]+)\b", bloque)} - ignorar
        huerfanos = sorted(leidos - emitidos)
        revisar(not huerfanos,
                f"{estructura}: la app lee {sorted(leidos)}"
                + (f" — NO EXISTEN {huerfanos}" if huerfanos else ""))


def test_derivados_declarados(src):
    """Toda cifra calculada aquí debe estar declarada con anotar_derivado().

    El margen por unidad de negocio se publicó semanas como si viniera del
    reporte: se calculaba (precio − costo) / precio y daba tres puntos más que
    Power BI. Nadie podía notarlo mirando la app, y el código tampoco lo decía.

    Esto busca las divisiones que producen un porcentaje sobre datos del
    reporte y exige que cada una esté declarada. No detecta todo cálculo
    posible —eso sería resolver el problema de la parada— pero sí la forma
    exacta que ya se nos coló dos veces.
    """
    print("\n6. Cifras calculadas: todas declaradas")
    lineas = src.split("\n")
    declarados = set()
    for m in re.finditer(r'anotar_derivado\(\s*\n?\s*"([a-z_0-9]+)",\s*"([a-z_0-9]+)",\s*"([a-z_0-9]+)"', src):
        declarados.add(m.groups())

    # Patrón: se asigna un campo de un desglose a partir de una división.
    sospechas = []
    for i, l in enumerate(lineas):
        m = re.search(r'"([a-z_0-9]+)":\s*\(?\s*round\(\s*\w+\[?[^)]*\]?\s*/\s*', l)
        if not m:
            continue
        campo = m.group(1)
        # ¿A qué desglose pertenece? Se busca hacia atrás el res["x"] = [ o similar.
        bloque = "\n".join(lineas[max(0, i - 14):i])
        dm = re.findall(r'res\["([a-z_0-9]+)"\]\s*=', bloque)
        # Algunos desgloses se arman dentro de una función auxiliar y no hay un
        # res["x"] = delante. En ese caso vale una declaración cercana.
        cerca = re.findall(r'anotar_derivado\(\s*\n?\s*"[a-z_0-9]+",\s*\n?\s*'
                           r'(?:clave|"[a-z_0-9]+")', bloque)
        desglose = dm[-1] if dm else ("(declarado cerca)" if cerca else None)
        # Sin desglose identificable y sin declaración cerca, el cálculo no es
        # de una cifra del reporte: es metadato nuestro (la cobertura de la
        # validación, por ejemplo). Este test vigila los datos publicados, no
        # las estadísticas sobre ellos.
        if desglose is None:
            continue
        sospechas.append((campo, desglose, i + 1))

    if not sospechas:
        revisar(True, "no se detectaron divisiones sin declarar")
    for campo, desglose, ln in sospechas:
        ok = (desglose == "(declarado cerca)"
              or any(d[1] == desglose and d[2] == campo for d in declarados))
        revisar(ok, f"L{ln}: {desglose}.{campo} "
                    + ("declarado" if ok else "CALCULA y no está declarado"))
    revisar(len(declarados) > 0,
            f"{len(declarados)} cifra(s) declaradas como calculadas")


def test_sin_nombres_repetidos(src, archivo):
    """Ninguna función puede estar definida dos veces en el mismo archivo.

    Python no avisa: la segunda definición pisa a la primera en silencio. En
    fetch_powerbi.py había dos valor_de_tarjeta() con firmas distintas, y las
    cinco lecturas de Consumo morían con "takes 4 positional arguments but 5
    were given". Como cada una va en su propio try, el fallo no se veía en
    ningún sitio salvo en el diagnóstico del JSON publicado.

    En un archivo de más de cuatro mil líneas esto no se detecta leyendo.
    """
    print(f"\n7. Nombres repetidos en {archivo}")
    tree = ast.parse(src)
    nombres = collections.Counter(
        n.name for n in tree.body if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef)))
    repes = {k: v for k, v in nombres.items() if v > 1}
    revisar(not repes,
            f"{len(nombres)} funciones, "
            + ("ninguna repetida" if not repes
               else f"REPETIDAS: {', '.join(f'{k} (×{v})' for k, v in repes.items())}"))


def test_variable_antes_de_asignar(src, archivo):
    """Variables que se usan antes de la línea donde se asignan.

    Python no avisa de esto al importar: falla recién cuando la ejecución llega
    a esa línea. En una corrida nocturna eso significa perder el dato y verlo un
    día después. Pasó con `margen_ds_id`, que se asignaba cien líneas más abajo
    de donde el detalle de costos ya lo usaba, y se llevó por delante las tres
    consultas de detalle.

    Solo mira variables locales de cada función: si un nombre se asigna en
    alguna parte de la función, deja de ser global, y usarlo antes de su primera
    asignación es un error aunque el módulo tenga una global con ese nombre.
    """
    print(f"\n8. Variables usadas antes de asignarse ({archivo})")
    try:
        arbol = ast.parse(src)
    except SyntaxError as e:
        revisar(False, f"{archivo} no compila: {e}")
        return

    for fn in [n for n in ast.walk(arbol)
               if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))]:
        # Los nombres declarados global/nonlocal no son locales de la función.
        externos = set()
        for n in ast.walk(fn):
            if isinstance(n, (ast.Global, ast.Nonlocal)):
                externos.update(n.names)
        params = {a.arg for a in
                  list(fn.args.args) + list(fn.args.posonlyargs) +
                  list(fn.args.kwonlyargs)}
        for a in (fn.args.vararg, fn.args.kwarg):
            if a:
                params.add(a.arg)

        # Solo el cuerpo directo de la función. Una comprensión o un lambda
        # tienen su propio ámbito y su variable de bucle se lee antes de
        # escribirse en el texto ("[f(x) for x in xs]"), que es correcto:
        # mirarlas daría decenas de falsas alarmas y el test se ignoraría.
        ANIDADOS = (ast.Lambda, ast.ListComp, ast.SetComp, ast.DictComp,
                    ast.GeneratorExp, ast.FunctionDef, ast.AsyncFunctionDef,
                    ast.ClassDef)
        asigna, usa = {}, {}
        pila = list(ast.iter_child_nodes(fn))
        while pila:
            n = pila.pop()
            if isinstance(n, ANIDADOS):
                # Un `def` anidado no se recorre, pero sí ata su nombre: la
                # función auxiliar existe desde su línea de definición.
                nom = getattr(n, "name", None)
                if nom:
                    asigna[nom] = min(asigna.get(nom, n.lineno), n.lineno)
                continue
            pila.extend(ast.iter_child_nodes(n))
            if isinstance(n, ast.Name):
                d = asigna if isinstance(n.ctx, ast.Store) else usa
                d[n.id] = min(d.get(n.id, n.lineno), n.lineno)
            elif isinstance(n, ast.ExceptHandler) and n.name:
                asigna[n.name] = min(asigna.get(n.name, n.lineno), n.lineno)

        malas = []
        for nombre, linea_uso in usa.items():
            if nombre in params or nombre in externos or nombre not in asigna:
                continue
            if linea_uso < asigna[nombre]:
                malas.append((nombre, linea_uso, asigna[nombre]))

        if malas:
            for nombre, uso, asg in sorted(malas, key=lambda x: x[1]):
                revisar(False, f"{archivo}: en {fn.name}(), '{nombre}' se usa en la "
                               f"línea {uso} pero se asigna recién en la {asg}")
        else:
            revisar(True, f"{fn.name}()")



def test_js_sin_nombres_repetidos(html):
    """Funciones definidas dos veces en juanito.html.

    En JavaScript la segunda definición gana en silencio: la app sigue
    funcionando, pero ejecuta la versión vieja. Pasó al restaurar un bloque
    borrado por error — el arreglo que se acababa de hacer dejó de aplicarse y
    la pantalla mostraba lo de antes sin ningún aviso.

    El equivalente en Python ya se revisa; faltaba el del navegador, que es
    donde el fallo es mudo.
    """
    print("\n9. Funciones repetidas en juanito.html")
    nombres = re.findall(r"^function\s+([A-Za-z_$][\w$]*)\s*\(", html, re.M)
    repes = {n: c for n, c in Counter(nombres).items() if c > 1}
    revisar(not repes,
            f"{len(nombres)} funciones, ninguna repetida" if not repes
            else "REPETIDAS: " + ", ".join(f"{n} (×{c})" for n, c in repes.items()))



def test_nombres_sin_definir(src, archivo):
    """Nombres que el módulo usa y nadie define. Son NameError esperando.

    _COMPRAS_FILTROS se usaba en las dos consultas del ratio consumo/compra y
    no estaba definido en ninguna parte del archivo. Cada llamada moría con
    NameError, el try/except de arriba lo imprimía y seguía, y el dataset
    'compras' no llegaba a crearse nunca. El KPI salió vacío en la app durante
    semanas con el reporte lleno en Power BI, y ninguna prueba lo miraba.

    Se comparan los nombres LEÍDOS a nivel de módulo contra los definidos —
    asignaciones, funciones, clases, imports, parámetros y locales — más los
    builtins. Lo que sobra no existe.
    """
    try:
        arbol = ast.parse(src)
    except SyntaxError as e:
        revisar(False, f"{archivo}: no se puede analizar ({e})")
        return

    definidos = set(dir(builtins))
    for n in ast.walk(arbol):
        if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            definidos.add(n.name)
            for a in list(getattr(n, "args", ast.arguments(
                    posonlyargs=[], args=[], kwonlyargs=[], kw_defaults=[], defaults=[])).args or []):
                definidos.add(a.arg)
            for extra in ("vararg", "kwarg"):
                v = getattr(getattr(n, "args", None), extra, None)
                if v:
                    definidos.add(v.arg)
            for a in getattr(getattr(n, "args", None), "kwonlyargs", []) or []:
                definidos.add(a.arg)
        elif isinstance(n, (ast.Import, ast.ImportFrom)):
            for a in n.names:
                definidos.add((a.asname or a.name).split(".")[0])
        elif isinstance(n, ast.Name) and isinstance(n.ctx, (ast.Store, ast.Del)):
            definidos.add(n.id)
        elif isinstance(n, ast.arg):
            definidos.add(n.arg)
        elif isinstance(n, ast.ExceptHandler) and n.name:
            definidos.add(n.name)
        elif isinstance(n, (ast.Global, ast.Nonlocal)):
            definidos.update(n.names)
        elif isinstance(n, ast.comprehension):
            for t in ast.walk(n.target):
                if isinstance(t, ast.Name):
                    definidos.add(t.id)

    faltan = sorted({n.id for n in ast.walk(arbol)
                     if isinstance(n, ast.Name) and isinstance(n.ctx, ast.Load)
                     and n.id not in definidos})
    revisar(not faltan,
            f"{archivo}: usa nombres que nadie define: {', '.join(faltan[:6])}"
            if faltan else f"{archivo}: ningún nombre sin definir")


def test_metadatos_que_lee_la_app(html):
    """Sub-campos de los bloques de metadatos que la app lee y no existen.

    La franja de confianza leía `val.fallos`, y el dato escribe `no_cuadran`.
    Resultado: decía "sin descuadres de coherencia" SIEMPRE, hubiera o no, y
    nadie lo notó porque la frase que salía era la buena.

    Se compara contra los campos que escribe cada script, no contra un
    summaries.json concreto: así vale aunque la corrida del día no traiga el
    bloque. La ventana es la función donde se asigna la variable, para no
    confundirla con otra del mismo nombre en otra parte del archivo.
    """
    ESCRIBEN = {
        "validacion": {"cuadran", "no_cuadran", "sin_datos"},
        "validacion_tarjetas": {"iguales", "distintas", "sin_correspondencia"},
        "cobertura": {"kpis_del_reporte", "kpis_totales", "pct"},
    }
    JS = {"length", "filter", "map", "reduce", "forEach", "slice", "join",
          "includes", "replace", "toLowerCase", "toFixed", "sort", "find",
          "some", "every", "push", "split", "trim", "startsWith"}
    lineas = html.split("\n")
    for i, l in enumerate(lineas):
        m = re.search(r"const\s+(\w+)\s*=\s*DATA\??\.\s*(\w+)\s*\|\|", l)
        if not m:
            continue
        var, blq = m.group(1), m.group(2)
        if blq not in ESCRIBEN:
            continue
        j = i + 1
        while j < len(lineas) and not re.match(r"^\}", lineas[j]):
            j += 1
        usa = set(re.findall(r"\b" + var + r"\.(\w+)", "\n".join(lineas[i:j]))) - JS
        falta = sorted(usa - ESCRIBEN[blq])
        revisar(not falta,
                f"la app lee {blq}.{', '.join(falta)} y nadie lo escribe "
                f"(línea {i + 1}); los campos reales son "
                f"{', '.join(sorted(ESCRIBEN[blq]))}"
                if falta else f"{blq}: la app lee campos que existen")


def test_metas_no_escritas_a_mano(html, metas):
    """Metas escritas dentro del código en vez de leerse de metas.json.

    Había dieciséis: "meta 98%", "meta ≥52%", "const META = 0.02", umbrales de
    color con "> 52". Cambiar una meta en el archivo dejaba media app diciendo
    la vieja, y un tablero que se contradice a sí mismo no se usa para decidir.

    Se busca cada meta publicada, escrita como la escribiría una persona, en
    cualquier parte del HTML que no sea un comentario.
    """
    print("\n10. Metas escritas a mano en juanito.html")
    valores = []
    for ind in (metas.get("indicadores") or []):
        v = ind.get("meta")
        if v is None:
            continue
        if ind.get("unidad") in ("pct", "ratio"):
            valores.append((ind["clave"], f"{v * 100:g}%"))
        elif ind.get("unidad") == "dias":
            valores.append((ind["clave"], f"{v:g} d"))

    lineas = [l for l in html.split("\n")
              if not l.strip().startswith(("//", "*", "/*"))]
    cuerpo = "\n".join(lineas)
    malas = []
    for clave, txt in valores:
        # "meta 98%", "meta ≥52%", "meta <2%", "meta de 15%"
        pat = re.compile(r"meta\s*(?:de\s*)?[≥≤<>]?\s*" + re.escape(txt), re.I)
        for m in pat.finditer(cuerpo):
            malas.append(f"{clave}: «{m.group(0)}»")
    revisar(not malas,
            f"{len(valores)} metas publicadas, ninguna escrita a mano en el HTML"
            if not malas else "ESCRITAS A MANO: " + "; ".join(sorted(set(malas))[:5]))



def main():
    src_pbi = leer("scripts", "fetch_powerbi.py")
    src_ser = leer("scripts", "fetch_series.py")
    html = leer("juanito.html")

    test_claves_internas(src_pbi)
    test_orden_carga_build(src_pbi)
    test_derivados_declarados(src_pbi)
    test_sin_nombres_repetidos(src_pbi, 'fetch_powerbi.py')
    test_sin_nombres_repetidos(leer('scripts', 'fetch_series.py'), 'fetch_series.py')
    test_funciones_usadas(src_pbi, "dax_", "fetch_powerbi.py")
    test_funciones_usadas(src_ser, "serie_", "fetch_series.py")
    test_variable_antes_de_asignar(src_pbi, "fetch_powerbi.py")
    test_variable_antes_de_asignar(src_ser, "fetch_series.py")
    test_js_sin_nombres_repetidos(html)
    try:
        metas_json = json.loads(leer("data", "metas.json"))
    except Exception as e:
        metas_json = {}
        revisar(False, f"metas.json ilegible: {e}")
    test_metas_no_escritas_a_mano(html, metas_json)
    test_metadatos_que_lee_la_app(html)
    test_nombres_sin_definir(src_pbi, 'fetch_powerbi.py')
    test_nombres_sin_definir(src_ser, 'fetch_series.py')
    salidas = test_construccion(cargar_fetch_powerbi())
    test_campos(salidas, html)

    print()
    if FALLOS:
        print(f"{len(FALLOS)} problema(s) — corregir ANTES de correr el workflow:")
        for f in FALLOS:
            print(f"  · {f}")
        return 1
    print("Todo conectado. El workflow puede correrse.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
