#!/usr/bin/env python3
"""Comprueba que la corrida produjo un dato de HOY antes de publicarlo.

Existe porque una corrida verde con datos de ayer es peor que una roja: nadie
mira las verdes. Si el extractor falla a medias y deja el JSON anterior
intacto, sin esto el workflow publica lo mismo de ayer, termina en verde, y la
app muestra un dato viejo sin que nadie se entere.

Al fallar, GitHub manda el correo de "workflow failed" al dueño del
repositorio — que es la única alarma que funciona sin que nadie esté mirando.

Uso:  python scripts/verificar_publicado.py [--dias-max N]
"""

import datetime
import json
import pathlib
import sys

RUTA = pathlib.Path("data/latest/summaries.json")


def fecha_lima(ahora=None):
    """Hoy en Lima (UTC-5, sin horario de verano) con el formato de la app."""
    ahora = ahora or datetime.datetime.now(datetime.timezone.utc)
    return (ahora - datetime.timedelta(hours=5)).strftime("%d/%m/%Y")


def revisar(datos, hoy):
    """Devuelve (motivo_del_fallo | None, resumen_para_el_log)."""
    fecha = datos.get("fecha")
    diag = datos.get("diagnostico") or []
    lineas = [f"fecha publicada: {fecha} · esperada: {hoy} · diagnósticos: {len(diag)}"]
    for x in diag[:10]:
        lineas.append(f"  - {x.get('consulta')} [{x.get('http')}] "
                      f"{str(x.get('error'))[:140]}")

    # Los desgloses son lo primero que desaparece cuando Power BI limita las
    # peticiones, y desaparecen en silencio: el JSON sigue siendo válido y la
    # corrida sigue en verde. Se comprueban por nombre.
    esperados = [
        ("control_interno", "por_area"),
        ("control_interno", "planes_por_planta"),
        ("cuentas_por_cobrar", "tramos"),
        ("cuentas_por_pagar", "proveedores_criticos"),
        ("margen_variable", "por_cliente"),
        ("mermas", "por_uen"),
        ("compras", "faltantes"),
        ("fill_rate", "por_grupo"),
        ("sop_inventario", "dead_por_categoria"),
    ]
    reportes = (datos.get("empresas", {}).get("PAUNO", {}) or {}).get("reportes", {})
    vacios = [f"{r}/{k}" for r, k in esperados if not (reportes.get(r, {}) or {}).get(k)]
    lineas.append(f"desgloses con datos: {len(esperados) - len(vacios)}/{len(esperados)}")

    # Contradicciones entre cifras del propio Power BI. No impiden publicar
    # —ya están publicadas y marcadas— pero sí tienen que llegar por correo:
    # un total que no es la suma de sus partes es un dato en el que alguien
    # va a apoyar una decisión.
    # Diferencias ya investigadas y explicadas. No encienden la alarma: una
    # alarma que suena todos los días por lo mismo deja de mirarse, y entonces
    # el día que aparece algo nuevo tampoco se ve.
    conocidos = set()
    ruta_con = pathlib.Path("data/descuadres_conocidos.json")
    if ruta_con.exists():
        try:
            conocidos = {x["prueba"] for x in
                         json.loads(ruta_con.read_text(encoding="utf-8"))["conocidos"]}
        except Exception as e:
            lineas.append(f"  (no se pudo leer descuadres_conocidos.json: {e})")

    # Un 429 no es un descuadre: es que la corrida pidió más consultas de las
    # que el espacio de trabajo admite y varias series llegaron vacías. Los
    # descuadres que salen después son consecuencia, no causa, y perseguirlos
    # lleva al sitio equivocado. Se dice primero y con todas las letras.
    todos_diag = (datos.get("diagnostico") or [])
    try:
        ser = json.loads(pathlib.Path("data/latest/series.json").read_text(encoding="utf-8"))
        todos_diag = todos_diag + (ser.get("diagnostico") or [])
    except Exception:
        pass
    throttled = [d for d in todos_diag if str(d.get("http")) == "429"]
    if throttled:
        lineas.append(f"⚠ PETICIONES LIMITADAS: {len(throttled)} consultas "
                      f"devolvieron 429. Las series que falten y los descuadres "
                      f"de abajo son consecuencia de eso, no de los datos.")

    val = datos.get("validacion") or {}
    todos = val.get("no_cuadran") or []
    # El veredicto lo pone validar_datos al publicar, que normaliza los
    # espacios dobles de los nombres de producto. Aquí se respeta en vez de
    # repetir el emparejamiento: hecho dos veces, las dos versiones
    # discrepaban y el mismo hallazgo salía "conocido" en un sitio y "nuevo"
    # en el otro. La comparación por nombre queda de respaldo para datos
    # antiguos, que aún no traen la marca.
    def _nuevo(x):
        if "conocido" in x:
            return not x["conocido"]
        return x["prueba"] not in conocidos

    descuadres = [x for x in todos if _nuevo(x)]
    ya_vistos = [x for x in todos if not _nuevo(x)]
    lineas.append(f"coherencia: {val.get('cuadran', '?')} cuadran · "
                  f"{len(descuadres)} no cuadran")
    for x in descuadres[:8]:
        # Las comprobaciones que no comparan dos números llegan sin cifras:
        # su detalle ES el hallazgo.
        if x.get("publicado") is None:
            lineas.append(f"  ✗ {x['prueba']}: {x.get('detalle') or 'sin detalle'}")
        else:
            lineas.append(f"  ✗ {x['prueba']}: publicado {x['publicado']:,} · "
                          f"esperado {x['esperado']:,} ({x['diferencia_pct']}%)")
    for x in ya_vistos:
        lineas.append(f"  · {x['prueba']}: diferencia conocida y explicada")

    # Diferencias contra las tarjetas del propio reporte. Estas no admiten
    # lista de conocidos: si la app no dice lo que dice Power BI, es un error.
    tar = datos.get("validacion_tarjetas") or {}
    distintas = tar.get("distintas") or []
    lineas.append(f"tarjetas: {tar.get('iguales', 0)} iguales · "
                  f"{len(distintas)} distintas")
    for x in distintas[:8]:
        lineas.append(f"  ✗ {x['kpi']}: Power BI {x['en_powerbi']} · "
                      f"Juanito {x['en_juanito']}")

    if fecha != hoy:
        return (f"el dato quedó con fecha {fecha}, no {hoy}: la corrida terminó "
                "sin actualizar nada"), lineas
    if vacios:
        return f"desgloses vacíos: {', '.join(vacios)}", lineas
    if distintas:
        return (f"{len(distintas)} cifras difieren de la tarjeta de Power BI: "
                + "; ".join(x["kpi"] for x in distintas[:4])), lineas
    if descuadres:
        return (f"{len(descuadres)} descuadres nuevos: "
                + "; ".join(x["prueba"] for x in descuadres[:4])), lineas
    return None, lineas


def main():
    if not RUTA.exists():
        sys.exit(f"ABORTA: no existe {RUTA}")
    datos = json.loads(RUTA.read_text(encoding="utf-8"))
    motivo, lineas = revisar(datos, fecha_lima())
    print("\n".join(lineas))
    if motivo:
        sys.exit(f"ABORTA: {motivo}")
    print("OK — dato de hoy y desgloses completos.")


if __name__ == "__main__":
    main()
