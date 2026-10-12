"""
verificar_split.py — Comprueba que la separación de BOMs salió correcta.

Compara el reporte consolidado contra los PDFs que generó el splitter y avisa
de cualquier diferencia:

  · BOMs del reporte que no tienen archivo generado.
  · Archivos generados que no corresponden a ningún BOM del reporte.
  · Componentes que se PERDIERON al cortar.
  · Componentes de OTRO BOM que se colaron en el archivo (contaminación).
  · Archivos que el procesador de redline no podría leer (sin tabla, sin
    'End of Report', formato no reconocido).
  · Portada con un item distinto al del archivo.

La verificación es INDEPENDIENTE del splitter: no usa su lógica de corte, sino
que lee los componentes de la columna Item de cada tabla. Así un fallo del
corte se detecta en vez de repetirse.

Uso:
    1. Ajusta las rutas en CONFIGURACIÓN.
    2. python verificar_split.py
    3. Revisa la consola y el Excel con el detalle.

Requiere:  pip install pymupdf openpyxl
"""

import os
import re
from pathlib import Path

try:
    import fitz
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
except ImportError as exc:
    print(f"Falta una dependencia: {exc}")
    print("Ejecuta: pip install pymupdf openpyxl")
    raise SystemExit(1)

__version__ = "1.0.0"

# ── CONFIGURACIÓN ───────────────────────────────────────────────────────────
CARPETA_REPORTES = "Input_PDFs"      # reportes consolidados (entrada del splitter)
CARPETA_SALIDA   = "Split_Output"    # PDFs separados (salida del splitter)
INFORME_XLSX     = "Verificacion_Split.xlsx"
# ────────────────────────────────────────────────────────────────────────────

ITEM_RE = re.compile(r"\bItem:\s*([A-Z0-9]{2,4}-[A-Z0-9-]{4,})", re.I)
SEP_RE  = re.compile(r"-{5,}")
GRUPO_RE = re.compile(r"-+")
EOR_RE  = re.compile(r"End of Report", re.I)


def _norm(texto):
    return str(texto).strip().upper().replace("‐", "-").replace("‑", "-")


def _spans(page):
    for b in page.get_text("dict")["blocks"]:
        if b["type"] != 0:
            continue
        for line in b["lines"]:
            for s in line["spans"]:
                yield s


def lineas(page):
    """[(y, texto, x0)] de la página, ordenado por y."""
    por_y = {}
    for s in _spans(page):
        por_y.setdefault(round(s["bbox"][1], 1), []).append(s)
    out = []
    for y in sorted(por_y):
        ss = sorted(por_y[y], key=lambda s: s["bbox"][0])
        out.append((y, "".join(s["text"] for s in ss),
                    min(s["bbox"][0] for s in ss)))
    return out


# ════════════════════════════════════════════════════════════════════════════
# LOCALIZAR LA COLUMNA ITEM
# ════════════════════════════════════════════════════════════════════════════

def tablas_del_doc(doc):
    """[(pagina, y_de_la_linea_de_guiones, x_columna_item, char_w)] — una por
    cada tabla de componentes encontrada."""
    tablas = []
    for pn in range(len(doc)):
        for y, texto, _ in lineas(doc[pn]):
            if texto.count("-") < 9 or not SEP_RE.search(texto):
                continue
            grupos = [m.start() for m in GRUPO_RE.finditer(texto)]
            if len(grupos) < 4:
                continue                 # línea de guiones de otra sección
            cajas = [s["bbox"] for s in _spans(doc[pn])
                     if abs(s["bbox"][1] - y) < 0.6]
            x_ini = min(b[0] for b in cajas)
            x_fin = max(b[2] for b in cajas)
            largo = len(texto.rstrip())
            if not largo:
                continue
            char_w = (x_fin - x_ini) / largo
            tablas.append((pn, y, x_ini + grupos[3] * char_w, char_w))
    return tablas


def columna_item(doc):
    """(x_de_la_columna_item, ancho_de_caracter) o (None, None)."""
    tablas = tablas_del_doc(doc)
    if not tablas:
        return None, None
    return tablas[0][2], tablas[0][3]


def _es_codigo(texto):
    """¿Tiene forma de código de part number? Descarta guiones sueltos y
    palabras de descripción que casualmente lleven un guion."""
    v = _norm(texto)
    if len(v) < 6 or "-" not in v:
        return False
    if set(v) <= set("-"):
        return False                      # la propia línea de guiones
    return any(ch.isdigit() for ch in v)  # un part number siempre lleva dígitos


def componentes_de(doc, desde_pagina=0, desde_y=None,
                   hasta_pagina=None, hasta_y=None):
    """Códigos de la COLUMNA ITEM que están DEBAJO de una línea de guiones.

    Exigir que estén bajo la línea de guiones evita contar la descripción del
    encabezado ('MULTI-MED …') o el propio separador como si fueran componentes.
    """
    tablas = tablas_del_doc(doc)
    if not tablas:
        return []
    hasta_pagina = len(doc) - 1 if hasta_pagina is None else hasta_pagina

    encontrados = []
    for pn in range(desde_pagina, min(hasta_pagina + 1, len(doc))):
        # Línea de guiones vigente en esta página (la última por encima)
        de_pagina = [t for t in tablas if t[0] == pn]
        anteriores = [t for t in tablas if t[0] < pn]
        if not de_pagina and not anteriores:
            continue
        x_item = (de_pagina[0][2] if de_pagina else anteriores[-1][2])
        char_w = (de_pagina[0][3] if de_pagina else anteriores[-1][3])
        tol = char_w * 1.5
        y_tabla = de_pagina[0][1] if de_pagina else -1

        for x0, y0, x1, y1, t, *_ in doc[pn].get_text("words"):
            if abs(x0 - x_item) > tol:
                continue
            if y0 <= y_tabla + 1:
                continue                  # encabezado o separador
            if pn == desde_pagina and desde_y is not None and y0 < desde_y:
                continue
            if pn == hasta_pagina and hasta_y is not None and y0 >= hasta_y:
                continue
            if _es_codigo(t):
                encontrados.append(_norm(t))
    return encontrados


# ════════════════════════════════════════════════════════════════════════════
# LEER EL REPORTE CONSOLIDADO
# ════════════════════════════════════════════════════════════════════════════

def boms_del_reporte(doc):
    """{item: [componentes]} — lo que el reporte original contiene por cada BOM."""
    x_item, char_w = columna_item(doc)
    if x_item is None:
        return {}, None

    # Dónde empieza cada BOM
    marcas = []
    for pn in range(len(doc)):
        for y, texto, _ in lineas(doc[pn]):
            m = ITEM_RE.search(texto)
            if m:
                codigo = _norm(m.group(1).rstrip(".,;:"))
                if marcas and marcas[-1][0] == codigo:
                    continue        # encabezado repetido del mismo BOM
                marcas.append((codigo, pn, y))

    resultado = {}
    for i, (codigo, pn, y) in enumerate(marcas):
        if i + 1 < len(marcas):
            _, pn_fin, y_fin = marcas[i + 1]
        else:
            pn_fin, y_fin = len(doc) - 1, None
        comps = componentes_de(doc, pn, y, pn_fin, y_fin)
        # El propio item puede aparecer en la columna; no es un componente suyo
        resultado.setdefault(codigo, []).extend(c for c in comps if c != codigo)
    return resultado, x_item


# ════════════════════════════════════════════════════════════════════════════
# REVISAR UN PDF SEPARADO
# ════════════════════════════════════════════════════════════════════════════

def revisar_archivo(ruta):
    """Devuelve (componentes, avisos) del PDF separado."""
    avisos = []
    doc = fitz.open(ruta)
    try:
        texto_total = "".join(doc[i].get_text() for i in range(len(doc)))

        x_item, char_w = columna_item(doc)
        if x_item is None:
            avisos.append("sin tabla de componentes reconocible")
            return [], avisos, None

        if not EOR_RE.search(texto_total):
            avisos.append("sin 'End of Report' (el redline no podrá delimitarlo)")

        # Item declarado en el documento
        item_doc = None
        for pn in range(len(doc)):
            for _, texto, _ in lineas(doc[pn]):
                m = ITEM_RE.search(texto)
                if m:
                    item_doc = _norm(m.group(1).rstrip(".,;:"))
                    break
            if item_doc:
                break

        comps = [c for c in componentes_de(doc) if c != item_doc]
        return comps, avisos, item_doc
    finally:
        doc.close()


def validable_por_redline(ruta):
    """¿bom_layout lo aceptaría? Solo si el módulo está disponible."""
    try:
        from bom_layout import detect_layout, validate_layout
    except ImportError:
        return None, "bom_layout.py no está en la carpeta (no se comprobó)"
    doc = fitz.open(ruta)
    try:
        ok, msg = validate_layout(doc, detect_layout(doc))
        return ok, msg
    except Exception as e:
        return False, repr(e)
    finally:
        doc.close()


# ════════════════════════════════════════════════════════════════════════════
# INFORME
# ════════════════════════════════════════════════════════════════════════════

def crear_informe(filas, ruta):
    wb = Workbook(); ws = wb.active; ws.title = "Verificación"
    cab = ["bom", "archivo", "estado", "comp_reporte", "comp_archivo",
           "faltantes", "sobrantes", "detalle"]
    fill = PatternFill("solid", fgColor="1F4E78")
    fnt = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    thin = Side(style="thin", color="D0D0D0")
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    rojo = Font(name="Arial", size=10, color="C00000")

    for j, h in enumerate(cab, 1):
        c = ws.cell(1, j, h); c.fill = fill; c.font = fnt; c.border = bd
        c.alignment = Alignment(horizontal="center")
    for i, f in enumerate(filas, 2):
        for j, k in enumerate(cab, 1):
            c = ws.cell(i, j, f.get(k, "")); c.border = bd
            if f.get("estado") != "OK":
                c.font = rojo
    for col, w in zip("ABCDEFGH", [20, 36, 12, 14, 14, 34, 34, 46]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    wb.save(str(ruta))


def main():
    base = Path(__file__).parent
    carpeta_rep = base / CARPETA_REPORTES
    carpeta_sal = base / CARPETA_SALIDA

    print("=" * 66)
    print(f" Verificación del Split de BOMs v{__version__}")
    print("=" * 66)

    if not carpeta_rep.is_dir():
        print(f"No existe la carpeta '{CARPETA_REPORTES}'.")
        return
    if not carpeta_sal.is_dir():
        print(f"No existe la carpeta '{CARPETA_SALIDA}'.")
        return

    # ── Reportes consolidados ───────────────────────────────────────────────
    vistos, reportes = set(), []
    for f in sorted(list(carpeta_rep.glob("*.pdf")) + list(carpeta_rep.glob("*.PDF"))):
        clave = str(f.resolve()).lower()
        if clave not in vistos:
            vistos.add(clave); reportes.append(f)

    esperado = {}
    for rep in reportes:
        doc = fitz.open(rep)
        try:
            boms, x_item = boms_del_reporte(doc)
            if x_item is None:
                print(f"  AVISO: '{rep.name}' no tiene tabla de componentes reconocible")
            for item, comps in boms.items():
                esperado.setdefault(item, []).extend(comps)
        finally:
            doc.close()
    print(f"Reportes leídos : {len(reportes)}")
    print(f"BOMs esperados  : {len(esperado)}")

    # ── PDFs separados ──────────────────────────────────────────────────────
    vistos, salidas = set(), []
    for f in sorted(list(carpeta_sal.glob("*.pdf")) + list(carpeta_sal.glob("*.PDF"))):
        clave = str(f.resolve()).lower()
        if clave not in vistos:
            vistos.add(clave); salidas.append(f)
    print(f"Archivos generados: {len(salidas)}\n")

    # Emparejar cada archivo con su BOM por el nombre
    por_item = {}
    for f in salidas:
        stem = _norm(f.stem)
        destino = next((it for it in esperado if it == stem), None)
        if destino is None:
            destino = next((it for it in esperado if it in stem), None)
        por_item.setdefault(destino, []).append(f)

    filas, problemas = [], 0

    for item in sorted(esperado):
        comps_rep = esperado[item]
        archivos = por_item.get(item, [])

        if not archivos:
            problemas += 1
            print(f"  [FALTA]  {item}: no se generó ningún archivo")
            filas.append({"bom": item, "archivo": "", "estado": "FALTA",
                          "comp_reporte": len(comps_rep), "comp_archivo": 0,
                          "faltantes": ", ".join(sorted(set(comps_rep))[:12]),
                          "sobrantes": "", "detalle": "sin archivo generado"})
            continue

        for ruta in archivos:
            comps_arch, avisos, item_doc = revisar_archivo(ruta)
            ok_layout, msg_layout = validable_por_redline(ruta)
            if ok_layout is False:
                avisos.append(f"el redline lo rechazaría: {msg_layout}")
            if item_doc and item_doc != item:
                avisos.append(f"la portada dice '{item_doc}'")

            faltan = sorted(set(comps_rep) - set(comps_arch))
            sobran = sorted(set(comps_arch) - set(comps_rep))

            if faltan:
                avisos.append(f"{len(faltan)} componente(s) perdido(s)")
            if sobran:
                avisos.append(f"{len(sobran)} componente(s) de otro BOM")

            estado = "OK" if not avisos else "REVISAR"
            if estado != "OK":
                problemas += 1
                print(f"  [REVISAR] {item} -> {ruta.name}")
                for a in avisos:
                    print(f"            · {a}")
                if faltan:
                    print(f"            · faltan: {', '.join(faltan[:8])}"
                          + (" …" if len(faltan) > 8 else ""))
                if sobran:
                    print(f"            · sobran: {', '.join(sobran[:8])}"
                          + (" …" if len(sobran) > 8 else ""))

            filas.append({
                "bom": item, "archivo": ruta.name, "estado": estado,
                "comp_reporte": len(set(comps_rep)),
                "comp_archivo": len(set(comps_arch)),
                "faltantes": ", ".join(faltan[:12]),
                "sobrantes": ", ".join(sobran[:12]),
                "detalle": " | ".join(avisos),
            })

    # Archivos que no corresponden a ningún BOM
    huerfanos = por_item.get(None, [])
    for f in huerfanos:
        problemas += 1
        print(f"  [SOBRA]  {f.name}: no corresponde a ningún BOM del reporte")
        filas.append({"bom": "", "archivo": f.name, "estado": "SOBRA",
                      "comp_reporte": 0, "comp_archivo": 0, "faltantes": "",
                      "sobrantes": "", "detalle": "sin BOM correspondiente"})

    crear_informe(filas, base / INFORME_XLSX)

    correctos = sum(1 for f in filas if f["estado"] == "OK")
    print("\n" + "=" * 66)
    print(f" Correctos      : {correctos}")
    print(f" Con problemas  : {problemas}")
    print(f" Informe        : {INFORME_XLSX}")
    if problemas == 0:
        print("\n La separación salió bien: todos los BOMs están completos.")
    print("=" * 66)


if __name__ == "__main__":
    main()
