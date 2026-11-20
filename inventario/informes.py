"""Informe en PDF y etiquetas con código QR."""
from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from .campos import CAMPOS_POR_CLAVE, dias_aviso, estado_garantia, fecha_a_date, mostrar_fecha, normalizar
from .pdf import A4, MM, PDF

AZUL = "#0071E3"
TEXTO = "#1D1D1F"
SECUNDARIO = "#6E6E73"
BORDE = "#E5E5EA"
FONDO = "#F5F5F7"
ROJO = "#FF3B30"
AMBAR = "#FF9500"
VERDE = "#34C759"
GRIS = "#8E8E93"
PALETA = ["#0071E3", "#5E5CE6", "#30B0C7", "#34C759", "#FF9500", "#FF2D55", "#AF52DE", "#A2845E"]


# =====================================================================
#  Códigos QR
# =====================================================================
_PATRON_QR = re.compile(r"^\s*INV:(\d+)\b", re.I)


def texto_qr(eq: dict) -> str:
    """Contenido del QR de un equipo. Empieza por «INV:<id>» para que el programa lo reconozca al
    escanearlo con un lector; el resto es legible desde cualquier móvil."""
    partes = [f"INV:{eq['id']}"]
    if eq.get("nombre_equipo"):
        partes.append(eq["nombre_equipo"])
    if eq.get("num_serie"):
        partes.append(f"S/N {eq['num_serie']}")
    if eq.get("persona"):
        partes.append(eq["persona"])
    return " | ".join(partes)[:200]


def id_desde_qr(texto: str) -> int | None:
    m = _PATRON_QR.match(texto or "")
    return int(m.group(1)) if m else None


def matriz_qr(texto: str) -> list[list[bool]]:
    import qrcode  # dependencia ligera, sólo se usa para calcular la matriz
    from qrcode.constants import ERROR_CORRECT_M
    q = qrcode.QRCode(border=0, error_correction=ERROR_CORRECT_M)
    q.add_data(texto)
    q.make(fit=True)
    return q.get_matrix()


# =====================================================================
#  Etiquetas
# =====================================================================
@dataclass(frozen=True)
class FormatoEtiqueta:
    nombre: str
    pagina: tuple[float, float]       # puntos
    ancho: float                      # mm
    alto: float                       # mm
    columnas: int
    filas: int
    margen_sup: float                 # mm
    margen_izq: float                 # mm
    sep_h: float = 0.0                # separación horizontal entre etiquetas (mm)
    sep_v: float = 0.0                # separación vertical (mm)


FORMATOS_ETIQUETA = [
    FormatoEtiqueta("A4 · 3×8 · 70×37 mm (Apli 1274 / Avery 3474)", A4, 70, 37, 3, 8, 0.5, 0),
    FormatoEtiqueta("A4 · 3×7 · 63,5×38,1 mm (Avery L7160)", A4, 63.5, 38.1, 3, 7, 15.15, 7.21, 2.54, 0),
    FormatoEtiqueta("A4 · 2×7 · 99,1×38,1 mm (Avery L7163)", A4, 99.1, 38.1, 2, 7, 15.15, 4.65, 2.54, 0),
    FormatoEtiqueta("A4 · 4×10 · 48,5×25,4 mm (Apli 1285)", A4, 48.5, 25.4, 4, 10, 21.5, 8.0, 0, 0),
    FormatoEtiqueta("Impresora de etiquetas · 62×29 mm (Brother DK-11209)", (62 * MM, 29 * MM), 62, 29, 1, 1, 0, 0),
    FormatoEtiqueta("Impresora de etiquetas · 54×25 mm (DYMO 11352)", (54 * MM, 25 * MM), 54, 25, 1, 1, 0, 0),
]


def generar_etiquetas(ruta: str | Path, equipos: list[dict], formato: FormatoEtiqueta,
                      saltar: int = 0, guias: bool = False) -> Path:
    """Genera un PDF con una etiqueta por equipo. `saltar` deja huecos al principio (para
    aprovechar hojas ya empezadas). `guias` dibuja el contorno de cada etiqueta."""
    pdf = PDF(formato.pagina, titulo="Etiquetas de inventario")
    por_hoja = formato.columnas * formato.filas
    lista: list[dict | None] = [None] * max(0, saltar) + list(equipos)
    for n, eq in enumerate(lista):
        if n % por_hoja == 0:
            pdf.nueva_pagina()
        if eq is None:
            continue
        k = n % por_hoja
        fila, col = divmod(k, formato.columnas)
        x = (formato.margen_izq + col * (formato.ancho + formato.sep_h)) * MM
        y = (formato.margen_sup + fila * (formato.alto + formato.sep_v)) * MM
        _dibujar_etiqueta(pdf, eq, x, y, formato.ancho * MM, formato.alto * MM, guias)
    return pdf.guardar(ruta)


def _dibujar_etiqueta(pdf: PDF, eq: dict, x: float, y: float, w: float, h: float, guias: bool):
    if guias:
        pdf.rect(x + 1, y + 1, w - 2, h - 2, borde="#C7C7CC", grosor=0.4, radio=4)
    margen = max(4.0, h * 0.09)
    lado = h - 2 * margen
    matriz = matriz_qr(texto_qr(eq))
    tam = lado / len(matriz)
    pdf.matriz(x + margen, y + margen, matriz, tam)
    tx = x + margen + lado + margen * 0.9
    ancho_txt = x + w - margen * 0.7 - tx
    escala = min(1.0, h / (37 * MM))
    titulo = eq.get("nombre_equipo") or eq.get("num_serie") or f"Equipo {eq['id']}"
    lineas = [(titulo, 10.5 * escala, True, TEXTO)]
    if eq.get("persona"):
        lineas.append((eq["persona"], 8 * escala, False, TEXTO))
    if eq.get("num_serie"):
        lineas.append((f"S/N {eq['num_serie']}", 7.5 * escala, False, SECUNDARIO))
    detalle = " · ".join(v for v in (eq.get("marca"), eq.get("modelo")) if v)
    if detalle:
        lineas.append((detalle, 7 * escala, False, SECUNDARIO))
    if eq.get("delegacion"):
        lineas.append((eq["delegacion"], 7 * escala, False, SECUNDARIO))
    lineas.append((f"ID {eq['id']}", 6.5 * escala, True, AZUL))
    alto_total = sum(t * 1.25 for _, t, _, _ in lineas)
    while alto_total > lado and len(lineas) > 2:  # quitar líneas secundarias si no caben
        lineas.pop(-2)
        alto_total = sum(t * 1.25 for _, t, _, _ in lineas)
    cy = y + margen + (lado - alto_total) / 2
    for texto, tam_f, negrita, color in lineas:
        cy += tam_f * 1.25
        pdf.texto(tx, cy - tam_f * 0.25, texto, tam_f, negrita, color, ancho_max=ancho_txt)


# =====================================================================
#  Informe
# =====================================================================
def _pie(pdf: PDF, generado: str):
    total = len(pdf.paginas)
    for i, ops in enumerate(pdf.paginas, start=1):
        pdf._actual = ops
        pdf.linea(40, pdf.alto - 32, pdf.ancho - 40, pdf.alto - 32, BORDE, 0.5)
        pdf.texto(40, pdf.alto - 20, generado, 7.5, color=SECUNDARIO)
        pdf.texto(pdf.ancho - 40, pdf.alto - 20, f"Página {i} de {total}", 7.5, color=SECUNDARIO, alin="der")


def _grafico(pdf: PDF, x: float, y: float, w: float, titulo: str, items: list[tuple[str, int]], color: str,
             filas: int = 6) -> float:
    """Gráfico de barras horizontales dentro de una tarjeta. Devuelve la altura usada."""
    if len(items) > filas:
        resto = sum(n for _, n in items[filas - 1:])
        items = items[: filas - 1] + [(f"Otros ({len(items) - filas + 1})", resto)]
    alto_fila = 15
    h = 34 + max(1, len(items)) * alto_fila + 8
    pdf.rect(x, y, w, h, relleno="#FFFFFF", borde=BORDE, grosor=0.6, radio=8)
    pdf.texto(x + 12, y + 20, titulo, 10, True)
    if not items:
        pdf.texto(x + 12, y + 40, "Sin datos", 8.5, color=SECUNDARIO)
        return h
    maximo = max(n for _, n in items) or 1
    etiqueta_w = w * 0.38
    num_w = 28
    barra_max = w - 24 - etiqueta_w - num_w
    for i, (valor, n) in enumerate(items):
        cy = y + 34 + i * alto_fila
        pdf.texto(x + 12, cy + 10, str(valor), 8.5, ancho_max=etiqueta_w - 6)
        largo = max(4, barra_max * n / maximo)
        pdf.rect(x + 12 + etiqueta_w, cy + 3.5, largo, 6.5, relleno=color, radio=3.2)
        pdf.texto(x + w - 12, cy + 10, str(n), 8.5, True, SECUNDARIO, alin="der")
    return h


def _tabla(pdf: PDF, y: float, columnas: list[tuple[str, str, float]], filas: list[dict],
           colores: dict[int, str] | None = None, titulo: str = "", generado: str = "") -> float:
    """Tabla paginada. columnas = [(clave, cabecera, ancho_relativo)]. Devuelve la y final."""
    x0, ancho = 40, pdf.ancho - 80
    total_rel = sum(c[2] for c in columnas)
    anchos = [ancho * c[2] / total_rel for c in columnas]
    alto_fila = 15
    limite = pdf.alto - 50

    def cabecera(yc: float) -> float:
        pdf.rect(x0, yc, ancho, 18, relleno=FONDO, radio=4)
        cx = x0
        for (clave, texto, _), w in zip(columnas, anchos):
            pdf.texto(cx + 5, yc + 12, texto.upper(), 6.8, True, SECUNDARIO, ancho_max=w - 8)
            cx += w
        return yc + 20

    if titulo:
        if y + 60 > limite:
            pdf.nueva_pagina()
            y = 50
        pdf.texto(x0, y + 12, titulo, 12, True)
        y += 22
    y = cabecera(y)
    for i, fila in enumerate(filas):
        if y + alto_fila > limite:
            pdf.nueva_pagina()
            y = cabecera(50)
        if i % 2:
            pdf.rect(x0, y, ancho, alto_fila, relleno="#FAFAFC")
        cx = x0
        color_fila = (colores or {}).get(i, TEXTO)
        for (clave, _, _), w in zip(columnas, anchos):
            valor = fila.get(clave, "")
            pdf.texto(cx + 5, y + 10.5, str(valor), 7.8, color=color_fila, ancho_max=w - 8)
            cx += w
        y += alto_fila
    pdf.linea(x0, y + 1, x0 + ancho, y + 1, BORDE, 0.5)
    return y + 8


def generar_informe(ruta: str | Path, equipos: list[dict], descripcion: str = "", usuario: str = "",
                    incluir_listado: bool = True) -> Path:
    hoy = date.today()
    dias = dias_aviso()
    generado = f"Generado el {datetime.now():%d/%m/%Y %H:%M}" + (f" por {usuario}" if usuario else "")
    pdf = PDF(A4, titulo="Informe de inventario", autor=usuario)
    pdf.nueva_pagina()
    W = pdf.ancho

    # ---- cabecera
    pdf.rect(0, 0, W, 92, relleno=AZUL)
    pdf.texto(40, 46, "Informe de inventario", 22, True, "#FFFFFF")
    pdf.texto(40, 68, f"{hoy:%d/%m/%Y}" + (f" · {descripcion}" if descripcion else ""), 10,
              color="#DCEBFF", ancho_max=W - 80)

    # ---- cifras
    estados = [estado_garantia(e.get("fin_garantia", ""), hoy) for e in equipos]
    no_baja = [i for i, e in enumerate(equipos) if e.get("estado") != "Baja"]
    cifras = [
        (len(equipos), "Equipos", AZUL),
        (sum(1 for e in equipos if (e.get("estado") or "Activo") == "Activo"), "Activos", VERDE),
        (sum(1 for i in no_baja if estados[i] == "caducada"), "Garantía caducada", ROJO),
        (sum(1 for i in no_baja if estados[i] == "proxima"), f"Caducan ≤ {dias} días", AMBAR),
        (len({normalizar(e.get('persona')) for e in equipos if e.get("persona")}), "Personas", GRIS),
    ]
    y = 112
    sep = 10
    w_c = (W - 80 - sep * 4) / 5
    for i, (n, texto, color) in enumerate(cifras):
        x = 40 + i * (w_c + sep)
        pdf.rect(x, y, w_c, 58, relleno="#FFFFFF", borde=BORDE, grosor=0.6, radio=8)
        pdf.texto(x + 12, y + 30, str(n), 20, True, color)
        pdf.texto(x + 12, y + 46, texto.replace("≤", "<="), 8, color=SECUNDARIO, ancho_max=w_c - 20)
    y += 58 + 18

    # ---- gráficos (2 columnas)
    def contar(clave):
        return Counter((e.get(clave) or "(sin dato)") for e in equipos).most_common()
    anios = Counter((str(fecha_a_date(e.get("fecha_compra", "")).year) if fecha_a_date(e.get("fecha_compra", ""))
                     else "(sin dato)") for e in equipos)
    graficos = [("Por delegación", contar("delegacion")), ("Por marca", contar("marca")),
                ("Por sistema operativo", contar("os")), ("Por estado", contar("estado")),
                ("Por año de compra", sorted(anios.items(), key=lambda t: t[0], reverse=True)),
                ("Por memoria RAM", contar("ram"))]
    w_g = (W - 80 - 12) / 2
    for fila in range(0, len(graficos), 2):
        alturas = []
        for col in range(2):
            if fila + col >= len(graficos):
                continue
            titulo, items = graficos[fila + col]
            alturas.append(_grafico(pdf, 40 + col * (w_g + 12), y, w_g, titulo, items,
                                    PALETA[(fila + col) % len(PALETA)]))
        y += max(alturas) + 12

    # ---- garantías que requieren atención
    criticos = sorted((e for i, e in enumerate(equipos) if i in set(no_baja) and estados[i] in ("caducada", "proxima")),
                      key=lambda e: e.get("fin_garantia", ""))
    filas, colores = [], {}
    for i, e in enumerate(criticos):
        d = (fecha_a_date(e["fin_garantia"]) - hoy).days
        filas.append({**e, "fin_garantia": mostrar_fecha(e["fin_garantia"]),
                      "_dias": f"hace {-d}" if d < 0 else f"en {d}"})
        colores[i] = ROJO if d < 0 else "#C76E00"
    pdf.nueva_pagina()
    y = 50
    if filas:
        y = _tabla(pdf, y, [("fin_garantia", "Expira", 1.1), ("_dias", "Días", 0.8), ("nombre_equipo", "Equipo", 1.4),
                            ("persona", "Persona", 1.6), ("delegacion", "Delegación", 1.2),
                            ("marca", "Marca", 0.9), ("num_serie", "Nº serie", 1.2)],
                   filas, colores, titulo=f"Garantías caducadas o que caducan en {dias} días ({len(filas)})")
    else:
        pdf.texto(40, y + 12, "Garantías", 12, True)
        pdf.texto(40, y + 32, "No hay garantías caducadas ni próximas a caducar.", 9, color=SECUNDARIO)
        y += 50

    # ---- listado completo
    if incluir_listado and equipos:
        lista = [{**e, "fin_garantia": mostrar_fecha(e.get("fin_garantia", "")),
                  "_equipo": " ".join(v for v in (e.get("marca"), e.get("modelo")) if v)} for e in equipos]
        y = _tabla(pdf, y + 14, [("persona", "Persona", 1.5), ("delegacion", "Delegación", 1.1),
                                 ("nombre_equipo", "Nombre", 1.2), ("_equipo", "Marca / modelo", 1.7),
                                 ("num_serie", "Nº serie", 1.1), ("fin_garantia", "Garantía", 0.9),
                                 ("estado", "Estado", 0.8)],
                   lista, titulo=f"Listado de equipos ({len(lista)})")
    _pie(pdf, generado)
    return pdf.guardar(ruta)


# =====================================================================
#  Periodos de garantía
# =====================================================================
def periodos_garantia(hoy: date | None = None) -> list[tuple[str, date | None, date | None]]:
    """[(nombre, desde, hasta)] para filtrar garantías que caducan en un periodo."""
    hoy = hoy or date.today()
    t = (hoy.month - 1) // 3
    ini_t = date(hoy.year, 3 * t + 1, 1)
    fin_t = date(hoy.year + (1 if t == 3 else 0), 1 if t == 3 else 3 * t + 4, 1)
    sig_fin = date(fin_t.year + (1 if fin_t.month == 10 else 0), 1 if fin_t.month == 10 else fin_t.month + 3, 1)
    from datetime import timedelta
    un_dia = timedelta(days=1)
    return [
        (f"Caducadas y próximas ({dias_aviso()} días)", None, hoy + timedelta(days=dias_aviso())),
        ("Ya caducadas", None, hoy - un_dia),
        (f"Este trimestre (T{t + 1} {hoy.year})", ini_t, fin_t - un_dia),
        ("Próximo trimestre", fin_t, sig_fin - un_dia),
        (f"Este año ({hoy.year})", date(hoy.year, 1, 1), date(hoy.year, 12, 31)),
        (f"Próximo año ({hoy.year + 1})", date(hoy.year + 1, 1, 1), date(hoy.year + 1, 12, 31)),
    ]


__all__ = ["FORMATOS_ETIQUETA", "generar_etiquetas", "generar_informe", "id_desde_qr", "texto_qr",
           "periodos_garantia", "CAMPOS_POR_CLAVE"]
