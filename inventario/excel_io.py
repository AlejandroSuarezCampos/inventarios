"""Importación y exportación de Excel / CSV con el formato del inventario."""
from __future__ import annotations

import csv
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .campos import (CAMPOS, CAMPOS_POR_CLAVE, CLAVES, GRUPOS, clave_para_cabecera,
                     fecha_a_date, parsear_fecha, valor_texto, estado_garantia)


@dataclass
class ResultadoLectura:
    filas: list[dict] = field(default_factory=list)
    hojas_leidas: list[str] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)
    columnas_reconocidas: dict[str, list[str]] = field(default_factory=dict)  # hoja -> etiquetas
    columnas_ignoradas: dict[str, list[str]] = field(default_factory=dict)


def _mapear_fila(valores) -> dict[int, str]:
    """Devuelve {índice_columna: clave} para las celdas que parecen cabeceras."""
    mapa: dict[int, str] = {}
    usadas: set[str] = set()
    for i, v in enumerate(valores):
        clave = clave_para_cabecera(v)
        if clave and clave not in usadas:
            mapa[i] = clave
            usadas.add(clave)
    return mapa


def _convertir(clave: str, valor) -> str:
    if CAMPOS_POR_CLAVE[clave].tipo == "fecha":
        return parsear_fecha(valor)
    return valor_texto(valor)


def _procesar_tabla(nombre: str, filas: list[tuple], res: ResultadoLectura) -> None:
    # 1) localizar la fila de cabeceras (la que más cabeceras conocidas tenga en las 30 primeras)
    mejor_idx, mejor_mapa = -1, {}
    for idx, fila in enumerate(filas[:30]):
        mapa = _mapear_fila(fila)
        if len(mapa) > len(mejor_mapa):
            mejor_idx, mejor_mapa = idx, mapa
    if len(mejor_mapa) < 2:
        res.avisos.append(f"Hoja '{nombre}': no se encontraron cabeceras reconocibles; se omite.")
        return

    # 2) una segunda fila de cabeceras justo debajo (cabeceras partidas en 2 filas)
    inicio = mejor_idx + 1
    if inicio < len(filas):
        extra = _mapear_fila(filas[inicio])
        nuevas = {i: c for i, c in extra.items()
                  if i not in mejor_mapa and c not in mejor_mapa.values()}
        if nuevas and len(extra) >= len(nuevas):
            mejor_mapa.update(nuevas)
            inicio += 1

    cab = filas[mejor_idx]
    res.hojas_leidas.append(nombre)
    res.columnas_reconocidas[nombre] = [CAMPOS_POR_CLAVE[c].etiqueta for c in mejor_mapa.values()]
    ignoradas = []
    for i, v in enumerate(cab):
        if i not in mejor_mapa and v not in (None, "") and clave_para_cabecera(v) is None:
            ignoradas.append(str(v))
    if ignoradas:
        res.columnas_ignoradas[nombre] = ignoradas

    # 3) leer datos
    for num, fila in enumerate(filas[inicio:], start=inicio + 1):
        datos = {c: "" for c in CLAVES}
        for i, clave in mejor_mapa.items():
            if i < len(fila):
                datos[clave] = _convertir(clave, fila[i])
        if not any(datos.values()):
            continue
        # evitar volver a importar una fila de cabeceras repetida
        if len(_mapear_fila(fila)) >= max(2, len(mejor_mapa) // 2):
            continue
        datos["_origen"] = f"{nombre}!{num}"
        res.filas.append(datos)


def leer_excel(ruta: str | Path) -> ResultadoLectura:
    ruta = Path(ruta)
    res = ResultadoLectura()
    if ruta.suffix.lower() in (".csv", ".txt"):
        return leer_csv(ruta)
    wb = load_workbook(ruta, data_only=True, read_only=False)
    try:
        for ws in wb.worksheets:
            if ws.sheet_state != "visible" and len(wb.worksheets) > 1:
                continue
            filas = [tuple(f) for f in ws.iter_rows(values_only=True)]
            _procesar_tabla(ws.title, filas, res)
    finally:
        wb.close()
    if not res.hojas_leidas:
        res.avisos.append("No se encontró ninguna hoja con el formato del inventario.")
    return res


def leer_csv(ruta: str | Path) -> ResultadoLectura:
    ruta = Path(ruta)
    res = ResultadoLectura()
    texto = None
    for cod in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            texto = ruta.read_text(encoding=cod)
            break
        except UnicodeDecodeError:
            continue
    if texto is None:
        res.avisos.append("No se pudo leer el archivo CSV (codificación desconocida).")
        return res
    try:
        dialecto = csv.Sniffer().sniff(texto[:4096], delimiters=";,\t|")
    except csv.Error:
        dialecto = csv.excel
        dialecto.delimiter = ";"
    filas = [tuple(f) for f in csv.reader(texto.splitlines(), dialecto)]
    _procesar_tabla(ruta.name, filas, res)
    return res


def leer_texto_tabla(texto: str, nombre: str = "Portapapeles") -> ResultadoLectura:
    """Interpreta celdas copiadas de Excel (texto separado por tabuladores).

    Si las filas copiadas incluyen las cabeceras se usan para saber qué es cada columna. Si no,
    se asume el orden de columnas de la plantilla (PERSONA, USUARIO PC, …), ignorando la columna
    separadora vacía que la plantilla tiene entre los dos bloques.
    """
    res = ResultadoLectura()
    lineas = [l for l in texto.replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    while lineas and not lineas[-1].strip():
        lineas.pop()
    filas = [tuple(c.strip() for c in l.split("\t")) for l in lineas]
    if not filas:
        res.avisos.append("El portapapeles está vacío.")
        return res
    cabeceras = max((len(_mapear_fila(f)) for f in filas[:30]), default=0)
    if cabeceras >= 2:
        _procesar_tabla(nombre, filas, res)
        return res

    ancho = max(len(f) for f in filas)
    columnas = list(range(ancho))
    # columna separadora de la plantilla (E): vacía en todas las filas
    if ancho >= 6 and all((len(f) <= 4 or not f[4]) for f in filas):
        columnas.pop(4)
    if ancho - (1 if len(columnas) < ancho else 0) > len(CLAVES):
        res.avisos.append(f"Se han copiado más columnas ({ancho}) de las que tiene el inventario; "
                          "las sobrantes se ignoran.")
    res.hojas_leidas.append(nombre)
    usadas = list(zip(columnas, CLAVES))
    res.columnas_reconocidas[nombre] = [CAMPOS_POR_CLAVE[c].etiqueta for _, c in usadas]
    res.avisos.append("No se han copiado las cabeceras: se asume el orden de columnas de la plantilla. "
                      "Revisa la vista previa.")
    for n, fila in enumerate(filas, start=1):
        datos = {c: "" for c in CLAVES}
        for i, clave in usadas:
            if i < len(fila):
                datos[clave] = _convertir(clave, fila[i])
        if any(datos.values()):
            datos["_origen"] = f"{nombre}!{n}"
            res.filas.append(datos)
    return res


# ---------------------------------------------------------------- exportación

_AZUL = PatternFill("solid", fgColor="1F4E78")
_AZUL_CLARO = PatternFill("solid", fgColor="DDEBF7")
_GRIS = PatternFill("solid", fgColor="EDEDED")
_ROJO = PatternFill("solid", fgColor="F8CBAD")
_AMARILLO = PatternFill("solid", fgColor="FFF2CC")
_FINO = Side(style="thin", color="808080")
_MEDIO = Side(style="medium", color="000000")
_BORDE = Border(left=_FINO, right=_FINO, top=_FINO, bottom=_FINO)
_BORDE_CAB = Border(left=_MEDIO, right=_MEDIO, top=_MEDIO, bottom=_MEDIO)
_CENTRO = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _disposicion() -> list[tuple[int, str]]:
    """Columnas de la hoja: lista (columna_excel, clave). Deja una columna separadora entre grupos,
    igual que la plantilla original (A-D propietario, E separador, F-S equipo)."""
    cols, col, grupo_ant = [], 1, None
    for c in CAMPOS:
        if grupo_ant is not None and c.grupo != grupo_ant:
            col += 1  # separador
        cols.append((col, c.clave))
        col += 1
        grupo_ant = c.grupo
    return cols


def exportar_excel(ruta: str | Path, equipos: list[dict], titulo_hoja: str = "Inventario") -> Path:
    ruta = Path(ruta)
    wb = Workbook()
    ws = wb.active
    ws.title = titulo_hoja[:31]
    disp = _disposicion()

    # Filas 1-2: cabeceras de grupo
    por_grupo: dict[str, list[int]] = {}
    for col, clave in disp:
        por_grupo.setdefault(CAMPOS_POR_CLAVE[clave].grupo, []).append(col)
    for grupo, cols in por_grupo.items():
        a, b = min(cols), max(cols)
        ws.merge_cells(start_row=1, start_column=a, end_row=2, end_column=b)
        celda = ws.cell(row=1, column=a, value=GRUPOS[grupo])
        celda.font = Font(bold=True, color="FFFFFF", size=12)
        celda.fill = _AZUL
        celda.alignment = _CENTRO
        for r in (1, 2):
            for c in range(a, b + 1):
                ws.cell(row=r, column=c).border = _BORDE_CAB

    # Filas 3-4: cabeceras de columna
    for col, clave in disp:
        campo = CAMPOS_POR_CLAVE[clave]
        ws.merge_cells(start_row=3, start_column=col, end_row=4, end_column=col)
        celda = ws.cell(row=3, column=col, value=campo.cabecera)
        celda.font = Font(bold=True)
        celda.fill = _AZUL_CLARO
        celda.alignment = _CENTRO
        for r in (3, 4):
            ws.cell(row=r, column=col).border = _BORDE_CAB
        ancho = max(10, min(40, campo.ancho / 7))
        ws.column_dimensions[get_column_letter(col)].width = ancho

    # separadores estrechos
    usadas = {c for c, _ in disp}
    for c in range(1, max(usadas) + 1):
        if c not in usadas:
            ws.column_dimensions[get_column_letter(c)].width = 1.5

    # Datos desde la fila 5
    hoy = date.today()
    for n, eq in enumerate(equipos):
        fila = 5 + n
        for col, clave in disp:
            campo = CAMPOS_POR_CLAVE[clave]
            valor = eq.get(clave, "") or ""
            if campo.tipo == "fecha":
                d = fecha_a_date(valor)
                celda = ws.cell(row=fila, column=col, value=d if d else (valor or None))
                if d:
                    celda.number_format = "DD/MM/YYYY"
                if clave == "fin_garantia":
                    est = estado_garantia(valor, hoy)
                    if est == "caducada":
                        celda.fill = _ROJO
                    elif est == "proxima":
                        celda.fill = _AMARILLO
            else:
                celda = ws.cell(row=fila, column=col, value=valor or None)
                if n % 2:
                    celda.fill = _GRIS
            celda.border = _BORDE
            celda.alignment = Alignment(vertical="center", wrap_text=campo.tipo == "largo")

    ultima_col = get_column_letter(max(usadas))
    ultima_fila = max(5, 4 + len(equipos))
    ws.auto_filter.ref = f"A4:{ultima_col}{ultima_fila}"
    ws.freeze_panes = "B5"
    ws.sheet_view.zoomScale = 90
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.print_title_rows = "1:4"

    wb.save(ruta)
    return ruta


def exportar_csv(ruta: str | Path, equipos: list[dict]) -> Path:
    ruta = Path(ruta)
    with ruta.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow([c.cabecera for c in CAMPOS])
        for eq in equipos:
            fila = []
            for c in CAMPOS:
                v = eq.get(c.clave, "") or ""
                if c.tipo == "fecha":
                    d = fecha_a_date(v)
                    v = d.strftime("%d/%m/%Y") if d else v
                fila.append(v)
            w.writerow(fila)
    return ruta


def exportar_plantilla(ruta: str | Path) -> Path:
    return exportar_excel(ruta, [], "Inventario")
