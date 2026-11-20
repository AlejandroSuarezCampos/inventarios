import shutil
from datetime import date, datetime, timedelta
from pathlib import Path

import pytest
from openpyxl import load_workbook

from inventario.campos import clave_para_cabecera, estado_garantia, parsear_fecha
from inventario.db import BaseDatos
from inventario.excel_io import exportar_csv, exportar_excel, leer_excel

FIXTURE = Path(__file__).parent / "fixtures" / "plantilla_original.xlsx"


@pytest.fixture
def excel_usuario(tmp_path):
    """Plantilla original del usuario con algunas filas de datos a partir de la fila 5."""
    destino = tmp_path / "Libro1.xlsx"
    shutil.copy(FIXTURE, destino)
    wb = load_workbook(destino)
    ws = wb.active
    ws["A5"], ws["B5"], ws["C5"], ws["D5"] = "Ana López", "alopez", 123456789, "Madrid"
    ws["F5"], ws["G5"], ws["H5"], ws["I5"] = "PC-MAD-01", "Dell", "Latitude 5420", "Windows 11 Pro"
    ws["J5"], ws["K5"], ws["L5"], ws["M5"] = "i5-1145G7", "16 GB", "DDR4", "512 GB SSD"
    ws["N5"], ws["O5"], ws["P5"], ws["Q5"] = "ABC123", '14"', "M365", "Defender"
    ws["R5"], ws["S5"] = datetime(2022, 3, 1), datetime(2025, 3, 1)
    ws["A6"], ws["F6"], ws["N6"], ws["R6"], ws["S6"] = "Luis Pérez", "PC-BCN-02", "XYZ789", "15/06/2024", "15/06/2027"
    wb.save(destino)
    return destino


def test_cabeceras_originales_reconocidas():
    wb = load_workbook(FIXTURE)
    fila = [c.value for c in wb.active[3]]
    claves = [clave_para_cabecera(v) for v in fila if v]
    assert None not in claves
    assert len(claves) == 18
    assert "nombre_equipo" in claves  # 'Nombre quipo'
    assert "fin_garantia" in claves   # 'EXPIACION GARANTIA'


def test_leer_excel_usuario(excel_usuario):
    res = leer_excel(excel_usuario)
    assert len(res.filas) == 2
    a, b = res.filas
    assert a["persona"] == "Ana López"
    assert a["teamviewer"] == "123456789"
    assert a["nombre_equipo"] == "PC-MAD-01"
    assert a["fecha_compra"] == "2022-03-01"
    assert a["fin_garantia"] == "2025-03-01"
    assert b["fecha_compra"] == "2024-06-15"
    assert b["num_serie"] == "XYZ789"


def test_plantilla_vacia_sin_filas():
    assert leer_excel(FIXTURE).filas == []


def test_ida_y_vuelta(tmp_path, excel_usuario):
    bd = BaseDatos(tmp_path / "t.db")
    for f in leer_excel(excel_usuario).filas:
        bd.insertar({k: v for k, v in f.items() if not k.startswith("_")})
    assert bd.contar() == 2
    salida = tmp_path / "export.xlsx"
    exportar_excel(salida, bd.listar())
    vuelta = leer_excel(salida).filas
    assert len(vuelta) == 2
    assert vuelta[0]["num_serie"] == "ABC123"
    assert vuelta[0]["fin_garantia"] == "2025-03-01"
    assert vuelta[0]["estado"] == "Activo"
    exportar_csv(tmp_path / "e.csv", bd.listar())
    assert len(leer_excel(tmp_path / "e.csv").filas) == 2


def test_duplicados_y_actualizacion(tmp_path):
    bd = BaseDatos(tmp_path / "t.db")
    i = bd.insertar({"num_serie": "abc-123", "persona": "A"})
    assert bd.buscar_duplicado({"num_serie": "ABC 123"})["id"] == i
    assert bd.buscar_duplicado({"num_serie": "OTRO", "nombre_equipo": "X"}) is None
    assert bd.actualizar(i, {"num_serie": "abc-123", "persona": "B"})
    assert bd.obtener(i)["persona"] == "B"
    assert any(h["accion"] == "Modificación" for h in bd.historial())
    bd.eliminar([i])
    assert bd.contar() == 0


def test_copia_y_restaurar(tmp_path):
    bd = BaseDatos(tmp_path / "t.db")
    bd.insertar({"persona": "A"})
    bd.copia_seguridad(tmp_path / "copia.db")
    bd.insertar({"persona": "B"})
    bd.restaurar(tmp_path / "copia.db")
    assert bd.contar() == 1


def test_fechas():
    assert parsear_fecha("01/02/2023") == "2023-02-01"
    assert parsear_fecha(44927) == "2023-01-01"
    assert parsear_fecha("") == ""
    assert parsear_fecha("pendiente") == "pendiente"
    hoy = date(2025, 1, 1)
    assert estado_garantia("2024-12-31", hoy) == "caducada"
    assert estado_garantia((hoy + timedelta(days=30)).isoformat(), hoy) == "proxima"
    assert estado_garantia("2030-01-01", hoy) == "vigente"
    assert estado_garantia("", hoy) == "sin_dato"
