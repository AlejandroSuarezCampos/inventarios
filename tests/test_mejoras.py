"""Tests de las mejoras: deshacer, varias personas, pegar, PDF, QR, detección y periodos."""
from datetime import date

import pytest

from inventario.db import BaseDatos, Conflicto
from inventario.detectar import interpretar
from inventario.excel_io import leer_texto_tabla
from inventario.informes import (FORMATOS_ETIQUETA, generar_etiquetas, generar_informe, id_desde_qr,
                                 matriz_qr, periodos_garantia, texto_qr)
from inventario.pdf import PDF


@pytest.fixture
def bd(tmp_path):
    return BaseDatos(tmp_path / "t.db")


def test_deshacer_y_rehacer(bd):
    i = bd.insertar({"persona": "Ana"})
    bd.actualizar_campo(i, "persona", "Luis")
    bd.eliminar([i])
    assert bd.contar() == 0
    assert bd.deshacer()[0].startswith("Eliminar")
    assert bd.obtener(i)["persona"] == "Luis"
    bd.deshacer()
    assert bd.obtener(i)["persona"] == "Ana"
    bd.rehacer()
    assert bd.obtener(i)["persona"] == "Luis"
    bd.deshacer()
    bd.deshacer()  # deshace el alta
    assert bd.contar() == 0
    assert bd.deshacer() is None


def test_operacion_agrupada(bd):
    with bd.operacion("Importar"):
        for k in range(5):
            bd.insertar({"persona": f"p{k}"})
    assert bd.contar() == 5 and bd.nombre_deshacer() == "Importar"
    bd.deshacer()
    assert bd.contar() == 0


def test_dos_personas_fusion_y_conflicto(tmp_path):
    a, b = BaseDatos(tmp_path / "x.db"), BaseDatos(tmp_path / "x.db")
    a.usuario, b.usuario = "ana", "luis"
    i = a.insertar({"persona": "P", "marca": "Dell", "ram": "8 GB"})
    base = b.obtener(i)
    a.actualizar(i, {**base, "marca": "HP"}, base=base)          # Ana cambia la marca
    b.actualizar(i, {**base, "ram": "16 GB"}, base=base)         # Luis cambia la RAM: se fusiona
    assert a.obtener(i)["marca"] == "HP" and a.obtener(i)["ram"] == "16 GB"
    with pytest.raises(Conflicto) as exc:                        # Luis cambia la marca sin saberlo
        b.actualizar(i, {**base, "marca": "Lenovo"}, base=base)
    assert exc.value.campos[0][:3] == ("marca", "Dell", "HP")
    b.actualizar(i, {**base, "marca": "Lenovo"}, base=base, forzar=True)
    assert a.obtener(i)["marca"] == "Lenovo"
    assert {h["usuario"] for h in a.historial()} >= {"ana", "luis"}
    a.eliminar([i])
    with pytest.raises(Conflicto) as exc:
        b.actualizar_campo(i, "persona", "X", esperado="P")
    assert exc.value.eliminado


def test_deshacer_respeta_cambios_de_otros(tmp_path):
    a, b = BaseDatos(tmp_path / "x.db"), BaseDatos(tmp_path / "x.db")
    i = a.insertar({"persona": "P", "marca": "Dell"})
    a.actualizar_campo(i, "marca", "HP", esperado="Dell")
    b.actualizar_campo(i, "marca", "Acer", esperado="HP")        # otra persona lo cambia después
    _, avisos = a.deshacer()
    assert a.obtener(i)["marca"] == "Acer" and avisos


def test_historial_por_equipo_y_adjuntos(bd, tmp_path):
    i = bd.insertar({"persona": "A"})
    j = bd.insertar({"persona": "B"})
    bd.actualizar_campo(i, "ram", "8 GB")
    assert len(bd.historial(equipo_id=i)) == 2 and len(bd.historial(equipo_id=j)) == 1
    f = tmp_path / "factura.pdf"
    f.write_bytes(b"x")
    bd.anadir_adjunto(i, f)
    bd.anadir_adjunto(i, f)
    assert [p.name for p in bd.listar_adjuntos(i)] == ["factura (2).pdf", "factura.pdf"]
    bd.eliminar_adjunto(i, "factura.pdf")
    assert bd.contar_adjuntos(i) == 1


def test_pegar_desde_excel():
    # sin cabeceras, con la columna separadora vacía de la plantilla
    fila = "Ana\tana\t123\tMadrid\t\tPC-1\tDell\tLat\tW11\ti5\t8 GB\tDDR4\t256\tSN1\t14\tM365\tDef\t01/02/2023\t01/02/2026"
    r = leer_texto_tabla(fila + "\n")
    assert r.filas[0]["nombre_equipo"] == "PC-1" and r.filas[0]["fin_garantia"] == "2026-02-01"
    r = leer_texto_tabla("Nº de serie\tPersona\r\nX9\tLuis\r\n")
    assert r.filas[0]["num_serie"] == "X9" and r.filas[0]["persona"] == "Luis"


def test_qr():
    eq = {"id": 42, "nombre_equipo": "PC-1", "num_serie": "S1", "persona": "Ana"}
    assert texto_qr(eq).startswith("INV:42 | PC-1")
    assert id_desde_qr(texto_qr(eq)) == 42 and id_desde_qr("hola") is None
    assert len(matriz_qr(texto_qr(eq))) >= 21


def _equipos(n=30):
    return [{"id": k, "persona": f"Núñez {k}", "nombre_equipo": f"PC-{k}", "num_serie": f"S{k}", "marca": "Dell",
             "delegacion": "Madrid", "fin_garantia": "2020-01-01" if k % 3 else "2099-01-01", "estado": "Activo",
             "fecha_compra": "2019-01-01"} for k in range(1, n + 1)]


def test_pdf_informe_y_etiquetas(tmp_path):
    pypdf = pytest.importorskip("pypdf")
    generar_informe(tmp_path / "i.pdf", _equipos(80), "prueba", "ana")
    texto = "".join(p.extract_text() for p in pypdf.PdfReader(tmp_path / "i.pdf").pages)
    assert "Informe de inventario" in texto and "Núñez 80" in texto
    for f in FORMATOS_ETIQUETA:
        generar_etiquetas(tmp_path / "e.pdf", _equipos(30), f, saltar=1)
        paginas = len(pypdf.PdfReader(tmp_path / "e.pdf").pages)
        assert paginas == -(-31 // (f.columnas * f.filas))


def test_pdf_ancho_texto():
    assert PDF.ancho_texto("WWW", 10) > PDF.ancho_texto("iii", 10)
    assert PDF().recortar("texto muy largo de prueba", 40, 10).endswith("…")


def test_detectar_interpretar():
    r = interpretar({"nombre": "PC1", "fabricante": "LENOVO", "modelo": "20W0", "modelo_version": "ThinkPad T14",
                     "serie": "To be filled by O.E.M.", "os": "Microsoft Windows 11 Pro", "os_display": "24H2",
                     "cpu": "Intel(R) Core(TM) i7-1185G7 @ 3.00GHz", "ram_modulos": ["17179869184"],
                     "ram_tipos": [34], "discos": [{"nombre": "WD", "tipo": "3", "tam": "1000204886016", "bus": "11"}],
                     "monitores": [{"h": 34, "v": 19}], "antivirus": ["Windows Defender", "ESET"],
                     "office_c2r": "ProPlus2021Volume", "teamviewer": None})
    assert r["marca"] == "Lenovo" and r["modelo"] == "ThinkPad T14" and r["os"] == "Windows 11 Pro 24H2"
    assert r["cpu"] == "Intel Core i7-1185G7" and r["ram"] == "16 GB" and r["tipo_ram"] == "DDR5"
    assert r["almacenamiento"] == "1 TB HDD" and r["pantalla"] == '15.6"'
    assert r["office"] == "Office 2021 Pro Plus" and r["antivirus"] == "Microsoft Defender, ESET"
    assert "num_serie" not in r and "teamviewer" not in r


def test_periodos_garantia():
    p = periodos_garantia(date(2026, 11, 20))
    nombres = [n for n, _, _ in p]
    assert any("T4 2026" in n for n in nombres)
    trimestre = next(x for x in p if "T4" in x[0])
    assert trimestre[1] == date(2026, 10, 1) and trimestre[2] == date(2026, 12, 31)
    proximo = next(x for x in p if x[0] == "Próximo trimestre")
    assert proximo[1] == date(2027, 1, 1) and proximo[2] == date(2027, 3, 31)
