"""Servidor central + clientes con modo sin conexión."""
import socket
import time

import pytest

from inventario.db import Conflicto, SinConexion
from inventario.remoto import ClienteRemoto, ErrorServidor, esperar_conexion, probar_conexion
from inventario.servidor import ServidorInventario


def _puerto_libre() -> int:
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    p = s.getsockname()[1]
    s.close()
    return p


@pytest.fixture
def entorno(tmp_path, monkeypatch):
    monkeypatch.setenv("XDG_DATA_HOME", str(tmp_path / "cliente"))
    monkeypatch.setenv("APPDATA", str(tmp_path / "cliente"))
    puerto = _puerto_libre()
    srv = ServidorInventario(tmp_path / "servidor", puerto=puerto, clave="secreta")
    srv.iniciar()
    clientes = []

    def cliente(**kw):
        c = ClienteRemoto(f"127.0.0.1:{puerto}", "secreta", intervalo=0.2, **kw)
        c.cliente = f"cliente{len(clientes)}"   # en la realidad, cada PC tiene su propio identificador
        clientes.append(c)
        return c
    yield srv, cliente, puerto, tmp_path
    for c in clientes:
        c.cerrar()
    try:
        srv.detener()
    except Exception:  # noqa: BLE001
        pass


def _esperar_desconexion(c, segundos=5):
    fin = time.time() + segundos
    while time.time() < fin:
        c.version_datos()
        if not c.en_linea:
            return True
        time.sleep(0.1)
    return False


def test_con_conexion_y_dos_clientes(entorno):
    srv, cliente, puerto, _ = entorno
    assert probar_conexion(f"127.0.0.1:{puerto}", "secreta")["app"] == "InventarioPC"
    a, b = cliente(), cliente()
    a.usuario, b.usuario = "ana", "luis"
    i = a.insertar({"persona": "Ana", "marca": "Dell"})
    assert i > 0 and srv.bd.obtener(i)["persona"] == "Ana"
    b.sincronizar()
    assert b.obtener(i)["marca"] == "Dell"
    base = b.obtener(i)
    a.actualizar_campo(i, "marca", "HP", esperado="Dell")
    with pytest.raises(Conflicto):                              # Luis no ha visto el cambio de Ana
        b.actualizar(i, {**base, "marca": "Lenovo"}, base=base)
    b.actualizar(i, {**base, "ram": "16 GB"}, base=base)        # otro campo: se fusiona
    assert srv.bd.obtener(i)["marca"] == "HP" and srv.bd.obtener(i)["ram"] == "16 GB"
    assert {h["usuario"] for h in srv.bd.historial()} >= {"ana", "luis"}
    # deshacer: cada cliente deshace lo suyo
    assert a.nombre_deshacer()
    a.deshacer()
    assert srv.bd.obtener(i)["marca"] == "Dell"
    with a.operacion("Importar"):
        for k in range(3):
            a.insertar({"persona": f"p{k}"})
    assert srv.bd.contar() == 4 and len(a.ultimos_ids) == 3
    a.deshacer()
    assert srv.bd.contar() == 1


def test_sin_conexion_consultar_y_crear(entorno):
    srv, cliente, puerto, tmp_path = entorno
    c = cliente()
    i = c.insertar({"persona": "Del servidor", "num_serie": "S1"})
    srv.detener()
    assert _esperar_desconexion(c)
    # consultar lo que había
    assert c.obtener(i)["persona"] == "Del servidor" and not c.puede_modificar(i)
    # no se puede modificar ni borrar lo del servidor
    with pytest.raises(SinConexion):
        c.actualizar_campo(i, "persona", "X", esperado="Del servidor")
    with pytest.raises(SinConexion):
        c.eliminar([i])
    # sí se pueden crear equipos (y retocarlos mientras están pendientes)
    p1 = c.insertar({"persona": "Nuevo 1", "num_serie": "S1"})   # mismo nº de serie: no hay clave
    p2 = c.insertar({"persona": "Nuevo 2"})
    assert p1 < 0 and p2 < 0 and c.pendientes == 2 and c.puede_modificar(p1)
    c.actualizar_campo(p2, "ram", "8 GB")
    c.eliminar([p2])
    with c.operacion("Importar sin conexión"):
        c.insertar({"persona": "Importado"})
    f = tmp_path / "factura.pdf"
    f.write_bytes(b"%PDF")
    c.anadir_adjunto(p1, f)
    assert {e["persona"] for e in c.listar()} == {"Del servidor", "Nuevo 1", "Importado"}
    # vuelve el servidor (mismo puerto y carpeta): se envían los pendientes
    srv2 = ServidorInventario(tmp_path / "servidor", puerto=puerto, clave="secreta")
    srv2.iniciar()
    try:
        assert esperar_conexion(c)
        assert c.pendientes == 0
        personas = sorted(e["persona"] for e in srv2.bd.listar())
        assert personas == ["Del servidor", "Importado", "Nuevo 1"]
        nuevo = next(e for e in srv2.bd.listar() if e["persona"] == "Nuevo 1")
        assert [a["nombre"] for a in srv2.bd.info_adjuntos(nuevo["id"])] == ["factura.pdf"]
        assert all(e["id"] > 0 for e in c.listar())
        assert any(h["accion"] == "Sincronización" for h in srv2.bd.historial())
    finally:
        srv2.detener()


def test_arranque_sin_servidor_usa_la_copia_local(entorno):
    srv, cliente, puerto, _ = entorno
    c = cliente()
    c.insertar({"persona": "Guardado"})
    c.cerrar()
    srv.detener()
    c2 = cliente()                                   # se abre sin servidor
    assert not c2.en_linea and [e["persona"] for e in c2.listar()] == ["Guardado"]
    assert c2.ultima_sincronizacion


def test_adjuntos_por_el_servidor(entorno, tmp_path):
    srv, cliente, _, _ = entorno
    c = cliente()
    i = c.insertar({"persona": "A"})
    f = tmp_path / "foto.jpg"
    f.write_bytes(b"123456")
    c.anadir_adjunto(i, f)
    assert [a["nombre"] for a in c.info_adjuntos(i)] == ["foto.jpg"]
    assert c.ruta_adjunto(i, "foto.jpg").read_bytes() == b"123456"
    c.eliminar_adjunto(i, "foto.jpg")
    assert c.info_adjuntos(i) == []


def test_clave_incorrecta_y_duplicados(entorno):
    srv, cliente, puerto, _ = entorno
    with pytest.raises(ErrorServidor):
        probar_conexion(f"127.0.0.1:{puerto}", "otra")
    c = cliente()
    ops = {"nombre": "x", "ops": [{"op": "insertar", "datos": {"persona": "Z"}, "uuid": "abc"}]}
    r1 = c._api("POST", "/api/lote", ops)["resultados"]
    r2 = c._api("POST", "/api/lote", ops)["resultados"]    # reintento tras un corte: no se duplica
    assert r1 == r2 and srv.bd.contar() == 1


def test_solo_misma_red_servidor(entorno):
    """El servidor rechaza equipos de otra red; el cliente sigue en modo sin conexión y sincroniza al volver."""
    from inventario.remoto import FueraDeRed
    srv, cliente, puerto, _ = entorno
    c = cliente()
    i = c.insertar({"persona": "Antes"})
    srv.permitir_este_equipo = False          # simula que el cliente está en otra red
    with pytest.raises(FueraDeRed):
        probar_conexion(f"127.0.0.1:{puerto}", "secreta")
    assert _esperar_desconexion(c) and c.fuera_de_red
    with pytest.raises(FueraDeRed):
        c.actualizar_campo(i, "persona", "X", esperado="Antes")
    p = c.insertar({"persona": "Creado fuera de la red"})
    assert p < 0 and srv.bd.contar() == 1      # no ha llegado nada al servidor
    assert any("red local" in t for _, _, t in srv.actividad)
    srv.permitir_este_equipo = True           # vuelve a la red
    assert esperar_conexion(c) and not c.fuera_de_red and c.pendientes == 0
    assert sorted(e["persona"] for e in srv.bd.listar()) == ["Antes", "Creado fuera de la red"]


def test_redes_del_servidor_y_del_cliente(entorno):
    import ipaddress
    from inventario.remoto import FueraDeRed, comprobar_misma_red
    srv, *_ = entorno
    srv.redes.fijar([ipaddress.IPv4Network("192.168.1.0/24")])
    assert srv.ip_permitida("192.168.1.50") and srv.ip_permitida("::ffff:192.168.1.7")
    assert not srv.ip_permitida("192.168.2.5") and not srv.ip_permitida("8.8.8.8")
    assert not srv.ip_permitida("10.8.0.6")   # p. ej. un cliente por VPN con otra red
    redes = [ipaddress.IPv4Network("192.168.1.0/24")]
    comprobar_misma_red("http://192.168.1.20:8765", redes)
    with pytest.raises(FueraDeRed):
        comprobar_misma_red("http://10.0.0.5:8765", redes)
    with pytest.raises(FueraDeRed):
        comprobar_misma_red("http://203.0.113.9:8765", redes)
    comprobar_misma_red("http://127.0.0.1:8765", redes)   # el propio equipo


def test_interpretar_redes():
    from inventario.red import interpretar_ip_addr, interpretar_ipconfig
    es = ("   Dirección IPv4. . . . . . . . . . . . . . : 192.168.1.20(Preferido)\n"
          "   Máscara de subred . . . . . . . . . . . . : 255.255.255.0\n"
          "   Puerta de enlace predeterminada . . . . . : 192.168.1.1\n"
          "   IPv4 Address. . . . . . . . . . . : 10.20.0.15\n"
          "   Subnet Mask . . . . . . . . . . . : 255.255.0.0\n")
    assert [str(r) for r in interpretar_ipconfig(es)] == ["192.168.1.0/24", "10.20.0.0/16"]
    assert [str(r) for r in interpretar_ip_addr("inet 172.16.5.4/20 brd 172.16.15.255")] == ["172.16.0.0/20"]
