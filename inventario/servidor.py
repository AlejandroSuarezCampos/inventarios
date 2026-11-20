"""Servidor central del inventario.

Se ejecuta con ``InventarioPC.exe --servidor`` en el equipo que va a guardar los datos. Los demás
equipos se conectan a él (Archivo → Conectar a un servidor). Es un servicio HTTP con JSON, sin
dependencias, protegido con una clave compartida que se envía en la cabecera ``X-Clave``.

Todos los datos (base de datos, adjuntos y copias) quedan en la carpeta del servidor.
"""
from __future__ import annotations

import hmac
import json
import secrets
import shutil
import socket
import tempfile
import threading
import time
import traceback
from collections import deque
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

from . import __version__
from .db import BaseDatos, Conflicto, carpeta_datos
from .red import RedesCacheadas, describir, en_alguna_red

PUERTO_POR_DEFECTO = 8765
TAMANO_MAXIMO = 300 * 1024 * 1024  # 300 MB por petición (adjuntos)


def carpeta_servidor() -> Path:
    ruta = carpeta_datos() / "servidor"
    ruta.mkdir(parents=True, exist_ok=True)
    return ruta


def direcciones_locales() -> list[str]:
    """IP de este equipo en la red local (para decírsela a los clientes)."""
    ips: list[str] = []
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("10.255.255.255", 1))
        ips.append(s.getsockname()[0])
        s.close()
    except OSError:
        pass
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            ip = info[4][0]
            if ip not in ips and not ip.startswith("127."):
                ips.append(ip)
    except OSError:
        pass
    return ips or ["127.0.0.1"]


def _conflicto_json(c: Conflicto) -> dict:
    return {"equipo_id": c.equipo_id, "eliminado": c.eliminado, "campos": [list(x) for x in c.campos]}


class ServidorInventario:
    def __init__(self, carpeta: str | Path | None = None, puerto: int | None = None, clave: str | None = None):
        self.carpeta = Path(carpeta) if carpeta else carpeta_servidor()
        self.carpeta.mkdir(parents=True, exist_ok=True)
        self.ruta_cfg = self.carpeta / "servidor.json"
        cfg = self._leer_cfg()
        self.puerto = int(puerto if puerto is not None else (cfg.get("puerto") or PUERTO_POR_DEFECTO))
        self.clave = clave or cfg.get("clave") or secrets.token_urlsafe(9)
        self._guardar_cfg({**cfg, "puerto": self.puerto, "clave": self.clave})
        self.bd = BaseDatos(self.carpeta / "inventario.db", varios_hilos=True)
        self.bd.usuario = "servidor"
        # equipos ya recibidos de clientes que trabajaron sin conexión (evita duplicados si se reintenta)
        self.bd.con.execute("CREATE TABLE IF NOT EXISTS sincronizados (uuid TEXT PRIMARY KEY, equipo_id INTEGER)")
        self.bd.con.commit()
        self.lock = threading.RLock()
        self.pilas: dict[str, tuple[list, list]] = {}
        self.clientes: dict[str, dict] = {}
        self.actividad: deque = deque(maxlen=300)
        # Sólo se atiende a equipos de la misma red local que el servidor
        self.redes = RedesCacheadas(60)
        self.permitir_este_equipo = True
        self.rechazados: dict[str, float] = {}
        self.httpd: ThreadingHTTPServer | None = None
        self._hilo: threading.Thread | None = None
        self._parar = threading.Event()

    # -------------------------------------------------------------- configuración
    def _leer_cfg(self) -> dict:
        try:
            return json.loads(self.ruta_cfg.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return {}

    def _guardar_cfg(self, cfg: dict) -> None:
        self.ruta_cfg.write_text(json.dumps(cfg, indent=2), encoding="utf-8")

    def nueva_clave(self) -> str:
        self.clave = secrets.token_urlsafe(9)
        self._guardar_cfg({**self._leer_cfg(), "clave": self.clave})
        return self.clave

    # -------------------------------------------------------------- ciclo de vida
    def iniciar(self) -> None:
        servidor = self

        class Manejador(_Manejador):
            srv = servidor

        self.httpd = ThreadingHTTPServer(("0.0.0.0", self.puerto), Manejador)
        self.httpd.daemon_threads = True
        self.puerto = self.httpd.server_address[1]
        self._hilo = threading.Thread(target=self.httpd.serve_forever, daemon=True)
        self._hilo.start()
        threading.Thread(target=self._copias_periodicas, daemon=True).start()
        self.anotar("servidor", f"En marcha en el puerto {self.puerto}")

    def detener(self) -> None:
        self._parar.set()
        if self.httpd:
            self.httpd.shutdown()
            self.httpd.server_close()
        with self.lock:
            self.bd.commit()

    def _copias_periodicas(self) -> None:
        while not self._parar.wait(5):
            with self.lock:
                try:
                    self.bd.copia_automatica(maximo=30)
                except Exception:  # noqa: BLE001
                    pass
            self._parar.wait(3600)

    def ip_permitida(self, ip: str) -> bool:
        """¿La petición viene de un equipo de la misma red local que el servidor?"""
        ip = ip.split("%")[0]
        if ip.startswith("::ffff:"):
            ip = ip[7:]
        if ip.startswith("127.") or ip == "::1":
            return self.permitir_este_equipo
        redes = [r for r in self.redes.obtener() if not r.is_loopback]
        return en_alguna_red(ip, redes)

    def redes_permitidas(self) -> str:
        return describir(self.redes.obtener())

    def anotar(self, quien: str, texto: str) -> None:
        self.actividad.appendleft((datetime.now().strftime("%H:%M:%S"), quien, texto))

    # -------------------------------------------------------------- operaciones
    def ejecutar(self, metodo: str, ruta: str, consulta: dict, cuerpo: bytes, cliente: str, usuario: str):
        """Devuelve (código, objeto JSON | bytes, tipo)."""
        partes = [unquote(p) for p in ruta.strip("/").split("/")]
        if partes[:1] != ["api"]:
            return 404, {"error": "no encontrado"}, None
        partes = partes[1:]
        datos = {}
        if cuerpo and metodo in ("POST", "PUT") and not (partes[:1] == ["adjuntos"] and metodo == "POST"):
            datos = json.loads(cuerpo.decode("utf-8"))
        bd = self.bd
        with self.lock:
            bd.usuario = usuario or "desconocido"
            bd._deshacer, bd._rehacer = self.pilas.setdefault(cliente, ([], []))
            try:
                return self._despachar(metodo, partes, consulta, datos, cuerpo)
            except Conflicto as c:
                return 409, {"conflicto": _conflicto_json(c)}, None
            finally:
                bd.usuario = "servidor"

    def _despachar(self, metodo, partes, consulta, datos, cuerpo):
        bd = self.bd
        q = lambda k, d=None: (consulta.get(k) or [d])[0]  # noqa: E731
        if metodo == "GET" and partes == ["estado"]:
            return 200, {"app": "InventarioPC", "version": __version__, "revision": bd.revision(),
                         "equipos": bd.contar(), "servidor": socket.gethostname(),
                         "deshacer": bd.nombre_deshacer(), "rehacer": bd.nombre_rehacer()}, None
        if metodo == "GET" and partes == ["datos"]:
            return 200, {"revision": bd.revision(), "equipos": bd.listar(), "historial": bd.historial(3000),
                         "adjuntos": bd.todos_adjuntos()}, None
        if metodo == "GET" and partes == ["historial"]:
            eid = q("equipo_id")
            return 200, bd.historial(int(q("limite", 2000)), int(eid) if eid else None), None
        if metodo == "DELETE" and partes == ["historial"]:
            bd.vaciar_historial()
            return 200, {"ok": True}, None
        if metodo == "POST" and partes == ["registrar"]:
            bd.registrar(datos.get("equipo_id"), datos.get("equipo", ""), datos.get("accion", ""), datos.get("detalle", ""))
            bd.commit()
            return 200, {"ok": True}, None
        if metodo == "POST" and partes == ["equipos"]:
            i = bd.insertar(datos["datos"])
            return 200, {"id": i}, None
        if metodo == "PUT" and len(partes) == 2 and partes[0] == "equipos":
            cambiado = bd.actualizar(int(partes[1]), datos["datos"], base=datos.get("base"),
                                     forzar=bool(datos.get("forzar")))
            return 200, {"cambiado": cambiado}, None
        if metodo == "POST" and partes == ["eliminar"]:
            return 200, {"n": bd.eliminar([int(i) for i in datos["ids"]])}, None
        if metodo == "POST" and partes == ["lote"]:
            return 200, {"resultados": self._lote(datos.get("nombre", "Operación"), datos.get("ops", []))}, None
        if metodo == "POST" and partes in (["deshacer"], ["rehacer"]):
            res = bd.deshacer() if partes == ["deshacer"] else bd.rehacer()
            return 200, ({"nombre": res[0], "avisos": res[1]} if res else {"nombre": None, "avisos": []}), None
        if partes[:1] == ["adjuntos"]:
            return self._adjuntos(metodo, partes[1:], consulta, cuerpo)
        if metodo == "GET" and partes == ["copia"]:
            tmp = Path(tempfile.mkdtemp()) / "copia.db"
            bd.copia_seguridad(tmp)
            contenido = tmp.read_bytes()
            shutil.rmtree(tmp.parent, ignore_errors=True)
            return 200, contenido, "application/octet-stream"
        return 404, {"error": "no encontrado"}, None

    def _lote(self, nombre: str, ops: list[dict]) -> list:
        bd = self.bd
        resultados = []
        with bd.operacion(nombre):
            for op in ops:
                tipo = op.get("op")
                if tipo == "insertar":
                    uuid = op.get("uuid")
                    previo = bd.con.execute("SELECT equipo_id FROM sincronizados WHERE uuid=?", (uuid,)).fetchone() \
                        if uuid else None
                    if previo:
                        resultados.append(previo[0])
                        continue
                    nuevo = bd.insertar(op["datos"])
                    if uuid:
                        bd.con.execute("INSERT INTO sincronizados (uuid, equipo_id) VALUES (?, ?)", (uuid, nuevo))
                    resultados.append(nuevo)
                elif tipo == "actualizar":
                    resultados.append(bd.actualizar(int(op["id"]), op["datos"], base=op.get("base"),
                                                    forzar=bool(op.get("forzar"))))
                elif tipo == "actualizar_campo":
                    resultados.append(bd.actualizar_campo(int(op["id"]), op["clave"], op["valor"],
                                                          esperado=op.get("esperado"), forzar=bool(op.get("forzar"))))
                elif tipo == "eliminar":
                    resultados.append(bd.eliminar([int(i) for i in op["ids"]]))
                elif tipo == "vaciar":
                    resultados.append(bd.vaciar(op.get("motivo", "")))
                elif tipo == "registrar":
                    bd.registrar(op.get("equipo_id"), op.get("equipo", ""), op.get("accion", ""), op.get("detalle", ""))
                    resultados.append(True)
                else:
                    raise ValueError(f"Operación desconocida: {tipo}")
        return resultados

    def _adjuntos(self, metodo, partes, consulta, cuerpo):
        bd = self.bd
        if not partes or not partes[0].isdigit():
            return 404, {"error": "no encontrado"}, None
        eid = int(partes[0])
        if metodo == "GET" and len(partes) == 1:
            return 200, bd.info_adjuntos(eid), None
        if metodo == "GET" and len(partes) == 2:
            ruta = bd.ruta_adjunto(eid, Path(partes[1]).name)
            if not ruta.is_file():
                return 404, {"error": "no existe"}, None
            return 200, ruta.read_bytes(), "application/octet-stream"
        if metodo == "POST" and len(partes) == 1:
            nombre = Path((consulta.get("nombre") or ["adjunto"])[0]).name
            tmp = Path(tempfile.mkdtemp()) / nombre
            tmp.write_bytes(cuerpo)
            try:
                destino = bd.anadir_adjunto(eid, tmp)
            finally:
                shutil.rmtree(tmp.parent, ignore_errors=True)
            return 200, {"nombre": destino.name}, None
        if metodo == "DELETE" and len(partes) == 2:
            bd.eliminar_adjunto(eid, Path(partes[1]).name)
            return 200, {"ok": True}, None
        return 404, {"error": "no encontrado"}, None


def _describir(metodo: str, ruta: str, codigo: int, cuerpo: bytes) -> str:
    """Texto legible de una petición para la ventana del servidor."""
    if codigo == 409:
        return "Conflicto: otra persona había cambiado los mismos campos"
    if codigo >= 400:
        return f"Error al atender la petición ({codigo})"
    partes = ruta.strip("/").split("/")[1:]
    if partes == ["lote"]:
        try:
            nombre = json.loads(cuerpo.decode("utf-8")).get("nombre", "")
        except ValueError:
            nombre = ""
        return f"Guardó: {nombre}" if nombre else "Guardó cambios"
    textos = {("POST", "equipos"): "Dio de alta un equipo", ("PUT", "equipos"): "Modificó un equipo",
              ("POST", "eliminar"): "Eliminó equipos", ("POST", "deshacer"): "Deshizo su último cambio",
              ("POST", "rehacer"): "Rehízo un cambio", ("POST", "adjuntos"): "Añadió un adjunto",
              ("DELETE", "adjuntos"): "Quitó un adjunto", ("DELETE", "historial"): "Vació el historial",
              ("POST", "registrar"): "Exportó datos"}
    return textos.get((metodo, partes[0] if partes else ""), f"{metodo} {ruta}")


class _Manejador(BaseHTTPRequestHandler):
    srv: ServidorInventario
    protocol_version = "HTTP/1.1"
    server_version = "InventarioPC"

    def log_message(self, formato, *args):  # silenciar el registro por consola
        pass

    def _responder(self, codigo: int, contenido, tipo: str | None = None):
        if isinstance(contenido, (bytes, bytearray)):
            datos, tipo = bytes(contenido), tipo or "application/octet-stream"
        else:
            datos, tipo = json.dumps(contenido, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8"
        self.send_response(codigo)
        self.send_header("Content-Type", tipo)
        self.send_header("Content-Length", str(len(datos)))
        self.end_headers()
        self.wfile.write(datos)

    def _atender(self, metodo: str):
        srv = self.srv
        ip = self.client_address[0]
        if not srv.ip_permitida(ip):
            self._leer_cuerpo()
            ahora_ = time.time()
            if ahora_ - srv.rechazados.get(ip, 0) > 60:   # no llenar la actividad con reintentos
                srv.anotar(ip, "Rechazado: no está en la red local del servidor")
            srv.rechazados[ip] = ahora_
            return self._responder(403, {"error": "fuera_de_red", "redes": srv.redes_permitidas()})
        clave = self.headers.get("X-Clave", "")
        if not hmac.compare_digest(clave.encode(), srv.clave.encode()):
            self._leer_cuerpo()
            srv.anotar(self.client_address[0], "Clave incorrecta")
            return self._responder(401, {"error": "clave incorrecta"})
        cliente = self.headers.get("X-Cliente", self.client_address[0])
        usuario = unquote(self.headers.get("X-Usuario", ""))
        srv.clientes[cliente] = {"usuario": usuario, "ip": self.client_address[0], "ultima": time.time()}
        try:
            cuerpo = self._leer_cuerpo()
            url = urlparse(self.path)
            codigo, contenido, tipo = srv.ejecutar(metodo, url.path, parse_qs(url.query), cuerpo, cliente, usuario)
            if metodo != "GET" or codigo >= 400:
                srv.anotar(usuario or cliente, _describir(metodo, url.path, codigo, cuerpo))
            self._responder(codigo, contenido, tipo)
        except Exception as exc:  # noqa: BLE001
            srv.anotar(usuario or cliente, f"Error: {exc}")
            traceback.print_exc()
            self._responder(500, {"error": str(exc)})

    def _leer_cuerpo(self) -> bytes:
        n = int(self.headers.get("Content-Length") or 0)
        if n > TAMANO_MAXIMO:
            raise ValueError("Petición demasiado grande")
        return self.rfile.read(n) if n else b""

    def do_GET(self):
        self._atender("GET")

    def do_POST(self):
        self._atender("POST")

    def do_PUT(self):
        self._atender("PUT")

    def do_DELETE(self):
        self._atender("DELETE")
