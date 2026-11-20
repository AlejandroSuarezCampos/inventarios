"""Cliente del servidor central con modo sin conexión.

Ofrece la misma interfaz que :class:`inventario.db.BaseDatos`, pero los datos viven en el servidor.

* **Con conexión**: todas las lecturas y escrituras van al servidor; se guarda una copia local.
* **Sin conexión**: se consulta la copia local de la última sincronización (sólo lectura) y se pueden
  **crear equipos nuevos**, que quedan «pendientes» en este equipo. Al volver la conexión se envían
  al servidor automáticamente. Como ningún campo actúa de clave, los equipos creados sin conexión
  nunca entran en conflicto con los del servidor.
"""
from __future__ import annotations

import json
import shutil
import socket
import sqlite3
import threading
import time
import uuid
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from urllib import error as urlerror
from urllib import request as urlrequest
from urllib.parse import quote, urlparse

from .campos import CLAVES
from .db import BaseDatos, Conflicto, SinConexion, ahora, carpeta_datos, guardar_config, leer_config, usuario_actual

MENSAJE_SIN_CONEXION = ("No hay conexión con el servidor.\n\nMientras tanto puedes consultar los datos de la última "
                        "sincronización y crear equipos nuevos: se enviarán al servidor cuando vuelva la conexión.")


class ErrorServidor(Exception):
    """El servidor ha rechazado la petición (clave incorrecta, error interno…)."""


def normalizar_url(texto: str) -> str:
    texto = texto.strip().rstrip("/")
    if not texto:
        raise ValueError("Indica la dirección del servidor")
    if "://" not in texto:
        texto = "http://" + texto
    u = urlparse(texto)
    if not u.hostname:
        raise ValueError("Dirección no válida")
    puerto = u.port or 8765
    return f"{u.scheme}://{u.hostname}:{puerto}"


def _peticion(url: str, clave: str, metodo: str, ruta: str, datos=None, cuerpo: bytes | None = None,
              timeout: float = 8, cliente: str = "", usuario: str = ""):
    """Hace una petición al servidor. Devuelve el JSON (o bytes). Lanza SinConexion, Conflicto o ErrorServidor."""
    cabeceras = {"X-Clave": clave, "X-Cliente": cliente, "X-Usuario": quote(usuario or "")}
    if datos is not None:
        cuerpo = json.dumps(datos, ensure_ascii=False).encode("utf-8")
        cabeceras["Content-Type"] = "application/json"
    req = urlrequest.Request(url + ruta, data=cuerpo if cuerpo is not None else (b"" if metodo != "GET" else None),
                             method=metodo, headers=cabeceras)
    try:
        with urlrequest.urlopen(req, timeout=timeout) as r:
            contenido = r.read()
            if r.headers.get_content_type() == "application/json":
                return json.loads(contenido.decode("utf-8"))
            return contenido
    except urlerror.HTTPError as e:
        try:
            info = json.loads(e.read().decode("utf-8"))
        except Exception:  # noqa: BLE001
            info = {}
        if e.code == 409 and "conflicto" in info:
            c = info["conflicto"]
            raise Conflicto(c["equipo_id"], [tuple(x) for x in c["campos"]], c["eliminado"]) from None
        if e.code == 401:
            raise ErrorServidor("La clave del servidor no es correcta.") from None
        raise ErrorServidor(f"El servidor ha respondido con un error ({e.code}): {info.get('error', '')}") from None
    except (urlerror.URLError, socket.timeout, TimeoutError, ConnectionError, OSError) as e:
        raise SinConexion(MENSAJE_SIN_CONEXION) from e


def probar_conexion(url: str, clave: str) -> dict:
    return _peticion(normalizar_url(url), clave, "GET", "/api/estado", timeout=5)


def id_cliente() -> str:
    cfg = leer_config()
    if not cfg.get("id_cliente"):
        cfg["id_cliente"] = uuid.uuid4().hex
        guardar_config(cfg)
    return cfg["id_cliente"]


class ClienteRemoto:
    remoto = True
    firma = classmethod(BaseDatos.firma.__func__)  # type: ignore[attr-defined]
    _limpiar = staticmethod(BaseDatos._limpiar)
    _descripcion = staticmethod(BaseDatos._descripcion)

    def __init__(self, url: str, clave: str, conectar: bool = True, intervalo: float = 5.0):
        self.url = normalizar_url(url)
        self.clave = clave
        self.usuario = usuario_actual()
        self.cliente = id_cliente()
        u = urlparse(self.url)
        nombre = f"{u.hostname}_{u.port}".replace(":", "_")
        self.carpeta = carpeta_datos() / "servidores" / nombre
        self.carpeta.mkdir(parents=True, exist_ok=True)
        self.ruta = self.carpeta / "copia_local.db"
        self.con = sqlite3.connect(str(self.ruta), timeout=15)
        self.con.row_factory = sqlite3.Row
        self._crear_cache()
        self._en_linea = False
        self._estado: dict = {}
        self._lote: list | None = None
        self._nombre_lote = ""
        self._nivel = 0
        self.ultimos_ids: list[int] = []
        self._cambios_locales = 0
        self._parar = threading.Event()
        self._hilo_ok: bool | None = None
        self._hilo_estado: dict = {}
        if conectar:
            try:
                self.sincronizar()
            except SinConexion:
                pass
        self._hilo = threading.Thread(target=self._comprobar_periodicamente, args=(intervalo,), daemon=True)
        self._hilo.start()

    # -------------------------------------------------------------- copia local
    def _crear_cache(self) -> None:
        columnas = ",\n".join(f"{c} TEXT NOT NULL DEFAULT ''" for c in CLAVES)
        self.con.executescript(f"""
            CREATE TABLE IF NOT EXISTS equipos (id INTEGER PRIMARY KEY, {columnas},
                creado TEXT NOT NULL DEFAULT '', modificado TEXT NOT NULL DEFAULT '');
            CREATE TABLE IF NOT EXISTS pendientes (id INTEGER PRIMARY KEY AUTOINCREMENT, uuid TEXT NOT NULL,
                datos TEXT NOT NULL, creado TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS historial (id INTEGER PRIMARY KEY, datos TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS adjuntos_info (equipo_id INTEGER PRIMARY KEY, datos TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS meta (clave TEXT PRIMARY KEY, valor TEXT NOT NULL);
        """)
        existentes = {r["name"] for r in self.con.execute("PRAGMA table_info(equipos)")}
        for c in CLAVES:
            if c not in existentes:
                self.con.execute(f"ALTER TABLE equipos ADD COLUMN {c} TEXT NOT NULL DEFAULT ''")
        self.con.commit()

    def _meta(self, clave: str, defecto: str = "") -> str:
        f = self.con.execute("SELECT valor FROM meta WHERE clave=?", (clave,)).fetchone()
        return f[0] if f else defecto

    def _poner_meta(self, clave: str, valor) -> None:
        self.con.execute("INSERT OR REPLACE INTO meta (clave, valor) VALUES (?, ?)", (clave, str(valor)))

    @property
    def ultima_sincronizacion(self) -> str:
        return self._meta("ultima_sincronizacion")

    @property
    def descripcion(self) -> str:
        u = urlparse(self.url)
        return f"Servidor {u.hostname}:{u.port}"

    @property
    def en_linea(self) -> bool:
        return self._en_linea

    @property
    def pendientes(self) -> int:
        return self.con.execute("SELECT COUNT(*) FROM pendientes").fetchone()[0]

    def puede_modificar(self, equipo_id: int) -> bool:
        return equipo_id < 0 or self._en_linea

    # -------------------------------------------------------------- comunicación
    def _api(self, metodo: str, ruta: str, datos=None, cuerpo: bytes | None = None, timeout: float = 10):
        try:
            res = _peticion(self.url, self.clave, metodo, ruta, datos, cuerpo, timeout, self.cliente, self.usuario)
        except SinConexion:
            self._en_linea = False
            raise
        self._en_linea = True
        return res

    def _requiere_conexion(self) -> None:
        if not self._en_linea:
            raise SinConexion(MENSAJE_SIN_CONEXION)

    def _comprobar_periodicamente(self, intervalo: float) -> None:
        """Hilo en segundo plano: comprueba si el servidor responde (sin tocar la base de datos local)."""
        while not self._parar.wait(intervalo):
            try:
                self._hilo_estado = _peticion(self.url, self.clave, "GET", "/api/estado", timeout=3,
                                              cliente=self.cliente, usuario=self.usuario)
                self._hilo_ok = True
            except (SinConexion, ErrorServidor):
                self._hilo_ok = False
            except Exception:  # noqa: BLE001
                self._hilo_ok = False

    def version_datos(self):
        """La llama la interfaz periódicamente: sincroniza si hace falta y devuelve un valor que cambia
        cuando hay que volver a pintar los datos."""
        ok, estado = self._hilo_ok, self._hilo_estado
        self._hilo_ok = None
        if ok is True:
            vuelve = not self._en_linea
            self._en_linea = True
            self._estado = estado
            if vuelve or self.pendientes or str(estado.get("revision")) != self._meta("revision"):
                try:
                    self.sincronizar()
                except (SinConexion, ErrorServidor):
                    pass
        elif ok is False:
            self._en_linea = False
        return (self._meta("revision"), self._en_linea, self.pendientes, self._cambios_locales)

    def revision(self) -> int:
        try:
            return int(self._meta("revision", "0"))
        except ValueError:
            return 0

    def sincronizar(self) -> dict:
        """Envía los equipos creados sin conexión y descarga los datos del servidor."""
        self._estado = self._api("GET", "/api/estado", timeout=5)
        enviados = self._enviar_pendientes()
        self._descargar()
        return {"enviados": enviados}

    def _enviar_pendientes(self) -> int:
        filas = self.con.execute("SELECT * FROM pendientes ORDER BY id").fetchall()
        if not filas:
            return 0
        ops = [{"op": "insertar", "datos": json.loads(f["datos"]), "uuid": f["uuid"]} for f in filas]
        ops.append({"op": "registrar", "equipo_id": None, "equipo": f"{len(filas)} equipos",
                    "accion": "Sincronización", "detalle": "Equipos creados sin conexión"})
        nombre = f"Sincronizar {len(filas)} equipo{'s' if len(filas) != 1 else ''} creado{'s' if len(filas) != 1 else ''} sin conexión"
        res = self._api("POST", "/api/lote", {"nombre": nombre, "ops": ops}, timeout=60)["resultados"]
        for f, nuevo_id in zip(filas, res):
            carpeta = self._carpeta_pendiente(f["id"])
            if carpeta.is_dir():
                for archivo in sorted(carpeta.iterdir()):
                    if archivo.is_file():
                        self._api("POST", f"/api/adjuntos/{int(nuevo_id)}?nombre={quote(archivo.name)}",
                                  cuerpo=archivo.read_bytes(), timeout=300)
                shutil.rmtree(carpeta, ignore_errors=True)
            self.con.execute("DELETE FROM pendientes WHERE id=?", (f["id"],))
            self.con.commit()
        self.ultimos_ids = [int(r) for r in res if isinstance(r, int)]
        return len(filas)

    def _descargar(self) -> None:
        datos = self._api("GET", "/api/datos", timeout=60)
        cols = ["id"] + CLAVES + ["creado", "modificado"]
        with self.con:
            self.con.execute("DELETE FROM equipos")
            self.con.executemany(f"INSERT INTO equipos ({', '.join(cols)}) VALUES ({', '.join('?' for _ in cols)})",
                                 [[e.get(c, "") if c != "id" else e["id"] for c in cols] for e in datos["equipos"]])
            self.con.execute("DELETE FROM historial")
            self.con.executemany("INSERT INTO historial (id, datos) VALUES (?, ?)",
                                 [(h["id"], json.dumps(h, ensure_ascii=False)) for h in datos["historial"]])
            self.con.execute("DELETE FROM adjuntos_info")
            self.con.executemany("INSERT INTO adjuntos_info (equipo_id, datos) VALUES (?, ?)",
                                 [(int(k), json.dumps(v)) for k, v in datos.get("adjuntos", {}).items()])
            self._poner_meta("revision", datos["revision"])
            self._poner_meta("ultima_sincronizacion", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    def _refrescar(self) -> None:
        """Tras escribir en el servidor: actualizar la copia local."""
        try:
            self._estado = self._api("GET", "/api/estado", timeout=5)
            self._descargar()
        except SinConexion:
            pass

    # -------------------------------------------------------------- consultas
    @staticmethod
    def _fila_pendiente(f) -> dict:
        d = {c: "" for c in CLAVES}
        d.update(json.loads(f["datos"]))
        d.update({"id": -int(f["id"]), "creado": f["creado"], "modificado": f["creado"], "_pendiente": True})
        return d

    def listar(self) -> list[dict]:
        filas = [dict(f) for f in self.con.execute("SELECT * FROM equipos ORDER BY id")]
        filas += [self._fila_pendiente(f) for f in self.con.execute("SELECT * FROM pendientes ORDER BY id")]
        return filas

    def obtener(self, equipo_id: int) -> dict | None:
        if equipo_id < 0:
            f = self.con.execute("SELECT * FROM pendientes WHERE id=?", (-equipo_id,)).fetchone()
            return self._fila_pendiente(f) if f else None
        f = self.con.execute("SELECT * FROM equipos WHERE id=?", (equipo_id,)).fetchone()
        return dict(f) if f else None

    def contar(self) -> int:
        return len(self.listar())

    def valores_distintos(self, clave: str) -> list[str]:
        if clave not in CLAVES:
            return []
        return sorted({e[clave] for e in self.listar() if e.get(clave)}, key=str.lower)

    def firmas(self) -> dict[tuple, int]:
        return {self.firma(e): e["id"] for e in self.listar()}

    def historial(self, limite: int = 2000, equipo_id: int | None = None) -> list[dict]:
        if equipo_id is not None and equipo_id < 0:
            return []
        if self._en_linea:
            try:
                ruta = f"/api/historial?limite={int(limite)}" + (f"&equipo_id={int(equipo_id)}" if equipo_id else "")
                return self._api("GET", ruta)
            except SinConexion:
                pass
        filas = [json.loads(f[0]) for f in self.con.execute("SELECT datos FROM historial ORDER BY id DESC")]
        if equipo_id is not None:
            filas = [h for h in filas if h.get("equipo_id") == equipo_id]
        return filas[:limite]

    # -------------------------------------------------------------- deshacer
    def nombre_deshacer(self) -> str | None:
        return self._estado.get("deshacer") if self._en_linea else None

    def nombre_rehacer(self) -> str | None:
        return self._estado.get("rehacer") if self._en_linea else None

    def deshacer(self):
        self._requiere_conexion()
        r = self._api("POST", "/api/deshacer")
        self._refrescar()
        return (r["nombre"], r["avisos"]) if r.get("nombre") else None

    def rehacer(self):
        self._requiere_conexion()
        r = self._api("POST", "/api/rehacer")
        self._refrescar()
        return (r["nombre"], r["avisos"]) if r.get("nombre") else None

    def olvidar_deshacer(self) -> None:
        pass

    # -------------------------------------------------------------- escritura
    @contextmanager
    def operacion(self, nombre: str):
        """Agrupa cambios: con conexión se envían juntos al servidor (una sola operación que se puede
        deshacer). Sin conexión sólo se admiten altas, que quedan pendientes."""
        self._nivel += 1
        if self._nivel == 1:
            self._lote = [] if self._en_linea else None
            self._nombre_lote = nombre
        try:
            yield
        except BaseException:
            if self._nivel == 1:
                if self._lote is None:
                    self.con.rollback()
                self._lote = None
            self._nivel -= 1
            raise
        self._nivel -= 1
        if self._nivel == 0:
            lote, self._lote = self._lote, None
            if lote is None:            # sin conexión: los pendientes ya están en la copia local
                self.con.commit()
                self._cambios_locales += 1
                return
            if not lote:
                return
            try:
                res = self._api("POST", "/api/lote", {"nombre": nombre, "ops": lote}, timeout=120)["resultados"]
            except SinConexion:
                if all(op["op"] in ("insertar", "registrar") for op in lote):
                    for op in lote:
                        if op["op"] == "insertar":
                            self._guardar_pendiente(op["datos"])
                    self.con.commit()
                    self.ultimos_ids = []
                    self._cambios_locales += 1
                    return
                raise
            self.ultimos_ids = [r for op, r in zip(lote, res) if op["op"] == "insertar"]
            self._refrescar()

    def _guardar_pendiente(self, datos: dict) -> int:
        d = self._limpiar(datos)
        if not d["estado"]:
            d["estado"] = "Activo"
        cur = self.con.execute("INSERT INTO pendientes (uuid, datos, creado) VALUES (?, ?, ?)",
                               (uuid.uuid4().hex, json.dumps(d, ensure_ascii=False), ahora()))
        return -int(cur.lastrowid)

    def insertar(self, datos: dict, registrar: bool = True, commit: bool = True) -> int | None:
        if self._lote is not None:
            self._lote.append({"op": "insertar", "datos": self._limpiar(datos)})
            return None
        if self._nivel or not self._en_linea:
            nuevo = self._guardar_pendiente(datos)
            if not self._nivel:
                self.con.commit()
                self._cambios_locales += 1
            return nuevo
        try:
            nuevo = self._api("POST", "/api/equipos", {"datos": self._limpiar(datos)})["id"]
        except SinConexion:
            nuevo = self._guardar_pendiente(datos)
            self.con.commit()
            self._cambios_locales += 1
            return nuevo
        self._refrescar()
        return nuevo

    def _actualizar_pendiente(self, equipo_id: int, datos: dict) -> bool:
        actual = self.obtener(equipo_id)
        if actual is None:
            return False
        d = self._limpiar({**actual, **datos})
        self.con.execute("UPDATE pendientes SET datos=? WHERE id=?", (json.dumps(d, ensure_ascii=False), -equipo_id))
        if not self._nivel:
            self.con.commit()
        self._cambios_locales += 1
        return True

    def actualizar(self, equipo_id: int, datos: dict, registrar: bool = True, commit: bool = True,
                   base: dict | None = None, forzar: bool = False) -> bool:
        if equipo_id < 0:
            return self._actualizar_pendiente(equipo_id, datos)
        op = {"op": "actualizar", "id": equipo_id, "datos": self._limpiar(datos),
              "base": self._limpiar(base) if base is not None else None, "forzar": forzar}
        if self._lote is not None:
            self._lote.append(op)
            return True
        self._requiere_conexion()
        r = self._api("PUT", f"/api/equipos/{int(equipo_id)}",
                      {"datos": op["datos"], "base": op["base"], "forzar": forzar})
        self._refrescar()
        return r["cambiado"]

    def actualizar_campo(self, equipo_id: int, clave: str, valor: str, esperado: str | None = None,
                         forzar: bool = False) -> bool:
        if equipo_id < 0:
            return self._actualizar_pendiente(equipo_id, {clave: valor})
        op = {"op": "actualizar_campo", "id": equipo_id, "clave": clave, "valor": valor,
              "esperado": esperado, "forzar": forzar}
        if self._lote is not None:
            self._lote.append(op)
            return True
        self._requiere_conexion()
        self._api("POST", "/api/lote", {"nombre": "Modificación", "ops": [op]})
        self._refrescar()
        return True

    def eliminar(self, ids: list[int]) -> int:
        locales = [i for i in ids if i < 0]
        remotos = [i for i in ids if i >= 0]
        if remotos:
            if self._lote is not None:
                self._lote.append({"op": "eliminar", "ids": remotos})
            else:
                self._requiere_conexion()
        for i in locales:
            self.con.execute("DELETE FROM pendientes WHERE id=?", (-i,))
            shutil.rmtree(self._carpeta_pendiente(-i), ignore_errors=True)
        self.con.commit()
        if locales:
            self._cambios_locales += 1
        if remotos and self._lote is None:
            n = self._api("POST", "/api/eliminar", {"ids": remotos})["n"]
            self._refrescar()
            return n + len(locales)
        return len(ids)

    def vaciar(self, motivo: str = "") -> int:
        self._requiere_conexion()
        if self._lote is not None:
            self._lote.append({"op": "vaciar", "motivo": motivo})
            return self.con.execute("SELECT COUNT(*) FROM equipos").fetchone()[0]
        r = self._api("POST", "/api/lote", {"nombre": "Vaciar inventario", "ops": [{"op": "vaciar", "motivo": motivo}]})
        self._refrescar()
        return r["resultados"][0]

    def registrar(self, equipo_id, equipo: str, accion: str, detalle: str) -> None:
        op = {"op": "registrar", "equipo_id": equipo_id, "equipo": equipo, "accion": accion, "detalle": detalle}
        if self._lote is not None:
            self._lote.append(op)
            return
        if self._en_linea:
            try:
                self._api("POST", "/api/registrar", op)
            except SinConexion:
                pass

    def commit(self) -> None:
        self.con.commit()

    def vaciar_historial(self) -> None:
        self._requiere_conexion()
        self._api("DELETE", "/api/historial")
        self._refrescar()

    # -------------------------------------------------------------- adjuntos
    def _carpeta_pendiente(self, pid: int) -> Path:
        return self.carpeta / "adjuntos_pendientes" / str(int(pid))

    def carpeta_adjuntos(self, equipo_id: int | None = None) -> Path:
        base = self.carpeta / "adjuntos"
        if equipo_id is None:
            return base
        if equipo_id < 0:
            return self._carpeta_pendiente(-equipo_id)
        return base / str(int(equipo_id))

    def info_adjuntos(self, equipo_id: int) -> list[dict]:
        if equipo_id < 0:
            carpeta = self._carpeta_pendiente(-equipo_id)
            if not carpeta.is_dir():
                return []
            return [{"nombre": p.name, "tamano": p.stat().st_size, "fecha": p.stat().st_mtime}
                    for p in sorted(carpeta.iterdir(), key=lambda p: p.name.lower()) if p.is_file()]
        if self._en_linea:
            try:
                info = self._api("GET", f"/api/adjuntos/{int(equipo_id)}")
                self.con.execute("INSERT OR REPLACE INTO adjuntos_info (equipo_id, datos) VALUES (?, ?)",
                                 (equipo_id, json.dumps(info)))
                self.con.commit()
                return info
            except SinConexion:
                pass
        f = self.con.execute("SELECT datos FROM adjuntos_info WHERE equipo_id=?", (equipo_id,)).fetchone()
        return json.loads(f[0]) if f else []

    def listar_adjuntos(self, equipo_id: int) -> list[Path]:
        return [self.carpeta_adjuntos(equipo_id) / a["nombre"] for a in self.info_adjuntos(equipo_id)]

    def contar_adjuntos(self, equipo_id: int) -> int:
        return len(self.info_adjuntos(equipo_id))

    def ruta_adjunto(self, equipo_id: int, nombre: str) -> Path:
        """Ruta local del adjunto; si hace falta lo descarga del servidor."""
        destino = self.carpeta_adjuntos(equipo_id) / Path(nombre).name
        if equipo_id < 0:
            return destino
        if self._en_linea:
            try:
                contenido = self._api("GET", f"/api/adjuntos/{int(equipo_id)}/{quote(Path(nombre).name)}", timeout=300)
                destino.parent.mkdir(parents=True, exist_ok=True)
                destino.write_bytes(contenido)
                return destino
            except SinConexion:
                pass
        if destino.is_file():
            return destino  # descargado en una conexión anterior
        raise SinConexion("Este adjunto no se ha descargado nunca en este equipo y ahora no hay conexión "
                          "con el servidor.")

    def anadir_adjunto(self, equipo_id: int, origen) -> Path:
        origen = Path(origen)
        if equipo_id < 0:
            carpeta = self._carpeta_pendiente(-equipo_id)
            carpeta.mkdir(parents=True, exist_ok=True)
            destino = carpeta / origen.name
            n = 2
            while destino.exists():
                destino = carpeta / f"{origen.stem} ({n}){origen.suffix}"
                n += 1
            shutil.copy2(origen, destino)
            return destino
        self._requiere_conexion()
        r = self._api("POST", f"/api/adjuntos/{int(equipo_id)}?nombre={quote(origen.name)}",
                      cuerpo=origen.read_bytes(), timeout=300)
        return self.carpeta_adjuntos(equipo_id) / r["nombre"]

    def eliminar_adjunto(self, equipo_id: int, nombre: str) -> None:
        if equipo_id < 0:
            (self._carpeta_pendiente(-equipo_id) / Path(nombre).name).unlink(missing_ok=True)
            return
        self._requiere_conexion()
        self._api("DELETE", f"/api/adjuntos/{int(equipo_id)}/{quote(Path(nombre).name)}")

    # -------------------------------------------------------------- copias
    def copia_seguridad(self, destino) -> Path:
        self._requiere_conexion()
        destino = Path(destino)
        destino.parent.mkdir(parents=True, exist_ok=True)
        destino.write_bytes(self._api("GET", "/api/copia", timeout=300))
        return destino

    def restaurar(self, origen) -> None:
        raise ErrorServidor("Con un servidor central, las copias se restauran en el propio servidor "
                            "(ventana del servidor → Restaurar copia).")

    def copia_automatica(self, maximo: int = 10) -> None:
        """Las copias automáticas las hace el servidor."""

    def cerrar(self) -> None:
        self._parar.set()
        try:
            self.con.commit()
            self.con.close()
        except sqlite3.Error:
            pass


def esperar_conexion(cliente: ClienteRemoto, segundos: float = 10) -> bool:
    """Utilidad para pruebas: espera a que el hilo detecte el servidor."""
    fin = time.time() + segundos
    while time.time() < fin:
        cliente.version_datos()
        if cliente.en_linea:
            return True
        time.sleep(0.2)
    return False
