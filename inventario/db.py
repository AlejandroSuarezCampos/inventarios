"""Capa de acceso a datos (SQLite).

Ningún campo del inventario actúa como clave: cada equipo se identifica sólo por su ID interno.

Además de las operaciones básicas incluye:
* Deshacer / rehacer de todas las operaciones de la sesión.
* Edición segura cuando varias personas usan la misma base de datos (p. ej. en una carpeta de red):
  al guardar sólo se aplican los campos que ha cambiado el usuario y, si otra persona ha cambiado
  esos mismos campos entretanto, se lanza :class:`Conflicto` en lugar de sobrescribir sin avisar.
* Nombre del usuario de Windows en el historial.
* Archivos adjuntos por equipo (en una carpeta junto a la base de datos).
"""
from __future__ import annotations

import getpass
import json
import os
import shutil
import socket
import sqlite3
import sys
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

from .campos import CAMPOS, CAMPOS_POR_CLAVE, CLAVES

APP_NOMBRE = "InventarioPC"


def carpeta_datos() -> Path:
    """Carpeta de datos del usuario (en Windows: %APPDATA%\\InventarioPC)."""
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA") or Path.home() / "AppData" / "Roaming")
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_DATA_HOME") or Path.home() / ".local" / "share")
    ruta = base / APP_NOMBRE
    ruta.mkdir(parents=True, exist_ok=True)
    return ruta


def ruta_config() -> Path:
    return carpeta_datos() / "config.json"


def leer_config() -> dict:
    try:
        return json.loads(ruta_config().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def guardar_config(cfg: dict) -> None:
    try:
        ruta_config().write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")
    except OSError:
        pass


def ruta_bd_por_defecto() -> Path:
    cfg = leer_config()
    if cfg.get("ruta_bd"):
        return Path(cfg["ruta_bd"])
    return carpeta_datos() / "inventario.db"


def usuario_actual() -> str:
    """Nombre con el que se firman los cambios: el de Preferencias o el usuario de Windows."""
    nombre = (leer_config().get("usuario") or "").strip()
    if nombre:
        return nombre
    try:
        usuario = getpass.getuser()
    except Exception:  # noqa: BLE001
        usuario = os.environ.get("USERNAME") or os.environ.get("USER") or "desconocido"
    try:
        equipo = socket.gethostname()
    except OSError:
        equipo = ""
    return f"{usuario}@{equipo}" if equipo else usuario


def ahora() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


class Conflicto(Exception):
    """Otra persona ha modificado (o eliminado) el equipo mientras se editaba."""

    def __init__(self, equipo_id: int, campos: list[tuple[str, str, str, str]] | None = None,
                 eliminado: bool = False):
        self.equipo_id = equipo_id
        self.campos = campos or []      # (clave, valor_original, valor_de_otro, valor_mio)
        self.eliminado = eliminado
        super().__init__(self.describir())

    def describir(self) -> str:
        if self.eliminado:
            return "Otra persona ha eliminado este equipo mientras lo editabas."
        lineas = ["Otra persona ha cambiado estos campos mientras los editabas:", ""]
        for clave, _orig, suyo, mio in self.campos:
            lineas.append(f"• {CAMPOS_POR_CLAVE[clave].etiqueta}: suyo «{suyo or '(vacío)'}» · "
                          f"tuyo «{mio or '(vacío)'}»")
        return "\n".join(lineas)


class BaseDatos:
    def __init__(self, ruta: str | os.PathLike | None = None):
        self.ruta = Path(ruta) if ruta else ruta_bd_por_defecto()
        self.ruta.parent.mkdir(parents=True, exist_ok=True)
        self.usuario = usuario_actual()
        self._conectar()
        self._deshacer: list[dict] = []
        self._rehacer: list[dict] = []
        self._op: dict | None = None
        self._nivel_op = 0

    def _conectar(self) -> None:
        # timeout alto: en una carpeta de red otra persona puede tener la BD bloqueada un instante
        self.con = sqlite3.connect(str(self.ruta), timeout=15)
        self.con.row_factory = sqlite3.Row
        self.con.execute("PRAGMA foreign_keys = ON")
        self.con.execute("PRAGMA busy_timeout = 15000")
        self._crear_esquema()

    # ------------------------------------------------------------ esquema
    def _crear_esquema(self) -> None:
        columnas = ",\n".join(f"{c} TEXT NOT NULL DEFAULT ''" for c in CLAVES)
        self.con.executescript(f"""
            CREATE TABLE IF NOT EXISTS equipos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                {columnas},
                creado TEXT NOT NULL DEFAULT '',
                modificado TEXT NOT NULL DEFAULT ''
            );
            CREATE TABLE IF NOT EXISTS historial (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TEXT NOT NULL,
                equipo_id INTEGER,
                equipo TEXT NOT NULL DEFAULT '',
                accion TEXT NOT NULL,
                detalle TEXT NOT NULL DEFAULT '',
                usuario TEXT NOT NULL DEFAULT ''
            );
            CREATE INDEX IF NOT EXISTS idx_hist_fecha ON historial(fecha);
        """)
        # Migraciones desde versiones anteriores
        existentes = {r["name"] for r in self.con.execute("PRAGMA table_info(equipos)")}
        for c in CLAVES:
            if c not in existentes:
                self.con.execute(f"ALTER TABLE equipos ADD COLUMN {c} TEXT NOT NULL DEFAULT ''")
        hist = {r["name"] for r in self.con.execute("PRAGMA table_info(historial)")}
        if "usuario" not in hist:
            self.con.execute("ALTER TABLE historial ADD COLUMN usuario TEXT NOT NULL DEFAULT ''")
        self.con.execute("CREATE INDEX IF NOT EXISTS idx_hist_equipo ON historial(equipo_id)")
        self.con.commit()

    def cerrar(self) -> None:
        self.con.close()

    def version_datos(self) -> int:
        """Cambia cada vez que OTRA conexión (otra persona) guarda cambios en la BD."""
        return self.con.execute("PRAGMA data_version").fetchone()[0]

    # ------------------------------------------------------------ consultas
    def listar(self) -> list[dict]:
        filas = self.con.execute("SELECT * FROM equipos ORDER BY id").fetchall()
        return [dict(f) for f in filas]

    def obtener(self, equipo_id: int) -> dict | None:
        f = self.con.execute("SELECT * FROM equipos WHERE id=?", (equipo_id,)).fetchone()
        return dict(f) if f else None

    def contar(self) -> int:
        return self.con.execute("SELECT COUNT(*) FROM equipos").fetchone()[0]

    def valores_distintos(self, clave: str) -> list[str]:
        if clave not in CLAVES:
            return []
        filas = self.con.execute(
            f"SELECT DISTINCT {clave} FROM equipos WHERE {clave} <> '' ORDER BY {clave} COLLATE NOCASE"
        ).fetchall()
        return [f[0] for f in filas]

    @classmethod
    def firma(cls, datos: dict) -> tuple:
        """Huella de un equipo con TODOS sus campos (no hay ningún campo que actúe de clave).
        Dos registros sólo se consideran iguales si coinciden en todos los campos."""
        d = cls._limpiar(datos)
        if not d["estado"]:
            d["estado"] = "Activo"
        return tuple(" ".join(d[c].split()).casefold() for c in CLAVES)

    def firmas(self) -> dict[tuple, int]:
        """{firma: id} de todos los equipos guardados."""
        return {self.firma(dict(f)): f["id"] for f in self.con.execute("SELECT * FROM equipos")}

    # ------------------------------------------------------------ operaciones (deshacer)
    @contextmanager
    def operacion(self, nombre: str):
        """Agrupa varios cambios en una sola operación que se puede deshacer de una vez.
        Hace commit al terminar (o rollback si hay un error)."""
        self._nivel_op += 1
        if self._nivel_op == 1:
            self._op = {"nombre": nombre, "pasos": []}
        try:
            yield
        except BaseException:
            if self._nivel_op == 1:
                self.con.rollback()
                self._op = None
            self._nivel_op -= 1
            raise
        self._nivel_op -= 1
        if self._nivel_op == 0:
            self.con.commit()
            op, self._op = self._op, None
            if op and op["pasos"]:
                self._deshacer.append(op)
                del self._deshacer[:-100]
                self._rehacer.clear()

    def _paso(self, *paso) -> None:
        if self._op is not None:
            self._op["pasos"].append(paso)

    def nombre_deshacer(self) -> str | None:
        return self._deshacer[-1]["nombre"] if self._deshacer else None

    def nombre_rehacer(self) -> str | None:
        return self._rehacer[-1]["nombre"] if self._rehacer else None

    def olvidar_deshacer(self) -> None:
        self._deshacer.clear()
        self._rehacer.clear()

    def _aplicar_pasos(self, pasos, inverso: bool) -> list[str]:
        """Aplica (o revierte) una lista de pasos. Devuelve avisos de campos que no se han podido
        restaurar porque otra persona los cambió después."""
        avisos: list[str] = []
        for paso in (reversed(pasos) if inverso else pasos):
            tipo = paso[0]
            if tipo == "insert":
                _, eid, fila = paso
                if inverso:
                    self._borrar_fila(eid)
                else:
                    self._insertar_fila(fila)
            elif tipo == "delete":
                _, eid, fila = paso
                if inverso:
                    self._insertar_fila(fila)
                else:
                    self._borrar_fila(eid)
            elif tipo == "update":
                _, eid, antes, despues = paso
                origen, destino = (despues, antes) if inverso else (antes, despues)
                actual = self.obtener(eid)
                if actual is None:
                    avisos.append(f"El equipo #{eid} ya no existe.")
                    continue
                nuevo = dict(actual)
                for c in CLAVES:
                    if origen[c] == destino[c]:
                        continue
                    if actual[c] == origen[c]:
                        nuevo[c] = destino[c]
                    elif actual[c] != destino[c]:
                        avisos.append(f"#{eid} {CAMPOS_POR_CLAVE[c].etiqueta}: lo ha cambiado otra persona")
                self._escribir(eid, nuevo)
        return avisos

    def deshacer(self) -> tuple[str, list[str]] | None:
        if not self._deshacer:
            return None
        op = self._deshacer.pop()
        avisos = self._aplicar_pasos(op["pasos"], inverso=True)
        self.registrar(None, op["nombre"], "Deshacer", "; ".join(avisos))
        self.con.commit()
        self._rehacer.append(op)
        return op["nombre"], avisos

    def rehacer(self) -> tuple[str, list[str]] | None:
        if not self._rehacer:
            return None
        op = self._rehacer.pop()
        avisos = self._aplicar_pasos(op["pasos"], inverso=False)
        self.registrar(None, op["nombre"], "Rehacer", "; ".join(avisos))
        self.con.commit()
        self._deshacer.append(op)
        return op["nombre"], avisos

    # ------------------------------------------------------------ primitivas de escritura
    def _insertar_fila(self, fila: dict) -> None:
        cols = ["id"] + CLAVES + ["creado", "modificado"]
        existe = self.obtener(fila["id"]) is not None
        if existe:  # muy improbable: el ID se ha reutilizado; se inserta con un ID nuevo
            cols = cols[1:]
        self.con.execute(f"INSERT INTO equipos ({', '.join(cols)}) VALUES ({', '.join('?' for _ in cols)})",
                         [fila.get(c, "") for c in cols])

    def _borrar_fila(self, eid: int) -> None:
        self.con.execute("DELETE FROM equipos WHERE id=?", (eid,))

    def _escribir(self, eid: int, d: dict) -> None:
        sets = ", ".join(f"{c}=?" for c in CLAVES)
        self.con.execute(f"UPDATE equipos SET {sets}, modificado=? WHERE id=?",
                         [d[c] for c in CLAVES] + [ahora(), eid])

    # ------------------------------------------------------------ escritura
    @staticmethod
    def _limpiar(datos: dict) -> dict:
        return {c: ("" if datos.get(c) is None else str(datos.get(c)).strip()) for c in CLAVES}

    @staticmethod
    def _descripcion(d: dict) -> str:
        partes = [d.get("nombre_equipo"), d.get("persona"), d.get("num_serie")]
        return " / ".join(p for p in partes if p) or f"#{d.get('id', '?')}"

    def insertar(self, datos: dict, registrar: bool = True, commit: bool = True) -> int:
        with self.operacion("Alta de equipo"):
            d = self._limpiar(datos)
            if not d["estado"]:
                d["estado"] = "Activo"
            t = ahora()
            cols = ", ".join(CLAVES + ["creado", "modificado"])
            marcas = ", ".join("?" for _ in range(len(CLAVES) + 2))
            cur = self.con.execute(f"INSERT INTO equipos ({cols}) VALUES ({marcas})",
                                   [d[c] for c in CLAVES] + [t, t])
            nuevo = cur.lastrowid
            self._paso("insert", nuevo, self.obtener(nuevo))
            if registrar:
                self.registrar(nuevo, self._descripcion(d), "Alta", "")
        return nuevo

    def actualizar(self, equipo_id: int, datos: dict, registrar: bool = True, commit: bool = True,
                   base: dict | None = None, forzar: bool = False) -> bool:
        """Guarda los datos de un equipo.

        Si se pasa `base` (los valores que había cuando se empezó a editar) sólo se aplican los campos
        que el usuario ha cambiado respecto a `base`; los cambios que otra persona haya hecho en otros
        campos se conservan. Si otra persona ha cambiado los MISMOS campos se lanza `Conflicto`
        (salvo con `forzar=True`, que impone los valores del usuario).
        """
        with self.operacion("Modificación"):
            actual = self.obtener(equipo_id)
            if actual is None:
                if base is not None:
                    raise Conflicto(equipo_id, eliminado=True)
                return False
            d = self._limpiar(datos)
            if base is not None:
                b = self._limpiar(base)
                mios = {c: d[c] for c in CLAVES if d[c] != b[c]}
                conflictos = [(c, b[c], actual[c], d[c]) for c in mios
                              if (actual[c] or "") != b[c] and (actual[c] or "") != d[c]]
                if conflictos and not forzar:
                    raise Conflicto(equipo_id, conflictos)
                d = {c: (mios[c] if c in mios else (actual[c] or "")) for c in CLAVES}
            cambios = [(c, actual[c], d[c]) for c in CLAVES if (actual[c] or "") != d[c]]
            if not cambios:
                return False
            self._escribir(equipo_id, d)
            self._paso("update", equipo_id, {c: actual[c] or "" for c in CLAVES}, dict(d))
            if registrar:
                detalle = "; ".join(f"{CAMPOS_POR_CLAVE[c].etiqueta}: '{a}' → '{n}'" for c, a, n in cambios)
                self.registrar(equipo_id, self._descripcion(d), "Modificación", detalle)
        return True

    def actualizar_campo(self, equipo_id: int, clave: str, valor: str, esperado: str | None = None,
                         forzar: bool = False) -> bool:
        """Cambia un único campo de un equipo. `esperado` es el valor que el usuario veía antes de
        editar: si otra persona lo ha cambiado entretanto se lanza `Conflicto`."""
        actual = self.obtener(equipo_id)
        if actual is None:
            if esperado is not None:
                raise Conflicto(equipo_id, eliminado=True)
            return False
        if clave not in CLAVES:
            return False
        base = dict(actual)
        if esperado is not None:
            base[clave] = esperado
        nuevo = dict(base)
        nuevo[clave] = valor
        return self.actualizar(equipo_id, nuevo, base=base, forzar=forzar)

    def vaciar(self, motivo: str = "") -> int:
        """Elimina todos los equipos (se usa al reemplazar el inventario por un Excel)."""
        with self.operacion("Vaciar inventario"):
            filas = self.listar()
            for f in filas:
                self._paso("delete", f["id"], f)
            self.con.execute("DELETE FROM equipos")
            self.registrar(None, "(todos)", "Inventario vaciado", motivo or f"{len(filas)} equipos eliminados")
        return len(filas)

    def eliminar(self, ids: list[int]) -> int:
        n = 0
        with self.operacion("Eliminar" if len(ids) == 1 else f"Eliminar {len(ids)} equipos"):
            for i in ids:
                d = self.obtener(i)
                if d is None:
                    continue
                self._borrar_fila(i)
                self._paso("delete", i, d)
                self.registrar(i, self._descripcion(d), "Baja (eliminado)",
                               json.dumps({c: d[c] for c in CLAVES if d[c]}, ensure_ascii=False))
                n += 1
        return n

    def registrar(self, equipo_id, equipo: str, accion: str, detalle: str) -> None:
        self.con.execute(
            "INSERT INTO historial (fecha, equipo_id, equipo, accion, detalle, usuario) VALUES (?,?,?,?,?,?)",
            (ahora(), equipo_id, equipo, accion, detalle, self.usuario))

    def commit(self) -> None:
        self.con.commit()

    def historial(self, limite: int = 2000, equipo_id: int | None = None) -> list[dict]:
        if equipo_id is None:
            filas = self.con.execute("SELECT * FROM historial ORDER BY id DESC LIMIT ?", (limite,)).fetchall()
        else:
            filas = self.con.execute("SELECT * FROM historial WHERE equipo_id=? ORDER BY id DESC LIMIT ?",
                                     (equipo_id, limite)).fetchall()
        return [dict(f) for f in filas]

    def vaciar_historial(self) -> None:
        self.con.execute("DELETE FROM historial")
        self.con.commit()

    # ------------------------------------------------------------ adjuntos
    def carpeta_adjuntos(self, equipo_id: int | None = None) -> Path:
        base = self.ruta.parent / f"{self.ruta.stem}_adjuntos"
        return base if equipo_id is None else base / str(int(equipo_id))

    def listar_adjuntos(self, equipo_id: int) -> list[Path]:
        carpeta = self.carpeta_adjuntos(equipo_id)
        if not carpeta.is_dir():
            return []
        return sorted((p for p in carpeta.iterdir() if p.is_file()), key=lambda p: p.name.lower())

    def contar_adjuntos(self, equipo_id: int) -> int:
        return len(self.listar_adjuntos(equipo_id))

    def anadir_adjunto(self, equipo_id: int, origen: str | os.PathLike) -> Path:
        origen = Path(origen)
        carpeta = self.carpeta_adjuntos(equipo_id)
        carpeta.mkdir(parents=True, exist_ok=True)
        destino = carpeta / origen.name
        n = 2
        while destino.exists():
            destino = carpeta / f"{origen.stem} ({n}){origen.suffix}"
            n += 1
        shutil.copy2(origen, destino)
        eq = self.obtener(equipo_id) or {"id": equipo_id}
        self.registrar(equipo_id, self._descripcion(eq), "Adjunto añadido", destino.name)
        self.con.commit()
        return destino

    def eliminar_adjunto(self, equipo_id: int, nombre: str) -> None:
        """Mueve el adjunto a la papelera de adjuntos (no se borra definitivamente)."""
        archivo = self.carpeta_adjuntos(equipo_id) / nombre
        if not archivo.is_file():
            return
        papelera = self.carpeta_adjuntos() / "_papelera" / str(int(equipo_id))
        papelera.mkdir(parents=True, exist_ok=True)
        destino = papelera / f"{datetime.now():%Y%m%d_%H%M%S}_{archivo.name}"
        shutil.move(str(archivo), str(destino))
        eq = self.obtener(equipo_id) or {"id": equipo_id}
        self.registrar(equipo_id, self._descripcion(eq), "Adjunto eliminado", nombre)
        self.con.commit()

    # ------------------------------------------------------------ copias
    def copia_seguridad(self, destino: str | os.PathLike) -> Path:
        destino = Path(destino)
        destino.parent.mkdir(parents=True, exist_ok=True)
        self.con.commit()
        dst = sqlite3.connect(str(destino))
        with dst:
            self.con.backup(dst)
        dst.close()
        return destino

    def restaurar(self, origen: str | os.PathLike) -> None:
        origen = Path(origen)
        prueba = sqlite3.connect(str(origen))  # validar que es una BD del programa
        try:
            prueba.execute("SELECT id FROM equipos LIMIT 1")
        finally:
            prueba.close()
        self.copia_seguridad(self.ruta.with_suffix(".antes_de_restaurar.db"))
        self.con.close()
        shutil.copyfile(origen, self.ruta)
        self._conectar()
        self.olvidar_deshacer()

    def copia_automatica(self, maximo: int = 10) -> None:
        """Guarda una copia diaria en la carpeta 'copias' junto a la BD."""
        carpeta = self.ruta.parent / "copias"
        try:
            carpeta.mkdir(exist_ok=True)
        except OSError:
            return
        destino = carpeta / f"{self.ruta.stem}_{datetime.now():%Y%m%d}.db"
        if not destino.exists():
            try:
                self.copia_seguridad(destino)
            except (sqlite3.Error, OSError):
                return
        copias = sorted(carpeta.glob(f"{self.ruta.stem}_*.db"))
        for viejo in copias[:-maximo]:
            try:
                viejo.unlink()
            except OSError:
                pass


__all__ = ["BaseDatos", "CAMPOS", "Conflicto", "carpeta_datos", "leer_config", "guardar_config",
           "usuario_actual"]
