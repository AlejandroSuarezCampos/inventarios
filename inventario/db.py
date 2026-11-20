"""Capa de acceso a datos (SQLite)."""
from __future__ import annotations

import json
import os
import shutil
import sqlite3
import sys
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


def ahora() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


class BaseDatos:
    def __init__(self, ruta: str | os.PathLike | None = None):
        self.ruta = Path(ruta) if ruta else ruta_bd_por_defecto()
        self.ruta.parent.mkdir(parents=True, exist_ok=True)
        self.con = sqlite3.connect(str(self.ruta))
        self.con.row_factory = sqlite3.Row
        self.con.execute("PRAGMA foreign_keys = ON")
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
                detalle TEXT NOT NULL DEFAULT ''
            );
            CREATE INDEX IF NOT EXISTS idx_hist_fecha ON historial(fecha);
        """)
        # Migración: añadir columnas nuevas si la BD es de una versión anterior
        existentes = {r["name"] for r in self.con.execute("PRAGMA table_info(equipos)")}
        for c in CLAVES:
            if c not in existentes:
                self.con.execute(f"ALTER TABLE equipos ADD COLUMN {c} TEXT NOT NULL DEFAULT ''")
        self.con.commit()

    def cerrar(self) -> None:
        self.con.close()

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

    # ------------------------------------------------------------ escritura
    @staticmethod
    def _limpiar(datos: dict) -> dict:
        return {c: ("" if datos.get(c) is None else str(datos.get(c)).strip()) for c in CLAVES}

    @staticmethod
    def _descripcion(d: dict) -> str:
        partes = [d.get("nombre_equipo"), d.get("persona"), d.get("num_serie")]
        return " / ".join(p for p in partes if p) or f"#{d.get('id', '?')}"

    def insertar(self, datos: dict, registrar: bool = True, commit: bool = True) -> int:
        d = self._limpiar(datos)
        if not d["estado"]:
            d["estado"] = "Activo"
        t = ahora()
        cols = ", ".join(CLAVES + ["creado", "modificado"])
        marcas = ", ".join("?" for _ in range(len(CLAVES) + 2))
        cur = self.con.execute(f"INSERT INTO equipos ({cols}) VALUES ({marcas})",
                               [d[c] for c in CLAVES] + [t, t])
        nuevo = cur.lastrowid
        if registrar:
            self.registrar(nuevo, self._descripcion(d), "Alta", "")
        if commit:
            self.con.commit()
        return nuevo

    def actualizar(self, equipo_id: int, datos: dict, registrar: bool = True,
                   commit: bool = True, solo_no_vacios: bool = False) -> bool:
        anterior = self.obtener(equipo_id)
        if anterior is None:
            return False
        d = self._limpiar(datos)
        if solo_no_vacios:
            d = {c: (v if v else anterior[c]) for c, v in d.items()}
        cambios = [(c, anterior[c], d[c]) for c in CLAVES if (anterior[c] or "") != d[c]]
        if not cambios:
            return False
        sets = ", ".join(f"{c}=?" for c in CLAVES)
        self.con.execute(f"UPDATE equipos SET {sets}, modificado=? WHERE id=?",
                         [d[c] for c in CLAVES] + [ahora(), equipo_id])
        if registrar:
            detalle = "; ".join(f"{CAMPOS_POR_CLAVE[c].etiqueta}: '{a}' → '{n}'" for c, a, n in cambios)
            self.registrar(equipo_id, self._descripcion(d), "Modificación", detalle)
        if commit:
            self.con.commit()
        return True

    def actualizar_campo(self, equipo_id: int, clave: str, valor: str) -> bool:
        """Cambia un único campo de un equipo (registra el cambio en el historial)."""
        actual = self.obtener(equipo_id)
        if actual is None or clave not in CLAVES:
            return False
        actual[clave] = valor
        return self.actualizar(equipo_id, actual)

    def vaciar(self, motivo: str = "") -> int:
        """Elimina todos los equipos (se usa al reemplazar el inventario por un Excel)."""
        n = self.contar()
        self.con.execute("DELETE FROM equipos")
        self.registrar(None, "(todos)", "Inventario vaciado", motivo or f"{n} equipos eliminados")
        return n

    def eliminar(self, ids: list[int]) -> int:
        n = 0
        for i in ids:
            d = self.obtener(i)
            if d is None:
                continue
            self.con.execute("DELETE FROM equipos WHERE id=?", (i,))
            self.registrar(i, self._descripcion(d), "Baja (eliminado)",
                           json.dumps({c: d[c] for c in CLAVES if d[c]}, ensure_ascii=False))
            n += 1
        self.con.commit()
        return n

    def registrar(self, equipo_id, equipo: str, accion: str, detalle: str) -> None:
        self.con.execute(
            "INSERT INTO historial (fecha, equipo_id, equipo, accion, detalle) VALUES (?,?,?,?,?)",
            (ahora(), equipo_id, equipo, accion, detalle))

    def commit(self) -> None:
        self.con.commit()

    def historial(self, limite: int = 2000) -> list[dict]:
        filas = self.con.execute(
            "SELECT * FROM historial ORDER BY id DESC LIMIT ?", (limite,)).fetchall()
        return [dict(f) for f in filas]

    def vaciar_historial(self) -> None:
        self.con.execute("DELETE FROM historial")
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
        # Validar que es una BD del programa
        prueba = sqlite3.connect(str(origen))
        try:
            prueba.execute("SELECT id FROM equipos LIMIT 1")
        finally:
            prueba.close()
        self.copia_seguridad(self.ruta.with_suffix(".antes_de_restaurar.db"))
        self.con.close()
        shutil.copyfile(origen, self.ruta)
        self.con = sqlite3.connect(str(self.ruta))
        self.con.row_factory = sqlite3.Row
        self._crear_esquema()

    def copia_automatica(self, maximo: int = 10) -> None:
        """Guarda una copia diaria en la carpeta 'copias' junto a la BD."""
        carpeta = self.ruta.parent / "copias"
        carpeta.mkdir(exist_ok=True)
        destino = carpeta / f"inventario_{datetime.now():%Y%m%d}.db"
        if not destino.exists():
            try:
                self.copia_seguridad(destino)
            except sqlite3.Error:
                return
        copias = sorted(carpeta.glob("inventario_*.db"))
        for viejo in copias[:-maximo]:
            try:
                viejo.unlink()
            except OSError:
                pass


__all__ = ["BaseDatos", "CAMPOS", "carpeta_datos", "leer_config", "guardar_config"]
