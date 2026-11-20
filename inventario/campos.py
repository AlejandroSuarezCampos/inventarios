"""Definición de los campos del inventario y utilidades de normalización."""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta


@dataclass(frozen=True)
class Campo:
    clave: str            # nombre de la columna en la base de datos
    etiqueta: str         # texto mostrado en la interfaz
    cabecera: str         # cabecera usada al exportar a Excel
    grupo: str            # "propietario" | "equipo" | "otros"
    tipo: str = "texto"   # "texto" | "fecha" | "largo"
    ancho: int = 120      # ancho de columna en la tabla (px)
    alias: tuple = field(default_factory=tuple)  # otras cabeceras aceptadas al importar


CAMPOS: list[Campo] = [
    Campo("persona", "Persona", "PERSONA", "propietario", ancho=150,
          alias=("PROPIETARIO", "NOMBRE", "EMPLEADO", "RESPONSABLE")),
    Campo("usuario_pc", "Usuario PC", "USUARIO PC", "propietario", ancho=110,
          alias=("USUARIO", "USER", "LOGIN")),
    Campo("teamviewer", "TeamViewer", "TEAMVIEWER", "propietario", ancho=110,
          alias=("ID TEAMVIEWER", "TEAM VIEWER", "TV")),
    Campo("delegacion", "Delegación/CT", "DELEGACION/CT", "propietario", ancho=120,
          alias=("DELEGACION", "CT", "CENTRO", "CENTRO DE TRABAJO", "SEDE", "OFICINA")),
    Campo("nombre_equipo", "Nombre equipo", "NOMBRE EQUIPO", "equipo", ancho=130,
          alias=("NOMBRE QUIPO", "HOSTNAME", "EQUIPO", "NOMBRE PC", "NOMBRE DEL EQUIPO")),
    Campo("marca", "Marca", "MARCA", "equipo", ancho=90, alias=("FABRICANTE",)),
    Campo("modelo", "Modelo", "MODELO", "equipo", ancho=130),
    Campo("os", "Sistema operativo", "OS", "equipo", ancho=120,
          alias=("SO", "SISTEMA OPERATIVO", "S.O.")),
    Campo("cpu", "CPU", "CPU", "equipo", ancho=140, alias=("PROCESADOR",)),
    Campo("ram", "RAM", "RAM", "equipo", ancho=70, alias=("MEMORIA", "MEMORIA RAM")),
    Campo("tipo_ram", "Tipo de RAM", "TIPO DE RAM", "equipo", ancho=90, alias=("TIPO RAM",)),
    Campo("almacenamiento", "Almacenamiento", "ALMACENAMIENTO", "equipo", ancho=120,
          alias=("DISCO", "DISCO DURO", "HDD", "SSD")),
    Campo("num_serie", "Nº de serie", "Nº DE SERIE", "equipo", ancho=120,
          alias=("N DE SERIE", "NUMERO DE SERIE", "NO DE SERIE", "SERIE", "S/N", "SN", "N SERIE",
                 "NUM SERIE", "SERIAL")),
    Campo("pantalla", "Pantalla", "PANTALLA", "equipo", ancho=100, alias=("MONITOR",)),
    Campo("office", "Office", "OFFICE", "equipo", ancho=100, alias=("OFIMATICA", "MICROSOFT OFFICE")),
    Campo("antivirus", "Antivirus", "ANTIVIRUS", "equipo", ancho=100),
    Campo("fecha_compra", "Fecha de compra", "FECHA DE COMPRA", "equipo", tipo="fecha", ancho=100,
          alias=("FECHA COMPRA", "COMPRA")),
    Campo("fin_garantia", "Expiración garantía", "EXPIRACION GARANTIA", "equipo", tipo="fecha",
          ancho=110, alias=("EXPIACION GARANTIA", "GARANTIA", "FIN GARANTIA", "FIN DE GARANTIA",
                            "EXPIRACION DE GARANTIA", "VENCIMIENTO GARANTIA")),
    Campo("estado", "Estado", "ESTADO", "otros", ancho=90, alias=("SITUACION",)),
    Campo("observaciones", "Observaciones", "OBSERVACIONES", "otros", tipo="largo", ancho=200,
          alias=("NOTAS", "COMENTARIOS", "OBS")),
]

CAMPOS_POR_CLAVE = {c.clave: c for c in CAMPOS}
CLAVES = [c.clave for c in CAMPOS]

ESTADOS = ["Activo", "En almacén", "En reparación", "Prestado", "Baja"]

GRUPOS = {
    "propietario": "PROPIETARIO DEL ORDENADOR",
    "equipo": "DATOS DEL ORDENADOR",
    "otros": "OTROS",
}

DIAS_AVISO_GARANTIA = 90


def normalizar(texto) -> str:
    """Mayúsculas, sin tildes y sin signos: 'Nº de Serie' -> 'NDESERIE'."""
    if texto is None:
        return ""
    t = str(texto).replace("º", "").replace("ª", "")
    t = unicodedata.normalize("NFKD", t)
    t = "".join(ch for ch in t if not unicodedata.combining(ch))
    return re.sub(r"[^A-Z0-9]", "", t.upper())


# Mapa cabecera normalizada -> clave del campo
_MAPA_CABECERAS: dict[str, str] = {}
for _c in CAMPOS:
    for _txt in (_c.cabecera, _c.etiqueta, _c.clave, *_c.alias):
        _MAPA_CABECERAS.setdefault(normalizar(_txt), _c.clave)


def clave_para_cabecera(cabecera) -> str | None:
    n = normalizar(cabecera)
    if not n:
        return None
    return _MAPA_CABECERAS.get(n)


# ------------------------------------------------------------------ fechas

_FORMATOS_FECHA = ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%Y-%m-%d", "%Y/%m/%d",
                   "%d/%m/%y", "%d-%m-%y", "%m/%Y", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M")


def parsear_fecha(valor) -> str:
    """Devuelve la fecha en formato ISO (AAAA-MM-DD) o '' si no se reconoce.

    Si el texto no es una fecha reconocible se devuelve tal cual para no perder datos.
    """
    if valor is None or valor == "":
        return ""
    if isinstance(valor, datetime):
        return valor.date().isoformat()
    if isinstance(valor, date):
        return valor.isoformat()
    if isinstance(valor, (int, float)):
        # Número de serie de fecha de Excel
        if 1 < valor < 80000:
            return (datetime(1899, 12, 30) + timedelta(days=float(valor))).date().isoformat()
        return str(valor)
    texto = str(valor).strip()
    if not texto:
        return ""
    for fmt in _FORMATOS_FECHA:
        try:
            return datetime.strptime(texto, fmt).date().isoformat()
        except ValueError:
            continue
    if re.fullmatch(r"\d{4}", texto):  # sólo año
        return f"{texto}-01-01"
    return texto


def fecha_a_date(iso: str) -> date | None:
    if not iso:
        return None
    try:
        return date.fromisoformat(iso[:10])
    except ValueError:
        return None


def mostrar_fecha(iso: str) -> str:
    d = fecha_a_date(iso)
    return d.strftime("%d/%m/%Y") if d else (iso or "")


def estado_garantia(iso: str, hoy: date | None = None) -> str:
    """'caducada' | 'proxima' | 'vigente' | 'sin_dato'"""
    d = fecha_a_date(iso)
    if d is None:
        return "sin_dato"
    hoy = hoy or date.today()
    if d < hoy:
        return "caducada"
    if (d - hoy).days <= DIAS_AVISO_GARANTIA:
        return "proxima"
    return "vigente"


def valor_texto(valor) -> str:
    """Convierte un valor de celda a texto limpio (sin '.0' en enteros)."""
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    if isinstance(valor, (datetime, date)):
        return valor.strftime("%d/%m/%Y")
    return str(valor).strip()
