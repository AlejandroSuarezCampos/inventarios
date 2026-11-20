"""Detección de las redes locales del equipo.

Se usa para que el servidor sólo atienda a equipos de **su misma red local** y para que el cliente
sepa, sin llegar a conectarse, si está en la red del servidor. Fuera de ella el cliente trabaja sin
conexión (consultar la última copia y crear equipos nuevos que se envían al volver a la red).
"""
from __future__ import annotations

import ipaddress
import json
import os
import re
import socket
import subprocess
import time

Red = ipaddress.IPv4Network

_IP = r"(\d{1,3}(?:\.\d{1,3}){3})"


def _ejecutar(orden: list[str], timeout: float = 15) -> str:
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if os.name == "nt" else 0
    r = subprocess.run(orden, capture_output=True, timeout=timeout, creationflags=flags)
    for cod in ("utf-8", "cp850", "cp1252", "latin-1"):
        try:
            return r.stdout.decode(cod)
        except UnicodeDecodeError:
            continue
    return ""


def _red(ip: str, prefijo_o_mascara) -> Red | None:
    try:
        return ipaddress.IPv4Network(f"{ip}/{prefijo_o_mascara}", strict=False)
    except ValueError:
        return None


def interpretar_ipconfig(texto: str) -> list[Red]:
    """Extrae (IP, máscara) de la salida de `ipconfig` (español, inglés, francés, alemán…)."""
    redes: list[Red] = []
    ip_pendiente = None
    for linea in texto.splitlines():
        valores = re.findall(_IP, linea)
        if not valores or ":" not in linea:
            continue
        etiqueta = linea.split(":")[0].lower()
        valor = valores[-1]
        if "ipv4" in etiqueta or "ip address" in etiqueta or "dirección ip" in etiqueta:
            ip_pendiente = valor
        elif valor.startswith("255.") and ip_pendiente:
            red = _red(ip_pendiente, valor)
            if red:
                redes.append(red)
            ip_pendiente = None
    return redes


def interpretar_ip_addr(texto: str) -> list[Red]:
    """Salida de `ip -o -f inet addr show` (Linux): «inet 192.168.1.20/24 brd …»."""
    redes = []
    for ip, pref in re.findall(_IP + r"/(\d{1,2})", texto):
        red = _red(ip, pref)
        if red:
            redes.append(red)
    return redes


def _windows() -> list[Red]:
    try:
        salida = _ejecutar(["powershell.exe", "-NoProfile", "-NonInteractive", "-Command",
                            "Get-NetIPAddress -AddressFamily IPv4 | Select-Object IPAddress,PrefixLength | "
                            "ConvertTo-Json -Compress"])
        datos = json.loads(salida[salida.find("["):] if "[" in salida else salida[salida.find("{"):])
        if isinstance(datos, dict):
            datos = [datos]
        redes = [r for r in (_red(d["IPAddress"], d["PrefixLength"]) for d in datos) if r]
        if redes:
            return redes
    except Exception:  # noqa: BLE001
        pass
    try:
        return interpretar_ipconfig(_ejecutar(["ipconfig"]))
    except Exception:  # noqa: BLE001
        return []


def _otros() -> list[Red]:
    for orden in (["ip", "-o", "-f", "inet", "addr", "show"], ["ifconfig"]):
        try:
            texto = _ejecutar(orden)
        except Exception:  # noqa: BLE001
            continue
        redes = interpretar_ip_addr(texto)
        if not redes:  # ifconfig: «inet 192.168.1.20 netmask 255.255.255.0»
            for ip, mascara in re.findall(r"inet (?:addr:)?" + _IP + r".*?(?:netmask|Mask:)\s*" + _IP, texto):
                red = _red(ip, mascara)
                if red:
                    redes.append(red)
        if redes:
            return redes
    return []


def _respaldo() -> list[Red]:
    """Si no se puede leer la configuración de red: la IP principal con máscara /24."""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("10.255.255.255", 1))
        ip = s.getsockname()[0]
        s.close()
        return [ipaddress.IPv4Network(f"{ip}/24", strict=False)]
    except OSError:
        return []


def redes_locales() -> list[Red]:
    """Redes IPv4 a las que está conectado este equipo (sin repetir; incluye 127.0.0.0/8)."""
    redes = _windows() if os.name == "nt" else _otros()
    if not [r for r in redes if not r.is_loopback]:
        redes += _respaldo()
    redes.append(ipaddress.IPv4Network("127.0.0.0/8"))
    unicas: list[Red] = []
    for r in redes:
        if r.prefixlen == 32 and not r.is_loopback:  # direcciones sueltas (p. ej. VPN punto a punto)
            continue
        if r not in unicas:
            unicas.append(r)
    return unicas


def en_alguna_red(ip: str, redes: list[Red]) -> bool:
    try:
        direccion = ipaddress.ip_address(ip.split("%")[0])
    except ValueError:
        return False
    if getattr(direccion, "ipv4_mapped", None):
        direccion = direccion.ipv4_mapped
    if direccion.is_loopback:
        return True
    return any(direccion in r for r in redes if direccion.version == r.version)


def resolver(host: str) -> str | None:
    try:
        return socket.gethostbyname(host)
    except OSError:
        return None


def describir(redes: list[Red]) -> str:
    return ", ".join(str(r) for r in redes if not r.is_loopback) or "sólo este equipo"


class RedesCacheadas:
    """Redes locales recalculadas como mucho cada `segundos` (detectarlas puede tardar ~1 s)."""

    def __init__(self, segundos: float = 30):
        self.segundos = segundos
        self._redes: list[Red] = []
        self._cuando = 0.0

    def obtener(self) -> list[Red]:
        if time.time() - self._cuando > self.segundos:
            self._redes = redes_locales()
            self._cuando = time.time()
        return self._redes

    def fijar(self, redes: list[Red]) -> None:
        self._redes = list(redes)
        self._cuando = time.time() + 10 ** 9  # no volver a calcular (pruebas)
