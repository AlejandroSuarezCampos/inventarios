"""Detección automática de los datos del equipo en el que se ejecuta el programa.

En Windows se consulta WMI/CIM y el registro mediante PowerShell (incluido en Windows 10/11),
así que no hace falta instalar nada. En otros sistemas se obtiene lo que se puede de /proc y /sys.
"""
from __future__ import annotations

import base64
import json
import math
import os
import platform
import re
import socket
import subprocess
from pathlib import Path

_PS = r"""
$ErrorActionPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$r = @{}
$r.nombre = $env:COMPUTERNAME
$r.usuario_env = $env:USERNAME
$cs = Get-CimInstance Win32_ComputerSystem
$r.fabricante = $cs.Manufacturer
$r.modelo = $cs.Model
$r.usuario = $cs.UserName
$r.ram_total = $cs.TotalPhysicalMemory
$csp = Get-CimInstance Win32_ComputerSystemProduct
$r.modelo_version = $csp.Version
$bios = Get-CimInstance Win32_BIOS
$r.serie = $bios.SerialNumber
$os = Get-CimInstance Win32_OperatingSystem
$r.os = $os.Caption
$r.os_version = $os.Version
$cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$r.os_display = $cv.DisplayVersion
$r.cpu = (Get-CimInstance Win32_Processor | Select-Object -First 1).Name
$mem = @(Get-CimInstance Win32_PhysicalMemory)
$r.ram_modulos = @($mem | ForEach-Object { [string]$_.Capacity })
$r.ram_tipos = @($mem | ForEach-Object { $_.SMBIOSMemoryType })
$r.ram_velocidad = @($mem | ForEach-Object { $_.Speed })
$disks = @()
try {
  Get-PhysicalDisk | ForEach-Object {
    $disks += @{ nombre = $_.FriendlyName; tipo = [string]$_.MediaType; tam = [string]$_.Size; bus = [string]$_.BusType }
  }
} catch {}
if ($disks.Count -eq 0) {
  Get-CimInstance Win32_DiskDrive | ForEach-Object {
    $disks += @{ nombre = $_.Model; tipo = [string]$_.MediaType; tam = [string]$_.Size; bus = [string]$_.InterfaceType }
  }
}
$r.discos = $disks
$mons = @()
Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorBasicDisplayParams | ForEach-Object {
  $mons += @{ h = $_.MaxHorizontalImageSize; v = $_.MaxVerticalImageSize }
}
$r.monitores = $mons
$r.antivirus = @(Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct | ForEach-Object { $_.displayName })
$c2r = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
$r.office_c2r = $c2r.ProductReleaseIds
$r.office_version = $c2r.VersionToReport
$office = @()
foreach ($k in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
                 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*')) {
  Get-ItemProperty $k | Where-Object { $_.DisplayName -match '^(Microsoft Office|Microsoft 365|Office 16|LibreOffice)' } |
    ForEach-Object { $office += $_.DisplayName }
}
$r.office_lista = $office
$tv = (Get-ItemProperty 'HKLM:\SOFTWARE\TeamViewer').ClientID
if (-not $tv) { $tv = (Get-ItemProperty 'HKLM:\SOFTWARE\WOW6432Node\TeamViewer').ClientID }
$r.teamviewer = $tv
$r | ConvertTo-Json -Compress -Depth 5
"""

_MARCAS = [
    (r"lenovo", "Lenovo"), (r"\bdell\b", "Dell"), (r"hewlett|^hp\b|\bhp\b", "HP"),
    (r"asus", "ASUS"), (r"acer", "Acer"), (r"micro-star|\bmsi\b", "MSI"), (r"microsoft", "Microsoft"),
    (r"apple", "Apple"), (r"toshiba|dynabook", "Toshiba"), (r"fujitsu", "Fujitsu"), (r"samsung", "Samsung"),
    (r"gigabyte", "Gigabyte"), (r"huawei", "Huawei"), (r"medion", "Medion"), (r"vmware", "VMware"),
    (r"innotek|virtualbox", "VirtualBox"), (r"qemu", "QEMU"),
]

_TIPOS_RAM = {20: "DDR", 21: "DDR2", 22: "DDR2", 24: "DDR3", 26: "DDR4", 27: "LPDDR", 28: "LPDDR2",
              29: "LPDDR3", 30: "LPDDR4", 34: "DDR5", 35: "LPDDR5"}

_SERIES_FALSAS = {"", "0", "none", "default string", "to be filled by o.e.m.", "system serial number",
                  "not applicable", "n/a", "123456789", "chassis serial number"}

_PULGADAS = [10.1, 11.6, 12.5, 13.3, 13.5, 14, 15.6, 16, 17.3, 18.5, 19, 20, 21.5, 22, 23, 23.8, 24, 25,
             27, 28, 31.5, 32, 34]

_TAMANOS_DISCO = [32, 64, 120, 128, 240, 250, 256, 480, 500, 512, 960, 1000, 2000, 4000, 8000]


def _marca(texto: str) -> str:
    t = (texto or "").strip()
    for patron, nombre in _MARCAS:
        if re.search(patron, t, re.I):
            return nombre
    return t.replace(" Inc.", "").replace(" Corporation", "").strip()


def _limpiar_cpu(cpu: str) -> str:
    cpu = re.sub(r"\((R|TM|tm|r)\)", "", cpu or "")
    cpu = re.sub(r"\bCPU\b", "", cpu)
    cpu = re.sub(r"\s+@.*$", "", cpu)
    cpu = re.sub(r"\b(\d+)-Core Processor\b", r"\1 núcleos", cpu)
    cpu = re.sub(r"\bProcessor\b", "", cpu)
    return " ".join(cpu.split())


def _gb(n_bytes: float) -> str:
    gib = n_bytes / 1024 ** 3
    redondeo = int(round(gib))
    # la RAM total que da Windows es algo menor que la instalada: ajustar a potencia/escalón habitual
    for escalon in (2, 4, 6, 8, 12, 16, 20, 24, 32, 48, 64, 96, 128):
        if abs(gib - escalon) / escalon < 0.08:
            redondeo = escalon
            break
    return f"{redondeo} GB"


def _tam_disco(n_bytes: float) -> str:
    gb = n_bytes / 1e9
    if gb <= 0:
        return ""
    cercano = min(_TAMANOS_DISCO, key=lambda t: abs(gb - t) / t)
    if abs(gb - cercano) / cercano < 0.07:
        gb = cercano
    if gb >= 1000:
        tb = gb / 1000
        return f"{tb:g} TB" if tb == int(tb) else f"{tb:.1f} TB"
    return f"{int(round(gb))} GB"


def _lista(x) -> list:
    if x is None:
        return []
    return x if isinstance(x, list) else [x]


def _office(datos: dict) -> str:
    ids = str(datos.get("office_c2r") or "")
    if ids:
        partes = []
        for pid in ids.split(","):
            pid = pid.strip()
            if not pid:
                continue
            if "O365" in pid or "M365" in pid:
                nombre = "Microsoft 365"
            else:
                m = re.search(r"(2016|2019|2021|2024)", pid)
                nombre = f"Office {m.group(1)}" if m else pid
                if "ProPlus" in pid:
                    nombre += " Pro Plus"
                elif "HomeBusiness" in pid:
                    nombre += " Hogar y Empresas"
                elif "HomeStudent" in pid:
                    nombre += " Hogar y Estudiantes"
                elif "Standard" in pid:
                    nombre += " Standard"
            if nombre not in partes and not pid.startswith(("Visio", "Project", "Language", "Proofing")):
                partes.append(nombre)
        if partes:
            return ", ".join(partes)
    lista = [str(x) for x in _lista(datos.get("office_lista")) if x]
    if lista:
        lista.sort(key=len)
        return lista[0].replace(" - es-es", "").replace(" - en-us", "")
    return ""


def interpretar(datos: dict) -> dict:
    """Convierte la información en bruto de Windows en los campos del inventario."""
    r: dict[str, str] = {}
    r["nombre_equipo"] = str(datos.get("nombre") or "").strip()
    usuario = str(datos.get("usuario") or datos.get("usuario_env") or "")
    r["usuario_pc"] = usuario.split("\\")[-1].strip()
    marca = _marca(str(datos.get("fabricante") or ""))
    r["marca"] = marca
    modelo = str(datos.get("modelo") or "").strip()
    version = str(datos.get("modelo_version") or "").strip()
    # Lenovo guarda el nombre comercial ("ThinkPad T14 Gen 2") en Version y un código en Model
    if marca == "Lenovo" and version and version.lower() not in ("none", "lenovo"):
        modelo = version
    r["modelo"] = modelo
    sistema = str(datos.get("os") or "").replace("Microsoft ", "").strip()
    if sistema and datos.get("os_display"):
        sistema += f" {datos['os_display']}"
    r["os"] = sistema
    r["cpu"] = _limpiar_cpu(str(datos.get("cpu") or ""))

    modulos = [float(x) for x in _lista(datos.get("ram_modulos")) if str(x).strip().isdigit()]
    if modulos:
        r["ram"] = _gb(sum(modulos))
    elif datos.get("ram_total"):
        r["ram"] = _gb(float(datos["ram_total"]))
    tipos = [int(t) for t in _lista(datos.get("ram_tipos")) if str(t).isdigit() and int(t) in _TIPOS_RAM]
    if tipos:
        r["tipo_ram"] = _TIPOS_RAM[tipos[0]]
    else:
        # SMBIOSMemoryType = 0 en algunos equipos: deducir por velocidad
        vel = [int(v) for v in _lista(datos.get("ram_velocidad")) if str(v).isdigit()]
        if vel:
            v = max(vel)
            r["tipo_ram"] = "DDR5" if v >= 4400 else "DDR4" if v >= 2133 else "DDR3" if v >= 800 else ""

    discos = []
    for d in _lista(datos.get("discos")):
        if not isinstance(d, dict):
            continue
        bus = str(d.get("bus") or "").upper()
        if bus in ("USB", "7", "SD", "12"):  # 7 = USB en Get-PhysicalDisk
            continue
        try:
            tam = _tam_disco(float(d.get("tam") or 0))
        except ValueError:
            tam = ""
        if not tam:
            continue
        tipo = str(d.get("tipo") or "")
        nombre = str(d.get("nombre") or "")
        if tipo in ("4", "SSD") or "SSD" in nombre.upper() or "NVME" in nombre.upper() or bus in ("NVME", "17"):
            etiqueta = "SSD NVMe" if (bus in ("NVME", "17") or "NVME" in nombre.upper()) else "SSD"
        elif tipo in ("3", "HDD"):
            etiqueta = "HDD"
        else:
            etiqueta = ""
        discos.append(f"{tam} {etiqueta}".strip())
    if discos:
        r["almacenamiento"] = " + ".join(discos)

    r["num_serie"] = ""
    serie = str(datos.get("serie") or "").strip()
    if serie.lower() not in _SERIES_FALSAS:
        r["num_serie"] = serie

    pantallas = []
    for m in _lista(datos.get("monitores")):
        try:
            h, v = float(m.get("h") or 0), float(m.get("v") or 0)
        except (AttributeError, ValueError):
            continue
        if h > 0 and v > 0:
            pulgadas = math.hypot(h, v) / 2.54
            cercana = min(_PULGADAS, key=lambda p: abs(p - pulgadas))
            if abs(cercana - pulgadas) <= 0.5:
                pulgadas = cercana
            pantallas.append(f'{pulgadas:.1f}"'.replace('.0"', '"'))
    if pantallas:
        r["pantalla"] = " + ".join(pantallas)

    r["office"] = _office(datos)
    antivirus = []
    for a in _lista(datos.get("antivirus")):
        a = str(a or "").strip()
        if a and a not in antivirus:
            antivirus.append(a.replace("Windows Defender", "Microsoft Defender"))
    r["antivirus"] = ", ".join(antivirus)
    if datos.get("teamviewer"):
        r["teamviewer"] = str(datos["teamviewer"]).strip()
    return {k: v for k, v in r.items() if v}


def _windows() -> dict:
    script = base64.b64encode(_PS.encode("utf-16-le")).decode("ascii")
    flags = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
    salida = subprocess.run(
        ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
         "-EncodedCommand", script],
        capture_output=True, timeout=90, creationflags=flags)
    texto = salida.stdout.decode("utf-8", errors="replace").strip()
    inicio = texto.find("{")
    if inicio < 0:
        raise RuntimeError("PowerShell no devolvió datos. " + salida.stderr.decode("utf-8", "replace")[:300])
    return json.loads(texto[inicio:])


def _leer(ruta: str) -> str:
    try:
        return Path(ruta).read_text(encoding="utf-8", errors="replace").strip()
    except OSError:
        return ""


def _otro_sistema() -> dict:
    r = {"nombre_equipo": socket.gethostname(), "usuario_pc": os.environ.get("USER", "")}
    r["marca"] = _marca(_leer("/sys/class/dmi/id/sys_vendor"))
    r["modelo"] = _leer("/sys/class/dmi/id/product_name")
    serie = _leer("/sys/class/dmi/id/product_serial")
    if serie.lower() not in _SERIES_FALSAS:
        r["num_serie"] = serie
    r["os"] = f"{platform.system()} {platform.release()}"
    for linea in _leer("/proc/cpuinfo").splitlines():
        if linea.startswith("model name"):
            r["cpu"] = _limpiar_cpu(linea.split(":", 1)[1])
            break
    for linea in _leer("/proc/meminfo").splitlines():
        if linea.startswith("MemTotal"):
            r["ram"] = _gb(int(linea.split()[1]) * 1024)
            break
    return {k: v for k, v in r.items() if v}


def detectar() -> dict:
    """Devuelve {clave_campo: valor} con los datos detectados del equipo actual."""
    if os.name == "nt":
        return interpretar(_windows())
    return _otro_sistema()
