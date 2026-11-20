"""Tema visual moderno (inspirado en macOS) para ttk, sin dependencias externas.

Los elementos redondeados (botones, campos, tarjetas, barras de desplazamiento) se generan
como imágenes PNG con transparencia en tiempo de ejecución y se usan como elementos
«image» de ttk con escalado de 9 partes, de modo que funcionan sobre cualquier fondo.
"""
from __future__ import annotations

import base64
import os
import struct
import sys
import tkinter as tk
import zlib
from tkinter import font as tkfont
from tkinter import ttk

# ------------------------------------------------------------------ paletas
CLARO = {
    "fondo": "#F5F5F7",        # fondo de la zona de contenido
    "lateral": "#E9E9EE",      # barra lateral
    "tarjeta": "#FFFFFF",
    "texto": "#1D1D1F",
    "secundario": "#6E6E73",
    "terciario": "#A1A1A6",
    "borde": "#D2D2D7",
    "borde_suave": "#E5E5EA",
    "acento": "#0071E3",
    "acento_hover": "#0077ED",
    "acento_pulsado": "#0062C4",
    "acento_suave": "#E8F1FD",
    "boton": "#E8E8ED",
    "boton_hover": "#DEDEE3",
    "boton_pulsado": "#D2D2D7",
    "rojo": "#FF3B30",
    "rojo_suave": "#FFE9E7",
    "rojo_pulsado": "#FFD6D2",
    "ambar": "#FF9500",
    "ambar_texto": "#C76E00",
    "ambar_suave": "#FFF4DE",
    "verde": "#34C759",
    "verde_suave": "#E5F8EA",
    "fila_alterna": "#FAFAFC",
    "lateral_sel": "#D9D9E0",
    "lateral_hover": "#E0E0E6",
    "pulgar": "#C7C7CC",
    "pulgar_hover": "#A1A1A6",
    "morado": "#5E5CE6",
}

OSCURO = {
    "fondo": "#1C1C1E",
    "lateral": "#252528",
    "tarjeta": "#2C2C2E",
    "texto": "#F5F5F7",
    "secundario": "#A1A1A6",
    "terciario": "#6E6E73",
    "borde": "#48484A",
    "borde_suave": "#3A3A3C",
    "acento": "#0A84FF",
    "acento_hover": "#2D96FF",
    "acento_pulsado": "#0070DD",
    "acento_suave": "#1A3350",
    "boton": "#3A3A3C",
    "boton_hover": "#48484A",
    "boton_pulsado": "#545456",
    "rojo": "#FF453A",
    "rojo_suave": "#3D2422",
    "rojo_pulsado": "#5C2C28",
    "ambar": "#FF9F0A",
    "ambar_texto": "#FFB340",
    "ambar_suave": "#3B321F",
    "verde": "#30D158",
    "verde_suave": "#1E3B26",
    "fila_alterna": "#313134",
    "lateral_sel": "#3A3A3E",
    "lateral_hover": "#313135",
    "pulgar": "#5A5A5E",
    "pulgar_hover": "#7C7C80",
    "morado": "#7D7AFF",
}

C = dict(CLARO)  # paleta activa (se modifica en sitio con aplicar_modo antes de crear la interfaz)
MODOS = ["Según Windows", "Claro", "Oscuro"]


def windows_en_oscuro() -> bool:
    """¿Tiene Windows activado el modo oscuro para aplicaciones?"""
    if os.name != "nt":
        return False
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as k:
            return winreg.QueryValueEx(k, "AppsUseLightTheme")[0] == 0
    except OSError:
        return False


def aplicar_modo(modo: str) -> bool:
    """Activa la paleta clara u oscura. Devuelve True si queda en oscuro."""
    oscuro = modo == "Oscuro" or (modo not in ("Claro", "Oscuro") and windows_en_oscuro())
    C.clear()
    C.update(OSCURO if oscuro else CLARO)
    return oscuro


def barra_titulo_oscura(ventana: tk.Misc, oscuro: bool) -> None:
    """Pone oscura la barra de título de Windows 10/11 (DWMWA_USE_IMMERSIVE_DARK_MODE)."""
    if os.name != "nt" or not oscuro:
        return
    try:
        import ctypes
        ventana.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(ventana.winfo_id())
        valor = ctypes.c_int(1)
        for atributo in (20, 19):
            if ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, atributo, ctypes.byref(valor),
                                                          ctypes.sizeof(valor)) == 0:
                break
    except Exception:  # noqa: BLE001
        pass


PALETA_GRAFICOS = ["#0071E3", "#5E5CE6", "#30B0C7", "#34C759", "#FF9500", "#FF2D55", "#AF52DE", "#A2845E"]


# ------------------------------------------------------------------ PNG
def _png(ancho: int, alto: int, pixeles: list[bytes]) -> bytes:
    raw = b"".join(b"\x00" + fila for fila in pixeles)

    def trozo(tipo: bytes, datos: bytes) -> bytes:
        return struct.pack(">I", len(datos)) + tipo + datos + struct.pack(">I", zlib.crc32(tipo + datos) & 0xFFFFFFFF)

    return (b"\x89PNG\r\n\x1a\n" + trozo(b"IHDR", struct.pack(">IIBBBBB", ancho, alto, 8, 6, 0, 0, 0))
            + trozo(b"IDAT", zlib.compress(raw, 9)) + trozo(b"IEND", b""))


def _rgb(color: str) -> tuple[int, int, int]:
    color = color.lstrip("#")
    return int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)


def _cobertura(x: float, y: float, w: float, h: float, r: float) -> bool:
    """¿El punto (x, y) cae dentro del rectángulo redondeado (0,0)-(w,h) de radio r?"""
    if x < 0 or y < 0 or x > w or y > h:
        return False
    cx = r if x < r else (w - r if x > w - r else x)
    cy = r if y < r else (h - r if y > h - r else y)
    return (x - cx) ** 2 + (y - cy) ** 2 <= r * r


def rect_redondeado(w: int, h: int, r: float, relleno: str | None, borde: str | None = None,
                    grosor: float = 1.0, ss: int = 4) -> bytes:
    """PNG RGBA de un rectángulo redondeado antialias con borde opcional."""
    fr = _rgb(relleno) if relleno else None
    br = _rgb(borde) if borde else None
    filas = []
    muestras = [(i + 0.5) / ss for i in range(ss)]
    total = ss * ss
    for py in range(h):
        fila = bytearray()
        for px in range(w):
            dentro = interior = 0
            for sy in muestras:
                for sx in muestras:
                    x, y = px + sx, py + sy
                    if _cobertura(x, y, w, h, r):
                        dentro += 1
                        if br is None or _cobertura(x - grosor, y - grosor, w - 2 * grosor, h - 2 * grosor,
                                                    max(0.0, r - grosor)):
                            interior += 1
            if dentro == 0:
                fila += b"\x00\x00\x00\x00"
                continue
            a_borde = (dentro - interior) / total
            a_relleno = interior / total if fr else 0.0
            alfa = a_borde + a_relleno
            if alfa <= 0:
                fila += b"\x00\x00\x00\x00"
                continue
            if br and fr:
                col = tuple((br[i] * a_borde + fr[i] * a_relleno) / alfa for i in range(3))
            elif br:
                col = br
            else:
                col = fr
            fila += bytes((int(col[0]), int(col[1]), int(col[2]), int(round(255 * alfa))))
        filas.append(bytes(fila))
    return _png(w, h, filas)


def chevron(color: str, tam: int = 10) -> bytes:
    """Flecha hacia abajo (para combobox)."""
    filas = []
    rgb = _rgb(color)
    ss = 4
    for py in range(tam):
        fila = bytearray()
        for px in range(tam):
            n = 0
            for sy in range(ss):
                for sx in range(ss):
                    x = (px + (sx + 0.5) / ss) / tam
                    y = (py + (sy + 0.5) / ss) / tam
                    # V entre y=0.3 y y=0.7, trazo de grosor ~0.16
                    yc = 0.32 + 0.38 * (1 - abs(x - 0.5) * 2)
                    if 0.12 < x < 0.88 and abs(y - yc) < 0.09:
                        n += 1
            fila += bytes((*rgb, int(255 * n / (ss * ss))))
        filas.append(bytes(fila))
    return _png(tam, tam, filas)


def _dist_segmento(px_, py_, ax, ay, bx, by) -> float:
    dx, dy = bx - ax, by - ay
    t = max(0.0, min(1.0, ((px_ - ax) * dx + (py_ - ay) * dy) / (dx * dx + dy * dy)))
    cx, cy = ax + t * dx, ay + t * dy
    return ((px_ - cx) ** 2 + (py_ - cy) ** 2) ** 0.5


def indicador(tam: int, margen: int, tipo: str, marcado: bool, fondo: str, borde: str, acento: str,
              marca: str = "#FFFFFF", ss: int = 4) -> bytes:
    """Casilla de verificación (tipo='casilla') o botón de opción (tipo='radio') antialias.
    Se deja `margen` píxeles transparentes a la derecha para separarlo del texto."""
    w, h = tam + margen, tam
    r_esq = tam * 0.28 if tipo == "casilla" else tam / 2
    relleno = _rgb(acento if marcado else fondo)
    linea = _rgb(acento if marcado else borde)
    blanco = _rgb(marca)
    filas = []
    for py in range(h):
        fila = bytearray()
        for px_ in range(w):
            cob = cob_int = cob_marca = 0
            for sy in range(ss):
                for sx in range(ss):
                    x, y = px_ + (sx + 0.5) / ss, py + (sy + 0.5) / ss
                    if x > tam:
                        continue
                    if _cobertura(x, y, tam, tam, r_esq):
                        cob += 1
                        if _cobertura(x - 1, y - 1, tam - 2, tam - 2, max(0.0, r_esq - 1)):
                            cob_int += 1
                            if marcado:
                                u, v = x / tam, y / tam
                                if tipo == "casilla":
                                    d = min(_dist_segmento(u, v, 0.26, 0.52, 0.43, 0.69),
                                            _dist_segmento(u, v, 0.43, 0.69, 0.75, 0.33))
                                    if d < 0.075:
                                        cob_marca += 1
                                elif (u - 0.5) ** 2 + (v - 0.5) ** 2 < 0.17 ** 2:
                                    cob_marca += 1
            total = ss * ss
            if cob == 0:
                fila += b"\x00\x00\x00\x00"
                continue
            a_borde = (cob - cob_int) / total
            a_marca = cob_marca / total
            a_rell = (cob_int - cob_marca) / total
            alfa = a_borde + a_marca + a_rell
            col = [(linea[i] * a_borde + blanco[i] * a_marca + relleno[i] * a_rell) / alfa for i in range(3)]
            fila += bytes((int(col[0]), int(col[1]), int(col[2]), int(round(255 * alfa))))
        filas.append(bytes(fila))
    return _png(w, h, filas)


# ------------------------------------------------------------------ tipografía
def familia_fuente(raiz: tk.Misc) -> str:
    disponibles = set(tkfont.families(raiz))
    if sys.platform == "darwin":
        preferidas = ("SF Pro Text", ".AppleSystemUIFont", "Helvetica Neue")
    elif os.name == "nt":
        preferidas = ("Segoe UI Variable Text", "Segoe UI")
    else:
        preferidas = ("Inter", "SF Pro Text", "Cantarell", "Noto Sans", "DejaVu Sans")
    for f in preferidas:
        if f in disponibles:
            return f
    return tkfont.nametofont("TkDefaultFont", raiz).actual("family")


class Tema:
    """Aplica el tema a una ventana raíz y guarda las imágenes (para que no las libere el GC)."""

    def __init__(self, raiz: tk.Tk):
        self.raiz = raiz
        self.imagenes: dict[str, tk.PhotoImage] = {}
        try:
            self.escala = max(1.0, float(raiz.winfo_fpixels("1i")) / 96.0)
        except tk.TclError:
            self.escala = 1.0
        self.familia = familia_fuente(raiz)
        f = self.familia
        self.fuente = (f, 10)
        self.fuente_sm = (f, 9)
        self.fuente_bold = (f, 10, "bold")
        self.fuente_titulo = (f, 22, "bold")
        self.fuente_subtitulo = (f, 13, "bold")
        self.fuente_cifra = (f, 26, "bold")
        self.fuente_seccion = (f, 9, "bold")
        self._fuentes_base()
        self._estilo()

    def px(self, v: float) -> int:
        return max(1, int(round(v * self.escala)))

    def img(self, nombre: str, datos_png: bytes) -> tk.PhotoImage:
        im = tk.PhotoImage(master=self.raiz, data=base64.b64encode(datos_png).decode("ascii"), format="png")
        self.imagenes[nombre] = im
        return im

    def rr(self, nombre: str, r: float, relleno: str | None, borde: str | None = None, grosor: float = 1.0):
        r = self.px(r)
        tam = 2 * r + self.px(4)
        return self.img(nombre, rect_redondeado(tam, tam, r, relleno, borde, grosor * self.escala))

    # -------------------------------------------------------------- fuentes
    def _fuentes_base(self):
        for nombre in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont", "TkCaptionFont",
                       "TkSmallCaptionFont", "TkIconFont", "TkTooltipFont"):
            try:
                tkfont.nametofont(nombre, self.raiz).configure(family=self.familia, size=10)
            except tk.TclError:
                pass
        r = self.raiz
        r.option_add("*TCombobox*Listbox.font", self.fuente)
        r.option_add("*TCombobox*Listbox.background", C["tarjeta"])
        r.option_add("*TCombobox*Listbox.foreground", C["texto"])
        r.option_add("*TCombobox*Listbox.selectBackground", C["acento"])
        r.option_add("*TCombobox*Listbox.selectForeground", "#FFFFFF")
        r.option_add("*TCombobox*Listbox.borderWidth", 0)
        r.option_add("*TCombobox*Listbox.relief", "flat")
        r.option_add("*Menu.background", C["tarjeta"])
        r.option_add("*Menu.activeBackground", C["acento"])
        r.option_add("*Menu.activeForeground", "#FFFFFF")
        r.option_add("*Menu.relief", "flat")
        r.option_add("*Menu.font", self.fuente)

    # -------------------------------------------------------------- estilos
    def _estilo(self):
        s = ttk.Style(self.raiz)
        s.theme_use("clam")
        self.raiz.configure(background=C["fondo"])
        pad = self.px

        s.configure(".", background=C["fondo"], foreground=C["texto"], font=self.fuente,
                    bordercolor=C["borde"], lightcolor=C["fondo"], darkcolor=C["fondo"],
                    troughcolor=C["fondo"], focuscolor=C["acento"], selectbackground=C["acento"],
                    selectforeground="#FFFFFF", insertcolor=C["texto"])

        # ---- marcos y etiquetas
        s.configure("TFrame", background=C["fondo"])
        s.configure("Lateral.TFrame", background=C["lateral"])
        s.configure("TLabel", background=C["fondo"], foreground=C["texto"])
        s.configure("Lateral.TLabel", background=C["lateral"], foreground=C["secundario"], font=self.fuente_sm)
        s.configure("LateralTitulo.TLabel", background=C["lateral"], foreground=C["texto"], font=(self.familia, 12, "bold"))
        s.configure("LateralSeccion.TLabel", background=C["lateral"], foreground=C["terciario"], font=self.fuente_seccion)
        s.configure("Titulo.TLabel", font=self.fuente_titulo)
        s.configure("Subtitulo.TLabel", foreground=C["secundario"])
        s.configure("Secundario.TLabel", foreground=C["secundario"], font=self.fuente_sm)
        s.configure("Seccion.TLabel", foreground=C["secundario"], font=self.fuente_seccion)

        # ---- tarjetas (fondo blanco redondeado con borde sutil)
        tarjeta = self.rr("tarjeta", 12, C["tarjeta"], C["borde_suave"])
        s.element_create("Tarjeta.bg", "image", tarjeta, border=pad(12), sticky="nsew")
        s.layout("Tarjeta.TFrame", [("Tarjeta.bg", {"sticky": "nsew"})])
        s.configure("Tarjeta.TFrame", padding=pad(14))
        s.configure("EnTarjeta.TFrame", background=C["tarjeta"])
        s.configure("EnTarjeta.TLabel", background=C["tarjeta"], foreground=C["texto"])
        s.configure("EnTarjetaSec.TLabel", background=C["tarjeta"], foreground=C["secundario"], font=self.fuente_sm)
        s.configure("EnTarjetaCampo.TLabel", background=C["tarjeta"], foreground=C["secundario"])
        s.configure("EnTarjetaTitulo.TLabel", background=C["tarjeta"], foreground=C["texto"], font=self.fuente_subtitulo)
        s.configure("Cifra.TLabel", background=C["tarjeta"], font=self.fuente_cifra)
        for nombre, color in (("Rojo", C["rojo"]), ("Ambar", C["ambar"]), ("Verde", C["verde"]),
                              ("Azul", C["acento"]), ("Gris", C["secundario"])):
            s.configure(f"Cifra{nombre}.TLabel", background=C["tarjeta"], foreground=color, font=self.fuente_cifra)

        # ---- botones
        self._boton(s, "TButton", C["boton"], C["boton_hover"], C["boton_pulsado"], C["texto"])
        self._boton(s, "Acento.TButton", C["acento"], C["acento_hover"], C["acento_pulsado"], "#FFFFFF", negrita=True)
        self._boton(s, "Peligro.TButton", C["boton"], C["rojo_suave"], C["rojo_pulsado"], C["rojo"])
        # botón plano (sólo texto, fondo al pasar el ratón)
        transp = self.rr("plano_n", 7, None)
        hover = self.rr("plano_h", 7, C["boton"])
        puls = self.rr("plano_p", 7, C["boton_hover"])
        s.element_create("Plano.bg", "image", transp, ("disabled", transp), ("pressed", puls), ("active", hover),
                         border=pad(7), sticky="nsew")
        s.layout("Plano.TButton", [("Plano.bg", {"sticky": "nsew", "children": [
            ("Button.padding", {"sticky": "nsew", "children": [("Button.label", {"sticky": "nsew"})]})]})])
        s.configure("Plano.TButton", padding=(pad(10), pad(6)), foreground=C["acento"], anchor="center")
        s.map("Plano.TButton", foreground=[("disabled", C["terciario"])])

        # ---- barra lateral (elementos de navegación)
        n = self.rr("lat_n", 7, None)
        h = self.rr("lat_h", 7, C["lateral_hover"])
        sel = self.rr("lat_s", 7, C["lateral_sel"])
        s.element_create("Lateral.bg", "image", n, ("disabled", n), ("selected", sel), ("pressed", sel), ("active", h),
                         border=pad(7), sticky="nsew")
        s.layout("Lateral.TButton", [("Lateral.bg", {"sticky": "nsew", "children": [
            ("Button.padding", {"sticky": "nsew", "children": [("Button.label", {"sticky": "w"})]})]})])
        s.configure("Lateral.TButton", padding=(pad(10), pad(7)), background=C["lateral"], anchor="w",
                    foreground=C["texto"])

        # ---- control segmentado (pestañas de las fichas)
        seg_n = self.rr("seg_n", 7, None)
        seg_h = self.rr("seg_h", 7, C["boton"])
        seg_s = self.rr("seg_s", 7, C["tarjeta"], C["borde_suave"])
        s.element_create("Segmento.bg", "image", seg_n, ("disabled", seg_n), ("selected", seg_s), ("active", seg_h),
                         border=pad(7), sticky="nsew")
        s.layout("Segmento.TButton", [("Segmento.bg", {"sticky": "nsew", "children": [
            ("Button.padding", {"sticky": "nsew", "children": [("Button.label", {"sticky": "nsew"})]})]})])
        s.configure("Segmento.TButton", padding=(pad(16), pad(6)), foreground=C["secundario"], anchor="center")
        s.map("Segmento.TButton", foreground=[("disabled", C["terciario"]), ("selected", C["texto"])],
              background=[("disabled", C["fondo"]), ("active", C["fondo"]), ("selected", C["fondo"]),
                          ("pressed", C["fondo"])])

        # ---- campos de texto
        campo = self.rr("campo", 7, C["tarjeta"], C["borde"])
        campo_f = self.rr("campo_f", 7, C["tarjeta"], C["acento"], 2.0)
        campo_d = self.rr("campo_d", 7, C["fondo"], C["borde_suave"])
        s.element_create("Campo.bg", "image", campo, ("focus", campo_f), ("disabled", campo_d),
                         border=pad(8), sticky="nsew")
        s.layout("TEntry", [("Campo.bg", {"sticky": "nsew", "children": [
            ("Entry.padding", {"sticky": "nsew", "children": [("Entry.textarea", {"sticky": "nsew"})]})]})])
        s.configure("TEntry", padding=(pad(9), pad(5)), foreground=C["texto"], fieldbackground=C["tarjeta"])

        flecha = self.img("chevron", chevron(C["secundario"], pad(10)))
        s.element_create("Flecha.img", "image", flecha, sticky="")
        s.layout("TCombobox", [("Campo.bg", {"sticky": "nsew", "children": [
            ("Flecha.img", {"side": "right", "sticky": ""}),
            ("Combobox.padding", {"sticky": "nsew", "children": [("Combobox.textarea", {"sticky": "nsew"})]})]})])
        s.configure("TCombobox", padding=(pad(9), pad(5), pad(9), pad(5)), foreground=C["texto"],
                    fieldbackground=C["tarjeta"], background=C["fondo"], arrowsize=0)
        s.map("TCombobox", fieldbackground=[("readonly", C["tarjeta"])],
              background=[("readonly", C["fondo"]), ("disabled", C["fondo"]), ("focus", C["fondo"])],
              selectbackground=[("readonly", C["tarjeta"]), ("!focus", C["tarjeta"])],
              selectforeground=[("readonly", C["texto"]), ("!focus", C["texto"])],
              foreground=[("disabled", C["terciario"])])

        # ---- tabla
        s.layout("Treeview", [("Treeview.treearea", {"sticky": "nsew"})])
        s.configure("Treeview", background=C["tarjeta"], fieldbackground=C["tarjeta"], foreground=C["texto"],
                    rowheight=pad(30), borderwidth=0, relief="flat", font=self.fuente)
        s.map("Treeview", background=[("selected", C["acento"])], foreground=[("selected", "#FFFFFF")])
        s.configure("Treeview.Heading", background=C["tarjeta"], foreground=C["secundario"],
                    font=self.fuente_seccion, relief="flat", borderwidth=0, padding=(pad(8), pad(8)),
                    lightcolor=C["tarjeta"], darkcolor=C["tarjeta"], bordercolor=C["borde_suave"])
        s.map("Treeview.Heading", background=[("active", C["fondo"])], foreground=[("active", C["texto"])])

        # ---- barras de desplazamiento finas
        pulgar = self.rr("pulgar", 4, C["pulgar"])
        pulgar_h = self.rr("pulgar_h", 4, C["pulgar_hover"])
        for orient in ("Vertical", "Horizontal"):
            s.element_create(f"{orient}.Fino.thumb", "image", pulgar, ("active", pulgar_h), ("pressed", pulgar_h),
                             border=pad(4), sticky="nsew")
            s.layout(f"{orient}.TScrollbar", [(f"{orient}.Scrollbar.trough", {"sticky": "nsew", "children": [
                (f"{orient}.Fino.thumb", {"expand": "1", "sticky": "nsew"})]})])
            s.configure(f"{orient}.TScrollbar", troughcolor=C["tarjeta"], background=C["tarjeta"],
                        bordercolor=C["tarjeta"], lightcolor=C["tarjeta"], darkcolor=C["tarjeta"],
                        arrowsize=pad(9), width=pad(9), gripcount=0, relief="flat")

        # ---- casillas y radios (indicadores dibujados)
        tam, margen = pad(16), pad(8)
        for tipo, clase in (("casilla", "Checkbutton"), ("radio", "Radiobutton")):
            imgs = {}
            for nombre, marcado, fondo_i, borde_i, acento_i in (
                    ("off", False, C["tarjeta"], C["borde"], C["acento"]),
                    ("on", True, C["tarjeta"], C["borde"], C["acento"]),
                    ("off_d", False, C["fondo"], C["borde_suave"], C["terciario"]),
                    ("on_d", True, C["fondo"], C["borde_suave"], C["terciario"])):
                imgs[nombre] = self.img(f"{tipo}_{nombre}",
                                        indicador(tam, margen, tipo, marcado, fondo_i, borde_i, acento_i))
            s.element_create(f"{clase}.ind", "image", imgs["off"],
                             ("disabled", "selected", imgs["on_d"]), ("disabled", imgs["off_d"]),
                             ("selected", imgs["on"]), sticky="")
            s.layout(f"T{clase}", [(f"{clase}.padding", {"sticky": "nswe", "children": [
                (f"{clase}.ind", {"side": "left", "sticky": ""}),
                (f"{clase}.label", {"side": "left", "sticky": "nswe"})]})])
            s.configure(f"T{clase}", background=C["tarjeta"], foreground=C["texto"], padding=(0, pad(3)),
                        focuscolor=C["tarjeta"])
            s.map(f"T{clase}", background=[("active", C["tarjeta"]), ("disabled", C["tarjeta"])],
                  foreground=[("disabled", C["terciario"])])

        # ttk rellena las zonas transparentes con el «background» del estilo: el fondo debe coincidir
        # con el del contenedor. Variantes «Blanco.*» para widgets dentro de tarjetas.
        s.map(".", background=[("disabled", C["fondo"])])
        for base in ("TButton", "Acento.TButton", "Peligro.TButton", "Plano.TButton"):
            s.map(base, background=[("disabled", C["fondo"]), ("active", C["fondo"]), ("pressed", C["fondo"])])
            s.configure(f"Blanco.{base}", background=C["tarjeta"])
            s.map(f"Blanco.{base}", background=[("disabled", C["tarjeta"]), ("active", C["tarjeta"]),
                                                ("pressed", C["tarjeta"])])
        for base in ("TEntry", "TCombobox"):
            s.configure(f"Blanco.{base}", background=C["tarjeta"])
            s.map(f"Blanco.{base}", background=[("disabled", C["tarjeta"]), ("readonly", C["tarjeta"]),
                                                ("focus", C["tarjeta"])])
        s.map("TEntry", background=[("disabled", C["fondo"]), ("focus", C["fondo"])])
        s.map("Lateral.TButton", background=[("disabled", C["lateral"]), ("active", C["lateral"]),
                                             ("pressed", C["lateral"]), ("selected", C["lateral"])])

        s.configure("Rojo.Plano.TButton", foreground=C["rojo"])
        s.map("Rojo.Plano.TButton", foreground=[("disabled", C["terciario"])])
        s.configure("TSeparator", background=C["borde_suave"])
        s.configure("TPanedwindow", background=C["fondo"])

    def _boton(self, s: ttk.Style, estilo: str, n: str, h: str, p: str, texto: str, negrita: bool = False):
        base = estilo.replace(".TButton", "").replace("TButton", "Base")
        im_n = self.rr(f"{base}_n", 7, n)
        im_h = self.rr(f"{base}_h", 7, h)
        im_p = self.rr(f"{base}_p", 7, p)
        im_d = self.rr(f"{base}_d", 7, C["borde_suave"])
        s.element_create(f"{base}.bg", "image", im_n, ("disabled", im_d), ("pressed", im_p), ("active", im_h),
                         border=self.px(8), sticky="nsew")
        s.layout(estilo, [(f"{base}.bg", {"sticky": "nsew", "children": [
            ("Button.padding", {"sticky": "nsew", "children": [("Button.label", {"sticky": "nsew"})]})]})])
        s.configure(estilo, padding=(self.px(14), self.px(7)), foreground=texto, anchor="center",
                    font=self.fuente_bold if negrita else self.fuente)
        s.map(estilo, foreground=[("disabled", C["terciario"])])


class GraficoBarras(tk.Canvas):
    """Gráfico de barras horizontales minimalista. Llama a `al_pulsar(valor)` al hacer clic en una fila."""

    def __init__(self, padre, tema: Tema, color: str = C["acento"], al_pulsar=None, max_filas: int = 7):
        super().__init__(padre, background=C["tarjeta"], highlightthickness=0, borderwidth=0,
                         height=tema.px(28) * max_filas)
        self.tema = tema
        self.color = color
        self.al_pulsar = al_pulsar
        self.max_filas = max_filas
        self.datos: list[tuple[str, int]] = []
        self.bind("<Configure>", lambda e: self.dibujar())
        self.bind("<Motion>", self._mover)
        self.bind("<Leave>", lambda e: self.configure(cursor=""))
        self.bind("<Button-1>", self._pulsar)
        self._filas: list[tuple[int, int, str]] = []

    def poner_datos(self, datos: list[tuple[str, int]]):
        if len(datos) > self.max_filas:
            resto = sum(n for _, n in datos[self.max_filas - 1:])
            datos = datos[: self.max_filas - 1] + [(f"Otros ({len(datos) - self.max_filas + 1})", resto)]
        self.datos = datos
        self.dibujar()

    def dibujar(self):
        self.delete("all")
        self._filas = []
        t = self.tema
        ancho = max(self.winfo_width(), 200)
        if not self.datos:
            self.create_text(ancho / 2, t.px(40), text="Sin datos", fill=C["terciario"], font=t.fuente)
            return
        alto_fila = t.px(28)
        etiqueta_w = min(int(ancho * 0.42), t.px(170))
        num_w = t.px(40)
        barra_max = max(10, ancho - etiqueta_w - num_w - t.px(12))
        maximo = max(n for _, n in self.datos) or 1
        fuente = t.fuente
        for i, (valor, n) in enumerate(self.datos):
            y = i * alto_fila + alto_fila / 2
            texto = str(valor)
            # recortar etiquetas largas
            f = tkfont.Font(family=fuente[0], size=fuente[1])
            while f.measure(texto) > etiqueta_w - t.px(8) and len(texto) > 3:
                texto = texto[:-2] + "…" if not texto.endswith("…") else texto[:-2] + "…"
            self.create_text(0, y, text=texto, anchor="w", fill=C["texto"], font=fuente)
            largo = max(t.px(6), barra_max * n / maximo)
            x0 = etiqueta_w
            h = t.px(10)
            self._barra(x0, y - h / 2, x0 + largo, y + h / 2, self.color)
            self.create_text(ancho, y, text=str(n), anchor="e", fill=C["secundario"], font=t.fuente_bold)
            self._filas.append((int(i * alto_fila), int((i + 1) * alto_fila), str(valor)))

    def _barra(self, x0, y0, x1, y1, color):
        r = (y1 - y0) / 2
        if x1 - x0 < 2 * r:
            x1 = x0 + 2 * r
        self.create_oval(x0, y0, x0 + 2 * r, y1, fill=color, outline=color)
        self.create_oval(x1 - 2 * r, y0, x1, y1, fill=color, outline=color)
        self.create_rectangle(x0 + r, y0, x1 - r, y1, fill=color, outline=color)

    def _fila_en(self, y) -> str | None:
        for a, b, v in self._filas:
            if a <= y < b:
                return v
        return None

    def _mover(self, e):
        v = self._fila_en(e.y)
        clicable = self.al_pulsar and v and not v.startswith("Otros (") and v != "(sin dato)"
        self.configure(cursor="hand2" if clicable else "")

    def _pulsar(self, e):
        v = self._fila_en(e.y)
        if self.al_pulsar and v and not v.startswith("Otros (") and v != "(sin dato)":
            self.al_pulsar(v)


class CampoBusqueda(ttk.Frame):
    """Entrada con icono de lupa y texto de ayuda (placeholder)."""

    def __init__(self, padre, tema: Tema, variable: tk.StringVar, ayuda: str = "Buscar", ancho: int = 28):
        super().__init__(padre)
        self.var = variable
        self.entrada = ttk.Entry(self, textvariable=variable, width=ancho, style="Busqueda.TEntry")
        s = ttk.Style(self)
        s.configure("Busqueda.TEntry", padding=(tema.px(30), tema.px(6), tema.px(9), tema.px(6)))
        self.entrada.pack(fill="x", expand=True)
        self.lupa = tk.Label(self, text="⌕", bg=C["tarjeta"], fg=C["secundario"],
                             font=(tema.familia, 14), bd=0, padx=0, pady=0)
        self.lupa.place(x=tema.px(9), rely=0.5, anchor="w")
        self.ayuda = tk.Label(self, text=ayuda, bg=C["tarjeta"], fg=C["terciario"], font=tema.fuente, bd=0)
        self.ayuda.place(x=tema.px(31), rely=0.5, anchor="w")
        for w in (self.lupa, self.ayuda):
            w.bind("<Button-1>", lambda e: self.entrada.focus_set())
        variable.trace_add("write", lambda *a: self._actualizar())
        self._actualizar()

    def _actualizar(self):
        if self.var.get():
            self.ayuda.place_forget()
        else:
            self.ayuda.place(x=self.tema_x(), rely=0.5, anchor="w")

    def tema_x(self):
        return int(self.lupa.place_info().get("x", 9)) + 22

    def focus_set(self):
        self.entrada.focus_set()
        self.entrada.select_range(0, "end")

    def bind_entrada(self, evento, func):
        self.entrada.bind(evento, func)
