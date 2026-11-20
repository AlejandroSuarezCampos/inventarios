"""Generador de PDF mínimo (sin dependencias) para informes y etiquetas.

Usa las fuentes estándar Helvetica / Helvetica-Bold (incluidas en cualquier visor de PDF) con
codificación WinAnsi, que cubre los caracteres del español (á, é, ñ, ü, ¿, ¡, €…).
Las coordenadas son en puntos (1/72") con el origen en la esquina SUPERIOR izquierda.
"""
from __future__ import annotations

import zlib
from pathlib import Path

A4 = (595.28, 841.89)
MM = 72 / 25.4

# Anchos (1/1000 em) de los caracteres 32-255 en cp1252 (WinAnsiEncoding), de las métricas AFM estándar.
_ANCHOS_HELVETICA = [
    278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556,
    556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556, 1015, 667, 667, 722, 722, 667,
    611, 778, 722, 278, 500, 667, 556, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667,
    667, 611, 278, 278, 278, 469, 556, 333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500,
    222, 833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584,
    761, 556, 0, 222, 556, 333, 1000, 556, 556, 333, 1000, 667, 333, 1000, 0, 611, 0, 0, 222, 222,
    333, 333, 350, 556, 1000, 333, 1000, 500, 333, 944, 0, 500, 667, 278, 333, 556, 556, 556, 556,
    260, 556, 333, 737, 370, 556, 584, 333, 737, 333, 400, 584, 333, 333, 333, 556, 537, 278, 333,
    333, 365, 556, 834, 834, 834, 611, 667, 667, 667, 667, 667, 667, 1000, 722, 667, 667, 667, 667,
    278, 278, 278, 278, 722, 722, 778, 778, 778, 778, 778, 584, 778, 722, 722, 722, 722, 667, 667,
    611, 556, 556, 556, 556, 556, 556, 889, 500, 556, 556, 556, 556, 278, 278, 278, 278, 556, 556,
    556, 556, 556, 556, 556, 584, 611, 556, 556, 556, 556, 500, 556, 500
]
_ANCHOS_HELVETICA_BOLD = [
    278, 333, 474, 556, 556, 889, 722, 238, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556,
    556, 556, 556, 556, 556, 556, 556, 333, 333, 584, 584, 584, 611, 975, 722, 722, 722, 722, 667,
    611, 778, 722, 278, 556, 722, 611, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667,
    667, 611, 333, 278, 333, 584, 556, 333, 556, 611, 556, 611, 556, 333, 611, 611, 278, 278, 556,
    278, 889, 611, 611, 611, 611, 389, 556, 333, 611, 556, 778, 556, 556, 500, 389, 280, 389, 584,
    761, 556, 0, 278, 556, 500, 1000, 556, 556, 333, 1000, 667, 333, 1000, 0, 611, 0, 0, 278, 278,
    500, 500, 350, 556, 1000, 333, 1000, 556, 333, 944, 0, 500, 667, 278, 333, 556, 556, 556, 556,
    280, 556, 333, 737, 370, 556, 584, 333, 737, 333, 400, 584, 333, 333, 333, 611, 556, 278, 333,
    333, 365, 556, 834, 834, 834, 611, 722, 722, 722, 722, 722, 722, 1000, 722, 667, 667, 667, 667,
    278, 278, 278, 278, 722, 722, 778, 778, 778, 778, 778, 584, 778, 722, 722, 722, 722, 667, 667,
    611, 556, 556, 556, 556, 556, 556, 889, 556, 556, 556, 556, 556, 278, 278, 278, 278, 611, 611,
    611, 611, 611, 611, 611, 584, 611, 611, 611, 611, 611, 556, 611, 556
]


def _rgb(color: str) -> tuple[float, float, float]:
    c = color.lstrip("#")
    return int(c[0:2], 16) / 255, int(c[2:4], 16) / 255, int(c[4:6], 16) / 255


def _cod(texto: str) -> bytes:
    return str(texto).encode("cp1252", errors="replace")


def _esc(b: bytes) -> bytes:
    return b.replace(b"\\", b"\\\\").replace(b"(", b"\\(").replace(b")", b"\\)").replace(b"\r", b"").replace(b"\n", b" ")


class PDF:
    def __init__(self, tamano: tuple[float, float] = A4, titulo: str = "", autor: str = ""):
        self.ancho, self.alto = tamano
        self.titulo = titulo
        self.autor = autor
        self.paginas: list[list[bytes]] = []
        self._actual: list[bytes] | None = None

    # ------------------------------------------------------------ páginas
    def nueva_pagina(self) -> None:
        self._actual = []
        self.paginas.append(self._actual)

    def _op(self, s: str | bytes) -> None:
        if self._actual is None:
            self.nueva_pagina()
        self._actual.append(s.encode("latin-1") if isinstance(s, str) else s)

    def _y(self, y: float) -> float:
        return self.alto - y

    # ------------------------------------------------------------ texto
    @staticmethod
    def ancho_texto(texto: str, tam: float, negrita: bool = False) -> float:
        tabla = _ANCHOS_HELVETICA_BOLD if negrita else _ANCHOS_HELVETICA
        total = 0
        for b in _cod(texto):
            total += tabla[b - 32] if 32 <= b <= 255 else 500
        return total * tam / 1000

    def recortar(self, texto: str, ancho: float, tam: float, negrita: bool = False) -> str:
        texto = str(texto or "").replace("\n", " ")
        if self.ancho_texto(texto, tam, negrita) <= ancho:
            return texto
        while texto and self.ancho_texto(texto + "…", tam, negrita) > ancho:
            texto = texto[:-1]
        return texto.rstrip() + "…" if texto else ""

    def partir(self, texto: str, ancho: float, tam: float, negrita: bool = False) -> list[str]:
        """Divide un texto en líneas que caben en `ancho`."""
        lineas, actual = [], ""
        for palabra in str(texto or "").split():
            prueba = f"{actual} {palabra}".strip()
            if self.ancho_texto(prueba, tam, negrita) <= ancho or not actual:
                actual = prueba
            else:
                lineas.append(actual)
                actual = palabra
        if actual:
            lineas.append(actual)
        return lineas

    def texto(self, x: float, y: float, texto: str, tam: float = 10, negrita: bool = False,
              color: str = "#1D1D1F", alin: str = "izq", ancho_max: float | None = None) -> float:
        """Escribe texto con la LÍNEA BASE en y. Devuelve el ancho escrito."""
        if ancho_max:
            texto = self.recortar(texto, ancho_max, tam, negrita)
        w = self.ancho_texto(texto, tam, negrita)
        if alin == "der":
            x -= w
        elif alin == "centro":
            x -= w / 2
        r, g, b = _rgb(color)
        fuente = "F2" if negrita else "F1"
        self._op(f"BT /{fuente} {tam:.2f} Tf {r:.3f} {g:.3f} {b:.3f} rg {x:.2f} {self._y(y):.2f} Td (".encode()
                 + _esc(_cod(texto)) + b") Tj ET")
        return w

    # ------------------------------------------------------------ formas
    def rect(self, x: float, y: float, w: float, h: float, relleno: str | None = None,
             borde: str | None = None, grosor: float = 0.75, radio: float = 0) -> None:
        if relleno is None and borde is None:
            return
        partes = []
        if relleno:
            r, g, b = _rgb(relleno)
            partes.append(f"{r:.3f} {g:.3f} {b:.3f} rg")
        if borde:
            r, g, b = _rgb(borde)
            partes.append(f"{r:.3f} {g:.3f} {b:.3f} RG {grosor:.2f} w")
        x0, y0 = x, self._y(y + h)
        if radio <= 0:
            partes.append(f"{x0:.2f} {y0:.2f} {w:.2f} {h:.2f} re")
        else:
            rr = min(radio, w / 2, h / 2)
            k = 0.5523 * rr
            x1, y1 = x0 + w, y0 + h
            partes.append(
                f"{x0 + rr:.2f} {y0:.2f} m {x1 - rr:.2f} {y0:.2f} l "
                f"{x1 - rr + k:.2f} {y0:.2f} {x1:.2f} {y0 + rr - k:.2f} {x1:.2f} {y0 + rr:.2f} c "
                f"{x1:.2f} {y1 - rr:.2f} l "
                f"{x1:.2f} {y1 - rr + k:.2f} {x1 - rr + k:.2f} {y1:.2f} {x1 - rr:.2f} {y1:.2f} c "
                f"{x0 + rr:.2f} {y1:.2f} l "
                f"{x0 + rr - k:.2f} {y1:.2f} {x0:.2f} {y1 - rr + k:.2f} {x0:.2f} {y1 - rr:.2f} c "
                f"{x0:.2f} {y0 + rr:.2f} l "
                f"{x0:.2f} {y0 + rr - k:.2f} {x0 + rr - k:.2f} {y0:.2f} {x0 + rr:.2f} {y0:.2f} c h")
        partes.append("B" if relleno and borde else ("f" if relleno else "S"))
        self._op("q " + " ".join(partes) + " Q")

    def linea(self, x1: float, y1: float, x2: float, y2: float, color: str = "#D2D2D7", grosor: float = 0.5,
              discontinua: bool = False) -> None:
        r, g, b = _rgb(color)
        guion = "[2 2] 0 d " if discontinua else ""
        self._op(f"q {r:.3f} {g:.3f} {b:.3f} RG {grosor:.2f} w {guion}{x1:.2f} {self._y(y1):.2f} m "
                 f"{x2:.2f} {self._y(y2):.2f} l S Q")

    def matriz(self, x: float, y: float, modulos: list[list[bool]], tam_modulo: float, color: str = "#000000"):
        """Dibuja una matriz de módulos (p. ej. un código QR)."""
        r, g, b = _rgb(color)
        trozos = [f"q {r:.3f} {g:.3f} {b:.3f} rg"]
        for fi, fila in enumerate(modulos):
            ci = 0
            while ci < len(fila):
                if fila[ci]:
                    ini = ci
                    while ci < len(fila) and fila[ci]:
                        ci += 1
                    trozos.append(f"{x + ini * tam_modulo:.3f} {self._y(y + (fi + 1) * tam_modulo):.3f} "
                                  f"{(ci - ini) * tam_modulo:.3f} {tam_modulo:.3f} re")
                else:
                    ci += 1
        trozos.append("f Q")
        self._op(" ".join(trozos))

    # ------------------------------------------------------------ salida
    def bytes(self) -> bytes:
        if not self.paginas:
            self.nueva_pagina()
        objetos: list[bytes] = []

        def nuevo(contenido: bytes) -> int:
            objetos.append(contenido)
            return len(objetos)

        catalogo = nuevo(b"")  # se rellena al final
        raiz_paginas = nuevo(b"")
        f1 = nuevo(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
        f2 = nuevo(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>")
        hijos = []
        for ops in self.paginas:
            datos = zlib.compress(b"\n".join(ops))
            cont = nuevo(b"<< /Length %d /Filter /FlateDecode >>\nstream\n" % len(datos) + datos + b"\nendstream")
            pag = nuevo((f"<< /Type /Page /Parent {raiz_paginas} 0 R /MediaBox [0 0 {self.ancho:.2f} {self.alto:.2f}] "
                         f"/Resources << /Font << /F1 {f1} 0 R /F2 {f2} 0 R >> >> /Contents {cont} 0 R >>").encode())
            hijos.append(pag)
        objetos[raiz_paginas - 1] = (f"<< /Type /Pages /Count {len(hijos)} /Kids ["
                                     + " ".join(f"{h} 0 R" for h in hijos) + "] >>").encode()
        objetos[catalogo - 1] = f"<< /Type /Catalog /Pages {raiz_paginas} 0 R >>".encode()
        info = nuevo(b"<< /Title (" + _esc(_cod(self.titulo)) + b") /Author (" + _esc(_cod(self.autor))
                     + b") /Producer (InventarioPC) >>")

        salida = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
        posiciones = []
        for i, obj in enumerate(objetos, start=1):
            posiciones.append(len(salida))
            salida += f"{i} 0 obj\n".encode() + obj + b"\nendobj\n"
        xref = len(salida)
        salida += f"xref\n0 {len(objetos) + 1}\n0000000000 65535 f \n".encode()
        for p in posiciones:
            salida += f"{p:010d} 00000 n \n".encode()
        salida += (f"trailer\n<< /Size {len(objetos) + 1} /Root {catalogo} 0 R /Info {info} 0 R >>\n"
                   f"startxref\n{xref}\n%%EOF\n").encode()
        return bytes(salida)

    def guardar(self, ruta: str | Path) -> Path:
        ruta = Path(ruta)
        ruta.write_bytes(self.bytes())
        return ruta
