"""Utilidades compartidas por la ventana principal y los diálogos."""
from __future__ import annotations

import os
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import ttk

from .campos import CLAVES, mostrar_fecha, normalizar
from .tema import C, barra_titulo_oscura

TITULO = "Inventario de Equipos"


def recurso(nombre: str) -> Path:
    """Ruta a un recurso empaquetado (funciona dentro del .exe de PyInstaller)."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base / "assets" / nombre


def abrir_archivo(ruta: Path) -> None:
    try:
        if os.name == "nt":
            os.startfile(str(ruta))  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(ruta)])
        else:
            subprocess.Popen(["xdg-open", str(ruta)])
    except OSError:
        pass


def texto_fila(eq: dict) -> str:
    return " ".join(normalizar(eq.get(c, "")) for c in CLAVES) + " " + " ".join(
        normalizar(mostrar_fecha(eq.get(c, ""))) for c in ("fecha_compra", "fin_garantia"))


def tabla_con_scroll(padre, **kw) -> ttk.Treeview:
    """Treeview con barras de desplazamiento finas dentro de `padre` (rejilla 2x2)."""
    tv = ttk.Treeview(padre, **kw)
    ys = ttk.Scrollbar(padre, orient="vertical", command=tv.yview)
    xs = ttk.Scrollbar(padre, orient="horizontal", command=tv.xview)
    tv.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
    tv.grid(row=0, column=0, sticky="nsew")
    ys.grid(row=0, column=1, sticky="ns", padx=(2, 0))
    xs.grid(row=1, column=0, sticky="ew", pady=(2, 0))
    padre.rowconfigure(0, weight=1)
    padre.columnconfigure(0, weight=1)
    return tv


def tarjeta(padre, **kw) -> tuple[ttk.Frame, ttk.Frame]:
    """Devuelve (tarjeta, interior). El interior tiene el fondo de la tarjeta para colocar contenido."""
    kw.setdefault("padding", 14)
    t = ttk.Frame(padre, style="Tarjeta.TFrame", **kw)
    interior = ttk.Frame(t, style="EnTarjeta.TFrame")
    interior.pack(fill="both", expand=True)
    return t, interior


def foco_en_texto(widget) -> bool:
    """¿El foco está en un campo de texto? (para no interceptar Ctrl+Z, Ctrl+V… mientras se escribe)"""
    return isinstance(widget, (tk.Entry, ttk.Entry, tk.Text, ttk.Combobox, tk.Spinbox, ttk.Spinbox))


def en_segundo_plano(raiz: tk.Misc, funcion, al_terminar, al_fallar) -> None:
    """Ejecuta `funcion()` en un hilo y llama a `al_terminar(resultado)` o `al_fallar(error)` en el hilo
    de la interfaz."""
    estado: dict = {}

    def trabajo():
        try:
            estado["ok"] = funcion()
        except Exception as exc:  # noqa: BLE001
            estado["error"] = exc

    hilo = threading.Thread(target=trabajo, daemon=True)
    hilo.start()

    def comprobar():
        if hilo.is_alive():
            raiz.after(100, comprobar)
        elif "error" in estado:
            al_fallar(estado["error"])
        else:
            al_terminar(estado.get("ok"))
    raiz.after(100, comprobar)


class Dialogo(tk.Toplevel):
    """Ventana de diálogo con título grande, subtítulo y pie de botones."""

    def __init__(self, app, titulo: str, subtitulo: str = ""):
        super().__init__(app)
        self.app = app
        self.title(titulo)
        self.transient(app)
        self.configure(background=C["fondo"])
        self.cuerpo = ttk.Frame(self, padding=(app.tema.px(24), app.tema.px(18)))
        self.cuerpo.pack(fill="both", expand=True)
        cab = ttk.Frame(self.cuerpo)
        cab.pack(fill="x", pady=(0, app.tema.px(12)))
        self.cabecera = cab
        ttk.Label(cab, text=titulo, font=(app.tema.familia, 17, "bold")).pack(anchor="w")
        self.lbl_subtitulo = ttk.Label(cab, text=subtitulo, style="Subtitulo.TLabel")
        if subtitulo:
            self.lbl_subtitulo.pack(anchor="w", pady=(2, 0))
        self.bind("<Escape>", lambda e: self.destroy())
        barra_titulo_oscura(self, getattr(app, "oscuro", False))

    def pie(self, texto_ok: str, comando, texto_info: str = "", texto_cancelar: str = "Cancelar") -> ttk.Frame:
        pie = ttk.Frame(self.cuerpo)
        pie.pack(fill="x", side="bottom", pady=(self.app.tema.px(14), 0))
        if texto_info:
            ttk.Label(pie, text=texto_info, style="Secundario.TLabel").pack(side="left")
        if texto_ok:
            ttk.Button(pie, text=texto_ok, style="Acento.TButton", command=comando).pack(side="right")
        ttk.Button(pie, text=texto_cancelar, command=self.destroy).pack(side="right", padx=(0, 8))
        self.barra_pie = pie
        return pie

    def centrar(self, modal: bool = True):
        self.update_idletasks()
        x = self.app.winfo_rootx() + (self.app.winfo_width() - self.winfo_width()) // 2
        y = self.app.winfo_rooty() + max(0, (self.app.winfo_height() - self.winfo_height()) // 3)
        self.geometry(f"+{max(0, x)}+{max(0, y)}")
        if modal:
            try:
                self.grab_set()
            except tk.TclError:
                pass
