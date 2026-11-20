"""Ventana del servidor central (InventarioPC.exe --servidor)."""
from __future__ import annotations

import os
import sys
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from . import __version__
from .servidor import ServidorInventario, direcciones_locales
from .tema import C, MODOS, Tema, aplicar_modo, barra_titulo_oscura
from .db import leer_config
from .ui_comun import abrir_archivo, recurso, tabla_con_scroll, tarjeta

TITULO = "Servidor de inventario"


def _carpeta_inicio() -> Path | None:
    if os.name != "nt":
        return None
    return Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"


def _archivo_inicio() -> Path | None:
    c = _carpeta_inicio()
    return c / "InventarioPC Servidor.vbs" if c else None


def _orden_servidor() -> list[str]:
    if getattr(sys, "frozen", False):
        return [sys.executable, "--servidor", "--minimizado"]
    return [sys.executable, str(Path(sys.argv[0]).resolve()), "--servidor", "--minimizado"]


class VentanaServidor(tk.Tk):
    def __init__(self, srv: ServidorInventario, minimizado: bool = False):
        super().__init__()
        self.srv = srv
        self.oscuro = aplicar_modo(leer_config().get("tema", MODOS[0]))
        self.tema = Tema(self)
        px = self.tema.px
        self.title(TITULO)
        self.geometry(f"{px(860)}x{px(640)}")
        self.minsize(px(640), px(480))
        try:
            if os.name == "nt" and recurso("icono.ico").exists():
                self.iconbitmap(default=str(recurso("icono.ico")))
        except tk.TclError:
            pass
        cuerpo = ttk.Frame(self, padding=(px(26), px(22)))
        cuerpo.pack(fill="both", expand=True)

        cab = ttk.Frame(cuerpo)
        cab.pack(fill="x")
        izq = ttk.Frame(cab)
        izq.pack(side="left")
        ttk.Label(izq, text="Servidor de inventario", style="Titulo.TLabel").pack(anchor="w")
        fila = ttk.Frame(izq)
        fila.pack(anchor="w")
        tk.Label(fila, text="●", fg=C["verde"], bg=C["fondo"], font=(self.tema.familia, 11)).pack(side="left")
        ttk.Label(fila, text=f"En marcha · versión {__version__} · deja esta ventana abierta (puedes minimizarla)",
                  style="Subtitulo.TLabel").pack(side="left", padx=(px(4), 0))
        ttk.Button(cab, text="Detener servidor", style="Peligro.TButton", command=self._cerrar).pack(side="right")

        # datos de conexión
        ttk.Label(cuerpo, text="PARA CONECTAR LOS DEMÁS EQUIPOS", style="Seccion.TLabel").pack(
            anchor="w", padx=px(4), pady=(px(18), px(6)))
        t, inn = tarjeta(cuerpo)
        t.pack(fill="x")
        inn.columnconfigure(1, weight=1)
        ttk.Label(inn, text="Dirección", style="EnTarjetaCampo.TLabel").grid(row=0, column=0, sticky="w")
        dirs = "   ·   ".join(f"{ip}:{srv.puerto}" for ip in direcciones_locales())
        ttk.Label(inn, text=dirs, style="EnTarjetaTitulo.TLabel").grid(row=0, column=1, sticky="w", padx=px(12))
        ttk.Button(inn, text="Copiar", style="Blanco.Plano.TButton",
                   command=lambda: self._copiar(dirs.split("   ·   ")[0])).grid(row=0, column=2)
        ttk.Label(inn, text="Clave", style="EnTarjetaCampo.TLabel").grid(row=1, column=0, sticky="w", pady=(px(6), 0))
        self.lbl_clave = ttk.Label(inn, text=srv.clave, style="EnTarjetaTitulo.TLabel")
        self.lbl_clave.grid(row=1, column=1, sticky="w", padx=px(12), pady=(px(6), 0))
        b = ttk.Frame(inn, style="EnTarjeta.TFrame")
        b.grid(row=1, column=2, pady=(px(6), 0))
        ttk.Button(b, text="Copiar", style="Blanco.Plano.TButton", command=lambda: self._copiar(srv.clave)).pack(side="left")
        ttk.Button(b, text="Cambiar", style="Blanco.Plano.TButton", command=self._nueva_clave).pack(side="left")
        ttk.Label(inn, text="Red", style="EnTarjetaCampo.TLabel").grid(row=2, column=0, sticky="w", pady=(px(6), 0))
        self.lbl_red = ttk.Label(inn, text="", style="EnTarjeta.TLabel")
        self.lbl_red.grid(row=2, column=1, columnspan=2, sticky="w", padx=px(12), pady=(px(6), 0))
        ttk.Label(inn, text="En cada equipo: Archivo → Conectar a un servidor central, y escribe la dirección y la "
                            "clave. Sólo se atiende a equipos de la misma red local: desde otras redes (Internet, "
                            "VPN…) los equipos trabajan sin conexión y envían lo que creen al volver a esta red. "
                            "La primera vez, Windows puede pedir permiso en el cortafuegos: pulsa «Permitir».",
                  style="EnTarjetaSec.TLabel", wraplength=px(760), justify="left").grid(
            row=3, column=0, columnspan=3, sticky="w", pady=(px(10), 0))

        opciones = ttk.Frame(cuerpo)
        opciones.pack(fill="x", pady=(px(10), 0))
        self.var_inicio = tk.BooleanVar(value=bool(_archivo_inicio() and _archivo_inicio().exists()))
        if _carpeta_inicio():
            ttk.Checkbutton(opciones, text="Iniciar el servidor al encender el equipo", variable=self.var_inicio,
                            command=self._inicio_windows, style="Fondo.TCheckbutton").pack(side="left")
        ttk.Button(opciones, text="Restaurar copia…", style="Plano.TButton", command=self._restaurar).pack(side="right")
        ttk.Button(opciones, text="Carpeta de datos", style="Plano.TButton",
                   command=lambda: abrir_archivo(srv.carpeta)).pack(side="right")

        # clientes y actividad
        rejilla = ttk.Frame(cuerpo)
        rejilla.pack(fill="both", expand=True, pady=(px(14), 0))
        rejilla.columnconfigure(0, weight=1, uniform="r")
        rejilla.columnconfigure(1, weight=1, uniform="r")
        rejilla.rowconfigure(1, weight=1)
        ttk.Label(rejilla, text="EQUIPOS CONECTADOS", style="Seccion.TLabel").grid(row=0, column=0, sticky="w", padx=px(4))
        ttk.Label(rejilla, text="ACTIVIDAD", style="Seccion.TLabel").grid(row=0, column=1, sticky="w", padx=px(16))
        t1, m1 = tarjeta(rejilla, padding=px(6))
        t1.grid(row=1, column=0, sticky="nsew", pady=(px(6), 0))
        self.tv_clientes = tabla_con_scroll(m1, columns=("usuario", "ip", "visto"), show="headings")
        for c, txt, w in (("usuario", "USUARIO", 150), ("ip", "IP", 95), ("visto", "ÚLTIMA VEZ", 80)):
            self.tv_clientes.heading(c, text=txt, anchor="w")
            self.tv_clientes.column(c, width=px(w), stretch=c == "usuario")
        t2, m2 = tarjeta(rejilla, padding=px(6))
        t2.grid(row=1, column=1, sticky="nsew", padx=(px(12), 0), pady=(px(6), 0))
        self.tv_act = tabla_con_scroll(m2, columns=("hora", "quien", "que"), show="headings")
        for c, txt, w in (("hora", "HORA", 64), ("quien", "QUIÉN", 120), ("que", "QUÉ", 220)):
            self.tv_act.heading(c, text=txt, anchor="w")
            self.tv_act.column(c, width=px(w), stretch=c == "que")
        self.lbl_pie = ttk.Label(cuerpo, text="", style="Secundario.TLabel")
        self.lbl_pie.pack(anchor="w", pady=(px(8), 0))

        self.protocol("WM_DELETE_WINDOW", self._cerrar)
        barra_titulo_oscura(self, self.oscuro)
        if minimizado:
            self.after(200, self.iconify)
        self._refrescar()

    def _copiar(self, texto: str):
        self.clipboard_clear()
        self.clipboard_append(texto)
        self.lbl_pie.configure(text=f"Copiado: {texto}")

    def _nueva_clave(self):
        if messagebox.askyesno(TITULO, "Los equipos que ya están conectados tendrán que escribir la clave nueva.\n\n"
                                       "¿Generar una clave nueva?"):
            self.lbl_clave.configure(text=self.srv.nueva_clave())

    def _inicio_windows(self):
        archivo = _archivo_inicio()
        if not archivo:
            return
        try:
            if self.var_inicio.get():
                orden = " ".join(f'""{p}""' if " " in p else p for p in _orden_servidor())
                archivo.write_text(f'CreateObject("WScript.Shell").Run "{orden}", 7, False\r\n', encoding="utf-8")
            elif archivo.exists():
                archivo.unlink()
        except OSError as exc:
            messagebox.showerror(TITULO, f"No se pudo cambiar el inicio automático:\n{exc}")

    def _restaurar(self):
        origen = filedialog.askopenfilename(parent=self, filetypes=[("Base de datos", "*.db"), ("Todos", "*.*")],
                                            initialdir=str(self.srv.carpeta / "copias"))
        if not origen or not messagebox.askyesno(
                TITULO, "Se sustituirán los datos del servidor por los de la copia (antes se guarda una copia del "
                        "estado actual). Los equipos conectados verán el cambio automáticamente.\n\n¿Continuar?",
                icon="warning"):
            return
        try:
            with self.srv.lock:
                self.srv.bd.restaurar(origen)
                for pilas in self.srv.pilas.values():
                    pilas[0].clear()
                    pilas[1].clear()
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(TITULO, f"No es una copia válida:\n{exc}")
            return
        self.srv.anotar("servidor", f"Copia restaurada: {Path(origen).name}")
        messagebox.showinfo(TITULO, "Copia restaurada.")

    def _refrescar(self):
        ahora = time.time()
        self.tv_clientes.delete(*self.tv_clientes.get_children())
        for cid, info in sorted(self.srv.clientes.items(), key=lambda x: -x[1]["ultima"]):
            seg = int(ahora - info["ultima"])
            visto = "ahora" if seg < 15 else (f"hace {seg // 60} min" if seg < 3600 else
                                             datetime.fromtimestamp(info["ultima"]).strftime("%d/%m %H:%M"))
            self.tv_clientes.insert("", "end", values=(info.get("usuario") or "?", info["ip"], visto))
        self.tv_act.delete(*self.tv_act.get_children())
        for hora, quien, que in list(self.srv.actividad)[:200]:
            self.tv_act.insert("", "end", values=(hora, quien, que))
        self.lbl_red.configure(text=f"Sólo equipos de la red {self.srv.redes_permitidas()}")
        with self.srv.lock:
            n = self.srv.bd.contar()
        self.lbl_pie.configure(text=f"{n} equipos en el inventario · datos en {self.srv.carpeta}")
        self.after(2000, self._refrescar)

    def _cerrar(self):
        activos = sum(1 for i in self.srv.clientes.values() if time.time() - i["ultima"] < 60)
        aviso = (f"Hay {activos} equipo(s) conectados. " if activos else "")
        if not messagebox.askyesno(TITULO, aviso + "Si detienes el servidor, los demás equipos pasarán a trabajar "
                                                   "sin conexión (sólo consultar y crear equipos nuevos).\n\n"
                                                   "¿Detener el servidor?", icon="warning"):
            return
        self.srv.detener()
        self.destroy()


def main(argv: list[str]) -> int:
    puerto = None
    if "--puerto" in argv:
        puerto = int(argv[argv.index("--puerto") + 1])
    carpeta = None
    if "--datos" in argv:
        carpeta = argv[argv.index("--datos") + 1]
    try:
        srv = ServidorInventario(carpeta, puerto=puerto)
        srv.iniciar()
    except OSError as exc:
        if "--sin-ventana" in argv:
            print(f"No se pudo iniciar el servidor: {exc}")
            return 1
        r = tk.Tk()
        r.withdraw()
        messagebox.showerror(TITULO, f"No se pudo iniciar el servidor (¿ya hay uno en marcha en este equipo?):\n\n{exc}")
        r.destroy()
        return 1
    if "--sin-ventana" in argv:
        print(f"Servidor en marcha en {', '.join(direcciones_locales())}:{srv.puerto} · clave {srv.clave}")
        try:
            while True:
                time.sleep(3600)
        except KeyboardInterrupt:
            srv.detener()
        return 0
    VentanaServidor(srv, minimizado="--minimizado" in argv).mainloop()
    return 0
