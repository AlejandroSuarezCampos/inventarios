"""Interfaz gráfica (Tkinter) del inventario de equipos."""
from __future__ import annotations

import os
import subprocess
import sys
import tkinter as tk
from collections import Counter
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from . import __version__
from .campos import (CAMPOS, CAMPOS_POR_CLAVE, CLAVES, DIAS_AVISO_GARANTIA, ESTADOS,
                     estado_garantia, fecha_a_date, mostrar_fecha, normalizar, parsear_fecha)
from .db import BaseDatos, guardar_config, leer_config
from .excel_io import exportar_csv, exportar_excel, exportar_plantilla, leer_excel
from .tema import C, PALETA_GRAFICOS, CampoBusqueda, GraficoBarras, Tema

TITULO = "Inventario de Equipos"
TODOS = "Garantía"
# Primer valor de cada filtro = «sin filtrar»
TODOS_FILTRO = {"delegacion": "Delegaciones", "marca": "Marcas", "os": "Sistemas", "estado": "Estados"}
FILTROS_GARANTIA = [TODOS, "Vigente", "Caduca pronto", "Caducada", "Sin fecha"]
_MAPA_GARANTIA = {FILTROS_GARANTIA[1]: "vigente", FILTROS_GARANTIA[2]: "proxima",
                  FILTROS_GARANTIA[3]: "caducada", FILTROS_GARANTIA[4]: "sin_dato"}


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
    """Devuelve (tarjeta, interior). El interior tiene fondo blanco para colocar contenido."""
    kw.setdefault("padding", 14)
    t = ttk.Frame(padre, style="Tarjeta.TFrame", **kw)
    interior = ttk.Frame(t, style="EnTarjeta.TFrame")
    interior.pack(fill="both", expand=True)
    return t, interior


# =====================================================================
#  Base de los diálogos
# =====================================================================
class Dialogo(tk.Toplevel):
    def __init__(self, app: "App", titulo: str, subtitulo: str = ""):
        super().__init__(app)
        self.app = app
        self.title(titulo)
        self.transient(app)
        self.configure(background=C["fondo"])
        self.cuerpo = ttk.Frame(self, padding=(app.tema.px(24), app.tema.px(18)))
        self.cuerpo.pack(fill="both", expand=True)
        cab = ttk.Frame(self.cuerpo)
        cab.pack(fill="x", pady=(0, app.tema.px(14)))
        ttk.Label(cab, text=titulo, font=(app.tema.familia, 17, "bold")).pack(anchor="w")
        if subtitulo:
            ttk.Label(cab, text=subtitulo, style="Subtitulo.TLabel").pack(anchor="w", pady=(2, 0))
        self.bind("<Escape>", lambda e: self.destroy())

    def pie(self, texto_ok: str, comando, texto_info: str = "") -> ttk.Frame:
        pie = ttk.Frame(self.cuerpo)
        pie.pack(fill="x", side="bottom", pady=(self.app.tema.px(16), 0))
        if texto_info:
            ttk.Label(pie, text=texto_info, style="Secundario.TLabel").pack(side="left")
        ttk.Button(pie, text=texto_ok, style="Acento.TButton", command=comando).pack(side="right")
        ttk.Button(pie, text="Cancelar", command=self.destroy).pack(side="right", padx=(0, 8))
        return pie

    def centrar(self):
        self.update_idletasks()
        x = self.app.winfo_rootx() + (self.app.winfo_width() - self.winfo_width()) // 2
        y = self.app.winfo_rooty() + max(0, (self.app.winfo_height() - self.winfo_height()) // 3)
        self.geometry(f"+{max(0, x)}+{max(0, y)}")
        self.grab_set()


# =====================================================================
#  Formulario de alta / edición
# =====================================================================
class FormularioEquipo(Dialogo):
    def __init__(self, app: "App", equipo: dict | None = None, titulo: str = "Nuevo equipo"):
        sub = ""
        if equipo and equipo.get("id"):
            sub = " · ".join(x for x in (equipo.get("persona"), equipo.get("delegacion"), equipo.get("num_serie")) if x)
        super().__init__(app, titulo, sub or "Rellena los datos del equipo y de la persona que lo usa.")
        self.equipo = equipo
        self.resultado: dict | None = None
        self.resizable(True, False)
        self.widgets: dict[str, tk.Widget] = {}
        self.vars: dict[str, tk.StringVar] = {}
        px = app.tema.px

        info = ""
        if equipo and equipo.get("id"):
            info = f"ID {equipo['id']} · creado {equipo.get('creado', '')[:16]} · modificado {equipo.get('modificado', '')[:16]}"
        self.pie("Guardar", self._guardar, info)

        rejilla = ttk.Frame(self.cuerpo)
        rejilla.pack(fill="both", expand=True)
        rejilla.columnconfigure(0, weight=1, uniform="c")
        rejilla.columnconfigure(1, weight=2, uniform="c")

        izq = ttk.Frame(rejilla)
        izq.grid(row=0, column=0, sticky="nsew", padx=(0, px(16)))
        der = ttk.Frame(rejilla)
        der.grid(row=0, column=1, sticky="nsew")

        secciones = {}
        for padre, clave, titulo_sec, expandir in ((izq, "propietario", "PROPIETARIO DEL ORDENADOR", False),
                                                   (izq, "otros", "ESTADO Y OBSERVACIONES", True),
                                                   (der, "equipo", "DATOS DEL ORDENADOR", True)):
            ttk.Label(padre, text=titulo_sec, style="Seccion.TLabel").pack(anchor="w", padx=px(4), pady=(0, px(6)))
            t, interior = tarjeta(padre)
            t.pack(fill="both", expand=expandir, pady=(0, px(12)))
            secciones[clave] = interior

        contadores = {"propietario": 0, "equipo": 0, "otros": 0}
        for campo in CAMPOS:
            marco = secciones[campo.grupo]
            idx = contadores[campo.grupo]
            contadores[campo.grupo] += 1
            if campo.grupo == "equipo":
                r, c = idx // 2, idx % 2
            else:
                r, c = idx, 0
            marco.columnconfigure(c, weight=1, uniform="campos")
            celda = ttk.Frame(marco, style="EnTarjeta.TFrame")
            celda.grid(row=r, column=c, sticky="ew", padx=(0 if c == 0 else px(12), 0), pady=(0, px(7)))
            ttk.Label(celda, text=campo.etiqueta, style="EnTarjetaCampo.TLabel").pack(anchor="w", pady=(0, px(2)))
            valor = (equipo or {}).get(campo.clave, "") or ""
            if campo.tipo == "largo":
                w = tk.Text(celda, width=30, height=3, wrap="word", font=app.tema.fuente, relief="flat",
                            bd=0, padx=px(8), pady=px(6), highlightthickness=1,
                            highlightbackground=C["borde"], highlightcolor=C["acento"],
                            background=C["tarjeta"], foreground=C["texto"], insertbackground=C["texto"])
                w.insert("1.0", valor)
                w.pack(fill="both", expand=True)
            elif campo.tipo == "fecha":
                var = tk.StringVar(value=mostrar_fecha(valor))
                fila = ttk.Frame(celda, style="EnTarjeta.TFrame")
                fila.pack(fill="x")
                w = ttk.Entry(fila, textvariable=var, width=12, style="Blanco.TEntry")
                w.pack(side="left", fill="x", expand=True)
                ttk.Button(fila, text="Hoy", style="Blanco.Plano.TButton",
                           command=lambda v=var: v.set(date.today().strftime("%d/%m/%Y"))).pack(side="left", padx=(4, 0))
                if campo.clave == "fin_garantia":
                    mb = ttk.Menubutton(fila, text="+ años", style="Blanco.Plano.TButton")
                    menu = tk.Menu(mb, tearoff=False)
                    for anios in (1, 2, 3, 4, 5):
                        menu.add_command(label=f"Fecha de compra + {anios} año{'s' if anios > 1 else ''}",
                                         command=lambda a=anios: self._garantia_desde_compra(a))
                    mb["menu"] = menu
                    mb.pack(side="left")
                self.vars[campo.clave] = var
            elif campo.clave == "estado":
                var = tk.StringVar(value=valor or "Activo")
                w = ttk.Combobox(celda, textvariable=var, values=ESTADOS, width=24, style="Blanco.TCombobox")
                w.pack(fill="x")
                self.vars[campo.clave] = var
            else:
                var = tk.StringVar(value=valor)
                w = ttk.Combobox(celda, textvariable=var, width=24, values=app.bd.valores_distintos(campo.clave),
                                 style="Blanco.TCombobox")
                w.pack(fill="x")
                self.vars[campo.clave] = var
            self.widgets[campo.clave] = w

        self.bind("<Return>", lambda e: self._guardar() if not isinstance(e.widget, tk.Text) else None)
        self.widgets["persona"].focus_set()
        self.centrar()

    def _garantia_desde_compra(self, anios: int):
        d = fecha_a_date(parsear_fecha(self.vars["fecha_compra"].get()))
        if not d:
            messagebox.showwarning(TITULO, "Introduce primero una fecha de compra válida.", parent=self)
            return
        try:
            fin = d.replace(year=d.year + anios)
        except ValueError:  # 29 de febrero
            fin = d.replace(year=d.year + anios, day=28)
        self.vars["fin_garantia"].set(fin.strftime("%d/%m/%Y"))

    def datos(self) -> dict:
        d = {}
        for campo in CAMPOS:
            if campo.tipo == "largo":
                d[campo.clave] = self.widgets[campo.clave].get("1.0", "end").strip()
            else:
                d[campo.clave] = self.vars[campo.clave].get().strip()
        return d

    def _guardar(self):
        d = self.datos()
        for clave in ("fecha_compra", "fin_garantia"):
            if d[clave]:
                iso = parsear_fecha(d[clave])
                if not fecha_a_date(iso) and not messagebox.askyesno(
                        TITULO, f"«{d[clave]}» no es una fecha reconocible (formato DD/MM/AAAA) en "
                                f"«{CAMPOS_POR_CLAVE[clave].etiqueta}».\n\n¿Guardarlo igualmente como texto?",
                        parent=self):
                    self.widgets[clave].focus_set()
                    return
                d[clave] = iso
        compra, fin = fecha_a_date(d["fecha_compra"]), fecha_a_date(d["fin_garantia"])
        if compra and fin and fin < compra:
            if not messagebox.askyesno(TITULO, "La expiración de garantía es anterior a la fecha de compra.\n"
                                               "¿Guardar de todos modos?", parent=self):
                return
        if not any(v for k, v in d.items() if k != "estado"):
            messagebox.showwarning(TITULO, "El equipo no tiene ningún dato.", parent=self)
            return
        self.resultado = d
        self.destroy()


# =====================================================================
#  Diálogo de importación
# =====================================================================
class _Cancelado(Exception):
    pass


class DialogoImportar(Dialogo):
    """Importa las filas del Excel. Ningún campo se usa como clave: cada fila es un equipo."""
    OPCIONES = [
        ("omitir_identicas", "Añadir las filas, saltando las que sean idénticas (en todos los campos) "
                             "a un equipo ya guardado"),
        ("anadir", "Añadir todas las filas como equipos nuevos"),
        ("reemplazar", "Reemplazar todo el inventario actual por el contenido del Excel"),
    ]

    def __init__(self, app: "App", ruta: Path):
        try:
            res = leer_excel(ruta)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(TITULO, f"No se pudo leer el archivo:\n{ruta}\n\n{exc}", parent=app)
            raise _Cancelado from exc
        if not res.filas:
            avisos = "\n".join(res.avisos) or "El archivo no contiene filas con datos."
            messagebox.showwarning(TITULO, f"No hay equipos que importar.\n\n{avisos}", parent=app)
            raise _Cancelado

        super().__init__(app, "Importar inventario", f"{ruta.name} · hojas: {', '.join(res.hojas_leidas)}")
        self.res = res
        self.ruta = ruta
        self.importado = False
        px = app.tema.px
        self.geometry(f"{px(1040)}x{px(680)}")
        self.minsize(px(720), px(480))

        # filas idénticas (todos los campos iguales) a equipos guardados o repetidas en el archivo
        guardadas = app.bd.firmas()
        self.identicas: set[int] = set()
        self.repetidas: set[int] = set()
        vistas: set[tuple] = set()
        for i, fila in enumerate(res.filas):
            f = app.bd.firma(fila)
            if f in guardadas:
                self.identicas.add(i)
            elif f in vistas:
                self.repetidas.add(i)
            vistas.add(f)

        self.pie(f"Importar", self._importar)

        cifras = ttk.Frame(self.cuerpo)
        cifras.pack(fill="x")
        for i, (n, texto, estilo) in enumerate((
                (len(res.filas), "Filas con datos", "CifraAzul.TLabel"),
                (app.bd.contar(), "Equipos guardados ahora", "CifraGris.TLabel"),
                (len(self.identicas), "Idénticas a equipos guardados", "CifraAmbar.TLabel"))):
            t, inn = tarjeta(cifras)
            t.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else px(12), 0))
            cifras.columnconfigure(i, weight=1, uniform="cif")
            ttk.Label(inn, text=str(n), style=estilo).pack(anchor="w")
            ttk.Label(inn, text=texto, style="EnTarjetaSec.TLabel").pack(anchor="w")

        notas = []
        for hoja, cols in res.columnas_reconocidas.items():
            notas.append(("✓", C["verde"], f"Columnas reconocidas en «{hoja}»: {', '.join(cols)}"))
        for hoja, cols in res.columnas_ignoradas.items():
            notas.append(("!", C["ambar"], f"Columnas no reconocidas (se ignoran) en «{hoja}»: {', '.join(cols)}"))
        if self.repetidas:
            notas.append(("!", C["ambar"], f"{len(self.repetidas)} fila(s) idénticas a otra fila del propio archivo."))
        for aviso in res.avisos:
            notas.append(("!", C["ambar"], aviso))
        if notas:
            caja = ttk.Frame(self.cuerpo)
            caja.pack(fill="x", pady=(px(12), 0))
            for icono, color, texto in notas:
                f = ttk.Frame(caja)
                f.pack(fill="x", pady=1)
                tk.Label(f, text=icono, fg=color, bg=C["fondo"], font=app.tema.fuente_bold, width=2).pack(side="left")
                ttk.Label(f, text=texto, style="Secundario.TLabel", wraplength=px(940), justify="left").pack(side="left")

        ttk.Label(self.cuerpo, text="CÓMO IMPORTAR", style="Seccion.TLabel").pack(
            anchor="w", padx=px(4), pady=(px(14), px(6)))
        t, op = tarjeta(self.cuerpo)
        t.configure(padding=(px(14), px(8)))
        t.pack(fill="x")
        self.modo = tk.StringVar(value="omitir_identicas" if (self.identicas or self.repetidas) else "anadir")
        for valor, texto in self.OPCIONES:
            ttk.Radiobutton(op, text=texto, value=valor, variable=self.modo).pack(anchor="w")

        ttk.Label(self.cuerpo, text="VISTA PREVIA", style="Seccion.TLabel").pack(anchor="w", padx=px(4),
                                                                               pady=(px(14), px(6)))
        t, marco = tarjeta(self.cuerpo)
        t.configure(padding=px(6))
        t.pack(fill="both", expand=True)
        cols = ["_estado"] + CLAVES
        tv = tabla_con_scroll(marco, columns=cols, show="headings", height=8)
        tv.heading("_estado", text="FILA")
        tv.column("_estado", width=px(110), stretch=False)
        for c in CAMPOS:
            tv.heading(c.clave, text=c.etiqueta.upper())
            tv.column(c.clave, width=px(c.ancho), stretch=False)
        tv.tag_configure("identica", background=C["ambar_suave"])
        for i, fila in enumerate(res.filas[:1000]):
            if i in self.identicas:
                estado, tags = "● Ya guardada", ("identica",)
            elif i in self.repetidas:
                estado, tags = "● Repetida", ("identica",)
            else:
                estado, tags = "● Nueva", ()
            valores = [estado] + [mostrar_fecha(fila[c.clave]) if c.tipo == "fecha" else fila[c.clave]
                                  for c in CAMPOS]
            tv.insert("", "end", values=valores, tags=tags)
        self.centrar()

    def _importar(self):
        bd, modo = self.app.bd, self.modo.get()
        if modo == "reemplazar" and bd.contar() and not messagebox.askyesno(
                TITULO, f"Se eliminarán los {bd.contar()} equipos actuales y se sustituirán por las "
                        f"{len(self.res.filas)} filas del Excel.\n(Antes se guarda una copia de seguridad.)\n\n"
                        "¿Continuar?", icon="warning", parent=self):
            return
        nuevos = omitidos = 0
        try:
            if modo == "reemplazar":
                bd.copia_seguridad(bd.ruta.parent / "copias" /
                                   f"antes_de_importar_{datetime.now():%Y%m%d_%H%M%S}.db")
                bd.vaciar(f"Reemplazado por {self.ruta.name}")
            vistas = bd.firmas() if modo == "omitir_identicas" else {}
            for fila in self.res.filas:
                datos = {k: v for k, v in fila.items() if not k.startswith("_")}
                if modo == "omitir_identicas":
                    f = bd.firma(datos)
                    if f in vistas:
                        omitidos += 1
                        continue
                    vistas[f] = 0
                bd.insertar(datos, commit=False)
                nuevos += 1
            bd.registrar(None, self.ruta.name, "Importación",
                         f"{nuevos} añadidos, {omitidos} omitidos por ser idénticos"
                         + (" (inventario reemplazado)" if modo == "reemplazar" else ""))
            bd.commit()
        except Exception as exc:  # noqa: BLE001
            bd.con.rollback()
            messagebox.showerror(TITULO, f"Error durante la importación. No se ha guardado nada.\n\n{exc}", parent=self)
            return
        self.importado = True
        self.destroy()
        texto = f"Importación completada.\n\nEquipos añadidos: {nuevos}"
        if omitidos:
            texto += f"\nFilas idénticas omitidas: {omitidos}"
        messagebox.showinfo(TITULO, texto, parent=self.app)


# =====================================================================
#  Ventana principal
# =====================================================================
class App(tk.Tk):
    PAGINAS = (("Inventario", "▦"), ("Resumen", "◔"), ("Historial", "↺"))

    def __init__(self, ruta_bd: str | None = None):
        super().__init__()
        self.title(TITULO)
        self.tema = Tema(self)
        px = self.tema.px
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        ancho, alto = min(px(1440), int(sw * 0.94)), min(px(860), int(sh * 0.88))
        self.geometry(f"{ancho}x{alto}+{(sw - ancho) // 2}+{max(0, (sh - alto) // 3)}")
        self.minsize(min(px(1000), sw - 40), min(px(600), sh - 80))
        if os.name == "nt" and (sw < px(1500) or sh < px(900)):
            try:
                self.state("zoomed")
            except tk.TclError:
                pass
        self._icono()

        try:
            self.bd = BaseDatos(ruta_bd)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(TITULO, f"No se pudo abrir la base de datos:\n{exc}")
            raise
        self.bd.copia_automatica()

        self.equipos: list[dict] = []
        self.visibles: list[dict] = []
        self.orden = ("id", False)
        cfg = leer_config()
        self.columnas_ocultas: set[str] = set(cfg.get("columnas_ocultas", []))
        self.pagina_actual = 0
        self.estado = tk.StringVar()
        self._editor = None
        self._fila_contexto: str | None = None
        self._clave_contexto: str | None = None

        self._menu()
        self._lateral()
        self.contenido = ttk.Frame(self)
        self.contenido.pack(side="left", fill="both", expand=True)
        self.paginas: list[ttk.Frame] = []
        self._pagina_inventario()
        self._pagina_resumen()
        self._pagina_historial()

        self._atajos()
        self.mostrar(0)
        self.recargar()
        self.protocol("WM_DELETE_WINDOW", self._salir)
        self.after(400, self._aviso_garantias)

    # ------------------------------------------------------------ apariencia
    def _icono(self):
        try:
            ico = recurso("icono.ico")
            png = recurso("icono.png")
            if os.name == "nt" and ico.exists():
                self.iconbitmap(default=str(ico))
            if png.exists():
                self._img_icono = tk.PhotoImage(file=str(png))
                if os.name != "nt":
                    self.iconphoto(True, self._img_icono)
                factor = max(1, round(256 / self.tema.px(30)))
                self._img_logo = self._img_icono.subsample(factor, factor)
        except tk.TclError:
            pass

    def _lateral(self):
        px = self.tema.px
        lat = ttk.Frame(self, style="Lateral.TFrame", padding=(px(12), px(18), px(12), px(14)), width=px(210))
        lat.pack(side="left", fill="y")
        lat.pack_propagate(False)
        sep = tk.Frame(self, width=1, bg=C["borde"])
        sep.pack(side="left", fill="y")

        cab = ttk.Frame(lat, style="Lateral.TFrame")
        cab.pack(fill="x", padx=px(6), pady=(0, px(22)))
        if getattr(self, "_img_logo", None):
            tk.Label(cab, image=self._img_logo, bg=C["lateral"], bd=0).pack(side="left", padx=(0, px(10)))
        tx = ttk.Frame(cab, style="Lateral.TFrame")
        tx.pack(side="left")
        ttk.Label(tx, text="Inventario", style="LateralTitulo.TLabel").pack(anchor="w")
        ttk.Label(tx, text="Equipos informáticos", style="Lateral.TLabel").pack(anchor="w")

        ttk.Label(lat, text="GENERAL", style="LateralSeccion.TLabel").pack(anchor="w", padx=px(10), pady=(0, px(4)))
        self.botones_nav: list[ttk.Button] = []
        for i, (nombre, icono) in enumerate(self.PAGINAS):
            b = ttk.Button(lat, text=f"{icono}   {nombre}", style="Lateral.TButton", command=lambda i=i: self.mostrar(i))
            b.pack(fill="x", pady=1)
            self.botones_nav.append(b)

        ttk.Label(lat, text="DATOS", style="LateralSeccion.TLabel").pack(anchor="w", padx=px(10), pady=(px(18), px(4)))
        for texto, cmd in (("⤓   Importar Excel", self.importar),
                           ("⤒   Exportar Excel", lambda: self.exportar("excel", "vista")),
                           ("⧉   Copia de seguridad", self.copia)):
            ttk.Button(lat, text=texto, style="Lateral.TButton", command=cmd).pack(fill="x", pady=1)

        pie = ttk.Frame(lat, style="Lateral.TFrame")
        pie.pack(side="bottom", fill="x", padx=px(10))
        self.lbl_lateral_total = ttk.Label(pie, text="", style="Lateral.TLabel")
        self.lbl_lateral_total.pack(anchor="w")
        self.lbl_lateral_bd = ttk.Label(pie, text="", style="Lateral.TLabel", wraplength=px(170))
        self.lbl_lateral_bd.pack(anchor="w", pady=(2, 0))
        ttk.Label(pie, text=f"Versión {__version__}", style="Lateral.TLabel").pack(anchor="w", pady=(px(6), 0))

    def mostrar(self, indice: int):
        self.pagina_actual = indice
        for i, (p, b) in enumerate(zip(self.paginas, self.botones_nav)):
            if i == indice:
                p.pack(fill="both", expand=True)
                b.state(["selected"])
            else:
                p.pack_forget()
                b.state(["!selected"])
        self._cambio_pagina()

    def _cabecera(self, padre, titulo: str) -> tuple[ttk.Frame, tk.StringVar]:
        px = self.tema.px
        cab = ttk.Frame(padre)
        cab.pack(fill="x", pady=(0, px(16)))
        izq = ttk.Frame(cab)
        izq.pack(side="left")
        ttk.Label(izq, text=titulo, style="Titulo.TLabel").pack(anchor="w")
        var = tk.StringVar()
        ttk.Label(izq, textvariable=var, style="Subtitulo.TLabel").pack(anchor="w")
        acciones = ttk.Frame(cab)
        acciones.pack(side="right", anchor="s")
        return acciones, var

    # ------------------------------------------------------------ menú
    def _menu(self):
        barra = tk.Menu(self)
        m = tk.Menu(barra, tearoff=False)
        m.add_command(label="Importar Excel / CSV…", accelerator="Ctrl+I", command=self.importar)
        m.add_separator()
        m.add_command(label="Exportar a Excel (vista actual)…", accelerator="Ctrl+E",
                      command=lambda: self.exportar("excel", "vista"))
        m.add_command(label="Exportar a Excel (todo)…", command=lambda: self.exportar("excel", "todo"))
        m.add_command(label="Exportar a Excel (seleccionados)…", command=lambda: self.exportar("excel", "seleccion"))
        m.add_command(label="Exportar a CSV (vista actual)…", command=lambda: self.exportar("csv", "vista"))
        m.add_command(label="Guardar plantilla Excel vacía…", command=self.plantilla)
        m.add_separator()
        m.add_command(label="Crear copia de seguridad…", command=self.copia)
        m.add_command(label="Restaurar copia de seguridad…", command=self.restaurar)
        m.add_command(label="Usar otra base de datos (p. ej. en red)…", command=self.cambiar_bd)
        m.add_command(label="Abrir carpeta de datos", command=lambda: abrir_archivo(self.bd.ruta.parent))
        m.add_separator()
        m.add_command(label="Salir", command=self._salir)
        barra.add_cascade(label="Archivo", menu=m)

        e = tk.Menu(barra, tearoff=False)
        e.add_command(label="Nuevo equipo", accelerator="Ctrl+N", command=self.nuevo)
        e.add_command(label="Editar", accelerator="Intro", command=self.editar)
        e.add_command(label="Duplicar", accelerator="Ctrl+D", command=self.duplicar)
        e.add_command(label="Eliminar", accelerator="Supr", command=self.eliminar)
        e.add_separator()
        self.menu_estado = tk.Menu(e, tearoff=False)
        for est in ESTADOS:
            self.menu_estado.add_command(label=est, command=lambda s=est: self.cambiar_campo("estado", s))
        e.add_cascade(label="Cambiar estado de los seleccionados", menu=self.menu_estado)
        e.add_command(label="Editar un campo en los seleccionados…", command=self.cambiar_campo_dialogo)
        e.add_command(label="Editar celda", accelerator="F2", command=self._f2)
        e.add_separator()
        e.add_command(label="Buscar", accelerator="Ctrl+F", command=self._enfocar_busqueda)
        e.add_command(label="Seleccionar todo", accelerator="Ctrl+A", command=self.seleccionar_todo)
        barra.add_cascade(label="Equipos", menu=e)

        v = tk.Menu(barra, tearoff=False)
        for i, (nombre, _) in enumerate(self.PAGINAS):
            v.add_command(label=nombre, accelerator=f"Ctrl+{i + 1}", command=lambda i=i: self.mostrar(i))
        v.add_separator()
        v.add_command(label="Columnas visibles…", command=self.dialogo_columnas)
        v.add_command(label="Actualizar", accelerator="F5", command=self.recargar)
        barra.add_cascade(label="Ver", menu=v)

        a = tk.Menu(barra, tearoff=False)
        a.add_command(label="Formato del Excel de importación", command=self.ayuda_formato)
        a.add_command(label="Acerca de…", command=self.acerca)
        barra.add_cascade(label="Ayuda", menu=a)
        self.config(menu=barra)

        self.menu_contexto = tk.Menu(self, tearoff=False)
        self.menu_contexto.add_command(label="Editar", command=self.editar)
        self.menu_contexto.add_command(label="Duplicar", command=self.duplicar)
        self.menu_contexto.add_command(label="Editar esta celda", command=lambda: self._editar_celda_contexto())
        self.menu_contexto.add_command(label="Editar un campo en la selección…",
                                       command=lambda: self.cambiar_campo_dialogo(self._clave_contexto))
        self.menu_contexto.add_cascade(label="Cambiar estado", menu=self.menu_estado)
        self.menu_contexto.add_command(label="Copiar fila", command=self.copiar_fila)
        self.menu_contexto.add_separator()
        self.menu_contexto.add_command(label="Exportar seleccionados…", command=lambda: self.exportar("excel", "seleccion"))
        self.menu_contexto.add_command(label="Eliminar", command=self.eliminar)

    def _enfocar_busqueda(self):
        self.mostrar(0)
        self.busqueda.focus_set()
        return "break"

    def _atajos(self):
        self.bind_all("<Control-n>", lambda e: self.nuevo())
        self.bind_all("<Control-i>", lambda e: self.importar())
        self.bind_all("<Control-e>", lambda e: self.exportar("excel", "vista"))
        self.bind_all("<Control-f>", lambda e: self._enfocar_busqueda())
        for i in range(len(self.PAGINAS)):
            self.bind_all(f"<Control-Key-{i + 1}>", lambda e, i=i: self.mostrar(i))
        self.bind_all("<F5>", lambda e: self.recargar())
        self.tabla.bind("<Control-d>", lambda e: self.duplicar())
        self.tabla.bind("<Control-a>", lambda e: (self.seleccionar_todo(), "break")[-1])
        self.tabla.bind("<Control-c>", lambda e: self.copiar_fila())
        self.tabla.bind("<Delete>", lambda e: self.eliminar())
        self.tabla.bind("<Return>", lambda e: self.editar())
        self.tabla.bind("<Double-1>", self._doble_click)
        self.tabla.bind("<F2>", self._f2)
        self.tabla.bind("<MouseWheel>", lambda e: self._cerrar_editor(True), add="+")
        self.tabla.bind("<Button-4>", lambda e: self._cerrar_editor(True), add="+")
        self.tabla.bind("<Button-5>", lambda e: self._cerrar_editor(True), add="+")
        self.tabla.bind("<Button-3>", self._click_derecho)
        self.tabla.bind("<<TreeviewSelect>>", lambda e: self._actualizar_estado())

    # ------------------------------------------------------------ página inventario
    def _pagina_inventario(self):
        px = self.tema.px
        p = ttk.Frame(self.contenido, padding=(px(28), px(22), px(28), px(16)))
        self.paginas.append(p)

        acciones, self.sub_inventario = self._cabecera(p, "Inventario")
        ttk.Button(acciones, text="＋  Nuevo equipo", style="Acento.TButton", command=self.nuevo).pack(side="right")
        ttk.Button(acciones, text="Exportar", command=lambda: self.exportar("excel", "vista")).pack(side="right", padx=(0, 8))
        ttk.Button(acciones, text="Importar", command=self.importar).pack(side="right", padx=(0, 8))

        filtros = ttk.Frame(p)
        filtros.pack(fill="x", pady=(0, px(12)))
        self.var_buscar = tk.StringVar()
        self.busqueda = CampoBusqueda(filtros, self.tema, self.var_buscar,
                                      "Buscar persona, equipo, nº de serie…", ancho=16)
        self.var_buscar.trace_add("write", lambda *a: self._programar_filtro())
        self.busqueda.bind_entrada("<Escape>", lambda e: self.var_buscar.set(""))

        self.filtros: dict[str, tk.StringVar] = {}
        self.combos: dict[str, ttk.Combobox] = {}
        # los filtros se colocan desde la derecha; el buscador ocupa el espacio que sobre
        self.btn_limpiar = ttk.Button(filtros, text="Limpiar", style="Plano.TButton", command=self.limpiar_filtros)
        self.btn_limpiar.pack(side="right")
        self.var_garantia = tk.StringVar(value=TODOS)
        cb = ttk.Combobox(filtros, textvariable=self.var_garantia, values=FILTROS_GARANTIA, state="readonly", width=11)
        cb.pack(side="right", padx=(px(8), px(4)))
        cb.bind("<<ComboboxSelected>>", lambda e: (self.aplicar_filtro(), self.tabla.focus_set()))
        for clave, todos in reversed(list(TODOS_FILTRO.items())):
            var = tk.StringVar(value=todos)
            cb = ttk.Combobox(filtros, textvariable=var, state="readonly", width=10)
            cb.pack(side="right", padx=(px(8), 0))
            cb.bind("<<ComboboxSelected>>", lambda e: (self.aplicar_filtro(), self.tabla.focus_set()))
            self.filtros[clave], self.combos[clave] = var, cb
        self.busqueda.pack(side="left", fill="x", expand=True)

        # pie (se empaqueta antes que la tabla para que siempre sea visible)
        pie = ttk.Frame(p)
        pie.pack(fill="x", side="bottom", pady=(px(10), 0))
        for color, texto in ((C["rojo"], "Garantía caducada"),
                             (C["ambar"], f"Caduca en ≤ {DIAS_AVISO_GARANTIA} días"),
                             (C["terciario"], "Baja")):
            tk.Label(pie, text="●", fg=color, bg=C["fondo"], font=(self.tema.familia, 11)).pack(side="left")
            ttk.Label(pie, text=texto, style="Secundario.TLabel").pack(side="left", padx=(2, px(14)))
        ttk.Label(pie, textvariable=self.estado, style="Secundario.TLabel").pack(side="right")

        barra_sel = ttk.Frame(p)
        barra_sel.pack(fill="x", side="bottom", pady=(px(8), 0))
        self.botones_seleccion = []
        for texto, cmd, estilo in (("Editar", self.editar, "Plano.TButton"),
                                   ("Duplicar", self.duplicar, "Plano.TButton"),
                                   ("Copiar", self.copiar_fila, "Plano.TButton"),
                                   ("Eliminar", self.eliminar, "Rojo.Plano.TButton")):
            b = ttk.Button(barra_sel, text=texto, style=estilo, command=cmd)
            b.pack(side="left", padx=(0, px(4)))
            self.botones_seleccion.append(b)
        ttk.Label(barra_sel, text="Doble clic en una celda para editarla · Intro abre la ficha completa · F2 edita",
                  style="Secundario.TLabel").pack(side="right")

        t, marco = tarjeta(p)
        t.configure(padding=(px(6), px(4), px(6), px(6)))
        t.pack(fill="both", expand=True)
        cols = ["id"] + CLAVES
        self.tabla = tabla_con_scroll(marco, columns=cols, show="headings", selectmode="extended")
        self.tabla.heading("id", text="ID", command=lambda: self.ordenar("id"))
        self.tabla.column("id", width=px(50), anchor="e", stretch=False)
        for c in CAMPOS:
            self.tabla.heading(c.clave, text=c.etiqueta.upper(), command=lambda k=c.clave: self.ordenar(k),
                               anchor="center" if c.tipo == "fecha" else "w")
            self.tabla.column(c.clave, width=px(c.ancho), stretch=False,
                              anchor="center" if c.tipo == "fecha" else "w")
        self.tabla.tag_configure("par", background=C["fila_alterna"])
        self.tabla.tag_configure("caducada", background=C["rojo_suave"])
        self.tabla.tag_configure("proxima", background=C["ambar_suave"])
        self.tabla.tag_configure("baja", foreground=C["terciario"])
        self._aplicar_columnas_visibles()
        self._id_filtro = None

    # ------------------------------------------------------------ página resumen
    def _pagina_resumen(self):
        px = self.tema.px
        p = ttk.Frame(self.contenido, padding=(px(28), px(22), px(28), px(16)))
        self.paginas.append(p)
        _, self.sub_resumen = self._cabecera(p, "Resumen")

        fila = ttk.Frame(p)
        fila.pack(fill="x")
        self.cifras: dict[str, tk.StringVar] = {}
        for i, (clave, texto, estilo) in enumerate((
                ("total", "Equipos en total", "CifraAzul.TLabel"),
                ("activos", "Activos", "CifraVerde.TLabel"),
                ("caducadas", "Garantía caducada", "CifraRojo.TLabel"),
                ("proximas", f"Caducan en ≤ {DIAS_AVISO_GARANTIA} días", "CifraAmbar.TLabel"),
                ("sin_garantia", "Sin fecha de garantía", "CifraGris.TLabel"))):
            t, inn = tarjeta(fila)
            t.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else px(12), 0))
            fila.columnconfigure(i, weight=1, uniform="cifras")
            var = tk.StringVar(value="0")
            ttk.Label(inn, textvariable=var, style=estilo).pack(anchor="w")
            ttk.Label(inn, text=texto, style="EnTarjetaSec.TLabel").pack(anchor="w")
            self.cifras[clave] = var

        rejilla = ttk.Frame(p)
        rejilla.pack(fill="x", pady=(px(14), 0))
        self.graficos: dict[str, GraficoBarras] = {}
        grupos = [("delegacion", "Por delegación"), ("marca", "Por marca"), ("os", "Por sistema operativo"),
                  ("estado", "Por estado"), ("office", "Por Office"), ("antivirus", "Por antivirus"),
                  ("ram", "Por memoria RAM"), ("anio", "Por año de compra")]
        for i, (clave, titulo) in enumerate(grupos):
            t, inn = tarjeta(rejilla)
            t.grid(row=i // 4, column=i % 4, sticky="nsew", padx=(0 if i % 4 == 0 else px(12), 0),
                   pady=(0, px(12)))
            ttk.Label(inn, text=titulo, style="EnTarjetaTitulo.TLabel").pack(anchor="w", pady=(0, px(8)))
            al_pulsar = (lambda v, k=clave: self._filtrar_desde_resumen(k, v)) if clave in self.filtros else None
            g = GraficoBarras(inn, self.tema, PALETA_GRAFICOS[i % len(PALETA_GRAFICOS)], al_pulsar, max_filas=5)
            g.pack(fill="both", expand=True)
            self.graficos[clave] = g
        for c in range(4):
            rejilla.columnconfigure(c, weight=1, uniform="graf")

        ttk.Label(p, text="GARANTÍAS QUE REQUIEREN ATENCIÓN", style="Seccion.TLabel").pack(
            anchor="w", padx=px(4), pady=(px(2), px(6)))
        t, marco = tarjeta(p)
        t.configure(padding=px(6))
        t.pack(fill="both", expand=True)
        cols = ("fin_garantia", "dias", "nombre_equipo", "persona", "delegacion", "marca", "modelo", "num_serie")
        self.tabla_garantias = tabla_con_scroll(marco, columns=cols, show="headings", height=5)
        for c in cols:
            texto = "DÍAS" if c == "dias" else CAMPOS_POR_CLAVE[c].etiqueta.upper()
            self.tabla_garantias.heading(c, text=texto, anchor="e" if c == "dias" else "w")
            self.tabla_garantias.column(c, width=px(70 if c == "dias" else 150), anchor="e" if c == "dias" else "w")
        self.tabla_garantias.tag_configure("caducada", foreground=C["rojo"])
        self.tabla_garantias.tag_configure("proxima", foreground="#C76E00")
        self.tabla_garantias.bind("<Double-1>", self._editar_desde_garantias)

    # ------------------------------------------------------------ página historial
    def _pagina_historial(self):
        px = self.tema.px
        p = ttk.Frame(self.contenido, padding=(px(28), px(22), px(28), px(16)))
        self.paginas.append(p)
        acciones, self.sub_historial = self._cabecera(p, "Historial")
        ttk.Button(acciones, text="Exportar a Excel", command=self.exportar_historial).pack(side="right")
        ttk.Button(acciones, text="Vaciar", style="Peligro.TButton", command=self.vaciar_historial).pack(
            side="right", padx=(0, 8))
        t, marco = tarjeta(p)
        t.configure(padding=px(6))
        t.pack(fill="both", expand=True)
        cols = ("fecha", "accion", "equipo", "detalle")
        self.tabla_hist = tabla_con_scroll(marco, columns=cols, show="headings")
        for c, txt, w in (("fecha", "FECHA", 150), ("accion", "ACCIÓN", 130), ("equipo", "EQUIPO", 280),
                          ("detalle", "DETALLE", 800)):
            self.tabla_hist.heading(c, text=txt, anchor="w")
            self.tabla_hist.column(c, width=px(w), stretch=c == "detalle")
        for accion, color in (("Alta", C["verde"]), ("Modificación", C["acento"]), ("Baja (eliminado)", C["rojo"]),
                              ("Importación", "#5E5CE6"), ("Exportación", C["secundario"])):
            self.tabla_hist.tag_configure(accion, foreground=color)

    # ------------------------------------------------------------ datos
    def recargar(self):
        self.equipos = self.bd.listar()
        for eq in self.equipos:
            eq["_texto"] = texto_fila(eq)
            eq["_garantia"] = estado_garantia(eq.get("fin_garantia", ""))
        for clave, cb in self.combos.items():
            valores = sorted({eq[clave] for eq in self.equipos if eq.get(clave)}, key=str.lower)
            if clave == "estado":
                valores = [e for e in ESTADOS if e in valores] + [v for v in valores if v not in ESTADOS]
            cb["values"] = [TODOS_FILTRO[clave]] + valores
            if self.filtros[clave].get() not in cb["values"]:
                self.filtros[clave].set(TODOS_FILTRO[clave])
        n = len(self.equipos)
        no_baja = [e for e in self.equipos if e.get("estado") != "Baja"]
        cad = sum(1 for e in no_baja if e["_garantia"] == "caducada")
        prox = sum(1 for e in no_baja if e["_garantia"] == "proxima")
        sub = f"{n} equipo{'s' if n != 1 else ''}"
        if cad:
            sub += f" · {cad} con garantía caducada"
        if prox:
            sub += f" · {prox} caducan pronto"
        self.sub_inventario.set(sub)
        self.lbl_lateral_total.configure(text=f"{n} equipos registrados")
        self.lbl_lateral_bd.configure(text=f"Base de datos: {self.bd.ruta.name}")
        self.aplicar_filtro()
        self._cambio_pagina()

    def _programar_filtro(self):
        if self._id_filtro:
            self.after_cancel(self._id_filtro)
        self._id_filtro = self.after(200, self.aplicar_filtro)

    def aplicar_filtro(self):
        self._id_filtro = None
        tokens = [normalizar(t) for t in self.var_buscar.get().split() if normalizar(t)]
        fil = {k: v.get() for k, v in self.filtros.items() if v.get() != TODOS_FILTRO[k]}
        gar = _MAPA_GARANTIA.get(self.var_garantia.get())
        res = []
        for eq in self.equipos:
            if any(eq.get(k, "") != v for k, v in fil.items()):
                continue
            if gar and eq["_garantia"] != gar:
                continue
            if tokens and not all(t in eq["_texto"] for t in tokens):
                continue
            res.append(eq)
        clave, desc = self.orden
        if clave == "id":
            res.sort(key=lambda e: e["id"], reverse=desc)
        else:
            res.sort(key=lambda e: _clave_orden(e.get(clave, "")), reverse=desc)
        self.visibles = res
        hay_filtro = bool(tokens or fil or gar)
        self.btn_limpiar.state(["!disabled"] if hay_filtro else ["disabled"])
        self._pintar_tabla()

    def _pintar_tabla(self):
        if getattr(self, "_editor", None):
            self._cerrar_editor(True)
        seleccion = set(self.tabla.selection())
        self.tabla.delete(*self.tabla.get_children())
        for n, eq in enumerate(self.visibles):
            self.tabla.insert("", "end", iid=str(eq["id"]), values=self._valores_fila(eq),
                              tags=self._tags_fila(eq, n))
        restaurar = [i for i in seleccion if self.tabla.exists(i)]
        if restaurar:
            self.tabla.selection_set(restaurar)
        for c in CAMPOS + [None]:
            clave = c.clave if c else "id"
            texto = c.etiqueta.upper() if c else "ID"
            if self.orden[0] == clave:
                texto += "  ↓" if self.orden[1] else "  ↑"
            self.tabla.heading(clave, text=texto)
        self._actualizar_estado()

    @staticmethod
    def _valores_fila(eq: dict) -> list:
        return [eq["id"]] + [mostrar_fecha(eq[c.clave]) if c.tipo == "fecha"
                             else (eq[c.clave] or "").replace("\n", " ") for c in CAMPOS]

    @staticmethod
    def _tags_fila(eq: dict, n: int) -> list[str]:
        tags = []
        if eq.get("estado") == "Baja":
            tags.append("baja")
        elif eq["_garantia"] in ("caducada", "proxima"):
            tags.append(eq["_garantia"])
        if (not tags or tags == ["baja"]) and n % 2:
            tags.append("par")
        return tags

    def _refrescar_equipo(self, equipo_id: int):
        """Vuelve a leer un equipo de la BD y actualiza su fila sin repintar toda la tabla."""
        nuevo = self.bd.obtener(equipo_id)
        if not nuevo:
            return
        nuevo["_texto"] = texto_fila(nuevo)
        nuevo["_garantia"] = estado_garantia(nuevo.get("fin_garantia", ""))
        for lista in (self.equipos, self.visibles):
            for i, eq in enumerate(lista):
                if eq["id"] == equipo_id:
                    lista[i] = nuevo
                    break
        iid = str(equipo_id)
        if self.tabla.exists(iid):
            n = self.tabla.index(iid)
            self.tabla.item(iid, values=self._valores_fila(nuevo), tags=self._tags_fila(nuevo, n))

    # ------------------------------------------------------------ edición en la propia tabla
    def _columna_en(self, x: int) -> str | None:
        col = self.tabla.identify_column(x)  # '#3'
        try:
            idx = int(col.lstrip("#")) - 1
        except ValueError:
            return None
        visibles = list(self.tabla["displaycolumns"])
        if visibles in (["#all"], ("#all",)) or visibles == ["#all"]:
            visibles = list(self.tabla["columns"])
        return visibles[idx] if 0 <= idx < len(visibles) else None

    def editar_celda(self, iid: str, clave: str):
        """Abre un campo de edición encima de la celda (Intro/Tab guarda, Esc cancela)."""
        self._cerrar_editor(guardar=False)
        if clave not in CLAVES or not self.tabla.exists(iid):
            return
        self.tabla.see(iid)
        self._ver_columna(clave)
        self.update_idletasks()
        caja = self.tabla.bbox(iid, clave)
        if not caja:
            return
        x, y, w, h = caja
        eq = next((e for e in self.equipos if str(e["id"]) == iid), None)
        if eq is None:
            return
        campo = CAMPOS_POR_CLAVE[clave]
        valor = mostrar_fecha(eq[clave]) if campo.tipo == "fecha" else (eq[clave] or "")
        var = tk.StringVar(value=valor)
        if campo.tipo == "fecha":
            w_ed = ttk.Entry(self.tabla, textvariable=var, style="Blanco.TEntry")
        else:
            valores = ESTADOS if clave == "estado" else self.bd.valores_distintos(clave)
            w_ed = ttk.Combobox(self.tabla, textvariable=var, values=valores, style="Blanco.TCombobox")
        alto = max(h, self.tema.px(32))
        w_ed.place(x=x, y=y + (h - alto) // 2, width=max(w, self.tema.px(140)), height=alto)
        w_ed.focus_set()
        w_ed.select_range(0, "end")
        w_ed.icursor("end")
        self._editor = (w_ed, var, iid, clave)
        w_ed.bind("<Return>", lambda e: (self._cerrar_editor(True), self.tabla.focus_set(), "break")[-1])
        w_ed.bind("<KP_Enter>", lambda e: (self._cerrar_editor(True), self.tabla.focus_set(), "break")[-1])
        w_ed.bind("<Escape>", lambda e: (self._cerrar_editor(False), self.tabla.focus_set(), "break")[-1])
        w_ed.bind("<Tab>", lambda e: (self._mover_editor(+1), "break")[-1])
        w_ed.bind("<ISO_Left_Tab>", lambda e: (self._mover_editor(-1), "break")[-1])
        w_ed.bind("<Shift-Tab>", lambda e: (self._mover_editor(-1), "break")[-1])
        w_ed.bind("<FocusOut>", lambda e: self.after(80, self._cerrar_si_sin_foco))

    def _ver_columna(self, clave: str):
        """Desplaza la tabla horizontalmente para que la columna `clave` quede a la vista."""
        cols = list(self.tabla["displaycolumns"])
        if cols in (["#all"],):
            cols = list(self.tabla["columns"])
        anchos = [int(self.tabla.column(c, "width")) for c in cols]
        total = sum(anchos) or 1
        inicio = sum(anchos[:cols.index(clave)]) if clave in cols else 0
        fin = inicio + (anchos[cols.index(clave)] if clave in cols else 0)
        a, b = self.tabla.xview()
        vis_a, vis_b = a * total, b * total
        if inicio < vis_a or fin > vis_b:
            self.tabla.xview_moveto(max(0.0, (inicio - 20) / total))

    def _cerrar_si_sin_foco(self):
        ed = getattr(self, "_editor", None)
        if not ed:
            return
        try:
            foco = self.focus_get()
        except (KeyError, tk.TclError):
            foco = None
        # el desplegable del combobox no cuenta como perder el foco
        if foco is not None and (foco == ed[0] or str(foco).startswith(str(ed[0]))):
            return
        if foco is not None and "popdown" in str(foco):
            return
        self._cerrar_editor(True)

    def _mover_editor(self, paso: int):
        ed = getattr(self, "_editor", None)
        if not ed:
            return
        _, _, iid, clave = ed
        self._cerrar_editor(True)
        visibles = [c for c in self.tabla["displaycolumns"] if c in CLAVES] or \
            [c for c in CLAVES if c not in self.columnas_ocultas]
        i = visibles.index(clave) + paso if clave in visibles else 0
        if 0 <= i < len(visibles):
            self.editar_celda(iid, visibles[i])

    def _cerrar_editor(self, guardar: bool):
        ed = getattr(self, "_editor", None)
        if not ed:
            return
        self._editor = None
        w_ed, var, iid, clave = ed
        nuevo = var.get().strip()
        try:
            w_ed.destroy()
        except tk.TclError:
            pass
        if not guardar:
            return
        if CAMPOS_POR_CLAVE[clave].tipo == "fecha":
            nuevo = parsear_fecha(nuevo)
        eq = next((e for e in self.equipos if str(e["id"]) == iid), None)
        if eq is None or (eq.get(clave) or "") == nuevo:
            return
        self.bd.actualizar_campo(int(iid), clave, nuevo)
        self._refrescar_equipo(int(iid))
        if self.pagina_actual == 0:
            self.estado.set(f"Guardado: {CAMPOS_POR_CLAVE[clave].etiqueta} = «{mostrar_fecha(nuevo) if CAMPOS_POR_CLAVE[clave].tipo == 'fecha' else nuevo}»")

    def _actualizar_estado(self):
        sel = len(self.tabla.selection())
        txt = f"Mostrando {len(self.visibles)} de {len(self.equipos)}"
        if sel:
            txt += f" · {sel} seleccionado{'s' if sel != 1 else ''}"
        self.estado.set(txt)
        for b in self.botones_seleccion:
            b.state(["!disabled"] if sel else ["disabled"])

    def ordenar(self, clave: str):
        actual, desc = self.orden
        self.orden = (clave, not desc if actual == clave else False)
        self.aplicar_filtro()

    def limpiar_filtros(self):
        for k, v in self.filtros.items():
            v.set(TODOS_FILTRO[k])
        self.var_garantia.set(TODOS)
        self.var_buscar.set("")
        self.aplicar_filtro()

    def seleccionados(self) -> list[int]:
        return [int(i) for i in self.tabla.selection()]

    def seleccionar_todo(self):
        self.tabla.selection_set(self.tabla.get_children())

    # ------------------------------------------------------------ acciones
    def nuevo(self, base: dict | None = None):
        f = FormularioEquipo(self, base, "Nuevo equipo")
        self.wait_window(f)
        if f.resultado:
            nuevo = self.bd.insertar(f.resultado)
            self.recargar()
            self._seleccionar(nuevo)

    def editar(self, equipo_id: int | None = None):
        if equipo_id is None:
            sel = self.seleccionados()
            if not sel:
                messagebox.showinfo(TITULO, "Selecciona un equipo para editar.")
                return
            equipo_id = sel[0]
        eq = self.bd.obtener(equipo_id)
        if not eq:
            return
        f = FormularioEquipo(self, eq, eq.get("nombre_equipo") or eq.get("persona") or f"Equipo {eq['id']}")
        self.wait_window(f)
        if f.resultado:
            self.bd.actualizar(equipo_id, f.resultado)
            self.recargar()
            self._seleccionar(equipo_id)

    def duplicar(self):
        sel = self.seleccionados()
        if not sel:
            return
        eq = self.bd.obtener(sel[0]) or {}
        base = {k: eq.get(k, "") for k in CLAVES}
        for k in ("num_serie", "persona", "usuario_pc", "teamviewer", "nombre_equipo"):
            base[k] = ""
        self.nuevo(base)

    def eliminar(self):
        sel = self.seleccionados()
        if not sel:
            return
        if len(sel) == 1:
            eq = self.bd.obtener(sel[0]) or {}
            pregunta = (f"¿Eliminar el equipo «{eq.get('nombre_equipo') or '-'}» de "
                        f"{eq.get('persona') or '-'}?")
        else:
            pregunta = f"¿Eliminar los {len(sel)} equipos seleccionados?"
        pregunta += "\n\nConsejo: si el equipo se ha retirado puedes marcarlo como «Baja» en lugar de eliminarlo."
        if messagebox.askyesno(TITULO, pregunta, icon="warning"):
            self.bd.eliminar(sel)
            self.recargar()

    def cambiar_campo(self, clave: str, valor: str):
        sel = self.seleccionados()
        if not sel:
            messagebox.showinfo(TITULO, "Selecciona uno o varios equipos.")
            return
        for i in sel:
            eq = self.bd.obtener(i)
            if eq:
                eq[clave] = valor
                self.bd.actualizar(i, eq, commit=False)
        self.bd.commit()
        self.recargar()

    def cambiar_campo_dialogo(self, clave: str | None = None):
        """Cambia cualquier campo en todos los equipos seleccionados."""
        sel = self.seleccionados()
        if not sel:
            messagebox.showinfo(TITULO, "Selecciona uno o varios equipos.")
            return
        d = Dialogo(self, "Editar un campo en la selección",
                    f"El valor se aplicará a {len(sel)} equipo{'s' if len(sel) != 1 else ''}")
        etiquetas = [c.etiqueta for c in CAMPOS]
        var_campo = tk.StringVar(value=CAMPOS_POR_CLAVE[clave].etiqueta if clave else etiquetas[0])
        var = tk.StringVar()

        def clave_actual() -> str:
            return next(c.clave for c in CAMPOS if c.etiqueta == var_campo.get())

        def aceptar():
            k = clave_actual()
            valor = var.get().strip()
            if CAMPOS_POR_CLAVE[k].tipo == "fecha":
                valor = parsear_fecha(valor)
            d.destroy()
            self.cambiar_campo(k, valor)
        d.pie("Aplicar", aceptar)
        t, inn = tarjeta(d.cuerpo)
        t.pack(fill="both", expand=True)
        ttk.Label(inn, text="Campo", style="EnTarjetaCampo.TLabel").pack(anchor="w", pady=(0, 2))
        cb_campo = ttk.Combobox(inn, textvariable=var_campo, values=etiquetas, state="readonly", width=40,
                                style="Blanco.TCombobox")
        cb_campo.pack(fill="x", pady=(0, self.tema.px(10)))
        ttk.Label(inn, text="Nuevo valor (vacío = borrar el campo)", style="EnTarjetaCampo.TLabel").pack(
            anchor="w", pady=(0, 2))
        cb = ttk.Combobox(inn, textvariable=var, width=40, style="Blanco.TCombobox")
        cb.pack(fill="x")

        def al_cambiar_campo(_e=None):
            k = clave_actual()
            cb["values"] = ESTADOS if k == "estado" else self.bd.valores_distintos(k)
        cb_campo.bind("<<ComboboxSelected>>", al_cambiar_campo)
        al_cambiar_campo()
        cb.focus_set()
        d.bind("<Return>", lambda e: aceptar())
        d.centrar()

    def copiar_fila(self):
        sel = self.seleccionados()
        if not sel:
            return
        lineas = ["\t".join(c.cabecera for c in CAMPOS)]
        por_id = {e["id"]: e for e in self.equipos}
        for i in sel:
            eq = por_id.get(i, {})
            lineas.append("\t".join(mostrar_fecha(eq.get(c.clave, "")) if c.tipo == "fecha"
                                    else (eq.get(c.clave, "") or "").replace("\n", " ") for c in CAMPOS))
        self.clipboard_clear()
        self.clipboard_append("\n".join(lineas))
        self.estado.set(f"{len(sel)} fila(s) copiadas al portapapeles · se pueden pegar en Excel")

    def _seleccionar(self, equipo_id: int):
        iid = str(equipo_id)
        if self.tabla.exists(iid):
            self.tabla.selection_set(iid)
            self.tabla.see(iid)
            self.tabla.focus(iid)

    def _doble_click(self, evento):
        if self.tabla.identify_region(evento.x, evento.y) != "cell":
            return
        fila = self.tabla.identify_row(evento.y)
        if not fila:
            return
        clave = self._columna_en(evento.x)
        if clave in CLAVES:
            self.editar_celda(fila, clave)
        else:
            self.editar(int(fila))
        return "break"

    def _f2(self, _evento=None):
        foco = self.tabla.focus() or (self.tabla.selection() or [None])[0]
        if foco:
            visibles = [c for c in CLAVES if c not in self.columnas_ocultas]
            if visibles:
                self.editar_celda(foco, visibles[0])
        return "break"

    def _editar_celda_contexto(self):
        if self._fila_contexto and self._clave_contexto in CLAVES:
            self.editar_celda(self._fila_contexto, self._clave_contexto)

    def _click_derecho(self, evento):
        self._cerrar_editor(True)
        fila = self.tabla.identify_row(evento.y)
        clave = self._columna_en(evento.x)
        self._fila_contexto = fila
        self._clave_contexto = clave if clave in CLAVES else None
        if fila and fila not in self.tabla.selection():
            self.tabla.selection_set(fila)
        if self.tabla.selection():
            self.menu_contexto.tk_popup(evento.x_root, evento.y_root)

    # ------------------------------------------------------------ importar / exportar
    def importar(self):
        ruta = filedialog.askopenfilename(
            title="Importar inventario", parent=self,
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("CSV", "*.csv *.txt"), ("Todos", "*.*")])
        if not ruta:
            return
        if Path(ruta).suffix.lower() == ".xls":
            messagebox.showerror(TITULO, "El formato .xls antiguo no está soportado.\n"
                                         "Ábrelo en Excel y guárdalo como .xlsx.")
            return
        try:
            d = DialogoImportar(self, Path(ruta))
        except _Cancelado:
            return
        self.wait_window(d)
        if d.importado:
            self.limpiar_filtros()
            self.recargar()
            self.mostrar(0)

    def _pedir_destino(self, extension: str, nombre: str) -> Path | None:
        tipos = [("Excel", "*.xlsx")] if extension == ".xlsx" else [("CSV (separado por ;)", "*.csv")]
        ruta = filedialog.asksaveasfilename(parent=self, defaultextension=extension, filetypes=tipos,
                                            initialfile=nombre)
        return Path(ruta) if ruta else None

    def exportar(self, formato: str, alcance: str):
        if alcance == "todo":
            datos = self.equipos
        elif alcance == "seleccion":
            ids = set(self.seleccionados())
            datos = [e for e in self.visibles if e["id"] in ids]
            if not datos:
                messagebox.showinfo(TITULO, "No hay equipos seleccionados.")
                return
        else:
            datos = self.visibles
        ext = ".xlsx" if formato == "excel" else ".csv"
        destino = self._pedir_destino(ext, f"inventario_{datetime.now():%Y%m%d}{ext}")
        if not destino:
            return
        try:
            if formato == "excel":
                exportar_excel(destino, datos)
            else:
                exportar_csv(destino, datos)
        except PermissionError:
            messagebox.showerror(TITULO, "No se puede escribir el archivo.\n¿Está abierto en Excel?")
            return
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(TITULO, f"Error al exportar:\n{exc}")
            return
        self.bd.registrar(None, destino.name, "Exportación", f"{len(datos)} equipos")
        self.bd.commit()
        if messagebox.askyesno(TITULO, f"Se han exportado {len(datos)} equipos a:\n{destino}\n\n¿Abrir el archivo?"):
            abrir_archivo(destino)

    def plantilla(self):
        destino = self._pedir_destino(".xlsx", "plantilla_inventario.xlsx")
        if destino:
            try:
                exportar_plantilla(destino)
            except PermissionError:
                messagebox.showerror(TITULO, "No se puede escribir el archivo.\n¿Está abierto en Excel?")
                return
            messagebox.showinfo(TITULO, f"Plantilla guardada en:\n{destino}")

    def exportar_historial(self):
        destino = self._pedir_destino(".xlsx", f"historial_{datetime.now():%Y%m%d}.xlsx")
        if not destino:
            return
        from openpyxl import Workbook
        from openpyxl.styles import Font
        wb = Workbook()
        ws = wb.active
        ws.title = "Historial"
        ws.append(["Fecha", "Acción", "Equipo", "Detalle"])
        for c in ws[1]:
            c.font = Font(bold=True)
        for h in self.bd.historial(100000):
            ws.append([h["fecha"], h["accion"], h["equipo"], h["detalle"]])
        for col, w in zip("ABCD", (20, 18, 40, 100)):
            ws.column_dimensions[col].width = w
        try:
            wb.save(destino)
        except PermissionError:
            messagebox.showerror(TITULO, "No se puede escribir el archivo.\n¿Está abierto en Excel?")
            return
        abrir_archivo(destino)

    def vaciar_historial(self):
        if messagebox.askyesno(TITULO, "¿Borrar todo el historial de cambios?", icon="warning"):
            self.bd.vaciar_historial()
            self._pintar_historial()

    # ------------------------------------------------------------ copias
    def copia(self):
        destino = filedialog.asksaveasfilename(parent=self, defaultextension=".db",
                                               filetypes=[("Base de datos", "*.db")],
                                               initialfile=f"copia_inventario_{datetime.now():%Y%m%d_%H%M}.db")
        if destino:
            self.bd.copia_seguridad(destino)
            messagebox.showinfo(TITULO, f"Copia de seguridad guardada en:\n{destino}")

    def restaurar(self):
        origen = filedialog.askopenfilename(parent=self, filetypes=[("Base de datos", "*.db"), ("Todos", "*.*")],
                                            initialdir=str(self.bd.ruta.parent / "copias"))
        if not origen:
            return
        if not messagebox.askyesno(TITULO, "Se sustituirán los datos actuales por los de la copia.\n"
                                           "(Se guardará antes una copia del estado actual.)\n\n¿Continuar?",
                                   icon="warning"):
            return
        try:
            self.bd.restaurar(origen)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(TITULO, f"No es una copia válida:\n{exc}")
            return
        self.recargar()
        messagebox.showinfo(TITULO, "Copia restaurada correctamente.")

    def cambiar_bd(self):
        ruta = filedialog.asksaveasfilename(
            parent=self, title="Elegir o crear base de datos", defaultextension=".db",
            filetypes=[("Base de datos", "*.db")], confirmoverwrite=False,
            initialdir=str(self.bd.ruta.parent), initialfile=self.bd.ruta.name)
        if not ruta:
            return
        try:
            nueva = BaseDatos(ruta)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(TITULO, f"No se pudo abrir la base de datos:\n{exc}")
            return
        self.bd.cerrar()
        self.bd = nueva
        cfg = leer_config()
        cfg["ruta_bd"] = str(Path(ruta).resolve())
        guardar_config(cfg)
        self.recargar()
        messagebox.showinfo(TITULO, f"Usando la base de datos:\n{ruta}\n\n"
                                    "Se recordará la próxima vez que abras el programa.")

    # ------------------------------------------------------------ columnas
    def _aplicar_columnas_visibles(self):
        self.tabla["displaycolumns"] = ["id"] + [c for c in CLAVES if c not in self.columnas_ocultas]

    def dialogo_columnas(self):
        d = Dialogo(self, "Columnas visibles", "Elige qué columnas se muestran en la tabla.")
        variables = {}

        def aceptar():
            self.columnas_ocultas = {k for k, v in variables.items() if not v.get()}
            cfg = leer_config()
            cfg["columnas_ocultas"] = sorted(self.columnas_ocultas)
            guardar_config(cfg)
            self._aplicar_columnas_visibles()
            d.destroy()
        d.pie("Aplicar", aceptar)
        t, inn = tarjeta(d.cuerpo)
        t.pack(fill="both", expand=True)
        for i, c in enumerate(CAMPOS):
            v = tk.BooleanVar(value=c.clave not in self.columnas_ocultas)
            ttk.Checkbutton(inn, text=c.etiqueta, variable=v).grid(row=i % 10, column=i // 10, sticky="w",
                                                                    padx=(0, self.tema.px(24)))
            variables[c.clave] = v
        d.centrar()

    # ------------------------------------------------------------ resumen / historial
    def _cambio_pagina(self):
        if self.pagina_actual == 1:
            self._pintar_resumen()
        elif self.pagina_actual == 2:
            self._pintar_historial()

    def _pintar_resumen(self):
        eqs = self.equipos
        hoy = date.today()
        self.cifras["total"].set(str(len(eqs)))
        self.cifras["activos"].set(str(sum(1 for e in eqs if (e.get("estado") or "Activo") == "Activo")))
        no_baja = [e for e in eqs if e.get("estado") != "Baja"]
        self.cifras["caducadas"].set(str(sum(1 for e in no_baja if e["_garantia"] == "caducada")))
        self.cifras["proximas"].set(str(sum(1 for e in no_baja if e["_garantia"] == "proxima")))
        self.cifras["sin_garantia"].set(str(sum(1 for e in no_baja if e["_garantia"] == "sin_dato")))
        personas = len({normalizar(e["persona"]) for e in eqs if e.get("persona")})
        delegaciones = len({e["delegacion"] for e in eqs if e.get("delegacion")})
        self.sub_resumen.set(f"{personas} personas · {delegaciones} delegaciones · actualizado {datetime.now():%H:%M}")

        for clave, g in self.graficos.items():
            if clave == "anio":
                cont = Counter((str(fecha_a_date(e["fecha_compra"]).year) if fecha_a_date(e["fecha_compra"])
                                else "(sin dato)") for e in eqs)
                items = sorted(cont.items(), key=lambda x: x[0], reverse=True)
            else:
                cont = Counter((e.get(clave) or "(sin dato)") for e in eqs)
                items = cont.most_common()
            g.poner_datos(items)

        tg = self.tabla_garantias
        tg.delete(*tg.get_children())
        criticos = [e for e in no_baja if e["_garantia"] in ("caducada", "proxima")]
        criticos.sort(key=lambda e: e["fin_garantia"])
        for e in criticos:
            dias = (fecha_a_date(e["fin_garantia"]) - hoy).days
            texto_dias = f"hace {-dias}" if dias < 0 else f"en {dias}"
            tg.insert("", "end", iid=str(e["id"]), tags=(e["_garantia"],),
                      values=(mostrar_fecha(e["fin_garantia"]), texto_dias, e["nombre_equipo"], e["persona"],
                              e["delegacion"], e["marca"], e["modelo"], e["num_serie"]))

    def _filtrar_desde_resumen(self, clave: str, valor: str):
        if clave not in self.filtros:
            return
        self.limpiar_filtros()
        self.filtros[clave].set(valor)
        self.mostrar(0)
        self.aplicar_filtro()

    def _editar_desde_garantias(self, _e):
        sel = self.tabla_garantias.selection()
        if sel:
            self.editar(int(sel[0]))
            self._pintar_resumen()

    def _pintar_historial(self):
        self.tabla_hist.delete(*self.tabla_hist.get_children())
        hist = self.bd.historial()
        for h in hist:
            self.tabla_hist.insert("", "end", tags=(h["accion"],),
                                   values=(h["fecha"], h["accion"], h["equipo"], h["detalle"].replace("\n", " ")))
        self.sub_historial.set(f"Últimos {len(hist)} cambios registrados" if hist else "Todavía no hay cambios")

    def _aviso_garantias(self):
        cfg = leer_config()
        if not cfg.get("avisar_garantias", True):
            return
        hoy = date.today().isoformat()
        if cfg.get("ultimo_aviso") == hoy:
            return
        prox = [e for e in self.equipos if e["_garantia"] == "proxima" and e.get("estado") != "Baja"]
        if prox:
            cfg["ultimo_aviso"] = hoy
            guardar_config(cfg)
            if messagebox.askyesno(TITULO, f"Hay {len(prox)} equipo(s) cuya garantía caduca en los próximos "
                                           f"{DIAS_AVISO_GARANTIA} días.\n\n¿Ver el listado?"):
                self.mostrar(1)

    # ------------------------------------------------------------ ayuda
    def ayuda_formato(self):
        txt = ("El programa importa archivos .xlsx (y .csv) con el formato de la plantilla:\n\n"
               "• Filas 1-2: títulos de grupo (PROPIETARIO DEL ORDENADOR / DATOS DEL ORDENADOR).\n"
               "• Fila 3 (o cualquier fila de las 30 primeras): cabeceras de columna.\n"
               "• Datos a partir de la fila siguiente a las cabeceras.\n\n"
               "Columnas reconocidas (da igual el orden, mayúsculas o tildes):\n" +
               ", ".join(c.cabecera for c in CAMPOS) +
               "\n\nNingún campo se usa como identificador: cada fila es un equipo y todos los campos "
               "(también el Nº de serie) se pueden editar. Al importar puedes añadir todas las filas, "
               "saltar las que sean idénticas en todos los campos o reemplazar el inventario completo.\n\n"
               "Puedes descargar una plantilla vacía desde Archivo → Guardar plantilla Excel vacía.")
        messagebox.showinfo("Formato de importación", txt)

    def acerca(self):
        messagebox.showinfo("Acerca de", f"{TITULO}\nVersión {__version__}\n\n"
                                         f"Datos guardados en:\n{self.bd.ruta}\n\n"
                                         "Copias automáticas diarias en la carpeta «copias».")

    def _salir(self):
        try:
            self.bd.commit()
            self.bd.cerrar()
        finally:
            self.destroy()


def _clave_orden(valor: str):
    """Orden natural: números como números, el resto como texto sin tildes."""
    v = (valor or "").strip()
    if not v:
        return (2, "")
    try:
        return (0, float(v.replace(",", ".").split()[0]))
    except (ValueError, IndexError):
        return (1, normalizar(v))


def _autoprueba(argv: list[str]) -> int:
    """Comprueba que el ejecutable arranca: crea la ventana, importa/exporta un Excel y sale.
    Uso: InventarioPC.exe --selftest <carpeta_salida>"""
    import tempfile
    salida = Path(argv[0]) if argv else Path(tempfile.gettempdir())
    salida.mkdir(parents=True, exist_ok=True)
    informe = salida / "selftest.txt"
    try:
        db = salida / "selftest.db"
        if db.exists():
            db.unlink()
        app = App(str(db))
        app._aviso_garantias = lambda: None
        app.bd.insertar({"persona": "Prueba", "nombre_equipo": "PC-TEST", "num_serie": "SN-TEST",
                         "fin_garantia": "2030-01-01"})
        app.recargar()
        xlsx = salida / "selftest.xlsx"
        exportar_excel(xlsx, app.equipos)
        filas = leer_excel(xlsx).filas
        assert len(filas) == 1 and filas[0]["num_serie"] == "SN-TEST", filas
        app.mostrar(1)
        app.update()
        app.mostrar(2)
        app.update()
        app.after(500, app._salir)
        app.mainloop()
        informe.write_text("OK\n", encoding="utf-8")
        return 0
    except Exception:  # noqa: BLE001
        import traceback
        informe.write_text("ERROR\n" + traceback.format_exc(), encoding="utf-8")
        return 1


def main(argv: list[str] | None = None):
    argv = sys.argv[1:] if argv is None else argv
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:  # noqa: BLE001
            pass
    if argv and argv[0] == "--selftest":
        return _autoprueba(argv[1:])
    ruta_bd = None
    if argv and argv[0].lower().endswith(".db"):
        ruta_bd = argv[0]
    app = App(ruta_bd)
    if argv and argv[0].lower().endswith((".xlsx", ".xlsm", ".csv")):
        def abrir():
            try:
                d = DialogoImportar(app, Path(argv[0]))
            except _Cancelado:
                return
            app.wait_window(d)
            if d.importado:
                app.recargar()
        app.after(300, abrir)
    app.mainloop()


if __name__ == "__main__":
    sys.exit(main())
