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
from .campos import (CAMPOS, CAMPOS_POR_CLAVE, CLAVES, DIAS_AVISO_GARANTIA, ESTADOS, GRUPOS,
                     estado_garantia, fecha_a_date, mostrar_fecha, normalizar, parsear_fecha)
from .db import BaseDatos, carpeta_datos, guardar_config, leer_config
from .excel_io import exportar_csv, exportar_excel, exportar_plantilla, leer_excel

TITULO = "Inventario de Equipos"
TODOS = "(Todos)"
FILTROS_GARANTIA = [TODOS, "Vigente", f"Caduca en ≤ {DIAS_AVISO_GARANTIA} días", "Caducada", "Sin fecha"]
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


# =====================================================================
#  Formulario de alta / edición
# =====================================================================
class FormularioEquipo(tk.Toplevel):
    def __init__(self, app: "App", equipo: dict | None = None, titulo: str = "Nuevo equipo"):
        super().__init__(app)
        self.app = app
        self.equipo = equipo
        self.resultado: dict | None = None
        self.title(titulo)
        self.transient(app)
        self.resizable(True, False)
        self.widgets: dict[str, tk.Widget] = {}
        self.vars: dict[str, tk.StringVar] = {}

        cuerpo = ttk.Frame(self, padding=12)
        cuerpo.pack(fill="both", expand=True)
        cuerpo.columnconfigure(0, weight=1)
        cuerpo.columnconfigure(1, weight=2)

        marcos = {
            "propietario": ttk.LabelFrame(cuerpo, text=GRUPOS["propietario"].capitalize(), padding=10),
            "equipo": ttk.LabelFrame(cuerpo, text=GRUPOS["equipo"].capitalize(), padding=10),
            "otros": ttk.LabelFrame(cuerpo, text="Estado y observaciones", padding=10),
        }
        marcos["propietario"].grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        marcos["equipo"].grid(row=0, column=1, rowspan=2, sticky="nsew")
        marcos["otros"].grid(row=1, column=0, sticky="nsew", padx=(0, 8), pady=(8, 0))

        filas = {"propietario": 0, "equipo": 0, "otros": 0}
        for campo in CAMPOS:
            marco = marcos[campo.grupo]
            if campo.grupo == "equipo":
                # dos columnas de campos en el bloque de datos del ordenador
                idx = filas["equipo"]
                r, c = idx // 2, (idx % 2) * 2
                filas["equipo"] += 1
            else:
                r, c = filas[campo.grupo], 0
                filas[campo.grupo] += 1
            ttk.Label(marco, text=campo.etiqueta + ":").grid(row=r, column=c, sticky="w", padx=(0, 6), pady=3)
            marco.columnconfigure(c + 1, weight=1)
            valor = (equipo or {}).get(campo.clave, "") or ""
            if campo.tipo == "largo":
                w = tk.Text(marco, width=34, height=5, wrap="word", font=app.fuente)
                w.insert("1.0", valor)
                w.grid(row=r, column=c + 1, sticky="ew", pady=3)
            elif campo.tipo == "fecha":
                var = tk.StringVar(value=mostrar_fecha(valor))
                fr = ttk.Frame(marco)
                fr.grid(row=r, column=c + 1, sticky="ew", pady=3)
                w = ttk.Entry(fr, textvariable=var, width=12)
                w.pack(side="left", fill="x", expand=True)
                ttk.Button(fr, text="Hoy", width=4,
                           command=lambda v=var: v.set(date.today().strftime("%d/%m/%Y"))).pack(side="left", padx=(3, 0))
                if campo.clave == "fin_garantia":
                    mb = ttk.Menubutton(fr, text="+", width=2)
                    menu = tk.Menu(mb, tearoff=False)
                    for anios in (1, 2, 3, 4, 5):
                        menu.add_command(label=f"Compra + {anios} año{'s' if anios > 1 else ''}",
                                         command=lambda a=anios: self._garantia_desde_compra(a))
                    mb["menu"] = menu
                    mb.pack(side="left", padx=(3, 0))
                self.vars[campo.clave] = var
            elif campo.clave == "estado":
                var = tk.StringVar(value=valor or "Activo")
                w = ttk.Combobox(marco, textvariable=var, values=ESTADOS, width=30)
                w.grid(row=r, column=c + 1, sticky="ew", pady=3)
                self.vars[campo.clave] = var
            else:
                var = tk.StringVar(value=valor)
                w = ttk.Combobox(marco, textvariable=var, width=28,
                                 values=app.bd.valores_distintos(campo.clave))
                w.grid(row=r, column=c + 1, sticky="ew", pady=3)
                self.vars[campo.clave] = var
            self.widgets[campo.clave] = w

        if equipo and equipo.get("id"):
            info = f"ID {equipo['id']}  ·  creado {equipo.get('creado', '')}  ·  modificado {equipo.get('modificado', '')}"
            ttk.Label(cuerpo, text=info, foreground="#666").grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))

        botones = ttk.Frame(cuerpo)
        botones.grid(row=3, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(botones, text="Guardar", command=self._guardar, style="Accent.TButton").pack(side="right")
        ttk.Button(botones, text="Cancelar", command=self.destroy).pack(side="right", padx=6)

        self.bind("<Return>", lambda e: self._guardar() if not isinstance(e.widget, tk.Text) else None)
        self.bind("<Escape>", lambda e: self.destroy())
        self.widgets["persona"].focus_set()
        self.update_idletasks()
        self._centrar()
        self.grab_set()

    def _centrar(self):
        x = self.app.winfo_rootx() + (self.app.winfo_width() - self.winfo_width()) // 2
        y = self.app.winfo_rooty() + (self.app.winfo_height() - self.winfo_height()) // 3
        self.geometry(f"+{max(0, x)}+{max(0, y)}")

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
                if not fecha_a_date(iso):
                    messagebox.showerror(TITULO, f"La fecha '{d[clave]}' en «{CAMPOS_POR_CLAVE[clave].etiqueta}» "
                                                 "no es válida.\nUsa el formato DD/MM/AAAA.", parent=self)
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
        dup = self.app.bd.buscar_duplicado(d, excluir_id=(self.equipo or {}).get("id"))
        if dup:
            campo = "nº de serie" if d["num_serie"] else "nombre de equipo"
            if not messagebox.askyesno(
                    TITULO, f"Ya existe un equipo con el mismo {campo}:\n\n"
                            f"{dup.get('nombre_equipo') or '-'}  ·  {dup.get('persona') or '-'}  ·  "
                            f"{dup.get('num_serie') or '-'}\n\n¿Guardar de todos modos?", parent=self):
                return
        self.resultado = d
        self.destroy()


# =====================================================================
#  Diálogo de importación
# =====================================================================
class DialogoImportar(tk.Toplevel):
    OPCIONES = [
        ("actualizar", "Actualizar los existentes con los datos del Excel"),
        ("completar", "Completar sólo los campos vacíos de los existentes"),
        ("omitir", "Omitir los equipos que ya existen"),
        ("duplicar", "Añadirlos igualmente como equipos nuevos"),
    ]

    def __init__(self, app: "App", ruta: Path):
        super().__init__(app)
        self.app = app
        self.ruta = ruta
        self.title("Importar inventario")
        self.transient(app)
        self.geometry("980x600")
        self.minsize(700, 450)
        self.importado = False

        try:
            self.res = leer_excel(ruta)
        except Exception as exc:  # noqa: BLE001
            self.destroy()
            messagebox.showerror(TITULO, f"No se pudo leer el archivo:\n{ruta}\n\n{exc}", parent=app)
            raise _Cancelado from exc

        if not self.res.filas:
            self.destroy()
            avisos = "\n".join(self.res.avisos) or "El archivo no contiene filas con datos."
            messagebox.showwarning(TITULO, f"No hay equipos que importar.\n\n{avisos}", parent=app)
            raise _Cancelado

        # clasificar filas en nuevas / existentes
        self.existentes: dict[int, dict] = {}
        for i, fila in enumerate(self.res.filas):
            dup = app.bd.buscar_duplicado(fila)
            if dup:
                self.existentes[i] = dup
        # duplicados dentro del propio archivo
        vistos, internos = set(), 0
        for fila in self.res.filas:
            k = normalizar(fila.get("num_serie")) or normalizar(fila.get("nombre_equipo"))
            if k:
                internos += k in vistos
                vistos.add(k)

        cont = ttk.Frame(self, padding=12)
        cont.pack(fill="both", expand=True)

        resumen = (f"Archivo: {ruta.name}\n"
                   f"Hojas leídas: {', '.join(self.res.hojas_leidas)}\n"
                   f"Filas con datos: {len(self.res.filas)}   ·   Nuevos: {len(self.res.filas) - len(self.existentes)}"
                   f"   ·   Ya existentes (mismo nº de serie o nombre): {len(self.existentes)}")
        if internos:
            resumen += f"\nAtención: {internos} fila(s) repetidas dentro del propio archivo."
        ttk.Label(cont, text=resumen, justify="left").pack(anchor="w")
        for hoja, cols in self.res.columnas_reconocidas.items():
            ttk.Label(cont, text=f"Columnas reconocidas en «{hoja}»: {', '.join(cols)}",
                      foreground="#2E7D32", wraplength=940, justify="left").pack(anchor="w", pady=(4, 0))
        for hoja, cols in self.res.columnas_ignoradas.items():
            ttk.Label(cont, text=f"Columnas no reconocidas (se ignoran) en «{hoja}»: {', '.join(cols)}",
                      foreground="#B26A00", wraplength=940, justify="left").pack(anchor="w")
        for aviso in self.res.avisos:
            ttk.Label(cont, text=aviso, foreground="#B26A00").pack(anchor="w")

        # vista previa
        marco = ttk.LabelFrame(cont, text="Vista previa", padding=4)
        marco.pack(fill="both", expand=True, pady=8)
        cols = ["_estado"] + CLAVES
        tv = ttk.Treeview(marco, columns=cols, show="headings", height=10)
        tv.heading("_estado", text="Importación")
        tv.column("_estado", width=90, stretch=False)
        for c in CAMPOS:
            tv.heading(c.clave, text=c.etiqueta)
            tv.column(c.clave, width=c.ancho, stretch=False)
        tv.tag_configure("existe", background="#FFF4D6")
        for i, fila in enumerate(self.res.filas[:1000]):
            existe = i in self.existentes
            valores = ["Existe" if existe else "Nuevo"]
            valores += [mostrar_fecha(fila[c.clave]) if c.tipo == "fecha" else fila[c.clave] for c in CAMPOS]
            tv.insert("", "end", values=valores, tags=("existe",) if existe else ())
        ys = ttk.Scrollbar(marco, orient="vertical", command=tv.yview)
        xs = ttk.Scrollbar(marco, orient="horizontal", command=tv.xview)
        tv.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        tv.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        marco.rowconfigure(0, weight=1)
        marco.columnconfigure(0, weight=1)

        self.modo = tk.StringVar(value="actualizar")
        if self.existentes:
            op = ttk.LabelFrame(cont, text="¿Qué hacer con los equipos que ya existen?", padding=6)
            op.pack(fill="x")
            for valor, texto in self.OPCIONES:
                ttk.Radiobutton(op, text=texto, value=valor, variable=self.modo).pack(anchor="w")

        bot = ttk.Frame(cont)
        bot.pack(fill="x", pady=(8, 0))
        ttk.Button(bot, text=f"Importar {len(self.res.filas)} fila(s)", style="Accent.TButton",
                   command=self._importar).pack(side="right")
        ttk.Button(bot, text="Cancelar", command=self.destroy).pack(side="right", padx=6)
        self.bind("<Escape>", lambda e: self.destroy())
        self.grab_set()

    def _importar(self):
        bd, modo = self.app.bd, self.modo.get()
        nuevos = actualizados = omitidos = 0
        try:
            for i, fila in enumerate(self.res.filas):
                datos = {k: v for k, v in fila.items() if not k.startswith("_")}
                existente = self.existentes.get(i)
                if existente is None:
                    # puede haberse creado en esta misma importación (fila repetida en el archivo)
                    existente = bd.buscar_duplicado(datos) if modo != "duplicar" else None
                if existente is None or modo == "duplicar":
                    bd.insertar(datos, commit=False)
                    nuevos += 1
                elif modo == "omitir":
                    omitidos += 1
                else:
                    if modo == "actualizar":
                        # no borrar datos existentes con celdas vacías del Excel
                        datos = {k: (v if v else existente.get(k, "")) for k, v in datos.items()}
                        cambiado = bd.actualizar(existente["id"], datos, commit=False)
                    else:  # completar
                        datos = {k: (existente.get(k) or v) for k, v in datos.items()}
                        cambiado = bd.actualizar(existente["id"], datos, commit=False)
                    if cambiado:
                        actualizados += 1
                    else:
                        omitidos += 1
            bd.registrar(None, self.ruta.name, "Importación",
                         f"{nuevos} nuevos, {actualizados} actualizados, {omitidos} sin cambios")
            bd.commit()
        except Exception as exc:  # noqa: BLE001
            bd.con.rollback()
            messagebox.showerror(TITULO, f"Error durante la importación. No se ha guardado nada.\n\n{exc}", parent=self)
            return
        self.importado = True
        self.destroy()
        messagebox.showinfo(TITULO, f"Importación completada.\n\nNuevos: {nuevos}\nActualizados: {actualizados}\n"
                                    f"Sin cambios / omitidos: {omitidos}", parent=self.app)


class _Cancelado(Exception):
    pass


# =====================================================================
#  Ventana principal
# =====================================================================
class App(tk.Tk):
    def __init__(self, ruta_bd: str | None = None):
        super().__init__()
        self.title(TITULO)
        self.geometry("1400x780")
        self.minsize(900, 500)
        self._configurar_estilo()
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

        self._menu()
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=6, pady=(6, 0))
        self._pestana_inventario()
        self._pestana_resumen()
        self._pestana_historial()
        self.nb.bind("<<NotebookTabChanged>>", self._cambio_pestana)

        self.estado = tk.StringVar()
        ttk.Label(self, textvariable=self.estado, anchor="w", padding=(8, 3), relief="sunken").pack(fill="x", side="bottom")

        self._atajos()
        self.recargar()
        self.protocol("WM_DELETE_WINDOW", self._salir)
        self.after(400, self._aviso_garantias)

    # ------------------------------------------------------------ apariencia
    def _configurar_estilo(self):
        estilo = ttk.Style(self)
        temas = estilo.theme_names()
        for t in ("vista", "winnative", "clam"):
            if t in temas:
                estilo.theme_use(t)
                break
        self.fuente = ("Segoe UI", 10) if os.name == "nt" else ("TkDefaultFont", 10)
        estilo.configure("Treeview", rowheight=24)
        estilo.configure("Treeview.Heading", font=(self.fuente[0], 10, "bold"))
        estilo.configure("Accent.TButton", font=(self.fuente[0], 10, "bold"))
        estilo.configure("Titulo.TLabel", font=(self.fuente[0], 13, "bold"))
        estilo.configure("Cifra.TLabel", font=(self.fuente[0], 22, "bold"))

    def _icono(self):
        try:
            ico = recurso("icono.ico")
            png = recurso("icono.png")
            if os.name == "nt" and ico.exists():
                self.iconbitmap(default=str(ico))
            elif png.exists():
                self._img_icono = tk.PhotoImage(file=str(png))
                self.iconphoto(True, self._img_icono)
        except tk.TclError:
            pass

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
        e.add_command(label="Cambiar delegación de los seleccionados…",
                      command=lambda: self.cambiar_campo_dialogo("delegacion"))
        e.add_separator()
        e.add_command(label="Buscar", accelerator="Ctrl+F", command=lambda: self.ent_buscar.focus_set())
        e.add_command(label="Seleccionar todo", accelerator="Ctrl+A", command=self.seleccionar_todo)
        barra.add_cascade(label="Equipos", menu=e)

        v = tk.Menu(barra, tearoff=False)
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
        self.menu_contexto.add_cascade(label="Cambiar estado", menu=self.menu_estado)
        self.menu_contexto.add_command(label="Copiar fila", command=self.copiar_fila)
        self.menu_contexto.add_separator()
        self.menu_contexto.add_command(label="Exportar seleccionados…", command=lambda: self.exportar("excel", "seleccion"))
        self.menu_contexto.add_command(label="Eliminar", command=self.eliminar)

    def _atajos(self):
        self.bind_all("<Control-n>", lambda e: self.nuevo())
        self.bind_all("<Control-i>", lambda e: self.importar())
        self.bind_all("<Control-e>", lambda e: self.exportar("excel", "vista"))
        self.bind_all("<Control-f>", lambda e: (self.nb.select(0), self.ent_buscar.focus_set(), "break")[-1])
        self.bind_all("<F5>", lambda e: self.recargar())
        self.tabla.bind("<Control-d>", lambda e: self.duplicar())
        self.tabla.bind("<Control-a>", lambda e: (self.seleccionar_todo(), "break")[-1])
        self.tabla.bind("<Control-c>", lambda e: self.copiar_fila())
        self.tabla.bind("<Delete>", lambda e: self.eliminar())
        self.tabla.bind("<Return>", lambda e: self.editar())
        self.tabla.bind("<Double-1>", self._doble_click)
        self.tabla.bind("<Button-3>", self._click_derecho)
        self.tabla.bind("<<TreeviewSelect>>", lambda e: self._actualizar_estado())

    # ------------------------------------------------------------ pestaña inventario
    def _pestana_inventario(self):
        p = ttk.Frame(self.nb, padding=6)
        self.nb.add(p, text="  Inventario  ")

        barra = ttk.Frame(p)
        barra.pack(fill="x")
        for texto, cmd in (("➕ Nuevo", self.nuevo), ("✏ Editar", self.editar), ("⧉ Duplicar", self.duplicar),
                           ("🗑 Eliminar", self.eliminar)):
            ttk.Button(barra, text=texto, command=cmd).pack(side="left", padx=(0, 4))
        ttk.Separator(barra, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Button(barra, text="📥 Importar Excel", command=self.importar).pack(side="left", padx=(0, 4))
        ttk.Button(barra, text="📤 Exportar Excel", command=lambda: self.exportar("excel", "vista")).pack(side="left")

        filtros = ttk.Frame(p)
        filtros.pack(fill="x", pady=(8, 6))
        ttk.Label(filtros, text="Buscar:").pack(side="left")
        self.var_buscar = tk.StringVar()
        self.ent_buscar = ttk.Entry(filtros, textvariable=self.var_buscar, width=26)
        self.ent_buscar.pack(side="left", padx=(4, 12))
        self.var_buscar.trace_add("write", lambda *a: self._programar_filtro())
        self.ent_buscar.bind("<Escape>", lambda e: self.var_buscar.set(""))

        self.filtros: dict[str, tk.StringVar] = {}
        self.combos: dict[str, ttk.Combobox] = {}
        for clave, etiqueta in (("delegacion", "Delegación"), ("marca", "Marca"), ("os", "SO"), ("estado", "Estado")):
            ttk.Label(filtros, text=etiqueta + ":").pack(side="left")
            var = tk.StringVar(value=TODOS)
            cb = ttk.Combobox(filtros, textvariable=var, state="readonly", width=13)
            cb.pack(side="left", padx=(4, 8))
            cb.bind("<<ComboboxSelected>>", lambda e: self.aplicar_filtro())
            self.filtros[clave], self.combos[clave] = var, cb
        ttk.Label(filtros, text="Garantía:").pack(side="left")
        self.var_garantia = tk.StringVar(value=TODOS)
        cb = ttk.Combobox(filtros, textvariable=self.var_garantia, values=FILTROS_GARANTIA, state="readonly", width=17)
        cb.pack(side="left", padx=(4, 10))
        cb.bind("<<ComboboxSelected>>", lambda e: self.aplicar_filtro())
        ttk.Button(filtros, text="Limpiar filtros", command=self.limpiar_filtros).pack(side="left")

        marco = ttk.Frame(p)
        marco.pack(fill="both", expand=True)
        cols = ["id"] + CLAVES
        self.tabla = ttk.Treeview(marco, columns=cols, show="headings", selectmode="extended")
        self.tabla.heading("id", text="ID", command=lambda: self.ordenar("id"))
        self.tabla.column("id", width=50, anchor="e", stretch=False)
        for c in CAMPOS:
            self.tabla.heading(c.clave, text=c.etiqueta, command=lambda k=c.clave: self.ordenar(k))
            self.tabla.column(c.clave, width=c.ancho, stretch=False,
                              anchor="center" if c.tipo == "fecha" else "w")
        self.tabla.tag_configure("caducada", background="#F9D3CF")
        self.tabla.tag_configure("proxima", background="#FFF1C2")
        self.tabla.tag_configure("baja", foreground="#8A8A8A")
        ys = ttk.Scrollbar(marco, orient="vertical", command=self.tabla.yview)
        xs = ttk.Scrollbar(marco, orient="horizontal", command=self.tabla.xview)
        self.tabla.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tabla.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        marco.rowconfigure(0, weight=1)
        marco.columnconfigure(0, weight=1)
        self._aplicar_columnas_visibles()

        leyenda = ttk.Frame(p)
        leyenda.pack(fill="x", pady=(4, 0))
        for color, texto in (("#F9D3CF", "Garantía caducada"),
                             ("#FFF1C2", f"Garantía caduca en ≤ {DIAS_AVISO_GARANTIA} días")):
            tk.Label(leyenda, bg=color, width=2, relief="solid", bd=1).pack(side="left")
            ttk.Label(leyenda, text=texto).pack(side="left", padx=(4, 14))
        ttk.Label(leyenda, text="Doble clic para editar · clic derecho para más opciones · "
                                "clic en la cabecera para ordenar", foreground="#666").pack(side="right")
        self._id_filtro = None

    # ------------------------------------------------------------ pestaña resumen
    def _pestana_resumen(self):
        p = ttk.Frame(self.nb, padding=10)
        self.nb.add(p, text="  Resumen  ")
        self.pest_resumen = p
        self.tarjetas = ttk.Frame(p)
        self.tarjetas.pack(fill="x")
        self.cifras: dict[str, tk.StringVar] = {}
        for clave, texto in (("total", "Equipos"), ("activos", "Activos"), ("caducadas", "Garantía caducada"),
                             ("proximas", f"Caduca ≤ {DIAS_AVISO_GARANTIA} días"), ("sin_garantia", "Sin fecha garantía"),
                             ("personas", "Personas")):
            f = ttk.Frame(self.tarjetas, padding=10, relief="groove")
            f.pack(side="left", padx=(0, 10), fill="x", expand=True)
            var = tk.StringVar(value="0")
            ttk.Label(f, textvariable=var, style="Cifra.TLabel").pack()
            ttk.Label(f, text=texto).pack()
            self.cifras[clave] = var

        rejilla = ttk.Frame(p)
        rejilla.pack(fill="both", expand=True, pady=(12, 0))
        self.tablas_resumen: dict[str, ttk.Treeview] = {}
        grupos = [("delegacion", "Por delegación/CT"), ("marca", "Por marca"), ("os", "Por sistema operativo"),
                  ("estado", "Por estado"), ("office", "Por Office"), ("antivirus", "Por antivirus"),
                  ("ram", "Por RAM"), ("anio", "Por año de compra")]
        for i, (clave, titulo) in enumerate(grupos):
            lf = ttk.LabelFrame(rejilla, text=titulo, padding=4)
            lf.grid(row=i // 4, column=i % 4, sticky="nsew", padx=4, pady=4)
            tv = ttk.Treeview(lf, columns=("valor", "n", "barra"), show="headings", height=7)
            tv.heading("valor", text="Valor")
            tv.heading("n", text="Nº")
            tv.heading("barra", text="")
            tv.column("valor", width=140)
            tv.column("n", width=45, anchor="e", stretch=False)
            tv.column("barra", width=110)
            sb = ttk.Scrollbar(lf, orient="vertical", command=tv.yview)
            tv.configure(yscrollcommand=sb.set)
            tv.pack(side="left", fill="both", expand=True)
            sb.pack(side="left", fill="y")
            tv.bind("<Double-1>", lambda e, k=clave, t=tv: self._filtrar_desde_resumen(k, t))
            self.tablas_resumen[clave] = tv
        for c in range(4):
            rejilla.columnconfigure(c, weight=1)
        for r in range(2):
            rejilla.rowconfigure(r, weight=1)

        lf = ttk.LabelFrame(p, text="Garantías caducadas o próximas a caducar", padding=4)
        lf.pack(fill="both", expand=True, pady=(8, 0))
        cols = ("fin_garantia", "dias", "nombre_equipo", "persona", "delegacion", "marca", "modelo", "num_serie")
        self.tabla_garantias = ttk.Treeview(lf, columns=cols, show="headings", height=6)
        for c in cols:
            texto = "Días" if c == "dias" else CAMPOS_POR_CLAVE[c].etiqueta
            self.tabla_garantias.heading(c, text=texto)
            self.tabla_garantias.column(c, width=70 if c == "dias" else 140, anchor="e" if c == "dias" else "w")
        self.tabla_garantias.tag_configure("caducada", background="#F9D3CF")
        self.tabla_garantias.tag_configure("proxima", background="#FFF1C2")
        sb = ttk.Scrollbar(lf, orient="vertical", command=self.tabla_garantias.yview)
        self.tabla_garantias.configure(yscrollcommand=sb.set)
        self.tabla_garantias.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        self.tabla_garantias.bind("<Double-1>", self._editar_desde_garantias)
        ttk.Label(p, text="Doble clic en una fila de las tablas para filtrar el inventario o abrir el equipo.",
                  foreground="#666").pack(anchor="w", pady=(4, 0))

    # ------------------------------------------------------------ pestaña historial
    def _pestana_historial(self):
        p = ttk.Frame(self.nb, padding=6)
        self.nb.add(p, text="  Historial  ")
        barra = ttk.Frame(p)
        barra.pack(fill="x", pady=(0, 6))
        ttk.Button(barra, text="Exportar historial a Excel…", command=self.exportar_historial).pack(side="left")
        ttk.Button(barra, text="Vaciar historial", command=self.vaciar_historial).pack(side="left", padx=6)
        cols = ("fecha", "accion", "equipo", "detalle")
        self.tabla_hist = ttk.Treeview(p, columns=cols, show="headings")
        for c, t, w in (("fecha", "Fecha", 140), ("accion", "Acción", 120), ("equipo", "Equipo", 260),
                        ("detalle", "Detalle", 800)):
            self.tabla_hist.heading(c, text=t)
            self.tabla_hist.column(c, width=w, stretch=c == "detalle")
        sb = ttk.Scrollbar(p, orient="vertical", command=self.tabla_hist.yview)
        self.tabla_hist.configure(yscrollcommand=sb.set)
        self.tabla_hist.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")

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
            cb["values"] = [TODOS] + valores
            if self.filtros[clave].get() not in cb["values"]:
                self.filtros[clave].set(TODOS)
        self.aplicar_filtro()
        self._cambio_pestana()

    def _programar_filtro(self):
        if self._id_filtro:
            self.after_cancel(self._id_filtro)
        self._id_filtro = self.after(200, self.aplicar_filtro)

    def aplicar_filtro(self):
        self._id_filtro = None
        tokens = [normalizar(t) for t in self.var_buscar.get().split() if normalizar(t)]
        fil = {k: v.get() for k, v in self.filtros.items() if v.get() != TODOS}
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
        self._pintar_tabla()

    def _pintar_tabla(self):
        seleccion = set(self.tabla.selection())
        self.tabla.delete(*self.tabla.get_children())
        for eq in self.visibles:
            valores = [eq["id"]] + [mostrar_fecha(eq[c.clave]) if c.tipo == "fecha"
                                    else eq[c.clave].replace("\n", " ") for c in CAMPOS]
            tags = []
            if eq["_garantia"] in ("caducada", "proxima") and eq.get("estado") != "Baja":
                tags.append(eq["_garantia"])
            if eq.get("estado") == "Baja":
                tags.append("baja")
            iid = str(eq["id"])
            self.tabla.insert("", "end", iid=iid, values=valores, tags=tags)
        restaurar = [i for i in seleccion if self.tabla.exists(i)]
        if restaurar:
            self.tabla.selection_set(restaurar)
        for c in CAMPOS + [None]:
            clave = c.clave if c else "id"
            texto = c.etiqueta if c else "ID"
            if self.orden[0] == clave:
                texto += " ▼" if self.orden[1] else " ▲"
            self.tabla.heading(clave, text=texto)
        self._actualizar_estado()

    def _actualizar_estado(self):
        sel = len(self.tabla.selection())
        txt = f"{len(self.visibles)} de {len(self.equipos)} equipos mostrados"
        if sel:
            txt += f"   ·   {sel} seleccionado(s)"
        txt += f"   ·   Base de datos: {self.bd.ruta}"
        self.estado.set(txt)

    def ordenar(self, clave: str):
        actual, desc = self.orden
        self.orden = (clave, not desc if actual == clave else False)
        self.aplicar_filtro()

    def limpiar_filtros(self):
        for v in self.filtros.values():
            v.set(TODOS)
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
        f = FormularioEquipo(self, eq, f"Editar equipo · {eq.get('nombre_equipo') or eq.get('persona') or eq['id']}")
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

    def cambiar_campo_dialogo(self, clave: str):
        sel = self.seleccionados()
        if not sel:
            messagebox.showinfo(TITULO, "Selecciona uno o varios equipos.")
            return
        top = tk.Toplevel(self)
        top.title(f"Cambiar {CAMPOS_POR_CLAVE[clave].etiqueta}")
        top.transient(self)
        fr = ttk.Frame(top, padding=12)
        fr.pack(fill="both", expand=True)
        ttk.Label(fr, text=f"Nuevo valor para {len(sel)} equipo(s):").pack(anchor="w")
        var = tk.StringVar()
        cb = ttk.Combobox(fr, textvariable=var, values=self.bd.valores_distintos(clave), width=35)
        cb.pack(fill="x", pady=6)
        cb.focus_set()

        def aceptar():
            top.destroy()
            self.cambiar_campo(clave, var.get().strip())
        ttk.Button(fr, text="Aceptar", command=aceptar).pack(side="right")
        ttk.Button(fr, text="Cancelar", command=top.destroy).pack(side="right", padx=6)
        top.bind("<Return>", lambda e: aceptar())
        top.bind("<Escape>", lambda e: top.destroy())
        top.grab_set()

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
        self.estado.set(f"{len(sel)} fila(s) copiadas al portapapeles (se pueden pegar en Excel).")

    def _seleccionar(self, equipo_id: int):
        iid = str(equipo_id)
        if self.tabla.exists(iid):
            self.tabla.selection_set(iid)
            self.tabla.see(iid)
            self.tabla.focus(iid)

    def _doble_click(self, evento):
        if self.tabla.identify_region(evento.x, evento.y) == "cell":
            fila = self.tabla.identify_row(evento.y)
            if fila:
                self.editar(int(fila))

    def _click_derecho(self, evento):
        fila = self.tabla.identify_row(evento.y)
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
        visibles = ["id"] + [c for c in CLAVES if c not in self.columnas_ocultas]
        self.tabla["displaycolumns"] = visibles

    def dialogo_columnas(self):
        top = tk.Toplevel(self)
        top.title("Columnas visibles")
        top.transient(self)
        fr = ttk.Frame(top, padding=12)
        fr.pack(fill="both", expand=True)
        variables = {}
        for i, c in enumerate(CAMPOS):
            v = tk.BooleanVar(value=c.clave not in self.columnas_ocultas)
            ttk.Checkbutton(fr, text=c.etiqueta, variable=v).grid(row=i % 10, column=i // 10, sticky="w", padx=6)
            variables[c.clave] = v

        def aceptar():
            self.columnas_ocultas = {k for k, v in variables.items() if not v.get()}
            cfg = leer_config()
            cfg["columnas_ocultas"] = sorted(self.columnas_ocultas)
            guardar_config(cfg)
            self._aplicar_columnas_visibles()
            top.destroy()
        b = ttk.Frame(fr)
        b.grid(row=11, column=0, columnspan=3, sticky="e", pady=(10, 0))
        ttk.Button(b, text="Aceptar", command=aceptar).pack(side="right")
        ttk.Button(b, text="Cancelar", command=top.destroy).pack(side="right", padx=6)
        top.grab_set()

    # ------------------------------------------------------------ resumen / historial
    def _cambio_pestana(self, _evento=None):
        try:
            actual = self.nb.index(self.nb.select())
        except tk.TclError:
            return
        if actual == 1:
            self._pintar_resumen()
        elif actual == 2:
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
        self.cifras["personas"].set(str(len({normalizar(e["persona"]) for e in eqs if e.get("persona")})))

        for clave, tv in self.tablas_resumen.items():
            tv.delete(*tv.get_children())
            if clave == "anio":
                cont = Counter((fecha_a_date(e["fecha_compra"]).year if fecha_a_date(e["fecha_compra"]) else "(sin dato)")
                               for e in eqs)
                items = sorted(cont.items(), key=lambda x: str(x[0]), reverse=True)
            else:
                cont = Counter((e.get(clave) or "(sin dato)") for e in eqs)
                items = cont.most_common()
            maximo = max(cont.values(), default=1)
            for valor, n in items:
                barra = "█" * max(1, round(12 * n / maximo))
                tv.insert("", "end", values=(valor, n, barra))

        tg = self.tabla_garantias
        tg.delete(*tg.get_children())
        criticos = [e for e in no_baja if e["_garantia"] in ("caducada", "proxima")]
        criticos.sort(key=lambda e: e["fin_garantia"])
        for e in criticos:
            dias = (fecha_a_date(e["fin_garantia"]) - hoy).days
            tg.insert("", "end", iid=str(e["id"]), tags=(e["_garantia"],),
                      values=(mostrar_fecha(e["fin_garantia"]), dias, e["nombre_equipo"], e["persona"],
                              e["delegacion"], e["marca"], e["modelo"], e["num_serie"]))

    def _filtrar_desde_resumen(self, clave: str, tv: ttk.Treeview):
        sel = tv.selection()
        if not sel or clave not in self.filtros:
            return
        valor = tv.item(sel[0], "values")[0]
        if valor == "(sin dato)":
            return
        self.limpiar_filtros()
        self.filtros[clave].set(valor)
        self.nb.select(0)
        self.aplicar_filtro()

    def _editar_desde_garantias(self, _e):
        sel = self.tabla_garantias.selection()
        if sel:
            self.editar(int(sel[0]))
            self._pintar_resumen()

    def _pintar_historial(self):
        self.tabla_hist.delete(*self.tabla_hist.get_children())
        for h in self.bd.historial():
            self.tabla_hist.insert("", "end", values=(h["fecha"], h["accion"], h["equipo"],
                                                      h["detalle"].replace("\n", " ")))

    def _aviso_garantias(self):
        if not leer_config().get("avisar_garantias", True):
            return
        hoy = date.today().isoformat()
        cfg = leer_config()
        if cfg.get("ultimo_aviso") == hoy:
            return
        prox = [e for e in self.equipos if e["_garantia"] == "proxima" and e.get("estado") != "Baja"]
        if prox:
            cfg["ultimo_aviso"] = hoy
            guardar_config(cfg)
            if messagebox.askyesno(TITULO, f"Hay {len(prox)} equipo(s) cuya garantía caduca en los próximos "
                                           f"{DIAS_AVISO_GARANTIA} días.\n\n¿Ver el listado?"):
                self.nb.select(1)

    # ------------------------------------------------------------ ayuda
    def ayuda_formato(self):
        txt = ("El programa importa archivos .xlsx (y .csv) con el formato de la plantilla:\n\n"
               "• Filas 1-2: títulos de grupo (PROPIETARIO DEL ORDENADOR / DATOS DEL ORDENADOR).\n"
               "• Fila 3 (o cualquier fila de las 30 primeras): cabeceras de columna.\n"
               "• Datos a partir de la fila siguiente a las cabeceras.\n\n"
               "Columnas reconocidas (da igual el orden, mayúsculas o tildes):\n" +
               ", ".join(c.cabecera for c in CAMPOS) +
               "\n\nLos equipos se identifican por el Nº DE SERIE (o el nombre de equipo si no hay nº de serie) "
               "para no duplicarlos al volver a importar.\n\n"
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
        app.nb.select(1)
        app.update()
        app.after(500, app._salir)
        app.mainloop()
        informe.write_text("OK\n", encoding="utf-8")
        return 0
    except Exception as exc:  # noqa: BLE001
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
