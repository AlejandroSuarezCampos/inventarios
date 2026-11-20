"""Ventana principal del inventario de equipos (Tkinter)."""
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
from .campos import (CAMPOS, CAMPOS_POR_CLAVE, CLAVES, DIAS_AVISO_GARANTIA, ESTADOS, dias_aviso,
                     estado_garantia, fecha_a_date, fijar_dias_aviso, mostrar_fecha, normalizar, parsear_fecha)
from .db import BaseDatos, Conflicto, guardar_config, leer_config, usuario_actual
from .dialogos import (Cancelado, DialogoCampoMasivo, DialogoColumnas, DialogoEtiquetas, DialogoImportar,
                       DialogoPreferencias, DialogoTexto, FormularioEquipo)
from .excel_io import exportar_csv, exportar_excel, exportar_plantilla, leer_excel, leer_texto_tabla
from .informes import generar_informe, id_desde_qr, periodos_garantia
from .tema import (C, MODOS, PALETA_GRAFICOS, CampoBusqueda, GraficoBarras, Tema, aplicar_modo,
                   barra_titulo_oscura)
from .ui_comun import (TITULO, abrir_archivo, foco_en_texto, recurso, tabla_con_scroll, tarjeta, texto_fila)

TODOS = "Garantía"
# Primer valor de cada filtro = «sin filtrar»
TODOS_FILTRO = {"delegacion": "Delegaciones", "marca": "Marcas", "os": "Sistemas", "estado": "Estados"}
FILTROS_GARANTIA = [TODOS, "Vigente", "Caduca pronto", "Caducada", "Sin fecha"]
_MAPA_GARANTIA = {FILTROS_GARANTIA[1]: "vigente", FILTROS_GARANTIA[2]: "proxima",
                  FILTROS_GARANTIA[3]: "caducada", FILTROS_GARANTIA[4]: "sin_dato"}

# Compatibilidad con código anterior
_Cancelado = Cancelado


class App(tk.Tk):
    PAGINAS = (("Inventario", "▦"), ("Resumen", "◔"), ("Historial", "↺"))

    def __init__(self, ruta_bd: str | None = None):
        super().__init__()
        self.cfg = leer_config()
        fijar_dias_aviso(self.cfg.get("dias_aviso", DIAS_AVISO_GARANTIA))
        self.oscuro = aplicar_modo(self.cfg.get("tema", MODOS[0]))
        self.title(TITULO)
        self.tema = Tema(self)
        px = self.tema.px
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        ancho, alto = min(px(1440), int(sw * 0.94)), min(px(860), int(sh * 0.88))
        self.geometry(f"{ancho}x{alto}+{(sw - ancho) // 2}+{max(0, (sh - alto) // 3)}")
        self.minsize(min(px(860), sw - 40), min(px(560), sh - 80))
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
        self.orden = tuple(self.cfg.get("orden_tabla", ("id", False)))
        self.columnas_ocultas: set[str] = set(self.cfg.get("columnas_ocultas", []))
        self.pagina_actual = 0
        self.estado = tk.StringVar()
        self._editor = None
        self._fila_contexto: str | None = None
        self._clave_contexto: str | None = None
        self._arrastre: dict | None = None
        self._compacto = None
        self._estrecho = None
        self._cambios_pendientes = False

        self._menu()
        self._lateral()
        self.contenido = ttk.Frame(self)
        self.contenido.pack(side="left", fill="both", expand=True)
        self.paginas: list[ttk.Frame] = []
        self._pagina_inventario()
        self._pagina_resumen()
        self._pagina_historial()

        self._atajos()
        self._actualizar_textos_aviso()
        self.mostrar(0)
        self.recargar()
        barra_titulo_oscura(self, self.oscuro)
        self._version_bd = self.bd.version_datos()
        self.protocol("WM_DELETE_WINDOW", self._salir)
        self.bind("<Configure>", self._al_redimensionar, add="+")
        self.after(300, self._ajustar_disposicion)
        self.after(400, self._aviso_garantias)
        self.after(2000, self._vigilar_cambios)

    # ------------------------------------------------------------ configuración
    def guardar_cfg(self):
        guardar_config(self.cfg)

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

    # ------------------------------------------------------------ barra lateral
    def _lateral(self):
        px = self.tema.px
        lat = ttk.Frame(self, style="Lateral.TFrame", padding=(px(12), px(18), px(12), px(14)), width=px(214))
        lat.pack(side="left", fill="y")
        lat.pack_propagate(False)
        self.lateral = lat
        tk.Frame(self, width=1, bg=C["borde"]).pack(side="left", fill="y")
        self._lat_ocultables: list[tk.Widget] = []

        cab = ttk.Frame(lat, style="Lateral.TFrame")
        cab.pack(fill="x", padx=px(6), pady=(0, px(20)))
        if getattr(self, "_img_logo", None):
            tk.Label(cab, image=self._img_logo, bg=C["lateral"], bd=0).pack(side="left", padx=(0, px(10)))
        tx = ttk.Frame(cab, style="Lateral.TFrame")
        tx.pack(side="left")
        ttk.Label(tx, text="Inventario", style="LateralTitulo.TLabel").pack(anchor="w")
        ttk.Label(tx, text="Equipos informáticos", style="Lateral.TLabel").pack(anchor="w")
        self._lat_ocultables.append(tx)

        self._nav: list[tuple[ttk.Button, str, str]] = []

        def seccion(texto, pady=(0, px(4))):
            lbl = ttk.Label(lat, text=texto, style="LateralSeccion.TLabel")
            lbl.pack(anchor="w", padx=px(10), pady=pady)
            self._lat_ocultables.append(lbl)
            return lbl

        def boton(icono, texto, cmd, padre=lat):
            b = ttk.Button(padre, text=f"{icono}   {texto}", style="Lateral.TButton", command=cmd)
            b.pack(fill="x", pady=1)
            self._nav.append((b, icono, texto))
            return b

        seccion("GENERAL")
        self.botones_nav = [boton(icono, nombre, lambda i=i: self.mostrar(i))
                            for i, (nombre, icono) in enumerate(self.PAGINAS)]

        self.lbl_vistas = seccion("VISTAS GUARDADAS", (px(16), px(4)))
        self.marco_vistas = ttk.Frame(lat, style="Lateral.TFrame")
        self.marco_vistas.pack(fill="x")
        self._lat_ocultables.append(self.marco_vistas)

        seccion("DATOS", (px(16), px(4)))
        boton("⌖", "Añadir este equipo", lambda: self.nuevo(detectar=True))
        boton("⤓", "Importar Excel", self.importar)
        boton("⤒", "Exportar Excel", lambda: self.exportar("excel", "vista"))
        boton("▤", "Informe PDF", lambda: self.informe("vista"))

        pie = ttk.Frame(lat, style="Lateral.TFrame")
        pie.pack(side="bottom", fill="x", padx=px(10))
        self.lbl_lateral_total = ttk.Label(pie, text="", style="Lateral.TLabel")
        self.lbl_lateral_total.pack(anchor="w")
        self.lbl_lateral_bd = ttk.Label(pie, text="", style="Lateral.TLabel", wraplength=px(180))
        self.lbl_lateral_bd.pack(anchor="w", pady=(2, 0))
        self.lbl_lateral_usuario = ttk.Label(pie, text="", style="Lateral.TLabel", wraplength=px(180))
        self.lbl_lateral_usuario.pack(anchor="w", pady=(2, 0))
        ttk.Label(pie, text=f"Versión {__version__}", style="Lateral.TLabel").pack(anchor="w", pady=(px(6), 0))
        self._lat_ocultables.append(pie)
        self._pintar_vistas()
        self._orden_lateral = [(w, w.pack_info()) for w in lat.pack_slaves()]

    def _modo_compacto(self, compacto: bool):
        """Barra lateral sólo con iconos en pantallas estrechas."""
        if compacto == self._compacto:
            return
        self._compacto = compacto
        px = self.tema.px
        self.lateral.configure(width=px(62) if compacto else px(214),
                               padding=(px(8), px(18), px(8), px(14)) if compacto else (px(12), px(18), px(12), px(14)))
        for b, icono, texto in self._nav:
            if b.winfo_exists():
                b.configure(text=icono if compacto else f"{icono}   {texto}")
        if not hasattr(self, "_orden_lateral"):
            self._orden_lateral = [(w, w.pack_info()) for w in self.lateral.pack_slaves()]
        if compacto:
            for w in self._lat_ocultables:
                if w.winfo_manager() == "pack":
                    w.pack_forget()
        else:  # volver a colocar todo en el orden original
            for w, _ in self._orden_lateral:
                w.pack_forget()
            for w, info in self._orden_lateral:
                info = dict(info)
                info.pop("in", None)
                w.pack(**info)

    def mostrar(self, indice: int):
        self._cerrar_editor(True)
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
        cab.pack(fill="x", pady=(0, px(14)))
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
        m.add_command(label="Pegar filas copiadas de Excel", accelerator="Ctrl+V", command=self.pegar)
        m.add_separator()
        m.add_command(label="Exportar a Excel (vista actual)…", accelerator="Ctrl+E",
                      command=lambda: self.exportar("excel", "vista"))
        m.add_command(label="Exportar a Excel (todo)…", command=lambda: self.exportar("excel", "todo"))
        m.add_command(label="Exportar a Excel (seleccionados)…", command=lambda: self.exportar("excel", "seleccion"))
        m.add_command(label="Exportar a CSV (vista actual)…", command=lambda: self.exportar("csv", "vista"))
        m.add_command(label="Guardar plantilla Excel vacía…", command=self.plantilla)
        m.add_separator()
        m.add_command(label="Informe PDF (vista actual)…", accelerator="Ctrl+P", command=lambda: self.informe("vista"))
        m.add_command(label="Informe PDF (todo el inventario)…", command=lambda: self.informe("todo"))
        m.add_command(label="Etiquetas con QR de los seleccionados…", command=lambda: self.etiquetas("seleccion"))
        m.add_command(label="Etiquetas con QR de la vista actual…", command=lambda: self.etiquetas("vista"))
        m.add_separator()
        m.add_command(label="Crear copia de seguridad…", command=self.copia)
        m.add_command(label="Restaurar copia de seguridad…", command=self.restaurar)
        m.add_command(label="Usar otra base de datos (p. ej. en red)…", command=self.cambiar_bd)
        m.add_command(label="Abrir carpeta de datos", command=lambda: abrir_archivo(self.bd.ruta.parent))
        m.add_separator()
        m.add_command(label="Preferencias…", accelerator="Ctrl+,", command=self.preferencias)
        m.add_separator()
        m.add_command(label="Salir", command=self._salir)
        barra.add_cascade(label="Archivo", menu=m)

        e = tk.Menu(barra, tearoff=False)
        self.menu_editar = e
        e.add_command(label="Deshacer", accelerator="Ctrl+Z", command=self.deshacer)
        e.add_command(label="Rehacer", accelerator="Ctrl+Y", command=self.rehacer)
        e.add_separator()
        e.add_command(label="Nuevo equipo", accelerator="Ctrl+N", command=self.nuevo)
        e.add_command(label="Añadir este equipo (detectar datos)…", command=lambda: self.nuevo(detectar=True))
        e.add_command(label="Abrir ficha", accelerator="Intro", command=self.editar)
        e.add_command(label="Editar celda", accelerator="F2", command=self._f2)
        e.add_command(label="Duplicar", accelerator="Ctrl+D", command=self.duplicar)
        e.add_command(label="Eliminar", accelerator="Supr", command=self.eliminar)
        e.add_separator()
        self.menu_estado = tk.Menu(e, tearoff=False)
        for est in ESTADOS:
            self.menu_estado.add_command(label=est, command=lambda s=est: self.cambiar_campo("estado", s))
        e.add_cascade(label="Cambiar estado de los seleccionados", menu=self.menu_estado)
        e.add_command(label="Editar un campo en los seleccionados…", command=self.cambiar_campo_dialogo)
        e.add_separator()
        e.add_command(label="Buscar", accelerator="Ctrl+F", command=self._enfocar_busqueda)
        e.add_command(label="Seleccionar todo", accelerator="Ctrl+A", command=self.seleccionar_todo)
        barra.add_cascade(label="Equipos", menu=e)

        v = tk.Menu(barra, tearoff=False)
        for i, (nombre, _) in enumerate(self.PAGINAS):
            v.add_command(label=nombre, accelerator=f"Ctrl+{i + 1}", command=lambda i=i: self.mostrar(i))
        v.add_separator()
        v.add_command(label="Guardar los filtros actuales como vista…", command=self.guardar_vista)
        v.add_command(label="Columnas…", command=self.dialogo_columnas)
        v.add_command(label="Actualizar", accelerator="F5", command=self.recargar)
        barra.add_cascade(label="Ver", menu=v)

        a = tk.Menu(barra, tearoff=False)
        a.add_command(label="Formato del Excel de importación", command=self.ayuda_formato)
        a.add_command(label="Atajos de teclado", command=self.ayuda_atajos)
        a.add_command(label="Acerca de…", command=self.acerca)
        barra.add_cascade(label="Ayuda", menu=a)
        self.config(menu=barra)

        self.menu_contexto = tk.Menu(self, tearoff=False)
        self.menu_contexto.add_command(label="Abrir ficha", command=self.editar)
        self.menu_contexto.add_command(label="Editar esta celda", command=self._editar_celda_contexto)
        self.menu_contexto.add_command(label="Editar un campo en la selección…",
                                       command=lambda: self.cambiar_campo_dialogo(self._clave_contexto))
        self.menu_contexto.add_cascade(label="Cambiar estado", menu=self.menu_estado)
        self.menu_contexto.add_separator()
        self.menu_contexto.add_command(label="Historial del equipo", command=lambda: self.editar(pestana="historial"))
        self.menu_contexto.add_command(label="Adjuntos del equipo", command=lambda: self.editar(pestana="adjuntos"))
        self.menu_contexto.add_command(label="Imprimir etiquetas…", command=lambda: self.etiquetas("seleccion"))
        self.menu_contexto.add_separator()
        self.menu_contexto.add_command(label="Duplicar", command=self.duplicar)
        self.menu_contexto.add_command(label="Copiar filas", command=self.copiar_fila)
        self.menu_contexto.add_command(label="Exportar seleccionados…", command=lambda: self.exportar("excel", "seleccion"))
        self.menu_contexto.add_separator()
        self.menu_contexto.add_command(label="Eliminar", command=self.eliminar)

    def _enfocar_busqueda(self):
        self.mostrar(0)
        self.busqueda.focus_set()
        return "break"

    def _atajo(self, funcion, permitir_en_texto: bool = False):
        """Crea un manejador de atajo que no interfiere al escribir en un campo de texto."""
        def manejador(evento=None):
            if not permitir_en_texto and foco_en_texto(self.focus_get()):
                return None
            if self.grab_current() not in (None, self):
                return None
            funcion()
            return "break"
        return manejador

    def _atajos(self):
        self.bind_all("<Control-n>", self._atajo(self.nuevo, True))
        self.bind_all("<Control-i>", self._atajo(self.importar, True))
        self.bind_all("<Control-e>", self._atajo(lambda: self.exportar("excel", "vista"), True))
        self.bind_all("<Control-p>", self._atajo(lambda: self.informe("vista"), True))
        self.bind_all("<Control-f>", lambda e: self._enfocar_busqueda())
        self.bind_all("<Control-comma>", self._atajo(self.preferencias, True))
        self.bind_all("<Control-z>", self._atajo(self.deshacer))
        self.bind_all("<Control-y>", self._atajo(self.rehacer))
        self.bind_all("<Control-Z>", self._atajo(self.rehacer))  # Ctrl+Mayús+Z
        self.bind_all("<Control-v>", self._atajo(self.pegar))
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
        self.tabla.bind("<Button-3>", self._click_derecho)
        self.tabla.bind("<<TreeviewSelect>>", lambda e: self._actualizar_estado())
        self.tabla.bind("<Key>", self._tecla_en_tabla)
        # arrastrar cabeceras para reordenar columnas
        self.tabla.bind("<ButtonPress-1>", self._arrastre_inicio, add="+")
        self.tabla.bind("<B1-Motion>", self._arrastre_mover, add="+")
        self.tabla.bind("<ButtonRelease-1>", self._arrastre_fin, add="+")
        for ev in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.tabla.bind(ev, lambda e: self._cerrar_editor(True), add="+")

    # ------------------------------------------------------------ página inventario
    def _pagina_inventario(self):
        px = self.tema.px
        p = ttk.Frame(self.contenido, padding=(px(28), px(22), px(28), px(16)))
        self.paginas.append(p)

        acciones, self.sub_inventario = self._cabecera(p, "Inventario")
        ttk.Button(acciones, text="＋  Nuevo equipo", style="Acento.TButton", command=self.nuevo).pack(side="right")
        ttk.Button(acciones, text="Exportar", command=lambda: self.exportar("excel", "vista")).pack(side="right", padx=(0, 8))
        ttk.Button(acciones, text="Importar", command=self.importar).pack(side="right", padx=(0, 8))
        self.btn_rehacer = ttk.Button(acciones, text="↷", width=3, style="Plano.TButton", command=self.rehacer)
        self.btn_rehacer.pack(side="right", padx=(0, 8))
        self.btn_deshacer = ttk.Button(acciones, text="↶  Deshacer", style="Plano.TButton", command=self.deshacer)
        self.btn_deshacer.pack(side="right")

        # barra de búsqueda y filtros (se reorganiza en dos filas en pantallas estrechas)
        self.barra_filtros = ttk.Frame(p)
        self.barra_filtros.pack(fill="x", pady=(0, px(12)))
        self.zona_busqueda = ttk.Frame(self.barra_filtros)
        self.zona_filtros = ttk.Frame(self.barra_filtros)
        self.var_buscar = tk.StringVar()
        self.busqueda = CampoBusqueda(self.zona_busqueda, self.tema, self.var_buscar,
                                      "Buscar o escanear un QR…", ancho=16)
        ttk.Button(self.zona_busqueda, text="☆ Guardar vista", style="Plano.TButton",
                   command=self.guardar_vista).pack(side="right", padx=(px(6), 0))
        self.busqueda.pack(side="left", fill="x", expand=True)
        self.var_buscar.trace_add("write", lambda *a: self._programar_filtro())
        self.busqueda.bind_entrada("<Escape>", lambda e: self.var_buscar.set(""))
        self.busqueda.bind_entrada("<Return>", lambda e: self._comprobar_qr(forzar=True))

        self.filtros: dict[str, tk.StringVar] = {}
        self.combos: dict[str, ttk.Combobox] = {}
        for clave, todos in TODOS_FILTRO.items():
            var = tk.StringVar(value=todos)
            cb = ttk.Combobox(self.zona_filtros, textvariable=var, state="readonly", width=10)
            cb.pack(side="left", padx=(0, px(8)))
            cb.bind("<<ComboboxSelected>>", lambda e: (self.aplicar_filtro(), self.tabla.focus_set()))
            self.filtros[clave], self.combos[clave] = var, cb
        self.var_garantia = tk.StringVar(value=TODOS)
        cb = ttk.Combobox(self.zona_filtros, textvariable=self.var_garantia, values=FILTROS_GARANTIA,
                          state="readonly", width=11)
        cb.pack(side="left", padx=(0, px(4)))
        cb.bind("<<ComboboxSelected>>", lambda e: (self.aplicar_filtro(), self.tabla.focus_set()))
        self.btn_limpiar = ttk.Button(self.zona_filtros, text="Limpiar", style="Plano.TButton",
                                      command=self.limpiar_filtros)
        self.btn_limpiar.pack(side="left")
        self._disposicion_filtros(False)

        # pie (se empaqueta antes que la tabla para que siempre sea visible)
        pie = ttk.Frame(p)
        pie.pack(fill="x", side="bottom", pady=(px(10), 0))
        for color, attr, texto in ((C["rojo"], None, "Garantía caducada"),
                                   (C["ambar"], "lbl_leyenda_proxima", ""),
                                   (C["terciario"], None, "Baja")):
            tk.Label(pie, text="●", fg=color, bg=C["fondo"], font=(self.tema.familia, 11)).pack(side="left")
            lbl = ttk.Label(pie, text=texto, style="Secundario.TLabel")
            lbl.pack(side="left", padx=(2, px(14)))
            if attr:
                setattr(self, attr, lbl)
        ttk.Label(pie, textvariable=self.estado, style="Secundario.TLabel").pack(side="right")

        barra_sel = ttk.Frame(p)
        barra_sel.pack(fill="x", side="bottom", pady=(px(8), 0))
        self.botones_seleccion = []
        for texto, cmd, estilo in (("Abrir ficha", self.editar, "Plano.TButton"),
                                   ("Duplicar", self.duplicar, "Plano.TButton"),
                                   ("Etiquetas", lambda: self.etiquetas("seleccion"), "Plano.TButton"),
                                   ("Copiar", self.copiar_fila, "Plano.TButton"),
                                   ("Eliminar", self.eliminar, "Rojo.Plano.TButton")):
            b = ttk.Button(barra_sel, text=texto, style=estilo, command=cmd)
            b.pack(side="left", padx=(0, px(2)))
            self.botones_seleccion.append(b)
        self.lbl_ayuda_tabla = ttk.Label(barra_sel, text="Doble clic en una celda para editarla · "
                                                         "arrastra las cabeceras para reordenar",
                                         style="Secundario.TLabel")
        self.lbl_ayuda_tabla.pack(side="right")

        t, marco = tarjeta(p, padding=(px(6), px(4), px(6), px(6)))
        t.pack(fill="both", expand=True)
        cols = ["id"] + CLAVES
        self.tabla = tabla_con_scroll(marco, columns=cols, show="headings", selectmode="extended")
        self.tabla.heading("id", text="ID", command=lambda: self.ordenar("id"))
        anchos = self.cfg.get("anchos_columnas", {})
        self.tabla.column("id", width=int(anchos.get("id", px(50))), anchor="e", stretch=False)
        for c in CAMPOS:
            self.tabla.heading(c.clave, text=c.etiqueta.upper(), command=lambda k=c.clave: self.ordenar(k),
                               anchor="center" if c.tipo == "fecha" else "w")
            self.tabla.column(c.clave, width=int(anchos.get(c.clave, px(c.ancho))), stretch=False,
                              anchor="center" if c.tipo == "fecha" else "w")
        self.tabla.tag_configure("par", background=C["fila_alterna"])
        self.tabla.tag_configure("caducada", background=C["rojo_suave"])
        self.tabla.tag_configure("proxima", background=C["ambar_suave"])
        self.tabla.tag_configure("baja", foreground=C["terciario"])
        self._aplicar_columnas_visibles()
        self._id_filtro = None
        self._actualizar_textos_aviso()

    def _disposicion_filtros(self, estrecho: bool):
        if estrecho == self._estrecho:
            return
        self._estrecho = estrecho
        self.zona_busqueda.pack_forget()
        self.zona_filtros.pack_forget()
        if estrecho:
            self.zona_busqueda.pack(fill="x")
            self.zona_filtros.pack(fill="x", pady=(self.tema.px(8), 0))
        else:
            self.zona_filtros.pack(side="right")
            self.zona_busqueda.pack(side="left", fill="x", expand=True, padx=(0, self.tema.px(10)))

    def _al_redimensionar(self, evento):
        if evento.widget is self:
            self.after_idle(self._ajustar_disposicion)

    def _ajustar_disposicion(self):
        px = self.tema.px
        ancho = self.winfo_width()
        if ancho <= 1:
            return
        self._modo_compacto(ancho < px(1200))
        self.update_idletasks()
        contenido = self.contenido.winfo_width()
        necesario = self.zona_filtros.winfo_reqwidth() + px(290)
        self._disposicion_filtros(contenido - px(56) < necesario)
        if self.lbl_ayuda_tabla.winfo_manager():
            if contenido < px(1200):
                self.lbl_ayuda_tabla.pack_forget()
        elif contenido >= px(1200):
            self.lbl_ayuda_tabla.pack(side="right")

    def _actualizar_textos_aviso(self):
        d = dias_aviso()
        if hasattr(self, "lbl_leyenda_proxima"):
            self.lbl_leyenda_proxima.configure(text=f"Caduca en ≤ {d} días")
        if hasattr(self, "lbl_cifra_proximas"):
            self.lbl_cifra_proximas.configure(text=f"Caducan en ≤ {d} días")

    # ------------------------------------------------------------ página resumen
    def _pagina_resumen(self):
        px = self.tema.px
        p = ttk.Frame(self.contenido, padding=(px(28), px(22), px(28), px(16)))
        self.paginas.append(p)
        acciones, self.sub_resumen = self._cabecera(p, "Resumen")
        ttk.Button(acciones, text="Informe PDF", style="Acento.TButton",
                   command=lambda: self.informe("todo")).pack(side="right")

        fila = ttk.Frame(p)
        fila.pack(fill="x")
        self.cifras: dict[str, tk.StringVar] = {}
        for i, (clave, texto, estilo) in enumerate((
                ("total", "Equipos en total", "CifraAzul.TLabel"),
                ("activos", "Activos", "CifraVerde.TLabel"),
                ("caducadas", "Garantía caducada", "CifraRojo.TLabel"),
                ("proximas", "", "CifraAmbar.TLabel"),
                ("sin_garantia", "Sin fecha de garantía", "CifraGris.TLabel"))):
            t, inn = tarjeta(fila)
            t.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else px(12), 0))
            fila.columnconfigure(i, weight=1, uniform="cifras")
            var = tk.StringVar(value="0")
            ttk.Label(inn, textvariable=var, style=estilo).pack(anchor="w")
            lbl = ttk.Label(inn, text=texto, style="EnTarjetaSec.TLabel")
            lbl.pack(anchor="w")
            if clave == "proximas":
                self.lbl_cifra_proximas = lbl
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

        cab = ttk.Frame(p)
        cab.pack(fill="x", pady=(px(2), px(6)))
        ttk.Label(cab, text="GARANTÍAS", style="Seccion.TLabel").pack(side="left", padx=px(4))
        ttk.Button(cab, text="Exportar a Excel", style="Plano.TButton",
                   command=self._exportar_garantias).pack(side="right")
        ttk.Button(cab, text="Etiquetas", style="Plano.TButton",
                   command=lambda: self._etiquetas_lista(self._lista_garantias)).pack(side="right")
        self.var_periodo = tk.StringVar()
        self.cb_periodo = ttk.Combobox(cab, textvariable=self.var_periodo, state="readonly", width=32)
        self.cb_periodo.pack(side="right", padx=px(8))
        self.cb_periodo.bind("<<ComboboxSelected>>", lambda e: self._pintar_garantias())
        self.lbl_garantias = ttk.Label(cab, text="", style="Secundario.TLabel")
        self.lbl_garantias.pack(side="left", padx=px(8))
        t, marco = tarjeta(p, padding=px(6))
        t.pack(fill="both", expand=True)
        cols = ("fin_garantia", "dias", "nombre_equipo", "persona", "delegacion", "marca", "modelo", "num_serie")
        self.tabla_garantias = tabla_con_scroll(marco, columns=cols, show="headings", height=5)
        for c in cols:
            texto = "DÍAS" if c == "dias" else CAMPOS_POR_CLAVE[c].etiqueta.upper()
            self.tabla_garantias.heading(c, text=texto, anchor="e" if c == "dias" else "w")
            self.tabla_garantias.column(c, width=px(70 if c == "dias" else 150), anchor="e" if c == "dias" else "w")
        self.tabla_garantias.tag_configure("caducada", foreground=C["rojo"])
        self.tabla_garantias.tag_configure("proxima", foreground=C["ambar_texto"])
        self.tabla_garantias.bind("<Double-1>", self._editar_desde_garantias)
        self._lista_garantias: list[dict] = []

    # ------------------------------------------------------------ página historial
    def _pagina_historial(self):
        px = self.tema.px
        p = ttk.Frame(self.contenido, padding=(px(28), px(22), px(28), px(16)))
        self.paginas.append(p)
        acciones, self.sub_historial = self._cabecera(p, "Historial")
        ttk.Button(acciones, text="Exportar a Excel", command=self.exportar_historial).pack(side="right")
        ttk.Button(acciones, text="Vaciar", style="Peligro.TButton", command=self.vaciar_historial).pack(
            side="right", padx=(0, 8))
        t, marco = tarjeta(p, padding=px(6))
        t.pack(fill="both", expand=True)
        cols = ("fecha", "usuario", "accion", "equipo", "detalle")
        self.tabla_hist = tabla_con_scroll(marco, columns=cols, show="headings")
        for c, txt, w in (("fecha", "FECHA", 150), ("usuario", "USUARIO", 150), ("accion", "ACCIÓN", 130),
                          ("equipo", "EQUIPO", 260), ("detalle", "DETALLE", 700)):
            self.tabla_hist.heading(c, text=txt, anchor="w")
            self.tabla_hist.column(c, width=px(w), stretch=c == "detalle")
        for accion, color in (("Alta", C["verde"]), ("Modificación", C["acento"]), ("Baja (eliminado)", C["rojo"]),
                              ("Importación", C["morado"]), ("Exportación", C["secundario"]),
                              ("Deshacer", C["ambar_texto"]), ("Rehacer", C["ambar_texto"]),
                              ("Adjunto añadido", C["morado"]), ("Adjunto eliminado", C["secundario"])):
            self.tabla_hist.tag_configure(accion, foreground=color)
        self.tabla_hist.bind("<Double-1>", self._abrir_desde_historial)

    # ------------------------------------------------------------ datos
    def recargar(self):
        self._cerrar_editor(True)
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
        self.lbl_lateral_usuario.configure(text=f"Usuario: {self.bd.usuario}")
        self.aplicar_filtro()
        self._cambio_pagina()
        self._actualizar_deshacer()
        self._version_bd = self.bd.version_datos()

    def _programar_filtro(self):
        if self._id_filtro:
            self.after_cancel(self._id_filtro)
        self._id_filtro = self.after(250, self._filtrar_o_qr)

    def _filtrar_o_qr(self):
        if not self._comprobar_qr():
            self.aplicar_filtro()

    def _comprobar_qr(self, forzar: bool = False) -> bool:
        """Si en el buscador se ha escaneado un QR de etiqueta («INV:<id> | …») abre ese equipo."""
        eid = id_desde_qr(self.var_buscar.get())
        if eid is None:
            return False
        self.var_buscar.set("")
        self.aplicar_filtro()
        if not any(e["id"] == eid for e in self.equipos):
            self.estado.set(f"El equipo con ID {eid} no existe en esta base de datos.")
            self.bell()
            return True
        self._seleccionar(eid)
        self.after(50, lambda: self.editar(eid))
        return True

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
        if self._editor:
            self._cerrar_editor(True)
        seleccion = set(self.tabla.selection())
        foco = self.tabla.focus()
        yv = self.tabla.yview()[0]
        self.tabla.delete(*self.tabla.get_children())
        for n, eq in enumerate(self.visibles):
            self.tabla.insert("", "end", iid=str(eq["id"]), values=self._valores_fila(eq),
                              tags=self._tags_fila(eq, n))
        restaurar = [i for i in seleccion if self.tabla.exists(i)]
        if restaurar:
            self.tabla.selection_set(restaurar)
        if foco and self.tabla.exists(foco):
            self.tabla.focus(foco)
        self.tabla.yview_moveto(yv)
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
        self._actualizar_deshacer()
        self._version_bd = self.bd.version_datos()

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
        self.cfg["orden_tabla"] = list(self.orden)
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

    def _seleccionar(self, ids):
        ids = [ids] if isinstance(ids, int) else list(ids)
        existentes = [str(i) for i in ids if self.tabla.exists(str(i))]
        if existentes:
            self.tabla.selection_set(existentes)
            self.tabla.see(existentes[0])
            self.tabla.focus(existentes[0])

    # ------------------------------------------------------------ deshacer / rehacer
    def _actualizar_deshacer(self):
        d, r = self.bd.nombre_deshacer(), self.bd.nombre_rehacer()
        self.btn_deshacer.state(["!disabled"] if d else ["disabled"])
        self.btn_rehacer.state(["!disabled"] if r else ["disabled"])
        self.menu_editar.entryconfigure(0, label=f"Deshacer «{d}»" if d else "Deshacer",
                                        state="normal" if d else "disabled")
        self.menu_editar.entryconfigure(1, label=f"Rehacer «{r}»" if r else "Rehacer",
                                        state="normal" if r else "disabled")

    def deshacer(self):
        self._cerrar_editor(True)
        res = self.bd.deshacer()
        if res is None:
            self.estado.set("No hay nada que deshacer.")
            return
        nombre, avisos = res
        self.recargar()
        self.estado.set(f"↶ Deshecho: {nombre}")
        if avisos:
            messagebox.showwarning(TITULO, "Se ha deshecho, pero algunos campos los había cambiado otra persona "
                                           "después y se han dejado como estaban:\n\n" + "\n".join(avisos[:15]))

    def rehacer(self):
        self._cerrar_editor(True)
        res = self.bd.rehacer()
        if res is None:
            self.estado.set("No hay nada que rehacer.")
            return
        nombre, avisos = res
        self.recargar()
        self.estado.set(f"↷ Rehecho: {nombre}")
        if avisos:
            messagebox.showwarning(TITULO, "Algunos campos no se han podido rehacer:\n\n" + "\n".join(avisos[:15]))

    # ------------------------------------------------------------ cambios de otras personas
    def _vigilar_cambios(self):
        intervalo = int(self.cfg.get("refresco_seg", 5) or 0)
        try:
            if intervalo:
                version = self.bd.version_datos()
                if version != self._version_bd or self._cambios_pendientes:
                    ocupado = self._editor is not None or self.grab_current() is not None
                    if ocupado:
                        self._cambios_pendientes = True
                    else:
                        self._cambios_pendientes = False
                        self.recargar()
                        self.estado.set(f"↻ Actualizado con cambios de otra persona ({datetime.now():%H:%M:%S})")
        except Exception:  # noqa: BLE001  (p. ej. carpeta de red no disponible un momento)
            pass
        self.after(max(2, intervalo or 5) * 1000, self._vigilar_cambios)

    def _resolver_conflicto(self, c: Conflicto) -> str:
        """Pregunta qué hacer ante un conflicto. Devuelve 'mio', 'suyo' o 'cancelar'."""
        if c.eliminado:
            return "eliminado"
        r = messagebox.askyesnocancel(
            TITULO, c.describir() + "\n\n¿Guardar TUS valores en esos campos?\n\n"
                                    "• Sí: se guardan tus valores.\n• No: se conservan los de la otra persona "
                                    "(el resto de tus cambios sí se guarda).\n• Cancelar: no guardar nada.",
            icon="warning")
        return "mio" if r else ("suyo" if r is False else "cancelar")

    # ------------------------------------------------------------ edición en la propia tabla
    def _columna_en(self, x: int) -> str | None:
        col = self.tabla.identify_column(x)
        try:
            idx = int(col.lstrip("#")) - 1
        except ValueError:
            return None
        visibles = list(self.tabla["displaycolumns"])
        if visibles == ["#all"]:
            visibles = list(self.tabla["columns"])
        return visibles[idx] if 0 <= idx < len(visibles) else None

    def _ver_columna(self, clave: str):
        cols = list(self.tabla["displaycolumns"])
        if cols == ["#all"]:
            cols = list(self.tabla["columns"])
        anchos = [int(self.tabla.column(c, "width")) for c in cols]
        total = sum(anchos) or 1
        inicio = sum(anchos[:cols.index(clave)]) if clave in cols else 0
        fin = inicio + (anchos[cols.index(clave)] if clave in cols else 0)
        a, b = self.tabla.xview()
        if inicio < a * total or fin > b * total:
            self.tabla.xview_moveto(max(0.0, (inicio - 20) / total))

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
        self._editor = (w_ed, var, iid, clave, eq[clave] or "")
        w_ed.bind("<Return>", lambda e: (self._cerrar_editor(True), self.tabla.focus_set(), "break")[-1])
        w_ed.bind("<KP_Enter>", lambda e: (self._cerrar_editor(True), self.tabla.focus_set(), "break")[-1])
        w_ed.bind("<Escape>", lambda e: (self._cerrar_editor(False), self.tabla.focus_set(), "break")[-1])
        w_ed.bind("<Tab>", lambda e: (self._mover_editor(+1), "break")[-1])
        w_ed.bind("<ISO_Left_Tab>", lambda e: (self._mover_editor(-1), "break")[-1])
        w_ed.bind("<Shift-Tab>", lambda e: (self._mover_editor(-1), "break")[-1])
        w_ed.bind("<FocusOut>", lambda e: self.after(80, self._cerrar_si_sin_foco))

    def _cerrar_si_sin_foco(self):
        ed = self._editor
        if not ed:
            return
        try:
            foco = self.focus_get()
        except (KeyError, tk.TclError):
            foco = None
        if foco is not None and (foco == ed[0] or str(foco).startswith(str(ed[0])) or "popdown" in str(foco)):
            return
        self._cerrar_editor(True)

    def _mover_editor(self, paso: int):
        ed = self._editor
        if not ed:
            return
        iid, clave = ed[2], ed[3]
        self._cerrar_editor(True)
        visibles = [c for c in self.tabla["displaycolumns"] if c in CLAVES]
        i = visibles.index(clave) + paso if clave in visibles else 0
        if 0 <= i < len(visibles):
            self.editar_celda(iid, visibles[i])

    def _cerrar_editor(self, guardar: bool):
        ed = self._editor
        if not ed:
            return
        self._editor = None
        w_ed, var, iid, clave, original = ed
        nuevo = var.get().strip()
        try:
            w_ed.destroy()
        except tk.TclError:
            pass
        if not guardar:
            return
        campo = CAMPOS_POR_CLAVE[clave]
        if campo.tipo == "fecha":
            nuevo = parsear_fecha(nuevo)
        if nuevo == original:
            return
        try:
            self.bd.actualizar_campo(int(iid), clave, nuevo, esperado=original)
        except Conflicto as c:
            decision = self._resolver_conflicto(c)
            if decision == "mio":
                self.bd.actualizar_campo(int(iid), clave, nuevo, esperado=original, forzar=True)
            elif decision == "eliminado":
                messagebox.showwarning(TITULO, c.describir())
                self.recargar()
                return
        self._refrescar_equipo(int(iid))
        valor = mostrar_fecha(nuevo) if campo.tipo == "fecha" else nuevo
        self.estado.set(f"Guardado: {campo.etiqueta} = «{valor}» · Ctrl+Z para deshacer")

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
            visibles = [c for c in self.tabla["displaycolumns"] if c in CLAVES]
            if visibles:
                self.editar_celda(foco, visibles[0])
        return "break"

    def _tecla_en_tabla(self, evento):
        """Escribir con la tabla seleccionada lleva al buscador (así funciona un lector de QR)."""
        if evento.state & 0x4 or evento.state & 0x20000 or evento.state & 0x8 and os.name != "nt":
            return None  # Ctrl / Alt
        if evento.char and evento.char.isprintable() and not evento.keysym.startswith(("F", "Up", "Down")):
            self.busqueda.entrada.focus_set()
            self.busqueda.entrada.insert("end", evento.char)
            return "break"
        return None

    # ------------------------------------------------------------ columnas
    def _aplicar_columnas_visibles(self):
        orden = [c for c in self.cfg.get("orden_columnas", []) if c in CLAVES]
        orden += [c for c in CLAVES if c not in orden]
        self.tabla["displaycolumns"] = ["id"] + [c for c in orden if c not in self.columnas_ocultas]

    def _guardar_columnas(self):
        visibles = [c for c in self.tabla["displaycolumns"] if c in CLAVES]
        ocultas = [c for c in self.cfg.get("orden_columnas", CLAVES) if c in CLAVES and c not in visibles]
        self.cfg["orden_columnas"] = visibles + [c for c in CLAVES if c not in visibles and c not in ocultas] + ocultas
        self.cfg["anchos_columnas"] = {c: int(self.tabla.column(c, "width")) for c in ["id"] + CLAVES}
        self.cfg["columnas_ocultas"] = sorted(self.columnas_ocultas)

    def _arrastre_inicio(self, e):
        self._arrastre = None
        if self.tabla.identify_region(e.x, e.y) == "heading":
            col = self._columna_en(e.x)
            if col in CLAVES:
                self._arrastre = {"col": col, "x": e.x, "moviendo": False}

    def _arrastre_mover(self, e):
        if self._arrastre and abs(e.x - self._arrastre["x"]) > 12:
            self._arrastre["moviendo"] = True
            self.tabla.configure(cursor="fleur" if os.name != "nt" else "size_we")

    def _arrastre_fin(self, e):
        datos, self._arrastre = self._arrastre, None
        if not datos or not datos["moviendo"]:
            return None
        self.tabla.configure(cursor="")
        destino = self._columna_en(e.x)
        origen = datos["col"]
        if destino not in CLAVES or destino == origen:
            return "break"
        cols = [c for c in self.tabla["displaycolumns"]]
        i_dest = cols.index(destino)
        derecha = i_dest > cols.index(origen)
        cols.remove(origen)
        cols.insert(cols.index(destino) + (1 if derecha else 0), origen)
        self.tabla["displaycolumns"] = cols
        self._guardar_columnas()
        self.guardar_cfg()
        self.estado.set(f"Columna «{CAMPOS_POR_CLAVE[origen].etiqueta}» movida")
        return "break"

    def dialogo_columnas(self):
        self._guardar_columnas()
        d = DialogoColumnas(self)
        self.wait_window(d)
        if d.resultado is None:
            return
        if d.restablecer:
            self.columnas_ocultas = set()
            self.cfg.pop("orden_columnas", None)
            self.cfg.pop("anchos_columnas", None)
            px = self.tema.px
            self.tabla.column("id", width=px(50))
            for c in CAMPOS:
                self.tabla.column(c.clave, width=px(c.ancho))
        else:
            self.columnas_ocultas = d.resultado
        self._aplicar_columnas_visibles()
        self._guardar_columnas()
        self.guardar_cfg()

    # ------------------------------------------------------------ vistas guardadas
    def _estado_filtros(self) -> dict:
        return {"buscar": self.var_buscar.get(),
                "filtros": {k: v.get() for k, v in self.filtros.items() if v.get() != TODOS_FILTRO[k]},
                "garantia": self.var_garantia.get() if self.var_garantia.get() != TODOS else "",
                "orden": list(self.orden)}

    def descripcion_filtros(self, estado: dict | None = None) -> str:
        estado = estado or self._estado_filtros()
        partes = [f"{CAMPOS_POR_CLAVE[k].etiqueta}: {v}" for k, v in estado["filtros"].items()]
        if estado.get("garantia"):
            partes.append(f"Garantía: {estado['garantia'].lower()}")
        if estado.get("buscar"):
            partes.append(f"búsqueda «{estado['buscar']}»")
        return " · ".join(partes) if partes else "Todos los equipos"

    def guardar_vista(self):
        estado = self._estado_filtros()
        if not estado["filtros"] and not estado["garantia"] and not estado["buscar"]:
            messagebox.showinfo(TITULO, "Primero aplica algún filtro o búsqueda; después guárdalo como vista "
                                        "para tenerlo a un clic en la barra lateral.")
            return
        d = DialogoTexto(self, "Guardar vista", "Nombre de la vista", self.descripcion_filtros(estado)[:40],
                         self.descripcion_filtros(estado))
        self.wait_window(d)
        if not d.resultado:
            return
        vistas = [v for v in self.cfg.get("vistas", []) if v.get("nombre") != d.resultado]
        vistas.append({"nombre": d.resultado, **estado})
        self.cfg["vistas"] = vistas
        self.guardar_cfg()
        self._pintar_vistas()
        self.estado.set(f"Vista «{d.resultado}» guardada en la barra lateral")

    def _pintar_vistas(self):
        for w in self.marco_vistas.winfo_children():
            w.destroy()
        self._nav = [n for n in self._nav if n[0].winfo_exists()]
        vistas = self.cfg.get("vistas", [])
        if not vistas:
            ttk.Label(self.marco_vistas, text="Filtra la tabla y pulsa «☆ Guardar vista»", style="Lateral.TLabel",
                      wraplength=self.tema.px(170)).pack(anchor="w", padx=self.tema.px(10))
            return
        for i, v in enumerate(vistas):
            b = ttk.Button(self.marco_vistas, text=f"☆   {v['nombre']}", style="Lateral.TButton",
                           command=lambda i=i: self.aplicar_vista(i))
            b.pack(fill="x", pady=1)
            b.bind("<Button-3>", lambda e, i=i: self._menu_vista(e, i))

    def aplicar_vista(self, i: int):
        vistas = self.cfg.get("vistas", [])
        if not 0 <= i < len(vistas):
            return
        v = vistas[i]
        self.mostrar(0)
        for k, var in self.filtros.items():
            valor = v.get("filtros", {}).get(k)
            var.set(valor if valor else TODOS_FILTRO[k])
        self.var_garantia.set(v.get("garantia") or TODOS)
        if v.get("orden"):
            self.orden = tuple(v["orden"])
        self.var_buscar.set(v.get("buscar", ""))
        self.aplicar_filtro()
        self.estado.set(f"Vista «{v['nombre']}»: {len(self.visibles)} equipos")

    def _menu_vista(self, evento, i: int):
        m = tk.Menu(self, tearoff=False)
        m.add_command(label="Aplicar", command=lambda: self.aplicar_vista(i))
        m.add_command(label="Actualizar con los filtros actuales", command=lambda: self._actualizar_vista(i))
        m.add_command(label="Cambiar nombre…", command=lambda: self._renombrar_vista(i))
        m.add_separator()
        m.add_command(label="Eliminar vista", command=lambda: self._eliminar_vista(i))
        m.tk_popup(evento.x_root, evento.y_root)

    def _actualizar_vista(self, i: int):
        vistas = self.cfg.get("vistas", [])
        vistas[i] = {"nombre": vistas[i]["nombre"], **self._estado_filtros()}
        self.guardar_cfg()

    def _renombrar_vista(self, i: int):
        vistas = self.cfg.get("vistas", [])
        d = DialogoTexto(self, "Cambiar nombre", "Nombre de la vista", vistas[i]["nombre"])
        self.wait_window(d)
        if d.resultado:
            vistas[i]["nombre"] = d.resultado
            self.guardar_cfg()
            self._pintar_vistas()

    def _eliminar_vista(self, i: int):
        vistas = self.cfg.get("vistas", [])
        if messagebox.askyesno(TITULO, f"¿Eliminar la vista «{vistas[i]['nombre']}»?\n(No se borra ningún equipo.)"):
            vistas.pop(i)
            self.guardar_cfg()
            self._pintar_vistas()

    # ------------------------------------------------------------ acciones sobre equipos
    def nuevo(self, base: dict | None = None, detectar: bool = False):
        self._cerrar_editor(True)
        f = FormularioEquipo(self, base, "Nuevo equipo", detectar_al_abrir=detectar)
        self.wait_window(f)
        if f.resultado:
            nuevo = self.bd.insertar(f.resultado)
            self.recargar()
            if not any(e["id"] == nuevo for e in self.visibles):
                self.limpiar_filtros()
            self._seleccionar(nuevo)
            self.estado.set("Equipo añadido · Ctrl+Z para deshacer")

    def editar(self, equipo_id: int | None = None, pestana: str = "datos"):
        self._cerrar_editor(True)
        if equipo_id is None:
            sel = self.seleccionados()
            if not sel:
                messagebox.showinfo(TITULO, "Selecciona un equipo.")
                return
            equipo_id = sel[0]
        eq = self.bd.obtener(equipo_id)
        if not eq:
            return
        base = dict(eq)
        valores = eq
        while True:
            f = FormularioEquipo(self, valores, eq.get("nombre_equipo") or eq.get("persona") or f"Equipo {eq['id']}",
                                 pestana=pestana)
            self.wait_window(f)
            if not f.resultado:
                self._refrescar_equipo(equipo_id)  # puede haber cambiado (adjuntos, historial)
                return
            try:
                self.bd.actualizar(equipo_id, f.resultado, base=base)
            except Conflicto as c:
                decision = self._resolver_conflicto(c)
                if decision == "eliminado":
                    if messagebox.askyesno(TITULO, c.describir() + "\n\n¿Guardarlo como un equipo nuevo?"):
                        nuevo = self.bd.insertar(f.resultado)
                        self.recargar()
                        self._seleccionar(nuevo)
                    else:
                        self.recargar()
                    return
                if decision == "mio":
                    self.bd.actualizar(equipo_id, f.resultado, base=base, forzar=True)
                elif decision == "suyo":
                    actual = self.bd.obtener(equipo_id) or {}
                    mis = dict(f.resultado)
                    for clave, *_ in c.campos:
                        mis[clave] = actual.get(clave, "")
                    self.bd.actualizar(equipo_id, mis, base=base)
                else:  # cancelar: volver a la ficha con los datos actuales y mis cambios
                    actual = self.bd.obtener(equipo_id)
                    if not actual:
                        return
                    base = dict(actual)
                    valores = {**actual, **f.resultado}
                    pestana = "datos"
                    continue
            break
        self.recargar()
        self._seleccionar(equipo_id)
        self.estado.set("Cambios guardados · Ctrl+Z para deshacer")

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
        pregunta += ("\n\nSe puede deshacer con Ctrl+Z. Si el equipo se ha retirado, también puedes "
                     "marcarlo como «Baja» en lugar de eliminarlo.")
        if messagebox.askyesno(TITULO, pregunta, icon="warning"):
            self.bd.eliminar(sel)
            self.recargar()
            self.estado.set(f"{len(sel)} equipo(s) eliminado(s) · Ctrl+Z para deshacer")

    def cambiar_campo(self, clave: str, valor: str):
        sel = self.seleccionados()
        if not sel:
            messagebox.showinfo(TITULO, "Selecciona uno o varios equipos.")
            return
        etiqueta = CAMPOS_POR_CLAVE[clave].etiqueta
        with self.bd.operacion(f"{etiqueta} en {len(sel)} equipo{'s' if len(sel) != 1 else ''}"):
            for i in sel:
                self.bd.actualizar_campo(i, clave, valor)
        self.recargar()
        self.estado.set(f"{etiqueta} cambiado en {len(sel)} equipo(s) · Ctrl+Z para deshacer")

    def cambiar_campo_dialogo(self, clave: str | None = None):
        sel = self.seleccionados()
        if not sel:
            messagebox.showinfo(TITULO, "Selecciona uno o varios equipos.")
            return
        d = DialogoCampoMasivo(self, len(sel), clave)
        self.wait_window(d)
        if d.resultado:
            self.cambiar_campo(*d.resultado)

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
        self.estado.set(f"{len(sel)} fila(s) copiadas · se pueden pegar en Excel o aquí con Ctrl+V")

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

    # ------------------------------------------------------------ importar / pegar / exportar
    def importar(self, ruta: str | None = None):
        ruta = ruta or filedialog.askopenfilename(
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
        except Cancelado:
            return
        self._tras_importar(d)

    def _tras_importar(self, d: DialogoImportar):
        try:
            self.wait_window(d)
        except tk.TclError:  # el diálogo ya se ha cerrado
            pass
        if d.importado:
            self.limpiar_filtros()
            self.recargar()
            self.mostrar(0)
            self._seleccionar(d.ids_nuevos)

    def pegar(self):
        try:
            texto = self.clipboard_get()
        except tk.TclError:
            texto = ""
        if not texto.strip():
            messagebox.showinfo(TITULO, "El portapapeles está vacío.\n\nCopia filas en Excel (Ctrl+C) y pulsa "
                                        "Ctrl+V con la tabla del inventario seleccionada.")
            return
        res = leer_texto_tabla(texto)
        if not res.filas:
            messagebox.showinfo(TITULO, "El portapapeles no contiene filas de datos reconocibles.")
            return
        try:
            d = DialogoImportar(self, res=res, nombre="Portapapeles")
        except Cancelado:
            return
        self._tras_importar(d)

    def _pedir_destino(self, extension: str, nombre: str) -> Path | None:
        tipos = {".xlsx": [("Excel", "*.xlsx")], ".csv": [("CSV (separado por ;)", "*.csv")],
                 ".pdf": [("PDF", "*.pdf")]}[extension]
        ruta = filedialog.asksaveasfilename(parent=self, defaultextension=extension, filetypes=tipos,
                                            initialfile=nombre)
        return Path(ruta) if ruta else None

    def _equipos_alcance(self, alcance: str) -> list[dict] | None:
        if alcance == "todo":
            return self.equipos
        if alcance == "seleccion":
            ids = set(self.seleccionados())
            datos = [e for e in self.visibles if e["id"] in ids]
            if not datos:
                messagebox.showinfo(TITULO, "No hay equipos seleccionados.")
                return None
            return datos
        return self.visibles

    def exportar(self, formato: str, alcance: str):
        datos = self._equipos_alcance(alcance)
        if datos is None:
            return
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

    def informe(self, alcance: str = "vista"):
        datos = self._equipos_alcance(alcance)
        if datos is None:
            return
        if not datos:
            messagebox.showinfo(TITULO, "No hay equipos que incluir en el informe.")
            return
        destino = self._pedir_destino(".pdf", f"informe_inventario_{datetime.now():%Y%m%d}.pdf")
        if not destino:
            return
        descripcion = "Todo el inventario" if alcance == "todo" else self.descripcion_filtros()
        self.configure(cursor="watch")
        self.update_idletasks()
        try:
            generar_informe(destino, datos, descripcion, self.bd.usuario)
        except PermissionError:
            messagebox.showerror(TITULO, "No se puede escribir el archivo. ¿Está abierto en otro programa?")
            return
        finally:
            self.configure(cursor="")
        self.bd.registrar(None, destino.name, "Exportación", f"Informe PDF de {len(datos)} equipos")
        self.bd.commit()
        abrir_archivo(destino)

    def etiquetas(self, alcance: str = "seleccion"):
        datos = self._equipos_alcance(alcance)
        if datos is None:
            return
        self._etiquetas_lista(datos)

    def _etiquetas_lista(self, datos: list[dict]):
        if not datos:
            messagebox.showinfo(TITULO, "No hay equipos para imprimir etiquetas.")
            return
        d = DialogoEtiquetas(self, datos)
        self.wait_window(d)

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
        ws.append(["Fecha", "Usuario", "Acción", "Equipo", "Detalle"])
        for c in ws[1]:
            c.font = Font(bold=True)
        for h in self.bd.historial(100000):
            ws.append([h["fecha"], h.get("usuario", ""), h["accion"], h["equipo"], h["detalle"]])
        for col, w in zip("ABCDE", (20, 22, 18, 40, 100)):
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

    # ------------------------------------------------------------ copias y base de datos
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
        self.cfg["ruta_bd"] = str(Path(ruta).resolve())
        self.guardar_cfg()
        self.recargar()
        messagebox.showinfo(TITULO, f"Usando la base de datos:\n{ruta}\n\n"
                                    "Se recordará la próxima vez que abras el programa. Si otras personas "
                                    "la usan a la vez, verás sus cambios automáticamente.")

    # ------------------------------------------------------------ preferencias
    def preferencias(self):
        tema_antes = self.cfg.get("tema", MODOS[0])
        d = DialogoPreferencias(self)
        self.wait_window(d)
        if not d.cambios:
            return
        self.cfg.update(d.cambios)
        self.guardar_cfg()
        fijar_dias_aviso(self.cfg["dias_aviso"])
        self.bd.usuario = usuario_actual()
        self._actualizar_textos_aviso()
        self.recargar()
        if self.cfg["tema"] != tema_antes and messagebox.askyesno(
                TITULO, "El nuevo tema se aplica al reiniciar el programa.\n\n¿Reiniciar ahora?"):
            self.reiniciar()

    def reiniciar(self):
        self._salir(destruir=False)
        if getattr(sys, "frozen", False):
            orden = [sys.executable]
        else:
            orden = [sys.executable, str(Path(sys.argv[0]).resolve())]
        try:
            subprocess.Popen(orden + [str(self.bd.ruta)], close_fds=True)
        finally:
            self.destroy()

    # ------------------------------------------------------------ resumen / historial
    def _cambio_pagina(self):
        if self.pagina_actual == 1:
            self._pintar_resumen()
        elif self.pagina_actual == 2:
            self._pintar_historial()

    def _pintar_resumen(self):
        eqs = self.equipos
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

        periodos = periodos_garantia()
        nombres = [p[0] for p in periodos]
        actual = self.var_periodo.get()
        self.cb_periodo["values"] = nombres
        if actual not in nombres:
            # conservar la elección aunque cambie el texto (p. ej. al cambiar los días de aviso)
            idx = next((i for i, n in enumerate(nombres) if n.split(" (")[0] == actual.split(" (")[0]), 0)
            self.var_periodo.set(nombres[idx])
        self._pintar_garantias()

    def _pintar_garantias(self):
        hoy = date.today()
        periodo = next((p for p in periodos_garantia() if p[0] == self.var_periodo.get()), periodos_garantia()[0])
        _, desde, hasta = periodo
        lista = []
        for e in self.equipos:
            if e.get("estado") == "Baja":
                continue
            d = fecha_a_date(e.get("fin_garantia", ""))
            if d is None or (desde and d < desde) or (hasta and d > hasta):
                continue
            lista.append(e)
        lista.sort(key=lambda e: e["fin_garantia"])
        self._lista_garantias = lista
        tg = self.tabla_garantias
        tg.delete(*tg.get_children())
        for e in lista:
            dias = (fecha_a_date(e["fin_garantia"]) - hoy).days
            texto_dias = f"hace {-dias}" if dias < 0 else ("hoy" if dias == 0 else f"en {dias}")
            tg.insert("", "end", iid=str(e["id"]), tags=("caducada" if dias < 0 else "proxima",),
                      values=(mostrar_fecha(e["fin_garantia"]), texto_dias, e["nombre_equipo"], e["persona"],
                              e["delegacion"], e["marca"], e["modelo"], e["num_serie"]))
        self.lbl_garantias.configure(text=f"{len(lista)} equipo{'s' if len(lista) != 1 else ''}")

    def _exportar_garantias(self):
        if not self._lista_garantias:
            messagebox.showinfo(TITULO, "No hay equipos en el periodo elegido.")
            return
        nombre = self.var_periodo.get().split(" (")[0].replace(" ", "_").lower()
        destino = self._pedir_destino(".xlsx", f"garantias_{nombre}_{datetime.now():%Y%m%d}.xlsx")
        if not destino:
            return
        try:
            exportar_excel(destino, self._lista_garantias, "Garantías")
        except PermissionError:
            messagebox.showerror(TITULO, "No se puede escribir el archivo.\n¿Está abierto en Excel?")
            return
        abrir_archivo(destino)

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
            self.tabla_hist.insert("", "end", iid=f"h{h['id']}", tags=(h["accion"],),
                                   values=(h["fecha"], h.get("usuario", ""), h["accion"], h["equipo"],
                                           h["detalle"].replace("\n", " ")))
        self._ids_hist = {f"h{h['id']}": h.get("equipo_id") for h in hist}
        usuarios = len({h.get("usuario") for h in hist if h.get("usuario")})
        self.sub_historial.set(f"Últimos {len(hist)} cambios · {usuarios} usuario{'s' if usuarios != 1 else ''}"
                               if hist else "Todavía no hay cambios")

    def _abrir_desde_historial(self, _e):
        sel = self.tabla_hist.selection()
        if sel:
            eid = getattr(self, "_ids_hist", {}).get(sel[0])
            if eid and self.bd.obtener(eid):
                self.editar(eid, pestana="historial")

    def _aviso_garantias(self):
        if not self.cfg.get("avisar_garantias", True):
            return
        hoy = date.today().isoformat()
        if self.cfg.get("ultimo_aviso") == hoy:
            return
        prox = [e for e in self.equipos if e["_garantia"] == "proxima" and e.get("estado") != "Baja"]
        if prox:
            self.cfg["ultimo_aviso"] = hoy
            self.guardar_cfg()
            if messagebox.askyesno(TITULO, f"Hay {len(prox)} equipo(s) cuya garantía caduca en los próximos "
                                           f"{dias_aviso()} días.\n\n¿Ver el listado?"):
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
               "También puedes copiar filas en Excel y pegarlas aquí con Ctrl+V.")
        messagebox.showinfo("Formato de importación", txt)

    def ayuda_atajos(self):
        messagebox.showinfo("Atajos de teclado", "\n".join((
            "Ctrl+Z / Ctrl+Y — Deshacer / rehacer",
            "Ctrl+N — Nuevo equipo",
            "Ctrl+F — Buscar (también para escanear códigos QR)",
            "Ctrl+V — Pegar filas copiadas de Excel",
            "Ctrl+C — Copiar las filas seleccionadas",
            "Ctrl+I / Ctrl+E — Importar / exportar Excel",
            "Ctrl+P — Informe PDF de la vista actual",
            "Ctrl+1 / 2 / 3 — Inventario / Resumen / Historial",
            "Ctrl+, — Preferencias",
            "",
            "Doble clic en una celda o F2 — Editar en la tabla",
            "Tab / Mayús+Tab — Siguiente / anterior campo al editar",
            "Intro — Abrir la ficha completa del equipo",
            "Supr — Eliminar los seleccionados",
            "F5 — Actualizar")))

    def acerca(self):
        messagebox.showinfo("Acerca de", f"{TITULO}\nVersión {__version__}\n\n"
                                         f"Base de datos:\n{self.bd.ruta}\n\n"
                                         f"Adjuntos:\n{self.bd.carpeta_adjuntos()}\n\n"
                                         f"Usuario: {self.bd.usuario}\n\n"
                                         "Copias automáticas diarias en la carpeta «copias».")

    def _salir(self, destruir: bool = True):
        try:
            self._cerrar_editor(True)
            self._guardar_columnas()
            self.guardar_cfg()
            self.bd.commit()
            self.bd.cerrar()
        except Exception:  # noqa: BLE001
            pass
        finally:
            # cancelar temporizadores pendientes para no dejar avisos al cerrar
            try:
                for ident in self.tk.splitlist(self.tk.call("after", "info")):
                    self.after_cancel(ident)
            except tk.TclError:
                pass
            if destruir:
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


# =====================================================================
#  Autoprueba (la usa la compilación automática para verificar el .exe)
# =====================================================================
def _autoprueba(argv: list[str]) -> int:
    """Uso: InventarioPC.exe --selftest <carpeta_salida>"""
    import tempfile
    import traceback
    salida = Path(argv[0]) if argv else Path(tempfile.gettempdir())
    salida.mkdir(parents=True, exist_ok=True)
    informe = salida / "selftest.txt"
    lineas: list[str] = []
    try:
        os.environ.setdefault("APPDATA", str(salida / "appdata"))
        db = salida / "selftest.db"
        if db.exists():
            db.unlink()
        for modo in ("Claro", "Oscuro"):
            cfg = leer_config()
            cfg["tema"] = modo
            guardar_config(cfg)
            app = App(str(db))
            app._aviso_garantias = lambda: None
            i = app.bd.insertar({"persona": "Prueba", "nombre_equipo": "PC-TEST", "num_serie": "SN-TEST",
                                 "fin_garantia": "2030-01-01"})
            app.bd.actualizar_campo(i, "num_serie", "SN-2", esperado="SN-TEST")
            assert app.bd.obtener(i)["num_serie"] == "SN-2"
            app.bd.deshacer()
            assert app.bd.obtener(i)["num_serie"] == "SN-TEST", "deshacer"
            app.recargar()
            xlsx = salida / "selftest.xlsx"
            exportar_excel(xlsx, app.equipos)
            filas = leer_excel(xlsx).filas
            assert any(f["num_serie"] == "SN-TEST" for f in filas), filas
            assert leer_texto_tabla("PERSONA\tNº DE SERIE\nAna\tX1").filas[0]["num_serie"] == "X1"
            generar_informe(salida / "selftest_informe.pdf", app.equipos, "prueba", "selftest")
            from .informes import FORMATOS_ETIQUETA, generar_etiquetas
            generar_etiquetas(salida / "selftest_etiquetas.pdf", app.equipos, FORMATOS_ETIQUETA[0])
            for k in range(3):
                app.mostrar(k)
                app.update()
            lineas.append(f"tema {modo}: OK ({app.bd.contar()} equipos, usuario {app.bd.usuario})")
            app.after(300, app._salir)
            app.mainloop()
        if os.name == "nt":
            from .detectar import detectar
            datos = detectar()
            lineas.append("detectar: " + ", ".join(f"{k}={v}" for k, v in datos.items()))
            assert datos.get("nombre_equipo") and datos.get("os"), datos
        informe.write_text("OK\n" + "\n".join(lineas) + "\n", encoding="utf-8")
        return 0
    except Exception:  # noqa: BLE001
        informe.write_text("ERROR\n" + "\n".join(lineas) + "\n" + traceback.format_exc(), encoding="utf-8")
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
        app.after(300, lambda: app.importar(argv[0]))
    app.mainloop()


if __name__ == "__main__":
    sys.exit(main())
