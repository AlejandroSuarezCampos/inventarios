"""Diálogos de la aplicación."""
from __future__ import annotations

import tkinter as tk
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .campos import CAMPOS, CAMPOS_POR_CLAVE, CLAVES, ESTADOS, fecha_a_date, mostrar_fecha, parsear_fecha
from .db import SinConexion
from .excel_io import ResultadoLectura, leer_excel
from .tema import C, MODOS
from .ui_comun import (TITULO, Dialogo, abrir_archivo, en_segundo_plano, tabla_con_scroll, tarjeta)


class Cancelado(Exception):
    pass


# =====================================================================
#  Ficha de equipo: Datos · Historial · Adjuntos
# =====================================================================
class FormularioEquipo(Dialogo):
    def __init__(self, app, equipo: dict | None = None, titulo: str = "Nuevo equipo",
                 detectar_al_abrir: bool = False, pestana: str = "datos", solo_lectura: bool = False):
        sub = ""
        if equipo and equipo.get("id"):
            sub = " · ".join(x for x in (equipo.get("persona"), equipo.get("delegacion"), equipo.get("num_serie")) if x)
        if solo_lectura:
            sub = "Sin conexión con el servidor: sólo lectura" + (f" · {sub}" if sub else "")
        elif (equipo or {}).get("_pendiente"):
            sub = "Pendiente de enviar al servidor" + (f" · {sub}" if sub else "")
        super().__init__(app, titulo, sub or "Rellena los datos del equipo y de la persona que lo usa.")
        if solo_lectura:
            self.lbl_subtitulo.configure(foreground=C["ambar_texto"])
        self.solo_lectura = solo_lectura
        self.equipo = equipo
        self.equipo_id = (equipo or {}).get("id")
        self.resultado: dict | None = None
        self.resizable(True, True)
        self.widgets: dict[str, tk.Widget] = {}
        self.vars: dict[str, tk.StringVar] = {}
        px = app.tema.px

        info = ""
        if self.equipo_id and self.equipo_id > 0:
            info = (f"ID {self.equipo_id} · creado {equipo.get('creado', '')[:16]} · "
                    f"modificado {equipo.get('modificado', '')[:16]}")
        pie = self.pie(None if solo_lectura else "Guardar", self._guardar, "",
                       texto_cancelar="Cerrar" if solo_lectura else "Cancelar")
        self.btn_detectar = ttk.Button(pie, text="⌖  Detectar este equipo", style="Plano.TButton",
                                       command=self.detectar)
        self.btn_detectar.pack(side="left")
        if info:
            ttk.Label(pie, text=info, style="Secundario.TLabel").pack(side="left", padx=(px(10), 0))

        # pestañas (sólo para equipos ya guardados)
        self.paginas: dict[str, ttk.Frame] = {}
        self.botones_pest: dict[str, ttk.Button] = {}
        if solo_lectura:
            self.btn_detectar.pack_forget()
        if self.equipo_id:
            barra = ttk.Frame(self.cuerpo)
            barra.pack(fill="x", pady=(0, px(12)))
            n_hist = len(app.bd.historial(500, self.equipo_id))
            n_adj = app.bd.contar_adjuntos(self.equipo_id)
            for clave, texto in (("datos", "Datos"), ("historial", f"Historial ({n_hist})"),
                                 ("adjuntos", f"Adjuntos ({n_adj})")):
                b = ttk.Button(barra, text=texto, style="Segmento.TButton", command=lambda c=clave: self.mostrar(c))
                b.pack(side="left", padx=(0, px(4)))
                self.botones_pest[clave] = b
        self.contenedor = ttk.Frame(self.cuerpo)
        self.contenedor.pack(fill="both", expand=True)

        self.paginas["datos"] = self._pagina_datos(self.contenedor)
        if self.equipo_id:
            self.paginas["historial"] = self._pagina_historial(self.contenedor)
            self.paginas["adjuntos"] = self._pagina_adjuntos(self.contenedor)
        self.mostrar("datos")

        if solo_lectura:
            for w in self.widgets.values():
                try:
                    w.configure(state="disabled")
                except tk.TclError:
                    pass
            for b in self._botones_fecha:
                b.state(["disabled"])
        else:
            self.bind("<Return>", lambda e: self._guardar() if not isinstance(e.widget, (tk.Text, ttk.Treeview)) else None)
        self.widgets["persona"].focus_set()
        self.centrar()
        if self.equipo_id:  # fijar el tamaño para que las pestañas no cambien la ventana
            self.update_idletasks()
            self.contenedor.configure(width=self.contenedor.winfo_width(), height=self.contenedor.winfo_height())
            self.contenedor.pack_propagate(False)
        if pestana != "datos" and pestana in self.paginas:
            self.mostrar(pestana)
        if detectar_al_abrir:
            self.after(150, self.detectar)

    # -------------------------------------------------------------- pestañas
    def mostrar(self, clave: str):
        for c, p in self.paginas.items():
            if c == clave:
                p.pack(fill="both", expand=True)
            else:
                p.pack_forget()
        for c, b in self.botones_pest.items():
            b.state(["selected"] if c == clave else ["!selected"])

    def _pagina_datos(self, padre) -> ttk.Frame:
        px = self.app.tema.px
        app = self.app
        equipo = self.equipo
        rejilla = ttk.Frame(padre)
        rejilla.columnconfigure(0, weight=1, uniform="c")
        rejilla.columnconfigure(1, weight=2, uniform="c")
        izq = ttk.Frame(rejilla)
        izq.grid(row=0, column=0, sticky="nsew", padx=(0, px(16)))
        der = ttk.Frame(rejilla)
        der.grid(row=0, column=1, sticky="nsew")
        rejilla.rowconfigure(0, weight=1)

        self._botones_fecha: list[ttk.Widget] = []
        secciones = {}
        for cont, clave, titulo_sec, expandir in ((izq, "propietario", "PROPIETARIO DEL ORDENADOR", False),
                                                  (izq, "otros", "ESTADO Y OBSERVACIONES", True),
                                                  (der, "equipo", "DATOS DEL ORDENADOR", True)):
            ttk.Label(cont, text=titulo_sec, style="Seccion.TLabel").pack(anchor="w", padx=px(4), pady=(0, px(6)))
            t, interior = tarjeta(cont)
            t.pack(fill="both", expand=expandir, pady=(0, px(12)))
            secciones[clave] = interior

        contadores = {"propietario": 0, "equipo": 0, "otros": 0}
        for campo in CAMPOS:
            marco = secciones[campo.grupo]
            idx = contadores[campo.grupo]
            contadores[campo.grupo] += 1
            r, c = (idx // 2, idx % 2) if campo.grupo == "equipo" else (idx, 0)
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
                b_hoy = ttk.Button(fila, text="Hoy", style="Blanco.Plano.TButton",
                                   command=lambda v=var: v.set(date.today().strftime("%d/%m/%Y")))
                b_hoy.pack(side="left", padx=(4, 0))
                self._botones_fecha.append(b_hoy)
                if campo.clave == "fin_garantia":
                    mb = ttk.Menubutton(fila, text="+ años", style="Blanco.Plano.TButton")
                    menu = tk.Menu(mb, tearoff=False)
                    for anios in (1, 2, 3, 4, 5):
                        menu.add_command(label=f"Fecha de compra + {anios} año{'s' if anios > 1 else ''}",
                                         command=lambda a=anios: self._garantia_desde_compra(a))
                    mb["menu"] = menu
                    mb.pack(side="left")
                    self._botones_fecha.append(mb)
                self.vars[campo.clave] = var
            else:
                var = tk.StringVar(value=valor or ("Activo" if campo.clave == "estado" else ""))
                valores = ESTADOS if campo.clave == "estado" else app.bd.valores_distintos(campo.clave)
                w = ttk.Combobox(celda, textvariable=var, width=24, values=valores, style="Blanco.TCombobox")
                w.pack(fill="x")
                self.vars[campo.clave] = var
            self.widgets[campo.clave] = w
        return rejilla

    def _pagina_historial(self, padre) -> ttk.Frame:
        px = self.app.tema.px
        p = ttk.Frame(padre)
        t, marco = tarjeta(p, padding=px(6))
        t.pack(fill="both", expand=True)
        tv = tabla_con_scroll(marco, columns=("fecha", "usuario", "accion", "detalle"), show="headings")
        for c, txt, w in (("fecha", "FECHA", 140), ("usuario", "USUARIO", 140), ("accion", "ACCIÓN", 120),
                          ("detalle", "DETALLE", 560)):
            tv.heading(c, text=txt, anchor="w")
            tv.column(c, width=px(w), stretch=c == "detalle")
        for accion, color in (("Alta", C["verde"]), ("Modificación", C["acento"]), ("Baja (eliminado)", C["rojo"]),
                              ("Adjunto añadido", C["morado"]), ("Adjunto eliminado", C["secundario"])):
            tv.tag_configure(accion, foreground=color)
        hist = self.app.bd.historial(500, self.equipo_id)
        for h in hist:
            tv.insert("", "end", tags=(h["accion"],),
                      values=(h["fecha"], h.get("usuario", ""), h["accion"], h["detalle"].replace("\n", " ")))
        if not hist:
            ttk.Label(p, text="Este equipo todavía no tiene cambios registrados.", style="Secundario.TLabel").pack(
                anchor="w", pady=(px(6), 0))
        ttk.Label(p, text="Pasa el ratón por encima o amplía la ventana para leer el detalle completo.",
                  style="Secundario.TLabel").pack(anchor="w", pady=(px(6), 0))
        self._tv_hist = tv
        return p

    def _pagina_adjuntos(self, padre) -> ttk.Frame:
        px = self.app.tema.px
        p = ttk.Frame(padre)
        barra = ttk.Frame(p)
        barra.pack(fill="x", pady=(0, px(8)))
        b_anadir = ttk.Button(barra, text="＋  Añadir archivos…", style="Acento.TButton", command=self._adj_anadir)
        b_anadir.pack(side="left")
        ttk.Button(barra, text="Abrir", command=self._adj_abrir).pack(side="left", padx=(px(8), 0))
        if not self.app.bd.remoto:
            ttk.Button(barra, text="Abrir carpeta", command=self._adj_carpeta).pack(side="left", padx=(px(8), 0))
        b_quitar = ttk.Button(barra, text="Quitar", style="Peligro.TButton", command=self._adj_quitar)
        b_quitar.pack(side="right")
        if self.solo_lectura:
            b_anadir.state(["disabled"])
            b_quitar.state(["disabled"])
        t, marco = tarjeta(p, padding=px(6))
        t.pack(fill="both", expand=True)
        tv = tabla_con_scroll(marco, columns=("nombre", "tipo", "tamano", "fecha"), show="headings")
        for c, txt, w, anc in (("nombre", "ARCHIVO", 380, "w"), ("tipo", "TIPO", 80, "w"),
                               ("tamano", "TAMAÑO", 90, "e"), ("fecha", "MODIFICADO", 140, "w")):
            tv.heading(c, text=txt, anchor=anc)
            tv.column(c, width=px(w), anchor=anc, stretch=c == "nombre")
        tv.bind("<Double-1>", lambda e: self._adj_abrir())
        tv.bind("<Delete>", lambda e: self._adj_quitar())
        donde = ("en el servidor central" if self.app.bd.remoto else
                 f"en la carpeta «{self.app.bd.carpeta_adjuntos().name}» junto a la base de datos")
        ttk.Label(p, text=f"Facturas, documentos de garantía, fotos… Se guardan {donde}.",
                  style="Secundario.TLabel").pack(anchor="w", pady=(px(6), 0))
        self._tv_adj = tv
        self._pintar_adjuntos()
        return p

    def _pintar_adjuntos(self):
        tv = self._tv_adj
        tv.delete(*tv.get_children())
        archivos = self.app.bd.info_adjuntos(self.equipo_id)
        for a in archivos:
            tam = int(a.get("tamano", 0))
            tam_txt = f"{tam / 1024 / 1024:.1f} MB" if tam > 1024 * 1024 else f"{max(1, tam // 1024)} KB"
            tv.insert("", "end", iid=a["nombre"], values=(
                a["nombre"], Path(a["nombre"]).suffix.lstrip(".").upper(), tam_txt,
                datetime.fromtimestamp(a.get("fecha", 0)).strftime("%d/%m/%Y %H:%M")))
        if "adjuntos" in self.botones_pest:
            self.botones_pest["adjuntos"].configure(text=f"Adjuntos ({len(archivos)})")

    def _adj_anadir(self):
        rutas = filedialog.askopenfilenames(parent=self, title="Añadir adjuntos")
        for r in rutas:
            try:
                self.app.bd.anadir_adjunto(self.equipo_id, r)
            except SinConexion as exc:
                messagebox.showwarning(TITULO, str(exc), parent=self)
                break
            except OSError as exc:
                messagebox.showerror(TITULO, f"No se pudo copiar «{Path(r).name}»:\n{exc}", parent=self)
        self._pintar_adjuntos()

    def _adj_sel(self) -> list[str]:
        return list(self._tv_adj.selection())

    def _adj_abrir(self):
        for nombre in self._adj_sel()[:5]:
            try:
                abrir_archivo(self.app.bd.ruta_adjunto(self.equipo_id, nombre))
            except SinConexion as exc:
                messagebox.showwarning(TITULO, str(exc), parent=self)
                return

    def _adj_carpeta(self):
        carpeta = self.app.bd.carpeta_adjuntos(self.equipo_id)
        carpeta.mkdir(parents=True, exist_ok=True)
        abrir_archivo(carpeta)

    def _adj_quitar(self):
        sel = self._adj_sel()
        if not sel:
            return
        if messagebox.askyesno(TITULO, f"¿Quitar {len(sel)} adjunto(s) de este equipo?\n\n"
                                       "Se mueven a la papelera de adjuntos, no se borran definitivamente.", parent=self):
            try:
                for nombre in sel:
                    self.app.bd.eliminar_adjunto(self.equipo_id, nombre)
            except SinConexion as exc:
                messagebox.showwarning(TITULO, str(exc), parent=self)
            self._pintar_adjuntos()

    # -------------------------------------------------------------- detección
    def detectar(self):
        from .detectar import detectar
        self.btn_detectar.configure(text="Detectando…")
        self.btn_detectar.state(["disabled"])
        self.configure(cursor="watch")

        def fin(datos):
            self._restaurar_boton()
            if not datos:
                messagebox.showwarning(TITULO, "No se ha podido detectar ningún dato.", parent=self)
                return
            d = DialogoDetectar(self.app, datos, self.datos(), parent=self)
            self.wait_window(d)
            for clave, valor in (d.elegidos or {}).items():
                if clave in self.vars:
                    self.vars[clave].set(mostrar_fecha(valor) if CAMPOS_POR_CLAVE[clave].tipo == "fecha" else valor)

        def error(exc):
            self._restaurar_boton()
            messagebox.showerror(TITULO, f"No se pudo detectar el equipo:\n{exc}", parent=self)
        en_segundo_plano(self, detectar, fin, error)

    def _restaurar_boton(self):
        try:
            self.configure(cursor="")
            self.btn_detectar.configure(text="⌖  Detectar este equipo")
            self.btn_detectar.state(["!disabled"])
        except tk.TclError:
            pass

    # -------------------------------------------------------------- guardar
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
                    self.mostrar("datos")
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


class DialogoDetectar(Dialogo):
    """Muestra los datos detectados y deja elegir cuáles copiar a la ficha."""

    def __init__(self, app, detectado: dict, actuales: dict, parent=None):
        super().__init__(app, "Datos detectados", "Marca los campos que quieres rellenar con lo detectado.")
        if parent is not None:
            self.transient(parent)
        self.elegidos: dict | None = None
        px = app.tema.px
        self.vars: dict[str, tk.BooleanVar] = {}

        def aceptar():
            self.elegidos = {k: detectado[k] for k, v in self.vars.items() if v.get()}
            self.destroy()
        self.pie("Rellenar campos", aceptar)
        t, inn = tarjeta(self.cuerpo)
        t.pack(fill="both", expand=True)
        ttk.Label(inn, text="CAMPO", style="EnTarjetaSec.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(inn, text="DETECTADO", style="EnTarjetaSec.TLabel").grid(row=0, column=1, sticky="w", padx=px(12))
        ttk.Label(inn, text="VALOR ACTUAL", style="EnTarjetaSec.TLabel").grid(row=0, column=2, sticky="w")
        fila = 1
        for c in CAMPOS:
            if c.clave not in detectado:
                continue
            nuevo = detectado[c.clave]
            actual = actuales.get(c.clave, "")
            v = tk.BooleanVar(value=(not actual) or actual == nuevo)
            ttk.Checkbutton(inn, text=c.etiqueta, variable=v).grid(row=fila, column=0, sticky="w", pady=1)
            ttk.Label(inn, text=nuevo, style="EnTarjeta.TLabel").grid(row=fila, column=1, sticky="w", padx=px(12))
            ttk.Label(inn, text=actual or "—", style="EnTarjetaSec.TLabel").grid(row=fila, column=2, sticky="w")
            self.vars[c.clave] = v
            fila += 1
        self.centrar()


# =====================================================================
#  Importación (archivo o portapapeles)
# =====================================================================
class DialogoImportar(Dialogo):
    """Importa filas. Ningún campo se usa como clave: cada fila es un equipo."""
    OPCIONES = [
        ("omitir_identicas", "Añadir las filas, saltando las que sean idénticas (en todos los campos) "
                             "a un equipo ya guardado"),
        ("anadir", "Añadir todas las filas como equipos nuevos"),
        ("reemplazar", "Reemplazar todo el inventario actual por estas filas"),
    ]

    def __init__(self, app, ruta: Path | None = None, res: ResultadoLectura | None = None,
                 nombre: str | None = None):
        if res is None:
            try:
                res = leer_excel(ruta)
            except Exception as exc:  # noqa: BLE001
                messagebox.showerror(TITULO, f"No se pudo leer el archivo:\n{ruta}\n\n{exc}", parent=app)
                raise Cancelado from exc
        if not res.filas:
            avisos = "\n".join(res.avisos) or "No hay filas con datos."
            messagebox.showwarning(TITULO, f"No hay equipos que importar.\n\n{avisos}", parent=app)
            raise Cancelado
        self.nombre = nombre or (ruta.name if ruta else "Portapapeles")
        titulo = "Pegar desde Excel" if ruta is None else "Importar inventario"
        super().__init__(app, titulo, f"{self.nombre} · {', '.join(res.hojas_leidas)}")
        self.res = res
        self.importado = False
        self.ids_nuevos: list[int] = []
        px = app.tema.px
        self.geometry(f"{px(1040)}x{px(680)}")
        self.minsize(px(720), px(480))

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

        self.pie("Importar", self._importar)

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
            notas.append(("!", C["ambar"], f"{len(self.repetidas)} fila(s) idénticas a otra fila de los datos."))
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
        t, op = tarjeta(self.cuerpo, padding=(px(14), px(8)))
        t.pack(fill="x")
        self.modo = tk.StringVar(value="omitir_identicas" if (self.identicas or self.repetidas) else "anadir")
        for valor, texto in self.OPCIONES:
            rb = ttk.Radiobutton(op, text=texto, value=valor, variable=self.modo)
            rb.pack(anchor="w")
            if valor == "reemplazar" and not app.bd.en_linea:
                rb.state(["disabled"])
        if not app.bd.en_linea:
            ttk.Label(op, text="Sin conexión con el servidor: los equipos se guardarán como pendientes y se "
                               "enviarán al servidor cuando vuelva la conexión.", style="EnTarjetaSec.TLabel",
                      wraplength=px(900)).pack(anchor="w", pady=(px(4), 0))

        ttk.Label(self.cuerpo, text="VISTA PREVIA", style="Seccion.TLabel").pack(anchor="w", padx=px(4),
                                                                               pady=(px(14), px(6)))
        t, marco = tarjeta(self.cuerpo, padding=px(6))
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
                        f"{len(self.res.filas)} filas importadas.\n(Antes se guarda una copia de seguridad y "
                        "también se puede deshacer con Ctrl+Z.)\n\n¿Continuar?", icon="warning", parent=self):
            return
        nuevos = omitidos = 0
        try:
            if modo == "reemplazar":
                bd.copia_seguridad(bd.ruta.parent / "copias" /
                                   f"antes_de_importar_{datetime.now():%Y%m%d_%H%M%S}.db")
            with bd.operacion(f"Importar {self.nombre}"):
                if modo == "reemplazar":
                    bd.vaciar(f"Reemplazado por {self.nombre}")
                vistas = bd.firmas() if modo == "omitir_identicas" else {}
                for fila in self.res.filas:
                    datos = {k: v for k, v in fila.items() if not k.startswith("_")}
                    if modo == "omitir_identicas":
                        f = bd.firma(datos)
                        if f in vistas:
                            omitidos += 1
                            continue
                        vistas[f] = 0
                    self.ids_nuevos.append(bd.insertar(datos))
                    nuevos += 1
                bd.registrar(None, self.nombre, "Importación",
                             f"{nuevos} añadidos, {omitidos} omitidos por ser idénticos"
                             + (" (inventario reemplazado)" if modo == "reemplazar" else ""))
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror(TITULO, f"Error durante la importación. No se ha guardado nada.\n\n{exc}", parent=self)
            return
        if None in self.ids_nuevos:  # con servidor los ID llegan al terminar la operación
            self.ids_nuevos = list(getattr(bd, "ultimos_ids", []))
        self.importado = True
        self.destroy()
        texto = f"Importación completada.\n\nEquipos añadidos: {nuevos}"
        if bd.remoto and not bd.en_linea:
            texto += " (pendientes de enviar al servidor)"
        if omitidos:
            texto += f"\nFilas idénticas omitidas: {omitidos}"
        texto += "\n\nSi algo no ha ido como esperabas, Ctrl+Z deshace la importación completa."
        messagebox.showinfo(TITULO, texto, parent=self.app)


# =====================================================================
#  Etiquetas QR
# =====================================================================
class DialogoEtiquetas(Dialogo):
    def __init__(self, app, equipos: list[dict]):
        from .informes import FORMATOS_ETIQUETA
        super().__init__(app, "Imprimir etiquetas", f"{len(equipos)} equipo{'s' if len(equipos) != 1 else ''} · "
                                                    "con código QR para identificarlos al escanearlos")
        self.equipos = equipos
        self.formatos = {f.nombre: f for f in FORMATOS_ETIQUETA}
        px = app.tema.px
        cfg = app.cfg
        self.var_formato = tk.StringVar(value=cfg.get("formato_etiqueta") if cfg.get("formato_etiqueta") in self.formatos
                                        else FORMATOS_ETIQUETA[0].nombre)
        self.var_saltar = tk.StringVar(value="0")
        self.var_guias = tk.BooleanVar(value=bool(cfg.get("guias_etiqueta", False)))
        self.pie("Crear PDF", self._crear)
        t, inn = tarjeta(self.cuerpo)
        t.pack(fill="both", expand=True)
        ttk.Label(inn, text="Formato de etiqueta", style="EnTarjetaCampo.TLabel").pack(anchor="w")
        cb = ttk.Combobox(inn, textvariable=self.var_formato, values=list(self.formatos), state="readonly",
                          width=56, style="Blanco.TCombobox")
        cb.pack(fill="x", pady=(px(2), px(10)))
        cb.bind("<<ComboboxSelected>>", lambda e: self._resumen())
        fila = ttk.Frame(inn, style="EnTarjeta.TFrame")
        fila.pack(fill="x")
        ttk.Label(fila, text="Saltar las primeras", style="EnTarjetaCampo.TLabel").pack(side="left")
        sp = ttk.Combobox(fila, values=[str(i) for i in range(80)], width=4, textvariable=self.var_saltar,
                          style="Blanco.TCombobox")
        sp.pack(side="left", padx=px(6))
        sp.bind("<<ComboboxSelected>>", lambda e: self._resumen())
        sp.bind("<KeyRelease>", lambda e: self._resumen())
        ttk.Label(fila, text="etiquetas (para aprovechar una hoja ya empezada)", style="EnTarjetaSec.TLabel").pack(
            side="left")
        ttk.Checkbutton(inn, text="Dibujar el contorno de cada etiqueta (útil para recortar en papel normal)",
                        variable=self.var_guias).pack(anchor="w", pady=(px(10), 0))
        self.lbl = ttk.Label(inn, text="", style="EnTarjetaSec.TLabel")
        self.lbl.pack(anchor="w", pady=(px(10), 0))
        self._resumen()
        self.centrar()

    def _saltar(self) -> int:
        try:
            return max(0, int(self.var_saltar.get()))
        except ValueError:
            return 0

    def _resumen(self):
        f = self.formatos[self.var_formato.get()]
        por_hoja = f.columnas * f.filas
        total = len(self.equipos) + self._saltar()
        hojas = -(-total // por_hoja)
        unidad = "etiqueta" if por_hoja == 1 else "hoja"
        self.lbl.configure(text=f"{len(self.equipos)} etiquetas de {f.ancho:g}×{f.alto:g} mm · {hojas} {unidad}"
                                f"{'s' if hojas != 1 else ''}. Imprime el PDF al 100 % (sin «ajustar a página»).")

    def _crear(self):
        from .informes import generar_etiquetas
        destino = filedialog.asksaveasfilename(parent=self, defaultextension=".pdf", filetypes=[("PDF", "*.pdf")],
                                               initialfile=f"etiquetas_{datetime.now():%Y%m%d_%H%M}.pdf")
        if not destino:
            return
        f = self.formatos[self.var_formato.get()]
        try:
            generar_etiquetas(destino, self.equipos, f, self._saltar(), self.var_guias.get())
        except PermissionError:
            messagebox.showerror(TITULO, "No se puede escribir el archivo. ¿Está abierto en otro programa?", parent=self)
            return
        self.app.cfg["formato_etiqueta"] = f.nombre
        self.app.cfg["guias_etiqueta"] = self.var_guias.get()
        self.app.guardar_cfg()
        self.destroy()
        abrir_archivo(Path(destino))


# =====================================================================
#  Preferencias
# =====================================================================
REFRESCOS = {"Desactivado": 0, "Cada 5 segundos": 5, "Cada 15 segundos": 15, "Cada 30 segundos": 30,
             "Cada minuto": 60}


class DialogoPreferencias(Dialogo):
    def __init__(self, app):
        super().__init__(app, "Preferencias", "Se guardan para tu usuario de Windows en este equipo.")
        self.cambios: dict = {}
        cfg = app.cfg
        px = app.tema.px
        self.var_dias = tk.StringVar(value=str(cfg.get("dias_aviso", 90)))
        self.var_avisar = tk.BooleanVar(value=cfg.get("avisar_garantias", True))
        self.var_usuario = tk.StringVar(value=cfg.get("usuario", ""))
        self.var_tema = tk.StringVar(value=cfg.get("tema", MODOS[0]))
        actual = cfg.get("refresco_seg", 5)
        self.var_refresco = tk.StringVar(value=next((k for k, v in REFRESCOS.items() if v == actual), "Cada 5 segundos"))
        self.pie("Guardar", self._guardar)

        def seccion(titulo):
            ttk.Label(self.cuerpo, text=titulo, style="Seccion.TLabel").pack(anchor="w", padx=px(4), pady=(0, px(6)))
            t, inn = tarjeta(self.cuerpo)
            t.pack(fill="x", pady=(0, px(12)))
            return inn

        inn = seccion("GARANTÍAS")
        f = ttk.Frame(inn, style="EnTarjeta.TFrame")
        f.pack(fill="x")
        ttk.Label(f, text="Avisar cuando falten", style="EnTarjeta.TLabel").pack(side="left")
        ttk.Combobox(f, textvariable=self.var_dias, values=["30", "60", "90", "120", "180", "365"], width=6,
                     style="Blanco.TCombobox").pack(side="left", padx=px(6))
        ttk.Label(f, text="días o menos para que caduque la garantía", style="EnTarjeta.TLabel").pack(side="left")
        ttk.Checkbutton(inn, text="Avisar al abrir el programa (una vez al día)", variable=self.var_avisar).pack(
            anchor="w", pady=(px(8), 0))

        inn = seccion("HISTORIAL")
        ttk.Label(inn, text="Nombre con el que se firman tus cambios", style="EnTarjetaCampo.TLabel").pack(anchor="w")
        ttk.Entry(inn, textvariable=self.var_usuario, width=40, style="Blanco.TEntry").pack(fill="x", pady=(px(2), 0))
        ttk.Label(inn, text=f"Si lo dejas vacío se usa el usuario de Windows ({_usuario_windows()}).",
                  style="EnTarjetaSec.TLabel").pack(anchor="w", pady=(px(4), 0))

        inn = seccion("VARIAS PERSONAS A LA VEZ")
        f = ttk.Frame(inn, style="EnTarjeta.TFrame")
        f.pack(fill="x")
        ttk.Label(f, text="Mostrar los cambios de otras personas", style="EnTarjeta.TLabel").pack(side="left")
        ttk.Combobox(f, textvariable=self.var_refresco, values=list(REFRESCOS), state="readonly", width=18,
                     style="Blanco.TCombobox").pack(side="left", padx=px(6))
        ttk.Label(inn, text="Útil si la base de datos está en una carpeta compartida de red.",
                  style="EnTarjetaSec.TLabel").pack(anchor="w", pady=(px(4), 0))

        inn = seccion("APARIENCIA")
        f = ttk.Frame(inn, style="EnTarjeta.TFrame")
        f.pack(fill="x")
        ttk.Label(f, text="Tema", style="EnTarjeta.TLabel").pack(side="left")
        ttk.Combobox(f, textvariable=self.var_tema, values=MODOS, state="readonly", width=16,
                     style="Blanco.TCombobox").pack(side="left", padx=px(6))
        ttk.Label(f, text="(se aplica al reiniciar el programa)", style="EnTarjetaSec.TLabel").pack(side="left")
        self.centrar()

    def _guardar(self):
        try:
            dias = int(self.var_dias.get())
            if not 1 <= dias <= 3650:
                raise ValueError
        except ValueError:
            messagebox.showerror(TITULO, "Los días de aviso deben ser un número entre 1 y 3650.", parent=self)
            return
        self.cambios = {"dias_aviso": dias, "avisar_garantias": self.var_avisar.get(),
                        "usuario": self.var_usuario.get().strip(), "tema": self.var_tema.get(),
                        "refresco_seg": REFRESCOS.get(self.var_refresco.get(), 5)}
        self.destroy()


def _usuario_windows() -> str:
    import getpass
    import socket
    try:
        return f"{getpass.getuser()}@{socket.gethostname()}"
    except Exception:  # noqa: BLE001
        return "usuario"


# =====================================================================
#  Edición de un campo en muchos equipos
# =====================================================================
class DialogoCampoMasivo(Dialogo):
    def __init__(self, app, n: int, clave: str | None = None):
        super().__init__(app, "Editar un campo en la selección",
                         f"El valor se aplicará a {n} equipo{'s' if n != 1 else ''}")
        self.resultado: tuple[str, str] | None = None
        etiquetas = [c.etiqueta for c in CAMPOS]
        self.var_campo = tk.StringVar(value=CAMPOS_POR_CLAVE[clave].etiqueta if clave in CAMPOS_POR_CLAVE
                                      else etiquetas[0])
        self.var = tk.StringVar()
        self.pie("Aplicar", self._aceptar)
        t, inn = tarjeta(self.cuerpo)
        t.pack(fill="both", expand=True)
        ttk.Label(inn, text="Campo", style="EnTarjetaCampo.TLabel").pack(anchor="w", pady=(0, 2))
        cb_campo = ttk.Combobox(inn, textvariable=self.var_campo, values=etiquetas, state="readonly", width=40,
                                style="Blanco.TCombobox")
        cb_campo.pack(fill="x", pady=(0, app.tema.px(10)))
        ttk.Label(inn, text="Nuevo valor (vacío = borrar el campo)", style="EnTarjetaCampo.TLabel").pack(
            anchor="w", pady=(0, 2))
        self.cb = ttk.Combobox(inn, textvariable=self.var, width=40, style="Blanco.TCombobox")
        self.cb.pack(fill="x")
        cb_campo.bind("<<ComboboxSelected>>", lambda e: self._valores())
        self._valores()
        self.cb.focus_set()
        self.bind("<Return>", lambda e: self._aceptar())
        self.centrar()

    def _clave(self) -> str:
        return next(c.clave for c in CAMPOS if c.etiqueta == self.var_campo.get())

    def _valores(self):
        k = self._clave()
        self.cb["values"] = ESTADOS if k == "estado" else self.app.bd.valores_distintos(k)

    def _aceptar(self):
        k = self._clave()
        valor = self.var.get().strip()
        if CAMPOS_POR_CLAVE[k].tipo == "fecha":
            valor = parsear_fecha(valor)
        self.resultado = (k, valor)
        self.destroy()


# =====================================================================
#  Columnas visibles
# =====================================================================
class DialogoColumnas(Dialogo):
    def __init__(self, app):
        super().__init__(app, "Columnas", "Elige qué columnas se muestran. Para cambiar el orden, arrastra la "
                                          "cabecera de una columna sobre otra en la tabla.")
        self.resultado: set[str] | None = None
        self.restablecer = False
        self.variables: dict[str, tk.BooleanVar] = {}
        pie = self.pie("Aplicar", self._aceptar)
        ttk.Button(pie, text="Restablecer orden y anchos", style="Plano.TButton",
                   command=self._restablecer).pack(side="left")
        t, inn = tarjeta(self.cuerpo)
        t.pack(fill="both", expand=True)
        for i, c in enumerate(CAMPOS):
            v = tk.BooleanVar(value=c.clave not in app.columnas_ocultas)
            ttk.Checkbutton(inn, text=c.etiqueta, variable=v).grid(row=i % 10, column=i // 10, sticky="w",
                                                                    padx=(0, app.tema.px(24)))
            self.variables[c.clave] = v
        self.centrar()

    def _aceptar(self):
        self.resultado = {k for k, v in self.variables.items() if not v.get()}
        self.destroy()

    def _restablecer(self):
        self.restablecer = True
        self.resultado = set()
        self.destroy()


# =====================================================================
#  Pregunta de texto sencilla
# =====================================================================
class DialogoTexto(Dialogo):
    def __init__(self, app, titulo: str, etiqueta: str, valor: str = "", subtitulo: str = ""):
        super().__init__(app, titulo, subtitulo)
        self.resultado: str | None = None
        self.var = tk.StringVar(value=valor)
        self.pie("Aceptar", self._aceptar)
        t, inn = tarjeta(self.cuerpo)
        t.pack(fill="both", expand=True)
        ttk.Label(inn, text=etiqueta, style="EnTarjetaCampo.TLabel").pack(anchor="w", pady=(0, 2))
        e = ttk.Entry(inn, textvariable=self.var, width=40, style="Blanco.TEntry")
        e.pack(fill="x")
        e.focus_set()
        e.select_range(0, "end")
        self.bind("<Return>", lambda ev: self._aceptar())
        self.centrar()

    def _aceptar(self):
        if self.var.get().strip():
            self.resultado = self.var.get().strip()
            self.destroy()


# =====================================================================
#  Conexión a un servidor central
# =====================================================================
class DialogoServidor(Dialogo):
    def __init__(self, app):
        super().__init__(app, "Conectar a un servidor central",
                         "Los datos se guardarán en el servidor. Sin conexión podrás consultar la última copia "
                         "y crear equipos nuevos.")
        self.resultado: tuple[str, str] | None = None
        cfg = app.cfg
        px = app.tema.px
        self.var_url = tk.StringVar(value=cfg.get("servidor", ""))
        self.var_clave = tk.StringVar(value=cfg.get("clave_servidor", ""))
        pie = self.pie("Conectar", self._conectar)
        ttk.Button(pie, text="Probar conexión", style="Plano.TButton", command=self._probar).pack(side="left")
        t, inn = tarjeta(self.cuerpo)
        t.pack(fill="both", expand=True)
        ttk.Label(inn, text="Dirección del servidor", style="EnTarjetaCampo.TLabel").pack(anchor="w")
        e = ttk.Entry(inn, textvariable=self.var_url, width=44, style="Blanco.TEntry")
        e.pack(fill="x", pady=(px(2), 0))
        ttk.Label(inn, text="Por ejemplo 192.168.1.20:8765 (aparece en la ventana del servidor)",
                  style="EnTarjetaSec.TLabel").pack(anchor="w", pady=(px(2), px(10)))
        ttk.Label(inn, text="Clave", style="EnTarjetaCampo.TLabel").pack(anchor="w")
        ttk.Entry(inn, textvariable=self.var_clave, width=44, style="Blanco.TEntry", show="•").pack(
            fill="x", pady=(px(2), 0))
        self.lbl = ttk.Label(inn, text="", style="EnTarjetaSec.TLabel", wraplength=px(420))
        self.lbl.pack(anchor="w", pady=(px(10), 0))
        e.focus_set()
        self.bind("<Return>", lambda ev: self._conectar())
        self.centrar()

    def _probar(self) -> bool:
        from .remoto import ErrorServidor, probar_conexion
        self.lbl.configure(text="Conectando…", foreground=C["secundario"])
        self.update_idletasks()
        try:
            info = probar_conexion(self.var_url.get(), self.var_clave.get())
        except (SinConexion, ErrorServidor, ValueError) as exc:
            texto = str(exc).split("\n")[0] if not isinstance(exc, SinConexion) else \
                "No se puede conectar. Comprueba la dirección, que el servidor esté en marcha y el cortafuegos."
            self.lbl.configure(text="✕  " + texto, foreground=C["rojo"])
            return False
        self.lbl.configure(text=f"✓  Conectado a «{info.get('servidor')}» · {info.get('equipos')} equipos · "
                                f"versión {info.get('version')}", foreground=C["verde"])
        return True

    def _conectar(self):
        if self._probar():
            from .remoto import normalizar_url
            self.resultado = (normalizar_url(self.var_url.get()), self.var_clave.get())
            self.destroy()
        elif messagebox.askyesno(TITULO, "Ahora no hay conexión con ese servidor.\n\n¿Guardar la configuración "
                                         "igualmente? Se trabajará sin conexión hasta que responda.", parent=self):
            from .remoto import normalizar_url
            try:
                self.resultado = (normalizar_url(self.var_url.get()), self.var_clave.get())
            except ValueError as exc:
                messagebox.showerror(TITULO, str(exc), parent=self)
                return
            self.destroy()
