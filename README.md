# Inventario de Equipos (InventarioPC)

Programa de escritorio para gestionar el inventario de ordenadores de la empresa.
Se distribuye como **un único `InventarioPC.exe`**: funciona en cualquier Windows 10/11 de 64 bits
**sin tener Python instalado** y sin instalación (basta con copiarlo y hacer doble clic).

## Descargar el .exe

Cada vez que se sube código, GitHub Actions compila el ejecutable en Windows, lo prueba y lo publica:

* **Releases** del repositorio → última «Compilación N» → `InventarioPC.exe`
* o bien **Actions** → «Compilar InventarioPC.exe» → artefacto `InventarioPC-windows`.

> La primera vez Windows SmartScreen puede avisar de que el programa no está firmado:
> «Más información» → «Ejecutar de todas formas».

## Funciones

* Alta, edición, duplicado y baja de equipos (formulario con los mismos bloques que el Excel:
  *Propietario del ordenador* y *Datos del ordenador*, más Estado y Observaciones).
* **Importar Excel** (`.xlsx`) con el formato de la plantilla (`tests/fixtures/plantilla_original.xlsx`),
  y también `.csv`:
  * Detecta sola la fila de cabeceras (en la plantilla, fila 3) y los datos que van debajo.
  * Reconoce las columnas aunque cambien de orden, de mayúsculas o tengan tildes/erratas
    (`Nombre quipo`, `EXPIACION GARANTIA`…).
  * Vista previa antes de importar y detección de equipos ya existentes por **Nº de serie**
    (o nombre de equipo si no hay nº de serie): actualizar, completar sólo campos vacíos,
    omitir o duplicar.
  * Fechas en formato fecha de Excel o texto (`dd/mm/aaaa`, `aaaa-mm-dd`…).
* **Exportar a Excel** con el mismo formato de la plantilla (grupos en filas 1‑2, cabeceras en 3‑4,
  datos desde la fila 5, autofiltro, garantías coloreadas). Exporta la vista filtrada, todo o la selección.
  También a CSV y plantilla vacía.
* Búsqueda instantánea en todos los campos, filtros por delegación, marca, SO, estado y garantía,
  orden por cualquier columna, columnas visibles configurables, copiar filas al portapapeles.
* **Control de garantías**: filas en rojo (caducada) y amarillo (caduca en ≤ 90 días), aviso al abrir,
  botón «Compra + N años» para calcular la expiración.
* **Resumen**: totales y recuentos por delegación, marca, SO, estado, Office, antivirus, RAM y año de compra.
* **Historial** de todos los cambios (quién/qué campo cambió, importaciones, exportaciones) exportable a Excel.
* Base de datos SQLite en `%APPDATA%\InventarioPC\inventario.db`, con **copia automática diaria**
  (se guardan las 10 últimas) y copias/restauración manuales. Se puede usar una base de datos en una
  carpeta compartida de red (*Archivo → Usar otra base de datos*).

## Ejecutar desde el código

```bash
pip install -r requirements.txt
python app.py
```

## Compilar el .exe en un Windows propio

Requiere Python 3.10+ (sólo en el equipo que compila). Ejecuta `build.bat` o:

```bash
pip install -r requirements-build.txt
python -m PyInstaller --noconfirm --clean InventarioPC.spec
```

El resultado queda en `dist\InventarioPC.exe`.

## Tests

```bash
python -m pytest -q tests
```

## Estructura

| Archivo | Contenido |
|---|---|
| `app.py` | Punto de entrada |
| `inventario/campos.py` | Definición de columnas, alias de cabeceras y fechas |
| `inventario/db.py` | Base de datos SQLite, historial y copias de seguridad |
| `inventario/excel_io.py` | Importación/exportación Excel y CSV |
| `inventario/gui.py` | Interfaz gráfica (Tkinter) |
| `InventarioPC.spec` | Configuración de PyInstaller |
| `.github/workflows/build.yml` | Compilación automática del .exe en Windows |
