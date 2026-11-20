# Inventario de Equipos (InventarioPC)

Programa de escritorio para gestionar el inventario de ordenadores de la empresa.
Funciona en cualquier Windows 10/11 de 64 bits **sin tener Python instalado**, con instalador o como un
único `InventarioPC.exe` portable (basta con copiarlo y hacer doble clic).

## Descargar

Cada vez que se sube código, GitHub Actions compila el programa en Windows, lo prueba y lo publica en
**Releases** («Inventario de Equipos X.Y.Z (compilación N)») con dos archivos:

* `InventarioPC-Setup-X.Y.Z.exe` — **instalador** (recomendado). No necesita permisos de administrador.
  Crea accesos directos en el escritorio y el menú Inicio, añade «Abrir con» para los `.xlsx` (sin
  quitarle a Excel la asociación) e incluye desinstalador.
* `InventarioPC.exe` — **portable**: un único archivo que se ejecuta sin instalar.

Ninguno necesita Python. Los datos (`%APPDATA%\InventarioPC`) se conservan al actualizar o desinstalar.

> **SmartScreen**: mientras el programa no esté firmado, Windows puede avisar la primera vez
> («Más información» → «Ejecutar de todas formas»). Para firmarlo hace falta un certificado de firma de
> código (de pago, emitido por una entidad como Sectigo, DigiCert, Certum…). Basta con añadir al
> repositorio los secretos `CERT_PFX_BASE64` (el `.pfx` en base64) y `CERT_PASSWORD`: la compilación
> firma automáticamente el `.exe` y el instalador.

Para publicar una versión «oficial» se crea una etiqueta `vX.Y.Z` (con la versión de `inventario/__init__.py`).

## Funciones

### Inventario
* Alta, edición, duplicado y baja de equipos. **Todos los campos son editables** y **ninguno actúa como
  clave**: cada equipo tiene un ID interno y puede haber varios con el mismo nº de serie, nombre, etc.
* **Edición en la propia tabla**: doble clic en una celda (o F2); Intro guarda, Tab pasa al siguiente campo.
* **Ficha del equipo** con tres pestañas: *Datos*, *Historial* (todos sus cambios, con quién y cuándo) y
  *Adjuntos* (facturas, garantías, fotos… guardados en una carpeta junto a la base de datos).
* **Deshacer / rehacer** (Ctrl+Z / Ctrl+Y) de cualquier operación: ediciones, bajas, cambios masivos,
  importaciones completas…
* **Editar un campo en muchos equipos a la vez** y cambio rápido de estado.
* **Detectar este equipo**: rellena nombre, usuario, marca, modelo, Windows, CPU, RAM y tipo, discos,
  nº de serie, pantalla, Office, antivirus e ID de TeamViewer leyendo los datos de Windows.
* Búsqueda instantánea, filtros por delegación, marca, SO, estado y garantía, orden por cualquier columna.
* **Vistas guardadas**: guarda una combinación de filtros con nombre y aplícala con un clic desde la barra lateral.
* **Columnas**: se recuerdan anchos y orden; se reordenan arrastrando las cabeceras y se pueden ocultar.

### Excel
* **Importar** `.xlsx`/`.csv` con el formato de la plantilla (`tests/fixtures/plantilla_original.xlsx`):
  detecta la fila de cabeceras, tolera cambios de orden, tildes y erratas (`Nombre quipo`,
  `EXPIACION GARANTIA`…), vista previa y tres modos (añadir, saltar filas idénticas en todos los campos,
  reemplazar con copia previa).
* **Pegar desde Excel**: copia filas en Excel y pulsa Ctrl+V en la tabla (con o sin cabeceras).
* **Exportar** con el mismo formato de la plantilla (vista filtrada, todo o selección), CSV y plantilla vacía.

### Garantías, informes y etiquetas
* Avisos de garantía con **días configurables** (Preferencias), filas en rojo/ámbar y aviso al abrir.
* En *Resumen*, listado de garantías por periodo (caducadas y próximas, ya caducadas, **este trimestre**,
  próximo trimestre, este año, próximo año) exportable a Excel.
* **Informe PDF** con cifras, gráficos, garantías que requieren atención y listado completo.
* **Etiquetas con código QR** en PDF para varios formatos de hoja A4 (Apli/Avery) e impresoras de
  etiquetas (Brother, DYMO). Al **escanear el QR** con un lector (o escribirlo en el buscador) se abre la
  ficha del equipo; desde un móvil se lee el nombre, nº de serie y persona.

### Varias personas a la vez
* La base de datos puede estar en una **carpeta compartida de red** (*Archivo → Usar otra base de datos*).
* Los cambios de otras personas aparecen solos (intervalo configurable).
* Al guardar sólo se aplican los campos que has cambiado tú; si otra persona ha cambiado **los mismos
  campos** mientras editabas, el programa avisa y deja elegir qué valor conservar (nunca se sobrescribe sin avisar).
* El **historial** registra el usuario de Windows (o el nombre que se indique en Preferencias).

### Apariencia
* Diseño claro y **modo oscuro** (automático según Windows, o forzado en Preferencias).
* En pantallas pequeñas la barra lateral se compacta a iconos y los filtros pasan a una segunda fila.

Base de datos SQLite en `%APPDATA%\InventarioPC\inventario.db`, con **copia automática diaria** (se
guardan las 10 últimas) y copias/restauración manuales.

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

El resultado queda en `dist\InventarioPC.exe`. Para crear el instalador, con
[Inno Setup 6](https://jrsoftware.org/isinfo.php) instalado:

```bash
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" /DMyAppVersion=1.1.0 installer\InventarioPC.iss
```

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
| `inventario/gui.py` | Ventana principal (Tkinter) |
| `inventario/dialogos.py` | Ficha del equipo, importación, etiquetas, preferencias… |
| `inventario/tema.py` | Tema visual claro/oscuro dibujado sin dependencias |
| `inventario/detectar.py` | Detección de los datos del equipo (Windows: CIM/registro) |
| `inventario/informes.py` | Informe PDF y etiquetas QR |
| `inventario/pdf.py` | Generador de PDF mínimo |
| `installer/InventarioPC.iss` | Instalador (Inno Setup) |
| `InventarioPC.spec` | Configuración de PyInstaller |
| `.github/workflows/build.yml` | Compilación automática del .exe en Windows |
