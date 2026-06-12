# Extractor de PDF Semántico — Editor Semántico
### Herramienta de Automatización Editorial — Paleontología Mexicana, UNAM

Aplicación de escritorio desarrollada en Python para el equipo editorial de la revista **Paleontología Mexicana** (Instituto de Geología, UNAM). Convierte artículos científicos en PDF a HTML editorial, EPUB y XML JATS (orientado a SciELO SPS) con estilos tipográficos fieles al diseño de la revista, optimizando el flujo de trabajo de maquetación.

---

## ¿Qué hace?

- Extrae y clasifica automáticamente los bloques de texto de un PDF científico (títulos, resúmenes, cuerpo, referencias, tablas, figuras, etc.)
- Genera un HTML limpio con tipografía editorial (Source Serif 4, Times New Roman)
- Genera un EPUB 2.0 compatible con lectores como Calibre, Thorium Reader y Apple Books
- Genera XML en formato JATS orientado a SciELO SPS, con validación estructural integrada
- Vincula autores con sus perfiles ORCID
- Inserta tablas Excel y figuras en la posición exacta del texto mediante texto ancla
- Maneja artículos en dos columnas, guiones de corte tipográfico y saltos de página

---

## Arquitectura

La aplicación funciona como una **app de escritorio híbrida**:

```
main.py (punto de entrada)
  │
  ├── levanta server.py (FastAPI + uvicorn) en un hilo de fondo
  │     └── server.py usa los módulos de core/ para procesar el PDF
  │         y construir HTML / EPUB / XML
  │
  └── abre una ventana nativa con pywebview que carga
        http://127.0.0.1:8765 → static/index.html (interfaz SPA)
```

- **Backend (`server.py`)**: expone la lógica de `core/` como una API REST (carga de PDF, autores, afiliaciones, referencias, figuras, tablas, exportación con vista previa, validación XML, manejo de sesión).
- **Frontend (`static/`)**: una SPA en HTML/CSS/JS con navegación tipo *stepper* entre 6 secciones (PDF, Autores, Afiliaciones, Referencias, Figuras, Tablas) más un panel de Configuración, que consume la API mediante `fetch`.
- **`core/`**: lógica de negocio pura (sin dependencias de UI), reutilizable tanto por `server.py` como por la versión legacy de escritorio.

---

## Tecnologías y requisitos del sistema

El proyecto está desarrollado en **Python 3.13+**. Las dependencias principales son:

| Paquete | Uso |
|---|---|
| `fastapi` / `uvicorn` / `starlette` | Backend / servidor local |
| `pywebview` | Ventana nativa que embebe la interfaz web |
| `PyMuPDF` (fitz) | Lectura y extracción de texto/metadatos de PDF |
| `lxml` | Validación de XML JATS |
| `openpyxl` | Lectura de archivos Excel (autores, tablas) |
| `pydantic` | Validación de datos de la API |
| `pyinstaller` | Empaquetado del ejecutable `.exe` |

Consulta `requirements.txt` para la lista completa con versiones exactas.

---

## Instalación

### 1. Clonar el repositorio

```bash
git clone https://github.com/tu-usuario/Automatizaci-n-Editorial.git
cd Automatizaci-n-Editorial
```

### 2. Crear y activar un entorno virtual

Es obligatorio ejecutar la aplicación dentro de un entorno virtual para evitar conflictos con otras dependencias del sistema.

**Windows:**
```bash
python -m venv venv
venv\Scripts\activate
```

**macOS / Linux:**
```bash
python3 -m venv venv
source venv/bin/activate
```

> ⚠️ La interfaz nativa (`pywebview`) y el empaquetado con PyInstaller están pensados y probados para **Windows**.

### 3. Instalar dependencias

```bash
pip install -r requirements.txt
```

### 4. Ejecutar la aplicación

```bash
cd Programa
python main.py
```

Esto levanta el servidor FastAPI en `http://127.0.0.1:8765` y abre la ventana de la aplicación automáticamente.

> ⚠️ Recuerda activar el entorno virtual **cada vez** que abras una nueva terminal antes de ejecutar el programa.

---

## Compilar el ejecutable (.exe)

El ejecutable se genera con PyInstaller a partir de `Programa/app.spec`. **Debe correrse desde dentro de la carpeta `Programa/`**, ya que el `.spec` usa rutas relativas a esa carpeta (para incluir `static/`, `core/`, etc.).

```bash
cd Programa
pip install pyinstaller   # si no está instalado
pyinstaller --clean app.spec
```

El resultado queda en `Programa/dist/EditorSemantico/EditorSemantico.exe`.

> Las carpetas `build/` y `dist/` se generan automáticamente y **no se suben al repositorio** (ver `.gitignore`). Cualquiera puede regenerarlas siguiendo estos pasos.

---

## Estructura del repositorio

```
Automatizaci-n-Editorial/
│
├── README.md                          →  Este archivo
├── requirements.txt                   →  Dependencias del proyecto
├── .gitignore                         →  Archivos/carpetas excluidos del repo (venv, build, dist, __pycache__...)
│
├── Material de apoyo PDF y resultados de funciones/
│   ├── Errores/                       →  Casos de prueba con errores conocidos
│   ├── Material inicial, PDF/         →  PDFs de artículos y archivos de entrada para pruebas
│   ├── Resultados de EPUB/            →  EPUBs generados por el programa
│   ├── Resultados de HTML/            →  HTMLs generados por el programa
│   └── Resultados de XML/             →  XMLs JATS generados por el programa
│
├── Programa/                          →  Código fuente de la aplicación
│   ├── main.py                        →  Punto de entrada (FastAPI + pywebview)
│   ├── server.py                      →  API REST (FastAPI) — reemplaza la app de escritorio anterior
│   ├── app.spec                       →  Configuración de PyInstaller para generar el .exe
│   ├── GUIA_rapida_ExtractorPDF.txt   →  Guía de uso rápido para usuarios finales
│   ├── sesion_guardada.json           →  Estado de sesión persistido
│   ├── pyrightconfig.json             →  Configuración del linter/tipado
│   │
│   ├── core/                          →  Lógica de negocio pura (sin UI)
│   │   ├── constans.py                →  Clasificaciones, colores, CSS editorial
│   │   ├── pdf_processor.py           →  Extracción y clasificación de bloques del PDF
│   │   ├── html_exporter.py           →  Construcción del HTML editorial
│   │   ├── epub_exporter.py           →  Empaquetado EPUB 2.0
│   │   ├── jats_exporterv2.py         →  Generación de XML JATS (SciELO SPS)
│   │   ├── xml_validator.py           →  Validación estructural del XML JATS
│   │   └── utils.py                   →  Funciones auxiliares compartidas (ORCID, regex, etc.)
│   │
│   ├── static/                        →  Frontend (interfaz SPA)
│   │   ├── index.html                 →  Estructura de la interfaz
│   │   ├── app.js                     →  Lógica de la interfaz / consumo de la API
│   │   ├── style.css                  →  Estilos visuales
│   │   └── assets/                    →  Recursos gráficos
│   │
│   └── ui/                            →  ⚠️ Código legacy (versión anterior de escritorio)
│       ├── app_window.py              →  App con customtkinter — reemplazada por server.py + static/
│       ├── app_v2.py                  →  Lanzador de la app legacy (LimpiadorEditorialApp)
│       ├── tabs/                      →  Pestañas de la interfaz legacy (PDF, Autores, etc.)
│       └── widgets/                   →  Widgets de la interfaz legacy
│
└── venv/                               →  Entorno virtual de Python (no se sube a GitHub)
```

> **Nota sobre `Programa/ui/`:** `app_window.py`, `app_v2.py`, `tabs/` y `widgets/` corresponden a la **versión anterior** de la aplicación (interfaz nativa con `customtkinter`). Se conservan como referencia, pero **la aplicación actual no los usa** — el punto de entrada vigente es `main.py`.

---

## Flujo de trabajo

```
1. Cargar PDF       →  El programa analiza y clasifica los bloques automáticamente
2. Autores / ORCID  →  Excel con columnas Autor | ORCID  (o agregar manualmente)
3. Afiliaciones     →  .txt con líneas numeradas  (ej: 1 Institución, Ciudad...)
4. Referencias      →  .txt con referencias numeradas
5. Figuras          →  Imágenes con pie de figura y texto ancla
6. Tablas           →  Excel con una hoja por tabla
7. Exportar HTML    →  Genera el artículo con vista previa
8. Exportar EPUB    →  No requiere generar HTML primero
9. Exportar XML     →  Genera JATS compatible con SciELO SPS, con validación integrada
```

---

## Características técnicas

| Funcionalidad | Detalle |
|---|---|
| Detección de columnas | Automática por página (1 o 2 columnas) |
| Continuación de párrafos | Regla universal por puntuación final |
| Guiones tipográficos | Eliminados automáticamente al unir líneas |
| Cornisas | Filtradas por posición (top/bottom 5%) |
| ORCID | Acepta link completo o solo los números |
| Tablas | Excel multi-hoja, una hoja = una tabla |
| Figuras | Inserción por texto ancla o al final del documento |
| Referencias | Solo desde .txt externo (nunca del PDF) |
| EPUB | Generado directamente desde el HTML clasificado, sin pasos intermedios |
| XML | Formato JATS orientado a SciELO SPS, con metadatos, front, body y validación de estructura |
| Vista previa | Endpoints `/preview` permiten previsualizar HTML y XML antes de exportar |

---

## Estilos editoriales aplicados

- **Fuente cuerpo:** Source Serif 4, 12pt, justificado
- **Abstract en inglés:** Times New Roman, 12pt, cursiva, gris
- **Encabezados de sección:** 13pt, centrado, con línea divisora
- **Subtítulos numerados:** 12pt bold / 12pt bold italic
- **Tablas:** Colores institucionales (#1b5e9a / #cbeefb)
- **ORCID:** Subrayado verde (#A6CE39), link directo al perfil del autor

---

## Lectores EPUB compatibles

El EPUB generado cumple la especificación EPUB 2.0 y ha sido probado en los siguientes lectores:

| Lector | Sistema | Resultado |
|---|---|---|
| Calibre | Windows / Mac / Linux | ✅ Compatible |
| Thorium Reader | Windows / Mac / Linux | ✅ Compatible |
| Apple Books | Mac / iPhone / iPad | ✅ Compatible |

> Se recomienda **Thorium Reader** para usuarios finales en Windows por su interfaz moderna y cumplimiento estricto del estándar EPUB. Descarga en [thorium-reader.org](https://thorium-reader.org).

---

## Colaboración y control de versiones

### Gestión de ramas

- **Rama `main`:** Exclusivamente para código estable y funcional. **No realizar modificaciones directas.**
- **Nuevas funciones:** Crear una rama independiente (ej. `feature/nueva-funcion`), desarrollar y probar, luego abrir un *Pull Request* para revisión antes de integrar a `main`.

### Archivos generados (no versionados)

`venv/`, `Programa/build/`, `Programa/dist/`, `__pycache__/` y archivos `.pyc` están excluidos vía `.gitignore`. Se regeneran localmente siguiendo los pasos de instalación y compilación de este README.

### Reporte de errores

- El proceso de QA consiste en generar los archivos HTML, EPUB y XML, y compararlos con el PDF original.
- Los errores se documentan **exclusivamente en la pestaña Issues** de este repositorio.
- Al crear un Issue, asigna la etiqueta `bug` e incluye: descripción del problema, página exacta del PDF donde ocurre y, si es posible, adjunta el PDF y el archivo generado.

---

## Estado del proyecto

> ✅ Versión funcional — PDF a HTML
> ✅ Versión funcional — PDF a EPUB
> ✅ Versión funcional — PDF a XML JATS (SciELO SPS), con validación integrada
> ✅ Migración de interfaz de escritorio (customtkinter) a interfaz web embebida (FastAPI + pywebview)

---

## Colaboradores

- Ileana Verónica Lee Obando — Ingeniería en Computación
- David Alejandro Galicia Cárdenas — Licenciatura en Informática
- Erick Isaac Echeverria Goicochea — Ingeniería en Computación

Servicio Social de Programación Editorial
Instituto de Geología, Universidad Nacional Autónoma de México (UNAM)

---

*Desarrollado y mantenido por el equipo del Servicio Social de Programación Editorial, Instituto de Geología, UNAM.*