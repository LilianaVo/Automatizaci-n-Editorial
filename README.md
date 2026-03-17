# Extractor de PDF Semántico
### Herramienta de Automatización Editorial — Paleontología Mexicana, UNAM

Aplicación de escritorio desarrollada en Python para el equipo editorial de la revista **Paleontología Mexicana** (Instituto de Geología, UNAM). Convierte artículos científicos en PDF a HTML editorial con estilos tipográficos fieles al diseño de la revista, optimizando el flujo de trabajo de maquetación.

---

## ¿Qué hace?

- Extrae y clasifica automáticamente los bloques de texto de un PDF científico (títulos, resúmenes, cuerpo, referencias, tablas, figuras, etc.)
- Genera un HTML limpio con tipografía editorial (Source Serif 4, Times New Roman)
- Vincula autores con sus perfiles ORCID
- Inserta tablas Excel y figuras en la posición exacta del texto mediante texto ancla
- Maneja artículos en dos columnas, guiones de corte tipográfico y saltos de página

---

## Tecnologías y requisitos del sistema

El proyecto está desarrollado en **Python 3.10 o superior**. Para ejecutar la aplicación es necesario instalar las siguientes dependencias:

- `PyMuPDF` (fitz) — lectura y extracción de texto y metadatos de archivos PDF
- `customtkinter` — interfaz gráfica con diseño moderno
- `Pillow` — procesamiento de imágenes
- `openpyxl` — lectura de archivos Excel

---

## Instalación

### 1. Clonar el repositorio

```bash
git clone https://github.com/tu-usuario/tu-repositorio.git
cd tu-repositorio
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

### 3. Instalar dependencias

```bash
pip install customtkinter pymupdf Pillow openpyxl
```

### 4. Ejecutar la aplicación

```bash
python app.py
```

> ⚠️ Recuerda activar el entorno virtual **cada vez** que abras una nueva terminal antes de ejecutar el programa.

---

## Estructura del repositorio

```
app.py                       →  Código fuente principal de la aplicación
app.spec                     →  Configuración para compilar el ejecutable (.exe)
GUIA_rapida_ExtractorPDF.txt →  Guía de uso rápido para usuarios finales
README.md                    →  Este archivo
build/                       →  Archivos temporales generados al compilar (no modificar)
dist/                        →  Aquí aparece el .exe listo para distribuir
Material Apoyo/              →  PDFs, Excels y archivos de prueba usados en desarrollo
venv/                        →  Entorno virtual de Python (no se sube a GitHub)
```

> **Para generar el ejecutable (.exe) sin necesidad de Python:**
> ```bash
> pip install pyinstaller
> pyinstaller app.spec
> ```
> El `.exe` resultante aparece en `dist/` y puede ejecutarse en cualquier Windows.

---

## Flujo de trabajo

```
1. Cargar PDF       →  El programa analiza y clasifica los bloques automáticamente
2. Autores / ORCID  →  Excel con columnas Autor | ORCID  (o agregar manualmente)
3. Afiliaciones     →  .txt con líneas numeradas  (ej: 1 Institución, Ciudad...)
4. Referencias      →  .txt con referencias numeradas
5. Figuras          →  Imágenes con pie de figura y texto ancla
6. Tablas           →  Excel con una hoja por tabla
7. Exportar HTML    →  Botón verde "HTML"
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

---

## Estilos editoriales aplicados

- **Fuente cuerpo:** Source Serif 4, 12pt, justificado
- **Abstract en inglés:** Times New Roman, 12pt, cursiva, gris
- **Encabezados de sección:** 13pt, centrado, con línea divisora
- **Subtítulos numerados:** 12pt bold / 12pt bold italic
- **Tablas:** Colores institucionales (#1b5e9a / #cbeefb)
- **ORCID:** Subrayado verde (#A6CE39), link directo al perfil del autor

---

## Colaboración y control de versiones

### Gestión de ramas

- **Rama `main`:** Exclusivamente para código estable y funcional. **No realizar modificaciones directas.**
- **Nuevas funciones:** Crear una rama independiente (ej. `feature/nueva-funcion`), desarrollar y probar, luego abrir un *Pull Request* para revisión antes de integrar a `main`.

### Reporte de errores

- El proceso de QA consiste en generar los archivos HTML y compararlos con el PDF original.
- Los errores se documentan **exclusivamente en la pestaña Issues** de este repositorio.
- Al crear un Issue, asigna la etiqueta `bug` e incluye: descripción del problema, página exacta del PDF donde ocurre y, si es posible, adjunta el PDF y el HTML generado.

---

## Estado del proyecto

> ✅ Versión funcional — PDF a HTML  
> 🔧 En desarrollo — Exportación XML y EPUB

---

## Colaboradores

- Ileana Verónica Lee Obando - Ingeniería en Computación
- David Alejandro Galicia Cárdenas - Licenciatura en Informática
- Erick -  Ingeniería en Computación

Servicio Social de Programación Editorial  
Instituto de Geología, Universidad Nacional Autónoma de México (UNAM)

---

*Desarrollado y mantenido por el equipo del Servicio Social de Programación Editorial, Instituto de Geología, UNAM.*