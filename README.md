# Herramienta de Extracción y Limpieza Editorial (PDF a HTML)

**Proyecto de Automatización Editorial - Instituto de Geología, UNAM**

Este repositorio contiene el código fuente de una aplicación de escritorio desarrollada en Python, diseñada para optimizar la conversión de artículos científicos de formato PDF a HTML semántico. La herramienta extrae el contenido respetando la estructura original y permite la clasificación manual de los bloques de texto a través de una interfaz gráfica de usuario (GUI).

## Tecnologías y requisitos del sistema

El proyecto está desarrollado en **Python 3**. Para ejecutar la aplicación correctamente en un entorno local, es necesario instalar las siguientes dependencias:

* `PyMuPDF` (fitz): Para la lectura y extracción de texto y metadatos de los archivos PDF.
* `customtkinter`: Para el renderizado de la interfaz gráfica con diseño moderno.

**Comando de instalación:**
```bash
pip install pymupdf customtkinter

```

## Instrucciones de uso

1. **Clonar el repositorio:** Descarga el proyecto en tu equipo local utilizando GitHub Desktop o mediante el comando `git clone`.
2. **Ejecutar la aplicación:** Abre una terminal en la carpeta del proyecto y ejecuta el archivo principal:
```bash
python Codigo.py

```

3. **Cargar el documento:** En la interfaz de la aplicación, haz clic en **"Cargar PDF"** y selecciona el artículo a procesar.
4. **Clasificación de bloques:** La herramienta dividirá el texto en segmentos. Utiliza el menú desplegable junto a cada bloque para asignarle la etiqueta HTML correspondiente (`Normal`, `Título`, `Pie de Figura` o `Ignorar`).
5. **Exportación:** Una vez clasificada la información relevante, haz clic en **"Generar HTML Limpio"** para guardar el archivo final estructurado.

## Flujo de Trabajo y colaboración

Para mantener la integridad del código y asegurar un desarrollo ordenado, el equipo deberá seguir estos lineamientos:

### 1. Gestión de ramas (Branches)

* **Rama `main`:** Esta rama es estrictamente para código estable y funcional. **No se deben realizar modificaciones directas en esta rama.**
* **Desarrollo de nuevas características (Erick):** Para trabajar en la optimización de extracción de tablas o nuevas funciones, es obligatorio crear una rama independiente (ej. `feature/extraccion-tablas`). Una vez finalizado y probado el código, se deberá crear un *Pull Request* para su revisión antes de integrarlo a `main`.

### 2. Control de Calidad y Reporte de Errores (David)

* El proceso de QA consiste en generar los archivos HTML y compararlos línea por línea con el PDF original.
* **Reporte de anomalías:** Si se detecta pérdida de formato, texto duplicado o problemas de anclaje de imágenes, el error **debe documentarse exclusivamente en la pestaña de "Issues"** de este repositorio.
* Al crear un *Issue*, asigna la etiqueta `bug` e incluye una descripción detallada del problema, mencionando la página exacta del PDF donde ocurre la falla para facilitar su corrección.

---

*Desarrollado y mantenido por el equipo de Ingeniería en Computación y Análisis de Datos.*

