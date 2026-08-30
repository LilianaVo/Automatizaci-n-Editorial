"""core/pdf_processor.py
Lógica pura de extracción y clasificación de bloques a partir de un PDF.
No depende de ningún framework de UI (customtkinter, tkinter, etc.).

Uso:
    from core.pdf_processor import procesar_pdf

    resultado = procesar_pdf("articulo.pdf")
    bloques   = resultado["bloques"]               # list[dict]
    figuras   = resultado["figuras"]               # list[dict]
    tablas    = resultado["tablas"]                # list[dict]
    body_size = resultado["body_size"]              # int  (tamaño de fuente dominante)
    resumen   = resultado["resumen"]                # str  (estadísticas en texto)
    metadatos = resultado["metadatos_detectados"]   # dict (volumen, número, DOI, fechas...)
"""

from __future__ import annotations

import os
import re
import tempfile
import shutil
from collections import Counter
from typing import Any

import fitz  # PyMuPDF

from core.constans import CLASE_COMPAT as _CLASE_COMPAT
from core.utils import (
    es_como_citar as _es_como_citar,
    es_fecha_mss as _es_fecha_mss,
    es_doi as _es_doi,
    limpiar_prefijo_pie_figura as _limpiar_prefijo_pie_figura,
    limpiar_prefijo_titulo_tabla as _limpiar_prefijo_titulo_tabla,
    extraer_fechas_mss as _extraer_fechas_mss,
    fecha_mss_a_iso as _fecha_mss_a_iso,
    extraer_doi as _extraer_doi,
    extraer_volumen_pagina as _extraer_volumen_pagina,
    extraer_issn as _extraer_issn,
    es_linea_masthead as _es_linea_masthead,
    es_linea_filiacion as _es_linea_filiacion,
)
from core.zonas import detectar_breakpoints, CursorZonas

# ─────────────────────────────────────────────────────────────────────────────
# Helpers de bajo nivel sobre bloques fitz
# ─────────────────────────────────────────────────────────────────────────────

def _info_fuente(block: dict) -> tuple[float, bool, bool, str]:
    """Extrae (tamaño_promedio, bold, italic, fuente_dominante) de un bloque."""
    sizes: list[float] = []
    bold = italic = False
    fuentes: list[str] = []

    for line in block.get("lines", []):
        for span in line.get("spans", []):
            sizes.append(span["size"])
            fnt = span.get("font", "").lower()
            fuentes.append(fnt)
            if span["flags"] & (1 << 4):
                bold = True
            if span["flags"] & (1 << 1):
                italic = True
            if re.search(r"bold|black|heavy|semibold|demi", fnt):
                bold = True
            if re.search(r"italic|oblique|kursiv", fnt):
                italic = True

    avg = sum(sizes) / len(sizes) if sizes else 10.0
    dom = Counter(fuentes).most_common(1)[0][0] if fuentes else ""
    return avg, bold, italic, dom


def _texto_bloque(block: dict) -> str:
    """Extrae texto del bloque reconectando palabras cortadas por guión."""
    lineas: list[str] = []
    for line in block.get("lines", []):
        lineas.append("".join(s["text"] for s in line.get("spans", [])))

    _DASH_CORTES = r"-\u00ad\u2010\u2011\u2012\u2013\u2014"
    resultado = ""
    for i, linea in enumerate(lineas):
        if i == 0:
            resultado = linea
        else:
            if re.search(rf"[{_DASH_CORTES}]\s*$", resultado) and \
               re.match(r"^\s*[a-záéíóúñüa-z]", linea):
                resultado = re.sub(rf"[{_DASH_CORTES}]\s*$", "", resultado) + linea.lstrip()
            else:
                resultado = resultado + " " + linea

    resultado = resultado.replace("\u00ad", "")
    resultado = re.sub(r"\s+", " ", resultado).strip()
    resultado = re.sub(
        r"(\w)[\u2010\u2011\u2012\u2013\u2014-]\s+([a-záéíóúñüa-z])",
        r"\1\2", resultado)
    return resultado


# ─────────────────────────────────────────────────────────────────────────────
# Clasificador automático de bloques
# ─────────────────────────────────────────────────────────────────────────────

_SECCIONES_EXACTAS_CLAS = {
    "resumen", "abstract", "resumen no técnico", "non-technical abstract",
    "palabras clave", "keywords", "introducción", "introduction",
    "conclusiones", "conclusions", "referencias", "references",
    "agradecimientos", "acknowledgements", "discusión", "discussion",
    "metodología", "methods", "resultados", "results",
    "contribuciones de los autores", "contribucción de autores",
    "contribución de autores", "authors' contribution",
    "conflicto de intereses", "conflict of interest",
    "conflicts of interest", "competing interests",
    "paleontología sistemática", "systematic palaeontology",
}

_pat_afil_letra = re.compile(
    r"^[a-zA-Z]{1,3}\s+[A-ZÁÉÍÓÚÑÜ][a-zA-ZÁÉÍÓÚÑÜáéíóúñü]"
)


class _EstadoPortada:
    """Estado local para clasificar la portada (título/autores/filiaciones):
    hace falta recordar si ya se vio el título, si el 'espacio' para un
    subtítulo (traducción del título) ya se cerró, y si ya se entró en la
    zona de filiaciones (para que un fragmento de filiación cortado a la
    mitad por el extractor de PDF —sin el marcador '1 '/'2 ' al principio—
    no se pierda como 'Cuerpo')."""
    __slots__ = ("titulo_visto", "titulo_cerrado", "en_filiacion")

    def __init__(self):
        self.titulo_visto = False
        self.titulo_cerrado = False
        self.en_filiacion = False


def _clasificar_portada(
    texto: str, size: float, bold: bool, italic: bool,
    body_size: int, estado: _EstadoPortada,
) -> str:
    """Clasifica un bloque de la PORTADA del artículo (todo lo que va antes
    de cualquier ancla de Resumen/Abstract/Palabras clave/Cuerpo): cornisa
    editorial, título, autores, filiaciones, email.

    A diferencia de clasificar_auto(), acá el contenido manda primero. Los
    umbrales de tamaño de letra fijos (>=13, ==9...) que antes decidían casi
    todo estaban calibrados para el layout de una sola revista y se rompen
    con cualquier otra (o incluso dentro del mismo documento: en la práctica
    los 4 bloques de una misma línea de autores pueden salir con 4 tamaños
    de letra distintos, según cómo el extractor de PDF fragmentó el texto).

    El tamaño de letra solo se usa para UNA cosa que sí es una convención
    editorial real: el título es visiblemente más grande que el cuerpo del
    texto — y eso se mide en relación a body_size, no contra un número fijo,
    así que se adapta a la tipografía de cada revista.
    """
    t = texto.strip()

    # 1) Cornisa editorial (nombre de la revista, ISSN, volumen/núm/año,
    #    rango de fechas) — siempre, sin importar tamaño.
    if _es_linea_masthead(t):
        return "Encabezado sección"

    # 2) Metadatos de contacto/publicación — por contenido, como en el resto
    #    del documento.
    if _es_como_citar(t):
        return "Cómo citar"
    if _es_fecha_mss(t) or _es_doi(t):
        return "Fecha manuscrito"
    if re.search(r"@[\w\-\.]+\.\w{2,}", t) and len(t) < 120:
        return "Email / Metadatos"

    # 3) Filiación institucional — por contenido, no por tamaño. Una vez
    #    dentro de la zona de filiaciones, un fragmento sin marcador propio
    #    (porque el extractor de PDF lo cortó a la mitad, p. ej. "mayo s/n
    #    esquina Fuerte de Loreto..." sin el "1 " que abrió esa filiación)
    #    sigue considerándose parte de la MISMA filiación, no "Cuerpo" —
    #    hasta que aparezca otra cosa reconocible (email, otra filiación, un
    #    encabezado, etc.) que indique que ya se salió de esa zona.
    if _es_linea_filiacion(t):
        estado.en_filiacion = True
        return "Filiación"
    if estado.en_filiacion:
        return "Filiación"

    # 4) Título: el primer bloque "grande" (relativo al cuerpo) que no
    #    calzó ninguna regla anterior. Si el siguiente bloque es igual de
    #    grande, se interpreta como subtítulo (p. ej. la traducción del
    #    título) — es la única parte de la portada donde el tamaño decide
    #    algo, porque acá sí es una convención editorial genuina.
    es_titulo_candidato = size >= body_size + 2 or (size >= body_size + 1 and bold)

    if not estado.titulo_visto:
        if es_titulo_candidato:
            estado.titulo_visto = True
            return "Título principal"
        # Todavía no apareció el título real (esto es ruido de portada sin
        # patrón propio, p. ej. el nombre de la revista solo): NO tocar
        # titulo_cerrado todavía, para no cerrarle el paso al subtítulo
        # antes de que el título mismo haya aparecido.
        return "Cuerpo"

    if not estado.titulo_cerrado:
        estado.titulo_cerrado = True
        if es_titulo_candidato:
            return "Título secundario"

    # 5) Lo que sigue (autores, o cualquier residuo de portada sin patrón
    #    propio) no tiene una clase dedicada en el vocabulario de salida:
    #    los autores reales se cargan aparte, desde la pestaña ORCID.
    return "Cuerpo"


def clasificar_auto(texto: str, size: float, bold: bool, italic: bool,
                    font: str, body_size: int) -> str:
    """Clasifica un bloque de texto según sus atributos tipográficos y contenido."""
    t, t_low = texto.strip(), texto.strip().lower()
    f_low = (font or "").lower()
    es_bold = bold or bool(re.search(r"bold|black|heavy|semibold|demi", f_low))
    es_italic = italic or bool(re.search(r"italic|oblique|kursiv", f_low))

    if t_low in _SECCIONES_EXACTAS_CLAS:
        return "Encabezado sección"
    if _es_como_citar(t):
        return "Cómo citar"
    if _es_fecha_mss(t) or _es_doi(t):
        return "Fecha manuscrito"
    if re.match(r"^(palabras\s+clave|keywords)\s*[:\.]", t_low):
        return "Palabras clave"
    if re.match(r"^(tabla|table)\s+\d+[\.\:\s]", t_low):
        return "Título tabla"
    if re.search(r"@[\w\-\.]+\.\w{2,}", t) and len(t) < 120:
        return "Email / Metadatos"

    s = round(size)
    if re.match(r"^\d+\.\s+\S", t) and not re.match(r"^\d+\.\d", t) and s <= 12:
        return "Subencabezado"
    if re.match(r"^\d+\.\d+", t) and s <= 12:
        return "Subencabezado-bajo"
    if s >= 13 and es_bold and not es_italic:
        return "Título principal"
    if s >= 13 and es_bold and es_italic:
        return "Título secundario"
    if s == 13 and not es_bold and not es_italic:
        return "Autores"
    if s == 12 and es_italic:
        return "Email / Metadatos"
    if s == 12 and not es_bold:
        return "Normal"
    if s == 10 and es_bold:
        return "Encabezado sección"
    if s == 10 and not es_bold:
        return "Subencabezado"

    if _pat_afil_letra.match(t) and not re.search(r"\.\s+[A-Z]", t[:30]):
        if re.search(
            r",|\b(university|instituto|department|lab|center|centre|"
            r"museum|college|faculty|school|national|natural)\b",
            t, re.IGNORECASE
        ):
            return "Filiación"

    if s == 9:
        is_tnr = "times" in font or "roman" in font
        if not is_tnr:
            if re.match(r"^(\d+|[a-zA-Z])\s+[A-ZÁÉÍÓÚÑÜ]", t):
                if re.search(
                    r",|\b(university|instituto|department|lab|center|centre|"
                    r"museum|college|faculty|school|national|natural)\b",
                    t, re.IGNORECASE
                ):
                    return "Filiación"
        if is_tnr or len(t) > 100:
            return "Resumen / Abstract"
        if len(t) < 80:
            return "Filiación"
        return "Resumen / Abstract"

    if len(t) < 4:
        return "Ignorar"
    return "Normal"


# ─────────────────────────────────────────────────────────────────────────────
# Extracción de figuras desde el PDF
# ─────────────────────────────────────────────────────────────────────────────

def extraer_figuras(doc: fitz.Document, ruta_pdf: str,
                    cache_dir: str | None = None) -> tuple[list[dict], str | None]:
    """
    Extrae imágenes del PDF y propone pie por cercanía.
    Devuelve (lista_figuras, directorio_temporal).
    El llamador es responsable de limpiar el directorio temporal cuando ya no
    lo necesite (shutil.rmtree).
    """
    base = os.path.splitext(os.path.basename(ruta_pdf))[0]
    base = re.sub(r"[^A-Za-z0-9_-]+", "_", base).strip("_") or "pdf"
    base = base[:24]

    if cache_dir:
        shutil.rmtree(cache_dir, ignore_errors=True)

    try:
        fig_dir = tempfile.mkdtemp(prefix=f"pm_fig_{base}_")
    except Exception:
        fig_dir = os.path.join(os.getcwd(), f"_pm_fig_{base}")
        os.makedirs(fig_dir, exist_ok=True)

    pat_caption = re.compile(
        r"^\s*(figura|fig\.?|figure|table|tabla)\s*\d+[a-z]?"
        r"(?:\s*-\s*[a-z])?[\.\):]?\s*",
        re.IGNORECASE,
    )

    encontrados: list[dict] = []
    for pnum in range(len(doc)):
        page = doc.load_page(pnum)
        blocks = page.get_text("dict").get("blocks", [])

        text_blocks = []
        for block in blocks:
            if block.get("type") != 0:
                continue
            txt = _texto_bloque(block).strip()
            if len(txt) < 3:
                continue
            x0, y0, x1, y1 = block.get("bbox", (0, 0, 0, 0))
            text_blocks.append({"texto": txt, "rect": fitz.Rect(x0, y0, x1, y1)})

        vistos: set = set()
        for img_idx, img in enumerate(page.get_images(full=True), 1):
            xref = img[0]
            try:
                img_data = doc.extract_image(xref)
            except Exception:
                continue
            if not img_data or "image" not in img_data:
                continue

            ext = (img_data.get("ext") or "png").lower()
            ext = re.sub(r"[^a-z0-9]", "", ext) or "png"
            if ext == "jpx":
                ext = "jpg"

            rects = page.get_image_rects(xref) or []
            for occ_idx, rect in enumerate(rects, 1):
                if rect.width < 32 or rect.height < 32:
                    continue

                key = (xref, round(rect.x0, 1), round(rect.y0, 1),
                       round(rect.x1, 1), round(rect.y1, 1))
                if key in vistos:
                    continue
                vistos.add(key)

                nom = f"p{pnum+1:03d}_img{img_idx:02d}_{occ_idx:02d}.{ext}"
                ruta_img = os.path.join(fig_dir, nom)
                try:
                    with open(ruta_img, "wb") as f:
                        f.write(img_data["image"])
                except Exception:
                    continue

                pie_auto = ""
                candidatos = []
                for tb in text_blocks:
                    tb_rect = tb["rect"]
                    overlap = max(
                        0.0,
                        min(rect.x1, tb_rect.x1) - max(rect.x0, tb_rect.x0)
                    )
                    ancho_ref = max(1.0, min(rect.width, tb_rect.width))
                    if overlap / ancho_ref < 0.15:
                        continue
                    dy = tb_rect.y0 - rect.y1
                    if dy < -8 or dy > 180:
                        continue
                    txt = tb["texto"].strip()
                    if not pat_caption.match(txt):
                        continue
                    candidatos.append((abs(dy), txt))

                if candidatos:
                    candidatos.sort(key=lambda c: c[0])
                    pie_auto = _limpiar_prefijo_pie_figura(candidatos[0][1])

                encontrados.append({
                    "ruta": ruta_img,
                    "pie": pie_auto,
                    "ancla": "",
                    "pagina": pnum + 1,
                    "origen": "auto_pdf",
                    "_y": rect.y0,
                    "_x": rect.x0,
                })

    encontrados.sort(
        key=lambda f: (f.get("pagina", 0), f.get("_y", 0), f.get("_x", 0))
    )
    for fig in encontrados:
        fig.pop("_y", None)
        fig.pop("_x", None)

    return encontrados, fig_dir


# ─────────────────────────────────────────────────────────────────────────────
# Extracción de tablas desde el PDF
# ─────────────────────────────────────────────────────────────────────────────

def extraer_tablas(doc: fitz.Document,
                   ruta_pdf: str) -> tuple[list[dict], dict[int, list], str]:
    """
    Intenta extraer tablas usando fitz.find_tables().
    Devuelve (tablas_auto, rects_por_pagina, mensaje_diagnostico).
    """
    base = os.path.splitext(os.path.basename(ruta_pdf))[0]
    base = re.sub(r"[^A-Za-z0-9_-]+", "_", base).strip("_") or "pdf"
    base = base[:24]

    try:
        tab_dir = tempfile.mkdtemp(prefix=f"pm_tab_{base}_")
    except Exception:
        tab_dir = os.path.join(os.getcwd(), f"_pm_tab_{base}")
        os.makedirs(tab_dir, exist_ok=True)

    tablas_auto: list[dict] = []
    rects_por_pagina: dict[int, list] = {}
    paginas_con_tablas = 0
    paginas_con_find_tables = 0
    errores_find_tables: list[str] = []
    errores_extract: list[str] = []
    errores_guardado: list[str] = []

    try:
        import openpyxl  # noqa: F401
        _tiene_openpyxl = True
    except ImportError:
        _tiene_openpyxl = False

    for pnum in range(len(doc)):
        page = doc.load_page(pnum)

        if not hasattr(page, "find_tables"):
            break
        paginas_con_find_tables += 1

        try:
            tabs = page.find_tables()
        except Exception as exc:
            errores_find_tables.append(str(exc)[:80])
            continue

        if not tabs or not tabs.tables:
            continue

        paginas_con_tablas += 1
        rects_por_pagina[pnum] = []

        for t_idx, tabla in enumerate(tabs.tables, 1):
            rect = fitz.Rect(tabla.bbox)
            rects_por_pagina[pnum].append(rect)

            if not _tiene_openpyxl:
                continue

            try:
                df_data = tabla.extract()
            except Exception as exc:
                errores_extract.append(f"p{pnum+1}t{t_idx}: {str(exc)[:60]}")
                continue

            if not df_data:
                continue

            nom = f"tabla_p{pnum+1:03d}_{t_idx:02d}.xlsx"
            ruta_xlsx = os.path.join(tab_dir, nom)
            try:
                import openpyxl as xl
                wb = xl.Workbook()
                ws = wb.active
                ws.title = f"Tabla_{t_idx}"
                for fila in df_data:
                    ws.append([c if c is not None else "" for c in fila])
                wb.save(ruta_xlsx)
                wb.close()
            except Exception as exc:
                errores_guardado.append(f"p{pnum+1}t{t_idx}: {str(exc)[:60]}")
                continue

            tablas_auto.append({
                "ruta": ruta_xlsx,
                "hoja": f"Tabla_{t_idx}",
                "rotulo": "",
                "descripcion": "",
                "ancla": "",
                "pagina": pnum + 1,
                "origen": "auto_pdf",
            })

    # Diagnóstico
    if tablas_auto:
        diag = (
            f"se extrajeron {len(tablas_auto)} tabla(s) "
            f"en {paginas_con_tablas} pagina(s)"
        )
    elif paginas_con_find_tables == 0:
        diag = "tu version de PyMuPDF no expone find_tables()"
    elif errores_find_tables:
        diag = "find_tables() fallo: " + " | ".join(errores_find_tables)
    elif errores_extract:
        diag = ("se detectaron tablas pero no se pudieron extraer: "
                + " | ".join(errores_extract))
    elif errores_guardado:
        diag = ("se detectaron tablas pero no se pudieron guardar: "
                + " | ".join(errores_guardado))
    else:
        diag = "find_tables() no detecto tablas en el PDF"

    return tablas_auto, rects_por_pagina, diag


# ─────────────────────────────────────────────────────────────────────────────
# Metadatos del artículo (volumen, número, páginas, DOI, fechas)
# ─────────────────────────────────────────────────────────────────────────────

def extraer_metadatos_articulo(bloques: list[dict]) -> dict[str, Any]:
    """
    Recorre los bloques ya clasificados buscando datos de la cabecera del
    artículo (volumen, número, año, páginas, DOI, ISSN, fechas de manuscrito)
    y los extrae a un diccionario plano, listo para mostrar en un formulario
    editable.

    No modifica los bloques ni su clasificación — solo lee.

    Las claves que pueden venir vacías ('' o {}) si no se detectó el dato:
      volumen, numero, anio, pagina_inicio, pagina_fin, doi, issn,
      fecha_recibido, fecha_corregido, fecha_aceptado (texto tal como aparece)
      fecha_recibido_iso, fecha_corregido_iso, fecha_aceptado_iso (formato XML)
    """
    metadatos: dict[str, Any] = {
        "volumen": "", "numero": "", "anio": "",
        "pagina_inicio": "", "pagina_fin": "",
        "doi": "", "issn": "",
        "fecha_recibido": "", "fecha_corregido": "", "fecha_aceptado": "",
        "fecha_recibido_iso": "", "fecha_corregido_iso": "", "fecha_aceptado_iso": "",
    }

    for b in bloques:
        texto = b.get("contenido", "")
        if not texto:
            continue

        # El encabezado de revista (volumen/núm./año/páginas) suele venir
        # junto con el ISSN en los primeros bloques del documento.
        if not metadatos["volumen"]:
            vol_info = _extraer_volumen_pagina(texto)
            if vol_info:
                metadatos.update({k: v for k, v in vol_info.items() if v})

        if not metadatos["issn"]:
            issn = _extraer_issn(texto)
            if issn:
                metadatos["issn"] = issn

        # El DOI y las fechas de manuscrito normalmente viven en el bloque
        # ya clasificado como "Fecha manuscrito" por clasificar_auto().
        if b.get("clasificacion") == "Fecha manuscrito" or _es_fecha_mss(texto) or _es_doi(texto):
            if not metadatos["doi"]:
                doi = _extraer_doi(texto)
                if doi:
                    metadatos["doi"] = doi

            fechas = _extraer_fechas_mss(texto)
            for clave_es, valor in fechas.items():
                campo = f"fecha_{clave_es}"
                if not metadatos.get(campo):
                    metadatos[campo] = valor
                    iso = _fecha_mss_a_iso(valor)
                    if iso:
                        metadatos[f"{campo}_iso"] = iso

    return metadatos


# ─────────────────────────────────────────────────────────────────────────────
# Pipeline principal
# ─────────────────────────────────────────────────────────────────────────────

def procesar_pdf(ruta: str) -> dict[str, Any]:
    """
    Extrae, clasifica y limpia los bloques de texto de un PDF.

    Retorna un diccionario con las claves:
      bloques    – list[dict]  cada dict tiene: contenido, clasificacion,
                               size, bold, italic, pnum
      figuras    – list[dict]  ruta, pie, ancla, pagina, origen
      tablas     – list[dict]  ruta, hoja, titulo, ancla, pagina, origen
      body_size  – int
      resumen    – str
      diag_tablas – str
      fig_dir    – str | None   directorio temporal de imágenes (limpiar con shutil.rmtree)
      tab_dir    – str | None   directorio temporal de tablas   (limpiar con shutil.rmtree)
    """
    doc = fitz.open(ruta)

    # ── Tamaño de fuente dominante ────────────────────────────────────────────
    all_sizes: list[int] = []
    for pnum in range(len(doc)):
        for block in doc.load_page(pnum).get_text("dict")["blocks"]:
            if block["type"] == 0:
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        all_sizes.append(round(span["size"]))
    body_size: int = Counter(all_sizes).most_common(1)[0][0] if all_sizes else 12

    # ── Detectar encabezados/pies de página repetidos (cornisas) ─────────────
    # En vez de listar a mano cada posible patrón de cornisa (DOI, apellido del
    # autor + "et al.", título corrido...), se cuenta qué textos aparecen en la
    # franja superior/inferior de varias páginas distintas. Si un texto se
    # repite así, es casi con certeza un encabezado/pie recurrente — sin
    # importar su contenido. Esto generaliza a cualquier artículo, revista o
    # apellido de autor, sin mantenimiento manual.
    _conteo_franja: Counter[str] = Counter()
    for _pnum in range(len(doc)):
        _page = doc.load_page(_pnum)
        _page_h = _page.rect.height
        for _block in _page.get_text("dict")["blocks"]:
            if _block["type"] != 0:
                continue
            _by0, _by1 = _block["bbox"][1], _block["bbox"][3]
            _en_franja = _by1 < _page_h * 0.12 or _by0 > _page_h * 0.90
            if not _en_franja:
                continue
            _texto_norm = re.sub(r"\d+", "#", _texto_bloque(_block).strip().lower())
            _texto_norm = re.sub(r"\s+", " ", _texto_norm).strip()
            if _texto_norm and len(_texto_norm) >= 6:
                _conteo_franja[_texto_norm] += 1

    _textos_cornisa_repetidos = {t for t, n in _conteo_franja.items() if n >= 2}

    # ── Tablas automáticas ────────────────────────────────────────────────────
    tablas_auto, rects_tablas_por_pagina, diag_tablas = extraer_tablas(doc, ruta)

    def _en_bloque_tabla(pnum: int, bbox) -> bool:
        rects = rects_tablas_por_pagina.get(pnum, [])
        if not rects:
            return False
        brect = fitz.Rect(bbox)
        area_b = max(1.0, brect.width * brect.height)
        for trect in rects:
            inter = brect & trect
            if inter.is_empty:
                continue
            if (inter.width * inter.height) / area_b >= 0.40:
                return True
        return False

    _pat_cornisa_txt = re.compile(
        r"^(https?://doi\.org/|doi\.org/|\d{1,3}$"
        r"|paleontolog[íi]a mexicana\s+vol\.)",
        re.IGNORECASE,
    )

    def _es_cornisa(by0: float, by1: float, page_h: float, pnum: int) -> bool:
        if pnum == 0:
            return False
        return by1 < page_h * 0.05 or by0 > page_h * 0.95

    # ── PASO 1: extraer bloques crudos ────────────────────────────────────────
    raw: list[dict] = []
    for pnum in range(len(doc)):
        page = doc.load_page(pnum)
        page_h = page.rect.height
        page_w = page.rect.width
        all_blocks = page.get_text("dict")["blocks"]

        # Detectar dos columnas
        text_blocks_page = [b for b in all_blocks if b["type"] == 0]
        cx_list = [(b["bbox"][0] + b["bbox"][2]) / 2 for b in text_blocks_page]
        mid = page_w * 0.55
        left_cx  = [cx for cx in cx_list if cx < mid]
        right_cx = [cx for cx in cx_list if cx >= mid]
        two_cols = len(left_cx) >= 2 and len(right_cx) >= 2

        if two_cols:
            left_blocks  = sorted(
                [b for b in all_blocks if (b["bbox"][0] + b["bbox"][2]) / 2 < mid],
                key=lambda b: b["bbox"][1])
            right_blocks = sorted(
                [b for b in all_blocks if (b["bbox"][0] + b["bbox"][2]) / 2 >= mid],
                key=lambda b: b["bbox"][1])
            ordered_blocks = left_blocks + right_blocks
        else:
            ordered_blocks = sorted(all_blocks, key=lambda b: b["bbox"][1])

        for block in ordered_blocks:
            if block["type"] == 1:
                x0, y0, x1, y1 = block["bbox"]
                w_px, h_px = abs(x1 - x0), abs(y1 - y0)
                ignorar = (w_px < 32 and h_px < 32) or \
                          _es_cornisa(y0, y1, page_h, pnum)
                raw.append({
                    "texto": f"[IMAGEN {w_px:.0f}×{h_px:.0f}px]",
                    "clasificacion": "Ignorar" if ignorar else "Imagen",
                    "size": 0, "bold": False, "italic": False,
                    "imagen": True, "pnum": pnum,
                })
                continue
            if block["type"] != 0:
                continue
            bx0, by0, bx1, by1 = block["bbox"]
            if _es_cornisa(by0, by1, page_h, pnum):
                continue
            if _en_bloque_tabla(pnum, (bx0, by0, bx1, by1)):
                continue
            size, bold, italic, font = _info_fuente(block)
            texto = _texto_bloque(block)
            if not texto or len(texto) < 3:
                continue
            if pnum > 0 and (by1 < page_h * 0.12 or by0 > page_h * 0.90):
                if _pat_cornisa_txt.search(texto.strip()[:100]):
                    continue
                texto_norm = re.sub(r"\d+", "#", texto.strip().lower())
                texto_norm = re.sub(r"\s+", " ", texto_norm).strip()
                if texto_norm in _textos_cornisa_repetidos:
                    continue
            raw.append({
                "texto": texto, "size": size,
                "bold": bold, "italic": italic, "font": font,
                "imagen": False, "clasificacion": None,
                "pnum": pnum,
            })

    # ── PASO 2: detectar anclas (enfoque tipo SciELO Markov / autómata) ───────
    # Se recorren todos los bloques UNA vez y se registran TODAS las
    # ocurrencias de cada ancla reconocida (Resumen/Abstract, Palabras
    # clave/Keywords, encabezados de sección del cuerpo, Referencias) — no
    # solo la primera. Ver core/zonas.py para el porqué.
    _pat_nivel1 = re.compile(r"^\d+\.\s*\S")
    _pat_nivel2 = re.compile(r"^\d+\.\d+")

    textos_raw = [
        ("" if r["imagen"] or r["clasificacion"] == "Ignorar" else r["texto"].strip())
        for r in raw
    ]
    sizes_raw = [r.get("size", 0) for r in raw]
    breakpoints = detectar_breakpoints(textos_raw, sizes_raw)
    cursor = CursorZonas(breakpoints)
    estado_portada = _EstadoPortada()

    # Clúster de cornisa editorial: el nombre de la revista, cuando va solo
    # (sin ISSN ni volumen en el mismo bloque, p. ej. "Paleontología
    # Mexicana" en su propia línea), no tiene un patrón de contenido propio
    # para reconocerlo aislado. Pero si dentro de los primeros bloques del
    # documento aparece una línea que SÍ es reconocible como cornisa
    # (ISSN/volumen/rango de fechas), todo lo que va antes de ese punto es
    # casi con certeza parte del mismo bloque editorial.
    _masthead_hasta = -1
    for _k in range(min(6, len(textos_raw))):
        if textos_raw[_k] and _es_linea_masthead(textos_raw[_k]):
            _masthead_hasta = _k
            break

    # ── PASO 3: clasificar según la zona vigente en cada bloque ──────────────
    bloques_raw: list[dict] = []
    for i, r in enumerate(raw):
        if r["imagen"]:
            bloques_raw.append({
                "contenido": r["texto"],
                "clasificacion": r["clasificacion"],
                "size": 0, "bold": False, "italic": False,
                "imagen": True, "pnum": r.get("pnum", 0),
            })
            continue

        texto = r["texto"].strip()
        t_low = texto.lower()
        zona = cursor.avanzar(i)
        es_encabezado_actual = zona is not None and zona.idx == i

        if zona is None:
            # "Portada": antes de cualquier ancla reconocida (título,
            # autores, filiaciones, email) — por contenido, no por tamaño
            # de letra fijo (ver _clasificar_portada).
            if i <= _masthead_hasta:
                cls = "Encabezado sección"
            else:
                cls = _clasificar_portada(
                    texto, r["size"], r["bold"], r["italic"], body_size, estado_portada)

        elif zona.tipo == "resumen":
            if es_encabezado_actual:
                cls = "Encabezado sección"
            elif _es_como_citar(texto):
                cls = "Cómo citar"
            elif _es_fecha_mss(texto) or _es_doi(texto):
                cls = "Fecha manuscrito"
            else:
                cls = "Cuerpo del abstract"

        elif zona.tipo == "palabras_clave":
            if es_encabezado_actual and not zona.inline:
                cls = "Encabezado sección"
            elif not es_encabezado_actual and _es_como_citar(texto):
                # "Cómo citar" suele venir metido justo después de las
                # palabras clave, antes del DOI/fechas de manuscrito.
                cls = "Cómo citar"
            elif not es_encabezado_actual and (_es_fecha_mss(texto) or _es_doi(texto)):
                cls = "Fecha manuscrito"
            else:
                # Encabezado inline ("Palabras clave: ...") o contenido que
                # sigue a un encabezado en su propio bloque: en ambos casos
                # el texto en sí son las palabras clave.
                cls = "Palabras clave"

        elif zona.tipo == "cuerpo":
            # Un encabezado numerado ("1. Introducción", "4.1 Metodología")
            # es Subencabezado/Subencabezado-bajo aunque sea el que abrió la
            # zona — igual que en la versión anterior. "Encabezado sección"
            # queda solo para anclas SIN numerar (p. ej. "Referencias" no
            # aplica acá, pero si un artículo escribe "Conclusiones" sin
            # número, cae en esta rama).
            if _pat_nivel1.match(texto) and not _pat_nivel2.match(texto):
                cls = "Subencabezado"
            elif _pat_nivel2.match(texto):
                cls = "Subencabezado-bajo"
            elif es_encabezado_actual:
                cls = "Encabezado sección"
            elif _es_como_citar(texto):
                cls = "Cómo citar"
            elif _es_fecha_mss(texto) or _es_doi(texto):
                cls = "Fecha manuscrito"
            elif re.search(r"\b(?:tabla|table)\s+\d+[\.\:\s]", t_low):
                cls = "Título tabla"
            elif re.match(r"^(figura|figure)\s+\d+[\.\:\s]", t_low):
                cls = "Pie de figura"
            else:
                cls = "Cuerpo"

        else:  # zona.tipo == "referencias"
            if es_encabezado_actual:
                cls = "Encabezado sección"
            # Nota: aquí solo se reconoce el encabezado explícito "Cómo
            # citar"/"How to cite" (no la regla "parece una cita" de
            # _es_como_citar, que dentro de la lista de referencias
            # coincidiría con casi cualquier renglón normal).
            elif re.match(r"^(cómo citar|how to cite)", t_low):
                cls = "Cómo citar"
            elif _es_fecha_mss(texto) or _es_doi(texto):
                cls = "Fecha manuscrito"
            else:
                cls = "Referencia"

        bloques_raw.append({
            "contenido": texto,
            "clasificacion": _CLASE_COMPAT.get(cls, cls),
            "size": r["size"],
            "bold": r["bold"],
            "italic": r["italic"],
            "pnum": r.get("pnum", 0),
        })

    # ── PASO 3b: suprimir filas de tabla del PDF ──────────────────────────────
    _pat_fila_tabla = re.compile(r"^(1[89]\d{2}|20\d{2})\s+\S")
    _pat_lista_crono_taxon = re.compile(
        r"^(1[89]\d{2}|20\d{2})\s+[A-Za-zÁÉÍÓÚÑáéíóúñü\-]+"
        r"(?:\s+[A-Za-zÁÉÍÓÚÑáéíóúñü\-]+){0,4}\s*;",
        re.IGNORECASE,
    )

    def _parece_fila_tabla(item: dict) -> bool:
        t = item["contenido"].strip()
        if item["clasificacion"] != "Cuerpo":
            return False
        if _pat_lista_crono_taxon.match(t):
            return False
        if _pat_fila_tabla.match(t):
            return True
        return False

    bloques_sin_tabla: list[dict] = []
    en_modo_tabla = False
    for item in bloques_raw:
        cls_i = item["clasificacion"]
        if cls_i == "Título tabla":
            en_modo_tabla = True
            bloques_sin_tabla.append(item)
            continue
        if not en_modo_tabla and _parece_fila_tabla(item):
            en_modo_tabla = True
            continue
        if en_modo_tabla:
            es_fin = cls_i in (
                "Subencabezado", "Subencabezado-bajo",
                "Encabezado sección", "Cómo citar", "Fecha manuscrito",
            ) or (
                cls_i == "Cuerpo" and
                len(item["contenido"]) > 200 and
                item["contenido"].rstrip()[-1] in ".?!"
            )
            if es_fin:
                en_modo_tabla = False
                bloques_sin_tabla.append(item)
        else:
            bloques_sin_tabla.append(item)
    bloques_raw = bloques_sin_tabla

    # ── PASO 4: separar encabezados embebidos ─────────────────────────────────
    HEADERS_EMBEBIDOS = [
        "Non-technical Abstract", "Non-Technical Abstract",
        "Resumen no técnico", "Resumen no Técnico",
        "Acknowledgments", "Acknowledgements", "Agradecimientos",
        "Conflicts of interest", "Conflict of interest", "Competing interests",
        "Conflicto de intereses", "Authors' Contribution",
        "Contribuciones de los autores", "Contribucción de autores",
        "Contribucción de Autores", "Contribución de Autores",
    ]
    _pat_abstract_solo = re.compile(r"(?<!\w)(Abstract)(?!\w)", re.IGNORECASE)
    _pat_tabla_embebida = re.compile(
        r"(?<!\w)((?:Tabla|Table)\s+\d+[\.\:\s][^\n]{0,200})", re.IGNORECASE
    )

    bloques_clean: list[dict] = []
    for item in bloques_raw:
        txt = item["contenido"]
        partido = False

        for hdr in HEADERS_EMBEBIDOS:
            idx = txt.find(hdr)
            if idx > 0:
                antes = txt[:idx].strip()
                if antes:
                    b1 = dict(item)
                    b1["contenido"] = antes
                    b1["clasificacion"] = clasificar_auto(
                        antes, item["size"], item["bold"],
                        item["italic"], "", body_size)
                    bloques_clean.append(b1)
                bloques_clean.append({
                    "contenido": hdr,
                    "clasificacion": "Encabezado sección",
                    "size": item["size"], "bold": item["bold"],
                    "italic": item["italic"], "pnum": item.get("pnum", 0),
                })
                despues = txt[idx + len(hdr):].strip()
                if despues:
                    b3 = dict(item)
                    b3["contenido"] = despues
                    bloques_clean.append(b3)
                partido = True
                break

        if partido:
            continue

        if not partido:
            m_abs = _pat_abstract_solo.search(txt)
            if m_abs and m_abs.start() > 0:
                ctx_antes = txt[max(0, m_abs.start() - 20):m_abs.start()].lower()
                es_non_technical = "technical" in ctx_antes or "técnico" in ctx_antes
                if not es_non_technical:
                    antes = txt[:m_abs.start()].strip()
                    despues = txt[m_abs.end():].strip()
                    if antes:
                        b1 = dict(item)
                        b1["contenido"] = antes
                        b1["clasificacion"] = clasificar_auto(
                            antes, item["size"], item["bold"],
                            item["italic"], "", body_size)
                        bloques_clean.append(b1)
                    bloques_clean.append({
                        "contenido": "Abstract",
                        "clasificacion": "Encabezado sección",
                        "size": item["size"], "bold": item["bold"],
                        "italic": item["italic"], "pnum": item.get("pnum", 0),
                    })
                    if despues:
                        b3 = dict(item)
                        b3["contenido"] = despues
                        bloques_clean.append(b3)
                    partido = True

        if item["clasificacion"] == "Cuerpo" and not partido:
            m = _pat_tabla_embebida.search(txt)
            if m and m.start() > 0:
                antes = txt[:m.start()].strip()
                titulo_tab = m.group(1).strip()
                despues = txt[m.end():].strip()
                if antes:
                    b1 = dict(item)
                    b1["contenido"] = antes
                    bloques_clean.append(b1)
                bloques_clean.append({
                    "contenido": titulo_tab,
                    "clasificacion": "Título tabla",
                    "size": item["size"], "bold": item["bold"],
                    "italic": item["italic"], "pnum": item.get("pnum", 0),
                })
                if despues:
                    b3 = dict(item)
                    b3["contenido"] = despues
                    bloques_clean.append(b3)
                continue
            elif m and m.start() == 0:
                item["clasificacion"] = "Título tabla"

        if not partido:
            bloques_clean.append(item)

    # ── PASO 5: fusionar bloques fragmentados por la extracción de PDF ────────
    # PyMuPDF a veces corta un párrafo lógico (un párrafo del cuerpo, una
    # referencia, una filiación) en dos o más bloques crudos —por un salto de
    # columna o de página, o simplemente por cómo el PDF fue generado— sin que
    # eso tenga nada que ver con dónde termina realmente ese párrafo. Antes
    # solo se volvían a pegar los bloques de "Cuerpo"; "Referencia" y
    # "Filiación" quedaban afuera a propósito, así que una referencia o una
    # filiación cortada por ese tipo de salto se quedaba partida para
    # siempre en dos bloques editables sueltos. Ahora se fusionan igual,
    # conservando su clasificación real (antes _vaciar() forzaba "Cuerpo"
    # en todo lo que fusionaba).
    NO_FUSIONAR = {
        "Título principal", "Título secundario",
        "Email / Metadatos", "Cómo citar",
        "Fecha manuscrito", "Encabezado sección",
        "Subencabezado", "Palabras clave",
    }
    FUSIONABLES = {"Cuerpo", "Cuerpo del abstract", "Referencia", "Filiación"}

    def _es_continuacion(anterior: str, siguiente: str,
                         pnum_ant: int, pnum_sig: int) -> bool:
        if abs(pnum_sig - pnum_ant) > 5:
            return False
        ant = anterior.rstrip()
        if not ant:
            return False
        if ant[-1] in ".?!:":
            return False
        return True

    fusionados: list[dict] = []
    buf_parrafos: list[tuple[str, int]] = []
    buf_item: dict | None = None
    buf_clase: str | None = None

    def _vaciar() -> None:
        nonlocal buf_item, buf_clase
        if buf_parrafos and buf_item is not None:
            merged = dict(buf_item)
            merged["contenido"] = "\n\n".join(t for t, _ in buf_parrafos)
            merged["clasificacion"] = buf_clase or "Cuerpo"
            fusionados.append(merged)
        buf_parrafos.clear()
        buf_item = None
        buf_clase = None

    for item in bloques_clean:
        cls  = item["clasificacion"]
        pnum = item.get("pnum", 0)

        if cls == "Ignorar" and item.get("imagen"):
            # Ícono decorativo incrustado en medio de una línea de texto
            # (p. ej. los íconos ORCID junto a cada autor): se conserva en
            # la lista para que se vea/edite si hace falta, pero NO corta
            # la fusión del texto que sigue — si lo hiciera, "Autor A" e
            # "ícono" e "Autor B" quedarían como 3 fragmentos sueltos en vez
            # de una sola línea de autores.
            fusionados.append(item)
            continue

        if cls in ("Imagen", "Ignorar", "Pie de figura", "Título tabla"):
            _vaciar()
            fusionados.append(item)
            continue

        if cls not in FUSIONABLES and cls != "Subencabezado-bajo":
            _vaciar()
            fusionados.append(item)
            continue

        # "Subencabezado-bajo" solo tiene sentido incrustado dentro de un
        # bloque de Cuerpo (ver marcador "§SUB§" más abajo); si el buffer
        # vigente es de otra clase fusionable (Referencia/Filiación/Cuerpo
        # del abstract), primero se cierra ese buffer antes de arrancar uno
        # de Cuerpo nuevo.
        clase_efectiva = "Cuerpo" if cls == "Subencabezado-bajo" else cls

        if clase_efectiva != buf_clase:
            _vaciar()
            buf_item = item
            buf_clase = clase_efectiva

        if cls == "Subencabezado-bajo":
            buf_parrafos.append(("§SUB§" + item["contenido"], pnum))
            continue

        txt_nuevo = item["contenido"]
        if buf_parrafos:
            ultimo_txt, ultimo_pnum = buf_parrafos[-1]
            if ultimo_txt.startswith("§SUB§"):
                buf_parrafos.append((txt_nuevo, pnum))
            elif _es_continuacion(ultimo_txt, txt_nuevo, ultimo_pnum, pnum):
                t = ultimo_txt.rstrip()
                if t.endswith("-"):
                    buf_parrafos[-1] = (t[:-1] + txt_nuevo.lstrip(), pnum)
                else:
                    buf_parrafos[-1] = (t + " " + txt_nuevo.lstrip(), pnum)
            else:
                buf_parrafos.append((txt_nuevo, pnum))
        else:
            buf_parrafos.append((txt_nuevo, pnum))

    _vaciar()
    bloques_utiles = fusionados

    # ── Figuras automáticas ───────────────────────────────────────────────────
    figuras_auto, fig_dir = extraer_figuras(doc, ruta)

    # ── Metadatos del artículo (volumen, número, páginas, DOI, fechas) ────────
    metadatos_detectados = extraer_metadatos_articulo(bloques_utiles)

    # ── Resumen estadístico ───────────────────────────────────────────────────
    conteo = Counter(b["clasificacion"] for b in bloques_utiles)
    resumen_str = "  |  ".join(f"{k}: {v}" for k, v in conteo.most_common(6))
    resumen_str += f"  |  Figuras auto: {len(figuras_auto)}"
    resumen_str += f"  |  Tablas auto: {len(tablas_auto)}"
    if not tablas_auto and diag_tablas:
        resumen_str += f"  |  diag tablas: {diag_tablas}"

    doc.close()

    return {
        "bloques":              bloques_utiles,
        "figuras":              figuras_auto,
        "tablas":                tablas_auto,
        "body_size":             body_size,
        "resumen":               resumen_str,
        "diag_tablas":           diag_tablas,
        "fig_dir":               fig_dir,
        "metadatos_detectados": metadatos_detectados,
    }