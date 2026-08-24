"""
core/utils.py
Funciones utilitarias puras — sin dependencias de UI ni de fitz/PDF.
Usadas por extractor.py, html_exporter.py, epub_exporter.py y jats_exporter.py.
"""

from __future__ import annotations
import re
import os
import base64
import unicodedata


# ─── Escape HTML ──────────────────────────────────────────────────────────────

def esc(t: str) -> str:
    """Escapa caracteres especiales HTML básicos."""
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def esc_con_etiquetas_editoriales(t: str) -> str:
    """Escapa HTML y formatea etiquetas editoriales puntuales.
    Solo aplica cuando llevan punto final: 'Sinonimia.', 'Material.',
    'Descripción.' y similares.
    """
    texto = esc(t)
    patron = re.compile(
        r"\b("
        r"Sinonimia\.|"
        r"Material(?: examinado)?\.| Referred material\.|"
        r"Descripción\.|Description\.|"
        r"Etología\.|"
        r"Especie tipo\.|Type species\.|"
        r"Medidas\.|Dimensions\.|"
        r"Distribución\.|Occurrence\.|"
        r"Otras localidades\.|Other occurrences\.|"
        r"Discusión\.|Remarks\.|"
        r"Comentarios\.|Comments\.|"
        r"Repositorio\.|Repository\.|"
        r"Localidad\.|Locality\.|"
        r"Nuevo registro\.|"
        r"Horizonte\.|Horizonte y localidad\.|Horizon and locality\."
        r")\s*",
        re.IGNORECASE
    )

    def _reemplazo(m: re.Match) -> str:
        etiqueta = m.group(1)
        prefijo = "" if m.start() == 0 else "<br>"
        if etiqueta in ("Sinonimia.", "Locality Abbreviations:"):
            return f"{prefijo}<strong>{etiqueta}</strong><br>"
        return f"{prefijo}<strong>{etiqueta}</strong> "

    return patron.sub(_reemplazo, texto).strip()


# ─── ORCID ────────────────────────────────────────────────────────────────────

_ORCID_SVG = (
    "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmci"
    "IHZpZXdCb3g9IjAgMCAyNCAyNCI+PGNpcmNsZSBjeD0iMTIiIGN5PSIxMiIgcj0iMTIiIGZpbGw9"
    "IiNBNkNFMzkiLz48cGF0aCBkPSJNNy41IDVoMXY3LjVoLTF6TTkuMyA3LjhDOS4zIDYgMTAuNCA1"
    "IDEyLjEgNWMxLjggMCAyLjkgMSAyLjkgMi44VjEyaDEuMlY3LjhjMC0yLjQtMS40LTMuOC0zLjct"
    "My44QzkuOCA0IDguMSA1IDguMSA3LjhWMTJoMS4yVjcuOHoiIGZpbGw9IiNmZmYiLz48L3N2Zz4="
)

_ORCID_BASE   = "https://orcid.org/"
_ORCID_SEARCH = "https://orcid.org/orcid-search/search?searchQuery="


def insertar_orcid(texto: str, autores_orcid: list[dict] | None = None) -> str:
    """Construye el HTML de autores con links ORCID.

    Si autores_orcid está disponible, lo usa directamente (nombres + IDs exactos).
    Si no, hace fallback a búsqueda por nombre en orcid.org usando el texto crudo del PDF.
    """
    if autores_orcid:
        partes = []
        for a in autores_orcid:
            nombre = a["nombre"].strip()
            orcid  = a["orcid"].strip()
            if not nombre:
                continue
            nombre_esc = esc(nombre)
            if orcid:
                url = _ORCID_BASE + orcid
                ico = (
                    f'<a class="orcid-icon" href="{url}" target="_blank" '
                    f'title="ORCID: {orcid}">'
                    f'<img src="{_ORCID_SVG}" alt="ORCID"></a>'
                )
                partes.append(
                    f'<a class="orcid-autor" href="{url}" target="_blank" '
                    f'title="ORCID: {orcid}">{nombre_esc}</a>{ico}'
                )
            else:
                q   = re.sub(r"\s+", "+", nombre)
                url = _ORCID_SEARCH + q
                partes.append(
                    f'<a class="orcid-autor" href="{url}" target="_blank" '
                    f'title="Buscar en ORCID">{nombre_esc}</a>'
                )
        return "; ".join(partes)

    # Fallback: parsear el texto crudo del PDF
    partes = [p.strip() for p in texto.split(";") if p.strip()]
    result = []
    for parte in partes:
        nombre_raw = re.sub(r"[\d,\*\u00b9\u00b2\u00b3\u2070-\u209f]+$", "", parte).strip()
        if not nombre_raw:
            continue
        nombre_esc = esc(nombre_raw)
        q   = re.sub(r"\s+", "+", nombre_raw)
        url = _ORCID_SEARCH + q
        ico = (
            f'<a class="orcid-icon" href="{url}" target="_blank" title="Buscar en ORCID">'
            f'<img src="{_ORCID_SVG}" alt="ORCID"></a>'
        )
        result.append(
            f'<a class="orcid-autor" href="{url}" target="_blank" '
            f'title="Buscar {nombre_esc} en ORCID">{nombre_esc}</a>{ico}'
        )
    return "; ".join(result)


# ─── Referencias ──────────────────────────────────────────────────────────────

def parsear_referencias(texto: str) -> list[str]:
    """Parsea referencias numeradas desde texto plano.
    Soporta formatos: '1. Texto', '1) Texto', '[1] Texto'.
    """
    patron = re.compile(r'^\s*(?:\[?\d+[\.\)\]]\s*)', re.MULTILINE)
    partes = patron.split(texto)
    refs   = [p.strip() for p in partes if p.strip()]
    if refs:
        return refs
    return [l.strip() for l in texto.splitlines() if l.strip()]


# ─── Detectores de tipo de bloque ─────────────────────────────────────────────

def es_como_citar(t: str) -> bool:
    """Detecta si un bloque es la sección 'Cómo citar'."""
    if re.match(r"(cómo citar|how to cite)", t.strip().lower()):
        return True
    # Continuación de cita: bloque corto que termina con patrón de revista
    if (len(t) < 400 and
            re.search(r",\s*\d+\s*\(\d+\)\s*,\s*\d+\s*[–\-]\s*\d+", t) and
            not re.search(r"https?://", t)):
        return True
    return False


def _normalizar(t: str) -> str:
    """Quita acentos, colapsa espacios y pasa a minúsculas. Uso interno para
    comparar encabezados de sección sin importar tildes/mayúsculas."""
    norm = unicodedata.normalize("NFKD", t)
    norm = "".join(ch for ch in norm if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", norm).strip().lower()


def es_encabezado_resumen(t: str) -> bool:
    """Coincidencia estricta para encabezados de resumen.
    Evita falsos positivos como 'Abstracto'.
    """
    norm = _normalizar(t)
    return bool(re.match(
        r"^(resumen|abstract|resumen no tecnico|non-technical abstract)\s*[:\.]?$",
        norm
    ))


# ─── Anclas semánticas del artículo (enfoque tipo SciELO Markup) ──────────────
# Estas funciones detectan los encabezados que marcan el INICIO de cada zona
# del artículo (Resumen, Palabras clave, Cuerpo, Referencias), para que el
# clasificador ya no dependa solo del tamaño de letra sino de estas "anclas"
# de contenido, igual que hace el autómata de SciELO Markup.

def es_encabezado_palabras_clave(t: str) -> bool:
    """Detecta si el bloque ES (nada más que) el encabezado de la sección
    'Palabras clave' / 'Keywords', sin las palabras clave mismas."""
    norm = _normalizar(t)
    return bool(re.match(r"^(palabras\s+clave|keywords)\s*[:\.]?$", norm))


def es_inicio_palabras_clave(t: str) -> bool:
    """Detecta si el bloque EMPIEZA con 'Palabras clave:' seguido de las
    palabras en la misma línea (el caso más común en los PDF)."""
    norm = _normalizar(t)
    return bool(re.match(r"^(palabras\s+clave|keywords)\s*[:\.]", norm))


def es_encabezado_referencias(t: str) -> bool:
    """Detecta el encabezado de la sección de referencias bibliográficas."""
    norm = _normalizar(t)
    return norm in (
        "referencias", "references", "referencias bibliograficas",
        "bibliografia", "bibliography", "literatura citada",
    )


def es_encabezado_cuerpo_inicio(t: str) -> bool:
    """Detecta encabezados típicos donde arranca el cuerpo del artículo
    (p. ej. 'Introducción'), con o sin numeración delante ('1. Introducción')."""
    norm = _normalizar(t)
    norm = re.sub(r"^\d+[\.\)]?\s*", "", norm)
    return norm in (
        "introduccion", "introduction",
        "paleontologia sistematica", "systematic palaeontology",
        "material y metodos", "materiales y metodos", "metodologia",
        "material and methods", "materials and methods", "methods",
    )


def es_fecha_mss(t: str) -> bool:
    """Detecta si un bloque contiene fechas de manuscrito."""
    return bool(re.search(
        r"(manuscrito\s+recibido|manuscrito\s+corregido|manuscrito\s+aceptado"
        r"|manuscript\s+received|manuscript\s+revised|manuscript\s+accepted)",
        t, re.IGNORECASE))


def es_doi(t: str) -> bool:
    """Solo marca como fecha/DOI si es un bloque pequeño de metadatos,
    NO si es una referencia bibliográfica.
    """
    if not re.search(r"https?://doi\.org/", t):
        return False
    if len(t) > 200:
        return False
    if re.search(r"\.\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]", t) and not es_fecha_mss(t):
        return False
    return True


# ─── Extractores de metadatos del artículo ────────────────────────────────────
# A diferencia de los detectores anteriores (es_fecha_mss, es_doi), que solo
# regresan True/False, estos SÍ extraen el dato real del texto del bloque.

_MESES_ES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}


def extraer_fechas_mss(t: str) -> dict[str, str]:
    """Extrae las fechas de recepción, corrección y aceptación del manuscrito.

    Soporta el patrón usado por Paleontología Mexicana:
        'Manuscrito recibido: Junio 5, 2025.'
        'Manuscrito corregido: Diciembre 10, 2025.'
        'Manuscrito aceptado: Diciembre 12, 2025.'
    y su variante en inglés (received/revised/accepted).

    Devuelve dict con claves 'recibido', 'corregido', 'aceptado'.
    Las claves ausentes en el texto no se incluyen.
    """
    etiquetas = {
        "recibido":  r"manuscrito\s+recibido|manuscript\s+received",
        "corregido": r"manuscrito\s+corregido|manuscript\s+revised",
        "aceptado":  r"manuscrito\s+aceptado|manuscript\s+accepted",
    }

    resultado: dict[str, str] = {}
    for clave, patron_etiqueta in etiquetas.items():
        m = re.search(
            rf"(?:{patron_etiqueta})\s*:?\s*"
            rf"([A-Za-zÁÉÍÓÚñÑ]+\s+\d{{1,2}}\s*,\s*\d{{4}})",
            t, re.IGNORECASE,
        )
        if m:
            resultado[clave] = m.group(1).strip().rstrip(".")
    return resultado


def fecha_mss_a_iso(fecha_texto: str) -> str | None:
    """Convierte 'Junio 5, 2025' → '2025-06-05' (formato ISO para XML-JATS).
    Devuelve None si no se reconoce el formato.
    """
    m = re.match(r"([A-Za-zÁÉÍÓÚñÑ]+)\s+(\d{1,2})\s*,\s*(\d{4})", fecha_texto.strip())
    if not m:
        return None
    mes_texto, dia, anio = m.groups()
    mes_num = _MESES_ES.get(mes_texto.strip().lower())
    if not mes_num:
        return None
    return f"{int(anio):04d}-{mes_num:02d}-{int(dia):02d}"


def extraer_doi(t: str) -> str | None:
    """Extrae el DOI limpio (sin el prefijo de URL) de un bloque de texto.

    'https://doi.org/10.22201/igl.05437652e.2026.15.1.410 Manuscrito...'
    → '10.22201/igl.05437652e.2026.15.1.410'
    """
    m = re.search(r"https?://doi\.org/(\S+)", t)
    if not m:
        return None
    doi = m.group(1).rstrip(".,;")
    return doi


def extraer_volumen_pagina(t: str) -> dict[str, str]:
    """Extrae volumen, número, año y páginas del encabezado de la revista.

    Soporta el patrón usado por Paleontología Mexicana:
        'Volumen 15, núm. 1, 2026, p. 85 – 108'
        'Volume 15, no. 1, 2026, p. 85-108'

    Devuelve dict con las claves presentes entre:
        'volumen', 'numero', 'anio', 'pagina_inicio', 'pagina_fin'
    """
    resultado: dict[str, str] = {}

    m = re.search(
        r"vol(?:umen|ume)?\.?\s*(\d+)\s*,\s*"
        r"(?:n[uú]m\.?|no\.?|number)\s*(\d+)\s*,\s*"
        r"(\d{4})"
        r"(?:\s*,\s*p\.?\s*(\d+)\s*[–\-]\s*(\d+))?",
        t, re.IGNORECASE,
    )
    if m:
        resultado["volumen"] = m.group(1)
        resultado["numero"]  = m.group(2)
        resultado["anio"]    = m.group(3)
        if m.group(4) and m.group(5):
            resultado["pagina_inicio"] = m.group(4)
            resultado["pagina_fin"]    = m.group(5)
        return resultado

    # Fallback: solo páginas, sin volumen/número (ej. "p. 85 – 108" suelto)
    m2 = re.search(r"\bp\.?\s*(\d+)\s*[–\-]\s*(\d+)\b", t)
    if m2:
        resultado["pagina_inicio"] = m2.group(1)
        resultado["pagina_fin"]    = m2.group(2)

    return resultado


def extraer_issn(t: str) -> str | None:
    """Extrae el ISSN del encabezado de la revista.
    Soporta 'ISSN:2007-5189' y 'ISSN: 2007-5189'.
    """
    m = re.search(r"ISSN\s*:?\s*(\d{4}-\d{3}[\dXx])", t)
    return m.group(1) if m else None


# ─── Limpieza de texto ────────────────────────────────────────────────────────

def limpiar_prefijo_pie_figura(texto: str) -> str:
    """Quita 'Figura 1.'/'Fig. 2:' del inicio del pie."""
    t = re.sub(r"\s+", " ", texto).strip()
    t = re.sub(
        r"^\s*(figura|fig\.?|figure)\s*\d+[a-z]?(?:\s*-\s*[a-z])?[\.\ ):]?\s*",
        "",
        t,
        flags=re.IGNORECASE,
    )
    return t.strip(" -:\t")


def limpiar_prefijo_titulo_tabla(texto: str) -> str:
    """Quita 'Tabla 1.'/'Table 2:' del inicio del titulo."""
    t = re.sub(r"\s+", " ", texto).strip()
    t = re.sub(
        r"^\s*(tabla|table)\s*\d+[a-z]?(?:\s*-\s*[a-z])?[\.\ ):]?\s*",
        "",
        t,
        flags=re.IGNORECASE,
    )
    return t.strip(" -:\t")


# ─── Afiliaciones ─────────────────────────────────────────────────────────────

def split_afiliaciones_linea(linea: str) -> list[tuple[str, str]]:
    """Extrae afiliaciones (numero/letra + texto) desde una linea.

    Soporta casos como:
      '1 Institucion ... 2 Otra ...'
      'a Department of Earth Sciences ... b Faculty of ...'
      'a Department of Earth Sciences, University of California'  (una sola)
    """
    t = re.sub(r"\s+", " ", (linea or "")).strip()
    if not t:
        return []

    # ── Caso 1: línea empieza con marcador único y es una sola afiliación ──
    m_simple = re.match(
        r"^([0-9]+|[A-Za-z]{1,3})\s*[\)\.\:\-\*]?\s+([A-ZÁÉÍÓÚÑÜ].*)$",
        t, re.DOTALL
    )
    if m_simple:
        label = m_simple.group(1).strip()
        body  = m_simple.group(2).strip(" -:\t")
        _pat_interno = re.compile(
            r"[\.;:\)]\s+([0-9]+|[A-Za-z]{1,3})\s+[A-ZÁÉÍÓÚÑÜ]"
        )
        if not _pat_interno.search(body):
            return [(label, body)] if body else []

    # ── Caso 2: varias afiliaciones en la misma línea ──────────────────────
    pat = re.compile(r"([0-9]+|[A-Za-z]{1,3})\s+(?=[A-Za-zÁÉÍÓÚÑÜáéíóúñü])")
    starts: list[tuple[int, int]] = []
    for m in pat.finditer(t):
        s, e = m.start(1), m.end(1)
        if s == 0:
            starts.append((s, e))
            continue
        j = s - 1
        while j >= 0 and t[j].isspace():
            j -= 1
        if j >= 0 and t[j] in ".;:)":
            starts.append((s, e))

    seen_pos = set()
    uniq: list[tuple[int, int]] = []
    for s, e in starts:
        if s in seen_pos:
            continue
        seen_pos.add(s)
        uniq.append((s, e))
    starts = sorted(uniq, key=lambda x: x[0])

    if not starts:
        return []

    out: list[tuple[str, str]] = []
    for i, (s, e) in enumerate(starts):
        lim   = starts[i + 1][0] if i + 1 < len(starts) else len(t)
        label = t[s:e].strip()
        body  = t[e:lim].strip(" -.:;\t")
        if body:
            out.append((label, body))
    return out


# ─── Imágenes ─────────────────────────────────────────────────────────────────

def img_to_base64(path: str) -> str:
    """Devuelve data-URI base64 de la imagen."""
    ext  = os.path.splitext(path)[1].lower().lstrip(".")
    mime = {"jpg": "jpeg", "jpeg": "jpeg", "png": "png",
            "gif": "gif",  "webp": "webp"}.get(ext, "jpeg")
    with open(path, "rb") as f:
        data = base64.b64encode(f.read()).decode()
    return f"data:image/{mime};base64,{data}"