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


def es_encabezado_resumen(t: str) -> bool:
    """Coincidencia estricta para encabezados de resumen.
    Evita falsos positivos como 'Abstracto'.
    """
    norm = unicodedata.normalize("NFKD", t)
    norm = "".join(ch for ch in norm if not unicodedata.combining(ch))
    norm = re.sub(r"\s+", " ", norm).strip().lower()
    return bool(re.match(
        r"^(resumen|abstract|resumen no tecnico|non-technical abstract)\s*[:\.]?$",
        norm
    ))


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