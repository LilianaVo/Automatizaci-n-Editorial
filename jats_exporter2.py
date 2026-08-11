from __future__ import annotations

"""Generador JATS 1.3.

El flujo general es:
1) normalizar/limpiar bloques,
2) extraer metadatos (titulos, autores, ISSN/DOI, afiliaciones),
3) construir front/article-meta,
4) mapear bloques a abstracts/cuerpo/referencias,
5) serializar XML legible.
"""

import os
import re
import unicodedata
import xml.etree.ElementTree as ET
from xml.dom import minidom


# Namespaces usados por JATS para enlaces e interoperabilidad.
XLINK_NS = "http://www.w3.org/1999/xlink"
MML_NS = "http://www.w3.org/1998/Math/MathML"
XML_NS = "http://www.w3.org/XML/1998/namespace"

ET.register_namespace("xlink", XLINK_NS)
ET.register_namespace("mml", MML_NS)


# Patrones de texto para detectar metadata editorial en encabezados.
_META_HEADING_RE = re.compile(
    r"issn|volumen\s+\d|vol\.\s*\d|num\.\s*\d|p\.\s*\d+|"
    r"enero|febrero|marzo|abril|mayo|junio|julio|agosto|"
    r"septiembre|octubre|noviembre|diciembre|"
    r"january|february|march|april|may|june|july|august|"
    r"september|october|november|december|\(\d{4}\)",
    re.IGNORECASE,
)

_DOI_URL_RE = re.compile(r"https?://doi\.org/(\S+)", re.IGNORECASE)
_DOI_CORE_RE = re.compile(r"(10\.\d{4,9}/\S+)", re.IGNORECASE)
_EMAIL_RE = re.compile(r"[\w.\-+%]+@[\w.\-]+\.\w{2,}")
_ISSN_RE = re.compile(r"\b\d{4}-\d{3}[\dXx]\b")
_AFF_LABEL_RE = re.compile(r"^([0-9]+|[A-Za-z]{1,3})\s*[\)\.\:\-]?\s+(.*)$", re.DOTALL)


def _split_affiliations_line(linea: str) -> list[tuple[str, str]]:
    """Extrae multiples afiliaciones de una sola linea.

    Ejemplos:
      "1 Institucion ... 2 Otra ..."
      "a Department of Earth Sciences ... b Faculty of ..."
      "a Department of Earth Sciences, University of California"  (una sola)
    """
    t = _clean_text(linea)
    if not t:
        return []

    # ── Caso 1: línea con un solo marcador (número o letra) ──────────────────
    m_simple = re.match(
        r"^([0-9]+|[A-Za-z]{1,3})\s*[\)\.\:\-\*]?\s+([A-ZÁÉÍÓÚÑÜ].*)$",
        t, re.DOTALL
    )
    if m_simple:
        label = m_simple.group(1).strip()
        body  = _clean_text(m_simple.group(2).strip(" -:\t"))
        _pat_interno = re.compile(
            r"[\.;:\)]\s+([0-9]+|[A-Za-z]{1,3})\s+[A-ZÁÉÍÓÚÑÜ]"
        )
        if not _pat_interno.search(body):
            return [(label, body)] if body else []

    # ── Caso 2: varias afiliaciones en la misma línea ────────────────────────
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

    uniq = []
    seen = set()
    for s, e in starts:
        if s in seen:
            continue
        seen.add(s)
        uniq.append((s, e))
    starts = sorted(uniq, key=lambda x: x[0])

    if not starts:
        m = _AFF_LABEL_RE.match(t)
        if not m:
            return []
        label, body = m.group(1).strip(), _clean_text(m.group(2))
        return [(label, body)] if body else []

    out: list[tuple[str, str]] = []
    for i, (s, e) in enumerate(starts):
        lim = starts[i + 1][0] if i + 1 < len(starts) else len(t)
        label = t[s:e].strip()
        body = _clean_text(t[e:lim].strip(" -.:;\t"))
        if body:
            out.append((label, body))
    return out


# --- Helpers de texto --------------------------------------------------------
def _norm(texto: str) -> str:
    """Normaliza texto para comparaciones robustas (sin tildes, minúsculas)."""
    t = unicodedata.normalize("NFKD", texto or "")
    t = "".join(ch for ch in t if not unicodedata.combining(ch))
    t = re.sub(r"\s+", " ", t).strip().lower()
    return t


def _clean_text(texto: str) -> str:
    """Limpia espacios y caracteres invisibles frecuentes en PDFs."""
    t = re.sub(r"[\u00ad\ufffc\ufffe]", "", texto or "")
    t = re.sub(r"\s+", " ", t).strip()
    return t


def _split_paragraphs(texto: str) -> list[str]:
    """Divide texto en parrafos y elimina marcadores internos de subtitulos."""
    if not texto:
        return []
    parts = [p.strip() for p in re.split(r"\n{2,}", texto) if p.strip()]
    if not parts:
        parts = [_clean_text(texto)]
    out = []
    for p in parts:
        p = p.replace("§SUB§", "").strip()
        if p:
            out.append(p)
    return out


# --- Extractores de metadatos ------------------------------------------------
def _extract_doi(bloques: list[dict]) -> str:
    """Busca el primer DOI en formato URL o en formato canónico 10.xxxx/..."""
    for b in bloques:
        txt = b.get("contenido", "")
        if not txt:
            continue
        m_url = _DOI_URL_RE.search(txt)
        if m_url:
            return m_url.group(1).rstrip(".),;")
        m_core = _DOI_CORE_RE.search(txt)
        if m_core:
            return m_core.group(1).rstrip(".),;")
    return ""


def _extract_issn(bloques: list[dict]) -> str:
    """Extrae el primer ISSN encontrado en los bloques."""
    for b in bloques:
        txt = b.get("contenido", "")
        if not txt:
            continue
        m = _ISSN_RE.search(txt)
        if m:
            return m.group(0).upper()
    return ""


def _clean_orcid(orcid_raw: str) -> str:
    """Normaliza ORCID y devuelve solo el identificador 0000-0000-0000-0000."""
    m = re.search(r"(\d{4}-\d{4}-\d{4}-\d{3}[\dX])", orcid_raw or "", re.IGNORECASE)
    return m.group(1).upper() if m else ""


# --- Parseo de autores/afiliaciones -----------------------------------------
def _authors_from_manual(autores_orcid: list[dict]) -> list[dict]:
    """Construye autores desde la pestaña manual (nombre + ORCID opcional)."""
    out = []
    for a in autores_orcid or []:
        nombre = _clean_text(a.get("nombre", ""))
        orcid = _clean_orcid(a.get("orcid", ""))
        if not nombre:
            continue
        out.append({"nombre": nombre, "orcid": orcid})
    return out


def _authors_from_pdf(bloques: list[dict]) -> list[dict]:
    """Fallback: toma autores detectados en el PDF cuando no hay carga manual."""
    out = []
    for b in bloques:
        if b.get("clasificacion") != "Autores":
            continue
        txt = b.get("contenido", "")
        for raw in [p.strip() for p in txt.split(";") if p.strip()]:
            nom = re.sub(r"[\d,\*\u00b9\u00b2\u00b3\u2070-\u209f]+$", "", raw).strip()
            if nom:
                out.append({"nombre": nom, "orcid": ""})
    return out


def _parse_affiliations_txt(afiliaciones_txt: str) -> tuple[list[dict], list[str]]:
    """Parsea afiliaciones desde .txt y correos de correspondencia.

    Los IDs XML (aff1, aff2, ...) siempre se generan únicos.
    La etiqueta visible puede conservar el prefijo original del txt (numero o letra).
    """
    afils: list[dict] = []
    emails: list[str] = []
    id_seq = 1
    for raw in (afiliaciones_txt or "").splitlines():
        linea = _clean_text(raw)
        if not linea:
            continue

        mail = _EMAIL_RE.search(linea)
        if linea.startswith("*") and mail:
            emails.append(mail.group(0))
            continue

        segs = _split_affiliations_line(linea)
        if segs:
            for label, resto in segs:
                afils.append({"id": f"aff{id_seq}", "label": label, "text": resto})
                id_seq += 1
        else:
            afils.append({"id": f"aff{id_seq}", "label": str(id_seq), "text": linea})
            id_seq += 1
    return afils, emails


def _fallback_affiliations_from_blocks(bloques: list[dict]) -> tuple[list[dict], list[str]]:
    """Fallback cuando no hay txt: usa bloques clasificados de afiliacion/email."""
    afils: list[dict] = []
    emails: list[str] = []
    seq = 1
    for b in bloques:
        cls = b.get("clasificacion")
        txt = _clean_text(b.get("contenido", ""))
        if not txt:
            continue
        if cls == "Filiación":
            afils.append({"id": f"aff{seq}", "label": str(seq), "text": txt})
            seq += 1
        elif cls == "Email / Metadatos":
            for em in _EMAIL_RE.findall(txt):
                if em not in emails:
                    emails.append(em)
    return afils, emails


# --- Referencias y keywords --------------------------------------------------
def _strip_ref_prefix(ref: str) -> str:
    """Quita prefijos numerados tipo '1.' o '[2]' de una referencia."""
    return re.sub(r"^\s*(?:\[?\d+[\.\)\]]\s*)", "", (ref or "").strip())


def _parse_keywords(texto: str) -> list[str]:
    """Separa keywords por coma/punto y coma y limpia etiqueta inicial."""
    t = _clean_text(texto)
    t = re.sub(r"^(palabras\s+clave|keywords)\s*[:\.]?\s*", "", t, flags=re.IGNORECASE)
    if not t:
        return []
    parts = [p.strip(" .;") for p in re.split(r"[;,]", t) if p.strip(" .;")]
    return parts


# --- Tablas (Excel -> XML JATS) ---------------------------------------------
def _table_rows_from_excel(ruta: str, hoja: str | None) -> tuple[list[list[str]], str]:
    """Lee una hoja de Excel y la transforma a matriz de strings."""
    try:
        import openpyxl
    except Exception:
        return [], "openpyxl no disponible"

    try:
        wb = openpyxl.load_workbook(ruta, data_only=True)
        ws = wb[hoja] if hoja and hoja in wb.sheetnames else wb.active
        rows: list[list[str]] = []
        for row in ws.iter_rows(values_only=True):
            vals = []
            has_content = False
            for cell in row:
                txt = "" if cell is None else _clean_text(str(cell))
                if txt:
                    has_content = True
                vals.append(txt)
            if has_content:
                rows.append(vals)
        wb.close()
        return rows, ""
    except Exception as exc:
        return [], str(exc)


def _append_table_xml(table_wrap: ET.Element, ruta: str, hoja: str | None) -> None:
    """Inserta contenido tabular JATS (<table>) dentro de <table-wrap>."""
    rows, err = _table_rows_from_excel(ruta, hoja)
    if not rows:
        p = ET.SubElement(table_wrap, "p")
        p.text = f"[No se pudo leer tabla: {err or 'tabla vacia'}]"
        return

    max_cols = max(len(r) for r in rows)
    norm_rows = [r + [""] * (max_cols - len(r)) for r in rows]

    tbl = ET.SubElement(table_wrap, "table")
    thead = ET.SubElement(tbl, "thead")
    trh = ET.SubElement(thead, "tr")
    for cell in norm_rows[0]:
        th = ET.SubElement(trh, "th")
        th.text = cell

    tbody = ET.SubElement(tbl, "tbody")
    for row in norm_rows[1:]:
        tr = ET.SubElement(tbody, "tr")
        for cell in row:
            td = ET.SubElement(tr, "td")
            td.text = cell


# --- Salida XML --------------------------------------------------------------
def _pretty_xml(root: ET.Element) -> str:
    """Serializa XML con sangría legible; fallback a serialización simple."""
    raw = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    try:
        parsed = minidom.parseString(raw)
        pretty = parsed.toprettyxml(indent="  ", encoding="utf-8").decode("utf-8")
        lines = [ln for ln in pretty.splitlines() if ln.strip()]
        return "\n".join(lines) + "\n"
    except Exception:
        return raw.decode("utf-8", errors="replace")


def build_jats_xml(
    *,
    bloques: list[dict],
    referencias_externas: list[str],
    autores_orcid: list[dict],
    afiliaciones_txt: str,
    figuras: list[dict],
    tablas: list[dict],
) -> str:
    """Genera un documento JATS 1.3 a partir de bloques semánticos de la app."""
    # Copia defensiva de bloques limpios
    bks = []
    for b in bloques or []:
        bks.append(
            {
                "contenido": b.get("contenido", ""),
                "clasificacion": b.get("clasificacion", "Cuerpo"),
                "italic": bool(b.get("italic", False)),
            }
        )

    title_es = next(
        (_clean_text(b["contenido"]) for b in bks if b["clasificacion"] == "Título principal" and _clean_text(b["contenido"])),
        "Articulo",
    )
    title_en = next(
        (_clean_text(b["contenido"]) for b in bks if b["clasificacion"] == "Título secundario" and _clean_text(b["contenido"])),
        "",
    )

    authors = _authors_from_manual(autores_orcid)
    if not authors:
        authors = _authors_from_pdf(bks)

    affs, emails = _parse_affiliations_txt(afiliaciones_txt)
    if not affs and not emails:
        affs, emails = _fallback_affiliations_from_blocks(bks)

    # Estructura del artículo
    root = ET.Element(
        "article",
        {
            "article-type": "research-article",
            "dtd-version": "1.3",
            f"{{{XML_NS}}}lang": "es",
            "specific-use": "generated-by-editor-semantico",
        },
    )

    front = ET.SubElement(root, "front")
    journal_meta = ET.SubElement(front, "journal-meta")
    journal_id = ET.SubElement(journal_meta, "journal-id", {"journal-id-type": "publisher-id"})
    journal_id.text = "Paleontologia Mexicana"
    journal_title_group = ET.SubElement(journal_meta, "journal-title-group")
    journal_title = ET.SubElement(journal_title_group, "journal-title")
    journal_title.text = "Paleontologia Mexicana"

    article_meta = ET.SubElement(front, "article-meta")
    issn = _extract_issn(bks)
    issn_node = ET.SubElement(journal_meta, "issn", {"pub-type": "epub"})
    issn_node.text = issn or "0000-0000"

    doi = _extract_doi(bks)
    if doi:
        aid = ET.SubElement(article_meta, "article-id", {"pub-id-type": "doi"})
        aid.text = doi

    title_group = ET.SubElement(article_meta, "title-group")
    article_title = ET.SubElement(title_group, "article-title")
    article_title.text = title_es
    if title_en:
        trans_title_group = ET.SubElement(title_group, "trans-title-group", {f"{{{XML_NS}}}lang": "en"})
        trans_title = ET.SubElement(trans_title_group, "trans-title")
        trans_title.text = title_en

    if authors:
        contrib_group = ET.SubElement(article_meta, "contrib-group")
        for a in authors:
            contrib = ET.SubElement(contrib_group, "contrib", {"contrib-type": "author"})
            if a.get("orcid"):
                cid = ET.SubElement(
                    contrib,
                    "contrib-id",
                    {"contrib-id-type": "orcid", "authenticated": "false"},
                )
                cid.text = f"https://orcid.org/{a['orcid']}"
            sname = ET.SubElement(contrib, "string-name")
            sname.text = a["nombre"]

    for af in affs:
        aff = ET.SubElement(article_meta, "aff", {"id": af["id"]})
        label = ET.SubElement(aff, "label")
        label.text = af["label"]
        aff.text = (aff.text or "") + " "
        inst = ET.SubElement(aff, "institution")
        inst.text = af["text"]

    if emails:
        author_notes = ET.SubElement(article_meta, "author-notes")
        corresp = ET.SubElement(author_notes, "corresp", {"id": "cor1"})
        corresp.text = "; ".join(emails)

    # JATS 1.3 Publishing exige pub-date o pub-date-not-available en article-meta.
    ET.SubElement(article_meta, "pub-date-not-available")

    # Parseo de bloques para abstracts, keywords, body y notas
    abstracts: dict[str, list[str]] = {"es": [], "en": [], "plain_es": [], "plain_en": []}
    kwds: dict[str, list[str]] = {"es": [], "en": [], "general": []}
    refs_from_blocks: list[str] = []
    cite_notes: list[str] = []
    date_notes: list[str] = []
    front_notes: list[str] = []
    body_stream: list[dict] = []

    abs_mode: str | None = None
    in_refs = False

    for b in bks:
        cls = b["clasificacion"]
        txt = _clean_text(b["contenido"])
        if not txt:
            continue

        if cls == "Encabezado sección":
            low = _norm(txt)
            if low in ("referencias", "references"):
                in_refs = True
                abs_mode = None
                continue

            # Front-matter editorial (ISSN, volumen, etc.)
            if _META_HEADING_RE.search(low) or low == "paleontologia mexicana":
                front_notes.append(txt)
                continue

            if low == "resumen":
                abs_mode = "es"
                in_refs = False
                continue
            if low == "abstract":
                abs_mode = "en"
                in_refs = False
                continue
            if low == "resumen no tecnico":
                abs_mode = "plain_es"
                in_refs = False
                continue
            if low == "non-technical abstract":
                abs_mode = "plain_en"
                in_refs = False
                continue

            in_refs = False
            abs_mode = None
            body_stream.append({"kind": "h", "text": txt})
            continue

        if in_refs:
            if cls in ("Referencia", "Cuerpo", "Normal"):
                refs_from_blocks.append(txt)
            continue

        if cls == "Palabras clave":
            vals = _parse_keywords(txt)
            if abs_mode in ("es", "plain_es"):
                kwds["es"].extend(vals)
            elif abs_mode in ("en", "plain_en"):
                kwds["en"].extend(vals)
            else:
                kwds["general"].extend(vals)
            continue

        if cls == "Cómo citar":
            cite_notes.append(txt)
            continue

        if cls == "Fecha manuscrito":
            date_notes.append(txt)
            continue

        if cls in ("Título principal", "Título secundario", "Autores", "Filiación", "Email / Metadatos"):
            continue
        if cls in ("Ignorar", "Imagen", "Pie de figura", "Título tabla"):
            continue

        if abs_mode and cls in ("Cuerpo", "Normal", "Resumen / Abstract"):
            abstracts[abs_mode].extend(_split_paragraphs(txt))
            continue

        if cls in ("Subencabezado", "Subencabezado-bajo"):
            body_stream.append({"kind": "h", "text": txt})
            abs_mode = None
            continue

        if cls in ("Cuerpo", "Normal", "Resumen / Abstract"):
            body_stream.append({"kind": "p", "text": txt})
            continue

    # Abstracts
    if abstracts["es"]:
        abs_es = ET.SubElement(article_meta, "abstract", {f"{{{XML_NS}}}lang": "es"})
        for ptxt in abstracts["es"]:
            p = ET.SubElement(abs_es, "p")
            p.text = ptxt

    if abstracts["plain_es"]:
        abs_plain_es = ET.SubElement(
            article_meta,
            "abstract",
            {"abstract-type": "plain-language-summary", f"{{{XML_NS}}}lang": "es"},
        )
        for ptxt in abstracts["plain_es"]:
            p = ET.SubElement(abs_plain_es, "p")
            p.text = ptxt

    if abstracts["en"]:
        abs_en = ET.SubElement(article_meta, "trans-abstract", {f"{{{XML_NS}}}lang": "en"})
        for ptxt in abstracts["en"]:
            p = ET.SubElement(abs_en, "p")
            p.text = ptxt

    if abstracts["plain_en"]:
        abs_plain_en = ET.SubElement(
            article_meta,
            "trans-abstract",
            {"abstract-type": "plain-language-summary", f"{{{XML_NS}}}lang": "en"},
        )
        for ptxt in abstracts["plain_en"]:
            p = ET.SubElement(abs_plain_en, "p")
            p.text = ptxt

    # Keywords
    def _append_kwd_group(lang: str, items: list[str]) -> None:
        uniques = []
        seen = set()
        for k in items:
            kk = _clean_text(k).strip(" .;")
            if kk and kk.lower() not in seen:
                seen.add(kk.lower())
                uniques.append(kk)
        if not uniques:
            return
        kg = ET.SubElement(article_meta, "kwd-group", {"kwd-group-type": "author-generated", f"{{{XML_NS}}}lang": lang})
        for k in uniques:
            kw = ET.SubElement(kg, "kwd")
            kw.text = k

    _append_kwd_group("es", kwds["es"])
    _append_kwd_group("en", kwds["en"])
    if kwds["general"]:
        _append_kwd_group("es", kwds["general"])

    # Notas editoriales capturadas del front
    if front_notes or cite_notes or date_notes:
        cmg = ET.SubElement(article_meta, "custom-meta-group")
        for i, txt in enumerate(front_notes, 1):
            cm = ET.SubElement(cmg, "custom-meta")
            mn = ET.SubElement(cm, "meta-name")
            mn.text = f"front-note-{i}"
            mv = ET.SubElement(cm, "meta-value")
            mv.text = txt
        for i, txt in enumerate(cite_notes, 1):
            cm = ET.SubElement(cmg, "custom-meta")
            mn = ET.SubElement(cm, "meta-name")
            mn.text = f"como-citar-{i}"
            mv = ET.SubElement(cm, "meta-value")
            mv.text = txt
        for i, txt in enumerate(date_notes, 1):
            cm = ET.SubElement(cmg, "custom-meta")
            mn = ET.SubElement(cm, "meta-name")
            mn.text = f"fecha-manuscrito-{i}"
            mv = ET.SubElement(cm, "meta-value")
            mv.text = txt

    # Body
    body = ET.SubElement(root, "body")
    current_sec: ET.Element | None = None

    def _ensure_sec() -> ET.Element:
        nonlocal current_sec
        if current_sec is None:
            current_sec = ET.SubElement(body, "sec")
        return current_sec

    for item in body_stream:
        if item["kind"] == "h":
            current_sec = ET.SubElement(body, "sec")
            title = ET.SubElement(current_sec, "title")
            title.text = item["text"]
            continue
        for ptxt in _split_paragraphs(item["text"]):
            sec = _ensure_sec()
            p = ET.SubElement(sec, "p")
            p.text = ptxt

    # Tablas al final del body
    if tablas:
        sec_tables = ET.SubElement(body, "sec", {"sec-type": "tables"})
        t_title = ET.SubElement(sec_tables, "title")
        t_title.text = "Tablas"
        for i, tab in enumerate(tablas, 1):
            tw = ET.SubElement(sec_tables, "table-wrap", {"id": f"tbl{i}"})
            label = ET.SubElement(tw, "label")
            label.text = f"Tabla {i}"
            caption = ET.SubElement(tw, "caption")
            cp = ET.SubElement(caption, "p")
            cp.text = _clean_text(tab.get("titulo", "")) or f"Tabla {i}"
            _append_table_xml(tw, tab.get("ruta", ""), tab.get("hoja"))

    # Figuras al final del body
    if figuras:
        sec_figs = ET.SubElement(body, "sec", {"sec-type": "figures"})
        f_title = ET.SubElement(sec_figs, "title")
        f_title.text = "Figuras"
        for i, fig in enumerate(figuras, 1):
            fg = ET.SubElement(sec_figs, "fig", {"id": f"fig{i}"})
            label = ET.SubElement(fg, "label")
            label.text = f"Figura {i}"
            cap = ET.SubElement(fg, "caption")
            p = ET.SubElement(cap, "p")
            p.text = _clean_text(fig.get("pie", "")) or f"Figura {i}"
            g = ET.SubElement(fg, "graphic")
            g.set(f"{{{XLINK_NS}}}href", os.path.basename(fig.get("ruta", f"fig{i}.png")))

    # Back
    back = ET.SubElement(root, "back")
    refs = referencias_externas if referencias_externas else refs_from_blocks
    refs = [_strip_ref_prefix(r) for r in refs if _strip_ref_prefix(r)]
    if refs:
        ref_list = ET.SubElement(back, "ref-list")
        rtitle = ET.SubElement(ref_list, "title")
        rtitle.text = "Referencias"
        for i, ref in enumerate(refs, 1):
            r = ET.SubElement(ref_list, "ref", {"id": f"R{i}"})
            rl = ET.SubElement(r, "label")
            rl.text = str(i)
            mc = ET.SubElement(r, "mixed-citation")
            mc.text = ref

    return _pretty_xml(root)
