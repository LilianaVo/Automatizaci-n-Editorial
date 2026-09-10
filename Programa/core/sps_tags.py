"""core/sps_tags.py
Biblioteca de etiquetas SciELO Markup / SPS (SciELO Publishing Schema) con su
correspondencia a los nombres de la JATS Publishing Tag Library 1.4.

Fuentes:
  · La marcación de corchetes de SciELO Markup en artículos reales de
    Paleontología Mexicana (a1-55-70) → nombres de corchete y anidación real.
  · La guía «Etiquetas_SciELO_para_Editor_Semantico.md» (listas oficiales de
    SciELO PC Programs) → vocabularios de atributos verificados.
  · JATS Publishing Tag Library 1.4 (https://jats.nlm.nih.gov/publishing/
    tag-library/1.4/) → el campo `jats` con el nombre exacto del elemento.

Solo datos + reglas: no toca la UI ni el exportador. Por cada etiqueta define su
equivalente JATS (`jats`), dónde es válida (`padres`), qué admite (`hijos`) y sus
atributos con vocabularios controlados. Con eso la interfaz puede ofrecer solo
etiquetas válidas por contexto e impedir anidaciones inválidas.

Clave canónica interna = nombre de corchete de Markup (p. ej. "figgrp", "tabwrap").
`jats` = None en agrupadores de Markup sin elemento JATS 1:1 (p. ej. `confgrp`,
`thesgrp`): sus hijos se emiten directos en el elemento contenedor.
"""

from __future__ import annotations

SPS_VERSION = "1.9"
JATS_TAG_LIBRARY = "1.4"

# ─────────────────────────────────────────────────────────────────────────────
# Vocabularios controlados (de la guía .md, verificados contra SciELO/JATS)
# ─────────────────────────────────────────────────────────────────────────────

# doctopic (en [doc]): código SciELO de 2 letras → (descripción, article-type JATS).
# "" en article-type = sin equivalente estándar directo (lo resuelve el conversor).
DOCTOPICS: list[tuple[str, str, str]] = [
    ("oa", "Artículo original / de investigación", "research-article"),
    ("ra", "Artículo de revisión",                 "review-article"),
    ("cr", "Informe de caso",                       "case-report"),
    ("ed", "Editorial",                             "editorial"),
    ("er", "Corrección / errata",                   "correction"),
    ("in", "Entrevista",                            "interview"),
    ("le", "Carta",                                 "letter"),
    ("co", "Comentario",                            "article-commentary"),
    ("an", "Anuncio",                               "announcement"),
    ("pr", "Nota de prensa",                        "in-brief"),
    ("ab", "Resúmenes",                             "abstract"),
    ("pv", "Punto de vista",                        "article-commentary"),
    ("sc", "Comunicación breve",                    "rapid-communication"),
    ("ax", "Anexo",                                 ""),
    ("ct", "Ensayo clínico",                        ""),
    ("mt", "Metodología",                           ""),
    ("rc", "Recuento",                              ""),
    ("rn", "Nota de investigación",                 ""),
    ("tr", "Informe técnico",                       ""),
    ("up", "Actualización",                         ""),
]
DOCTOPIC_CODES = [c for c, _, _ in DOCTOPICS]
DOCTOPIC_A_ARTICLE_TYPE = {c: at for c, _, at in DOCTOPICS}
DOCTOPIC_DESC = {c: d for c, d, _ in DOCTOPICS}

# sec-type (en [sec]/[subsec]): valores atómicos; los combinados usan «|».
SEC_TYPE = [
    "nd", "intro", "materials", "methods", "results", "discussion",
    "conclusions", "cases", "subjects", "supplementary-material",
]
# xref @ref-type — Apéndice B de la guía (lista SciELO).
XREF_REF_TYPE = [
    "aff", "app", "author-notes", "bibr", "boxed-text", "contrib", "corresp",
    "disp-formula", "fig", "fn", "kwd", "list", "plate", "scheme", "sec",
    "statement", "supplementary-material", "table", "chem", "other",
]
# ref @reftype — tipo de referencia bibliográfica.
REF_TYPE = [
    "journal", "book", "confproc", "thesis", "webpage", "data", "report",
    "patent", "newspaper", "legal-doc", "database", "software", "other",
]
AUTHIDTP = ["orcid", "lattes", "researchid", "scopus"]            # authorid @authidtp
PUBID_IDTYPE = [                                                   # pubid @idtype
    "doi", "pmid", "pmcid", "art-access-id", "coden", "doaj", "medline",
    "manuscript", "rrn", "pii", "publisher-id", "sici", "other",
]
LIST_TYPE = ["order", "bullet", "alpha-lower", "alpha-upper",
             "roman-lower", "roman-upper", "simple"]               # list @listtype
ROLE = ["nd", "coord", "ed", "org", "tr"]                         # author/authors @role
TD_ALIGN = ["left", "center", "right", "justify"]                 # td @align
SN2 = ["y", "n"]                                                  # sí/no
SN3 = ["y", "n", "nd"]                                            # sí/no/no-definido

# ─────────────────────────────────────────────────────────────────────────────
# Conjuntos de apoyo
# ─────────────────────────────────────────────────────────────────────────────
INLINE = ["xref", "sup", "sub", "italic", "bold", "sc", "namedcontent"]

TEXTO_INLINE = [
    "p", "sectitle", "caption", "label", "li", "kwd", "funding",
    "source", "arttitle", "chptitle", "pubname", "publoc", "moreinfo",
    "confname", "degree", "doctitle", "toctitle", "corresp", "td", "th",
]
_INLINE_ANIDABLE = ["sup", "sub", "italic", "bold", "sc", "namedcontent"]

_TXT = "#TEXTO"
_INL = "#INLINE"


def _a(req=False, valores=None):
    return {"req": req, "valores": valores}


# ─────────────────────────────────────────────────────────────────────────────
# Definición de etiquetas
#   jats  – elemento JATS 1.4 equivalente (None = agrupador Markup sin 1:1)
#   tipo  – raiz|meta|estructural|bloque|inline|cita|tabla
#   grupo – agrupación para la UI
#   etiqueta – nombre amable (ES)
#   padres/hijos – reglas de anidación (padres = autoridad)
#   attrs – {nombre: {"req": bool, "valores": [...]|None}}
#   vacio – elemento vacío
#   nota  – aclaración (opcional)
# ─────────────────────────────────────────────────────────────────────────────
TAGS: dict[str, dict] = {

    # ── Raíz / documento ──────────────────────────────────────────────────────
    "doc": {
        "jats": "article", "tipo": "raiz", "grupo": "documento",
        "etiqueta": "Documento",
        "padres": ["#RAIZ"],
        "hijos": ["doi", "doctitle", "toctitle", "author", "normaff", "corresp",
                  "xmlabstr", "kwdgrp", "hist", "xmlbody", "refs", "ack"],
        "attrs": {
            "sps": _a(True), "acron": _a(), "jtitle": _a(), "stitle": _a(),
            "issn": _a(), "pissn": _a(), "eissn": _a(), "pubname": _a(),
            "license": _a(), "volid": _a(), "issueno": _a(), "dateiso": _a(),
            "artdate": _a(), "season": _a(), "order": _a(), "fpage": _a(),
            "lpage": _a(), "pagcount": _a(),
            "doctopic": _a(valores=DOCTOPIC_CODES),
            "language": _a(valores=["es", "en"]),
        },
    },

    # ── Front / metadatos ─────────────────────────────────────────────────────
    "doi": {"jats": "article-id", "tipo": "meta", "grupo": "front",
            "etiqueta": "DOI", "padres": ["doc"], "hijos": [_TXT], "attrs": {}},
    "doctitle": {"jats": "article-title", "tipo": "meta", "grupo": "front",
                 "etiqueta": "Título del artículo", "padres": ["doc"],
                 "hijos": [_TXT, _INL], "attrs": {"language": _a(True, ["es", "en"])}},
    "toctitle": {"jats": None, "tipo": "meta", "grupo": "front",
                 "etiqueta": "Título para TOC", "padres": ["doc"],
                 "hijos": [_TXT, _INL], "attrs": {},
                 "nota": "Título para tabla de contenidos; sin elemento JATS 1:1."},

    "author": {"jats": "contrib", "tipo": "meta", "grupo": "front",
               "etiqueta": "Autor", "padres": ["doc"],
               "hijos": ["surname", "fname", "xref", "authorid"],
               "attrs": {"role": _a(valores=ROLE), "rid": _a(),
                         "corresp": _a(valores=SN2), "deceased": _a(valores=SN2),
                         "eqcontr": _a(valores=SN3)}},
    "surname": {"jats": "surname", "tipo": "meta", "grupo": "front",
                "etiqueta": "Apellido", "padres": ["author", "pauthor"],
                "hijos": [_TXT], "attrs": {}},
    "fname": {"jats": "given-names", "tipo": "meta", "grupo": "front",
              "etiqueta": "Nombre(s)", "padres": ["author", "pauthor"],
              "hijos": [_TXT], "attrs": {}},
    "authorid": {"jats": "contrib-id", "tipo": "meta", "grupo": "front",
                 "etiqueta": "ID de autor (ORCID…)", "padres": ["author"],
                 "hijos": [_TXT], "attrs": {"authidtp": _a(True, AUTHIDTP)}},

    "normaff": {"jats": "aff", "tipo": "meta", "grupo": "front",
                "etiqueta": "Afiliación", "padres": ["doc"],
                "hijos": ["label", "orgdiv1", "orgdiv2", "orgname",
                          "city", "state", "country", "zipcode"],
                "attrs": {"id": _a(True), "ncountry": _a(), "norgname": _a(),
                          "icountry": _a()}},
    "orgdiv1": {"jats": "institution", "tipo": "meta", "grupo": "front",
                "etiqueta": "Dependencia (nivel 1)", "padres": ["normaff"],
                "hijos": [_TXT], "attrs": {},
                "nota": "JATS: <institution content-type=\"orgdiv1\">."},
    "orgdiv2": {"jats": "institution", "tipo": "meta", "grupo": "front",
                "etiqueta": "Dependencia (nivel 2)", "padres": ["normaff"],
                "hijos": [_TXT], "attrs": {},
                "nota": "JATS: <institution content-type=\"orgdiv2\">."},
    "orgname": {"jats": "institution", "tipo": "meta", "grupo": "front",
                "etiqueta": "Institución", "padres": ["normaff", "thesgrp"],
                "hijos": [_TXT], "attrs": {},
                "nota": "JATS: <institution content-type=\"orgname\">."},
    "city": {"jats": "city", "tipo": "meta", "grupo": "front",
             "etiqueta": "Ciudad", "padres": ["normaff", "confgrp"],
             "hijos": [_TXT], "attrs": {}},
    "state": {"jats": "state", "tipo": "meta", "grupo": "front",
              "etiqueta": "Estado/Provincia", "padres": ["normaff", "confgrp"],
              "hijos": [_TXT], "attrs": {}},
    "country": {"jats": "country", "tipo": "meta", "grupo": "front",
                "etiqueta": "País", "padres": ["normaff", "confgrp"],
                "hijos": [_TXT], "attrs": {}},
    "zipcode": {"jats": "postal-code", "tipo": "meta", "grupo": "front",
                "etiqueta": "Código postal", "padres": ["normaff"],
                "hijos": [_TXT], "attrs": {}},

    "corresp": {"jats": "corresp", "tipo": "meta", "grupo": "front",
                "etiqueta": "Correspondencia", "padres": ["doc"],
                "hijos": [_TXT, _INL], "attrs": {"id": _a(True)}},

    "xmlabstr": {"jats": "abstract", "tipo": "estructural", "grupo": "front",
                 "etiqueta": "Resumen", "padres": ["doc"],
                 "hijos": ["sectitle", "p"], "attrs": {"language": _a(True, ["es", "en"])}},
    "kwdgrp": {"jats": "kwd-group", "tipo": "estructural", "grupo": "front",
               "etiqueta": "Grupo de palabras clave", "padres": ["doc"],
               "hijos": ["sectitle", "kwd"], "attrs": {"language": _a(True, ["es", "en"])}},
    "kwd": {"jats": "kwd", "tipo": "bloque", "grupo": "front",
            "etiqueta": "Palabra clave", "padres": ["kwdgrp"],
            "hijos": [_TXT, _INL], "attrs": {}},

    "hist": {"jats": "history", "tipo": "estructural", "grupo": "front",
             "etiqueta": "Historial", "padres": ["doc"],
             "hijos": ["received", "revised", "accepted"], "attrs": {}},
    "received": {"jats": "date", "tipo": "meta", "grupo": "front",
                 "etiqueta": "Recibido", "padres": ["hist"],
                 "hijos": [_TXT], "attrs": {"dateiso": _a(True)},
                 "nota": "JATS: <date date-type=\"received\">."},
    "revised": {"jats": "date", "tipo": "meta", "grupo": "front",
                "etiqueta": "Corregido", "padres": ["hist"],
                "hijos": [_TXT], "attrs": {"dateiso": _a(True)},
                "nota": "JATS: <date date-type=\"rev-recd\">."},
    "accepted": {"jats": "date", "tipo": "meta", "grupo": "front",
                 "etiqueta": "Aceptado", "padres": ["hist"],
                 "hijos": [_TXT], "attrs": {"dateiso": _a(True)},
                 "nota": "JATS: <date date-type=\"accepted\">."},

    # ── Cuerpo ────────────────────────────────────────────────────────────────
    "xmlbody": {"jats": "body", "tipo": "estructural", "grupo": "cuerpo",
                "etiqueta": "Cuerpo", "padres": ["doc"],
                "hijos": ["sec"], "attrs": {}},
    "sec": {"jats": "sec", "tipo": "estructural", "grupo": "cuerpo",
            "etiqueta": "Sección", "padres": ["xmlbody"],
            "hijos": ["sectitle", "p", "subsec", "figgrp", "tabwrap", "list"],
            "attrs": {"sec-type": _a(valores=SEC_TYPE)}},
    "subsec": {"jats": "sec", "tipo": "estructural", "grupo": "cuerpo",
               "etiqueta": "Subsección", "padres": ["sec", "subsec"],
               "hijos": ["sectitle", "p", "subsec", "figgrp", "tabwrap", "list"],
               "attrs": {"sec-type": _a(valores=SEC_TYPE)}},
    "sectitle": {"jats": "title", "tipo": "bloque", "grupo": "cuerpo",
                 "etiqueta": "Título de sección",
                 "padres": ["sec", "subsec", "xmlabstr", "kwdgrp", "refs", "ack"],
                 "hijos": [_TXT, _INL], "attrs": {}},
    "p": {"jats": "p", "tipo": "bloque", "grupo": "cuerpo",
          "etiqueta": "Párrafo", "padres": ["sec", "subsec", "xmlabstr", "ack"],
          "hijos": [_TXT, _INL, "funding"], "attrs": {}},
    "list": {"jats": "list", "tipo": "estructural", "grupo": "cuerpo",
             "etiqueta": "Lista", "padres": ["sec", "subsec", "p"],
             "hijos": ["li"], "attrs": {"listtype": _a(valores=LIST_TYPE)}},
    "li": {"jats": "list-item", "tipo": "bloque", "grupo": "cuerpo",
           "etiqueta": "Elemento de lista", "padres": ["list"],
           "hijos": [_TXT, _INL], "attrs": {}},
    "funding": {"jats": "funding-statement", "tipo": "bloque", "grupo": "cuerpo",
                "etiqueta": "Financiamiento", "padres": ["p"],
                "hijos": [_TXT, _INL], "attrs": {}},

    # ── Figuras y tablas ──────────────────────────────────────────────────────
    "figgrp": {"jats": "fig", "tipo": "estructural", "grupo": "figuras-tablas",
               "etiqueta": "Figura", "padres": ["sec", "subsec"],
               "hijos": ["graphic", "label", "caption"], "attrs": {"id": _a(True)}},
    "graphic": {"jats": "graphic", "tipo": "estructural", "grupo": "figuras-tablas",
                "etiqueta": "Imagen", "padres": ["figgrp"],
                "hijos": [], "vacio": True, "attrs": {"href": _a(True)}},
    "tabwrap": {"jats": "table-wrap", "tipo": "estructural", "grupo": "figuras-tablas",
                "etiqueta": "Tabla", "padres": ["sec", "subsec"],
                "hijos": ["label", "caption", "table"], "attrs": {"id": _a(True)}},
    "table": {"jats": "table", "tipo": "tabla", "grupo": "figuras-tablas",
              "etiqueta": "Cuerpo de tabla", "padres": ["tabwrap"],
              "hijos": ["thead", "tbody"], "attrs": {}},
    "thead": {"jats": "thead", "tipo": "tabla", "grupo": "figuras-tablas",
              "etiqueta": "Encabezado de tabla", "padres": ["table"],
              "hijos": ["tr"], "attrs": {}},
    "tbody": {"jats": "tbody", "tipo": "tabla", "grupo": "figuras-tablas",
              "etiqueta": "Cuerpo de filas", "padres": ["table"],
              "hijos": ["tr"], "attrs": {}},
    "tr": {"jats": "tr", "tipo": "tabla", "grupo": "figuras-tablas",
           "etiqueta": "Fila", "padres": ["thead", "tbody"],
           "hijos": ["th", "td"], "attrs": {}},
    "th": {"jats": "th", "tipo": "tabla", "grupo": "figuras-tablas",
           "etiqueta": "Celda de encabezado", "padres": ["tr"],
           "hijos": [_TXT, _INL],
           "attrs": {"align": _a(valores=TD_ALIGN), "colspan": _a(), "rowspan": _a()}},
    "td": {"jats": "td", "tipo": "tabla", "grupo": "figuras-tablas",
           "etiqueta": "Celda", "padres": ["tr"], "hijos": [_TXT, _INL],
           "attrs": {"align": _a(valores=TD_ALIGN), "colspan": _a(), "rowspan": _a()}},
    "label": {"jats": "label", "tipo": "bloque", "grupo": "figuras-tablas",
              "etiqueta": "Rótulo", "padres": ["figgrp", "tabwrap", "normaff"],
              "hijos": [_TXT, _INL], "attrs": {}},
    "caption": {"jats": "caption", "tipo": "bloque", "grupo": "figuras-tablas",
                "etiqueta": "Leyenda/Descripción", "padres": ["figgrp", "tabwrap"],
                "hijos": [_TXT, _INL], "attrs": {}},

    # ── Referencias (back) ────────────────────────────────────────────────────
    "refs": {"jats": "ref-list", "tipo": "estructural", "grupo": "referencias",
             "etiqueta": "Lista de referencias", "padres": ["doc"],
             "hijos": ["sectitle", "ref"], "attrs": {}},
    "ref": {"jats": "element-citation", "tipo": "cita", "grupo": "referencias",
            "etiqueta": "Referencia", "padres": ["refs"],
            "hijos": ["authors", "date", "source", "arttitle", "chptitle",
                      "pages", "volid", "issueno", "pubname", "publoc", "pubid",
                      "series", "confgrp", "thesgrp", "moreinfo", "url"],
            "attrs": {"id": _a(True), "reftype": _a(valores=REF_TYPE)}},
    "authors": {"jats": "person-group", "tipo": "estructural", "grupo": "referencias",
                "etiqueta": "Grupo de autores (cita)", "padres": ["ref"],
                "hijos": ["pauthor"], "attrs": {"role": _a(valores=ROLE)}},
    "pauthor": {"jats": "name", "tipo": "estructural", "grupo": "referencias",
                "etiqueta": "Autor (cita)", "padres": ["authors"],
                "hijos": ["surname", "fname"], "attrs": {}},
    "date": {"jats": "year", "tipo": "cita", "grupo": "referencias",
             "etiqueta": "Fecha (cita)", "padres": ["ref"],
             "hijos": [_TXT], "attrs": {"dateiso": _a(), "specyear": _a()},
             "nota": "En la cita JATS se emite como <year> (y afines)."},
    "source": {"jats": "source", "tipo": "cita", "grupo": "referencias",
               "etiqueta": "Fuente (revista/libro)", "padres": ["ref"],
               "hijos": [_TXT, _INL], "attrs": {}},
    "arttitle": {"jats": "article-title", "tipo": "cita", "grupo": "referencias",
                 "etiqueta": "Título del artículo (cita)", "padres": ["ref"],
                 "hijos": [_TXT, _INL], "attrs": {}},
    "chptitle": {"jats": "chapter-title", "tipo": "cita", "grupo": "referencias",
                 "etiqueta": "Título de capítulo (cita)", "padres": ["ref"],
                 "hijos": [_TXT, _INL], "attrs": {}},
    "pages": {"jats": "page-range", "tipo": "cita", "grupo": "referencias",
              "etiqueta": "Páginas (cita)", "padres": ["ref"],
              "hijos": [_TXT], "attrs": {}},
    "volid": {"jats": "volume", "tipo": "cita", "grupo": "referencias",
              "etiqueta": "Volumen (cita)", "padres": ["ref"],
              "hijos": [_TXT], "attrs": {}},
    "issueno": {"jats": "issue", "tipo": "cita", "grupo": "referencias",
                "etiqueta": "Número (cita)", "padres": ["ref"],
                "hijos": [_TXT], "attrs": {}},
    "pubname": {"jats": "publisher-name", "tipo": "cita", "grupo": "referencias",
                "etiqueta": "Editorial (cita)", "padres": ["ref"],
                "hijos": [_TXT, _INL], "attrs": {}},
    "publoc": {"jats": "publisher-loc", "tipo": "cita", "grupo": "referencias",
               "etiqueta": "Lugar de edición (cita)", "padres": ["ref"],
               "hijos": [_TXT, _INL], "attrs": {}},
    "pubid": {"jats": "pub-id", "tipo": "cita", "grupo": "referencias",
              "etiqueta": "ID de publicación (cita)", "padres": ["ref"],
              "hijos": [_TXT], "attrs": {"idtype": _a(True, PUBID_IDTYPE)}},
    "series": {"jats": "series", "tipo": "cita", "grupo": "referencias",
               "etiqueta": "Serie (cita)", "padres": ["ref"],
               "hijos": [_TXT, _INL], "attrs": {}},
    "confgrp": {"jats": None, "tipo": "estructural", "grupo": "referencias",
                "etiqueta": "Congreso (cita)", "padres": ["ref"],
                "hijos": ["confname", "no", "city", "state", "country"], "attrs": {},
                "nota": "Agrupador Markup; en JATS los conf-* van directos en la cita."},
    "confname": {"jats": "conf-name", "tipo": "cita", "grupo": "referencias",
                 "etiqueta": "Nombre del congreso", "padres": ["confgrp"],
                 "hijos": [_TXT, _INL], "attrs": {}},
    "no": {"jats": "conf-num", "tipo": "cita", "grupo": "referencias",
           "etiqueta": "Número de congreso", "padres": ["confgrp"],
           "hijos": [_TXT], "attrs": {}},
    "thesgrp": {"jats": None, "tipo": "estructural", "grupo": "referencias",
                "etiqueta": "Tesis (cita)", "padres": ["ref"],
                "hijos": ["degree", "orgname"], "attrs": {},
                "nota": "Agrupador Markup; sin wrapper JATS propio."},
    "degree": {"jats": None, "tipo": "cita", "grupo": "referencias",
               "etiqueta": "Grado (tesis)", "padres": ["thesgrp"],
               "hijos": [_TXT, _INL], "attrs": {},
               "nota": "Sin elemento JATS 1:1 en element-citation."},
    "moreinfo": {"jats": "comment", "tipo": "cita", "grupo": "referencias",
                 "etiqueta": "Info adicional (cita)", "padres": ["ref"],
                 "hijos": [_TXT, _INL], "attrs": {}},
    "url": {"jats": "ext-link", "tipo": "cita", "grupo": "referencias",
            "etiqueta": "URL (cita)", "padres": ["ref"],
            "hijos": [_TXT], "attrs": {}},

    # ── Back varios ───────────────────────────────────────────────────────────
    "ack": {"jats": "ack", "tipo": "estructural", "grupo": "back",
            "etiqueta": "Agradecimientos", "padres": ["doc"],
            "hijos": ["sectitle", "p"], "attrs": {}},

    # ── Inline ────────────────────────────────────────────────────────────────
    "xref": {"jats": "xref", "tipo": "inline", "grupo": "inline",
             "etiqueta": "Referencia cruzada", "padres": list(TEXTO_INLINE),
             "hijos": [_TXT], "attrs": {"ref-type": _a(True, XREF_REF_TYPE),
                                        "rid": _a(True), "label": _a()}},
    "sup": {"jats": "sup", "tipo": "inline", "grupo": "inline",
            "etiqueta": "Superíndice", "padres": list(TEXTO_INLINE) + _INLINE_ANIDABLE,
            "hijos": [_TXT, _INL], "attrs": {}},
    "sub": {"jats": "sub", "tipo": "inline", "grupo": "inline",
            "etiqueta": "Subíndice", "padres": list(TEXTO_INLINE) + _INLINE_ANIDABLE,
            "hijos": [_TXT, _INL], "attrs": {}},
    "italic": {"jats": "italic", "tipo": "inline", "grupo": "inline",
               "etiqueta": "Cursiva (taxón, término)",
               "padres": list(TEXTO_INLINE) + _INLINE_ANIDABLE,
               "hijos": [_TXT, _INL], "attrs": {}},
    "bold": {"jats": "bold", "tipo": "inline", "grupo": "inline",
             "etiqueta": "Negrita", "padres": list(TEXTO_INLINE) + _INLINE_ANIDABLE,
             "hijos": [_TXT, _INL], "attrs": {}},
    "sc": {"jats": "sc", "tipo": "inline", "grupo": "inline",
           "etiqueta": "Versalita", "padres": list(TEXTO_INLINE) + _INLINE_ANIDABLE,
           "hijos": [_TXT, _INL], "attrs": {}},
    "namedcontent": {"jats": "named-content", "tipo": "inline", "grupo": "inline",
                     "etiqueta": "Contenido nombrado",
                     "padres": list(TEXTO_INLINE) + _INLINE_ANIDABLE,
                     "hijos": [_TXT, _INL], "attrs": {"content-type": _a()}},
}

# Nombres de corchete Markup por confirmar contra el manual (el JATS es firme).
POR_CONFIRMAR = {"sub", "italic", "bold", "sc", "namedcontent"}

GRUPOS_UI = ["documento", "front", "cuerpo", "figuras-tablas", "referencias",
             "back", "inline"]

# ─────────────────────────────────────────────────────────────────────────────
# API de consulta (la consumen la UI y el validador)
# ─────────────────────────────────────────────────────────────────────────────

def existe(tag: str) -> bool:
    return tag in TAGS

def jats_de(tag: str) -> str | None:
    d = TAGS.get(tag)
    return d.get("jats") if d else None

def es_inline(tag: str) -> bool:
    return TAGS.get(tag, {}).get("tipo") == "inline"

def es_vacio(tag: str) -> bool:
    return bool(TAGS.get(tag, {}).get("vacio"))

def admite_texto(tag: str) -> bool:
    return _TXT in TAGS.get(tag, {}).get("hijos", [])

def _hijos_expandidos(tag: str) -> set[str]:
    hijos = TAGS.get(tag, {}).get("hijos", [])
    out: set[str] = set()
    for h in hijos:
        if h == _INL:
            out.update(INLINE)
        elif h != _TXT:
            out.add(h)
    return out

def hijos_validos(tag: str) -> list[str]:
    return sorted(_hijos_expandidos(tag))

def padres_validos(tag: str) -> list[str]:
    return list(TAGS.get(tag, {}).get("padres", []))

def puede_anidar(padre: str, hijo: str) -> bool:
    """Regla central para «impedir anidación inválida»."""
    if hijo not in TAGS or padre not in TAGS:
        return False
    if hijo in _hijos_expandidos(padre):
        return True
    return padre in TAGS[hijo].get("padres", [])

def tags_validas_en(padre: str, solo_inline: bool | None = None) -> list[str]:
    cand = _hijos_expandidos(padre)
    for t, d in TAGS.items():
        if padre in d.get("padres", []):
            cand.add(t)
    if solo_inline is True:
        cand = {t for t in cand if es_inline(t)}
    elif solo_inline is False:
        cand = {t for t in cand if not es_inline(t)}
    return sorted(cand)

def atributos(tag: str) -> dict:
    return TAGS.get(tag, {}).get("attrs", {})

def valores_control(tag: str, attr: str) -> list[str] | None:
    return TAGS.get(tag, {}).get("attrs", {}).get(attr, {}).get("valores")

def article_type_de_doctopic(codigo: str) -> str:
    """Convierte el código doctopic de SciELO (p. ej. 'oa') a article-type JATS."""
    return DOCTOPIC_A_ARTICLE_TYPE.get(codigo, "")
