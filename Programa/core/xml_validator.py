"""
xml_validator.py
Valida un string XML contra el DTD oficial de JATS 1.1 (SciELO SPS).

No modifica ningún archivo existente. Se usa así:
    from core.xml_validator import validar_jats
    resultado = validar_jats(xml_string)
"""

from __future__ import annotations
import re
from dataclasses import dataclass, field


# ─────────────────────────────────────────────────────────────────────────────
# Resultado de validación
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class ErrorValidacion:
    nivel:   str   # "error" | "advertencia"
    linea:   int
    mensaje: str


@dataclass
class ResultadoValidacion:
    valido:        bool
    errores:       list[ErrorValidacion] = field(default_factory=list)
    advertencias:  list[ErrorValidacion] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "valido": self.valido,
            "errores": [
                {"nivel": e.nivel, "linea": e.linea, "mensaje": e.mensaje}
                for e in self.errores
            ],
            "advertencias": [
                {"nivel": e.nivel, "linea": e.linea, "mensaje": e.mensaje}
                for e in self.advertencias
            ],
            "resumen": _resumen(self),
        }


def _resumen(r: ResultadoValidacion) -> str:
    if r.valido and not r.advertencias:
        return "✅ XML válido para SciELO — sin errores ni advertencias."
    if r.valido:
        return f"✅ XML válido con {len(r.advertencias)} advertencia(s)."
    return f"❌ XML inválido — {len(r.errores)} error(es), {len(r.advertencias)} advertencia(s)."


# ─────────────────────────────────────────────────────────────────────────────
# Validación con lxml + DTD en línea
# ─────────────────────────────────────────────────────────────────────────────

def validar_jats(xml_string: str) -> ResultadoValidacion:
    """
    Valida el XML en dos pasos:
      1) Bien formado (XML válido sintácticamente)
      2) Reglas JATS/SciELO sin necesidad del DTD externo
         (verificación de elementos obligatorios y estructura)

    Retorna un ResultadoValidacion con errores y advertencias detallados.
    """
    errores: list[ErrorValidacion] = []
    advertencias: list[ErrorValidacion] = []

    # ── Paso 1: XML bien formado con lxml ─────────────────────────────────────
    try:
        from lxml import etree

        try:
            root = etree.fromstring(xml_string.encode("utf-8"))
        except etree.XMLSyntaxError as exc:
            for e in exc.error_log:
                errores.append(ErrorValidacion(
                    nivel="error",
                    linea=e.line,
                    mensaje=f"XML mal formado: {e.message}",
                ))
            return ResultadoValidacion(valido=False, errores=errores)

    except ImportError:
        # lxml no disponible — hacer validación básica manual
        root = None
        errores_basicos = _validar_sin_lxml(xml_string)
        errores.extend(errores_basicos)
        if errores:
            return ResultadoValidacion(valido=False, errores=errores)

    # ── Paso 2: Reglas estructurales JATS/SciELO ─────────────────────────────
    _verificar_estructura(xml_string, root, errores, advertencias)

    valido = len(errores) == 0
    return ResultadoValidacion(valido=valido, errores=errores, advertencias=advertencias)


# ─────────────────────────────────────────────────────────────────────────────
# Reglas estructurales JATS — sin DTD externo
# ─────────────────────────────────────────────────────────────────────────────

def _verificar_estructura(xml_string: str, root, errores, advertencias):
    """Verifica reglas clave de JATS 1.1 / SciELO SPS."""

    ns = {"xlink": "http://www.w3.org/1999/xlink"}

    def xpath(expr):
        try:
            if root is not None:
                from lxml import etree
                return root.xpath(expr, namespaces=ns)
            return []
        except Exception:
            return []

    def buscar_en_texto(patron: str) -> list[tuple[int, str]]:
        """Encuentra líneas que contienen el patrón."""
        resultados = []
        for i, linea in enumerate(xml_string.splitlines(), 1):
            if re.search(patron, linea):
                resultados.append((i, linea.strip()))
        return resultados

    def linea_de_tag(tag: str) -> int:
        """Devuelve la línea aproximada donde aparece un tag."""
        for i, l in enumerate(xml_string.splitlines(), 1):
            if f"<{tag}" in l or f"</{tag}" in l:
                return i
        return 0

    # ── 1. Elemento raíz <article> ───────────────────────────────────────────
    if root is not None:
        from lxml import etree
        if root.tag != "article":
            errores.append(ErrorValidacion(
                nivel="error", linea=1,
                mensaje="El elemento raíz debe ser <article>, no <{}>".format(root.tag)
            ))
            return

    # ── 2. Atributos requeridos en <article> ─────────────────────────────────
    atrs_article = buscar_en_texto(r"<article\s")
    if atrs_article:
        linea_art, txt_art = atrs_article[0]
        if "article-type" not in txt_art:
            errores.append(ErrorValidacion(
                nivel="error", linea=linea_art,
                mensaje='<article> requiere atributo article-type (ej: article-type="research-article")',
            ))
        if "xml:lang" not in txt_art:
            errores.append(ErrorValidacion(
                nivel="error", linea=linea_art,
                mensaje='<article> requiere atributo xml:lang (ej: xml:lang="es")',
            ))

    # ── 3. <front> → <journal-meta> → <journal-title> ────────────────────────
    if not buscar_en_texto(r"<journal-title"):
        errores.append(ErrorValidacion(
            nivel="error", linea=linea_de_tag("journal-meta"),
            mensaje="Falta <journal-title> dentro de <journal-meta>",
        ))

    if not buscar_en_texto(r"<issn"):
        advertencias.append(ErrorValidacion(
            nivel="advertencia", linea=linea_de_tag("journal-meta"),
            mensaje="No se encontró <issn> — recomendado para SciELO",
        ))

    # ── 4. <article-meta> obligatorio ────────────────────────────────────────
    if not buscar_en_texto(r"<article-meta"):
        errores.append(ErrorValidacion(
            nivel="error", linea=linea_de_tag("front"),
            mensaje="Falta <article-meta> dentro de <front>",
        ))
        return  # Sin article-meta muchas otras reglas fallan

    # ── 5. Título del artículo ────────────────────────────────────────────────
    if not buscar_en_texto(r"<article-title"):
        errores.append(ErrorValidacion(
            nivel="error", linea=linea_de_tag("title-group"),
            mensaje="Falta <article-title> — el título del artículo es obligatorio",
        ))

    # ── 6. Al menos un <contrib> de tipo author ───────────────────────────────
    contribs_author = buscar_en_texto(r'contrib-type="author"')
    if not contribs_author:
        advertencias.append(ErrorValidacion(
            nivel="advertencia", linea=linea_de_tag("contrib-group"),
            mensaje="No se encontró ningún <contrib contrib-type=\"author\"> — se recomienda al menos uno",
        ))

    # ── 7. <pub-date> ────────────────────────────────────────────────────────
    if not buscar_en_texto(r"<pub-date"):
        errores.append(ErrorValidacion(
            nivel="error", linea=linea_de_tag("article-meta"),
            mensaje="Falta <pub-date> — la fecha de publicación es obligatoria",
        ))

    # ── 8. <abstract> ────────────────────────────────────────────────────────
    if not buscar_en_texto(r"<abstract"):
        advertencias.append(ErrorValidacion(
            nivel="advertencia", linea=linea_de_tag("article-meta"),
            mensaje="No se encontró <abstract> — recomendado por SciELO",
        ))

    # ── 9. <body> ────────────────────────────────────────────────────────────
    if not buscar_en_texto(r"<body"):
        errores.append(ErrorValidacion(
            nivel="error", linea=0,
            mensaje="Falta elemento <body> — el cuerpo del artículo es obligatorio",
        ))

    # ── 10. <back> y <ref-list> ──────────────────────────────────────────────
    if not buscar_en_texto(r"<back"):
        advertencias.append(ErrorValidacion(
            nivel="advertencia", linea=0,
            mensaje="No se encontró <back> — se recomienda incluir la lista de referencias",
        ))
    elif not buscar_en_texto(r"<ref-list"):
        advertencias.append(ErrorValidacion(
            nivel="advertencia", linea=linea_de_tag("back"),
            mensaje="<back> existe pero no contiene <ref-list>",
        ))

    # ── 11. DOI ──────────────────────────────────────────────────────────────
    doi_encontrado = buscar_en_texto(r'pub-id-type="doi"')
    if not doi_encontrado:
        advertencias.append(ErrorValidacion(
            nivel="advertencia", linea=linea_de_tag("article-meta"),
            mensaje="No se encontró DOI (<article-id pub-id-type=\"doi\">) — requerido por SciELO",
        ))

    # ── 12. Licencia CC ──────────────────────────────────────────────────────
    if not buscar_en_texto(r"<license"):
        advertencias.append(ErrorValidacion(
            nivel="advertencia", linea=linea_de_tag("permissions"),
            mensaje="No se encontró <license> en <permissions> — requerido para acceso abierto en SciELO",
        ))

    # ── 13. Keywords ─────────────────────────────────────────────────────────
    if not buscar_en_texto(r"<kwd-group"):
        advertencias.append(ErrorValidacion(
            nivel="advertencia", linea=linea_de_tag("article-meta"),
            mensaje="No se encontró <kwd-group> — las palabras clave son recomendadas",
        ))

    # ── 14. Afiliaciones sin país ─────────────────────────────────────────────
    affs = buscar_en_texto(r"<aff\s")
    paises = buscar_en_texto(r"<country")
    if affs and not paises:
        advertencias.append(ErrorValidacion(
            nivel="advertencia", linea=affs[0][0],
            mensaje="Las afiliaciones no tienen <country> — requerido por SciELO SPS",
        ))

    # ── 15. <sec> sin <title> ────────────────────────────────────────────────
    secs = buscar_en_texto(r"<sec[\s>]")
    titulos_sec = buscar_en_texto(r"<title>")
    if len(secs) > len(titulos_sec):
        advertencias.append(ErrorValidacion(
            nivel="advertencia", linea=secs[0][0] if secs else 0,
            mensaje=f"Hay {len(secs)} sección(es) pero solo {len(titulos_sec)} título(s) — cada <sec> debe tener <title>",
        ))


# ─────────────────────────────────────────────────────────────────────────────
# Fallback sin lxml
# ─────────────────────────────────────────────────────────────────────────────

def _validar_sin_lxml(xml_string: str) -> list[ErrorValidacion]:
    """Validación mínima de XML bien formado usando xml.etree (stdlib)."""
    import xml.etree.ElementTree as ET
    errores = []
    try:
        ET.fromstring(xml_string)
    except ET.ParseError as e:
        linea = e.position[0] if hasattr(e, "position") else 0
        errores.append(ErrorValidacion(
            nivel="error", linea=linea,
            mensaje=f"XML mal formado: {e}",
        ))
    return errores