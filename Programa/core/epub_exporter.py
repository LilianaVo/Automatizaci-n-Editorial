"""
core/epub_exporter.py
Función pura de construcción de EPUB 2.0.
No depende de tkinter, customtkinter ni de ningún framework de UI.

Uso:
    from core.epub_exporter import build_epub

    epub_bytes = build_epub(
        html_str = "<html>…</html>",   # salida de html_exporter.build_html()
        titulo   = "Título del artículo",
        autores  = ["Autor A", "Autor B"],
        doi      = "https://doi.org/10.xxxx/yyyy",   # puede ser cadena vacía
        secciones = [("introduccion", "Introducción"), …],  # list[tuple[ancla, texto]]
    )
    with open("articulo.epub", "wb") as f:
        f.write(epub_bytes)
"""

from __future__ import annotations

import re
import uuid
import zipfile
import io

from core.utils import esc


# ─── Helpers internos ─────────────────────────────────────────────────────────

def _html_a_xhtml(html_str: str) -> tuple[str, str]:
    """Separa CSS y body del HTML generado, y convierte el body a XHTML válido.

    Devuelve (css_str, body_xhtml).
    """
    # Extraer CSS
    css_match = re.search(r"<style>(.*?)</style>", html_str, re.DOTALL)
    css_str   = css_match.group(1) if css_match else ""
    # Quitar @import de Google Fonts (no funciona offline en EPUB)
    css_str = re.sub(r"@import\s+url\([^)]+\)\s*;?\s*", "", css_str)

    # Extraer body
    body_match = re.search(r"<body>(.*?)</body>", html_str, re.DOTALL)
    body_str   = body_match.group(1) if body_match else html_str

    # Normalizar a XHTML: etiquetas vacías auto-cerradas
    body_str = re.sub(r"<br\s*>",               "<br/>",       body_str)
    body_str = re.sub(r"<hr\s*>",               "<hr/>",       body_str)
    body_str = re.sub(r"<img([^>]*[^/])\s*>",   r"<img\1/>",   body_str)
    body_str = re.sub(r"<input([^>]*[^/])\s*>", r"<input\1/>", body_str)
    # Escapar & sueltos que no sean entidades ya válidas
    body_str = re.sub(
        r"&(?!amp;|lt;|gt;|quot;|apos;|#\d+;|#x[\da-fA-F]+;)",
        "&amp;",
        body_str,
    )

    return css_str, body_str


# ─── Función pública principal ────────────────────────────────────────────────

def build_epub(
    html_str:  str,
    titulo:    str,
    autores:   list[str],
    doi:       str = "",
    secciones: list[tuple[str, str]] | None = None,
) -> bytes:
    """Construye un EPUB 2.0 completo y lo devuelve como bytes.

    Parámetros
    ----------
    html_str:
        HTML completo producido por ``html_exporter.build_html()``.
    titulo:
        Título del artículo (para metadatos OPF y TOC).
    autores:
        Lista de nombres de autores (strings limpios, sin ORCID).
    doi:
        DOI del artículo como URL (``https://doi.org/…``).
        Cadena vacía si no está disponible.
    secciones:
        Lista de tuplas ``(ancla, texto_seccion)`` para el TOC.
        Si es ``None`` o vacía, el TOC contendrá solo el título del artículo.

    Retorna
    -------
    bytes — contenido binario del archivo .epub listo para escribir en disco.
    """
    if not secciones:
        secciones = []

    css_str, body_str = _html_a_xhtml(html_str)

    uid_libro = str(uuid.uuid4())

    # ── mimetype ──────────────────────────────────────────────────────────────
    mimetype_bytes = b"application/epub+zip"

    # ── META-INF/container.xml ────────────────────────────────────────────────
    container_xml = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<container version="1.0" '
        'xmlns="urn:oasis:names:tc:opendocument:xmlns:container">\n'
        '  <rootfiles>\n'
        '    <rootfile full-path="OEBPS/content.opf" '
        'media-type="application/oebps-package+xml"/>\n'
        '  </rootfiles>\n'
        '</container>'
    ).encode("utf-8")

    # ── OEBPS/content.opf ─────────────────────────────────────────────────────
    autores_opf = "\n    ".join(
        f"<dc:creator>{esc(a)}</dc:creator>" for a in autores
    ) if autores else "<dc:creator>Autor desconocido</dc:creator>"

    doi_opf = (
        f'\n    <dc:identifier id="doi">{esc(doi)}</dc:identifier>'
        if doi else ""
    )

    content_opf = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<package xmlns="http://www.idpf.org/2007/opf" '
        'version="2.0" unique-identifier="uid">\n'
        '  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:opf="http://www.idpf.org/2007/opf">\n'
        f'    <dc:identifier id="uid">{uid_libro}</dc:identifier>\n'
        f'    <dc:title>{esc(titulo)}</dc:title>\n'
        f'    {autores_opf}{doi_opf}\n'
        '    <dc:language>es</dc:language>\n'
        '    <dc:publisher>'
        'Paleontología Mexicana — Instituto de Geología, UNAM'
        '</dc:publisher>\n'
        '  </metadata>\n'
        '  <manifest>\n'
        '    <item id="article" href="article.xhtml" '
        'media-type="application/xhtml+xml"/>\n'
        '    <item id="css" href="style/main.css" '
        'media-type="text/css"/>\n'
        '    <item id="ncx" href="toc.ncx" '
        'media-type="application/x-dtbncx+xml"/>\n'
        '  </manifest>\n'
        '  <spine toc="ncx">\n'
        '    <itemref idref="article"/>\n'
        '  </spine>\n'
        '</package>'
    ).encode("utf-8")

    # ── OEBPS/toc.ncx ─────────────────────────────────────────────────────────
    nav_points = ""
    for i, (ancla, txt_sec) in enumerate(secciones, 1):
        nav_points += (
            f'  <navPoint id="nav{i}" playOrder="{i}">\n'
            f'    <navLabel><text>{esc(txt_sec)}</text></navLabel>\n'
            f'    <content src="article.xhtml"/>\n'
            f'  </navPoint>\n'
        )
    if not nav_points:
        nav_points = (
            '  <navPoint id="nav1" playOrder="1">\n'
            f'    <navLabel><text>{esc(titulo)}</text></navLabel>\n'
            '    <content src="article.xhtml"/>\n'
            '  </navPoint>\n'
        )

    toc_ncx = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<!DOCTYPE ncx PUBLIC "-//NISO//DTD ncx 2005-1//EN" '
        '"http://www.daisy.org/z3986/2005/ncx-2005-1.dtd">\n'
        '<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1">\n'
        '  <head>\n'
        f'    <meta name="dtb:uid" content="{uid_libro}"/>\n'
        '    <meta name="dtb:depth" content="1"/>\n'
        '    <meta name="dtb:totalPageCount" content="0"/>\n'
        '    <meta name="dtb:maxPageNumber" content="0"/>\n'
        '  </head>\n'
        f'  <docTitle><text>{esc(titulo)}</text></docTitle>\n'
        '  <navMap>\n'
        f'{nav_points}'
        '  </navMap>\n'
        '</ncx>'
    ).encode("utf-8")

    # ── OEBPS/article.xhtml ───────────────────────────────────────────────────
    article_xhtml = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<!DOCTYPE html>\n'
        '<html xmlns="http://www.w3.org/1999/xhtml" lang="es" xml:lang="es">\n'
        '<head>\n'
        f'  <title>{esc(titulo)}</title>\n'
        '  <link href="style/main.css" rel="stylesheet" type="text/css"/>\n'
        '</head>\n'
        f'<body>\n{body_str}\n</body>\n'
        '</html>'
    ).encode("utf-8")

    css_bytes = css_str.encode("utf-8")

    # ── Empaquetar en ZIP (EPUB 2.0) ──────────────────────────────────────────
    # ebooklib a veces genera ZIPs corruptos; construirlo a mano garantiza
    # que el mimetype quede SIN comprimir y SIN extra fields, como exige la spec.
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        # mimetype: DEBE ser la primera entrada, sin compresión
        info_mime = zipfile.ZipInfo("mimetype")
        info_mime.compress_type = zipfile.ZIP_STORED
        zf.writestr(info_mime, mimetype_bytes)

        zf.writestr("META-INF/container.xml", container_xml)
        zf.writestr("OEBPS/content.opf",      content_opf)
        zf.writestr("OEBPS/toc.ncx",          toc_ncx)
        zf.writestr("OEBPS/article.xhtml",    article_xhtml)
        zf.writestr("OEBPS/style/main.css",   css_bytes)

    return buf.getvalue()