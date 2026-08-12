"""
core/constants.py
Constantes de dominio: clasificaciones semánticas, colores de UI,
estilos de textbox y CSS editorial para exportación HTML.
"""

# ─── Clasificaciones semánticas ───────────────────────────────────────────────

OPCIONES = [
    "Cuerpo",
    "Cuerpo del abstract",
    "Título principal", "Título secundario",
    "Encabezado sección",
    "Subencabezado", "Subencabezado-bajo",
    "Palabras clave",
    "Referencia",
    "Cómo citar", "Fecha manuscrito",
    "Título tabla", "Pie de figura",
    "Filiación", "Email / Metadatos",
    "Ignorar",
]

# Mapeo de clases antiguas → clase equivalente actual (compatibilidad)
CLASE_COMPAT = {
    "Normal":             "Cuerpo",
    "Autores":            "Cuerpo",         # autores vienen de la pestaña ORCID
    "Resumen / Abstract": "Cuerpo",         # el estilo lo da el Encabezado sección
    "Imagen":             "Ignorar",
}

# ─── Colores de UI por clase ───────────────────────────────────────────────────
# Lista de (clase, hex_color, nombre_legible)
# Usada para la leyenda y para colorear los frames de bloques en la UI.

COLORES_UI = [
    ("Título principal",    "#1a237e", "Azul marino"),
    ("Título secundario",   "#283593", "Azul índigo"),
    ("Encabezado sección",  "#0277bd", "Azul"),
    ("Subencabezado",       "#00695c", "Verde azulado"),
    ("Subencabezado-bajo",  "#00796b", "Verde azulado claro"),
    ("Cuerpo",              "#212121", "Negro"),
    ("Cuerpo del abstract", "#455a64", "Gris azulado"),
    ("Palabras clave",      "#6a1b9a", "Morado"),
    ("Referencia",          "#424242", "Gris"),
    ("Cómo citar",          "#e65100", "Naranja"),
    ("Fecha manuscrito",    "#bf360c", "Rojo ladrillo"),
    ("Título tabla",        "#1565c0", "Azul tabla"),
    ("Pie de figura",       "#558b2f", "Verde oliva"),
    ("Filiación",           "#2e7d32", "Verde"),
    ("Email / Metadatos",   "#388e3c", "Verde medio"),
    ("Ignorar",             "#c62828", "Rojo"),
]

# Acceso rápido: clase → color hex
COLOR_POR_CLASE: dict[str, str] = {c: col for c, col, _ in COLORES_UI}

# ─── RF-05: Categorías / doctopic del documento ───────────────────────────────
# Cada entrada: (clave interna, etiqueta visible, article-type JATS, subject SciELO)
DOCTOPICS = [
    ("research-article", "Artículo de investigación", "research-article", "Research Article"),
    ("review-article",   "Artículo de revisión",       "review-article",   "Review Article"),
    ("case-report",      "Nota / Reporte de caso",     "case-report",      "Case Report"),
    ("other",            "Otro",                       "other",            "Other"),
]
DOCTOPIC_POR_CLAVE = {d[0]: d for d in DOCTOPICS}
DOCTOPIC_DEFAULT = "research-article"

# ─── RF-42: Licencias de la revista ────────────────────────────────────────────
# Catálogo de licencias comunes (con su href real de creativecommons.org) más
# una opción "otra" que activa un campo de texto libre en la UI, para el caso
# excepcional de una licencia institucional propia.
# Cada entrada: (clave interna, etiqueta visible, href, texto legal en español)
LICENCIAS = [
    ("cc-by-4.0", "CC BY 4.0 — Atribución",
     "https://creativecommons.org/licenses/by/4.0/",
     "Distribuido bajo una licencia Creative Commons Attribution 4.0 International (CC BY 4.0)."),
    ("cc-by-sa-4.0", "CC BY-SA 4.0 — Atribución compartir igual",
     "https://creativecommons.org/licenses/by-sa/4.0/",
     "Distribuido bajo una licencia Creative Commons Attribution-ShareAlike 4.0 "
     "International (CC BY-SA 4.0)."),
    ("cc-by-nc-4.0", "CC BY-NC 4.0 — Atribución no comercial",
     "https://creativecommons.org/licenses/by-nc/4.0/",
     "Distribuido bajo una licencia Creative Commons Attribution-NonCommercial 4.0 "
     "International (CC BY-NC 4.0)."),
    ("cc-by-nc-sa-4.0", "CC BY-NC-SA 4.0 — Atribución no comercial compartir igual",
     "https://creativecommons.org/licenses/by-nc-sa/4.0/",
     "Distribuido bajo una licencia Creative Commons Attribution-NonCommercial-ShareAlike "
     "4.0 International (CC BY-NC-SA 4.0)."),
    ("cc-by-nd-4.0", "CC BY-ND 4.0 — Atribución sin derivadas",
     "https://creativecommons.org/licenses/by-nd/4.0/",
     "Distribuido bajo una licencia Creative Commons Attribution-NoDerivatives 4.0 "
     "International (CC BY-ND 4.0)."),
    ("cc-by-nc-nd-4.0", "CC BY-NC-ND 4.0 — Atribución no comercial sin derivadas",
     "https://creativecommons.org/licenses/by-nc-nd/4.0/",
     "Distribuido bajo una licencia Creative Commons Atribución-NoComercial-"
     "SinDerivadas 4.0 Internacional (CC BY-NC-ND 4.0)."),
    ("todos-los-derechos", "Todos los derechos reservados",
     "",
     "Todos los derechos reservados."),
]
LICENCIA_POR_CLAVE = {l[0]: l for l in LICENCIAS}
LICENCIA_DEFAULT = "cc-by-nc-nd-4.0"

# Clave especial que la UI usa para mostrar el campo de texto libre.
LICENCIA_CLAVE_OTRA = "otra"

# ─── Estilos de textbox por clase (UI) ────────────────────────────────────────
# Cada entrada: (font_size, weight, slant, fg_color, bg_color)
# Usados por _aplicar_estilo_textbox() y _crear_bloque_ui() en la UI.

ESTILO_POR_CLASE: dict[str, tuple] = {
    "Título principal":    (17, "bold",   "roman",  "#ffffff", "#1a237e"),
    "Título secundario":   (16, "bold",   "italic", "#e8eaf6", "#283593"),
    "Encabezado sección":  (14, "bold",   "roman",  "#e3f2fd", "#0277bd"),
    "Subencabezado":       (13, "bold",   "roman",  "#e0f2f1", "#00695c"),
    "Subencabezado-bajo":  (12, "normal", "italic", "#e0f2f1", "#00796b"),
    "Cuerpo":              (12, "normal", "roman",  "#e2e8f0", "#1a1a2e"),
    "Cuerpo del abstract": (12, "normal", "roman",  "#b0bec5", "#263238"),
    "Palabras clave":      (11, "normal", "roman",  "#f3e5f5", "#4a148c"),
    "Referencia":          (11, "normal", "roman",  "#cccccc", "#2a2a2a"),
    "Cómo citar":          (11, "normal", "italic", "#fff3e0", "#bf360c"),
    "Fecha manuscrito":    (11, "normal", "italic", "#fbe9e7", "#8d2b07"),
    "Título tabla":        (12, "bold",   "roman",  "#e3f2fd", "#0d47a1"),
    "Pie de figura":       (11, "normal", "italic", "#f1f8e9", "#33691e"),
    "Filiación":           (11, "normal", "roman",  "#e8f5e9", "#1b5e20"),
    "Email / Metadatos":   (11, "normal", "italic", "#e8f5e9", "#1b5e20"),
    "Ignorar":             (11, "normal", "roman",  "#ffcdd2", "#7f0000"),
}

# ─── CSS editorial para exportación HTML ──────────────────────────────────────
# También se usa en el EPUB (sin el @import de Google Fonts).

HTML_CSS = """
<style>
  @import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:ital,wght@0,400;0,700;1,400;1,700&display=swap');
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  :root {
    --serif: 'Source Serif 4', 'Georgia', serif;
    --sans:  'Helvetica Neue', Arial, sans-serif;
    --tnr:   'Times New Roman', Times, serif;
    --texto: #1a1a1a; --gris: #555; --linea: #bbb;
  }
  body {
    font-family: var(--serif); font-size: 12pt; line-height: 1.65;
    color: var(--texto); background: #fff;
    max-width: 880px; margin: 0 auto; padding: 32px 52px 64px; hyphens: auto;
  }
  h1.titulo-principal  { font-family:var(--serif); font-size:14pt; font-weight:700; font-style:normal;  line-height:1.3; margin:20px 0 6px; }
  h2.titulo-secundario { font-family:var(--serif); font-size:14pt; font-weight:700; font-style:italic;  line-height:1.3; color:#222; margin-bottom:14px; }
  .autores     { font-family:var(--serif); font-size:13pt; font-weight:400; font-style:normal; margin:12px 0 4px; line-height:1.6; }
  .autores a.orcid-autor { color:#1a3a5c; text-decoration:underline; text-underline-offset:3px; text-decoration-color:#A6CE39; text-decoration-thickness:2px; }
  .autores a.orcid-autor:hover { text-decoration-color:#1a3a5c; }
  .autores .orcid-icon img { width:16px; height:16px; vertical-align:middle; margin-left:2px; }
  .filiaciones { font-family:var(--serif); font-size:9pt; font-weight:400; font-style:normal; color:#666; margin:2px 0; line-height:1.5; }
  .email       { font-family:var(--serif); font-size:9pt;  font-style:italic; color:#1a5276; margin-bottom:6px; }
  h2.seccion { font-family:var(--serif); font-size:13pt; font-weight:700; text-align:center; margin:24px 0 10px; }
  h2.seccion.meta { font-size:9pt; font-weight:400; margin:3px 0; }
  h2.seccion.con-linea { border-top:1px solid var(--linea); padding-top:18px; margin-top:32px; }
  h2.seccion.gris { color:#666; }
  h3.subseccion { font-family:var(--serif); font-size:12pt; font-weight:700; margin:20px 0 8px; }
  h3.subseccion.primer-nivel1 { border-top:1px solid var(--linea); padding-top:18px; margin-top:32px; }
  h3.subseccion-bajo{ font-family:var(--serif); font-size:12pt; font-weight:700; font-style:italic; margin:16px 0 6px; }
  p.resumen  { font-family:var(--serif); font-size:9pt; text-align:justify; text-indent:1.2em; margin-bottom:7px; }
  p.abstract { font-family:var(--tnr);   font-size:12pt; color:#666; font-style:italic; text-align:justify; text-indent:1.2em; margin-bottom:7px; }
  p { font-family:var(--serif); font-size:10pt; text-align:justify; text-indent:1.4em; margin-bottom:8px; font-style:normal; }
  p.sin-sangria { text-indent:0; }
  p.cuerpo { font-family:var(--serif); font-size:12pt; font-style:normal !important; font-weight:normal; text-align:justify; text-indent:1.4em; margin-bottom:9px; color:var(--texto); }
  .keywords { font-family:var(--serif); font-size:9pt; margin:4px 0 16px; text-indent:0; }
  .keywords strong { font-weight:700; }
  ol.referencias { padding-left:2em; margin:8px 0 16px; }
  ol.referencias li { font-family:var(--serif); font-size:10pt; margin-bottom:6px; line-height:1.5; }
  .post-referencias { margin-top:28px; border-top:1px solid var(--linea); padding-top:14px; }
  .como-citar { font-family:var(--serif); font-size:9pt; margin-bottom:10px; line-height:1.5; }
  .fechas-manuscrito ul { list-style:disc; padding-left:1.6em; margin:6px 0 10px; }
  .fechas-manuscrito li { font-family:var(--serif); font-size:9pt; margin-bottom:4px; }
  .fechas-manuscrito a { color:#1a5276; }
  .figuras-finales { margin-top:28px; border-top:1px solid var(--linea); padding-top:14px; }
  .figuras-finales h2 { font-family:var(--serif); font-size:10pt; font-weight:700; text-align:center; margin-bottom:14px; text-transform:uppercase; letter-spacing:0.05em; }
  figure { margin:18px auto; text-align:center; }
  figure img { max-width:100%; border:1px solid var(--linea); }
  figcaption { font-family:var(--serif); font-size:9pt; color:#1a1a1a; margin-top:5px; text-align:left; line-height:1.4; }
  .tabla-wrapper { margin:20px auto 24px; overflow-x:auto; }
  .tabla-titulo { font-family:var(--serif); font-size:9pt; font-weight:400; color:#666; margin-bottom:6px; }
  .tabla-titulo strong { font-weight:700; }
  table.pm-tabla { border-collapse:collapse; width:100%; font-family:var(--serif); font-size:9pt; }
  table.pm-tabla thead tr th { background:#1b5e9a; color:#fff; font-weight:700; padding:6px 10px; border:1px solid #155080; text-align:center; }
  table.pm-tabla tbody tr td { background:#cbeefb; color:#1a1a1a; padding:4px 10px; border:1px solid #9fd8f0; vertical-align:top; }
  table.pm-tabla tbody tr:nth-child(even) td { background:#b8e6f8; }
</style>
"""

# RF-42: Licencia por defecto de la revista 

LICENCIA_TEXTO_DEFAULT = (
    "Distribuido bajo una licencia Creative Commons Atribución-NoComercial-"
    "SinDerivadas 4.0 Internacional (CC BY-NC-ND 4.0)."
)
LICENCIA_HREF_DEFAULT = "https://creativecommons.org/licenses/by-nc-nd/4.0/"