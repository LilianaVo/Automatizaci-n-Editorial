import customtkinter as ctk
from tkinter import filedialog
from PIL import Image as PILImage
import fitz
import re
import os
import ebooklib
from ebooklib import epub
from collections import Counter
import base64
import shutil

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# ─── CSS exportado ────────────────────────────────────────────────────────────
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
  h1.titulo-principal { font-family:var(--serif); font-size:14pt; font-weight:700; line-height:1.3; margin:20px 0 8px; }
  h2.titulo-secundario { font-family:var(--serif); font-size:14pt; font-weight:700; font-style:italic; color:#333; margin-bottom:16px; line-height:1.3; }
  .autores { font-family:var(--serif); font-size:13pt; margin:12px 0 4px; line-height:1.6; }
  .autores a.orcid-autor { color:#1a3a5c; text-decoration:underline; text-underline-offset:3px; text-decoration-color:#A6CE39; text-decoration-thickness:2px; }
  .autores a.orcid-autor:hover { text-decoration-color:#1a3a5c; }
  .autores .orcid-icon img { width:16px; height:16px; vertical-align:middle; margin-left:2px; }
  .filiaciones { font-family:var(--serif); font-size:9pt; color:var(--gris); margin:3px 0; line-height:1.5; }
  .email { font-family:var(--serif); font-size:12pt; font-style:italic; color:#1a5276; margin-bottom:6px; }
  h2.seccion { font-family:var(--serif); font-size:10pt; font-weight:700; text-align:center; margin:20px 0 8px; }
  /* Subtítulos nivel 1 (1. 2. 3.) — negritas */
  h3.subseccion { font-family:var(--serif); font-size:10pt; font-weight:700; margin:18px 0 6px; }
  /* Subtítulos nivel 2+ (1.1 1.2 2.3) — sin negritas */
  h3.subseccion-bajo { font-family:var(--serif); font-size:10pt; font-weight:400; font-style:italic; margin:14px 0 5px; }
  p.resumen { font-family:var(--serif); font-size:9pt; text-align:justify; text-indent:1.2em; margin-bottom:7px; }
  p.abstract { font-family:var(--tnr); font-size:9pt; font-style:italic; text-align:justify; text-indent:1.2em; margin-bottom:7px; }
  p { font-family:var(--serif); font-size:10pt; text-align:justify; text-indent:1.4em; margin-bottom:8px; font-style:normal; }
  p.sin-sangria { text-indent:0; }
  /* NO aplicar cursiva automática a párrafos de cuerpo */
  /* Cuerpo del artículo: Source Serif Pro 12pt (10+2), desde sección 1 hasta Referencias */
  p.cuerpo { font-family:var(--serif); font-size:12pt; text-align:justify; text-indent:1.4em; margin-bottom:9px; font-style:normal; color:var(--texto); }
  .keywords { font-family:var(--serif); font-size:9pt; margin:4px 0 16px; text-indent:0; }
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
  figcaption { font-family:var(--serif); font-size:9pt; color:var(--gris); margin-top:5px; text-align:left; line-height:1.4; }
  /* ── Tablas ── */
  .tabla-wrapper { margin:20px auto 24px; overflow-x:auto; }
  .tabla-titulo { font-family:var(--serif); font-size:9pt; font-weight:700; margin-bottom:6px; }
  table.pm-tabla { border-collapse:collapse; width:100%; font-family:var(--serif); font-size:9pt; }
  table.pm-tabla thead tr th { background:#1b5e9a; color:#fff; font-weight:700; padding:6px 10px; border:1px solid #155080; text-align:center; }
  table.pm-tabla tbody tr td { background:#cbeefb; color:#1a1a1a; padding:4px 10px; border:1px solid #9fd8f0; vertical-align:top; }
  table.pm-tabla tbody tr:nth-child(even) td { background:#b8e6f8; }
</style>
"""

# ─── Clasificaciones ──────────────────────────────────────────────────────────
OPCIONES = [
    "Normal", "Cuerpo", "Título principal", "Título secundario", "Autores",
    "Filiación", "Email / Metadatos", "Cómo citar", "Fecha manuscrito",
    "Encabezado sección", "Subencabezado", "Subencabezado-bajo",
    "Resumen / Abstract", "Palabras clave", "Referencia",
    "Título tabla", "Pie de figura", "Imagen", "Ignorar",
]

COLORES_UI = [
    ("Título principal",   "#1a3a5c", "Azul oscuro"),
    ("Título secundario",  "#1e5f74", "Azul medio"),
    ("Autores",            "#145a32", "Verde oscuro"),
    ("Filiación",          "#1d4e1d", "Verde musgo"),
    ("Email / Metadatos",  "#1d3d1d", "Verde bajo"),
    ("Cómo citar",         "#4a3000", "Café dorado"),
    ("Fecha manuscrito",   "#3d2800", "Café oscuro"),
    ("Encabezado sección", "#1f538d", "Azul índigo"),
    ("Subencabezado",      "#2471a3", "Azul cielo — bold"),
    ("Subencabezado-bajo", "#1a6e9e", "Azul cielo — sin bold"),
    ("Resumen / Abstract", "#5a3a1d", "Siena"),
    ("Palabras clave",     "#4a235a", "Morado"),
    ("Referencia",         "#3a3030", "Gris oscuro"),
    ("Título tabla",       "#1b3a5c", "Azul tabla"),
    ("Cuerpo",             "#1a2a1a", "Verde muy oscuro — cuerpo art."),
    ("Pie de figura",      "#3a3a1b", "Oliva"),
    ("Imagen",             "#2b3a2b", "Verde gris"),
    ("Ignorar",            "#3a2b2b", "Rojo oscuro"),
]
COLOR_POR_CLASE = {c: col for c, col, _ in COLORES_UI}

# ─── ORCID SVG embebido ───────────────────────────────────────────────────────
_ORCID_SVG = (
    "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmci"
    "IHZpZXdCb3g9IjAgMCAyNCAyNCI+PGNpcmNsZSBjeD0iMTIiIGN5PSIxMiIgcj0iMTIiIGZpbGw9"
    "IiNBNkNFMzkiLz48cGF0aCBkPSJNNy41IDVoMXY3LjVoLTF6TTkuMyA3LjhDOS4zIDYgMTAuNCA1"
    "IDEyLjEgNWMxLjggMCAyLjkgMSAyLjkgMi44VjEyaDEuMlY3LjhjMC0yLjQtMS40LTMuOC0zLjct"
    "My44QzkuOCA0IDguMSA1IDguMSA3LjhWMTJoMS4yVjcuOHoiIGZpbGw9IiNmZmYiLz48L3N2Zz4="
)


# ─── Utilidades ───────────────────────────────────────────────────────────────

def esc(t: str) -> str:
    return t.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")


def _insertar_orcid(texto: str) -> str:
    """
    Recibe el texto crudo de autores, separa por ';', limpia superíndices,
    envuelve cada nombre en un <a> subrayado que apunta a búsqueda ORCID.
    """
    partes = [p.strip() for p in texto.split(";") if p.strip()]
    result = []
    for parte in partes:
        nombre_raw = re.sub(r"[\d,\*\u00b9\u00b2\u00b3\u2070-\u209f]+$", "", parte).strip()
        if not nombre_raw:
            continue
        nombre_esc = esc(nombre_raw)
        q   = re.sub(r"\s+", "+", nombre_raw)
        url = f"https://orcid.org/orcid-search/search?searchQuery={q}"
        ico = (f'<a class="orcid-icon" href="{url}" target="_blank" title="Buscar en ORCID">'
               f'<img src="{_ORCID_SVG}" alt="ORCID"></a>')
        result.append(
            f'<a class="orcid-autor" href="{url}" target="_blank" '
            f'title="Buscar {nombre_esc} en ORCID">{nombre_esc}</a>{ico}')
    return "; ".join(result)


def _parsear_referencias(texto: str) -> list[str]:
    patron = re.compile(r'^\s*(?:\[?\d+[\.\)\]]\s*)', re.MULTILINE)
    partes = patron.split(texto)
    refs   = [p.strip() for p in partes if p.strip()]
    if refs:
        return refs
    return [l.strip() for l in texto.splitlines() if l.strip()]


def _es_como_citar(t: str) -> bool:
    return bool(re.match(r"(cómo citar|how to cite)", t.strip().lower()))


def _es_fecha_mss(t: str) -> bool:
    return bool(re.search(
        r"(manuscrito\s+recibido|manuscrito\s+corregido|manuscrito\s+aceptado"
        r"|manuscript\s+received|manuscript\s+revised|manuscript\s+accepted)",
        t, re.IGNORECASE))


def _es_doi(t: str) -> bool:
    return bool(re.search(r"https?://doi\.org/", t))


def _img_to_base64(path: str) -> str:
    """Devuelve data-URI base64 de la imagen."""
    ext  = os.path.splitext(path)[1].lower().lstrip(".")
    mime = {"jpg": "jpeg", "jpeg": "jpeg", "png": "png",
            "gif": "gif",  "webp": "webp"}.get(ext, "jpeg")
    with open(path, "rb") as f:
        data = base64.b64encode(f.read()).decode()
    return f"data:image/{mime};base64,{data}"


def _excel_a_html_tabla(ruta: str) -> str:
    """
    Convierte la primera hoja de un .xlsx en una tabla HTML.
    Requiere openpyxl (pip install openpyxl).
    """
    try:
        import openpyxl
        wb = openpyxl.load_workbook(ruta, data_only=True)
        ws = wb.active
        filas = list(ws.iter_rows(values_only=True))
        if not filas:
            return "<p><em>[Tabla vacía]</em></p>"

        html = ['<table class="pm-tabla">']
        # Primera fila como encabezado
        html.append("<thead><tr>")
        for celda in filas[0]:
            val = "" if celda is None else str(celda)
            html.append(f"<th>{esc(val)}</th>")
        html.append("</tr></thead><tbody>")
        for fila in filas[1:]:
            if all(c is None for c in fila):
                continue   # saltar filas vacías
            html.append("<tr>")
            for celda in fila:
                val = "" if celda is None else str(celda)
                html.append(f"<td>{esc(val)}</td>")
            html.append("</tr>")
        html.append("</tbody></table>")
        return "\n".join(html)
    except ImportError:
        return "<p><em>[Instala openpyxl para renderizar tablas: pip install openpyxl]</em></p>"
    except Exception as e:
        return f"<p><em>[Error al leer tabla: {esc(str(e))}]</em></p>"


# ─────────────────────────────────────────────────────────────────────────────
class LimpiadorEditorialApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Limpiador Editorial Pro - UNAM")
        self.geometry("1180x920")
        self.minsize(980, 660)

        self.datos_bloques:        list[dict] = []
        self.referencias_externas: list[str]  = []
        self.figuras_manuales:     list[dict] = []
        self.tablas_manuales:      list[dict] = []   # {"ruta", "titulo", "_var"}

        # ── Encabezado ──────────────────────────────────────────
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", padx=20, pady=(14, 0))
        ctk.CTkLabel(top, text="Extractor de PDF Semántico",
                     font=ctk.CTkFont(size=22, weight="bold")).pack(side="left")

        # ── Botones de exportar (siempre visibles) ───────────────
        exp_f = ctk.CTkFrame(self, fg_color="transparent")
        exp_f.pack(pady=(6, 2))
        for texto, cmd, fg, hv in [
            ("🌐  HTML",  self.evento_exportar_html,  "#28a745", "#1e7a35"),
            ("📄  XML",   self.evento_exportar_xml,   "#e67e22", "#b35a00"),
            ("📚  EPUB",  self.evento_exportar_epub,  "#8e44ad", "#6c1f8a"),
        ]:
            ctk.CTkButton(exp_f, text=texto, command=cmd,
                          fg_color=fg, hover_color=hv,
                          width=140, font=ctk.CTkFont(size=12)
                          ).pack(side="left", padx=5)

        # ── Pestañas ─────────────────────────────────────────────
        self.tabs = ctk.CTkTabview(self)
        self.tabs.pack(fill="both", expand=True, padx=14, pady=(4, 2))

        self.tabs.add("📄  PDF")
        self.tabs.add("📋  Referencias")
        self.tabs.add("🖼️  Figuras")

        self._construir_tab_pdf()
        self._construir_tab_referencias()
        self._construir_tab_figuras()

        # ── Status bar ───────────────────────────────────────────
        self._status = ctk.CTkLabel(
            self, text="Listo.", anchor="w",
            font=ctk.CTkFont(size=10), text_color="#888")
        self._status.pack(fill="x", padx=16, pady=(0, 6))

    # ═════════════════════════════════════════════════════════════
    # TAB 1 — PDF
    # ═════════════════════════════════════════════════════════════

    def _construir_tab_pdf(self):
        tab = self.tabs.tab("📄  PDF")

        # ── Fila 1: botones de acción ─────────────────────────────
        btn_f = ctk.CTkFrame(tab, fg_color="transparent")
        btn_f.pack(fill="x", pady=(6, 2))

        self._btn_cargar = ctk.CTkButton(
            btn_f, text="📂  Cargar PDF",
            command=self.evento_cargar_archivo,
            fg_color="#2e86c1", hover_color="#1a5276",
            width=160, font=ctk.CTkFont(size=12))
        self._btn_cargar.pack(side="left", padx=5)

        ctk.CTkButton(btn_f, text="🎨  Leyenda",
                      command=self._toggle_leyenda,
                      fg_color="#3a3a3a", hover_color="#555",
                      width=100, height=28, font=ctk.CTkFont(size=11)
                      ).pack(side="left", padx=5)

        self._stats_lbl = ctk.CTkLabel(
            btn_f, text="", font=ctk.CTkFont(size=10), text_color="#999")
        self._stats_lbl.pack(side="left", padx=10)

        # ── Fila 2: zona segmentada + filtro fino ─────────────────
        nav_f = ctk.CTkFrame(tab, fg_color="transparent")
        nav_f.pack(fill="x", pady=(0, 3))

        # Zonas principales — segmented button
        self._zona_seg = ctk.CTkSegmentedButton(
            nav_f,
            values=["📋 Portada", "📖 Cuerpo", "🖼 Medios"],
            command=self._cambiar_zona,
            font=ctk.CTkFont(size=12),
            width=320,
        )
        self._zona_seg.set("📋 Portada")
        self._zona_seg.pack(side="left", padx=(4, 12))

        ctk.CTkLabel(nav_f, text="Filtro:",
                     font=ctk.CTkFont(size=11)).pack(side="left", padx=(0, 2))
        self._filtro_menu = ctk.CTkOptionMenu(
            nav_f, values=["Todos"] + OPCIONES,
            command=self._aplicar_filtro, width=170,
            font=ctk.CTkFont(size=11))
        self._filtro_menu.set("Todos")
        self._filtro_menu.pack(side="left", padx=4)

        # ── Leyenda (oculta) ──────────────────────────────────────
        self._leyenda_visible = False
        self._leyenda_panel   = ctk.CTkFrame(tab, fg_color="#1e1e1e", corner_radius=8)
        self._construir_leyenda(tab)

        # ── Banner de completado ──────────────────────────────────
        self._banner = ctk.CTkLabel(
            tab, text="",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color="#1b5e20", fg_color="#c8e6c9",
            corner_radius=6, height=34, anchor="center")

        # ── Scroll de bloques ─────────────────────────────────────
        self.frame_scroll = ctk.CTkScrollableFrame(
            tab, label_text="Bloques extraídos")
        self.frame_scroll.pack(fill="both", expand=True, padx=0, pady=(0, 2))

    # ═════════════════════════════════════════════════════════════
    # TAB 2 — REFERENCIAS
    # ═════════════════════════════════════════════════════════════

    def _construir_tab_referencias(self):
        tab = self.tabs.tab("📋  Referencias")

        info = ctk.CTkLabel(
            tab,
            text=(
                "Carga un .txt con las referencias numeradas.\n"
                "Formatos aceptados:   1. Texto...   |   1) Texto...   |   [1] Texto...\n"
                "Si cargas referencias aquí, los bloques 'Referencia' del PDF se ignorarán automáticamente."
            ),
            font=ctk.CTkFont(size=11), justify="left", text_color="#aaa")
        info.pack(anchor="w", padx=12, pady=(12, 6))

        btn_f = ctk.CTkFrame(tab, fg_color="transparent")
        btn_f.pack(fill="x", padx=10, pady=(0, 8))

        ctk.CTkButton(btn_f, text="📂  Cargar .txt de referencias",
                      command=self.evento_cargar_referencias,
                      fg_color="#5d4037", hover_color="#3e2723",
                      width=230, font=ctk.CTkFont(size=12)
                      ).pack(side="left", padx=5)

        ctk.CTkButton(btn_f, text="🗑  Limpiar",
                      command=self._limpiar_referencias,
                      fg_color="#6d4c41", hover_color="#4e342e",
                      width=100, font=ctk.CTkFont(size=11)
                      ).pack(side="left", padx=5)

        self._refs_count_lbl = ctk.CTkLabel(
            btn_f, text="Sin referencias cargadas",
            font=ctk.CTkFont(size=11), text_color="#888")
        self._refs_count_lbl.pack(side="left", padx=10)

        # Lista previa
        self._refs_scroll = ctk.CTkScrollableFrame(
            tab, label_text="Referencias cargadas")
        self._refs_scroll.pack(fill="both", expand=True, padx=10, pady=(0, 6))

    # ═════════════════════════════════════════════════════════════
    # TAB 3 — FIGURAS + TABLAS
    # ═════════════════════════════════════════════════════════════

    def _construir_tab_figuras(self):
        tab = self.tabs.tab("🖼️  Figuras")

        # ── Segmented: Figuras / Tablas ───────────────────────────
        seg = ctk.CTkSegmentedButton(
            tab,
            values=["🖼️  Figuras", "📊  Tablas"],
            command=self._cambiar_panel_media,
            font=ctk.CTkFont(size=12), width=280)
        seg.set("🖼️  Figuras")
        seg.pack(pady=(10, 6))

        # ── Panel Figuras ─────────────────────────────────────────
        self._panel_figs = ctk.CTkFrame(tab, fg_color="transparent")
        self._panel_figs.pack(fill="both", expand=True)

        info_f = ctk.CTkLabel(
            self._panel_figs,
            text="Agrega imágenes con pie de figura. Numeración automática (Figura 1, 2…).\n"
                 "Se insertan al final del HTML como en InDesign.",
            font=ctk.CTkFont(size=11), justify="left", text_color="#aaa")
        info_f.pack(anchor="w", padx=12, pady=(4, 6))

        bf = ctk.CTkFrame(self._panel_figs, fg_color="transparent")
        bf.pack(fill="x", padx=10, pady=(0, 6))
        ctk.CTkButton(bf, text="➕  Agregar figura",
                      command=self._agregar_figura,
                      fg_color="#1565c0", hover_color="#0d47a1",
                      width=160, font=ctk.CTkFont(size=12)).pack(side="left", padx=5)
        ctk.CTkButton(bf, text="🗑  Limpiar",
                      command=self._limpiar_figuras,
                      fg_color="#6d4c41", hover_color="#4e342e",
                      width=100, font=ctk.CTkFont(size=11)).pack(side="left", padx=5)
        self._figs_count_lbl = ctk.CTkLabel(bf, text="Sin figuras",
                                             font=ctk.CTkFont(size=11), text_color="#888")
        self._figs_count_lbl.pack(side="left", padx=10)

        self._figs_scroll = ctk.CTkScrollableFrame(
            self._panel_figs, label_text="Figuras cargadas")
        self._figs_scroll.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        # ── Panel Tablas ──────────────────────────────────────────
        self._panel_tabs = ctk.CTkFrame(tab, fg_color="transparent")
        # (oculto hasta que el usuario pulse "Tablas")

        info_t = ctk.CTkLabel(
            self._panel_tabs,
            text="Importa archivos Excel (.xlsx). Numeración automática (Tabla 1, 2…).\n"
                 "Escribe el título de la tabla. Para insertar en el cuerpo usa el marcador [Tabla N].",
            font=ctk.CTkFont(size=11), justify="left", text_color="#aaa")
        info_t.pack(anchor="w", padx=12, pady=(4, 6))

        bt = ctk.CTkFrame(self._panel_tabs, fg_color="transparent")
        bt.pack(fill="x", padx=10, pady=(0, 6))
        ctk.CTkButton(bt, text="➕  Agregar tabla (.xlsx)",
                      command=self._agregar_tabla,
                      fg_color="#4a148c", hover_color="#311b92",
                      width=190, font=ctk.CTkFont(size=12)).pack(side="left", padx=5)
        ctk.CTkButton(bt, text="🗑  Limpiar",
                      command=self._limpiar_tablas,
                      fg_color="#6d4c41", hover_color="#4e342e",
                      width=100, font=ctk.CTkFont(size=11)).pack(side="left", padx=5)
        self._tabs_count_lbl = ctk.CTkLabel(bt, text="Sin tablas",
                                             font=ctk.CTkFont(size=11), text_color="#888")
        self._tabs_count_lbl.pack(side="left", padx=10)

        self._tabs_scroll = ctk.CTkScrollableFrame(
            self._panel_tabs, label_text="Tablas cargadas")
        self._tabs_scroll.pack(fill="both", expand=True, padx=10, pady=(0, 6))

    def _cambiar_panel_media(self, valor: str):
        if valor == "🖼️  Figuras":
            self._panel_tabs.pack_forget()
            self._panel_figs.pack(fill="both", expand=True)
        else:
            self._panel_figs.pack_forget()
            self._panel_tabs.pack(fill="both", expand=True)

    # ── Helpers generales ────────────────────────────────────────

    def _set_status(self, msg: str):
        self._status.configure(text=msg)
        self.update_idletasks()

    # Clases que pertenecen a cada zona del segmented button
    _ZONA_CLASES = {
        "📋 Portada": {
            "Título principal", "Título secundario", "Autores", "Filiación",
            "Email / Metadatos", "Cómo citar", "Fecha manuscrito",
            "Encabezado sección", "Resumen / Abstract", "Palabras clave",
            "Normal",
        },
        "📖 Cuerpo": {
            "Subencabezado", "Subencabezado-bajo", "Cuerpo",
            "Encabezado sección", "Referencia", "Título tabla",
        },
        "🖼 Medios": {
            "Imagen", "Pie de figura", "Ignorar",
        },
    }

    def _cambiar_zona(self, zona: str):
        """Muestra solo los bloques de la zona seleccionada."""
        clases_visibles = self._ZONA_CLASES.get(zona, None)
        filtro_actual   = self._filtro_menu.get()
        for b in self.datos_bloques:
            cls = b["menu"].get()
            visible = True
            if clases_visibles is not None:
                visible = cls in clases_visibles
            if filtro_actual != "Todos" and cls != filtro_actual:
                visible = False
            if visible:
                b["frame"].pack(fill="x", padx=8, pady=2)
            else:
                b["frame"].pack_forget()

    def _mostrar_banner(self, msg: str, color_bg="#c8e6c9", color_txt="#1b5e20"):
        """Muestra el banner verde de completado y lo oculta tras 4 segundos."""
        self._banner.configure(text=msg, fg_color=color_bg, text_color=color_txt)
        # Ocultar el scroll temporalmente, insertar banner, volver a mostrar scroll
        self.frame_scroll.pack_forget()
        self._banner.pack(fill="x", padx=8, pady=(0, 4))
        self.frame_scroll.pack(fill="both", expand=True, padx=0, pady=(0, 2))
        self.after(4000, self._ocultar_banner)

    def _ocultar_banner(self):
        self._banner.pack_forget()

    def _actualizar_stats(self, *_):
        if not self.datos_bloques:
            self._stats_lbl.configure(text=""); return
        c = Counter(b["menu"].get() for b in self.datos_bloques)
        self._stats_lbl.configure(
            text="  ".join(f"{k[:5]}:{v}" for k, v in c.most_common(6)))

    def _aplicar_filtro(self, valor: str):
        zona = self._zona_seg.get() if hasattr(self, "_zona_seg") else None
        clases_zona = self._ZONA_CLASES.get(zona) if zona else None
        for b in self.datos_bloques:
            cls = b["menu"].get()
            en_zona   = (clases_zona is None or cls in clases_zona)
            en_filtro = (valor == "Todos" or cls == valor)
            if en_zona and en_filtro:
                b["frame"].pack(fill="x", padx=8, pady=2)
            else:
                b["frame"].pack_forget()

    # ── Leyenda ───────────────────────────────────────────────────

    def _construir_leyenda(self, parent):
        cols = 4
        for i, (cls, color, etiqueta) in enumerate(COLORES_UI):
            row, col = divmod(i, cols)
            celda = ctk.CTkFrame(self._leyenda_panel, fg_color="transparent")
            celda.grid(row=row, column=col, padx=10, pady=4, sticky="w")
            sw = ctk.CTkFrame(celda, width=18, height=18,
                               fg_color=color, corner_radius=3)
            sw.pack(side="left", padx=(0, 5)); sw.pack_propagate(False)
            ctk.CTkLabel(celda, text=f"{cls}  ({etiqueta})",
                         font=ctk.CTkFont(size=10),
                         text_color="#ddd").pack(side="left")

    def _toggle_leyenda(self):
        self._leyenda_visible = not self._leyenda_visible
        if self._leyenda_visible:
            self._leyenda_panel.pack(fill="x", padx=0, pady=(0, 6))
        else:
            self._leyenda_panel.pack_forget()

    # ═════════════════════════════════════════════════════════════
    # REFERENCIAS externas
    # ═════════════════════════════════════════════════════════════

    def evento_cargar_referencias(self):
        ruta = filedialog.askopenfilename(
            title="Referencias (.txt)",
            filetypes=[("Texto", "*.txt"), ("Todos", "*.*")])
        if not ruta: return
        try:
            with open(ruta, encoding="utf-8", errors="replace") as f:
                contenido = f.read()
            self.referencias_externas = _parsear_referencias(contenido)
            self._refrescar_lista_refs()
            n = len(self.referencias_externas)
            self._refs_count_lbl.configure(
                text=f"✓ {n} referencias  ({os.path.basename(ruta)})",
                text_color="#a5d6a7")
            self._set_status(f"✓ {n} referencias cargadas desde '{os.path.basename(ruta)}'")
        except Exception as e:
            self._set_status(f"❌ Error: {e}")

    def _limpiar_referencias(self):
        self.referencias_externas = []
        self._refrescar_lista_refs()
        self._refs_count_lbl.configure(
            text="Sin referencias cargadas", text_color="#888")
        self._set_status("Referencias limpiadas.")

    def _refrescar_lista_refs(self):
        for w in self._refs_scroll.winfo_children():
            w.destroy()
        for i, ref in enumerate(self.referencias_externas, 1):
            frame = ctk.CTkFrame(self._refs_scroll, fg_color="#252525", corner_radius=4)
            frame.pack(fill="x", padx=4, pady=2)
            ctk.CTkLabel(frame, text=f"{i}.", width=32,
                         font=ctk.CTkFont(size=10, weight="bold"),
                         text_color="#aaa").pack(side="left", padx=(6, 2), pady=4)
            ctk.CTkLabel(frame, text=ref[:160] + ("…" if len(ref) > 160 else ""),
                         wraplength=780, justify="left",
                         font=ctk.CTkFont(size=10)).pack(
                             side="left", padx=4, pady=4, fill="x", expand=True)

    # ═════════════════════════════════════════════════════════════
    # FIGURAS manuales
    # ═════════════════════════════════════════════════════════════

    def _agregar_figura(self):
        rutas = filedialog.askopenfilenames(
            title="Selecciona imagen(es)",
            filetypes=[
                ("Imágenes", "*.jpg *.jpeg *.png *.gif *.webp *.tif *.tiff"),
                ("Todos", "*.*"),
            ])
        if not rutas: return
        for ruta in rutas:
            self.figuras_manuales.append({"ruta": ruta, "pie": ""})
        self._refrescar_lista_figuras()
        self._set_status(f"✓ {len(self.figuras_manuales)} figura(s) en total.")

    def _limpiar_figuras(self):
        self.figuras_manuales = []
        self._refrescar_lista_figuras()
        self._figs_count_lbl.configure(text="Sin figuras", text_color="#888")
        self._set_status("Figuras limpiadas.")

    def _refrescar_lista_figuras(self):
        for w in self._figs_scroll.winfo_children():
            w.destroy()

        for i, fig in enumerate(self.figuras_manuales):
            num = i + 1
            frame = ctk.CTkFrame(self._figs_scroll,
                                  fg_color="#1e2a1e", corner_radius=6)
            frame.pack(fill="x", padx=4, pady=4)
            frame.columnconfigure(2, weight=1)

            # Miniatura
            try:
                img_pil = PILImage.open(fig["ruta"])
                img_pil.thumbnail((80, 80))
                thumb   = ctk.CTkImage(img_pil, size=img_pil.size)
                ctk.CTkLabel(frame, image=thumb, text="").grid(
                    row=0, column=0, rowspan=2, padx=(8, 6), pady=6)
            except Exception:
                ctk.CTkLabel(frame, text="🖼️", font=ctk.CTkFont(size=30),
                             width=80).grid(row=0, column=0, rowspan=2,
                                            padx=(8, 6), pady=6)

            # Número de figura (fijo)
            ctk.CTkLabel(frame,
                         text=f"Figura {num}",
                         font=ctk.CTkFont(size=11, weight="bold"),
                         text_color="#a5d6a7").grid(
                row=0, column=1, padx=(0, 8), pady=(6, 0), sticky="w")

            # Nombre de archivo
            ctk.CTkLabel(frame,
                         text=os.path.basename(fig["ruta"]),
                         font=ctk.CTkFont(size=9), text_color="#888").grid(
                row=1, column=1, padx=(0, 8), pady=(0, 6), sticky="w")

            # Campo de texto para el pie de figura
            pie_var = ctk.StringVar(value=fig["pie"])
            entry = ctk.CTkEntry(
                frame,
                placeholder_text=f"Pie de la Figura {num}…",
                textvariable=pie_var,
                font=ctk.CTkFont(size=11),
                height=36)
            entry.grid(row=0, column=2, columnspan=2,
                       rowspan=2, padx=(0, 8), pady=8, sticky="ew")

            # Guardar referencia al StringVar para leer al exportar
            fig["_var"] = pie_var

            # Botón eliminar
            def _borrar(idx=i):
                # Leer pies actualizados antes de borrar
                self._sync_pies()
                self.figuras_manuales.pop(idx)
                self._refrescar_lista_figuras()
                self._set_status(f"Figura {idx+1} eliminada.")

            ctk.CTkButton(frame, text="✕", width=28, height=28,
                          fg_color="#b71c1c", hover_color="#7f0000",
                          command=_borrar,
                          font=ctk.CTkFont(size=12)).grid(
                row=0, column=4, padx=(0, 6), pady=6)

        n = len(self.figuras_manuales)
        self._figs_count_lbl.configure(
            text=f"{n} figura{'s' if n != 1 else ''}" if n else "Sin figuras",
            text_color="#a5d6a7" if n else "#888")

    def _sync_pies(self):
        """Sincroniza los campos de texto con self.figuras_manuales."""
        for fig in self.figuras_manuales:
            if "_var" in fig:
                fig["pie"] = fig["_var"].get()

    def _sync_titulos_tablas(self):
        for tab in self.tablas_manuales:
            if "_var_tit" in tab:
                tab["titulo"] = tab["_var_tit"].get()
            if "_var_anc" in tab:
                tab["ancla"]  = tab["_var_anc"].get()

    # ─── Tablas ───────────────────────────────────────────────────

    def _agregar_tabla(self):
        rutas = filedialog.askopenfilenames(
            title="Selecciona archivo(s) Excel",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        if not rutas: return
        for ruta in rutas:
            self.tablas_manuales.append({"ruta": ruta, "titulo": "", "ancla": ""})
        self._refrescar_lista_tablas()
        self._set_status(f"✓ {len(self.tablas_manuales)} tabla(s) en total.")

    def _limpiar_tablas(self):
        self.tablas_manuales = []
        self._refrescar_lista_tablas()
        self._tabs_count_lbl.configure(text="Sin tablas", text_color="#888")
        self._set_status("Tablas limpiadas.")

    def _refrescar_lista_tablas(self):
        for w in self._tabs_scroll.winfo_children():
            w.destroy()

        for i, tab_item in enumerate(self.tablas_manuales):
            num   = i + 1
            frame = ctk.CTkFrame(self._tabs_scroll, fg_color="#1a1a2e", corner_radius=6)
            frame.pack(fill="x", padx=4, pady=6)
            frame.columnconfigure(2, weight=1)

            # ── Columna izquierda: icono + etiqueta ──
            ctk.CTkLabel(frame, text="📊",
                         font=ctk.CTkFont(size=26), width=48).grid(
                row=0, column=0, rowspan=3, padx=(8, 4), pady=8, sticky="n")

            ctk.CTkLabel(frame, text=f"Tabla {num}",
                         font=ctk.CTkFont(size=12, weight="bold"),
                         text_color="#ce93d8").grid(
                row=0, column=1, padx=(0, 6), pady=(8, 0), sticky="w")

            ctk.CTkLabel(frame,
                         text=os.path.basename(tab_item["ruta"]),
                         font=ctk.CTkFont(size=9), text_color="#666").grid(
                row=1, column=1, padx=(0, 6), pady=0, sticky="w")

            # ── Campo 1: Título ──
            tit_var = ctk.StringVar(value=tab_item.get("titulo", ""))
            ctk.CTkEntry(
                frame,
                placeholder_text=f"Título de la Tabla {num}…",
                textvariable=tit_var,
                font=ctk.CTkFont(size=11), height=30
            ).grid(row=0, column=2, columnspan=2,
                   padx=(0, 8), pady=(8, 2), sticky="ew")
            tab_item["_var_tit"] = tit_var

            # ── Campo 2: Texto ancla ──
            ctk.CTkLabel(frame,
                         text="📍 Pega aquí el texto del párrafo donde va la tabla:",
                         font=ctk.CTkFont(size=9), text_color="#ce93d8").grid(
                row=2, column=1, columnspan=3, padx=(0, 8), pady=(4, 0), sticky="w")

            anc_var = ctk.StringVar(value=tab_item.get("ancla", ""))
            ctk.CTkEntry(
                frame,
                placeholder_text='Ej: "la Dra. Elena Centeno (Tabla 1)."',
                textvariable=anc_var,
                font=ctk.CTkFont(size=10), height=30
            ).grid(row=3, column=1, columnspan=3,
                   padx=(0, 8), pady=(0, 8), sticky="ew")
            tab_item["_var_anc"] = anc_var

            # ── Botón eliminar ──
            def _borrar_t(idx=i):
                self._sync_titulos_tablas()
                self.tablas_manuales.pop(idx)
                self._refrescar_lista_tablas()
                self._set_status(f"Tabla {idx+1} eliminada.")

            ctk.CTkButton(frame, text="✕", width=28, height=28,
                          fg_color="#b71c1c", hover_color="#7f0000",
                          command=_borrar_t,
                          font=ctk.CTkFont(size=12)).grid(
                row=0, column=4, padx=(0, 6), pady=8)

        n = len(self.tablas_manuales)
        self._tabs_count_lbl.configure(
            text=f"{n} tabla{'s' if n != 1 else ''}" if n else "Sin tablas",
            text_color="#ce93d8" if n else "#888")

    # ═════════════════════════════════════════════════════════════
    # EXTRACCIÓN PDF
    # ═════════════════════════════════════════════════════════════

    def _info_fuente(self, block):
        sizes, bold, italic, fuentes = [], False, False, []
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                sizes.append(span["size"])
                fuentes.append(span.get("font", "").lower())
                if span["flags"] & (1 << 4): bold   = True
                if span["flags"] & (1 << 1): italic = True
        avg = sum(sizes)/len(sizes) if sizes else 10.0
        dom = Counter(fuentes).most_common(1)[0][0] if fuentes else ""
        return avg, bold, italic, dom

    def _texto_bloque(self, block):
        partes = ["".join(s["text"] for s in l.get("spans", []))
                  for l in block.get("lines", [])]
        t = re.sub(r"\s+", " ", " ".join(partes)).strip()
        return re.sub(r"(\w)-\s+(\w)", r"\1\2", t)

    def _clasificar_auto(self, texto, size, bold, italic, font, body_size):
        t, t_low = texto.strip(), texto.strip().lower()

        secciones = {
            "resumen","abstract","resumen no técnico","non-technical abstract",
            "palabras clave","keywords","introducción","introduction",
            "conclusiones","conclusions","referencias","references",
            "agradecimientos","acknowledgements","discusión","discussion",
            "metodología","methods","resultados","results",
            "contribuciones de los autores",
            "paleontología sistemática","systematic palaeontology",
        }
        if t_low in secciones: return "Encabezado sección"
        if _es_como_citar(t):  return "Cómo citar"
        if _es_fecha_mss(t) or _es_doi(t): return "Fecha manuscrito"
        if re.match(r"^(palabras\s+clave|keywords)\s*[:\.]", t_low):
            return "Palabras clave"
        if re.search(r"@[\w\-\.]+\.\w{2,}", t) and len(t) < 120:
            return "Email / Metadatos"

        s = round(size)
        # ── Subtítulos numerados ─────────────────────────────────────────────
        # Nivel 1:  "4. Texto"      (sin punto decimal) → bold
        # Nivel 2+: "4.1 Texto"    (con punto decimal) → sin bold
        if re.match(r"^\d+\.\s+\S", t) and not re.match(r"^\d+\.\d", t) and s <= 12:
            return "Subencabezado"
        if re.match(r"^\d+\.\d+", t) and s <= 12:
            return "Subencabezado-bajo"
        if s >= 13 and bold and not italic: return "Título principal"
        if s >= 13 and bold and italic:     return "Título secundario"
        if s == 13 and not bold and not italic: return "Autores"
        if s == 12 and italic:  return "Email / Metadatos"
        if s == 12 and not bold: return "Normal"
        if s == 10 and bold:    return "Encabezado sección"
        if s == 10 and not bold: return "Subencabezado"
        if s == 9:
            is_tnr = "times" in font or "roman" in font
            if is_tnr or len(t) > 100: return "Resumen / Abstract"
            if len(t) < 50: return "Filiación"
            return "Resumen / Abstract"
        if len(t) < 4: return "Ignorar"
        return "Normal"

    def evento_cargar_archivo(self):
        ruta = filedialog.askopenfilename(
            title="Selecciona el PDF",
            filetypes=[("Archivos PDF", "*.pdf")])
        if not ruta: return

        # ── Bloquear UI mientras carga ────────────────────────────
        self._btn_cargar.configure(
            state="disabled",
            text="⏳  Analizando…",
            fg_color="#555"
        )
        self._status.configure(text_color="#888")   # resetear color
        self._set_status("⏳  Analizando PDF, por favor espera…")
        self.update()   # forzar redibujado antes de bloquear el hilo

        for w in self.frame_scroll.winfo_children():
            w.destroy()
        self.datos_bloques.clear()

        try:
            doc = fitz.open(ruta)

            all_sizes = []
            for pnum in range(len(doc)):
                for block in doc.load_page(pnum).get_text("dict")["blocks"]:
                    if block["type"] == 0:
                        for line in block.get("lines", []):
                            for span in line.get("spans", []):
                                all_sizes.append(round(span["size"]))
            body_size = Counter(all_sizes).most_common(1)[0][0] if all_sizes else 12

            # ── Filtro de cornisas: solo por posición, desde página 2 ──────
            def _es_cornisa(by0, by1, page_h, pnum):
                if pnum == 0:
                    return False
                return by1 < page_h * 0.09 or by0 > page_h * 0.91

            # ─── PASO 1: extraer todos los bloques de texto crudos ────────
            raw = []   # list of {"texto", "size", "bold", "italic", "font"}
            for pnum in range(len(doc)):
                page   = doc.load_page(pnum)
                page_h = page.rect.height
                for block in page.get_text("dict")["blocks"]:
                    if block["type"] == 1:
                        x0, y0, x1, y1 = block["bbox"]
                        w_px, h_px = abs(x1-x0), abs(y1-y0)
                        ignorar = (w_px < 32 and h_px < 32) or \
                                  _es_cornisa(y0, y1, page_h, pnum)
                        raw.append({
                            "texto": f"[IMAGEN {w_px:.0f}×{h_px:.0f}px]",
                            "clasificacion": "Ignorar" if ignorar else "Imagen",
                            "size": 0, "bold": False, "italic": False,
                            "imagen": True,
                        })
                        continue
                    if block["type"] != 0:
                        continue
                    bx0, by0, bx1, by1 = block["bbox"]
                    if _es_cornisa(by0, by1, page_h, pnum):
                        continue
                    size, bold, italic, font = self._info_fuente(block)
                    texto = self._texto_bloque(block)
                    if not texto or len(texto) < 3:
                        continue
                    raw.append({
                        "texto": texto, "size": size,
                        "bold": bold, "italic": italic, "font": font,
                        "imagen": False, "clasificacion": None,
                    })

            # ─── PASO 2: detectar zonas ───────────────────────────────────
            # Zona A (portada/frente): antes del primer encabezado numerado
            # Zona B (cuerpo):         desde primer "N. Texto" hasta "Referencias"
            # Zona C (post-refs):      después de "Referencias"
            #
            # Patrón de encabezado numerado nivel-1: "4.Texto" o "4. Texto"
            _pat_nivel1 = re.compile(r"^\d+\.\s*\S")
            _pat_nivel2 = re.compile(r"^\d+\.\d+")
            _SECCIONES_EXACTAS = {
                "resumen", "abstract", "resumen no técnico",
                "non-technical abstract", "palabras clave", "keywords",
                "referencias", "references", "conclusiones", "conclusions",
                "agradecimientos", "acknowledgements", "discusión",
                "resultados", "introducción", "metodología",
                "contribuciones de los autores",
            }

            zona_b_inicio = None   # índice en raw donde empieza el cuerpo
            zona_b_fin    = None   # índice en raw donde empieza "Referencias"

            for i, r in enumerate(raw):
                if r["imagen"] or r["clasificacion"] == "Ignorar":
                    continue
                t_low = r["texto"].strip().lower()
                s     = round(r["size"])

                # Primer encabezado numerado nivel-1 → inicio zona B
                if zona_b_inicio is None:
                    if _pat_nivel1.match(r["texto"].strip()) and \
                       not _pat_nivel2.match(r["texto"].strip()) and \
                       s <= 12:
                        zona_b_inicio = i

                # "Referencias" o "References" exacto → fin zona B
                if zona_b_inicio is not None and zona_b_fin is None:
                    if t_low in ("referencias", "references"):
                        zona_b_fin = i

            # ─── PASO 3: clasificar cada bloque según su zona ─────────────
            bloques_raw = []
            for i, r in enumerate(raw):
                if r["imagen"]:
                    bloques_raw.append({
                        "contenido": r["texto"],
                        "clasificacion": r["clasificacion"],
                        "size": 0, "bold": False, "italic": False,
                    })
                    continue

                texto = r["texto"].strip()
                t_low = texto.lower()
                s     = round(r["size"])

                # ── Zona B: cuerpo del artículo ──────────────────────────
                en_b = zona_b_inicio is not None and \
                       i >= zona_b_inicio and \
                       (zona_b_fin is None or i < zona_b_fin)

                if en_b:
                    if _pat_nivel1.match(texto) and not _pat_nivel2.match(texto):
                        cls = "Subencabezado"
                    elif _pat_nivel2.match(texto):
                        cls = "Subencabezado-bajo"
                    elif _es_como_citar(texto):
                        cls = "Cómo citar"
                    elif _es_fecha_mss(texto) or _es_doi(texto):
                        cls = "Fecha manuscrito"
                    elif t_low in _SECCIONES_EXACTAS:
                        cls = "Encabezado sección"
                    elif re.search(r"\btabla\s+\d+[\.\:\s]", t_low):
                        cls = "Título tabla"   # "Tabla 1. descripción"
                    else:
                        cls = "Cuerpo"
                else:
                    # ── Zona A / C: clasificación normal ─────────────────
                    if i == zona_b_fin:
                        cls = "Encabezado sección"   # la palabra "Referencias"
                    else:
                        cls = self._clasificar_auto(
                            texto, r["size"], r["bold"],
                            r["italic"], r["font"], body_size)

                bloques_raw.append({
                    "contenido": texto,
                    "clasificacion": cls,
                    "size": r["size"],
                    "bold": r["bold"],
                    "italic": r["italic"],
                })

            # ─── PASO 3b: suprimir filas de tabla del PDF ─────────────────
            # Regla universal para todas las revistas:
            # 1. Al detectar "Título tabla" → activar modo tabla
            # 2. También detectar filas-de-tabla por patrón (año + autor, celdas cortas)
            #    incluso sin haber visto el título todavía
            # 3. Salir del modo cuando aparece un párrafo real (>200 chars + punto final)
            #    o un encabezado/subencabezado

            # Patrón universal de fila de tabla académica:
            # - Empieza con año (1800-2099) seguido de texto corto
            # - O es un bloque muy corto (<80 chars) con formato de celda
            _pat_fila_tabla = re.compile(
                r"^(1[89]\d{2}|20\d{2})\s+\S"   # año AAAA + contenido
            )

            def _parece_fila_tabla(item):
                t   = item["contenido"].strip()
                cls = item["clasificacion"]
                if cls != "Cuerpo": return False
                # Fila que empieza con año
                if _pat_fila_tabla.match(t): return True
                # Bloque muy corto de celda (sin punto al final, sin mayúscula inicial
                # que sea inicio de oración real)
                if len(t) < 80 and not t.endswith(".") and not t.endswith("?"):
                    # Si es todo números, años o nombres cortos → celda
                    if re.match(r"^[\w\s\-\.,;&éáíóúüñÉÁÍÓÚÜÑ]{3,79}$", t):
                        return True
                return False

            bloques_sin_tabla_pdf = []
            en_modo_tabla = False
            for item in bloques_raw:
                cls_i = item["clasificacion"]

                # Activar por título explícito
                if cls_i == "Título tabla":
                    en_modo_tabla = True
                    bloques_sin_tabla_pdf.append(item)
                    continue

                # Activar por patrón de fila (sin necesitar el título)
                if not en_modo_tabla and _parece_fila_tabla(item):
                    en_modo_tabla = True
                    # No agregar esta fila — es contenido de tabla del PDF
                    continue

                if en_modo_tabla:
                    # ¿Salimos del modo tabla?
                    es_fin = cls_i in (
                        "Subencabezado", "Subencabezado-bajo",
                        "Encabezado sección", "Cómo citar",
                        "Fecha manuscrito",
                    ) or (
                        cls_i == "Cuerpo" and
                        len(item["contenido"]) > 200 and
                        item["contenido"].rstrip()[-1] in ".?!"
                    )
                    if es_fin:
                        en_modo_tabla = False
                        bloques_sin_tabla_pdf.append(item)
                    # else: fila de tabla del PDF → descartar
                else:
                    bloques_sin_tabla_pdf.append(item)
            bloques_raw = bloques_sin_tabla_pdf
            HEADERS_EMBEBIDOS = [
                "Non-technical Abstract", "Resumen no técnico",
                "Resumen no Técnico", "Non-Technical Abstract",
            ]
            # Patrón que identifica títulos de tabla embebidos en un bloque mayor.
            # Captura todo desde "Tabla N." hasta el fin de esa "oración" (hasta \n
            # o hasta el final del texto).
            _pat_tabla_embebida = re.compile(
                r'(?<!\w)(Tabla\s+\d+[\.\:\s][^\n]{0,200})',
                re.IGNORECASE
            )

            bloques_clean = []
            for item in bloques_raw:
                txt    = item["contenido"]
                partido = False

                # 1. Separar encabezados de sección embebidos
                for hdr in HEADERS_EMBEBIDOS:
                    idx = txt.find(hdr)
                    if idx > 0:
                        antes = txt[:idx].strip()
                        if antes:
                            b1 = dict(item); b1["contenido"] = antes
                            b1["clasificacion"] = self._clasificar_auto(
                                antes, item["size"], item["bold"],
                                item["italic"], "", body_size)
                            bloques_clean.append(b1)
                        bloques_clean.append({
                            "contenido": hdr, "clasificacion": "Encabezado sección",
                            "size": item["size"], "bold": item["bold"],
                            "italic": item["italic"],
                        })
                        despues = txt[idx+len(hdr):].strip()
                        if despues:
                            b3 = dict(item); b3["contenido"] = despues
                            bloques_clean.append(b3)
                        partido = True
                        break

                if partido:
                    continue

                # 2. Separar títulos de tabla embebidos en bloques Cuerpo
                #    Ej: "...texto previo. Tabla 1. Lista de trabajos…\nTexto sig."
                if item["clasificacion"] == "Cuerpo":
                    m = _pat_tabla_embebida.search(txt)
                    if m and m.start() > 0:
                        # hay texto ANTES del título de tabla
                        antes = txt[:m.start()].strip()
                        titulo_tab = m.group(1).strip()
                        despues = txt[m.end():].strip()

                        if antes:
                            b1 = dict(item); b1["contenido"] = antes
                            bloques_clean.append(b1)

                        bloques_clean.append({
                            "contenido": titulo_tab,
                            "clasificacion": "Título tabla",
                            "size": item["size"],
                            "bold": item["bold"],
                            "italic": item["italic"],
                        })

                        if despues:
                            b3 = dict(item); b3["contenido"] = despues
                            bloques_clean.append(b3)
                        continue   # no agregar el bloque original
                    elif m and m.start() == 0:
                        # el bloque empieza directamente con "Tabla N." → ya
                        # debería haberse clasificado arriba, pero por si acaso:
                        item["clasificacion"] = "Título tabla"

                bloques_clean.append(item)

            # ─── PASO 5: fusionar bloques Cuerpo/Normal consecutivos ──────
            # Encabezados de nivel-1 rompen la fusión → cada sección queda
            # como: [Subencabezado] + [Cuerpo fusionado]
            # Clases que NUNCA se fusionan:
            NO_FUSIONAR = {
                "Título principal", "Título secundario", "Autores",
                "Filiación", "Email / Metadatos", "Cómo citar",
                "Fecha manuscrito", "Encabezado sección",
                "Subencabezado",
                "Palabras clave", "Referencia", "Título tabla",
                "Pie de figura", "Imagen", "Ignorar",
            }
            # "Subencabezado-bajo" se embebe en el bloque Cuerpo con §SUB§
            fusionados = []
            buf_txt  = []
            buf_item = None

            def _vaciar():
                nonlocal buf_item
                if buf_txt and buf_item is not None:
                    merged = dict(buf_item)
                    merged["contenido"]     = "\n\n".join(buf_txt)
                    merged["clasificacion"] = "Cuerpo"
                    fusionados.append(merged)
                buf_txt.clear()
                buf_item = None

            for item in bloques_clean:
                cls = item["clasificacion"]
                if cls in NO_FUSIONAR:
                    _vaciar()
                    fusionados.append(item)
                elif cls == "Subencabezado-bajo":
                    if buf_item is None:
                        buf_item = item
                    buf_txt.append("§SUB§" + item["contenido"])
                else:
                    # Cuerpo / Normal / Resumen / Abstract → acumular
                    if buf_item is None:
                        buf_item = item
                    buf_txt.append(item["contenido"])
            _vaciar()
            bloques_utiles = fusionados

            for item in bloques_utiles:
                self._crear_bloque_ui(item)

            conteo  = Counter(b["clasificacion"] for b in bloques_utiles)
            resumen = "  |  ".join(f"{k}: {v}" for k,v in conteo.most_common(6))
            self._status.configure(text_color="#66bb6a")
            self._set_status(
                f"✅  Análisis completo — {len(bloques_utiles)} bloques  |  base: {body_size}pt  |  {resumen}")
            self._mostrar_banner(
                f"✅  Análisis completo — {len(bloques_utiles)} bloques extraídos")
            self._actualizar_stats()
            # Mostrar zona Portada por defecto al cargar
            self._zona_seg.set("📋 Portada")
            self._cambiar_zona("📋 Portada")

        except Exception as e:
            self._set_status(f"❌ Error: {e}")
            import traceback; traceback.print_exc()

        finally:
            # ── Restaurar botón siempre, haya error o no ─────────
            self._btn_cargar.configure(
                state="normal",
                text="📂  Cargar PDF",
                fg_color="#2e86c1"
            )

    def _crear_bloque_ui(self, item):
        cls  = item["clasificacion"]
        cont = item["contenido"]
        size = item.get("size", 10)
        bold = item.get("bold", False)
        ital = item.get("italic", False)

        color   = COLOR_POR_CLASE.get(cls, "#2b2b2b")
        preview = (cont[:150]+"…") if len(cont)>150 else cont

        frame = ctk.CTkFrame(self.frame_scroll, fg_color=color, corner_radius=5)
        frame.pack(fill="x", padx=8, pady=2)
        frame.columnconfigure(1, weight=1)

        badge = f"{size:.0f}pt"+(" B" if bold else "")+(" I" if ital else "")
        ctk.CTkLabel(frame, text=badge, width=56,
                     font=ctk.CTkFont(size=9), text_color="#bbb"
                     ).grid(row=0, column=0, padx=(6,0), pady=5, sticky="w")
        ctk.CTkLabel(frame, text=preview, wraplength=670,
                     justify="left", anchor="w",
                     font=ctk.CTkFont(size=11)
                     ).grid(row=0, column=1, padx=6, pady=5, sticky="ew")

        menu = ctk.CTkOptionMenu(frame, values=OPCIONES, width=175,
                                  font=ctk.CTkFont(size=11),
                                  command=self._actualizar_stats)
        menu.set(cls)
        menu.grid(row=0, column=2, padx=(0,8), pady=5)

        self.datos_bloques.append({
            "contenido": cont, "menu": menu,
            "italic": ital,    "frame": frame,
        })

    # ═════════════════════════════════════════════════════════════
    # EXPORTAR HTML
    # ═════════════════════════════════════════════════════════════

    def evento_exportar_html(self):
        if not self.datos_bloques:
            self._set_status("⚠ Primero carga un PDF."); return
        ruta = filedialog.asksaveasfilename(
            defaultextension=".html",
            filetypes=[("Archivo HTML", "*.html")])
        if not ruta: return

        self._sync_pies()
        self._sync_titulos_tablas()

        # Decidir qué referencias usar
        # Si hay externas → usar esas; los bloques "Referencia" del PDF → ignorar
        usar_refs_externas = bool(self.referencias_externas)
        refs_a_usar = (self.referencias_externas if usar_refs_externas
                       else [b["contenido"] for b in self.datos_bloques
                             if b["menu"].get() == "Referencia"])

        # ── Separar bloques ─────────────────────────────────────────
        # Primero detectamos en qué índice empieza la sección de
        # Referencias para no recolectar DOIs sueltos de las citas.
        idx_refs_start = None
        for i, b in enumerate(self.datos_bloques):
            if b["menu"].get() == "Encabezado sección" and re.search(
                    r"referencia|reference", b["contenido"], re.I):
                idx_refs_start = i
                break

        cuerpo_bloques = []
        como_citar_lst = []
        fechas_mss_lst = []
        pies_pendientes= []

        for i, b in enumerate(self.datos_bloques):
            cls = b["menu"].get()
            # ¿Este bloque cae dentro de la zona de referencias del PDF?
            dentro_de_refs = (idx_refs_start is not None and i > idx_refs_start)

            if cls == "Cómo citar":
                if not dentro_de_refs:      # ignorar si viene de dentro de refs
                    como_citar_lst.append(b)
            elif cls == "Fecha manuscrito":
                if not dentro_de_refs:      # idem – evita DOIs de las citas
                    fechas_mss_lst.append(b)
            elif cls == "Imagen":
                pass                        # imágenes del PDF siempre ignoradas
            elif cls == "Referencia" and usar_refs_externas:
                pass                        # ignorar duplicados cuando hay .txt
            else:
                cuerpo_bloques.append(b)

        # Deduplicar por primeros 80 chars (por si el PDF los repite)
        def _dedup(lst):
            seen, out = set(), []
            for b in lst:
                key = b["contenido"].strip()[:80]
                if key not in seen:
                    seen.add(key); out.append(b)
            return out
        como_citar_lst = _dedup(como_citar_lst)
        fechas_mss_lst = _dedup(fechas_mss_lst)

        # Generar HTML
        lineas = [
            "<!DOCTYPE html>", '<html lang="es">',
            "<head>",
            '  <meta charset="UTF-8">',
            '  <meta name="viewport" content="width=device-width, initial-scale=1.0">',
            "  <title>Artículo</title>",
            HTML_CSS,
            "</head>", "<body><article>",
        ]

        en_refs = False

        for b in cuerpo_bloques:
            cls   = b["menu"].get()
            texto = esc(b["contenido"])
            ital  = b.get("italic", False)
            if cls == "Ignorar": continue

            # ── Si hay referencias externas y estamos dentro de la sección
            #    de referencias del PDF, suprimir TODO excepto nuevos
            #    encabezados de sección (que resetean en_refs).
            if usar_refs_externas and en_refs and cls != "Encabezado sección":
                continue

            if cls == "Título principal":
                lineas.append(f'<h1 class="titulo-principal">{texto}</h1>')
            elif cls == "Título secundario":
                lineas.append(f'<h2 class="titulo-secundario">{texto}</h2>')
            elif cls == "Autores":
                # Pasar el contenido crudo (sin escapar) para que _insertar_orcid
                # pueda dividir correctamente por ";" y luego escapar por dentro
                lineas.append(f'<p class="autores sin-sangria">{_insertar_orcid(b["contenido"])}</p>')
            elif cls == "Filiación":
                lineas.append(f'<p class="filiaciones sin-sangria">{texto}</p>')
            elif cls == "Email / Metadatos":
                txt_link = re.sub(r"([\w\.\-]+@[\w\-\.]+\.\w{2,})",
                                  r'<a href="mailto:\1">\1</a>', texto)
                lineas.append(f'<p class="email sin-sangria">{txt_link}</p>')
            elif cls == "Encabezado sección":
                en_refs = bool(re.search(r"referencia|reference", texto, re.I))
                lineas.append(f'<h2 class="seccion">{texto}</h2>')
                if en_refs and refs_a_usar:
                    lineas.append('<ol class="referencias">')
                    for ref in refs_a_usar:
                        lineas.append(f"  <li>{esc(ref)}</li>")
                    lineas.append("</ol>")
            elif cls == "Subencabezado":
                lineas.append(f'<h3 class="subseccion">{texto}</h3>')
            elif cls == "Subencabezado-bajo":
                lineas.append(f'<h3 class="subseccion-bajo">{texto}</h3>')
            elif cls == "Resumen / Abstract":
                tag = "abstract" if ital else "resumen"
                lineas.append(f'<p class="{tag} sin-sangria">{texto}</p>')
            elif cls == "Palabras clave":
                lineas.append(f'<p class="keywords sin-sangria">{texto}</p>')
            elif cls in ("Normal", "Cuerpo"):
                if en_refs and not usar_refs_externas:
                    for parte in b["contenido"].split("\n\n"):
                        parte = parte.strip()
                        if parte:
                            lineas.append(f'<p style="padding-left:1.5em;text-indent:-1.5em;font-size:10pt;">{esc(parte)}</p>')
                else:
                    for parte in b["contenido"].split("\n\n"):
                        parte = parte.strip()
                        if not parte: continue
                        if parte.startswith("§SUB§"):
                            lineas.append(f'<h3 class="subseccion-bajo">{esc(parte[5:])}</h3>')
                        else:
                            tag_p = 'class="cuerpo"' if cls == "Cuerpo" else ""
                            lineas.append(f'<p {tag_p}>{esc(parte)}</p>'.replace("  >", ">"))
            elif cls == "Referencia":
                lineas.append(f'<p style="padding-left:1.5em;text-indent:-1.5em;font-size:10pt;">{texto}</p>')
            elif cls == "Título tabla":
                pass   # suprimido del flujo inline — las tablas van al final
            elif cls == "Pie de figura":
                pies_pendientes.append(texto)

        # Post-referencias
        if como_citar_lst or fechas_mss_lst:
            lineas.append('<div class="post-referencias">')
            for b in como_citar_lst:
                lineas.append(f'<p class="como-citar">{esc(b["contenido"])}</p>')
            if fechas_mss_lst:
                lineas.append('<div class="fechas-manuscrito"><ul>')
                doi_items = []
                doi_seen  = set()
                for b in fechas_mss_lst:
                    t      = b["contenido"].strip()
                    doi_m  = re.search(r"(https?://doi\.org/\S+)", t)
                    if doi_m:
                        doi_clean = re.sub(r"\s+", "", doi_m.group(1))
                        if doi_clean not in doi_seen:
                            doi_seen.add(doi_clean)
                            doi_items.append(doi_clean)
                        t = t.replace(doi_m.group(1), "").strip(" .")
                    for parte in re.split(
                            r'(?<=[.!?])\s+(?=Manuscrito|Manuscript)', t):
                        parte = parte.strip()
                        if parte:
                            lineas.append(f"  <li>{esc(parte)}</li>")
                for doi in doi_items:
                    lineas.append(
                        f'  <li><a href="{doi}" target="_blank">{esc(doi)}</a></li>')
                lineas.append("</ul></div>")
            lineas.append("</div>")

        # ── Tablas: inserción inline o al final ──────────────────────────
        # Primero cerramos el article para tener el HTML completo del cuerpo.
        lineas.append("</article></body></html>")
        html_body = "\n".join(lineas)

        if self.tablas_manuales:
            # Ordenar las tablas por posición de su ancla en el HTML
            tablas_ordenadas = []
            for idx_t, t_item in enumerate(self.tablas_manuales, 1):
                ancla  = t_item.get("ancla", "").strip()
                titulo = t_item.get("titulo", "") or f"Tabla {idx_t}"
                thtml  = _excel_a_html_tabla(t_item["ruta"])
                bloque = (
                    f'\n<div class="tabla-wrapper">\n'
                    f'<p class="tabla-titulo">Tabla {idx_t}. {esc(titulo)}</p>\n'
                    f'{thtml}\n</div>\n'
                )
                pos = -1
                if ancla:
                    # Normalizar: quitar soft-hyphens, guiones de corte + salto, espacios
                    ancla_norm = re.sub(r"­", "", ancla)          # soft-hyphen Unicode
                    ancla_norm = re.sub(r"-\s+", "", ancla_norm)  # guion + espacio (corte PDF)
                    ancla_norm = re.sub(r"\s+", " ", ancla_norm).strip()
                    # Usar los ÚLTIMOS 60 chars: fin del párrafo = posición exacta
                    # de inserción, evita que el PDF haya partido el bloque antes
                    muestra = ancla_norm[-60:].strip()
                    # Regex tolerante a tags HTML y soft-hyphens embebidos
                    escaped = re.sub(r"([.+*?()\[\]{}\\|^$])", r"\\\1", muestra)
                    pattern = re.sub(r"\\ ", r"(?:(?:­|-)?\\s*(?:<[^>]+>)?\\s*)+", escaped)
                    m = re.search(pattern, html_body, re.IGNORECASE | re.DOTALL)
                    if m:
                        cierre = html_body.find("</p>", m.end())
                        pos = cierre + 4 if cierre != -1 else -1
                tablas_ordenadas.append((pos, idx_t, bloque))

            # Ordenar: primero las que tienen ancla (por posición asc), luego sin ancla
            tablas_ordenadas.sort(key=lambda x: (x[0] == -1, x[0], x[1]))

            # Insertar de atrás para adelante para no desplazar posiciones
            tablas_inline  = [(p, b) for p, _, b in tablas_ordenadas if p != -1]
            tablas_al_final = [b for p, _, b in tablas_ordenadas if p == -1]

            # Insertar inline de mayor a menor posición
            for pos, bloque in sorted(tablas_inline, key=lambda x: -x[0]):
                html_body = html_body[:pos] + bloque + html_body[pos:]

            # Las que no tuvieron ancla van antes del cierre </article>
            if tablas_al_final:
                seccion = (
                    '\n<div class="figuras-finales">\n'
                    '<h2 style="text-align:center;font-size:10pt;font-weight:700;'
                    'margin-bottom:14px;">Tablas</h2>\n'
                    + "".join(tablas_al_final)
                    + "</div>\n"
                )
                html_body = html_body.replace("</article>", seccion + "</article>", 1)

        # Figuras manuales al final — insertar antes de </article>
        figs = self.figuras_manuales
        if figs:
            figs_html = '\n<div class="figuras-finales">\n<h2>Figuras</h2>\n'
            pie_idx = 0
            for i, fig in enumerate(figs, 1):
                pie_txt = fig.get("pie", "") or next(
                    (pies_pendientes[pie_idx + k]
                     for k in range(len(pies_pendientes) - pie_idx)
                     if pie_idx + k < len(pies_pendientes)),
                    "")
                try:
                    src = _img_to_base64(fig["ruta"])
                except Exception:
                    src = "imagen.jpg"
                figs_html += f'<figure id="fig{i}">\n'
                figs_html += f'  <img src="{src}" alt="Figura {i}">\n'
                cap = f"<strong>Figura {i}.</strong> {esc(pie_txt)}" if pie_txt \
                      else f"<strong>Figura {i}.</strong>"
                figs_html += f"  <figcaption>{cap}</figcaption>\n</figure>\n"
            figs_html += "</div>\n"
            html_body = html_body.replace("</article>", figs_html + "</article>", 1)

        with open(ruta, "w", encoding="utf-8") as f:
            f.write(html_body)
        self._set_status(f"✓ HTML guardado en: {ruta}")

    # ═════════════════════════════════════════════════════════════
    # EXPORTAR XML
    # ═════════════════════════════════════════════════════════════

    def evento_exportar_xml(self):
        if not self.datos_bloques:
            self._set_status("⚠ Primero carga un PDF."); return
        ruta = filedialog.asksaveasfilename(
            defaultextension=".xml", filetypes=[("XML", "*.xml")])
        if not ruta: return

        usar_refs_externas = bool(self.referencias_externas)
        refs_xml = (self.referencias_externas if usar_refs_externas
                    else [b["contenido"] for b in self.datos_bloques
                          if b["menu"].get() == "Referencia"])

        def ex(t): return t.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")

        head = ['<?xml version="1.0" encoding="UTF-8"?>',
                '<article article-type="research-article" xml:lang="es"',
                '         xmlns:xlink="http://www.w3.org/1999/xlink">',
                "  <front><article-meta>"]
        body = ["  <body>"]
        back = ["  <back>"]
        sec  = False

        for b in self.datos_bloques:
            cls = b["menu"].get()
            txt = ex(b["contenido"])
            if cls in ("Ignorar","Imagen","Referencia"): continue
            if   cls == "Título principal":
                head.append(f"    <title-group><article-title>{txt}</article-title></title-group>")
            elif cls == "Título secundario":
                head.append(f"    <title-group><trans-title>{txt}</trans-title></title-group>")
            elif cls == "Autores":
                head.append(f"    <contrib-group><contrib><string-name>{txt}</string-name></contrib></contrib-group>")
            elif cls == "Filiación":
                head.append(f"    <aff>{txt}</aff>")
            elif cls == "Palabras clave":
                head.append(f"    <kwd-group><kwd>{txt}</kwd></kwd-group>")
            elif cls in ("Cómo citar","Fecha manuscrito"):
                back.append(f"    <notes><p>{txt}</p></notes>")
            elif cls == "Encabezado sección":
                if sec: body.append("    </sec>")
                body.append(f"    <sec><title>{txt}</title>"); sec = True
            elif cls == "Subencabezado":
                body.append(f"      <title>{txt}</title>")
            elif cls in ("Normal","Resumen / Abstract"):
                body.append(f"      <p>{txt}</p>")
            elif cls == "Pie de figura":
                body.append(f"      <fig><caption><p>{txt}</p></caption></fig>")

        if sec: body.append("    </sec>")
        if refs_xml:
            body.append("    <ref-list>")
            for ref in refs_xml:
                body.append(f"      <ref><mixed-citation>{ex(ref)}</mixed-citation></ref>")
            body.append("    </ref-list>")

        body.append("  </body>")
        back.append("  </back>\n</article>")
        head += ["  </article-meta></front>"]

        with open(ruta, "w", encoding="utf-8") as f:
            f.write("\n".join(head + body + back))
        self._set_status(f"✓ XML guardado en: {ruta}")

    # ═════════════════════════════════════════════════════════════
    # EXPORTAR EPUB
    # ═════════════════════════════════════════════════════════════

    def evento_exportar_epub(self):
        if not self.datos_bloques:
            self._set_status("⚠ Primero carga un PDF."); return
        ruta = filedialog.asksaveasfilename(
            defaultextension=".epub", filetypes=[("EPUB","*.epub")])
        if not ruta: return

        usar_refs_externas = bool(self.referencias_externas)
        refs_epub = (self.referencias_externas if usar_refs_externas
                     else [b["contenido"] for b in self.datos_bloques
                           if b["menu"].get() == "Referencia"])

        book = epub.EpubBook()
        book.set_title("Artículo de Revista")
        book.set_language("es")
        cap  = epub.EpubHtml(title="Contenido",
                              file_name="articulo.xhtml", lang="es")
        html = "<h1>Artículo</h1>\n"

        for b in self.datos_bloques:
            cls = b["menu"].get()
            txt = b["contenido"]
            if cls in ("Ignorar","Imagen","Referencia"): continue
            if   cls == "Título principal":          html += f"<h1>{txt}</h1>\n"
            elif cls == "Título secundario":         html += f"<h2><em>{txt}</em></h2>\n"
            elif cls == "Autores":                   html += f"<p><strong>{txt}</strong></p>\n"
            elif cls in ("Filiación","Email / Metadatos"): html += f"<p><small>{txt}</small></p>\n"
            elif cls == "Encabezado sección":        html += f"<h2>{txt}</h2>\n"
            elif cls == "Subencabezado":             html += f"<h3>{txt}</h3>\n"
            elif cls == "Resumen / Abstract":        html += f"<p><em>{txt}</em></p>\n"
            elif cls == "Palabras clave":            html += f"<p><strong>{txt}</strong></p>\n"
            elif cls == "Normal":                    html += f"<p>{txt}</p>\n"
            elif cls == "Pie de figura":             html += f"<p><small><em>{txt}</em></small></p>\n"
            elif cls in ("Cómo citar","Fecha manuscrito"): html += f"<p><small>{txt}</small></p>\n"

        if refs_epub:
            html += ("<h2>Referencias</h2>\n<ol>\n"
                     + "".join(f"  <li>{r}</li>\n" for r in refs_epub)
                     + "</ol>\n")

        cap.content = html
        book.add_item(cap)
        book.spine = ["nav", cap]
        epub.write_epub(ruta, book, {})
        self._set_status(f"✓ EPUB guardado en: {ruta}")


if __name__ == "__main__":
    app = LimpiadorEditorialApp()
    app.mainloop()