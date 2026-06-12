"""
ui/app_window.py
Clase LimpiadorEditorialApp — ventana principal rediseñada (v3).

Cambios visuales respecto a v2:
  • Sidebar blanca con iconos sutiles y texto oscuro (en lugar de azul oscuro)
  • Ítem activo con fondo azul suave + acento lateral redondeado
  • Topbar más delgada con logo UNAM real o placeholder elegante
  • Stepper más compacto y legible
  • Espaciados más generosos y tipografía más consistente
"""

import re
import unicodedata
from collections import Counter

import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import os

# ── core ──────────────────────────────────────────────────────────────────────
from core.jats_exporterv2  import build_jats_xml
from core.html_exporter    import build_html
from core.epub_exporter    import build_epub
from core.pdf_processor    import procesar_pdf
from core.constans import (
    OPCIONES,
    CLASE_COMPAT as _CLASE_COMPAT,
    COLORES_UI,
    COLOR_POR_CLASE,
    ESTILO_POR_CLASE,
)
from core.utils import (
    esc,
    parsear_referencias          as _parsear_referencias,
    es_como_citar                as _es_como_citar,
    es_encabezado_resumen        as _es_encabezado_resumen,
    es_fecha_mss                 as _es_fecha_mss,
    es_doi                       as _es_doi,
    limpiar_prefijo_pie_figura   as _limpiar_prefijo_pie_figura,
    limpiar_prefijo_titulo_tabla as _limpiar_prefijo_titulo_tabla,
    split_afiliaciones_linea     as _split_afiliaciones_linea,
)

# ── ui ────────────────────────────────────────────────────────────────────────
from ui.tabs       import tab_pdf, tab_autores, tab_afiliaciones, tab_referencias, tab_figuras
from ui.widgets    import bloque_widget

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

# ─── Paleta de colores ────────────────────────────────────────────────────────
C = {
    # Fondos generales
    "bg":            "#F7F8FC",   # fondo general muy suave
    "panel_bg":      "#FFFFFF",
    "panel_brd":     "#E4E9F0",

    # Topbar
    "topbar_bg":     "#FFFFFF",
    "topbar_brd":    "#E4E9F0",

    # Sidebar — ahora clara
    "sidebar_bg":    "#FFFFFF",
    "sidebar_brd":   "#E4E9F0",
    "sidebar_sel_bg":"#EEF2FF",  # azul muy pálido para ítem activo
    "sidebar_hov":   "#F5F7FF",  # hover sutilísimo
    "sidebar_accent":"#1B2A4A",  # barra lateral del ítem activo

    # Acento / botones primarios
    "accent":        "#1B2A4A",
    "accent_hov":    "#243860",
    "accent_light":  "#EEF2FF",

    # Botones secundarios
    "btn_sec":       "#F4F6FA",
    "btn_sec_hov":   "#E9EDF5",
    "btn_brd":       "#DDE3ED",

    # Texto
    "text_main":     "#1A2236",
    "text_sub":      "#5A6478",
    "text_light":    "#9AA3B5",
    "text_sidebar":  "#3A4558",   # texto de sidebar (oscuro, no blanco)

    # Exportar
    "export_brd":    "#CBD5E1",

    # Stepper
    "step_done":     "#1B2A4A",
    "step_active":   "#1B2A4A",
    "step_todo":     "#D1D8E6",

    # Misc
    "drop_bg":       "#F8FAFC",
    "drop_brd":      "#CBD5E1",
    "green":         "#16A34A",
    "status_bg":     "#F7F8FC",
    "status_brd":    "#E4E9F0",
}

# Secciones del sidebar: (clave, texto_icono, etiqueta, subtitulo)
SECCIONES = [
    ("pdf",          "PDF",          "PDF",           "Documento"),
    ("autores",      "Autores",      "Autores",        "ORCID"),
    ("afiliaciones", "Afil.",        "Afiliaciones",   "Instituciones"),
    ("referencias",  "Refs.",        "Referencias",    "Bibliografía"),
    ("figuras",      "Figs.",        "Figuras",        "Imágenes"),
    ("tablas",       "Tablas",       "Tablas",         "Datos"),
    ("config",       "Config.",      "Configuración",  "Preferencias"),
]

# Íconos SVG-style como texto Unicode limpio por sección
ICONOS = {
    "pdf":          "󰈙",   # fallback: texto
    "autores":      "󰀄",
    "afiliaciones": "󰏔",
    "referencias":  "󰈙",
    "figuras":      "󰋯",
    "tablas":       "󰓫",
    "config":       "󰒓",
}

# Fallback de íconos simples si la fuente no los soporta
ICONOS_SIMPLE = {
    "pdf":          "▣",
    "autores":      "◉",
    "afiliaciones": "◈",
    "referencias":  "◎",
    "figuras":      "◫",
    "tablas":       "▦",
    "config":       "◌",
}

# Pasos del stepper
PASOS      = ["PDF", "Autores", "Afiliaciones", "Referencias", "Figuras", "Tablas", "Exportar"]
PASOS_KEYS = ["pdf", "autores", "afiliaciones", "referencias", "figuras", "tablas", "exportar"]


# ─────────────────────────────────────────────────────────────────────────────

class LimpiadorEditorialApp(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.title("Editor Semántico — Paleontología Mexicana")
        self.geometry("1340x880")
        self.minsize(1100, 720)
        self.configure(fg_color=C["bg"])

        # ── Estado compartido ─────────────────────────────────────────────────
        self.datos_bloques:        list[dict] = []
        self.referencias_externas: list[str]  = []
        self.figuras_manuales:     list[dict] = []
        self.tablas_manuales:      list[dict] = []
        self.autores_orcid:        list[dict] = []
        self.afiliaciones_txt:     str        = ""
        self._vista_estructura:    bool       = True
        self._auto_fig_dir:        str | None = None
        self._auto_tab_dir:        str | None = None
        self._diag_tablas_auto:    str        = ""
        self._buscar_resultados:   list[int]  = []
        self._buscar_cursor:       int        = -1
        self._seccion_activa:      str        = "pdf"

        # Widgets que se llenan luego
        self._sidebar_btns:  dict = {}
        self._step_labels:   dict = {}
        self._step_circles:  dict = {}
        self._step_lines:    dict = {}
        self._paneles:       dict = {}

        self._construir_topbar()
        self._construir_cuerpo()
        self._construir_statusbar()
        self._construir_stepper()

        self._activar_seccion("pdf")

    # ═════════════════════════════════════════════════════════════════════════
    # Topbar
    # ═════════════════════════════════════════════════════════════════════════

    def _construir_topbar(self):
        topbar = ctk.CTkFrame(
            self,
            fg_color=C["topbar_bg"],
            corner_radius=0,
            height=64,
            border_width=0)
        topbar.pack(fill="x", side="top")
        topbar.pack_propagate(False)

        # Borde inferior sutil
        ctk.CTkFrame(
            topbar, height=1, fg_color=C["topbar_brd"], corner_radius=0
        ).place(relx=0, rely=1.0, relwidth=1.0, anchor="sw")

        # ── Logo UNAM + título ────────────────────────────────────────────────
        brand = ctk.CTkFrame(topbar, fg_color="transparent")
        brand.pack(side="left", padx=(20, 0), pady=0, fill="y")

        # Intenta cargar logo real; si no, placeholder elegante
        logo_path = os.path.join(os.path.dirname(__file__), "assets", "unam_logo.png")
        logo_frame = ctk.CTkFrame(
            brand, width=44, height=44, corner_radius=4,
            fg_color="transparent")
        logo_frame.pack(side="left", padx=(0, 14))
        logo_frame.pack_propagate(False)

        if os.path.exists(logo_path):
            try:
                img = Image.open(logo_path).resize((40, 40), Image.LANCZOS)
                self._logo_img = ImageTk.PhotoImage(img)
                ctk.CTkLabel(
                    logo_frame, image=self._logo_img, text=""
                ).place(relx=0.5, rely=0.5, anchor="center")
            except Exception:
                self._logo_placeholder(logo_frame)
        else:
            self._logo_placeholder(logo_frame)

        # Separador vertical
        ctk.CTkFrame(
            brand, width=1, fg_color=C["panel_brd"]
        ).pack(side="left", padx=(0, 14), pady=10, fill="y")

        # Texto título
        title_block = ctk.CTkFrame(brand, fg_color="transparent")
        title_block.pack(side="left", fill="y")

        ctk.CTkLabel(
            title_block,
            text="Editor Semántico",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color=C["text_main"],
            anchor="sw"
        ).pack(anchor="w", pady=(14, 0))

        ctk.CTkLabel(
            title_block,
            text="Paleontología Mexicana",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=C["text_sub"],
            anchor="nw"
        ).pack(anchor="w", pady=(1, 14))

        # ── Botones exportar (derecha) ─────────────────────────────────────────
        export_bar = ctk.CTkFrame(topbar, fg_color="transparent")
        export_bar.pack(side="right", padx=(0, 16), pady=0, fill="y")

        ctk.CTkLabel(
            export_bar,
            text="Exportar a:",
            font=ctk.CTkFont(size=12),
            text_color=C["text_sub"]
        ).pack(side="left", padx=(0, 8))

        # HTML — primario
        ctk.CTkButton(
            export_bar,
            text="HTML",
            command=self.evento_exportar_html,
            fg_color=C["accent"],
            hover_color=C["accent_hov"],
            text_color="#FFFFFF",
            width=80,
            height=34,
            corner_radius=7,
            font=ctk.CTkFont(size=12, weight="bold"),
            border_width=0
        ).pack(side="left", padx=(0, 6))

        # XML — secundario
        ctk.CTkButton(
            export_bar,
            text="XML",
            command=self.evento_exportar_xml,
            fg_color=C["btn_sec"],
            hover_color=C["btn_sec_hov"],
            text_color=C["text_main"],
            width=72,
            height=34,
            corner_radius=7,
            font=ctk.CTkFont(size=12),
            border_width=1,
            border_color=C["btn_brd"]
        ).pack(side="left", padx=(0, 6))

        # EPUB — secundario
        ctk.CTkButton(
            export_bar,
            text="EPUB",
            command=self.evento_exportar_epub,
            fg_color=C["btn_sec"],
            hover_color=C["btn_sec_hov"],
            text_color=C["text_main"],
            width=72,
            height=34,
            corner_radius=7,
            font=ctk.CTkFont(size=12),
            border_width=1,
            border_color=C["btn_brd"]
        ).pack(side="left", padx=(0, 10))

        # Engrane config
        ctk.CTkButton(
            export_bar,
            text="⚙",
            command=lambda: self._activar_seccion("config"),
            fg_color="transparent",
            hover_color=C["btn_sec"],
            text_color=C["text_sub"],
            width=34,
            height=34,
            corner_radius=7,
            font=ctk.CTkFont(size=15),
            border_width=0
        ).pack(side="left")

    def _logo_placeholder(self, parent):
        """Escudo UNAM placeholder con estilo."""
        inner = ctk.CTkFrame(
            parent, width=44, height=44, corner_radius=6,
            fg_color=C["accent"])
        inner.place(relx=0.5, rely=0.5, anchor="center")
        inner.pack_propagate(False)
        ctk.CTkLabel(
            inner,
            text="UNAM",
            font=ctk.CTkFont(size=8, weight="bold"),
            text_color="#FFFFFF"
        ).place(relx=0.5, rely=0.5, anchor="center")

    # ═════════════════════════════════════════════════════════════════════════
    # Cuerpo: sidebar + contenido
    # ═════════════════════════════════════════════════════════════════════════

    def _construir_cuerpo(self):
        self._cuerpo = ctk.CTkFrame(self, fg_color=C["bg"], corner_radius=0)
        self._cuerpo.pack(fill="both", expand=True)
        self._cuerpo.grid_columnconfigure(1, weight=1)
        self._cuerpo.grid_rowconfigure(0, weight=1)

        self._construir_sidebar()
        self._construir_panel_contenido()

    def _construir_sidebar(self):
        sidebar = ctk.CTkFrame(
            self._cuerpo,
            fg_color=C["sidebar_bg"],
            corner_radius=0,
            width=220)
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)

        # Borde derecho sutil
        ctk.CTkFrame(
            sidebar, width=1, fg_color=C["sidebar_brd"], corner_radius=0
        ).place(relx=1.0, rely=0, relheight=1.0, anchor="ne")

        # Encabezado de sección
        ctk.CTkLabel(
            sidebar,
            text="NAVEGACIÓN",
            font=ctk.CTkFont(size=9, weight="bold"),
            text_color=C["text_light"],
            anchor="w"
        ).pack(anchor="w", padx=20, pady=(22, 10))

        # Botones de sección
        for key, _icon, label, sub in SECCIONES:
            self._sidebar_btns[key] = self._crear_btn_sidebar(
                sidebar, key, label, sub)

        # Spacer
        spacer = ctk.CTkFrame(sidebar, fg_color="transparent")
        spacer.pack(fill="both", expand=True)

        # Pie del sidebar: logo/versión
        pie = ctk.CTkFrame(sidebar, fg_color="transparent")
        pie.pack(fill="x", padx=16, pady=(0, 16))

        ctk.CTkFrame(pie, height=1, fg_color=C["sidebar_brd"]).pack(
            fill="x", pady=(0, 12))

        pie_inner = ctk.CTkFrame(pie, fg_color="transparent")
        pie_inner.pack(fill="x")

        # Icono ammonite placeholder
        ammonite_frame = ctk.CTkFrame(
            pie_inner, width=36, height=36, corner_radius=18,
            fg_color=C["btn_sec"])
        ammonite_frame.pack(side="left", padx=(0, 10))
        ammonite_frame.pack_propagate(False)
        ctk.CTkLabel(
            ammonite_frame,
            text="◎",
            font=ctk.CTkFont(size=14),
            text_color=C["text_light"]
        ).place(relx=0.5, rely=0.5, anchor="center")

        pie_text = ctk.CTkFrame(pie_inner, fg_color="transparent")
        pie_text.pack(side="left")
        ctk.CTkLabel(
            pie_text,
            text="Paleontología Mexicana",
            font=ctk.CTkFont(size=10, weight="bold"),
            text_color=C["text_sidebar"],
            anchor="w"
        ).pack(anchor="w")
        ctk.CTkLabel(
            pie_text,
            text="Editorial • UNAM",
            font=ctk.CTkFont(size=9),
            text_color=C["text_light"],
            anchor="w"
        ).pack(anchor="w")

    def _crear_btn_sidebar(self, parent, key, label, sub):
        # Contenedor externo para el ítem
        btn_frame = ctk.CTkFrame(
            parent,
            fg_color="transparent",
            corner_radius=0,
            cursor="hand2")
        btn_frame.pack(fill="x", padx=0, pady=0)

        # Barra de acento izquierda (visible solo cuando está activo)
        accent_bar = ctk.CTkFrame(
            btn_frame, width=3, fg_color="transparent", corner_radius=2)
        accent_bar.place(x=0, rely=0.15, relheight=0.7)

        # Inner con padding y radio para el fondo de hover/selección
        inner = ctk.CTkFrame(
            btn_frame,
            fg_color="transparent",
            corner_radius=8)
        inner.pack(fill="x", padx=(10, 10), pady=2)

        # Número/índice como indicador visual de paso
        idx = [k for k, *_ in SECCIONES].index(key)
        num_frame = ctk.CTkFrame(
            inner, width=30, height=30, corner_radius=15,
            fg_color=C["btn_sec"])
        num_frame.pack(side="left", padx=(10, 10), pady=10)
        num_frame.pack_propagate(False)

        num_lbl = ctk.CTkLabel(
            num_frame,
            text=str(idx + 1),
            font=ctk.CTkFont(size=10, weight="bold"),
            text_color=C["text_light"])
        num_lbl.place(relx=0.5, rely=0.5, anchor="center")

        # Texto
        text_frame = ctk.CTkFrame(inner, fg_color="transparent")
        text_frame.pack(side="left", pady=10, fill="x", expand=True)

        main_lbl = ctk.CTkLabel(
            text_frame,
            text=label,
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=C["text_sidebar"],
            anchor="w")
        main_lbl.pack(anchor="w")

        sub_lbl = ctk.CTkLabel(
            text_frame,
            text=sub,
            font=ctk.CTkFont(size=10),
            text_color=C["text_light"],
            anchor="w")
        sub_lbl.pack(anchor="w")

        # Eventos
        def _on_click(e=None):
            self._activar_seccion(key)

        def _on_enter(e=None):
            if self._seccion_activa != key:
                inner.configure(fg_color=C["sidebar_hov"])

        def _on_leave(e=None):
            if self._seccion_activa != key:
                inner.configure(fg_color="transparent")

        for w in [btn_frame, inner, num_frame, num_lbl, text_frame, main_lbl, sub_lbl]:
            w.bind("<Button-1>", _on_click)
            w.bind("<Enter>", _on_enter)
            w.bind("<Leave>", _on_leave)

        return {
            "frame":      btn_frame,
            "inner":      inner,
            "num_frame":  num_frame,
            "num_lbl":    num_lbl,
            "main":       main_lbl,
            "sub":        sub_lbl,
            "accent_bar": accent_bar,
        }

    def _activar_seccion(self, key: str):
        prev = self._seccion_activa

        # Restaurar el anterior
        if prev in self._sidebar_btns:
            b = self._sidebar_btns[prev]
            b["inner"].configure(fg_color="transparent")
            b["main"].configure(
                text_color=C["text_sidebar"],
                font=ctk.CTkFont(size=13, weight="bold"))
            b["sub"].configure(text_color=C["text_light"])
            b["num_frame"].configure(fg_color=C["btn_sec"])
            b["num_lbl"].configure(text_color=C["text_light"])
            b["accent_bar"].configure(fg_color="transparent")

        self._seccion_activa = key

        # Resaltar el nuevo
        if key in self._sidebar_btns:
            b = self._sidebar_btns[key]
            b["inner"].configure(fg_color=C["sidebar_sel_bg"])
            b["main"].configure(
                text_color=C["accent"],
                font=ctk.CTkFont(size=13, weight="bold"))
            b["sub"].configure(text_color=C["accent"])
            b["num_frame"].configure(fg_color=C["accent"])
            b["num_lbl"].configure(text_color="#FFFFFF")
            b["accent_bar"].configure(fg_color=C["accent"])

        # Mostrar panel
        for k, panel in self._paneles.items():
            if k == key:
                panel.grid()
            else:
                panel.grid_remove()

        self._actualizar_stepper(key)

    # ═════════════════════════════════════════════════════════════════════════
    # Panel de contenido
    # ═════════════════════════════════════════════════════════════════════════

    def _construir_panel_contenido(self):
        contenedor = ctk.CTkFrame(
            self._cuerpo, fg_color=C["bg"], corner_radius=0)
        contenedor.grid(row=0, column=1, sticky="nsew")
        contenedor.grid_rowconfigure(0, weight=1)
        contenedor.grid_columnconfigure(0, weight=1)
        self._contenedor = contenedor
        self._crear_paneles()

    def _crear_paneles(self):
        def _panel(key):
            p = ctk.CTkFrame(
                self._contenedor,
                fg_color=C["bg"],
                corner_radius=0)
            p.grid(row=0, column=0, sticky="nsew")
            self._paneles[key] = p
            return p

        panel_pdf = _panel("pdf")
        self._panel_pdf = panel_pdf
        _panel("autores")
        _panel("afiliaciones")
        _panel("referencias")
        _panel("figuras")

        self._shim_tabs = _TabShim(self._paneles)
        self.tabs = self._shim_tabs

        tab_pdf.construir(self)
        tab_autores.construir(self)
        tab_afiliaciones.construir(self)
        tab_referencias.construir(self)
        tab_figuras.construir(self)

        _panel("tablas")

        panel_cfg = _panel("config")
        self._construir_panel_config(panel_cfg)

        for p in self._paneles.values():
            p.grid_remove()

    def _construir_panel_config(self, parent):
        ctk.CTkLabel(
            parent,
            text="Configuración",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=C["text_main"]
        ).pack(anchor="w", padx=36, pady=(36, 4))
        ctk.CTkLabel(
            parent,
            text="Preferencias de la aplicación",
            font=ctk.CTkFont(size=13),
            text_color=C["text_sub"]
        ).pack(anchor="w", padx=36)

    # ═════════════════════════════════════════════════════════════════════════
    # Stepper (parte inferior)
    # ═════════════════════════════════════════════════════════════════════════

    def _construir_stepper(self):
        stepper_bg = ctk.CTkFrame(
            self,
            fg_color="#FFFFFF",
            corner_radius=0,
            height=76,
            border_width=0)
        stepper_bg.pack(fill="x", side="bottom", before=self._status_bar)
        stepper_bg.pack_propagate(False)

        # Borde superior
        ctk.CTkFrame(
            stepper_bg, height=1, fg_color=C["topbar_brd"], corner_radius=0
        ).place(relx=0, rely=0, relwidth=1.0)

        inner = ctk.CTkFrame(stepper_bg, fg_color="transparent")
        inner.place(relx=0.5, rely=0.5, anchor="center")

        for i, (paso, key) in enumerate(zip(PASOS, PASOS_KEYS)):
            col_frame = ctk.CTkFrame(inner, fg_color="transparent")
            col_frame.pack(side="left", padx=2)

            # Círculo numerado
            circle = ctk.CTkFrame(
                col_frame,
                width=30,
                height=30,
                corner_radius=15,
                fg_color=C["step_todo"])
            circle.pack(anchor="center")
            circle.pack_propagate(False)

            num_lbl = ctk.CTkLabel(
                circle,
                text=str(i + 1),
                font=ctk.CTkFont(size=10, weight="bold"),
                text_color="#FFFFFF")
            num_lbl.place(relx=0.5, rely=0.5, anchor="center")

            # Etiqueta del paso
            lbl = ctk.CTkLabel(
                col_frame,
                text=paso,
                font=ctk.CTkFont(size=10),
                text_color=C["text_light"])
            lbl.pack(anchor="center", pady=(5, 0))

            self._step_circles[key] = circle
            self._step_labels[key]  = lbl

            # Línea conectora
            if i < len(PASOS) - 1:
                line = ctk.CTkFrame(
                    inner,
                    width=32,
                    height=2,
                    fg_color=C["step_todo"],
                    corner_radius=1)
                line.pack(side="left", padx=0, pady=(0, 20))
                self._step_lines[f"{key}_line"] = line

        self._stepper_frame = stepper_bg

    def _actualizar_stepper(self, key_activo: str):
        if key_activo not in PASOS_KEYS:
            return
        idx_activo = PASOS_KEYS.index(key_activo)
        for i, key in enumerate(PASOS_KEYS):
            if key not in self._step_circles:
                continue
            circle = self._step_circles[key]
            lbl    = self._step_labels[key]
            if i < idx_activo:
                circle.configure(fg_color=C["step_done"])
                lbl.configure(
                    text_color=C["text_sub"],
                    font=ctk.CTkFont(size=10, weight="bold"))
            elif i == idx_activo:
                circle.configure(fg_color=C["step_active"])
                lbl.configure(
                    text_color=C["text_main"],
                    font=ctk.CTkFont(size=10, weight="bold"))
            else:
                circle.configure(fg_color=C["step_todo"])
                lbl.configure(
                    text_color=C["text_light"],
                    font=ctk.CTkFont(size=10))

            # Línea entre pasos
            line_key = f"{key}_line"
            if line_key in self._step_lines:
                self._step_lines[line_key].configure(
                    fg_color=C["step_done"] if i < idx_activo else C["step_todo"])

    # ═════════════════════════════════════════════════════════════════════════
    # Status bar
    # ═════════════════════════════════════════════════════════════════════════

    def _construir_statusbar(self):
        self._status_bar = ctk.CTkFrame(
            self,
            fg_color=C["status_bg"],
            corner_radius=0,
            height=28,
            border_width=0)
        self._status_bar.pack(fill="x", side="bottom")
        self._status_bar.pack_propagate(False)

        # Borde superior
        ctk.CTkFrame(
            self._status_bar, height=1, fg_color=C["status_brd"], corner_radius=0
        ).place(relx=0, rely=0, relwidth=1.0)

        # Punto verde + texto
        dot_frame = ctk.CTkFrame(self._status_bar, fg_color="transparent")
        dot_frame.pack(side="left", padx=(14, 0))

        self._status_dot = ctk.CTkFrame(
            dot_frame, width=8, height=8, corner_radius=4,
            fg_color="#22C55E")
        self._status_dot.pack(side="left", padx=(0, 6))

        self._status = ctk.CTkLabel(
            dot_frame,
            text="Listo para comenzar",
            anchor="w",
            font=ctk.CTkFont(size=11),
            text_color=C["text_sub"])
        self._status.pack(side="left")

        ctk.CTkLabel(
            self._status_bar,
            text="v2.0.0",
            font=ctk.CTkFont(size=11),
            text_color=C["text_light"]
        ).pack(side="right", padx=14)

    # ═════════════════════════════════════════════════════════════════════════
    # Helpers de UI generales
    # ═════════════════════════════════════════════════════════════════════════

    def _set_status(self, msg: str):
        # Quitar prefijo "● " heredado
        clean = msg.lstrip("● ").lstrip("●").strip()
        self._status.configure(text=clean)
        self.update_idletasks()

    def _toggle_leyenda(self):
        self._leyenda_visible = not self._leyenda_visible
        if self._leyenda_visible:
            self._leyenda_panel.pack(fill="x", padx=0, pady=(0, 6))
        else:
            self._leyenda_panel.pack_forget()

    def _mostrar_banner(self, msg: str, color_bg="#DCFCE7", color_txt="#166534"):
        self._banner.configure(text=msg, fg_color=color_bg, text_color=color_txt)
        self.frame_scroll.pack_forget()
        self._banner.pack(fill="x", padx=8, pady=(0, 4))
        self.frame_scroll.pack(fill="both", expand=True, padx=0, pady=(0, 2))
        self.after(4000, self._ocultar_banner)

    def _ocultar_banner(self):
        self._banner.pack_forget()

    def _actualizar_stats(self, *_):
        if not self.datos_bloques:
            self._stats_lbl.configure(text="")
            return
        c = Counter(b["menu"].get() for b in self.datos_bloques)
        self._stats_lbl.configure(
            text="  ".join(f"{k[:5]}:{v}" for k, v in c.most_common(6)))

    def _aplicar_filtro(self, valor: str):
        for b in self.datos_bloques:
            cls = b["menu"].get()
            if valor == "Todos" or cls == valor:
                b["frame"].pack(fill="x", padx=8, pady=2)
            else:
                b["frame"].pack_forget()
        self._limpiar_busqueda_highlight()

    def _cambiar_panel_media(self, valor: str):
        tab_figuras.cambiar_panel_media(self, valor)

    # ═════════════════════════════════════════════════════════════════════════
    # Buscador
    # ═════════════════════════════════════════════════════════════════════════

    @staticmethod
    def _quitar_acentos(texto: str) -> str:
        return "".join(
            c for c in unicodedata.normalize("NFD", texto)
            if unicodedata.category(c) != "Mn"
        )

    def _buscar_en_bloques(self):
        query = self._entry_buscar.get().strip()
        self._limpiar_busqueda_highlight()
        if not query or not self.datos_bloques:
            self._lbl_buscar_cnt.configure(text="")
            return

        query_low = self._quitar_acentos(query.lower())
        self._buscar_resultados = [
            i for i, b in enumerate(self.datos_bloques)
            if query_low in self._quitar_acentos(b["_txtbox"].get("1.0", "end").lower())
        ]
        total = len(self._buscar_resultados)
        if total == 0:
            self._lbl_buscar_cnt.configure(text="0 resultados", text_color="#EF4444")
            self._buscar_cursor = -1
            return

        for idx in self._buscar_resultados:
            b = self.datos_bloques[idx]
            if "_color_original" not in b:
                b["_color_original"] = b["frame"].cget("fg_color")
            b["frame"].configure(border_width=2, border_color="#FACC15")

        self._buscar_cursor = 0
        self._ir_a_resultado_actual()

    def _navegar_busqueda(self, direccion: int):
        if not self._buscar_resultados:
            self._buscar_en_bloques()
            return
        total = len(self._buscar_resultados)
        self._buscar_cursor = (self._buscar_cursor + direccion) % total
        self._ir_a_resultado_actual()

    def _ir_a_resultado_actual(self):
        if not self._buscar_resultados:
            return
        total  = len(self._buscar_resultados)
        cursor = self._buscar_cursor
        idx_activo = self._buscar_resultados[cursor]

        self._lbl_buscar_cnt.configure(
            text=f"{cursor + 1} / {total}", text_color=C["text_sub"])

        for idx in self._buscar_resultados:
            self.datos_bloques[idx]["frame"].configure(
                border_width=2, border_color="#FACC15")

        frame_activo = self.datos_bloques[idx_activo]["frame"]
        frame_activo.configure(border_width=2, border_color="#F97316")
        self._resaltar_texto_en_textbox(idx_activo)
        self.after(40, lambda f=frame_activo: self._scroll_hasta_frame(f))

    def _resaltar_texto_en_textbox(self, idx: int):
        query = self._entry_buscar.get().strip()
        if not query:
            return
        tb = self.datos_bloques[idx]["_txtbox"]
        try:
            tb.tag_remove("busqueda", "1.0", "end")
        except Exception:
            pass
        contenido     = tb.get("1.0", "end-1c")
        query_low     = self._quitar_acentos(query.lower())
        contenido_low = self._quitar_acentos(contenido.lower())
        start = 0
        while True:
            pos = contenido_low.find(query_low, start)
            if pos == -1:
                break
            antes   = contenido[:pos]
            fila    = antes.count("\n") + 1
            col_i   = pos - antes.rfind("\n") - 1
            fin     = pos + len(query)
            despues = contenido[:fin]
            fila_f  = despues.count("\n") + 1
            col_f   = fin - despues.rfind("\n") - 1
            try:
                tb.tag_add("busqueda", f"{fila}.{col_i}", f"{fila_f}.{col_f}")
            except Exception:
                pass
            start = pos + 1
        try:
            tb.tag_config("busqueda", background="#F97316", foreground="#000000")
        except Exception:
            pass

    def _scroll_hasta_frame(self, frame):
        try:
            canvas = self.frame_scroll._parent_canvas
            frame.update_idletasks()
            canvas.update_idletasks()
            fy = frame.winfo_y()
            fh = frame.winfo_height()
            _, _, _, total_h = canvas.bbox("all")
            if not total_h or total_h <= 0:
                return
            visible_h = canvas.winfo_height()
            target_y  = fy - (visible_h - fh) // 2
            target_y  = max(0, min(target_y, total_h - visible_h))
            canvas.yview_moveto(target_y / total_h)
        except Exception:
            pass

    def _limpiar_busqueda_highlight(self):
        for b in self.datos_bloques:
            try:
                color_orig = b.get("_color_original", b["frame"].cget("fg_color"))
                b["frame"].configure(border_width=0, border_color=color_orig)
            except Exception:
                pass
            try:
                b["_txtbox"].tag_remove("busqueda", "1.0", "end")
            except Exception:
                pass
        self._buscar_resultados = []
        self._buscar_cursor     = -1
        try:
            self._lbl_buscar_cnt.configure(text="")
        except Exception:
            pass

    # ═════════════════════════════════════════════════════════════════════════
    # Delegaciones a tabs
    # ═════════════════════════════════════════════════════════════════════════

    def _agregar_autor(self):           tab_autores.agregar_autor(self)
    def _sync_autores(self):            tab_autores.sync_autores(self)
    def _refrescar_lista_autores(self): tab_autores.refrescar_lista(self)
    def _aplicar_autores(self):         tab_autores.sync_autores(self)
    def _limpiar_autores(self):         tab_autores.limpiar_autores(self)
    def _cargar_autores_excel(self):    tab_autores.cargar_autores_excel(self)

    def _cargar_afiliaciones(self):     tab_afiliaciones.cargar_afiliaciones(self)
    def _limpiar_afiliaciones(self):    tab_afiliaciones.limpiar_afiliaciones(self)
    def _refrescar_afiliaciones(self):  tab_afiliaciones.refrescar(self)

    def evento_cargar_referencias(self):  tab_referencias.cargar_referencias(self)
    def _limpiar_referencias(self):       tab_referencias.limpiar_referencias(self)
    def _refrescar_lista_refs(self):      tab_referencias.refrescar_lista(self)

    def _agregar_figura(self):              tab_figuras.agregar_figura(self)
    def _limpiar_figuras(self):             tab_figuras.limpiar_figuras(self)
    def _refrescar_lista_figuras(self):     tab_figuras.refrescar_lista_figuras(self)
    def _sync_pies(self):                   tab_figuras.sync_pies(self)
    def _limpiar_cache_figuras_auto(self):  tab_figuras.limpiar_cache_figuras_auto(self)

    def _agregar_tabla(self):               tab_figuras.agregar_tabla(self)
    def _limpiar_tablas(self):              tab_figuras.limpiar_tablas(self)
    def _refrescar_lista_tablas(self):      tab_figuras.refrescar_lista_tablas(self)
    def _sync_titulos_tablas(self):         tab_figuras.sync_titulos_tablas(self)
    def _limpiar_cache_tablas_auto(self):   tab_figuras.limpiar_cache_tablas_auto(self)
    def _remover_tablas_auto(self):         tab_figuras.remover_tablas_auto(self)

    def _crear_bloque_ui(self, item: dict):
        bloque_widget.crear_bloque_ui(self, item)

    # ═════════════════════════════════════════════════════════════════════════
    # Sincronización
    # ═════════════════════════════════════════════════════════════════════════

    def _sync_contenidos_bloques(self):
        for b in self.datos_bloques:
            tb = b.get("_txtbox")
            if tb is not None:
                b["contenido"] = tb.get("1.0", "end-1c")

    # ═════════════════════════════════════════════════════════════════════════
    # Evento — Cargar PDF
    # ═════════════════════════════════════════════════════════════════════════

    def evento_cargar_archivo(self):
        ruta = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if not ruta:
            return

        for b in self.datos_bloques:
            b["frame"].destroy()
        self.datos_bloques.clear()
        self._buscar_resultados = []
        self._buscar_cursor     = -1
        try:
            self._entry_buscar.delete(0, "end")
            self._lbl_buscar_cnt.configure(text="")
        except Exception:
            pass
        self._limpiar_cache_figuras_auto()
        self.figuras_manuales = []
        self._refrescar_lista_figuras()
        self._limpiar_cache_tablas_auto()
        self._remover_tablas_auto()
        self._refrescar_lista_tablas()

        self._btn_cargar.configure(
            state="disabled",
            text="Procesando…",
            fg_color="#93C5FD")

        try:
            resultado = procesar_pdf(ruta)

            bloques_utiles         = resultado["bloques"]
            figuras_auto           = resultado["figuras"]
            tablas_auto            = resultado["tablas"]
            body_size              = resultado["body_size"]
            resumen                = resultado["resumen"]
            self._diag_tablas_auto = resultado.get("diag_tablas", "")
            self._auto_fig_dir     = resultado.get("fig_dir")

            for item in bloques_utiles:
                self._crear_bloque_ui(item)

            self.figuras_manuales = figuras_auto
            self._refrescar_lista_figuras()
            self.tablas_manuales.extend(tablas_auto)
            self._refrescar_lista_tablas()

            self._set_status(
                f"Análisis completo — {len(bloques_utiles)} bloques  "
                f"|  base: {body_size}pt  |  {resumen}"
            )
            self._status.configure(text_color=C["text_sub"])
            self._status_dot.configure(fg_color="#22C55E")

            if not tablas_auto and self._diag_tablas_auto:
                diag_msg = (
                    f"No se extrajeron tablas automáticamente: "
                    f"{self._diag_tablas_auto}"
                )
                self._mostrar_banner(diag_msg, color_bg="#FEE2E2", color_txt="#991B1B")
                try:
                    messagebox.showwarning("Diagnóstico de tablas", diag_msg)
                except Exception:
                    pass
            else:
                self._mostrar_banner(
                    f"✓  Análisis completo — {len(bloques_utiles)} bloques extraídos")

            self._actualizar_stats()
            self._aplicar_filtro("Todos")

        except Exception as e:
            self._set_status(f"Error: {e}")
            self._status.configure(text_color="#EF4444")
            self._status_dot.configure(fg_color="#EF4444")
            import traceback
            traceback.print_exc()

        finally:
            self._btn_cargar.configure(
                state="normal",
                text="Seleccionar PDF",
                fg_color=C["accent"])

    # ═════════════════════════════════════════════════════════════════════════
    # Eventos — Exportar
    # ═════════════════════════════════════════════════════════════════════════

    def evento_exportar_html(self) -> str | None:
        if not self.datos_bloques:
            self._set_status("Primero carga un PDF.")
            return None
        ruta = filedialog.asksaveasfilename(
            defaultextension=".html",
            filetypes=[("Archivo HTML", "*.html")])
        if not ruta:
            return None

        self._sync_contenidos_bloques()
        self._sync_pies()
        self._sync_titulos_tablas()
        self._sync_autores()

        bloques_snapshot = [
            {
                "contenido":     b["contenido"],
                "clasificacion": b["menu"].get(),
                "italic":        b.get("italic", False),
            }
            for b in self.datos_bloques
        ]

        html_str = build_html(
            bloques              = bloques_snapshot,
            referencias_externas = self.referencias_externas,
            autores_orcid        = self.autores_orcid,
            afiliaciones_txt     = self.afiliaciones_txt,
            figuras              = self.figuras_manuales,
            tablas               = self.tablas_manuales,
        )

        with open(ruta, "w", encoding="utf-8") as f:
            f.write(html_str)
        self._set_status(f"HTML guardado en: {ruta}")
        return html_str

    def evento_exportar_xml(self):
        self.evento_exportar_jats()

    def evento_exportar_jats(self):
        if not self.datos_bloques:
            self._set_status("Primero carga un PDF.")
            return

        ruta_xml = filedialog.asksaveasfilename(
            defaultextension=".xml",
            filetypes=[("JATS XML", "*.xml"), ("Archivo XML", "*.xml")])
        if not ruta_xml:
            return

        try:
            self._sync_contenidos_bloques()
            self._sync_pies()
            self._sync_titulos_tablas()
            self._sync_autores()

            bloques_snapshot = [
                {
                    "contenido":     b.get("contenido", ""),
                    "clasificacion": b["menu"].get(),
                    "italic":        b.get("italic", False),
                }
                for b in self.datos_bloques
            ]

            xml_jats = build_jats_xml(
                bloques              = bloques_snapshot,
                referencias_externas = self.referencias_externas,
                autores_orcid        = self.autores_orcid,
                afiliaciones_txt     = self.afiliaciones_txt,
                figuras              = self.figuras_manuales,
                tablas               = self.tablas_manuales,
            )

            with open(ruta_xml, "w", encoding="utf-8") as f:
                f.write(xml_jats)
            self._set_status(f"JATS XML guardado en: {ruta_xml}")

        except Exception as exc:
            self._set_status(f"Error al generar JATS XML: {exc}")

    def evento_exportar_epub(self):
        if not self.datos_bloques:
            self._set_status("Primero carga un PDF.")
            return

        ruta_epub = filedialog.asksaveasfilename(
            defaultextension=".epub",
            filetypes=[("Archivo EPUB", "*.epub")])
        if not ruta_epub:
            return

        self._sync_contenidos_bloques()
        self._sync_pies()
        self._sync_titulos_tablas()
        self._sync_autores()

        bloques_snapshot = [
            {
                "contenido":     b["contenido"],
                "clasificacion": b["menu"].get(),
                "italic":        b.get("italic", False),
            }
            for b in self.datos_bloques
        ]

        html_str = build_html(
            bloques              = bloques_snapshot,
            referencias_externas = self.referencias_externas,
            autores_orcid        = self.autores_orcid,
            afiliaciones_txt     = self.afiliaciones_txt,
            figuras              = self.figuras_manuales,
            tablas               = self.tablas_manuales,
        )

        if not html_str or len(html_str) < 100:
            self._set_status("El HTML generado está vacío.")
            return

        titulo_art = next(
            (b["contenido"].strip() for b in self.datos_bloques
             if b["menu"].get() == "Título principal"),
            "Artículo",
        )
        autores_lista = [
            a["nombre"].strip() for a in self.autores_orcid
            if a.get("nombre", "").strip()
        ] if self.autores_orcid else []

        doi_art = ""
        for b in self.datos_bloques:
            m = re.search(r"https?://doi\.org/\S+", b["contenido"])
            if m:
                doi_art = m.group(0).rstrip(".")
                break

        secciones = [
            (re.sub(r"[^a-z0-9]", "-", b["contenido"].strip().lower())[:40].strip("-"),
             b["contenido"].strip())
            for b in self.datos_bloques
            if b["menu"].get() == "Encabezado sección" and b["contenido"].strip()
        ]

        try:
            epub_bytes = build_epub(
                html_str  = html_str,
                titulo    = titulo_art,
                autores   = autores_lista,
                doi       = doi_art,
                secciones = secciones,
            )
            with open(ruta_epub, "wb") as f:
                f.write(epub_bytes)
            self._set_status(f"EPUB guardado en: {ruta_epub}")
        except Exception as exc:
            self._set_status(f"Error al generar EPUB: {exc}")


# ─────────────────────────────────────────────────────────────────────────────
# _TabShim
# ─────────────────────────────────────────────────────────────────────────────

_TAB_KEY_MAP = {
    "📄  PDF":           "pdf",
    "👥  Autores":       "autores",
    "🏛️  Afiliaciones":  "afiliaciones",
    "📋  Referencias":   "referencias",
    "🖼️  Figuras":       "figuras",
    # Sin emoji (fallback)
    "PDF":               "pdf",
    "Autores":           "autores",
    "Afiliaciones":      "afiliaciones",
    "Referencias":       "referencias",
    "Figuras":           "figuras",
}

# Mapeo por palabra clave (tolerante a variantes de emoji)
_KEYWORD_MAP = {
    "pdf":          "pdf",
    "autores":      "autores",
    "afiliaciones": "afiliaciones",
    "referencias":  "referencias",
    "figuras":      "figuras",
}

def _normalizar_tab_key(nombre: str) -> str:
    """Extrae la clave de panel a partir del nombre del tab, tolerante a emojis."""
    # Intento directo
    if nombre in _TAB_KEY_MAP:
        return _TAB_KEY_MAP[nombre]
    # Buscar por palabra clave en el string (case-insensitive, sin emojis)
    nombre_lower = nombre.lower()
    for kw, key in _KEYWORD_MAP.items():
        if kw in nombre_lower:
            return key
    return nombre


class _TabShim:
    """Shim que imita CTkTabview.tab(nombre) devolviendo el panel correcto."""

    def __init__(self, paneles: dict):
        self._paneles = paneles

    def tab(self, nombre: str):
        key = _normalizar_tab_key(nombre)
        return self._paneles.get(key)