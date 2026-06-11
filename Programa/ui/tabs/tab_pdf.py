"""
ui/tabs/tab_pdf.py
Panel "PDF" rediseñado (v3) — alineado con la inspiración limpia.

Cambios visuales:
  • Encabezado con subtítulo descriptivo
  • Toolbar con separadores y más aire
  • Drop zone más grande y centrada, con icono mejorado
  • Info del documento con subrayado en el título (como la inspo)
  • Fondo general #F7F8FC (mismo que app_window)
"""
import unicodedata
import customtkinter as ctk
from core.constans import OPCIONES, COLORES_UI, COLOR_POR_CLASE, ESTILO_POR_CLASE

# ── Paleta (debe coincidir con app_window.py v3) ──────────────────────────────
C = {
    "bg":            "#F7F8FC",
    "panel_bg":      "#FFFFFF",
    "panel_brd":     "#E4E9F0",
    "accent":        "#1B2A4A",
    "accent_hov":    "#243860",
    "accent_light":  "#EEF2FF",
    "text_main":     "#1A2236",
    "text_sub":      "#5A6478",
    "text_light":    "#9AA3B5",
    "toolbar_bg":    "#FFFFFF",
    "toolbar_brd":   "#E4E9F0",
    "drop_bg":       "#F8FAFC",
    "drop_brd":      "#DDE3ED",
    "drop_brd_dash": "#CBD5E1",
    "scroll_bg":     "#FFFFFF",
    "btn_sec":       "#F4F6FA",
    "btn_sec_hov":   "#E9EDF5",
    "btn_brd":       "#DDE3ED",
}


# ─────────────────────────────────────────────────────────────────────────────
# Construcción del panel
# ─────────────────────────────────────────────────────────────────────────────

def construir(app):
    """Construye todos los widgets del panel PDF."""
    tab = app.tabs.tab("📄  PDF")
    tab.configure(fg_color=C["bg"])

    # ── Encabezado ─────────────────────────────────────────────────────────────
    _construir_header(app, tab)

    # ── Toolbar ────────────────────────────────────────────────────────────────
    _construir_toolbar(app, tab)

    # ── Leyenda (oculta por defecto) ──────────────────────────────────────────
    app._leyenda_visible = False
    app._leyenda_panel = ctk.CTkFrame(
        tab,
        fg_color=C["panel_bg"],
        corner_radius=10,
        border_width=1,
        border_color=C["panel_brd"])
    _construir_leyenda(app)

    # ── Drop zone ─────────────────────────────────────────────────────────────
    app._drop_zone = _construir_drop_zone(app, tab)

    # ── Banner de completado ──────────────────────────────────────────────────
    app._banner = ctk.CTkLabel(
        tab, text="",
        font=ctk.CTkFont(size=12, weight="bold"),
        text_color="#166534",
        fg_color="#DCFCE7",
        corner_radius=6,
        height=34,
        anchor="center")

    # ── Scroll de bloques ─────────────────────────────────────────────────────
    app.frame_scroll = ctk.CTkScrollableFrame(
        tab,
        fg_color=C["scroll_bg"],
        label_text="",
        corner_radius=10,
        border_width=1,
        border_color=C["panel_brd"],
        scrollbar_button_color="#DDE3ED",
        scrollbar_button_hover_color="#C8D0DF")
    app.frame_scroll.pack(fill="both", expand=True, padx=24, pady=(0, 12))

    # ── Info del documento ────────────────────────────────────────────────────
    _construir_info_doc(app, tab)


def _construir_header(app, parent):
    """Encabezado con título y subtítulo descriptivo."""
    header = ctk.CTkFrame(parent, fg_color="transparent")
    header.pack(fill="x", padx=28, pady=(24, 0))

    # Icono de sección
    icon_frame = ctk.CTkFrame(
        header, width=38, height=38, corner_radius=8,
        fg_color=C["accent_light"])
    icon_frame.pack(side="left", padx=(0, 12))
    icon_frame.pack_propagate(False)
    ctk.CTkLabel(
        icon_frame,
        text="▣",
        font=ctk.CTkFont(size=16),
        text_color=C["accent"]
    ).place(relx=0.5, rely=0.5, anchor="center")

    title_col = ctk.CTkFrame(header, fg_color="transparent")
    title_col.pack(side="left", fill="y")

    ctk.CTkLabel(
        title_col,
        text="Documento PDF",
        font=ctk.CTkFont(size=16, weight="bold"),
        text_color=C["text_main"],
        anchor="w"
    ).pack(anchor="w")

    ctk.CTkLabel(
        title_col,
        text="Carga el archivo PDF del artículo que deseas procesar",
        font=ctk.CTkFont(size=11),
        text_color=C["text_light"],
        anchor="w"
    ).pack(anchor="w")


def _construir_toolbar(app, parent):
    """Toolbar con botones, filtro y buscador."""
    # Contenedor con fondo blanco y borde
    toolbar_outer = ctk.CTkFrame(
        parent,
        fg_color=C["toolbar_bg"],
        corner_radius=10,
        border_width=1,
        border_color=C["toolbar_brd"],
        height=54)
    toolbar_outer.pack(fill="x", padx=24, pady=(14, 8))
    toolbar_outer.pack_propagate(False)

    # Grupo izquierdo
    left = ctk.CTkFrame(toolbar_outer, fg_color="transparent")
    left.pack(side="left", fill="y", padx=(12, 0))

    app._btn_cargar = ctk.CTkButton(
        left,
        text="Seleccionar PDF",
        command=app.evento_cargar_archivo,
        fg_color=C["accent"],
        hover_color=C["accent_hov"],
        text_color="#FFFFFF",
        width=155,
        height=34,
        corner_radius=7,
        font=ctk.CTkFont(size=12, weight="bold"),
        border_width=0)
    app._btn_cargar.pack(side="left", padx=(0, 8), pady=10)

    ctk.CTkButton(
        left,
        text="Leyenda",
        command=app._toggle_leyenda,
        fg_color=C["btn_sec"],
        hover_color=C["btn_sec_hov"],
        text_color=C["text_main"],
        width=76,
        height=34,
        corner_radius=7,
        font=ctk.CTkFont(size=12),
        border_width=1,
        border_color=C["btn_brd"]
    ).pack(side="left", padx=(0, 8), pady=10)

    # Separador
    ctk.CTkFrame(
        left, width=1, height=26, fg_color=C["toolbar_brd"]
    ).pack(side="left", padx=4, pady=14)

    ctk.CTkLabel(
        left,
        text="Filtrar:",
        font=ctk.CTkFont(size=12),
        text_color=C["text_sub"]
    ).pack(side="left", padx=(8, 6), pady=10)

    app._filtro_menu = ctk.CTkOptionMenu(
        left,
        values=["Todos"] + OPCIONES,
        command=app._aplicar_filtro,
        fg_color=C["btn_sec"],
        button_color=C["accent"],
        button_hover_color=C["accent_hov"],
        text_color=C["text_main"],
        dropdown_fg_color="#FFFFFF",
        dropdown_hover_color=C["btn_sec"],
        dropdown_text_color=C["text_main"],
        width=185,
        height=34,
        corner_radius=7,
        font=ctk.CTkFont(size=12))
    app._filtro_menu.set("Todos")
    app._filtro_menu.pack(side="left", padx=(0, 8), pady=10)

    app._stats_lbl = ctk.CTkLabel(
        left,
        text="",
        font=ctk.CTkFont(size=11),
        text_color=C["text_light"])
    app._stats_lbl.pack(side="left", padx=(4, 0))

    # Grupo derecho — buscador
    right = ctk.CTkFrame(toolbar_outer, fg_color="transparent")
    right.pack(side="right", fill="y", padx=(0, 12))

    app._btn_buscar_prev = ctk.CTkButton(
        right,
        text="▲",
        width=28,
        height=34,
        corner_radius=7,
        fg_color=C["btn_sec"],
        hover_color=C["btn_sec_hov"],
        text_color=C["text_main"],
        font=ctk.CTkFont(size=10),
        border_width=1,
        border_color=C["btn_brd"],
        command=lambda: app._navegar_busqueda(-1))
    app._btn_buscar_prev.pack(side="right", padx=(2, 0), pady=10)

    app._btn_buscar_next = ctk.CTkButton(
        right,
        text="▼",
        width=28,
        height=34,
        corner_radius=7,
        fg_color=C["btn_sec"],
        hover_color=C["btn_sec_hov"],
        text_color=C["text_main"],
        font=ctk.CTkFont(size=10),
        border_width=1,
        border_color=C["btn_brd"],
        command=lambda: app._navegar_busqueda(+1))
    app._btn_buscar_next.pack(side="right", padx=(2, 0), pady=10)

    app._lbl_buscar_cnt = ctk.CTkLabel(
        right,
        text="",
        width=52,
        font=ctk.CTkFont(size=11),
        text_color=C["text_sub"])
    app._lbl_buscar_cnt.pack(side="right", padx=(0, 4), pady=10)

    app._entry_buscar = ctk.CTkEntry(
        right,
        placeholder_text="Buscar en bloques...",
        width=195,
        height=34,
        corner_radius=7,
        fg_color="#FFFFFF",
        border_color=C["btn_brd"],
        text_color=C["text_main"],
        placeholder_text_color=C["text_light"],
        font=ctk.CTkFont(size=12))
    app._entry_buscar.pack(side="right", padx=(0, 4), pady=10)
    app._entry_buscar.bind("<Return>",      lambda e: app._navegar_busqueda(+1))
    app._entry_buscar.bind("<Shift-Return>", lambda e: app._navegar_busqueda(-1))
    app._entry_buscar.bind("<KeyRelease>",   lambda e: app._buscar_en_bloques())


def _construir_drop_zone(app, parent):
    """Zona de arrastrar y soltar con diseño limpio y centrado."""
    # Marco exterior con borde punteado visual (simulado con color suave)
    zone = ctk.CTkFrame(
        parent,
        fg_color=C["drop_bg"],
        corner_radius=12,
        border_width=2,
        border_color=C["drop_brd_dash"])
    zone.pack(fill="x", padx=24, pady=(0, 10))

    # Contenido centrado
    inner = ctk.CTkFrame(zone, fg_color="transparent")
    inner.pack(pady=(40, 36))

    # Icono de documento grande
    icon_outer = ctk.CTkFrame(
        inner,
        width=80,
        height=80,
        corner_radius=16,
        fg_color=C["panel_bg"],
        border_width=1,
        border_color=C["panel_brd"])
    icon_outer.pack(anchor="center")
    icon_outer.pack_propagate(False)

    ctk.CTkLabel(
        icon_outer,
        text="▣",
        font=ctk.CTkFont(size=30),
        text_color="#C8D0DF"
    ).place(relx=0.5, rely=0.5, anchor="center")

    # Badge "+" sobre el icono
    plus = ctk.CTkFrame(
        icon_outer,
        width=24,
        height=24,
        corner_radius=12,
        fg_color=C["accent"])
    plus.place(relx=1.0, rely=1.0, anchor="se", x=-4, y=-4)
    plus.pack_propagate(False)
    ctk.CTkLabel(
        plus,
        text="+",
        font=ctk.CTkFont(size=14, weight="bold"),
        text_color="#FFFFFF"
    ).place(relx=0.5, rely=0.5, anchor="center")

    ctk.CTkLabel(
        inner,
        text="Arrastra y suelta tu PDF aquí",
        font=ctk.CTkFont(size=15, weight="bold"),
        text_color=C["text_main"]
    ).pack(pady=(20, 3))

    ctk.CTkLabel(
        inner,
        text="o haz clic para seleccionar un archivo",
        font=ctk.CTkFont(size=12),
        text_color=C["text_sub"]
    ).pack()

    ctk.CTkButton(
        inner,
        text="Seleccionar PDF",
        command=app.evento_cargar_archivo,
        fg_color=C["accent"],
        hover_color=C["accent_hov"],
        text_color="#FFFFFF",
        width=175,
        height=40,
        corner_radius=8,
        font=ctk.CTkFont(size=12, weight="bold"),
        border_width=0
    ).pack(pady=(20, 6))

    ctk.CTkLabel(
        inner,
        text="Formatos soportados: PDF",
        font=ctk.CTkFont(size=11),
        text_color=C["text_light"]
    ).pack()

    return zone


def _construir_info_doc(app, parent):
    """Panel inferior: info del documento + consejos."""
    row = ctk.CTkFrame(parent, fg_color="transparent")
    row.pack(fill="x", padx=24, pady=(0, 10))
    row.grid_columnconfigure(0, weight=3)
    row.grid_columnconfigure(1, weight=2)

    # ── Card Info ─────────────────────────────────────────────────────────────
    info_card = ctk.CTkFrame(
        row,
        fg_color=C["panel_bg"],
        corner_radius=10,
        border_width=1,
        border_color=C["panel_brd"])
    info_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

    # Encabezado con subrayado (acento azul)
    info_header = ctk.CTkFrame(info_card, fg_color="transparent")
    info_header.pack(fill="x", padx=18, pady=(14, 0))

    ctk.CTkLabel(
        info_header,
        text="Información del documento",
        font=ctk.CTkFont(size=12, weight="bold"),
        text_color=C["text_main"],
        anchor="w"
    ).pack(anchor="w")

    # Subrayado acento
    ctk.CTkFrame(
        info_card, height=2, fg_color=C["accent"], corner_radius=1
    ).pack(fill="x", padx=18, pady=(4, 10))

    # Campos en fila
    fields_row = ctk.CTkFrame(info_card, fg_color="transparent")
    fields_row.pack(fill="x", padx=18, pady=(0, 14))

    for label in ["Nombre:", "Páginas:", "Tamaño:", "Estado:"]:
        col = ctk.CTkFrame(fields_row, fg_color="transparent")
        col.pack(side="left", expand=True, anchor="w")

        ctk.CTkLabel(
            col,
            text=label,
            font=ctk.CTkFont(size=10, weight="bold"),
            text_color=C["text_sub"]
        ).pack(anchor="w")

        val_text = "Sin cargar" if label == "Estado:" else "—"
        val_lbl = ctk.CTkLabel(
            col,
            text=val_text,
            font=ctk.CTkFont(size=12),
            text_color=C["text_light"])
        val_lbl.pack(anchor="w", pady=(2, 0))

        # Badge para Estado
        if label == "Estado:":
            badge = ctk.CTkFrame(
                col,
                fg_color="#F1F5F9",
                corner_radius=5)
            badge.pack(anchor="w", pady=(2, 0))
            ctk.CTkLabel(
                badge,
                text="Sin cargar",
                font=ctk.CTkFont(size=10),
                text_color=C["text_sub"]
            ).pack(padx=8, pady=2)
            val_lbl.pack_forget()  # Usar el badge en su lugar

    # ── Card Consejos ─────────────────────────────────────────────────────────
    tips_card = ctk.CTkFrame(
        row,
        fg_color=C["panel_bg"],
        corner_radius=10,
        border_width=1,
        border_color=C["panel_brd"])
    tips_card.grid(row=0, column=1, sticky="nsew", padx=(0, 0))

    tips_header = ctk.CTkFrame(tips_card, fg_color="transparent")
    tips_header.pack(anchor="w", padx=18, pady=(14, 0))

    ctk.CTkLabel(
        tips_header,
        text="💡",
        font=ctk.CTkFont(size=13)
    ).pack(side="left", padx=(0, 6))

    ctk.CTkLabel(
        tips_header,
        text="Consejos",
        font=ctk.CTkFont(size=12, weight="bold"),
        text_color=C["accent"]
    ).pack(side="left")

    # Subrayado acento
    ctk.CTkFrame(
        tips_card, height=2, fg_color=C["accent"], corner_radius=1
    ).pack(fill="x", padx=18, pady=(4, 10))

    tips = [
        "Asegúrate de que el PDF sea texto seleccionable",
        "Evita PDFs escaneados para mejor extracción",
    ]
    for tip in tips:
        tip_row = ctk.CTkFrame(tips_card, fg_color="transparent")
        tip_row.pack(fill="x", padx=18, pady=(0, 8))

        ctk.CTkLabel(
            tip_row,
            text="•",
            font=ctk.CTkFont(size=13),
            text_color=C["accent"],
            width=14
        ).pack(side="left", anchor="n", pady=1)

        ctk.CTkLabel(
            tip_row,
            text=tip,
            font=ctk.CTkFont(size=11),
            text_color=C["text_sub"],
            wraplength=230,
            justify="left"
        ).pack(side="left", padx=(4, 0))


def _construir_leyenda(app):
    """Grid de etiquetas de clasificación."""
    cols = 4
    for i, (cls, color, etiqueta) in enumerate(COLORES_UI):
        row, col = divmod(i, cols)
        celda = ctk.CTkFrame(app._leyenda_panel, fg_color="transparent")
        celda.grid(row=row, column=col, padx=12, pady=5, sticky="w")

        dot = ctk.CTkFrame(celda, width=14, height=14,
                           fg_color=color, corner_radius=4)
        dot.pack(side="left", padx=(0, 6))
        dot.pack_propagate(False)

        ctk.CTkLabel(
            celda,
            text=f"{cls}  ({etiqueta})",
            font=ctk.CTkFont(size=11),
            text_color=C["text_main"]
        ).pack(side="left")