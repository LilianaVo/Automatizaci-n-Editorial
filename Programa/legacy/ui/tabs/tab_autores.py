"""
ui/tabs/tab_autores.py
Panel "Autores" rediseñado (v3) — tema claro, consistente con app_window v3.
"""
import re
import customtkinter as ctk
from tkinter import filedialog

# ── Paleta (misma que app_window v3) ─────────────────────────────────────────
C = {
    "bg":          "#F7F8FC",
    "panel_bg":    "#FFFFFF",
    "panel_brd":   "#E4E9F0",
    "accent":      "#1B2A4A",
    "accent_hov":  "#243860",
    "accent_light":"#EEF2FF",
    "text_main":   "#1A2236",
    "text_sub":    "#5A6478",
    "text_light":  "#9AA3B5",
    "btn_sec":     "#F4F6FA",
    "btn_sec_hov": "#E9EDF5",
    "btn_brd":     "#DDE3ED",
    "danger":      "#EF4444",
    "danger_hov":  "#DC2626",
    "success":     "#16A34A",
    "success_hov": "#15803D",
    "row_bg":      "#F8FAFC",
    "row_brd":     "#E4E9F0",
    "row_num":     "#EEF2FF",
}


# ─────────────────────────────────────────────────────────────────────────────
# Construcción del tab
# ─────────────────────────────────────────────────────────────────────────────

def construir(app):
    tab = app.tabs.tab("👥  Autores")
    tab.configure(fg_color=C["bg"])

    # ── Encabezado ────────────────────────────────────────────────────────────
    header = ctk.CTkFrame(tab, fg_color="transparent")
    header.pack(fill="x", padx=28, pady=(24, 0))

    icon_frame = ctk.CTkFrame(
        header, width=38, height=38, corner_radius=8,
        fg_color=C["accent_light"])
    icon_frame.pack(side="left", padx=(0, 12))
    icon_frame.pack_propagate(False)
    ctk.CTkLabel(
        icon_frame, text="◉",
        font=ctk.CTkFont(size=16),
        text_color=C["accent"]
    ).place(relx=0.5, rely=0.5, anchor="center")

    title_col = ctk.CTkFrame(header, fg_color="transparent")
    title_col.pack(side="left", fill="y")
    ctk.CTkLabel(
        title_col, text="Autores",
        font=ctk.CTkFont(size=16, weight="bold"),
        text_color=C["text_main"], anchor="w"
    ).pack(anchor="w")
    ctk.CTkLabel(
        title_col,
        text="Agrega autores con su ORCID o importa desde Excel",
        font=ctk.CTkFont(size=11),
        text_color=C["text_light"], anchor="w"
    ).pack(anchor="w")

    # ── Barra de acciones ─────────────────────────────────────────────────────
    actions = ctk.CTkFrame(
        tab, fg_color=C["panel_bg"],
        corner_radius=10,
        border_width=1,
        border_color=C["panel_brd"],
        height=54)
    actions.pack(fill="x", padx=24, pady=(14, 8))
    actions.pack_propagate(False)

    ctk.CTkButton(
        actions, text="+ Agregar autor",
        command=app._agregar_autor,
        fg_color=C["accent"], hover_color=C["accent_hov"],
        text_color="#FFFFFF",
        width=145, height=34, corner_radius=7,
        font=ctk.CTkFont(size=12, weight="bold"),
        border_width=0
    ).pack(side="left", padx=(12, 6), pady=10)

    ctk.CTkButton(
        actions, text="Cargar Excel",
        command=app._cargar_autores_excel,
        fg_color=C["btn_sec"], hover_color=C["btn_sec_hov"],
        text_color=C["text_main"],
        width=120, height=34, corner_radius=7,
        font=ctk.CTkFont(size=12),
        border_width=1, border_color=C["btn_brd"]
    ).pack(side="left", padx=(0, 6), pady=10)

    ctk.CTkButton(
        actions, text="Limpiar todo",
        command=app._limpiar_autores,
        fg_color=C["btn_sec"], hover_color="#FEE2E2",
        text_color=C["danger"],
        width=115, height=34, corner_radius=7,
        font=ctk.CTkFont(size=12),
        border_width=1, border_color="#FECACA"
    ).pack(side="left", padx=(0, 8), pady=10)

    # Separador
    ctk.CTkFrame(
        actions, width=1, height=26, fg_color=C["panel_brd"]
    ).pack(side="left", padx=4, pady=14)

    app._autores_lbl = ctk.CTkLabel(
        actions, text="Sin autores cargados",
        font=ctk.CTkFont(size=12),
        text_color=C["text_light"])
    app._autores_lbl.pack(side="left", padx=12)

    # ── Encabezado de columnas ────────────────────────────────────────────────
    hdr = ctk.CTkFrame(
        tab, fg_color=C["btn_sec"],
        corner_radius=0,
        border_width=1,
        border_color=C["panel_brd"],
        height=36)
    hdr.pack(fill="x", padx=24, pady=0)
    hdr.pack_propagate(False)
    hdr.columnconfigure(1, weight=2)
    hdr.columnconfigure(2, weight=1)

    ctk.CTkLabel(
        hdr, text="#", width=40,
        font=ctk.CTkFont(size=11, weight="bold"),
        text_color=C["text_sub"]
    ).grid(row=0, column=0, padx=(12, 0), pady=8)

    ctk.CTkLabel(
        hdr, text="Apellido, Nombre",
        font=ctk.CTkFont(size=11, weight="bold"),
        text_color=C["text_sub"], anchor="w"
    ).grid(row=0, column=1, padx=8, pady=8, sticky="ew")

    ctk.CTkLabel(
        hdr, text="ORCID  (0000-0001-2345-6789)",
        font=ctk.CTkFont(size=11, weight="bold"),
        text_color=C["text_sub"], anchor="w"
    ).grid(row=0, column=2, padx=8, pady=8, sticky="ew")

    ctk.CTkLabel(
        hdr, text="", width=40
    ).grid(row=0, column=3, padx=8)

    # ── Lista scrollable ──────────────────────────────────────────────────────
    app._autores_scroll = ctk.CTkScrollableFrame(
        tab,
        fg_color=C["panel_bg"],
        corner_radius=0,
        border_width=1,
        border_color=C["panel_brd"],
        scrollbar_button_color=C["btn_brd"],
        scrollbar_button_hover_color="#C8D0DF")
    app._autores_scroll.pack(fill="both", expand=True, padx=24, pady=(0, 12))

    # ── Hint de formato ───────────────────────────────────────────────────────
    hint = ctk.CTkFrame(tab, fg_color="transparent")
    hint.pack(fill="x", padx=24, pady=(0, 8))
    ctk.CTkLabel(
        hint,
        text="El ORCID puede ser el link completo (https://orcid.org/0000-...) o solo los números.",
        font=ctk.CTkFont(size=10),
        text_color=C["text_light"],
        anchor="w"
    ).pack(anchor="w")


# ─────────────────────────────────────────────────────────────────────────────
# Lógica de UI
# ─────────────────────────────────────────────────────────────────────────────

def agregar_autor(app):
    sync_autores(app)
    app.autores_orcid.append({"nombre": "", "orcid": ""})
    refrescar_lista(app)


def sync_autores(app):
    for a in app.autores_orcid:
        if "_var_nom" in a:
            a["nombre"] = a["_var_nom"].get().strip()
        if "_var_orc" in a:
            raw = a["_var_orc"].get().strip()
            m = re.search(r"(\d{4}-\d{4}-\d{4}-\d{3}[\dX])", raw, re.IGNORECASE)
            a["orcid"] = m.group(1) if m else raw


def refrescar_lista(app):
    for w in app._autores_scroll.winfo_children():
        w.destroy()

    for i, autor in enumerate(app.autores_orcid):
        # Fila alternada
        row_color = C["panel_bg"] if i % 2 == 0 else C["row_bg"]
        row = ctk.CTkFrame(
            app._autores_scroll,
            fg_color=row_color,
            corner_radius=0,
            border_width=1,
            border_color=C["row_brd"])
        row.pack(fill="x", padx=0, pady=(0, 1))
        row.columnconfigure(1, weight=2)
        row.columnconfigure(2, weight=1)

        # Número de fila
        num_bg = ctk.CTkFrame(
            row, width=40, height=36,
            fg_color=C["row_num"],
            corner_radius=0)
        num_bg.grid(row=0, column=0, padx=0, pady=0, sticky="nsew")
        num_bg.pack_propagate(False)
        ctk.CTkLabel(
            num_bg,
            text=str(i + 1),
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=C["accent"]
        ).place(relx=0.5, rely=0.5, anchor="center")

        # Entrada nombre
        var_nom = ctk.StringVar(value=autor.get("nombre", ""))
        ctk.CTkEntry(
            row,
            textvariable=var_nom,
            placeholder_text="Apellido, Nombre",
            font=ctk.CTkFont(size=12),
            height=36,
            fg_color="#FFFFFF",
            border_color=C["btn_brd"],
            text_color=C["text_main"],
            placeholder_text_color=C["text_light"],
            corner_radius=6
        ).grid(row=0, column=1, padx=(8, 6), pady=6, sticky="ew")
        autor["_var_nom"] = var_nom

        # Entrada ORCID
        var_orc = ctk.StringVar(value=autor.get("orcid", ""))
        ctk.CTkEntry(
            row,
            textvariable=var_orc,
            placeholder_text="0000-0001-2345-6789",
            font=ctk.CTkFont(size=12),
            height=36,
            fg_color="#FFFFFF",
            border_color=C["btn_brd"],
            text_color=C["text_main"],
            placeholder_text_color=C["text_light"],
            corner_radius=6
        ).grid(row=0, column=2, padx=(0, 6), pady=6, sticky="ew")
        autor["_var_orc"] = var_orc

        # Botón eliminar
        def _borrar(idx=i):
            sync_autores(app)
            app.autores_orcid.pop(idx)
            refrescar_lista(app)

        ctk.CTkButton(
            row,
            text="✕",
            width=30,
            height=30,
            corner_radius=6,
            fg_color=C["btn_sec"],
            hover_color="#FEE2E2",
            text_color=C["danger"],
            font=ctk.CTkFont(size=11),
            border_width=1,
            border_color="#FECACA",
            command=_borrar
        ).grid(row=0, column=3, padx=(0, 10), pady=6)

    n = len(app.autores_orcid)
    app._autores_lbl.configure(
        text=f"{n} autor{'es' if n != 1 else ''} cargados" if n else "Sin autores cargados",
        text_color=C["accent"] if n else C["text_light"])


def limpiar_autores(app):
    app.autores_orcid = []
    refrescar_lista(app)


def cargar_autores_excel(app):
    ruta = filedialog.askopenfilename(
        title="Selecciona Excel de autores",
        filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
    if not ruta:
        return
    try:
        import openpyxl
        wb = openpyxl.load_workbook(ruta, data_only=True)
        ws = wb.active
        filas = list(ws.iter_rows(values_only=True))
        wb.close()
    except ImportError:
        app._set_status("Instala openpyxl: pip install openpyxl")
        return
    except Exception as e:
        app._set_status(f"Error leyendo Excel: {e}")
        return

    sync_autores(app)
    nuevos = 0
    for fila in filas:
        if not fila or all(c is None for c in fila):
            continue
        nombre = str(fila[0]).strip() if fila[0] else ""
        orcid  = str(fila[1]).strip() if len(fila) > 1 and fila[1] else ""
        if nombre.lower() in ("autor", "nombre", "author", "name"):
            continue
        if not nombre:
            continue
        m = re.search(r"(\d{4}-\d{4}-\d{4}-\d{3}[\dX])", orcid)
        orcid_limpio = m.group(1) if m else orcid
        app.autores_orcid.append({"nombre": nombre, "orcid": orcid_limpio})
        nuevos += 1

    refrescar_lista(app)
    app._set_status(f"{nuevos} autor(es) importados desde Excel.")