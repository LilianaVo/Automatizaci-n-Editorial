"""
ui/tabs/tab_figuras.py
Construcción del tab "🖼️  Figuras" (panel Figuras + panel Tablas)
y toda su lógica de UI: agregar, borrar, sync, refresh, limpiar.
"""
import os
import re
import shutil
import tempfile
import customtkinter as ctk
from tkinter import filedialog
from PIL import Image as PILImage
import fitz

from core.utils import (
    limpiar_prefijo_pie_figura   as _limpiar_prefijo_pie_figura,
    limpiar_prefijo_titulo_tabla as _limpiar_prefijo_titulo_tabla,
)


# ─────────────────────────────────────────────────────────────────────────────
# Construcción del tab
# ─────────────────────────────────────────────────────────────────────────────

def construir(app):
    tab = app.tabs.tab("🖼️  Figuras")

    seg = ctk.CTkSegmentedButton(
        tab,
        values=["🖼️  Figuras", "📊  Tablas"],
        command=app._cambiar_panel_media,
        font=ctk.CTkFont(size=15), width=280)
    seg.set("🖼️  Figuras")
    seg.pack(pady=(10, 6))

    # ── Panel Figuras ─────────────────────────────────────────────────────────
    app._panel_figs = ctk.CTkFrame(tab, fg_color="transparent")
    app._panel_figs.pack(fill="both", expand=True)

    ctk.CTkLabel(
        app._panel_figs,
        text="Agrega imágenes con pie de figura. Numeración automática (Figura 1, 2…).\n",
        font=ctk.CTkFont(size=15), justify="left", text_color="#aaa"
    ).pack(anchor="w", padx=12, pady=(4, 6))

    bf = ctk.CTkFrame(app._panel_figs, fg_color="transparent")
    bf.pack(fill="x", padx=10, pady=(0, 6))
    ctk.CTkButton(bf, text="➕  Agregar figura",
                  command=app._agregar_figura,
                  fg_color="#1565c0", hover_color="#0d47a1",
                  width=170, font=ctk.CTkFont(size=15)).pack(side="left", padx=5)
    ctk.CTkButton(bf, text="🗑  Limpiar",
                  command=app._limpiar_figuras,
                  fg_color="#c62828", hover_color="#8b0000",
                  width=110, font=ctk.CTkFont(size=15)).pack(side="left", padx=5)
    app._figs_count_lbl = ctk.CTkLabel(bf, text="Sin figuras",
                                        font=ctk.CTkFont(size=15), text_color="#888")
    app._figs_count_lbl.pack(side="left", padx=10)

    app._figs_scroll = ctk.CTkScrollableFrame(app._panel_figs, label_text="Figuras cargadas")
    app._figs_scroll.pack(fill="both", expand=True, padx=10, pady=(0, 6))

    # ── Panel Tablas (oculto) ─────────────────────────────────────────────────
    app._panel_tabs = ctk.CTkFrame(tab, fg_color="transparent")

    ctk.CTkLabel(
        app._panel_tabs,
        text="Importa archivos Excel (.xlsx). Numeración automática (Tabla 1, 2…).\n"
             "Escribe el pie de la tabla y el texto del cuerpo correspondiente a insertar.",
        font=ctk.CTkFont(size=15), justify="left", text_color="#aaa"
    ).pack(anchor="w", padx=12, pady=(4, 6))

    bt = ctk.CTkFrame(app._panel_tabs, fg_color="transparent")
    bt.pack(fill="x", padx=10, pady=(0, 6))
    ctk.CTkButton(bt, text="➕  Agregar tabla (.xlsx)",
                  command=app._agregar_tabla,
                  fg_color="#1565c0", hover_color="#0d47a1",
                  width=190, font=ctk.CTkFont(size=15)).pack(side="left", padx=5)
    ctk.CTkButton(bt, text="🗑  Limpiar",
                  command=app._limpiar_tablas,
                  fg_color="#c62828", hover_color="#8b0000",
                  width=110, font=ctk.CTkFont(size=15)).pack(side="left", padx=5)
    app._tabs_count_lbl = ctk.CTkLabel(bt, text="Sin tablas",
                                        font=ctk.CTkFont(size=15), text_color="#888")
    app._tabs_count_lbl.pack(side="left", padx=10)

    app._tabs_scroll = ctk.CTkScrollableFrame(app._panel_tabs, label_text="Tablas cargadas")
    app._tabs_scroll.pack(fill="both", expand=True, padx=10, pady=(0, 6))


# ─────────────────────────────────────────────────────────────────────────────
# Cambio de panel Figuras ↔ Tablas
# ─────────────────────────────────────────────────────────────────────────────

def cambiar_panel_media(app, valor: str):
    if valor == "🖼️  Figuras":
        app._panel_tabs.pack_forget()
        app._panel_figs.pack(fill="both", expand=True)
    else:
        app._panel_figs.pack_forget()
        app._panel_tabs.pack(fill="both", expand=True)


# ─────────────────────────────────────────────────────────────────────────────
# Figuras — lógica
# ─────────────────────────────────────────────────────────────────────────────

def limpiar_cache_figuras_auto(app):
    if app._auto_fig_dir and os.path.isdir(app._auto_fig_dir):
        try:
            shutil.rmtree(app._auto_fig_dir)
        except Exception:
            pass
    app._auto_fig_dir = None


def agregar_figura(app):
    sync_pies(app)
    rutas = filedialog.askopenfilenames(
        title="Selecciona imagen(es)",
        filetypes=[
            ("Imágenes", "*.jpg *.jpeg *.png *.gif *.webp *.bmp *.tiff"),
            ("Todos", "*.*"),
        ])
    if not rutas:
        return
    for ruta in rutas:
        app.figuras_manuales.append({"ruta": ruta, "pie": "", "ancla": ""})
    refrescar_lista_figuras(app)
    app._set_status(f"✓ {len(app.figuras_manuales)} figura(s) en total.")


def limpiar_figuras(app):
    limpiar_cache_figuras_auto(app)
    app.figuras_manuales = []
    refrescar_lista_figuras(app)
    app._figs_count_lbl.configure(text="Sin figuras", text_color="#888")
    app._set_status("Figuras limpiadas.")


def refrescar_lista_figuras(app):
    for w in app._figs_scroll.winfo_children():
        w.destroy()

    for i, fig in enumerate(app.figuras_manuales):
        num   = i + 1
        frame = ctk.CTkFrame(app._figs_scroll, fg_color="#1e2a1e", corner_radius=6)
        frame.pack(fill="x", padx=4, pady=6)
        frame.columnconfigure(2, weight=1)

        try:
            img_pil = PILImage.open(fig["ruta"])
            img_pil.thumbnail((72, 72))
            thumb = ctk.CTkImage(img_pil, size=img_pil.size)
            ctk.CTkLabel(frame, image=thumb, text="").grid(
                row=0, column=0, rowspan=3, padx=(8, 6), pady=8)
        except Exception:
            ctk.CTkLabel(frame, text="🖼️", font=ctk.CTkFont(size=28),
                         width=72).grid(row=0, column=0, rowspan=3,
                                        padx=(8, 6), pady=8)

        ctk.CTkLabel(frame, text=f"Figura {num}",
                     font=ctk.CTkFont(size=15, weight="bold"),
                     text_color="#a5d6a7").grid(
            row=0, column=1, padx=(0, 8), pady=(8, 0), sticky="w")

        subt = os.path.basename(fig["ruta"])
        if fig.get("origen") == "auto_pdf" and fig.get("pagina"):
            subt = f'{subt} - pagina {fig["pagina"]} (auto)'
        ctk.CTkLabel(frame, text=subt,
                     font=ctk.CTkFont(size=15), text_color="#666").grid(
            row=1, column=1, padx=(0, 8), sticky="w")

        pie_var = ctk.StringVar(value=fig.get("pie", ""))
        ctk.CTkEntry(frame,
                     placeholder_text=f"Pie de la Figura {num}…",
                     textvariable=pie_var,
                     font=ctk.CTkFont(size=15), height=36).grid(
            row=0, column=2, columnspan=2,
            padx=(0, 8), pady=(8, 2), sticky="ew")
        fig["_var"] = pie_var

        ctk.CTkLabel(frame,
                     text="📍 Pega aquí el texto del párrafo donde va la figura:",
                     font=ctk.CTkFont(size=15), text_color="#a5d6a7").grid(
            row=2, column=1, columnspan=3, padx=(0, 8), pady=(4, 0), sticky="w")

        anc_var = ctk.StringVar(value=fig.get("ancla", ""))
        ctk.CTkEntry(frame,
                     placeholder_text='Ej: "Se muestra en la Figura 1A-B."',
                     textvariable=anc_var,
                     font=ctk.CTkFont(size=15), height=36).grid(
            row=3, column=1, columnspan=3,
            padx=(0, 8), pady=(0, 8), sticky="ew")
        fig["_var_anc"] = anc_var

        def _borrar(idx=i):
            sync_pies(app)
            app.figuras_manuales.pop(idx)
            refrescar_lista_figuras(app)
            app._set_status(f"Figura {idx+1} eliminada.")

        ctk.CTkButton(frame, text="✕", width=28, height=34,
                      fg_color="#c62828", hover_color="#8b0000",
                      command=_borrar,
                      font=ctk.CTkFont(size=15)).grid(
            row=0, column=4, padx=(0, 6), pady=8)

    n = len(app.figuras_manuales)
    app._figs_count_lbl.configure(
        text=f"{n} figura{'s' if n != 1 else ''}" if n else "Sin figuras",
        text_color="#a5d6a7" if n else "#888")


def sync_pies(app):
    for fig in app.figuras_manuales:
        if "_var" in fig:
            fig["pie"] = fig["_var"].get()
        if "_var_anc" in fig:
            fig["ancla"] = fig["_var_anc"].get()


# ─────────────────────────────────────────────────────────────────────────────
# Tablas — lógica
# ─────────────────────────────────────────────────────────────────────────────

def limpiar_cache_tablas_auto(app):
    if app._auto_tab_dir and os.path.isdir(app._auto_tab_dir):
        try:
            shutil.rmtree(app._auto_tab_dir)
        except Exception:
            pass
    app._auto_tab_dir = None


def remover_tablas_auto(app):
    app.tablas_manuales = [
        t for t in app.tablas_manuales if t.get("origen") != "auto_pdf"
    ]


def agregar_tabla(app):
    sync_titulos_tablas(app)
    rutas = filedialog.askopenfilenames(
        title="Selecciona archivo(s) Excel",
        filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
    if not rutas:
        return
    try:
        import openpyxl
    except ImportError:
        app._set_status("❌ Instala openpyxl: pip install openpyxl")
        return

    for ruta in rutas:
        try:
            wb = openpyxl.load_workbook(ruta, data_only=True)
            hojas = wb.sheetnames
            wb.close()
            for hoja in hojas:
                app.tablas_manuales.append({
                    "ruta":   ruta,
                    "hoja":   hoja,
                    "titulo": "",
                    "ancla":  ""
                })
        except Exception as e:
            app._set_status(f"❌ Error leyendo {os.path.basename(ruta)}: {e}")
            return

    refrescar_lista_tablas(app)
    app._set_status(f"✓ {len(app.tablas_manuales)} tabla(s) detectadas.")


def limpiar_tablas(app):
    limpiar_cache_tablas_auto(app)
    app.tablas_manuales = []
    refrescar_lista_tablas(app)
    app._tabs_count_lbl.configure(text="Sin tablas", text_color="#888")
    app._set_status("Tablas limpiadas.")


def refrescar_lista_tablas(app):
    for w in app._tabs_scroll.winfo_children():
        w.destroy()

    for i, tab_item in enumerate(app.tablas_manuales):
        num   = i + 1
        frame = ctk.CTkFrame(app._tabs_scroll, fg_color="#1a1a2e", corner_radius=6)
        frame.pack(fill="x", padx=4, pady=6)
        frame.columnconfigure(2, weight=1)

        ctk.CTkLabel(frame, text="📊",
                     font=ctk.CTkFont(size=26), width=48).grid(
            row=0, column=0, rowspan=3, padx=(8, 4), pady=8, sticky="n")

        ctk.CTkLabel(frame, text=f"Tabla {num}",
                     font=ctk.CTkFont(size=15, weight="bold"),
                     text_color="#ce93d8").grid(
            row=0, column=1, padx=(0, 6), pady=(8, 0), sticky="w")

        hoja     = tab_item.get("hoja", "")
        archivo  = os.path.basename(tab_item["ruta"])
        subtitulo = f"{archivo}  >  {hoja}" if hoja else archivo
        if tab_item.get("origen") == "auto_pdf" and tab_item.get("pagina"):
            subtitulo += f'  -  pagina {tab_item["pagina"]} (auto)'
        ctk.CTkLabel(frame, text=subtitulo,
                     font=ctk.CTkFont(size=15), text_color="#888").grid(
            row=1, column=1, padx=(0, 6), pady=0, sticky="w")

        tit_var = ctk.StringVar(value=tab_item.get("titulo", ""))
        ctk.CTkEntry(
            frame,
            placeholder_text=f"Título de la Tabla {num}…",
            textvariable=tit_var,
            font=ctk.CTkFont(size=15), height=36
        ).grid(row=0, column=2, columnspan=2,
               padx=(0, 8), pady=(8, 2), sticky="ew")
        tab_item["_var_tit"] = tit_var

        ctk.CTkLabel(frame,
                     text="📍 Pega aquí el texto del párrafo donde va la tabla:",
                     font=ctk.CTkFont(size=15), text_color="#ce93d8").grid(
            row=2, column=1, columnspan=3, padx=(0, 8), pady=(4, 0), sticky="w")

        anc_var = ctk.StringVar(value=tab_item.get("ancla", ""))
        ctk.CTkEntry(
            frame,
            placeholder_text='Ej: "la Dra. Elena Centeno (Tabla 1)."',
            textvariable=anc_var,
            font=ctk.CTkFont(size=15), height=36
        ).grid(row=3, column=1, columnspan=3,
               padx=(0, 8), pady=(0, 8), sticky="ew")
        tab_item["_var_anc"] = anc_var

        def _borrar_t(idx=i):
            sync_titulos_tablas(app)
            app.tablas_manuales.pop(idx)
            refrescar_lista_tablas(app)
            app._set_status(f"Tabla {idx+1} eliminada.")

        ctk.CTkButton(frame, text="✕", width=28, height=34,
                      fg_color="#c62828", hover_color="#8b0000",
                      command=_borrar_t,
                      font=ctk.CTkFont(size=15)).grid(
            row=0, column=4, padx=(0, 6), pady=8)

    n = len(app.tablas_manuales)
    app._tabs_count_lbl.configure(
        text=f"{n} tabla{'s' if n != 1 else ''}" if n else "Sin tablas",
        text_color="#ce93d8" if n else "#888")


def sync_titulos_tablas(app):
    for tab in app.tablas_manuales:
        if "_var_tit" in tab:
            tab["titulo"] = tab["_var_tit"].get()
        if "_var_anc" in tab:
            tab["ancla"]  = tab["_var_anc"].get()