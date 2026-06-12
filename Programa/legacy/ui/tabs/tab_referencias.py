"""
ui/tabs/tab_referencias.py
Construcción del tab "📋  Referencias" y su lógica de UI.
"""
import os
import customtkinter as ctk
from tkinter import filedialog
from core.utils import parsear_referencias as _parsear_referencias


# ─────────────────────────────────────────────────────────────────────────────
# Construcción del tab
# ─────────────────────────────────────────────────────────────────────────────

def construir(app):
    tab = app.tabs.tab("📋  Referencias")

    ctk.CTkLabel(
        tab,
        text=(
            "Carga un .txt con las referencias numeradas.\n"
            "Formatos aceptados:   1. Texto...   |   1) Texto...   |   [1] Texto...\n"
        ),
        font=ctk.CTkFont(size=15), justify="left", text_color="#aaa"
    ).pack(anchor="w", padx=12, pady=(12, 6))

    btn_f = ctk.CTkFrame(tab, fg_color="transparent")
    btn_f.pack(fill="x", padx=10, pady=(0, 8))

    ctk.CTkButton(btn_f, text="📂  Cargar .txt de referencias",
                  command=app.evento_cargar_referencias,
                  fg_color="#c62828", hover_color="#8b0000",
                  width=230, font=ctk.CTkFont(size=15)
                  ).pack(side="left", padx=5)
    ctk.CTkButton(btn_f, text="🗑  Limpiar",
                  command=app._limpiar_referencias,
                  fg_color="#c62828", hover_color="#8b0000",
                  width=110, font=ctk.CTkFont(size=15)
                  ).pack(side="left", padx=5)

    app._refs_count_lbl = ctk.CTkLabel(
        btn_f, text="Sin referencias cargadas",
        font=ctk.CTkFont(size=15), text_color="#888")
    app._refs_count_lbl.pack(side="left", padx=10)

    app._refs_scroll = ctk.CTkScrollableFrame(tab, label_text="Referencias cargadas")
    app._refs_scroll.pack(fill="both", expand=True, padx=10, pady=(0, 6))


# ─────────────────────────────────────────────────────────────────────────────
# Lógica de UI
# ─────────────────────────────────────────────────────────────────────────────

def cargar_referencias(app):
    ruta = filedialog.askopenfilename(
        title="Referencias (.txt)",
        filetypes=[("Texto", "*.txt"), ("Todos", "*.*")])
    if not ruta:
        return
    try:
        with open(ruta, encoding="utf-8", errors="replace") as f:
            contenido = f.read()
        app.referencias_externas = _parsear_referencias(contenido)
        refrescar_lista(app)
        n = len(app.referencias_externas)
        app._refs_count_lbl.configure(
            text=f"✓ {n} referencias  ({os.path.basename(ruta)})",
            text_color="#a5d6a7")
        app._set_status(f"✓ {n} referencias cargadas desde '{os.path.basename(ruta)}'")
    except Exception as e:
        app._set_status(f"❌ Error: {e}")


def limpiar_referencias(app):
    app.referencias_externas = []
    refrescar_lista(app)
    app._refs_count_lbl.configure(text="Sin referencias cargadas", text_color="#888")
    app._set_status("Referencias limpiadas.")


def refrescar_lista(app):
    for w in app._refs_scroll.winfo_children():
        w.destroy()
    for i, ref in enumerate(app.referencias_externas, 1):
        frame = ctk.CTkFrame(app._refs_scroll, fg_color="#252525", corner_radius=4)
        frame.pack(fill="x", padx=4, pady=2)
        ctk.CTkLabel(frame, text=f"{i}.", width=32,
                     font=ctk.CTkFont(size=15, weight="bold"),
                     text_color="#aaa").pack(side="left", padx=(6, 2), pady=4)
        ctk.CTkLabel(frame, text=ref[:160] + ("…" if len(ref) > 160 else ""),
                     wraplength=780, justify="left",
                     font=ctk.CTkFont(size=15)).pack(
            side="left", padx=4, pady=4, fill="x", expand=True)