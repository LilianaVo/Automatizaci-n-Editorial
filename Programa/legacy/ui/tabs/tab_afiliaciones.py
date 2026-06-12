"""
ui/tabs/tab_afiliaciones.py
Construcción del tab "🏛️  Afiliaciones" y su lógica de UI.
"""
import re
import customtkinter as ctk
from tkinter import filedialog
from core.utils import split_afiliaciones_linea as _split_afiliaciones_linea


# ─────────────────────────────────────────────────────────────────────────────
# Construcción del tab
# ─────────────────────────────────────────────────────────────────────────────

def construir(app):
    tab = app.tabs.tab("🏛️  Afiliaciones")

    ctk.CTkLabel(
        tab,
        text=(
            "Carga un .txt con las afiliaciones, una por línea:\n"
            "    1 Colección de Paleontología, Facultad...\n"
            "    a Department of Paleontology, Institute...\n"
            "    * correo@unam.mx\n"
            "Acepta prefijos con número o letra (1, 2, a, b, ...).\n"
            "Los correos se vinculan automáticamente."
        ),
        font=ctk.CTkFont(size=15), justify="left", text_color="#aaa"
    ).pack(anchor="w", padx=14, pady=(12, 6))

    bf = ctk.CTkFrame(tab, fg_color="transparent")
    bf.pack(fill="x", padx=10, pady=(0, 6))

    ctk.CTkButton(bf, text="📂  Cargar .txt",
                  command=app._cargar_afiliaciones,
                  fg_color="#1b5e20", hover_color="#2e7d32",
                  width=170, font=ctk.CTkFont(size=15)).pack(side="left", padx=5)
    ctk.CTkButton(bf, text="🗑  Limpiar",
                  command=app._limpiar_afiliaciones,
                  fg_color="#c62828", hover_color="#8b0000",
                  width=110, font=ctk.CTkFont(size=15)).pack(side="left", padx=5)
    app._afil_lbl = ctk.CTkLabel(
        bf, text="Sin afiliaciones cargadas",
        font=ctk.CTkFont(size=15), text_color="#888")
    app._afil_lbl.pack(side="left", padx=10)

    app._afil_scroll = ctk.CTkScrollableFrame(tab, label_text="Afiliaciones cargadas")
    app._afil_scroll.pack(fill="both", expand=True, padx=10, pady=(0, 6))


# ─────────────────────────────────────────────────────────────────────────────
# Lógica de UI
# ─────────────────────────────────────────────────────────────────────────────

def cargar_afiliaciones(app):
    ruta = filedialog.askopenfilename(
        title="Selecciona .txt de afiliaciones",
        filetypes=[("Texto", "*.txt"), ("Todos", "*.*")])
    if not ruta:
        return
    with open(ruta, encoding="utf-8", errors="replace") as f:
        app.afiliaciones_txt = f.read()
    refrescar(app)


def limpiar_afiliaciones(app):
    app.afiliaciones_txt = ""
    for w in app._afil_scroll.winfo_children():
        w.destroy()
    app._afil_lbl.configure(text="Sin afiliaciones cargadas", text_color="#888")


def refrescar(app):
    for w in app._afil_scroll.winfo_children():
        w.destroy()

    lineas_raw = [l for l in app.afiliaciones_txt.splitlines() if l.strip()]
    lineas = []
    for linea in lineas_raw:
        t = linea.strip()
        if re.match(r"^\*\s*[\w\.\-]+@[\w\-\.]+\.\w{2,}", t):
            lineas.append(t)
            continue
        segs = _split_afiliaciones_linea(t)
        if segs:
            for lbl, txt in segs:
                lineas.append(f"{lbl} {txt}")
        else:
            lineas.append(t)

    for linea in lineas:
        frame = ctk.CTkFrame(app._afil_scroll, fg_color="#1a2a1a", corner_radius=4)
        frame.pack(fill="x", padx=2, pady=2)
        ctk.CTkLabel(frame, text=linea, font=ctk.CTkFont(size=15),
                     justify="left", anchor="w", wraplength=700).pack(
            padx=10, pady=5, fill="x")

    n = len(lineas)
    app._afil_lbl.configure(
        text=f"{n} afiliación{'es' if n != 1 else ''}" if n else "Sin afiliaciones",
        text_color="#a5d6a7" if n else "#888")