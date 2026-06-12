"""
ui/widgets/bloque_widget.py
Widget de bloque editable (v3) — tema claro, consistente con app_window v3.

Cambios visuales:
  • Colores de fondo claros en lugar de oscuros
  • Badge de tamaño con pill suave
  • Bordes más delicados
  • Dropdown de clase con estilo claro
"""
import customtkinter as ctk
from core.constans import OPCIONES, COLOR_POR_CLASE, ESTILO_POR_CLASE

# ── Paleta general para el widget ────────────────────────────────────────────
C = {
    "accent":     "#1B2A4A",
    "accent_hov": "#243860",
    "btn_brd":    "#DDE3ED",
    "text_sub":   "#5A6478",
    "text_light": "#9AA3B5",
}


def crear_bloque_ui(app, item: dict):
    """Crea un frame de bloque editable y lo registra en app.datos_bloques."""
    cls  = item["clasificacion"]
    cont = item["contenido"]
    size = item.get("size", 10)
    bold = item.get("bold", False)
    ital = item.get("italic", False)

    color = COLOR_POR_CLASE.get(cls, "#F1F5F9")

    # Frame contenedor del bloque
    frame = ctk.CTkFrame(
        app.frame_scroll,
        fg_color=color,
        corner_radius=8,
        border_width=1,
        border_color=_lighten_border(color))
    frame.pack(fill="x", padx=8, pady=3)
    frame.columnconfigure(0, weight=0)
    frame.columnconfigure(1, weight=1)
    frame.columnconfigure(2, weight=0)

    # ── Badge de tamaño/estilo ────────────────────────────────────────────────
    badge_parts = [f"{size:.0f}pt"]
    if bold: badge_parts.append("B")
    if ital: badge_parts.append("I")
    badge_text = "  ".join(badge_parts)

    badge_pill = ctk.CTkFrame(
        frame,
        fg_color="#FFFFFF",
        corner_radius=5,
        border_width=1,
        border_color=C["btn_brd"])
    badge_pill.grid(row=0, column=0, padx=(8, 4), pady=(8, 8), sticky="nw")

    ctk.CTkLabel(
        badge_pill,
        text=badge_text,
        font=ctk.CTkFont(size=10),
        text_color=C["text_sub"]
    ).pack(padx=6, pady=2)

    # ── Textbox ───────────────────────────────────────────────────────────────
    lineas_aprox = max(2, min(10, len(cont) // 80 + cont.count("\n") + 1))
    altura_px    = lineas_aprox * 20 + 14

    fsize, weight, slant, fg, bg = ESTILO_POR_CLASE.get(
        cls, (12, "normal", "roman", "#1E293B", "#FFFFFF"))

    txt_box = ctk.CTkTextbox(
        frame,
        font=ctk.CTkFont(size=fsize, weight=weight),
        fg_color=bg,
        text_color=fg,
        border_color=C["btn_brd"],
        border_width=1,
        corner_radius=6,
        wrap="word",
        height=altura_px,
        activate_scrollbars=False)
    txt_box.insert("1.0", cont)
    txt_box.grid(row=0, column=1, padx=(4, 6), pady=(8, 8), sticky="ew")

    # ── Menú de clasificación ─────────────────────────────────────────────────
    menu = ctk.CTkOptionMenu(
        frame,
        values=OPCIONES,
        width=168,
        height=32,
        corner_radius=6,
        fg_color="#FFFFFF",
        button_color=C["accent"],
        button_hover_color=C["accent_hov"],
        text_color=C["text_sub"],
        dropdown_fg_color="#FFFFFF",
        dropdown_hover_color="#F4F6FA",
        dropdown_text_color="#1A2236",
        font=ctk.CTkFont(size=11),
        command=lambda v, f=frame, tb=txt_box: _on_clase_cambiada(app, v, f, tb))
    menu.set(cls)
    menu.grid(row=0, column=2, padx=(0, 8), pady=(8, 8), sticky="ne")

    app.datos_bloques.append({
        "contenido":       cont,
        "menu":            menu,
        "italic":          ital,
        "frame":           frame,
        "_txtbox":         txt_box,
        "_color_original": color,
    })


def _lighten_border(hex_color: str) -> str:
    """Devuelve un color de borde ligeramente más oscuro que el fondo del bloque."""
    try:
        r = int(hex_color[1:3], 16)
        g = int(hex_color[3:5], 16)
        b = int(hex_color[5:7], 16)
        factor = 0.88
        r2 = max(0, int(r * factor))
        g2 = max(0, int(g * factor))
        b2 = max(0, int(b * factor))
        return f"#{r2:02x}{g2:02x}{b2:02x}"
    except Exception:
        return "#DDE3ED"


def _on_clase_cambiada(app, nueva_clase: str, frame, txt_box):
    """Cambia color del frame y estilo del textbox al reclasificar."""
    nuevo_color = COLOR_POR_CLASE.get(nueva_clase, "#F1F5F9")
    frame.configure(
        fg_color=nuevo_color,
        border_color=_lighten_border(nuevo_color))
    _aplicar_estilo_textbox(txt_box, nueva_clase)
    app._actualizar_stats(nueva_clase)
    for b in app.datos_bloques:
        if b["frame"] is frame:
            b["_color_original"] = nuevo_color
            break


def _aplicar_estilo_textbox(txt_box, cls: str):
    """Aplica fuente y colores al CTkTextbox según la clase semántica."""
    fsize, weight, slant, fg, bg = ESTILO_POR_CLASE.get(
        cls, (12, "normal", "roman", "#1E293B", "#FFFFFF"))
    txt_box.configure(
        font=ctk.CTkFont(size=fsize, weight=weight),
        text_color=fg,
        fg_color=bg,
        border_color=C["btn_brd"])