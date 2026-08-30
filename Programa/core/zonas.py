"""
core/zonas.py
Detección de "anclas" (encabezados de sección reconocibles) y asignación de
zona/estado a cada bloque de texto de un artículo científico — enfoque tipo
autómata (SciELO Markup): se recorre el documento en orden y se va llevando
la cuenta de en qué parte del artículo se está.

Por qué este módulo existe
---------------------------
La versión anterior (una variable por tipo de ancla: idx_resumen,
idx_palabras_clave, zona_b_inicio, zona_b_fin) detectaba cada ancla UNA SOLA
VEZ ("primera coincidencia gana") y asumía un orden fijo Resumen → Palabras
clave → Cuerpo → Referencias. Eso rompe con artículos reales que repiten el
ciclo (p. ej. Resumen ES + Abstract EN + Resumen no técnico + Non-technical
Abstract, cada uno con su propio "Palabras clave"/"Keywords"): todo lo que
había entre la PRIMERA "Palabras clave" y la PRIMERA sección numerada del
cuerpo se etiquetaba como "Palabras clave", sin importar que en el medio
hubiera un "Abstract" completo, fechas de manuscrito, etc.

Este módulo en cambio:
  1. Escanea TODAS las ocurrencias de cada ancla reconocida (no solo la
     primera) y arma una lista ordenada de "breakpoints".
  2. Para clasificar el bloque i, busca cuál fue el breakpoint más reciente
     antes de i y usa su tipo — así una zona se cierra automáticamente en
     cuanto aparece la siguiente ancla, sea cual sea, sin importar el orden
     ni cuántas veces se repita el ciclo. Esto también lo vuelve tolerante
     a revistas con otro orden de secciones (p. ej. Agradecimientos antes
     de Referencias en vez de después).

Usado por core/pdf_processor.py y core/docx_processor.py — una sola
implementación de la detección de anclas para ambos formatos.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from core.utils import (
    es_encabezado_resumen,
    es_encabezado_palabras_clave,
    es_inicio_palabras_clave,
    es_encabezado_referencias,
    es_encabezado_seccion_cuerpo,
)

_RE_NIVEL1 = re.compile(r"^\d+\.\s+\S")   # "1. Introducción" (no "1.2 ...")
_RE_NIVEL2 = re.compile(r"^\d+\.\d+")     # "1.2 Subsección"


@dataclass(frozen=True)
class Breakpoint:
    idx: int              # índice del bloque que dispara el ancla
    tipo: str             # "resumen" | "palabras_clave" | "cuerpo" | "referencias"
    seccion: str | None   # p. ej. "introduccion", "conflicto_intereses"...
                           # (hoy solo trazabilidad; útil a futuro para SPS)
    inline: bool = False  # True si el encabezado y el contenido vienen
                           # juntos en el mismo bloque (p. ej. una sola línea
                           # "Palabras clave: brecha de género, ..."), en vez
                           # de encabezado y contenido en bloques separados.


def detectar_breakpoints(
    textos: list[str],
    sizes: list[float] | None = None,
) -> list[Breakpoint]:
    """Recorre los bloques UNA vez y devuelve TODAS las anclas reconocidas,
    en orden de aparición.

    textos: lista de contenido de cada bloque, en el mismo orden en que se
            recorrerá el documento (el índice i es el mismo que se usará
            después para clasificar).
    sizes:  tamaños de letra opcionales (uno por bloque), para exigir que el
            patrón numérico genérico "1. Xxx" tenga un tamaño razonable de
            encabezado (<=12), igual que exigía la versión anterior. Si no
            se pasa, no se aplica ese filtro (útil para DOCX, donde el
            tamaño no siempre es representativo).
    """
    breakpoints: list[Breakpoint] = []

    for i, t in enumerate(textos):
        if not t or not t.strip():
            continue
        texto = t.strip()

        if es_encabezado_resumen(texto):
            breakpoints.append(Breakpoint(i, "resumen", None))
            continue

        if es_encabezado_palabras_clave(texto):
            breakpoints.append(Breakpoint(i, "palabras_clave", None, inline=False))
            continue
        if es_inicio_palabras_clave(texto):
            breakpoints.append(Breakpoint(i, "palabras_clave", None, inline=True))
            continue

        if es_encabezado_referencias(texto):
            breakpoints.append(Breakpoint(i, "referencias", None))
            continue

        seccion = es_encabezado_seccion_cuerpo(texto)
        if seccion:
            breakpoints.append(Breakpoint(i, "cuerpo", seccion))
            continue

        # Fallback genérico: encabezado numerado "1. Xxx" no reconocido por
        # nombre (p. ej. "1. Estructuras de desigualdad de género..."), que
        # también cuenta como ancla de cuerpo si tiene tamaño de encabezado.
        if _RE_NIVEL1.match(texto) and not _RE_NIVEL2.match(texto):
            size_ok = (
                sizes is None
                or i >= len(sizes)
                or sizes[i] is None
                or round(sizes[i]) <= 12
            )
            if size_ok:
                breakpoints.append(Breakpoint(i, "cuerpo", None))
                continue

    return breakpoints


class CursorZonas:
    """Cursor de avance monótono sobre una lista de breakpoints.

    Se usa recorriendo los bloques en orden creciente de índice (0, 1, 2...)
    para saber en qué zona cae cada uno, sin tener que re-escanear toda la
    lista de breakpoints en cada bloque (avanza en O(1) amortizado).
    """

    def __init__(self, breakpoints: list[Breakpoint]):
        self._bps = breakpoints
        self._pos = 0
        self._actual: Breakpoint | None = None

    def avanzar(self, i: int) -> Breakpoint | None:
        """Devuelve el breakpoint más reciente con idx <= i, o None si el
        bloque i todavía no cruzó ninguna ancla (zona "portada": título,
        autores, filiaciones — antes de cualquier Resumen/Abstract)."""
        while self._pos < len(self._bps) and self._bps[self._pos].idx <= i:
            self._actual = self._bps[self._pos]
            self._pos += 1
        return self._actual