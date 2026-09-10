"""core/sps_documento.py
Modelo de documento (árbol) para la marcación SciELO Markup, con parser y
serializador del formato de corchetes, y validación de anidación contra la
biblioteca de etiquetas (core/sps_tags).

Este es el núcleo compartido de las vistas: la Vista de Etiquetas (Markup plano)
y la Vista de Bloques serán proyecciones de este mismo árbol. Aquí sólo va el
modelo + serialización de corchetes; los puentes a «bloques» y a JATS se agregan
en fases siguientes.

Contenido mixto: un nodo puede contener texto y otros nodos intercalados
(necesario para inline dentro de un párrafo). Por eso `hijos` es una lista de
`str` (texto) o `Nodo`.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from core import sps_tags

# ─────────────────────────────────────────────────────────────────────────────
# Modelo
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class Nodo:
    tag: str                                   # nombre de corchete (p, sec, xref…)
    attrs: dict[str, str] = field(default_factory=dict)
    hijos: list = field(default_factory=list)  # list[str | Nodo]  (texto o nodos)

    # --- utilidades de árbol ---
    def hijos_nodo(self):
        return [h for h in self.hijos if isinstance(h, Nodo)]

    def texto_plano(self) -> str:
        out = []
        for h in self.hijos:
            out.append(h if isinstance(h, str) else h.texto_plano())
        return "".join(out)

    def __eq__(self, otro) -> bool:
        if not isinstance(otro, Nodo):
            return NotImplemented
        return (self.tag == otro.tag
                and self.attrs == otro.attrs
                and self.hijos == otro.hijos)


ROOT_TAG = "#RAIZ"   # nodo raíz virtual que contiene a [doc]

# ─────────────────────────────────────────────────────────────────────────────
# Parser de corchetes  →  árbol
# ─────────────────────────────────────────────────────────────────────────────

_OPEN = re.compile(r"\[([a-zA-Z][\w-]*)((?:\s+[\w:-]+=\"[^\"]*\")*)\s*\]")
_CLOSE = re.compile(r"\[/([a-zA-Z][\w-]*)\]")
_ATTR = re.compile(r"([\w:-]+)=\"([^\"]*)\"")


def _parse_attrs(raw: str) -> dict[str, str]:
    return {k: v for k, v in _ATTR.findall(raw or "")}


def parsear(texto: str) -> Nodo:
    """Convierte marcación de corchetes en un árbol con raíz `#RAIZ`.

    Tolerante: si un cierre no coincide con la etiqueta abierta más reciente,
    cierra hacia atrás hasta encontrarla (o lo ignora si no está abierta), para
    no romperse ante marcación imperfecta.
    """
    raiz = Nodo(ROOT_TAG)
    pila: list[Nodo] = [raiz]
    i, n = 0, len(texto)

    def _add_text(s: str):
        if not s:
            return
        tope = pila[-1]
        if tope.hijos and isinstance(tope.hijos[-1], str):
            tope.hijos[-1] += s           # fusiona texto contiguo
        else:
            tope.hijos.append(s)

    while i < n:
        if texto[i] == "[":
            mo = _OPEN.match(texto, i)
            mc = _CLOSE.match(texto, i)
            if mc:
                tag = mc.group(1)
                # cerrar: si el tope coincide, pop; si no, buscar hacia atrás
                if pila[-1].tag == tag:
                    pila.pop()
                elif any(nd.tag == tag for nd in pila[1:]):
                    while len(pila) > 1 and pila[-1].tag != tag:
                        pila.pop()
                    if len(pila) > 1:
                        pila.pop()
                # si no estaba abierta, se ignora el cierre huérfano
                i = mc.end()
                continue
            if mo:
                nodo = Nodo(mo.group(1), _parse_attrs(mo.group(2)))
                pila[-1].hijos.append(nodo)
                if not sps_tags.es_vacio(nodo.tag):
                    pila.append(nodo)     # los vacíos no abren contexto
                i = mo.end()
                continue
            # un '[' que no forma etiqueta: texto literal
            _add_text("[")
            i += 1
            continue
        # texto hasta el próximo '['
        j = texto.find("[", i)
        if j == -1:
            _add_text(texto[i:])
            break
        _add_text(texto[i:j])
        i = j
    return raiz


# ─────────────────────────────────────────────────────────────────────────────
# Serializador  árbol  →  corchetes
# ─────────────────────────────────────────────────────────────────────────────

def _attrs_str(attrs: dict[str, str]) -> str:
    return "".join(f' {k}="{v}"' for k, v in attrs.items())


def serializar(nodo: Nodo) -> str:
    """Árbol → marcación de corchetes. El nodo raíz `#RAIZ` no se emite."""
    partes: list[str] = []

    def _walk(nd: Nodo, es_raiz: bool):
        if not es_raiz:
            partes.append(f"[{nd.tag}{_attrs_str(nd.attrs)}]")
        for h in nd.hijos:
            if isinstance(h, str):
                partes.append(h)
            else:
                _walk(h, False)
        if not es_raiz:
            partes.append(f"[/{nd.tag}]")

    _walk(nodo, nodo.tag == ROOT_TAG)
    return "".join(partes)


# ─────────────────────────────────────────────────────────────────────────────
# Validación de anidación contra la biblioteca
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class Incidencia:
    tipo: str        # "tag-desconocida" | "anidacion" | "attr-desconocido" | "valor-invalido" | "attr-requerido"
    detalle: str


def validar(raiz: Nodo) -> list[Incidencia]:
    """Recorre el árbol y reporta: etiquetas no registradas, anidaciones
    inválidas, atributos desconocidos, valores fuera del vocabulario controlado
    y atributos requeridos ausentes."""
    inc: list[Incidencia] = []

    def _walk(nd: Nodo):
        padre = nd.tag
        for hijo in nd.hijos_nodo():
            if not sps_tags.existe(hijo.tag):
                inc.append(Incidencia("tag-desconocida",
                                      f"«{hijo.tag}» no está en la biblioteca"))
            elif padre != ROOT_TAG and not sps_tags.puede_anidar(padre, hijo.tag):
                inc.append(Incidencia("anidacion",
                                      f"«{hijo.tag}» no es válido dentro de «{padre}»"))
            # atributos del hijo
            if sps_tags.existe(hijo.tag):
                defs = sps_tags.atributos(hijo.tag)
                for a, v in hijo.attrs.items():
                    if a not in defs:
                        inc.append(Incidencia("attr-desconocido",
                                              f"«{hijo.tag}» no define el atributo «{a}»"))
                        continue
                    vals = defs[a].get("valores")
                    if vals and v not in vals and not _valor_compuesto_ok(v, vals):
                        inc.append(Incidencia("valor-invalido",
                                              f"«{hijo.tag} {a}=\"{v}\"» fuera del vocabulario"))
                for a, d in defs.items():
                    if d.get("req") and a not in hijo.attrs:
                        inc.append(Incidencia("attr-requerido",
                                              f"«{hijo.tag}» requiere el atributo «{a}»"))
            _walk(hijo)

    _walk(raiz)
    return inc


def _valor_compuesto_ok(valor: str, vals: list[str]) -> bool:
    """Permite valores combinados con «|» (p. ej. sec-type=\"materials|methods\")
    si cada parte pertenece al vocabulario."""
    if "|" not in valor:
        return False
    return all(p in vals for p in valor.split("|"))


# ─────────────────────────────────────────────────────────────────────────────
# Cobertura (para diagnóstico): tags y atributos vistos vs. biblioteca
# ─────────────────────────────────────────────────────────────────────────────

def cobertura(raiz: Nodo) -> dict:
    vistos_tags: set[str] = set()
    faltantes: set[str] = set()

    def _walk(nd: Nodo):
        for h in nd.hijos_nodo():
            vistos_tags.add(h.tag)
            if not sps_tags.existe(h.tag):
                faltantes.add(h.tag)
            _walk(h)

    _walk(raiz)
    return {
        "tags_vistas": sorted(vistos_tags),
        "tags_faltantes": sorted(faltantes),
        "en_biblioteca_sin_usar": sorted(set(sps_tags.TAGS) - vistos_tags),
    }
