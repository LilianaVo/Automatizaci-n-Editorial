"""core/sps_bridge.py
Puente entre el estado actual del Editor (lista de bloques clasificados +
metadatos/autores/figuras/tablas/referencias) y el árbol de etiquetas SPS
(core/sps_documento.Nodo).

La clasificación de bloques es una GUÍA: este puente produce un primer borrador
del documento etiquetado que luego el usuario afina en las vistas. No pretende
una conversión perfecta desde la clasificación gruesa; sí un árbol VÁLIDO según
la biblioteca (core/sps_tags), que alimenta la Vista de Etiquetas (solo lectura
por ahora) y, más adelante, la edición sincronizada.

Punto de entrada:
    arbol   = estado_a_arbol(estado)      # Nodo raíz (#RAIZ → doc)
    markup  = estado_a_markup(estado)     # str en formato de corchetes de Markup
"""

from __future__ import annotations

import re
from typing import Any

from core.sps_documento import Nodo, ROOT_TAG, serializar

# Clasificaciones (de core.constans.OPCIONES) → rol en el árbol.
_SECCION = {"Encabezado sección", "Subencabezado"}
_SUBSECCION = {"Subencabezado-bajo"}
_PARRAFO = {"Cuerpo", "Normal"}
_ABSTRACT = {"Cuerpo del abstract", "Resumen / Abstract"}
_IGNORAR = {"Ignorar", "Imagen"}
# Clases que NO van en el cuerpo (se colocan en front/back o se omiten aquí).
_FUERA_CUERPO = _ABSTRACT | {
    "Título principal", "Título secundario", "Palabras clave", "Referencia",
    "Filiación", "Email / Metadatos", "Cómo citar", "Fecha manuscrito",
}

_RE_LABEL = re.compile(r"^\s*((?:tabla|table|figura|fig\.?|figure)\s+\d+\w?)\s*[\.\:\-]?\s*(.*)$",
                       re.IGNORECASE | re.DOTALL)


# ─────────────────────────────────────────────────────────────────────────────
# Helpers de construcción
# ─────────────────────────────────────────────────────────────────────────────

def _n(tag: str, attrs: dict | None = None, hijos: list | None = None) -> Nodo:
    return Nodo(tag, dict(attrs or {}), list(hijos or []))


def _txt(tag: str, texto: str, attrs: dict | None = None) -> Nodo:
    return Nodo(tag, dict(attrs or {}), [texto] if texto else [])


def _partir_label_caption(texto: str) -> tuple[str, str]:
    """«Tabla 1. Descripción» → ("Tabla 1", "Descripción")."""
    m = _RE_LABEL.match(texto.strip())
    if m:
        return m.group(1).strip(), (m.group(2) or "").strip()
    return "", texto.strip()


def _partir_nombre(nombre: str) -> tuple[str, str]:
    """Devuelve (apellido, nombres) de forma heurística."""
    nombre = (nombre or "").strip()
    if "," in nombre:
        ap, _, nom = nombre.partition(",")
        return ap.strip(), nom.strip()
    partes = nombre.split()
    if len(partes) >= 2:
        return partes[-1], " ".join(partes[:-1])
    return nombre, ""


# ─────────────────────────────────────────────────────────────────────────────
# Front matter
# ─────────────────────────────────────────────────────────────────────────────

def _doc_attrs(estado: dict) -> dict:
    m = estado.get("metadatos") or {}
    attrs = {"sps": "1.9", "doctopic": "oa", "language": "es"}
    if m.get("volumen"):       attrs["volid"] = str(m["volumen"])
    if m.get("numero"):        attrs["issueno"] = str(m["numero"])
    if m.get("pagina_inicio"): attrs["fpage"] = str(m["pagina_inicio"])
    if m.get("pagina_fin"):    attrs["lpage"] = str(m["pagina_fin"])
    if m.get("issn"):          attrs["issn"] = str(m["issn"])
    if m.get("anio"):          attrs["dateiso"] = f"{m['anio']}0000"
    return attrs


def _front_autores(doc: Nodo, estado: dict) -> None:
    for a in (estado.get("autores_orcid") or []):
        nombre = a.get("nombre") if isinstance(a, dict) else str(a)
        if not (nombre or "").strip():
            continue
        ap, nom = _partir_nombre(nombre)
        hijos = [_txt("surname", ap), _txt("fname", nom)]
        orcid = a.get("orcid") if isinstance(a, dict) else ""
        if orcid:
            hijos.append(_txt("authorid", orcid.strip(), {"authidtp": "orcid"}))
        doc.hijos.append(_n("author", {"role": "nd"}, hijos))


def _front_afiliaciones(doc: Nodo, estado: dict) -> None:
    txt = (estado.get("afiliaciones_txt") or "").strip()
    if not txt:
        return
    for i, linea in enumerate([l for l in txt.splitlines() if l.strip()], 1):
        doc.hijos.append(_n("normaff", {"id": f"aff{i}"},
                            [_txt("orgname", linea.strip())]))


def _front_abstract_keywords(doc: Nodo, bloques: list[dict]) -> None:
    # Resumen: bloques marcados como abstract, agrupados.
    parr = [b["contenido"] for b in bloques if b.get("clasificacion") in _ABSTRACT]
    if parr:
        doc.hijos.append(_n("xmlabstr", {"language": "es"},
                            [_txt("sectitle", "Resumen")] +
                            [_txt("p", p) for p in parr]))
    # Palabras clave.
    for b in bloques:
        if b.get("clasificacion") == "Palabras clave":
            cuerpo = re.sub(r"^\s*(palabras\s+clave|keywords)\s*[:\.]?\s*",
                            "", b["contenido"], flags=re.IGNORECASE)
            kws = [k.strip() for k in re.split(r"[;,]", cuerpo) if k.strip()]
            doc.hijos.append(_n("kwdgrp", {"language": "es"},
                                [_txt("sectitle", "Palabras clave")] +
                                [_txt("kwd", k) for k in kws]))
            break


def _front_hist(doc: Nodo, estado: dict) -> None:
    m = estado.get("metadatos") or {}
    campos = [("fecha_recibido_iso", "received"), ("fecha_corregido_iso", "revised"),
              ("fecha_aceptado_iso", "accepted")]
    hijos = []
    for iso_key, tag in campos:
        if m.get(iso_key):
            hijos.append(_txt(tag, "", {"dateiso": m[iso_key].replace("-", "")}))
    if hijos:
        doc.hijos.append(_n("hist", {}, hijos))


# ─────────────────────────────────────────────────────────────────────────────
# Cuerpo
# ─────────────────────────────────────────────────────────────────────────────

def _construir_cuerpo(bloques: list[dict]) -> Nodo:
    body = _n("xmlbody")
    sec_actual: Nodo | None = None
    sub_actual: Nodo | None = None
    contador_fig = contador_tab = 0

    def _asegurar_sec() -> Nodo:
        nonlocal sec_actual, sub_actual
        if sec_actual is None:
            sec_actual = _n("sec")
            body.hijos.append(sec_actual)
            sub_actual = None
        return sec_actual

    def _contenedor() -> Nodo:
        return sub_actual if sub_actual is not None else _asegurar_sec()

    for b in bloques:
        cls = b.get("clasificacion", "")
        texto = (b.get("contenido") or "").strip()
        if not texto or cls in _IGNORAR or cls in _FUERA_CUERPO:
            continue

        if cls in _SECCION:
            sec_actual = _n("sec", {}, [_txt("sectitle", texto)])
            body.hijos.append(sec_actual)
            sub_actual = None
        elif cls in _SUBSECCION:
            sub_actual = _n("subsec", {}, [_txt("sectitle", texto)])
            _asegurar_sec().hijos.append(sub_actual)
        elif cls == "Título tabla":
            contador_tab += 1
            lab, cap = _partir_label_caption(texto)
            hijos = []
            if lab: hijos.append(_txt("label", lab))
            hijos.append(_txt("caption", cap or texto))
            _contenedor().hijos.append(_n("tabwrap", {"id": f"t{contador_tab}"}, hijos))
        elif cls == "Pie de figura":
            contador_fig += 1
            lab, cap = _partir_label_caption(texto)
            hijos = [_n("graphic", {"href": ""})]
            if lab: hijos.append(_txt("label", lab))
            hijos.append(_txt("caption", cap or texto))
            _contenedor().hijos.append(_n("figgrp", {"id": f"f{contador_fig}"}, hijos))
        elif cls in _PARRAFO:
            _contenedor().hijos.append(_txt("p", texto))
        else:
            # clase no mapeada explícitamente → párrafo por defecto (guía)
            _contenedor().hijos.append(_txt("p", texto))
        # otros (Filiación, Email, Cómo citar, Fecha manuscrito, Título principal,
        # Palabras clave, Referencia) se tratan fuera del cuerpo o se omiten aquí.
    return body


# ─────────────────────────────────────────────────────────────────────────────
# Back (referencias)
# ─────────────────────────────────────────────────────────────────────────────

def _construir_refs(estado: dict, bloques: list[dict]) -> Nodo | None:
    refs_txt = list(estado.get("referencias_externas") or [])
    if not refs_txt:
        refs_txt = [b["contenido"] for b in bloques
                    if b.get("clasificacion") == "Referencia"]
    refs_txt = [r for r in refs_txt if (r or "").strip()]
    if not refs_txt:
        return None
    refs = _n("refs", {}, [_txt("sectitle", "Referencias")])
    for i, r in enumerate(refs_txt, 1):
        # Primer borrador: la cita cruda va en [source] (placeholder estructurable).
        refs.hijos.append(_n("ref", {"id": f"B{i}", "reftype": "other"},
                             [_txt("source", r.strip())]))
    return refs


# ─────────────────────────────────────────────────────────────────────────────
# Ensamblado
# ─────────────────────────────────────────────────────────────────────────────

def estado_a_arbol(estado: dict) -> Nodo:
    """Construye el árbol SPS (raíz #RAIZ → doc) a partir del estado del Editor."""
    bloques = estado.get("bloques") or []
    doc = _n("doc", _doc_attrs(estado))

    # DOI (si hay)
    doi = (estado.get("metadatos") or {}).get("doi")
    if doi:
        doc.hijos.append(_txt("doi", doi))

    # Título principal → doctitle
    for b in bloques:
        if b.get("clasificacion") == "Título principal" and (b.get("contenido") or "").strip():
            doc.hijos.append(_txt("doctitle", b["contenido"].strip(), {"language": "es"}))
            break

    _front_autores(doc, estado)
    _front_afiliaciones(doc, estado)
    _front_abstract_keywords(doc, bloques)
    _front_hist(doc, estado)

    doc.hijos.append(_construir_cuerpo(bloques))

    refs = _construir_refs(estado, bloques)
    if refs is not None:
        doc.hijos.append(refs)

    raiz = _n(ROOT_TAG, {}, [doc])
    return raiz


def estado_a_markup(estado: dict, sangria: bool = True) -> str:
    """Marcación de corchetes de Markup a partir del estado. Con `sangria`,
    inserta saltos de línea legibles entre bloques estructurales (solo estético;
    no altera el contenido de texto)."""
    arbol = estado_a_arbol(estado)
    if not sangria:
        return serializar(arbol)
    return _serializar_legible(arbol)


# Etiquetas estructurales tras las que conviene un salto de línea al mostrar.
_BLOQUE_LINEA = {"doc", "doctitle", "toctitle", "doi", "author", "normaff",
                 "corresp", "xmlabstr", "kwdgrp", "hist", "xmlbody", "sec",
                 "subsec", "sectitle", "p", "figgrp", "tabwrap", "list", "li",
                 "refs", "ref", "ack"}


def _serializar_legible(raiz: Nodo) -> str:
    """Como serializar(), pero con saltos de línea antes de cada etiqueta de
    bloque para lectura humana en la Vista de Etiquetas."""
    partes: list[str] = []

    def _walk(nd: Nodo, es_raiz: bool, prof: int):
        if not es_raiz:
            if nd.tag in _BLOQUE_LINEA and partes:
                partes.append("\n" + "  " * max(0, prof - 1))
            attrs = "".join(f' {k}="{v}"' for k, v in nd.attrs.items())
            partes.append(f"[{nd.tag}{attrs}]")
        for h in nd.hijos:
            if isinstance(h, str):
                partes.append(h)
            else:
                _walk(h, False, prof + 1)
        if not es_raiz:
            partes.append(f"[/{nd.tag}]")

    _walk(raiz, raiz.tag == ROOT_TAG, 0)
    return "".join(partes).lstrip("\n")
