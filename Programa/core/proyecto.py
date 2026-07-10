"""
proyecto.py
Serialización de un proyecto del Editor Semántico a un archivo portable .pmz.

Un .pmz es un ZIP autocontenido con:
  - session.json   → los datos del proyecto (bloques y su clasificación, autores,
                     afiliaciones, referencias, metadatos, pdf_info, y la lista de
                     figuras/tablas con rutas RELATIVAS a los recursos).
  - recursos/tablas/*.xlsx  → los .xlsx de las tablas.
  - recursos/figuras/*      → las imágenes de las figuras.

Así "el archivo" y "la sesión" son el mismo artefacto: guardar = empaquetar el
estado; abrir/reanudar = reconstruirlo. El módulo es puro (opera sobre un dict de
estado y rutas); server.py decide dónde viven los archivos.
"""
from __future__ import annotations

import os
import json
import shutil
import zipfile
from datetime import datetime
from typing import Any

FORMATO = "editor-semantico-proyecto"
VERSION = 1

# Claves de datos que se serializan tal cual a session.json.
_CLAVES_JSON = [
    "bloques",
    "referencias_externas",
    "autores_orcid",
    "afiliaciones_txt",
    "metadatos",
    "pdf_info",
]

# Claves derivadas (no persistibles) que nunca deben ir a session.json.
_DERIVADAS = ("preview", "sugiere_unir_siguiente")


def _entrada_recurso(item: dict, arcname: str | None) -> dict:
    """Copia un item (figura/tabla) sin su 'ruta' absoluta ni claves derivadas,
    poniendo en su lugar la ruta relativa dentro del zip ('archivo')."""
    entrada = {k: v for k, v in item.items()
               if k not in ("ruta",) and k not in _DERIVADAS}
    if arcname is not None:
        entrada["archivo"] = arcname
    return entrada


def empaquetar(estado: dict[str, Any], destino_zip: str,
               fuente: dict | None = None) -> None:
    """Escribe un .pmz a ``destino_zip`` a partir de ``estado``.

    ``estado`` es el dict de sesión del servidor (con figuras_manuales /
    tablas_manuales que traen ruta absoluta a archivos en disco).
    """
    session: dict[str, Any] = {
        "formato":  FORMATO,
        "version":  VERSION,
        "guardado": datetime.now().isoformat(timespec="seconds"),
    }
    if fuente:
        session["fuente"] = fuente
    for k in _CLAVES_JSON:
        session[k] = estado.get(k)

    os.makedirs(os.path.dirname(os.path.abspath(destino_zip)), exist_ok=True)
    figuras_meta: list[dict] = []
    tablas_meta:  list[dict] = []

    with zipfile.ZipFile(destino_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, f in enumerate(estado.get("figuras_manuales", []) or []):
            ruta = f.get("ruta", "")
            arc  = None
            if ruta and os.path.isfile(ruta):
                arc = f"recursos/figuras/{i:03d}_{os.path.basename(ruta)}"
                zf.write(ruta, arc)
            figuras_meta.append(_entrada_recurso(f, arc))

        for i, t in enumerate(estado.get("tablas_manuales", []) or []):
            ruta = t.get("ruta", "")
            arc  = None
            if ruta and os.path.isfile(ruta):
                arc = f"recursos/tablas/{i:03d}_{os.path.basename(ruta)}"
                zf.write(ruta, arc)
            tablas_meta.append(_entrada_recurso(t, arc))

        session["figuras_manuales"] = figuras_meta
        session["tablas_manuales"]  = tablas_meta
        zf.writestr("session.json",
                    json.dumps(session, ensure_ascii=False, indent=2))


def es_pmz(ruta: str) -> bool:
    """True si el archivo parece un .pmz válido (zip con session.json)."""
    try:
        with zipfile.ZipFile(ruta, "r") as zf:
            return "session.json" in zf.namelist()
    except (zipfile.BadZipFile, OSError):
        return False


def cargar(origen_zip: str, work_dir: str) -> dict[str, Any]:
    """Extrae los recursos de ``origen_zip`` en ``work_dir`` y devuelve el estado
    reconstruido, con las rutas de figuras/tablas apuntando a los archivos
    extraídos (ruta absoluta).

    Lanza ValueError si el archivo no es un proyecto válido.
    """
    if not zipfile.is_zipfile(origen_zip):
        raise ValueError("El archivo no es un proyecto válido (.pmz).")

    os.makedirs(work_dir, exist_ok=True)
    with zipfile.ZipFile(origen_zip, "r") as zf:
        if "session.json" not in zf.namelist():
            raise ValueError("El .pmz no contiene session.json.")
        session = json.loads(zf.read("session.json").decode("utf-8"))
        for name in zf.namelist():
            # Solo extraemos recursos; evitamos rutas fuera de work_dir (zip-slip).
            if not name.startswith("recursos/") or name.endswith("/"):
                continue
            destino = os.path.normpath(os.path.join(work_dir, name))
            if not destino.startswith(os.path.abspath(work_dir)):
                continue
            os.makedirs(os.path.dirname(destino), exist_ok=True)
            with zf.open(name) as src, open(destino, "wb") as dst:
                shutil.copyfileobj(src, dst)

    estado: dict[str, Any] = {k: session.get(k) for k in _CLAVES_JSON}

    def _reconstruir(items: list[dict] | None) -> list[dict]:
        salida = []
        for it in items or []:
            it = dict(it)
            arc = it.pop("archivo", None)
            it["ruta"] = os.path.join(work_dir, arc) if arc else ""
            salida.append(it)
        return salida

    estado["figuras_manuales"] = _reconstruir(session.get("figuras_manuales"))
    estado["tablas_manuales"]  = _reconstruir(session.get("tablas_manuales"))
    estado["_meta"] = {
        "guardado": session.get("guardado"),
        "version":  session.get("version"),
        "fuente":   session.get("fuente"),
    }
    return estado
