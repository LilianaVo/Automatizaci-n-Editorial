"""
server.py
Backend FastAPI — Editor Semántico Paleontología Mexicana
Reemplaza app_window.py. Expone la lógica de core/ como endpoints REST.
"""

from __future__ import annotations

import os
import re
import sys
import shutil
import tempfile
import subprocess
import base64
from pathlib import Path
from typing import Any

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel

# ── core ──────────────────────────────────────────────────────────────────────
from core.pdf_processor    import procesar_pdf
from core.jats_exporterv2  import build_jats_xml
from core.html_exporter    import build_html
from core.epub_exporter    import build_epub
from core.xml_validator    import validar_jats
from core.constans import (
    OPCIONES,
    CLASE_COMPAT,
    COLORES_UI,
    COLOR_POR_CLASE,
)

# ─────────────────────────────────────────────────────────────────────────────
app = FastAPI(title="Editor Semántico — Paleontología Mexicana")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── Directorio de static (HTML/CSS/JS) ───────────────────────────────────────
BASE_DIR = Path(__file__).parent
STATIC_DIR = BASE_DIR / "static"
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

# ── Estado de sesión en memoria (una sesión a la vez — uso de escritorio) ────
_estado: dict[str, Any] = {
    "bloques":             [],   # list[dict] con clasificacion y contenido
    "referencias_externas":[],
    "figuras_manuales":    [],
    "tablas_manuales":     [],
    "autores_orcid":       [],
    "afiliaciones_txt":    "",
    "fig_dir":             None, # carpeta temporal de figuras extraídas
    "pdf_info":            {},   # nombre, páginas, tamaño
    "metadatos":           {},   # volumen, número, año, páginas, DOI, ISSN, fechas mss.
}


# ═════════════════════════════════════════════════════════════════════════════
# Modelos Pydantic
# ═════════════════════════════════════════════════════════════════════════════

class BloqueUpdate(BaseModel):
    idx: int
    contenido: str | None = None
    clasificacion: str | None = None

class BloquesBulkUpdate(BaseModel):
    bloques: list[dict]   # [{idx, contenido?, clasificacion?}]

class BloqueDividir(BaseModel):
    idx: int
    texto_nuevo: str
    contenido_restante: str
    clasificacion_nuevo: str | None = None

class BloqueNuevo(BaseModel):
    contenido: str
    clasificacion: str | None = None
    insertar_despues: int | None = None   # idx del bloque tras el cual insertar; None = al final

class UnirBloques(BaseModel):
    idx_a: int   # bloque que absorbe (queda)
    idx_b: int   # bloque que se une (desaparece)

class AutoresPayload(BaseModel):
    autores: list[dict]   # [{nombre, orcid}]

class AfiliacionesPayload(BaseModel):
    texto: str

class ReferenciasPayload(BaseModel):
    referencias: list[str]

class FigurasPayload(BaseModel):
    figuras: list[dict]   # [{pie, ruta?}]

class TablasPayload(BaseModel):
    tablas: list[dict]    # [{titulo, contenido}]

class ExportPayload(BaseModel):
    formato: str          # "html" | "xml" | "epub"
    ruta_destino: str     # ruta absoluta elegida por el usuario (desde pywebview dialog)

class MetadatosPayload(BaseModel):
    """Metadatos editoriales del artículo, confirmados/corregidos por el usuario.
    Todos los campos son opcionales: solo se actualizan los que vengan presentes.
    """
    volumen:             str | None = None
    numero:              str | None = None
    anio:                str | None = None
    pagina_inicio:       str | None = None
    pagina_fin:          str | None = None
    doi:                 str | None = None
    issn:                str | None = None
    fecha_recibido:      str | None = None
    fecha_corregido:     str | None = None
    fecha_aceptado:      str | None = None
    fecha_recibido_iso:  str | None = None
    fecha_corregido_iso: str | None = None
    fecha_aceptado_iso:  str | None = None


# ═════════════════════════════════════════════════════════════════════════════
# Endpoints — Meta / Config
# ═════════════════════════════════════════════════════════════════════════════

@app.get("/")
def root():
    """Sirve el index.html de la app."""
    return FileResponse(str(STATIC_DIR / "index.html"))


@app.get("/api/config")
def get_config():
    """Devuelve constantes de UI que el frontend necesita."""
    return {
        "opciones":      OPCIONES,
        "colores":       {c: col for c, col, _ in COLORES_UI},
        "clase_compat":  CLASE_COMPAT,
    }


# ═════════════════════════════════════════════════════════════════════════════
# Endpoints — PDF
# ═════════════════════════════════════════════════════════════════════════════

@app.post("/api/pdf/cargar")
async def cargar_pdf(file: UploadFile = File(...)):
    """
    Recibe el PDF, lo procesa con core.pdf_processor y devuelve los bloques
    clasificados, figuras y tablas detectadas.
    """
    # Guardar en temporal
    sufijo = Path(file.filename).suffix or ".pdf"
    with tempfile.NamedTemporaryFile(delete=False, suffix=sufijo) as tmp:
        contenido = await file.read()
        tmp.write(contenido)
        ruta_tmp = tmp.name

    try:
        resultado = procesar_pdf(ruta_tmp)
    except Exception as e:
        os.unlink(ruta_tmp)
        raise HTTPException(status_code=422, detail=f"Error procesando PDF: {e}")

    # Limpiar fig_dir anterior si existía
    if _estado["fig_dir"] and os.path.isdir(_estado["fig_dir"]):
        shutil.rmtree(_estado["fig_dir"], ignore_errors=True)

    # Guardar en estado de sesión
    # Convertimos los bloques al formato que usa el frontend:
    # {id, contenido, clasificacion, italic, size, bold}
    bloques_ui = []
    for i, b in enumerate(resultado["bloques"]):
        cls = CLASE_COMPAT.get(b.get("clasificacion", "Cuerpo"), b.get("clasificacion", "Cuerpo"))
        bloques_ui.append({
            "id":            i,
            "contenido":     b.get("contenido", ""),
            "clasificacion": cls,
            "italic":        b.get("italic", False),
            "bold":          b.get("bold", False),
            "size":          b.get("size", 12),
        })

    _estado["bloques"]              = bloques_ui
    _estado["figuras_manuales"]     = resultado.get("figuras", [])
    _estado["tablas_manuales"]      = resultado.get("tablas", [])
    _estado["referencias_externas"] = []
    _estado["fig_dir"]              = resultado.get("fig_dir")
    _estado["metadatos"]            = resultado.get("metadatos_detectados", {})
    _estado["pdf_info"]             = {
        "nombre":  file.filename,
        "paginas": resultado.get("body_size", "?"),   # reutilizamos para info
        "resumen": resultado.get("resumen", ""),
        "tamanio": f"{len(contenido) / 1024:.1f} KB",
    }

    os.unlink(ruta_tmp)

    return {
        "ok":        True,
        "info":      _estado["pdf_info"],
        "bloques":   bloques_ui,
        "figuras":   _estado["figuras_manuales"],
        "tablas":    _estado["tablas_manuales"],
        "metadatos": _estado["metadatos"],
        "resumen":   resultado.get("resumen", ""),
    }


@app.get("/api/pdf/info")
def get_pdf_info():
    return _estado["pdf_info"]


@app.post("/api/pdf/limpiar")
def limpiar_pdf():
    """Borra solo los bloques y la info del PDF. Conserva autores, afiliaciones y referencias."""
    if _estado["fig_dir"] and os.path.isdir(_estado["fig_dir"]):
        shutil.rmtree(_estado["fig_dir"], ignore_errors=True)
    _estado["bloques"]          = []
    _estado["figuras_manuales"] = []
    _estado["tablas_manuales"]  = []
    _estado["fig_dir"]          = None
    _estado["pdf_info"]         = {}
    _estado["metadatos"]        = {}
    _guardar_sesion()
    return {"ok": True}


class RutaPDFPayload(BaseModel):
    ruta: str

@app.post("/api/pdf/cargar-ruta")
def cargar_pdf_por_ruta(payload: RutaPDFPayload):
    """
    Versión para PyWebView: recibe la ruta absoluta del PDF en disco
    (ya que PyWebView no puede hacer drag-and-drop con File API).
    """
    ruta = payload.ruta
    if not os.path.isfile(ruta):
        raise HTTPException(status_code=404, detail=f"Archivo no encontrado: {ruta}")

    try:
        resultado = procesar_pdf(ruta)
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Error procesando PDF: {e}")

    if _estado["fig_dir"] and os.path.isdir(_estado["fig_dir"]):
        shutil.rmtree(_estado["fig_dir"], ignore_errors=True)

    bloques_ui = []
    for i, b in enumerate(resultado["bloques"]):
        cls = CLASE_COMPAT.get(b.get("clasificacion", "Cuerpo"), b.get("clasificacion", "Cuerpo"))
        bloques_ui.append({
            "id":            i,
            "contenido":     b.get("contenido", ""),
            "clasificacion": cls,
            "italic":        b.get("italic", False),
            "bold":          b.get("bold", False),
            "size":          b.get("size", 12),
        })

    tam = os.path.getsize(ruta)
    _estado["bloques"]              = bloques_ui
    _estado["figuras_manuales"]     = resultado.get("figuras", [])
    _estado["tablas_manuales"]      = resultado.get("tablas", [])
    _estado["referencias_externas"] = []
    _estado["fig_dir"]              = resultado.get("fig_dir")
    _estado["metadatos"]            = resultado.get("metadatos_detectados", {})
    _estado["pdf_info"]             = {
        "nombre":  Path(ruta).name,
        "paginas": resultado.get("body_size", "?"),
        "resumen": resultado.get("resumen", ""),
        "tamanio": f"{tam / 1024:.1f} KB",
    }

    return {
        "ok":        True,
        "info":      _estado["pdf_info"],
        "bloques":   bloques_ui,
        "figuras":   _estado["figuras_manuales"],
        "tablas":    _estado["tablas_manuales"],
        "metadatos": _estado["metadatos"],
    }


# ── Exportar preview (para desarrollo en browser sin PyWebView) ───────────────

from fastapi.responses import Response as FastAPIResponse

@app.post("/api/exportar/html/preview")
def exportar_html_preview():
    if not _estado["bloques"]:
        raise HTTPException(status_code=400, detail="No hay bloques cargados.")
    html_str = build_html(
        bloques=_bloques_snapshot(),
        referencias_externas=_estado["referencias_externas"],
        autores_orcid=_estado["autores_orcid"],
        afiliaciones_txt=_estado["afiliaciones_txt"],
        figuras=_estado["figuras_manuales"],
        tablas=_estado["tablas_manuales"],
    )
    return FastAPIResponse(content=html_str, media_type="text/html",
        headers={"Content-Disposition": "attachment; filename=articulo.html"})


@app.get("/api/exportar/html/vista-previa")
def vista_previa_html():
    """Devuelve el HTML generado sin header de descarga — para mostrarlo en iframe."""
    if not _estado["bloques"]:
        raise HTTPException(status_code=400, detail="No hay bloques cargados.")
    html_str = build_html(
        bloques=_bloques_snapshot(),
        referencias_externas=_estado["referencias_externas"],
        autores_orcid=_estado["autores_orcid"],
        afiliaciones_txt=_estado["afiliaciones_txt"],
        figuras=_estado["figuras_manuales"],
        tablas=_estado["tablas_manuales"],
    )
    return FastAPIResponse(content=html_str, media_type="text/html")

@app.post("/api/exportar/xml/preview")
def exportar_xml_preview():
    if not _estado["bloques"]:
        raise HTTPException(status_code=400, detail="No hay bloques cargados.")
    xml_str = build_jats_xml(
        bloques=_bloques_snapshot(),
        referencias_externas=_estado["referencias_externas"],
        autores_orcid=_estado["autores_orcid"],
        afiliaciones_txt=_estado["afiliaciones_txt"],
        figuras=_estado["figuras_manuales"],
        tablas=_estado["tablas_manuales"],
    )
    return FastAPIResponse(content=xml_str, media_type="application/xml",
        headers={"Content-Disposition": "attachment; filename=articulo.xml"})


# ═════════════════════════════════════════════════════════════════════════════
# Endpoints — Bloques
# ═════════════════════════════════════════════════════════════════════════════

@app.get("/api/bloques")
def get_bloques():
    return {"bloques": _estado["bloques"]}


@app.patch("/api/bloques/{idx}")
def actualizar_bloque(idx: int, datos: BloqueUpdate):
    if idx < 0 or idx >= len(_estado["bloques"]):
        raise HTTPException(status_code=404, detail="Bloque no encontrado")
    b = _estado["bloques"][idx]
    if datos.contenido is not None:
        b["contenido"] = datos.contenido
    if datos.clasificacion is not None:
        b["clasificacion"] = datos.clasificacion
    return {"ok": True, "bloque": b}


@app.delete("/api/bloques/{idx}")
def eliminar_bloque(idx: int):
    """Elimina un bloque por completo (acción irreversible, distinta de
    'Ignorar', que solo lo excluye del export pero lo conserva en la lista).
    """
    if idx < 0 or idx >= len(_estado["bloques"]):
        raise HTTPException(status_code=404, detail="Bloque no encontrado")

    _estado["bloques"].pop(idx)
    for nuevo_idx, b in enumerate(_estado["bloques"]):
        b["id"] = nuevo_idx

    return {"ok": True, "bloques": _estado["bloques"]}


@app.post("/api/bloques/dividir")
def dividir_bloque(datos: BloqueDividir):
    """Divide un bloque en dos a partir de una selección de texto."""
    idx = datos.idx
    if idx < 0 or idx >= len(_estado["bloques"]):
        raise HTTPException(status_code=404, detail="Bloque no encontrado")

    if not datos.texto_nuevo.strip():
        raise HTTPException(status_code=400, detail="El texto seleccionado está vacío.")

    original = _estado["bloques"][idx]
    clasificacion_nuevo = datos.clasificacion_nuevo or original.get("clasificacion", "Cuerpo")

    bloque_nuevo = {
        "id":            idx + 1,
        "contenido":     datos.texto_nuevo.strip(),
        "clasificacion": clasificacion_nuevo,
        "italic":        original.get("italic", False),
        "bold":          original.get("bold", False),
        "size":          original.get("size", 12),
    }

    original["contenido"] = datos.contenido_restante.strip()

    _estado["bloques"].insert(idx + 1, bloque_nuevo)
    for nuevo_idx, b in enumerate(_estado["bloques"]):
        b["id"] = nuevo_idx

    return {"ok": True, "bloques": _estado["bloques"]}




@app.post("/api/bloques/agregar")
def agregar_bloque(datos: BloqueNuevo):
    """Crea un bloque nuevo escrito manualmente por el usuario.

    Si se provee insertar_despues (idx de un bloque existente), el nuevo
    bloque se inserta justo debajo de ese. Si no se provee, se agrega al final.
    """
    if not datos.contenido.strip():
        raise HTTPException(status_code=400, detail="El contenido no puede estar vacio.")

    clasificacion = datos.clasificacion or "Cuerpo"

    bloque_nuevo = {
        "contenido":     datos.contenido.strip(),
        "clasificacion": clasificacion,
        "italic":        False,
        "bold":          False,
        "size":          12,
    }

    n = len(_estado["bloques"])
    idx_insertar = datos.insertar_despues

    if idx_insertar is not None and 0 <= idx_insertar < n:
        _estado["bloques"].insert(idx_insertar + 1, bloque_nuevo)
    else:
        _estado["bloques"].append(bloque_nuevo)

    for nuevo_idx, b in enumerate(_estado["bloques"]):
        b["id"] = nuevo_idx

    return {"ok": True, "bloques": _estado["bloques"]}

@app.post("/api/bloques/unir")
def unir_bloques(datos: UnirBloques):
    """Une dos bloques: el contenido de idx_b se append al de idx_a, separado
    por un espacio. idx_b se elimina. Siempre se reindexan los ids."""
    n = len(_estado["bloques"])
    a, b = datos.idx_a, datos.idx_b
    if not (0 <= a < n and 0 <= b < n):
        raise HTTPException(status_code=404, detail="Bloque no encontrado")
    if a == b:
        raise HTTPException(status_code=400, detail="No puedes unir un bloque consigo mismo")

    bloque_a = _estado["bloques"][a]
    bloque_b = _estado["bloques"][b]

    bloque_a["contenido"] = (bloque_a["contenido"].rstrip() + " " + bloque_b["contenido"].lstrip()).strip()

    _estado["bloques"].pop(b)
    for nuevo_idx, bl in enumerate(_estado["bloques"]):
        bl["id"] = nuevo_idx

    return {"ok": True, "bloques": _estado["bloques"]}


@app.put("/api/bloques")
def actualizar_bloques_bulk(payload: BloquesBulkUpdate):
    """Actualiza múltiples bloques de una vez (sync desde frontend)."""
    for item in payload.bloques:
        idx = item.get("idx") or item.get("id")
        if idx is None:
            continue
        if 0 <= idx < len(_estado["bloques"]):
            b = _estado["bloques"][idx]
            if "contenido" in item:
                b["contenido"] = item["contenido"]
            if "clasificacion" in item:
                b["clasificacion"] = item["clasificacion"]
    return {"ok": True}


# ═════════════════════════════════════════════════════════════════════════════
# Endpoints — Autores
# ═════════════════════════════════════════════════════════════════════════════

@app.get("/api/autores")
def get_autores():
    return {"autores": _estado["autores_orcid"]}


@app.put("/api/autores")
def set_autores(payload: AutoresPayload):
    _estado["autores_orcid"] = payload.autores
    return {"ok": True, "total": len(payload.autores)}


class RutaExcelPayload(BaseModel):
    ruta: str

def _parsear_excel_autores(ruta: str) -> tuple[list[dict], int]:
    """Lee un Excel de autores y devuelve (lista_autores, num_importados)."""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(ruta, data_only=True)
        ws = wb.active
        filas = list(ws.iter_rows(values_only=True))
        wb.close()
    except ImportError:
        raise HTTPException(status_code=500, detail="Instala openpyxl: pip install openpyxl")
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Error leyendo Excel: {e}")

    nuevos = []
    for fila in filas:
        if not fila or all(c is None for c in fila):
            continue
        nombre = str(fila[0]).strip() if fila[0] else ""
        orcid  = str(fila[1]).strip() if len(fila) > 1 and fila[1] else ""
        if nombre.lower() in ("autor", "nombre", "author", "name"):
            continue
        if not nombre:
            continue
        m = re.search(r"(\d{4}-\d{4}-\d{4}-\d{3}[\dX])", orcid, re.IGNORECASE)
        orcid_limpio = m.group(1) if m else orcid
        nuevos.append({"nombre": nombre, "orcid": orcid_limpio})
    return nuevos, len(nuevos)


@app.post("/api/autores/excel")
def importar_autores_excel_ruta(payload: RutaExcelPayload):
    """Para PyWebView — recibe ruta absoluta del Excel."""
    if not os.path.isfile(payload.ruta):
        raise HTTPException(status_code=404, detail="Archivo no encontrado")
    nuevos, total = _parsear_excel_autores(payload.ruta)
    _estado["autores_orcid"].extend(nuevos)
    return {"ok": True, "autores": _estado["autores_orcid"], "importados": total}


@app.post("/api/autores/excel-upload")
async def importar_autores_excel_upload(file: UploadFile = File(...)):
    """Para desarrollo en browser — recibe el archivo directamente."""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(await file.read())
        ruta_tmp = tmp.name
    try:
        nuevos, total = _parsear_excel_autores(ruta_tmp)
    finally:
        os.unlink(ruta_tmp)
    _estado["autores_orcid"].extend(nuevos)
    return {"ok": True, "autores": _estado["autores_orcid"], "importados": total}


# ═════════════════════════════════════════════════════════════════════════════
# Endpoints — Afiliaciones
# ═════════════════════════════════════════════════════════════════════════════

@app.get("/api/afiliaciones")
def get_afiliaciones():
    return {"texto": _estado["afiliaciones_txt"]}


@app.put("/api/afiliaciones")
def set_afiliaciones(payload: AfiliacionesPayload):
    _estado["afiliaciones_txt"] = payload.texto
    return {"ok": True}


# ═════════════════════════════════════════════════════════════════════════════
# Endpoints — Afiliaciones (carga desde .txt)
# ═════════════════════════════════════════════════════════════════════════════

class RutaTxtPayload(BaseModel):
    ruta: str

@app.post("/api/afiliaciones/txt")
def cargar_afiliaciones_txt(payload: RutaTxtPayload):
    """PyWebView: carga afiliaciones desde ruta de .txt en disco."""
    if not os.path.isfile(payload.ruta):
        raise HTTPException(status_code=404, detail="Archivo no encontrado")
    with open(payload.ruta, encoding="utf-8", errors="replace") as f:
        texto = f.read()
    _estado["afiliaciones_txt"] = texto
    return {"ok": True, "texto": texto}

@app.post("/api/afiliaciones/txt-upload")
async def cargar_afiliaciones_txt_upload(file: UploadFile = File(...)):
    """Browser fallback: recibe el .txt directamente."""
    contenido = await file.read()
    texto = contenido.decode("utf-8", errors="replace")
    _estado["afiliaciones_txt"] = texto
    return {"ok": True, "texto": texto}


# ═════════════════════════════════════════════════════════════════════════════
# Endpoints — Metadatos editoriales (volumen, número, año, páginas, DOI, ISSN,
# fechas de manuscrito)
# ═════════════════════════════════════════════════════════════════════════════
#
# NOTA IMPORTANTE: estos metadatos (detectados automáticamente y luego
# confirmados/corregidos por el usuario aquí) todavía NO se pasan a
# build_jats_xml() ni a build_html(). Esos exportadores (core/jats_exporterv2.py
# y core/html_exporter.py) tienen su propia extracción interna independiente
# (ver _extract_doi, _extract_year_and_pages, _parse_manuscript_dates en
# jats_exporterv2.py), que vuelve a leer directamente de los bloques del PDF.
# Conectar ambas fuentes es trabajo pendiente (Paso 4): hay que decidir si el
# dato corregido manualmente por el usuario debe tener prioridad sobre el
# extraído automáticamente antes de inyectarlo en el XML/HTML.

@app.get("/api/metadatos")
def get_metadatos():
    return {"metadatos": _estado["metadatos"]}


@app.put("/api/metadatos")
def set_metadatos(payload: MetadatosPayload):
    """Actualiza (parcialmente) los metadatos editoriales del artículo.
    Solo se sobrescriben los campos presentes en el payload; el resto
    conserva el valor ya detectado/guardado previamente.
    """
    actualizados = payload.model_dump(exclude_unset=True)
    _estado["metadatos"].update(actualizados)
    return {"ok": True, "metadatos": _estado["metadatos"]}


# ═════════════════════════════════════════════════════════════════════════════
# Endpoints — Referencias
# ═════════════════════════════════════════════════════════════════════════════

@app.get("/api/referencias")
def get_referencias():
    return {"referencias": _estado["referencias_externas"]}

@app.put("/api/referencias")
def set_referencias(payload: ReferenciasPayload):
    _estado["referencias_externas"] = payload.referencias
    return {"ok": True, "total": len(payload.referencias)}

@app.post("/api/referencias/txt")
def cargar_referencias_txt(payload: RutaTxtPayload):
    """PyWebView: carga referencias desde ruta de .txt en disco."""
    if not os.path.isfile(payload.ruta):
        raise HTTPException(status_code=404, detail="Archivo no encontrado")
    try:
        from core.utils import parsear_referencias
        with open(payload.ruta, encoding="utf-8", errors="replace") as f:
            contenido = f.read()
        refs = parsear_referencias(contenido)
        _estado["referencias_externas"] = refs
        return {"ok": True, "referencias": refs, "total": len(refs)}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/referencias/txt-upload")
async def cargar_referencias_txt_upload(file: UploadFile = File(...)):
    """Browser fallback."""
    try:
        from core.utils import parsear_referencias
        contenido = (await file.read()).decode("utf-8", errors="replace")
        refs = parsear_referencias(contenido)
        _estado["referencias_externas"] = refs
        return {"ok": True, "referencias": refs, "total": len(refs)}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# ═════════════════════════════════════════════════════════════════════════════
# Endpoints — Figuras
# ═════════════════════════════════════════════════════════════════════════════

def _figura_a_dict(fig: dict) -> dict:
    """Convierte una figura a dict con imagen en base64 si existe."""
    f = {k: v for k, v in fig.items() if not k.startswith("_")}
    ruta = f.get("ruta") or ""
    if ruta and os.path.isfile(ruta):
        try:
            with open(ruta, "rb") as fh:
                f["img_b64"] = base64.b64encode(fh.read()).decode()
            ext = Path(ruta).suffix.lower().lstrip(".")
            f["img_mime"] = f"image/{ext if ext in ('png','jpg','jpeg','gif','webp') else 'png'}"
        except Exception:
            f["img_b64"] = None
    return f

@app.get("/api/figuras")
def get_figuras():
    return {"figuras": [_figura_a_dict(f) for f in _estado["figuras_manuales"]]}

@app.put("/api/figuras")
def set_figuras(payload: FigurasPayload):
    _estado["figuras_manuales"] = payload.figuras
    return {"ok": True, "total": len(payload.figuras)}

class RutasImagenesPayload(BaseModel):
    rutas: list[str]   # lista de rutas absolutas (desde PyWebView)

@app.post("/api/figuras/agregar-rutas")
def agregar_figuras_por_rutas(payload: RutasImagenesPayload):
    """PyWebView: agrega imágenes por sus rutas en disco."""
    agregadas = 0
    for ruta in payload.rutas:
        if os.path.isfile(ruta):
            _estado["figuras_manuales"].append({"ruta": ruta, "pie": "", "ancla": ""})
            agregadas += 1
    return {
        "ok": True,
        "agregadas": agregadas,
        "figuras": [_figura_a_dict(f) for f in _estado["figuras_manuales"]],
    }

@app.post("/api/figuras/agregar-upload")
async def agregar_figuras_upload(files: list[UploadFile] = File(...)):
    """Browser fallback: recibe imágenes como upload y las guarda en temp."""
    if _estado["fig_dir"] is None:
        _estado["fig_dir"] = tempfile.mkdtemp(prefix="editor_figs_")
    agregadas = 0
    for file in files:
        ext  = Path(file.filename).suffix or ".png"
        dest = os.path.join(_estado["fig_dir"], f"fig_{len(_estado['figuras_manuales'])+agregadas+1}{ext}")
        with open(dest, "wb") as fh:
            fh.write(await file.read())
        _estado["figuras_manuales"].append({"ruta": dest, "pie": "", "ancla": ""})
        agregadas += 1
    return {
        "ok": True,
        "agregadas": agregadas,
        "figuras": [_figura_a_dict(f) for f in _estado["figuras_manuales"]],
    }

class ActualizarFiguraPayload(BaseModel):
    idx: int
    pie:   str | None = None
    ancla: str | None = None

@app.patch("/api/figuras/{idx}")
def actualizar_figura(idx: int, datos: ActualizarFiguraPayload):
    if idx < 0 or idx >= len(_estado["figuras_manuales"]):
        raise HTTPException(status_code=404, detail="Figura no encontrada")
    f = _estado["figuras_manuales"][idx]
    if datos.pie   is not None: f["pie"]   = datos.pie
    if datos.ancla is not None: f["ancla"] = datos.ancla
    return {"ok": True}

@app.delete("/api/figuras/{idx}")
def eliminar_figura(idx: int):
    if idx < 0 or idx >= len(_estado["figuras_manuales"]):
        raise HTTPException(status_code=404, detail="Figura no encontrada")
    _estado["figuras_manuales"].pop(idx)
    return {"ok": True, "figuras": [_figura_a_dict(f) for f in _estado["figuras_manuales"]]}


# ═════════════════════════════════════════════════════════════════════════════
# Endpoints — Tablas
# ═════════════════════════════════════════════════════════════════════════════

# Vista previa: cuántas filas/columnas se muestran en la tarjeta de cada tabla.
PREVIEW_MAX_FILAS = 8
PREVIEW_MAX_COLS  = 8

def _leer_filas_xlsx(ruta: str, hoja: str | None,
                     max_filas: int | None = None,
                     max_cols:  int | None = None) -> dict:
    """Lee una hoja de un .xlsx a matriz de strings para vista previa.

    El .xlsx es la fuente de verdad: se relee cada vez, así que refleja lo que
    el usuario haya editado en Excel. Devuelve un dict con la matriz ya capada
    y metadatos de cuántas filas/columnas tiene en total.
    """
    res = {
        "ok": False, "filas": [], "n_filas": 0, "n_cols": 0,
        "truncada_filas": False, "truncada_cols": False, "error": "",
        "mtime": 0.0,
    }
    if not ruta or not os.path.isfile(ruta):
        res["error"] = "archivo no encontrado"
        return res
    try:
        res["mtime"] = os.path.getmtime(ruta)   # para detectar ediciones en Excel
    except OSError:
        pass
    try:
        import openpyxl
    except ImportError:
        res["error"] = "openpyxl no disponible"
        return res
    try:
        wb = openpyxl.load_workbook(ruta, data_only=True, read_only=True)
        ws = wb[hoja] if hoja and hoja in wb.sheetnames else wb.active
        filas: list[list[str]] = []
        for fila in ws.iter_rows(values_only=True):
            vals = ["" if c is None else str(c) for c in fila]
            if any(v.strip() for v in vals):
                filas.append(vals)
        wb.close()
    except Exception as e:
        res["error"] = str(e)
        return res

    n_cols = max((len(f) for f in filas), default=0)
    filas = [f + [""] * (n_cols - len(f)) for f in filas]   # normaliza ancho
    res["n_filas"], res["n_cols"] = len(filas), n_cols

    if max_filas is not None and len(filas) > max_filas:
        filas = filas[:max_filas]
        res["truncada_filas"] = True
    if max_cols is not None and n_cols > max_cols:
        filas = [f[:max_cols] for f in filas]
        res["truncada_cols"] = True

    res["ok"], res["filas"] = True, filas
    return res

def _tablas_para_frontend() -> list[dict]:
    """Lista de tablas para la UI: datos limpios + vista previa (matriz capada).

    ``preview`` es un campo derivado (no se persiste en ``_estado``); se recalcula
    en cada respuesta releyendo el .xlsx.
    """
    salida = []
    for t in _estado["tablas_manuales"]:
        limpio = {k: v for k, v in t.items()
                  if not k.startswith("_") and k != "preview"}
        limpio["preview"] = _leer_filas_xlsx(
            t.get("ruta", ""), t.get("hoja"),
            max_filas=PREVIEW_MAX_FILAS, max_cols=PREVIEW_MAX_COLS)
        salida.append(limpio)
    return salida

@app.get("/api/tablas")
def get_tablas():
    return {"tablas": _tablas_para_frontend()}

@app.put("/api/tablas")
def set_tablas(payload: TablasPayload):
    # ``preview`` es derivado: no debe persistirse si el frontend lo reenvía.
    _estado["tablas_manuales"] = [
        {k: v for k, v in t.items() if k != "preview"} for t in payload.tablas
    ]
    return {"ok": True, "total": len(_estado["tablas_manuales"])}

class RutasExcelTablasPayload(BaseModel):
    rutas: list[str]

@app.post("/api/tablas/agregar-rutas")
def agregar_tablas_por_rutas(payload: RutasExcelTablasPayload):
    """PyWebView: agrega tablas desde rutas de Excel en disco."""
    try:
        import openpyxl
    except ImportError:
        raise HTTPException(status_code=500, detail="Instala openpyxl")
    agregadas = 0
    for ruta in payload.rutas:
        if not os.path.isfile(ruta):
            continue
        try:
            wb = openpyxl.load_workbook(ruta, data_only=True)
            for hoja in wb.sheetnames:
                ws = wb[hoja]
                filas = list(ws.iter_rows(values_only=True))
                preview = "\n".join(
                    "\t".join(str(c) if c is not None else "" for c in fila)
                    for fila in filas[:8]
                )
                _estado["tablas_manuales"].append({
                    "ruta":     ruta,
                    "hoja":     hoja,
                    "titulo":   "",
                    "ancla":    "",
                    "contenido": preview,
                })
                agregadas += 1
            wb.close()
        except Exception as e:
            raise HTTPException(status_code=422, detail=f"Error leyendo {Path(ruta).name}: {e}")
    return {"ok": True, "agregadas": agregadas, "tablas": _tablas_para_frontend()}

@app.post("/api/tablas/agregar-upload")
async def agregar_tablas_upload(files: list[UploadFile] = File(...)):
    """Browser fallback."""
    try:
        import openpyxl
    except ImportError:
        raise HTTPException(status_code=500, detail="Instala openpyxl")
    agregadas = 0
    for file in files:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(await file.read())
            ruta_tmp = tmp.name
        try:
            wb = openpyxl.load_workbook(ruta_tmp, data_only=True)
            for hoja in wb.sheetnames:
                ws = wb[hoja]
                filas = list(ws.iter_rows(values_only=True))
                preview = "\n".join(
                    "\t".join(str(c) if c is not None else "" for c in fila)
                    for fila in filas[:8]
                )
                _estado["tablas_manuales"].append({
                    "ruta":     ruta_tmp,
                    "hoja":     hoja,
                    "titulo":   "",
                    "ancla":    "",
                    "contenido": preview,
                })
                agregadas += 1
            wb.close()
        except Exception as e:
            os.unlink(ruta_tmp)
            raise HTTPException(status_code=422, detail=str(e))
    return {"ok": True, "agregadas": agregadas, "tablas": _tablas_para_frontend()}

class ActualizarTablaPayload(BaseModel):
    idx:    int
    titulo: str | None = None
    ancla:  str | None = None

@app.patch("/api/tablas/{idx}")
def actualizar_tabla(idx: int, datos: ActualizarTablaPayload):
    if idx < 0 or idx >= len(_estado["tablas_manuales"]):
        raise HTTPException(status_code=404, detail="Tabla no encontrada")
    t = _estado["tablas_manuales"][idx]
    if datos.titulo is not None: t["titulo"] = datos.titulo
    if datos.ancla  is not None: t["ancla"]  = datos.ancla
    return {"ok": True}

@app.delete("/api/tablas/{idx}")
def eliminar_tabla(idx: int):
    if idx < 0 or idx >= len(_estado["tablas_manuales"]):
        raise HTTPException(status_code=404, detail="Tabla no encontrada")
    _estado["tablas_manuales"].pop(idx)
    return {"ok": True, "tablas": _tablas_para_frontend()}

def _abrir_en_sistema(ruta: str) -> None:
    """Abre un archivo con la aplicación predeterminada del sistema operativo.

    El .xlsx sigue siendo la fuente de verdad: el usuario lo edita en Excel y el
    programa lo relee (ver _leer_filas_xlsx). En escritorio (pywebview) el backend
    corre en la máquina del usuario, así que esto abre Excel localmente.
    """
    if sys.platform.startswith("win"):
        os.startfile(ruta)                       # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", ruta])
    else:
        subprocess.Popen(["xdg-open", ruta])

@app.post("/api/tablas/{idx}/abrir-excel")
def abrir_tabla_excel(idx: int):
    """Abre el .xlsx de la tabla en Excel para editarlo (RF-29)."""
    if idx < 0 or idx >= len(_estado["tablas_manuales"]):
        raise HTTPException(status_code=404, detail="Tabla no encontrada")
    ruta = _estado["tablas_manuales"][idx].get("ruta", "")
    if not ruta or not os.path.isfile(ruta):
        raise HTTPException(
            status_code=404,
            detail="El archivo .xlsx de esta tabla ya no existe en disco.",
        )
    try:
        _abrir_en_sistema(ruta)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"No se pudo abrir el archivo: {e}")
    return {"ok": True, "ruta": ruta}


# ═════════════════════════════════════════════════════════════════════════════
# Endpoints — Exportar
# ═════════════════════════════════════════════════════════════════════════════

def _bloques_snapshot() -> list[dict]:
    """Convierte bloques de estado al formato que esperan los exportadores.
    Si hay referencias externas cargadas, corta todo lo que viene después
    del encabezado de sección "Referencias" del PDF — sin importar cómo
    estén clasificados esos bloques (Cuerpo, Referencia, Cómo citar, etc.).
    Así se evitan duplicados al exportar.
    """
    bloques = _estado["bloques"]
    hay_refs_externas = bool(_estado["referencias_externas"])

    if hay_refs_externas:
        # Encontrar el índice del encabezado "Referencias" en el PDF
        idx_corte = None
        for i, b in enumerate(bloques):
            cls  = b.get("clasificacion", "")
            cont = b.get("contenido", "").lower().strip()
            if cls == "Encabezado sección" and re.search(r"referencia|reference", cont):
                idx_corte = i
                break

        # Si encontró el encabezado, cortar ahí (incluye el encabezado,
        # lo excluye el html_exporter al detectar en_refs=True y sustituye
        # con las referencias externas)
        if idx_corte is not None:
            bloques = bloques[:idx_corte + 1]  # +1 para incluir el encabezado

    return [
        {
            "contenido":     b["contenido"],
            "clasificacion": b["clasificacion"],
            "italic":        b.get("italic", False),
        }
        for b in bloques
    ]


@app.post("/api/exportar/html")
def exportar_html(payload: ExportPayload):
    if not _estado["bloques"]:
        raise HTTPException(status_code=400, detail="No hay bloques cargados. Carga un PDF primero.")
    try:
        html_str = build_html(
            bloques              = _bloques_snapshot(),
            referencias_externas = _estado["referencias_externas"],
            autores_orcid        = _estado["autores_orcid"],
            afiliaciones_txt     = _estado["afiliaciones_txt"],
            figuras              = _estado["figuras_manuales"],
            tablas               = _estado["tablas_manuales"],
        )
        with open(payload.ruta_destino, "w", encoding="utf-8") as f:
            f.write(html_str)
        return {"ok": True, "ruta": payload.ruta_destino}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/exportar/xml")
def exportar_xml(payload: ExportPayload):
    if not _estado["bloques"]:
        raise HTTPException(status_code=400, detail="No hay bloques cargados. Carga un PDF primero.")
    try:
        xml_str = build_jats_xml(
            bloques              = _bloques_snapshot(),
            referencias_externas = _estado["referencias_externas"],
            autores_orcid        = _estado["autores_orcid"],
            afiliaciones_txt     = _estado["afiliaciones_txt"],
            figuras              = _estado["figuras_manuales"],
            tablas               = _estado["tablas_manuales"],
        )
        with open(payload.ruta_destino, "w", encoding="utf-8") as f:
            f.write(xml_str)
        return {"ok": True, "ruta": payload.ruta_destino}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/exportar/epub")
def exportar_epub(payload: ExportPayload):
    if not _estado["bloques"]:
        raise HTTPException(status_code=400, detail="No hay bloques cargados. Carga un PDF primero.")
    try:
        snap = _bloques_snapshot()

        html_str = build_html(
            bloques              = snap,
            referencias_externas = _estado["referencias_externas"],
            autores_orcid        = _estado["autores_orcid"],
            afiliaciones_txt     = _estado["afiliaciones_txt"],
            figuras              = _estado["figuras_manuales"],
            tablas               = _estado["tablas_manuales"],
        )

        titulo_art = next(
            (b["contenido"].strip() for b in snap if b["clasificacion"] == "Título principal"),
            "Artículo",
        )
        autores_lista = [
            a["nombre"].strip() for a in _estado["autores_orcid"]
            if a.get("nombre", "").strip()
        ]
        doi_art = ""
        for b in snap:
            m = re.search(r"https?://doi\.org/\S+", b["contenido"])
            if m:
                doi_art = m.group(0).rstrip(".")
                break
        secciones = [
            (re.sub(r"[^a-z0-9]", "-", b["contenido"].strip().lower())[:40].strip("-"),
             b["contenido"].strip())
            for b in snap
            if b["clasificacion"] == "Encabezado sección" and b["contenido"].strip()
        ]

        epub_bytes = build_epub(
            html_str  = html_str,
            titulo    = titulo_art,
            autores   = autores_lista,
            doi       = doi_art,
            secciones = secciones,
        )
        with open(payload.ruta_destino, "wb") as f:
            f.write(epub_bytes)
        return {"ok": True, "ruta": payload.ruta_destino}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/validar/xml")
def validar_xml_endpoint():
    """
    Genera el XML del estado actual y lo valida contra las reglas JATS/SciELO.
    No guarda ningún archivo — solo valida y devuelve el resultado.
    """
    if not _estado["bloques"]:
        raise HTTPException(
            status_code=400,
            detail="No hay bloques cargados. Carga un PDF primero."
        )
    try:
        xml_str = build_jats_xml(
            bloques              = _bloques_snapshot(),
            referencias_externas = _estado["referencias_externas"],
            autores_orcid        = _estado["autores_orcid"],
            afiliaciones_txt     = _estado["afiliaciones_txt"],
            figuras              = _estado["figuras_manuales"],
            tablas               = _estado["tablas_manuales"],
        )
        resultado = validar_jats(xml_str)
        return resultado.to_dict()
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# ═════════════════════════════════════════════════════════════════════════════
# Endpoints — Estado general
# ═════════════════════════════════════════════════════════════════════════════

@app.get("/api/estado")
def get_estado():
    """Resumen rápido del estado actual (para el stepper del frontend)."""
    return {
        "tiene_pdf":       bool(_estado["bloques"]),
        "num_bloques":     len(_estado["bloques"]),
        "num_autores":     len(_estado["autores_orcid"]),
        "tiene_afiliaciones": bool(_estado["afiliaciones_txt"].strip()),
        "num_referencias": len(_estado["referencias_externas"]),
        "num_figuras":     len(_estado["figuras_manuales"]),
        "num_tablas":      len(_estado["tablas_manuales"]),
        "tiene_metadatos": bool(_estado["metadatos"]) and any(_estado["metadatos"].values()),
        "pdf_info":        _estado["pdf_info"],
    }


@app.delete("/api/estado")
def resetear_estado():
    """Limpia toda la sesión."""
    if _estado["fig_dir"] and os.path.isdir(_estado["fig_dir"]):
        shutil.rmtree(_estado["fig_dir"], ignore_errors=True)
    _estado.update({
        "bloques": [], "referencias_externas": [], "figuras_manuales": [],
        "tablas_manuales": [], "autores_orcid": [], "afiliaciones_txt": "",
        "fig_dir": None, "pdf_info": {}, "metadatos": {},
    })
    return {"ok": True}


# ─────────────────────────────────────────────────────────────────────────────
# Arranque directo (desarrollo)
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import uvicorn
    uvicorn.run("server:app", host="127.0.0.1", port=8000, reload=False)