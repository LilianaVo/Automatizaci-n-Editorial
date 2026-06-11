# app.spec
# PyInstaller spec para Editor Semántico — Paleontología Mexicana
# Genera un .exe de ventana única con PyWebView + FastAPI embebido.
#
# Uso:
#   pyinstaller app.spec
#
# El ejecutable quedará en dist/EditorSemantico/EditorSemantico.exe

from PyInstaller.utils.hooks import collect_data_files, collect_submodules
import sys
import os

# ─── Rutas base ───────────────────────────────────────────────────────────────
block_cipher = None
BASE = os.path.dirname(os.path.abspath(SPEC))  # directorio de Programa/

# ─── Archivos de datos a incluir ──────────────────────────────────────────────
datas = [
    # Frontend completo (HTML, CSS, JS)
    (os.path.join(BASE, "static"),  "static"),

    # Assets de UI (logo, iconos, etc.)
    # (os.path.join(BASE, "ui", "assets"), os.path.join("ui", "assets")),
]

# Datos de paquetes externos que los necesitan
datas += collect_data_files("webview")
datas += collect_data_files("uvicorn")

# ─── Hidden imports ───────────────────────────────────────────────────────────
# FastAPI / uvicorn / starlette los necesitan explícitamente
hiddenimports = [
    # Uvicorn
    "uvicorn.logging",
    "uvicorn.loops",
    "uvicorn.loops.auto",
    "uvicorn.protocols",
    "uvicorn.protocols.http",
    "uvicorn.protocols.http.auto",
    "uvicorn.protocols.websockets",
    "uvicorn.protocols.websockets.auto",
    "uvicorn.lifespan",
    "uvicorn.lifespan.on",

    # FastAPI / Starlette
    "fastapi",
    "starlette",
    "starlette.routing",
    "starlette.staticfiles",
    "starlette.responses",
    "starlette.middleware.cors",
    "anyio",
    "anyio.from_thread",

    # Pydantic
    "pydantic",
    "pydantic.deprecated.class_validators",

    # PyWebView
    "webview",
    "webview.platforms.winforms",   # Windows
    # "webview.platforms.cocoa",    # macOS — descomentar si compilas en Mac
    # "webview.platforms.gtk",      # Linux — descomentar si compilas en Linux

    # Core de la app
    "core.pdf_processor",
    "core.jats_exporterv2",
    "core.html_exporter",
    "core.epub_exporter",
    "core.constans",
    "core.utils",

    # Dependencias de core
    "fitz",       # PyMuPDF
    "openpyxl",
    "ebooklib",
]

hiddenimports += collect_submodules("uvicorn")
hiddenimports += collect_submodules("starlette")

# ─── Análisis ─────────────────────────────────────────────────────────────────
a = Analysis(
    ["main.py"],          # nuevo entry point
    pathex        = [BASE],
    binaries      = [],
    datas         = datas,
    hiddenimports = hiddenimports,
    hookspath     = [],
    hooksconfig   = {},
    runtime_hooks = [],
    excludes      = [
        "tkinter",        # ya no lo necesitamos
        "customtkinter",  # ya no lo necesitamos
        "matplotlib",
        "numpy",
        "pandas",
        "scipy",
        "PIL",            # quitar si algún módulo de core lo usa
    ],
    win_no_prefer_redirects = False,
    win_private_assemblies  = False,
    cipher                  = block_cipher,
    noarchive               = False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries = True,
    name             = "EditorSemantico",
    debug            = False,
    bootloader_ignore_signals = False,
    strip            = False,
    upx              = True,
    console          = False,      # Sin ventana de consola
    # icon           = os.path.join(BASE, "ui", "assets", "icon.ico"),
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip = False,
    upx   = True,
    upx_exclude = [],
    name  = "EditorSemantico",
)
