"""
main.py
Entry point de la aplicación de escritorio.
Lanza FastAPI en un hilo y abre PyWebView como ventana nativa.
"""

from __future__ import annotations

import sys
import time
import threading
import socket
from pathlib import Path

import uvicorn
import webview

HOST = "127.0.0.1"
PORT = 8765
URL  = f"http://{HOST}:{PORT}"

# Cuando el usuario confirma cerrar (con o sin guardar), JS pone esto en True
# para que el manejador del evento 'closing' permita el cierre.
_permitir_cierre = {"v": False}

# Espejo en Python del estado "hay cambios sin guardar" (lo sincroniza JS vía
# set_sin_guardar). Permite que el cierre limpio sea instantáneo sin consultar
# a la interfaz (sin evaluate_js), evitando cualquier bloqueo del hilo GUI.
_sin_guardar = {"v": False}


class AppAPI:
    """
    Métodos accesibles desde JavaScript via window.pywebview.api.*
    IMPORTANTE — Windows PyWebView requiere el formato exacto:
        "Descripción (*.ext1;*.ext2)"
    """

    def abrir_pdf(self) -> str | None:
        rutas = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Archivos PDF (*.pdf)",),
        )
        return rutas[0] if rutas else None

    def abrir_documento(self) -> str | None:
        """RF-02 — Selector de artículo: PDF o Word (.docx) en un solo diálogo."""
        rutas = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=(
                "Artículo (*.pdf;*.docx)",
                "Archivos PDF (*.pdf)",
                "Documento Word (*.docx)",
            ),
        )
        return rutas[0] if rutas else None

    def guardar_html(self) -> str | None:
        ruta = webview.windows[0].create_file_dialog(
            webview.SAVE_DIALOG,
            save_filename="articulo.html",
            file_types=("Archivo HTML (*.html)",),
        )
        return ruta if isinstance(ruta, str) else (ruta[0] if ruta else None)

    def guardar_xml(self) -> str | None:
        ruta = webview.windows[0].create_file_dialog(
            webview.SAVE_DIALOG,
            save_filename="articulo.xml",
            file_types=("Archivo XML (*.xml)",),
        )
        return ruta if isinstance(ruta, str) else (ruta[0] if ruta else None)

    def guardar_epub(self) -> str | None:
        ruta = webview.windows[0].create_file_dialog(
            webview.SAVE_DIALOG,
            save_filename="articulo.epub",
            file_types=("Archivo EPUB (*.epub)",),
        )
        return ruta if isinstance(ruta, str) else (ruta[0] if ruta else None)

    def abrir_excel(self) -> str | None:
        """Un solo Excel (autores)."""
        rutas = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Archivo Excel (*.xlsx;*.xls)",),
        )
        return rutas[0] if rutas else None

    def abrir_excels_multiples(self) -> list[str]:
        """Múltiples Excel (tablas)."""
        rutas = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=True,
            file_types=("Archivo Excel (*.xlsx;*.xls)",),
        )
        return list(rutas) if rutas else []

    def abrir_txt(self) -> str | None:
        """Archivo .txt (afiliaciones o referencias)."""
        rutas = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Archivo de texto (*.txt)",),
        )
        return rutas[0] if rutas else None

    def abrir_imagenes(self) -> list[str]:
        """Múltiples imágenes (figuras)."""
        rutas = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=True,
            file_types=("Imágenes (*.jpg;*.jpeg;*.png;*.gif;*.webp;*.bmp)",),
        )
        return list(rutas) if rutas else []

    def abrir_pmz(self) -> str | None:
        """Abrir un proyecto del Editor Semántico (.pmz)."""
        rutas = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Proyecto Editor Semántico (*.pmz)",),
        )
        return rutas[0] if rutas else None

    def guardar_pmz(self, nombre_sugerido: str = "proyecto.pmz") -> str | None:
        """Guardar el proyecto actual como .pmz."""
        ruta = webview.windows[0].create_file_dialog(
            webview.SAVE_DIALOG,
            save_filename=nombre_sugerido or "proyecto.pmz",
            file_types=("Proyecto Editor Semántico (*.pmz)",),
        )
        return ruta if isinstance(ruta, str) else (ruta[0] if ruta else None)

    def seleccionar_carpeta(self) -> str | None:
        """Abre diálogo para elegir carpeta de salida predeterminada."""
        resultado = webview.windows[0].create_file_dialog(
            webview.FOLDER_DIALOG,
        )
        if resultado:
            return resultado[0] if isinstance(resultado, (list, tuple)) else resultado
        return None

    def set_sin_guardar(self, v: bool) -> None:
        """JS sincroniza aquí el estado 'hay cambios sin guardar'."""
        _sin_guardar["v"] = bool(v)

    def cerrar_confirmado(self) -> None:
        """Llamado desde JS cuando el usuario confirma cerrar (con o sin guardar)."""
        _permitir_cierre["v"] = True
        webview.windows[0].destroy()


# ─────────────────────────────────────────────────────────────────────────────

def _puerto_libre(host: str, port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex((host, port)) != 0


def _iniciar_servidor() -> None:
    from server import app as fastapi_app
    uvicorn.run(fastapi_app, host=HOST, port=PORT,
                log_level="error", access_log=False)


def _esperar_servidor(timeout: float = 15.0) -> bool:
    inicio = time.time()
    while time.time() - inicio < timeout:
        if not _puerto_libre(HOST, PORT):
            return True
        time.sleep(0.1)
    return False


def _al_cerrar() -> bool:
    """Manejador del evento 'closing' de la ventana.

    IMPORTANTE: no se puede llamar evaluate_js de forma síncrona aquí, porque
    este handler corre en el hilo de la GUI y evaluate_js esperaría a ese mismo
    hilo → deadlock (la app se cuelga como 'Not Responding'). Por eso: si el
    cierre aún no está confirmado, se lanza la consulta a la interfaz en un hilo
    aparte (donde evaluate_js sí puede despachar a la GUI, ya libre) y se aborta
    este cierre. La interfaz decide y llama a cerrar_confirmado()."""
    if _permitir_cierre["v"]:
        return True
    if not _sin_guardar["v"]:
        return True                       # sin cambios: cerrar directo, sin tocar JS

    def _preguntar() -> None:
        try:
            webview.windows[0].evaluate_js(
                "typeof App !== 'undefined' && App._alIntentarCerrar && App._alIntentarCerrar()"
            )
        except Exception:
            # Si algo falla al consultar la interfaz, no dejar la app atrapada.
            _permitir_cierre["v"] = True
            try:
                webview.windows[0].destroy()
            except Exception:
                pass

    threading.Thread(target=_preguntar, daemon=True).start()
    return False                          # abortar este cierre; decide la interfaz


def main() -> None:
    if not _puerto_libre(HOST, PORT):
        print(f"[WARN] Puerto {PORT} ya en uso.")

    hilo = threading.Thread(target=_iniciar_servidor, daemon=True)
    hilo.start()

    if not _esperar_servidor():
        print("[ERROR] El servidor no arrancó a tiempo.")
        sys.exit(1)

    ventana = webview.create_window(
        title       = "Editor Semántico — Paleontología Mexicana",
        url         = URL,
        js_api      = AppAPI(),
        width       = 1340,
        height      = 880,
        min_size    = (1100, 720),
        resizable   = True,
        text_select = True,
    )

    try:
        ventana.events.closing += _al_cerrar
    except Exception as e:
        print(f"[WARN] No se pudo enlazar el evento de cierre: {e}")

    webview.start(debug=False)


if __name__ == "__main__":
    main()