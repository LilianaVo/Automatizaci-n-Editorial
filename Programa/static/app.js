/**
 * app.js — Editor Semántico · Paleontología Mexicana · UNAM
 * Lógica completa del frontend. Se comunica con FastAPI via fetch()
 * y con diálogos nativos via window.pywebview.api.*
 */

"use strict";

// ─────────────────────────────────────────────────────────────────────────────
// Estado local
// ─────────────────────────────────────────────────────────────────────────────
const State = {
  seccionActiva : "pdf",
  bloques       : [],   // [{id, contenido, clasificacion, italic, bold, size}]
  autores       : [],   // [{nombre, orcid}]
  metadatos     : {},   // {volumen, numero, anio, pagina_inicio, pagina_fin, doi, issn, fecha_recibido, ...}
  config        : { opciones: [], colores: {} },
  tienePDF      : false,
  proyectos     : [],   // RF-03: pestañas abiertas [{id, titulo, activo, sin_guardar, tiene_contenido, ruta}]
  activo        : "",   // id del proyecto activo
  _afilTimer    : null,
};


// ─────────────────────────────────────────────────────────────────────────────
// Wrapper fetch  (GET / POST JSON / PUT JSON / POST FormData)
// ─────────────────────────────────────────────────────────────────────────────
// ¿La petición modifica el proyecto? (para marcar "cambios sin guardar", RF-04)
// Se excluyen proyecto (guardar/abrir), exportaciones y validación.
function _esMutacion(path) {
  return path.startsWith("/api/") &&
    !path.startsWith("/api/proyecto") &&   // cubre /api/proyecto/… y /api/proyectos… (pestañas)
    !path.startsWith("/api/exportar/") &&
    !path.startsWith("/api/validar");
}
// ¿Alguna pestaña (con contenido) tiene cambios sin guardar? El active se toma
// del vivo State.sinGuardar; las demás, de la última lista del backend.
function _algunaSinGuardar() {
  const activoSucio = !!(State.tienePDF && State.sinGuardar);
  const otras = (State.proyectos || []).some(
    p => p.id !== State.activo && p.tiene_contenido && p.sin_guardar
  );
  return activoSucio || otras;
}
// Fija el estado "hay cambios sin guardar" del proyecto ACTIVO y sincroniza con
// Python (main.py) el agregado de TODAS las pestañas, que se usa para decidir si
// avisar al cerrar la ventana.
function _dirty(v) {
  v = !!v;
  if (typeof State === "undefined") return;
  const cambio = State.sinGuardar !== v;
  State.sinGuardar = v;
  // Reflejar en la entrada del proyecto activo (para el agregado y los puntos).
  const p = (State.proyectos || []).find(p => p.id === State.activo);
  if (p) p.sin_guardar = v;
  if (!cambio) return;                        // sin cambio real: no re-renderizar ni empujar
  if (typeof App !== "undefined" && App._renderTabs) App._renderTabs();
  try { window.pywebview?.api?.set_sin_guardar?.(_algunaSinGuardar()); } catch (_) {}
}
function _marcarSinGuardar(path) {
  if (_esMutacion(path)) _dirty(true);
}

const API = {
  async get(path) {
    const r = await fetch(path);
    if (!r.ok) { const t = await r.text(); throw new Error(t); }
    return r.json();
  },
  async post(path, body) {
    const r = await fetch(path, {
      method  : "POST",
      headers : { "Content-Type": "application/json" },
      body    : JSON.stringify(body),
    });
    if (!r.ok) { const t = await r.text(); throw new Error(t); }
    _marcarSinGuardar(path);
    return r.json();
  },
  async put(path, body) {
    const r = await fetch(path, {
      method  : "PUT",
      headers : { "Content-Type": "application/json" },
      body    : JSON.stringify(body),
    });
    if (!r.ok) { const t = await r.text(); throw new Error(t); }
    _marcarSinGuardar(path);
    return r.json();
  },
  async patch(path, body) {
    const r = await fetch(path, {
      method  : "PATCH",
      headers : { "Content-Type": "application/json" },
      body    : JSON.stringify(body),
    });
    if (!r.ok) { const t = await r.text(); throw new Error(t); }
    _marcarSinGuardar(path);
    return r.json();
  },
  async delete(path) {
    const r = await fetch(path, { method: "DELETE" });
    if (!r.ok) { const t = await r.text(); throw new Error(t); }
    _marcarSinGuardar(path);
    return r.json();
  },
  async postForm(path, fd) {
    const r = await fetch(path, { method: "POST", body: fd });
    if (!r.ok) { const t = await r.text(); throw new Error(t); }
    _marcarSinGuardar(path);
    return r.json();
  },
  async postBlob(path) {
    const r = await fetch(path, { method: "POST" });
    if (!r.ok) { const t = await r.text(); throw new Error(t); }
    return r.blob();
  },
};


// ─────────────────────────────────────────────────────────────────────────────
// Helpers de UI
// ─────────────────────────────────────────────────────────────────────────────
function $(id)    { return document.getElementById(id); }
function esc(str) {
  return String(str ?? "")
    .replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")
    .replace(/"/g,"&quot;");
}

function setStatus(msg, tipo = "ok") {
  const dot  = $("status-dot");
  const text = $("status-text");
  if (!dot || !text) return;
  text.textContent = msg;
  dot.className = "status-dot"
    + (tipo === "error" ? " status-dot--error"
     : tipo === "warn"  ? " status-dot--warn"
     : tipo === "idle"  ? " status-dot--idle" : "");
}

let _toastTimer = null;
function showToast(msg, ms = 2800) {
  const t = $("toast");
  if (!t) return;
  t.textContent = msg;
  t.classList.add("toast--visible");
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => t.classList.remove("toast--visible"), ms);
}

function showLoading(on) {
  let ov = $("loading-overlay");
  if (on) {
    if (!ov) {
      ov = document.createElement("div");
      ov.id = "loading-overlay";
      ov.className = "loading-overlay";
      ov.innerHTML = '<div class="loading-spinner"></div>';
      document.body.appendChild(ov);
    }
    ov.style.display = "flex";
  } else if (ov) {
    ov.style.display = "none";
  }
}

function _progresoSteps() {
  /**
   * Evalúa cada paso y devuelve su estado:
   * "complete" — tiene datos suficientes ✅
   * "warn"     — tiene algo pero incompleto ⚠️
   * "empty"    — sin datos
   */
  const e = State;
  return {
    pdf: (() => {
      if (!e.tienePDF) return { estado: "empty", msg: "Sin PDF cargado" };
      const n = e.bloques.length;
      return { estado: "complete", msg: `${n} bloques clasificados` };
    })(),
    metadatos: (() => {
      const m = e.metadatos || {};
      const valores = Object.values(m).filter(v => (v || "").toString().trim());
      if (valores.length === 0) return { estado: "empty", msg: "Sin metadatos detectados" };
      const claves = ["volumen", "numero", "anio", "doi", "issn"];
      const faltantes = claves.filter(k => !(m[k] || "").toString().trim());
      if (faltantes.length > 0) {
        return { estado: "warn", msg: `Metadatos incompletos — falta: ${faltantes.join(", ")}` };
      }
      return { estado: "complete", msg: "Metadatos completos ✓" };
    })(),
    autores: (() => {
      const n = e.autores?.length || 0;
      if (n === 0) return { estado: "empty", msg: "Sin autores agregados" };
      const sinOrcid = e.autores.filter(a => !a.orcid?.trim()).length;
      if (sinOrcid > 0) return { estado: "warn", msg: `${n} autor(es) — ${sinOrcid} sin ORCID` };
      return { estado: "complete", msg: `${n} autor(es) con ORCID ✓` };
    })(),
    afiliaciones: (() => {
      const txt = (State._afilTxt || "").trim();
      if (!txt) return { estado: "empty", msg: "Sin afiliaciones" };
      const lineas = txt.split("\n").filter(l => l.trim()).length;
      return { estado: "complete", msg: `${lineas} afiliación(es) ✓` };
    })(),
    referencias: (() => {
      const n = State._numRefs || 0;
      if (n === 0) return { estado: "empty", msg: "Sin referencias cargadas" };
      return { estado: "complete", msg: `${n} referencia(s) ✓` };
    })(),
    figuras: (() => {
      const n = State._numFiguras || 0;
      if (n === 0) return { estado: "empty", msg: "Sin figuras (opcional)" };
      const sinPie = State._figurasSinPie || 0;
      if (sinPie > 0) return { estado: "warn", msg: `${n} figura(s) — ${sinPie} sin pie de figura` };
      return { estado: "complete", msg: `${n} figura(s) ✓` };
    })(),
    tablas: (() => {
      const n = State._numTablas || 0;
      if (n === 0) return { estado: "empty", msg: "Sin tablas (opcional)" };
      return { estado: "complete", msg: `${n} tabla(s) ✓` };
    })(),
    exportar: (() => {
      const tienePDF  = e.tienePDF;
      const tieneAut  = (e.autores?.length || 0) > 0;
      if (!tienePDF) return { estado: "empty", msg: "Completa los pasos anteriores" };
      if (!tieneAut) return { estado: "warn",  msg: "Faltan autores para exportar" };
      return { estado: "complete", msg: "Listo para exportar ✓" };
    })(),
  };
}

function actualizarStepper() {
  const orden  = ["pdf","metadatos","autores","afiliaciones","referencias","figuras","tablas","exportar"];
  const idx    = orden.indexOf(State.seccionActiva);
  const progreso = _progresoSteps();

  // Iconos SVG para cada estado
  const iconoCheck  = `<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M4 10l4 4 8-8"/></svg>`;
  const iconoWarn   = `<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2"><path d="M10 4v7M10 14v1"/><circle cx="10" cy="10" r="8"/></svg>`;

  orden.forEach((s, i) => {
    const el = $(`step-${s}`);
    if (!el) return;

    const p = progreso[s];
    const esActivo = i === idx;

    // Limpiar clases de estado
    el.classList.remove("step--active","step--done","step--complete","step--warn");

    if (esActivo) {
      el.classList.add("step--active");
    } else if (p.estado === "complete") {
      el.classList.add("step--complete");
      // Cambiar ícono del círculo a checkmark
      const circle = el.querySelector(".step-circle");
      if (circle) circle.innerHTML = iconoCheck;
    } else if (p.estado === "warn") {
      el.classList.add("step--warn");
      const circle = el.querySelector(".step-circle");
      if (circle) circle.innerHTML = iconoWarn;
    }

    // Actualizar tooltip
    let tip = el.querySelector(".step-tooltip");
    if (!tip) {
      tip = document.createElement("div");
      tip.className = "step-tooltip";
      el.appendChild(tip);
    }
    tip.textContent = p.msg;
  });

  // Colorear líneas entre pasos
  const lineas = document.querySelectorAll(".step-line");
  lineas.forEach((linea, i) => {
    linea.classList.remove("step-line--complete","step-line--warn");
    const p = progreso[orden[i]];
    if (p?.estado === "complete") linea.classList.add("step-line--complete");
    else if (p?.estado === "warn") linea.classList.add("step-line--warn");
  });
}

function actualizarSidebar(seccion) {
  document.querySelectorAll(".nav-item").forEach(btn =>
    btn.classList.toggle("nav-item--active", btn.dataset.seccion === seccion)
  );
}


// ─────────────────────────────────────────────────────────────────────────────
// Modal de confirmación personalizado (reemplaza window.confirm)
// Uso: const ok = await _modalConfirmar({ titulo, preview, mensaje, btnOk, peligro })
// ─────────────────────────────────────────────────────────────────────────────
function _modalConfirmar({ titulo = "Confirmar", preview = "", mensaje = "", btnOk = "Aceptar", peligro = false } = {}) {
  return new Promise(resolve => {
    // Reutilizar o crear el overlay
    let overlay = document.getElementById("modal-confirmar-overlay");
    if (!overlay) {
      overlay = document.createElement("div");
      overlay.id = "modal-confirmar-overlay";
      overlay.className = "modal-overlay";
      overlay.innerHTML = `
        <div class="modal-confirmar" role="dialog" aria-modal="true">
          <h3 class="modal-confirmar-titulo" id="modal-confirmar-titulo"></h3>
          <p  class="modal-confirmar-preview" id="modal-confirmar-preview"></p>
          <p  class="modal-confirmar-msg"     id="modal-confirmar-msg"></p>
          <div class="modal-confirmar-btns">
            <button class="btn btn--ghost" id="modal-confirmar-cancelar">Cancelar</button>
            <button class="btn"            id="modal-confirmar-ok">Aceptar</button>
          </div>
        </div>
      `;
      document.body.appendChild(overlay);
    }

    // Rellenar contenido
    document.getElementById("modal-confirmar-titulo").textContent  = titulo;
    const prevEl = document.getElementById("modal-confirmar-preview");
    if (preview) { prevEl.textContent = `"${preview}"`; prevEl.style.display = "block"; }
    else         {                                        prevEl.style.display = "none"; }
    document.getElementById("modal-confirmar-msg").textContent = mensaje;

    const btnOkEl = document.getElementById("modal-confirmar-ok");
    btnOkEl.textContent = btnOk;
    btnOkEl.className   = peligro ? "btn btn--danger" : "btn btn--primary";

    // Mostrar
    overlay.style.display = "flex";
    btnOkEl.focus();

    // Handlers (frescos cada vez para evitar listeners duplicados)
    const cerrar = (val) => {
      overlay.style.display = "none";
      document.removeEventListener("keydown", onKey);
      resolve(val);
    };

    const onKey = (e) => {
      if (e.key === "Escape") cerrar(false);
      if (e.key === "Enter")  cerrar(true);
    };

    btnOkEl.onclick = () => cerrar(true);
    document.getElementById("modal-confirmar-cancelar").onclick = () => cerrar(false);
    overlay.onclick = (e) => { if (e.target === overlay) cerrar(false); };
    document.addEventListener("keydown", onKey);
  });
}


// ─────────────────────────────────────────────────────────────────────────────
// Objeto principal App  (expuesto globalmente para los onclick del HTML)
// ─────────────────────────────────────────────────────────────────────────────
const App = {

  // ══════════════════════════════════════════════════════════════════════════
  // NAVEGACIÓN
  // ══════════════════════════════════════════════════════════════════════════

  irSeccion(seccion) {
    document.querySelectorAll(".panel").forEach(p => p.style.display = "none");
    const panel = $(`panel-${seccion}`);
    if (panel) panel.style.display = "flex";

    State.seccionActiva = seccion;
    actualizarSidebar(seccion);
    actualizarStepper();

    // Cargar datos según sección
    if (seccion === "referencias") App._cargarReferencias();
    if (seccion === "figuras")     App._cargarFiguras();
    if (seccion === "tablas")      App._cargarTablas();
    if (seccion === "autores")     App._renderAutores();
    if (seccion === "metadatos")   App._renderMetadatos();
  },


  // ══════════════════════════════════════════════════════════════════════════
  // PDF — carga
  // ══════════════════════════════════════════════════════════════════════════

  async seleccionarPDF() {
    try {
      if (window.pywebview) {
        // RF-02 — un solo selector para PDF y Word; si la app aún no expone
        // abrir_documento, cae en abrir_pdf.
        const api = window.pywebview.api;
        const ruta = api.abrir_documento ? await api.abrir_documento()
                                         : await api.abrir_pdf();
        if (ruta) await App._cargarPorRuta(ruta);
      } else {
        // Fallback desarrollo: input[type=file]
        const input = document.createElement("input");
        input.type   = "file";
        input.accept = ".pdf,.docx";
        input.onchange = async () => {
          if (input.files[0]) await App._cargarPorUpload(input.files[0]);
        };
        input.click();
      }
    } catch (e) {
      setStatus("Error: " + e.message, "error");
    }
  },

  async _cargarPorRuta(ruta) {
    const docx = /\.docx$/i.test(ruta);
    setStatus(`Procesando ${docx ? "Word" : "PDF"}...`, "idle");
    showLoading(true);
    try {
      const endpoint = docx ? "/api/docx/cargar-ruta" : "/api/pdf/cargar-ruta";
      const data = await API.post(endpoint, { ruta });
      // Guardar en historial antes de aplicar resultado
      const nombre = ruta.split(/[\\/]/).pop();
      Historial.agregar(ruta, nombre);
      App._aplicarResultadoPDF(data);
    } catch (e) {
      showLoading(false);
      setStatus("Error: " + e.message, "error");
      showToast(`No se pudo procesar el ${docx ? "Word" : "PDF"}`, 4000);
    }
  },

  async _cargarPorUpload(file) {
    const docx = /\.docx$/i.test(file.name || "");
    setStatus(`Procesando ${docx ? "Word" : "PDF"}...`, "idle");
    showLoading(true);
    try {
      const fd = new FormData();
      fd.append("file", file);
      const endpoint = docx ? "/api/docx/cargar" : "/api/pdf/cargar";
      const data = await API.postForm(endpoint, fd);
      App._aplicarResultadoPDF(data);
    } catch (e) {
      showLoading(false);
      setStatus("Error: " + e.message, "error");
      showToast(`No se pudo procesar el ${docx ? "Word" : "PDF"}`, 4000);
    }
  },

  _aplicarResultadoPDF(data) {
    showLoading(false);
    State.bloques   = data.bloques || [];
    State.metadatos = data.metadatos || {};
    State.tienePDF  = true;

    // Info del documento
    const info = data.info || {};
    const setVal = (id, val) => { const el=$(id); if(el) el.textContent = val||"—"; };
    setVal("info-nombre",  info.nombre);
    setVal("info-paginas", info.paginas);
    setVal("info-tamanio", info.tamanio);
    const badge = $("info-estado");
    if (badge) { badge.textContent = "Cargado"; badge.className = "badge badge--green"; }

    // Ocultar dropzone, mostrar bloques
    const dz = $("dropzone");
    const bc = $("bloques-container");
    if (dz) dz.style.display = "none";
    if (bc) bc.style.display = "block";

    // Mostrar controles de filtro y botón cambiar PDF
    const btnLey = $("btn-leyenda");
    const grpFil = $("filtro-grupo");
    const btnCambiar = $("btn-cambiar-pdf");
    if (btnLey)     btnLey.style.display     = "inline-flex";
    if (grpFil)     grpFil.style.display     = "flex";
    if (btnCambiar) btnCambiar.style.display = "inline-flex";

    App._poblarFiltro();
    App._renderBloques(State.bloques);
    App._renderMetadatos();
    actualizarStepper();

    // Ocultar historial mientras hay PDF activo
    const hist = document.getElementById("historial-recientes");
    if (hist) hist.style.display = "none";

    const n = State.bloques.length;
    setStatus(`PDF cargado — ${n} bloques detectados`);
    showToast(`✓ PDF procesado · ${n} bloques`);
  },

  // ── Drag & drop ─────────────────────────────────────────────────────────

  onDragOver(e) {
    e.preventDefault();
    $("dropzone")?.classList.add("dropzone--over");
  },
  onDragLeave() {
    $("dropzone")?.classList.remove("dropzone--over");
  },
  async onDrop(e) {
    e.preventDefault();
    $("dropzone")?.classList.remove("dropzone--over");
    const file = e.dataTransfer?.files?.[0];
    const esDocx = /\.docx$/i.test(file?.name || "") ||
      file?.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    if (file?.type === "application/pdf" || esDocx) {
      await App._cargarPorUpload(file);
    } else {
      showToast("Solo se aceptan archivos PDF o Word (.docx)");
    }
  },


  // ══════════════════════════════════════════════════════════════════════════
  // BLOQUES — render y edición
  // ══════════════════════════════════════════════════════════════════════════

  _poblarFiltro() {
    const sel = $("filtro-clase");
    if (!sel) return;
    const clases = ["Todos", ...new Set(State.bloques.map(b => b.clasificacion))];
    sel.innerHTML = clases.map(c => `<option value="${c}">${esc(c)}</option>`).join("");
  },

  filtrarBloques(clase) {
    const lista = clase === "Todos"
      ? State.bloques
      : State.bloques.filter(b => b.clasificacion === clase);
    App._renderBloques(lista);
  },

  _renderBloques(bloques) {
    const container = $("bloques-container");
    if (!container) return;

    const opciones = State.config.opciones || [];

    container.innerHTML = `
      <div class="bloques-toolbar">
        <span style="font-weight:600">${bloques.length} bloque${bloques.length !== 1 ? "s" : ""}</span>
        <button class="btn-agregar-bloque" onclick="App._agregarBloqueAlFinal()" title="Agregar bloque nuevo al final">
          <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2.2" width="14" height="14"><path d="M10 4v12M4 10h12"/></svg>
          Agregar bloque
        </button>
        <span style="margin-left:auto;font-size:11px;color:var(--text-light)">
          Clic en el texto para editar · Selector para reclasificar
        </span>
      </div>
      <div class="bloques-list" id="bloques-list"></div>
    `;

    const lista = $("bloques-list");
    bloques.forEach(b => {
      const idx = b.id ?? 0;
      const div = document.createElement("div");
      div.className       = "bloque-item";
      div.dataset.clase   = b.clasificacion;
      div.dataset.idx     = idx;
      div.addEventListener("contextmenu", (e) => App._mostrarMenuContextual(e, idx));

      // Construimos el select con la opción correcta pre-seleccionada
      const optsConSelected = opciones
        .map(o => `<option value="${esc(o)}"${o === b.clasificacion ? " selected" : ""}>${esc(o)}</option>`)
        .join("");

      div.innerHTML = `
        <div class="bloque-num">${idx + 1}</div>
        <textarea class="bloque-texto" id="bloque-texto-${idx}"
          oninput="this.style.height='auto';this.style.height=this.scrollHeight+'px'"
          onblur="App._onBloqueTextoBlur(${idx}, this.value)"
          onmouseup="App._onSeleccionTexto(${idx}, this)"
          onkeyup="App._onSeleccionTexto(${idx}, this)"
        >${esc(b.contenido)}</textarea>
        <select class="bloque-select"
          onchange="App._onBloqueClaseChange(${idx}, this.value, this.closest('.bloque-item'))">
          ${optsConSelected}
        </select>
        <button class="bloque-del" title="Marcar como Ignorar"
          onclick="App._ignorarBloque(${idx}, this.closest('.bloque-item'))">✕</button>
        <button class="bloque-eliminar" title="Eliminar bloque por completo (no se puede deshacer)"
          onclick="App._eliminarBloque(${idx})">🗑</button>
      `;

      lista.appendChild(div);

      // Botón "+" entre bloques para insertar debajo de este
      const btnEntre = document.createElement("button");
      btnEntre.className = "btn-insertar-entre";
      btnEntre.title = "Insertar bloque nuevo aquí";
      btnEntre.innerHTML = `<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2.2" width="12" height="12"><path d="M10 4v12M4 10h12"/></svg>`;
      btnEntre.onclick = () => App._mostrarModalAgregar(idx);
      lista.appendChild(btnEntre);
    });

    // Auto-altura inicial de todas las textareas
    lista.querySelectorAll(".bloque-texto").forEach(ta => {
      ta.style.height = "auto";
      ta.style.height = ta.scrollHeight + "px";
    });
  },

  async _onBloqueTextoBlur(idx, valor) {
    const b = State.bloques.find(b => b.id === idx);
    if (b) b.contenido = valor;
    try { await API.patch(`/api/bloques/${idx}`, { idx, contenido: valor }); }
    catch (_) {}
  },

  async _onBloqueClaseChange(idx, clase, rowEl) {
    const b = State.bloques.find(b => b.id === idx);
    if (b) b.clasificacion = clase;
    if (rowEl) rowEl.dataset.clase = clase;
    try { await API.patch(`/api/bloques/${idx}`, { idx, clasificacion: clase }); }
    catch (_) {}
  },

  _ignorarBloque(idx, rowEl) {
    App._onBloqueClaseChange(idx, "Ignorar", rowEl);
    if (rowEl) {
      const sel = rowEl.querySelector(".bloque-select");
      if (sel) sel.value = "Ignorar";
    }
  },

  async _eliminarBloque(idx) {
    const b = State.bloques.find(b => b.id === idx);
    const preview = b ? b.contenido.slice(0, 60).trim() : "";
    const textoPreview = preview + (preview.length === 60 ? "…" : "");

    const confirmado = await _modalConfirmar({
      titulo   : "Eliminar bloque",
      preview  : textoPreview,
      mensaje  : 'Esta acción no se puede deshacer. Si solo quieres excluirlo del export, usa "Ignorar" (✕) en su lugar.',
      btnOk    : "Eliminar",
      peligro  : true,
    });
    if (!confirmado) return;

    try {
      const data = await API.delete(`/api/bloques/${idx}`);
      State.bloques = data.bloques;
      App._poblarFiltro();
      App._renderBloques(State.bloques);
      showToast("Bloque eliminado");
    } catch (e) {
      showToast("Error al eliminar el bloque: " + e.message, 4000);
    }
  },

  // ── Agregar bloque nuevo ──────────────────────────────────────────────────

  _agregarBloqueAlFinal() {
    App._mostrarModalAgregar(null);
  },

  _mostrarModalAgregar(insertarDespues) {
    // Reutilizar o crear modal
    let overlay = $("modal-agregar-bloque");
    if (!overlay) {
      overlay = document.createElement("div");
      overlay.id = "modal-agregar-bloque";
      overlay.className = "modal-overlay";

      const opciones = State.config.opciones || [];
      const optsHtml = opciones.map(o => `<option value="${esc(o)}">${esc(o)}</option>`).join("");

      overlay.innerHTML = `
        <div class="modal-confirmar" role="dialog" aria-modal="true" style="width:520px;max-width:94vw">
          <h3 class="modal-confirmar-titulo">Agregar bloque nuevo</h3>
          <p class="modal-confirmar-msg" id="modal-agregar-pos"></p>
          <textarea id="modal-agregar-texto" rows="6"
            placeholder="Escribe o pega el texto del bloque aquí…"
            style="width:100%;margin-top:4px;padding:10px 12px;border:1.5px solid var(--border-2);
                   border-radius:var(--radius-sm);font-size:13px;font-family:inherit;
                   background:var(--surface);color:var(--text-main);resize:vertical;
                   line-height:1.55;min-height:100px;"></textarea>
          <div style="display:flex;align-items:center;gap:10px;margin-top:10px">
            <label style="font-size:12px;color:var(--text-sub);white-space:nowrap">Tipo de bloque:</label>
            <select id="modal-agregar-clase" style="flex:1;padding:6px 10px;border:1.5px solid var(--border-2);
              border-radius:var(--radius-sm);background:var(--surface);color:var(--text-main);font-size:13px">
              ${optsHtml}
            </select>
          </div>
          <div class="modal-confirmar-btns" style="margin-top:14px">
            <button class="btn--ghost" id="modal-agregar-cancelar">Cancelar</button>
            <button class="btn--primary" id="modal-agregar-ok">Agregar bloque</button>
          </div>
        </div>
      `;
      document.body.appendChild(overlay);
    }

    // Actualizar mensaje de posición
    const posEl = $("modal-agregar-pos");
    if (posEl) {
      posEl.textContent = insertarDespues !== null
        ? `Se insertará después del bloque ${insertarDespues + 1}.`
        : "Se agregará al final de la lista.";
    }

    // Limpiar textarea
    const ta = $("modal-agregar-texto");
    if (ta) { ta.value = ""; }

    overlay.style.display = "flex";
    if (ta) ta.focus();

    // Guardar posición en el overlay para usarla al confirmar
    overlay.dataset.insertarDespues = insertarDespues ?? "";

    const cerrar = () => {
      overlay.style.display = "none";
      document.removeEventListener("keydown", onKey);
    };

    const confirmar = async () => {
      const contenido     = ($("modal-agregar-texto")?.value || "").trim();
      const clasificacion = $("modal-agregar-clase")?.value || "Cuerpo";
      const pos           = overlay.dataset.insertarDespues;
      const insertar_despues = pos !== "" ? parseInt(pos) : null;

      if (!contenido) {
        showToast("Escribe algo antes de agregar el bloque.");
        return;
      }

      cerrar();
      try {
        const data = await API.post("/api/bloques/agregar", {
          contenido, clasificacion, insertar_despues,
        });
        State.bloques = data.bloques;
        App._poblarFiltro();
        App._renderBloques(State.bloques);
        showToast("✓ Bloque agregado");

        // Hacer scroll al bloque recién creado
        const nuevoPosicion = insertar_despues !== null ? insertar_despues + 1 : State.bloques.length - 1;
        setTimeout(() => {
          const nuevo = document.querySelector(`[data-idx="${nuevoPosicion}"]`);
          if (nuevo) nuevo.scrollIntoView({ behavior: "smooth", block: "center" });
        }, 100);
      } catch (e) {
        showToast("Error al agregar el bloque: " + e.message, 4000);
      }
    };

    const onKey = (e) => {
      if (e.key === "Escape") cerrar();
    };

    $("modal-agregar-ok").onclick     = confirmar;
    $("modal-agregar-cancelar").onclick = cerrar;
    overlay.onclick = (e) => { if (e.target === overlay) cerrar(); };
    document.addEventListener("keydown", onKey);
  },

  // ── Unir bloques ─────────────────────────────────────────────────────────

  async _unirConSiguiente(idx) {
    const bloques = State.bloques;
    const posActual = bloques.findIndex(b => b.id === idx);
    if (posActual === -1 || posActual >= bloques.length - 1) {
      showToast("No hay un bloque siguiente para unir.");
      return;
    }
    const siguiente = bloques[posActual + 1];
    const previewA = bloques[posActual].contenido.slice(0, 40).trim();
    const previewB = siguiente.contenido.slice(0, 40).trim();

    const confirmado = await _modalConfirmar({
      titulo  : "Unir bloques",
      preview : `"${previewA}…" + "${previewB}…"`,
      mensaje : "El texto del bloque siguiente se añadirá al final de este. El bloque siguiente desaparecerá.",
      btnOk   : "Unir",
      peligro : false,
    });
    if (!confirmado) return;

    try {
      const data = await API.post("/api/bloques/unir", { idx_a: idx, idx_b: siguiente.id });
      State.bloques = data.bloques;
      App._poblarFiltro();
      App._renderBloques(State.bloques);
      showToast("✓ Bloques unidos");
    } catch (e) {
      showToast("Error al unir bloques: " + e.message, 4000);
    }
  },

  // ── Menú contextual (clic derecho en tarjeta de bloque) ───────────────────

  _mostrarMenuContextual(e, idx) {
    e.preventDefault();

    // Cerrar cualquier menú previo
    App._cerrarMenuContextual();

    // Detectar si hay texto seleccionado en el textarea de este bloque
    const ta = $(`bloque-texto-${idx}`);
    const haySeleccion = ta &&
      ta.selectionEnd > ta.selectionStart &&
      ta.value.slice(ta.selectionStart, ta.selectionEnd).trim().length > 0;

    const bloques = State.bloques;
    const posActual = bloques.findIndex(b => b.id === idx);
    const hayBloqueSiguiente = posActual !== -1 && posActual < bloques.length - 1;

    const menu = document.createElement("div");
    menu.id = "ctx-menu-bloque";
    menu.className = "ctx-menu";
    menu.innerHTML = `
      <button class="ctx-menu-item" data-action="agregar">
        <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2" width="14" height="14"><path d="M10 4v12M4 10h12"/></svg>
        Agregar bloque aquí
      </button>
      <button class="ctx-menu-item ${haySeleccion ? "" : "ctx-menu-item--disabled"}" data-action="dividir">
        <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2" width="14" height="14"><path d="M4 10h12M10 4v3M10 13v3"/></svg>
        Dividir desde selección
      </button>
      <button class="ctx-menu-item ${hayBloqueSiguiente ? "" : "ctx-menu-item--disabled"}" data-action="unir">
        <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2" width="14" height="14"><path d="M4 7h12M4 13h12M10 4v12"/></svg>
        Unir con el siguiente
      </button>
    `;

    // Posición junto al cursor
    menu.style.position = "fixed";
    menu.style.top  = `${Math.min(e.clientY, window.innerHeight - 150)}px`;
    menu.style.left = `${Math.min(e.clientX, window.innerWidth  - 220)}px`;
    document.body.appendChild(menu);

    // Handlers
    menu.addEventListener("click", (ev) => {
      const btn = ev.target.closest("[data-action]");
      if (!btn || btn.classList.contains("ctx-menu-item--disabled")) return;
      App._cerrarMenuContextual();
      const action = btn.dataset.action;
      if (action === "agregar")  App._mostrarModalAgregar(idx);
      if (action === "dividir")  { if (ta) App._dividirBloqueDesdeSeleccion(idx, ta); }
      if (action === "unir")     App._unirConSiguiente(idx);
    });

    // Cerrar al hacer clic fuera (usamos click, no mousedown, para no
    // interferir con el click en las opciones del propio menú)
    const cerrarSiFuera = (ev) => {
      if (!menu.contains(ev.target)) {
        App._cerrarMenuContextual();
        document.removeEventListener("click", cerrarSiFuera);
      }
    };
    setTimeout(() => {
      document.addEventListener("click", cerrarSiFuera);
      document.addEventListener("keydown", (ev) => {
        if (ev.key === "Escape") {
          App._cerrarMenuContextual();
          document.removeEventListener("click", cerrarSiFuera);
        }
      }, { once: true });
    }, 0);
  },

  _cerrarMenuContextual() {
    const m = document.getElementById("ctx-menu-bloque");
    if (m) m.remove();
  },

  // ── Dividir bloque desde selección de texto ────────────────────────────

  _onSeleccionTexto(idx, textareaEl) {
    const inicio = textareaEl.selectionStart;
    const fin    = textareaEl.selectionEnd;
    const hayTexto = fin > inicio && textareaEl.value.slice(inicio, fin).trim().length > 0;

    let boton = $("btn-flotante-dividir");

    if (!hayTexto) {
      if (boton) boton.style.display = "none";
      return;
    }

    if (!boton) {
      boton = document.createElement("button");
      boton.id = "btn-flotante-dividir";
      boton.className = "btn-flotante-dividir";
      boton.textContent = "✂ Hacer bloque desde selección";
      document.body.appendChild(boton);
    }

    boton.dataset.idx = idx;
    boton.onclick = () => App._dividirBloqueDesdeSeleccion(idx, textareaEl);

    const rect = textareaEl.getBoundingClientRect();
    boton.style.display = "block";
    boton.style.position = "fixed";
    boton.style.top  = `${rect.top - 36}px`;
    boton.style.left = `${rect.left + 8}px`;
  },

  async _dividirBloqueDesdeSeleccion(idx, textareaEl) {
    const inicio = textareaEl.selectionStart;
    const fin    = textareaEl.selectionEnd;
    const completo = textareaEl.value;

    const texto_nuevo        = completo.slice(inicio, fin).trim();
    const contenido_restante = (completo.slice(0, inicio) + completo.slice(fin)).trim();

    const boton = $("btn-flotante-dividir");
    if (boton) boton.style.display = "none";

    if (!texto_nuevo) return;

    if (!contenido_restante) {
      showToast("La selección abarca todo el bloque — no hay nada que dividir.");
      return;
    }

    try {
      const data = await API.post("/api/bloques/dividir", {
        idx, texto_nuevo, contenido_restante,
      });
      State.bloques = data.bloques;
      App._poblarFiltro();
      App._renderBloques(State.bloques);
      showToast("Bloque dividido en dos ✓");
    } catch (e) {
      showToast("Error al dividir el bloque: " + e.message, 4000);
    }
  },

  // ── Leyenda ──────────────────────────────────────────────────────────────

  mostrarLeyenda() {
    const cont = $("leyenda-contenido");
    const cols = State.config.colores || {};
    if (!cont) return;
    cont.innerHTML = Object.entries(cols).map(([nombre, color]) => `
      <div class="leyenda-item">
        <div class="leyenda-dot" style="background:${color}"></div>
        <span class="leyenda-name">${esc(nombre)}</span>
        <span style="font-size:11px;color:var(--text-light)">${color}</span>
      </div>
    `).join("");
    $("modal-leyenda").style.display = "flex";
  },

  cerrarLeyenda() {
    $("modal-leyenda").style.display = "none";
  },


  // ══════════════════════════════════════════════════════════════════════════
  // AUTORES
  // ══════════════════════════════════════════════════════════════════════════

  async agregarAutor() {
    App._leerInputsAutores();   // guardar valores actuales antes de redibujar
    State.autores.push({ nombre: "", orcid: "", afiliaciones: "" });
    App._renderAutores();
    await App._pushAutores();
    // Hacer foco en el último input de nombre
    const rows = document.querySelectorAll(".autor-row");
    if (rows.length) {
      const lastInput = rows[rows.length - 1].querySelector(".autor-input");
      lastInput?.focus();
    }
  },

  async limpiarAutores() {
    State.autores = [];
    App._renderAutores();
    await App._pushAutores();
    showToast("Autores eliminados");
  },

  async cargarAutoresExcel() {
    try {
      if (window.pywebview) {
        const ruta = await window.pywebview.api.abrir_excel();
        if (!ruta) return;
        setStatus("Importando Excel...", "idle");
        const data = await API.post("/api/autores/excel", { ruta });
        State.autores = data.autores || [];
        App._renderAutores();
        showToast(`✓ ${data.importados} autores importados`);
        setStatus("Autores importados desde Excel");
      } else {
        const input = document.createElement("input");
        input.type   = "file";
        input.accept = ".xlsx,.xls";
        input.onchange = async () => {
          if (!input.files[0]) return;
          const fd = new FormData();
          fd.append("file", input.files[0]);
          const data = await API.postForm("/api/autores/excel-upload", fd);
          State.autores = data.autores || [];
          App._renderAutores();
          showToast(`✓ ${data.importados} autores importados`);
        };
        input.click();
      }
    } catch (e) {
      setStatus("Error importando Excel: " + e.message, "error");
      showToast("Error al importar Excel", 4000);
    }
  },

  _renderAutores() {
    const lista = $("autores-list");
    if (!lista) return;

    if (State.autores.length === 0) {
      lista.innerHTML = `
        <div class="empty-state">
          <svg viewBox="0 0 48 48" fill="none" stroke="currentColor" stroke-width="1.2"
            style="width:48px;height:48px;opacity:.25">
            <circle cx="20" cy="16" r="8"/><path d="M4 40c0-8.8 7.2-16 16-16s16 7.2 16 16"/>
          </svg>
          <p>Sin autores cargados</p>
          <p class="empty-hint">Haz clic en "Agregar autor" para comenzar</p>
        </div>`;
      return;
    }

    lista.innerHTML = State.autores.map((a, i) => `
      <div class="autor-row" data-autor-idx="${i}">
        <div class="autor-num">${i + 1}</div>
        <input class="autor-input" type="text"
          placeholder="Apellido, Nombre"
          value="${esc(a.nombre || "")}"
          onblur="App._onAutorBlur(${i}, 'nombre', this.value)"
        />
        <input class="autor-input" type="text"
          placeholder="0000-0001-2345-6789"
          value="${esc(a.orcid || "")}"
          onblur="App._onAutorBlur(${i}, 'orcid', this.value)"
        />
        <input class="autor-input autor-input--aff" type="text"
          placeholder="Afiliación(es), ej. 1 o 1,2"
          title="Número(s) o letra(s) de afiliación tal como aparecen junto al nombre del autor en el PDF"
          value="${esc(a.afiliaciones || "")}"
          onblur="App._onAutorBlur(${i}, 'afiliaciones', this.value)"
        />
        <button class="autor-del" onclick="App._eliminarAutor(${i})" title="Eliminar">✕</button>
      </div>
    `).join("");
  },

  async _onAutorBlur(idx, campo, valor) {
    if (State.autores[idx]) {
      State.autores[idx][campo] = valor.trim();
      await App._pushAutores();
    }
  },

  async _eliminarAutor(idx) {
    App._leerInputsAutores();
    State.autores.splice(idx, 1);
    App._renderAutores();
    await App._pushAutores();
  },

  /** Lee los valores actuales de los inputs y los guarda en State.autores */
  _leerInputsAutores() {
    document.querySelectorAll(".autor-row").forEach(row => {
      const i = parseInt(row.dataset.autorIdx, 10);
      if (isNaN(i) || !State.autores[i]) return;
      const inputs = row.querySelectorAll(".autor-input");
      if (inputs[0]) State.autores[i].nombre       = inputs[0].value.trim();
      if (inputs[1]) State.autores[i].orcid        = inputs[1].value.trim();
      if (inputs[2]) State.autores[i].afiliaciones = inputs[2].value.trim();
    });
  },

  async _pushAutores() {
    try { await API.put("/api/autores", { autores: State.autores }); }
    catch (_) {}
  },


  // ══════════════════════════════════════════════════════════════════════════
  // METADATOS EDITORIALES (volumen, número, año, páginas, DOI, ISSN, fechas)
  // ══════════════════════════════════════════════════════════════════════════
  //
  // NOTA: estos valores se detectan automáticamente al cargar el PDF
  // (resultado["metadatos_detectados"] en core/pdf_processor.py) y el usuario
  // puede corregirlos aquí. Por ahora esta corrección NO se inyecta en el
  // XML/HTML exportado — los exportadores (core/jats_exporterv2.py,
  // core/html_exporter.py) siguen extrayendo estos mismos datos por su cuenta
  // directamente de los bloques del PDF. Conectar ambas fuentes es trabajo
  // pendiente (ver server.py, sección Endpoints — Metadatos editoriales).

  _CAMPOS_METADATOS: [
    ["volumen",             "meta-volumen"],
    ["numero",              "meta-numero"],
    ["anio",                "meta-anio"],
    ["pagina_inicio",       "meta-pagina-inicio"],
    ["pagina_fin",          "meta-pagina-fin"],
    ["issn",                "meta-issn"],
    ["doi",                 "meta-doi"],
    ["idioma",               "meta-idioma"],
    ["fecha_recibido",      "meta-fecha-recibido"],
    ["fecha_corregido",     "meta-fecha-corregido"],
    ["fecha_aceptado",      "meta-fecha-aceptado"],
  ],

  /** Pinta los inputs del panel Metadatos desde State.metadatos. */
  _renderMetadatos() {
    const m = State.metadatos || {};
    App._CAMPOS_METADATOS.forEach(([campo, id]) => {
      const el = $(id);
      if (!el) return;
      // 'idioma' no se detecta automáticamente del PDF; "es" es el default
      // razonable para Paleontología Mexicana hasta que el usuario lo cambie.
      el.value = m[campo] || (campo === "idioma" ? "es" : "");
    });

    const badge = $("metadatos-estado");
    if (badge) {
      if (!State.tienePDF) {
        badge.textContent = "Sin PDF cargado";
        badge.className = "badge badge--gray";
      } else {
        const valores = Object.values(m).filter(v => (v || "").toString().trim());
        if (valores.length === 0) {
          badge.textContent = "Nada detectado";
          badge.className = "badge badge--gray";
        } else {
          badge.textContent = "Detectado desde PDF";
          badge.className = "badge badge--green";
        }
      }
    }
  },

  /** Handler de onblur de cada input de metadatos: actualiza State y guarda en el backend. */
  async syncMetadato(campo, valor) {
    const v = (valor || "").trim();
    State.metadatos = State.metadatos || {};
    State.metadatos[campo] = v;
    actualizarStepper();
    try {
      await API.put("/api/metadatos", { [campo]: v });
    } catch (e) {
      showToast("No se pudo guardar el metadato", 3000);
    }
  },


  // ══════════════════════════════════════════════════════════════════════════
  // AFILIACIONES
  // ══════════════════════════════════════════════════════════════════════════

  async cargarAfiliacionesTxt() {
    try {
      if (window.pywebview) {
        const ruta = await window.pywebview.api.abrir_txt();
        if (!ruta) return;
        const data = await API.post("/api/afiliaciones/txt", { ruta });
        const ta = $("afiliaciones-txt");
        if (ta) ta.value = data.texto;
        showToast("✓ Afiliaciones cargadas desde .txt");
        setStatus("Afiliaciones cargadas");
      } else {
        const input = document.createElement("input");
        input.type = "file"; input.accept = ".txt";
        input.onchange = async () => {
          if (!input.files[0]) return;
          const fd = new FormData();
          fd.append("file", input.files[0]);
          const data = await API.postForm("/api/afiliaciones/txt-upload", fd);
          const ta = $("afiliaciones-txt");
          if (ta) ta.value = data.texto;
          showToast("✓ Afiliaciones cargadas");
        };
        input.click();
      }
    } catch (e) {
      setStatus("Error: " + e.message, "error");
      showToast("Error al cargar afiliaciones", 4000);
    }
  },

  limpiarAfiliaciones() {
    const ta = $("afiliaciones-txt");
    if (ta) ta.value = "";
    API.put("/api/afiliaciones", { texto: "" }).catch(() => {});
    showToast("Afiliaciones limpiadas");
  },

  syncAfiliaciones() {
    clearTimeout(State._afilTimer);
    const txt = $("afiliaciones-txt");
    if (txt) {
      State._afilTxt = txt.value;
      actualizarStepper();
    }
    State._afilTimer = setTimeout(async () => {
      if (!txt) return;
      try { await API.put("/api/afiliaciones", { texto: txt.value }); }
      catch (_) {}
    }, 700);
  },


  // ══════════════════════════════════════════════════════════════════════════
  // REFERENCIAS
  // ══════════════════════════════════════════════════════════════════════════

  async cargarReferenciassTxt() {
    try {
      if (window.pywebview) {
        const ruta = await window.pywebview.api.abrir_txt();
        if (!ruta) return;
        setStatus("Cargando referencias...", "idle");
        const data = await API.post("/api/referencias/txt", { ruta });
        App._renderRefs(data.referencias || []);
        showToast(`✓ ${data.total} referencias cargadas`);
        setStatus(`${data.total} referencias cargadas`);
      } else {
        const input = document.createElement("input");
        input.type = "file"; input.accept = ".txt";
        input.onchange = async () => {
          if (!input.files[0]) return;
          const fd = new FormData();
          fd.append("file", input.files[0]);
          const data = await API.postForm("/api/referencias/txt-upload", fd);
          App._renderRefs(data.referencias || []);
          showToast(`✓ ${data.total} referencias cargadas`);
        };
        input.click();
      }
    } catch (e) {
      setStatus("Error: " + e.message, "error");
      showToast("Error al cargar referencias", 4000);
    }
  },

  async limpiarReferencias() {
    await API.put("/api/referencias", { referencias: [] }).catch(() => {});
    App._renderRefs([]);
    showToast("Referencias limpiadas");
  },

  async _cargarReferencias() {
    try {
      const data = await API.get("/api/referencias");
      let refs = data.referencias || [];
      if (refs.length === 0) {
        refs = State.bloques
          .filter(b => b.clasificacion === "Referencia")
          .map(b => b.contenido);
      }
      App._renderRefs(refs);
    } catch (e) {
      setStatus("Error cargando referencias: " + e.message, "error");
    }
  },

  _renderRefs(refs) {
    State._numRefs = refs.length;
    actualizarStepper();
    const lista = $("refs-list");
    const count = $("refs-count");
    if (!lista) return;
    if (count) count.textContent = `${refs.length} referencia${refs.length !== 1 ? "s" : ""}`;
    if (refs.length === 0) {
      lista.innerHTML = `
        <div class="empty-state">
          <svg viewBox="0 0 48 48" fill="none" stroke="currentColor" stroke-width="1.2"
            style="width:48px;height:48px;opacity:.25">
            <rect x="6" y="6" width="36" height="36" rx="2"/>
            <path d="M14 16h20M14 24h20M14 32h12"/>
          </svg>
          <p>Sin referencias cargadas</p>
          <p class="empty-hint">Carga un .txt o se detectan automáticamente del PDF</p>
        </div>`;
      return;
    }
    lista.innerHTML = refs.map((r, i) => `
      <div class="ref-item">
        <span class="ref-num">${i + 1}.</span>
        <span>${esc(r)}</span>
      </div>
    `).join("");
  },


  // ══════════════════════════════════════════════════════════════════════════
  // FIGURAS
  // ══════════════════════════════════════════════════════════════════════════

  async agregarFiguras() {
    try {
      if (window.pywebview) {
        const rutas = await window.pywebview.api.abrir_imagenes();
        if (!rutas || rutas.length === 0) return;
        setStatus("Agregando figuras...", "idle");
        const data = await API.post("/api/figuras/agregar-rutas", { rutas });
        App._renderFiguras(data.figuras || []);
        showToast(`✓ ${data.agregadas} figura(s) agregadas`);
        setStatus(`${data.agregadas} figura(s) agregadas`);
      } else {
        const input = document.createElement("input");
        input.type = "file"; input.accept = "image/*"; input.multiple = true;
        input.onchange = async () => {
          if (!input.files.length) return;
          const fd = new FormData();
          for (const f of input.files) fd.append("files", f);
          const data = await API.postForm("/api/figuras/agregar-upload", fd);
          App._renderFiguras(data.figuras || []);
          showToast(`✓ ${data.agregadas} figura(s) agregadas`);
        };
        input.click();
      }
    } catch (e) {
      setStatus("Error: " + e.message, "error");
      showToast("Error al agregar figuras", 4000);
    }
  },

  async limpiarFiguras() {
    await API.put("/api/figuras", { figuras: [] }).catch(() => {});
    App._renderFiguras([]);
    showToast("Figuras limpiadas");
  },

  async _cargarFiguras() {
    try {
      const data = await API.get("/api/figuras");
      App._renderFiguras(data.figuras || []);
    } catch (e) {
      setStatus("Error cargando figuras: " + e.message, "error");
    }
  },

  _renderFiguras(figs) {
    State._numFiguras    = figs.length;
    State._figurasSinPie = figs.filter(f => !f.pie?.trim()).length;
    actualizarStepper();
    const grid  = $("figuras-grid");
    const count = $("figs-count");
    if (!grid) return;
    if (count) count.textContent = `${figs.length} figura${figs.length !== 1 ? "s" : ""}`;

    if (figs.length === 0) {
      grid.innerHTML = `
        <div class="empty-state">
          <svg viewBox="0 0 48 48" fill="none" stroke="currentColor" stroke-width="1.2"
            style="width:48px;height:48px;opacity:.25">
            <rect x="4" y="8" width="40" height="32" rx="2"/>
            <path d="M4 32l10-10 8 8 8-10 10 12"/><circle cx="16" cy="20" r="4"/>
          </svg>
          <p>Sin figuras</p>
          <p class="empty-hint">Haz clic en "+ Agregar imagen" o carga un PDF</p>
        </div>`;
      return;
    }

    grid.innerHTML = figs.map((f, i) => {
      const pie   = esc(f.pie   || "");
      const ancla = esc(f.ancla || "");
      const nombre = esc(f.ruta ? f.ruta.split(/[\\/]/).pop() : `Figura ${i + 1}`);
      const imgHtml = f.img_b64
        ? `<img class="figura-thumb" src="data:${f.img_mime || "image/png"};base64,${f.img_b64}" alt="Figura ${i + 1}">`
        : `<div class="figura-thumb-placeholder">
             <svg viewBox="0 0 32 32" fill="none" stroke="currentColor" stroke-width="1"
               style="width:28px;height:28px;opacity:.4">
               <rect x="2" y="4" width="28" height="22" rx="1"/>
               <path d="M2 20l7-7 5 5 5-6 7 8"/>
             </svg>
           </div>`;
      return `
        <div class="figura-row">
          ${imgHtml}
          <div class="figura-body-edit">
            <div class="figura-num-label">Figura ${i + 1}</div>
            <div class="figura-filename">${nombre}</div>
            <input class="figura-input" type="text" value="${pie}"
              placeholder="Pie de figura…"
              onblur="App._syncFigura(${i}, 'pie', this.value)" />
            <div class="figura-ancla-label">📍 Párrafo donde va la figura:</div>
            <input class="figura-input" type="text" value="${ancla}"
              placeholder='Ej: "Se muestra en la Figura 1A-B."'
              onblur="App._syncFigura(${i}, 'ancla', this.value)" />
          </div>
          <button class="figura-del" onclick="App._eliminarFigura(${i})" title="Eliminar">✕</button>
        </div>`;
    }).join("");
  },

  async _syncFigura(idx, campo, valor) {
    try {
      await fetch(`/api/figuras/${idx}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ idx, [campo]: valor }),
      });
      _dirty(true);
    } catch (_) {}
  },

  async _eliminarFigura(idx) {
    try {
      const data = await fetch(`/api/figuras/${idx}`, { method: "DELETE" }).then(r => r.json());
      _dirty(true);
      App._renderFiguras(data.figuras || []);
    } catch (e) {
      setStatus("Error: " + e.message, "error");
    }
  },


  // ══════════════════════════════════════════════════════════════════════════
  // TABLAS
  // ══════════════════════════════════════════════════════════════════════════

  async agregarTablas() {
    try {
      if (window.pywebview) {
        const rutas = await window.pywebview.api.abrir_excels_multiples();
        if (!rutas || rutas.length === 0) return;
        setStatus("Importando tablas...", "idle");
        const data = await API.post("/api/tablas/agregar-rutas", { rutas });
        App._renderTablas(data.tablas || []);
        showToast(`✓ ${data.agregadas} tabla(s) importadas`);
        setStatus(`${data.agregadas} tabla(s) importadas`);
      } else {
        const input = document.createElement("input");
        input.type = "file"; input.accept = ".xlsx,.xls"; input.multiple = true;
        input.onchange = async () => {
          if (!input.files.length) return;
          const fd = new FormData();
          for (const f of input.files) fd.append("files", f);
          const data = await API.postForm("/api/tablas/agregar-upload", fd);
          App._renderTablas(data.tablas || []);
          showToast(`✓ ${data.agregadas} tabla(s) importadas`);
        };
        input.click();
      }
    } catch (e) {
      setStatus("Error: " + e.message, "error");
      showToast("Error al importar tablas", 4000);
    }
  },

  async limpiarTablas() {
    await API.put("/api/tablas", { tablas: [] }).catch(() => {});
    App._renderTablas([]);
    showToast("Tablas limpiadas");
  },

  async _cargarTablas() {
    try {
      const data = await API.get("/api/tablas");
      App._renderTablas(data.tablas || []);
    } catch (e) {
      setStatus("Error cargando tablas: " + e.message, "error");
    }
  },

  _renderTablas(tabs) {
    State._numTablas = tabs.length;
    actualizarStepper();
    const lista = $("tablas-list");
    const count = $("tabs-count");
    if (!lista) return;
    if (count) count.textContent = `${tabs.length} tabla${tabs.length !== 1 ? "s" : ""}`;

    if (tabs.length === 0) {
      lista.innerHTML = `
        <div class="empty-state">
          <svg viewBox="0 0 48 48" fill="none" stroke="currentColor" stroke-width="1.2"
            style="width:48px;height:48px;opacity:.25">
            <rect x="4" y="6" width="40" height="36" rx="2"/>
            <path d="M4 18h40M4 30h40M20 18v24M32 18v24"/>
          </svg>
          <p>Sin tablas</p>
          <p class="empty-hint">Haz clic en "+ Agregar Excel" o carga un PDF</p>
        </div>`;
      return;
    }

    lista.innerHTML = tabs.map((t, i) => {
      const rotulo      = esc(t.rotulo || "");
      const descripcion = esc(t.descripcion || t.titulo || "");   // compat
      const ancla   = esc(t.ancla  || "");
      const archivo = esc(t.ruta ? t.ruta.split(/[\\/]/).pop() : "");
      const hoja    = esc(t.hoja || "");
      const previewHTML = App._miniTablaHTML(t.preview);
      const subtitle = archivo ? (hoja ? `${archivo} › ${hoja}` : archivo) : `Tabla ${i + 1}`;
      const hayEsq  = i < tabs.length - 1;
      const sugiere = !!t.sugiere_unir_siguiente;
      const unirHTML = hayEsq ? `
        <div class="tabla-unir">
          ${sugiere ? `<div class="tabla-unir-hint">🔗 Parece la continuación de una tabla partida por salto de página.</div>` : ""}
          <button class="tabla-btn-unir${sugiere ? " sugerido" : ""}" onclick="App._unirSiguiente(${i})">
            ⬍ Unir con la tabla siguiente
          </button>
        </div>` : "";
      return `
        <div class="tabla-row">
          <div class="tabla-row-header">
            <div class="tabla-row-icon">
              <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6">
                <rect x="2" y="3" width="16" height="14" rx="1"/>
                <path d="M2 8h16M2 13h16M7 8v9M13 8v9"/>
              </svg>
            </div>
            <div class="tabla-row-meta">
              <div class="tabla-row-num">Tabla ${i + 1}</div>
              <div class="tabla-row-file">${subtitle}</div>
            </div>
            <button class="tabla-btn-excel" onclick="App._editarEnExcel(${i})"
              title="Abrir el archivo en Excel para editarlo">
              <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6">
                <path d="M12 3H5a1 1 0 00-1 1v12a1 1 0 001 1h10a1 1 0 001-1V7z"/>
                <path d="M12 3v4h4M7 11l3 3M10 11l-3 3"/>
              </svg>
              Editar en Excel
            </button>
            <button class="tabla-row-del" onclick="App._eliminarTabla(${i})" title="Eliminar">✕</button>
          </div>
          <div class="tabla-row-body">
            <div class="figura-ancla-label">🏷️ Rótulo (etiqueta):</div>
            <input class="tabla-input" type="text" value="${rotulo}"
              placeholder='Ej: "Tabla 1"'
              onblur="App._syncTabla(${i}, 'rotulo', this.value)" />
            <div class="figura-ancla-label">📝 Descripción (leyenda):</div>
            <input class="tabla-input" type="text" value="${descripcion}"
              placeholder="Ej: Coeficientes de correlación de la sección La Joya…"
              onblur="App._syncTabla(${i}, 'descripcion', this.value)" />
            <div class="figura-ancla-label">📍 Párrafo donde va la tabla:</div>
            <input class="tabla-input" type="text" value="${ancla}"
              placeholder='Ej: "...la Dra. Elena Centeno (Tabla 1)."'
              onblur="App._syncTabla(${i}, 'ancla', this.value)" />
            ${previewHTML}
            ${unirHTML}
          </div>
        </div>`;
    }).join("");
  },

  // Construye el HTML de la mini-tabla de vista previa a partir del objeto
  // `preview` que devuelve el backend ({ok, filas, n_filas, n_cols, ...}).
  _miniTablaHTML(preview) {
    if (!preview || !preview.ok) {
      const msg = preview && preview.error ? preview.error : "Sin datos";
      return `<div class="tabla-preview-empty">⚠ ${esc(msg)}</div>`;
    }
    const filas = preview.filas || [];
    if (filas.length === 0) return `<div class="tabla-preview-empty">Tabla vacía</div>`;

    const [head, ...body] = filas;
    const th   = head.map(c => `<th>${esc(c)}</th>`).join("");
    const rows = body.map(f =>
      `<tr>${f.map(c => `<td>${esc(c)}</td>`).join("")}</tr>`).join("");

    const nf = preview.n_filas || filas.length;
    const nc = preview.n_cols  || (head ? head.length : 0);
    const nota = (nf || nc)
      ? `<div class="tabla-preview-nota">${nf} fila${nf !== 1 ? "s" : ""} × ${nc} columna${nc !== 1 ? "s" : ""}</div>`
      : "";

    return `
      <div class="tabla-preview">
        <table class="tabla-mini">
          <thead><tr>${th}</tr></thead>
          <tbody>${rows}</tbody>
        </table>
      </div>${nota}`;
  },

  async _syncTabla(idx, campo, valor) {
    try {
      await fetch(`/api/tablas/${idx}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ idx, [campo]: valor }),
      });
      _dirty(true);
    } catch (_) {}
  },

  async _eliminarTabla(idx) {
    try {
      const data = await fetch(`/api/tablas/${idx}`, { method: "DELETE" }).then(r => r.json());
      _dirty(true);
      App._renderTablas(data.tablas || []);
    } catch (e) {
      setStatus("Error: " + e.message, "error");
    }
  },

  // RF-29 — Abre el .xlsx de la tabla en Excel y arma la detección de cambios.
  async _editarEnExcel(idx) {
    try {
      // Captura el mtime actual para saber si hubo edición al regresar.
      const antes = await API.get("/api/tablas");
      const t = (antes.tablas || [])[idx];
      App._editWatch = { idx, mtime: (t && t.preview) ? t.preview.mtime || 0 : 0 };

      await API.post(`/api/tablas/${idx}/abrir-excel`, {});
      showToast("Abriendo en Excel… edita, guarda y vuelve a la app");
      setStatus(`Editando Tabla ${idx + 1} en Excel…`, "idle");
    } catch (e) {
      App._editWatch = null;
      showToast("No se pudo abrir en Excel: " + (e.message || ""), 4000);
    }
  },

  // RF-28 — Une la tabla idx con la siguiente en un .xlsx combinado.
  async _unirSiguiente(idx) {
    const ok = await _modalConfirmar({
      titulo:  "Unir tablas",
      mensaje: `Se unirán la Tabla ${idx + 1} y la Tabla ${idx + 2} en una sola. ` +
               `Si la segunda repite el encabezado, se quitará automáticamente. ` +
               `El resultado queda en un Excel que podrás seguir editando.`,
      btnOk:   "Unir",
    });
    if (!ok) return;
    try {
      setStatus("Uniendo tablas…", "idle");
      const data = await API.post(`/api/tablas/${idx}/unir-siguiente`, {});
      App._renderTablas(data.tablas || []);
      showToast(data.encabezado_quitado
        ? "✓ Tablas unidas (encabezado repetido eliminado)"
        : "✓ Tablas unidas");
      setStatus("Tablas unidas");
    } catch (e) {
      showToast("No se pudieron unir: " + (e.message || ""), 4000);
      setStatus("Error al unir tablas", "error");
    }
  },

  // Al volver el foco a la app: relee la tabla editada y avisa si cambió.
  async _revisarEdicionExcel() {
    const w = App._editWatch;
    if (!w) return;
    App._editWatch = null;
    try {
      const data = await API.get("/api/tablas");
      App._renderTablas(data.tablas || []);
      const t = (data.tablas || [])[w.idx];
      const nuevoM = (t && t.preview) ? t.preview.mtime || 0 : 0;
      if (nuevoM && w.mtime && nuevoM > w.mtime) {
        showToast(`✓ Tabla ${w.idx + 1} actualizada`);
        setStatus("Vista previa de la tabla actualizada");
      } else {
        setStatus("Sin cambios detectados en la tabla");
      }
    } catch (_) {}
  },

  // ── Proyecto (RF-04): menú Archivo, guardar/abrir/autoguardar ───────────────
  _toggleMenu(ev, id) {
    ev.stopPropagation();
    const dd = document.getElementById(id);
    const abrir = dd && dd.style.display !== "block";
    App._cerrarMenus();
    if (dd && abrir) dd.style.display = "block";
  },

  _cerrarMenus() {
    document.querySelectorAll(".menu-dropdown").forEach(d => (d.style.display = "none"));
  },

  // RF-03 — "Nuevo" abre otra PESTAÑA (proyecto vacío), sin cerrar el actual.
  async proyectoNuevo() {
    App._cerrarMenus();
    try {
      await API.post("/api/proyectos/nuevo", {});
      location.reload();   // re-inicializa la interfaz sobre la pestaña nueva
    } catch (e) {
      showToast("No se pudo crear el proyecto: " + (e.message || ""), 4000);
    }
  },

  async proyectoAbrir() {
    App._cerrarMenus();
    const ruta = window.pywebview?.api?.abrir_pmz
      ? await window.pywebview.api.abrir_pmz()
      : null;
    if (!ruta) return;   // cancelado
    try {
      setStatus("Abriendo proyecto…", "idle");
      // El backend lo carga en una pestaña nueva si la actual ya tiene contenido.
      await API.post("/api/proyecto/abrir", { ruta });
      localStorage.setItem("proyectoRuta", ruta);
      showToast("✓ Proyecto abierto");
      location.reload();   // re-inicializa la interfaz desde el estado cargado
    } catch (e) {
      showToast("No se pudo abrir el proyecto: " + (e.message || ""), 4000);
      setStatus("Error al abrir proyecto", "error");
    }
  },

  // ── RF-03 — Pestañas (varios proyectos a la vez) ────────────────────────────
  _renderTabs() {
    const cont = $("tabbar-tabs");
    if (!cont) return;
    const ps = State.proyectos || [];
    cont.innerHTML = ps.map(p => {
      const activo = p.id === State.activo;
      // El punto de "sin guardar" del activo se toma del estado vivo.
      const sucio = activo ? !!(State.tienePDF && State.sinGuardar)
                           : (p.tiene_contenido && p.sin_guardar);
      const titulo = _escHtml(p.titulo || "Proyecto sin título");
      const cls = "tab" + (activo ? " tab--activa" : "") + (sucio ? " tab--sucia" : "");
      const rid = p.id.replace(/'/g, "\\'");
      return `
        <div class="${cls}" title="${titulo}" onclick="App.proyectoActivar('${rid}')">
          ${sucio ? '<span class="tab-punto" title="Cambios sin guardar"></span>' : ""}
          <span class="tab-titulo">${titulo}</span>
          <button class="tab-cerrar" title="Cerrar pestaña"
                  onclick="App.proyectoCerrar(event, '${rid}')">✕</button>
        </div>`;
    }).join("");
  },

  async _cargarProyectos() {
    try {
      const data = await API.get("/api/proyectos");
      State.proyectos = data.proyectos || [];
      State.activo    = data.activo || "";
      App._renderTabs();
    } catch (e) {
      console.error("[proyectos]", e);
    }
  },

  async proyectoActivar(id) {
    if (id === State.activo) return;   // ya activa
    try {
      setStatus("Cambiando de proyecto…", "idle");
      await API.post("/api/proyectos/activar", { id });
      location.reload();               // recarga toda la vista sobre la pestaña elegida
    } catch (e) {
      showToast("No se pudo cambiar de proyecto: " + (e.message || ""), 4000);
    }
  },

  async proyectoCerrar(ev, id) {
    ev.stopPropagation();
    const p = (State.proyectos || []).find(p => p.id === id);
    const esActiva = id === State.activo;
    const sucia = esActiva ? !!(State.tienePDF && State.sinGuardar)
                           : (p && p.tiene_contenido && p.sin_guardar);
    if (sucia) {
      const ok = await _modalConfirmar({
        titulo:  "Cerrar proyecto",
        mensaje: "Esta pestaña tiene cambios sin guardar. Si la cierras, se perderán. ¿Cerrar de todas formas?",
        btnOk:   "Cerrar sin guardar", peligro: true,
      });
      if (!ok) return;
    }
    try {
      await API.delete(`/api/proyectos/${id}`);
      if (esActiva) {
        location.reload();             // cambió el proyecto activo → recargar vista
      } else {
        await App._cargarProyectos();  // solo actualizar la barra de pestañas
        try { window.pywebview?.api?.set_sin_guardar?.(_algunaSinGuardar()); } catch (_) {}
      }
    } catch (e) {
      showToast("No se pudo cerrar la pestaña: " + (e.message || ""), 4000);
    }
  },

  async proyectoGuardar() {
    App._cerrarMenus();
    if (State.proyectoRuta) await App._guardarEn(State.proyectoRuta);
    else await App.proyectoGuardarComo();
  },

  async proyectoGuardarComo() {
    App._cerrarMenus();
    const base = ((State.pdfInfo && State.pdfInfo.nombre) || "proyecto")
      .replace(/\.[^.]+$/, "");
    const ruta = window.pywebview?.api?.guardar_pmz
      ? await window.pywebview.api.guardar_pmz(base + ".pmz")
      : null;
    if (!ruta) return;
    await App._guardarEn(ruta);
  },

  async _guardarEn(ruta) {
    try {
      setStatus("Guardando proyecto…", "idle");
      const data = await API.post("/api/proyecto/guardar", { ruta });
      State.proyectoRuta = data.ruta;
      if (data.proyectos) { State.proyectos = data.proyectos; State.activo = data.activo; }
      _dirty(false);                            // también re-renderiza las pestañas
      localStorage.setItem("proyectoRuta", data.ruta);
      App._actualizarTituloProyecto();
      App._renderTabs();                        // el título de la pestaña ya cambió
      showToast("✓ Proyecto guardado");
      setStatus("Proyecto guardado");
    } catch (e) {
      showToast("No se pudo guardar: " + (e.message || ""), 4000);
      setStatus("Error al guardar proyecto", "error");
    }
  },

  // ── Título del proyecto en la barra superior (RF-04 · visual) ───────────────
  _tituloProyecto() {
    if (!State.proyectoRuta) return "Proyecto sin título";
    return State.proyectoRuta.split(/[\\/]/).pop().replace(/\.pmz$/i, "");
  },

  _actualizarTituloProyecto() {
    const el = $("proyecto-titulo");
    if (!el) return;
    const t = App._tituloProyecto();
    el.textContent = t;
    el.classList.toggle("proyecto-titulo--sin", !State.proyectoRuta);
    el.title = "Proyecto actual: " + t;
  },

  // ── Aviso de cambios sin guardar al cerrar (RF-04) ──────────────────────────
  // La llama main.py (evento closing de la ventana) vía window.pywebview.
  _hayCambiosSinGuardar() {
    return _algunaSinGuardar();
  },

  // La llama main.py (hilo aparte) al intentar cerrar la ventana.
  async _alIntentarCerrar() {
    if (!App._hayCambiosSinGuardar()) { App._cerrarApp(); return; }
    await App._confirmarCierre();
  },

  _cerrarApp() {
    if (window.pywebview?.api?.cerrar_confirmado)
      window.pywebview.api.cerrar_confirmado();
  },

  async _confirmarCierre() {
    const eleccion = await App._modalCierre();
    if (eleccion === "cancelar") return;
    if (eleccion === "guardar") {
      await App.proyectoGuardar();
      if (State.sinGuardar) return;   // el usuario canceló el diálogo de guardar
    }
    App._cerrarApp();
  },

  _modalCierre() {
    return new Promise(resolve => {
      let overlay = document.getElementById("modal-cierre-overlay");
      if (!overlay) {
        overlay = document.createElement("div");
        overlay.id = "modal-cierre-overlay";
        overlay.className = "modal-overlay";
        overlay.innerHTML = `
          <div class="modal-confirmar" role="dialog" aria-modal="true">
            <h3 class="modal-confirmar-titulo">Cambios sin guardar</h3>
            <p class="modal-confirmar-msg">Tienes cambios sin guardar en este proyecto. ¿Qué quieres hacer antes de cerrar?</p>
            <div class="modal-confirmar-btns modal-cierre-btns">
              <button class="btn btn--ghost"   id="cierre-cancelar">Cancelar</button>
              <button class="btn btn--danger"  id="cierre-sin">Cerrar sin guardar</button>
              <button class="btn btn--primary" id="cierre-guardar">Guardar y cerrar</button>
            </div>
          </div>`;
        document.body.appendChild(overlay);
      }
      overlay.style.display = "flex";
      const cerrar = (val) => {
        overlay.style.display = "none";
        document.removeEventListener("keydown", onKey);
        resolve(val);
      };
      const onKey = (e) => { if (e.key === "Escape") cerrar("cancelar"); };
      document.getElementById("cierre-cancelar").onclick = () => cerrar("cancelar");
      document.getElementById("cierre-sin").onclick      = () => cerrar("sin-guardar");
      document.getElementById("cierre-guardar").onclick  = () => cerrar("guardar");
      overlay.onclick = (e) => { if (e.target === overlay) cerrar("cancelar"); };
      document.addEventListener("keydown", onKey);
      document.getElementById("cierre-guardar").focus();
    });
  },

  async _ofrecerReabrirUltimo(ruta) {
    const nombre = ruta.split(/[\\/]/).pop();
    const ok = await _modalConfirmar({
      titulo:  "Reabrir proyecto",
      mensaje: `¿Quieres reabrir el último proyecto guardado «${nombre}»?`,
      btnOk:   "Reabrir",
    });
    if (!ok) return;
    try {
      await API.post("/api/proyecto/abrir", { ruta });
      location.reload();
    } catch (e) {
      showToast("No se pudo reabrir: " + (e.message || ""), 4000);
      localStorage.removeItem("proyectoRuta");   // la ruta ya no es válida
    }
  },


  // ══════════════════════════════════════════════════════════════════════════
  // EXPORTAR
  // ══════════════════════════════════════════════════════════════════════════

  async exportar(formato) {
    if (!State.tienePDF) {
      showToast("Primero carga un PDF");
      return;
    }

    // Guardar autores antes de exportar
    App._leerInputsAutores();
    await App._pushAutores();

    try {
      if (window.pywebview) {
        // ── Producción: usa carpeta predeterminada si existe, si no abre diálogo ──
        const carpeta = localStorage.getItem("carpetaSalida");
        let ruta;

        if (carpeta) {
          // Construir ruta directamente sin diálogo
          const ext      = formato === "xml" ? "xml" : formato;
          const nombre   = `articulo.${ext}`;
          const sep      = carpeta.includes("/") ? "/" : "\\";
          ruta = carpeta.replace(/[/\\]$/, "") + sep + nombre;
        } else {
          // Sin carpeta configurada → diálogo nativo como antes
          const fn = { html: "guardar_html", xml: "guardar_xml", epub: "guardar_epub" }[formato];
          ruta = await window.pywebview.api[fn]();
          if (!ruta) return;   // usuario canceló
        }

        setStatus(`Exportando ${formato.toUpperCase()}...`, "idle");
        showLoading(true);
        await API.post(`/api/exportar/${formato}`, { formato, ruta_destino: ruta });
        showLoading(false);
        setStatus(`${formato.toUpperCase()} guardado`);
        showToast(`✓ ${formato.toUpperCase()} guardado en: ${ruta}`);

      } else {
        // ── Desarrollo: descarga en el browser ──
        setStatus(`Exportando ${formato.toUpperCase()}...`, "idle");
        showLoading(true);
        const blob = await API.postBlob(`/api/exportar/${formato}/preview`);
        const url  = URL.createObjectURL(blob);
        const a    = document.createElement("a");
        a.href     = url;
        a.download = `articulo.${formato === "xml" ? "xml" : formato}`;
        a.click();
        URL.revokeObjectURL(url);
        showLoading(false);
        setStatus(`${formato.toUpperCase()} descargado`);
        showToast(`✓ ${formato.toUpperCase()} descargado`);
      }
    } catch (e) {
      showLoading(false);
      setStatus("Error al exportar: " + e.message, "error");
      showToast("Error al exportar: " + e.message, 4500);
    }
  },

  // ══════════════════════════════════════════════════════════════════════════
  // VALIDAR XML JATS
  // ══════════════════════════════════════════════════════════════════════════

  cambiarPDF() {
    // Mostrar modal de confirmación en lugar del feo window.confirm
    const modal = $("modal-confirmar-pdf");
    if (modal) modal.style.display = "flex";
  },

  _cerrarConfirmarPDF() {
    const modal = $("modal-confirmar-pdf");
    if (modal) modal.style.display = "none";
  },

  async _confirmarCambiarPDF() {
    App._cerrarConfirmarPDF();
    showLoading(true);

    try {
      try {
        // Reset total en el servidor: bloques, metadatos, autores,
        // afiliaciones, referencias, figuras y tablas — todo el artículo
        // anterior se descarta, se empieza desde cero.
        await API.delete("/api/estado");
      } catch (_) { /* aunque falle la llamada, igual limpiamos localmente */ }

      // ── Reset total del estado local ────────────────────────────────────
      State.bloques      = [];
      State.metadatos    = {};
      State.autores      = [];
      State.tienePDF     = false;
      State._afilTxt     = "";
      State._numRefs     = 0;
      State._numFiguras  = 0;
      State._figurasSinPie = 0;
      State._numTablas   = 0;

      // ── Panel PDF ────────────────────────────────────────────────────────
      const dz = $("dropzone");
      const bc = $("bloques-container");
      if (dz) dz.style.display = "flex";
      if (bc) { bc.style.display = "none"; bc.innerHTML = ""; }

      $("btn-leyenda")     && ($("btn-leyenda").style.display     = "none");
      $("filtro-grupo")    && ($("filtro-grupo").style.display    = "none");
      $("btn-cambiar-pdf") && ($("btn-cambiar-pdf").style.display = "none");

      const setInfo = (id, v) => { const el = $(id); if (el) el.textContent = v || "—"; };
      setInfo("info-nombre",  "—");
      setInfo("info-paginas", "—");
      setInfo("info-tamanio", "—");
      const badge = $("info-estado");
      if (badge) { badge.textContent = "Sin cargar"; badge.className = "badge badge--gray"; }

      // ── Resto de paneles: repintar todos en blanco ──────────────────────
      App._renderMetadatos();
      App._renderAutores();
      const ta = $("afiliaciones-txt");
      if (ta) ta.value = "";
      App._renderRefs([]);
      App._renderFiguras([]);
      App._renderTablas([]);

      actualizarStepper();
      setStatus("Listo para cargar nuevo PDF");
      showToast("Sesión reiniciada — carga el nuevo archivo");

      // Volver a mostrar historial si hay entradas
      Historial.renderizar();
    } catch (e) {
      console.error("[_confirmarCambiarPDF]", e);
      setStatus("Error al reiniciar: " + e.message, "error");
      showToast("Algo falló al reiniciar — revisa la consola", 4000);
    } finally {
      showLoading(false);
    }
  },

  // ══════════════════════════════════════════════════════════════════════════
  // VISTA PREVIA HTML
  // ══════════════════════════════════════════════════════════════════════════

  async verPreview() {
    if (!State.tienePDF) {
      showToast("Primero carga un PDF");
      return;
    }

    // Guardar autores antes de generar
    App._leerInputsAutores();
    await App._pushAutores();

    // Abrir modal con spinner
    const modal   = $("modal-preview");
    const iframe  = $("preview-iframe");
    const loading = $("preview-loading");
    const nombre  = $("preview-nombre");
    if (!modal) return;

    modal.style.display   = "flex";
    loading.style.display = "flex";
    iframe.style.display  = "none";
    if (nombre) nombre.textContent = State.pdfInfo?.nombre || "";

    // Cuando el iframe termine de cargar, ocultar spinner
    iframe.onload = () => {
      loading.style.display = "none";
      iframe.style.display  = "block";
    };

    // Cargar la URL directamente — el servidor genera el HTML al vuelo
    iframe.src = `http://127.0.0.1:8765/api/exportar/html/vista-previa`;
  },

  cerrarPreview() {
    const modal  = $("modal-preview");
    const iframe = $("preview-iframe");
    if (modal)  modal.style.display  = "none";
    if (iframe) { iframe.src = "about:blank"; iframe.style.display = "none"; }
  },

  async validarXML() {
    if (!State.tienePDF) {
      showToast("Primero carga un PDF");
      return;
    }
    App._leerInputsAutores();
    await App._pushAutores();

    setStatus("Validando XML...", "idle");
    showLoading(true);
    try {
      const resultado = await API.post("/api/validar/xml", {});
      showLoading(false);
      App._mostrarResultadoValidacion(resultado);
    } catch (e) {
      showLoading(false);
      setStatus("Error al validar: " + e.message, "error");
      showToast("Error al validar XML", 4000);
    }
  },

  _mostrarResultadoValidacion(r) {
    setStatus(r.valido ? "XML válido ✅" : "XML con errores ❌", r.valido ? "ok" : "error");

    const errores      = r.errores      || [];
    const advertencias = r.advertencias || [];

    const filaHTML = (item, tipo) => {
      const color = tipo === "error" ? "#EF4444" : "#D97706";
      const icono = tipo === "error" ? "✕" : "⚠";
      const linea = item.linea > 0
        ? `<span style="color:#9AA3B5;font-size:10px;margin-left:6px">línea ${item.linea}</span>`
        : "";
      return `
        <div style="display:flex;align-items:flex-start;gap:10px;padding:8px 0;
          border-bottom:1px solid #E4E9F0;">
          <span style="color:${color};font-weight:700;font-size:13px;flex-shrink:0">${icono}</span>
          <span style="font-size:12px;color:#1A2236;flex:1">${esc(item.mensaje)}${linea}</span>
        </div>`;
    };

    let cuerpo = `
      <div style="margin-bottom:16px;padding:12px 16px;border-radius:8px;
        background:${r.valido ? "#DCFCE7" : "#FEE2E2"};
        color:${r.valido ? "#16A34A" : "#EF4444"};
        font-size:13px;font-weight:600;">
        ${esc(r.resumen)}
      </div>`;

    if (errores.length > 0) {
      cuerpo += `<div style="font-size:11px;font-weight:700;color:#EF4444;
        text-transform:uppercase;letter-spacing:.04em;margin-bottom:6px;">
        Errores (${errores.length})</div>`;
      cuerpo += errores.map(e => filaHTML(e, "error")).join("");
    }

    if (advertencias.length > 0) {
      cuerpo += `<div style="font-size:11px;font-weight:700;color:#D97706;
        text-transform:uppercase;letter-spacing:.04em;
        margin-top:${errores.length ? "16px" : "0"};margin-bottom:6px;">
        Advertencias (${advertencias.length})</div>`;
      cuerpo += advertencias.map(a => filaHTML(a, "advertencia")).join("");
    }

    if (errores.length === 0 && advertencias.length === 0) {
      cuerpo += `<div style="text-align:center;padding:24px;color:#9AA3B5;font-size:13px;">
        No se encontraron problemas.</div>`;
    }

    // Reutiliza el modal de leyenda que ya existe en el HTML
    const contenido = $("leyenda-contenido");
    const titulo    = document.querySelector("#modal-leyenda .modal-header h3");
    if (contenido) contenido.innerHTML = cuerpo;
    if (titulo)    titulo.textContent  = "Validación JATS / SciELO";
    const modal = $("modal-leyenda");
    if (modal) modal.style.display = "flex";
  },

};   // fin App


// ─────────────────────────────────────────────────────────────────────────────
// Inicialización
// ─────────────────────────────────────────────────────────────────────────────
async function init() {
  try {
    // 1. Cargar configuración (opciones de clasificación + colores)
    State.config = await API.get("/api/config");

    // 2. Ver si ya hay un PDF cargado en el servidor (por recarga de ventana)
    const estado = await API.get("/api/estado");

    if (estado.tiene_pdf) {
      State.tienePDF = true;
      const bloqData = await API.get("/api/bloques");
      State.bloques  = bloqData.bloques || [];

      const info = estado.pdf_info || {};
      const sv = (id, v) => { const el=$(id); if(el) el.textContent = v||"—"; };
      sv("info-nombre",  info.nombre);
      sv("info-paginas", info.paginas);
      sv("info-tamanio", info.tamanio);
      const badge = $("info-estado");
      if (badge) { badge.textContent = "Cargado"; badge.className = "badge badge--green"; }

      const dz = $("dropzone");
      const bc = $("bloques-container");
      if (dz) dz.style.display = "none";
      if (bc) bc.style.display = "block";
      $("btn-leyenda")      && ($("btn-leyenda").style.display      = "inline-flex");
      $("filtro-grupo")     && ($("filtro-grupo").style.display     = "flex");
      $("btn-cambiar-pdf")  && ($("btn-cambiar-pdf").style.display  = "inline-flex");

      App._poblarFiltro();
      App._renderBloques(State.bloques);

      // 2b. Recuperar metadatos editoriales ya detectados/guardados
      try {
        const metaData = await API.get("/api/metadatos");
        State.metadatos = metaData.metadatos || {};
      } catch (_) {
        State.metadatos = {};
      }
    }

    // 3. Cargar autores
    const autData = await API.get("/api/autores");
    State.autores  = autData.autores || [];

    // 4. Cargar afiliaciones
    const afilData = await API.get("/api/afiliaciones");
    const ta = $("afiliaciones-txt");
    if (ta && afilData.texto) ta.value = afilData.texto;
    State._afilTxt = afilData.texto || "";

    // 4b. Cargar conteos para el indicador de progreso
    const refsData = await API.get("/api/referencias");
    State._numRefs = (refsData.referencias || []).length;

    const figsData = await API.get("/api/figuras");
    const figs = figsData.figuras || [];
    State._numFiguras    = figs.length;
    State._figurasSinPie = figs.filter(f => !f.pie?.trim()).length;

    const tabsData = await API.get("/api/tablas");
    State._numTablas = (tabsData.tablas || []).length;

    // 5. Mostrar sección inicial
    App.irSeccion("pdf");

    // 6. Mostrar carpeta de salida guardada (si existe)
    App._mostrarCarpeta(localStorage.getItem("carpetaSalida"));

    // 7. Renderizar historial de PDFs recientes
    Historial.renderizar();

    // 8. Proyectos/pestañas (RF-03) + estado del proyecto activo (RF-04).
    await App._cargarProyectos();
    const activo = (State.proyectos || []).find(p => p.id === State.activo);
    // La ruta del proyecto activo la dicta el backend (autoritativo por pestaña).
    State.proyectoRuta = (activo && activo.ruta) || null;
    _dirty(activo ? !!activo.sin_guardar : false);
    App._actualizarTituloProyecto();

    if (estado.tiene_pdf) {
      setStatus("Listo para continuar");
    } else {
      setStatus("Listo para comenzar");
      // Sin autoguardado: si esta pestaña está vacía, ofrecer reabrir el último
      // proyecto que guardaste (útil al reiniciar la app).
      const ultimo = localStorage.getItem("proyectoRuta");
      if (ultimo && !State.proyectoRuta) App._ofrecerReabrirUltimo(ultimo);
    }

  } catch (e) {
    setStatus("Error iniciando: " + e.message, "error");
    console.error("[init]", e);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Historial de PDFs recientes
// ─────────────────────────────────────────────────────────────────────────────
const Historial = {
  MAX: 5,
  KEY: "historialPDFs",

  cargar() {
    try { return JSON.parse(localStorage.getItem(this.KEY) || "[]"); }
    catch { return []; }
  },

  guardar(lista) {
    localStorage.setItem(this.KEY, JSON.stringify(lista));
  },

  /** Agrega una entrada {ruta, nombre, fecha} y recorta a MAX */
  agregar(ruta, nombre) {
    const lista = this.cargar().filter(e => e.ruta !== ruta); // evita duplicados
    lista.unshift({ ruta, nombre, fecha: Date.now() });
    this.guardar(lista.slice(0, this.MAX));
    this.renderizar();
  },

  limpiar() {
    this.guardar([]);
    this.renderizar();
  },

  renderizar() {
    const lista   = this.cargar();
    const wrapper = document.getElementById("historial-recientes");
    const ul      = document.getElementById("historial-lista");
    if (!wrapper || !ul) return;

    // Solo mostrar si hay entradas Y no hay PDF cargado
    if (lista.length === 0 || State.tienePDF) {
      wrapper.style.display = "none";
      return;
    }

    wrapper.style.display = "block";
    ul.innerHTML = lista.map(e => {
      const fecha = _fechaRelativa(e.fecha);
      return `
        <li class="historial-item" onclick="App._cargarPorRuta('${e.ruta.replace(/\\/g, "\\\\")}')">
          <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" class="historial-item-icon">
            <path d="M4 3h8l4 4v11a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1z"/>
            <path d="M12 3v4h4"/>
          </svg>
          <span class="historial-item-nombre">${_escHtml(e.nombre)}</span>
          <span class="historial-item-fecha">${fecha}</span>
        </li>`;
    }).join("");
  },
};

function _fechaRelativa(ts) {
  const diff = Date.now() - ts;
  const min  = Math.floor(diff / 60000);
  const hrs  = Math.floor(diff / 3600000);
  const dias = Math.floor(diff / 86400000);
  if (min  <  1) return "Justo ahora";
  if (min  < 60) return `Hace ${min} min`;
  if (hrs  < 24) return `Hace ${hrs} h`;
  if (dias <  7) return `Hace ${dias} día${dias > 1 ? "s" : ""}`;
  return new Date(ts).toLocaleDateString("es-MX", { day:"2-digit", month:"short" });
}

function _escHtml(s) {
  return s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

// Exponer limpiar en App
App.limpiarHistorial = () => { Historial.limpiar(); showToast("Historial borrado"); };
const Tema = {
  aplicar(modo) {
    document.documentElement.setAttribute("data-theme", modo);
    localStorage.setItem("tema", modo);
    // Marcar botón activo en la sección de config
    const btnClaro  = document.getElementById("btnTemaClaro");
    const btnOscuro = document.getElementById("btnTemaOscuro");
    if (btnClaro && btnOscuro) {
      btnClaro.classList.toggle("config-theme-btn--active",  modo === "light");
      btnOscuro.classList.toggle("config-theme-btn--active", modo === "dark");
    }
  },

  init() {
    const guardado = localStorage.getItem("tema");
    if (guardado === "dark") {
      this.aplicar("dark");
    } else {
      this.aplicar("light");
    }
  },

  toggle() {
    const actual = document.documentElement.getAttribute("data-theme");
    this.aplicar(actual === "dark" ? "light" : "dark");
  },
};

// Exponer en App
App.toggleTema = () => Tema.toggle();
App.setTema    = (modo) => Tema.aplicar(modo);

// ── Carpeta de salida ─────────────────────────────────────────────────────
App.seleccionarCarpeta = async () => {
  if (!window.pywebview) {
    showToast("Solo disponible en la app de escritorio");
    return;
  }
  const carpeta = await window.pywebview.api.seleccionar_carpeta();
  if (!carpeta) return;
  localStorage.setItem("carpetaSalida", carpeta);
  App._mostrarCarpeta(carpeta);
  showToast("✓ Carpeta de salida guardada");
};

App.limpiarCarpeta = () => {
  localStorage.removeItem("carpetaSalida");
  App._mostrarCarpeta(null);
  showToast("Carpeta eliminada — se pedirá ubicación al exportar");
};

App._mostrarCarpeta = (ruta) => {
  const el     = document.getElementById("carpeta-salida-ruta");
  const btnDel = document.getElementById("carpeta-salida-borrar");
  if (!el) return;
  if (ruta) {
    el.textContent = ruta;
    el.classList.remove("carpeta-vacia");
    if (btnDel) btnDel.style.display = "inline-flex";
  } else {
    el.textContent = "Sin carpeta configurada — se pedirá al exportar";
    el.classList.add("carpeta-vacia");
    if (btnDel) btnDel.style.display = "none";
  }
};

// Inicializar tema antes de que cargue el resto para evitar flash
Tema.init();

// Ocultar el botón flotante de "dividir bloque" si se hace clic fuera de él
// y fuera de cualquier textarea de bloque.
document.addEventListener("mousedown", (e) => {
  const boton = document.getElementById("btn-flotante-dividir");
  if (!boton || boton.style.display === "none") return;
  const dentroDeTextarea = e.target.classList?.contains("bloque-texto");
  const esElBoton = e.target === boton;
  if (!dentroDeTextarea && !esElBoton) {
    boton.style.display = "none";
  }
});

// Exponer App en window para que main.py (evaluate_js) pueda invocarlo:
// `const App` NO crea la propiedad window.App por sí solo.
window.App = App;

// RF-29 — Al recuperar el foco (p. ej. al volver de Excel), revisa si la tabla
// que se estaba editando cambió y refresca su vista previa.
window.addEventListener("focus", () => { App._revisarEdicionExcel(); });

// RF-04 — Cerrar los menús (Archivo) al hacer clic fuera.
document.addEventListener("click", () => App._cerrarMenus());

// Arrancar cuando esté listo (PyWebView o browser)
if (window.pywebview) {
  window.addEventListener("pywebviewready", init);
} else {
  document.addEventListener("DOMContentLoaded", init);
}