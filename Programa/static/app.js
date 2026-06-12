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
  config        : { opciones: [], colores: {} },
  tienePDF      : false,
  _afilTimer    : null,
};


// ─────────────────────────────────────────────────────────────────────────────
// Wrapper fetch  (GET / POST JSON / PUT JSON / POST FormData)
// ─────────────────────────────────────────────────────────────────────────────
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
    return r.json();
  },
  async put(path, body) {
    const r = await fetch(path, {
      method  : "PUT",
      headers : { "Content-Type": "application/json" },
      body    : JSON.stringify(body),
    });
    if (!r.ok) { const t = await r.text(); throw new Error(t); }
    return r.json();
  },
  async patch(path, body) {
    const r = await fetch(path, {
      method  : "PATCH",
      headers : { "Content-Type": "application/json" },
      body    : JSON.stringify(body),
    });
    if (!r.ok) { const t = await r.text(); throw new Error(t); }
    return r.json();
  },
  async postForm(path, fd) {
    const r = await fetch(path, { method: "POST", body: fd });
    if (!r.ok) { const t = await r.text(); throw new Error(t); }
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
  const orden  = ["pdf","autores","afiliaciones","referencias","figuras","tablas","exportar"];
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
  },


  // ══════════════════════════════════════════════════════════════════════════
  // PDF — carga
  // ══════════════════════════════════════════════════════════════════════════

  async seleccionarPDF() {
    try {
      if (window.pywebview) {
        const ruta = await window.pywebview.api.abrir_pdf();
        if (ruta) await App._cargarPorRuta(ruta);
      } else {
        // Fallback desarrollo: input[type=file]
        const input = document.createElement("input");
        input.type   = "file";
        input.accept = ".pdf";
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
    setStatus("Procesando PDF...", "idle");
    showLoading(true);
    try {
      const data = await API.post("/api/pdf/cargar-ruta", { ruta });
      App._aplicarResultadoPDF(data);
    } catch (e) {
      showLoading(false);
      setStatus("Error: " + e.message, "error");
      showToast("No se pudo procesar el PDF", 4000);
    }
  },

  async _cargarPorUpload(file) {
    setStatus("Procesando PDF...", "idle");
    showLoading(true);
    try {
      const fd = new FormData();
      fd.append("file", file);
      const data = await API.postForm("/api/pdf/cargar", fd);
      App._aplicarResultadoPDF(data);
    } catch (e) {
      showLoading(false);
      setStatus("Error: " + e.message, "error");
      showToast("No se pudo procesar el PDF", 4000);
    }
  },

  _aplicarResultadoPDF(data) {
    showLoading(false);
    State.bloques  = data.bloques || [];
    State.tienePDF = true;

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
    if (file?.type === "application/pdf") {
      await App._cargarPorUpload(file);
    } else {
      showToast("Solo se aceptan archivos PDF");
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
    const optsHtml = opciones.map(o => `<option value="${esc(o)}">${esc(o)}</option>`).join("");

    container.innerHTML = `
      <div class="bloques-toolbar">
        <span style="font-weight:600">${bloques.length} bloque${bloques.length !== 1 ? "s" : ""}</span>
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

      // Construimos el select con la opción correcta pre-seleccionada
      const optsConSelected = opciones
        .map(o => `<option value="${esc(o)}"${o === b.clasificacion ? " selected" : ""}>${esc(o)}</option>`)
        .join("");

      div.innerHTML = `
        <div class="bloque-num">${idx + 1}</div>
        <textarea class="bloque-texto"
          oninput="this.style.height='auto';this.style.height=this.scrollHeight+'px'"
          onblur="App._onBloqueTextoBlur(${idx}, this.value)"
        >${esc(b.contenido)}</textarea>
        <select class="bloque-select"
          onchange="App._onBloqueClaseChange(${idx}, this.value, this.closest('.bloque-item'))">
          ${optsConSelected}
        </select>
        <button class="bloque-del" title="Marcar como Ignorar"
          onclick="App._ignorarBloque(${idx}, this.closest('.bloque-item'))">✕</button>
      `;

      lista.appendChild(div);
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
    State.autores.push({ nombre: "", orcid: "" });
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
      if (inputs[0]) State.autores[i].nombre = inputs[0].value.trim();
      if (inputs[1]) State.autores[i].orcid  = inputs[1].value.trim();
    });
  },

  async _pushAutores() {
    try { await API.put("/api/autores", { autores: State.autores }); }
    catch (_) {}
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
    } catch (_) {}
  },

  async _eliminarFigura(idx) {
    try {
      const data = await fetch(`/api/figuras/${idx}`, { method: "DELETE" }).then(r => r.json());
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
      const titulo  = esc(t.titulo || "");
      const ancla   = esc(t.ancla  || "");
      const archivo = esc(t.ruta ? t.ruta.split(/[\\/]/).pop() : "");
      const hoja    = esc(t.hoja || "");
      const preview = t.contenido || "";
      const subtitle = archivo ? (hoja ? `${archivo} › ${hoja}` : archivo) : `Tabla ${i + 1}`;
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
            <button class="tabla-row-del" onclick="App._eliminarTabla(${i})" title="Eliminar">✕</button>
          </div>
          <div class="tabla-row-body">
            <input class="tabla-input" type="text" value="${titulo}"
              placeholder="Título de la tabla…"
              onblur="App._syncTabla(${i}, 'titulo', this.value)" />
            <div class="figura-ancla-label">📍 Párrafo donde va la tabla:</div>
            <input class="tabla-input" type="text" value="${ancla}"
              placeholder='Ej: "...la Dra. Elena Centeno (Tabla 1)."'
              onblur="App._syncTabla(${i}, 'ancla', this.value)" />
            ${preview ? `<div class="tabla-preview-text">${esc(preview)}</div>` : ""}
          </div>
        </div>`;
    }).join("");
  },

  async _syncTabla(idx, campo, valor) {
    try {
      await fetch(`/api/tablas/${idx}`, {
        method: "PATCH",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ idx, [campo]: valor }),
      });
    } catch (_) {}
  },

  async _eliminarTabla(idx) {
    try {
      const data = await fetch(`/api/tablas/${idx}`, { method: "DELETE" }).then(r => r.json());
      App._renderTablas(data.tablas || []);
    } catch (e) {
      setStatus("Error: " + e.message, "error");
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
        // ── Producción: diálogo nativo → guardar en disco ──
        const fn = { html: "guardar_html", xml: "guardar_xml", epub: "guardar_epub" }[formato];
        const ruta = await window.pywebview.api[fn]();
        if (!ruta) return;   // usuario canceló

        setStatus(`Exportando ${formato.toUpperCase()}...`, "idle");
        showLoading(true);
        await API.post(`/api/exportar/${formato}`, { formato, ruta_destino: ruta });
        showLoading(false);
        setStatus(`${formato.toUpperCase()} guardado`);
        showToast(`✓ ${formato.toUpperCase()} exportado correctamente`);

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

  _confirmarCambiarPDF() {
    App._cerrarConfirmarPDF();

    // Limpiar solo los bloques y la info del PDF
    State.bloques  = [];
    State.tienePDF = false;

    // Resetear la UI del PDF
    const dz = $("dropzone");
    const bc = $("bloques-container");
    if (dz) dz.style.display = "flex";
    if (bc) { bc.style.display = "none"; bc.innerHTML = ""; }

    $("btn-leyenda")     && ($("btn-leyenda").style.display     = "none");
    $("filtro-grupo")    && ($("filtro-grupo").style.display    = "none");
    $("btn-cambiar-pdf") && ($("btn-cambiar-pdf").style.display = "none");

    // Resetear info del documento
    sv("info-nombre",  "—");
    sv("info-paginas", "—");
    sv("info-tamanio", "—");
    const badge = $("info-estado");
    if (badge) { badge.textContent = "Sin cargar"; badge.className = "badge badge--gray"; }

    // Limpiar solo bloques en el servidor
    API.post("/api/pdf/limpiar", {}).catch(() => {});

    actualizarStepper();
    setStatus("Listo para cargar nuevo PDF");
    showToast("PDF eliminado — carga el nuevo archivo");
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

    if (estado.tiene_pdf && estado.sesion_restaurada) {
      setStatus("Sesión restaurada ✓");
      showToast("✓ Se recuperó tu sesión anterior — puedes continuar donde lo dejaste", 4500);
    } else {
      setStatus("Listo para comenzar");
    }

  } catch (e) {
    setStatus("Error iniciando: " + e.message, "error");
    console.error("[init]", e);
  }
}

// Arrancar cuando esté listo (PyWebView o browser)
if (window.pywebview) {
  window.addEventListener("pywebviewready", init);
} else {
  document.addEventListener("DOMContentLoaded", init);
}