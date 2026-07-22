// pdf-view.js — read-only PDF viewer with annotations.
// Uses pdf.js (Mozilla) for rendering, pdf-lib for save, WebCrypto (ECDSA P-256)
// for real digital signatures. Loaded lazily from CDN only when a PDF is opened.
//
// Exports: openPdf, closePdf, isPdfMode, getPdfInfo

const PDFJS_VERSION = "3.11.174";
const PDFJS_BASE = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${PDFJS_VERSION}`;
const PDFLIB_SRC = "https://unpkg.com/pdf-lib@1.17.1/dist/pdf-lib.min.js";

let pdfjsPromise = null, pdfLibPromise = null;
let pdfDoc = null;
let pdfBytes = null;
let pdfScale = 1.2;
let pdfCurrentPage = 1;
let pdfMode = false;
let pendingRender = 0;
let hooks = {};
let annotations = [];           // [{ id, page (1-based), x, y, w, h, type, data }]
let annSeq = 0;
let placementMode = null;        // { type: 'comment'|'stamp'|'signature', data? }
let selectedAnn = null;
let stampBuffer = null;          // dataURL for stamp image pending placement
let sigSigBuffer = null;         // dataURL for signature image pending placement
let sigMetaBuffer = null;         // { signer, ts } pending placement
let stickerBuffer = null;        // { shape, color, label } pending placement
let curveKeys = null;            // ECDSA key pair (cached in IndexedDB)
let contextMenuEl = null;
let hintEl = null;
let wheelZoomListener = null;

function $(id) { return document.getElementById(id); }

function loadPdfjs() {
  if (pdfjsPromise) return pdfjsPromise;
  pdfjsPromise = new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = `${PDFJS_BASE}/pdf.min.js`;
    s.onload = () => {
      try { window.pdfjsLib.GlobalWorkerOptions.workerSrc = `${PDFJS_BASE}/pdf.worker.min.js`; resolve(window.pdfjsLib); }
      catch (e) { reject(e); }
    };
    s.onerror = () => reject(new Error("Failed to load pdf.js from CDN"));
    document.head.appendChild(s);
  });
  return pdfjsPromise;
}
function loadPdfLib() {
  if (pdfLibPromise) return pdfLibPromise;
  pdfLibPromise = new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = PDFLIB_SRC;
    s.onload = () => resolve(window.PDFLib);
    s.onerror = () => reject(new Error("Failed to load pdf-lib from CDN"));
    document.head.appendChild(s);
  });
  return pdfLibPromise;
}

// ---- ECDSA signing key (persisted in IndexedDB so the same identity can re-sign) ----
function openDb() {
  return new Promise((res, rej) => {
    const r = indexedDB.open("word-editor-pdf", 1);
    r.onupgradeneeded = () => { r.result.createObjectStore("keys"); };
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
}
async function getKey() {
  if (curveKeys) return curveKeys;
  try {
    const db = await openDb();
    const stored = await new Promise((res, rej) => {
      const tx = db.transaction("keys", "readonly").objectStore("keys").get("ecdsa");
      tx.onsuccess = () => res(tx.result); tx.onerror = () => rej(tx.error);
    });
    if (stored) { curveKeys = stored; return curveKeys; }
    const kp = await crypto.subtle.generateKey({ name: "ECDSA", namedCurve: "P-256" }, false, ["sign", "verify"]);
    await new Promise((res, rej) => {
      const tx = db.transaction("keys", "readwrite").objectStore("keys").put(kp, "ecdsa");
      tx.onsuccess = () => res(); tx.onerror = () => rej(tx.error);
    });
    curveKeys = kp;
  } catch(_) {
    curveKeys = await crypto.subtle.generateKey({ name: "ECDSA", namedCurve: "P-256" }, false, ["sign", "verify"]);
  }
  return curveKeys;
}
async function exportPubBase64() {
  const kp = await getKey();
  const spki = await crypto.subtle.exportKey("spki", kp.publicKey);
  return btoa(String.fromCharCode(...new Uint8Array(spki)));
}
async function signBytes(bytes) {
  const kp = await getKey();
  const sig = await crypto.subtle.sign({ name: "ECDSA", hash: { name: "SHA-256" } }, kp.privateKey, bytes);
  return btoa(String.fromCharCode(...new Uint8Array(sig)));
}
async function sha256(bytes) {
  const h = await crypto.subtle.digest("SHA-256", bytes);
  return Array.from(new Uint8Array(h)).map(b => b.toString(16).padStart(2, "0")).join("");
}

// ---- entities ----
function newAnn(type, page, x, y, w, h, data) {
  const a = { id: ++annSeq, type, page, x, y, w, h, data };
  annotations.push(a);
  return a;
}

// ---- toolbar ----
function buildPdfToolbar() {
  const bar = $("pdf-toolbar");
  if (!bar) return;
  bar.innerHTML = "";
  bar.className = "pdf-toolbar";
  const mk = (label, title, onClick, cls = "") => {
    const b = document.createElement("button");
    b.type = "button"; b.title = title; b.innerHTML = label;
    if (cls) b.className = cls;
    b.addEventListener("click", onClick);
    bar.appendChild(b);
    return b;
  };
  const sep = () => { const s = document.createElement("span"); s.className = "pdf-sep"; bar.appendChild(s); };
  mk("‹", "Previous page", () => goToPage(pdfCurrentPage - 1));
  const pl = document.createElement("span");
  pl.className = "pdf-page-label"; pl.id = "pdf-page-label"; bar.appendChild(pl);
  mk("›", "Next page", () => goToPage(pdfCurrentPage + 1));
  sep();
  mk("−", "Zoom out", () => { setScale(pdfScale / 1.2); });
  const z = document.createElement("span");
  z.className = "pdf-zoom-label"; z.id = "pdf-zoom-label"; z.title = "Reset to 100%";
  z.textContent = "100%"; z.addEventListener("click", () => setScale(1.2)); bar.appendChild(z);
  mk("+", "Zoom in", () => setScale(pdfScale * 1.2));
  sep();
  mk("⤡", "Fit screen", () => fitScreen(), "pdf-fit");
  mk("⤢", "Fit width", () => fitWidth(), "pdf-fit");
  sep();
  mk("💬", "Add comment", () => startPlacement("comment"));
  mk("🏷", "Add sticker (sign-here / arrow / star / check)", () => openStickerPicker(), "pdf-sticker");
  mk("🖋", "Digital sign", () => openSignatureModal());
  mk("📛", "Wet stamp — upload any image", () => uploadStamp());
  sep();
  mk("💾", "Save as new PDF (Ctrl/⌘+S)", saveAsNewPdf, "pdf-save");
  mk("🖨", "Print", printPdf);
}

function updatePageLabel() {
  const el = $("pdf-page-label");
  if (el && pdfDoc) el.textContent = `Page ${pdfCurrentPage} / ${pdfDoc.numPages}`;
  const z = $("pdf-zoom-label");
  if (z) z.textContent = Math.round((pdfScale / 1.2) * 100) + "%";
}

function setScale(s) {
  pdfScale = Math.max(0.3, Math.min(5, s));
  renderAll();
}
function pageBaseSize() {
  if (!pdfDoc) return { w: 612, h: 792 };
  // pdf.js getViewport({scale:1}) gives points (1/72") since PDF default user units
  return null; // we'll get per-page inside renderAll
}
function fitWidth() {
  if (!pdfDoc) return;
  const wrap = $("editor-wrap");
  const avail = (wrap ? wrap.clientWidth : 800) - 48;
  pdfDoc.getPage(1).then(p => {
    const v = p.getViewport({ scale: 1 });
    pdfScale = Math.max(0.3, Math.min(5, avail / v.width));
    renderAll();
  });
}
function fitScreen() {
  if (!pdfDoc) return;
  const wrap = $("editor-wrap");
  const availW = (wrap ? wrap.clientWidth : 800) - 48;
  const availH = (wrap ? wrap.clientHeight : 600) - 48;
  pdfDoc.getPage(1).then(p => {
    const v = p.getViewport({ scale: 1 });
    pdfScale = Math.max(0.3, Math.min(5, Math.min(availW / v.width, availH / v.height)));
    renderAll();
  });
}

async function goToPage(n) {
  if (!pdfDoc) return;
  n = Math.max(1, Math.min(pdfDoc.numPages, n));
  pdfCurrentPage = n;
  updatePageLabel();
  const tgt = document.getElementById(`pdf-page-${n}`);
  if (tgt) {
    const wrap = $("editor-wrap");
    if (wrap) wrap.scrollTo({ top: tgt.offsetTop - wrap.offsetTop - 24, behavior: "smooth" });
  }
  // mark current
  document.querySelectorAll(".pdf-page-wrap").forEach(w => w.classList.remove("pdf-page-current"));
  if (tgt) tgt.classList.add("pdf-page-current");
}

// ---- rendering: page canvases + annotation overlay ----
async function renderAll() {
  if (!pdfDoc) return;
  const token = ++pendingRender;
  const view = $("pdf-view");
  if (!view) return;
  view.innerHTML = "";
  for (let p = 1; p <= pdfDoc.numPages; p++) {
    const page = await pdfDoc.getPage(p);
    if (token !== pendingRender) return;
    const cap = document.createElement("div");
    cap.className = "pdf-page-wrap";
    cap.id = `pdf-page-${p}`;
    cap.dataset.page = String(p);
    const canvas = document.createElement("canvas");
    cap.appendChild(canvas);
    const overlay = document.createElement("div");
    overlay.className = "pdf-ann-layer";
    cap.appendChild(overlay);
    view.appendChild(cap);
    const ctx = canvas.getContext("2d");
    const viewport = page.getViewport({ scale: pdfScale });
    const dpr = window.devicePixelRatio || 1;
    canvas.width = Math.floor(viewport.width * dpr);
    canvas.height = Math.floor(viewport.height * dpr);
    canvas.style.width = Math.floor(viewport.width) + "px";
    canvas.style.height = Math.floor(viewport.height) + "px";
    overlay.style.width = Math.floor(viewport.width) + "px";
    overlay.style.height = Math.floor(viewport.height) + "px";
    cap.dataset.pageW = String(viewport.width); // screen px at current scale (== pdf*scale)
    cap.dataset.pageH = String(viewport.height);
    cap.dataset.pageWPts = String(page.getViewport({ scale: 1 }).width);
    cap.dataset.pageHPts = String(page.getViewport({ scale: 1 }).height);
    ctx.scale(dpr, dpr);
    await page.render({ canvasContext: ctx, viewport }).promise;
    if (token !== pendingRender) return;
    if (p === pdfCurrentPage) cap.classList.add("pdf-page-current");
    // paint annotations for this page
    for (const a of annotations) if (a.page === p) paintAnnotation(overlay, a, cap);
  }
  updatePageLabel();
  updatePageCountHint();
}

function paintAnnotation(overlay, a, cap) {
  const el = document.createElement("div");
  el.className = "pdf-ann ann-" + a.type;
  el.dataset.id = String(a.id);
  el.style.left = (a.x * pdfScale) + "px";
  el.style.top = (a.y * pdfScale) + "px";
  el.style.width = (a.w * pdfScale) + "px";
  el.style.height = (a.h * pdfScale) + "px";
  if (a.type === "comment") {
    const textEl = document.createElement("div");
    textEl.className = "ann-text";
    textEl.textContent = a.data.text;
    el.appendChild(textEl);
    el.title = "Comment — drag body to move, drag corner to resize, double-click to edit";
    el.addEventListener("dblclick", () => {
      const t = prompt("Edit comment:", a.data.text);
      if (t !== null) { a.data.text = t; textEl.textContent = t; }
    });
  } else if (a.type === "stamp") {
    const img = document.createElement("img");
    img.src = a.data.url; img.draggable = false;
    el.appendChild(img);
    el.title = "Wet stamp — drag body to move, drag corner to resize";
  } else if (a.type === "signature") {
    const img = document.createElement("img");
    img.src = a.data.url; img.draggable = false;
    el.appendChild(img);
    const tag = document.createElement("span");
    tag.className = "sig-tag";
    tag.textContent = "✓ " + (a.data.signer || "Signed");
    tag.title = "Signed by " + (a.data.signer || "?") + " at " + (a.data.ts || "");
    el.appendChild(tag);
    el.title = "Digital signature — drag body to move, drag corner to resize";
  } else if (a.type === "sticker") {
    paintSticker(el, a);
    el.title = "Sticker — drag body to move, drag corner to resize, double-click to change color";
  }
  const del = document.createElement("button");
  del.className = "ann-del";
  del.textContent = "×";
  del.title = "Delete";
  del.addEventListener("click", (e) => { e.stopPropagation(); annotations = annotations.filter(x => x.id !== a.id); el.remove(); updatePageCountHint(); });
  el.appendChild(del);
  // resize handle (bottom-right corner)
  const resize = document.createElement("div");
  resize.className = "ann-resize";
  resize.title = "Drag to resize";
  enableResize(resize, el, a);
  el.appendChild(resize);
  enableDrag(el, a);
  el.addEventListener("mousedown", () => selectAnn(a.id), true);
  overlay.appendChild(el);
}

function selectAnn(id) {
  document.querySelectorAll(".pdf-ann.selected").forEach(el => el.classList.remove("selected"));
  const el = document.querySelector(`.pdf-ann[data-id="${id}"]`);
  if (el) el.classList.add("selected");
  selectedAnn = annotations.find(a => a.id === id) || null;
}
function deselectAll() {
  document.querySelectorAll(".pdf-ann.selected").forEach(el => el.classList.remove("selected"));
  selectedAnn = null;
}

// Single global drag state machine — robust against mouseup outside the window,
// iframe focus loss, or repeated mousedowns. Uses pointer events with capture.
let dragState = null;
let dragListenersInstalled = false;
function installDragListeners() {
  if (dragListenersInstalled) return;
  dragListenersInstalled = true;
  const onMove = (e) => {
    if (!dragState) return;
    const dx = (e.clientX - dragState.startX) / pdfScale;
    const dy = (e.clientY - dragState.startY) / pdfScale;
    dragState.ann.x = Math.max(0, dragState.origX + dx);
    dragState.ann.y = Math.max(0, dragState.origY + dy);
    dragState.el.style.left = (dragState.ann.x * pdfScale) + "px";
    dragState.el.style.top = (dragState.ann.y * pdfScale) + "px";
  };
  const onUp = () => {
    if (!dragState) return;
    if (dragState.el && dragState.el.releasePointerCapture && dragState.pointerId != null) {
      try { dragState.el.releasePointerCapture(dragState.pointerId); } catch {}
    }
    dragState = null;
    document.body.classList.remove("pdf-dragging");
  };
  window.addEventListener("pointermove", onMove);
  window.addEventListener("pointerup", onUp);
  window.addEventListener("pointercancel", onUp);
  // backstops for non-pointer browsers / lost capture
  window.addEventListener("mouseup", onUp);
  window.addEventListener("blur", onUp);
}
function enableDrag(el, a) {
  el.addEventListener("pointerdown", (e) => {
    if (e.target.classList.contains("ann-del") || e.target.classList.contains("ann-resize")) return;
    if (e.button !== 0 && e.pointerType === "mouse") return;
    if (dragState || resizeState) return; // already dragging or resizing
    installDragListeners();
    dragState = { el, ann: a, startX: e.clientX, startY: e.clientY, origX: a.x, origY: a.y, pointerId: e.pointerId };
    try { el.setPointerCapture(e.pointerId); } catch {}
    document.body.classList.add("pdf-dragging");
    e.preventDefault();
  });
}

// Resize handle on the bottom-right corner of any annotation — uses the same
// robust global state machine as drag.
let resizeState = null;
let resizeListenersInstalled = false;
function installResizeListeners() {
  if (resizeListenersInstalled) return;
  resizeListenersInstalled = true;
  const onMove = (e) => {
    if (!resizeState) return;
    const dx = (e.clientX - resizeState.startX) / pdfScale;
    const dy = (e.clientY - resizeState.startY) / pdfScale;
    const minW = 30, minH = 24;
    resizeState.ann.w = Math.max(minW, resizeState.origW + dx);
    resizeState.ann.h = Math.max(minH, resizeState.origH + dy);
    resizeState.el.style.width = (resizeState.ann.w * pdfScale) + "px";
    resizeState.el.style.height = (resizeState.ann.h * pdfScale) + "px";
    // if it's a sticker, the painted inner SVG fills the box; no repaint needed.
    // if it's a comment, the text wraps automatically (CSS word-wrap).
  };
  const onUp = () => {
    if (!resizeState) return;
    if (resizeState.el && resizeState.el.releasePointerCapture && resizeState.pointerId != null) {
      try { resizeState.el.releasePointerCapture(resizeState.pointerId); } catch {}
    }
    resizeState = null;
    document.body.classList.remove("pdf-dragging");
  };
  window.addEventListener("pointermove", onMove);
  window.addEventListener("pointerup", onUp);
  window.addEventListener("pointercancel", onUp);
  window.addEventListener("mouseup", onUp);
  window.addEventListener("blur", onUp);
}
function enableResize(handle, el, a) {
  handle.addEventListener("pointerdown", (e) => {
    if (e.button !== 0 && e.pointerType === "mouse") return;
    if (dragState || resizeState) return;
    installResizeListeners();
    resizeState = { el, ann: a, startX: e.clientX, startY: e.clientY, origW: a.w, origH: a.h, pointerId: e.pointerId };
    try { handle.setPointerCapture(e.pointerId); } catch {}
    document.body.classList.add("pdf-dragging");
    e.stopPropagation();
    e.preventDefault();
  });
}

// ---- sticker shapes ----
const STICKER_SHAPES = {
  "sign-here": null,  // drawn as a callout tag
  arrow: null,        // drawn as a thick arrow
  star: null,         // 5-point star
  check: null,        // check mark
  circle: null,       // empty circle outline
};
const STICKER_COLORS = ["#e74c3c", "#f39c12", "#f1c40f", "#27ae60", "#2b579a", "#9b59b6", "#000000"];

function paintSticker(el, a) {
  const shape = a.data.shape || "sign-here";
  const color = a.data.color || "#e74c3c";
  const label = a.data.label || (shape === "sign-here" ? "Sign Here" : "");
  el.style.background = "transparent";
  if (shape === "sign-here") {
    el.style.background = color;
    el.style.color = "#fff";
    el.style.borderRadius = "6px";
    el.style.display = "flex";
    el.style.alignItems = "center";
    el.style.justifyContent = "center";
    el.style.fontWeight = "700";
    el.style.fontSize = "14px";
    el.style.boxShadow = "0 2px 4px rgba(0,0,0,0.2)";
    el.style.transform = "rotate(-3deg)";
    const span = document.createElement("span");
    span.style.padding = "4px 10px";
    span.style.textAlign = "center";
    span.style.pointerEvents = "none";
    span.textContent = label || "Sign Here";
    el.appendChild(span);
  } else {
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("viewBox", "0 0 100 100");
    svg.setAttribute("width", "100%");
    svg.setAttribute("height", "100%");
    svg.style.pointerEvents = "none";
    svg.style.overflow = "visible";
    let path = "";
    const fill = color;
    if (shape === "arrow") {
      path = '<path d="M5,35 L60,35 L60,15 L95,50 L60,85 L60,65 L5,65 Z" fill="' + fill + '"/>';
    } else if (shape === "star") {
      const cx = 50, cy = 48, r1 = 45, r2 = 18;
      let pts = "";
      for (let i = 0; i < 10; i++) {
        const r = (i % 2 === 0) ? r1 : r2;
        const a2 = (Math.PI / 5) * i - Math.PI / 2;
        pts += (cx + r * Math.cos(a2)).toFixed(1) + "," + (cy + r * Math.sin(a2)).toFixed(1) + " ";
      }
      path = '<polygon points="' + pts.trim() + '" fill="' + fill + '" stroke="rgba(0,0,0,0.15)" stroke-width="1"/>';
    } else if (shape === "check") {
      path = '<path d="M15,55 L40,80 L85,25" stroke="' + fill + '" stroke-width="14" stroke-linecap="round" stroke-linejoin="round" fill="none"/>';
    } else if (shape === "circle") {
      path = '<circle cx="50" cy="50" r="44" fill="none" stroke="' + fill + '" stroke-width="6"/>';
    }
    svg.innerHTML = path;
    el.appendChild(svg);
  }
  // double-click stickers (except sign-here) to change color
  if (shape !== "sign-here") {
    el.addEventListener("dblclick", () => {
      const idx = STICKER_COLORS.indexOf(color);
      a.data.color = STICKER_COLORS[(idx + 1) % STICKER_COLORS.length];
      el.innerHTML = "";
      paintStickerShapeOnly(el, a);
      el.appendChild(delRefFor(a, el));
      el.appendChild(resizeRefFor(a, el));
    });
  }
}
// re-paint a sticker's inner SVG/visual when color changes
function paintStickerShapeOnly(el, a) {
  paintSticker(el, a);
}
function delRefFor(a, el) {
  const del = document.createElement("button");
  del.className = "ann-del"; del.textContent = "×"; del.title = "Delete";
  del.addEventListener("click", (e) => { e.stopPropagation(); annotations = annotations.filter(x => x.id !== a.id); el.remove(); updatePageCountHint(); });
  return del;
}
function resizeRefFor(a, el) {
  const resize = document.createElement("div");
  resize.className = "ann-resize"; resize.title = "Drag to resize";
  enableResize(resize, el, a);
  return resize;
}

function openStickerPicker() {
  // overlay modal
  const overlay = document.createElement("div");
  overlay.className = "sig-overlay";
  const box = document.createElement("div");
  box.className = "sig-modal sticker-modal";
  const shapeChoices = Object.keys(STICKER_SHAPES);
  const shapeLabels = { "sign-here": "Sign Here", arrow: "Arrow", star: "Star", check: "Check", circle: "Circle" };
  box.innerHTML = `
    <h3>Add a sticker</h3>
    <div class="sticker-shapes"></div>
    <div class="sticker-colors"></div>
    <label class="sig-name" style="margin-top: 8px;">Label text (Sign Here only) <input type="text" value="Sign Here" placeholder="Sign Here" /></label>
    <div class="sig-actions">
      <button class="sig-cancel">Cancel</button>
      <button class="sig-ok">Place sticker</button>
    </div>
  `;
  overlay.appendChild(box);
  document.body.appendChild(overlay);
  let chosenShape = "sign-here";
  let chosenColor = "#e74c3c";
  // shape preview renders
  const shapesHost = box.querySelector(".sticker-shapes");
  for (const s of shapeChoices) {
    const btn = document.createElement("button");
    btn.className = "sticker-choice" + (s === chosenShape ? " active" : "");
    btn.title = shapeLabels[s];
    const preview = document.createElement("div");
    preview.style.width = "52px"; preview.style.height = "52px"; preview.style.display = "flex"; preview.style.alignItems = "center"; preview.style.justifyContent = "center";
    paintSticker(preview, { data: { shape: s, color: chosenColor, label: s === "sign-here" ? "Sign Here" : "" } });
    btn.appendChild(preview);
    btn.addEventListener("click", () => {
      chosenShape = s;
      box.querySelectorAll(".sticker-choice").forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      // re-paint already-active shape previews keep current color; no-op
    });
    shapesHost.appendChild(btn);
  }
  const colorsHost = box.querySelector(".sticker-colors");
  for (const c of STICKER_COLORS) {
    const sw = document.createElement("button");
    sw.className = "sticker-swatch" + (c === chosenColor ? " active" : "");
    sw.style.background = c;
    sw.title = c;
    sw.addEventListener("click", () => {
      chosenColor = c;
      box.querySelectorAll(".sticker-swatch").forEach(b => b.classList.remove("active"));
      sw.classList.add("active");
      // re-paint shape previews with new color
      box.querySelectorAll(".sticker-choice").forEach((b, i) => {
        const s = shapeChoices[i];
        const pv = b.querySelector("div");
        if (!pv) return;
        pv.innerHTML = "";
        paintSticker(pv, { data: { shape: s, color: chosenColor, label: s === "sign-here" ? "Sign Here" : "" } });
      });
    });
    colorsHost.appendChild(sw);
  }
  box.querySelector(".sig-cancel").addEventListener("click", () => overlay.remove());
  box.querySelector(".sig-ok").addEventListener("click", () => {
    const label = (box.querySelector(".sig-name input").value || "").trim() || "Sign Here";
    stickerBuffer = { shape: chosenShape, color: chosenColor, label };
    overlay.remove();
    startPlacement("sticker");
  });
}

function updatePageCountHint() {
  if (!pdfDoc) return;
  const annCount = annotations.length;
  const sigCount = annotations.filter(a => a.type === "signature").length;
  const pc = $("pagecount");
  if (pc) {
    let s = `${pdfDoc.numPages} page${pdfDoc.numPages === 1 ? "" : "s"} (PDF)`;
    if (annCount) s += ` · ${annCount} annotation${annCount === 1 ? "" : "s"}`;
    if (sigCount) s += ` · 🔒 signed`;
    pc.textContent = s;
  }
}

// ---- placement mode ----
function startPlacement(type, data) {
  placementMode = { type, data };
  showHint(`Click on the page to place the ${type}.`);
  const view = $("pdf-view");
  if (view) view.style.cursor = "crosshair";
}
function showHint(text) {
  if (!hintEl) {
    hintEl = document.createElement("div");
    hintEl.className = "pdf-hint";
    document.body.appendChild(hintEl);
  }
  hintEl.textContent = text;
  hintEl.classList.remove("hidden");
}
function hideHint() { if (hintEl) hintEl.classList.add("hidden"); }

function onPdfViewClick(e) {
  if (!placementMode) return;
  const cap = e.target.closest(".pdf-page-wrap");
  if (!cap) return;
  const overlay = cap.querySelector(".pdf-ann-layer");
  if (!overlay) return;
  const rect = overlay.getBoundingClientRect();
  const xPx = e.clientX - rect.left;
  const yPx = e.clientY - rect.top;
  const x = xPx / pdfScale;
  const y = yPx / pdfScale;
  const page = parseInt(cap.dataset.page, 10);
  const pageHPts = parseFloat(cap.dataset.pageHPts);
  if (placementMode.type === "comment") {
    const text = prompt("Comment text:");
    if (!text || !text.trim()) { cancelPlacement(); return; }
    newAnn("comment", page, x, y, 180, 60, { text: text.trim() });
    renderAnnotationsOnly(page);
  } else if (placementMode.type === "stamp") {
    if (!stampBuffer) { cancelPlacement(); return; }
    newAnn("stamp", page, x, y, 120, 120, { url: stampBuffer });
    renderAnnotationsOnly(page);
  } else if (placementMode.type === "signature") {
    if (!sigSigBuffer) { cancelPlacement(); return; }
    newAnn("signature", page, x, y, 200, 80, { url: sigSigBuffer, signer: sigMetaBuffer.signer, ts: sigMetaBuffer.ts });
    renderAnnotationsOnly(page);
  } else if (placementMode.type === "sticker") {
    if (!stickerBuffer) { cancelPlacement(); return; }
    const shape = stickerBuffer.shape;
    const label = stickerBuffer.label || (shape === "sign-here" ? "Sign Here" : "");
    newAnn("sticker", page, x, y, shape === "sign-here" ? 140 : 80, shape === "sign-here" ? 40 : 80, { shape, color: stickerBuffer.color, label });
    renderAnnotationsOnly(page);
  }
  cancelPlacement();
}
function cancelPlacement() {
  placementMode = null;
  hideHint();
  const view = $("pdf-view");
  if (view) view.style.cursor = "";
  stampBuffer = null;
  sigSigBuffer = null;
  sigMetaBuffer = null;
  stickerBuffer = null;
}
function renderAnnotationsOnly(page) {
  const cap = document.getElementById(`pdf-page-${page}`);
  if (!cap) return;
  const overlay = cap.querySelector(".pdf-ann-layer");
  if (!overlay) return;
  overlay.innerHTML = "";
  for (const a of annotations) if (a.page === page) paintAnnotation(overlay, a, cap);
  updatePageCountHint();
}

// ---- wet stamp upload ----
function uploadStamp() {
  const inp = document.createElement("input");
  inp.type = "file"; inp.accept = "image/png,image/jpeg,image/svg+xml";
  inp.addEventListener("change", async () => {
    const f = inp.files && inp.files[0];
    if (!f) return;
    const url = await fileToDataURL(f);
    stampBuffer = url;
    startPlacement("stamp");
  });
  inp.click();
}
function fileToDataURL(f) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result);
    r.onerror = () => rej(r.error);
    r.readAsDataURL(f);
  });
}

// ---- digital signature modal ----
function openSignatureModal() {
  const overlay = document.createElement("div");
  overlay.className = "sig-overlay";
  const box = document.createElement("div");
  box.className = "sig-modal";
  box.innerHTML = `
    <h3>Place a digital signature</h3>
    <div class="sig-tabs"><button data-tab="draw" class="active">Draw</button><button data-tab="type">Type name</button></div>
    <div class="sig-preview"><canvas width="520" height="160"></canvas><button class="sig-clear">Clear</button></div>
    <label class="sig-name">Signer name <input type="text" placeholder="Your name" /></label>
    <div class="sig-actions"><button class="sig-cancel">Cancel</button><button class="sig-ok">Sign &amp; place</button></div>
  `;
  overlay.appendChild(box);
  document.body.appendChild(overlay);
  const canvas = box.querySelector("canvas");
  const ctx = canvas.getContext("2d");
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.strokeStyle = "#1a3b8c"; ctx.lineWidth = 2.5; ctx.lineCap = "round";
  let drawing = false, last = null;
  function startStroke(x, y) { drawing = true; last = [x, y]; }
  function endStroke() { drawing = false; last = null; }
  function drawTo(x, y) {
    if (!drawing || !last) return;
    ctx.beginPath(); ctx.moveTo(last[0], last[1]); ctx.lineTo(x, y); ctx.stroke();
    last = [x, y];
  }
  function toCanvasCoords(e) {
    const r = canvas.getBoundingClientRect();
    const cx = (e.touches ? e.touches[0].clientX : e.clientX) - r.left;
    const cy = (e.touches ? e.touches[0].clientY : e.clientY) - r.top;
    return [cx * (canvas.width / r.width), cy * (canvas.height / r.height)];
  }
  canvas.addEventListener("mousedown", e => { const [x,y] = toCanvasCoords(e); startStroke(x,y); });
  canvas.addEventListener("mousemove", e => { const [x,y] = toCanvasCoords(e); drawTo(x,y); });
  canvas.addEventListener("mouseup", endStroke);
  canvas.addEventListener("mouseleave", endStroke);
  canvas.addEventListener("touchstart", e => { e.preventDefault(); const [x,y] = toCanvasCoords(e); startStroke(x,y); }, { passive: false });
  canvas.addEventListener("touchmove", e => { e.preventDefault(); const [x,y] = toCanvasCoords(e); drawTo(x,y); }, { passive: false });
  canvas.addEventListener("touchend", endStroke);
  function clearCanvas() { ctx.clearRect(0, 0, canvas.width, canvas.height); }
  box.querySelector(".sig-clear").addEventListener("click", clearCanvas);
  // tab switching
  let mode = "draw";
  box.querySelectorAll(".sig-tabs button").forEach(b => b.addEventListener("click", () => {
    box.querySelectorAll(".sig-tabs button").forEach(x => x.classList.remove("active"));
    b.classList.add("active");
    mode = b.dataset.tab;
    if (mode === "type") {
      clearCanvas();
      const name = (box.querySelector(".sig-name input").value || "").trim() || "Signature";
      ctx.fillStyle = "#1a3b8c";
      ctx.font = 'italic 64px "Brush Script MT", "Segoe Script", "Snell Roundhand", cursive';
      ctx.textBaseline = "middle";
      ctx.fillText(name, 20, canvas.height / 2);
    } else {
      clearCanvas();
    }
  }));
  box.querySelector(".sig-name input").addEventListener("input", () => {
    if (mode === "type") {
      clearCanvas();
      const name = (box.querySelector(".sig-name input").value || "").trim() || "Signature";
      ctx.fillStyle = "#1a3b8c";
      ctx.font = 'italic 64px "Brush Script MT", "Segoe Script", "Snell Roundhand", cursive';
      ctx.textBaseline = "middle";
      ctx.fillText(name, 20, canvas.height / 2);
    }
  });
  box.querySelector(".sig-cancel").addEventListener("click", () => overlay.remove());
  box.querySelector(".sig-ok").addEventListener("click", async () => {
    const name = (box.querySelector(".sig-name input").value || "Anonymous").trim();
    // ensure we have an ECDSA keypair ready (real signature crypto)
    await getKey();
    const dataUrl = canvas.toDataURL("image/png");
    sigSigBuffer = dataUrl;
    sigMetaBuffer = { signer: name, ts: new Date().toISOString() };
    overlay.remove();
    startPlacement("signature");
  });
}

// ---- save as new PDF ----
async function saveAsNewPdf() {
  if (!pdfBytes) return;
  if (hooks.setStatus) hooks.setStatus("saving");
  try {
    const lib = await loadPdfLib();
    const { PDFDocument, rgb, StandardFonts } = lib;
    const doc = await PDFDocument.load(pdfBytes);
    const helv = await doc.embedFont(StandardFonts.Helvetica);
    const helvBold = await doc.embedFont(StandardFonts.HelveticaBold);
    const helvOblique = await doc.embedFont(StandardFonts.HelveticaOblique);
    const pages = doc.getPages();

    for (const a of annotations) {
      const page = pages[a.page - 1];
      if (!page) continue;
      const { width: pw, height: ph } = page.getSize();
      // a.x, a.y are top-left in PDF points (0..pw, 0..ph)
      // pdf-lib uses bottom-up. Top-left y → bottom-left y = ph - (a.y + a.h)
      const x = a.x;
      const yBottom = ph - (a.y + a.h);
      if (a.type === "comment") {
        page.drawRectangle({ x, y: yBottom, width: a.w, height: a.h, color: rgb(1, 1, 0.75), borderColor: rgb(0.78, 0.65, 0.25), borderWidth: 1 });
        const lines = wrapText(a.data.text, a.w - 12, helv, 11);
        let yy = yBottom + a.h - 14;
        for (const ln of lines.slice(0, Math.floor(a.h / 13))) {
          page.drawText(ln, { x: x + 6, y: yy, size: 11, font: helv, color: rgb(0.1, 0.1, 0.15) });
          yy -= 13;
        }
      } else if (a.type === "stamp") {
        await embedImage(doc, page, a.data.url, x, yBottom, a.w, a.h);
      } else if (a.type === "signature") {
        await embedImage(doc, page, a.data.url, x, yBottom, a.w, a.h);
        page.drawText("Digitally signed", { x, y: yBottom - 10, size: 7, font: helvOblique, color: rgb(0.1, 0.4, 0.2) });
        page.drawText(`${a.data.signer || "Anonymous"} · ${(a.data.ts || "").slice(0,19).replace("T", " ")}`, { x, y: yBottom - 20, size: 7, font: helv, color: rgb(0.3, 0.3, 0.3) });
      } else if (a.type === "sticker") {
        drawStickerPdf(page, a, x, yBottom, helv, helvBold, rgb);
      }
    }
    saveSignedMetadata(doc, annotations);
    const saved = await doc.save({ useObjectStreams: false });
    // cryptographic signing of the final bytes (real ECDSA-SHA-256 over the saved PDF)
    const hasSig = annotations.some(a => a.type === "signature");
    let sigInfo = null;
    if (hasSig) {
      const hash = await sha256(saved);
      const sigB64 = await signBytes(saved);
      const pubB64 = await exportPubBase64();
      const sigAnn = annotations.find(a => a.type === "signature");
      sigInfo = {
        algorithm: "ECDSA-P-256-SHA-256",
        pdf_sha256: hash,
        signature_base64: sigB64,
        public_key_spki_base64: pubB64,
        signer: (sigAnn && sigAnn.data.signer) || "Anonymous",
        signed_at: (sigAnn && sigAnn.data.ts) || new Date().toISOString(),
      };
      const blobSig = new Blob([JSON.stringify(sigInfo, null, 2)], { type: "application/json" });
      downloadBlob(blobSig, (hooks.title || "document") + ".sig.json");
    }

    const blob = new Blob([saved], { type: "application/pdf" });
    downloadBlob(blob, (hooks.title || "document") + (hasSig ? "_signed" : "_edited") + ".pdf");
    if (hooks.setStatus) hooks.setStatus("saved");
  } catch (e) {
    console.error(e);
    if (hooks.setStatus) hooks.setStatus("error");
    alert("Could not save PDF: " + e.message);
  }
}

function saveSignedMetadata(doc, anns) {
  const sigAnns = anns.filter(a => a.type === "signature");
  if (!sigAnns.length) return;
  const meta = sigAnns.map(s => `${s.data.signer || "?"}|${s.data.ts || ""}`).join("; ");
  try {
    doc.setSubject("WordEditor: " + meta);
    doc.setKeywords(["word-editor", "signed", "ecdsa-sha-256"]);
    doc.setProducer("WordEditor PDF viewer");
    doc.setCreator("WordEditor");
  } catch {}
}

async function embedImage(doc, page, dataUrl, x, yBottom, w, h) {
  const isPng = dataUrl.startsWith("data:image/png");
  const bytes = dataURLToBytes(dataUrl);
  let img;
  if (isPng) img = await doc.embedPng(bytes);
  else img = await doc.embedJpg(bytes);
  page.drawImage(img, { x, y: yBottom, width: w, height: h });
}

// hexlike #rrggbb → {r,g,b} each 0..1 for pdf-lib rgb()
function hexToRgb(hex) {
  const h = (hex || "#000000").replace("#", "");
  return {
    r: parseInt(h.slice(0, 2), 16) / 255,
    g: parseInt(h.slice(2, 4), 16) / 255,
    b: parseInt(h.slice(4, 6), 16) / 255,
  };
}

function drawStickerPdf(page, a, x, yBottom, helv, helvBold, rgb) {
  const shape = a.data.shape || "sign-here";
  const color = a.data.color || "#e74c3c";
  const c = hexToRgb(color);
  const w = a.w, h = a.h;
  if (shape === "sign-here") {
    page.drawRectangle({ x, y: yBottom, width: w, height: h, color: rgb(c.r, c.g, c.b), borderColor: rgb(0, 0, 0), borderWidth: 1 });
    const label = a.data.label || "Sign Here";
    const size = Math.min(14, Math.floor(h * 0.45));
    const tw = helvBold.widthOfTextAtSize(label, size);
    page.drawText(label, { x: x + (w - tw) / 2, y: yBottom + h / 2 - size / 2 + 1, size, font: helvBold, color: rgb(1, 1, 1) });
  } else if (shape === "arrow") {
    // Approach: draw a large arrow approximation using polygons
    // Draw an orange-ish rectangle for the shaft and a triangle arrowhead.
    const shaftH = h * 0.35;
    const shaftY = yBottom + (h - shaftH) / 2;
    const shaftW = w * 0.7;
    page.drawRectangle({ x, y: shaftY, width: shaftW, height: shaftH, color: rgb(c.r, c.g, c.b) });
    // triangle head — approximate using three drawRectangle smudges... but pdf-lib has drawSvgPath
    const hx = x + shaftW;
    const tipX = x + w;
    const cy = yBottom + h / 2;
    try {
      page.drawSvgPath(`M ${shaftW} 0 L ${w} ${h/2} L ${shaftW} ${h} Z`,
        { x, y: yBottom + h, borderColor: rgb(c.r, c.g, c.b), color: rgb(c.r, c.g, c.b), borderWidth: 0, scale: 1 });
    } catch {
      // fallback if drawSvgPath not available — three thin rectangles triangulating the head
      page.drawRectangle({ x: hx, y: yBottom + h*0.05, width: w*0.3, height: h*0.45, color: rgb(c.r, c.g, c.b) });
      page.drawRectangle({ x: hx, y: yBottom + h*0.5,  width: w*0.3, height: h*0.45, color: rgb(c.r, c.g, c.b) });
      page.drawRectangle({ x: hx + w*0.25, y: yBottom + h*0.25, width: w*0.05, height: h*0.5, color: rgb(c.r, c.g, c.b) });
    }
  } else if (shape === "star") {
    // 5-point star path. pdf-lib uses bottom-up coords with origin at top-left of translate group.
    // Build the star in local coords sized to w×h, upside-down since y axis flips.
    const cx = w / 2, cy = h / 2, r1 = Math.min(w, h) * 0.5, r2 = r1 * 0.4;
    let d = "";
    for (let i = 0; i < 10; i++) {
      const r = (i % 2 === 0) ? r1 : r2;
      const ang = (Math.PI / 5) * i - Math.PI / 2;
      const px = cx + r * Math.cos(ang);
      const py = cy + r * Math.sin(ang);
      d += (i === 0 ? "M " : " L ") + px.toFixed(1) + " " + py.toFixed(1);
    }
    d += " Z";
    try {
      page.drawSvgPath(d, { x, y: yBottom + h, color: rgb(c.r, c.g, c.b), borderColor: rgb(0,0,0), borderWidth: 0.5 });
    } catch {}
  } else if (shape === "check") {
    // thick stroke check — pdf-lib doesn't support stroke thickness easily; draw filled polygons
    try {
      page.drawSvgPath("M 10 50 L 40 78 L 80 22", { x, y: yBottom + h, borderColor: rgb(c.r, c.g, c.b), borderWidth: Math.min(14, h*0.18), borderLineCap: "round", color: rgb(1,1,1) });
    } catch {
      page.drawText("✓", { x, y: yBottom, size: Math.min(h, w) * 0.9, font: helvBold, color: rgb(c.r, c.g, c.b) });
    }
  } else if (shape === "circle") {
    // Approximate circle with many short lines — pdf-lib supports drawEllipse since 1.16+
    try {
      page.drawEllipse({ x: x + w/2, y: yBottom + h/2, xScale: w/2 - 1, yScale: h/2 - 1, borderColor: rgb(c.r, c.g, c.b), borderWidth: 4 });
    } catch {
      // fallback: draw circle path
      const seg = 48;
      let d = "";
      for (let i = 0; i <= seg; i++) {
        const ang = (2 * Math.PI * i) / seg;
        const px = w/2 + (w/2 - 2) * Math.cos(ang);
        const py = h/2 + (h/2 - 2) * Math.sin(ang);
        d += (i === 0 ? "M " : " L ") + px.toFixed(1) + " " + py.toFixed(1);
      }
      d += " Z";
      try { page.drawSvgPath(d, { x, y: yBottom + h, borderColor: rgb(c.r, c.g, c.b), borderWidth: 4 }); } catch {}
    }
  }
}
function dataURLToBytes(d) {
  const b = atob(d.split(",")[1] || d);
  const arr = new Uint8Array(b.length);
  for (let i = 0; i < b.length; i++) arr[i] = b.charCodeAt(i);
  return arr;
}
function wrapText(text, maxWidth, font, size) {
  const words = text.split(/\s+/);
  const lines = [];
  let line = "";
  for (const w of words) {
    const test = line ? line + " " + w : w;
    if (font.widthOfTextAtSize(test, size) > maxWidth) {
      if (line) lines.push(line);
      line = w;
    } else {
      line = test;
    }
  }
  if (line) lines.push(line);
  return lines;
}
function downloadBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

// ---- print / download original ----
function printPdf() {
  if (!pdfBytes) return;
  const blob = new Blob([pdfBytes], { type: "application/pdf" });
  const url = URL.createObjectURL(blob);
  const w = window.open(url, "_blank");
  setTimeout(() => URL.revokeObjectURL(url), 30000);
  if (!w) alert("Please allow pop-ups to print the PDF.");
}

// ---- wheel zoom ----
function attachWheel() {
  const wrap = $("editor-wrap");
  if (!wrap) return;
  if (wheelZoomListener) wrap.removeEventListener("wheel", wheelZoomListener);
  wheelZoomListener = (e) => {
    if (!pdfMode) return;
    if (e.ctrlKey || e.metaKey) {
      e.preventDefault();
      const factor = e.deltaY < 0 ? 1.1 : 1 / 1.1;
      setScale(pdfScale * factor);
    }
  };
  wrap.addEventListener("wheel", wheelZoomListener, { passive: false });
}

// ---- context menu ----
function ensureContextMenu() {
  if (contextMenuEl) return contextMenuEl;
  contextMenuEl = document.createElement("div");
  contextMenuEl.className = "pdf-ctx hidden";
  document.body.appendChild(contextMenuEl);
  return contextMenuEl;
}
function showContextMenu(e, cap) {
  const m = ensureContextMenu();
  const page = parseInt(cap.dataset.page, 10);
  const overlay = cap.querySelector(".pdf-ann-layer");
  const rect = overlay.getBoundingClientRect();
  const x = (e.clientX - rect.left) / pdfScale;
  const y = (e.clientY - rect.top) / pdfScale;
  const ann = e.target.closest(".pdf-ann");
  const items = [];
  if (ann) {
    items.push({ label: "Edit comment", fn: () => {
      const a = annotations.find(x => x.id == ann.dataset.id);
      if (a) {
        const t = prompt("Edit comment:", a.data.text);
        if (t !== null) { a.data.text = t; const te = ann.querySelector(".ann-text"); if (te) te.textContent = t; }
      }
    } });
    items.push({ label: "Delete annotation", fn: () => { annotations = annotations.filter(a => a.id != ann.dataset.id); ann.remove(); updatePageCountHint(); } });
  } else {
    items.push({ label: "➕ Add comment here", fn: () => { const t = prompt("Comment text:"); if (t && t.trim()) { newAnn("comment", page, x, y, 180, 60, { text: t.trim() }); renderAnnotationsOnly(page); } } });
    items.push({ label: "➕ Add comment via dialog", fn: () => { startPlacement("comment"); } });
    items.push({ label: "🧿 Place wet stamp here", fn: () => { if (!stampBuffer) { uploadStamp(); } else { newAnn("stamp", page, x, y, 120, 120, { url: stampBuffer }); renderAnnotationsOnly(page); } } });
    items.push({ label: "🖋 Place signature here", fn: () => { openSignatureModal(); } });
    items.push({ label: "🏷 Place sticker here", fn: () => { openStickerPicker(); } });
    items.push({ sep: true });
    items.push({ label: "🔍 Zoom in", fn: () => setScale(pdfScale * 1.2) });
    items.push({ label: "🔍 Zoom out", fn: () => setScale(pdfScale / 1.2) });
    items.push({ label: "⤢ Fit width", fn: fitWidth });
    items.push({ label: "⤡ Fit screen", fn: fitScreen });
    items.push({ sep: true });
    items.push({ label: "💾 Save as new PDF", fn: saveAsNewPdf });
    items.push({ label: "🖨 Print", fn: printPdf });
  }
  m.innerHTML = "";
  for (const it of items) {
    if (it.sep) { const s = document.createElement("div"); s.className = "ctx-sep"; m.appendChild(s); continue; }
    const b = document.createElement("button");
    b.textContent = it.label;
    b.addEventListener("click", () => { hideContextMenu(); it.fn && it.fn(); });
    m.appendChild(b);
  }
  m.classList.remove("hidden");
  m.style.left = Math.min(e.clientX, window.innerWidth - 220) + "px";
  m.style.top = Math.min(e.clientY, window.innerHeight - 260) + "px";
}
function hideContextMenu() { if (contextMenuEl) contextMenuEl.classList.add("hidden"); }

// ---- global listeners setup ----
function setupInteraction() {
  const view = $("pdf-view");
  if (!view) return;
  view.addEventListener("click", onPdfViewClick);
  view.addEventListener("contextmenu", (e) => {
    if (!pdfMode) return;
    const cap = e.target.closest(".pdf-page-wrap");
    if (cap) { e.preventDefault(); showContextMenu(e, cap); }
  });
  document.addEventListener("click", (e) => {
    if (!pdfMode) return;
    if (contextMenuEl && !contextMenuEl.contains(e.target)) hideContextMenu();
    if (!e.target.closest(".pdf-ann")) deselectAll();
  });
  document.addEventListener("keydown", (e) => {
    if (!pdfMode) return;
    if (e.key === "Escape") { if (placementMode) { cancelPlacement(); } else { hideContextMenu(); deselectAll(); } }
    if (e.key === "Delete" && selectedAnn) {
      annotations = annotations.filter(a => a.id !== selectedAnn.id);
      const el = document.querySelector(`.pdf-ann[data-id="${selectedAnn.id}"]`);
      if (el) el.remove();
      selectedAnn = null;
      updatePageCountHint();
    }
    const mod = e.metaKey || e.ctrlKey;
    if (mod && e.key.toLowerCase() === "s") { e.preventDefault(); saveAsNewPdf(); }
    if (mod && e.key.toLowerCase() === "p") { e.preventDefault(); printPdf(); }
  });
  attachWheel();
  // re-fit on window resize
  window.addEventListener("resize", () => { if (pdfMode) { if (annotations.length || true) {} /* keep current scale on resize */ } });
}

// ---- public API ----
export async function openPdf(file, opts = {}) {
  hooks = opts;
  const lib = await loadPdfjs();
  pdfBytes = new Uint8Array(await file.arrayBuffer());
  pdfDoc = await lib.getDocument({ data: pdfBytes.slice() }).promise;
  pdfCurrentPage = 1;
  annotations = [];
  enterPdfMode();
  if (hooks.setTitle) hooks.setTitle(file.name.replace(/\.pdf$/i, ""));
  if (hooks.setStatus) hooks.setStatus("saved");
  fitWidth();
  setupInteraction();
  const wrap = $("editor-wrap");
  if (wrap) wrap.scrollTop = 0;
}
export function closePdf() {
  if (!pdfMode) return;
  pdfMode = false;
  if (pdfDoc) { try { pdfDoc.destroy(); } catch {} }
  pdfDoc = null; pdfBytes = null;
  annotations = []; annSeq = 0;
  pendingRender++;
  exitPdfMode();
  if (hooks.onClosed) hooks.onClosed();
}
export function isPdfMode() { return pdfMode; }
export function getPdfInfo() {
  if (!pdfDoc) return null;
  const sigCount = annotations.filter(a => a.type === "signature").length;
  const annCount = annotations.length;
  return { numPages: pdfDoc.numPages, currentPage: pdfCurrentPage, scale: pdfScale, annCount, sigCount };
}

function enterPdfMode() {
  pdfMode = true;
  const tb = $("toolbar"); if (tb) tb.classList.add("hidden");
  const fh = $("find-host"); if (fh) fh.classList.add("hidden");
  const po = $("page-outer"); if (po) po.classList.add("hidden");
  const view = $("pdf-view"); if (view) { view.classList.remove("hidden"); }
  buildPdfToolbar();
}
function exitPdfMode() {
  const tb = $("toolbar"); if (tb) tb.classList.remove("hidden");
  const po = $("page-outer"); if (po) po.classList.remove("hidden");
  const view = $("pdf-view"); if (view) { view.innerHTML = ""; view.classList.add("hidden"); }
  const bar = $("pdf-toolbar"); if (bar) { bar.innerHTML = ""; bar.classList.add("hidden"); }
  const fh = $("find-host"); if (fh) fh.classList.remove("hidden");
  const pc = $("pagecount"); if (pc) pc.textContent = "1 page";
  if (hintEl) hintEl.remove(), hintEl = null;
  if (contextMenuEl) contextMenuEl.classList.add("hidden");
}