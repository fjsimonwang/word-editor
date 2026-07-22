// pdf-view.js — read-only PDF viewer built on Mozilla pdf.js (lazy-loaded
// from CDN on first use). Renders pages to <canvas> inside #pdf-view and
// swaps the regular editing toolbar for a PDF-specific one.
//
// Exports:
//   openPdf(file, { setTitle, setStatus, onClosed, fitWidthDefault })
//   closePdf()
//   isPdfMode()

const PDFJS_VERSION = "3.11.174";
const PDFJS_BASE = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${PDFJS_VERSION}`;
let pdfLibPromise = null;
let pdfDoc = null;            // PDFDocumentProxy
let pdfData = null;           // Uint8Array (original bytes, for download/print)
let pdfScale = 1.2;
let pdfFitWidth = false;
let pdfCurrentPage = 1;
let pdfMode = false;
let pendingRender = null;     // token to cancel in-flight renders
let hooks = {};

function $(id) { return document.getElementById(id); }

// lazy-load pdf.js UMD build once
function loadPdfLib() {
  if (pdfLibPromise) return pdfLibPromise;
  pdfLibPromise = new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = `${PDFJS_BASE}/pdf.min.js`;
    s.onload = () => {
      try {
        window.pdfjsLib.GlobalWorkerOptions.workerSrc = `${PDFJS_BASE}/pdf.worker.min.js`;
        resolve(window.pdfjsLib);
      } catch (e) { reject(e); }
    };
    s.onerror = () => reject(new Error("Failed to load pdf.js from CDN — check your network connection."));
    document.head.appendChild(s);
  });
  return pdfLibPromise;
}

// Build a PDF-specific toolbar (prev / page indicator / next / zoom / fit / download / print / close)
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
  const sep = () => {
    const s = document.createElement("span");
    s.className = "pdf-sep";
    bar.appendChild(s);
  };

  const prevBtn = mk("‹", "Previous page", () => goToPage(pdfCurrentPage - 1));
  const pageLabel = document.createElement("span");
  pageLabel.className = "pdf-page-label";
  pageLabel.id = "pdf-page-label";
  bar.appendChild(pageLabel);
  const nextBtn = mk("›", "Next page", () => goToPage(pdfCurrentPage + 1));

  sep();
  mk("−", "Zoom out", () => { pdfFitWidth = false; setScale(pdfScale / 1.2); });
  const zoomLabel = document.createElement("span");
  zoomLabel.className = "pdf-zoom-label";
  zoomLabel.id = "pdf-zoom-label";
  zoomLabel.textContent = "100%";
  zoomLabel.title = "Reset zoom to 100%";
  zoomLabel.addEventListener("click", () => { pdfFitWidth = false; setScale(1.2); });
  bar.appendChild(zoomLabel);
  mk("+", "Zoom in", () => { pdfFitWidth = false; setScale(pdfScale * 1.2); });

  sep();
  mk("⤢", "Fit width", () => { pdfFitWidth = true; renderAll(); }, "pdf-fit");

  sep();
  mk("⤓", "Download original PDF", downloadPdf);
  mk("🖨", "Print", printPdf);
  sep();
  mk("✕ Close PDF", "Close PDF and return to editor", closePdf, "pdf-close");
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

async function goToPage(n) {
  if (!pdfDoc) return;
  n = Math.max(1, Math.min(pdfDoc.numPages, n));
  pdfCurrentPage = n;
  updatePageLabel();
  const target = document.getElementById(`pdf-page-${n}`);
  if (target) {
    const wrap = $("editor-wrap");
    if (wrap) wrap.scrollTo({ top: target.offsetTop - wrap.offsetTop - 24, behavior: "smooth" });
  }
}

function fitScaleForViewport() {
  if (!pdfDoc || !pdfFitWidth) return pdfScale;
  const wrap = $("editor-wrap");
  const avail = (wrap ? wrap.clientWidth : 800) - 64;
  return Math.max(0.3, Math.min(5, avail / 612 * 1.0)); // 612pt is US Letter width
}

async function renderAll() {
  if (!pdfDoc) return;
  const token = ++pendingRender;
  const view = $("pdf-view");
  if (!view) return;
  view.innerHTML = "";
  const scale = fitScaleForViewport();
  const wrap = $("editor-wrap");
  const avail = wrap ? wrap.clientWidth - 48 : 800;

  for (let p = 1; p <= pdfDoc.numPages; p++) {
    const page = await pdfDoc.getPage(p);
    if (token !== pendingRender) return; // superseded
    const cap = document.createElement("div");
    cap.className = "pdf-page-wrap";
    cap.id = `pdf-page-${p}`;
    const canvas = document.createElement("canvas");
    cap.appendChild(canvas);
    view.appendChild(cap);
    const ctx = canvas.getContext("2d");
    const viewport = page.getViewport({ scale });
    // device-pixel-ratio for crisp text
    const dpr = window.devicePixelRatio || 1;
    canvas.width = Math.floor(viewport.width * dpr);
    canvas.height = Math.floor(viewport.height * dpr);
    canvas.style.width = Math.floor(viewport.width) + "px";
    canvas.style.height = Math.floor(viewport.height) + "px";
    ctx.scale(dpr, dpr);
    // wait until scroll target page is rendered before painting later ones
    await page.render({ canvasContext: ctx, viewport }).promise;
    if (token !== pendingRender) return;
    if (p === pdfCurrentPage) cap.classList.add("pdf-page-current");
  }
  updatePageLabel();
}

function downloadPdf() {
  if (!pdfData) return;
  const blob = new Blob([pdfData], { type: "application/pdf" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = (hooks.title || "document") + ".pdf";
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function printPdf() {
  if (!pdfData) return;
  // Open the raw PDF in a new tab — browsers provide a native print dialog
  const blob = new Blob([pdfData], { type: "application/pdf" });
  const url = URL.createObjectURL(blob);
  const w = window.open(url, "_blank");
  setTimeout(() => URL.revokeObjectURL(url), 30000);
  if (!w) alert("Please allow pop-ups to print the PDF.");
}

// ---- public API ----

export async function openPdf(file, opts = {}) {
  hooks = opts;
  const lib = await loadPdfLib();
  pdfData = new Uint8Array(await file.arrayBuffer());
  pdfDoc = await lib.getDocument({ data: pdfData }).promise;
  pdfCurrentPage = 1;
  pdfFitWidth = !!opts.fitWidthDefault;
  pdfScale = pdfFitWidth ? fitScaleForViewport() : 1.2;
  enterPdfMode();
  if (hooks.setTitle) hooks.setTitle(file.name.replace(/\.pdf$/i, ""));
  if (hooks.setStatus) hooks.setStatus("saved");
  await renderAll();
  // scroll to first page
  const wrap = $("editor-wrap");
  if (wrap) wrap.scrollTop = 0;
}

export function closePdf() {
  if (!pdfMode) return;
  pdfMode = false;
  if (pdfDoc) { try { pdfDoc.destroy(); } catch {} }
  pdfDoc = null;
  pdfData = null;
  pendingRender = (pendingRender || 0) + 1; // cancel any in-flight render
  exitPdfMode();
  if (hooks.onClosed) hooks.onClosed();
}

export function isPdfMode() { return pdfMode; }

function enterPdfMode() {
  pdfMode = true;
  const tb = $("toolbar"); if (tb) tb.classList.add("hidden");
  const fh = $("find-host"); if (fh) fh.classList.add("hidden");
  const po = $("page-outer"); if (po) po.classList.add("hidden");
  const view = $("pdf-view"); if (view) view.classList.remove("hidden");
  buildPdfToolbar();
}

function exitPdfMode() {
  const tb = $("toolbar"); if (tb) tb.classList.remove("hidden");
  const po = $("page-outer"); if (po) po.classList.remove("hidden");
  const view = $("pdf-view"); if (view) { view.innerHTML = ""; view.classList.add("hidden"); }
  const bar = $("pdf-toolbar"); if (bar) { bar.innerHTML = ""; bar.classList.add("hidden"); }
  const fh = $("find-host"); if (fh) fh.classList.remove("hidden");
  // restore statusbar page indicator
  const pc = $("pagecount"); if (pc) pc.textContent = "1 page";
}

// expose for keyboard shortcuts / status bar
export function getPdfInfo() {
  if (!pdfDoc) return null;
  return { numPages: pdfDoc.numPages, currentPage: pdfCurrentPage, scale: pdfScale };
}