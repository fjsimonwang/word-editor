import {
  buildToolbar, attachEditorBehaviors, attachTableToolbar, attachContextMenu,
  attachImageEditing, initTooltips,
  createFindPanel, createTrackChanges, wrapSelectionComment,
  sanitizeHtml, insertHtmlAtCaret, insertTextAtCaret, openPageSetupDialog,
  openDialog, countWords, saveSelection,
  insertImage, openTableDialog, openLinkDialog, openSymbolDialog,
  openWordArtDialog, openShapeDialog, insertPageBreak, insertBlankPage,
} from "./editor.js";
import { History } from "./history.js";
import {
  importDocx, buildDocxFromHtml, supportsDocx, DEFAULT_PAGE_SETUP,
  htmlToPlainText, plainTextToHtml, exportStandaloneHtml,
} from "./docx.js";
import {
  Autosaver, SyncClient, createDocument, getDocument, listDocuments, deleteDocument,
  importDocxFile, putDocx, saveDocument, listVersions, getVersion, restoreVersion,
} from "./store.js";
import { openPdf, closePdf, isPdfMode, getPdfInfo } from "./pdf-view.js";

const LS_KEY = "word-editor:current-id";
const LS_USER = "word-editor:user";
const $ = (id) => document.getElementById(id);
const params = new URLSearchParams(location.search);

// page dimensions in inches
const PAGE_INCHES = {
  Letter: { w: 8.5, h: 11 }, A4: { w: 8.27, h: 11.69 },
  Legal: { w: 8.5, h: 14 }, A3: { w: 11.69, h: 16.54 },
};
const PAGE_GAP = 32; // px between simulated sheets

function setStatus(s) {
  const el = $("save-status");
  el.textContent = s === "conflict" ? "conflict!" : s;
  el.className = "status " + s;
}
function saveAs(blob, name) {
  const a = document.createElement("a");
  const url = URL.createObjectURL(blob);
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
function getUserName() {
  if (params.get("user")) return params.get("user").slice(0, 40);
  let u = localStorage.getItem(LS_USER);
  if (!u) {
    u = "Guest-" + Math.random().toString(36).slice(2, 6);
    localStorage.setItem(LS_USER, u);
  }
  return u;
}
function fmtTime(t) {
  return t ? new Date(t).toLocaleString() : "";
}

// ---------------------------------------------------------------
// Main app
// ---------------------------------------------------------------
async function main() {
  const editor = $("editor");
  const editHistory = new History(editor);
  editHistory.attach();
  const toolbarHost = $("toolbar");
  const titleInput = $("doc-title");
  const fileInput = $("file-input");
  const userName = getUserName();

  // ---- embed / config flags ----
  const embedded = params.get("embed") === "1" || window.parent !== window;
  const showToolbar = params.get("toolbar") !== "0";
  const showStatusbar = params.get("statusbar") !== "0";
  const mode = params.get("mode") === "view" ? "view" : "edit";
  if (embedded) document.body.classList.add("embed");
  if (!showToolbar) toolbarHost.classList.add("hidden");
  if (!showStatusbar) $("statusbar").classList.add("hidden");

  if (!supportsDocx()) {
    setStatus("error");
    const warn = document.createElement("div");
    warn.className = "banner";
    warn.textContent = "This browser lacks CompressionStream support; .docx import/export may fail. Use a recent Chrome, Edge, Firefox, or Safari.";
    toolbarHost.parentNode.insertBefore(warn, toolbarHost);
  }

  // ---- document state ----
  let current = { id: null, title: "Untitled", pageSetup: { ...DEFAULT_PAGE_SETUP } };
  let docComments = [];

  // ---- clean serialization (pagination spacers + UI marks stripped) ----
  function getCleanHtml() {
    const clone = editor.cloneNode(true);
    for (const el of clone.querySelectorAll("[data-pg]")) {
      el.style.marginTop = "";
      el.removeAttribute("data-pg");
      if (!el.getAttribute("style")) el.removeAttribute("style");
    }
    for (const el of clone.querySelectorAll(".comment-ref.active")) el.classList.remove("active");
    for (const m of clone.querySelectorAll("mark.find-hit")) {
      const p = m.parentNode;
      while (m.firstChild) p.insertBefore(m.firstChild, m);
      m.remove();
    }
    return clone.innerHTML;
  }

  // ---- visual pagination: simulate separate sheets ----
  let paginateTimer = null;
  let ruler = null;
  let pageMetrics = { count: 1, ph: 1056, gap: PAGE_GAP, mTop: 96, mBottom: 96, mLeft: 96, mRight: 96, pageWidth: 816 };
  function schedulePaginate() {
    clearTimeout(paginateTimer);
    paginateTimer = setTimeout(paginate, 180);
  }
  function paginate() {
    const page = $("page");
    // clear previous spacers/gaps
    for (const el of editor.querySelectorAll("[data-pg]")) {
      el.style.marginTop = "";
      el.removeAttribute("data-pg");
      if (!el.getAttribute("style")) el.removeAttribute("style");
    }
    for (const g of page.querySelectorAll(".page-gap")) g.remove();

    const s = current.pageSetup || DEFAULT_PAGE_SETUP;
    const dim = PAGE_INCHES[s.size] || PAGE_INCHES.Letter;
    const ph = (s.orientation === "landscape" ? dim.w : dim.h) * 96;
    const m = s.margins || DEFAULT_PAGE_SETUP.margins;
    const mTop = (m.top != null ? m.top : 1) * 96;
    const mBottom = (m.bottom != null ? m.bottom : 1) * 96;
    const contentH = ph - mTop - mBottom;
    if (contentH < 120) { page.style.minHeight = ph + "px"; return; }

    let pageEnd = mTop + contentH; // page-relative Y where current sheet's content area ends
    let count = 1;
    let forceNext = false;
    const gaps = [];
    for (const block of editor.children) {
      // free-floating objects (dragged shapes) don't participate in the flow
      if (getComputedStyle(block).position === "absolute") continue;
      const h = block.offsetHeight;
      let top = block.offsetTop;
      const needsPush = forceNext || top >= pageEnd || (top + h > pageEnd && h <= contentH);
      if (needsPush) {
        const nextStart = pageEnd + mBottom + PAGE_GAP + mTop;
        const delta = nextStart - top;
        if (delta > 0) {
          const existing = parseFloat(getComputedStyle(block).marginTop) || 0;
          block.style.marginTop = (existing + delta) + "px";
          block.setAttribute("data-pg", "1");
          top = block.offsetTop; // re-read: margin collapse may shift the result
        }
        gaps.push(pageEnd + mBottom);
        count++;
        pageEnd = top + contentH;
        forceNext = false;
      } else if (top + h > pageEnd) {
        // oversized block straddles sheets; no gap visuals inside it
        while (top + h > pageEnd) { pageEnd += contentH + mBottom + mTop; count++; }
      }
      if (block.classList && block.classList.contains("page-break")) forceNext = true;
    }
    page.style.minHeight = (pageEnd + mBottom) + "px";
    for (const y of gaps) {
      const g = document.createElement("div");
      g.className = "page-gap";
      g.style.top = y + "px";
      g.style.height = PAGE_GAP + "px";
      g.setAttribute("contenteditable", "false");
      page.appendChild(g);
    }
    pageMetrics = {
      count, ph, gap: PAGE_GAP, mTop, mBottom,
      mLeft: (m.left != null ? m.left : 1) * 96,
      mRight: (m.right != null ? m.right : 1) * 96,
      pageWidth: page.clientWidth,
    };
    renderChrome();
    updateCurrentPage();
    if (ruler) ruler.render();
  }
  window.addEventListener("resize", schedulePaginate);
  editor.addEventListener("load", schedulePaginate, true); // images finishing load

  // ---- headers / footers / page numbers (rendered per simulated sheet) ----
  function toRoman(n) {
    const map = [[1000,"m"],[900,"cm"],[500,"d"],[400,"cd"],[100,"c"],[90,"xc"],[50,"l"],[40,"xl"],[10,"x"],[9,"ix"],[5,"v"],[4,"iv"],[1,"i"]];
    let out = ""; for (const [v, s] of map) while (n >= v) { out += s; n -= v; } return out;
  }
  function toAlpha(n) {
    let out = ""; while (n > 0) { n--; out = String.fromCharCode(97 + (n % 26)) + out; n = Math.floor(n / 26); } return out;
  }
  function pageNumText(n, total, fmt) {
    switch (fmt) {
      case "roman": return toRoman(n);
      case "alpha": return toAlpha(n);
      case "page": return "Page " + n;
      case "pageOfN": return "Page " + n + " of " + total;
      default: return String(n);
    }
  }
  function getChrome() {
    const s = current.pageSetup;
    if (!s.chrome) s.chrome = {};
    return s.chrome;
  }
  function renderChrome() {
    const page = $("page");
    page.querySelectorAll(".pg-chrome").forEach((e) => e.remove());
    const ch = (current.pageSetup && current.pageSetup.chrome) || {};
    const met = pageMetrics;
    const hasHeader = ch.header && ch.header.text;
    const hasFooter = ch.footer && ch.footer.text;
    const pn = ch.pageNumber && ch.pageNumber.enabled ? ch.pageNumber : null;
    if (!hasHeader && !hasFooter && !pn) return;

    const contentW = met.pageWidth - met.mLeft - met.mRight;
    for (let k = 1; k <= met.count; k++) {
      const sheetTop = (k - 1) * (met.ph + met.gap);
      const mkBand = (yTop, def, place, type) => {
        const band = document.createElement("div");
        band.className = "pg-chrome";
        band.style.left = met.mLeft + "px";
        band.style.top = yTop + "px";
        band.style.width = contentW + "px";
        band.setAttribute("data-type", type);
        if (def && def.align) band.setAttribute("data-align", def.align);
        const cells = { left: document.createElement("span"), center: document.createElement("span"), right: document.createElement("span") };
        for (const [k2, c] of Object.entries(cells)) { c.className = "pg-cell pg-" + k2; c.contentEditable = "false"; band.appendChild(c); }
        if (def && def.text) cells[def.align || "center"].textContent = def.text;
        if (pn && pn.place && pn.place.startsWith(place)) {
          const where = pn.place.split("-")[1] || "center";
          const t = pageNumText(k, met.count, pn.format);
          cells[where].textContent = (cells[where].textContent ? cells[where].textContent + "  " : "") + t;
        }
        // double-click to edit text inline
        band.addEventListener("dblclick", () => {
          const align = band.getAttribute("data-align") || "center";
          const cell = band.querySelector(`.pg-cell.pg-${align}`);
          if (!cell) return;
          cell.contentEditable = "true";
          cell.focus();
          const sel = window.getSelection();
          const r = document.createRange();
          r.selectNodeContents(cell);
          sel.removeAllRanges();
          sel.addRange(r);
          const commit = () => {
            if (cell.contentEditable !== "true") return;
            cell.contentEditable = "false";
            const text = cell.textContent.replace(/\s+/g, " ").trim();
            const orig = def && def.text ? def.text : "";
            if (text && text !== orig) {
              const s = current.pageSetup || (current.pageSetup = {});
              const ct = s.chrome || (s.chrome = {});
              ct[type] = ct[type] || {};
              ct[type].text = text;
              ct[type].align = align;
              scheduleSave();
              schedulePaginate();
            }
          };
          cell.addEventListener("blur", commit, { once: true });
          cell.addEventListener("keydown", (ke) => {
            if (ke.key === "Enter") { ke.preventDefault(); cell.blur(); }
          });
        });
        page.appendChild(band);
      };
      // header band sits in the top margin; footer band in the bottom margin
      if (hasHeader || (pn && pn.place && pn.place.startsWith("header"))) {
        mkBand(sheetTop + Math.max(12, met.mTop / 2 - 8), ch.header, "header", "header");
      }
      if (hasFooter || (pn && pn.place && pn.place.startsWith("footer"))) {
        mkBand(sheetTop + met.ph - met.mBottom / 2 - 8, ch.footer, "footer", "footer");
      }
    }
  }

  // ---- current page indicator (updates while scrolling) ----
  function updateCurrentPage() {
    const met = pageMetrics;
    const page = $("page");
    const wrap = $("editor-wrap");
    const zoom = parseFloat(getComputedStyle(page).zoom) || 1;
    const pr = page.getBoundingClientRect();
    const wr = wrap.getBoundingClientRect();
    // sample a point ~1/3 down the viewport — the page occupying that band reads
    // as the "current" page (and this absorbs the ruler/padding offset above the page)
    const probe = (wr.top - pr.top) + wrap.clientHeight * 0.34;
    const unit = (met.ph + met.gap) * zoom;
    let cur = Math.floor(probe / unit) + 1;
    cur = Math.min(met.count, Math.max(1, cur));
    const pc = $("pagecount");
    if (pc) pc.textContent = `Page ${cur} of ${met.count}`;
  }
  $("editor-wrap").addEventListener("scroll", updateCurrentPage, { passive: true });

  // ---- rulers: fixed app chrome. The horizontal ruler is pinned right below
  // the toolbar (full width); the vertical ruler runs down the app window's
  // left edge. Each has a "track" that overlays the page — the horizontal track
  // spans the page's width, the vertical track the FIRST page's height (so it's
  // only visible while page 1 is in view). Both carry draggable margin markers.
  // Positions are computed from live rects, so they stay aligned through
  // scrolling, zoom, panel toggles and width-clamping. ----
  function attachRuler() {
    const app = $("app");
    const toolbar = $("toolbar");
    const wrap = $("editor-wrap");
    const page = $("page");

    const hbar = document.createElement("div");
    hbar.className = "ruler-h";
    hbar.innerHTML =
      '<div class="ruler-track">' +
      '<div class="ruler-shade left"></div><div class="ruler-shade right"></div>' +
      '<div class="ruler-marker left" title="Left margin — drag to adjust"></div>' +
      '<div class="ruler-marker right" title="Right margin — drag to adjust"></div>' +
      '</div>';
    toolbar.after(hbar);

    const vbar = document.createElement("div");
    vbar.className = "ruler-v";
    vbar.innerHTML =
      '<div class="ruler-vtrack">' +
      '<div class="ruler-vshade top"></div><div class="ruler-vshade bottom"></div>' +
      '<div class="ruler-vmarker top" title="Top margin — drag to adjust"></div>' +
      '<div class="ruler-vmarker bottom" title="Bottom margin — drag to adjust"></div>' +
      '</div>';
    app.appendChild(vbar);

    const track = hbar.querySelector(".ruler-track");
    const shadeL = hbar.querySelector(".ruler-shade.left");
    const shadeR = hbar.querySelector(".ruler-shade.right");
    const markL = hbar.querySelector(".ruler-marker.left");
    const markR = hbar.querySelector(".ruler-marker.right");
    const vtrack = vbar.querySelector(".ruler-vtrack");
    const vshadeT = vbar.querySelector(".ruler-vshade.top");
    const vshadeB = vbar.querySelector(".ruler-vshade.bottom");
    const vmarkT = vbar.querySelector(".ruler-vmarker.top");
    const vmarkB = vbar.querySelector(".ruler-vmarker.bottom");

    function pageDimsIn() {
      const s = current.pageSetup;
      const dim = PAGE_INCHES[s.size] || PAGE_INCHES.Letter;
      return s.orientation === "landscape" ? { w: dim.h, h: dim.w } : { w: dim.w, h: dim.h };
    }
    // rendered px per inch (horizontal follows the drawn page width; vertical is 96*zoom)
    function scales() {
      const { w, h } = pageDimsIn();
      const pr = page.getBoundingClientRect();
      const zoom = parseFloat(getComputedStyle(page).zoom) || 1;
      return { w, h, pr, pxH: pr.width / w, pxV: 96 * zoom, firstPageH: h * 96 * zoom };
    }
    // cheap: keep the tracks overlaying the page as it scrolls
    function reposition() {
      const pr = page.getBoundingClientRect();
      track.style.left = pr.left - hbar.getBoundingClientRect().left + "px";
      vtrack.style.top = pr.top - vbar.getBoundingClientRect().top + "px";
    }
    function render() {
      const s = current.pageSetup;
      const m = s.margins || { top: 1, right: 1, bottom: 1, left: 1 };
      const { w, h, pr, pxH, pxV, firstPageH } = scales();

      // vertical bar overlays the editor viewport, anchored at the app's left edge
      const ar = app.getBoundingClientRect();
      const wr = wrap.getBoundingClientRect();
      vbar.style.top = wr.top - ar.top + "px";
      vbar.style.height = wr.height + "px";

      // ----- horizontal track = the page's x-range within the bar -----
      track.style.left = pr.left - hbar.getBoundingClientRect().left + "px";
      track.style.width = pr.width + "px";
      track.querySelectorAll(".ruler-tick").forEach((t) => t.remove());
      for (let i = 0; i <= Math.floor(w); i++) {
        const t = document.createElement("div");
        t.className = "ruler-tick"; t.style.left = i * pxH + "px"; t.textContent = i;
        track.appendChild(t);
      }
      const lpx = (m.left != null ? m.left : 1) * pxH;
      const rpx = (m.right != null ? m.right : 1) * pxH;
      shadeL.style.width = lpx + "px"; shadeR.style.width = rpx + "px";
      markL.style.left = lpx + "px"; markR.style.left = pr.width - rpx + "px";

      // ----- vertical track = the first page's y-range within the bar -----
      vtrack.style.top = pr.top - vbar.getBoundingClientRect().top + "px";
      vtrack.style.height = firstPageH + "px";
      vtrack.querySelectorAll(".ruler-vtick").forEach((t) => t.remove());
      for (let i = 0; i <= Math.floor(h); i++) {
        const t = document.createElement("div");
        t.className = "ruler-vtick"; t.style.top = i * pxV + "px"; t.textContent = i;
        vtrack.appendChild(t);
      }
      const tpx = (m.top != null ? m.top : 1) * pxV;
      const bpx = (m.bottom != null ? m.bottom : 1) * pxV;
      vshadeT.style.top = "0px"; vshadeT.style.height = tpx + "px";
      vshadeB.style.top = firstPageH - bpx + "px"; vshadeB.style.height = bpx + "px";
      vmarkT.style.top = tpx + "px"; vmarkB.style.top = firstPageH - bpx + "px";
    }

    function makeHDrag(marker, side) {
      marker.addEventListener("mousedown", (e) => {
        e.preventDefault();
        const { w, pxH } = scales();
        const rect = track.getBoundingClientRect();
        const totalPx = rect.width;
        const mv = (ev) => {
          const x = Math.max(0, Math.min(totalPx, ev.clientX - rect.left));
          const s = current.pageSetup;
          const margins = { ...(s.margins || { top: 1, right: 1, bottom: 1, left: 1 }) };
          if (side === "left") margins.left = Math.max(0.2, Math.min(x / pxH, w - margins.right - 0.5));
          else margins.right = Math.max(0.2, Math.min((totalPx - x) / pxH, w - margins.left - 0.5));
          margins.left = Math.round(margins.left * 100) / 100;
          margins.right = Math.round(margins.right * 100) / 100;
          applyPageSetup({ ...s, margins });
          render();
        };
        const up = () => { document.removeEventListener("mousemove", mv); document.removeEventListener("mouseup", up); scheduleSave(); };
        document.addEventListener("mousemove", mv);
        document.addEventListener("mouseup", up);
      });
    }
    function makeVDrag(marker, side) {
      marker.addEventListener("mousedown", (e) => {
        e.preventDefault();
        const { h, pxV, firstPageH } = scales();
        const rect = vtrack.getBoundingClientRect();
        const mv = (ev) => {
          const y = Math.max(0, Math.min(firstPageH, ev.clientY - rect.top));
          const s = current.pageSetup;
          const margins = { ...(s.margins || { top: 1, right: 1, bottom: 1, left: 1 }) };
          if (side === "top") margins.top = Math.max(0.2, Math.min(y / pxV, h - margins.bottom - 0.5));
          else margins.bottom = Math.max(0.2, Math.min((firstPageH - y) / pxV, h - margins.top - 0.5));
          margins.top = Math.round(margins.top * 100) / 100;
          margins.bottom = Math.round(margins.bottom * 100) / 100;
          applyPageSetup({ ...s, margins });
          render();
        };
        const up = () => { document.removeEventListener("mousemove", mv); document.removeEventListener("mouseup", up); scheduleSave(); };
        document.addEventListener("mousemove", mv);
        document.addEventListener("mouseup", up);
      });
    }
    makeHDrag(markL, "left");
    makeHDrag(markR, "right");
    makeVDrag(vmarkT, "top");
    makeVDrag(vmarkB, "bottom");
    wrap.addEventListener("scroll", reposition, { passive: true });
    window.addEventListener("resize", render);
    render();
    return { render };
  }

  function applyPageSetup(setup) {
    current.pageSetup = setup || { ...DEFAULT_PAGE_SETUP };
    const s = current.pageSetup;
    const dim = PAGE_INCHES[s.size] || PAGE_INCHES.Letter;
    let w = dim.w, h = dim.h;
    if (s.orientation === "landscape") [w, h] = [h, w];
    const m = s.margins || DEFAULT_PAGE_SETUP.margins;
    const page = $("page");
    page.style.width = `min(${w * 96}px, 100%)`;
    page.style.minHeight = h * 96 + "px";
    page.style.padding = `${m.top * 96}px ${m.right * 96}px ${m.bottom * 96}px ${m.left * 96}px`;
    // print CSS
    let st = $("print-style");
    if (!st) {
      st = document.createElement("style");
      st.id = "print-style";
      document.head.appendChild(st);
    }
    st.textContent = `@page { size: ${s.size === "Letter" || s.size === "Legal" ? s.size.toLowerCase() : s.size} ${s.orientation}; margin: ${m.top}in ${m.right}in ${m.bottom}in ${m.left}in; }`;
    schedulePaginate();
  }

  function setEditorContent(html) {
    editor.innerHTML = html && html.trim() ? sanitizeHtml(html) : "<p><br></p>";
    if (editHistory) editHistory.reset();
    updateWordCount();
    schedulePaginate();
  }

  function updateWordCount() {
    const { words, chars } = countWords(editor);
    $("wc").textContent = `${words} word${words === 1 ? "" : "s"} · ${chars} chars`;
  }

  // ---- autosave + realtime ----
  const emitToHost = (event, data) => {
    if (window.parent !== window) {
      try { window.parent.postMessage({ we: 1, event, data }, "*"); } catch {}
    }
  };

  const track = createTrackChanges(editor, () => userName);

  const autosaver = new Autosaver(null, 1200, setStatus,
    (doc) => { // onSaved
      $("doc-info").textContent = `rev ${doc.rev}`;
      sync.sendUpdate({
        rev: doc.rev, title: doc.title, state: doc.state, pageSetup: doc.pageSetup,
        comments: doc.comments, trackChanges: doc.trackChanges,
      });
      emitToHost("save", { id: doc.id, rev: doc.rev, title: doc.title });
    },
    (server) => showConflictBanner(server) // onConflict (409)
  );

  const sync = new SyncClient(userName, {
    onPresence(users) {
      const host = $("presence");
      host.innerHTML = "";
      for (const u of users) {
        const chip = document.createElement("span");
        chip.className = "presence-chip";
        chip.title = u.user;
        chip.style.background = u.color;
        chip.textContent = (u.user || "?").trim()[0].toUpperCase();
        host.appendChild(chip);
      }
      emitToHost("presence", { users });
    },
    onUpdate(msg) {
      if (msg.rev != null && msg.rev <= autosaver.rev) return;
      if (autosaver.dirty || autosaver.inFlight) {
        showConflictBanner({ rev: msg.rev, title: msg.title, state: msg.state, pageSetup: msg.pageSetup, comments: msg.comments }, msg.from);
        return;
      }
      // clean local state: follow the remote edit
      autosaver.suspended = true;
      if (msg.title != null && msg.title !== titleInput.value) titleInput.value = msg.title;
      if (msg.pageSetup) applyPageSetup(msg.pageSetup);
      if (msg.state != null) setEditorContent(msg.state);
      if (msg.comments) { docComments = msg.comments; renderComments(); }
      if (msg.trackChanges != null) { track.setEnabled(!!msg.trackChanges); syncTrackUI(); }
      if (msg.rev != null) autosaver.setRev(msg.rev);
      autosaver.suspended = false;
      setStatus("synced");
      emitToHost("change", { source: "remote" });
    },
  });
  sync.connect();

  function showConflictBanner(server, from) {
    const banner = $("banner");
    banner.classList.remove("hidden");
    banner.innerHTML = "";
    const who = from ? `${from.user}` : "another session";
    banner.appendChild(document.createTextNode(`This document was changed by ${who}. `));
    const loadBtn = document.createElement("button");
    loadBtn.textContent = "Load their version";
    loadBtn.addEventListener("click", () => {
      autosaver.suspended = true;
      if (server.title != null) titleInput.value = server.title;
      if (server.pageSetup) applyPageSetup(server.pageSetup);
      setEditorContent(server.state || "");
      if (server.comments) { docComments = server.comments; renderComments(); }
      autosaver.setRev(server.rev || 0);
      autosaver.dirty = false;
      autosaver.suspended = false;
      banner.classList.add("hidden");
      setStatus("synced");
    });
    const keepBtn = document.createElement("button");
    keepBtn.textContent = "Keep mine (overwrite)";
    keepBtn.addEventListener("click", () => {
      autosaver.setRev(server.rev || 0);
      banner.classList.add("hidden");
      scheduleSave();
      autosaver.flush();
    });
    banner.append(loadBtn, keepBtn);
  }

  const scheduleSave = () => {
    autosaver.update({
      title: titleInput.value, state: getCleanHtml(), pageSetup: current.pageSetup,
      comments: docComments, trackChanges: track.isEnabled(),
    });
  };

  // ---------------------------------------------------------------
  // Comments
  // ---------------------------------------------------------------
  const commentsPanel = $("comments-panel");
  const commentsItems = $("comments-items");

  function commentAnchors(cid) {
    return editor.querySelectorAll(`.comment-ref[data-cid="${CSS.escape(cid)}"]`);
  }
  function jumpToComment(cid) {
    const anchors = commentAnchors(cid);
    if (!anchors.length) return;
    anchors[0].scrollIntoView({ block: "center", behavior: "smooth" });
    for (const a of anchors) a.classList.add("active");
    setTimeout(() => { for (const a of anchors) a.classList.remove("active"); }, 1600);
  }
  function renderComments(focusCid) {
    commentsItems.innerHTML = "";
    if (!docComments.length) {
      const li = document.createElement("li");
      li.className = "muted";
      li.textContent = "No comments yet — select text and right-click → Add comment.";
      commentsItems.appendChild(li);
      return;
    }
    for (const c of docComments) {
      const li = document.createElement("li");
      const card = document.createElement("div");
      card.className = "comment-card" + (c.resolved ? " resolved" : "") + (c.id === focusCid ? " focus" : "");
      const head = document.createElement("div");
      head.className = "comment-head";
      const author = document.createElement("span");
      author.className = "comment-author";
      author.textContent = c.author || "Unknown";
      const when = document.createElement("span");
      when.textContent = (c.resolved ? "resolved · " : "") + fmtTime(c.createdAt);
      head.append(author, when);
      const text = document.createElement("div");
      text.className = "comment-text";
      text.textContent = c.text;
      card.append(head, text);
      if (!commentAnchors(c.id).length) {
        const orphan = document.createElement("div");
        orphan.className = "comment-orphan";
        orphan.textContent = "(anchor text was removed)";
        card.appendChild(orphan);
      }
      for (const r of c.replies || []) {
        const rd = document.createElement("div");
        rd.className = "comment-reply";
        const ra = document.createElement("b");
        ra.textContent = (r.author || "?") + ": ";
        rd.appendChild(ra);
        rd.appendChild(document.createTextNode(r.text));
        card.appendChild(rd);
      }
      const actions = document.createElement("div");
      actions.className = "comment-actions";
      const replyBtn = document.createElement("button");
      replyBtn.textContent = "Reply";
      replyBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        const input = document.createElement("input");
        input.type = "text";
        input.placeholder = "Reply…";
        input.className = "comment-reply-input";
        actions.before(input);
        input.focus();
        input.addEventListener("keydown", (ev) => {
          if (ev.key === "Enter" && input.value.trim()) {
            c.replies = c.replies || [];
            c.replies.push({ author: userName, text: input.value.trim(), createdAt: Date.now() });
            renderComments(c.id);
            scheduleSave();
          } else if (ev.key === "Escape") input.remove();
        });
      });
      const resolveBtn = document.createElement("button");
      resolveBtn.textContent = c.resolved ? "Reopen" : "Resolve";
      resolveBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        c.resolved = !c.resolved;
        for (const a of commentAnchors(c.id)) a.classList.toggle("resolved", c.resolved);
        renderComments(c.id);
        scheduleSave();
      });
      const delBtn = document.createElement("button");
      delBtn.textContent = "Delete";
      delBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        if (!confirm("Delete this comment?")) return;
        for (const a of [...commentAnchors(c.id)]) {
          const p = a.parentNode;
          while (a.firstChild) p.insertBefore(a.firstChild, a);
          a.remove();
          p.normalize();
        }
        docComments = docComments.filter((x) => x.id !== c.id);
        renderComments();
        scheduleSave();
      });
      actions.append(replyBtn, resolveBtn, delBtn);
      card.appendChild(actions);
      card.addEventListener("click", () => jumpToComment(c.id));
      li.appendChild(card);
      commentsItems.appendChild(li);
    }
  }
  function openCommentsPanel(focusCid) {
    closePanels("comments-panel");
    commentsPanel.classList.remove("hidden");
    renderComments(focusCid);
    if (focusCid) jumpToComment(focusCid);
  }
  function addCommentFlow() {
    const ta = document.createElement("textarea");
    ta.rows = 3;
    ta.style.cssText = "width:100%;box-sizing:border-box;font:13px/1.4 inherit;padding:6px";
    ta.placeholder = "Write a comment…";
    openDialog("Add comment", ta, [
      { label: "Cancel" },
      {
        label: "Comment", primary: true,
        onClick: () => {
          const text = ta.value.trim();
          if (!text) return false;
          const cid = "c" + Date.now().toString(36) + Math.random().toString(36).slice(2, 5);
          if (!wrapSelectionComment(editor, cid)) {
            alert("Select some text to comment on first.");
            return;
          }
          docComments.push({ id: cid, author: userName, text, createdAt: Date.now(), resolved: false, replies: [] });
          openCommentsPanel(cid);
          scheduleSave();
        },
      },
    ]);
  }
  // clicking an anchor opens its comment
  editor.addEventListener("click", (e) => {
    const ref = e.target.closest && e.target.closest("span.comment-ref");
    if (ref && editor.contains(ref)) openCommentsPanel(ref.getAttribute("data-cid"));
  });

  // ---------------------------------------------------------------
  // Review (track changes) panel
  // ---------------------------------------------------------------
  const reviewPanel = $("review-panel");
  const reviewItems = $("review-items");
  const trackToggle = $("track-toggle");

  function syncTrackUI() {
    trackToggle.checked = track.isEnabled();
    const tb = toolbarHost.querySelector(".toolbar");
    if (tb && tb.trackBtn) tb.trackBtn.classList.toggle("active", track.isEnabled());
  }
  function renderReview() {
    reviewItems.innerHTML = "";
    const changes = track.list();
    if (!changes.length) {
      const li = document.createElement("li");
      li.className = "muted";
      li.textContent = track.isEnabled()
        ? "No pending changes. Edits are now being recorded."
        : "No pending changes. Turn on Track changes to record edits.";
      reviewItems.appendChild(li);
      return;
    }
    for (const ch of changes) {
      const li = document.createElement("li");
      const card = document.createElement("div");
      card.className = "review-item " + (ch.type === "insertion" ? "ins" : "del");
      const head = document.createElement("div");
      head.className = "comment-head";
      const badge = document.createElement("span");
      badge.className = "badge";
      badge.textContent = (ch.type === "insertion" ? "➕ inserted" : "➖ deleted") + " · " + ch.author;
      const when = document.createElement("span");
      when.textContent = ch.ts ? fmtTime(ch.ts) : "";
      head.append(badge, when);
      const snippet = document.createElement("span");
      snippet.className = "snippet";
      snippet.textContent = `“${ch.text}${ch.text.length >= 60 ? "…" : ""}”`;
      const actions = document.createElement("div");
      actions.className = "comment-actions";
      const acc = document.createElement("button");
      acc.textContent = "Accept";
      acc.addEventListener("click", (e) => { e.stopPropagation(); track.accept(ch.node); renderReview(); });
      const rej = document.createElement("button");
      rej.textContent = "Reject";
      rej.addEventListener("click", (e) => { e.stopPropagation(); track.reject(ch.node); renderReview(); });
      actions.append(acc, rej);
      card.append(head, snippet, actions);
      card.addEventListener("click", () => {
        ch.node.scrollIntoView({ block: "center", behavior: "smooth" });
        ch.node.classList.add("flash");
        setTimeout(() => ch.node.classList.remove("flash"), 1200);
      });
      li.appendChild(card);
      reviewItems.appendChild(li);
    }
  }
  function openReviewPanel() {
    closePanels("review-panel");
    reviewPanel.classList.remove("hidden");
    syncTrackUI();
    renderReview();
  }
  trackToggle.addEventListener("change", () => {
    track.setEnabled(trackToggle.checked);
    syncTrackUI();
    renderReview();
    scheduleSave();
  });
  $("btn-accept-all").addEventListener("click", () => {
    if (!track.list().length) return;
    if (!confirm("Accept all tracked changes?")) return;
    track.acceptAll();
    renderReview();
  });
  $("btn-reject-all").addEventListener("click", () => {
    if (!track.list().length) return;
    if (!confirm("Reject all tracked changes?")) return;
    track.rejectAll();
    renderReview();
  });

  // ---------------------------------------------------------------
  // Panels bookkeeping
  // ---------------------------------------------------------------
  const PANELS = ["library", "history", "comments-panel", "review-panel"];
  function closePanels(except) {
    for (const id of PANELS) {
      if (id !== except) $(id).classList.add("hidden");
    }
  }

  // ---- toolbar / behaviors ----
  const findPanel = createFindPanel(editor, $("find-host"));
  if (showToolbar && mode === "edit") {
    toolbarHost.appendChild(buildToolbar(editor, {
      history: editHistory,
      toggleFind: () => findPanel.toggle(),
      print: () => window.print(),
      pageSetup: () => openPageSetupDialog(current.pageSetup, (setup) => {
        applyPageSetup(setup);
        scheduleSave();
      }),
      addComment: () => addCommentFlow(),
      toggleTrack: () => {
        track.setEnabled(!track.isEnabled());
        syncTrackUI();
        scheduleSave();
        return track.isEnabled();
      },
    }));
  }
  attachEditorBehaviors(editor, { track, history: editHistory });
  attachTableToolbar(editor, $("page-outer"));
  attachImageEditing(editor, $("page"));
  ruler = attachRuler();
  initTooltips();
  attachContextMenu(editor, {
    addComment: () => addCommentFlow(),
    openComment: (cid) => openCommentsPanel(cid),
    acceptChange: (node) => { track.accept(node); if (!reviewPanel.classList.contains("hidden")) renderReview(); },
    rejectChange: (node) => { track.reject(node); if (!reviewPanel.classList.contains("hidden")) renderReview(); },
  });

  if (mode === "view") {
    editor.contentEditable = "false";
    toolbarHost.classList.add("hidden");
    document.body.classList.add("view-mode");
  }

  // ---- load initial document ----
  async function loadDocument(id) {
    const meta = await getDocument(id);
    if (!meta) return false;
    autosaver.setId(meta.id, meta.rev || 0);
    current = { id: meta.id, title: meta.title, pageSetup: meta.pageSetup || { ...DEFAULT_PAGE_SETUP } };
    docComments = meta.comments || [];
    track.setEnabled(!!meta.trackChanges);
    syncTrackUI();
    titleInput.value = meta.title;
    applyPageSetup(current.pageSetup);
    setEditorContent(meta.state);
    renderComments();
    $("doc-info").textContent = `rev ${meta.rev || 0}`;
    sync.join(meta.id);
    setStatus("ready");
    if (!embedded) localStorage.setItem(LS_KEY, meta.id);
    return true;
  }

  const requestedId = params.get("doc");
  let loaded = false;
  if (requestedId) loaded = await loadDocument(requestedId);
  if (!loaded) {
    const savedId = localStorage.getItem(LS_KEY);
    if (savedId) loaded = await loadDocument(savedId);
  }
  if (!loaded) {
    const meta = await createDocument("Untitled");
    await loadDocument(meta.id);
  }

  let wcTimer = null;
  editor.addEventListener("input", () => {
    scheduleSave();
    schedulePaginate();
    clearTimeout(wcTimer);
    wcTimer = setTimeout(() => {
      updateWordCount();
      if (!reviewPanel.classList.contains("hidden")) renderReview();
    }, 300);
    emitToHost("change", { source: "local" });
  });
  titleInput.addEventListener("input", scheduleSave);

  // ---- New ----
  $("btn-new").addEventListener("click", async () => {
    await autosaver.flush();
    const meta = await createDocument("Untitled");
    await loadDocument(meta.id);
    editor.focus();
  });

  // ---- Open file ----
  $("btn-open").addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", async () => {
    const file = fileInput.files && fileInput.files[0];
    if (!file) return;
    const name = file.name.toLowerCase();
    // PDF: open in read-only viewer
    if (name.endsWith(".pdf")) {
      setStatus("saving");
      try {
        if (isPdfMode()) closePdf();
        await openPdf(file, {
          title: file.name.replace(/\.pdf$/i, ""),
          setTitle: (t) => { titleInput.value = t; },
          setStatus: (s) => setStatus(s),
          onClosed: () => { updateWordCount(); },
          fitWidthDefault: true,
        });
        setStatus("saved");
        const info = getPdfInfo();
        if (info) $("pagecount").textContent = `${info.numPages} page${info.numPages === 1 ? "" : "s"} (PDF)`;
      } catch (e) {
        console.error(e);
        setStatus("error");
        alert("Could not open PDF: " + e.message);
      } finally {
        fileInput.value = "";
      }
      return;
    }
    // Editing formats: close PDF view if active and edit normally
    if (isPdfMode()) closePdf();
    setStatus("saving");
    try {
      let html, pageSetup = { ...DEFAULT_PAGE_SETUP }, comments = [];
      if (name.endsWith(".txt")) {
        html = plainTextToHtml(await file.text());
      } else if (name.endsWith(".html") || name.endsWith(".htm")) {
        html = sanitizeHtml(await file.text());
      } else {
        const res = await importDocx(file);
        html = res.html;
        pageSetup = res.pageSetup;
        comments = res.comments || [];
      }
      const title = file.name.replace(/\.(docx|txt|html?|pdf)$/i, "");
      const seeded = name.endsWith(".docx")
        ? await importDocxFile(file)
        : await createDocument(title);
      autosaver.setId(seeded.id, seeded.rev || 0);
      current = { id: seeded.id, title, pageSetup };
      docComments = comments;
      titleInput.value = title;
      applyPageSetup(pageSetup);
      setEditorContent(html);
      renderComments();
      const doc = await saveDocument(seeded.id, { title, state: getCleanHtml(), pageSetup, comments });
      autosaver.setRev(doc.rev);
      if (!embedded) localStorage.setItem(LS_KEY, seeded.id);
      sync.join(seeded.id);
      setStatus("saved");
    } catch (e) {
      console.error(e);
      setStatus("error");
      alert("Could not open file: " + e.message);
    } finally {
      fileInput.value = "";
    }
  });

  // ---- File menu ----
  const fileMenu = $("file-menu");
  function closeFileMenu() { fileMenu.classList.add("hidden"); exportMenu.classList.add("hidden"); }
  if ($("btn-file")) $("btn-file").addEventListener("click", (e) => {
    e.stopPropagation();
    fileMenu.classList.toggle("hidden");
    exportMenu.classList.add("hidden");
    const im = $("insert-menu");
    if (im) im.classList.add("hidden"); // clicking File closes the Insert menu
  });
  document.addEventListener("click", (e) => {
    if (!$("file-menu-wrap").contains(e.target)) closeFileMenu();
  });
  if ($("btn-print")) $("btn-print").addEventListener("click", () => {
    closeFileMenu();
    if (isPdfMode()) { const info = getPdfInfo(); if (info) { /* delegate to pdf-view's print via its toolbar button */ const btn = document.querySelector("#pdf-toolbar button[title='Print']"); if (btn) btn.click(); return; } }
    window.print();
  });
  if ($("btn-close")) $("btn-close").addEventListener("click", async () => {
    closeFileMenu();
    // PDF: close viewer and return to editor
    if (isPdfMode()) { closePdf(); return; }
    // Document: flush, clear editor, create a fresh blank doc
    await autosaver.flush();
    if (current && current.id) sync.leave(current.id);
    const meta = await createDocument("Untitled");
    await loadDocument(meta.id);
    localStorage.removeItem(LS_KEY);
    editor.focus();
  });
  fileMenu.addEventListener("click", (e) => {
    const b = e.target.closest("button");
    if (!b || b.id === "btn-export") return; // keep menu open to show the Export flyout
    closeFileMenu();
  });

  // ---- Export menu ----
  const exportMenu = $("export-menu");
  $("btn-export").addEventListener("click", (e) => { e.stopPropagation(); exportMenu.classList.toggle("hidden"); });
  document.addEventListener("click", (e) => {
    if (!$("export-menu-wrap").contains(e.target)) exportMenu.classList.add("hidden");
  });
  async function doExport(fmt) {
    const title = titleInput.value || "document";
    setStatus("saving");
    try {
      if (fmt === "docx") {
        const blob = await buildDocxFromHtml(getCleanHtml(), { title, pageSetup: current.pageSetup, comments: docComments });
        saveAs(blob, title + ".docx");
        await putDocx(current.id, blob);
      } else if (fmt === "html") {
        saveAs(new Blob([exportStandaloneHtml(getCleanHtml(), title)], { type: "text/html" }), title + ".html");
      } else if (fmt === "txt") {
        saveAs(new Blob([htmlToPlainText(getCleanHtml())], { type: "text/plain" }), title + ".txt");
      } else if (fmt === "pdf") {
        window.print();
      }
      setStatus("saved");
    } catch (e) {
      console.error(e);
      setStatus("error");
      alert("Export failed: " + e.message);
    }
  }
  exportMenu.addEventListener("click", (e) => {
    const b = e.target.closest("button[data-fmt]");
    if (!b) return;
    exportMenu.classList.add("hidden");
    doExport(b.dataset.fmt);
  });

  // ---- Insert menu (top bar) ----
  const insertMenu = $("insert-menu");
  $("btn-insert").addEventListener("click", (e) => {
    e.stopPropagation();
    closeFileMenu();
    insertMenu.classList.toggle("hidden");
  });
  document.addEventListener("click", (e) => {
    if (!$("insert-menu-wrap").contains(e.target)) insertMenu.classList.add("hidden");
  });
  insertMenu.addEventListener("click", (e) => {
    const b = e.target.closest("button[data-ins]");
    if (!b) return;
    insertMenu.classList.add("hidden");
    const a = b.dataset.ins;
    if (a === "pagebreak") insertPageBreak(editor);
    else if (a === "blankpage") insertBlankPage(editor);
    else if (a === "picture") insertImage(editor);
    else if (a === "table") openTableDialog(editor);
    else if (a === "link") openLinkDialog(editor);
    else if (a === "hr") insertHtmlAtCaret(editor, "<hr>");
    else if (a === "symbol") openSymbolDialog(editor);
    else if (a === "wordart") openWordArtDialog(editor);
    else if (a === "shape") openShapeDialog(editor);
    else if (a === "header") openHeaderFooterDialog("header");
    else if (a === "footer") openHeaderFooterDialog("footer");
    else if (a === "pagenum") openPageNumberDialog();
  });

  // header / footer editing dialog
  function dlgField(labelText, control) {
    const l = document.createElement("label");
    l.className = "dlg-field";
    const s = document.createElement("span");
    s.textContent = labelText;
    l.append(s, control);
    return l;
  }
  function makeSelect(pairs, value) {
    const sel = document.createElement("select");
    for (const [v, label] of pairs) {
      const o = document.createElement("option");
      o.value = v; o.textContent = label;
      if (value === v) o.selected = true;
      sel.appendChild(o);
    }
    return sel;
  }
  function openHeaderFooterDialog(which) {
    const ch = getChrome();
    const cur = ch[which] || { text: "", align: "center" };
    const textIn = document.createElement("input");
    textIn.type = "text"; textIn.value = cur.text || "";
    textIn.placeholder = "Header/footer text";
    const alignSel = makeSelect([["left", "Left"], ["center", "Center"], ["right", "Right"]], cur.align || "center");
    const body = document.createElement("div");
    body.append(dlgField("Text", textIn), dlgField("Alignment", alignSel));
    const hint = document.createElement("div");
    hint.className = "dlg-hint";
    hint.textContent = `Shown in the ${which} of every page. Clear the text and Apply to remove it.`;
    body.appendChild(hint);
    openDialog(which === "header" ? "Edit header" : "Edit footer", body, [
      { label: "Cancel" },
      {
        label: "Apply", primary: true,
        onClick: () => {
          const text = textIn.value.trim();
          if (text) ch[which] = { text, align: alignSel.value };
          else delete ch[which];
          schedulePaginate();
          scheduleSave();
        },
      },
    ]);
  }
  function openPageNumberDialog() {
    const ch = getChrome();
    const cur = ch.pageNumber || { enabled: true, format: "arabic", place: "footer-center" };
    const enCk = document.createElement("input");
    enCk.type = "checkbox"; enCk.checked = cur.enabled !== false;
    const enWrap = document.createElement("label");
    enWrap.className = "dlg-field dlg-check";
    const enSpan = document.createElement("span");
    enSpan.textContent = "Show page numbers";
    enWrap.append(enCk, enSpan);
    const fmtSel = makeSelect([
      ["arabic", "1, 2, 3"], ["roman", "i, ii, iii"], ["alpha", "a, b, c"],
      ["page", "Page 1"], ["pageOfN", "Page 1 of N"],
    ], cur.format || "arabic");
    const placeSel = makeSelect([
      ["header-left", "Top left"], ["header-center", "Top center"], ["header-right", "Top right"],
      ["footer-left", "Bottom left"], ["footer-center", "Bottom center"], ["footer-right", "Bottom right"],
    ], cur.place || "footer-center");
    const body = document.createElement("div");
    body.append(enWrap, dlgField("Format", fmtSel), dlgField("Position", placeSel));
    openDialog("Page numbers", body, [
      { label: "Cancel" },
      {
        label: "Apply", primary: true,
        onClick: () => {
          ch.pageNumber = { enabled: enCk.checked, format: fmtSel.value, place: placeSel.value };
          schedulePaginate();
          scheduleSave();
        },
      },
    ]);
  }

  // ---- Library ----
  const library = $("library");
  const libraryItems = $("library-items");
  $("btn-list").addEventListener("click", async () => {
    if (!library.classList.contains("hidden")) { library.classList.add("hidden"); return; }
    closePanels("library");
    library.classList.remove("hidden");
    await renderLibrary();
  });
  async function renderLibrary() {
    libraryItems.innerHTML = "";
    const docs = await listDocuments();
    if (!docs.length) {
      const li = document.createElement("li");
      li.textContent = "No saved documents yet.";
      li.className = "muted";
      libraryItems.appendChild(li);
      return;
    }
    for (const d of docs) {
      const li = document.createElement("li");
      li.className = "lib-item" + (d.id === current.id ? " current" : "");
      const open = document.createElement("button");
      open.className = "lib-open";
      open.innerHTML = `<span class="lib-title"></span><span class="lib-date"></span>`;
      open.querySelector(".lib-title").textContent = d.title || "Untitled";
      open.querySelector(".lib-date").textContent = fmtTime(d.updatedAt);
      open.addEventListener("click", async () => {
        await autosaver.flush();
        await loadDocument(d.id);
        editor.focus();
      });
      const del = document.createElement("button");
      del.className = "lib-del";
      del.title = "Delete document";
      del.textContent = "🗑";
      del.addEventListener("click", async (e) => {
        e.stopPropagation();
        if (!confirm(`Delete "${d.title}"? This cannot be undone.`)) return;
        await deleteDocument(d.id);
        if (d.id === current.id) {
          const meta = await createDocument("Untitled");
          await loadDocument(meta.id);
        }
        await renderLibrary();
      });
      li.append(open, del);
      libraryItems.appendChild(li);
    }
  }

  // ---- Version history ----
  const history = $("history");
  const historyItems = $("history-items");
  $("btn-history").addEventListener("click", async () => {
    if (!history.classList.contains("hidden")) { history.classList.add("hidden"); return; }
    closePanels("history");
    history.classList.remove("hidden");
    await renderHistory();
  });
  async function renderHistory() {
    historyItems.innerHTML = "";
    const versions = await listVersions(current.id);
    if (!versions.length) {
      const li = document.createElement("li");
      li.textContent = "No versions yet — versions are snapshotted as you edit.";
      li.className = "muted";
      historyItems.appendChild(li);
      return;
    }
    for (const v of [...versions].reverse()) {
      const li = document.createElement("li");
      li.className = "lib-item";
      const info = document.createElement("button");
      info.className = "lib-open";
      info.innerHTML = `<span class="lib-title"></span><span class="lib-date"></span>`;
      info.querySelector(".lib-title").textContent = `rev ${v.rev} — ${v.title}`;
      info.querySelector(".lib-date").textContent = fmtTime(v.t);
      info.addEventListener("click", async () => {
        if (!confirm(`Restore version from ${fmtTime(v.t)}? Current content is snapshotted first.`)) return;
        const doc = await restoreVersion(current.id, v.index);
        autosaver.setRev(doc.rev);
        titleInput.value = doc.title;
        if (doc.pageSetup) applyPageSetup(doc.pageSetup);
        setEditorContent(doc.state);
        docComments = doc.comments || [];
        renderComments();
        setStatus("saved");
        history.classList.add("hidden");
      });
      li.appendChild(info);
      historyItems.appendChild(li);
    }
  }

  // ---- Comments / Review buttons ----
  $("btn-comments").addEventListener("click", () => {
    if (!commentsPanel.classList.contains("hidden")) { commentsPanel.classList.add("hidden"); return; }
    openCommentsPanel();
  });
  $("btn-review").addEventListener("click", () => {
    if (!reviewPanel.classList.contains("hidden")) { reviewPanel.classList.add("hidden"); return; }
    openReviewPanel();
  });

  // ---- statusbar: zoom + spellcheck ----
  $("zoom").addEventListener("change", () => {
    $("page").style.zoom = $("zoom").value;
    schedulePaginate();
  });
  $("spellcheck-toggle").addEventListener("change", (e) => {
    editor.spellcheck = e.target.checked;
    editor.blur(); editor.focus();
  });

  // ---- keyboard shortcuts ----
  window.addEventListener("keydown", (e) => {
    // Esc closes the PDF viewer and returns to editing
    // Esc in PDF mode is handled by the pdf-view module (deselect / cancel placement) — no auto-close
    const mod = e.metaKey || e.ctrlKey;
    if (!mod) return;
    const k = e.key.toLowerCase();
    if (k === "s") {
      e.preventDefault();
      if (e.shiftKey) doExport("docx");
      else { scheduleSave(); autosaver.flush(); }
    } else if (k === "f" && !e.shiftKey) {
      e.preventDefault();
      findPanel.toggle();
    } else if (k === "p") {
      e.preventDefault();
      window.print();
    } else if (k === "m" && e.altKey) {
      e.preventDefault();
      addCommentFlow(); // Ctrl/⌘+Alt+M — Word's add-comment shortcut
    } else if (k === "e" && e.shiftKey) {
      e.preventDefault(); // Ctrl/⌘+Shift+E — Word's track-changes toggle
      track.setEnabled(!track.isEnabled());
      syncTrackUI();
      scheduleSave();
    }
  });

  window.addEventListener("beforeunload", () => { autosaver.flush(); });

  // ---------------------------------------------------------------
  // postMessage API (JS SDK bridge)
  // ---------------------------------------------------------------
  window.addEventListener("message", async (ev) => {
    const m = ev.data;
    if (!m || m.we !== 1 || !m.cmd) return;
    const reply = (result, error) => {
      try { ev.source.postMessage({ we: 1, re: m.id, result, error }, "*"); } catch {}
    };
    try {
      switch (m.cmd) {
        case "getContent": reply({ html: getCleanHtml() }); break;
        case "getText": reply({ text: htmlToPlainText(getCleanHtml()) }); break;
        case "setContent": setEditorContent(m.args && m.args.html || ""); scheduleSave(); reply({ ok: true }); break;
        case "insertText": insertTextAtCaret(editor, (m.args && m.args.text) || ""); scheduleSave(); reply({ ok: true }); break;
        case "insertHtml": insertHtmlAtCaret(editor, sanitizeHtml((m.args && m.args.html) || "")); scheduleSave(); reply({ ok: true }); break;
        case "getMeta": reply({ id: current.id, title: titleInput.value, rev: autosaver.rev, pageSetup: current.pageSetup, trackChanges: track.isEnabled(), commentCount: docComments.length, ...countWords(editor) }); break;
        case "setTitle": titleInput.value = String((m.args && m.args.title) || ""); scheduleSave(); reply({ ok: true }); break;
        case "save": { scheduleSave(); await autosaver.flush(); reply({ ok: true, rev: autosaver.rev }); break; }
        case "loadDocument": reply({ ok: await loadDocument(m.args && m.args.id) }); break;
        case "setMode": {
          const v = m.args && m.args.mode === "view";
          editor.contentEditable = v ? "false" : "true";
          toolbarHost.classList.toggle("hidden", v || !showToolbar);
          reply({ ok: true });
          break;
        }
        case "find": reply({ matches: findPanel.find((m.args && m.args.query) || "", m.args || {}) }); break;
        case "replaceAll": reply({ replaced: findPanel.replaceAll((m.args && m.args.query) || "", (m.args && m.args.replacement) || "", m.args || {}) }); break;
        case "focus": editor.focus(); reply({ ok: true }); break;
        case "addComment": {
          const text = ((m.args && m.args.text) || "").trim();
          if (!text) { reply(null, "text required"); break; }
          const cid = "c" + Date.now().toString(36) + Math.random().toString(36).slice(2, 5);
          if (!wrapSelectionComment(editor, cid)) { reply(null, "no selection to anchor the comment"); break; }
          docComments.push({ id: cid, author: (m.args && m.args.author) || userName, text, createdAt: Date.now(), resolved: false, replies: [] });
          renderComments();
          scheduleSave();
          reply({ id: cid });
          break;
        }
        case "getComments": reply({ comments: docComments }); break;
        case "setTrackChanges": track.setEnabled(!!(m.args && m.args.on)); syncTrackUI(); scheduleSave(); reply({ ok: true, on: track.isEnabled() }); break;
        case "getChanges": reply({ changes: track.list().map(({ node, ...rest }) => rest) }); break;
        case "acceptAllChanges": track.acceptAll(); scheduleSave(); reply({ ok: true }); break;
        case "rejectAllChanges": track.rejectAll(); scheduleSave(); reply({ ok: true }); break;
        case "undo": if (editHistory) editHistory.undo(); reply({ ok: true, canUndo: editHistory && editHistory.canUndo(), canRedo: editHistory && editHistory.canRedo() }); break;
        case "redo": if (editHistory) editHistory.redo(); reply({ ok: true, canUndo: editHistory && editHistory.canUndo(), canRedo: editHistory && editHistory.canRedo() }); break;
        case "canUndo": reply({ canUndo: !!(editHistory && editHistory.canUndo()), canRedo: !!(editHistory && editHistory.canRedo()) }); break;

        // ---- formatting: execCommand-based ----
        case "format": {
          const cmd = String((m.args && m.args.cmd) || "");
          const val = m.args && m.args.value != null ? String(m.args.value) : null;
          if (!cmd || !/^[a-zA-Z]+$/.test(cmd)) { reply(null, "invalid format cmd"); break; }
          editor.focus();
          let ok = false;
          try { ok = document.execCommand(cmd, false, val); } catch (e) { reply(null, String(e.message || e)); break; }
          scheduleSave(); schedulePaginate();
          reply({ ok, html: getCleanHtml() });
          break;
        }
        case "getSelectedText": {
          const sel = window.getSelection();
          reply({ text: sel ? sel.toString() : "", html: sel && sel.rangeCount ? sel.getRangeAt(0).cloneContents().textContent : "" });
          break;
        }

        // ---- inserts that don't need a dialog ----
        case "insertImage": {
          const src = String((m.args && m.args.src) || "");
          if (!src) { reply(null, "src required"); break; }
          const alt = String((m.args && m.args.alt) || "");
          const w = m.args && m.args.width != null ? Number(m.args.width) : null;
          const h = m.args && m.args.height != null ? Number(m.args.height) : null;
          let img = `<img src="${src.replace(/"/g, "&quot;")}"${alt ? ` alt="${alt.replace(/"/g, "&quot;")}"` : ""}${w != null ? ` width="${w}"` : ""}${h != null ? ` height="${h}"` : ""}>`;
          insertHtmlAtCaret(editor, img); scheduleSave(); schedulePaginate(); reply({ ok: true });
          break;
        }
        case "insertLink": {
          const href = String((m.args && m.args.href) || "");
          if (!href) { reply(null, "href required"); break; }
          const text = String((m.args && m.args.text) || href);
          insertHtmlAtCaret(editor, `<a href="${href.replace(/"/g, "&quot;")}">${text.replace(/[<>&]/g, (c) => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;" })[c])}</a>`);
          scheduleSave(); reply({ ok: true });
          break;
        }
        case "insertTable": {
          const rows = Math.max(1, Math.min(50, parseInt(m.args && m.args.rows, 10) || 2));
          const cols = Math.max(1, Math.min(20, parseInt(m.args && m.args.cols, 10) || 2));
          let html = '<table><tbody>';
          for (let r = 0; r < rows; r++) {
            html += "<tr>";
            for (let c = 0; c < cols; c++) html += "<td><br></td>";
            html += "</tr>";
          }
          html += "</tbody></table><p><br></p>";
          insertHtmlAtCaret(editor, html); scheduleSave(); schedulePaginate(); reply({ ok: true });
          break;
        }
        case "insertSymbol": {
          const ch = String((m.args && m.args.char) || "");
          if (!ch) { reply(null, "char required"); break; }
          insertTextAtCaret(editor, ch); scheduleSave(); reply({ ok: true });
          break;
        }
        case "insertPageBreak": insertPageBreak(editor); scheduleSave(); reply({ ok: true }); break;
        case "insertBlankPage": insertBlankPage(editor); scheduleSave(); reply({ ok: true }); break;
        case "insertHr": insertHtmlAtCaret(editor, "<hr>"); scheduleSave(); reply({ ok: true }); break;

        // ---- headers / footers / page numbers ----
        case "setHeader": case "setFooter": {
          const which = m.cmd === "setHeader" ? "header" : "footer";
          const ch = getChrome();
          const text = (m.args && m.args.text != null) ? String(m.args.text) : null;
          const align = (m.args && m.args.align) || "center";
          if (text === null || text === "") delete ch[which];
          else ch[which] = { text, align: ["left", "center", "right"].includes(align) ? align : "center" };
          schedulePaginate(); scheduleSave(); reply({ ok: true });
          break;
        }
        case "setPageNumbers": {
          const ch = getChrome();
          if (m.args && m.args.enabled === false) delete ch.pageNumber;
          else {
            ch.pageNumber = {
              enabled: m.args && m.args.enabled !== false,
              format: (m.args && m.args.format) || "arabic",
              place: (m.args && m.args.place) || "footer-center",
            };
          }
          schedulePaginate(); scheduleSave(); reply({ ok: true });
          break;
        }

        // ---- page setup / zoom ----
        case "setPageSetup": {
          const s = current.pageSetup || { ...DEFAULT_PAGE_SETUP };
          const next = { ...s };
          if (m.args && m.args.size && PAGE_INCHES[m.args.size]) next.size = m.args.size;
          if (m.args && m.args.orientation) next.orientation = m.args.orientation;
          if (m.args && m.args.margins) {
            next.margins = { ...(s.margins || DEFAULT_PAGE_SETUP.margins), ...m.args.margins };
          }
          applyPageSetup(next); scheduleSave(); reply({ ok: true, pageSetup: next });
          break;
        }
        case "getPageSetup": reply({ pageSetup: current.pageSetup }); break;
        case "setZoom": {
          const z = parseFloat(m.args && m.args.zoom);
          if (!isFinite(z) || z <= 0) { reply(null, "invalid zoom"); break; }
          const sel = $("zoom");
          if (sel) sel.value = String(z);
          $("page").style.zoom = String(z);
          schedulePaginate(); reply({ ok: true, zoom: z });
          break;
        }
        case "getZoom": reply({ zoom: parseFloat(getComputedStyle($("page")).zoom) || 1 }); break;

        // ---- library / versions ----
        case "listDocuments": reply({ documents: await listDocuments() }); break;
        case "newDocument": {
          await autosaver.flush();
          const meta = await createDocument(String((m.args && m.args.title) || "Untitled"));
          await loadDocument(meta.id);
          reply({ ok: true, id: meta.id });
          break;
        }
        case "deleteDocument": {
          const id = String((m.args && m.args.id) || "");
          if (!id || id === current.id) { reply(null, "refusing to delete current doc; call newDocument first"); break; }
          await deleteDocument(id);
          reply({ ok: true });
          break;
        }
        case "listVersions": reply({ versions: await listVersions(current.id) }); break;
        case "restoreVersion": {
          const idx = parseInt(m.args && m.args.index, 10);
          if (!isFinite(idx)) { reply(null, "index required"); break; }
          const doc = await restoreVersion(current.id, idx);
          autosaver.setRev(doc.rev);
          if (doc.title != null) titleInput.value = doc.title;
          if (doc.pageSetup) applyPageSetup(doc.pageSetup);
          setEditorContent(doc.state);
          docComments = doc.comments || [];
          renderComments();
          setStatus("saved");
          reply({ ok: true, rev: doc.rev });
          break;
        }

        // ---- export ----
        case "export": {
          const fmt = String((m.args && m.args.fmt) || "html");
          const title = (titleInput.value || "document").trim();
          try {
            if (fmt === "docx") { const blob = await buildDocxFromHtml(getCleanHtml(), { title, pageSetup: current.pageSetup, comments: docComments }); saveAs(blob, title + ".docx"); }
            else if (fmt === "html") { saveAs(new Blob([exportStandaloneHtml(getCleanHtml(), title)], { type: "text/html" }), title + ".html"); }
            else if (fmt === "txt") { saveAs(new Blob([htmlToPlainText(getCleanHtml())], { type: "text/plain" }), title + ".txt"); }
            else if (fmt === "pdf") { window.print(); }
            else { reply(null, "unknown fmt: " + fmt); break; }
            setStatus("saved");
            reply({ ok: true });
          } catch (e) { reply(null, String(e.message || e)); }
          break;
        }
        case "previewPrint": window.print(); reply({ ok: true }); break;

        default: reply(null, "unknown command: " + m.cmd);
      }
    } catch (e) {
      reply(null, String(e.message || e));
    }
  });

  updateWordCount();
  schedulePaginate();
  emitToHost("ready", { id: current.id, title: current.title });
  if (mode === "edit") editor.focus();
}

main().catch((e) => {
  console.error(e);
  setStatus("error");
});
