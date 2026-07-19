// editor.js — contentEditable editing surface: toolbar, dialogs, find & replace,
// table operations, format painter, sanitization, custom context menu,
// track changes engine, comment anchors. No libraries.

const FONTS = [
  "Arial", "Calibri", "Cambria", "Courier New", "Garamond", "Georgia",
  "Helvetica", "Impact", "Tahoma", "Times New Roman", "Trebuchet MS", "Verdana",
];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];
const HIGHLIGHTS = ["#ffff00", "#00ff00", "#00ffff", "#ff00ff", "#ff9999", "#99ccff", "#cccccc", "transparent"];
const SYMBOLS = [
  "—", "–", "…", "•", "·", "§", "¶", "†", "‡", "°", "±", "×", "÷", "≈", "≠", "≤", "≥", "∞", "µ",
  "π", "Ω", "∑", "√", "∫", "α", "β", "γ", "δ", "ε", "θ", "λ", "σ", "φ", "ω",
  "→", "←", "↑", "↓", "⇒", "⇔", "✓", "✗", "★", "☆", "♦", "♣", "♠", "♥",
  "©", "®", "™", "€", "£", "¥", "¢", "«", "»", "“", "”", "‘", "’", "½", "¼", "¾", "²", "³",
];

// ---------------------------------------------------------------
// DOM helpers
// ---------------------------------------------------------------
function el(tag, attrs = {}, children = []) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === "class") e.className = v;
    else e.setAttribute(k, String(v));
  }
  for (const c of children) e.appendChild(typeof c === "string" ? document.createTextNode(c) : c);
  return e;
}
function btn(label, title, onClick, cls = "") {
  const b = el("button", { type: "button", title, class: cls });
  b.innerHTML = label;
  b.addEventListener("mousedown", (ev) => ev.preventDefault()); // keep editor selection
  b.addEventListener("click", (ev) => { ev.preventDefault(); onClick(b); });
  return b;
}
function sep() { return el("span", { class: "sep" }); }
function alignIcon(spec) {
  const rects = spec.map(([x, w], i) =>
    `<rect x="${x}" y="${(3 + i * 3.3).toFixed(2)}" width="${w}" height="1.6" rx=".8"/>`).join("");
  return `<svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor" aria-hidden="true">${rects}</svg>`;
}
const ALIGN_ICONS = {
  left:    alignIcon([[2, 12], [2, 9], [2, 11], [2, 7]]),
  center:  alignIcon([[2, 12], [3.5, 9], [2.5, 11], [4.5, 7]]),
  right:   alignIcon([[2, 12], [5, 9], [3, 11], [7, 7]]),
  justify: alignIcon([[2, 12], [2, 12], [2, 12], [2, 7]]),
};
// Curved undo/redo arrows.
const UNDO_ICON = `<svg width="17" height="17" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3.5 8h6a3 3 0 0 1 0 6H6"/><path d="M6 5 3 8l3 3"/></svg>`;
const REDO_ICON = `<svg width="17" height="17" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12.5 8h-6a3 3 0 0 0 0 6H10"/><path d="M10 5l3 3-3 3"/></svg>`;

// ---------------------------------------------------------------
// Tooltips — styled hover popovers driven by the title attribute.
// The native title is moved to data-tip on first hover so the browser's
// own (slow, unstyled) tooltip never double-shows.
// ---------------------------------------------------------------
let tipEl = null, tipTimer = null;
function hideTip() {
  clearTimeout(tipTimer);
  if (tipEl) tipEl.classList.add("hidden");
}
function showTip(target) {
  const text = target.getAttribute("data-tip");
  if (!text || !target.isConnected) return;
  if (!tipEl) { tipEl = el("div", { class: "tooltip hidden", role: "tooltip" }); document.body.appendChild(tipEl); }
  tipEl.textContent = text;
  tipEl.classList.remove("hidden");
  const r = target.getBoundingClientRect();
  const tr = tipEl.getBoundingClientRect();
  let left = r.left + r.width / 2 - tr.width / 2;
  left = Math.max(6, Math.min(left, window.innerWidth - tr.width - 6));
  let top = r.bottom + 6;
  if (top + tr.height > window.innerHeight - 6) top = r.top - tr.height - 6;
  tipEl.style.left = left + "px";
  tipEl.style.top = top + "px";
}
export function initTooltips() {
  const SCOPE = "#toolbar,#topbar,.table-toolbar,.img-toolbar,.find-panel";
  document.addEventListener("mouseover", (e) => {
    const t = e.target.closest && e.target.closest("[title],[data-tip]");
    if (!t || !t.closest(SCOPE)) return;
    if (t.hasAttribute("title")) { t.setAttribute("data-tip", t.getAttribute("title")); t.removeAttribute("title"); }
    clearTimeout(tipTimer);
    tipTimer = setTimeout(() => showTip(t), 350);
  });
  document.addEventListener("mouseout", (e) => {
    if (e.target.closest && e.target.closest("[data-tip]")) hideTip();
  });
  document.addEventListener("mousedown", hideTip, true);
  window.addEventListener("blur", hideTip);
  document.addEventListener("scroll", hideTip, true);
}
function exec(cmd, val) { document.execCommand(cmd, false, val); }
function escText(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
function fireInput(editor) {
  editor.dispatchEvent(new Event("input", { bubbles: true }));
}
function unwrapNode(node) {
  const parent = node.parentNode;
  if (!parent) return;
  while (node.firstChild) parent.insertBefore(node.firstChild, node);
  parent.removeChild(node);
  parent.normalize();
}

// ---------------------------------------------------------------
// Selection utilities
// ---------------------------------------------------------------
let savedRange = null;
export function saveSelection(editor) {
  const sel = window.getSelection();
  if (sel.rangeCount && editor.contains(sel.anchorNode)) savedRange = sel.getRangeAt(0).cloneRange();
}
export function restoreSelection(editor) {
  if (!savedRange) return false;
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(savedRange);
  return true;
}
export function insertHtmlAtCaret(editor, html) {
  editor.focus();
  restoreSelection(editor);
  exec("insertHTML", html);
  saveSelection(editor);
}
export function insertTextAtCaret(editor, text) {
  editor.focus();
  restoreSelection(editor);
  exec("insertText", text);
  saveSelection(editor);
}

// Wrap the current selection in a span with the given inline style.
function applySpanStyle(editor, styleProps) {
  editor.focus();
  const sel = window.getSelection();
  // if the live selection was lost/collapsed (e.g. focus moved to a toolbar
  // control), fall back to the last saved non-collapsed range.
  if (!sel.rangeCount || sel.isCollapsed) {
    if (savedRange && !savedRange.collapsed) {
      sel.removeAllRanges();
      sel.addRange(savedRange);
    } else return;
  }
  if (!sel.rangeCount || sel.isCollapsed) return;
  const range = sel.getRangeAt(0);
  const span = document.createElement("span");
  for (const [k, v] of Object.entries(styleProps)) span.style[k] = v;
  try {
    range.surroundContents(span);
  } catch {
    const frag = range.extractContents();
    span.appendChild(frag);
    range.insertNode(span);
  }
  sel.removeAllRanges();
  const r = document.createRange();
  r.selectNodeContents(span);
  sel.addRange(r);
  saveSelection(editor);
  fireInput(editor);
}

function closestBlock(node, editor) {
  let n = node && node.nodeType === 3 ? node.parentElement : node;
  while (n && n !== editor) {
    if (/^(P|DIV|H[1-6]|LI|BLOCKQUOTE|TD|TH|PRE)$/.test(n.tagName)) return n;
    n = n.parentElement;
  }
  return null;
}
function selectedBlocks(editor) {
  const sel = window.getSelection();
  if (!sel.rangeCount) return [];
  const range = sel.getRangeAt(0);
  const blocks = [];
  const walker = document.createTreeWalker(editor, NodeFilter.SHOW_ELEMENT);
  let n;
  while ((n = walker.nextNode())) {
    if (/^(P|DIV|H[1-6]|LI|BLOCKQUOTE)$/.test(n.tagName) && range.intersectsNode(n)) blocks.push(n);
  }
  if (!blocks.length) {
    const b = closestBlock(sel.anchorNode, editor);
    if (b) blocks.push(b);
  }
  return blocks;
}

// ---------------------------------------------------------------
// HTML sanitizer (paste + loaded content)
// ---------------------------------------------------------------
const ALLOWED = {
  p: [], div: [], br: [], hr: [], span: ["data-cid"], font: ["face", "size", "color"],
  b: [], strong: [], i: [], em: [], u: [], s: [], strike: [],
  ins: ["data-author", "data-ts"], del: ["data-author", "data-ts"],
  sub: [], sup: [], mark: [], code: [], kbd: [], samp: [], tt: [], pre: [],
  h1: [], h2: [], h3: [], h4: [], h5: [], h6: [], blockquote: [],
  ul: [], ol: ["start"], li: [],
  table: [], thead: [], tbody: [], tfoot: [], tr: [],
  td: ["colspan", "rowspan"], th: ["colspan", "rowspan"],
  a: ["href"], img: ["src", "width", "height", "alt"],
  figure: [], figcaption: [], section: [], article: [],
};
const ALLOWED_STYLES = new Set([
  "color", "background-color", "font-size", "font-family", "font-weight", "font-style",
  "text-decoration", "text-decoration-line", "text-align", "text-indent",
  "line-height", "vertical-align", "width", "height",
  // image editing + shapes
  "border", "border-width", "border-style", "border-color", "border-radius",
  "float", "margin", "margin-top", "margin-right", "margin-bottom", "margin-left",
  "display", "clip-path", "background", "box-shadow", "padding", "transform",
  // free-floating shape objects
  "position", "left", "top", "right", "bottom", "z-index",
]);
const KEEP_CLASSES = ["page-break", "pb", "comment-ref", "resolved", "tc-ins", "tc-del", "wordart", "shape"];
// classes matching these prefixes are also preserved (WordArt / shape variants)
const KEEP_CLASS_RE = /^(wa-\d+|shape-[a-z]+)$/;
const DROP_TAGS = new Set(["script", "style", "meta", "link", "head", "title", "iframe", "object", "embed", "applet", "noscript", "svg", "math", "template", "form", "input", "button", "select", "textarea", "audio", "video"]);

function sanitizeStyle(el, out) {
  const style = el.getAttribute && el.getAttribute("style");
  if (!style) return;
  const parts = [];
  for (const decl of style.split(";")) {
    const i = decl.indexOf(":");
    if (i < 0) continue;
    const k = decl.slice(0, i).trim().toLowerCase();
    const v = decl.slice(i + 1).trim();
    if (!ALLOWED_STYLES.has(k)) continue;
    if (/url\s*\(|expression|javascript:/i.test(v)) continue;
    // never let pasted content pin itself to the viewport
    if (k === "position" && !/^(absolute|relative|static)$/i.test(v)) continue;
    parts.push(`${k}:${v}`);
  }
  if (parts.length) out.setAttribute("style", parts.join(";"));
}

function sanitizeNode(node, doc) {
  if (node.nodeType === 3) return doc.createTextNode(node.textContent);
  if (node.nodeType !== 1) return null;
  const tag = node.tagName.toLowerCase();
  if (DROP_TAGS.has(tag)) return null;
  const allowedAttrs = ALLOWED[tag];
  if (allowedAttrs === undefined) {
    // unknown tag: unwrap, keep children
    const frag = doc.createDocumentFragment();
    for (const c of node.childNodes) {
      const s = sanitizeNode(c, doc);
      if (s) frag.appendChild(s);
    }
    return frag;
  }
  const out = doc.createElement(tag);
  for (const a of allowedAttrs) {
    const v = node.getAttribute(a);
    if (v == null) continue;
    if (a === "href" && !/^(https?:|mailto:|#)/i.test(v.trim())) continue;
    if (a === "src" && !/^(data:image\/|https?:|blob:)/i.test(v.trim())) continue;
    out.setAttribute(a, v);
  }
  sanitizeStyle(node, out);
  // preserve our structural classes only
  if (node.classList) {
    const keep = [...node.classList].filter((c) => KEEP_CLASSES.includes(c) || KEEP_CLASS_RE.test(c));
    if (keep.length) out.className = keep.join(" ");
  }
  for (const c of node.childNodes) {
    const s = sanitizeNode(c, doc);
    if (s) out.appendChild(s);
  }
  return out;
}

export function sanitizeHtml(html) {
  const doc = new DOMParser().parseFromString(`<div>${html || ""}</div>`, "text/html");
  const root = doc.body.firstChild;
  const container = document.createElement("div");
  for (const c of root.childNodes) {
    const s = sanitizeNode(c, document);
    if (s) container.appendChild(s);
  }
  return container.innerHTML;
}

// ---------------------------------------------------------------
// Dialogs
// ---------------------------------------------------------------
export function openDialog(title, bodyEl, actions) {
  closeDialog();
  const overlay = el("div", { class: "dlg-overlay", id: "dlg-overlay" });
  const box = el("div", { class: "dlg" });
  box.appendChild(el("div", { class: "dlg-title" }, [title]));
  const body = el("div", { class: "dlg-body" });
  body.appendChild(bodyEl);
  box.appendChild(body);
  const bar = el("div", { class: "dlg-actions" });
  for (const a of actions) {
    const b = el("button", { type: "button", class: a.primary ? "primary" : "" }, [a.label]);
    b.addEventListener("click", () => { if (a.onClick && a.onClick() === false) return; closeDialog(); });
    bar.appendChild(b);
  }
  box.appendChild(bar);
  overlay.appendChild(box);
  overlay.addEventListener("mousedown", (e) => { if (e.target === overlay) closeDialog(); });
  document.body.appendChild(overlay);
  const firstInput = box.querySelector("input,select,textarea");
  if (firstInput) firstInput.focus();
  overlay.addEventListener("keydown", (e) => {
    if (e.key === "Escape") closeDialog();
    if (e.key === "Enter" && e.target.tagName === "INPUT") {
      const primary = bar.querySelector(".primary");
      if (primary) { e.preventDefault(); primary.click(); }
    }
  });
  return overlay;
}
export function closeDialog() {
  const d = document.getElementById("dlg-overlay");
  if (d) d.remove();
}
function field(label, input) {
  return el("label", { class: "dlg-field" }, [el("span", {}, [label]), input]);
}

export function openLinkDialog(editor) {
  saveSelection(editor);
  const sel = window.getSelection();
  const selText = sel.rangeCount ? sel.toString() : "";
  const textIn = el("input", { type: "text", value: selText });
  const urlIn = el("input", { type: "text", placeholder: "https://…" });
  const body = el("div", {}, [field("Text", textIn), field("URL", urlIn)]);
  openDialog("Insert link", body, [
    { label: "Cancel" },
    {
      label: "Insert", primary: true,
      onClick: () => {
        let url = urlIn.value.trim();
        if (!url) return false;
        if (!/^(https?:|mailto:)/i.test(url)) url = "https://" + url;
        const text = textIn.value.trim() || url;
        const a = document.createElement("a");
        a.href = url;
        a.textContent = text;
        insertHtmlAtCaret(editor, a.outerHTML);
      },
    },
  ]);
}

export function openTableDialog(editor) {
  saveSelection(editor);
  const rowsIn = el("input", { type: "number", value: "3", min: "1", max: "50" });
  const colsIn = el("input", { type: "number", value: "3", min: "1", max: "20" });
  const headerIn = el("input", { type: "checkbox", checked: "" });
  const body = el("div", {}, [
    field("Rows", rowsIn), field("Columns", colsIn),
    el("label", { class: "dlg-field dlg-check" }, [headerIn, el("span", {}, ["Header row"])]),
  ]);
  openDialog("Insert table", body, [
    { label: "Cancel" },
    {
      label: "Insert", primary: true,
      onClick: () => {
        const rows = Math.min(50, Math.max(1, parseInt(rowsIn.value, 10) || 3));
        const cols = Math.min(20, Math.max(1, parseInt(colsIn.value, 10) || 3));
        let html = "<table><tbody>";
        for (let r = 0; r < rows; r++) {
          html += "<tr>";
          const tag = headerIn.checked && r === 0 ? "th" : "td";
          for (let c = 0; c < cols; c++) html += `<${tag}><br></${tag}>`;
          html += "</tr>";
        }
        html += "</tbody></table><p><br></p>";
        insertHtmlAtCaret(editor, html);
      },
    },
  ]);
}

export function openSymbolDialog(editor) {
  saveSelection(editor);
  const grid = el("div", { class: "symbol-grid" });
  for (const s of SYMBOLS) {
    const b = el("button", { type: "button", class: "symbol" }, [s]);
    b.addEventListener("click", () => { insertTextAtCaret(editor, s); closeDialog(); });
    grid.appendChild(b);
  }
  openDialog("Insert symbol", grid, [{ label: "Close" }]);
}

export function openPageSetupDialog(current, onApply) {
  const cur = current || {};
  const sizeSel = el("select", {});
  for (const s of ["Letter", "A4", "Legal", "A3"]) {
    const o = el("option", { value: s }, [s]);
    if (cur.size === s) o.selected = true;
    sizeSel.appendChild(o);
  }
  const orientSel = el("select", {});
  for (const o of ["portrait", "landscape"]) {
    const opt = el("option", { value: o }, [o[0].toUpperCase() + o.slice(1)]);
    if (cur.orientation === o) opt.selected = true;
    orientSel.appendChild(opt);
  }
  const m = cur.margins || { top: 1, right: 1, bottom: 1, left: 1 };
  const mIn = {};
  for (const side of ["top", "bottom", "left", "right"]) {
    mIn[side] = el("input", { type: "number", step: "0.1", min: "0", max: "3", value: String(m[side]) });
  }
  const body = el("div", {}, [
    field("Paper size", sizeSel),
    field("Orientation", orientSel),
    el("div", { class: "dlg-grid2" }, [
      field("Top margin (in)", mIn.top), field("Bottom (in)", mIn.bottom),
      field("Left (in)", mIn.left), field("Right (in)", mIn.right),
    ]),
  ]);
  openDialog("Page setup", body, [
    { label: "Cancel" },
    {
      label: "Apply", primary: true,
      onClick: () => {
        const margins = {};
        for (const side of ["top", "bottom", "left", "right"]) {
          margins[side] = Math.max(0, Math.min(3, parseFloat(mIn[side].value) || 1));
        }
        onApply({ size: sizeSel.value, orientation: orientSel.value, margins });
      },
    },
  ]);
}

export function insertImage(editor) {
  saveSelection(editor);
  const input = el("input", { type: "file", accept: "image/*" });
  input.addEventListener("change", () => {
    const file = input.files && input.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const probe = new Image();
      probe.onload = () => {
        let w = probe.naturalWidth, h = probe.naturalHeight;
        const MAXW = 624;
        if (w > MAXW) { h = Math.round(h * (MAXW / w)); w = MAXW; }
        insertHtmlAtCaret(editor, `<img src="${reader.result}" width="${w}" height="${h}" alt="">`);
      };
      probe.src = String(reader.result);
    };
    reader.readAsDataURL(file);
  });
  input.click();
}

// ---------------------------------------------------------------
// Insert menu content: page break / blank page / WordArt / shapes
// ---------------------------------------------------------------
const PAGEBREAK_HTML = '<p class="page-break"><br></p><p><br></p>';
const BLANKPAGE_HTML = '<p class="page-break"><br></p><p><br></p><p class="page-break"><br></p><p><br></p>';
export function insertPageBreak(editor) { insertHtmlAtCaret(editor, PAGEBREAK_HTML); }
export function insertBlankPage(editor) { insertHtmlAtCaret(editor, BLANKPAGE_HTML); }

// WordArt: decorative text styled entirely by CSS classes (round-trips through
// the sanitizer as span + wa-N class; keeps its text in .docx export).
export function openWordArtDialog(editor) {
  saveSelection(editor);
  const initial = (window.getSelection().toString() || "").trim() || "WordArt";
  const textIn = el("input", { type: "text", value: initial });
  const styles = el("div", { class: "wordart-styles" });
  let chosen = 1;
  for (let i = 1; i <= 8; i++) {
    const b = el("button", { type: "button", class: "wa-swatch wordart wa-" + i + (i === 1 ? " sel" : "") });
    b.textContent = "Aa";
    b.addEventListener("click", () => {
      chosen = i;
      styles.querySelectorAll("button").forEach((x) => x.classList.remove("sel"));
      b.classList.add("sel");
    });
    styles.appendChild(b);
  }
  const body = el("div", {}, [
    field("Text", textIn),
    el("div", { class: "dlg-field dlg-field-col" }, [el("span", {}, ["Style"]), styles]),
  ]);
  openDialog("Insert WordArt", body, [
    { label: "Cancel" },
    {
      label: "Insert", primary: true,
      onClick: () => {
        const text = textIn.value.trim();
        if (!text) return false;
        insertHtmlAtCaret(editor, `<span class="wordart wa-${chosen}">${escText(text)}</span>&nbsp;`);
      },
    },
  ]);
}

// Shapes: self-contained spans styled with allowed inline CSS so they survive
// the sanitizer and re-render on reload. Inserted as absolutely-positioned
// objects that the user can drag anywhere on the page.
const SHAPES = [
  { type: "rect",     name: "Rectangle",         style: "width:120px;height:72px;background:#4a90d9;border-radius:2px" },
  { type: "rounded",  name: "Rounded rectangle", style: "width:120px;height:72px;background:#4a90d9;border-radius:16px" },
  { type: "ellipse",  name: "Ellipse",           style: "width:96px;height:96px;background:#50b477;border-radius:50%" },
  { type: "triangle", name: "Triangle",          style: "width:96px;height:84px;background:#e8823a;clip-path:polygon(50% 0,100% 100%,0 100%)" },
  { type: "star",     name: "Star",              style: "width:96px;height:92px;background:#f0c419;clip-path:polygon(50% 0,61% 35%,98% 35%,68% 57%,79% 91%,50% 70%,21% 91%,32% 57%,2% 35%,39% 35%)" },
  { type: "arrow",    name: "Arrow",             style: "width:120px;height:64px;background:#7a55c8;clip-path:polygon(0 30%,60% 30%,60% 8%,100% 50%,60% 92%,60% 70%,0 70%)" },
  { type: "line",     name: "Line",              style: "width:140px;height:4px;background:#444" },
  { type: "diamond",  name: "Diamond",           style: "width:92px;height:92px;background:#d9534f;clip-path:polygon(50% 0,100% 50%,50% 100%,0 50%)" },
  { type: "pentagon", name: "Pentagon",          style: "width:96px;height:92px;background:#20a4a4;clip-path:polygon(50% 0,100% 38%,82% 100%,18% 100%,0 38%)" },
  { type: "hexagon",  name: "Hexagon",           style: "width:104px;height:92px;background:#e0567b;clip-path:polygon(25% 0,75% 0,100% 50%,75% 100%,25% 100%,0 50%)" },
  { type: "chevron",  name: "Chevron",           style: "width:120px;height:64px;background:#5b8c2a;clip-path:polygon(0 0,70% 0,100% 50%,70% 100%,0 100%,30% 50%)" },
  { type: "cross",    name: "Cross",             style: "width:92px;height:92px;background:#8a6d3b;clip-path:polygon(35% 0,65% 0,65% 35%,100% 35%,100% 65%,65% 65%,65% 100%,35% 100%,35% 65%,0 65%,0 35%,35% 35%)" },
];
// A shape as a free-floating object (position:absolute inside the editor).
// `fill` overrides the shape's default background color.
function makeShapeEl(s, left, top, fill) {
  const span = document.createElement("span");
  span.className = "shape shape-" + s.type;
  let style = s.style;
  if (fill) style = style.replace(/background:[^;]+/, "background:" + fill);
  span.setAttribute("style", `position:absolute;left:${left}px;top:${top}px;${style}`);
  return span;
}

// Insert-shape picker — works like the symbol picker (a popup grid), with a
// color selector so the shape is inserted in the chosen fill color.
export function openShapeDialog(editor) {
  saveSelection(editor);
  const colorIn = el("input", { type: "color", value: "#4a90d9", title: "Shape color" });
  const colorRow = el("div", { class: "shape-color-row" }, [el("span", {}, ["Color"]), colorIn]);
  // live-tint the preview swatches to the chosen color
  const tint = () => grid.querySelectorAll(".shape-preview").forEach((p) => { p.style.background = colorIn.value; });
  colorIn.addEventListener("input", tint);
  const grid = el("div", { class: "shape-grid" });
  for (const s of SHAPES) {
    const b = el("button", { type: "button", class: "shape-pick", title: s.name });
    b.innerHTML = `<span class="shape-preview shape-${s.type}"></span>`;
    b.addEventListener("click", () => {
      // stagger inserts so several don't stack exactly
      const n = editor.querySelectorAll(".shape").length;
      const off = (n % 6) * 22;
      const span = makeShapeEl(s, 90 + off, 90 + off, colorIn.value);
      editor.appendChild(span);
      fireInput(editor);
      closeDialog();
      if (typeof editor._selectShape === "function") editor._selectShape(span);
    });
    grid.appendChild(b);
  }
  tint();
  const body = el("div", {}, [colorRow, grid]);
  openDialog("Insert shape", body, [{ label: "Close" }]);
}

// ---------------------------------------------------------------
// Find & replace
// ---------------------------------------------------------------
export function createFindPanel(editor, host) {
  const panel = el("div", { class: "find-panel hidden", id: "find-panel" });
  const findIn = el("input", { type: "text", placeholder: "Find…" });
  const replIn = el("input", { type: "text", placeholder: "Replace with…" });
  const caseCk = el("input", { type: "checkbox", title: "Match case" });
  const regexCk = el("input", { type: "checkbox", title: "Regular expression" });
  const count = el("span", { class: "find-count" }, [""]);
  let hits = [];
  let cur = -1;

  function clearHighlights() {
    for (const m of editor.querySelectorAll("mark.find-hit")) {
      const parent = m.parentNode;
      while (m.firstChild) parent.insertBefore(m.firstChild, m);
      parent.removeChild(m);
      parent.normalize();
    }
    hits = []; cur = -1;
    count.textContent = "";
  }

  function blockTextSegments() {
    // group text nodes by nearest block so matches never cross block boundaries
    const segs = new Map();
    const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
    let n;
    while ((n = walker.nextNode())) {
      const block = closestBlock(n, editor) || editor;
      if (!segs.has(block)) segs.set(block, { nodes: [], text: "" });
      const seg = segs.get(block);
      seg.nodes.push({ node: n, start: seg.text.length });
      seg.text += n.textContent;
    }
    return segs;
  }

  function buildRegex() {
    const q = findIn.value;
    if (!q) return null;
    const flags = caseCk.checked ? "g" : "gi";
    try {
      return regexCk.checked ? new RegExp(q, flags) : new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), flags);
    } catch { return null; }
  }

  function search() {
    clearHighlights();
    const re = buildRegex();
    if (!re) return;
    const segs = blockTextSegments();
    const ranges = [];
    for (const { nodes, text } of segs.values()) {
      let m;
      while ((m = re.exec(text)) !== null) {
        if (m[0] === "") { re.lastIndex++; continue; }
        const start = m.index, end = m.index + m[0].length;
        const locate = (off, isEnd) => {
          for (let i = nodes.length - 1; i >= 0; i--) {
            const { node, start: s } = nodes[i];
            if (off > s || (off === s && (!isEnd || i === 0))) return { node, off: off - s };
            if (off === s && isEnd) return { node: nodes[i - 1].node, off: nodes[i - 1] ? off - nodes[i - 1].start : 0 };
          }
          return { node: nodes[0].node, off: 0 };
        };
        const a = locate(start, false), b = locate(end, true);
        const range = document.createRange();
        try {
          range.setStart(a.node, Math.min(a.off, a.node.textContent.length));
          range.setEnd(b.node, Math.min(b.off, b.node.textContent.length));
          ranges.push(range);
        } catch {}
      }
    }
    // wrap from last to first so offsets stay valid
    for (let i = ranges.length - 1; i >= 0; i--) {
      const mark = document.createElement("mark");
      mark.className = "find-hit";
      try {
        ranges[i].surroundContents(mark);
      } catch {
        const frag = ranges[i].extractContents();
        mark.appendChild(frag);
        ranges[i].insertNode(mark);
      }
    }
    hits = [...editor.querySelectorAll("mark.find-hit")];
    if (hits.length) { cur = 0; focusHit(); }
    count.textContent = hits.length ? `${cur + 1}/${hits.length}` : "0 results";
  }

  function focusHit() {
    hits.forEach((h, i) => h.classList.toggle("current", i === cur));
    if (hits[cur]) {
      hits[cur].scrollIntoView({ block: "center", behavior: "smooth" });
      count.textContent = `${cur + 1}/${hits.length}`;
    }
  }
  function move(dir) {
    if (!hits.length) { search(); return; }
    cur = (cur + dir + hits.length) % hits.length;
    focusHit();
  }
  function replaceCurrent() {
    if (!hits.length) { search(); if (!hits.length) return; }
    const m = hits[cur];
    m.replaceWith(document.createTextNode(replIn.value));
    fireInput(editor);
    search();
  }
  function replaceAll() {
    search();
    if (!hits.length) return;
    for (const m of hits) m.replaceWith(document.createTextNode(replIn.value));
    fireInput(editor);
    clearHighlights();
    count.textContent = "replaced";
  }

  findIn.addEventListener("input", () => search());
  panel.append(
    findIn,
    btn("↑", "Previous", () => move(-1)),
    btn("↓", "Next", () => move(1)),
    el("label", { class: "find-opt", title: "Match case" }, [caseCk, "Aa"]),
    el("label", { class: "find-opt", title: "Regular expression" }, [regexCk, ".*"]),
    count,
    replIn,
    btn("Replace", "Replace current", replaceCurrent),
    btn("All", "Replace all", replaceAll),
    btn("✕", "Close", () => api.close()),
  );
  host.appendChild(panel);

  const api = {
    toggle() {
      panel.classList.toggle("hidden");
      if (!panel.classList.contains("hidden")) {
        const sel = window.getSelection().toString();
        if (sel && sel.length < 100) findIn.value = sel;
        findIn.focus(); findIn.select();
        if (findIn.value) search();
      } else clearHighlights();
    },
    close() { panel.classList.add("hidden"); clearHighlights(); },
    isOpen() { return !panel.classList.contains("hidden"); },
    find(q, opts = {}) { findIn.value = q; caseCk.checked = !!opts.matchCase; regexCk.checked = !!opts.regex; panel.classList.remove("hidden"); search(); return hits.length; },
    replaceAll(q, r, opts = {}) { findIn.value = q; replIn.value = r; caseCk.checked = !!opts.matchCase; regexCk.checked = !!opts.regex; search(); const n = hits.length; if (n) replaceAll(); return n; },
    clear: clearHighlights,
  };
  return api;
}

// ---------------------------------------------------------------
// Table operations (shared by table toolbar and context menu)
// ---------------------------------------------------------------
function cellIndexInRow(cell) {
  let i = 0;
  for (const c of cell.parentElement.children) {
    if (c === cell) return i;
    i += Math.max(1, parseInt(c.getAttribute("colspan"), 10) || 1);
  }
  return i;
}
function cellAtIndex(row, idx) {
  let i = 0;
  for (const c of row.children) {
    const span = Math.max(1, parseInt(c.getAttribute("colspan"), 10) || 1);
    if (idx >= i && idx < i + span) return c;
    i += span;
  }
  return row.lastElementChild;
}
function tableInsertRow(cell, below) {
  const row = cell.parentElement;
  const nCells = [...row.children].reduce((n, c) => n + Math.max(1, parseInt(c.getAttribute("colspan"), 10) || 1), 0);
  const tr = document.createElement("tr");
  for (let i = 0; i < nCells; i++) {
    const td = document.createElement("td");
    td.innerHTML = "<br>";
    tr.appendChild(td);
  }
  row.parentElement.insertBefore(tr, below ? row.nextSibling : row);
}
function tableInsertCol(cell, right) {
  const idx = cellIndexInRow(cell) + (right ? 1 : 0);
  const tbl = cell.closest("table");
  for (const row of tbl.querySelectorAll("tr")) {
    const ref = cellAtIndex(row, idx);
    const td = document.createElement("td");
    td.innerHTML = "<br>";
    if (ref && cellIndexInRow(ref) >= idx) row.insertBefore(td, ref);
    else row.appendChild(td);
  }
}
function tableDeleteRow(cell) {
  const tbl = cell.closest("table");
  cell.parentElement.remove();
  if (tbl && !tbl.querySelector("tr")) tbl.remove();
}
function tableDeleteCol(cell) {
  const idx = cellIndexInRow(cell);
  const tbl = cell.closest("table");
  for (const row of tbl.querySelectorAll("tr")) {
    const c = cellAtIndex(row, idx);
    if (c) {
      const span = Math.max(1, parseInt(c.getAttribute("colspan"), 10) || 1);
      if (span > 1) c.setAttribute("colspan", span - 1);
      else c.remove();
    }
  }
  if (!tbl.querySelector("td,th")) tbl.remove();
}
function tableDelete(cell) {
  const tbl = cell.closest("table");
  if (tbl) tbl.remove();
}
function tableMergeRight(cell) {
  const next = cell.nextElementSibling;
  if (!next) return;
  const span = Math.max(1, parseInt(cell.getAttribute("colspan"), 10) || 1);
  const nspan = Math.max(1, parseInt(next.getAttribute("colspan"), 10) || 1);
  cell.setAttribute("colspan", span + nspan);
  if (next.textContent.trim()) cell.innerHTML += " " + next.innerHTML;
  next.remove();
}
function tableMergeDown(cell) {
  const row = cell.parentElement;
  const nextRow = row.nextElementSibling;
  if (!nextRow) return;
  const idx = cellIndexInRow(cell);
  const below = cellAtIndex(nextRow, idx);
  if (!below || cellIndexInRow(below) !== idx) return;
  const span = Math.max(1, parseInt(cell.getAttribute("rowspan"), 10) || 1);
  const bspan = Math.max(1, parseInt(below.getAttribute("rowspan"), 10) || 1);
  cell.setAttribute("rowspan", span + bspan);
  if (below.textContent.trim()) cell.innerHTML += " " + below.innerHTML;
  below.remove();
}
function tableSplitCell(cell) {
  const colspan = Math.max(1, parseInt(cell.getAttribute("colspan"), 10) || 1);
  const rowspan = Math.max(1, parseInt(cell.getAttribute("rowspan"), 10) || 1);
  cell.removeAttribute("colspan");
  for (let i = 1; i < colspan; i++) {
    const td = document.createElement("td");
    td.innerHTML = "<br>";
    cell.parentElement.insertBefore(td, cell.nextSibling);
  }
  if (rowspan > 1) {
    cell.removeAttribute("rowspan");
    const idx = cellIndexInRow(cell);
    let row = cell.parentElement;
    for (let r = 1; r < rowspan; r++) {
      row = row.nextElementSibling;
      if (!row) break;
      const ref = cellAtIndex(row, idx);
      for (let i = 0; i < colspan; i++) {
        const td = document.createElement("td");
        td.innerHTML = "<br>";
        if (ref && cellIndexInRow(ref) >= idx) row.insertBefore(td, ref);
        else row.appendChild(td);
      }
    }
  }
}
function currentCell(editor) {
  const sel = window.getSelection();
  if (!sel.rangeCount) return null;
  let n = sel.anchorNode;
  if (n && n.nodeType === 3) n = n.parentElement;
  while (n && n !== editor) {
    if (n.tagName === "TD" || n.tagName === "TH") return n;
    n = n.parentElement;
  }
  return null;
}

export function attachTableToolbar(editor, wrap) {
  const bar = el("div", { class: "table-toolbar hidden", id: "table-toolbar" });
  const withCell = (fn) => () => { const c = currentCell(editor); if (c) { fn(c); fireInput(editor); } };

  bar.append(
    el("span", { class: "tt-label" }, ["Table:"]),
    btn("+Row ↑", "Insert row above", withCell((c) => tableInsertRow(c, false))),
    btn("+Row ↓", "Insert row below", withCell((c) => tableInsertRow(c, true))),
    btn("+Col ←", "Insert column left", withCell((c) => tableInsertCol(c, false))),
    btn("+Col →", "Insert column right", withCell((c) => tableInsertCol(c, true))),
    sep(),
    btn("Merge →", "Merge with cell to the right", withCell(tableMergeRight)),
    btn("Merge ↓", "Merge with cell below", withCell(tableMergeDown)),
    btn("Split", "Unmerge cell", withCell(tableSplitCell)),
    sep(),
    btn("Shade", "Cell background", withCell((c) => {
      const input = el("input", { type: "color", value: "#f5f5dc" });
      input.addEventListener("input", () => { c.style.backgroundColor = input.value; fireInput(editor); });
      input.click();
    })),
    btn("Del row", "Delete row", withCell(tableDeleteRow)),
    btn("Del col", "Delete column", withCell(tableDeleteCol)),
    btn("Del table", "Delete table", withCell(tableDelete)),
  );
  wrap.appendChild(bar);

  document.addEventListener("selectionchange", () => {
    const sel = window.getSelection();
    const inEditor = sel.rangeCount && editor.contains(sel.anchorNode);
    const cell = inEditor ? currentCell(editor) : null;
    bar.classList.toggle("hidden", !cell);
  });
  return bar;
}

// ---------------------------------------------------------------
// Custom context menu
// ---------------------------------------------------------------
export function attachContextMenu(editor, actions = {}) {
  let menu = null;
  const close = () => { if (menu) { menu.remove(); menu = null; } };
  document.addEventListener("mousedown", (e) => { if (menu && !menu.contains(e.target)) close(); });
  document.addEventListener("scroll", close, true);
  window.addEventListener("blur", close);
  window.addEventListener("resize", close);
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") close(); });

  function item(label, action, opts = {}) {
    return { label, action, ...opts };
  }

  editor.addEventListener("contextmenu", (e) => {
    if (editor.contentEditable !== "true") return; // view mode: browser default
    e.preventDefault();
    close();
    const sel = window.getSelection();
    saveSelection(editor);
    const target = e.target.nodeType === 3 ? e.target.parentElement : e.target;
    const hasSel = sel && !sel.isCollapsed;
    const items = [];

    items.push(item("✂️ Cut", () => exec("cut"), { disabled: !hasSel }));
    items.push(item("📋 Copy", () => exec("copy"), { disabled: !hasSel }));
    items.push(item("📥 Paste", async () => {
      editor.focus();
      restoreSelection(editor);
      try {
        const text = await navigator.clipboard.readText();
        if (text) exec("insertText", text);
      } catch {
        alert("Your browser blocked script paste — press Ctrl/⌘+V instead.");
      }
    }));
    items.push("-");
    items.push(item("💬 Add comment", () => actions.addComment && actions.addComment(), { disabled: !actions.addComment }));

    const cref = target.closest && target.closest("span.comment-ref");
    if (cref && actions.openComment) {
      items.push(item("👁 View comment", () => actions.openComment(cref.getAttribute("data-cid"))));
    }

    const tracked = target.closest && target.closest("ins.tc-ins, del.tc-del");
    if (tracked && actions.acceptChange) {
      items.push("-");
      items.push(item("✓ Accept change", () => actions.acceptChange(tracked)));
      items.push(item("✗ Reject change", () => actions.rejectChange(tracked)));
    }

    items.push("-");
    const a = target.closest && target.closest("a[href]");
    if (a) {
      items.push(item("🔗 Open link", () => window.open(a.href, "_blank", "noopener")));
      items.push(item("⛓ Remove link", () => { unwrapNode(a); fireInput(editor); }));
    } else {
      items.push(item("🔗 Insert link…", () => openLinkDialog(editor)));
    }

    const cell = target.closest && target.closest("td,th");
    if (cell && editor.contains(cell)) {
      items.push("-");
      items.push(item("Insert row below", () => { tableInsertRow(cell, true); fireInput(editor); }));
      items.push(item("Insert column right", () => { tableInsertCol(cell, true); fireInput(editor); }));
      items.push(item("Delete row", () => { tableDeleteRow(cell); fireInput(editor); }));
      items.push(item("Delete column", () => { tableDeleteCol(cell); fireInput(editor); }));
      items.push(item("Delete table", () => { tableDelete(cell); fireInput(editor); }));
    }

    items.push("-");
    items.push(item("Select all", () => { editor.focus(); exec("selectAll"); }));

    // build menu
    menu = el("div", { class: "ctx-menu" });
    for (const it of items) {
      if (it === "-") {
        if (menu.lastChild && !menu.lastChild.classList.contains("ctx-sep")) {
          menu.appendChild(el("div", { class: "ctx-sep" }));
        }
        continue;
      }
      const b = el("button", { type: "button" }, []);
      b.textContent = it.label;
      if (it.disabled) b.disabled = true;
      b.addEventListener("mousedown", (ev) => ev.preventDefault());
      b.addEventListener("click", () => { close(); it.action && it.action(); });
      menu.appendChild(b);
    }
    document.body.appendChild(menu);
    // position within viewport
    const mw = menu.offsetWidth, mh = menu.offsetHeight;
    let x = e.clientX, y = e.clientY;
    if (x + mw > window.innerWidth - 8) x = window.innerWidth - mw - 8;
    if (y + mh > window.innerHeight - 8) y = window.innerHeight - mh - 8;
    menu.style.left = x + "px";
    menu.style.top = y + "px";
  });
}

// ---------------------------------------------------------------
// Picture editing — select an image (or shape), show resize handles and a
// floating toolbar for border / rounding / alignment / alt / replace / delete.
// ---------------------------------------------------------------
export function attachImageEditing(editor, page) {
  const SELECTABLE = "img, span.shape";
  let sel = null;

  const overlay = el("div", { class: "img-overlay hidden" });
  const HANDLES = ["nw", "ne", "se", "sw", "n", "e", "s", "w"];
  for (const h of HANDLES) overlay.appendChild(el("div", { class: "img-handle h-" + h, "data-h": h }));
  const bar = el("div", { class: "img-toolbar hidden" });
  page.appendChild(overlay);
  page.appendChild(bar);

  const isImg = () => sel && sel.tagName === "IMG";
  const changed = () => { reposition(); fireInput(editor); };

  const rgbToHex = (rgb) => {
    const m = String(rgb).match(/(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
    if (!m) return "#4a90d9";
    return "#" + [m[1], m[2], m[3]].map((n) => (+n).toString(16).padStart(2, "0")).join("");
  };

  // ---- toolbar controls ----
  // fill color (shapes only)
  const fillColor = el("input", { type: "color", value: "#4a90d9", title: "Fill color" });
  fillColor.addEventListener("input", () => {
    if (!sel || isImg()) return;
    sel.style.background = fillColor.value;
    changed();
  });

  const borderColor = el("input", { type: "color", value: "#333333", title: "Border color" });
  const thick = el("select", { title: "Border" });
  for (const [v, l] of [["0", "No border"], ["1", "1 px"], ["2", "2 px"], ["3", "3 px"], ["4", "4 px"], ["6", "6 px"]]) {
    thick.appendChild(el("option", { value: v }, [l]));
  }
  const applyBorder = () => {
    if (!sel) return;
    const w = parseInt(thick.value, 10);
    sel.style.border = w ? `${w}px solid ${borderColor.value}` : "";
    changed();
  };
  borderColor.addEventListener("input", applyBorder);
  thick.addEventListener("change", applyBorder);

  // rounded-corner toggle — reflects the current state (highlighted when rounded)
  const radiusBtn = btn("⬭", "Rounded corners", () => {
    if (!sel) return;
    const rounded = parseFloat(getComputedStyle(sel).borderTopLeftRadius) > 2;
    sel.style.borderRadius = rounded ? "0" : "12px";
    syncRadiusBtn();
    changed();
  });
  function syncRadiusBtn() {
    if (!sel) return;
    const rounded = parseFloat(getComputedStyle(sel).borderTopLeftRadius) > 2;
    radiusBtn.classList.toggle("active", rounded);
  }

  function setAlign(mode) {
    if (!sel) return;
    sel.style.float = ""; sel.style.display = ""; sel.style.margin = "";
    if (mode === "left") { sel.style.float = "left"; sel.style.margin = "4px 12px 4px 0"; }
    else if (mode === "right") { sel.style.float = "right"; sel.style.margin = "4px 0 4px 12px"; }
    else if (mode === "center") { sel.style.display = "block"; sel.style.margin = "8px auto"; }
    changed();
  }
  const alignBtns = [
    btn(ALIGN_ICONS.left, "Inline with text", () => setAlign("inline")),
    btn("⇤", "Float left, wrap text", () => setAlign("left")),
    btn(ALIGN_ICONS.center, "Center on its own line", () => setAlign("center")),
    btn("⇥", "Float right, wrap text", () => setAlign("right")),
  ];

  const altBtn = btn("Alt", "Set description (alt text)", () => {
    if (!sel) return;
    const v = prompt("Description (alt text):", sel.getAttribute("alt") || "");
    if (v != null) { sel.setAttribute("alt", v); fireInput(editor); }
  });
  const replaceBtn = btn("↻ Replace", "Replace image", () => {
    if (!isImg()) return;
    const input = el("input", { type: "file", accept: "image/*" });
    input.addEventListener("change", () => {
      const f = input.files && input.files[0];
      if (!f) return;
      const rd = new FileReader();
      rd.onload = () => { sel.src = String(rd.result); changed(); };
      rd.readAsDataURL(f);
    });
    input.click();
  });
  const delBtn = btn("🗑", "Delete", () => {
    if (!sel) return;
    const s = sel; deselect(); s.remove(); fireInput(editor);
  });

  bar.append(fillColor, borderColor, thick, radiusBtn, sep(), ...alignBtns, sep(), altBtn, replaceBtn, delBtn);

  // ---- geometry ----
  function offsetInPage(node) {
    let x = 0, y = 0;
    while (node && node !== page) { x += node.offsetLeft; y += node.offsetTop; node = node.offsetParent; }
    return { x, y };
  }
  function reposition() {
    if (!sel || !editor.contains(sel)) { deselect(); return; }
    const { x, y } = offsetInPage(sel);
    const w = sel.offsetWidth, h = sel.offsetHeight;
    overlay.style.left = x + "px"; overlay.style.top = y + "px";
    overlay.style.width = w + "px"; overlay.style.height = h + "px";
    overlay.classList.remove("hidden");
    bar.classList.remove("hidden");
    replaceBtn.style.display = isImg() ? "" : "none";
    fillColor.style.display = isImg() ? "none" : ""; // fill is shape-only
    const bw = bar.offsetWidth;
    let bx = x;
    if (bx + bw > page.clientWidth) bx = Math.max(0, page.clientWidth - bw);
    bar.style.left = bx + "px";
    bar.style.top = Math.max(0, y - bar.offsetHeight - 8) + "px";
  }
  function select(elm) {
    sel = elm;
    const cs = getComputedStyle(elm);
    const bwid = parseInt(cs.borderTopWidth, 10) || 0;
    const hasBorder = cs.borderTopStyle !== "none" && bwid > 0;
    thick.value = hasBorder && [1, 2, 3, 4, 6].includes(bwid) ? String(bwid) : "0";
    if (hasBorder) borderColor.value = rgbToHex(cs.borderTopColor);
    if (elm.tagName !== "IMG") fillColor.value = rgbToHex(cs.backgroundColor);
    syncRadiusBtn();
    reposition();
    editor.focus(); // so Delete / Escape reach the keydown handler below
  }
  function deselect() {
    sel = null;
    overlay.classList.add("hidden");
    bar.classList.add("hidden");
  }
  function externalDeselect() { deselect(); }
  const isFloating = (elm) => elm && getComputedStyle(elm).position === "absolute";
  // let the shape picker select a freshly-inserted shape
  editor._selectShape = (elm) => { if (elm) select(elm); };

  editor.addEventListener("mousedown", (e) => {
    const t = e.target.closest && e.target.closest(SELECTABLE);
    if (t && editor.contains(t)) {
      e.preventDefault();
      select(t);
      if (isFloating(t)) startMove(t, e);
    } else if (!overlay.contains(e.target) && !bar.contains(e.target)) {
      deselect();
    }
  });

  // ---- move dragging (free-floating shapes) ----
  function startMove(elm, e) {
    const zoom = parseFloat(getComputedStyle(page).zoom) || 1;
    const x0 = e.clientX, y0 = e.clientY;
    const left0 = parseFloat(elm.style.left) || 0;
    const top0 = parseFloat(elm.style.top) || 0;
    let moved = false;
    const mv = (ev) => {
      moved = true;
      elm.style.left = (left0 + (ev.clientX - x0) / zoom) + "px";
      elm.style.top = (top0 + (ev.clientY - y0) / zoom) + "px";
      reposition();
    };
    const up = () => {
      document.removeEventListener("mousemove", mv);
      document.removeEventListener("mouseup", up);
      if (moved) fireInput(editor);
    };
    document.addEventListener("mousemove", mv);
    document.addEventListener("mouseup", up);
  }
  // keep the caret from landing "inside" a shape span on click
  editor.addEventListener("click", (e) => {
    const t = e.target.closest && e.target.closest(SELECTABLE);
    if (t) e.preventDefault();
  });
  editor.addEventListener("input", () => { if (sel) reposition(); });
  editor.addEventListener("keydown", (e) => {
    if (!sel) return;
    if (e.key === "Delete" || e.key === "Backspace") { e.preventDefault(); delBtn.click(); }
    else if (e.key === "Escape") deselect();
  });
  window.addEventListener("resize", () => { if (sel) reposition(); });

  // ---- resize dragging ----
  let drag = null;
  overlay.addEventListener("mousedown", (e) => {
    const h = e.target.getAttribute && e.target.getAttribute("data-h");
    if (!h || !sel) return;
    e.preventDefault(); e.stopPropagation();
    const zoom = parseFloat(getComputedStyle(page).zoom) || 1;
    drag = { h, x0: e.clientX, y0: e.clientY, w0: sel.offsetWidth, h0: sel.offsetHeight, zoom,
             ratio: sel.offsetWidth / Math.max(1, sel.offsetHeight) };
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  });
  function onMove(e) {
    if (!drag) return;
    const dx = (e.clientX - drag.x0) / drag.zoom;
    const dy = (e.clientY - drag.y0) / drag.zoom;
    let w = drag.w0, h = drag.h0;
    if (drag.h.includes("e")) w = drag.w0 + dx;
    if (drag.h.includes("w")) w = drag.w0 - dx;
    if (drag.h.includes("s")) h = drag.h0 + dy;
    if (drag.h.includes("n")) h = drag.h0 - dy;
    const corner = drag.h.length === 2;
    if (corner) h = w / drag.ratio;
    w = Math.max(16, Math.round(w));
    h = Math.max(10, Math.round(h));
    if (isImg()) {
      sel.setAttribute("width", w);
      sel.style.width = w + "px";
      if (corner) { sel.removeAttribute("height"); sel.style.height = "auto"; }
      else { sel.setAttribute("height", h); sel.style.height = h + "px"; }
    } else {
      sel.style.width = w + "px";
      sel.style.height = h + "px";
    }
    reposition();
  }
  function onUp() {
    drag = null;
    document.removeEventListener("mousemove", onMove);
    document.removeEventListener("mouseup", onUp);
    fireInput(editor);
  }

  return { deselect: externalDeselect };
}

// ---------------------------------------------------------------
// Track changes engine
// ---------------------------------------------------------------
export function createTrackChanges(editor, getAuthor) {
  let enabled = false;
  const now = () => String(Date.now());

  function makeWrapper(tag) {
    const n = document.createElement(tag);
    n.className = tag === "ins" ? "tc-ins" : "tc-del";
    n.setAttribute("data-author", getAuthor());
    n.setAttribute("data-ts", now());
    return n;
  }

  function ownIns(node) {
    let n = node && node.nodeType === 3 ? node.parentElement : node;
    const ins = n && n.closest ? n.closest("ins.tc-ins") : null;
    return ins && editor.contains(ins) && ins.getAttribute("data-author") === getAuthor() ? ins : null;
  }

  function intersectedBlocks(range) {
    const blocks = [];
    const walker = document.createTreeWalker(editor, NodeFilter.SHOW_ELEMENT);
    let n;
    while ((n = walker.nextNode())) {
      if (/^(P|DIV|H[1-6]|LI|BLOCKQUOTE|TD|TH|PRE)$/.test(n.tagName) && range.intersectsNode(n)) blocks.push(n);
    }
    const leaves = blocks.filter((b) => !blocks.some((o) => o !== b && b.contains(o)));
    if (!leaves.length) leaves.push(editor);
    return leaves;
  }
  function clampToBlock(range, block) {
    const sub = document.createRange();
    sub.selectNodeContents(block);
    try {
      if (sub.comparePoint(range.startContainer, range.startOffset) === 0) {
        sub.setStart(range.startContainer, range.startOffset);
      }
      if (sub.comparePoint(range.endContainer, range.endOffset) === 0) {
        sub.setEnd(range.endContainer, range.endOffset);
      }
    } catch { return null; }
    return sub;
  }

  function wrapRangeDel(sub) {
    if (!sub || sub.collapsed) return null;
    const frag = sub.extractContents();
    const scratch = document.createElement("div");
    scratch.appendChild(frag);
    // deleting our own unaccepted insertion really removes it
    for (const ins of [...scratch.querySelectorAll("ins.tc-ins")]) {
      if (ins.getAttribute("data-author") === getAuthor()) ins.remove();
    }
    // flatten nested tracked deletions (content is already marked deleted)
    for (const d of [...scratch.querySelectorAll("del.tc-del")]) unwrapNode(d);
    if (!scratch.textContent.length && !scratch.querySelector("img")) return null;
    const del = makeWrapper("del");
    while (scratch.firstChild) del.appendChild(scratch.firstChild);
    sub.insertNode(del);
    return del;
  }

  function deleteTracked(direction, granularity = "character") {
    const sel = window.getSelection();
    if (!sel.rangeCount) return;
    if (sel.isCollapsed) {
      try { sel.modify("extend", direction, granularity); } catch {}
      if (sel.isCollapsed) return; // document edge
    }
    const range = sel.getRangeAt(0);
    // selection entirely inside our own insertion: plain delete
    const insA = ownIns(range.startContainer), insB = ownIns(range.endContainer);
    if (insA && insA === insB) {
      range.deleteContents();
      if (!insA.textContent && !insA.querySelector("img")) {
        const r = document.createRange();
        r.setStartBefore(insA);
        insA.remove();
        sel.removeAllRanges();
        r.collapse(true);
        sel.addRange(r);
      }
      fireInput(editor);
      return;
    }
    const blocks = intersectedBlocks(range);
    const dels = [];
    for (let i = blocks.length - 1; i >= 0; i--) {
      const sub = clampToBlock(range, blocks[i]);
      const d = wrapRangeDel(sub);
      if (d) dels.unshift(d); // document order
    }
    sel.removeAllRanges();
    if (dels.length) {
      const r = document.createRange();
      if (direction === "backward") r.setStartBefore(dels[0]);
      else r.setStartAfter(dels[dels.length - 1]);
      r.collapse(true);
      sel.addRange(r);
    }
    fireInput(editor);
  }

  function placeCaretInside(node) {
    const sel = window.getSelection();
    const r = document.createRange();
    r.selectNodeContents(node);
    r.collapse(false);
    sel.removeAllRanges();
    sel.addRange(r);
  }

  function insertTracked(text) {
    const sel = window.getSelection();
    if (!sel.rangeCount) return;
    if (!sel.isCollapsed) deleteTracked("forward");
    const range = window.getSelection().getRangeAt(0);
    const ins = makeWrapper("ins");
    ins.textContent = text;
    range.insertNode(ins);
    placeCaretInside(ins);
    fireInput(editor);
  }

  function insertTrackedHtml(html) {
    const sel = window.getSelection();
    if (!sel.rangeCount) return;
    if (!sel.isCollapsed) deleteTracked("forward");
    const tpl = document.createElement("template");
    tpl.innerHTML = html;
    const hasBlock = tpl.content.querySelector("p,div,h1,h2,h3,h4,h5,h6,ul,ol,table,blockquote,pre");
    if (hasBlock) {
      // block-level paste isn't wrapped (Word tracks it structurally; we accept the gap)
      exec("insertHTML", html);
      fireInput(editor);
      return;
    }
    const range = window.getSelection().getRangeAt(0);
    const ins = makeWrapper("ins");
    ins.appendChild(tpl.content);
    range.insertNode(ins);
    placeCaretInside(ins);
    fireInput(editor);
  }

  function strictlyInsideOwnIns(sel, edge) {
    const ins = ownIns(sel.anchorNode);
    if (!ins) return false;
    const r0 = document.createRange();
    r0.selectNodeContents(ins);
    if (edge === "start") r0.setEnd(sel.anchorNode, sel.anchorOffset);
    else r0.setStart(sel.anchorNode, sel.anchorOffset);
    return r0.toString().length > 0;
  }

  function onBeforeInput(e) {
    if (!enabled) return;
    const t = e.inputType;
    if (t === "insertText") {
      const sel = window.getSelection();
      if (sel.isCollapsed && ownIns(sel.anchorNode)) return; // native typing extends own ins
      e.preventDefault();
      insertTracked(e.data || "");
    } else if (t === "deleteContentBackward" || t === "deleteWordBackward") {
      const sel = window.getSelection();
      if (sel.isCollapsed && strictlyInsideOwnIns(sel, "start")) return; // native delete of own insertion
      e.preventDefault();
      deleteTracked("backward", t.includes("Word") ? "word" : "character");
    } else if (t === "deleteContentForward" || t === "deleteWordForward") {
      const sel = window.getSelection();
      if (sel.isCollapsed && strictlyInsideOwnIns(sel, "end")) return;
      e.preventDefault();
      deleteTracked("forward", t.includes("Word") ? "word" : "character");
    } else if (t === "insertFromDrop") {
      e.preventDefault();
      const dt = e.dataTransfer;
      const html = dt && dt.getData("text/html");
      const text = dt && dt.getData("text/plain");
      if (html) insertTrackedHtml(sanitizeHtml(html));
      else if (text) insertTracked(text);
    }
    // insertParagraph, IME composition, formatting commands: native behavior
  }
  editor.addEventListener("beforeinput", onBeforeInput);

  // cut while tracking: copy manually, then mark deleted
  editor.addEventListener("cut", (e) => {
    if (!enabled) return;
    const sel = window.getSelection();
    if (!sel.rangeCount || sel.isCollapsed) return;
    e.preventDefault();
    const div = document.createElement("div");
    div.appendChild(sel.getRangeAt(0).cloneContents());
    try {
      e.clipboardData.setData("text/html", div.innerHTML);
      e.clipboardData.setData("text/plain", sel.toString());
    } catch {}
    deleteTracked("forward");
  });

  function accept(node) {
    if (!node) return;
    if (node.matches && node.matches("ins.tc-ins")) unwrapNode(node);
    else node.remove();
    fireInput(editor);
  }
  function reject(node) {
    if (!node) return;
    if (node.matches && node.matches("ins.tc-ins")) node.remove();
    else unwrapNode(node);
    fireInput(editor);
  }

  return {
    setEnabled(v) { enabled = !!v; },
    isEnabled() { return enabled; },
    insertText: insertTracked,
    insertHtml: insertTrackedHtml,
    accept,
    reject,
    acceptAll() {
      for (const n of [...editor.querySelectorAll("ins.tc-ins, del.tc-del")]) {
        n.tagName === "INS" ? unwrapNode(n) : n.remove();
      }
      fireInput(editor);
    },
    rejectAll() {
      for (const n of [...editor.querySelectorAll("ins.tc-ins, del.tc-del")]) {
        n.tagName === "INS" ? n.remove() : unwrapNode(n);
      }
      fireInput(editor);
    },
    list() {
      return [...editor.querySelectorAll("ins.tc-ins, del.tc-del")].map((n, i) => ({
        index: i,
        node: n,
        type: n.tagName === "INS" ? "insertion" : "deletion",
        author: n.getAttribute("data-author") || "?",
        ts: parseInt(n.getAttribute("data-ts"), 10) || null,
        text: (n.textContent || "").slice(0, 60),
      }));
    },
  };
}

// ---------------------------------------------------------------
// Comment anchors
// ---------------------------------------------------------------
export function wrapSelectionComment(editor, cid) {
  editor.focus();
  restoreSelection(editor);
  const sel = window.getSelection();
  if (!sel.rangeCount) return false;
  let range = sel.getRangeAt(0);
  if (range.collapsed) {
    // expand to the word under the caret
    try {
      sel.modify("move", "backward", "word");
      sel.modify("extend", "forward", "word");
      range = sel.getRangeAt(0);
    } catch {}
    if (range.collapsed) return false;
  }
  // clamp to a single block so the anchor never swallows block structure
  const startBlock = closestBlock(range.startContainer, editor);
  if (startBlock && !startBlock.contains(range.endContainer)) {
    range.setEnd(startBlock, startBlock.childNodes.length);
  }
  if (!range.toString().trim()) return false;
  const span = document.createElement("span");
  span.className = "comment-ref";
  span.setAttribute("data-cid", cid);
  try {
    range.surroundContents(span);
  } catch {
    const frag = range.extractContents();
    span.appendChild(frag);
    range.insertNode(span);
  }
  saveSelection(editor);
  fireInput(editor);
  return true;
}

// ---------------------------------------------------------------
// Format painter
// ---------------------------------------------------------------
const TAGGABLE_BLOCKS = /^(P|H[1-6]|BLOCKQUOTE|PRE)$/;

// Format painter: capture character + paragraph formatting at the caret/anchor.
function captureFormat(editor) {
  const sel = window.getSelection();
  if (!sel.rangeCount) return null;
  let n = sel.anchorNode;
  if (n && n.nodeType === 3) n = n.parentElement;
  if (!n || !editor.contains(n)) return null;
  const cs = getComputedStyle(n);
  const edCs = getComputedStyle(editor);
  const td = cs.textDecorationLine || cs.textDecoration || "";
  const inline = {};
  if (parseInt(cs.fontWeight, 10) >= 600) inline.fontWeight = "bold";
  if (cs.fontStyle && cs.fontStyle !== "normal") inline.fontStyle = cs.fontStyle;
  if (td && td !== "none") inline.textDecoration = td;
  if (cs.color && cs.color !== edCs.color) inline.color = cs.color;
  const bg = cs.backgroundColor;
  if (bg && bg !== "rgba(0, 0, 0, 0)") inline.backgroundColor = bg;
  const fam = cs.fontFamily.split(",")[0].replace(/['"]/g, "");
  const edFam = edCs.fontFamily.split(",")[0].replace(/['"]/g, "");
  if (fam !== edFam) inline.fontFamily = fam;
  const size = Math.round(parseFloat(cs.fontSize) * 0.75) + "pt";
  const edSize = Math.round(parseFloat(edCs.fontSize) * 0.75) + "pt";
  if (size !== edSize) inline.fontSize = size;

  const block = closestBlock(n, editor);
  let blockFmt = null;
  if (block) {
    const bcs = getComputedStyle(block);
    const align = block.style.textAlign || bcs.textAlign;
    blockFmt = {
      tag: block.tagName,
      align: align && align !== "start" && align !== "left" ? align : "",
      lineHeight: block.style.lineHeight || (bcs.lineHeight !== "normal" ? bcs.lineHeight : ""),
      indent: block.style.paddingLeft || block.style.marginLeft || "",
    };
  }
  return { inline, block: blockFmt };
}

function changeBlockTag(block, tag) {
  if (block.tagName === tag) return block;
  const nb = document.createElement(tag.toLowerCase());
  while (block.firstChild) nb.appendChild(block.firstChild);
  for (const a of [...block.attributes]) { if (a.name === "class") continue; nb.setAttribute(a.name, a.value); }
  block.replaceWith(nb);
  return nb;
}

// Apply captured paragraph formatting to the blocks overlapping the selection
// (or the caret's block when the selection is collapsed).
function applyBlockFormat(editor, bf) {
  if (!bf) return;
  let targets = selectedBlocks(editor);
  if (!targets.length) {
    const sel = window.getSelection();
    let n = sel.rangeCount ? sel.anchorNode : null;
    if (n && n.nodeType === 3) n = n.parentElement;
    const b = n && closestBlock(n, editor);
    if (b) targets = [b];
  }
  for (const b of targets) {
    if (bf.align) b.style.textAlign = bf.align;
    if (bf.lineHeight) b.style.lineHeight = bf.lineHeight;
    if (bf.indent) b.style.paddingLeft = bf.indent;
    if (bf.tag && TAGGABLE_BLOCKS.test(bf.tag) && TAGGABLE_BLOCKS.test(b.tagName)) {
      changeBlockTag(b, bf.tag);
    }
  }
  fireInput(editor);
}

// ---------------------------------------------------------------
// Toolbar
// ---------------------------------------------------------------
export function buildToolbar(editor, hooks = {}) {
  const bar = el("div", { class: "toolbar" });
  try { document.execCommand("styleWithCSS", false, true); } catch {}

  const refocus = () => { editor.focus(); };

  // undo/redo
  const undoBtn = btn(UNDO_ICON, "Undo (Ctrl+Z)", () => {
    if (hooks.history) hooks.history.undo(); else exec("undo");
    refocus();
  });
  const redoBtn = btn(REDO_ICON, "Redo (Ctrl+Y / Ctrl+Shift+Z)", () => {
    if (hooks.history) hooks.history.redo(); else exec("redo");
    refocus();
  });
  bar.append(undoBtn, redoBtn, sep());
  if (hooks.history) {
    undoBtn.disabled = true;
    redoBtn.disabled = true;
    hooks.history.onUpdate = (u, r) => { undoBtn.disabled = !u; redoBtn.disabled = !r; };
  }

  // format painter + clear
  // Single click loads format from the caret/selection and applies it once to
  // the next selection. Double-click locks the brush (apply many times).
  // Escape or another click cancels.
  let painter = null;
  let painterSticky = false;
  const painterBtn = btn("🖌", "Format painter — pick up formatting, then select target (double-click to lock)", (b) => {
    if (painter) { cancelPainter(); return; }
    painter = captureFormat(editor);
    if (painter) b.classList.add("active");
  });
  function cancelPainter() {
    painter = null; painterSticky = false;
    painterBtn.classList.remove("active");
  }
  painterBtn.addEventListener("dblclick", (e) => {
    e.preventDefault();
    painter = captureFormat(editor);
    painterSticky = !!painter;
    painterBtn.classList.add("active");
  });
  editor.addEventListener("keyup", (e) => { if (e.key === "Escape" && painter) cancelPainter(); });
  editor.addEventListener("mouseup", () => {
    if (!painter) return;
    const sel = window.getSelection();
    saveSelection(editor);
    applyBlockFormat(editor, painter.block);
    if (!sel.isCollapsed) applySpanStyle(editor, painter.inline);
    if (!painterSticky) cancelPainter();
  });
  bar.append(painterBtn, btn("⌫", "Clear formatting", () => { exec("removeFormat"); exec("formatBlock", "<p>"); refocus(); }), sep());

  // paragraph style
  const styleSel = el("select", { title: "Paragraph style" });
  for (const [v, label] of [["P", "Normal"], ["H1", "Heading 1"], ["H2", "Heading 2"], ["H3", "Heading 3"], ["H4", "Heading 4"], ["H5", "Heading 5"], ["H6", "Heading 6"], ["BLOCKQUOTE", "Quote"], ["PRE", "Code block"]]) {
    styleSel.appendChild(el("option", { value: v }, [label]));
  }
  styleSel.addEventListener("change", () => { exec("formatBlock", "<" + styleSel.value.toLowerCase() + ">"); refocus(); });
  bar.appendChild(styleSel);

  // font family / size
  const fontSel = el("select", { title: "Font family" });
  fontSel.appendChild(el("option", { value: "" }, ["Font"]));
  for (const f of FONTS) {
    const o = el("option", { value: f }, [f]);
    o.style.fontFamily = f;
    fontSel.appendChild(o);
  }
  fontSel.addEventListener("change", () => { if (fontSel.value) exec("fontName", fontSel.value); refocus(); });
  bar.appendChild(fontSel);

  const sizeSel = el("select", { title: "Font size (pt)" });
  sizeSel.appendChild(el("option", { value: "" }, ["Size"]));
  for (const s of SIZES) sizeSel.appendChild(el("option", { value: String(s) }, [String(s)]));
  sizeSel.addEventListener("change", () => {
    if (sizeSel.value) applySpanStyle(editor, { fontSize: sizeSel.value + "pt" });
    refocus();
  });
  bar.appendChild(sizeSel);

  bar.appendChild(sep());

  const fmtBtns = {
    bold: btn("<b>B</b>", "Bold (Ctrl+B)", () => { exec("bold"); refocus(); }),
    italic: btn("<i>I</i>", "Italic (Ctrl+I)", () => { exec("italic"); refocus(); }),
    underline: btn("<u>U</u>", "Underline (Ctrl+U)", () => { exec("underline"); refocus(); }),
    strikeThrough: btn("<s>S</s>", "Strikethrough", () => { exec("strikeThrough"); refocus(); }),
    subscript: btn("x₂", "Subscript", () => { exec("subscript"); refocus(); }),
    superscript: btn("x²", "Superscript", () => { exec("superscript"); refocus(); }),
  };
  bar.append(...Object.values(fmtBtns), sep());

  // colors
  let activeFg = "#c0392b";
  const colorInput = el("input", { type: "color", value: activeFg, title: "Text color" });
  const colorBtn = btn('<span style="border-bottom:3px solid #c0392b">A</span>', "Text color", () => colorInput.click());
  function applyFg(color) {
    restoreSelection(editor);
    exec("foreColor", color);
    activeFg = color;
    colorBtn.firstChild.style.borderBottomColor = color;
    colorInput.value = color;
    saveSelection(editor);
  }
  colorInput.addEventListener("change", () => { applyFg(colorInput.value); refocus(); });
  // keep the active color for subsequent typing at a collapsed caret
  editor.addEventListener("keydown", function _fgKey(e) {
    if (e.key.length !== 1 || e.ctrlKey || e.metaKey || e.altKey) return;
    const s = window.getSelection();
    if (!s || !s.isCollapsed || !s.rangeCount || !editor.contains(s.anchorNode)) return;
    exec("foreColor", activeFg);
  });
  // reflect the current caret's text color in the button indicator
  function sampleColor() {
    const s = window.getSelection();
    if (!s || !s.rangeCount || !editor.contains(s.anchorNode)) return;
    const node = s.anchorNode;
    const el = node.nodeType === 3 ? node.parentElement : node;
    if (!el) return;
    colorBtn.firstChild.style.borderBottomColor = getComputedStyle(el).color;
  }
  editor.addEventListener("mouseup", sampleColor);
  editor.addEventListener("keyup", sampleColor);

  const hlWrap = el("span", { class: "dropdown" });
  const hlBtn = btn('<span style="background:#ffff00;padding:0 4px">ab</span>', "Highlight", () => hlMenu.classList.toggle("hidden"));
  const hlMenu = el("div", { class: "dropdown-menu hidden" });
  for (const c of HIGHLIGHTS) {
    const sw = el("button", { type: "button", class: "swatch", title: c });
    sw.style.background = c === "transparent" ? "linear-gradient(to top right, #fff 45%, #f00 50%, #fff 55%)" : c;
    sw.addEventListener("mousedown", (e) => e.preventDefault());
    sw.addEventListener("click", () => {
      restoreSelection(editor);
      try { exec("hiliteColor", c === "transparent" ? "inherit" : c); } catch { exec("backColor", c); }
      hlMenu.classList.add("hidden");
    });
    hlMenu.appendChild(sw);
  }
  hlWrap.append(hlBtn, hlMenu);
  document.addEventListener("click", (e) => { if (!hlWrap.contains(e.target)) hlMenu.classList.add("hidden"); });

  bar.append(colorBtn, colorInput, hlWrap, sep());

  // alignment
  const alignBtns = {
    justifyLeft: btn(ALIGN_ICONS.left, "Align left", () => { exec("justifyLeft"); refocus(); }),
    justifyCenter: btn(ALIGN_ICONS.center, "Align center", () => { exec("justifyCenter"); refocus(); }),
    justifyRight: btn(ALIGN_ICONS.right, "Align right", () => { exec("justifyRight"); refocus(); }),
    justifyFull: btn(ALIGN_ICONS.justify, "Justify", () => { exec("justifyFull"); refocus(); }),
  };
  bar.append(...Object.values(alignBtns), sep());

  // lists + indent
  bar.append(
    btn("• ⎯", "Bullet list", () => { exec("insertUnorderedList"); refocus(); }),
    btn("1. ⎯", "Numbered list", () => { exec("insertOrderedList"); refocus(); }),
    btn("⇥", "Increase indent", () => { exec("indent"); refocus(); }),
    btn("⇤", "Decrease indent", () => { exec("outdent"); refocus(); }),
  );

  // line spacing
  const lsSel = el("select", { title: "Line spacing" });
  lsSel.appendChild(el("option", { value: "" }, ["↕"]));
  for (const v of ["1", "1.15", "1.5", "2", "2.5"]) lsSel.appendChild(el("option", { value: v }, [v]));
  lsSel.addEventListener("change", () => {
    if (!lsSel.value) return;
    for (const b of selectedBlocks(editor)) b.style.lineHeight = lsSel.value;
    lsSel.value = "";
    fireInput(editor);
    refocus();
  });
  bar.append(lsSel, sep());

  // review
  const commentBtn = btn("💬", "Add comment on selection", () => hooks.addComment && hooks.addComment());
  const trackBtn = btn("📝", "Track changes on/off", (b) => {
    if (!hooks.toggleTrack) return;
    const on = hooks.toggleTrack();
    b.classList.toggle("active", on);
  });
  bar.append(commentBtn, trackBtn, sep());
  bar.trackBtn = trackBtn; // let the app reflect externally-set state

  // find / page setup / print
  bar.append(
    btn("🔍", "Find & replace (Ctrl+F)", () => hooks.toggleFind && hooks.toggleFind()),
    btn("📄", "Page setup", () => hooks.pageSetup && hooks.pageSetup()),
    btn("🖨", "Print / PDF (Ctrl+P)", () => hooks.print && hooks.print()),
  );

  // active-state reflection
  document.addEventListener("selectionchange", () => {
    const sel = window.getSelection();
    if (!sel.rangeCount || !editor.contains(sel.anchorNode)) return;
    for (const [cmd, b] of Object.entries(fmtBtns)) {
      try { b.classList.toggle("active", document.queryCommandState(cmd)); } catch {}
    }
    for (const [cmd, b] of Object.entries(alignBtns)) {
      try { b.classList.toggle("active", document.queryCommandState(cmd)); } catch {}
    }
    const block = closestBlock(sel.anchorNode, editor);
    if (block) {
      const t = block.tagName;
      styleSel.value = /^H[1-6]$/.test(t) ? t : t === "BLOCKQUOTE" ? "BLOCKQUOTE" : t === "PRE" ? "PRE" : "P";
    }
    // reflect current font family / size at the caret
    let n = sel.anchorNode;
    if (n && n.nodeType === 3) n = n.parentElement;
    if (n && editor.contains(n)) {
      const cs = getComputedStyle(n);
      const ff = (cs.fontFamily || "").split(",")[0].trim().replace(/^["']|["']$/g, "");
      let matchedFont = "";
      for (const opt of fontSel.options) {
        if (opt.value && opt.value.toLowerCase() === ff.toLowerCase()) { matchedFont = opt.value; break; }
      }
      fontSel.value = matchedFont;
      const pt = Math.round(parseFloat(cs.fontSize) * 0.75);
      let matchedSize = "";
      for (const opt of sizeSel.options) {
        if (opt.value && Number(opt.value) === pt) { matchedSize = opt.value; break; }
      }
      sizeSel.value = matchedSize;
    }
  });

  return bar;
}

// ---------------------------------------------------------------
// Editor behaviors: paste sanitize, selection tracking, shortcuts
// ---------------------------------------------------------------
export function attachEditorBehaviors(editor, opts = {}) {
  // selection bookkeeping for toolbar actions. No save on "focus": refocusing
  // after a dialog would clobber the saved range with a collapsed caret.
  editor.addEventListener("keyup", () => saveSelection(editor));
  editor.addEventListener("mouseup", () => saveSelection(editor));
  document.addEventListener("selectionchange", () => {
    const sel = window.getSelection();
    if (sel.rangeCount && editor.contains(sel.anchorNode)) saveSelection(editor);
  });

  // sanitize pasted content; route through track-changes when recording
  editor.addEventListener("paste", (e) => {
    const html = e.clipboardData && e.clipboardData.getData("text/html");
    const text = e.clipboardData && e.clipboardData.getData("text/plain");
    if (opts.track && opts.track.isEnabled()) {
      e.preventDefault();
      if (html) opts.track.insertHtml(sanitizeHtml(html));
      else if (text) opts.track.insertText(text);
      return;
    }
    if (html) {
      e.preventDefault();
      exec("insertHTML", sanitizeHtml(html));
    }
    // plain text falls through to native handling
  });

  // ctrl/cmd-click opens links
  editor.addEventListener("click", (e) => {
    const a = e.target.closest && e.target.closest("a[href]");
    if (a && (e.metaKey || e.ctrlKey)) {
      e.preventDefault();
      window.open(a.href, "_blank", "noopener");
    }
  });

  editor.addEventListener("keydown", (e) => {
    const mod = e.metaKey || e.ctrlKey;
    if (!mod) {
      // Tab inside a table cell moves to the next cell
      if (e.key === "Tab") {
        const sel = window.getSelection();
        let n = sel.anchorNode;
        if (n && n.nodeType === 3) n = n.parentElement;
        const cell = n && n.closest && n.closest("td,th");
        if (cell && editor.contains(cell)) {
          e.preventDefault();
          const cells = [...cell.closest("table").querySelectorAll("td,th")];
          const i = cells.indexOf(cell);
          const next = cells[i + (e.shiftKey ? -1 : 1)];
          if (next) {
            const r = document.createRange();
            r.selectNodeContents(next);
            r.collapse(true);
            sel.removeAllRanges();
            sel.addRange(r);
          }
          return;
        }
        e.preventDefault();
        exec(e.shiftKey ? "outdent" : "indent");
      }
      return;
    }
    const k = e.key.toLowerCase();
    if (k === "z") {
      e.preventDefault();
      if (opts.history) (e.shiftKey ? opts.history.redo() : opts.history.undo());
      else exec(e.shiftKey ? "redo" : "undo");
      return;
    }
    if (k === "y") {
      e.preventDefault();
      if (opts.history) opts.history.redo(); else exec("redo");
      return;
    }
    if (k === "k") { e.preventDefault(); openLinkDialog(editor); }
    else if (e.altKey && /^[0-6]$/.test(e.key)) {
      e.preventDefault();
      exec("formatBlock", e.key === "0" ? "<p>" : `<h${e.key}>`);
    }
  });
}

export function countWords(editor) {
  let src = editor;
  if (editor.querySelector("del.tc-del")) {
    src = editor.cloneNode(true);
    for (const d of src.querySelectorAll("del.tc-del")) d.remove();
  }
  const text = src.innerText || src.textContent || "";
  const words = (text.match(/[\p{L}\p{N}]+(?:['’-][\p{L}\p{N}]+)*/gu) || []).length;
  const chars = text.replace(/\s/g, "").length;
  return { words, chars };
}
