// docx.js — dependency-free .docx read/write.
// Uses the browser's native CompressionStream/DecompressionStream ("deflate-raw")
// for ZIP inflate/deflate, plus manual ZIP container parsing and OOXML (Word ML)
// XML mapping. No external libraries.
//
// Import: paragraphs, runs (b/i/u/strike/sub/sup/color/highlight/size/font),
// headings, alignment, indentation, line spacing, native numbered/bulleted
// lists (numbering.xml + style-based numbering), tables (gridSpan/vMerge/shading),
// inline images, hyperlinks, page setup (sectPr), document title (core.xml),
// tracked changes (w:ins/w:del), comments (comments.xml + comment ranges).
// Export: the same set, generating a complete valid OOXML package.

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

export function supportsDocx() {
  return typeof DecompressionStream !== "undefined" && typeof CompressionStream !== "undefined";
}

// ---------------- CRC32 ----------------
const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();
function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

// ---------------- raw deflate via native streams ----------------
async function streamCollect(readable) {
  const reader = readable.getReader();
  const chunks = [];
  let len = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    len += value.length;
  }
  const out = new Uint8Array(len);
  let o = 0;
  for (const c of chunks) { out.set(c, o); o += c.length; }
  return out;
}
async function inflateRaw(data) {
  const ds = new DecompressionStream("deflate-raw");
  const w = ds.writable.getWriter();
  w.write(data);
  w.close();
  return streamCollect(ds.readable);
}
async function deflateRaw(data) {
  const cs = new CompressionStream("deflate-raw");
  const w = cs.writable.getWriter();
  w.write(data);
  w.close();
  return streamCollect(cs.readable);
}

// ---------------- ZIP read ----------------
export async function unzip(arrayBuffer) {
  const view = new DataView(arrayBuffer);
  const bytes = new Uint8Array(arrayBuffer);
  const len = arrayBuffer.byteLength;
  let eocd = -1;
  for (let i = len - 22; i >= 0 && i >= len - 65557; i--) {
    if (view.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error("Not a ZIP/DOCX file");
  const cdCount = view.getUint16(eocd + 10, true);
  let cdOffset = view.getUint32(eocd + 16, true);
  const files = new Map();
  for (let i = 0; i < cdCount; i++) {
    if (view.getUint32(cdOffset, true) !== 0x02014b50) break;
    const method = view.getUint16(cdOffset + 10, true);
    const compSize = view.getUint32(cdOffset + 20, true);
    const nameLen = view.getUint16(cdOffset + 28, true);
    const extraLen = view.getUint16(cdOffset + 30, true);
    const commentLen = view.getUint16(cdOffset + 32, true);
    const localHeader = view.getUint32(cdOffset + 42, true);
    const name = new TextDecoder().decode(bytes.subarray(cdOffset + 46, cdOffset + 46 + nameLen));
    const lNameLen = view.getUint16(localHeader + 26, true);
    const lExtraLen = view.getUint16(localHeader + 28, true);
    const dataStart = localHeader + 30 + lNameLen + lExtraLen;
    const compData = bytes.subarray(dataStart, dataStart + compSize);
    let content;
    if (method === 0) content = new Uint8Array(compData);
    else if (method === 8) content = await inflateRaw(compData);
    else throw new Error("Unsupported zip method " + method);
    files.set(name, content);
    cdOffset += 46 + nameLen + extraLen + commentLen;
  }
  return files;
}

// ---------------- ZIP write ----------------
function u16(n) { return new Uint8Array([n & 0xff, (n >>> 8) & 0xff]); }
function u32(n) { return new Uint8Array([n & 0xff, (n >>> 8) & 0xff, (n >>> 16) & 0xff, (n >>> 24) & 0xff]); }
function concat(arrs) {
  let len = 0; for (const a of arrs) len += a.length;
  const out = new Uint8Array(len); let o = 0;
  for (const a of arrs) { out.set(a, o); o += a.length; }
  return out;
}
export async function zip(files) {
  const enc = new TextEncoder();
  const localParts = [];
  const central = [];
  let offset = 0;
  for (const [name, data] of files) {
    const nameBytes = enc.encode(name);
    const crc = crc32(data);
    const comp = await deflateRaw(data);
    const local = concat([
      u32(0x04034b50), u16(20), u16(0), u16(8), u16(0), u16(0),
      u32(crc), u32(comp.length), u32(data.length),
      u16(nameBytes.length), u16(0), nameBytes, comp,
    ]);
    localParts.push(local);
    central.push(concat([
      u32(0x02014b50), u16(20), u16(20), u16(0), u16(8), u16(0), u16(0),
      u32(crc), u32(comp.length), u32(data.length),
      u16(nameBytes.length), u16(0), u16(0), u16(0), u16(0), u32(0), u32(offset), nameBytes,
    ]));
    offset += local.length;
  }
  const centralBuf = concat(central);
  const eocd = concat([
    u32(0x06054b50), u16(0), u16(0), u16(files.size), u16(files.size),
    u32(centralBuf.length), u32(offset), u16(0),
  ]);
  return new Blob([concat([...localParts, centralBuf, eocd])], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

// ---------------- helpers ----------------
function escXml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&apos;")
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
}
function escHtml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
function child(parent, ns, name) {
  for (const c of parent.children) if (c.namespaceURI === ns && c.localName === name) return c;
  return null;
}
function children(parent, ns, name) {
  const out = [];
  for (const c of parent.children) if (c.namespaceURI === ns && c.localName === name) out.push(c);
  return out;
}
function attr(el, name) {
  // OOXML attributes are namespaced (w:val etc.); DOMParser exposes both forms.
  return el.getAttribute("w:val") !== null && name === "val"
    ? el.getAttribute("w:val")
    : (el.getAttributeNS(W, name) ?? el.getAttribute("w:" + name) ?? el.getAttribute(name));
}
function boolProp(el) {
  if (!el) return false;
  const v = attr(el, "val");
  return !(v === "0" || v === "false" || v === "none");
}
function findDesc(el, localName) {
  const found = el.getElementsByTagName("*");
  for (const n of found) if (n.localName === localName) return n;
  return null;
}

const TW_PER_PX = 15;       // 1px = 0.75pt = 15 twips
const EMU_PER_PX = 9525;

// Word named highlight colors <-> css hex
const HIGHLIGHT_COLORS = {
  yellow: "#ffff00", green: "#00ff00", cyan: "#00ffff", magenta: "#ff00ff",
  blue: "#0000ff", red: "#ff0000", darkBlue: "#000080", darkCyan: "#008080",
  darkGreen: "#008000", darkMagenta: "#800080", darkRed: "#800000", darkYellow: "#808000",
  darkGray: "#808080", lightGray: "#c0c0c0", black: "#000000", white: "#ffffff",
};
const HEX_TO_HIGHLIGHT = Object.fromEntries(
  Object.entries(HIGHLIGHT_COLORS).map(([k, v]) => [v.slice(1), k])
);

export const PAGE_SIZES = {
  Letter: { w: 12240, h: 15840 },
  A4: { w: 11906, h: 16838 },
  Legal: { w: 12240, h: 20160 },
  A3: { w: 16838, h: 23811 },
};
export const DEFAULT_PAGE_SETUP = {
  size: "Letter", orientation: "portrait",
  margins: { top: 1, right: 1, bottom: 1, left: 1 }, // inches
};

const IMG_MIME = {
  png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif",
  bmp: "image/bmp", webp: "image/webp", svg: "image/svg+xml", tif: "image/tiff", tiff: "image/tiff",
};

function bytesToDataUrl(bytes, mime) {
  let bin = "";
  const CHUNK = 0x8000;
  for (let i = 0; i < bytes.length; i += CHUNK) {
    bin += String.fromCharCode.apply(null, bytes.subarray(i, i + CHUNK));
  }
  return `data:${mime};base64,${btoa(bin)}`;
}
function dataUrlToBytes(url) {
  const i = url.indexOf(",");
  const meta = url.slice(5, i);
  const b64 = meta.includes("base64");
  const body = url.slice(i + 1);
  const mime = meta.split(";")[0] || "application/octet-stream";
  if (b64) {
    const bin = atob(body);
    const bytes = new Uint8Array(bin.length);
    for (let j = 0; j < bin.length; j++) bytes[j] = bin.charCodeAt(j);
    return { bytes, mime };
  }
  return { bytes: new TextEncoder().encode(decodeURIComponent(body)), mime };
}

// W3CDTF timestamp from ms-epoch, ISO string, or nothing
function tsIso(v) {
  let d;
  if (v == null || v === "") d = new Date();
  else if (/^\d+$/.test(String(v))) d = new Date(parseInt(v, 10));
  else d = new Date(v);
  if (isNaN(d.getTime())) d = new Date();
  return d.toISOString().replace(/\.\d+Z$/, "Z");
}

// ============================================================
// IMPORT: .docx -> { html, pageSetup, title, comments }
// ============================================================

function parseXml(text) {
  const doc = new DOMParser().parseFromString(text, "application/xml");
  if (doc.getElementsByTagName("parsererror")[0]) throw new Error("XML parse error");
  return doc;
}
function decodePart(files, name) {
  const data = files.get(name);
  return data ? new TextDecoder().decode(data) : null;
}

function parseRels(files, relPath) {
  const rels = new Map();
  const text = decodePart(files, relPath);
  if (!text) return rels;
  const doc = parseXml(text);
  for (const rel of doc.getElementsByTagName("Relationship")) {
    rels.set(rel.getAttribute("Id"), {
      type: rel.getAttribute("Type") || "",
      target: rel.getAttribute("Target") || "",
      mode: rel.getAttribute("TargetMode") || "Internal",
    });
  }
  return rels;
}

function parseNumbering(files) {
  // numId -> { ilvl -> "bullet" | "decimal" | ... }
  const text = decodePart(files, "word/numbering.xml");
  const fmtByNum = new Map();
  if (!text) return fmtByNum;
  const doc = parseXml(text);
  const abstracts = new Map();
  for (const ab of doc.getElementsByTagNameNS(W, "abstractNum")) {
    const id = ab.getAttributeNS(W, "abstractNumId") ?? ab.getAttribute("w:abstractNumId");
    const levels = {};
    for (const lvl of children(ab, W, "lvl")) {
      const ilvl = lvl.getAttributeNS(W, "ilvl") ?? lvl.getAttribute("w:ilvl");
      const numFmt = child(lvl, W, "numFmt");
      levels[ilvl] = numFmt ? attr(numFmt, "val") : "decimal";
    }
    abstracts.set(id, levels);
  }
  for (const num of doc.getElementsByTagNameNS(W, "num")) {
    const numId = num.getAttributeNS(W, "numId") ?? num.getAttribute("w:numId");
    const absRef = child(num, W, "abstractNumId");
    const absId = absRef ? attr(absRef, "val") : null;
    fmtByNum.set(numId, abstracts.get(absId) || {});
  }
  return fmtByNum;
}

// styles.xml: styleId -> {numId, ilvl} for styles that carry their own numbering
function parseStylesNumbering(files) {
  const map = new Map();
  const text = decodePart(files, "word/styles.xml");
  if (!text) return map;
  let doc;
  try { doc = parseXml(text); } catch { return map; }
  for (const style of doc.getElementsByTagNameNS(W, "style")) {
    const id = (style.getAttributeNS(W, "styleId") ?? style.getAttribute("w:styleId") ?? "").toLowerCase();
    if (!id) continue;
    const pPr = child(style, W, "pPr");
    const numPr = pPr && child(pPr, W, "numPr");
    if (!numPr) continue;
    const numIdEl = child(numPr, W, "numId");
    const ilvlEl = child(numPr, W, "ilvl");
    if (!numIdEl) continue;
    const numId = attr(numIdEl, "val");
    if (!numId || numId === "0") continue;
    map.set(id, { numId, ilvl: ilvlEl ? parseInt(attr(ilvlEl, "val"), 10) || 0 : 0 });
  }
  return map;
}

// word/comments.xml: numeric id -> {author, date, text}
function parseComments(files) {
  const map = new Map();
  const text = decodePart(files, "word/comments.xml");
  if (!text) return map;
  let doc;
  try { doc = parseXml(text); } catch { return map; }
  for (const c of doc.getElementsByTagNameNS(W, "comment")) {
    const id = c.getAttributeNS(W, "id") ?? c.getAttribute("w:id");
    if (id == null) continue;
    const author = c.getAttributeNS(W, "author") ?? c.getAttribute("w:author") ?? "";
    const date = c.getAttributeNS(W, "date") ?? c.getAttribute("w:date") ?? "";
    const paras = children(c, W, "p").map((p) => p.textContent);
    map.set(id, { author, date, text: paras.join("\n").trim() });
  }
  return map;
}

function resolveTarget(target) {
  // rel targets are relative to word/
  if (target.startsWith("/")) return target.slice(1);
  const parts = ("word/" + target).split("/");
  const out = [];
  for (const p of parts) {
    if (p === "..") out.pop();
    else if (p !== ".") out.push(p);
  }
  return out.join("/");
}

function imageFromDrawing(node, ctx) {
  const blip = findDesc(node, "blip");
  if (!blip) return "";
  const embed = blip.getAttributeNS(R, "embed") || blip.getAttribute("r:embed");
  const rel = embed && ctx.rels.get(embed);
  if (!rel) return "";
  const partName = resolveTarget(rel.target);
  const data = ctx.files.get(partName);
  if (!data) return "";
  const ext = (partName.split(".").pop() || "").toLowerCase();
  const mime = IMG_MIME[ext];
  if (!mime) return ""; // emf/wmf etc. — browser can't render
  let style = "";
  const extent = findDesc(node, "extent");
  if (extent) {
    const cx = parseInt(extent.getAttribute("cx"), 10);
    const cy = parseInt(extent.getAttribute("cy"), 10);
    if (cx > 0 && cy > 0) {
      style = ` width="${Math.round(cx / EMU_PER_PX)}" height="${Math.round(cy / EMU_PER_PX)}"`;
    }
  }
  return `<img src="${bytesToDataUrl(data, mime)}"${style} alt="">`;
}

function runToHtml(r, ctx) {
  const rPr = child(r, W, "rPr");
  let out = "";
  let pageBreak = false;
  const segs = [];
  for (const c of r.children) {
    if (c.namespaceURI !== W && c.localName !== "drawing") continue;
    if (c.localName === "t") segs.push({ t: c.textContent });
    else if (c.localName === "delText") segs.push({ t: c.textContent });
    else if (c.localName === "tab") segs.push({ t: "\t" });
    else if (c.localName === "br") {
      const type = attr(c, "type");
      if (type === "page") pageBreak = true;
      else segs.push({ br: true });
    }
    else if (c.localName === "drawing") segs.push({ raw: imageFromDrawing(c, ctx) });
    else if (c.localName === "noBreakHyphen") segs.push({ t: "‑" });
  }
  let open = "", close = "";
  if (rPr) {
    if (boolProp(child(rPr, W, "b"))) { open += "<b>"; close = "</b>" + close; }
    if (boolProp(child(rPr, W, "i"))) { open += "<i>"; close = "</i>" + close; }
    const u = child(rPr, W, "u");
    if (u && attr(u, "val") !== "none") { open += "<u>"; close = "</u>" + close; }
    if (boolProp(child(rPr, W, "strike"))) { open += "<s>"; close = "</s>" + close; }
    const va = child(rPr, W, "vertAlign");
    if (va) {
      const v = attr(va, "val");
      if (v === "superscript") { open += "<sup>"; close = "</sup>" + close; }
      else if (v === "subscript") { open += "<sub>"; close = "</sub>" + close; }
    }
    const styles = [];
    const color = child(rPr, W, "color");
    if (color) {
      const v = attr(color, "val");
      if (v && v !== "auto") styles.push("color:#" + v.toLowerCase());
    }
    const hl = child(rPr, W, "highlight");
    if (hl) {
      const v = attr(hl, "val");
      if (v && v !== "none") styles.push("background-color:" + (HIGHLIGHT_COLORS[v] || v));
    } else {
      const shd = child(rPr, W, "shd");
      if (shd) {
        const fill = shd.getAttributeNS(W, "fill") ?? shd.getAttribute("w:fill");
        if (fill && fill !== "auto") styles.push("background-color:#" + fill.toLowerCase());
      }
    }
    const sz = child(rPr, W, "sz");
    if (sz) {
      const v = parseInt(attr(sz, "val"), 10);
      if (v) styles.push("font-size:" + v / 2 + "pt");
    }
    const rFonts = child(rPr, W, "rFonts");
    if (rFonts) {
      const f = rFonts.getAttributeNS(W, "ascii") ?? rFonts.getAttribute("w:ascii");
      if (f) styles.push("font-family:" + f);
    }
    if (styles.length) { open += `<span style="${styles.join(";")}">`; close = "</span>" + close; }
  }
  for (const seg of segs) {
    if (seg.raw !== undefined) { out += seg.raw; continue; }
    if (seg.br) { out += "<br>"; continue; }
    out += open + escHtml(seg.t).replace(/\t/g, "&nbsp;&nbsp;&nbsp;&nbsp;") + close;
  }
  return { html: out, pageBreak };
}

// Processes inline-level children of a paragraph (runs, hyperlinks, tracked
// changes, comment range markers) into HTML.
function inlineToHtml(parent, ctx) {
  let html = "";
  let pageBreak = false;
  for (const node of parent.children) {
    if (node.namespaceURI !== W) continue;
    if (node.localName === "r") {
      const r = runToHtml(node, ctx);
      html += r.html;
      pageBreak = pageBreak || r.pageBreak;
    } else if (node.localName === "hyperlink") {
      const id = node.getAttributeNS(R, "id") || node.getAttribute("r:id");
      const rel = id && ctx.rels.get(id);
      const href = rel && rel.mode === "External" ? rel.target : null;
      const inner = inlineToHtml(node, ctx);
      pageBreak = pageBreak || inner.pageBreak;
      html += href ? `<a href="${escHtml(href)}">${inner.html}</a>` : inner.html;
    } else if (node.localName === "ins" || node.localName === "del") {
      // tracked changes -> <ins>/<del> wrappers the editor's review UI understands
      const author = node.getAttributeNS(W, "author") ?? node.getAttribute("w:author") ?? "";
      const date = node.getAttributeNS(W, "date") ?? node.getAttribute("w:date") ?? "";
      const inner = inlineToHtml(node, ctx);
      pageBreak = pageBreak || inner.pageBreak;
      const tag = node.localName;
      html += `<${tag} class="tc-${tag}" data-author="${escHtml(author)}" data-ts="${escHtml(date)}">${inner.html}</${tag}>`;
    } else if (node.localName === "commentRangeStart") {
      const id = node.getAttributeNS(W, "id") ?? node.getAttribute("w:id");
      if (id != null && ctx.comments && ctx.comments.has(id)) {
        // if the range crosses paragraphs the HTML parser auto-closes the span
        // at the paragraph end; the anchor then covers the first portion
        html += `<span class="comment-ref" data-cid="c${id}">`;
      }
    } else if (node.localName === "commentRangeEnd") {
      const id = node.getAttributeNS(W, "id") ?? node.getAttribute("w:id");
      if (id != null && ctx.comments && ctx.comments.has(id)) html += `</span>`;
    } else if (node.localName === "smartTag" || node.localName === "sdt" || node.localName === "sdtContent") {
      // unwrap containers we don't model (content controls)
      const target = node.localName === "sdt" ? (child(node, W, "sdtContent") || node) : node;
      const inner = inlineToHtml(target, ctx);
      html += inner.html;
      pageBreak = pageBreak || inner.pageBreak;
    }
  }
  return { html, pageBreak };
}

function paragraphStyleInfo(p, ctx) {
  const pPr = child(p, W, "pPr");
  const info = { tag: "p", styles: [], numId: null, ilvl: 0, pageBreakBefore: false, listTag: null };
  if (!pPr) return info;
  const pStyle = child(pPr, W, "pStyle");
  if (pStyle) {
    const v = (attr(pStyle, "val") || "").toLowerCase();
    const m = v.match(/heading(\d)/);
    if (m && +m[1] >= 1 && +m[1] <= 6) info.tag = "h" + m[1];
    else if (v === "title") info.tag = "h1";
    else if (v === "quote" || v === "intensequote") info.tag = "blockquote";
    else {
      // style-based numbering: explicit style->numPr map, else Word's built-in
      // ListBullet / ListNumber style-name conventions
      const styleNum = ctx && ctx.styleNumbering && ctx.styleNumbering.get(v);
      if (styleNum) {
        info.numId = styleNum.numId;
        info.ilvl = styleNum.ilvl;
      } else {
        const lm = v.match(/^list(bullet|number)(\d)?$/);
        if (lm) {
          info.listTag = lm[1] === "bullet" ? "ul" : "ol";
          info.ilvl = lm[2] ? parseInt(lm[2], 10) - 1 : 0;
        }
      }
    }
  }
  const jc = child(pPr, W, "jc");
  if (jc) {
    const v = attr(jc, "val");
    if (v === "both" || v === "distribute") info.styles.push("text-align:justify");
    else if (v === "center" || v === "right" || v === "left") info.styles.push("text-align:" + v);
    else if (v === "start") info.styles.push("text-align:left");
    else if (v === "end") info.styles.push("text-align:right");
  }
  const ind = child(pPr, W, "ind");
  if (ind) {
    const left = parseInt(ind.getAttributeNS(W, "left") ?? ind.getAttribute("w:left"), 10);
    const first = parseInt(ind.getAttributeNS(W, "firstLine") ?? ind.getAttribute("w:firstLine"), 10);
    const hang = parseInt(ind.getAttributeNS(W, "hanging") ?? ind.getAttribute("w:hanging"), 10);
    if (left) info.styles.push("margin-left:" + Math.round(left / TW_PER_PX) + "px");
    if (first) info.styles.push("text-indent:" + Math.round(first / TW_PER_PX) + "px");
    else if (hang) info.styles.push("text-indent:-" + Math.round(hang / TW_PER_PX) + "px");
  }
  const spacing = child(pPr, W, "spacing");
  if (spacing) {
    const line = parseInt(spacing.getAttributeNS(W, "line") ?? spacing.getAttribute("w:line"), 10);
    const rule = spacing.getAttributeNS(W, "lineRule") ?? spacing.getAttribute("w:lineRule");
    if (line && (!rule || rule === "auto")) {
      const lh = Math.round((line / 240) * 100) / 100;
      if (lh !== 1) info.styles.push("line-height:" + lh);
    }
  }
  if (boolProp(child(pPr, W, "pageBreakBefore"))) info.pageBreakBefore = true;
  const numPr = child(pPr, W, "numPr");
  if (numPr) {
    const numIdEl = child(numPr, W, "numId");
    const ilvlEl = child(numPr, W, "ilvl");
    info.numId = numIdEl ? attr(numIdEl, "val") : info.numId;
    info.ilvl = ilvlEl ? parseInt(attr(ilvlEl, "val"), 10) || 0 : info.ilvl;
    if (info.numId === "0") info.numId = null; // numId 0 = "no numbering"
  }
  return info;
}

function tableToHtml(tbl, ctx) {
  // Build a model first so vMerge continuation cells can extend rowspans.
  const rows = [];
  const merges = new Map(); // gridCol -> model cell currently spanning down
  for (const tr of children(tbl, W, "tr")) {
    const row = [];
    let gridCol = 0;
    for (const tc of children(tr, W, "tc")) {
      const tcPr = child(tc, W, "tcPr");
      let colspan = 1, vMerge = null, shd = null, widthCss = null;
      if (tcPr) {
        const gs = child(tcPr, W, "gridSpan");
        if (gs) colspan = parseInt(attr(gs, "val"), 10) || 1;
        const vm = child(tcPr, W, "vMerge");
        if (vm) vMerge = attr(vm, "val") || "continue";
        const sh = child(tcPr, W, "shd");
        if (sh) {
          const fill = sh.getAttributeNS(W, "fill") ?? sh.getAttribute("w:fill");
          if (fill && fill !== "auto") shd = "#" + fill.toLowerCase();
        }
        const tcW = child(tcPr, W, "tcW");
        if (tcW) {
          const type = tcW.getAttributeNS(W, "type") ?? tcW.getAttribute("w:type");
          const wv = parseInt(tcW.getAttributeNS(W, "w") ?? tcW.getAttribute("w:w"), 10);
          if (wv > 0 && type === "dxa") widthCss = Math.round(wv / TW_PER_PX) + "px";
          else if (wv > 0 && type === "pct") widthCss = Math.round(wv / 50) + "%";
        }
      }
      // skip grid columns still covered by an active rowspan from above
      while (merges.has(gridCol) && vMerge !== "continue") gridCol += merges.get(gridCol).colspan;
      if (vMerge === "continue") {
        const origin = merges.get(gridCol);
        if (origin) { origin.rowspan++; gridCol += origin.colspan; continue; }
        vMerge = null; // continuation without a restart above: treat as normal
      }
      let inner = "";
      for (const block of tc.children) {
        if (block.namespaceURI !== W) continue;
        if (block.localName === "p") inner += renderParagraph(block, ctx);
        else if (block.localName === "tbl") inner += tableToHtml(block, ctx);
      }
      const cell = { colspan, rowspan: 1, shd, widthCss, inner: inner || "<p><br></p>" };
      row.push(cell);
      if (vMerge === "restart") merges.set(gridCol, cell);
      else merges.delete(gridCol);
      gridCol += colspan;
    }
    rows.push(row);
  }
  let html = '<table><tbody>';
  for (const row of rows) {
    html += "<tr>";
    for (const c of row) {
      const attrs = [];
      if (c.colspan > 1) attrs.push(`colspan="${c.colspan}"`);
      if (c.rowspan > 1) attrs.push(`rowspan="${c.rowspan}"`);
      const st = [];
      if (c.shd) st.push("background-color:" + c.shd);
      if (c.widthCss) st.push("width:" + c.widthCss);
      if (st.length) attrs.push(`style="${st.join(";")}"`);
      html += `<td ${attrs.join(" ")}>${c.inner}</td>`;
    }
    html += "</tr>";
  }
  return html + "</tbody></table>";
}

function renderParagraph(p, ctx) {
  const info = paragraphStyleInfo(p, ctx);
  const inner = inlineToHtml(p, ctx);
  const attrStr = info.styles.length ? ` style="${info.styles.join(";")}"` : "";
  let html = "";
  if (info.pageBreakBefore || inner.pageBreak) html += `<p class="page-break"><br></p>`;
  html += `<${info.tag}${attrStr}>${inner.html || "<br>"}</${info.tag}>`;
  return html;
}

function parseSectPr(sectPr) {
  const setup = JSON.parse(JSON.stringify(DEFAULT_PAGE_SETUP));
  const pgSz = child(sectPr, W, "pgSz");
  if (pgSz) {
    let w = parseInt(pgSz.getAttributeNS(W, "w") ?? pgSz.getAttribute("w:w"), 10) || 12240;
    let h = parseInt(pgSz.getAttributeNS(W, "h") ?? pgSz.getAttribute("w:h"), 10) || 15840;
    const orient = pgSz.getAttributeNS(W, "orient") ?? pgSz.getAttribute("w:orient");
    setup.orientation = orient === "landscape" ? "landscape" : "portrait";
    if (setup.orientation === "landscape") [w, h] = [h, w];
    setup.size = "Letter";
    for (const [name, dim] of Object.entries(PAGE_SIZES)) {
      if (Math.abs(dim.w - w) < 30 && Math.abs(dim.h - h) < 30) { setup.size = name; break; }
    }
  }
  const pgMar = child(sectPr, W, "pgMar");
  if (pgMar) {
    for (const side of ["top", "right", "bottom", "left"]) {
      const v = parseInt(pgMar.getAttributeNS(W, side) ?? pgMar.getAttribute("w:" + side), 10);
      if (v >= 0) setup.margins[side] = Math.round((v / 1440) * 100) / 100;
    }
  }
  return setup;
}

export async function importDocx(fileOrBuffer) {
  const buf = fileOrBuffer.arrayBuffer ? await fileOrBuffer.arrayBuffer() : fileOrBuffer;
  const files = await unzip(buf);
  const docXmlText = decodePart(files, "word/document.xml");
  if (!docXmlText) throw new Error("Not a valid .docx (no word/document.xml)");
  const ctx = {
    files,
    rels: parseRels(files, "word/_rels/document.xml.rels"),
    numFmt: parseNumbering(files),
    styleNumbering: parseStylesNumbering(files),
    comments: parseComments(files),
  };
  const doc = parseXml(docXmlText);
  const body = doc.getElementsByTagNameNS(W, "body")[0];
  if (!body) throw new Error("Malformed document.xml");

  // First pass: flat list of blocks, list items tagged with (numId, ilvl).
  const items = [];
  let pageSetup = null;
  for (const node of body.children) {
    if (node.namespaceURI !== W) continue;
    if (node.localName === "p") {
      const info = paragraphStyleInfo(node, ctx);
      if (info.numId != null && ctx.numFmt.has(info.numId)) {
        const fmt = (ctx.numFmt.get(info.numId) || {})[String(info.ilvl)] || "decimal";
        const inner = inlineToHtml(node, ctx);
        items.push({ li: true, ilvl: Math.min(info.ilvl, 8), tag: fmt === "bullet" ? "ul" : "ol", html: inner.html || "<br>" });
      } else if (info.listTag) {
        const inner = inlineToHtml(node, ctx);
        items.push({ li: true, ilvl: Math.min(info.ilvl, 8), tag: info.listTag, html: inner.html || "<br>" });
      } else {
        items.push({ html: renderParagraph(node, ctx) });
      }
    } else if (node.localName === "tbl") {
      items.push({ html: tableToHtml(node, ctx) });
    } else if (node.localName === "sectPr") {
      pageSetup = parseSectPr(node);
    }
  }

  // Second pass: assemble nested lists.
  let html = "";
  const stack = [];
  const closeOne = () => { html += `</${stack.pop()}>`; };
  for (const item of items) {
    if (!item.li) {
      while (stack.length) closeOne();
      html += item.html;
      continue;
    }
    while (stack.length > item.ilvl + 1) closeOne();
    if (stack.length === item.ilvl + 1 && stack[stack.length - 1] !== item.tag) closeOne();
    while (stack.length < item.ilvl + 1) { html += `<${item.tag}>`; stack.push(item.tag); }
    html += `<li>${item.html}</li>`;
  }
  while (stack.length) closeOne();

  // document title from core.xml
  let title = null;
  const coreText = decodePart(files, "docProps/core.xml");
  if (coreText) {
    try {
      const core = parseXml(coreText);
      const t = core.getElementsByTagNameNS("http://purl.org/dc/elements/1.1/", "title")[0];
      if (t && t.textContent.trim()) title = t.textContent.trim();
    } catch {}
  }

  // comments -> the editor's client-side format
  const comments = [...ctx.comments.entries()].map(([id, c]) => ({
    id: "c" + id,
    author: c.author || "Unknown",
    text: c.text || "",
    createdAt: c.date ? (Date.parse(c.date) || Date.now()) : Date.now(),
    resolved: false,
    replies: [],
  }));

  return { html, pageSetup: pageSetup || { ...DEFAULT_PAGE_SETUP }, title, comments };
}

// Back-compat wrapper.
export async function readDocxHtml(file) {
  return (await importDocx(file)).html;
}

// ============================================================
// EXPORT: HTML -> .docx package
// ============================================================

function parseInlineStyle(style) {
  const o = {};
  if (!style) return o;
  for (const part of style.split(";")) {
    const i = part.indexOf(":");
    if (i < 0) continue;
    const k = part.slice(0, i).trim().toLowerCase();
    const v = part.slice(i + 1).trim();
    if (k) o[k] = v;
  }
  return o;
}
const NAMED_COLORS = { black:"000000", white:"ffffff", red:"ff0000", blue:"0000ff", green:"008000", yellow:"ffff00", cyan:"00ffff", magenta:"ff00ff", gray:"808080", grey:"808080", silver:"c0c0c0", maroon:"800000", olive:"808000", purple:"800080", teal:"008080", navy:"000080", orange:"ffa500", lime:"00ff00", aqua:"00ffff", fuchsia:"ff00ff" };
function cssColorToHex(v) {
  if (!v) return null;
  v = v.trim().toLowerCase();
  if (v === "transparent" || v === "inherit" || v === "initial") return null;
  if (v.startsWith("#")) {
    let h = v.slice(1);
    if (h.length === 3) h = h.split("").map(c => c + c).join("");
    if (h.length === 8) h = h.slice(0, 6);
    return h.length === 6 ? h : null;
  }
  const rgb = v.match(/rgba?\(([^)]+)\)/);
  if (rgb) {
    const parts = rgb[1].split(",").map(s => parseFloat(s.trim()));
    if (parts.length === 4 && parts[3] === 0) return null;
    if (parts.length >= 3) return parts.slice(0, 3).map(n => Math.max(0, Math.min(255, Math.round(n))).toString(16).padStart(2, "0")).join("");
  }
  return NAMED_COLORS[v] || null;
}
function cssSizeToHalfPoints(v) {
  if (!v) return null;
  let m = v.match(/([\d.]+)\s*pt/);
  if (m) return Math.round(parseFloat(m[1]) * 2);
  m = v.match(/([\d.]+)\s*px/);
  if (m) return Math.round(parseFloat(m[1]) * 1.5);
  m = v.match(/([\d.]+)\s*em/);
  if (m) return Math.round(parseFloat(m[1]) * 22);
  return null;
}
function cssLenToTwips(v) {
  if (!v) return null;
  let m = String(v).match(/(-?[\d.]+)\s*px/);
  if (m) return Math.round(parseFloat(m[1]) * TW_PER_PX);
  m = String(v).match(/(-?[\d.]+)\s*pt/);
  if (m) return Math.round(parseFloat(m[1]) * 20);
  m = String(v).match(/(-?[\d.]+)\s*in/);
  if (m) return Math.round(parseFloat(m[1]) * 1440);
  return null;
}

// ---- runs ----
function spanProps(el, props) {
  const np = { ...props };
  const st = parseInlineStyle(el.getAttribute("style") || "");
  if (st["color"]) np.color = cssColorToHex(st["color"]) || np.color;
  if (st["background-color"]) {
    const h = cssColorToHex(st["background-color"]);
    if (h) np.highlight = h;
  }
  const sz = cssSizeToHalfPoints(st["font-size"]);
  if (sz) np.sz = sz;
  if (st["font-family"]) np.font = st["font-family"].split(",")[0].replace(/['"]/g, "").trim();
  if (st["font-weight"] === "bold" || parseInt(st["font-weight"], 10) >= 600) np.b = true;
  if (st["font-weight"] === "normal") np.b = false;
  if (st["font-style"] === "italic") np.i = true;
  const td = st["text-decoration"] || st["text-decoration-line"] || "";
  if (td.includes("underline")) np.u = true;
  if (td.includes("line-through")) np.s = true;
  if (st["vertical-align"] === "super") np.sup = true;
  if (st["vertical-align"] === "sub") np.sub = true;
  return np;
}

function collectRuns(node, props, runs) {
  for (const kid of node.childNodes) {
    if (kid.nodeType === 3) {
      const t = kid.textContent;
      if (t !== "") runs.push({ ...props, text: t });
      continue;
    }
    if (kid.nodeType !== 1) continue;
    const el = kid;
    const tag = el.tagName.toLowerCase();
    let np = { ...props };
    if (tag === "b" || tag === "strong") np.b = true;
    else if (tag === "i" || tag === "em") np.i = true;
    else if (tag === "ins") {
      // tracked insertion (review) vs plain <ins> (underline)
      if (el.classList.contains("tc-ins")) {
        np.ins = { author: el.getAttribute("data-author") || "Author", date: el.getAttribute("data-ts") || "" };
      } else np.u = true;
    }
    else if (tag === "u") np.u = true;
    else if (tag === "del") {
      if (el.classList.contains("tc-del")) {
        np.del = { author: el.getAttribute("data-author") || "Author", date: el.getAttribute("data-ts") || "" };
      } else np.s = true;
    }
    else if (tag === "s" || tag === "strike") np.s = true;
    else if (tag === "sub") np.sub = true;
    else if (tag === "sup") np.sup = true;
    else if (tag === "mark") np.highlight = np.highlight || "ffff00";
    else if (tag === "code" || tag === "kbd" || tag === "samp" || tag === "tt") np.font = "Courier New";
    else if (tag === "br") { runs.push({ ...props, br: true, pb: el.classList.contains("pb") }); continue; }
    else if (tag === "img") { runs.push({ ...props, img: el }); continue; }
    else if (tag === "a") {
      const href = el.getAttribute("href");
      if (href && /^(https?:|mailto:)/i.test(href)) np.link = href;
    }
    else if (tag === "span" || tag === "font") {
      if (el.classList && el.classList.contains("comment-ref")) {
        const cid = el.getAttribute("data-cid");
        if (cid) np.cmt = cid;
      }
      const face = el.getAttribute("face");
      if (face) np.font = face;
      const size = el.getAttribute("size");
      if (size) np.sz = (parseInt(size, 10) * 2 + 12) || np.sz;
      const fcolor = el.getAttribute("color");
      if (fcolor) np.color = cssColorToHex(fcolor) || np.color;
    }
    if (el.getAttribute && el.getAttribute("style")) np = spanProps(el, np);
    collectRuns(el, np, runs);
  }
}

function runPropsXml(run, opts = {}) {
  const rPr = [];
  if (opts.hyperlink) rPr.push(`<w:rStyle w:val="Hyperlink"/>`);
  if (run.font) rPr.push(`<w:rFonts w:ascii="${escXml(run.font)}" w:hAnsi="${escXml(run.font)}"/>`);
  if (run.b) rPr.push(`<w:b/>`);
  if (run.i) rPr.push(`<w:i/>`);
  if (run.s) rPr.push(`<w:strike/>`);
  if (run.color) rPr.push(`<w:color w:val="${escXml(run.color)}"/>`);
  if (run.sz) rPr.push(`<w:sz w:val="${run.sz}"/><w:szCs w:val="${run.sz}"/>`);
  if (run.highlight) {
    const named = HEX_TO_HIGHLIGHT[run.highlight.toLowerCase()];
    if (named) rPr.push(`<w:highlight w:val="${named}"/>`);
    else rPr.push(`<w:shd w:val="clear" w:color="auto" w:fill="${escXml(run.highlight)}"/>`);
  }
  if (run.u) rPr.push(`<w:u w:val="single"/>`);
  if (run.sub) rPr.push(`<w:vertAlign w:val="subscript"/>`);
  if (run.sup) rPr.push(`<w:vertAlign w:val="superscript"/>`);
  return rPr.length ? `<w:rPr>${rPr.join("")}</w:rPr>` : "";
}

function imageRunXml(run, ctx) {
  const media = ctx.images.get(run.img);
  if (!media) return "";
  const cx = Math.max(1, Math.round(media.w * EMU_PER_PX));
  const cy = Math.max(1, Math.round(media.h * EMU_PER_PX));
  const id = media.docPrId;
  return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:docPr id="${id}" name="Picture ${id}"/>` +
    `<a:graphic xmlns:a="${A}"><a:graphicData uri="${PIC}">` +
    `<pic:pic xmlns:pic="${PIC}">` +
    `<pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${media.relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}

function runToXml(run, ctx, opts = {}) {
  if (run.img) return imageRunXml(run, ctx);
  if (run.br) return run.pb ? `<w:r><w:br w:type="page"/></w:r>` : `<w:r><w:br/></w:r>`;
  // deleted runs must use w:delText instead of w:t
  const tTag = run.del ? "w:delText" : "w:t";
  return `<w:r>${runPropsXml(run, opts)}<${tTag} xml:space="preserve">${escXml(run.text)}</${tTag}></w:r>`;
}

// Groups consecutive runs sharing the same link into w:hyperlink wrappers.
function linkGroupToXml(runs, ctx) {
  let out = "";
  let i = 0;
  while (i < runs.length) {
    const link = runs[i].link;
    if (!link) { out += runToXml(runs[i], ctx); i++; continue; }
    let j = i;
    let inner = "";
    while (j < runs.length && runs[j].link === link) {
      inner += runToXml(runs[j], ctx, { hyperlink: !runs[j].img });
      j++;
    }
    let relId = ctx.hrefRels.get(link);
    if (!relId) {
      relId = "rId" + ctx.nextRelId++;
      ctx.hrefRels.set(link, relId);
      ctx.rels.push(`<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${escXml(link)}" TargetMode="External"/>`);
    }
    out += `<w:hyperlink r:id="${relId}">${inner}</w:hyperlink>`;
    i = j;
  }
  return out;
}

// Full inline serialization: comment range markers (milestones) outermost,
// tracked-change wrappers (w:ins / w:del) next, hyperlinks innermost.
function runsToXml(runs, ctx) {
  let out = "";
  let openCmt = null;
  const closeCmt = () => {
    const id = ctx.commentId(openCmt);
    out += `<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`;
    openCmt = null;
  };
  let i = 0;
  while (i < runs.length) {
    const r = runs[i];
    const cmt = r.cmt && ctx.commentId(r.cmt) != null ? r.cmt : null;
    if (cmt !== openCmt) {
      if (openCmt) closeCmt();
      if (cmt) { out += `<w:commentRangeStart w:id="${ctx.commentId(cmt)}"/>`; openCmt = cmt; }
    }
    const kind = r.ins ? "ins" : r.del ? "del" : "";
    const kindKey = kind ? JSON.stringify(r[kind]) : "";
    let j = i + 1;
    while (j < runs.length) {
      const rj = runs[j];
      const cmtJ = rj.cmt && ctx.commentId(rj.cmt) != null ? rj.cmt : null;
      if (cmtJ !== cmt) break;
      const kindJ = rj.ins ? "ins" : rj.del ? "del" : "";
      if (kindJ !== kind) break;
      if (kind && JSON.stringify(rj[kind]) !== kindKey) break;
      j++;
    }
    const inner = linkGroupToXml(runs.slice(i, j), ctx);
    if (kind) {
      const meta = r[kind];
      out += `<w:${kind} w:id="${ctx.revId++}" w:author="${escXml(meta.author || "Author")}" w:date="${tsIso(meta.date)}">${inner}</w:${kind}>`;
    } else {
      out += inner;
    }
    i = j;
  }
  if (openCmt) closeCmt();
  return out;
}

// ---- paragraphs / blocks ----
function paragraphPropsXml(el, extra = []) {
  const pPr = [...extra];
  const st = el.style || {};
  const align = st.textAlign || el.getAttribute("align");
  if (align === "left" || align === "start") pPr.push(`<w:jc w:val="left"/>`);
  else if (align === "center") pPr.push(`<w:jc w:val="center"/>`);
  else if (align === "right" || align === "end") pPr.push(`<w:jc w:val="right"/>`);
  else if (align === "justify") pPr.push(`<w:jc w:val="both"/>`);
  const indParts = [];
  const ml = cssLenToTwips(st.marginLeft);
  if (ml) indParts.push(`w:left="${ml}"`);
  const ti = cssLenToTwips(st.textIndent);
  if (ti > 0) indParts.push(`w:firstLine="${ti}"`);
  else if (ti < 0) indParts.push(`w:hanging="${-ti}"`);
  if (indParts.length) pPr.push(`<w:ind ${indParts.join(" ")}/>`);
  const lh = parseFloat(st.lineHeight);
  if (lh && !String(st.lineHeight).match(/px|pt/) && lh > 0 && lh < 10) {
    pPr.push(`<w:spacing w:line="${Math.round(lh * 240)}" w:lineRule="auto"/>`);
  }
  return pPr;
}

// contentEditable represents an empty paragraph as <p><br></p>; that trailing
// br is a placeholder, not a real line break.
function stripPlaceholderBr(runs) {
  if (runs.length && runs[runs.length - 1].br && !runs[runs.length - 1].pb) {
    const hasText = runs.some((r) => (r.text && r.text.trim()) || r.img);
    if (!hasText) runs.pop();
  }
  return runs;
}

function paragraphToXml(el, ctx, extraPPr = []) {
  const pPr = paragraphPropsXml(el, extraPPr);
  const runs = [];
  collectRuns(el, {}, runs);
  stripPlaceholderBr(runs);
  const pPrXml = pPr.length ? `<w:pPr>${pPr.join("")}</w:pPr>` : "";
  return `<w:p>${pPrXml}${runsToXml(runs, ctx)}</w:p>`;
}

function listToXml(listEl, ctx, ilvl, numId) {
  const tag = listEl.tagName.toLowerCase();
  if (numId == null) {
    if (tag === "ol") { numId = ctx.nextNumId++; ctx.nums.push({ numId, abstract: 1 }); }
    else numId = 1; // shared bullet numbering
  }
  let out = "";
  for (const li of listEl.children) {
    const t = li.tagName.toLowerCase();
    if (t === "ul" || t === "ol") { out += listToXml(li, ctx, Math.min(ilvl + 1, 8), t === tag ? numId : null); continue; }
    if (t !== "li") continue;
    // split direct inline content from nested lists
    const nested = [];
    const clone = li.cloneNode(true);
    for (const sub of [...clone.children]) {
      const st = sub.tagName.toLowerCase();
      if (st === "ul" || st === "ol") { clone.removeChild(sub); }
    }
    const runs = [];
    collectRuns(clone, {}, runs);
    stripPlaceholderBr(runs);
    const numPr = `<w:numPr><w:ilvl w:val="${Math.min(ilvl, 8)}"/><w:numId w:val="${numId}"/></w:numPr>`;
    const pPr = paragraphPropsXml(li, [`<w:pStyle w:val="ListParagraph"/>`, numPr]);
    out += `<w:p><w:pPr>${pPr.join("")}</w:pPr>${runsToXml(runs, ctx)}</w:p>`;
    for (const sub of li.children) {
      const st = sub.tagName.toLowerCase();
      if (st === "ul" || st === "ol") nested.push(sub);
    }
    for (const sub of nested) out += listToXml(sub, ctx, Math.min(ilvl + 1, 8), sub.tagName.toLowerCase() === tag ? numId : null);
  }
  return out;
}

function tableToXml(tableEl, ctx) {
  // Build grid model from the HTML table, honoring colspan/rowspan.
  const trs = [...tableEl.querySelectorAll(":scope > tr, :scope > thead > tr, :scope > tbody > tr, :scope > tfoot > tr")];
  if (!trs.length) return "";
  const grid = []; // grid[row][col] = {cell, origin:bool} | undefined
  let nCols = 0;
  trs.forEach((tr, r) => {
    grid[r] = grid[r] || [];
    let c = 0;
    for (const cell of tr.children) {
      const t = cell.tagName.toLowerCase();
      if (t !== "td" && t !== "th") continue;
      while (grid[r][c] !== undefined) c++;
      const colspan = Math.max(1, parseInt(cell.getAttribute("colspan"), 10) || 1);
      const rowspan = Math.max(1, parseInt(cell.getAttribute("rowspan"), 10) || 1);
      for (let dr = 0; dr < rowspan; dr++) {
        grid[r + dr] = grid[r + dr] || [];
        for (let dc = 0; dc < colspan; dc++) {
          grid[r + dr][c + dc] = { cell, origin: dr === 0 && dc === 0, top: dr === 0, colspan };
        }
      }
      c += colspan;
      nCols = Math.max(nCols, c);
    }
    nCols = Math.max(nCols, grid[r].length);
  });
  if (!nCols) return "";
  const colW = Math.floor(9360 / nCols);
  let xml = `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>` +
    `<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>` +
    `</w:tblPr><w:tblGrid>${`<w:gridCol w:w="${colW}"/>`.repeat(nCols)}</w:tblGrid>`;
  for (let r = 0; r < grid.length; r++) {
    xml += "<w:tr>";
    let c = 0;
    while (c < nCols) {
      const slot = grid[r][c];
      if (!slot) { // empty filler cell
        xml += `<w:tc><w:tcPr><w:tcW w:w="${colW}" w:type="dxa"/></w:tcPr><w:p/></w:tc>`;
        c++;
        continue;
      }
      const { cell, origin, top, colspan } = slot;
      const tcPr = [`<w:tcW w:w="${colW * colspan}" w:type="dxa"/>`];
      if (colspan > 1) tcPr.push(`<w:gridSpan w:val="${colspan}"/>`);
      const rowspan = Math.max(1, parseInt(cell.getAttribute("rowspan"), 10) || 1);
      if (rowspan > 1) tcPr.push(origin ? `<w:vMerge w:val="restart"/>` : `<w:vMerge/>`);
      const bg = cssColorToHex((cell.style && cell.style.backgroundColor) || cell.getAttribute("bgcolor") || "");
      if (bg && (origin || top)) tcPr.push(`<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>`);
      let content = "";
      if (origin) {
        const isTh = cell.tagName.toLowerCase() === "th";
        const blocks = blockChildren(cell);
        if (blocks.length) content = blocks.map((b) => blockToXml(b, ctx)).join("");
        else if (isTh) {
          const runs = [];
          collectRuns(cell, { b: true }, runs);
          stripPlaceholderBr(runs);
          content = `<w:p>${runsToXml(runs, ctx)}</w:p>`;
        } else {
          content = paragraphToXml(cell, ctx);
        }
        if (!content.endsWith("</w:p>") && !content.endsWith("<w:p/>")) content += "<w:p/>"; // tc must end with a paragraph
      } else {
        content = "<w:p/>";
      }
      xml += `<w:tc><w:tcPr>${tcPr.join("")}</w:tcPr>${content || "<w:p/>"}</w:tc>`;
      c += colspan;
    }
    xml += "</w:tr>";
  }
  return xml + "</w:tbl>";
}

const BLOCK_TAGS = new Set(["p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "ul", "ol", "table", "blockquote", "pre", "hr", "li", "section", "article", "figure"]);
function blockChildren(el) {
  const out = [];
  for (const c of el.children) {
    if (BLOCK_TAGS.has(c.tagName.toLowerCase())) out.push(c);
  }
  // only treat as block container if ALL meaningful content is in blocks
  if (!out.length) return [];
  for (const n of el.childNodes) {
    if (n.nodeType === 3 && n.textContent.trim()) return [];
    if (n.nodeType === 1 && !BLOCK_TAGS.has(n.tagName.toLowerCase()) && n.textContent.trim()) return [];
  }
  return out;
}

function blockToXml(el, ctx) {
  const tag = el.tagName.toLowerCase();
  if (el.classList && el.classList.contains("page-break")) {
    return `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
  }
  if (/^h[1-6]$/.test(tag)) {
    return paragraphToXml(el, ctx, [`<w:pStyle w:val="Heading${tag[1]}"/>`]);
  }
  if (tag === "blockquote") {
    const blocks = blockChildren(el);
    if (blocks.length) return blocks.map((b) => paragraphToXml(b, ctx, [`<w:pStyle w:val="Quote"/>`])).join("");
    return paragraphToXml(el, ctx, [`<w:pStyle w:val="Quote"/>`]);
  }
  if (tag === "ul" || tag === "ol") return listToXml(el, ctx, 0, null);
  if (tag === "table") return tableToXml(el, ctx);
  if (tag === "hr") {
    return `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>`;
  }
  if (tag === "pre") {
    const lines = el.textContent.split("\n");
    return lines.map((ln) =>
      `<w:p><w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">${escXml(ln)}</w:t></w:r></w:p>`
    ).join("");
  }
  // p / div / other containers
  const blocks = blockChildren(el);
  if (blocks.length) return blocks.map((b) => blockToXml(b, ctx)).join("");
  return paragraphToXml(el, ctx);
}

// ---- image collection (async: sizes + bytes) ----
async function collectImages(container, ctx) {
  const imgs = [...container.querySelectorAll("img")];
  let n = 0;
  for (const img of imgs) {
    try {
      const src = img.getAttribute("src") || "";
      let bytes, mime;
      if (src.startsWith("data:")) {
        ({ bytes, mime } = dataUrlToBytes(src));
      } else if (src) {
        const resp = await fetch(src);
        const blob = await resp.blob();
        mime = blob.type || "image/png";
        bytes = new Uint8Array(await blob.arrayBuffer());
      } else continue;
      let ext = Object.entries(IMG_MIME).find(([, m]) => m === mime)?.[0] || "png";
      if (ext === "jpeg") ext = "jpg";
      // dimensions: explicit attrs/styles win, else decode
      let w = parseFloat(img.getAttribute("width")) || parseFloat(img.style.width) || 0;
      let h = parseFloat(img.getAttribute("height")) || parseFloat(img.style.height) || 0;
      if (!w || !h) {
        const dims = await new Promise((resolve) => {
          const probe = new Image();
          probe.onload = () => resolve({ w: probe.naturalWidth, h: probe.naturalHeight });
          probe.onerror = () => resolve({ w: 300, h: 200 });
          probe.src = src;
        });
        if (!w && !h) { w = dims.w; h = dims.h; }
        else if (!h) h = w * (dims.h / dims.w || 0.66);
        else if (!w) w = h * (dims.w / dims.h || 1.5);
      }
      // cap at printable width (~6.5in @ 96dpi)
      const MAXW = 624;
      if (w > MAXW) { h = h * (MAXW / w); w = MAXW; }
      n++;
      const name = `image${n}.${ext}`;
      const relId = "rId" + ctx.nextRelId++;
      ctx.rels.push(`<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${name}"/>`);
      ctx.media.push({ name, bytes, ext });
      ctx.images.set(img, { relId, w: Math.round(w), h: Math.round(h), docPrId: n });
    } catch (e) {
      console.warn("image export skipped:", e);
    }
  }
}

// ---- package parts ----
function contentTypesXml(ctx) {
  const exts = new Set(ctx.media.map((m) => m.ext));
  let defaults = "";
  for (const e of exts) {
    const mime = IMG_MIME[e === "jpg" ? "jpeg" : e] || "image/" + e;
    defaults += `<Default Extension="${e}" ContentType="${mime}"/>`;
  }
  const commentsOverride = ctx.commentIds.size
    ? `<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>`
    : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` + defaults +
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
    `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
    `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>` +
    commentsOverride +
    `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>` +
    `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>` +
    `</Types>`;
}

const PKG_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;

function docRelsXml(ctx) {
  const commentsRel = ctx.commentIds.size
    ? `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>`
    : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>` +
    commentsRel +
    ctx.rels.join("") +
    `</Relationships>`;
}

function commentsXml(ctx) {
  let out = "";
  for (const [cid, num] of ctx.commentIds) {
    const c = ctx.commentMeta.get(cid) || {};
    const paras = [c.text || "", ...((c.replies || []).map((r) => `${r.author || "?"}: ${r.text || ""}`))].filter(Boolean);
    const body = paras.map((t) =>
      t.split("\n").map((ln) => `<w:p><w:r><w:t xml:space="preserve">${escXml(ln)}</w:t></w:r></w:p>`).join("")
    ).join("") || "<w:p/>";
    out += `<w:comment w:id="${num}" w:author="${escXml(c.author || "Unknown")}" w:date="${tsIso(c.createdAt)}" w:initials="${escXml(String(c.author || "?").slice(0, 2).toUpperCase())}">${body}</w:comment>`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:comments xmlns:w="${W}">${out}</w:comments>`;
}

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W}">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:color w:val="2F5496"/><w:sz w:val="40"/><w:szCs w:val="40"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="160" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:color w:val="2F5496"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="160" w:after="80"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:color w:val="2F5496"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="80" w:after="40"/><w:outlineLvl w:val="3"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:i/><w:color w:val="2F5496"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading5"><w:name w:val="heading 5"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="80" w:after="40"/><w:outlineLvl w:val="4"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:color w:val="2F5496"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading6"><w:name w:val="heading 6"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="40" w:after="0"/><w:outlineLvl w:val="5"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:color w:val="595959"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="80"/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="160"/><w:ind w:left="720" w:right="720"/></w:pPr><w:rPr><w:i/><w:color w:val="404040"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>
</w:styles>`;

const BULLET_GLYPHS = ["", "o", ""]; // Symbol bullet, Courier o, Wingdings square
const BULLET_FONTS = ["Symbol", "Courier New", "Wingdings"];
function numberingXml(ctx) {
  let bulletLvls = "", decimalLvls = "";
  for (let l = 0; l < 9; l++) {
    const ind = 720 * (l + 1);
    bulletLvls += `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${BULLET_GLYPHS[l % 3]}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${ind}" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="${BULLET_FONTS[l % 3]}" w:hAnsi="${BULLET_FONTS[l % 3]}" w:hint="default"/></w:rPr></w:lvl>`;
    const fmt = l % 3 === 0 ? "decimal" : l % 3 === 1 ? "lowerLetter" : "lowerRoman";
    decimalLvls += `<w:lvl w:ilvl="${l}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="%${l + 1}."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${ind}" w:hanging="360"/></w:pPr></w:lvl>`;
  }
  let nums = `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`;
  for (const n of ctx.nums) {
    nums += `<w:num w:numId="${n.numId}"><w:abstractNumId w:val="${n.abstract}"/></w:num>`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:numbering xmlns:w="${W}">` +
    `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${bulletLvls}</w:abstractNum>` +
    `<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${decimalLvls}</w:abstractNum>` +
    nums +
    `</w:numbering>`;
}

function coreXml(title) {
  const now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
    `<dc:title>${escXml(title || "")}</dc:title>` +
    `<dc:creator>Word-Compat Editor</dc:creator>` +
    `<cp:lastModifiedBy>Word-Compat Editor</cp:lastModifiedBy>` +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>` +
    `</cp:coreProperties>`;
}
const APP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Word-Compat Editor</Application></Properties>`;

function sectPrXml(setup) {
  const s = setup || DEFAULT_PAGE_SETUP;
  const dim = PAGE_SIZES[s.size] || PAGE_SIZES.Letter;
  let w = dim.w, h = dim.h;
  const landscape = s.orientation === "landscape";
  if (landscape) [w, h] = [h, w];
  const m = s.margins || DEFAULT_PAGE_SETUP.margins;
  const tw = (v) => Math.round((v != null ? v : 1) * 1440);
  return `<w:sectPr><w:pgSz w:w="${w}" w:h="${h}"${landscape ? ' w:orient="landscape"' : ""}/>` +
    `<w:pgMar w:top="${tw(m.top)}" w:right="${tw(m.right)}" w:bottom="${tw(m.bottom)}" w:left="${tw(m.left)}" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;
}

export function htmlToDocumentXml(html, ctx, pageSetup) {
  const container = document.createElement("div");
  container.innerHTML = html || "<p></p>";
  return domToDocumentXml(container, ctx, pageSetup);
}
function domToDocumentXml(container, ctx, pageSetup) {
  let body = "";
  const blocks = blockChildren(container);
  if (blocks.length) {
    for (const el of container.children) body += blockToXml(el, ctx);
  } else if (container.textContent.trim() || container.querySelector("img")) {
    body = paragraphToXml(container, ctx);
  }
  if (!body) body = `<w:p/>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:pic="${PIC}">` +
    `<w:body>${body}${sectPrXml(pageSetup)}</w:body></w:document>`;
}

export async function buildDocxFromHtml(html, opts = {}) {
  const enc = new TextEncoder();
  const ctx = {
    rels: [], media: [], images: new Map(),
    hrefRels: new Map(), nextRelId: 10,
    nums: [], nextNumId: 2,
    revId: 1,
    commentMeta: new Map((opts.comments || []).map((c) => [c.id, c])),
    commentIds: new Map(),
  };
  ctx.commentId = (cid) => {
    if (!ctx.commentMeta.has(cid)) return null;
    if (!ctx.commentIds.has(cid)) ctx.commentIds.set(cid, ctx.commentIds.size);
    return ctx.commentIds.get(cid);
  };
  const container = document.createElement("div");
  container.innerHTML = html || "<p></p>";
  await collectImages(container, ctx);
  const documentXml = domToDocumentXml(container, ctx, opts.pageSetup);
  const files = new Map([
    ["[Content_Types].xml", enc.encode(contentTypesXml(ctx))],
    ["_rels/.rels", enc.encode(PKG_RELS)],
    ["word/document.xml", enc.encode(documentXml)],
    ["word/_rels/document.xml.rels", enc.encode(docRelsXml(ctx))],
    ["word/styles.xml", enc.encode(STYLES)],
    ["word/numbering.xml", enc.encode(numberingXml(ctx))],
    ["docProps/core.xml", enc.encode(coreXml(opts.title))],
    ["docProps/app.xml", enc.encode(APP_XML)],
  ]);
  if (ctx.commentIds.size) files.set("word/comments.xml", enc.encode(commentsXml(ctx)));
  for (const m of ctx.media) files.set("word/media/" + m.name, m.bytes);
  return zip(files);
}

// ============================================================
// Other formats
// ============================================================

export function htmlToPlainText(html) {
  const div = document.createElement("div");
  div.innerHTML = html || "";
  const out = [];
  const walk = (node) => {
    for (const c of node.childNodes) {
      if (c.nodeType === 3) { out.push(c.textContent); continue; }
      if (c.nodeType !== 1) continue;
      const tag = c.tagName.toLowerCase();
      if (tag === "br") { out.push("\n"); continue; }
      if (tag === "del" && c.classList.contains("tc-del")) continue; // deleted (tracked) text
      const isBlock = BLOCK_TAGS.has(tag) || tag === "tr";
      if (tag === "td" || tag === "th") { walk(c); out.push("\t"); continue; }
      walk(c);
      if (isBlock) out.push("\n");
    }
  };
  walk(div);
  return out.join("").replace(/\t\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim() + "\n";
}

export function plainTextToHtml(text) {
  return String(text).split(/\r?\n/).map((ln) => `<p>${escHtml(ln) || "<br>"}</p>`).join("");
}

export function exportStandaloneHtml(html, title) {
  return `<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>${escHtml(title || "Document")}</title>
<style>
body{font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:1.4;max-width:8.5in;margin:24px auto;padding:0 1in;color:#111}
table{border-collapse:collapse;width:100%;margin:8px 0}
td,th{border:1px solid #bbb;padding:4px 8px;vertical-align:top}
img{max-width:100%}
blockquote{border-left:3px solid #2b579a;margin-left:0;padding-left:12px;color:#555;font-style:italic}
h1,h2,h3{color:#2F5496}
ins.tc-ins{background:#e7f7e7;color:#14652f;text-decoration:underline}
del.tc-del{background:#fdeaea;color:#b02a2a;text-decoration:line-through}
span.comment-ref{background:#fff3bf;border-bottom:2px solid #f4a806}
.page-break{page-break-after:always;border:none}
</style></head><body>${html}</body></html>`;
}
