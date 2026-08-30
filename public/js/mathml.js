// mathml.js — dependency-free math equation support: a small LaTeX-ish input
// language, compiled to MathML (which Chrome/Firefox/Safari all render
// natively via MathML Core — no rendering library needed), plus MathML<->OMML
// (Office Math Markup Language, the <m:oMath> vocabulary Word actually saves)
// converters so equations round-trip through real .docx files.
//
// Scope is deliberately a practical subset, not full LaTeX: fractions, roots,
// sub/superscripts, sums/products/integrals with limits, greek letters, common
// operators/relations, and parenthesized groups. Good enough for the equations
// that show up in contracts/reports (interest formulas, simple statistics);
// not a TeX engine.
//
// This module works both in the browser and in Node (via server/docxNode.mjs's
// jsdom shim), exactly like docx.js — it only touches DOM APIs that jsdom
// provides (DOMParser, document.createElementNS is NOT required: everything
// here is string-in/string-out plus DOM *reads* via a parsed XML document).

const MATHML_NS = "http://www.w3.org/1998/Math/MathML";
const M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

// ---------------------------------------------------------------
// LaTeX-subset -> MathML
// ---------------------------------------------------------------

const GREEK = {
  alpha: "α", beta: "β", gamma: "γ", delta: "δ", epsilon: "ε", zeta: "ζ", eta: "η",
  theta: "θ", iota: "ι", kappa: "κ", lambda: "λ", mu: "μ", nu: "ν", xi: "ξ",
  pi: "π", rho: "ρ", sigma: "σ", tau: "τ", upsilon: "υ", phi: "φ", chi: "χ",
  psi: "ψ", omega: "ω",
  Gamma: "Γ", Delta: "Δ", Theta: "Θ", Lambda: "Λ", Xi: "Ξ", Pi: "Π", Sigma: "Σ",
  Upsilon: "Υ", Phi: "Φ", Psi: "Ψ", Omega: "Ω",
};
// single-symbol commands with no arguments
const SYMBOLS = {
  infty: "∞", pm: "±", mp: "∓", times: "×", cdot: "⋅", div: "÷",
  leq: "≤", geq: "≥", neq: "≠", approx: "≈", equiv: "≡", propto: "∝",
  to: "→", rightarrow: "→", leftarrow: "←", leftrightarrow: "↔", Rightarrow: "⇒",
  in: "∈", notin: "∉", subset: "⊂", subseteq: "⊆", cup: "∪", cap: "∩",
  forall: "∀", exists: "∃", partial: "∂", nabla: "∇", cdots: "⋯", ldots: "…",
  degree: "°", angle: "∠", perp: "⊥", parallel: "∥", sim: "∼", ell: "ℓ",
};
const NARY = { sum: "∑", prod: "∏", int: "∫", oint: "∮", coprod: "∐" };
const GREEK_CHARS = new Set(Object.values(GREEK));
const REVERSE_GREEK = Object.fromEntries(Object.entries(GREEK).map(([k, v]) => [v, k]));

class LatexParser {
  constructor(src) {
    this.s = src;
    this.i = 0;
  }
  peek() { return this.s[this.i]; }
  eof() { return this.i >= this.s.length; }
  skipSpace() { while (!this.eof() && /\s/.test(this.peek())) this.i++; }

  // top-level: a sequence of expr terms until eof or a stop char (used for group bodies)
  parseRow(stopChars = "") {
    const items = [];
    for (;;) {
      this.skipSpace();
      if (this.eof() || stopChars.includes(this.peek())) break;
      items.push(this.parseAtomWithScripts());
    }
    return items.length === 1 ? items[0] : { t: "row", items };
  }

  // one atom, then optional ^{}/_{} attached to it
  parseAtomWithScripts() {
    let base = this.parseAtom();
    this.skipSpace();
    let sup = null, sub = null;
    while (this.peek() === "^" || this.peek() === "_") {
      const isSup = this.peek() === "^";
      this.i++;
      const arg = this.parseGroupOrSingle();
      if (isSup) sup = arg; else sub = arg;
      this.skipSpace();
    }
    if (sup && sub) return { t: "subsup", base, sub, sup };
    if (sup) return { t: "sup", base, sup };
    if (sub) return { t: "sub", base, sub };
    return base;
  }

  // {...} group, or a single next atom (no scripts) if no brace follows —
  // e.g. x^2 without braces
  parseGroupOrSingle() {
    this.skipSpace();
    if (this.peek() === "{") {
      this.i++;
      const row = this.parseRow("}");
      this.skipSpace();
      if (this.peek() === "}") this.i++;
      return row;
    }
    return this.parseSingleToken();
  }

  // a single "letter-ish" token: one char, or one \command (without further
  // script/group consumption) — used for bare x^2, a_i
  parseSingleToken() {
    this.skipSpace();
    if (this.peek() === "\\") return this.parseCommand(false);
    const c = this.peek();
    this.i++;
    return this.classify(c);
  }

  parseAtom() {
    this.skipSpace();
    const c = this.peek();
    if (c === undefined) return { t: "row", items: [] };
    if (c === "\\") return this.parseCommand(true);
    if (c === "(" || c === "[" || c === "{") return this.parseDelimited();
    this.i++;
    return this.classify(c);
  }

  // parses a plain ( ... ) or [ ... ] group as a fenced <m:d>/mfenced-like row
  parseDelimited() {
    const open = this.peek();
    const close = open === "(" ? ")" : open === "[" ? "]" : "}";
    this.i++;
    const row = this.parseRow(close);
    this.skipSpace();
    if (this.peek() === close) this.i++;
    return { t: "fenced", open, close, body: row };
  }

  // \left( ... \right) — same as a plain delimited group, auto-sizing is a
  // rendering detail MathML/OMML handle on their own via mo/stretchy or dPr
  parseLeftRight() {
    this.skipSpace();
    const openTok = this.readBraceless();
    const open = openTok === "." ? "" : openTok;
    const row = this.parseRow(""); // consumed up to \right below via recursion guard
    return { open, row };
  }

  readBraceless() {
    this.skipSpace();
    if (this.peek() === "\\") {
      const start = this.i;
      this.i++;
      let name = "";
      while (!this.eof() && /[a-zA-Z]/.test(this.peek())) name += this.s[this.i++];
      if (name === "") { this.i = start + 1; return this.s[start + 1] || ""; } // \{ \}
      return { "{": "{", "}": "}" }[name] || name;
    }
    const c = this.peek();
    this.i++;
    return c;
  }

  parseCommand(allowScriptless) {
    this.i++; // consume backslash
    let name = "";
    while (!this.eof() && /[a-zA-Z]/.test(this.peek())) name += this.s[this.i++];
    if (name === "") {
      // escaped punctuation like \{ \} \\ \,
      const c = this.s[this.i++] || "";
      return { t: "op", v: c };
    }
    if (name === "frac") {
      const num = this.parseGroupOrSingle();
      const den = this.parseGroupOrSingle();
      return { t: "frac", num, den };
    }
    if (name === "sqrt") {
      this.skipSpace();
      let deg = null;
      if (this.peek() === "[") {
        this.i++;
        deg = this.parseRow("]");
        this.skipSpace();
        if (this.peek() === "]") this.i++;
      }
      const body = this.parseGroupOrSingle();
      return deg ? { t: "root", deg, body } : { t: "sqrt", body };
    }
    if (name === "text") {
      const body = this.parseGroupOrSingle();
      return { t: "text", body };
    }
    if (name === "left") {
      const open = this.readBraceless();
      const row = this.parseUntilRight();
      return { t: "fenced", open, close: row.close, body: row.row };
    }
    if (name === "right") return { t: "row", items: [] }; // stray \right, ignore
    if (NARY[name]) {
      this.skipSpace();
      let sub = null, sup = null;
      // limits can come in either order: \sum_{i=1}^{n} or \sum^{n}_{i=1}
      for (let guard = 0; guard < 2; guard++) {
        if (this.peek() === "_") { this.i++; sub = this.parseGroupOrSingle(); }
        else if (this.peek() === "^") { this.i++; sup = this.parseGroupOrSingle(); }
        this.skipSpace();
      }
      return { t: "nary", chr: NARY[name], sub, sup };
    }
    if (GREEK[name]) return { t: "id", v: GREEK[name] };
    if (SYMBOLS[name]) return { t: "op", v: SYMBOLS[name] };
    if (name === "quad") return { t: "op", v: "  " };
    // unknown command: render its name literally so nothing silently vanishes
    return { t: "text", body: { t: "row", items: [...name].map((c) => this.classify(c)) } };
  }

  // used by \left ... \right<delim>
  parseUntilRight() {
    const items = [];
    for (;;) {
      this.skipSpace();
      if (this.eof()) return { row: items.length === 1 ? items[0] : { t: "row", items }, close: "" };
      if (this.s.startsWith("\\right", this.i)) {
        this.i += 6;
        const close = this.readBraceless();
        return { row: items.length === 1 ? items[0] : { t: "row", items }, close };
      }
      items.push(this.parseAtomWithScripts());
    }
  }

  classify(c) {
    if (/[0-9.]/.test(c)) {
      // greedily consume a full number
      let n = c;
      while (!this.eof() && /[0-9.]/.test(this.peek())) n += this.s[this.i++];
      return { t: "num", v: n };
    }
    if (/[a-zA-Z]/.test(c)) return { t: "id", v: c };
    if ("+-=<>/!,;:".includes(c)) return { t: "op", v: c };
    if (c === "*") return { t: "op", v: "⋅" };
    if (GREEK_CHARS.has(c)) return { t: "id", v: c };
    return { t: "op", v: c };
  }
}

function escXmlText(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// AST -> MathML string
function astToMathML(node) {
  if (node == null) return "<mrow/>";
  switch (node.t) {
    case "row": return `<mrow>${node.items.map(astToMathML).join("")}</mrow>`;
    case "num": return `<mn>${escXmlText(node.v)}</mn>`;
    case "id": return `<mi>${escXmlText(node.v)}</mi>`;
    case "op": return `<mo>${escXmlText(node.v)}</mo>`;
    case "text": return `<mtext>${escXmlText(flattenText(node.body))}</mtext>`;
    case "frac": return `<mfrac>${astToMathML(node.num)}${astToMathML(node.den)}</mfrac>`;
    case "sqrt": return `<msqrt>${astToMathML(node.body)}</msqrt>`;
    case "root": return `<mroot>${astToMathML(node.body)}${astToMathML(node.deg)}</mroot>`;
    case "sup": return `<msup>${astToMathML(node.base)}${astToMathML(node.sup)}</msup>`;
    case "sub": return `<msub>${astToMathML(node.base)}${astToMathML(node.sub)}</msub>`;
    case "subsup": return `<msubsup>${astToMathML(node.base)}${astToMathML(node.sub)}${astToMathML(node.sup)}</msubsup>`;
    case "fenced": {
      const open = node.open ? `<mo>${escXmlText(node.open)}</mo>` : "";
      const close = node.close ? `<mo>${escXmlText(node.close)}</mo>` : "";
      return `<mrow>${open}${astToMathML(node.body)}${close}</mrow>`;
    }
    case "nary": {
      const chr = `<mo>${escXmlText(node.chr)}</mo>`;
      if (!node.sub && !node.sup) return chr;
      if (node.sub && node.sup) return `<munderover>${chr}${astToMathML(node.sub)}${astToMathML(node.sup)}</munderover>`;
      if (node.sub) return `<munder>${chr}${astToMathML(node.sub)}</munder>`;
      return `<mover>${chr}${astToMathML(node.sup)}</mover>`;
    }
    default: return "<mrow/>";
  }
}

function flattenText(node) {
  if (node == null) return "";
  if (node.t === "row") return node.items.map(flattenText).join("");
  if (node.v !== undefined) return node.v;
  return "";
}

// Compiles a LaTeX-subset string to a full `<math>...</math>` MathML string,
// suitable for `element.innerHTML = latexToMathML(src)` (browsers create real
// namespaced MathML nodes from that markup automatically, same as <svg>).
export function latexToMathML(latex) {
  let ast;
  try {
    ast = new LatexParser(String(latex || "")).parseRow("");
  } catch {
    ast = { t: "text", body: { t: "row", items: [{ t: "id", v: String(latex || "") }] } };
  }
  return `<math xmlns="${MATHML_NS}" display="inline">${astToMathML(ast)}</math>`;
}

// ---------------------------------------------------------------
// MathML DOM -> LaTeX-ish source (best-effort; used so an equation imported
// from a real Word file, or edited from one previously inserted here, can be
// re-opened in the same plain-text editing box)
// ---------------------------------------------------------------

function mmlNodeToLatex(el) {
  if (!el) return "";
  const tag = el.localName || el.tagName;
  const kids = () => [...el.children];
  switch (tag) {
    case "math": case "mrow": case "mstyle": case "semantics":
      return kids().map(mmlNodeToLatex).join(" ").replace(/\s+/g, " ").trim();
    case "mn": return el.textContent;
    case "mi": {
      const v = el.textContent;
      const rev = REVERSE_GREEK[v];
      return rev ? "\\" + rev + " " : v;
    }
    case "mo": {
      const v = el.textContent;
      const rev = REVERSE_SYMBOLS[v];
      return rev ? "\\" + rev + " " : v;
    }
    case "mtext": return `\\text{${el.textContent}}`;
    case "mfrac": return `\\frac{${mmlNodeToLatex(kids()[0])}}{${mmlNodeToLatex(kids()[1])}}`;
    case "msqrt": return `\\sqrt{${kids().map(mmlNodeToLatex).join(" ")}}`;
    case "mroot": return `\\sqrt[${mmlNodeToLatex(kids()[1])}]{${mmlNodeToLatex(kids()[0])}}`;
    case "msup": return `${wrapTerm(kids()[0])}^{${mmlNodeToLatex(kids()[1])}}`;
    case "msub": return `${wrapTerm(kids()[0])}_{${mmlNodeToLatex(kids()[1])}}`;
    case "msubsup": return `${wrapTerm(kids()[0])}_{${mmlNodeToLatex(kids()[1])}}^{${mmlNodeToLatex(kids()[2])}}`;
    case "munder": return `${mmlNodeToLatex(kids()[0])}_{${mmlNodeToLatex(kids()[1])}}`;
    case "mover": return `${mmlNodeToLatex(kids()[0])}^{${mmlNodeToLatex(kids()[1])}}`;
    case "munderover": return `${mmlNodeToLatex(kids()[0])}_{${mmlNodeToLatex(kids()[1])}}^{${mmlNodeToLatex(kids()[2])}}`;
    default: return el.textContent || "";
  }
}
function wrapTerm(el) {
  const s = mmlNodeToLatex(el);
  return s.length === 1 ? s : `{${s}}`;
}
const REVERSE_SYMBOLS = Object.fromEntries(
  Object.entries({ ...SYMBOLS, ...NARY }).map(([k, v]) => [v, k])
);

// Accepts either a live MathML DOM element (the <math> root) or an HTML
// string containing one; returns a best-effort LaTeX-subset source string.
export function mathMLToLatex(mathEl) {
  if (typeof mathEl === "string") {
    const doc = new DOMParser().parseFromString(`<div xmlns="http://www.w3.org/1999/xhtml">${mathEl}</div>`, "text/html");
    mathEl = doc.body.querySelector("math");
  }
  if (!mathEl) return "";
  return mmlNodeToLatex(mathEl).trim();
}

// ---------------------------------------------------------------
// MathML -> OMML (for .docx export)
// ---------------------------------------------------------------

const NARY_CHARS = new Set(Object.values(NARY));

// In MathML, a big-operator's limits (munder/mover/munderover, or a bare <mo>
// for a sum with no explicit limits) and its summand/integrand are separate
// *siblings* in the enclosing row — e.g. \sum_{i=1}^n i^2 is
// <mrow><munderover>..</munderover><msup>i,2</msup></mrow>, not one nested
// object. OMML's <m:nary> instead bundles operator+limits+body into a single
// element (<m:e> holds the body). So converting a row to OMML can't just map
// each child independently: it has to detect a nary head as the first item
// and fold every following sibling in that row into that nary's <m:e>.
function naryHead(el) {
  const tag = el.localName || el.tagName;
  const kids = [...el.children];
  if (tag === "munder") return { chr: kids[0] && kids[0].textContent, sub: kids[1], sup: null };
  if (tag === "mover") return { chr: kids[0] && kids[0].textContent, sub: null, sup: kids[1] };
  if (tag === "munderover") return { chr: kids[0] && kids[0].textContent, sub: kids[1], sup: kids[2] };
  if (tag === "mo" && NARY_CHARS.has(el.textContent)) return { chr: el.textContent, sub: null, sup: null };
  return null;
}

// Converts a sequence of sibling MathML nodes (the children of one <mrow>,
// or one <m:e>/<m:num>/... slot) to OMML, applying the nary head+body fold
// described above.
function rowToOmml(list) {
  if (!list.length) return "";
  const head = naryHead(list[0]);
  if (head) {
    const body = list.slice(1);
    const chr = escXmlText(head.chr || "");
    return `<m:nary><m:naryPr><m:chr m:val="${chr}"/><m:limLoc m:val="undOvr"/>` +
      `<m:subHide m:val="${head.sub ? 0 : 1}"/><m:supHide m:val="${head.sup ? 0 : 1}"/><m:ctrlPr/></m:naryPr>` +
      `<m:sub>${slotToOmml(head.sub)}</m:sub><m:sup>${slotToOmml(head.sup)}</m:sup>` +
      `<m:e>${rowToOmml(body)}</m:e></m:nary>`;
  }
  return list.map(mmlToOmmlNode).join("");
}
// A single child slot (mfrac's numerator, msup's base, ...) is itself either
// one element or (if the source was an <mrow>) a row of several — always
// resolve it through rowToOmml so a nary head as sole content still folds
// correctly (its body is then simply empty).
function slotToOmml(el) {
  if (!el) return "";
  const tag = el.localName || el.tagName;
  if (tag === "mrow") return rowToOmml([...el.children]);
  return rowToOmml([el]);
}

function mmlToOmmlNode(el) {
  if (!el) return "";
  const tag = el.localName || el.tagName;
  const kids = [...el.children];
  switch (tag) {
    case "math": case "mstyle": case "semantics":
      return rowToOmml(kids);
    case "mrow":
      return rowToOmml(kids);
    case "mn": case "mi": case "mo": case "mtext":
      return `<m:r><m:t xml:space="preserve">${escXmlText(el.textContent)}</m:t></m:r>`;
    case "mfrac":
      return `<m:f><m:fPr><m:ctrlPr/></m:fPr><m:num>${slotToOmml(kids[0])}</m:num><m:den>${slotToOmml(kids[1])}</m:den></m:f>`;
    case "msqrt":
      return `<m:rad><m:radPr><m:degHide m:val="1"/><m:ctrlPr/></m:radPr><m:deg/><m:e>${rowToOmml(kids)}</m:e></m:rad>`;
    case "mroot":
      return `<m:rad><m:radPr><m:ctrlPr/></m:radPr><m:deg>${slotToOmml(kids[1])}</m:deg><m:e>${slotToOmml(kids[0])}</m:e></m:rad>`;
    case "msup":
      return `<m:sSup><m:sSupPr><m:ctrlPr/></m:sSupPr><m:e>${slotToOmml(kids[0])}</m:e><m:sup>${slotToOmml(kids[1])}</m:sup></m:sSup>`;
    case "msub":
      return `<m:sSub><m:sSubPr><m:ctrlPr/></m:sSubPr><m:e>${slotToOmml(kids[0])}</m:e><m:sub>${slotToOmml(kids[1])}</m:sub></m:sSub>`;
    case "msubsup":
      return `<m:sSubSup><m:sSubSupPr><m:ctrlPr/></m:sSubSupPr><m:e>${slotToOmml(kids[0])}</m:e><m:sub>${slotToOmml(kids[1])}</m:sub><m:sup>${slotToOmml(kids[2])}</m:sup></m:sSubSup>`;
    // munder/mover/munderover are only ever handled as a nary head via
    // rowToOmml/naryHead above — they never reach this switch directly.
    default:
      return rowToOmml(kids);
  }
}

// mathEl: a live `<math>` element (browser or jsdom DOM). Returns the inner
// XML for `<m:oMath>...</m:oMath>` (caller wraps the tags).
export function mmlToOmml(mathEl) {
  return rowToOmml([...mathEl.children]);
}

// ---------------------------------------------------------------
// OMML -> MathML (for .docx import)
// ---------------------------------------------------------------

function ns(el, name) {
  for (const c of el.children) if (c.namespaceURI === M && c.localName === name) return c;
  return null;
}
// OMML text runs (<m:t>) don't distinguish operator/number/identifier the way
// MathML tags do — Word just renders every run in math italic unless it's
// recognized as an operator glyph. Classify on the way in so mo/mn read back
// correctly (in particular so operators round-trip through their \command
// name again — see mmlNodeToLatex's "mo" case).
const OPERATOR_CHARS = new Set([...Object.values(SYMBOLS), "+", "-", "=", "<", ">", "(", ")", "[", "]", ",", ";", ":", "!", "/"]);
function mmlTagForText(text) {
  if (/^[0-9]+(\.[0-9]+)?$/.test(text)) return "mn";
  if (OPERATOR_CHARS.has(text)) return "mo";
  return "mi";
}

function ommlValAttr(el) {
  return el ? (el.getAttributeNS(M, "val") ?? el.getAttribute("m:val")) : null;
}

function ommlNodeToMathML(el) {
  if (!el) return "";
  const tag = el.localName;
  switch (tag) {
    case "oMath": case "oMathPara":
      return el.children.length
        ? [...el.children].filter((c) => c.namespaceURI === M).map(ommlNodeToMathML).join("")
        : "";
    case "r": {
      const t = ns(el, "t");
      const text = t ? t.textContent : "";
      return `<${mmlTagForText(text)}>${escXmlText(text)}</${mmlTagForText(text)}>`;
    }
    case "f": {
      const num = ns(el, "num"), den = ns(el, "den");
      return `<mfrac>${ommlGroupToMathML(num)}${ommlGroupToMathML(den)}</mfrac>`;
    }
    case "rad": {
      const deg = ns(el, "deg"), e = ns(el, "e");
      const radPr = ns(el, "radPr");
      const hidden = radPr && ommlValAttr(ns(radPr, "degHide")) === "1";
      if (hidden || !deg || !deg.children.length) return `<msqrt>${ommlGroupToMathML(e)}</msqrt>`;
      return `<mroot>${ommlGroupToMathML(e)}${ommlGroupToMathML(deg)}</mroot>`;
    }
    case "sSup": {
      const e = ns(el, "e"), sup = ns(el, "sup");
      return `<msup>${ommlGroupToMathML(e)}${ommlGroupToMathML(sup)}</msup>`;
    }
    case "sSub": {
      const e = ns(el, "e"), sub = ns(el, "sub");
      return `<msub>${ommlGroupToMathML(e)}${ommlGroupToMathML(sub)}</msub>`;
    }
    case "sSubSup": {
      const e = ns(el, "e"), sub = ns(el, "sub"), sup = ns(el, "sup");
      return `<msubsup>${ommlGroupToMathML(e)}${ommlGroupToMathML(sub)}${ommlGroupToMathML(sup)}</msubsup>`;
    }
    case "nary": {
      const naryPr = ns(el, "naryPr");
      const chr = naryPr ? (ommlValAttr(ns(naryPr, "chr")) || "∑") : "∑";
      const subHidden = naryPr && ommlValAttr(ns(naryPr, "subHide")) === "1";
      const supHidden = naryPr && ommlValAttr(ns(naryPr, "supHide")) === "1";
      const sub = ns(el, "sub"), sup = ns(el, "sup"), e = ns(el, "e");
      const chrXml = `<mo>${escXmlText(chr)}</mo>`;
      const body = ommlGroupToMathML(e);
      let head;
      if (!subHidden && !supHidden && sub && sup && (sub.children.length || sup.children.length)) {
        head = `<munderover>${chrXml}${ommlGroupToMathML(sub)}${ommlGroupToMathML(sup)}</munderover>`;
      } else if (!subHidden && sub && sub.children.length) {
        head = `<munder>${chrXml}${ommlGroupToMathML(sub)}</munder>`;
      } else if (!supHidden && sup && sup.children.length) {
        head = `<mover>${chrXml}${ommlGroupToMathML(sup)}</mover>`;
      } else {
        head = chrXml;
      }
      return `<mrow>${head}${body}</mrow>`;
    }
    case "d": {
      const dPr = ns(el, "dPr");
      const beg = dPr ? (ommlValAttr(ns(dPr, "begChr")) ?? "(") : "(";
      const end = dPr ? (ommlValAttr(ns(dPr, "endChr")) ?? ")") : ")";
      const e = ns(el, "e");
      return `<mrow><mo>${escXmlText(beg)}</mo>${ommlGroupToMathML(e)}<mo>${escXmlText(end)}</mo></mrow>`;
    }
    default:
      return "";
  }
}
// <m:e>/<m:num>/<m:den>/... wrap a sequence of run-level objects (m:r, m:f, ...)
function ommlGroupToMathML(groupEl) {
  if (!groupEl) return "<mrow/>";
  const parts = [...groupEl.children].filter((c) => c.namespaceURI === M).map(ommlNodeToMathML);
  return parts.length === 1 ? parts[0] : `<mrow>${parts.join("")}</mrow>`;
}

// oMathEl: a live `<m:oMath>` or `<m:oMathPara>` element (parsed from
// document.xml). Returns a full `<math>...</math>` MathML string.
export function ommlToMathML(oMathEl) {
  return `<math xmlns="${MATHML_NS}" display="inline">${ommlNodeToMathML(oMathEl)}</math>`;
}

export const MATHML_NAMESPACE = MATHML_NS;
export const OMML_NAMESPACE = M;
