"use strict";
// Zero-dependency Node HTTP + WebSocket server for the Word-Compatible Editor.
// - Serves public/ and /api/*
// - Documents stored as JSON (data/<id>.json) + .docx binaries (data/<id>.docx)
// - Version history (data/<id>.versions.json, capped)
// - Revision-based conflict detection (baseRev -> 409)
// - Hand-rolled WebSocket (/ws) for presence + live document sync
// - Optional bearer-token auth (AUTH_TOKEN env), optional save webhook (SAVE_WEBHOOK_URL env)

const http = require("http");
const https = require("https");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const PORT = process.env.PORT || 3001;
const HOST = process.env.HOST || "127.0.0.1";
const AUTH_TOKEN = process.env.AUTH_TOKEN || "";
const WEBHOOK = process.env.SAVE_WEBHOOK_URL || "";
const MAX_BODY = 64 * 1024 * 1024; // 64 MB
const MAX_VERSIONS = 30;
const VERSION_MIN_INTERVAL = 90 * 1000; // min ms between auto snapshots

const ROOT = __dirname;
const PUBLIC = path.join(ROOT, "public");
const DATA = process.env.DATA_DIR
  ? (path.isAbsolute(process.env.DATA_DIR) ? process.env.DATA_DIR : path.join(ROOT, process.env.DATA_DIR))
  : path.join(ROOT, "data");
fs.mkdirSync(DATA, { recursive: true });

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".ico": "image/x-icon",
  ".woff2": "font/woff2",
};

const SECURITY_HEADERS = {
  "X-Content-Type-Options": "nosniff",
  "X-Frame-Options": "SAMEORIGIN",
  "Referrer-Policy": "no-referrer",
};

function send(res, code, body, headers = {}) {
  res.writeHead(code, { "Cache-Control": "no-store", ...SECURITY_HEADERS, ...headers });
  res.end(body);
}
function sendJSON(res, code, obj) {
  send(res, code, JSON.stringify(obj), { "Content-Type": "application/json; charset=utf-8" });
}
function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on("data", (c) => {
      size += c.length;
      if (size > MAX_BODY) { req.destroy(); reject(new Error("body too large")); return; }
      chunks.push(c);
    });
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}
async function readJSON(req) {
  const raw = (await readBody(req)).toString();
  if (!raw) return {};
  try { return JSON.parse(raw); } catch { throw Object.assign(new Error("invalid JSON body"), { status: 400 }); }
}

// IDs are UUIDs we generate; reject anything else so ids can never traverse paths.
const ID_RE = /^[0-9a-fA-F-]{8,64}$/;
function validId(id) { return ID_RE.test(id); }

function docPath(id) { return path.join(DATA, `${id}.json`); }
function docxPath(id) { return path.join(DATA, `${id}.docx`); }
function versionsPath(id) { return path.join(DATA, `${id}.versions.json`); }

function readMeta(id) {
  try { return JSON.parse(fs.readFileSync(docPath(id), "utf8")); } catch { return null; }
}
function readVersions(id) {
  try { return JSON.parse(fs.readFileSync(versionsPath(id), "utf8")); } catch { return []; }
}
function snapshotVersion(id, meta, force = false) {
  if (meta.state == null) return;
  const versions = readVersions(id);
  const last = versions[versions.length - 1];
  if (!force && last && Date.now() - last.t < VERSION_MIN_INTERVAL) return;
  if (last && last.state === meta.state && last.title === meta.title) return;
  versions.push({
    t: Date.now(), rev: meta.rev, title: meta.title, state: meta.state,
    pageSetup: meta.pageSetup || null, comments: meta.comments || [],
  });
  while (versions.length > MAX_VERSIONS) versions.shift();
  fs.writeFileSync(versionsPath(id), JSON.stringify(versions));
}
function writeMeta(id, body, opts = {}) {
  const existing = readMeta(id) || { id, title: "Untitled", createdAt: Date.now(), rev: 0 };
  const next = {
    id,
    title: body.title != null ? String(body.title).slice(0, 300) : existing.title,
    state: body.state !== undefined ? body.state : existing.state,
    pageSetup: body.pageSetup !== undefined ? body.pageSetup : (existing.pageSetup || null),
    comments: body.comments !== undefined ? body.comments : (existing.comments || []),
    trackChanges: body.trackChanges !== undefined ? !!body.trackChanges : !!existing.trackChanges,
    createdAt: existing.createdAt,
    updatedAt: Date.now(),
    rev: (existing.rev || 0) + 1,
  };
  fs.writeFileSync(docPath(id), JSON.stringify(next));
  if (body.state !== undefined) snapshotVersion(id, next, opts.forceVersion);
  fireWebhook(next);
  return next;
}
function metaSummary(d) {
  return { id: d.id, title: d.title, updatedAt: d.updatedAt, createdAt: d.createdAt, rev: d.rev || 0 };
}

function fireWebhook(meta) {
  if (!WEBHOOK) return;
  try {
    const u = new URL(WEBHOOK);
    const mod = u.protocol === "https:" ? https : http;
    const payload = JSON.stringify({ event: "save", id: meta.id, title: meta.title, rev: meta.rev, updatedAt: meta.updatedAt });
    const req = mod.request(u, {
      method: "POST",
      headers: { "Content-Type": "application/json", "Content-Length": Buffer.byteLength(payload) },
      timeout: 5000,
    }, (res) => res.resume());
    req.on("error", (e) => console.error("webhook error:", e.message));
    req.end(payload);
  } catch (e) { console.error("webhook error:", e.message); }
}

function authorized(req, url) {
  if (!AUTH_TOKEN) return true;
  const h = req.headers["authorization"] || "";
  if (h === `Bearer ${AUTH_TOKEN}`) return true;
  if (url && url.searchParams.get("token") === AUTH_TOKEN) return true;
  return false;
}

// ---- server-side HTML helpers (zero-dependency) ----

const BLOCK_TAGS = new Set(["p","div","h1","h2","h3","h4","h5","h6","ol","ul","li","table","tr","pre","blockquote","hr"]);

function stripTags(html) {
  return html.replace(/<[^>]*>/g, "");
}
function htmlToPlainText(html) {
  let s = String(html || "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "</p>\n")
    .replace(/<\/h[1-6]>/gi, "</h$&>\n")
    .replace(/<\/li>/gi, "</li>\n")
    .replace(/<\/tr>/gi, "</tr>\n")
    .replace(/<\/div>/gi, "</div>\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(d));
  return s.replace(/\n{3,}/g, "\n\n").trim() + "\n";
}
function countWords(html) {
  const text = htmlToPlainText(html).trim();
  const words = text ? text.split(/\s+/).length : 0;
  return { words, chars: text.length };
}
function escHtml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function exportStandaloneHtml(html, title) {
  return `<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>${escHtml(title || "Document")}</title>
<style>body{font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:1.4;max-width:8.5in;margin:24px auto;padding:0 1in;color:#111}</style>
</head><body>${html}</body></html>`;
}
function applyFormat(html, cmd, value) {
  // Wraps the ENTIRE provided html fragment with the formatting elements.
  // The caller is responsible for extracting the precise sub-range to format.
  // Map keys are lowercase (cmd arrives lowercased by the caller).
  const map = {
    bold:                  (h) => `<strong>${h}</strong>`,
    italic:                (h) => `<em>${h}</em>`,
    underline:             (h) => `<span style="text-decoration:underline">${h}</span>`,
    strikethrough:         (h) => `<s>${h}</s>`,
    subscript:             (h) => `<sub>${h}</sub>`,
    superscript:           (h) => `<sup>${h}</sup>`,
    forecolor:             (h) => value ? `<span style="color:${escHtml(value)}">${h}</span>` : h,
    hilitecolor:           (h) => value ? `<span style="background:${escHtml(value)}">${h}</span>` : h,
    fontname:              (h) => value ? `<span style="font-family:${escHtml(value)}">${h}</span>` : h,
    fontsize:              (h) => value ? `<span style="font-size:${escHtml(value)}">${h}</span>` : h,
    formatblock:           (h) => value ? `<${value}>${h}</${value}>` : h,
    justifyleft:           (h) => `<div style="text-align:left">${h}</div>`,
    justifycenter:         (h) => `<div style="text-align:center">${h}</div>`,
    justifyright:          (h) => `<div style="text-align:right">${h}</div>`,
    justifyfull:           (h) => `<div style="text-align:justify">${h}</div>`,
    insertorderedlist:     (h) => `<ol>${h.split(/\n+/).filter(Boolean).map((l) => `<li>${l}</li>`).join("")}</ol>`,
    insertunorderedlist:   (h) => `<ul>${h.split(/\n+/).filter(Boolean).map((l) => `<li>${l}</li>`).join("")}</ul>`,
  };
  const fn = map[cmd];
  if (!fn) throw new Error("unknown format command: " + cmd);
  return fn(html);
}

// ---- API ----
async function api(req, res, url) {
  const method = req.method;
  const p = url.pathname;

  if (p === "/api/health") return sendJSON(res, 200, { ok: true, uptime: process.uptime() });

  if (!authorized(req, url)) return sendJSON(res, 401, { error: "unauthorized" });

  if (p === "/api/documents" && method === "GET") {
    const list = fs.readdirSync(DATA)
      .filter((f) => f.endsWith(".json") && !f.endsWith(".versions.json"))
      .map((f) => {
        try { return metaSummary(JSON.parse(fs.readFileSync(path.join(DATA, f), "utf8"))); }
        catch { return null; }
      })
      .filter(Boolean)
      .sort((a, b) => b.updatedAt - a.updatedAt);
    return sendJSON(res, 200, list);
  }

  if (p === "/api/documents" && method === "POST") {
    const body = await readJSON(req);
    const id = crypto.randomUUID();
    const doc = writeMeta(id, { title: body.title || "Untitled", state: body.state !== undefined ? body.state : null });
    return sendJSON(res, 200, doc);
  }

  if (p === "/api/documents/import" && method === "POST") {
    const body = await readJSON(req);
    if (!body.data) return sendJSON(res, 400, { error: "no data" });
    const id = crypto.randomUUID();
    const title = String(body.name || "Imported").replace(/\.(docx|txt|html?)$/i, "").slice(0, 300);
    fs.writeFileSync(docxPath(id), Buffer.from(body.data, "base64"));
    const doc = writeMeta(id, { title, state: null });
    return sendJSON(res, 200, { id: doc.id, title: doc.title, hasDocx: true });
  }

  let m = p.match(/^\/api\/documents\/([^/]+)$/);
  if (m) {
    const id = m[1];
    if (!validId(id)) return sendJSON(res, 400, { error: "bad id" });
    if (method === "GET") {
      const doc = readMeta(id);
      return doc ? sendJSON(res, 200, doc) : sendJSON(res, 404, { error: "not found" });
    }
    if (method === "PUT") {
      const existing = readMeta(id);
      if (!existing) return sendJSON(res, 404, { error: "not found" });
      const body = await readJSON(req);
      // Optimistic concurrency: caller may send the rev it based its edit on.
      if (body.baseRev != null && body.baseRev !== (existing.rev || 0)) {
        return sendJSON(res, 409, { error: "conflict", current: existing });
      }
      return sendJSON(res, 200, writeMeta(id, body));
    }
    if (method === "DELETE") {
      for (const f of [docPath(id), docxPath(id), versionsPath(id)]) {
        if (fs.existsSync(f)) fs.unlinkSync(f);
      }
      closeRoom(id);
      return sendJSON(res, 200, { ok: true });
    }
  }

  m = p.match(/^\/api\/documents\/([^/]+)\/docx$/);
  if (m) {
    const id = m[1];
    if (!validId(id)) return sendJSON(res, 400, { error: "bad id" });
    if (method === "PUT") {
      const body = await readJSON(req);
      if (!body.data) return sendJSON(res, 400, { error: "no data" });
      fs.writeFileSync(docxPath(id), Buffer.from(body.data, "base64"));
      return sendJSON(res, 200, { ok: true });
    }
    if (method === "GET") {
      if (!fs.existsSync(docxPath(id))) return sendJSON(res, 404, { error: "no docx" });
      const meta = readMeta(id);
      const name = ((meta && meta.title) || id).replace(/[^\w一-鿿 .-]+/g, "_");
      return send(res, 200, fs.readFileSync(docxPath(id)), {
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": `attachment; filename="${name}.docx"`,
      });
    }
  }

  m = p.match(/^\/api\/documents\/([^/]+)\/versions$/);
  if (m && method === "GET") {
    const id = m[1];
    if (!validId(id)) return sendJSON(res, 400, { error: "bad id" });
    const versions = readVersions(id).map((v, i) => ({
      index: i, t: v.t, rev: v.rev, title: v.title, size: v.state ? v.state.length : 0,
    }));
    return sendJSON(res, 200, versions);
  }

  m = p.match(/^\/api\/documents\/([^/]+)\/versions\/(\d+)$/);
  if (m && method === "GET") {
    const id = m[1];
    if (!validId(id)) return sendJSON(res, 400, { error: "bad id" });
    const v = readVersions(id)[Number(m[2])];
    return v ? sendJSON(res, 200, v) : sendJSON(res, 404, { error: "not found" });
  }

  m = p.match(/^\/api\/documents\/([^/]+)\/restore$/);
  if (m && method === "POST") {
    const id = m[1];
    if (!validId(id)) return sendJSON(res, 400, { error: "bad id" });
    const existing = readMeta(id);
    if (!existing) return sendJSON(res, 404, { error: "not found" });
    const body = await readJSON(req);
    const v = readVersions(id)[Number(body.index)];
    if (!v) return sendJSON(res, 404, { error: "version not found" });
    snapshotVersion(id, existing, true); // keep the pre-restore state recoverable
    const doc = writeMeta(id, { title: v.title, state: v.state, pageSetup: v.pageSetup, comments: v.comments || [] }, { forceVersion: true });
    broadcast(id, null, { type: "update", from: { id: "server", user: "Version restore", color: "#666" }, rev: doc.rev, title: doc.title, state: doc.state, pageSetup: doc.pageSetup, comments: doc.comments });
    return sendJSON(res, 200, doc);
  }

  // ---------------------------------------------------------------
  // RESTful document-content API (mirrors the SDK commands)
  // ---------------------------------------------------------------

  // GET /api/documents/{id}/content   → { html, text }
  // PUT /api/documents/{id}/content   → set content { html? }
  // POST /api/documents/{id}/insert-html → append { html }
  // POST /api/documents/{id}/insert-text → append { text }
  if ((m = p.match(/^\/api\/documents\/([^/]+)\/content$/))) {
    const id = m[1];
    if (!validId(id)) return sendJSON(res, 400, { error: "bad id" });
    const doc = readMeta(id);
    if (!doc) return sendJSON(res, 404, { error: "not found" });
    if (method === "GET") {
      const html = (doc.state || "").replace(/^(<p><br><\/p>\s*)+$/, "");
      const text = html ? htmlToPlainText(html) : "";
      const wc = countWords(html);
      return sendJSON(res, 200, { id: doc.id, title: doc.title, html, text, ...wc, pageSetup: doc.pageSetup || null, rev: doc.rev });
    }
    if (method === "PUT") {
      const body = await readJSON(req);
      const updated = writeMeta(id, { state: body.html !== undefined ? body.html : doc.state, title: body.title !== undefined ? body.title : doc.title });
      broadcast(id, null, { type: "update", from: { id: "api", user: "REST API", color: "#666" }, rev: updated.rev, title: updated.title, state: updated.state, pageSetup: updated.pageSetup, comments: updated.comments });
      return sendJSON(res, 200, { ok: true, rev: updated.rev });
    }
    return sendJSON(res, 405, { error: "method not allowed" });
  }

  // GET /api/documents/{id}/text   → { text }
  if ((m = p.match(/^\/api\/documents\/([^/]+)\/text$/))) {
    const id = m[1];
    if (!validId(id)) return sendJSON(res, 400, { error: "bad id" });
    const doc = readMeta(id);
    if (!doc) return sendJSON(res, 404, { error: "not found" });
    return sendJSON(res, 200, { text: htmlToPlainText(doc.state || "") });
  }

  // POST /api/documents/{id}/insert-html   append HTML to the stored document
  if ((m = p.match(/^\/api\/documents\/([^/]+)\/insert-html$/)) && method === "POST") {
    const id = m[1];
    if (!validId(id)) return sendJSON(res, 400, { error: "bad id" });
    const doc = readMeta(id);
    if (!doc) return sendJSON(res, 404, { error: "not found" });
    if (doc.state == null) doc.state = "<p><br></p>";
    const body = await readJSON(req);
    const fragment = String(body.html || "").trim();
    if (!fragment) return sendJSON(res, 400, { error: "html required" });
    const state = doc.state.replace(/<p><br><\/p>\s*$/, "") + fragment + "<p><br></p>";
    const updated = writeMeta(id, { state });
    broadcast(id, null, { type: "update", from: { id: "api", user: "REST API", color: "#666" }, rev: updated.rev, title: updated.title, state: updated.state });
    return sendJSON(res, 200, { ok: true, rev: updated.rev });
  }

  // POST /api/documents/{id}/insert-text   append plain-text (wraps in <p>)
  // shortcut — same as insert-html but auto-paragraphs
  if ((m = p.match(/^\/api\/documents\/([^/]+)\/insert-text$/)) && method === "POST") {
    const id = m[1];
    if (!validId(id)) return sendJSON(res, 400, { error: "bad id" });
    const doc = readMeta(id);
    if (!doc) return sendJSON(res, 404, { error: "not found" });
    const body = await readJSON(req);
    const text = String(body.text || "").trim();
    if (!text) return sendJSON(res, 400, { error: "text required" });
    const html = text.split(/\r?\n/).filter(Boolean).map((l) => `<p>${escHtml(l)}</p>`).join("");
    const state = (doc.state || "<p><br></p>").replace(/<p><br><\/p>\s*$/, "") + html + "<p><br></p>";
    const updated = writeMeta(id, { state });
    broadcast(id, null, { type: "update", from: { id: "api", user: "REST API", color: "#666" }, rev: updated.rev, title: updated.title, state: updated.state });
    return sendJSON(res, 200, { ok: true, rev: updated.rev });
  }

  // PUT /api/documents/{id}/title   → { title }
  if ((m = p.match(/^\/api\/documents\/([^/]+)\/title$/)) && method === "PUT") {
    const id = m[1];
    if (!validId(id)) return sendJSON(res, 400, { error: "bad id" });
    const doc = readMeta(id);
    if (!doc) return sendJSON(res, 404, { error: "not found" });
    const body = await readJSON(req);
    const updated = writeMeta(id, { title: String(body.title || doc.title).slice(0, 300) });
    return sendJSON(res, 200, { ok: true, rev: updated.rev, title: updated.title });
  }

  // GET /api/documents/{id}/meta  →  { id, title, rev, words, chars, pageSetup, trackChanges, commentCount }
  if ((m = p.match(/^\/api\/documents\/([^/]+)\/meta$/)) && method === "GET") {
    const id = m[1];
    if (!validId(id)) return sendJSON(res, 400, { error: "bad id" });
    const doc = readMeta(id);
    if (!doc) return sendJSON(res, 404, { error: "not found" });
    const wc = countWords(doc.state || "");
    return sendJSON(res, 200, {
      id: doc.id, title: doc.title, rev: doc.rev || 0,
      pageSetup: doc.pageSetup || null, trackChanges: !!doc.trackChanges,
      commentCount: (doc.comments || []).length, ...wc,
    });
  }
  // PUT /api/documents/{id}/meta  →  update { title?, pageSetup? }
  if ((m = p.match(/^\/api\/documents\/([^/]+)\/meta$/)) && method === "PUT") {
    const id = m[1];
    if (!validId(id)) return sendJSON(res, 400, { error: "bad id" });
    const doc = readMeta(id);
    if (!doc) return sendJSON(res, 404, { error: "not found" });
    const body = await readJSON(req);
    const patch = {};
    if (body.title !== undefined) patch.title = String(body.title).slice(0, 300);
    if (body.pageSetup !== undefined) patch.pageSetup = body.pageSetup;
    const updated = writeMeta(id, patch);
    return sendJSON(res, 200, { ok: true, rev: updated.rev, title: updated.title, pageSetup: updated.pageSetup });
  }

  // GET /api/documents/{id}/export?fmt=docx|html|txt  → download
  if ((m = p.match(/^\/api\/documents\/([^/]+)\/export$/))) {
    const id = m[1];
    if (!validId(id)) return sendJSON(res, 400, { error: "bad id" });
    const doc = readMeta(id);
    if (!doc) return sendJSON(res, 404, { error: "not found" });
    const fmt = url.searchParams.get("fmt") || "docx";
    const title = (doc.title || "document").replace(/[^\w .-]/g, "_");
    try {
      if (fmt === "txt") {
        const text = htmlToPlainText(doc.state || "");
        return send(res, 200, text, { "Content-Type": "text/plain; charset=utf-8", "Content-Disposition": `attachment; filename="${title}.txt"` });
      }
      if (fmt === "html") {
        const standalone = exportStandaloneHtml(doc.state || "", doc.title || "Document");
        return send(res, 200, standalone, { "Content-Type": "text/html; charset=utf-8", "Content-Disposition": `attachment; filename="${title}.html"` });
      }
      // docx — rely on the editor .docx binary if stored
      if (fmt === "docx" && fs.existsSync(docxPath(id))) {
        return send(res, 200, fs.readFileSync(docxPath(id)), {
          "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          "Content-Disposition": `attachment; filename="${title}.docx"`,
        });
      }
      return sendJSON(res, 400, { error: `unsupported or unavailable format: ${fmt}` });
    } catch (e) { return sendJSON(res, 500, { error: e.message }); }
  }

  // ---- comments ----
  // GET  /api/documents/{id}/comments     → { comments }
  // POST /api/documents/{id}/comments     → add { text, author? }
  if ((m = p.match(/^\/api\/documents\/([^/]+)\/comments$/))) {
    const id = m[1];
    if (!validId(id)) return sendJSON(res, 400, { error: "bad id" });
    const doc = readMeta(id);
    if (!doc) return sendJSON(res, 404, { error: "not found" });
    if (method === "GET") return sendJSON(res, 200, { comments: doc.comments || [] });
    if (method === "POST") {
      const body = await readJSON(req);
      const text = String(body.text || "").trim();
      if (!text) return sendJSON(res, 400, { error: "text required" });
      const c = { id: "c" + Date.now().toString(36) + Math.random().toString(36).slice(2, 5), author: body.author || "API", text, createdAt: Date.now(), resolved: false, replies: [] };
      const comments = [...(doc.comments || []), c];
      writeMeta(id, { comments });
      return sendJSON(res, 200, { ok: true, id: c.id });
    }
    return sendJSON(res, 405, { error: "method not allowed" });
  }

  // ---- track changes ----
  // GET  /api/documents/{id}/track-changes  → { enabled }
  // PUT  /api/documents/{id}/track-changes  → { enabled }
  if ((m = p.match(/^\/api\/documents\/([^/]+)\/track-changes$/))) {
    const id = m[1];
    if (!validId(id)) return sendJSON(res, 400, { error: "bad id" });
    const doc = readMeta(id);
    if (!doc) return sendJSON(res, 404, { error: "not found" });
    if (method === "GET") return sendJSON(res, 200, { enabled: !!doc.trackChanges });
    if (method === "PUT") {
      const body = await readJSON(req);
      const on = body.enabled === true || body.enabled === "true";
      writeMeta(id, { trackChanges: on });
      return sendJSON(res, 200, { ok: true, enabled: on });
    }
    return sendJSON(res, 405, { error: "method not allowed" });
  }

  // ---- page setup ----
  // GET  /api/documents/{id}/page-setup → { pageSetup }
  // PUT  /api/documents/{id}/page-setup → { pageSetup? }
  if ((m = p.match(/^\/api\/documents\/([^/]+)\/page-setup$/))) {
    const id = m[1];
    if (!validId(id)) return sendJSON(res, 400, { error: "bad id" });
    const doc = readMeta(id);
    if (!doc) return sendJSON(res, 404, { error: "not found" });
    if (method === "GET") return sendJSON(res, 200, { pageSetup: doc.pageSetup || null });
    if (method === "PUT") {
      const body = await readJSON(req);
      const merged = { ...(doc.pageSetup || {}), ...(body.pageSetup || {}) };
      if (body.size) merged.size = body.size;
      if (body.orientation) merged.orientation = body.orientation;
      if (body.margins) merged.margins = { ...((doc.pageSetup && doc.pageSetup.margins) || {}), ...body.margins };
      writeMeta(id, { pageSetup: merged });
      return sendJSON(res, 200, { ok: true, pageSetup: merged });
    }
    return sendJSON(res, 405, { error: "method not allowed" });
  }

  // ---- format transform — applies formatting commands to an HTML fragment ----
  // POST /api/format { html, cmd, value? } → { html }
  if (p === "/api/format" && method === "POST") {
    const body = await readJSON(req);
    const html = String(body.html || "");
    const cmd = String(body.cmd || "").toLowerCase();
    const val = body.value != null ? String(body.value) : null;
    if (!html || !cmd) return sendJSON(res, 400, { error: "html and cmd required" });
    let transformed;
    try { transformed = applyFormat(html, cmd, val); } catch (e) { return sendJSON(res, 400, { error: e.message }); }
    return sendJSON(res, 200, { html: transformed });
  }

  return sendJSON(res, 404, { error: "not found" });
}

// ---- static ----
function serveStatic(req, res, url) {
  let rel = decodeURIComponent(url.pathname);
  if (rel === "/" || rel === "") rel = "/index.html";
  const filePath = path.resolve(PUBLIC, "." + rel);
  if (filePath !== PUBLIC && !filePath.startsWith(PUBLIC + path.sep)) return send(res, 403, "forbidden");
  fs.readFile(filePath, (err, buf) => {
    if (err) return send(res, 404, "not found");
    send(res, 200, buf, { "Content-Type": MIME[path.extname(filePath).toLowerCase()] || "application/octet-stream" });
  });
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://${req.headers.host || "localhost"}`);
  try {
    if (url.pathname.startsWith("/api/")) await api(req, res, url);
    else serveStatic(req, res, url);
  } catch (e) {
    console.error(e);
    if (!res.headersSent) sendJSON(res, e.status || 500, { error: String(e.message || e) });
  }
});

// ============================================================
// WebSocket (RFC 6455) — hand-rolled, no dependencies.
// Rooms keyed by document id; relays presence, updates, cursors.
// ============================================================
const WS_GUID = "258EAFA5-E914-47DA-95CA-C5AB0DC85B11";
const rooms = new Map(); // docId -> Set<client>
const COLORS = ["#e91e63", "#2196f3", "#4caf50", "#ff9800", "#9c27b0", "#00bcd4", "#795548", "#607d8b"];
let clientSeq = 0;

function encodeFrame(opcode, payload) {
  const len = payload.length;
  let header;
  if (len < 126) {
    header = Buffer.from([0x80 | opcode, len]);
  } else if (len < 65536) {
    header = Buffer.alloc(4);
    header[0] = 0x80 | opcode; header[1] = 126;
    header.writeUInt16BE(len, 2);
  } else {
    header = Buffer.alloc(10);
    header[0] = 0x80 | opcode; header[1] = 127;
    header.writeBigUInt64BE(BigInt(len), 2);
  }
  return Buffer.concat([header, payload]);
}
function wsSend(client, obj) {
  if (client.socket.destroyed) return;
  try { client.socket.write(encodeFrame(1, Buffer.from(JSON.stringify(obj)))); } catch {}
}
function roomUsers(docId) {
  const set = rooms.get(docId);
  if (!set) return [];
  return [...set].map((c) => ({ id: c.id, user: c.user, color: c.color }));
}
function broadcast(docId, exceptClient, obj) {
  const set = rooms.get(docId);
  if (!set) return;
  for (const c of set) if (c !== exceptClient) wsSend(c, obj);
}
function leaveRoom(client) {
  const set = rooms.get(client.docId);
  if (!set) return;
  set.delete(client);
  if (set.size === 0) rooms.delete(client.docId);
  else broadcast(client.docId, null, { type: "presence", users: roomUsers(client.docId) });
}
function closeRoom(docId) {
  const set = rooms.get(docId);
  if (!set) return;
  for (const c of set) { try { c.socket.destroy(); } catch {} }
  rooms.delete(docId);
}
function handleMessage(client, text) {
  let msg;
  try { msg = JSON.parse(text); } catch { return; }
  if (msg.type === "hello") {
    const docId = String(msg.docId || "");
    if (!validId(docId)) return;
    if (client.docId) leaveRoom(client);
    client.docId = docId;
    client.user = String(msg.user || "Guest").slice(0, 60);
    if (!rooms.has(docId)) rooms.set(docId, new Set());
    rooms.get(docId).add(client);
    wsSend(client, { type: "welcome", id: client.id, color: client.color, users: roomUsers(docId) });
    broadcast(docId, client, { type: "presence", users: roomUsers(docId) });
    return;
  }
  if (!client.docId) return;
  const from = { id: client.id, user: client.user, color: client.color };
  if (msg.type === "update") {
    broadcast(client.docId, client, {
      type: "update", from, rev: msg.rev, title: msg.title, state: msg.state,
      pageSetup: msg.pageSetup, comments: msg.comments, trackChanges: msg.trackChanges,
    });
  } else if (msg.type === "cursor") {
    broadcast(client.docId, client, { type: "cursor", from, at: msg.at });
  }
}

server.on("upgrade", (req, socket) => {
  const url = new URL(req.url, `http://${req.headers.host || "localhost"}`);
  if (url.pathname !== "/ws" || !authorized(req, url)) {
    socket.write("HTTP/1.1 403 Forbidden\r\n\r\n"); socket.destroy(); return;
  }
  const key = req.headers["sec-websocket-key"];
  if (!key) { socket.destroy(); return; }
  const accept = crypto.createHash("sha1").update(key + WS_GUID).digest("base64");
  socket.write(
    "HTTP/1.1 101 Switching Protocols\r\n" +
    "Upgrade: websocket\r\nConnection: Upgrade\r\n" +
    `Sec-WebSocket-Accept: ${accept}\r\n\r\n`
  );
  socket.setNoDelay(true);

  const client = { id: "u" + (++clientSeq), color: COLORS[clientSeq % COLORS.length], socket, docId: null, user: "Guest" };
  let buf = Buffer.alloc(0);
  let fragments = [];

  socket.on("data", (chunk) => {
    buf = Buffer.concat([buf, chunk]);
    for (;;) {
      if (buf.length < 2) return;
      const fin = (buf[0] & 0x80) !== 0;
      const opcode = buf[0] & 0x0f;
      const masked = (buf[1] & 0x80) !== 0;
      let len = buf[1] & 0x7f;
      let off = 2;
      if (len === 126) {
        if (buf.length < 4) return;
        len = buf.readUInt16BE(2); off = 4;
      } else if (len === 127) {
        if (buf.length < 10) return;
        const big = buf.readBigUInt64BE(2);
        if (big > BigInt(MAX_BODY)) { socket.destroy(); return; }
        len = Number(big); off = 10;
      }
      const maskLen = masked ? 4 : 0;
      if (buf.length < off + maskLen + len) return;
      let payload = buf.subarray(off + maskLen, off + maskLen + len);
      if (masked) {
        const mask = buf.subarray(off, off + 4);
        const un = Buffer.allocUnsafe(len);
        for (let i = 0; i < len; i++) un[i] = payload[i] ^ mask[i % 4];
        payload = un;
      }
      buf = buf.subarray(off + maskLen + len);

      if (opcode === 8) { // close
        try { socket.write(encodeFrame(8, Buffer.alloc(0))); } catch {}
        socket.destroy();
        return;
      }
      if (opcode === 9) { // ping -> pong
        try { socket.write(encodeFrame(10, payload)); } catch {}
        continue;
      }
      if (opcode === 10) continue; // pong
      if (opcode === 1 || opcode === 2 || opcode === 0) {
        fragments.push(payload);
        if (fin) {
          const full = Buffer.concat(fragments);
          fragments = [];
          handleMessage(client, full.toString("utf8"));
        }
      }
    }
  });
  const cleanup = () => { if (client.docId) leaveRoom(client); };
  socket.on("close", cleanup);
  socket.on("error", () => { cleanup(); try { socket.destroy(); } catch {} });
});

// keepalive pings
setInterval(() => {
  for (const set of rooms.values()) {
    for (const c of set) {
      try { c.socket.write(encodeFrame(9, Buffer.alloc(0))); } catch {}
    }
  }
}, 30000).unref();

server.listen(PORT, HOST, () => console.log(`word-editor on http://${HOST}:${PORT}${AUTH_TOKEN ? " (auth enabled)" : ""}`));
