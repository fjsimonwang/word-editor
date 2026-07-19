/**
 * WordEditor REST API Client — use the editor from ANY HTTP client
 * (no iframe, no postMessage).
 *
 * Usage:
 *   const api = new WordEditorRestClient("http://localhost:3001");
 *   await api.createDocument("Report");
 *   await api.setContent("<p>Hello</p>");
 *   const meta = await api.getMeta();
 *   await api.insertHtml("<p>More</p>");
 *   await api.setPageSetup({ size: "A4", margins: { top: 1 } });
 *   console.log(meta.words);
 */
(function (global) {
  "use strict";

  class WordEditorRestClient {
    /**
     * @param {string} baseUrl  Server base URL (e.g. "http://localhost:3001")
     * @param {string} [docId]  Optional document ID (auto-creates if omitted)
     */
    constructor(baseUrl, docId) {
      this.baseUrl = baseUrl.replace(/\/+$/, "");
      this.docId = docId || null;
    }

    // ─── helpers ───
    async _fetch(method, path, body) {
      const opts = {
        method,
        headers: body ? { "Content-Type": "application/json" } : {},
      };
      if (body) opts.body = JSON.stringify(body);
      const res = await fetch(this.baseUrl + path, opts);
      const contentType = res.headers.get("content-type") || "";
      if (contentType.includes("application/json")) {
        const data = await res.json();
        if (!res.ok) throw new Error(data.error || `HTTP ${res.status}`);
        return data;
      }
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      return res;
    }

    async _json(method, path, body) {
      return this._fetch(method, path, body);
    }

    async _ensureDoc() {
      if (!this.docId) {
        const doc = await this._json("POST", "/api/documents", { title: "Untitled" });
        this.docId = doc.id;
      }
      return this.docId;
    }

    // ─── document management ───
    async createDocument(title) {
      const doc = await this._json("POST", "/api/documents", { title: title || "Untitled" });
      this.docId = doc.id;
      return { ok: true, id: doc.id, rev: doc.rev };
    }

    async loadDocument(id) {
      const doc = await this._json("GET", "/api/documents/" + id);
      this.docId = doc.id;
      return { ok: true, id: doc.id, rev: doc.rev };
    }

    async listDocuments() {
      const list = await this._json("GET", "/api/documents");
      return { documents: list };
    }

    async deleteDocument(id) {
      await this._json("DELETE", "/api/documents/" + id);
      return { ok: true };
    }

    async newDocument(title) {
      return this.createDocument(title);
    }

    // ─── content ───
    async getContent() {
      const id = await this._ensureDoc();
      const data = await this._json("GET", `/api/documents/${id}/content`);
      return { html: data.html };
    }

    async setContent(html) {
      const id = await this._ensureDoc();
      await this._json("PUT", `/api/documents/${id}/content`, { html });
      return { ok: true };
    }

    async getText() {
      const id = await this._ensureDoc();
      const data = await this._json("GET", `/api/documents/${id}/text`);
      return { text: data.text };
    }

    async insertHtml(html) {
      const id = await this._ensureDoc();
      await this._json("POST", `/api/documents/${id}/insert-html`, { html });
      return { ok: true };
    }

    async insertText(text) {
      const id = await this._ensureDoc();
      await this._json("POST", `/api/documents/${id}/insert-text`, { text });
      return { ok: true };
    }

    async getMeta() {
      const id = await this._ensureDoc();
      return this._json("GET", `/api/documents/${id}/meta`);
    }

    async setTitle(title) {
      const id = await this._ensureDoc();
      await this._json("PUT", `/api/documents/${id}/title`, { title });
      return { ok: true };
    }

    async save() {
      // REST API auto-saves on every write; no flush needed.
      const id = await this._ensureDoc();
      const doc = await this._json("GET", `/api/documents/${id}/content`);
      return { ok: true, rev: doc.rev };
    }

    async focus() {
      // no-op for REST; the client is responsible for its own UI focus
      return { ok: true };
    }

    // ─── formatting (operates on a provided HTML fragment) ───
    async format(cmd, value) {
      throw new Error("format() requires a selected HTML fragment; use formatHtml(cmd, html, value) instead");
    }

    async formatHtml(cmd, html, value) {
      const data = await this._json("POST", "/api/format", { cmd, html, value });
      return { ok: true, html: data.html };
    }

    async bold(html)        { return this.formatHtml("bold", html); }
    async italic(html)      { return this.formatHtml("italic", html); }
    async underline(html)   { return this.formatHtml("underline", html); }
    async strikeThrough(html) { return this.formatHtml("strikeThrough", html); }
    async subscript(html)   { return this.formatHtml("subscript", html); }
    async superscript(html) { return this.formatHtml("superscript", html); }
    async foreColor(color, html)    { return this.formatHtml("foreColor", html, color); }
    async hiliteColor(color, html)  { return this.formatHtml("hiliteColor", html, color); }
    async fontName(font, html)      { return this.formatHtml("fontName", html, font); }
    async fontSize(size, html)      { return this.formatHtml("fontSize", html, size); }
    async formatBlock(tag, html)    { return this.formatHtml("formatBlock", html, tag); }
    async justifyLeft(html)   { return this.formatHtml("justifyLeft", html); }
    async justifyCenter(html) { return this.formatHtml("justifyCenter", html); }
    async justifyRight(html)  { return this.formatHtml("justifyRight", html); }
    async justifyFull(html)   { return this.formatHtml("justifyFull", html); }
    async insertOrderedList(html)   { return this.formatHtml("insertOrderedList", html); }
    async insertUnorderedList(html) { return this.formatHtml("insertUnorderedList", html); }

    // ─── insert helpers ───
    async insertImage(src, opts) {
      const dims = opts ? ` width="${opts.width || ""}" height="${opts.height || ""}"` : "";
      const alt = opts && opts.alt ? ` alt="${opts.alt.replace(/"/g, "&quot;")}"` : "";
      return this.insertHtml(`<img src="${src.replace(/"/g, "&quot;")}"${alt}${dims}>`);
    }

    async insertLink(href, text) {
      return this.insertHtml(`<a href="${href.replace(/"/g, "&quot;")}">${(text || href).replace(/[<>&]/g, (c) => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;" })[c])}</a>`);
    }

    async insertTable(rows, cols) {
      rows = Math.max(1, Math.min(50, rows || 2));
      cols = Math.max(1, Math.min(20, cols || 2));
      let tbl = "<table><tbody>";
      for (let r = 0; r < rows; r++) {
        tbl += "<tr>";
        for (let c = 0; c < cols; c++) tbl += "<td><br></td>";
        tbl += "</tr>";
      }
      tbl += "</tbody></table><p><br></p>";
      return this.insertHtml(tbl);
    }

    async insertSymbol(ch) {
      return this.insertText(ch);
    }

    async insertPageBreak() {
      return this.insertHtml('<p class="page-break"><br></p><p><br></p>');
    }

    async insertBlankPage() {
      return this.insertHtml('<p class="page-break"><br></p><p><br></p><p class="page-break"><br></p><p><br></p>');
    }

    async insertHr() {
      return this.insertHtml("<hr>");
    }

    // ─── comments ───
    async addComment(text, author) {
      const id = await this._ensureDoc();
      const data = await this._json("POST", `/api/documents/${id}/comments`, { text, author });
      return { id: data.id };
    }

    async getComments() {
      const id = await this._ensureDoc();
      return this._json("GET", `/api/documents/${id}/comments`);
    }

    // ─── track changes ───
    async setTrackChanges(on) {
      const id = await this._ensureDoc();
      const data = await this._json("PUT", `/api/documents/${id}/track-changes`, { enabled: on });
      return { ok: true, on: data.enabled };
    }

    async getChanges() {
      // Tracked changes are stored in the document state as del/ins marks, not as a separate list.
      // Use getContent() and inspect the HTML for .tc-del / .tc-ins markers.
      return { changes: [] };
    }

    async acceptAllChanges() {
      throw new Error("acceptAllChanges requires the embed SDK (needs browser DOM)");
    }

    async rejectAllChanges() {
      throw new Error("rejectAllChanges requires the embed SDK (needs browser DOM)");
    }

    // ─── headers / footers / page numbers ───
    async setHeader(text, align) {
      const ps = await this.getPageSetup();
      const ch = ps.pageSetup && ps.pageSetup.chrome || {};
      if (text) ch.header = { text, align: align || "center" };
      else delete ch.header;
      return this._json("PUT", `/api/documents/${this.docId}/page-setup`, { pageSetup: { ...ps.pageSetup, chrome: ch } });
    }

    async setFooter(text, align) {
      const ps = await this.getPageSetup();
      const ch = ps.pageSetup && ps.pageSetup.chrome || {};
      if (text) ch.footer = { text, align: align || "center" };
      else delete ch.footer;
      return this._json("PUT", `/api/documents/${this.docId}/page-setup`, { pageSetup: { ...ps.pageSetup, chrome: ch } });
    }

    async setPageNumbers(opts) {
      const ps = await this.getPageSetup();
      const ch = ps.pageSetup && ps.pageSetup.chrome || {};
      if (opts && opts.enabled === false) delete ch.pageNumber;
      else ch.pageNumber = { enabled: true, format: (opts && opts.format) || "arabic", place: (opts && opts.place) || "footer-center" };
      return this._json("PUT", `/api/documents/${this.docId}/page-setup`, { pageSetup: { ...ps.pageSetup, chrome: ch } });
    }

    // ─── page setup / zoom ───
    async setPageSetup(o) {
      const id = await this._ensureDoc();
      await this._json("PUT", `/api/documents/${id}/page-setup`, { pageSetup: o, ...o });
      return { ok: true, pageSetup: (await this.getPageSetup()).pageSetup };
    }

    async getPageSetup() {
      const id = await this._ensureDoc();
      return this._json("GET", `/api/documents/${id}/page-setup`);
    }

    async setZoom(z) {
      // Zoom is a client-side property; stored as metadata for the embed scenario.
      return { ok: true, zoom: z };
    }

    async getZoom() {
      return { zoom: 1 };
    }

    // ─── versions ───
    async listVersions() {
      const id = await this._ensureDoc();
      const versions = await this._json("GET", `/api/documents/${id}/versions`);
      return { versions };
    }

    async restoreVersion(index) {
      const id = await this._ensureDoc();
      const data = await this._json("POST", `/api/documents/${id}/restore`, { index });
      return { ok: true, rev: data.rev };
    }

    // ─── find / replace ───
    async find(query, opts) {
      const content = await this.getContent();
      const html = content.html || "";
      const text = html.replace(/<[^>]*>/g, "").replace(/&nbsp;/g, " ").replace(/&amp;/g, "&");
      const flags = (opts && opts.matchCase) ? "g" : "gi";
      const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const re = opts && opts.regex ? new RegExp(query, flags) : new RegExp(escaped, flags);
      const matches = [];
      let m;
      while ((m = re.exec(text)) !== null) {
        matches.push({ index: m.index, text: m[0] });
        if (matches.length >= 1000) break;
      }
      return { matches };
    }

    async replaceAll(query, replacement, opts) {
      const id = await this._ensureDoc();
      const content = await this.getContent();
      const html = content.html || "";
      const flags = (opts && opts.matchCase) ? "g" : "gi";
      const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const re = opts && opts.regex ? new RegExp(query, flags) : new RegExp(escaped, "g");
      const replaced = html.replace(re, (match) => replacement.replace(/\$&/g, match));
      await this._json("PUT", `/api/documents/${id}/content`, { html: replaced });
      return { replaced: true };
    }

    // ─── export ───
    async exportDoc(fmt) {
      const id = await this._ensureDoc();
      const res = await fetch(`${this.baseUrl}/api/documents/${id}/export?fmt=${encodeURIComponent(fmt)}`);
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: `HTTP ${res.status}` }));
        throw new Error(err.error);
      }
      const blob = await res.blob();
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = `document.${fmt}`;
      a.click();
      URL.revokeObjectURL(a.href);
      return { ok: true };
    }

    async previewPrint() {
      throw new Error("previewPrint requires a browser; use exportDoc('pdf') for server-side");
    }

    // ─── lifecycle ───
    destroy() {
      this.docId = null;
    }
  }

  global.WordEditorRestClient = WordEditorRestClient;
})(window);
