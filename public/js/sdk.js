/**
 * DocEditor SDK — embed the Word-Compatible Editor in any web page.
 *
 * Usage:
 *   <script src="http://your-editor-host/js/sdk.js"></script>
 *   <script>
 *     const ed = DocEditor.init({
 *       container: "#editor-holder",        // element or selector
 *       baseUrl: "http://localhost:3001",   // editor server (defaults to script origin)
 *       docId: "…",                          // optional: open an existing document
 *       mode: "edit",                        // "edit" | "view"
 *       toolbar: true,
 *       statusbar: true,
 *       user: "Alice",                       // presence name
 *       onReady(info)   {},
 *       onChange(info)  {},
 *       onSave(info)    {},
 *       onPresence(p)   {},
 *       onError(err)    {},
 *     });
 *     await ed.getContent();     // -> {html}
 *     await ed.setContent("<p>Hello</p>");
 *     await ed.insertText("…"); await ed.insertHtml("<b>…</b>");
 *     await ed.getText();        // -> {text}
 *     await ed.getMeta();        // -> {id, title, rev, words, chars, pageSetup, trackChanges, commentCount}
 *     await ed.setTitle("Report");
 *     await ed.save();
 *     await ed.find("term", {matchCase:false, regex:false});
 *     await ed.replaceAll("a", "b", {});
 *     await ed.setMode("view");
 *     await ed.loadDocument(id);
 *     // review features:
 *     await ed.addComment("Please rephrase");        // anchors on current selection
 *     await ed.getComments();                        // -> {comments}
 *     await ed.setTrackChanges(true);
 *     await ed.getChanges();                         // -> {changes}
 *     await ed.acceptAllChanges(); await ed.rejectAllChanges();
 *   // ---- formatting (execCommand bridge) ----
 *   await ed.bold(); await ed.italic(); await ed.underline();
 *   await ed.formatBlock("H1"); await ed.foreColor("#ff0000");
 *   await ed.justifyCenter(); await ed.insertUnorderedList();
 *   // ---- inserts ----
 *   await ed.insertImage(url, { alt, width, height });
 *   await ed.insertLink(href, text); await ed.insertTable(3, 4);
 *   await ed.insertPageBreak(); await ed.insertHr(); await ed.insertSymbol("§");
 *   // ---- headers/footers/page numbers ----
 *   await ed.setHeader("Chapter 1", "right"); await ed.setPageNumbers({format:"roman", place:"footer-center"});
 *   // ---- page setup / zoom ----
 *   await ed.setPageSetup({ size: "A4", margins: { top: 1, bottom: 1 } });
 *   await ed.setZoom(1.25);
 *   // ---- library / versions ----
 *   await ed.listDocuments(); await ed.newDocument("Memo"); await ed.restoreVersion(0);
 *   // ---- export ----
 *   await ed.exportDoc("docx");  // "docx" | "pdf" | "html" | "txt"
 *   ed.destroy();
 * }
 */
(function (global) {
  "use strict";

  const scriptOrigin = (() => {
    try {
      const s = document.currentScript;
      return s ? new URL(s.src).origin : location.origin;
    } catch { return location.origin; }
  })();

  function init(opts = {}) {
    const container = typeof opts.container === "string"
      ? document.querySelector(opts.container)
      : opts.container;
    if (!container) throw new Error("DocEditor.init: container not found");

    const baseUrl = (opts.baseUrl || scriptOrigin).replace(/\/$/, "");
    const q = new URLSearchParams();
    q.set("embed", "1");
    if (opts.docId) q.set("doc", opts.docId);
    if (opts.mode === "view") q.set("mode", "view");
    if (opts.toolbar === false) q.set("toolbar", "0");
    if (opts.statusbar === false) q.set("statusbar", "0");
    if (opts.user) q.set("user", opts.user);

    const iframe = document.createElement("iframe");
    iframe.src = `${baseUrl}/?${q.toString()}`;
    iframe.style.cssText = "width:100%;height:100%;border:none;display:block;min-height:400px";
    iframe.allow = "clipboard-read; clipboard-write";
    container.appendChild(iframe);

    const pending = new Map();
    let seq = 0;
    let destroyed = false;

    function onMessage(ev) {
      if (destroyed) return;
      if (ev.source !== iframe.contentWindow) return;
      const m = ev.data;
      if (!m || m.we !== 1) return;
      if (m.re != null && pending.has(m.re)) {
        const { resolve, reject } = pending.get(m.re);
        pending.delete(m.re);
        m.error ? reject(new Error(m.error)) : resolve(m.result);
        return;
      }
      if (m.event) {
        const handler = {
          ready: opts.onReady, change: opts.onChange, save: opts.onSave,
          presence: opts.onPresence, error: opts.onError,
        }[m.event];
        if (handler) try { handler(m.data); } catch (e) { console.error(e); }
      }
    }
    window.addEventListener("message", onMessage);

    function call(cmd, args) {
      if (destroyed) return Promise.reject(new Error("editor destroyed"));
      return new Promise((resolve, reject) => {
        const id = ++seq;
        pending.set(id, { resolve, reject });
        iframe.contentWindow.postMessage({ we: 1, id, cmd, args }, "*");
        setTimeout(() => {
          if (pending.has(id)) {
            pending.delete(id);
            reject(new Error("DocEditor: timeout for " + cmd));
          }
        }, 15000);
      });
    }

    return {
      iframe,
      // ---- content ----
      getContent: () => call("getContent"),
      getText: () => call("getText"),
      setContent: (html) => call("setContent", { html }),
      insertText: (text) => call("insertText", { text }),
      insertHtml: (html) => call("insertHtml", { html }),
      getMeta: () => call("getMeta"),
      setTitle: (title) => call("setTitle", { title }),
      save: () => call("save"),
      loadDocument: (id) => call("loadDocument", { id }),
      setMode: (mode) => call("setMode", { mode }),
      find: (query, o) => call("find", { query, ...o }),
      replaceAll: (query, replacement, o) => call("replaceAll", { query, replacement, ...o }),
      focus: () => call("focus"),
      undo: () => call("undo"),
      redo: () => call("redo"),
      canUndo: () => call("canUndo"),
      // ---- formatting ----
      format: (cmd, value) => call("format", { cmd, value }),
      bold: () => call("format", { cmd: "bold" }),
      italic: () => call("format", { cmd: "italic" }),
      underline: () => call("format", { cmd: "underline" }),
      strikeThrough: () => call("format", { cmd: "strikeThrough" }),
      subscript: () => call("format", { cmd: "subscript" }),
      superscript: () => call("format", { cmd: "superscript" }),
      justifyLeft: () => call("format", { cmd: "justifyLeft" }),
      justifyCenter: () => call("format", { cmd: "justifyCenter" }),
      justifyRight: () => call("format", { cmd: "justifyRight" }),
      justifyFull: () => call("format", { cmd: "justifyFull" }),
      insertOrderedList: () => call("format", { cmd: "insertOrderedList" }),
      insertUnorderedList: () => call("format", { cmd: "insertUnorderedList" }),
      outdent: () => call("format", { cmd: "outdent" }),
      indent: () => call("format", { cmd: "indent" }),
      formatBlock: (tag) => call("format", { cmd: "formatBlock", value: tag }),
      foreColor: (color) => call("format", { cmd: "foreColor", value: color }),
      hiliteColor: (color) => call("format", { cmd: "hiliteColor", value: color }),
      fontName: (font) => call("format", { cmd: "fontName", value: font }),
      fontSize: (size) => call("format", { cmd: "fontSize", value: size }),
      getSelectedText: () => call("getSelectedText"),
      // ---- inserts ----
      insertImage: (src, o) => call("insertImage", { src, ...o }),
      insertLink: (href, text) => call("insertLink", { href, text }),
      insertTable: (rows, cols) => call("insertTable", { rows, cols }),
      insertSymbol: (ch) => call("insertSymbol", { char: ch }),
      insertPageBreak: () => call("insertPageBreak"),
      insertBlankPage: () => call("insertBlankPage"),
      insertHr: () => call("insertHr"),
      // ---- headers / footers / page numbers ----
      setHeader: (text, align) => call("setHeader", { text, align }),
      setFooter: (text, align) => call("setFooter", { text, align }),
      setPageNumbers: (o) => call("setPageNumbers", o),
      // ---- page setup / zoom ----
      setPageSetup: (o) => call("setPageSetup", o),
      getPageSetup: () => call("getPageSetup"),
      setZoom: (z) => call("setZoom", { zoom: z }),
      getZoom: () => call("getZoom"),
      // ---- library / versions ----
      listDocuments: () => call("listDocuments"),
      newDocument: (title) => call("newDocument", { title }),
      deleteDocument: (id) => call("deleteDocument", { id }),
      listVersions: () => call("listVersions"),
      restoreVersion: (index) => call("restoreVersion", { index }),
      // ---- export ----
      exportDoc: (fmt) => call("export", { fmt }),
      previewPrint: () => call("previewPrint"),
      // ---- review ----
      addComment: (text, author) => call("addComment", { text, author }),
      getComments: () => call("getComments"),
      setTrackChanges: (on) => call("setTrackChanges", { on }),
      getChanges: () => call("getChanges"),
      acceptAllChanges: () => call("acceptAllChanges"),
      rejectAllChanges: () => call("rejectAllChanges"),
      destroy() {
        destroyed = true;
        window.removeEventListener("message", onMessage);
        iframe.remove();
        pending.clear();
      },
    };
  }

  global.DocEditor = { init, version: "3.0.0" };
})(window);
