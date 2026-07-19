// history.js — custom undo/redo stack for the contentEditable editor.
// Snapshots pair editor.innerHTML (cleaned of transient find highlights) with
// a serialized selection so undo/redo restores both content and caret.
// Coalesces consecutive typing/deleting within a short window into one step,
// matching the granularity users expect from word processors.

// ---- selection serialization ----------------------------------------------
function nodePath(root, node) {
  const path = [];
  let cur = node;
  while (cur && cur !== root) {
    const parent = cur.parentNode;
    if (!parent) return null;
    let idx = 0;
    for (let i = 0; i < parent.childNodes.length; i++) {
      if (parent.childNodes[i] === cur) { idx = i; break; }
    }
    path.unshift(idx);
    cur = parent;
  }
  return cur === root ? path : null;
}

function nodeFromPath(root, path) {
  let cur = root;
  for (const idx of path) {
    if (!cur || !cur.childNodes || idx < 0 || idx >= cur.childNodes.length) return null;
    cur = cur.childNodes[idx];
  }
  return cur;
}

function getSelectionSnapshot(editor) {
  const sel = window.getSelection();
  if (!sel || !sel.rangeCount) return null;
  const r = sel.getRangeAt(0);
  if (!editor.contains(r.startContainer) || !editor.contains(r.endContainer)) return null;
  const sp = nodePath(editor, r.startContainer);
  const ep = nodePath(editor, r.endContainer);
  if (!sp || !ep) return null;
  return {
    start: { path: sp, offset: r.startOffset },
    end: { path: ep, offset: r.endOffset },
  };
}

function applySelectionSnapshot(editor, snap) {
  if (!snap) return;
  const sNode = nodeFromPath(editor, snap.start.path);
  const eNode = nodeFromPath(editor, snap.end.path);
  if (!sNode || !eNode) return;
  const sel = window.getSelection();
  if (!sel) return;
  const sMax = sNode.nodeType === 3 ? sNode.length || 0 : sNode.childNodes.length;
  const eMax = eNode.nodeType === 3 ? eNode.length || 0 : eNode.childNodes.length;
  const r = document.createRange();
  try {
    r.setStart(sNode, Math.min(snap.start.offset, sMax));
    r.setEnd(eNode, Math.min(snap.end.offset, eMax));
    sel.removeAllRanges();
    sel.addRange(r);
  } catch {}
}

// ---- History ---------------------------------------------------------------
const INSERT = new Set(["insertText", "insertCompositionText", "insertReplacementText"]);
const DELETE = new Set(["deleteContentBackward", "deleteContentForward", "deleteByComposition", "deleteByCut"]);

export class History {
  constructor(editor, opts = {}) {
    this.editor = editor;
    this.undoStack = [];
    this.redoStack = [];
    this.present = null;
    this.suspended = false;
    this.coalescing = false;
    this.lastType = null;
    this.lastTime = 0;
    this.lastWasSpace = false;
    this.onUpdate = null; // (canUndo, canRedo) => void
    this.coalesceMs = opts.coalesceMs != null ? opts.coalesceMs : 1200;
    this.limit = opts.limit || 500;
  }

  attach() {
    this._onInput = (e) => this._handleInput(e);
    this.editor.addEventListener("input", this._onInput);
    this.reset();
  }

  detach() {
    if (this._onInput) this.editor.removeEventListener("input", this._onInput);
  }

  // innerHTML with transient find-highlight marks unwrapped, so snapshots
  // never carry ephemeral UI state.
  _cleanHtml() {
    const clone = this.editor.cloneNode(true);
    for (const m of clone.querySelectorAll("mark.find-hit")) {
      const p = m.parentNode;
      if (!p) continue;
      while (m.firstChild) p.insertBefore(m.firstChild, m);
      p.removeChild(m);
      p.normalize();
    }
    return clone.innerHTML;
  }

  _snapshot() {
    return { html: this._cleanHtml(), sel: getSelectionSnapshot(this.editor) };
  }

  _notify() {
    if (this.onUpdate) try { this.onUpdate(this.canUndo(), this.canRedo()); } catch {}
  }

  // Re-baseline without recording history (used on load/import/remote follow).
  reset() {
    this.undoStack = [];
    this.redoStack = [];
    this.present = this._snapshot();
    this.coalescing = false;
    this.lastTime = 0;
    this.lastWasSpace = false;
    this._notify();
  }

  canUndo() { return this.undoStack.length > 0; }
  canRedo() { return this.redoStack.length > 0; }

  // Word-level coalescing: consecutive typing merges, with a new group
  // starting at the beginning of each word (a non-space after a space, or a
  // category switch between inserting and deleting). Whitespace attaches to
  // the preceding word so undo steps back one word plus its trailing space.
  _shouldCoalesce(type, data) {
    if (!type) return false;
    const isIns = INSERT.has(type);
    const isDel = DELETE.has(type);
    if (!isIns && !isDel) return false;
    if (isIns && DELETE.has(this.lastType)) return false;
    if (isDel && INSERT.has(this.lastType)) return false;
    if (isIns) {
      if (/\s/.test(data || "")) return true;        // space joins current word
      return !this.lastWasSpace;                       // new word starts fresh group
    }
    return true;                                       // deletes merge together
  }

  _handleInput(ev) {
    if (this.suspended) return;
    const type = ev && ev.inputType;
    const data = ev && ev.data;
    const now = Date.now();
    const snap = this._snapshot();

    if (this.present && snap.html === this.present.html) {
      // caret-only change (no content delta): keep caret fresh, record nothing
      this.present.sel = snap.sel;
      return;
    }

    const withinWindow = (now - this.lastTime) < this.coalesceMs;
    const coalesce = this.coalescing && withinWindow && this._shouldCoalesce(type, data);

    if (coalesce) {
      this.present = snap; // merge: before-state already on the undo stack
    } else {
      this.undoStack.push(this.present);
      this.present = snap;
      this.redoStack = [];
      while (this.undoStack.length > this.limit) this.undoStack.shift();
    }

    this.coalescing = INSERT.has(type) || DELETE.has(type);
    this.lastWasSpace = INSERT.has(type) ? /\s/.test(data || "") : DELETE.has(type) ? false : this.lastWasSpace;
    this.lastType = type;
    this.lastTime = now;
    this._notify();
  }

  _apply(snap) {
    this.suspended = true;
    try {
      this.editor.innerHTML = snap.html;
      applySelectionSnapshot(this.editor, snap.sel);
      // let the app react (autosave, paginate, word count) without re-recording
      this.editor.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "historyApply" }));
    } finally {
      this.suspended = false;
    }
    this._notify();
  }

  undo() {
    if (!this.undoStack.length) return;
    this.redoStack.push(this.present);
    this.present = this.undoStack.pop();
    this.coalescing = false;
    this.lastWasSpace = false;
    this._apply(this.present);
  }

  redo() {
    if (!this.redoStack.length) return;
    this.undoStack.push(this.present);
    this.present = this.redoStack.pop();
    this.coalescing = false;
    this.lastWasSpace = false;
    this._apply(this.present);
  }
}
