"use strict";
/* Loads tableau-to-excel.js into this Node process with just enough browser + Tableau globals, and returns its
 * internals (the file is one closure: a hook before its last "})();" exposes the names asked for). */
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const { DOMParser } = require("./shims.js");

const ROOT = process.env.EXT_ROOT || path.resolve(__dirname, "../..");

/** a DOM element that accepts anything */
function fakeElement(id) {
  const el = {
    id, style: {}, dataset: {}, children: [], textContent: "", innerHTML: "", title: "", value: "", checked: false, disabled: false,
    classList: { add() {}, remove() {}, toggle() {}, contains: () => false },
    setAttribute() {}, getAttribute: () => null, addEventListener() {}, removeEventListener() {},
    appendChild(c) { this.children.push(c); return c; }, append() {}, prepend() {}, remove() {}, focus() {}, click() { if (el.onclick) el.onclick(); },
    querySelector: () => fakeElement(), querySelectorAll: () => [], closest: () => null, insertBefore() {}, insertAdjacentElement() {}
  };
  return el;
}

function load({ dashboard, captureDownload, expose, patch }) {
  const elements = new Map();
  const byId = id => {
    if (!elements.has(id)) elements.set(id, fakeElement(id));
    return elements.get(id);
  };
  // the panel's elements: the export button and status line exist, everything else is absent
  ["export_button", "export_status"].forEach(byId);
  globalThis.window = globalThis;
  globalThis.DOMParser = DOMParser;
  globalThis.document = {
    addEventListener() {}, getElementById: id => elements.get(id) || null, createElement: tag => fakeElement(tag),
    head: fakeElement("head"), body: fakeElement("body"), querySelector: () => null, querySelectorAll: () => []
  };
  globalThis.alert = msg => console.error("[alert]", msg);
  globalThis.confirm = () => true;
  globalThis.ExcelJS = require(path.join(ROOT, "js/exceljs.min.js"));
  globalThis.JSZip = require(path.join(ROOT, "js/jszip.min.js"));
  const settings = new Map();
  globalThis.tableau = { extensions: {
    settings: { get: k => settings.get(k), set: (k, v) => settings.set(k, v), erase: k => settings.delete(k), saveAsync: async () => {} },
    environment: { context: "desktop" },
    dashboardContent: { dashboard },
    initializeAsync: () => new Promise(() => {})
  } };
  // downloads: <a>.click() on a blob URL → captureDownload(bytes)
  const blobs = new Map();
  globalThis.URL.createObjectURL = blob => { const u = "blob:" + blobs.size; blobs.set(u, blob); return u; };
  globalThis.URL.revokeObjectURL = () => {};
  const create = globalThis.document.createElement;
  globalThis.document.createElement = tag => {
    const el = create(tag);
    if (tag === "a") el.click = () => { const b = blobs.get(el.href); if (b) b.arrayBuffer().then(buf => captureDownload(new Uint8Array(buf), el.download)); };
    return el;
  };

  let src = fs.readFileSync(path.join(ROOT, "tableau-to-excel.js"), "utf8");
  (patch || []).forEach(([a, b]) => { if (!src.includes(a)) throw new Error("patch not found: " + a); src = src.replace(a, b); });
  const end = src.lastIndexOf("})();");
  if (end < 0) throw new Error("closure end not found");
  const names = expose.filter(n => new RegExp(`(function|const|let|class)\\s+${n}\\b`).test(src));
  src = src.slice(0, end) + `\nglobalThis.__T = { ${names.join(", ")} };\n` + src.slice(end);
  vm.runInThisContext(src, { filename: "tableau-to-excel.js" });
  return globalThis.__T;
}

module.exports = { load, ROOT };
