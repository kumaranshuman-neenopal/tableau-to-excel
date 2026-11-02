"use strict";
/* Reading an exported .xlsx in tests: its sheets, its Conversion Report (which visuals follow the slicers, the build
 * stamp), its package parts – and running the test exporters (run-gallery.js, export-twbx.js) as child processes. */
const fs = require("fs");
const os = require("os");
const path = require("path");
const { execFileSync } = require("child_process");

const ROOT = path.resolve(__dirname, "../..");
const ExcelJS = require(path.join(ROOT, "js/exceljs.min.js"));
const JSZip = require(path.join(ROOT, "js/jszip.min.js"));

/** a fresh temporary file path for an export */
function tempXlsx(name) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "tfx-test-"));
  return path.join(dir, name);
}

/** runs a test exporter; returns its printed log (throws with the log when it fails) */
function runExporter(script, args, env = {}) {
  try {
    return execFileSync(process.execPath, [path.join(__dirname, script), ...args],
      { env: { ...process.env, ...env }, encoding: "utf8", stdio: ["ignore", "pipe", "pipe"], maxBuffer: 64 * 1024 * 1024, timeout: 600000 });
  } catch (e) {
    throw new Error(`${script} failed:\n${e.stdout || ""}\n${e.stderr || ""}`);
  }
}

/** the export's sheets, report rows and package: { sheets, stamp, visuals: [{ worksheet, follows, note }], zip } */
async function inspect(file) {
  const bytes = fs.readFileSync(file);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes);
  const report = wb.getWorksheet("Conversion Report");
  const rows = [];
  if (report) report.eachRow(r => rows.push(r.values.slice(1).map(v => (v && typeof v === "object" && "richText" in v ? v.richText.map(x => x.text).join("") : v))));
  const text = rows.map(r => r.filter(v => v !== null && v !== undefined).join(" | "));
  const stampLine = text.find(t => /Export to Excel v/.test(t)) || "";
  const stamp = (stampLine.match(/Export to Excel v([\d.]+) \(([^)]*)\)/) || []).slice(1);
  // "Slicers: the visuals they filter": worksheet | Follows the slicers / Fixed values | note (merged cells repeat
  // their value across the row: the cells before and after the status give the worksheet and the note)
  const visuals = rows.map(r => {
    const at = r.findIndex(v => v === "Follows the slicers" || v === "Fixed values");
    if (at < 0) return null;
    const before = r.slice(0, at).filter(v => v !== null && v !== undefined);
    const after = r.slice(at + 1).filter(v => v !== null && v !== undefined);
    return { worksheet: String(before[before.length - 1]), follows: r[at] === "Follows the slicers", note: String(after[0] || "") };
  }).filter(Boolean);
  const zip = await JSZip.loadAsync(bytes);
  return { sheets: wb.worksheets.map(w => w.name), stamp: { version: stamp[0], build: stamp[1] }, visuals, text, zip };
}

/** every XML part of the package, as text */
async function xmlParts(zip) {
  const out = {};
  for (const n of Object.keys(zip.files).filter(n => /\.(xml|rels|vml)$/.test(n))) out[n] = await zip.file(n).async("string");
  return out;
}

/** the build stamp tableau-to-excel.js carries now */
function currentBuild() {
  const js = fs.readFileSync(path.join(ROOT, "tableau-to-excel.js"), "utf8");
  return { version: js.match(/const EXTENSION_VERSION = "([^"]*)"/)[1], build: js.match(/const EXTENSION_BUILD = "([^"]*)"/)[1] };
}

module.exports = { ROOT, tempXlsx, runExporter, inspect, xmlParts, currentBuild };
