"use strict";
/* The Visual Gallery test workbook (fixtures/gallery.twb: every chart type on one dashboard, two filter cards) exported
 * end to end against a mock of the Tableau API, and the file checked: its sheets, its build stamp, the visuals that
 * follow the slicers, and a package Excel opens without repairing. node --test tests/ */
const test = require("node:test");
const assert = require("node:assert/strict");
const { DOMParser } = require("./lib/shims.js");
const { tempXlsx, runExporter, inspect, xmlParts, currentBuild } = require("./lib/inspect.js");

const out = tempXlsx("gallery.xlsx");
let log = "", file = null;

test.before(async () => {
  log = runExporter("run-gallery.js", [out]);
  file = await inspect(out);
});

test("the export writes a file", () => {
  assert.match(log, /wrote /);
  assert.deepEqual(file.sheets, ["Visual Gallery", "Chart Data", "Slicer Data", "Conversion Report"]);
});

test("the Conversion Report carries the current build stamp", () => {
  assert.deepEqual(file.stamp, currentBuild());
});

test("the visuals follow the slicers (at least as many as before)", () => {
  assert.equal(file.visuals.length, 22);
  const follow = file.visuals.filter(v => v.follows).map(v => v.worksheet);
  assert.ok(follow.length >= 20, `${follow.length} of 22 follow: ${file.visuals.filter(v => !v.follows).map(v => v.worksheet).join(", ")}`);
  for (const name of ["KPI - Total Sales", "Text Table", "Highlight Table", "Packed Bubbles", "Filled Map"])
    assert.ok(follow.includes(name), `${name} follows the slicers`);
});

test("the package is well formed: no NaN, parseable XML, slicers on their tables", async () => {
  const parts = await xmlParts(file.zip);
  for (const [name, xml] of Object.entries(parts)) {
    assert.ok(!/>NaN<|"NaN"|\bNaN\b(?=[,)])/.test(xml), `${name} holds NaN (Excel would repair the file)`);
    const doc = new DOMParser().parseFromString(xml, "text/xml");
    assert.ok(doc && doc.documentElement, `${name} parses`);
  }
  const caches = Object.keys(parts).filter(n => /^xl\/slicerCaches\/slicerCache\d+\.xml$/.test(n));
  assert.equal(caches.length, 2, "a slicer cache per filter card");
  const tableIds = Object.keys(parts).filter(n => /^xl\/tables\/table\d+\.xml$/.test(n)).map(n => (parts[n].match(/<table\b[^>]*\sid="(\d+)"/) || [])[1]);
  for (const c of caches) {
    const id = (parts[c].match(/tableSlicerCache tableId="(\d+)"/) || [])[1];
    assert.ok(tableIds.includes(id), `${c} points at an existing table (${id})`);
  }
  const sheet = Object.entries(parts).find(([n, x]) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n) && /Slicer Data/.test(x));
  assert.ok(sheet, "the dashboard's values are formulas over the Slicer Data sheet");
});
