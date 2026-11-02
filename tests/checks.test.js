"use strict";
/* The panel's checks before an export (ui/checks.js), on the Visual Gallery dashboard: its Region card keeps three of
 * four regions, and a tiled extension object is added to it. node --test tests/ */
const test = require("node:test");
const assert = require("node:assert/strict");
const { load } = require("./lib/load.js");
const g = require("./lib/gallery.js");

const T = load({ dashboard: {}, captureDownload: () => {}, expose: ["applyWorkbookXml", "preExportChecks", "exportSummaryChecks"] });
console.log = console.info = console.warn = () => {};

let dashboard;
test.before(async () => {
  const { formatModel } = await T.applyWorkbookXml(g.workbookXml(), "Visual Gallery.twb", false, null);
  dashboard = g.dashboardFrom(formatModel);
  tableau.extensions.dashboardContent.dashboard = dashboard;
});

test("a card set to some values: every value goes into the file", async () => {
  const lines = await T.preExportChecks(dashboard);
  const set = lines.find(l => /Region = Central, East, West/.test(l.text));
  assert.ok(set, lines.map(l => l.text).join(" | "));
  assert.equal(set.level, "info");
  assert.match(set.text, /every value goes into the file/);
  assert.ok(lines.some(l => l.level === "ok" && /Ready to export/.test(l.text)));
});

test("a tiled extension squeezes the dashboard: a warning", async () => {
  const withExtension = { ...dashboard, objects: [...dashboard.objects,
    { id: 999, type: "extension", name: "Export2Sheet", isVisible: true, isFloating: false, position: { x: 0, y: 900 }, size: { width: 400, height: 120 } }] };
  const lines = await T.preExportChecks(withExtension);
  const tiled = lines.find(l => /extension is tiled/.test(l.text));
  assert.ok(tiled && tiled.level === "warn");
  assert.ok(!lines.some(l => /Ready to export/.test(l.text)));
  const floating = { ...withExtension, objects: withExtension.objects.map(o => (o.id === 999 ? { ...o, isFloating: true } : o)) };
  assert.ok(!(await T.preExportChecks(floating)).some(l => /extension is tiled/.test(l.text)));
});

test("after an export: how many visuals follow, and why the others do not", () => {
  const written = [{ dashboardName: "D", report: { slicers: [
    { worksheet: "A", follows: true, note: "follows the slicers" },
    { worksheet: "B", follows: false, note: "fixed values: a table calculation" }] } }];
  const lines = T.exportSummaryChecks(written);
  assert.equal(lines[0].text, "1 of 2 visuals follow the filters in Excel.");
  assert.equal(lines[1].text, "B keeps Tableau's values: a table calculation");
});
