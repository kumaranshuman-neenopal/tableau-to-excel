"use strict";
/* Real workbooks with their packed data, end to end. Client files stay out of the repository: these tests run when
 * TFX_WORKBOOKS names a folder that holds them, and are skipped otherwise (CI). With TFX_EXCEL=1 on Windows with Excel,
 * the exported formulas are also recalculated in Excel and compared with numbers computed from the workbook's own data.
 *   TFX_WORKBOOKS=C:\Users\me\Downloads node --test tests/
 * Ecommerce_Dashboard.twbx – a relationships data source of four CSV tables, 14 filter actions, 4 dashboards;
 * Untitled_17.twbx – one Excel sheet of 51 rows, every measure a ROUND(...) calculation. */
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");
const { tempXlsx, runExporter, inspect } = require("./lib/inspect.js");

const DIR = process.env.TFX_WORKBOOKS || "";
const ECOM = path.join(DIR, "Ecommerce_Dashboard.twbx");
const LOGI = path.join(DIR, "Untitled_17.twbx");
const has = f => !!DIR && fs.existsSync(f);
const EXCEL = process.env.TFX_EXCEL === "1" && process.platform === "win32";
const skipEcom = has(ECOM) ? false : "TFX_WORKBOOKS has no Ecommerce_Dashboard.twbx";
const skipLogi = has(LOGI) ? false : "TFX_WORKBOOKS has no Untitled_17.twbx";

/** the workbook exported (one dashboard, or ALL) → { log, file } */
async function exportOf(twbx, dashboard, env, name) {
  const out = tempXlsx(name);
  const log = runExporter("export-twbx.js", [twbx, dashboard, out], env);
  return { log, out, file: await inspect(out) };
}

/** Excel: the dropdown holding value set to it; the first formula cells over the slicer data before and after */
function excel(out, value, count = 4) {
  const ps = path.join(__dirname, "excel", "set-and-read.ps1");
  const json = execFileSync("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps, "-Xlsx", out, "-Value", value, "-Count", String(count)],
                            { encoding: "utf8", timeout: 300000 });
  return JSON.parse(json.trim().split(/\r?\n/).pop());
}

test("Ecommerce Sales Overview: related tables joined, every visual follows, an action slicer per chart", { skip: skipEcom }, async () => {
  const { log, file } = await exportOf(ECOM, "Sales Overview", { REALCSV: "1" }, "sales.xlsx");
  assert.match(log, /4 related tables joined on "order_items\.csv" – 10000 rows/);
  assert.equal(file.visuals.length, 8);
  assert.deepEqual(file.visuals.filter(v => !v.follows).map(v => v.worksheet), []);
  assert.equal(Object.keys(file.zip.files).filter(n => /slicerCache\d+\.xml$/.test(n)).length, 8, "8 action slicers");
  assert.ok(file.sheets.includes("Slicer Actions"));
});

test("Ecommerce exported at Channel = Website: the pie gets the other channels", { skip: skipEcom }, async () => {
  const { log } = await exportOf(ECOM, "Sales Overview", { REALCSV: "1", SELECT: "Channel=Website" }, "website.xlsx");
  assert.match(log, /\[Every value\] Sales Overview: Channel cleared/);
  assert.match(log, /\[Every value\] Revenue Share by Channel \(Pie\): 3 more slice\(s\)/);
});

test("Ecommerce Detail Analysis: the packed bubbles follow the slicers", { skip: skipEcom }, async () => {
  const { file } = await exportOf(ECOM, "Detail Analysis", { REALCSV: "1" }, "detail.xlsx");
  const bubbles = file.visuals.find(v => /Bubbles/.test(v.worksheet));
  assert.ok(bubbles && bubbles.follows, bubbles && bubbles.note);
});

test("Ecommerce, all four dashboards: a slicer table each, 13 action slicers", { skip: skipEcom }, async () => {
  const { file } = await exportOf(ECOM, "ALL", { REALCSV: "1" }, "all.xlsx");
  for (const s of ["KPI Overview", "Sales Overview", "Gauge & Donut", "Detail Analysis", "Slicer Data", "Slicer Data 2", "Slicer Data 3", "Slicer Data 4"])
    assert.ok(file.sheets.includes(s), `sheet ${s}`);
  assert.equal(Object.keys(file.zip.files).filter(n => /slicerCache\d+\.xml$/.test(n)).length, 13);
});

test("Ecommerce KPI Overview in Excel: Category = Electronics gives the CSV's numbers", { skip: skipEcom || (EXCEL ? false : "set TFX_EXCEL=1 (Windows with Excel)") }, async () => {
  const { out } = await exportOf(ECOM, "KPI Overview", { REALCSV: "1" }, "kpi.xlsx");
  const r = excel(out, "Electronics", 6);
  assert.equal(r.repaired, false, "opens without repair");
  assert.ok(r.dropdown, "the Category dropdown");
  // Total Revenue, Orders, Avg Order Value, Profit Margin, Total Profit, Units Sold – from order_items / orders / products
  // (Avg Order Value = revenue / orders, Profit Margin = profit / revenue)
  assert.deepEqual(r.before, [3166913.17, 4581, 691.31, 0.3, 941075.06, 20087]);
  assert.deepEqual(r.after, [1735482.52, 1586, 1094.25, 0.2, 350239.31, 3699]);
});

test("Logistics (every measure a ROUND calculation): the visuals follow the filters", { skip: skipLogi }, async () => {
  const { file } = await exportOf(LOGI, "Logistics & Supply Chain Operations Command Centre", { REALXLSX: "1" }, "logistics.xlsx");
  const follow = file.visuals.filter(v => v.follows).length;
  assert.ok(follow >= 22, `${follow} of ${file.visuals.length} follow`);
});

test("Logistics without the formula path: Tableau-computed values per dropdown value", { skip: skipLogi }, async () => {
  const { log, file } = await exportOf(LOGI, "Logistics & Supply Chain Operations Command Centre", { REALXLSX: "1", NOSLICERS: "1" }, "states.xlsx");
  assert.match(log, /\[States\] .*: 24 filter values/);
  assert.ok(file.sheets.some(s => /^States /.test(s)), "the hidden States sheet");
});

test("Logistics in Excel: Transport Mode = Air gives the data's numbers", { skip: skipLogi || (EXCEL ? false : "set TFX_EXCEL=1 (Windows with Excel)") }, async () => {
  const { out } = await exportOf(LOGI, "Logistics & Supply Chain Operations Command Centre", { REALXLSX: "1" }, "air.xlsx");
  const r = excel(out, "Air");
  assert.equal(r.repaired, false, "opens without repair");
  assert.ok(r.dropdown, "the Transport Mode dropdown");
  assert.ok(r.after.length > 0, "formula cells read");
  assert.notDeepEqual(r.after, r.before, "the values follow the dropdown");
});
