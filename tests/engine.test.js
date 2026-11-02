"use strict";
/* The slicer engine without a workbook: calculations of aggregates translated to Excel and evaluated as Tableau does,
 * and every kind of filter the Extensions API reports turned into a slicer / dropdown / range – or the reason it stays a
 * card – with the rows each keeps. node --test tests/ */
const test = require("node:test");
const assert = require("node:assert/strict");
const { load } = require("./lib/load.js");

const T = load({ dashboard: {}, captureDownload: () => {},
  expose: ["createSlicerModel", "tfParseFieldRef", "tfIsNull", "slicerParseCalc", "slicerEval", "slicerExcel", "slicerDatePart", "slicerMarkText"] });
console.log = console.info = () => {};

const dv = (v, f) => ({ value: v, nativeValue: v, formattedValue: f === undefined ? String(v) : f });
const nul = { value: null, nativeValue: null, formattedValue: "Null" };
const nan = { value: NaN, nativeValue: NaN, formattedValue: "Null" };

test("a range card's open end (NaN) counts as no value", () => {
  assert.equal(T.tfIsNull(nan), true);
  assert.equal(T.tfIsNull(dv(0)), false);
});

test("calculations of aggregates: functions, conditions, parameters", () => {
  const src = { id: "d", name: "d", columns: [{ name: "Sales", dataType: "float" }, { name: "Discount", dataType: "float" }],
                rows: [[100, 0.1], [250, 0.3], [-40, 0]].map(r => r.map(v => dv(v))) };
  const model = { fields: { "Parameters|parameter 3": { param: true, value: "0.15" } } };
  const R = { col: c => "C" + c, key: "K", vis: "V" };
  const cases = [
    ["ROUND(SUM([Sales])/3, 1)", 103.3, "ROUND((SUMIFS(C0,K,1,V,1)/3),1)"],
    ["ROUND(-2.5)", -3, null],                                                 // half away from zero, as Tableau and Excel
    ["IIF(AVG([Discount]) >= [Parameters].[Parameter 3], 1, 0)", 0, "IF((AVERAGEIFS(C1,K,1,V,1)>=0.15),1,0)"],
    ["IF SUM([Sales]) > 200 THEN ABS(MIN([Sales])) ELSEIF SUM([Sales]) > 0 THEN 2 ELSE 3 END", 40, null],
    ["IFNULL(MAX([Sales]), 0)", 250, "IFERROR(_xlfn.MAXIFS(C0,K,1,V,1),0)"],
    ["MAX(SUM([Sales]), 500)", 500, "MAX(SUMIFS(C0,K,1,V,1),500)"],
    ["DIV(INT(SUM([Sales])), 7)", 44, null],
    ["SUM([Sales]) / COUNT([Sales])", 310 / 3, null]
  ];
  for (const [formula, value, excel] of cases) {
    const e = T.slicerParseCalc(src, model, formula, 0);
    assert.ok(e, `parsed: ${formula}`);
    assert.ok(Math.abs(T.slicerEval(e, [0, 1, 2], src, null) - value) < 1e-9, `value of ${formula}`);
    if (excel) assert.equal(T.slicerExcel(e, R, "1"), excel, `Excel formula of ${formula}`);
  }
});

test("weekday and week levels: Tableau's weeks (Sunday first, week 1 holds 1 January)", () => {
  const d = (y, m, dd) => ({ y, m, d: dd });
  assert.deepEqual(["wd", "wk", "twk"].map(p => T.slicerDatePart(d(2024, 1, 1), p)), ["Monday", "1", "2023-12-31"]);
  assert.deepEqual(["wd", "wk", "twk"].map(p => T.slicerDatePart(d(2024, 1, 7), p)), ["Sunday", "2", "2024-01-07"]);
  assert.equal(T.slicerDatePart(d(2025, 12, 31), "wk"), "53");
  // the summary data's WEEKDAY: a name or 1–7 (1 = Sunday) – the same key as the rows'
  assert.equal(T.slicerMarkText(dv("Monday"), "wd", true), "Monday");
  assert.equal(T.slicerMarkText(dv(2, "Monday"), "wd", true), "Monday");
  assert.equal(T.slicerMarkText(dv(1), "wd", true), "Sunday");
  assert.equal(T.slicerMarkText(dv(14), "wk", true), "14");
});

/* ── filter cards: a little data source of 48 rows ───────────────────────────────────────────────────────────── */
const REGIONS = ["Central", "East", "South", "West"];
const rows = [];
for (let i = 0; i < 48; i++) {
  const d = new Date(Date.UTC(2024 + (i % 2), i % 12, 1 + (i % 28)));
  rows.push({ region: i === 47 ? null : REGIONS[i % 4], date: d, sales: 10 * i, flag: i % 3 === 0, year: 2020 + (i % 5), dt: new Date(d.getTime() + 13 * 3600e3) });
}
const COLS = [["Region", "string", r => r.region], ["Order Date", "date", r => r.date], ["Sales", "float", r => r.sales],
              ["Is Big", "bool", r => r.flag], ["Fiscal Year", "int", r => r.year], ["Ship Time", "date-time", r => r.dt]];
const cell = (t, v) => v === null ? nul : v instanceof Date ? { value: v.toISOString().slice(0, t === "date" ? 10 : 19), nativeValue: v, formattedValue: v.toISOString().slice(0, 10) } : dv(v);
const ds = { id: "ds1", name: "Orders", fields: COLS.map(([n]) => ({ name: n })),
  getLogicalTablesAsync: async () => [{ id: "t", caption: "Orders" }],
  getLogicalTableDataAsync: async () => ({ columns: COLS.map(([n, t], i) => ({ fieldName: n, dataType: t, index: i })),
    data: rows.map(r => COLS.map(([, t, f]) => cell(t, f(r)))), totalRowCount: rows.length }) };
const ws = { name: "S", getDataSourcesAsync: async () => [ds] };
const model = { fields: {}, sheets: {}, dashboards: {} };
const ref = s => T.tfParseFieldRef(`[ds1].[${s}]`);
const cat = (fieldName, applied, extra = {}) => ({ fieldName, filterType: "categorical", isExcludeMode: false, isAllSelected: false, appliedValues: applied, ...extra });
const range = (fieldName, lo, hi, extra = {}) => ({ fieldName, filterType: "range", minValue: lo, maxValue: hi, includeNullValues: false, ...extra });
const day = (y, m, d) => cell("date", new Date(Date.UTC(y, m, d)));

/** label, field, filter, card mode, expected kind ("text" = stays a card), rows kept of 48 */
const CASES = [
  ["categorical (All)", "none:Region:nk", cat("Region", REGIONS.map(v => dv(v)), { isAllSelected: true }), "checklist", "slicer", 48],
  ["2 of 4 values (checklist)", "none:Region:nk", cat("Region", [dv("East"), dv("West")]), "checklist", "slicer", 23],
  ["one value (dropdown)", "none:Region:nk", cat("Region", [dv("East")]), "dropdown", "list", 12],
  ["exclude East", "none:Region:nk", cat("Region", [dv("East")], { isExcludeMode: true }), "checklist", "slicer", 36],
  ["East and Null", "none:Region:nk", cat("Region", [dv("East"), nul]), "checklist", "slicer", 13],
  ["a value not in the data", "none:Region:nk", cat("Region", [dv("East"), dv("North")]), "checklist", "slicer", 12],
  ["boolean True", "none:Is Big:nk", cat("Is Big", [dv(true, "True")]), "radiolist", "list", 16],
  ["integer dimension", "none:Fiscal Year:ok", cat("Fiscal Year", [dv(2021), dv(2022)]), "checklist", "slicer", 20],
  ["YEAR(Order Date)", "yr:Order Date:ok", cat("YEAR(Order Date)", [dv(2024)]), "dropdown", "list", 24],
  ["QUARTER(Order Date)", "qr:Order Date:ok", cat("QUARTER(Order Date)", [dv("Q1")]), "checklist", "slicer", 12],
  ["MONTH(Order Date)", "mn:Order Date:ok", cat("MONTH(Order Date)", [dv("January"), dv("February")]), "checklist", "slicer", 8],
  ["month and year", "tmn:Order Date:ok", cat("MONTH(Order Date)", [{ value: "2024-03-01", nativeValue: new Date(Date.UTC(2024, 2, 1)), formattedValue: "March 2024" }]), "checklist", "slicer", 4],
  ["range of dates", "none:Order Date:qk", range("Order Date", day(2024, 2, 1), day(2024, 8, 30)), null, "range", 16],
  ["range of dates, open start", "none:Order Date:qk", range("Order Date", nan, day(2024, 8, 30)), null, "range", 20],
  ["range of date-times", "none:Ship Time:qk", range("Ship Time", cell("date-time", new Date(Date.UTC(2024, 2, 1, 13))), cell("date-time", new Date(Date.UTC(2024, 8, 30, 13)))), null, "range", 16],
  ["range of numbers", "none:Sales:qk", range("Sales", dv(100), dv(300)), null, "range", 21],
  ["range of numbers, open end", "none:Sales:qk", range("Sales", dv(100), nan), null, "range", 38],
  ["range of SUM(Sales)", "sum:Sales:qk", range("SUM(Sales)", dv(100), dv(5000)), null, "marks", null],
  ["range of YEAR(Order Date)", "yr:Order Date:qk", range("YEAR(Order Date)", dv(2024), dv(2024)), null, "range", 24],
  ["range of months", "tmn:Order Date:qk", range("MONTH(Order Date)", day(2024, 0, 1), day(2024, 5, 1)), null, "range", 12],
  ["WEEKDAY(Order Date)", "wd:Order Date:ok", cat("WEEKDAY(Order Date)", [dv("Monday")]), "dropdown", "list", null],
  ["relative: last 3 months", "none:Order Date:qk", { fieldName: "Order Date", filterType: "relative-date", periodType: "months", rangeType: "lastn", rangeN: 3, anchorDate: dv("2024-06-15") }, null, "range", 4],
  ["relative: year to date", "none:Order Date:qk", { fieldName: "Order Date", filterType: "relative-date", periodType: "years", rangeType: "todate", anchorDate: dv("2024-06-15") }, null, "range", 12],
  ["hierarchical (cube)", "none:Region:nk", { fieldName: "Region", filterType: "hierarchical" }, null, "text", null]
];

for (const [label, inner, f, mode, kind, kept] of CASES) {
  test(`filter card: ${label}`, async () => {
    const slicers = T.createSlicerModel();
    const res = await slicers.adoptCard({ filter: f, field: ref(inner), filterName: f.fieldName, visualName: f.fieldName, mode, model, layout: { heightPx: 300 } }, ws, model);
    if (kind === "text") { assert.equal(typeof res, "string", "stays a card"); return; }
    assert.notEqual(typeof res, "string", `adopted (${res})`);
    assert.equal(res.kind, kind);
    if (kept !== null) assert.equal(rows.map((_, i) => i).filter(i => res.pass(i)).length, kept, "rows kept");
  });
}
