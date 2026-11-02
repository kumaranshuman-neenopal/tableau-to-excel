"use strict";
/* A mock of Tableau's Extensions API for the "Visual Gallery" test workbook (every chart type on one dashboard), with
 * two filter cards added to its dashboard: Region (Central, East, West) and YEAR(Order Date) (2024, 2025) – both
 * applied to every sheet, as a filter "on all using this data source" is. Summary data is aggregated from synthetic
 * Superstore rows under those filters; the data source returns every row. */
const fs = require("fs");
const path = require("path");

let seed = 7;
const rnd = () => (seed = (seed * 16807) % 2147483647) / 2147483647;
const pick = a => a[Math.floor(rnd() * a.length)];
const SUB = { Furniture: ["Bookcases", "Chairs", "Furnishings", "Tables"],
  "Office Supplies": ["Appliances", "Art", "Binders", "Envelopes", "Fasteners", "Labels", "Paper", "Storage", "Supplies"],
  Technology: ["Accessories", "Copiers", "Machines", "Phones"] };
const STATES = [["California", 36.1, -119.6], ["New York", 42.1, -74.9], ["Texas", 31.0, -97.5], ["Washington", 47.4, -121.4],
  ["Florida", 27.7, -81.6], ["Illinois", 40.3, -89.0], ["Ohio", 40.3, -82.7], ["Georgia", 33.0, -83.6]];
const recs = [];
for (let i = 0; i < 2500; i++) {
  const cat = pick(Object.keys(SUB));
  const st = pick(STATES);
  const d = new Date(Date.UTC(2023 + Math.floor(rnd() * 4), Math.floor(rnd() * 12), 1 + Math.floor(rnd() * 28)));
  const sales = Math.round((20 + rnd() * 900) * (cat === "Technology" ? 1.6 : 1) * 100) / 100;
  recs.push({ Region: pick(["Central", "East", "South", "West"]), Category: cat, "Sub-Category": pick(SUB[cat]),
    Segment: pick(["Consumer", "Corporate", "Home Office"]), "Ship Mode": pick(["First Class", "Same Day", "Second Class", "Standard Class"]),
    date: d, Sales: sales, Profit: Math.round(sales * (rnd() * 0.6 - 0.15) * 100) / 100, Quantity: 1 + Math.floor(rnd() * 14),
    "Order ID": "CA-" + i, "Customer Name": "Customer " + Math.floor(rnd() * 120), State: st[0], lat: st[1], lon: st[2],
    days: Math.floor(rnd() * 7) });
}
/* GALLERY_SINGLE=1: the Region card is a Single Value dropdown with West picked.
 * GALLERY_RANGE=1: the year card is a range of Order Date (2024-03-01 – 2025-06-30) instead. */
/* GALLERY_PARAMS=1: the year card is replaced by Start Date / End Date parameter controls (2024-03-01, 2025-06-30)
 * and a Date Flag calculation ([Order Date] between them) that filters every sheet – the Retail dashboard's pattern. */
const SINGLE = !!process.env.GALLERY_SINGLE, PARAMS = !!process.env.GALLERY_PARAMS, RANGE = !!process.env.GALLERY_RANGE || PARAMS;
const state = { regions: SINGLE ? ["West"] : ["Central", "East", "West"], years: [2024, 2025],
                from: new Date(Date.UTC(2024, 2, 1)), to: new Date(Date.UTC(2025, 5, 30)) };
const passes = r => state.regions.includes(r.Region) &&
  (RANGE ? r.date >= state.from && r.date <= state.to : state.years.includes(r.date.getUTCFullYear()));

const fmtNum = n => n.toLocaleString("en-US", { maximumFractionDigits: 0 });
const dims = {
  Region: { t: "string", f: r => r.Region }, Category: { t: "string", f: r => r.Category },
  "Sub-Category": { t: "string", f: r => r["Sub-Category"] }, Segment: { t: "string", f: r => r.Segment },
  "Ship Mode": { t: "string", f: r => r["Ship Mode"] }, "Customer Name": { t: "string", f: r => r["Customer Name"] },
  "State/Province": { t: "string", f: r => r.State }, "Country/Region": { t: "string", f: () => "United States" },
  "MONTH(Order Date)": { t: "date", f: r => new Date(Date.UTC(r.date.getUTCFullYear(), r.date.getUTCMonth(), 1)),
                         fmt: v => v.toLocaleString("en-US", { month: "long", year: "numeric", timeZone: "UTC" }) },
  "YEAR(Order Date)": { t: "int", f: r => r.date.getUTCFullYear(), fmt: String },
  "QUARTER(Order Date)": { t: "string", f: r => "Q" + (Math.floor(r.date.getUTCMonth() / 3) + 1) },
  Quantity: { t: "int", f: r => r.Quantity, fmt: String },
  "Latitude (generated)": { t: "float", f: r => r.lat, fmt: String }, "Longitude (generated)": { t: "float", f: r => r.lon, fmt: String }
};
const meas = {
  "SUM(Sales)": { agg: g => g.reduce((s, r) => s + r.Sales, 0), fmt: v => "$" + fmtNum(v) },
  "SUM(Profit)": { agg: g => g.reduce((s, r) => s + r.Profit, 0), fmt: v => "$" + fmtNum(v) },
  "CNT(Order ID)": { agg: g => g.length, fmt: fmtNum, t: "int" },
  "AVG(Days to Ship Actual)": { agg: g => g.reduce((s, r) => s + r.days, 0) / g.length, fmt: v => v.toFixed(1) },
  "AGG(Profit Ratio)": { agg: g => g.reduce((s, r) => s + r.Profit, 0) / g.reduce((s, r) => s + r.Sales, 0), fmt: v => (v * 100).toFixed(1) + "%" }
};
const dv = (v, f) => ({ value: v, nativeValue: v, formattedValue: f ? f(v) : String(v) });
function summary(dimNames, measNames, measureValues) {
  const groups = new Map();
  recs.filter(passes).forEach(r => {
    const key = dimNames.map(d => String(+dims[d].f(r) || dims[d].f(r))).join("|");
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(r);
  });
  const columns = dimNames.map((d, i) => ({ fieldName: d, dataType: dims[d].t, index: i }));
  const data = [];
  for (const g of groups.values()) {
    const keyRow = dimNames.map(d => dv(dims[d].f(g[0]), dims[d].fmt));
    if (measureValues) measNames.forEach(m => data.push([...keyRow, dv(m), dv(meas[m].agg(g), meas[m].fmt)]));
    else data.push([...keyRow, ...measNames.map(m => dv(meas[m].agg(g), meas[m].fmt))]);
  }
  if (measureValues) columns.push({ fieldName: "Measure Names", dataType: "string" }, { fieldName: "Measure Values", dataType: "float" });
  else measNames.forEach(m => columns.push({ fieldName: m, dataType: meas[m].t || "float" }));
  return { columns, data };
}
const SHEETS = {
  "KPI - Total Sales": () => summary([], ["SUM(Sales)"]),
  "Text Table": () => summary(["Region", "Category"], ["SUM(Sales)", "SUM(Profit)"], true),
  "Highlight Table": () => summary(["Sub-Category", "Region"], ["SUM(Profit)"]),
  "Heat Map": () => summary(["Sub-Category", "Segment"], ["SUM(Profit)", "SUM(Sales)"]),
  "Vertical Bar": () => summary(["Category"], ["SUM(Sales)"]),
  "Horizontal Bar (sorted)": () => summary(["Sub-Category"], ["SUM(Sales)"]),
  "Stacked Bar": () => summary(["Region", "Segment"], ["SUM(Sales)"]),
  "Side-by-Side Bar": () => summary(["Region", "Segment"], ["SUM(Sales)"]),
  "Line (Continuous)": () => summary(["MONTH(Order Date)"], ["SUM(Sales)"]),
  "Line (Discrete, by Category)": () => summary(["YEAR(Order Date)", "QUARTER(Order Date)", "Category"], ["SUM(Profit)"]),
  "Area Chart": () => summary(["MONTH(Order Date)", "Category"], ["SUM(Sales)"]),
  // GALLERY_PIEX=1: like the Retail donut – a label calculation (all "") and an aggregate Tableau returns untyped, all null
  "Pie Chart": () => {
    const s = summary(["Segment"], ["SUM(Sales)"]);
    if (!process.env.GALLERY_PIEX) return s;
    s.columns.splice(1, 0, { fieldName: "Segment Label Filter", dataType: "string" });
    s.columns.push({ fieldName: "SUM(Quantity)", dataType: "unknown" });
    s.data = s.data.map(r => [r[0], dv(""), ...r.slice(1), { value: null, nativeValue: null, formattedValue: "Null" }]);
    return s;
  },
  "Scatter Plot": () => summary(["Region", "Customer Name"], ["SUM(Sales)", "SUM(Profit)"]),
  "Circle View": () => summary(["Category", "Segment", "Sub-Category"], ["SUM(Sales)"]),
  "Dual Combination": () => summary(["QUARTER(Order Date)"], ["SUM(Sales)", "SUM(Profit)"]),
  "Measure Values Bar": () => summary(["Region"], ["SUM(Sales)", "SUM(Profit)"], true),
  "Histogram": () => summary(["Quantity"], ["CNT(Order ID)"]),
  "Gantt Chart": () => summary(["Ship Mode", "MONTH(Order Date)"], ["AVG(Days to Ship Actual)"]),
  "Treemap": () => summary(["Sub-Category"], ["SUM(Profit)", "SUM(Sales)"]),
  "Packed Bubbles": () => summary(["Category", "Sub-Category"], ["SUM(Sales)"]),
  "Symbol Map": () => summary(["State/Province", "Country/Region", "Latitude (generated)", "Longitude (generated)"], ["SUM(Profit)", "SUM(Sales)"]),
  "Filled Map": () => summary(["State/Province", "Country/Region", "Latitude (generated)", "Longitude (generated)"], ["AGG(Profit Ratio)"])
};

const filtersOf = sheet => [
  { fieldName: "Region", filterType: "categorical", worksheetName: sheet, isExcludeMode: false, isAllSelected: false,
    appliedValues: state.regions.map(v => dv(v)), getDomainAsync: async () => ({ values: ["Central", "East", "South", "West"].map(v => dv(v)) }) },
  PARAMS
    ? { fieldName: "Date Flag", filterType: "categorical", worksheetName: sheet, isExcludeMode: false, isAllSelected: false,
        appliedValues: [dv(true)], getDomainAsync: async () => ({ values: [dv(true), dv(false)] }) }
  : RANGE
    ? { fieldName: "Order Date", filterType: "range", worksheetName: sheet, includeNullValues: false,
        minValue: { value: iso(state.from), nativeValue: state.from, formattedValue: iso(state.from) },
        maxValue: { value: iso(state.to), nativeValue: state.to, formattedValue: iso(state.to) } }
    : { fieldName: "YEAR(Order Date)", filterType: "categorical", worksheetName: sheet, isExcludeMode: false, isAllSelected: false,
        appliedValues: state.years.map(v => dv(v)), getDomainAsync: async () => ({ values: [2023, 2024, 2025, 2026].map(v => dv(v)) }) }
];

const iso = d => d.toISOString().slice(0, 10);
const RAW = [["Row ID", "int", (r, i) => i + 1], ["Order ID", "string", r => r["Order ID"]], ["Order Date", "date", r => r.date],
  ["Ship Mode", "string", r => r["Ship Mode"]], ["Customer Name", "string", r => r["Customer Name"]], ["Segment", "string", r => r.Segment],
  ["Country/Region", "string", () => "United States"], ["State/Province", "string", r => r.State], ["Region", "string", r => r.Region],
  ["Category", "string", r => r.Category], ["Sub-Category", "string", r => r["Sub-Category"]], ["Sales", "float", r => r.Sales],
  ["Quantity", "int", r => r.Quantity], ["Profit", "float", r => r.Profit]];
const dataSource = {
  id: "federated.10nnk8d1vgmw8q17yu76u06pnbcj", name: "Sample - Superstore",
  getLogicalTablesAsync: async () => [{ id: "Orders_ECFCA1FB690A41FE803BC071773BA862", caption: "Orders" }],
  getLogicalTableDataAsync: async () => ({
    columns: RAW.map(([n, t], i) => ({ fieldName: n, dataType: t, index: i })),
    data: recs.map((r, i) => RAW.map(([, t, f]) => { const v = f(r, i); return t === "date" ? { value: iso(v), nativeValue: v, formattedValue: iso(v) } : dv(v); })),
    totalRowCount: recs.length
  })
};
const worksheet = name => ({
  name, sheetType: "worksheet",
  getSummaryDataAsync: async () => SHEETS[name](),
  getFiltersAsync: async () => filtersOf(name),
  getSelectedMarksAsync: async () => ({ data: [] }),
  getDataSourcesAsync: async () => [dataSource]
});

/** the workbook XML with the two filter cards in the empty space beside the last sheet */
function workbookXml() {
  const xml = fs.readFileSync(path.join(__dirname, "../fixtures/gallery.twb"), "utf8");
  const ds = "[federated.10nnk8d1vgmw8q17yu76u06pnbcj]";
  const cards = `<zone h='6000' id='40' mode='${SINGLE ? "dropdown" : "checkdropdown"}' name='Vertical Bar' param='${ds}.[none:Region:nk]' type-v2='filter' w='15000' x='33611' y='87421' />` +
                (PARAMS ? `<zone h='3000' id='41' param='[Parameters].[Parameter 91]' type-v2='paramctrl' w='7000' x='50000' y='87421' />` +
                          `<zone h='3000' id='42' param='[Parameters].[Parameter 92]' type-v2='paramctrl' w='7000' x='57500' y='87421' />`
                 : RANGE ? `<zone h='6000' id='41' name='Vertical Bar' param='${ds}.[none:Order Date:qk]' type-v2='filter' w='15000' x='50000' y='87421' />`
                 : `<zone h='6000' id='41' mode='checkdropdown' name='Vertical Bar' param='${ds}.[yr:Order Date:ok]' type-v2='filter' w='15000' x='50000' y='87421' />`);
  const at = xml.indexOf("<zone h='11635' id='24'");
  if (at < 0) throw new Error("Filled Map zone not found");
  let out = xml.slice(0, at) + cards + xml.slice(at);
  if (PARAMS) {
    const params = `<column caption='Start Date' datatype='date' name='[Parameter 91]' param-domain-type='any' role='measure' type='quantitative' value='#2024-03-01#'><calculation class='tableau' formula='#2024-03-01#' /></column>` +
                   `<column caption='End Date' datatype='date' name='[Parameter 92]' param-domain-type='any' role='measure' type='quantitative' value='#2025-06-30#'><calculation class='tableau' formula='#2025-06-30#' /></column>`;
    out = out.replace(/(<datasource hasconnection='false' inline='true' name='Parameters' version='18.1'>\s*<aliases enabled='yes' \/>)/, `$1${params}`);
    const calc = `<column caption='Date Flag' datatype='boolean' name='[Calculation_DF]' role='dimension' type='nominal'><calculation class='tableau' ` +
                 `formula='If [Order Date] &gt;= [Parameters].[Parameter 91] AND [Order Date] &lt;= [Parameters].[Parameter 92] THEN TRUE ELSE FALSE END' /></column>`;
    out = out.replace(/(<datasource caption='Sample - Superstore' inline='true' name='federated.10nnk8d1vgmw8q17yu76u06pnbcj' version='18.1'>)/, `$1${calc}`);
    out = out.replace(/<view>/g, `<view><filter class='categorical' column='${ds}.[none:Calculation_DF:nk]'><groupfilter function='member' level='[none:Calculation_DF:nk]' member='true' /></filter>`);
  }
  return out;
}

function dashboardFrom(fmtModel) {
  const dash = fmtModel.dashboards["Visual Gallery"];
  const W = 1800, H = 3180;
  const kind = { worksheet: "worksheet", filter: "quick-filter", color: "legend", text: "text", bitmap: "image", paramctrl: "parameter-control" };
  const paramName = { "[Parameters].[Parameter 91]": "Start Date", "[Parameters].[Parameter 92]": "End Date" };
  const objects = dash.zones.filter(z => kind[z.type] && z.w && z.h).map(z => ({
    id: Number(z.id), name: z.type === "worksheet" ? z.name : z.type === "paramctrl" ? paramName[z.param] : `${z.type} ${z.id}`, type: kind[z.type], isVisible: true, isFloating: false,
    position: { x: z.x / 100000 * W, y: z.y / 100000 * H }, size: { width: z.w / 100000 * W, height: z.h / 100000 * H }
  }));
  const dateParam = (id, name, d) => ({ id, name, dataType: "date", currentValue: { value: iso(d), nativeValue: d, formattedValue: `${d.getUTCMonth() + 1}/${d.getUTCDate()}/${d.getUTCFullYear()}` } });
  return { name: "Visual Gallery", objects, worksheets: Object.keys(SHEETS).map(worksheet),
           getParametersAsync: async () => PARAMS ? [dateParam("[Parameters].[Parameter 91]", "Start Date", state.from), dateParam("[Parameters].[Parameter 92]", "End Date", state.to)] : [] };
}

module.exports = { workbookXml, dashboardFrom, recs, state, passes };
