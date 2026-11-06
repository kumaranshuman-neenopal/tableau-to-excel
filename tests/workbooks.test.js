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
const { tempXlsx, runExporter, inspect, xmlParts } = require("./lib/inspect.js");

const DIR = process.env.TFX_WORKBOOKS || "";
const ECOM = path.join(DIR, "Ecommerce_Dashboard.twbx");
const LOGI = path.join(DIR, "Untitled_17.twbx");
const has = f => !!DIR && fs.existsSync(f);
const EXCEL = process.env.TFX_EXCEL === "1" && process.platform === "win32";
const skipEcom = has(ECOM) ? false : "TFX_WORKBOOKS has no Ecommerce_Dashboard.twbx";
const skipLogi = has(LOGI) ? false : "TFX_WORKBOOKS has no Untitled_17.twbx";
const GAUGE = path.join(DIR, "Gauge Charts.twbx");
const skipGauge = has(GAUGE) ? false : "TFX_WORKBOOKS has no Gauge Charts.twbx";
const TEMPLATE = path.join(DIR, "Gauge Chart Template.twbx");
const skipTemplate = has(TEMPLATE) ? false : "TFX_WORKBOOKS has no Gauge Chart Template.twbx";
const NOI = path.join(DIR, "NOI Fake.twbx");
const skipNoi = has(NOI) ? false : "TFX_WORKBOOKS has no NOI Fake.twbx";

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
  // the scatter's Category: its filter card's dropdown does it; the pie gets its own above it, beside the Channel card
  assert.match(log, /\(Scatter\)" not drawn: the "Category" filter card filters by it already/);
  assert.match(log, /\[Actions\] "Channel \(Revenue Share by Channel \(Pie\)\)": 4 values of channel/);
  assert.equal(Object.keys(file.zip.files).filter(n => /slicerCache\d+\.xml$/.test(n)).length, 7, "7 action slicers");
  assert.ok(file.sheets.includes("Slicer Actions"));
  // flat chips, moving and sizing with their rows – a row group the [–] hides, holding no cell of a chart or table
  const parts = await xmlParts(file.zip);
  const slicers = Object.entries(parts).filter(([n]) => /slicers\//.test(n)).map(([, x]) => x).join("");
  assert.equal((slicers.match(/style="TfxChips"/g) || []).length, 7);
  assert.match(parts["xl/styles.xml"], /<x14:slicerStyle name="TfxChips">/);
  assert.equal((parts["xl/drawings/drawing1.xml"].match(/editAs="twoCell"/g) || []).length, 7);
  const grouped = (parts["xl/worksheets/sheet1.xml"].match(/<row [^>]*outlineLevel="1"[^>]*>/g) || []);
  assert.ok(grouped.length >= 7 && grouped.every(t => t.endsWith("/>")), "band rows grouped, empty");
});

test("Gauge Charts: gauges whose slices are Measure Values calculations recalculate with the slicers", { skip: skipGauge }, async () => {
  const { log, file } = await exportOf(GAUGE, "Gauge Chart Gallery", { REALCSV: "1" }, "gauges.xlsx");
  for (const s of ["Semi-Circle Gauge", "RAG Threshold Gauge", "Speedometer Needle Gauge", "270 Degree Arc Gauge", "Donut Progress Ring"])
    assert.match(log, new RegExp(`\\[Slicers\\] ${s}: follows the slicers \\(1 marks`), s);
  assert.match(log, /\[Slicers\] Gauge Small Multiples by Region: follows the slicers \(4 marks/);
  // a slice: its member's calculation over the rows the slicers keep (fill_r = MAX(MIN(1, MAX(0, on-time %)), 0.5) - 0.5)
  const data = (await xmlParts(file.zip))[`xl/worksheets/sheet${file.sheets.indexOf("Chart Data") + 1}.xml`];
  assert.ok(/<f>IF\(COUNTIFS\([^<]*MAX\(MIN\(1,MAX\(0,/.test(data.replace(/&apos;/g, "'")), "slice formulas");
});

test("Gauge Chart Template: a map gauge's dial and needle are formulas of its parameter cells", { skip: skipTemplate }, async () => {
  // its data is a .hyper extract: the summary rows Tableau gives (placeholder rows level × value, the needle's points)
  const dv = v => ({ value: v, nativeValue: v, formattedValue: String(v) }), NUL = { value: null, nativeValue: null, formattedValue: "Null" };
  const columns = ["level", "value", "point", "SUM(Sales)", "gauge", "Gauge Point"].map((n, i) => ({ fieldName: n, dataType: i < 3 ? "integer" : i === 3 ? "float" : "string", index: i }));
  const data = [];
  for (let l = 5; l >= 1; l--) for (let v = 1; v <= 401; v++) data.push([dv(l), dv(v), NUL, NUL, dv("pt"), NUL]);
  for (let p = 1; p <= 2; p++) data.push([NUL, NUL, dv(p), dv(2297200.86), NUL, dv("pt")]);
  const given = tempXlsx("template-summary.json");
  fs.writeFileSync(given, JSON.stringify({ "Base Gauge": { columns, data } }));
  const { file } = await exportOf(TEMPLATE, "Dashboard 1", { SUMMARY: given }, "template.xlsx");
  const parts = await xmlParts(file.zip);
  const data2 = parts[`xl/worksheets/sheet${file.sheets.indexOf("Chart Data") + 1}.xml`].replace(/&apos;/g, "'");
  // Level 1 (B7) / Over Target (B14) set the first band's end; the needle reads Target (B17)
  assert.match(data2, /<f>MAX\(D3,\(\(DEGREES\(ATAN2\([^<]*'Dashboard 1'!\$B\$7\/'Dashboard 1'!\$B\$14/);
  assert.match(data2, /<f>MAX\(0,MAX\(0,MIN\(\$D\$8,[^<]*'Dashboard 1'!\$B\$17/);
  // the parameter cells are numbers in Tableau's format, not text
  const dash = parts["xl/worksheets/sheet1.xml"];
  assert.match(dash, /<c r="B7" [^>]*><v>0\.4<\/v>/);
});

test("NOI Executive Console: layered gauge needles, backdrop fills, lollipops, overlaid bars, a button as a dropdown", { skip: skipNoi }, async () => {
  // its data is a .hyper extract: the summary rows of the sheets checked here, made up
  const dv = (v, f) => ({ value: v, nativeValue: v, formattedValue: f === undefined ? String(v) : f });
  const colsOf = (names, types) => names.map((n, i) => ({ fieldName: n, dataType: types[i], index: i }));
  const rad = d => d * Math.PI / 180;
  const dial = [];
  [["Rojo", -120, 60], ["Amarillo", 60, 80], ["Verde", 80, 120]].forEach(([name, a0, a1], id) => {
    const pts = [];
    for (let a = a0; a <= a1 + 1e-9; a += 2) pts.push([Math.sin(rad(a)), Math.cos(rad(a))]);
    for (let a = a1; a >= a0 - 1e-9; a -= 2) pts.push([0.75 * Math.sin(rad(a)), 0.75 * Math.cos(rad(a))]);
    pts.forEach((p, k) => dial.push([dv(name), dv(id + 1), dv(k + 1), dv(p[0]), dv(p[1])]));
  });
  // a needle: base left, tip (0.9 R), base right – at deg clockwise from 12
  const needle = deg => { const t = rad(deg), d = [Math.sin(t), Math.cos(t)], q = [Math.cos(t), -Math.sin(t)];
    return [[q[0] * 0.054, q[1] * 0.054], [d[0] * 0.9, d[1] * 0.9], [-q[0] * 0.054, -q[1] * 0.054]].map((p, k) => [dv(1), dv(k + 1), dv(p[0]), dv(p[1])]); };
  const needleCols = colsOf(["ID", "Path Aguja", "AGG(X Aguja)", "AGG(Y Aguja)"], ["integer", "integer", "float", "float"]);
  const months = ["Jan 24", "Feb 24", "Mar 24"];
  const ive = months.flatMap((m, k) => [[dv(m), dv(m), dv("General - Real Act Seleccionado"), dv(0.4 + k / 10), dv(0.4 + k / 10)],
                                        [dv(m), dv(m), dv("MIN(0)"), dv(0), dv(0.4 + k / 10)]]);
  // as Tableau gives them: four measures on the path One → Line → Zero → Line (copy), the copy = Line
  const proj = ["City Walk", "Explanada Pachuca"].flatMap((p, k) => [[dv(p), dv("Min. Gauge - One"), dv(1), dv(0)],
    [dv(p), dv("Gauge - Dynamic Project Chooser Line"), dv([1, -468.4][k]), dv(0)], [dv(p), dv("Min. Gauge - Zero"), dv(0), dv(0)],
    [dv(p), dv("Gauge - Dynamic Project Chooser Line (copy)"), dv([1, -468.4][k]), dv(0)]]);
  const given = tempXlsx("noi-summary.json");
  fs.writeFileSync(given, JSON.stringify({
    "Gauge Original NOI (2)": { columns: colsOf(["Segment Color", "Segment ID", "Path Order", "SUM(X)", "SUM(Y)"], ["string", "integer", "integer", "float", "float"]), data: dial },
    "Needle NOI Total NOI": { columns: needleCols, data: needle(30) },
    "Needle NOI Total NOI|Mensual / Acumulado=Acumulado": { columns: needleCols, data: needle(-60) },
    "Needle NOI Total NOI|Mes=Enero|Mensual / Acumulado=Acumulado": { columns: needleCols, data: needle(-90) },
    "Needle NOI Bolita NOI": { columns: colsOf(["MIN(Bolita)"], ["float"]), data: [[dv(0)]] },
    "Gauge Valor IvE": { columns: colsOf(["MY(Date)", "General - MY Axis", "Measure Names", "Measure Values", "AGG(General - Real Act Seleccionado (G))"], ["string", "string", "string", "float", "float"]), data: ive },
    "Gauge Proyectos Centros Comerciales A": { columns: colsOf([" Proyecto", "Measure Names", "Measure Values", "AGG(MIN(0))"], ["string", "string", "float", "float"]), data: proj },
    // General: the project groups; Detalle: the projects – other rows, more of them
    "Gauge Proyectos Mixto A": { columns: colsOf(["Parameter GP / P Mixto", "Measure Names", "Measure Values"], ["string", "string", "float"]),
      data: ["Capital Reforma", "Paseo Arcos Bosques"].flatMap((p, k) => [[dv(p), dv("Min. Gauge - One"), dv(1)], [dv(p), dv("Gauge - Dynamic Project Chooser Line"), dv([1, -0.5][k])]]) },
    "Gauge Proyectos Mixto A|Grupo Proyecto / Proyecto=Proyectos": { columns: colsOf(["Parameter GP / P Mixto", "Measure Names", "Measure Values"], ["string", "string", "float"]),
      data: ["Capital Reforma", "Paseo Arcos Bosques C.C.", "Paseo Arcos Bosques Marco II"].flatMap((p, k) => [[dv(p), dv("Min. Gauge - One"), dv(1)], [dv(p), dv("Gauge - Dynamic Project Chooser Line"), dv([0.3, 1, -2][k])]]) },
    // the % next to the bars: a measure on Text, nothing on Columns, the columns' field labels hidden
    "Gauge Proyectos Centros Comerciales B": { columns: colsOf([" Proyecto", "Gauge - Dynamic Project Chooser"], ["string", "float"]),
                                               data: [[dv("City Walk"), dv(2.43, "243.3%")], [dv("Explanada Pachuca"), dv(-4.68, "-468.4%")]] },
    // the General / Detalle buttons: Text Names labels the marks, Modo Val is what the action sets, EsActivo colours them
    "Button B": { columns: colsOf(["EsActivo", "Modo", "Text Names", "Modo Val", "SUM(X)", "AGG(MIN(1))"], ["bool", "string", "string", "integer", "integer", "integer"]),
                  data: [[dv(true, "True"), dv("Grupo de Proyectos"), dv("General"), dv(1), dv(1), dv(1)],
                         [dv(false, "False"), dv("Proyectos"), dv("Detalle"), dv(2), dv(2), dv(1)]] }
  }));
  // MIXED: the project gauge that uses Grupo Proyecto / Proyecto stays off the slicers (as in Tableau, whose data is a
  // .hyper), so the General / Detalle button sheet becomes its dropdown
  const { log, file } = await exportOf(NOI, "Executive Console", { SUMMARY: given, MIXED: "^Gauge Proyectos Mixto A$" }, "noi.xlsx");
  // Mes_TMP only feeds Mes (a bridge action): Tableau is set through Mes alone; Mes × Mensual / Acumulado read as pairs
  assert.match(log, /\[States\] parameter "Mes_TMP": set only through "Mes"/);
  assert.match(log, /\[States\] Executive Console: \d+ filter values \([1-9]\d* pair\(s\) of dropdowns\)/);
  // the Mensual / Acumulado button → a dropdown; only the visuals whose calculations use the parameter are read
  assert.match(log, /\[States\] parameter "Mensual \/ Acumulado": 2 values, \d+ visual\(s\) use it \(the "Buton A" button → a dropdown\)/);
  assert.doesNotMatch(log, /"PUENTE Param" button/, "a one-mark bridge sheet is not a button");
  const parts = await xmlParts(file.zip);
  const report = file.text.join("\n");
  const sheetXml = n => parts[`xl/worksheets/sheet${file.sheets.indexOf(n) + 1}.xml`].replace(/&apos;/g, "'").replace(/&quot;/g, '"');
  // the needle and the dot sheets are part of the gauge; the needle at 30° (150° from the dial's start at -120°) and
  // at -60° for Acumulado – Tableau's angles per value on the States sheet
  assert.match(report, /Needle NOI Total NOI.*Worksheet layered on a gauge.*The gauge's needle/);
  assert.match(report, /Needle NOI Bolita NOI.*Worksheet layered on a gauge.*The gauge's centre dot/);
  const states = sheetXml("States Executive Console");
  assert.match(states, /<row r="(\d+)"[^>]*>(?:(?!<\/row>).)*<v>-59\.99\d*<\/v>(?:(?!<\/row>).)*<v>29\.99\d*<\/v>/, "the needle's angle per value: one row");
  // Enero with Acumulado: a state of its own (key "0|Enero||1|Acumulado"), the needle at -90° there; A2 joins the changes
  assert.match(parts["xl/sharedStrings.xml"] || states, /0\|Enero\|\|1\|Acumulado/);
  assert.match(states, /<row r="(\d+)"[^>]*>(?:(?!<\/row>).)*<v>-59\.99\d*<\/v>(?:(?!<\/row>).)*<v>-90<\/v>/, "the pair's needle angle on the needle's row");
  assert.match(states, /<c r="A2"[^>]*><f>[^<]*MID\(/);
  // a key cell of one dropdown: its test made a number ((J4<>"x")=0 is never TRUE in Excel – TRUE / FALSE are not 0)
  assert.match(states, /<c r="[B-Z]2"[^>]*><f>IF\([^<]*IF\(0\+\(/);
  for (const n of file.sheets) assert.doesNotMatch(sheetXml(n), /<f>IF\(=|MATCH\(,/, `${n}: a formula with an empty key reference`);
  assert.match(sheetXml("Chart Data"), /<f>MAX\(0,MAX\(0,MIN\([^<]*'States Executive Console'!\$[B-Z]\$2[^<]*<\/f><v>149\.2<\/v>/, "the needle reads its own key cell (its dropdowns only)");
  // the orange panel image: its colour behind the cells
  assert.match(report, /Image\/Backgrounder\.png.*Background image.*Cell fill #FBB05C/);
  assert.match(parts["xl/styles.xml"], /rgb="FFFBB05C"/);
  // lollipops: thin bars + dots on one scale; the project bars: track and value overlaid, not stacked
  assert.match(report, /Gauge Valor IvE.*Column chart \(thin bars\) with a dot on each/);
  assert.match(report, /Gauge Proyectos Centros Comerciales A.*Overlaid bar chart/);
  // Detalle changes which rows the Mixto bars have: 3 slots by position, names and values read from the States sheet
  assert.match(log, /\[States\] Gauge Proyectos Mixto A: 3 rows by position/);
  assert.match(parts["xl/sharedStrings.xml"], /Gauge Proyectos Mixto A \(row 3\)/);
  // the % table has no heading over its measure, as in Tableau
  assert.doesNotMatch(parts["xl/sharedStrings.xml"], /<t[^>]*>Gauge - Dynamic Project Chooser<\/t>/);
  // the light track (Min. Gauge - One: #FFF2E3) behind, the orange value in front; the copy that repeats it is left out
  const bars = Object.entries(parts).find(([n, x]) => /^xl\/charts\/chart\d+\.xml$/.test(n) && /Chooser Line/.test(x))[1];
  const sers = [...bars.matchAll(/<c:ser>.*?<\/c:ser>/gs)].map(m => [(m[0].match(/<c:v>([^<]*)<\/c:v>/) || [])[1], (m[0].match(/srgbClr val="(\w+)"/) || [])[1]]);
  assert.deepEqual(sers.filter(([n]) => !/Zero/.test(n)), [["Min. Gauge - One", "FFF2E3"], ["Gauge - Dynamic Project Chooser Line", "FBB05C"]]);
  // the General / Detalle buttons: a dropdown of their labels (Tableau set to the value each sets), and button cells
  // green (#499202) while the dropdown holds their label, grey (#535355) otherwise
  assert.match(log, /\[States\] parameter "Grupo Proyecto \/ Proyecto": 2 values, \d+ visual\(s\) use it \(the "Button B" button → a dropdown\)/);
  const list = [...states.matchAll(/<c r="([A-Z]+)(\d+)"[^>]*t="s"[^>]*><v>(\d+)<\/v>/g)];
  const strs = [...(parts["xl/sharedStrings.xml"] || "").matchAll(/<si>(.*?)<\/si>/gs)].map(m => m[1].replace(/<[^>]+>/g, ""));
  const col = list.find(m => strs[+m[3]] === "General");
  assert.ok(col && list.some(m => m[1] === col[1] && strs[+m[3]] === "Detalle"), "the dropdown's list: General, Detalle");
  const dash = sheetXml("Executive Console");
  assert.match(dash, /<cfRule type="expression"[^>]*><formula>'Executive Console'!\$[A-Z]+\$\d+="General"<\/formula>/);
  assert.match(dash, /<cfRule type="expression"[^>]*><formula>'Executive Console'!\$[A-Z]+\$\d+="Detalle"<\/formula>/);
  assert.match(parts["xl/styles.xml"], /<dxf><fill><patternFill patternType="solid"><bgColor rgb="FF499202"\/>/);
  assert.match(parts["xl/styles.xml"], /<fgColor rgb="FF535355"\/>/);
  assert.match(report, /Grupo Proyecto \/ Proyecto.*Button sheet → dropdown \+ buttons coloured by the dropdown \(General \/ Detalle/);
});

test("Ecommerce exported at Channel = Website: the pie gets the other channels", { skip: skipEcom }, async () => {
  const { log } = await exportOf(ECOM, "Sales Overview", { REALCSV: "1", SELECT: "Channel=Website" }, "website.xlsx");
  assert.match(log, /\[Every value\] Sales Overview: Channel cleared/);
  assert.match(log, /\[Every value\] Revenue Share by Channel \(Pie\): 3 more slice\(s\)/);
});

test("Ecommerce exported at Device Type = Mobile: the donut gets the other device types", { skip: skipEcom }, async () => {
  const { log } = await exportOf(ECOM, "Gauge & Donut", { REALCSV: "1", SELECT: "Device Type=Mobile" }, "mobile.xlsx");
  assert.match(log, /\[Every value\] Revenue by Device Type \(Donut\): 2 more slice\(s\)/);
});

test("Ecommerce heatmap (WEEKDAY × hour): every cell follows the slicers", { skip: skipEcom }, async () => {
  const { log, file } = await exportOf(ECOM, "Sales Overview", { REALCSV: "1" }, "heatmap.xlsx");
  assert.match(log, /Orders by Weekday and Hour \(Heatmap\): follows the slicers \(168 marks/);
  assert.ok(file.visuals.find(v => /Heatmap/.test(v.worksheet)).follows);
});

test("Ecommerce Detail Analysis: the packed bubbles follow the slicers", { skip: skipEcom }, async () => {
  const { file } = await exportOf(ECOM, "Detail Analysis", { REALCSV: "1" }, "detail.xlsx");
  const bubbles = file.visuals.find(v => /Bubbles/.test(v.worksheet));
  assert.ok(bubbles && bubbles.follows, bubbles && bubbles.note);
});

test("Ecommerce, all four dashboards: a slicer table each, 11 action slicers", { skip: skipEcom }, async () => {
  const { file } = await exportOf(ECOM, "ALL", { REALCSV: "1" }, "all.xlsx");
  for (const s of ["KPI Overview", "Sales Overview", "Gauge & Donut", "Detail Analysis", "Slicer Data", "Slicer Data 2", "Slicer Data 3", "Slicer Data 4"])
    assert.ok(file.sheets.includes(s), `sheet ${s}`);
  // 7 + 3 (Detail Analysis: Customer Region is a card) + 1 (Gauge & Donut: the donut's Device Type, beside its card)
  assert.equal(Object.keys(file.zip.files).filter(n => /slicerCache\d+\.xml$/.test(n)).length, 11);
  // the chip style once, however many dashboards add it
  const styles = await file.zip.file("xl/styles.xml").async("string");
  assert.equal((styles.match(/<x14:slicerStyle name="TfxChips">/g) || []).length, 1);
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
