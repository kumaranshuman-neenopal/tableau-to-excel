"use strict";
/* node tests/lib/export-twbx.js <file.twbx> <dashboard | ALL> <out.xlsx>  –  any dashboard of any workbook through the
 * real exportToExcel(), with a mock Tableau API made from the workbook alone: synthetic rows for every data source (its
 * non-calculated columns), each worksheet's summary data aggregated from them (its source-column dimensions and SUM /
 * AVG / MIN / MAX / COUNT of source columns, its calculations through the extension's own evaluator), the worksheet's
 * filters as Tableau reports them, and applyFilterAsync / clearFilterAsync that the summary data follows. Prints the
 * [Filters] / [Slicers] / [Actions] / [States] / [Every value] lines.
 * Environment switches:
 *   REALCSV=1      a relationships data source served from the CSV files packed in the .twbx (one logical table each)
 *   REALXLSX=1     a single-table data source served from the packed Data/*.xlsx
 *   SELECT=f=v     the filter card of field f exported with the one value v
 *   ALL (as the dashboard)  every dashboard of the workbook in one file, as Tableau Cloud's "All dashboards"
 *   NOSLICERS / NODROPDOWN / NOJOIN / MIXED=<regex> / ZONEPOS=<json>  switch parts off or force a case (see below) */
const fs = require("fs");
const path = require("path");
const { load } = require("./load.js");
const [TWBX, DASH, OUTARG] = process.argv.slice(2);
const OUT = path.resolve(OUTARG || "out/any.xlsx");
const SELECT = process.env.SELECT;           // "field=value": that card keeps one value

let seed = 7;
const rnd = () => (seed = (seed * 16807) % 2147483647) / 2147483647;
const dv = (v, f) => ({ value: v, nativeValue: v, formattedValue: f === undefined ? String(v) : f });
const iso = d => d.toISOString().slice(0, 10);
const AGG = { sum: "SUM", avg: "AVG", min: "MIN", max: "MAX", cnt: "CNT", ctd: "CNTD" };

(async () => {
  const log = console.log.bind(console);
  const keep = /^\[(Filters|Slicers|Export\]|Parameters|States|Actions|Every value)/;
  console.log = (...a) => { if (keep.test(String(a[0]))) log(...a); };
  console.info = console.table = () => {};
  console.warn = (...a) => { if (/filter|slicer/i.test(String(a[0]))) log("WARN", ...a); };
  let saved = null;
  const ALL = DASH === "ALL";
  const patch = ALL ? [["function wantsAllDashboards() {", "function wantsAllDashboards() { return true; } function unusedWants() {"],
                       ["async function collectAllDashboards(current, onProgress) {", "async function collectAllDashboards(current) { return globalThis.__collect(current); } async function unusedCollect(current, onProgress) {"]] : [];
  // MIXED: visuals named like it keep away from the slicer formulas (a calculation they cannot translate, say)
  if (process.env.MIXED) patch.push(["const binding = slicers.bind(item.vm, item.name, filters, fmtModel);", "const binding = new RegExp(process.env.MIXED).test(item.name) ? \"test: kept off the formulas\" : slicers.bind(item.vm, item.name, filters, fmtModel);"]);
  const T = load({ dashboard: {}, captureDownload: b => { saved = b; }, expose: ["applyWorkbookXml", "exportToExcel", "tfFieldInfo", "tfParseFieldRef", "FORMAT_CONFIG", "slicerParseCalc", "slicerEval", "tableauToExcelNumFmt", "tfFormatNumber"], patch });
  if (process.env.NODROPDOWN) T.FORMAT_CONFIG.dropdownCards = false;
  if (process.env.NOSLICERS) T.FORMAT_CONFIG.slicers = false;
  // filters set through the API (applyFilterAsync / clearFilterAsync): field caption → the values kept
  const active = new Map();
  if (process.env.SELECT) { const [fn, v] = process.env.SELECT.split("="); active.set(fn, new Set([v])); }
  globalThis.__reads = 0;
  const zip = await JSZip.loadAsync(fs.readFileSync(TWBX));
  const xml = await zip.file(Object.keys(zip.files).find(n => /\.twb$/i.test(n) && !n.includes("/"))).async("string");
  const { formatModel: model } = await T.applyWorkbookXml(xml, path.basename(TWBX), false, zip);
  const doc = new DOMParser().parseFromString(xml, "text/xml");

  // ── synthetic rows per data source: members the workbook names (filters, aliases) or made-up ones
  const members = {};                       // "ds|column" → Set of values
  Array.from(doc.getElementsByTagName("groupfilter")).forEach(g => {
    const m = g.getAttribute("member"), lvl = g.getAttribute("level");
    if (!m || !lvl) return;
    const ref = T.tfParseFieldRef(lvl.includes("].[") ? lvl : `[x].${lvl}`);
    const anc = (() => { let p = g.parentNode; while (p && p.tagName !== "filter") p = p.parentNode; return p; })();
    const col = anc && T.tfParseFieldRef(anc.getAttribute("column") || "");
    const key = `${(col && col.ds) || ""}|${ref ? ref.name : lvl}`;
    (members[key] || (members[key] = new Set())).add(m.replace(/^"|"$/g, "").replace(/^&quot;|&quot;$/g, ""));
  });
  const sources = {};
  Object.values(model.fields).forEach(f => {
    if (!f || !f.ds || f.ds === "Parameters" || f.param || f.formula || f.bin || f.group || /^:|Measure Names|Number of Records/i.test(f.name)) return;
    const s = sources[f.ds] || (sources[f.ds] = { id: f.ds, name: f.ds, cols: [] });
    if (s.cols.some(c => c.name === f.name)) return;
    const t = String(f.datatype || "string");
    s.cols.push({ name: f.name, caption: f.caption || f.name, type: t === "datetime" ? "date-time" : t === "real" ? "float" : t === "integer" ? "int" : t });
  });
  Object.values(sources).forEach(s => {
    s.rows = [];
    for (let i = 0; i < 300; i++) s.rows.push(s.cols.map(c => {
      if (c.type === "string") {
        const m = members[`${s.id}|${c.name}`];
        const pool = m && m.size ? [...m] : [1, 2, 3, 4, 5].map(k => `${c.caption} ${k}`);
        return pool[Math.floor(rnd() * pool.length)];
      }
      if (/date/.test(c.type)) return new Date(Date.UTC(2024 + Math.floor(rnd() * 2), Math.floor(rnd() * 12), 1 + Math.floor(rnd() * 28)));
      if (c.type === "boolean") return rnd() < 0.5;
      return c.type === "int" ? Math.round(rnd() * 1000) : Math.round(rnd() * 100000) / 100;
    }));
    s.dataSource = { id: s.id, name: s.id, getLogicalTablesAsync: async () => [{ id: "t", caption: s.id }],
      getLogicalTableDataAsync: async () => ({ columns: s.cols.map((c, i) => ({ fieldName: c.caption, dataType: c.type, index: i })),
        data: s.rows.map(r => r.map((v, i) => v instanceof Date ? { value: iso(v), nativeValue: v, formattedValue: iso(v) } : dv(v))), totalRowCount: s.rows.length }) };
  });
  // REALCSV: a data source of related CSV tables packed in the .twbx served as Tableau does – one logical table per CSV
  // (id = the workbook's object id), the summary data aggregated from the tables joined here independently
  if (process.env.REALCSV) for (const [dsId, rel] of Object.entries(model.relationships || {})) {
    const s = sources[dsId]; if (!s) continue;
    const parseCsv = txt => { const lines = txt.split(String.fromCharCode(10)).map(l => l.split(String.fromCharCode(13)).join("")).filter(Boolean); const split = l => { const out = []; let cur = "", q = false; for (const ch of l) { if (ch === '"') q = !q; else if (ch === "," && !q) { out.push(cur); cur = ""; } else cur += ch; } out.push(cur); return out; }; return { head: split(lines[0]), rows: lines.slice(1).map(split) }; };
    const typeOf = name => { const f = Object.values(model.fields).find(x => x.ds === dsId && x.name === name) || {}; const t = f.datatype || "string"; return t === "real" ? "float" : t === "integer" ? "int" : t; };
    const capOf = name => { const f = Object.values(model.fields).find(x => x.ds === dsId && x.name === name) || {}; return f.caption || name; };
    const conv = (t, v) => v === "" ? null : t === "int" || t === "float" ? Number(v) : t === "date" ? new Date(v + "T00:00:00Z") : v;
    const tables = {};
    for (const [id, cap] of Object.entries(rel.objects)) {
      const f = Object.keys(zip.files).find(n => n.endsWith("/" + cap)); if (!f) continue;
      const { head, rows } = parseCsv(await zip.file(f).async("string"));
      tables[id] = { id, caption: cap, head, types: head.map(typeOf), rows: rows.map(r => r.map((v, i) => conv(typeOf(head[i]), v))) };
    }
    const ids = Object.keys(tables), base = ids.reduce((b, id) => tables[id].rows.length > tables[b].rows.length ? id : b, ids[0]);
    // mock join: base rows, each link's other table looked up by key (keys unique in these CSVs)
    let cols = tables[base].head.map((h, i) => ({ t: base, i })), rowsJ = tables[base].rows.map(r => ({ [base]: r }));
    const done = new Set([base]); let more = true;
    while (more) { more = false; for (const l of rel.links) { const [h, a] = done.has(l.a) && !done.has(l.b) ? [l.a, l.b] : done.has(l.b) && !done.has(l.a) ? [l.b, l.a] : [null]; if (!h) continue;
      const [kh, ka] = l.a === h ? l.pairs[0] : [l.pairs[0][1], l.pairs[0][0]]; const ih = tables[h].head.indexOf(kh), ia = tables[a].head.indexOf(ka);
      const idx = new Map(tables[a].rows.map(r => [r[ia], r])); rowsJ.forEach(o => { o[a] = o[h] ? idx.get(o[h][ih]) || null : null; }); done.add(a); more = true; } }
    s.cols = ids.flatMap(id => tables[id].head.map((h, i) => ({ name: h, caption: capOf(h), type: tables[id].types[i], t: id, i })));
    s.rows = rowsJ.map(o => s.cols.map(c => (o[c.t] ? o[c.t][c.i] : null)));
    const cell = (t, v) => v === null ? { value: null, nativeValue: null, formattedValue: "Null" } : v instanceof Date ? { value: iso(v), nativeValue: v, formattedValue: iso(v) } : dv(v);
    s.dataSource = { id: dsId, name: "E-commerce", fields: s.cols.map(c => ({ name: c.caption })),
      getLogicalTablesAsync: async () => ids.map(id => ({ id, caption: tables[id].caption })),
      getLogicalTableDataAsync: async id => ({ columns: tables[id].head.map((h, i) => ({ fieldName: capOf(h), dataType: tables[id].types[i], index: i })),
        data: tables[id].rows.map(r => r.map((v, i) => cell(tables[id].types[i], v))), totalRowCount: tables[id].rows.length }) };
    log(`[mock] ${dsId}: ${ids.length} CSV tables, ${s.rows.length} joined rows`);
  }
  if (process.env.REALXLSX) {
    const f = Object.keys(zip.files).find(n => /^Data\/.*\.xlsx$/i.test(n));
    const wbx = new ExcelJS.Workbook(); await wbx.xlsx.load(await zip.file(f).async("nodebuffer"));
    const sh = wbx.worksheets[0];
    const head = sh.getRow(1).values.slice(1).map(String);
    const dsId = Object.keys(sources).find(id => /^federated/.test(id)) || Object.keys(sources)[0];
    const s = sources[dsId];
    const fieldOf = name => Object.values(model.fields).find(x => x.ds === dsId && x.name === name) || {};
    const typeOf = name => { const t = fieldOf(name).datatype || "string"; return t === "real" ? "float" : t === "integer" ? "int" : t; };
    s.cols = head.map((h, i) => ({ name: h, caption: fieldOf(h).caption || h, type: typeOf(h), i }));
    s.rows = [];
    sh.eachRow((row, n) => { if (n === 1) return; s.rows.push(s.cols.map((c, i) => { let v = row.getCell(i + 1).value; if (v && typeof v === "object" && "result" in v) v = v.result;
      if (v === null || v === undefined || v === "") return null; if (/date/.test(c.type)) return v instanceof Date ? v : new Date(String(v).slice(0, 10) + "T00:00:00Z");
      return c.type === "int" || c.type === "float" ? Number(v) : String(v); })); });
    const cell = v => v === null ? { value: null, nativeValue: null, formattedValue: "Null" } : v instanceof Date ? { value: iso(v), nativeValue: v, formattedValue: iso(v) } : dv(v);
    s.dataSource = { id: dsId, name: "Data", fields: s.cols.map(c => ({ name: c.caption })), getLogicalTablesAsync: async () => [{ id: "Sheet1", caption: "Sheet1" }],
      getLogicalTableDataAsync: async () => ({ columns: s.cols.map((c, i) => ({ fieldName: c.caption, dataType: c.type, index: i })), data: s.rows.map(r => r.map(cell)), totalRowCount: s.rows.length }) };
    log(`[mock] ${dsId}: ${s.rows.length} rows from ${f}`);
  }
  if (process.env.NOJOIN) model.relationships = {};
  const colOf = ref => { const s = sources[ref.ds]; if (!s) return null; const i = s.cols.findIndex(c => c.name === ref.name || c.caption === ref.name); return i < 0 ? null : { s, i, c: s.cols[i] }; };
  const caption = ref => { const info = T.tfFieldInfo(model, ref); return (info && info.caption) || ref.name; };

  // ── summary data: the sheet's source-column dimensions, and aggregates of source columns
  // SUMMARY=<file.json>: { sheet: { columns, data } } – a sheet's summary data given (a .hyper extract this mock cannot read)
  const given = process.env.SUMMARY ? JSON.parse(fs.readFileSync(process.env.SUMMARY, "utf8")) : {};
  const summary = name => {
    if (given[name]) return { ...given[name], totalRowCount: given[name].data.length };
    const sh = model.sheets[name];
    if (!sh) return { columns: [], data: [] };
    const refs = [...(sh.rows || []), ...(sh.cols || []), ...(sh.panes || []).flatMap(p => p.encodings.map(e => e.field))].filter(Boolean);
    const dims = [], meas = [], seen = new Set();
    /** a measure: an aggregate of a column, or a calculation evaluated on the rows (null: neither) */
    const measOf = r => {
      const c = colOf(r), d = String(r.deriv || "none").toLowerCase();
      if (!c && d === "usr") {
        const info = T.tfFieldInfo(model, r), s0 = sources[r.ds];
        if (!info || !info.formula || !s0) return null;
        const cell = v => v === null || v === undefined ? { value: null, nativeValue: null, formattedValue: "Null" } : v instanceof Date ? { value: iso(v), nativeValue: v, formattedValue: iso(v) } : dv(v);
        s0.evalSrc = s0.evalSrc || { id: s0.id, name: s0.id, columns: s0.cols.map(k => ({ name: k.caption, dataType: k.type })), rows: s0.rows.map(row => row.map(cell)) };
        const e = T.slicerParseCalc(s0.evalSrc, model, info.formulaFull || info.formula, 0);
        return e ? { r, c: { s: s0 }, agg: "usr", e } : null;
      }
      return c && AGG[d] ? { r, c, agg: d } : null;
    };
    refs.forEach(r => {
      if (seen.has(r.inner)) return; seen.add(r.inner);
      const c = colOf(r), d = String(r.deriv || "none").toLowerCase();
      const m = measOf(r);
      if (m) meas.push(m);
      else if (c && (d === "none" || d === "attr" || /^(yr|qr|mn|dy|wd|wk|tyr|tqr|tmn|twk|tdy)$/.test(d))) dims.push({ r, c, d });
    });
    // Measure Names on the view: Tableau gives a row per member of its Measure Names filter (Measure Values: that
    // member's value, the axis measures beside it), then a row of the sheet's other measures
    const members = refs.some(r => /^Measure Names$/i.test(r.name))
      ? (sh.measureFilter || []).map(m => T.tfParseFieldRef(String(m).replace(/^"|"$/g, ""))).map(r => r && measOf(r)).filter(Boolean) : [];
    const axis = new Set([...(sh.rows || []), ...(sh.cols || [])].filter(Boolean).map(r => r.inner));
    if (!dims.length && !meas.length) return { columns: [], data: [] };
    const s = (dims[0] || meas[0]).c.s;
    const groups = new Map(), idx = new Map();
    const txt = v => (v instanceof Date ? iso(v) : String(v));
    const kept = row => [...active.entries()].every(([fn, set]) => { const i = s.cols.findIndex(c => c.caption === fn || c.name === fn); return i < 0 || set.has(txt(row[i])); });
    s.rows.forEach((row, ri) => {
      if (!kept(row)) return;
      const part = (x, v) => { if (!(v instanceof Date) || !/^(yr|qr|mn|dy|wd|wk|tyr|tqr|tmn|twk|tdy)$/.test(x.d)) return v; const y = v.getUTCFullYear(), m = v.getUTCMonth();
        return x.d === "yr" ? y : x.d === "qr" ? Math.floor(m / 3) + 1 : x.d === "mn" ? m + 1 : x.d === "dy" ? v.getUTCDate() : x.d === "tyr" ? new Date(Date.UTC(y, 0, 1)) : x.d === "tqr" ? new Date(Date.UTC(y, m - m % 3, 1)) : x.d === "tmn" ? new Date(Date.UTC(y, m, 1))
          // WEEKDAY as Tableau's summary data gives it (the day's name), WEEK as its number, a week by its Sunday
          : x.d === "wd" ? ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"][v.getUTCDay()]
          : x.d === "wk" ? Math.floor(((v - Date.UTC(y, 0, 1)) / 86400000 + new Date(Date.UTC(y, 0, 1)).getUTCDay()) / 7) + 1
          : x.d === "twk" ? new Date(v.getTime() - v.getUTCDay() * 86400000) : v; };
      const k = JSON.stringify(dims.map(x => x.c.s === s ? part(x, row[x.c.i]) : null));
      (groups.get(k) || groups.set(k, []).get(k)).push(row);
      (idx.get(k) || idx.set(k, []).get(k)).push(ri);
    });
    const NUL = { value: null, nativeValue: null, formattedValue: "Null" };
    const columns = [...dims.map(x => ({ fieldName: caption(x.r), fieldId: `[${x.r.ds}].[${x.r.inner}]`, dataType: x.c.c.type })),
                     ...(members.length ? [{ fieldName: "Measure Names", fieldId: "[:Measure Names]", dataType: "string" }] : []),
                     ...meas.map(x => ({ fieldName: `${x.agg === "usr" ? "AGG" : AGG[x.agg]}(${caption(x.r)})`, fieldId: `[${x.r.ds}].[${x.r.inner}]`, dataType: "float" })),
                     ...(members.length ? [{ fieldName: "Measure Values", fieldId: "[Multiple Values]", dataType: "float" }] : [])]
      .map((c, i) => ({ ...c, index: i }));
    const valueOf = (x, k, rows) => {
      if (x.agg === "usr") { const v = T.slicerEval(x.e, idx.get(k), s.evalSrc || x.c.s.evalSrc, null); return v === null ? NUL : dv(v); }
      // a blended (secondary) data source's measure: this mock groups only the primary's rows
      if (x.c.s !== s) return NUL;
      const vals = rows.map(r => r[x.c.i]).filter(v => v !== null && v !== undefined);
      const num = vals.map(Number);
      // MIN / MAX of a string field: the first / last name, as Tableau gives it
      const text = vals.some(v => typeof v === "string") ? vals.map(String).sort() : null;
      const v = x.agg === "sum" ? num.reduce((a, b) => a + b, 0) : x.agg === "avg" ? num.reduce((a, b) => a + b, 0) / num.length
        : x.agg === "min" ? (text ? text[0] : Math.min(...num)) : x.agg === "max" ? (text ? text[text.length - 1] : Math.max(...num))
        : x.agg === "cnt" ? vals.length : new Set(vals.map(String)).size;
      return dv(v);
    };
    const data = [...groups.entries()].flatMap(([k, rows]) => {
      const head = JSON.parse(k).map(v => { const raw = typeof v === "string" && /^\d{4}-\d\d-\d\dT/.test(v) ? new Date(v) : v; return raw instanceof Date ? { value: iso(raw), nativeValue: raw, formattedValue: iso(raw) } : raw === null ? NUL : dv(raw); });
      const values = meas.map(x => valueOf(x, k, rows));
      if (!members.length) return [[...head, ...values]];
      return [...members.map(m => [...head, dv(caption(m.r)), ...meas.map((x, j) => (axis.has(x.r.inner) ? values[j] : NUL)), valueOf(m, k, rows)]),
              [...head, NUL, ...values, NUL]];
    });
    return { columns, data, totalRowCount: data.length };
  };

  // ── the sheet's filters as the Extensions API reports them: its <filter> elements and its dashboard cards' fields
  const build = DNAME => {
  const dash = model.dashboards[DNAME];
  // ZONEPOS: zone id → [x, y, w, h] (the workbook's 0–100000 units), for a layout the .twb does not hold
  if (process.env.ZONEPOS && dash) Object.entries(JSON.parse(process.env.ZONEPOS)).forEach(([id, [x, y, w, h]]) => {
    dash.zones.filter(z => String(z.id) === id).forEach(z => Object.assign(z, { x, y, w, h }));
  });
  if (!dash) { log("no dashboard", DNAME, "– have:", Object.keys(model.dashboards).join(" | ")); process.exit(2); }
  const wsXml = {};
  Array.from(doc.getElementsByTagName("worksheet")).forEach(w => { wsXml[w.getAttribute("name")] = w; });
  const filtersOf = name => {
    const w = wsXml[name];
    if (!w) return [];
    const cols = new Map();
    Array.from(w.getElementsByTagName("filter")).forEach(f => cols.set(f.getAttribute("column"), f));
    dash.zones.filter(z => z.type === "filter" && z.name === name && z.param && !cols.has(z.param)).forEach(z => cols.set(z.param, null));
    const out = [];
    cols.forEach((f, column) => {
      const ref = T.tfParseFieldRef(column || "");
      if (!ref || /Measure Names/i.test(ref.name)) return;
      const d = String(ref.deriv || "none").toLowerCase();
      const c = colOf(ref);
      const cls = f ? f.getAttribute("class") : (c && /date|int|float|real/.test(c.c.type) && ref.type === "qk" ? "quantitative" : "categorical");
      const fieldName = AGG[d] ? `${AGG[d]}(${caption(ref)})` : d !== "none" && d !== "attr" ? `${d.toUpperCase()}(${caption(ref)})` : caption(ref);
      const base = { fieldName, fieldId: `[${ref.ds}].[${ref.inner}]`, worksheetName: name };
      if (cls === "quantitative") {
        let lo, hi;
        if (c && !AGG[d]) { const v = c.s.rows.map(r => r[c.i]); lo = v.reduce((a, b) => (a < b ? a : b)); hi = v.reduce((a, b) => (a > b ? a : b)); }
        const wrap = v => v === undefined ? { value: NaN, nativeValue: NaN, formattedValue: "Null" } : v instanceof Date ? { value: iso(v), nativeValue: v, formattedValue: iso(v) } : dv(v);
        out.push({ ...base, filterType: "range", minValue: wrap(lo), maxValue: wrap(hi), includeNullValues: false });
      } else if (cls === "relative-date") {
        out.push({ ...base, filterType: "relative-date", periodType: (f.getAttribute("period-type-v2") || f.getAttribute("period-type") || "year").toLowerCase(),
                   rangeType: (f.getAttribute("range-type") || "last").toLowerCase().replace("-", ""), rangeN: +(f.getAttribute("range-n") || 1) });
      } else {
        const vals = c ? [...new Set(c.s.rows.map(r => r[c.i]).map(v => v instanceof Date ? iso(v) : String(v)))] : ["A", "B"];
        let applied = vals, all = true;
        if (SELECT) { const [fn, v] = SELECT.split("="); if (fn === fieldName || fn === caption(ref)) { applied = [v]; all = false; } }
        out.push({ ...base, filterType: "categorical", isExcludeMode: false, isAllSelected: all, appliedValues: applied.map(v => dv(v)),
                   getDomainAsync: async () => ({ values: vals.map(v => dv(v)) }) });
      }
    });
    return out;
  };

  const W = dash.width || 1200, H = dash.height || 800;
  const kind = { worksheet: "worksheet", filter: "quick-filter", text: "text", bitmap: "image", paramctrl: "parameter-control", color: "legend" };
  const runText = z => (z.runs || []).map(r => r.text).join("").trim();
  const objects = dash.zones.filter(z => kind[z.type] && z.w && z.h).map(z => ({
    id: Number(z.id), type: kind[z.type], isVisible: !z.hidden, isFloating: false,
    name: z.type === "worksheet" ? z.name : z.type === "paramctrl" ? runText(z) || z.param : `${z.type} ${z.id}`,
    position: { x: z.x / 100000 * W, y: z.y / 100000 * H }, size: { width: z.w / 100000 * W, height: z.h / 100000 * H } }));
  const names = [...new Set(dash.zones.filter(z => z.type === "worksheet" && z.name).map(z => z.name))];
  const dsOf = name => {
    const sh = model.sheets[name] || {};
    const refs = [...(sh.rows || []), ...(sh.cols || [])].filter(Boolean);
    const ids = [...new Set([...refs.map(r => r.ds), ...dash.zones.filter(z => z.type === "filter" && z.name === name && z.param).map(z => (T.tfParseFieldRef(z.param) || {}).ds)])].filter(id => sources[id]);
    return ids.map(id => sources[id].dataSource);
  };
  const worksheets = names.map(name => ({ name, sheetType: "worksheet", getSummaryDataAsync: async () => { globalThis.__reads++; return summary(name); },
    applyFilterAsync: async (fieldName, values) => { active.set(fieldName, new Set(values.map(v => (v instanceof Date ? iso(v) : String(v))))); return fieldName; },
    clearFilterAsync: async fieldName => { active.delete(fieldName); return fieldName; },
    getFiltersAsync: async () => filtersOf(name), getSelectedMarksAsync: async () => ({ data: [] }), getDataSourcesAsync: async () => dsOf(name) }));
  // the workbook's parameters as the Extensions API gives them: name = caption, a number's value with its format
  const params = Object.entries(model.fields).filter(([k, f]) => /^Parameters\|/.test(k) && f.param).map(([, f]) => {
    const raw = String(f.value ?? "").replace(/^"|"$/g, ""), num = /^(real|integer|float|int)$/i.test(f.datatype || "") ? Number(raw) : NaN;
    const nf = isFinite(num) && f.defaultFormat ? T.tableauToExcelNumFmt(f.defaultFormat) : null;
    return { id: `[Parameters].[${f.name}]`, name: f.caption || f.name, dataType: f.datatype,
             currentValue: isFinite(num) ? { value: String(num), nativeValue: num, formattedValue: nf ? T.tfFormatNumber(num, nf) : String(num) } : dv(raw) };
  });
  return { name: DNAME, objects, worksheets, getParametersAsync: async () => params };
  };
  const names0 = ALL ? Object.keys(model.dashboards) : [DASH];
  const first = build(names0[0]);
  globalThis.__collect = async current => ({ targets: [current, ...names0.slice(1).map(n => { const d = build(n); return { dashboard: d, sheets: d.worksheets }; })], notes: [] });
  tableau.extensions.dashboardContent.dashboard = first;
  await T.exportToExcel();
  for (let i = 0; i < 100 && !saved; i++) await new Promise(r => setTimeout(r, 50));
  if (!saved) { log("NO FILE – " + document.getElementById("export_status").textContent); process.exit(1); }
  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, saved);
  log(`wrote ${OUT}`);
})().catch(e => { console.error(e); process.exit(1); });
