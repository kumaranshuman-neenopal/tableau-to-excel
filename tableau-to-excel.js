/* ══════════════════════════════════════════════════════════════════════════════════════════════════════
 * tableau-to-excel.js — Tableau dashboard extension: Export to Excel (NeenOpal)
 * ──────────────────────────────────────────────────────────────────────────────────────────────────────
 * ONE file, merged from two code bases that grew out of the same v0.6.0 build_table_copy.js (this file's former name):
 *   • the visual export engine (Export_extension_XML / Export2Sheet, src/ modules): tables, KPI cards,
 *     text boxes, images, backgrounds, Tableau-exact layout, NATIVE EXCEL CHARTS (bar, line, area, pie,
 *     donut, scatter, bubble, combo, treemap, waterfall, box plot, Gantt …), Conversion Report sheet;
 *   • the Tableau Cloud features of the former build_table_copy.js: the workbook loaded automatically through the
 *     backend (Worker), every visible dashboard read through hidden Embedding API views, the one-time
 *     Tableau sign-in, and ONE Excel sheet per dashboard.
 *
 * Each "═══ <path> ═══" banner below starts one of the original modules (same names, same order as the
 * module tree, so a change can be traced back to it). The src/ modules are folded in up to "Donut centre
 * totals and waterfall grand totals" (eab0bfb); this file is now the only source. Sections that belong to the merge:
 *   cloud/backend.js      auto-load from Tableau Cloud (Advanced → Backend URL)
 *   cloud/dashboards.js   the other dashboards (hidden views, sign-in, progress screen, sheet names)
 *   ui/workbook-store.js  applyWorkbookXml(): manual 📁 load and auto-load share one path
 *   export/export.js      exportToExcel() → writeDashboardSheet() per dashboard → charts injected per sheet
 *   export/report.js      writeConversionReports(): one report for every exported dashboard
 *   main.js               panel wiring (start-up)
 * Settings that change the output: FORMAT_CONFIG (config.js section).
 *
 * The page (index.html) must load, before this file:
 *   ./js/tableau.extensions.1.latest.js, ExcelJS 4.4 (window.ExcelJS) and JSZip 3.10 (window.JSZip).
 * Panel element ids used: twb_file_label, load_workbook_btn, export_button, export_status,
 *   export_all (checkbox "All dashboards"), backend_url, backend_key, backend_save (Advanced settings).
 * ══════════════════════════════════════════════════════════════════════════════════════════════════════ */
"use strict";
(function () {

if (typeof ExcelJS === "undefined" || typeof JSZip === "undefined") {
  console.error("[Export] ExcelJS and JSZip must be loaded (script tags in index.html) before tableau-to-excel.js");
}

/* ═══ format/number-format.js ═════════════════════════════════════════════════════════════════════ */
/* Tableau number formats → Excel number formats. */

/* Tableau text-format → Excel numFmt.
 * - strips Tableau's type prefix (n/c/p/e/*); locale "standard" codes (C1033…) → null (inferred instead),
 *   while short ones are formats ("p0%" = whole percent)
 * - moves thousands-scaling commas behind the decimals: "#,##0,.0K" → "#,##0.0,\"K\""
 * - quotes every literal letter (Excel rejects bare K, M, yrs …)
 * - validates the result; anything doubtful → null so we never write an invalid format */
function tableauToExcelNumFmt(raw) {
  if (!raw) return null;
  let s = String(raw);
  if (/^[A-Za-z]\d{3,}%?$/.test(s)) return null;
  if (/^[ncpes*]/i.test(s) && !/^[#0]/.test(s)) s = s.slice(1);
  const sections = tfSplitSections(s);
  if (!sections || sections.length > 4) return null;
  const out = sections.map(tfConvertSection);
  if (out.some(x => x === null)) return null;
  const res = out.join(";");
  return /[0#]/.test(res) ? res : null;
}

/**
 * A number as Excel shows it in a format, for text Excel does not format itself (label texts taken from
 * cells): literal text, digits with thousands separators, decimals (0 required, # optional), thousands
 * scaling (trailing commas: ,, = millions), percent; sections positive;negative;zero.
 * @param {number} value @param {string} fmt Excel number format @returns {string}
 */
function tfFormatNumber(value, fmt) {
  if (typeof value !== "number" || !isFinite(value)) return value == null ? "" : String(value);
  const secs = tfSplitSections(String(fmt || "General")) || ["General"];
  let sec = secs[0], sign = value < 0 ? "-" : "";
  if (value < 0 && secs.length > 1) { sec = secs[1]; sign = ""; }   // the negative section shows the size
  else if (value === 0 && secs.length > 2) sec = secs[2];
  const x = Math.abs(value);
  let pre = "", post = "", pattern = "", phase = 0;                 // 0 before the number, 1 in it, 2 after
  const text = t => { if (phase === 1) phase = 2; if (phase === 0) pre += t; else post += t; };
  for (let i = 0; i < sec.length; i++) {
    const ch = sec[i];
    if (ch === '"') { const j = sec.indexOf('"', i + 1); text(sec.slice(i + 1, j < 0 ? sec.length : j)); i = j < 0 ? sec.length : j; }
    else if (ch === "\\") { text(sec[i + 1] || ""); i++; }
    else if (ch === "[") { const j = sec.indexOf("]", i); i = j < 0 ? sec.length : j; }      // [Red], [>100]
    else if (ch === "_") { text(" "); i++; }                                                 // _) = a space
    else if (ch === "*") i++;                                                                 // fill character
    else if (/[0#?.,%]/.test(ch) && phase < 2) { phase = 1; pattern += ch; }
    else if (/general/i.test(sec.slice(i, i + 7)) && phase === 0) { pattern = "G"; phase = 1; i += 6; }
    else text(ch);
  }
  if (pattern === "G" || !/[0#?]/.test(pattern)) {
    return sign + pre + (pattern === "G" ? String(+x.toPrecision(10)) : "") + post;
  }
  const pct = (pattern.match(/%/g) || []).length;
  const p = pattern.replace(/%/g, "");
  const dot = p.indexOf(".");
  let ip = dot >= 0 ? p.slice(0, dot) : p, dp = dot >= 0 ? p.slice(dot + 1) : "";
  let scale = 0;                                                   // commas after the last digit: ÷ 1000 each
  const it = ip.match(/,+$/); if (it) { scale += it[0].length; ip = ip.slice(0, -it[0].length); }
  const dt = dp.match(/,+$/); if (dt) { scale += dt[0].length; dp = dp.slice(0, -dt[0].length); }
  const minDec = (dp.match(/0/g) || []).length, maxDec = (dp.match(/[0#?]/g) || []).length;
  let s = (x / Math.pow(1000, scale) * Math.pow(100, pct)).toFixed(maxDec);
  if (maxDec > minDec) s = s.replace(new RegExp(`0{1,${maxDec - minDec}}$`), "").replace(/\.$/, "");
  let [whole, frac] = s.split(".");
  const minInt = (ip.match(/0/g) || []).length;
  if (whole.length < minInt) whole = whole.padStart(minInt, "0");
  if (!minInt && whole === "0" && frac) whole = "";
  if (ip.includes(",")) whole = whole.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return sign + pre + whole + (frac ? "." + frac : "") + "%".repeat(pct) + post;
}

/** an Excel number format with literal text around the number in every section ("4.0" → "4.0 DAYS")
 * @param {string} fmt @param {string} prefix @param {string} suffix @returns {string} */
function tfWrapNumFmt(fmt, prefix, suffix) {
  const q = s => s ? `"${s.replace(/"/g, "")}"` : "";
  return (tfSplitSections(fmt || "General") || ["General"]).slice(0, 3).map(sec => q(prefix) + (sec || "General") + q(suffix)).join(";");
}

function tfSplitSections(s) {
  const secs = []; let cur = "", q = false;
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '"') q = !q;
    if (ch === "\\" && !q) { cur += ch + (s[i + 1] || ""); i++; continue; }
    if (ch === ";" && !q) { secs.push(cur); cur = ""; continue; }
    cur += ch;
  }
  if (q) return null;                        // unbalanced quotes
  secs.push(cur);
  return secs;
}

function tfConvertSection(sec) {
  // tokens: quoted literal | escaped char | number pattern | other char
  const toks = [];
  const re = /"[^"]*"|\\.|[#0?,.]+(?:[eE][+-][0#]+)?|[\s\S]/g;
  let m, numDone = false;
  while ((m = re.exec(sec))) {
    const t = m[0];
    if (t[0] === '"' || t[0] === "\\") { toks.push(t); continue; }
    if (/^[#0?,.]/.test(t) && /[#0?]/.test(t) && !numDone) {
      numDone = true;
      const [intPart, ...decParts] = t.split(".");
      const dec = decParts.join("");
      const intTrail = (intPart.match(/,+$/) || [""])[0];                 // "#,##0," → scale
      const decScale = (dec.match(/,+$/) || [""])[0];
      const intClean = intPart.slice(0, intPart.length - intTrail.length);
      const decClean = dec.replace(/,+$/, "");
      toks.push(intClean + (decParts.length ? "." + decClean : "") + intTrail + decScale);
      continue;
    }
    if (/[A-Za-z]/.test(t)) {                                              // literal letter → quote
      toks.push('"' + t + '"');
      continue;
    }
    toks.push(t);
  }
  // merge adjacent quoted literals: "K""g" → "Kg"
  return toks.join("").replace(/""/g, "");
}

/* Parse a Tableau formatted value, build an Excel format and only accept it
 * if it reproduces the same number (round-trip check). */
function inferExcelNumFmt(formatted, value) {
  if (formatted == null || typeof value !== "number" || !isFinite(value)) return null;
  const f = String(formatted).trim();
  const m = f.match(/^(\(?)([^\d(]*?)(-?)(\d[\d.,\s\u00A0']*)(.*?)(\)?)$/);
  if (!m) return null;
  const [, lpar, prefixRaw, minus, core, suffixRaw, rpar] = m;
  const neg = !!(lpar && rpar) || !!minus || /^-/.test(prefixRaw);
  const prefix = prefixRaw.replace(/-/g, "").trim() ? prefixRaw.replace(/-/g, "") : "";
  const suffix = suffixRaw;

  // decide decimal separator by trying both interpretations
  const tryParse = dec => {
    const grpSrc = dec === "." ? "[,\\s\\u00A0']" : "[.\\s\\u00A0']";
    const clean = core.trim().replace(new RegExp(grpSrc, "g"), "").replace(dec, ".");
    const n = parseFloat(clean);
    const decimals = clean.includes(".") ? clean.split(".")[1].length : 0;
    return { n, decimals, grouped: new RegExp(grpSrc).test(core.trim()) };
  };
  const unitScale = { "%": 100, K: 1e-3, M: 1e-6, B: 1e-9, Bn: 1e-9 };
  const unit = (suffix.trim().match(/^(%|K|M|Bn|B)\b/) || suffix.trim().match(/^%/) || [])[0] || null;
  const scaled = Math.abs(value) * (unit ? unitScale[unit] : 1);

  for (const dec of [".", ","]) {
    const p = tryParse(dec);
    if (!isFinite(p.n)) continue;
    const tol = 0.5 * Math.pow(10, -p.decimals) + 1e-9 * Math.max(1, scaled);
    let plainPercent = false;
    if (Math.abs(scaled - p.n) > tol) {
      if (unit === "%" && Math.abs(Math.abs(value) - p.n) <= tol) plainPercent = true; // value already ×100
      else continue;
    }
    // Tableau's "Automatic" format pads to ~6 significant digits ("348.000", "4.50000",
    // "24.86108%"); 3+ decimals are that padding, not a chosen format → keep what the value needs
    let decimals = p.decimals;
    if (decimals >= 3) {
      const frac = (String(p.n).split(".")[1] || "").length;
      decimals = Math.min(decimals, frac);
      if (decimals >= 3) decimals = Math.max(0, Math.min(decimals, 2 - Math.floor(Math.log10(Math.abs(p.n) || 1))));
    }
    let body = (p.grouped ? "#,##0" : "0") + (decimals ? "." + "0".repeat(decimals) : "");
    const q = t => t ? '"' + t.replace(/"/g, '""') + '"' : "";
    let rest = suffix;
    if (unit === "%" && !plainPercent) { body += "%"; rest = suffix.replace("%", ""); }
    else if (unit && unit !== "%") { body += ",".repeat({ K: 1, M: 2, B: 3, Bn: 3 }[unit]) + q(unit); rest = suffix.replace(unit, ""); }
    const pos = q(prefix) + body + q(rest);
    if (!neg) return pos;
    if (lpar && rpar) return pos + ";(" + pos + ")";
    return minus ? pos + ";" + q(prefix) + "-" + body + q(rest) : pos + ";-" + pos;
  }
  return null;
}


/* ═══ config.js ═══════════════════════════════════════════════════════════════════════════════════ */
/* Extension-wide settings and enumerations. */

/* ══════════════════════════════════════════════════════════════════════════
 * TABLEAU → EXCEL FORMAT ENGINE
 * ──────────────────────────────────────────────────────────────────────────
 * 1. parseTableauFormatting(xml)  → plain JSON "format model" (DOM based)
 * 2. createSheetFormatter(model, sheetName) → resolves the Tableau cascade
 *      defaults → workbook <style> → worksheet <style> → element → field
 *      → customized-label run
 * 3. tfExcel*() helpers → convert resolved props into ExcelJS styles
 * ══════════════════════════════════════════════════════════════════════════ */

const FORMAT_CONFIG = {
  substituteTableauFonts: true,   // "Tableau Book" etc. are only installed with Tableau Desktop
  tableauFontSubstitute: "Arial",
  writeNativeNumbers: true,       // write real numbers + Excel numFmt instead of text
  applyFallbackHeatmap: false,    // old red/green heatmap when Tableau has no color (not in Tableau → off)
  nativeCharts: true,             // bar/line/area/pie/scatter/combo → editable native Excel charts
  chartPolicy: "data",            // charts that cannot be drawn: "skip" | "data" (export their data as a plain table)
  visualsAsTables: false,         // default of the panel's "Charts as tables" switch (each workbook / browser remembers its choice)
  crosstabKeepsSparklines: true,  // with "Charts as tables": charts in zones shorter than SPARKLINE_MAX_HEIGHT_PX stay charts
                                  // (KPI trend lines – a table would not fit there)
  groupOverflowRows: true,        // collapse rows beyond ROW_GROUP_THRESHOLD into an expandable [+]/[-] group
  autoFilter: "largest",          // Excel allows ONE autofilter per sheet: "largest" table | "none"
  textBoxHeaders: true,           // dashboard text boxes over a table as its column headers (one box per column or per group of panes)
  textBoxTitle: false,            // text boxes are drawn in place now; true = also use the top one as the sheet title
  linkText: "url",                // URL-action cells: "url" shows the link itself, any other string is shown as the text
  sheetGridlines: false,          // Excel's cell grid on the dashboard sheet – Tableau dashboards have none
  printFitToWidth: true,          // printing / PDF: the dashboard one page wide, landscape when it is wider than tall
  tableauTextScale: null,         // Tableau label text vs its pixels (Windows display scaling) for pie label placement; null = the browser's devicePixelRatio
  conversionReport: true,         // a "Conversion Report" sheet: every visual, how it was converted (NATIVE … TABLE_FALLBACK) and why
  fallbackNotes: true,            // a note on each visual exported as its data table, naming the Tableau visual and the reason
  filterCards: true,              // dashboard filter controls → a card with their title and current selection ("(All)", "East, West" …)
  slicers: true,                  // a filter card of a field's values → an Excel slicer that filters every chart, table and KPI value
                                  // built from the same data source (their numbers become formulas over its rows, a hidden sheet)
  slicerMaxRows: 200000,          // a data source with more rows keeps its filter cards as text (the file would get too big and slow)
  dropdownCards: true,            // a Tableau dropdown card (Multiple Values (dropdown) too) → a dropdown cell, (All) or one value,
                                  // like the card; false: a slicer. A card with several – not all – values selected stays a slicer
  showTiming: false,              // true: also show the "⏱ dashboards … · sheets … · file …" line in the panel (always in the console)
  debug: false                    // true: dump the parsed workbook model, [Format] traces and full visual specs to the console
};

const VISUAL_TYPES = Object.freeze({
  TABLE: "TABLE",
  KPI: "KPI",
  HEATMAP: "HEATMAP",
  BAR: "BAR",
  COLUMN: "COLUMN",
  LINE: "LINE",
  AREA: "AREA",
  PIE: "PIE",
  SCATTER: "SCATTER",
  BUBBLE: "BUBBLE",
  COMBO: "COMBO",
  HISTOGRAM: "HISTOGRAM",
  TREEMAP: "TREEMAP",
  WATERFALL: "WATERFALL",
  BOXPLOT: "BOXPLOT",
  MAP: "MAP",
  MAP_FILLED: "MAP_FILLED",
  GANTT: "GANTT",
  CUSTOM: "CUSTOM",
  UNKNOWN: "UNKNOWN"
});

/* Approximate Tableau Desktop defaults. Compare against one of your exports
 * and tweak these once – everything explicitly set in the TWB overrides them. */
const TABLEAU_DEFAULTS = {
  worksheet:  { fontName: "Tableau Book", fontSize: 9, color: "FF333333" },
  fieldLabel: { bold: true },
  title:      { fontName: "Tableau Book", fontSize: 15, color: "FF333333" },
  dashTitle:  { fontName: "Tableau Book", fontSize: 18, color: "FF333333" },
  divider:    { visible: true, style: "thin", color: "FFD4D4D4", level: 1 }
};

/* XML element names used in <style-rule element='…'>. Arrays = aliases seen
 * across Tableau versions; unknown ones are simply ignored. */
const TF_ELEMENTS = {
  datalabel:  ["datalabel"],        // Marks card → Label → font (per pane)
  sheet:      ["all", "worksheet"],
  title:      ["title", "worksheet-title"],
  dashTitle:  ["dash-title", "dashboard-title", "title"],
  pane:       ["pane"],
  cell:       ["cell"],
  header:     ["header", "label"],   // "header" = box (shading/size), "label" = header TEXT (font/colour/underline/alignment)
  fieldLabel: ["field-labels", "field-labels-decoration"],
  divider:    ["table-div"]
};

const TF_DERIV_LABEL = {
  sum: "SUM", avg: "AVG", cnt: "CNT", ctd: "CNTD", min: "MIN", max: "MAX",
  med: "MEDIAN", attr: "ATTR", usr: "AGG", std: "STDEV", stdp: "STDEVP",
  var: "VAR", varp: "VARP", yr: "YEAR", qr: "QUARTER", mn: "MONTH", wk: "WEEK",
  dy: "DAY", hr: "HOUR", tyr: "YEAR", tqr: "QUARTER", tmn: "MONTH", twk: "WEEK", tdy: "DAY"
};


/* ═══ util.js ═════════════════════════════════════════════════════════════════════════════════════ */
/* Small shared helpers: logging, DOM child lookup, text/number normalisation, export file name. */

/* ── small helpers ─────────────────────────────────────────────────────── */
function tfLog(...a) { if (FORMAT_CONFIG.debug) console.log("[Format]", ...a); }

function tfKids(el, tag) {
  if (!el) return [];
  return Array.from(el.childNodes).filter(n => n.nodeType === 1 && (!tag || n.tagName === tag));
}

function tfKid(el, tag) { return tfKids(el, tag)[0] || null; }

function tfNorm(s) { return String(s == null ? "" : s).toLowerCase().replace(/\s+/g, " ").trim(); }

function tfNum(v) { const n = parseFloat(v); return isFinite(n) ? n : undefined; }

/** default name of the exported file: "<workbook> - <dashboard>.xlsx", without characters Windows forbids
 * @param {string | null} workbookFile the loaded .twb / .twbx @param {string} dashboardName @returns {string} */
function exportFileName(workbookFile, dashboardName) {
  const clean = s => String(s || "").replace(/[\\/:*?"<>|\u0000-\u001f]/g, "").replace(/\s+/g, " ").trim().replace(/[. ]+$/, "");
  const workbook = clean(String(workbookFile || "").replace(/\.twbx?$/i, ""));
  const dashboard = clean(dashboardName) || "Dashboard Export";
  const base = workbook && workbook.toLowerCase() !== dashboard.toLowerCase() ? `${workbook} - ${dashboard}` : dashboard;
  return base.slice(0, 150).replace(/[. ]+$/, "") + ".xlsx";
}

/** copy of o without its undefined properties
 * @template T @param {T} o @returns {T} */
function tfDefined(o) {
  const r = /** @type {T} */ ({});
  for (const k in o) if (o[k] !== undefined) r[k] = o[k];
  return r;
}


/* ═══ twb/field-ref.js ════════════════════════════════════════════════════════════════════════════ */
/* Tableau field references ("[ds].[sum:Sales:qk]") – parsing and comparison. */

/* ── Field references ──────────────────────────────────────────────────────
 * "[federated.0x].[sum:Sales:qk]" → {ds, deriv:"sum", name:"Sales", type:"qk"}
 * type: qk = continuous measure, nk/ok = discrete                          */
/** @param {string} ref @returns {FieldRef | null} */
function tfParseFieldRef(ref) {
  if (!ref) return null;
  ref = String(ref).trim().replace(/^"|"$/g, "");
  const parts = ref.match(/\[[^\]]*\]/g) || [];
  if (!parts.length) return { raw: ref, ds: null, deriv: null, name: ref, type: null, inner: ref };
  const inner = parts[parts.length - 1].slice(1, -1);
  const ds = parts.length > 1 ? parts[0].slice(1, -1) : null;
  const seg = inner.split(":");
  let deriv = null, name = inner, type = null;
  // "[usr:Calculation_1:ok:9]" → type is the nk/ok/qk segment, not the last one
  let ti = -1;
  for (let i = seg.length - 1; i >= 2; i--) if (/^(nk|ok|qk)$/.test(seg[i])) { ti = i; break; }
  if (ti >= 2) { deriv = seg[0] || null; type = seg[ti]; name = seg.slice(1, ti).join(":"); }
  else if (seg.length === 2 && seg[0] === "") name = seg[1];          // [:Measure Names]
  return { raw: ref, ds, deriv, name, type, inner };
}

/** a quick table calculation's measure: "[cum:sum:Sales:qk:7]" (running total) → "[sum:Sales:qk]"; null for any
 * other field @param {FieldRef} ref @returns {FieldRef | null} */
function tfTableCalcBase(ref) {
  if (!ref || !ref.deriv || !ref.type || !/^(sum|avg|cnt|ctd|min|max|med|attr|usr|std|stdp|var|varp|none):/i.test(ref.name || "")) return null;
  const inner = `${ref.name}:${ref.type}`;
  return tfParseFieldRef(ref.ds ? `[${ref.ds}].[${inner}]` : `[${inner}]`);
}

/** The field references in a shelf or label text. A relationship model's "Count of <table>" is one reference of
 * three parts: "[ds].[__tableau_internal_object_id__].[cnt:Orders.csv_…:qk]".
 * @param {string} text @returns {FieldRef[]} */
function tfExtractRefs(text) {
  if (!text) return [];
  const m = String(text).match(/(?:\[[^\]]+\]\.)?(?:\[__tableau_internal_object_id__\]\.)?\[[^\]]+\]/g) || [];
  return m.map(tfParseFieldRef).filter(Boolean);
}

function tfRefKey(r) { return ((r.ds || "") + "|" + r.inner).toLowerCase(); }

/** @param {Partial<FieldRef> | null} a @param {Partial<FieldRef> | null} b @returns {boolean} */
function tfSameField(a, b) {
  if (!a || !b) return false;
  if (a.inner.toLowerCase() === b.inner.toLowerCase()) return true;
  return tfNorm(a.name) === tfNorm(b.name) && (a.deriv || "none") === (b.deriv || "none");
}


/* ═══ charts/model/axes.js ════════════════════════════════════════════════════════════════════════ */
/* The worksheet's own axis and label settings on a chart spec: hidden axes, axis titles, Edit Axis ranges
 * and tick spacing, tick number formats, grid lines, axis rulers, mark label font, position and text. */

/* Edit Axis → the spec: fixed range ("fixed" / "fixedmin" / "fixedmax"), tick spacing, include zero */
function applySpace(spec, s, axis, noReverse = false) {
  const key = name => axis === "x" ? "x" + name : "value" + name;
  if (s.majorSpacing) spec[key("MajorUnit")] = s.majorSpacing;
  // e.g. a bump chart: rank 1 at the top (a butterfly's reversed wing is negated instead)
  if (s.reverse && !spec.mirrored && !noReverse) spec[axis === "x" ? "xReversed" : "valueReversed"] = true;
  if (/^fixed(min)?$/.test(s.rangeType || "") && s.min !== undefined) spec[key("Min")] = s.min;
  if (/^fixed(max)?$/.test(s.rangeType || "") && s.max !== undefined) spec[key("Max")] = s.max;
  if (s.domainExpand === "false" && axis !== "x") spec.includeZero = false;
}

/* an axis's tick format (Format → Axis → Numbers: the field's "label" text-format) */
const axisNumFmt = (fmt, ref) => (ref && tableauToExcelNumFmt(fmt.labelFormat(ref))) || undefined;

/* mark labels: the worksheet's label font, for bars the label alignment as Excel's label position, and a
 * label text around the value ("<AGG(Days)> DAYS") as literal text in the label's number format */
function applyLabels(spec, fmt, measures) {
  const dl = fmt.dataLabelStyle();
  if (dl.fontName || dl.fontSize || dl.color || dl.bold !== undefined) {
    spec.labelFont = { name: dl.fontName, size: dl.fontSize, color: dl.color ? dl.color.slice(2) : undefined, bold: dl.bold };
  }
  const tpl = fmt.labelTemplate();
  const at = tpl && (tpl.prefix.trim() || tpl.suffix.trim()) ? measures.findIndex(m => m && tfSameField(m, tpl.ref)) : -1;
  if (at === 0 || at === 1) {
    spec.series.forEach(s => {
      if (s.refLine || !!s.secondary !== (at === 1 && measures.length > 1) || s.labelTexts) return;
      s.labelNumFmt = tfWrapNumFmt(s.labelNumFmt || (s.secondary ? spec.secondaryNumFmt : spec.numFmt) || "General", tpl.prefix, tpl.suffix);
    });
  }
  if (spec.kind === "bar" || spec.kind === "combo") {
    const cell = fmt.markCellStyle(null);
    const pos = spec.barDir === "bar" ? { left: "inBase", center: "ctr", right: "inEnd" }[cell.hAlign]
                                     : { bottom: "inBase", middle: "ctr", top: "inEnd" }[cell.vAlign];
    if (pos) spec.labelPos = pos;
  }
}

/**
 * Label → Marks to Label: Tableau labels only some marks of each line / pane – Min/Max of a field (the earliest and
 * latest dates when the field is the category axis, else the lowest and highest values), Line Ends or Most Recent:
 * the points that keep their label (s.labelOnly).
 * @param {ChartSpec} spec @param {SheetFormatter} fmt @param {{ ref?: FieldRef | null } | undefined} catDim the innermost category field
 */
function applyLabelMarks(spec, fmt, catDim) {
  const m = fmt.markLabelsMarks();
  if (!m || !/^(range|line-ends|most-recent)$/.test(m.mode || "")) return;
  const onCategory = !m.field || !!(catDim && catDim.ref && tfSameField(tfParseFieldRef(m.field), catDim.ref));
  spec.series.forEach(s => {
    if (!s.labels || s.refLine || !s.values) return;
    const at = s.values.map((v, i) => num(v) === null ? -1 : i).filter(i => i >= 0);
    if (!at.length) return;
    const first = at[0], last = at[at.length - 1], keep = new Set();
    if (m.mode === "range" && !onCategory) {
      const byValue = [...at].sort((a, b) => /** @type {number} */ (s.values[a]) - /** @type {number} */ (s.values[b]));
      if (m.min) keep.add(byValue[0]);
      if (m.max) keep.add(byValue[byValue.length - 1]);
    } else if (m.mode === "range") {
      if (m.min) keep.add(first);
      if (m.max) keep.add(last);
    } else if (m.mode === "line-ends") {
      if (m.first) keep.add(first);
      if (m.last) keep.add(last);
    } else keep.add(last);
    s.labelOnly = [...keep];
  });
}

/**
 * Tableau hides axes ("Show Header" off), renames or removes axis titles (Edit Axis), and turns grid
 * lines and axis rulers off per worksheet; the Excel chart follows the same settings.
 * @param {ChartSpec} spec @param {ChartContext} ctx
 */
function tvApplyWorkbookAxes(spec, ctx) {
  const { vm, roles } = ctx;
  const fmt = vm.fmt;
  if (!fmt.hasModel || roles.source !== "twb" || spec.axesHidden || /^(pie|doughnut|treemap)$/.test(spec.kind)) return;

  if (spec.kind === "scatter" || spec.kind === "bubble") {
    const x = roles.cols.values[0], y = roles.rows.values[0];
    const xi = x ? fmt.axisInfo(x.ref, "cols", "0") : {}, yi = y ? fmt.axisInfo(y.ref, "rows", "0") : {};
    if (xi.hidden) spec.xAxisHidden = true;
    if (xi.title !== undefined) spec.xTitle = xi.title;
    if (yi.hidden) spec.valueAxisHidden = true;
    if (yi.title !== undefined) spec.valueTitle = yi.title;
    if (!fmt.gridlinesShown("rows")) spec.gridlines = false;
    if (!fmt.gridlinesShown("cols")) spec.xGridlines = false;
    if (x) applySpace(spec, fmt.axisSpace(x.ref, "cols", "0"), "x");
    if (y) applySpace(spec, fmt.axisSpace(y.ref, "rows", "0"), "value");
    spec.xAxisNumFmt = axisNumFmt(fmt, x && x.ref);
    spec.valueAxisNumFmt = axisNumFmt(fmt, y && y.ref);
    applyLabels(spec, fmt, [y && y.ref, x && x.ref]);
    return;
  }

  const valueShelf = roles.rows.values.length ? "rows" : roles.cols.values.length ? "cols" : null;
  if (!valueShelf) return;
  const catShelf = valueShelf === "rows" ? "cols" : "rows";
  const refs = roles[valueShelf].axisRefs.length ? roles[valueShelf].axisRefs : roles[valueShelf].values.map(v => v.ref);
  const primary = refs[0] ? fmt.axisInfo(refs[0], valueShelf, "0") : {};
  // Tick Marks: None leaves an axis with no labels – Tableau shows nothing there
  const noTicks = (ref, cls) => fmt.axisSpace(ref, valueShelf, cls).majorShow === "false";
  if (primary.hidden || (refs[0] && noTicks(refs[0], "0"))) spec.valueAxisHidden = true;
  if (primary.title !== undefined) spec.valueTitle = primary.title;
  if (refs[1] && spec.series.some(s => s.secondary)) {
    // class = which axis of that field: a second measure has its own first axis ("0"); the same field on
    // both axes has its second one ("1")
    const cls = tfSameField(refs[0], refs[1]) ? "1" : "0";
    const second = fmt.axisInfo(refs[1], valueShelf, cls);
    if (second.hidden || noTicks(refs[1], cls)) spec.secondaryAxisHidden = true;
    if (second.title !== undefined) spec.secondaryTitle = second.title;
    // Edit Axis → Fixed on the second axis: Excel's secondary axis gets the same range (else it scales
    // itself to the data and the line sits at a different height than in Tableau)
    const s2 = fmt.axisSpace(refs[1], valueShelf, cls);
    if (/^fixed(min)?$/.test(s2.rangeType || "") && s2.min !== undefined) spec.secondaryMin = s2.min;
    if (/^fixed(max)?$/.test(s2.rangeType || "") && s2.max !== undefined) spec.secondaryMax = s2.max;
  }
  if (!fmt.gridlinesShown(valueShelf)) spec.gridlines = false;
  // one axis reversed beside another that is not (nor synchronized to it) mirrors two panes – a centred funnel, a
  // butterfly – which one Excel axis cannot: the reversal is left out
  const mirror = refs.length > 1 && refs.slice(1).some(r => {
    const s = fmt.axisSpace(r, valueShelf, tfSameField(r, refs[0]) ? "1" : "0");
    return !s.reverse && !s.synchronized;
  });
  if (refs[0]) applySpace(spec, fmt.axisSpace(refs[0], valueShelf, "0"), "value", mirror);
  spec.valueAxisNumFmt = axisNumFmt(fmt, refs[0]);
  if (refs[1]) spec.secondaryAxisNumFmt = axisNumFmt(fmt, refs[1]);
  applyLabels(spec, fmt, refs);
  applyLabelMarks(spec, fmt, roles[catShelf].dims.slice(-1)[0]);
  // categories: a continuous pill draws an axis, discrete pills draw headers – hidden only if all are
  const dims = [...roles[valueShelf].dims.map(d => ({ d, shelf: valueShelf })), ...roles[catShelf].dims.map(d => ({ d, shelf: catShelf }))];
  const hidden = ({ d, shelf }) => !!d.ref && (d.continuous ? fmt.axisInfo(d.ref, shelf, "0").hidden === true : fmt.isLabelHidden(d.ref));
  if (dims.length && dims.every(hidden)) spec.categoryAxisHidden = true;
  // one header hidden among shown ones (Department shown, the colour's Race under it hidden): that level's labels
  // are left blank – Tableau draws no header row for it, and its many labels would overlap. A space, not "": with
  // empty cells Excel drops the multi-level axis, and Department sits under its first bar instead of its group
  else if (spec.categories && spec.categories.levels.length > 1) {
    const blank = new Set(roles[catShelf].dims.filter(d => d.ref && !d.continuous && fmt.isLabelHidden(d.ref))
      .map(d => spec.categories.names.indexOf(tvMeasureLabel(vm, d.ci))).filter(l => l >= 0));
    if (blank.size && blank.size < spec.categories.levels.length) {
      spec.categories = { ...spec.categories, levels: spec.categories.levels.map((lv, l) => (blank.has(l) ? lv.map(() => " ") : lv)),
                          blankLevels: [...blank] };
    }
  }
  // a continuous category (a date axis): Edit Axis → its title ("" = none – a sparkline's axis keeps its room for the dates)
  const axisDim = roles[catShelf].dims.find(d => d.ref && d.continuous);
  const catTitle = axisDim ? fmt.axisInfo(axisDim.ref, catShelf, "0").title : undefined;
  if (catTitle !== undefined) spec.categoryTitle = catTitle;
  // category labels: horizontal like Tableau's column headers, unless the header is rotated in the workbook
  const inner = roles[catShelf].dims.filter(d => d.ref && !d.continuous).slice(-1)[0];
  const rot = inner ? fmt.headerOrientation(inner.ref) : null;
  spec.categoryRotation = rot || 0;
  // Format → Cell width on the category header: Tableau gives every label that much room (a long axis scrolls),
  // so a label is cut short only if it does not fit that width – not the chart's width shared by all labels
  const slot = inner && catShelf === "cols" ? fmt.widthPx(inner.ref) : null;
  if (slot > 0) spec.categorySlotPx = slot;
  if (!fmt.axisLineShown(catShelf)) spec.axisLine = false;
}


/* ═══ data/values.js ══════════════════════════════════════════════════════════════════════════════ */
/* Tableau summary-data values: null checks, text/number access, natural + calendar ordering. */

/* ══════════════════════════════════════════════════════════════════════════
 * VIEW MODEL — rebuild the table Tableau actually draws from summary data
 *   1. pivot Measure Names / Measure Values into one column per measure
 *   2. merge stacked multi-pane rows (one block per pane → one row)
 *   3. sort rows like the view (manual sort, sort-by-measure, else natural)
 *   4. keep only visible columns, in visual order, with Tableau's header labels
 * ══════════════════════════════════════════════════════════════════════════ */
/** @param {DataValue} dv */
function tfIsNull(dv) {
  if (!dv) return true;
  const v = dv.nativeValue !== undefined ? dv.nativeValue : dv.value;
  return v === null || v === undefined || v === "%null%";
}

/** @param {DataValue} dv @returns {string} */
function tfDvText(dv) { return tfIsNull(dv) ? "" : String(dv.formattedValue != null ? dv.formattedValue : dv.value); }

/** @param {DataValue} dv @returns {number | null} */
function tfDvNum(dv) {
  if (tfIsNull(dv)) return null;
  const v = dv.nativeValue !== undefined ? dv.nativeValue : dv.value;
  if (typeof v === "number") return v;
  const n = parseFloat(String(v).replace(/[^0-9.\-eE]/g, ""));
  return isFinite(n) && /^[\s$€£¥(+-]*[\d.,]+/.test(String(v)) ? n : null;
}

/* discrete date parts arrive as names: "January" / "Jan" / "Monday" → calendar position */
const TF_MONTHS = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];

const TF_WEEKDAYS = ["sun", "mon", "tue", "wed", "thu", "fri", "sat"];

function tfCalendarRank(text) {
  const t = String(text || "").trim().toLowerCase();
  const m = t.match(/^(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sept?|oct|nov|dec)$/);
  if (m) return { kind: "month", rank: TF_MONTHS.indexOf(m[1].slice(0, 3)) };
  const d = t.match(/^(sunday|monday|tuesday|wednesday|thursday|friday|saturday|sun|mon|tue|wed|thu|fri|sat)$/);
  if (d) return { kind: "weekday", rank: TF_WEEKDAYS.indexOf(d[1].slice(0, 3)) };
  return null;
}

/** View order of two values: numbers, dates, ISO strings, month / weekday names, then text; nulls last.
 * @param {DataValue} a @param {DataValue} b @returns {number} */
function tfNaturalCompare(a, b) {
  const an = tfIsNull(a), bn = tfIsNull(b);
  if (an || bn) return an === bn ? 0 : an ? 1 : -1;           // nulls last
  const av = a.nativeValue !== undefined ? a.nativeValue : a.value;
  const bv = b.nativeValue !== undefined ? b.nativeValue : b.value;
  if (typeof av === "number" && typeof bv === "number") return av - bv;
  if (av instanceof Date && bv instanceof Date) return /** @type {any} */ (av) - /** @type {any} */ (bv);
  const iso = /^\d{4}-\d{2}-\d{2}/;
  if (typeof av === "string" && typeof bv === "string" && iso.test(av) && iso.test(bv)) return av < bv ? -1 : av > bv ? 1 : 0;
  const ra = tfCalendarRank(tfDvText(a)), rb = tfCalendarRank(tfDvText(b));   // MONTH / WEEKDAY names
  if (ra !== null && rb !== null && ra.kind === rb.kind) return ra.rank - rb.rank;
  return tfDvText(a).localeCompare(tfDvText(b), undefined, { numeric: true, sensitivity: "base" });
}


/* ═══ format/colors.js ════════════════════════════════════════════════════════════════════════════ */
/* Colour parsing and interpolation helpers. */

/* Tableau colours: "#rrggbb", "#rgb" or "#rrggbbaa" (aa=00 → transparent) → ARGB */
function tfArgb(v) {
  if (v == null) return undefined;
  let s = String(v).trim().replace(/^#/, "");
  if (/^[0-9a-f]{8}$/i.test(s)) {
    if (s.slice(6) === "00") return null;          // explicit "no fill"
    s = s.slice(0, 6);
  }
  if (/^[0-9a-f]{3}$/i.test(s)) s = s.split("").map(c => c + c).join("");
  if (!/^[0-9a-f]{6}$/i.test(s)) return undefined;
  return "FF" + s.toUpperCase();
}

function tfBrightness(argb) {
  const r = parseInt(argb.substring(2, 4), 16), g = parseInt(argb.substring(4, 6), 16), b = parseInt(argb.substring(6, 8), 16);
  return (r * 299 + g * 587 + b * 114) / 1000;
}

function tfInterpolate(c1, c2, t) {
  const p = (c, i) => parseInt(c.substring(i, i + 2), 16);
  const ch = i => Math.round(p(c1, i) + (p(c2, i) - p(c1, i)) * t).toString(16).padStart(2, "0");
  return ("FF" + ch(2) + ch(4) + ch(6)).toUpperCase();
}

function tfSample(colors, pos) {
  pos = Math.max(0, Math.min(1, pos));
  if (colors.length === 1) return colors[0];
  const seg = pos * (colors.length - 1);
  const i = Math.min(Math.floor(seg), colors.length - 2);
  return tfInterpolate(colors[i], colors[i + 1], seg - i);
}


/* ═══ format/palettes.js ══════════════════════════════════════════════════════════════════════════ */
/* Tableau built-in and automatic colour palettes. */

"use strict";

// ── TABLEAU BUILT-IN PALETTE LOOKUP ──────────────────────────────────────
const TABLEAU_BUILTIN_PALETTES = {
  "blue_10_0": ["#C7DDEA","#AFCFE1","#97C0D7","#7FAFCA","#689BC0","#5487B1","#4475A0","#356790","#2A5783"],
  "orange_10_0": ["#F3C184", "#F0AE62", "#EE9A42", "#EF882D", "#ED7420", "#E25F1D", "#CC531F", "#B54820", "#9E3D22"],
  "green_10_0": ["#B9D9AF", "#A3CF95", "#8BC57D", "#74BA67", "#5DAA56", "#4D984B", "#3D8743", "#31773F", "#24693D"],
  "red_10_0": ["#F3B8AB", "#EEA08E", "#EA8972", "#E9725B", "#EA5C4C", "#E6453C", "#D92C34", "#C71532", "#AE123A"],
  "purple_10_0": ["#ECC6E3", "#E2BCD8", "#D7AFCA", "#CB9FBC", "#BB85A8", "#AC759B", "#9F6B93", "#8E5B86", "#7C4D79"],
  "brown_10_0": ["#E8D5B4","#DEBE8A","#D8A66B","#D08F57","#C97A4B","#BF6740","#B45539","#AA4435","#9F3632"],
  "gray_10_0": ["#E5E5E5","#D4D6D8","#C0C4C8","#AAB0B6","#959DA5","#818A94","#6D7782","#5B6470","#49525E"],
  "gray_warm_10_0": ["#D8D1CE","#CBC2BE","#BBB0AB","#AB9F9A","#9A8E89","#887C77","#776B67","#685D59","#59504E"],
  "blue_teal_10_0": ["#B7D7D1","#9BC9C7","#80BBC0","#67ACC0","#529DBA","#448DAF","#3A7D9F","#336C91","#2C5985"],
  "orange_gold_10_0": ["#E8C85E","#EDB657","#F0A54A","#F08F32","#EF791F","#E96418","#D5531D","#BA4522","#9E3A26"],
  "green_gold_10_0": ["#E5C75A","#CDBE58","#B1B953","#97B64F","#7FB255","#67A957","#529B53","#348347","#146C36"],
  "red_gold_10_0": ["#E8C85A","#EDA951","#F08C4B","#EC7247","#E65E47","#DF4D47","#D33A45","#C32942","#B71D3E"],
  "orange_blue_diverging_10_0": ["#9E3D22","#B54820","#CC531F","#E25F1D","#ED7420","#F3D9BE","#97C0D7","#689BC0","#4475A0","#356790","#2B5C8A"],
  "red_green_diverging_10_0": ["#AE123A","#C61E3F","#D93443","#E94F4A","#F07A66","#F2B3A3","#8BC97D","#6DB65F","#539F50","#3B8447","#24693D"],
  "green_blue_diverging_10_0": ["#24693D","#347D46","#4E9854","#72B464","#9BCF89","#D4DDD9","#A9C9DC","#7DAACE","#5C8FBC","#4373A0","#2A5783"],
  "red_blue_diverging_10_0": ["#A90C38","#C71F3E","#DC393F","#EB5A4F","#F3A091","#E7E2DE","#B8D1E1","#86B0D1","#5F90BC","#44709C","#2E5A87"],
  "red_black_10_0": ["#AE123A","#C71F3E","#DC393F","#EB5A4F","#F3A091","#DDD9D5","#B9BCBC","#969DA1","#78818A","#606A75","#49525E"],
  "gold_purple_diverging_10_0": ["#AD9024","#B89B34","#C6AA50","#D3BA6D","#DDC892","#E3D7D1","#D7C1D2","#C9A5C3","#BB8AB2","#AC7299"],
  "red_green_gold_diverging_10_0": ["#BE2A3E","#D44344","#E75D49","#F07A47","#F2A14A","#E7C65A","#A8BE5E","#77AF5B","#55994E","#3B8746","#22763F"],
  "sunrise_sunset_diverging_10_0": ["#33608C","#556AA0","#7B67A6","#A664A2","#CD6C95","#EC7C79","#F2B15A","#ED8D46","#E56B44","#D24844","#B81840"],
  "orange_blue_white_diverging_10_0": ["#9E3D22","#B94B20","#D45A1D","#EC7420","#F3B562","#F5F1EC","#C6DDEA","#93BED8","#679BC1","#4677A5","#2B5C8A"],
  "red_green_white_diverging_10_0": ["#AE123A","#C8243F","#DC4947","#ED725E","#F3A38F","#F3F1EE","#B9DFAF","#8BC57D","#5DAA56","#3F8847","#24693D"],
  "green_blue_white_diverging_10_0": ["#24693D","#3E864B","#5BA557","#82C06F","#B7DFAE","#F2F3F1","#C7DDEA","#97C0D7","#689BC0","#4475A0","#2A5783"],
  "red_blue_white_diverging_10_0": ["#A90C38","#C71F3E","#DC393F","#EB5A4F","#F3A091","#F4F3F2","#C6DDEA","#93BED8","#679BC1","#4677A5","#2E5A87"],
  "red_black_white_diverging_10_0": ["#AE123A","#C8243F","#DC4947","#ED725E","#F3A38F","#F3F2F1","#D0D2D1","#A7ADB0","#838C93","#626C77","#49525E"],
  "tableau-blue-light": ["#EEF2F7","#E7EDF5","#DFE8F3","#D8E2F0","#D1DDEE","#CBD9ED","#C9D7F1","#C7D9F2","#C4D8F3"],
  "tableau-orange-light": ["#F6F2EE","#F7EBDD","#F8E3CC","#F9DBBE","#FAD5B3","#FBCFA8","#FCCB9F","#FFCC9E","#FFCC9E"],
  "tableau-orange-blue-light": ["#FFCC9E","#FAD1AB","#F4D8BC","#EEE0CC","#EAE4D8","#E8E8E8","#DFE7EF","#D5E0EC","#CBD9E9","#C7D9F1","#C4D8F3"],
  "tableau-map-blue-green": ["#F5F5C8","#EEF2B3","#E0EA9A","#CBE18F","#AFD695","#91CC9D","#72C3A8","#58BCB5","#41B7C4"],
  "tableau-map-temperatur": ["#529985","#669C76","#81A364","#A6B04E","#D2C63F","#F0D347","#F2C04A","#E7A24A","#D4824D","#C26B51"]
};

// ── HELPER: Get built-in palette colors by name ──────────────────────────
function getBuiltInPaletteColors(paletteName) {
  if (!paletteName) return null;
  
  if (TABLEAU_BUILTIN_PALETTES[paletteName]) {
    return TABLEAU_BUILTIN_PALETTES[paletteName];
  }
  
  const lowerName = paletteName.toLowerCase();
  for (const [key, colors] of Object.entries(TABLEAU_BUILTIN_PALETTES)) {
    if (key.toLowerCase() === lowerName) {
      return colors;
    }
  }
  
  const normalized = paletteName.toLowerCase().replace(/[-\s]+/g, '');
  for (const [key, colors] of Object.entries(TABLEAU_BUILTIN_PALETTES)) {
    const keyNormalized = key.toLowerCase().replace(/[-\s]+/g, '');
    if (keyNormalized.includes(normalized) || normalized.includes(keyNormalized)) {
      return colors;
    }
  }
  
  return null;
}

const AUTOMATIC_PALETTE_BY_MARK = {
  automatic:    "blue_teal_10_0",
  bar:          "blue_10_0",
  line:         "blue_10_0",
  area:         "blue_teal_10_0",
  square:       "blue_teal_10_0",
  circle:       "blue_10_0",
  shape:        "blue_10_0",
  text:         "blue_10_0",
  map:          "blue_teal_10_0",
  multipolygon: "blue_teal_10_0",
  pie:          "blue_10_0",
  ganttbar:     "blue_10_0",
  polygon:      "blue_10_0",
  density:      "blue_10_0",
  heatmap:      "blue_10_0"
};

const DEFAULT_AUTOMATIC_PALETTE = "blue_10_0";

function getAutomaticPaletteForMark(markClass) {
  const key = String(markClass || "Automatic").toLowerCase().replace(/[\s_-]+/g, "");
  const name = AUTOMATIC_PALETTE_BY_MARK[key] || DEFAULT_AUTOMATIC_PALETTE;
  if (!AUTOMATIC_PALETTE_BY_MARK[key]) {
    console.log(`[Auto Palette] Mark type "${markClass}" not in table, using ${name}`);
  }
  return {
    name,
    colors: TABLEAU_BUILTIN_PALETTES[name].map(c => "FF" + c.replace("#", "").toUpperCase())
  };
}

const TABLEAU_10 = ["FF4E79A7", "FFF28E2B", "FFE15759", "FF76B7B2", "FF59A14F",
                    "FFEDC948", "FFB07AA1", "FFFF9DA7", "FF9C755F", "FFBAB0AC"];


/* ═══ format/color-scale.js ═══════════════════════════════════════════════════════════════════════ */
/* Colour encodings → value-to-colour scales (categorical and continuous). */

/* ── Colour scales ─────────────────────────────────────────────────────── */
function tfBuildColorScale(fmt, enc, values) {
  const def = enc.def;
  if (enc.continuous) {
    const nums = values.filter(v => typeof v === "number" && isFinite(v));
    if (!nums.length) return null;
    const min = def && def.min !== undefined ? def.min : Math.min(...nums);
    const max = def && def.max !== undefined ? def.max : Math.max(...nums);
    let colors = def && def.customColors.length ? def.customColors : null;
    let diverging = /diverging/i.test((def && (def.paletteName || "")) + " " + (def && (def.paletteType || "")));
    if (!colors && def && def.paletteName && !/^automatic$/i.test(def.paletteName)) {
      const pc = getBuiltInPaletteColors(def.paletteName);
      if (pc) colors = pc.map(tfArgb);
    }
    if (!colors) {                     // Tableau "Automatic"
      if (min < 0 && max > 0) { colors = getBuiltInPaletteColors("orange_blue_diverging_10_0").map(tfArgb); diverging = true; }
      else colors = getAutomaticPaletteForMark(enc.markClass).colors;
    }
    if (def && def.reverse) colors = colors.slice().reverse();
    let center = def && def.center !== undefined ? def.center : null;
    if (diverging && center === null) center = (min < 0 && max > 0) ? 0 : (min + max) / 2;
    tfLog(`Gradient for ${enc.ref.inner}: ${colors.length} stops, range ${min}..${max}${center !== null ? ", center " + center : ""}`);
    return v => {
      if (typeof v !== "number" || !isFinite(v)) return null;
      let pos;
      if (center !== null) {
        const dev = Math.max(Math.abs(max - center), Math.abs(min - center)) || 1;
        pos = 0.5 + (v - center) / (2 * dev);       // symmetric, like "Use full colour range" = off
      } else pos = max === min ? 0.5 : (v - min) / (max - min);
      // Stepped Color (num-steps): the range in N equal bands, each one solid colour – evenly spaced over the
      // palette, ends included (2 steps on a diverging palette: its two end colours, below / above the centre)
      const steps = def && def.steps >= 2 ? Math.round(def.steps) : 0;
      if (steps) {
        const band = Math.min(steps - 1, Math.max(0, Math.floor(Math.max(0, Math.min(1, pos)) * steps)));
        return tfSample(colors, band / (steps - 1));
      }
      return tfSample(colors, pos);
    };
  }
  // categorical
  const map = {};
  if (def) {
    // one palette can hold a measure under several aggregations (SUM and AVG of the same calculation, each its own
    // colour): the entries of the measures this sheet shows (its Measure Names list) are applied last, so they win
    const bare = k => tfNorm(String(k).replace(/^"|"$/g, "").replace(/\\(.)/g, "$1"));
    const shown = new Set(((fmt.sheetModel || {}).measureFilter || []).map(bare));
    const entries = Object.entries(def.map);
    const mine = ([k]) => shown.has(bare(k));
    [...entries.filter(e => !mine(e)), ...entries.filter(mine)].forEach(([k, c]) => fmt.bucketAliases(k).forEach(a => { map[tfNorm(a)] = c; }));
  }
  const used = new Set(Object.values(map));
  const pool = TABLEAU_10.filter(c => !used.has(c));
  const auto = {};
  [...new Set(values.filter(v => v != null && v !== "").map(tfNorm))].filter(k => !(k in map)).sort().forEach((k, i) => {
    const p = pool.length ? pool : TABLEAU_10;
    auto[k] = p[i % p.length];
  });
  tfLog(`Categorical colours for ${enc.ref.inner}: ${Object.keys(map).length} from TWB, ${Object.keys(auto).length} auto-assigned`);
  return v => (v == null || v === "") ? null : (map[tfNorm(v)] || auto[tfNorm(v)] || null);
}


/* ═══ format/date-format.js ═══════════════════════════════════════════════════════════════════════ */
/* Tableau date label formats (Format → Header → Dates) applied to header values. */

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

/**
 * A header value in Tableau's custom date format: "iLLLLL" → "J" (month initial), "MMM" → "Jan",
 * "MMM yy" → "Jan 24", "yyyy" → "2024" … The value can be a full date or a date part Tableau sends as
 * text (month / weekday name, "Q1", year). null when the format needs a part the value does not have.
 * @param {string} raw Tableau text-format ("i…" / "*…" prefix + ICU-style pattern)
 * @param {string} text the formatted value from the summary data
 * @param {DataValue} [dv]
 * @returns {string | null}
 */
function tfFormatDateLabel(raw, text, dv) {
  const custom = /^\*/.test(String(raw || ""));          // "*mmm'yy": Tableau's own (Excel-style) tokens, not ICU
  const pattern = String(raw || "").replace(/^[i*]/, "");
  if (custom ? !/[ymdq]/i.test(pattern) : !/[yMLdEQq]/.test(pattern)) return null;
  const t = String(text || "").trim().toLowerCase();
  let date = null, month = null, weekday = null, quarter = null, year = null;
  const native = dv ? (dv.nativeValue !== undefined ? dv.nativeValue : dv.value) : null;
  if (native instanceof Date) date = native;
  else if (typeof native === "string" && /^\d{4}-\d{2}-\d{2}/.test(native)) date = new Date(native.slice(0, 10) + "T00:00:00Z");
  if (date && !isNaN(date.getTime())) {
    year = date.getUTCFullYear(); month = date.getUTCMonth(); weekday = date.getUTCDay(); quarter = Math.floor(month / 3) + 1;
  } else {
    date = null;
    const m = MONTHS.findIndex(n => n.toLowerCase() === t || n.slice(0, 3).toLowerCase() === t);
    if (m >= 0) { month = m; quarter = Math.floor(m / 3) + 1; }
    const d = DAYS.findIndex(n => n.toLowerCase() === t || n.slice(0, 3).toLowerCase() === t);
    if (d >= 0) weekday = d;
    const q = t.match(/^q([1-4])$/);
    if (q) quarter = +q[1];
    if (/^\d{4}$/.test(t)) year = +t;
  }
  if (month === null && weekday === null && quarter === null && year === null) return null;
  let missing = false;
  if (custom) return tfFormatCustomDate(pattern, { date, month, weekday, quarter, year });
  const out = pattern.replace(/'([^']*)'|y{1,4}|M{1,5}|L{1,5}|d{1,2}|E{1,5}|Q{1,4}|q{1,4}/g, (tok, literal) => {
    if (literal !== undefined) return literal;
    const c = tok[0], n = tok.length;
    const need = v => { if (v === null) missing = true; return v === null; };
    if (c === "y") return need(year) ? "" : n === 2 ? String(year).slice(-2) : String(year);
    if (c === "M" || c === "L") {
      if (need(month)) return "";
      return n >= 5 ? MONTHS[month][0] : n === 4 ? MONTHS[month] : n === 3 ? MONTHS[month].slice(0, 3)
           : n === 2 ? String(month + 1).padStart(2, "0") : String(month + 1);
    }
    if (c === "d") return need(date) ? "" : String(date.getUTCDate()).padStart(n, "0");
    if (c === "E") return need(weekday) ? "" : n >= 5 ? DAYS[weekday][0] : n === 4 ? DAYS[weekday] : DAYS[weekday].slice(0, 3);
    return need(quarter) ? "" : n >= 3 ? "Q" + quarter : String(quarter);
  });
  return missing ? null : out;
}


/**
 * Tableau's custom date format (the "*" text-formats, Excel-style): yyyy / yy year, m … mmmmm month
 * (1, 01, Jan, January, J), d / dd day, ddd / dddd weekday, q quarter; "…" and \x are literal text, every
 * other character (' / - . space …) is printed as it is. Time parts (h, n, s, AM/PM) and weeks are not
 * rebuilt here – null, so the caller keeps Tableau's own text.
 * @param {string} pattern @param {{ date: Date | null, month: number | null, weekday: number | null,
 *   quarter: number | null, year: number | null }} p @returns {string | null}
 */
function tfFormatCustomDate(pattern, p) {
  let missing = false;
  const need = v => { if (v === null) missing = true; return v === null; };
  const out = pattern.replace(/"([^"]*)"|\\(.)|am\/pm|ampm|a\/p|y{3,4}|y{1,2}|m{1,5}|d{1,4}|q|h{1,2}|n{1,2}|s{1,2}|w{1,2}/gi, (tok, quoted, escaped) => {
    if (quoted !== undefined) return quoted;
    if (escaped !== undefined) return escaped;
    const c = tok[0].toLowerCase(), n = tok.length;
    if (c === "y") return need(p.year) ? "" : n <= 2 ? String(p.year).slice(-2) : String(p.year);
    if (c === "m") {
      if (need(p.month)) return "";
      return n >= 5 ? MONTHS[p.month][0] : n === 4 ? MONTHS[p.month] : n === 3 ? MONTHS[p.month].slice(0, 3)
           : n === 2 ? String(p.month + 1).padStart(2, "0") : String(p.month + 1);
    }
    if (c === "d") {
      if (n >= 3) return need(p.weekday) ? "" : n === 4 ? DAYS[p.weekday] : DAYS[p.weekday].slice(0, 3);
      return need(p.date) ? "" : String(/** @type {Date} */ (p.date).getUTCDate()).padStart(n, "0");
    }
    if (c === "q") return need(p.quarter) ? "" : String(p.quarter);
    missing = true;                                     // time of day, week number: Tableau's own text is used
    return "";
  });
  return missing ? null : out;
}


/* ═══ twb/dashboard-text.js ═══════════════════════════════════════════════════════════════════════ */
/* Dashboard text boxes: header/title runs and zone borders. */

/* Text boxes count as column headers only if ALL hold:
 *   – ≥ 2 text zones, visible, in one row (same top edge ± tol)
 *   – their bottom edge touches the sheet's top edge (± tol)
 *   – they lie within the sheet's horizontal span
 *   – their number equals the number of visible columns
 * Positions are Tableau's own zone coordinates (0–100000), so the check is deterministic. */
const TF_ZONE_TOL = 1500;

function tfZoneText(z) { return (z.runs || []).map(r => r.text).join("").replace(/\u00C6\r?\n?/g, "\n").trim(); }

function tfTextBoxHeaders(dash, sheetName, nCols) {
  const strip = tfTextBoxHeaderStrip(dash, sheetName);
  if (!strip || nCols < 2 || strip.zones.length !== nCols) return null;
  return strip;
}

/* the row of text boxes sitting on a worksheet's top edge, left to right (≥ 2), whatever their number */
function tfTextBoxHeaderStrip(dash, sheetName) {
  const ws = dash.zones.find(z => z.name === sheetName && z.type === "worksheet");
  if (!ws) return null;
  const near = (a, b) => Math.abs(a - b) <= TF_ZONE_TOL;
  const cand = dash.zones.filter(z => z.type === "text" && !z.hidden && z.runs && tfZoneText(z) &&
    near(z.y + z.h, ws.y) && z.x >= ws.x - TF_ZONE_TOL && z.x + z.w <= ws.x + ws.w + TF_ZONE_TOL);
  if (cand.length < 2 || !cand.every(z => near(z.y, cand[0].y))) return null;
  cand.sort((a, b) => a.x - b.x);
  return { ws, zones: cand.map(z => ({ id: z.id, text: tfZoneText(z), x: z.x, w: z.w, h: z.h, runs: z.runs,
    props: (z.runs.find(r => r.text.trim()) || { props: {} }).props })) };
}

/* Dashboard title: Tableau's own title if shown, else the text box above every worksheet
 * with the largest font (ties → topmost). Returns runs for rich text, or null. */
function tfDashboardTitleRuns(model, dashName, usedZoneIds) {
  const dash = model && model.dashboards && model.dashboards[dashName];
  if (!dash) return null;
  const titleZone = dash.zones.find(z => z.type === "title" && !z.hidden);
  if (titleZone && dash.title && dash.title.runs.length) return dash.title.runs;
  if (!FORMAT_CONFIG.textBoxTitle) return null;
  const wsTop = Math.min(...dash.zones.filter(z => z.type === "worksheet").map(z => z.y));
  const size = z => (z.runs.find(r => r.text.trim()) || { props: {} }).props.fontSize || 9;
  const cand = dash.zones.filter(z => z.type === "text" && !z.hidden && z.runs && tfZoneText(z) &&
    !usedZoneIds.has(z.id) && z.y + z.h <= wsTop + TF_ZONE_TOL);
  if (!cand.length) return null;
  cand.sort((a, b) => size(b) - size(a) || a.y - b.y);
  return cand[0].runs;
}

function tfStrokeToBorder(size) {
  if (!size) return null;
  if (size <= 1) return "thin";
  if (size === 2) return "medium";
  return "thick";
}


/* ═══ twb/formatter.js ════════════════════════════════════════════════════════════════════════════ */
/* Resolves Tableau's formatting cascade for one worksheet (createSheetFormatter). */

function tfFormatsToProps(m) {
  const p = {};
  if (m["font-family"]) p.fontName = m["font-family"];
  if (m["font-size"]) p.fontSize = tfNum(m["font-size"]);
  if (m["font-weight"]) p.bold = /bold/i.test(m["font-weight"]);
  if (m["font-style"]) p.italic = /italic/i.test(m["font-style"]);
  if (m["text-decoration"]) { p.underline = /underline/i.test(m["text-decoration"]); p.strike = /line-through/i.test(m["text-decoration"]); }
  if (m["color"]) p.color = tfArgb(m["color"]) || undefined;
  if ("background-color" in m) p.bgColor = tfArgb(m["background-color"]);  // null = transparent
  if (/^(left|center|right)$/.test(m["text-align"] || "")) p.hAlign = m["text-align"];
  const va = { top: "top", center: "middle", middle: "middle", bottom: "bottom" }[m["vertical-align"]];
  if (va) p.vAlign = va;
  if (m["wrap"]) p.wrap = /^(true|on)$/i.test(m["wrap"]);
  if (m["text-format"]) p.numFmtRaw = m["text-format"];
  if (m["band-color"]) p.bandColor = tfArgb(m["band-color"]) || undefined;
  if (m["band-size"]) p.bandSize = tfNum(m["band-size"]);
  if (m["line-visibility"]) p.lineVisible = m["line-visibility"] !== "off";
  if (m["stroke-size"] != null) p.strokeSize = tfNum(m["stroke-size"]);
  if (m["stroke-color"]) p.strokeColor = tfArgb(m["stroke-color"]) || undefined;
  if (m["div-level"]) p.divLevel = tfNum(m["div-level"]);
  if (m["width"]) p.width = tfNum(m["width"]);
  if (m["display"]) p.display = m["display"] !== "false";
  if (m["display-field-labels"]) p.displayFieldLabels = m["display-field-labels"] !== "false";
  if (m["height"]) p.height = tfNum(m["height"]);
  if (m["text-orientation"] != null) p.textOrientation = tfNum(m["text-orientation"]);
  return p;
}

/* precedence inside one element: base < scope < field < field+scope */
function tfCollect(style, elements, opts = {}) {
  const out = {};
  if (!style) return out;
  for (const el of elements) {
    const tiers = [{}, {}, {}, {}];
    for (const f of style.rules[el] || []) {
      if ((f.dataClass || null) !== (opts.dataClass || null)) continue;
      if (f.scope && f.scope !== opts.scope) continue;
      if (f.field && !(opts.field && tfSameField(tfParseFieldRef(f.field), opts.field))) continue;
      tiers[(f.field ? 2 : 0) + (f.scope ? 1 : 0)][f.attr] = f.value;
    }
    Object.assign(out, ...tiers);
  }
  return tfFormatsToProps(out);
}

function tfMerge(...layers) {
  const out = {};
  layers.forEach(l => { if (l) for (const k in l) if (l[k] !== undefined) out[k] = l[k]; });
  return out;
}

function tfFieldInfo(model, ref) {
  if (!model || !ref) return null;
  const k = (tfTableCalcBase(ref) || ref).name.toLowerCase();      // a table calculation: its measure's field
  return model.fields[(ref.ds || "") + "|" + k] || model.fields["|" + k] || null;
}

function tfDisplayNames(model, ref) {
  const info = tfFieldInfo(model, ref);
  const cap = (info && info.caption) || ref.name;
  const agg = TF_DERIV_LABEL[(ref.deriv || "").toLowerCase()];
  const names = [];
  if (agg) names.push(`${agg}(${cap})`, `${agg}(${ref.name})`);
  names.push(cap, ref.name);
  if (ref.name === "Multiple Values") names.push("Measure Values");
  const alias = model && model.measureAliases && (model.measureAliases[tfRefKey(ref)] ||
    Object.entries(model.measureAliases).find(([k]) => k.endsWith("|" + ref.inner.toLowerCase()))?.[1]);
  if (alias) names.unshift(alias);
  return [...new Set(names.map(tfNorm))];
}

/** Format → Dashboard → Dashboard Shading: the dashboard's canvas colour (ARGB), or null
 * @param {FormatModel | null} model @param {string} dashboardName @returns {string | null} */
function tfDashboardShading(model, dashboardName) {
  const dash = model && model.dashboards ? model.dashboards[dashboardName] : null;
  return (dash && tfCollect(dash.style, ["table"]).bgColor) || null;
}

/** @param {FormatModel | null} model @param {string} sheetName */
/**
 * @param {FormatModel | null} model @param {string} sheetName
 * @param {Pane[]} [onlyPanes] scope the Marks-card methods to these panes (one layer of a dual-axis donut)
 */
function createSheetFormatter(model, sheetName, onlyPanes) {
  const sheet = (model && model.sheets && model.sheets[sheetName]) || null;
  const wb = model ? model.workbookStyle : null;
  const st = sheet ? sheet.style : null;
  const panes = onlyPanes || (sheet ? sheet.panes : []);
  /** @param {Pane} p @param {string} attr */
  const markRule = (p, attr) => (((p.style && p.style.rules.mark) || []).find(x => x.attr === attr) || {}).value;
  const matchCache = new Map();

  function base() {
    return tfMerge(TABLEAU_DEFAULTS.worksheet, tfCollect(wb, TF_ELEMENTS.sheet), tfCollect(st, TF_ELEMENTS.sheet));
  }

  /* summary-data column name ("SUM(Sales)") → Tableau field ref. Exact names only,
   * no substring matching (old code matched "Sales" to "Sales Person"). */
  function matchName(colName) {
    const key = tfNorm(colName);
    if (matchCache.has(key)) return matchCache.get(key);
    let best = null, bestScore = 0;
    const stripped = key.replace(/^[a-z]+\((.*)\)$/, "$1");
    for (const ref of sheet ? sheet.fieldRefs : []) {
      const names = tfDisplayNames(model, ref);
      let score = 0;
      if (names[0] === key) score = 3;
      else if (names.includes(key)) score = 2;
      else if (names.includes(stripped)) score = 1;
      if (score > bestScore) { best = ref; bestScore = score; }
    }
    matchCache.set(key, best);
    return best;
  }

  /* summary-data column id ("[federated.x].[pcto:sum:Sales:qk:2]") → the sheet's field ref; null when the id is not
   * a field instance of this sheet */
  function matchFieldId(id) {
    const r = id && /^\[/.test(String(id)) ? tfParseFieldRef(id) : null;
    if (!r || !r.type || !sheet) return null;
    return sheet.fieldRefs.find(x => x.inner.toLowerCase() === r.inner.toLowerCase()) || null;
  }

  function onShelf(ref, shelf) { return !!ref && (sheet ? sheet[shelf] : []).some(r => tfSameField(r, ref)); }
  // headers = DISCRETE pills on Rows/Columns; continuous pills (…:qk) draw axes, not headers
  function isHeaderField(ref) { return !!ref && ref.type !== "qk" && (onShelf(ref, "rows") || onShelf(ref, "cols")); }
  function shelfOf(ref) { return onShelf(ref, "rows") ? "rows" : onShelf(ref, "cols") ? "cols" : null; }

  function labelRunProps(ref) {
    if (!ref) return {};
    for (const p of panes) for (const r of p.labelRuns) if (r.refs.some(x => tfSameField(x, ref))) return r.props;
    return {};
  }
  /* panes that draw this field (multi-pane tables: each column has its own marks card) */
  function owningPanes(ref) {
    const owns = p => ref && (p.labelRuns.some(r => r.refs.some(x => tfSameField(x, ref))) ||
                              p.encodings.some(e => tfSameField(e.field, ref)));
    const hit = panes.filter(owns);
    if (hit.length) return hit;
    if (panes.length === 1) return panes;
    return panes.filter(p => !p.id);                            // the "All" pane
  }
  function paneStyle(ref) {
    const els = TF_ELEMENTS.pane.concat(TF_ELEMENTS.cell, TF_ELEMENTS.datalabel);
    return tfMerge(...owningPanes(ref).map(p => tfCollect(p.style, els, { field: ref })));
  }

  function numFmt(ref, fromHeader) {
    const els = fromHeader ? TF_ELEMENTS.header : TF_ELEMENTS.pane.concat(TF_ELEMENTS.cell);
    const raw = tfCollect(st, els, { field: ref, scope: fromHeader ? "rows" : undefined }).numFmtRaw
             || (ref && (tfFieldInfo(model, ref) || {}).defaultFormat);
    return raw ? tableauToExcelNumFmt(raw) : null;
  }

  return {
    hasModel: !!sheet,
    matchName,
    matchFieldId,
    isHeaderField,

    titleText() { return sheet && sheet.title ? sheet.title.text : null; },
    titleStyle() {
      const runs = sheet && sheet.title ? sheet.title.runs : [];
      const firstRun = (runs.find(r => r.text.trim()) || {}).props;
      return tfMerge(base(), TABLEAU_DEFAULTS.title, tfCollect(wb, TF_ELEMENTS.title), tfCollect(st, TF_ELEMENTS.title), firstRun);
    },

    /* Excel header row (field names) */
    fieldLabelStyle(ref, isMeasure) {
      return tfMerge(base(), TABLEAU_DEFAULTS.fieldLabel,
        tfCollect(wb, TF_ELEMENTS.fieldLabel, { scope: "rows" }),
        tfCollect(st, TF_ELEMENTS.fieldLabel, { field: ref, scope: isMeasure ? "cols" : "rows" }));
    },

    /* dimension values (row headers in Tableau) */
    headerCellStyle(ref) {
      const p = tfMerge(base(), tfCollect(wb, TF_ELEMENTS.header, { scope: "rows" }),
        tfCollect(st, TF_ELEMENTS.header, { field: ref, scope: "rows" }));
      p.numFmt = numFmt(ref, true);
      return p;
    },

    /** a header's text orientation in degrees (-90 = rotated up), null when Tableau's default (horizontal)
     * @param {FieldRef} ref */
    headerOrientation(ref) {
      if (!ref) return null;
      const p = tfMerge(tfCollect(wb, ["label"], { field: ref }), tfCollect(st, ["label"], { field: ref }),
        tfCollect(st, ["label"], { field: ref, scope: "cols" }), tfCollect(st, ["label"], { field: ref, scope: "rows" }));
      return typeof p.textOrientation === "number" ? p.textOrientation : null;
    },

    /* measure values (text marks in Tableau) */
    markCellStyle(ref) {
      const els = TF_ELEMENTS.pane.concat(TF_ELEMENTS.cell, TF_ELEMENTS.datalabel);
      const p = tfMerge(base(), tfCollect(wb, els),
        tfCollect(st, els, { field: ref, scope: "rows" }),
        paneStyle(ref), labelRunProps(ref));
      // cell position comes from Format → Alignment (pane/cell); the label editor's alignment only
      // aligns lines inside the label, so it is used only when no pane/cell alignment is set
      const cellAlign = tfMerge(tfCollect(st, els, { field: ref, scope: "rows" }), paneStyle(ref)).hAlign;
      if (cellAlign) p.hAlign = cellAlign;
      p.numFmt = numFmt(ref, false);
      p.explicitColor = !!(labelRunProps(ref).color || tfCollect(st, TF_ELEMENTS.cell, { field: ref }).color);
      return p;
    },

    divider(scope) {
      const p = tfCollect(st, TF_ELEMENTS.divider, { scope });
      const d = scope === "rows" ? { ...TABLEAU_DEFAULTS.divider } : { ...TABLEAU_DEFAULTS.divider, visible: false };
      if (p.lineVisible !== undefined) d.visible = p.lineVisible;
      if (p.strokeSize !== undefined) { d.style = tfStrokeToBorder(p.strokeSize); if (!d.style) d.visible = false; }
      if (p.strokeColor) d.color = p.strokeColor;
      if (p.divLevel) d.level = p.divLevel;
      return d;
    },

    banding() {
      const pane = tfCollect(st, TF_ELEMENTS.pane, { scope: "rows" });
      const header = tfCollect(st, TF_ELEMENTS.header, { scope: "rows" });
      const table = tfCollect(st, ["table"], { scope: "rows" });
      return { pane: pane.bandColor || null, header: header.bandColor || null,
               size: table.bandSize || pane.bandSize || header.bandSize || 1 };
    },

    widthPx(ref) {
      return tfCollect(st, TF_ELEMENTS.cell.concat(TF_ELEMENTS.header), { field: ref, scope: "cols" }).width
          || tfCollect(st, TF_ELEMENTS.cell.concat(TF_ELEMENTS.header), { field: ref }).width || null;
    },
    /* row height: Format → Cell height, which Tableau stores on the innermost Rows dimension when set by
       dragging the row border */
    rowHeightPx() {
      const own = tfCollect(st, TF_ELEMENTS.cell).height;
      if (own) return own;
      const dims = sheet ? sheet.rows.filter(r => r.type !== "qk") : [];
      for (let i = dims.length - 1; i >= 0; i--) {
        const h = tfCollect(st, TF_ELEMENTS.cell, { field: dims[i] }).height;
        if (h) return h;
      }
      return null;
    },

    /* fields on Text / Label – these are the "marks" that get coloured */
    textRefs() {
      const refs = [];
      panes.forEach(p => {
        p.encodings.filter(e => e.channel === "text" || e.channel === "label").forEach(e => refs.push(e.field));
        p.labelRuns.forEach(r => r.refs.forEach(x => refs.push(x)));
      });
      return refs;
    },

    colorEncoding() {
      for (const p of panes) {
        const out = this.colorEncodingOf(p);
        if (out) return out;
      }
      return null;
    },

    /** a Shape encoding's value → shape map ("Arrows/1-4.png", ":filled/circle"): sheet, then data source
     * @param {FieldRef} ref @returns {Record<string, string> | null} */
    shapeMap(ref) {
      const pool = [st, ...(model ? Object.values(model.datasourceStyles) : []), wb].filter(Boolean);
      for (const s of pool) {
        const def = (s.shapes || []).find(x => tfSameField(x.field, ref)) ||
                    (s.shapes || []).find(x => x.field && tfNorm(x.field.name) === tfNorm(ref.name));
        if (def) return def.map;
      }
      return null;
    },

    /** one pane's Color encoding (a sheet built from marks has one per pane) @param {any} p a pane */
    colorEncodingOf(p) {
      const enc = p.encodings.find(e => e.channel === "color" && e.field);
      if (!enc) return null;
      const ref = enc.field;
      const hasText = p.encodings.some(e => e.channel === "text" || e.channel === "label") || p.labelRuns.length > 0;
      const cls = String(p.markClass || "Automatic");
      const effective = /^automatic$/i.test(cls) && hasText ? "Text" : cls;
      const pool = [st, ...panes.map(x => x.style), model && model.datasourceStyles[ref.ds],
                    ...(model ? Object.values(model.datasourceStyles) : [])].filter(Boolean);
      let def = null;
      for (const s of pool) { def = s.encodings.find(e => tfSameField(e.field, ref)); if (def) break; }
      if (!def) for (const s of pool) { def = s.encodings.find(e => e.field && tfNorm(e.field.name) === tfNorm(ref.name)); if (def) break; }
      const paneRefs = [...p.encodings.filter(e => e.channel === "text" || e.channel === "label").map(e => e.field),
                        ...p.labelRuns.flatMap(r => r.refs)];
      /** a measure's own colour rule – Color on Measure Values gives every measure its own legend @param {FieldRef} r */
      const defFor = r => {
        for (const s of pool) { const d = s.encodings.find(e => tfSameField(e.field, r)); if (d) return d; }
        return null;
      };
      return { ref, def, defFor, markClass: effective, applyTo: /^text$/i.test(effective) ? "font" : "fill", paneRefs,
               continuous: ref.type ? /^q/.test(ref.type) : !!(def && def.type === "interpolated") };
    },

    shelfOf,
    sheetModel: sheet,

    /** Format → Shading → Worksheet, workbook < worksheet: ARGB; "none" = switched off (what is behind the
     * sheet shows); undefined = not set (Tableau draws the sheet white). A dashboard's shading colours the
     * dashboard only – sheets on it keep their own (designers set them to match).
     * @returns {string | undefined} */
    sheetShading() {
      let out;
      [wb, st].forEach(s => { const x = s ? tfCollect(s, ["table"]).bgColor : undefined; if (x !== undefined) out = x === null ? "none" : x; });
      return out;
    },
    /** the worksheet's own background colour (ARGB), null when it has none (white or switched off) */
    tableBackground() {
      const s = this.sheetShading();
      return s && s !== "none" ? s : null;
    },
    /* column-header height (px) stored on any Columns-shelf header field, incl. Measure Names */
    headerRowHeightPx() {
      if (!sheet) return null;
      let h = null;
      (st && st.rules.header || []).forEach(f => {
        if (f.attr !== "height" || !f.field) return;
        const r = tfParseFieldRef(f.field);
        if (r.name === "Measure Names" || onShelf(r, "cols")) h = Math.max(h || 0, tfNum(f.value) || 0);
      });
      return h;
    },
    /* URL actions whose source includes this sheet */
    urlActions(dashboardName) {
      return (model && model.actions || []).filter(a =>
        (a.worksheet ? a.worksheet === sheetName : (!a.dashboard || a.dashboard === dashboardName)) &&
        !(a.exclude || []).includes(sheetName));
    },

    /* header label hidden via Format → "Hide" (label display=false) */
    isLabelHidden(ref) { return tfCollect(st, ["label"], { field: ref }).display === false; },

    /* an axis on a shelf: "Show Header" off (display=false) and Edit Axis → title ("" = no title).
     * cls "0" / "1" = primary / secondary axis of a dual axis; rules without a class apply to both */
    axisInfo(ref, shelf, cls) {
      /** @type {{ hidden?: boolean, title?: string }} */
      const out = {};
      if (!st || !ref) return out;
      /** @type {Record<string, string>[]} */
      const tiers = [{}, {}];                                 // without scope < with scope
      for (const f of st.rules.axis || []) {
        if (f.attr !== "display" && f.attr !== "title") continue;
        if (!f.field || !tfSameField(tfParseFieldRef(f.field), ref)) continue;
        if (f.scope && f.scope !== shelf) continue;
        if (f.axisClass !== undefined && cls !== undefined && f.axisClass !== cls) continue;
        tiers[f.scope ? 1 : 0][f.attr] = f.value;
      }
      const v = { ...tiers[0], ...tiers[1] };
      if (v.display !== undefined) out.hidden = v.display === "false";
      if (v.title !== undefined) out.title = v.title;
      return out;
    },
    /* grid lines across a shelf's axis (Format → Lines → Grid Lines): off = stroke 0 or line-visibility off */
    gridlinesShown(shelf) {
      const g = tfMerge(tfCollect(wb, ["gridline"], { scope: shelf }), tfCollect(st, ["gridline"], { scope: shelf }));
      return !(g.strokeSize === 0 || g.lineVisible === false);
    },
    /* Analytics → Reference Line entries of the worksheet */
    referenceLines() { return (sheet && sheet.referenceLines) || []; },
    /* Edit Axis for a field on a shelf: fixed range, tick spacing, include zero */
    axisSpace(ref, shelf, cls) {
      /** @type {AxisSpace} */
      let out = {};
      if (!st || !st.spaces || !ref) return out;
      for (const s of st.spaces) {
        if (!s.field || !tfSameField(tfParseFieldRef(s.field), ref) || (s.scope && s.scope !== shelf)) continue;
        if (s.axisClass !== undefined && cls !== undefined && s.axisClass !== cls) continue;
        out = s;
      }
      return out;
    },
    /* a header's label format (Format → Header → Dates), e.g. "iLLLLL" = month initial */
    labelFormat(ref) { return ref ? tfCollect(st, ["label"], { field: ref }).numFmtRaw || null : null; },
    /** mark label font (datalabel element); its colour only when the user picked one, not "automatic"
     * @returns {Record<string, any>} */
    dataLabelStyle() {
      const raw = {};
      const pool = [wb, st, ...owningPanes(null).map(p => p.style)];      // workbook, worksheet, marks card
      for (const s of pool) for (const f of (s && s.rules.datalabel) || []) if (!f.field) raw[f.attr] = f.value;
      const p = tfMerge(...pool.map(s => tfCollect(s, ["datalabel"])));
      if (raw["color-mode"] && raw["color-mode"] !== "user") delete p.color;
      return p;
    },
    /** the mark label's text around its one field ("<AGG(Days)> DAYS" → prefix "", suffix " DAYS"); null when
     * the label is the default, or holds several fields
     * @returns {{ ref: FieldRef, prefix: string, suffix: string } | null} */
    labelTemplate() {
      const pane = owningPanes(null)[0];
      if (!pane || !pane.labelRuns.length) return null;
      const text = pane.labelRuns.map(r => r.text).join("");
      const tokens = [...text.matchAll(/<([^<>]+)>/g)];
      const ref = tokens.length === 1 ? tfParseFieldRef(tokens[0][1]) : null;
      if (!ref) return null;
      const t = tokens[0], flat = s => s.replace(/\u00C6[ \t]*(?:\r?\n)?|\r?\n/g, " ");
      return { ref, prefix: flat(text.slice(0, t.index)), suffix: flat(text.slice(t.index + t[0].length)) };
    },
    /** reference line formatting: line colour (and its opacity), width, dash, visibility; label font / format
     * @returns {Record<string, any>} */
    reflineStyle() {
      const raw = {};
      const pool = [wb, st, ...owningPanes(null).map(p => p.style)];
      for (const s of pool) for (const f of (s && s.rules.refline) || []) if (!f.field && !f.scope) raw[f.attr] = f.value;
      const alpha = /^#?[0-9a-f]{8}$/i.test(raw["stroke-color"] || "") ? parseInt(raw["stroke-color"].replace(/^#/, "").slice(6), 16) / 255 : 1;
      return { ...tfMerge(...pool.map(s => tfCollect(s, ["refline"]))), dash: raw["line-pattern-only"] || null, strokeAlpha: alpha };
    },

    /* axis rulers (Format → Lines → Axis Rulers) */
    axisLineShown(shelf) {
      const a = tfMerge(tfCollect(wb, ["axis"], { scope: shelf }), tfCollect(st, ["axis"], { scope: shelf }));
      return !(a.strokeSize === 0 || a.lineVisible === false);
    },
    /* "Show field labels for rows/columns" (worksheet display-field-labels) */
    fieldLabelsShown(scope) {
      const v = tfCollect(st, ["worksheet"], { scope }).displayFieldLabels;
      return v === undefined ? true : v;
    },
    /* Rows then Columns, discrete pills only, in shelf order */
    headerRefs() {
      return sheet ? [...sheet.rows, ...sheet.cols].filter(r => r.type !== "qk" && r.name !== "Measure Names") : [];
    },
    /* measure order: Measure Names manual sort, else the Measure Names filter order */
    measureOrder() {
      if (!sheet) return [];
      const ms = sheet.manualSorts.find(m => m.field && m.field.name === "Measure Names");
      const src = ms ? ms.order : sheet.measureFilter;
      return src.map(x => tfParseFieldRef(String(x).replace(/^"|"$/g, ""))).filter(Boolean);
    },
    /* Measure Names text ("Typical home value") → measure field ref */
    measureRefForName(text) {
      const key = tfNorm(text);
      const cands = [...this.measureOrder(), ...(sheet ? sheet.fieldRefs : [])];
      return cands.find(r => tfDisplayNames(model, r).includes(key)) || matchName(text);
    },
    /* panes in visual order with the fields they label */
    panesInOrder() {
      // a pane drawn on its own axis sits where that axis is on Columns / Rows (one pane per measure)
      const shelfPos = p => {
        const x = p.xAxisName ? sheet.cols.findIndex(r => tfSameField(r, tfParseFieldRef(p.xAxisName))) : -1;
        if (x >= 0) return x;
        return p.yAxisName ? sheet.rows.findIndex(r => tfSameField(r, tfParseFieldRef(p.yAxisName))) : -1;
      };
      return panes.map((p, i) => ({
        pane: p, i,
        order: p.xIndex !== undefined && p.xIndex !== null ? p.xIndex
          : shelfPos(p) >= 0 ? shelfPos(p) : (p.xAxisName ? 0 : (panes.length > 1 ? -1 : 0)),
        refs: [...p.labelRuns.flatMap(r => r.refs),
               ...p.encodings.filter(e => e.channel === "text" || e.channel === "label").map(e => e.field)]
      })).sort((a, b) => a.order - b.order || a.i - b.i);
    },
    /* ── pies and donuts ── */
    /** pie layers { outer, hole, inner } on a dual axis (MIN(0) twice); the later axis draws on top.
     *   single pie                                       → { outer }
     *   donut: + a Pie / Circle layer with nothing on its card splitting it into slices → { outer, hole }
     *   nested donut: two coloured pies, the top one smaller                           → { outer, inner }
     * A layer underneath that the top layer covers is dropped (plain pie). The id-less "All" pane draws nothing.
     * null: not a pie (or 3+ layers). */
    pieLayers() {
      // the drawn layers: on a dual axis the panes with their own axis (the "All" marks card, pane 1 in some
      // workbooks, has an id but no axis); on map layers the panes with an id
      const axisPanes = panes.filter(p => p.yAxisName || p.xAxisName);
      // on map layers the "All" pane has an id too, but no field of its own
      const layerPanes = panes.filter(p => p.id && (p.encodings || []).length);
      const drawn = axisPanes.length >= 2 ? axisPanes : panes.length > 1 && layerPanes.length ? layerPanes : panes;
      const isPieP = p => /^pie$/i.test(p.markClass || "");
      if (drawn.length === 1) return isPieP(drawn[0]) ? { outer: drawn[0], hole: null, inner: null } : null;
      if (drawn.length !== 2) return null;
      // a gauge's hole: a Circle layer never splits into wedges, whatever is on its Color (often Measure Names)
      const wedgeP = p => isPieP(p) && p.encodings.some(e => e.channel === "wedge-size");
      const wi = drawn.findIndex(wedgeP);
      if (wi >= 0 && /^circle$/i.test(drawn[1 - wi].markClass || "") && !drawn[1 - wi].encodings.some(e => e.channel === "wedge-size"))
        return { outer: drawn[wi], hole: 1 - wi > wi ? drawn[1 - wi] : null, inner: null };
      const splits = p => p.encodings.some(e => e.channel === "wedge-size" || (/^(color|lod)$/.test(e.channel) && e.field.type !== "qk"));
      const oi = drawn.findIndex(splits), hi = 1 - oi;
      if (oi < 0) return null;
      if (splits(drawn[hi])) {                                   // nested donut
        if (!drawn.every(isPieP)) return null;
        const size = p => parseFloat(markRule(p, "size")) || 1;
        const [bottom, top] = drawn;
        // a gauge's hole drawn as a second pie: smaller, on top, Marks colour white – the hole with its text
        const white = p => /^#?f{6}$/i.test(String(markRule(p, "mark-color") || "").trim());
        if (size(top) < size(bottom) && white(top) && !white(bottom)) return { outer: bottom, hole: top, inner: null };
        return size(top) < size(bottom) ? { outer: bottom, hole: null, inner: top } : { outer: top, hole: null, inner: null };
      }
      if (!isPieP(drawn[oi]) || !/^(pie|circle)$/i.test(drawn[hi].markClass || "")) return null;
      return { outer: drawn[oi], hole: hi > oi ? drawn[hi] : null, inner: null };
    },
    /** Marks → Color for a pane without a colour field, RRGGBB; Tableau blue if unset */
    markColor() {
      const v = panes.map(p => markRule(p, "mark-color")).find(Boolean);
      return ((v && tfArgb(v)) || TABLEAU_10[0]).slice(2);
    },
    /** the axis measure of each pane, e.g. one layer's own MIN(0) of a dual-axis donut */
    axisRefs() { return panes.map(p => p.yAxisName || p.xAxisName).filter(Boolean).map(tfParseFieldRef); },
    /** fields on the Marks card by channel ("color", "wedge-size", "text" …) @param {RegExp} channelRe */
    encodingRefs(channelRe) { return panes.flatMap(p => p.encodings.filter(e => channelRe.test(e.channel)).map(e => e.field)); },
    baseStyle: base,
    /** Marks → Size slider value; null = never moved */
    markSize() {
      for (const p of panes) {
        const v = parseFloat(markRule(p, "size"));
        if (isFinite(v)) return v;
      }
      return null;
    },
    /** Label → "Allow labels to overlap other marks" off (Tableau's default) → overlapping labels hidden */
    markLabelsCulled() { return panes.some(p => markRule(p, "mark-labels-cull") !== "false"); },
    /** Label → Marks to Label: "range" (Min/Max of a field), "line-ends", "most-recent" …; null = all marks */
    markLabelsMarks() {
      const p = panes.find(x => markRule(x, "mark-labels-mode") !== undefined);
      if (!p) return null;
      const r = a => markRule(p, a);
      return { mode: r("mark-labels-mode"), field: r("mark-labels-range-field"),
               min: r("mark-labels-range-min") !== "false", max: r("mark-labels-range-max") !== "false",
               first: r("mark-labels-line-first") !== "false", last: r("mark-labels-line-last") !== "false" };
    },
    /** Label → "Show mark labels"; unset = on when something is on Label */
    markLabelsShown() {
      return panes.some(p => {
        const f = markRule(p, "mark-labels-show");
        if (f !== undefined) return f === "true";
        return p.labelRuns.length > 0 || p.encodings.some(e => e.channel === "text" || e.channel === "label");
      });
    },
    /** label text as runs [{text, props}]: the Label editor, else Tableau's default (each field on Label on its
     * own line); text holds <[ds].[field]> placeholders */
    labelRunsTemplate() {
      const p = panes[0];
      if (!p) return [];
      if (p.labelRuns.length) return p.labelRuns.map(r => ({ text: r.text, props: r.props }));
      return p.encodings.filter(e => e.channel === "text" || e.channel === "label")
        .map((e, i) => ({ text: (i ? "\n" : "") + "<" + e.field.raw + ">", props: {} }));
    },
    hasCustomTooltip() { return !!(panes[0] && panes[0].tooltipRuns && panes[0].tooltipRuns.length); },
    /** tooltip text as runs: the Tooltip editor, else Tableau's default "Caption: <field>" per field in the view */
    tooltipTemplate() {
      const p = panes[0];
      if (p && p.tooltipRuns && p.tooltipRuns.length) return p.tooltipRuns.map(r => ({ text: r.text, props: r.props }));
      /** @type {FieldRef[]} */
      const refs = [];
      const add = r => { if (r && !refs.some(x => tfSameField(x, r))) refs.push(r); };
      // dual axis: a pane shows only its own axis measure, not the other layer's
      const axis = p && (p.yAxisName || p.xAxisName) ? tfParseFieldRef(p.yAxisName || p.xAxisName) : null;
      (sheet ? [...sheet.rows, ...sheet.cols] : []).filter(r => !axis || r.type !== "qk" || tfSameField(r, axis)).forEach(add);
      (p ? p.encodings : []).forEach(e => add(e.field));
      // dimensions first, then measures by instance name – as Tableau lists them on the dashboards seen so far
      refs.sort((a, b) => Number(a.type === "qk") - Number(b.type === "qk") || (a.type === "qk" ? a.inner.localeCompare(b.inner) : 0));
      // quick table calc "[pcto:sum:Qty:qk]" → "% of Total Qty along Table (Across)"
      const ALONG = { rows: " along Table (Across)", columns: " along Table (Down)" };
      const cap = r => /^pcto$/i.test(r.deriv || "")
        ? "% of Total " + this.captionFor({ ...r, deriv: r.name.split(":")[0], name: r.name.split(":").pop() }, r.name) +
          (ALONG[String(((sheet && sheet.tableCalcs) || {})[r.inner.toLowerCase()] || "").toLowerCase()] || "")
        : this.captionFor(r, r.name);
      return refs.map((r, i) => ({ text: (i ? "\n" : "") + cap(r) + ": <" + r.raw + ">", props: {} }));
    },
    /** font of a label run: worksheet font < Format → Label < the run's own font @param {Record<string, any>} runProps */
    labelFont(runProps) {
      return tfMerge(base(), tfCollect(wb, TF_ELEMENTS.datalabel), tfCollect(st, TF_ELEMENTS.datalabel),
        ...panes.map(p => tfCollect(p.style, TF_ELEMENTS.datalabel)), runProps);
    },
    /** Excel number format of a field's values @param {FieldRef} ref */
    numFmtFor: ref => numFmt(ref, false),
    /** the sheet's fit on a dashboard: "entire-view" | "fit-width" | "fit-height" | null (Standard) @param {string} dashboardName */
    fitMode(dashboardName) {
      const d = model && model.dashboards && model.dashboards[dashboardName];
      return (d && d.fit && d.fit[sheetName]) || null;
    },

    /* chart vs table: a visible continuous axis, or only chart marks with no text */
    isChart() {
      if (!sheet) return false;
      if (this.pieLayers()) return true;                         // pies usually carry labels
      for (const shelf of ["rows", "cols"]) {
        for (const r of sheet[shelf]) {
          const continuous = r.type === "qk" || r.name === "Multiple Values";
          if (!continuous) continue;
          const ax = tfCollect(st, ["axis"], { field: r, scope: shelf });
          if (ax.display !== false) return true;
        }
      }
      const chartMarks = /^(bar|line|area|pie|map|multipolygon|polygon|ganttbar|density)$/i;
      return panes.length > 0 && panes.every(p => chartMarks.test(p.markClass || "") &&
        !p.encodings.some(e => e.channel === "text" || e.channel === "label") && !p.labelRuns.length);
    },
    manualSortFor(ref) { return sheet ? sheet.manualSorts.find(m => tfSameField(m.field, ref)) : null; },
    naturalSortFor(ref) { return sheet ? (sheet.naturalSorts || []).find(m => tfSameField(m.field, ref)) || null : null; },
    measureSortFor(ref) { return sheet ? sheet.measureSorts.find(m => tfSameField(m.field, ref)) : null; },
    /** the sheet's quick table calculations over a measure ("cum:sum:Sales:qk:7" over "sum:Sales:qk") */
    tableCalcRefs(ref) {
      return ref && sheet ? sheet.fieldRefs.filter(r => { const b = tfTableCalcBase(r); return !!b && tfSameField(b, ref); }) : [];
    },
    /** a running total (Quick Table Calculation → Running Total) */
    isRunningTotal(ref) {
      return !!ref && (/^(cum|rsum)$/i.test(ref.deriv || "") || (sheet ? (sheet.runningTotals || []) : []).some(t => tfSameField(t, ref)));
    },
    captionFor(ref, fallbackName) {
      const base = tfTableCalcBase(ref);                       // a running total is captioned as its measure
      if (base) return this.captionFor(base, fallbackName);
      const info = tfFieldInfo(model, ref);
      let cap = info && info.caption ? info.caption
              : ref && !/^Calculation_/.test(ref.name) ? ref.name
              : String(fallbackName || "").replace(/^[A-Z]+\((.*)\)$/, "$1");
      const d = ref && (ref.deriv || "").toLowerCase();
      const pre = { avg: "Avg. ", cnt: "Count of ", ctd: "Count Distinct of ", min: "Min. ", max: "Max. ",
                    med: "Median ", std: "Std. dev. of ", var: "Variance of " }[d];
      return pre ? pre + cap : cap;
    },
    /* title text with <Sheet Name>, <[field]> and <[Parameters].[p]> substituted */
    renderTitle(valuesFor, sheetLabel) {
      const runs = sheet && sheet.title ? sheet.title.runs : null;
      const props = this.titleStyle();
      if (!runs || !runs.length) return { text: sheetLabel, props };
      let text = runs.map(r => r.text).join("").replace(/\u00C6\r?\n?/g, "\n");
      text = text.replace(/<([^<>]+)>/g, (m, inner) => {
        if (/^sheet name$/i.test(inner)) return sheetLabel;
        if (/^(workbook name|page name|page count|page number)$/i.test(inner)) return "";
        if (!/^\[/.test(inner)) return m;
        const ref = tfParseFieldRef(inner);
        if (!ref) return m;
        if (ref.ds === "Parameters") {
          const info = tfFieldInfo(model, ref) || {};
          return info.alias || info.value || "";
        }
        const vals = valuesFor(ref);
        if (!vals.length) return "";
        return vals.length <= 3 ? vals.join(", ") : "*";
      });
      return { text: text.trim(), props };
    },

    /* display names for Measure-Names buckets like "[ds].[sum:Sales:qk]" */
    bucketAliases(bucketKey) {
      if (!/^\[.*\]$/.test(bucketKey)) return [bucketKey];
      // Tableau escapes some characters in a member's field name ("[sum:gray>50\% (copy 2)_…]")
      const plain = bucketKey.replace(/\\(.)/g, "$1");
      const ref = tfParseFieldRef(plain);
      return ref ? [bucketKey, ...(plain !== bucketKey ? [plain] : []), ...tfDisplayNames(model, ref)] : [bucketKey];
    }
  };
}


/* ═══ format/excel-style.js ═══════════════════════════════════════════════════════════════════════ */
/* Resolved Tableau style props → ExcelJS font / fill / alignment, and cell writing. */

function tfExcelFont(p) {
  let name = p.fontName || "Arial";
  let bold = !!p.bold;
  if (FORMAT_CONFIG.substituteTableauFonts && /^tableau\b/i.test(name)) {
    if (/\b(bold|semibold|black|heavy)\b/i.test(name)) bold = true;  // weight lives in the family name
    name = FORMAT_CONFIG.tableauFontSubstitute;
  }
  const f = { name, size: p.fontSize || 9, bold, italic: !!p.italic };
  if (p.underline) f.underline = true;
  if (p.strike) f.strike = true;
  if (p.color) f.color = { argb: p.color };
  return f;
}

function tfExcelAlignment(p, isNumber) {
  // a number shrinks to its column (columns keep the dashboard's widths) instead of showing ####
  return { horizontal: p.hAlign || (isNumber ? "right" : "left"), vertical: p.vAlign || "middle", wrapText: !!p.wrap,
           ...(isNumber && !p.wrap ? { shrinkToFit: true } : {}) };
}

/** @param {string} argb @returns {import("exceljs").Fill | undefined} */
function tfExcelFill(argb) {
  return argb ? { type: "pattern", pattern: "solid", fgColor: { argb } } : undefined;
}

function tfRichRuns(runs, baseProps) {
  const out = [];
  runs.forEach(r => {
    const text = r.text.replace(/\u00C6\r?\n?/g, "\n");
    if (text) out.push({ text, font: tfExcelFont(tfMerge(baseProps, r.props)) });
  });
  return out;
}

/* writes value + style; returns true if the cell holds a real number */
function tfWriteCell(cell, dv, p, extra = {}) {
  const native = dv ? (dv.nativeValue !== undefined ? dv.nativeValue : dv.value) : undefined;
  const formatted = dv ? (dv.formattedValue != null ? dv.formattedValue : (native != null ? String(native) : "")) : "";
  let isNumber = false;
  if (dv && (dv.value === null || dv.nativeValue === null || dv.value === "%null%")) {   // Tableau Null → blank
    cell.value = "";
  } else if (FORMAT_CONFIG.writeNativeNumbers && typeof native === "number" && isFinite(native)) {
    const numFmt = p.numFmt || inferExcelNumFmt(formatted, native);
    if (numFmt) { cell.value = native; cell.numFmt = numFmt; isNumber = true; }
  }
  if (!isNumber && cell.value !== "") cell.value = formatted;
  if (extra.hyperlink) cell.value = { text: extra.hyperlinkText || extra.hyperlink, hyperlink: extra.hyperlink };
  const looksNumeric = isNumber || /^[^\d]{0,3}-?[\d.,]+[^\d]{0,3}$/.test(String(formatted).trim());
  const fill = extra.fill !== undefined ? extra.fill : p.bgColor;
  const font = tfExcelFont(extra.fontColor ? { ...p, color: extra.fontColor } : p);
  if (fill && !extra.fontColor && !p.explicitColor && tfBrightness(fill) < 128) font.color = { argb: "FFFFFFFF" };
  cell.font = font;
  cell.alignment = tfExcelAlignment(p, looksNumeric);
  if (fill) cell.fill = tfExcelFill(fill);
  if (extra.border) {
    const b = {};
    for (const k in extra.border) if (extra.border[k]) b[k] = extra.border[k];
    if (Object.keys(b).length) cell.border = b;
  }
  return isNumber;
}


/* ═══ twb/parser.js ═══════════════════════════════════════════════════════════════════════════════ */
/* TWB/TWBX XML → plain JSON format model (worksheets, panes, styles, dashboards, actions). */

function tfParseRunProps(run) {
  const a = n => run.getAttribute(n);
  const align = { "0": "left", "1": "center", "2": "right" }[a("fontalignment")];
  return tfDefined({
    fontName: a("fontname") || undefined,
    fontSize: tfNum(a("fontsize")),
    bold: a("bold") != null ? a("bold") === "true" : undefined,
    italic: a("italic") != null ? a("italic") === "true" : undefined,
    underline: a("underline") != null ? a("underline") === "true" : undefined,
    color: tfArgb(a("fontcolor")) || undefined,
    hAlign: align
  });
}

/* a dashboard zone's Layout pane formatting: background and border */
function tfZoneStyle(el) {
  if (!el) return undefined;
  const f = {};
  tfKids(el, "format").forEach(x => { f[x.getAttribute("attr")] = x.getAttribute("value"); });
  const style = tfDefined({
    bgColor: tfArgb(f["background-color"]) || undefined,
    borderColor: tfArgb(f["border-color"]) || undefined,
    borderStyle: f["border-style"] || undefined,
    borderWidth: tfNum(f["border-width"])
  });
  return Object.keys(style).length ? style : undefined;
}

function tfParseRuns(ftEl) {
  return tfKids(ftEl, "run").map(r => ({ text: r.textContent || "", props: tfParseRunProps(r) }));
}

function tfBucketKey(t) {
  let s = String(t == null ? "" : t).trim();
  if (/^".*"$/.test(s)) s = s.slice(1, -1).replace(/""/g, '"');
  else if (/^#.*#$/.test(s)) s = s.slice(1, -1);
  if (s === "%null%") s = "null";
  return s.toLowerCase();
}

function tfParseColorEncoding(enc) {
  const a = n => enc.getAttribute(n);
  const cp = enc.getElementsByTagName("color-palette")[0] || null;
  const customColors = cp
    ? Array.from(cp.getElementsByTagName("color")).map(c => tfArgb(c.textContent)).filter(Boolean)
    : [];
  const map = {};
  Array.from(enc.getElementsByTagName("map")).forEach(m => {
    const color = tfArgb(m.getAttribute("to"));
    if (!color) return;
    Array.from(m.getElementsByTagName("bucket")).forEach(b => { map[tfBucketKey(b.textContent)] = color; });
  });
  return {
    field: tfParseFieldRef(a("field")),
    type: a("type"),                                  // "palette" (categorical) | "interpolated" | …
    paletteName: a("palette") || (cp && cp.getAttribute("name")) || null,
    paletteType: cp ? cp.getAttribute("type") : null, // ordered-sequential | ordered-diverging | regular
    customColors,
    reverse: a("reverse") === "true" || (cp && cp.getAttribute("reverse") === "true"),
    // optional range settings (only used if present in your TWB)
    center: tfNum(a("center")), min: tfNum(a("min")), max: tfNum(a("max")),
    steps: tfNum(a("num-steps")),                     // Stepped Color: that many solid colours instead of a gradient
    map
  };
}

/** @param {Element | null} styleEl @returns {ParsedStyle} */
function tfParseStyle(styleEl) {
  /** @type {ParsedStyle["rules"]} */
  const rules = {};
  const encodings = [];
  if (!styleEl) return { rules, encodings };
  tfKids(styleEl, "style-rule").forEach(rule => {
    const el = rule.getAttribute("element") || "all";
    rules[el] = rules[el] || [];
    tfKids(rule, "format").forEach(f => rules[el].push(tfDefined({
      attr: f.getAttribute("attr"),
      value: f.getAttribute("value"),
      field: f.getAttribute("field") || undefined,
      scope: f.getAttribute("scope") || undefined,
      dataClass: f.getAttribute("data-class") || undefined,
      axisClass: f.getAttribute("class") || undefined
    })));
  });
  Array.from(styleEl.getElementsByTagName("encoding"))
    .filter(e => e.getAttribute("attr") === "color")
    .forEach(e => encodings.push(tfParseColorEncoding(e)));
  // Edit Axis: fixed range, tick spacing, "include zero" (encoding attr="space")
  const spaces = Array.from(styleEl.getElementsByTagName("encoding")).filter(e => e.getAttribute("attr") === "space").map(e => tfDefined({
    field: e.getAttribute("field") || undefined,
    scope: e.getAttribute("scope") || undefined,
    axisClass: e.getAttribute("class") || undefined,
    rangeType: e.getAttribute("range-type") || undefined,
    min: tfNum(e.getAttribute("min")),
    max: tfNum(e.getAttribute("max")),
    majorSpacing: tfNum(e.getAttribute("major-spacing")),
    majorShow: e.getAttribute("major-show") || undefined,      // "false": Edit Axis → Tick Marks → None
    domainExpand: e.getAttribute("domain-expand") || undefined,
    reverse: e.getAttribute("reverse") === "true" || undefined, // Edit Axis → Scale → Reversed (rank 1 at the top)
    fold: e.getAttribute("fold") === "true" || undefined,       // Dual Axis: this axis is folded onto the one before it
    synchronized: e.getAttribute("synchronized") === "true" || undefined   // … with the same scale (Synchronize Axis)
  }));
  // Shape encodings: value → shape ("Zoom Icons/Zoom in.png" = a custom shape, ":filled/circle" = Tableau's)
  const shapes = Array.from(styleEl.getElementsByTagName("encoding")).filter(e => e.getAttribute("attr") === "shape").map(e => {
    /** @type {Record<string, string>} */
    const map = {};
    Array.from(e.getElementsByTagName("map")).forEach(m => Array.from(m.getElementsByTagName("bucket"))
      .forEach(b => { map[tfBucketKey(b.textContent)] = m.getAttribute("to"); }));
    return { field: tfParseFieldRef(e.getAttribute("field")), map };
  }).filter(s => s.field);
  /** @type {ParsedStyle} */
  const out = { rules, encodings };
  if (spaces.length) out.spaces = spaces;
  if (shapes.length) out.shapes = shapes;
  return out;
}

function tfParseTitle(ownerEl) {
  const lo = tfKid(ownerEl, "layout-options") || tfKid(ownerEl, "layout");
  const t = lo && tfKid(lo, "title");
  const ft = t && tfKid(t, "formatted-text");
  if (!ft) return null;
  const runs = tfParseRuns(ft);
  return { text: runs.map(r => r.text).join("").trim(), runs };
}

/**
 * A calculation that is one number ("0.5", "MIN(0)"): an axis that only positions marks, e.g. the columns
 * of a table built from marks.
 * @param {Element | null} calc @returns {number | undefined}
 */
function tfConstantFormula(calc) {
  const m = calc && String(calc.getAttribute("formula") || "").match(/^\s*(?:(?:MIN|MAX|AVG|SUM|ATTR)\s*\(\s*(-?\d+(?:\.\d+)?)\s*\)|(-?\d+(?:\.\d+)?))\s*$/i);
  return m ? Number(m[1] ?? m[2]) : undefined;
}

/**
 * What kind of calculation a calculated field is, from its formula – for the conversion report (the export
 * writes the values Tableau computed; formulas are not translated).
 * @param {string} formula @returns {string}
 */
function tfCalcKind(formula) {
  const f = String(formula || "").replace(/\/\/[^\n]*|\/\*[\s\S]*?\*\//g, " ");
  if (/\{\s*(FIXED|INCLUDE|EXCLUDE)\b|\{\s*(SUM|AVG|MIN|MAX|COUNTD?)\s*\(/i.test(f)) return "LOD expression";
  if (/\b(WINDOW_\w+|RUNNING_\w+|RANK(_\w+)?|INDEX|FIRST|LAST|LOOKUP|PREVIOUS_VALUE|SIZE|TOTAL)\s*\(/i.test(f)) return "table calculation";
  if (/\b(SUM|AVG|MIN|MAX|COUNTD?|MEDIAN|ATTR|STDEVP?|VARP?|PERCENTILE)\s*\(/i.test(f)) return "aggregation";
  if (/\b(DATEPART|DATETRUNC|DATEADD|DATEDIFF|DATENAME|DATEPARSE|MAKEDATE|TODAY|NOW|YEAR|QUARTER|MONTH|WEEK|DAY)\s*\(/i.test(f)) return "date calculation";
  if (/\b(IF|CASE|IIF|ZN|IFNULL|ISNULL)\b/i.test(f)) return "conditional logic";
  if (/\b(MAKEPOINT|MAKELINE|BUFFER|DISTANCE|SCRIPT_\w+|RAWSQL\w*|REGEXP_\w+|SPLIT)\s*\(/i.test(f)) return "Tableau-specific function";
  return /[-+*/]/.test(f) ? "arithmetic" : "expression";
}

/** a datasource column's calculation: a constant, a bin, or a calculated field (its kind and formula)
 * @param {Element | null} calc */
function tfCalcInfo(calc) {
  if (!calc) return {};
  const cls = calc.getAttribute("class") || "";
  const formula = calc.getAttribute("formula") || "";
  if (cls === "bin") return { bin: true };
  if (cls === "categorical-bin") return { group: true };
  if (!formula) return {};
  return { constant: tfConstantFormula(calc), calcKind: tfCalcKind(formula), formula: formula.length > 160 ? formula.slice(0, 157) + "…" : formula,
           ...(formula.length > 160 ? { formulaFull: formula } : {}) };
}

/* Raise when the parser reads something new: a workbook model remembered by an older version is then
   parsed again from its stored XML (ui/workbook-store.js), so an update reaches workbooks loaded before it. */
const FORMAT_MODEL_VERSION = 10;

/** @param {string} xmlString the .twb XML @returns {FormatModel} */
function parseTableauFormatting(xmlString) {
  const doc = new DOMParser().parseFromString(xmlString, "text/xml");
  const root = doc.documentElement;
  const model = { version: FORMAT_MODEL_VERSION, workbookStyle: tfParseStyle(tfKid(root, "style")),
                  fields: {}, datasourceStyles: {}, sheets: {}, dashboards: {}, measureAliases: {} };

  // ── datasource columns: captions, roles, default number formats, colour maps
  tfKids(tfKid(root, "datasources"), "datasource").forEach(ds => {
    const dsName = ds.getAttribute("name");
    model.datasourceStyles[dsName] = tfParseStyle(tfKid(ds, "style"));
    tfKids(ds, "column").forEach(col => {
      const n = (col.getAttribute("name") || "").replace(/^\[|\]$/g, "");
      if (!n) return;
      const info = tfDefined({
        name: n, ds: dsName,
        caption: col.getAttribute("caption") || undefined,
        role: col.getAttribute("role") || undefined,
        datatype: col.getAttribute("datatype") || undefined,
        defaultFormat: col.getAttribute("default-format") || undefined,
        alias: col.getAttribute("alias") || undefined,      // parameters: current value's display text
        value: col.getAttribute("value") || undefined,
        param: col.getAttribute("param-domain-type") ? true : undefined,
        geoRole: col.getAttribute("semantic-role") || undefined,   // "[State].[Name]": Tableau geocodes it on a map
        ...tfCalcInfo(tfKid(col, "calculation"))
      });
      model.fields[dsName + "|" + n.toLowerCase()] = info;
      if (!model.fields["|" + n.toLowerCase()]) model.fields["|" + n.toLowerCase()] = info;
      // Measure Names aliases: key "[ds].[sum:HomeValue:qk]" → "Typical home value"
      if (n === ":Measure Names") {
        Array.from(col.getElementsByTagName("alias")).forEach(al => {
          const r = tfParseFieldRef(al.getAttribute("key"));
          if (r) model.measureAliases[tfRefKey(r)] = al.getAttribute("value");
        });
      }
    });
  });

  // ── worksheets
  tfKids(tfKid(root, "worksheets"), "worksheet").forEach(ws => {
    const name = ws.getAttribute("name");
    const table = tfKid(ws, "table");
    const sheet = {
      name,
      title: tfParseTitle(ws),
      style: tfParseStyle(table && tfKid(table, "style")),
      rows: tfExtractRefs(table && tfKid(table, "rows") && tfKid(table, "rows").textContent),
      cols: tfExtractRefs(table && tfKid(table, "cols") && tfKid(table, "cols").textContent),
      // Analysis → Totals → Show Row / Column Grand Totals (<rows total='true'> / <cols total='true'>)
      grandTotals: { rows: !!(table && tfKid(table, "rows") && tfKid(table, "rows").getAttribute("total") === "true"),
                     cols: !!(table && tfKid(table, "cols") && tfKid(table, "cols").getAttribute("total") === "true") },
      panes: [], fieldRefs: []
    };
    const view = table && tfKid(table, "view");
    // columns declared in this sheet's datasource-dependencies (covers blended/secondary sources)
    Array.from(ws.getElementsByTagName("datasource-dependencies")).forEach(dep => {
      const dsName = dep.getAttribute("datasource");
      tfKids(dep, "column").forEach(col => {
        const n = (col.getAttribute("name") || "").replace(/^\[|\]$/g, "");
        if (!n) return;
        const key = dsName + "|" + n.toLowerCase();
        const cur = model.fields[key] || {};
        // a calculation typed into a shelf exists only here (an unnamed calc: "avg(0)" placing a butterfly's labels)
        model.fields[key] = tfMerge(tfDefined({ name: n, ds: dsName, caption: col.getAttribute("caption") || undefined,
          role: col.getAttribute("role") || undefined, datatype: col.getAttribute("datatype") || undefined,
          defaultFormat: col.getAttribute("default-format") || undefined,
          alias: col.getAttribute("alias") || undefined, value: col.getAttribute("value") || undefined,
          ...tfCalcInfo(tfKid(col, "calculation")) }), cur);
        if (!model.fields["|" + n.toLowerCase()]) model.fields["|" + n.toLowerCase()] = model.fields[key];
      });
    });
    sheet.manualSorts = view ? Array.from(view.getElementsByTagName("manual-sort")).map(ms => ({
      field: tfParseFieldRef(ms.getAttribute("column")),
      direction: (ms.getAttribute("direction") || "ASC").toUpperCase(),
      order: Array.from(ms.getElementsByTagName("bucket")).map(b => b.textContent)
    })) : [];
    // Sort → Alphabetic, Descending (Z → A): a field's members in reverse natural order
    sheet.naturalSorts = view ? Array.from(view.getElementsByTagName("natural-sort")).map(ns => ({
      field: tfParseFieldRef(ns.getAttribute("column")),
      direction: (ns.getAttribute("direction") || "ASC").toUpperCase()
    })).filter(x => x.field) : [];
    sheet.measureSorts = view ? [
      ...Array.from(view.getElementsByTagName("shelf-sort-v2")).map(ss => ({
        field: tfParseFieldRef(ss.getAttribute("dimension-to-sort")),
        measure: tfParseFieldRef(ss.getAttribute("measure-to-sort-by")),
        direction: (ss.getAttribute("direction") || "ASC").toUpperCase() })),
      ...Array.from(view.getElementsByTagName("computed-sort")).map(cs => ({
        field: tfParseFieldRef(cs.getAttribute("column")),
        measure: tfParseFieldRef(cs.getAttribute("using")),
        direction: (cs.getAttribute("direction") || "ASC").toUpperCase() }))
    ].filter(x => x.field && x.measure) : [];
    // members kept by the Measure Names filter, in filter order
    sheet.measureFilter = [];
    // the fields the sheet filters by (a calculated filter's own datasource: two calculations can share a caption)
    sheet.filterRefs = view ? tfKids(view, "filter").map(f => tfParseFieldRef(f.getAttribute("column") || "")).filter(Boolean) : [];
    if (view) tfKids(view, "filter").forEach(f => {
      if (!/\[:Measure Names\]$/.test(f.getAttribute("column") || "")) return;
      Array.from(f.getElementsByTagName("groupfilter")).forEach(g => {
        const m = g.getAttribute("member");
        if (g.getAttribute("function") === "member" && m) sheet.measureFilter.push(m);
      });
    });
    const panesEl = table && tfKid(table, "panes");
    tfKids(panesEl, "pane").forEach(p => {
      const mark = tfKid(p, "mark");
      const encodings = tfKids(tfKid(p, "encodings")).map(e => ({
        channel: e.tagName, field: tfParseFieldRef(e.getAttribute("column"))
      })).filter(e => e.field);
      const cl = tfKid(p, "customized-label");
      const labelRuns = cl ? tfParseRuns(tfKid(cl, "formatted-text")) : [];
      const ct = tfKid(p, "customized-tooltip");                 // Tooltip editor text
      const tooltipRuns = ct ? tfParseRuns(tfKid(ct, "formatted-text")) : [];
      sheet.panes.push({
        id: p.getAttribute("id") || null,
        xIndex: tfNum(p.getAttribute("x-index")),
        xAxisName: p.getAttribute("x-axis-name") || null,
        yIndex: tfNum(p.getAttribute("y-index")),
        yAxisName: p.getAttribute("y-axis-name") || null,     // dual axis / multi-pane: which measure this pane draws
        markClass: mark ? mark.getAttribute("class") : "Automatic",
        hidden: p.getAttribute("hidden") === "true",          // a map layer switched off in Tableau
        encodings,
        labelRuns: labelRuns.map(r => ({ refs: tfExtractRefs(r.text.replace(/^<|>$/g, "")), text: r.text, props: r.props })),
        tooltipRuns,
        style: tfParseStyle(tfKid(p, "style"))
      });
    });
    // Analytics pane box plot: <reference-line formula='iqr' boxplot-whisker-type='…'>
    sheet.boxPlot = Array.from(ws.getElementsByTagName("reference-line")).some(rl =>
      rl.getAttribute("formula") === "iqr" || rl.getAttribute("boxplot-whisker-type") != null);
    // reference lines (Analytics → Reference Line); box-plot whiskers / distributions are not lines
    const refIds = new Set();
    sheet.referenceLines = Array.from(ws.getElementsByTagName("reference-line")).filter(rl => {
      const id = rl.getAttribute("id") || "";
      if (refIds.has(id) || rl.getAttribute("boxplot-whisker-type") != null) return false;
      refIds.add(id);
      return /^(average|mean|median|sum|total|min|max|constant)$/.test(rl.getAttribute("formula") || "");
    }).map(rl => tfDefined({
      formula: rl.getAttribute("formula"),
      labelType: rl.getAttribute("label-type") || "automatic",
      label: rl.getAttribute("label") || undefined,
      scope: rl.getAttribute("scope") || undefined,
      value: tfNum(rl.getAttribute("value")),
      axis: tfParseFieldRef(rl.getAttribute("axis-column")) || undefined,
      field: tfParseFieldRef(rl.getAttribute("value-column")) || undefined
    }));
    // every field the sheet references (for name matching incl. Measure Names)
    const seen = new Set();
    const add = r => { if (r && !seen.has(r.inner.toLowerCase())) { seen.add(r.inner.toLowerCase()); sheet.fieldRefs.push(r); } };
    sheet.runningTotals = [];                                    // waterfall: running-sum table calcs
    sheet.tableCalcs = {};                                       // quick table calc → its "compute using" (ordering-type)
    Array.from(ws.getElementsByTagName("column-instance")).forEach(ci => {
      const depDs = ci.parentNode && ci.parentNode.getAttribute && ci.parentNode.getAttribute("datasource");
      const ref = tfParseFieldRef((depDs ? "[" + depDs + "]." : "") + ci.getAttribute("name"));
      add(ref);
      const calc = tfKid(ci, "table-calc");
      if (ref && calc && /^(cumtotal|runningtotal)$/i.test(calc.getAttribute("type") || "")) sheet.runningTotals.push(ref);
      if (ref && calc) sheet.tableCalcs[ref.inner.toLowerCase()] = calc.getAttribute("ordering-type") || "";
    });
    sheet.rows.forEach(add); sheet.cols.forEach(add);
    sheet.panes.forEach(p => p.encodings.forEach(e => add(e.field)));
    model.sheets[name] = sheet;
  });

  // ── URL actions (not filter/highlight actions) → Excel hyperlinks
  model.actions = [];
  tfKids(tfKid(root, "actions"), "action").forEach(a => {
    const link = tfKid(a, "link"), src = tfKid(a, "source");
    if (!link || tfKid(a, "command")) return;                   // filter/highlight actions have a <command>
    const expr = link.getAttribute("expression") || "";
    if (!expr || /^tsl:/i.test(expr)) return;
    model.actions.push(tfDefined({
      caption: a.getAttribute("caption") || undefined, expression: expr,
      dashboard: src ? src.getAttribute("dashboard") || undefined : undefined,
      worksheet: src ? src.getAttribute("worksheet") || undefined : undefined,
      exclude: src ? Array.from(src.getElementsByTagName("exclude-sheet")).map(x => x.getAttribute("name")) : []
    }));
  });

  // ── dashboards (title + style)
  tfKids(tfKid(root, "dashboards"), "dashboard").forEach(d => {
    const zones = [];
    const walk = (z, parent) => tfKids(z, "zone").forEach(c => {
      zones.push(tfDefined({
        id: c.getAttribute("id"), name: c.getAttribute("name") || undefined,
        type: c.getAttribute("type-v2") || "worksheet",
        showTitle: c.getAttribute("show-title") !== "false",
        customTitle: c.getAttribute("custom-title") === "true" || undefined,   // runs: the title shown instead of the field's
        hidden: c.getAttribute("hidden-by-user") === "true",
        x: tfNum(c.getAttribute("x")), y: tfNum(c.getAttribute("y")),
        w: tfNum(c.getAttribute("w")), h: tfNum(c.getAttribute("h")),
        runs: tfKid(c, "formatted-text") ? tfParseRuns(tfKid(c, "formatted-text")) : undefined,
        param: c.getAttribute("param") || undefined,             // image zones: the file inside the .twbx; legends: their field
        mode: c.getAttribute("mode") || undefined,               // filter cards: dropdown, checkdropdown, checklist, radiolist, slider …
        scaled: c.getAttribute("is-scaled") === "1" || undefined,
        // Center Image is on unless switched off (is-centered='0')
        centered: c.getAttribute("is-centered") ? c.getAttribute("is-centered") === "1" : undefined,
        url: c.getAttribute("url") || undefined,                   // image / button zones: the link it opens
        parent,                                                     // the layout container it sits in
        style: tfZoneStyle(tfKid(c, "zone-style"))
      }));
      walk(c, c.getAttribute("id") || undefined);
    });
    walk(tfKid(d, "zones"), undefined);   // direct child only → phone/tablet layouts are skipped
    const size = tfKid(d, "size");
    const sid = tfKid(d, "simple-id");
    model.dashboards[d.getAttribute("name")] = {
      title: tfParseTitle(d), style: tfParseStyle(tfKid(d, "style")), zones,
      uuid: sid ? sid.getAttribute("uuid") || undefined : undefined,
      toggles: {},                          // zone id → the boolean field that shows it (dynamic zone visibility, below)
      width: size ? tfNum(size.getAttribute("maxwidth")) : undefined,
      height: size ? tfNum(size.getAttribute("maxheight")) : undefined
    };
  });

  // ── dynamic zone visibility (Layout → Control visibility using value): a zone is shown while a boolean field is
  //    True. <datagraph>: a zone's visibility input ← a field node's value output. The field usually tests a parameter
  //    ([Parameters].[Parameter 1]='Race') that buttons set (parameter actions): the zone is one of the views they
  //    switch between. dashboards[d].toggles[zone id] = { field, param, value } (param / value: when it is such a test)
  const graph = tfKid(tfKid(root, "datagraph"), "graph");
  if (graph) {
    const nodes = tfKids(tfKid(graph, "nodes"));
    const fieldOfOutput = new Map(nodes.filter(n => n.tagName === "single-value-field-node")
      .map(n => [n.getAttribute("value-output-guid"), n.getAttribute("fieldname")]));
    const fromOf = new Map(tfKids(tfKid(graph, "edges"), "edge").map(e => [e.getAttribute("to"), e.getAttribute("from")]));
    nodes.filter(n => n.tagName === "dashboard-zone-visibility-node").forEach(n => {
      const dash = Object.values(model.dashboards).find(x => x.uuid && x.uuid === n.getAttribute("dashboard-identifier"));
      const field = fieldOfOutput.get(fromOf.get(n.getAttribute("visibility-input-guid")) || "");
      if (!dash || !field) return;
      const ref = tfParseFieldRef(field);
      const info = ref && tfFieldInfo(model, ref);
      const test = info && String(info.formula || "").trim()
        .match(/^\[Parameters\]\.\[([^\]]+)\]\s*==?\s*(?:'((?:[^']|'')*)'|"((?:[^"]|"")*)")$/);
      const value = test && (test[2] !== undefined ? test[2].replace(/''/g, "'") : test[3].replace(/""/g, '"'));
      dash.toggles[n.getAttribute("zone-id")] = tfDefined({ field, param: test ? test[1] : undefined, value: test ? value : undefined });
    });
  }
  // ── parameter actions (a click on a sheet's mark sets a parameter: the "buttons" of a dashboard)
  model.paramActions = tfKids(tfKid(root, "actions"), "edit-parameter-action").map(a => {
    const src = tfKid(a, "source");
    const p = name => { const x = tfKids(tfKid(a, "params"), "param").find(q => q.getAttribute("name") === name); return x ? x.getAttribute("value") || "" : ""; };
    const target = p("target-parameter").match(/^\[Parameters\]\.\[(.+)\]$/);
    return tfDefined({ dashboard: src ? src.getAttribute("dashboard") || undefined : undefined,
                       worksheet: src ? src.getAttribute("worksheet") || undefined : undefined,
                       param: target ? target[1] : undefined, field: p("source-field") || undefined });
  }).filter(a => a.worksheet && a.param);

  // ── each sheet's fit on each dashboard (<window class='dashboard'> viewpoints: "entire-view", "fit-width" …)
  tfKids(tfKid(root, "windows"), "window").forEach(w => {
    const dash = w.getAttribute("class") === "dashboard" && model.dashboards[w.getAttribute("name")];
    if (!dash) return;
    dash.fit = {};
    tfKids(tfKid(w, "viewpoints"), "viewpoint").forEach(v => {
      const z = tfKid(v, "zoom");
      if (z) dash.fit[v.getAttribute("name")] = z.getAttribute("type");
    });
  });

  tfDebugDump(model);
  return model;
}

/* Prints every style element/attr found – use this to see what YOUR TWB contains */
function tfDebugDump(model) {
  if (!FORMAT_CONFIG.debug) return;
  const summarize = st => {
    const o = {};
    for (const [el, list] of Object.entries(st.rules)) o[el] = [...new Set(list.map(f => f.attr + (f.scope ? "@" + f.scope : "") + (f.field ? "[field]" : "")))].join(", ");
    return o;
  };
  console.log("[Format] Workbook style:", summarize(model.workbookStyle));
  for (const s of Object.values(model.sheets)) {
    console.log(`[Format] Sheet "${s.name}"`, {
      rules: summarize(s.style),
      marks: s.panes.map(p => p.markClass).join(", "),
      encodings: s.panes.flatMap(p => p.encodings.map(e => e.channel + "=" + e.field.inner)).join(", "),
      colorEncodings: s.style.encodings.length,
      labelRuns: s.panes.reduce((n, p) => n + p.labelRuns.length, 0)
    });
  }
}


/* ═══ charts/model/common.js ══════════════════════════════════════════════════════════════════════ */
/* Shared helpers for building renderer-neutral Excel chart specs. */

"use strict";

const TV_MAX_POINTS = 4000;

const TV_MAX_SERIES = 60;

const TV_DEFAULT_MARK_COLOR = "4E79A7";

function tvHex(argb) {
  if (!argb) return null;
  const s = String(argb).replace(/^#/, "").toUpperCase();
  return s.length === 8 ? s.slice(2) : s;
}

function tvText(dv) { return tfIsNull(dv) ? "Null" : tfDvText(dv); }

function tvMarkToken(cls) {
  const t = String(cls || "").toLowerCase().replace(/[\s_-]+/g, "");
  return t === "automatic" ? "" : t;
}

function tvPaneRule(pane, element, attrName) {
  const list = (pane && pane.style && pane.style.rules && pane.style.rules[element]) || [];
  const f = list.find(x => x.attr === attrName && !x.field && !x.scope);
  return f ? f.value : undefined;
}

/* ── formatting helpers ─────────────────────────────────────────────────── */
function tvChartFont(fmt) {
  const f = tfExcelFont(fmt.headerCellStyle(null));
  return { name: f.name, size: f.size || 9, color: tvHex(f.color && f.color.argb) || "333333" };
}

/** @param {ViewModel} vm @param {number} ci @returns {string} */
function tvNumFmt(vm, ci) {
  const c = vm.cols[ci];
  const st = vm.fmt.sheetModel ? vm.fmt.sheetModel.style : null;
  if (c && c.ref && st) {
    const axisRaw = tfCollect(st, ["axis"], { field: c.ref }).numFmtRaw;
    const axis = axisRaw ? tableauToExcelNumFmt(axisRaw) : null;
    if (axis) return axis;
  }
  const own = c && vm.fmt.markCellStyle(c.ref || null).numFmt;
  if (own) return own;
  for (const row of vm.rows) {
    const dv = row[ci];
    if (tfIsNull(dv)) continue;
    const n = tfDvNum(dv);
    if (n === null) continue;
    return inferExcelNumFmt(dv.formattedValue, n) || "General";
  }
  return "General";
}

/** @param {ViewModel} vm @param {number} ci @returns {string} */
function tvMeasureLabel(vm, ci) {
  const c = vm.cols[ci];
  return (c && (c.label || c.name)) || "Value";
}

/* mark of the pane that draws this measure (dual axis / multi-pane), else the sheet's.
 * Without a workbook, the live spec has one marks card per measure, in shelf order. */
/** @param {Roles} roles @param {FieldRef | null} ref @param {string} [fallback] @param {number} [index] @param {number} [count] @returns {string} */
function tvMeasureMark(roles, ref, fallback, index, count) {
  if (!roles.panes.length && roles.liveMarks && count > 1 && roles.liveMarks.length === count && roles.liveMarks[index]) {
    return roles.liveMarks[index];
  }
  if (ref) {
    const p = roles.panes.find(x => [x.yAxisName, x.xAxisName].some(n => n && tfSameField(tfParseFieldRef(n), ref)));
    const t = p && tvMarkToken(p.markClass);
    if (t) return t;
  }
  const all = roles.panes.find(x => !x.id) || roles.panes[0];
  return (all && tvMarkToken(all.markClass)) || fallback || "bar";
}

/** @param {Roles} roles @param {FieldRef | null} ref @returns {boolean} */
function tvLabelsOn(roles, ref) {
  const shown = roles.panes.some(p => tvPaneRule(p, "mark", "mark-labels-show") === "true");
  if (shown) return true;
  if (!ref) return false;
  return roles.labelRefs.some(r => tfSameField(r, ref) || (r.name === "Multiple Values")) ||
    roles.labelNames.some(n => n === tfNorm(ref.name));
}

/** @param {Roles} roles @returns {string} */
function tvMarkColor(roles) {
  for (const p of roles.panes) {
    const c = tvHex(tfArgb(tvPaneRule(p, "mark", "mark-color")));
    if (c) return c;
  }
  return TV_DEFAULT_MARK_COLOR;
}

/* ── colour scale for the Color encoding (TWB mapping or Tableau automatic) ─ */
/**
 * @param {ViewModel} vm @param {Roles} roles @param {string} [markToken]
 * @returns {((value: any, index?: number) => string | null) | null} value → hex colour
 */
function tvColorScale(vm, roles, markToken) {
  if (!roles.color) return null;
  const fmt = vm.fmt;
  /** @type {Partial<ReturnType<SheetFormatter["colorEncoding"]>>} */
  let enc = fmt.colorEncoding();
  if (!enc) {                                      // live spec only: Tableau automatic palettes
    const ci = roles.color.ci;
    const name = ci >= 0 ? vm.cols[ci].name : "Measure Names";
    enc = { ref: { name, inner: name }, def: null, continuous: !!roles.color.continuous, markClass: markToken || "bar" };
  }
  if (roles.color.measureNames) {
    const scale = tfBuildColorScale(fmt, { ...enc, continuous: false }, []);
    return (value, i) => tvHex(scale && scale(value)) || tvHex(TABLEAU_10[i % TABLEAU_10.length]);
  }
  const ci = roles.color.ci;
  const values = vm.rows.map(r => roles.color.continuous ? tfDvNum(r[ci]) : tvText(r[ci]));
  const scale = tfBuildColorScale(fmt, enc, values);
  return scale ? (value => tvHex(scale(value))) : null;
}

/* legend order of a discrete colour field: manual sort, else natural */
/** @param {ViewModel} vm @param {Roles} roles @returns {string[]} */
function tvColorValues(vm, roles) {
  const ci = roles.color.ci;
  const seen = new Map();
  vm.rows.forEach(r => { const t = tvText(r[ci]); if (!seen.has(t)) seen.set(t, r[ci]); });
  let values = [...seen.keys()];
  const manual = roles.color.ref ? vm.fmt.manualSortFor(roles.color.ref) : null;
  const byMeasure = !manual && roles.color.ref ? vm.fmt.measureSortFor(roles.color.ref) : null;
  if (manual) {
    const rank = new Map(manual.order.map((b, i) => [tfBucketKey(b), i]));
    values.sort((a, b) => (rank.has(tfNorm(a)) ? rank.get(tfNorm(a)) : 1e9) - (rank.has(tfNorm(b)) ? rank.get(tfNorm(b)) : 1e9));
    if (manual.direction === "DESC") values.reverse();
  } else if (byMeasure) {
    // sorted by a measure (Sort → Field: Value Ordered, ascending): its total per colour value
    const mi = vm.cols.findIndex(c => c.ref && tfSameField(c.ref, byMeasure.measure));
    const total = new Map(values.map(v => [v, 0]));
    if (mi >= 0) vm.rows.forEach(r => { const t = tvText(r[ci]); total.set(t, (total.get(t) || 0) + (tfDvNum(r[mi]) || 0)); });
    values.sort((a, b) => (total.get(a) - total.get(b)) * (byMeasure.direction === "DESC" ? -1 : 1) || tfNaturalCompare(seen.get(a), seen.get(b)));
  } else values.sort((a, b) => tfNaturalCompare(seen.get(a), seen.get(b)));
  return values;
}

/** a header's date label format (Format → Header → Dates: "iLLLLL" → J F M …)
 * @param {ViewModel} vm @param {number} ci @returns {(text: string, dv: DataValue) => string} */
function tvLabelFormatter(vm, ci) {
  const col = vm.cols[ci];
  const raw = col && col.ref && vm.fmt.labelFormat ? vm.fmt.labelFormat(col.ref) : null;
  if (!raw) return t => t;
  return (text, dv) => tfFormatDateLabel(raw, text, dv) ?? text;
}

/**
 * A marks card's label as lines of text segments: its custom label (Label → Text), else one line per Text
 * field; each field token replaced by valueOf(ref). A token split over runs still resolves, each segment
 * keeps the font of the run it starts in.
 * @param {any} pane @param {(ref: FieldRef) => string} valueOf
 * @returns {{ text: string, props: Record<string, any> }[][]}
 */
function tvPaneLabel(pane, valueOf) {
  const runs = pane.labelRuns.length ? pane.labelRuns.map(r => ({ text: String(r.text), props: r.props || {} }))
    : pane.encodings.filter(e => e.channel === "text" || e.channel === "label")
        .flatMap((e, i) => [...(i ? [{ text: "\n", props: {} }] : []), { text: `<${e.field.raw}>`, props: {} }]);
  const chars = runs.flatMap(r => [...r.text].map(ch => ({ ch, props: r.props })));
  const text = chars.map(c => c.ch).join("");
  /** @type {{ text: string, props: Record<string, any> }[][]} */
  const lines = [[]];
  const push = (t, props) => {
    if (!t) return;
    const line = lines[lines.length - 1], last = line[line.length - 1];
    if (last && last.props === props) last.text += t; else line.push({ text: t, props });
  };
  const re = /<([^<>]+)>|Æ[ \t]*(?:\r?\n)?|\r?\n/g;
  let at = 0;
  for (let m = re.exec(text); m; m = re.exec(text)) {
    for (let k = at; k < m.index; k++) push(chars[k].ch, chars[k].props);
    if (m[1] !== undefined) { const ref = tfParseFieldRef(m[1]); push(ref ? valueOf(ref) : m[0], chars[m.index].props); }
    else lines.push([]);
    at = m.index + m[0].length;
  }
  for (let k = at; k < text.length; k++) push(chars[k].ch, chars[k].props);
  return lines;
}

/* ── category axis: distinct label tuples in view order ─────────────────── */
/**
 * @param {ViewModel} vm
 * @param {number[]} catCis category columns, outer → inner
 * @param {boolean} sortNative sort by value (continuous / date axes) instead of view order
 * @returns {{ count: number, levels: string[][], indexOf: (row: DataValue[]) => number | undefined }}
 */
function tvCategories(vm, catCis, sortNative) {
  let order = [];
  const seen = new Map();
  vm.rows.forEach(row => {
    const labels = catCis.map(ci => tvText(row[ci]));
    const key = labels.join("\u0001");
    if (!seen.has(key)) { seen.set(key, true); order.push({ key, labels, dvs: catCis.map(ci => row[ci]) }); }
  });
  if (sortNative) {
    order.sort((a, b) => {
      for (let l = 0; l < catCis.length; l++) { const d = tfNaturalCompare(a.dvs[l], b.dvs[l]); if (d) return d; }
      return 0;
    });
  }
  if (!catCis.length) order = [{ key: "", labels: [""] }];
  const index = new Map(order.map((o, i) => [o.key, i]));
  // shown labels follow each header's date format (J F M …); the keys stay the raw values, so January,
  // June and July remain three categories even when all are labelled "J"
  const shown = catCis.map(ci => tvLabelFormatter(vm, ci));
  return {
    count: order.length,
    levels: (catCis.length ? catCis : [null]).map((_, l) => order.map(o => catCis.length ? shown[l](o.labels[l], o.dvs[l]) : o.labels[l])),
    dvs: order.map(o => o.dvs || []),
    indexOf: row => index.get(catCis.map(ci => tvText(row[ci])).join("\u0001"))
  };
}

/**
 * A continuous date on the category axis is a time axis in Tableau (ticks at whole years, quarters, months …): the
 * categories' dates as Excel serials and the step between marks, or null when a category is not a date.
 * @param {{ dvs?: DataValue[][] }} cats @returns {{ serials: number[], base: "days" | "months" | "years" } | null}
 */
function tvDateAxis(cats) {
  const dvs = (cats.dvs || []).map(d => d[0]);
  const isDate = dv => { const v = dv && (dv.nativeValue !== undefined ? dv.nativeValue : dv.value); return v instanceof Date || typeof v === "string"; };
  if (dvs.length < 2 || !dvs.every(isDate)) return null;
  // Tableau's dates arrive at midnight, sometimes in local time: the nearest whole day
  const serials = dvs.map(dv => { const s = tvExcelSerial(dv); return s === null ? null : Math.round(s); });
  if (serials.some(s => s === null)) return null;
  const gaps = serials.slice(1).map((s, i) => s - /** @type {number} */ (serials[i]));
  if (gaps.some(g => g <= 0)) return null;
  const step = Math.min(...gaps);
  return { serials: /** @type {number[]} */ (serials), base: step >= 365 ? "years" : step >= 28 ? "months" : "days" };
}

/**
 * @param {ViewModel} vm
 * @param {{ count: number, indexOf: (row: DataValue[]) => number | undefined }} cats
 * @param {number} valueCi
 * @param {(row: DataValue[]) => boolean} [filter]
 * @returns {(number | null)[]} one total per category, null = no value; .src: the measure column and each total's marks
 *   (a slicer turns the total into a formula over them – an array made from these values no longer has it)
 */
function tvSum(vm, cats, valueCi, filter) {
  const out = new Array(cats.count).fill(null);
  const marks = out.map(() => []);
  vm.rows.forEach(row => {
    if (filter && !filter(row)) return;
    const n = tfDvNum(row[valueCi]);
    if (n === null) return;
    const i = cats.indexOf(row);
    if (i === undefined) return;
    out[i] = (out[i] || 0) + n;
    marks[i].push(row);
  });
  Object.defineProperty(out, "src", { value: { ci: valueCi, rows: marks } });
  return out;
}

/** tvSum's totals picked / reordered by index (-1: no value), keeping their marks (.src)
 * @param {(number | null)[]} from @param {number[]} at the index in `from` of each value */
function tvPickSum(from, at) {
  const out = at.map(i => (i < 0 ? null : from[i]));
  const src = /** @type {any} */ (from).src;
  if (src) Object.defineProperty(out, "src", { value: { ci: src.ci, rows: at.map(i => (i < 0 ? [] : src.rows[i])) } });
  return out;
}

/** one value per mark (a scatter's x / y): measure ci of each row, with its mark (.src) @param {DataValue[][]} rows @param {number} ci */
function tvMarkValues(rows, ci) {
  const out = rows.map(r => tfDvNum(r[ci]));
  Object.defineProperty(out, "src", { value: { ci, rows: rows.map(r => [r]) } });
  return out;
}

/* ══════════════════════════════════════════════════════════════════════════
 * chart builders
 * ══════════════════════════════════════════════════════════════════════════ */
/** @param {ViewModel} vm @returns {Pick<ChartSpec, "font" | "background" | "gridlines">} */
function tvBaseSpec(vm) {
  return {
    font: tvChartFont(vm.fmt),
    // Worksheet shading: its colour, none (see-through: the container shows) or Tableau's white
    background: vm.fmt.sheetShading() === "none" ? null : tvHex(vm.fmt.tableBackground()) || "FFFFFF",
    gridlines: true
  };
}

/** @param {string} mark @returns {Pick<ChartSeries, "type" | "line" | "marker">} */
function tvSeriesType(mark) {
  if (mark === "line") return { type: "line" };
  if (mark === "area") return { type: "area" };
  if (/^(circle|shape|square)$/.test(mark)) return { type: "line", line: false, marker: true };
  return { type: "bar" };
}


/* ═══ charts/model/reflines.js ════════════════════════════════════════════════════════════════════ */
/* Analytics → Reference Line on a chart spec, valued, labelled and styled like Tableau. */

const COMPUTATION = { average: "Average", mean: "Average", median: "Median", sum: "Sum", total: "Total",
                      min: "Minimum", max: "Maximum", constant: "Constant" };

/* Excel number-format literal: "text" (quotes inside dropped) */
const lit = s => s ? `"${String(s).replace(/"/g, "")}"` : "";

/** the line's value over the marks of the chart; null = not computable
 * @param {ReferenceLine} rl @param {number[]} values @returns {number | null} */
function tvReferenceValue(rl, values) {
  if (rl.formula === "constant") return typeof rl.value === "number" ? rl.value : null;
  if (!values.length) return null;
  switch (rl.formula) {
    case "average": case "mean": return values.reduce((a, b) => a + b, 0) / values.length;
    case "median": {
      const s = [...values].sort((a, b) => a - b), m = Math.floor(s.length / 2);
      return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
    }
    case "sum": case "total": return values.reduce((a, b) => a + b, 0);
    case "min": return Math.min(...values);
    case "max": return Math.max(...values);
    default: return null;
  }
}

/**
 * The label as an Excel number format on the line's own value, so Excel prints it: "Average"
 * (automatic / computation), the value, or a custom text with <Value> / <Computation>.
 * @param {ReferenceLine} rl @param {string} valueFmt @returns {string | null}
 */
function tvReferenceLabelFormat(rl, valueFmt) {
  const name = COMPUTATION[rl.formula] || "";
  const positive = (String(valueFmt || "General").match(/^((?:"[^"]*"|[^;])*)/) || [])[1] || "General";
  if (rl.labelType === "none") return null;
  if (rl.labelType === "value") return positive;
  if (rl.labelType === "custom" && rl.label) {
    const parts = rl.label.split(/(<Value>|<Computation>)/i);
    if (!parts.some(p => /^<value>$/i.test(p))) return lit(parts.map(p => /^<computation>$/i.test(p) ? name : p.replace(/<[^<>]*>/g, "")).join(""));
    return parts.map(p => /^<value>$/i.test(p) ? positive : /^<computation>$/i.test(p) ? lit(name) : lit(p.replace(/<[^<>]*>/g, ""))).join("");
  }
  // automatic / computation: Tableau labels the line with the computation's name
  return rl.formula === "constant" ? positive : lit(name);
}

/* AGGREGATE's function per computation; option 6 skips errors – a mark the slicers leave empty is #N/A. In the file
 * it is _xlfn.AGGREGATE (an Excel 2010 function): bare, Excel reads an unknown function and shows #NAME? */
const REFLINE_AGGREGATE = { average: 1, mean: 1, max: 4, min: 5, sum: 9, total: 9, median: 12 };

/**
 * A reference line over series that follow the slicers: each of its cells computed from those series' cells
 * (writeChartData's cellFormula), so the line and its label move with them. null (Tableau's value stays) when the
 * series keep fixed values or the computation is a constant.
 * @param {ChartSpec} spec @param {ReferenceLine} rl @param {any[]} data the series the line is computed over
 */
function refLineFormula(spec, rl, data) {
  const fn = REFLINE_AGGREGATE[rl.formula];
  if (!fn || !data.length) return null;
  return (i, at) => {
    const follows = data.some(s => s.cellFormula) || (spec.valueFormula && data.some(s => s.values && s.values.src));
    const n = spec.categories.levels[0].length, ks = data.map(s => spec.series.indexOf(s));
    if (!follows || !n || ks.some(k => k < 0)) return null;
    return `_xlfn.AGGREGATE(${fn},6,${ks.map(k => `${at("series", k, 0)}:${at("series", k, n - 1)}`).join(",")})`;
  };
}

/**
 * Adds the worksheet's reference lines (table / pane scope) on the value axis: a flat line series on a
 * vertical value axis, or spec.refLines for horizontal bars (the writer draws those across the bars).
 * Lines per cell are Tableau's per-mark ticks – not drawn. Across horizontal bars on two axes a line is left
 * out and named in the conversion report: the writer would need a third axis group, and Excel cannot open that.
 * @param {ChartSpec} spec @param {ChartContext} ctx
 */
function tvApplyReferenceLines(spec, ctx) {
  const { vm, roles } = ctx;
  const fmt = vm.fmt;
  if (!fmt.hasModel || !fmt.referenceLines || spec.boxPlot || !/^(bar|line|area|combo)$/.test(spec.kind) || !spec.categories) return;
  const lines = fmt.referenceLines().filter(r => !r.scope || r.scope === "per-table" || r.scope === "per-pane");
  const valueShelf = roles.rows.values.length ? "rows" : roles.cols.values.length ? "cols" : null;
  if (!lines.length || !valueShelf) return;
  const measures = roles[valueShelf].values.map(v => v.ref).filter(Boolean);
  const drawn = spec.series.filter(s => !s.refLine && s.color !== null);                   // not invisible helper series
  const numbers = list => list.flatMap(s => s.values || []).filter(v => typeof v === "number" && isFinite(v));
  /* the series a line is computed over: its measure's (value-column, else the axis it sits on) –
     a dual-axis chart has one per measure; without a match, the primary axis's series */
  const seriesOf = rl => {
    const ref = rl.field || rl.axis;
    const m = ref && roles[valueShelf].values.find(v => v.ref && tfSameField(v.ref, ref));
    if (m) {
      const label = tvMeasureLabel(vm, m.ci).trim();
      const own = drawn.filter(s => { const n = s.name.trim(); return n === label || n.startsWith(label + " – "); });
      if (own.length) return own;
    }
    return drawn.filter(s => !s.secondary);
  };
  const style = fmt.reflineStyle();
  const hidden = style.lineVisible === false || style.strokeSize === 0;
  const valueFmt = (style.numFmtRaw && tableauToExcelNumFmt(style.numFmtRaw)) || spec.numFmt || "General";
  const twoAxes = spec.series.some(s => s.secondary);
  let left = 0;
  lines.forEach(rl => {
    if (rl.axis && measures.length && !measures.some(m => tfSameField(m, rl.axis))) return;   // another measure's axis
    const data = seriesOf(rl);
    const value = tvReferenceValue(rl, numbers(data));
    if (value === null) return;
    const labelFmt = tvReferenceLabelFormat(rl, valueFmt);
    if (hidden && !labelFmt) return;
    /** @type {RefLineStyle} */
    const line = { value, labelFmt, name: labelFmt ? tfFormatNumber(value, labelFmt) : "Reference line",
                   color: style.strokeColor ? style.strokeColor.slice(2) : "7F7F7F", alpha: style.strokeAlpha ?? 1,
                   width: style.strokeSize || 1, dash: style.dash === "dashed", hidden,
                   font: { color: style.color ? style.color.slice(2) : undefined, bold: style.bold } };
    if (spec.barDir === "bar" && data.some(s => (s.type || spec.kind) === "bar")) {
      if (twoAxes) left++; else (spec.refLines = spec.refLines || []).push(line);
    }
    // drawn on its measure's axis: on a dual-axis chart that may be the secondary one
    else spec.series.push({ name: line.name, type: "line", color: line.color, line: !hidden,
                            ...(data.length && data.every(s => s.secondary) ? { secondary: true } : {}),
                            marker: false, labels: false, values: spec.categories.levels[0].map(() => value), refLine: line,
                            cellFormula: refLineFormula(spec, rl, data) });
  });
  if (left) {
    const c = spec.conversion;
    const note = `reference line${left > 1 ? "s" : ""} left out: Excel cannot draw ${left > 1 ? "them" : "one"} across horizontal bars on two axes`;
    spec.conversion = { ...c, strategy: "APPROXIMATE", note: c && c.note ? `${c.note}; ${note}` : note };
  }
}


/* ═══ visual/classify.js ══════════════════════════════════════════════════════════════════════════ */
/* Visual type classification from marks, shelves and encodings. */

function isKPIViewModel(vm) {
  const vis = vm.order.map(i => vm.cols[i]);
  return vm.rows.length === 1 && vis.length >= 1 && vis.length <= 8 && !vis.some(c => c.isHeader);
}

function normalizeVisualToken(value) {
  return String(value == null ? "" : value).toLowerCase().replace(/[\s_-]+/g, "");
}

/* mark types of every marks card in the live visual specification
 * (marksSpecifications[].type: "bar" | "line" | "gantt-bar" | …) */
function visualSpecMarkTokens(spec) {
  if (!spec || typeof spec !== "object") return [];
  const tokens = [];
  const add = value => {
    if (value == null) return;
    if (Array.isArray(value)) value.forEach(add);
    else if (typeof value === "object") {
      [value.markType, value.markClass, value.mark, value.marksType, value.type, value.class].forEach(add);
    } else tokens.push(normalizeVisualToken(value));
  };
  if (Array.isArray(spec.marksSpecifications) && spec.marksSpecifications.length) {
    spec.marksSpecifications.forEach(add);                    // all cards → dual axis / combos are visible
  } else {
    [spec.markType, spec.markClass, spec.mark, spec.type, spec.visualType, spec.marksType, spec.marks].forEach(add);
    if (spec.activeMarksSpecification && typeof spec.activeMarksSpecification === "object") add(spec.activeMarksSpecification);
    (spec.panes || []).forEach(add);
  }
  return [...new Set(tokens.filter(t => t && t !== "automatic"))];
}

/* continuous MEASURE (draws a value axis) vs continuous date/dimension (draws a category axis) */
const TV_DATE_DERIVS = new Set(["yr", "qr", "mn", "wk", "dy", "hr", "mi", "sc", "my", "md", "mdy", "wd",
                                "tyr", "tqr", "tmn", "twk", "tdy", "thr", "tmi", "tsc"]);

function tvIsMeasureRef(model, ref) {
  if (!ref) return false;
  if (ref.name === "Multiple Values") return true;
  if (ref.type !== "qk") return false;
  if (TV_DATE_DERIVS.has(String(ref.deriv || "").toLowerCase())) return false;
  const info = tfFieldInfo(model, ref) || {};
  if (info.bin || /\(bin\)$/i.test(ref.name)) return false;           // a bin draws the histogram's category axis
  return !(info.role === "dimension" && /^date/i.test(info.datatype || ""));
}

/** a measure that is one number (MIN(0), AVG(1)): it only places marks or labels @param {FormatModel | null} model @param {FieldRef} ref */
function tvIsConstantRef(model, ref) {
  const info = ref && tfFieldInfo(model, ref);
  return !!info && info.constant !== undefined && !info.param;
}

/* what the TWB shelves say about the view: value axes, continuous dimension axes, geography,
 * and the evidence for treemaps, histograms, waterfalls and box plots */
function visualShelfShape(sheet, model) {
  if (!sheet) return null;
  const measures = shelf => sheet[shelf].filter(r => tvIsMeasureRef(model, r)).length;
  const contDim = shelf => sheet[shelf].some(r => r.type === "qk" && !tvIsMeasureRef(model, r));
  const all = [...sheet.rows, ...sheet.cols];
  const encoded = channel => sheet.panes.some(p => p.encodings.some(e => e.channel === channel));
  const datatype = r => String((tfFieldInfo(model, r) || {}).datatype || "");
  return {
    rowMeasures: measures("rows"), colMeasures: measures("cols"),
    continuousDimension: contDim("rows") || contDim("cols"),
    geo: all.some(r => /^(latitude|longitude)( \(generated\))?$/i.test(r.name)),
    filled: encoded("geometry"),
    shelfFields: all.filter(r => r.name !== "Measure Names").length,
    size: encoded("size"),
    binned: all.some(r => /\(bin\)$/i.test(r.name) || !!(tfFieldInfo(model, r) || {}).bin),
    countAxis: all.some(r => tvIsMeasureRef(model, r) && /^(cnt|ctd)$/i.test(r.deriv || "")),
    numericDiscreteDim: all.some(r => r.type === "ok" && !TV_DATE_DERIVS.has(String(r.deriv || "").toLowerCase()) &&
                                      /^(integer|real)$/i.test(datatype(r))),
    runningTotal: all.some(r => /^(cum|rsum)$/i.test(r.deriv || "") ||
                                (sheet.runningTotals || []).some(t => tfSameField(t, r))),
    boxPlot: !!sheet.boxPlot,
    // a measure on Color and nothing on Text / Label: Tableau's Automatic mark is a square (heat map)
    colorMeasure: sheet.panes.some(p => p.encodings.some(e => e.channel === "color" && tvIsMeasureRef(model, e.field))),
    text: sheet.panes.some(p => p.labelRuns.length || p.encodings.some(e => e.channel === "text" || e.channel === "label")),
    // a date on the shelves (discrete MONTH(…) or continuous) → Tableau's Automatic mark is a line
    dateDimension: all.some(r => !tvIsMeasureRef(model, r) && (TV_DATE_DERIVS.has(String(r.deriv || "").toLowerCase()) ||
                                                              (/^date/i.test(datatype(r)) && r.type !== "nk"))),
    shelfDims: all.filter(r => r.name !== "Measure Names" && !tvIsMeasureRef(model, r)).length
  };
}

/* the same evidence from the live visual specification (no workbook loaded) */
function liveSpecShape(spec) {
  if (!spec || typeof spec !== "object" || !(Array.isArray(spec.rowFields) || Array.isArray(spec.columnFields))) return null;
  const name = f => String(f == null ? "" : typeof f === "string" ? f : (f.name || f.fieldName || f.caption || ""));
  const shelf = [...(spec.rowFields || []), ...(spec.columnFields || [])].filter(f => !/^measure names$/i.test(name(f)));
  const encodings = (spec.marksSpecifications || []).flatMap(m => (m && m.encodings) || []);
  return {
    shelfFields: shelf.length,
    size: encodings.some(e => /^size$/i.test(String(e.type || e.encodingType || ""))),
    binned: shelf.some(f => /\(bin\)$/i.test(name(f)))
  };
}

/** Tableau's Automatic mark for the view's shelves: what an Automatic marks card draws
 * @param {ReturnType<typeof visualShelfShape>} shape @returns {string} */
function tvAutomaticMark(shape) {
  return !shape ? "" : shape.geo ? "map" : shape.rowMeasures && shape.colMeasures ? "circle"
    : shape.rowMeasures || shape.colMeasures ? (shape.binned ? "bar" : shape.continuousDimension || shape.dateDimension ? "line" : "bar")
    : shape.shelfFields === 0 && shape.size ? "square"                    // empty shelves + Size → treemap
    : shape.colorMeasure && !shape.text ? "square"                         // dimensions only, a measure on Color → heat map
    : "text";
}

/* mark type: live visual spec (current state) → TWB pane marks → Tableau's "Automatic" rules */
/** @param {any} spec live visual spec @param {ViewModel} vm @param {FormatModel | null} model */
function resolveVisualMarks(spec, vm, model) {
  const sheet = vm.fmt && vm.fmt.sheetModel;
  const shape = visualShelfShape(sheet, model);
  let tokens = visualSpecMarkTokens(spec);
  let source = tokens.length ? "live" : "none";
  const automatic = () => tvAutomaticMark(shape);
  if (!tokens.length && sheet) {
    // the marks drawn: every pane with an id when the sheet has several (the id-less "All" pane draws nothing)
    const drawn = sheet.panes.length > 1 && sheet.panes.some(p => p.id) ? sheet.panes.filter(p => p.id) : sheet.panes;
    tokens = [...new Set(drawn.map(p => normalizeVisualToken(p.markClass)).filter(t => t && t !== "automatic"))];
    // Automatic panes beside text / shape panes draw Tableau's automatic mark (bars of a butterfly chart beside its
    // label column): the chart mark comes first, text last
    const auto = automatic();
    if (tokens.length && tokens.every(t => /^(text|shape|circle|square)$/.test(t)) && /^(bar|line)$/.test(auto) &&
        drawn.some(p => !normalizeVisualToken(p.markClass) || normalizeVisualToken(p.markClass) === "automatic")) tokens.push(auto);
    tokens.sort((a, b) => Number(a === "text") - Number(b === "text"));
    if (tokens.length) source = "twb";
  }
  if (!tokens.length && shape) {
    source = "automatic";
    tokens = [automatic()];
  }
  return { tokens, shape, source };
}

/**
 * Rounded bars: a line from a constant (MIN(0)) to each value, one per member, drawn thick with round caps –
 * Line mark, Measure Names on Path, the constant among Measure Values.
 * @param {ViewModel} vm @param {FormatModel | null} model
 */
function isRoundedBar(vm, model) {
  const sheet = vm.fmt && vm.fmt.sheetModel;
  if (!sheet) return false;
  const path = sheet.panes.some(p => /^line$/i.test(p.markClass || "") && p.encodings.some(e => e.channel === "path" && e.field.name === "Measure Names"));
  const mv = [...sheet.rows, ...sheet.cols].some(r => r.name === "Multiple Values");
  const circles = sheet.panes.some(p => /^(circle|shape)$/i.test(p.markClass || ""));
  return path && mv && !circles && sheet.fieldRefs.some(r => tvIsConstantRef(model, r));
}

/**
 * A table built from marks: every axis on Rows / Columns is a constant ("MIN(0)", "0.5") that only places
 * the marks – one column of circles, squares, shapes or bars per axis, each pane showing its label.
 * @param {ViewModel} vm @param {FormatModel | null} model
 */
function isMarkTable(vm, model) {
  const sheet = vm.fmt && vm.fmt.sheetModel;
  if (!sheet) return false;
  const axes = [...sheet.rows, ...sheet.cols].filter(r => tvIsMeasureRef(model, r));
  if (!axes.length || !sheet.panes.some(p => p.labelRuns.length || p.encodings.some(e => e.channel === "text" || e.channel === "label"))) return false;
  // pies / donuts on MIN(0) axes are charts: only cell-like marks make a table
  if (!sheet.panes.every(p => /^(circle|square|shape|bar|text|automatic)?$/i.test(p.markClass || ""))) return false;
  return axes.every(r => {
    const info = tfFieldInfo(model, r);
    if (info && info.constant !== undefined) return true;
    // no formula to go by: the same value on every row of the data
    const ci = vm.cols.findIndex(c => c.ref && tfSameField(c.ref, r));
    if (ci < 0 || vm.rows.length < 2) return false;
    return new Set(vm.rows.map(row => tfDvNum(row[ci])).filter(v => v !== null)).size <= 1;
  });
}

/**
 * @param {any} spec live visual spec @param {ViewModel} vm @param {FormatModel | null} model
 * @param {string[]} [ev] collects the evidence behind the decision (shown in the conversion report)
 * @returns {VisualType}
 */
function classifyVisualType(spec, vm, model, ev = []) {
  const { tokens, shape, source } = resolveVisualMarks(spec, vm, model);
  const live = liveSpecShape(spec);
  const axes = shape ? shape.rowMeasures + shape.colMeasures : null;     // null → unknown (no workbook)
  const token = tokens[0] || "";
  ev.push(tokens.length ? `${tokens.join(" + ")} mark${tokens.length > 1 ? "s" : ""} (${source === "twb" ? "workbook" : source === "live" ? "live view" : "Tableau's automatic mark"})` : "no mark type known");
  if (shape) ev.push(`${shape.rowMeasures} measure axis on Rows, ${shape.colMeasures} on Columns`);
  const kpiOrTable = why => {
    const kpi = isKPIViewModel(vm);
    ev.push(why, kpi ? "one row of a few values: a KPI" : "drawn as cells");
    return kpi ? VISUAL_TYPES.KPI : VISUAL_TYPES.TABLE;
  };
  // generated lat/long on the shelves: a map whatever the mark (circle → symbol map)
  if (shape && shape.geo) {
    ev.push("generated latitude / longitude on the shelves: a map");
    const filled = shape.filled || tokens.some(t => t === "multipolygon" || t === "polygon");
    if (filled) ev.push(shape.filled ? "geometry encoding: filled areas" : "polygon marks");
    return filled ? VISUAL_TYPES.MAP_FILLED : VISUAL_TYPES.MAP;
  }
  if (axes && isMarkTable(vm, model)) return kpiOrTable("every axis is a constant that only places the marks: a table built from marks");
  if (tokens.includes("text") && tokens.every(t => /^(text|shape|circle|square)$/.test(t))) {
    return kpiOrTable("text marks");                                  // KPI tiles / buttons built from marks
  }
  const cartesian = tokens.filter(t => t === "bar" || t === "line" || t === "area");
  if (new Set(cartesian).size > 1) { ev.push("different marks per axis: a combination chart"); return VISUAL_TYPES.COMBO; }
  if (token === "line" && isRoundedBar(vm, model)) {
    ev.push("a thick line from a constant to each value (Path: Measure Names): rounded bars");
    return VISUAL_TYPES.BAR;
  }
  if (token === "ganttbar" || token === "gantt") {
    if (shape && shape.runningTotal) { ev.push("Gantt bars on a running total: a waterfall"); return VISUAL_TYPES.WATERFALL; }
    ev.push("Gantt bars: start on the axis, length from Size");
    return VISUAL_TYPES.GANTT;
  }
  if (shape && shape.boxPlot && axes > 0) { ev.push("box plot reference distribution (Analytics pane)"); return VISUAL_TYPES.BOXPLOT; }
  const byMark = {
    bar: VISUAL_TYPES.BAR, line: VISUAL_TYPES.LINE, area: VISUAL_TYPES.AREA, pie: VISUAL_TYPES.PIE,
    map: VISUAL_TYPES.MAP, multipolygon: VISUAL_TYPES.MAP_FILLED, polygon: VISUAL_TYPES.MAP_FILLED,
    heatmap: VISUAL_TYPES.MAP, density: VISUAL_TYPES.MAP,               // density marks: Tableau-only visual
    vizextension: VISUAL_TYPES.CUSTOM
  };
  if (token === "bar" && axes === 0) return kpiOrTable("bar marks without a measure axis");
  if (token === "bar" && ((shape && (shape.binned || (shape.countAxis && shape.numericDiscreteDim))) || (live && live.binned))) {
    ev.push(shape && shape.binned || live && live.binned ? "a bin field on the category axis: a histogram" : "counts per numeric value: a histogram");
    return VISUAL_TYPES.HISTOGRAM;
  }
  if (byMark[token]) {
    ev.push(token === "vizextension" ? "a viz extension draws this sheet" : /^(heatmap|density)$/.test(token) ? "density marks" : `${token} marks`);
    return byMark[token];
  }
  if (token === "circle" || token === "shape" || token === "square") {
    // nothing on Rows/Columns + Size → Tableau lays the marks out itself: treemap / packed bubbles
    const free = shape ? shape.shelfFields === 0 && shape.size : live ? live.shelfFields === 0 && live.size : false;
    if (free) {
      ev.push("nothing on Rows / Columns and a Size measure: Tableau packs the marks itself");
      return token === "square" ? VISUAL_TYPES.TREEMAP : VISUAL_TYPES.BUBBLE;
    }
    if (axes === null) { ev.push("no workbook: guessed from the mark"); return token === "circle" ? VISUAL_TYPES.SCATTER : VISUAL_TYPES.TABLE; }
    if (shape.rowMeasures && shape.colMeasures) { ev.push("a measure on both axes: a scatter plot"); return VISUAL_TYPES.SCATTER; }
    if (axes > 0) { ev.push("marks along one measure axis: a dot plot"); return VISUAL_TYPES.LINE; }
    if (token === "square") { ev.push("square marks in a grid of dimensions: a heat map"); return VISUAL_TYPES.HEATMAP; }
    return kpiOrTable(`${token} marks in a grid of dimensions`);
  }
  if (!token && vm.kind === "chart") { ev.push("chart without a known mark"); return VISUAL_TYPES.UNKNOWN; }
  return kpiOrTable(token ? `${token} marks` : "no chart marks");
}


/* ═══ charts/model/cartesian.js ═══════════════════════════════════════════════════════════════════ */
/* Bar / line / area / combo / funnel / histogram chart specs. */

/** a constant measure's value (MIN(0) → 0, AVG(1) → 1), else undefined @param {ChartContext} ctx @param {FieldRef | null} ref */
const constantOf = (ctx, ref) => ref && tvIsConstantRef(ctx.model || null, ref) ? tfFieldInfo(ctx.model, ref).constant : undefined;
/** a "% of Total" quick table calculation @param {FieldRef | null} ref */
const isPercentOfTotal = ref => !!ref && (/^pcto$/i.test(ref.deriv || "") || /^pcto:/i.test(ref.inner || ""));
/** a measure its formula negates ("-COUNT([Customers])"): the left wing of a butterfly chart */
const isNegated = (ctx, ref) => { const info = ref && ctx.model ? tfFieldInfo(ctx.model, ref) : null; return !!info && /^\s*-/.test(info.formula || ""); };
/** a number format that shows negative values without their minus sign (both wings of a butterfly read positive) */
const absFmt = fmt => { const p = (String(fmt || "General").match(/^((?:"[^"]*"|[^;])*)/) || [])[1] || "General"; return `${p};${p}`; };

/** @param {ChartContext} ctx @returns {ChartSpec[]} */
function tvCartesianSpecs(ctx) {
  const { vm, roles } = ctx;
  const valueShelf = roles.rows.values.length ? "rows" : roles.cols.values.length ? "cols" : null;
  if (!valueShelf) throw new Error("no continuous measure axis to chart");
  const catShelf = valueShelf === "rows" ? "cols" : "rows";
  const horizontal = valueShelf === "cols";
  // a constant 0 on an axis (MIN(0), AVG(0)) only places labels – a butterfly's names – or the start of rounded bars:
  // not data, so not a series
  const zero = m => constantOf(ctx, m.ref) === 0;
  if (roles[valueShelf].values.some(zero) && roles[valueShelf].values.some(m => !zero(m))) {
    const values = roles[valueShelf].values.filter(m => !zero(m));
    return tvCartesianSpecs({ ...ctx, roles: { ...roles, [valueShelf]: { ...roles[valueShelf], values } } });
  }
  const measures = roles[valueShelf].values;
  const mvMode = measures.some(m => m.mv);
  const catDims = [...roles[valueShelf].dims, ...roles[catShelf].dims];
  const color = roles.color;
  const colorCi = color && !color.measureNames && !color.continuous ? color.ci : -1;
  const colorLevel = colorCi >= 0 ? catDims.findIndex(d => d.ci === colorCi) : -1;
  // a pane of small multiples gets the whole view's colours (its own rows hold one member of each)
  const scale = ctx.colorScale !== undefined ? ctx.colorScale : tvColorScale(vm, roles, ctx.markToken);
  const fieldLabels = vm.fmt.hasModel ? vm.fmt.fieldLabelsShown(catShelf) : true;
  // rounded bars are lines in Tableau (thick, round caps): Excel's bars
  const asBars = t => ctx.roundedBar && t === "line" ? "bar" : t;
  const sheetMark = asBars(ctx.markToken);
  // an Automatic marks card draws the All card's mark, else Tableau's automatic mark for the shelves (a line over
  // dates beside a card of bars – not the bars of the other card)
  const allPane = roles.panes.find(p => !p.id);
  const inherit = (allPane && asBars(tvMarkToken(allPane.markClass))) || asBars((roles.source === "twb" && ctx.autoMark) || ctx.markToken);
  /** @param {FieldRef | null} ref @param {string} fallback @param {number} [index] @param {number} [count] */
  const measureMark = (ref, fallback, index, count) => {
    const own = ref && roles.panes.find(x => x.id && [x.yAxisName, x.xAxisName].some(n => n && tfSameField(tfParseFieldRef(n), ref)));
    if (own && !tvMarkToken(own.markClass)) return inherit;
    return asBars(tvMeasureMark(roles, ref, fallback, index, count));
  };
  // the colour's field is also the innermost category, its header hidden (Department, then Race – coloured, with no
  // header of its own): Tableau's clustered bars. One series per member side by side, with a colour legend, and the
  // outer field alone on the axis – each name under its group, where a level of blank labels would squeeze it
  const lastDim = roles[catShelf].dims[roles[catShelf].dims.length - 1];
  if (!ctx.clustered && colorCi >= 0 && lastDim && lastDim.ci === colorCi && !lastDim.continuous && catDims.length > 1 &&
      measures.length === 1 && vm.fmt.hasModel && lastDim.ref && vm.fmt.isLabelHidden(lastDim.ref) &&
      measureMark(measures[0].ref, inherit || sheetMark) === "bar") {
    return tvCartesianSpecs({ ...ctx, clustered: true, roles: { ...roles, [catShelf]: { ...roles[catShelf], dims: roles[catShelf].dims.slice(0, -1) } } });
  }

  // categories (Measure Names alone on the category shelf → one bar per measure)
  const mnCategory = mvMode && !catDims.length && roles.measureNames === catShelf;
  const cats = mnCategory
    ? { count: measures.length, levels: [measures.map(m => tvMeasureLabel(vm, m.ci))], indexOf: () => 0 }
    : tvCategories(vm, catDims.map(d => d.ci), catDims.some(d => d.continuous));
  if (cats.count > TV_MAX_POINTS) throw new Error(`${cats.count} categories – too many for an Excel chart`);
  const categoryTitle = fieldLabels && catDims.length ? catDims.map(d => tvMeasureLabel(vm, d.ci)).join(" / ") : "";
  const categories = { names: catDims.length ? catDims.map(d => tvMeasureLabel(vm, d.ci)) : [mnCategory ? "Measure Names" : ""], levels: cats.levels };
  // one continuous date across the chart: an Excel date axis, labelled at whole years / months like Tableau's
  const dateAxis = !horizontal && !mnCategory && catDims.length === 1 && catDims[0].continuous ? tvDateAxis(cats) : null;

  // per-category colours: colour = a category level, or a continuous measure
  let pointColors;
  if (scale && colorLevel >= 0) pointColors = cats.levels[colorLevel].map(v => scale(v));
  else if (scale && color && color.continuous && color.ci >= 0) pointColors = tvSum(vm, cats, color.ci).map(v => scale(v));

  // a dimension on Detail / Label that splits the marks of one category and colour: Tableau draws a line (or a dot)
  // per member – each becomes its own series, never their sum
  const splitDims = (() => {
    if (mnCategory) return [];
    const cand = [...roles.detailDims, ...roles.textDims].filter((ci, i, a) => ci >= 0 && a.indexOf(ci) === i && ci !== colorCi &&
                                                                    !catDims.some(d => d.ci === ci));
    if (!cand.length) return [];
    const seen = new Map();
    for (const r of vm.rows) {
      const k = cats.indexOf(r) + "\u0001" + (colorCi >= 0 ? tvText(r[colorCi]) : "");
      const v = cand.map(ci => tvText(r[ci])).join("\u0001");
      if (!seen.has(k)) seen.set(k, v); else if (seen.get(k) !== v) return cand;
    }
    return [];
  })();

  const seriesFor = (m, opts) => {
    const label = tvMeasureLabel(vm, m.ci);
    const labels = tvLabelsOn(roles, m.ref);
    if (splitDims.length && opts.type && opts.type.type === "line") {
      const keyCis = [...(colorCi >= 0 && colorLevel < 0 && !opts.noColorSplit ? [colorCi] : []), ...splitDims];
      const keyOf = r => keyCis.map(ci => tvText(r[ci])).join(" – ");
      const keys = [...new Set(vm.rows.map(keyOf))];
      if (keys.length > TV_MAX_SERIES) {
        throw new Error(`${keys.length} lines (one per ${splitDims.map(ci => tvMeasureLabel(vm, ci)).join(" / ")}) – too many for an Excel chart`);
      }
      // each line in its member's colour (the colour field's value on its marks), else the marks' colour
      const colourOf = k => { const r = colorCi >= 0 && scale ? vm.rows.find(x => keyOf(x) === k) : null; return r ? scale(tvText(r[colorCi])) : null; };
      return keys.map((k, i) => ({
        name: opts.prefix ? `${label} – ${k}` : k, values: tvSum(vm, cats, m.ci, r => keyOf(r) === k),
        color: colourOf(k) || opts.color || tvHex(TABLEAU_10[i % TABLEAU_10.length]), labels, ...opts.type
      }));
    }
    if (colorCi >= 0 && colorLevel < 0 && !opts.noColorSplit) {
      const values = tvColorValues(vm, roles);
      return values.map((v, i) => ({
        name: opts.prefix ? `${label} – ${v}` : v,
        values: tvSum(vm, cats, m.ci, r => tvText(r[colorCi]) === v),
        color: (scale && scale(v)) || tvHex(TABLEAU_10[i % TABLEAU_10.length]),
        labels, ...opts.type
      }));
    }
    return [{ name: label, values: tvSum(vm, cats, m.ci), color: opts.color, labels,
              pointColors: opts.noPointColors ? undefined : pointColors, ...opts.type,
              ...(opts.alpha !== undefined ? { alpha: opts.alpha } : {}) }];
  };

  // Tableau never draws a line across panes: categories of several levels (a year's quarters, a country's months) start
  // a new line at every change of an outer level – each such stretch its own series, the same colour, blank elsewhere
  const paneBreaks = series => {
    const levels = cats.levels;
    const broken = s => !s.refLine && ((s.type === "line" && s.line !== false) || s.type === "area");
    if (levels.length < 2 || !series.some(broken)) return series;
    const groupOf = i => levels.slice(0, -1).map(lv => lv[i]).join("\u0001");
    const ranges = [];
    for (let i = 0; i < cats.count; i++) {
      const g = groupOf(i), last = ranges[ranges.length - 1];
      if (last && last.g === g) last.to = i; else ranges.push({ g, from: i, to: i });
    }
    if (ranges.length < 2 || ranges.length * series.length > TV_MAX_SERIES) return series;
    return series.flatMap(s => !broken(s) ? [s]
      : ranges.map((r, k) => ({ ...s, name: k ? `${s.name} ` + "\u200b".repeat(k) : s.name,
                                values: tvPickSum(s.values, s.values.map((v, i) => i >= r.from && i <= r.to ? i : -1)),
                                pointColors: s.pointColors ? s.pointColors.map((c, i) => i >= r.from && i <= r.to ? c : null) : undefined })));
  };

  // stacked bars / areas of a "% of Total" that fills every category: Excel's 100% stacked chart
  const percentOfTotal = measures.length > 0 && measures.every(m => isPercentOfTotal(m.ref));
  const make = (series, extra) => {
    if (series.length > TV_MAX_SERIES) throw new Error(`${series.length} series – too many for an Excel chart`);
    const kinds = new Set(series.map(s => s.type));
    const kind = kinds.size > 1 || series.some(s => s.secondary) ? "combo" : ([...kinds][0] || "bar");
    const stackable = series.every(s => s.type === "bar" || s.type === "area");
    const stacked = stackable && series.length > 1 && !ctx.clustered &&
      ((colorCi >= 0 && colorLevel < 0) || (mvMode && color && color.measureNames && roles.measureNames !== catShelf));
    const totals = Array.from({ length: cats.count }, (_, i) => series.reduce((a, s) => a + ((s.values && s.values[i]) || 0), 0));
    const percent = stacked && percentOfTotal && totals.some(t => t) && totals.every(t => t === 0 || Math.abs(t - 1) < 0.005);
    // Tableau stacks columns in the colour legend's order from the top; Excel stacks its first series at the
    // bottom → reversed (bars across keep the legend's order from the axis outwards)
    const ordered = paneBreaks(stacked && !horizontal ? [...series].reverse() : series);
    const spec = { ...tvBaseSpec(vm), kind, barDir: horizontal ? "bar" : "col", stacked,
                   categories, categoryTitle, legend: series.length > 1, series: ordered, ...(dateAxis ? { dateAxis } : {}), ...extra };
    if (percent) {
      spec.percent = true;
      if (!/%/.test(spec.numFmt || "")) spec.numFmt = "0%";
    }
    return spec;
  };

  const axisIds = [...new Set(measures.map(m => m.axis || 0))].sort((a, b) => a - b);
  const paneOfAxis = k => {
    const refs = roles[valueShelf].axisRefs || [];
    const ref = refs[k];
    if (ref) {
      const nth = refs.slice(0, k).filter(r => tfSameField(r, ref) || (r.name === ref.name && r.name === "Multiple Values")).length;
      const axisName = p => (horizontal ? p.xAxisName : p.yAxisName);
      const hits = roles.panes.filter(p => { const n = axisName(p) && tfParseFieldRef(axisName(p)); return n && (tfSameField(n, ref) || (n.name === "Multiple Values" && ref.name === "Multiple Values")); });
      if (hits[nth]) return hits[nth];
    }
    return roles.panes.find(x => String(x.id) === String(k + 1)) || null;
  };
  const paneMarkOfAxis = k => { const p = paneOfAxis(k); return p ? asBars(tvMarkToken(p.markClass)) || (p.id ? inherit : "") : ""; };
  const paneColor = (m, i) => {
    const p = paneOfAxis(m.axis || 0);
    return (color && color.measureNames && scale ? scale(vm.cols[m.ci].name, i) : null) ||
           (p ? tvHex(tfArgb(tvPaneRule(p, "mark", "mark-color"))) : null) || tvHex(TABLEAU_10[i % TABLEAU_10.length]);
  };

  // an axis of custom shapes picked by a field (logos, ▲▼ icons: Shape mark + Shape encoding) is
  // decoration Excel cannot draw → chart the other measures instead of a meaningless extra series
  if (axisIds.length > 1) {
    const decorative = axisIds.filter(k => { const p = paneOfAxis(k); return p && tvMarkToken(p.markClass) === "shape" && p.encodings.some(e => e.channel === "shape"); });
    if (decorative.length && decorative.length < axisIds.length) {
      const keep = measures.filter(m => !decorative.includes(m.axis || 0));
      if (keep.length) { measures.splice(0, measures.length, ...keep); return tvCartesianSpecs({ ...ctx, roles: { ...roles, [valueShelf]: { ...roles[valueShelf], values: keep } } }); }
    }
  }

  // butterfly / population pyramid: two measures back to back – the left one negated by its formula, or drawn on a
  // reversed axis – as bars stacked around zero; the left wing's minus signs are hidden
  const uniqueMeasures = measures.filter((m, i) => measures.findIndex(x => x.ci === m.ci) === i);
  if (horizontal && uniqueMeasures.length === 2 && colorLevel < 0 && colorCi < 0) {
    const reversed = m => !m.mv && !!vm.fmt.axisSpace(m.ref, valueShelf, "0").reverse;
    const left = uniqueMeasures.findIndex(m => isNegated(ctx, m.ref) || reversed(m));
    const right = 1 - left;
    if (left >= 0 && !isNegated(ctx, uniqueMeasures[right].ref) && !reversed(uniqueMeasures[right])) {
      const flip = !isNegated(ctx, uniqueMeasures[left].ref);           // a reversed axis holds positive values
      const series = [left, right].map((k, i) => {
        const m = uniqueMeasures[k];
        const values = tvSum(vm, cats, m.ci).map(v => v === null ? null : k === left && flip ? -v : v);
        return { name: tvMeasureLabel(vm, m.ci).replace(/^AGG\(-(.*)\)$/i, "$1"), values, type: /** @type {"bar"} */ ("bar"),
                 color: paneColor(m, i), labels: tvLabelsOn(roles, m.ref), labelNumFmt: absFmt(tvNumFmt(vm, m.ci)) };
      });
      const spec = make(series, { numFmt: absFmt(tvNumFmt(vm, uniqueMeasures[right].ci)), valueTitle: "", gapWidth: 30 });
      return [{ ...spec, stacked: true, legend: true, mirrored: true,
                conversion: { strategy: "CONSTRUCTED", output: "Back-to-back bar chart (butterfly)",
                              note: "Excel has no butterfly chart: two series stacked around zero, the left one negated, its minus signs hidden" } }];
    }
  }

  // barbell (dumbbell): two values per category joined by a line (Measure Names on the line's Path) → markers for
  // both values and Excel's high-low line between them
  const pathMN = roles.panes.some(p => tvMarkToken(p.markClass) === "line" && p.encodings.some(e => e.channel === "path" && e.field.name === "Measure Names"));
  if (pathMN && mvMode && ctx.type !== VISUAL_TYPES.BAR && uniqueMeasures.length === 2 && !mnCategory && horizontal) {
    const linePane = roles.panes.find(p => tvMarkToken(p.markClass) === "line");
    const [a, b] = uniqueMeasures.map(m => tvSum(vm, cats, m.ci));
    const join = { name: "", color: (linePane && tvHex(tfArgb(tvPaneRule(linePane, "mark", "mark-color")))) || "A0A0A0", line: true, lineWidth: 2,
                   values: a.flatMap((v, i) => [v, b[i], null]), cats: a.flatMap((v, i) => v === null || b[i] === null ? [null, null, null] : [i, i, null]) };
    const dots = uniqueMeasures.map((m, i) => ({ name: tvMeasureLabel(vm, m.ci), color: paneColor(m, i), marker: "circle", markerSize: 9,
                                                 values: [a, b][i], cats: [a, b][i].map((v, c) => v === null ? null : c) }));
    const spec = make([{ name: "", values: a.map(() => null), type: "bar", color: null, labels: false }],
                      { numFmt: tvNumFmt(vm, uniqueMeasures[0].ci), valueTitle: "" });
    return [{ ...spec, legend: false, overlay: [join, ...dots],
              conversion: { strategy: "CONSTRUCTED", output: "Dumbbell: dots and joining lines over the bar chart's category bands",
                            note: "Excel has no dumbbell chart: both values as dots, joined by a line, drawn over horizontal category bands" } }];
  }
  if (pathMN && mvMode && ctx.type !== VISUAL_TYPES.BAR && uniqueMeasures.length === 2 && !mnCategory) {
    const linePane = roles.panes.find(p => tvMarkToken(p.markClass) === "line");
    const series = uniqueMeasures.map((m, i) => ({ name: tvMeasureLabel(vm, m.ci), values: tvSum(vm, cats, m.ci), type: /** @type {"line"} */ ("line"),
      line: false, marker: true, markerSize: 9, color: paneColor(m, i), labels: tvLabelsOn(roles, m.ref) }));
    const spec = make(series, { numFmt: tvNumFmt(vm, uniqueMeasures[0].ci), valueTitle: "" });
    return [{ ...spec, kind: "line", barDir: "col", stacked: false, legend: true,
              hiLowLines: { color: (linePane && tvHex(tfArgb(tvPaneRule(linePane, "mark", "mark-color")))) || "A0A0A0" },
              conversion: horizontal
                ? { strategy: "APPROXIMATE", output: "Line chart: markers joined by high-low lines (vertical)",
                    note: "drawn vertically: Excel joins markers with high-low lines only on vertical line charts" }
                : { strategy: "CONSTRUCTED", output: "Line chart: markers joined by high-low lines",
                    note: "Excel has no dumbbell chart: both values as markers, joined by high-low lines" } }];
  }

  // the same measure on both axes (styling: line + shape at the points, bar + line, Gantt cap on bars)
  // → drawn once: bar / area / line wins, a line gets markers when the other pane draws shapes
  if (!mvMode && axisIds.length > 1 && measures.every(m => m.ci === measures[0].ci)) {
    const toks = axisIds.map(k => paneMarkOfAxis(k) || sheetMark).filter(Boolean);
    if (toks.includes("bar") && toks.some(x => /^(circle|shape)$/.test(x))) {
      const m = measures[0];
      const base = { name: tvMeasureLabel(vm, m.ci), values: tvSum(vm, cats, m.ci), color: tvMarkColor(roles), pointColors };
      /** @type {ChartSeries[]} */
      const series = [{ ...base, type: "bar", labels: horizontal && tvLabelsOn(roles, m.ref) }];
      if (!horizontal) series.push({ ...base, name: base.name + " ", type: "line", line: false, marker: true, markerSize: 9, labels: tvLabelsOn(roles, m.ref) });
      const spec = make(series, { numFmt: tvNumFmt(vm, m.ci), gapWidth: 300, valueTitle: tvMeasureLabel(vm, m.ci) });
      const heads = horizontal ? [{ name: base.name, color: base.color, marker: "circle", markerSize: 9, values: base.values, cats: base.values.map((v, i) => v === null ? null : i) }] : undefined;
      return [{ ...spec, stacked: false, legend: false, overlay: heads, conversion: lollipopConversion(horizontal) }];
    }
    const primary = ["bar", "area", "line"].find(x => toks.includes(x)) || toks[0] || "bar";
    const type = { ...tvSeriesType(primary) };
    if (primary === "line" && toks.some(x => /^(circle|shape|square)$/.test(x))) type.marker = true;
    const m = measures[0];
    return [withTargets(make(seriesFor(m, { type, color: tvMarkColor(roles) }), { numFmt: tvNumFmt(vm, m.ci), valueTitle: tvMeasureLabel(vm, m.ci) }))];
  }

  // a measure on one axis and Measure Values on the other (Sales bars + MV lines) → combo, axis each
  const groupOf = k => measures.filter(m => (m.axis || 0) === k);
  const sameSets = axisIds.every(k => { const a = groupOf(k).map(m => m.ci).sort().join(), b = groupOf(axisIds[0]).map(m => m.ci).sort().join(); return a === b; });
  if (mvMode && axisIds.length > 1 && !sameSets && !mnCategory) {
    const series = [];
    axisIds.forEach((k, ai) => {
      const tok = paneMarkOfAxis(k) || sheetMark;
      groupOf(k).filter((m, i, a) => a.findIndex(x => x.ci === m.ci) === i).forEach(m => series.push({
        name: tvMeasureLabel(vm, m.ci), values: tvSum(vm, cats, m.ci), labels: tvLabelsOn(roles, m.ref),
        color: (color && color.measureNames && scale ? scale(vm.cols[m.ci].name, series.length) : null) || tvHex(TABLEAU_10[series.length % TABLEAU_10.length]),
        ...tvSeriesType(tok), secondary: axisIds.length === 2 && ai === 1
      }));
    });
    const first = groupOf(axisIds[0])[0], second = groupOf(axisIds[1])[0];
    const spec = make(series, { numFmt: tvNumFmt(vm, first.ci), valueTitle: groupOf(axisIds[0]).length === 1 ? tvMeasureLabel(vm, first.ci) : "Value",
                                secondaryNumFmt: tvNumFmt(vm, second.ci), secondaryTitle: groupOf(axisIds[1]).length === 1 ? tvMeasureLabel(vm, second.ci) : "Value" });
    spec.stacked = false;
    spec.legend = series.length > 1;
    return [spec];
  }

  // Measure Values: one chart, one series per measure
  if (mvMode) {
    // dual axis "MV + MV": every measure is listed once per axis, and each axis has its own
    // pane (id 1, 2, …) with its own mark → one entry per measure, marks per axis
    const axisCount = Math.max(...measures.map(m => (m.axis || 0) + 1));
    const axisMarks = Array.from({ length: axisCount }, (_, k) => {
      return (axisCount > 1 && paneMarkOfAxis(k)) || measureMark(null, sheetMark);
    });
    const unique = uniqueMeasures;
    const markSet = [...new Set(axisMarks.map(t => tvSeriesType(t).marker && tvSeriesType(t).line === false ? "dot" : tvSeriesType(t).type))];
    const type = tvSeriesType(axisMarks.find(t => tvSeriesType(t).type === "bar") || axisMarks[0]);

    if (!mnCategory && axisCount > 1 && markSet.length > 1) {
      if (markSet.includes("bar") && markSet.includes("dot")) {
        // lollipop: bar + circle on the same values. Vertical → thin bars + markers;
        // horizontal → Excel cannot mix bar and line orientations, so thin bars only
        const colorOf = (m, i) => (color && color.measureNames && scale ? scale(vm.cols[m.ci].name, i) : null) ||
          (unique.length === 1 ? tvMarkColor(roles) : tvHex(TABLEAU_10[i % TABLEAU_10.length]));
        const series = [];
        unique.forEach((m, i) => {
          const base = { name: tvMeasureLabel(vm, m.ci), values: tvSum(vm, cats, m.ci), color: colorOf(m, i), pointColors };
          series.push({ ...base, type: "bar", labels: false });
          if (!horizontal) series.push({ ...base, name: base.name + " ", type: "line", line: false, marker: true, markerSize: 9,
                                         labels: tvLabelsOn(roles, m.ref) });
        });
        if (horizontal) series.forEach(s => { s.labels = tvLabelsOn(roles, unique[0].ref); });
        const spec = make(series, { numFmt: tvNumFmt(vm, unique[0].ci), gapWidth: 300,
                                    valueTitle: unique.length === 1 ? tvMeasureLabel(vm, unique[0].ci) : "Value" });
        spec.stacked = false;                      // the circle sits on the bar end, never on top of it
        spec.legend = unique.length > 1;
        // across: each bar's head as a dot over its band
        if (horizontal) spec.overlay = series.map(s => ({ name: s.name, color: s.color, marker: "circle", markerSize: 9, values: s.values, cats: s.values.map((v, i) => v === null ? null : i) }));
        spec.conversion = lollipopConversion(horizontal);
        return [spec];
      }
      // e.g. area + line of the same measure: overlay both, sharing one axis
      const series = [];
      axisMarks.forEach((token, k) => unique.forEach((m, i) => series.push({
        name: tvMeasureLabel(vm, m.ci) + (k ? " " : ""), values: tvSum(vm, cats, m.ci), labels: k === axisCount - 1 && tvLabelsOn(roles, m.ref),
        color: (color && color.measureNames && scale ? scale(vm.cols[m.ci].name, i) : null) || tvHex(TABLEAU_10[i % TABLEAU_10.length]),
        ...tvSeriesType(token)
      })));
      const spec = make(series, { numFmt: tvNumFmt(vm, unique[0].ci), valueTitle: unique.length === 1 ? tvMeasureLabel(vm, unique[0].ci) : "Value" });
      spec.stacked = false;
      spec.legend = unique.length > 1;
      return [spec];
    }
    measures.splice(0, measures.length, ...unique);

    // funnel: one category per measure, each row carries exactly one of the measures
    // (per-stage calcs) → one value per stage, drawn centred
    const exclusive = catDims.length === 1 && measures.length > 1 && type.type === "bar" &&
      vm.rows.every(r => measures.filter(m => tfDvNum(r[m.ci]) !== null).length <= 1);
    if (exclusive) return tvFunnelSpecFromStages(ctx, cats, catDims[0].ci, measures, horizontal);

    if (mnCategory) {
      const values = measures.map(m => vm.rows.reduce((s, r) => s + (tfDvNum(r[m.ci]) || 0), 0));
      const pc = color && color.measureNames && scale ? measures.map((m, i) => scale(vm.cols[m.ci].name, i)) : undefined;
      return [make([{ name: "Value", values, color: tvMarkColor(roles), pointColors: pc,
                      labels: tvLabelsOn(roles, measures[0].ref), ...type }],
                   { numFmt: tvNumFmt(vm, measures[0].ci), valueTitle: "Value" })];
    }
    const series = measures.map((m, i) => ({
      name: tvMeasureLabel(vm, m.ci), values: tvSum(vm, cats, m.ci), labels: tvLabelsOn(roles, m.ref),
      color: (color && color.measureNames && scale ? scale(vm.cols[m.ci].name, i) : null) || tvHex(TABLEAU_10[i % TABLEAU_10.length]),
      ...type
    }));
    return [make(series, { numFmt: tvNumFmt(vm, measures[0].ci), valueTitle: "Value" })];
  }

  // one measure
  if (measures.length === 1) {
    const m = measures[0];
    const type = tvSeriesType(measureMark(m.ref, sheetMark));
    // a horizontal dot plot: each mark (detail included) a dot over its category band, one series per colour
    if (horizontal && type.type === "line" && type.line === false && !mnCategory) {
      const groups = colorCi >= 0 && colorLevel < 0 ? tvColorValues(vm, roles) : [null];
      if (groups.length <= TV_MAX_SERIES) {
        const overlay = groups.map((g, i) => {
          const own = vm.rows.filter(r => (g === null || tvText(r[colorCi]) === g) && tfDvNum(r[m.ci]) !== null && cats.indexOf(r) !== undefined);
          const c = g === null ? tvMarkColor(roles) : (scale && scale(g)) || tvHex(TABLEAU_10[i % TABLEAU_10.length]);
          return { name: g === null ? tvMeasureLabel(vm, m.ci) : g, color: c, marker: "circle", markerSize: 7,
                   values: own.map(r => tfDvNum(r[m.ci])), cats: own.map(r => cats.indexOf(r)) };
        });
        const spec = make([{ name: "", values: new Array(cats.count).fill(null), type: "bar", color: null, labels: false }],
                          { numFmt: tvNumFmt(vm, m.ci), valueTitle: tvMeasureLabel(vm, m.ci) });
        return [{ ...spec, legend: groups.length > 1, overlay,
                  conversion: { strategy: "CONSTRUCTED", output: "Dot plot: dots over the bar chart's category bands",
                                note: "Excel has no horizontal dot chart: every mark a dot (XY) over its category's band" } }];
      }
    }
    return [withTargets(make(seriesFor(m, { type, color: tvMarkColor(roles) }),
                             { numFmt: tvNumFmt(vm, m.ci), valueTitle: tvMeasureLabel(vm, m.ci) }))];
  }

  // several measures: dual axis / combo (different marks) or separate panes (same mark). Tableau's Dual Axis
  // folds the second axis onto the first; Synchronize Axis gives both one scale
  // each axis's own marks card (found by its axis, else its position), an Automatic one drawing what it inherits
  const marks = measures.map((m, i) => {
    const p = paneOfAxis(m.axis || 0);
    return p && p.id ? asBars(tvMarkToken(p.markClass)) || inherit : measureMark(m.ref, sheetMark, i, measures.length);
  });
  const space = (m, i) => vm.fmt.axisSpace(m.ref, valueShelf, measures.slice(0, i).some(x => tfSameField(x.ref, m.ref)) ? "1" : "0");
  const folded = measures.length === 2 && !!space(measures[1], 1).fold;
  const synced = measures.length === 2 && !!space(measures[1], 1).synchronized;
  if (new Set(marks).size > 1 && measures.length <= 4) {
    const series = [];
    const colourPanes = roles.panes.filter(p => p.encodings.some(e => e.channel === "color"));
    measures.forEach((m, i) => {
      // each axis has its own marks card: the colour field splits only the measure whose card holds it
      // (a KPI trend's min / max dots, not its area); the others keep their card's colour and opacity
      const pane = paneOfAxis(m.axis || 0);
      const split = !pane || !colourPanes.length || colourPanes.includes(pane);
      const own = pane ? tvHex(tfArgb(tvPaneRule(pane, "mark", "mark-color"))) : null;
      const opacity = pane ? Number(tvPaneRule(pane, "mark", "mark-transparency")) : NaN;
      seriesFor(m, {
        type: { ...tvSeriesType(marks[i]), secondary: measures.length === 2 && i === 1 && !synced },
        color: (color && color.measureNames && scale ? scale(vm.cols[m.ci].name, i) : null) || (!split && own) ||
               tvHex(TABLEAU_10[i % TABLEAU_10.length]),
        prefix: colorCi >= 0 && split, noPointColors: true, noColorSplit: !split,
        alpha: opacity >= 0 && opacity < 255 ? opacity / 255 : undefined
      }).forEach(s => series.push(s));
    });
    const spec = make(series, {
      numFmt: tvNumFmt(vm, measures[0].ci), valueTitle: tvMeasureLabel(vm, measures[0].ci),
      secondaryNumFmt: tvNumFmt(vm, measures[1].ci),
      secondaryTitle: measures.length === 2 ? tvMeasureLabel(vm, measures[1].ci) : undefined
    });
    // without Dual Axis Tableau draws each measure in a pane of its own: overlaid here, each on its own axis
    if (roles.source === "twb" && !folded) {
      spec.conversion = { strategy: "APPROXIMATE", note: "Tableau draws each measure in its own pane (no dual axis); Excel overlays them, each on its own axis" };
    }
    return [spec];
  }
  // the same mark on a dual axis: one chart. Bars are drawn one in the other, as Tableau draws them (bar in bar,
  // a progress bar over its track): the wider marks (the pane's Size) – else the longer values – on the primary axis
  // behind, the other on the secondary axis in front, narrower when Tableau draws it narrower. Synchronized axes share
  // one scale (the secondary axis is pinned to the primary's range and hidden); lines and areas of a synchronized
  // axis simply share the primary axis
  // a colour field on one axis's card splits only that axis's bars (a trend's highlighted min / max months)
  const colourCards = roles.panes.filter(p => p.encodings.some(e => e.channel === "color"));
  const coloured = m => { const p = paneOfAxis(m.axis || 0); return colorCi >= 0 && colorLevel < 0 && (!p || !colourCards.length || colourCards.includes(p)); };
  if (folded && (colorCi < 0 || measures.filter(coloured).length === 1)) {
    const type = tvSeriesType(marks[0]);
    const bars = type.type === "bar";
    const sizeOf = m => { const p = paneOfAxis(m.axis || 0); return (p && Number(tvPaneRule(p, "mark", "size"))) || 1; };
    const reach = m => Math.max(0, ...tvSum(vm, cats, m.ci).filter(v => v !== null).map(Math.abs));
    let order = [0, 1];
    if (bars) order.sort((a, b) => sizeOf(measures[b]) - sizeOf(measures[a]) || reach(measures[b]) - reach(measures[a]) || a - b);
    const [back, front] = order.map(k => measures[k]);
    const series = [back, front].flatMap((m, i) => seriesFor(m, {
      type: { ...type, secondary: i === 1 && (bars || !synced) }, color: paneColor(m, measures.indexOf(m)),
      noColorSplit: !coloured(m), noPointColors: true
    }).map(s => ({ ...s, labels: s.labels && constantOf(ctx, m.ref) === undefined })));
    const spec = make(series, { numFmt: tvNumFmt(vm, back.ci), valueTitle: tvMeasureLabel(vm, back.ci),
                                secondaryNumFmt: tvNumFmt(vm, front.ci), secondaryTitle: synced ? undefined : tvMeasureLabel(vm, front.ci) });
    // the coloured axis's bars stack by colour (one colour per category as a rule); the other draws whole bars
    spec.stacked = colorCi >= 0 && bars;
    spec.legend = false;
    if (bars && synced) { spec.secondarySync = true; spec.secondaryAxisHidden = true; }
    if (bars && sizeOf(front) < sizeOf(back) * 0.9) { spec.gapWidth = 50; spec.secondaryGapWidth = 250; }
    return [spec];
  }
  if (measures.length <= 6) {
    return measures.map((m, i) => withTargets(make(seriesFor(m, { type: tvSeriesType(marks[i]), color: tvMarkColor(roles) }),
      { numFmt: tvNumFmt(vm, m.ci), valueTitle: tvMeasureLabel(vm, m.ci), paneIndex: i, paneCount: measures.length })));
  }
  const series = measures.map((m, i) => ({ name: tvMeasureLabel(vm, m.ci), values: tvSum(vm, cats, m.ci),
    color: tvHex(TABLEAU_10[i % TABLEAU_10.length]), labels: tvLabelsOn(roles, m.ref), ...tvSeriesType(marks[i]) }));
  return [make(series, { numFmt: tvNumFmt(vm, measures[0].ci), valueTitle: "Value" })];

  /* bullet graph: a reference line per cell from another field (the target) over bars → a tick at each category's
   * target: dash markers over columns; across horizontal bars, vertical ticks drawn by the writer */
  function withTargets(spec) {
    const fmt = vm.fmt;
    const rl = fmt.hasModel && fmt.referenceLines
      ? fmt.referenceLines().find(r => r.scope === "per-cell" && r.field && r.axis && !tfSameField(r.field, r.axis)) : null;
    if (!rl || !spec.series.some(s => (s.type || spec.kind) === "bar") || spec.percent) return spec;
    const ti = vm.cols.findIndex(c => c.ref && tfSameField(c.ref, rl.field));
    if (ti < 0) return spec;
    // the line's value per cell: the target's average over the cell's marks (one mark per category as a rule)
    const sums = tvSum(vm, cats, ti), counts = new Array(cats.count).fill(0);
    vm.rows.forEach(r => { const i = cats.indexOf(r); if (i !== undefined && tfDvNum(r[ti]) !== null) counts[i]++; });
    const targets = sums.map((v, i) => v === null || !counts[i] ? null : v / counts[i]);
    const name = tvMeasureLabel(vm, ti);
    const conversion = { strategy: /** @type {"CONSTRUCTED"} */ ("CONSTRUCTED"), output: `${horizontal ? "Bar" : "Column"} chart with target ticks (bullet graph)`,
                         note: "Excel has no bullet graph: the actual values as bars, each category's target as a tick" };
    if (horizontal) return { ...spec, targets: { name, values: targets, color: "333333" }, conversion };
    const tick = { name, values: targets, type: /** @type {"line"} */ ("line"), line: false, marker: true, markerSymbol: "dash",
                   markerSize: 20, color: "333333", labels: false };
    return { ...spec, kind: "combo", series: [...spec.series, tick], conversion };
  }
}

/**
 * Names what an Excel chart cannot show of the Tableau view (the conversion becomes APPROXIMATE): bars that sum
 * Tableau's segments per detail member, bar widths from Size, a horizontal line or dot plot drawn vertically.
 * @param {ChartSpec} spec @param {ChartContext} ctx
 */
function tvNoteApproximations(spec, ctx) {
  if (spec.conversion || !/^(bar|line|area|combo)$/.test(spec.kind)) return;
  const { vm, roles } = ctx;
  const notes = [];
  const bars = spec.series.some(s => (s.type || spec.kind) === "bar" && s.color !== null);
  const lines = spec.series.filter(s => !s.refLine).every(s => (s.type || spec.kind) === "line");
  if (lines && spec.barDir === "bar") {
    spec.barDir = "col";                                   // Excel draws line charts vertically only
    notes.push("drawn vertically: Excel has no horizontal line or dot chart");
  }
  // bars sized by a measure (bar widths carry data); a size on another card (shapes beside the bars) is not the bars'
  const sizedBars = roles.panes.some(p => {
    const e = p.encodings.find(x => x.channel === "size");
    const mark = tvMarkToken(p.markClass) || (spec.kind === "bar" ? "bar" : "");
    return e && mark === "bar" && tvIsMeasureRef(ctx.model || null, e.field);
  });
  if (bars && sizedBars) notes.push("Tableau sizes these bars by a measure; Excel bars all have one width");
  if (bars) {
    const shelfDims = [...roles.rows.dims, ...roles.cols.dims].map(d => d.ci);
    const colorCi = roles.color && !roles.color.measureNames && roles.color.ci >= 0 ? roles.color.ci : -1;
    const detail = [...roles.detailDims, ...roles.textDims].filter(ci => ci !== colorCi && !shelfDims.includes(ci));
    if (detail.length) {
      const seen = new Map();
      const split = vm.rows.some(r => {
        const k = [...shelfDims, colorCi].map(ci => ci >= 0 ? String(r[ci] && r[ci].formattedValue) : "").join("\u0001");
        const v = detail.map(ci => String(r[ci] && r[ci].formattedValue)).join("\u0001");
        if (!seen.has(k)) { seen.set(k, v); return false; }
        return seen.get(k) !== v;
      });
      if (split) notes.push(`one bar per category: Tableau's segments per ${detail.map(ci => tvMeasureLabel(vm, ci)).join(" / ")} are summed (totals unchanged)`);
    }
  }
  if (notes.length) spec.conversion = { strategy: "APPROXIMATE", note: notes.join("; ") };
}

/** a lollipop's conversion: thin bars, and a dot at each bar's end */
function lollipopConversion(horizontal) {
  return { strategy: /** @type {"CONSTRUCTED"} */ ("CONSTRUCTED"), output: `${horizontal ? "Bar" : "Column"} chart (thin bars) with a dot on each`,
           note: "Excel has no lollipop chart: thin bars, each with a dot at its end" + (horizontal ? " (dots drawn over the category bands)" : "") };
}

/* ── funnel: stages = categories with one value each → stacked bars of an
 *    invisible half-gap + the value, so every stage is centred ─────────────── */
/**
 * @param {ChartContext} ctx @param {ReturnType<typeof tvCategories>} cats @param {number} catCi
 * @param {RoleValue[]} measures @param {boolean} horizontal
 * @returns {ChartSpec[]}
 */
function tvFunnelSpecFromStages(ctx, cats, catCi, measures, horizontal) {
  const { vm, roles } = ctx;
  const values = new Array(cats.count).fill(null);
  vm.rows.forEach(r => {
    const i = cats.indexOf(r);
    if (i === undefined) return;
    measures.forEach(m => { const n = tfDvNum(r[m.ci]); if (n !== null) values[i] = (values[i] || 0) + n; });
  });
  const max = Math.max(0, ...values.filter(v => v !== null));
  const scale = tvColorScale(vm, roles, "bar");
  const pointColors = scale && roles.color && roles.color.ci === catCi
    ? cats.levels[0].map((v, i) => scale(v) || tvHex(TABLEAU_10[i % TABLEAU_10.length])) : undefined;
  return [{
    ...tvBaseSpec(vm), kind: "bar", barDir: horizontal ? "bar" : "col", stacked: true, gapWidth: 15, legend: false,
    gridlines: false, valueAxisHidden: true,
    categories: { names: [tvMeasureLabel(vm, catCi)], levels: cats.levels }, categoryTitle: "",
    numFmt: tvNumFmt(vm, measures[0].ci), valueTitle: "",
    series: [
      { name: "Offset", type: "bar", values: values.map(v => v === null ? null : (max - v) / 2), color: null, labels: false },
      { name: tvMeasureLabel(vm, catCi), type: "bar", values, color: tvMarkColor(roles), pointColors, labels: roles.labelRefs.length > 0 }
    ],
    conversion: { strategy: "CONSTRUCTED", output: "Funnel: centred stacked bars",
                  note: "built from stacked bars with an invisible offset, so it opens in every Excel version" }
  }];
}

/* ── histogram: the bar chart with touching bars ─────────────────────────── */
/** @param {ChartContext} ctx @returns {ChartSpec[]} */
function tvHistogramSpecs(ctx) {
  return tvCartesianSpecs(ctx).map(s => ({ ...s, gapWidth: 0,
    conversion: { strategy: "CONSTRUCTED", output: `Column chart of Tableau's bins (no gaps)`,
                  note: "Tableau's own bins and counts as touching columns – the same bins Tableau draws, not Excel's automatic ones" } }));
}

/* ── small multiples: discrete dimensions beside the axes split Tableau's view into a grid of panes, each a little
 *    chart of its own. Excel draws no panes inside a chart: one chart per pane, laid out as Tableau's grid ──────── */
const TV_MAX_PANES = 48;

/** a chart's values as its value axis has to reach them (stacked marks add up) @param {ChartSpec} spec @returns {number[]} */
function tvAxisExtent(spec) {
  const out = [], stacks = new Map();
  spec.series.forEach(s => (s.values || []).forEach((v, i) => {
    const x = num(v);
    if (x === null) return;
    if (!spec.stacked) { out.push(x); return; }
    const t = stacks.get(i) || { pos: 0, neg: 0 };
    if (x >= 0) t.pos += x; else t.neg += x;
    stacks.set(i, t);
  }));
  stacks.forEach(t => out.push(t.pos, t.neg));
  return out;
}

/** where Tableau ends an axis reaching v: the next tick of a round step (1, 2, 2.5, 5 × 10ⁿ) giving about four ticks */
function tvNiceEnd(v) {
  if (!v) return 0;
  const a = Math.abs(v), raw = a / 4, mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const step = ([1, 2, 2.5, 5, 10].find(k => k * mag >= raw - 1e-12) || 10) * mag;
  return Math.sign(v) * +(Math.ceil(a / step - 1e-9) * step).toPrecision(12);
}

/**
 * Tableau's grid of panes – rows: the discrete dimensions on Rows beside the measure, columns: the discrete dimensions
 * on Columns outside the innermost field – as one chart spec per pane, its place and header texts on spec.grid. Every
 * pane spans the same dates and the same value range (Tableau's uniform axis), unless the axis is independent.
 * Null when the view is no such grid (the panes then stay one chart).
 * @param {ChartContext} ctx @returns {ChartSpec[] | null}
 */
function tvSmallMultipleSpecs(ctx) {
  const { vm, roles } = ctx;
  if (!roles.rows.values.length || roles.cols.values.length) return null;          // panes of columns: measures on Rows
  const rowDims = roles.rows.dims, colDims = roles.cols.dims, outer = colDims.slice(0, -1);
  if (!rowDims.length || !colDims.length || [...rowDims, ...outer].some(d => d.continuous || d.ci < 0)) return null;
  const keyOf = (r, dims) => dims.map(d => tvText(r[d.ci])).join("\u0001");
  /** the members of the dimensions in Tableau's order (natural sort), with their header text */
  const members = dims => {
    const seen = new Map();
    vm.rows.forEach(r => { const k = keyOf(r, dims); if (!seen.has(k)) seen.set(k, dims.map(d => r[d.ci])); });
    return [...seen.entries()].sort(([, a], [, b]) => {
      for (let l = 0; l < dims.length; l++) { const d = tfNaturalCompare(a[l], b[l]); if (d) return d; }
      return 0;
    }).map(([key, dvs]) => ({ key, label: dvs.map(dv => tfDvText(dv)).join(" / ") }));
  };
  const rowM = members(rowDims), colM = outer.length ? members(outer) : [{ key: "", label: "" }];
  const n = rowM.length * colM.length;
  if (n < 2 || n > TV_MAX_PANES) return null;
  const paneRoles = { ...roles, rows: { ...roles.rows, dims: [] }, cols: { ...roles.cols, dims: colDims.slice(-1) } };
  const rowRot = vm.fmt.hasModel ? vm.fmt.headerOrientation(rowDims[rowDims.length - 1].ref) : null;
  const rowHead = { rowVertical: rowRot === -90 || rowRot === 90, rowLen: Math.max(...rowM.map(m => m.label.length)) };
  const colorScale = tvColorScale(vm, roles, ctx.markToken);
  /** @type {ChartSpec[]} */
  const specs = [];
  for (let r = 0; r < rowM.length; r++) {
    for (let c = 0; c < colM.length; c++) {
      const rows = vm.rows.filter(x => keyOf(x, rowDims) === rowM[r].key && (!outer.length || keyOf(x, outer) === colM[c].key));
      if (!rows.length) continue;                                      // Tableau leaves an empty pane blank
      const pane = tvCartesianSpecs({ ...ctx, vm: { ...vm, rows }, roles: paneRoles, colorScale });
      if (pane.length !== 1) return null;                              // panes of panes: kept as one chart
      specs.push({ ...pane[0], legend: false,
                   grid: { row: r, col: c, rows: rowM.length, cols: colM.length, rowLabel: rowM[r].label, colLabel: colM[c].label, ...rowHead } });
    }
  }
  if (specs.length < 2) return null;
  // one time axis for all: the earliest and latest date of any pane; other categories have to be the same in each
  const dated = specs.filter(s => s.dateAxis);
  if (dated.length === specs.length) {
    const min = Math.min(...dated.map(s => s.dateAxis.serials[0])), max = Math.max(...dated.map(s => s.dateAxis.serials[s.dateAxis.serials.length - 1]));
    specs.forEach(s => { s.dateAxis = { ...s.dateAxis, min, max }; });
  } else if (dated.length || specs.some(s => JSON.stringify(s.categories.levels) !== JSON.stringify(specs[0].categories.levels))) {
    return null;
  }
  // Tableau's uniform axis: every pane on the whole view's range (an independent axis gives each pane its own)
  const ref = roles.rows.values[0].ref;
  const independent = ref && /independent/.test(vm.fmt.axisSpace(ref, "rows", "0").rangeType || "");
  if (!independent && !specs.some(s => s.percent)) {
    const all = specs.flatMap(tvAxisExtent);
    const lo = Math.min(0, ...all), hi = Math.max(0, ...all);
    specs.forEach(s => {
      // the range of the exported values, not of the workbook: dropped again when the values follow slicers
      s.dataRange = { min: s.valueMin === undefined, max: s.valueMax === undefined };
      if (s.valueMin === undefined) s.valueMin = tvNiceEnd(lo);
      if (s.valueMax === undefined) s.valueMax = tvNiceEnd(hi);
    });
  }
  const conversion = { strategy: "CONSTRUCTED", output: `Small multiples: ${specs.length} charts in Tableau's ${rowM.length} × ${colM.length} grid`,
                       note: "Excel draws no panes inside one chart: each pane is its own chart, on the same axes" };
  specs.forEach(s => { s.conversion = s.conversion ? { ...s.conversion, note: `${s.conversion.note}; ${conversion.note}` } : conversion; });
  return specs;
}


/* ═══ charts/model/roles.js ═══════════════════════════════════════════════════════════════════════ */
/* Which summary column sits on which shelf / encoding (TWB or live spec). */

/* ── field roles: which summary column sits on which shelf / encoding ───── */
/**
 * @param {ViewModel} vm
 * @param {FormatModel | null} model
 * @param {any} [liveSpec] the live visual specification, when no workbook was loaded
 * @returns {Roles}
 */
function tvRoles(vm, model, liveSpec) {
  const sheet = vm.fmt.sheetModel;
  /** @type {Roles} */
  const roles = { rows: { values: [], dims: [], axisRefs: [] }, cols: { values: [], dims: [], axisRefs: [] }, measureNames: null,
                  color: null, angle: -1, size: -1, textDims: [], detailDims: [],
                  labelRefs: [], labelNames: [], panes: sheet ? sheet.panes : [], source: "none" };
  const addDim = (list, ci) => {
    if (ci >= 0 && (vm.cols[ci].isHeader || !numeric(vm.cols[ci])) && !list.includes(ci)) list.push(ci);
  };
  const pivoted = vm.cols.map((c, i) => c.pivoted ? i : -1).filter(i => i >= 0);
  const numeric = c => /^(int|float|real|integer|number|double)/i.test(String(c.dataType || "")) || c.pivoted;

  if (sheet) {
    roles.source = "twb";
    const colOf = ref => {
      if (!ref) return -1;
      let i = vm.cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
      if (i < 0) {
        const names = tfDisplayNames(model, ref);
        i = vm.cols.findIndex(c => names.includes(tfNorm(c.name)));
      }
      return i;
    };
    for (const shelf of ["rows", "cols"]) {
      let axisNo = 0;                              // each measure / Measure Values entry = one axis ("A + B")
      for (const ref of sheet[shelf]) {
        if (ref.name === "Measure Names") { roles.measureNames = shelf; continue; }
        if (ref.name === "Multiple Values") {
          roles[shelf].axisRefs.push(ref);
          const axis = axisNo++;
          pivoted.forEach(ci => roles[shelf].values.push({ ci, ref: vm.cols[ci].ref, mv: true, axis }));
          continue;
        }
        const ci = colOf(ref);
        if (ci < 0) continue;
        if (tvIsMeasureRef(model, ref)) { roles[shelf].axisRefs.push(ref); roles[shelf].values.push({ ci, ref, axis: axisNo++ }); }
        else roles[shelf].dims.push({ ci, ref, continuous: ref.type === "qk" });
      }
    }
    const enc = channel => {
      for (const p of sheet.panes) {
        const e = p.encodings.find(x => x.channel === channel);
        if (e) return e.field;
      }
      return null;
    };
    const colorRef = enc("color");
    if (colorRef) {
      if (colorRef.name === "Measure Names") roles.color = { measureNames: true, ref: colorRef };
      else {
        const ci = colOf(colorRef);
        if (ci >= 0) roles.color = { ci, ref: colorRef, continuous: tvIsMeasureRef(model, colorRef) || colorRef.type === "qk" };
      }
    }
    const angleRef = enc("wedge-size");
    if (angleRef) roles.angle = angleRef.name === "Multiple Values" ? (pivoted[0] ?? -1) : colOf(angleRef);
    const sizeRef = enc("size");
    if (sizeRef) roles.size = colOf(sizeRef);
    sheet.panes.forEach(p => {
      p.encodings.filter(e => e.channel === "text" || e.channel === "label").forEach(e => {
        roles.labelRefs.push(e.field);
        if (!tvIsMeasureRef(model, e.field)) addDim(roles.textDims, colOf(e.field));
      });
      p.encodings.filter(e => e.channel === "lod").forEach(e => addDim(roles.detailDims, colOf(e.field)));
      p.labelRuns.forEach(r => r.refs.forEach(x => roles.labelRefs.push(x)));
    });
    return roles;
  }

  // ── no workbook loaded: use the live visual specification ──
  const spec = liveSpec && typeof liveSpec === "object" ? liveSpec : null;
  const byName = name => {
    const key = tfNorm(name);
    const strip = s => tfNorm(s).replace(/^[a-z]+\((.*)\)$/, "$1");
    let i = vm.cols.findIndex(c => tfNorm(c.name) === key);
    if (i < 0) i = vm.cols.findIndex(c => strip(c.name) === strip(name));
    return i;
  };
  const fieldName = f => f == null ? "" : typeof f === "string" ? f : (f.name || f.fieldName || f.caption || "");
  if (spec && (Array.isArray(spec.rowFields) || Array.isArray(spec.columnFields))) {
    roles.source = "live";
    for (const [shelf, list] of [["rows", spec.rowFields || []], ["cols", spec.columnFields || []]]) {
      list.forEach(f => {
        const name = fieldName(f);
        if (/^measure names$/i.test(name)) { roles.measureNames = shelf; return; }
        if (/^measure values$/i.test(name)) { pivoted.forEach(ci => roles[shelf].values.push({ ci, ref: vm.cols[ci].ref, mv: true })); return; }
        const ci = byName(name);
        if (ci < 0) return;
        if (numeric(vm.cols[ci]) && !vm.cols[ci].isHeader) roles[shelf].values.push({ ci, ref: vm.cols[ci].ref });
        else roles[shelf].dims.push({ ci, ref: vm.cols[ci].ref, continuous: /^date/i.test(String(vm.cols[ci].dataType || "")) });
      });
    }
    const marks = Array.isArray(spec.marksSpecifications) ? spec.marksSpecifications : [];
    roles.liveMarks = marks.map(m => tvMarkToken(m && (m.type || m.markType)));   // one marks card per measure
    const active = marks[spec.activeMarksSpecificationIndex] || marks[0];
    (active && active.encodings || []).forEach(e => {
      const name = fieldName(e.field);
      const type = String(e.type || e.encodingType || "").toLowerCase();
      if (type === "color") {
        if (/^measure names$/i.test(name)) roles.color = { measureNames: true };
        else { const ci = byName(name); if (ci >= 0) roles.color = { ci, continuous: numeric(vm.cols[ci]) && !vm.cols[ci].isHeader }; }
      } else if (type === "angle") roles.angle = byName(name);
      else if (type === "size") roles.size = byName(name);
      else if (type === "label" || type === "text") {
        roles.labelNames.push(tfNorm(name));
        const ci = byName(name);
        if (ci >= 0 && !numeric(vm.cols[ci])) addDim(roles.textDims, ci);
      } else if (type === "detail") addDim(roles.detailDims, byName(name));
    });
    if (roles.rows.values.length || roles.cols.values.length) return roles;
  }

  // ── nothing structural known: dimensions → categories, measures → rows ──
  roles.source = "summary";
  vm.cols.forEach((c, ci) => {
    if (c.isHeader) roles.cols.dims.push({ ci, ref: c.ref, continuous: /^date/i.test(String(c.dataType || "")) });
    else if (numeric(c)) roles.rows.values.push({ ci, ref: c.ref, mv: !!c.pivoted });
  });
  return roles;
}


/* ═══ charts/model/packing.js ═════════════════════════════════════════════════════════════════════ */
/* Circle packing for packed bubbles: largest bubble in the middle, the others around it. */

/* the two centres where a circle of radius r touches circles p and q (gap included) */
function touchingBoth(p, q, r, gap) {
  const r1 = p.r + r + gap, r2 = q.r + r + gap;
  const dx = q.x - p.x, dy = q.y - p.y, d = Math.hypot(dx, dy);
  if (!d || d > r1 + r2 || d < Math.abs(r1 - r2)) return [];
  const a = (r1 * r1 - r2 * r2 + d * d) / (2 * d);
  const h = Math.sqrt(Math.max(0, r1 * r1 - a * a));
  const mx = p.x + a * dx / d, my = p.y + a * dy / d;
  return [{ x: mx + h * dy / d, y: my - h * dx / d }, { x: mx - h * dy / d, y: my + h * dx / d }];
}

/**
 * Greedy packing, biggest first: every circle goes where it touches two placed circles, as close to
 * the centre as it fits. O(n³) – fine for the few hundred marks a packed-bubble view shows.
 * @param {number[]} radii
 * @param {number} [gap] space between circles, in radius units
 * @returns {{ centres: { x: number, y: number }[], box: { w: number, h: number, cx: number, cy: number } }}
 */
function tvPackCircles(radii, gap = 0) {
  const placed = [];
  const centres = new Array(radii.length);
  const order = radii.map((_, i) => i).sort((a, b) => radii[b] - radii[a]);
  const fits = (x, y, r) => placed.every(p => Math.hypot(p.x - x, p.y - y) >= p.r + r + gap - 1e-9);
  for (const i of order) {
    const r = radii[i];
    let best = null;
    if (!placed.length) best = { x: 0, y: 0 };
    else if (placed.length === 1) best = { x: placed[0].r + r + gap, y: 0 };
    else {
      for (let a = 0; a < placed.length; a++) {
        for (let b = a + 1; b < placed.length; b++) {
          for (const c of touchingBoth(placed[a], placed[b], r, gap)) {
            const d = Math.hypot(c.x, c.y);
            if ((!best || d < best.d) && fits(c.x, c.y, r)) best = { ...c, d };
          }
        }
      }
      if (!best) best = { x: Math.max(...placed.map(p => p.x + p.r)) + r + gap, y: 0 };
    }
    placed.push({ x: best.x, y: best.y, r });
    centres[i] = { x: best.x, y: best.y };
  }
  const x0 = Math.min(...placed.map(p => p.x - p.r)), x1 = Math.max(...placed.map(p => p.x + p.r));
  const y0 = Math.min(...placed.map(p => p.y - p.r)), y1 = Math.max(...placed.map(p => p.y + p.r));
  return { centres, box: { w: x1 - x0, h: y1 - y0, cx: (x0 + x1) / 2, cy: (y0 + y1) / 2 } };
}


/* ═══ charts/writer/constants.js ══════════════════════════════════════════════════════════════════ */
/* OOXML namespaces, relationship and content types, axis ids. */

const NS = {
  c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  xdr: "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
  rels: "http://schemas.openxmlformats.org/package/2006/relationships",
  cx: "http://schemas.microsoft.com/office/drawing/2014/chartex",
  cx1: "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006"
};

const REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

const REL_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

const REL_CHARTEX = "http://schemas.microsoft.com/office/2014/relationships/chartEx";

const CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml";

const CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

const CT_CHARTEX = "application/vnd.ms-office.chartex+xml";

const REL_CHARTSTYLE = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";

const REL_CHARTCOLORS = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";

const CT_CHARTSTYLE = "application/vnd.ms-office.chartstyle+xml";

const CT_CHARTCOLORS = "application/vnd.ms-office.chartcolorstyle+xml";

const EMU_PER_PX = 9525;

const AX = { cat: 50010, val: 50020, cat2: 50030, val2: 50040, x2: 50050, y2: 50060 };

const C15 = "http://schemas.microsoft.com/office/drawing/2012/chart";

/* XLSX package surgery: relationship and content types of the parts injectCharts adds */
const EMPTY_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS.rels}"></Relationships>`;


/* ═══ charts/writer/xml-util.js ═══════════════════════════════════════════════════════════════════ */
/* XML escaping, cell references and other small writer helpers. */

/* ── small helpers ─────────────────────────────────────────────────────── */
function esc(s) {
  return String(s == null ? "" : s)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function colName(i) {
  let s = "", n = i + 1;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - m) / 26); }
  return s;
}

/* 0-based (col,row) → 'Sheet'!$A$1[:$B$9] */
function cellRef(sheet, c1, r1, c2, r2) {
  const q = "'" + String(sheet).replace(/'/g, "''") + "'!";
  const a = "$" + colName(c1) + "$" + (r1 + 1);
  return q + a + (c2 === undefined ? "" : ":$" + colName(c2) + "$" + (r2 + 1));
}

function hex(c) {
  const s = String(c || "").replace(/^#/, "").toUpperCase();
  return s.length === 8 ? s.slice(2) : s;
}

function num(v) { return typeof v === "number" && isFinite(v) ? v : null; }

function attr(tag, name) {
  const m = tag.match(new RegExp("\\s" + name.replace(":", "\\:") + "=\"([^\"]*)\""));
  return m ? m[1] : null;
}

/* boundaries of each category level: a new group starts where any level ≤ l changes */
function levelStarts(levels) {
  const n = levels.length ? levels[0].length : 0;
  return levels.map((_, l) => Array.from({ length: n }, (__, i) =>
    i === 0 || levels.slice(0, l + 1).some(lv => lv[i] !== lv[i - 1])));
}


/* ═══ charts/writer/pie.js ════════════════════════════════════════════════════════════════════════ */
/* Pie / doughnut chart parts drawn the way Tableau draws them:
 *   pieChartXml     – the chart: slices with their colours (a selected mark outlined), a pinned plot area so
 *                     the diameter is Tableau's, an optional doughnut hole
 *   userShapesXml   – labels as text boxes at Tableau's label positions (Excel data labels wrap and shrink the
 *                     pie), the donut hole's own fill and the centre text
 *   tipShapeXml     – Tableau's tooltip: an invisible wedge over each slice whose hyperlink ScreenTip is the
 *                     tooltip text (Excel's own chart hover text cannot be changed)
 * Knows nothing about Tableau: colours, fonts and texts arrive resolved. */

const PIE_EMU_PER_PX = 9525;                                        // 96 dpi
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_CDR = "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing";
const REL_USER_SHAPES = NS.r + "/chartUserShapes";
const CT_USER_SHAPES = "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml";
const REL_HYPERLINK = NS.r + "/hyperlink";

/* ── text ─────────────────────────────────────────────────────────────── */
/** @typedef {{ name?: string, size?: number, bold?: boolean, italic?: boolean, underline?: boolean, color?: string }} PieFont */
/** @param {PieFont} [f] */
function fontAttrs(f) {
  f = f || {};
  return ` lang="en-US" sz="${Math.round((f.size || 9) * 100)}" b="${f.bold ? 1 : 0}" i="${f.italic ? 1 : 0}"` +
         (f.underline ? ' u="sng"' : "");
}
/** @param {PieFont} [f] */
function fontKids(f) {
  f = f || {};
  return `<a:solidFill><a:srgbClr val="${f.color || "333333"}"/></a:solidFill>` + (f.name ? `<a:latin typeface="${esc(f.name)}"/>` : "");
}
/** lines [[{text, font}]] → one paragraph per line, one run per piece @param {string} [algn] "l" | "ctr" | "r" */
function parasXml(lines, algn) {
  const pPr = algn ? `<a:pPr algn="${algn}"/>` : "";
  return lines.map(runs => {
    const rs = runs.filter(r => r.text !== "").map(r =>
      `<a:r><a:rPr${fontAttrs(r.font)}>${fontKids(r.font)}</a:rPr><a:t>${esc(r.text)}</a:t></a:r>`).join("");
    return rs ? `<a:p>${pPr}${rs}</a:p>` : `<a:p>${pPr}<a:endParaRPr lang="en-US"/></a:p>`;
  }).join("");
}
const richXml = lines => `<c:rich><a:bodyPr/><a:lstStyle/>${parasXml(lines)}</c:rich>`;
const txPrXml = (f, bodyPr) => `<c:txPr>${bodyPr || "<a:bodyPr/>"}<a:lstStyle/><a:p><a:pPr><a:defRPr${fontAttrs(f)}>${fontKids(f)}</a:defRPr></a:pPr>` +
                     `<a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
const NO_LINE = `<a:ln><a:noFill/></a:ln>`;
const fillXml = c => c ? `<a:solidFill><a:srgbClr val="${c}"/></a:solidFill>` : `<a:noFill/>`;

/* ── data: worksheet ranges (linked, "Select Data" shows them) or literals ── */
const sheetRef = (sheet, range) => `'${String(sheet).replace(/'/g, "''")}'!${range}`;
const strPts = vals => `<c:ptCount val="${vals.length}"/>` + vals.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join("");
const numPts = vals => `<c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>` +
  vals.map((v, i) => `<c:pt idx="${i}"><c:v>${Number(v)}</c:v></c:pt>`).join("");
const pieCatXml = (vals, ref) => ref
  ? `<c:cat><c:strRef><c:f>${esc(sheetRef(ref.sheet, ref.cat))}</c:f><c:strCache>${strPts(vals)}</c:strCache></c:strRef></c:cat>`
  : `<c:cat><c:strLit>${strPts(vals)}</c:strLit></c:cat>`;
const pieValXml = (vals, ref) => ref
  ? `<c:val><c:numRef><c:f>${esc(sheetRef(ref.sheet, ref.val))}</c:f><c:numCache>${numPts(vals)}</c:numCache></c:numRef></c:val>`
  : `<c:val><c:numLit>${numPts(vals)}</c:numLit></c:val>`;

/* ── chart frame ──────────────────────────────────────────────────────── */
function titleXml(title, font) {
  if (!title) return '<c:autoTitleDeleted val="1"/>';
  return `<c:title><c:tx>${richXml([[{ text: title, font }]])}</c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`;
}
/** plot: { x, y, w, h } fractions of the chart (inner plot area) – pins the pie's diameter */
function layoutXml(plot) {
  if (!plot) return "<c:layout/>";
  const v = n => Math.max(0, Math.min(1, n)).toFixed(4);
  return `<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/>` +
    `<c:x val="${v(plot.x)}"/><c:y val="${v(plot.y)}"/><c:w val="${v(plot.w)}"/><c:h val="${v(plot.h)}"/></c:manualLayout></c:layout>`;
}

const NS_C15 = "http://schemas.microsoft.com/office/drawing/2012/chart";

/** slice labels Excel takes from a range of cells ("Value From Cells", Excel 2013+): they move with the slices
 * @param {{ range: string, texts: string[], font: any }} l @param {boolean} doughnut */
function cellLabelsXml(l, doughnut) {
  // wrap="none": each label keeps Tableau's lines (Excel would wrap them to a narrow width, mid-word)
  const look = `<c:spPr><a:noFill/>${NO_LINE}</c:spPr>${txPrXml(l.font, '<a:bodyPr wrap="none"/>')}` +
    (doughnut ? "" : `<c:dLblPos val="outEnd"/>`) +
    `<c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/>` +
    `<c:showBubbleSize val="0"/>`;
  const fromCells = `<c:extLst><c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="${NS_C15}"><c15:showDataLabelsRange val="1"/></c:ext></c:extLst>`;
  // a label pushed off its neighbour (cellLabelShifts): its shift from Excel's own place
  const moved = (l.shifts || []).map((d, i) => d ? `<c:dLbl><c:idx val="${i}"/><c:layout><c:manualLayout>` +
    `<c:x val="${d.x.toFixed(4)}"/><c:y val="${d.y.toFixed(4)}"/></c:manualLayout></c:layout>${look}${fromCells}</c:dLbl>` : "").join("");
  const dLbls = `<c:dLbls>${moved}${look}${doughnut ? "" : `<c:showLeaderLines val="1"/>`}${fromCells}</c:dLbls>`;
  const range = `<c:extLst><c:ext uri="{02D57815-91ED-43cb-92C2-25804820EDAC}" xmlns:c15="${NS_C15}"><c15:datalabelsRange>` +
    `<c15:f>${esc(l.range)}</c15:f><c15:dlblRangeCache>${strPts(l.texts)}</c15:dlblRangeCache></c15:datalabelsRange></c:ext></c:extLst>`;
  return { dLbls, range };
}

/**
 * spec: { seriesName, categories[], values[], ref? {sheet, cat, val}, title?, titleFont?, font?, bg? (RRGGBB, null =
 *   see-through), plot?, holeSize? (% of the diameter → doughnut), hasShapes?, points: [{ color, selected }],
 *   cellLabels? { range, texts[], font } (labels from cells) }
 * @param {any} spec @returns {string}
 */
function pieChartXml(spec) {
  const dPts = (spec.points || []).map((p, i) =>
    `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/><c:spPr>${fillXml(p.color)}` +
    (p.selected ? `<a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>`
                : !p.color ? NO_LINE : `<a:ln w="9525"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>`) +
    `</c:spPr></c:dPt>`).join("");
  const lbl = spec.cellLabels ? cellLabelsXml(spec.cellLabels, !!spec.holeSize) : { dLbls: "", range: "" };
  const ser = `<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>${esc(spec.seriesName || "")}</c:v></c:tx>` +
    `${dPts}${lbl.dLbls}${pieCatXml(spec.categories, spec.ref)}${pieValXml(spec.values, spec.ref)}${lbl.range}</c:ser>`;
  const plot = spec.holeSize
    ? `<c:doughnutChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/>` +
      `<c:holeSize val="${Math.max(10, Math.min(90, Math.round(spec.holeSize)))}"/></c:doughnutChart>`    // Excel allows 10–90
    : `<c:pieChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/></c:pieChart>`;
  return XML_DECL +
    `<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">` +
    `<c:roundedCorners val="0"/><c:chart>${titleXml(spec.title, spec.titleFont)}` +
    `<c:plotArea>${layoutXml(spec.plot)}${plot}<c:spPr><a:noFill/>${NO_LINE}</c:spPr></c:plotArea>` +
    `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `<c:spPr>${fillXml(spec.bg)}${NO_LINE}</c:spPr>${txPrXml(spec.font)}` +
    (spec.hasShapes ? `<c:userShapes r:id="rId1"/>` : "") + `</c:chartSpace>`;
}

/**
 * Text boxes / shapes drawn on top of a chart (its chartDrawing part), anchored relative to the chart so they
 * move and scale with it; later items draw on top.
 * texts: [{ x, y, w, h (px inside the chart), lines: [[{text, font}]], align: "l"|"ctr"|"r", anchor: "t"|"ctr"|"b",
 *           shape?: preset geometry ("ellipse" …; default a borderless text box), fill?: RRGGBB,
 *           link?: the cell whose text the box shows ('Sheet'!$A$1) }]
 * @param {any[]} texts @param {{ w: number, h: number }} size chart size px @returns {string}
 */
function userShapesXml(texts, size) {
  const fx = (v, d) => Math.max(0, Math.min(1, v / d)).toFixed(5);
  // a box linked to a cell: its text is Excel's TxLink field (a plain run would never refresh), in the line's largest font
  const linked = (t, i) => {
    const runs = t.lines[0] || [], f = runs.reduce((a, r) => ((r.font && r.font.size) || 0) > ((a && a.size) || 0) ? r.font : a, null);
    const id = `{6A1D0C2E-0000-4000-8000-${String(i + 1).padStart(12, "0")}}`;
    return `<a:p><a:pPr algn="${t.align || "ctr"}"/><a:fld id="${id}" type="TxLink"><a:rPr${fontAttrs(f)}>${fontKids(f)}</a:rPr>` +
           `<a:t>${esc(runs.map(r => r.text).join(""))}</a:t></a:fld></a:p>`;
  };
  const paras = (t, i) => t.link ? linked(t, i)
    : t.lines && t.lines.length ? parasXml(t.lines, t.align || "ctr") : `<a:p><a:endParaRPr lang="en-US"/></a:p>`;
  return XML_DECL + `<c:userShapes xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:cdr="${NS_CDR}">` + texts.map((t, i) =>
    `<cdr:relSizeAnchor><cdr:from><cdr:x>${fx(t.x, size.w)}</cdr:x><cdr:y>${fx(t.y, size.h)}</cdr:y></cdr:from>` +
    `<cdr:to><cdr:x>${fx(t.x + t.w, size.w)}</cdr:x><cdr:y>${fx(t.y + t.h, size.h)}</cdr:y></cdr:to>` +
    `<cdr:sp macro="" textlink="${t.link ? esc(t.link) : ""}"><cdr:nvSpPr><cdr:cNvPr id="${i + 2}" name="${t.shape ? "Shape" : "Label"} ${i + 1}"/>` +
    `<cdr:cNvSpPr${t.shape ? "" : ' txBox="1"'}/></cdr:nvSpPr>` +
    `<cdr:spPr><a:prstGeom prst="${t.shape || "rect"}"><a:avLst/></a:prstGeom>${fillXml(t.fill)}${NO_LINE}</cdr:spPr>` +
    `<cdr:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="${t.anchor || "t"}"/><a:lstStyle/>` +
    `${paras(t, i)}</cdr:txBody></cdr:sp></cdr:relSizeAnchor>`).join("") + `</c:userShapes>`;
}

const xfrmXml = (x, y, w, h) => `<a:off x="${Math.round(x * PIE_EMU_PER_PX)}" y="${Math.round(y * PIE_EMU_PER_PX)}"/>` +
                                `<a:ext cx="${Math.round(w * PIE_EMU_PER_PX)}" cy="${Math.round(h * PIE_EMU_PER_PX)}"/>`;

/**
 * An invisible wedge over a slice: hovering shows its hyperlink ScreenTip (the Tableau tooltip). A 100 %
 * transparent fill still catches the mouse; noFill would not.
 * tip: { x, y, w, h (px inside the chart), prst ("pie" | "blockArc" | "ellipse"), adj: [avLst values], text }
 * @param {any} t @param {number} id @param {string} rid the hyperlink relationship
 */
function tipShapeXml(t, id, rid) {
  const gd = (t.adj || []).map((v, k) => `<a:gd name="adj${k + 1}" fmla="val ${Math.round(v)}"/>`).join("");
  const text = String(t.text).slice(0, 255);                   // Excel's ScreenTip limit
  return `<xdr:sp macro="" textlink=""><xdr:nvSpPr>` +
    `<xdr:cNvPr id="${id}" name="Tooltip ${id}"><a:hlinkClick xmlns:r="${NS.r}" r:id="${rid}" tooltip="${esc(text).replace(/\n/g, "&#xA;")}"/></xdr:cNvPr>` +
    `<xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm>${xfrmXml(t.x, t.y, t.w, t.h)}</a:xfrm><a:prstGeom prst="${t.prst}"><a:avLst>${gd}</a:avLst></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:solidFill>${NO_LINE}</xdr:spPr></xdr:sp>`;
}

/**
 * The anchor of a chart with tooltip wedges: chart and wedges grouped, so the wedges keep their exact px
 * positions over the slices whatever the cell sizes.
 * @param {any} chart ChartJob (col / row / offsets, widthPx / heightPx, name) @param {string} chartRid
 * @param {{ next: number }} ids shape ids @param {string[]} tipRids one hyperlink relationship per tip
 */
function groupedAnchorXml(chart, chartRid, ids, tipRids) {
  const w = chart.widthPx, h = chart.heightPx;
  const frame = `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${ids.next++}" name="${esc(chart.name || "Chart")}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm>${xfrmXml(0, 0, w, h)}</xdr:xfrm>` +
    `<a:graphic><a:graphicData uri="${NS.c}"><c:chart xmlns:c="${NS.c}" xmlns:r="${NS.r}" r:id="${chartRid}"/></a:graphicData></a:graphic></xdr:graphicFrame>`;
  const group = `<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="${ids.next++}" name="${esc((chart.name || "Chart") + " group")}"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>` +
    `<xdr:grpSpPr><a:xfrm>${xfrmXml(0, 0, w, h)}<a:chOff x="0" y="0"/><a:chExt cx="${Math.round(w * PIE_EMU_PER_PX)}" cy="${Math.round(h * PIE_EMU_PER_PX)}"/></a:xfrm></xdr:grpSpPr>` +
    frame + chart.tips.map((t, k) => tipShapeXml(t, ids.next++, tipRids[k])).join("") + `</xdr:grpSp>`;
  return `<xdr:oneCellAnchor><xdr:from><xdr:col>${chart.col}</xdr:col><xdr:colOff>${Math.round((chart.colOffPx || 0) * PIE_EMU_PER_PX)}</xdr:colOff>` +
    `<xdr:row>${chart.row}</xdr:row><xdr:rowOff>${Math.round((chart.rowOffPx || 0) * PIE_EMU_PER_PX)}</xdr:rowOff></xdr:from>` +
    `<xdr:ext cx="${Math.round(w * PIE_EMU_PER_PX)}" cy="${Math.round(h * PIE_EMU_PER_PX)}"/>${group}<xdr:clientData/></xdr:oneCellAnchor>`;
}


/* ═══ charts/model/pie.js ═════════════════════════════════════════════════════════════════════════ */
/* ══════════════════════════════════════════════════════════════════════════
 * PIES AND DONUTS – Tableau marks → pie model → chart jobs. One summary row = one Tableau mark = one slice.
 *   Angle (wedge-size)  → slice value; no Angle → equal slices (as Tableau draws them)
 *   Color               → slice fill: palette / custom colours, or gradient for a measure
 *   Detail / other dims → more slices, never merged
 *   Label / Tooltip     → Tableau's template filled with Tableau's formatted values
 *   Rows / Cols dims    → one pie per pane
 * A donut is a dual axis (MIN(0) twice): the slices plus a hole layer on top (its own label – usually the
 * total – drawn in the centre); a nested donut is two coloured pies, the top one smaller.
 * ══════════════════════════════════════════════════════════════════════════ */

/* Measured in Tableau 2026.2 (1000×800 dashboard): a view with nothing on Rows/Columns draws each pane as a
 * 160×160 px cell at the zone's top-left (inside the 4 px zone margin), pie centred in it. Size slider →
 * diameter is linear in the TWB value: untouched = 80 px, 1.4613 → 117 px (80 × 1.46); a pie bigger than its
 * cell is clipped to it. */
const PIE_DIAMETER_PX = 80;
const PIE_CELL_PX = 160;
const ZONE_MARGIN = 4;
/** @param {number | null | undefined} sizeValue Size slider */
const pieDiameterPx = sizeValue => sizeValue === null || sizeValue === undefined ? PIE_DIAMETER_PX : PIE_DIAMETER_PX * Math.max(0.01, sizeValue);

/** a pie label value's number format: the workbook's, else for % of Total Tableau's 2 decimals (its API's
 * formatted value has 4) @param {any} fmt @param {FieldRef} ref */
const pieNumFmt = (fmt, ref) => fmt.numFmtFor(ref) || (/^pcto/i.test(String(ref.deriv || "")) ? "0.00%" : null);

/** Excel chart font from Tableau props; measureName = Tableau's own font, for text measuring */
function pieFont(p) {
  const f = tfExcelFont(p || {});
  return { name: f.name, size: f.size, bold: f.bold, italic: f.italic, underline: !!f.underline,
           color: f.color ? f.color.argb.slice(2) : "333333", measureName: p && p.fontName };
}

/* ── label placement ────────────────────────────────────────────────────────
 * Tableau's pie label placement (measured in Tableau 2026.2): a label sits at its slice's outer mid-point, on
 * the side the slice faces – right half: label starts there, left half: ends there; top half: above, bottom
 * half: below – and is pushed back inside the pane where it would stick out. With "Allow labels to overlap
 * other marks" off, a label hitting an already placed one is hidden, biggest slice first. Tableau draws label
 * text at the screen's DPI scale (9 pt = 15 px at 125 %), the pane itself in plain px. */
const LABEL_CHAR_EM = 0.52;            // text width per character when there is no canvas to measure with
const LABEL_LINE_EM = 1.33;            // Tableau label line height (20 px lines for 15 px text)
const LABEL_PAD = 2;
/** @type {CanvasRenderingContext2D | null} */
let measureCtx = null;
function textWidthPx(text, font, px) {
  try {
    measureCtx = measureCtx || document.createElement("canvas").getContext("2d");
    measureCtx.font = `${font.italic ? "italic " : ""}${font.bold ? "bold " : ""}${px}px "${font.measureName || font.name}", Arial, sans-serif`;
    return measureCtx.measureText(text).width;
  } catch (e) {
    return text.length * px * LABEL_CHAR_EM;
  }
}

/**
 * Per point a label box {x, y, w, h, up} in pane px, or null (no label / hidden).
 * @param {any[]} points @param {{ cx: number, cy: number, R: number, w: number, h: number, cull: boolean, textScale: number }} g
 */
function pieLabelLayout(points, g) {
  const EPS = 1e-6;
  const total = points.reduce((s, p) => s + p.value, 0) || 1;
  let acc = 0;
  const boxes = points.map(p => {
    const mid = (acc + p.value / 2) / total * 2 * Math.PI;      // clockwise from 12 o'clock (firstSliceAng 0)
    acc += p.value;
    if (!p.label || !p.label.length) return null;
    const s = Math.sin(mid), c = Math.cos(mid);
    const right = s > EPS || (Math.abs(s) <= EPS && c > 0);    // exactly 12 / 6 o'clock: the way the circle turns
    const up = c > EPS || (Math.abs(c) <= EPS && s < 0);       // exactly 3 o'clock → below, 9 o'clock → above
    let w = 0, h = 0;
    p.label.forEach(line => {
      const px = f => (f.size || 9) * 4 / 3 * g.textScale;
      w = Math.max(w, line.reduce((t, x) => t + textWidthPx(x.text, x.font, px(x.font)), 0));
      h += (line.length ? Math.max(...line.map(x => px(x.font))) : px({})) * LABEL_LINE_EM;
    });
    w += 2 * LABEL_PAD;
    const ax = g.cx + g.R * s, ay = g.cy - g.R * c;
    return { x: Math.max(0, Math.min(g.w - w, right ? ax : ax - w)),
             y: Math.max(0, Math.min(g.h - h, up ? ay - h : ay)), w, h, up };
  });
  if (g.cull) {
    // collision on the text itself (box minus padding), a few px of slack: Tableau keeps labels that only touch
    const SLACK = 4;
    const hit = (a, b) => Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x) - 2 * LABEL_PAD > SLACK &&
                          Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y) > SLACK;
    const kept = [];
    points.map((p, i) => i).sort((a, b) => points[b].value - points[a].value || a - b).forEach(i => {
      if (!boxes[i]) return;
      if (kept.some(k => hit(k, boxes[i]))) boxes[i] = null;
      else kept.push(boxes[i]);
    });
  }
  return boxes;
}

/**
 * Labels Excel draws from cells (a pie that follows the slicers) sit outside each slice where Excel puts them, and
 * Excel never moves one off another: Tableau pushes crowded labels apart (two thin slices side by side), so do the
 * same – per point the shift, in chart fractions, from Excel's own place, or null. Excel keeps the shift when the
 * slicers resize the slices, so a label still follows its slice. Excel's place is estimated: the box just outside
 * the slice's mid-point, centred on the line from the centre; overlapping boxes are pushed apart along the axis
 * that needs the least move, inside the chart.
 * @param {any[]} points @param {{ cx: number, cy: number, R: number, w: number, h: number }} g px
 * @returns {({ x: number, y: number } | null)[]}
 */
function cellLabelShifts(points, g) {
  const GAP = 4, SPACE = 3;
  const total = points.reduce((s, p) => s + Math.abs(p.value || 0), 0) || 1;
  let acc = 0;
  const boxes = points.map(p => {
    const mid = (acc + Math.abs(p.value || 0) / 2) / total * 2 * Math.PI;   // clockwise from 12 o'clock
    acc += Math.abs(p.value || 0);
    if (!p.label || !p.label.length || !p.value) return null;
    let w = 0, h = 0;
    p.label.forEach(line => {
      const px = f => (f.size || 9) * 4 / 3;
      w = Math.max(w, line.reduce((t, x) => t + textWidthPx(x.text, x.font, px(x.font)), 0));
      h += (line.length ? Math.max(...line.map(x => px(x.font))) : px({})) * 1.2;
    });
    const s = Math.sin(mid), c = Math.cos(mid);
    const cx = g.cx + (g.R + GAP) * s + w / 2 * s, cy = g.cy - (g.R + GAP) * c - h / 2 * c;
    return { x: cx - w / 2, y: cy - h / 2, w, h, x0: cx - w / 2, y0: cy - h / 2 };
  });
  const live = boxes.filter(Boolean);
  const clamp = b => { b.x = Math.max(0, Math.min(g.w - b.w, b.x)); b.y = Math.max(0, Math.min(g.h - b.h, b.y)); };
  for (let pass = 0; pass < 30; pass++) {
    let moved = false;
    for (let i = 0; i < live.length; i++) for (let j = i + 1; j < live.length; j++) {
      const a = live[i], b = live[j];
      const ox = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x) + SPACE;
      const oy = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y) + SPACE;
      if (ox <= SPACE || oy <= SPACE) continue;
      moved = true;
      const k = ox <= oy ? "x" : "y", size = k === "x" ? "w" : "h";
      const [lo, hi] = a[k] + a[size] / 2 <= b[k] + b[size] / 2 ? [a, b] : [b, a];
      const d = Math.min(ox, oy) / 2;
      lo[k] -= d; hi[k] += d;
      clamp(lo); clamp(hi);
    }
    if (!moved) break;
  }
  return boxes.map(b => {
    if (!b || (Math.abs(b.x - b.x0) < 0.5 && Math.abs(b.y - b.y0) < 0.5)) return null;
    return { x: (b.x - b.x0) / g.w, y: (b.y - b.y0) / g.h };
  });
}

/* ── text ──────────────────────────────────────────────────────────────── */
/** a number in the sheet's own Excel format; null when there is no format or no number */
const formatNum = (v, nf) => v === null || v === undefined || !nf ? null : tfFormatNumber(v, nf);

/** a number shown the way Tableau showed a sibling value: same prefix / suffix, decimals, grouping, K/M/% scale */
function formatLike(dv, v) {
  const s = tfDvText(dv), n = tfDvNum(dv);
  const m = s.match(/^(\D*?)(-?\d[\d,]*(?:\.\d+)?)(.*)$/);
  if (!m || !n) return String(v);
  const shown = parseFloat(m[2].replace(/,/g, ""));
  const scale = shown ? n / shown : 1;
  const dec = (m[2].split(".")[1] || "").length;
  return m[1] + (v / scale).toLocaleString("en-US", { minimumFractionDigits: dec, maximumFractionDigits: dec,
    useGrouping: m[2].includes(",") || Math.abs(shown) < 1000 }) + m[3];
}

/**
 * runs [{text, props}] with <[field]> placeholders → lines [[{text, props}]]. Tableau usually splits a
 * placeholder over runs ("<", "[ds].[field]", ">"), so the joined text is rendered and every piece keeps the
 * props of the run it came from (a field value takes its field run's props).
 * @param {{ text: string, props?: any }[]} runs @param {(ref: FieldRef) => string} valueOf
 */
function renderTemplate(runs, valueOf) {
  const texts = runs.map(r => String(r.text || "").replace(/Æ\r?\n?/g, "\n"));
  const text = texts.join("");
  const owner = [];                                            // char index → run index
  texts.forEach((t, i) => { for (let k = 0; k < t.length; k++) owner.push(i); });
  const pieces = [];
  const push = (t, ri) => {
    const last = pieces[pieces.length - 1];
    if (last && last.ri === ri) last.text += t; else pieces.push({ text: t, ri });
  };
  const re = /<([^<>]+)>/g;
  let at = 0, m;
  while ((m = re.exec(text))) {
    for (let k = at; k < m.index; k++) push(text[k], owner[k]);
    const inner = m[1];
    let val = m[0];
    if (/^\[/.test(inner)) { const ref = tfParseFieldRef(inner); val = ref ? valueOf(ref) : ""; }
    else if (/^(sheet name|workbook name|page name|page count|page number)$/i.test(inner)) val = "";
    push(val, owner[m.index + 1]);
    at = m.index + m[0].length;
  }
  for (let k = at; k < text.length; k++) push(text[k], owner[k]);
  const lines = [[]];
  pieces.forEach(p => p.text.split(/\r?\n/).forEach((part, i) => {
    if (i) lines.push([]);
    if (part) lines[lines.length - 1].push({ text: part, props: runs[p.ri].props || {} });
  }));
  while (lines.length && !lines[lines.length - 1].length) lines.pop();
  return lines;
}
const plainText = lines => lines.map(line => line.map(x => x.text).join("")).join("\n");

/** summary rows selected in Tableau (getSelectedMarksAsync), matched on their dimension values */
function selectedRows(cols, data, tables) {
  const out = new Set();
  (tables || []).forEach(t => {
    const pairs = (t.columns || []).map((c, ti) => [cols.findIndex(x => x.name === c.fieldName), ti])
      .filter(([si]) => si >= 0 && !cols[si].numeric);
    if (!pairs.length) return;
    const keys = new Set((t.data || []).map(r => pairs.map(([, ti]) => tfDvText(r[ti])).join("\u0001")));
    data.forEach((r, i) => { if (keys.has(pairs.map(([si]) => tfDvText(r[si])).join("\u0001"))) out.add(i); });
  });
  return out;
}

/* ── model ─────────────────────────────────────────────────────────────── */
/**
 * One pie layer: its slices (value, colour, label, tooltip) per pane, its legend.
 * @param {SheetFormatter} fmt formatter scoped to the layer's pane @param {any} summary
 * @param {{ selectedTables?: any[] }} [opts]
 */
function buildPieSpec(fmt, summary, opts = {}) {
  // fieldId ("[ds].[pcto:sum:Qty:qk]") carries the aggregation – a % of total and its base measure can share
  // the fieldName "SUM(Qty)", so it wins over name matching when present
  const cols = (summary.columns || []).map((c, i) => {
    const name = c.fieldName || c.fieldId || `Col${i + 1}`;
    const byId = /:(qk|nk|ok)(:\d+)?\]$/.test(c.fieldId || "") ? tfParseFieldRef(c.fieldId) : null;
    return { name, ref: byId || fmt.matchName(name), numeric: /^(int|float|real|integer|number)$/i.test(String(c.dataType || "")) };
  });
  const colFor = ref => {
    if (!ref) return -1;
    let i = cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
    if (i < 0 && /^multiple values$/i.test(ref.name)) i = cols.findIndex(c => /^measure values$/i.test(c.name));
    if (i < 0 && /^measure names$/i.test(ref.name)) i = cols.findIndex(c => /^measure names$/i.test(c.name));
    return i;
  };

  const angleRef = fmt.encodingRefs(/^wedge-size$/i)[0] || null;
  const vi = colFor(angleRef);
  let marks = (summary.data || []).map((r, i) => ({ r, i }));
  // dual axis: summary data holds every layer's marks (a donut's hole is one extra row with the total);
  // this layer's rows are the ones where its own axis measure has a value
  const axisCol = fmt.axisRefs().length === 1 ? colFor(fmt.axisRefs()[0]) : -1;
  if (axisCol >= 0) marks = marks.filter(({ r }) => !tfIsNull(r[axisCol]));
  // both layers on one axis measure (MIN(0) twice): the other layer's rows are the ones without this layer's
  // colour member – a donut's hole total comes back as Country = Null
  const enc0 = fmt.colorEncoding(), ci0 = enc0 && !enc0.continuous ? colFor(enc0.ref) : -1;
  if (axisCol >= 0 && ci0 >= 0 && marks.some(({ r }) => !tfIsNull(r[ci0]))) marks = marks.filter(({ r }) => !tfIsNull(r[ci0]));
  if (vi >= 0) marks = marks.filter(({ r }) => { const v = tfDvNum(r[vi]); return v !== null && v > 0; });   // Tableau drops null / ≤ 0 wedges
  if (!marks.length) return null;

  // a field's value as Tableau shows it; % of total computed when the table calc isn't in the summary data
  const totals = {};
  const valueOf = r => ref => {
    const ci = colFor(ref);
    if (ci >= 0) return formatNum(tfDvNum(r[ci]), pieNumFmt(fmt, ref)) || tfDvText(r[ci]);
    if (!/^pcto/i.test(ref.deriv || "")) return "";
    const base = tfNorm(ref.name.split(":").pop());
    const bi = cols.findIndex(c => c.numeric && c.ref && tfNorm(c.ref.name) === base);
    if (bi < 0) return "";
    if (totals[bi] === undefined) totals[bi] = marks.reduce((s, m) => s + (tfDvNum(m.r[bi]) || 0), 0);
    // % of the whole table (Tableau's default scope); other "compute using" scopes are not replicated
    const nf = pieNumFmt(fmt, ref);
    const dec = nf ? ((nf.match(/\.(0+)/) || [, ""])[1].length) : 2;
    return totals[bi] ? (100 * (tfDvNum(r[bi]) || 0) / totals[bi]).toFixed(dec) + "%" : "";
  };

  // colour + legend
  const enc = fmt.colorEncoding();
  const ci = enc ? colFor(enc.ref) : -1;
  let colorAt = r => fmt.markColor();                          // no colour field: Marks → Color, else Tableau blue
  let legend = null;
  if (ci >= 0) {
    const pick = dv => tfIsNull(dv) ? null
      : enc.continuous ? (dv.nativeValue !== undefined ? dv.nativeValue : dv.value) : tfDvText(dv);
    const vals = marks.map(({ r }) => pick(r[ci]));
    const scale = tfBuildColorScale(fmt, enc, vals);
    if (scale) {
      const rgb = v => { const a = scale(v); return a ? a.slice(2) : null; };
      colorAt = r => rgb(pick(r[ci]));
      const title = fmt.captionFor(enc.ref, cols[ci].name);
      if (enc.continuous) {
        const nums = vals.filter(v => typeof v === "number");
        const lo = Math.min(...nums), hi = Math.max(...nums);
        legend = { title, items: [{ text: String(lo), color: rgb(lo) }, { text: String(hi), color: rgb(hi) }] };   // gradient: its two ends
      } else {
        const uniq = [...new Map(marks.map(({ r }) => [tfDvText(r[ci]), r[ci]])).values()].sort(tfNaturalCompare);
        legend = { title, items: uniq.map(dv => ({ text: tfDvText(dv), color: rgb(tfDvText(dv)) })) };
      }
    }
  }

  // panes: discrete fields on Rows / Columns → one pie each
  const hdr = shelf => cols.map((c, i) => c.ref && fmt.isHeaderField(c.ref) && fmt.shelfOf(c.ref) === shelf ? i : -1).filter(i => i >= 0);
  const rowHdr = hdr("rows"), colHdr = hdr("cols");
  const keyOf = (r, idx) => idx.map(i => tfDvText(r[i])).join(" | ");
  const hdrSet = new Set([...rowHdr, ...colHdr]);
  // a slice's name: its dimensions – not a field only another layer uses (a hole's "dummy" colour field)
  const sheetPanes = ((fmt.sheetModel || {}).panes) || [];
  const mineRefs = fmt.encodingRefs(/./);
  const otherLayerOnly = ref => !mineRefs.some(r => tfSameField(r, ref)) &&
    sheetPanes.some(p => (p.encodings || []).some(e => e.field && tfSameField(e.field, ref)));
  const dimIdx = cols.map((c, i) => !c.numeric && !hdrSet.has(i) && !/^measure values$/i.test(c.name) &&
                                    !(c.ref && !/^measure names$/i.test(c.name) && otherLayerOnly(c.ref)) ? i : -1).filter(i => i >= 0);

  const valueCaption = vi >= 0 ? fmt.captionFor(angleRef, cols[vi].name) : "Marks";
  const showLabels = fmt.markLabelsShown();
  const labelTpl = fmt.labelRunsTemplate(), tipTpl = fmt.tooltipTemplate();
  const selected = selectedRows(cols, summary.data || [], opts.selectedTables);
  // the label with each field's value as a part of its own (a slicer formula rebuilds those parts)
  const labelParts = r => {
    const parts = [];
    const lines = renderTemplate(labelTpl, ref => {
      const k = colFor(ref);
      parts.push({ ref, col: k, text: valueOf(r)(ref), nf: pieNumFmt(fmt, ref) });
      return `\u0002${parts.length - 1}\u0003`;
    });
    return { lines: lines.map(line => line.map(x => ({ text: x.text, font: pieFont(fmt.labelFont(x.props)) }))), parts };
  };
  const points = marks.map(({ r, i }) => ({
    row: i,
    rowKey: keyOf(r, rowHdr), colKey: keyOf(r, colHdr),
    category: dimIdx.map(k => tfDvText(r[k])).filter(Boolean).join(", ") || valueCaption,
    value: vi >= 0 ? tfDvNum(r[vi]) : 1,
    color: colorAt(r),
    selected: selected.has(i),
    label: showLabels ? renderTemplate(labelTpl, valueOf(r))
      .map(line => line.map(x => ({ text: x.text, font: pieFont(fmt.labelFont(x.props)) }))) : null,
    labelParts: showLabels ? labelParts(r) : null,
    tooltip: plainText(renderTemplate(tipTpl, valueOf(r)))
  }));
  const markSize = fmt.markSize();
  tfLog(`Pie: ${points.length} marks, angle=${vi >= 0 ? cols[vi].name : "none"}, colour=${ci >= 0 ? cols[ci].name : "none"}, ` +
        `labels=${showLabels}, selected=${selected.size}, size slider=${markSize === null ? "default" : markSize}`);
  const axis = shelf => ((fmt.sheetModel || {})[shelf] || []).some(r => r.type === "qk");
  return {
    columns: cols, data: summary.data || [], valueCaption, points, legend, markSize, valueCol: vi,
    cullLabels: fmt.markLabelsCulled(),
    rowKeys: [...new Set(points.map(p => p.rowKey))], colKeys: [...new Set(points.map(p => p.colKey))],
    stretch: { w: axis("cols"), h: axis("rows") },             // a continuous axis (a donut's MIN(0)) → the pane fills the zone
    paneOf: r => keyOf(r, rowHdr) + "\u0001" + keyOf(r, colHdr),   // the pane of any summary row (e.g. a hole row)
    colorRef: enc ? enc.ref : null,
    font: pieFont(fmt.labelFont({})),
    /** @type {any} */ hole: null, /** @type {any} */ inner: null, /** @type {any} */ centerText: null
  };
}

/**
 * Donut hole = the top layer of a dual axis (nothing on its card splitting it), drawn as Excel's doughnut hole.
 * Its label lives at the hole's level of detail (usually the grand total): Tableau returns one summary row per
 * pane for it, which carries the exact value – COUNTD and AVG too; without it the measures are re-aggregated
 * over the slices. @param {SheetFormatter} fmt formatter scoped to the hole's pane @param {any} pie
 */
function buildHoleSpec(fmt, pie) {
  const shown = fmt.markLabelsShown(), tpl = fmt.labelRunsTemplate();
  const ax = fmt.axisRefs()[0];
  const ai = ax ? pie.columns.findIndex(c => c.ref && tfSameField(c.ref, ax)) : -1;
  // the hole's own row: one the slices don't use (on one shared axis measure every row has an axis value)
  const sliceRows = new Set(pie.points.map(p => p.row));
  // map layers have no axis: every row is a candidate – the hole's own row is one no slice uses
  const axisRows = ai >= 0 ? pie.data.filter(r => !tfIsNull(r[ai])) : ax ? [] : pie.data;
  const spare = axisRows.filter(r => !sliceRows.has(pie.data.indexOf(r)));
  const own = new Map((spare.length ? spare : axisRows).map(r => [pie.paneOf(r), r]));
  const valueOf = pts => ref => {
    const k = pie.columns.findIndex(c => c.ref && tfSameField(c.ref, ref));
    if (k < 0) return "";
    const hit = own.get(pts[0].rowKey + "\u0001" + pts[0].colKey);
    if (hit) return formatNum(tfDvNum(hit[k]), pieNumFmt(fmt, ref)) || tfDvText(hit[k]);
    const dvs = pts.map(p => pie.data[p.row][k]);
    if (!pie.columns[k].numeric) { const u = [...new Set(dvs.map(tfDvText))]; return u.length === 1 ? u[0] : "*"; }
    const nums = dvs.map(tfDvNum).filter(v => v !== null);
    if (!nums.length) return "";
    const d = (ref.deriv || "").toLowerCase();
    // AVG / CNTD re-aggregated as a plain mean / sum of the slices, not from the raw rows
    const sum = nums.reduce((a, b) => a + b, 0);
    const v = d === "min" ? Math.min(...nums) : d === "max" ? Math.max(...nums) : d === "avg" ? sum / nums.length : sum;
    return formatNum(v, pieNumFmt(fmt, ref)) || formatLike(dvs.find(x => tfDvNum(x) !== null), v);
  };
  // hover text: the Tooltip editor if set; else, like Tableau, "Caption: value" for every field the hole's own
  // row has a value for, dimensions first, then by field name
  const tipFor = pts => {
    const hit = pts.length && own.get(pts[0].rowKey + "\u0001" + pts[0].colKey);
    if (!hit || fmt.hasCustomTooltip()) return plainText(renderTemplate(fmt.tooltipTemplate(), valueOf(pts)));
    return pie.columns.map((c, k) => ({ c, k })).filter(({ c, k }) => c.ref && !tfIsNull(hit[k]))
      .sort((a, b) => Number(a.c.numeric) - Number(b.c.numeric) || (a.c.numeric ? a.c.ref.inner.localeCompare(b.c.ref.inner) : 0))
      .map(({ c }) => fmt.captionFor(c.ref, c.name) + ": " + valueOf(pts)(c.ref)).join("\n");
  };
  return {
    markSize: fmt.markSize(),
    tipFor,
    color: fmt.markColor(),                                    // usually white = the background
    labelFor: pts => shown ? renderTemplate(tpl, valueOf(pts))
      .map(line => line.map(x => ({ text: x.text, font: pieFont(fmt.labelFont(x.props)) }))) : null,
    /** the label with each field's value as a part of its own (see buildPieSpec) */
    labelPartsFor: pts => {
      if (!shown) return null;
      const parts = [];
      const lines = renderTemplate(tpl, ref => {
        parts.push({ ref, col: pie.columns.findIndex(c => c.ref && tfSameField(c.ref, ref)), text: valueOf(pts)(ref), nf: pieNumFmt(fmt, ref) });
        return `\u0002${parts.length - 1}\u0003`;
      });
      return { lines: lines.map(line => line.map(x => ({ text: x.text, font: pieFont(fmt.labelFont(x.props)) }))), parts };
    }
  };
}

/**
 * Floating centre text: a dashboard text box, or another sheet (e.g. a total), whose centre lies inside a single
 * donut's hole (or a pie's circle) is drawn there with its Tableau text and fonts.
 * g: { stretch, D, holeD, cell, margin } px; dataOf(sheetName) → that sheet's summary data.
 * @returns {{ name: string | null, id: string, lines: any[] } | null} name: the sheet drawn there (null = a text box)
 */
function pieCenterText(model, dashName, sheetName, g, dataOf) {
  const dash = model && model.dashboards && model.dashboards[dashName];
  const z = dash && dash.width && dash.height && dash.zones.find(s => s.type === "worksheet" && s.name === sheetName);
  if (!z) return null;
  const px = (v, w) => v / 100000 * (w ? dash.width : dash.height);
  const W = px(z.w, 1), H = px(z.h, 0);
  const cx = px(z.x, 1) + g.margin + (g.stretch.w ? W - 2 * g.margin : Math.min(g.cell, W - 2 * g.margin)) / 2;
  const cy = px(z.y, 0) + (g.stretch.h ? H - g.margin : Math.min(g.cell, H - g.margin)) / 2;
  const R = (g.holeD || g.D) / 2;
  const hit = dash.zones.find(c => c !== z && !c.hidden && c.w && c.h && px(c.w, 1) < W && px(c.h, 0) < H &&
    (c.type === "text" ? !!(c.runs && tfZoneText(c)) : c.type === "worksheet" && c.name !== sheetName) &&
    Math.hypot(px(c.x + c.w / 2, 1) - cx, px(c.y + c.h / 2, 0) - cy) <= R);
  if (!hit) return null;
  if (hit.type === "text") return { name: null, id: hit.id, lines: renderTemplate(hit.runs, () => "")
    .map(line => line.map(x => ({ text: x.text, font: pieFont(tfMerge(TABLEAU_DEFAULTS.worksheet, x.props)) }))) };
  const sum = dataOf(hit.name);
  if (!sum || !sum.data || !sum.data.length) return null;
  const f = createSheetFormatter(model, hit.name);
  const refs = (sum.columns || []).map(c => /:(qk|nk|ok)(:\d+)?\]$/.test(c.fieldId || "") ? tfParseFieldRef(c.fieldId) : f.matchName(c.fieldName));
  const row = sum.data[0];
  const valueOf = ref => {
    const k = refs.findIndex(r => r && tfSameField(r, ref));
    return k < 0 ? "" : formatNum(tfDvNum(row[k]), f.numFmtFor(ref)) || tfDvText(row[k]);
  };
  const tpl = f.labelRunsTemplate();
  let lines = tpl.length ? renderTemplate(tpl, valueOf) : [];
  if (!lines.some(l => l.length)) lines = [[{ text: row.map(tfDvText).filter(Boolean).join(" "), props: {} }]];   // no label → its values
  return { name: hit.name, id: hit.id, lines: lines.map(line => line.map(x => ({ text: x.text, font: pieFont(f.labelFont(x.props)) }))) };
}

/**
 * The whole pie model of a sheet, or null when it isn't a pie / donut Tableau draws with one or two layers.
 * @param {FormatModel} model @param {string} sheetName @param {any} summary
 * @param {{ dashboardName: string, selectedTables?: any[], dataOf: (name: string) => any }} opts
 */
function buildPieModel(model, sheetName, summary, opts) {
  const fmt = createSheetFormatter(model, sheetName);
  const layers = fmt.pieLayers();
  if (!layers) return null;
  const scoped = pane => createSheetFormatter(model, sheetName, [pane]);
  const pie = buildPieSpec(scoped(layers.outer), summary, opts);
  if (!pie) return null;
  if (layers.hole) pie.hole = buildHoleSpec(scoped(layers.hole), pie);
  if (layers.inner) pie.inner = buildPieSpec(scoped(layers.inner), summary, opts);
  pie.gauge = !layers.inner && pieIsGauge(pie, (fmt.tableBackground() || "FFFFFFFF").slice(2));
  // headers hidden in Tableau (Show Header off on every discrete field of a shelf): the pies get no title from them
  const sm = fmt.sheetModel || {};
  const hiddenRefs = ((((sm.style || {}).rules || {}).label) || []).filter(r => r.attr === "display" && r.value === "false")
    .map(r => tfParseFieldRef(r.field)).filter(Boolean);
  const allHidden = shelf => { const dims = (sm[shelf] || []).filter(r => r && r.type !== "qk");
    return dims.length > 0 && dims.every(f => hiddenRefs.some(h => tfSameField(h, f))); };
  pie.hideRowHeaders = allHidden("rows");
  pie.hideColHeaders = allHidden("cols");
  const fit = fmt.fitMode(opts.dashboardName);                 // Entire View / Fit Width / Fit Height → the pane fills the zone
  if (/^(entire-view|fit-width)$/.test(fit || "")) pie.stretch.w = true;
  if (/^(entire-view|fit-height)$/.test(fit || "")) pie.stretch.h = true;
  if (pie.rowKeys.length * pie.colKeys.length === 1) {
    pie.centerText = pieCenterText(model, opts.dashboardName, sheetName,
      { stretch: pie.stretch, D: pieDiameterPx(pie.markSize), holeD: pie.hole ? pieDiameterPx(pie.hole.markSize) : 0,
        cell: PIE_CELL_PX, margin: ZONE_MARGIN }, opts.dataOf);
  }
  return pie;
}

/* ── output ───────────────────────────────────────────────────────────── */
/** a number format showing numbers the way Tableau showed this one ("$1.2M" → "$"#,##0.0,,"M"); null: not a number */
function excelFormatLike(text) {
  const m = String(text || "").match(/^(\D*?)(\d[\d,]*)(?:\.(\d+))?(\D*)$/);
  if (!m) return null;
  const q = s => (s ? `"${s.replace(/"/g, "")}"` : "");
  const unit = m[4].match(/^\s*([KMB%])/i);
  const scale = !unit ? "" : unit[1] === "%" ? "%" : ",".repeat({ k: 1, m: 2, b: 3 }[unit[1].toLowerCase()]);
  const suffix = unit && unit[1] === "%" ? m[4].replace("%", "") : m[4];
  const grouped = m[2].includes(",") || parseFloat(m[2]) < 1000;
  return `${q(m[1].replace(/-$/, ""))}${grouped ? "#,##0" : "0"}${m[3] ? "." + "0".repeat(m[3].length) : ""}${scale}${q(suffix)}`;
}

/**
 * A label's lines as formulas: each field part a TEXT() of its formula, in Tableau's number format (else one like
 * the value Tableau showed), the rest literal text. partFormula(part) → formula, or null (the part keeps its text).
 * @param {{ lines: any[][], parts: any[] }} lp @returns {{ expr: string, dynamic: boolean, text: string }[]}
 */
function labelLineFormulas(lp, partFormula) {
  const lit = s => `"${String(s).replace(/"/g, '""')}"`;
  return lp.lines.map(line => {
    let dynamic = false, text = "";
    const expr = line.map(x => x.text).join("").split(/(\u0002\d+\u0003)/).filter(Boolean).map(s => {
      const m = s.match(/^\u0002(\d+)\u0003$/);
      if (!m) { text += s; return lit(s); }
      const part = lp.parts[+m[1]];
      text += part.text;
      const v = partFormula(part);
      const nf = v && (part.nf || excelFormatLike(part.text));
      if (!nf) return lit(part.text);
      dynamic = true;
      return `TEXT(${v},${lit(nf)})`;
    }).join("&") || '""';
    return { expr, dynamic, text };
  });
}

/**
 * The slices on the chart data sheet – every summary column, the category / value Excel plots and Tableau's
 * label and tooltip text – so each chart is linked to real cells. Returns per pane the ranges its series uses.
 * A pie bound to the slicers (pie.slicerBinding) gets formulas: each slice's value, its label (Excel draws it from
 * the cell) and the donut's centre lines that show the slices' total (cells its text boxes link to).
 * @param {import("exceljs").Worksheet} ds @param {number} startRow 1-based @param {any} pie @param {string} name
 * @returns {{ panes: Map<string, { ref: any, innerRef: any, centre: (string | null)[] | null }>, nextRow: number }}
 */
function writePieData(ds, startRow, pie, name) {
  const nCols = pie.columns.length;
  const bound = pie.slicerBinding || null;
  ds.getCell(startRow, 1).value = name;
  ds.getCell(startRow, 1).font = { bold: true };
  // cell by cell: Row.values only takes an Array of ExcelJS's own realm
  const putRow = (r, values) => values.forEach((v, k) => { ds.getCell(r, k + 1).value = v; });
  putRow(startRow + 1, [...pie.columns.map(c => c.name), "Excel category", "Excel value", "Tableau label", "Tableau tooltip"]);
  ds.getRow(startRow + 1).font = { bold: true };
  const catCol = ds.getColumn(nCols + 1).letter, valCol = ds.getColumn(nCols + 2).letter, lblCol = ds.getColumn(nCols + 3).letter;
  const base = pie.valueCol >= 0 ? pie.columns[pie.valueCol] : null;
  /** a label part's formula: the slice's value or its share of the total, another measure of the slice; null = text */
  const fieldFormula = (part, at, total, p) => {
    const d = String(part.ref.deriv || "").toLowerCase();
    if (/^pcto/.test(d)) {
      if (base && base.ref && tfNorm(part.ref.name.split(":").pop()) === tfNorm(base.ref.name)) return `${at}/${total}`;
      // % of total of another measure: its own formula over the visible rows
      const f = part.col >= 0 && p ? bound.formulaFor([pie.data[p.row]], part.col, "cell") : null;
      if (f) return f;
    } else if (part.col === pie.valueCol) return at;
    else if (part.col >= 0 && pie.columns[part.col].numeric && p) {
      const f = bound.formulaFor([pie.data[p.row]], part.col, "cell");
      if (f) return f;
    }
    if (part.col >= 0 && pie.columns[part.col].numeric) bound.fixed++;
    return null;
  };
  let row = startRow + 2;
  const writeRows = (pts, layer) => {
    const first = row, last = row + pts.length - 1;
    const total = `SUMIF($${valCol}$${first}:$${valCol}$${last},">0")`;     // #N/A (a slice the slicers empty) left out
    pts.forEach(p => {
      const r = row++;
      putRow(r, [
        ...pie.data[p.row].map((dv, k) => pie.columns[k].numeric ? tfDvNum(dv) : tfDvText(dv)),
        p.category, p.value, p.label ? plainText(p.label) : "", p.tooltip]);
      if (!bound || layer !== pie) return;
      const at = `${valCol}${r}`;
      ds.getCell(at).value = { formula: bound.formulaFor([pie.data[p.row]], pie.valueCol, "chart"), result: p.value };
      bound.written++;
      const lines = p.labelParts ? labelLineFormulas(p.labelParts, part => fieldFormula(part, at, total, p)) : [];
      if (lines.some(l => l.dynamic))
        ds.getCell(`${lblCol}${r}`).value = { formula: `IF(ISNUMBER(${at}),${lines.map(l => l.expr).join("&CHAR(10)&")},"")`, result: plainText(p.label) };
    });
    return { sheet: ds.name, cat: `$${catCol}$${first}:$${catCol}$${last}`, val: `$${valCol}$${first}:$${valCol}$${last}`,
             lbl: `$${lblCol}$${first}:$${lblCol}$${last}`, total };
  };
  const single = pie.rowKeys.length * pie.colKeys.length === 1;
  const panes = new Map();
  pie.rowKeys.forEach(rk => pie.colKeys.forEach(ck => {
    const pts = pie.points.filter(p => p.rowKey === rk && p.colKey === ck);
    if (!pts.length) return;
    const ref = writeRows(pts, pie);
    const innerPts = pie.inner ? pie.inner.points.filter(p => p.rowKey === rk && p.colKey === ck) : [];
    const innerRef = innerPts.length ? writeRows(innerPts, pie.inner) : null;
    // the hole's label: a line showing the total of the slices' measure follows them; the rest keeps its text
    let centre = null;
    const lp = bound && pie.hole && !(single && pie.centerText) ? pie.hole.labelPartsFor(pts) : null;
    if (lp) {
      const lines = labelLineFormulas(lp, part => {
        if (part.col === pie.valueCol && /^(sum|cnt)$/.test(String(part.ref.deriv || "").toLowerCase())) return ref.total;
        if (part.col >= 0 && pie.columns[part.col].numeric) bound.fixed++;
        return null;
      });
      if (lines.some(l => l.dynamic)) {
        putRow(row, ["Centre text"]);
        centre = lines.map((l, i) => {
          if (!l.dynamic) return null;
          const cell = ds.getCell(row, i + 2);
          cell.value = { formula: l.expr, result: l.text };
          return sheetRef(ds.name, `$${cell.address.replace(/\d+$/, "")}$${row}`);
        });
        row++;
      }
    }
    panes.set(rk + "\u0001" + ck, { ref, innerRef, centre });
  }));
  return { panes, nextRow: row + 1 };
}

/**
 * Chart jobs for a pie sheet drawn in a block of w × h px: one exact-size chart per pane (a nested donut's inner
 * pie is a second, see-through chart over the ring's hole), Tableau's labels as text boxes, its tooltip on an
 * invisible wedge over each slice. Positions are px from the block's top-left.
 * @param {any} pie @param {{ w: number, h: number }} block @param {Map<string, any>} paneRefs (writePieData)
 * @param {{ name: string, bg: string, link: string }} o bg RRGGBB; link: where clicking a slice goes
 * @returns {{ x: number, y: number, w: number, h: number, xml: string, shapes: string | null, tips: any[], name: string }[]}
 */
function buildPieCharts(pie, block, paneRefs, o) {
  const M = ZONE_MARGIN;
  // Tableau's cell per pane, from the block's top-left inside the zone margin
  const fitW = (block.w - 2 * M) / pie.colKeys.length, fitH = (block.h - M) / pie.rowKeys.length;
  const pw = pie.stretch.w ? fitW : Math.min(PIE_CELL_PX, fitW);
  const ph = pie.stretch.h ? fitH : Math.min(PIE_CELL_PX, fitH);
  const textScale = FORMAT_CONFIG.tableauTextScale || (typeof window !== "undefined" && window.devicePixelRatio) || 1;
  // Tableau clips an oversized pie to its cell; Excel can't clip, so it fills the cell instead – and the hole /
  // inner ring shrink with it, so the ring keeps Tableau's proportions
  const full = pieDiameterPx(pie.markSize);
  const D = Math.min(full, pw, ph), k = D / full;
  const holeD = pie.hole ? Math.min(pieDiameterPx(pie.hole.markSize) * k, D) : 0;
  const innerD = pie.inner ? Math.min(pieDiameterPx(pie.inner.markSize) * k, D) : 0;
  const single = pie.rowKeys.length * pie.colKeys.length === 1;
  const labelsOf = (pts, R, cull) => {
    const boxes = pieLabelLayout(pts, { cx: pw / 2, cy: ph / 2, R, w: pw, h: ph, cull, textScale });
    return pts.map((p, k) => boxes[k] && { ...boxes[k], lines: p.label, anchor: boxes[k].up ? "b" : "t" }).filter(Boolean);
  };
  // Excel draws slice 1 from 12 o'clock clockwise with |value|; DrawingML angles are 60000ths of a degree
  // clockwise from 3 o'clock
  const tipsOf = (pts, d, hole, size) => {
    const x = (size.w - d) / 2, y = (size.h - d) / 2;
    const total = pts.reduce((s, p) => s + Math.abs(p.value || 0), 0);
    const ang = a => (((a % 360) + 360) % 360) * 60000;
    let a = -90;
    return pts.map(p => {
      const st = a;
      a += total ? 360 * Math.abs(p.value || 0) / total : 0;
      if (a === st || !p.tooltip) return null;
      return { x, y, w: d, h: d, prst: hole ? "blockArc" : "pie",
               adj: hole ? [ang(st), ang(a), 50000 * (1 - hole / d)] : [ang(st), ang(a)], text: p.tooltip, link: o.link };
    }).filter(Boolean);
  };
  // bound to the slicers: the slices change size, so Excel draws the labels from their cells, the centre lines
  // showing the total link to theirs, and the tooltip wedges (cut at Tableau's angles) are left out
  const bound = !!pie.slicerBinding || !!pie.switched;
  /** centre lines stacked around the middle, one text box each; a line with a cell shows that cell's text */
  const centreBoxes = (lines, links) => {
    const lh = lines.map(l => Math.max(...l.map(x => (x.font && x.font.size) || 9), 1) * 1.25 * 96 / 72);
    let y = ph / 2 - lh.reduce((a, b) => a + b, 0) / 2;
    return lines.map((l, i) => { const b = { x: 0, y, w: pw, h: lh[i], lines: [l], anchor: "ctr", link: links[i] || null }; y += lh[i]; return b; });
  };
  const out = [];
  pie.rowKeys.forEach((rk, ri) => pie.colKeys.forEach((ck, cj) => {
    const pts = pie.points.filter(p => p.rowKey === rk && p.colKey === ck);
    const refs = paneRefs.get(rk + "\u0001" + ck);
    if (!pts.length || !refs) return;
    const innerPts = innerD && refs.innerRef ? pie.inner.points.filter(p => p.rowKey === rk && p.colKey === ck) : [];
    // labels of the two rings are culled per ring, not against each other
    // a gauge's slice in the background colour hides part of the dial: see-through (on any background); its label
    // and tooltip stay as Tableau shows them
    const bgU = String(o.bg || "").toUpperCase();
    let shown = pts.map(p => pie.gauge && p.color && String(p.color).toUpperCase() === bgU ? { ...p, hidden: true } : p);
    // a gauge's label that is the same on every slice (its name: "BNI / Kredit Konsumer / Limit : 80%") is
    // shown once by Tableau, above the dial: kept on the visible slice nearest 12 o'clock, dropped on the others
    if (pie.gauge) {
      const textOf = p => (p.label ? JSON.stringify((p.label || []).map(l => (Array.isArray(l) ? l : [l]).map(x => (x && x.text !== undefined ? x.text : x)))) : "");
      const labelled = shown.map((p, i) => i).filter(i => textOf(shown[i]));
      if (labelled.length > 1 && labelled.every(i => textOf(shown[i]) === textOf(shown[labelled[0]]))) {
        const total = shown.reduce((t, p) => t + Math.abs(p.value || 0), 0) || 1;
        let a = 0;
        const mid = shown.map(p => { const st = a; a += 360 * Math.abs(p.value || 0) / total; return (st + a) / 2; });
        const top = i => Math.min(mid[i], 360 - mid[i]);
        const pick = labelled.filter(i => !shown[i].hidden && Math.abs(shown[i].value || 0) > 0)
                             .sort((x, y) => top(x) - top(y))[0];
        if (pick !== undefined) shown = shown.map((p, i) => (i === pick || !p.label ? p : { ...p, label: null }));
      }
    }
    const texts = bound ? [] : [...labelsOf(shown, D / 2, pie.cullLabels), ...(innerPts.length ? labelsOf(innerPts, innerD / 2, pie.inner.cullLabels) : [])];
    // Excel's hole is see-through (it shows the chart background): a hole in another colour is a filled circle
    // under the labels. Bound to the slicers, a donut is a pie under such a circle: Excel puts a pie's labels outside
    // it, as Tableau does, but a doughnut's on the ring
    const holeAsCircle = bound && holeD && !innerPts.length;
    if (holeD && (holeAsCircle || pie.hole.color.toUpperCase() !== o.bg.toUpperCase()))
      texts.unshift({ x: (pw - holeD) / 2, y: (ph - holeD) / 2, w: holeD, h: holeD, shape: "ellipse", fill: pie.hole.color, lines: [] });
    // centre: a floating text box / sheet over the hole wins over the hole's own label (Tableau draws it on top)
    const centre = (single && pie.centerText && pie.centerText.lines) || (pie.hole && pie.hole.labelFor(pts));
    if (centre && centre.length && !innerPts.length) {
      if (refs.centre) texts.push(...centreBoxes(centre, refs.centre));
      else texts.push({ x: 0, y: 0, w: pw, h: ph, lines: centre, anchor: "ctr" });
    }
    const labelFont = (pts.find(p => p.label && p.label.length && p.label[0].length) || { label: [[{ font: pie.font }]] }).label[0][0].font;
    const cellLabels = bound && pts.some(p => p.label) ? { range: sheetRef(refs.ref.sheet, refs.ref.lbl), texts: pts.map(p => p.label ? plainText(p.label) : ""), font: labelFont,
      shifts: !pie.switched && (holeAsCircle || !holeD) ? cellLabelShifts(pts, { cx: pw / 2, cy: ph / 2, R: D / 2, w: pw, h: ph }) : [] } : null;
    const x = M + cj * pw, y = ri * ph;
    const spec = {
      seriesName: pie.valueCaption, categories: pts.map(p => p.category), values: pts.map(p => p.value), ref: refs.ref,
      title: [pie.hideRowHeaders ? null : rk, pie.hideColHeaders ? null : ck].filter(Boolean).join(" | ") || null, titleFont: pie.font, font: pie.font, bg: o.bg,
      plot: { x: (1 - D / pw) / 2, y: (1 - D / ph) / 2, w: D / pw, h: D / ph },
      holeSize: holeAsCircle ? null : holeD ? 100 * holeD / D : innerPts.length ? 100 * innerD / D : null,
      points: shown.map(p => ({ color: p.hidden ? null : p.color, selected: p.selected })), hasShapes: texts.length > 0, cellLabels
    };
    const holeTip = holeD && pie.hole.tipFor(pts);
    out.push({ x, y, w: pw, h: ph, name: o.name, xml: pieChartXml(spec),
               shapes: texts.length ? userShapesXml(texts, { w: pw, h: ph }) : null,
               // a gauge stays a plain, editable Excel chart: no tooltip wedges grouped over it
               tips: pie.gauge || bound ? [] : [...tipsOf(shown, D, holeD || (innerPts.length ? innerD : 0), { w: pw, h: ph }),
                      ...(holeTip ? [{ x: (pw - holeD) / 2, y: (ph - holeD) / 2, w: holeD, h: holeD, prst: "ellipse", text: holeTip, link: o.link }] : [])] });
    if (!innerPts.length) return;
    // nested donut: the inner pie is its own see-through chart exactly over the ring's hole (drawn after = on top);
    // its labels stay in the ring chart so they can sit over the ring
    const innerTexts = centre && centre.length ? [{ x: 0, y: 0, w: innerD, h: innerD, lines: centre, anchor: "ctr" }] : [];
    const innerSpec = {
      seriesName: pie.inner.valueCaption, categories: innerPts.map(p => p.category), values: innerPts.map(p => p.value),
      ref: refs.innerRef, title: null, font: pie.font, bg: null, plot: { x: 0, y: 0, w: 1, h: 1 }, hasShapes: innerTexts.length > 0,
      points: innerPts.map(p => ({ color: p.color, selected: p.selected }))
    };
    out.push({ x: x + (pw - innerD) / 2, y: y + (ph - innerD) / 2, w: innerD, h: innerD, name: o.name + " (inner)",
               xml: pieChartXml(innerSpec), shapes: innerTexts.length ? userShapesXml(innerTexts, { w: innerD, h: innerD }) : null,
               tips: tipsOf(innerPts, innerD, 0, { w: innerD, h: innerD }) });
  }));
  return out;
}



/* ═══ charts/model/gauge.js ════════════════════════════════════════════════════════════════════════ */
/* Gauges. Tableau has no gauge mark, so authors build one of two ways:
 *  (a) from pies – a value slice, a remainder, often a thin "pointer" slice (the needle) and a slice painted the
 *      background colour that hides the bottom half, with a circle on top as the hole. That goes through the pie
 *      path (pieLayers accepts the coloured circle hole; a slice in the background colour is written see-through).
 *  (b) from map layers or polygons – band arcs drawn as polygons from MAKEPOINT(r·SIN θ, −r·COS θ) points, a
 *      needle line, tick lines and text marks. The summary data only carries the points' inputs (level, value,
 *      point …), not their positions, so this module recomputes every layer's points by evaluating the sheet's
 *      Tableau calculations row by row, measures the dial (centre, radii, band angles), the needle and the labels
 *      from that geometry, and writes Excel's own gauge: a doughnut dial – the bands plus a see-through remainder –
 *      and a pie needle, both turned to the dial's start angle, with Tableau's labels as text boxes.
 * Anything it cannot evaluate or recognise returns null, and the sheet is exported exactly as before. */

/* ── a small Tableau calculation evaluator ─────────────────────────────── */
/* Enough of the calculation language for geometry: numbers, strings, booleans, NULL, fields, parameters,
 * + − * / % ^, comparisons, AND / OR / NOT, IN (…), IF / ELSEIF / ELSE, CASE … WHEN [IN (…)], IIF, LOD braces
 * {FIXED …: expr} (one value: the gauge's KPI), the math functions, MIN / MAX, aggregations over one mark,
 * MAKEPOINT and MAKELINE. Anything else throws, and the caller gives up. */
class TfCalcError extends Error {}

function tfCalcTokens(src) {
  const out = [];
  const s = String(src);
  let i = 0;
  while (i < s.length) {
    const c = s[i];
    if (/\s/.test(c)) { i++; continue; }
    if (c === "/" && s[i + 1] === "/") { while (i < s.length && s[i] !== "\n") i++; continue; }        // comment
    if (c === "/" && s[i + 1] === "*") { const e = s.indexOf("*/", i + 2); i = e < 0 ? s.length : e + 2; continue; }
    if (/[0-9]/.test(c) || (c === "." && /[0-9]/.test(s[i + 1] || ""))) {
      const m = /^(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?/.exec(s.slice(i));
      out.push({ t: "num", v: parseFloat(m[0]) }); i += m[0].length; continue;
    }
    if (c === "'" || c === '"') {
      let j = i + 1, v = "";
      while (j < s.length) { if (s[j] === c && s[j + 1] === c) { v += c; j += 2; continue; } if (s[j] === c) break; v += s[j++]; }
      out.push({ t: "str", v }); i = j + 1; continue;
    }
    if (c === "[") {                                           // [field] or [ds].[field]; ]] escapes ]
      const parts = [];
      let j = i;
      for (;;) {
        let k = j + 1, v = "";
        while (k < s.length) { if (s[k] === "]" && s[k + 1] === "]") { v += "]"; k += 2; continue; } if (s[k] === "]") break; v += s[k++]; }
        parts.push(v); j = k + 1;
        if (s[j] === "." && s[j + 1] === "[") { j++; continue; }
        break;
      }
      out.push({ t: "field", v: parts }); i = j; continue;
    }
    const two = s.slice(i, i + 2);
    if (["<=", ">=", "<>", "!=", "=="].includes(two)) { out.push({ t: "op", v: two === "==" ? "=" : two === "!=" ? "<>" : two }); i += 2; continue; }
    if ("+-*/%^=<>(),{}:".includes(c)) { out.push({ t: c === "(" || c === ")" || c === "," || c === "{" || c === "}" || c === ":" ? c : "op", v: c }); i++; continue; }
    const m = /^[A-Za-z_][A-Za-z0-9_]*/.exec(s.slice(i));
    if (!m) throw new TfCalcError(`unexpected "${c}"`);
    out.push({ t: "id", v: m[0].toUpperCase() }); i += m[0].length;
  }
  return out;
}

/** formula text → AST @param {string} src */
function tfCalcParse(src) {
  const tk = tfCalcTokens(src);
  let p = 0;
  const peek = () => tk[p] || { t: "eof" };
  const isKw = (k) => peek().t === "id" && peek().v === k;
  const take = (t, v) => {
    const x = peek();
    if (x.t !== t || (v !== undefined && x.v !== v)) throw new TfCalcError(`expected ${v || t}`);
    p++; return x;
  };
  const kw = k => { if (!isKw(k)) throw new TfCalcError(`expected ${k}`); p++; };
  const list = () => { take("("); const a = []; if (peek().t !== ")") { a.push(expr()); while (peek().t === ",") { p++; a.push(expr()); } } take(")"); return a; };
  const expr = () => or();
  const or = () => { let l = and(); while (isKw("OR")) { p++; l = { k: "or", l, r: and() }; } return l; };
  const and = () => { let l = not(); while (isKw("AND")) { p++; l = { k: "and", l, r: not() }; } return l; };
  const not = () => isKw("NOT") ? (p++, { k: "not", a: not() }) : cmp();
  const cmp = () => {
    let l = add();
    for (;;) {
      if (peek().t === "op" && ["=", "<>", "<", "<=", ">", ">="].includes(peek().v)) { const op = tk[p++].v; l = { k: "bin", op, l, r: add() }; continue; }
      if (isKw("IN")) { p++; l = { k: "in", a: l, list: list() }; continue; }
      return l;
    }
  };
  const add = () => { let l = mul(); while (peek().t === "op" && (peek().v === "+" || peek().v === "-")) { const op = tk[p++].v; l = { k: "bin", op, l, r: mul() }; } return l; };
  const mul = () => { let l = pow(); while (peek().t === "op" && ["*", "/", "%"].includes(peek().v)) { const op = tk[p++].v; l = { k: "bin", op, l, r: pow() }; } return l; };
  const pow = () => { const l = unary(); if (peek().t === "op" && peek().v === "^") { p++; return { k: "bin", op: "^", l, r: pow() }; } return l; };
  const unary = () => {
    if (peek().t === "op" && (peek().v === "-" || peek().v === "+")) { const op = tk[p++].v; const a = unary(); return op === "-" ? { k: "neg", a } : a; }
    return primary();
  };
  const primary = () => {
    const x = peek();
    if (x.t === "num") { p++; return { k: "lit", v: x.v }; }
    if (x.t === "str") { p++; return { k: "lit", v: x.v }; }
    if (x.t === "field") { p++; return { k: "field", parts: x.v }; }
    if (x.t === "(") { p++; const e = expr(); take(")"); return e; }
    if (x.t === "{") {                                         // LOD: one value for the whole view
      p++;
      if (isKw("FIXED") || isKw("INCLUDE") || isKw("EXCLUDE")) {
        p++;
        while (peek().t !== ":" && peek().t !== "eof") p++;   // the dimensions do not matter for one gauge
        take(":");
      }
      const e = expr(); take("}"); return { k: "lod", a: e };
    }
    if (x.t === "id") {
      p++;
      if (x.v === "TRUE") return { k: "lit", v: true };
      if (x.v === "FALSE") return { k: "lit", v: false };
      if (x.v === "NULL") return { k: "lit", v: null };
      if (x.v === "IF") {
        const arms = [];
        const c = expr(); kw("THEN"); arms.push([c, expr()]);
        let other = { k: "lit", v: null };
        for (;;) {
          if (isKw("ELSEIF")) { p++; const c2 = expr(); kw("THEN"); arms.push([c2, expr()]); continue; }
          if (isKw("ELSE")) { p++; other = expr(); }
          kw("END"); break;
        }
        return { k: "if", arms, other };
      }
      if (x.v === "CASE") {
        const subject = expr();
        const arms = [];
        let other = { k: "lit", v: null };
        while (isKw("WHEN")) {
          p++;
          let match;
          if (isKw("IN")) { p++; match = { inList: list() }; }
          else { const vals = [expr()]; while (peek().t === ",") { p++; vals.push(expr()); } match = { vals }; }
          kw("THEN"); arms.push([match, expr()]);
        }
        if (isKw("ELSE")) { p++; other = expr(); }
        kw("END");
        return { k: "case", subject, arms, other };
      }
      if (peek().t === "(") return { k: "call", f: x.v, args: list() };
      throw new TfCalcError(`unknown word ${x.v}`);
    }
    throw new TfCalcError("unexpected end of formula");
  };
  const ast = expr();
  if (peek().t !== "eof") throw new TfCalcError("unexpected text after the formula");
  return ast;
}

const TF_CALC_NUM = v => (typeof v === "number" && isFinite(v) ? v : v === null || v === undefined ? null : typeof v === "boolean" ? +v : NaN);
const TF_CALC_FN = {
  PI: () => Math.PI, SIN: Math.sin, COS: Math.cos, TAN: Math.tan, ASIN: Math.asin, ACOS: Math.acos, ATAN: Math.atan,
  ATAN2: Math.atan2, SQRT: Math.sqrt, ABS: Math.abs, EXP: Math.exp, LN: Math.log, SIGN: Math.sign, FLOOR: Math.floor,
  CEILING: Math.ceil, INT: Math.trunc, FLOAT: v => v, SQUARE: v => v * v, RADIANS: v => v * Math.PI / 180,
  DEGREES: v => v * 180 / Math.PI, POWER: (a, b) => Math.pow(a, b), LOG: (a, b) => b === undefined ? Math.log10(a) : Math.log(a) / Math.log(b),
  ROUND: (a, n) => { const k = Math.pow(10, n || 0); return Math.round(a * k) / k; },
  DIV: (a, b) => Math.trunc(a / b)
};
const TF_CALC_AGG = new Set(["SUM", "AVG", "MEDIAN", "ATTR", "MIN", "MAX", "AGG", "TOTAL", "WINDOW_SUM", "WINDOW_AVG", "WINDOW_MAX", "WINDOW_MIN"]);

/**
 * Evaluate an AST. env.field(parts) → a field's value for the current mark (throws when it cannot be found).
 * One mark at a time: an aggregation of one mark is its value; MIN / MAX with two or more arguments compare them.
 * @param {any} n @param {{ field: (parts: string[]) => any }} env
 */
function tfCalcEval(n, env) {
  const ev = x => tfCalcEval(x, env);
  switch (n.k) {
    case "lit": return n.v;
    case "field": return env.field(n.parts);
    case "lod": return ev(n.a);
    case "neg": { const v = TF_CALC_NUM(ev(n.a)); return v === null ? null : -v; }
    case "not": { const v = ev(n.a); return v === null ? null : !v; }
    case "and": { const l = ev(n.l); if (l === false) return false; const r = ev(n.r); return l === null || r === null ? (r === false ? false : null) : !!(l && r); }
    case "or": { const l = ev(n.l); if (l === true) return true; const r = ev(n.r); return l === null || r === null ? (r === true ? true : null) : !!(l || r); }
    case "in": { const a = ev(n.a); return a === null ? null : n.list.some(x => ev(x) === a); }
    case "bin": {
      const l = ev(n.l), r = ev(n.r);
      if (l === null || r === null) return null;
      if (["=", "<>", "<", "<=", ">", ">="].includes(n.op)) {
        const a = typeof l === "string" || typeof r === "string" ? String(l) : TF_CALC_NUM(l);
        const b = typeof l === "string" || typeof r === "string" ? String(r) : TF_CALC_NUM(r);
        switch (n.op) { case "=": return a === b; case "<>": return a !== b; case "<": return a < b; case "<=": return a <= b; case ">": return a > b; default: return a >= b; }
      }
      if (n.op === "+" && (typeof l === "string" || typeof r === "string")) return String(l) + String(r);
      const a = TF_CALC_NUM(l), b = TF_CALC_NUM(r);
      if (isNaN(a) || isNaN(b)) throw new TfCalcError(`not a number in ${n.op}`);
      switch (n.op) {
        case "+": return a + b; case "-": return a - b; case "*": return a * b;
        case "/": return b === 0 ? null : a / b; case "%": return b === 0 ? null : a % b; default: return Math.pow(a, b);
      }
    }
    case "if": {
      for (const [c, v] of n.arms) if (ev(c) === true) return ev(v);
      return ev(n.other);
    }
    case "case": {
      const s = ev(n.subject);
      for (const [m, v] of n.arms) {
        const hit = m.inList ? m.inList.some(x => ev(x) === s) : m.vals.some(x => ev(x) === s);
        if (s !== null && hit) return ev(v);
      }
      return ev(n.other);
    }
    case "call": {
      const f = n.f, args = n.args;
      if (f === "IIF") { const c = ev(args[0]); return c === true ? ev(args[1]) : c === false ? ev(args[2]) : (args[3] ? ev(args[3]) : null); }
      if (f === "IFNULL") { const v = ev(args[0]); return v === null ? ev(args[1]) : v; }
      if (f === "ZN") { const v = ev(args[0]); return v === null ? 0 : v; }
      if (f === "ISNULL") return ev(args[0]) === null;
      if (f === "MAKEPOINT") {
        const lat = TF_CALC_NUM(ev(args[0])), lon = TF_CALC_NUM(ev(args[1]));
        return lat === null || lon === null || isNaN(lat) || isNaN(lon) ? null : { point: [lon, lat] };   // x = longitude, y = latitude
      }
      if (f === "MAKELINE") { const a = ev(args[0]), b = ev(args[1]); return a && b && a.point && b.point ? { line: [a.point, b.point] } : null; }
      if ((f === "MIN" || f === "MAX") && args.length > 1) {
        const vals = args.map(a => TF_CALC_NUM(ev(a)));
        if (vals.some(v => v === null)) return null;
        return f === "MIN" ? Math.min(...vals) : Math.max(...vals);
      }
      if (TF_CALC_AGG.has(f) && args.length === 1) return ev(args[0]);
      if (f === "COUNT" || f === "COUNTD") return ev(args[0]) === null ? 0 : 1;
      if (TF_CALC_FN[f]) {
        const vals = args.map(a => TF_CALC_NUM(ev(a)));
        if (vals.some(v => v === null)) return null;
        if (vals.some(v => isNaN(v))) throw new TfCalcError(`${f}: not a number`);
        const v = TF_CALC_FN[f](...vals);
        return typeof v === "number" && !isFinite(v) ? null : v;
      }
      throw new TfCalcError(`function ${f} is not supported`);
    }
    default: throw new TfCalcError("unknown expression");
  }
}

/* ── the sheet's fields for the evaluator ───────────────────────────────── */
/**
 * Resolves a field reference for one summary row: a parameter's current value, the row's own column, a calculated
 * field evaluated from its formula, or – for an aggregate the row does not carry (a gauge's KPI) – the field's one
 * value anywhere in the summary data.
 * @param {FormatModel} model @param {any} summary @param {any[]} cols [{ name, ref, numeric }]
 */
function tfCalcContext(model, summary, cols) {
  const fields = model.fields || {};
  // a calculation's name can exist in several data sources with different formulas: the field of the given data
  // source wins, then any data source's; a parameter is never a field
  const fieldInfo = (name, ds) => {
    const lower = String(name).toLowerCase();
    const own = ds && fields[ds + "|" + lower];
    if (own && !own.param) return own;
    const k = Object.keys(fields).find(key => key.endsWith("|" + lower) && fields[key] && !fields[key].param);
    return k ? fields[k] : null;
  };
  const paramInfo = name => fields["Parameters|" + String(name).toLowerCase()] || null;
  const strip = s => String(s || "").replace(/^(sum|avg|min|max|agg|attr|median|count|countd|none|usr|clct|year|month|day|quarter|week)\(\s*(.*)\)$/i, "$2");
  const colIndex = new Map();
  const colFor = (name, ds) => {
    const key = String(ds || "") + "|" + String(name).toLowerCase();
    if (colIndex.has(key)) return colIndex.get(key);
    const info = fieldInfo(name, ds);
    const wanted = new Set([String(name).toLowerCase(), info && info.caption ? String(info.caption).toLowerCase() : null].filter(Boolean));
    let i = cols.findIndex(c => c.ref && wanted.has(String(c.ref.name).toLowerCase()));
    if (i < 0) i = cols.findIndex(c => wanted.has(String(c.name).toLowerCase()) || wanted.has(strip(c.name).toLowerCase()));
    colIndex.set(key, i);
    return i;
  };
  const scalar = new Map();
  const anyValue = ci => {
    if (scalar.has(ci)) return scalar.get(ci);
    const r = (summary.data || []).find(row => !tfIsNull(row[ci]));
    const v = r ? (cols[ci].numeric ? tfDvNum(r[ci]) : tfDvText(r[ci])) : null;
    scalar.set(ci, v);
    return v;
  };
  const asts = new Map();
  const astOf = (name, formula) => {
    if (!asts.has(name)) asts.set(name, tfCalcParse(formula));
    return asts.get(name);
  };
  const formulaOf = info => {
    const f = info && (info.formulaFull || info.formula);
    return f && !/…$/.test(f) ? f : null;
  };
  /** an evaluator bound to one summary row (null: no row – the KPI-level values only) */
  const forRow = (row, dsHint) => {
    const memo = new Map();
    const dsStack = [dsHint || ""];
    const env = {
      depth: 0,
      field(parts) {
        const name = parts[parts.length - 1];
        const ds = parts.length > 1 ? parts[0] : dsStack[dsStack.length - 1];
        if (/^parameters$/i.test(ds) || (!ds && paramInfo(name) && !fieldInfo(name))) {
          const p = paramInfo(name);
          if (!p) throw new TfCalcError(`unknown parameter ${name}`);
          const v = p.constant !== undefined ? p.constant : p.value;
          const n = typeof v === "number" ? v : parseFloat(String(v).replace(/^"|"$/g, ""));
          return isFinite(n) ? n : String(v).replace(/^"|"$/g, "");
        }
        const mkey = ds + "|" + name;
        if (memo.has(mkey)) return memo.get(mkey);
        const ci = colFor(name, ds);
        const info = fieldInfo(name, ds), formula = formulaOf(info);
        const located = !!formula && (info.datatype === "spatial" || /MAKEPOINT|MAKELINE/i.test(formula));
        let v;
        if (row && ci >= 0 && !tfIsNull(row[ci]) && !located) v = cols[ci].numeric ? tfDvNum(row[ci]) : tfDvText(row[ci]);
        else {
          if (formula) {
            if (++env.depth > 40) throw new TfCalcError("calculations nest too deep");
            dsStack.push(info.ds || ds);
            try { v = tfCalcEval(astOf((info.ds || ds) + "|" + name, formula), env); }
            finally { dsStack.pop(); env.depth--; }
          } else if (ci >= 0) v = anyValue(ci);              // an aggregate the row does not carry: its one value
          else throw new TfCalcError(`field ${name} not found`);
        }
        memo.set(mkey, v);
        return v;
      }
    };
    return env;
  };
  return { forRow, colFor, fieldInfo, formulaOf, paramInfo };
}

/* ── geometry ───────────────────────────────────────────────────────────── */
/** least-squares circle through points (Kåsa): { cx, cy } @param {number[][]} pts */
function tfFitCircleCentre(pts) {
  let sx = 0, sy = 0, sxx = 0, syy = 0, sxy = 0, sxz = 0, syz = 0, sz = 0;
  const n = pts.length;
  for (const [x, y] of pts) {
    const z = x * x + y * y;
    sx += x; sy += y; sxx += x * x; syy += y * y; sxy += x * y; sxz += x * z; syz += y * z; sz += z;
  }
  // solve [sxx sxy sx; sxy syy sy; sx sy n] [D E F]ᵀ = −[sxz syz sz]ᵀ  (x² + y² + Dx + Ey + F = 0)
  const A = [[sxx, sxy, sx], [sxy, syy, sy], [sx, sy, n]], b = [-sxz, -syz, -sz];
  const det = m => m[0][0] * (m[1][1] * m[2][2] - m[1][2] * m[2][1]) - m[0][1] * (m[1][0] * m[2][2] - m[1][2] * m[2][0]) + m[0][2] * (m[1][0] * m[2][1] - m[1][1] * m[2][0]);
  const d = det(A);
  if (Math.abs(d) < 1e-12) return null;
  const col = k => A.map((r, i) => r.map((v, j) => j === k ? b[i] : v));
  const D = det(col(0)) / d, E = det(col(1)) / d;
  return { cx: -D / 2, cy: -E / 2 };
}

/**
 * A dial's centre from its outlines: three points along one arc define the arc's centre exactly; triples that
 * straddle a corner (outer arc → inner arc) do not, so the median of all triples' centres wins.
 * @param {number[][][]} paths ordered outlines @returns {{ cx: number, cy: number } | null}
 */
function tfArcCentreGuess(paths) {
  const xs = [], ys = [];
  for (const path of paths) {
    const n = path.length;
    if (n < 6) continue;
    const k = Math.max(1, Math.floor(n / 40));                 // points a little apart: less rounding noise
    for (let i = 0; i + 2 * k < n; i += k) {
      const [ax, ay] = path[i], [bx, by] = path[i + k], [cx, cy] = path[i + 2 * k];
      const d = 2 * (ax * (by - cy) + bx * (cy - ay) + cx * (ay - by));
      const span = Math.hypot(cx - ax, cy - ay);
      if (Math.abs(d) < 1e-12 * Math.max(1, span * span)) continue;   // straight: no circle
      const a2 = ax * ax + ay * ay, b2 = bx * bx + by * by, c2 = cx * cx + cy * cy;
      xs.push((a2 * (by - cy) + b2 * (cy - ay) + c2 * (ay - by)) / d);
      ys.push((a2 * (cx - bx) + b2 * (ax - cx) + c2 * (bx - ax)) / d);
    }
  }
  if (xs.length < 3) return null;
  const med = v => { const s = v.slice().sort((p, q) => p - q); return s[Math.floor(s.length / 2)]; };
  return { cx: med(xs), cy: med(ys) };
}

/**
 * The centre of a dial: its band outlines lie on two concentric circles (outer and inner edge) – a plain circle
 * fit is biased on such partial arcs. Starting from the algebraic fit, the origin and the points' centroid, each
 * guess is refined (Gauss–Newton) so every point lies on one of the two circles; the best fit wins.
 * @param {number[][]} pts @returns {{ cx: number, cy: number, rms: number } | null}
 */
function tfFitDialCentre(pts, paths) {
  if (pts.length < 3) return null;
  const mean = pts.reduce((a, [x, y]) => [a[0] + x / pts.length, a[1] + y / pts.length], [0, 0]);
  const guesses = [tfArcCentreGuess(paths || [pts]), tfFitCircleCentre(pts), { cx: 0, cy: 0 }, { cx: mean[0], cy: mean[1] }].filter(Boolean);
  let best = null;
  for (const g of guesses) {
    let c = { cx: g.cx, cy: g.cy }, rms = Infinity;
    for (let it = 0; it < 40; it++) {
      const d = pts.map(([x, y]) => Math.hypot(x - c.cx, y - c.cy));
      const hi = Math.max(...d), lo = Math.min(...d.filter(v => v > 0.03 * hi));
      const two = hi - lo > 0.05 * hi, mid = (hi + lo) / 2;
      const group = d.map(v => v <= 0.03 * hi ? -1 : two && v < mid ? 0 : 1);   // −1: the centre itself
      const rad = [0, 1].map(k => { const ix = d.map((v, i) => i).filter(i => group[i] === k); return ix.length ? ix.reduce((s, i) => s + d[i], 0) / ix.length : 0; });
      let a11 = 0, a12 = 0, a22 = 0, b1 = 0, b2 = 0, sse = 0, n = 0;
      d.forEach((di, i) => {
        if (group[i] < 0) return;
        const ux = (pts[i][0] - c.cx) / di, uy = (pts[i][1] - c.cy) / di, e = di - rad[group[i]];
        a11 += ux * ux; a12 += ux * uy; a22 += uy * uy; b1 += ux * e; b2 += uy * e; sse += e * e; n++;
      });
      rms = n ? Math.sqrt(sse / n) / (hi || 1) : Infinity;
      const det = a11 * a22 - a12 * a12;
      if (Math.abs(det) < 1e-12) break;
      const dx = (a22 * b1 - a12 * b2) / det, dy = (a11 * b2 - a12 * b1) / det;
      c = { cx: c.cx + dx, cy: c.cy + dy };
      if (Math.hypot(dx, dy) < 1e-10 * (hi || 1)) break;
    }
    if (!best || rms < best.rms) best = { ...c, rms };
  }
  return best;
}

/** degrees clockwise from 12 o'clock of (x, y) around the centre */
const tfGaugeAngle = (x, y, c) => Math.atan2(x - c.cx, y - c.cy) * 180 / Math.PI;

/** a polygon's angular range [from, to] (degrees clockwise from 12), unwrapped along its outline */
function tfAngularRange(pts, c) {
  let prev = null, off = 0, lo = Infinity, hi = -Infinity;
  for (const [x, y] of pts) {
    if (Math.hypot(x - c.cx, y - c.cy) < 1e-9) continue;     // the centre itself has no angle
    let a = tfGaugeAngle(x, y, c) + off;
    if (prev !== null) {
      while (a - prev > 180) { a -= 360; off -= 360; }
      while (a - prev < -180) { a += 360; off += 360; }
    }
    prev = a;
    lo = Math.min(lo, a); hi = Math.max(hi, a);
  }
  return isFinite(lo) ? { from: lo, to: hi } : null;
}

/** Tableau's "Automatic" number format: thousands separators; whole numbers from 1,000 up, else up to 2 decimals */
function tfAutoNumber(v) {
  const dec = Math.abs(v) >= 1000 || Number.isInteger(v) ? 0 : 2;
  return v.toLocaleString("en-US", { minimumFractionDigits: 0, maximumFractionDigits: dec });
}

/* ── map / polygon gauge model ───────────────────────────────────────────── */
/**
 * A gauge drawn with map layers or polygons, rebuilt from its own calculations; null when the sheet is not one
 * (or anything cannot be evaluated). The result describes the dial in degrees clockwise from 12 o'clock:
 *   { kind: "map", start, span, bands: [{ from, to, color, name }], holeRatio, needle: { angle, color } | null,
 *     texts: [{ x, y (gauge units, centre 0,0, radius 1), lines, size }], extent: { x0, x1, y0, y1 }, stroke }
 * @param {FormatModel} model @param {string} sheetName @param {any} summary
 */
function buildMapGaugeModel(model, sheetName, summary) {
  const no = why => { tfLog(`Gauge (map) ${sheetName}: not rebuilt – ${why}`); return null; };
  try {
    const fmt = createSheetFormatter(model, sheetName);
    const sheet = model.sheets && model.sheets[sheetName];
    if (!sheet || !summary || !(summary.data || []).length) return no("no sheet or no data");
    const panes = (sheet.panes || []).filter(p => p.encodings && p.encodings.length && !p.hidden);   // hidden layers are not drawn
    const polyPanes = panes.filter(p => /^polygon$/i.test(p.markClass || ""));
    if (!polyPanes.length) return no("no polygon layer");
    const cols = (summary.columns || []).map((c, i) => {
      const name = c.fieldName || c.fieldId || `Col${i + 1}`;
      const byId = /:(qk|nk|ok)(:\d+)?\]$/.test(c.fieldId || "") ? tfParseFieldRef(c.fieldId) : null;
      return { name, ref: byId || fmt.matchName(name), numeric: /^(int|float|real|integer|number)$/i.test(String(c.dataType || "")) };
    });
    const ctx = tfCalcContext(model, summary, cols);
    const xyRefs = { x: (sheet.cols || []).find(r => r.type === "qk"), y: (sheet.rows || []).find(r => r.type === "qk") };
    const geoGenerated = r => r && /\(generated\)$/i.test(r.name || "");

    /** a pane's marks: the summary rows that belong to it, each with its evaluated position */
    const marksOf = pane => {
      const enc = pane.encodings;
      const spatial = enc.map(e => e.field).find(f => {
        const info = ctx.fieldInfo(f.name, f.ds);
        return info && (info.datatype === "spatial" || /MAKEPOINT|MAKELINE/i.test(ctx.formulaOf(info) || ""));
      });
      const useXY = !spatial && xyRefs.x && xyRefs.y && !geoGenerated(xyRefs.x) && !geoGenerated(xyRefs.y);
      if (!spatial && !useXY) return null;
      // this pane's rows: its location field (or both axis fields) has a value
      const keyCols = spatial ? [ctx.colFor(spatial.name, spatial.ds)] : [ctx.colFor(xyRefs.x.name, xyRefs.x.ds), ctx.colFor(xyRefs.y.name, xyRefs.y.ds)];
      if (keyCols.some(ci => ci < 0)) return null;
      // layers share columns (two text layers on one location field): a layer's rows carry all of its own fields
      enc.forEach(e => {
        if (/^parameters$/i.test(e.field.ds || "")) return;
        const ci = ctx.colFor(e.field.name, e.field.ds);
        if (ci >= 0 && !keyCols.includes(ci)) keyCols.push(ci);
      });
      const out = [];
      (summary.data || []).forEach((row, ri) => {
        if (keyCols.some(ci => tfIsNull(row[ci]))) return;
        let xy = null;
        if (spatial) { const v = ctx.forRow(row, spatial.ds).field([spatial.name]); xy = v && v.point; }
        else {
          const env = ctx.forRow(row, xyRefs.x.ds);
          const x = env.field([xyRefs.x.name]), y = env.field([xyRefs.y.name]);
          xy = typeof x === "number" && typeof y === "number" ? [x, y] : null;
        }
        if (xy && xy.every(isFinite)) out.push({ row, ri, xy });
      });
      return out;
    };

    // dial: every polygon pane's shapes, one per colour member (and other discrete detail), in drawing order
    const shapes = [];
    for (const pane of polyPanes) {
      const pfmt = createSheetFormatter(model, sheetName, [pane]);
      const marks = marksOf(pane);
      if (!marks || marks.length < 6) return no("a polygon layer's positions could not be computed");
      const pathF = (pane.encodings.find(e => e.channel === "path") || {}).field;
      const pi = pathF ? ctx.colFor(pathF.name, pathF.ds) : -1;
      const enc = pfmt.colorEncoding();
      const ci = enc ? ctx.colFor(enc.ref.name, enc.ref.ds) : -1;
      const groupCols = pane.encodings.filter(e => /^(color|lod|detail)$/.test(e.channel) && e.field.type !== "qk")
        .map(e => ctx.colFor(e.field.name, e.field.ds)).filter(i => i >= 0 && i !== pi);
      const groups = new Map();
      marks.forEach(m => {
        const key = groupCols.map(i => tfDvText(m.row[i])).join("\u0001");
        if (!groups.has(key)) groups.set(key, { key, marks: [], first: m.ri });
        groups.get(key).marks.push(m);
      });
      const vals = [...groups.values()].map(g => ci >= 0 ? (enc.continuous ? tfDvNum(g.marks[0].row[ci]) : tfDvText(g.marks[0].row[ci])) : null);
      const scale = ci >= 0 ? tfBuildColorScale(pfmt, enc, vals) : null;
      [...groups.values()].forEach((g, k) => {
        if (pi >= 0) g.marks.sort((a, b) => (tfDvNum(a.row[pi]) || 0) - (tfDvNum(b.row[pi]) || 0));
        const argb = scale ? scale(vals[k]) : null;
        shapes.push({ pts: g.marks.map(m => m.xy), color: (argb ? argb.slice(2) : pfmt.markColor()).toUpperCase(),
                      name: ci >= 0 ? tfDvText(g.marks[0].row[ci]) : "", first: g.first });
      });
    }
    shapes.sort((a, b) => a.first - b.first);                 // Tableau draws in summary order: later on top
    const all = shapes.flatMap(s => s.pts);
    // first centre: the median of the outlines' arc centres (robust: a needle or hub does not pull it off)
    let c = tfArcCentreGuess(shapes.map(sh => sh.pts)) || tfFitDialCentre(all, shapes.map(sh => sh.pts));
    if (!c) return no("no circle fits the polygons");
    // what each polygon is: a dial band (an arc between two radii), the needle (a thin shape from the centre
    // outwards – often a triangle) or the hub (a small shape round the centre); anything else → not a gauge
    const dist = ([x, y]) => Math.hypot(x - c.cx, y - c.cy);
    const Rall = Math.max(...all.map(dist));
    shapes.forEach(sh => {
      const d = sh.pts.map(dist), hi = Math.max(...d), lo = Math.min(...d);
      if (hi < 0.25 * Rall) { sh.role = "hub"; return; }
      const far = sh.pts.filter((p, i) => d[i] > 0.3 * hi);
      const fr = far.length > 1 ? tfAngularRange(far, c) : { from: 0, to: 0 };
      if (lo < 0.2 * Rall && hi > 0.35 * Rall && fr && fr.to - fr.from < 25) { sh.role = "needle"; return; }
      const inner = d.filter(v => v > 0.03 * hi), loArc = inner.length ? Math.min(...inner) : 0;
      const onEdge = d.filter(v => Math.abs(v - hi) < 0.04 * hi || Math.abs(v - loArc) < 0.04 * hi || v < 0.03 * hi).length;
      sh.role = onEdge >= 0.85 * d.length ? "band" : "other";
    });
    const odd = shapes.find(sh => sh.role === "other");
    if (odd) return no(`polygon "${odd.name}" is neither a dial band, a needle nor a hub`);
    const bandShapes = shapes.filter(sh => sh.role === "band");
    if (!bandShapes.length) return no("no polygon is a dial arc");
    // the centre again from the arcs alone (the needle and hub only disturb the fit)
    c = tfFitDialCentre(bandShapes.flatMap(sh => sh.pts), bandShapes.map(sh => sh.pts)) || c;
    const arcPts = bandShapes.flatMap(sh => sh.pts);
    const radii = arcPts.map(([x, y]) => Math.hypot(x - c.cx, y - c.cy));
    const R = Math.max(...radii);
    if (!(R > 0)) return no("zero radius");
    const inner = radii.filter(r => r > 0.03 * R);
    const r0 = inner.length ? Math.min(...inner) : 0;
    // a dial: (almost) every outline point lies on the outer or the inner circle, or at the centre
    const onCircle = radii.filter(r => Math.abs(r - R) < 0.04 * R || (r0 > 0 && Math.abs(r - r0) < 0.04 * R) || r < 0.03 * R).length;
    if (onCircle < 0.85 * radii.length) return no("the polygons are not arcs of one circle");
    const ranges = bandShapes.map(s => ({ ...s, range: tfAngularRange(s.pts, c) })).filter(s => s.range && s.range.to - s.range.from > 0.5);
    if (!ranges.length) return no("no arc has an angular range");
    // one angle frame for all arcs: the first arc's, the others shifted by whole turns to overlap it
    const ref = (ranges[0].range.from + ranges[0].range.to) / 2;
    ranges.forEach(s => {
      const mid = (s.range.from + s.range.to) / 2;
      const k = Math.round((ref - mid) / 360);
      s.range = { from: s.range.from + 360 * k, to: s.range.to + 360 * k };
    });
    const start = Math.min(...ranges.map(s => s.range.from)), end = Math.max(...ranges.map(s => s.range.to));
    const span = Math.min(360, end - start);
    if (span < 30) return no("the arcs span under 30°");
    // the visible bands: at each angle, the arc drawn last that covers it (overlapping arcs = layered bands)
    const bands = [];
    const STEP = 0.25;
    for (let a = start + STEP / 2; a < start + span; a += STEP) {
      let top = null;
      for (const s of ranges) if (a >= s.range.from - 1e-6 && a <= s.range.to + 1e-6) top = s;
      const color = top ? top.color : null, name = top ? top.name : "";
      const last = bands[bands.length - 1];
      if (last && last.color === color && last.name === name) last.to = a + STEP / 2;
      else bands.push({ from: a - STEP / 2, to: a + STEP / 2, color, name });
    }
    // snap band edges to the arcs' own end angles (the sampling is only 0.25° fine)
    const edges = ranges.flatMap(s => [s.range.from, s.range.to]);
    bands.forEach(b => {
      const snap = v => { const e = edges.find(x => Math.abs(x - v) <= STEP); return e === undefined ? v : e; };
      b.from = snap(b.from); b.to = snap(b.to);
    });

    // the other layers: a needle (a line from the centre outwards) and text
    const toUnit = ([x, y]) => [(x - c.cx) / R, (y - c.cy) / R];
    // an angle in the dial's own frame: within 180° of the dial's middle (a dial can sweep more than 180°)
    const mid = start + span / 2;
    const wrap = a => { while (a < mid - 180) a += 360; while (a > mid + 180) a -= 360; return a; };
    let needle = null, hub = null;
    const needleShape = shapes.find(sh => sh.role === "needle");
    if (needleShape) {
      const d = needleShape.pts.map(([x, y]) => Math.hypot(x - c.cx, y - c.cy)), far = Math.max(...d);
      const tip = needleShape.pts[d.indexOf(far)];
      needle = { angle: wrap(tfGaugeAngle(tip[0], tip[1], c)), length: far / R, color: needleShape.color };
    }
    const hubShape = shapes.find(sh => sh.role === "hub");
    if (hubShape) hub = { color: hubShape.color, ratio: Math.max(...hubShape.pts.map(([x, y]) => Math.hypot(x - c.cx, y - c.cy))) / R };
    const texts = [];
    const extentPts = all.map(toUnit);
    for (const pane of panes) {
      if (polyPanes.includes(pane)) continue;
      const mark = String(pane.markClass || "").toLowerCase();
      const marks = marksOf(pane);
      if (!marks || !marks.length) continue;
      const pfmt = createSheetFormatter(model, sheetName, [pane]);
      marks.forEach(m => extentPts.push(toUnit(m.xy)));
      if (mark === "line" && marks.length >= 2) {
        const d = marks.map(m => Math.hypot(m.xy[0] - c.cx, m.xy[1] - c.cy) / R);
        const near = Math.min(...d), far = Math.max(...d);
        if (near < 0.2 && far > 0.4 && (!needle || far > needle.length)) {
          const tip = marks[d.indexOf(far)];
          needle = { angle: wrap(tfGaugeAngle(tip.xy[0], tip.xy[1], c)), length: far, color: pfmt.markColor().toUpperCase() };
        }
        continue;
      }
      if (mark === "text" || pfmt.markLabelsShown()) {
        if (mark !== "text") continue;                         // labels of other marks are not reproduced
        const tpl = pfmt.labelRunsTemplate();
        if (!tpl.length) continue;
        marks.forEach(m => {
          const valueOf = ref => {
            const fmtOf = v => {
              const nf = pfmt.numFmtFor(ref) || (/^parameters$/i.test(ref.ds || "") ? (ctx.paramInfo(ref.name) || {}).defaultFormat : null);
              return typeof v === "number" ? formatNum(v, nf) || tfAutoNumber(v) : String(v);
            };
            try {
              if (/^parameters$/i.test(ref.ds || "")) { const v = ctx.forRow(null).field(["Parameters", ref.name]); return v === null ? "" : fmtOf(v); }
              const ci = ctx.colFor(ref.name, ref.ds);
              // the row's own value as Tableau formatted it
              if (ci >= 0 && !tfIsNull(m.row[ci])) return formatNum(tfDvNum(m.row[ci]), pfmt.numFmtFor(ref)) || tfDvText(m.row[ci]);
              // a value another row carries (the KPI on the needle layer): Tableau's formatting of that row
              if (ci >= 0) { const r2 = (summary.data || []).find(r => !tfIsNull(r[ci])); if (r2) return formatNum(tfDvNum(r2[ci]), pfmt.numFmtFor(ref)) || tfDvText(r2[ci]); }
              const v = ctx.forRow(m.row, ref.ds).field([ref.name]);
              return v === null || v === undefined || typeof v === "object" ? "" : fmtOf(v);
            } catch (e) { return ""; }
          };
          const lines = renderTemplate(tpl, valueOf).map(line => line.map(x => ({ text: x.text, font: pieFont(pfmt.labelFont(x.props)) })));
          if (lines.length && lines.some(l => l.some(x => x.text.trim()))) {
            const [ux, uy] = toUnit(m.xy);
            // cull: "Allow labels to overlap other marks" is off – Tableau hides such a label rather than overlap
            texts.push({ x: ux, y: uy, lines, cull: pfmt.markLabelsCulled() });
          }
        });
      }
    }
    const xs = extentPts.map(p => p[0]), ys = extentPts.map(p => p[1]);
    const strokeRule = (polyPanes[0].style && polyPanes[0].style.rules && polyPanes[0].style.rules.mark || [])
      .find(r => r.attr === "stroke-color");
    tfLog(`Gauge (map): ${ranges.length} arcs → ${bands.length} bands over ${span.toFixed(1)}° from ${start.toFixed(1)}°, ` +
          `hole ${(r0 / R).toFixed(2)}, needle ${needle ? needle.angle.toFixed(1) + "°" : "none"}, hub ${hub ? "yes" : "no"}, ${texts.length} texts`);
    return {
      kind: "map", start, span, bands, holeRatio: r0 / R, needle, hub, texts,
      extent: { x0: Math.min(...xs), x1: Math.max(...xs), y0: Math.min(...ys), y1: Math.max(...ys) },
      stroke: strokeRule ? (tfArgb(strokeRule.value) || "").slice(2) || null : null,
      font: pieFont(fmt.labelFont({}))
    };
  } catch (e) {
    tfLog(`Gauge (map): not rebuilt – ${e.message}`);
    return null;
  }
}

/* ── output ─────────────────────────────────────────────────────────────── */
/**
 * The gauge on the chart data sheet: one row per dial band and the needle's three wedges, in degrees, so the
 * chart's series are linked to cells. @returns {{ dial: any, needle: any, nextRow: number }}
 */
function writeGaugeData(ds, startRow, gauge, name) {
  ds.getCell(startRow, 1).value = name;
  ds.getCell(startRow, 1).font = { bold: true };
  ["Dial band", "Degrees", "From (°)", "To (°)", "Colour"].forEach((h, k) => { ds.getCell(startRow + 1, k + 1).value = h; });
  ds.getRow(startRow + 1).font = { bold: true };
  let row = startRow + 2;
  const dial = gaugeDialSlices(gauge);
  const first = row;
  dial.forEach(s => {
    [s.name, s.value, s.from, s.to, s.color ? "#" + s.color : "(see-through)"].forEach((v, k) => { ds.getCell(row, k + 1).value = v; });
    row++;
  });
  const dialRef = { sheet: ds.name, cat: `$A$${first}:$A$${row - 1}`, val: `$B$${first}:$B$${row - 1}` };
  let needleRef = null;
  const nd = gaugeNeedleSlices(gauge);
  if (nd) {
    row++;
    ds.getCell(row, 1).value = "Needle"; ds.getCell(row, 1).font = { bold: true };
    const f = ++row;
    nd.forEach(s => { ds.getCell(row, 1).value = s.name; ds.getCell(row, 2).value = s.value; row++; });
    needleRef = { sheet: ds.name, cat: `$A$${f}:$A$${row - 1}`, val: `$B$${f}:$B$${row - 1}` };
  }
  return { panes: { dial: dialRef, needle: needleRef }, nextRow: row + 1 };
}

/** the dial's wedges from its start angle: the bands, then the see-through rest of the circle */
function gaugeDialSlices(gauge) {
  const out = gauge.bands.map(b => ({ name: b.name || "Band", value: +(b.to - b.from).toFixed(4),
    from: +(b.from - gauge.start).toFixed(2), to: +(b.to - gauge.start).toFixed(2), color: b.color }));
  const rest = 360 - gauge.span;
  if (rest > 0.01) out.push({ name: "(hidden)", value: +rest.toFixed(4), from: +gauge.span.toFixed(2), to: 360, color: null });
  return out;
}

/** the needle as three pie wedges from the dial's start: see-through, the needle, see-through */
function gaugeNeedleSlices(gauge) {
  if (!gauge.needle) return null;
  const w = 1.6;                                               // degrees: a thin but visible needle
  const at = Math.max(0, Math.min(gauge.span, gauge.needle.angle - gauge.start));
  const before = Math.max(0, at - w / 2);
  return [{ name: "Before the needle", value: +before.toFixed(4) },
          { name: "Needle", value: w },
          { name: "After the needle", value: +(360 - before - w).toFixed(4) }];
}

/**
 * Excel's gauge: a doughnut dial and a pie needle in one plot area, both turned to the dial's start
 * (firstSliceAng, whole degrees clockwise from 12 o'clock). spec: { dial, needle, firstSliceAng, holeSize, plot,
 * bg, font, needleColor, stroke, hasShapes, ref: { dial, needle } }
 */
function gaugeChartXml(spec) {
  const pt = (i, color, line) => `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/><c:spPr>${fillXml(color)}` +
    (line && color ? `<a:ln w="9525"><a:solidFill><a:srgbClr val="${line}"/></a:solidFill></a:ln>` : NO_LINE) + `</c:spPr></c:dPt>`;
  const ser = (idx, name, slices, ref, line) => `<c:ser><c:idx val="${idx}"/><c:order val="${idx}"/><c:tx><c:v>${esc(name)}</c:v></c:tx>` +
    slices.map((s, i) => pt(i, s.color, line)).join("") +
    `${pieCatXml(slices.map(s => s.name), ref)}${pieValXml(slices.map(s => s.value), ref)}</c:ser>`;
  const ang = `<c:firstSliceAng val="${spec.firstSliceAng}"/>`;
  const dial = spec.holeSize
    ? `<c:doughnutChart><c:varyColors val="1"/>${ser(0, "Dial", spec.dial, spec.ref.dial, spec.stroke)}${ang}` +
      `<c:holeSize val="${Math.max(10, Math.min(90, Math.round(spec.holeSize)))}"/></c:doughnutChart>`
    : `<c:pieChart><c:varyColors val="1"/>${ser(0, "Dial", spec.dial, spec.ref.dial, spec.stroke)}${ang}</c:pieChart>`;
  const needle = spec.needle
    ? `<c:pieChart><c:varyColors val="1"/>${ser(1, "Needle", spec.needle.map((s, i) => ({ ...s, color: i === 1 ? spec.needleColor : null })), spec.ref.needle, null)}${ang}</c:pieChart>`
    : "";
  return XML_DECL +
    `<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">` +
    `<c:roundedCorners val="0"/><c:chart><c:autoTitleDeleted val="1"/>` +
    `<c:plotArea>${layoutXml(spec.plot)}${dial}${needle}<c:spPr><a:noFill/>${NO_LINE}</c:spPr></c:plotArea>` +
    `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `<c:spPr>${fillXml(spec.bg)}${NO_LINE}</c:spPr>${txPrXml(spec.font)}` +
    (spec.hasShapes ? `<c:userShapes r:id="rId1"/>` : "") + `</c:chartSpace>`;
}

/**
 * The chart job for a map / polygon gauge in a block of w × h px: the dial sized and placed the way Tableau's map
 * fits all marks into the zone (but never taller than the block: Excel cannot draw a plot area outside its
 * chart), Tableau's text marks as text boxes at their own positions, a hub over the needle's pivot.
 * @param {any} gauge buildMapGaugeModel @param {{ w: number, h: number }} block @param {any} refs writeGaugeData
 * @param {{ name: string, bg: string }} o @returns {any[]}
 */
function buildGaugeCharts(gauge, block, refs, o) {
  const M = ZONE_MARGIN;
  const w = Math.max(40, block.w - 2 * M), h = Math.max(40, block.h - M);
  // px per gauge unit (centre 0,0, radius 1): every mark fits – Tableau's map zooms to its marks – and the whole
  // circle stays inside the chart (Excel cannot draw a plot area outside it), so the hidden half needs room too
  const ex = gauge.extent, padX = 0.08 * (ex.x1 - ex.x0 || 2), padY = 0.08 * (ex.y1 - ex.y0 || 2);
  const x0 = Math.min(ex.x0 - padX, -1), x1 = Math.max(ex.x1 + padX, 1), y0 = ex.y0 - padY, y1 = ex.y1 + padY;
  let k = Math.min(w / (x1 - x0), w / 2, h / 2, h / (y1 - y0));
  let cyMin = 0, cyMax = 0;
  for (let guard = 0; guard < 60; guard++, k *= 0.96) {
    cyMin = Math.max(k, y1 * k);                               // the circle's top and the highest mark inside
    cyMax = Math.min(h - k, h + y0 * k);                       // its bottom and the lowest mark inside
    if (cyMin <= cyMax) break;
  }
  k = Math.max(4, k);
  const D = 2 * k;
  const cx = Math.max(-x0 * k, Math.min(w - x1 * k, w / 2 - (ex.x0 + ex.x1) / 2 * k));
  const cy = Math.max(cyMin, Math.min(cyMax, h / 2 + (ex.y0 + ex.y1) / 2 * k));
  const texts = [];
  // Tableau's map fits the marks into the whole zone; Excel's dial is often smaller (the whole circle, hidden half
  // included, must fit in the chart). The text shrinks by the same ratio, so labels keep Tableau's spacing.
  const kTableau = Math.min(w / (x1 - x0), h / (y1 - y0));
  const ratio = Math.min(1, k / (kTableau || k));
  const scaled = f => ({ ...f, size: Math.max(6, Math.round((f.size || 9) * ratio * 2) / 2) });
  // boxes as Excel draws the text (points at 96 dpi – the screen's scaling does not apply inside Excel)
  const boxes = gauge.texts.map((t, order) => {
    const lines = t.lines.map(l => l.map(x => ({ ...x, font: scaled(x.font) })));
    const pt = Math.max(...lines.flatMap(l => l.map(x => x.font.size)), 6);
    const tw = Math.max(...lines.map(l => l.reduce((s, x) => s + textWidthPx(x.text, x.font, x.font.size * 96 / 72), 0))) + 4;
    const th = lines.length * pt * 96 / 72 * LABEL_LINE_EM;
    const x = cx + t.x * k, y = cy - t.y * k;
    return { x: Math.max(0, Math.min(w - tw, x - tw / 2)), y: Math.max(0, Math.min(h - th, y - th / 2)), w: tw, h: th,
             lines, anchor: "ctr", align: "ctr", cull: t.cull !== false, pt, order };
  });
  // Tableau hides a label that would overlap another rather than draw them on top of each other ("Allow labels to
  // overlap other marks" off): labels that must show first, then the bigger text, then Tableau's order
  const overlap = (a, b) => a.x < b.x + b.w - 1 && b.x < a.x + a.w - 1 && a.y < b.y + b.h - 1 && b.y < a.y + a.h - 1;
  const placed = [];
  boxes.slice().sort((a, b) => (a.cull - b.cull) || (b.pt - a.pt) || (a.order - b.order)).forEach(b => {
    if (b.cull && placed.some(p => overlap(b, p))) { tfLog(`Gauge: label "${b.lines.flat().map(x => x.text).join("").trim()}" hidden – it would overlap another, as in Tableau`); return; }
    placed.push(b);
  });
  placed.sort((a, b) => a.order - b.order).forEach(({ cull, pt, order, ...box }) => texts.push(box));
  if (gauge.hub || gauge.needle) {                             // the pivot: Tableau's hub, else a small disc in the needle's colour
    const hub = gauge.hub ? Math.max(6, D * gauge.hub.ratio) : Math.max(6, D * 0.05);
    texts.push({ x: cx - hub / 2, y: cy - hub / 2, w: hub, h: hub, shape: "ellipse", fill: gauge.hub ? gauge.hub.color : gauge.needle.color, lines: [] });
  }
  const startInt = Math.round(((gauge.start % 360) + 360) % 360) % 360;
  const spec = {
    dial: gaugeDialSlices(gauge), needle: gaugeNeedleSlices(gauge), firstSliceAng: startInt,
    holeSize: gauge.holeRatio > 0.05 ? 100 * gauge.holeRatio : null,
    plot: { x: (cx - D / 2) / w, y: (cy - D / 2) / h, w: D / w, h: D / h },
    bg: o.bg, font: gauge.font, needleColor: gauge.needle ? gauge.needle.color : null,
    stroke: gauge.stroke, hasShapes: texts.length > 0, ref: refs
  };
  return [{ x: M, y: 0, w, h, name: o.name, xml: gaugeChartXml(spec),
            shapes: texts.length ? userShapesXml(texts, { w, h }) : null, tips: [] }];
}

/** a pie / donut gauge: one slice painted the background colour (Tableau's trick to hide the bottom half) */
function pieIsGauge(pie, bg) {
  const bgU = String(bg || "FFFFFF").toUpperCase();
  // each pie on its own: small multiples (one gauge per member) share the points list
  const pies = new Map();
  pie.points.forEach(p => { const k = `${p.rowKey}\u0001${p.colKey}`; if (!pies.has(k)) pies.set(k, []); pies.get(k).push(p); });
  return [...pies.values()].some(pts => {
    const total = pts.reduce((s, p) => s + Math.abs(p.value || 0), 0);
    return total > 0 && pts.some(p => p.color && String(p.color).toUpperCase() === bgU &&
                                      Math.abs(p.value) / total >= 0.2 && Math.abs(p.value) / total <= 0.6);
  });
}


/* ═══ charts/model/specialized.js ═════════════════════════════════════════════════════════════════ */
/* Pie, scatter, waterfall, box plot, Gantt, treemap, symbol map and packed bubble chart specs. */

/** @param {ChartContext} ctx @returns {ChartSpec[]} */
function tvPieSpec(ctx) {
  const { vm, roles } = ctx;
  let catCi = roles.color && !roles.color.measureNames && !roles.color.continuous ? roles.color.ci : -1;
  if (catCi < 0) {
    const dims = [...roles.rows.dims, ...roles.cols.dims];
    catCi = dims.length ? dims[0].ci : vm.cols.findIndex(c => c.isHeader);
  }
  let angleCi = roles.angle;
  if (angleCi < 0) {
    const v = [...roles.rows.values, ...roles.cols.values][0];
    angleCi = v ? v.ci : vm.cols.findIndex(c => !c.isHeader && c.name !== (vm.cols[catCi] || {}).name);
  }
  const scale = tvColorScale(vm, roles, "pie");
  const spec = /** @type {ChartSpec} */ ({ ...tvBaseSpec(vm), kind: ctx.doughnut ? "doughnut" : "pie", legend: true });
  // Tableau's donut: a pie of slices (colour / angle) and, on the second axis, a smaller plain pie – the hole,
  // in the background colour, carrying the total as its label
  const piePanes = roles.panes.filter(p => tvMarkToken(p.markClass) === "pie");
  const slicing = p => p.encodings.some(e => e.channel === "wedge-size" || e.channel === "color");
  const slicePane = ctx.doughnut ? piePanes.find(slicing) || null : null;
  const holePane = ctx.doughnut ? piePanes.find(p => p.id && !slicing(p)) || null : null;
  let categories, values, cats = null;
  if ((roles.color && roles.color.measureNames) || (catCi < 0 && vm.cols.some(c => c.pivoted))) {
    const measures = vm.cols.map((c, i) => c.pivoted ? i : -1).filter(i => i >= 0);
    categories = { names: ["Measure Names"], levels: [measures.map(ci => tvMeasureLabel(vm, ci))] };
    values = measures.map(ci => vm.rows.reduce((s, r) => s + (tfDvNum(r[ci]) || 0), 0));
    angleCi = measures[0];
  } else {
    if (catCi < 0 || angleCi < 0) throw new Error("pie has no colour dimension or angle measure");
    cats = tvCategories(vm, [catCi], false);
    categories = { names: [tvMeasureLabel(vm, catCi)], levels: cats.levels };
    values = tvSum(vm, cats, angleCi);
  }
  // a dual-pie donut's hole contributes a mark with no category: no value, or the total of the slices → not a slice
  const total = values.reduce((s, v) => s + (v || 0), 0);
  const nullish = l => /^(null|%null%|\(null\))?$/i.test(String(l == null ? "" : l).trim());
  const isHole = i => nullish(categories.levels[0][i]) &&
    (!values[i] || (ctx.doughnut && Math.abs(values[i] - (total - values[i])) <= Math.abs(values[i]) * 0.005));
  const keep = categories.levels[0].map((_, i) => !isHole(i));
  const kept = keep.map((k, i) => k ? i : -1).filter(i => i >= 0);
  if (keep.includes(false)) {
    categories = { ...categories, levels: categories.levels.map(lv => lv.filter((_, i) => keep[i])) };
    values = values.filter((_, i) => keep[i]);
  }
  if (categories.levels[0].length > 200) throw new Error("too many pie slices for an Excel chart");
  const angleRef = vm.cols[angleCi] && vm.cols[angleCi].ref;
  const catRef = catCi >= 0 && vm.cols[catCi].ref;
  // a "% of Total" quick table calc on Label → Excel's percentage label instead of the raw value
  const labelPct = roles.labelRefs.some(r => /^pcto$/i.test(r.deriv || ""));
  const labelValue = !!angleRef && roles.labelRefs.some(r => !/^pcto$/i.test(r.deriv || "") && tfSameField(r, angleRef)) ||
    (!labelPct && tvLabelsOn(roles, angleRef));
  const labelCat = !!catRef && roles.labelRefs.some(r => tfSameField(r, catRef));
  spec.categories = categories;
  spec.numFmt = tvNumFmt(vm, angleCi);
  spec.series = [{
    name: tvMeasureLabel(vm, angleCi),
    color: tvMarkColor(roles),
    values,
    pointColors: categories.levels[0].map((v, i) => (scale && scale(v, i)) || tvHex(TABLEAU_10[i % TABLEAU_10.length])),
    labels: labelValue || labelCat || labelPct,
    labelParts: { value: labelValue, category: labelCat, percent: labelPct },
    labelNumFmt: labelPct && !labelValue ? "0.0%" : undefined
  }];
  if (slicePane && cats) {
    // the field values as Tableau writes them (measures in their number format, empty for null)
    const text = (row, ref) => {
      const ci = vm.cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
      if (ci < 0 || !row || tfIsNull(row[ci])) return "";
      const n = tfDvNum(row[ci]);
      return vm.cols[ci].isHeader || n === null ? tvText(row[ci]) : tfFormatNumber(n, tvNumFmt(vm, ci));
    };
    const rowsOf = new Map();
    vm.rows.forEach(r => { const i = cats.indexOf(r); if (i !== undefined && keep[i] && !rowsOf.has(i)) rowsOf.set(i, r); });
    const flat = lines => lines.map(l => l.map(s => s.text).join("")).filter(t => t.trim()).join("\n");
    // slice labels: only the slices' own marks card (blank until its label fields have values)
    const texts = kept.map(i => flat(tvPaneLabel(slicePane, ref => text(rowsOf.get(i), ref))));
    const labelled = texts.some(Boolean) && slicePane.style && tvPaneRule(slicePane, "mark", "mark-labels-show") !== "false";
    Object.assign(spec.series[0], { labels: labelled, labelTexts: labelled ? texts : undefined, labelParts: undefined, labelNumFmt: undefined });
    if (holePane) {
      // the hole's label: its fields over all the slices (the total), in its runs' fonts
      const sumOf = ref => {
        const ci = vm.cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
        if (ci < 0) return "";
        if (vm.cols[ci].isHeader) return "";
        const s = vm.rows.reduce((acc, r) => { const i = cats.indexOf(r); return i !== undefined && keep[i] ? acc + (tfDvNum(r[ci]) || 0) : acc; }, 0);
        return tfFormatNumber(s, tvNumFmt(vm, ci));
      };
      spec.centerLabel = tvPaneLabel(holePane, sumOf).filter(l => l.some(s => s.text.trim())).map(l => l.map(s => ({
        text: s.text, bold: !!s.props.bold, size: s.props.fontSize, color: tvHex(s.props.color) || undefined, font: s.props.fontName })));
      // the hole's size against the ring's (the marks' Size sliders)
      const sizeOf = p => Number(tvPaneRule(p, "mark", "size"));
      const ratio = sizeOf(holePane) / sizeOf(slicePane);
      if (ratio > 0 && ratio < 1) spec.holeSize = Math.round(Math.max(10, Math.min(90, ratio * 100)));
    }
  }
  return [spec];
}

/**
 * @param {ChartContext} ctx
 * @param {{ lines?: boolean }} [opts] lines: a Line mark with a measure on both axes – the points joined in order of x
 * @returns {ChartSpec[]}
 */
function tvScatterSpec(ctx, opts = {}) {
  const { vm, roles } = ctx;
  const xm = roles.cols.values[0], ym = roles.rows.values[0];
  const scale = tvColorScale(vm, roles, ctx.markToken);
  const spec = /** @type {ChartSpec} */ ({ ...tvBaseSpec(vm), kind: "scatter", numFmt: tvNumFmt(vm, ym.ci), xNumFmt: tvNumFmt(vm, xm.ci),
                 valueTitle: tvMeasureLabel(vm, ym.ci), xTitle: tvMeasureLabel(vm, xm.ci) });
  const point = r => ({ x: tfDvNum(r[xm.ci]), y: tfDvNum(r[ym.ci]) });
  const rows = vm.rows.filter(r => { const p = point(r); return p.x !== null && p.y !== null; });
  if (opts.lines) rows.sort((a, b) => point(a).x - point(b).x || point(a).y - point(b).y);
  if (rows.length > TV_MAX_POINTS) throw new Error(`${rows.length} marks – too many for an Excel scatter chart`);
  const colorCi = roles.color && !roles.color.measureNames ? roles.color.ci : -1;
  // Label = a dimension (store name …) → Excel "value from cells" labels with that text, when
  // few enough marks to stay readable; numbers only when a measure itself is on Label
  const labelDim = roles.textDims[0] ?? -1;
  const measureLabel = roles.labelRefs.some(r => tfSameField(r, ym.ref) || tfSameField(r, xm.ref));
  const textLabels = labelDim >= 0 && rows.length <= 40;
  const labels = textLabels || measureLabel || (labelDim < 0 && tvLabelsOn(roles, ym.ref));
  const labelOf = r => tvText(r[labelDim]);
  // a Size measure is part of the meaning (Tableau's circles sized by it): Excel's bubble chart, area = size
  const sizeCi = !opts.lines && roles.size >= 0 && !vm.cols[roles.size].isHeader ? roles.size : -1;
  const sizeOf = r => { const n = tfDvNum(r[sizeCi]); return n === null ? null : Math.abs(n); };
  // a size that is never negative is the measure itself (each value knows its mark); a bubble's area is its magnitude
  const sizes = own => own.every(r => (tfDvNum(r[sizeCi]) ?? 0) >= 0) ? tvMarkValues(own, sizeCi) : own.map(sizeOf);
  if (colorCi >= 0 && !roles.color.continuous) {
    const values = tvColorValues(vm, roles);
    if (values.length > TV_MAX_SERIES) throw new Error("too many colour values for an Excel chart");
    spec.series = values.map((v, i) => {
      const own = rows.filter(r => tvText(r[colorCi]) === v);
      return { name: v, color: (scale && scale(v)) || tvHex(TABLEAU_10[i % TABLEAU_10.length]),
               x: tvMarkValues(own, xm.ci), y: tvMarkValues(own, ym.ci), labels, ...(opts.lines ? { line: true } : {}),
               ...(sizeCi >= 0 ? { size: sizes(own) } : {}),
               labelTexts: textLabels ? own.map(labelOf) : undefined };
    });
  } else {
    spec.series = [{
      name: tvMeasureLabel(vm, ym.ci), color: tvMarkColor(roles), x: tvMarkValues(rows, xm.ci), y: tvMarkValues(rows, ym.ci), labels,
      ...(opts.lines ? { line: true } : {}), ...(sizeCi >= 0 ? { size: sizes(rows) } : {}),
      labelTexts: textLabels ? rows.map(labelOf) : undefined,
      pointColors: colorCi >= 0 && scale ? rows.map(r => scale(tfDvNum(r[colorCi]))) : undefined
    }];
  }
  if (sizeCi >= 0) Object.assign(spec, { kind: "bubble", sizeTitle: tvMeasureLabel(vm, sizeCi), sizeNumFmt: tvNumFmt(vm, sizeCi) });
  spec.legend = spec.series.length > 1;
  return [spec];
}

/* value axis + category dimensions shared by waterfall / box plot */
/** @param {ChartContext} ctx */
function tvAxisLayout(ctx) {
  const { roles } = ctx;
  const valueShelf = roles.rows.values.length ? "rows" : roles.cols.values.length ? "cols" : null;
  if (!valueShelf) throw new Error("no continuous measure axis to chart");
  const catShelf = valueShelf === "rows" ? "cols" : "rows";
  return { valueShelf, horizontal: valueShelf === "cols", measure: roles[valueShelf].values[0],
           catDims: [...roles[valueShelf].dims, ...roles[catShelf].dims] };
}

/* ── waterfall: Gantt bars sized by -measure on a running total →
 *    stacked columns: invisible base + increase / decrease, each step in its category's colour, and the
 *    Text field written over each floating bar ─────────────────────────────── */
/** @param {ChartContext} ctx @returns {ChartSpec[]} */
function tvWaterfallSpec(ctx) {
  const { vm, roles } = ctx;
  const { horizontal, measure, catDims, valueShelf } = tvAxisLayout(ctx);
  const fieldLabels = vm.fmt.hasModel ? vm.fmt.fieldLabelsShown(valueShelf === "rows" ? "cols" : "rows") : true;
  const cats = tvCategories(vm, catDims.map(d => d.ci), catDims.some(d => d.continuous));
  if (cats.count > TV_MAX_POINTS) throw new Error(`${cats.count} categories – too many for an Excel chart`);
  const running = tvSum(vm, cats, measure.ci);
  if (running.some(v => v !== null && v < 0)) throw new Error("running total goes below zero – not drawable as stacked columns");
  const base = [], up = [], down = [], ends = [];
  let prev = 0;
  running.forEach(v => {
    const cur = v === null ? prev : v;
    const delta = cur - prev;
    base.push(Math.min(prev, cur));
    up.push(delta > 0 ? delta : null);
    down.push(delta < 0 ? -delta : null);
    ends.push(Math.max(prev, cur));                             // the top of the floating bar
    prev = cur;
  });
  const markColor = tvMarkColor(roles);
  const scale = tvColorScale(vm, roles, "ganttbar");
  // colour = the category (each step its own colour), or a measure (usually the step itself)
  const colorLevel = roles.color && !roles.color.continuous ? catDims.findIndex(d => d.ci === roles.color.ci) : -1;
  const stepColors = !scale ? null : colorLevel >= 0 ? cats.levels[colorLevel].map(v => scale(v))
    : roles.color && roles.color.continuous && roles.color.ci >= 0 ? tvSum(vm, cats, roles.color.ci).map(v => scale(v)) : null;
  const label = tvMeasureLabel(vm, measure.ci);
  const numFmt = tvNumFmt(vm, measure.ci);
  // mark labels: the Text field (the step, or the running total) over each bar; Excel writes stacked labels
  // inside the bars, so an invisible line along the bar tops carries them above, as text in its format
  const textCi = roles.labelRefs.map(r => vm.cols.findIndex(c => c.ref && tfSameField(c.ref, r))).find(i => i >= 0);
  const labelCi = textCi === undefined ? measure.ci : textCi;
  const labels = tvLabelsOn(roles, vm.cols[labelCi].ref);
  const labelFmt = tvNumFmt(vm, labelCi);
  const labelValues = labelCi === measure.ci ? running : tvSum(vm, cats, labelCi);
  const above = labels && !horizontal;
  // decreases are drawn as positive heights but labelled as the negative step, like Tableau
  const positive = (numFmt.match(/^((?:"[^"]*"|[^;])*)/) || [])[1] || "General";
  const downFmt = positive === "General" ? "-General" : "-" + positive;
  /** @type {ChartSeries[]} */
  const series = [
    { name: "Base", type: "bar", values: base, color: null, labels: false },
    { name: label + " (increase)", type: "bar", values: up, color: markColor, pointColors: stepColors, labels: labels && !above },
    { name: label + " (decrease)", type: "bar", values: down, color: markColor, pointColors: stepColors, labels: labels && !above, labelNumFmt: downFmt }
  ];
  // Analysis → Totals → Show Grand Totals on the category shelf: Tableau ends with a full bar from zero to the
  // final running total (drawn in grey)
  const catShelf = valueShelf === "rows" ? "cols" : "rows";
  const totals = vm.fmt.sheetModel && vm.fmt.sheetModel.grandTotals && vm.fmt.sheetModel.grandTotals[catShelf];
  if (totals && running.length) {
    const n = running.length;
    series.forEach(s => { s.values = [...s.values, null]; if (s.pointColors) s.pointColors = [...s.pointColors, null]; });
    series.push({ name: "Grand Total", type: "bar", values: [...Array(n).fill(null), prev], color: "9E9E9E", labels: labels && !above });
    cats.levels = cats.levels.map((lv, i) => [...lv, i === 0 ? "Grand Total" : ""]);
    ends.push(prev); labelValues.push(labelCi === measure.ci ? prev : null);
  }
  if (above) {
    series.push({ name: label + " (labels)", type: "line", values: ends, color: null, line: false, marker: false, labels: true,
                  labelTexts: labelValues.map(v => v === null ? "" : tfFormatNumber(v, labelFmt)) });
  }
  return [{
    // Tableau's Gantt bars float thin, about as wide as the gaps between them
    ...tvBaseSpec(vm), kind: "bar", barDir: horizontal ? "bar" : "col", stacked: true, gapWidth: 100, legend: false,
    conversion: { strategy: "CONSTRUCTED", output: "Waterfall: stacked columns on an invisible running base",
                  note: "built from stacked columns (invisible base, increases, decreases) so Tableau's step colours and labels stay – Excel's own waterfall cannot colour each step" },
    // "Allow labels to overlap other marks" off (Tableau's default): overlapping labels are left out
    labelCull: !roles.panes.some(p => tvPaneRule(p, "mark", "mark-labels-cull") === "false"),
    categories: { names: catDims.length ? catDims.map(d => tvMeasureLabel(vm, d.ci)) : [""], levels: cats.levels },
    categoryTitle: fieldLabels ? catDims.map(d => tvMeasureLabel(vm, d.ci)).join(" / ") : "",
    numFmt, valueTitle: label, series,
    // what bindWaterfall needs to rebuild the steps with formulas (series: base, increase, decrease[, grand total][, labels])
    waterfall: { running, ref: vm.cols[measure.ci].ref, totals: !!(totals && running.length),
                 labels: above ? { running: labelCi === measure.ci, fmt: labelFmt,
                                   // the step itself (the running total's own measure) – Tableau's usual waterfall label
                                   step: labelCi !== measure.ci && !!vm.cols[labelCi].ref && !!tfTableCalcBase(vm.cols[measure.ci].ref) &&
                                         tfSameField(vm.cols[labelCi].ref, tfTableCalcBase(vm.cols[measure.ci].ref)) } : null }
  }];
}

/**
 * A waterfall following the slicers: each step a formula of the running total's base measure (checked against
 * Tableau's steps), the running total adding them up, and the invisible base, the increase / decrease bars, the
 * grand total and the labels computed from it – cells of the chart's data (writeChartData: helpers, cellFormula).
 * false when the steps cannot be rebuilt from the rows: the waterfall keeps Tableau's values.
 * @param {ChartSpec} spec @param {any} binding vm.slicerBinding
 */
function bindWaterfall(spec, binding) {
  const w = /** @type {any} */ (spec).waterfall, run = w && w.running, src = run && run.src;
  const base = w && w.ref && tfTableCalcBase(w.ref);
  if (!src || !base || run.some(v => v === null)) return false;
  const n = run.length, steps = run.map((v, i) => v - (i ? run[i - 1] : 0));
  const f = binding.measureFormulas(base, src.rows, src.ci, steps, "zero");
  if (!f) return false;
  const RT = "Running total";
  const cur = (i, at) => at("helper", RT, Math.min(i, n - 1));
  const prev = (i, at) => (i ? at("helper", RT, i - 1) : "0");
  /** @type {any} */ (spec).helpers = [
    { name: "Step", results: steps, formula: i => (i < n ? f[i] : null) },
    { name: RT, results: run, formula: (i, at) => (i >= n ? null : i ? `${prev(i, at)}+${at("helper", "Step", i)}` : at("helper", "Step", i)) }
  ];
  const S = spec.series;
  S[0].cellFormula = (i, at) => (i < n ? `MIN(${prev(i, at)},${cur(i, at)})` : null);
  S[1].cellFormula = (i, at) => (i < n ? `IF(${cur(i, at)}>${prev(i, at)},${cur(i, at)}-${prev(i, at)},NA())` : null);
  S[2].cellFormula = (i, at) => (i < n ? `IF(${cur(i, at)}<${prev(i, at)},${prev(i, at)}-${cur(i, at)},NA())` : null);
  if (w.totals) S[3].cellFormula = (i, at) => (i === n ? cur(i, at) : null);
  if (w.labels) {
    const lab = S[S.length - 1];
    lab.cellFormula = (i, at) => (i < n ? `MAX(${prev(i, at)},${cur(i, at)})` : cur(i, at));
    const fmt = `"${String(w.labels.fmt || "General").replace(/"/g, '""')}"`;
    if (w.labels.running) lab.labelFormula = (i, at) => `TEXT(${cur(i, at)},${fmt})`;
    // the step's label; the grand total's is its running total
    else if (w.labels.step) lab.labelFormula = (i, at) => `TEXT(${i < n ? at("helper", "Step", i) : cur(i, at)},${fmt})`;
  }
  binding.written += n;
  return true;
}

/* ── box plot: quartiles per category → line chart with up/down bars (box),
 *    high-low lines (whiskers) and a dash at the median ───────────────────── */
function tvQuantile(sorted, p) {                      // linear interpolation, as Tableau / QUARTILE.INC
  const pos = (sorted.length - 1) * p, lo = Math.floor(pos), hi = Math.ceil(pos);
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (pos - lo);
}

/** @param {ChartContext} ctx @returns {ChartSpec[]} */
function tvBoxPlotSpec(ctx) {
  const { vm, roles } = ctx;
  const { measure, catDims } = tvAxisLayout(ctx);
  const cats = tvCategories(vm, catDims.map(d => d.ci), catDims.some(d => d.continuous));
  if (cats.count > 200) throw new Error("too many boxes for an Excel chart");
  const groups = Array.from({ length: cats.count }, () => []);
  vm.rows.forEach(r => {
    const n = tfDvNum(r[measure.ci]);
    const i = cats.indexOf(r);
    if (n !== null && i !== undefined) groups[i].push(n);
  });
  const stats = groups.map(g => {
    if (!g.length) return null;
    const s = g.slice().sort((a, b) => a - b);
    const q1 = tvQuantile(s, 0.25), med = tvQuantile(s, 0.5), q3 = tvQuantile(s, 0.75), iqr = q3 - q1;
    // standard whiskers: the furthest marks within 1.5 × IQR of the box
    const low = s.find(v => v >= q1 - 1.5 * iqr), high = [...s].reverse().find(v => v <= q3 + 1.5 * iqr);
    return { q1, med, q3, low, high };
  });
  const col = k => stats.map(st => st ? st[k] : null);
  const boxColor = tvMarkColor(roles);
  return [{
    ...tvBaseSpec(vm), kind: "line", boxPlot: { color: boxColor }, legend: false, includeZero: false,
    conversion: { strategy: "CONSTRUCTED", output: "Box plot: quartile boxes (up/down bars), whiskers (high-low lines), median marks",
                  note: "Tableau's quartiles and 1.5 × IQR whiskers computed from the marks – Excel's box & whisker chart would compute its own statistics" },
    categories: { names: catDims.length ? catDims.map(d => tvMeasureLabel(vm, d.ci)) : [""], levels: cats.levels },
    categoryTitle: catDims.map(d => tvMeasureLabel(vm, d.ci)).join(" / "),
    numFmt: tvNumFmt(vm, measure.ci), valueTitle: tvMeasureLabel(vm, measure.ci),
    // up/down bars span the FIRST and LAST series → Q1 … Q3; high-low lines span all series
    series: [
      { name: "Lower quartile", type: "line", values: col("q1"), color: boxColor, line: false, marker: false },
      { name: "Lower whisker", type: "line", values: col("low"), color: boxColor, line: false, marker: false },
      { name: "Median", type: "line", values: col("med"), color: "333333", line: false, marker: true, markerSymbol: "dash", markerSize: 14 },
      { name: "Upper whisker", type: "line", values: col("high"), color: boxColor, line: false, marker: false },
      { name: "Upper quartile", type: "line", values: col("q3"), color: boxColor, line: false, marker: false }
    ]
  }];
}

/* ── Gantt: per row, marks start at the date axis value and last Size days →
 *    stacked horizontal bars of alternating invisible gaps and visible bars ─ */
const TV_GANTT_MAX_SERIES = 250;

function tvExcelSerial(dv) {
  if (tfIsNull(dv)) return null;
  const v = dv.nativeValue !== undefined ? dv.nativeValue : dv.value;
  if (typeof v === "number") return v;                         // already a serial / numeric axis
  const t = v instanceof Date ? v.getTime() : Date.parse(String(v));
  return isNaN(t) ? null : t / 86400000 + 25569;               // days since 1899-12-30 (UTC)
}

/** @param {ChartContext} ctx @returns {ChartSpec[]} */
function tvGanttSpec(ctx) {
  const { vm, roles } = ctx;
  const startDim = [...roles.cols.dims, ...roles.rows.dims].find(d => d.continuous);
  if (!startDim) throw new Error("no continuous date axis for the Gantt bars");
  if (roles.size < 0) throw new Error("no Size measure for the Gantt bar length");
  const startShelf = roles.cols.dims.includes(startDim) ? "cols" : "rows";
  const catDims = roles[startShelf === "cols" ? "rows" : "cols"].dims.filter(d => !d.continuous);
  const cats = tvCategories(vm, catDims.map(d => d.ci), false);
  const colorScale = tvColorScale(vm, roles, "ganttbar");
  const colorCi = roles.color && !roles.color.measureNames ? roles.color.ci : -1;
  const markColor = tvMarkColor(roles);
  const marks = Array.from({ length: cats.count }, () => []);
  vm.rows.forEach(r => {
    const start = tvExcelSerial(r[startDim.ci]), len = tfDvNum(r[roles.size]), i = cats.indexOf(r);
    if (start === null || len === null || i === undefined) return;
    const c = colorCi >= 0 && colorScale ? colorScale(roles.color.continuous ? tfDvNum(r[colorCi]) : tvText(r[colorCi])) : null;
    marks[i].push({ start, end: start + Math.max(0, len), color: c || markColor });
  });
  const all = marks.flat();
  if (!all.length) throw new Error("no Gantt marks with a start date and a length");
  const axisMin = Math.floor(Math.min(...all.map(m => m.start)));
  const axisMax = Math.ceil(Math.max(...all.map(m => m.end)));
  const depth = Math.max(...marks.map(m => m.length));
  if (2 * depth > TV_GANTT_MAX_SERIES) throw new Error(`${depth} bars in one row – too many for an Excel chart`);
  /** @type {ChartSeries[]} */
  const series = [];
  for (let k = 0; k < depth; k++) {
    series.push({ name: `Gap ${k + 1}`, type: "bar", values: [], color: null, labels: false });
    series.push({ name: `Bar ${k + 1}`, type: "bar", values: [], color: markColor, pointColors: [], labels: false });
  }
  marks.forEach((row, i) => {
    row.sort((a, b) => a.start - b.start);
    let edge = 0;
    row.forEach((m, k) => {
      if (m.start < edge - 1e-9) throw new Error("overlapping Gantt bars in one row – not drawable as stacked bars");
      series[2 * k].values[i] = k === 0 ? m.start : m.start - edge;
      series[2 * k + 1].values[i] = m.end - m.start;
      series[2 * k + 1].pointColors[i] = m.color;
      edge = m.end;
    });
    for (let k = row.length; k < depth; k++) { series[2 * k].values[i] = null; series[2 * k + 1].values[i] = null; }
  });
  const dateFmt = axisMax - axisMin > 400 ? "mmm yyyy" : "d mmm yyyy";
  return [{
    ...tvBaseSpec(vm), kind: "bar", barDir: "bar", stacked: true, gapWidth: 40, legend: false,
    conversion: { strategy: "CONSTRUCTED", output: "Gantt: stacked bars with invisible gaps on a date axis",
                  note: "Excel has no Gantt chart: each bar starts at its date and lasts its Size, the gaps before it invisible" },
    valueMin: axisMin, valueMax: axisMax, includeZero: false,
    categories: { names: catDims.length ? catDims.map(d => tvMeasureLabel(vm, d.ci)) : [""], levels: cats.levels },
    categoryTitle: catDims.map(d => tvMeasureLabel(vm, d.ci)).join(" / "),
    numFmt: dateFmt, valueTitle: tvMeasureLabel(vm, startDim.ci), series
  }];
}

/* ── treemap: nothing on Rows/Columns, Size = measure, Label/Detail = the
 *    hierarchy → Excel treemap (chartex), each tile in its Tableau colour ── */
/** @param {ChartContext} ctx @returns {ChartSpec[]} */
function tvTreemapSpec(ctx) {
  const { vm, roles } = ctx;
  if (roles.size < 0) throw new Error("no Size measure for the treemap");
  let levelCis = [...roles.detailDims, ...roles.textDims].filter((ci, i, a) => a.indexOf(ci) === i);
  if (roles.color && !roles.color.continuous && roles.color.ci >= 0 && !levelCis.includes(roles.color.ci)) levelCis.unshift(roles.color.ci);
  if (!levelCis.length) levelCis = vm.cols.map((c, i) => c.isHeader ? i : -1).filter(i => i >= 0);
  if (!levelCis.length) throw new Error("no dimension to split the treemap");
  // outermost level = the one with the fewest distinct values
  const distinct = ci => new Set(vm.rows.map(r => tvText(r[ci]))).size;
  levelCis.sort((a, b) => distinct(a) - distinct(b));
  const cats = tvCategories(vm, levelCis, false);
  if (cats.count > 1000) throw new Error(`${cats.count} tiles – too many for an Excel treemap`);
  const sizes = tvSum(vm, cats, roles.size);
  const scale = tvColorScale(vm, roles, "square");
  let pointColors;
  if (scale && roles.color.continuous) pointColors = tvSum(vm, cats, roles.color.ci).map(v => scale(v));
  else if (scale && roles.color.ci >= 0) {
    const level = levelCis.indexOf(roles.color.ci);
    pointColors = cats.levels[level].map((v, i) => scale(v) || tvHex(TABLEAU_10[i % TABLEAU_10.length]));
  }
  // drop empty / negative tiles: a treemap cannot draw them
  const keep = sizes.map((v, i) => v !== null && v > 0 ? i : -1).filter(i => i >= 0);
  // Excel builds the hierarchy from consecutive rows: keep each parent's tiles together,
  // biggest parent first, biggest tile first inside it (Tableau's treemap order)
  const pathKey = (i, l) => cats.levels.slice(0, l + 1).map(lv => lv[i]).join("\u0001");
  const totals = cats.levels.map((_, l) => {
    const m = new Map();
    keep.forEach(i => m.set(pathKey(i, l), (m.get(pathKey(i, l)) || 0) + sizes[i]));
    return m;
  });
  keep.sort((a, b) => {
    for (let l = 0; l < cats.levels.length; l++) {
      const ka = pathKey(a, l), kb = pathKey(b, l);
      if (ka === kb) continue;
      return (totals[l].get(kb) - totals[l].get(ka)) || (ka < kb ? -1 : 1);
    }
    return 0;
  });
  return [{
    ...tvBaseSpec(vm), kind: "treemap", legend: false,
    categories: { names: levelCis.map(ci => tvMeasureLabel(vm, ci)), levels: cats.levels.map(lv => keep.map(i => lv[i])) },
    numFmt: tvNumFmt(vm, roles.size),
    series: [{ name: tvMeasureLabel(vm, roles.size), values: tvPickSum(sizes, keep), color: tvMarkColor(roles),
               pointColors: pointColors ? keep.map(i => pointColors[i]) : undefined, labels: true,
               // Size measure also on Label → "Quality Issue 20K", like Tableau
               labelParts: { category: true, value: roles.labelRefs.some(r => tfSameField(r, vm.cols[roles.size].ref || {})) } }]
  }];
}

/* ── packed bubbles: one bubble per mark, area = Size → bubble chart; the bubbles are packed here
 *    (largest in the middle) and scaled to the chart's size when the chart XML is written ──── */
const TV_MAX_BUBBLES = 200;

/** @param {ChartContext} ctx @returns {ChartSpec[]} */
function tvPackedBubbleSpec(ctx) {
  const { vm, roles } = ctx;
  if (roles.size < 0) throw new Error("no Size measure for the bubbles");
  const colorDim = roles.color && !roles.color.measureNames && !roles.color.continuous && roles.color.ci >= 0 ? roles.color.ci : -1;
  let dims = [colorDim, ...roles.textDims, ...roles.detailDims].filter((ci, i, a) => ci >= 0 && ci !== roles.size && a.indexOf(ci) === i);
  if (!dims.length) dims = vm.cols.map((c, i) => c.isHeader ? i : -1).filter(i => i >= 0);
  if (!dims.length) throw new Error("no dimension to split the bubbles");
  const cats = tvCategories(vm, dims, false);
  if (cats.count > TV_MAX_BUBBLES) throw new Error(`${cats.count} bubbles – too many to pack in an Excel chart`);
  const sizes = tvSum(vm, cats, roles.size);
  const keep = sizes.map((v, i) => v !== null && v > 0 ? i : -1).filter(i => i >= 0);   // a bubble needs an area
  if (!keep.length) throw new Error("no bubble with a positive size");
  const max = Math.max(...keep.map(i => sizes[i]));
  const radii = keep.map(i => Math.sqrt(sizes[i] / max));
  const { centres, box } = tvPackCircles(radii, 0.03);
  const scale = tvColorScale(vm, roles, "circle");
  let pointColors;
  if (scale && roles.color && roles.color.continuous && roles.color.ci >= 0) {
    const totals = tvSum(vm, cats, roles.color.ci);
    pointColors = keep.map(i => scale(totals[i]));
  } else if (scale && colorDim >= 0) {
    const level = dims.indexOf(colorDim);
    pointColors = keep.map((i, k) => scale(cats.levels[level][i]) || tvHex(TABLEAU_10[k % TABLEAU_10.length]));
  }
  // Tableau's label and tooltip per bubble: its template filled with the bubble's values – Tableau's own formatted
  // value when the bubble is one row, else the rows' sum in the field's format
  const rowsOf = keep.map(() => /** @type {any[][]} */ ([]));
  const at = new Map(keep.map((i, k) => [i, k]));
  vm.rows.forEach(row => { const k = at.get(cats.indexOf(row)); if (k !== undefined) rowsOf[k].push(row); });
  const valueOf = k => ref => {
    const ci = vm.cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
    const rows = rowsOf[k];
    if (ci < 0 || !rows.length) return "";
    if (rows.length === 1 || vm.cols[ci].isHeader) return tfDvText(rows[0][ci]);
    const nums = rows.map(r => tfDvNum(r[ci])).filter(v => v !== null);
    const nf = vm.fmt.numFmtFor ? vm.fmt.numFmtFor(ref) : null;
    return nums.length ? (nf ? tfFormatNumber(nums.reduce((a, b) => a + b, 0), nf) : String(nums.reduce((a, b) => a + b, 0))) : "";
  };
  const text = lines => lines.map(line => line.map(x => x.text).join("")).join("\n");
  const labelTpl = vm.fmt.hasModel ? vm.fmt.labelRunsTemplate() : [];
  const tipTpl = vm.fmt.hasModel ? vm.fmt.tooltipTemplate() : [];
  // without a label template: the label dimension (Label, else Detail, else the innermost level)
  const labelLevel = dims.indexOf(roles.textDims[0] ?? roles.detailDims[0] ?? dims[dims.length - 1]);
  const labels = labelTpl.length ? vm.fmt.markLabelsShown() : roles.textDims.length > 0 || roles.labelRefs.length > 0;
  const labelTexts = labels ? keep.map((i, k) => labelTpl.length ? text(renderTemplate(labelTpl, valueOf(k))) : cats.levels[labelLevel][i]) : undefined;
  const tooltips = tipTpl.length ? keep.map((i, k) => text(renderTemplate(tipTpl, valueOf(k)))) : undefined;
  return [{
    ...tvBaseSpec(vm), kind: "bubble", gridlines: false, axesHidden: true, legend: false,
    xTitle: "Bubble x", numFmt: "General", xNumFmt: "General",
    sizeTitle: tvMeasureLabel(vm, roles.size), sizeNumFmt: tvNumFmt(vm, roles.size), packed: box,
    series: [{
      name: "Bubble y", color: tvMarkColor(roles), x: centres.map(p => p.x), y: centres.map(p => p.y),
      size: keep.map(i => sizes[i]), pointColors, labels, labelTexts, tooltips
    }],
    labelCull: vm.fmt.hasModel ? vm.fmt.markLabelsCulled() : true
  }];
}


/* ═══ charts/model/index.js ═══════════════════════════════════════════════════════════════════════ */
/* Entry point: visual model → Excel chart specs (throws when not representable). */

/* ══════════════════════════════════════════════════════════════════════════
 * TABLEAU VISUAL → EXCEL CHART SPECS
 * ──────────────────────────────────────────────────────────────────────────
 * buildExcelChartSpecs(visualModel, model) reads
 *   – live summary data (already pivoted / merged / sorted by buildViewModel)
 *   – TWB shelves, panes, marks, encodings, colours, number formats, labels
 *   – the live visual specification when no workbook was loaded
 * and returns renderer-neutral chart specs (ChartSpec, src/types.d.ts) for charts/writer.
 * It throws when a visual cannot be represented faithfully, so the caller
 * can fall back to the Tableau image renderer or the data table.
 * ══════════════════════════════════════════════════════════════════════════ */

/**
 * @param {VisualModel} visualModel
 * @param {FormatModel | null} model
 * @returns {ChartSpec[]}
 */
function buildExcelChartSpecs(visualModel, model) {
  const vm = visualModel.viewModel;
  if (!vm || !vm.rows.length) throw new Error("visual has no data rows");
  const roles = tvRoles(vm, model, visualModel.source && visualModel.source.visualSpec);
  const markToken = visualModel.metadata.markToken || "";
  /** @type {ChartContext} */
  const ctx = { vm, roles, markToken, type: visualModel.type, model,
                autoMark: tvAutomaticMark(visualShelfShape(vm.fmt && vm.fmt.sheetModel, model)),
                roundedBar: visualModel.type === VISUAL_TYPES.BAR && isRoundedBar(vm, model) };
  if (roles.source === "summary" && /^(circle|shape|square)$/.test(markToken)) {
    throw new Error("mark layout unknown – load the workbook file for this visual");
  }
  let specs;
  if (visualModel.type === VISUAL_TYPES.HISTOGRAM) specs = tvHistogramSpecs(ctx);
  else if (visualModel.type === VISUAL_TYPES.WATERFALL) specs = tvWaterfallSpec(ctx);
  else if (visualModel.type === VISUAL_TYPES.BOXPLOT) specs = tvBoxPlotSpec(ctx);
  else if (visualModel.type === VISUAL_TYPES.GANTT) specs = tvGanttSpec(ctx);
  else if (visualModel.type === VISUAL_TYPES.TREEMAP) specs = tvTreemapSpec(ctx);
  else if (visualModel.type === VISUAL_TYPES.BUBBLE) specs = tvPackedBubbleSpec(ctx);
  else if (visualModel.type === VISUAL_TYPES.PIE) {
    ctx.doughnut = roles.panes.filter(p => tvMarkToken(p.markClass) === "pie").length >= 2;
    specs = tvPieSpec(ctx);
  } else if (visualModel.type === VISUAL_TYPES.SCATTER && roles.rows.values.length && roles.cols.values.length) {
    specs = tvScatterSpec(ctx);
  } else if (visualModel.type === VISUAL_TYPES.LINE && roles.rows.values.length && roles.cols.values.length &&
             ![...roles.rows.values, ...roles.cols.values].some(v => v.mv)) {
    specs = tvScatterSpec(ctx, { lines: true });         // a line of one measure against another (a Pareto curve)
  } else {
    specs = tvSmallMultipleSpecs(ctx) || tvCartesianSpecs(ctx);
  }
  // reference lines last: a line Excel cannot draw adds its note to the conversion the approximations named
  specs.forEach(s => { s.name = visualModel.metadata.worksheetName; s.rolesSource = roles.source; tvApplyWorkbookAxes(s, ctx); tvNoteApproximations(s, ctx);
                       tvApplyReferenceLines(s, ctx); });
  return specs;
}


/* ═══ export/conditional-format.js ════════════════════════════════════════════════════════════════ */
/* Fallback conditional formatting for tables without Tableau colours. */

function isNumeric(value) {
  if (value === null || value === undefined) return false;
  
  const cleaned = String(value)
    .replace(/,/g, "")
    .replace(/%/g, "")
    .trim();
  
  return cleaned !== "" && !isNaN(Number(cleaned));
}

function shouldSkipConditionalFormatting(columnName, value, fieldRole = null) {
  if (!columnName) return false;
  
  const lowerColName = columnName.toLowerCase();
  
  const categoricalKeywords = [
    'year', 'quarter', 'month', 'date', 'time', 
    'category', 'region', 'product', 'name', 'id',
    'country', 'city', 'state', 'department', 'type',
    'status', 'group', 'segment', 'class'
  ];
  
  const geoKeywords = [
    'latitude', 'lat', 'latitud', 'ycoord', 'y_coord',
    'longitude', 'long', 'lon', 'longitud', 'xcoord', 'x_coord',
    'location', 'geo', 'geography', 'coordinates', 'coords',
    'postal', 'zip', 'zipcode', 'postcode', 'address'
  ];
  
  for (const keyword of geoKeywords) {
    if (lowerColName.includes(keyword)) {
      console.log(`🗺️ Skipping conditional formatting for geographic column: "${columnName}"`);
      return true;
    }
  }
  
  for (const keyword of categoricalKeywords) {
    if (lowerColName.includes(keyword)) {
      return true;
    }
  }
  
  if (value !== undefined && value !== null) {
    const strValue = String(value).trim();
    const numValue = parseFloat(strValue);
    
    if (!isNaN(numValue)) {
      if (lowerColName.includes('lat') && numValue >= -90 && numValue <= 90) {
        console.log(`🗺️ Skipping conditional formatting for latitude column: "${columnName}" (value: ${numValue})`);
        return true;
      }
      
      if (lowerColName.includes('lon') && numValue >= -180 && numValue <= 180) {
        console.log(`🗺️ Skipping conditional formatting for longitude column: "${columnName}" (value: ${numValue})`);
        return true;
      }
      
      if (lowerColName.match(/lat/i) && numValue >= -90 && numValue <= 90) {
        console.log(`🗺️ Skipping conditional formatting for geographic coordinate column: "${columnName}"`);
        return true;
      }
      
      if (lowerColName.match(/lon/i) && numValue >= -180 && numValue <= 180) {
        console.log(`🗺️ Skipping conditional formatting for geographic coordinate column: "${columnName}"`);
        return true;
      }
    }
    
    if (/^\d{4}$/.test(strValue)) {
      const yearNum = parseInt(strValue, 10);
      if (yearNum >= 1900 && yearNum <= 2100) {
        return true;
      }
    }
  }
  
  return false;
}

function isGeographicCoordinate(value, columnName) {
  if (!isNumeric(value)) return false;
  
  const lowerColName = columnName.toLowerCase();
  const numValue = parseFloat(String(value).trim());
  
  if (lowerColName.includes('latitude') || lowerColName.includes('lat')) {
    return numValue >= -90 && numValue <= 90;
  }
  
  if (lowerColName.includes('longitude') || lowerColName.includes('lon')) {
    return numValue >= -180 && numValue <= 180;
  }
  
  if (numValue >= -90 && numValue <= 90 && (lowerColName.includes('coord') || lowerColName.includes('geo'))) {
    return true;
  }
  
  if ((numValue >= -180 && numValue <= -90) || (numValue >= 90 && numValue <= 180)) {
    if (lowerColName.includes('coord') || lowerColName.includes('geo')) {
      return true;
    }
  }
  
  return false;
}

function isNumericForFormatting(value, columnName) {
  if (!isNumeric(value)) return false;
  
  if (columnName) {
    const strValue = String(value).trim();
    if (/^\d{4}$/.test(strValue)) {
      const yearNum = parseInt(strValue, 10);
      if (yearNum >= 1900 && yearNum <= 2100) {
        return false;
      }
    }
    
    if (isGeographicCoordinate(value, columnName)) {
      return false;
    }
  }
  
  return true;
}

function getZeroCenteredColor(value, maxAbs) {
  if (maxAbs === 0) {
    return "FFFFFF";
  }
  
  const intensity = Math.min(Math.abs(value) / maxAbs, 1);
  
  let r = 255;
  let g = 255;
  let b = 0;
  
  if (value > 0) {
    r = Math.round(255 * (1 - intensity));
    g = 255;
    b = Math.round(100 * (1 - intensity));
  } else if (value < 0) {
    r = 255;
    g = Math.round(255 * (1 - intensity));
    b = Math.round(100 * (1 - intensity));
  }
  
  return (
    "FF" +
    r.toString(16).padStart(2, "0") +
    g.toString(16).padStart(2, "0") +
    b.toString(16).padStart(2, "0")
  ).toUpperCase();
}

/* ── FALLBACK: Hardcoded conditional formatting (only used if no XML colors) ── */
function applyConditionalFormattingToTable(worksheet, rows, headers, startRow, startCol, colWidths) {
  headers.forEach((header, colIndex) => {
    const columnName = header.fieldName || header.fieldId || `Column_${colIndex + 1}`;
    const firstValue = rows[0]?.[colIndex]?.formattedValue || rows[0]?.[colIndex]?.value;
    
    let shouldSkip = shouldSkipConditionalFormatting(columnName, firstValue);
    
    if (shouldSkip) {
      console.log(`⚠️ Skipping conditional formatting for column: "${columnName}"`);
      return;
    }
    
    const numericValues = [];
    
    rows.forEach(row => {
      const value = row[colIndex]?.formattedValue || row[colIndex]?.value;
      if (isNumericForFormatting(value, columnName)) {
        let numValue = String(value).replace(/,/g, "").replace(/%/g, "");
        numericValues.push(parseFloat(numValue));
      }
    });
    
    if (numericValues.length === 0) return;
    
    const maxAbs = Math.max(...numericValues.map(v => Math.abs(v)));
    
    rows.forEach((row, rowIndex) => {
      const rawValue = row[colIndex]?.formattedValue || row[colIndex]?.value;
      
      if (!isNumericForFormatting(rawValue, columnName)) return;
      
      let numValue = String(rawValue).replace(/,/g, "").replace(/%/g, "");
      const value = parseFloat(numValue);
      
      const excelRow = startRow + rowIndex;
      const excelCol = startCol + colIndex;
      const cell = worksheet.getCell(excelRow + 1, excelCol + 1);
      
      const color = getZeroCenteredColor(value, maxAbs);
      
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: color }
      };
    });
  });
}


/* ═══ export/mark-cells.js ════════════════════════════════════════════════════════════════════════ */
/* A table built from marks (one constant axis per column) drawn in cells: each column takes the look of the
 * marks in its pane – circles / squares as the cell's fill, arrow shapes as a glyph before the label, bars as
 * Excel data bars – coloured as Tableau colours them. */

const DEFAULT_MARK = "FF4E79A7";                       // Tableau's default mark colour

/* Tableau's "Arrows" shape palettes: the file number is the direction, the same in every set */
const ARROW_BY_INDEX = ["", "↓", "↘", "→", "↗", "↑", "↖", "←", "↙"];
/* Tableau's own shapes (":filled/circle" …) */
const BUILT_IN = { circle: "●", square: "■", triangle: "▲", diamond: "◆", plus: "+", cross: "✕", asterisk: "✱",
                   "dot-circle": "◉", "dot-square": "▣" };

/** the character drawn for a shape: an arrow's direction, or the shape itself @param {string} name */
function tvShapeGlyph(name) {
  if (!name) return null;
  const arrow = name.match(/^Arrows\/\w+-(\d)\.png$/i);
  if (arrow) return ARROW_BY_INDEX[+arrow[1]] || null;
  const own = name.match(/^:(filled\/)?([a-z-]+)$/i);
  if (own) {
    const g = BUILT_IN[own[2].toLowerCase()];
    if (!g) return null;
    return own[1] ? g : ({ "●": "○", "■": "□", "▲": "△", "◆": "◇" })[g] || g;
  }
  return "●";                                           // a custom picture: a dot in its colour
}

/** a number format with a glyph before the value in every section (two at least: no minus before the glyph)
 * @param {string | undefined} fmt @param {string} glyph */
function tvGlyphNumFmt(fmt, glyph) {
  const secs = (tfSplitSections(fmt || "General") || ["General"]).slice(0, 3).map(s => s || "General");
  if (secs.length === 1) secs.push("-" + secs[0]);
  return secs.map(s => `"${glyph} "${s}`).join(";");
}

/**
 * @typedef {{ kind: "fill" | "glyph" | "bar", colorAt: (row: number) => string | null,
 *             glyphAt: (row: number) => string | null, group: string }} MarkColumn
 * kind fill: circles / squares / shapes without a Shape field; glyph: a Shape field (arrows);
 * bar: bars, sized by their label value. group: columns whose bars share one scale (same colour field).
 */

/**
 * How each value column of a table built from marks looks.
 * @param {SheetFormatter} fmt @param {ViewColumn[]} cols @param {any[][]} rows @param {number[]} order
 * @returns {Map<number, MarkColumn>}
 */
function buildMarkPlan(fmt, cols, rows, order) {
  /** @type {Map<number, MarkColumn>} */
  const plan = new Map();
  const panes = fmt.panesInOrder();
  const colOf = ref => cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
  order.forEach(ci => {
    const c = cols[ci];
    if (c.isHeader || !c.ref) return;
    const pp = panes.find(p => p.refs.some(r => tfSameField(r, c.ref)));
    if (!pp) return;
    const pane = pp.pane, mark = String(pane.markClass || "").toLowerCase();
    if (!/^(circle|square|shape|bar)$/.test(mark)) return;
    // colour: the pane's Color field, else its Marks card colour
    const enc = fmt.colorEncodingOf(pane);
    const fixedRule = ((pane.style && pane.style.rules && pane.style.rules.mark) || []).find(f => f.attr === "mark-color" && !f.field);
    const fixed = (fixedRule && tfArgb(fixedRule.value)) || DEFAULT_MARK;
    let colorAt = () => fixed;
    if (enc) {
      const k = colOf(enc.ref);
      if (k >= 0) {
        const vals = rows.map(r => tfIsNull(r[k]) ? null : enc.continuous ? tfDvNum(r[k]) : tfDvText(r[k]));
        const scale = tfBuildColorScale(fmt, enc, vals);
        if (scale) { const colors = vals.map(v => v === null ? null : scale(v)); colorAt = i => colors[i] || fixed; }
      }
    }
    const shapeEnc = pane.encodings.find(e => e.channel === "shape" && e.field);
    const group = enc ? `color:${enc.ref.inner}` : `col:${ci}`;
    if (mark === "bar") {
      plan.set(ci, { kind: "bar", colorAt, glyphAt: () => null, group });
    } else if (mark === "shape" && shapeEnc) {
      const map = fmt.shapeMap(shapeEnc.field) || {};
      const k = colOf(shapeEnc.field);
      const glyphs = rows.map(r => k >= 0 ? tvShapeGlyph(map[tfBucketKey(tfDvText(r[k]))]) : null);
      plan.set(ci, { kind: "glyph", colorAt, glyphAt: i => glyphs[i] || null, group });
    } else {
      plan.set(ci, { kind: "fill", colorAt, glyphAt: () => null, group });
    }
  });
  return plan;
}

/**
 * Data bars for the bar columns, written after the cells: one rule per cell when the colour varies by row.
 * Columns of one group share a scale; a column of negative values (the left half of a diverging bar) grows
 * leftwards from its right edge.
 * @param {import("exceljs").Worksheet} worksheet @param {Map<number, MarkColumn>} plan
 * @param {ViewColumn[]} cols @param {any[][]} rows @param {number[]} order
 * @param {number} firstRow 0-based sheet row of the first data row
 * @param {(k: number) => number} colAt 0-based sheet column of the table's k-th field
 */
function writeMarkBars(worksheet, plan, cols, rows, order, firstRow, colAt) {
  const bars = order.map((ci, k) => ({ ci, k, m: plan.get(ci) })).filter(x => x.m && x.m.kind === "bar");
  /** @type {Map<string, number>} */
  const maxOf = new Map();
  bars.forEach(({ ci, m }) => rows.forEach(r => {
    const v = tfDvNum(r[ci]);
    if (v !== null) maxOf.set(m.group, Math.max(maxOf.get(m.group) || 0, Math.abs(v)));
  }));
  bars.forEach(({ ci, k, m }) => {
    const M = maxOf.get(m.group);
    if (!M) return;
    const values = rows.map(r => tfDvNum(r[ci]));
    const negative = values.some(v => v !== null && v < 0) && !values.some(v => v !== null && v > 0);
    const cfvo = negative ? [{ type: "num", value: -M }, { type: "num", value: 0 }] : [{ type: "num", value: 0 }, { type: "num", value: M }];
    const rule = argb => ({ type: "dataBar", gradient: false, border: false, minLength: 0, maxLength: 100, cfvo,
      color: { argb }, negativeFillColor: { argb }, negativeBarColorSameAsPositive: true, axisPosition: "auto" });
    const addr = i => worksheet.getCell(firstRow + i + 1, colAt(k) + 1).address;
    const colors = values.map((v, i) => v === null ? null : m.colorAt(i));
    const one = colors.filter(Boolean);
    if (one.length && one.every(x => x === one[0])) {
      worksheet.addConditionalFormatting(/** @type {any} */ ({ ref: `${addr(0)}:${addr(rows.length - 1)}`, rules: [rule(one[0])] }));
    } else {
      colors.forEach((argb, i) => {
        if (argb) worksheet.addConditionalFormatting(/** @type {any} */ ({ ref: addr(i), rules: [rule(argb)] }));
      });
    }
  });
}


/* ═══ export/images.js ════════════════════════════════════════════════════════════════════════════ */
/* Dashboard image objects (logos, icons) → pictures on the sheet, from the files packaged in the .twbx. */

/** @typedef {{ type: "png" | "jpeg" | "gif" | "svg" | null, width: number, height: number }} ImageInfo */

/* below this size (px) an image is a divider line or spacer, not a picture */
const MIN_IMAGE_PX = 8;

/**
 * The file's format and pixel size from its header: PNG IHDR, JPEG SOF, GIF screen descriptor, SVG
 * width / height / viewBox. type null = not a format Excel or the browser can draw.
 * @param {Uint8Array} data @param {string} [path] @returns {ImageInfo}
 */
function imageInfo(data, path = "") {
  const b = data || new Uint8Array(0);
  const u16 = i => (b[i] << 8) | b[i + 1], u32 = i => ((b[i] << 24) >>> 0) + (b[i + 1] << 16) + (b[i + 2] << 8) + b[i + 3];
  if (b.length > 24 && b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return { type: "png", width: u32(16), height: u32(20) };
  if (b.length > 10 && b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return { type: "gif", width: b[6] | (b[7] << 8), height: b[8] | (b[9] << 8) };
  if (b.length > 4 && b[0] === 0xff && b[1] === 0xd8) {
    for (let i = 2; i + 9 < b.length;) {
      if (b[i] !== 0xff) { i++; continue; }
      const m = b[i + 1];
      // SOF0–SOF15 carry the frame size (not DHT C4, JPG C8, DAC CC)
      if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) return { type: "jpeg", width: u16(i + 7), height: u16(i + 5) };
      if (m === 0xd8 || m === 0x01 || (m >= 0xd0 && m <= 0xd7)) { i += 2; continue; }
      i += 2 + u16(i + 2);
    }
    return { type: "jpeg", width: 0, height: 0 };
  }
  const head = new TextDecoder().decode(b.subarray(0, 4096));
  if (/\.svg$/i.test(path) || /<svg[\s>]/i.test(head)) {
    const tag = (head.match(/<svg[^>]*>/i) || [""])[0];
    const attr = n => { const m = tag.match(new RegExp(`\\s${n}\\s*=\\s*["']([\\d.]+)(px)?["']`, "i")); return m ? +m[1] : 0; };
    const vb = (tag.match(/viewBox\s*=\s*["']\s*[-\d.]+[\s,]+[-\d.]+[\s,]+([\d.]+)[\s,]+([\d.]+)/i) || []).slice(1).map(Number);
    return { type: "svg", width: attr("width") || vb[0] || 0, height: attr("height") || vb[1] || 0 };
  }
  return { type: null, width: 0, height: 0 };
}

/**
 * How each image object can be drawn on the grid: "backdrop" (a background behind other objects – not
 * exported, it would cover the cells), "overlay" (an icon on top of a worksheet / text box: floats over
 * that block, host = its id), "tiny" (dividers, spacers) or "block" (a logo with its own zone).
 * Only visible objects count, as Tableau shows them.
 * @param {any[]} objects Tableau DashboardObjects
 * @returns {Map<string, { kind: "backdrop" | "overlay" | "tiny" | "block", host?: string }>}
 */
function classifyImageObjects(objects) {
  const box = o => ({ x: o.position.x, y: o.position.y, w: (o.size || {}).width || 0, h: (o.size || {}).height || 0 });
  const inside = (px, py, r) => px >= r.x && px <= r.x + r.w && py >= r.y && py <= r.y + r.h;
  const shown = (objects || []).filter(o => o && o.position && o.isVisible !== false);
  const others = shown.filter(o => /^(worksheet|text|quick-filter|filter|parameter-control|parameter|legend)$/.test(o.type));
  const out = new Map();
  shown.filter(o => o.type === "image").forEach(o => {
    const r = box(o);
    if (r.w < MIN_IMAGE_PX || r.h < MIN_IMAGE_PX) return out.set(String(o.id), { kind: "tiny" });
    if (others.some(p => { const q = box(p); return inside(q.x + q.w / 2, q.y + q.h / 2, r); })) return out.set(String(o.id), { kind: "backdrop" });
    const host = others.find(p => /^(worksheet|text)$/.test(p.type) && inside(r.x + r.w / 2, r.y + r.h / 2, box(p)));
    out.set(String(o.id), host ? { kind: "overlay", host: String(host.id) } : { kind: "block" });
  });
  return out;
}

/**
 * The picture's rectangle inside its box (px), as Tableau draws an image object: Fit Image scales it to
 * fit keeping its proportions, otherwise it keeps its own size (shrunk only when larger than the box);
 * Center Image (on unless switched off) centres it, otherwise it sits top-left.
 * @param {number} natW @param {number} natH @param {number} boxW @param {number} boxH
 * @param {boolean} [scaled] @param {boolean} [centered] undefined = on
 * @returns {{ x: number, y: number, w: number, h: number }}
 */
function fitImage(natW, natH, boxW, boxH, scaled, centered) {
  const nw = natW || boxW, nh = natH || boxH;
  const k = scaled ? Math.min(boxW / nw, boxH / nh) : Math.min(1, boxW / nw, boxH / nh);
  const w = Math.max(1, Math.round(nw * k)), h = Math.max(1, Math.round(nh * k));
  const center = centered !== false;
  return { x: center ? Math.round((boxW - w) / 2) : 0, y: center ? Math.round((boxH - h) / 2) : 0, w, h };
}

const B64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

const B64_INDEX = (() => { const t = new Int16Array(128).fill(-1); for (let i = 0; i < 64; i++) t[B64.charCodeAt(i)] = i; return t; })();

/** base64 text (line breaks, padding allowed) → bytes @param {string} text @returns {Uint8Array} */
function fromBase64(text) {
  const s = String(text || ""), out = new Uint8Array(Math.ceil(s.length * 3 / 4));
  let n = 0, buf = 0, bits = 0;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i), v = c < 128 ? B64_INDEX[c] : -1;
    if (v < 0) continue;                                         // white space, "="
    buf = ((buf << 6) | v) & 0xffffff; bits += 6;
    if (bits >= 8) { bits -= 8; out[n++] = (buf >> bits) & 255; }
  }
  return out.subarray(0, n);
}

/**
 * The custom shapes embedded in the workbook (<external><shapes>: icons used by Shape marks), keyed
 * "shape:<name>" next to the packaged images.
 * @param {string} xmlString the .twb XML @returns {Record<string, WorkbookImage>}
 */
function extractCustomShapes(xmlString) {
  /** @type {Record<string, WorkbookImage>} */
  const out = {};
  const at = String(xmlString || "").lastIndexOf("<external>");
  if (at < 0) return out;
  const unescape = s => s.replace(/&apos;/g, "'").replace(/&quot;/g, '"').replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
  const re = /<shape name=(?:'([^']*)'|"([^"]*)")\s*>([\s\S]*?)<\/shape>/g;
  re.lastIndex = at;
  for (let m = re.exec(xmlString); m; m = re.exec(xmlString)) out["shape:" + unescape(m[1] ?? m[2])] = { data: fromBase64(m[3]) };
  return out;
}

/** @param {Uint8Array} bytes @returns {string} */
function toBase64(bytes) {
  let out = "";
  for (let i = 0; i < bytes.length; i += 3) {
    const n = (bytes[i] << 16) | ((bytes[i + 1] || 0) << 8) | (bytes[i + 2] || 0);
    out += B64[(n >> 18) & 63] + B64[(n >> 12) & 63] + (i + 1 < bytes.length ? B64[(n >> 6) & 63] : "=") + (i + 2 < bytes.length ? B64[n & 63] : "=");
  }
  return out;
}

/* a canvas to draw on, when the page has one (the browser; not the Node test runs) */
function makeCanvas(w, h) {
  try {
    const c = typeof document !== "undefined" && document.createElement ? document.createElement("canvas") : null;
    if (!c || typeof c.getContext !== "function") return null;
    c.width = w; c.height = h;
    return c;
  } catch (e) {
    return null;
  }
}

/* the file drawn at w × h px into a PNG (SVG has to be rasterised for Excel; big photos are shrunk) */
async function rasterize(data, type, w, h) {
  const canvas = makeCanvas(w, h);
  if (!canvas || typeof Image === "undefined" || typeof URL === "undefined" || !URL.createObjectURL) return null;
  const url = URL.createObjectURL(new Blob([data], { type: type === "svg" ? "image/svg+xml" : "image/" + type }));
  try {
    const img = new Image();
    await new Promise((resolve, reject) => { img.onload = resolve; img.onerror = () => reject(new Error("image could not be decoded")); img.src = url; });
    const ctx = canvas.getContext("2d");
    ctx.drawImage(img, 0, 0, w, h);
    const dataUrl = canvas.toDataURL("image/png");
    return dataUrl.slice(dataUrl.indexOf(",") + 1);
  } catch (e) {
    console.warn("[Images] could not draw the image:", e.message);
    return null;
  } finally {
    URL.revokeObjectURL(url);
  }
}

/**
 * The picture for Excel at its drawn size: PNG / JPEG / GIF as packaged; SVG rasterised; a large file
 * shown small (a 1 MB icon in a 40 px zone) re-drawn at twice its size so the workbook stays small.
 * null when it cannot be drawn here (SVG without a canvas, unknown format).
 * @param {{ data: Uint8Array }} file @param {ImageInfo} info @param {number} w @param {number} h drawn size, px
 * @returns {Promise<{ base64: string, extension: "png" | "jpeg" | "gif" } | null>}
 */
async function prepareImage(file, info, w, h) {
  const big = file.data.length > 150000 && info.width > 3 * w && info.height > 3 * h;
  if (info.type === "svg" || (big && info.type !== "gif")) {
    const png = await rasterize(file.data, info.type, Math.max(1, Math.round(w * 2)), Math.max(1, Math.round(h * 2)));
    if (png) return { base64: png, extension: "png" };
    if (info.type === "svg") return null;
  }
  return info.type ? { base64: toBase64(file.data), extension: /** @type {"png" | "jpeg" | "gif"} */ (info.type) } : null;
}

/** a URL Excel opens as a web link ("www.site.com/x" → "https://www.site.com/x"); null for anything else
 * @param {string | undefined} url @returns {string | null} */
function webLink(url) {
  const u = String(url || "").trim();
  if (!u || /^(javascript|data|file|vbscript):/i.test(u)) return null;
  if (/^(https?|mailto):/i.test(u)) return u;
  return /^[\w-]+(\.[\w-]+)+([/?#].*)?$/.test(u) ? "https://" + u : null;
}

/**
 * The sheet position px from the top-left of cell (col, row), as an anchor ExcelJS writes as is:
 * cell + offset in EMU (fractional anchors would be scaled by ExcelJS's own column width guess).
 * @param {number} col @param {number} row 0-based @param {number} x @param {number} y px
 * @param {(c: number) => number} colPx @param {(r: number) => number} rowPx final column widths / row heights
 */
function nativeAnchor(col, row, x, y, colPx, rowPx) {
  let c = col, r = row, dx = Math.max(0, x), dy = Math.max(0, y);
  for (let guard = 0; guard < 500 && colPx(c) > 0 && dx >= colPx(c); guard++) { dx -= colPx(c); c++; }
  for (let guard = 0; guard < 5000 && dy >= rowPx(r); guard++) { dy -= rowPx(r); r++; }
  return { nativeCol: c, nativeColOff: Math.round(dx * 9525), nativeRow: r, nativeRowOff: Math.round(dy * 9525) };
}


/* ═══ export/layout.js ════════════════════════════════════════════════════════════════════════════ */
/* Dashboard layout → Excel grid: positions, sizes and collision handling. */

/* ── Layout constants ─────────────────────────────────────────────────── */
const PX_PER_COL  = 90;

const PX_PER_ROW  = 20;          // one Excel row (15 pt): dashboard y maps 1:1 onto the sheet


const COL_GAP     = 0;

const ROW_GAP     = 0;            // Tableau zones touch; their padding is inside the zone

const ROW_GROUP_THRESHOLD = 8;

const EXCEL_ROW_PX = 20;

// default Excel row (15pt)
const EXCEL_COL_PX = 75;

// a 10-character column – the narrowest width setColumnWidths assigns
const CHART_DATA_SHEET = "Chart Data";

/* pixel box of a chart/image visual: the dashboard zone size minus Tableau's title strip */
function graphicBox(layout, vm) {
  const widthPx = Math.max(40, Math.round((layout && layout.widthPx) || 480));
  const zoneH = Math.round((layout && layout.heightPx) || 320);
  const titleRows = vm.showTitle ? 1 : 0;
  const heightPx = Math.max(60, zoneH - (titleRows ? 28 : 0));
  return { widthPx, heightPx, titleRows,
           // the zone's width on the dashboard grid, so the block lines up with the blocks above / below it
           gridW: layout && layout.widthPx ? layout.gridW : Math.max(2, Math.ceil(widthPx / EXCEL_COL_PX)),
           // the zone's own height on the sheet, so the block below starts where it does on the dashboard
           rows: Math.max(titleRows + 1, Math.round(zoneH / EXCEL_ROW_PX)) };
}

/* zone edges closer than this (px) are one column boundary */
const EDGE_MERGE_PX = 6;

/**
 * Excel columns from the dashboard's own zone edges: every left / right edge of a laid-out block is a column
 * boundary, so each column is as wide (px) as the gap between two edges and every block spans exactly its
 * zone – positions and widths as on the dashboard, whatever the content. A table's fields get their boundaries
 * at the widths the table draws (splitPx, from the zone's left edge; running into free space beside the zone,
 * and scaled down only when that is not enough); KPI tiles split their zone (split). A field or tile crossed by
 * another block's edge spans several columns.
 * @param {{ layout?: any, split?: number[], splitPx?: number[] }[]} items
 * @returns {{ colPx: number[], span: (item: any) => { gridCol: number, gridW: number, cuts: number[] | null,
 *   fields: { offset: number, span: number }[] | null } | null }}
 */
function buildColumnGrid(items) {
  const laid = items.filter(it => it.layout && typeof it.layout.xPx === "number" && it.layout.widthPx > 0);
  if (!laid.length) return { colPx: [], span: () => null };
  const minX = Math.min(...laid.map(it => it.layout.xPx));
  const merge = (list, tol = EDGE_MERGE_PX) => {
    const out = [];
    [...list].sort((a, b) => a - b).forEach(e => { if (!out.length || e - out[out.length - 1] > tol) out.push(e); });
    return out;
  };
  const sideOf = it => [it.layout.xPx - minX, it.layout.xPx - minX + it.layout.widthPx];
  let edges = merge(laid.flatMap(sideOf));
  const nearest = v => edges.reduce((best, e, i) => Math.abs(e - v) < Math.abs(edges[best] - v) ? i : best, 0);
  const extra = [];
  const rowsOverlap = (a, b) => a.layout.yPx < b.layout.yPx + (b.layout.heightPx || 0) - 2 &&
                                b.layout.yPx < a.layout.yPx + (a.layout.heightPx || 0) - 2;
  /* how far a table may run to the right: up to the next block beside it (sharing its rows) – a table's
     columns keep their widths as long as there is room, instead of being squeezed */
  const roomRight = (R, it) => {
    let limit = Infinity;
    laid.forEach(o => {
      if (o === it) return;
      const [l] = sideOf(o);
      if (l >= R - EDGE_MERGE_PX && rowsOverlap(it, o)) limit = Math.min(limit, Math.max(R, l));
    });
    return limit;
  };
  // A table's field boundaries are column edges too. Every block keeps its own edges: where another block's
  // edge falls inside a field, the field spans both columns (merged cells), so tables, KPI tiles and charts all
  // keep their exact widths on one shared set of columns.
  /** @type {Map<any, number[]>} */
  const fieldEdges = new Map();
  laid.forEach(it => {
    if (!it.splitPx || !it.splitPx.length) return;
    const [l, r] = sideOf(it), L = edges[nearest(l)], R = edges[nearest(r)];
    const total = it.splitPx.reduce((a, b) => a + b, 0);
    const room = roomRight(R, it) - L;                                  // squeezed only when that is not enough
    const k = total > room ? room / total : 1;
    let at = L;
    const own = [L];
    it.splitPx.forEach(w => { at += w * k; own.push(Math.round(at)); });
    extra.push(...own);
    fieldEdges.set(it, own);
  });
  // KPI tiles: Tableau gives each the same share of the zone – their boundaries are column edges, whatever other
  // blocks' edges fall in between (a tile then spans several columns)
  /** @type {Map<any, number[]>} */
  const splitCuts = new Map();
  laid.forEach(it => {
    if (it.splitPx && it.splitPx.length) return;
    if (!it.split || it.split.length < 2) return;
    const [l, r] = sideOf(it), L = edges[nearest(l)], R = edges[nearest(r)];
    const total = it.split.reduce((a, b) => a + b, 0);
    let at = L;
    const cuts = it.split.slice(0, -1).map(w => { at += (R - L) * w / total; return Math.round(at); });
    extra.push(...cuts);
    splitCuts.set(it, cuts);
  });
  edges = merge([...edges, ...extra], 2);
  const colPx = edges.slice(1).map((e, i) => e - edges[i]);
  return {
    colPx,
    span: it => {
      if (!laid.includes(it)) return null;
      const [l, r] = sideOf(it), a = nearest(l), b = nearest(r);
      // split blocks: the column (from the block's first) where each next part starts
      const cuts = splitCuts.has(it) ? splitCuts.get(it).map(e => nearest(e) - a) : null;
      // tables: each field's first column (from the block's first) and how many columns it spans
      let fields = null;
      if (fieldEdges.has(it)) {
        const idx = fieldEdges.get(it).map(nearest);
        fields = idx.slice(0, -1).map((c, k) => ({ offset: c - idx[0], span: Math.max(1, idx[k + 1] - c) }));
        fields.forEach((f, k) => { if (k && f.offset < fields[k - 1].offset + fields[k - 1].span) f.offset = fields[k - 1].offset + fields[k - 1].span; });
      }
      return { gridCol: a, gridW: Math.max(1, b - a), cuts, fields };
    }
  };
}

function getExcelColumnName(colIndex) {
  let columnName = "";
  let dividend = colIndex + 1;

  while (dividend > 0) {
    let modulo = (dividend - 1) % 26;
    columnName = String.fromCharCode(65 + modulo) + columnName;
    dividend = Math.floor((dividend - modulo) / 26);
  }

  return columnName;
}

/* ── Range tracker ────────────────────────────────────────────────────── */
function makeRangeTracker() {
  let minR = Infinity, minC = Infinity, maxR = 0, maxC = 0;
  return {
    update(r, c) {
      if (r < minR) minR = r;
      if (c < minC) minC = c;
      if (r > maxR) maxR = r;
      if (c > maxC) maxC = c;
    },
    toRange() {
      return {
        s: { r: minR === Infinity ? 0 : minR, c: minC === Infinity ? 0 : minC },
        e: { r: maxR, c: maxC },
      };
    },
  };
}

/* edges a few px apart (tiled containers with padding) → one grid line, so blocks that line up on
 * the dashboard land in the same row / column */
const SNAP_PX = 8;

/** @param {number[]} values @param {number} [tolerance] @returns {(v: number) => number} */
function snapEdges(values, tolerance = SNAP_PX) {
  const anchor = new Map();
  let current = null;
  [...new Set(values)].sort((a, b) => a - b).forEach(v => {
    if (current === null || v - current > tolerance) current = v;
    anchor.set(v, current);
  });
  return v => anchor.has(v) ? anchor.get(v) : v;
}

/* ── buildLayoutMap ─────────────────────────────────────────────────── */
/* Tableau's DashboardObject: position = { x, y }, size = { width, height } (px) */
/* Tableau's DashboardObjectType → the kinds laid out on the sheet ("filter" / "parameter" as older
 * test fixtures spell them); a DashboardObject's id is its zone id in the workbook */
const OBJECT_KIND = Object.freeze({ worksheet: "worksheet", "quick-filter": "filter", filter: "filter",
                                           "parameter-control": "parameter", parameter: "parameter", text: "text", image: "image" });

function buildLayoutMap(dashboardObjects, titleMap = {}) {
  const map = new Map();
  // images with their own zone (logos) take part in the layout; backgrounds, dividers and icons
  // floating over a sheet do not (export/images.js)
  const images = classifyImageObjects(dashboardObjects || []);

  const positionableObjects = (dashboardObjects || []).filter(
    (obj) => OBJECT_KIND[obj.type] &&
             obj.position &&
             typeof obj.position.x === "number" &&
             typeof obj.position.y === "number" &&
             (obj.type !== "image" || (images.get(String(obj.id)) || {}).kind === "block")
  );

  if (positionableObjects.length === 0) return map;

  let minX = Infinity, minY = Infinity;
  positionableObjects.forEach((obj) => {
    if (obj.position.x < minX) minX = obj.position.x;
    if (obj.position.y < minY) minY = obj.position.y;
  });

  const sizeOf = obj => obj.size || obj.position;
  const snapX = snapEdges(positionableObjects.flatMap(o => [o.position.x, o.position.x + (sizeOf(o).width || 0)]));
  const snapY = snapEdges(positionableObjects.map(o => o.position.y));

  positionableObjects.forEach((obj) => {
    const px = obj.position;
    const { width, height } = sizeOf(obj);
    const gridCol = Math.round((snapX(px.x) - minX) / PX_PER_COL);
    const gridRow = Math.round((snapY(px.y) - minY) / PX_PER_ROW);
    const gridW = Math.max(2, Math.round((width || 180) / PX_PER_COL));
    const gridH = Math.max(3, Math.round((height || 60) / PX_PER_ROW));
    
    const kind = OBJECT_KIND[obj.type];
    let displayName = "";
    if (kind === "text" || kind === "image") displayName = "";
    else if (kind === "worksheet") {
      // a title written with Tableau's <Sheet Name> token names the sheet itself
      displayName = (titleMap[obj.name] && titleMap[obj.name].replace(/<sheet name>/gi, obj.name))
                 || (obj.title && obj.title.trim() ? obj.title.trim() : null)
                 || obj.name;
    } else {
      displayName = (obj.name || "Filter").replace(/[_-]/g, " ");
    }

    // text boxes and images have no unique name – keyed by their zone id
    const key = kind === "text" || kind === "image" ? `${kind}:${obj.id}` : obj.name || `filter_${gridRow}_${gridCol}`;
    // a quick filter can carry its worksheet's name – the worksheet keeps its own position and size
    const entry = {
      type: kind,
      id: obj.id,
      gridRow: Math.max(0, gridRow),
      gridCol: Math.max(0, gridCol),
      gridW,
      // right edge on the grid: tiles in a strip end where the next one starts
      gridRight: width ? Math.round((snapX(px.x + width) - minX) / PX_PER_COL) : gridCol + gridW,
      gridH,
      widthPx: width || null,
      heightPx: height || null,
      xPx: px.x,
      yPx: px.y,
      displayName,
      originalName: obj.name
    };
    if (kind === "filter") map.set(`filter:${obj.id}`, entry);       // the filter card's own spot, by zone id
    if (kind !== "worksheet" && map.has(key) && map.get(key).type === "worksheet") return;
    map.set(key, entry);
  });

  return map;
}

/* ── grouped table rows ─────────────────────────────────────────────────
 * A long table shows its first ROW_GROUP_THRESHOLD rows and collapses the rest into an outline group.
 * Excel hides whole sheet rows, so a chart / image / card beside the table would lose the same rows:
 * the table keeps visible every row such a neighbour uses (it then fills the band like the Tableau
 * zone) and only the rows below are collapsed. Tables beside each other keep their own threshold. */
function isGroupedTable(item) {
  return !!(FORMAT_CONFIG.groupOverflowRows && item.vm && item.type === "worksheet" && !item.isKPI && !item.box &&
            item.vm.rows.length > ROW_GROUP_THRESHOLD);
}

/**
 * Sheet rows a table's header takes: the height Tableau stores for it in 20 px rows (merged down), so a tall
 * header keeps the sheet's rows at 20 px – a single tall row would push the blocks beside the table down.
 * @param {ViewModel} vm
 */
function headerRowSpan(vm) {
  if (!vm.showHeaderRow) return 0;
  // Tableau's stored header height, else the height of the dashboard text boxes drawn as the headers
  const hpx = (vm.fmt && vm.fmt.headerRowHeightPx ? vm.fmt.headerRowHeightPx() : 0) || vm.headerPx || 0;
  return hpx ? Math.max(1, Math.round(hpx / PX_PER_ROW)) : 1;
}

/* sheet row of a table's first data row */
function tableDataStart(item) {
  return item.gridRow + (item.vm.showTitle ? 1 : 0) + headerRowSpan(item.vm);
}

/** sets item.visibleRows on every table that will be grouped @param {ExportItem[]} items */
function setTableVisibleRows(items) {
  items.forEach(t => {
    if (!isGroupedTable(t)) return;
    const start = tableDataStart(t);
    const end = t.gridRow + t.allocatedRows;
    let need = ROW_GROUP_THRESHOLD;
    // the rows Tableau shows in the zone (at the sheet's own row height) stay visible
    if (t.layout && t.layout.heightPx) {
      const headPx = ((t.vm.showTitle ? 1 : 0) + headerRowSpan(t.vm)) * PX_PER_ROW;
      need = Math.max(need, Math.floor((t.layout.heightPx - headPx) / (t.vm.fmt.rowHeightPx() || PX_PER_ROW)));
    }
    items.forEach(o => {
      if (o === t || isGroupedTable(o)) return;
      const oEnd = o.gridRow + (o.allocatedRows || o.rowCount || 0);
      if (o.gridRow < end && oEnd > t.gridRow) need = Math.max(need, oEnd - start);   // shares sheet rows
    });
    t.visibleRows = Math.min(t.vm.rows.length, need);
  });
}

/* rows the writer will produce – used for layout before writing */
function viewModelHeight(vm) {
  const grouped = FORMAT_CONFIG.groupOverflowRows && vm.rows.length > ROW_GROUP_THRESHOLD;
  return (vm.showTitle ? 1 : 0) + headerRowSpan(vm) + vm.rows.length + (grouped ? 1 : 0);
}

function resolveCollisions(zones) {
if (!zones || zones.length === 0) return zones;

// Full row/header/note height a zone actually occupies on the sheet,
// including its own hidden-but-physically-present grouped rows.
function getVisualHeight(zone) {
  return (zone.allocatedRows || zone.rowCount || 5) + ROW_GAP;
}

// ── Group by EXACT gridRow (not a rounded bucket) ──
const rowGroups = new Map();
zones.forEach(zone => {
  const key = zone.gridRow;
  if (!rowGroups.has(key)) rowGroups.set(key, []);
  rowGroups.get(key).push(zone);
});

// ── Horizontal packing within each group ──
const processedGroups = [];
for (const [, group] of rowGroups) {
  group.sort((a, b) => a.gridCol - b.gridCol);
  const minRow = Math.min(...group.map(z => z.gridRow));
  group.forEach(z => { z.gridRow = minRow; });
  let cursor = group[0].gridCol;
  for (const item of group) {
    if (item.gridCol < cursor) item.gridCol = cursor;
    cursor = item.gridCol + item.gridW + COL_GAP;
  }
  const groupBottom = minRow + Math.max(...group.map(z => getVisualHeight(z))) + ROW_GAP;
  processedGroups.push({ minRow, items: group, bottom: groupBottom });
}

// ── Vertical pushes: a row group moves down as a whole, below every earlier group it overlaps
//    (not only the group just above), so blocks that share a top edge on the dashboard stay level ──
processedGroups.sort((a, b) => a.minRow - b.minRow);
// How far `lower` must move to clear the items of `upper` that are really above it (sharing columns). Only
// those count: a short title beside a tall table must not push what sits under the title below the table.
const clearance = (upper, lower) => {
  let need = 0;
  upper.items.forEach(u => lower.items.forEach(l => {
    if (u.gridCol < l.gridCol + l.gridW && u.gridCol + u.gridW > l.gridCol) {
      need = Math.max(need, u.gridRow + getVisualHeight(u) + ROW_GAP - lower.minRow);
    }
  }));
  return need;
};
for (let i = 1; i < processedGroups.length; i++) {
  const lower = processedGroups[i];
  let pushBy = 0;
  for (let j = 0; j < i; j++) pushBy = Math.max(pushBy, clearance(processedGroups[j], lower));
  if (pushBy > 0) {
    lower.items.forEach(z => { z.gridRow += pushBy; });
    lower.minRow += pushBy;
    lower.bottom += pushBy;
  }
}

// ── Full pairwise safety-net scan across ALL zones, not just adjacent groups ──
let fullPassChanged = true;
const MAX_PASSES = 20;
let pass = 0;
while (fullPassChanged && pass < MAX_PASSES) {
  fullPassChanged = false;
  pass++;
  for (let i = 0; i < zones.length; i++) {
    for (let j = i + 1; j < zones.length; j++) {
      const a = zones[i], b = zones[j];
      const upper = a.gridRow <= b.gridRow ? a : b;
      const lower = a.gridRow <= b.gridRow ? b : a;
      const upperBottom = upper.gridRow + getVisualHeight(upper) + ROW_GAP;
      const colOverlap = upper.gridCol < lower.gridCol + lower.gridW &&
                          upper.gridCol + upper.gridW > lower.gridCol;
      const rowOverlap = lower.gridRow < upperBottom;
      if (colOverlap && rowOverlap) {
        lower.gridRow += (upperBottom - lower.gridRow);
        fullPassChanged = true;
      }
    }
  }
}

return zones;
}


/* ═══ export/visual-writers.js ════════════════════════════════════════════════════════════════════ */
/* Tables and KPI cards written with Tableau formatting; widths and autofilters. */

/* =============================================================================
 * TABLEAU-FORMATTED WRITERS
 * All styling below comes from the TWB format model (createSheetFormatter).
 * If no workbook was loaded, the formatter falls back to TABLEAU_DEFAULTS.
 * ============================================================================= */

/**
 * A colour legend at its dashboard position, as Tableau draws it: the field caption in bold, then a coloured
 * square and the value per item.
 * @param {import("exceljs").Worksheet} worksheet @param {ExportItem} item @param {any} rangeTracker
 * @param {Record<number, number>} colWidths
 */
function writeLegendBlock(worksheet, item, rangeTracker, colWidths) {
  const r = item.gridRow, C = item.gridCol;
  const font = tfExcelFont(item.fmt.baseStyle());
  const title = worksheet.getCell(r + 1, C + 1);
  title.value = item.legend.title;
  title.font = { ...font, bold: true };
  rangeTracker.update(r, C);
  item.legend.items.forEach((it, i) => {
    worksheet.getCell(r + 2 + i, C + 1).value = { richText: [
      { text: "■ ", font: { ...font, size: (font.size || 9) + 3, color: { argb: "FF" + (it.color || "CCCCCC") } } },
      { text: it.text, font }
    ] };
    rangeTracker.update(r + 1 + i, C);
  });
  if (!item.layout || !item.layout.widthPx) {                 // no zone width to keep: fit the longest entry
    const chars = Math.max(item.legend.title.length, ...item.legend.items.map(it => it.text.length + 2));
    colWidths[C] = Math.max(colWidths[C] || 0, chars + 2);
  }
}

function writeTableauTitle(worksheet, r, C, text, p, span) {
  const cell = worksheet.getCell(r + 1, C + 1);
  cell.value = text;
  cell.font = tfExcelFont(p);
  cell.alignment = { horizontal: p.hAlign || "left", vertical: "middle", wrapText: false };
  if (p.bgColor) cell.fill = tfExcelFill(p.bgColor);
  if (span > 1) worksheet.mergeCells(r + 1, C + 1, r + 1, C + span);
  if ((p.fontSize || 0) > 11) worksheet.getRow(r + 1).height = Math.round(p.fontSize * 1.6);
}

/* Which columns get Tableau's mark colour, and what colour each row gets.
 * Tableau colours the MARKS of the pane that holds the Color encoding. */
function buildColorPlan(fmt, colInfo, rows) {
  const enc = fmt.colorEncoding();
  const textRefs = fmt.textRefs();
  let markIdx = colInfo.map((c, i) => (!c.isHeader && (c.pivoted || (c.ref && textRefs.some(t => tfSameField(t, c.ref))))) ? i : -1).filter(i => i >= 0);
  if (!markIdx.length) markIdx = colInfo.map((c, i) => c.isHeader ? -1 : i).filter(i => i >= 0);
  if (!enc) return { enc: null, markIdx: new Set(markIdx), colorAt: () => null };

  if (enc.paneRefs && enc.paneRefs.length) {                    // scope to the colour's own pane
    const scoped = markIdx.filter(i => colInfo[i].pivoted
      ? enc.paneRefs.some(r => r.name === "Multiple Values")
      : enc.paneRefs.some(r => tfSameField(r, colInfo[i].ref)));
    if (scoped.length) markIdx = scoped;
  }
  // Color on Measure Values: each measure column is coloured by its own measure's rule (its own legend in
  // Tableau); a measure with no rule of its own keeps its plain style
  if (enc.ref.name === "Multiple Values" && enc.defFor) {
    const perCol = new Map();
    markIdx.forEach(i => {
      const c = colInfo[i];
      if (!c.pivoted || !c.ref) return;
      const def = enc.defFor(c.ref);
      if (!def) return;
      const vals = rows.map(row => tfIsNull(row[i]) ? null : tfDvNum(row[i]));
      const scale = tfBuildColorScale(fmt, { ...enc, ref: c.ref, def, continuous: true }, vals);
      if (scale) perCol.set(i, vals.map(v => (v === null ? null : scale(v)) || null));
    });
    if (perCol.size) {
      return { enc, markIdx: new Set(markIdx), colorAt: () => null,
               colorAtCell: (rowIdx, ci) => (perCol.get(ci) || [])[rowIdx] || null };
    }
  }
  let colorIdx = colInfo.findIndex(c => c.ref && tfSameField(c.ref, enc.ref));
  if (colorIdx < 0) {
    console.log(`[Format] Colour field ${enc.ref.inner} not present in summary data – no mark colours`);
    return { enc, markIdx: new Set(markIdx), colorAt: () => null };
  }
  const pick = dv => tfIsNull(dv) ? null
    : enc.continuous ? (dv.nativeValue !== undefined ? dv.nativeValue : dv.value) : tfDvText(dv);
  const vals = rows.map(row => pick(row[colorIdx]));
  const scale = tfBuildColorScale(fmt, enc, vals);
  const colors = scale ? vals.map(v => scale(v)) : [];
  return { enc, markIdx: new Set(markIdx), colorAt: i => colors[i] || null };
}

function borderSide(d) { return d && d.visible && d.style ? { style: d.style, color: { argb: d.color } } : undefined; }

function writeKPICardStacked(worksheet, vm, originRow, originCol, rangeTracker) {
  let r = originRow;
  const C = originCol;
  const { fmt, cols, rows, order } = vm;
  const plan = buildColorPlan(fmt, cols, rows);
  if (vm.showTitle) {
    writeTableauTitle(worksheet, r, C, vm.title.text, vm.title.props, 2);
    rangeTracker.update(r, C); rangeTracker.update(r, C + 1);
    r++;
  }
  order.forEach((ci, k) => {
    const info = cols[ci];
    tfWriteCell(worksheet.getCell(r + k + 1, C + 1), { formattedValue: info.label, value: info.label }, fmt.fieldLabelStyle(info.ref, true));
    const vp = fmt.markCellStyle(info.ref);
    const color = plan.markIdx.has(ci) ? plan.colorAt(0) : null;
    const extra = {};
    if (color && plan.enc.applyTo === "fill") extra.fill = color;
    if (color && plan.enc.applyTo === "font" && !vp.explicitColor) extra.fontColor = color;
    tfWriteCell(worksheet.getCell(r + k + 1, C + 2), rows[0][ci], vp, extra);
    if (vm.slicerBinding) vm.slicerBinding.bindCell(worksheet.getCell(r + k + 1, C + 2), [rows[0]], ci);
    rangeTracker.update(r + k, C); rangeTracker.update(r + k, C + 1);
  });
  return (vm.showTitle ? 1 : 0) + order.length;
}

/**
 * Each table column's width in px, as the table is drawn: the width Tableau stores for the field (or for
 * Measure Names, or its text-box header), else an estimate from its label and values in their font.
 * @param {ViewModel} vm @returns {number[]} one per vm.order entry
 */
function tableColumnPx(vm) {
  const { fmt, cols, rows, order } = vm;
  return order.map(ci => {
    const info = cols[ci];
    const measureNamesRef = info.pivoted ? tfParseFieldRef("[:Measure Names]") : null;
    const px = fmt.widthPx(info.ref) || (measureNamesRef && fmt.widthPx(measureNamesRef)) || info.zoneWidthPx;
    if (px) return px;
    const style = info.isHeader ? fmt.headerCellStyle(info.ref) : fmt.markCellStyle(info.ref);
    const scale = (style.fontSize || 9) / 11;
    let len = Math.min((info.label || "").length, 24);
    for (let i = 0; i < Math.min(rows.length, 200); i++) {
      len = Math.max(len, tfDvText(rows[i][ci]).split("\n").reduce((m, l) => Math.max(m, l.length), 0));
    }
    return Math.min(50, Math.ceil(len * scale * 1.15) + 2) * 7;
  });
}

/**
 * @param {number} [visibleRows] rows kept visible before the rest is grouped
 * @param {{ offset: number, span: number }[] | null} [fieldCols] each field's first column (from originCol) and how
 *   many columns it spans on the dashboard grid; null = one column per field
 */
function writeRegularTable(worksheet, vm, originRow, originCol, rangeTracker, allTablesInfo, colWidths, exactWidths,
                                  visibleRows = ROW_GROUP_THRESHOLD, fieldCols = null) {
  let r = originRow;
  const C = originCol;
  const { fmt, cols, rows, order } = vm;
  const numCols = order.length;
  // field k's first sheet column, and its last: a field crossed by another block's column edge spans both columns
  const grid = fieldCols && fieldCols.length === numCols ? fieldCols : null;
  const colAt = k => C + (grid ? grid[k].offset : k);
  const lastAt = k => colAt(k) + (grid ? grid[k].span : 1) - 1;
  const lastCol = lastAt(numCols - 1);
  const wideFields = !!grid && grid.some(f => f.span > 1);
  if (!fmt.hasModel) console.log(`[Format] No TWB format info for "${vm.title.text}" – using Tableau defaults`);

  const plan = buildColorPlan(fmt, cols, rows);
  // a table built from marks: each column looks like its pane's marks (fills, arrows, data bars)
  const marks = vm.markTable ? buildMarkPlan(fmt, cols, rows, order) : new Map();
  const tableBg = fmt.tableBackground();                       // Format → Shading → Worksheet
  const rowDiv = fmt.divider("rows");
  const colDiv = fmt.divider("cols");
  const band = fmt.banding();
  const headerIdx = vm.headerOrder.length ? vm.headerOrder
                  : cols.map((c, i) => c.isHeader ? i : -1).filter(i => i >= 0);

  // ── Title (respects the dashboard zone's "Show title") ──
  if (vm.showTitle) {
    const tp = vm.title.props.bgColor ? vm.title.props : { ...vm.title.props, bgColor: tableBg };
    writeTableauTitle(worksheet, r, C, vm.title.text, tp, lastCol - C + 1);
    rangeTracker.update(r, C);
    rangeTracker.update(r, lastCol);
    r++;
  }

  // ── Header row: Measure Names aliases / captions; hidden field labels stay blank ──
  const styles = {};
  order.forEach(ci => { styles[ci] = cols[ci].isHeader ? fmt.headerCellStyle(cols[ci].ref) : fmt.markCellStyle(cols[ci].ref); });
  // ── Column widths: Tableau pixel width → text-box header width → content estimate ──
  const widthOf = {};
  order.forEach((ci, k) => {
    const info = cols[ci];
    const excelCol = colAt(k);
    const measureNamesRef = info.pivoted ? tfParseFieldRef("[:Measure Names]") : null;
    const px = fmt.widthPx(info.ref) || (measureNamesRef && fmt.widthPx(measureNamesRef)) || info.zoneWidthPx;
    if (px) {
      widthOf[ci] = Math.round(px / 7 * 10) / 10;
      exactWidths[excelCol] = Math.max(exactWidths[excelCol] || 0, widthOf[ci]);
      return;
    }
    const scale = (styles[ci].fontSize || 9) / 11;
    let len = Math.min(info.label.length, 24);
    for (let i = 0; i < Math.min(rows.length, 200); i++) {
      const longestLine = tfDvText(rows[i][ci]).split("\n").reduce((m, l) => Math.max(m, l.length), 0);
      len = Math.max(len, longestLine);
    }
    widthOf[ci] = Math.min(50, Math.ceil(len * scale * 1.15) + 2);
    colWidths[excelCol] = Math.max(colWidths[excelCol] || 0, widthOf[ci]);
  });

  const numericCol = ci => { const d = rows.find(rw => !tfIsNull(rw[ci])); return !!d && typeof (d[ci].nativeValue !== undefined ? d[ci].nativeValue : d[ci].value) === "number"; };

  let headerRow = r;
  if (vm.showHeaderRow) {
    const span = headerRowSpan(vm);
    /** @type {[number, number, number, number][]} merged after every header cell is written */
    const merges = [];
    order.forEach((ci, k) => {
      const info = cols[ci];
      let p = fmt.fieldLabelStyle(info.ref, !info.isHeader);
      if (info.labelProps) p = tfMerge(p, { bold: false }, info.labelProps);      // dashboard text box font
      // header label sits over its column: same alignment as the values (numbers right, text left)
      if (!p.hAlign) p.hAlign = info.isHeader ? styles[ci].hAlign : (styles[ci].hAlign || (numericCol(ci) ? "right" : "left"));
      if (/\n/.test(info.label) || info.label.length > 18) p.wrap = true;
      const cell = worksheet.getCell(r + 1, colAt(k) + 1);
      tfWriteCell(cell, { formattedValue: info.label, value: info.label }, p, {
        fill: p.bgColor || tableBg || undefined,
        border: { bottom: borderSide(rowDiv), right: k < numCols - 1 ? borderSide(colDiv) : undefined }
      });
      // a dashboard text box as the header: its runs (fonts, sizes, line breaks) as rich text
      if (info.labelRuns && info.labelRuns.length && !info.headerCovered) {
        cell.value = { richText: tfRichRuns(info.labelRuns, p) };
        cell.alignment = { ...(cell.alignment || {}), wrapText: true };
      }
      // Tableau's header height as 20 px rows merged down, and a header over several panes merged across;
      // the label wraps in it, centred as Tableau draws it (which also keeps it clear of the filter buttons
      // on the last row)
      const across = info.headerSpan || 1;
      const to = lastAt(Math.min(numCols - 1, k + across - 1));     // its last sheet column
      if (!info.headerCovered && (span > 1 || to > colAt(k))) {
        if (span > 1 || across > 1) cell.alignment = { ...(cell.alignment || {}), wrapText: true };
        for (let i = 0; i < span; i++) for (let c = colAt(k); c <= to; c++) {
          if (i || c > colAt(k)) worksheet.getCell(r + 1 + i, c + 1).border = cell.border;
        }
        merges.push([r + 1, colAt(k) + 1, r + span, to + 1]);
      }
      if (span > 1) for (let i = 0; i < span; i++) worksheet.getRow(r + 1 + i).height = 15;
      rangeTracker.update(r + span - 1, lastAt(k));
    });
    merges.forEach(m => worksheet.mergeCells(...m));
    // wrapped labels and no height stored: at most two lines, not a row Excel grows to fit every word
    if (span === 1 && !fmt.headerRowHeightPx() && order.some(ci => /\n/.test(cols[ci].label) || cols[ci].label.length > 18)) {
      worksheet.getRow(r + 1).height = 26;
    }
    headerRow = r + span - 1;                                // the filter buttons go on the header's last row
    r += span;
  }

  const dataStartRow = r;
  const totalRows = rows.length;
  const keepRows = Math.max(ROW_GROUP_THRESHOLD, visibleRows || 0);   // rows a block beside the table uses stay visible
  const needsGrouping = FORMAT_CONFIG.groupOverflowRows && totalRows > keepRows;
  const level = Math.max(1, rowDiv.level || 1);
  const rowHeightPx = fmt.rowHeightPx();

  rows.forEach((row, rowIdx) => {
    const next = rows[rowIdx + 1];
    const dividerHere = !next || !headerIdx.length || headerIdx.slice(0, level).some(hi =>
      tfDvText(row[hi]) !== tfDvText(next[hi]));
    const banded = rowIdx % 2 === 1;                 // banding per VISUAL row (after pivot/merge)
    const rowColor = plan.colorAt(rowIdx);

    order.forEach((ci, k) => {
      let p = styles[ci];
      const ownRef = cols[ci].rowStyleRefs && cols[ci].rowStyleRefs[rowIdx];    // a joined indicator column
      if (ownRef && ownRef !== cols[ci].ref) p = fmt.markCellStyle(ownRef);
      const extra = {
        border: {
          bottom: dividerHere ? borderSide(rowDiv) : undefined,
          right: k < numCols - 1 ? borderSide(colDiv) : undefined
        }
      };
      // fill precedence: mark colour > own shading > banding > worksheet background
      const bandFill = banded ? (cols[ci].isHeader ? band.header : band.pane) : null;
      extra.fill = p.bgColor || bandFill || tableBg || undefined;
      const mk = marks.get(ci);
      const markColor = plan.colorAtCell ? plan.colorAtCell(rowIdx, ci) : rowColor;
      if (markColor && plan.markIdx.has(ci) && !mk) {
        if (plan.enc.applyTo === "fill") extra.fill = markColor;
        else if (!p.explicitColor) extra.fontColor = markColor;
      }
      const cellColor = vm.cellFill ? vm.cellFill(rowIdx, ci) : null;     // a matrix: each cell its own mark's colour
      if (cellColor) extra.fill = cellColor;
      let dv = row[ci];
      const text = tfDvText(dv);
      if (mk && !tfIsNull(dv)) {
        const color = mk.colorAt(rowIdx);
        if (mk.kind === "fill" && color) extra.fill = color;              // circle / square: the cell is the mark
        if (mk.kind === "glyph") {                                        // arrow shape before the label, in its colour
          if (color && !p.explicitColor) extra.fontColor = color;
          const glyph = mk.glyphAt(rowIdx);
          const native = dv.nativeValue !== undefined ? dv.nativeValue : dv.value;
          if (glyph && typeof native === "number") p = { ...p, numFmt: tvGlyphNumFmt(p.numFmt || inferExcelNumFmt(text, native) || "General", glyph) };
          else if (glyph) dv = { ...dv, formattedValue: `${glyph} ${text}` };
        }
      }
      if (cols[ci].link) {                                      // URL action → clickable cell
        const url = cols[ci].link.expression.replace(/<([^<>]+)>/g, (m, inner) => {
          const part = cols[ci].link.parts.find(x => ("[" + x.token.inner + "]") === inner || x.token.raw === inner);
          return part && part.ci >= 0 ? tfDvText(row[part.ci]) : "";
        }).trim();
        if (/^(https?:|mailto:|ftp:)/i.test(url)) {
          extra.hyperlink = url;
          extra.hyperlinkText = FORMAT_CONFIG.linkText === "url" ? url : FORMAT_CONFIG.linkText;
          p = { ...p, underline: true, color: p.explicitColor ? p.color : "FF0563C1" };
        }
      }
      const pp = /\n/.test(text) ? { ...p, wrap: true } : p;    // wrap only on real line breaks
      const cell = worksheet.getCell(r + 1, colAt(k) + 1);
      tfWriteCell(cell, dv, pp, extra);
      if (vm.slicerBinding && !cols[ci].isHeader) vm.slicerBinding.bindCell(cell, [row], ci);
      if (lastAt(k) > colAt(k)) {                                // a field over several columns: one merged cell
        for (let c = colAt(k) + 1; c <= lastAt(k); c++) worksheet.getCell(r + 1, c + 1).style = cell.style;
        worksheet.mergeCells(r + 1, colAt(k) + 1, r + 1, lastAt(k) + 1);
      }
      rangeTracker.update(r, lastAt(k));
    });

    const excelRow = worksheet.getRow(r + 1);
    if (rowHeightPx) excelRow.height = Math.round(rowHeightPx * 0.75);   // only a height Tableau stores
    if (needsGrouping && rowIdx >= keepRows) {
      excelRow.outlineLevel = 1;
      excelRow.hidden = true;
    }
    r++;
  });

  if (marks.size) writeMarkBars(worksheet, marks, cols, rows, order, dataStartRow, colAt);

  if (needsGrouping) {
    const hiddenCount = totalRows - keepRows;
    const noteCell = worksheet.getCell(r + 1, C + 1);
    noteCell.value = `${hiddenCount} rows are hidden — use the row group controls [+] / [-] on the left to expand or collapse`;
    noteCell.font = { italic: true, size: 9, color: { argb: "FF888888" } };
    noteCell.alignment = { horizontal: "left", vertical: "middle" };
    rangeTracker.update(r, C);
    r++;
    worksheet.properties.outlineProperties = { summaryBelow: false, summaryRight: false };
  }

  if (!plan.enc && FORMAT_CONFIG.applyFallbackHeatmap) {
    applyConditionalFormattingToTable(worksheet, rows.map(rw => order.map(ci => rw[ci])),
      order.map(ci => ({ fieldName: cols[ci].name })), dataStartRow, C, {});
  }

  // no filter buttons on a table built from marks, nor on merged fields (a button would sit inside the field)
  if (allTablesInfo && !vm.markTable && !wideFields) {
    allTablesInfo.push({ name: vm.title.text, headerRow: vm.showHeaderRow ? headerRow : dataStartRow,
      leftCol: C, rightCol: lastCol,
      bottomRow: dataStartRow + totalRows - 1,             // last DATA row – the grouping note stays outside the filter
      rowCount: totalRows, hasHeader: vm.showHeaderRow });
  }
  return r - originRow;
}

function applyAutoFilters(worksheet, allTablesInfo) {
  if (FORMAT_CONFIG.autoFilter === "none" || !allTablesInfo || !allTablesInfo.length) return;
  // Excel allows a single autofilter per sheet → put it on the largest table that has a header row
  const t = allTablesInfo.filter(x => x.hasHeader).sort((a, b) => b.rowCount - a.rowCount)[0];
  if (!t) return;
  worksheet.autoFilter = `${getExcelColumnName(t.leftCol)}${t.headerRow + 1}:${getExcelColumnName(t.rightCol)}${t.bottomRow + 1}`;
}

function setColumnWidths(worksheet, colWidths, exactWidths = {}) {
  const keys = [...Object.keys(colWidths), ...Object.keys(exactWidths)].map(Number);
  const maxColIdx = keys.length ? Math.max(...keys) : 0;
  for (let ci = 0; ci <= maxColIdx; ci++) {
    if (exactWidths[ci]) { worksheet.getColumn(ci + 1).width = exactWidths[ci]; continue; }
    worksheet.getColumn(ci + 1).width = Math.min((colWidths[ci] || 10), 50);
  }
}


/* ═══ export/cell-writers.js ══════════════════════════════════════════════════════════════════════ */
/* Low-level cell writers: titles, headers, data cells, filter tables. */

/* ── Quick filter / parameter control ───────────────────────────────────
 * As Tableau draws one: its caption, then the value(s) in a white box with a thin grey border, across the
 * control's zone. */
function writeIndividualFilterTable(worksheet, filterName, filterValues, originRow, originCol, rangeTracker, width = 1) {
  const C = originCol, W = Math.max(1, width);
  const cleanName = filterName.replace(/_(Filter|filter)_\d+$/, "").replace(/_(Values|values)_\d+$/, "");
  const font = tfExcelFont({ ...TABLEAU_DEFAULTS.worksheet, fontSize: 9, color: "FF333333" });     // Tableau fonts → their substitute
  const span = r => { if (W > 1) { try { worksheet.mergeCells(r + 1, C + 1, r + 1, C + W); } catch (e) { /* already merged */ } } };
  const caption = worksheet.getCell(originRow + 1, C + 1);
  caption.value = cleanName;
  caption.font = { ...font, bold: true };
  caption.alignment = { horizontal: "left", vertical: "bottom" };
  span(originRow);
  rangeTracker.update(originRow, C);
  rangeTracker.update(originRow, C + W - 1);
  const values = Array.isArray(filterValues) ? filterValues : [filterValues];
  /** @type {Partial<import("exceljs").Border>} */
  const edge = { style: "thin", color: { argb: "FFBFBFBF" } };
  values.forEach((value, idx) => {
    const r = originRow + 1 + idx;
    for (let c = C; c < C + W; c++) {
      const cell = worksheet.getCell(r + 1, c + 1);
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
      cell.border = { top: idx === 0 ? edge : undefined, bottom: idx === values.length - 1 ? edge : undefined,
                      left: c === C ? edge : undefined, right: c === C + W - 1 ? edge : undefined };
    }
    const cell = worksheet.getCell(r + 1, C + 1);
    cell.value = value;
    cell.font = font;
    cell.alignment = { horizontal: "left", vertical: "middle", shrinkToFit: true };
    span(r);
    rangeTracker.update(r, C + W - 1);
  });
  return 1 + values.length;
}

function writeDashboardTitle(worksheet, dashboardName, originRow, originCol, rangeTracker, titleProps, titleRuns) {
  let r = originRow;
  const C = originCol;
  const p = titleProps || tfMerge(TABLEAU_DEFAULTS.dashTitle);
  writeTableauTitle(worksheet, r, C, dashboardName, p, 5);
  if (titleRuns && titleRuns.length) {                    // text-box title → rich text, one run per Tableau run
    const rich = tfRichRuns(titleRuns, tfMerge(TABLEAU_DEFAULTS.worksheet));
    const cell = worksheet.getCell(r + 1, C + 1);
    cell.value = { richText: rich };
    cell.alignment = { horizontal: (titleRuns[0].props.hAlign || "left"), vertical: "top", wrapText: true };
    const lineSizes = [];
    let cur = 0;
    rich.forEach(x => x.text.split("\n").forEach((seg, i) => {
      if (i > 0) { lineSizes.push(cur); cur = 0; }
      if (seg) cur = Math.max(cur, x.font.size || 9);
    }));
    lineSizes.push(cur);
    worksheet.getRow(r + 1).height = Math.ceil(lineSizes.reduce((a, b) => a + (b || 9) * 1.3, 0) + 4);
  }
  rangeTracker.update(r, C);
  rangeTracker.update(r, C + 4);
  r++;
  r++;
  return r - originRow;
}


/* ═══ export/sheet-data.js ════════════════════════════════════════════════════════════════════════ */
/* Fetching worksheet summary data, filters and filter-value sheets from Tableau. */

async function fetchAllSheetsData(sheets, concurrency = 4) {
const results = [];
const queue = [...sheets];

async function worker() {
  while (queue.length) {
    const sheet = queue.shift();
    try {
      const [dataResult, visualSpecResult] = await Promise.allSettled([
        // every mark, not only the selected ones: with a mark selected Tableau returns just that mark otherwise
        sheet.getSummaryDataAsync({ ignoreSelection: true }),
        typeof sheet.getVisualSpecificationAsync === "function"
          ? sheet.getVisualSpecificationAsync()
          : Promise.reject(new Error("getVisualSpecificationAsync is unavailable"))
      ]);
      if (dataResult.status === "rejected") throw dataResult.reason;
      results.push({
        sheet,
        data: dataResult.value,
        visualSpec: visualSpecResult.status === "fulfilled" ? visualSpecResult.value : null,
        visualSpecError: visualSpecResult.status === "rejected" ? visualSpecResult.reason : null,
        error: null
      });
    } catch (err) {
      results.push({ sheet, data: null, error: err });
    }
  }
}

const workers = Array(concurrency).fill().map(() => worker());
await Promise.all(workers);
return results;
}

/* ── Filter Extraction ─────────────────────────────────────────────────── */
async function extractFilterValuesPerField(sheets) {
  const filterMap = new Map();
  
  for (const worksheet of sheets) {
    let filters = [];
    
    try {
      filters = await worksheet.getFiltersAsync();
    } catch (e) {
      console.warn(`Could not get filters for ${worksheet.name}:`, e);
      continue;
    }
    
    filters.forEach(filter => {
      const ignoredFields = ["Measure Names", "Measure Values"];
      if (ignoredFields.includes(filter.fieldName) || /^Month\(/i.test(filter.fieldName)) {
        return;
      }
      
      if (!filterMap.has(filter.fieldName)) {
        filterMap.set(filter.fieldName, {
          values: new Set(),
          worksheetNames: new Set()
        });
      }
      
      const filterData = filterMap.get(filter.fieldName);
      filterData.worksheetNames.add(worksheet.name);
      
      let value = "";
      
      if (filter.filterType === "categorical") {
        const appliedValues = (filter.appliedValues || []).map(v => v.formattedValue);
        appliedValues.forEach(v => filterData.values.add(v));
        value = appliedValues.join(", ");
      } else if (filter.filterType === "range") {
        // date ranges arrive as "6/1/2025 12:00:00 AM" – a midnight time is not part of the filter
        const end = v => String((v && v.formattedValue) || "").replace(/\s+12:00:00\s*AM$|\s+00:00:00$/i, "");
        value = end(filter.minValue) + " - " + end(filter.maxValue);
        filterData.values.add(value);
      } else {
        value = filter.filterType;
        filterData.values.add(value);
      }
    });
  }
  
  const result = {};
  for (const [fieldName, data] of filterMap.entries()) {
    result[fieldName] = Array.from(data.values);
  }
  
  return result;
}

/** The marks selected on a sheet (a selected pie slice is outlined, as Tableau shows it); [] when unavailable.
 * @param {any} sheet Tableau worksheet @returns {Promise<any[]>} */
async function fetchSelectedMarks(sheet) {
  try {
    const marks = typeof sheet.getSelectedMarksAsync === "function" ? await sheet.getSelectedMarksAsync() : null;
    return (marks && marks.data) || [];
  } catch (e) {
    console.warn(`[Export] selected marks of "${sheet.name}" unavailable: ${e.message}`);
    return [];
  }
}

/* ── Check if worksheet is a filter value table ───────────────────────── */
function isFilterValueWorksheet(sheetName, summaryData) {
  if (/filter[_\- ]?\d+/i.test(sheetName) || 
      /values[_\- ]?\d+/i.test(sheetName) ||
      /_filter_\d+/i.test(sheetName)) {
    return true;
  }
  
  if (summaryData && summaryData.columns && summaryData.columns.length === 1) {
    const firstColumnName = summaryData.columns[0]?.fieldName || "";
    if (firstColumnName === "Values" || firstColumnName === "VALUE") {
      return true;
    }
  }
  
  return false;
}


/* ═══ export/kpi-card.js ══════════════════════════════════════════════════════════════════════════ */
/* KPI cards and dashboard text boxes: formatted text in a block that fills its dashboard zone. */

/* Tableau's formatted text: "Æ" marks a line break – usually followed by a newline, sometimes only a space */
const LINE_BREAK = /Æ[ \t]*(?:\r?\n)?|\r?\n/;

/* white edges between neighbouring cards – Tableau's tile gutters are narrower than a grid column */
/** @type {Partial<import("exceljs").Border>} */
const KPI_GUTTER = { style: "medium", color: { argb: "FFFFFFFF" } };

/**
 * @typedef {{ text?: string, ci?: number, props: Record<string, any> }} KpiSegment  static text or a value column
 * @typedef {{ segments: KpiSegment[], hAlign: "left" | "center" | "right", sizePt: number }} KpiLine
 * @typedef {{ lines: KpiLine[] }} KpiTile
 * @typedef {{ title: { text: string, props: Record<string, any> } | null, tiles: KpiTile[], background: string | null,
 *             padTop: number, padBottom: number, lineHeightsPt: number[], rows: number, widthPx: number[],
 *             zoneWidthPx: number | null }} KpiCard
 */

/**
 * Tableau formatted text (custom label, text box) → lines of segments. Field tokens "<[ds].[field]>"
 * may be split over several runs ("<", "[Parameters].[P 10]", "> Stores by <" …), so they are found
 * in the joined text; each piece keeps the font of the run it comes from.
 * @param {TextRun[]} runs
 * @param {(inner: string, props: Record<string, any>) => KpiSegment | null} token a "<…>" token → segment (null = drop)
 * @param {(props: Record<string, any>) => Record<string, any>} styled static text props from a run's props
 * @returns {KpiSegment[][]}
 */
function formattedLines(runs, token, styled) {
  const full = runs.map(r => r.text).join("");
  const starts = [];
  runs.reduce((pos, r) => { starts.push(pos); return pos + r.text.length; }, 0);
  const runAt = i => { let k = 0; while (k + 1 < starts.length && starts[k + 1] <= i) k++; return k; };
  const lines = [[]];
  const plain = (from, to) => {                              // static text, cut at run and line boundaries
    for (let i = from; i < to;) {
      const k = runAt(i), end = Math.min(to, k + 1 < starts.length ? starts[k + 1] : full.length);
      full.slice(i, end).split(LINE_BREAK).forEach((part, n) => {
        if (n > 0) lines.push([]);
        if (part) lines[lines.length - 1].push({ text: part, props: styled(runs[k].props) });
      });
      i = end;
    }
  };
  let last = 0;
  for (const m of full.matchAll(/<([^<>]+)>/g)) {
    plain(last, m.index);
    const seg = token(m[1], runs[runAt(m.index + 1)].props);  // the run holding the field name
    if (seg) lines[lines.length - 1].push(seg);
    last = m.index + m[0].length;
  }
  plain(last, full.length);
  while (lines.length > 1 && !lines[lines.length - 1].length) lines.pop();          // trailing line breaks
  return lines;
}

/**
 * The tile as Tableau draws it: its custom label (lines, fonts, sizes, colours, alignment), or
 * Tableau's default text layout; sized to fill the dashboard zone. null = Measure Names on Rows –
 * Tableau draws that as a two-column text table, which the stacked label | value writer matches.
 * @param {ViewModel} vm @param {string} sheetName @param {{ heightPx?: number, widthPx?: number } | null} [layout]
 * @returns {KpiCard | null}
 */
function buildKpiCard(vm, sheetName, layout) {
  const fmt = vm.fmt;
  const sheet = fmt.sheetModel;
  const visible = vm.order.filter(ci => !vm.cols[ci].isHeader);
  const colOf = ref => vm.cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
  const labelPane = sheet ? sheet.panes.find(p => p.labelRuns.length) : null;
  /** @type {KpiTile[]} */
  let tiles;
  if (labelPane) {
    const styled = props => tfMerge(fmt.markCellStyle(null), props);
    const lines = formattedLines(labelPane.labelRuns, (inner, props) => {
      if (/^sheet name$/i.test(inner)) return { text: sheetName, props: styled(props) };
      const ref = /^\[/.test(inner) ? tfParseFieldRef(inner) : null;
      const ci = ref ? colOf(ref) : -1;
      return ci >= 0 ? { ci, props: tfMerge(fmt.markCellStyle(vm.cols[ci].ref), props) } : null;
    }, styled);
    tiles = [{ lines: lines.map(segments => kpiLine(segments, fmt)) }];
  } else {
    const onShelf = shelf => !!sheet && sheet[shelf].some(r => r.name === "Measure Names");
    if (onShelf("rows")) return null;
    if (onShelf("cols")) {                                  // Measure Names on Columns: a tile per measure, name above value
      const nameStyle = fmt.headerCellStyle(tfParseFieldRef("[:Measure Names]"));
      tiles = visible.map(ci => ({ lines: [
        kpiLine([{ text: vm.cols[ci].label || vm.cols[ci].name, props: nameStyle }], fmt),
        kpiLine([{ ci, props: fmt.markCellStyle(vm.cols[ci].ref) }], fmt)
      ] }));
    } else {                                                // Text = the measures: one value per line
      tiles = [{ lines: visible.map(ci => kpiLine([{ ci, props: fmt.markCellStyle(vm.cols[ci].ref) }], fmt)) }];
    }
  }
  const depth = Math.max(...tiles.map(t => t.lines.length));
  const lineHeightsPt = Array.from({ length: depth }, (_, i) =>
    Math.ceil(Math.max(...tiles.map(t => (t.lines[i] ? t.lines[i].sizePt : 9))) * 1.35 + 2));
  const title = vm.showTitle ? vm.title : null;
  // padding rows (default 20 px) fill the zone's height around the lines, as Tableau centres the label
  const used = (title ? 20 : 0) + lineHeightsPt.reduce((s, h) => s + h * 4 / 3, 0);
  const pad = layout && layout.heightPx ? Math.max(0, Math.round((layout.heightPx - used) / 20)) : 0;
  const vAlign = fmt.markCellStyle(null).vAlign || "middle";
  const padTop = vAlign === "top" ? 0 : vAlign === "bottom" ? pad : Math.floor(pad / 2);
  const background = fmt.tableBackground();
  const widthPx = tiles.map(t => Math.max(60, ...t.lines.map(l =>
    l.segments.reduce((w, s) => w + textOf(s, vm).length * (s.props.fontSize || 9) * 4 / 3 * 0.58, 0) + 18)));
  return { title, tiles, background, padTop, padBottom: pad - padTop, lineHeightsPt,
           rows: (title ? 1 : 0) + pad + depth, widthPx, zoneWidthPx: (layout && layout.widthPx) || null };
}

/**
 * A dashboard text box (banner, title, note) as Tableau draws it: its runs – fonts, sizes, colours,
 * alignment – on the zone's background colour, filling the zone. Only parameter values can be fields
 * in a dashboard text box; other tokens (page name …) have no value in an export.
 * @param {DashboardZone} zone @param {{ heightPx?: number, widthPx?: number } | null} layout @param {FormatModel | null} model
 * @returns {KpiCard | null}
 */
function buildTextCard(zone, layout, model) {
  const base = tfMerge(TABLEAU_DEFAULTS.worksheet);
  const styled = props => tfMerge(base, props);
  const lines = formattedLines(zone.runs || [], (inner, props) => {
    const ref = /^\[/.test(inner) ? tfParseFieldRef(inner) : null;
    const info = ref && ref.ds === "Parameters" ? tfFieldInfo(model, ref) || {} : null;
    const text = info ? String(info.alias || info.value || "").replace(/^"|"$/g, "") : "";
    return text ? { text, props: styled(props) } : null;
  }, styled);
  if (!lines.some(l => l.some(s => String(s.text).trim()))) return null;
  /** @type {KpiLine[]} */
  const kpiLines = lines.map(segments => ({ segments,
    hAlign: /** @type {KpiLine["hAlign"]} */ ((segments.find(s => s.props.hAlign) || { props: { hAlign: "left" } }).props.hAlign),
    sizePt: Math.max(9, ...segments.map(s => s.props.fontSize || 9)) }));
  const lineHeightsPt = kpiLines.map(l => Math.ceil(l.sizePt * 1.35 + 2));
  const used = lineHeightsPt.reduce((s, h) => s + h * 4 / 3, 0);
  const pad = layout && layout.heightPx ? Math.max(0, Math.round((layout.heightPx - used) / 20)) : 0;
  const widthPx = [Math.max(40, ...kpiLines.map(l => l.segments.reduce((w, s) => w + s.text.length * (s.props.fontSize || 9) * 4 / 3 * 0.58, 0) + 12))];
  return { title: null, tiles: [{ lines: kpiLines }], background: (zone.style && zone.style.bgColor) || null,
           padTop: Math.floor(pad / 2), padBottom: pad - Math.floor(pad / 2), lineHeightsPt, rows: pad + kpiLines.length,
           widthPx, zoneWidthPx: (layout && layout.widthPx) || null };
}

/** a line's alignment: the pane's Format → Alignment, else the label editor's, else centred
 * @returns {KpiLine} */
function kpiLine(segments, fmt) {
  const cellAlign = tfMerge(fmt.markCellStyle(null)).hAlign;
  const runAlign = (segments.find(s => s.props.hAlign) || { props: {} }).props.hAlign;
  const hAlign = /** @type {KpiLine["hAlign"]} */ (/^(left|right)$/.test(cellAlign || runAlign || "") ? (cellAlign || runAlign) : "center");
  return { segments, hAlign, sizePt: Math.max(9, ...segments.map(s => s.props.fontSize || 9)) };
}

function textOf(s, vm) { return s.ci !== undefined ? tfDvText(vm.rows[0][s.ci]) : String(s.text || ""); }

/**
 * Writes the card over gridW columns from (originRow, originCol); returns the rows used.
 * @param {import("exceljs").Worksheet} worksheet @param {KpiCard} card @param {ViewModel | null} vm null for a text box
 * @param {number} originRow @param {number} originCol @param {number} gridW
 * @param {{ update(r: number, c: number): void }} rangeTracker @param {Record<number, number>} colWidths
 * @param {number[] | null} [cuts] the column (from originCol) where each next tile starts – equal shares of the zone
 */
function writeKpiCard(worksheet, card, vm, originRow, originCol, gridW, rangeTracker, colWidths, cuts = null) {
  let r = originRow;
  const C = originCol;
  if (card.title) {
    writeTableauTitle(worksheet, r, C, card.title.text, card.title.props, gridW);
    rangeTracker.update(r, C); rangeTracker.update(r, C + gridW - 1);
    r++;
  }
  const depth = card.lineHeightsPt.length;
  const bodyRows = card.padTop + depth + card.padBottom;
  if (card.background) {
    for (let rr = r; rr < r + bodyRows; rr++) {
      for (let c = C; c < C + gridW; c++) {
        const cell = worksheet.getCell(rr + 1, c + 1);
        cell.fill = tfExcelFill(card.background);
        if (c === C || c === C + gridW - 1) cell.border = { ...(c === C ? { left: KPI_GUTTER } : {}), ...(c === C + gridW - 1 ? { right: KPI_GUTTER } : {}) };
      }
    }
  }
  const plan = vm ? buildColorPlan(vm.fmt, vm.cols, vm.rows) : null;
  const n = card.tiles.length;
  const span = Math.max(1, Math.floor(gridW / n));
  // the tile's columns: between its boundaries on the grid (equal shares of the zone), else an equal count
  const byCuts = cuts && cuts.length === n - 1 && cuts.every((c, i) => c > (i ? cuts[i - 1] : 0) && c < gridW);
  card.tiles.forEach((tile, k) => {
    const c0 = byCuts ? C + (k ? cuts[k - 1] : 0) : C + k * span;
    const c1 = byCuts ? C + (k < n - 1 ? cuts[k] : gridW) - 1 : k === n - 1 ? C + gridW - 1 : c0 + span - 1;
    // as wide as the tile's zone (and never narrower than its text); Excel width units ≈ 7 px
    const tilePx = Math.max(card.widthPx[k], card.zoneWidthPx ? card.zoneWidthPx / n : 0);
    const perCol = Math.ceil(tilePx / (c1 - c0 + 1) / 7);
    for (let c = c0; c <= c1; c++) colWidths[c] = Math.max(colWidths[c] || 0, perCol);
    tile.lines.forEach((ln, i) => {
      const row = r + card.padTop + i;
      if (c1 > c0) worksheet.mergeCells(row + 1, c0 + 1, row + 1, c1 + 1);
      const cell = worksheet.getCell(row + 1, c0 + 1);
      const only = ln.segments.length === 1 ? ln.segments[0] : null;
      if (only && only.ci !== undefined) {                  // a single value: keep the number and its format
        const color = plan && plan.enc && plan.markIdx.has(only.ci) && plan.enc.applyTo === "font" && !only.props.explicitColor ? plan.colorAt(0) : null;
        tfWriteCell(cell, vm.rows[0][only.ci], only.props, { fill: card.background || undefined, fontColor: color || undefined });
        if (vm.slicerBinding) vm.slicerBinding.bindCell(cell, [vm.rows[0]], only.ci);
      } else if (ln.segments.length) {
        cell.value = { richText: ln.segments.map(s => ({ text: textOf(s, vm), font: tfExcelFont(s.props) })) };
        if (card.background) cell.fill = tfExcelFill(card.background);
      }
      cell.alignment = { horizontal: ln.hAlign, vertical: "middle", indent: ln.hAlign === "center" ? 0 : 1, wrapText: false };
      // a merged range has one style (ExcelJS): both outer edges go on it, Excel draws only the outline
      if (card.background) cell.border = { ...(k === 0 ? { left: KPI_GUTTER } : {}), ...(k === n - 1 ? { right: KPI_GUTTER } : {}) };
      const xr = worksheet.getRow(row + 1);
      xr.height = Math.max(xr.height || 15, card.lineHeightsPt[i]);
      rangeTracker.update(row, c0); rangeTracker.update(row, c1);
    });
  });
  rangeTracker.update(r + bodyRows - 1, C + gridW - 1);
  return (card.title ? 1 : 0) + bodyRows;
}


/* ═══ export/backgrounds.js ═══════════════════════════════════════════════════════════════════════ */
/* Dashboard shading and container / object backgrounds (Layout pane → Background) → the colour of the
 * cells around and under the blocks on the sheet; thin coloured zones (dividers, accent lines) → borders. */

/**
 * A block on the sheet (0-based rows / columns, inclusive) and the zone it comes from.
 * own: the block's own opaque colour (ARGB); "white" = a sheet Tableau draws white (left unfilled);
 * undefined = see-through, the zone behind it shows.
 * @typedef {{ zoneId: string, top: number, left: number, bottom: number, right: number, own?: string }} BgBlock
 */

/**
 * A divider: along the top of row `at` ("h", columns from–to) or the left of column `at` ("v", rows from–to).
 * @typedef {{ dir: "h" | "v", at: number, from: number, to: number, color: string, style: "thin" | "medium" | "thick" }} BgLine
 */

/* thinner than this (px on the dashboard) a coloured zone is a line, not an area */
const LINE_PX = 8;

/* zone coordinate (0–100000) → grid edge: piecewise linear through the edges of the blocks laid out from
 * zones, never decreasing (the layout keeps the dashboard's order) */
function axisMap(anchors, lo, hi) {
  const pts = [[0, lo], ...anchors, [100000, hi]].sort((a, b) => a[0] - b[0] || a[1] - b[1]);
  for (let i = 1; i < pts.length; i++) pts[i][1] = Math.max(pts[i][1], pts[i - 1][1]);
  return v => {
    if (v <= pts[0][0]) return pts[0][1];
    for (let i = 1; i < pts.length; i++) {
      if (v > pts[i][0]) continue;
      const [x0, g0] = pts[i - 1], [x1, g1] = pts[i];
      return x1 === x0 ? g1 : g0 + (g1 - g0) * (v - x0) / (x1 - x0);
    }
    return pts[pts.length - 1][1];
  };
}

/**
 * The colour each cell shows behind the blocks, as on the dashboard: the dashboard shading at the back,
 * then every container / backdrop zone with a background from the largest to the smallest, then each
 * block's own colour. A zone colours the blocks whose zones sit inside it (their centres) and the cells
 * between them – out to the zone's own edges, placed on the grid between the blocks' edges, so a header
 * bar spans the sheet as it spans the dashboard. Zones thinner than a few px are dividers: lines.
 * @param {DashboardZone[]} zones @param {string | null} dashColor ARGB
 * @param {BgBlock[]} blocks
 * @param {{ top: number, left: number, bottom: number, right: number }} canvas the dashboard's cells
 * @param {{ w: number, h: number }} dashPx the dashboard's size in px
 * @returns {{ colorAt: (row: number, col: number) => string | null, lines: BgLine[] }} colorAt: ARGB, null = leave the cell
 */
function backgroundPlan(zones, dashColor, blocks, canvas, dashPx) {
  const zoneById = new Map((zones || []).map(z => [String(z.id), z]));
  const blockZones = new Set(blocks.map(b => b.zoneId));
  const placed = blocks.map(b => ({ b, z: zoneById.get(b.zoneId) })).filter(p => p.z && p.z.w > 0 && p.z.h > 0);
  const mapX = axisMap(placed.flatMap(({ b, z }) => [[z.x, b.left], [z.x + z.w, b.right + 1]]), canvas.left, canvas.right + 1);
  const mapY = axisMap(placed.flatMap(({ b, z }) => [[z.y, b.top], [z.y + z.h, b.bottom + 1]]), canvas.top, canvas.bottom + 1);
  const center = z => ({ x: z.x + z.w / 2, y: z.y + z.h / 2 });
  const inside = (p, z) => p.x >= z.x && p.x <= z.x + z.w && p.y >= z.y && p.y <= z.y + z.h;
  // a zone in a hidden container is not on the dashboard either
  const shown = z => { for (let q = z, n = 0; q && n < 50; q = zoneById.get(String(q.parent)), n++) if (q.hidden) return false; return true; };
  const coloured = (zones || []).map((z, order) => ({ z, order }))
    .filter(({ z }) => z.style && z.style.bgColor && !blockZones.has(String(z.id)) && z.w > 0 && z.h > 0 && shown(z));
  const px = z => ({ w: z.w * dashPx.w / 100000, h: z.h * dashPx.h / 100000 });

  /** @type {BgLine[]} */
  const lines = [];
  coloured.forEach(({ z }) => {
    const { w, h } = px(z);
    if (Math.min(w, h) >= LINE_PX || Math.max(w, h) < LINE_PX) return;
    const thick = Math.min(w, h), style = thick < 1.5 ? "thin" : thick < 2.5 ? "medium" : "thick";
    const c = center(z);
    const line = h < w
      ? { dir: "h", at: Math.round(mapY(c.y)), from: Math.round(mapX(z.x)), to: Math.round(mapX(z.x + z.w)) - 1 }
      : { dir: "v", at: Math.round(mapX(c.x)), from: Math.round(mapY(z.y)), to: Math.round(mapY(z.y + z.h)) - 1 };
    if (line.to >= line.from) lines.push(/** @type {BgLine} */ ({ ...line, color: z.style.bgColor, style }));
  });

  const layers = coloured
    .filter(({ z }) => { const { w, h } = px(z); return Math.min(w, h) >= LINE_PX; })
    .map(({ z, order }) => {
      const members = placed.filter(p => inside(center(p.z), z)).map(p => p.b);
      const rect = { top: Math.round(mapY(z.y)), left: Math.round(mapX(z.x)),
                     bottom: Math.round(mapY(z.y + z.h)) - 1, right: Math.round(mapX(z.x + z.w)) - 1 };
      members.forEach(b => {
        rect.top = Math.min(rect.top, b.top); rect.left = Math.min(rect.left, b.left);
        rect.bottom = Math.max(rect.bottom, b.bottom); rect.right = Math.max(rect.right, b.right);
      });
      if (rect.bottom < rect.top || rect.right < rect.left) return null;
      return { color: z.style.bgColor, area: z.w * z.h, order, members: new Set(members), ...rect };
    })
    .filter(Boolean)
    .sort((a, b) => b.area - a.area || a.order - b.order);             // back to front
  const covers = (b, r, c) => r >= b.top && r <= b.bottom && c >= b.left && c <= b.right;
  const colorAt = (r, c) => {
    const block = blocks.find(b => covers(b, r, c));
    if (block) {
      if (block.own === "white") return null;
      if (block.own) return block.own;
      for (let i = layers.length - 1; i >= 0; i--) if (layers[i].members.has(block)) return layers[i].color;
      return dashColor || null;
    }
    for (let i = layers.length - 1; i >= 0; i--) if (covers(layers[i], r, c)) return layers[i].color;
    return dashColor || null;
  };
  return { colorAt, lines };
}


/* ═══ visual/workbook-match.js ════════════════════════════════════════════════════════════════════ */
/* Does the loaded workbook belong to the running dashboard? */

/* does the loaded workbook belong to the running dashboard? (a different or stale .twb would
 * silently apply the wrong titles, colours and sorts) → { level: "ok" | "partial" | "mismatch", … } */
function checkWorkbookMatch(model, dashboard) {
  if (!model || !dashboard) return null;
  const names = (dashboard.worksheets || []).map(w => w.name);
  const known = new Set(Object.keys(model.sheets || {}).map(tfNorm));
  const missing = names.filter(n => !known.has(tfNorm(n)));
  const matched = names.length - missing.length;
  const dashboardFound = Object.keys(model.dashboards || {}).some(d => tfNorm(d) === tfNorm(dashboard.name));
  const level = !missing.length && dashboardFound ? "ok"
    : names.length && matched / names.length >= 0.5 ? "partial" : "mismatch";
  return { level, dashboard: dashboard.name, dashboardFound, matched, total: names.length, missing };
}

function describeWorkbookMatch(match) {
  if (!match || match.level === "ok") return "";
  const parts = [];
  if (!match.dashboardFound) parts.push(`dashboard "${match.dashboard}" is not in this workbook`);
  if (match.missing.length) {
    parts.push(`${match.matched}/${match.total} worksheets found` +
      ` (missing: ${match.missing.slice(0, 3).join(", ")}${match.missing.length > 3 ? ", …" : ""})`);
  }
  return (match.level === "mismatch" ? "⚠ Wrong workbook? " : "⚠ Workbook partly matches: ") + parts.join("; ");
}


/* ═══ ui/status.js ════════════════════════════════════════════════════════════════════════════════ */
/* Panel status line and loaded-workbook label. */

/* loaded-workbook label, with a warning when the file does not match this dashboard */
function showWorkbookLabel(fileName, details, model) {
  const label = document.getElementById("twb_file_label");
  if (!label) return;
  const match = checkWorkbookMatch(model, tableau.extensions.dashboardContent.dashboard);
  const warning = describeWorkbookMatch(match);
  label.textContent = `${warning ? "" : "✅ "}${fileName} — ${details}` + (warning ? `\n${warning}` : "");
  label.title = match && match.missing.length ? "Not found in the workbook:\n" + match.missing.join("\n") : "";
  label.classList.toggle("status-warning", !!warning);
  if (warning) console.warn("[Workbook]", warning, match);
}

/* progress / outcome line under the Export button */
function setExportStatus(text, warning) {
  const target = document.getElementById("export_status");
  if (!target) return;
  target.textContent = text;
  target.title = "";
  target.classList.toggle("status-warning", !!warning);
}

function appendExportStatus(text) {
  const target = document.getElementById("export_status");
  if (target && target.textContent) target.textContent += " · " + text;
}

function updateVisualStatus(statuses, note) {
  const target = document.getElementById("export_status");
  if (!target) return;
  if (note) {                                      // e.g. the loaded workbook does not match
    updateVisualStatus(statuses);
    target.textContent = note + " · " + target.textContent;
    target.classList.add("status-warning");
    return;
  }
  if (!statuses || !statuses.length) {
    target.textContent = "No worksheet visual status available";
    return;
  }
  // once converted: how (NATIVE … TABLE_FALLBACK, as on the report sheet); before that, which renderer
  if (statuses.every(item => item.strategy)) {
    const order = ["NATIVE", "CONSTRUCTED", "APPROXIMATE", "TABLE_FALLBACK", "IMAGE_FALLBACK", "UNSUPPORTED", "FAILED"];
    const words = { NATIVE: "native", CONSTRUCTED: "constructed", APPROXIMATE: "approximated", TABLE_FALLBACK: "as data tables",
                    IMAGE_FALLBACK: "as pictures", UNSUPPORTED: "unsupported", FAILED: "failed" };
    const tables = statuses.filter(item => item.asTable).length;          // "Charts as tables": chosen, not a fallback
    const parts = order.map(s => [statuses.filter(item => item.strategy === s && !item.asTable).length, words[s]]).filter(([n]) => n);
    if (tables) parts.push([tables, "as tables (Charts as tables)"]);
    const flagged = statuses.filter(item => item.strategy !== "NATIVE" && item.strategy !== "CONSTRUCTED");
    target.textContent = `${statuses.length} visuals: ${parts.map(([n, w]) => `${n} ${w}`).join(", ")} – see the Conversion Report sheet`;
    target.title = flagged.map(w => `${w.worksheet} (${w.visual}): ${w.strategy} – ${w.reason}`).join("\n");
    target.classList.toggle("status-warning", statuses.some(item => item.strategy === "FAILED"));
    return;
  }
  const count = renderer => statuses.filter(item => item.renderer === renderer).length;
  const parts = [
    [count("cell"), "tables/KPIs"],
    [count("excel-chart"), "Excel charts"],
    [count("tableau-image"), "Tableau images"],
    [count("data-fallback"), "as data"]
  ].filter(([n]) => n).map(([n, label]) => `${n} ${label}`);
  const warnings = statuses.filter(item => item.status !== "success" && item.status !== "pending");
  target.textContent = `${statuses.length} visuals: ${parts.join(", ")}` +
    (warnings.length ? ` – ${warnings.length} warning${warnings.length === 1 ? "" : "s"} (see console)` : "");
  target.title = warnings.map(w => `${w.worksheet}: ${w.reason}`).join("\n");
  target.classList.toggle("status-warning", warnings.length > 0);
}

/* plain message on the loaded-workbook label (auto-load progress, problems) */
function showWorkbookStatus(text, warning) {
  const label = document.getElementById("twb_file_label");
  if (!label) return;
  label.textContent = text;
  label.title = "";
  label.classList.toggle("status-warning", !!warning);
}

/* what was left out of a multi-dashboard export, and why – one line each under the status */
function appendExportNotes(notes) {
  (notes || []).forEach(n => console.warn("[Export] " + n));
  const target = document.getElementById("export_status");
  if (!target || !notes || !notes.length) return;
  target.textContent = [target.textContent, ...notes.map(n => "⚠️ " + n)].filter(Boolean).join("\n");
  target.classList.add("status-warning");
}


/* ═══ cloud/backend.js ════════════════════════════════════════════════════════════════════════════ */
/* =============================================================================
 * AUTO-FETCH FROM TABLEAU CLOUD (Phase 1)
 * -----------------------------------------------------------------------------
 * Replaces the manual "Load Workbook" step on Tableau Cloud:
 *   1. ask the backend (Worker) which published workbook holds this dashboard
 *   2. download that workbook's XML through the backend
 *   3. parse it exactly like a manually uploaded file (applyWorkbookXml)
 * The manual 📁 button stays as a fallback (Tableau Desktop, backend down).
 * The PAT never reaches this file — only the backend holds it.
 * ============================================================================= */

const DEFAULT_BACKEND_URL = "";      // production: the hosted Worker URL, e.g. "https://export-to-excel-backend.<account>.workers.dev"
const EXTENSION_VERSION = "2.0.2";
/** which build made an export: shown in the Conversion Report and the panel (old cached files are easy to spot) */
const EXTENSION_BUILD = "2026-10-30 multi-gauge";
console.info(`[Export to Excel] v${EXTENSION_VERSION} · build ${EXTENSION_BUILD}`);
/** Links shown with the licence status (leave "" to hide). */
const PRODUCT_LINKS = { buy: "", help: "" };

/** The "Advanced: backend connection" box is for development: hidden once DEFAULT_BACKEND_URL is set,
 *  unless this browser opts in with localStorage tfx_showAdvanced = "1". */
function advancedBackendAllowed() {
  if (!DEFAULT_BACKEND_URL) return true;
  try { return localStorage.getItem("tfx_showAdvanced") === "1"; } catch (e) { return false; }
}

/** Read a value saved by this extension: workbook settings first, then this browser. */
function readSaved(key) {
  let v = null;
  try { v = tableau.extensions.settings.get(key); } catch (e) { /* ignore */ }
  if (!v) { try { v = localStorage.getItem("tfx_" + key); } catch (e) { /* ignore */ } }
  return v || "";
}

/** Save values to workbook settings (works in edit mode) and this browser (works always). */
async function writeSaved(pairs) {
  Object.keys(pairs).forEach(k => {
    const v = pairs[k];
    try { localStorage.setItem("tfx_" + k, v); } catch (e) { /* ignore */ }
    try {
      if (v === "") tableau.extensions.settings.erase(k);
      else tableau.extensions.settings.set(k, v);
    } catch (e) { /* ignore */ }
  });
  try { await tableau.extensions.settings.saveAsync(); }
  catch (e) { console.warn("[AutoFetch] settings.saveAsync failed (viewing mode?) — kept in this browser only:", e.message); }
}

function getBackendConfig() {
  // with a production URL built in, an old saved (tunnel) address must not override it
  const saved = advancedBackendAllowed() ? readSaved("backendUrl") : "";
  return {
    url: (saved || DEFAULT_BACKEND_URL).trim().replace(/\/+$/, ""),
    key: readSaved("backendKey").trim(),
  };
}

/* -----------------------------------------------------------------------------
 * Licence (marketplace backend). The workbook author saves a licence key once (it is stored with the
 * workbook); every backend call sends it, with Tableau's anonymous uniqueUserId for seat counting.
 * A single-client backend has no /licence route: LICENCE then stays null and nothing here shows.
 * --------------------------------------------------------------------------- */
/** @type {null | { valid: boolean, reason?: string|null, message?: string, plan?: string, tenant?: string,
 *   daysLeft?: number, expiresAt?: string, seats?: number, seatsUsed?: number }} */
let LICENCE = null;

class LicenceError extends Error {
  constructor(info) { super(info.message || "Licence problem"); this.licence = info; }
}
function tenantKey() { return readSaved("tenantKey").trim(); }
function userIdForSeats() {
  try { return String((tableau.extensions.environment || {}).uniqueUserId || ""); } catch (e) { return ""; }
}
/** Lite mode: a marketplace backend that says the licence is missing, invalid or over. */
function liteMode() { return !!(LICENCE && LICENCE.valid === false); }

const LICENCE_REASONS = {
  missing_key: "no licence key yet",
  invalid_key: "the licence key is not valid",
  trial_expired: "the trial has ended",
  expired: "the licence has expired",
  canceled: "the subscription was cancelled",
  suspended: "the licence is suspended",
  seat_limit: "all licensed users for this month are in use",
};

/** Ask the backend for the licence. Sets LICENCE (null when the backend has no licensing). */
async function checkLicence() {
  if (!getBackendConfig().url) { LICENCE = null; return LICENCE; }
  try {
    const res = await backendFetch("/licence");
    LICENCE = await res.json();
  } catch (e) {
    if (e instanceof LicenceError) LICENCE = e.licence;                 // missing / invalid key, ended licence
    else if (/^Backend 404/.test(e.message)) LICENCE = null;           // single-client backend: no licensing
    else { console.warn("[Licence] check failed:", e.message); LICENCE = null; }   // backend down: export as before
  }
  refreshLicenceBox();
  applyLiteMode();
  return LICENCE;
}

/** Lite mode turns off "All dashboards" (it needs the backend); this dashboard still exports. */
function applyLiteMode() {
  const allBox = /** @type {HTMLInputElement | null} */ (document.getElementById("export_all"));
  if (!allBox) return;
  if (liteMode()) {
    allBox.checked = false;
    allBox.disabled = true;
    allBox.title = "Needs a valid licence";
  } else if (allBox.title === "Needs a valid licence") {
    allBox.disabled = false;
    allBox.title = "";
  }
}

function licenceLine() {
  const L = LICENCE;
  if (!L) return "";
  const users = Number.isFinite(L.seatsUsed) && L.seats ? ` · ${L.seatsUsed}/${L.seats} users this month` : "";
  if (L.valid) {
    const until = L.expiresAt ? new Date(L.expiresAt).toLocaleDateString() : "";
    if (L.plan === "trial") return `🕒 Trial${L.tenant ? " — " + L.tenant : ""} · ${L.daysLeft} day${L.daysLeft === 1 ? "" : "s"} left${users}`;
    return `✅ Pro licence${L.tenant ? " — " + L.tenant : ""}${until ? " · valid until " + until : ""}${users}`;
  }
  const why = LICENCE_REASONS[L.reason] || L.message || "licence problem";
  return `⚠️ Lite mode — ${why}. This dashboard still exports; Tableau formatting from the cloud and "All dashboards" need a licence.`;
}

/** The licence panel: a status line for everyone, and the key box for the workbook author. */
function ensureLicenceBox() {
  let box = document.getElementById("licence_box");
  if (box) return box;
  const exportBtn = document.getElementById("export_button");
  if (!exportBtn || !exportBtn.parentNode) return null;
  box = document.createElement("div");
  box.id = "licence_box";
  box.style.cssText = "margin:6px 0 8px;font-size:11px;line-height:1.4;";
  box.innerHTML = `<div id="licence_line" style="color:#3c4a44"></div>
    <div id="licence_links" style="margin-top:2px"></div>
    <div id="licence_setup" style="display:none;margin-top:6px">
      <label for="licence_key" style="display:block;font-weight:600;margin-bottom:2px">Licence key</label>
      <div style="display:flex;gap:6px">
        <input id="licence_key" type="password" autocomplete="off" spellcheck="false" placeholder="eox_…"
               style="flex:1;min-width:0;font-size:11px;padding:4px 6px">
        <button id="licence_save" class="btn-load" style="width:auto;padding:4px 10px">Save</button>
      </div>
      <div style="color:#6b7a73;margin-top:2px">Saved with this workbook, so viewers don't need it.</div>
    </div>
    <div style="color:#9aa6a0;margin-top:2px">v${EXTENSION_VERSION} · ${EXTENSION_BUILD}</div>`;
  exportBtn.parentNode.insertBefore(box, exportBtn);
  return box;
}
function refreshLicenceBox() {
  const box = ensureLicenceBox();
  if (!box) return;
  const line = /** @type {HTMLElement} */ (box.querySelector("#licence_line"));
  line.textContent = licenceLine();
  const links = /** @type {HTMLElement} */ (box.querySelector("#licence_links"));
  links.innerHTML = "";
  const addLink = (href, text) => {
    if (!href) return;
    const a = document.createElement("a");
    a.href = href; a.target = "_blank"; a.rel = "noopener"; a.textContent = text; a.style.marginRight = "10px";
    links.appendChild(a);
  };
  if (LICENCE && (!LICENCE.valid || LICENCE.plan === "trial")) addLink(PRODUCT_LINKS.buy, "Buy or renew a licence");
  if (LICENCE) addLink(PRODUCT_LINKS.help, "Help");
  // the key box: only where the workbook can be saved (authoring), and only with a marketplace backend
  let authoring = false;
  try { authoring = (tableau.extensions.environment || {}).mode === "authoring"; } catch (e) { /* ignore */ }
  const setup = /** @type {HTMLElement} */ (box.querySelector("#licence_setup"));
  setup.style.display = authoring && LICENCE ? "block" : "none";
}

/** Wire the key box; saving a key re-checks the licence and reloads the workbook. */
function setupLicence(dashboard, onReload) {
  const box = ensureLicenceBox();
  if (!box) return;
  const input = /** @type {HTMLInputElement} */ (box.querySelector("#licence_key"));
  const save = /** @type {HTMLButtonElement} */ (box.querySelector("#licence_save"));
  input.value = tenantKey();
  save.addEventListener("click", async () => {
    save.disabled = true;
    await writeSaved({ tenantKey: input.value.trim() });
    await checkLicence();
    const start = () => autoLoadWorkbook(dashboard);
    await (onReload ? onReload(start) : start());
    save.disabled = false;
  });
}
/** The panel's "Charts as tables" switch, under "All dashboards". Remembered per workbook (and browser). */
function setupVisualsAsTables() {
  if (document.getElementById("visuals_as_tables")) return;
  const allBox = document.getElementById("export_all");
  const anchor = allBox ? (allBox.closest("label") || allBox.parentElement) : document.getElementById("export_button");
  if (!anchor || !anchor.parentNode) return;
  const row = document.createElement("label");
  row.id = "visuals_as_tables_row";
  row.style.cssText = "display:flex;align-items:center;gap:6px;margin:4px 0 6px;font-size:12px;cursor:pointer;";
  row.title = "Off: charts are exported as Excel charts. On: each chart is exported as a table of its data, " +
              "coloured with the chart's Tableau colours.";
  row.innerHTML = `<input type="checkbox" id="visuals_as_tables"><span>Charts as tables (Tableau colours)</span>`;
  if (allBox) anchor.parentNode.insertBefore(row, anchor.nextSibling);
  else anchor.parentNode.insertBefore(row, anchor);
  const box = /** @type {HTMLInputElement} */ (row.querySelector("input"));
  box.checked = visualsAsTables();
  box.addEventListener("change", () => { writeSaved({ visualsAs: box.checked ? "tables" : "charts" }); });
}

/** Tableau's "Configure…" menu: show the licence key box. */
function openLicenceSetup() {
  const box = ensureLicenceBox();
  if (!box) return;
  /** @type {HTMLElement} */ (box.querySelector("#licence_setup")).style.display = "block";
  try { if (typeof box.scrollIntoView === "function") box.scrollIntoView({ block: "center" }); } catch (e) { /* ignore */ }
  /** @type {HTMLElement} */ (box.querySelector("#licence_key")).focus();
}

/** Call the backend. Throws a readable Error on network failure or non-2xx. */
async function backendFetch(path, init = {}) {
  const { url, key } = getBackendConfig();
  if (!url) throw new Error("No backend URL configured");
  const headers = { ...(init.headers || {}) };
  if (key) headers["X-Spike-Key"] = key;
  // licence headers only with a licence key (marketplace backend): a single-client backend's CORS allows just
  // Content-Type and X-Spike-Key – any other header makes the browser block the request ("could not reach")
  const tk = tenantKey();
  if (tk) {
    headers["X-Tenant-Key"] = tk;
    const uid = userIdForSeats();
    if (uid) headers["X-User-Id"] = uid;
  }
  // a backend that never answers (tunnel down, laptop asleep) must not leave the workbook "loading" for ever
  const ctrl = typeof AbortController === "function" ? new AbortController() : null;
  const timer = ctrl ? setTimeout(() => ctrl.abort(), BACKEND_TIMEOUT_MS) : null;
  let res;
  try { res = await fetch(url + path, { ...init, headers, ...(ctrl ? { signal: ctrl.signal } : {}) }); }
  catch (e) {
    if (ctrl && ctrl.signal.aborted) throw new Error(`The backend (${url}) did not answer within ${Math.round(BACKEND_TIMEOUT_MS / 1000)}s — is it running?`);
    throw new Error(`Could not reach the backend (${url}) — is it running?`);
  }
  finally { if (timer) clearTimeout(timer); }
  if (!res.ok) {
    let text = "";
    try { text = await res.text(); } catch (e) { /* ignore */ }
    let msg = text, body = null;
    try { body = JSON.parse(text); msg = body.error || text; } catch (e) { /* not JSON */ }
    if (body && body.error === "licence") {
      // the licence ran out (or no key): switch to Lite mode and say why
      const info = { valid: false, reason: body.reason, message: body.message, ...(body.licence || {}) };
      info.valid = false;
      if (path !== "/licence") { LICENCE = info; refreshLicenceBox(); applyLiteMode(); }
      throw new LicenceError(info);
    }
    throw new Error(`Backend ${res.status}: ${String(msg).slice(0, 300)}`);
  }
  return res;
}

/** "N titles, formatting for M sheets (K with colour)" – the loaded-workbook label */
function describeModel(titleMap, formatModel, images) {
  const sheets = Object.values((formatModel && formatModel.sheets) || {});
  const colour = sheets.filter(s => (s.panes || []).some(p => (p.encodings || []).some(e => e.channel === "color"))).length;
  const imageCount = images ? Object.keys(images).length : 0;
  return `${Object.keys(titleMap || {}).length} titles, formatting for ${sheets.length} sheets (${colour} with colour)` +
    (imageCount ? `, ${imageCount} image${imageCount > 1 ? "s" : ""}` : "");
}

/**
 * Safety check: does the downloaded XML really contain this dashboard and its worksheets?
 * Catches a stale saved workbook ID (e.g. the workbook was copied with "Save As").
 */
function twbMatchesDashboard(xmlString, dashboard) {
  const doc = new DOMParser().parseFromString(xmlString, "text/xml");
  const root = doc && doc.documentElement;
  if (!root || root.nodeName !== "workbook") return false;
  const names = tag => {
    const holder = getDirectChildByTag(root, tag + "s");
    const out = new Set();
    if (!holder) return out;
    for (let i = 0; i < holder.childNodes.length; i++) {
      const c = holder.childNodes[i];
      if (c.nodeType === 1 && c.tagName === tag) out.add(c.getAttribute("name"));
    }
    return out;
  };
  const dashboards = names("dashboard"), worksheets = names("worksheet");
  return dashboards.has(dashboard.name) && dashboard.worksheets.every(ws => worksheets.has(ws.name));
}

/** Which published workbook is this? Saved answer first, otherwise ask the backend. */
async function resolveWorkbook(dashboard) {
  const savedId = readSaved("twbWorkbookId");
  if (savedId) return { id: savedId, name: readSaved("twbWorkbookName") || "workbook", fromSaved: true };

  const res = await backendFetch("/resolve", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ dashboardName: dashboard.name, worksheetNames: dashboard.worksheets.map(w => w.name) }),
  });
  const r = await res.json();
  const pick = c => ({ id: c.workbookId, name: c.workbookName, project: c.projectName });
  if (r.decision === "unique") return pick(r.candidates[0]);
  if (r.decision === "ambiguous") {
    const exact = r.candidates.filter(c => c.exact);
    const list = (exact.length ? exact : r.candidates.filter(c => c.score > 0)).map(pick);
    return { ambiguous: true, candidates: list };
  }
  throw new Error(`no published workbook contains dashboard "${dashboard.name}"`);
}

/** the open "pick a workbook" prompt – a newer load or a manual 📁 load closes it (its promise gets null) */
/** @type {{ box: HTMLElement, resolve: (v: any) => void } | null} */
let PENDING_PICKER = null;
function cancelWorkbookPicker() {
  if (!PENDING_PICKER) return;
  const p = PENDING_PICKER;
  PENDING_PICKER = null;
  p.box.remove();
  p.resolve(null);
}

/** Several workbooks match (e.g. copies in two projects): let the user pick once. Resolves null if cancelled. */
function askUserToChooseWorkbook(candidates) {
  cancelWorkbookPicker();
  return new Promise(resolve => {
    const label = document.getElementById("twb_file_label");
    const box = document.createElement("div");
    box.id = "twb_picker";
    box.className = "twb-picker";
    const select = document.createElement("select");
    candidates.forEach((c, i) => {
      const o = document.createElement("option");
      o.value = String(i);
      o.textContent = `${c.project ? c.project + " / " : ""}${c.name}`;
      select.appendChild(o);
    });
    const ok = document.createElement("button");
    ok.className = "btn-load";
    ok.textContent = "Use this workbook";
    ok.addEventListener("click", () => {
      PENDING_PICKER = null;
      box.remove();
      resolve(candidates[Number(select.value)]);
    });
    box.appendChild(select);
    box.appendChild(ok);
    if (label && label.parentNode) label.parentNode.insertBefore(box, label.nextSibling);
    else document.body.appendChild(box);
    PENDING_PICKER = { box, resolve };
  });
}

/* -----------------------------------------------------------------------------
 * Export button state. Export is clickable only once the workbook is ready (loaded, or its loading has
 * ended – failed or not set up – so the export can still run); while it loads, or while the user still has
 * to pick between workbooks, the button is disabled and says why. A running export is never started twice.
 * --------------------------------------------------------------------------- */
/** @type {"loading" | "choosing" | "ready"} */
let WORKBOOK_PHASE = "ready";
let EXPORT_RUNNING = false;
/** @type {string | null} the button's own label ("Export to EXCEL") */
let EXPORT_LABEL = null;
/** every new workbook load gets a number: an older one that ends late must not undo a newer state */
let WORKBOOK_LOAD_GEN = 0;

function refreshExportButton() {
  const b = /** @type {HTMLButtonElement | null} */ (document.getElementById("export_button"));
  if (!b || EXPORT_RUNNING) return;
  if (EXPORT_LABEL === null) EXPORT_LABEL = b.textContent;
  if (WORKBOOK_PHASE === "loading") { b.disabled = true; b.textContent = "⏳ Loading workbook…"; }
  else if (WORKBOOK_PHASE === "choosing") { b.disabled = true; b.textContent = "⬆ Pick the workbook above, then click \u201cUse this workbook\u201d"; }
  else { b.disabled = false; b.textContent = EXPORT_LABEL; }
}
/** @param {"loading" | "choosing" | "ready"} phase */
function setWorkbookPhase(phase) { WORKBOOK_PHASE = phase; refreshExportButton(); }

/**
 * Track a workbook load: Export waits (disabled) until it ends. A load still running after
 * WORKBOOK_WATCHDOG_MS releases Export anyway, with a note – never a button that stays grey for ever.
 * @param {() => Promise<any>} start begins the load – called after this load's number is taken, so the
 *   load sees itself as the newest one @returns {Promise<any>}
 */
function trackWorkbookLoad(start) {
  const mine = ++WORKBOOK_LOAD_GEN;
  setWorkbookPhase("loading");
  let promise;
  try { promise = start(); } catch (e) { promise = Promise.reject(e); }
  const watchdog = setTimeout(() => {
    if (mine === WORKBOOK_LOAD_GEN && WORKBOOK_PHASE === "loading") {
      showWorkbookStatus("⚠️ The workbook is taking long to load — you can export now, but Tableau formatting may be missing", true);
      setWorkbookPhase("ready");
    }
  }, WORKBOOK_WATCHDOG_MS);
  return Promise.resolve(promise).catch(() => {}).then(() => {
    clearTimeout(watchdog);
    if (mine === WORKBOOK_LOAD_GEN) setWorkbookPhase("ready");
  });
}

/** Auto-load on open. Never throws: on any problem it explains and leaves 📁 as the fallback. */
async function autoLoadWorkbook(dashboard) {
  const env = tableau.extensions.environment || {};
  const savedFile = readSaved("twbFileName");
  const fallback = savedFile ? ` — using previously loaded ${savedFile}` : " — use 📁 to load the workbook manually";

  if (env.context === "desktop") {
    if (!savedFile) showWorkbookStatus("Tableau Desktop: auto-load needs Tableau Cloud — use 📁 to load the workbook", true);
    return;
  }
  if (!getBackendConfig().url) {
    if (!savedFile) showWorkbookStatus("Auto-load not set up (Advanced → Backend URL)" + fallback, true);
    return;
  }
  await checkLicence();
  if (liteMode()) {
    showWorkbookStatus(`Lite mode (${LICENCE_REASONS[LICENCE.reason] || "licence problem"})` +
      (savedFile ? ` — using previously loaded ${savedFile}` : " — use 📁 to load the workbook for Tableau formatting"), true);
    return;
  }

  const gen = WORKBOOK_LOAD_GEN;                        // a newer load (or a manual 📁 load) supersedes this one
  try {
    for (let attempt = 0; attempt < 2; attempt++) {
      showWorkbookStatus("⏳ Finding this workbook on Tableau Cloud…");
      let wb = await resolveWorkbook(dashboard);
      if (gen !== WORKBOOK_LOAD_GEN) return;
      if (wb.ambiguous) {
        showWorkbookStatus(`${wb.candidates.length} published workbooks contain this dashboard — pick the right one, then click \u201cUse this workbook\u201d:`);
        setWorkbookPhase("choosing");
        wb = await askUserToChooseWorkbook(wb.candidates);
        if (!wb || gen !== WORKBOOK_LOAD_GEN) return;        // closed by a newer load / a manual load
        setWorkbookPhase("loading");
      }

      showWorkbookStatus(`⏳ Downloading "${wb.name}" from Tableau Cloud…`);
      let xml;
      try {
        xml = await (await backendFetch("/twb/" + encodeURIComponent(wb.id))).text();
      } catch (e) {
        if (wb.fromSaved && /Backend 404/.test(e.message)) {          // saved workbook was deleted/moved
          await writeSaved({ twbWorkbookId: "", twbWorkbookName: "" });
          continue;
        }
        throw e;
      }

      if (!twbMatchesDashboard(xml, dashboard)) {
        if (wb.fromSaved) {                                            // saved ID points at another workbook
          await writeSaved({ twbWorkbookId: "", twbWorkbookName: "" });
          continue;
        }
        throw new Error("the downloaded workbook doesn't contain this dashboard");
      }

      if (gen !== WORKBOOK_LOAD_GEN) return;                // superseded while downloading: keep the newer workbook
      const { titleMap, formatModel, images } = await applyWorkbookXml(xml, wb.name, false);
      setLoadedWorkbookId(wb.id);            // lets Phase 3 find the other dashboards' views
      if (!wb.fromSaved) await writeSaved({ twbWorkbookId: wb.id, twbWorkbookName: wb.name });
      showWorkbookLabel(`${wb.name} (auto, Tableau Cloud)`, describeModel(titleMap, formatModel, images), formatModel);
      return;
    }
    throw new Error("could not identify this workbook");
  } catch (e) {
    if (gen !== WORKBOOK_LOAD_GEN) return;
    console.warn("[AutoFetch]", e);
    showWorkbookStatus(`⚠️ Auto-load failed: ${e.message}${fallback}`, true);
  }
}

/** Advanced panel: backend URL + key, saved with the workbook. Saving reconnects.
 *  onReload(start) runs the new auto-load (start) so Export waits for it, and returns its promise. */
function setupBackendSettings(dashboard, onReload) {
  const urlIn = /** @type {HTMLInputElement} */ (document.getElementById("backend_url"));
  const keyIn = /** @type {HTMLInputElement} */ (document.getElementById("backend_key"));
  const saveBtn = /** @type {HTMLButtonElement} */ (document.getElementById("backend_save"));
  if (!urlIn || !keyIn || !saveBtn) return;
  if (!advancedBackendAllowed()) {                       // production build: the backend address is built in
    const holder = urlIn.closest("details") || urlIn.parentElement;
    if (holder) /** @type {HTMLElement} */ (holder).style.display = "none";
    return;
  }
  const cfg = getBackendConfig();
  urlIn.value = readSaved("backendUrl");
  urlIn.placeholder = DEFAULT_BACKEND_URL || "https://…trycloudflare.com";
  keyIn.value = cfg.key;
  saveBtn.addEventListener("click", async () => {
    saveBtn.disabled = true;
    await writeSaved({ backendUrl: urlIn.value.trim(), backendKey: keyIn.value.trim() });
    const start = () => autoLoadWorkbook(dashboard);
    await (onReload ? onReload(start) : start());
    saveBtn.disabled = false;
  });
}


/* ═══ ui/workbook-store.js ════════════════════════════════════════════════════════════════════════ */
/* Loading the .twb/.twbx, remembering it (settings + IndexedDB) and the save location. */

/* =============================================================================
 * parseTwbXmlInBrowser(xmlString) - Parses TITLES from XML
 * ============================================================================= */
function parseTwbXmlInBrowser(xmlString) {
  const doc = new DOMParser().parseFromString(xmlString, "text/xml");
  const titleMap = {};

  const worksheetNodes = doc.getElementsByTagName("worksheet");

  for (let i = 0; i < worksheetNodes.length; i++) {
    const wsNode = worksheetNodes[i];
    const internalName = wsNode.getAttribute("name");
    if (!internalName) continue;

    const layoutNode = getDirectChildByTag(wsNode, "layout-options")
                    || getDirectChildByTag(wsNode, "layout");
    if (!layoutNode) continue;

    const titleNode = getDirectChildByTag(layoutNode, "title");
    if (!titleNode) continue;

    const fmtNode = titleNode.getElementsByTagName("formatted-text")[0];
    if (!fmtNode) continue;

    const runNodes = fmtNode.getElementsByTagName("run");
    const displayTitle = Array.from(runNodes)
      .map(r => r.textContent || "")
      .join("")
      .trim();

    if (displayTitle) {
      titleMap[internalName] = displayTitle;
      console.log(`[TWB parse] Title: "${internalName}" → "${displayTitle}"`);
    }
  }

  return titleMap;
}

function getDirectChildByTag(parent, tagName) {
  for (let i = 0; i < parent.childNodes.length; i++) {
    const child = parent.childNodes[i];
    if (child.nodeType === 1 && child.tagName === tagName) return child;
  }
  return null;
}

/* =============================================================================
 * Workbook file handle – the save dialog opens in the loaded workbook's folder.
 * Pages cannot read a file's path, but a FileSystemFileHandle from showOpenFilePicker
 * can be passed to showSaveFilePicker({ startIn }). It is kept in IndexedDB so the
 * next Tableau session still starts there. Runtimes without the File System Access
 * API keep the plain <input type="file"> + download behaviour.
 * ============================================================================= */
let workbookFileHandle = null;

const HANDLE_DB = "mark2table", HANDLE_STORE = "handles", HANDLE_KEY = "workbook";

function openHandleDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(HANDLE_DB, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(HANDLE_STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function storeWorkbookHandle(handle) {
  try {
    const db = await openHandleDb();
    db.transaction(HANDLE_STORE, "readwrite").objectStore(HANDLE_STORE).put(handle, HANDLE_KEY);
  } catch (e) {
    console.warn("[Workbook] could not remember the workbook folder:", e.message);
  }
}

/* parsed workbook model: Tableau's extension settings cannot hold a large model
 * (saveAsync fails) and it would be lost on the next reload → keep a copy here too, with the workbook XML
 * so a newer version of the extension can parse it again */
async function storeFormatModel(fileName, model, titleMap, xml) {
  try {
    const db = await openHandleDb();
    // one entry per workbook file: any dashboard of a workbook loaded once is matched again later
    db.transaction(HANDLE_STORE, "readwrite").objectStore(HANDLE_STORE)
      .put({ fileName, model, titleMap, xml, savedAt: Date.now() }, "model:" + String(fileName).toLowerCase());
  } catch (e) {
    console.warn("[Workbook] could not keep the workbook formatting in browser storage:", e.message);
  }
}

/** a model parsed by an older version of the extension: false = parsed by this one */
function isFormatModelStale(model) { return !!model && model.version !== FORMAT_MODEL_VERSION; }

/* a remembered workbook parsed by an older version: parsed again from its XML, and stored again */
function freshen(rec) {
  if (!isFormatModelStale(rec.model) || !rec.xml) return rec;
  try {
    const model = parseTableauFormatting(rec.xml);
    console.log(`[Workbook] ${rec.fileName}: formatting read again by this version of the extension`);
    storeFormatModel(rec.fileName, model, rec.titleMap, rec.xml);
    return { ...rec, model };
  } catch (e) {
    console.warn(`[Workbook] ${rec.fileName}: could not read the stored workbook again:`, e.message);
    return rec;
  }
}

async function restoreFormatModels() {
  try {
    const db = await openHandleDb();
    return await new Promise(resolve => {
      const out = [];
      const req = db.transaction(HANDLE_STORE).objectStore(HANDLE_STORE).openCursor();
      req.onsuccess = () => {
        const cur = req.result;
        if (!cur) return resolve(out);
        const key = String(cur.key);
        if ((key.startsWith("model:") || key === "formatModel") && cur.value && cur.value.model) out.push(freshen(cur.value));
        cur.continue();
      };
      req.onerror = () => resolve(out);
    });
  } catch (e) {
    return [];
  }
}

async function restoreWorkbookHandle() {
  try {
    const db = await openHandleDb();
    return await new Promise(resolve => {
      const req = db.transaction(HANDLE_STORE).objectStore(HANDLE_STORE).get(HANDLE_KEY);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => resolve(null);
    });
  } catch (e) {
    return null;
  }
}

/* =============================================================================
 * Packaged images – the files the dashboards' image objects show (logos, icons). Read from the .twbx
 * on load and kept in IndexedDB under their own key, so a remembered workbook still exports them.
 * ============================================================================= */
/** @type {{ file: string | null, images: Record<string, WorkbookImage> } | null} */
let IMAGE_CACHE = null;

/** "Image/logo.png" → its bytes, for every image zone of the workbook's dashboards found in the archive
 * @param {JSZip} zip @param {FormatModel} model @returns {Promise<Record<string, WorkbookImage>>} */
async function extractWorkbookImages(zip, model) {
  const wanted = new Set();
  Object.values(model.dashboards || {}).forEach(d => (d.zones || []).forEach(z => { if (z.type === "bitmap" && z.param) wanted.add(z.param); }));
  const files = Object.values(zip.files).filter(f => !f.dir);
  const norm = p => String(p).replace(/\\/g, "/").toLowerCase();
  const base = p => norm(p).split("/").pop();
  /** @type {Record<string, WorkbookImage>} */
  const images = {};
  for (const path of wanted) {
    // packaged as written in the zone ("Image/x.png"); a zone that points at the author's disk matches by file name
    const f = files.find(x => norm(x.name) === norm(path)) || files.find(x => base(x.name) === base(path));
    if (f) images[path] = { data: await f.async("uint8array") };
  }
  return images;
}

async function storeWorkbookImages(fileName, images) {
  try {
    const db = await openHandleDb();
    db.transaction(HANDLE_STORE, "readwrite").objectStore(HANDLE_STORE)
      .put({ fileName, images, savedAt: Date.now() }, "images:" + String(fileName).toLowerCase());
  } catch (e) {
    console.warn("[Workbook] could not keep the workbook images in browser storage:", e.message);
  }
}

/** the packaged images of the workbook the formatting comes from ({} for a .twb or none)
 * @returns {Promise<Record<string, WorkbookImage>>} */
async function getWorkbookImages() {
  if (IMAGE_CACHE && IMAGE_CACHE.file === FORMAT_MODEL_FILE) return IMAGE_CACHE.images;
  if (!FORMAT_MODEL_FILE || typeof indexedDB === "undefined") return {};
  try {
    const db = await openHandleDb();
    /** @type {any} */
    const rec = await new Promise(resolve => {
      const req = db.transaction(HANDLE_STORE).objectStore(HANDLE_STORE).get("images:" + FORMAT_MODEL_FILE.toLowerCase());
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => resolve(null);
    });
    IMAGE_CACHE = { file: FORMAT_MODEL_FILE, images: (rec && rec.images) || {} };
    return IMAGE_CACHE.images;
  } catch (e) {
    return {};
  }
}

/* =============================================================================
 * loadWorkbookFile() - Loads .twb/.twbx and parses BOTH titles AND colors
 * ============================================================================= */
async function loadWorkbookFile() {
  if (typeof window.showOpenFilePicker === "function") {
    try {
      const [handle] = await window.showOpenFilePicker({
        types: [{ description: "Tableau workbook", accept: { "application/octet-stream": [".twb", ".twbx"] } }]
      });
      const file = await handle.getFile();
      const titleMap = await readWorkbookFile(file);
      if (Object.keys(titleMap).length || FORMAT_MODEL_CACHE) {
        workbookFileHandle = handle;
        await storeWorkbookHandle(handle);
      }
      return titleMap;
    } catch (err) {
      if (err && err.name === "AbortError") return {};            // picker cancelled
      console.warn("[loadWorkbookFile] file picker unavailable, using file input:", err && err.message);
    }
  }
  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".twb,.twbx";
    input.style.display = "none";
    document.body.appendChild(input);

    input.onchange = async (event) => {
      document.body.removeChild(input);
      const file = /** @type {HTMLInputElement} */ (event.target).files[0];
      if (!file) {
        console.log("[loadWorkbookFile] No file selected");
        resolve({});
        return;
      }
      resolve(await readWorkbookFile(file));
    };

    input.oncancel = () => {
      document.body.removeChild(input);
      resolve({});
    };

    input.click();
  });
}

/* parses a picked .twb/.twbx: titles + format model, saved to the extension settings */
async function readWorkbookFile(file) {
  console.log(`[loadWorkbookFile] Reading: ${file.name}`);

  try {
    let xmlString;
    let zip = null;
    const ext = file.name.split(".").pop().toLowerCase();

    if (ext === "twb") {
      xmlString = await new Promise((res, rej) => {
        const reader = new FileReader();
        reader.onload = (e) => res(e.target.result);
        reader.onerror = () => rej(new Error("FileReader failed reading .twb"));
        reader.readAsText(file, "utf-8");
      });
    } else if (ext === "twbx") {
      const arrayBuffer = await new Promise((res, rej) => {
        const reader = new FileReader();
        reader.onload = (e) => res(e.target.result);
        reader.onerror = () => rej(new Error("FileReader failed reading .twbx"));
        reader.readAsArrayBuffer(file);
      });

      zip = await JSZip.loadAsync(arrayBuffer);
      const twbEntry = Object.values(zip.files).find(
        f => !f.dir && /\.twb$/i.test(f.name)
      );

      if (!twbEntry) {
        // an Excel file picked by mistake (an earlier export sits next to the workbook): say so plainly
        if (zip.file("xl/workbook.xml")) throw new Error(`"${file.name}" is an Excel file (an export?), not a Tableau workbook – pick the dashboard's .twbx (or .twb) instead.`);
        const entries = Object.values(zip.files).filter(f => !f.dir).map(f => f.name).slice(0, 20);
        throw new Error(`No .twb file found inside the .twbx archive. Entries: ${entries.join(", ") || "none"}`);
      }

      xmlString = await twbEntry.async("string");
    } else {
      throw new Error(`Unsupported file type ".${ext}". Please select a .twb or .twbx file.`);
    }

    // same parsing path as the Tableau Cloud auto-fetch; a manual file is also saved into the workbook settings
    const { titleMap, formatModel, images } = await applyWorkbookXml(xmlString, file.name, true, zip);
    showWorkbookLabel(`${file.name} (manual)`, describeModel(titleMap, formatModel, images), formatModel);

    return titleMap;

  } catch (err) {
    console.error("[loadWorkbookFile] Error:", err.message);
    alert(`Could not read workbook file:\n${err.message}`);
    return {};
  }
}

/* The save dialog → handle, null (no picker API / picker failed → plain download), or "cancelled". It needs a
 * click's user activation: it runs from the Save button's click (saveExport). */
async function chooseSaveTarget(suggestedName) {
  if (typeof window.showSaveFilePicker !== "function") return null;
  const options = {
    suggestedName,
    types: [{ description: "Excel workbook",
              accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }]
  };
  if (!workbookFileHandle) workbookFileHandle = await restoreWorkbookHandle();
  try {
    return await window.showSaveFilePicker(workbookFileHandle ? { ...options, startIn: workbookFileHandle } : options);
  } catch (err) {
    if (err && err.name === "AbortError") return "cancelled";
    if (workbookFileHandle) {
      try {                                   // stale handle (file moved / deleted) → default folder
        return await window.showSaveFilePicker(options);
      } catch (err2) {
        if (err2 && err2.name === "AbortError") return "cancelled";
      }
    }
    console.warn("[Export] save dialog unavailable, using browser download:", err && err.message);
    return null;
  }
}

/* =============================================================================
 * Saving the finished file. The save dialog is opened only once the file is built, by its own click (💾 Save
 * Excel file): a file picked before the export would sit empty on disk while the export runs – opened then,
 * Excel reports an invalid file, and the browser's final swap into it fails while Excel holds it.
 * ============================================================================= */
/** @type {{ blob: Blob, name: string } | null} the last export, until it is saved */
let pendingExport = null;

function downloadBlob(blob, name) {
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = name;
  link.click();
  URL.revokeObjectURL(link.href);
}

/** the panel's Save button, right under Export (made the first time) */
function saveButton() {
  let b = /** @type {HTMLButtonElement | null} */ (document.getElementById("save_button"));
  if (b) return b;
  const exp = document.getElementById("export_button");
  b = /** @type {HTMLButtonElement} */ (document.createElement("button"));
  b.id = "save_button";
  b.type = "button";
  b.className = exp ? String(exp.className || "").replace(/\bhas-tip\b/, "").trim() : "";
  b.style.marginTop = "6px";
  b.textContent = "💾 Save Excel file";
  b.hidden = true;
  b.addEventListener("click", () => { saveExport(); });
  if (exp && exp.insertAdjacentElement) exp.insertAdjacentElement("afterend", b);
  return b;
}

/** The built file: a Save button that opens the save dialog (or, without that dialog, a download right away).
 * → true when the Save button waits for a click */
function offerSave(blob, name) {
  if (typeof window.showSaveFilePicker !== "function") { downloadBlob(blob, name); return false; }
  pendingExport = { blob, name };
  const b = saveButton();
  b.hidden = false;
  b.disabled = false;
  return true;
}

/** The Save button's click: where to save, then the file written in one go. A file Excel has open cannot be
 * replaced – the button stays, to save again once it is closed (or under another name). */
async function saveExport() {
  if (!pendingExport) return;
  const b = saveButton(), { blob, name } = pendingExport;
  b.disabled = true;
  try {
    const target = await chooseSaveTarget(name);
    if (target === "cancelled") { appendExportStatus("Not saved yet – click 💾 Save Excel file to choose where"); return; }
    if (!target) {
      downloadBlob(blob, name);
      appendExportStatus(`downloaded as ${name}`);
    } else {
      const writable = await target.createWritable();
      try { await writable.write(blob); await writable.close(); }
      catch (e) { try { await writable.abort(); } catch (_) { /* already closed */ } throw e; }
      console.log(`[Export] saved to ${target.name}`);
      appendExportStatus(`saved as ${target.name} – you can open it now`);
    }
    pendingExport = null;
    b.hidden = true;
  } catch (err) {
    console.error("[Export] save failed", err);
    const locked = (err && /NoModificationAllowed|InvalidState/.test(err.name)) || /in use|locked|being used|access/i.test(String(err && err.message));
    appendExportStatus(locked
      ? "Not saved: the file is open in Excel. Close it there, then click 💾 Save Excel file again (or save under another name)."
      : `Not saved: ${err && err.message}. Click 💾 Save Excel file to try again.`);
  } finally {
    b.disabled = false;
  }
}

/* =============================================================================
 * applyWorkbookXml() – the one place a workbook becomes the active one (manual 📁 load and
 * Tableau Cloud auto-fetch alike): titles + format model parsed, kept in memory and in browser
 * storage (IndexedDB, with the XML so a newer extension version can parse it again), images and
 * custom shapes extracted. persistToSettings: a manual file is also saved into the workbook's
 * extension settings (old behaviour); the auto-fetch is not – it downloads the latest version on
 * every open.
 * ============================================================================= */
/** titles from the most recent load (auto or manual) @type {any} */
let TITLE_MAP_CACHE = null;

/** the workbook the formatting comes from: { id (Tableau Cloud workbook id, auto-load only), name, xml }
 *  – the other dashboards (tab order, views) are found through it */
let LOADED_WORKBOOK = null;

/** @returns {{ id: string | null, name: string, xml: string } | null} */
function loadedWorkbook() { return LOADED_WORKBOOK; }

/** set by the auto-load once it knows which published workbook this is @param {string} id */
function setLoadedWorkbookId(id) { if (LOADED_WORKBOOK) LOADED_WORKBOOK.id = id; }

/**
 * @param {string} xmlString @param {string} sourceName file / workbook name shown and used for the export's file name
 * @param {boolean} persistToSettings @param {JSZip | null} [zip] the .twbx archive (logos and other images)
 */
async function applyWorkbookXml(xmlString, sourceName, persistToSettings, zip = null) {
  const titleMap = parseTwbXmlInBrowser(xmlString);
  console.log(`[Workbook] ${sourceName}: parsed ${Object.keys(titleMap).length} titles`);
  // one DOM-based pass extracts all formatting (fonts, colours, number formats, shelves, marks …)
  const formatModel = parseTableauFormatting(xmlString);
  TITLE_MAP_CACHE = titleMap;
  FORMAT_MODEL_CACHE = formatModel;
  FORMAT_MODEL_FILE = sourceName;
  LOADED_WORKBOOK = { id: null, name: sourceName, xml: xmlString };       // id is set by the auto-load
  await storeFormatModel(sourceName, formatModel, titleMap, xmlString);
  // logos: only a .twbx carries the image files; custom shapes (button icons) are inside the workbook XML
  const images = { ...(zip ? await extractWorkbookImages(zip, formatModel) : {}), ...extractCustomShapes(xmlString) };
  IMAGE_CACHE = { file: sourceName, images };
  if (Object.keys(images).length) await storeWorkbookImages(sourceName, images);
  if (persistToSettings) {
    tableau.extensions.settings.set("twbTitleMap", JSON.stringify(titleMap));
    try {
      tableau.extensions.settings.set("twbFormatModel", JSON.stringify(formatModel));
    } catch (e) {
      try { tableau.extensions.settings.erase("twbFormatModel"); } catch (e2) { /* nothing saved */ }
      console.warn("[Workbook] Format model too large for settings – kept in memory only:", e.message);
    }
    tableau.extensions.settings.set("twbFileName", sourceName);
    try {
      await tableau.extensions.settings.saveAsync();
    } catch (e) {
      console.warn("[Workbook] settings.saveAsync failed (model kept in memory):", e.message);
    }
  }
  return { titleMap, formatModel, images };
}

/* =============================================================================
 * getTitleMap() - Gets titles from settings   * ============================================================================= */
function getTitleMap() {
  if (TITLE_MAP_CACHE) return TITLE_MAP_CACHE;          // latest auto-fetch or manual load
  try {
    const saved = tableau.extensions.settings.get("twbTitleMap");
    if (!saved) return RESTORED_TITLE_MAP || {};
    return JSON.parse(saved);
  } catch (err) {
    console.warn("[getTitleMap] Could not read settings:", err.message);
    return {};
  }
}

/* =============================================================================
 * getFormatModel() - format model from memory, else from extension settings
 * ============================================================================= */
let FORMAT_MODEL_CACHE = null;

/* the .twb / .twbx FORMAT_MODEL_CACHE was read from – the export's default file name starts with it */
let FORMAT_MODEL_FILE = null;

/** @returns {string | null} */
function formatModelFileName() { return FORMAT_MODEL_FILE; }

function getFormatModel() {
  if (FORMAT_MODEL_CACHE) return FORMAT_MODEL_CACHE;
  try {
    const saved = tableau.extensions.settings.get("twbFormatModel");
    FORMAT_MODEL_CACHE = saved ? JSON.parse(saved) : null;
    if (FORMAT_MODEL_CACHE) FORMAT_MODEL_FILE = tableau.extensions.settings.get("twbFileName") || null;
  } catch (err) {
    console.warn("[getFormatModel] Could not read settings:", err.message);
  }
  return FORMAT_MODEL_CACHE;
}

/* workbook model for THIS dashboard: memory / settings first, else the remembered workbook
 * whose worksheets match the dashboard best (no Load Workbook needed again) */
let RESTORED_TITLE_MAP = null;

let RESTORED_FILE_NAME = null;

async function ensureFormatModel() {
  const dashboard = tableau.extensions.dashboardContent.dashboard;
  const rank = m => { const r = checkWorkbookMatch(m, dashboard); return !r ? -1 : r.level === "ok" ? 2 + r.matched : r.level === "partial" ? 1 + r.matched / (r.total || 1) : 0; };
  const current = getFormatModel();
  if (current && !isFormatModelStale(current) && rank(current) >= 2) return current;
  const saved = await restoreFormatModels();
  /** @type {{ model: FormatModel, titleMap?: Record<string, string>, fileName: string, xml?: string } | null} */
  let best = null;
  // a stale model in memory / settings loses to the same workbook parsed again by this version
  let bestRank = current ? rank(current) - (isFormatModelStale(current) ? 0.5 : 0) : 0;
  saved.forEach(rec => { const r = rank(rec.model); if (r > bestRank) { best = rec; bestRank = r; } });
  if (best) {
    FORMAT_MODEL_CACHE = best.model;
    FORMAT_MODEL_FILE = best.fileName;
    RESTORED_TITLE_MAP = best.titleMap || null;
    RESTORED_FILE_NAME = best.fileName;
    if (best.xml && !LOADED_WORKBOOK) LOADED_WORKBOOK = { id: null, name: best.fileName, xml: best.xml };
    console.log(`[Workbook] using remembered workbook ${best.fileName} for this dashboard`);
  }
  return FORMAT_MODEL_CACHE;
}

/* main.js restores the cached model from the extension settings at start-up */
function setFormatModelCache(model, fileName) {
  FORMAT_MODEL_CACHE = model;
  FORMAT_MODEL_FILE = fileName || null;
}


/* ═══ visual/icon-sheet.js ════════════════════════════════════════════════════════════════════════ */
/* Button / icon worksheets: a custom shape drawn on empty shelves (zoom, info, navigation icons). Tableau
 * shows the icon; its fields are only flags for actions, so the export draws the icon, not a table. */

/**
 * The icon a worksheet shows when it is only Shape marks on empty Rows / Columns, without text: the shape
 * the workbook maps its current value to, else its Marks card shape. null = not an icon sheet.
 * @param {FormatModel | null} model @param {string} sheetName @param {SummaryData} summary
 * @returns {{ shape: string | null } | null} shape: a custom shape ("Zoom Icons/Zoom in.png"); null = one of
 *   Tableau's own shapes (circle, square …), which the export leaves out
 */
function tvIconSheet(model, sheetName, summary) {
  const sheet = model && model.sheets ? model.sheets[sheetName] : null;
  if (!sheet || !sheet.panes.length || sheet.rows.length || sheet.cols.length) return null;
  if (!sheet.panes.every(p => /^shape$/i.test(p.markClass || ""))) return null;
  if (sheet.panes.some(p => p.labelRuns.length || p.encodings.some(e => e.channel === "text" || e.channel === "label"))) return null;
  const pane = sheet.panes[0];
  let shape = null;
  // Shape = a field: the shape of the value it has now (the workbook's map, sheet before data source)
  const enc = pane.encodings.find(e => e.channel === "shape");
  if (enc) {
    const styles = [sheet.style, ...Object.values(model.datasourceStyles || {}), model.workbookStyle];
    const def = styles.flatMap(s => (s && s.shapes) || []).find(s => tfSameField(s.field, enc.field));
    const fmt = createSheetFormatter(model, sheetName);
    const ci = (summary.columns || []).findIndex(c => tfSameField(fmt.matchName(c.fieldName), enc.field));
    // the first mark whose value has a shape (the marks of a button sit on one spot)
    if (def && ci >= 0) shape = (summary.data || []).map(r => def.map[tfBucketKey(tfDvText(r[ci]))]).find(Boolean) || null;
  }
  // else the Marks card shape
  if (!shape) {
    const f = ((pane.style && pane.style.rules && pane.style.rules.mark) || []).find(x => x.attr === "shape" && !x.field);
    shape = f ? f.value : null;
  }
  return { shape: shape && !shape.startsWith(":") ? shape : null };     // ":filled/circle" … = Tableau's own
}


/* ═══ visual/image-renderer.js ════════════════════════════════════════════════════════════════════ */
/* Tableau-drawn images via createVizImageAsync (spec building, SVG → PNG). */

/* createVizImageAsync input spec, built from the same field roles as the native charts
 * (tvRoles, charts/model/roles.js): real shelves, colour, size, labels and view order.
 * One value axis → v1 spec; several measures (combo / dual axis, Measure Values) → v2 spec.
 * The API draws bar / line / area / square / circle / text marks only and has one field per
 * shelf and no Detail channel, so several category levels are joined into one ordered field. */
const VIZ_IMAGE_DISCRETE_PALETTES = new Set(["tableau10_10_0", "tableau20_10_0", "color_blind_10_0", "seattle_grays_10_0",
  "traffic_light_10_0", "superfishel_stone_10_0", "miller_stone_10_0", "nuriel_stone_10_0", "jewel_bright_10_0", "summer_10_0",
  "winter_10_0", "green_orange_cyan_yellow_10_0", "blue_red_brown_10_0", "purple_pink_gray_10_0", "tableau-10",
  "tableau-10-medium", "tableau-20", "cyclic_10_0"]);

const VIZ_IMAGE_ORDER_FIELD = "__order";

/** @param {VisualModel} visualModel @param {number} width @param {number} height @returns {Record<string, any>} */
function buildVizImageSpec(visualModel, width, height) {
  const vm = visualModel.viewModel;
  const T = VISUAL_TYPES;
  const type = visualModel.type;
  if (!vm || !vm.rows.length) throw new Error("visual has no data rows");
  const roles = tvRoles(vm, visualModel.formatModel, visualModel.source && visualModel.source.visualSpec);
  const markToken = visualModel.metadata.markToken || "";

  // readable, unique field names
  const taken = new Set([VIZ_IMAGE_ORDER_FIELD, "__mark"]);
  const keys = vm.cols.map((c, i) => {
    let k = String(c.label || c.name || `Field${i + 1}`);
    while (taken.has(k)) k += " ";
    taken.add(k);
    return k;
  });
  const kinds = new Map();                                   // column → "discrete" | "continuous"
  const field = (ci, continuous, extra) => {
    kinds.set(ci, continuous ? "continuous" : "discrete");
    return { field: keys[ci], type: continuous ? "continuous" : "discrete", ...extra };
  };
  const title = ci => tvMeasureLabel(vm, ci);

  // colour: TWB palette / mapping → named discrete palette or custom gradient end points
  let color = null;
  if (roles.color && !roles.color.measureNames && roles.color.ci >= 0) {
    const ci = roles.color.ci;
    if (roles.color.continuous) {
      const scale = tvColorScale(vm, roles, markToken);
      const nums = vm.rows.map(r => tfDvNum(r[ci])).filter(n => n !== null);
      if (scale && nums.length) {
        const min = Math.min(...nums), max = Math.max(...nums);
        color = min < 0 && max > 0
          ? field(ci, true, { palette: "custom-diverging", start: "#" + scale(min), end: "#" + scale(max) })
          : field(ci, true, { palette: "custom-sequential", end: "#" + scale(max) });
      } else color = field(ci, true);
    } else {
      const enc = visualModel.formatModel && vm.fmt.colorEncoding ? vm.fmt.colorEncoding() : null;
      const named = enc && enc.def && enc.def.paletteName;
      color = field(ci, false, { palette: VIZ_IMAGE_DISCRETE_PALETTES.has(named) ? named : "tableau10_10_0" });
    }
  }
  const sizeField = roles.size >= 0 ? field(roles.size, true) : null;
  const labelDim = roles.textDims[0] ?? roles.detailDims[0] ?? -1;
  /* No Detail channel: without a per-mark field Tableau would aggregate the marks (a scatter of
   * customers collapses to one point per colour). Label them; with many marks the label is an
   * invisible zero-width key so the marks stay separate without covering the chart. */
  const synthetic = {};
  const markLabel = ci => {
    const distinct = new Set(vm.rows.map(r => tvText(r[ci])));
    if (distinct.size <= 40 || roles.textDims.includes(ci)) return field(ci, false);
    const index = new Map([...distinct].map((v, i) => [v, i]));
    synthetic.__mark = row => index.get(tvText(row[ci])).toString(2).replace(/0/g, "​").replace(/1/g, "‌");
    return { field: "__mark", type: "discrete" };
  };

  /** @type {Record<string, any>} createVizImageAsync input */
  let spec = null;
  let categoryKey = null, categoryOf = null;
  const base = {
    description: (visualModel.title && visualModel.title.text) || visualModel.metadata.worksheetName,
    markcolor: "#" + tvMarkColor(roles)
  };

  if (type === T.MAP) {
    const lat = vm.cols.findIndex(c => /latitude/i.test(c.name));
    const lon = vm.cols.findIndex(c => /longitude/i.test(c.name));
    if (lat < 0 || lon < 0) throw new Error("the map has no latitude/longitude in its summary data");
    const encoding = { columns: field(lon, true, { hidden: true }), rows: field(lat, true, { hidden: true }) };
    if (color) encoding.color = color;
    if (sizeField) encoding.size = sizeField;
    const place = roles.detailDims.find(ci => ci !== lat && ci !== lon);
    if (place !== undefined) encoding.text = markLabel(place);
    spec = { ...base, mark: "circle", encoding };
  } else if (type === T.TREEMAP || type === T.BUBBLE) {
    if (roles.size < 0 || labelDim < 0) throw new Error("no Size measure and label dimension to lay out the marks");
    const encoding = { size: sizeField, text: field(labelDim, false) };
    if (color) encoding.color = color;
    spec = { ...base, mark: type === T.TREEMAP ? "square" : "circle", encoding };
  } else if (type === T.SCATTER) {
    const xm = roles.cols.values[0], ym = roles.rows.values[0];
    if (!xm || !ym) throw new Error("scatter needs a measure on Rows and on Columns");
    const encoding = { columns: field(xm.ci, true, { title: title(xm.ci) }), rows: field(ym.ci, true, { title: title(ym.ci) }) };
    if (color) encoding.color = color;
    if (sizeField) encoding.size = sizeField;
    if (labelDim >= 0) encoding.text = markLabel(labelDim);
    spec = { ...base, mark: "circle", encoding };
  } else {
    // bar / line / area / histogram / combo: one ordered category field + one or more value axes
    const valueShelf = roles.rows.values.length ? "rows" : roles.cols.values.length ? "cols" : null;
    if (!valueShelf) throw new Error("no continuous measure axis to draw");
    const catShelf = valueShelf === "rows" ? "cols" : "rows";
    const measures = roles[valueShelf].values;
    const catCis = [...roles[catShelf].dims, ...roles[valueShelf].dims].map(d => d.ci);
    if (catCis.length) {
      const order = new Map();
      categoryOf = row => catCis.map(ci => tvText(row[ci])).join(" · ");
      vm.rows.forEach(r => { const k = categoryOf(r); if (!order.has(k)) order.set(k, order.size); });
      if (catCis.some(ci => /^date/i.test(String(vm.cols[ci].dataType || "")))) {
        const firstRow = new Map();
        vm.rows.forEach(r => { const k = categoryOf(r); if (!firstRow.has(k)) firstRow.set(k, r); });
        const sorted = [...order.keys()].sort((a, b) => {
          for (const ci of catCis) { const d = tfNaturalCompare(firstRow.get(a)[ci], firstRow.get(b)[ci]); if (d) return d; }
          return 0;
        });
        order.clear();
        sorted.forEach((k, i) => order.set(k, i));
      }
      categoryKey = { name: catCis.map(ci => keys[ci]).join(" · "), order };
    }
    const catField = categoryKey ? { field: categoryKey.name, type: "discrete" } : null;
    const markOf = token => token === "line" ? "line" : token === "area" ? "area"
      : /^(circle|shape)$/.test(token) ? "circle" : token === "square" ? "square" : "bar";
    const shelfName = s => s === "cols" ? "columns" : "rows";
    if (measures.length === 1) {
      const m = measures[0];
      const encoding = {
        [shelfName(valueShelf)]: field(m.ci, true, { title: title(m.ci) }),
        sort: catField ? { field: categoryKey.name, sortby: VIZ_IMAGE_ORDER_FIELD, direction: "ascending" } : undefined
      };
      if (catField) encoding[shelfName(catShelf)] = catField;
      if (color) encoding.color = color;
      if (sizeField) encoding.size = sizeField;
      if (labelDim >= 0 && !catCis.includes(labelDim)) encoding.text = markLabel(labelDim);
      if (!encoding.sort) delete encoding.sort;
      spec = { ...base, mark: markOf(tvMeasureMark(roles, m.ref, markToken)), encoding };
    } else {
      const marks = measures.map((m, i) => markOf(tvMeasureMark(roles, m.ref, markToken, i, measures.length)));
      measures.forEach(m => field(m.ci, true));
      spec = {
        version: 2,
        description: base.description,
        vizlayout: { size: { width, height }, showcolorlegend: !!color },
        [shelfName(catShelf)]: catField ? [catField] : [],
        [shelfName(valueShelf)]: measures.map(m => ({ field: keys[m.ci], type: "continuous", title: title(m.ci) })),
        encodingaxis: shelfName(valueShelf),
        defaultencoding: { mark: marks[0] },
        encodings: marks.map(mark => (color ? { mark, color } : { mark }))
      };
    }
  }
  if (!spec.version) spec.size = { width, height };

  // data: every encoded column + the joined category and its order helper
  const perCategory = new Map();
  if (categoryKey) vm.rows.forEach(r => { const k = categoryOf(r); perCategory.set(k, (perCategory.get(k) || 0) + 1); });
  spec.data = {
    values: vm.rows.map(row => {
      const out = {};
      kinds.forEach((kind, ci) => {
        out[keys[ci]] = kind === "continuous" ? tfDvNum(row[ci]) : tvText(row[ci]);
      });
      for (const name in synthetic) out[name] = synthetic[name](row);
      if (categoryKey) {
        const k = categoryOf(row);
        out[categoryKey.name] = k;
        out[VIZ_IMAGE_ORDER_FIELD] = categoryKey.order.get(k) / perCategory.get(k);   // SUM over the category = its rank
      }
      return out;
    })
  };
  return spec;
}

function svgToPngDataUrl(svg, width, height) {
  return new Promise((resolve, reject) => {
    const blob = new Blob([svg], { type: "image/svg+xml" });
    const url = URL.createObjectURL(blob);
    const image = new Image();
    image.onload = () => {
      try {
        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        canvas.getContext("2d").drawImage(image, 0, 0, width, height);
        URL.revokeObjectURL(url);
        resolve(canvas.toDataURL("image/png"));
      } catch (err) {
        URL.revokeObjectURL(url);
        reject(err);
      }
    };
    image.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("Tableau returned an SVG that could not be rasterized"));
    };
    image.src = url;
  });
}

/** @param {VisualModel} visualModel @param {number} width @param {number} height @returns {Promise<string>} PNG, base64 */
async function renderTableauImage(visualModel, width, height) {
  if (typeof tableau === "undefined" || !tableau.extensions ||
      typeof tableau.extensions.createVizImageAsync !== "function") {
    throw new Error("Tableau image API is unavailable in this runtime");
  }
  const spec = buildVizImageSpec(visualModel, width, height);
  const svg = await tableau.extensions.createVizImageAsync(spec);
  if (typeof svg !== "string" || !svg.trim()) throw new Error("Tableau image API returned no SVG");
  return svgToPngDataUrl(svg, width, height);
}


/* ═══ visual/matrix.js ════════════════════════════════════════════════════════════════════════════ */
/* Heat maps, highlight tables, calendars and crosstabs as Tableau draws them: a matrix – the Rows shelf's members
 * down, the Columns shelf's members across, one cell per mark holding its label (else its size measure, else its
 * colour measure, else its shape), filled in the mark's colour. And the guard that keeps a table of marks from
 * dropping what the marks show. */

/** @param {SheetModel} sheet the panes Tableau draws: each with an id when there are several */
const drawnPanes = sheet => sheet.panes.length > 1 && sheet.panes.some(p => p.id) ? sheet.panes.filter(p => p.id) : sheet.panes;

/**
 * Pivots the view model into the matrix Tableau shows, in place. Only a view of dimensions on both shelves (no axes, no
 * Measure Names) whose marks each hold one value; anything else keeps its one-row-per-mark table.
 * @param {ViewModel} vm @param {FormatModel | null} model @returns {boolean} pivoted
 */
function pivotMatrix(vm, model) {
  const fmt = vm.fmt, sheet = fmt.sheetModel;
  if (!sheet || vm.markTable) return false;
  const shelf = [...sheet.rows, ...sheet.cols];
  if (shelf.some(r => r.name === "Measure Names" || r.type === "qk" || tvIsMeasureRef(model, r))) return false;
  const colRefs = sheet.cols;
  const ciOf = ref => vm.cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
  const colCis = colRefs.map(ciOf);
  if (!colRefs.length || colCis.some(ci => ci < 0)) return false;
  const rowCis = vm.headerOrder.filter(ci => !colCis.includes(ci) && sheet.rows.some(r => tfSameField(r, vm.cols[ci].ref)));
  if (!rowCis.length) return false;
  const panes = drawnPanes(sheet);
  if (panes.length !== 1) return false;
  const pane = panes[0];
  const enc = ch => pane.encodings.filter(e => e.channel === ch).map(e => e.field);
  const text = [...enc("text"), ...enc("label"), ...pane.labelRuns.flatMap(r => r.refs)].filter((r, i, a) => a.findIndex(x => tfSameField(x, r)) === i);
  if (text.length > 1) return false;
  const size = enc("size").find(r => tvIsMeasureRef(model, r));
  const color = enc("color")[0] || null;
  const valueRef = text[0] || size || (color && tvIsMeasureRef(model, color) ? color : null) || enc("shape")[0];
  const valueCi = valueRef ? ciOf(valueRef) : -1;
  if (valueCi < 0) return false;
  const colorCi = color ? ciOf(color) : -1;

  // one mark per cell: a cell that would need two values is not a matrix
  const colKey = r => colCis.map(ci => tfDvText(r[ci])).join("\u0001");
  const rowKey = r => rowCis.map(ci => tfDvText(r[ci])).join("\u0001");
  /** @type {Map<string, { head: DataValue[], cells: Map<string, DataValue[]> }>} */
  const groups = new Map();
  for (const r of vm.rows) {
    const k = rowKey(r);
    if (!groups.has(k)) groups.set(k, { head: r, cells: new Map() });
    const cells = groups.get(k).cells, ck = colKey(r);
    if (cells.has(ck) && tfDvText(cells.get(ck)[valueCi]) !== tfDvText(r[valueCi])) return false;
    cells.set(ck, r);
  }
  // the Columns members in Tableau's order: manual sort, else natural
  /** @type {Map<string, DataValue[]>} */
  const tuples = new Map();
  vm.rows.forEach(r => { const k = colKey(r); if (!tuples.has(k)) tuples.set(k, colCis.map(ci => r[ci])); });
  const ranks = colRefs.map(ref => {
    const m = fmt.manualSortFor(ref);
    return m ? { map: new Map(m.order.map((b, i) => [tfBucketKey(b), i])), desc: m.direction === "DESC" } : null;
  });
  const keys = [...tuples.keys()].sort((a, b) => {
    const ta = tuples.get(a), tb = tuples.get(b);
    for (let l = 0; l < colCis.length; l++) {
      let d = 0;
      if (ranks[l]) {
        const ra = ranks[l].map.get(tfNorm(tfDvText(ta[l]))), rb = ranks[l].map.get(tfNorm(tfDvText(tb[l])));
        d = (ra === undefined ? 1e9 : ra) - (rb === undefined ? 1e9 : rb);
        if (ranks[l].desc) d = -d;
      }
      if (!d) d = tfNaturalCompare(ta[l], tb[l]);
      if (d) return d;
    }
    return 0;
  });

  // each mark's colour, from the colour field over all marks (Tableau's palette / range)
  /** @type {(row: DataValue[]) => string | null} */
  let colorAt = () => null;
  const colorEnc = color && colorCi >= 0 ? fmt.colorEncodingOf(pane) : null;
  if (colorEnc) {
    const pick = dv => tfIsNull(dv) ? null : colorEnc.continuous ? tfDvNum(dv) : tfDvText(dv);
    const scale = tfBuildColorScale(fmt, colorEnc, vm.rows.map(r => pick(r[colorCi])));
    if (scale) colorAt = r => { const v = pick(r[colorCi]); return v === null ? null : scale(v) || null; };
  }

  const value = vm.cols[valueCi];
  const blank = { value: null, nativeValue: null, formattedValue: "" };
  /** @type {ViewColumn[]} */
  const cols = [...rowCis.map(ci => vm.cols[ci]),
    ...keys.map(k => {
      const label = tuples.get(k).map(tfDvText).join(" / ");
      // matrixKey: the Columns members of this column (a mark is its row's members and these – slicer formulas)
      return { name: label, label, ref: value.ref, dataType: value.dataType, pivoted: true, isHeader: false, matrixKey: tuples.get(k) };
    })];
  const colFields = colCis.map(ci => vm.cols[ci].ref);
  /** @type {(string | null)[][]} */
  const fills = [];
  const rows = [...groups.values()].map(g => {
    fills.push(keys.map(k => g.cells.has(k) ? colorAt(g.cells.get(k)) : null));
    return [...rowCis.map(ci => g.head[ci]), ...keys.map(k => g.cells.has(k) ? g.cells.get(k)[valueCi] : blank)];
  });
  const shownRow = rowCis.map((ci, i) => vm.order.includes(ci) ? i : -1).filter(i => i >= 0);
  vm.cols = cols;
  vm.rows = rows;
  vm.order = [...shownRow, ...keys.map((_, k) => rowCis.length + k)];
  vm.headerOrder = rowCis.map((_, i) => i);
  vm.showHeaderRow = true;
  vm.matrix = { colDims: colCis, valueCi, colorCi, colFields };
  vm.cellFill = (rowIdx, ci) => ci >= rowCis.length && fills[rowIdx] ? fills[rowIdx][ci - rowCis.length] : null;
  vm.notes.push(`matrix: ${groups.size} rows × ${keys.length} columns of ${value.name}`);
  return true;
}

/**
 * A table of marks must show what the marks show: when its cells would hold no value at all while the marks encode
 * measures (circles sized and coloured on a timeline), the table is lossy – the export writes every column instead
 * (TABLE_FALLBACK). Sets vm.lossless / vm.fallbackReason.
 * @param {ViewModel} vm @param {FormatModel | null} model @returns {boolean}
 */
function checkLossless(vm, model) {
  const sheet = vm.fmt.sheetModel;
  if (!sheet || vm.markTable || vm.matrix) return true;
  if (vm.order.some(ci => !vm.cols[ci].isHeader)) return true;           // the cells carry values (coloured as the marks)
  // measures by colour / size / label, and any field picking the marks' shape (a KPI's ▲▼)
  const encoded = drawnPanes(sheet).flatMap(p => p.encodings.filter(e => (/^(color|size|text|label)$/.test(e.channel) && tvIsMeasureRef(model, e.field)) ||
                                                                         e.channel === "shape")
    .map(e => ({ channel: e.channel === "label" ? "text" : e.channel, ref: e.field })));
  if (!encoded.length) return true;
  const names = [...new Set(encoded.map(e => {
    const c = vm.cols.find(x => x.ref && tfSameField(x.ref, e.ref));
    return `${c ? c.name : e.ref.name} (${e.channel})`;
  }))];
  vm.lossless = false;
  vm.fallbackReason = `its marks show ${names.join(", ")}, which a table of its headers would leave out – every column exported`;
  return false;
}


/* ═══ visual/view-model.js ════════════════════════════════════════════════════════════════════════ */
/* Summary data → view model (pivot, merge, sort, visible columns, titles). */

/**
 * Dashboard text boxes over a table built from marks, where one box can head several panes (e.g. a
 * diverging bar split into a negative and a positive axis): adjacent value columns whose panes are coloured
 * by the same field share a box. Matched only when the groups and the boxes line up one to one; the row
 * headers take the width left of the first box. Sets the columns' labels, widths and header spans.
 * @param {any} dash @param {string} sheetName @param {ViewColumn[]} cols @param {number[]} order
 * @param {SheetFormatter} fmt @returns {{ ids: string[], headerPx: number } | null}
 */
function markTableHeaders(dash, sheetName, cols, order, fmt) {
  const strip = tfTextBoxHeaderStrip(dash, sheetName);
  if (!strip || !dash.width) return null;
  const panes = fmt.panesInOrder();
  const colourOf = ci => {
    const p = panes.find(x => x.refs.some(r => tfSameField(r, cols[ci].ref)));
    const e = p && p.pane.encodings.find(x => x.channel === "color" && x.field);
    return e ? tfRefKey(e.field) : null;
  };
  // the pane a column is drawn in: a pane on its own axis is one column of the table in Tableau (its label may
  // hold several fields, e.g. "12% ▲")
  const paneOf = ci => {
    const k = panes.findIndex(x => x.refs.some(r => tfSameField(r, cols[ci].ref)));
    return k >= 0 ? "pane:" + k : null;
  };
  const values = order.filter(ci => !cols[ci].isHeader), heads = order.filter(ci => cols[ci].isHeader);
  /** @param {(ci: number) => string | null} keyOf @returns {{ key: string | null, cols: number[] }[]} */
  const groupBy = keyOf => {
    const groups = [];
    values.forEach(ci => {
      const key = keyOf(ci), last = groups[groups.length - 1];
      if (last && key && last.key === key) last.cols.push(ci); else groups.push({ key, cols: [ci] });
    });
    return groups;
  };
  // A strip that starts at the worksheet's left edge also heads the row-header columns (Department, Race …):
  // its first boxes go to those, the rest to the value groups – by colour, else by pane
  let headBoxes = 0, groups = null;
  if (heads.length && Math.abs(strip.zones[0].x - strip.ws.x) <= TF_ZONE_TOL) {
    for (const keyOf of [colourOf, paneOf]) {
      const g = groupBy(keyOf);
      if (heads.length + g.length === strip.zones.length) { headBoxes = heads.length; groups = g; break; }
    }
  }
  if (!groups) {
    const g = groupBy(colourOf);
    if (g.length !== strip.zones.length) return null;
    groups = g;
  }
  const px = units => units / 100000 * dash.width;
  heads.slice(0, headBoxes).forEach((ci, k) => {
    const z = strip.zones[k];
    cols[ci].label = z.text;
    cols[ci].labelProps = z.props;
    cols[ci].labelRuns = z.runs;
    cols[ci].zoneWidthPx = px(z.w);
  });
  groups.forEach((g, k) => {
    const z = strip.zones[headBoxes + k], first = cols[g.cols[0]];
    first.label = z.text;
    first.labelProps = z.props;
    first.labelRuns = z.runs;
    first.headerSpan = g.cols.length;
    g.cols.forEach((ci, j) => {
      if (j) { cols[ci].label = ""; cols[ci].headerCovered = true; }
      cols[ci].zoneWidthPx = px(z.w) / g.cols.length;
    });
  });
  if (!headBoxes) {
    const left = px(strip.zones[0].x - strip.ws.x);
    if (heads.length && left > 0) heads.forEach(ci => { cols[ci].zoneWidthPx = left / heads.length; });
  }
  return { ids: strip.zones.map(z => z.id), headerPx: Math.max(...strip.zones.map(z => z.h)) / 100000 * (dash.height || 0) };
}

/**
 * Indicator fields in one mark label – e.g. a green "▲" field and a red "▼" field after a percentage, only one of
 * them set on any row – are one spot in Tableau's cell. Written as separate columns they leave a gap wherever the
 * first is empty, so consecutive ones of the same pane are joined into one column; each cell keeps the style
 * (colour) of the field it came from. Only short text (≤ 3 characters) that never overlaps on a row is joined.
 * @param {SheetFormatter} fmt @param {ViewColumn[]} cols @param {any[][]} rows @param {number[]} order
 * @returns {number[]} the order with each joined run replaced by one new column
 */
function mergeIndicatorColumns(fmt, cols, rows, order) {
  if (!fmt || !fmt.hasModel || !rows.length) return order;
  const panes = fmt.panesInOrder();
  const paneOf = ci => cols[ci].ref ? panes.findIndex(x => x.refs.some(r => tfSameField(r, cols[ci].ref))) : -1;
  const isIndicator = ci => {
    const c = cols[ci];
    if (c.isHeader || c.pivoted || !c.ref || paneOf(ci) < 0) return false;
    return rows.every(r => {
      const dv = r[ci];
      if (tfIsNull(dv)) return true;
      const v = dv.nativeValue !== undefined ? dv.nativeValue : dv.value;
      return typeof v !== "number" && tfDvText(dv).trim().length <= 3;
    });
  };
  const out = [];
  for (let k = 0; k < order.length;) {
    const ci = order[k];
    let run = [ci];
    if (isIndicator(ci)) {
      while (k + run.length < order.length && isIndicator(order[k + run.length]) &&
             paneOf(order[k + run.length]) === paneOf(ci)) run.push(order[k + run.length]);
    }
    const filled = (r, c) => !tfIsNull(r[c]) && tfDvText(r[c]).trim() !== "";
    const exclusive = run.length > 1 && rows.every(r => run.filter(c => filled(r, c)).length <= 1);
    if (!exclusive) { out.push(ci); k++; continue; }
    // the joined column: the first field's identity, each row's value and style from the field that is set
    const at = cols.length;
    const srcOf = rows.map(r => run.find(c => filled(r, c)) ?? run[0]);
    cols.push({ ...cols[run[0]], name: cols[run[0]].name, label: cols[run[0]].label,
                rowStyleRefs: srcOf.map(c => cols[c].ref), joined: run.map(c => cols[c].name) });
    rows.forEach((r, i) => { r[at] = r[srcOf[i]]; });
    out.push(at);
    k += run.length;
  }
  return out;
}

/* Tableau names a quick table calculation after its measure: a running sum and the plain sum both arrive as
 * "SUM(Value Ordered)". Of the columns sharing a name, the one holding the calculation takes the sheet's
 * table-calculation field: a running sum holds the plain column's grand total (its last value), a percent of
 * total adds up to 1; otherwise the calculation is the first of them, as Tableau lists it (cum… before sum…). */
function splitTableCalcColumns(cols, rows, fmt) {
  const groups = new Map();
  cols.forEach((c, i) => { if (c.ref) { const k = tfNorm(c.name); groups.set(k, [...(groups.get(k) || []), i]); } });
  groups.forEach(idx => {
    if (idx.length < 2 || !fmt.hasModel || idx.every(i => cols[i].exact)) return;
    const calcs = fmt.tableCalcRefs(cols[idx[0]].ref);
    if (!calcs.length) return;
    const nums = i => rows.map(r => tfDvNum(r[i])).filter(v => v !== null);
    const sum = i => nums(i).reduce((a, b) => a + b, 0);
    const near = (a, b) => Math.abs(a - b) <= Math.max(1e-9, Math.abs(b) * 1e-6);
    let free = [...idx];
    calcs.forEach(calc => {
      if (free.length < 2) return;
      const test = fmt.isRunningTotal(calc) ? (i, j) => nums(i).some(v => near(v, sum(j)))
                 : /^pcto$/i.test(calc.deriv || "") ? (i) => near(sum(i), 1) || near(sum(i), 100) : null;
      let pick = test ? free.find(i => free.some(j => j !== i && test(i, j))) : undefined;
      if (pick === undefined) pick = free[0];
      cols[pick].ref = calc;
      free = free.filter(i => i !== pick);
    });
  });
}

/**
 * @param {FormatModel | null} model
 * @param {string} sheetName
 * @param {SummaryData} summary
 * @param {{ dashboardName?: string, displayName?: string, [key: string]: any }} [opts]
 * @returns {ViewModel}
 */
function buildViewModel(model, sheetName, summary, opts = {}) {
  const fmt = createSheetFormatter(model, sheetName);
  const notes = [];
  /** @type {ViewColumn[]} */
  // a column's field: its id when Tableau gives one ("[ds].[pcto:sum:Sales:qk]" – exact, even for table
  // calculations captioned like their measure), else its caption ("SUM(Sales)")
  let cols = (summary.columns || []).map((c, i) => {
    const name = c.fieldName || c.fieldId || `Col${i + 1}`;
    const byId = fmt.matchFieldId(c.fieldId);
    return { name, dataType: c.dataType, ref: byId || fmt.matchName(name), exact: !!byId };
  });
  let rows = (summary.data || []).map(r => r.slice());
  splitTableCalcColumns(cols, rows, fmt);

  // ── 1. pivot Measure Names / Measure Values ─────────────────────────────
  const mnI = cols.findIndex(c => /^measure names$/i.test(c.name));
  const mvI = cols.findIndex(c => /^measure values$/i.test(c.name));
  if (mnI >= 0 && mvI >= 0) {
    const keyIdx = cols.map((_, i) => i).filter(i => i !== mnI && i !== mvI);
    const order = fmt.measureOrder();
    const measures = [];                         // [{text, ref, rank}]
    const groups = new Map();
    rows.forEach(r => {
      const text = tfDvText(r[mnI]);
      if (!measures.some(m => m.text === text)) {
        const ref = fmt.measureRefForName(text);
        const pos = ref ? order.findIndex(o => tfSameField(o, ref)) : -1;
        measures.push({ text, ref, rank: pos >= 0 ? pos : 1e6 + measures.length });
      }
      const key = keyIdx.map(i => tfDvText(r[i])).join("\u0001");
      if (!groups.has(key)) groups.set(key, { keyRow: keyIdx.map(i => r[i]), vals: {} });
      const g = groups.get(key);
      if (g.vals[text] === undefined) g.vals[text] = r[mvI];
    });
    measures.sort((a, b) => a.rank - b.rank);
    cols = [...keyIdx.map(i => cols[i]),
            ...measures.map(m => ({ name: m.text, ref: m.ref, pivoted: true, dataType: "float" }))];
    rows = [...groups.values()].map(g => [...g.keyRow, ...measures.map(m => g.vals[m.text] || null)]);
    notes.push(`pivoted ${measures.length} measures → ${rows.length} rows`);
  }

  // column roles
  cols.forEach(c => {
    c.isHeader = !c.pivoted && fmt.hasModel ? fmt.isHeaderField(c.ref)
               : !c.pivoted && !/^(int|float|real|integer|number)$/i.test(String(c.dataType || ""));
  });

  // ── 2. merge stacked multi-pane rows ────────────────────────────────────
  const hdrIdx = cols.map((c, i) => c.isHeader ? i : -1).filter(i => i >= 0);
  if (hdrIdx.length && rows.length > 1) {
    const groups = new Map();
    rows.forEach(r => {
      const key = hdrIdx.map(i => tfDvText(r[i])).join("\u0001");
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(r);
    });
    if (groups.size < rows.length) {
      let conflict = false;
      const merged = [];
      for (const g of groups.values()) {
        const out = g[0].slice();
        for (let ci = 0; ci < cols.length && !conflict; ci++) {
          const vals = [...new Set(g.filter(r => !tfIsNull(r[ci])).map(r => tfDvText(r[ci])))];
          if (vals.length > 1) conflict = true;
          else { const hit = g.find(r => !tfIsNull(r[ci])); if (hit) out[ci] = hit[ci]; }
        }
        if (conflict) break;
        merged.push(out);
      }
      if (!conflict) { notes.push(`merged ${rows.length} stacked rows → ${merged.length}`); rows = merged; }
      else notes.push("rows share header values but differ – not merged");
    }
  }

  // ── 3. sort like the view ───────────────────────────────────────────────
  const sortLevels = [];
  (fmt.hasModel ? fmt.headerRefs() : []).forEach(ref => {
    const ci = cols.findIndex(c => c.isHeader && c.ref && tfSameField(c.ref, ref));
    if (ci < 0) return;
    const manual = fmt.manualSortFor(ref);
    const byMeasure = fmt.measureSortFor(ref);
    const natural = fmt.naturalSortFor ? fmt.naturalSortFor(ref) : null;
    let mi = -1;
    if (byMeasure) mi = cols.findIndex(c => c.ref && tfSameField(c.ref, byMeasure.measure));
    // sort measure not in the summary data (e.g. a funnel sorted by a hidden field) → the row's own total
    sortLevels.push({ ci, manual, byMeasure, natural, mi });
  });
  if (sortLevels.length) {
    // measure sort aggregates the measure over the rows sharing the same outer path
    const aggCache = sortLevels.map((lvl, L) => {
      if (!lvl.byMeasure) return null;
      const m = new Map();
      rows.forEach(r => {
        const k = sortLevels.slice(0, L + 1).map(x => tfDvText(r[x.ci])).join("\u0001");
        const v = lvl.mi >= 0 ? tfDvNum(r[lvl.mi])
          : cols.reduce((s, c, i) => s + (!c.isHeader && tfDvNum(r[i]) !== null ? tfDvNum(r[i]) : 0), 0);
        m.set(k, (m.get(k) || 0) + (v || 0));
      });
      return m;
    });
    const manualRank = lvl => {
      if (!lvl.manual) return null;
      const map = new Map(lvl.manual.order.map((b, i) => [tfBucketKey(b), i]));
      return map;
    };
    const ranks = sortLevels.map(manualRank);
    rows.sort((a, b) => {
      for (let L = 0; L < sortLevels.length; L++) {
        const lvl = sortLevels[L];
        let d = 0;
        if (ranks[L]) {
          const ra = ranks[L].get(tfNorm(tfDvText(a[lvl.ci]))), rb = ranks[L].get(tfNorm(tfDvText(b[lvl.ci])));
          d = (ra === undefined ? 1e9 : ra) - (rb === undefined ? 1e9 : rb);
          if (lvl.manual.direction === "DESC") d = -d;
        } else if (lvl.byMeasure) {
          const path = r => sortLevels.slice(0, L + 1).map(x => tfDvText(r[x.ci])).join("\u0001");
          d = (aggCache[L].get(path(a)) || 0) - (aggCache[L].get(path(b)) || 0);
          if (lvl.byMeasure.direction === "DESC") d = -d;
        }
        if (!d) d = tfNaturalCompare(a[lvl.ci], b[lvl.ci]) * (!lvl.manual && !lvl.byMeasure && lvl.natural && lvl.natural.direction === "DESC" ? -1 : 1);
        if (d) return d;
      }
      return 0;
    });
  }

  // ── URL actions → link columns (kept visible even when the field is only on Detail)
  const linkActions = fmt.urlActions(opts.dashboardName);
  linkActions.forEach(a => {
    const refs = tfExtractRefs(a.expression);
    if (!refs.length) return;
    const ci = cols.findIndex(c => !c.isHeader && c.ref && tfNorm(c.ref.name) === tfNorm(refs[0].name));
    if (ci < 0) return;
    cols[ci].link = { caption: a.caption, expression: a.expression,
      parts: refs.map(r => ({ token: r, ci: cols.findIndex(c => c.ref && tfNorm(c.ref.name) === tfNorm(r.name)) })) };
  });

  // ── 4. visible columns in visual order + header labels ──────────────────
  let order;
  if (!fmt.hasModel) {
    order = cols.map((_, i) => i);
  } else {
    const panes = fmt.panesInOrder();
    const headerRefs = fmt.headerRefs();
    /** @type {{ i: number, g: number, a: number, b: number }[]} column, group, then sort keys */
    const place = [];
    cols.forEach((c, i) => {
      if (c.isHeader) {
        if (fmt.isLabelHidden(c.ref)) return;                       // Format → Hide header
        const pos = headerRefs.findIndex(r => tfSameField(r, c.ref));
        place.push({ i, g: 0, a: pos, b: 0 });
      } else if (c.pivoted) {
        place.push({ i, g: 1, a: 0, b: i });
      } else if (!c.ref) {
        place.push({ i, g: 2, a: 1e6, b: i });                     // unknown → keep, at the end
      } else {
        /** @type {{ a: number, b: number } | null} */
        let hit = null;
        panes.forEach((p, pi) => {
          const k = p.refs.findIndex(r => tfSameField(r, c.ref));
          if (k >= 0 && !hit) hit = { a: pi, b: k };
        });
        if (hit) place.push({ i, g: 2, a: hit.a, b: hit.b });      // shown as a mark label
        else if (c.link) {                                         // link column: in the pane that carries the field
          const pi = panes.findIndex(p => p.pane.encodings.some(e => tfSameField(e.field, c.ref)));
          place.push({ i, g: 2, a: pi >= 0 ? pi : 1e5, b: 999 });
        }
        // else: detail / tooltip / colour-only / axis measure → not visible
      }
    });
    place.sort((x, y) => x.g - y.g || x.a - y.a || x.b - y.b);
    order = place.map(p => p.i);
    if (!order.length) { order = cols.map((_, i) => i); notes.push("no visible columns resolved – exporting all"); }
  }

  const labelsRows = fmt.fieldLabelsShown("rows"), labelsCols = fmt.fieldLabelsShown("cols");
  cols.forEach(c => {
    if (c.pivoted) c.label = c.name;                                          // Measure Names alias
    else if (c.isHeader) {
      const shown = fmt.shelfOf(c.ref) === "cols" ? labelsCols : labelsRows;
      c.label = shown ? fmt.captionFor(c.ref, c.name) : "";
    } else c.label = fmt.captionFor(c.ref, c.name);
    if (c.link && c.link.caption) c.label = c.link.caption;
  });

  order = mergeIndicatorColumns(fmt, cols, rows, order);

  // ── dashboard text boxes drawn as column headers (strict match, else captions stay) ──
  const dashM = model && opts.dashboardName && model.dashboards && model.dashboards[opts.dashboardName];
  let headerZoneIds = [];
  let headerPx = 0;
  if (FORMAT_CONFIG.textBoxHeaders && dashM && !labelsRows) {
    const tb = tfTextBoxHeaders(dashM, sheetName, order.length);
    if (tb) {
      order.forEach((ci, k) => {
        const z = tb.zones[k];
        cols[ci].label = z.text;
        cols[ci].labelProps = z.props;
        if (dashM.width) cols[ci].zoneWidthPx = z.w / 100000 * dashM.width;
      });
      headerZoneIds = tb.zones.map(z => z.id);
      notes.push(`headers from ${tb.zones.length} dashboard text boxes`);
    } else {
      const grouped = markTableHeaders(dashM, sheetName, cols, order, fmt);
      if (grouped) {
        headerZoneIds = grouped.ids;
        headerPx = grouped.headerPx;
        notes.push(`headers from ${grouped.ids.length} dashboard text boxes over grouped panes`);
      }
    }
  }
  const showHeaderRow = order.some(i => cols[i].label);

  // title
  const valuesFor = ref => {
    const ci = cols.findIndex(c => c.ref && tfSameField(c.ref, ref));
    if (ci < 0) return [];
    return [...new Set(rows.map(r => tfDvText(r[ci])).filter(Boolean))];
  };
  const title = fmt.renderTitle(valuesFor, opts.displayName || sheetName);
  const dash = model && opts.dashboardName && model.dashboards && model.dashboards[opts.dashboardName];
  const zone = dash && dash.zones.find(z => z.name === sheetName && z.type === "worksheet");
  const showTitle = zone ? zone.showTitle : true;

  if (notes.length) tfLog(`View model "${sheetName}": ${notes.join("; ")}`);
  return {
    fmt, cols, rows, order, notes, showHeaderRow, title, showTitle, headerZoneIds, headerPx,
    dashboardName: opts.dashboardName,
    kind: fmt.isChart() ? "chart" : "table",
    headerOrder: sortLevels.map(l => l.ci)          // all header cols (incl. hidden) outer → inner
  };
}


/* ═══ visual/visual-model.js ══════════════════════════════════════════════════════════════════════ */
/* View model + classification → visual model, and renderer selection. */

/**
 * @param {FormatModel | null} model
 * @param {string} sheetName
 * @param {SummaryData} summary
 * @param {{ dashboardName?: string, displayName?: string, visualSpec?: any, visualSpecError?: any, layout?: any }} [opts]
 * @returns {VisualModel}
 */
function buildVisualModel(model, sheetName, summary, opts = {}) {
  const vm = buildViewModel(model, sheetName, summary, opts);
  const spec = opts.visualSpec || null;
  const marks = resolveVisualMarks(spec, vm, model);
  const evidence = [];
  const type = classifyVisualType(spec, vm, model, evidence);
  vm.markTable = (type === VISUAL_TYPES.TABLE || type === VISUAL_TYPES.KPI) && isMarkTable(vm, model);
  // drawn as cells: a heat map / crosstab as Tableau's matrix, and never a table that drops what its marks show
  if (type === VISUAL_TYPES.HEATMAP || type === VISUAL_TYPES.TABLE) {
    if (pivotMatrix(vm, model)) evidence.push("dimensions on Rows and Columns, one mark per cell: a matrix");
    else checkLossless(vm, model);
  }
  const dimensions = vm.order.filter(i => vm.cols[i].isHeader).map(i => vm.cols[i]);
  const measures = vm.order.filter(i => !vm.cols[i].isHeader).map(i => vm.cols[i]);
  return {
    type,
    title: vm.title,
    data: { rows: vm.rows, columns: vm.cols, dimensions, measures },
    axes: spec && (spec.axes || spec.axis) || {},
    encodings: spec && (spec.encodings || spec.encoding) || {},
    panes: spec && spec.panes || [],
    style: spec && (spec.style || spec.styles) || {},
    layout: opts.layout || null,
    source: { visualSpec: spec },
    formatModel: model || null,
    metadata: {
      worksheetName: sheetName,
      markClass: vm.fmt && vm.fmt.colorEncoding ? ((vm.fmt.colorEncoding() || {}).markClass || null) : null,
      markToken: marks.tokens[0] || "",
      markTokens: marks.tokens,
      markSource: marks.source,
      visualSpecAvailable: !!spec,
      visualSpecError: opts.visualSpecError ? String(opts.visualSpecError.message || opts.visualSpecError) : null,
      visualSpecKeys: spec && typeof spec === "object" ? Object.keys(spec) : [],
      imageApiAvailable: typeof tableau !== "undefined" && tableau.extensions &&
        typeof tableau.extensions.createVizImageAsync === "function"
    },
    viewModel: vm,
    diagnostics: [],
    evidence
  };
}

/* -----------------------------------------------------------------------------
 * "Charts as tables" (panel switch): every chart is exported as a crosstab – a table of its marks, coloured
 * with the chart's Tableau colours (colour encoding: category colours or the colour scale) – instead of a
 * native Excel chart. Tables, KPI cards and heat maps are cells either way. Charts in very small zones (KPI
 * trend lines) stay charts while FORMAT_CONFIG.crosstabKeepsSparklines is on.
 * --------------------------------------------------------------------------- */
const SPARKLINE_MAX_HEIGHT_PX = 80;

/** The panel's choice for this workbook / browser, else FORMAT_CONFIG.visualsAsTables. */
function visualsAsTables() {
  const saved = readSaved("visualsAs");
  return saved ? saved === "tables" : !!FORMAT_CONFIG.visualsAsTables;
}
/** A chart small enough to be a KPI trend line (sparkline): its dashboard zone is very short. */
function isSparklineZone(layout) {
  return !!(layout && layout.heightPx && layout.heightPx < SPARKLINE_MAX_HEIGHT_PX);
}

/* cell-native → cells; mappable charts → native Excel chart; the rest → Tableau image → data table;
 * opts.asTable ("Charts as tables") → every chart as a table */
/** @param {VisualModel} visualModel @param {{ asTable?: boolean }} [opts] @returns {RendererDecision} */
function chooseVisualRenderer(visualModel, opts = {}) {
  const type = visualModel.type;
  if (VISUAL_RENDERERS.cellTypes.has(type)) return VISUAL_RENDERERS.cell;
  if (opts.asTable) return VISUAL_RENDERERS.crosstab;
  if (FORMAT_CONFIG.nativeCharts && VISUAL_RENDERERS.chartTypes.has(type)) return VISUAL_RENDERERS.chart;
  return imageOrFallback(visualModel);
}

/** @param {VisualModel} visualModel @param {string} [why] @returns {RendererDecision} */
function imageOrFallback(visualModel, why) {
  const prefix = why ? why + "; " : "";
  if (VISUAL_RENDERERS.imageTypes.has(visualModel.type) && visualModel.metadata.imageApiAvailable) {
    return { ...VISUAL_RENDERERS.image, reason: prefix + VISUAL_RENDERERS.image.reason };
  }
  const reason = !visualModel.metadata.imageApiAvailable && VISUAL_RENDERERS.imageTypes.has(visualModel.type)
    ? "Tableau image API unavailable – exported as data"
    : `${visualModel.type.toLowerCase()} visuals cannot be drawn in Excel – exported as data`;
  return { ...VISUAL_RENDERERS.fallback, reason: prefix + reason };
}

const VISUAL_RENDERERS = Object.freeze({
  cell: Object.freeze({ renderer: "cell", status: "success", reason: "cell-native visual" }),
  chart: Object.freeze({ renderer: "excel-chart", status: "pending", reason: "native Excel chart" }),
  image: Object.freeze({ renderer: "tableau-image", status: "pending", reason: "Tableau image renderer selected" }),
  cellTypes: /** @type {Set<VisualType>} */ (new Set([VISUAL_TYPES.TABLE, VISUAL_TYPES.KPI, VISUAL_TYPES.HEATMAP])),
  chartTypes: /** @type {Set<VisualType>} */ (new Set([VISUAL_TYPES.BAR, VISUAL_TYPES.COLUMN, VISUAL_TYPES.LINE, VISUAL_TYPES.AREA,
                       VISUAL_TYPES.PIE, VISUAL_TYPES.SCATTER, VISUAL_TYPES.COMBO, VISUAL_TYPES.HISTOGRAM,
                       VISUAL_TYPES.WATERFALL, VISUAL_TYPES.BOXPLOT, VISUAL_TYPES.GANTT, VISUAL_TYPES.TREEMAP,
                       VISUAL_TYPES.BUBBLE])),
  // createVizImageAsync draws bar / line / area / square / circle / text marks only (no pie, gantt or polygon
  // marks). Maps – symbol or filled – are neither: without the basemap they don't read as maps, so they are
  // exported as a table of their marks (data-fallback)
  imageTypes: /** @type {Set<VisualType>} */ (new Set([VISUAL_TYPES.BAR, VISUAL_TYPES.COLUMN, VISUAL_TYPES.LINE, VISUAL_TYPES.AREA, VISUAL_TYPES.SCATTER,
                       VISUAL_TYPES.COMBO, VISUAL_TYPES.HISTOGRAM, VISUAL_TYPES.TREEMAP, VISUAL_TYPES.BUBBLE])),
  fallback: Object.freeze({ renderer: "data-fallback", status: "warning", reason: "visual renderer not implemented yet" }),
  // chosen, not a fallback: the table is the requested output
  crosstab: Object.freeze({ renderer: "data-fallback", status: "success", reason: "exported as a table (Charts as tables)", crosstab: true })
});


/* ═══ visual/semantics.js ═════════════════════════════════════════════════════════════════════════ */
/* What a Tableau visual is, in Tableau's terms, and the evidence for it: the shelves, marks and encodings of the
 * worksheet read as one visual ("Stacked Bar", "Bump Chart", "Filled Map" …). The conversion report shows it; the
 * classifier (classify.js) decides how the visual is drawn, this names what it is. */

/* quick table calculations, as their derivation prefix ("pcto:sum:Sales:qk") */
const TABLE_CALCS = { pcto: "% of Total", cum: "Running Total", rsum: "Running Total", rank: "Rank", pctrank: "Percentile",
                      diff: "Difference", pdiff: "% Difference", mavg: "Moving Average", ytd: "YTD Total",
                      cgr: "Compound Growth Rate", yoy: "Year over Year Growth", ytdgr: "YTD Growth" };

const GEO_GENERATED = /^(latitude|longitude) \(generated\)$/i;

/**
 * A field reference taken apart: its quick table calculations (outer → inner), its aggregation or date part, the
 * underlying field and how Tableau labels it ("% of Total of SUM(Sales)").
 * @param {FormatModel | null} model @param {FieldRef} ref
 */
function describeField(model, ref) {
  if (ref.name === "Measure Names") return { label: "Measure Names", calcs: [], agg: null, info: null, name: ref.name };
  if (ref.name === "Multiple Values") return { label: "Measure Values", calcs: [], agg: null, info: null, name: ref.name };
  const chain = ref.deriv ? [ref.deriv.toLowerCase()] : [];
  let name = ref.name || "";
  for (let m = name.match(/^([a-z_]+):(.+)$/i); m && (TABLE_CALCS[m[1].toLowerCase()] || TF_DERIV_LABEL[m[1].toLowerCase()] ||
       /^(none|usr)$/i.test(m[1])); m = name.match(/^([a-z_]+):(.+)$/i)) {
    chain.push(m[1].toLowerCase());
    name = m[2];
  }
  const calcs = chain.filter(c => TABLE_CALCS[c]);
  const agg = chain.find(c => !TABLE_CALCS[c]) || null;
  const info = model ? tfFieldInfo(model, { ds: ref.ds, deriv: null, name, inner: name, type: ref.type }) : null;
  const caption = (info && info.caption) || (/^Calculation_\d+$/i.test(name) ? "Calculation" : name.replace(/_/g, " "));
  const AGG = !agg || agg === "none" ? null : agg === "usr" ? "AGG" : (TF_DERIV_LABEL[agg] || agg.toUpperCase());
  let label = AGG ? `${AGG}(${caption})` : caption;
  if (calcs.length) label = calcs.map(c => TABLE_CALCS[c]).join(" of ") + " of " + label;
  return { label, calcs, agg, info, name };
}

/** @param {Pane} p @returns {string} the pane's mark ("" = Automatic) */
const markOf = p => String(p.markClass || "").toLowerCase().replace(/[\s_-]+/g, "").replace(/^automatic$/, "");

/**
 * @param {VisualModel} visualModel
 * @param {FormatModel | null} model
 * @returns {VisualSemantics}
 */
function describeVisual(visualModel, model) {
  const vm = visualModel.viewModel;
  const sheet = vm && vm.fmt ? vm.fmt.sheetModel : null;
  const type = visualModel.type;
  const evidence = [...(visualModel.evidence || [])];
  const source = visualModel.metadata.markSource;
  const confidence = sheet && source === "twb" ? "high" : sheet || source === "live" ? "medium" : "low";
  /** @type {VisualSemantics} */
  const out = { visual: "", family: type, confidence, evidence, encodings: {}, calculations: [], flags: {} };
  if (!sheet) {
    out.visual = { TABLE: "Text Table", KPI: "KPI", CUSTOM: "Viz Extension" }[type] || titleCase(type);
    if (!visualModel.formatModel) evidence.push("no workbook loaded: recognised from the live data only");
    return out;
  }
  const label = r => describeField(model, r).label;
  const isMeasure = r => tvIsMeasureRef(model, r);
  const panes = sheet.panes.length > 1 && sheet.panes.some(p => p.id) ? sheet.panes.filter(p => p.id) : sheet.panes;
  // an Automatic card draws the All card's mark, else Tableau's automatic mark for the shelves
  const allPane = sheet.panes.find(p => !p.id);
  const inherit = (allPane && markOf(allPane)) || tvAutomaticMark(visualShelfShape(sheet, model));
  const marks = [...new Set(panes.map(p => markOf(p) || inherit))];
  const enc = ch => {
    const refs = [];
    panes.forEach(p => p.encodings.filter(e => e.channel === ch).forEach(e => { if (!refs.some(r => tfSameField(r, e.field))) refs.push(e.field); }));
    return refs;
  };
  const color = enc("color"), size = enc("size"), shape = enc("shape"), detail = enc("lod"), path = enc("path");
  const text = [...enc("text"), ...enc("label")];
  panes.forEach(p => p.labelRuns.forEach(r => r.refs.forEach(x => { if (!text.some(t => tfSameField(t, x))) text.push(x); })));
  const shelf = name => sheet[name].filter(r => r.name !== "__tableau_internal_object_id__");
  out.encodings = {
    rows: shelf("rows").map(label), columns: shelf("cols").map(label), color: color.map(label), size: size.map(label),
    shape: shape.map(label), label: text.map(label), detail: detail.map(label), path: path.map(label),
    angle: enc("wedge-size").map(label), tooltip: enc("tooltip").map(label)
  };
  Object.keys(out.encodings).forEach(k => { if (!out.encodings[k].length) delete out.encodings[k]; });

  // calculated fields behind the view: Tableau computed them – the export writes their values
  const seen = new Set();
  [...sheet.rows, ...sheet.cols, ...panes.flatMap(p => p.encodings.map(e => e.field))].forEach(r => {
    const d = describeField(model, r);
    if (d.calcs.length && !seen.has(d.label)) { seen.add(d.label); out.calculations.push({ name: d.label, kind: "quick table calculation" }); }
    if (d.info && d.info.calcKind && !d.info.param && !seen.has(d.name)) {
      seen.add(d.name);
      out.calculations.push({ name: d.info.caption || d.name, kind: d.info.calcKind });
    }
  });

  const rowsM = sheet.rows.filter(isMeasure), colsM = sheet.cols.filter(isMeasure);
  const horizontal = colsM.length > 0 && rowsM.length === 0;
  const valueRefs = horizontal ? colsM : rowsM;
  const catShelf = horizontal ? "rows" : "cols";
  const dimsOf = s => sheet[s].filter(r => !isMeasure(r) && r.name !== "Measure Names" && !GEO_GENERATED.test(r.name) &&
                                        r.name !== "__tableau_internal_object_id__");
  const calcsOf = r => describeField(model, r).calcs;
  const colorDim = color.find(r => !isMeasure(r) && r.name !== "Measure Names");
  const colorMeasure = color.find(r => isMeasure(r));
  const constant = r => { const d = describeField(model, r); return !!d.info && d.info.constant !== undefined; };
  const negated = r => { const d = describeField(model, r); return !!d.info && /^\s*-/.test(d.info.formula || ""); };
  const spaces = (sheet.style && sheet.style.spaces) || [];
  const flags = out.flags;
  // Dual Axis: an axis folded onto the one before it (the TWB's fold); measures without it get panes of their own
  flags.dualAxis = spaces.some(s => s.fold);
  flags.separatePanes = !flags.dualAxis && (rowsM.length > 1 || colsM.length > 1);
  flags.percentOfTotal = valueRefs.some(r => calcsOf(r).includes("pcto"));
  flags.runningTotal = valueRefs.some(r => calcsOf(r).includes("cum") || calcsOf(r).includes("rsum"));
  flags.rank = valueRefs.some(r => calcsOf(r).includes("rank"));
  flags.reversedAxis = spaces.some(s => s.reverse);
  flags.constantAxis = valueRefs.some(constant);
  flags.smallMultiples = dimsOf(horizontal ? "cols" : "rows").length > 0 && valueRefs.length > 0;
  flags.detail = detail.some(r => !isMeasure(r));
  if (flags.dualAxis) evidence.push(`dual axis${spaces.some(s => s.fold && s.synchronized) ? " (synchronized)" : ""}`);
  if (flags.separatePanes) evidence.push("several measures, each in its own pane");
  if (flags.percentOfTotal) evidence.push("percent-of-total table calculation on the value axis");
  if (flags.runningTotal) evidence.push("running-total table calculation on the value axis");
  if (flags.rank) evidence.push("rank table calculation on the value axis");
  if (flags.reversedAxis) evidence.push("reversed axis");
  if (flags.smallMultiples) evidence.push("a dimension on the value shelf: one pane per member (small multiples)");
  if (colorDim) evidence.push(`colour by the dimension ${label(colorDim)}`);
  else if (colorMeasure) evidence.push(`colour by the measure ${label(colorMeasure)}`);
  if (size.length) evidence.push(`size by ${size.map(label).join(", ")}`);
  if (path.length) evidence.push(`path by ${path.map(label).join(", ")}`);
  if (flags.detail) evidence.push(`detail: ${detail.filter(r => !isMeasure(r)).map(label).join(", ")}`);

  const refLines = sheet.referenceLines || [];
  const bulletTarget = refLines.some(rl => rl.scope === "per-cell" && rl.field && rl.axis && !tfSameField(rl.field, rl.axis));
  const dateParts = [...sheet.rows, ...sheet.cols].map(r => String(r.deriv || "").toLowerCase());
  const has = m => marks.includes(m);
  const shelfDims = [...dimsOf("rows"), ...dimsOf("cols")];
  const bars = horizontal ? "Bar" : "Column";
  const name = () => {
    switch (type) {
      case VISUAL_TYPES.MAP: case VISUAL_TYPES.MAP_FILLED: {
        if (visualModel.gauge) { evidence.push("polygon arcs + needle rebuilt from the workbook's calculations"); return "Gauge"; }
        const geoDims = [...detail, ...shelfDims, ...color, ...text].filter(r => !isMeasure(r) && !GEO_GENERATED.test(r.name) &&
                                                                             !!(describeField(model, r).info || {}).geoRole);
        flags.geographic = true;
        if (has("density") || has("heatmap")) return "Density Map";
        if (!geoDims.length && has("pie")) return "Donut / Pie built on map layers";
        if (!geoDims.length) { evidence.push("no geographic field: map layers used as a canvas"); return "Map with Layers"; }
        if (has("pie")) return "Pies on a Map";
        if (panes.length > 1 && marks.length > 1) return "Map with Layers";
        if (has("line") && path.length) return "Flow Map";
        if (has("polygon")) return "Polygon Map";
        if (has("multipolygon") || type === VISUAL_TYPES.MAP_FILLED) return "Filled Map";
        return "Symbol Map";
      }
      case VISUAL_TYPES.PIE:
        if (visualModel.pie && visualModel.pie.gauge) { evidence.push("a slice in the background colour hides part of the dial"); return "Gauge"; }
        if (visualModel.pie && (visualModel.pie.inner || (visualModel.pie.layers || 0) > 1)) return "Nested Donut";
        if ((visualModel.pie && visualModel.pie.hole) || (vm.fmt.pieLayers && (vm.fmt.pieLayers() || {}).hole)) return "Donut Chart";
        return shelfDims.length ? "Multiple Pies" : "Pie Chart";
      case VISUAL_TYPES.HEATMAP:
        if (dateParts.includes("wk") && (dateParts.includes("wd") || dateParts.includes("dy"))) return "Calendar Heat Map";
        return text.some(isMeasure) ? "Highlight Table" : "Heat Map";
      case VISUAL_TYPES.TABLE:
        if (vm.markTable) return "Table built from marks";
        if (has("shape") && dimsOf("rows").length && dimsOf("cols").length) return "KPI Indicator Matrix";
        if (has("circle") && size.length && shelfDims.some(r => r.type === "qk")) return "Circle Timeline";
        return dimsOf("rows").length && dimsOf("cols").length ? "Crosstab" : "Text Table";
      case VISUAL_TYPES.KPI: return "KPI (BAN)";
      case VISUAL_TYPES.HISTOGRAM: return "Histogram";
      case VISUAL_TYPES.WATERFALL: return "Waterfall Chart";
      case VISUAL_TYPES.BOXPLOT: return "Box Plot";
      case VISUAL_TYPES.GANTT: return "Gantt Chart";
      case VISUAL_TYPES.TREEMAP: return "Treemap";
      case VISUAL_TYPES.BUBBLE: return "Packed Bubbles";
      case VISUAL_TYPES.CUSTOM: return "Viz Extension";
      case VISUAL_TYPES.SCATTER:
        if (size.length) return "Scatter Plot with Size (bubble)";
        if (refLines.length >= 2 || (colorMeasure || colorDim) && refLines.length) return "Quadrant Chart";
        return "Scatter Plot";
      case VISUAL_TYPES.COMBO:
        if (valueRefs.some(r => calcsOf(r).includes("pcto") && calcsOf(r).includes("cum"))) return "Pareto Chart";
        return `Combination Chart (${marks.filter(Boolean).join(" + ")}, dual axis)`;
      case VISUAL_TYPES.AREA:
        if (flags.percentOfTotal && colorDim) return "100% Stacked Area";
        if (colorDim) return "Stacked Area";
        if (flags.smallMultiples) return "Small Multiple Areas";
        return "Area Chart";
      case VISUAL_TYPES.LINE:
        if (path.some(r => r.name === "Measure Names") && valueRefs.some(r => r.name === "Multiple Values")) {
          return has("circle") ? "Barbell (Dumbbell) Chart" : flags.constantAxis || hasConstantMeasure(model, sheet) ? "Rounded Bar Chart" : "Line Chart";
        }
        if (rowsM.length && colsM.length) return flags.runningTotal && flags.percentOfTotal ? "Pareto Curve (cumulative % against cumulative %)" : "Line (measure against measure)";
        if (flags.rank) return "Bump Chart";
        if (!has("line") && marks.some(m => /^(circle|shape|square)$/.test(m))) return "Dot Plot";
        if (valueRefs.length && valueRefs.every(r => r.name === valueRefs[0].name) && valueRefs.length === 2 && has("circle") &&
            sheet[catShelf].some(r => /^(yr|tyr)$/i.test(r.deriv || "") && r.type !== "qk")) return "Slope Chart";
        if (flags.runningTotal) return "Cumulative Line Chart";
        if (flags.smallMultiples) return "Small Multiple Lines (sparklines)";
        if (flags.dualAxis && valueRefs.length > 1 && !valueRefs.every(r => tfSameField(r, valueRefs[0]))) return "Dual-axis Line Chart";
        if (flags.separatePanes && valueRefs.length > 1) return "Line Charts in separate panes (one per measure)";
        if (colorDim || (flags.detail && has("line"))) return "Multi-line Chart";
        return "Line Chart";
      case VISUAL_TYPES.BAR: {
        if (path.some(r => r.name === "Measure Names")) return "Rounded Bar Chart";
        if (bulletTarget) return "Bullet Graph";
        if (valueRefs.some(negated) || (valueRefs.some(r => r.name === "Multiple Values") && sheet.fieldRefs.some(negated))) return "Butterfly (Diverging) Bar Chart";
        if (flags.percentOfTotal && valueRefs.some(constant)) return "Progress Bar";
        if (has("circle") && has("bar")) return "Lollipop Chart";
        if (flags.percentOfTotal && colorDim) return `100% Stacked ${bars}`;
        if (flags.smallMultiples) return `Small Multiple ${bars}s`;
        if (flags.dualAxis && valueRefs.length > 1 && marks.every(m => m === "bar")) return "Bar in Bar";
        if (flags.dualAxis && valueRefs.length > 1) return `Combination Chart (${marks.filter(Boolean).join(" + ")}, dual axis)`;
        if (flags.separatePanes && valueRefs.length > 1) return `${bars} Charts in separate panes (one per measure)`;
        const catDimRefs = dimsOf(catShelf);
        if (colorDim && !catDimRefs.some(r => tfSameField(r, colorDim)) && !sheet[horizontal ? "cols" : "rows"].some(r => tfSameField(r, colorDim))) return `Stacked ${bars}`;
        // side by side: the colour field is the innermost of several category fields (bars grouped by the outer one)
        if (colorDim && catDimRefs.length > 1 && tfSameField(catDimRefs[catDimRefs.length - 1], colorDim)) return `Side-by-side (grouped) ${bars}`;
        if (colorDim) return `${bars} Chart coloured by ${label(colorDim)}`;
        if (flags.detail) return `Stacked ${bars} (by detail)`;
        if (refLines.length) return `${bars} Chart with Reference Line`;
        if (dateParts.some(d => TV_DATE_DERIVS.has(d))) return `${bars} Chart over Time`;
        return horizontal ? "Horizontal Bar Chart" : "Column Chart";
      }
      default: return titleCase(type);
    }
  };
  out.visual = name();
  return out;
}

/** a measure on the sheet that is a constant (MIN(0), AVG(1)): it only places the marks */
function hasConstantMeasure(model, sheet) {
  return sheet.fieldRefs.some(r => { const d = describeField(model, r); return !!d.info && d.info.constant !== undefined && !d.info.param; });
}

function titleCase(t) { return String(t || "").toLowerCase().replace(/_/g, " ").replace(/\b\w/g, c => c.toUpperCase()); }


/* ═══ visual/strategy.js ══════════════════════════════════════════════════════════════════════════ */
/* How each visual was converted, in the terms of the conversion report:
 *   NATIVE          a genuine Excel equivalent (chart type or formatted cells)
 *   CONSTRUCTED     Excel has no such chart: built from Excel primitives that keep its meaning
 *   APPROXIMATE     an Excel chart that keeps the analytical meaning, with a visible difference (named)
 *   TABLE_FALLBACK  no Excel chart would read correctly: the visual's data as a table
 *   IMAGE_FALLBACK  a picture of the visual (Tableau-rendered), not editable
 *   UNSUPPORTED     not exported
 *   FAILED          exporting it went wrong
 * A correct table always beats a chart that would mislead: anything not provably meaningful falls back. */

const STRATEGY = Object.freeze({
  NATIVE: "NATIVE", CONSTRUCTED: "CONSTRUCTED", APPROXIMATE: "APPROXIMATE", TABLE_FALLBACK: "TABLE_FALLBACK",
  IMAGE_FALLBACK: "IMAGE_FALLBACK", UNSUPPORTED: "UNSUPPORTED", FAILED: "FAILED"
});

/** report order, strongest first */
const STRATEGY_ORDER = ["NATIVE", "CONSTRUCTED", "APPROXIMATE", "TABLE_FALLBACK", "IMAGE_FALLBACK", "UNSUPPORTED", "FAILED"];

const FIDELITY = { NATIVE: "High", CONSTRUCTED: "High", APPROXIMATE: "Medium", TABLE_FALLBACK: "Exact data",
                   IMAGE_FALLBACK: "Picture only", UNSUPPORTED: "None", FAILED: "None" };

/** an Excel chart in words ("Stacked column chart", "Combo chart: column + line, secondary axis") @param {ChartSpec} spec */
function chartOutputName(spec) {
  if (spec.conversion && spec.conversion.output) return spec.conversion.output;
  const dir = spec.barDir === "bar" ? "bar" : "column";
  const grouping = spec.percent ? "100% stacked " : spec.stacked ? "stacked " : "";
  const kinds = [...new Set(spec.series.filter(s => !s.refLine).map(s => s.type || spec.kind))];
  const extras = spec.series.some(s => s.refLine) || (spec.refLines && spec.refLines.length) ? " with reference line" : "";
  switch (spec.kind) {
    case "bar": return cap(`${grouping || "clustered "}${dir} chart`) + extras;
    case "line": return (spec.series.filter(s => !s.refLine).every(s => s.line === false) ? "Line chart (markers only)" : "Line chart") + extras;
    case "area": return cap(`${grouping}area chart`) + extras;
    case "combo": return (kinds.length === 1 && kinds[0] === "bar" ? `Bar-in-bar ${dir} chart (two axes${spec.secondarySync ? ", one scale" : ""})`
                          : `Combo chart: ${kinds.map(k => k === "bar" ? dir : k).join(" + ")}` +
                            (spec.series.some(s => s.secondary) ? `, secondary axis${spec.secondarySync ? " (synchronized)" : ""}` : "")) + extras;
    case "pie": return "Pie chart";
    case "doughnut": return "Doughnut chart";
    case "scatter": return spec.series.some(s => s.line) ? "Scatter chart with lines" : "Scatter chart" + extras;
    case "bubble": return "Bubble chart";
    case "treemap": return "Treemap (Excel 2016+)";
    default: return cap(`${spec.kind} chart`);
  }
}

/**
 * The conversion of one exported visual, once the export has drawn it (renderer, status and chart specs final).
 * @param {VisualModel} visualModel @param {VisualSemantics} [semantics]
 * @returns {Conversion}
 */
function conversionOf(visualModel, semantics) {
  const make = (strategy, excelOutput, reason, fidelity) => ({ strategy, excelOutput, reason, fidelity: fidelity || FIDELITY[strategy] });
  const renderer = visualModel.renderer;
  const why = String(visualModel.statusReason || "");
  if (visualModel.status === "warning" && /injection failed/i.test(why)) return make(STRATEGY.FAILED, "Nothing", why);
  if (visualModel.gauge) return make(STRATEGY.CONSTRUCTED, "Gauge: doughnut dial + pie needle",
    "Excel has no gauge chart: built Excel's standard way – the dial's bands as a doughnut (the rest of the circle see-through), " +
    "the needle as a thin pie slice, both turned to the dial's start; positions recomputed from the workbook's calculations, " +
    "Tableau's text marks as text boxes");
  if (visualModel.pie && visualModel.pie.gauge) return make(STRATEGY.NATIVE, `${visualModel.pie.hole ? "Doughnut" : "Pie"} chart (gauge)`,
    "Direct Excel equivalent: Tableau's slices, colours and pointer slice; the slice Tableau paints in the background colour " +
    "is see-through" + (visualModel.pie.hole ? "; the hole with its text" : ""));
  if (visualModel.pie) {
    const pie = visualModel.pie;
    const donut = pie.hole || pie.inner;
    return make(STRATEGY.NATIVE, `${donut ? "Doughnut" : "Pie"} chart${pie.inner ? " (two rings)" : ""}`,
      "Direct Excel equivalent: Tableau's slices, colours and labels" + (pie.hole ? ", the hole with its text" : "") + "; tooltips as hover text");
  }
  if (renderer === "excel-chart") {
    const specs = visualModel.chartSpecs || [];
    const rank = s => STRATEGY_ORDER.indexOf((s.conversion && s.conversion.strategy) || defaultStrategy(s));
    const weakest = specs.slice().sort((a, b) => rank(b) - rank(a))[0];
    if (!weakest) return make(STRATEGY.FAILED, "Nothing", "no chart was produced");
    const strategy = (weakest.conversion && weakest.conversion.strategy) || defaultStrategy(weakest);
    const output = [...new Set(specs.map(chartOutputName))].join("; ") + (specs.length > 1 ? ` (${specs.length} panes)` : "");
    const note = specs.map(s => s.conversion && s.conversion.note).filter(Boolean)[0];
    return make(strategy, output, note || (strategy === STRATEGY.APPROXIMATE ? "Excel's closest chart; the layout differs" : "Direct Excel equivalent"),
                weakest.conversion && weakest.conversion.fidelity);
  }
  if (renderer === "tableau-image") return make(STRATEGY.IMAGE_FALLBACK, "Picture (Tableau-rendered) + its data on the Visual Data sheet",
                                                  why || "no editable Excel equivalent");
  if (renderer === "data-fallback" && visualModel.crosstab) {
    return make(STRATEGY.CONSTRUCTED, "Table (crosstab) with the chart's Tableau colours",
                "Chosen in the panel: Charts as tables – every mark as a row, coloured like the chart", "Exact data");
  }
  if (renderer === "data-fallback") return make(STRATEGY.TABLE_FALLBACK, "Data table", fallbackReason(visualModel, semantics, why));
  if (renderer === "cell") {
    const vm = visualModel.viewModel;
    if (vm && vm.lossless === false) return make(STRATEGY.TABLE_FALLBACK, "Data table", vm.fallbackReason || "its marks encode values the cells would not show");
    if (vm && vm.matrix) return make(STRATEGY.CONSTRUCTED, "Matrix of cells with Tableau's colour per cell",
                                     "Excel has no heat-map chart: a matrix with the same rows, columns and colours");
    if (vm && vm.markTable) return make(STRATEGY.CONSTRUCTED, "Cells drawn like Tableau's marks (fills, arrows, data bars)",
                                        "a table built from marks: each column styled like its marks");
    if (visualModel.type === VISUAL_TYPES.KPI) return make(STRATEGY.NATIVE, "KPI cells", "Direct Excel equivalent: the values in Tableau's format");
    if (visualModel.type === VISUAL_TYPES.HEATMAP) return make(STRATEGY.CONSTRUCTED, "Table with Tableau's colour per row", "Excel has no heat-map chart");
    return make(STRATEGY.NATIVE, "Formatted cell table", "Direct Excel equivalent: Tableau's table as cells");
  }
  return make(STRATEGY.UNSUPPORTED, "Nothing", why || "no renderer");
}

/** @param {ChartSpec} spec */
function defaultStrategy(spec) {
  if (spec.kind === "bubble" && spec.packed) return STRATEGY.APPROXIMATE;
  return STRATEGY.NATIVE;
}

/** why a visual became a data table, in words a reader of the report understands */
function fallbackReason(visualModel, semantics, why) {
  const t = visualModel.type;
  const visual = semantics ? semantics.visual : "";
  if (t === VISUAL_TYPES.MAP || t === VISUAL_TYPES.MAP_FILLED) {
    const kind = {
      "Density Map": "density marks have no Excel equivalent",
      "Polygon Map": "custom polygons (hex bins, shapes) cannot be drawn by an Excel chart",
      "Flow Map": "paths between locations cannot be drawn on an Excel map",
      "Map with Layers": "several map layers on one map have no Excel equivalent",
      "Pies on a Map": "pies placed on a map have no Excel equivalent",
      "Donut / Pie built on map layers": "a donut built from map layers (no geography) – Excel's charts cannot stack the layers",
      "Filled Map": "Excel's filled map cannot reproduce Tableau's geography and colour scale reliably",
      "Symbol Map": "Excel has no symbol map: marks sized and coloured at locations"
    }[visual] || "Excel cannot reproduce this map";
    return `${kind}; every mark exported with its location fields and values`;
  }
  if (t === VISUAL_TYPES.CUSTOM) return "a viz extension draws this sheet; Excel has no equivalent";
  const m = why.match(/native chart not possible \(([^)]*)\)/i);
  if (m) return `no faithful Excel chart: ${m[1]}`;
  return why.replace(/ – exported as data$/, "") || "no meaningful Excel equivalent";
}

function cap(s) { return s.charAt(0).toUpperCase() + s.slice(1); }


/* ═══ export/data-sheet.js ════════════════════════════════════════════════════════════════════════ */
/* The data behind visuals exported as pictures (IMAGE_FALLBACK): each one's marks as a plain table on a "Visual Data"
 * sheet, so the numbers travel with the picture (which links to them). */

const DATA_SHEET = "Visual Data";

/**
 * Appends one visual's data (every column but the latitude / longitude Tableau generates) under a title.
 * @param {import("exceljs").Workbook} workbook @param {string} title @param {ViewModel} vm
 * @returns {string} the link to the block ("#'Visual Data'!A5")
 */
function appendVisualData(workbook, title, vm) {
  const ws = workbook.getWorksheet(DATA_SHEET) || workbook.addWorksheet(DATA_SHEET);
  const font = { name: "Arial", size: 9 };
  const start = ws.rowCount ? ws.rowCount + 2 : 1;
  const cols = vm.cols.map((c, i) => i).filter(i => !/^(latitude|longitude) \(generated\)$/i.test(vm.cols[i].name));
  ws.getCell(start, 1).value = title;
  ws.getCell(start, 1).font = { ...font, size: 11, bold: true };
  cols.forEach((ci, k) => {
    const cell = ws.getCell(start + 1, k + 1);
    cell.value = vm.cols[ci].label || vm.cols[ci].name;
    cell.font = { ...font, bold: true };
    cell.border = { bottom: { style: "thin", color: { argb: "FFD4D4D4" } } };
    ws.getColumn(k + 1).width = Math.max(ws.getColumn(k + 1).width || 0, Math.min(40, String(cell.value).length + 4));
  });
  vm.rows.forEach((row, r) => cols.forEach((ci, k) => {
    const dv = row[ci], cell = ws.getCell(start + 2 + r, k + 1);
    const v = tfIsNull(dv) ? null : dv.nativeValue !== undefined ? dv.nativeValue : dv.value;
    cell.value = typeof v === "number" ? v : tfIsNull(dv) ? null : tfDvText(dv);
    if (typeof v === "number") { const nf = inferExcelNumFmt(dv.formattedValue, v); if (nf) cell.numFmt = nf; }   // as Tableau shows it
    cell.font = font;
  }));
  return `#'${DATA_SHEET}'!A${start}`;
}


/* ═══ export/report.js ════════════════════════════════════════════════════════════════════════════ */
/* The conversion report: a sheet after the dashboard that lists every visual – what Tableau shows, how Excel shows
 * it (NATIVE … TABLE_FALLBACK), the fidelity and why – then the dashboard's other objects, the filters and
 * parameters the data reflects, and what was left out. Nothing is converted silently. */

const REPORT_SHEET = "Conversion Report";

const SLICER_REPORT_NOTE = "Slicers: the visuals that follow them compute their numbers with formulas over the data source's rows " +
  `(the hidden "Slicer Data" sheet), so a click in a slicer filters them all; the others keep Tableau's values. Charts and ` +
  "tables keep the bars, points and rows Tableau drew: a value the filter left out at export adds to those, but a category, " +
  "row or line only it has does not appear (export with every value selected to have them all).";

const STRATEGY_FILL = { NATIVE: "FFE2F0D9", CONSTRUCTED: "FFDDEBF7", APPROXIMATE: "FFFFF2CC", TABLE_FALLBACK: "FFEDEDED",
                        IMAGE_FALLBACK: "FFE4DFEC", UNSUPPORTED: "FFF8CBAD", FAILED: "FFF4B183" };

/** "15 NATIVE · 3 CONSTRUCTED · 2 TABLE_FALLBACK" @param {{ strategy: string }[]} visuals */
function strategyCounts(visuals) {
  return STRATEGY_ORDER.map(s => [s, visuals.filter(v => v.strategy === s).length]).filter(([, n]) => n);
}

/**
 * @param {import("exceljs").Workbook} workbook
 * @param {{ dashboard: string, workbookFile: string | null, visuals: ReportVisual[],
 *           objects: { name: string, kind: string, output: string }[], skipped: { name: string, reason: string }[],
 *           filters: { name: string, values: string[] }[], parameters: { name: string, value: string }[] }} report
 */
function writeConversionReport(workbook, report) {
  const names = workbook.worksheets.map(w => w.name);
  const ws = workbook.addWorksheet(names.includes(REPORT_SHEET) ? REPORT_SHEET + " (export)" : REPORT_SHEET);
  ws.views = [{ state: "frozen", ySplit: 5, showGridLines: false }];
  const font = { name: "Arial", size: 9 };
  const bold = { ...font, bold: true };
  let r = 1;
  const put = (row, col, value, f = font, extra = {}) => {
    const cell = ws.getCell(row, col);
    cell.value = value;
    cell.font = f;
    cell.alignment = { vertical: "top", wrapText: true, ...(extra.alignment || {}) };
    if (extra.fill) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: extra.fill } };
    if (extra.border) cell.border = extra.border;
    return cell;
  };
  put(r, 1, "Tableau → Excel conversion report", { ...font, size: 14, bold: true }, { alignment: { wrapText: false } });
  ws.mergeCells(r, 1, r, 6);
  r++;
  put(r, 1, `Dashboard: ${report.dashboard}` + (report.workbookFile ? `   ·   Workbook: ${report.workbookFile}` : "   ·   no workbook loaded") +
           `   ·   Export to Excel v${EXTENSION_VERSION} (${EXTENSION_BUILD})`,
      font, { alignment: { wrapText: false } });
  ws.mergeCells(r, 1, r, 6);
  r++;
  const counts = strategyCounts(report.visuals);
  put(r, 1, `${report.visuals.length} visual${report.visuals.length === 1 ? "" : "s"}: ` +
      (counts.map(([s, n]) => `${n} ${s}`).join("   ·   ") || "none"), bold, { alignment: { wrapText: false } });
  ws.mergeCells(r, 1, r, 6);
  r += 2;

  const head = ["#", "Worksheet", "Tableau visual", "Strategy", "Excel output", "Fidelity", "Reason", "Detected from", "Encodings", "Calculations"];
  const widths = [4, 24, 24, 16, 30, 11, 52, 52, 48, 34];
  const line = { style: "thin", color: { argb: "FFD4D4D4" } };
  head.forEach((h, i) => put(r, i + 1, h, bold, { fill: "FFF2F2F2", border: { bottom: line } }));
  widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  const headerRow = r;
  report.visuals.forEach((v, k) => {
    r++;
    const enc = Object.entries(v.semantics.encodings || {}).map(([ch, list]) => `${ch}: ${list.join(", ")}`).join("\n");
    const calcs = (v.semantics.calculations || []).map(c => `${c.name} (${c.kind})`).join("\n");
    const cells = [k + 1, v.worksheet, v.semantics.visual, v.strategy, v.excelOutput, v.fidelity, v.reason,
                   `${v.semantics.evidence.join("; ")} [confidence: ${v.semantics.confidence}]`, enc, calcs];
    cells.forEach((value, i) => put(r, i + 1, value, i === 3 ? bold : font,
      { fill: i === 3 ? STRATEGY_FILL[v.strategy] : undefined, border: { bottom: line } }));
  });
  ws.autoFilter = { from: { row: headerRow, column: 1 }, to: { row: r, column: head.length } };

  const section = title => { r += 2; put(r, 1, title, { ...bold, size: 11 }, { alignment: { wrapText: false } }); ws.mergeCells(r, 1, r, 6); };
  if (report.objects.length) {
    section("Other dashboard objects");
    report.objects.forEach(o => { r++; put(r, 2, o.name); put(r, 3, o.kind); put(r, 5, o.output); });
  }
  if (report.slicers && report.slicers.length) {
    section("Slicers: the visuals they filter");
    report.slicers.forEach(s => { r++; put(r, 2, s.worksheet); put(r, 3, s.follows ? "Follows the slicers" : "Fixed values"); put(r, 5, s.note); ws.mergeCells(r, 5, r, 7); });
  }
  if (report.filters.length || report.parameters.length) {
    section("Filters and parameters (the exported data reflects their state at export time)");
    report.filters.forEach(f => { r++; put(r, 2, f.name); put(r, 3, "Filter"); put(r, 5, f.values.join(", ")); });
    report.parameters.forEach(p => { r++; put(r, 2, p.name); put(r, 3, "Parameter"); put(r, 5, p.value); });
  }
  if (report.skipped.length) {
    section("Not exported");
    report.skipped.forEach(s => { r++; put(r, 2, s.name); put(r, 5, s.reason); ws.mergeCells(r, 5, r, 7); });
  }
  section("Notes");
  [
    ...(report.slicers && report.slicers.length ? [SLICER_REPORT_NOTE] : []),
    "Values are Tableau's own results for the view: calculated fields, LOD expressions and table calculations are exported as values, not translated into Excel formulas.",
    "Filters, parameters and sets apply as they were at export time; Excel has no live link back to Tableau.",
    "Tooltips: pies, donuts and packed bubbles show Tableau's tooltip text on hover; elsewhere the tooltip fields are in the exported data only.",
    "Actions, highlighting and drill-down are interactive Tableau features and are not reproduced.",
    "TABLE_FALLBACK means no Excel chart would show the visual correctly: its data is exported as a table instead (the cell carries a note)."
  ].forEach(t => { r++; put(r, 2, t, font, { alignment: { wrapText: true } }); ws.mergeCells(r, 2, r, 7); ws.getRow(r).height = 24; });
  return ws;
}

/**
 * Several dashboards in one file: one report sheet with a section per dashboard (its visuals, other objects,
 * filters / parameters, what was left out), then the notes once. One dashboard: exactly writeConversionReport.
 * @param {import("exceljs").Workbook} workbook
 * @param {Parameters<typeof writeConversionReport>[1][]} reports
 */
function writeConversionReports(workbook, reports) {
  if (reports.length === 1) return writeConversionReport(workbook, reports[0]);
  const names = workbook.worksheets.map(w => w.name);
  const ws = workbook.addWorksheet(names.includes(REPORT_SHEET) ? REPORT_SHEET + " (export)" : REPORT_SHEET);
  ws.views = [{ showGridLines: false }];
  const font = { name: "Arial", size: 9 };
  const bold = { ...font, bold: true };
  const line = { style: "thin", color: { argb: "FFD4D4D4" } };
  let r = 1;
  const put = (row, col, value, f = font, extra = {}) => {
    const cell = ws.getCell(row, col);
    cell.value = value;
    cell.font = f;
    cell.alignment = { vertical: "top", wrapText: true, ...(extra.alignment || {}) };
    if (extra.fill) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: extra.fill } };
    if (extra.border) cell.border = extra.border;
    return cell;
  };
  const banner = (text, f) => { put(r, 1, text, f, { alignment: { wrapText: false } }); ws.mergeCells(r, 1, r, 6); };
  const all = reports.flatMap(rep => rep.visuals);
  banner("Tableau → Excel conversion report", { ...font, size: 14, bold: true });
  r++;
  banner(`${reports.length} dashboards` + (reports[0].workbookFile ? `   ·   Workbook: ${reports[0].workbookFile}` : "   ·   no workbook loaded") +
         `   ·   Export to Excel v${EXTENSION_VERSION} (${EXTENSION_BUILD})`, font);
  r++;
  banner(`${all.length} visual${all.length === 1 ? "" : "s"}: ` + (strategyCounts(all).map(([s, n]) => `${n} ${s}`).join("   ·   ") || "none"), bold);

  const head = ["#", "Worksheet", "Tableau visual", "Strategy", "Excel output", "Fidelity", "Reason", "Detected from", "Encodings", "Calculations"];
  const widths = [4, 24, 24, 16, 30, 11, 52, 52, 48, 34];
  widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  const section = title => { r += 2; banner(title, { ...bold, size: 11 }); };

  reports.forEach(report => {
    r += 2;
    const counts = strategyCounts(report.visuals);
    banner(`Dashboard: ${report.dashboard}`, { ...bold, size: 12 });
    r++;
    banner(`${report.visuals.length} visual${report.visuals.length === 1 ? "" : "s"}: ` +
      (counts.map(([s, n]) => `${n} ${s}`).join("   ·   ") || "none"), font);
    r++;
    head.forEach((h, i) => put(r, i + 1, h, bold, { fill: "FFF2F2F2", border: { bottom: line } }));
    report.visuals.forEach((v, k) => {
      r++;
      const enc = Object.entries(v.semantics.encodings || {}).map(([ch, list]) => `${ch}: ${list.join(", ")}`).join("\n");
      const calcs = (v.semantics.calculations || []).map(c => `${c.name} (${c.kind})`).join("\n");
      const cells = [k + 1, v.worksheet, v.semantics.visual, v.strategy, v.excelOutput, v.fidelity, v.reason,
                     `${v.semantics.evidence.join("; ")} [confidence: ${v.semantics.confidence}]`, enc, calcs];
      cells.forEach((value, i) => put(r, i + 1, value, i === 3 ? bold : font,
        { fill: i === 3 ? STRATEGY_FILL[v.strategy] : undefined, border: { bottom: line } }));
    });
    if (report.objects.length) {
      section("Other dashboard objects");
      report.objects.forEach(o => { r++; put(r, 2, o.name); put(r, 3, o.kind); put(r, 5, o.output); });
    }
    if (report.slicers && report.slicers.length) {
      section("Slicers: the visuals they filter");
      report.slicers.forEach(s => { r++; put(r, 2, s.worksheet); put(r, 3, s.follows ? "Follows the slicers" : "Fixed values"); put(r, 5, s.note); ws.mergeCells(r, 5, r, 7); });
    }
    if (report.filters.length || report.parameters.length) {
      section("Filters and parameters (the exported data reflects their state at export time)");
      report.filters.forEach(f => { r++; put(r, 2, f.name); put(r, 3, "Filter"); put(r, 5, f.values.join(", ")); });
      report.parameters.forEach(p => { r++; put(r, 2, p.name); put(r, 3, "Parameter"); put(r, 5, p.value); });
    }
    if (report.skipped.length) {
      section("Not exported");
      report.skipped.forEach(s => { r++; put(r, 2, s.name); put(r, 5, s.reason); ws.mergeCells(r, 5, r, 7); });
    }
  });
  section("Notes");
  [
    ...(reports.some(rep => rep.slicers && rep.slicers.length) ? [SLICER_REPORT_NOTE] : []),
    "Values are Tableau's own results for the view: calculated fields, LOD expressions and table calculations are exported as values, not translated into Excel formulas.",
    "Filters, parameters and sets apply as they were at export time; Excel has no live link back to Tableau.",
    "Tooltips: pies, donuts and packed bubbles show Tableau's tooltip text on hover; elsewhere the tooltip fields are in the exported data only.",
    "Actions, highlighting and drill-down are interactive Tableau features and are not reproduced.",
    "TABLE_FALLBACK means no Excel chart would show the visual correctly: its data is exported as a table instead (the cell carries a note)."
  ].forEach(t => { r++; put(r, 2, t, font, { alignment: { wrapText: true } }); ws.mergeCells(r, 2, r, 7); ws.getRow(r).height = 24; });
  return ws;
}


/* ═══ charts/writer/chart-data.js ═════════════════════════════════════════════════════════════════ */
/* Chart source data → worksheet cells (charts stay editable and linked). */

/**
 * @param {import("exceljs").Worksheet} ws the (hidden) chart data sheet
 * @param {ChartSpec} spec
 * @param {number} startRow 0-based first row to write
 * @returns {ChartRefs}
 */
/* share of a chart's height Excel gives category names turned on their side */
const TURNED_LABEL_SHARE = 0.45;

/** a name cut to at most room characters: each word shortened alike ("Marketing Associate" → "Mark. Asso."), so names
 * that share a first word stay apart; a word left whole when it fits, "…" when even that is too long */
function shortName(name, room) {
  if (name.length <= room) return name;
  const words = name.trim().split(/\s+/);
  for (let per = Math.max(...words.map(w => w.length)); per >= 1; per--) {
    const t = words.map(w => (w.length > per + 1 ? w.slice(0, per) + "." : w)).join(" ");
    if (t.length <= room) return t;
  }
  return name.slice(0, Math.max(1, room - 1)) + "…";
}

function writeChartData(ws, spec, startRow) {
  const sheet = ws.name;
  const put = (r, c, v, fmt) => {
    const cell = ws.getCell(r + 1, c + 1);
    cell.value = v;
    if (fmt) cell.numFmt = fmt;
  };
  // a value that follows the slicers: a formula over the slicer table's visible rows, Tableau's number as its result
  const valueAt = (values, i) => {
    const f = spec.valueFormula ? spec.valueFormula(values, i) : null;
    return f ? { formula: f, result: values[i] } : values[i];
  };
  ws.getCell(startRow + 1, 1).font = { bold: true };

  if (spec.kind === "scatter" || spec.kind === "bubble") {
    const per = spec.kind === "bubble" ? 3 : 2;                 // x, y (, bubble size) columns per series
    let maxLen = 0;
    /** @type {ChartRefs["series"]} */
    const series = spec.series.map((s, k) => {
      const cx = per * k, cy = cx + 1, n = s.x.length;
      put(startRow, cx, (spec.xTitle || "X") + (spec.series.length > 1 ? " – " + s.name : ""));
      put(startRow, cy, s.name);
      for (let i = 0; i < n; i++) {
        if (num(s.x[i]) !== null) put(startRow + 1 + i, cx, valueAt(s.x, i), spec.xNumFmt);
        if (num(s.y[i]) !== null) put(startRow + 1 + i, cy, valueAt(s.y, i), spec.numFmt);
      }
      maxLen = Math.max(maxLen, n);
      /** @type {ChartRefs["series"][number]} */
      const ref = { tx: cellRef(sheet, cy, startRow),
                    x: cellRef(sheet, cx, startRow + 1, cx, startRow + Math.max(1, n)),
                    y: cellRef(sheet, cy, startRow + 1, cy, startRow + Math.max(1, n)) };
      if (per === 3) {
        const cs = cx + 2;
        put(startRow, cs, (spec.sizeTitle || "Size") + (spec.series.length > 1 ? " – " + s.name : ""));
        s.size.forEach((v, i) => { if (num(v) !== null) put(startRow + 1 + i, cs, valueAt(s.size, i), spec.sizeNumFmt); });
        ref.size = cellRef(sheet, cs, startRow + 1, cs, startRow + Math.max(1, n));
      }
      return ref;
    });
    const lblCol = per * spec.series.length;
    spec.series.forEach((s, k) => {
      if (!s.labelTexts) return;
      const c = lblCol + k;
      put(startRow, c, s.name + " – label");
      s.labelTexts.forEach((v, i) => put(startRow + 1 + i, c, v));
      series[k].lbl = cellRef(sheet, c, startRow + 1, c, startRow + Math.max(1, s.labelTexts.length));
    });
    return { series, nextRow: startRow + maxLen + 3 };
  }

  // a group's name over a hidden level (Job Role over its races): Excel wraps it at spaces within the group's width,
  // but a word longer than that runs into the next one – cut short with "…", as Tableau cuts its header
  const blanks = spec.categories.blankLevels || [];
  const groups = levelStarts(spec.categories.levels);           // from the full names: two cut alike stay two groups
  spec.categories.starts = groups;
  if (blanks.length && spec.widthPx) {
    const lv0 = spec.categories.levels, n = lv0[0].length;
    const charPx = (spec.font.size || 9) * 4 / 3 * 0.5, slot = spec.widthPx * 0.9 / Math.max(1, n);   // Arial's average character
    spec.categories.levels = lv0.map((lv, l) => {
      if (blanks.includes(l) || l === lv0.length - 1) return lv;
      const starts = groups[l];
      return lv.map((name, i) => {
        if (!starts[i]) return name;
        let size = 1;
        while (i + size < n && !starts[i + size]) size++;
        const room = Math.max(1, Math.floor((slot * size - 4) / charPx));
        return String(name).split(/\s+/).some(w => w.length > room) ? String(name).slice(0, Math.max(1, room - 1)).trimEnd() + "…" : name;
      });
    });
  }
  // names turned on their side (categoryLabelFit) under a short chart: Excel leaves them about half its height and cuts
  // the rest down to a letter ("A.") – shortened here to what fits, a little of every word ("Mark. Asso."), as
  // Tableau's scrolling view shows them in full
  if (spec.widthPx && spec.heightPx && spec.categories.levels.length === 1 && !spec.dateAxis && !spec.keepNames) {
    const fit = categoryLabelFit(spec, { w: spec.widthPx * 0.85, h: spec.heightPx * 0.72 });
    if (fit && fit.rot) {
      const fontPx = (spec.font.size || 9) * 4 / 3, slot = spec.widthPx * 0.85 / Math.max(1, spec.categories.levels[0].length) * fit.skip;
      const lines = Math.max(1, Math.floor(slot / (fontPx * 1.25)));          // a wide slot: Excel wraps a turned name
      const room = Math.max(4, Math.floor((spec.heightPx * TURNED_LABEL_SHARE - 6) / (fontPx * 0.55))) * lines;
      spec.categories.levels = [spec.categories.levels[0].map(n => shortName(String(n ?? ""), room))];
    }
  }
  const levels = spec.categories.levels;
  const L = levels.length;
  const N = L ? levels[0].length : 0;
  // treemap: every level written on every row – Excel reads each row as one full hierarchy path
  const starts = spec.kind === "treemap" ? levels.map(lv => lv.map(() => true)) : groups;
  spec.categories.names.forEach((name, l) => put(startRow, l, name || ""));
  spec.series.forEach((s, k) => put(startRow, L + k, s.name));
  // a date axis reads real dates
  const dates = spec.dateAxis ? spec.dateAxis.serials : null;
  const dateFmt = spec.dateAxis && { years: "yyyy", months: "mmm yyyy", days: "d mmm yyyy" }[spec.dateAxis.base];
  // cells computed from other cells of the chart (a waterfall following the slicers): series with cellFormula, label
  // texts with labelFormula, and helper columns (spec.helpers) after the labels; at(kind, key, i) → that cell's
  // address – "series" k, "helper" name
  const lblCount = spec.series.filter(s => s.labelTexts).length;
  const helpers = spec.helpers || [];
  const helperCol = name => L + spec.series.length + lblCount + helpers.findIndex(h => h.name === name);
  const at = (kind, key, i) => `${getExcelColumnName(kind === "series" ? L + /** @type {number} */ (key) : helperCol(String(key)))}${startRow + 2 + i}`;
  for (let i = 0; i < N; i++) {
    if (dates) put(startRow + 1 + i, 0, dates[i], dateFmt);
    else for (let l = 0; l < L; l++) if (l === L - 1 || starts[l][i]) put(startRow + 1 + i, l, levels[l][i]);
    spec.series.forEach((s, k) => {
      const fmt = s.secondary ? spec.secondaryNumFmt : spec.numFmt;
      const f = s.cellFormula ? s.cellFormula(i, at) : null;
      if (f) put(startRow + 1 + i, L + k, num(s.values[i]) === null ? { formula: f } : { formula: f, result: s.values[i] }, fmt);
      else if (num(s.values[i]) !== null) put(startRow + 1 + i, L + k, valueAt(s.values, i), fmt);
    });
  }
  // label texts taken from cells (Excel's "Value From Cells"), after the series columns
  let lblCol = L + spec.series.length;
  const lbl = spec.series.map(s => {
    if (!s.labelTexts) return null;
    const c = lblCol++;
    put(startRow, c, s.name + " – label");
    s.labelTexts.forEach((t, i) => {
      const f = s.labelFormula ? s.labelFormula(i, at) : null;
      put(startRow + 1 + i, c, f ? { formula: f, result: t } : t);
    });
    return cellRef(sheet, c, startRow + 1, c, startRow + Math.max(1, N));
  });
  helpers.forEach(h => {
    const c = helperCol(h.name);
    put(startRow, c, h.name);
    for (let i = 0; i < N; i++) { const f = h.formula(i, at); if (f) put(startRow + 1 + i, c, { formula: f, result: h.results[i] }, spec.numFmt); }
  });
  return {
    cat: cellRef(sheet, 0, startRow + 1, L - 1, startRow + N),
    series: spec.series.map((s, k) => ({
      tx: cellRef(sheet, L + k, startRow),
      val: cellRef(sheet, L + k, startRow + 1, L + k, startRow + N),
      ...(lbl[k] ? { lbl: lbl[k] } : {})
    })),
    nextRow: startRow + N + 3
  };
}


/* ═══ charts/writer/drawingml.js ══════════════════════════════════════════════════════════════════ */
/* DrawingML chart XML (bar / line / area / pie / scatter / bubble / combo). */

function solid(color, alpha) {
  if (!color) return "<a:noFill/>";
  return `<a:solidFill><a:srgbClr val="${hex(color)}">${alpha ? `<a:alpha val="${alpha}"/>` : ""}</a:srgbClr></a:solidFill>`;
}

function line(color, w) { return color ? `<a:ln w="${w || 9525}">${solid(color)}</a:ln>` : `<a:ln><a:noFill/></a:ln>`; }

function runProps(font, o, tag) {
  const sz = Math.round((o.size || font.size || 9) * 100);
  return `<a:${tag} lang="en-US" sz="${sz}" b="${o.bold ? 1 : 0}">${solid(o.color || font.color)}` +
    `<a:latin typeface="${esc(font.name)}"/><a:cs typeface="${esc(font.name)}"/></a:${tag}>`;
}

function txPr(font, o = {}) {
  const def = runProps(font, o, "defRPr").replace(' lang="en-US"', "");
  // rot 0 is written too: it keeps Excel from turning labels it finds crowded
  return `<c:txPr><a:bodyPr${o.rot !== undefined ? ` rot="${o.rot}" vert="horz"` : ""}${o.noWrap ? ' wrap="none"' : ""}/><a:lstStyle/>` +
    `<a:p><a:pPr>${def}</a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>`;
}

/** @param {string} text @param {any} font @param {boolean} vertical reads upwards @param {boolean} [flat] level even where
 * Excel would turn it (a value axis title) */
function title(text, font, vertical, flat = false) {
  if (!text) return "";
  const def = runProps(font, {}, "defRPr").replace(' lang="en-US"', "");
  return `<c:title><c:tx><c:rich><a:bodyPr${vertical ? ' rot="-5400000" vert="horz"' : flat ? ' rot="0" vert="horz"' : ""}/><a:lstStyle/>` +
    `<a:p><a:pPr>${def}</a:pPr><a:r>${runProps(font, {}, "rPr")}<a:t>${esc(text)}</a:t></a:r></a:p></c:rich></c:tx>` +
    `<c:overlay val="0"/></c:title>`;
}

function strCache(values) {
  return `<c:ptCount val="${values.length}"/>` +
    values.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join("");
}

function numCache(values) {
  return `<c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>` +
    values.map((v, i) => num(v) === null ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("");
}

function serTx(ref, name) {
  return `<c:tx><c:strRef><c:f>${esc(ref)}</c:f><c:strCache>${strCache([name])}</c:strCache></c:strRef></c:tx>`;
}

/** @param {string} ref @param {any[][]} levels @param {{ serials: number[] } | null} [dates] a date axis: the dates as serials */
/** groups: each level's group starts (writeChartData's, from the full names), else worked out from levels */
function catXml(ref, levels, dates = null, groups = null) {
  if (dates) {
    return `<c:cat><c:numRef><c:f>${esc(ref)}</c:f><c:numCache><c:formatCode>yyyy\\-mm\\-dd</c:formatCode><c:ptCount val="${dates.serials.length}"/>` +
      dates.serials.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("") + `</c:numCache></c:numRef></c:cat>`;
  }
  if (levels.length <= 1) {
    return `<c:cat><c:strRef><c:f>${esc(ref)}</c:f><c:strCache>${strCache(levels[0] || [])}</c:strCache></c:strRef></c:cat>`;
  }
  const starts = groups || levelStarts(levels);
  const n = levels[0].length;
  // innermost level first, as Excel writes it
  const lvls = levels.map((lv, l) => `<c:lvl>${lv.map((v, i) =>
      (l === levels.length - 1 || starts[l][i]) ? `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>` : "").join("")}</c:lvl>`).reverse();
  return `<c:cat><c:multiLvlStrRef><c:f>${esc(ref)}</c:f><c:multiLvlStrCache><c:ptCount val="${n}"/>${lvls.join("")}` +
    `</c:multiLvlStrCache></c:multiLvlStrRef></c:cat>`;
}

function valXml(tag, ref, values) {
  return `<c:${tag}><c:numRef><c:f>${esc(ref)}</c:f><c:numCache>${numCache(values)}</c:numCache></c:numRef></c:${tag}>`;
}

/** mark labels: the worksheet's label font (else the chart font); labels set from the workbook never wrap
 * @param {ChartSpec} spec */
function labelTxPr(spec) {
  const lf = spec.labelFont;
  if (!lf) return txPr(spec.font, spec.labelPos || spec.series.some(s => s.labelOnly) ? { noWrap: true } : {});
  return txPr({ name: lf.name || spec.font.name, size: lf.size || spec.font.size, color: lf.color || spec.font.color },
              { bold: lf.bold, noWrap: true });
}

/** @param {ChartSpec} spec @param {ChartSeries} s @param {string | null} pos @param {number[]} [hidden] points without a label */
function dLbls(spec, s, pos, hidden) {
  if (!s.labels) return "";
  const only = s.labelOnly ? new Set(s.labelOnly) : null;
  if (only) hidden = [...new Set([...(hidden || []), ...(s.values || []).map((_, i) => i).filter(i => !only.has(i))])].sort((a, b) => a - b);
  const drop = (hidden || []).map(i => `<c:dLbl><c:idx val="${i}"/><c:delete val="1"/></c:dLbl>`).join("");
  if (s.labelTexts) {
    // a doughnut takes no label position (Excel refuses the file)
    return `<c:dLbls>${drop}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${labelTxPr(spec)}` +
      `${spec.kind === "doughnut" ? "" : `<c:dLblPos val="${pos || "r"}"/>`}<c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/>` +
      `<c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>` +
      `<c:extLst><c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="${C15}">` +
      `<c15:showDataLabelsRange val="1"/></c:ext></c:extLst></c:dLbls>`;
  }
  const fmt = s.labelNumFmt || (s.secondary ? spec.secondaryNumFmt : spec.numFmt);
  const showVal = s.labelParts ? (s.labelParts.value ? 1 : 0) : 1;
  const showCat = s.labelParts && s.labelParts.category ? 1 : 0;
  const shows = `<c:showLegendKey val="0"/><c:showVal val="${showVal}"/><c:showCatName val="${showCat}"/><c:showSerName val="0"/>` +
    `<c:showPercent val="${s.labelParts && s.labelParts.percent ? 1 : 0}"/><c:showBubbleSize val="0"/>`;
  // a label Tableau keeps on a line's first / last mark stays inside the pane: beside the point, not centred over it
  const marks = (s.values || []).map((v, i) => num(v) === null ? -1 : i).filter(i => i >= 0);
  const side = i => only && pos === "t" && marks.length > 1 ? (i === marks[0] ? "r" : i === marks[marks.length - 1] ? "l" : null) : null;
  const placed = only ? [...only].filter(i => side(i)) : [];
  if (placed.length) {
    const one = i => `<c:dLbl><c:idx val="${i}"/>${fmt ? `<c:numFmt formatCode="${esc(fmt)}" sourceLinked="0"/>` : ""}` +
      `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${labelTxPr(spec)}<c:dLblPos val="${side(i)}"/>${shows}</c:dLbl>`;
    const del = i => `<c:dLbl><c:idx val="${i}"/><c:delete val="1"/></c:dLbl>`;
    const items = [...(hidden || []).map(i => [i, del(i)]), ...placed.map(i => [i, one(i)])].sort((a, b) => +a[0] - +b[0]);
    return `<c:dLbls>${items.map(x => x[1]).join("")}${fmt ? `<c:numFmt formatCode="${esc(fmt)}" sourceLinked="0"/>` : ""}` +
      `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${labelTxPr(spec)}<c:dLblPos val="${pos}"/>${shows}</c:dLbls>`;
  }
  return `<c:dLbls>${drop}${fmt ? `<c:numFmt formatCode="${esc(fmt)}" sourceLinked="0"/>` : ""}` +
    `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${labelTxPr(spec)}` +
    `${pos ? `<c:dLblPos val="${pos}"/>` : ""}<c:showLegendKey val="0"/><c:showVal val="${showVal}"/>` +
    `<c:showCatName val="${showCat}"/><c:showSerName val="0"/><c:showPercent val="${s.labelParts && s.labelParts.percent ? 1 : 0}"/><c:showBubbleSize val="0"/>` +
    `${spec.kind === "pie" || spec.kind === "doughnut" ? '<c:showLeaderLines val="1"/>' : ""}</c:dLbls>`;
}

function markerXml(symbol, color, size) {
  if (!symbol || symbol === "none") return `<c:marker><c:symbol val="none"/></c:marker>`;
  return `<c:marker><c:symbol val="${symbol}"/><c:size val="${size || 6}"/>` +
    `<c:spPr>${solid(color, 85000)}${line(color, 9525)}</c:spPr></c:marker>`;
}

/* a reference line's stroke: colour with its opacity, width (px), dashes; none when switched off */
function refLineLn(rl) {
  if (rl.hidden) return `<a:ln><a:noFill/></a:ln>`;
  return `<a:ln w="${Math.round(rl.width * 9525)}"><a:solidFill><a:srgbClr val="${hex(rl.color)}">` +
    `${rl.alpha < 1 ? `<a:alpha val="${Math.round(rl.alpha * 100000)}"/>` : ""}</a:srgbClr></a:solidFill>` +
    `${rl.dash ? '<a:prstDash val="dash"/>' : ""}</a:ln>`;
}

/* the line's label on one of its points (idx): an Excel number format on the line's value ("Average",
 * "Avg. $1.2M"); valueFlag picks the value Excel formats (showVal = y, showCatName = x of an XY line) */
function refLineLabel(spec, rl, pos, valueFlag = "showVal", idx = 0) {
  if (!rl.labelFmt) return "";
  const font = { name: spec.font.name, size: spec.font.size, color: rl.font.color || "555555" };
  const flags = on => ["showLegendKey", "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize"]
    .map(f => `<c:${f} val="${on && f === valueFlag ? 1 : 0}"/>`).join("");
  return `<c:dLbls><c:dLbl><c:idx val="${idx}"/><c:numFmt formatCode="${esc(rl.labelFmt)}" sourceLinked="0"/>` +
    `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${txPr(font, { bold: rl.font.bold, noWrap: true })}` +
    `<c:dLblPos val="${pos}"/>${flags(true)}</c:dLbl>${flags(false)}</c:dLbls>`;
}

/* a vertical reference line across horizontal bars: XY points (value, 0) → (value, 1) on the hidden x2 / y2 axes */
function refLineXY(spec, rl, k) {
  const lit = vals => `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>` +
    vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("") + `</c:numLit>`;
  return `<c:ser><c:idx val="${k}"/><c:order val="${k}"/><c:tx><c:v>${esc(rl.name)}</c:v></c:tx>` +
    `<c:spPr>${refLineLn(rl)}</c:spPr><c:marker><c:symbol val="none"/></c:marker>${refLineLabel(spec, rl, "r", "showCatName")}` +
    `<c:xVal>${lit([rl.value, rl.value])}</c:xVal><c:yVal>${lit([0, 1])}</c:yVal><c:smooth val="0"/></c:ser>`;
}

/* bullet targets across horizontal bars: one vertical tick per category at its target, on the hidden x2 / y2 axes
 * (y2 runs 0 – 1 bottom to top; the first category is the top band). Blank points between ticks break the line. */
function targetTicksXY(spec, k) {
  const t = spec.targets, n = t.values.length;
  const xs = [], ys = [];
  t.values.forEach((v, i) => {
    const mid = 1 - (i + 0.5) / n, half = 0.32 / n;
    xs.push(v, v, null); ys.push(mid - half, mid + half, null);
  });
  const lit = vals => `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>` +
    vals.map((v, i) => num(v) === null ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("") + `</c:numLit>`;
  return `<c:ser><c:idx val="${k}"/><c:order val="${k}"/><c:tx><c:v>${esc(t.name || "Target")}</c:v></c:tx>` +
    `<c:spPr><a:ln w="28575" cap="flat">${solid(t.color || "333333")}</a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker>` +
    `<c:xVal>${lit(xs)}</c:xVal><c:yVal>${lit(ys)}</c:yVal><c:smooth val="0"/></c:ser>`;
}

/* marks over horizontal bars' category bands (dots of a dot plot or lollipop, a dumbbell's joining lines): XY points
 * (value, band) on the hidden x2 / y2 axes – a band's centre is 1 − (i + ½) / n; a null band breaks the line */
function overlayXY(spec, o, k) {
  const n = spec.categories && spec.categories.levels.length ? spec.categories.levels[0].length : 1;
  const lit = vals => `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>` +
    vals.map((v, i) => num(v) === null ? "" : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("") + `</c:numLit>`;
  const ys = o.cats.map(c => c === null || c === undefined ? null : +(1 - (c + 0.5) / n).toFixed(6));
  const xs = o.values.map((v, i) => ys[i] === null ? null : v);
  const ln = o.line ? `<a:ln w="${Math.round((o.lineWidth || 1.5) * 9525)}" cap="rnd">${solid(o.color)}</a:ln>` : `<a:ln w="19050"><a:noFill/></a:ln>`;
  return `<c:ser><c:idx val="${k}"/><c:order val="${k}"/><c:tx><c:v>${esc(o.name || "Marks")}</c:v></c:tx>` +
    `<c:spPr>${ln}</c:spPr>${markerXml(o.marker || "none", o.color, o.markerSize || 7)}` +
    `<c:xVal>${lit(xs)}</c:xVal><c:yVal>${lit(ys)}</c:yVal><c:smooth val="0"/></c:ser>`;
}

/* a hidden value axis pinned to [min, max] (the XY overlay's x2 / y2) */
function hiddenValAx(id, cross, pos, min, max) {
  return `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/><c:max val="${max}"/><c:min val="${min}"/></c:scaling>` +
    `<c:delete val="1"/><c:axPos val="${pos}"/><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/>` +
    `<c:minorTickMark val="none"/><c:tickLblPos val="none"/><c:crossAx val="${cross}"/><c:crosses val="max"/>` +
    `<c:crossBetween val="midCat"/></c:valAx>`;
}

/** Tableau-like axis ticks: a "nice" step (1 / 2 / 2.5 / 5 × 10ⁿ) giving about `ticks` intervals
 * @param {number} range @param {number} ticks @returns {number | null} */
function niceUnit(range, ticks) {
  if (!(range > 0) || !(ticks > 0)) return null;
  const raw = range / ticks, pow = Math.pow(10, Math.floor(Math.log10(raw))), n = raw / pow;
  return +((n <= 1 ? 1 : n <= 2 ? 2 : n <= 2.5 ? 2.5 : n <= 5 ? 5 : 10) * pow).toPrecision(12);
}

/**
 * The primary value axis as Tableau draws it: ticks at the workbook's spacing, else a "nice" step about every
 * 55 px (110 px across); the range starts at zero (without "Include zero": at a tick below the data) and ends
 * just past the data, so no empty tick sits above the marks. Stacked bars / areas count as their totals.
 * Synchronized axes always get an explicit range: the secondary axis is pinned to it.
 * @param {ChartSpec} spec @param {{ w: number, h: number } | null} axisPlot approximate plot area in px
 * @returns {{ fixed: { min?: number, max?: number }, unit: number | null }}
 */
function valueRange(spec, axisPlot) {
  // 100 % stacked: 0 – 100 %, a tick about every 55 px as Tableau's (a short plot: 0 %, 50 %, 100 %)
  if (spec.percent) return { fixed: { min: 0, max: 1 }, unit: spec.valueMajorUnit ||
    (axisPlot && axisPlot.h < SHORT_PLOT_PX ? niceUnit(1, Math.max(2, Math.round(axisPlot.h / 55))) : 0.2) };
  const fixed = { min: spec.valueMin, max: spec.valueMax };
  // values that follow slicers can grow past Tableau's: only the workbook's own range and spacing are kept, Excel
  // scales the rest (from zero when every value is positive, as Tableau's axis starts). Dots, targets and reference
  // lines over bars are drawn on an axis pinned to this range: such a chart keeps Tableau's
  // (a waterfall following them computes its bars from cells of its own: spec.helpers)
  const follows = ((spec.valueFormula && spec.series.some(s => s.values && /** @type {any} */ (s.values).src)) || /** @type {any} */ (spec).helpers) &&
                  !spec.overlay && !spec.targets && !(spec.refLines && spec.refLines.length);
  if (follows) {
    const all = spec.series.flatMap(s => s.values || []).filter(v => num(v) !== null);
    if (fixed.min === undefined && spec.includeZero !== false && all.length && all.every(v => v >= 0)) fixed.min = 0;
    // a short plot: Excel's own step crowds its few px with labels (0 10 20 … 50) – Tableau's, about every 55 px, of
    // the values at export (a selection that grows them past it only adds ticks)
    let unit = spec.valueMajorUnit || null;
    if (!unit && axisPlot && spec.barDir !== "bar" && axisPlot.h < SHORT_PLOT_PX && all.length && !spec.stacked && !spec.valueAxisHidden) {
      const lo = Math.min(0, ...all), hi = Math.max(0, ...all);
      unit = niceUnit(hi - lo, Math.max(2, Math.round(axisPlot.h / 55)));
    }
    return { fixed, unit };
  }
  const vals = [], stacks = new Map();
  spec.series.filter(s => (!s.secondary || spec.secondarySync) && s.values).forEach(s => {
    const type = s.type || spec.kind;
    if (spec.stacked && !s.refLine && (type === "bar" || type === "area")) {
      const g = stacks.get(type) || [];
      stacks.set(type, g);
      s.values.forEach((v, i) => {
        const x = num(v);
        if (x === null) return;
        const t = g[i] || (g[i] = { pos: 0, neg: 0 });
        if (x >= 0) t.pos += x; else t.neg += x;
      });
    } else s.values.forEach(v => { const x = num(v); if (x !== null) vals.push(x); });
  });
  stacks.forEach(g => g.forEach(t => { if (t) vals.push(t.pos, t.neg); }));
  (spec.refLines || []).forEach(r => vals.push(r.value));
  (spec.targets ? spec.targets.values : []).forEach(v => { if (num(v) !== null) vals.push(v); });
  (spec.overlay || []).forEach(o => o.values.forEach(v => { if (num(v) !== null) vals.push(v); }));
  const ticks = axisPlot ? Math.max(2, Math.round(spec.barDir === "bar" ? axisPlot.w / 110 : axisPlot.h / 55)) : null;
  const aligned = (v, u) => Math.abs(v / u - Math.round(v / u)) < 1e-9;
  const round = v => +v.toPrecision(12);
  const pinned = fixed.min !== undefined && fixed.max !== undefined;
  // Excel labels ticks from the axis minimum: a fixed minimum off the tick step starts at the step below
  // it, so the labels are round numbers as on Tableau's axis (a hidden axis keeps the exact range)
  const dates = /[dmy]/i.test(spec.numFmt || "") && !/[0#]/.test(spec.numFmt || "");   // Gantt: Excel's own date steps
  const settle = (range, unit) => {
    if (!unit || range.min === undefined || aligned(range.min, unit)) return unit;
    if (spec.valueAxisHidden || dates) return null;
    range.min = round(Math.floor(range.min / unit) * unit);
    return unit;
  };
  // the workbook's spacing, unless it would crowd the axis with ticks (data far beyond what it was set for)
  const spacing = span => spec.valueMajorUnit && span / spec.valueMajorUnit <= 40 ? spec.valueMajorUnit : null;
  if (!vals.length || pinned || !(ticks || spec.includeZero === false || spec.secondarySync || (spec.refLines && spec.refLines.length))) {
    const span = pinned ? fixed.max - fixed.min : 0;
    const unit = pinned ? spacing(span) || niceUnit(span, ticks || 5) : spec.valueMajorUnit || null;
    return { fixed, unit: settle(fixed, unit) };
  }
  const zero = spec.includeZero !== false;
  let lo = fixed.min ?? Math.min(...vals), hi = fixed.max ?? Math.max(...vals);
  if (zero) { lo = Math.min(0, lo); hi = Math.max(0, hi); }
  if (!(hi > lo)) hi = lo + (Math.abs(lo) || 1);
  const unit = spacing(hi - lo) || niceUnit(hi - lo, ticks || 4);
  // labels above the marks need room inside the plot
  const outside = spec.series.some(s => s.labels && !s.refLine) && !/^(inBase|ctr|inEnd)$/.test(spec.labelPos || "");
  const pad = (hi - lo) * (outside ? 0.12 : 0.04);
  const range = { min: fixed.min ?? (zero && lo >= 0 ? 0 : round(Math.floor((lo - pad) / unit) * unit)),
                  max: fixed.max ?? (zero && hi <= 0 ? 0 : round(hi + pad)) };
  return { fixed: range, unit: settle(range, unit) };
}

/** data labels taken from cells (scatter / bubble: Label = a dimension; waterfall steps over their bars)
 * @param {ChartSeries} s @param {ChartRefs["series"][number]} r */
function labelRangeXml(s, r) {
  if (!s.labelTexts || !r.lbl) return "";
  return `<c:extLst><c:ext uri="{02D57815-91ED-43cb-92C2-25804820EDAC}" xmlns:c15="${C15}"><c15:datalabelsRange>` +
    `<c15:f>${esc(r.lbl)}</c15:f><c15:dlblRangeCache>${strCache(s.labelTexts)}</c15:dlblRangeCache>` +
    `</c15:datalabelsRange></c:ext></c:extLst>`;
}

/** @param {ChartSpec} spec @param {ChartSeries} s @param {number} k @param {ChartRefs} refs @param {string} type
 *  @param {number[]} [hiddenLabels] points whose data label is left out */
function seriesXml(spec, s, k, refs, type, hiddenLabels) {
  const r = refs.series[k];
  const head = `<c:idx val="${k}"/><c:order val="${k}"/>${serTx(r.tx, s.name)}`;
  const pc = s.pointColors || [];
  const levels = spec.categories ? spec.categories.levels : [];
  // the workbook's opacity (Color → Opacity), in DrawingML's 1/1000 %
  const opacity = s.alpha !== undefined && s.alpha < 1 ? Math.round(Math.max(0, s.alpha) * 100000) : null;
  if (type === "bar") {
    const dpts = pc.map((c, i) => c ? `<c:dPt><c:idx val="${i}"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/>` +
      `<c:spPr>${solid(c, opacity)}<a:ln><a:noFill/></a:ln></c:spPr></c:dPt>` : "").join("");
    const pos = spec.labelPos || (spec.stacked ? null : "outEnd");
    return `<c:ser>${head}<c:spPr>${solid(s.color, opacity)}<a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="0"/>` +
      `${dpts}${dLbls(spec, s, pos)}${catXml(refs.cat, levels, spec.dateAxis, spec.categories.starts)}${valXml("val", r.val, s.values)}${labelRangeXml(s, r)}</c:ser>`;
  }
  if (type === "line" && s.refLine) {         // reference line: flat, no markers; label above the 2nd point, as Tableau's
    const at = Math.min(1, Math.max(0, (s.values || []).length - 1));
    return `<c:ser>${head}<c:spPr>${refLineLn(s.refLine)}</c:spPr><c:marker><c:symbol val="none"/></c:marker>` +
      `${refLineLabel(spec, s.refLine, "t", "showVal", at)}${catXml(refs.cat, levels, spec.dateAxis, spec.categories.starts)}${valXml("val", r.val, s.values)}<c:smooth val="0"/></c:ser>`;
  }
  if (type === "line") {
    const lineSp = s.line === false ? `<a:ln w="28575"><a:noFill/></a:ln>`
      : `<a:ln w="22225" cap="rnd">${solid(s.color, opacity)}<a:round/></a:ln>`;
    const symbol = s.marker ? (s.markerSymbol || "circle") : "none";
    const size = s.markerSize || 7;
    const dpts = s.marker ? pc.map((c, i) => c ? `<c:dPt><c:idx val="${i}"/>${markerXml(symbol, c, size)}<c:bubble3D val="0"/></c:dPt>` : "").join("") : "";
    return `<c:ser>${head}<c:spPr>${lineSp}</c:spPr>${markerXml(symbol, s.color, size)}${dpts}` +
      `${dLbls(spec, s, "t", hiddenLabels)}${catXml(refs.cat, levels, spec.dateAxis, spec.categories.starts)}${valXml("val", r.val, s.values)}<c:smooth val="0"/>${labelRangeXml(s, r)}</c:ser>`;
  }
  if (type === "area") {
    return `<c:ser>${head}<c:spPr>${solid(s.color, opacity !== null || s.alpha !== undefined ? opacity : spec.stacked ? null : 75000)}<a:ln><a:noFill/></a:ln></c:spPr>` +
      `${dLbls(spec, s, null)}${catXml(refs.cat, levels, spec.dateAxis, spec.categories.starts)}${valXml("val", r.val, s.values)}</c:ser>`;
  }
  if (type === "pie") {
    const dpts = pc.map((c, i) => `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/>` +
      `<c:spPr>${solid(c || s.color)}${line("FFFFFF", 12700)}</c:spPr></c:dPt>`).join("");
    return `<c:ser>${head}${dpts}${dLbls(spec, s, spec.kind === "pie" ? "bestFit" : null)}` +
      `${catXml(refs.cat, levels, spec.dateAxis, spec.categories.starts)}${valXml("val", r.val, s.values)}${labelRangeXml(s, r)}</c:ser>`;
  }
  if (type === "bubble") {
    // packed bubbles: opaque with a white outline like Tableau; map marks slightly see-through
    const fill = c => `${solid(c, spec.packed ? null : 80000)}${line("FFFFFF", 9525)}`;
    const dpts = pc.map((c, i) => c ? `<c:dPt><c:idx val="${i}"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/>` +
      `<c:spPr>${fill(c)}</c:spPr></c:dPt>` : "").join("");
    return `<c:ser>${head}<c:spPr>${fill(s.color)}</c:spPr><c:invertIfNegative val="0"/>${dpts}` +
      `${dLbls(spec, s, spec.packed ? "ctr" : "r", hiddenLabels)}${valXml("xVal", r.x, s.x)}${valXml("yVal", r.y, s.y)}` +
      `${valXml("bubbleSize", r.size, s.size)}<c:bubble3D val="0"/>${labelRangeXml(s, r)}</c:ser>`;
  }
  // scatter; a line of measure against measure joins its points in order, without markers (Tableau's line mark)
  const dpts = pc.map((c, i) => c ? `<c:dPt><c:idx val="${i}"/>${markerXml("circle", c, 7)}<c:bubble3D val="0"/></c:dPt>` : "").join("");
  if (s.line) {
    return `<c:ser>${head}<c:spPr><a:ln w="22225" cap="rnd">${solid(s.color, opacity)}<a:round/></a:ln></c:spPr>` +
      `${markerXml(s.marker ? "circle" : "none", s.color, 7)}${dLbls(spec, s, "r")}${valXml("xVal", r.x, s.x)}${valXml("yVal", r.y, s.y)}` +
      `<c:smooth val="0"/>${labelRangeXml(s, r)}</c:ser>`;
  }
  return `<c:ser>${head}<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>${markerXml("circle", s.color, 7)}${dpts}` +
    `${dLbls(spec, s, "r")}${valXml("xVal", r.x, s.x)}${valXml("yVal", r.y, s.y)}<c:smooth val="0"/>` +
    labelRangeXml(s, r) + `</c:ser>`;
}

/** @param {ChartSpec} spec @param {number} id @param {number} cross @param {{ deleted?: boolean, rot?: number, skip?: number }} [o]
 *  rot: label rotation (60000ths of a degree), skip: every n-th label shown */
function catAxis(spec, id, cross, o = {}) {
  if (spec.dateAxis) return dateAxisXml(spec, id, cross, o);
  const horizontal = spec.barDir === "bar";
  const multi = spec.categories && spec.categories.levels.length > 1;
  return `<c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="${horizontal ? "maxMin" : "minMax"}"/></c:scaling>` +
    `<c:delete val="${o.deleted ? 1 : 0}"/><c:axPos val="${horizontal ? "l" : "b"}"/>` +
    `${o.deleted ? "" : title(spec.categoryTitle, spec.font, horizontal)}` +
    `<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/>` +
    `<c:tickLblPos val="${o.deleted ? "none" : spec.valueReversed ? "nextTo" : "low"}"/><c:spPr>${spec.axisLine === false ? "<a:ln><a:noFill/></a:ln>" : line("D4D4D4", 9525)}</c:spPr>${txPr(spec.font, { rot: o.rot })}` +
    `<c:crossAx val="${cross}"/><c:crosses val="${spec.valueReversed && !o.deleted ? "max" : "autoZero"}"/><c:auto val="1"/><c:lblAlgn val="ctr"/>` +
    `<c:lblOffset val="100"/>${o.rot !== undefined ? `<c:tickLblSkip val="${o.skip || 1}"/>` : ""}<c:noMultiLvlLbl val="${multi ? 0 : 1}"/></c:catAx>`;
}

/** a continuous date axis: Excel's date axis, its labels as far apart as dateTicks found room for
 * @param {ChartSpec} spec @param {number} id @param {number} cross @param {{ deleted?: boolean, rot?: number }} o */
function dateAxisXml(spec, id, cross, o) {
  const t = spec.dateTicks || { majorUnit: 1, majorTimeUnit: "years", numFmt: "yyyy" };
  const d = spec.dateAxis;                                // panes of small multiples share one span (max before min)
  const span = (d.max !== undefined ? `<c:max val="${d.max}"/>` : "") + (d.min !== undefined ? `<c:min val="${d.min}"/>` : "");
  return `<c:dateAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/>${span}</c:scaling>` +
    `<c:delete val="${o.deleted ? 1 : 0}"/><c:axPos val="b"/>${o.deleted ? "" : title(spec.categoryTitle, spec.font, false)}` +
    `<c:numFmt formatCode="${esc(t.numFmt)}" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/>` +
    `<c:tickLblPos val="${o.deleted ? "none" : spec.valueReversed ? "nextTo" : "low"}"/><c:spPr>${spec.axisLine === false ? "<a:ln><a:noFill/></a:ln>" : line("D4D4D4", 9525)}</c:spPr>${txPr(spec.font, { rot: o.rot })}` +
    `<c:crossAx val="${cross}"/><c:crosses val="${spec.valueReversed && !o.deleted ? "max" : "autoZero"}"/><c:auto val="0"/><c:lblOffset val="100"/>` +
    `<c:baseTimeUnit val="${spec.dateAxis.base}"/><c:majorUnit val="${t.majorUnit}"/><c:majorTimeUnit val="${t.majorTimeUnit}"/></c:dateAx>`;
}

/**
 * Tableau labels a time axis at whole units – years, else quarters, months or days – as close together as the labels
 * fit: the first step whose labels fit the plot width.
 * @param {ChartSpec} spec @param {{ w: number, h: number } | null} axisPlot
 * @returns {{ majorUnit: number, majorTimeUnit: "days" | "months" | "years", numFmt: string } | null}
 */
function dateTicks(spec, axisPlot) {
  const d = spec.dateAxis;
  if (!d) return null;
  const s = d.serials, first = d.min !== undefined ? d.min : s[0], last = d.max !== undefined ? d.max : s[s.length - 1];
  const month = v => { const t = new Date((v - 25569) * 86400000); return t.getUTCFullYear() * 12 + t.getUTCMonth(); };
  const months = month(last) - month(first), days = last - first;
  const charPx = ((spec.font && spec.font.size) || 9) * 4 / 3 * 0.55;          // average character of the label font
  const width = axisPlot ? axisPlot.w : 400;
  /** @type {[number, "days" | "months" | "years", string][]} */
  const steps = [
    ...(d.base === "days" ? /** @type {[number, "days", string][]} */ ([[1, "days", "d mmm"], [7, "days", "d mmm"], [14, "days", "d mmm"]]) : []),
    ...(d.base !== "years" ? /** @type {[number, "months", string][]} */ ([[1, "months", "mmm yyyy"], [3, "months", "mmm yyyy"], [6, "months", "mmm yyyy"]]) : []),
    [1, "years", "yyyy"], [2, "years", "yyyy"], [5, "years", "yyyy"], [10, "years", "yyyy"], [25, "years", "yyyy"]
  ];
  for (const [unit, of, fmt] of steps) {
    const count = Math.floor((of === "days" ? days : of === "months" ? months : months / 12) / unit) + 1;
    if (count * (fmt.length * charPx + 12) <= width) return { majorUnit: unit, majorTimeUnit: of, numFmt: fmt };
  }
  return { majorUnit: 50, majorTimeUnit: "years", numFmt: "yyyy" };
}

/* Tableau axes "include zero" by default; Excel would otherwise auto-scale from a non-zero minimum */
function zeroScaling(values) {
  const nums = values.filter(v => num(v) !== null);
  if (!nums.length) return "";
  if (Math.min(...nums) >= 0) return `<c:min val="0"/>`;
  if (Math.max(...nums) <= 0) return `<c:max val="0"/>`;
  return "";
}

/* axis ticks: Tableau's automatic axis drops decimals and abbreviates thousands (100K, 1.5M);
 * labels and the data cells keep the full number format */
function axisFmt(fmt, values) {
  const nums = (values || []).filter(v => num(v) !== null).map(Math.abs);
  if (!fmt || !nums.length) return fmt || "General";
  const max = Math.max(...nums);
  if (max < 1000) {
    if (/%/.test(fmt)) return fmt;
    if (max >= 10) return fmt.replace(/0.0+/g, "0");
    return fmt.replace(/0.0{2,}/g, "0.0");
  }
  const plain = fmt.match(/^("[^"]*")?(#,##0|0)(.0+)?$/);
  if (plain && max >= 1e6) return (plain[1] || "") + '#,##0.0,,"M"';
  if (plain && max >= 1e4) return (plain[1] || "") + '#,##0,"K"';
  return fmt.replace(/0.0+/g, "0");
}

/**
 * @param {ChartSpec} spec @param {number} id @param {number} cross
 * @param {{ pos?: string, crosses?: string, grid?: boolean, fixed?: { min?: number, max?: number }, title?: string,
 *           numFmt?: string, values?: (number | null)[], deleted?: boolean, hidden?: boolean, lowLabels?: boolean, midCat?: boolean,
 *           majorUnit?: number | null, tickFmt?: string, reversed?: boolean }} [o] tickFmt = the workbook's own tick format, used as is;
 *   hidden = kept for its scale but not drawn (no labels, tick marks or line) – Excel draws the series of a deleted value
 *   axis on the other axis group's value axis
 */
function valAxis(spec, id, cross, o = {}) {
  const horizontal = spec.barDir === "bar" && spec.kind !== "scatter";
  const pos = o.pos || (horizontal ? "b" : "l");
  const crosses = o.crosses || (horizontal ? "max" : "autoZero");
  const grid = o.grid !== false && spec.gridlines !== false
    ? `<c:majorGridlines><c:spPr>${line("EBEBEB", 9525)}</c:spPr></c:majorGridlines>` : "";
  const fixed = o.fixed || {};                       // CT_Scaling order: orientation, max, min
  const scale = fixed.max !== undefined || fixed.min !== undefined
    ? (fixed.max !== undefined ? `<c:max val="${fixed.max}"/>` : "") + (fixed.min !== undefined ? `<c:min val="${fixed.min}"/>` : "")
    : spec.includeZero === false || !o.values ? "" : zeroScaling(o.values);
  return `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="${o.reversed ? "maxMin" : "minMax"}"/>${scale}</c:scaling><c:delete val="${o.deleted ? 1 : 0}"/>` +
    `<c:axPos val="${pos}"/>${grid}${title(o.title, spec.font, (pos === "l" || pos === "r") && !o.titleFlat, !!o.titleFlat)}` +
    `<c:numFmt formatCode="${esc(o.tickFmt || axisFmt(o.numFmt || "General", o.values))}" sourceLinked="0"/><c:majorTickMark val="none"/>` +
    `<c:minorTickMark val="none"/><c:tickLblPos val="${o.hidden ? "none" : o.lowLabels ? "low" : "nextTo"}"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>` +
    `${txPr(spec.font)}<c:crossAx val="${cross}"/><c:crosses val="${crosses}"/>` +
    `<c:crossBetween val="${o.midCat ? "midCat" : "between"}"/>${o.majorUnit ? `<c:majorUnit val="${o.majorUnit}"/>` : ""}</c:valAx>`;
}

/* Excel's largest bubble: diameter / the plot area's shorter side for a bubble scale s (1 = 100 %).
 * Measured in Excel (one bubble, scales 10–300 %, three plot shapes): 0.03 at 10 %, 0.234 at 100 %,
 * 0.485 at 300 % – A·s / (1 + B·s) fits within 1 %. */
const BUBBLE_A = 0.3014, BUBBLE_B = 0.288;
const bubbleRatio = s => BUBBLE_A * s / (1 + BUBBLE_B * s);
const bubbleScaleFor = ratio => ratio >= bubbleRatio(3) ? 3 : Math.max(0.01, ratio / (BUBBLE_A - BUBBLE_B * ratio));

/** a packed-bubble chart's plot area, as fractions of the chart (it fills the chart: no axes) */
const bubbleFrame = spec => ({ x: 0.02, y: 0.03, w: spec.legend ? 0.74 : 0.96, h: 0.94 });

/**
 * Tableau's tooltip on each packed bubble: an invisible circle over it (px inside the chart) whose hyperlink
 * ScreenTip is the tooltip text, placed with the same fit Excel draws the bubbles at.
 * @param {ChartSpec} spec @param {{ widthPx: number, heightPx: number }} size @param {string} link
 */
function bubbleTips(spec, size, link) {
  if (spec.kind !== "bubble" || !spec.packed) return [];
  const f = bubbleFrame(spec);
  const plot = { w: size.widthPx * f.w, h: size.heightPx * f.h };
  const fit = xyFit(spec, plot);
  if (!fit.unit) return [];
  return spec.series.flatMap(s => {
    if (!s.tooltips || !s.size) return [];
    const max = Math.max(...s.size.filter(v => num(v) !== null && v > 0));
    return s.size.map((v, i) => {
      if (!s.tooltips[i] || !(v > 0)) return null;
      const r = Math.sqrt(v / max) * fit.unit;
      const cx = size.widthPx * f.x + (s.x[i] - fit.x.min) * fit.unit, cy = size.heightPx * f.y + (fit.y.max - s.y[i]) * fit.unit;
      return { x: cx - r, y: cy - r, w: 2 * r, h: 2 * r, prst: "ellipse", text: s.tooltips[i], link };
    }).filter(Boolean);
  });
}

/**
 * Packed bubbles: fixed axis ranges and the bubble scale for the plot area's size in px, so the bubbles touch
 * without overlapping.
 * @param {ChartSpec} spec @param {{ w: number, h: number } | null} plot
 * @returns {{ bubbleScale: number, unit?: number, x?: { min: number, max: number }, y?: { min: number, max: number } }}
 */
function xyFit(spec, plot) {
  if (!plot || !spec.packed) return { bubbleScale: 100 };
  const W = plot.w, H = plot.h, shorter = Math.min(W, H);
  const b = spec.packed;                                      // radius units: largest bubble radius = 1
  const want = 0.98 * Math.min(W / b.w, H / b.h);             // px per radius unit that fits the packing
  const bubbleScale = Math.max(1, Math.round(100 * bubbleScaleFor(2 * want / shorter)));
  const unit = bubbleRatio(bubbleScale / 100) * shorter / 2;  // px per radius unit as Excel draws it
  return { bubbleScale, unit, x: { min: b.cx - W / 2 / unit, max: b.cx + W / 2 / unit },
           y: { min: b.cy - H / 2 / unit, max: b.cy + H / 2 / unit } };
}

/**
 * Tableau's column headers stay horizontal while every name fits its slot; otherwise Tableau turns them to read
 * upwards and, where even turned labels would overlap, shows every second (third …) one. The turn (degrees) and the
 * label step for a vertical chart's category axis, or null when the labels stay as they are (or the workbook turns them).
 * @param {ChartSpec} spec @param {{ w: number, h: number } | null} axisPlot @returns {{ rot: number, skip: number } | null}
 */
function categoryLabelFit(spec, axisPlot) {
  const levels = spec.categories ? spec.categories.levels : [];
  if (!axisPlot || spec.barDir === "bar" || spec.categoryRotation || spec.categoryAxisHidden || spec.dateAxis ||
      levels.length !== 1 || !levels[0].length) return null;
  const names = levels[0].map(c => String(c == null ? "" : c));
  const fontPx = ((spec.font && spec.font.size) || 9) * 4 / 3;
  const slot = axisPlot.w / names.length;
  if (names.every(n => n.length * fontPx * 0.5 <= slot - 4)) return null;     // average character: half the font size
  // names of a few words (Human Resources): Excel wraps a horizontal label at its spaces – kept horizontal like
  // Tableau's headers while every word fits its slot, where a turned one would break mid-word
  if (names.every(n => { const w = n.trim().split(/\s+/); return w.length <= 3 && w.every(x => x.length * fontPx * 0.5 <= slot - 4); })) return null;
  // Tableau scrolls a chart whose category cells (Format → Cell width) are wider than the view, its names horizontal
  // while they fit a cell; an Excel chart cannot scroll, so every n-th is shown instead, as many as fit side by side
  if (spec.categorySlotPx > slot && names.every(n => n.length * fontPx * 0.5 <= spec.categorySlotPx - 4)) {
    return { rot: 0, skip: Math.max(1, Math.ceil((Math.max(...names.map(n => n.length)) * fontPx * 0.5 + 8) / slot)) };
  }
  // numbers (bins, years) stay horizontal like an axis: every n-th, as far apart as the longest needs
  if (names.every(n => /^-?[\d.,]+%?$/.test(n.trim()))) {
    return { rot: 0, skip: Math.max(1, Math.ceil((Math.max(...names.map(n => n.length)) * fontPx * 0.55 + 10) / slot)) };
  }
  return { rot: -90, skip: Math.max(1, Math.ceil(fontPx * 1.4 / slot)) };      // a turned label's line, with spacing
}

/** @param {ChartSpec} spec @param {ChartRefs} refs @param {{ w: number, h: number } | null} [plot] plot area in px (manual layout)
 *  @param {{ w: number, h: number } | null} [axisPlot] approximate plot area of a chart with axes, for Tableau-like tick spacing */
function plotAreaXml(spec, refs, plot = null, axisPlot = null) {
  spec.categoryFit = categoryLabelFit(spec, axisPlot);
  spec.dateTicks = dateTicks(spec, axisPlot);
  const k = spec.kind;
  if (k === "pie" || k === "doughnut") {
    const ser = seriesXml(spec, spec.series[0], 0, refs, "pie");
    return k === "pie"
      ? `<c:pieChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/></c:pieChart>`
      : `<c:doughnutChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/><c:holeSize val="${spec.holeSize || 55}"/></c:doughnutChart>`;
  }
  if (k === "scatter" || k === "bubble") {
    const fit = xyFit(spec, plot);
    // packed bubbles: a label is centred on its bubble; with "allow labels to overlap" off (Tableau's default) one
    // that would overlap a bigger bubble's label is hidden, as Tableau hides it
    const fontPx = (spec.font && spec.font.size || 9) * 4 / 3, charPx = fontPx * 0.55;
    const overlapping = s => {
      if (!spec.packed || !s.labelTexts || !fit.unit || spec.labelCull === false) return [];
      const boxes = s.labelTexts.map((t, i) => {
        const lines = String(t || "").split("\n");
        const w = Math.max(...lines.map(l => l.length)) * charPx, h = lines.length * fontPx * 1.2;
        const cx = (s.x[i] - fit.x.min) * fit.unit, cy = (fit.y.max - s.y[i]) * fit.unit;
        return { i, x: cx - w / 2, y: cy - h / 2, w, h, size: s.size[i] || 0 };
      });
      const kept = [], hidden = [];
      boxes.sort((a, b) => b.size - a.size).forEach(b => {
        const hit = kept.some(o => b.x < o.x + o.w && o.x < b.x + b.w && b.y < o.y + o.h && o.y < b.y + b.h);
        (hit ? hidden : kept).push(b);
      });
      return hidden.map(b => b.i);
    };
    const sers = spec.series.map((s, i) => seriesXml(spec, s, i, refs, k, overlapping(s))).join("");
    const group = k === "bubble"
      ? `<c:bubbleChart><c:varyColors val="0"/>${sers}<c:bubbleScale val="${fit.bubbleScale}"/><c:showNegBubbles val="0"/>` +
        `<c:sizeRepresents val="area"/><c:axId val="${AX.cat}"/><c:axId val="${AX.val}"/></c:bubbleChart>`
      : `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${sers}` +
        `<c:axId val="${AX.cat}"/><c:axId val="${AX.val}"/></c:scatterChart>`;
    return group +
      valAxis(spec, AX.cat, AX.val, { pos: "b", crosses: "autoZero", title: spec.axesHidden || spec.xAxisHidden ? null : spec.xTitle,
        numFmt: spec.xNumFmt, tickFmt: spec.xAxisNumFmt, midCat: true, grid: spec.xGridlines !== false, lowLabels: true, values: spec.series.flatMap(s => s.x),
        fixed: fit.x || { min: spec.xMin, max: spec.xMax }, deleted: spec.axesHidden || spec.xAxisHidden, majorUnit: spec.xMajorUnit, reversed: spec.xReversed }) +
      valAxis(spec, AX.val, AX.cat, { pos: "l", crosses: "autoZero", title: spec.axesHidden || spec.valueAxisHidden ? null : spec.valueTitle,
        numFmt: spec.numFmt, tickFmt: spec.valueAxisNumFmt, midCat: true, lowLabels: true, values: spec.series.flatMap(s => s.y),
        fixed: fit.y || { min: spec.valueMin, max: spec.valueMax }, deleted: spec.axesHidden || spec.valueAxisHidden, majorUnit: spec.valueMajorUnit,
        reversed: spec.valueReversed });
  }
  // labels written over the bars (waterfall steps): as Tableau, a label that would overlap one already
  // placed is left out – positions estimated from the plot size and the axis range
  const { fixed, unit } = valueRange(spec, axisPlot);
  /** @type {Map<number, number[]>} */
  const culled = new Map();
  if (axisPlot && spec.labelCull !== false && fixed.min !== undefined && fixed.max !== undefined && fixed.max > fixed.min) {
    const size = (spec.labelFont && spec.labelFont.size) || spec.font.size || 9, h = size * 4 / 3 + 2;
    spec.series.forEach((s, i) => {
      if (!s.labelTexts || (s.type || k) !== "line") return;
      const slot = axisPlot.w / Math.max(1, s.values.length), kept = [], hide = [];
      s.values.forEach((v, j) => {
        const t = s.labelTexts[j], n = num(v);
        if (n === null || !t) return;
        const w = String(t).length * size * 0.62 + 4, x = (j + 0.5) * slot;
        const y = axisPlot.h * (fixed.max - n) / (fixed.max - fixed.min) - h / 2 - 3;   // a label just above the point
        const box = { l: x - w / 2, r: x + w / 2, t: y - h / 2, b: y + h / 2 };
        if (kept.some(q => q.l < box.r && box.l < q.r && q.t < box.b && box.t < q.b)) hide.push(j); else kept.push(box);
      });
      if (hide.length) culled.set(i, hide);
    });
  }
  // bar / line / area / combo: one chart group per (type, axis)
  const groups = [];
  spec.series.forEach((s, i) => {
    const type = s.type || k;
    const key = type + (s.secondary ? "2" : "1");
    let g = groups.find(x => x.key === key);
    if (!g) groups.push(g = { key, type, secondary: !!s.secondary, items: [] });
    g.items.push(seriesXml(spec, s, i, refs, type, culled.get(i)));
  });
  const hasSecondary = groups.some(g => g.secondary);
  const xml = groups.map(g => {
    const ax = g.secondary ? `<c:axId val="${AX.cat2}"/><c:axId val="${AX.val2}"/>` : `<c:axId val="${AX.cat}"/><c:axId val="${AX.val}"/>`;
    const sers = g.items.join("");
    if (g.type === "bar") {
      const grouping = spec.percent ? "percentStacked" : spec.stacked ? "stacked" : "clustered";
      const gap = g.secondary && spec.secondaryGapWidth !== undefined ? spec.secondaryGapWidth : spec.gapWidth ?? 60;
      return `<c:barChart><c:barDir val="${spec.barDir || "col"}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${sers}` +
        `<c:gapWidth val="${gap}"/>${spec.stacked || spec.overlap ? '<c:overlap val="100"/>' : ""}${ax}</c:barChart>`;
    }
    if (g.type === "area") {
      return `<c:areaChart><c:grouping val="${spec.percent ? "percentStacked" : spec.stacked ? "stacked" : "standard"}"/><c:varyColors val="0"/>${sers}${ax}</c:areaChart>`;
    }
    // box plot: high-low lines = whiskers, up/down bars between first (Q1) and last (Q3) series = box
    const box = spec.boxPlot
      ? `<c:hiLowLines><c:spPr>${line(spec.boxPlot.color, 12700)}</c:spPr></c:hiLowLines>` +
        `<c:upDownBars><c:gapWidth val="${spec.gapWidth ?? 80}"/>` +
        `<c:upBars><c:spPr>${solid(spec.boxPlot.color, 35000)}${line(spec.boxPlot.color, 12700)}</c:spPr></c:upBars>` +
        `<c:downBars><c:spPr>${solid(spec.boxPlot.color, 35000)}${line(spec.boxPlot.color, 12700)}</c:spPr></c:downBars></c:upDownBars>`
      : spec.hiLowLines ? `<c:hiLowLines><c:spPr>${line(spec.hiLowLines.color, 19050)}</c:spPr></c:hiLowLines>` : "";
    return `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}${box}<c:marker val="1"/>${ax}</c:lineChart>`;
  }).join("");
  const axisValues = secondary => spec.series.filter(s => !!s.secondary === secondary).flatMap(s => s.values);
  // reference lines across horizontal bars: XY lines on hidden x2 / y2 axes, x2 pinned to the bars' value range – never
  // beside a secondary axis: Excel cannot open a workbook with a chart of three axis groups (the model leaves those out)
  let overlay = "", overlayAxes = "";
  const targets = spec.barDir === "bar" && spec.targets ? 1 : 0;
  const dots = spec.barDir === "bar" && spec.overlay ? spec.overlay : [];
  if (spec.barDir === "bar" && !hasSecondary && ((spec.refLines && spec.refLines.length) || targets || dots.length) &&
      fixed.min !== undefined && fixed.max !== undefined) {
    const k0 = spec.series.length + (spec.refLines || []).length + targets;
    overlay = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>` +
      (spec.refLines || []).map((rl, i) => refLineXY(spec, rl, spec.series.length + i)).join("") +
      (targets ? targetTicksXY(spec, spec.series.length + (spec.refLines || []).length) : "") +
      dots.map((o, j) => overlayXY(spec, o, k0 + j)).join("") +
      `<c:axId val="${AX.x2}"/><c:axId val="${AX.y2}"/></c:scatterChart>`;
    overlayAxes = hiddenValAx(AX.x2, AX.y2, "t", fixed.min, fixed.max) + hiddenValAx(AX.y2, AX.x2, "r", 0, 1);
  }
  // a hidden value axis beside a second axis group is kept, not drawn: Excel would draw a deleted axis's series on the
  // other group's axis, losing the scale Tableau gives each axis
  let axes = catAxis(spec, AX.cat, AX.val, { deleted: spec.categoryAxisHidden,
                                     // switched views of other names (writeSwitchedChart): Excel turns them as each one needs
                                     rot: spec.categoryAuto ? undefined : (spec.categoryRotation || (spec.categoryFit ? spec.categoryFit.rot : 0)) * 60000,
                                     skip: spec.categoryFit ? spec.categoryFit.skip : 1 }) +
    valAxis(spec, AX.val, AX.cat, { title: spec.rowTitle !== undefined ? spec.rowTitle : spec.valueAxisHidden ? null : spec.valueTitle,
                                    titleFlat: spec.rowTitle !== undefined && !(spec.grid && spec.grid.rowVertical),
                                    // whole-percent ticks read as Tableau's axis: 0 %, 50 %, 100 %, not 50.0 %
                                    numFmt: unit && /%/.test(spec.numFmt || "") && Math.abs(unit * 100 - Math.round(unit * 100)) < 1e-9
                                      ? "0%" : spec.numFmt, values: axisValues(false), fixed,
                                    deleted: spec.valueAxisHidden && !hasSecondary && spec.rowTitle === undefined,
                                    hidden: spec.valueAxisHidden && (hasSecondary || spec.rowTitle !== undefined),
                                    majorUnit: unit, tickFmt: spec.valueAxisNumFmt, reversed: spec.valueReversed });
  if (hasSecondary) {
    axes += catAxis(spec, AX.cat2, AX.val2, { deleted: true }) +
      valAxis(spec, AX.val2, AX.cat2, { pos: spec.barDir === "bar" ? "t" : "r", crosses: "max", grid: false,
        title: spec.secondaryAxisHidden ? null : spec.secondaryTitle, numFmt: spec.secondaryNumFmt || spec.numFmt, values: axisValues(true),
        hidden: spec.secondaryAxisHidden, tickFmt: spec.secondaryAxisNumFmt,
        // synchronized: the same range as the primary axis, so both measures are drawn to one scale;
        // else the workbook's own fixed range on the second axis, if it has one
        ...(spec.secondarySync ? { fixed, majorUnit: unit }
          : (spec.secondaryMin !== undefined || spec.secondaryMax !== undefined)
            ? { fixed: { min: spec.secondaryMin, max: spec.secondaryMax } } : {}) });
  }
  return xml + overlay + axes + overlayAxes;
}

/* small multiples: room round Tableau's pane grid for the headers and axes (px) */
const GRID_TOP = 18, GRID_BOTTOM = 22, GRID_PAD = 1, GRID_OVERLAP = 14, GRID_OVERLAP_V = 10, GRID_AXIS = 34;
/* a chart this short (px) with no value axis is a sparkline: its plot is laid out by hand (chartXml) */
const SPARKLINE_MAX_PX = 110;
/* a value axis this short (px) gets Tableau's tick spacing even when the chart follows the slicers (valueRange) */
const SHORT_PLOT_PX = 150;

/** the grid's left margin (px): the row headers (turned: one line high; else as wide as the longest, at most 120 px)
 * and the value axis labels unless the workbook hides them @param {any} g spec.grid @param {ChartSpec} spec */
function gridLeft(g, spec) {
  const fontPx = ((spec.font && spec.font.size) || 9) * 4 / 3;
  const head = !g.rowLen ? 0 : g.rowVertical ? fontPx + 6 : Math.min(120, g.rowLen * fontPx * 0.55 + 10);
  return Math.round(head + (spec.valueAxisHidden ? 0 : GRID_AXIS) + 4);
}

/** a pane chart's margins round its plot area (px): headers and axes only on the grid's outer panes; inner edges reach
 * into the neighbouring panes (the charts are see-through), so labels at a pane's edge keep their room and Excel never
 * moves a plot area to fit them – every pane's plot stays exactly at its place in the grid
 * @param {any} g spec.grid @param {ChartSpec} spec */
function gridInset(g, spec) {
  return { l: (g.col === 0 ? gridLeft(g, spec) : GRID_OVERLAP) + GRID_PAD, r: (g.col === g.cols - 1 ? 0 : GRID_OVERLAP) + GRID_PAD,
           t: (g.row === 0 ? GRID_TOP : GRID_OVERLAP_V) + GRID_PAD, b: (g.row === g.rows - 1 ? GRID_BOTTOM : GRID_OVERLAP_V) + GRID_PAD };
}

/** a pane chart's frame inside its block (px): the panes share the block evenly, the outer ones also hold the
 * headers and axes, so every plot area is the same size
 * @param {any} g spec.grid @param {ChartSpec} spec @param {number} W @param {number} H */
function gridFrame(g, spec, W, H) {
  const GRID_LEFT = gridLeft(g, spec);
  const pw = Math.max(12, (W - GRID_LEFT) / g.cols), ph = Math.max(12, (H - GRID_TOP - GRID_BOTTOM) / g.rows);
  const x = g.col === 0 ? 0 : GRID_LEFT + g.col * pw - GRID_OVERLAP, y = g.row === 0 ? 0 : GRID_TOP + g.row * ph - GRID_OVERLAP_V;
  const x2 = GRID_LEFT + (g.col + 1) * pw + (g.col === g.cols - 1 ? 0 : GRID_OVERLAP);
  const y2 = GRID_TOP + (g.row + 1) * ph + (g.row === g.rows - 1 ? GRID_BOTTOM : GRID_OVERLAP_V);
  return { x: Math.round(x), y: Math.round(y), w: Math.round(x2) - Math.round(x), h: Math.round(y2) - Math.round(y) };
}

/** chart part XML (xl/charts/chartN.xml)
 * @param {ChartSpec} spec @param {ChartRefs} refs
 * @param {{ widthPx?: number, heightPx?: number }} [size] the chart frame, for maps / packed bubbles
 * @returns {string} */
function chartXml(spec, refs, size) {
  // a pane of small multiples: the column's header above the top row, the row's header as the left column's axis
  // title, value labels on the left column and the category axis on the bottom row only (Tableau's shared axes)
  const g = spec.grid;
  if (g) {
    spec = { ...spec, legend: false, categoryTitle: "", background: null, categoryAxisHidden: spec.categoryAxisHidden || g.row !== g.rows - 1,
             valueAxisHidden: spec.valueAxisHidden || g.col !== 0, valueTitle: null, rowTitle: g.col === 0 ? g.rowLabel : undefined };
  }
  // no axes: the plot area fills the frame (room for the legend), so its size in px is known
  const frame = spec.axesHidden && spec.packed ? bubbleFrame(spec) : null;
  const plot = frame && { w: ((size && size.widthPx) || 600) * frame.w, h: ((size && size.heightPx) || 400) * frame.h };
  const gi = g && gridInset(g, spec);
  const axisPlot = !size || !size.widthPx ? null
    : gi ? { w: size.widthPx - gi.l - gi.r, h: (size.heightPx || 300) - gi.t - gi.b }
    : { w: size.widthPx * 0.85,
        // each extra level of a multi-level category axis takes a row of labels below the plot
        h: (size.heightPx || 300) * 0.72 - Math.max(0, ((spec.categories && spec.categories.levels.length) || 1) - 1) * (spec.font.size || 9) * 4 / 3 * 1.6 };
  const layout = frame
    ? `<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/>` +
      `<c:x val="${frame.x}"/><c:y val="${frame.y}"/><c:w val="${frame.w}"/><c:h val="${frame.h}"/></c:manualLayout></c:layout>`
    : "<c:layout/>";
  const legend = spec.legend
    ? `<c:legend><c:legendPos val="r"/><c:overlay val="0"/>${txPr(spec.font)}</c:legend>` : "";
  // a donut's centre text: the chart title laid over the hole, the ring centred in a fixed plot area
  const center = spec.kind === "doughnut" && spec.centerLabel && spec.centerLabel.length ? spec.centerLabel : null;
  let titleXml = "", plotLayout = layout;
  if (center) {
    const W = (size && size.widthPx) || 400, H = (size && size.heightPx) || 300;
    const box = { x: 0.04, y: 0.04, w: spec.legend ? 0.66 : 0.92, h: 0.92 };
    const pt = s => s.size || spec.font.size || 9;
    const tw = Math.max(...center.map(l => l.reduce((w, s) => w + String(s.text).length * pt(s) * 0.62, 0))) + 10;
    const th = center.reduce((h, l) => h + Math.max(...l.map(pt)) * 4 / 3 * 1.25, 0) + 6;
    const cx = box.x + box.w / 2, cy = box.y + box.h / 2;
    titleXml = `<c:title><c:tx><c:rich><a:bodyPr wrap="none"/><a:lstStyle/>` + center.map(l => `<a:p><a:pPr algn="ctr"><a:defRPr/></a:pPr>` +
        l.map(s => `<a:r>${runProps({ name: s.font || spec.font.name, size: pt(s), color: s.color || spec.font.color }, { bold: s.bold }, "rPr")}` +
          `<a:t>${esc(s.text)}</a:t></a:r>`).join("") + `</a:p>`).join("") + `</c:rich></c:tx>` +
      `<c:layout><c:manualLayout><c:xMode val="edge"/><c:yMode val="edge"/>` +
      `<c:x val="${Math.max(0, cx - tw / 2 / W).toFixed(4)}"/><c:y val="${Math.max(0, cy - th / 2 / H).toFixed(4)}"/></c:manualLayout></c:layout>` +
      `<c:overlay val="1"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:title>`;
    plotLayout = `<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/>` +
      `<c:x val="${box.x}"/><c:y val="${box.y}"/><c:w val="${box.w}"/><c:h val="${box.h}"/></c:manualLayout></c:layout>`;
  }
  // a sparkline (a short chart, no value axis, no legend – a KPI tile's trend): the plot fills the chart above its
  // category labels, as in Tableau; Excel's own layout keeps a margin that leaves it a sliver
  const H0 = size && size.heightPx;
  if (!g && !center && !frame && H0 && H0 <= SPARKLINE_MAX_PX && spec.valueAxisHidden && !spec.legend && !spec.title &&
      /^(line|area|bar|combo)$/.test(spec.kind)) {
    const W = (size && size.widthPx) || 200, f = v => Math.max(0, Math.min(1, v)).toFixed(4);
    const labels = spec.categoryAxisHidden ? 0 : (spec.font.size || 9) * 4 / 3 * 1.3 + 4 + (spec.categoryTitle ? (spec.font.size || 9) * 4 / 3 * 1.3 : 0);
    plotLayout = `<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/>` +
      `<c:x val="${f(4 / W)}"/><c:y val="${f(2 / H0)}"/><c:w val="${f((W - 8) / W)}"/><c:h val="${f((H0 - 2 - labels) / H0)}"/></c:manualLayout></c:layout>`;
  }
  if (g) {
    // every pane's plot area at the same place inside its chart, so the grid lines up like Tableau's panes
    const W = (size && size.widthPx) || 200, H = (size && size.heightPx) || 150, i = gridInset(g, spec);
    const f = v => Math.max(0, Math.min(1, v)).toFixed(4);
    plotLayout = `<c:layout><c:manualLayout><c:layoutTarget val="inner"/><c:xMode val="edge"/><c:yMode val="edge"/>` +
      `<c:x val="${f(i.l / W)}"/><c:y val="${f(i.t / H)}"/><c:w val="${f((W - i.l - i.r) / W)}"/><c:h val="${f((H - i.t - i.b) / H)}"/></c:manualLayout></c:layout>`;
    // the row's header on one line left of its pane (turned or not, as in the workbook), cut short with ".." like Tableau's
    if (g.col === 0 && spec.rowTitle) {
      const charPx = (spec.font.size || 9) * 4 / 3 * 0.55;
      const room = Math.floor((g.rowVertical ? H - i.t - i.b : i.l - (spec.valueAxisHidden ? 0 : GRID_AXIS) - 8) / charPx);
      if (spec.rowTitle.length > room) spec.rowTitle = spec.rowTitle.slice(0, Math.max(1, room - 2)).trimEnd() + "..";
    }
    if (g.row === 0 && g.colLabel) {
      const tw = String(g.colLabel).length * (spec.font.size || 9) * 4 / 3 * 0.55 + 8;
      titleXml = title(g.colLabel, spec.font, false).replace("<c:overlay", `<c:layout><c:manualLayout><c:xMode val="edge"/><c:yMode val="edge"/>` +
        `<c:x val="${f((i.l + (W - i.l - i.r) / 2 - tw / 2) / W)}"/><c:y val="0"/></c:manualLayout></c:layout><c:overlay`);
    }
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">` +
    `<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/>` +
    `<c:chart>${titleXml}<c:autoTitleDeleted val="${titleXml ? 0 : 1}"/><c:plotArea>${plotLayout}${plotAreaXml(spec, refs, plot, axisPlot)}` +
    `<c:spPr><a:noFill/>${g ? line("D9D9D9", 6350) : "<a:ln><a:noFill/></a:ln>"}</c:spPr></c:plotArea>${legend}` +
    `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `<c:spPr>${spec.background === null ? "<a:noFill/>" : solid(spec.background || "FFFFFF")}<a:ln><a:noFill/></a:ln></c:spPr>${txPr(spec.font)}` +
    `<c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/>` +
    `<c:pageSetup/></c:printSettings></c:chartSpace>`;
}


/* ═══ charts/writer/chartex.js ════════════════════════════════════════════════════════════════════ */
/* Excel 2016+ chartex parts (treemap) and their drawing anchors. */

/* Excel 2016+ "chartex" part – treemap. Category levels outer → inner in the spec,
 * inner (leaf) level first in the XML. Excel only accepts chartex formulas that are
 * hidden workbook names (_xlchart.v1.N), so refs here are those names, not ranges. */
/** @param {ChartSpec} spec @param {Pick<ChartRefs, "cat" | "series">} refs workbook names, not ranges @param {string} uid series GUID @returns {string} */
function chartExXml(spec, refs, uid) {
  const levels = spec.categories.levels;
  const n = levels.length ? levels[0].length : 0;
  const s = spec.series[0];
  const lvl = lv => `<cx:lvl ptCount="${n}">${lv.map((v, i) => `<cx:pt idx="${i}">${esc(v)}</cx:pt>`).join("")}</cx:lvl>`;
  const vals = s.values.map((v, i) => num(v) === null ? "" : `<cx:pt idx="${i}">${v}</cx:pt>`).join("");
  const tile = c => `<cx:spPr>${solid(c)}${line("FFFFFF", 12700)}</cx:spPr>`;
  const font = spec.font || /** @type {ChartSpec["font"]} */ ({});
  const labelPr = `<cx:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${Math.round((font.size || 9) * 100)}">` +
    `<a:latin typeface="${esc(font.name || "Arial")}"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></cx:txPr>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<cx:chartSpace xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:cx="${NS.cx}">` +
    `<cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:f>${esc(refs.cat)}</cx:f>${levels.slice().reverse().map(lvl).join("")}</cx:strDim>` +
    `<cx:numDim type="size"><cx:f>${esc(refs.series[0].val)}</cx:f><cx:lvl ptCount="${n}" formatCode="${esc(spec.numFmt || "General")}">${vals}</cx:lvl>` +
    `</cx:numDim></cx:data></cx:chartData>` +
    `<cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="treemap" uniqueId="${uid}">` +
    `<cx:tx><cx:txData><cx:f>${esc(refs.series[0].tx)}</cx:f><cx:v>${esc(s.name)}</cx:v></cx:txData></cx:tx>` +
    tile(s.color) +
    (s.pointColors || []).map((c, i) => c ? `<cx:dataPt idx="${i}">${tile(c)}</cx:dataPt>` : "").join("") +
    (s.labels ? `<cx:dataLabels pos="inEnd">` +
      (s.labelParts && s.labelParts.value ? `<cx:numFmt formatCode="${esc(spec.numFmt || "General")}" sourceLinked="0"/>` : "") +
      `${labelPr}<cx:visibility seriesName="0" categoryName="1" value="${s.labelParts && s.labelParts.value ? 1 : 0}"/>` +
      (s.labelParts && s.labelParts.value ? `<cx:separator>${"\n"}</cx:separator>` : "") + `</cx:dataLabels>` : "") +
    `<cx:dataId val="0"/><cx:layoutPr><cx:parentLabelLayout val="${levels.length > 1 ? "banner" : "none"}"/></cx:layoutPr>` +
    `</cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart>` +
    `<cx:spPr>${spec.background === null ? "<a:noFill/>" : solid(spec.background || "FFFFFF")}<a:ln><a:noFill/></a:ln></cx:spPr></cx:chartSpace>`;
}

/* chartex frames sit in mc:AlternateContent; older Excel shows the fallback rectangle */
/** @param {ChartJob} chart @param {number} id @param {string} rid */
function anchorExXml(chart, id, rid) {
  const cx = Math.max(1, Math.round(chart.widthPx * EMU_PER_PX));
  const cy = Math.max(1, Math.round(chart.heightPx * EMU_PER_PX));
  const name = esc(chart.name || "Chart " + id);
  return `<xdr:oneCellAnchor><xdr:from><xdr:col>${chart.col}</xdr:col><xdr:colOff>${Math.round((chart.colOffPx || 0) * EMU_PER_PX)}</xdr:colOff>` +
    `<xdr:row>${chart.row}</xdr:row><xdr:rowOff>${Math.round((chart.rowOffPx || 0) * EMU_PER_PX)}</xdr:rowOff></xdr:from>` +
    `<xdr:ext cx="${cx}" cy="${cy}"/>` +
    `<mc:AlternateContent xmlns:mc="${NS.mc}"><mc:Choice xmlns:cx1="${NS.cx1}" Requires="cx1">` +
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic>` +
    `<a:graphicData uri="${NS.cx}"><cx:chart xmlns:cx="${NS.cx}" xmlns:r="${NS.r}" r:id="${rid}"/></a:graphicData>` +
    `</a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback>` +
    `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr>` +
    `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
    `<a:solidFill><a:prstClr val="white"/></a:solidFill><a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln></xdr:spPr>` +
    `<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1100"/>` +
    `<a:t>This treemap needs Excel 2016 or later.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback></mc:AlternateContent>` +
    `<xdr:clientData/></xdr:oneCellAnchor>`;
}


/* ═══ charts/writer/chartex-style.js ══════════════════════════════════════════════════════════════ */
/* Excel 2016 default chart style / colour style parts required by chartex charts. */

/* Excel 2016 default chart style / colour style parts (chartex charts will not load without them) */
const CHARTEX_STYLE = "<cs:chartStyle xmlns:cs=\"http://schemas.microsoft.com/office/drawing/2012/chartStyle\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" id=\"410\"><cs:axisTitle><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val=\"bg1\"><a:lumMod val=\"65000\"/></a:schemeClr></a:solidFill><a:ln w=\"19050\"><a:solidFill><a:schemeClr val=\"bg1\"/></a:solidFill></a:ln></cs:spPr><cs:defRPr sz=\"900\"/></cs:axisTitle><cs:categoryAxis><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></cs:fontRef><cs:spPr><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"15000\"/><a:lumOff val=\"85000\"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr><cs:defRPr sz=\"900\"/></cs:categoryAxis><cs:chartArea mods=\"allowNoFillOverride allowNoLineOverride\"><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val=\"bg1\"/></a:solidFill><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"15000\"/><a:lumOff val=\"85000\"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr><cs:defRPr sz=\"1000\"/></cs:chartArea><cs:dataLabel><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\"/></cs:fontRef><cs:defRPr sz=\"900\"/></cs:dataLabel><cs:dataLabelCallout><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"dk1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val=\"lt1\"/></a:solidFill><a:ln><a:solidFill><a:schemeClr val=\"dk1\"><a:lumMod val=\"25000\"/><a:lumOff val=\"75000\"/></a:schemeClr></a:solidFill></a:ln></cs:spPr><cs:defRPr sz=\"900\"/><cs:bodyPr rot=\"0\" spcFirstLastPara=\"1\" vertOverflow=\"clip\" horzOverflow=\"clip\" vert=\"horz\" wrap=\"square\" lIns=\"36576\" tIns=\"18288\" rIns=\"36576\" bIns=\"18288\" anchor=\"ctr\" anchorCtr=\"1\"><a:spAutoFit/></cs:bodyPr></cs:dataLabelCallout><cs:dataPoint><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"><cs:styleClr val=\"auto\"/></cs:fillRef><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:ln w=\"19050\"><a:solidFill><a:schemeClr val=\"lt1\"/></a:solidFill></a:ln></cs:spPr></cs:dataPoint><cs:dataPoint3D><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"><cs:styleClr val=\"auto\"/></cs:fillRef><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></cs:spPr></cs:dataPoint3D><cs:dataPointLine><cs:lnRef idx=\"0\"><cs:styleClr val=\"auto\"/></cs:lnRef><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:ln w=\"28575\" cap=\"rnd\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:round/></a:ln></cs:spPr></cs:dataPointLine><cs:dataPointMarker><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"><cs:styleClr val=\"auto\"/></cs:fillRef><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:ln w=\"9525\"><a:solidFill><a:schemeClr val=\"lt1\"/></a:solidFill></a:ln></cs:spPr></cs:dataPointMarker><cs:dataPointMarkerLayout symbol=\"circle\" size=\"5\"/><cs:dataPointWireframe><cs:lnRef idx=\"0\"><cs:styleClr val=\"auto\"/></cs:lnRef><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:ln w=\"28575\" cap=\"rnd\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:round/></a:ln></cs:spPr></cs:dataPointWireframe><cs:dataTable><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></cs:fontRef><cs:spPr><a:ln w=\"9525\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"15000\"/><a:lumOff val=\"85000\"/></a:schemeClr></a:solidFill></a:ln></cs:spPr><cs:defRPr sz=\"900\"/></cs:dataTable><cs:downBar><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"dk1\"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val=\"dk1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></a:solidFill><a:ln w=\"9525\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></a:solidFill></a:ln></cs:spPr></cs:downBar><cs:dropLine><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"35000\"/><a:lumOff val=\"65000\"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:dropLine><cs:errorBar><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:errorBar><cs:floor><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef></cs:floor><cs:gridlineMajor><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"15000\"/><a:lumOff val=\"85000\"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:gridlineMajor><cs:gridlineMinor><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"15000\"/><a:lumOff val=\"85000\"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:gridlineMinor><cs:hiLoLine><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"75000\"/><a:lumOff val=\"25000\"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:hiLoLine><cs:leaderLine><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"35000\"/><a:lumOff val=\"65000\"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr></cs:leaderLine><cs:legend><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></cs:fontRef><cs:defRPr sz=\"900\"/></cs:legend><cs:plotArea mods=\"allowNoFillOverride allowNoLineOverride\"><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef></cs:plotArea><cs:plotArea3D mods=\"allowNoFillOverride allowNoLineOverride\"><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef></cs:plotArea3D><cs:seriesAxis><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></cs:fontRef><cs:spPr><a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"15000\"/><a:lumOff val=\"85000\"/></a:schemeClr></a:solidFill><a:round/></a:ln></cs:spPr><cs:defRPr sz=\"900\"/></cs:seriesAxis><cs:seriesLine><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:ln w=\"9525\" cap=\"flat\"><a:solidFill><a:srgbClr val=\"D9D9D9\"/></a:solidFill><a:round/></a:ln></cs:spPr></cs:seriesLine><cs:title><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></cs:fontRef><cs:defRPr sz=\"1400\"/></cs:title><cs:trendline><cs:lnRef idx=\"0\"><cs:styleClr val=\"auto\"/></cs:lnRef><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef><cs:spPr><a:ln w=\"19050\" cap=\"rnd\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"sysDash\"/></a:ln></cs:spPr></cs:trendline><cs:trendlineLabel><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></cs:fontRef><cs:defRPr sz=\"900\"/></cs:trendlineLabel><cs:upBar><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"dk1\"/></cs:fontRef><cs:spPr><a:solidFill><a:schemeClr val=\"lt1\"/></a:solidFill><a:ln w=\"9525\"><a:solidFill><a:schemeClr val=\"tx1\"><a:lumMod val=\"15000\"/><a:lumOff val=\"85000\"/></a:schemeClr></a:solidFill></a:ln></cs:spPr></cs:upBar><cs:valueAxis><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"><a:lumMod val=\"65000\"/><a:lumOff val=\"35000\"/></a:schemeClr></cs:fontRef><cs:defRPr sz=\"900\"/></cs:valueAxis><cs:wall><cs:lnRef idx=\"0\"/><cs:fillRef idx=\"0\"/><cs:effectRef idx=\"0\"/><cs:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></cs:fontRef></cs:wall></cs:chartStyle>";

const CHARTEX_COLORS = "<cs:colorStyle xmlns:cs=\"http://schemas.microsoft.com/office/drawing/2012/chartStyle\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" meth=\"cycle\" id=\"10\"><a:schemeClr val=\"accent1\"/><a:schemeClr val=\"accent2\"/><a:schemeClr val=\"accent3\"/><a:schemeClr val=\"accent4\"/><a:schemeClr val=\"accent5\"/><a:schemeClr val=\"accent6\"/><cs:variation/><cs:variation><a:lumMod val=\"60000\"/></cs:variation><cs:variation><a:lumMod val=\"80000\"/><a:lumOff val=\"20000\"/></cs:variation><cs:variation><a:lumMod val=\"80000\"/></cs:variation><cs:variation><a:lumMod val=\"60000\"/><a:lumOff val=\"40000\"/></cs:variation><cs:variation><a:lumMod val=\"50000\"/></cs:variation><cs:variation><a:lumMod val=\"70000\"/><a:lumOff val=\"30000\"/></cs:variation><cs:variation><a:lumMod val=\"70000\"/></cs:variation><cs:variation><a:lumMod val=\"50000\"/><a:lumOff val=\"50000\"/></cs:variation></cs:colorStyle>";


/* ═══ charts/writer/package.js ════════════════════════════════════════════════════════════════════ */
/* XLSX package surgery: drawing / chart parts, relationships, content types. */

/** <sheetPr> children in schema order (tabColor, outlinePr, pageSetUpPr): ExcelJS writes pageSetUpPr first
 * when a sheet has both fit-to-page and outline groups, and Excel then refuses to open the file.
 * Anything else inside <sheetPr> leaves it as it is. @param {string} xml @returns {string} */
function orderSheetPr(xml) {
  return xml.replace(/<sheetPr\b([^>]*)>([\s\S]*?)<\/sheetPr>/, (m, attrs, inner) => {
    const pick = tag => (inner.match(new RegExp(`<${tag}\\b[^>]*?(?:\\/>|>[\\s\\S]*?<\\/${tag}>)`)) || [""])[0];
    const out = pick("tabColor") + pick("outlinePr") + pick("pageSetUpPr");
    return out.length === inner.length ? `<sheetPr${attrs}>${out}</sheetPr>` : m;
  });
}

/**
 * Row groups the way Excel itself writes them. ExcelJS marks every grouped row collapsed="1" and leaves the
 * sheet's outlineLevelRow at 0 – Excel then reports "a problem with some content" and repairs the file. Excel's
 * own form: <sheetFormatPr outlineLevelRow> = the deepest level; the hidden detail rows carry only hidden +
 * outlineLevel; collapsed="1" sits on the summary row next to each hidden group (above it when summaryBelow="0").
 * @param {string} xml a worksheet part @returns {string}
 */
function fixRowOutlines(xml) {
  const rowTag = /<row\b[^>]*?(\/?)>/g;
  const levelOf = tag => +((tag.match(/\soutlineLevel="(\d+)"/) || [])[1] || 0);
  const numOf = tag => +((tag.match(/\sr="(\d+)"/) || [])[1] || 0);
  const tags = [...xml.matchAll(rowTag)].map(m => m[0]);
  const deepest = Math.max(0, ...tags.map(levelOf));
  if (!deepest) return xml;
  const above = /<outlinePr\b[^>]*\ssummaryBelow="(0|false)"/.test(xml);
  // hidden grouped rows → their summary rows (the row just before / after each unbroken run)
  const grouped = new Set(tags.filter(t => levelOf(t) && /\shidden="1"/.test(t)).map(numOf));
  const summaries = new Set();
  grouped.forEach(r => {
    if (above && !grouped.has(r - 1) && r > 1) summaries.add(r - 1);
    if (!above && !grouped.has(r + 1)) summaries.add(r + 1);
  });
  const present = new Set(tags.map(numOf));
  let out = xml.replace(rowTag, tag => {
    let t = tag.replace(/\scollapsed="(1|true)"/, "");
    if (summaries.has(numOf(t))) t = t.replace(/^<row\b/, '<row collapsed="1"');
    return t;
  });
  // a summary row with no cells has no <row> yet: an empty one, in row order
  const missing = [...summaries].filter(r => !present.has(r)).sort((a, b) => a - b);
  missing.forEach(r => {
    const next = [...out.matchAll(rowTag)].find(m => numOf(m[0]) > r);
    const empty = `<row r="${r}" collapsed="1"/>`;
    out = next ? out.slice(0, next.index) + empty + out.slice(next.index)
               : out.replace("</sheetData>", empty + "</sheetData>");
  });
  return out.replace(/<sheetFormatPr\b[^>]*?\/?>/, tag =>
    /\soutlineLevelRow="/.test(tag) ? tag.replace(/\soutlineLevelRow="\d+"/, ` outlineLevelRow="${deepest}"`)
                                    : tag.replace(/^<sheetFormatPr\b/, `<sheetFormatPr outlineLevelRow="${deepest}"`));
}

/**
 * Rich text with an empty run – e.g. a KPI line "vs Prev = 5 | ▲ 12%" whose arrow field is empty on this row –
 * makes Excel report "a problem with some content" ("Repaired Records: String properties from
 * /xl/sharedStrings.xml"): every run needs at least one character. Empty runs are dropped from every cell;
 * a cell left with no run at all keeps its style and no value. Runs of spaces stay (they are visible).
 * @param {any} workbook an ExcelJS workbook
 */
function dropEmptyRichTextRuns(workbook) {
  let dropped = 0;
  workbook.eachSheet(ws => ws.eachRow({ includeEmpty: false }, row => row.eachCell({ includeEmpty: false }, cell => {
    if (cell.master && cell.master !== cell) return;                    // merged: only the top-left cell owns the value
    const v = cell.value;
    if (!v || typeof v !== "object" || !Array.isArray(v.richText)) return;
    const runs = v.richText.filter(r => r && r.text !== undefined && r.text !== null && String(r.text) !== "");
    if (runs.length === v.richText.length) return;
    dropped += v.richText.length - runs.length;
    cell.value = runs.length ? { ...v, richText: runs } : null;
  })));
  if (dropped) console.log(`[Export] ${dropped} empty rich-text run(s) removed (Excel rejects them)`);
}

/* -----------------------------------------------------------------------------
 * Filter cards: a dashboard filter control as its title and current selection, at its dashboard spot.
 * --------------------------------------------------------------------------- */
const FILTER_CARD_MAX_VALUES = 3;
/** time limit for asking Tableau a filtered field's full list of values (filter cards' "(All)") */
const FILTER_DOMAIN_TIMEOUT_MS = (typeof window !== "undefined" && /** @type {any} */ (window).TFX_FILTER_DOMAIN_TIMEOUT_MS) || 10000;

/** What a filter control shows as its selection – "(All)", "East, West", "1/1/2023 – 12/31/2024",
 *  "Last 3 months" … – or null when it cannot be told (that card is left out). @param {any} f a Tableau filter */
function filterCardText(f) {
  const day = v => String(v == null ? "" : v).replace(/\s+12:00:00\s*AM$|\s+00:00:00$/i, "");
  const fv = v => (v ? day(v.formattedValue !== undefined && v.formattedValue !== null ? v.formattedValue : v.value) : "");
  const type = String(f.filterType || "").toLowerCase().replace(/[^a-z]/g, "");
  if (type === "categorical") {
    if (f.isAllSelected === true) return "(All)";
    const vals = (f.appliedValues || []).map(fv).filter(x => x !== "");
    if (!vals.length) return f.isExcludeMode ? "(All)" : "(None)";
    const shown = vals.slice(0, FILTER_CARD_MAX_VALUES).join(", ") +
                  (vals.length > FILTER_CARD_MAX_VALUES ? ` (+${vals.length - FILTER_CARD_MAX_VALUES} more)` : "");
    return f.isExcludeMode ? `All except ${shown}` : shown;
  }
  if (type === "range") {
    const a = fv(f.minValue), b = fv(f.maxValue);
    return a || b ? `${a} – ${b}` : "(All)";
  }
  if (type === "relativedate") {
    const unit = String(f.periodType || "").toLowerCase().replace(/s$/, "");
    const n = Number(f.rangeN) || 0, r = String(f.rangeType || "").toLowerCase();
    if (!unit) return null;
    if (r === "current") return `This ${unit}`;
    if (r === "todate") return `${unit[0].toUpperCase()}${unit.slice(1)} to date`;
    if (r === "last") return `Last ${unit}`;
    if (r === "next") return `Next ${unit}`;
    if (r === "lastn") return `Last ${n} ${unit}s`;
    if (r === "nextn") return `Next ${n} ${unit}s`;
    return null;
  }
  return null;
}

/**
 * The dashboard's filter controls as cards: the twb tells which worksheet and field each control filters
 * (and its custom title); the worksheet's filters tell the selection. A control whose filter can't be found
 * is left out – a card never shows a guess.
 * @param {any} dashboard @param {any[]} sheets @param {any} fmtModel @param {Map<string, any>} layoutMap
 * @returns {Promise<ExportItem[]>}
 */
async function buildFilterCards(dashboard, sheets, fmtModel, layoutMap) {
  const dash = fmtModel && fmtModel.dashboards && fmtModel.dashboards[dashboard.name];
  if (!dash) return [];
  const shown = new Set((dashboard.objects || []).filter(o => OBJECT_KIND[o.type] === "filter" && o.isVisible !== false).map(o => String(o.id)));
  const norm = x => String(x || "").toLowerCase().replace(/[^a-z0-9]/g, "");
  const seen = new Set(), cards = [], filtersOf = new Map(), domains = new Map();
  for (const z of dash.zones) {
    if (z.type !== "filter" || !z.param || z.hidden || seen.has(String(z.id))) continue;
    seen.add(String(z.id));
    const layout = layoutMap.get(`filter:${z.id}`);
    if (!shown.has(String(z.id)) || !layout) continue;                 // not on screen (or hidden by the user)
    const ref = tfParseFieldRef(z.param);
    if (!ref) continue;
    const ws = sheets.find(s => s.name === z.name);
    if (!ws || typeof ws.getFiltersAsync !== "function") continue;
    if (!filtersOf.has(ws.name)) {
      try {
        const raw = (await ws.getFiltersAsync()) || [];
        const all = await allSelectedFields(raw, p => p, domains);
        filtersOf.set(ws.name, raw.map(f => all.has(f.fieldName) ? { ...plainFilter(f), isAllSelected: true } : f));
      } catch (e) { filtersOf.set(ws.name, []); }
    }
    const filters = filtersOf.get(ws.name);
    const info = tfFieldInfo(fmtModel, ref);
    const names = [info && info.caption, ref.name].filter(Boolean).map(norm);
    let f = filters.find(x => names.includes(norm(x.fieldName)));
    if (!f) {                                                            // a date part: "MONTH(Date)" for [tmn:Date]
      const near = filters.filter(x => names.some(n => n && norm(x.fieldName).includes(n)));
      if (near.length === 1) f = near[0];
    }
    if (!f) { console.log(`[Filters] "${z.name}" – no filter on ${ref.name}: card left out`); continue; }
    const text = filterCardText(f);
    if (text === null) { console.log(`[Filters] ${f.fieldName}: a ${f.filterType} filter – card left out`); continue; }
    // the card's title: its custom title (Edit Title…), else the field's name; none when the title is hidden
    const sheetRules = (fmtModel.sheets[z.name] && fmtModel.sheets[z.name].style && fmtModel.sheets[z.name].style.rules) || {};
    const custom = (sheetRules["quick-filter"] || []).find(r => r.attr === "title" && r.field && tfSameField(tfParseFieldRef(r.field), ref));
    const title = !z.showTitle ? "" : (custom && custom.value) || f.fieldName || (info && info.caption) || ref.name;
    cards.push({ type: "filterValue", isFilterCard: true, name: `filter:${z.id}`, visualName: title || f.fieldName,
                 filterName: title, values: [text], layout, rowCount: 2, filter: f, field: ref, sheetName: ws.name, mode: z.mode || null });
  }
  if (cards.length) console.log(`[Filters] ${cards.length} filter card(s): ${cards.map(c => `${c.visualName} = ${c.values[0]}`).join(" | ")}`);
  return cards;
}

/** the package with every worksheet's <sheetPr> in schema order and row groups in Excel's own form
 * (unchanged buffer when nothing moves) @param {ArrayBuffer | Uint8Array} buffer */
async function fixSheetProperties(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  let changed = false;
  for (const name of Object.keys(zip.files).filter(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))) {
    const xml = await zip.file(name).async("string");
    const fixed = fixRowOutlines(orderSheetPr(xml));
    if (fixed !== xml) { zip.file(name, fixed); changed = true; }
  }
  return changed ? zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }) : buffer;
}

/** @param {ChartJob} chart @param {number} id @param {string} rid */
function anchorXml(chart, id, rid) {
  const cx = Math.max(1, Math.round(chart.widthPx * EMU_PER_PX));
  const cy = Math.max(1, Math.round(chart.heightPx * EMU_PER_PX));
  return `<xdr:oneCellAnchor><xdr:from><xdr:col>${chart.col}</xdr:col><xdr:colOff>${Math.round((chart.colOffPx || 0) * EMU_PER_PX)}</xdr:colOff>` +
    `<xdr:row>${chart.row}</xdr:row><xdr:rowOff>${Math.round((chart.rowOffPx || 0) * EMU_PER_PX)}</xdr:rowOff></xdr:from>` +
    `<xdr:ext cx="${cx}" cy="${cy}"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>` +
    `<xdr:cNvPr id="${id}" name="${esc(chart.name || "Chart " + id)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic>` +
    `<a:graphicData uri="${NS.c}"><c:chart xmlns:c="${NS.c}" xmlns:r="${NS.r}" r:id="${rid}"/></a:graphicData>` +
    `</a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>`;
}

function relTarget(relsXml, id) {
  const tag = (relsXml.match(new RegExp(`<Relationship\\b[^>]*\\sId="${id}"[^>]*>`)) || [])[0];
  return tag ? attr(tag, "Target") : null;
}

function nextRelId(relsXml) {
  const ids = [...relsXml.matchAll(/\sId="rId(\d+)"/g)].map(m => +m[1]);
  return "rId" + ((ids.length ? Math.max(...ids) : 0) + 1);
}

function addRel(relsXml, id, type, target) {
  return relsXml.replace("</Relationships>", `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`);
}

function resolve(baseDir, target) {
  if (target.startsWith("/")) return target.slice(1);
  const parts = (baseDir + target).split("/");
  const out = [];
  parts.forEach(p => { if (p === "..") out.pop(); else if (p && p !== ".") out.push(p); });
  return out.join("/");
}

function dirOf(path) { return path.replace(/[^/]+$/, ""); }

function relsPathOf(path) { return dirOf(path) + "_rels/" + path.replace(/^.*\//, "") + ".rels"; }

function nextFreeIndex(zip, prefix, ext) {
  let n = 1;
  while (zip.file(`${prefix}${n}${ext}`)) n++;
  return n;
}

function addOverride(ct, part, type) {
  if (ct.includes(`PartName="/${part}"`)) return ct;
  return ct.replace("</Types>", `<Override PartName="/${part}" ContentType="${type}"/></Types>`);
}

/* <drawing> must precede these elements in CT_Worksheet */
function insertDrawingTag(sheetXml, rid) {
  const tag = `<drawing r:id="${rid}"/>`;
  const after = ["<legacyDrawing", "<legacyDrawingHF", "<drawingHF", "<picture", "<oleObjects", "<controls",
                 "<webPublishItems", "<tableParts", "<extLst", "</worksheet>"];
  for (const t of after) {
    const i = sheetXml.indexOf(t);
    if (i >= 0) return sheetXml.slice(0, i) + tag + sheetXml.slice(i);
  }
  return sheetXml;
}

/* opts: { sheetIndex, charts: [{ spec, refs, col, row, widthPx, heightPx, name }] } */
/**
 * Adds the charts to the dashboard sheet of a written XLSX.
 * @param {ArrayBuffer | Uint8Array} buffer workbook.xlsx.writeBuffer() output
 * @param {{ sheetIndex?: number, sheetName?: string, charts: ChartJob[] }} opts
 * @returns {Promise<ArrayBuffer | Uint8Array>}
 */
async function injectCharts(buffer, opts) {
  const charts = (opts.charts || []).filter(Boolean);
  if (!charts.length) return buffer;
  const zip = await JSZip.loadAsync(buffer);
  const read = p => zip.file(p) ? zip.file(p).async("string") : Promise.resolve(null);

  const wbXml = await read("xl/workbook.xml");
  const wbRels = await read("xl/_rels/workbook.xml.rels");
  const sheetTags = wbXml.match(/<sheet\b[^>]*>/g) || [];
  // the dashboard's sheet: by name when given (several dashboards, hidden data sheets in between), else by position
  const unescapeXml = v => String(v || "").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
  const sheetTag = opts.sheetName != null ? sheetTags.find(t => unescapeXml(attr(t, "name")) === opts.sheetName)
    : sheetTags[opts.sheetIndex || 0];
  if (!sheetTag) throw new Error("Dashboard sheet not found in workbook package");
  const sheetPath = resolve("xl/", relTarget(wbRels, attr(sheetTag, "r:id")));
  const sheetRelsPath = relsPathOf(sheetPath);
  let sheetXml = await read(sheetPath);
  let sheetRels = (await read(sheetRelsPath)) || EMPTY_RELS;
  let ct = await read("[Content_Types].xml");

  if (!/<worksheet\b[^>]*\sxmlns:r=/.test(sheetXml)) {
    sheetXml = sheetXml.replace(/<worksheet\b/, `<worksheet xmlns:r="${NS.r}"`);
  }

  // re-use the drawing ExcelJS created for images, else create one
  let drawingPath, drawingXml, drawingRels;
  const existingRid = (sheetXml.match(/<drawing\b[^>]*r:id="([^"]+)"/) || [])[1];
  if (existingRid) {
    drawingPath = resolve(dirOf(sheetPath), relTarget(sheetRels, existingRid));
    drawingXml = await read(drawingPath);
    drawingRels = (await read(relsPathOf(drawingPath))) || EMPTY_RELS;
  } else {
    const n = nextFreeIndex(zip, "xl/drawings/drawing", ".xml");
    drawingPath = `xl/drawings/drawing${n}.xml`;
    drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}"></xdr:wsDr>`;
    drawingRels = EMPTY_RELS;
    const rid = nextRelId(sheetRels);
    sheetRels = addRel(sheetRels, rid, REL_DRAWING, `../drawings/drawing${n}.xml`);
    sheetXml = insertDrawingTag(sheetXml, rid);
    ct = addOverride(ct, drawingPath, CT_DRAWING);
    zip.file(drawingPath, drawingXml);                    // claim the name now: charts' label parts are drawings too
  }
  if (!/<xdr:wsDr\b[^>]*\sxmlns:a=/.test(drawingXml)) {
    drawingXml = drawingXml.replace(/<xdr:wsDr\b/, `<xdr:wsDr xmlns:a="${NS.a}"`);
  }

  const ids = [...drawingXml.matchAll(/<xdr:cNvPr\b[^>]*\sid="(\d+)"/g)].map(m => +m[1]);
  let shapeId = (ids.length ? Math.max(...ids) : 1) + 1;
  let anchors = "";
  // chartex data ranges → hidden defined names, as Excel writes them
  const usedNames = [...wbXml.matchAll(/name="_xlchart\.v1\.(\d+)"/g)].map(m => +m[1]);
  let nameIndex = usedNames.length ? Math.max(...usedNames) + 1 : 0;
  const newNames = [];
  const defineName = ref => {
    const name = `_xlchart.v1.${nameIndex++}`;
    newNames.push(`<definedName name="${name}" hidden="1">${esc(ref)}</definedName>`);
    return name;
  };
  // a switched chart's names (writeSwitchedChart): its ranges picked by the switch
  charts.forEach(chart => (chart.names || []).forEach(n => newNames.push(`<definedName name="${n.name}">${esc(n.formula)}</definedName>`)));
  for (const chart of charts) {
    const rid = nextRelId(drawingRels);
    if (chart.spec.kind === "treemap") {
      const n = nextFreeIndex(zip, "xl/charts/chartEx", ".xml");
      const chartPath = `xl/charts/chartEx${n}.xml`;
      const uid = `{6F1D3C1E-0000-4000-8000-${String(n).padStart(12, "0")}}`;
      const named = { cat: defineName(chart.refs.cat),
                      series: chart.refs.series.map(s => ({ tx: defineName(s.tx), val: defineName(s.val) })) };
      zip.file(chartPath, chartExXml(chart.spec, named, uid));
      ct = addOverride(ct, chartPath, CT_CHARTEX);
      const xmlHead = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;
      const styleN = nextFreeIndex(zip, "xl/charts/style", ".xml");
      zip.file(`xl/charts/style${styleN}.xml`, xmlHead + CHARTEX_STYLE);
      zip.file(`xl/charts/colors${styleN}.xml`, xmlHead + CHARTEX_COLORS);
      ct = addOverride(ct, `xl/charts/style${styleN}.xml`, CT_CHARTSTYLE);
      ct = addOverride(ct, `xl/charts/colors${styleN}.xml`, CT_CHARTCOLORS);
      zip.file(relsPathOf(chartPath), addRel(addRel(EMPTY_RELS, "rId1", REL_CHARTSTYLE, `style${styleN}.xml`),
                                             "rId2", REL_CHARTCOLORS, `colors${styleN}.xml`));
      drawingRels = addRel(drawingRels, rid, REL_CHARTEX, `../charts/chartEx${n}.xml`);
      anchors += anchorExXml(chart, shapeId++, rid);
      continue;
    }
    const n = nextFreeIndex(zip, "xl/charts/chart", ".xml");
    const chartPath = `xl/charts/chart${n}.xml`;
    // a pie / donut arrives as finished chart XML, with its label shapes and tooltip wedges
    zip.file(chartPath, chart.xml || chartXml(chart.spec, chart.refs, chart));
    ct = addOverride(ct, chartPath, CT_CHART);
    drawingRels = addRel(drawingRels, rid, REL_CHART, `../charts/chart${n}.xml`);
    if (chart.shapes) {                                    // labels / centre text: the chart's own drawing part
      const d = nextFreeIndex(zip, "xl/drawings/drawing", ".xml");
      const shapesPath = `xl/drawings/drawing${d}.xml`;
      zip.file(shapesPath, chart.shapes);
      ct = addOverride(ct, shapesPath, CT_USER_SHAPES);
      zip.file(relsPathOf(chartPath), addRel(EMPTY_RELS, "rId1", REL_USER_SHAPES, `../drawings/drawing${d}.xml`));
    }
    if (chart.tips && chart.tips.length) {                // tooltips: hyperlink ScreenTips on wedges over the slices
      const tipRids = chart.tips.map(t => {
        const tid = nextRelId(drawingRels);
        drawingRels = addRel(drawingRels, tid, REL_HYPERLINK, esc(t.link));
        return tid;
      });
      const ids = { next: shapeId };
      anchors += groupedAnchorXml(chart, rid, ids, tipRids);
      shapeId = ids.next;
      continue;
    }
    anchors += anchorXml(chart, shapeId++, rid);
  }
  // charts first: pictures ExcelJS placed (icons over a chart) stay in front of them
  drawingXml = drawingXml.replace(/(<xdr:wsDr\b[^>]*>)/, (m) => m + anchors);
  if (newNames.length) {                         // CT_Workbook: definedNames follow sheets
    const defs = newNames.join("");
    const wb = wbXml.includes("</definedNames>") ? wbXml.replace("</definedNames>", defs + "</definedNames>")
      : wbXml.includes("<definedNames/>") ? wbXml.replace("<definedNames/>", `<definedNames>${defs}</definedNames>`)
      : wbXml.replace("</sheets>", `</sheets><definedNames>${defs}</definedNames>`);
    zip.file("xl/workbook.xml", wb);
  }

  zip.file(sheetPath, sheetXml);
  zip.file(sheetRelsPath, sheetRels);
  zip.file(drawingPath, drawingXml);
  zip.file(relsPathOf(drawingPath), drawingRels);
  zip.file("[Content_Types].xml", ct);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}


/* ═══ charts/writer/index.js ══════════════════════════════════════════════════════════════════════ */
/* ══════════════════════════════════════════════════════════════════════════
 * EXCEL NATIVE CHART WRITER
 * ──────────────────────────────────────────────────────────────────────────
 * ExcelJS cannot author charts, so native charts are produced in two steps:
 *   1. writeChartData(ws, spec, row) – the chart's source data is written to a
 *      (hidden) data sheet with ExcelJS, so every chart stays editable and is
 *      linked to real cells.
 *   2. injectCharts(buffer, opts) – after workbook.xlsx.writeBuffer() the XLSX
 *      package is opened with JSZip and DrawingML chart parts are added to the
 *      dashboard sheet (re-using the drawing ExcelJS made for images, if any).
 *
 * Chart spec (produced by charts/model):
 *   { kind: "bar"|"line"|"area"|"pie"|"doughnut"|"scatter"|"bubble"|"combo"|"treemap",
 *     barDir: "col"|"bar", stacked, gapWidth, categories: { names:[], levels:[[…]] },
 *     series: [{ name, type, color (null = invisible), values|x+y(+size), pointColors, secondary,
 *                line, marker, markerSymbol, markerSize, labels, labelNumFmt }],
 *     numFmt, secondaryNumFmt, xNumFmt, valueTitle, secondaryTitle, valueMin, valueMax,
 *     boxPlot: { color } (line chart → up/down bars + high-low lines),
 *     categoryTitle, legend, gridlines, font: { name, size, color },
 *     axesHidden + packed (packed bubbles): axis ranges fitted to the chart size }
 * "treemap" is written as an Excel 2016+ chartex part with an older-Excel fallback shape.
 * Colours are 6-digit hex ("4E79A7"); ARGB ("FF4E79A7") is accepted too.
 * ══════════════════════════════════════════════════════════════════════════ */

const ExcelChartWriter = { writeChartData, injectCharts, chartXml, chartExXml, bubbleTips };


/* ═══ cloud/dashboards.js ═════════════════════════════════════════════════════════════════════════ */
/* =============================================================================
 * OTHER DASHBOARDS (Phase 3) — read dashboards the extension is NOT placed on.
 * -----------------------------------------------------------------------------
 * The Extensions API only sees its own dashboard. For the others we open each one
 * in an invisible Embedding API viz, in the viewer's OWN Tableau session (so row-level
 * security and permissions apply), read every worksheet's data, then close it.
 * The result has the shape the export pipeline uses for the live dashboard:
 *   { dashboard: { name, objects, worksheets, getParametersAsync },
 *     sheets: [{ name, getSummaryDataAsync, getVisualSpecificationAsync, getFiltersAsync, getSelectedMarksAsync }] }
 * Everything is read up front into plain snapshots, so nothing keeps talking to a closed viz.
 * ============================================================================= */

const TFX_WINDOW = typeof window !== "undefined" ? /** @type {any} */ (window) : {};
const VIZ_LOAD_TIMEOUT_MS = TFX_WINDOW.TFX_VIZ_TIMEOUT_MS || 90000;   // window override = tests only
// No "size known" from Tableau within this time = the viz is showing Tableau's sign-in page.
const SIGNIN_DETECT_MS = TFX_WINDOW.TFX_SIGNIN_DETECT_MS || 15000;
// How long the user gets to complete the one-time sign-in.
const SIGNIN_WAIT_MS = TFX_WINDOW.TFX_SIGNIN_WAIT_MS || 300000;
// Upper bound for one live Extensions API read of the current dashboard.
const LIVE_READ_TIMEOUT_MS = TFX_WINDOW.TFX_LIVE_READ_TIMEOUT_MS || 120000;
// How many other dashboards load at the same time (each one runs a full Tableau view in the browser).
const PARALLEL_DASHBOARDS = TFX_WINDOW.TFX_PARALLEL_DASHBOARDS || 3;
/** "12.3s" – for the timing lines in the console and the panel. */
const secs = ms => `${(ms / 1000).toFixed(1)}s`;
// the backend's answer time limit, and how long Export may wait for a workbook before it is released anyway
const BACKEND_TIMEOUT_MS = TFX_WINDOW.TFX_BACKEND_TIMEOUT_MS || 60000;
const WORKBOOK_WATCHDOG_MS = TFX_WINDOW.TFX_WORKBOOK_WATCHDOG_MS || 120000;
// Worksheets of the LIVE dashboard read at the same time (as fetchAllSheetsData does). Hidden (embedded) vizzes
// are always read one call at a time – Tableau mixes up overlapping reads inside one embedded session.
const SHEET_READ_CONCURRENCY = TFX_WINDOW.TFX_SHEET_READ_CONCURRENCY || 4;

/** Visible dashboards in Tableau tab order, from the workbook XML (<windows>). */
function dashboardTabOrder(xmlString) {
  const doc = new DOMParser().parseFromString(xmlString, "text/xml");
  const root = doc && doc.documentElement;
  if (!root) return [];
  const windows = getDirectChildByTag(root, "windows");
  const out = [];
  if (windows) {
    for (let i = 0; i < windows.childNodes.length; i++) {
      const w = windows.childNodes[i];
      if (w.nodeType === 1 && w.tagName === "window" && w.getAttribute("class") === "dashboard" &&
          w.getAttribute("hidden") !== "true") out.push(w.getAttribute("name"));
    }
  }
  if (out.length) return out;
  const dashboards = getDirectChildByTag(root, "dashboards");       // fallback: definition order
  if (!dashboards) return [];
  for (let i = 0; i < dashboards.childNodes.length; i++) {
    const d = dashboards.childNodes[i];
    if (d.nodeType === 1 && d.tagName === "dashboard") out.push(d.getAttribute("name"));
  }
  return out;
}

/** Dashboard size from the XML (fixed-size dashboards), so the hidden viz renders at its real size. */
function dashboardPixelSize(xmlString, name) {
  const doc = new DOMParser().parseFromString(xmlString, "text/xml");
  const d = [...doc.getElementsByTagName("dashboard")].find(x => x.getAttribute("name") === name);
  const size = d && getDirectChildByTag(d, "size");
  const num = k => (size && +size.getAttribute(k)) || 0;
  return { width: num("maxwidth") || num("minwidth") || 1200, height: num("maxheight") || num("minheight") || 900 };
}

/* -----------------------------------------------------------------------------
 * The embed room: a same-origin child frame that holds Tableau's Embedding API and every hidden /
 * sign-in viz. Loaded into this page, the Embedding API also answers the messages meant for the
 * Extensions API – after the first multi-dashboard export, every read of the live dashboard went
 * unanswered (the second export hung on "Choose where to save…"). In its own frame it hears only
 * its vizzes. The room is removed after each export, so the next one starts clean.
 * --------------------------------------------------------------------------- */
/** @type {{ frame: HTMLIFrameElement, ready: Promise<Window>, lib?: Promise<void> } | null} */
let EMBED_ROOM = null;
const ROOM_HIDDEN_CSS = "position:fixed;left:0;top:0;width:100vw;height:100vh;border:0;opacity:0;pointer-events:none;";
const ROOM_DEBUG_CSS = "position:fixed;left:0;top:0;width:100vw;height:100vh;border:0;z-index:9999;background:#fff;outline:3px dashed #b3261e;";

/** The room's window (created on first use). */
function embedRoom() {
  if (EMBED_ROOM && EMBED_ROOM.frame.isConnected) return EMBED_ROOM.ready;
  const frame = document.createElement("iframe");
  frame.id = "tfx_embed_room";
  frame.setAttribute("aria-hidden", "true");
  frame.setAttribute("title", "Tableau dashboards being read");
  frame.style.cssText = ROOM_HIDDEN_CSS;
  const ready = new Promise(resolve => frame.addEventListener("load", () => resolve(/** @type {Window} */ (frame.contentWindow)), { once: true }));
  frame.srcdoc = "<!doctype html><html><head><meta charset=\"utf-8\"><style>html,body{margin:0;padding:0;background:#fff}</style></head><body></body></html>";
  document.body.appendChild(frame);
  EMBED_ROOM = { frame, ready };
  return ready;
}

/** Tableau's Embedding API (pod-specific URL), loaded INSIDE the room. */
async function loadEmbeddingApi(libUrl) {
  const win = await embedRoom();
  if (win.customElements && win.customElements.get("tableau-viz")) return;
  const room = /** @type {any} */ (EMBED_ROOM);
  if (!room.lib) {
    room.lib = withTimeout(new Promise((resolve, reject) => {
      const script = win.document.createElement("script");
      script.type = "module";
      script.src = libUrl;
      script.onload = () => resolve(undefined);
      script.onerror = () => reject(new Error("the script did not load"));
      win.document.head.appendChild(script);
    }).then(() => win.customElements.whenDefined("tableau-viz")), 60000, "Loading Tableau's Embedding API")
      .catch(e => { room.lib = null; throw new Error(`could not load Tableau's Embedding API (${e.message})`); });
  }
  return room.lib;
}

/**
 * One-shot debug switch: localStorage.setItem("tfx_debugVizOnce", "1") shows the hidden dashboards
 * on screen for the NEXT export only, then switches itself off (see clearDebugVizSwitch).
 * The old persistent key "tfx_debugViz" is ignored and removed, so it can't stay on by accident.
 */
function debugVizOn() {
  try {
    localStorage.removeItem("tfx_debugViz");
    return localStorage.getItem("tfx_debugVizOnce") === "1";
  } catch (e) { return false; }
}
function clearDebugVizSwitch() {
  try { localStorage.removeItem("tfx_debugVizOnce"); localStorage.removeItem("tfx_debugViz"); } catch (e) { /* ignore */ }
}

/**
 * Host for the hidden vizzes. It sits INSIDE the visible area but fully transparent:
 * Chrome pauses rendering of cross-origin frames that are off-screen, and a paused
 * Tableau viz never becomes interactive — so "off-screen" is not an option.
 */
async function hiddenVizHost() {
  const win = await embedRoom();
  // on screen but transparent: Chrome pauses frames it can't see, and a paused viz never finishes loading
  /** @type {any} */ (EMBED_ROOM).frame.style.cssText = debugVizOn() ? ROOM_DEBUG_CSS : ROOM_HIDDEN_CSS;
  return win.document.body;
}

/**
 * Open one view invisibly; resolves with the <tableau-viz> once it's interactive.
 * Tracks Tableau's earlier "size known" event so a timeout can say WHICH stage failed:
 * never loaded (sign-in / cookies / embedding blocked) vs loaded but never finished drawing.
 */
async function openHiddenViz(embedUrl, size) {
  const host = await hiddenVizHost();
  return new Promise((resolve, reject) => {
    const viz = /** @type {any} */ (host.ownerDocument.createElement("tableau-viz"));
    viz.setAttribute("src", embedUrl);
    viz.setAttribute("toolbar", "hidden");
    viz.setAttribute("hide-tabs", "");
    viz.setAttribute("width", String(size.width));
    viz.setAttribute("height", String(size.height));
    // Parallel vizzes overlap at the top-left so every one stays inside the visible area –
    // Chrome pauses frames that are off-screen, and a paused viz never finishes loading.
    if (!debugVizOn()) viz.style.cssText = "position:absolute;left:0;top:0;";
    let sizeKnown = false;
    viz.addEventListener("firstvizsizeknown", () => { sizeKnown = true; });
    // Signed-out embeds never report a size (Tableau shows its sign-in page instead).
    const signInTimer = setTimeout(() => {
      if (sizeKnown) return;
      clearTimeout(timer);
      viz.remove();
      const err = /** @type {any} */ (new Error("needs Tableau sign-in"));
      err.code = "SIGN_IN_NEEDED";
      reject(err);
    }, SIGNIN_DETECT_MS);
    const timer = setTimeout(() => {
      clearTimeout(signInTimer);
      viz.remove();
      const secs = VIZ_LOAD_TIMEOUT_MS / 1000;
      reject(new Error(sizeKnown
        ? `loaded but never finished drawing within ${secs}s (browser paused it?)`
        : `never loaded within ${secs}s — usually a Tableau sign-in/cookie problem, or embedding blocked for this site`));
    }, VIZ_LOAD_TIMEOUT_MS);
    viz.addEventListener("firstinteractive", () => { clearTimeout(timer); clearTimeout(signInTimer); resolve(viz); });
    viz.addEventListener("vizloaderror", e => {
      clearTimeout(timer);
      clearTimeout(signInTimer);
      viz.remove();
      let detail = e && e.detail;
      try { detail = JSON.stringify(detail); } catch (x) { /* keep as is */ }
      reject(new Error("Tableau refused to open it: " + detail));
    });
    host.appendChild(viz);
  });
}

/* ── plain copies: only what the export uses, so nothing keeps a reference to a closed viz ──────── */
/** A Date made in another frame (the embed room) fails `instanceof Date` here – rebuilt as this frame's Date. */
const ownRealm = x => (x && typeof x === "object" && !(x instanceof Date) && Object.prototype.toString.call(x) === "[object Date]")
  ? new Date(x.getTime()) : x;
const plainValue = v => (v ? { value: ownRealm(v.value), nativeValue: ownRealm(v.nativeValue), formattedValue: v.formattedValue } : v);

/** a DataTable (summary data, selected marks): columns incl. fieldId (pies match measures by it) */
function plainSummary(t) {
  return {
    columns: (t.columns || []).map(c => ({ fieldName: c.fieldName, fieldId: c.fieldId, dataType: c.dataType, index: c.index,
                                           isReferenced: c.isReferenced })),
    data: (t.data || []).map(r => r.map(plainValue)),
  };
}
function plainFilter(f) {
  const fv = v => (v ? { formattedValue: v.formattedValue, value: ownRealm(v.value) } : v);
  return { fieldName: f.fieldName, filterType: f.filterType,
           appliedValues: (f.appliedValues || []).map(fv), minValue: fv(f.minValue), maxValue: fv(f.maxValue),
           isAllSelected: f.isAllSelected, isExcludeMode: f.isExcludeMode,           // filter cards: "(All)", "All except …"
           periodType: f.periodType, rangeType: f.rangeType, rangeN: f.rangeN };
}
function plainParameter(p) {
  return { name: p.name, currentValue: p.currentValue ? { value: ownRealm(p.currentValue.value), formattedValue: p.currentValue.formattedValue } : null };
}
/** a DashboardObject: what the layout uses (the id is the zone id in the workbook – text boxes / images need it) */
function plainObject(o) {
  return { id: o.id, type: o.type, name: o.name, isVisible: o.isVisible !== false, isFloating: !!o.isFloating,
           position: o.position ? { x: o.position.x, y: o.position.y } : undefined,
           size: o.size ? { width: o.size.width, height: o.size.height } : undefined };
}

async function readWorksheetSummary(ws) {
  if (typeof ws.getSummaryDataReaderAsync === "function") {
    const reader = await ws.getSummaryDataReaderAsync(undefined, { ignoreSelection: true });
    try { return plainSummary(await reader.getAllPagesAsync()); }
    finally { try { await reader.releaseAsync(); } catch (e) { /* ignore */ } }
  }
  return plainSummary(await ws.getSummaryDataAsync({ ignoreSelection: true }));
}

/** a worksheet stand-in answering from what was read (same methods the export calls on a live worksheet) */
function snapshotSheet(name, s) {
  const answer = (value, error) => async () => { if (error) throw error; return value; };
  return {
    name,
    getSummaryDataAsync: answer(s.summary, s.summaryError),
    getVisualSpecificationAsync: answer(s.visualSpec, s.visualSpec ? null : (s.visualSpecError || new Error("getVisualSpecificationAsync is unavailable"))),
    getFiltersAsync: async () => s.filters || [],
    getSelectedMarksAsync: async () => ({ data: s.selected || [] }),
    // slicers read the rows of a filtered field's data source: the live worksheet's own (this dashboard stays
    // open), or the copies read while a hidden view was open (another dashboard: its view is closed by now)
    ...(s.live && typeof s.live.getDataSourcesAsync === "function" ? { getDataSourcesAsync: () => s.live.getDataSourcesAsync() }
      : s.dataSources ? { getDataSourcesAsync: async () => s.dataSources } : {})
  };
}

/**
 * Copies of a worksheet's data sources that have one of its filtered fields – read now, while a hidden view is open
 * (it is closed before the slicers are made). Every logical table's rows (at most FORMAT_CONFIG.slicerMaxRows + 1
 * each, so "too many rows" is still told): of related tables the slicers pick the one that holds the field
 * (readSlicerSource). Each data source is read once per dashboard (cache). Never throws: no copy → the filter card
 * stays a card.
 * @param {any} ws @param {any[]} filters @param {Map<string, Promise<any>>} cache @returns {Promise<any[] | null>}
 */
async function freezeDataSources(ws, filters, cache) {
  if (!FORMAT_CONFIG.slicers || typeof ws.getDataSourcesAsync !== "function") return null;
  const filtered = (filters || []).filter(f => f && /^(categorical|range|relative-?date)$/i.test(String(f.filterType || "")) &&
                                                !/^Measure (Names|Values)$/.test(f.fieldName || "")).map(f => tfNorm(f.fieldName));
  if (!filtered.length) return null;
  let sources = [];
  try { sources = (await ws.getDataSourcesAsync()) || []; } catch (e) { return null; }
  const out = [];
  for (const ds of sources) {
    const fields = (ds.fields || []).map(f => ({ name: f.name }));
    if (!fields.some(f => filtered.includes(tfNorm(f.name)))) continue;      // no filtered field: not needed
    if (!cache.has(ds.id)) {
      cache.set(ds.id, (async () => {
        const tables = typeof ds.getLogicalTablesAsync === "function" ? await ds.getLogicalTablesAsync() : [];
        const max = FORMAT_CONFIG.slicerMaxRows + 1;
        /** @type {Map<string, { columns: any[], data: any[] }>} each table's rows (related tables: all of them) */
        const rows = new Map();
        for (const t of tables) {
          let table = null;
          if (typeof ds.getLogicalTableDataReaderAsync === "function") {
            const reader = await ds.getLogicalTableDataReaderAsync(t.id, 10000, { ignoreAliases: false });
            try { table = reader.totalRowCount > max ? { columns: [], data: new Array(reader.totalRowCount) } : await reader.getAllPagesAsync(max); }
            finally { await reader.releaseAsync(); }
          } else if (typeof ds.getLogicalTableDataAsync === "function") {
            table = await ds.getLogicalTableDataAsync(t.id, { ignoreAliases: false, maxRows: max });
          }
          if (table) rows.set(t.id, { columns: (table.columns || []).map(c => ({ fieldName: c.fieldName, dataType: c.dataType })), data: table.data || [] });
        }
        const plainTables = tables.map(t => ({ id: t.id, caption: t.caption }));
        return { id: ds.id, name: ds.name, fields,
                 getLogicalTablesAsync: async () => plainTables,
                 // the copy answers as Tableau would (the reader's row count tells "too many rows"); maxRows: a first look
                 getLogicalTableDataAsync: async (id, opts) => {
                   const table = rows.get(id);
                   if (!table) throw new Error("its rows were not read");
                   return opts && opts.maxRows && table.data.length > opts.maxRows ? { ...table, data: table.data.slice(0, opts.maxRows) } : table;
                 } };
      })().catch(e => { console.warn(`[Slicers] could not read "${ds.name}" while its dashboard was open:`, e.message); return null; }));
    }
    const copy = await cache.get(ds.id);
    if (copy) out.push(copy);
  }
  return out.length ? out : null;
}

/** Reads one worksheet completely: data, visual specification, filters, selected marks.
 *  wrap(promise, what) bounds each call (live dashboard) or passes it through. */
/**
 * Filters whose every value is selected – Tableau's "(All)". Tableau's isAllSelected is not always set
 * (it is often missing for embedded views, and not set when each value was checked one by one), so a
 * categorical filter without it is compared with its field's full list of values (asked once per field).
 * Any problem: the field is simply not marked (its card then lists the values, as before).
 * @param {any[]} raw Tableau filter objects @param {(p: Promise<any>, what: string) => Promise<any>} wrap
 * @param {Map<string, Promise<number>>} domains value counts per field, shared by the dashboard's worksheets
 * @returns {Promise<Set<string>>} the field names of filters with all values selected
 */
async function allSelectedFields(raw, wrap, domains) {
  const all = new Set();
  for (const f of raw || []) {
    if (!f || String(f.filterType || "").toLowerCase() !== "categorical" || f.isAllSelected === true || f.isExcludeMode) continue;
    const applied = (f.appliedValues || []).length;
    if (!applied || typeof f.getDomainAsync !== "function" || /^Measure (Names|Values)$/.test(f.fieldName || "")) continue;
    if (!domains.has(f.fieldName)) {
      domains.set(f.fieldName, withTimeout(wrap(f.getDomainAsync("database"), `Values of "${f.fieldName}"`), FILTER_DOMAIN_TIMEOUT_MS, "Filter values")
        .then(d => (d && Array.isArray(d.values) ? d.values.length : 0)).catch(() => 0));
    }
    const total = await domains.get(f.fieldName);
    if (total > 0 && applied >= total) all.add(f.fieldName);
  }
  return all;
}

async function readSheet(ws, wrap, plain, hidden = false, domains = new Map(), sources = new Map()) {
  const s = { summary: null, summaryError: null, visualSpec: null, visualSpecError: null, filters: [], selected: [], live: null, dataSources: null };
  const readFilters = async () => {
    try {
      const raw = (await wrap(ws.getFiltersAsync(), `Filters of "${ws.name}"`)) || [];
      const all = FORMAT_CONFIG.filterCards ? await allSelectedFields(raw, wrap, domains) : new Set();
      // a filter with every value selected carries isAllSelected (a plain copy then – Tableau's object is kept otherwise)
      s.filters = raw.map(f => all.has(f.fieldName) ? { ...plainFilter(f), isAllSelected: true } : plain ? plainFilter(f) : f);
    }
    catch (e) { /* export works without filter values */ }
  };
  if (hidden) {
    // Hidden on the dashboard (dynamic zone visibility): the export skips it anyway, so its data and visual
    // specification are not fetched – an empty table lets the export record it as hidden. Its filters are
    // still read: the export's filter notes come from every worksheet.
    s.summary = { columns: [], data: [] };
    await readFilters();
    return snapshotSheet(ws.name, s);
  }
  const readSummary = async () => {
    try { s.summary = plain ? await wrap(readWorksheetSummary(ws), `Reading "${ws.name}"`)
                            : await wrap(ws.getSummaryDataAsync({ ignoreSelection: true }), `Reading "${ws.name}"`); }
    catch (e) { s.summaryError = e; }
  };
  const readSpec = async () => {
    if (typeof ws.getVisualSpecificationAsync !== "function") return;
    try { s.visualSpec = await wrap(ws.getVisualSpecificationAsync(), `Visual specification of "${ws.name}"`); }
    catch (e) { s.visualSpecError = e; }
  };
  if (plain) {
    // A hidden (embedded) viz: ONE call at a time. Tableau keeps one data reader per embedded session –
    // overlapping calls there come back mixed up (wrong columns, internal-error), so nothing overlaps.
    await readSummary();
    await readSpec();
    await readFilters();
    // the rows of the data sources its filters use: read now, while the hidden view is open
    try { s.dataSources = await freezeDataSources(ws, s.filters, sources); } catch (e) { s.dataSources = null; }
    // a freshly opened hidden viz has nothing selected: its selected marks are always empty – not asked for
  } else {
    // The live dashboard (Extensions API): data + visual specification together, as fetchAllSheetsData does.
    await Promise.all([readSummary(), readSpec()]);
    await readFilters();
    s.live = ws;                                        // this dashboard stays open: its data sources are read when needed
    if (typeof ws.getSelectedMarksAsync === "function") {
      try {
        const marks = await wrap(ws.getSelectedMarksAsync(), `Selected marks of "${ws.name}"`);
        s.selected = ((marks && marks.data) || []).map(t => plain ? plainSummary(t) : t);
      } catch (e) { /* no selection outline */ }
    }
  }
  return snapshotSheet(ws.name, s);
}

/* ═══ export/view-switch.js ═══════════════════════════════════════════════════════════════════════ */
/* Views a parameter's buttons switch between (Gender / Race) as ONE Excel chart that a dropdown switches, with no
 * macros: the dropdown sits where Tableau's buttons are, and a cell of the chart data sheet holds its value (the
 * switch). Every view writes its data as usual; the chart reads them through defined names –
 * CHOOSE(MATCH(switch, {"Gender","Race"}, 0), the Gender view's range, the Race view's range) – so its categories
 * change with the switch. Each series of every view is a series of its own (it keeps its colour): the views not
 * picked read a column of #N/A, which Excel does not draw. A pie lists the slices of every view, each one #N/A
 * while its view is not picked. The title and the legend line above the chart are formulas of the switch too. */

/** what kind of chart a view is drawn as – views switch in one chart only when they are of one kind; null: it can't
 * @param {any} item @returns {string | null} */
function switchFamily(item) {
  const vm = item && item.visualModel;
  if (!vm || vm.gauge) return null;
  // a pie: one pane; an inner layer only when it is a disc of one colour (a donut's centre), which stays as it is
  if (vm.pie) return (vm.pie.inner && new Set(vm.pie.inner.points.map(p => String(p.color))).size > 1) ||
    vm.pie.rowKeys.length * vm.pie.colKeys.length !== 1 ? null : "pie";
  if (vm.renderer !== "excel-chart" || !vm.chartSpecs || vm.chartSpecs.length !== 1) return null;
  const s = /** @type {any} */ (vm.chartSpecs[0]);
  if (s.grid || s.waterfall || /^(scatter|bubble|treemap)$/.test(s.kind)) return null;
  return [s.kind, s.barDir, !!s.stacked, !!s.percent, s.dateAxis ? "date" : ""].join("|");
}

/**
 * A dashboard's switches: per parameter whose buttons switch views (the workbook's toggles), the view of each value –
 * one chart each, all of one kind (switchFamily), one of them shown – and the buttons (sheets with its parameter
 * action). Anything else (two charts per value, a table …) stays as withAlternateViews draws it.
 * @param {any[]} items @param {any} dashM @param {Set<string>} alternates @param {FormatModel} model
 * @param {string} dashName @param {(item: any) => string} zoneOf
 */
function findViewSwitches(items, dashM, alternates, model, dashName, zoneOf) {
  /** @type {Map<string, Map<string, Set<string>>>} parameter → value → its zones */
  const byParam = new Map();
  Object.entries((dashM && dashM.toggles) || {}).forEach(([zone, t]) => {
    if (!t.param || t.value === undefined) return;
    if (!byParam.has(t.param)) byParam.set(t.param, new Map());
    const values = byParam.get(t.param);
    if (!values.has(t.value)) values.set(t.value, new Set());
    values.get(t.value).add(zone);
  });
  const out = [];
  byParam.forEach((values, param) => {
    if (values.size < 2) return;
    const views = [...values].map(([value, zones]) => ({ value, zones,
      charts: items.filter(it => it.type === "worksheet" && zones.has(zoneOf(it)) && switchFamily(it)) }));
    if (views.some(v => v.charts.length !== 1) || new Set(views.map(v => switchFamily(v.charts[0]))).size !== 1) return;
    const shown = views.filter(v => !alternates.has(zoneOf(v.charts[0])));
    if (shown.length !== 1) return;
    const buttons = ((model && model.paramActions) || []).filter(a => a.param === param && (!a.dashboard || a.dashboard === dashName))
      .map(a => items.find(it => it.name === a.worksheet)).filter(Boolean);
    if (!buttons.length) return;
    // text boxes the buttons switch too (a view's title): the shown one, and each value's at the same place across
    const zoneAt = it => (dashM.zones || []).find(z => String(z.id) === zoneOf(it));
    const textOf = it => it.textCard.tiles[0].lines.map(l => l.segments.map(sg => sg.text).join(""));
    const texts = items.filter(it => it.textCard && shown[0].zones.has(zoneOf(it))).map(it => {
      const z = zoneAt(it);
      const byValue = views.map(v => {
        if (v === shown[0]) return textOf(it);
        const twins = items.filter(o => o.textCard && v.zones.has(zoneOf(o)) && zoneAt(o) && z && Math.abs(zoneAt(o).x - z.x) < 500);
        twins.sort((a, b) => Math.abs(zoneAt(a).y - z.y) - Math.abs(zoneAt(b).y - z.y));
        return twins.length ? textOf(twins[0]) : null;
      });
      return byValue.every(Boolean) ? { item: it, lines: byValue } : null;
    }).filter(Boolean);
    out.push({ texts,
      param, value: shown[0].value, host: shown[0].charts[0], buttons,
      views: views.map(v => ({ value: v.value, item: v.charts[0] })),
      // the other values' zones (their charts, legends …) drawn below the dashboard until now: the switch shows them
      dropped: items.filter(it => alternates.has(zoneOf(it)) && views.some(v => v.zones.has(zoneOf(it)))),
      /** @type {any} */ sw: null
    });
  });
  return out;
}

/** a string as an Excel formula literal */
const xlText = s => `"${String(s).replace(/"/g, '""')}"`;

/**
 * The switch's cell on the chart data sheet (its value: the view shown; once the dropdown is written, a link to it).
 * pick(parts) → a formula taking the part of the view picked.
 * @param {import("exceljs").Worksheet} ds @param {number} row 0-based @param {any} group findViewSwitches
 */
function makeSwitch(ds, row, group) {
  const values = group.views.map(v => v.value);
  ds.getCell(row + 1, 1).value = `Switch: ${group.param.replace(/_\d{6,}$/, "")}`;
  ds.getCell(row + 1, 1).font = { bold: true };
  const cell = ds.getCell(row + 1, 2);
  cell.value = group.value;
  const ref = sheetRef(ds.name, `$B$${row + 1}`);
  const idx = `MATCH(${ref},{${values.map(xlText).join(",")}},0)`;
  return { cell, ref, values, pick: parts => `CHOOSE(${idx},${parts.join(",")})`,
           is: value => `${ref}=${xlText(value)}` };
}

/**
 * A pie whose views switch: each view's slices written as usual (writePieData), then one list of every view's slices
 * – its value and label cells #N/A / empty unless its view is picked – that the chart draws, its labels from cells.
 * Returns the panes (writePieData's) and the pie to draw (the shown view's, with every view's slices).
 * @param {import("exceljs").Worksheet} ds @param {number} startRow 1-based @param {any} group @param {string} name
 */
function writeSwitchedPie(ds, startRow, group, name) {
  const sw = group.sw;
  const host = group.host.visualModel.pie;
  let row = startRow;
  const parts = group.views.map(v => {
    const pie = v.item.visualModel.pie;
    const data = writePieData(ds, row, pie, v.item.name);
    row = data.nextRow;
    return { value: v.value, pie, pane: [...data.panes.values()][0] };
  });
  ds.getCell(row, 1).value = `${name} (switched)`;
  ds.getCell(row, 1).font = { bold: true };
  ["Slice", "Value", "Label"].forEach((t, k) => { ds.getCell(row + 1, k + 1).value = t; });
  const first = row + 2;
  let r = first;
  const cellOf = (range, i) => { const m = String(range).match(/\$([A-Z]+)\$(\d+)/); return `${m[1]}${+m[2] + i}`; };
  const points = [];
  parts.forEach(p => {
    const shown = p.pie === host;
    p.pie.points.forEach((pt, i) => {
      ds.getCell(r, 1).value = pt.category;
      ds.getCell(r, 2).value = { formula: `IF(${sw.is(p.value)},${cellOf(p.pane.ref.val, i)},NA())`, result: shown ? pt.value : { error: "#N/A" } };
      ds.getCell(r, 3).value = { formula: `IF(${sw.is(p.value)},${cellOf(p.pane.ref.lbl, i)},"")`, result: shown && pt.label ? plainText(pt.label) : "" };
      points.push({ ...pt, rowKey: host.rowKeys[0], colKey: host.colKeys[0], value: shown ? pt.value : null });
      r++;
    });
  });
  const last = r - 1;
  const ref = { sheet: ds.name, cat: `$A$${first}:$A$${last}`, val: `$B$${first}:$B$${last}`, lbl: `$C$${first}:$C$${last}`,
                total: `SUMIF($B$${first}:$B$${last},">0")` };
  // the hole's text: each line the picked view's (its cell when it follows the slicers, else its text)
  let centre = null;
  const hostLines = host.hole && !host.centerText ? host.hole.labelFor(host.points) : null;
  if (hostLines && hostLines.length) {
    ds.getCell(r, 1).value = "Centre text";
    centre = hostLines.map((line, li) => {
      const texts = parts.map(p => {
        const own = p.pie.hole ? p.pie.hole.labelFor(p.pie.points) : [];
        const linked = p.pane.centre && p.pane.centre[li];
        return linked || xlText(own[li] ? plainText([own[li]]) : "");
      });
      const cell = ds.getCell(r, li + 2);
      cell.value = { formula: sw.pick(texts), result: plainText([line]) };
      return sheetRef(ds.name, `$${getExcelColumnName(li + 1)}$${r}`);
    });
    r++;
  }
  const pie = { ...host, points, switched: true,
                hole: host.hole ? { ...host.hole, labelFor: () => hostLines || [] } : null };
  const innerRef = host.inner ? parts.find(p => p.pie === host).pane.innerRef : null;
  return { panes: new Map([[host.rowKeys[0] + "\u0001" + host.colKeys[0], { ref, innerRef, centre }]]), pie, nextRow: r + 1 };
}

/**
 * A chart whose views switch: each view's data written as usual (writeChartData, after prep: slicer formulas …) with
 * a column of #N/A beside it, then the chart's spec – the shown view's, with every view's series – and its refs:
 * defined names that pick the picked view's categories and each series' values (#N/A for the other views).
 * @param {import("exceljs").Worksheet} ds @param {{ row: number, names?: number }} chartData @param {any} group
 * @param {(item: any, spec: any) => void} prep
 * @returns {{ spec: any, refs: any, names: { name: string, formula: string }[], parts: any[] }}
 */
function writeSwitchedChart(ds, chartData, group, prep) {
  const sw = group.sw;
  const id = chartData.names = (chartData.names || 0) + 1;
  // views of other names (Department / Job Role): one axis shows each, so Excel fits them – the names kept whole
  const levelsOf = v => JSON.stringify(/** @type {any} */ (v.item.visualModel.chartSpecs[0]).categories.levels);
  const otherNames = new Set(group.views.map(levelsOf)).size > 1;
  const parts = group.views.map(v => {
    const spec = /** @type {any} */ (v.item.visualModel.chartSpecs[0]);
    prep(v.item, spec);
    if (otherNames) spec.keepNames = true;
    const start = chartData.row;
    const refs = writeChartData(ds, spec, start);
    chartData.row = refs.nextRow;
    const L = spec.categories.levels.length, N = L ? spec.categories.levels[0].length : 0;
    const free = L + spec.series.length + spec.series.filter(s => s.labelTexts).length + (spec.helpers || []).length;
    ds.getCell(start + 1, free + 1).value = "(not shown)";
    for (let i = 0; i < N; i++) ds.getCell(start + 2 + i, free + 1).value = { formula: "NA()", result: { error: "#N/A" } };
    return { value: v.value, item: v.item, spec, refs,
             na: cellRef(ds.name, free, start + 1, free, start + Math.max(1, N)),
             blank: cellRef(ds.name, free + 1, start + 1, free + 1, start + Math.max(1, N)) };
  });
  const host = parts.find(p => p.item === group.host);
  const n = host.spec.categories.levels.length ? host.spec.categories.levels[0].length : 0;
  const names = [{ name: `_Switch${id}_Categories`, formula: sw.pick(parts.map(p => p.refs.cat)) }];
  const series = [], seriesRefs = [];
  parts.forEach((p, k) => p.spec.series.forEach((s, j) => {
    const r = p.refs.series[j];
    const pick = (own, other) => sw.pick(parts.map((q, m) => (m === k ? own : other(q))));
    const val = `_Switch${id}_S${series.length + 1}`;
    names.push({ name: val, formula: pick(r.val, q => q.na) });
    let lbl = null;
    if (r.lbl) { lbl = `${val}_Labels`; names.push({ name: lbl, formula: pick(r.lbl, q => q.blank) }); }
    const shown = p === host;
    series.push({ ...s, values: shown ? s.values : new Array(n).fill(null),
                  ...(s.labelTexts ? { labelTexts: shown ? s.labelTexts : new Array(n).fill("") } : {}) });
    seriesRefs.push({ tx: r.tx, val: `[0]!${val}`, ...(lbl ? { lbl: `[0]!${lbl}` } : {}) });
  }));
  const spec = { ...host.spec, series };
  // other names in another view (Department / Job Role): no rotation of the shown one's – Excel turns crowded names
  if (otherNames && !host.spec.categoryRotation) spec.categoryAuto = true;
  // an axis range of one view would clip another's values: kept only when every view has it
  ["valueMin", "valueMax"].forEach(k => { if (parts.some(p => p.spec[k] !== host.spec[k])) delete spec[k]; });
  return { spec, refs: { cat: `[0]!_Switch${id}_Categories`, series: seriesRefs }, names, parts };
}

/**
 * The legend line above a switched chart: one cell per entry ("■ name"), each a formula of the switch, coloured by
 * the picked view (conditional formatting) – entries placed across the chart's columns at the widest view's widths.
 * false when the row has no room (something else in its cells): the chart keeps its own legend.
 * @param {import("exceljs").Worksheet} ws @param {number} row 1-based @param {number} col 0-based first column
 * @param {number[]} colPx the grid's column widths @param {number} widthPx @param {any[]} parts writeSwitchedChart's
 * @param {any} sw @param {any} font
 */
function writeSwitchLegend(ws, row, col, colPx, widthPx, parts, sw, font) {
  const entries = parts.map(p => {
    const list = p.spec.series.filter(s => s.color && !s.refLine)
      .map(s => ({ name: String(s.name), color: String(s.color).replace(/^#/, "").slice(-6).toUpperCase() }));
    return p.spec.stacked && p.spec.barDir !== "bar" ? list.reverse() : list;       // stacked columns: the legend's order
  });
  const charPx = (font.size || 9) * 4 / 3 * 0.5;                 // Arial's average character
  const slots = Math.max(...entries.map(e => e.length));
  // too long for the chart's width: each entry an equal share, its name cut at the end ("South East As…", as Tableau)
  const fullPx = n => (n.length + 3) * charPx + 10;
  const total = Array.from({ length: slots }, (_, i) => Math.max(...entries.map(e => (e[i] ? fullPx(e[i].name) : 0)))).reduce((a, b) => a + b, 0);
  const room = total > widthPx ? Math.max(4, Math.floor((widthPx / slots - 10) / charPx) - 3) : 0;
  entries.forEach(e => e.forEach(x => { if (room && x.name.length > room) x.name = x.name.slice(0, room - 1).trimEnd() + "…"; }));
  /** each column's left edge, px from the chart's */
  const edges = [];
  for (let c = col, x = 0; c < col + 400 && x <= widthPx; c++) { edges.push({ c, x }); x += colPx[c] || 0; }
  // each entry in the column where its room begins (a cell's text runs on over the empty cells after it)
  const cols = [];
  for (let i = 0, x = 0; i < slots; i++) {
    const at = [...edges].reverse().find(e => e.x <= x + 2 && (!cols.length || e.c > cols[cols.length - 1])) ||
               edges.find(e => !cols.length || e.c > cols[cols.length - 1]);
    const need = Math.max(...entries.map(e => (e[i] ? (e[i].name.length + 3) * charPx : 0))) + 10;
    if (!at || (at.x + need > widthPx + 8 && cols.length)) break;
    cols.push(at.c);
    x = Math.max(x, at.x) + need;
  }
  if (!cols.length) return false;
  // each name cut to the room its cell really has, up to the next entry's
  const xOf = c => (edges.find(e => e.c === c) || { x: 0 }).x;
  cols.forEach((c, i) => {
    const roomPx = (i + 1 < cols.length ? xOf(cols[i + 1]) : widthPx) - xOf(c) - 6;
    const fit = Math.max(3, Math.floor(roomPx / charPx) - 2);
    entries.forEach(e => { if (e[i] && e[i].name.length > fit) e[i].name = e[i].name.replace(/…$/, "").slice(0, fit - 1).trimEnd() + "…"; });
  });
  for (let cc = col; cc <= Math.max(...cols) + 1; cc++) {
    const v = ws.getCell(row, cc + 1).value;
    if (v !== null && v !== undefined && v !== "") return false;
  }
  const shown = Math.max(0, parts.findIndex(p => p.value === sw.cell.value));
  cols.forEach((cc, i) => {
    const cell = ws.getCell(row, cc + 1);
    cell.value = { formula: sw.pick(entries.map(e => xlText(e[i] ? `■ ${e[i].name}` : ""))),
                   result: entries[shown][i] ? `■ ${entries[shown][i].name}` : "" };
    cell.font = font;
    cell.alignment = { vertical: "bottom", wrapText: false };
    ws.addConditionalFormatting({ ref: cell.address, rules: parts.map((p, k) => entries[k][i] ? {
      type: "expression", priority: k + 1, formulae: [sw.is(p.value)],
      style: { font: { color: { argb: "FF" + entries[k][i].color } } } } : null).filter(Boolean) });
  });
  return true;
}

/* px between the dashboard and the views its buttons switch to, drawn below it (room for their header) */
const ALTERNATE_GAP_PX = 60;

/**
 * The views a dashboard's buttons switch between (dynamic zone visibility on a parameter: Gender / Race, Department /
 * Job Role): Tableau shows one at a time, Excel cannot switch – so the hidden ones are drawn too, moved below the
 * dashboard at their own place across, as the dashboard would look with the other choice. Returns the dashboard with
 * those objects shown and moved (its methods kept), the moved zone ids, and the px where they begin.
 * @param {any} dashboard @param {FormatModel | null} model
 */
function withAlternateViews(dashboard, model) {
  const dash = model && model.dashboards && model.dashboards[dashboard.name];
  const toggles = (dash && dash.toggles) || {};
  const objects = dashboard.objects || [];
  const moved = objects.filter(o => o && o.position && o.size && o.isVisible === false && toggles[String(o.id)]);
  if (!moved.length) return { dashboard, alternates: new Set(), top: 0 };
  const shown = objects.filter(o => o && o.position && o.size && o.isVisible !== false);
  const top = Math.max(0, ...shown.map(o => o.position.y + o.size.height)) + ALTERNATE_GAP_PX;
  const dy = top - Math.min(...moved.map(o => o.position.y));
  const out = objects.map(o => (moved.includes(o) ? { ...plainObject(o), isVisible: true, position: { x: o.position.x, y: o.position.y + dy } } : o));
  console.log(`[Views] ${dashboard.name}: ${moved.length} object(s) a button shows instead – drawn below the dashboard`);
  return { dashboard: Object.create(dashboard, { objects: { value: out, enumerable: true } }), alternates: new Set(moved.map(o => String(o.id))), top };
}

/** Worksheets hidden on a dashboard by dynamic zone visibility (isVisible === false) – except the views a parameter's
 * buttons switch to (the workbook's toggles): the export draws those too (withAlternateViews). */
function hiddenWorksheetNames(objects, dashName) {
  const model = getFormatModel();
  const dash = model && model.dashboards && dashName ? model.dashboards[dashName] : null;
  const toggled = dash && dash.toggles ? dash.toggles : {};
  return new Set((objects || []).filter(o => o && o.type === "worksheet" && o.isVisible === false && !toggled[String(o.id)]).map(o => o.name));
}

/**
 * readSheet for every worksheet; results in the worksheets' order. A hidden (embedded) viz: one worksheet at a
 * time (see readSheet). The live dashboard: SHEET_READ_CONCURRENCY at a time, like fetchAllSheetsData.
 */
async function readSheets(worksheets, wrap, plain, objects, dashName) {
  const hidden = hiddenWorksheetNames(objects, dashName);
  /** @type {Map<string, Promise<number>>} each filtered field's value count, asked once for the dashboard */
  const domains = new Map();
  /** @type {Map<string, Promise<any>>} the data sources copied for slicers, each read once for the dashboard */
  const sources = new Map();
  const list = [...(worksheets || [])];
  const out = new Array(list.length);
  let next = 0;
  const worker = async () => {
    while (next < list.length) {
      const i = next++;
      out[i] = await readSheet(list[i], wrap, plain, hidden.has(list[i].name), domains, sources);
    }
  };
  const lanes = plain ? 1 : SHEET_READ_CONCURRENCY;
  await Promise.all(Array.from({ length: Math.min(lanes, list.length) }, worker));
  return out;
}

/** Read every worksheet of the dashboard shown in `viz` into the export's input shape. */
async function readEmbeddedDashboard(viz) {
  const sheet = viz.workbook.activeSheet;
  if (!sheet || sheet.sheetType !== "dashboard") throw new Error("opened view is not a dashboard");
  const pass = p => p;
  const sheets = await readSheets(sheet.worksheets, pass, true, sheet.objects, sheet.name);
  let params = [];
  try {
    const wb = viz.workbook;
    params = typeof wb.getParametersAsync === "function" ? ((await wb.getParametersAsync()) || []).map(plainParameter) : [];
  } catch (e) { /* no parameter controls exported */ }
  const objects = (sheet.objects || []).map(plainObject);
  return { dashboard: { name: sheet.name, objects, worksheets: sheets, getParametersAsync: async () => params }, sheets };
}

/**
 * Loading screen over the extension while other dashboards are read. The hidden vizzes sit
 * underneath it (they must stay on screen for Chrome to render them), so the user sees progress
 * instead of dashboards flashing by. The debug view and the sign-in panel sit above it.
 */
function showExportOverlay(title, detail) {
  if (!document.getElementById("tfx_loading_css")) {
    const css = document.createElement("style");
    css.id = "tfx_loading_css";
    css.textContent =
      "#tfx_loading{position:fixed;inset:0;z-index:9000;background:#fff;display:flex;flex-direction:column;" +
      "align-items:center;justify-content:center;gap:10px;padding:16px;text-align:center;" +
      "font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;color:#1c2b26;}" +
      "#tfx_loading .tfx-spin{width:28px;height:28px;border:3px solid #dfe6e2;border-top-color:#1c7a4d;" +
      "border-radius:50%;animation:tfxspin .8s linear infinite;}" +
      "#tfx_loading .tfx-title{font-size:13px;font-weight:600;}" +
      "#tfx_loading .tfx-detail{font-size:11px;color:#5b6b65;max-width:260px;line-height:1.4;}" +
      "@keyframes tfxspin{to{transform:rotate(360deg)}}";
    document.head.appendChild(css);
  }
  let overlay = document.getElementById("tfx_loading");
  if (!overlay) {
    overlay = document.createElement("div");
    overlay.id = "tfx_loading";
    overlay.setAttribute("role", "status");
    overlay.setAttribute("aria-live", "polite");
    overlay.innerHTML = '<div class="tfx-spin"></div><div class="tfx-title"></div><div class="tfx-detail"></div>';
    document.body.appendChild(overlay);
  }
  overlay.querySelector(".tfx-title").textContent = title;
  overlay.querySelector(".tfx-detail").textContent = detail || "";
}

function hideExportOverlay() {
  const overlay = document.getElementById("tfx_loading");
  if (overlay) overlay.remove();
}

/** Rejects with a readable error if `promise` doesn't settle in time (nothing may hang the export). */
function withTimeout(promise, ms, what) {
  let t;
  return Promise.race([
    promise,
    new Promise((_, reject) => { t = setTimeout(() => reject(new Error(`${what} took longer than ${ms / 1000}s`)), ms); }),
  ]).finally(() => clearTimeout(t));
}

/**
 * Read the CURRENT dashboard's data, visual specifications, filters, selected marks and parameters up front,
 * through the Extensions API, BEFORE the Embedding API is loaded. Once embedded vizzes are connected inside
 * this frame, Extensions API calls can stop getting answers (both libraries talk to Tableau through
 * postMessage), which left the export stuck on "Building Excel…". Returns the same { dashboard, sheets }
 * shape, with each sheet answering from the snapshot.
 */
async function snapshotLiveDashboard(current) {
  const bounded = (p, what) => withTimeout(p, LIVE_READ_TIMEOUT_MS, what);
  const sheets = await readSheets(current.sheets, bounded, false, current.dashboard && current.dashboard.objects,
                                 current.dashboard && current.dashboard.name);
  let params = [];
  try {
    const d = current.dashboard;
    params = typeof d.getParametersAsync === "function" ? (await bounded(d.getParametersAsync(), "Parameters")) || [] : [];
  } catch (e) { /* no parameter controls exported */ }
  const d = current.dashboard;
  return { dashboard: { name: d.name, objects: d.objects, worksheets: sheets, getParametersAsync: async () => params }, sheets };
}

/**
 * One-time Tableau sign-in for embedded dashboards.
 * The extension's frame can't reuse the user's normal Tableau login, so Tableau shows its own
 * "Sign in to Tableau Cloud" button inside the embedded view. We show that view (with a short
 * explanation) and wait until it becomes interactive = signed in. Resolves true / false (skipped).
 */
function promptTableauSignIn(embedUrl) {
  return new Promise(async resolve => {
    const overlay = document.createElement("div");
    overlay.id = "tfx_signin";
    overlay.style.cssText = "position:fixed;inset:0;z-index:10000;background:#fff;display:flex;flex-direction:column;";
    const bar = document.createElement("div");
    bar.style.cssText = "display:flex;gap:8px;align-items:center;padding:8px 10px;border-bottom:1px solid #dfe6e2;font-size:12px;line-height:1.4;";
    const msg = document.createElement("div");
    msg.style.flex = "1";
    msg.textContent = "To export the other dashboards, Tableau needs you to sign in once. Click \u201cSign in to Tableau Cloud\u201d below.";
    const skip = document.createElement("button");
    skip.className = "btn-load";
    skip.style.width = "auto";
    skip.textContent = "Skip";
    bar.appendChild(msg);
    bar.appendChild(skip);
    overlay.appendChild(bar);
    document.body.appendChild(overlay);

    // the viz lives in the embed room (with the Embedding API): the room is shown under the bar meanwhile
    const win = await embedRoom();
    const room = /** @type {any} */ (EMBED_ROOM).frame;
    const top = bar.offsetHeight || 40;
    room.style.cssText = `position:fixed;left:0;top:${top}px;width:100vw;height:calc(100vh - ${top}px);border:0;z-index:10001;background:#fff;`;
    const viz = win.document.createElement("tableau-viz");
    viz.setAttribute("src", embedUrl);
    viz.setAttribute("toolbar", "hidden");
    viz.setAttribute("hide-tabs", "");
    viz.setAttribute("width", String(Math.max(300, room.clientWidth || 0)));
    viz.setAttribute("height", String(Math.max(300, room.clientHeight || 0)));

    let finished = false;
    const finish = ok => {
      if (finished) return;
      finished = true;
      clearTimeout(timer);
      viz.remove();
      room.style.cssText = debugVizOn() ? ROOM_DEBUG_CSS : ROOM_HIDDEN_CSS;
      overlay.remove();
      resolve(ok);
    };
    const timer = setTimeout(() => finish(false), SIGNIN_WAIT_MS);
    viz.addEventListener("firstinteractive", () => finish(true));
    skip.addEventListener("click", () => finish(false));
    win.document.body.appendChild(viz);
  });
}

/**
 * All visible dashboards in tab order: the current one from the Extensions API (live
 * filters), the others via hidden vizzes. Never throws for a single dashboard — failures
 * become `notes`, and the export continues with whatever could be read.
 */
async function collectAllDashboards(current, onProgress) {
  const notes = [];
  const wb = loadedWorkbook();
  if (!wb || !wb.xml) return { targets: [current], notes: ["Other dashboards: load the workbook first."] };
  const order = dashboardTabOrder(wb.xml);
  if (!order.includes(current.dashboard.name)) order.unshift(current.dashboard.name);
  if (order.length === 1) return { targets: [current], notes };
  if (!wb.id) {
    return { targets: [current], notes: ["Other dashboards are only available when the workbook is auto-loaded from Tableau Cloud (not from a manual file)."] };
  }

  // Current dashboard first — before any embedded viz exists in this frame.
  const tLive = performance.now();
  current = await snapshotLiveDashboard(current);
  console.log(`[Timing] "${current.dashboard.name}" (this dashboard): read in ${secs(performance.now() - tLive)}`);

  let views, libUrl;
  try {
    const r = await (await backendFetch(`/workbooks/${encodeURIComponent(wb.id)}/views`)).json();
    views = r.views || [];
    libUrl = r.embedLibUrl;
    if (!libUrl || !views.every(v => v.embedUrl)) throw new Error("backend is outdated — update worker.js");
    await loadEmbeddingApi(libUrl);
  } catch (e) {
    return { targets: [current], notes: [`Other dashboards skipped: ${e.message}`] };
  }

  const others = order.filter(n => n !== current.dashboard.name);
  /** @type {Map<string, { target?: any, note?: string }>} name → what reading it gave */
  const results = new Map();
  let finished = 0, abort = null, signInPromise = null;
  // count the current dashboard too (it is already read), so 2 dashboards show "1 of 2", then "2 of 2"
  const report = (name, stage) => { if (onProgress) onProgress(name, finished + 1, others.length + 1, stage); };
  /** One-time sign-in shared by the parallel readers: only one panel, everyone waits for it. */
  const ensureSignedIn = (embedUrl, name) => {
    if (!signInPromise) {
      report(name, "signin");
      signInPromise = promptTableauSignIn(embedUrl);
    }
    return signInPromise;
  };

  /** Read one dashboard through a hidden viz; never throws – a failure becomes its note. */
  const readOne = async name => {
    if (abort) { results.set(name, { note: `"${name}": skipped (${abort})` }); return; }
    const tStart = performance.now();
    const view = views.find(v => v.name === name);
    if (!view) { results.set(name, { note: `"${name}": not published as a tab — skipped` }); return; }
    let viz = null;
    try {
      const size = dashboardPixelSize(wb.xml, name);
      try {
        viz = await openHiddenViz(view.embedUrl, size);
      } catch (e) {
        if (e.code !== "SIGN_IN_NEEDED") throw e;
        if (!(await ensureSignedIn(view.embedUrl, name))) throw new Error("Tableau sign-in was skipped");
        viz = await openHiddenViz(view.embedUrl, size);       // signed in now → open hidden again
      }
      const tOpened = performance.now();
      results.set(name, { target: await readEmbeddedDashboard(viz) });
      console.log(`[Timing] "${name}": opened in ${secs(tOpened - tStart)}, read in ${secs(performance.now() - tOpened)}`);
    } catch (e) {
      results.set(name, { note: `"${name}": ${e.message}` });
      if (/sign-in|refused|Embedding API|never finished drawing/i.test(e.message)) abort = "same problem as the previous dashboard";
    } finally {
      if (viz) viz.remove();
      finished++;
      report(name);
    }
  };

  // All of them start together, PARALLEL_DASHBOARDS at a time. A needed sign-in is still asked only once
  // (ensureSignedIn shares one panel and everyone waits for it), and if embedding is broken they fail
  // together – no slower than finding out on one dashboard first.
  report(others[0]);
  const queue = others.slice();
  const pool = Array.from({ length: Math.min(PARALLEL_DASHBOARDS, queue.length) }, async () => {
    while (queue.length) await readOne(queue.shift());
  });
  await Promise.all(pool);

  // Tableau tab order, whatever order they finished in (the notes too).
  const targets = [];
  for (const name of order) {
    if (name === current.dashboard.name) { targets.push(current); continue; }
    const r = results.get(name);
    if (r && r.target) targets.push(r.target);
    else if (r && r.note) notes.push(r.note);
  }
  removeHiddenVizHost();
  clearDebugVizSwitch();
  return { targets, notes };
}

/** Remove the hidden-viz host once reading is done, so nothing (debug view included) covers the extension. */
function removeHiddenVizHost() {
  if (EMBED_ROOM) { EMBED_ROOM.frame.remove(); EMBED_ROOM = null; }
}

/* =============================================================================
 * uniqueExcelSheetName() — a valid, unique Excel sheet name for a dashboard.
 * Excel rules: max 31 chars; none of \ / ? * [ ] : ; not empty; must not start or
 * end with an apostrophe; "History" is reserved; names are unique ignoring case.
 * usedLower: Set of lower-cased names already taken (updated here).
 * ============================================================================= */
function uniqueExcelSheetName(dashboardName, usedLower) {
  let base = String(dashboardName || "Dashboard Export")
    .replace(/[\\\/\*\?\[\]:]/g, "")
    .replace(/^'+|'+$/g, "")
    .slice(0, 31)
    .replace(/^'+|'+$/g, "");
  if (!base.trim()) base = "Dashboard Export";
  if (base.toLowerCase() === "history") base = "History (1)";
  let name = base, n = 2;
  while (usedLower.has(name.toLowerCase())) {
    const suffix = ` (${n++})`;
    name = base.slice(0, 31 - suffix.length).replace(/[\s']+$/, "") + suffix;
  }
  usedLower.add(name.toLowerCase());
  return name;
}


/* ═══ export/slicers.js ═══════════════════════════════════════════════════════════════════════════ */
/* Tableau filter cards → Excel slicers that cross-filter the dashboard. The rows of the filtered field's data source go
 * to one hidden table ("Slicer Data"); a slicer on one of its columns hides the rows of the values not picked, and every
 * chart value, table cell and KPI value that follows the slicers is a formula over the table's visible rows
 * (SUBTOTAL(103) = 1) – one click filters every chart and table of that data source, in every dashboard of the file.
 *   Each visual gets a key column: the mark (row of its summary data) each data-source row belongs to, found by the
 * visual's dimensions (Region, MONTH(Order Date) …) under the sheet's other filters. A measure becomes SUMIFS,
 * AVERAGEIFS, MINIFS, MAXIFS, COUNTIFS or a distinct count over the rows of its key, or a calculation of those
 * (SUM([Profit]) / SUM([Sales])). Before a visual is bound, those formulas are worked out on the rows with Tableau's
 * current selection and have to give Tableau's own numbers: a visual that does not (a calculated dimension, an LOD
 * expression, a table calculation, a filter that cannot be told) keeps its fixed values, and the report says why. */

const SLICER_DATA_SHEET = "Slicer Data";
const SLICER_LISTS_SHEET = "Slicer Lists";                       // the dropdown cells' lists of values
const SLICER_DATES_COL = "ZZ";                                    // …and, far right there, the dates the date dropdowns make
const SLICER_BOUND_COL = "ZY";                                    // …and the Min / Max of the ranges of an aggregate
const SLICER_TABLE = "SlicerData";
const SLICER_ALL = "(All)";

/**
 * A calculated filter that keeps the dates between two date parameters – Tableau's usual "Start Date / End Date"
 * pattern: [Order Date] >= [Parameters].[Start] AND [Order Date] <= [Parameters].[End] (also inside IF … THEN TRUE
 * [ELSE FALSE] END – without ELSE the rows outside are Null, which the filter's True leaves out as well). Returns the
 * date field, the parameters and whether each bound is strict; null for anything else.
 * @param {string} formula
 */
function slicerParamRange(formula) {
  let f = String(formula || "").replace(/\/\/[^\n]*/g, "").replace(/\s+/g, " ").trim();
  const wrapped = f.match(/^IF (.+?) THEN TRUE(?: ELSE FALSE)? END$/i);
  if (wrapped) f = wrapped[1];
  const cmp = "\\(?\\s*\\[([^\\]]+)\\]\\s*(>=|>|<=|<)\\s*\\[Parameters\\]\\.\\[([^\\]]+)\\]\\s*\\)?";
  const m = f.match(new RegExp(`^\\(?\\s*${cmp}\\s+AND\\s+${cmp}\\s*\\)?$`, "i"));
  if (!m) return null;
  const a = { date: m[1], op: m[2], param: m[3] }, b = { date: m[4], op: m[5], param: m[6] };
  if (a.date.toLowerCase() !== b.date.toLowerCase()) return null;
  const lo = [a, b].find(x => x.op[0] === ">"), hi = [a, b].find(x => x.op[0] === "<");
  if (!lo || !hi) return null;
  return { date: a.date, lo: lo.param, hi: hi.param, loStrict: lo.op === ">", hiStrict: hi.op === "<" };
}
/**
 * The days a relative date filter keeps at export ("Last 3 months", "Year to date" …): Tableau counts from the
 * anchor (today unless the filter names another day), whole periods – "Last 3 months" ends with this month. Returns
 * { lo, hi, text } (UTC dates), or null for a period or range it does not know.
 * @param {any} f a Tableau RelativeDateFilter
 */
function slicerRelativeRange(f) {
  const unit = String(f.periodType || "").toLowerCase().replace(/s$/, "");
  const r = String(f.rangeType || "").toLowerCase().replace(/[^a-z]/g, ""), n = Math.max(1, Math.round(Number(f.rangeN) || 1));
  const raw = f.anchorDate && typeof f.anchorDate === "object" && !(f.anchorDate instanceof Date) ? f.anchorDate.value : f.anchorDate;
  const a0 = raw ? new Date(raw) : new Date();
  if (isNaN(+a0)) return null;
  const y = a0.getFullYear(), m = a0.getMonth(), d = a0.getDate();
  const day = 86400000, anchor = Date.UTC(y, m, d);
  /** the first day of the period k periods from the anchor's */
  const start = k => {
    switch (unit) {
      case "year": return Date.UTC(y + k, 0, 1);
      case "quarter": return Date.UTC(y, m - (m % 3) + 3 * k, 1);
      case "month": return Date.UTC(y, m + k, 1);
      case "week": return anchor - new Date(anchor).getUTCDay() * day + 7 * k * day;
      case "day": return anchor + k * day;
      default: return NaN;
    }
  };
  const end = k => start(k + 1) - day;
  const span = { current: [0, 0], last: [-1, -1], next: [1, 1], lastn: [-(n - 1), 0], nextn: [0, n - 1] }[r];
  const [lo, hi] = r === "todate" ? [start(0), anchor] : span ? [start(span[0]), end(span[1])] : [NaN, NaN];
  if (isNaN(lo) || isNaN(hi)) return null;
  const text = filterCardText({ ...f, filterType: "relative-date" });
  return { lo: new Date(lo), hi: new Date(hi), text: text || "relative dates" };
}

/* a slicer as Excel draws it (default style): its caption, its buttons, a little padding; px */
const SLICER_HEADER_PX = 34, SLICER_ITEM_PX = 26, SLICER_PAD_PX = 10, SLICER_ITEM_MIN_W_PX = 90, SLICER_MAX_ITEM_ROWS = 8;
/** TWB derivations of a date a slicer key can be built from (YEAR, QUARTER, MONTH, DAY, and their truncated forms) */
const SLICER_DATE_PARTS = new Set(["yr", "qr", "mn", "dy", "tyr", "tqr", "tmn", "tdy"]);
/** Tableau aggregations → the slicer formulas' aggregations */
const SLICER_AGGS = { sum: "SUM", avg: "AVG", min: "MIN", max: "MAX", cnt: "COUNT", ctd: "COUNTD" };

/** a date value as { y, m, d } – Tableau's dates arrive at midnight, sometimes in local time: the nearest day
 * @param {DataValue} dv */
function slicerDate(dv) {
  const s = tvExcelSerial(dv);
  if (s === null) return null;
  const t = new Date(Math.round(s - 25569) * 86400000);
  return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate() };
}

/** @param {{ y: number, m: number, d: number }} d @param {string} deriv */
function slicerDatePart(d, deriv) {
  const two = n => String(n).padStart(2, "0");
  switch (deriv) {
    case "yr": case "tyr": return String(d.y);
    case "qr": return String(Math.ceil(d.m / 3));
    case "tqr": return `${d.y}-Q${Math.ceil(d.m / 3)}`;
    case "mn": return String(d.m);
    case "tmn": return `${d.y}-${two(d.m)}`;
    case "dy": return String(d.d);
    default: return `${d.y}-${two(d.m)}-${two(d.d)}`;
  }
}

/**
 * A data-source value as the key text of a field derived as `deriv` (YEAR → "2023", MONTH → "7", a truncated month →
 * "2023-07", a date → "2023-07-14", anything else its value); null when it cannot be told.
 * @param {DataValue} dv @param {string} deriv @param {boolean} isDate the column holds dates
 */
function slicerRowText(dv, deriv, isDate) {
  if (tfIsNull(dv)) return "\u0000";
  if (isDate || SLICER_DATE_PARTS.has(deriv)) { const d = slicerDate(dv); return d ? slicerDatePart(d, deriv) : null; }
  return String(dv.nativeValue !== undefined ? dv.nativeValue : dv.value);
}

/** The same key text of a mark's or a filter's value: summary data has YEAR as 2023, MONTH as 7 or "July", QUARTER as
 * "Q3", a truncated month as its first day. @param {DataValue} dv @param {string} deriv @param {boolean} isDate */
function slicerMarkText(dv, deriv, isDate) {
  if (tfIsNull(dv)) return "\u0000";
  const v = dv.nativeValue !== undefined ? dv.nativeValue : dv.value;
  if (/^(yr|qr|mn|dy)$/.test(deriv) && !(v instanceof Date)) {
    const text = tfDvText(dv).trim();
    if (deriv === "mn") { const r = tfCalendarRank(text); if (r && r.kind === "month") return String(r.rank + 1); }
    if (deriv === "qr") { const q = text.match(/^Q?\s*([1-4])$/i); if (q) return q[1]; }
    const n = typeof v === "number" ? v : Number(String(v).trim());
    return isFinite(n) ? String(Math.round(n)) : null;
  }
  return slicerRowText(dv, deriv, isDate);
}

/** the worksheet's data source that has the field (a blend lists several), else its first; null when Tableau gives none
 * @param {any} ws @param {string} fieldName */
async function slicerDataSourceOf(ws, fieldName) {
  if (typeof ws.getDataSourcesAsync !== "function") return null;
  const sources = (await ws.getDataSourcesAsync()) || [];
  const has = ds => (ds.fields || []).some(f => tfNorm(f.name) === tfNorm(fieldName));
  return sources.find(has) || sources[0] || null;
}

/** a logical table's column names (one row read) – which of a data source's related tables holds a field
 * @param {any} ds @param {string} id */
async function slicerTableColumns(ds, id) {
  let t;
  if (typeof ds.getLogicalTableDataReaderAsync === "function") {
    const reader = await ds.getLogicalTableDataReaderAsync(id, 1, { ignoreAliases: false });
    try { t = await reader.getPageAsync(0); } finally { await reader.releaseAsync(); }
  } else t = await ds.getLogicalTableDataAsync(id, { ignoreAliases: false, maxRows: 1 });
  return ((t && t.columns) || []).map(c => tfNorm(c.fieldName));
}

/**
 * The rows of the data source a worksheet draws the field from: one logical table – of related tables (Sheet1 +
 * Vacency), the one that holds the field – at most FORMAT_CONFIG.slicerMaxRows rows. Another table's fields are not
 * in these rows: their cards stay cards, and a visual that needs them keeps fixed values (the check against Tableau's
 * numbers decides).
 * @param {any} ws a Tableau worksheet @param {string} fieldName @returns {Promise<SlicerSource>}
 * @typedef {{ id: string, name: string, columns: { name: string, dataType: string }[], rows: DataValue[][] }} SlicerSource
 */
async function readSlicerSource(ws, fieldName) {
  const ds = await slicerDataSourceOf(ws, fieldName);
  if (!ds) throw new Error("Tableau does not give this worksheet's data source here");
  if (typeof ds.getLogicalTablesAsync !== "function") throw new Error("the worksheet's data source cannot be read");
  const tables = await ds.getLogicalTablesAsync();
  if (!tables.length) throw new Error(`"${ds.name}" has no tables`);
  let tableId = tables[0].id;
  if (tables.length > 1) {
    const name = tfNorm(fieldName);
    const cols = [];
    for (const t of tables) cols.push(await slicerTableColumns(ds, t.id));
    let k = cols.findIndex(c => c.includes(name));
    // a calculated field is in no table: the table with the most columns – the facts the related tables describe
    if (k < 0) k = cols.reduce((best, c, i) => (c.length > cols[best].length ? i : best), 0);
    tableId = tables[k].id;
    console.log(`[Slicers] ${ds.name}: ${tables.length} related tables – "${tables[k].caption || tableId}" for "${fieldName}"`);
  }
  const max = FORMAT_CONFIG.slicerMaxRows;
  const tooMany = n => new Error(`"${ds.name}" has ${n} rows, more than FORMAT_CONFIG.slicerMaxRows`);
  let table;
  if (typeof ds.getLogicalTableDataReaderAsync === "function") {
    const reader = await ds.getLogicalTableDataReaderAsync(tableId, 10000, { ignoreAliases: false });
    try {
      if (reader.totalRowCount > max) throw tooMany(reader.totalRowCount);
      table = await reader.getAllPagesAsync(max);
    } finally { await reader.releaseAsync(); }
  } else if (typeof ds.getLogicalTableDataAsync === "function") {
    table = await ds.getLogicalTableDataAsync(tableId, { ignoreAliases: false, maxRows: max + 1 });
    if (table.data.length > max) throw tooMany(`over ${max}`);
  } else {
    throw new Error("this Tableau version cannot read a data source's rows");
  }
  console.log(`[Slicers] ${ds.name}: ${table.data.length} rows, ${table.columns.length} columns`);
  return { id: ds.id, name: ds.name, rows: table.data,
           columns: table.columns.map(c => ({ name: c.fieldName, dataType: String(c.dataType || "").toLowerCase() })) };
}

/** a workbook field → its column in the source rows (by caption, else by name); -1 when the rows do not have it. A
 * calculated field the rows lack is computed row by row when it can be (slicerCalcColumn) and becomes a column of
 * its own; the check against Tableau's numbers decides whether it is right.
 * @param {SlicerSource} source @param {FormatModel} model @param {FieldRef} ref */
function slicerColumnOf(source, model, ref) {
  const info = tfFieldInfo(model, { ...ref, deriv: null });
  const names = [info && info.caption, ref.name].filter(Boolean).map(tfNorm);
  const col = source.columns.findIndex(c => names.includes(tfNorm(c.name)));
  return col >= 0 || !info || !(info.formulaFull || info.formula) || info.param ? col
    : slicerCalcColumn(source, model, info.caption || info.name, info.formulaFull || info.formula, info.ds);
}

/* ── calculations computed row by row ─────────────────────────────────────────────────────────────────────────
 * A row-level calculated field ([Category] + " - " + [Sub-Category], IF YEAR([Order Date]) = 2022 THEN [Sales] END)
 * and a FIXED level of detail ({FIXED [Customer ID] : COUNT([Order ID])}) are computed for every source row and added
 * to the source as columns: slicers, mark keys and the SUMIFS-style formulas then use them like any other column.
 * A FIXED value is computed over all the rows – as Tableau computes it before the sheet's dimension filters (a slicer
 * is one), so it stays as it is when a slicer changes. Parameters take their value at export. Anything else (table
 * calculations, INCLUDE / EXCLUDE, unknown functions) is not computed: the field stays out of the rows. */

/** a calculation's tokens: fields ([ds].[name]), strings, #dates#, numbers, words, operators */
function slicerTokenize(formula) {
  const s = String(formula).replace(/\/\*[\s\S]*?\*\//g, " ").replace(/\/\/[^\n]*/g, " ");
  const re = /\s+|((?:\[[^\]]*\]\.)?\[[^\]]*\])|('(?:[^']|'')*'|"(?:[^"]|"")*")|(#[^#]*#)|(\d+(?:\.\d+)?(?:[eE][-+]?\d+)?)|([A-Za-z_][A-Za-z_0-9]*)|(<=|>=|<>|!=|==|&&|\|\|)|([-+*/%=<>(),{}:^])/y;
  const out = [];
  let m;
  re.lastIndex = 0;
  while (re.lastIndex < s.length && (m = re.exec(s))) {
    if (m[0].trim()) out.push(m[0]);
  }
  if (re.lastIndex < s.length) return null;                    // a character it does not know
  return out;
}

/** tokens → a syntax tree (null when it is not a calculation this evaluator reads) */
function slicerParseRowCalc(tokens) {
  if (!tokens) return null;
  let i = 0;
  const up = () => String(tokens[i] || "").toUpperCase();
  const fail = () => { throw new Error("calc"); };
  const take = t => { if (up() !== t) fail(); i++; };
  const expr = () => or();
  const or = () => { let a = and(); while (up() === "OR" || tokens[i] === "||") { i++; a = { t: "or", a, b: and() }; } return a; };
  const and = () => { let a = not(); while (up() === "AND" || tokens[i] === "&&") { i++; a = { t: "and", a, b: not() }; } return a; };
  const not = () => (up() === "NOT" ? (i++, { t: "not", a: not() }) : cmp());
  const cmp = () => {
    const a = add();
    const op = tokens[i];
    if (["=", "==", "!=", "<>", "<", "<=", ">", ">="].includes(op)) { i++; return { t: "cmp", op: op === "==" ? "=" : op === "<>" ? "!=" : op, a, b: add() }; }
    return a;
  };
  const add = () => { let a = mul(); while (tokens[i] === "+" || tokens[i] === "-") { const op = tokens[i++]; a = { t: "bin", op, a, b: mul() }; } return a; };
  const mul = () => { let a = unary(); while (tokens[i] === "*" || tokens[i] === "/" || tokens[i] === "%" || tokens[i] === "^") { const op = tokens[i++]; a = { t: "bin", op, a, b: unary() }; } return a; };
  const unary = () => (tokens[i] === "-" ? (i++, { t: "neg", a: unary() }) : tokens[i] === "+" ? (i++, unary()) : primary());
  const primary = () => {
    const t = tokens[i++];
    if (t === undefined) fail();
    if (t === "(") { const e = expr(); take(")"); return e; }
    if (/^\d/.test(t)) return { t: "lit", v: Number(t) };
    if (t[0] === "'" || t[0] === '"') return { t: "lit", v: t.slice(1, -1).replace(t[0] === "'" ? /''/g : /""/g, t[0]) };
    if (t[0] === "#") { const d = new Date(t.slice(1, -1).trim().replace(" ", "T") + (/T|:/.test(t) ? "" : "T00:00:00") + "Z"); return isNaN(+d) ? fail() : { t: "lit", v: d }; }
    if (t[0] === "[") return { t: "field", raw: t };
    if (t === "{") {                                             // {FIXED [a], [b] : AGG(…)} – FIXED only
      const kind = up();
      if (kind === "INCLUDE" || kind === "EXCLUDE") fail();
      if (kind !== "FIXED") { const body = expr(); take("}"); return { t: "lod", dims: [], body }; }   // {MAX(…)}: the whole table
      i++;
      const dims = [];
      while (tokens[i] !== ":") { dims.push(expr()); if (tokens[i] === ",") i++; else if (tokens[i] !== ":") fail(); }
      i++;
      const body = expr();
      take("}");
      return { t: "lod", dims, body };
    }
    const w = t.toUpperCase();
    if (w === "TRUE" || w === "FALSE") return { t: "lit", v: w === "TRUE" };
    if (w === "NULL") return { t: "lit", v: null };
    if (w === "IF") {
      const branches = [];
      let c = expr(); take("THEN"); branches.push([c, expr()]);
      while (up() === "ELSEIF") { i++; c = expr(); take("THEN"); branches.push([c, expr()]); }
      let otherwise = null;
      if (up() === "ELSE") { i++; otherwise = expr(); }
      take("END");
      return { t: "if", branches, otherwise };
    }
    if (w === "CASE") {
      const subject = expr(), whens = [];
      while (up() === "WHEN") { i++; const v = expr(); take("THEN"); whens.push([v, expr()]); }
      let otherwise = null;
      if (up() === "ELSE") { i++; otherwise = expr(); }
      take("END");
      return { t: "case", subject, whens, otherwise };
    }
    if (/^[A-Za-z_]/.test(t) && tokens[i] === "(") {
      i++;
      const args = [];
      while (tokens[i] !== ")") { args.push(expr()); if (tokens[i] === ",") i++; else if (tokens[i] !== ")") fail(); }
      i++;
      return { t: "fn", name: w, args };
    }
    return fail();
  };
  try {
    const e = expr();
    return i === tokens.length ? e : null;
  } catch (err) { return null; }
}

const SLICER_ROW_AGGS = new Set(["SUM", "AVG", "MIN", "MAX", "COUNT", "COUNTD", "ATTR", "MEDIAN"]);
const SLICER_DAY_MS = 86400000;

/**
 * A calculation computed for every source row, added to the source as a column; its index, or -1 when it cannot be
 * (it is remembered either way). key: the column's name (a field's caption, or the inline expression's text).
 * @param {SlicerSource} source @param {FormatModel} model @param {string} key @param {string} formula @param {string} [ds]
 */
function slicerCalcColumn(source, model, key, formula, ds) {
  const S = /** @type {any} */ (source);
  S.calcCols = S.calcCols || new Map();
  if (S.calcCols.has(key)) return S.calcCols.get(key);
  S.calcCols.set(key, -1);                                     // a calculation that refers to itself stays out
  const ast = slicerParseRowCalc(slicerTokenize(formula));
  if (!ast) return -1;
  const rows = source.rows;
  /** a source value as the calculation sees it: a number, string, boolean, date (UTC midnight for a date) or null */
  const valueOf = (dv, type) => {
    if (tfIsNull(dv)) return null;
    if (/date/.test(type)) {
      const s = tvExcelSerial(dv);
      if (s === null) return null;
      return new Date(Math.round((/time/.test(type) ? s : Math.round(s)) * SLICER_DAY_MS - 25569 * SLICER_DAY_MS));
    }
    if (/^bool/.test(type)) return dv.value === true || /^true$/i.test(String(dv.value));
    const v = dv.nativeValue !== undefined ? dv.nativeValue : dv.value;
    if (/^(int|integer|float|real|number|double)/.test(type)) return tfDvNum(dv);
    return v === null || v === undefined ? null : String(v);
  };
  /** a field of the calculation → (row → value); a parameter → its value at export */
  const fieldFn = raw => {
    const ref = tfParseFieldRef(raw.includes("].[") ? raw : ds ? `[${ds}].${raw}` : raw);
    if (!ref) throw new Error("field");
    if (/^parameters$/i.test(ref.ds || "") || /^\[Parameters\]/i.test(raw)) {
      const p = model.fields["Parameters|" + ref.name.toLowerCase()];
      if (!p) throw new Error("param");
      const lit = slicerParseRowCalc(slicerTokenize(String(p.value || "")));
      const v = lit && lit.t === "lit" ? lit.v : p.value === undefined ? null : Number(p.value);
      return () => v;
    }
    const col = slicerColumnOf(source, model, { ...ref, deriv: "none" });
    if (col < 0) throw new Error("field");
    const type = source.columns[col].dataType;
    return r => valueOf(rows[r][col], type);
  };
  const num = v => (v === null || v === undefined ? null : v instanceof Date ? v.getTime() / SLICER_DAY_MS + 25569 : typeof v === "boolean" ? (v ? 1 : 0) : typeof v === "number" ? v : isFinite(Number(v)) && String(v).trim() !== "" ? Number(v) : null);
  const truthy = v => v === true || (typeof v === "number" && v !== 0);
  const eq = (a, b) => (a instanceof Date || b instanceof Date ? num(a) === num(b) : typeof a === "number" || typeof b === "number" ? num(a) === num(b) : String(a) === String(b));
  const dayOf = v => (v instanceof Date ? v : v === null ? null : typeof v === "number" ? new Date((v - 25569) * SLICER_DAY_MS) : new Date(String(v)));
  const trunc = (part, d) => {
    if (!d) return null;
    const y = d.getUTCFullYear(), m = d.getUTCMonth();
    switch (part) {
      case "year": return new Date(Date.UTC(y, 0, 1));
      case "quarter": return new Date(Date.UTC(y, m - (m % 3), 1));
      case "month": return new Date(Date.UTC(y, m, 1));
      case "week": return new Date(Date.UTC(y, m, d.getUTCDate() - d.getUTCDay()));
      case "day": return new Date(Date.UTC(y, m, d.getUTCDate()));
      default: throw new Error("datepart");
    }
  };
  const part = (p, d) => {
    if (!d) return null;
    switch (p) {
      case "year": return d.getUTCFullYear();
      case "quarter": return Math.floor(d.getUTCMonth() / 3) + 1;
      case "month": return d.getUTCMonth() + 1;
      case "day": return d.getUTCDate();
      case "weekday": return d.getUTCDay() + 1;
      case "week": { const j = Date.UTC(d.getUTCFullYear(), 0, 1); return Math.floor(((+trunc("week", d) - j) / SLICER_DAY_MS + new Date(j).getUTCDay()) / 7) + 1; }
      default: throw new Error("datepart");
    }
  };
  const lower = v => String(v).toLowerCase().replace(/^'|'$/g, "");
  /** node → (row, group) → value; group: the rows an aggregate inside a FIXED computes over */
  const compile = n => {
    switch (n.t) {
      case "lit": return () => n.v;
      case "field": return fieldFn(n.raw);
      case "neg": { const a = compile(n.a); return (r, g) => { const v = num(a(r, g)); return v === null ? null : -v; }; }
      case "not": { const a = compile(n.a); return (r, g) => { const v = a(r, g); return v === null ? null : !truthy(v); }; }
      case "and": { const a = compile(n.a), b = compile(n.b); return (r, g) => truthy(a(r, g)) && truthy(b(r, g)); }
      case "or": { const a = compile(n.a), b = compile(n.b); return (r, g) => truthy(a(r, g)) || truthy(b(r, g)); }
      case "cmp": {
        const a = compile(n.a), b = compile(n.b);
        return (r, g) => {
          const x = a(r, g), y = b(r, g);
          if (x === null || y === null) return null;
          if (n.op === "=") return eq(x, y);
          if (n.op === "!=") return !eq(x, y);
          const p = typeof x === "string" && typeof y === "string" ? x.localeCompare(y) : num(x) - num(y);
          return n.op === "<" ? p < 0 : n.op === "<=" ? p <= 0 : n.op === ">" ? p > 0 : p >= 0;
        };
      }
      case "bin": {
        const a = compile(n.a), b = compile(n.b);
        return (r, g) => {
          const x = a(r, g), y = b(r, g);
          if (x === null || y === null) return null;
          if (n.op === "+" && (typeof x === "string" || typeof y === "string")) return String(x) + String(y);
          if (x instanceof Date && n.op === "-" && y instanceof Date) return (x.getTime() - y.getTime()) / SLICER_DAY_MS;
          if (x instanceof Date && (n.op === "+" || n.op === "-")) return new Date(x.getTime() + (n.op === "+" ? 1 : -1) * num(y) * SLICER_DAY_MS);
          const p = num(x), q = num(y);
          if (p === null || q === null) return null;
          switch (n.op) {
            case "+": return p + q;
            case "-": return p - q;
            case "*": return p * q;
            case "/": return q === 0 ? null : p / q;
            case "%": return q === 0 ? null : p % q;
            default: return Math.pow(p, q);
          }
        };
      }
      case "if": {
        const br = n.branches.map(([c, v]) => [compile(c), compile(v)]), o = n.otherwise ? compile(n.otherwise) : () => null;
        return (r, g) => { for (const [c, v] of br) if (truthy(c(r, g))) return v(r, g); return o(r, g); };
      }
      case "case": {
        const s = compile(n.subject), wh = n.whens.map(([v, x]) => [compile(v), compile(x)]), o = n.otherwise ? compile(n.otherwise) : () => null;
        return (r, g) => { const x = s(r, g); for (const [v, y] of wh) if (x !== null && eq(x, v(r, g))) return y(r, g); return o(r, g); };
      }
      case "lod": {
        // {FIXED dims : body}: the body's aggregates over every row with the same dims, worked out once
        const dims = n.dims.map(compile), body = compile(n.body);
        let byRow = null;
        return r => {
          if (!byRow) {
            const groups = new Map(), keyOf = q => dims.map(d => { const v = d(q); return v instanceof Date ? +v : String(v); }).join("\u0001");
            rows.forEach((_, q) => { const k = keyOf(q); if (!groups.has(k)) groups.set(k, []); groups.get(k).push(q); });
            const val = new Map();
            groups.forEach((g, k) => val.set(k, body(-1, g)));
            byRow = rows.map((_, q) => val.get(keyOf(q)));
          }
          return byRow[r];
        };
      }
      case "fn": return compileFn(n);
      default: throw new Error("node");
    }
  };
  const compileFn = n => {
    const A = n.args.map(a => (a.t === "lit" ? () => a.v : null));
    const args = () => n.args.map(compile);
    if (SLICER_ROW_AGGS.has(n.name)) {
      if (n.args.length !== 1) throw new Error("agg");
      const a = compile(n.args[0]);
      return (r, g) => {
        if (!g) throw new Error("aggregate at row level");   // an aggregate calculation: not computed row by row
        const vals = g.map(q => a(q, null)).filter(v => v !== null && v !== undefined);
        if (n.name === "COUNT") return vals.length;
        if (n.name === "COUNTD") return new Set(vals.map(v => (v instanceof Date ? +v : String(v)))).size;
        if (n.name === "ATTR") return new Set(vals.map(v => (v instanceof Date ? +v : String(v)))).size === 1 ? vals[0] : null;
        if (!vals.length) return null;
        if (n.name === "MIN" || n.name === "MAX") {
          const dates = vals[0] instanceof Date, ns = vals.map(num);
          const m = n.name === "MIN" ? Math.min(...ns) : Math.max(...ns);
          return dates ? new Date((m - 25569) * SLICER_DAY_MS) : m;
        }
        const ns = vals.map(num).filter(v => v !== null);
        if (n.name === "SUM") return ns.reduce((s, v) => s + v, 0);
        if (n.name === "AVG") return ns.reduce((s, v) => s + v, 0) / ns.length;
        const sorted = ns.sort((x, y) => x - y), h = Math.floor(sorted.length / 2);
        return sorted.length % 2 ? sorted[h] : (sorted[h - 1] + sorted[h]) / 2;
      };
    }
    const f = args();
    const one = fn => (r, g) => { const v = f[0](r, g); return v === null ? null : fn(v, r, g); };
    const lit = k => { const v = A[k] ? A[k]() : null; if (v === null) throw new Error("literal"); return lower(v); };
    switch (n.name) {
      case "ZN": return (r, g) => num(f[0](r, g)) ?? 0;
      case "IFNULL": return (r, g) => { const v = f[0](r, g); return v === null ? f[1](r, g) : v; };
      case "ISNULL": return (r, g) => f[0](r, g) === null;
      case "IIF": return (r, g) => { const c = f[0](r, g); return c === null ? (f[3] ? f[3](r, g) : null) : truthy(c) ? f[1](r, g) : f[2](r, g); };
      case "ABS": return one(v => Math.abs(num(v)));
      case "SIGN": return one(v => Math.sign(num(v)));
      case "SQRT": return one(v => (num(v) < 0 ? null : Math.sqrt(num(v))));
      case "ROUND": return (r, g) => { const v = num(f[0](r, g)), d = f[1] ? num(f[1](r, g)) || 0 : 0; return v === null ? null : Math.round(v * 10 ** d) / 10 ** d; };
      case "INT": return one(v => Math.trunc(num(v)));
      case "FLOOR": return one(v => Math.floor(num(v)));
      case "CEILING": return one(v => Math.ceil(num(v)));
      case "FLOAT": return one(v => num(v));
      case "POWER": return (r, g) => { const a = num(f[0](r, g)), b = num(f[1](r, g)); return a === null || b === null ? null : a ** b; };
      case "DIV": return (r, g) => { const a = num(f[0](r, g)), b = num(f[1](r, g)); return a === null || !b ? null : Math.trunc(a / b); };
      case "STR": return one(v => (v instanceof Date ? v.toISOString().slice(0, 10) : typeof v === "number" ? String(+v.toPrecision(15)) : String(v)));
      case "UPPER": return one(v => String(v).toUpperCase());
      case "LOWER": return one(v => String(v).toLowerCase());
      case "TRIM": return one(v => String(v).trim());
      case "LEN": return one(v => String(v).length);
      case "LEFT": return (r, g) => { const v = f[0](r, g); return v === null ? null : String(v).slice(0, num(f[1](r, g))); };
      case "RIGHT": return (r, g) => { const v = f[0](r, g), k = num(f[1](r, g)); return v === null ? null : k > 0 ? String(v).slice(-k) : ""; };
      case "MID": return (r, g) => { const v = f[0](r, g); if (v === null) return null; const s = num(f[1](r, g)) - 1; return f[2] ? String(v).substr(s, num(f[2](r, g))) : String(v).slice(s); };
      case "CONTAINS": return (r, g) => { const v = f[0](r, g), s = f[1](r, g); return v === null || s === null ? null : String(v).includes(String(s)); };
      case "STARTSWITH": return (r, g) => { const v = f[0](r, g), s = f[1](r, g); return v === null || s === null ? null : String(v).startsWith(String(s)); };
      case "ENDSWITH": return (r, g) => { const v = f[0](r, g), s = f[1](r, g); return v === null || s === null ? null : String(v).endsWith(String(s)); };
      case "YEAR": return one(v => part("year", dayOf(v)));
      case "QUARTER": return one(v => part("quarter", dayOf(v)));
      case "MONTH": return one(v => part("month", dayOf(v)));
      case "DAY": return one(v => part("day", dayOf(v)));
      case "DATE": return one(v => trunc("day", dayOf(v)));
      case "DATETRUNC": { const p = lit(0); return (r, g) => trunc(p, dayOf(f[1](r, g))); }
      case "DATEPART": { const p = lit(0); return (r, g) => part(p, dayOf(f[1](r, g))); }
      case "DATEDIFF": {
        const p = lit(0);
        return (r, g) => {
          const a = dayOf(f[1](r, g)), b = dayOf(f[2](r, g));
          if (!a || !b) return null;
          if (p === "day") return Math.round((+trunc("day", b) - +trunc("day", a)) / SLICER_DAY_MS);
          if (p === "week") return Math.round((+trunc("week", b) - +trunc("week", a)) / (7 * SLICER_DAY_MS));
          const months = (b.getUTCFullYear() - a.getUTCFullYear()) * 12 + b.getUTCMonth() - a.getUTCMonth();
          return p === "month" ? months : p === "quarter" ? Math.floor((part("quarter", b) - part("quarter", a)) + (b.getUTCFullYear() - a.getUTCFullYear()) * 4) : p === "year" ? b.getUTCFullYear() - a.getUTCFullYear() : (() => { throw new Error("datepart"); })();
        };
      }
      case "DATEADD": {
        const p = lit(0);
        return (r, g) => {
          const k = num(f[1](r, g)), d = dayOf(f[2](r, g));
          if (k === null || !d) return null;
          if (p === "day") return new Date(+d + k * SLICER_DAY_MS);
          if (p === "week") return new Date(+d + 7 * k * SLICER_DAY_MS);
          const m = p === "month" ? k : p === "quarter" ? 3 * k : p === "year" ? 12 * k : (() => { throw new Error("datepart"); })();
          return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + m, d.getUTCDate()));
        };
      }
      default: throw new Error(`function ${n.name}`);
    }
  };
  let values;
  try {
    const fn = compile(ast);
    values = rows.map((_, r) => fn(r, null));
  } catch (err) {
    return -1;                                                  // an aggregate calculation, or something not computed
  }
  const kinds = new Set(values.filter(v => v !== null && v !== undefined).map(v => (v instanceof Date ? "date" : typeof v)));
  const type = kinds.size === 1 ? { date: "date", number: "float", boolean: "boolean", string: "string" }[[...kinds][0]] : kinds.size ? "string" : "float";
  const col = source.columns.length;
  source.columns.push({ name: key, dataType: type, calc: true });
  rows.forEach((row, r) => {
    const v = values[r];
    row[col] = v === null || v === undefined ? { value: null, nativeValue: null, formattedValue: "Null" }
      : v instanceof Date ? { value: v.toISOString().slice(0, 10), nativeValue: v, formattedValue: v.toISOString().slice(0, 10) }
      : { value: type === "string" ? String(v) : v, nativeValue: type === "string" ? String(v) : v, formattedValue: String(v) };
  });
  S.calcCols.set(key, col);
  return col;
}

/** a filter's field name ("Region", "YEAR(Order Date)") → the field and its date part */
function slicerFilterField(fieldName) {
  const m = String(fieldName || "").match(/^(YEAR|QUARTER|MONTH|DAY)\((.+)\)$/i);
  const deriv = m ? { year: "yr", quarter: "qr", month: "mn", day: "dy" }[m[1].toLowerCase()] : "none";
  const name = m ? m[2] : String(fieldName || "");
  return { ref: /** @type {FieldRef} */ ({ raw: name, ds: null, deriv, name, type: "nk", inner: name }), deriv };
}

/**
 * A measure as aggregations of source columns: [sum:Sales] → SUM(Sales); a calculated field of aggregations
 * (SUM([Profit])/SUM([Sales]), SUM([Sales])/COUNTD([Customer Name])) → its expression; null for anything else
 * (table calculations, LOD expressions, row-level calculations, functions other than + − × ÷ and ZN).
 * @param {SlicerSource} source @param {FormatModel} model @param {FieldRef | null} ref
 * @returns {SlicerExpr | null}
 * @typedef {{ op: string, fn?: string, col?: number, v?: number, a?: SlicerExpr, b?: SlicerExpr }} SlicerExpr
 */
function slicerMeasureExpr(source, model, ref) {
  if (!ref) return null;
  const deriv = String(ref.deriv || "").toLowerCase();
  // percent of total (a quick table calculation over the whole view): the mark's value over every mark's
  if (/^pcto/.test(deriv)) {
    const base = tfTableCalcBase(ref), e = base && slicerMeasureExpr(source, model, base);
    return e && !JSON.stringify(e).includes('"COUNTD"') ? { op: "pct", a: e } : null;
  }
  if (SLICER_AGGS[deriv]) {
    const col = slicerColumnOf(source, model, ref);
    return col >= 0 ? { op: "agg", fn: SLICER_AGGS[deriv], col } : null;
  }
  const info = deriv === "usr" ? tfFieldInfo(model, ref) : null;
  return info && info.formula ? slicerParseCalc(source, model, info.formula, 0) : null;
}

/** a Tableau calculation of aggregations → SlicerExpr, null when it holds anything else */
function slicerParseCalc(source, model, formula, depth) {
  if (depth > 4) return null;
  const tokens = slicerTokenize(formula);
  if (!tokens) return null;
  let i = 0;
  const fail = () => { throw new Error("not translatable"); };
  const expect = t => { if (tokens[i++] !== t) fail(); };
  /** @returns {SlicerExpr} */
  const expr = () => {
    let a = term();
    while (tokens[i] === "+" || tokens[i] === "-") { const op = tokens[i++]; a = { op, a, b: term() }; }
    return a;
  };
  const term = () => {
    let a = unary();
    while (tokens[i] === "*" || tokens[i] === "/") { const op = tokens[i++]; a = { op, a, b: unary() }; }
    return a;
  };
  const unary = () => (tokens[i] === "-" ? (i++, { op: "neg", a: unary() }) : primary());
  const primary = () => {
    const t = tokens[i++];
    if (t === undefined) fail();
    if (t === "(") { const e = expr(); expect(")"); return e; }
    if (/^\d/.test(t)) return { op: "num", v: Number(t) };
    if (t[0] === "[") {                       // another calculated field – of aggregations, else not translatable
      const info = tfFieldInfo(model, tfParseFieldRef(t));
      const e = info && info.formula ? slicerParseCalc(source, model, info.formula, depth + 1) : null;
      return e || fail();
    }
    const fn = t.toUpperCase();
    expect("(");
    if (fn === "ZN") { const e = expr(); expect(")"); return { op: "zn", a: e }; }
    // TOTAL(SUM([Sales])): the aggregate over every mark of the visual (a table calculation over the whole view)
    if (fn === "TOTAL") { const e = expr(); expect(")"); return JSON.stringify(e).includes('"COUNTD"') ? fail() : { op: "total", a: e }; }
    const agg = { SUM: "SUM", AVG: "AVG", MIN: "MIN", MAX: "MAX", COUNT: "COUNT", COUNTD: "COUNTD" }[fn];
    if (!agg) fail();
    if (tokens[i] && tokens[i][0] === "[" && tokens[i + 1] === ")") {      // SUM([Sales])
      const field = tokens[i++];
      expect(")");
      const col = slicerColumnOf(source, model, tfParseFieldRef(field));
      return col >= 0 ? { op: "agg", fn: agg, col } : fail();
    }
    // SUM(IF [Year] = 2022 THEN [Sales] END), MIN(0): the expression computed row by row (slicerCalcColumn)
    const start = i;
    for (let depthP = 0; i < tokens.length; i++) {
      if (tokens[i] === "(" || tokens[i] === "{") depthP++;
      else if (tokens[i] === ")" || tokens[i] === "}") { if (!depthP) break; depthP--; }
    }
    const text = tokens.slice(start, i).join(" ");
    expect(")");
    const col = slicerCalcColumn(source, model, `(${text})`, text);
    return col >= 0 ? { op: "agg", fn: agg, col } : fail();
  };
  try {
    const e = expr();
    return i === tokens.length ? e : null;
  } catch (err) { return null; }
}

/** the expression's value over some source rows, as Tableau computes it (null: no value) @param {SlicerExpr} e
 * @param {number[]} rows indexes into source.rows @param {SlicerSource} source @returns {number | null} */
function slicerEval(e, rows, source, all) {
  if (e.op === "num") return /** @type {number} */ (e.v);
  if (e.op === "zn") return slicerEval(/** @type {SlicerExpr} */ (e.a), rows, source, all) ?? 0;
  if (e.op === "neg") { const a = slicerEval(/** @type {SlicerExpr} */ (e.a), rows, source, all); return a === null ? null : -a; }
  if (e.op === "total") return all ? slicerEval(/** @type {SlicerExpr} */ (e.a), all, source, all) : null;
  if (e.op === "pct") {                                        // all: the rows of every mark of the visual
    const a = slicerEval(/** @type {SlicerExpr} */ (e.a), rows, source, all), t = all ? slicerEval(/** @type {SlicerExpr} */ (e.a), all, source, all) : null;
    return a === null || !t ? null : a / t;
  }
  if (e.op === "agg") {
    const dvs = rows.map(r => source.rows[r][/** @type {number} */ (e.col)]).filter(dv => !tfIsNull(dv));
    if (e.fn === "COUNT") return dvs.length;
    if (e.fn === "COUNTD") return new Set(dvs.map(dv => String(dv.nativeValue !== undefined ? dv.nativeValue : dv.value))).size;
    const nums = dvs.map(tfDvNum).filter(n => n !== null);
    if (!nums.length) return null;
    if (e.fn === "SUM") return nums.reduce((s, n) => s + n, 0);
    if (e.fn === "AVG") return nums.reduce((s, n) => s + n, 0) / nums.length;
    return e.fn === "MIN" ? Math.min(...nums) : Math.max(...nums);
  }
  const a = slicerEval(/** @type {SlicerExpr} */ (e.a), rows, source, all), b = slicerEval(/** @type {SlicerExpr} */ (e.b), rows, source, all);
  if (a === null || b === null) return null;
  if (e.op === "/") return b === 0 ? null : a / b;
  return e.op === "+" ? a + b : e.op === "-" ? a - b : a * b;
}

/** the expression as an Excel formula over the rows of key k that the slicers leave visible
 * @param {SlicerExpr} e @param {{ col: (c: number) => string, key: string, vis: string }} R ranges @param {string} k */
function slicerExcel(e, R, k) {
  if (e.op === "num") return String(e.v);
  if (e.op === "zn") return `IFERROR(${slicerExcel(/** @type {SlicerExpr} */ (e.a), R, k)},0)`;
  if (e.op === "neg") return `-(${slicerExcel(/** @type {SlicerExpr} */ (e.a), R, k)})`;
  // percent of total: over the visible rows of every mark of the visual (any key)
  if (e.op === "pct") return `(${slicerExcel(/** @type {SlicerExpr} */ (e.a), R, k)})/(${slicerExcel(/** @type {SlicerExpr} */ (e.a), R, '">0"')})`;
  if (e.op === "total") return `(${slicerExcel(/** @type {SlicerExpr} */ (e.a), R, '">0"')})`;
  if (e.op === "agg") {
    const F = R.col(/** @type {number} */ (e.col)), on = `${R.key},${k},${R.vis},1`;
    switch (e.fn) {
      case "SUM": return `SUMIFS(${F},${on})`;
      case "AVG": return `AVERAGEIFS(${F},${on})`;
      case "MIN": return `_xlfn.MINIFS(${F},${on})`;
      case "MAX": return `_xlfn.MAXIFS(${F},${on})`;
      case "COUNT": return `COUNTIFS(${on},${F},"<>")`;
      default: return `IFERROR(ROWS(_xlfn.UNIQUE(_xlfn._xlws.FILTER(${F},(${R.key}=${k})*(${R.vis}=1)*(${F}<>"")))),0)`;
    }
  }
  return `(${slicerExcel(/** @type {SlicerExpr} */ (e.a), R, k)}${e.op}${slicerExcel(/** @type {SlicerExpr} */ (e.b), R, k)})`;
}

/** an expression that can fail in Excel (a division, an average, min or max of no rows) */
const slicerMayFail = e => !!e && (e.op === "/" || e.op === "pct" || (e.op === "agg" && /^(AVG|MIN|MAX)$/.test(e.fn || "")) ||
                                   slicerMayFail(e.a) || slicerMayFail(e.b));

/**
 * One export's slicers, shared by every dashboard of the file: the source rows, the slicer fields with Tableau's
 * selection, the columns of the hidden table, and the visuals bound to it.
 */
function createSlicerModel() {
  /** @type {SlicerSource | null} */
  let source = null;
  /**
   * kind "slicer": an Excel slicer (selected: the items ticked); "list": a dropdown cell (current: its value, null =
   * (All)); "range": From / To date cells (lo / hi: day serials). pass(r): row r is in the selection. control: the
   * cells the dashboard sheet holds for a list / range, set when its card is written.
   * @type {{ kind: string, col: number, deriv: string, isDate: boolean, header: string, values: any[], items: string[],
   *          selected: Set<string> | null, pass: (r: number) => boolean, tableCol: number, [k: string]: any }[]}
   */
  const fields = [];
  /**
   * Ranges of an aggregate (SUM(Sales) from … to, a Profit Ratio of at least …): Tableau keeps the marks whose value
   * is in them – not rows. Their Min / Max cells (control) are read through cells of the hidden lists sheet (loRef /
   * hiRef, known before the card is written), which the bound visuals' values test (bind: keepMarks).
   * @type {{ kind: "marks", ref: FieldRef, header: string, lo: number | null, hi: number | null, loRef: string, hiRef: string, control: any }[]}
   */
  const markFields = [];
  /** cells of the lists sheet that show a card's cell (a marks range's Min / Max) @type {{ field: any, end: string }[]} */
  const boundCells = [];
  const bound = (field, end) => { boundCells.push({ field, end }); return sheetRef(SLICER_LISTS_SHEET, `$${SLICER_BOUND_COL}$${boundCells.length}`); };
  /** extra table columns after Row and Visible @type {{ header: string, values: any[] }[]} */
  const columns = [];
  const headers = new Set(["row", "visible"]);
  const fieldCols = new Map();              // source column → table column (measures' fields)
  let locked = false;                       // a visual is bound: the slicer fields are final
  let failed = null;                        // why the data source could not be read
  const uniqueHeader = name => {
    let h = String(name || "Column").replace(/[\r\n]+/g, " ").trim() || "Column", n = 2;
    while (headers.has(h.toLowerCase())) h = `${name} (${n++})`;
    headers.add(h.toLowerCase());
    return h;
  };
  const addColumn = (header, values) => { columns.push({ header: uniqueHeader(header), values }); return columns.length + 1; };
  const N = () => (source ? source.rows.length : 0);
  const range = c => cellRef(SLICER_DATA_SHEET, c, 1, c, N());
  const visible = r => fields.every(f => f.pass(r));
  const shownBySlicers = r => fields.every(f => f.kind !== "slicer" || f.pass(r));     // the rows the table's filter hides
  let lists = 0;                            // dropdown fields: one column each on the lists sheet
  /** other dropdowns' lists (a date's year / month / day) on the lists sheet @type {{ col: number, values: any[] }[]} */
  const extraLists = [];
  /** the dates the date parameters' dropdowns make (writeParamDateCard), in a column of the lists sheet of their own
   * @type {{ formula: string, result: Date }[]} */
  const dateCells = [];

  /** the table column of a source column (written the first time a formula needs it) */
  const fieldColumn = col => {
    if (!fieldCols.has(col)) {
      const c = /** @type {SlicerSource} */ (source).columns[col];
      const numeric = /^(int|integer|float|real|number|double)/.test(c.dataType);
      const values = /** @type {SlicerSource} */ (source).rows.map(r => tfIsNull(r[col]) ? null
        : numeric ? tfDvNum(r[col]) : slicerRowText(r[col], "none", /date/.test(c.dataType)));
      fieldCols.set(col, addColumn(c.name, values));
    }
    return /** @type {number} */ (fieldCols.get(col));
  };

  /**
   * The values of source column col kept between two bounds → a "range" field: days of a date column (From / To
   * cells), numbers of a number column (Min / Max cells). loDv / hiDv: DataValues of its bounds, missing = the data's
   * first / last value (an "at least" / "at most" filter). A reason (string) when the column has neither. opts:
   * includeNull, loStrict / hiStrict (> / < instead of >= / <=), calc (the calculated filter it stands for), relative
   * (the relative date filter's text it stands for).
   */
  const rangeField = (col, name, loDv, hiDv, opts) => {
    // dates: the days of the rows (a date at midnight may arrive in local time: its nearest day; a date-time: its day)
    const type = /** @type {SlicerSource} */ (source).columns[col].dataType;
    const isDate = /date/.test(type), timed = /time/.test(type);
    const valueOf = dv => {
      if (!dv || tfIsNull(dv)) return null;
      if (!isDate) return tfDvNum(dv);
      const s = tvExcelSerial(dv);
      return s === null ? null : timed ? Math.floor(s + 1e-9) : Math.round(s);
    };
    const serials = /** @type {SlicerSource} */ (source).rows.map(r => valueOf(r[col]));
    let min = Infinity, max = -Infinity;
    serials.forEach(s => { if (s !== null) { if (s < min) min = s; if (s > max) max = s; } });
    if (min === Infinity) return isDate ? "has no dates in the data source" : "has no numbers in the data source";
    const lo = valueOf(loDv) ?? min, hi = valueOf(hiDv) ?? max;
    const above = opts.loStrict ? v => v > lo : v => v >= lo, below = opts.hiStrict ? v => v < hi : v => v <= hi;
    /** @type {any} */
    const field = { kind: "range", col, deriv: "none", isDate, numeric: !isDate, header: uniqueHeader(name), values: serials, items: [], selected: null,
                    tableCol: -1, lo, hi, min, max, includeNull: !!opts.includeNull, loStrict: !!opts.loStrict, hiStrict: !!opts.hiStrict,
                    calc: opts.calc || null, relative: opts.relative || null, control: null,
                    pass: r => serials[r] === null ? field.includeNull : above(serials[r]) && below(serials[r]) };
    columns.push({ header: field.header, values: serials, date: isDate });
    field.tableCol = columns.length + 1;
    fields.push(field);
    return field;
  };

  /** A range of an aggregate's card → a "marks" field (see markFields), or a reason it stays text.
   * @param {any} card @param {FieldRef} ref @param {any} f */
  const markRange = (card, ref, f) => {
    if (locked) return "visuals are already bound to the other slicers";
    const known = markFields.find(m => tfSameField(m.ref, ref));
    if (known) return known;
    if (!slicerMeasureExpr(/** @type {SlicerSource} */ (source), card.model, ref)) return `${ref.name}: not a sum, average, min, max or count of the rows, nor a calculation of those`;
    /** @type {any} */
    const field = { kind: "marks", ref, header: uniqueHeader(card.filterName || card.visualName || ref.name),
                    lo: tfIsNull(f.minValue) ? null : tfDvNum(f.minValue), hi: tfIsNull(f.maxValue) ? null : tfDvNum(f.maxValue), control: null };
    field.loRef = bound(field, "from");
    field.hiRef = bound(field, "to");
    markFields.push(field);
    console.log(`[Slicers] "${field.header}": the marks from ${field.lo} to ${field.hi} (Min / Max cells)`);
    return field;
  };

  /** a control field's condition on table row R (a dropdown's value, a range's From / To – or its bounds when a card
   * is not on the sheet) */
  const condOf = (f, R) => {
    const c = `${getExcelColumnName(f.tableCol)}${R}`;
    if (f.kind === "list") return `OR(${f.control.cell}="${SLICER_ALL}",${c}=${f.control.cell})`;
    const from = (f.control && f.control.from) || String(f.lo), to = (f.control && f.control.to) || String(f.hi);
    return `IF(ISNUMBER(${c}),AND(OR(${from}="",${c}${f.loStrict ? ">" : ">="}${from}),OR(${to}="",${c}${f.hiStrict ? "<" : "<="}${to})),` +
           `${f.includeNull ? "TRUE" : "FALSE"})`;
  };
  /** Visible on row R: SUBTOTAL(103) – 0 on a row the slicers hide – and 0 when one of these controls leaves the row out */
  const visibleFormula = (R, controls) => {
    const conds = controls.filter(f => f.kind === "list" ? f.control : true).map(f => condOf(f, R));
    return conds.length ? `IF(AND(${conds.join(",")}),SUBTOTAL(103,A${R}),0)` : `SUBTOTAL(103,A${R})`;
  };

  return {
    get active() { return (fields.length > 0 || markFields.length > 0) && !!source; },

    /** a dropdown's list of values on the lists sheet (the same values: the same list) → its range, for Data Validation
     * @param {any[]} values */
    listRef(values) {
      const key = JSON.stringify(values);
      let l = extraLists.find(x => x.key === key);
      if (!l) { l = { key, col: lists++, values }; extraLists.push(l); }
      const letter = getExcelColumnName(l.col);
      return sheetRef(SLICER_LISTS_SHEET, `$${letter}$1:$${letter}$${values.length}`);
    },
    get fields() { return fields; },
    /** a date made from a card's dropdowns: its formula in a cell of the hidden lists sheet → that cell's reference */
    dateCell(formula, result) {
      dateCells.push({ formula, result });
      return sheetRef(SLICER_LISTS_SHEET, `$${SLICER_DATES_COL}$${dateCells.length}`);
    },
    get failure() { return failed; },

    /**
     * A filter card → a field the visuals follow: a card of a field's values becomes a slicer – a dropdown cell when
     * Tableau's card picks one value (Single Value dropdown / list / slider) – and a range of dates becomes From / To
     * date cells. Returns the field, or a reason (string) the card stays text.
     * @param {ExportItem & { filter: any, field: FieldRef, sheetName: string, mode: string | null }} card @param {any} ws @param {FormatModel} model
     */
    async adoptCard(card, ws, model) {
      const f = card.filter, ref = card.field;
      const type = String((f && f.filterType) || "").toLowerCase().replace(/[^a-z]/g, "");
      const isRange = type === "range", isRelative = type === "relativedate";
      if (type !== "categorical" && !isRange && !isRelative) return `a ${f && f.filterType} filter cannot follow in Excel`;
      if (!ref || /^(measure names|:measure names)$/i.test(ref.name) || (ref.type === "qk" && type === "categorical")) return "not a filter of a dimension's values";
      let deriv = String(ref.deriv || "none").toLowerCase();
      if ((isRange || isRelative) && deriv === "tdy") deriv = "none";     // exact date
      // a range of an aggregate (SUM(Sales), a calculation of aggregates) keeps marks, not rows: Min / Max cells the
      // bound visuals' values test (markRange)
      const aggRange = isRange && (!!SLICER_AGGS[deriv] || deriv === "usr");
      if ((isRange || isRelative) && deriv !== "none" && !aggRange) return `a range of ${TF_DERIV_LABEL[deriv] || deriv}(${ref.name})`;
      if (!isRange && !isRelative && deriv !== "none" && !SLICER_DATE_PARTS.has(deriv)) return `a filter on ${TF_DERIV_LABEL[deriv] || deriv}(${ref.name})`;
      if (failed) return failed;
      if (!source) {
        try { source = await readSlicerSource(ws, f.fieldName); }
        catch (e) { failed = `the data source's rows could not be read: ${e.message}`; return failed; }
      } else {
        const ds = await slicerDataSourceOf(ws, f.fieldName);
        if (!ds || ds.id !== source.id) return `another data source than "${source.name}" (slicers filter one data source)`;
      }
      if (aggRange) return markRange(card, ref, f);
      const col = slicerColumnOf(source, model, ref);
      if (col < 0) return `"${ref.name}" is not a column of "${source.name}"`;
      const isDate = /date/.test(source.columns[col].dataType);
      const isNumber = /^(int|integer|float|real|number|double)/.test(source.columns[col].dataType);
      if (deriv !== "none" && !isDate) return `"${ref.name}" holds no dates`;
      if (isRelative && !isDate) return `a relative date filter of "${ref.name}", which holds no dates`;
      if (isRange && !isDate && !isNumber) return `a range of "${ref.name}", which holds neither dates nor numbers`;
      const ranged = isRange || isRelative;
      const known = fields.find(x => x.col === col && x.deriv === deriv && (x.kind === "range") === ranged && !x.calc);
      if (known) return known;
      if (locked) return "visuals are already bound to the other slicers";
      if (ranged) {
        // a relative date filter: the days it keeps at export ("Last 3 months" from today) as From / To dates
        const rel = isRelative ? slicerRelativeRange(f) : null;
        if (isRelative && !rel) return `"${ref.name}": a relative date filter of a period Excel cannot tell`;
        const iso = d => ({ value: d.toISOString().slice(0, 10) });
        const field = rangeField(col, card.filterName || card.visualName || source.columns[col].name,
                                 rel ? iso(rel.lo) : f.minValue, rel ? iso(rel.hi) : f.maxValue,
                                 { includeNull: !!f.includeNullValues, relative: rel ? rel.text : null });
        if (typeof field === "string") return `"${ref.name}" ${field}`;
        console.log(`[Slicers] "${field.header}": ${field.isDate ? "dates" : "numbers"} ${field.lo}–${field.hi} of ${field.min}–${field.max}` +
                    (rel ? ` (${rel.text} at export)` : field.isDate ? " (From / To cells)" : " (Min / Max cells)"));
        return field;
      }
      const values = source.rows.map(r => slicerRowText(r[col], deriv, isDate) ?? "");
      const items = [...new Set(values)].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
      let selected = null;
      if (!f.isAllSelected) {
        const applied = (f.appliedValues || []).map(dv => slicerMarkText(dv, deriv, isDate));
        if (applied.some(v => v === null || !items.includes(v))) return "its selected values are not values of the data source";
        selected = new Set(f.isExcludeMode ? items.filter(v => !applied.includes(v)) : applied);
        if (selected.size === items.length) selected = null;
      }
      // a card that picks one value – or a dropdown of several (checkdropdown) with all or one picked, when
      // FORMAT_CONFIG.dropdownCards – → a dropdown cell; else a slicer (a dropdown cell holds one value)
      const current = selected ? [...selected][0] : null;
      // a list (checklist) in a zone too short for a slicer's caption and one row of buttons: a dropdown as well
      const tooShort = !!(card.layout && card.layout.heightPx) && card.layout.heightPx < SLICER_HEADER_PX + SLICER_ITEM_PX + SLICER_PAD_PX;
      const oneValue = /^(dropdown|radiolist|slider)$/i.test(card.mode || "") ||
                       (FORMAT_CONFIG.dropdownCards && (/^checkdropdown$/i.test(card.mode || "") || tooShort));
      const single = oneValue && (!selected || selected.size === 1) && current !== "\u0000";
      /** @type {any} */
      const field = single
        ? { kind: "list", col, deriv, isDate, header: "", values, items, selected, tableCol: -1, current, control: null,
            listCol: lists++, pass: r => current === null || values[r] === current }
        : { kind: "slicer", col, deriv, isDate, header: "", values, items, selected, tableCol: -1,
            pass: r => !selected || selected.has(values[r]) };
      field.header = uniqueHeader(card.filterName || card.visualName || source.columns[col].name);
      columns.push({ header: field.header, values: values.map(v => (v === "\u0000" ? null : v)) });
      field.tableCol = columns.length + 1;
      if (single) {
        const choices = [SLICER_ALL, ...items.filter(v => v !== "\u0000")], letter = getExcelColumnName(field.listCol);
        field.choices = choices;
        field.listRef = sheetRef(SLICER_LISTS_SHEET, `$${letter}$1:$${letter}$${choices.length}`);
      }
      fields.push(field);
      console.log(`[Slicers] "${field.header}": ${items.length} values, ${selected ? selected.size + " selected" : "all selected"}` +
                  (single ? " (dropdown)" : ""));
      return field;
    },

    /**
     * A calculated filter keeping a date between two date parameters (slicerParamRange) → a "range" field whose
     * From / To are the parameters' cells on the dashboard. spec: { calc: the calculation's FieldRef, date: the date
     * field's name, lo / hi: the parameters' current DataValues, loStrict / hiStrict }. Returns the field or a reason.
     * @param {any} spec @param {any} ws a worksheet filtered by it @param {FormatModel} model
     */
    async adoptParamRange(spec, ws, model) {
      if (failed) return failed;
      if (!source) {
        try { source = await readSlicerSource(ws, spec.date); }
        catch (e) { failed = `the data source's rows could not be read: ${e.message}`; return failed; }
      } else {
        const ds = await slicerDataSourceOf(ws, spec.date);
        if (!ds || ds.id !== source.id) return `another data source than "${source.name}" (slicers filter one data source)`;
      }
      const col = slicerColumnOf(source, model, /** @type {any} */ ({ name: spec.date, ds: spec.calc.ds, deriv: null }));
      if (col < 0) return `"${spec.date}" is not a column of "${source.name}"`;
      if (!/date/.test(source.columns[col].dataType)) return `"${spec.date}" holds no dates`;
      const known = fields.find(x => x.calc && x.calc.name === spec.calc.name && (x.calc.ds || "") === (spec.calc.ds || ""));
      if (known) return known;
      if (locked) return "visuals are already bound to the other slicers";
      const field = rangeField(col, source.columns[col].name, spec.lo, spec.hi, { calc: spec.calc, loStrict: spec.loStrict, hiStrict: spec.hiStrict });
      if (typeof field === "string") return `"${spec.date}" ${field}`;
      console.log(`[Slicers] "${field.header}": dates ${field.lo}–${field.hi} of ${field.min}–${field.max} (date parameters)`);
      return field;
    },

    /**
     * Binds a visual to the slicers: its marks found in the source rows, each measure checked against Tableau's
     * numbers. Returns the binding (vm.slicerBinding) – or a reason (string) the visual keeps fixed values.
     * @param {ViewModel} vm @param {string} sheetName @param {any[]} filters the sheet's Tableau filters @param {FormatModel} model
     */
    bind(vm, sheetName, filters, model) {
      if (!source || (!fields.length && !markFields.length)) return "no slicer";
      const src = source;
      // Tableau filters this sheet by every slicer field (else the slicer would filter what Tableau does not)
      const resolved = (filters || []).map(f => {
        const { ref, deriv } = slicerFilterField(f.fieldName);
        const col = slicerColumnOf(src, model, ref);
        return { f, col, deriv, isDate: col >= 0 && /date/.test(src.columns[col].dataType) };
      });
      // which fields Tableau filters it by: a calculated date filter by the sheet's own filter list (its datasource:
      // two calculations can share a caption)
      const sheetFilters = ((model && model.sheets[sheetName]) || {}).filterRefs || [];
      const applies = fd => fd.calc ? sheetFilters.some(r => r.name === fd.calc.name && (r.ds || "") === (fd.calc.ds || ""))
        : resolved.some(x => x.col === fd.col && x.deriv === fd.deriv);
      // a slicer filters every visual (it hides the table's rows); a dropdown or date range only those it applies to
      const missing = fields.filter(fd => fd.kind === "slicer" && !applies(fd));
      if (missing.length) return `Tableau does not filter it by ${missing.map(m => m.header).join(", ")}`;
      const own = fields.filter(fd => fd.kind !== "slicer" && applies(fd));
      const shown = r => fields.every(fd => (fd.kind !== "slicer" && !own.includes(fd)) || fd.pass(r));
      // its dimensions → source columns: headers, and continuous dimensions such as a date axis (MONTH(Order Date));
      // the latitude / longitude Tableau generates follow the geographic fields
      const heads = [], colHeads = [];
      const generated = c => /^(latitude|longitude) \(generated\)$/i.test(c.name);
      /** a dimension field → its source column, or the reason it has none */
      const headOf = (ref, label) => {
        if (!ref) return `"${label}" is not a workbook field`;
        const deriv = String(ref.deriv || "none").toLowerCase();
        const col = slicerColumnOf(src, model, ref);
        if (col < 0) return `"${label}" is a calculation or not a column of "${src.name}"`;
        const isDate = /date/.test(src.columns[col].dataType);
        if (!(deriv === "none" || deriv === "attr" || (isDate && SLICER_DATE_PARTS.has(deriv)))) return `"${label}" (${deriv}) cannot be matched to the rows`;
        return { col, deriv: deriv === "attr" ? "none" : deriv, isDate };
      };
      // a dimension the rows do not have (a calculation, a constant such as "Total") is left out when the others
      // still tell every mark apart – it then adds no marks of its own (a colour or label calculation per mark)
      const dropped = [];
      for (let ci = 0; ci < vm.cols.length; ci++) {
        const c = vm.cols[ci];
        if (c.pivoted || generated(c)) continue;
        const deriv = String((c.ref && c.ref.deriv) || "none").toLowerCase();
        // a value: an aggregate, a calculation, or a percent of total of one (pcto:cnt:Gender – a measure, not a dimension)
        if (!c.isHeader && (!c.ref || SLICER_AGGS[deriv] || deriv === "usr" || /^pcto/.test(deriv))) continue;
        const h = headOf(c.ref, c.label || c.name);
        if (typeof h === "string") {
          if (!c.ref || slicerColumnOf(src, model, c.ref) < 0) { dropped.push({ ci, reason: h }); continue; }
          return h;
        }
        heads.push({ ci, ...h });
      }
      if (dropped.length) {
        const seen = new Map();
        for (const r of vm.rows) {
          const k = heads.map(h => slicerMarkText(r[h.ci], h.deriv, h.isDate)).join("\u0001");
          const d = dropped.map(x => tfDvText(r[x.ci])).join("\u0001");
          if (seen.has(k) && seen.get(k) !== d) return dropped.map(x => x.reason).join("; ");
          seen.set(k, d);
        }
      }
      // a matrix: the fields on Columns are in each value column's members (vm.cols[ci].matrixKey)
      // (a field the rows do not have – a colour calculation spread over columns – is left out: the check of every
      // value against Tableau's below tells whether it split marks)
      ((vm.matrix && vm.matrix.colFields) || []).forEach((ref, l) => {
        const h = headOf(ref, ref ? ref.name : "?");
        if (typeof h !== "string") colHeads.push({ ...h, l });
        else if (ref && slicerColumnOf(src, model, ref) >= 0) colHeads.push({ fail: h });
      });
      const colFail = colHeads.find(h => h.fail);
      if (colFail) return colFail.fail;
      // marks: one key per distinct combination of dimension values (a matrix cell: its row's and its column's)
      const ids = new Map();
      /** the mark of row r's value in column ci, or 0 when it cannot be told */
      const idOf = (r, ci) => {
        const t = heads.map(h => slicerMarkText(r[h.ci], h.deriv, h.isDate));
        if (colHeads.length) {
          const mk = vm.cols[ci] && vm.cols[ci].matrixKey;
          if (!mk) return 0;
          colHeads.forEach(h => t.push(slicerMarkText(mk[h.l], h.deriv, h.isDate)));
        }
        if (t.some(x => x === null)) return 0;
        const key = t.join("\u0001");
        if (!ids.has(key)) ids.set(key, ids.size + 1);
        return /** @type {number} */ (ids.get(key));
      };
      const valueCols = vm.cols.map((c, ci) => ci).filter(ci => !heads.some(h => h.ci === ci) && !dropped.some(x => x.ci === ci) && !generated(vm.cols[ci]));
      vm.rows.forEach(r => valueCols.forEach(ci => idOf(r, ci)));
      // the sheet's other filters on the rows (a filter of a measure or a calculation applies to marks, not rows)
      const tests = [];
      resolved.forEach(({ f, col, deriv, isDate }) => {
        if (col < 0 || fields.some(fd => fd.col === col && fd.deriv === deriv)) return;
        const type = String(f.filterType || "").toLowerCase();
        if (type === "categorical" && !f.isAllSelected) {
          const set = new Set((f.appliedValues || []).map(dv => slicerMarkText(dv, deriv, isDate)));
          tests.push(r => set.has(slicerRowText(r[col], deriv, isDate)) !== !!f.isExcludeMode);
        } else if (type === "range" && deriv === "none" && isDate) {
          const day = dv => { const s = tfIsNull(dv) ? null : tvExcelSerial(dv); return s === null ? null : Math.round(s); };
          const lo = f.minValue ? day(f.minValue) : null, hi = f.maxValue ? day(f.maxValue) : null;
          tests.push(r => { const v = day(r[col]); return v === null ? !!f.includeNullValues : (lo === null || v >= lo) && (hi === null || v <= hi); });
        }
      });
      // every source row → its mark (null: no mark of this visual, or filtered out by the sheet)
      const keys = src.rows.map(r => {
        if (!tests.every(t => t(r))) return null;
        return ids.get([...heads, ...colHeads].map(h => slicerRowText(r[h.col], h.deriv, h.isDate)).join("\u0001")) || null;
      });
      /** @type {Map<number, number[]>} the rows of each mark under Tableau's current selection */
      const rowsOf = new Map();
      keys.forEach((k, r) => { if (k && shown(r)) { if (!rowsOf.has(k)) rowsOf.set(k, []); /** @type {number[]} */ (rowsOf.get(k)).push(r); } });
      const allRows = [...rowsOf.values()].flat();             // every mark's (a percent of total's whole)

      // each measure: translated, then checked mark by mark against Tableau's value
      /** share: a percent of total within each stack (100 % stacked bars): the base measure's formula, which a 100 %
       * stacked Excel chart turns into the same shares – never a table cell or KPI value
       * @type {Map<number, { expr: SlicerExpr } | { reason: string, share?: SlicerExpr }>} */
      const measures = new Map();
      const near = (v, t) => v !== null && Math.abs(v - t) <= 1e-6 * Math.max(1, Math.abs(t));
      /** a percent of total Tableau computes within a partition (each month's bar, all the races in it): the heads
       * but one (the stacked dimension) – the base measure of each mark over its partition's sum gives every value */
      const shareWithin = (base, ci) => {
        const vals = vm.rows.map(r => ({ r, t: tfDvNum(r[ci]), b: slicerEval(base, rowsOf.get(idOf(r, ci)) || [], src, allRows) }));
        if (vals.some(x => x.t !== null && x.b === null)) return false;
        return heads.length > 1 && heads.some((_, skip) => {
          const keyOf = r => heads.filter((h, j) => j !== skip).map(h => tfDvText(r[h.ci])).join("\u0001");
          const sums = new Map();
          vals.forEach(x => { if (x.b !== null) sums.set(keyOf(x.r), (sums.get(keyOf(x.r)) || 0) + x.b); });
          return vals.every(x => x.t === null || near(sums.get(keyOf(x.r)) ? x.b / sums.get(keyOf(x.r)) : 0, x.t));
        });
      };
      valueCols.forEach(ci => {
        const c = vm.cols[ci], label = c.label || c.name;
        const expr = slicerMeasureExpr(src, model, c.ref);
        if (!expr) { measures.set(ci, { reason: `${label}: not a sum, average, min, max or (distinct) count of a data-source field, nor a calculation of those` }); return; }
        for (const r of vm.rows) {
          const t = tfDvNum(r[ci]);
          if (t === null) continue;
          const v = slicerEval(expr, rowsOf.get(idOf(r, ci)) || [], src, allRows);
          if (!near(v, t)) {
            const reason = `${label}: the data source's rows give ${v === null ? "no value" : +v.toPrecision(6)} where Tableau shows ${tfDvText(r[ci])}`;
            const base = expr.op === "pct" ? /** @type {SlicerExpr} */ (expr.a) : null;
            measures.set(ci, base && shareWithin(base, ci) ? { reason: `${label}: a percent of total within each stack (it follows only as a 100 % stacked chart)`, share: base } : { reason });
            return;
          }
        }
        measures.set(ci, { expr });
      });
      const ok = [...measures.values()].filter(m => "expr" in m || m.share).length;
      if (!ok) return [...measures.values()].map(m => /** @type {any} */ (m).reason)[0] || "no measure";
      locked = true;

      let keyCol = -1;
      // Visible – or, when a dropdown / date range does not apply to it, a Visible of its own (writeTable: c.vis)
      const controls = fields.filter(fd => fd.kind !== "slicer");
      let visCol = own.length === controls.length ? 1 : -1;
      const R = { col: c => range(fieldColumn(c)), get key() { return range(keyCol); },
                  get vis() {
                    if (visCol < 0) { columns.push({ header: uniqueHeader(`Visible: ${sheetName}`), values: src.rows.map((_, r) => (shown(r) ? 1 : 0)), vis: own }); visCol = columns.length + 1; }
                    return range(visCol);
                  } };
      /**
       * The formula of a value made of these marks' measure ci; null when it cannot follow the slicers.
       * "chart": a mark the slicers leave empty is #N/A (no point, no bar); "cell": an empty cell.
       */
      const formula = (rows, ci, mode, shares) => {
        const m = measures.get(ci);
        const f = m && !("expr" in m) && shares && m.share ? exprFormula(m.share, rows, ci, mode)
          : m && "expr" in m ? exprFormula(m.expr, rows, ci, mode) : null;
        return f && markTests.length ? keepMarks(f, rows, ci, mode) : f;
      };
      // ranges of an aggregate this sheet is filtered by (a marks field): a value shows only while its mark's
      // aggregate is between the card's Min and Max (either left empty: no bound) – else a gap / an empty cell
      const markTests = [], markReasons = [];
      markFields.filter(mf => sheetFilters.some(r => r.name === mf.ref.name && (r.ds || "") === (mf.ref.ds || ""))).forEach(mf => {
        const expr = slicerMeasureExpr(src, model, mf.ref);
        if (expr) markTests.push({ mf, expr });
        else markReasons.push(`${mf.header}: its measure cannot be computed from the rows (its Min / Max do not filter this visual)`);
      });
      const keepMarks = (f, rows, ci, mode) => {
        const conds = markTests.map(t => {
          const g = exprFormula(t.expr, rows, ci, "zero");
          return g && `AND(OR(${t.mf.loRef}="",${g}>=${t.mf.loRef}),OR(${t.mf.hiRef}="",${g}<=${t.mf.hiRef}))`;
        });
        if (conds.some(c => !c)) return f;
        return `IF(AND(${conds.join(",")}),${f},${mode === "chart" ? "NA()" : mode === "zero" ? "0" : '""'})`;
      };
      /** "zero": a mark the slicers leave empty counts 0 (a waterfall step) */
      const exprFormula = (expr, rows, ci, mode) => {
        if (!rows.length) return null;
        const ks = [...new Set(rows.map(r => idOf(r, ci)))];
        if (ks.some(k => !k)) return null;
        if (keyCol < 0) keyCol = addColumn(`Key: ${sheetName}`, keys);
        const empty = mode === "chart" ? "NA()" : mode === "zero" ? "0" : '""';
        if (ks.length > 1) {                               // several marks summed into one value: only sums and counts add up
          if (expr.op !== "agg" || !/^(SUM|COUNT)$/.test(expr.fn || "")) return null;
          const k = `{${ks.join(",")}}`;
          return `IF(SUM(COUNTIFS(${R.key},${k},${R.vis},1))=0,${empty},SUM(${slicerExcel(expr, R, k)}))`;
        }
        const k = String(ks[0]);
        const body = slicerExcel(expr, R, k);
        return `IF(COUNTIFS(${R.key},${k},${R.vis},1)=0,${empty},${slicerMayFail(expr) ? `IFERROR(${body},${empty})` : body})`;
      };
      /** a measure the visual does not show itself (a waterfall's step under its running total) – translated, and
       * checked against the values Tableau's own numbers give for it: per value its marks' rows (groups), its number */
      const extraExprs = new Map();
      const checkedExpr = (ref, groups, ci, expected) => {
        if (!extraExprs.has(ref.inner)) {
          const expr = slicerMeasureExpr(src, model, ref);
          const ok = !!expr && groups.every((rows, i) => {
            if (expected[i] === null) return true;
            const ks = [...new Set(rows.map(r => idOf(r, ci)))];
            if (!rows.length || ks.some(k => !k)) return false;
            const v = slicerEval(expr, ks.flatMap(k => rowsOf.get(k) || []), src, allRows);
            return v !== null && Math.abs(v - expected[i]) <= 1e-6 * Math.max(1, Math.abs(expected[i]));
          });
          extraExprs.set(ref.inner, ok ? expr : null);
        }
        return extraExprs.get(ref.inner);
      };
      // (a share counts as followed: where it cannot be – not a 100 % stacked chart – the report's "x of y values" says so)
      const reasons = [...measures.entries()].filter(([ci, m]) => !("expr" in m) && !m.share && vm.order.includes(ci)).map(([, m]) => /** @type {any} */ (m).reason)
        .concat(markReasons);
      console.log(`[Slicers] ${sheetName}: follows the slicers (${ids.size} marks, ${ok} measure${ok === 1 ? "" : "s"})` +
                  (reasons.length ? ` – fixed: ${reasons.join("; ")}` : ""));
      const binding = {
        sheetName, reasons,
        /** values written as formulas, and values that stayed fixed (what the report says of the visual) */
        written: 0, fixed: 0,
        /** a written number cell → its formula (Tableau's number stays as the cached result) */
        bindCell(cell, rows, ci) {
          if (typeof cell.value !== "number") return;
          const f = formula(rows, ci, "cell");
          if (f) { cell.value = { formula: f, result: cell.value }; binding.written++; } else binding.fixed++;
        },
        /** the formula of the value of these rows' measure ci (a pie's slice, a label's measure); null: it stays fixed */
        formulaFor: (rows, ci, mode) => formula(rows, ci, mode),
        /** formulas of a measure the visual does not show (ref) for each group of its rows (marks of column ci), when
         * it gives Tableau's numbers (expected); null otherwise. mode as formulaFor's ("zero": an empty step is 0) */
        measureFormulas(ref, groups, ci, expected, mode) {
          const expr = checkedExpr(ref, groups, ci, expected);
          if (!expr) return null;
          const out = groups.map(rows => exprFormula(expr, rows, ci, mode));
          return out.every(f => f) ? out : null;
        },
        /** a chart value whose array knows its marks (values.src: tvSum) → its formula. shares: a 100 % stacked chart
         * without value labels – a percent of total within each stack may be its base measure (Excel draws the shares) */
        valueFormula(values, i, shares) {
          const s = values && values.src;
          const f = s ? formula(s.rows[i], s.ci, "chart", shares) : null;
          if (f) binding.written++; else binding.fixed++;
          return f;
        }
      };
      return binding;
    },

    /** the hidden "Slicer Data" sheet: Row, Visible, the slicer fields, measure fields and keys; rows hidden as Tableau's
     * selection on the slicers leaves them. Visible is SUBTOTAL(103) – 0 on a row the slicers hide – and 0 too when a
     * dropdown cell or the From / To cells leave the row out. The dropdowns' lists go to a hidden sheet of their own.
     * @param {import("exceljs").Workbook} workbook */
    writeTable(workbook) {
      const ws = workbook.addWorksheet(SLICER_DATA_SHEET, { state: "hidden" });
      const head = ["Row", "Visible", ...columns.map(c => c.header)];
      head.forEach((h, c) => { ws.getCell(1, c + 1).value = h; });
      // the dropdown / From–To conditions (a dropdown no dashboard shows leaves its field out)
      const controls = fields.filter(f => f.kind !== "slicer");
      const dates = columns.map(c => !!c.date);
      for (let r = 0; r < N(); r++) {
        const row = ws.getRow(r + 2);
        row.getCell(1).value = r + 1;
        row.getCell(2).value = { formula: visibleFormula(r + 2, controls), result: visible(r) ? 1 : 0 };
        columns.forEach((c, k) => {
          if (c.values[r] === null || c.values[r] === undefined) return;
          const cell = row.getCell(k + 3);
          cell.value = c.vis ? { formula: visibleFormula(r + 2, c.vis), result: c.values[r] } : c.values[r];
          if (dates[k]) cell.numFmt = "yyyy-mm-dd";
        });
        if (!shownBySlicers(r)) row.hidden = true;
      }
      ws.columns.forEach((col, c) => { col.width = c < 2 ? 8 : 16; });
      const listed = fields.filter(f => f.kind === "list");
      if (listed.length || extraLists.length || dateCells.length || boundCells.length) {
        const ls = workbook.addWorksheet(SLICER_LISTS_SHEET, { state: "hidden" });
        listed.forEach(f => f.choices.forEach((v, i) => { const cell = ls.getCell(i + 1, f.listCol + 1); cell.value = v; cell.numFmt = "@"; }));
        extraLists.forEach(l => l.values.forEach((v, i) => { ls.getCell(i + 1, l.col + 1).value = v; }));
        dateCells.forEach((d, i) => { const cell = ls.getCell(`${SLICER_DATES_COL}${i + 1}`); cell.value = { formula: d.formula, result: d.result }; cell.numFmt = "m/d/yyyy"; });
        // a marks range's Min / Max as its card shows them (its bound when the card is not on any sheet)
        boundCells.forEach((b, i) => {
          const at = b.field.control && b.field.control[b.end], v = b.end === "from" ? b.field.lo : b.field.hi;
          ls.getCell(`${SLICER_BOUND_COL}${i + 1}`).value = at ? { formula: `IF(${at}="","",${at})`, result: v ?? "" } : (v ?? "");
        });
      }
      return { sheetName: ws.name, ref: `A1:${getExcelColumnName(head.length - 1)}${N() + 1}`, headers: head,
               visibleFormula: visibleFormula(2, controls), columnFormulas: columns.map(c => (c.vis ? visibleFormula(2, c.vis) : null)) };
    },

    /** the package parts' plan for injectSlicers @param {any[]} placed slicers placed on the dashboard sheets */
    packagePlan(table, placed) {
      const slicerFields = fields.filter(f => f.kind === "slicer");
      const caches = slicerFields.map(f => ({
        name: "Slicer_" + (f.header.replace(/[^A-Za-z0-9_]/g, "_").replace(/^(\d)/, "_$1") || "Field"),
        sourceName: f.header, column: table.headers.indexOf(f.header) + 1,
        filter: f.selected ? [...f.selected].map(v => (v === "\u0000" ? "" : v)) : null
      }));
      caches.forEach((c, i) => { let n = 2; while (caches.slice(0, i).some(x => x.name === c.name)) c.name = `${c.name.replace(/\d+$/, "")}${n++}`; });
      const taken = new Set();                       // a slicer's name is unique in the workbook: "Region", "Region 2" …
      return { table, caches, slicers: placed.map(p => {
        let name = p.caption, n = 2;
        while (taken.has(name.toLowerCase())) name = `${p.caption} ${n++}`;
        taken.add(name.toLowerCase());
        return { ...p, name, cache: caches[slicerFields.indexOf(p.field)].name };
      }) };
    },

    /** the slicer fields of a dashboard none of whose visuals could follow them: dropped again (nothing is bound yet) */
    discard() {
      if (locked) return;
      fields.length = 0;
      columns.length = 0;
      lists = 0;
      extraLists.length = 0;
      fieldCols.clear();
      headers.clear();
      headers.add("row").add("visible");
    },

    /** how big a card's slicer is drawn: its zone, and tall enough for its buttons (columns across a wide zone)
     * @param {{ items: string[] }} field @param {{ widthPx?: number, heightPx?: number }} layout */
    /** a slicer in its card's zone: as many columns of buttons as its width holds; the zone's height, the buttons
     * scrolling when there are more (the dashboard's layout stays Tableau's) – taller only when the zone cannot hold
     * the caption and one row of buttons */
    sizeOf(field, layout) {
      const w = layout.widthPx || 160, n = field.items.length;
      const columnCount = Math.max(1, Math.min(n, Math.floor(w / SLICER_ITEM_MIN_W_PX)));
      const rows = Math.min(SLICER_MAX_ITEM_ROWS, Math.ceil(n / columnCount));
      const least = SLICER_HEADER_PX + SLICER_ITEM_PX + SLICER_PAD_PX, full = SLICER_HEADER_PX + rows * SLICER_ITEM_PX + SLICER_PAD_PX;
      const zone = layout.heightPx || 0;
      return { columnCount, heightPx: zone >= least ? zone : Math.min(full, Math.max(zone, least)) };
    }
  };
}

/**
 * Binds a pie / donut to the slicers: its slices are the marks (their summary rows), its angle measure the value each
 * slice must rebuild from the data source's rows. Returns the binding, or why the pie keeps Tableau's values.
 * @param {ReturnType<typeof createSlicerModel>} slicers @param {any} pie @param {string} sheetName @param {any[]} filters @param {FormatModel} model
 */
function bindPie(slicers, pie, sheetName, filters, model) {
  if (pie.valueCol < 0) return "its slices have no size measure";
  const rows = pie.points.map(p => pie.data[p.row]);
  const vm = /** @type {any} */ ({
    // an aggregate is a measure even when Tableau gives it no number type (a calculation that is all null)
    cols: pie.columns.map(c => {
      const d = String((c.ref && c.ref.deriv) || "").toLowerCase();
      return { name: c.name, label: c.name, ref: c.ref, isHeader: !c.numeric && !(SLICER_AGGS[d] || d === "usr") };
    }),
    rows, order: [pie.valueCol], matrix: null
  });
  const binding = slicers.bind(vm, sheetName, filters, model);
  if (typeof binding === "string") return binding;
  if (rows.some(r => !binding.formulaFor([r], pie.valueCol, "chart")))
    return binding.reasons[0] || `${pie.valueCaption}: its slices cannot be rebuilt from the data source's rows`;
  return binding;
}

const DATE_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

/**
 * A date parameter's control, drawn like Tableau's: its title over one boxed row of Month / Day / Year dropdowns (the
 * card's zone split in three: item.splitCuts) – read as Tableau shows the date, 6/5/2022. The date they make is a
 * cell of the hidden lists sheet (slicers.dateCell): the From (or To) of each range a calculated filter keeps between
 * date parameters (adoptParamRange) – a day past the month's end is its last day. A date parameter no visual on the
 * sheet uses (item.dateOnly) gets the same dropdowns and a note saying so. Shown on another dashboard again, the card
 * shows the first one's date.
 * @param {import("exceljs").Worksheet} worksheet @param {any} item @param {any} tracker
 * @param {ReturnType<typeof createSlicerModel>} slicers (its lists sheet holds the dropdowns' lists and the date)
 */
function writeParamDateCard(worksheet, item, tracker, slicers) {
  const uses = item.paramFields || [];
  const first = uses[0];
  const serial = first ? (first.end === "from" ? first.field.lo : first.field.hi) : item.paramSerial;
  const date = new Date(Math.round((serial - 25569) * 86400000));
  const C = item.gridCol, W = Math.max(1, item.gridW), top = item.gridRow + 2;      // 1-based row of the dropdowns
  writeIndividualFilterTable(worksheet, item.filterName, [""], item.gridRow, C, tracker, W);
  const title = worksheet.getCell(item.gridRow + 1, C + 1);
  if (!uses.length) title.note = "No visual on this sheet uses this parameter: its date is here as Tableau shows it, " +
                                 "for reference – picking another one changes nothing on this sheet.";
  const lead = worksheet.getCell(top, C + 1);
  const font = lead.font;
  const abs = cell => { const m = String(cell.address).match(/^([A-Z]+)(\d+)$/); return sheetRef(worksheet.name, `$${m[1]}$${m[2]}`); };
  const owner = uses.map(pf => pf.field.control && pf.field.control[pf.end]).find(Boolean);
  if (owner) {                                         // shown again: the first card's date
    lead.value = { formula: owner, result: date };
    lead.numFmt = "m/d/yyyy";
    return;
  }
  const cuts = item.splitCuts && item.splitCuts.length === 2 ? item.splitCuts : W >= 3 ? [1, 2] : null;
  let ref;
  if (cuts) {
    const starts = [0, ...cuts], ends = [...cuts, W];
    if (W > 1) { try { worksheet.unMergeCells(top, C + 1, top, C + W); } catch (e) { /* not merged */ } }
    /** @type {Partial<import("exceljs").Border>} */
    const edge = { style: "thin", color: { argb: "FFBFBFBF" } }, inner = { style: "hair", color: { argb: "FFD9D9D9" } };
    // the years: those of the dates the ranges cover (and this one's)
    const ranges = uses.length ? uses.map(pf => pf.field) : slicers.fields.filter(f => f.kind === "range");
    const yearOf = s => new Date(Math.round((s - 25569) * 86400000)).getUTCFullYear();
    const y0 = ranges.length ? Math.min(...ranges.map(f => yearOf(f.min))) : date.getUTCFullYear() - 2;
    const y1 = ranges.length ? Math.max(...ranges.map(f => yearOf(f.max))) : date.getUTCFullYear() + 2;
    const years = [];
    for (let y = Math.min(y0, date.getUTCFullYear()); y <= Math.max(y1, date.getUTCFullYear()); y++) years.push(y);
    const parts = [
      { value: DATE_MONTHS[date.getUTCMonth()], list: DATE_MONTHS, name: "month" },
      { value: date.getUTCDate(), list: Array.from({ length: 31 }, (_, i) => i + 1), name: "day" },
      { value: date.getUTCFullYear(), list: years, name: "year" }
    ];
    const cells = parts.map((p, k) => {
      const c0 = C + starts[k], c1 = C + Math.max(starts[k] + 1, ends[k]) - 1;
      if (c1 > c0) { try { worksheet.mergeCells(top, c0 + 1, top, c1 + 1); } catch (e) { /* already merged */ } }
      const cell = worksheet.getCell(top, c0 + 1);
      cell.value = p.value;
      cell.font = font;
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };   // white on any dashboard background
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = { top: edge, bottom: edge, left: k ? inner : edge, right: k === parts.length - 1 ? edge : inner };
      cell.dataValidation = { type: "list", allowBlank: true, formulae: [slicers.listRef(p.list)], showErrorMessage: true,
                              errorTitle: item.filterName, error: `Pick a ${p.name} from the list.` };
      return abs(cell);
    });
    const [m, d, y] = cells, mi = `MATCH(${m},{${DATE_MONTHS.map(x => `"${x}"`).join(",")}},0)`;
    ref = slicers.dateCell(`IF(OR(${y}="",${m}="",${d}=""),"",DATE(${y},${mi},MIN(${d},DAY(EOMONTH(DATE(${y},${mi},1),0)))))`, date);
  } else {                                             // too narrow for three: the date typed in
    lead.value = date;
    lead.numFmt = "m/d/yyyy";
    lead.dataValidation = { type: "date", operator: "greaterThan", allowBlank: true, formulae: [new Date(Date.UTC(1900, 0, 1))],
                            showErrorMessage: true, errorTitle: item.filterName, error: "A date, e.g. 1/1/2024." };
    ref = abs(lead);
  }
  uses.forEach(pf => { pf.field.control = pf.field.control || {}; pf.field.control[pf.end] = ref; });
}

/** what a filter card became, for the Conversion Report @param {any} field */
function controlOutput(field) {
  if (field.kind === "list") return "Dropdown cell (its value as in Tableau)";
  if (field.kind === "marks") return "Min / Max cells: the visuals show the marks whose value is between them";
  if (field.relative) return `From / To date cells: ${field.relative} at export`;
  return field.numeric ? "Min / Max cells (the range as in Tableau)" : "From / To date cells (the range as in Tableau)";
}

/**
 * A filter card the visuals follow through cells, drawn as Tableau's card: its title, then a dropdown cell (a list of
 * (All) and the field's values) or From / To date cells (dates within the data's). The first card of a field owns
 * the cells the Visible formulas read; another card of it (another dashboard) shows them.
 * @param {import("exceljs").Worksheet} worksheet @param {any} item @param {any} tracker
 */
function writeControlCard(worksheet, item, tracker) {
  const field = item.control.field, list = field.kind === "list";
  const numeric = !list && (field.kind === "marks" || field.numeric);              // Min / Max, not From / To dates
  const dateOf = serial => new Date(Math.round((serial - 25569) * 86400000));
  const values = list ? [field.current === null ? SLICER_ALL : field.current]
    : numeric ? [field.lo ?? "", field.hi ?? ""] : [dateOf(field.lo), dateOf(field.hi)];
  writeIndividualFilterTable(worksheet, item.filterName || field.header, values, item.gridRow, item.gridCol, tracker, item.gridW);
  const cells = values.map((v, i) => worksheet.getCell(item.gridRow + 2 + i, item.gridCol + 1));
  const ref = cell => { const m = String(cell.address).match(/^([A-Z]+)(\d+)$/); return sheetRef(worksheet.name, `$${m[1]}$${m[2]}`); };
  if (!list) cells.forEach((c, i) => { c.numFmt = numeric ? (i ? '"Max   "General;"Max   "-General' : '"Min   "General;"Min   "-General') : i ? '"To   "d mmm yyyy' : '"From "d mmm yyyy'; });
  // a relative date filter: the window it kept at export, and that the cells hold that window, not "the last 3 months"
  if (field.relative) worksheet.getCell(item.gridRow + 1, item.gridCol + 1).note =
    `Tableau: ${field.relative}. Excel keeps the dates it covered at export – change From / To for another window.`;
  const owner = field.control;
  if (owner) {                                         // shown again: the cells of the card that owns the field
    const refs = list ? [owner.cell] : [owner.from, owner.to];
    cells.forEach((c, i) => { c.value = { formula: refs[i], result: values[i] }; });
    return;
  }
  if (list) {
    const c = cells[0];
    c.numFmt = "@";                                    // a value picked stays text ("2024", not the number)
    c.dataValidation = { type: "list", allowBlank: false, formulae: [field.listRef], showErrorMessage: true,
                         errorTitle: field.header, error: "Pick a value from the list." };
    field.control = { cell: ref(c) };
  } else if (numeric) {
    // a number; empty: no bound on that side ("at least" / "at most")
    cells.forEach(c => {
      c.dataValidation = field.kind === "marks"
        ? { type: "decimal", operator: "between", allowBlank: true, formulae: [-1e15, 1e15], showErrorMessage: true, errorTitle: field.header, error: "A number." }
        : { type: "decimal", operator: "between", allowBlank: true, formulae: [field.min, field.max], showErrorMessage: true,
            errorTitle: field.header, error: `A number from ${field.min} to ${field.max}.` };
    });
    field.control = { from: ref(cells[0]), to: ref(cells[1]) };
  } else {
    const fmt = s => dateOf(s).toISOString().slice(0, 10);
    cells.forEach(c => {
      c.dataValidation = { type: "date", operator: "between", allowBlank: true, formulae: [dateOf(field.min), dateOf(field.max)],
                           showErrorMessage: true, errorTitle: field.header, error: `A date from ${fmt(field.min)} to ${fmt(field.max)}.` };
    });
    field.control = { from: ref(cells[0]), to: ref(cells[1]) };
  }
}

/**
 * Adds the slicer table, the slicer caches and the slicers (drawn where the filter cards were) to a written XLSX –
 * the parts Excel itself writes for a table slicer.
 * @param {ArrayBuffer | Uint8Array} buffer
 * @param {{ table: { sheetName: string, ref: string, headers: string[] }, caches: { name: string, sourceName: string, column: number, filter: string[] | null }[],
 *           slicers: { sheetName: string, name: string, caption: string, cache: string, columnCount: number,
 *                      from: { col: number, colOff: number, row: number, rowOff: number }, to: { col: number, colOff: number, row: number, rowOff: number } }[] }} plan
 */
async function injectSlicers(buffer, plan) {
  const zip = await JSZip.loadAsync(buffer);
  const read = p => zip.file(p) ? zip.file(p).async("string") : Promise.resolve(null);
  const NS_X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
  const NS_X15 = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
  const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const REL_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
  const REL_SLICER = "http://schemas.microsoft.com/office/2007/relationships/slicer";
  const REL_SLICER_CACHE = "http://schemas.microsoft.com/office/2007/relationships/slicerCache";
  let wbXml = await read("xl/workbook.xml");
  let wbRels = await read("xl/_rels/workbook.xml.rels");
  let ct = await read("[Content_Types].xml");
  const unescapeXml = v => String(v || "").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
  const sheetPathOf = name => {
    const tag = (wbXml.match(/<sheet\b[^>]*>/g) || []).find(t => unescapeXml(attr(t, "name")) === name);
    if (!tag) throw new Error(`sheet "${name}" not found in the workbook package`);
    return resolve("xl/", relTarget(wbRels, attr(tag, "r:id")));
  };
  /** an <ext> into the part's <extLst> (created before the closing tag when missing) */
  const addExt = (xml, closing, ext) => xml.includes("</extLst>") ? xml.replace(/<\/extLst>(?![\s\S]*<\/extLst>)/, ext + "</extLst>")
    : xml.replace(closing, `<extLst>${ext}</extLst>${closing}`);

  // ── the table over the slicer data, filtered as Tableau's selection is ──
  const { table } = plan;
  const dataPath = sheetPathOf(table.sheetName);
  let dataXml = await read(dataPath);
  let dataRels = (await read(relsPathOf(dataPath))) || EMPTY_RELS;
  const tableN = nextFreeIndex(zip, "xl/tables/table", ".xml");
  const ids = [];
  for (const p of Object.keys(zip.files).filter(n => /^xl\/tables\/table\d+\.xml$/.test(n))) ids.push(+attr(((await read(p)) || "").match(/<table\b[^>]*>/)[0], "id") || 0);
  const tableId = Math.max(0, ...ids) + 1;
  const filterCols = plan.caches.filter(c => c.filter).map(c => `<filterColumn colId="${c.column - 1}"><filters${c.filter.includes("") ? ' blank="1"' : ""}>` +
    c.filter.filter(v => v !== "").map(v => `<filter val="${esc(v)}"/>`).join("") + `</filters></filterColumn>`).join("");
  zip.file(`xl/tables/table${tableN}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<table xmlns="${NS_MAIN}" id="${tableId}" name="${SLICER_TABLE}" displayName="${SLICER_TABLE}" ref="${table.ref}" totalsRowShown="0">` +
    `<autoFilter ref="${table.ref}">${filterCols}</autoFilter><tableColumns count="${table.headers.length}">` +
    table.headers.map((h, i) => {
      const f = i === 1 ? table.visibleFormula || "SUBTOTAL(103,A2)" : i > 1 && table.columnFormulas ? table.columnFormulas[i - 2] : null;
      return `<tableColumn id="${i + 1}" name="${esc(h)}"${f ? `><calculatedColumnFormula>${esc(f)}</calculatedColumnFormula></tableColumn>` : "/>"}`;
    }).join("") +
    `</tableColumns><tableStyleInfo name="TableStyleLight1" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`);
  ct = addOverride(ct, `xl/tables/table${tableN}.xml`, "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml");
  const tableRid = nextRelId(dataRels);
  dataRels = addRel(dataRels, tableRid, REL_TABLE, `../tables/table${tableN}.xml`);
  if (!/<worksheet\b[^>]*\sxmlns:r=/.test(dataXml)) dataXml = dataXml.replace(/<worksheet\b/, `<worksheet xmlns:r="${NS.r}"`);
  const tablePart = `<tableParts count="1"><tablePart r:id="${tableRid}"/></tableParts>`;
  dataXml = dataXml.includes("<extLst") ? dataXml.replace("<extLst", tablePart + "<extLst") : dataXml.replace("</worksheet>", tablePart + "</worksheet>");
  zip.file(dataPath, dataXml);
  zip.file(relsPathOf(dataPath), dataRels);

  // ── a slicer cache per field (workbook part, named in definedNames and the workbook's slicer list) ──
  const cacheRids = [];
  plan.caches.forEach(c => {
    const n = nextFreeIndex(zip, "xl/slicerCaches/slicerCache", ".xml");
    zip.file(`xl/slicerCaches/slicerCache${n}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<slicerCacheDefinition xmlns="${NS_X14}" xmlns:mc="${NS.mc}" mc:Ignorable="x" xmlns:x="${NS_MAIN}" name="${esc(c.name)}" sourceName="${esc(c.sourceName)}">` +
      `<extLst><x:ext uri="{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}" xmlns:x15="${NS_X15}"><x15:tableSlicerCache tableId="${tableId}" column="${c.column}"/></x:ext></extLst>` +
      `</slicerCacheDefinition>`);
    ct = addOverride(ct, `xl/slicerCaches/slicerCache${n}.xml`, "application/vnd.ms-excel.slicerCache+xml");
    const rid = nextRelId(wbRels);
    wbRels = addRel(wbRels, rid, REL_SLICER_CACHE, `slicerCaches/slicerCache${n}.xml`);
    cacheRids.push(rid);
  });
  if (!plan.caches.length) {                     // only dropdown / date cells: the table alone
    zip.file("[Content_Types].xml", ct);
    return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  }
  const names = plan.caches.map(c => `<definedName name="${esc(c.name)}">#N/A</definedName>`).join("");
  wbXml = wbXml.includes("</definedNames>") ? wbXml.replace("</definedNames>", names + "</definedNames>")
    : wbXml.includes("<definedNames/>") ? wbXml.replace("<definedNames/>", `<definedNames>${names}</definedNames>`)
    : wbXml.replace("</sheets>", `</sheets><definedNames>${names}</definedNames>`);
  wbXml = addExt(wbXml, "</workbook>", `<ext uri="{46BE6895-7355-4a93-B00E-2C351335B9C9}" xmlns:x15="${NS_X15}">` +
    `<x15:slicerCaches xmlns:x14="${NS_X14}">${cacheRids.map(r => `<x14:slicerCache r:id="${r}"/>`).join("")}</x15:slicerCaches></ext>`);
  zip.file("xl/workbook.xml", wbXml);
  zip.file("xl/_rels/workbook.xml.rels", wbRels);

  // ── the slicers of each dashboard sheet: a slicers part, and each slicer drawn where its filter card was ──
  for (const sheetName of [...new Set(plan.slicers.map(s => s.sheetName))]) {
    const own = plan.slicers.filter(s => s.sheetName === sheetName);
    const sheetPath = sheetPathOf(sheetName);
    let sheetXml = await read(sheetPath);
    let sheetRels = (await read(relsPathOf(sheetPath))) || EMPTY_RELS;
    if (!/<worksheet\b[^>]*\sxmlns:r=/.test(sheetXml)) sheetXml = sheetXml.replace(/<worksheet\b/, `<worksheet xmlns:r="${NS.r}"`);
    const n = nextFreeIndex(zip, "xl/slicers/slicer", ".xml");
    zip.file(`xl/slicers/slicer${n}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<slicers xmlns="${NS_X14}" xmlns:mc="${NS.mc}" mc:Ignorable="x" xmlns:x="${NS_MAIN}">` +
      own.map(s => `<slicer name="${esc(s.name)}" cache="${esc(s.cache)}" caption="${esc(s.caption)}" columnCount="${s.columnCount}" rowHeight="${SLICER_ITEM_PX * EMU_PER_PX}"/>`).join("") +
      `</slicers>`);
    ct = addOverride(ct, `xl/slicers/slicer${n}.xml`, "application/vnd.ms-excel.slicer+xml");
    const slicerRid = nextRelId(sheetRels);
    sheetRels = addRel(sheetRels, slicerRid, REL_SLICER, `../slicers/slicer${n}.xml`);
    sheetXml = addExt(sheetXml, "</worksheet>", `<ext uri="{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}" xmlns:x15="${NS_X15}">` +
      `<x14:slicerList xmlns:x14="${NS_X14}"><x14:slicer r:id="${slicerRid}"/></x14:slicerList></ext>`);

    // the sheet's drawing (charts and pictures are in it already), else a new one
    let drawingPath, drawingXml;
    const existingRid = (sheetXml.match(/<drawing\b[^>]*r:id="([^"]+)"/) || [])[1];
    if (existingRid) {
      drawingPath = resolve(dirOf(sheetPath), relTarget(sheetRels, existingRid));
      drawingXml = await read(drawingPath);
    } else {
      const d = nextFreeIndex(zip, "xl/drawings/drawing", ".xml");
      drawingPath = `xl/drawings/drawing${d}.xml`;
      drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}"></xdr:wsDr>`;
      const rid = nextRelId(sheetRels);
      sheetRels = addRel(sheetRels, rid, REL_DRAWING, `../drawings/drawing${d}.xml`);
      sheetXml = insertDrawingTag(sheetXml, rid);
      ct = addOverride(ct, drawingPath, CT_DRAWING);
    }
    if (!/<xdr:wsDr\b[^>]*\sxmlns:a=/.test(drawingXml)) drawingXml = drawingXml.replace(/<xdr:wsDr\b/, `<xdr:wsDr xmlns:a="${NS.a}"`);
    const used = [...drawingXml.matchAll(/<xdr:cNvPr\b[^>]*\sid="(\d+)"/g)].map(m => +m[1]);
    let shapeId = (used.length ? Math.max(...used) : 1) + 1;
    const marker = (tag, p) => `<xdr:${tag}><xdr:col>${p.col}</xdr:col><xdr:colOff>${p.colOff}</xdr:colOff><xdr:row>${p.row}</xdr:row><xdr:rowOff>${p.rowOff}</xdr:rowOff></xdr:${tag}>`;
    const anchors = own.map(s => `<xdr:twoCellAnchor editAs="oneCell">${marker("from", s.from)}${marker("to", s.to)}` +
      `<mc:AlternateContent xmlns:mc="${NS.mc}"><mc:Choice xmlns:sle15="http://schemas.microsoft.com/office/drawing/2012/slicer" Requires="sle15">` +
      `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${shapeId++}" name="${esc(s.name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
      `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/slicer">` +
      `<sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="${esc(s.name)}"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice>` +
      `<mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr>` +
      `<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val="white"/></a:solidFill></xdr:spPr>` +
      `<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1100"/>` +
      `<a:t>Slicer "${esc(s.caption)}": table slicers need Excel 2013 or later.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback>` +
      `</mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>`).join("");
    drawingXml = drawingXml.replace("</xdr:wsDr>", anchors + "</xdr:wsDr>");     // over the charts: a slicer is clicked
    zip.file(drawingPath, drawingXml);
    zip.file(sheetPath, sheetXml);
    zip.file(relsPathOf(sheetPath), sheetRels);
  }
  zip.file("[Content_Types].xml", ct);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}


/* ═══ export/export.js ════════════════════════════════════════════════════════════════════════════ */
/* The export: Tableau dashboard → formatted XLSX with native charts. */
const makeExportFileName = exportFileName;

/* the longest a Tableau-rendered picture may take (it can stall once embedded dashboards were opened) */
const IMAGE_RENDER_TIMEOUT_MS = 30000;

/* =========================================================================
 * exportToExcel() — one click → ONE .xlsx, ONE sheet per dashboard.
 *   Tableau Cloud with "All dashboards" ticked: every visible dashboard in tab order (the others are read
 *   through hidden embedded views – cloud/dashboards.js); otherwise only the dashboard the extension sits on.
 *   Each dashboard is written by writeDashboardSheet() (tables, KPI cards, text, images, native charts);
 *   the charts are injected sheet by sheet after ExcelJS has written the file; one Conversion Report covers all.
 * ========================================================================= */
async function exportToExcel() {
  const dashboard = tableau.extensions.dashboardContent.dashboard;
  const sheets = dashboard.worksheets;
  const btn = /** @type {HTMLButtonElement} */ (document.getElementById("export_button"));
  const btnText = btn.textContent;
  btn.disabled = true;
  /** @type {string[]} dashboards left out, and why */
  let notes = [];

  try {
    if (!(await ensureFormatModel()) && !window.confirm(
        "No Tableau workbook is loaded for this dashboard.\n\n" +
        "Without it the extension cannot tell which sheets are charts, so they will be exported as tables.\n\n" +
        "Cancel, then click 📁 Load Workbook to export real charts – or OK to export tables anyway.")) {
      setExportStatus("Export stopped – click 📁 Load Workbook first so charts are exported as charts", true);
      return;
    }
    const wantAll = wantsAllDashboards();
    // "<workbook> - <dashboard>.xlsx" for one dashboard, "<workbook>.xlsx" when every dashboard can be read
    // (the workbook auto-loaded from Tableau Cloud – a manual file exports only this dashboard)
    const wbFile = formatModelFileName();
    const allName = wantAll && wbFile && loadedWorkbook().id ? String(wbFile).replace(/\.twbx?$/i, "") : null;
    const exportFileName = makeExportFileName(wbFile, allName || dashboard.name);
    // a new export replaces one not saved yet
    pendingExport = null;
    if (document.getElementById("save_button")) saveButton().hidden = true;

    const tExport = performance.now();
    let tReadMs = 0;
    const titleMap = getTitleMap();
    console.log(`[Export] Using ${Object.keys(titleMap).length} titles`);
    const fmtModel = await ensureFormatModel();
    console.log(`[Export] Format model: ${fmtModel ? Object.keys(fmtModel.sheets).length + " sheets" : "none – load the workbook for exact formatting"}`);

    // ── which dashboards: the current one, or every visible one (Tableau Cloud) ──
    const current = { dashboard, sheets };
    let targets = [current];
    if (wantAll) {
      const collected = await collectAllDashboards(current, (name, i, n, stage) => {
        if (stage === "signin") {
          btn.textContent = "🔐 Waiting for Tableau sign-in…";
          hideExportOverlay();                       // the sign-in panel needs to be seen
          return;
        }
        btn.textContent = `⏳ Reading dashboards (${i}/${n} done)…`;
        showExportOverlay(`Reading dashboards — ${i} of ${n} done`, `Opening them in the background (up to ${PARALLEL_DASHBOARDS} at a time) to read their tables and charts.`);
      });
      targets = collected.targets;
      notes = collected.notes || [];
      tReadMs = performance.now() - tExport;
    }
    btn.textContent = "⏳ Building Excel…";
    if (targets.length > 1) showExportOverlay("Building Excel…", `${targets.length} dashboards, one sheet each.`);

    // ── one sheet per dashboard ──
    const workbook = new ExcelJS.Workbook();
    // the helper sheets' names are taken: a dashboard called "Chart Data" becomes "Chart Data (2)"
    const usedSheetNames = new Set([CHART_DATA_SHEET, DATA_SHEET, REPORT_SHEET, SLICER_DATA_SHEET, SLICER_LISTS_SHEET].map(s => s.toLowerCase()));
    const chartData = { sheet: null, row: 0 };                  // one hidden "Chart Data" sheet for every dashboard
    const slicers = FORMAT_CONFIG.slicers && fmtModel ? createSlicerModel() : null;   // one slicer table for every dashboard
    const written = [];
    for (const target of targets) {
      const sheetName = uniqueExcelSheetName(target.dashboard.name, usedSheetNames);
      const result = await writeDashboardSheet(workbook, sheetName, target, { titleMap, fmtModel, chartData, slicers });
      if (result) written.push(result);
      else {
        usedSheetNames.delete(sheetName.toLowerCase());
        console.log(`[Export] "${target.dashboard.name}" has no exportable data – no sheet added`);
        if (targets.length > 1) notes.push(`"${target.dashboard.name}": no data in any visible worksheet — no sheet added`);
      }
    }
    if (!written.length) throw new Error("No data found in any visible worksheet.");
    const tSheetsDone = performance.now();
    // the slicers' data source rows: a hidden table every bound formula reads; Excel recalculates them on opening
    const slicerTable = slicers && slicers.active ? slicers.writeTable(workbook) : null;
    if (slicerTable) workbook.calcProperties = { ...(workbook.calcProperties || {}), fullCalcOnLoad: true };

    if (FORMAT_CONFIG.conversionReport) writeConversionReports(workbook, written.map(w => w.report));

    setExportStatus("Building the Excel file…");
    dropEmptyRichTextRuns(workbook);
    /** @type {any} the XLSX bytes: ExcelJS's buffer, then injectCharts' Uint8Array */
    let buffer = await workbook.xlsx.writeBuffer();
    // fit-to-page next to collapsed row groups: ExcelJS writes <sheetPr> out of order – put it right
    // and row groups in Excel's own form (else Excel asks to repair the file)
    if (written.some(w => w.worksheet.properties.outlineProperties)) buffer = await fixSheetProperties(buffer);
    // native charts, sheet by sheet: sized to their blocks now that the final column widths / row heights are known
    for (const w of written) {
      const jobs = w.finalizeCharts();
      if (!jobs.length) continue;
      try {
        buffer = await ExcelChartWriter.injectCharts(buffer, { sheetName: w.worksheet.name, charts: jobs });
        console.log(`[Charts] ${w.worksheet.name}: ${jobs.length} native Excel chart(s) added`);
      } catch (err) {
        console.error(`[Charts] ${w.worksheet.name}: could not add native charts`, err);
        jobs.forEach(job => {
          const status = w.visualStatuses.find(entry => entry.worksheet === job.item.name);
          if (status) { status.status = "warning"; status.reason = `chart injection failed: ${err.message}`; }
        });
      }
    }
    // the slicers where the filter cards were (after the charts: they share the sheet's drawing, slicers on top)
    if (slicerTable) {
      try {
        const placed = written.flatMap(w => w.finalizeSlicers());
        buffer = await injectSlicers(buffer, slicers.packagePlan(slicerTable, placed));
        console.log(`[Slicers] ${placed.length} slicer(s) over ${slicerTable.ref} of "${SLICER_DATA_SHEET}"`);
      } catch (err) {
        console.error("[Slicers] could not add the slicers", err);
        notes.push(`slicers could not be added (${err.message}) – the filtered values stay as Tableau shows them`);
      }
    }

    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const waitsForSave = offerSave(blob, exportFileName);           // saved by its own click, once complete
    // the panel's summary: every visual of every dashboard ("Dashboard / Sheet" when there are several), plus
    // where the file went and what was left out
    const multi = written.length > 1;
    const visualStatuses = written.flatMap(w => multi
      ? w.visualStatuses.map(s => ({ ...s, worksheet: `${w.dashboardName} / ${s.worksheet}` })) : w.visualStatuses);
    const workbookWarning = [...new Set(written.map(w => w.workbookWarning).filter(Boolean))].join(" · ");
    updateVisualStatus(visualStatuses, workbookWarning);
    appendExportStatus((multi ? `${written.length} dashboards, one sheet each – ` : "") +
      (waitsForSave ? "the Excel file is ready – click 💾 Save Excel file to save it" : `downloaded as ${exportFileName}`));
    const timing = `⏱ ${tReadMs ? `dashboards ${secs(tReadMs)} · ` : ""}sheets ${secs(tSheetsDone - tExport - tReadMs)} · ` +
      `file ${secs(performance.now() - tSheetsDone)} · total ${secs(performance.now() - tExport)}`;
    console.log(`[Timing] ${timing}`);
    if (FORMAT_CONFIG.showTiming) appendExportStatus(timing);
    appendExportNotes(notes);

    console.log(`✅ Export completed with Tableau formatting and native charts — ${written.length} sheet(s): ${written.map(w => w.worksheet.name).join(", ")}`);

  } catch (err) {
    console.error("[Export]", err);
    setExportStatus("Export failed: " + err.message, true);
    appendExportNotes(notes);
    alert("Export failed. Check console (F12) for details.\n\n" + err.message);
  } finally {
    removeHiddenVizHost();
    hideExportOverlay();
    btn.textContent = btnText;
    btn.disabled = false;
  }
}

/* "All dashboards" ticked (default) on Tableau Cloud / Server, with the workbook loaded and more than one
 * dashboard in it → every visible dashboard. Unticked or Tableau Desktop: only this one. (A manual file has no
 * Tableau Cloud id: collectAllDashboards then exports this one and says why.) */
function wantsAllDashboards() {
  const allBox = /** @type {HTMLInputElement} */ (document.getElementById("export_all"));
  const env = tableau.extensions.environment || {};
  if ((allBox && !allBox.checked) || env.context === "desktop" || liteMode()) return false;
  const wb = loadedWorkbook();
  return !!(wb && wb.xml && dashboardTabOrder(wb.xml).length > 1);
}

/* =========================================================================
 * writeDashboardSheet() — ONE dashboard onto ONE new Excel sheet: tables, KPI cards, text boxes, images,
 * filters / parameters and native charts at their dashboard positions.
 * target: { dashboard: { name, objects, worksheets, getParametersAsync }, sheets: Tableau worksheets (live, or
 * snapshots of another dashboard) }; ctx: { titleMap, fmtModel, chartData: the shared hidden data sheet }.
 * Returns null (and adds no sheet) when there is nothing to export, else
 * { worksheet, dashboardName, visualStatuses, workbookWarning, report, finalizeCharts() → ChartJob[] }
 * (finalizeCharts runs after workbook.xlsx.writeBuffer(), with the sheet's final column widths / row heights).
 * ========================================================================= */
async function writeDashboardSheet(workbook, sheetName, target, ctx) {
    const sheets = target.sheets;
    const { titleMap, fmtModel, chartData } = ctx;
    // the views its buttons switch to are drawn below it (their objects shown and moved there)
    const { dashboard, alternates } = withAlternateViews(target.dashboard, fmtModel);
    /** a zone hidden on the dashboard – not one of those views */
    const zoneHidden = z => !!z.hidden && !alternates.has(String(z.id));
    const filterValuesMap = await extractFilterValuesPerField(sheets);
    console.log(`📊 [${dashboard.name}] Filter values per field:`, filterValuesMap);
    const workbookWarning = fmtModel && isFormatModelStale(fmtModel)
      ? "⚠ The workbook was loaded with an older version of the extension – click 📁 Load Workbook once more for the latest formatting"
      : fmtModel ? describeWorkbookMatch(checkWorkbookMatch(fmtModel, dashboard))
      : "⚠ Workbook not loaded – charts can only be recognised from the workbook, so they may be exported as tables. Click 📁 Load Workbook, then export again";
    if (workbookWarning) console.warn("[Export]", workbookWarning);

    const layoutMap = buildLayoutMap(dashboard.objects || [], titleMap);

    // ── 1. Build DZV map ── (a sheet squeezed into a 1 px zone is a common way to hide it: not shown either)
    const dzvMap = {};
    (dashboard.objects || [])
      .filter(obj => obj.type === "worksheet")
      .forEach(obj => {
        const tiny = obj.size && (obj.size.width < 8 || obj.size.height < 8);
        dzvMap[obj.name] = obj.isVisible !== false && !tiny;
      });
    console.log("[DZV] Visibility map:", dzvMap);

    // ── 2. Fetch all sheets in parallel ──
    setExportStatus(`Reading ${sheets.length} worksheets of "${dashboard.name}" from Tableau…`);
    const allSheetsData = await fetchAllSheetsData(sheets);

    /** @type {{ name: string, shape: string, layout: any }[]} button / icon sheets, drawn as their icon */
    const iconSheets = [];
    /** @type {ExportItem[]} */
    const filterValueItems = [];
    /** @type {ExportItem[]} */
    const dataWorksheetItems = [];
    const visualStatuses = [];
    /** sheets drawn inside a donut's hole (its centre text): not exported as blocks of their own */
    const pieCentreSheets = new Set();
    /** @type {ExportItem[]} colour legends of pies, drawn at their dashboard position */
    const pieLegendItems = [];
    /** @type {{ name: string, reason: string }[]} worksheets not exported, and why (conversion report) */
    const skipped = [];
    /* A pie's colour legends shown on the dashboard (the chart has none of its own): the TWB legend zones of the
       sheet, matched to each ring by the field they show, laid out at their dashboard position. TWB zones use
       0–100000 units: they are mapped to API pixels through the sheet's own zone. */
    /** @returns {ExportItem[]} */
    const pieLegends = (pie, sheetName, vm) => {
      const dashM = fmtModel && fmtModel.dashboards && fmtModel.dashboards[dashboard.name];
      const sheetZone = dashM && dashM.zones.find(z => z.type === "worksheet" && z.name === sheetName && z.w && z.h);
      const api = (dashboard.objects || []).find(o => o.type === "worksheet" && o.name === sheetName && o.size);
      if (!sheetZone || !api) return [];
      const zones = dashM.zones.filter(z => z.type === "color" && z.name === sheetName && !zoneHidden(z));
      const sx = api.size.width / sheetZone.w, sy = api.size.height / sheetZone.h;
      return [pie, pie.inner].filter(s => s && s.legend).map((s, i) => {
        const z = zones.find(z => s.colorRef && z.param && tfSameField(tfParseFieldRef(z.param), s.colorRef)) || (i === 0 && zones.length === 1 ? zones[0] : null);
        if (!z) return null;
        const key = `legend:${z.id}`;
        const obj = { type: "worksheet", name: key, id: z.id, isVisible: true,
                      position: { x: api.position.x + (z.x - sheetZone.x) * sx, y: api.position.y + (z.y - sheetZone.y) * sy },
                      size: { width: z.w * sx, height: z.h * sy } };
        const layout = buildLayoutMap([...(dashboard.objects || []), obj], titleMap).get(key);
        return layout ? { type: /** @type {"legend"} */ ("legend"), name: key, visualName: s.legend.title, layout, legend: s.legend, fmt: vm.fmt, columns: [],
                          rowCount: 1 + s.legend.items.length } : null;
      }).filter(Boolean);
    };

    for (const { sheet, data, visualSpec, visualSpecError, error } of allSheetsData) {
      if (error) {
        console.warn(`Skipping "${sheet.name}": ${error.message}`);
        skipped.push({ name: sheet.name, reason: `FAILED – Tableau did not return its data: ${error.message}` });
        continue;
      }

      // ---- DZV check ----
      if (dzvMap[sheet.name] === false) {
        console.log(`[DZV] Skipping hidden sheet: "${sheet.name}"`);
        skipped.push({ name: sheet.name, reason: "hidden on the dashboard (not visible, or squeezed into a tiny zone)" });
        continue;
      }

      let summaryData = data;  // already fetched
      if (!summaryData.columns || summaryData.columns.length === 0) {
        console.warn(`[Export] Skipping "${sheet.name}": summary data has no columns`);
        skipped.push({ name: sheet.name, reason: "no data: Tableau returned no columns for it" });
        continue;
      }
      if (!summaryData.data || summaryData.data.length === 0) {
        console.warn(`[Export] Skipping "${sheet.name}": summary data has no rows`);
        skipped.push({ name: sheet.name, reason: "no data: no marks with the current filters" });
        continue;
      }

      const layout = layoutMap.get(sheet.name);
      const visualName = (layout && layout.displayName) ? layout.displayName : sheet.name;

      // a button / icon sheet (a custom shape on empty shelves): drawn as its icon, not as a table of its flags
      const icon = fmtModel ? tvIconSheet(fmtModel, sheet.name, summaryData) : null;
      if (icon) {
        if (icon.shape && layout) iconSheets.push({ name: sheet.name, shape: icon.shape, layout });
        else {
          console.log(`[Icons] ${sheet.name}: a Tableau shape – not exported`);
          skipped.push({ name: sheet.name, reason: "a button / icon drawn with one of Tableau's built-in shapes" });
        }
        continue;
      }

      if (isFilterValueWorksheet(sheet.name, summaryData)) {
        let matchedFilterName = null;
        let matchedValues = null;

        for (const [filterField, values] of Object.entries(filterValuesMap)) {
          const cleanFilterField = filterField.toLowerCase().replace(/[^a-z]/g, '');
          const cleanSheetName = sheet.name.toLowerCase().replace(/[^a-z]/g, '');
          if (cleanSheetName.includes(cleanFilterField) || cleanFilterField.includes(cleanSheetName)) {
            matchedFilterName = filterField;
            matchedValues = values;
            break;
          }
        }

        if (!matchedValues) {
          matchedFilterName = visualName;
          matchedValues = [`Values: ${summaryData.data.length} items`];
        }

        filterValueItems.push({
          type: "filterValue",
          name: sheet.name,
          visualName: visualName,
          filterName: matchedFilterName,
          values: matchedValues,
          layout: layout,
          rowCount: 1 + (Array.isArray(matchedValues) ? matchedValues.length : 1),
          originalData: summaryData
        });
      } else {
        // ── NEW: rebuild the visual table (pivot, merge, sort, visible columns, title) ──
        const visualModel = buildVisualModel(fmtModel, sheet.name, summaryData,
          { dashboardName: dashboard.name, displayName: visualName, visualSpec, visualSpecError, layout });
        if (FORMAT_CONFIG.debug) console.log(`[VisualSpec] ${sheet.name}:`, {
          available: visualModel.metadata.visualSpecAvailable,
          keys: visualModel.metadata.visualSpecKeys,
          type: visualModel.type,
          marks: visualModel.metadata.markTokens,
          markSource: visualModel.metadata.markSource,
          error: visualModel.metadata.visualSpecError,
          spec: visualModel.source.visualSpec
        });
        // "Charts as tables": this chart as a crosstab (KPI trend lines in tiny zones stay charts)
        const asTable = visualsAsTables() && !VISUAL_RENDERERS.cellTypes.has(visualModel.type) &&
                        !(FORMAT_CONFIG.crosstabKeepsSparklines && isSparklineZone(layout));
        visualModel.crosstab = asTable;
        let renderDecision = chooseVisualRenderer(visualModel, { asTable });
        // a pie or donut (one or two layers): drawn the way Tableau draws it – slices, labels, hole, tooltips
        if (!asTable && fmtModel && FORMAT_CONFIG.nativeCharts && visualModel.viewModel.fmt.pieLayers()) {
          const pie = buildPieModel(fmtModel, sheet.name, summaryData, {
            dashboardName: dashboard.name, selectedTables: await fetchSelectedMarks(sheet),
            dataOf: name => (allSheetsData.find(x => x.sheet.name === name) || { data: null }).data
          });
          if (pie) {
            visualModel.pie = pie;
            visualModel.type = VISUAL_TYPES.PIE;
            renderDecision = VISUAL_RENDERERS.chart;
            if (pie.centerText) {                              // the text drawn in the hole is not exported again
              if (pie.centerText.name) pieCentreSheets.add(pie.centerText.name);
              else visualModel.viewModel.headerZoneIds.push(pie.centerText.id);
            }
            pieLegendItems.push(...pieLegends(pie, sheet.name, visualModel.viewModel));
          }
        }
        // a gauge drawn with map layers / polygons: rebuilt from its calculations as Excel's doughnut + needle gauge
        if (!asTable && fmtModel && FORMAT_CONFIG.nativeCharts && !visualModel.pie &&
            ((fmtModel.sheets[sheet.name] || {}).panes || []).some(p => /^polygon$/i.test(p.markClass || ""))) {
          const gauge = buildMapGaugeModel(fmtModel, sheet.name, summaryData);
          if (gauge) { visualModel.gauge = gauge; renderDecision = VISUAL_RENDERERS.chart; }
        }
        if (renderDecision.renderer === "excel-chart" && !visualModel.pie && !visualModel.gauge) {
          try {
            visualModel.chartSpecs = buildExcelChartSpecs(visualModel, fmtModel);
            // Tableau draws legends as separate dashboard cards, never inside the view: the chart keeps a
            // legend only when the dashboard shows a colour / size / shape legend for this sheet
            const dashZones = fmtModel && fmtModel.dashboards && fmtModel.dashboards[dashboard.name] ? fmtModel.dashboards[dashboard.name].zones : null;
            const legendZone = dashZones && dashZones.find(z => /^(color|size|shape)$/.test(z.type) && z.name === sheet.name && !zoneHidden(z));
            if (dashZones && !legendZone) {
              visualModel.chartSpecs.forEach(s => { s.legend = false; });
            } else if (legendZone) {
              // the card a strip right above the view (Tableau's horizontal colour legend): written as one line of
              // cells there (legendStrip) – inside the chart it would take a third of a small plot
              const view = dashZones.find(z => z.type === "worksheet" && z.name === sheet.name && z.w && z.h);
              if (view && legendZone.w >= view.w * 0.6 && legendZone.y + legendZone.h <= view.y + view.h * 0.1 &&
                  legendZone.y + legendZone.h >= view.y - view.h * 0.15) visualModel.legendStrip = true;
            }
          } catch (err) {
            console.warn(`[Visual] ${sheet.name}: native chart not possible – ${err.message}`);
            renderDecision = imageOrFallback(visualModel, `native chart not possible (${err.message})`);
          }
        }
        visualModel.renderer = renderDecision.renderer;
        visualModel.status = renderDecision.status;
        visualModel.statusReason = renderDecision.reason;
        visualStatuses.push({
          worksheet: sheet.name,
          type: visualModel.type,
          renderer: renderDecision.renderer,
          status: renderDecision.status,
          reason: renderDecision.reason,
          visualSpecAvailable: visualModel.metadata.visualSpecAvailable,
          imageApiAvailable: visualModel.metadata.imageApiAvailable
        });
        console.log(`[Visual] ${sheet.name}: ${visualModel.type} -> ${renderDecision.renderer} (${renderDecision.reason})`);

        if (renderDecision.renderer === "data-fallback" && visualModel.viewModel.kind === "chart" &&
            FORMAT_CONFIG.chartPolicy === "skip" && !asTable) {
          console.warn(`[Export] "${sheet.name}" is a chart - skipped by FORMAT_CONFIG.chartPolicy`);
          skipped.push({ name: sheet.name, reason: `UNSUPPORTED – ${renderDecision.reason} (chartPolicy "skip")` });
          continue;
        }
        let vm = visualModel.viewModel;
        if (!vm.rows.length || !vm.order.length) {
          console.warn(`[Export] "${sheet.name}" produced an empty formatted view; rebuilding from live summary data`);
          const rawVisualModel = buildVisualModel(null, sheet.name, summaryData,
            { dashboardName: dashboard.name, displayName: visualName, visualSpec, visualSpecError, layout });
          if (rawVisualModel.viewModel.rows.length && rawVisualModel.viewModel.order.length) {
            vm = rawVisualModel.viewModel;
            visualModel.viewModel = vm;
            visualModel.data.rows = vm.rows;
            visualModel.data.columns = vm.cols;
            visualModel.data.dimensions = vm.order.filter(i => vm.cols[i].isHeader).map(i => vm.cols[i]);
            visualModel.data.measures = vm.order.filter(i => !vm.cols[i].isHeader).map(i => vm.cols[i]);
            visualModel.statusReason = "TWB view was empty; exported live summary data";
          }
        }
        if (!vm.rows.length || !vm.order.length) {
          console.warn(`[Export] Skipping "${sheet.name}": no renderable rows or columns after fallback`);
          skipped.push({ name: sheet.name, reason: "FAILED – no rows or columns left to write" });
          continue;
        }
        // a chart exported as data – or a table whose cells would drop what its marks show: axis/colour/detail fields
        // are the data – show every column, except the latitude / longitude Tableau generates only to place a map's marks
        if ((visualModel.renderer === "data-fallback" && !VISUAL_RENDERERS.cellTypes.has(visualModel.type)) || vm.lossless === false) {
          vm.order = vm.cols.map((_, i) => i).filter(i => !/^(latitude|longitude) \(generated\)$/i.test(vm.cols[i].name));
          vm.cols.forEach(c => { if (!c.label) c.label = c.name; });
          vm.showHeaderRow = true;
        }
        const isGraphic = visualModel.renderer === "excel-chart" || visualModel.renderer === "tableau-image";
        const box = isGraphic ? graphicBox(layout, vm) : null;
        const isKPI = !isGraphic && isKPIViewModel(vm);
        // KPI tile: rebuilt from its Tableau label and filling its zone (null = drawn as a label | value table)
        const kpiCard = isKPI ? buildKpiCard(vm, sheet.name, layout) : null;
        const kpiW = kpiCard && layout && layout.widthPx ? Math.max(kpiCard.tiles.length, layout.gridRight - layout.gridCol) : null;
        // a table whose column headers are dashboard text boxes above its zone starts where those boxes start –
        // in Tableau the headers sit there, directly under the title, not at the top of the worksheet zone
        let placed = layout;
        if (layout && !box && !kpiCard && vm.headerZoneIds && vm.headerZoneIds.length) {
          const tops = vm.headerZoneIds.map(id => layoutMap.get(`text:${id}`)).filter(Boolean).map(l => l.gridRow);
          const top = tops.length ? Math.min(...tops) : layout.gridRow;
          if (top < layout.gridRow) placed = { ...layout, gridRow: top };
        }
        dataWorksheetItems.push({
          name: sheet.name,
          visualName: vm.title.text,
          layout: placed,
          isKPI,
          kpiCard,
          type: "worksheet",
          vm,
          visualModel,
          box,                                                 // chart/image: pixel size from the dashboard zone
          fixedGridW: box ? box.gridW : kpiW,
          columns: vm.order,                                   // only its length is used for layout
          rowCount: box ? box.rows : kpiCard ? kpiCard.rows : isKPI ? (vm.showTitle ? 1 : 0) + vm.order.length : viewModelHeight(vm)
        });
      }
    }

    // ── parameter controls (e.g. Start Date / End Date): not filters, so read them separately ──
    /** @type {{ name: string, value: string }[]} every parameter's value at export time (conversion report) */
    const reportParams = [];
    try {
      const params = typeof dashboard.getParametersAsync === "function" ? await dashboard.getParametersAsync() : [];
      const controls = (dashboard.objects || []).filter(o => OBJECT_KIND[o.type] === "parameter" && o.isVisible !== false);
      const used = new Set();
      const norm = s => String(s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
      // the workbook's control zones: which parameter each shows ([Parameters].[Parameter 3]) and its custom title
      // ("Current Start Date") – the dashboard object may carry that title instead of the parameter's name
      const ctrlZones = new Map(((fmtModel && fmtModel.dashboards[dashboard.name] || { zones: [] }).zones)
        .filter(z => z.type === "paramctrl" && z.param).map(z => [String(z.id), z]));
      const paramIdOf = p => norm(p.id || "") || norm(`[Parameters].[${p.name}]`);
      params.forEach(param => {
        const raw = param.currentValue ? (param.currentValue.formattedValue ?? String(param.currentValue.value)) : "";
        const value = String(raw).replace(/\s+12:00:00\s*AM$|\s+00:00:00$/i, "");   // dates: no midnight time
        reportParams.push({ name: param.name, value });
        const info = fmtModel && Object.values(fmtModel.fields).find(f => f.param && norm(f.caption || f.name) === norm(param.name));
        const ids = [paramIdOf(param), info && norm(`[Parameters].[${info.name}]`)].filter(Boolean);
        const control = controls.find(o => !used.has(o) && ctrlZones.has(String(o.id)) && ids.includes(norm(ctrlZones.get(String(o.id)).param))) ||
                        controls.find(o => !used.has(o) && !ctrlZones.has(String(o.id)) && norm(o.name) === norm(param.name)) ||
                        controls.find(o => !used.has(o) && !ctrlZones.has(String(o.id)) && norm(o.name).includes(norm(param.name)));
        if (!control) return;                                      // parameter not shown on this dashboard
        used.add(control);
        const zone = ctrlZones.get(String(control.id));
        const title = (zone && zone.customTitle && tfZoneText(zone)) || param.name;
        filterValueItems.push({
          type: "filterValue", isParameter: true, name: param.name, visualName: title,
          filterName: title, values: [value], layout: layoutMap.get(control.name), rowCount: 2, param
        });
      });
      if (params.length) console.log(`[Parameters] ${used.size} of ${params.length} parameters have a control on the dashboard`);
    } catch (err) {
      console.warn("[Parameters] could not read parameters:", err.message);
    }

    // ── dashboard filter controls (Select Chain: (All) …): a card each, where Tableau shows it ──
    if (FORMAT_CONFIG.filterCards && fmtModel) {
      try {
        const taken = new Set(filterValueItems.map(it => it.layout).filter(Boolean));   // a spot already used (filter-like sheet)
        // a card only fills empty space: one that would sit on a worksheet, text box, image or parameter control
        // (e.g. a filter floating over a chart) is left out, so the rest of the layout never moves for it
        const blocks = [];
        layoutMap.forEach((l, key) => { if (l && /^(worksheet|text|image|parameter)$/.test(l.type) && !key.startsWith("filter:")) blocks.push(l); });
        const rect = l => ({ r0: l.gridRow, r1: l.gridRow + Math.max(1, l.gridH || 1), c0: l.gridCol, c1: Math.max(l.gridCol + 1, l.gridRight || l.gridCol + (l.gridW || 1)) });
        const hits = (a, b) => a.r0 < b.r1 && b.r0 < a.r1 && a.c0 < b.c1 && b.c0 < a.c1;
        (await buildFilterCards(dashboard, sheets, fmtModel, layoutMap)).forEach(card => {
          if (taken.has(card.layout)) return;
          const me = { ...rect(card.layout), r1: card.layout.gridRow + 2 };
          const under = blocks.find(b => b !== card.layout && hits(me, rect(b)));
          if (under) { console.log(`[Filters] "${card.visualName}" sits on "${under.originalName || under.type}" – card left out`); return; }
          filterValueItems.push(card);
        });
      } catch (err) {
        console.warn("[Filters] filter cards left out:", err.message);
      }
    }

    // ── filter cards of a field's values → Excel slicers; the visuals whose numbers the data source's rows rebuild
    //    follow them (vm.slicerBinding: their chart values, table cells and KPI values become formulas)
    const slicers = ctx.slicers;
    if (slicers) {
      for (const card of filterValueItems.filter(it => it.isFilterCard)) {
        const ws = sheets.find(s => s.name === card.sheetName);
        card.model = fmtModel;                                       // a range of an aggregate translates its measure
        const field = ws ? await slicers.adoptCard(card, ws, fmtModel) : "its worksheet is not on the dashboard";
        if (typeof field === "string") { console.log(`[Slicers] "${card.visualName}" stays a card: ${field}`); continue; }
        if (field.kind !== "slicer") {                                   // a dropdown cell, or From / To date cells
          card.control = { field };
          card.rowCount = field.kind === "list" ? 2 : 3;                // title + the value, or title + Min / Max (From / To)
          continue;
        }
        const size = slicers.sizeOf(field, card.layout);
        card.slicer = { field, ...size };
        card.rowCount = Math.max(card.rowCount, Math.ceil(size.heightPx / EXCEL_ROW_PX));
      }
      // Start / End date parameters a calculated filter keeps a date between (Date Flag: [Order Date] >= [Start Date]
      // AND [Order Date] <= [End Date]): their cells become the range's From / To
      const dateParams = filterValueItems.filter(it => it.isParameter && it.param && it.layout && /date/i.test(String(it.param.dataType || "")));
      if (dateParams.length) {
        // the parameter's control: by its id ([Parameters].[Parameter 1]), else by its caption
        const paramCard = name => {
          const id = tfNorm(`[Parameters].[${name}]`);
          const info = fmtModel.fields["Parameters|" + String(name).toLowerCase()];
          const names = [name, info && info.caption].filter(Boolean).map(tfNorm);
          return dateParams.find(it => it.param.id && tfNorm(it.param.id) === id) ||
                 dateParams.find(it => names.includes(tfNorm(it.param.name)));
        };
        const seen = new Set();
        for (const item of dataWorksheetItems) {
          for (const ref of ((fmtModel.sheets[item.name] || {}).filterRefs || [])) {
            const key = `${ref.ds || ""}|${ref.name}`;
            if (seen.has(key)) continue;
            seen.add(key);
            const info = tfFieldInfo(fmtModel, ref);
            const spec = info && slicerParamRange(info.formulaFull || info.formula);
            const lo = spec && paramCard(spec.lo), hi = spec && paramCard(spec.hi);
            if (!spec || (!lo && !hi)) continue;
            // a bound whose parameter has no control on the dashboard stays at the workbook's value of it (Current
            // date End 1 beside a Current date End control): only the other one is a date cell
            const fixedBound = name => {
              const p = fmtModel.fields["Parameters|" + String(name).toLowerCase()];
              const m = p && String(p.value || "").match(/^#(\d{4}-\d\d-\d\d)(?:[ T][\d:]+)?#$/);
              return m ? { value: m[1], formattedValue: m[1] } : null;
            };
            const loDv = lo ? lo.param.currentValue : fixedBound(spec.lo), hiDv = hi ? hi.param.currentValue : fixedBound(spec.hi);
            if (!loDv || !hiDv) continue;
            const ws = sheets.find(s => s.name === item.name);
            const field = await slicers.adoptParamRange({ calc: ref, date: spec.date, lo: loDv, hi: hiDv,
                                                          loStrict: spec.loStrict, hiStrict: spec.hiStrict }, ws, fmtModel);
            const between = `dates between ${lo ? lo.param.name : spec.lo} and ${hi ? hi.param.name : spec.hi}`;
            if (typeof field === "string") { console.log(`[Slicers] ${info.caption || ref.name} (${between}) stays fixed: ${field}`); continue; }
            [[lo, "from"], [hi, "to"]].forEach(([card, end]) => {
              if (!card) return;
              (card.paramFields = card.paramFields || []).push({ field, end });
              card.rowCount = 2;                                 // title, Month / Day / Year dropdowns
            });
          }
        }
      }
      if (filterValueItems.some(it => it.slicer || it.control || it.paramFields)) {
        for (const item of dataWorksheetItems) {
          const vmod = item.visualModel;
          const pie = vmod && vmod.pie;
          if (!item.vm || !vmod || vmod.gauge || vmod.renderer === "tableau-image" || (pie && (pie.gauge || pie.inner))) {
            item.slicerNote = pie && pie.inner ? "fixed values: a nested donut keeps Tableau's numbers"
              : "fixed values: gauges and pictures keep Tableau's numbers";
            continue;
          }
          const sheet = sheets.find(s => s.name === item.name);
          let filters = [];
          try { filters = sheet && typeof sheet.getFiltersAsync === "function" ? (await sheet.getFiltersAsync()) || [] : []; }
          catch (e) { /* bound without its filters: the check against Tableau's numbers decides */ }
          if (pie) {
            const binding = bindPie(slicers, pie, item.name, filters, fmtModel);
            if (typeof binding === "string") item.slicerNote = `fixed values: ${binding}`;
            else {
              item.vm.slicerBinding = pie.slicerBinding = binding;
              item.slicerNote = "follows the slicers (Excel draws its labels; Tableau's tooltips are left out" +
                (pie.centerText ? "; the text over its centre keeps Tableau's numbers)" : ")");
            }
            continue;
          }
          const binding = slicers.bind(item.vm, item.name, filters, fmtModel);
          if (typeof binding === "string") item.slicerNote = `fixed values: ${binding}`;
          else {
            item.vm.slicerBinding = binding;
            item.slicerNote = binding.reasons.length ? `follows the slicers, except ${binding.reasons.join("; ")}` : "follows the slicers";
          }
        }
        if (!dataWorksheetItems.some(it => it.vm && it.vm.slicerBinding)) {
          console.log(`[Slicers] ${dashboard.name}: no visual can follow the slicers – the filter cards stay cards`);
          filterValueItems.forEach(it => { if (it.slicer || it.control || it.paramFields) { delete it.slicer; delete it.control; delete it.paramFields; it.rowCount = 2; } });
          slicers.discard();
        }
      }
      // the other date parameters' controls (no visual here uses them): the same Month / Day / Year dropdowns, so every
      // date card looks alike – the lists sheet that holds them is written while the slicers are
      if (slicers.active) {
        filterValueItems.filter(it => it.isParameter && it.param && it.layout && !it.paramFields && /date/i.test(String(it.param.dataType || "")))
          .forEach(it => {
            const serial = tvExcelSerial(it.param.currentValue);
            if (serial !== null) { it.dateOnly = true; it.paramSerial = Math.round(serial); }
          });
      }
    }

    // Tableau's number + trend tile: a chart in the zone right under / over a KPI card, same position and
    // width, takes the card's columns so the two line up as one card
    const near = (a, b) => Math.abs(a - b) <= 8;
    dataWorksheetItems.filter(it => it.kpiCard && it.fixedGridW && it.layout).forEach(card => {
      const c = card.layout;
      dataWorksheetItems.forEach(other => {
        const o = other.layout;
        if (other === card || !other.box || !o || !o.widthPx || !near(o.xPx, c.xPx) || !near(o.widthPx, c.widthPx)) return;
        if (near(o.yPx, c.yPx + c.heightPx) || near(c.yPx, o.yPx + o.heightPx)) { other.fixedGridW = card.fixedGridW; other.pairedCard = true; }
      });
    });

    // ── dashboard text boxes (banners, titles, notes): drawn where Tableau puts them, as Tableau formats them
    const dashZones = (fmtModel && fmtModel.dashboards && fmtModel.dashboards[dashboard.name] || { zones: [] }).zones;
    const headerZoneIds = new Set(dataWorksheetItems.flatMap(it => it.vm ? it.vm.headerZoneIds : []));   // already table headers
    /** @type {ExportItem[]} */
    const textItems = [];
    (dashboard.objects || []).filter(o => OBJECT_KIND[o.type] === "text" && o.isVisible !== false).forEach(o => {
      const zone = dashZones.find(z => z.type === "text" && String(z.id) === String(o.id));
      if (!zone || zoneHidden(zone) || headerZoneIds.has(zone.id)) return;
      const layout = layoutMap.get(`text:${o.id}`);
      const card = buildTextCard(zone, layout, fmtModel);
      if (!card) return;
      textItems.push({ type: "text", name: `text:${o.id}`, visualName: tfZoneText(zone), layout, textCard: card,
        rowCount: card.rows, columns: [], fixedGridW: layout && layout.widthPx ? Math.max(1, layout.gridRight - layout.gridCol) : null });
    });
    if (textItems.length) console.log(`[Text] ${textItems.length} dashboard text box(es)`);

    // ── dashboard images (logos, icons) from the files packaged in the .twbx: a logo in its own zone is
    //    a block of the layout, an icon over a sheet floats on that sheet's block, backgrounds are left out
    const imageKinds = classifyImageObjects(dashboard.objects || []);
    const imageObjects = (dashboard.objects || []).filter(o => o.type === "image" && o.isVisible !== false && imageKinds.has(String(o.id)));
    const workbookImages = imageObjects.length || iconSheets.length ? await getWorkbookImages() : {};
    /** @type {ExportItem[]} */
    const imageItems = [];
    /** @type {(ImageBlock & { host: string, xPx: number, yPx: number })[]} */
    const overlayImages = [];
    imageObjects.forEach(o => {
      const cls = imageKinds.get(String(o.id));
      const zone = dashZones.find(z => z.type === "bitmap" && String(z.id) === String(o.id));
      if (!zone || !zone.param || zoneHidden(zone)) return;
      if (cls.kind === "backdrop" || cls.kind === "tiny") { console.log(`[Images] ${zone.param}: ${cls.kind === "tiny" ? "divider" : "background"} – not exported`); return; }
      const file = workbookImages[zone.param];
      if (!file) { console.warn(`[Images] ${zone.param} is not in the loaded workbook – load the .twbx to export it`); return; }
      const info = imageInfo(file.data, zone.param);
      if (!info.type) { console.warn(`[Images] ${zone.param}: not a PNG, JPEG, GIF or SVG file`); return; }
      /** @type {ImageBlock} */
      const image = { file, info, zone, url: webLink(zone.url), widthPx: o.size.width, heightPx: o.size.height };
      if (cls.kind === "overlay") { overlayImages.push({ ...image, host: cls.host, xPx: o.position.x, yPx: o.position.y }); return; }
      const layout = layoutMap.get(`image:${o.id}`);
      if (!layout) return;
      imageItems.push({ type: "image", name: `image:${o.id}`, visualName: zone.param, layout, image, columns: [],
        rowCount: Math.max(1, Math.round(o.size.height / EXCEL_ROW_PX)),
        fixedGridW: layout.widthPx ? Math.max(1, layout.gridRight - layout.gridCol) : null });
    });
    // icon sheets: the shape centred in the sheet's zone at about Tableau's mark size
    iconSheets.forEach(({ name, shape, layout }) => {
      const file = workbookImages["shape:" + shape];
      if (!file) { console.warn(`[Icons] ${name}: shape "${shape}" not in the loaded workbook`); return; }
      const info = imageInfo(file.data, shape);
      const w = layout.widthPx || 40, h = layout.heightPx || 40;
      const box = Math.min(28, Math.max(12, Math.round(Math.min(w, h) * 0.5)));
      if (!info.type) return;
      imageItems.push({ type: "image", name, visualName: name, layout, columns: [],
        image: { file, info, zone: { param: shape, scaled: true, centered: true }, url: null, widthPx: box, heightPx: box },
        rowCount: Math.max(1, Math.round(h / EXCEL_ROW_PX)),
        fixedGridW: layout.widthPx ? Math.max(1, layout.gridRight - layout.gridCol) : null });
    });
    if (imageItems.length || overlayImages.length) console.log(`[Images] ${imageItems.length + overlayImages.length} image(s)`);

    const allItems = [...filterValueItems, ...dataWorksheetItems.filter(it => !pieCentreSheets.has(it.name)),
                      ...textItems, ...imageItems, ...pieLegendItems];

    if (allItems.length === 0) {
      return null;   // nothing exportable on this dashboard → no sheet for it
    }

    console.table(visualStatuses);

    // columns from the dashboard's zone edges: every block keeps its Tableau position and width; a table gets
    // a column per field inside its zone, split by how much its fields hold
    allItems.forEach(item => {
      if (item.vm && item.type === "worksheet" && !item.box && !item.kpiCard && !item.isKPI) {
        item.splitPx = tableColumnPx(item.vm);                 // the widths Tableau draws the table at
      } else if (item.isKPI && !item.kpiCard) item.split = [1, 1];
      else if (item.kpiCard && item.kpiCard.tiles.length > 1) item.split = item.kpiCard.tiles.map(() => 1);
      else if (item.paramFields || item.dateOnly) item.split = [4, 3, 4];  // a date's Month / Day / Year dropdowns
    });
    const grid = buildColumnGrid(allItems);
    const placedItems = allItems.map((item, idx) => {
      const span = grid.span(item);
      if (item.layout && span) {
        // a table covers every column it writes (it may run past its zone into free space)
        const fieldsEnd = span.fields && span.fields.length ? span.fields[span.fields.length - 1].offset + span.fields[span.fields.length - 1].span : 0;
        const gridW = Math.max(span.gridW, fieldsEnd);
        return { ...item, gridRow: item.layout.gridRow, gridCol: span.gridCol, gridW, allocatedRows: item.rowCount,
                 splitCuts: span.cuts, fieldCols: span.fields };
      }
      if (item.layout) {
        const l = item.layout;
        let calculatedWidth = l.gridW;

        if (item.fixedGridW) {
          calculatedWidth = item.fixedGridW;
        } else if (item.type === "filterValue") {
          calculatedWidth = Math.max(l.gridW, 3);
        } else if (item.isKPI) {
          calculatedWidth = Math.max(l.gridW, Math.min(6, item.columns.length + 1));
        } else {
          calculatedWidth = Math.max(l.gridW, Math.min(25, item.columns.length + 1));
        }

        return {
          ...item,
          gridRow: l.gridRow,
          gridCol: l.gridCol,
          gridW: calculatedWidth,
          allocatedRows: item.rowCount,
        };
      } else {
        const itemsPerRow = 3;
        const rowIdx = Math.floor(idx / itemsPerRow);
        const colIdx = idx % itemsPerRow;
        const gridRow = rowIdx * (Math.max(8, item.rowCount) + ROW_GAP);
        const gridCol = colIdx * 12;
        const gridW = item.fixedGridW || ((item.type === "filterValue") ? 4 : (item.isKPI ? 5 : 12));

        return {
          ...item,
          gridRow: gridRow,
          gridCol: gridCol,
          gridW: gridW,
          allocatedRows: item.rowCount,
        };
      }
    });

    resolveCollisions(placedItems);

// SNAP-OUT-OF-GROUPED-ROWS PASS
// Excel hides entire physical rows, not per-column cells, so a long table's collapsed rows would
// also hide whatever sits beside it. Charts, images, cards and filter lists that share rows with
// the table keep those rows visible (setTableVisibleRows); a table that would start inside another
// table's collapsed rows is pushed below that table's full physical row range.
let snapChanged = true;
const MAX_SNAP_PASSES = 10;
let snapPass = 0;

while (FORMAT_CONFIG.groupOverflowRows && snapChanged && snapPass < MAX_SNAP_PASSES) {
  snapChanged = false;
  snapPass++;
  setTableVisibleRows(placedItems);               // charts / cards beside a long table keep its rows visible

  for (const item of placedItems) {
    for (const other of placedItems) {
      if (item === other) continue;
      if (!other.visibleRows || other.visibleRows >= other.vm.rows.length) continue;  // other hides no rows

      const otherTotal = other.allocatedRows || other.rowCount || 0;
      const otherVisibleEnd = tableDataStart(other) + other.visibleRows;               // first hidden row
      const otherPhysicalEnd = other.gridRow + otherTotal + ROW_GAP;

      if (item.gridRow >= otherVisibleEnd && item.gridRow < otherPhysicalEnd) {
        item.gridRow = otherPhysicalEnd;
        snapChanged = true;
      }
    }
  }
  // a table moved below another's collapsed rows can now sit on what was under it: everything below moves
  // down with it (resolveCollisions only ever moves items down, and only where they really overlap)
  if (snapChanged) resolveCollisions(placedItems);
}
setTableVisibleRows(placedItems);

    const worksheet = workbook.addWorksheet(sheetName);
    // the sheet as the dashboard's canvas: no cell grid, as in Tableau
    worksheet.views = [{ showGridLines: !!FORMAT_CONFIG.sheetGridlines }];

    /** @type {Record<number, number>} */
    const colWidths = {};
    const exactWidths = {};
    const allTablesInfo = [];
    const tracker = makeRangeTracker();

    let currentRow = 0;
    const dashboardName = dashboard.name || "Dashboard Export";
    let dashTitleProps = null;
    const dashFmt = fmtModel && fmtModel.dashboards && fmtModel.dashboards[dashboard.name];
    if (dashFmt) {
      const run = dashFmt.title && (dashFmt.title.runs.find(x => x.text.trim()) || {}).props;
      dashTitleProps = tfMerge(TABLEAU_DEFAULTS.dashTitle,
        tfCollect(fmtModel.workbookStyle, TF_ELEMENTS.dashTitle), tfCollect(dashFmt.style, TF_ELEMENTS.dashTitle), run);
    }
    const usedZones = new Set(dataWorksheetItems.flatMap(it => it.vm ? it.vm.headerZoneIds : []));
    const titleRuns = tfDashboardTitleRuns(fmtModel, dashboard.name, usedZones);
    // Tableau shows the dashboard title only when its title zone is on (text boxes are drawn in place)
    const titleShown = !dashFmt || dashFmt.zones.some(z => z.type === "title" && !z.hidden);
    const titleHeight = writeDashboardTitle(worksheet, titleShown ? dashboardName : "", currentRow, 0, tracker, dashTitleProps,
      titleShown ? titleRuns : null);
    currentRow += titleHeight;


    const adjustedItems = placedItems.map(item => ({
      ...item,
      gridRow: item.gridRow + currentRow
    }));
    // the views the buttons switch to (withAlternateViews) start two rows under the dashboard's blocks: a row for
    // their header (the layout closes up empty rows, so the gap between them is not kept)
    const zoneOf = it => (it.layout && it.layout.id !== undefined ? String(it.layout.id) : "");
    // views of one kind that a dropdown switches in one chart (findViewSwitches): the ones Tableau hides are not drawn
    const switches = fmtModel ? findViewSwitches(adjustedItems, dashFmt, alternates, fmtModel, dashboard.name, zoneOf) : [];
    const switchOf = new Map(switches.map(g => [g.host, g]));
    switches.forEach(g => g.dropped.forEach(it => { const at = adjustedItems.indexOf(it); if (at >= 0) adjustedItems.splice(at, 1); }));
    if (switches.length) console.log(`[Views] ${dashboard.name}: ${switches.map(g => g.views.map(v => v.item.visualName || v.item.name).join(" / ")).join(", ")} – one chart each, switched by a dropdown`);
    const altItems = adjustedItems.filter(it => alternates.has(zoneOf(it)));
    const dashEnd = Math.max(0, ...adjustedItems.filter(it => !alternates.has(zoneOf(it))).map(it => it.gridRow + Math.max(1, it.allocatedRows || 1)));
    if (altItems.length) {
      const shift = dashEnd + 2 - Math.min(...altItems.map(it => it.gridRow));
      if (shift > 0) altItems.forEach(it => { it.gridRow += shift; });
    }

    // native charts: data goes to a hidden sheet now, chart parts are injected after writeBuffer()
    const chartJobs = [];
    const reserveGraphicBlock = item => {
      for (let c = item.gridCol; c < item.gridCol + item.gridW; c++) colWidths[c] = Math.max(colWidths[c] || 0, 10);
      tracker.update(item.gridRow, item.gridCol);
      tracker.update(item.gridRow + item.allocatedRows - 1, item.gridCol + item.gridW - 1);
    };
    const writeGraphicTitle = item => {
      if (!item.vm.showTitle) return 0;
      writeTableauTitle(worksheet, item.gridRow, item.gridCol, item.vm.title.text, item.vm.title.props, item.gridW);
      // switched views: the picked one's title
      const group = switchOf.get(item);
      const titles = group ? group.views.map(v => (v.item.vm && v.item.vm.showTitle ? v.item.vm.title.text : "") || "") : [];
      if (group && group.sw && new Set(titles).size > 1)
        worksheet.getCell(item.gridRow + 1, item.gridCol + 1).value = { formula: group.sw.pick(titles.map(xlText)), result: item.vm.title.text };
      return 1;
    };
    /** a chart's spec before its data is written: its values as formulas of the slicers, fitted to its block */
    const prepSpec = (item, spec, widthPx, heightPx) => {
      if (item.vm.slicerBinding) {
        const binding = item.vm.slicerBinding;
        // a 100 % stacked chart without value labels: Excel draws each stack's shares itself from the base measure
        const shares = !!spec.percent && !spec.series.some(s => s.labels || s.labelTexts);
        spec.valueFormula = shares ? (values, i) => binding.valueFormula(values, i, true) : binding.valueFormula;
        // a waterfall's bars are computed from its steps (cells of its own), not taken from the marks
        if (/** @type {any} */ (spec).waterfall && bindWaterfall(spec, item.vm.slicerBinding)) {
          spec.valueFormula = undefined;
          delete spec.valueMin; delete spec.valueMax;              // Excel scales the axis to the steps shown
        }
        // an axis range taken from Tableau's values would clip the values another selection gives: Excel scales it
        if (spec.dataRange && spec.dataRange.min) delete spec.valueMin;
        if (spec.dataRange && spec.dataRange.max) delete spec.valueMax;
      }
      spec.widthPx = widthPx;                          // writeChartData fits category names to it
      spec.heightPx = heightPx;
    };
    /** a switch's cell on the chart data sheet, made with its first chart */
    const switchFor = group => {
      if (!group.sw) { group.sw = makeSwitch(chartDataSheetOf(), chartData.row, group); chartData.row += 2; }
      return group.sw;
    };
    const switchedNote = group => ` – switches between ${group.views.map(v => v.item.visualName || v.item.name).join(" / ")} with a dropdown, as Tableau's buttons do`;

    /** @type {{ item: ExportItem, top: number, panes: Map<string, any> }[]} pies, laid out with the final column widths */
    const pieJobs = [];
    /** @type {ExportItem[]} filter cards drawn as slicers, placed with the final column widths */
    const slicerJobs = [];
    let controlCards = 0;                                          // dropdown / From–To cards written
    // one hidden data sheet for the charts of every dashboard in the file (ctx.chartData: { sheet, row })
    const chartDataSheetOf = () => {
      if (!chartData.sheet) chartData.sheet = workbook.addWorksheet(CHART_DATA_SHEET, { state: "hidden" });
      return chartData.sheet;
    };

    for (let i = 0; i < adjustedItems.length; i++) {
      const item = adjustedItems[i];

      if (item.visualModel && item.visualModel.gauge) {
        const status = visualStatuses.find(entry => entry.worksheet === item.name);
        const titleRows = writeGraphicTitle(item);
        const data = writeGaugeData(chartDataSheetOf(), chartData.row + 1, item.visualModel.gauge, item.name);
        chartData.row = data.nextRow - 1;
        pieJobs.push({ item, top: item.gridRow + titleRows, panes: data.panes, gauge: true });
        reserveGraphicBlock(item);
        item.visualModel.status = "success";
        item.visualModel.statusReason = "Excel gauge (doughnut dial + pie needle) rebuilt from the workbook's calculations";
        if (status) { status.status = "success"; status.reason = item.visualModel.statusReason; }
        continue;
      }

      if (item.visualModel && item.visualModel.pie) {
        const status = visualStatuses.find(entry => entry.worksheet === item.name);
        const group = switchOf.get(item);
        if (group) switchFor(group);
        const titleRows = writeGraphicTitle(item);
        const data = group ? writeSwitchedPie(chartDataSheetOf(), chartData.row + 1, group, item.visualName || item.name)
          : writePieData(chartDataSheetOf(), chartData.row + 1, item.visualModel.pie, item.name);
        chartData.row = data.nextRow - 1;
        pieJobs.push({ item, top: item.gridRow + titleRows, panes: data.panes, pie: group ? /** @type {any} */ (data).pie : null });
        reserveGraphicBlock(item);
        const pie = item.visualModel.pie;
        item.visualModel.status = "success";
        item.visualModel.statusReason = `native Excel ${pie.hole || pie.inner ? "donut" : "pie"} chart ` +
          (pie.slicerBinding ? "following the slicers, its labels drawn by Excel from formula cells" : "with Tableau labels and tooltips") +
          (group ? switchedNote(group) : "");
        if (status) { status.status = "success"; status.reason = item.visualModel.statusReason; }
        continue;
      }

      if (item.type === "legend") {
        writeLegendBlock(worksheet, item, tracker, colWidths);
        continue;
      }

      if (item.visualModel && item.visualModel.renderer === "excel-chart") {
        const status = visualStatuses.find(entry => entry.worksheet === item.name);
        const group = switchOf.get(item);
        if (group) switchFor(group);
        const titleRows = writeGraphicTitle(item);
        const specs = item.visualModel.chartSpecs;
        const paneH = Math.floor(item.box.heightPx / specs.length);     // separate panes stack vertically
        if (group) {
          const w = writeSwitchedChart(chartDataSheetOf(), chartData, group, (it, spec) => prepSpec(it, spec, item.box.widthPx, paneH));
          // the legend line above it switches too
          if (item.visualModel.legendStrip && w.spec.legend && item.gridRow > 0 &&
              writeSwitchLegend(worksheet, item.gridRow, item.gridCol, grid.colPx, item.box.widthPx, w.parts, group.sw, tfExcelFont(item.vm.fmt.baseStyle())))
            w.spec.legend = false;
          chartJobs.push({ spec: w.spec, refs: w.refs, names: w.names, item, name: item.visualName || item.name,
            col: item.gridCol, row: item.gridRow + titleRows, rowOffPx: 0, widthPx: item.box.widthPx, heightPx: paneH,
            top: item.gridRow + titleRows, pane: 0, panes: 1 });
          reserveGraphicBlock(item);
          item.visualModel.status = "success";
          item.visualModel.statusReason = `native Excel ${w.spec.kind} chart${switchedNote(group)}`;
          if (status) { status.status = "success"; status.reason = item.visualModel.statusReason; }
          continue;
        }
        // Tableau's colour legend in a strip above the view: one line of "■ name" in the row above the chart, in the
        // legend's order, cut with "…" where the chart's width ends (Tableau scrolls it); the chart then has none
        const lead = specs[0];
        const above = item.visualModel.legendStrip && specs.length === 1 && lead.legend && item.gridRow > 0
          ? worksheet.getCell(item.gridRow, item.gridCol + 1) : null;
        if (above && (above.value === null || above.value === "")) {
          const font = tfExcelFont(item.vm.fmt.baseStyle());
          const entries = lead.series.filter(s => s.color && !s.refLine).map(s => ({ name: s.name, color: String(s.color).replace(/^#/, "").slice(-6).toUpperCase() }));
          if (lead.stacked && lead.barDir !== "bar") entries.reverse();        // stacked columns: the legend's order
          const charPx = (font.size || 9) * 4 / 3 * 0.55, room = (item.box.widthPx || 300) / charPx;
          const runs = [];
          let used = 0;
          for (const e of entries) {
            const need = e.name.length + 5;
            if (used + need > room) { runs.push({ text: " …", font }); break; }
            runs.push({ text: "■ ", font: { ...font, size: (font.size || 9) + 2, color: { argb: "FF" + e.color } } }, { text: e.name + "   ", font });
            used += need;
          }
          if (runs.length) {
            above.value = { richText: runs };
            above.alignment = { vertical: "bottom", wrapText: false };
            specs.forEach(s => { s.legend = false; });
          }
        }
        specs.forEach((spec, k) => {
          prepSpec(item, spec, item.box.widthPx, paneH);
          const refs = ExcelChartWriter.writeChartData(chartDataSheetOf(), spec, chartData.row);
          chartData.row = refs.nextRow;
          const offPx = k * paneH;
          chartJobs.push({
            spec, refs, item, name: item.visualName || item.name,
            col: item.gridCol, row: item.gridRow + titleRows + Math.floor(offPx / EXCEL_ROW_PX),
            rowOffPx: offPx % EXCEL_ROW_PX, widthPx: item.box.widthPx, heightPx: paneH,
            top: item.gridRow + titleRows, pane: k, panes: specs.length
          });
        });
        reserveGraphicBlock(item);
        item.visualModel.status = "success";
        item.visualModel.statusReason = `native Excel ${[...new Set(specs.map(s => s.kind))].join("/")} chart` +
          (specs.length > 1 ? ` (${specs.length} panes${specs[0].grid ? ` in a ${specs[0].grid.rows} × ${specs[0].grid.cols} grid` : ""})` : "");
        if (status) { status.status = "success"; status.reason = item.visualModel.statusReason; }
        continue;
      }

      if (item.visualModel && item.visualModel.renderer === "tableau-image") {
        const imageWidth = item.box ? item.box.widthPx : Math.max(320, item.gridW * PX_PER_COL);
        const imageHeight = item.box ? item.box.heightPx : Math.max(160, (item.allocatedRows || item.rowCount || 8) * PX_PER_ROW);
        const status = visualStatuses.find(entry => entry.worksheet === item.name);
        try {
          // bounded: after other dashboards were read through embedded views, Tableau may never answer
          const imageBase64 = await withTimeout(renderTableauImage(item.visualModel, imageWidth, imageHeight),
            IMAGE_RENDER_TIMEOUT_MS, "Tableau's picture of the visual");
          const titleRows = writeGraphicTitle(item);
          const imageId = workbook.addImage({ base64: imageBase64, extension: "png" });
          // a picture is not data: the visual's marks go to the Visual Data sheet too, and the picture links to them
          const dataLink = appendVisualData(workbook, item.visualName || item.name, item.vm);
          worksheet.addImage(imageId, /** @type {any} */ ({
            tl: { col: item.gridCol, row: item.gridRow + titleRows },
            ext: { width: imageWidth, height: imageHeight },
            hyperlinks: { hyperlink: dataLink, tooltip: "The data of this picture (Visual Data sheet)" }
          }));
          reserveGraphicBlock(item);
          item.visualModel.status = "success";
          item.visualModel.statusReason = "Tableau-rendered picture; its data on the Visual Data sheet";
          if (status) {
            status.status = "success";
            status.reason = item.visualModel.statusReason;
          }
          continue;
        } catch (err) {
          item.visualModel.renderer = "data-fallback";
          item.visualModel.status = "warning";
          item.visualModel.statusReason = `image renderer failed: ${err.message}`;
          // exported as its data: every column, as for any chart that falls back to a table
          item.vm.order = item.vm.cols.map((_, i) => i).filter(i => !/^(latitude|longitude) \(generated\)$/i.test(item.vm.cols[i].name));
          item.vm.cols.forEach(c => { if (!c.label) c.label = c.name; });
          item.vm.showHeaderRow = true;
          if (status) {
            status.renderer = "data-fallback";
            status.status = "warning";
            status.reason = item.visualModel.statusReason;
          }
          console.warn(`[Visual] ${item.name}: image renderer failed; using data fallback`, err);
        }
      }

      if (item.type === "filterValue" && item.slicer) {           // drawn by injectSlicers, over its block
        reserveGraphicBlock(item);
        slicerJobs.push(item);
        continue;
      }
      if (item.type === "filterValue" && item.control) {          // a dropdown cell / From–To cells the visuals follow
        writeControlCard(worksheet, item, tracker);
        controlCards++;
        continue;
      }
      if (item.type === "filterValue" && (item.paramFields || item.dateOnly)) {   // a date parameter: Month / Day / Year dropdowns
        writeParamDateCard(worksheet, item, tracker, ctx.slicers);
        controlCards++;
        continue;
      }

      if (item.type === "filterValue") {
        writeIndividualFilterTable(
          worksheet,
          item.filterName,
          item.values,
          item.gridRow,
          item.gridCol,
          tracker,
          item.gridW
        );
        if (!item.layout) colWidths[item.gridCol] = Math.max(colWidths[item.gridCol] || 0, 35);
      } else if (item.image) {
        tracker.update(item.gridRow, item.gridCol);                 // the picture goes in once the columns are sized
        tracker.update(item.gridRow + item.allocatedRows - 1, item.gridCol + item.gridW - 1);
      } else if (item.textCard) {
        writeKpiCard(worksheet, item.textCard, null, item.gridRow, item.gridCol, item.gridW, tracker, colWidths);
      } else if (item.kpiCard) {
        writeKpiCard(worksheet, item.kpiCard, item.vm, item.gridRow, item.gridCol, item.gridW, tracker, colWidths, item.splitCuts);
      } else if (item.isKPI) {
        writeKPICardStacked(worksheet, item.vm, item.gridRow, item.gridCol, tracker);
        colWidths[item.gridCol] = Math.max(colWidths[item.gridCol] || 0, 22);
        colWidths[item.gridCol + 1] = Math.max(colWidths[item.gridCol + 1] || 0, 18);
      } else {
        writeRegularTable(worksheet, item.vm, item.gridRow, item.gridCol, tracker, allTablesInfo, colWidths, exactWidths, item.visibleRows,
                          item.fieldCols);
      }
    }

    // ── the switches (findViewSwitches): Tableau's buttons become one dropdown of their values, on the shown value's
    //    button; the switch cell the charts read links to it
    const cellText = v => (v && typeof v === "object" ? (v.richText ? v.richText.map(t => t.text).join("") : v.text || "") : String(v ?? "")).trim();
    const switched = switches.filter(g => g.sw);
    switched.forEach(g => {
      const lower = g.sw.values.map(v => String(v).toLowerCase());
      let target = null;
      const others = [];
      g.buttons.forEach(button => {
        for (let rr = button.gridRow; rr < button.gridRow + Math.max(1, button.allocatedRows); rr++) {
          for (let cc = button.gridCol; cc < button.gridCol + Math.max(1, button.gridW); cc++) {
            const cell = worksheet.getCell(rr + 1, cc + 1);
            const t = cellText(cell.value).toLowerCase();
            if (!lower.includes(t)) continue;
            if (!target && t === String(g.value).toLowerCase()) target = cell;
            else others.push(cell);
          }
        }
      });
      if (!target) target = others.shift() || worksheet.getCell(g.buttons[0].gridRow + 1, g.buttons[0].gridCol + 1);
      others.forEach(c => { if (c !== target) c.value = null; });
      target.value = g.value;
      target.numFmt = '@" ▾"';                     // reads as a dropdown (the value itself stays the view's name)
      target.dataValidation = { type: "list", allowBlank: false, formulae: [`"${g.sw.values.map(v => String(v).replace(/"/g, "")).join(",")}"`],
                                showErrorMessage: true, errorTitle: "View", error: `Pick one of: ${g.sw.values.join(", ")}.` };
      target.note = `Switches the view, as Tableau's buttons do: ${g.views.map(v => `${v.value} → ${v.item.visualName || v.item.name}`).join("; ")}`;
      g.sw.cell.value = { formula: sheetRef(worksheet.name, target.address.replace(/([A-Z]+)(\d+)/, "$$$1$$$2")), result: g.value };
    });
    const switchedParams = new Set(switched.map(g => g.param));
    // the text boxes they switch (a view's title): each line that differs, a formula of the switch
    switched.forEach(g => g.texts.forEach(t => {
      const it = t.item, shownLines = t.lines[g.views.findIndex(v => v.value === g.value)];
      shownLines.forEach((text, li) => {
        const options = t.lines.map(lines => lines[li] || "");
        if (new Set(options).size < 2 || !text.trim()) return;
        for (let rr = it.gridRow; rr < it.gridRow + Math.max(1, it.allocatedRows); rr++) {
          for (let cc = it.gridCol; cc < it.gridCol + Math.max(1, it.gridW); cc++) {
            const cell = worksheet.getCell(rr + 1, cc + 1);
            if (cellText(cell.value) !== text.trim()) continue;
            const font = cell.value && cell.value.richText && cell.value.richText[0] ? cell.value.richText[0].font : cell.font;
            cell.value = { formula: g.sw.pick(options.map(xlText)), result: text };
            if (font) cell.font = font;
            return;
          }
        }
      });
    }));

    // ── the views the buttons switch to (withAlternateViews), below the dashboard: a header over them that links back
    //    up, and each button a link to the view it shows (Excel cannot switch them in place)
    if (altItems.length) {
      const link = (r, c) => `#${sheetRef(worksheet.name, `${getExcelColumnName(c)}${r + 1}`)}`;
      const left = Math.min(...altItems.map(it => it.gridCol)), top = Math.min(...altItems.map(it => it.gridRow));
      const head = worksheet.getCell(Math.max(dashEnd, top - 1) + 1, left + 1);     // the row left for it above them
      if (head.value === null || head.value === undefined || head.value === "") {
        head.value = { text: "Other views – in Tableau a button switches the view above to one of these   ↑ back to the dashboard", hyperlink: link(0, 0) };
        head.font = { bold: true, size: 10, color: { argb: "FF595959" } };
      }
      const dashM = fmtModel && fmtModel.dashboards[dashboard.name];
      const textOf = v => (v && typeof v === "object" ? (v.richText ? v.richText.map(t => t.text).join("") : v.text || "") : String(v ?? "")).trim();
      ((fmtModel && fmtModel.paramActions) || []).filter(a => !a.dashboard || a.dashboard === dashboard.name).forEach(a => {
        if (switchedParams.has(a.param)) return;
        const button = adjustedItems.find(it => it.name === a.worksheet);
        const ref = a.field ? tfParseFieldRef(a.field) : null;
        const info = ref ? tfFieldInfo(fmtModel, ref) : null;
        const lit = info && String(info.formula || "").trim().match(/^'((?:[^']|'')*)'$|^"((?:[^"]|"")*)"$/);
        if (!button || !lit || !dashM) return;
        const value = lit[1] !== undefined ? lit[1].replace(/''/g, "'") : lit[2].replace(/""/g, '"');
        // the zones this button's value shows, and where they are now (above if Tableau shows them, else below)
        const zones = new Set(Object.entries(dashM.toggles || {}).filter(([, t]) => t.param === a.param && t.value === value).map(([z]) => z));
        const views = adjustedItems.filter(it => zones.has(zoneOf(it)) && it.type === "worksheet");
        if (!views.length) return;
        const r = Math.min(...views.map(it => it.gridRow)), c = Math.min(...views.map(it => it.gridCol));
        for (let rr = button.gridRow; rr < button.gridRow + Math.max(1, button.allocatedRows); rr++) {
          for (let cc = button.gridCol; cc < button.gridCol + Math.max(1, button.gridW); cc++) {
            const cell = worksheet.getCell(rr + 1, cc + 1);
            if (textOf(cell.value).toLowerCase() !== value.toLowerCase()) continue;
            cell.value = { text: textOf(cell.value), hyperlink: link(r, c), tooltip: `Go to the ${views.map(v => v.visualName || v.name).join(", ")} view` };
            return;
          }
        }
      });
    }

    // ── how each visual was converted (NATIVE … TABLE_FALLBACK), and why: the report sheet, the panel status, and a
    //    note on every visual exported as its data so the reader of the sheet knows what it stands for
    /** @type {ReportVisual[]} */
    const reportVisuals = [];
    adjustedItems.filter(it => it.type === "worksheet" && it.visualModel).forEach(item => {
      const semantics = describeVisual(item.visualModel, fmtModel);
      const conversion = conversionOf(item.visualModel, semantics);
      reportVisuals.push({ worksheet: item.name, semantics, ...conversion });
      const status = visualStatuses.find(entry => entry.worksheet === item.name);
      if (status) Object.assign(status, { strategy: conversion.strategy, visual: semantics.visual, asTable: !!item.visualModel.crosstab });
      if (FORMAT_CONFIG.fallbackNotes && conversion.strategy === STRATEGY.TABLE_FALLBACK) {
        worksheet.getCell(item.gridRow + 1, item.gridCol + 1).note =
          `Tableau visual: ${semantics.visual}\nExported as: its data (TABLE_FALLBACK)\nWhy: ${conversion.reason}`;
      }
    });


    // a KPI card directly above / below a chart of the same width and colour (Tableau's number + trend
    // tile): the layout gap between them gets the card colour too, so they read as one card
    const sameColor = (a, b) => !!a && !!b && String(a).replace(/^#/, "").slice(-6).toUpperCase() === String(b).replace(/^#/, "").slice(-6).toUpperCase();
    adjustedItems.filter(it => it.kpiCard && it.kpiCard.background).forEach(card => {
      const color = card.kpiCard.background;
      const cardEnd = card.gridRow + card.allocatedRows;
      adjustedItems.forEach(other => {
        if (other === card || other.gridCol !== card.gridCol || Math.abs(other.gridW - card.gridW) > 1) return;
        const specs = other.visualModel && other.visualModel.renderer === "excel-chart" ? other.visualModel.chartSpecs : null;
        if (!sameColor(specs && specs.length ? specs[0].background : null, color)) return;
        const otherEnd = other.gridRow + other.allocatedRows;
        const [from, to] = other.gridRow >= cardEnd ? [cardEnd, other.gridRow] : otherEnd <= card.gridRow ? [otherEnd, card.gridRow] : [0, 0];
        if (to - from < 1 || to - from > 2) return;
        for (let rr = from; rr < to; rr++) {
          for (let c = card.gridCol; c < card.gridCol + card.gridW; c++) {
            const cell = worksheet.getCell(rr + 1, c + 1);
            cell.fill = tfExcelFill(color);
            if (c === card.gridCol) cell.border = { left: KPI_GUTTER };
            if (c === card.gridCol + card.gridW - 1) cell.border = { ...(cell.border || {}), right: KPI_GUTTER };
          }
        }
      });
    });

    // bordered layout containers (Layout pane → Border, e.g. a card around a column of sheets): their
    // outline around the blocks they hold
    const zoneById = new Map(dashZones.map(z => [String(z.id), z]));
    const within = (zoneId, containerId) => {
      for (let z = zoneById.get(String(zoneId)); z && z.parent; z = zoneById.get(String(z.parent))) if (String(z.parent) === containerId) return true;
      return false;
    };
    dashZones.filter(z => /^layout-/.test(z.type) && !z.hidden && z.style && z.style.borderStyle && z.style.borderStyle !== "none" &&
                          (z.style.borderWidth || 0) > 0).forEach(container => {
      const members = adjustedItems.filter(it => it.layout && it.layout.id !== undefined && within(it.layout.id, String(container.id)));
      if (!members.length) return;
      const top = Math.min(...members.map(m => m.gridRow)), left = Math.min(...members.map(m => m.gridCol));
      const bottom = Math.max(...members.map(m => m.gridRow + m.allocatedRows)) - 1;
      const right = Math.max(...members.map(m => m.gridCol + m.gridW)) - 1;
      /** @type {Partial<import("exceljs").Border>} */
      const side = { style: /** @type {import("exceljs").BorderStyle} */ (tfStrokeToBorder(container.style.borderWidth) || "thin"),
                     color: { argb: container.style.borderColor || "FFD4D4D4" } };
      const edge = (r, c, sides) => { const cell = worksheet.getCell(r + 1, c + 1); cell.border = { ...(cell.border || {}), ...sides }; };
      for (let c = left; c <= right; c++) { edge(top, c, { top: side }); edge(bottom, c, { bottom: side }); }
      for (let r = top; r <= bottom; r++) { edge(r, left, { left: side }); edge(r, right, { right: side }); }
    });

    // dashboard shading, container / backdrop backgrounds and each object's own: the cells around and under
    // the blocks take the colour Tableau shows there; cells a block coloured itself keep their colour
    const dashColor = tfDashboardShading(fmtModel, dashboard.name);
    const bgBlocks = adjustedItems.filter(it => it.layout && it.layout.id !== undefined).map(it => {
      const zone = zoneById.get(String(it.layout.id));
      const zoneBg = (zone && zone.style && zone.style.bgColor) || undefined;
      // a worksheet draws its shading over its zone: its colour, Tableau's white, or see-through
      const shading = it.type === "worksheet" && it.vm ? it.vm.fmt.sheetShading() : "none";
      const own = shading === "none" ? zoneBg : shading || "white";
      return { zoneId: String(it.layout.id), top: it.gridRow, left: it.gridCol,
               bottom: it.gridRow + it.allocatedRows - 1, right: it.gridCol + it.gridW - 1, own };
    });
    // the dashboard's extent in px (its objects), for the size of zones given in 0–100000 units
    const dashExt = (dashboard.objects || []).reduce((a, o) => o.position && o.size
      ? { w: Math.max(a.w, o.position.x + o.size.width), h: Math.max(a.h, o.position.y + o.size.height) } : a, { w: 0, h: 0 });
    if (bgBlocks.length && (dashColor || dashZones.some(z => z.style && z.style.bgColor))) {
      const used = tracker.toRange();
      const plan = backgroundPlan(dashZones, dashColor, bgBlocks,
        { top: Math.min(...bgBlocks.map(b => b.top)), left: 0, bottom: used.e.r, right: used.e.c },
        { w: (dashFmt && dashFmt.width) || dashExt.w || 1200, h: (dashFmt && dashFmt.height) || dashExt.h || 800 });
      for (let r = 0; r <= used.e.r; r++) {
        for (let c = 0; c <= used.e.c; c++) {
          const color = plan.colorAt(r, c);
          if (!color) continue;
          const cell = worksheet.getCell(r + 1, c + 1);
          const target = cell.isMerged ? cell.master : cell;
          if (target.fill && target.fill.type === "pattern" && target.fill.pattern !== "none") continue;
          target.style = { ...target.style, fill: tfExcelFill(color) };      // a style of its own, not shared
        }
      }
      // dividers and accent lines: borders along the block edges and through the gaps, never across a block
      plan.lines.forEach(ln => {
        /** @type {Partial<import("exceljs").Border>} */
        const side = { style: ln.style, color: { argb: ln.color } };
        for (let k = ln.from; k <= ln.to; k++) {
          const across = bgBlocks.some(b => ln.dir === "h" ? b.top < ln.at && ln.at <= b.bottom && k >= b.left && k <= b.right
                                                         : b.left < ln.at && ln.at <= b.right && k >= b.top && k <= b.bottom);
          if (across) continue;
          const cell = ln.dir === "h" ? worksheet.getCell(ln.at + 1, k + 1) : worksheet.getCell(k + 1, ln.at + 1);
          const key = ln.dir === "h" ? "top" : "left";
          if (cell.border && cell.border[key]) continue;                   // a table's own rule stays
          cell.style = { ...cell.style, border: { ...(cell.border || {}), [key]: side } };
        }
      });
    }

    // the grid's columns at their dashboard widths (Excel: px = 7 × stored width); content no longer widens them
    grid.colPx.forEach((px, c) => { exactWidths[c] = Math.max(0.1, Math.round(px / 7 * 100) / 100); });
    setColumnWidths(worksheet, colWidths, exactWidths);
    applyAutoFilters(worksheet, allTablesInfo);
    if (FORMAT_CONFIG.printFitToWidth) {
      // printing / Save as PDF: the whole dashboard, one page wide, oriented like the dashboard
      const used = tracker.toRange();
      worksheet.pageSetup = { ...worksheet.pageSetup, orientation: dashExt.w >= dashExt.h ? "landscape" : "portrait",
        fitToPage: true, fitToWidth: 1, fitToHeight: 0, horizontalCentered: true,
        margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 },
        printArea: `A$1:${getExcelColumnName(used.e.c)}$${used.e.r + 1}` };          // ExcelJS adds the column's $
    }

    // every row a definite 15 pt (20 px): without one, Excel derives the row height from the display's scaling
    // (e.g. 19 px at 125 %), while charts and pictures keep their size – they would run into the block below.
    // Rows with wrapped text keep Excel's fit.
    {
      const lastRow = tracker.toRange().e.r + 1;
      for (let r = 1; r <= lastRow; r++) {
        const row = worksheet.getRow(r);
        if (row.height || row.hidden) continue;
        let wraps = false;
        row.eachCell(cell => { if (cell.alignment && cell.alignment.wrapText) wraps = true; });
        if (!wraps) row.height = 15;
      }
      // a row taller than 20 px (a big-font text box line) pushes everything below it down: give the excess back
      // in the following rows that hold nothing (no value, border or merge, no chart or picture over them), so
      // blocks further down start where they do on the dashboard
      const covered = new Set();
      [...chartJobs.map(j => j.item), ...adjustedItems.filter(it => it.image)].forEach(it => {
        for (let r = it.gridRow; r < it.gridRow + (it.allocatedRows || 0); r++) covered.add(r);
      });
      overlayImages.forEach(img => {
        const host = adjustedItems.find(it => it.layout && String(it.layout.id) === img.host);
        if (host) for (let r = host.gridRow; r < host.gridRow + (host.allocatedRows || 0); r++) covered.add(r);
      });
      const holdsNothing = ri => {
        if (covered.has(ri)) return false;
        let used = false;
        worksheet.getRow(ri + 1).eachCell(cell => {
          const b = cell.border || {};
          if (cell.isMerged || (cell.value !== null && cell.value !== undefined && cell.value !== "") || b.top || b.bottom || b.left || b.right) used = true;
        });
        return !used;
      };
      const dashTop = Math.min(...adjustedItems.filter(it => it.layout).map(it => it.gridRow));
      let drift = 0;
      for (let ri = isFinite(dashTop) ? dashTop : lastRow; ri < lastRow; ri++) {
        const row = worksheet.getRow(ri + 1);
        if (row.hidden) continue;
        const px = row.height ? Math.round(row.height * 4 / 3) : EXCEL_ROW_PX;
        if (drift > 0 && px === EXCEL_ROW_PX && holdsNothing(ri)) {
          const give = Math.min(drift, EXCEL_ROW_PX - 1);
          row.height = Math.round((EXCEL_ROW_PX - give) * 0.75 * 4) / 4;
          drift -= give;
        } else drift += px - EXCEL_ROW_PX;
        if (drift < 0) drift = 0;
      }
    }
    // final column widths / row heights in px: charts and pictures are fitted to their blocks with them
    // a column's stored width already holds Excel's padding: 7 px per unit (Calibri 11), 64 px when unset
    const colPx = ci => { const w = worksheet.getColumn(ci + 1).width; return w ? Math.round(w * 7) : 64; };
    const rowPx = ri => {                            // collapsed rows take no space on screen
      const row = worksheet.getRow(ri + 1);
      if (row.hidden) return 0;
      return row.height ? Math.round(row.height * 4 / 3) : EXCEL_ROW_PX;
    };
    const blockPx = it => {
      let w = 0, h = 0;
      for (let c = it.gridCol; c < it.gridCol + it.gridW; c++) w += colPx(c);
      for (let r = it.gridRow; r < it.gridRow + it.allocatedRows; r++) h += rowPx(r);
      return { w, h };
    };
    /* a picture at (x, y) px from the block's top-left cell, w × h px, linked like its zone */
    const addPicture = async (image, col, row, x, y, w, h) => {
      const pic = await prepareImage(image.file, image.info, w, h);
      if (!pic) { console.warn(`[Images] ${image.zone.param}: cannot be drawn in this browser`); return false; }
      const id = workbook.addImage({ base64: pic.base64, extension: pic.extension });
      worksheet.addImage(id, /** @type {any} */ ({
        tl: nativeAnchor(col, row, x, y, colPx, rowPx), ext: { width: w, height: h }, editAs: "oneCell",
        ...(image.url ? { hyperlinks: { hyperlink: image.url, tooltip: image.url } } : {})
      }));
      return true;
    };
    let picturesAdded = 0;
    for (const item of adjustedItems.filter(it => it.image)) {
      const img = item.image, block = blockPx(item);
      // Fit Image / Center Image inside the block, never larger than Tableau draws it in its zone
      const fit = fitImage(img.info.width, img.info.height, block.w, block.h, img.zone.scaled, img.zone.centered);
      const cap = fitImage(img.info.width, img.info.height, img.widthPx, img.heightPx, img.zone.scaled, img.zone.centered);
      const w = Math.min(fit.w, cap.w), h = Math.min(fit.h, cap.h);
      const centered = img.zone.centered !== false;
      const x = centered ? Math.round((block.w - w) / 2) : 0, y = centered ? Math.round((block.h - h) / 2) : 0;
      if (await addPicture(img, item.gridCol, item.gridRow, x, y, w, h)) picturesAdded++;
    }
    for (const img of overlayImages) {
      // an icon over a sheet: the same spot of that sheet's block, scaled with the block
      const host = adjustedItems.find(it => it.layout && String(it.layout.id) === img.host);
      if (!host || !host.layout.widthPx || !host.layout.heightPx) continue;
      const block = blockPx(host);
      const sx = block.w / host.layout.widthPx, sy = block.h / host.layout.heightPx, k = Math.max(0.5, Math.min(1.5, sx, sy));
      const fit = fitImage(img.info.width, img.info.height, img.widthPx * k, img.heightPx * k, img.zone.scaled, img.zone.centered);
      if (await addPicture(img, host.gridCol, host.gridRow, Math.round((img.xPx - host.layout.xPx) * sx) + fit.x,
                           Math.round((img.yPx - host.layout.yPx) * sy) + fit.y, fit.w, fit.h)) picturesAdded++;
    }
    if (picturesAdded) console.log(`[Images] ${picturesAdded} picture(s) added`);

    // the Conversion Report's part for this dashboard (written once for all dashboards by exportToExcel)
    const objects = [
      ...textItems.map(it => ({ name: String(it.visualName || "").slice(0, 60) || it.name, kind: "Text box", output: "Cells, with Tableau's fonts" })),
      ...imageItems.map(it => ({ name: it.visualName, kind: iconSheets.some(s => s.name === it.name) ? "Button / icon sheet" : "Image",
                                 output: "Picture" })),
      ...pieLegendItems.map(it => ({ name: it.visualName, kind: "Colour legend", output: "Cells: a colour key per item" })),
      ...filterValueItems.map(it => ({ name: it.filterName, kind: it.isParameter ? "Parameter control" : "Filter card",
                                       output: it.isParameter ? (it.paramFields ? "Date dropdowns: the visuals follow them" : it.dateOnly ? "Date dropdowns (no visual on this sheet uses it)" : "Cells: its value")
                                         : it.slicer ? "Excel slicer (its selection as in Tableau)"
                                         : it.control ? controlOutput(it.control.field)
                                         : "Cells: the selected values" })),
      ...[...pieCentreSheets].map(name => ({ name, kind: "Worksheet in a donut's hole", output: "The donut's centre text" }))
    ];
    const report = {
      dashboard: dashboardName, workbookFile: formatModelFileName(), visuals: reportVisuals, objects, skipped,
      filters: Object.entries(filterValuesMap).map(([name, values]) => ({ name, values: values.map(String) })),
      parameters: reportParams,
      // which visuals follow the slicers, and why the others keep Tableau's values
      slicers: slicerJobs.length || controlCards ? dataWorksheetItems.filter(it => it.slicerNote).map(it => {
        const b = it.vm && it.vm.slicerBinding;
        if (!b) return { worksheet: it.name, follows: false, note: it.slicerNote };
        if (!b.written) return { worksheet: it.name, follows: false, note: `fixed values: this ${it.visualModel.type.toLowerCase()} is not drawn from values a formula can give` };
        const part = b.fixed ? `partly (${b.written} of ${b.written + b.fixed} values; the rest keep Tableau's numbers)` : "";
        return { worksheet: it.name, follows: true, note: part ? it.slicerNote.replace("follows the slicers", `follows the slicers ${part}`) : it.slicerNote };
      }) : []
    };
    report.slicers.forEach(s => console.log(`[Slicers] ${s.worksheet}: ${s.note}`));

    // the slicers at their blocks, now that the final column widths / row heights are known
    const finalizeSlicers = () => slicerJobs.map(item => {
      let w = 0, h = 0;
      for (let c = item.gridCol; c < item.gridCol + item.gridW; c++) w += colPx(c);
      for (let r = item.gridRow; r < item.gridRow + item.allocatedRows; r++) h += rowPx(r);
      const at = (x, y) => { const a = nativeAnchor(item.gridCol, item.gridRow, x, y, colPx, rowPx);
                             return { col: a.nativeCol, colOff: a.nativeColOff, row: a.nativeRow, rowOff: a.nativeRowOff }; };
      return { sheetName: worksheet.name, field: item.slicer.field, caption: item.filterName || item.slicer.field.header,
               columnCount: item.slicer.columnCount, from: at(0, 0), to: at(Math.max(60, w), Math.max(SLICER_HEADER_PX + SLICER_ITEM_PX, h)) };
    });

    // native charts, once ExcelJS has written the file: pies and donuts get Tableau's pane cells inside the block
    // (slices link back to the block's first cell); every chart is fitted to its block's final px size
    const finalizeCharts = () => {
      /** @type {ChartJob[]} */
      const pieChartJobs = [];
      pieJobs.forEach(({ item, top, panes, gauge, pie }) => {
        let w = 0, h = 0;
        for (let c = item.gridCol; c < item.gridCol + item.gridW; c++) w += colPx(c);
        for (let r = top; r < item.gridRow + item.allocatedRows; r++) h += rowPx(r);
        const bg = (item.vm.fmt.tableBackground() || "FFFFFFFF").slice(2);
        const link = `#'${worksheet.name.replace(/'/g, "''")}'!${worksheet.getCell(top + 1, item.gridCol + 1).address}`;
        (gauge ? buildGaugeCharts(item.visualModel.gauge, { w, h }, panes, { name: item.visualName || item.name, bg })
               : buildPieCharts(pie || item.visualModel.pie, { w, h }, panes, { name: item.visualName || item.name, bg, link })).forEach(c => {
          const at = nativeAnchor(item.gridCol, top, c.x, c.y, colPx, rowPx);
          pieChartJobs.push({ spec: /** @type {any} */ ({ kind: "pie" }), refs: /** @type {any} */ (null), item, name: c.name,
            col: at.nativeCol, row: at.nativeRow, colOffPx: at.nativeColOff / 9525, rowOffPx: at.nativeRowOff / 9525,
            widthPx: c.w, heightPx: c.h, xml: c.xml, shapes: c.shapes, tips: c.tips });
        });
      });
      if (chartJobs.length || pieChartJobs.length) {
        // fit every chart to its own block now that the final column widths/row heights are known:
        // tables widen the columns, so a chart kept at its Tableau pixel size would end short of the
        // tables / charts aligned with it on the dashboard → stretch it to the block's edges
        chartJobs.forEach(job => {
          let w = 0, h = 0;
          for (let c = job.item.gridCol; c < job.item.gridCol + job.item.gridW; c++) w += colPx(c);
          for (let r = job.top; r < job.item.gridRow + job.item.allocatedRows; r++) h += rowPx(r);
          if (job.spec.grid) {
            // small multiples: each pane at its place in Tableau's grid
            const f = gridFrame(job.spec.grid, job.spec, w, h);
            const at = nativeAnchor(job.item.gridCol, job.top, f.x, f.y, colPx, rowPx);
            Object.assign(job, { col: at.nativeCol, colOffPx: at.nativeColOff / 9525, row: at.nativeRow, rowOffPx: at.nativeRowOff / 9525,
                                 widthPx: Math.max(12, f.w), heightPx: Math.max(12, f.h) });
            return;
          }
          job.widthPx = Math.max(40, w);
          if (job.item.pairedCard) { job.colOffPx = 2; job.widthPx = Math.max(40, w - 4); }     // the card's white edges
          // stacked panes split the block's real height (rows may be taller, shorter or collapsed), each from its own spot
          const paneH = Math.floor(h / job.panes);
          let r = job.top, off = job.pane * paneH;
          for (let guard = 0; guard < 5000 && off >= rowPx(r) && (rowPx(r) > 0 || off > 0); guard++) { off -= rowPx(r); r++; }
          job.row = r;
          job.rowOffPx = off;
          job.heightPx = Math.max(20, job.pane === job.panes - 1 ? h - job.pane * paneH : paneH);
          // packed bubbles: Tableau's tooltip on each bubble (a click follows the link to the chart's own cell)
          const link = `#'${worksheet.name.replace(/'/g, "''")}'!${worksheet.getCell(job.row + 1, job.col + 1).address}`;
          const tips = ExcelChartWriter.bubbleTips(job.spec, job, link);
          if (tips.length) job.tips = tips;
        });
        return [...chartJobs, ...pieChartJobs];
      }
      return [];
    };

    return { worksheet, dashboardName, visualStatuses, workbookWarning, report, finalizeCharts, finalizeSlicers };
}


/* ═══ main.js ═════════════════════════════════════════════════════════════════════════════════════ */
/* Entry point: wires the panel once Tableau has initialised the extension –
 *   📁 Load workbook (manual .twb / .twbx), the loaded-workbook label (remembered workbook), the Tableau Cloud
 *   auto-load (backend, Advanced settings) and Export (waits for a running auto-load first). */

/* the loaded-workbook label at start-up: the workbook saved with the dashboard, else one remembered in this browser */
function restoreWorkbookLabel() {
  const fileLabel = document.getElementById("twb_file_label");
  if (!fileLabel) return Promise.resolve();
  const savedFileName = tableau.extensions.settings.get("twbFileName");
  const savedTitleMap = tableau.extensions.settings.get("twbTitleMap");
  const savedFormat = tableau.extensions.settings.get("twbFormatModel");
  if (savedFileName && savedTitleMap) {
    const titleCount = Object.keys(JSON.parse(savedTitleMap)).length;
    if (savedFormat) setFormatModelCache(JSON.parse(savedFormat), savedFileName);
    return ensureFormatModel().then(model => {
      if (model) {
        showWorkbookLabel(savedFileName, `${titleCount} titles, formatting for ${Object.keys(model.sheets || {}).length} sheets`, model);
      } else {
        fileLabel.textContent = `${savedFileName} — formatting not available after reload.
⚠ Click 📁 Load Workbook again so charts are exported as charts`;
        fileLabel.classList.add("status-warning");
      }
    });
  }
  fileLabel.textContent = "No workbook loaded — click 📁 to load titles and colors";
  return ensureFormatModel().then(model => {
    if (model && RESTORED_FILE_NAME) {
      showWorkbookLabel(RESTORED_FILE_NAME, `remembered from an earlier load, formatting for ${Object.keys(model.sheets || {}).length} sheets`, model);
    } else {
      fileLabel.textContent = "⚠ No workbook loaded — click 📁 Load Workbook, otherwise charts are exported as tables";
      fileLabel.classList.add("status-warning");
    }
  });
}

document.addEventListener("DOMContentLoaded", () => {
  tableau.extensions.initializeAsync({ configure: () => openLicenceSetup() }).then(() => {
    const dashboard = tableau.extensions.dashboardContent.dashboard;
    /** the export waits on this while the workbook is being loaded automatically */
    let WORKBOOK_READY = Promise.resolve();

    const loadBtn = /** @type {HTMLButtonElement} */ (document.getElementById("load_workbook_btn"));
    if (loadBtn) {
      loadBtn.addEventListener("click", async () => {
        loadBtn.disabled = true;
        loadBtn.textContent = "⏳ Loading...";
        const before = FORMAT_MODEL_CACHE;
        try { await loadWorkbookFile(); } catch (e) { console.warn("[Workbook] manual load failed:", e); }
        // a workbook loaded by hand wins: close an open "pick a workbook" prompt, stop a running auto-load
        if (FORMAT_MODEL_CACHE && FORMAT_MODEL_CACHE !== before) {
          WORKBOOK_LOAD_GEN++;
          cancelWorkbookPicker();
          setWorkbookPhase("ready");
        }
        loadBtn.disabled = false;
        loadBtn.textContent = "📁 Load workbook file manually (.twb / .twbx)";
      });
    }

    const exportBtn = /** @type {HTMLButtonElement} */ (document.getElementById("export_button"));
    if (exportBtn) EXPORT_LABEL = exportBtn.textContent;

    // the remembered workbook first, then Phase 1: fetch the workbook from Tableau Cloud automatically
    // (falls back to 📁 on any problem; Tableau Desktop keeps the manual load). Export stays disabled meanwhile.
    WORKBOOK_READY = trackWorkbookLoad(() => restoreWorkbookLabel()
      .catch(err => console.warn("[Workbook] could not restore the remembered workbook:", err))
      .then(() => autoLoadWorkbook(dashboard)));
    setupBackendSettings(dashboard, start => (WORKBOOK_READY = trackWorkbookLoad(start)));
    setupLicence(dashboard, start => (WORKBOOK_READY = trackWorkbookLoad(start)));
    setupVisualsAsTables();

    if (exportBtn) {
      exportBtn.addEventListener("click", async () => {
        // only a ready workbook, and never twice at once (the button is disabled then anyway)
        if (WORKBOOK_PHASE !== "ready" || EXPORT_RUNNING) return;
        EXPORT_RUNNING = true;
        try { await exportToExcel(); }
        finally { EXPORT_RUNNING = false; refreshExportButton(); }
      });
    }

  }).catch((err) => {
    console.error("[Tableau init]", err);
    alert("Failed to connect to Tableau: " + err.message);
    if (document.getElementById("export_button")) {
      /** @type {HTMLButtonElement} */ (document.getElementById("export_button")).disabled = true;
    }
  });
});

})();
