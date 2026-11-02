"use strict";
/* node tests/lib/run-gallery.js <out.xlsx>  –  the "Visual Gallery" test workbook (fixtures/gallery.twb: every chart
 * type on one dashboard, with Region and YEAR(Order Date) filter cards) through the real exportToExcel(), against a
 * mock of the Tableau API (gallery.js). Prints the export's [Slicers] / [Filters] lines and "wrote <file>". */
const fs = require("fs");
const path = require("path");
const { load } = require("./load.js");

const OUT = path.resolve(process.argv[2] || "gallery.xlsx");

(async () => {
  const log = console.log.bind(console);
  console.log = (...a) => { if (/^\[(Slicers|Filters|Actions|States|Every value)\]/.test(String(a[0]))) log(...a); };
  console.info = console.table = () => {};
  console.warn = (...a) => log("WARN", ...a);
  let saved = null;
  const T = load({ dashboard: {}, captureDownload: bytes => { saved = bytes; }, expose: ["applyWorkbookXml", "exportToExcel"] });
  const g = require("./gallery.js");
  const { formatModel } = await T.applyWorkbookXml(g.workbookXml(), "Visual Gallery.twb", false, null);
  tableau.extensions.dashboardContent.dashboard = g.dashboardFrom(formatModel);
  await T.exportToExcel();
  for (let i = 0; i < 100 && !saved; i++) await new Promise(r => setTimeout(r, 50));
  if (!saved) { log("NO FILE – " + document.getElementById("export_status").textContent); process.exit(1); }
  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, saved);
  log(`wrote ${OUT}`);
})().catch(e => { console.error(e); process.exit(1); });
