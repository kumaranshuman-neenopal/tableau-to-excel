# Mark2Table / Export2Sheet

> Tableau dashboard extension that converts dashboards into editable, formatted Excel workbooks.

[![Repository](https://img.shields.io/badge/GitHub-kumaranshuman--sudo-181717?style=flat-square&logo=github)](https://github.com/kumaranshuman-sudo/tableau-to-excel)
[![Runtime](https://img.shields.io/badge/Runtime-Tableau%20Desktop-1C1C1C?style=flat-square)](https://www.tableau.com/products/desktop)
[![Build](https://img.shields.io/badge/Build-none%20(one%20file)-FFCF00?style=flat-square)](#repository-layout)

## What it solves

Tableau dashboards are optimized for interactive analysis, while Excel is often still required for downstream editing, reporting, and distribution. Mark2Table bridges that gap by exporting dashboard content into a workbook while preserving as much semantic and visual structure as possible: each object at its dashboard position, with Tableau's fonts, colours and number formats.

The exporter recognises what each visual is from its shelves, marks and encodings, and picks the closest representation that keeps its meaning, in this order:

| Strategy | Excel output | Examples |
|---|---|---|
| **NATIVE** | a genuine Excel equivalent | text tables, KPIs, bar / column / line / area (stacked, 100 % stacked), pie, donut, scatter, bubble, combo and dual axis, treemap |
| **CONSTRUCTED** | Excel primitives that keep the meaning | heat map / highlight table / calendar (a matrix with Tableau's colour per cell), waterfall, box plot, Gantt, funnel, bullet graph, butterfly, dumbbell, lollipop, horizontal dot plot, histogram of Tableau's bins, tables built from marks |
| **APPROXIMATE** | Excel's closest chart, the difference named | packed bubbles (Excel's bubble chart), bars Tableau sizes by a measure, bars summing Tableau's detail segments |
| **TABLE_FALLBACK** | the visual's data as a table | symbol / filled / density / polygon / flow maps, map layers, viz extensions, anything no Excel chart would show correctly |
| **IMAGE_FALLBACK** | a Tableau-rendered picture, its data on a *Visual Data* sheet | charts Tableau can draw when the native chart is impossible |

A correct table always beats a misleading chart: a visual is never drawn as an Excel chart that merely looks similar. Every export adds a **Conversion Report** sheet listing each visual, the Tableau visual it is, the strategy, the Excel output, the fidelity, the reason and the evidence it was recognised from, plus the filters and parameters the data reflects. A visual exported as data carries a cell note with the same reason.

Dashboard formatting – palettes, number formats, layout and workbook styles – is applied throughout.

On Tableau Cloud the workbook is loaded automatically through a small backend (Worker) and every visible dashboard of the workbook can be exported in one go, one Excel sheet per dashboard.

## Architecture

```text
Tableau Dashboard
       │
       ▼
Tableau Extensions API
       │
       ├── Workbook / TWB / TWBX parser
       │        └── formatting cascade + assets
       │
       ├── Dashboard data extraction
       │
       ▼
View model (pivot, sort, matrix) ─── visual classification (type + evidence)
       │
       ▼
Semantics (Tableau visual, encodings) ─── conversion strategy (NATIVE … TABLE_FALLBACK)
       │
       ▼
Renderer-neutral chart model
       │
       ├── Cell writer ───────────────► XLSX cells
       ├── Chart writer ──────────────► Native Excel charts
       ├── Image/data fallback ───────► XLSX pictures / data tables
       └── Report writer ─────────────► Conversion Report sheet
```

## Engineering highlights

- **TWB/TWBX parsing:** reads workbook metadata and formatting rules rather than treating the dashboard as a screenshot.
- **Renderer-neutral chart model:** separates Tableau visual interpretation from Excel DrawingML generation.
- **Native chart generation:** supported visuals become editable Excel charts.
- **Formatting pipeline:** palettes, number formats, dimensions and workbook styles are resolved before writing.
- **Tableau Cloud auto-load:** the backend resolves which published workbook holds the dashboard and serves its XML; the Personal Access Token stays on the backend.
- **Every dashboard in one export:** the other dashboards are read through hidden Embedding API views in the viewer's own Tableau session (row-level security applies), one Excel sheet each.
- **No build step:** the whole extension is one script, `build_table_copy.js`, loaded by `index.html` next to ExcelJS and JSZip.

## Run it in Tableau Desktop

Serve the repository folder on `http://localhost:5500` – the URL in `Export.trex` – with any static server (for example VS Code **Live Server**, or `npx http-server -p 5500 -c-1`). In Tableau, add an Extension object to a dashboard and choose **Access Local Extensions → Export.trex** (or **Reload** an existing one).

In the panel, **📁 Load workbook file manually** reads the dashboard's `.twbx` (or `.twb`) for its formatting. Load the packaged `.twbx` to also export logos, custom shapes and other images: a `.twb` does not contain them. **All dashboards** (ticked by default) exports every visible dashboard of the workbook, one sheet each; untick it to export only this dashboard.

## Tableau Cloud

1. Host the backend (Worker) that holds the Tableau Personal Access Token: `/resolve` names the published workbook a dashboard belongs to, `/twb/<id>` returns its XML.
2. Set `DEFAULT_BACKEND_URL` in `build_table_copy.js` (section `cloud/backend.js`), or enter the URL and key under **Advanced: backend connection** in the panel – they are saved with the workbook.
3. The workbook then loads automatically when the extension opens; the 📁 button stays as the fallback. The first multi-dashboard export may ask the viewer to sign in to Tableau once.

## Deploy

1. Host the repository files (`index.html`, `build_table_copy.js`, `js/`) on any HTTPS web server (Tableau requires HTTPS for non-localhost extensions). ExcelJS 4.4 and JSZip 3.10 load from `js/` – no outside website is needed.
2. In `Export.trex`, set `<source-location><url>` to the hosted `index.html`, and raise `extension-version` for each release.
3. Tableau Server / Cloud: add the URL to the extension allow list (Settings → Extensions).
4. Hand out the updated `Export.trex`; users add it through **Access Local Extensions**.

## How an export runs

Everything below lives in `build_table_copy.js`; each `═══ <path> ═══` banner there starts the module named here. `export/export.js` → `exportToExcel()` drives one export, calling `writeDashboardSheet()` once per dashboard:

1. **Read** – every worksheet's summary data from Tableau (`export/sheet-data.js`) and the dashboard objects' positions; the loaded workbook's format model (`twb/parser.js`, cached by `ui/workbook-store.js`).
2. **Model** – per worksheet a view model (`visual/view-model.js`: pivot, sort, visible columns, headers), a visual type with the evidence for it (`visual/classify.js`), heat maps and crosstabs pivoted into Tableau's matrix (`visual/matrix.js`), then a renderer (`visual/visual-model.js`): cells, a native chart, an image or a data table. Summary columns are matched to the workbook's fields by Tableau's field id (`fieldId`), else by caption.
3. **Lay out** – every object at its dashboard position: columns from the zone edges, rows at 20 px (`export/layout.js`).
4. **Write** – tables, KPI cards, text, filters and pictures as cells (`export/*-writers.js`, `export/kpi-card.js`, `export/mark-cells.js`), backgrounds (`export/backgrounds.js`), then native charts injected into the XLSX package (`charts/writer/`).

5. **Report** – each visual's semantics (`visual/semantics.js`: "Bump Chart", "Filled Map" …, its shelves, encodings and calculated fields) and conversion (`visual/strategy.js`) go to the Conversion Report sheet (`export/report.js`) and the panel's status line.

Pies and donuts have their own path (`charts/model/pie.js` → `charts/writer/pie.js`): Tableau's pane cells and mark size, labels as text boxes at Tableau's label positions, the donut hole with its total, nested donuts, one pie per pane, a colour legend at its dashboard position, and Tableau's tooltip on each slice – an invisible wedge whose hyperlink ScreenTip is the tooltip text.

What the chart model reads from the workbook, beyond shelves and marks: Dual Axis and Synchronize Axis (`fold` / `synchronized` on the axis), reversed axes (bump charts), constants that only place labels or bar starts (`MIN(0)`), negated measures (butterfly charts), per-cell reference lines (bullet targets), "% of Total" calculations (100 % stacked), Detail / Label fields that split marks (one line per member, never their sum), and pane boundaries (Tableau never draws a line across panes).

Adding support for a new kind of visual usually means a rule in `visual/classify.js` and either a chart spec builder in `charts/model/` or a cell writer in `export/`; a construction or approximation names itself in `spec.conversion` so the report says what it is.

Settings that change the output are in the `config.js` section (`FORMAT_CONFIG`): `conversionReport` (the report sheet), `fallbackNotes` (notes on visuals exported as data), `chartPolicy`, `slicers` / `slicerMaxRows` and more; `debug: true` logs the parsed workbook model and formatting traces to the browser console.

### Slicers (`export/slicers.js`)

A dashboard filter card becomes an Excel control where the card was, with Tableau's selection, and the charts, tables and KPI values of the same data source **cross-filter** with it:

| Tableau card | Excel |
|---|---|
| Multiple values (dropdown, list) of a field (Region, YEAR(Order Date) …) | a **slicer** |
| Single value (dropdown, list, slider) | a **dropdown cell** (Data Validation list: (All) and the field's values) |
| Range of dates (slider, range of dates) | **From** / **To** date cells (only dates within the data's) |
| Start / End date parameters that a calculated filter compares a date with (`[Order Date] >= [Start Date] AND [Order Date] <= [End Date]`, also inside `IF … THEN TRUE ELSE FALSE END`) | each parameter's **Year / Month / Day dropdowns** (Excel has no pop-up calendar without macros) and the date they make – the visuals follow it; a day past the month's end is its last day |

A dropdown or date range filters only the visuals Tableau filters by it: such a visual reads its own `Visible: <sheet>` column. Relative-date filters and ranges of measures (they filter marks, not rows) stay text cards.

1. The data source's rows (one logical table, up to `slicerMaxRows`) go to a hidden **Slicer Data** sheet as an Excel table: a `Visible` column (`SUBTOTAL(103)` – 0 on a row the slicers hide – and 0 as well when a dropdown cell or the From / To cells leave the row out; the dropdowns' lists are on a hidden **Slicer Lists** sheet), the slicer fields, the measure fields the formulas need, and per visual a key column – the mark each row belongs to, found by the visual's dimensions under the sheet's other filters (a date range, a categorical filter).
2. Each measure becomes a formula over the visible rows of its mark: `SUMIFS`, `AVERAGEIFS`, `MINIFS`, `MAXIFS`, `COUNTIFS`, a distinct count (`UNIQUE` / `FILTER`, Excel 2021 / 365), or a calculation of those (`SUM([Profit])/SUM([Sales])`). Chart values (`tvSum` keeps each value's marks), table cells and KPI values are written as these formulas, with Tableau's number as the cached result; a mark the slicers leave empty is a gap in a chart (`#N/A`) and a blank cell in a table.
3. Before a visual is bound, the formulas are worked out on the rows with Tableau's selection and must give Tableau's own numbers. A visual that does not – a calculated or LOD dimension the rows lack, a table calculation, a filter that cannot be told – keeps its fixed values; so do nested donuts, gauges, packed bubbles and pictures. The Conversion Report lists which visuals follow the slicers and why the others do not.
   A pie or donut that follows the slicers (`bindPie()`) has a formula for each slice's value and for its label: the label is rebuilt from Tableau's label text, with each value (and % of total) as `TEXT()` in Tableau's number format, and Excel draws it from that cell ("Value From Cells"), so it moves with its slice. A donut is drawn as a pie under a circle in the hole's colour (Excel puts a doughnut's labels on the ring); a centre line showing the slices' total links to a cell. Tableau's tooltips are left out of such a pie, and a text box or sheet floating over its centre keeps Tableau's numbers.
4. `injectSlicers()` adds the table, slicer caches and slicers to the package as Excel writes them for a table slicer; one slicer cache per field is shared by every dashboard sheet of the file.

Axes of bound charts scale with the data (only the workbook's own fixed range stays). Charts and tables keep the bars, points and rows Tableau drew: a value the filter left out at export adds to those, but a category or row only it has cannot appear – export with every value selected to have them all. Tableau's filter actions (click a mark to filter) have no Excel counterpart.

## Repository layout

```text
index.html                extension panel: loads the Tableau library, ExcelJS, JSZip and build_table_copy.js
build_table_copy.js       the whole extension, one section per module:
                            format/ config util twb/   parsing and the formatting cascade
                            data/ visual/              view model, matrix pivot, classification, semantics, strategy
                            charts/model/ writer/      renderer-neutral chart specs → DrawingML / chartex parts
                            export/                    layout, workbook writing, slicers, Conversion Report, Visual Data sheet
                            ui/ cloud/                 panel status, workbook storage, backend auto-load, other dashboards
                            main.js                    panel wiring
js/                       Tableau Extensions API library, ExcelJS 4.4, JSZip 3.10
Export.trex               extension manifest (points Tableau at the hosted index.html)
.github/workflows/ci.yml  CI: syntax check of build_table_copy.js, page and manifest checks
```

## Current scope and limitations

This is a **Tableau dashboard extension**: it runs inside a dashboard in Tableau Desktop (or Tableau Server / Cloud once allowed there), not as a standalone exporter. End-to-end validation therefore requires Tableau and representative workbooks.

Visual fidelity follows the strategies above; the Conversion Report names every construction, approximation and fallback. Known limits:

- **Maps** of every kind are exported as tables of their marks (location fields and values): Excel's map charts cannot reproduce Tableau's geography, layers or colour scales reliably.
- **Calculations** – calculated fields, LOD expressions, table calculations – arrive as Tableau's computed values. Only the values of visuals that follow a slicer become Excel formulas (sums, averages, min / max, counts and calculations of those); the report lists each calculation and its kind.
- **Filters and parameters**: a filter card of a field's values becomes a slicer (several values) or a dropdown cell (one value), and a range of dates From / To cells, that cross-filter the visuals of its data source (see *Slicers*); relative-date and measure filters and parameters apply as they were at export time, and the report lists them. Excel has no live link back to Tableau.
- **Tooltips** appear on hover for pies, donuts (not when they follow a slicer) and packed bubbles; elsewhere their fields are in the data only. Actions, highlighting and drill-down are not reproduced.
- **Shapes and sizes** Excel cannot draw (custom shape images in charts, bar widths by a measure, rounded bar ends) are left out and named in the report.

The project should be evaluated against representative dashboards containing the visual types and formatting features it claims to support.

## Development

```bash
node --check build_table_copy.js
```

For regression testing, use representative Tableau workbooks and verify:

1. dashboard dimensions and object ordering;
2. extracted values and filters;
3. number formats and palettes;
4. chart type and series mapping;
5. native Excel chart editability;
6. fallback behaviour for unsupported visuals;
7. the Conversion Report: every visual's Tableau visual, strategy and reason as expected.

## License

See the repository license for details.
