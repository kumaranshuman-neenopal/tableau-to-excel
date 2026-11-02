#!/usr/bin/env node
"use strict";
/* node tools/stamp.js [--check] [--version x.y.z]
 *
 * The build stamp of tableau-to-excel.js – EXTENSION_BUILD, shown in the panel and every Conversion Report – is the
 * date and a hash of the code itself, so a file made by an older (cached) script is told at a glance. The same hash
 * is the ?v= of the script in index.html, so browsers load each new build instead of a cached one.
 *   (no flag)        writes the stamp and the ?v= when the code changed (the pre-commit hook runs this)
 *   --check          fails when they do not match the code (CI)
 *   --version x.y.z  also sets EXTENSION_VERSION and Export.trex's extension-version (a release)
 */
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const ROOT = path.resolve(__dirname, "..");
const JS = path.join(ROOT, "tableau-to-excel.js");
const HTML = path.join(ROOT, "index.html");
const TREX = path.join(ROOT, "Export.trex");
const BUILD_LINE = /^const EXTENSION_BUILD = "([^"]*)";$/m;
const VERSION_LINE = /^const EXTENSION_VERSION = "([^"]*)";$/m;
const SCRIPT_TAG = /<script src="\.\/tableau-to-excel\.js(?:\?v=([^"]*))?"><\/script>/;

const args = process.argv.slice(2);
const check = args.includes("--check");
const vi = args.indexOf("--version");
const newVersion = vi >= 0 ? args[vi + 1] : null;
if (vi >= 0 && !/^\d+\.\d+\.\d+$/.test(newVersion || "")) { console.error("--version needs x.y.z"); process.exit(2); }

let js = fs.readFileSync(JS, "utf8");
let html = fs.readFileSync(HTML, "utf8");
if (!BUILD_LINE.test(js) || !VERSION_LINE.test(js)) { console.error("EXTENSION_BUILD / EXTENSION_VERSION not found in tableau-to-excel.js"); process.exit(2); }
if (!SCRIPT_TAG.test(html)) { console.error("the tableau-to-excel.js script tag was not found in index.html"); process.exit(2); }

if (newVersion) js = js.replace(VERSION_LINE, `const EXTENSION_VERSION = "${newVersion}";`);
// the hash of the code without its own stamp (writing the stamp does not change it); line endings normalised
const hash = crypto.createHash("sha256").update(js.replace(BUILD_LINE, "").replace(/\r\n/g, "\n")).digest("hex").slice(0, 8);
const current = js.match(BUILD_LINE)[1];
const tagged = (html.match(SCRIPT_TAG) || [])[1] || "";
const stamped = current.endsWith(` ${hash}`) && tagged === hash;

if (check) {
  if (stamped) { console.log(`build stamp ok: ${current}`); process.exit(0); }
  console.error(`build stamp is stale: EXTENSION_BUILD "${current}", index.html ?v=${tagged || "(none)"}, code hash ${hash}` +
                " – run node tools/stamp.js (the pre-commit hook does)");
  process.exit(1);
}

if (!stamped || newVersion) {
  const date = new Date().toISOString().slice(0, 10);
  js = js.replace(BUILD_LINE, `const EXTENSION_BUILD = "${date} ${hash}";`);
  html = html.replace(SCRIPT_TAG, `<script src="./tableau-to-excel.js?v=${hash}"></script>`);
  fs.writeFileSync(JS, js);
  fs.writeFileSync(HTML, html);
  console.log(`build stamp: ${date} ${hash}`);
} else console.log(`build stamp unchanged: ${current}`);

if (newVersion) {
  const trex = fs.readFileSync(TREX, "utf8").replace(/extension-version="[^"]*"/, `extension-version="${newVersion}"`);
  fs.writeFileSync(TREX, trex);
  console.log(`version ${newVersion} (tableau-to-excel.js, Export.trex)`);
}
