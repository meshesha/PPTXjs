// Renders every fixture with the legacy (v1, jQuery) PPTXjs inside jsdom and
// stores the produced HTML + any errors in test/baseline/. The v2 port is
// compared against these files to catch regressions introduced by the rewrite.
//
// usage: node scripts/legacy-baseline.mjs [fixture.pptx ...]
import { JSDOM, VirtualConsole } from "jsdom";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const legacy = (f) => fs.readFileSync(path.join(root, "legacy", f), "utf8");
const outDir = path.join(root, "test", "baseline");
fs.mkdirSync(outDir, { recursive: true });

function listFixtures() {
  const dir = path.join(root, "test", "fixtures");
  const out = [];
  for (const e of fs.readdirSync(dir, { recursive: true })) {
    if (String(e).endsWith(".pptx")) out.push(path.join(dir, String(e)));
  }
  return out;
}

async function renderLegacy(file) {
  const errors = [];
  const vc = new VirtualConsole();
  vc.on("jsdomError", (e) => errors.push(String(e.detail?.stack || e.message || e).split("\n").slice(0, 3).join(" | ")));
  vc.on("error", (...a) => errors.push(a.map(String).join(" ").slice(0, 300)));
  const dom = new JSDOM(`<!doctype html><html><body><div id="result"></div></body></html>`, {
    runScripts: "outside-only",
    pretendToBeVisual: true,
    virtualConsole: vc,
  });
  const w = dom.window;
  // jsdom has no object URLs; v1 uses them for video/audio.
  w.URL.createObjectURL = () => "blob:stub";
  for (const f of ["js/jquery-1.11.3.min.js", "js/jszip.min.js", "js/filereader.js", "js/dingbat.js", "js/pptxjs.js"]) {
    let src = legacy(f);
    if (f === "js/pptxjs.js") {
      // v1 defines Object.prototype.set (issues #22/#28/#30), which breaks every
      // later Object.defineProperty call and crashes V8 under jsdom. Swap it for
      // a plain helper with identical semantics so a baseline can be produced.
      src = src
        .replace("Object.prototype.set = function (parts, value) {", "var __set = function (obj, parts, value) {")
        .replace("var obj = this;", "")
        .replace("node.set(path, value)", "__set(node, path, value)");
    }
    w.eval(src);
  }
  const buf = fs.readFileSync(file);
  // Allocate in the jsdom realm: JSZip v2 detects input type with instanceof.
  const ab = new w.Uint8Array(buf).buffer;
  // Deliver the bytes on a later tick so v1's try/catch does not swallow render errors.
  let finished = false;
  w.JSZipUtils.getBinaryContent = (_url, cb) =>
    setTimeout(() => {
      try {
        cb(null, ab);
      } catch (e) {
        errors.push(String(e.stack || e).split("\n").slice(0, 3).join(" | "));
      }
      finished = true;
    }, 0);
  // Call the load handler synchronously so exceptions surface in the catch above.
  w.FileReaderJS.setupBlob = (_blob, opts) => opts.on.load({ target: { result: ab } });
  w.onerror = (msg, _s, line, _c, err) => {
    errors.push(String(err?.stack || msg).split("\n").slice(0, 3).join(" | "));
  };
  const started = Date.now();
  try {
    w.eval(`$("#result").pptxToHtml({ pptxFileUrl: "x.pptx", slideMode: false, keyBoardShortCut: false });`);
  } catch (e) {
    errors.push(String(e.stack || e).split("\n").slice(0, 3).join(" | "));
  }
  while (!finished) await new Promise((r) => setTimeout(r, 20));
  const html = w.document.getElementById("result").innerHTML;
  const slides = w.document.querySelectorAll("#result .slide").length;
  w.close();
  return { html, slides, errors, ms: Date.now() - started };
}

const files = process.argv.slice(2).length ? process.argv.slice(2) : listFixtures();
const summary = {};
for (const f of files) {
  const name = path.basename(f, ".pptx");
  process.stdout.write(`${name} ... `);
  const r = await renderLegacy(f);
  fs.writeFileSync(path.join(outDir, `${name}.legacy.html`), r.html);
  summary[name] = { slides: r.slides, errors: r.errors, ms: r.ms, bytes: r.html.length };
  console.log(`${r.slides} slides, ${r.errors.length} errors, ${r.ms}ms`);
  for (const e of r.errors.slice(0, 2)) console.log("   ", e.slice(0, 220));
}
fs.writeFileSync(path.join(outDir, "legacy-summary.json"), JSON.stringify(summary, null, 2));
