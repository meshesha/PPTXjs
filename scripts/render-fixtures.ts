// Renders every fixture with the v2 engine (Node, no DOM) and writes viewable
// HTML pages to test/output/. Prints slide counts, warnings and timings.
//
// usage: npm run fixtures [-- name-filter]
import fs from "node:fs";
import path from "node:path";
import { PptxPackage } from "../src/core/package";
import { renderPresentation } from "../src/render/presentation";
import { BASE_CSS } from "../src/viewer/styles";

const root = process.cwd();
const fixturesDir = path.join(root, "test", "fixtures");
const outDir = path.join(root, "test", "output");
fs.mkdirSync(outDir, { recursive: true });
const filter = process.argv[2];

const files: string[] = [];
for (const e of fs.readdirSync(fixturesDir, { recursive: true })) {
  const f = String(e);
  if (f.endsWith(".pptx") && (!filter || f.includes(filter))) files.push(path.join(fixturesDir, f));
}

const summary: Record<string, unknown> = {};
for (const file of files.sort()) {
  const name = path.basename(file, ".pptx");
  const warnings: string[] = [];
  const t0 = Date.now();
  try {
    const pkg = PptxPackage.fromBytes(fs.readFileSync(file));
    const stackOf = (c: unknown) =>
      process.env.STACK && c instanceof Error ? " :: " + String(c.stack).split("\n").slice(1, 4).join(" <- ") : "";
    const res = renderPresentation(pkg, { onWarning: (w) => warnings.push(w.message + stackOf(w.cause)) });
    const failed = res.slides.filter((s) => s.error);
    const html =
      `<!doctype html><meta charset="utf-8"><title>${name}</title><style>${BASE_CSS}\n${res.css}\n` +
      `body{background:#e5e5e5;margin:0;padding:20px} .slide{margin:0 auto 20px;box-shadow:0 2px 8px rgba(0,0,0,.25)}</style>` +
      `<div class="pptxjs">${res.slides.map((s) => s.html).join("\n")}</div>`;
    fs.writeFileSync(path.join(outDir, name + ".html"), html);
    summary[name] = { slides: res.slides.length, failed: failed.length, warnings: warnings.length, ms: Date.now() - t0 };
    console.log(`${name.padEnd(12)} ${String(res.slides.length).padStart(3)} slides  ${failed.length} failed  ${warnings.length} warnings  ${Date.now() - t0}ms`);
    for (const w of [...new Set(warnings)].slice(0, 4)) console.log("    ", w.slice(0, 240));
  } catch (e) {
    console.log(`${name.padEnd(12)} FATAL ${(e as Error).stack?.split("\n").slice(0, 3).join(" | ")}`);
    summary[name] = { fatal: String(e) };
  }
}
fs.writeFileSync(path.join(outDir, "summary.json"), JSON.stringify(summary, null, 2));
