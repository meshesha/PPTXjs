// Structural comparison between v1 baseline HTML (test/baseline/*.legacy.html)
// and v2 output (test/output/*.html): per slide it compares the number of SVG
// shapes, images, text blocks and tables and the visible text.
//
// usage: node scripts/compare-baseline.mjs [name-filter]
import { JSDOM } from "jsdom";
import fs from "node:fs";
import path from "node:path";

const root = process.cwd();
const filter = process.argv[2];
const baseDir = path.join(root, "test", "baseline");
const outDir = path.join(root, "test", "output");

function stats(html) {
  const doc = new JSDOM(html).window.document;
  return [...doc.querySelectorAll(".slide")]
    .filter((s) => !s.parentElement.closest(".slide"))
    .map((s) => ({
      svg: s.querySelectorAll("svg.drawing").length,
      img: s.querySelectorAll("img").length,
      block: s.querySelectorAll("div.block").length,
      table: s.querySelectorAll("table").length,
      text: s.textContent.replace(/\s+/g, " ").trim(),
    }));
}

let total = 0,
  diffs = 0;
for (const f of fs.readdirSync(baseDir).filter((f) => f.endsWith(".legacy.html"))) {
  const name = f.replace(".legacy.html", "");
  if (filter && !name.includes(filter)) continue;
  const outFile = path.join(outDir, name + ".html");
  if (!fs.existsSync(outFile)) continue;
  const a = stats(fs.readFileSync(path.join(baseDir, f), "utf8"));
  const b = stats(fs.readFileSync(outFile, "utf8"));
  if (a.length === 0) {
    console.log(`${name}: v1 rendered nothing (crash); v2 rendered ${b.length} slides`);
    continue;
  }
  const lines = [];
  if (a.length !== b.length) lines.push(`slide count v1=${a.length} v2=${b.length}`);
  for (let i = 0; i < Math.min(a.length, b.length); i++) {
    total++;
    const d = [];
    for (const k of ["svg", "img", "block", "table"]) if (a[i][k] !== b[i][k]) d.push(`${k} ${a[i][k]}->${b[i][k]}`);
    if (a[i].text !== b[i].text) {
      let j = 0;
      while (j < a[i].text.length && a[i].text[j] === b[i].text[j]) j++;
      d.push(`text@${j}: v1=${JSON.stringify(a[i].text.slice(j, j + 40))} v2=${JSON.stringify(b[i].text.slice(j, j + 40))}`);
    }
    if (d.length) {
      diffs++;
      lines.push(`  slide ${i + 1}: ${d.join("; ")}`);
    }
  }
  console.log(`${name}: ${lines.length ? "\n" + lines.slice(0, 12).join("\n") + (lines.length > 12 ? `\n  ... ${lines.length - 12} more` : "") : "identical structure"}`);
}
console.log(`\n${diffs} of ${total} comparable slides differ`);
