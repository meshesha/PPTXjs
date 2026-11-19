#!/usr/bin/env node
// Converts a .pptx to a standalone HTML page and/or .odp, and can apply simple
// edits. No browser needed.
//
//   node examples/node/convert.mjs deck.pptx --html out.html --odp out.odp
//   node examples/node/convert.mjs deck.pptx --replace "{{name}}=ACME" --pptx filled.pptx
import { readFile, writeFile } from "node:fs/promises";
import { BASE_CSS, PresentationDocument, exportOdpBytes, exportPptxBytes, renderPresentation } from "@meshesha/pptxjs";

const [input, ...args] = process.argv.slice(2);
if (!input) {
  console.error("usage: convert.mjs <file.pptx> [--html out.html] [--odp out.odp] [--pptx out.pptx] [--replace key=value]");
  process.exit(1);
}
const opt = (name) => {
  const i = args.indexOf(name);
  return i === -1 ? undefined : args[i + 1];
};

const doc = await PresentationDocument.load(await readFile(input));

const replace = opt("--replace");
if (replace) {
  const [key, value] = replace.split("=");
  doc.slides().forEach((_, slide) => {
    for (const shape of doc.getShapes(slide)) {
      if (shape.paragraphs && shape.text.includes(key)) doc.setShapeText(slide, shape.id, shape.text.split(key).join(value));
    }
  });
}

if (opt("--html")) {
  const res = renderPresentation(doc.pkg, { onWarning: (w) => console.warn("warning:", w.message) });
  const page =
    `<!doctype html><meta charset="utf-8"><title>${input}</title>` +
    `<style>${BASE_CSS}\n${res.css}\nbody{background:#ddd;margin:0;padding:24px}.pptxjs .slide{margin:0 auto 24px;box-shadow:0 2px 8px #0003}</style>` +
    `<div class="pptxjs">${res.slides.map((s) => s.html).join("\n")}</div>`;
  await writeFile(opt("--html"), page);
  console.log(`HTML: ${res.slides.length} slides -> ${opt("--html")} (${res.renderTime} ms)`);
}
if (opt("--odp")) {
  await writeFile(opt("--odp"), exportOdpBytes(doc));
  console.log(`ODP -> ${opt("--odp")}`);
}
if (opt("--pptx")) {
  await writeFile(opt("--pptx"), exportPptxBytes(doc));
  console.log(`PPTX -> ${opt("--pptx")}`);
}
