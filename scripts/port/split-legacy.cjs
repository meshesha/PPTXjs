// One-off migration tool (kept for traceability): splits the inner functions of
// legacy/js/pptxjs.js into TypeScript modules under src/render/, and generates
// the cross-module imports. Functions that were rewritten by hand (XML parser,
// zip access, colour maths, DOM post-processing ...) are skipped.
//
// usage: node scripts/port/split-legacy.cjs
const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..", "..");
const lines = fs.readFileSync(path.join(root, "legacy/js/pptxjs.js"), "utf8").replace(/\r/g, "").split("\n");

const MODULES = {
  nodes: ["processNodesInSlide", "processGroupSpNode", "processSpNode", "processCxnSpNode", "processPicNode", "processGraphicFrameNode", "processSpPrNode", "genDiagram", "getPosition", "getSize"],
  shape: ["genShape"],
  geometry: ["shapePie", "shapeGear", "shapeArc", "shapeSnipRoundRect", "shapePolygon"],
  text: ["genTextBody", "genBuChar", "getHtmlBullet", "getDingbatToUnicode", "getLayoutAndMasterNode", "genSpanElement", "getPregraphMargn", "getVerticalMargins", "getHorizontalAlign", "getPregraphDir", "getVerticalAlign", "getContentDir", "getFontType", "getFontColorPr", "getFontSize", "getFontBold", "getFontItalic", "getFontDecoration", "getTextHorizontalAlign", "getTextVerticalAlign"],
  table: ["genTable", "getTableCellParams", "getTableBorders"],
  chart: ["genChart", "extractChartData"],
  border: ["getBorder"],
  background: ["getBackground", "getSlideBackgroundFill", "getBgGradientFill", "getBgPicFill"],
  fill: ["getShapeFill", "getFillType", "getGradientFill", "getPicFill", "getPatternFill", "getLinerGrandient", "getSvgGradient", "getMiddleStops", "SVGangle", "getSvgImagePattern", "getBase64ImageDimensions"],
  color: ["getSolidFill", "getSchemeColorFromTheme", "toHex", "hslToRgb", "hueToRgb", "getColorName2Hex", "applyShade", "applyTint", "applyLumOff", "applyLumMod", "applyHueMod", "applySatMod", "rgba2hex"],
};

// Hand-written replacements live in these modules; imports are generated for them too.
const EXTERNAL = {
  "../core/xml": ["getTextByPathList", "getTextByPathStr", "setTextByPathList", "eachElement"],
  "../core/package": ["getMimeType", "extractFileExtension", "isVideoLink", "IsVideoLink"],
  "../utils/base64": ["base64ArrayBuffer"],
  "../utils/html": ["escapeHtml", "escapeAttr", "sanitizeUrl"],
  "./units": ["slideFactor", "fontSizeFactor", "angleToDegrees", "rtl_langs_array"],
  "../utils/tinycolor": ["tinycolor"],
  "../dingbat": ["dingbat_unicode"],
};

// ---- locate functions -------------------------------------------------------
const starts = [];
lines.forEach((l, i) => {
  const m = l.match(/^        function ([A-Za-z0-9_]+)\s*\(([^)]*)\)\s*\{/);
  if (m) starts.push({ name: m[1], params: m[2], line: i });
});
const END_OF_INNER = lines.findIndex((l) => l.includes("/////////////////////////////////////tXml"));
const ranges = {};
for (let k = 0; k < starts.length; k++) {
  const s = starts[k];
  let end = k + 1 < starts.length ? starts[k + 1].line : END_OF_INNER;
  // leave comment lines that directly precede the next function to that function
  while (end > s.line && /^\s*(\/\/|\/\*\*|\*|\*\/|$)/.test(lines[end - 1]) && !/^\s*\}\s*$/.test(lines[end - 1])) end--;
  let begin = s.line;
  while (begin > 0 && /^\s*(\/\*\*|\*|\*\/)/.test(lines[begin - 1])) begin--;
  ranges[s.name] = { begin, end, params: s.params, line: s.line };
}

const owner = {};
for (const [mod, fns] of Object.entries(MODULES)) for (const f of fns) owner[f] = mod;

function dedent(l) {
  return l.startsWith("        ") ? l.slice(8) : l.replace(/^\s{0,7}/, "");
}

function convertSignature(line, params) {
  const typed = params
    .split(",")
    .map((p) => p.trim())
    .filter(Boolean)
    .map((p) => `${p}?: any`)
    .join(", ");
  return line.replace(/function ([A-Za-z0-9_]+)\s*\([^)]*\)/, `export function $1(${typed}): any`);
}

for (const [mod, fns] of Object.entries(MODULES)) {
  const body = [];
  for (const f of fns) {
    const r = ranges[f];
    if (!r) throw new Error("missing function " + f);
    const chunk = lines.slice(r.begin, r.end).map(dedent);
    const sigIdx = r.line - r.begin;
    chunk[sigIdx] = convertSignature(chunk[sigIdx], r.params);
    body.push(`// ported from legacy/js/pptxjs.js:${r.line + 1}`);
    body.push(...chunk, "");
  }
  const text = body.join("\n");
  const stripped = text.replace(/\/\/.*$/gm, "").replace(/"(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*'/g, '""');
  const imports = {};
  for (const [f, m] of Object.entries(owner)) {
    if (m === mod) continue;
    if (new RegExp(`(^|[^.\\w$])${f}\\s*\\(`).test(stripped)) (imports["./" + m] ||= []).push(f);
  }
  for (const [from, names] of Object.entries(EXTERNAL)) {
    for (const n of names) if (new RegExp(`(^|[^.\\w$])${n}\\b`).test(stripped)) (imports[from] ||= []).push(n);
  }
  const header = [
    "/* eslint-disable */",
    "// Ported from PPTXjs v1 (legacy/js/pptxjs.js). Internal rendering code: the",
    "// node objects come from the simplified XML tree, so most values are `any`.",
    'import type { WarpObj } from "./context";',
    ...Object.entries(imports).map(([from, names]) => `import { ${[...new Set(names)].sort().join(", ")} } from "${from}";`),
    "",
  ];
  fs.mkdirSync(path.join(root, "src/render"), { recursive: true });
  fs.writeFileSync(path.join(root, "src/render", mod + ".ts"), header.join("\n") + text);
  console.log(mod.padEnd(12), fns.length, "functions", text.split("\n").length, "lines");
}
