/**
 * Renders every .pptx in test/fixtures (the sample deck plus files attached to
 * GitHub issues) and checks that no slide fails and the HTML is well formed.
 * Issue-specific assertions guard the fixes.
 */
import fs from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { PptxPackage } from "../../src/core/package";
import { renderPresentation, type RenderedPresentation } from "../../src/render/presentation";

const dir = path.resolve(__dirname, "../fixtures");
const files = fs
  .readdirSync(dir, { recursive: true })
  .map(String)
  .filter((f) => f.endsWith(".pptx"))
  .sort();

const cache = new Map<string, RenderedPresentation>();
function render(name: string): RenderedPresentation {
  if (!cache.has(name)) {
    const pkg = PptxPackage.fromBytes(fs.readFileSync(path.join(dir, name)));
    cache.set(name, renderPresentation(pkg));
  }
  return cache.get(name)!;
}
const fixture = (n: string) => files.find((f) => path.basename(f) === n + ".pptx");

/** Counts unbalanced container tags in an HTML fragment. */
function unbalanced(html: string): string[] {
  const depth: Record<string, number> = {};
  for (const m of html.matchAll(/<(\/?)(div|span|svg|table|tr|td)\b[^>]*?(\/?)>/gi)) {
    if (m[3] === "/") continue;
    const t = m[2].toLowerCase();
    depth[t] = (depth[t] ?? 0) + (m[1] ? -1 : 1);
  }
  return Object.entries(depth)
    .filter(([, v]) => v !== 0)
    .map(([k, v]) => `${k}:${v}`);
}

describe.skipIf(files.length === 0)("fixtures", () => {
  it.each(files)("%s renders every slide", (file) => {
    const res = render(file);
    expect(res.slides.length).toBeGreaterThan(0);
    const failed = res.slides.filter((s) => s.error).map((s) => `${s.number}: ${s.error?.message}`);
    expect(failed).toEqual([]);
    for (const s of res.slides) expect(unbalanced(s.html), `slide ${s.number}`).toEqual([]);
    expect(res.width).toBeGreaterThan(0);
    // no executable markup ever reaches the output
    for (const s of res.slides) {
      expect(s.html).not.toMatch(/<script\b/i);
      expect(s.html).not.toMatch(/href='\s*javascript:/i);
    }
  });

  it.runIf(!!fixture("issue-42"))("#42: file without docProps/app.xml opens", () => {
    expect(render(fixture("issue-42")!).appVersion).toBe(0);
  });

  it.runIf(!!fixture("issue-39"))("#39: picture inside mc:AlternateContent is rendered", () => {
    const s = render(fixture("issue-39")!).slides[10];
    expect(s.html).toContain("<img src='data:image/");
  });

  it.runIf(!!fixture("issue-40"))("#40: Wingdings bullets become real Unicode bullets", () => {
    const html = render(fixture("issue-40")!).slides[1].html;
    expect(html).toContain("&#10625;");
    expect(html).not.toContain("&#151;");
  });

  it.runIf(!!fixture("issue-13"))("startAt of auto-numbered bullets is honoured", () => {
    expect(render(fixture("issue-13")!).slides[19].html).toContain(">3. </div>");
  });

  it.runIf(!!fixture("issue-20") || !!fixture("Sample_12"))("#20: emits <span>, never <sapn>", () => {
    expect(render(fixture("Sample_12")!).slides.some((s) => s.html.includes("<sapn"))).toBe(false);
  });

  it.runIf(!!fixture("Sample_12"))("charts are rendered as inline SVG (no d3/nvd3)", () => {
    const res = render(fixture("Sample_12")!);
    expect(res.slides.some((s) => s.html.includes("class=\"pptx-chart\""))).toBe(true);
  });

  it.runIf(!!fixture("Sample_12"))("group shapes have no debug border", () => {
    expect(render(fixture("Sample_12")!).slides.every((s) => !s.html.includes("border:1px solid red"))).toBe(true);
  });
});
