/**
 * Security regression tests: crafted presentations must never produce
 * executable HTML, pollute prototypes, or exhaust memory.
 */
// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
import { PptxPackage } from "../../src/core/package";
import { renderPresentation } from "../../src/render/presentation";
import { PptxError } from "../../src/errors";
import { makePptx, sp } from "../helpers/make-pptx";

function renderHtml(bytes: Uint8Array, opts = {}): string {
  const res = renderPresentation(PptxPackage.fromBytes(bytes), opts);
  return res.slides.map((s) => s.html).join("") + "<style>" + res.css + "</style>";
}

/** Parses HTML with the browser parser and reports anything executable. */
function dangerous(html: string): string[] {
  const doc = new DOMParser().parseFromString(`<body>${html}</body>`, "text/html");
  const found: string[] = [];
  doc.querySelectorAll("script, object, embed, base, meta, link").forEach((e) => found.push(e.tagName));
  doc.querySelectorAll("*").forEach((e) => {
    for (const a of Array.from(e.attributes)) {
      if (/^on/i.test(a.name)) found.push(`${e.tagName}[${a.name}]`);
      if (/^(href|src|xlink:href|action|formaction)$/i.test(a.name) && /^\s*(javascript|vbscript|data:text)/i.test(a.value)) found.push(`${e.tagName}[${a.name}=${a.value}]`);
    }
  });
  if ((doc.querySelectorAll("iframe").length)) found.push("iframe");
  return found;
}

const PAYLOADS = [
  `&lt;img src=x onerror=alert(1)&gt;`,
  `&lt;/style&gt;&lt;script&gt;alert(1)&lt;/script&gt;`,
  `"'&gt;&lt;svg onload=alert(1)&gt;`,
  `<![CDATA[<script>alert(1)</script>]]>`,
];

describe("XSS hardening", () => {
  it.each(PAYLOADS)("text payload %s is inert", (payload) => {
    const html = renderHtml(makePptx([{ shapes: sp(2, payload) }]));
    expect(dangerous(html)).toEqual([]);
  });

  it("attribute payloads (font names, shape names) cannot break out", () => {
    const evilFont = `x' onmouseover='alert(1)' a='`;
    const shapes =
      sp(2, "hello", { rPr: "", extra: "" }).replace('name="Shape 2"', `name="${evilFont}"`).replace("<a:rPr lang=\"en-US\"/>", `<a:rPr lang="en-US"><a:latin typeface="${evilFont}"/></a:rPr>`) +
      sp(3, "x", { pPr: `<a:pPr><a:buFont typeface="${evilFont}"/><a:buChar char="'"/></a:pPr>` });
    const html = renderHtml(makePptx([{ shapes }]));
    expect(dangerous(html)).toEqual([]);
  });

  it("javascript: hyperlinks are dropped (issue #3)", () => {
    const run = `<a:rPr lang="en-US"><a:hlinkClick r:id="rId5"/></a:rPr>`;
    const shapes = sp(2, "click me").replace('<a:rPr lang="en-US"/>', run);
    const html = renderHtml(makePptx([{ shapes, rels: [["rId5", "hyperlink", "javascript:alert(document.cookie)", true]] }]));
    expect(html).toContain("click me");
    expect(dangerous(html)).toEqual([]);
    expect(html).not.toMatch(/javascript:/i);
  });

  it("external links open with rel=noopener", () => {
    const shapes = sp(2, "site").replace('<a:rPr lang="en-US"/>', `<a:rPr lang="en-US"><a:hlinkClick r:id="rId5"/></a:rPr>`);
    const html = renderHtml(makePptx([{ shapes, rels: [["rId5", "hyperlink", "https://example.com/", true]] }]));
    expect(html).toContain("href='https://example.com/'");
    expect(html).toContain("rel='noopener noreferrer'");
  });

  it("external video links are not embedded unless enabled", () => {
    const pic =
      `<p:pic><p:nvPicPr><p:cNvPr id="4" name="Video"/><p:cNvPicPr/><p:nvPr><a:videoFile r:link="rId6"/></p:nvPr></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="rId7"/></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr></p:pic>`;
    const bytes = makePptx(
      [{ shapes: pic, rels: [["rId6", "video", "https://evil.example/embed", true], ["rId7", "image", "../media/image1.png"]] }],
      { extra: { "ppt/media/image1.png": new Uint8Array([137, 80, 78, 71]) } },
    );
    const off = renderHtml(bytes);
    expect(off).not.toContain("<iframe");
    expect(off).toContain("pptx-media-link");
    const on = renderHtml(bytes, { externalMedia: true });
    expect(on).toContain("sandbox='allow-scripts allow-same-origin allow-presentation'");
  });

  it("no prototype pollution from element names", () => {
    const shapes = sp(2, "x").replace("<p:txBody>", `<__proto__ polluted="1"/><p:txBody>`);
    renderHtml(makePptx([{ shapes }]));
    expect(({} as any).polluted).toBeUndefined();
    expect((Object.prototype as any).set).toBeUndefined();
  });
});

describe("resource limits", () => {
  it("rejects non-zip input", () => {
    expect(() => PptxPackage.fromBytes(strToU8("hello world, definitely not a zip"))).toThrow(PptxError);
  });

  it("rejects zips that are not OOXML", () => {
    expect(() => PptxPackage.fromBytes(zipSync({ "a.txt": strToU8("x") }))).toThrowError(/Content_Types/);
  });

  it("enforces the uncompressed size limit (zip bomb)", () => {
    const big = new Uint8Array(5 * 1024 * 1024); // compresses to a few KB
    const bomb = zipSync({ "[Content_Types].xml": strToU8("<Types/>"), "big.bin": big });
    expect(bomb.length).toBeLessThan(100 * 1024);
    let err: unknown;
    try {
      PptxPackage.fromBytes(bomb, { maxUncompressedSize: 1024 * 1024 });
    } catch (e) {
      err = e;
    }
    expect(err).toBeInstanceOf(PptxError);
    expect((err as PptxError).code).toBe("LIMIT_EXCEEDED");
  });

  it("enforces the entry count limit", () => {
    const files: Record<string, Uint8Array> = { "[Content_Types].xml": strToU8("<Types/>") };
    for (let i = 0; i < 50; i++) files[`f${i}`] = strToU8("x");
    expect(() => PptxPackage.fromBytes(zipSync(files), { maxEntries: 10 })).toThrowError(/entries/);
  });
});
