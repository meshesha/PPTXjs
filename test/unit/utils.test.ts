import { describe, expect, it } from "vitest";
import { tinycolor } from "../../src/utils/tinycolor";
import { base64ArrayBuffer } from "../../src/utils/base64";
import { escapeHtml, sanitizeUrl } from "../../src/utils/html";
import { imageSize, imageSizeFromDataUrl } from "../../src/utils/image-size";
import { alphaNumeric, applyNumericBullets, getNumTypeNum, romanize } from "../../src/render/numbering";
import { custGeomToSvgPaths, evalGuide } from "../../src/render/custgeom";
import { parseXmlSimplified } from "../../src/core/xml";
import { dingbatToUnicode } from "../../src/dingbat";
import { getHtmlBullet } from "../../src/render/text";
import { PNG_1PX } from "../helpers/make-pptx";

describe("tinycolor port", () => {
  it("converts like TinyColor 1.4.2", () => {
    expect(tinycolor("FF0000").toHex()).toBe("ff0000");
    expect(tinycolor("#4472C4").toHsl().h).toBeCloseTo(218.4, 0);
    expect(tinycolor({ h: 0, s: 1, l: 0.25, a: 1 }).toHex()).toBe("800000");
    expect(tinycolor("ff000080").toHex8()).toBe("ff000080");
    const c = tinycolor("00ff00");
    c.setAlpha(0.5);
    expect(c.toHex8()).toBe("00ff0080");
    expect(c.toRgbString()).toBe("rgba(0, 255, 0, 0.5)");
    expect(tinycolor("red").toHexString()).toBe("#ff0000");
    expect(tinycolor("nonsense").isValid()).toBe(false);
  });
});

describe("base64", () => {
  it("matches btoa", () => {
    for (const len of [0, 1, 2, 3, 4, 5, 255]) {
      const bytes = Uint8Array.from({ length: len }, (_, i) => (i * 37) & 255);
      expect(base64ArrayBuffer(bytes)).toBe(Buffer.from(bytes).toString("base64"));
    }
  });
});

describe("html safety", () => {
  it("escapeHtml does not double escape parser output", () => {
    expect(escapeHtml("a &amp; <b>")).toBe("a &amp; &lt;b&gt;");
  });

  it.each([
    ["javascript:alert(1)", ""],
    ["JaVaScRiPt:alert(1)", ""],
    [" java\tscript:alert(1)", ""],
    ["&#106;avascript:alert(1)", ""],
    ["data:text/html,<script>", ""],
    ["vbscript:msgbox", ""],
    ["https://example.com/?a=1&amp;b=2", "https://example.com/?a=1&amp;b=2"],
    ["mailto:x@example.com", "mailto:x@example.com"],
    ["#slide2", "#slide2"],
    ["docs/page.html", "docs/page.html"],
  ])("sanitizeUrl(%s)", (input, expected) => {
    expect(sanitizeUrl(input)).toBe(expected);
  });
});

describe("image size", () => {
  it("reads PNG headers", () => {
    expect(imageSize(PNG_1PX)).toEqual([1, 1]);
    expect(imageSizeFromDataUrl("data:image/png;base64," + Buffer.from(PNG_1PX).toString("base64"))).toEqual([1, 1]);
    expect(imageSizeFromDataUrl("not a data url")).toEqual([0, 0]);
  });
});

describe("numbering", () => {
  it("formats schemes", () => {
    expect(getNumTypeNum("arabicPeriod", 3)).toBe("3. ");
    expect(getNumTypeNum("alphaLcParenR", 2)).toBe("b) ");
    expect(getNumTypeNum("alphaUcPeriod", 27)).toBe("AA. ");
    expect(getNumTypeNum("romanLcPeriod", 4)).toBe("iv. ");
    expect(getNumTypeNum("arabicParenBoth", 1)).toBe("(1) ");
    expect(getNumTypeNum("hebrew2Minus", 1)).toBe("א-");
    expect(romanize(1994)).toBe("MCMXCIV");
    expect(alphaNumeric(28, false)).toBe("bb");
  });

  it("numbers placeholders per level and honours startAt", () => {
    const ph = (t: string, l: number, s?: number) =>
      `<div data-bulltname = '${t}' data-bulltlvl = '${l}'${s ? ` data-bulltstart = '${s}'` : ""} class='numeric-bullet-style'></div>`;
    const html = ph("arabicPeriod", 0, 3) + ph("arabicPeriod", 0) + ph("alphaLcPeriod", 1) + ph("alphaLcPeriod", 1) + ph("arabicPeriod", 0);
    const out = applyNumericBullets(html);
    expect(out.match(/>([^<]+)<\/div>/g)).toEqual([">3. </div>", ">4. </div>", ">a. </div>", ">b. </div>", ">5. </div>"]);
  });
});

describe("custom geometry", () => {
  it("evaluates guide formulas", () => {
    const env = { w: 200, h: 100, adj: 50000 };
    expect(evalGuide("*/ w adj 100000", env)).toBe(100);
    expect(evalGuide("+- w 10 20", env)).toBe(190);
    expect(evalGuide("pin 0 adj 25000", env)).toBe(25000);
    expect(evalGuide("?: -1 1 2", env)).toBe(2);
    expect(evalGuide("max w h", env)).toBe(200);
  });

  it("handles several paths, single lnTo, arcs and quadratic curves (issues #41, #43)", () => {
    const geom = parseXmlSimplified(`<a:custGeom><a:gdLst><a:gd name="half" fmla="*/ w 1 2"/></a:gdLst><a:pathLst>
      <a:path w="100" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="half" y="100"/></a:lnTo><a:close/></a:path>
      <a:path w="100" h="100" fill="none"><a:moveTo><a:pt x="0" y="50"/></a:moveTo><a:arcTo wR="50" hR="50" stAng="10800000" swAng="10800000"/><a:quadBezTo><a:pt x="50" y="0"/><a:pt x="0" y="0"/></a:quadBezTo></a:path>
      </a:pathLst></a:custGeom>`)["a:custGeom"];
    const paths = custGeomToSvgPaths(geom, 200, 200);
    expect(paths).toHaveLength(2);
    expect(paths[0].d).toBe("M0,0 L100,200 z");
    expect(paths[1].fill).toBe(false);
    expect(paths[1].d).toMatch(/^M0,100 A100,100 0 0,1 200,100 Q100,0 0,0$/);
  });
});

describe("bullets", () => {
  it("maps symbol-font characters to Unicode (issue #40)", () => {
    expect(getHtmlBullet("Symbol", "")).toBe("&#8226;");
    expect(getHtmlBullet("Wingdings", "")).toBe("&#9632;");
    expect(getHtmlBullet("Wingdings 2", "")).toBe("&#10625;");
    expect(getHtmlBullet("Arial", "&#8226;")).toBe("&#8226;");
    expect(dingbatToUnicode("Webdings", "!")).toBe(128375);
  });
});

describe("fill helpers", () => {
  it("getFillType tolerates missing nodes (issue #33)", async () => {
    const { getFillType } = await import("../../src/render/fill");
    expect(getFillType(undefined)).toBe("");
    expect(getFillType(null)).toBe("");
    expect(getFillType({ "a:solidFill": {} })).toBe("SOLID_FILL");
  });
});
