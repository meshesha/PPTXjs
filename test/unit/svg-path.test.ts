import { describe, expect, it } from "vitest";
import { arcToCubic, geometryFromSvgMarkup, parseSvgPath, toEnhancedPath } from "../../src/export/svg-path";

describe("SVG path conversion", () => {
  it("normalises relative and shorthand commands", () => {
    const cmds = parseSvgPath("m10,10 h20 v10 l-5 5 z M0 0 L1 1 2 2");
    expect(cmds).toEqual([
      { c: "M", p: [10, 10] },
      { c: "L", p: [30, 10] },
      { c: "L", p: [30, 20] },
      { c: "L", p: [25, 25] },
      { c: "Z", p: [] },
      { c: "M", p: [0, 0] },
      { c: "L", p: [1, 1] },
      { c: "L", p: [2, 2] },
    ]);
  });

  it("expands S and T shorthands", () => {
    const cmds = parseSvgPath("M0,0 C0,10 10,10 10,0 S20,-10 20,0 Q25,10 30,0 T40,0");
    expect(cmds[2]).toEqual({ c: "C", p: [10, -10, 20, -10, 20, 0] });
    expect(cmds[4]).toEqual({ c: "Q", p: [35, -10, 40, 0] });
  });

  it("approximates arcs with cubic curves ending exactly at the target", () => {
    const segs = arcToCubic(0, 50, 50, 50, 0, false, true, 100, 50);
    expect(segs.length).toBe(2);
    expect(segs.at(-1)!.slice(4)).toEqual([100, 50]);
    // the top of the half circle passes near (50, 0)
    expect(segs[0][5]).toBeCloseTo(0, 5);
  });

  it("extracts renderer geometry and writes integer enhanced paths", () => {
    const html =
      "<svg class='drawing'><defs><marker><path d='M0,0 L5,5'/></marker></defs>" +
      "<rect x='0' y='0' width='100' height='50' fill='red' stroke='blue'/>" +
      "<ellipse cx='50' cy='25' rx='50' ry='25' fill='none' stroke='none'/></svg><div>text</div>";
    const parts = geometryFromSvgMarkup(html);
    expect(parts).toHaveLength(2);
    expect(parts[1].fill).toBe(false);
    const path = toEnhancedPath(parts, 100, 50);
    expect(path.startsWith("M 0 0 L 21600 0 L 21600 21600 L 0 21600 Z N")).toBe(true);
    expect(path).toMatch(/F S N$/);
    expect(path).not.toMatch(/\d\.\d/);
  });
});
