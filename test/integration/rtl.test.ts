/**
 * Right-to-left support (Hebrew, Arabic ...).
 */
import { describe, expect, it } from "vitest";
import { strFromU8, unzipSync } from "fflate";
import { PptxPackage } from "../../src/core/package";
import { renderPresentation } from "../../src/render/presentation";
import { firstStrongDirection, isRtlLanguage } from "../../src/render/bidi";
import { PresentationDocument } from "../../src/model/document";
import { exportOdpBytes } from "../../src/export/odp";
import { makePptx, sp } from "../helpers/make-pptx";

const HEB = "שלום עולם";
const deck = (pPr = "") => makePptx([{ shapes: sp(2, HEB, { pPr }) + sp(3, "Hello world", { y: 2743200 }) }]);
const render = (bytes: Uint8Array, textDirection?: "file" | "auto" | "rtl") =>
  renderPresentation(PptxPackage.fromBytes(bytes), { textDirection }).slides[0].html;
/** paragraph direction classes in document order */
const dirs = (html: string) => [...html.matchAll(/slide-prgrph [^']*?(pregraph-(?:rtl|ltr|inherit))/g)].map((m) => m[1]);

describe("bidi helpers", () => {
  it("detects RTL languages by primary subtag and script", () => {
    for (const l of ["he-IL", "he", "iw-IL", "ar-EG", "ar-SA", "fa-IR", "ur-PK", "yi", "ckb-IQ", "ku-Arab-IQ", "pa-Arab"]) expect(isRtlLanguage(l), l).toBe(true);
    for (const l of ["en-US", "fr", "ku-Latn", "", undefined]) expect(isRtlLanguage(l), String(l)).toBe(false);
  });

  it("finds the first strong direction", () => {
    expect(firstStrongDirection("  123 שלום abc")).toBe("rtl");
    expect(firstStrongDirection("(1) abc שלום")).toBe("ltr");
    expect(firstStrongDirection("مرحبا")).toBe("rtl");
    expect(firstStrongDirection("12 - 3")).toBeNull();
  });
});

describe("rendering", () => {
  it("honours explicit rtl=\"1\" in every mode", () => {
    for (const mode of ["file", "auto", "rtl"] as const) {
      const html = render(deck('<a:pPr algn="r" rtl="1"/>'), mode);
      expect(dirs(html)[0], mode).toBe("pregraph-rtl");
      // the text container inherits the paragraph direction instead of resetting it to LTR
      expect(html).toContain("direction: inherit");
      expect(html).not.toContain("direction: initial");
    }
  });

  it("file mode (default) keeps paragraphs without direction LTR, like PowerPoint", () => {
    expect(dirs(render(deck()))).toEqual(["pregraph-inherit", "pregraph-inherit"]);
  });

  it("auto mode detects Hebrew paragraphs and leaves English ones LTR", () => {
    expect(dirs(render(deck(), "auto"))).toEqual(["pregraph-rtl", "pregraph-ltr"]);
  });

  it("rtl mode makes paragraphs without direction RTL", () => {
    expect(dirs(render(deck(), "rtl"))).toEqual(["pregraph-rtl", "pregraph-rtl"]);
  });

  it("explicit rtl=\"0\" wins over the policy", () => {
    expect(dirs(render(deck('<a:pPr rtl="0"/>'), "rtl"))[0]).toBe("pregraph-ltr");
  });

  it("presentation default text style direction is used", () => {
    const pkg = PptxPackage.fromBytes(deck());
    pkg.setText("ppt/presentation.xml", pkg.text("ppt/presentation.xml")!.replace('<a:lvl1pPr>', '<a:lvl1pPr rtl="1">'));
    expect(dirs(renderPresentation(pkg).slides[0].html)).toEqual(["pregraph-rtl", "pregraph-rtl"]);
  });
});

describe("editing", () => {
  it("marks new Hebrew text as RTL and right-aligned; keeps explicit directions", () => {
    const doc = new PresentationDocument(PptxPackage.fromBytes(makePptx([{ shapes: sp(2, "x") }])));
    const id = doc.addTextBox(0, { x: 10, y: 10, width: 300, height: 50 }, `${HEB}\nEnglish line`);
    const paras = doc.getShape(0, id).paragraphs!;
    expect(paras[0]).toMatchObject({ rtl: true, align: "right" });
    expect(paras[1].rtl).toBeUndefined();
    doc.setShapeText(0, "2", HEB);
    expect(doc.getShape(0, "2").paragraphs![0].rtl).toBe(true);
    doc.setShapeText(0, "2", [{ runs: [{ text: HEB }], rtl: false }]);
    expect(doc.getShape(0, "2").paragraphs![0].rtl).toBe(false);
    // the edited deck renders the RTL paragraph RTL even in "file" mode
    expect(dirs(renderPresentation(doc.pkg).slides[0].html)).toContain("pregraph-rtl");
  });
});

describe("ODP export", () => {
  it("writes RTL paragraphs with rl-tb writing mode and right alignment", () => {
    const doc = new PresentationDocument(PptxPackage.fromBytes(deck('<a:pPr algn="r" rtl="1" marL="457200"/>')));
    const content = strFromU8(unzipSync(exportOdpBytes(doc))["content.xml"]);
    expect(content).toMatch(/<style:paragraph-properties fo:text-align="right" style:writing-mode="rl-tb"/);
    expect(content).toContain(HEB);
    // the English paragraph stays LTR
    expect(content).toMatch(/fo:text-align="left"\/>/);
  });

  it("applies the textDirection option", () => {
    const doc = new PresentationDocument(PptxPackage.fromBytes(deck()));
    const plain = strFromU8(unzipSync(exportOdpBytes(doc))["content.xml"]);
    const auto = strFromU8(unzipSync(exportOdpBytes(doc, { textDirection: "auto" }))["content.xml"]);
    expect(plain).not.toContain("rl-tb");
    expect(auto).toContain("rl-tb");
  });
});

describe("alignment of policy-RTL paragraphs", () => {
  it("auto mode right-aligns detected RTL paragraphs that have no alignment", () => {
    const html = render(deck(), "auto");
    const classes = [...html.matchAll(/class='slide-prgrph ([^']*)'/g)].map((m) => m[1]);
    expect(classes[0]).toMatch(/^h-right-rtl pregraph-rtl/);
    expect(classes[1]).not.toMatch(/h-right/);
    expect(html).not.toContain("slide-prgrph undefined");
  });
});
