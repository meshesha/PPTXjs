/**
 * Editing model + export round trips.
 */
import { describe, expect, it } from "vitest";
import { unzipSync, strFromU8 } from "fflate";
import { PptxPackage } from "../../src/core/package";
import { parseXmlDocument } from "../../src/core/xmldom";
import { PresentationDocument } from "../../src/model/document";
import { renderPresentation, listSlides } from "../../src/render/presentation";
import { exportPptxBytes } from "../../src/export/pptx";
import { exportOdpBytes, ODP_MIME } from "../../src/export/odp";
import { makePptx, sp, PNG_1PX } from "../helpers/make-pptx";

const doc = (slides = 2) =>
  new PresentationDocument(
    PptxPackage.fromBytes(
      makePptx(Array.from({ length: slides }, (_, i) => ({ shapes: sp(2, `Hello ${i + 1}`) + sp(3, "", { ph: '<p:ph type="title"/>' }) }))),
    ),
  );
const reopen = (d: PresentationDocument) => PptxPackage.fromBytes(exportPptxBytes(d));

describe("xmldom", () => {
  it("round-trips XML losslessly", () => {
    const src = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a:r xmlns:a="x"><a:t xml:space="preserve"> two  spaces &amp; &lt;tag&gt; </a:t><!-- c --><a:b c='q"'/><![CDATA[<x>]]></a:r>`;
    const out = parseXmlDocument(src).toString();
    expect(out).toContain(`<a:t xml:space="preserve"> two  spaces &amp; &lt;tag&gt; </a:t>`);
    expect(out).toContain(`c="q&quot;"`);
    expect(out).toContain(`<![CDATA[<x>]]>`);
    expect(parseXmlDocument(out).root.child("a:t")!.text).toBe(" two  spaces & <tag> ");
  });

  it("rejects malformed XML", () => {
    expect(() => parseXmlDocument("<a><b></a>")).toThrow();
  });
});

describe("PresentationDocument", () => {
  it("lists slides and shapes, resolving placeholder bounds from the layout/master", () => {
    const d = doc();
    expect(d.slideCount).toBe(2);
    const shapes = d.getShapes(0);
    expect(shapes.map((s) => s.kind)).toEqual(["textbox", "shape"]);
    expect(shapes[0].text).toBe("Hello 1");
    expect(shapes[0].bounds).toEqual({ x: 96, y: 96, width: 384, height: 96 });
    // title placeholder without xfrm inherits from the master
    expect(shapes[1].placeholder?.type).toBe("title");
    expect(shapes[1].bounds!.width).toBeCloseTo(1104, 0);
  });

  it("moves, resizes, rotates and edits text; undo/redo restore exactly", () => {
    const d = doc();
    const before = d.pkg.text("ppt/slides/slide1.xml");
    d.setShapeBounds(0, "2", { x: 10, y: 20, width: 200 });
    d.setShapeRotation(0, "2", 90);
    d.setShapeText(0, "2", "Line one\nLine & <two>");
    d.setTextStyle(0, "2", { bold: true, size: 32, color: "#ff0000", align: "center" });
    const s = d.getShape(0, "2");
    expect(s.bounds).toEqual({ x: 10, y: 20, width: 200, height: 96 });
    expect(s.rotation).toBe(90);
    expect(s.paragraphs!.map((p) => p.runs.map((r) => r.text).join(""))).toEqual(["Line one", "Line & <two>"]);
    expect(s.paragraphs![0].runs[0]).toMatchObject({ bold: true, size: 32, color: "#FF0000" });
    expect(s.paragraphs![0].align).toBe("center");
    expect(d.canUndo).toBe(true);
    while (d.canUndo) d.undo();
    expect(d.pkg.text("ppt/slides/slide1.xml")).toBe(before);
    d.redo();
    expect(d.getShape(0, "2").bounds!.x).toBe(10);
  });

  it("sets a placeholder's bounds by materialising its xfrm", () => {
    const d = doc();
    d.setShapeBounds(0, "3", { x: 50 });
    const b = d.getShape(0, "3").bounds!;
    expect(b.x).toBe(50);
    expect(b.width).toBeCloseTo(1104, 0);
  });

  it("adds text boxes, shapes and images; deletes and reorders shapes", () => {
    const d = doc(1);
    const tb = d.addTextBox(0, { x: 0, y: 0, width: 100, height: 40 }, "New text", { bold: true, color: "#00ff00" });
    const shape = d.addShape(0, "ellipse", { x: 10, y: 10, width: 50, height: 50 }, { fill: "#123456", text: "In shape" });
    const img = d.addImage(0, PNG_1PX, "image/png", { x: 5, y: 5, width: 10, height: 10 });
    const kinds = d.getShapes(0).map((s) => [s.id, s.kind]);
    expect(kinds).toEqual([["2", "textbox"], ["3", "shape"], [tb, "textbox"], [shape, "shape"], [img, "picture"]]);
    expect(d.getShape(0, shape).fill).toBe("#123456");
    expect(d.pkg.list().some((p) => p.startsWith("ppt/media/image"))).toBe(true);
    expect(d.pkg.text("[Content_Types].xml")).toContain('Extension="png"');
    d.reorderShape(0, img, "back");
    expect(d.getShapes(0)[0].id).toBe(img);
    d.reorderShape(0, img, "forward");
    expect(d.getShapes(0)[1].id).toBe(img);
    d.reorderShape(0, img, "front");
    expect(d.getShapes(0).at(-1)!.id).toBe(img);
    const dup = d.duplicateShape(0, tb);
    expect(d.getShape(0, dup).text).toBe("New text");
    d.deleteShape(0, tb);
    expect(d.getShapes(0).some((s) => s.id === tb)).toBe(false);
    // the edited deck renders
    const res = renderPresentation(reopen(d));
    expect(res.slides[0].error).toBeUndefined();
    expect(res.slides[0].html).toContain("In shape");
  });

  it("adds, duplicates, moves, hides and deletes slides", () => {
    const d = doc(2);
    const added = d.addSlide(0);
    expect(added).toBe(1);
    expect(d.slideCount).toBe(3);
    const copy = d.duplicateSlide(0);
    expect(d.getShapes(copy)[0].text).toBe("Hello 1");
    d.moveSlide(0, 3);
    expect(d.getShapes(3)[0].text).toBe("Hello 1");
    d.setSlideHidden(3, true);
    expect(d.slides()[3].hidden).toBe(true);
    d.deleteSlide(0);
    expect(d.slideCount).toBe(3);
    const pkg = reopen(d);
    expect(listSlides(pkg)).toHaveLength(3);
    const res = renderPresentation(pkg);
    expect(res.slides.every((s) => !s.error)).toBe(true);
    // undo everything
    while (d.canUndo) d.undo();
    expect(d.slideCount).toBe(2);
  });

  it("rolls back a failing transaction", () => {
    const d = doc(1);
    const before = d.pkg.text("ppt/slides/slide1.xml");
    expect(() => d.setShapeFill(0, "2", "not-a-colour")).toThrow();
    expect(d.pkg.text("ppt/slides/slide1.xml")).toBe(before);
    expect(d.canUndo).toBe(false);
  });

  it("emits change events", () => {
    const d = doc(1);
    const events: string[] = [];
    d.on("change", (c) => events.push(`${c.type}:${c.label}`));
    d.setShapeFill(0, "2", "#ff0000");
    d.addSlide();
    d.undo();
    expect(events).toEqual(["slide:Fill colour", "structure:Add slide", "structure:Add slide"]);
  });
});

describe("rendering rules", () => {
  it("#38: slide-level showMasterSp=0 hides layout and master shapes", () => {
    const pkg = (attrs?: string) => PptxPackage.fromBytes(makePptx([{ shapes: sp(2, "Own"), attrs }]));
    const shown = renderPresentation(pkg()).slides[0].html;
    const hidden = renderPresentation(pkg('showMasterSp="0"')).slides[0].html;
    expect(shown).toContain("Title Placeholder 1");
    expect(hidden).not.toContain("Title Placeholder 1");
    expect(hidden).toContain("Own");
  });

  it("#15: slide order follows presentation.xml, not file names", () => {
    const pkg = PptxPackage.fromBytes(makePptx([{ shapes: sp(2, "First file") }, { shapes: sp(2, "Second file") }], { reverseOrder: true }));
    const res = renderPresentation(pkg);
    expect(res.slides.map((s) => s.file)).toEqual(["ppt/slides/slide2.xml", "ppt/slides/slide1.xml"]);
  });

  it("#42: renders without docProps/app.xml", () => {
    const res = renderPresentation(PptxPackage.fromBytes(makePptx([{ shapes: sp(2, "x") }], { noAppXml: true })));
    expect(res.slides[0].error).toBeUndefined();
  });

  it("hidden slides can be skipped", () => {
    const pkg = PptxPackage.fromBytes(makePptx([{ shapes: sp(2, "a") }, { shapes: sp(2, "b"), attrs: 'show="0"' }]));
    expect(renderPresentation(pkg).slides[1].hidden).toBe(true);
    expect(renderPresentation(pkg, { skipHiddenSlides: true }).slides).toHaveLength(1);
  });
});

describe("export", () => {
  it("pptx export is lossless for untouched parts and drops unused media", () => {
    const d = doc(1);
    const img = d.addImage(0, PNG_1PX, "png");
    d.deleteShape(0, img); // rel remains -> media still referenced
    const files = unzipSync(exportPptxBytes(d));
    expect(Object.keys(files)[0]).toBe("[Content_Types].xml");
    expect(strFromU8(files["ppt/theme/theme1.xml"])).toBe(d.pkg.text("ppt/theme/theme1.xml"));
    d.deleteSlide(0);
    const files2 = unzipSync(exportPptxBytes(d));
    expect(Object.keys(files2).some((f) => f.startsWith("ppt/media/"))).toBe(false);
  });

  it("odp export produces a valid ODF package", () => {
    const d = doc(2);
    d.addShape(0, "rightArrow", { x: 10, y: 10, width: 100, height: 50 }, { fill: "#ff0000", text: "Arrow" });
    d.addImage(1, PNG_1PX, "png");
    d.setShapeRotation(0, "2", 45);
    const bytes = exportOdpBytes(d);
    const files = unzipSync(bytes);
    const names = Object.keys(files);
    expect(names[0]).toBe("mimetype");
    expect(strFromU8(files.mimetype)).toBe(ODP_MIME);
    // mimetype must be stored (uncompressed): its bytes appear verbatim at offset 38
    expect(new TextDecoder().decode(bytes.slice(38, 38 + ODP_MIME.length))).toBe(ODP_MIME);
    for (const f of ["content.xml", "styles.xml", "meta.xml", "META-INF/manifest.xml"]) {
      expect(() => parseXmlDocument(strFromU8(files[f])), f).not.toThrow();
    }
    const content = strFromU8(files["content.xml"]);
    expect(content.match(/<draw:page /g)).toHaveLength(2);
    expect(content).toContain(">Hello 1<");
    expect(content).toContain('draw:type="right-arrow"');
    // explicit geometry, integer coordinates (PowerPoint ignores named shape types)
    expect(content).toMatch(/draw:enhanced-path="M [0-9]+ [0-9]+ /);
    expect(content.toLowerCase()).toContain('draw:fill-color="#ff0000"');
    // geometric shapes use the master's "other" text style: no inherited body bullet
    expect(content).not.toMatch(/• <\/text:span><text:span[^>]*>Arrow/);
    expect(content).toContain("draw:transform=\"rotate (");
    expect(content).toMatch(/xlink:href="Pictures\/img1\.png"/);
    expect(strFromU8(files["META-INF/manifest.xml"])).toContain("Pictures/img1.png");
  });
});
