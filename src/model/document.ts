/**
 * Editable presentation model. All edits are applied to the real OOXML parts
 * of the package (lossless round trip), so exporting back to .pptx keeps
 * everything that was not edited exactly as it was.
 */
import { PptxPackage, normalizePath, extractFileExtension, getMimeType } from "../core/package";
import { parseXmlDocument, x, XElement, XDocument, XText } from "../core/xmldom";
import { PptxError } from "../errors";
import { readSource, type PptxSource } from "../source";
import { Emitter, type Unsubscribe } from "../utils/emitter";
import { imageSize } from "../utils/image-size";
import { EMU_PER_PX } from "../render/units";
import { firstStrongDirection } from "../render/bidi";

// ---------------------------------------------------------------------------
// types
// ---------------------------------------------------------------------------

/** Rectangle in CSS pixels (96 dpi). */
export interface Bounds {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface TextRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  /** font size in points */
  size?: number;
  /** `#rrggbb` */
  color?: string;
  /** font family */
  font?: string;
}

export interface TextParagraph {
  runs: TextRun[];
  align?: "left" | "center" | "right" | "justify";
  /** indentation level 0..8 */
  level?: number;
  /**
   * Right-to-left paragraph (Hebrew, Arabic ...). When omitted on write, the
   * existing paragraph's direction is kept, and new paragraphs whose first
   * strong character is RTL are marked RTL automatically.
   */
  rtl?: boolean;
}

export type ShapeKind = "shape" | "textbox" | "picture" | "table" | "chart" | "diagram" | "group" | "connector" | "media" | "other";

export interface ShapeInfo {
  /** `p:cNvPr/@id`, unique within the slide */
  id: string;
  name: string;
  kind: ShapeKind;
  /** position and size in px; `null` when not resolvable */
  bounds: Bounds | null;
  /** rotation in degrees */
  rotation: number;
  placeholder?: { type?: string; idx?: string };
  /** preset geometry, e.g. "rect", "ellipse" */
  geometry?: string;
  /** solid fill `#rrggbb` when set directly on the shape */
  fill?: string;
  /** text content, `null` for shapes that cannot hold text */
  paragraphs: TextParagraph[] | null;
  /** plain text (paragraphs joined by "\n") */
  text: string;
}

export interface SlideRef {
  index: number;
  file: string;
  hidden: boolean;
}

export interface DocumentChange {
  /** "slide" = one slide's content changed; "structure" = slides added/removed/reordered */
  type: "slide" | "structure";
  /** affected slide index for "slide" changes */
  index?: number;
  /** description of the operation (for undo menus) */
  label: string;
}

export interface DocumentEvents {
  change: DocumentChange;
  undo: { label: string };
  redo: { label: string };
}

export type PresetGeometry =
  | "rect"
  | "roundRect"
  | "ellipse"
  | "triangle"
  | "rtTriangle"
  | "diamond"
  | "parallelogram"
  | "trapezoid"
  | "pentagon"
  | "hexagon"
  | "octagon"
  | "star5"
  | "rightArrow"
  | "leftArrow"
  | "upArrow"
  | "downArrow"
  | "line"
  | (string & {});

// ---------------------------------------------------------------------------
// constants
// ---------------------------------------------------------------------------
const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const CT_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const SHAPE_TAGS = ["p:sp", "p:pic", "p:graphicFrame", "p:grpSp", "p:cxnSp", "p:contentPart", "mc:AlternateContent"];
const FILL_TAGS = ["a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill"];
/** child order of a:spPr (CT_ShapeProperties) */
const SPPR_ORDER = ["a:xfrm", "a:custGeom", "a:prstGeom", ...FILL_TAGS, "a:ln", "a:effectLst", "a:effectDag", "a:scene3d", "a:sp3d", "a:extLst"];
/** child order of a:rPr (CT_TextCharacterProperties) */
const RPR_ORDER = ["a:ln", ...FILL_TAGS, "a:effectLst", "a:effectDag", "a:highlight", "a:uLnTx", "a:uLn", "a:uFillTx", "a:uFill", "a:latin", "a:ea", "a:cs", "a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst"];

const px2emu = (px: number) => Math.round(px * EMU_PER_PX);
const emu2px = (emu: string | number | undefined) => (emu === undefined ? NaN : Number(emu) / EMU_PER_PX);

function cleanHex(color: string): string {
  const c = color.trim().replace(/^#/, "");
  if (/^[0-9a-fA-F]{6}$/.test(c)) return c.toUpperCase();
  if (/^[0-9a-fA-F]{3}$/.test(c)) return c.split("").map((ch) => ch + ch).join("").toUpperCase();
  throw new PptxError("EXPORT_FAILED", `Invalid colour "${color}" (expected #rrggbb)`);
}

/** Inserts `child` into `parent` respecting the schema order given in `order`. */
function insertOrdered(parent: XElement, child: XElement, order: string[]): void {
  const rank = order.indexOf(child.name);
  if (rank === -1) {
    parent.append(child);
    return;
  }
  let idx = parent.children.length;
  for (let i = 0; i < parent.children.length; i++) {
    const c = parent.children[i];
    if (c.type === "element") {
      const r = order.indexOf(c.name);
      if (r > rank) {
        idx = i;
        break;
      }
    }
  }
  parent.insert(child, idx);
}

function relsPathOf(part: string): string {
  const i = part.lastIndexOf("/");
  return part.slice(0, i + 1) + "_rels/" + part.slice(i + 1) + ".rels";
}

function relativeTarget(fromPart: string, toPart: string): string {
  const from = fromPart.split("/").slice(0, -1);
  const to = toPart.split("/");
  let i = 0;
  while (i < from.length && i < to.length - 1 && from[i] === to[i]) i++;
  return [...Array(from.length - i).fill(".."), ...to.slice(i)].join("/");
}

function resolveRel(fromPart: string, target: string): string {
  if (target.startsWith("/")) return normalizePath(target);
  const dir = fromPart.split("/").slice(0, -1).join("/");
  return normalizePath(dir + "/" + target);
}

// ---------------------------------------------------------------------------
// document
// ---------------------------------------------------------------------------
interface UndoEntry {
  label: string;
  parts: Map<string, Uint8Array | null>;
  change: DocumentChange;
}

export class PresentationDocument {
  readonly pkg: PptxPackage;
  private readonly emitter = new Emitter<DocumentEvents>();
  private undoStack: UndoEntry[] = [];
  private redoStack: UndoEntry[] = [];
  private recording: Map<string, Uint8Array | null> | null = null;
  /** maximum number of undo steps (default 100) */
  undoLimit = 100;

  constructor(pkg: PptxPackage) {
    this.pkg = pkg;
    pkg.onWrite((path, before) => {
      if (this.recording && !this.recording.has(path)) this.recording.set(path, before ? before.slice() : null);
    });
  }

  static async load(source: PptxSource): Promise<PresentationDocument> {
    const bytes = await readSource(source);
    return new PresentationDocument(PptxPackage.fromBytes(bytes));
  }

  on<K extends keyof DocumentEvents>(event: K, handler: (p: DocumentEvents[K]) => void): Unsubscribe {
    return this.emitter.on(event, handler);
  }

  // ------------------------------------------------------------- undo / redo
  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }
  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }
  get undoLabel(): string | undefined {
    return this.undoStack[this.undoStack.length - 1]?.label;
  }
  get redoLabel(): string | undefined {
    return this.redoStack[this.redoStack.length - 1]?.label;
  }

  /**
   * Runs `fn` as one undoable step. Nested calls join the outer step.
   * If `fn` throws, all of its writes are rolled back.
   */
  transact<T>(change: DocumentChange, fn: () => T): T {
    if (this.recording) return fn();
    const rec = new Map<string, Uint8Array | null>();
    this.recording = rec;
    let result: T;
    try {
      result = fn();
    } catch (e) {
      this.recording = null;
      for (const [path, data] of rec) this.pkg.restore(path, data);
      throw e;
    }
    this.recording = null;
    if (rec.size) {
      this.undoStack.push({ label: change.label, parts: rec, change });
      if (this.undoStack.length > this.undoLimit) this.undoStack.shift();
      this.redoStack = [];
      this.emitter.emit("change", change);
    }
    return result;
  }

  undo(): boolean {
    const entry = this.undoStack.pop();
    if (!entry) return false;
    this.redoStack.push(this.swap(entry));
    this.emitter.emit("undo", { label: entry.label });
    this.emitter.emit("change", entry.change);
    return true;
  }

  redo(): boolean {
    const entry = this.redoStack.pop();
    if (!entry) return false;
    this.undoStack.push(this.swap(entry));
    this.emitter.emit("redo", { label: entry.label });
    this.emitter.emit("change", entry.change);
    return true;
  }

  private swap(entry: UndoEntry): UndoEntry {
    const current = new Map<string, Uint8Array | null>();
    for (const [path, data] of entry.parts) {
      current.set(path, this.pkg.bytes(path)?.slice() ?? null);
      this.pkg.restore(path, data);
    }
    return { label: entry.label, parts: current, change: entry.change };
  }

  // ------------------------------------------------------------- xml helpers
  private read(path: string): XDocument {
    const text = this.pkg.text(path);
    if (text === null) throw new PptxError("INVALID_FILE", `Part ${path} not found`);
    return parseXmlDocument(text);
  }

  private write(path: string, doc: XDocument): void {
    this.pkg.setText(path, doc.toString());
  }

  private readRels(part: string): XDocument {
    const p = relsPathOf(part);
    const text = this.pkg.text(p);
    if (text !== null) return parseXmlDocument(text);
    return new XDocument(x("Relationships", { xmlns: NS_PKG_REL }));
  }

  private rels(part: string): Array<{ id: string; type: string; target: string; external: boolean; el: XElement }> {
    return this.readRels(part)
      .root.childrenNamed("Relationship")
      .map((el) => ({
        id: el.attr("Id") ?? "",
        type: el.attr("Type") ?? "",
        target: el.attr("TargetMode") === "External" ? (el.attr("Target") ?? "") : resolveRel(part, el.attr("Target") ?? ""),
        external: el.attr("TargetMode") === "External",
        el,
      }));
  }

  /** Adds a relationship from `part` to `targetPart` and returns its id. */
  private addRel(part: string, type: string, targetPart: string, external = false): string {
    const doc = this.readRels(part);
    const used = new Set(doc.root.childrenNamed("Relationship").map((e) => e.attr("Id")));
    let n = used.size + 1;
    while (used.has("rId" + n)) n++;
    const id = "rId" + n;
    doc.root.append(
      x("Relationship", {
        Id: id,
        Type: type,
        Target: external ? targetPart : relativeTarget(part, targetPart),
        TargetMode: external ? "External" : undefined,
      }),
    );
    this.write(relsPathOf(part), doc);
    return id;
  }

  private contentTypes(): XDocument {
    return this.read("[Content_Types].xml");
  }

  private ensureDefaultContentType(ext: string, mime: string): void {
    const doc = this.contentTypes();
    const has = doc.root.childrenNamed("Default").some((d) => (d.attr("Extension") ?? "").toLowerCase() === ext.toLowerCase());
    if (!has) {
      insertOrdered(doc.root, x("Default", { Extension: ext.toLowerCase(), ContentType: mime }), ["Default", "Override"]);
      this.write("[Content_Types].xml", doc);
    }
  }

  private uniquePartName(dir: string, base: string, ext: string): string {
    let n = 1;
    while (this.pkg.has(`${dir}/${base}${n}.${ext}`)) n++;
    return `${dir}/${base}${n}.${ext}`;
  }

  // ------------------------------------------------------------- slides
  private presentation(): XDocument {
    return this.read("ppt/presentation.xml");
  }

  /** Slides in presentation order. */
  slides(): SlideRef[] {
    const pres = this.presentation();
    const rels = new Map(this.rels("ppt/presentation.xml").map((r) => [r.id, r]));
    const list = pres.root.path("p:sldIdLst")?.childrenNamed("p:sldId") ?? [];
    return list
      .map((s) => rels.get(s.attr("r:id") ?? ""))
      .filter((r): r is NonNullable<typeof r> => !!r && this.pkg.has(r.target))
      .map((r, index) => {
        const show = this.pkg.text(r.target)?.match(/<p:sld\b[^>]*\bshow="(0|false)"/);
        return { index, file: r.target, hidden: !!show };
      });
  }

  get slideCount(): number {
    return this.slides().length;
  }

  private slideFile(index: number): string {
    const s = this.slides()[index];
    if (!s) throw new PptxError("INVALID_FILE", `Slide index ${index} out of range`);
    return s.file;
  }

  /** Moves a slide to a new position. */
  moveSlide(from: number, to: number): void {
    this.transact({ type: "structure", label: "Move slide" }, () => {
      const pres = this.presentation();
      const lst = pres.root.path("p:sldIdLst");
      const items = lst?.childrenNamed("p:sldId") ?? [];
      if (!lst || from < 0 || from >= items.length) throw new PptxError("INVALID_FILE", "Slide index out of range");
      to = Math.max(0, Math.min(items.length - 1, to));
      if (from === to) return;
      const item = items[from];
      lst.remove(item);
      const rest = lst.childrenNamed("p:sldId");
      const anchor = rest[to];
      lst.insert(item, anchor ? lst.children.indexOf(anchor) : lst.children.length);
      this.write("ppt/presentation.xml", pres);
    });
  }

  /** Hides or shows a slide in slide-show mode. */
  setSlideHidden(index: number, hidden: boolean): void {
    const file = this.slideFile(index);
    this.transact({ type: "slide", index, label: hidden ? "Hide slide" : "Show slide" }, () => {
      const doc = this.read(file);
      doc.root.setAttr("show", hidden ? "0" : undefined);
      this.write(file, doc);
    });
  }

  /** Deletes a slide (and its notes page). */
  deleteSlide(index: number): void {
    const file = this.slideFile(index);
    this.transact({ type: "structure", label: "Delete slide" }, () => {
      const pres = this.presentation();
      const presRels = this.readRels("ppt/presentation.xml");
      const rel = presRels.root.childrenNamed("Relationship").find((r) => resolveRel("ppt/presentation.xml", r.attr("Target") ?? "") === file);
      const rid = rel?.attr("Id");
      const lst = pres.root.path("p:sldIdLst");
      const item = lst?.childrenNamed("p:sldId").find((s) => s.attr("r:id") === rid);
      if (item) lst!.remove(item);
      if (rel) presRels.root.remove(rel);
      this.write("ppt/presentation.xml", pres);
      this.write(relsPathOf("ppt/presentation.xml"), presRels);
      // notes slide belongs to this slide only
      for (const r of this.rels(file)) {
        if (r.type === REL + "notesSlide") this.removePart(r.target);
      }
      this.removePart(file);
    });
  }

  private removePart(part: string): void {
    this.pkg.delete(part);
    this.pkg.delete(relsPathOf(part));
    const ct = this.contentTypes();
    const ov = ct.root.childrenNamed("Override").find((o) => normalizePath(o.attr("PartName") ?? "") === part);
    if (ov) {
      ct.root.remove(ov);
      this.write("[Content_Types].xml", ct);
    }
  }

  private registerSlide(file: string, afterIndex: number): void {
    const ct = this.contentTypes();
    ct.root.append(x("Override", { PartName: "/" + file, ContentType: CT_SLIDE }));
    this.write("[Content_Types].xml", ct);
    const rid = this.addRel("ppt/presentation.xml", REL + "slide", file);
    const pres = this.presentation();
    const lst = pres.root.path("p:sldIdLst") ?? (() => {
      const l = x("p:sldIdLst");
      insertOrdered(pres.root, l, ["p:sldMasterIdLst", "p:notesMasterIdLst", "p:handoutMasterIdLst", "p:sldIdLst", "p:sldSz", "p:notesSz"]);
      return l;
    })();
    const items = lst.childrenNamed("p:sldId");
    const maxId = Math.max(255, ...items.map((s) => Number(s.attr("id")) || 0));
    const el = x("p:sldId", { id: maxId + 1, "r:id": rid });
    const anchor = items[afterIndex + 1];
    lst.insert(el, anchor ? lst.children.indexOf(anchor) : lst.children.length);
    this.write("ppt/presentation.xml", pres);
  }

  /** Duplicates a slide; the copy is inserted after the original. Returns its index. */
  duplicateSlide(index: number): number {
    const file = this.slideFile(index);
    return this.transact({ type: "structure", label: "Duplicate slide" }, () => {
      const newFile = this.uniquePartName("ppt/slides", "slide", "xml");
      this.pkg.setBytes(newFile, this.pkg.bytes(file)!.slice());
      const rels = this.readRels(file);
      for (const r of rels.root.childrenNamed("Relationship")) {
        if (r.attr("Type") === REL + "notesSlide") rels.root.remove(r); // notes are not shared
      }
      this.write(relsPathOf(newFile), rels);
      this.registerSlide(newFile, index);
      return index + 1;
    });
  }

  /**
   * Adds a new slide after `afterIndex` (default: at the end) using the layout
   * of slide `layoutOf` (default: the "blank" layout, or the last slide's).
   * Returns the new slide index.
   */
  addSlide(afterIndex?: number, layoutOf?: number): number {
    const slides = this.slides();
    const after = afterIndex ?? slides.length - 1;
    return this.transact({ type: "structure", label: "Add slide" }, () => {
      let layout: string | undefined;
      if (layoutOf !== undefined) layout = this.rels(this.slideFile(layoutOf)).find((r) => r.type === REL + "slideLayout")?.target;
      if (!layout) {
        // prefer a layout of type "blank"
        for (const p of this.pkg.list()) {
          if (/^ppt\/slideLayouts\/[^/]+\.xml$/.test(p) && /<p:sldLayout\b[^>]*\btype="blank"/.test(this.pkg.text(p) ?? "")) {
            layout = p;
            break;
          }
        }
      }
      if (!layout && slides.length) layout = this.rels(slides[slides.length - 1].file).find((r) => r.type === REL + "slideLayout")?.target;
      if (!layout) layout = this.pkg.list().find((p) => /^ppt\/slideLayouts\/[^/]+\.xml$/.test(p));
      if (!layout) throw new PptxError("INVALID_FILE", "The presentation has no slide layouts");
      const newFile = this.uniquePartName("ppt/slides", "slide", "xml");
      const doc = new XDocument(
        x(
          "p:sld",
          { "xmlns:a": NS_A, "xmlns:r": NS_R, "xmlns:p": NS_P },
          x(
            "p:cSld",
            {},
            x(
              "p:spTree",
              {},
              x("p:nvGrpSpPr", {}, x("p:cNvPr", { id: 1, name: "" }), x("p:cNvGrpSpPr"), x("p:nvPr")),
              x("p:grpSpPr", {}, x("a:xfrm", {}, x("a:off", { x: 0, y: 0 }), x("a:ext", { cx: 0, cy: 0 }), x("a:chOff", { x: 0, y: 0 }), x("a:chExt", { cx: 0, cy: 0 }))),
            ),
          ),
          x("p:clrMapOvr", {}, x("a:masterClrMapping")),
        ),
      );
      this.write(newFile, doc);
      this.addRel(newFile, REL + "slideLayout", layout);
      this.registerSlide(newFile, after);
      return after + 1;
    });
  }

  // ------------------------------------------------------------- shapes
  private spTree(doc: XDocument): XElement {
    const t = doc.root.path("p:cSld", "p:spTree");
    if (!t) throw new PptxError("INVALID_FILE", "Slide has no shape tree");
    return t;
  }

  private findShape(doc: XDocument, id: string): XElement {
    const cNvPr = this.spTree(doc).find((e) => e.name === "p:cNvPr" && e.attr("id") === String(id));
    // cNvPr -> nv*Pr -> shape
    const shape = cNvPr?.parent?.parent;
    if (!shape || !SHAPE_TAGS.includes(shape.name)) throw new PptxError("INVALID_FILE", `Shape ${id} not found`);
    return shape;
  }

  private nextShapeId(doc: XDocument): number {
    let max = 1;
    for (const e of doc.root.findAll((el) => el.name === "p:cNvPr")) max = Math.max(max, Number(e.attr("id")) || 0);
    return max + 1;
  }

  private xfrmOf(shape: XElement, create: boolean): XElement | undefined {
    if (shape.name === "p:graphicFrame") {
      let xf = shape.child("p:xfrm");
      if (!xf && create) {
        xf = x("p:xfrm");
        shape.insertAfterAny(xf, ["p:nvGraphicFramePr"]);
      }
      return xf;
    }
    const prName = shape.name === "p:grpSp" ? "p:grpSpPr" : "p:spPr";
    let pr = shape.child(prName);
    if (!pr) {
      if (!create) return undefined;
      pr = x(prName);
      shape.insertAfterAny(pr, ["p:nvSpPr", "p:nvPicPr", "p:nvCxnSpPr", "p:nvGrpSpPr", "p:blipFill"]);
    }
    let xf = pr.child("a:xfrm");
    if (!xf && create) {
      xf = x("a:xfrm");
      pr.insert(xf, 0);
    }
    return xf;
  }

  /** Bounds stored on the shape itself, or inherited from its layout/master placeholder. */
  private resolveBounds(file: string, shape: XElement): Bounds | null {
    const xf = this.xfrmOf(shape, false);
    const off = xf?.child("a:off");
    const ext = xf?.child("a:ext");
    if (off && ext) {
      return { x: emu2px(off.attr("x")), y: emu2px(off.attr("y")), width: emu2px(ext.attr("cx")), height: emu2px(ext.attr("cy")) };
    }
    const ph = shape.find((e) => e.name === "p:ph");
    if (!ph) return null;
    const layout = this.rels(file).find((r) => r.type === REL + "slideLayout")?.target;
    const master = layout ? this.rels(layout).find((r) => r.type === REL + "slideMaster")?.target : undefined;
    for (const part of [layout, master]) {
      if (!part || !this.pkg.has(part)) continue;
      const doc = this.read(part);
      const type = ph.attr("type");
      const idx = ph.attr("idx");
      const match = doc.root.find((e) => {
        if (e.name !== "p:ph") return false;
        if (idx !== undefined && e.attr("idx") === idx) return true;
        return type !== undefined && e.attr("type") === type;
      });
      const sp = match?.parent?.parent?.parent;
      if (sp) {
        const b = this.resolveBounds(part, sp);
        if (b) return b;
      }
    }
    return null;
  }

  private shapeInfo(file: string, shape: XElement): ShapeInfo {
    const cNvPr = shape.find((e) => e.name === "p:cNvPr");
    const ph = shape.elements[0]?.find((e) => e.name === "p:ph");
    let kind: ShapeKind = "other";
    switch (shape.name) {
      case "p:sp":
        kind = shape.find((e) => e.name === "p:cNvSpPr")?.attr("txBox") === "1" ? "textbox" : "shape";
        break;
      case "p:pic":
        kind = shape.find((e) => e.name === "a:videoFile" || e.name === "a:audioFile") ? "media" : "picture";
        break;
      case "p:grpSp":
        kind = "group";
        break;
      case "p:cxnSp":
        kind = "connector";
        break;
      case "p:graphicFrame": {
        const uri = shape.path("a:graphic", "a:graphicData")?.attr("uri") ?? "";
        kind = uri.endsWith("/table") ? "table" : uri.endsWith("/chart") ? "chart" : uri.endsWith("/diagram") ? "diagram" : "other";
        break;
      }
    }
    const xf = this.xfrmOf(shape, false);
    const rot = Number(xf?.attr("rot") ?? 0) / 60000;
    const spPr = shape.child("p:spPr");
    const fillHex = spPr?.path("a:solidFill", "a:srgbClr")?.attr("val");
    const txBody = shape.child("p:txBody");
    const paragraphs = txBody ? readParagraphs(txBody) : shape.name === "p:sp" ? [] : null;
    return {
      id: cNvPr?.attr("id") ?? "",
      name: cNvPr?.attr("name") ?? "",
      kind,
      bounds: this.resolveBounds(file, shape),
      rotation: rot,
      placeholder: ph ? { type: ph.attr("type"), idx: ph.attr("idx") } : undefined,
      geometry: spPr?.child("a:prstGeom")?.attr("prst") ?? (spPr?.child("a:custGeom") ? "custom" : undefined),
      fill: fillHex ? "#" + fillHex : undefined,
      paragraphs,
      text: paragraphs ? paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join("\n") : "",
    };
  }

  /** Top-level shapes of a slide (in z-order, back to front). */
  getShapes(slideIndex: number): ShapeInfo[] {
    const file = this.slideFile(slideIndex);
    const tree = this.spTree(this.read(file));
    return tree.elements.filter((e) => SHAPE_TAGS.includes(e.name) && e.name !== "mc:AlternateContent").map((e) => this.shapeInfo(file, e));
  }

  getShape(slideIndex: number, id: string): ShapeInfo {
    const file = this.slideFile(slideIndex);
    return this.shapeInfo(file, this.findShape(this.read(file), id));
  }

  /** Edits a slide's XML in one undoable step. */
  private editSlide<T>(slideIndex: number, label: string, fn: (doc: XDocument, file: string) => T): T {
    const file = this.slideFile(slideIndex);
    return this.transact({ type: "slide", index: slideIndex, label }, () => {
      const doc = this.read(file);
      const r = fn(doc, file);
      this.write(file, doc);
      return r;
    });
  }

  /** Moves and/or resizes a shape (px). Missing fields keep their current value. */
  setShapeBounds(slideIndex: number, id: string, bounds: Partial<Bounds>): void {
    this.editSlide(slideIndex, "Move/resize shape", (doc, file) => {
      const shape = this.findShape(doc, id);
      const cur = this.resolveBounds(file, shape) ?? { x: 0, y: 0, width: 100, height: 100 };
      const b = { ...cur, ...bounds };
      if (b.width < 1) b.width = 1;
      if (b.height < 1) b.height = 1;
      const xf = this.xfrmOf(shape, true)!;
      const off = xf.child("a:off") ?? xf.insert(x("a:off"), 0).child("a:off")!;
      const ext = xf.child("a:ext") ?? xf.insertAfterAny(x("a:ext"), ["a:off"]).child("a:ext")!;
      off.setAttr("x", px2emu(b.x)).setAttr("y", px2emu(b.y));
      ext.setAttr("cx", px2emu(b.width)).setAttr("cy", px2emu(b.height));
      if (shape.name === "p:grpSp") {
        // keep children where they are relative to the group
        const chOff = xf.child("a:chOff");
        const chExt = xf.child("a:chExt");
        if (!chOff || !chExt) {
          xf.append(x("a:chOff", { x: px2emu(cur.x), y: px2emu(cur.y) }));
          xf.append(x("a:chExt", { cx: px2emu(cur.width), cy: px2emu(cur.height) }));
        }
      }
    });
  }

  /** Sets the rotation in degrees. */
  setShapeRotation(slideIndex: number, id: string, degrees: number): void {
    this.editSlide(slideIndex, "Rotate shape", (doc) => {
      const shape = this.findShape(doc, id);
      const xf = this.xfrmOf(shape, true)!;
      const rot = Math.round((((degrees % 360) + 360) % 360) * 60000);
      xf.setAttr("rot", rot === 0 ? undefined : rot);
    });
  }

  deleteShape(slideIndex: number, id: string): void {
    this.editSlide(slideIndex, "Delete shape", (doc) => {
      this.findShape(doc, id).detach();
    });
  }

  /** Changes z-order: "front", "back", "forward" or "backward". */
  reorderShape(slideIndex: number, id: string, where: "front" | "back" | "forward" | "backward"): void {
    this.editSlide(slideIndex, "Arrange shape", (doc) => {
      const shape = this.findShape(doc, id);
      const parent = shape.parent!;
      const siblings = parent.elements.filter((e) => SHAPE_TAGS.includes(e.name));
      const pos = siblings.indexOf(shape);
      let target: number;
      if (where === "front") target = siblings.length - 1;
      else if (where === "back") target = 0;
      else if (where === "forward") target = Math.min(siblings.length - 1, pos + 1);
      else target = Math.max(0, pos - 1);
      if (target === pos) return;
      parent.remove(shape);
      // insert so that the shape ends up at position `target` among its shape siblings
      const rest = parent.elements.filter((e) => SHAPE_TAGS.includes(e.name));
      if (target >= rest.length) {
        const last = rest[rest.length - 1];
        parent.insert(shape, last ? parent.children.indexOf(last) + 1 : parent.children.length);
      } else {
        parent.insert(shape, parent.children.indexOf(rest[target]));
      }
    });
  }

  /** Duplicates a shape, offset by (dx, dy) px. Returns the new id. */
  duplicateShape(slideIndex: number, id: string, dx = 16, dy = 16): string {
    return this.editSlide(slideIndex, "Duplicate shape", (doc, file) => {
      const shape = this.findShape(doc, id);
      const copy = shape.clone();
      let next = this.nextShapeId(doc);
      for (const c of [copy, ...copy.findAll((e) => e.name === "p:cNvPr")]) {
        if (c.name === "p:cNvPr") c.setAttr("id", next++);
      }
      const cNvPr = copy.find((e) => e.name === "p:cNvPr");
      const newId = cNvPr?.attr("id") ?? String(next);
      shape.parent!.insert(copy, shape.parent!.children.indexOf(shape) + 1);
      const b = this.resolveBounds(file, shape);
      if (b) {
        const xf = this.xfrmOf(copy, true)!;
        const off = xf.child("a:off") ?? xf.insert(x("a:off"), 0).child("a:off")!;
        const ext = xf.child("a:ext") ?? xf.insertAfterAny(x("a:ext"), ["a:off"]).child("a:ext")!;
        off.setAttr("x", px2emu(b.x + dx)).setAttr("y", px2emu(b.y + dy));
        ext.setAttr("cx", px2emu(b.width)).setAttr("cy", px2emu(b.height));
      }
      return newId;
    });
  }

  /**
   * Replaces the text of a shape. A string is split into paragraphs on "\n";
   * existing paragraph and run formatting is reused where possible.
   */
  setShapeText(slideIndex: number, id: string, text: string | TextParagraph[]): void {
    this.editSlide(slideIndex, "Edit text", (doc) => {
      const shape = this.findShape(doc, id);
      if (shape.name !== "p:sp") throw new PptxError("INVALID_FILE", "Only shapes and text boxes can hold text");
      const txBody = ensureTxBody(shape);
      writeParagraphs(txBody, typeof text === "string" ? text.split(/\r?\n/).map((t) => ({ runs: [{ text: t }] })) : text);
    });
  }

  /** Applies character formatting to all text of a shape. `null` removes a property. */
  setTextStyle(slideIndex: number, id: string, style: { bold?: boolean | null; italic?: boolean | null; underline?: boolean | null; size?: number | null; color?: string | null; font?: string | null; align?: TextParagraph["align"] | null }): void {
    this.editSlide(slideIndex, "Format text", (doc) => {
      const shape = this.findShape(doc, id);
      const txBody = shape.child("p:txBody");
      if (!txBody) return;
      for (const p of txBody.childrenNamed("a:p")) {
        if (style.align !== undefined) {
          const pPr = p.child("a:pPr") ?? p.insert(x("a:pPr"), 0).child("a:pPr")!;
          pPr.setAttr("algn", style.align === null ? undefined : ALIGN_TO_OOXML[style.align]);
        }
        const rPrs = [...p.childrenNamed("a:r").map((r) => r.child("a:rPr") ?? r.insert(x("a:rPr", { lang: "en-US" }), 0).child("a:rPr")!)];
        const end = p.child("a:endParaRPr");
        if (end) rPrs.push(end);
        for (const rPr of rPrs) applyRunStyle(rPr, style);
      }
    });
  }

  /** Sets (or with `null` removes) a solid fill colour of a shape. */
  setShapeFill(slideIndex: number, id: string, color: string | null): void {
    this.editSlide(slideIndex, "Fill colour", (doc) => {
      const shape = this.findShape(doc, id);
      const spPr = shape.child("p:spPr") ?? (() => {
        const pr = x("p:spPr");
        shape.insertAfterAny(pr, ["p:nvSpPr", "p:nvPicPr", "p:nvCxnSpPr", "p:blipFill"]);
        return pr;
      })();
      for (const t of FILL_TAGS) spPr.child(t)?.detach();
      insertOrdered(spPr, color === null ? x("a:noFill") : x("a:solidFill", {}, x("a:srgbClr", { val: cleanHex(color) })), SPPR_ORDER);
    });
  }

  /** Sets the outline colour and width (px); `null` colour removes the outline. */
  setShapeLine(slideIndex: number, id: string, color: string | null, widthPx = 1): void {
    this.editSlide(slideIndex, "Outline", (doc) => {
      const shape = this.findShape(doc, id);
      const spPr = shape.child("p:spPr");
      if (!spPr) return;
      spPr.child("a:ln")?.detach();
      const ln = color === null ? x("a:ln", {}, x("a:noFill")) : x("a:ln", { w: px2emu(widthPx) }, x("a:solidFill", {}, x("a:srgbClr", { val: cleanHex(color) })));
      insertOrdered(spPr, ln, SPPR_ORDER);
    });
  }

  /** Adds a text box. Returns the new shape id. */
  addTextBox(slideIndex: number, bounds: Bounds, text = "", style: Omit<TextRun, "text"> = {}): string {
    return this.editSlide(slideIndex, "Add text box", (doc) => {
      const id = this.nextShapeId(doc);
      const sp = x(
        "p:sp",
        {},
        x("p:nvSpPr", {}, x("p:cNvPr", { id, name: `TextBox ${id - 1}` }), x("p:cNvSpPr", { txBox: 1 }), x("p:nvPr")),
        x("p:spPr", {}, xfrm(bounds), x("a:prstGeom", { prst: "rect" }, x("a:avLst")), x("a:noFill")),
        x("p:txBody", {}, x("a:bodyPr", { wrap: "square", rtlCol: 0 }, x("a:spAutoFit")), x("a:lstStyle")),
      );
      // new RTL text is right-aligned, as PowerPoint does for Hebrew/Arabic text boxes
      writeParagraphs(
        sp.child("p:txBody")!,
        text.split(/\r?\n/).map((t) => ({ runs: [{ ...style, text: t }], align: firstStrongDirection(t) === "rtl" ? ("right" as const) : undefined })),
      );
      this.spTree(doc).append(sp);
      return String(id);
    });
  }

  /** Adds a preset shape (rectangle, ellipse, arrow ...). Returns the new shape id. */
  addShape(slideIndex: number, geometry: PresetGeometry, bounds: Bounds, options: { fill?: string; line?: string; text?: string } = {}): string {
    return this.editSlide(slideIndex, "Add shape", (doc) => {
      const id = this.nextShapeId(doc);
      const spPr = x("p:spPr", {}, xfrm(bounds), x("a:prstGeom", { prst: geometry }, x("a:avLst")));
      if (options.fill) spPr.append(x("a:solidFill", {}, x("a:srgbClr", { val: cleanHex(options.fill) })));
      if (options.line) spPr.append(x("a:ln", { w: 12700 }, x("a:solidFill", {}, x("a:srgbClr", { val: cleanHex(options.line) }))));
      const sp = x(
        "p:sp",
        {},
        x("p:nvSpPr", {}, x("p:cNvPr", { id, name: `${geometry} ${id - 1}` }), x("p:cNvSpPr"), x("p:nvPr")),
        spPr,
        // theme-based style so the shape looks native when no fill is given
        x(
          "p:style",
          {},
          x("a:lnRef", { idx: 2 }, x("a:schemeClr", { val: "accent1" }, x("a:shade", { val: 50000 }))),
          x("a:fillRef", { idx: 1 }, x("a:schemeClr", { val: "accent1" })),
          x("a:effectRef", { idx: 0 }, x("a:schemeClr", { val: "accent1" })),
          x("a:fontRef", { idx: "minor" }, x("a:schemeClr", { val: "lt1" })),
        ),
        x("p:txBody", {}, x("a:bodyPr", { rtlCol: 0, anchor: "ctr" }), x("a:lstStyle")),
      );
      writeParagraphs(sp.child("p:txBody")!, (options.text ?? "").split(/\r?\n/).map((t) => ({ runs: [{ text: t }], align: "center" as const })));
      this.spTree(doc).append(sp);
      return String(id);
    });
  }

  /**
   * Adds a picture. `bounds` defaults to the image's natural size, centred.
   * Returns the new shape id.
   */
  addImage(slideIndex: number, data: Uint8Array, mimeOrExt: string, bounds?: Partial<Bounds>): string {
    const ext = (mimeOrExt.includes("/") ? MIME_EXT[mimeOrExt.toLowerCase()] : mimeOrExt.replace(/^\./, "").toLowerCase()) ?? "";
    const mime = getMimeType(ext);
    if (!ext || !mime.startsWith("image/")) throw new PptxError("INVALID_FILE", `Unsupported image type: ${mimeOrExt}`);
    const file = this.slideFile(slideIndex);
    return this.transact({ type: "slide", index: slideIndex, label: "Add picture" }, () => {
      const media = this.uniquePartName("ppt/media", "image", ext);
      this.pkg.setBytes(media, data);
      this.ensureDefaultContentType(ext, mime);
      const rid = this.addRel(file, REL + "image", media);
      const doc = this.read(file);
      const id = this.nextShapeId(doc);
      const natural = imageSize(data) ?? [320, 240];
      const size = this.slideSize();
      let w = bounds?.width ?? natural[0];
      let h = bounds?.height ?? (bounds?.width ? (bounds.width * natural[1]) / natural[0] : natural[1]);
      const fit = Math.min(1, (size.width * 0.8) / w, (size.height * 0.8) / h);
      if (!bounds?.width) {
        w *= fit;
        h *= fit;
      }
      const b: Bounds = { x: bounds?.x ?? (size.width - w) / 2, y: bounds?.y ?? (size.height - h) / 2, width: w, height: h };
      const pic = x(
        "p:pic",
        {},
        x("p:nvPicPr", {}, x("p:cNvPr", { id, name: `Picture ${id - 1}` }), x("p:cNvPicPr", {}, x("a:picLocks", { noChangeAspect: 1 })), x("p:nvPr")),
        x("p:blipFill", {}, x("a:blip", { "r:embed": rid }), x("a:stretch", {}, x("a:fillRect"))),
        x("p:spPr", {}, xfrm(b), x("a:prstGeom", { prst: "rect" }, x("a:avLst"))),
      );
      this.spTree(doc).append(pic);
      this.write(file, doc);
      return String(id);
    });
  }

  /** Slide size in px. */
  slideSize(): { width: number; height: number } {
    const sz = this.presentation().root.child("p:sldSz");
    return { width: emu2px(sz?.attr("cx") ?? 9144000), height: emu2px(sz?.attr("cy") ?? 6858000) };
  }

  /** Speaker notes (plain text) of a slide. */
  getNotes(slideIndex: number): string {
    const rel = this.rels(this.slideFile(slideIndex)).find((r) => r.type === REL + "notesSlide");
    if (!rel || !this.pkg.has(rel.target)) return "";
    const doc = this.read(rel.target);
    const body = doc.root.findAll((e) => e.name === "p:sp").find((sp) => sp.find((e) => e.name === "p:ph")?.attr("type") === "body");
    return body?.child("p:txBody") ? readParagraphs(body.child("p:txBody")!).map((p) => p.runs.map((r) => r.text).join("")).join("\n") : "";
  }

  /** Serialises the (edited) presentation to .pptx bytes. */
  toPptx(): Uint8Array {
    return this.pkg.toZip();
  }
}

// ---------------------------------------------------------------------------
// text helpers
// ---------------------------------------------------------------------------
const ALIGN_FROM_OOXML: Record<string, TextParagraph["align"]> = { l: "left", ctr: "center", r: "right", just: "justify", dist: "justify" };
const ALIGN_TO_OOXML: Record<NonNullable<TextParagraph["align"]>, string> = { left: "l", center: "ctr", right: "r", justify: "just" };
const MIME_EXT: Record<string, string> = {
  "image/png": "png",
  "image/jpeg": "jpeg",
  "image/jpg": "jpeg",
  "image/gif": "gif",
  "image/svg+xml": "svg",
  "image/bmp": "bmp",
  "image/webp": "webp",
  "image/tiff": "tiff",
};

function xfrm(b: Bounds): XElement {
  return x("a:xfrm", {}, x("a:off", { x: px2emu(b.x), y: px2emu(b.y) }), x("a:ext", { cx: px2emu(b.width), cy: px2emu(b.height) }));
}

function ensureTxBody(sp: XElement): XElement {
  let tx = sp.child("p:txBody");
  if (!tx) {
    tx = x("p:txBody", {}, x("a:bodyPr", { rtlCol: 0, anchor: "ctr" }), x("a:lstStyle"));
    sp.insertAfterAny(tx, ["p:nvSpPr", "p:spPr", "p:style"]);
  }
  if (!tx.child("a:bodyPr")) tx.insert(x("a:bodyPr"), 0);
  return tx;
}

function readRunStyle(rPr: XElement | undefined): Omit<TextRun, "text"> {
  if (!rPr) return {};
  const s: Omit<TextRun, "text"> = {};
  const b = rPr.attr("b");
  const i = rPr.attr("i");
  const u = rPr.attr("u");
  const sz = rPr.attr("sz");
  if (b !== undefined) s.bold = b === "1" || b === "true";
  if (i !== undefined) s.italic = i === "1" || i === "true";
  if (u !== undefined) s.underline = u !== "none";
  if (sz !== undefined) s.size = Number(sz) / 100;
  const clr = rPr.path("a:solidFill", "a:srgbClr")?.attr("val");
  if (clr) s.color = "#" + clr;
  const font = rPr.child("a:latin")?.attr("typeface");
  if (font) s.font = font;
  return s;
}

export function readParagraphs(txBody: XElement): TextParagraph[] {
  return txBody.childrenNamed("a:p").map((p) => {
    const runs: TextRun[] = [];
    for (const c of p.elements) {
      if (c.name === "a:r" || c.name === "a:fld") runs.push({ text: c.child("a:t")?.text ?? "", ...readRunStyle(c.child("a:rPr")) });
      else if (c.name === "a:br") runs.push({ text: "\u000b" });
    }
    const pPr = p.child("a:pPr");
    const para: TextParagraph = { runs };
    const algn = pPr?.attr("algn");
    if (algn) para.align = ALIGN_FROM_OOXML[algn];
    const lvl = pPr?.attr("lvl");
    if (lvl) para.level = Number(lvl);
    const rtl = pPr?.attr("rtl");
    if (rtl !== undefined) para.rtl = rtl === "1" || rtl === "true";
    return para;
  });
}

function applyRunStyle(rPr: XElement, style: { bold?: boolean | null; italic?: boolean | null; underline?: boolean | null; size?: number | null; color?: string | null; font?: string | null }): void {
  if (style.bold !== undefined) rPr.setAttr("b", style.bold === null ? undefined : style.bold ? 1 : 0);
  if (style.italic !== undefined) rPr.setAttr("i", style.italic === null ? undefined : style.italic ? 1 : 0);
  if (style.underline !== undefined) rPr.setAttr("u", style.underline === null ? undefined : style.underline ? "sng" : "none");
  if (style.size !== undefined) rPr.setAttr("sz", style.size === null ? undefined : Math.round(Math.max(1, Math.min(4000, style.size)) * 100));
  if (style.color !== undefined) {
    for (const t of FILL_TAGS) rPr.child(t)?.detach();
    if (style.color !== null) insertOrdered(rPr, x("a:solidFill", {}, x("a:srgbClr", { val: cleanHex(style.color) })), RPR_ORDER);
  }
  if (style.font !== undefined) {
    rPr.child("a:latin")?.detach();
    rPr.child("a:cs")?.detach();
    if (style.font !== null) {
      insertOrdered(rPr, x("a:latin", { typeface: style.font }), RPR_ORDER);
      insertOrdered(rPr, x("a:cs", { typeface: style.font }), RPR_ORDER);
    }
  }
}

/** Replaces the paragraphs of a txBody, keeping existing pPr/rPr as templates. */
function writeParagraphs(txBody: XElement, paragraphs: TextParagraph[]): void {
  const old = txBody.childrenNamed("a:p");
  for (const p of old) txBody.remove(p);
  if (paragraphs.length === 0) paragraphs = [{ runs: [] }];
  paragraphs.forEach((para, i) => {
    const tpl = old[Math.min(i, old.length - 1)];
    const pPr = tpl?.child("a:pPr")?.clone();
    const tplRPr = tpl?.childrenNamed("a:r")[0]?.child("a:rPr") ?? tpl?.child("a:endParaRPr");
    const p = x("a:p");
    if (pPr) p.append(pPr);
    // direction: explicit value, else keep the template's, else detect from the text
    let rtl = para.rtl;
    if (rtl === undefined && p.child("a:pPr")?.attr("rtl") === undefined) {
      rtl = firstStrongDirection(para.runs.map((r) => r.text).join("")) === "rtl" ? true : undefined;
    }
    if (para.align || para.level !== undefined || rtl !== undefined) {
      const pp = p.child("a:pPr") ?? p.insert(x("a:pPr"), 0).child("a:pPr")!;
      if (para.align) pp.setAttr("algn", ALIGN_TO_OOXML[para.align]);
      if (para.level !== undefined) pp.setAttr("lvl", para.level || undefined);
      if (rtl !== undefined) pp.setAttr("rtl", rtl ? 1 : 0);
    }
    for (const run of para.runs) {
      const parts = run.text.split("\u000b");
      parts.forEach((part, j) => {
        if (j > 0) {
          const brPr = tplRPr ? tplRPr.clone() : x("a:rPr", { lang: "en-US" });
          brPr.name = "a:rPr";
          p.append(x("a:br", {}, brPr));
        }
        if (part === "" && parts.length > 1) return;
        const rPr = tplRPr ? tplRPr.clone() : x("a:rPr", { lang: "en-US", dirty: 0 });
        rPr.name = "a:rPr";
        applyRunStyle(rPr, run);
        const t = x("a:t");
        t.children = [new XText(escapeText(part))];
        p.append(x("a:r", {}, rPr, t));
      });
    }
    const end = tplRPr ? tplRPr.clone() : x("a:endParaRPr", { lang: "en-US", dirty: 0 });
    end.name = "a:endParaRPr";
    p.append(end);
    // a:p must come after a:bodyPr / a:lstStyle
    txBody.append(p);
  });
}

function escapeText(s: string): string {
  return s.replace(/[&<>]/g, (c) => (c === "&" ? "&amp;" : c === "<" ? "&lt;" : "&gt;"));
}

export { extractFileExtension };
