/**
 * Presentation-level rendering: resolves slides, layouts, masters and themes
 * and turns every slide into an HTML string. DOM-free, so it also runs in Web
 * Workers and Node.
 */
import { PptxPackage, resolveTarget } from "../core/package";
import { asArray, getTextByPathList, type XmlNode } from "../core/xml";
import { base64ArrayBuffer } from "../utils/base64";
import { RenderContext, type NodeTables, type RenderOptions, type ResourceRef, type WarpObj } from "./context";
import { getBackground, getSlideBackgroundFill } from "./background";
import { processNodesInSlide } from "./nodes";
import { slideFactor } from "./units";

const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const REL_SLIDE = REL_NS + "slide";
const REL_LAYOUT = REL_NS + "slideLayout";
const REL_MASTER = REL_NS + "slideMaster";
const REL_THEME = REL_NS + "theme";
const REL_NOTES = REL_NS + "notesSlide";
const REL_DIAGRAM_DRAWING = "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing";
const CT_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const CT_LAYOUT = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml";

export interface SlideInfo {
  /** 1-based position in the presentation. */
  number: number;
  /** Package path, e.g. `ppt/slides/slide3.xml`. */
  file: string;
  /** Slide is hidden in PowerPoint's slide show (`show="0"`). */
  hidden: boolean;
}

export interface RenderedSlide extends SlideInfo {
  /** HTML of the slide (a single `<div class="slide">` element). */
  html: string;
  /** Plain-text speaker notes, if any. */
  notes: string;
  /** Set when the slide could not be rendered; `html` then holds a placeholder. */
  error?: Error;
}

export interface RenderedPresentation {
  /** Slide size in CSS pixels. */
  width: number;
  height: number;
  slides: RenderedSlide[];
  /** CSS for classes generated while rendering; inject once per document. */
  css: string;
  /** `data:` URL of the embedded thumbnail, if present. */
  thumbnail?: string;
  /** Office application version that saved the file (e.g. 16). */
  appVersion: number;
  /** Total render time in milliseconds. */
  renderTime: number;
}

export interface Relationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

/** Reads `_rels/<part>.rels` and resolves targets relative to `part`. */
export function readRelationships(pkg: PptxPackage, part: string): Relationship[] {
  const slash = part.lastIndexOf("/");
  const relsPath = part.slice(0, slash + 1) + "_rels/" + part.slice(slash + 1) + ".rels";
  const rels = pkg.xmlFresh(relsPath);
  const list = asArray(getTextByPathList(rels, ["Relationships", "Relationship"]));
  return list
    .filter((r: any) => r && r.attrs && r.attrs.Id !== undefined)
    .map((r: any) => {
      const external = r.attrs.TargetMode === "External";
      const target = String(r.attrs.Target ?? "");
      return {
        id: r.attrs.Id,
        type: String(r.attrs.Type ?? ""),
        target: external ? target : resolveTarget(part, target),
        external,
      };
    });
}

function toResObj(rels: Relationship[], skipTypes: string[] = []): Record<string, ResourceRef> {
  const obj: Record<string, ResourceRef> = {};
  for (const r of rels) {
    if (skipTypes.includes(r.type)) continue;
    obj[r.id] = { type: r.type.replace(REL_NS, ""), target: r.target };
  }
  return obj;
}

/** Returns the slides in presentation order (`p:sldIdLst`), with a fallback to content types. */
export function listSlides(pkg: PptxPackage): SlideInfo[] {
  const pres = pkg.xmlFresh("ppt/presentation.xml");
  const rels = readRelationships(pkg, "ppt/presentation.xml");
  const byId = new Map(rels.map((r) => [r.id, r]));
  const out: SlideInfo[] = [];
  for (const sld of asArray(getTextByPathList(pres, ["p:presentation", "p:sldIdLst", "p:sldId"]))) {
    const rid = sld?.attrs?.["r:id"];
    const rel = rid ? byId.get(rid) : undefined;
    if (rel && rel.type === REL_SLIDE && pkg.has(rel.target)) {
      out.push({ number: out.length + 1, file: rel.target, hidden: false });
    }
  }
  if (out.length === 0) {
    // Fallback (v1 behaviour): slides from [Content_Types].xml sorted by number.
    const ct = pkg.xmlFresh("[Content_Types].xml");
    const files = asArray(getTextByPathList(ct, ["Types", "Override"]))
      .filter((o: any) => o?.attrs?.ContentType === CT_SLIDE)
      .map((o: any) => String(o.attrs.PartName).replace(/^\//, ""))
      .sort((a: string, b: string) => slideNum(a) - slideNum(b));
    files.forEach((f: string) => out.push({ number: out.length + 1, file: f, hidden: false }));
  }
  for (const s of out) {
    const show = getTextByPathList(pkg.xml(s.file), ["p:sld", "attrs", "show"]);
    s.hidden = show === "0" || show === "false";
  }
  return out;
}

function slideNum(file: string): number {
  const m = /(\d+)\.xml$/.exec(file);
  return m ? Number(m[1]) : 0;
}

/** Layout files listed in the content types (kept for API compatibility). */
export function listLayouts(pkg: PptxPackage): string[] {
  const ct = pkg.xmlFresh("[Content_Types].xml");
  return asArray(getTextByPathList(ct, ["Types", "Override"]))
    .filter((o: any) => o?.attrs?.ContentType === CT_LAYOUT)
    .map((o: any) => String(o.attrs.PartName).replace(/^\//, ""));
}

/** Reads app version, slide size and the default text style. */
export function readPresentationInfo(ctx: RenderContext): { width: number; height: number } {
  const pkg = ctx.pkg;
  const app = pkg.xmlFresh("docProps/app.xml");
  const appVersion = parseInt(getTextByPathList(app, ["Properties", "AppVersion"]));
  ctx.appVersion = Number.isFinite(appVersion) ? appVersion : 0; // issue #42: app.xml is optional

  const pres = pkg.xmlFresh("ppt/presentation.xml");
  const sz = getTextByPathList(pres, ["p:presentation", "p:sldSz", "attrs"]) ?? {};
  const cx = parseInt(sz.cx) || 9144000; // default 10in x 7.5in
  const cy = parseInt(sz.cy) || 6858000;
  ctx.defaultTextStyle = getTextByPathList(pres, ["p:presentation", "p:defaultTextStyle"]);
  const inc = ctx.options.incSlide ?? {};
  ctx.slideWidth = (cx * slideFactor + (inc.width ?? 0)) | 0;
  ctx.slideHeight = (cy * slideFactor + (inc.height ?? 0)) | 0;
  return { width: ctx.slideWidth, height: ctx.slideHeight };
}

export function indexNodes(content: XmlNode | null): NodeTables {
  const tables: NodeTables = { idTable: {}, idxTable: {}, typeTable: {} };
  if (!content) return tables;
  const root = content[Object.keys(content)[0]];
  const spTree = getTextByPathList(root, ["p:cSld", "p:spTree"]);
  if (!spTree) return tables;
  for (const key of Object.keys(spTree)) {
    if (key === "p:nvGrpSpPr" || key === "p:grpSpPr" || key === "attrs") continue;
    for (const node of asArray(spTree[key])) {
      const nv = node?.["p:nvSpPr"];
      const id = getTextByPathList(nv, ["p:cNvPr", "attrs", "id"]);
      const idx = getTextByPathList(nv, ["p:nvPr", "p:ph", "attrs", "idx"]);
      const type = getTextByPathList(nv, ["p:nvPr", "p:ph", "attrs", "type"]);
      if (id !== undefined) tables.idTable[id] = node;
      if (idx !== undefined) tables.idxTable[idx] = node;
      if (type !== undefined) tables.typeTable[type] = node;
    }
  }
  return tables;
}

/** Builds the per-slide context (layout, master, theme, diagram, relationships). */
export function buildWarpObj(ctx: RenderContext, slideFile: string): WarpObj {
  const pkg = ctx.pkg;
  // Step 1: slide relationships -> layout
  const slideRels = readRelationships(pkg, slideFile);
  const layoutFile = slideRels.find((r) => r.type === REL_LAYOUT)?.target ?? "";
  const diagramFile = slideRels.find((r) => r.type === REL_DIAGRAM_DRAWING)?.target ?? "";
  const slideResObj = toResObj(slideRels, [REL_LAYOUT]);

  const slideLayoutContent = layoutFile ? pkg.xmlFresh(layoutFile) : null;
  const slideLayoutTables = indexNodes(slideLayoutContent);
  const clrOvr = getTextByPathList(slideLayoutContent, ["p:sldLayout", "p:clrMapOvr", "a:overrideClrMapping", "attrs"]);

  // Step 2: layout relationships -> master
  const layoutRels = layoutFile ? readRelationships(pkg, layoutFile) : [];
  const masterFile = layoutRels.find((r) => r.type === REL_MASTER)?.target ?? "";
  const layoutResObj = toResObj(layoutRels, [REL_MASTER]);
  const slideMasterContent = masterFile ? pkg.xmlFresh(masterFile) : null;
  const slideMasterTextStyles = getTextByPathList(slideMasterContent, ["p:sldMaster", "p:txStyles"]);
  const slideMasterTables = indexNodes(slideMasterContent);

  // Step 3: master relationships -> theme
  const masterRels = masterFile ? readRelationships(pkg, masterFile) : [];
  const themeFile = masterRels.find((r) => r.type === REL_THEME)?.target ?? "";
  const masterResObj = toResObj(masterRels, [REL_THEME]);
  const themeContent = themeFile ? pkg.xmlFresh(themeFile) : null;
  const themeResObj = themeFile ? toResObj(readRelationships(pkg, themeFile)) : {};

  // Step 4: SmartArt drawing (dsp: namespace mapped to p: as in v1)
  let digramFileContent: XmlNode | null = null;
  let diagramResObj: Record<string, ResourceRef> = {};
  if (diagramFile) {
    const raw = pkg.xmlFresh(diagramFile);
    if (raw) digramFileContent = JSON.parse(JSON.stringify(raw).replace(/dsp:/g, "p:"));
    diagramResObj = toResObj(readRelationships(pkg, diagramFile));
  }

  const slideContent = pkg.xmlFresh(slideFile);

  return {
    ctx,
    zip: pkg,
    slideFile,
    slideLayoutContent,
    slideLayoutTables,
    slideMasterContent,
    slideMasterTables,
    slideContent,
    slideResObj,
    slideMasterTextStyles,
    layoutResObj,
    masterResObj,
    themeContent,
    themeResObj,
    digramFileContent,
    diagramResObj,
    defaultTextStyle: ctx.defaultTextStyle,
    layoutClrOverride: clrOvr,
  };
}

/** Renders the shapes of one slide (background + spTree) to HTML. */
export function renderSlide(ctx: RenderContext, info: SlideInfo, index: number): RenderedSlide {
  ctx.currentSlide = info.number;
  const size = { width: ctx.slideWidth, height: ctx.slideHeight };
  const notes = readNotes(ctx.pkg, info.file);
  const attrs =
    `class='slide${info.hidden ? " slide-hidden" : ""}' data-slide='${info.number}' ` +
    `data-file='${info.file}' id='${ctx.idPrefix}-slide-${slideNum(info.file)}'`;
  try {
    const warpObj = buildWarpObj(ctx, info.file);
    let bgResult = "";
    if (ctx.processFullTheme === true) bgResult = getBackground(warpObj, size, index);
    let bgColor = "";
    if (ctx.processFullTheme === "colorsAndImageOnly") bgColor = getSlideBackgroundFill(warpObj, index);

    let html = `<div ${attrs} style='width:${size.width}px; height:${size.height}px;${bgColor}'>` + bgResult;
    const nodes = getTextByPathList(warpObj.slideContent, ["p:sld", "p:cSld", "p:spTree"]);
    if (nodes) {
      for (const nodeKey of Object.keys(nodes)) {
        for (const node of asArray(nodes[nodeKey])) {
          html += processNodesInSlide(nodeKey, node, nodes, warpObj, "slide");
        }
      }
    }
    html += bgResult ? "</div></div>" : "</div>";
    return { ...info, html, notes };
  } catch (e) {
    const err = e instanceof Error ? e : new Error(String(e));
    ctx.warn(`Slide ${info.number} (${info.file}) could not be rendered: ${err.message}`, err);
    const html =
      `<div ${attrs} style='width:${size.width}px; height:${size.height}px;'>` +
      `<div class='slide-error'>Slide ${info.number} could not be rendered.</div></div>`;
    return { ...info, html, notes, error: err };
  }
}

/** Extracts the speaker notes text of a slide. */
export function readNotes(pkg: PptxPackage, slideFile: string): string {
  const rel = readRelationships(pkg, slideFile).find((r) => r.type === REL_NOTES);
  if (!rel) return "";
  const notes = pkg.xml(rel.target);
  const spTree = getTextByPathList(notes, ["p:notes", "p:cSld", "p:spTree"]);
  const paragraphs: string[] = [];
  for (const sp of asArray(spTree?.["p:sp"])) {
    const phType = getTextByPathList(sp, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]);
    if (phType !== "body") continue;
    for (const p of asArray(getTextByPathList(sp, ["p:txBody", "a:p"]))) {
      const runs = asArray(p?.["a:r"]).map((r: any) => (typeof r?.["a:t"] === "string" ? r["a:t"] : ""));
      paragraphs.push(runs.join(""));
    }
  }
  return paragraphs.join("\n").trim();
}

/** Generates the CSS for the classes registered in `ctx.styleTable`. */
export function genGlobalCSS(ctx: RenderContext): string {
  let css = "";
  for (const key of Object.keys(ctx.styleTable)) {
    const s = ctx.styleTable[key];
    css += " ." + s.name + (s.suffix ? s.suffix : "") + "{" + s.text + "}\n";
  }
  return css;
}

export interface RenderPresentationOptions extends RenderOptions {
  /** Only render these slide numbers (1-based). Default: all. */
  slides?: number[];
  /** Skip slides marked hidden (default `false`). */
  skipHiddenSlides?: boolean;
  /** Called after each slide with (done, total). */
  onProgress?: (done: number, total: number) => void;
}

/** Renders a whole presentation to HTML strings. */
export function renderPresentation(pkg: PptxPackage, options: RenderPresentationOptions = {}): RenderedPresentation {
  const started = Date.now();
  const ctx = new RenderContext(pkg, options);
  const { width, height } = readPresentationInfo(ctx);
  ctx.tableStyles = pkg.xmlFresh("ppt/tableStyles.xml");

  let thumbnail: string | undefined;
  const thumb = pkg.bytes("docProps/thumbnail.jpeg");
  if (thumb) thumbnail = "data:image/jpeg;base64," + base64ArrayBuffer(thumb);

  let infos = listSlides(pkg);
  if (options.skipHiddenSlides) infos = infos.filter((s) => !s.hidden);
  if (options.slides) {
    const wanted = new Set(options.slides);
    infos = infos.filter((s) => wanted.has(s.number));
  }
  const slides: RenderedSlide[] = [];
  infos.forEach((info, i) => {
    slides.push(renderSlide(ctx, info, info.number - 1));
    options.onProgress?.(i + 1, infos.length);
  });
  return {
    width,
    height,
    slides,
    css: genGlobalCSS(ctx),
    thumbnail,
    appVersion: ctx.appVersion,
    renderTime: Date.now() - started,
  };
}
