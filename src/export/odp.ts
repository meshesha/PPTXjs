/**
 * Export to OpenDocument Presentation (.odp, ODF 1.3).
 *
 * The converter walks the slide trees with the same style resolution as the
 * HTML renderer (theme colours, placeholder inheritance, master text styles),
 * so colours, fonts and positions match what the viewer shows. Supported:
 * text boxes and placeholders (runs with font/size/colour/bold/italic/
 * underline, alignment, bullets/numbering as text, indents), preset shapes
 * (mapped to ODF custom shapes), custom geometry (draw:path), pictures,
 * tables, charts (as embedded SVG), groups, connectors, rotation, slide
 * backgrounds (colour or picture), layout/master decorations and speaker notes.
 */
import { strToU8, zipSync } from "fflate";
import { PptxPackage, extractFileExtension, getMimeType } from "../core/package";
import { asArray, decodeXmlEntities, getTextByPathList } from "../core/xml";
import { PresentationDocument } from "../model/document";
import { RenderContext } from "../render/context";
import { buildWarpObj, listSlides, readNotes, readPresentationInfo } from "../render/presentation";
import { getShapeFill } from "../render/fill";
import { getBorder } from "../render/border";
import { getSolidFill } from "../render/color";
import { getSlideBackgroundFill } from "../render/background";
import { genBuChar, genSpanElement, getHorizontalAlign, getPregraphDir, getPregraphMargn, getVerticalAlign } from "../render/text";
import { custGeomToSvgPaths } from "../render/custgeom";
import { processSpNode } from "../render/nodes";
import { ENHANCED_SIZE, geometryFromSvgMarkup, parseSvgPath, toEnhancedPath, type GeometryPart } from "./svg-path";
import { getNumTypeNum } from "../render/numbering";
import { parseChart } from "../charts/model";
import { renderChartSvg } from "../charts/svg";
import { themePalette } from "../render/chart";
import { slideFactor } from "../render/units";
import { PptxError } from "../errors";

export const ODP_MIME = "application/vnd.oasis.opendocument.presentation";

export interface OdpExportOptions {
  /** Include shapes from slide layouts/masters (logos, decorations). Default true. */
  includeMasterShapes?: boolean;
  /** Include speaker notes. Default true. */
  includeNotes?: boolean;
  /** Skip slides marked hidden. Default false. */
  skipHiddenSlides?: boolean;
  /** Direction policy for paragraphs without an explicit direction (see RenderOptions.textDirection). */
  textDirection?: "file" | "auto" | "rtl";
}

// ---------------------------------------------------------------------------
// helpers
// ---------------------------------------------------------------------------
const esc = (s: string) => s.replace(/[&<>"']/g, (c) => (c === "&" ? "&amp;" : c === "<" ? "&lt;" : c === ">" ? "&gt;" : c === '"' ? "&quot;" : "&apos;"));
/** px (96 dpi) -> cm string */
const cm = (px: number) => (Number.isFinite(px) ? Math.round(((px * 2.54) / 96) * 1000) / 1000 : 0) + "cm";
const num = (v: unknown) => {
  const n = parseFloat(String(v));
  return Number.isFinite(n) ? n : NaN;
};

interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
  rot: number;
  flipH: boolean;
  flipV: boolean;
}
type MapBox = (b: Box) => Box;
const identity: MapBox = (b) => b;

/** Parses the first colour of a CSS text ("#rrggbb", "#rrggbbaa" or rgb()). */
function cssColor(css: unknown): { color: string; opacity?: number } | null {
  if (typeof css !== "string") return null;
  let m = /#([0-9a-fA-F]{8})\b/.exec(css);
  if (m) {
    const a = parseInt(m[1].slice(6), 16) / 255;
    return { color: "#" + m[1].slice(0, 6), opacity: a };
  }
  m = /#([0-9a-fA-F]{6})\b/.exec(css);
  if (m) return { color: "#" + m[1] };
  const r = /rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*(?:,\s*([\d.]+))?\s*\)/.exec(css);
  if (r) {
    const hex = "#" + [r[1], r[2], r[3]].map((v) => Number(v).toString(16).padStart(2, "0")).join("");
    return { color: hex, opacity: r[4] !== undefined ? Number(r[4]) : undefined };
  }
  return null;
}

/** Parses a CSS declaration list into a map. */
function cssDecls(css: string): Record<string, string> {
  const out: Record<string, string> = {};
  for (const part of decodeXmlEntities(css).split(";")) {
    const i = part.indexOf(":");
    if (i > 0) out[part.slice(0, i).trim().toLowerCase()] = part.slice(i + 1).trim();
  }
  return out;
}

const PRESET_TO_ODF: Record<string, string> = {
  rect: "rectangle",
  roundRect: "round-rectangle",
  ellipse: "ellipse",
  triangle: "isosceles-triangle",
  rtTriangle: "right-triangle",
  diamond: "diamond",
  parallelogram: "parallelogram",
  trapezoid: "trapezoid",
  pentagon: "pentagon",
  hexagon: "hexagon",
  octagon: "octagon",
  star4: "star4",
  star5: "star5",
  star8: "star8",
  star24: "star24",
  rightArrow: "right-arrow",
  leftArrow: "left-arrow",
  upArrow: "up-arrow",
  downArrow: "down-arrow",
  leftRightArrow: "left-right-arrow",
  upDownArrow: "up-down-arrow",
  heart: "heart",
  cloud: "cloud",
  plus: "cross",
  can: "can",
  cube: "cube",
  smileyFace: "smiley",
  sun: "sun",
  moon: "moon",
  lightningBolt: "lightning",
  chevron: "chevron",
  homePlate: "pentagon-right",
  donut: "ring",
  noSmoking: "forbidden",
  frame: "frame",
  blockArc: "block-arc",
  wedgeRectCallout: "rectangular-callout",
  wedgeRoundRectCallout: "round-rectangular-callout",
  wedgeEllipseCallout: "round-callout",
  cloudCallout: "cloud-callout",
  flowChartProcess: "flowchart-process",
  flowChartDecision: "flowchart-decision",
  flowChartTerminator: "flowchart-terminator",
  flowChartAlternateProcess: "flowchart-alternate-process",
  flowChartDocument: "flowchart-document",
  flowChartConnector: "flowchart-connector",
  flowChartInputOutput: "flowchart-data",
  flowChartPredefinedProcess: "flowchart-predefined-process",
  flowChartManualInput: "flowchart-manual-input",
  flowChartPreparation: "flowchart-preparation",
};

// ---------------------------------------------------------------------------
// style registry
// ---------------------------------------------------------------------------
class Styles {
  private maps = new Map<string, Map<string, string>>();
  private defs: string[] = [];
  readonly fillImages: string[] = [];

  /** Returns a style name for (family, properties-xml), creating it once. */
  get(family: string, prefix: string, props: string, parent?: string): string {
    let map = this.maps.get(family);
    if (!map) this.maps.set(family, (map = new Map()));
    const key = props + "|" + (parent ?? "");
    let name = map.get(key);
    if (!name) {
      name = prefix + (map.size + 1);
      map.set(key, name);
      this.defs.push(
        `<style:style style:name="${name}" style:family="${family}"${parent ? ` style:parent-style-name="${parent}"` : ""}>${props}</style:style>`,
      );
    }
    return name;
  }

  xml(): string {
    return this.defs.join("");
  }
}

class Media {
  readonly files: Array<{ path: string; data: Uint8Array; mime: string }> = [];
  private byKey = new Map<string, string>();
  add(key: string, data: Uint8Array, ext: string): string {
    const known = this.byKey.get(key);
    if (known) return known;
    const path = `Pictures/img${this.files.length + 1}.${ext.toLowerCase()}`;
    this.files.push({ path, data, mime: getMimeType(ext) || "application/octet-stream" });
    this.byKey.set(key, path);
    return path;
  }
}

// ---------------------------------------------------------------------------
// converter
// ---------------------------------------------------------------------------
class OdpWriter {
  private readonly styles = new Styles();
  private readonly media = new Media();
  private readonly fillImageDefs: string[] = [];
  private shapeCounter = 0;
  private pageW = 0;
  private pageH = 0;

  constructor(
    private readonly pkg: PptxPackage,
    private readonly opts: OdpExportOptions,
  ) {}

  build(): Uint8Array {
    const ctx = new RenderContext(this.pkg, { themeProcess: true, textDirection: this.opts.textDirection });
    const { width, height } = readPresentationInfo(ctx);
    this.pageW = width;
    this.pageH = height;
    ctx.tableStyles = this.pkg.xmlFresh("ppt/tableStyles.xml");
    let slides = listSlides(this.pkg);
    if (this.opts.skipHiddenSlides) slides = slides.filter((s) => !s.hidden);

    let pages = "";
    slides.forEach((info, i) => {
      ctx.currentSlide = info.number;
      let warpObj: any;
      try {
        warpObj = buildWarpObj(ctx, info.file);
      } catch (e) {
        throw new PptxError("EXPORT_FAILED", `Slide ${info.number}: ${(e as Error).message}`, e);
      }
      const dp = this.pageStyle(warpObj, info.number - 1, info.hidden);
      let body = this.bgPicture;
      if (this.opts.includeMasterShapes !== false) body += this.masterShapes(warpObj);
      const tree = getTextByPathList(warpObj.slideContent, ["p:sld", "p:cSld", "p:spTree"]);
      body += this.tree(tree, warpObj, "slide", identity);
      let notes = "";
      if (this.opts.includeNotes !== false) {
        const text = readNotes(this.pkg, info.file);
        if (text) {
          const paras = text.split("\n").map((l) => `<text:p>${esc(l)}</text:p>`).join("");
          notes =
            `<presentation:notes draw:style-name="${dp}"><draw:frame presentation:class="notes" svg:x="2cm" svg:y="13cm" svg:width="17cm" svg:height="12cm">` +
            `<draw:text-box>${paras}</draw:text-box></draw:frame></presentation:notes>`;
        }
      }
      pages += `<draw:page draw:name="page${i + 1}" draw:style-name="${dp}" draw:master-page-name="Default">${body}${notes}</draw:page>`;
    });

    const content =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<office:document-content ${NS} office:version="1.3">` +
      `<office:automatic-styles>${this.styles.xml()}</office:automatic-styles>` +
      `<office:body><office:presentation>${pages}</office:presentation></office:body></office:document-content>`;
    const orientation = width >= height ? "landscape" : "portrait";
    const stylesXml =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<office:document-styles ${NS} office:version="1.3">` +
      `<office:styles>${this.fillImageDefs.join("")}` +
      `<style:default-style style:family="graphic"><style:graphic-properties draw:shadow="hidden" svg:stroke-color="#3465a4" draw:fill-color="#729fcf"/>` +
      `<style:paragraph-properties style:text-autospace="ideograph-alpha" style:writing-mode="lr-tb"/>` +
      `<style:text-properties fo:font-size="18pt" style:font-name="Calibri" fo:font-family="Calibri" fo:language="en" fo:country="US"/></style:default-style>` +
      `</office:styles>` +
      `<office:automatic-styles><style:page-layout style:name="PM1"><style:page-layout-properties fo:margin-top="0cm" fo:margin-bottom="0cm" fo:margin-left="0cm" fo:margin-right="0cm" fo:page-width="${cm(width)}" fo:page-height="${cm(height)}" style:print-orientation="${orientation}"/></style:page-layout>` +
      `<style:style style:name="Mdp1" style:family="drawing-page"><style:drawing-page-properties draw:background-size="full" draw:fill="none"/></style:style></office:automatic-styles>` +
      `<office:master-styles><draw:layer-set><draw:layer draw:name="layout"/><draw:layer draw:name="background"/><draw:layer draw:name="backgroundobjects"/><draw:layer draw:name="controls"/><draw:layer draw:name="measurelines"/></draw:layer-set><style:master-page style:name="Default" style:page-layout-name="PM1" draw:style-name="Mdp1"/></office:master-styles>` +
      `</office:document-styles>`;
    const title = this.title();
    const meta =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<office:document-meta ${NS} office:version="1.3"><office:meta><meta:generator>PPTXjs</meta:generator>` +
      (title ? `<dc:title>${esc(title)}</dc:title>` : "") +
      `<meta:creation-date>${new Date().toISOString().replace(/\.\d+Z$/, "")}</meta:creation-date></office:meta></office:document-meta>`;
    const manifest =
      `<?xml version="1.0" encoding="UTF-8"?>` +
      `<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">` +
      `<manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="${ODP_MIME}"/>` +
      `<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>` +
      `<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>` +
      `<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>` +
      this.media.files.map((f) => `<manifest:file-entry manifest:full-path="${f.path}" manifest:media-type="${f.mime}"/>`).join("") +
      `</manifest:manifest>`;

    // "mimetype" must be the first entry and stored uncompressed
    const files: Record<string, [Uint8Array, { level: 0 | 6 }]> = {
      mimetype: [strToU8(ODP_MIME), { level: 0 }],
      "content.xml": [strToU8(content), { level: 6 }],
      "styles.xml": [strToU8(stylesXml), { level: 6 }],
      "meta.xml": [strToU8(meta), { level: 6 }],
      "META-INF/manifest.xml": [strToU8(manifest), { level: 6 }],
    };
    for (const f of this.media.files) files[f.path] = [f.data, { level: 0 }];
    return zipSync(files);
  }

  private title(): string {
    const core = this.pkg.xmlFresh("docProps/core.xml");
    const t = getTextByPathList(core, ["cp:coreProperties", "dc:title"]);
    return typeof t === "string" ? decodeXmlEntities(t) : "";
  }

  // ------------------------------------------------------------- background
  /** Full-slide picture used as background (more portable than a bitmap page fill). */
  private bgPicture = "";

  private pageStyle(warpObj: any, index: number, hidden: boolean): string {
    this.bgPicture = "";
    let props = `draw:background-size="full"`;
    let css = "";
    try {
      css = String(getSlideBackgroundFill(warpObj, index) ?? "");
    } catch {
      css = "";
    }
    const img = /url\((["']?)(data:([^;]+);base64,([^"')]+))\1\)/.exec(css);
    if (img) {
      const ext = img[3].split("/")[1]?.replace("jpeg", "jpg") ?? "png";
      const data = base64ToBytes(img[4]);
      const path = this.media.add("bg:" + img[4].slice(0, 64) + img[4].length, data, ext);
      const gs = this.graphicStyle(null, null);
      this.bgPicture =
        `<draw:frame draw:style-name="${gs}" draw:layer="layout" draw:name="Background" svg:x="0cm" svg:y="0cm" svg:width="${cm(this.pageW)}" svg:height="${cm(this.pageH)}">` +
        `<draw:image xlink:href="${path}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`;
      props += ` draw:fill="none"`;
    } else {
      const c = cssColor(css);
      props += c ? ` draw:fill="solid" draw:fill-color="${c.color}"` : ` draw:fill="solid" draw:fill-color="#ffffff"`;
    }
    if (hidden) props += ` presentation:visibility="hidden"`;
    return this.styles.get("drawing-page", "dp", `<style:drawing-page-properties ${props}/>`);
  }

  private masterShapes(warpObj: any): string {
    let out = "";
    const slideShow = getTextByPathList(warpObj.slideContent, ["p:sld", "attrs", "showMasterSp"]);
    if (slideShow === "0") return "";
    const layoutShow = getTextByPathList(warpObj.slideLayoutContent, ["p:sldLayout", "attrs", "showMasterSp"]);
    if (layoutShow !== "0") {
      const mtree = getTextByPathList(warpObj.slideMasterContent, ["p:sldMaster", "p:cSld", "p:spTree"]);
      out += this.tree(mtree, warpObj, "slideMasterBg", identity, true);
    }
    const ltree = getTextByPathList(warpObj.slideLayoutContent, ["p:sldLayout", "p:cSld", "p:spTree"]);
    out += this.tree(ltree, warpObj, "slideLayoutBg", identity, true);
    return out;
  }

  // ------------------------------------------------------------- shape tree
  /** Converts the children of a shape tree / group in document order. */
  private tree(tree: any, warpObj: any, source: string, map: MapBox, skipPlaceholders = false): string {
    if (!tree) return "";
    const items: Array<{ key: string; node: any; order: number }> = [];
    for (const key of Object.keys(tree)) {
      if (key === "attrs" || key === "p:nvGrpSpPr" || key === "p:grpSpPr") continue;
      for (const node of asArray(tree[key])) items.push({ key, node, order: Number(node?.attrs?.order ?? 0) });
    }
    items.sort((a, b) => a.order - b.order);
    let out = "";
    for (const { key, node } of items) {
      if (skipPlaceholders && getTextByPathList(node, [nvName(key), "p:nvPr", "p:ph"]) !== undefined) continue;
      try {
        out += this.node(key, node, warpObj, source, map);
      } catch (e) {
        warpObj.ctx.warn(`ODP export: shape skipped (${(e as Error).message})`, e);
      }
    }
    return out;
  }

  private node(key: string, node: any, warpObj: any, source: string, map: MapBox): string {
    switch (key) {
      case "p:sp":
        return this.shape(node, warpObj, source, map);
      case "p:pic":
        return this.picture(node, warpObj, source, map);
      case "p:cxnSp":
        return this.connector(node, warpObj, source, map);
      case "p:grpSp":
        return this.group(node, warpObj, source, map);
      case "p:graphicFrame":
        return this.graphicFrame(node, warpObj, source, map);
      case "mc:AlternateContent": {
        const fb = node["mc:Fallback"] ?? asArray(node["mc:Choice"])[0];
        return fb ? this.tree(fb, warpObj, source, map) : "";
      }
      default:
        return "";
    }
  }

  private xfrmBox(xfrm: any): Box | null {
    const off = xfrm?.["a:off"]?.attrs;
    const ext = xfrm?.["a:ext"]?.attrs;
    if (!off || !ext) return null;
    return {
      x: num(off.x) * slideFactor,
      y: num(off.y) * slideFactor,
      w: num(ext.cx) * slideFactor,
      h: num(ext.cy) * slideFactor,
      rot: num(xfrm.attrs?.rot) / 60000 || 0,
      flipH: xfrm.attrs?.flipH === "1",
      flipV: xfrm.attrs?.flipV === "1",
    };
  }

  /** Layout and master counterparts of a placeholder (same lookup as the renderer). */
  private placeholderNodes(node: any, nv: string, warpObj: any): { layout: any; master: any; idx: any; type: any } {
    const idx = getTextByPathList(node, [nv, "p:nvPr", "p:ph", "attrs", "idx"]);
    let type = getTextByPathList(node, [nv, "p:nvPr", "p:ph", "attrs", "type"]);
    let layout: any, master: any;
    if (idx !== undefined) {
      layout = warpObj.slideLayoutTables.idxTable[idx];
      master = type !== undefined ? warpObj.slideMasterTables.typeTable[type] : warpObj.slideMasterTables.idxTable[idx];
    } else if (type !== undefined) {
      layout = warpObj.slideLayoutTables.typeTable[type];
      master = warpObj.slideMasterTables.typeTable[type];
    }
    if (type === undefined) {
      if (getTextByPathList(node, [nv, "p:cNvSpPr", "attrs", "txBox"]) === "1") type = "textBox";
      else type = getTextByPathList(layout, ["p:nvSpPr", "p:nvPr", "p:ph", "attrs", "type"]) ?? "obj";
    }
    return { layout, master, idx, type };
  }

  private resolveBox(node: any, layout: any, master: any, xfrmPath: string[], map: MapBox): Box | null {
    const b =
      this.xfrmBox(getTextByPathList(node, xfrmPath)) ??
      this.xfrmBox(getTextByPathList(layout, xfrmPath)) ??
      this.xfrmBox(getTextByPathList(master, xfrmPath));
    return b ? map(b) : null;
  }

  /** svg:x/y/width/height or draw:transform for rotated shapes. */
  private geomAttrs(b: Box): string {
    if (!b.rot) return `svg:x="${cm(b.x)}" svg:y="${cm(b.y)}" svg:width="${cm(b.w)}" svg:height="${cm(b.h)}"`;
    // OOXML rotates clockwise around the centre; ODF rotates counter-clockwise around the origin then translates
    const a = (-b.rot * Math.PI) / 180;
    const cx = b.x + b.w / 2;
    const cy = b.y + b.h / 2;
    const tx = cx - ((b.w / 2) * Math.cos(a) + (b.h / 2) * Math.sin(a));
    const ty = cy - (-(b.w / 2) * Math.sin(a) + (b.h / 2) * Math.cos(a));
    return `svg:width="${cm(b.w)}" svg:height="${cm(b.h)}" draw:transform="rotate (${a.toFixed(6)}) translate (${cm(tx)} ${cm(ty)})"`;
  }

  private graphicStyle(fill: { color: string; opacity?: number } | null, stroke: any, extra = ""): string {
    let p = "";
    if (fill) {
      p += ` draw:fill="solid" draw:fill-color="${fill.color}"`;
      if (fill.opacity !== undefined && fill.opacity < 1) p += ` draw:opacity="${Math.round(fill.opacity * 100)}%"`;
    } else p += ` draw:fill="none"`;
    const sc = stroke && typeof stroke === "object" ? cssColor(stroke.color) : null;
    if (sc) {
      const w = num(stroke.width);
      p += ` draw:stroke="${stroke.strokeDasharray && stroke.strokeDasharray !== "0" ? "dash" : "solid"}" svg:stroke-color="${sc.color}" svg:stroke-width="${cm(Number.isFinite(w) ? w * (96 / 72) : 1)}"`;
    } else p += ` draw:stroke="none"`;
    p += ` draw:shadow="hidden"` + extra;
    return this.styles.get("graphic", "gr", `<style:graphic-properties${p}/>`);
  }

  private shape(node: any, warpObj: any, source: string, map: MapBox): string {
    const { layout, master, idx, type } = this.placeholderNodes(node, "p:nvSpPr", warpObj);
    const b = this.resolveBox(node, layout, master, ["p:spPr", "a:xfrm"], map);
    if (!b) return "";
    const name = esc(decodeXmlEntities(String(getTextByPathList(node, ["p:nvSpPr", "p:cNvPr", "attrs", "name"]) ?? "")));
    let fill: { color: string; opacity?: number } | null = null;
    try {
      fill = cssColor(getShapeFill(node, undefined, false, warpObj, source));
    } catch {
      fill = null;
    }
    let stroke: any = null;
    try {
      stroke = node["p:spPr"] ? getBorder(node, undefined, true, "shape", warpObj) : null;
    } catch {
      stroke = null;
    }
    const prst = getTextByPathList(node, ["p:spPr", "a:prstGeom", "attrs", "prst"]);
    const custGeom = getTextByPathList(node, ["p:spPr", "a:custGeom"]);
    const txBody = node["p:txBody"];
    const userDrawn = getTextByPathList(node, ["p:nvSpPr", "p:nvPr", "attrs", "userDrawn"]);
    const showText = txBody && (source === "slide" || userDrawn === "1");
    const vClass = String(getVerticalAlign(node, layout, master, type) ?? "");
    const anchor = vClass.includes("v-mid") ? "middle" : vClass.includes("v-down") ? "bottom" : "top";
    const bodyPr = txBody?.["a:bodyPr"]?.attrs ?? {};
    const ins = (v: unknown, d: number) => cm((Number.isFinite(num(v)) ? num(v) : d) * slideFactor);
    const pad = ` fo:padding-left="${ins(bodyPr.lIns, 91440)}" fo:padding-right="${ins(bodyPr.rIns, 91440)}" fo:padding-top="${ins(bodyPr.tIns, 45720)}" fo:padding-bottom="${ins(bodyPr.bIns, 45720)}" draw:textarea-vertical-align="${anchor}" draw:auto-grow-height="false" fo:wrap-option="${bodyPr.wrap === "none" ? "no-wrap" : "wrap"}"`;
    // the renderer formats text of geometric shapes with the master's "other" style
    const textType = (prst !== undefined || custGeom) && type !== "diagram" && type !== "textBox" ? "shape" : type;
    const text = showText ? this.textBody(txBody, node, layout, master, textType, idx, warpObj) : "";
    const isPlainBox = (prst === undefined || prst === "rect") && !custGeom && !fill && !(stroke && typeof stroke === "object" && cssColor(stroke.color));
    const id = ++this.shapeCounter;

    if (isPlainBox) {
      if (!text) return "";
      const gs = this.graphicStyle(null, null, pad);
      return `<draw:frame draw:style-name="${gs}" draw:layer="layout" draw:name="${name || "Text " + id}" ${this.geomAttrs(b)}><draw:text-box>${text}</draw:text-box></draw:frame>`;
    }
    const gs = this.graphicStyle(fill, stroke, pad);
    // Geometry is written as an explicit draw:enhanced-path (named ODF shape
    // types are not understood by every office suite, PowerPoint ignores them).
    // Coordinates are in the shape's own, unscaled pixel space.
    const raw = this.resolveBox(node, layout, master, ["p:spPr", "a:xfrm"], identity) ?? b;
    const w = Math.max(1, raw.w),
      h = Math.max(1, raw.h);
    let parts: GeometryPart[] = [];
    if (custGeom) {
      parts = custGeomToSvgPaths(custGeom, w, h).map((p) => ({ cmds: parseSvgPath(p.d), fill: p.fill, stroke: p.stroke }));
    } else if (prst === undefined || prst === "rect") {
      parts = [{ cmds: parseSvgPath(`M0,0 L${w},0 L${w},${h} L0,${h} Z`), fill: true, stroke: true }];
    } else {
      // reuse the renderer's SVG drawing of the preset shape so geometry matches the viewer
      try {
        parts = geometryFromSvgMarkup(String(processSpNode(node, undefined, warpObj, source) ?? ""));
      } catch {
        parts = [];
      }
      if (!parts.length) parts = [{ cmds: parseSvgPath(`M0,0 L${w},0 L${w},${h} L0,${h} Z`), fill: true, stroke: true }];
    }
    if (!parts.length) return "";
    const odfType = prst && PRESET_TO_ODF[prst] ? PRESET_TO_ODF[prst] : "non-primitive";
    const mirror = (b.flipH ? ` draw:mirror-horizontal="true"` : "") + (b.flipV ? ` draw:mirror-vertical="true"` : "");
    return (
      `<draw:custom-shape draw:style-name="${gs}" draw:layer="layout" draw:name="${name}" ${this.geomAttrs(b)}>${text}` +
      `<draw:enhanced-geometry svg:viewBox="0 0 ${ENHANCED_SIZE} ${ENHANCED_SIZE}" draw:type="${odfType}" draw:enhanced-path="${toEnhancedPath(parts, w, h)}"${mirror}/></draw:custom-shape>`
    );
  }

  /** Converts a txBody to text:p elements with resolved character styles. */
  private textBody(txBody: any, spNode: any, layout: any, master: any, type: any, idx: any, warpObj: any): string {
    const pFontStyle = getTextByPathList(spNode, ["p:style", "a:fontRef"]);
    const counters: Array<{ lvl: number; type: string; n: number }> = [];
    let out = "";
    asArray(txBody["a:p"]).forEach((pNode: any, i: number) => {
      const lvl = Number(getTextByPathList(pNode, ["a:pPr", "attrs", "lvl"]) ?? 0) || 0;
      // direction and alignment (physical left/right so RTL paragraphs stay correct)
      let rtl = false;
      try {
        rtl = getPregraphDir(pNode, txBody, idx, type, warpObj) === "pregraph-rtl";
      } catch {
        rtl = false;
      }
      let align = rtl ? "right" : "left";
      try {
        const h = String(getHorizontalAlign(pNode, txBody, idx, type, rtl ? "pregraph-rtl" : "pregraph-ltr", warpObj) ?? "");
        align = h.includes("h-mid") ? "center" : h.includes("h-right") ? "right" : h.includes("h-left") ? "left" : h.includes("h-just") || h.includes("h-dist") ? "justify" : align;
      } catch {
        /* keep default */
      }
      // bullet
      let bullet = "";
      let isBullet = false;
      try {
        const bu = genBuChar(pNode, i, spNode, txBody, pFontStyle, idx, type, warpObj);
        const html = String(bu?.[0] ?? "");
        if (html) {
          isBullet = true;
          const nm = /data-bulltname = '([^']*)'(?:.*?data-bulltstart = '(\d+)')?/.exec(html);
          if (nm) {
            while (counters.length && counters[counters.length - 1].lvl > lvl) counters.pop();
            const top = counters[counters.length - 1];
            if (top && top.lvl === lvl && top.type === nm[1]) top.n++;
            else {
              if (top && top.lvl === lvl) counters.pop();
              counters.push({ lvl, type: nm[1], n: Number(nm[2]) || 1 });
            }
            bullet = getNumTypeNum(nm[1], counters[counters.length - 1].n);
          } else {
            const ch = />([^<>]+)<\/div><\/div>\s*$/.exec(html);
            bullet = ch ? decodeXmlEntities(ch[1]) + " " : "";
          }
        }
      } catch {
        bullet = "";
      }
      let indent = 0;
      try {
        indent = num(getPregraphMargn(pNode, idx, type, isBullet, warpObj)?.[1]) || 0;
      } catch {
        indent = 0;
      }
      const ps = this.styles.get(
        "paragraph",
        "P",
        `<style:paragraph-properties fo:text-align="${align}"${rtl ? ` style:writing-mode="rl-tb"` : ""}${indent > 0 ? ` fo:margin-${rtl ? "right" : "left"}="${cm(indent)}"` : ""}/>`,
      );
      // runs in document order
      const runs: Array<{ key: string; node: any; order: number }> = [];
      for (const k of ["a:r", "a:br", "a:fld"]) for (const r of asArray(pNode[k])) runs.push({ key: k, node: r, order: Number(r?.attrs?.order ?? 0) });
      runs.sort((a, b) => a.order - b.order);
      let spans = "";
      if (bullet) spans += `<text:span text:style-name="${this.runStyle(runs.find((r) => r.key !== "a:br")?.node ?? pNode, pNode, txBody, pFontStyle, layout, idx, type, 1, warpObj, isBullet)}">${esc(bullet)}</text:span>`;
      runs.forEach(({ key, node }, j) => {
        if (key === "a:br") {
          spans += "<text:line-break/>";
          return;
        }
        const t = typeof node["a:t"] === "string" ? decodeXmlEntities(node["a:t"]) : "";
        if (!t) return;
        const ts = this.runStyle(node, pNode, txBody, pFontStyle, layout, idx, type, runs.length, warpObj, isBullet, j);
        spans += `<text:span text:style-name="${ts}">${odfText(t)}</text:span>`;
      });
      out += `<text:p text:style-name="${ps}">${spans}</text:p>`;
    });
    return out;
  }

  /** Text style from the renderer's resolved span CSS. */
  private runStyle(rNode: any, pNode: any, txBody: any, pFontStyle: any, layout: any, idx: any, type: any, n: number, warpObj: any, isBullet: boolean, j = 0): string {
    let decl: Record<string, string> = {};
    try {
      const html = String(genSpanElement(rNode, j, pNode, txBody, pFontStyle, layout, idx, type, n, warpObj, isBullet) ?? "");
      const cls = /class='text-block ([^']*)'/.exec(html)?.[1]?.trim();
      const entry = cls ? Object.values(warpObj.ctx.styleTable).find((s: any) => s.name === cls) : undefined;
      decl = { ...cssDecls((entry as any)?.text ?? ""), ...cssDecls(/style='([^']*)'/.exec(html)?.[1] ?? "") };
    } catch {
      decl = {};
    }
    let p = "";
    const size = num(decl["font-size"]);
    if (Number.isFinite(size)) p += ` fo:font-size="${Math.round(size * 0.75 * 10) / 10}pt"`;
    const color = cssColor(decl["color"]);
    if (color) p += ` fo:color="${color.color}"`;
    const fam = decl["font-family"]?.split(",")[0]?.replace(/["']/g, "").trim();
    if (fam && fam !== "inherit") p += ` fo:font-family="${esc(fam)}" style:font-family-asian="${esc(fam)}" style:font-family-complex="${esc(fam)}"`;
    if (decl["font-weight"] === "bold" || num(decl["font-weight"]) >= 600) p += ` fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"`;
    if (decl["font-style"] === "italic") p += ` fo:font-style="italic" style:font-style-asian="italic" style:font-style-complex="italic"`;
    const deco = decl["text-decoration"] ?? decl["text-decoration-line"] ?? "";
    if (deco.includes("underline")) p += ` style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"`;
    if (deco.includes("line-through")) p += ` style:text-line-through-style="solid"`;
    if (decl["text-transform"] === "uppercase") p += ` fo:text-transform="uppercase"`;
    return this.styles.get("text", "T", `<style:text-properties${p}/>`);
  }

  private resObj(warpObj: any, source: string): Record<string, { type: string; target: string }> {
    return source === "slideMasterBg" ? warpObj.masterResObj : source === "slideLayoutBg" ? warpObj.layoutResObj : warpObj.slideResObj;
  }

  private picture(node: any, warpObj: any, source: string, map: MapBox): string {
    const rid = getTextByPathList(node, ["p:blipFill", "a:blip", "attrs", "r:embed"]);
    const target = rid ? this.resObj(warpObj, source)[rid]?.target : undefined;
    const data = target ? this.pkg.bytes(target) : null;
    const { layout, master } = this.placeholderNodes(node, "p:nvPicPr", warpObj);
    const b = this.resolveBox(node, layout, master, ["p:spPr", "a:xfrm"], map);
    if (!data || !b || !target) return "";
    const path = this.media.add(target, data, extractFileExtension(target) || "png");
    const gs = this.graphicStyle(null, null);
    const name = esc(decodeXmlEntities(String(getTextByPathList(node, ["p:nvPicPr", "p:cNvPr", "attrs", "name"]) ?? "")));
    const descr = getTextByPathList(node, ["p:nvPicPr", "p:cNvPr", "attrs", "descr"]);
    return (
      `<draw:frame draw:style-name="${gs}" draw:layer="layout" draw:name="${name}" ${this.geomAttrs(b)}>` +
      `<draw:image xlink:href="${path}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>` +
      (descr ? `<svg:desc>${esc(decodeXmlEntities(String(descr)))}</svg:desc>` : "") +
      `</draw:frame>`
    );
  }

  private connector(node: any, warpObj: any, _source: string, map: MapBox): string {
    const b = this.resolveBox(node, undefined, undefined, ["p:spPr", "a:xfrm"], map);
    if (!b) return "";
    let stroke: any = null;
    try {
      stroke = getBorder(node, undefined, true, "shape", warpObj);
    } catch {
      stroke = null;
    }
    const gs = this.graphicStyle(null, stroke && typeof stroke === "object" ? stroke : { color: "#000000", width: 1 });
    const x1 = b.flipH ? b.x + b.w : b.x;
    const x2 = b.flipH ? b.x : b.x + b.w;
    const y1 = b.flipV ? b.y + b.h : b.y;
    const y2 = b.flipV ? b.y : b.y + b.h;
    return `<draw:line draw:style-name="${gs}" draw:layer="layout" svg:x1="${cm(x1)}" svg:y1="${cm(y1)}" svg:x2="${cm(x2)}" svg:y2="${cm(y2)}"/>`;
  }

  private group(node: any, warpObj: any, source: string, map: MapBox): string {
    const xf = getTextByPathList(node, ["p:grpSpPr", "a:xfrm"]);
    const off = xf?.["a:off"]?.attrs;
    const ext = xf?.["a:ext"]?.attrs;
    const chOff = xf?.["a:chOff"]?.attrs;
    const chExt = xf?.["a:chExt"]?.attrs;
    let childMap = map;
    if (off && ext && chOff && chExt) {
      const sx = num(chExt.cx) ? num(ext.cx) / num(chExt.cx) : 1;
      const sy = num(chExt.cy) ? num(ext.cy) / num(chExt.cy) : 1;
      const ox = num(off.x) * slideFactor,
        oy = num(off.y) * slideFactor;
      const cx = num(chOff.x) * slideFactor,
        cy = num(chOff.y) * slideFactor;
      childMap = (b) => map({ ...b, x: ox + (b.x - cx) * sx, y: oy + (b.y - cy) * sy, w: b.w * sx, h: b.h * sy });
    }
    const inner = this.tree(node, warpObj, source, childMap);
    return inner ? `<draw:g>${inner}</draw:g>` : "";
  }

  private graphicFrame(node: any, warpObj: any, source: string, map: MapBox): string {
    const b = this.resolveBox(node, undefined, undefined, ["p:xfrm"], map);
    if (!b) return "";
    const data = getTextByPathList(node, ["a:graphic", "a:graphicData"]);
    const uri = String(data?.attrs?.uri ?? "");
    if (uri.endsWith("/table")) return this.table(data["a:tbl"], b, warpObj);
    if (uri.endsWith("/chart")) {
      const rid = getTextByPathList(data, ["c:chart", "attrs", "r:id"]);
      const target = rid ? this.resObj(warpObj, source)[rid]?.target : undefined;
      const part = target ? this.pkg.xmlFresh(target) : null;
      const model = part
        ? parseChart(part, (spPr) => {
            const f = spPr?.["a:solidFill"];
            try {
              return f ? cssColor("#" + getSolidFill(f, undefined, undefined, warpObj))?.color : undefined;
            } catch {
              return undefined;
            }
          })
        : null;
      if (!model) return "";
      const svg = renderChartSvg(model, { width: b.w, height: b.h, palette: themePalette(warpObj) });
      const path = this.media.add("chart:" + target, strToU8(svg), "svg");
      return `<draw:frame draw:style-name="${this.graphicStyle(null, null)}" ${this.geomAttrs(b)}><draw:image xlink:href="${path}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame>`;
    }
    if (uri.endsWith("/diagram") && warpObj.digramFileContent) {
      const tree = getTextByPathList(warpObj.digramFileContent, ["p:drawing", "p:spTree"]);
      const xf = { x: b.x, y: b.y };
      return `<draw:g>${this.tree(tree, warpObj, "diagramBg", (c) => map({ ...c, x: c.x + xf.x, y: c.y + xf.y }))}</draw:g>`;
    }
    return "";
  }

  private table(tbl: any, b: Box, warpObj: any): string {
    if (!tbl) return "";
    const cols = asArray(getTextByPathList(tbl, ["a:tblGrid", "a:gridCol"]));
    let xml = `<table:table>`;
    for (const c of cols) {
      const w = num(c?.attrs?.w) * slideFactor;
      const cs = this.styles.get("table-column", "co", `<style:table-column-properties style:column-width="${cm(w)}"/>`);
      xml += `<table:table-column table:style-name="${cs}"/>`;
    }
    for (const tr of asArray(tbl["a:tr"])) {
      const h = num(tr?.attrs?.h) * slideFactor;
      const rs = this.styles.get("table-row", "ro", `<style:table-row-properties style:min-row-height="${cm(h)}"/>`);
      xml += `<table:table-row table:style-name="${rs}">`;
      for (const tc of asArray(tr["a:tc"])) {
        if (tc?.attrs?.hMerge === "1" || tc?.attrs?.vMerge === "1") {
          xml += `<table:covered-table-cell/>`;
          continue;
        }
        let fill: { color: string } | null = null;
        const sf = getTextByPathList(tc, ["a:tcPr", "a:solidFill"]);
        if (sf) {
          try {
            fill = cssColor("#" + getSolidFill(sf, undefined, undefined, warpObj));
          } catch {
            fill = null;
          }
        }
        const cellStyle = this.styles.get(
          "table-cell",
          "ce",
          `<style:graphic-properties draw:fill="${fill ? "solid" : "none"}"${fill ? ` draw:fill-color="${fill.color}"` : ""}/>` +
            `<style:table-cell-properties fo:border="0.5pt solid #808080" fo:padding="0.1cm"/>`,
        );
        const span = (tc?.attrs?.gridSpan ? ` table:number-columns-spanned="${tc.attrs.gridSpan}"` : "") + (tc?.attrs?.rowSpan ? ` table:number-rows-spanned="${tc.attrs.rowSpan}"` : "");
        const paras = asArray(getTextByPathList(tc, ["a:txBody", "a:p"]))
          .map((p: any) => {
            const t = asArray(p["a:r"]).map((r: any) => (typeof r?.["a:t"] === "string" ? decodeXmlEntities(r["a:t"]) : "")).join("");
            return `<text:p>${odfText(t)}</text:p>`;
          })
          .join("");
        xml += `<table:table-cell table:style-name="${cellStyle}" office:value-type="string"${span}>${paras || "<text:p/>"}</table:table-cell>`;
      }
      xml += `</table:table-row>`;
    }
    xml += `</table:table>`;
    return `<draw:frame draw:style-name="${this.graphicStyle(null, null)}" ${this.geomAttrs(b)}>${xml}</draw:frame>`;
  }
}

function nvName(key: string): string {
  return key === "p:pic" ? "p:nvPicPr" : key === "p:grpSp" ? "p:nvGrpSpPr" : key === "p:graphicFrame" ? "p:nvGraphicFramePr" : key === "p:cxnSp" ? "p:nvCxnSpPr" : "p:nvSpPr";
}

/** Escapes text and keeps runs of spaces / tabs (text:s, text:tab). */
function odfText(t: string): string {
  return esc(t)
    .replace(/\t/g, "<text:tab/>")
    .replace(/ {2,}/g, (m) => ` <text:s text:c="${m.length - 1}"/>`)
    .replace(/\u000b/g, "<text:line-break/>");
}

function base64ToBytes(b64: string): Uint8Array {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

const NS =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ' +
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" ' +
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ' +
  'xmlns:xlink="http://www.w3.org/1999/xlink" ' +
  'xmlns:dc="http://purl.org/dc/elements/1.1/" ' +
  'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" ' +
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" ' +
  'xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"';

/** Converts a presentation to .odp bytes. */
export function exportOdpBytes(input: PresentationDocument | PptxPackage, options: OdpExportOptions = {}): Uint8Array {
  const pkg = input instanceof PresentationDocument ? input.pkg : input;
  try {
    return new OdpWriter(pkg, options).build();
  } catch (e) {
    if (e instanceof PptxError) throw e;
    throw new PptxError("EXPORT_FAILED", "ODP export failed: " + (e as Error).message, e);
  }
}

/** Converts a presentation to an .odp `Blob`. */
export function exportOdp(input: PresentationDocument | PptxPackage, options: OdpExportOptions = {}): Blob {
  return new Blob([exportOdpBytes(input, options) as BlobPart], { type: ODP_MIME });
}
