import type { PptxPackage } from "../core/package";
import type { XmlNode } from "../core/xml";
import { PptxError } from "../errors";

/** Options that influence how slides are rendered to HTML. */
export interface RenderOptions {
  /**
   * How much of the theme to apply.
   * - `true` (default): backgrounds and shapes from layouts/masters
   * - `"colorsAndImageOnly"`: only slide background colour/image
   * - `false`: no theme processing
   */
  themeProcess?: boolean | "colorsAndImageOnly";
  /** Render video and audio elements (default `true`). */
  mediaProcess?: boolean;
  /**
   * Embed externally linked videos (e.g. YouTube) in a sandboxed iframe.
   * Default `false`: a poster image with a link is shown instead, so opening
   * an untrusted file never loads third-party content automatically.
   */
  externalMedia?: boolean;
  /**
   * Direction of paragraphs that do not declare one (`a:pPr/@rtl`):
   * - `"file"` (default): as PowerPoint, left to right unless the file says otherwise
   * - `"auto"`: from the first strong character (Hebrew/Arabic text becomes RTL), like HTML `dir="auto"`
   * - `"rtl"`: right to left
   * Paragraphs with an explicit direction are never changed.
   */
  textDirection?: "file" | "auto" | "rtl";
  /** Extra pixels added to the slide size (v1 `incSlide`). */
  incSlide?: { width?: number; height?: number };
  /**
   * Measures the rendered width (px) of an HTML fragment. Used to lay out
   * bullet text. The browser renderer injects a DOM-based implementation;
   * in Node/workers widths are treated as 0.
   */
  measureHtmlWidth?: (html: string) => number;
  /** Receives non-fatal problems found while rendering. */
  onWarning?: (warning: PptxError) => void;
  /** Prefix for generated element ids/classes, to allow several viewers per page. */
  idPrefix?: string;
}

export interface StyleEntry {
  name: string;
  text: string;
  suffix?: string;
}

/** Presentation-wide state shared by all slides during one render. */
export class RenderContext {
  readonly pkg: PptxPackage;
  readonly options: RenderOptions;
  /** Generated CSS classes (v1 `styleTable`). */
  styleTable: Record<string, StyleEntry> = {};
  appVersion = 0;
  defaultTextStyle: XmlNode | undefined;
  tableStyles: XmlNode | null = null;
  slideWidth = 0;
  slideHeight = 0;
  /** Counter for unique chart ids. */
  chartCounter = 0;
  /** 1-based number of the slide currently being rendered. */
  currentSlide = 0;
  readonly idPrefix: string;

  constructor(pkg: PptxPackage, options: RenderOptions = {}) {
    this.pkg = pkg;
    this.options = options;
    this.idPrefix = options.idPrefix ?? "pptx";
  }

  get processFullTheme(): boolean | "colorsAndImageOnly" {
    return this.options.themeProcess ?? true;
  }

  get mediaProcess(): boolean {
    return this.options.mediaProcess ?? true;
  }

  get textDirection(): "file" | "auto" | "rtl" {
    return this.options.textDirection ?? "file";
  }

  get externalMedia(): boolean {
    return this.options.externalMedia ?? false;
  }

  /** True when a text measurer is available (bullet paragraphs are shrunk to fit their text). */
  get canMeasure(): boolean {
    return typeof this.options.measureHtmlWidth === "function";
  }

  measureHtmlWidth(html: string): number {
    return this.options.measureHtmlWidth ? this.options.measureHtmlWidth(html) : 0;
  }

  /** Object URLs created for media; revoke them with {@link revokeObjectUrls}. */
  readonly objectUrls: string[] = [];

  /** Creates (and tracks) an object URL for embedded media. Returns "" when unsupported. */
  createObjectUrl(blob: Blob): string {
    if (typeof URL === "undefined" || typeof URL.createObjectURL !== "function") return "";
    try {
      const url = URL.createObjectURL(blob);
      this.objectUrls.push(url);
      return url;
    } catch {
      return "";
    }
  }

  revokeObjectUrls(): void {
    for (const url of this.objectUrls.splice(0)) {
      try {
        URL.revokeObjectURL(url);
      } catch {
        /* ignore */
      }
    }
  }

  warn(message: string, cause?: unknown): void {
    const w = new PptxError("RENDER_WARNING", message, cause);
    w.slide = this.currentSlide || undefined;
    this.options.onWarning?.(w);
  }
}

export interface ResourceRef {
  type: string;
  target: string;
}

/**
 * Per-slide bundle of parsed parts (v1 name `warpObj` kept so the ported code
 * stays recognisable).
 */
export interface WarpObj {
  ctx: RenderContext;
  zip: PptxPackage;
  slideFile: string;
  slideLayoutContent: XmlNode | null;
  slideLayoutTables: NodeTables;
  slideMasterContent: XmlNode | null;
  slideMasterTables: NodeTables;
  slideContent: XmlNode | null;
  slideResObj: Record<string, ResourceRef>;
  slideMasterTextStyles: XmlNode | undefined;
  layoutResObj: Record<string, ResourceRef>;
  masterResObj: Record<string, ResourceRef>;
  themeContent: XmlNode | null;
  themeResObj: Record<string, ResourceRef>;
  digramFileContent: XmlNode | null;
  diagramResObj: Record<string, ResourceRef>;
  defaultTextStyle: XmlNode | undefined;
  /** `<p:clrMapOvr><a:overrideClrMapping>` attributes of the layout, if any. */
  layoutClrOverride?: Record<string, string>;
  "loaded-images"?: Record<string, string>;
  [key: string]: any;
}

export interface NodeTables {
  idTable: Record<string, any>;
  idxTable: Record<string, any>;
  typeTable: Record<string, any>;
}
