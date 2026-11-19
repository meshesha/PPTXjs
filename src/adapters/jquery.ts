/**
 * jQuery plugin with a v1-compatible API.
 *
 * ```js
 * import $ from "jquery";
 * import { installJQueryPlugin } from "@meshesha/pptxjs/jquery";
 * installJQueryPlugin($);
 * $("#result").pptxToHtml({ pptxFileUrl: "deck.pptx", slideMode: true });
 * ```
 * With the UMD build (`<script src="pptxjs.umd.js">`) the plugin is installed
 * automatically when `window.jQuery` exists.
 */
import { PptxViewer, type ViewerOptions, type SlideTransition } from "../viewer/viewer";
import { PptxEditor, type EditorOptions } from "../editor/editor";

/** Options of PPTXjs v1 (`$.fn.pptxToHtml`). */
export interface LegacyOptions {
  pptxFileUrl?: string;
  fileInputId?: string;
  slidesScale?: string | number;
  slideMode?: boolean;
  slideType?: string;
  keyBoardShortCut?: boolean;
  mediaProcess?: boolean;
  themeProcess?: boolean | "colorsAndImageOnly";
  incSlide?: { width?: number; height?: number };
  slideModeConfig?: {
    first?: number;
    nav?: boolean;
    navTxtColor?: string;
    keyBoardShortCut?: boolean;
    showSlideNum?: boolean;
    showTotalSlideNum?: boolean;
    autoSlide?: number | boolean;
    randomAutoSlide?: boolean;
    loop?: boolean;
    background?: string | false;
    transition?: string;
    transitionTime?: number;
  };
  /** v2 options can be mixed in */
  [key: string]: unknown;
}

/** Converts v1 options to v2 {@link ViewerOptions}. */
export function fromLegacyOptions(o: LegacyOptions): ViewerOptions {
  const sm = o.slideModeConfig ?? {};
  const scale = o.slidesScale !== undefined && o.slidesScale !== "" ? parseFloat(String(o.slidesScale)) / 100 : undefined;
  const transition: SlideTransition =
    sm.transition === "slid" ? "slide" : sm.transition === "fade" ? "fade" : sm.transition === "random" ? "random" : sm.transition === "default" ? "none" : "fade";
  const v2 = { ...o } as Record<string, unknown>;
  for (const k of ["pptxFileUrl", "fileInputId", "slidesScale", "slideMode", "slideType", "keyBoardShortCut", "slideModeConfig", "revealjsPath", "revealjsConfig", "jsZipV2"]) delete v2[k];
  return {
    ...(v2 as ViewerOptions),
    mode: o.slideMode ? "slideshow" : "list",
    scale: scale && Number.isFinite(scale) ? scale : "fit",
    keyboard: o.keyBoardShortCut || sm.keyBoardShortCut ? "global" : true,
    mediaProcess: o.mediaProcess,
    themeProcess: o.themeProcess,
    incSlide: o.incSlide,
    slideshow: {
      first: sm.first,
      nav: sm.nav,
      showSlideNumber: sm.showSlideNum !== false,
      autoplay: sm.autoSlide === true ? 5 : typeof sm.autoSlide === "number" ? sm.autoSlide : false,
      random: sm.randomAutoSlide,
      loop: sm.loop,
      background: sm.background || undefined,
      transition,
      transitionDuration: sm.transitionTime,
    },
  };
}

const DATA_KEY = "pptxjs-viewer";

type JQueryStaticLike = any;

/** Registers `$.fn.pptxToHtml` (v1 compatible), `$.fn.pptxViewer` and `$.fn.pptxEditor`. */
export function installJQueryPlugin($: JQueryStaticLike): void {
  if (!$ || !$.fn) throw new Error("installJQueryPlugin: jQuery instance required");

  $.fn.pptxToHtml = function (this: JQueryStaticLike, options: LegacyOptions = {}) {
    return this.each(function (this: HTMLElement) {
      const existing: PptxViewer | undefined = $(this).data(DATA_KEY);
      existing?.destroy();
      const viewer = new PptxViewer(this, fromLegacyOptions(options));
      $(this).data(DATA_KEY, viewer);
      if (options.pptxFileUrl) viewer.load(options.pptxFileUrl).catch(() => undefined);
      if (options.fileInputId) {
        const input = document.getElementById(options.fileInputId) as HTMLInputElement | null;
        input?.addEventListener("change", () => {
          const file = input.files?.[0];
          if (file) viewer.load(file).catch(() => undefined);
        });
      }
    });
  };

  /** `$(el).pptxViewer(options)` creates a viewer; `$(el).pptxViewer()` returns it. */
  $.fn.pptxViewer = function (this: JQueryStaticLike, options?: ViewerOptions & { src?: string | Blob }) {
    if (options === undefined) return this.first().data(DATA_KEY);
    return this.each(function (this: HTMLElement) {
      ($(this).data(DATA_KEY) as PptxViewer | undefined)?.destroy();
      const viewer = new PptxViewer(this, options);
      $(this).data(DATA_KEY, viewer);
      if (options.src) viewer.load(options.src).catch(() => undefined);
    });
  };

  $.fn.pptxEditor = function (this: JQueryStaticLike, options?: EditorOptions & { src?: string | Blob }) {
    if (options === undefined) return this.first().data("pptxjs-editor");
    return this.each(function (this: HTMLElement) {
      ($(this).data("pptxjs-editor") as PptxEditor | undefined)?.destroy();
      const editor = new PptxEditor(this, options);
      $(this).data("pptxjs-editor", editor);
      if (options.src) editor.load(options.src).catch(() => undefined);
    });
  };
}

// auto-install for <script> usage
declare const window: { jQuery?: JQueryStaticLike } | undefined;
if (typeof window !== "undefined" && window && window.jQuery && window.jQuery.fn && !window.jQuery.fn.pptxViewer) {
  installJQueryPlugin(window.jQuery);
}

declare global {
  interface JQuery {
    /** PPTXjs v1-compatible plugin. */
    pptxToHtml(options?: LegacyOptions): JQuery;
    /** Returns the viewer of the first element. */
    pptxViewer(): PptxViewer | undefined;
    /** Creates a viewer on every matched element. */
    pptxViewer(options: ViewerOptions & { src?: string | Blob }): JQuery;
    pptxEditor(): PptxEditor | undefined;
    pptxEditor(options: EditorOptions & { src?: string | Blob }): JQuery;
  }
}
