/**
 * PPTXjs v2 - render, edit and export PowerPoint presentations in the browser.
 *
 * @example Browser
 * ```ts
 * import { PptxViewer } from "@meshesha/pptxjs";
 * const viewer = new PptxViewer("#slides", { mode: "slideshow" });
 * await viewer.load("deck.pptx");
 * ```
 *
 * @example Node / SSR / Web Worker (no DOM)
 * ```ts
 * import { renderToHtml } from "@meshesha/pptxjs";
 * const { slides, css } = await renderToHtml(bytes);
 * ```
 */
import { PptxPackage, type PackageLimits } from "./core/package";
import { renderPresentation, type RenderedPresentation, type RenderPresentationOptions } from "./render/presentation";
import { readSource, type PptxSource } from "./source";
import { BASE_CSS } from "./viewer/styles";

export { PptxError, type PptxErrorCode } from "./errors";
export { PptxPackage, type PackageLimits } from "./core/package";
export { readSource, type PptxSource } from "./source";
export {
  renderPresentation,
  listSlides,
  readNotes,
  type RenderedPresentation,
  type RenderedSlide,
  type RenderPresentationOptions,
  type SlideInfo,
} from "./render/presentation";
export type { RenderOptions } from "./render/context";
export { PptxViewer, type ViewerOptions, type ViewerEvents, type ViewerMode, type SlideshowOptions, type SlideTransition } from "./viewer/viewer";
export { BASE_CSS } from "./viewer/styles";
export { parseChart, type ChartModel, type ChartSeries, type ChartGroup } from "./charts/model";
export { renderChartSvg, DEFAULT_PALETTE } from "./charts/svg";
export { PptxEditor, type EditorOptions, type EditorEvents, type ShapeRef } from "./editor/editor";
export { PresentationDocument, type ShapeInfo, type TextParagraph, type TextRun } from "./model/document";
export { exportPptx, exportPptxBytes, PPTX_MIME, type PptxExportOptions } from "./export/pptx";
export { exportOdp, exportOdpBytes, ODP_MIME, type OdpExportOptions } from "./export/odp";

export const VERSION = "2.0.0-alpha.0";

export interface LoadOptions extends RenderPresentationOptions {
  limits?: PackageLimits;
  fetchOptions?: RequestInit;
  signal?: AbortSignal;
}

/** Opens a .pptx package from a URL, Blob/File or bytes (no rendering). */
export async function loadPptx(source: PptxSource, options: Pick<LoadOptions, "limits" | "fetchOptions" | "signal"> = {}): Promise<PptxPackage> {
  const bytes = await readSource(source, options);
  return PptxPackage.fromBytes(bytes, options.limits);
}

/**
 * Renders a presentation to HTML strings without touching the DOM.
 * Works in browsers, Web Workers and Node.js.
 */
export async function renderToHtml(source: PptxSource, options: LoadOptions = {}): Promise<RenderedPresentation> {
  const pkg = await loadPptx(source, options);
  return renderPresentation(pkg, options);
}

/**
 * Renders a presentation into one self-contained HTML string (all slides, the
 * base stylesheet and the generated CSS), e.g. for server-side rendering.
 */
export async function renderToHtmlString(source: PptxSource, options: LoadOptions = {}): Promise<string> {
  const res = await renderToHtml(source, options);
  return (
    `<div class="pptxjs"><style>${BASE_CSS}\n${res.css}</style>` +
    res.slides.map((s) => s.html).join("\n") +
    `</div>`
  );
}
export { defineCustomElements, PptxViewerElement, PptxEditorElement } from "./adapters/element";
