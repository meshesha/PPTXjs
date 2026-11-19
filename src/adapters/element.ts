/**
 * Standard Web Components: `<pptx-viewer>` and `<pptx-editor>`.
 * They work in plain HTML and in every framework that supports custom
 * elements (Angular, Svelte, Lit, Solid, ...).
 *
 * ```html
 * <pptx-viewer src="deck.pptx" mode="slideshow"></pptx-viewer>
 * ```
 * Attributes: `src`, `mode` ("list" | "slideshow"), `scale` ("fit" | number),
 * `transition`, `autoplay` (seconds), `loop`. For other options set the
 * `options` property before `src`. Events: `pptx-load`, `pptx-error`,
 * `pptx-slidechange` (CustomEvent with `detail`).
 */
import { PptxViewer, type ViewerOptions, type SlideTransition } from "../viewer/viewer";
import { PptxEditor, type EditorOptions } from "../editor/editor";
import type { PptxSource } from "../source";

const BaseElement: typeof HTMLElement = typeof HTMLElement !== "undefined" ? HTMLElement : (class {} as unknown as typeof HTMLElement);

export class PptxViewerElement extends BaseElement {
  static observedAttributes = ["src", "mode", "scale", "transition", "autoplay", "loop"];
  /** Extra options (set before `src`). */
  options: ViewerOptions = {};
  private _viewer: PptxViewer | null = null;
  private _source: PptxSource | null = null;

  /** Underlying viewer (after the element is connected). */
  get viewer(): PptxViewer | null {
    return this._viewer;
  }

  /** Accepts a URL string or bytes/Blob (property only). */
  get source(): PptxSource | null {
    return this._source;
  }
  set source(v: PptxSource | null) {
    this._source = v;
    if (v && this._viewer) this._viewer.load(v).catch(() => undefined);
  }

  private buildOptions(): ViewerOptions {
    const scaleAttr = this.getAttribute("scale");
    const autoplay = Number(this.getAttribute("autoplay"));
    return {
      ...this.options,
      mode: (this.getAttribute("mode") as ViewerOptions["mode"]) ?? this.options.mode,
      scale: scaleAttr ? (scaleAttr === "fit" ? "fit" : Number(scaleAttr)) : this.options.scale,
      slideshow: {
        ...this.options.slideshow,
        transition: (this.getAttribute("transition") as SlideTransition) ?? this.options.slideshow?.transition,
        autoplay: autoplay > 0 ? autoplay : this.options.slideshow?.autoplay,
        loop: this.hasAttribute("loop") || this.options.slideshow?.loop,
      },
    };
  }

  connectedCallback(): void {
    if (this._viewer) return;
    this.style.display ||= "block";
    const v = new PptxViewer(this, this.buildOptions());
    v.on("load", ({ presentation }) => this.dispatchEvent(new CustomEvent("pptx-load", { detail: presentation })));
    v.on("error", ({ error }) => this.dispatchEvent(new CustomEvent("pptx-error", { detail: error })));
    v.on("slidechange", ({ index }) => this.dispatchEvent(new CustomEvent("pptx-slidechange", { detail: { index } })));
    this._viewer = v;
    const src = this._source ?? this.getAttribute("src");
    if (src) v.load(src).catch(() => undefined);
  }

  disconnectedCallback(): void {
    this._viewer?.destroy();
    this._viewer = null;
  }

  attributeChangedCallback(name: string, old: string | null, value: string | null): void {
    if (!this._viewer || old === value) return;
    if (name === "src") {
      if (value) this._viewer.load(value).catch(() => undefined);
    } else {
      this._viewer.setOptions(this.buildOptions());
    }
  }
}

export class PptxEditorElement extends BaseElement {
  static observedAttributes = ["src"];
  options: EditorOptions = {};
  private _editor: PptxEditor | null = null;

  get editor(): PptxEditor | null {
    return this._editor;
  }

  connectedCallback(): void {
    if (this._editor) return;
    this.style.display ||= "block";
    const e = new PptxEditor(this, this.options);
    e.on("load", ({ slides }) => this.dispatchEvent(new CustomEvent("pptx-load", { detail: { slides } })));
    e.on("error", ({ error }) => this.dispatchEvent(new CustomEvent("pptx-error", { detail: error })));
    e.on("change", ({ label }) => this.dispatchEvent(new CustomEvent("pptx-change", { detail: { label } })));
    this._editor = e;
    const src = this.getAttribute("src");
    if (src) e.load(src).catch(() => undefined);
  }

  disconnectedCallback(): void {
    this._editor?.destroy();
    this._editor = null;
  }

  attributeChangedCallback(name: string, old: string | null, value: string | null): void {
    if (name === "src" && this._editor && value && old !== value) this._editor.load(value).catch(() => undefined);
  }
}

/** Registers `<pptx-viewer>` and `<pptx-editor>` (idempotent). */
export function defineCustomElements(prefix = "pptx"): void {
  if (typeof customElements === "undefined") return;
  if (!customElements.get(`${prefix}-viewer`)) customElements.define(`${prefix}-viewer`, class extends PptxViewerElement {});
  if (!customElements.get(`${prefix}-editor`)) customElements.define(`${prefix}-editor`, class extends PptxEditorElement {});
}
