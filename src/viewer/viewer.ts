/**
 * PptxViewer: framework-agnostic browser component that renders a .pptx into
 * a container element. Supports a scrolling list of slides or a slideshow with
 * navigation, keyboard control, autoplay, transitions and fullscreen.
 */
import { PptxPackage, type PackageLimits } from "../core/package";
import { PptxError } from "../errors";
import { RenderContext } from "../render/context";
import {
  genGlobalCSS,
  listSlides,
  readPresentationInfo,
  renderSlide,
  type RenderedPresentation,
  type RenderedSlide,
  type RenderPresentationOptions,
  type SlideInfo,
} from "../render/presentation";
import { base64ArrayBuffer } from "../utils/base64";
import { Emitter, type Unsubscribe } from "../utils/emitter";
import { readSource, type PptxSource } from "../source";
import { BASE_CSS } from "./styles";
import { VIEWER_CSS } from "./viewer-styles";

export type ViewerMode = "list" | "slideshow";
export type SlideTransition = "none" | "fade" | "slide" | "zoom" | "random";

export interface SlideshowOptions {
  /** 1-based slide to start with (default 1). */
  first?: number;
  /** Show the navigation toolbar (default true). */
  nav?: boolean;
  /** Show "current / total" in the toolbar (default true). */
  showSlideNumber?: boolean;
  /** Seconds between slides when autoplaying; `false` disables (default false). */
  autoplay?: number | false;
  /** Pick a random next slide when autoplaying (default false). */
  random?: boolean;
  /** Wrap around after the last slide (default false). */
  loop?: boolean;
  /** Transition effect (default "fade"). */
  transition?: SlideTransition;
  /** Transition duration in seconds (default 0.4). */
  transitionDuration?: number;
  /** Background of the stage in slideshow/fullscreen, e.g. "#000". */
  background?: string;
}

export interface ViewerOptions extends RenderPresentationOptions {
  /** "list" (default) shows all slides; "slideshow" shows one at a time. */
  mode?: ViewerMode;
  /**
   * "fit" (default) scales slides to the container width; "contain" fits each
   * slide inside the container's width and height; a number is a fixed scale.
   */
  scale?: "fit" | "contain" | number;
  /** Slideshow settings. */
  slideshow?: SlideshowOptions;
  /**
   * Keyboard shortcuts: `true` (default) when the viewer has focus,
   * `"global"` to listen on the whole document, `false` to disable.
   */
  keyboard?: boolean | "global";
  /** Show a progress bar while loading (default true). */
  showProgress?: boolean;
  /** Zip-bomb protection limits. */
  limits?: PackageLimits;
  /** Extra options for `fetch()` when loading from a URL (headers, credentials ...). */
  fetchOptions?: RequestInit;
  /** Render slides in chunks and yield to the browser in between (default true). */
  incremental?: boolean;
  /**
   * Direction of the viewer UI (toolbar order, arrow icons and keys, slide
   * transition): "auto" (default) follows the container's CSS direction, so a
   * page with `dir="rtl"` gets a right-to-left viewer. Slide content is not
   * affected; see `textDirection` for paragraphs.
   */
  dir?: "auto" | "ltr" | "rtl";
  /** Gap between slides in list mode, in px (default 24). */
  gap?: number;
}

export interface ViewerEvents {
  /** Fired when loading starts. */
  loadstart: { source: PptxSource };
  /** Rendering progress (0..1). */
  progress: { done: number; total: number; ratio: number };
  /** All slides rendered. */
  load: { presentation: RenderedPresentation };
  /** Loading/rendering failed. */
  error: { error: PptxError };
  /** Non-fatal rendering problem. */
  warning: { warning: PptxError };
  /** Current slide changed (slideshow mode). */
  slidechange: { index: number; slide: RenderedSlide };
  /** A single slide was re-rendered (editor). */
  sliderender: { index: number; element: HTMLElement };
  /** Mode switched. */
  modechange: { mode: ViewerMode };
  /** Fullscreen state changed. */
  fullscreenchange: { fullscreen: boolean };
}

let instanceCounter = 0;

/** Direction of an element from the nearest dir attribute, else its computed CSS direction. */
function elementDirection(el: HTMLElement): "ltr" | "rtl" {
  const attr = el.closest("[dir]")?.getAttribute("dir")?.toLowerCase();
  if (attr === "rtl" || attr === "ltr") return attr;
  try {
    return getComputedStyle(el).direction === "rtl" ? "rtl" : "ltr";
  } catch {
    return "ltr";
  }
}


const svgIcon = (d: string) =>
  `<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true" focusable="false" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${d}</svg>`;
const ICONS: Record<string, string> = {
  first: svgIcon('<path d="M6 5v14M18 6l-8 6 8 6z"/>'),
  prev: svgIcon('<path d="M15 6l-6 6 6 6"/>'),
  next: svgIcon('<path d="M9 6l6 6-6 6"/>'),
  last: svgIcon('<path d="M18 5v14M6 6l8 6-8 6z"/>'),
  play: svgIcon('<path d="M7 5l12 7-12 7z"/>'),
  pause: svgIcon('<path d="M8 5v14M16 5v14"/>'),
  fullscreen: svgIcon('<path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/>'),
};

/**
 * Yields to the browser so it can paint progress. requestAnimationFrame alone
 * never fires in background tabs, so a timer guarantees progress there.
 */
function nextFrame(): Promise<void> {
  return new Promise((resolve) => {
    let done = false;
    const finish = () => {
      if (!done) {
        done = true;
        resolve();
      }
    };
    if (typeof requestAnimationFrame === "function" && !(typeof document !== "undefined" && document.hidden)) requestAnimationFrame(finish);
    setTimeout(finish, 16);
  });
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string, attrs?: Record<string, string>): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (attrs) for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v);
  return e;
}

/** Parses one slide's HTML into an element (each slide separately, so a malformed slide cannot swallow its neighbours). */
function htmlToElement(html: string): HTMLElement {
  const tpl = document.createElement("template");
  tpl.innerHTML = html;
  const first = tpl.content.firstElementChild as HTMLElement | null;
  return first ?? document.createElement("div");
}

export { elementDirection };

export class PptxViewer {
  readonly container: HTMLElement;
  readonly id: string;
  private options: ViewerOptions;
  private readonly emitter = new Emitter<ViewerEvents>();
  private root!: HTMLElement;
  private styleEl!: HTMLStyleElement;
  private stage!: HTMLElement;
  private slidesEl!: HTMLElement;
  private toolbar!: HTMLElement;
  private progressEl!: HTMLElement;
  private counterEl!: HTMLElement;
  private playBtn!: HTMLButtonElement;
  private measureEl: HTMLElement | null = null;
  private wrappers: HTMLElement[] = [];
  private ctx: RenderContext | null = null;
  private pkg: PptxPackage | null = null;
  private presentation: RenderedPresentation | null = null;
  private slideInfos: SlideInfo[] = [];
  private current = 0;
  private mode: ViewerMode;
  private autoplayTimer: ReturnType<typeof setTimeout> | null = null;
  private resizeObserver: ResizeObserver | null = null;
  private abort: AbortController | null = null;
  private loadToken = 0;
  private readonly cleanup: Array<() => void> = [];
  private destroyed = false;

  constructor(container: HTMLElement | string, options: ViewerOptions = {}) {
    const c = typeof container === "string" ? document.querySelector<HTMLElement>(container) : container;
    if (!c) throw new PptxError("INVALID_FILE", `PptxViewer: container ${String(container)} not found`);
    this.container = c;
    this.options = { ...options };
    this.mode = options.mode ?? "list";
    this.id = options.idPrefix ?? `pptxjs-${++instanceCounter}`;
    this.buildDom();
  }

  // ------------------------------------------------------------------ events
  on<K extends keyof ViewerEvents>(event: K, handler: (payload: ViewerEvents[K]) => void): Unsubscribe {
    return this.emitter.on(event, handler);
  }
  off<K extends keyof ViewerEvents>(event: K, handler: (payload: ViewerEvents[K]) => void): void {
    this.emitter.off(event, handler);
  }

  // ------------------------------------------------------------------ getters
  /** Number of slides currently rendered. */
  get slideCount(): number {
    return this.wrappers.length;
  }
  /** 0-based index of the current slide (slideshow mode). */
  get currentIndex(): number {
    return this.current;
  }
  get currentMode(): ViewerMode {
    return this.mode;
  }
  /** The last rendered presentation (HTML strings, CSS, notes ...). */
  get result(): RenderedPresentation | null {
    return this.presentation;
  }
  /** The loaded package (used by the editor and exporters). */
  get package(): PptxPackage | null {
    return this.pkg;
  }
  /** Root element created inside the container. */
  get element(): HTMLElement {
    return this.root;
  }
  /** The `.slide` element of slide `index` (0-based). */
  getSlideElement(index: number): HTMLElement | null {
    return (this.wrappers[index]?.firstElementChild as HTMLElement) ?? null;
  }
  /** Current scale factor applied to slides. */
  get scale(): number {
    return this.computeScale();
  }

  // ------------------------------------------------------------------ loading
  /**
   * Loads and renders a presentation.
   * @param source URL, File/Blob, ArrayBuffer or Uint8Array
   */
  async load(source: PptxSource): Promise<RenderedPresentation> {
    this.assertAlive();
    const token = ++this.loadToken;
    this.abort?.abort();
    const abort = new AbortController();
    this.abort = abort;
    this.reset();
    this.emitter.emit("loadstart", { source });
    this.setProgress(0.01);
    try {
      const bytes = await readSource(source, { signal: abort.signal, fetchOptions: this.options.fetchOptions });
      if (token !== this.loadToken) throw new PptxError("ABORTED", "Superseded by a newer load() call");
      const pkg = PptxPackage.fromBytes(bytes, this.options.limits);
      return await this.renderPackage(pkg, {}, token);
    } catch (e) {
      const error = e instanceof PptxError ? e : new PptxError("INVALID_FILE", (e as Error)?.message ?? String(e), e);
      if (error.code !== "ABORTED") {
        this.setProgress(null);
        this.showError(error);
        this.emitter.emit("error", { error });
      }
      throw error;
    }
  }

  /**
   * Renders an already opened package (used by the editor after modifications).
   * @param options.startIndex slide to show afterwards (slideshow mode)
   */
  async renderPackage(pkg: PptxPackage, options: { startIndex?: number } = {}, token = ++this.loadToken): Promise<RenderedPresentation> {
    const started = Date.now();
    this.pkg = pkg;
    this.ctx?.revokeObjectUrls();
    const ctx = new RenderContext(pkg, {
      ...this.options,
      idPrefix: this.id,
      measureHtmlWidth: this.options.measureHtmlWidth ?? ((html) => this.measure(html)),
      onWarning: (w) => {
        this.options.onWarning?.(w);
        this.emitter.emit("warning", { warning: w });
      },
    });
    this.ctx = ctx;
    const { width, height } = readPresentationInfo(ctx);
    ctx.tableStyles = pkg.xmlFresh("ppt/tableStyles.xml");
    let infos = listSlides(pkg);
    if (this.options.skipHiddenSlides) infos = infos.filter((s) => !s.hidden);
    if (this.options.slides) {
      const wanted = new Set(this.options.slides);
      infos = infos.filter((s) => wanted.has(s.number));
    }
    this.slideInfos = infos;
    this.slidesEl.textContent = "";
    this.wrappers = [];
    const slides: RenderedSlide[] = [];
    const incremental = this.options.incremental !== false;
    let lastYield = Date.now();
    for (let i = 0; i < infos.length; i++) {
      if (token !== this.loadToken) throw new PptxError("ABORTED", "Superseded by a newer load() call");
      const slide = renderSlide(ctx, infos[i], infos[i].number - 1);
      slides.push(slide);
      this.appendSlide(slide, width, height);
      const ratio = (i + 1) / infos.length;
      this.setProgress(ratio);
      this.options.onProgress?.(i + 1, infos.length);
      this.emitter.emit("progress", { done: i + 1, total: infos.length, ratio });
      if (incremental && Date.now() - lastYield > 32) {
        // keep the page responsive on large decks
        this.updateCss(ctx);
        await nextFrame();
        lastYield = Date.now();
      }
    }
    const thumb = pkg.bytes("docProps/thumbnail.jpeg");
    const presentation: RenderedPresentation = {
      width,
      height,
      slides,
      css: genGlobalCSS(ctx),
      thumbnail: thumb ? "data:image/jpeg;base64," + base64ArrayBuffer(thumb) : undefined,
      appVersion: ctx.appVersion,
      renderTime: Date.now() - started,
    };
    this.presentation = presentation;
    this.updateCss(ctx);
    this.setProgress(null);
    this.applyLayout();
    const wanted = options.startIndex ?? (this.options.slideshow?.first ?? 1) - 1;
    const first = Math.min(Math.max(wanted, 0), Math.max(0, slides.length - 1));
    this.current = this.mode === "slideshow" ? first : 0;
    this.showCurrent(false);
    if (this.mode === "slideshow" && this.options.slideshow?.autoplay) this.play();
    this.emitter.emit("load", { presentation });
    return presentation;
  }

  /** Re-renders a single slide from the (possibly modified) package. */
  rerenderSlide(index: number): HTMLElement | null {
    if (!this.ctx || !this.pkg || !this.presentation) return null;
    const info = this.slideInfos[index];
    const wrap = this.wrappers[index];
    if (!info || !wrap) return null;
    const slide = renderSlide(this.ctx, info, info.number - 1);
    this.presentation.slides[index] = slide;
    const element = htmlToElement(slide.html);
    wrap.replaceChildren(element);
    this.updateCss(this.ctx);
    this.applyLayout();
    this.emitter.emit("sliderender", { index, element });
    return element;
  }

  // ------------------------------------------------------------------ navigation
  goTo(index: number): void {
    const n = this.wrappers.length;
    if (!n) return;
    const loop = this.options.slideshow?.loop;
    if (index < 0) index = loop ? n - 1 : 0;
    if (index >= n) index = loop ? 0 : n - 1;
    if (index === this.current && this.mode === "slideshow") return;
    this.current = index;
    if (this.mode === "slideshow") this.showCurrent(true);
    else this.wrappers[index].scrollIntoView({ behavior: "smooth", block: "start" });
    // during incremental loading the presentation object is not complete yet
    const slide = this.presentation?.slides[index];
    if (slide) this.emitter.emit("slidechange", { index, slide });
  }
  next(): void {
    const s = this.options.slideshow;
    if (s?.random && this.autoplayTimer !== null && this.wrappers.length > 1) {
      let r = this.current;
      while (r === this.current) r = Math.floor(Math.random() * this.wrappers.length);
      this.goTo(r);
    } else {
      this.goTo(this.current + 1);
    }
  }
  prev(): void {
    this.goTo(this.current - 1);
  }
  first(): void {
    this.goTo(0);
  }
  last(): void {
    this.goTo(this.wrappers.length - 1);
  }

  /** Starts autoplay (slideshow mode). */
  play(seconds?: number): void {
    const s = seconds ?? (this.options.slideshow?.autoplay || 5);
    this.pause();
    const tick = () => {
      if (!this.options.slideshow?.loop && !this.options.slideshow?.random && this.current >= this.wrappers.length - 1) {
        this.pause();
        return;
      }
      this.next();
      this.autoplayTimer = setTimeout(tick, s * 1000);
    };
    this.autoplayTimer = setTimeout(tick, s * 1000);
    this.updateToolbar();
  }
  pause(): void {
    if (this.autoplayTimer !== null) clearTimeout(this.autoplayTimer);
    this.autoplayTimer = null;
    this.updateToolbar();
  }
  get playing(): boolean {
    return this.autoplayTimer !== null;
  }

  setMode(mode: ViewerMode): void {
    if (mode === this.mode) return;
    this.mode = mode;
    if (mode === "list") this.pause();
    this.applyLayout();
    this.showCurrent(false);
    this.emitter.emit("modechange", { mode });
  }

  /** Updates options at runtime (scale, slideshow settings ...). */
  setOptions(options: Partial<ViewerOptions>): void {
    this.options = { ...this.options, ...options, slideshow: { ...this.options.slideshow, ...options.slideshow } };
    if (options.dir) this.applyDir();
    if (options.mode) this.setMode(options.mode);
    this.applyLayout();
  }

  async enterFullscreen(): Promise<void> {
    if (document.fullscreenElement !== this.root && this.root.requestFullscreen) await this.root.requestFullscreen();
  }
  async exitFullscreen(): Promise<void> {
    if (document.fullscreenElement === this.root) await document.exitFullscreen();
  }
  async toggleFullscreen(): Promise<void> {
    if (document.fullscreenElement === this.root) await this.exitFullscreen();
    else await this.enterFullscreen();
  }

  /** Removes all DOM, listeners, timers and object URLs. */
  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    this.loadToken++;
    this.abort?.abort();
    this.pause();
    this.ctx?.revokeObjectUrls();
    this.resizeObserver?.disconnect();
    for (const fn of this.cleanup.splice(0)) fn();
    this.measureEl?.remove();
    this.root.remove();
    this.emitter.clear();
  }

  /** True when the viewer UI is right to left. */
  get isRtl(): boolean {
    return this.root?.getAttribute("dir") === "rtl";
  }

  private applyDir() {
    const opt = this.options.dir ?? "auto";
    this.root.setAttribute("dir", opt === "auto" ? elementDirection(this.container) : opt);
  }

  // ------------------------------------------------------------------ internals
  private assertAlive() {
    if (this.destroyed) throw new PptxError("ABORTED", "PptxViewer has been destroyed");
  }

  private reset() {
    this.pause();
    this.slidesEl.textContent = "";
    this.wrappers = [];
    this.presentation = null;
    this.current = 0;
    this.root.querySelector(".pptxjs-error")?.remove();
    this.updateToolbar();
  }

  private buildDom() {
    const root = el("div", "pptxjs pptxjs-viewer", { "data-pptxjs-id": this.id, tabindex: "0", role: "region", "aria-roledescription": "presentation" });
    root.setAttribute("aria-label", "Presentation");
    this.styleEl = el("style");
    this.progressEl = el("div", "pptxjs-progress", { role: "progressbar", "aria-valuemin": "0", "aria-valuemax": "100" });
    this.progressEl.appendChild(el("div", "pptxjs-progress-bar"));
    this.progressEl.hidden = true;
    this.stage = el("div", "pptxjs-stage");
    this.slidesEl = el("div", "pptxjs-slides");
    this.stage.appendChild(this.slidesEl);

    this.toolbar = el("div", "pptxjs-toolbar", { role: "toolbar", "aria-label": "Slide navigation" });
    const btn = (cls: string, label: string, icon: string, onClick: () => void) => {
      const b = el("button", "pptxjs-btn " + cls, { type: "button", "aria-label": label, title: label });
      b.innerHTML = ICONS[icon];
      b.addEventListener("click", onClick);
      this.toolbar.appendChild(b);
      return b;
    };
    btn("pptxjs-first", "First slide", "first", () => this.first());
    btn("pptxjs-prev", "Previous slide", "prev", () => this.prev());
    this.counterEl = el("span", "pptxjs-counter", { "aria-live": "polite" });
    this.toolbar.appendChild(this.counterEl);
    btn("pptxjs-next", "Next slide", "next", () => this.next());
    btn("pptxjs-last", "Last slide", "last", () => this.last());
    this.playBtn = btn("pptxjs-play", "Play", "play", () => (this.playing ? this.pause() : this.play()));
    btn("pptxjs-fs-btn", "Fullscreen", "fullscreen", () => void this.toggleFullscreen());

    root.append(this.styleEl, this.progressEl, this.stage, this.toolbar);
    this.container.appendChild(root);
    this.root = root;
    this.applyDir();
    this.styleEl.textContent = BASE_CSS + VIEWER_CSS;

    // keyboard
    const kb = this.options.keyboard ?? true;
    if (kb) {
      const target: HTMLElement | Document = kb === "global" ? document : root;
      const onKey = (ev: Event) => this.onKey(ev as KeyboardEvent);
      target.addEventListener("keydown", onKey);
      this.cleanup.push(() => target.removeEventListener("keydown", onKey));
    }
    const onFs = () => {
      const fs = document.fullscreenElement === this.root;
      this.root.classList.toggle("pptxjs-fullscreen", fs);
      this.applyLayout();
      this.emitter.emit("fullscreenchange", { fullscreen: fs });
    };
    document.addEventListener("fullscreenchange", onFs);
    this.cleanup.push(() => document.removeEventListener("fullscreenchange", onFs));

    // links between slides (<a data-pptx-slide-link>) navigate inside the viewer
    const onClick = (ev: MouseEvent) => {
      const a = (ev.target as Element | null)?.closest?.("a[data-pptx-slide-link]");
      if (!a || !this.root.contains(a)) return;
      ev.preventDefault();
      const file = `ppt/slides/slide${a.getAttribute("data-pptx-slide-link")}.xml`;
      const index = this.slideInfos.findIndex((s) => s.file === file);
      if (index !== -1) this.goTo(index);
    };
    root.addEventListener("click", onClick);
    this.cleanup.push(() => root.removeEventListener("click", onClick));

    if (typeof ResizeObserver !== "undefined") {
      this.resizeObserver = new ResizeObserver(() => this.applyLayout());
      this.resizeObserver.observe(this.container);
    }
  }

  private onKey(ev: KeyboardEvent) {
    const t = ev.target as HTMLElement | null;
    if (t && (t.isContentEditable || /^(INPUT|TEXTAREA|SELECT)$/.test(t.tagName))) return;
    if (ev.altKey || ev.ctrlKey || ev.metaKey) return;
    let handled = true;
    // in a right-to-left UI the left arrow moves forward
    let key = ev.key;
    if (this.isRtl) key = key === "ArrowLeft" ? "ArrowRight" : key === "ArrowRight" ? "ArrowLeft" : key;
    switch (key) {
      case "ArrowRight":
      case "ArrowDown":
      case "PageDown":
      case " ":
        this.next();
        break;
      case "ArrowLeft":
      case "ArrowUp":
      case "PageUp":
      case "Backspace":
        this.prev();
        break;
      case "Home":
        this.first();
        break;
      case "End":
        this.last();
        break;
      case "f":
      case "F":
        void this.toggleFullscreen();
        break;
      case "p":
      case "P":
        if (this.mode === "slideshow") {
          if (this.playing) this.pause();
          else this.play();
        }
        break;
      default:
        handled = false;
    }
    if (handled && (this.mode === "slideshow" || ev.key !== " ")) ev.preventDefault();
  }

  private appendSlide(slide: RenderedSlide, width: number, height: number) {
    const wrap = el("div", "pptxjs-slide-wrap", { "data-index": String(this.wrappers.length) });
    wrap.style.setProperty("--pptx-w", width + "px");
    wrap.style.setProperty("--pptx-h", height + "px");
    wrap.setAttribute("role", "group");
    wrap.setAttribute("aria-roledescription", "slide");
    wrap.setAttribute("aria-label", `Slide ${slide.number}`);
    wrap.appendChild(htmlToElement(slide.html));
    this.slidesEl.appendChild(wrap);
    this.wrappers.push(wrap);
    this.sizeWrapper(wrap, this.computeScale());
  }

  private updateCss(ctx: RenderContext) {
    const scope = `[data-pptxjs-id="${this.id}"]`;
    const generated = genGlobalCSS(ctx)
      .split("\n")
      .filter(Boolean)
      .map((rule) => scope + rule)
      .join("\n");
    // only plain colour values are accepted (no CSS injection through options)
    const bg = (this.options.slideshow?.background ?? "").replace(/[^#a-zA-Z0-9(),.% -]/g, "");
    const extra = bg ? `${scope}.pptxjs-mode-slideshow .pptxjs-stage, ${scope}.pptxjs-fullscreen { background: ${bg}; }` : "";
    this.styleEl.textContent = BASE_CSS + VIEWER_CSS + extra + "\n" + generated;
  }

  private computeScale(): number {
    const s = this.options.scale ?? "fit";
    if (typeof s === "number") return s > 0 ? s : 1;
    const w = this.presentation?.width ?? this.ctx?.slideWidth ?? 0;
    const h = this.presentation?.height ?? this.ctx?.slideHeight ?? 0;
    if (!w) return 1;
    const fullscreen = document.fullscreenElement === this.root;
    const avail = (fullscreen ? window.innerWidth : this.stage.clientWidth || this.container.clientWidth) - (fullscreen ? 0 : 4);
    let scale = avail > 0 ? avail / w : 1;
    if ((this.mode === "slideshow" || fullscreen || s === "contain") && h) {
      const toolbarH = this.toolbar.hidden ? 0 : this.toolbar.offsetHeight || 44;
      let availH = Infinity;
      if (fullscreen) availH = window.innerHeight - toolbarH;
      else if (s === "contain") {
        const cs = getComputedStyle(this.container);
        const pad = (parseFloat(cs.paddingTop) || 0) + (parseFloat(cs.paddingBottom) || 0);
        availH = this.container.clientHeight - pad - toolbarH - 2;
      }
      if (availH > 0) scale = Math.min(scale, availH / h);
    }
    return scale > 0 && Number.isFinite(scale) ? scale : 1;
  }

  private sizeWrapper(wrap: HTMLElement, scale: number) {
    wrap.style.setProperty("--pptx-scale", String(scale));
  }

  private applyLayout() {
    if (this.destroyed) return;
    const ss = this.mode === "slideshow";
    this.root.classList.toggle("pptxjs-mode-slideshow", ss);
    this.root.classList.toggle("pptxjs-mode-list", !ss);
    const nav = this.options.slideshow?.nav ?? true;
    this.toolbar.hidden = !(ss && nav) || this.wrappers.length === 0;
    this.root.style.setProperty("--pptx-gap", (this.options.gap ?? 24) + "px");
    const dur = this.options.slideshow?.transitionDuration ?? 0.4;
    this.root.style.setProperty("--pptx-transition", dur + "s");
    const scale = this.computeScale();
    for (const w of this.wrappers) this.sizeWrapper(w, scale);
    this.updateToolbar();
  }

  private showCurrent(animate: boolean) {
    const ss = this.mode === "slideshow";
    let transition = this.options.slideshow?.transition ?? "fade";
    if (transition === "random") transition = (["fade", "slide", "zoom"] as const)[Math.floor(Math.random() * 3)];
    this.wrappers.forEach((w, i) => {
      const active = i === this.current;
      w.classList.toggle("pptxjs-active", active);
      w.hidden = ss && !active;
      if (ss && active && animate && transition !== "none") {
        w.classList.remove("pptxjs-anim-fade", "pptxjs-anim-slide", "pptxjs-anim-zoom");
        void w.offsetWidth; // restart animation
        w.classList.add("pptxjs-anim-" + transition);
      }
      w.setAttribute("aria-hidden", ss && !active ? "true" : "false");
    });
    this.updateToolbar();
  }

  private updateToolbar() {
    if (!this.counterEl) return;
    const total = this.wrappers.length;
    const show = this.options.slideshow?.showSlideNumber ?? true;
    this.counterEl.textContent = show && total ? `${this.current + 1} / ${total}` : "";
    if (this.playBtn) {
      this.playBtn.innerHTML = this.playing ? ICONS.pause : ICONS.play;
      this.playBtn.setAttribute("aria-label", this.playing ? "Pause" : "Play");
      this.playBtn.title = this.playing ? "Pause" : "Play";
    }
  }

  private setProgress(ratio: number | null) {
    if (this.options.showProgress === false) return;
    if (ratio === null) {
      this.progressEl.hidden = true;
      return;
    }
    this.progressEl.hidden = false;
    const pct = Math.round(ratio * 100);
    (this.progressEl.firstElementChild as HTMLElement).style.width = pct + "%";
    this.progressEl.setAttribute("aria-valuenow", String(pct));
  }

  private showError(error: PptxError) {
    const box = el("div", "pptxjs-error", { role: "alert" });
    box.textContent = error.message;
    this.stage.prepend(box);
  }

  /** Measures the width of an HTML fragment (used for bullet layout). */
  private measure(html: string): number {
    if (!this.measureEl) {
      this.measureEl = el("div", "pptxjs pptxjs-measure", { "aria-hidden": "true" });
      this.measureEl.style.cssText = "position:absolute;left:-10000px;top:0;visibility:hidden;white-space:nowrap;pointer-events:none;";
      document.body.appendChild(this.measureEl);
    }
    this.measureEl.innerHTML = html;
    const first = this.measureEl.firstElementChild as HTMLElement | null;
    if (!first) return 0;
    first.style.position = "absolute";
    first.style.whiteSpace = "nowrap";
    const w = first.getBoundingClientRect().width;
    this.measureEl.textContent = "";
    return w;
  }
}
