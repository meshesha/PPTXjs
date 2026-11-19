/**
 * PptxEditor: in-browser presentation editor built on PptxViewer (rendering)
 * and PresentationDocument (lossless OOXML model).
 *
 * Features: slide thumbnails, select / move / resize shapes, inline text
 * editing (double click), text formatting, fill colour, add text box / shape /
 * picture, duplicate / delete / arrange shapes, add / duplicate / delete /
 * reorder / hide slides, undo / redo, keyboard shortcuts and export to .pptx
 * and .odp.
 */
import { PptxPackage, type PackageLimits } from "../core/package";
import { PptxError } from "../errors";
import { PresentationDocument, type Bounds, type ShapeInfo, type PresetGeometry } from "../model/document";
import { readSource, type PptxSource } from "../source";
import { Emitter, type Unsubscribe } from "../utils/emitter";
import { PptxViewer, elementDirection, type ViewerOptions } from "../viewer/viewer";
import { exportPptx } from "../export/pptx";
import { exportOdp } from "../export/odp";
import { EDITOR_CSS } from "./editor-styles";

export interface EditorOptions extends Omit<ViewerOptions, "mode" | "keyboard" | "slideshow"> {
  /** Show the built-in toolbar (default true). */
  toolbar?: boolean;
  /** Show the slide thumbnail panel (default true). */
  thumbnails?: boolean;
  /** Disable all editing (selection still works). */
  readOnly?: boolean;
  /** File name used by download() without extension (default "presentation"). */
  fileName?: string;
  limits?: PackageLimits;
}

export interface ShapeRef {
  slide: number;
  id: string;
}

export interface EditorEvents {
  load: { slides: number };
  error: { error: PptxError };
  selectionchange: { selection: ShapeRef | null; shape: ShapeInfo | null };
  slidechange: { index: number };
  change: { label: string };
}

const HANDLES = ["nw", "n", "ne", "e", "se", "s", "sw", "w"] as const;
type Handle = (typeof HANDLES)[number];

const SHAPES: Array<[PresetGeometry, string]> = [
  ["rect", "Rectangle"],
  ["roundRect", "Rounded rectangle"],
  ["ellipse", "Ellipse"],
  ["triangle", "Triangle"],
  ["diamond", "Diamond"],
  ["pentagon", "Pentagon"],
  ["hexagon", "Hexagon"],
  ["star5", "Star"],
  ["rightArrow", "Arrow right"],
  ["leftArrow", "Arrow left"],
  ["upArrow", "Arrow up"],
  ["downArrow", "Arrow down"],
  ["heart", "Heart"],
  ["cloud", "Cloud"],
];

function h<K extends keyof HTMLElementTagNameMap>(tag: K, attrs: Record<string, string> = {}, text?: string): HTMLElementTagNameMap[K] {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === "class") e.className = v;
    else e.setAttribute(k, v);
  }
  if (text !== undefined) e.textContent = text;
  return e;
}

export class PptxEditor {
  readonly container: HTMLElement;
  readonly viewer: PptxViewer;
  private doc: PresentationDocument | null = null;
  private readonly options: EditorOptions;
  private readonly emitter = new Emitter<EditorEvents>();
  private readonly root: HTMLElement;
  private readonly toolbarEl: HTMLElement;
  private readonly thumbsEl: HTMLElement;
  private readonly stageEl: HTMLElement;
  private readonly overlay: HTMLElement;
  private readonly fileInput: HTMLInputElement;
  private readonly controls: Record<string, HTMLElement> = {};
  private selection: ShapeRef | null = null;
  private textEditor: HTMLTextAreaElement | null = null;
  private drag: {
    kind: "move" | Handle;
    startX: number;
    startY: number;
    start: Bounds;
    current: Bounds;
    moved: boolean;
    pointerId: number;
  } | null = null;
  private readonly cleanup: Array<() => void> = [];
  private docUnsub: Unsubscribe | null = null;
  private rendering: Promise<unknown> = Promise.resolve();

  constructor(container: HTMLElement | string, options: EditorOptions = {}) {
    const c = typeof container === "string" ? document.querySelector<HTMLElement>(container) : container;
    if (!c) throw new PptxError("INVALID_FILE", `PptxEditor: container ${String(container)} not found`);
    this.container = c;
    this.options = options;

    this.root = h("div", { class: "pptxjs-editor", tabindex: "0" });
    this.toolbarEl = h("div", { class: "pptxjs-editor-toolbar", role: "toolbar", "aria-label": "Editor" });
    const body = h("div", { class: "pptxjs-editor-body" });
    this.thumbsEl = h("div", { class: "pptxjs-editor-thumbs", role: "listbox", "aria-label": "Slides" });
    this.stageEl = h("div", { class: "pptxjs-editor-stage" });
    body.append(this.thumbsEl, this.stageEl);
    const style = h("style");
    style.textContent = EDITOR_CSS;
    this.fileInput = h("input", { type: "file", accept: "image/png,image/jpeg,image/gif,image/svg+xml,image/webp,image/bmp", hidden: "" });
    this.root.append(style, this.toolbarEl, body, this.fileInput);
    this.container.appendChild(this.root);
    // UI direction: explicit option, or the container's CSS direction (dir="rtl" pages)
    const dirOpt = options.dir ?? "auto";
    const uiDir = dirOpt === "auto" ? elementDirection(this.container) : dirOpt;
    this.root.setAttribute("dir", uiDir);
    if (options.toolbar === false) this.toolbarEl.hidden = true;
    if (options.thumbnails === false) this.thumbsEl.hidden = true;

    this.viewer = new PptxViewer(this.stageEl, {
      ...options,
      scale: options.scale ?? "contain",
      mode: "slideshow",
      dir: uiDir,
      keyboard: false,
      slideshow: { nav: false, transition: "none" },
      gap: 0,
    });
    this.overlay = h("div", { class: "pptxjs-editor-overlay" });
    this.overlay.hidden = true;
    for (const hd of HANDLES) this.overlay.appendChild(h("div", { class: `pptxjs-handle pptxjs-handle-${hd}`, "data-handle": hd }));

    this.buildToolbar();
    this.bindEvents();
    this.viewer.on("error", ({ error }) => this.emitter.emit("error", { error }));
    this.viewer.on("slidechange", ({ index }) => {
      this.updateThumbSelection();
      this.emitter.emit("slidechange", { index });
    });
  }

  // ------------------------------------------------------------------ public API
  on<K extends keyof EditorEvents>(event: K, handler: (p: EditorEvents[K]) => void): Unsubscribe {
    return this.emitter.on(event, handler);
  }

  /** The editable document (available after load()). */
  get document(): PresentationDocument | null {
    return this.doc;
  }

  get currentSlide(): number {
    return this.viewer.currentIndex;
  }

  get selected(): ShapeRef | null {
    return this.selection;
  }

  /** Loads a presentation for editing. */
  async load(source: PptxSource): Promise<void> {
    try {
      const bytes = await readSource(source, { fetchOptions: this.options.fetchOptions });
      const pkg = PptxPackage.fromBytes(bytes, this.options.limits);
      await this.open(new PresentationDocument(pkg));
    } catch (e) {
      const error = e instanceof PptxError ? e : new PptxError("INVALID_FILE", (e as Error).message, e);
      this.emitter.emit("error", { error });
      throw error;
    }
  }

  /** Opens an existing PresentationDocument. */
  async open(doc: PresentationDocument): Promise<void> {
    this.docUnsub?.();
    this.doc = doc;
    this.select(null);
    this.docUnsub = doc.on("change", (change) => void this.onDocChange(change.type, change.index, change.label));
    await this.viewer.renderPackage(doc.pkg, { startIndex: 0 });
    this.buildThumbs();
    this.updateToolbar();
    this.emitter.emit("load", { slides: doc.slideCount });
  }

  goToSlide(index: number): void {
    this.commitText();
    this.select(null);
    this.viewer.goTo(index);
  }

  /** Selects a shape (or clears the selection with `null`). */
  select(ref: ShapeRef | null): void {
    this.selection = ref;
    this.positionOverlay();
    this.updateToolbar();
    let shape: ShapeInfo | null = null;
    if (ref && this.doc) {
      try {
        shape = this.doc.getShape(ref.slide, ref.id);
      } catch {
        shape = null;
      }
    }
    this.emitter.emit("selectionchange", { selection: ref, shape });
  }

  undo(): void {
    this.commitText();
    this.doc?.undo();
  }
  redo(): void {
    this.commitText();
    this.doc?.redo();
  }

  /** Exports the edited presentation as .pptx. */
  exportPptx(): Blob {
    this.commitText();
    if (!this.doc) throw new PptxError("EXPORT_FAILED", "No document loaded");
    return exportPptx(this.doc);
  }

  /** Exports the edited presentation as .odp. */
  exportOdp(): Blob {
    this.commitText();
    if (!this.doc) throw new PptxError("EXPORT_FAILED", "No document loaded");
    return exportOdp(this.doc);
  }

  /** Triggers a browser download of the presentation. */
  download(format: "pptx" | "odp" = "pptx", fileName = this.options.fileName ?? "presentation"): void {
    const blob = format === "odp" ? this.exportOdp() : this.exportPptx();
    const url = URL.createObjectURL(blob);
    const a = h("a", { href: url, download: `${fileName}.${format}` });
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  destroy(): void {
    this.docUnsub?.();
    for (const fn of this.cleanup.splice(0)) fn();
    this.viewer.destroy();
    this.root.remove();
  }

  // ------------------------------------------------------------------ document sync
  private async onDocChange(type: "slide" | "structure", index: number | undefined, label: string): Promise<void> {
    const run = async () => {
      if (!this.doc) return;
      if (type === "structure" || index === undefined) {
        const keep = Math.min(this.viewer.currentIndex, this.doc.slideCount - 1);
        await this.viewer.renderPackage(this.doc.pkg, { startIndex: Math.max(0, keep) });
        this.buildThumbs();
      } else {
        this.viewer.rerenderSlide(index);
        this.updateThumb(index);
        if (index !== this.viewer.currentIndex) this.viewer.goTo(index);
      }
      // drop the selection if the shape disappeared
      if (this.selection) {
        try {
          this.doc.getShape(this.selection.slide, this.selection.id);
        } catch {
          this.selection = null;
        }
      }
      this.positionOverlay();
      this.updateToolbar();
      this.emitter.emit("change", { label });
    };
    this.rendering = this.rendering.then(run, run);
    await this.rendering;
  }

  /** Runs a document edit, reporting errors instead of throwing from UI handlers. */
  private edit(fn: (doc: PresentationDocument, slide: number) => void): void {
    if (!this.doc || this.options.readOnly) return;
    try {
      fn(this.doc, this.viewer.currentIndex);
    } catch (e) {
      const error = e instanceof PptxError ? e : new PptxError("RENDER_WARNING", (e as Error).message, e);
      this.emitter.emit("error", { error });
    }
  }

  // ------------------------------------------------------------------ thumbnails
  private buildThumbs(): void {
    this.thumbsEl.textContent = "";
    for (let i = 0; i < this.viewer.slideCount; i++) {
      const item = h("div", { class: "pptxjs-thumb", role: "option", tabindex: "-1", "data-index": String(i) });
      item.appendChild(h("span", { class: "pptxjs-thumb-num" }, String(i + 1)));
      item.appendChild(h("div", { class: "pptxjs-thumb-frame" }));
      item.addEventListener("click", () => this.goToSlide(i));
      this.thumbsEl.appendChild(item);
      this.updateThumb(i);
    }
    this.updateThumbSelection();
  }

  private updateThumb(index: number): void {
    const item = this.thumbsEl.children[index] as HTMLElement | undefined;
    const slide = this.viewer.getSlideElement(index);
    const res = this.viewer.result;
    if (!item || !slide || !res) return;
    const frame = item.querySelector(".pptxjs-thumb-frame") as HTMLElement;
    const width = 150;
    const scale = width / res.width;
    frame.style.width = width + "px";
    frame.style.height = res.height * scale + "px";
    const clone = slide.cloneNode(true) as HTMLElement;
    clone.querySelectorAll("video, audio, iframe").forEach((e) => e.remove());
    clone.removeAttribute("id");
    clone.querySelectorAll("[id]").forEach((e) => e.removeAttribute("id"));
    clone.style.transform = `scale(${scale})`;
    clone.style.transformOrigin = "0 0";
    clone.style.pointerEvents = "none";
    frame.replaceChildren(clone);
    item.classList.toggle("pptxjs-thumb-hidden", slide.classList.contains("slide-hidden"));
  }

  private updateThumbSelection(): void {
    Array.from(this.thumbsEl.children).forEach((c, i) => {
      const active = i === this.viewer.currentIndex;
      c.classList.toggle("pptxjs-thumb-active", active);
      c.setAttribute("aria-selected", String(active));
      if (active) (c as HTMLElement).scrollIntoView?.({ block: "nearest" });
    });
  }

  // ------------------------------------------------------------------ toolbar
  private buildToolbar(): void {
    const group = () => {
      const g = h("div", { class: "pptxjs-tb-group" });
      this.toolbarEl.appendChild(g);
      return g;
    };
    const button = (g: HTMLElement, key: string, label: string, text: string, onClick: () => void) => {
      const b = h("button", { type: "button", class: "pptxjs-tb-btn", title: label, "aria-label": label }, text);
      b.addEventListener("click", () => {
        onClick();
        this.root.focus({ preventScroll: true });
      });
      g.appendChild(b);
      this.controls[key] = b;
      return b;
    };
    const sel = (fn: (doc: PresentationDocument, ref: ShapeRef) => void) => () => {
      const ref = this.selection;
      if (ref) this.edit((doc) => fn(doc, ref));
    };

    let g = group();
    button(g, "undo", "Undo (Ctrl+Z)", "↶", () => this.undo());
    button(g, "redo", "Redo (Ctrl+Y)", "↷", () => this.redo());

    g = group();
    button(g, "addText", "Add text box", "T+", () =>
      this.edit((doc, s) => {
        const size = doc.slideSize();
        const id = doc.addTextBox(s, { x: size.width / 2 - 150, y: size.height / 2 - 30, width: 300, height: 60 }, "New text", { size: 24 });
        this.select({ slide: s, id });
      }),
    );
    const shapeSel = h("select", { class: "pptxjs-tb-select", title: "Add shape", "aria-label": "Add shape" });
    shapeSel.appendChild(h("option", { value: "" }, "Shape…"));
    for (const [geom, label] of SHAPES) shapeSel.appendChild(h("option", { value: geom }, label));
    shapeSel.addEventListener("change", () => {
      const geom = shapeSel.value;
      shapeSel.value = "";
      if (!geom) return;
      this.edit((doc, s) => {
        const size = doc.slideSize();
        const id = doc.addShape(s, geom, { x: size.width / 2 - 80, y: size.height / 2 - 60, width: 160, height: 120 });
        this.select({ slide: s, id });
      });
    });
    g.appendChild(shapeSel);
    this.controls.addShape = shapeSel;
    button(g, "addImage", "Add picture", "🖼", () => this.fileInput.click());

    g = group();
    button(g, "bold", "Bold (Ctrl+B)", "B", sel((doc, r) => doc.setTextStyle(r.slide, r.id, { bold: !this.firstRun()?.bold })));
    button(g, "italic", "Italic (Ctrl+I)", "I", sel((doc, r) => doc.setTextStyle(r.slide, r.id, { italic: !this.firstRun()?.italic })));
    button(g, "underline", "Underline (Ctrl+U)", "U", sel((doc, r) => doc.setTextStyle(r.slide, r.id, { underline: !this.firstRun()?.underline })));
    const size = h("input", { type: "number", min: "1", max: "400", step: "1", class: "pptxjs-tb-num", title: "Font size (pt)", "aria-label": "Font size" });
    size.addEventListener("change", sel((doc, r) => doc.setTextStyle(r.slide, r.id, { size: Number(size.value) || 18 })));
    g.appendChild(size);
    this.controls.size = size;
    const color = h("input", { type: "color", class: "pptxjs-tb-color", title: "Text colour", "aria-label": "Text colour" });
    color.addEventListener("change", sel((doc, r) => doc.setTextStyle(r.slide, r.id, { color: color.value })));
    g.appendChild(color);
    this.controls.color = color;
    const align = h("select", { class: "pptxjs-tb-select", title: "Alignment", "aria-label": "Alignment" });
    for (const [v, l] of [["", "Align…"], ["left", "Left"], ["center", "Center"], ["right", "Right"], ["justify", "Justify"]]) align.appendChild(h("option", { value: v }, l));
    align.addEventListener("change", () => {
      const v = align.value as "left" | "center" | "right" | "justify" | "";
      align.value = "";
      if (v) sel((doc, r) => doc.setTextStyle(r.slide, r.id, { align: v }))();
    });
    g.appendChild(align);
    this.controls.align = align;

    g = group();
    const fill = h("input", { type: "color", class: "pptxjs-tb-color", title: "Fill colour", "aria-label": "Fill colour" });
    fill.addEventListener("change", sel((doc, r) => doc.setShapeFill(r.slide, r.id, fill.value)));
    g.appendChild(h("span", { class: "pptxjs-tb-label" }, "Fill"));
    g.appendChild(fill);
    this.controls.fill = fill;
    button(g, "noFill", "No fill", "∅", sel((doc, r) => doc.setShapeFill(r.slide, r.id, null)));
    button(g, "front", "Bring to front", "⭱", sel((doc, r) => doc.reorderShape(r.slide, r.id, "front")));
    button(g, "back", "Send to back", "⭳", sel((doc, r) => doc.reorderShape(r.slide, r.id, "back")));
    button(g, "dupShape", "Duplicate (Ctrl+D)", "⎘", sel((doc, r) => this.select({ slide: r.slide, id: doc.duplicateShape(r.slide, r.id) })));
    button(g, "delShape", "Delete (Del)", "🗑", () => this.deleteSelection());

    g = group();
    g.appendChild(h("span", { class: "pptxjs-tb-label" }, "Slide"));
    button(g, "addSlide", "New slide", "+", () => this.edit((doc, s) => this.afterStructure(doc.addSlide(s, s))));
    button(g, "dupSlide", "Duplicate slide", "⎘", () => this.edit((doc, s) => this.afterStructure(doc.duplicateSlide(s))));
    button(g, "upSlide", "Move slide up", "↑", () => this.edit((doc, s) => s > 0 && (doc.moveSlide(s, s - 1), this.afterStructure(s - 1))));
    button(g, "downSlide", "Move slide down", "↓", () => this.edit((doc, s) => s < doc.slideCount - 1 && (doc.moveSlide(s, s + 1), this.afterStructure(s + 1))));
    button(g, "hideSlide", "Hide / show slide", "👁", () => this.edit((doc, s) => doc.setSlideHidden(s, !doc.slides()[s].hidden)));
    button(g, "delSlide", "Delete slide", "✕", () =>
      this.edit((doc, s) => {
        if (doc.slideCount > 1) {
          doc.deleteSlide(s);
          this.afterStructure(Math.max(0, s - 1));
        }
      }),
    );

    g = group();
    button(g, "exportPptx", "Download .pptx", "⬇ PPTX", () => this.download("pptx"));
    button(g, "exportOdp", "Download .odp", "⬇ ODP", () => this.download("odp"));
    this.updateToolbar();
  }

  private afterStructure(index: number): void {
    // structural changes re-render asynchronously; jump once done
    void this.rendering.then(() => this.goToSlide(index));
  }

  private firstRun() {
    if (!this.selection || !this.doc) return undefined;
    try {
      return this.doc.getShape(this.selection.slide, this.selection.id).paragraphs?.[0]?.runs[0];
    } catch {
      return undefined;
    }
  }

  private updateToolbar(): void {
    const has = !!this.selection;
    const ro = !!this.options.readOnly || !this.doc;
    let shape: ShapeInfo | null = null;
    if (has && this.doc) {
      try {
        shape = this.doc.getShape(this.selection!.slide, this.selection!.id);
      } catch {
        shape = null;
      }
    }
    const textual = !!shape && shape.paragraphs !== null;
    const set = (key: string, enabled: boolean) => {
      const c = this.controls[key] as HTMLButtonElement | undefined;
      if (c) c.disabled = !enabled;
    };
    set("undo", !ro && !!this.doc?.canUndo);
    set("redo", !ro && !!this.doc?.canRedo);
    for (const k of ["addText", "addShape", "addImage", "addSlide", "dupSlide", "upSlide", "downSlide", "hideSlide", "delSlide"]) set(k, !ro);
    for (const k of ["bold", "italic", "underline", "size", "color", "align"]) set(k, !ro && textual);
    for (const k of ["fill", "noFill"]) set(k, !ro && has && shape?.kind !== "group" && shape?.kind !== "table" && shape?.kind !== "chart");
    for (const k of ["front", "back", "dupShape", "delShape"]) set(k, !ro && has);
    set("exportPptx", !!this.doc);
    set("exportOdp", !!this.doc);
    const run = shape?.paragraphs?.[0]?.runs[0];
    (this.controls.bold as HTMLElement | undefined)?.classList.toggle("pptxjs-on", !!run?.bold);
    (this.controls.italic as HTMLElement | undefined)?.classList.toggle("pptxjs-on", !!run?.italic);
    (this.controls.underline as HTMLElement | undefined)?.classList.toggle("pptxjs-on", !!run?.underline);
    if (this.controls.size) (this.controls.size as HTMLInputElement).value = run?.size ? String(run.size) : "";
    if (this.controls.color && run?.color) (this.controls.color as HTMLInputElement).value = run.color.toLowerCase();
    if (this.controls.fill && shape?.fill) (this.controls.fill as HTMLInputElement).value = shape.fill.toLowerCase();
  }

  // ------------------------------------------------------------------ selection overlay
  private positionOverlay(): void {
    const ref = this.selection;
    if (!ref || !this.doc || ref.slide !== this.viewer.currentIndex) {
      this.overlay.hidden = true;
      return;
    }
    let b: Bounds | null = null;
    let rot = 0;
    try {
      const s = this.doc.getShape(ref.slide, ref.id);
      b = s.bounds;
      rot = s.rotation;
    } catch {
      b = null;
    }
    const slideEl = this.viewer.getSlideElement(ref.slide);
    const wrap = slideEl?.parentElement;
    if (!b || !wrap) {
      this.overlay.hidden = true;
      return;
    }
    if (this.overlay.parentElement !== wrap) wrap.appendChild(this.overlay);
    this.placeOverlay(this.drag?.current ?? b, rot);
    this.overlay.hidden = false;
  }

  private placeOverlay(b: Bounds, rot = 0): void {
    const s = this.viewer.scale;
    Object.assign(this.overlay.style, {
      left: b.x * s + "px",
      top: b.y * s + "px",
      width: Math.max(1, b.width * s) + "px",
      height: Math.max(1, b.height * s) + "px",
      transform: rot ? `rotate(${rot}deg)` : "",
    });
  }

  private shapeElements(ref: ShapeRef): HTMLElement[] {
    const wrapper = this.viewer.getSlideElement(ref.slide)?.querySelector(`[data-shape-id="${CSS.escape(ref.id)}"]`);
    return wrapper ? (Array.from(wrapper.children) as HTMLElement[]) : [];
  }

  private deleteSelection(): void {
    const ref = this.selection;
    if (!ref) return;
    this.edit((doc) => doc.deleteShape(ref.slide, ref.id));
    this.select(null);
  }

  // ------------------------------------------------------------------ events
  private bindEvents(): void {
    const on = <K extends keyof HTMLElementEventMap>(el: HTMLElement | Document, type: K, fn: (ev: HTMLElementEventMap[K]) => void) => {
      el.addEventListener(type, fn as EventListener);
      this.cleanup.push(() => el.removeEventListener(type, fn as EventListener));
    };

    on(this.stageEl, "pointerdown", (ev) => this.onPointerDown(ev));
    on(this.stageEl, "dblclick", (ev) => this.onDoubleClick(ev));
    on(this.root, "keydown", (ev) => this.onKey(ev));
    on(document, "pointermove", (ev) => this.onPointerMove(ev as PointerEvent));
    on(document, "pointerup", (ev) => this.onPointerUp(ev as PointerEvent));
    on(this.fileInput, "change", () => void this.onImagePicked());
    const ro = typeof ResizeObserver !== "undefined" ? new ResizeObserver(() => this.positionOverlay()) : null;
    ro?.observe(this.stageEl);
    this.cleanup.push(() => ro?.disconnect());
    this.viewer.on("slidechange", () => this.positionOverlay());
  }

  private onPointerDown(ev: PointerEvent): void {
    if (ev.button !== 0 || !this.doc) return;
    const target = ev.target as HTMLElement;
    if (this.textEditor && target === this.textEditor) return;
    this.commitText();
    const handle = target.closest<HTMLElement>("[data-handle]");
    if (handle && this.selection && !this.options.readOnly) {
      this.startDrag(ev, handle.dataset.handle as Handle);
      return;
    }
    if (target.closest(".pptxjs-editor-overlay") && this.selection && !this.options.readOnly) {
      this.startDrag(ev, "move");
      return;
    }
    const slideEl = this.viewer.getSlideElement(this.viewer.currentIndex);
    if (!slideEl || !slideEl.contains(target)) {
      if (!target.closest(".pptxjs-editor-overlay")) this.select(null);
      return;
    }
    const shapeEl = target.closest<HTMLElement>("[data-shape-id]");
    if (!shapeEl) {
      this.select(null);
      return;
    }
    this.select({ slide: this.viewer.currentIndex, id: shapeEl.dataset.shapeId! });
    if (!this.options.readOnly) this.startDrag(ev, "move");
    ev.preventDefault();
  }

  private startDrag(ev: PointerEvent, kind: "move" | Handle): void {
    const ref = this.selection;
    if (!ref || !this.doc) return;
    let b: Bounds | null = null;
    try {
      b = this.doc.getShape(ref.slide, ref.id).bounds;
    } catch {
      b = null;
    }
    if (!b) return;
    this.drag = { kind, startX: ev.clientX, startY: ev.clientY, start: { ...b }, current: { ...b }, moved: false, pointerId: ev.pointerId };
    ev.preventDefault();
  }

  private onPointerMove(ev: PointerEvent): void {
    const d = this.drag;
    if (!d || ev.pointerId !== d.pointerId) return;
    const s = this.viewer.scale || 1;
    const dx = (ev.clientX - d.startX) / s;
    const dy = (ev.clientY - d.startY) / s;
    if (!d.moved && Math.abs(dx) + Math.abs(dy) < 2) return;
    d.moved = true;
    const b = { ...d.start };
    if (d.kind === "move") {
      b.x += dx;
      b.y += dy;
      if (ev.shiftKey) {
        if (Math.abs(dx) > Math.abs(dy)) b.y = d.start.y;
        else b.x = d.start.x;
      }
    } else {
      if (d.kind.includes("e")) b.width = Math.max(4, d.start.width + dx);
      if (d.kind.includes("s")) b.height = Math.max(4, d.start.height + dy);
      if (d.kind.includes("w")) {
        b.width = Math.max(4, d.start.width - dx);
        b.x = d.start.x + d.start.width - b.width;
      }
      if (d.kind.includes("n")) {
        b.height = Math.max(4, d.start.height - dy);
        b.y = d.start.y + d.start.height - b.height;
      }
      if (ev.shiftKey && d.start.height > 0) {
        // keep aspect ratio
        const ratio = d.start.width / d.start.height;
        if (d.kind === "n" || d.kind === "s") b.width = b.height * ratio;
        else b.height = b.width / ratio;
      }
    }
    d.current = b;
    this.placeOverlay(b);
    if (d.kind === "move" && this.selection) {
      // live preview: shift the rendered elements
      const tx = (b.x - d.start.x) * s;
      const ty = (b.y - d.start.y) * s;
      for (const el of this.shapeElements(this.selection)) el.style.translate = `${tx / s}px ${ty / s}px`;
    }
  }

  private onPointerUp(ev: PointerEvent): void {
    const d = this.drag;
    if (!d || ev.pointerId !== d.pointerId) return;
    this.drag = null;
    const ref = this.selection;
    if (!d.moved || !ref) return;
    const c = d.current;
    const round = (n: number) => Math.round(n * 100) / 100;
    this.edit((doc) => doc.setShapeBounds(ref.slide, ref.id, { x: round(c.x), y: round(c.y), width: round(c.width), height: round(c.height) }));
  }

  private onDoubleClick(ev: MouseEvent): void {
    if (this.options.readOnly || !this.doc) return;
    const ref = this.selection;
    if (!ref) return;
    let shape: ShapeInfo;
    try {
      shape = this.doc.getShape(ref.slide, ref.id);
    } catch {
      return;
    }
    if (shape.paragraphs === null || !shape.bounds) return;
    ev.preventDefault();
    this.openTextEditor(ref, shape);
  }

  private openTextEditor(ref: ShapeRef, shape: ShapeInfo): void {
    const wrap = this.viewer.getSlideElement(ref.slide)?.parentElement;
    if (!wrap || !shape.bounds) return;
    const s = this.viewer.scale;
    const run = shape.paragraphs?.[0]?.runs[0];
    // dir="auto": Hebrew/Arabic text is typed right to left
    const ta = h("textarea", { class: "pptxjs-editor-text", "aria-label": "Shape text", spellcheck: "true", dir: "auto" });
    ta.value = shape.text.replace(/\u000b/g, "\n");
    Object.assign(ta.style, {
      left: shape.bounds.x * s + "px",
      top: shape.bounds.y * s + "px",
      width: Math.max(80, shape.bounds.width * s) + "px",
      height: Math.max(40, shape.bounds.height * s) + "px",
      fontSize: Math.max(10, (run?.size ?? 18) * (96 / 72) * s) + "px",
      fontWeight: run?.bold ? "bold" : "normal",
      fontStyle: run?.italic ? "italic" : "normal",
    });
    ta.dataset.slide = String(ref.slide);
    ta.dataset.id = ref.id;
    ta.dataset.original = ta.value;
    ta.addEventListener("keydown", (e) => {
      e.stopPropagation();
      if (e.key === "Escape") {
        ta.value = ta.dataset.original ?? ta.value;
        this.commitText();
      } else if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
        this.commitText();
      }
    });
    ta.addEventListener("blur", () => this.commitText());
    wrap.appendChild(ta);
    this.overlay.hidden = true;
    this.textEditor = ta;
    ta.focus();
    ta.select();
  }

  /** Writes the inline text editor back to the document. */
  private commitText(): void {
    const ta = this.textEditor;
    if (!ta) return;
    this.textEditor = null;
    const value = ta.value;
    const original = ta.dataset.original;
    const slide = Number(ta.dataset.slide);
    const id = ta.dataset.id!;
    ta.remove();
    if (value !== original) this.edit((doc) => doc.setShapeText(slide, id, value));
    this.positionOverlay();
    this.root.focus({ preventScroll: true });
  }

  private onKey(ev: KeyboardEvent): void {
    const t = ev.target as HTMLElement;
    if (t.tagName === "INPUT" || t.tagName === "SELECT" || t.tagName === "TEXTAREA") return;
    const mod = ev.ctrlKey || ev.metaKey;
    const ref = this.selection;
    const nudge = ev.shiftKey ? 10 : 1;
    const rtl = this.root.getAttribute("dir") === "rtl";
    let handled = true;
    if (mod && ev.key.toLowerCase() === "z") {
      if (ev.shiftKey) this.redo();
      else this.undo();
    }
    else if (mod && ev.key.toLowerCase() === "y") this.redo();
    else if (mod && ev.key.toLowerCase() === "d" && ref) this.edit((doc) => this.select({ slide: ref.slide, id: doc.duplicateShape(ref.slide, ref.id) }));
    else if (mod && ev.key.toLowerCase() === "b" && ref) this.edit((doc) => doc.setTextStyle(ref.slide, ref.id, { bold: !this.firstRun()?.bold }));
    else if (mod && ev.key.toLowerCase() === "i" && ref) this.edit((doc) => doc.setTextStyle(ref.slide, ref.id, { italic: !this.firstRun()?.italic }));
    else if (mod && ev.key.toLowerCase() === "u" && ref) this.edit((doc) => doc.setTextStyle(ref.slide, ref.id, { underline: !this.firstRun()?.underline }));
    else if ((ev.key === "Delete" || ev.key === "Backspace") && ref) this.deleteSelection();
    else if (ev.key === "Escape") this.select(null);
    else if (ev.key === "Enter" && ref && this.doc) {
      try {
        const shape = this.doc.getShape(ref.slide, ref.id);
        if (shape.paragraphs !== null) this.openTextEditor(ref, shape);
      } catch {
        /* ignore */
      }
    } else if (ref && ev.key.startsWith("Arrow") && this.doc) {
      const b = this.doc.getShape(ref.slide, ref.id).bounds;
      if (b) {
        const dx = ev.key === "ArrowLeft" ? -nudge : ev.key === "ArrowRight" ? nudge : 0;
        const dy = ev.key === "ArrowUp" ? -nudge : ev.key === "ArrowDown" ? nudge : 0;
        this.edit((doc) => doc.setShapeBounds(ref.slide, ref.id, { x: b.x + dx, y: b.y + dy }));
      }
    } else if (!ref && (ev.key === "PageDown" || ev.key === "ArrowDown" || ev.key === (rtl ? "ArrowLeft" : "ArrowRight"))) this.goToSlide(this.viewer.currentIndex + 1);
    else if (!ref && (ev.key === "PageUp" || ev.key === "ArrowUp" || ev.key === (rtl ? "ArrowRight" : "ArrowLeft"))) this.goToSlide(this.viewer.currentIndex - 1);
    else handled = false;
    if (handled) ev.preventDefault();
  }

  private async onImagePicked(): Promise<void> {
    const file = this.fileInput.files?.[0];
    this.fileInput.value = "";
    if (!file) return;
    const bytes = new Uint8Array(await file.arrayBuffer());
    const ext = file.name.split(".").pop() ?? "";
    this.edit((doc, s) => {
      const id = doc.addImage(s, bytes, file.type || ext);
      this.select({ slide: s, id });
    });
  }
}
