// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { PptxViewer } from "../../src/viewer/viewer";
import { PptxEditor } from "../../src/editor/editor";
import { PptxError } from "../../src/errors";
import { makePptx, sp } from "../helpers/make-pptx";

const deck = (n = 3) => makePptx(Array.from({ length: n }, (_, i) => ({ shapes: sp(2, `Slide ${i + 1}`) + sp(3, `Second ${i + 1}`, { x: 914400 * 4 }) })));

let host: HTMLElement;
beforeEach(() => {
  host = document.createElement("div");
  document.body.appendChild(host);
});
afterEach(() => {
  host.remove();
  document.body.innerHTML = "";
});

describe("PptxViewer", () => {
  it("renders all slides in list mode and emits events", async () => {
    const viewer = new PptxViewer(host, { scale: 1 });
    const progress: number[] = [];
    const loaded = vi.fn();
    viewer.on("progress", (p) => progress.push(p.done));
    viewer.on("load", loaded);
    const res = await viewer.load(deck(3));
    expect(res.slides).toHaveLength(3);
    expect(viewer.slideCount).toBe(3);
    expect(progress).toEqual([1, 2, 3]);
    expect(loaded).toHaveBeenCalledOnce();
    expect(host.querySelectorAll(".pptxjs-slide-wrap")).toHaveLength(3);
    expect(host.querySelector(".slide")!.textContent).toContain("Slide 1");
    // generated CSS is scoped to this viewer instance
    const css = host.querySelector("style")!.textContent!;
    expect(css).toContain(`[data-pptxjs-id="${viewer.id}"]`);
    viewer.destroy();
    expect(host.children).toHaveLength(0);
  });

  it("tags top-level shapes for hit testing", async () => {
    const viewer = new PptxViewer(host);
    await viewer.load(deck(1));
    const ids = Array.from(host.querySelectorAll("[data-shape-id]")).map((e) => e.getAttribute("data-shape-id"));
    expect(ids).toEqual(["2", "3"]);
    viewer.destroy();
  });

  it("navigates in slideshow mode with buttons and keyboard", async () => {
    const viewer = new PptxViewer(host, { mode: "slideshow", slideshow: { transition: "none" } });
    await viewer.load(deck(3));
    const visible = () => Array.from(host.querySelectorAll<HTMLElement>(".pptxjs-slide-wrap")).filter((w) => !w.hidden).length;
    expect(visible()).toBe(1);
    expect(host.querySelector(".pptxjs-counter")!.textContent).toBe("1 / 3");
    const changes: number[] = [];
    viewer.on("slidechange", ({ index }) => changes.push(index));
    host.querySelector<HTMLButtonElement>(".pptxjs-next")!.click();
    expect(viewer.currentIndex).toBe(1);
    viewer.element.dispatchEvent(new KeyboardEvent("keydown", { key: "End", bubbles: true }));
    expect(viewer.currentIndex).toBe(2);
    viewer.element.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }));
    expect(viewer.currentIndex).toBe(2); // no loop
    viewer.element.dispatchEvent(new KeyboardEvent("keydown", { key: "Home", bubbles: true }));
    expect(changes).toEqual([1, 2, 0]);
    viewer.setMode("list");
    expect(visible()).toBe(3);
    viewer.destroy();
  });

  it("does not hijack keys typed into inputs or when keyboard is off", async () => {
    const viewer = new PptxViewer(host, { mode: "slideshow", keyboard: "global" });
    await viewer.load(deck(2));
    const input = document.createElement("input");
    document.body.appendChild(input);
    const ev = new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true, cancelable: true });
    input.dispatchEvent(ev);
    expect(viewer.currentIndex).toBe(0);
    expect(ev.defaultPrevented).toBe(false);
    viewer.destroy();
  });

  it("loops and autoplays", async () => {
    vi.useFakeTimers();
    try {
      const viewer = new PptxViewer(host, { mode: "slideshow", incremental: false, slideshow: { loop: true, transition: "none" } });
      await viewer.load(deck(2));
      viewer.play(1);
      expect(viewer.playing).toBe(true);
      vi.advanceTimersByTime(1000);
      expect(viewer.currentIndex).toBe(1);
      vi.advanceTimersByTime(1000);
      expect(viewer.currentIndex).toBe(0);
      viewer.pause();
      expect(viewer.playing).toBe(false);
      viewer.destroy();
    } finally {
      vi.useRealTimers();
    }
  });

  it("reports invalid files through the error event and a rejected promise", async () => {
    const viewer = new PptxViewer(host);
    const onError = vi.fn();
    viewer.on("error", onError);
    await expect(viewer.load(new Uint8Array([1, 2, 3]))).rejects.toBeInstanceOf(PptxError);
    expect(onError).toHaveBeenCalledOnce();
    expect(host.querySelector(".pptxjs-error")!.textContent).toMatch(/not a valid/);
    viewer.destroy();
  });

  it("loads from URLs via fetch and reports HTTP errors", async () => {
    const bytes = deck(1);
    const fetchMock = vi.fn(async (url: string) =>
      url.endsWith("ok.pptx") ? new Response(bytes as BlobPart) : new Response("nope", { status: 404 }),
    );
    vi.stubGlobal("fetch", fetchMock);
    try {
      const viewer = new PptxViewer(host);
      await viewer.load("https://example.com/ok.pptx");
      expect(viewer.slideCount).toBe(1);
      await expect(viewer.load("https://example.com/missing.pptx")).rejects.toMatchObject({ code: "FETCH_FAILED" });
      viewer.destroy();
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it("accepts Blob/File sources", async () => {
    const viewer = new PptxViewer(host);
    await viewer.load(new Blob([deck(2) as BlobPart]));
    expect(viewer.slideCount).toBe(2);
    viewer.destroy();
  });

  it("re-renders a single slide", async () => {
    const viewer = new PptxViewer(host);
    await viewer.load(deck(2));
    const pkg = viewer.package!;
    pkg.setText("ppt/slides/slide2.xml", pkg.text("ppt/slides/slide2.xml")!.replace("Slide 2", "Changed"));
    const el = viewer.rerenderSlide(1)!;
    expect(el.textContent).toContain("Changed");
    viewer.destroy();
  });
});

describe("PptxEditor", () => {
  it("loads, selects, edits text, undoes and exports", async () => {
    const editor = new PptxEditor(host);
    await editor.load(deck(2));
    expect(host.querySelectorAll(".pptxjs-thumb")).toHaveLength(2);
    const selections: Array<string | undefined> = [];
    editor.on("selectionchange", ({ selection }) => selections.push(selection?.id));
    editor.select({ slide: 0, id: "2" });
    expect(selections).toEqual(["2"]);
    editor.document!.setShapeText(0, "2", "Edited");
    await new Promise((r) => setTimeout(r, 0));
    expect(editor.viewer.getSlideElement(0)!.textContent).toContain("Edited");
    editor.undo();
    await new Promise((r) => setTimeout(r, 0));
    expect(editor.viewer.getSlideElement(0)!.textContent).toContain("Slide 1");
    // keyboard delete removes the selected shape
    editor.select({ slide: 0, id: "3" });
    host.querySelector<HTMLElement>(".pptxjs-editor")!.dispatchEvent(new KeyboardEvent("keydown", { key: "Delete", bubbles: true }));
    await new Promise((r) => setTimeout(r, 0));
    expect(editor.document!.getShapes(0).map((s) => s.id)).toEqual(["2"]);
    const pptx = editor.exportPptx();
    expect(pptx.type).toBe("application/vnd.openxmlformats-officedocument.presentationml.presentation");
    expect(pptx.size).toBeGreaterThan(1000);
    const odp = editor.exportOdp();
    expect(odp.type).toBe("application/vnd.oasis.opendocument.presentation");
    editor.destroy();
  });

  it("structural edits re-render the deck and thumbnails", async () => {
    const editor = new PptxEditor(host);
    await editor.load(deck(2));
    editor.document!.addSlide();
    await new Promise((r) => setTimeout(r, 50));
    expect(host.querySelectorAll(".pptxjs-thumb")).toHaveLength(3);
    expect(editor.viewer.slideCount).toBe(3);
    editor.destroy();
  });

  it("read-only mode blocks edits from the UI", async () => {
    const editor = new PptxEditor(host, { readOnly: true });
    await editor.load(deck(1));
    editor.select({ slide: 0, id: "2" });
    host.querySelector<HTMLElement>(".pptxjs-editor")!.dispatchEvent(new KeyboardEvent("keydown", { key: "Delete", bubbles: true }));
    expect(editor.document!.getShapes(0)).toHaveLength(2);
    editor.destroy();
  });
});

describe("right-to-left UI", () => {
  it("follows the page direction, mirrors arrow keys and keeps slides LTR", async () => {
    host.setAttribute("dir", "rtl");
    const viewer = new PptxViewer(host, { mode: "slideshow", slideshow: { transition: "none" } });
    await viewer.load(deck(3));
    expect(viewer.isRtl).toBe(true);
    expect(viewer.element.getAttribute("dir")).toBe("rtl");
    viewer.element.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowLeft", bubbles: true }));
    expect(viewer.currentIndex).toBe(1); // left = forward in RTL
    viewer.element.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight", bubbles: true }));
    expect(viewer.currentIndex).toBe(0);
    // slide content is isolated from the host direction
    const css = host.querySelector("style")!.textContent!;
    expect(css).toMatch(/\.pptxjs \.slide \{ direction: ltr; unicode-bidi: isolate; \}/);
    viewer.destroy();
    host.removeAttribute("dir");
  });

  it("dir option overrides the page direction", async () => {
    host.setAttribute("dir", "rtl");
    const viewer = new PptxViewer(host, { dir: "ltr" });
    expect(viewer.isRtl).toBe(false);
    viewer.setOptions({ dir: "rtl" });
    expect(viewer.isRtl).toBe(true);
    viewer.destroy();
    host.removeAttribute("dir");
  });

  it("editor UI is RTL on RTL pages", async () => {
    host.setAttribute("dir", "rtl");
    const editor = new PptxEditor(host);
    await editor.load(deck(2));
    expect(host.querySelector(".pptxjs-editor")!.getAttribute("dir")).toBe("rtl");
    expect(editor.viewer.isRtl).toBe(true);
    editor.destroy();
    host.removeAttribute("dir");
  });
});
