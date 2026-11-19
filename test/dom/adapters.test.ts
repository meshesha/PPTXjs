// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from "vitest";
import { createElement, createRef } from "react";
import { createRoot } from "react-dom/client";
import { act } from "react";
import { createApp, h, nextTick } from "vue";
import jQuery from "jquery";
import { PptxViewer as ReactViewer } from "../../src/adapters/react";
import { PptxViewer as VueViewer } from "../../src/adapters/vue";
import { installJQueryPlugin, fromLegacyOptions } from "../../src/adapters/jquery";
import { defineCustomElements } from "../../src/adapters/element";
import type { PptxViewer } from "../../src/viewer/viewer";
import { makePptx, sp } from "../helpers/make-pptx";

(globalThis as any).IS_REACT_ACT_ENVIRONMENT = true;
const deck = makePptx([{ shapes: sp(2, "Adapter slide") }, { shapes: sp(2, "Second") }]);
const until = async (cond: () => boolean, ms = 3000) => {
  const t0 = Date.now();
  while (!cond()) {
    if (Date.now() - t0 > ms) throw new Error("timeout");
    await new Promise((r) => setTimeout(r, 10));
  }
};

afterEach(() => {
  document.body.innerHTML = "";
});

describe("React adapter", () => {
  it("renders, exposes the viewer through a ref and cleans up", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const ref = createRef<PptxViewer | null>();
    const onLoad = vi.fn();
    const root = createRoot(host);
    await act(async () => {
      root.render(createElement(ReactViewer, { ref, src: deck, onLoad }));
    });
    await until(() => onLoad.mock.calls.length > 0);
    expect(onLoad.mock.calls[0][0].slides).toHaveLength(2);
    expect(ref.current?.slideCount).toBe(2);
    expect(host.textContent).toContain("Adapter");
    await act(async () => root.unmount());
    expect(host.querySelector(".pptxjs-viewer")).toBeNull();
  });
});

describe("Vue adapter", () => {
  it("renders and emits load", async () => {
    const host = document.createElement("div");
    document.body.appendChild(host);
    const onLoad = vi.fn();
    const app = createApp({ render: () => h(VueViewer, { src: deck, onLoad }) });
    app.mount(host);
    await nextTick();
    await until(() => onLoad.mock.calls.length > 0);
    expect(host.querySelectorAll(".pptxjs-slide-wrap")).toHaveLength(2);
    app.unmount();
    expect(host.querySelector(".pptxjs-viewer")).toBeNull();
  });
});

describe("jQuery adapter", () => {
  it("keeps the v1 $.fn.pptxToHtml API working", async () => {
    installJQueryPlugin(jQuery);
    const url = URL.createObjectURL(new Blob([deck as BlobPart]));
    vi.stubGlobal("fetch", async () => new Response(deck as BlobPart));
    try {
      const $el = jQuery("<div id='result'></div>").appendTo(document.body);
      $el.pptxToHtml({ pptxFileUrl: url, slideMode: true, slideModeConfig: { transition: "slid", autoSlide: false } });
      const viewer = $el.pptxViewer() as PptxViewer;
      await until(() => viewer.slideCount === 2);
      expect(viewer.currentMode).toBe("slideshow");
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it("maps legacy options", () => {
    const o = fromLegacyOptions({ slideMode: true, slidesScale: "50", keyBoardShortCut: true, slideModeConfig: { autoSlide: 3, transition: "slid", loop: true } });
    expect(o).toMatchObject({ mode: "slideshow", scale: 0.5, keyboard: "global", slideshow: { autoplay: 3, transition: "slide", loop: true } });
  });
});

describe("Web Components", () => {
  it("<pptx-viewer> renders and dispatches pptx-load", async () => {
    defineCustomElements();
    const el = document.createElement("pptx-viewer") as HTMLElement & { source: unknown; viewer: PptxViewer };
    const loaded = vi.fn();
    el.addEventListener("pptx-load", loaded);
    el.setAttribute("mode", "slideshow");
    el.source = deck;
    document.body.appendChild(el);
    await until(() => loaded.mock.calls.length > 0);
    expect(el.viewer.currentMode).toBe("slideshow");
    el.remove();
    expect(el.viewer).toBeNull();
  });
});
