# Framework integration

PPTXjs is framework-agnostic: `PptxViewer` and `PptxEditor` are plain classes that take a container
element. The bindings below are thin wrappers that handle mounting, reloading on `src` changes and cleanup.
React, Vue, jQuery and Angular are optional peer dependencies; install only what you use.

- [React](#react)
- [Vue 3](#vue-3)
- [Angular](#angular)
- [jQuery](#jquery)
- [Web Components](#web-components)
- [Svelte, Solid, Lit and others](#svelte-solid-lit-and-others)
- [Node.js and SSR](#nodejs-and-ssr)
- [Web Workers](#web-workers)

## React

```tsx
import { PptxViewer, PptxEditor } from "@meshesha/pptxjs/react";
```

| Prop | Type | |
|---|---|---|
| `src` | `PptxSource \| null` | Changing it reloads |
| `options` | `ViewerOptions` / `EditorOptions` | Read on mount. Use the ref + `setOptions()` to change them later |
| `className`, `style` | | Applied to the host `<div>` |
| `onLoad`, `onError`, `onSlideChange` | | Viewer callbacks |
| `onLoad`, `onError`, `onChange`, `onSelectionChange` | | Editor callbacks |

The ref exposes the underlying instance:

```tsx
const viewer = useRef<import("@meshesha/pptxjs").PptxViewer>(null);
<PptxViewer ref={viewer} src={file} />
<button onClick={() => viewer.current?.toggleFullscreen()}>Fullscreen</button>
```

File upload:

```tsx
const [file, setFile] = useState<File | null>(null);
<input type="file" accept=".pptx" onChange={(e) => setFile(e.target.files?.[0] ?? null)} />
<PptxViewer src={file} />
```

Next.js: the components use the DOM, so render them on the client only (`"use client"`, or `dynamic(() => …, { ssr: false })`).
On the server, use `renderToHtml()`.

## Vue 3

```vue
<script setup lang="ts">
import { ref } from "vue";
import { PptxViewer, PptxEditor } from "@meshesha/pptxjs/vue";

const viewer = ref();   // viewer.value.viewer -> PptxViewer instance
</script>

<template>
  <PptxViewer ref="viewer" src="/deck.pptx" :options="{ mode: 'slideshow' }"
              @load="(p) => console.log(p.slides.length)" @error="console.error" @slidechange="(i) => …" />
  <PptxEditor src="/deck.pptx" @change="(label) => …" />
</template>
```

Global registration: `app.use(PptxPlugin)`. For Nuxt, wrap the component in `<ClientOnly>`.

## Angular

Angular libraries have to be compiled with the Angular compiler, so PPTXjs exposes standard custom elements,
which Angular supports natively:

```ts
// main.ts
import { definePptxElements } from "@meshesha/pptxjs/angular";
definePptxElements();
```

```ts
import { Component, CUSTOM_ELEMENTS_SCHEMA, ElementRef, ViewChild } from "@angular/core";
import type { PptxViewerElement } from "@meshesha/pptxjs/angular";

@Component({
  selector: "app-deck",
  standalone: true,
  schemas: [CUSTOM_ELEMENTS_SCHEMA],
  template: `
    <pptx-viewer #viewer [attr.src]="url" mode="slideshow" transition="fade"
                 (pptx-load)="loaded($event)" (pptx-slidechange)="changed($event)"></pptx-viewer>
    <button (click)="viewer.viewer?.next()">Next</button>`,
})
export class DeckComponent {
  url = "/assets/deck.pptx";
  @ViewChild("viewer") viewerRef!: ElementRef<PptxViewerElement>;
  loaded(e: Event) { console.log((e as CustomEvent).detail.slides.length); }
  changed(e: Event) { console.log((e as CustomEvent).detail.index); }
}
```

If you prefer a regular Angular component with `@Input()`/`@Output()`, copy
[examples/angular/pptx-viewer.component.ts](../examples/angular/pptx-viewer.component.ts) into your project.

## jQuery

```js
import $ from "jquery";
import { installJQueryPlugin } from "@meshesha/pptxjs/jquery";
installJQueryPlugin($);
```

With the UMD bundle (`<script src="pptxjs.umd.js">` after jQuery) the plugin installs itself.

```js
// v1 API (options are translated, see migration guide)
$("#result").pptxToHtml({ pptxFileUrl: "deck.pptx", fileInputId: "upload", slideMode: true });

// v2 style
$("#result").pptxViewer({ src: "deck.pptx", mode: "slideshow" });
const viewer = $("#result").pptxViewer();   // instance

$("#edit").pptxEditor({ src: "deck.pptx" });
```

TypeScript users get typings for `$.fn.pptxToHtml`, `$.fn.pptxViewer` and `$.fn.pptxEditor`.

## Web Components

```js
import { defineCustomElements } from "@meshesha/pptxjs";
defineCustomElements();      // the UMD bundle does this automatically
```

```html
<pptx-viewer src="deck.pptx" mode="slideshow" scale="fit" transition="slide" autoplay="5" loop></pptx-viewer>
<pptx-editor src="deck.pptx" style="display:block;height:600px"></pptx-editor>
```

| Attribute | |
|---|---|
| `src` | URL (reloads when changed) |
| `mode` | `list` / `slideshow` |
| `scale` | `fit` / `contain` / number |
| `transition`, `autoplay`, `loop` | Slideshow |

Properties: `options` (any `ViewerOptions`, set before connecting), `source` (accepts `File`/`Blob`/bytes),
`viewer` / `editor` (the instance). Events: `pptx-load`, `pptx-error`, `pptx-slidechange`, and `pptx-change` for the editor.
`defineCustomElements("my")` registers `<my-viewer>` / `<my-editor>` instead.

## Svelte, Solid, Lit and others

Use the custom elements, or the class directly:

```svelte
<script>
  import { onMount, onDestroy } from "svelte";
  import { PptxViewer } from "@meshesha/pptxjs";
  export let src;
  let el, viewer;
  onMount(() => { viewer = new PptxViewer(el); viewer.load(src); });
  onDestroy(() => viewer?.destroy());
  $: viewer && src && viewer.load(src);
</script>
<div bind:this={el}></div>
```

## Node.js and SSR

The parser, renderer, document model and exporters need no DOM:

```js
import { renderToHtml, PresentationDocument, exportOdpBytes, BASE_CSS } from "@meshesha/pptxjs";
import { readFile, writeFile } from "node:fs/promises";

const res = await renderToHtml(await readFile("deck.pptx"));
await writeFile("deck.html", `<!doctype html><style>${BASE_CSS}${res.css}</style><div class="pptxjs">${res.slides.map((s) => s.html).join("")}</div>`);

const doc = await PresentationDocument.load(await readFile("deck.pptx"));
await writeFile("deck.odp", exportOdpBytes(doc));
```

Without a DOM, text is not measured, so paragraphs with bullets use their full width instead of shrinking
to the text. Everything else is identical. See [examples/node/convert.mjs](../examples/node/convert.mjs).

## Web Workers

To keep very large decks off the main thread, render in a worker and insert the HTML:

```js
// worker.js
import { renderToHtml } from "@meshesha/pptxjs";
self.onmessage = async (e) => self.postMessage(await renderToHtml(e.data));

// main thread
worker.postMessage(fileBytes, [fileBytes.buffer]);
worker.onmessage = ({ data }) => { /* data.slides[i].html, data.css */ };
```

Video and audio use object URLs created in the worker. Create them on the main thread instead
(`mediaProcess: false` in the worker), or use `PptxViewer`, which renders incrementally and stays responsive.
