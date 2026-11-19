# PPTXjs

[![MIT License](https://img.shields.io/badge/license-MIT-blue.svg?style=flat)](LICENSE)

**Render, present, edit and export PowerPoint (`.pptx`) files in the browser, with no server and no jQuery.**

PPTXjs v2 is a TypeScript rewrite of the original jQuery plugin. It renders slides to HTML/SVG, runs
slideshows, lets users edit presentations, and saves the result as `.pptx` or `.odp`.

```ts
import { PptxViewer } from "@meshesha/pptxjs";

const viewer = new PptxViewer("#slides", { mode: "slideshow" });
await viewer.load("deck.pptx"); // URL, File, Blob, ArrayBuffer or Uint8Array
```

| | |
|---|---|
| **Viewer** | scrolling slide list or slideshow · fit-to-width/contain scaling · keyboard, fullscreen, autoplay, transitions · progress events |
| **Editor** | select, move, resize, rotate · inline text editing and formatting · add text, shapes and pictures · slide management · undo/redo |
| **Export** | lossless `.pptx` round trip · `.odp` (OpenDocument) conversion · HTML strings for SSR |
| **Frameworks** | React · Vue 3 · Angular · jQuery (v1-compatible API) · Web Components · vanilla JS |
| **Runtime** | browsers, Web Workers and Node.js (the renderer has no DOM dependency) · one runtime dependency ([fflate](https://github.com/101arrowz/fflate)) |
| **Safety** | output is XSS-safe by construction · URL sanitising · zip-bomb limits · no prototype pollution |

> v2 is currently `2.0.0-alpha.0`. See [CHANGELOG.md](CHANGELOG.md) for what changed and
> [docs/migration-v1-to-v2.md](docs/migration-v1-to-v2.md) if you are upgrading.

## Contents

- [Installation](#installation)
- [Quick start](#quick-start) (vanilla, React, Vue, Angular, jQuery, Web Component, Node.js)
- [Viewer options](#viewer-options)
- [Editing and export](#editing-and-export)
- [What is supported](#what-is-supported)
- [Security](#security)
- [Browser support](#browser-support)
- [Documentation](#documentation)
- [Development](#development)

## Installation

```bash
npm install @meshesha/pptxjs
```

Or without a bundler:

```html
<script src="https://unpkg.com/@meshesha/pptxjs/dist/pptxjs.umd.js"></script>
<script>
  const viewer = new PPTXjs.PptxViewer("#slides");
  viewer.load("deck.pptx");
</script>
```

The styles are injected automatically. `@meshesha/pptxjs/style.css` is also provided if you prefer a
static stylesheet, for example to satisfy a strict Content Security Policy.

## Quick start

### Vanilla JS

```ts
import { PptxViewer } from "@meshesha/pptxjs";

const viewer = new PptxViewer(document.getElementById("slides")!, {
  mode: "list",          // or "slideshow"
  scale: "fit",          // "fit" | "contain" | number
});

viewer.on("load", ({ presentation }) => console.log(presentation.slides.length, "slides"));
viewer.on("error", ({ error }) => console.error(error.code, error.message));

await viewer.load("/decks/quarterly.pptx");

// file input
input.addEventListener("change", () => viewer.load(input.files![0]));
```

### React

```tsx
import { useRef } from "react";
import { PptxViewer, PptxEditor } from "@meshesha/pptxjs/react";
import type { PptxViewer as Viewer } from "@meshesha/pptxjs";

export function Deck({ url }: { url: string }) {
  const ref = useRef<Viewer>(null);
  return (
    <>
      <button onClick={() => ref.current?.next()}>Next</button>
      <PptxViewer ref={ref} src={url} options={{ mode: "slideshow" }} onLoad={(p) => console.log(p.slides.length)} />
    </>
  );
}
```

### Vue 3

```vue
<script setup lang="ts">
import { PptxViewer } from "@meshesha/pptxjs/vue";
</script>

<template>
  <PptxViewer src="/deck.pptx" :options="{ mode: 'slideshow' }" @load="(p) => console.log(p)" />
</template>
```

`app.use(PptxPlugin)` registers `<PptxViewer>` and `<PptxEditor>` globally.

### Angular

```ts
// main.ts
import { definePptxElements } from "@meshesha/pptxjs/angular";
definePptxElements();

// deck.component.ts
@Component({
  selector: "app-deck",
  standalone: true,
  schemas: [CUSTOM_ELEMENTS_SCHEMA],
  template: `<pptx-viewer [attr.src]="url" mode="slideshow" (pptx-load)="onLoad($event)"></pptx-viewer>`,
})
export class DeckComponent {
  url = "/assets/deck.pptx";
  onLoad(e: Event) {}
}
```

A standalone component that wraps the viewer class directly is available in
[examples/angular/pptx-viewer.component.ts](examples/angular/pptx-viewer.component.ts).

### jQuery (v1 compatible)

Existing v1 code keeps working:

```js
import $ from "jquery";
import { installJQueryPlugin } from "@meshesha/pptxjs/jquery";
installJQueryPlugin($); // automatic with the UMD <script> when window.jQuery exists

$("#result").pptxToHtml({
  pptxFileUrl: "deck.pptx",
  fileInputId: "upload",
  slideMode: true,
  slideModeConfig: { autoSlide: 5, loop: true, transition: "fade" },
});

const viewer = $("#result").pptxViewer(); // the underlying PptxViewer
```

### Web Component

```html
<script src="pptxjs.umd.js"></script>
<pptx-viewer src="deck.pptx" mode="slideshow" transition="slide" autoplay="5" loop></pptx-viewer>
<pptx-editor src="deck.pptx" style="height: 600px"></pptx-editor>
```

### Node.js, SSR and Web Workers

```ts
import { renderToHtml, renderToHtmlString } from "@meshesha/pptxjs";
import { readFile } from "node:fs/promises";

const { slides, css, width, height } = await renderToHtml(await readFile("deck.pptx"));
slides[0].html;   // HTML of the first slide
slides[0].notes;  // speaker notes as text

const page = await renderToHtmlString(await readFile("deck.pptx")); // one self-contained fragment
```

## Viewer options

| Option | Type | Default | Description |
|---|---|---|---|
| `mode` | `"list" \| "slideshow"` | `"list"` | All slides stacked, or one slide at a time |
| `scale` | `"fit" \| "contain" \| number` | `"fit"` | Fit the container width, fit width and height, or a fixed factor |
| `slideshow.first` | `number` | `1` | First slide shown |
| `slideshow.nav` | `boolean` | `true` | Show the navigation toolbar |
| `slideshow.showSlideNumber` | `boolean` | `true` | Show "3 / 12" in the toolbar |
| `slideshow.autoplay` | `number \| false` | `false` | Seconds between slides |
| `slideshow.loop` / `random` | `boolean` | `false` | Wrap around / random order while autoplaying |
| `slideshow.transition` | `"none" \| "fade" \| "slide" \| "zoom" \| "random"` | `"fade"` | Transition between slides |
| `slideshow.transitionDuration` | `number` | `0.4` | Seconds |
| `slideshow.background` | CSS colour | – | Stage colour in slideshow/fullscreen |
| `keyboard` | `boolean \| "global"` | `true` | Shortcuts when the viewer has focus, or page-wide |
| `textDirection` | `"file" \| "auto" \| "rtl"` | `"file"` | Direction of paragraphs that don't declare one: as in the file (PowerPoint), detected from the text (Hebrew/Arabic → RTL), or RTL |
| `dir` | `"auto" \| "ltr" \| "rtl"` | `"auto"` | Viewer/editor UI direction; `auto` follows the page (`<html dir="rtl">`) |
| `themeProcess` | `boolean \| "colorsAndImageOnly"` | `true` | How much of the slide master/layout to draw |
| `mediaProcess` | `boolean` | `true` | Render embedded video/audio |
| `externalMedia` | `boolean` | `false` | Embed linked videos (YouTube…) in a sandboxed iframe |
| `skipHiddenSlides` | `boolean` | `false` | Leave out slides hidden in PowerPoint |
| `slides` | `number[]` | all | Render only these slide numbers |
| `limits` | `{ maxUncompressedSize, maxEntries }` | 512 MiB / 20 000 | Zip-bomb protection |
| `fetchOptions` | `RequestInit` | – | Headers or credentials for URL loading |
| `onWarning` | `(w: PptxError) => void` | – | Non-fatal rendering problems |

Keyboard: <kbd>←</kbd> <kbd>→</kbd> <kbd>Space</kbd> <kbd>PgUp</kbd> <kbd>PgDn</kbd> to navigate, <kbd>Home</kbd>/<kbd>End</kbd>,
<kbd>F</kbd> for fullscreen, <kbd>P</kbd> to play or pause.

Events: `loadstart`, `progress`, `load`, `error`, `warning`, `slidechange`, `modechange`, `fullscreenchange`, `sliderender`.

The full reference is in [docs/api.md](docs/api.md).

## Editing and export

```ts
import { PptxEditor } from "@meshesha/pptxjs";

const editor = new PptxEditor("#editor");
await editor.load("deck.pptx");
editor.download("pptx");          // or "odp"
```

The editor has a toolbar, slide thumbnails, drag-to-move, resize handles, double-click text editing,
formatting (bold, italic, underline, size, colour, alignment, fill), adding text boxes, shapes and pictures,
arranging, and slide add/duplicate/move/hide/delete. It supports undo/redo and the usual shortcuts.

For programmatic editing without any UI (also in Node.js):

```ts
import { PresentationDocument, exportPptx, exportOdp } from "@meshesha/pptxjs";

const doc = await PresentationDocument.load(bytes);
const id = doc.addTextBox(0, { x: 80, y: 60, width: 400, height: 60 }, "Hello", { size: 32, bold: true });
doc.setShapeFill(0, id, "#FFE699");
doc.duplicateSlide(0);
doc.undo();
const pptx = exportPptx(doc);   // Blob, lossless: untouched parts are byte-identical
const odp = exportOdp(doc);     // Blob, OpenDocument Presentation
```

See [docs/editing-and-export.md](docs/editing-and-export.md).

## What is supported

- **Text:** fonts, sizes, colours, bold, italic, underline, strike, caps, highlight, hyperlinks (safe), bullets (characters, pictures, Symbol/Wingdings mapped to Unicode), auto-numbering (all schemes incl. Hebrew, `startAt`), alignment, indents
- **Right-to-left (Hebrew, Arabic, Persian, Urdu …):** RTL paragraphs, bullets and numbering; mixed RTL/LTR text with correct punctuation; optional auto-detection for files without direction info (`textDirection: "auto"`); RTL viewer and editor UI; Hebrew typed in the editor is saved as RTL; RTL kept in ODP export. See [docs/rtl.md](docs/rtl.md)
- **Shapes:** about 180 preset geometries, custom geometry (all path commands and shape guides), rotation, flip, solid/gradient/pattern/picture fills, outlines, arrows, shadows and glow (partial)
- **Pictures:** PNG, JPEG, GIF, SVG, BMP, WebP and TIFF (browser dependent); EMF/WMF are not supported by browsers
- **Tables:** custom and theme table styles, merged cells
- **Charts:** bar/column (clustered, stacked, 100%), line, area, pie, doughnut, scatter and combos, rendered as dependency-free SVG with theme colours
- **SmartArt:** drawn from the cached drawing part
- **Media:** embedded video and audio; linked videos with opt-in sandboxing
- **Slides:** backgrounds (colour, gradient, picture), layout and master inheritance, hidden slides, slide order from `presentation.xml`, speaker notes

## Security

Presentations are untrusted input. PPTXjs v2:

- escapes all document text and attribute values in the XML parser, so the generated HTML cannot contain markup that came from the file
- allows only `http(s)`, `mailto`, `tel`, `ftp` and relative links (`javascript:`, `data:` and `vbscript:` are removed) and opens external links with `rel="noopener noreferrer"`
- never embeds third-party pages unless you opt in with `externalMedia`, and then sandboxes them
- limits uncompressed size and entry count (zip bombs)
- never writes to `Object.prototype` and ignores `__proto__`/`constructor` keys
- adds no global event handlers: keyboard handling is scoped to the viewer by default

Details are in [docs/security.md](docs/security.md). To report a vulnerability, see [SECURITY.md](SECURITY.md).

## Browser support

The latest two versions of Chrome, Edge, Firefox and Safari, plus Node.js 18 or later. Internet Explorer is no
longer supported; use PPTXjs v1 (in `legacy/`) if you need IE11.

## Documentation

| | |
|---|---|
| [docs/api.md](docs/api.md) | Complete API reference |
| [docs/frameworks.md](docs/frameworks.md) | React, Vue, Angular, jQuery, Web Components and Node.js in depth |
| [docs/editing-and-export.md](docs/editing-and-export.md) | Editor UI, `PresentationDocument`, PPTX/ODP export and limitations |
| [docs/rtl.md](docs/rtl.md) | Right-to-left languages (Hebrew, Arabic …) |
| [docs/migration-v1-to-v2.md](docs/migration-v1-to-v2.md) | Upgrading from the jQuery plugin |
| [docs/security.md](docs/security.md) | Threat model and mitigations |
| [docs/developer-guide.md](docs/developer-guide.md) | Architecture, code layout, testing and contributing |
| [docs/github-issues-report.md](docs/github-issues-report.md) | Status of every GitHub issue and the open feature requests |

## Development

```bash
npm install
npm run dev          # demo app on http://localhost:5175
npm test             # unit, integration, security and DOM tests (Vitest)
npm run lint
npm run typecheck
npm run build        # dist/: ESM, CJS, UMD, CSS, type declarations
```

## License

MIT © 2017–present Meshesha. See [LICENSE](LICENSE).
