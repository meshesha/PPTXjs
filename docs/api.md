# API reference

All exports are available from `@meshesha/pptxjs` (ESM and CJS) and from the global `PPTXjs` in the UMD bundle.
Framework bindings live in sub-paths (`/react`, `/vue`, `/angular`, `/jquery`); see [frameworks.md](frameworks.md).

- [Sources](#sources)
- [PptxViewer](#pptxviewer)
- [PptxEditor](#pptxeditor)
- [PresentationDocument](#presentationdocument)
- [Headless rendering](#headless-rendering)
- [Export](#export)
- [Charts](#charts)
- [PptxPackage](#pptxpackage)
- [Errors](#errors)
- [Styling](#styling)

---

## Sources

```ts
type PptxSource = string | URL | Blob | ArrayBuffer | ArrayBufferView;
```

A string or `URL` is fetched with `fetch()`; pass `fetchOptions` for headers or credentials. `File` is a `Blob`.
`javascript:` and `vbscript:` URLs are rejected.

---

## PptxViewer

```ts
new PptxViewer(container: HTMLElement | string, options?: ViewerOptions)
```

Creates the viewer DOM inside `container` (an element or a CSS selector).

### Methods

| Method | Description |
|---|---|
| `load(source): Promise<RenderedPresentation>` | Loads and renders a presentation. A newer `load()` cancels an older one (the older promise rejects with code `ABORTED`). |
| `renderPackage(pkg, { startIndex? })` | Renders an already opened `PptxPackage`. |
| `rerenderSlide(index): HTMLElement \| null` | Re-renders one slide from the package (after edits). |
| `goTo(index)`, `next()`, `prev()`, `first()`, `last()` | Navigation. In list mode the slide is scrolled into view. |
| `play(seconds?)`, `pause()` | Autoplay (slideshow mode). |
| `setMode("list" \| "slideshow")` | Switch modes. |
| `setOptions(partial)` | Change options at runtime (scale, slideshow settings, mode). |
| `enterFullscreen()`, `exitFullscreen()`, `toggleFullscreen()` | Fullscreen API on the viewer element. |
| `getSlideElement(index): HTMLElement \| null` | The rendered `.slide` element. |
| `on(event, handler): () => void` / `off(event, handler)` | Events (see below). `on` returns an unsubscribe function. |
| `destroy()` | Removes DOM, listeners, timers and object URLs. |

### Properties

`slideCount`, `currentIndex`, `currentMode`, `playing`, `scale`, `result` (the last `RenderedPresentation`),
`package` (the `PptxPackage`), `element` (root element), `id` (instance id used to scope CSS).

### Options (`ViewerOptions`)

| Option | Type | Default | |
|---|---|---|---|
| `mode` | `"list" \| "slideshow"` | `"list"` | |
| `scale` | `"fit" \| "contain" \| number` | `"fit"` | `fit` uses the container width; `contain` uses width and height. |
| `slideshow` | `SlideshowOptions` | | See below. |
| `keyboard` | `boolean \| "global"` | `true` | `true`: only while the viewer has focus. Keys typed into inputs are never captured. |
| `showProgress` | `boolean` | `true` | Progress bar while loading. |
| `incremental` | `boolean` | `true` | Yield to the browser between slides so large decks keep the page responsive. |
| `gap` | `number` | `24` | Space between slides in list mode (px). |
| `limits` | `PackageLimits` | | `maxUncompressedSize` (512 MiB), `maxEntries` (20 000). |
| `fetchOptions` | `RequestInit` | | Used when loading from a URL. |
| `textDirection` | `"file" \| "auto" \| "rtl"` | `"file"` | Paragraphs without an explicit direction: as PowerPoint (LTR), detected from the first strong character, or RTL. See [rtl.md](rtl.md). |
| `dir` | `"auto" \| "ltr" \| "rtl"` | `"auto"` | Direction of the viewer UI (toolbar, arrow keys, transition). `auto` follows the nearest `dir` attribute / CSS direction of the container. Slide content is not affected. `isRtl` reports the result. |
| `themeProcess` | `boolean \| "colorsAndImageOnly"` | `true` | |
| `mediaProcess` | `boolean` | `true` | Embedded video/audio. |
| `externalMedia` | `boolean` | `false` | Embed linked videos in a sandboxed iframe; otherwise poster + link. |
| `skipHiddenSlides` | `boolean` | `false` | |
| `slides` | `number[]` | | Only render these (1-based) slide numbers. |
| `incSlide` | `{ width?, height? }` | | Extra pixels added to the slide size (v1 option). |
| `idPrefix` | `string` | auto | Prefix for generated ids (`<prefix>-slide-3`, `<prefix>-chart-0`). |
| `measureHtmlWidth` | `(html) => number` | DOM based | Override text measurement. |
| `onProgress` | `(done, total) => void` | | |
| `onWarning` | `(w: PptxError) => void` | | |

`SlideshowOptions`: `first` (1), `nav` (true), `showSlideNumber` (true), `autoplay` (false or seconds),
`random` (false), `loop` (false), `transition` (`"fade"`, also `"none" | "slide" | "zoom" | "random"`),
`transitionDuration` (0.4 s), `background` (CSS colour).

### Events (`ViewerEvents`)

| Event | Payload |
|---|---|
| `loadstart` | `{ source }` |
| `progress` | `{ done, total, ratio }` |
| `load` | `{ presentation: RenderedPresentation }` |
| `error` | `{ error: PptxError }` |
| `warning` | `{ warning: PptxError }` |
| `slidechange` | `{ index, slide: RenderedSlide }` |
| `sliderender` | `{ index, element }` |
| `modechange` | `{ mode }` |
| `fullscreenchange` | `{ fullscreen }` |

### Keyboard

| Key | Action |
|---|---|
| <kbd>→</kbd> <kbd>↓</kbd> <kbd>PgDn</kbd> <kbd>Space</kbd> | next slide |
| <kbd>←</kbd> <kbd>↑</kbd> <kbd>PgUp</kbd> <kbd>Backspace</kbd> | previous slide |
| <kbd>Home</kbd> / <kbd>End</kbd> | first / last |
| <kbd>F</kbd> | fullscreen |
| <kbd>P</kbd> | play / pause (slideshow) |

Links between slides inside a presentation navigate within the viewer.

---

## PptxEditor

```ts
new PptxEditor(container: HTMLElement | string, options?: EditorOptions)
```

`EditorOptions` extends the viewer options (except `mode`, `keyboard` and `slideshow`) with:

| Option | Default | |
|---|---|---|
| `toolbar` | `true` | Built-in toolbar. |
| `thumbnails` | `true` | Slide panel. |
| `readOnly` | `false` | Selection only, no edits. |
| `fileName` | `"presentation"` | Name used by `download()`. |

| Method / property | Description |
|---|---|
| `load(source)` | Opens a presentation. |
| `open(doc: PresentationDocument)` | Opens an existing document model. |
| `document` | The `PresentationDocument` (all editing APIs). |
| `viewer` | The embedded `PptxViewer`. |
| `select({ slide, id } \| null)`, `selected` | Selection. |
| `goToSlide(index)`, `currentSlide` | Navigation. |
| `undo()`, `redo()` | History. |
| `exportPptx(): Blob`, `exportOdp(): Blob` | Export. |
| `download("pptx" \| "odp", fileName?)` | Triggers a download. |
| `on(event, handler)` | `load`, `error`, `selectionchange`, `slidechange`, `change`. |
| `destroy()` | Cleanup. |

Editor shortcuts: <kbd>Del</kbd> delete · arrows nudge 1 px (<kbd>Shift</kbd> 10 px) · <kbd>Ctrl</kbd>+<kbd>Z</kbd>/<kbd>Y</kbd> undo/redo ·
<kbd>Ctrl</kbd>+<kbd>D</kbd> duplicate · <kbd>Ctrl</kbd>+<kbd>B</kbd>/<kbd>I</kbd>/<kbd>U</kbd> formatting · <kbd>Enter</kbd> or double-click to edit text
(<kbd>Ctrl</kbd>+<kbd>Enter</kbd> or click outside to commit, <kbd>Esc</kbd> to cancel) · <kbd>Shift</kbd>+drag constrains movement or keeps the aspect ratio when resizing.

---

## PresentationDocument

Editable model working directly on the OOXML parts. Every method is one undoable step. Units are CSS pixels (96 dpi).
Shape ids are the `p:cNvPr/@id` values (strings).

```ts
const doc = await PresentationDocument.load(source);   // or new PresentationDocument(pkg)
```

### Slides

| Method | |
|---|---|
| `slides(): SlideRef[]`, `slideCount` | `{ index, file, hidden }` in presentation order |
| `slideSize(): { width, height }` | |
| `addSlide(afterIndex?, layoutOf?)` | New empty slide using the layout of slide `layoutOf` (default: a "blank" layout). Returns its index. |
| `duplicateSlide(index)` | Copy inserted after the original (speaker notes are not shared). Returns the new index. |
| `moveSlide(from, to)` | |
| `deleteSlide(index)` | Also removes the slide's notes page. |
| `setSlideHidden(index, hidden)` | |
| `getNotes(index): string` | Speaker notes as text. |

### Shapes

| Method | |
|---|---|
| `getShapes(slide): ShapeInfo[]` | Top-level shapes, back to front |
| `getShape(slide, id): ShapeInfo` | |
| `setShapeBounds(slide, id, { x?, y?, width?, height? })` | Placeholders inheriting position from the layout get their own `xfrm` |
| `setShapeRotation(slide, id, degrees)` | |
| `deleteShape(slide, id)` | |
| `duplicateShape(slide, id, dx = 16, dy = 16): string` | Returns the new id |
| `reorderShape(slide, id, "front" \| "back" \| "forward" \| "backward")` | |
| `setShapeText(slide, id, text: string \| TextParagraph[])` | Keeps existing paragraph/run formatting as a template |
| `setTextStyle(slide, id, { bold?, italic?, underline?, size?, color?, font?, align? })` | Applies to all runs; `null` removes a property |
| `setShapeFill(slide, id, "#rrggbb" \| null)` | `null` = no fill |
| `setShapeLine(slide, id, "#rrggbb" \| null, widthPx?)` | |
| `addTextBox(slide, bounds, text?, style?): string` | |
| `addShape(slide, geometry, bounds, { fill?, line?, text? }): string` | Any OOXML preset name (`"rect"`, `"ellipse"`, `"star5"`, `"rightArrow"` …) |
| `addImage(slide, bytes, mimeOrExt, bounds?): string` | PNG/JPEG/GIF/SVG/BMP/WebP/TIFF; default size is the natural size, centred and scaled to fit |

```ts
interface ShapeInfo {
  id: string; name: string;
  kind: "shape" | "textbox" | "picture" | "table" | "chart" | "diagram" | "group" | "connector" | "media" | "other";
  bounds: { x: number; y: number; width: number; height: number } | null;
  rotation: number;
  placeholder?: { type?: string; idx?: string };
  geometry?: string;         // preset name or "custom"
  fill?: string;             // direct solid fill
  paragraphs: TextParagraph[] | null;  // null: cannot hold text
  text: string;
}
interface TextParagraph { runs: TextRun[]; align?: "left" | "center" | "right" | "justify"; level?: number; rtl?: boolean /* auto-detected when omitted */ }
interface TextRun { text: string; bold?: boolean; italic?: boolean; underline?: boolean; size?: number /*pt*/; color?: string; font?: string }
```

### History and events

`undo()`, `redo()`, `canUndo`, `canRedo`, `undoLabel`, `redoLabel`, `undoLimit` (100).
`transact(change, fn)` groups several calls into one step; if `fn` throws, all its writes are rolled back.
`on("change" | "undo" | "redo", handler)`. `change` carries `{ type: "slide" | "structure", index?, label }`.

`toPptx(): Uint8Array` serialises the package.

---

## Headless rendering

These functions work in browsers, Web Workers and Node.js.

```ts
loadPptx(source, { limits?, fetchOptions?, signal? }): Promise<PptxPackage>
renderToHtml(source, options?): Promise<RenderedPresentation>
renderToHtmlString(source, options?): Promise<string>    // slides + CSS in one <div class="pptxjs">
renderPresentation(pkg, options?): RenderedPresentation  // synchronous
listSlides(pkg): SlideInfo[]
readNotes(pkg, slideFile): string
```

```ts
interface RenderedPresentation {
  width: number; height: number;     // px
  slides: RenderedSlide[];
  css: string;                       // generated classes; inject together with BASE_CSS
  thumbnail?: string;                // data: URL
  appVersion: number;
  renderTime: number;                // ms
}
interface RenderedSlide { number: number; file: string; hidden: boolean; html: string; notes: string; error?: Error }
```

A slide that cannot be rendered does not fail the whole presentation: it gets `error` and a placeholder, and
`onWarning` is called.

---

## Export

```ts
exportPptx(docOrPkg, { pruneUnusedMedia = true }): Blob
exportPptxBytes(docOrPkg, options): Uint8Array
exportOdp(docOrPkg, { includeMasterShapes = true, includeNotes = true, skipHiddenSlides = false, textDirection = "file" }): Blob
exportOdpBytes(docOrPkg, options): Uint8Array
PPTX_MIME, ODP_MIME
```

---

## Charts

```ts
parseChart(chartSpaceXml, resolveFill): ChartModel | null
renderChartSvg(model, { width, height, palette?, fontFamily? }): string
DEFAULT_PALETTE
```

---

## PptxPackage

Low-level access to the zip package.

```ts
PptxPackage.fromBytes(bytes, limits?)
pkg.list(): string[]; pkg.has(path); pkg.bytes(path); pkg.text(path); pkg.xml(path)
pkg.setBytes(path, data); pkg.setText(path, text); pkg.delete(path); pkg.clone(); pkg.toZip()
```

---

## Errors

All errors are `PptxError` instances with a `code`:

| Code | Meaning |
|---|---|
| `INVALID_FILE` | Not a readable `.pptx` |
| `LIMIT_EXCEEDED` | A zip limit was exceeded |
| `FETCH_FAILED` | Network error or HTTP status ≥ 400 |
| `RENDER_WARNING` | Non-fatal; reported through `onWarning` / the `warning` event, `slide` is set |
| `EXPORT_FAILED` | Export failed or invalid input to an edit method (e.g. a bad colour) |
| `ABORTED` | Superseded or destroyed |

---

## Styling

- `BASE_CSS`: slide styles (scoped to `.pptxjs`), also shipped as `@meshesha/pptxjs/style.css` with the viewer and editor chrome.
- Generated per-presentation classes are scoped to `[data-pptxjs-id="…"]`, so several viewers can share a page.
- Slides are isolated from common host CSS resets (`box-sizing: border-box`, `max-width: 100%` …).
- Viewer chrome classes: `.pptxjs-viewer`, `.pptxjs-toolbar`, `.pptxjs-btn`, `.pptxjs-counter`, `.pptxjs-progress`, `.pptxjs-slide-wrap`, `.pptxjs-error`, and the modes `.pptxjs-mode-list`, `.pptxjs-mode-slideshow`, `.pptxjs-fullscreen`.
- Each top-level shape is wrapped in `<div class="pptx-shape" data-shape-id="…" style="display:contents">`, which is useful for click handlers.
