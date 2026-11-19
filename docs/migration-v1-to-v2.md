# Migrating from PPTXjs v1 to v2

v2 is a rewrite in TypeScript without jQuery. Most v1 code keeps working through the jQuery adapter, but you
can remove five script tags and the jQuery dependency.

## The shortest path (no code changes)

Replace the v1 scripts and stylesheets:

```html
<!-- v1: remove all of these -->
<link rel="stylesheet" href="./css/pptxjs.css">
<link rel="stylesheet" href="./css/nv.d3.min.css">
<script src="./js/jszip.min.js"></script>
<script src="./js/filereader.js"></script>
<script src="./js/d3.min.js"></script>
<script src="./js/nv.d3.min.js"></script>
<script src="./js/dingbat.js"></script>
<script src="./js/pptxjs.js"></script>
<script src="./js/divs2slides.js"></script>

<!-- v2: keep jQuery (only if the rest of your page uses it), add one bundle -->
<script src="jquery.min.js"></script>
<script src="pptxjs.umd.js"></script>
```

`$("#result").pptxToHtml({...})` keeps working. The options are translated:

| v1 option | v2 |
|---|---|
| `pptxFileUrl` | `viewer.load(url)` |
| `fileInputId` | listens to the input and calls `load(file)` |
| `slideMode: true` | `mode: "slideshow"` |
| `slidesScale: "50"` | `scale: 0.5` (default is now `"fit"`) |
| `keyBoardShortCut` | `keyboard: "global"` (otherwise keys work when the viewer has focus) |
| `mediaProcess`, `themeProcess`, `incSlide` | unchanged |
| `slideModeConfig.first / nav / loop` | `slideshow.first / nav / loop` |
| `slideModeConfig.showSlideNum` | `slideshow.showSlideNumber` |
| `slideModeConfig.autoSlide` | `slideshow.autoplay` (seconds) |
| `slideModeConfig.randomAutoSlide` | `slideshow.random` |
| `slideModeConfig.transition: "slid" \| "fade" \| "default" \| "random"` | `"slide" \| "fade" \| "none" \| "random"` |
| `slideModeConfig.transitionTime` | `slideshow.transitionDuration` |
| `slideModeConfig.background` | `slideshow.background` |
| `jsZipV2` | not needed; v2 has no JSZip dependency and no page reload |
| `slideType: "revealjs"`, `revealjsPath`, `revealjsConfig` | removed; use `renderToHtml()` and wrap the slides in Reveal.js sections yourself |
| `slideModeConfig.navTxtColor`, `showPlayPauseBtn` | removed; style `.pptxjs-toolbar` with CSS |

## Recommended: use the v2 API

```js
// v1
$("#result").pptxToHtml({ pptxFileUrl: "deck.pptx", slideMode: true, slideModeConfig: { autoSlide: 5 } });

// v2
import { PptxViewer } from "@meshesha/pptxjs";
const viewer = new PptxViewer("#result", { mode: "slideshow", slideshow: { autoplay: 5 } });
await viewer.load("deck.pptx");
```

Benefits: a returned `Promise` and `error` events (v1 errors could not be caught, issue #14), `destroy()`,
several viewers on one page, typed options and events.

## Behaviour changes

| Area | v1 | v2 |
|---|---|---|
| Errors | Uncaught exceptions; one bad slide stopped everything | `load()` rejects with `PptxError`; a slide that fails is replaced by a placeholder and reported via `warning` |
| Slide order | Sorted by file name | Order from `presentation.xml` (as in PowerPoint) |
| Charts | d3 + nvd3, drawn after rendering | Inline SVG drawn during rendering, with theme colours |
| Text | Spaces replaced with `&nbsp;`, so text could only wrap mid-word | Real spaces with `white-space: pre-wrap`; wraps at words like PowerPoint |
| Markup | `<sapn>` | `<span>` (issue #20) |
| Groups | Debug red border | No border |
| Bullets | Symbol/Wingdings private-use characters (needed the fonts installed) | Mapped to Unicode (issue #40) |
| Numbering | `startAt` ignored | `startAt` honoured |
| RTL text | Text in RTL paragraphs had an LTR base direction (punctuation and numbers at the wrong end); RTL pages mirrored slides | Correct RTL base direction; slides isolated from page direction; `textDirection` and `dir` options |
| Linked videos | Unsandboxed `<iframe>` | Poster + link by default; `externalMedia: true` for a sandboxed iframe |
| Keyboard | Captured every key on the page (blocked typing elsewhere) | Only when the viewer has focus, never inside inputs |
| Globals | Patched `Object.prototype.set` (broke Vue, ExtJS, jQuery; issues #22/#28/#30) | No globals |
| CSS | Global `.slide` rules | Scoped to `.pptxjs` / the viewer instance |
| DOM ids | `#all_slides_warpper`, `#chartN`, `audio_player` (duplicated) | Unique per viewer (`<idPrefix>-slide-N`, `<idPrefix>-chart-N`) |
| Browsers | IE11 and newer | Evergreen browsers and Node.js 18+ |

## CSS selectors

If you styled v1 output:

- `.slide` still exists and keeps its inner structure (`div.block`, `.slide-prgrph`, `.text-block` …).
- `#all_slides_warpper` is now `.pptxjs-slides`; each slide is wrapped in `.pptxjs-slide-wrap`.
- `.slides-loadnig-msg` / `.slides-loading-progress-bar` are now `.pptxjs-progress` / `.pptxjs-progress-bar`.
- divs2slides' `.slides-toolbar` is now `.pptxjs-toolbar`.
