# Changelog

## 2.0.0-alpha.0 (unreleased)

Complete rewrite as a modern, framework-agnostic TypeScript library. See
[docs/migration-v1-to-v2.md](docs/migration-v1-to-v2.md).

### Added

- `PptxViewer`: list and slideshow modes, fit/contain/fixed scaling, keyboard navigation, fullscreen,
  autoplay (loop, random), fade/slide/zoom transitions, progress and slide events, incremental rendering,
  several viewers per page, `destroy()`.
- `PptxEditor`: select, move, resize, inline text editing, formatting, fill, add text box, shape or picture,
  duplicate, delete, arrange, slide add/duplicate/move/hide/delete, undo/redo, thumbnails, keyboard shortcuts.
- `PresentationDocument`: lossless OOXML editing model with transactions and undo/redo; works in Node.js.
- Export to `.pptx` (lossless round trip) and `.odp` (OpenDocument Presentation).
- Headless rendering for Node.js, SSR and Web Workers: `renderToHtml`, `renderToHtmlString`, `renderPresentation`.
- Framework bindings: React, Vue 3, Angular (custom elements), jQuery (v1-compatible `pptxToHtml`), Web Components.
- Dependency-free SVG charts (bar/column clustered, stacked and 100%, line, area, pie, doughnut, scatter, combos) with theme colours.
- Right-to-left support: `textDirection` option (`file`/`auto`/`rtl`), `dir` option for an RTL viewer/editor UI (automatic on `dir="rtl"` pages), RTL paragraphs in the editor model and ODP export, RTL language detection by BCP-47 subtag.
- Speaker notes, hidden slides (`skipHiddenSlides`), `slides` filter, links between slides.
- `PptxError` with error codes; per-slide error isolation; `warning` events.
- TypeScript declarations, ESM, CJS and UMD builds, `style.css`.
- Test suite (unit, integration, security, DOM, adapters), demo app, documentation.

### Changed

- No jQuery, JSZip, FileReader.js, d3, nvd3 or tinycolor. The only dependency is `fflate`.
- Slide order follows `presentation.xml` (#15).
- Text wraps at word boundaries (`white-space: pre-wrap` instead of `&nbsp;`).
- Linked videos are not embedded unless `externalMedia: true` (then sandboxed).
- Keyboard shortcuts are active only while the viewer has focus (`keyboard: "global"` to opt in).
- CSS is scoped to the viewer; slides are isolated from host-page CSS resets.
- Internet Explorer is no longer supported.

### Fixed

- Crashes: #44, #43, #42, #41, #39, #33, #32, #31, #26, #21.
- Prototype pollution breaking other libraries: #30, #28, #22, #17 (and likely #29).
- Rendering: #40 (Symbol/Wingdings bullets), #38 (slide `showMasterSp`), #20 (`<sapn>`), #18 (linked video URLs),
  #12 (custom geometry, bullets, wrapping), numbering `startAt`, unclosed tables, group debug border.
- RTL: text inside RTL paragraphs was reset to LTR (`direction: initial`), misplacing punctuation and numbers;
  every non-RTL-language run was forced to `direction: ltr`; only 6 locales were recognised as RTL;
  `<html dir="rtl">` pages mirrored slide layouts.
- Error handling: #14.
- Security: XSS through document content and `javascript:` links (#3 follow-up), reverse tabnabbing,
  unsandboxed iframes, zip bombs, page freeze from a busy-wait, object-URL leaks. See [docs/security.md](docs/security.md).

### Removed

- Options `jsZipV2`, `slideType: "revealjs"`, `revealjsPath`, `revealjsConfig`, `navTxtColor`, `showPlayPauseBtn`.
- Runtime script loading.

## 1.21.1 and earlier

See the [v1 README](legacy/README-v1.md) and the [GitHub releases](https://github.com/meshesha/PPTXjs/releases).
