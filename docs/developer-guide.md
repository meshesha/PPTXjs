# Developer guide

This guide is for people working on PPTXjs itself: architecture, code layout, conventions, tests and releases.

## Architecture

```
                 ┌──────────────┐   bytes    ┌──────────────┐  simplified XML  ┌────────────────────┐
  PptxSource ──▶ │  source.ts   │──────────▶ │ PptxPackage  │────────────────▶ │  render/*  (HTML)  │
 (URL/File/…)    │ fetch / Blob │            │ fflate unzip │  (core/xml.ts)   │  DOM-free strings  │
                 └──────────────┘            │ limits,cache │                  └─────────┬──────────┘
                                             └──────┬───────┘                            │
                                                    │ lossless XML (core/xmldom.ts)      ▼
                                             ┌──────┴────────────┐              ┌────────────────────┐
                                             │ PresentationDoc.  │──re-render──▶│ PptxViewer (DOM)   │
                                             │ edits + undo      │              │ list / slideshow   │
                                             └──────┬────────────┘              └─────────┬──────────┘
                                                    │                                     │
                                       ┌────────────┴───────────┐                ┌────────┴─────────┐
                                       │ export/pptx  export/odp│                │ PptxEditor (UI)  │
                                       └────────────────────────┘                └──────────────────┘
                                   adapters/: React · Vue · Angular (custom elements) · jQuery · Web Components
```

Principles:

1. **The renderer is pure.** `render/*` turns parsed XML into HTML strings, with no DOM and no globals.
   It runs in Node and workers, and is easy to test.
2. **One context per render.** Presentation-wide state (generated CSS classes, chart counter, theme
   options, warnings, object URLs) lives in `RenderContext`. Per-slide data (layout, master, theme,
   relationships) lives in the `WarpObj` built by `buildWarpObj()`. Nothing is module-global except the
   XML document-order counter.
3. **Safe by construction.** The XML parser emits HTML-safe strings, so the large ported renderer can keep
   concatenating strings without auditing every call site. URLs go through `sanitizeUrl()`.
4. **Edits are made to the model.** The editor never edits the HTML. It changes OOXML through
   `PresentationDocument` and re-renders.

## Source layout

| Path | Lines | Purpose |
|---|---|---|
| `src/index.ts` | | Public API |
| `src/source.ts` | | `PptxSource` → bytes (fetch, Blob, buffers) |
| `src/errors.ts` | | `PptxError` and codes |
| `src/core/package.ts` | 280 | Zip package: limits, caching, mutation hooks, `toZip()` |
| `src/core/xml.ts` | 320 | Simplified-tree XML parser (tXml port), entity safety, AlternateContent resolution, tree helpers |
| `src/core/xmldom.ts` | 310 | Lossless XML DOM for editing |
| `src/render/presentation.ts` | 340 | Slide list, relationships, warpObj, per-slide error isolation, notes, CSS |
| `src/render/context.ts` | | `RenderContext`, `RenderOptions`, `WarpObj` |
| `src/render/nodes.ts` | 570 | Shape tree dispatch: shapes, pictures, media, groups, graphic frames, diagrams |
| `src/render/shape.ts` | 7 000 | `genShape`: 180+ preset geometries (ported from v1) |
| `src/render/text.ts` | 2 000 | Text bodies, paragraphs, runs, bullets, fonts, alignment |
| `src/render/{fill,color,border,background,table}.ts` | | Fills, theme colour maths, outlines, slide backgrounds, tables |
| `src/render/custgeom.ts` | 230 | Custom geometry and the shape-guide formula language |
| `src/render/numbering.ts` | | Auto-numbered bullets |
| `src/render/chart.ts`, `src/charts/*` | | Chart model and SVG renderer |
| `src/viewer/*` | 700 | `PptxViewer`, styles |
| `src/model/document.ts` | 1 040 | `PresentationDocument` editing model with undo |
| `src/editor/*` | 770 | `PptxEditor` UI |
| `src/export/pptx.ts`, `odp.ts`, `svg-path.ts` | | Exporters |
| `src/adapters/*` | | React, Vue, jQuery, Angular, Web Components |
| `src/utils/*` | | tinycolor port, base64, HTML escaping, image size, event emitter |
| `legacy/` | | PPTXjs v1, for reference and baseline comparison only |
| `scripts/` | | Maintenance tools (see below) |
| `demo/` | | Demo app (`npm run dev`) |

### About the ported renderer

`render/{shape,text,table,fill,color,border,background,nodes,geometry}.ts` were split automatically out of
the 14 000-line v1 closure (`scripts/port/split-legacy.cjs`), then fixed by hand (`scripts/port/fix-ported.cjs`).
They keep v1's style (`var`, `any`, loose equality) so a function can still be compared side by side with
`legacy/js/pptxjs.js`. Every function starts with a `// ported from legacy/js/pptxjs.js:<line>` comment.
New code is strict TypeScript. When you substantially change a ported function, convert it to modern style.

The simplified XML tree (`node["p:spPr"]["a:xfrm"]["attrs"]["rot"]`) groups children by tag and loses order
between different tags. `attrs.order` is a global post-order counter that records document order. The
renderer uses it as z-index, and `custgeom.ts`/`odp.ts` sort by it.

## Conventions

- TypeScript `strict`. Public APIs are fully typed and documented with TSDoc.
- No runtime dependencies besides `fflate`. Framework packages are optional peers.
- Never write to globals or prototypes. No `innerHTML` with unescaped data. URLs go through `sanitizeUrl`.
- Units: the public API uses CSS pixels (96 dpi). OOXML uses EMU (`EMU_PER_PX = 9525`). Angles are 60000ths of a degree.
- Errors: throw `PptxError` with a code. Inside the renderer, prefer `ctx.warn()` and keep going.

## Tests

```bash
npm test                 # everything (Vitest)
npm run test:watch
npm run test:coverage
```

| Suite | What it covers |
|---|---|
| `test/unit/*` | XML parser (safety, AlternateContent, pollution), tinycolor port, base64, URL sanitising, image sizes, numbering, custom geometry, bullet mapping, SVG path conversion |
| `test/integration/fixtures.test.ts` | Renders every `.pptx` in `test/fixtures` (sample deck and files attached to GitHub issues). Checks no slide fails, HTML is balanced and nothing is executable, plus per-issue regressions |
| `test/integration/security.test.ts` | Crafted files with XSS payloads in text, attributes, CDATA, hyperlinks and video links; prototype pollution; zip bomb and entry limits (HTML checked with a real HTML parser) |
| `test/integration/document.test.ts` | xmldom round trip, every `PresentationDocument` operation, undo/redo/rollback, rendering rules, PPTX and ODP export validity |
| `test/dom/*` | `PptxViewer` (modes, navigation, keyboard, autoplay, fetch, errors, re-render), `PptxEditor`, React/Vue/jQuery/Web Component adapters (jsdom) |

`test/helpers/make-pptx.ts` builds valid minimal presentations in memory. Use it for new tests instead of
adding binary fixtures when possible.

### Regression tooling against v1

| Command | Purpose |
|---|---|
| `node scripts/legacy-baseline.mjs` | Renders all fixtures with v1 under jsdom and writes `test/baseline/*.legacy.html` |
| `npm run fixtures` | Renders all fixtures with v2 and writes viewable pages to `test/output/*.html` |
| `node scripts/compare-baseline.mjs [name]` | Structural diff (shapes, images, tables, text per slide) between v1 and v2 |
| `scripts/validate-powerpoint.ps1 <dir>` | Windows only: opens exported `.pptx`/`.odp` files in PowerPoint, exports PDF and PNG previews |
| `/legacy/compare.html?f=<fixture>` | Shows v1's rendering in the browser (serve the repo root) |

## Adding support for something

1. Write a failing test, using `makePptx()` with the XML you need or a real file in `test/fixtures/`.
2. Find where the element is handled. `processNodesInSlide` (nodes.ts) is the entry point for shapes,
   `genTextBody` (text.ts) for text, `getShapeFill`/`getSolidFill` for colours.
3. Keep the renderer DOM-free. If something must be measured, add an option to `RenderContext` the way `measureHtmlWidth` works.
4. If the feature affects export, update `export/odp.ts`. PPTX export needs no changes because it is lossless.
5. Run `npm run fixtures` and check `test/output/` in a browser for visual changes.

## Building and releasing

```bash
npm run typecheck && npm run lint && npm test
npm run build            # dist/ (ESM, CJS, UMD, CSS, .d.ts)
npm run build:demo       # dist-demo/ static demo site
npm publish --access public
```

The `exports` map in `package.json` defines the entry points: `.`, `./react`, `./vue`, `./angular`, `./jquery`, `./style.css`.
Update `CHANGELOG.md` and `VERSION` in `src/index.ts`.

CI (`.github/workflows/ci.yml`) runs typecheck, lint, tests and the build on Node 20 and 22.
