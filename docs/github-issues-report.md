# GitHub issues: status report

Snapshot of all 40 issues of [meshesha/PPTXjs](https://github.com/meshesha/PPTXjs/issues) (27 open, 13 closed),
reviewed against v2. Where an issue attached a file, it is in `test/fixtures/issues/issue-<n>.pptx` and covered
by the fixture test suite: every slide renders without errors.

## Bugs fixed in v2

| Issue | Title / symptom | Root cause | Fix | Covered by |
|---|---|---|---|---|
| [#44](https://github.com/meshesha/PPTXjs/issues/44) | `can't access property "p:drawing", node is null` | Missing SmartArt drawing part returned `null`, but only `undefined` was checked | Tree helpers treat `null` as missing; drawing lookup made safe | fixture `issue-44` (26 slides) |
| [#43](https://github.com/meshesha/PPTXjs/issues/43) | `can't access property "x", off is undefined` | Shape inheriting its position from the layout placeholder; only the slide's own `xfrm` was read | Position/size fall back to layout, then master | fixture `issue-43` |
| [#42](https://github.com/meshesha/PPTXjs/issues/42) | `app is null` | `docProps/app.xml` is optional but was required | App version optional | fixture `issue-42`, unit test |
| [#41](https://github.com/meshesha/PPTXjs/issues/41) | `pathNodes.attrs is undefined` | Custom geometry with several `<a:path>` elements | New custom-geometry engine (all paths, single `lnTo`, multiple arcs, `quadBezTo`, shape guides) | fixtures `issue-41`, `issue-41b`, unit tests |
| [#40](https://github.com/meshesha/PPTXjs/issues/40) | `` instead of `•` | Bullet char inherited from the master but its font was not, and Symbol/Wingdings private-use code points were emitted as-is | Bullet font inherited independently; Symbol/Wingdings/Webdings mapped to Unicode | fixture `issue-40`, unit test |
| [#39](https://github.com/meshesha/PPTXjs/issues/39) | `node['p:blipFill'] is undefined` | Picture wrapped in `mc:AlternateContent` (PowerPoint for Mac) | Markup-compatibility blocks resolved in the parser | fixture `issue-39`, unit test |
| [#38](https://github.com/meshesha/PPTXjs/issues/38) | Master formatting overrides every slide | Slide-level `showMasterSp="0"` ignored (only the layout's flag was read) | Slide flag hides layout and master graphics | integration test |
| [#33](https://github.com/meshesha/PPTXjs/issues/33) | `Cannot read properties of undefined (reading 'a:noFill')` | Fill helper called with a missing node | Missing nodes tolerated; slide-level error isolation | fixtures |
| [#32](https://github.com/meshesha/PPTXjs/issues/32) | `reading 'w'` in `genShape` | Same as #41 | as #41 | fixture `issue-32` (38 slides) |
| [#31](https://github.com/meshesha/PPTXjs/issues/31) | Corrupted zip: can't find end of central directory | JSZip v2 errors, and invalid input was not detected | fflate; input validated with clear `INVALID_FILE` errors | security tests |
| [#30](https://github.com/meshesha/PPTXjs/issues/30), [#28](https://github.com/meshesha/PPTXjs/issues/28), [#22](https://github.com/meshesha/PPTXjs/issues/22), [#17](https://github.com/meshesha/PPTXjs/issues/17) | `X[g].exec is not a function`; breaks Vue/ExtJS; prototype pollution | `Object.prototype.set` assigned at runtime | Removed; parser blocks `__proto__`/`constructor` | unit and security tests |
| [#29](https://github.com/meshesha/PPTXjs/issues/29) | Downloads larger than 1 MB fail on a page using PPTXjs | Most likely the same prototype pollution: other code saw the extra enumerable `set` property | Fixed by the above (not reproduced; no sample) | – |
| [#26](https://github.com/meshesha/PPTXjs/issues/26), [#27](https://github.com/meshesha/PPTXjs/issues/27), [#21](https://github.com/meshesha/PPTXjs/issues/21) | "When parsing some PowerPoint files an error is reported" (screenshots only) | Various crashes on unexpected structure | Errors are now isolated per slide and reported via `warning`; the crash paths above are fixed | fixtures |
| [#20](https://github.com/meshesha/PPTXjs/issues/20) | `<sapn>` instead of `<span>` | Typo | Fixed | fixture test |
| [#18](https://github.com/meshesha/PPTXjs/issues/18) | Linked video URL mangled ("address splicing") | External relationship targets were rewritten with `replace("../", "ppt/")` | External targets kept verbatim; internal targets resolved per OPC | security test |
| [#15](https://github.com/meshesha/PPTXjs/issues/15) | Incorrect page order | Order taken from file names | Order from `p:sldIdLst` | integration test, fixture `issue-15` |
| [#14](https://github.com/meshesha/PPTXjs/issues/14) | Errors cannot be caught | Callback-based API with uncaught exceptions | `load()` returns a Promise; `error`/`warning` events; `PptxError` codes | DOM tests |
| [#12](https://github.com/meshesha/PPTXjs/issues/12) | File displays incorrectly | Several: custom geometry, bullets, text wrapping (`&nbsp;`) | Improved (custom geometry, bullets, word wrapping). 3D effects are still approximations (see feature requests) | fixture `issue-12` (53 slides) |
| [#3](https://github.com/meshesha/PPTXjs/issues/3) (closed) | XSS in hyperlinks | v1 escaped HTML but still allowed `javascript:` | URL allow-list, plus parser-level escaping of all content | security tests |
| [#7](https://github.com/meshesha/PPTXjs/issues/7)–[#11](https://github.com/meshesha/PPTXjs/issues/11), [#13](https://github.com/meshesha/PPTXjs/issues/13), [#16](https://github.com/meshesha/PPTXjs/issues/16) (closed) | Earlier crashes | Fixed in v1.21 | Still pass in v2 | fixtures `issue-7` … `issue-16` |

In addition, the code audit found and fixed these bugs, which had no issue: unclosed `<table>` tags (later
slides nested inside earlier ones), `controls` on iframes and duplicate `id="audio_player"`, a red debug
border on every group, `stroke-width="NaN"`, a stray `'` in every SVG tag, `undefined` printed as an
attribute, a bitwise `&` used instead of `&&`, an undefined variable (`cocacl_h`) in colour maths, implicit
globals (`txBoxVal`, `is_first_br`, `dgmQuickStyleFileName`, `cx`, gear-shape variables), a busy-wait loop for
image sizes, object-URL leaks, a page-wide `keydown` handler blocking typing, `dingbat.js` missing from the
demo page, and slides being lost when the HTML was concatenated. See [security.md](security.md) for the
security findings.

## Feature requests: decisions needed

| # | Request | Status in v2 | Suggested next step | Effort |
|---|---|---|---|---|
| [#25](https://github.com/meshesha/PPTXjs/issues/25) | npm package | **Done:** `@meshesha/pptxjs`, ESM/CJS/UMD and types | Decide the final package name (the unscoped `pptxjs` is taken on npm) and publish | S |
| [#19](https://github.com/meshesha/PPTXjs/issues/19) | Convert pptx to HTML in Node.js | **Done:** `renderToHtml()` / `renderToHtmlString()` | – | – |
| [#23](https://github.com/meshesha/PPTXjs/issues/23) | Web Worker / independent API | **Partly done:** DOM-free API that works in workers; framework adapters | Ship a ready-made worker entry and a `useWorker` viewer option | M |
| [#24](https://github.com/meshesha/PPTXjs/issues/24) | Convert HTML slides back to PowerPoint | **Done for edited decks:** editor + lossless `.pptx` export (and `.odp`) | Optionally a "create presentation from scratch" API (new empty deck with default theme) | M |
| [#6](https://github.com/meshesha/PPTXjs/issues/6) | Animations | Not supported | Parse `p:timing` for the common entrance/exit effects (fade, fly-in, appear) and play them in slideshow mode, plus slide transitions from `p:transition` | L |
| [#2](https://github.com/meshesha/PPTXjs/issues/2) | Reveal.js-compatible output | v1's partial support was removed (it loaded scripts at runtime) | Provide a documented `toRevealSections()` helper that returns `<section>` markup | S |
| [#1](https://github.com/meshesha/PPTXjs/issues/1) | EMF/WMF images | Browsers cannot display them; shown as broken images, as in v1 | Use the PNG fallback that many files embed, or add an optional EMF→SVG converter (separate package) | M–L |
| [#12](https://github.com/meshesha/PPTXjs/issues/12) | Better fidelity: 3D shapes, effects | 3D extrusions/bevels approximated; shadows/glow partial | Add CSS/SVG approximations for bevel and extrusion, plus gradient fills in ODP export | L |
| – | Editor: rotation handle, multi-select, grouping, table/chart editing, rich-text runs | Not yet (model supports rotation) | Prioritise by user demand | M each |
| – | Speaker notes editing; comments; thumbnails view in viewer | Notes are read-only | `setNotes()` in the model; notes panel in the viewer | S–M |
| – | Accessibility: alt text output, reading order, screen-reader mode | Partial (ARIA on viewer chrome) | Emit `alt` from `descr`, `aria-label` per slide, expose an outline/text-only mode | M |

Effort: S = about a day, M = a few days, L = a week or more.
