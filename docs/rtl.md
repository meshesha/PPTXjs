# Right-to-left languages (Hebrew, Arabic, Persian, Urdu …)

## Summary

| | Behaviour |
|---|---|
| Paragraphs marked RTL in the file (`a:pPr rtl="1"`, from the paragraph, layout, master or presentation defaults) | Rendered RTL: base direction, alignment, bullets and numbering (including Hebrew `א-` numbering) on the right |
| Paragraphs without a direction | Left to right by default, as PowerPoint does. `textDirection: "auto"` detects RTL text; `textDirection: "rtl"` makes them RTL |
| Mixed text (`שלום, 2024 (test).`) | The Unicode bidi algorithm runs with the paragraph's base direction, so punctuation and numbers land in the right place |
| Runs with an RTL language tag (`lang="he-IL"`, `ar-EG`, `fa`, `ur`, `yi`, `ckb`, `ku-Arab` …) | Marked RTL |
| Viewer and editor UI | RTL on pages with `dir="rtl"` (or `dir: "rtl"`): mirrored toolbar, ← moves forward, mirrored slide transition, slide panel on the right |
| Host page direction | Never mirrors slide layouts; slides keep their PowerPoint coordinates |
| Editor | Hebrew/Arabic typed into a shape is saved with `rtl="1"`; new RTL text boxes are right-aligned; the text box uses `dir="auto"` |
| ODP export | RTL paragraphs get `style:writing-mode="rl-tb"`, physical alignment and right-side indents; `textDirection` is supported |

## Options

```ts
new PptxViewer(el, {
  textDirection: "auto", // "file" (default) | "auto" | "rtl"
  dir: "rtl",            // "auto" (default: follows the page) | "ltr" | "rtl"
});
```

### `textDirection`

Decides the direction of paragraphs for which the presentation doesn't say. An explicit `rtl="0"` or `rtl="1"`
in the file always wins.

- `"file"` (default): exactly like PowerPoint. Files authored in Hebrew/Arabic Office mark their paragraphs, so they render correctly.
- `"auto"`: the paragraph's first strong character decides, like HTML `dir="auto"`. Use this for files from tools that don't write direction information, such as some converters and generators. Detected RTL paragraphs with no alignment are right-aligned.
- `"rtl"`: every paragraph without a direction is RTL. Use this for decks that are entirely in an RTL language but lack the flags.

The same option exists for `renderToHtml()`, `renderPresentation()` and `exportOdp()`.

### `dir`

The direction of the viewer and editor chrome only. `"auto"` reads the nearest `dir` attribute of the
container, falling back to its CSS `direction`.

## Fonts

Browsers fall back to a system font when the presentation's font is not installed. For Hebrew and Arabic,
load suitable web fonts (for example Noto Sans Hebrew / Noto Naskh Arabic) on the page, or map the
presentation's fonts with CSS `@font-face` rules using the same family names.
