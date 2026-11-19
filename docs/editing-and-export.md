# Editing and export

PPTXjs edits presentations model-first. Every change is applied to the presentation's real OOXML parts,
and the affected slide is then re-rendered from that XML. As a result:

- what you see is what gets saved: the view is always a rendering of the actual file
- `.pptx` export is lossless: parts you didn't touch (themes, masters, charts, media, custom XML) are written back byte for byte
- every edit is undoable, because undo restores the previous bytes of the parts that changed

## The editor UI

```ts
import { PptxEditor } from "@meshesha/pptxjs";
const editor = new PptxEditor("#editor", { fileName: "q3-review" });
await editor.load(file);
```

| Area | Features |
|---|---|
| Toolbar | Undo/redo · add text box · add shape (14 presets) · add picture · bold/italic/underline · font size · text colour · alignment · fill colour / no fill · bring to front / send to back · duplicate / delete shape · new / duplicate / move up/down / hide / delete slide · download `.pptx` / `.odp` |
| Slide panel | Live thumbnails; click to switch; hidden slides are dimmed |
| Canvas | Click to select, drag to move (<kbd>Shift</kbd> = one axis), 8 resize handles (<kbd>Shift</kbd> = keep aspect ratio), double-click or <kbd>Enter</kbd> to edit text |
| Keyboard | <kbd>Del</kbd>, arrows (+<kbd>Shift</kbd>), <kbd>Ctrl</kbd>+<kbd>Z</kbd>/<kbd>Y</kbd>/<kbd>D</kbd>/<kbd>B</kbd>/<kbd>I</kbd>/<kbd>U</kbd>, <kbd>Esc</kbd>, <kbd>PgUp</kbd>/<kbd>PgDn</kbd> |

Hide the built-in chrome with `toolbar: false` or `thumbnails: false` and drive the editor from your own UI:

```ts
editor.on("selectionchange", ({ selection, shape }) => updateMyPanel(shape));
myBoldButton.onclick = () => {
  const s = editor.selected;
  if (s) editor.document!.setTextStyle(s.slide, s.id, { bold: true });
};
```

`readOnly: true` keeps selection and export available but blocks edits.

## Programmatic editing

`PresentationDocument` has the same capabilities without any UI and works in Node.js:

```ts
import { PresentationDocument, exportPptxBytes } from "@meshesha/pptxjs";

const doc = await PresentationDocument.load(bytes);

// fill a template
for (const shape of doc.getShapes(0)) {
  if (shape.text.includes("{{customer}}")) doc.setShapeText(0, shape.id, shape.text.replace("{{customer}}", "ACME"));
}

// several operations as one undo step
doc.transact({ type: "slide", index: 0, label: "Add logo" }, () => {
  const id = doc.addImage(0, logoBytes, "image/png", { x: 20, y: 20, width: 120 });
  doc.reorderShape(0, id, "front");
});

await fs.writeFile("out.pptx", exportPptxBytes(doc));
```

See the [API reference](api.md#presentationdocument) for every method.

### How edits map to OOXML

| Operation | XML change |
|---|---|
| Move / resize / rotate | `a:xfrm` (`a:off`, `a:ext`, `@rot`). Placeholders that inherit their position get their own `a:xfrm`. Groups keep `chOff`/`chExt`, so children keep their relative layout. |
| Text | `a:p` elements are rebuilt (new paragraphs starting with Hebrew/Arabic get `rtl="1"`; new RTL text boxes are right-aligned); the first run's `a:rPr` and each paragraph's `a:pPr` are reused as templates. |
| Formatting | `a:rPr@b/@i/@u/@sz`, `a:solidFill`, `a:latin`/`a:cs`, `a:pPr@algn`, inserted in schema order. |
| Fill / outline | `p:spPr` fill choice / `a:ln`, inserted in schema order. |
| Add shape / text box / picture | New `p:sp` / `p:pic` with the next free `p:cNvPr@id`. Pictures add `ppt/media/imageN.*`, a relationship and a content-type default. |
| Slides | `p:sldIdLst`, `presentation.xml.rels`, `[Content_Types].xml`; duplicated slides copy their relationships except the notes page. |

## Export

### PPTX

```ts
const blob = exportPptx(doc);                                // or editor.exportPptx()
const bytes = exportPptxBytes(doc, { pruneUnusedMedia: true });
```

The package is re-zipped with `[Content_Types].xml` first, as OPC recommends. Media no longer referenced
by any relationship (for example after deleting slides) is dropped unless `pruneUnusedMedia: false` is set.

Validation: exported files were opened in Microsoft PowerPoint 365 (via `scripts/validate-powerpoint.ps1`).
The test files had slides deleted, duplicated and moved, shapes, text boxes and pictures added, and text edited.
All of them opened and rendered correctly.

### ODP (OpenDocument Presentation)

```ts
const blob = exportOdp(doc, { includeMasterShapes: true, includeNotes: true, skipHiddenSlides: false });
```

The converter uses the same style resolution as the renderer, so colours, fonts and positions match the viewer:

| Converted | Notes |
|---|---|
| Text boxes, placeholders | Right-to-left paragraphs (`style:writing-mode="rl-tb"`), fonts, sizes, colours, bold, italic, underline, strike, caps, alignment, indents; inherited master/layout styles are resolved |
| Bullets, numbering | Written as text so they look identical everywhere (not as ODF lists) |
| Preset shapes, custom geometry | `draw:custom-shape` with an explicit `draw:enhanced-path` built from the renderer's geometry |
| Fill, outline, rotation, flip | Solid colours (the first colour of gradients), outline colour, width and dash |
| Pictures | Embedded as-is in `Pictures/` |
| Tables | `table:table` with column widths, merged cells and cell fills |
| Charts | Rendered to SVG and embedded as a picture |
| Groups, connectors | `draw:g`, `draw:line` |
| Backgrounds | Colour (page style) or picture (full-slide image) |
| Layout/master decorations | Non-placeholder shapes of the layout and master |
| Speaker notes | `presentation:notes` |

Validation: exported `.odp` files were opened in Microsoft PowerPoint 365. All slides imported, with
backgrounds, shapes, text and pictures.

Known limitations of ODP export:

- picture, gradient and pattern fills inside shapes become solid colours or no fill
- text effects (WordArt outlines, shadows, glow, 3D) are not converted
- charts are pictures, not editable chart objects
- SmartArt is converted from its cached drawing (shapes), not as SmartArt
- video and audio are not embedded; the poster image is kept
- animations and transitions are not converted

## Limitations of editing

- Only top-level shapes are selectable in the editor UI; a group moves as a whole. The model can edit
  any shape by id, including shapes inside groups.
- Text editing replaces a shape's text with plain paragraphs. Character formatting applies to the
  whole shape. Mixed formatting inside a paragraph is kept only until that shape's text is edited.
- Tables, charts and SmartArt can be moved, resized, duplicated and deleted, but not edited internally.
- Rotation from the UI is not available yet (`setShapeRotation()` works).
