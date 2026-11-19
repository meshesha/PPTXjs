/**
 * Angular integration.
 *
 * Angular libraries must be compiled with the Angular compiler (ng-packagr),
 * so instead of shipping decorators this entry point registers standard custom
 * elements, which Angular supports natively:
 *
 * ```ts
 * // main.ts
 * import { definePptxElements } from "@meshesha/pptxjs/angular";
 * definePptxElements();
 *
 * // any standalone component
 * @Component({
 *   selector: "app-deck",
 *   standalone: true,
 *   schemas: [CUSTOM_ELEMENTS_SCHEMA],
 *   template: `<pptx-viewer [attr.src]="url" mode="slideshow" (pptx-load)="loaded($event)"></pptx-viewer>`,
 * })
 * export class DeckComponent { url = "/assets/deck.pptx"; loaded(e: Event) {} }
 * ```
 *
 * A ready-to-copy standalone Angular component that wraps {@link PptxViewer}
 * directly is in `examples/angular/pptx-viewer.component.ts`.
 */
export { defineCustomElements as definePptxElements, PptxViewerElement, PptxEditorElement } from "./element";
export { PptxViewer, type ViewerOptions } from "../viewer/viewer";
export { PptxEditor, type EditorOptions } from "../editor/editor";
