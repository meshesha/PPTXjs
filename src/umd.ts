/**
 * Entry point of the UMD/IIFE bundle (`<script src="pptxjs.umd.js">`), exposed
 * as `window.PPTXjs`. It also installs the jQuery plugin when jQuery is present
 * (drop-in replacement for v1) and registers the `<pptx-viewer>` /
 * `<pptx-editor>` custom elements.
 */
import { defineCustomElements } from "./adapters/element";
export * from "./index";
export { installJQueryPlugin, fromLegacyOptions } from "./adapters/jquery";

defineCustomElements();
