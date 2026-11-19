/** Styles of the viewer chrome (toolbar, progress, slideshow, transitions). */
export const VIEWER_CSS = `
.pptxjs-viewer { position: relative; outline: none; font-family: system-ui, -apple-system, "Segoe UI", Roboto, sans-serif; }
.pptxjs-viewer:focus-visible { box-shadow: 0 0 0 2px #4775d1; }
.pptxjs-stage { position: relative; width: 100%; }
.pptxjs-slides { display: flex; flex-direction: column; align-items: center; gap: var(--pptx-gap, 24px); }
.pptxjs-slide-wrap { position: relative; width: calc(var(--pptx-w) * var(--pptx-scale, 1)); height: calc(var(--pptx-h) * var(--pptx-scale, 1)); flex: none; }
.pptxjs-slide-wrap[hidden] { display: none; }
.pptxjs-slide-wrap > .slide { position: absolute; top: 0; left: 0; margin: 0; transform: scale(var(--pptx-scale, 1)); transform-origin: 0 0; box-shadow: 0 1px 4px rgba(0,0,0,.18); }
.pptxjs-mode-slideshow .pptxjs-slides { gap: 0; }
.pptxjs-mode-slideshow .pptxjs-slide-wrap > .slide { box-shadow: none; }
.pptxjs-progress { height: 4px; background: rgba(0,0,0,.08); overflow: hidden; margin-bottom: 8px; border-radius: 2px; }
.pptxjs-progress[hidden] { display: none; }
.pptxjs-progress-bar { height: 100%; width: 0; background: #4775d1; transition: width .15s ease; }
.pptxjs-toolbar { display: flex; gap: 6px; align-items: center; justify-content: center; padding: 8px; }
.pptxjs-toolbar[hidden] { display: none; }
.pptxjs-btn { display: inline-grid; place-items: center; box-sizing: border-box; min-width: 36px; height: 32px; border: 1px solid rgba(0,0,0,.15); border-radius: 6px; background: #fff; color: #222; cursor: pointer; font-size: 15px; line-height: 1; }
.pptxjs-btn:hover { background: #f0f3fa; }
.pptxjs-btn:focus-visible { outline: 2px solid #4775d1; outline-offset: 1px; }
.pptxjs-counter { min-width: 72px; text-align: center; font-variant-numeric: tabular-nums; color: #333; font-size: 14px; }
.pptxjs-error { padding: 12px 16px; margin: 8px 0; border: 1px solid #e3a3a3; background: #fff4f4; color: #8a1c1c; border-radius: 6px; font-size: 14px; }
.pptxjs-fullscreen { background: #000; display: flex; flex-direction: column; justify-content: center; }
.pptxjs-fullscreen .pptxjs-toolbar { position: absolute; bottom: 0; left: 0; right: 0; opacity: 0; transition: opacity .2s; }
.pptxjs-fullscreen:hover .pptxjs-toolbar { opacity: .9; }
.pptxjs-fullscreen .pptxjs-counter { color: #eee; }
.pptxjs-anim-fade { animation: pptxjs-fade var(--pptx-transition, .4s) ease both; }
.pptxjs-anim-slide { animation: pptxjs-slide var(--pptx-transition, .4s) ease both; }
.pptxjs-anim-zoom { animation: pptxjs-zoom var(--pptx-transition, .4s) ease both; }
@keyframes pptxjs-fade { from { opacity: 0; } to { opacity: 1; } }
@keyframes pptxjs-slide { from { opacity: 0; transform: translateX(8%); } to { opacity: 1; transform: none; } }
@keyframes pptxjs-zoom { from { opacity: 0; transform: scale(.92); } to { opacity: 1; transform: none; } }
/* right-to-left UI: mirror the direction icons and the slide transition */
.pptxjs-viewer[dir="rtl"] :is(.pptxjs-first, .pptxjs-prev, .pptxjs-next, .pptxjs-last) svg { transform: scaleX(-1); }
.pptxjs-viewer[dir="rtl"] .pptxjs-anim-slide { animation-name: pptxjs-slide-rtl; }
@keyframes pptxjs-slide-rtl { from { opacity: 0; transform: translateX(-8%); } to { opacity: 1; transform: none; } }
@media (prefers-reduced-motion: reduce) { .pptxjs-anim-fade, .pptxjs-anim-slide, .pptxjs-anim-zoom { animation: none; } }
`;
