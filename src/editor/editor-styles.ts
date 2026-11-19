/** Styles of the editor chrome. */
export const EDITOR_CSS = `
.pptxjs-editor { display: flex; flex-direction: column; height: 100%; min-height: 420px; border: 1px solid #d0d4dc; border-radius: 8px; overflow: hidden; background: #f4f5f8; outline: none; font: 13px system-ui, -apple-system, "Segoe UI", Roboto, sans-serif; color: #1f2430; }
.pptxjs-editor-toolbar { display: flex; flex-wrap: wrap; gap: 6px; padding: 6px 8px; background: #fff; border-bottom: 1px solid #d0d4dc; }
.pptxjs-editor-toolbar[hidden] { display: none; }
.pptxjs-tb-group { display: flex; align-items: center; gap: 3px; padding-right: 6px; border-right: 1px solid #e3e6ec; }
.pptxjs-tb-group:last-child { border-right: 0; }
.pptxjs-tb-btn { min-width: 30px; height: 30px; padding: 0 7px; border: 1px solid transparent; border-radius: 5px; background: transparent; color: inherit; cursor: pointer; font: inherit; font-weight: 600; }
.pptxjs-tb-btn:hover:not(:disabled) { background: #eef1f7; border-color: #d6dbe6; }
.pptxjs-tb-btn.pptxjs-on { background: #dfe7fb; border-color: #9fb5ec; }
.pptxjs-tb-btn:disabled, .pptxjs-tb-num:disabled, .pptxjs-tb-select:disabled, .pptxjs-tb-color:disabled { opacity: .4; cursor: default; }
.pptxjs-tb-btn:focus-visible, .pptxjs-tb-select:focus-visible, .pptxjs-tb-num:focus-visible { outline: 2px solid #4775d1; outline-offset: 1px; }
.pptxjs-tb-select, .pptxjs-tb-num { height: 30px; border: 1px solid #d6dbe6; border-radius: 5px; background: #fff; font: inherit; padding: 0 4px; }
.pptxjs-tb-num { width: 56px; }
.pptxjs-tb-color { width: 30px; height: 30px; padding: 2px; border: 1px solid #d6dbe6; border-radius: 5px; background: #fff; }
.pptxjs-tb-label { color: #6a7180; font-size: 12px; padding: 0 2px; }
.pptxjs-editor-body { display: flex; flex: 1; min-height: 0; }
.pptxjs-editor-thumbs { width: 184px; flex: none; overflow-y: auto; padding: 10px 8px; background: #eceef3; border-right: 1px solid #d0d4dc; display: flex; flex-direction: column; gap: 10px; }
.pptxjs-editor-thumbs[hidden] { display: none; }
.pptxjs-thumb { position: relative; display: flex; gap: 6px; align-items: flex-start; cursor: pointer; border-radius: 4px; padding: 2px; }
.pptxjs-thumb-num { width: 14px; text-align: right; font-size: 11px; color: #6a7180; padding-top: 2px; }
.pptxjs-thumb-frame { position: relative; overflow: hidden; background: #fff; outline: 1px solid #c8cdd8; border-radius: 2px; }
.pptxjs-thumb-frame > .slide { position: absolute; top: 0; left: 0; margin: 0; }
.pptxjs-thumb-active .pptxjs-thumb-frame { outline: 2px solid #4775d1; }
.pptxjs-thumb-hidden .pptxjs-thumb-frame { opacity: .45; }
.pptxjs-editor-stage { flex: 1; min-width: 0; overflow: auto; padding: 20px; }
.pptxjs-editor-stage > .pptxjs-viewer { height: 100%; }
.pptxjs-editor-stage .pptxjs-slide-wrap { box-shadow: 0 2px 10px rgba(0,0,0,.15); }
.pptxjs-editor-stage .pptx-shape > * { cursor: move; }
.pptxjs-editor-overlay { position: absolute; z-index: 2147483000; border: 1.5px solid #4775d1; box-sizing: border-box; cursor: move; transform-origin: 50% 50%; }
.pptxjs-editor-overlay[hidden] { display: none; }
.pptxjs-handle { position: absolute; width: 9px; height: 9px; background: #fff; border: 1.5px solid #4775d1; border-radius: 2px; box-sizing: border-box; }
.pptxjs-handle-nw { left: -5px; top: -5px; cursor: nwse-resize; }
.pptxjs-handle-n { left: calc(50% - 4.5px); top: -5px; cursor: ns-resize; }
.pptxjs-handle-ne { right: -5px; top: -5px; cursor: nesw-resize; }
.pptxjs-handle-e { right: -5px; top: calc(50% - 4.5px); cursor: ew-resize; }
.pptxjs-handle-se { right: -5px; bottom: -5px; cursor: nwse-resize; }
.pptxjs-handle-s { left: calc(50% - 4.5px); bottom: -5px; cursor: ns-resize; }
.pptxjs-handle-sw { left: -5px; bottom: -5px; cursor: nesw-resize; }
.pptxjs-handle-w { left: -5px; top: calc(50% - 4.5px); cursor: ew-resize; }
.pptxjs-editor-text { position: absolute; z-index: 2147483001; box-sizing: border-box; padding: 4px 6px; border: 2px solid #4775d1; border-radius: 3px; background: rgba(255,255,255,.96); resize: none; font-family: Calibri, Arial, sans-serif; line-height: 1.2; color: #111; }
`;
