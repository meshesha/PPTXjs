/**
 * Base stylesheet for rendered slides (v1 css/pptxjs.css), scoped to the
 * `.pptxjs` container class so it cannot leak into the host page. The viewer
 * injects it automatically; it is also shipped as `@meshesha/pptxjs/style.css`.
 */
export const BASE_CSS = `
.pptxjs { position: relative; }
/* isolate slides from host-page CSS (Bootstrap/Tailwind set border-box, max-width:100% ...) */
.pptxjs .slide, .pptxjs .slide *, .pptxjs .slide *::before, .pptxjs .slide *::after { box-sizing: content-box; }
.pptxjs .slide img, .pptxjs .slide svg, .pptxjs .slide video, .pptxjs .slide iframe { max-width: none; max-height: none; }
.pptxjs .slide p, .pptxjs .slide table { margin: 0; }
/* slides are laid out in absolute LTR coordinates; a host page with dir="rtl" must not mirror them.
   Paragraph direction comes from the presentation (or the textDirection option). */
.pptxjs .slide { direction: ltr; unicode-bidi: isolate; }
.pptxjs .slide { line-height: normal; letter-spacing: normal; text-transform: none; text-align: left; font-size: 16px; }
.pptxjs .slide { position: relative; overflow: hidden; margin: 0 auto 24px; background: #fff; box-sizing: content-box; }
.pptxjs .slide-error { position: absolute; inset: 0; display: flex; align-items: center; justify-content: center; color: #a00; font: 16px sans-serif; background: #fff4f4; }
.pptxjs .slide div.block { position: absolute; top: 0; left: 0; width: 100%; line-height: 1; }
.pptxjs .slide div.content, .pptxjs .slide div.diagram-content { display: flex; flex-direction: column; }
.pptxjs .slide div.content-rtl { display: flex; flex-direction: column; direction: rtl; }
.pptxjs .slide .pregraph-rtl { direction: rtl; }
.pptxjs .slide .pregraph-ltr { direction: ltr; }
.pptxjs .slide .pregraph-inherit { direction: inherit; }
.pptxjs .slide .slide-prgrph { width: 100%; }
/* keep runs of spaces/tabs but wrap at word boundaries like PowerPoint (v1 used &nbsp;) */
.pptxjs .slide .text-block { white-space: pre-wrap; tab-size: 4; }
.pptxjs .slide .line-break-br::before { content: "\\A"; white-space: pre; }
.pptxjs .slide div.v-up { justify-content: flex-start; }
.pptxjs .slide div.v-mid { justify-content: center; }
.pptxjs .slide div.v-down { justify-content: flex-end; }
.pptxjs .slide div.h-left { justify-content: flex-start; align-items: flex-start; text-align: left; }
.pptxjs .slide div.h-left-rtl { justify-content: flex-end; align-items: flex-end; text-align: left; }
.pptxjs .slide div.h-mid { justify-content: center; align-items: center; text-align: center; }
.pptxjs .slide div.h-right { justify-content: flex-end; align-items: flex-end; text-align: right; }
.pptxjs .slide div.h-right-rtl { justify-content: flex-start; align-items: flex-start; text-align: right; }
.pptxjs .slide div.h-just, .pptxjs .slide div.h-dist { text-align: justify; }
.pptxjs .slide div.up-left { justify-content: flex-start; align-items: flex-start; text-align: left; }
.pptxjs .slide div.up-center { justify-content: flex-start; align-items: center; }
.pptxjs .slide div.up-right { justify-content: flex-start; align-items: flex-end; }
.pptxjs .slide div.center-left { justify-content: center; align-items: flex-start; text-align: left; }
.pptxjs .slide div.center-center { justify-content: center; align-items: center; }
.pptxjs .slide div.center-right { justify-content: center; align-items: flex-end; }
.pptxjs .slide div.down-left { justify-content: flex-end; align-items: flex-start; text-align: left; }
.pptxjs .slide div.down-center { justify-content: flex-end; align-items: center; }
.pptxjs .slide div.down-right { justify-content: flex-end; align-items: flex-end; }
.pptxjs .slide li.slide { margin: 10px 0; font-size: 18px; }
.pptxjs .slide table { position: absolute; }
.pptxjs .slide svg.drawing { position: absolute; overflow: visible; }
.pptxjs .slide svg.pptx-chart { display: block; }
.pptxjs .slide .pptx-media-link { position: absolute; left: 50%; top: 50%; transform: translate(-50%, -50%); width: 56px; height: 56px; border-radius: 50%; background: rgba(0,0,0,.6); color: #fff; font-size: 26px; line-height: 56px; text-align: center; text-decoration: none; }
`;
