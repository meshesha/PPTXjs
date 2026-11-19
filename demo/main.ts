import {
  PptxViewer,
  PptxEditor,
  PptxError,
  renderToHtml,
  loadPptx,
  exportOdp,
  BASE_CSS,
  defineCustomElements,
  type ViewerOptions,
  type RenderedPresentation,
} from "@meshesha/pptxjs";

const $ = <T extends HTMLElement = HTMLElement>(sel: string) => document.querySelector<T>(sel)!;
const SAMPLE_URL = "./samples/Sample_12.pptx";

let current: File | string = SAMPLE_URL;
let lastRender: RenderedPresentation | null = null;

// ---------------------------------------------------------------- event log
function log(msg: string, kind: "info" | "warn" | "error" = "info") {
  const li = document.createElement("li");
  li.className = kind;
  li.textContent = `${new Date().toLocaleTimeString()}  ${msg}`;
  const list = $("#log");
  list.prepend(li);
  while (list.children.length > 60) list.lastElementChild?.remove();
}

// ---------------------------------------------------------------- viewer
function viewerOptions(): ViewerOptions {
  const theme = $<HTMLSelectElement>("#opt-theme").value;
  const scale = $<HTMLSelectElement>("#opt-scale").value;
  return {
    mode: $<HTMLSelectElement>("#opt-mode").value as ViewerOptions["mode"],
    scale: scale === "fit" ? "fit" : Number(scale),
    themeProcess: theme === "true" ? true : theme === "false" ? false : "colorsAndImageOnly",
    skipHiddenSlides: $<HTMLInputElement>("#opt-hidden").checked,
    textDirection: $<HTMLSelectElement>("#opt-textdir").value as ViewerOptions["textDirection"],
    dir: $<HTMLSelectElement>("#opt-dir").value as ViewerOptions["dir"],
    externalMedia: $<HTMLInputElement>("#opt-external").checked,
    slideshow: {
      transition: $<HTMLSelectElement>("#opt-transition").value as "fade",
      autoplay: Number($<HTMLInputElement>("#opt-autoplay").value) || false,
      loop: $<HTMLInputElement>("#opt-loop").checked,
    },
  };
}

let viewer = createViewer();
function createViewer(): PptxViewer {
  const v = new PptxViewer("#viewer", viewerOptions());
  v.on("loadstart", () => log("loading…"));
  v.on("load", ({ presentation: p }) => log(`loaded ${p.slides.length} slides (${p.width}×${p.height}px) in ${p.renderTime} ms`));
  v.on("warning", ({ warning }) => log(warning.message, "warn"));
  v.on("error", ({ error }) => log(`${error.code}: ${error.message}`, "error"));
  v.on("slidechange", ({ index }) => log(`slide ${index + 1}`));
  v.on("modechange", ({ mode }) => log(`mode: ${mode}`));
  return v;
}

async function reloadViewer() {
  viewer.destroy();
  viewer = createViewer();
  await viewer.load(current).catch(() => undefined);
}

// options that can change without reloading
$("#opt-mode").addEventListener("change", () => viewer.setOptions({ mode: viewerOptions().mode }));
$("#opt-scale").addEventListener("change", () => viewer.setOptions({ scale: viewerOptions().scale }));
for (const id of ["#opt-transition", "#opt-loop"]) $(id).addEventListener("change", () => viewer.setOptions({ slideshow: viewerOptions().slideshow }));
$("#opt-autoplay").addEventListener("change", () => {
  const s = Number($<HTMLInputElement>("#opt-autoplay").value);
  viewer.setOptions({ slideshow: viewerOptions().slideshow });
  if (s > 0) {
    viewer.setMode("slideshow");
    $<HTMLSelectElement>("#opt-mode").value = "slideshow";
    viewer.play(s);
  } else viewer.pause();
});
// options that need a re-render
$("#opt-dir").addEventListener("change", () => viewer.setOptions({ dir: viewerOptions().dir }));
for (const id of ["#opt-theme", "#opt-hidden", "#opt-external", "#opt-textdir"]) $(id).addEventListener("change", () => void reloadViewer());
$("#btn-prev").addEventListener("click", () => viewer.prev());
$("#btn-next").addEventListener("click", () => viewer.next());
$("#btn-fs").addEventListener("click", () => void viewer.toggleFullscreen());

// ---------------------------------------------------------------- editor (created lazily)
let editor: PptxEditor | null = null;
async function ensureEditor() {
  if (editor) return;
  editor = new PptxEditor("#editor", { fileName: "edited-presentation" });
  editor.on("error", ({ error }) => log(`editor: ${error.message}`, "error"));
  editor.on("change", ({ label }) => log(`edit: ${label}`));
  await editor.load(current).catch(() => undefined);
}

// ---------------------------------------------------------------- tabs
const tabs = Array.from(document.querySelectorAll<HTMLButtonElement>('[role="tab"]'));
for (const tab of tabs) {
  tab.addEventListener("click", () => selectTab(tab));
  tab.addEventListener("keydown", (e) => {
    const i = tabs.indexOf(tab);
    if (e.key === "ArrowRight") selectTab(tabs[(i + 1) % tabs.length], true);
    if (e.key === "ArrowLeft") selectTab(tabs[(i - 1 + tabs.length) % tabs.length], true);
  });
}
function selectTab(tab: HTMLButtonElement, focus = false) {
  for (const t of tabs) {
    const on = t === tab;
    t.setAttribute("aria-selected", String(on));
    t.tabIndex = on ? 0 : -1;
    const panel = document.getElementById(t.getAttribute("aria-controls")!)!;
    panel.hidden = !on;
    panel.classList.toggle("active", on);
  }
  if (focus) tab.focus();
  if (tab.id === "t-editor") void ensureEditor();
  if (tab.id === "t-frameworks") mountWebComponent();
}

// ---------------------------------------------------------------- file loading
async function openSource(src: File | string) {
  current = src;
  $("#file-name").textContent = typeof src === "string" ? "Sample_12.pptx" : src.name;
  lastRender = null;
  $<HTMLButtonElement>("#btn-dl-html").disabled = true;
  await viewer.load(src).catch(() => undefined);
  if (editor) await editor.load(src).catch(() => undefined);
}
$<HTMLInputElement>("#file").addEventListener("change", (e) => {
  const f = (e.target as HTMLInputElement).files?.[0];
  if (f) void openSource(f);
});
$("#load-sample").addEventListener("click", () => void openSource(SAMPLE_URL));

// drag & drop anywhere on the page
const drop = $("#drop");
let dragDepth = 0;
document.addEventListener("dragenter", (e) => {
  if (e.dataTransfer?.types.includes("Files")) {
    dragDepth++;
    drop.hidden = false;
  }
});
document.addEventListener("dragleave", () => {
  if (--dragDepth <= 0) drop.hidden = true;
});
document.addEventListener("dragover", (e) => e.preventDefault());
document.addEventListener("drop", (e) => {
  e.preventDefault();
  dragDepth = 0;
  drop.hidden = true;
  const f = e.dataTransfer?.files[0];
  if (f) {
    if (!/\.pptx$/i.test(f.name)) log(`${f.name} is not a .pptx file`, "error");
    else void openSource(f);
  }
});

// ---------------------------------------------------------------- convert tab
$("#btn-render").addEventListener("click", async () => {
  const t0 = performance.now();
  try {
    lastRender = await renderToHtml(current);
    const ms = Math.round(performance.now() - t0);
    const size = lastRender.slides.reduce((n, s) => n + s.html.length, 0);
    $("#convert-stats").textContent = `${lastRender.slides.length} slides, ${(size / 1024).toFixed(0)} KB of HTML in ${ms} ms`;
    $<HTMLButtonElement>("#btn-dl-html").disabled = false;
    $<HTMLButtonElement>("#btn-dl-odp").disabled = false;
    const notes = $("#notes");
    notes.textContent = "";
    for (const s of lastRender.slides) {
      const li = document.createElement("li");
      li.textContent = (s.hidden ? "[hidden] " : "") + (s.notes || "(no notes)");
      notes.appendChild(li);
    }
  } catch (e) {
    $("#convert-stats").textContent = e instanceof PptxError ? `${e.code}: ${e.message}` : String(e);
  }
});
$<HTMLButtonElement>("#btn-dl-odp").disabled = false;
$("#btn-dl-html").addEventListener("click", () => {
  if (!lastRender) return;
  const page =
    `<!doctype html><meta charset="utf-8"><title>Presentation</title><style>${BASE_CSS}\n${lastRender.css}\n` +
    `body{background:#ddd;margin:0;padding:20px}.pptxjs .slide{margin:0 auto 20px;box-shadow:0 2px 8px #0003}</style>` +
    `<div class="pptxjs">${lastRender.slides.map((s) => s.html).join("\n")}</div>`;
  download(new Blob([page], { type: "text/html" }), "presentation.html");
});
$("#btn-dl-odp").addEventListener("click", async () => {
  try {
    download(exportOdp(await loadPptx(current)), "presentation.odp");
  } catch (e) {
    log(String((e as Error).message), "error");
  }
});
function download(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob);
  const a = Object.assign(document.createElement("a"), { href: url, download: name });
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ---------------------------------------------------------------- web component
let wcMounted = false;
function mountWebComponent() {
  if (wcMounted) return;
  wcMounted = true;
  defineCustomElements();
  const el = document.createElement("pptx-viewer") as HTMLElement & { source: unknown };
  el.setAttribute("mode", "slideshow");
  el.setAttribute("transition", "slide");
  el.source = current;
  $("#wc-host").appendChild(el);
}

void openSource(SAMPLE_URL);
