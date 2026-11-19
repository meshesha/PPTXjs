/**
 * Vue 3 bindings.
 *
 * ```vue
 * <script setup>
 * import { PptxViewer } from "@meshesha/pptxjs/vue";
 * </script>
 * <template>
 *   <PptxViewer src="/deck.pptx" :options="{ mode: 'slideshow' }" @load="onLoad" />
 * </template>
 * ```
 */
import { defineComponent, h, onBeforeUnmount, onMounted, ref, shallowRef, watch, type PropType } from "vue";
import { PptxViewer as Viewer, type ViewerOptions } from "../viewer/viewer";
import { PptxEditor as Editor, type EditorOptions } from "../editor/editor";
import type { PptxSource } from "../source";

export const PptxViewer = defineComponent({
  name: "PptxViewer",
  props: {
    src: { type: [String, Object] as PropType<PptxSource | null>, default: null },
    options: { type: Object as PropType<ViewerOptions>, default: () => ({}) },
  },
  emits: ["load", "error", "slidechange"],
  setup(props, { emit, expose }) {
    const host = ref<HTMLDivElement | null>(null);
    const viewer = shallowRef<Viewer | null>(null);
    const load = (src: PptxSource | null) => {
      if (viewer.value && src) viewer.value.load(src).catch(() => undefined);
    };
    onMounted(() => {
      const v = new Viewer(host.value!, props.options);
      v.on("load", ({ presentation }) => emit("load", presentation));
      v.on("error", ({ error }) => emit("error", error));
      v.on("slidechange", ({ index }) => emit("slidechange", index));
      viewer.value = v;
      load(props.src);
    });
    watch(
      () => props.src,
      (src) => load(src),
    );
    onBeforeUnmount(() => {
      viewer.value?.destroy();
      viewer.value = null;
    });
    expose({ viewer });
    return () => h("div", { ref: host });
  },
});

export const PptxEditor = defineComponent({
  name: "PptxEditor",
  props: {
    src: { type: [String, Object] as PropType<PptxSource | null>, default: null },
    options: { type: Object as PropType<EditorOptions>, default: () => ({}) },
  },
  emits: ["load", "error", "change", "selectionchange"],
  setup(props, { emit, expose }) {
    const host = ref<HTMLDivElement | null>(null);
    const editor = shallowRef<Editor | null>(null);
    const load = (src: PptxSource | null) => {
      if (editor.value && src) editor.value.load(src).catch(() => undefined);
    };
    onMounted(() => {
      const e = new Editor(host.value!, props.options);
      e.on("load", ({ slides }) => emit("load", slides));
      e.on("error", ({ error }) => emit("error", error));
      e.on("change", ({ label }) => emit("change", label));
      e.on("selectionchange", ({ selection }) => emit("selectionchange", selection));
      editor.value = e;
      load(props.src);
    });
    watch(
      () => props.src,
      (src) => load(src),
    );
    onBeforeUnmount(() => {
      editor.value?.destroy();
      editor.value = null;
    });
    expose({ editor });
    return () => h("div", { ref: host, style: { height: "600px" } });
  },
});

/** Vue plugin: `app.use(PptxPlugin)` registers <PptxViewer> and <PptxEditor> globally. */
export const PptxPlugin = {
  install(app: { component: (name: string, c: unknown) => unknown }) {
    app.component("PptxViewer", PptxViewer);
    app.component("PptxEditor", PptxEditor);
  },
};

export type { ViewerOptions, EditorOptions };
