/**
 * React bindings.
 *
 * ```tsx
 * import { PptxViewer } from "@meshesha/pptxjs/react";
 * <PptxViewer src="/deck.pptx" options={{ mode: "slideshow" }} onLoad={(p) => console.log(p.slides.length)} />
 * ```
 */
import { createElement, forwardRef, useEffect, useImperativeHandle, useRef, useState, type CSSProperties, type ForwardedRef } from "react";
import { PptxViewer as Viewer, type ViewerOptions } from "../viewer/viewer";
import { PptxEditor as Editor, type EditorOptions, type ShapeRef } from "../editor/editor";
import type { RenderedPresentation } from "../render/presentation";
import type { PptxError } from "../errors";
import type { PptxSource } from "../source";

export interface PptxViewerProps {
  /** URL, File/Blob, ArrayBuffer or Uint8Array. Changing it reloads. */
  src?: PptxSource | null;
  /** Viewer options; only read when the component mounts (use the ref to change them later). */
  options?: ViewerOptions;
  className?: string;
  style?: CSSProperties;
  onLoad?: (presentation: RenderedPresentation) => void;
  onError?: (error: PptxError) => void;
  onSlideChange?: (index: number) => void;
}

/** Renders a presentation. The ref exposes the underlying {@link Viewer} instance. */
export const PptxViewer = forwardRef(function PptxViewer(props: PptxViewerProps, ref: ForwardedRef<Viewer | null>) {
  const host = useRef<HTMLDivElement>(null);
  const [viewer, setViewer] = useState<Viewer | null>(null);
  const cb = useRef(props);
  cb.current = props;

  useEffect(() => {
    const v = new Viewer(host.current!, props.options);
    const offs = [
      v.on("load", ({ presentation }) => cb.current.onLoad?.(presentation)),
      v.on("error", ({ error }) => cb.current.onError?.(error)),
      v.on("slidechange", ({ index }) => cb.current.onSlideChange?.(index)),
    ];
    setViewer(v);
    return () => {
      offs.forEach((off) => off());
      v.destroy();
      setViewer(null);
    };
    // options are applied on mount only
  }, []);

  useEffect(() => {
    if (viewer && props.src) viewer.load(props.src).catch(() => undefined /* reported via onError */);
  }, [viewer, props.src]);

  useImperativeHandle(ref, () => viewer as Viewer, [viewer]);
  return createElement("div", { ref: host, className: props.className, style: props.style });
});

export interface PptxEditorProps {
  src?: PptxSource | null;
  options?: EditorOptions;
  className?: string;
  style?: CSSProperties;
  onLoad?: (slides: number) => void;
  onError?: (error: PptxError) => void;
  onChange?: (label: string) => void;
  onSelectionChange?: (selection: ShapeRef | null) => void;
}

/** Presentation editor. The ref exposes the underlying {@link Editor} instance (export, undo ...). */
export const PptxEditor = forwardRef(function PptxEditor(props: PptxEditorProps, ref: ForwardedRef<Editor | null>) {
  const host = useRef<HTMLDivElement>(null);
  const [editor, setEditor] = useState<Editor | null>(null);
  const cb = useRef(props);
  cb.current = props;

  useEffect(() => {
    const e = new Editor(host.current!, props.options);
    const offs = [
      e.on("load", ({ slides }) => cb.current.onLoad?.(slides)),
      e.on("error", ({ error }) => cb.current.onError?.(error)),
      e.on("change", ({ label }) => cb.current.onChange?.(label)),
      e.on("selectionchange", ({ selection }) => cb.current.onSelectionChange?.(selection)),
    ];
    setEditor(e);
    return () => {
      offs.forEach((off) => off());
      e.destroy();
      setEditor(null);
    };
  }, []);

  useEffect(() => {
    if (editor && props.src) editor.load(props.src).catch(() => undefined);
  }, [editor, props.src]);

  useImperativeHandle(ref, () => editor as Editor, [editor]);
  return createElement("div", { ref: host, className: props.className, style: { height: 600, ...props.style } });
});

export type { ViewerOptions, EditorOptions };
