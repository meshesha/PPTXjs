import { PptxError } from "./errors";

/** Anything a presentation can be loaded from. */
export type PptxSource = string | URL | Blob | ArrayBuffer | ArrayBufferView;

export interface ReadSourceOptions {
  signal?: AbortSignal;
  fetchOptions?: RequestInit;
}

/** Reads a {@link PptxSource} into bytes. */
export async function readSource(source: PptxSource, opts: ReadSourceOptions = {}): Promise<Uint8Array> {
  if (source instanceof Uint8Array) return source;
  if (source instanceof ArrayBuffer) return new Uint8Array(source);
  if (ArrayBuffer.isView(source)) return new Uint8Array(source.buffer, source.byteOffset, source.byteLength);
  if (typeof Blob !== "undefined" && source instanceof Blob) {
    if (typeof source.arrayBuffer === "function") return new Uint8Array(await source.arrayBuffer());
    // older engines / jsdom: FileReader fallback
    return new Uint8Array(
      await new Promise<ArrayBuffer>((resolve, reject) => {
        const r = new FileReader();
        r.onload = () => resolve(r.result as ArrayBuffer);
        r.onerror = () => reject(new PptxError("INVALID_FILE", "Could not read the file", r.error));
        r.readAsArrayBuffer(source);
      }),
    );
  }
  if (typeof source === "string" || source instanceof URL) {
    const url = String(source);
    if (url.trim() === "") throw new PptxError("FETCH_FAILED", "Empty URL");
    if (/^\s*(javascript|vbscript):/i.test(url)) throw new PptxError("FETCH_FAILED", "Unsupported URL scheme");
    let res: Response;
    try {
      res = await fetch(url, { ...opts.fetchOptions, signal: opts.signal });
    } catch (e) {
      if ((e as Error)?.name === "AbortError") throw new PptxError("ABORTED", "Loading was aborted", e);
      throw new PptxError("FETCH_FAILED", `Could not load ${url}: ${(e as Error)?.message ?? e}`, e);
    }
    if (!res.ok) throw new PptxError("FETCH_FAILED", `Could not load ${url}: HTTP ${res.status}`);
    return new Uint8Array(await res.arrayBuffer());
  }
  throw new PptxError("INVALID_FILE", "Unsupported source type");
}
