/**
 * Access to the OPC (zip) package of a .pptx file.
 *
 * The whole archive is inflated once, up front, with fflate. Every later read is
 * synchronous, which keeps the (large, synchronous) renderer simple and lets us
 * enforce size limits against zip bombs before anything is rendered.
 */
import { unzipSync, zipSync, strFromU8, strToU8 } from "fflate";
import { parseXmlSimplified, type XmlNode } from "./xml";
import { PptxError } from "../errors";

export interface PackageLimits {
  /** Maximum total uncompressed size in bytes (default 512 MiB). */
  maxUncompressedSize?: number;
  /** Maximum number of entries in the archive (default 20 000). */
  maxEntries?: number;
}

export const DEFAULT_LIMITS: Required<PackageLimits> = {
  maxUncompressedSize: 512 * 1024 * 1024,
  maxEntries: 20000,
};

const ZIP_MAGIC = [0x50, 0x4b]; // "PK"

export class PptxPackage {
  private readonly files: Map<string, Uint8Array>;
  private readonly xmlCache = new Map<string, XmlNode | null>();
  private readonly textCache = new Map<string, string | null>();

  private constructor(files: Map<string, Uint8Array>) {
    this.files = files;
  }

  /** Inflates a .pptx archive. Throws {@link PptxError} for invalid input. */
  static fromBytes(input: ArrayBuffer | Uint8Array, limits: PackageLimits = {}): PptxPackage {
    const lim = { ...DEFAULT_LIMITS, ...limits };
    const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
    if (bytes.byteLength < 22 || bytes[0] !== ZIP_MAGIC[0] || bytes[1] !== ZIP_MAGIC[1]) {
      throw new PptxError("INVALID_FILE", "The file is not a valid .pptx (zip) archive.");
    }
    let total = 0;
    let count = 0;
    let entries: Record<string, Uint8Array>;
    try {
      entries = unzipSync(bytes, {
        filter(file) {
          count++;
          total += file.originalSize;
          if (count > lim.maxEntries) {
            throw new PptxError("LIMIT_EXCEEDED", `The archive has more than ${lim.maxEntries} entries.`);
          }
          if (total > lim.maxUncompressedSize) {
            throw new PptxError("LIMIT_EXCEEDED", `The archive expands to more than ${lim.maxUncompressedSize} bytes.`);
          }
          return true;
        },
      });
    } catch (e) {
      if (e instanceof PptxError) throw e;
      throw new PptxError("INVALID_FILE", "The file is not a valid .pptx (zip) archive: " + (e as Error).message, e);
    }
    const map = new Map<string, Uint8Array>();
    for (const [name, data] of Object.entries(entries)) map.set(normalizePath(name), data);
    if (!map.has("[Content_Types].xml")) {
      throw new PptxError("INVALID_FILE", "The archive is not an Office Open XML package ([Content_Types].xml missing).");
    }
    return new PptxPackage(map);
  }

  has(path: string): boolean {
    return this.files.has(normalizePath(path));
  }

  /** Names of all parts in the package. */
  list(): string[] {
    return [...this.files.keys()];
  }

  bytes(path: string): Uint8Array | null {
    return this.files.get(normalizePath(path)) ?? null;
  }

  text(path: string): string | null {
    const key = normalizePath(path);
    if (this.textCache.has(key)) return this.textCache.get(key)!;
    const data = this.files.get(key);
    let text: string | null = null;
    if (data) {
      text = strFromU8(data);
      if (text.charCodeAt(0) === 0xfeff) text = text.slice(1); // BOM
    }
    this.textCache.set(key, text);
    return text;
  }

  /**
   * Parses a part into the simplified XML tree, or returns `null` when the part
   * is missing or malformed. Results are cached per package. Callers must not
   * mutate the returned tree (v1's renderer only ever adds missing defaults).
   */
  xml(path: string): XmlNode | null {
    const key = normalizePath(path);
    if (this.xmlCache.has(key)) return this.xmlCache.get(key)!;
    const text = this.text(key);
    let tree: XmlNode | null = null;
    if (text !== null) {
      try {
        tree = parseXmlSimplified(text);
      } catch {
        tree = null;
      }
    }
    this.xmlCache.set(key, tree);
    return tree;
  }

  /** Uncached parse: returns a fresh tree that the caller may mutate. */
  xmlFresh(path: string): XmlNode | null {
    const text = this.text(path);
    if (text === null) return null;
    try {
      return parseXmlSimplified(text);
    } catch {
      return null;
    }
  }

  /** Raw zip entries (used by the exporter to rebuild the package). */
  entries(): Map<string, Uint8Array> {
    return this.files;
  }

  // ------------------------------------------------------------- mutation
  private writeHooks: Array<(path: string, before: Uint8Array | null) => void> = [];

  /** Registers a hook called before a part is written or deleted (used for undo). */
  onWrite(hook: (path: string, before: Uint8Array | null) => void): () => void {
    this.writeHooks.push(hook);
    return () => {
      this.writeHooks = this.writeHooks.filter((h) => h !== hook);
    };
  }

  private invalidate(key: string) {
    this.xmlCache.delete(key);
    this.textCache.delete(key);
  }

  /** Replaces (or adds) a part. */
  setBytes(path: string, data: Uint8Array): void {
    const key = normalizePath(path);
    for (const h of this.writeHooks) h(key, this.files.get(key) ?? null);
    this.files.set(key, data);
    this.invalidate(key);
  }

  /** Replaces (or adds) a text part (UTF-8). */
  setText(path: string, text: string): void {
    this.setBytes(path, strToU8(text));
  }

  /** Removes a part. */
  delete(path: string): void {
    const key = normalizePath(path);
    if (!this.files.has(key)) return;
    for (const h of this.writeHooks) h(key, this.files.get(key) ?? null);
    this.files.delete(key);
    this.invalidate(key);
  }

  /** Restores a part to given bytes (or deletes it when `null`) without calling hooks. */
  restore(path: string, data: Uint8Array | null): void {
    const key = normalizePath(path);
    if (data === null) this.files.delete(key);
    else this.files.set(key, data);
    this.invalidate(key);
  }

  /** Creates a deep copy of the package. */
  clone(): PptxPackage {
    const map = new Map<string, Uint8Array>();
    for (const [k, v] of this.files) map.set(k, v.slice());
    return new PptxPackage(map);
  }

  /**
   * Serialises the package to a zip archive. `[Content_Types].xml` is written
   * first as recommended by the OPC specification.
   */
  toZip(): Uint8Array {
    const entries: Record<string, Uint8Array | [Uint8Array, { level: 0 }]> = {};
    const ct = this.files.get("[Content_Types].xml");
    if (ct) entries["[Content_Types].xml"] = ct;
    for (const [k, v] of this.files) {
      if (k === "[Content_Types].xml") continue;
      // already-compressed media is stored, everything else deflated
      entries[k] = /\.(png|jpe?g|gif|mp4|m4v|webm|mp3|m4a|wav|zip|emf|wmf)$/i.test(k) ? [v, { level: 0 }] : v;
    }
    return zipSync(entries, { level: 6 });
  }
}

/** Strips leading "/" and resolves "." / ".." segments inside the package. */
export function normalizePath(path: string): string {
  const out: string[] = [];
  for (const seg of path.replace(/\\/g, "/").split("/")) {
    if (seg === "" || seg === ".") continue;
    if (seg === "..") out.pop();
    else out.push(seg);
  }
  return out.join("/");
}

/** Resolves a relationship target relative to the part that owns the .rels file. */
export function resolveTarget(sourcePart: string, target: string): string {
  if (target.startsWith("/")) return normalizePath(target);
  const dir = sourcePart.split("/").slice(0, -1).join("/");
  return normalizePath(dir + "/" + target);
}

export function getMimeType(ext: string | undefined): string {
  switch ((ext || "").toLowerCase()) {
    case "jpg":
    case "jpeg":
      return "image/jpeg";
    case "png":
      return "image/png";
    case "gif":
      return "image/gif";
    case "bmp":
      return "image/bmp";
    case "webp":
      return "image/webp";
    case "emf":
      return "image/x-emf"; // no native browser support
    case "wmf":
      return "image/x-wmf"; // no native browser support
    case "svg":
      return "image/svg+xml";
    case "tif":
    case "tiff":
      return "image/tiff";
    case "mp4":
    case "m4v":
      return "video/mp4";
    case "webm":
      return "video/webm";
    case "ogg":
    case "ogv":
      return "video/ogg";
    case "avi":
      return "video/avi";
    case "mpg":
    case "mpeg":
      return "video/mpeg";
    case "wmv":
      return "video/x-ms-wmv";
    case "mov":
      return "video/quicktime";
    case "mp3":
      return "audio/mpeg";
    case "m4a":
      return "audio/mp4";
    case "wav":
      return "audio/wav";
    case "oga":
      return "audio/ogg";
    default:
      return "";
  }
}

export function extractFileExtension(filename: string): string {
  const i = filename.lastIndexOf(".");
  return i === -1 ? "" : filename.slice(i + 1);
}

/** True for absolute http(s) URLs (external media links). */
export function isVideoLink(target: string): boolean {
  return /^https?:\/\//i.test(target);
}
export { isVideoLink as IsVideoLink };
