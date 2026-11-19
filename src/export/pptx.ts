/**
 * Export back to .pptx. Because every edit is applied to the original OOXML
 * parts, this is a lossless re-zip of the package (optionally dropping media
 * that is no longer referenced after deletions).
 */
import { PptxPackage, normalizePath } from "../core/package";
import { PresentationDocument } from "../model/document";

export const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

export interface PptxExportOptions {
  /** Remove media parts that no relationship points to anymore (default true). */
  pruneUnusedMedia?: boolean;
}

function referencedParts(pkg: PptxPackage): Set<string> {
  const used = new Set<string>();
  for (const path of pkg.list()) {
    if (!path.endsWith(".rels")) continue;
    const text = pkg.text(path) ?? "";
    const source = path.replace(/_rels\/([^/]+)\.rels$/, "$1");
    const dir = source.split("/").slice(0, -1).join("/");
    for (const m of text.matchAll(/<Relationship\b[^>]*>/g)) {
      const tag = m[0];
      if (/TargetMode="External"/.test(tag)) continue;
      const t = /Target="([^"]*)"/.exec(tag)?.[1];
      if (!t) continue;
      used.add(t.startsWith("/") ? normalizePath(t) : normalizePath(dir + "/" + t));
    }
  }
  return used;
}

/** Serialises a presentation to .pptx bytes. */
export function exportPptxBytes(input: PresentationDocument | PptxPackage, options: PptxExportOptions = {}): Uint8Array {
  const pkg = input instanceof PresentationDocument ? input.pkg : input;
  if (options.pruneUnusedMedia !== false) {
    const used = referencedParts(pkg);
    const copy = pkg.clone();
    for (const p of copy.list()) {
      if (/^ppt\/media\//.test(p) && !used.has(p)) copy.restore(p, null);
    }
    return copy.toZip();
  }
  return pkg.toZip();
}

/** Serialises a presentation to a .pptx `Blob`. */
export function exportPptx(input: PresentationDocument | PptxPackage, options: PptxExportOptions = {}): Blob {
  const bytes = exportPptxBytes(input, options);
  return new Blob([bytes as BlobPart], { type: PPTX_MIME });
}
