/**
 * Right-to-left (bidi) helpers: RTL languages, first-strong-character
 * detection and the paragraph-direction rule used by the renderer.
 */
import { asArray, decodeXmlEntities, getTextByPathList } from "../core/xml";

/** Text direction policy for paragraphs that do not declare a direction. */
export type TextDirection = "file" | "auto" | "rtl";

/** Primary language subtags written right to left. */
const RTL_LANGS = new Set(["ar", "he", "iw", "fa", "ur", "yi", "ji", "ps", "dv", "ckb", "sd", "ug", "syr", "arc", "nqo", "ks", "ku-arab", "pa-arab", "az-arab"]);

/** True for BCP-47 tags of right-to-left languages (`he-IL`, `ar-EG`, `fa`, `ku-Arab-IQ` ...). */
export function isRtlLanguage(lang: unknown): boolean {
  if (typeof lang !== "string" || lang === "") return false;
  const parts = lang.toLowerCase().split(/[-_]/);
  if (RTL_LANGS.has(parts[0])) return true;
  // explicit Arabic/Hebrew script subtag, e.g. "ku-Arab", "pa-Arab", "yi-Hebr"
  return parts.slice(1).some((p) => p === "arab" || p === "hebr" || p === "syrc" || p === "thaa" || p === "nkoo");
}

const RTL_CHAR = /[֐-ࣿיִ-﷿ﹰ-﻿\u{10800}-\u{10FFF}\u{1E800}-\u{1EFFF}]/u;
const STRONG = /[\p{L}]/u;

/** Direction of the first strong (letter) character, like HTML `dir="auto"`. */
export function firstStrongDirection(text: string): "rtl" | "ltr" | null {
  for (const ch of text) {
    if (RTL_CHAR.test(ch)) return "rtl";
    if (STRONG.test(ch)) return "ltr";
  }
  return null;
}

/** Plain text of an `a:p` node (simplified tree). */
export function paragraphText(pNode: any): string {
  let s = "";
  for (const key of ["a:r", "a:fld"]) {
    for (const r of asArray(pNode?.[key])) {
      const t = r?.["a:t"];
      if (typeof t === "string") s += decodeXmlEntities(t);
    }
  }
  return s;
}

/**
 * Final direction of a paragraph. An explicit `rtl` attribute (paragraph,
 * list style, layout, master; already looked up by the caller) always wins,
 * then the presentation default text style, then the `textDirection` policy.
 */
export function resolveParagraphRtl(pNode: any, explicit: unknown, warpObj: any): boolean {
  if (explicit !== undefined && explicit !== null) return explicit === "1" || explicit === "true";
  const presDefault = getTextByPathList(warpObj?.defaultTextStyle, ["a:lvl1pPr", "attrs", "rtl"]);
  if (presDefault !== undefined) return presDefault === "1" || presDefault === "true";
  const mode: TextDirection = warpObj?.ctx?.textDirection ?? "file";
  if (mode === "rtl") return true;
  if (mode === "auto") return firstStrongDirection(paragraphText(pNode)) === "rtl";
  return false;
}

/** True when `resolveParagraphRtl` would decide without an explicit value (for "inherit" handling). */
export function hasDirectionPolicy(warpObj: any): boolean {
  const mode: TextDirection = warpObj?.ctx?.textDirection ?? "file";
  return mode !== "file" || getTextByPathList(warpObj?.defaultTextStyle, ["a:lvl1pPr", "attrs", "rtl"]) !== undefined;
}
