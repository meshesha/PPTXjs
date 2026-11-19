import { decodeXmlEntities, encodeHtmlEntities } from "../core/xml";

/**
 * Escapes text for HTML. Values coming from the XML tree are already
 * entity-encoded (see core/xml.ts); decoding first avoids double escaping.
 */
export function escapeHtml(text: unknown): string {
  if (text === undefined || text === null) return "";
  return encodeHtmlEntities(decodeXmlEntities(String(text)));
}

/** Same as {@link escapeHtml}; named for attribute contexts for readability. */
export const escapeAttr = escapeHtml;

const SAFE_URL = /^(?:https?:|mailto:|tel:|ftp:|#)/i;

/**
 * Returns `url` if it uses a safe scheme (http, https, mailto, tel, ftp or an
 * in-page anchor), otherwise `""`. Blocks `javascript:`, `data:`, `vbscript:`
 * and similar hyperlink targets taken from the document (v1 issue #3).
 */
export function sanitizeUrl(url: unknown): string {
  if (typeof url !== "string") return "";
  // decode entities and strip control characters/whitespace browsers ignore
  const decoded = decodeXmlEntities(url).replace(/[\u0000-\u001F\u007F\s]+/g, "");
  if (decoded === "") return "";
  if (SAFE_URL.test(decoded)) return encodeHtmlEntities(decodeXmlEntities(url).trim());
  // relative URLs without a scheme are allowed
  if (!/^[a-z][a-z0-9+.-]*:/i.test(decoded)) return encodeHtmlEntities(decodeXmlEntities(url).trim());
  return "";
}
