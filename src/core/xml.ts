/**
 * Minimal XML parser producing the "simplified" object tree that the renderer
 * works on. It is a TypeScript port of the customised tXml (MIT, Tobias Nickel)
 * that PPTXjs v1 embedded, with these deliberate changes:
 *
 *  - **Output is HTML-safe.** Every text node and attribute value is decoded and
 *    re-encoded so that `& < > " '` only ever appear as entities. The renderer
 *    builds HTML by string concatenation; this guarantees document content can
 *    never break out of a text node or a quoted attribute (XSS hardening).
 *  - **No prototype pollution.** Elements or attributes named `__proto__`,
 *    `constructor` or `prototype` are dropped, and nothing is ever written to
 *    `Object.prototype` (v1 issues #22, #28, #30).
 *  - CDATA sections are kept as (escaped) text instead of being dropped.
 *
 * Simplified tree shape (unchanged from v1 so the ported renderer keeps working):
 *   - a child element appears as `node[tagName]`; repeated tags become arrays
 *   - attributes live in `node.attrs`, plus a document-order counter `attrs.order`
 *   - an element whose only child is text simplifies to that string
 */

/** A node in the simplified XML tree. */
export type XmlNode = { [key: string]: any; attrs?: XmlAttrs };
export type XmlAttrs = { [name: string]: any; order?: number };

interface RawElement {
  tagName: string;
  attributes?: Record<string, string | null>;
  children?: Array<RawElement | string>;
}

const FORBIDDEN_KEYS = new Set(["__proto__", "constructor", "prototype"]);

// ---------------------------------------------------------------------------
// entity handling
// ---------------------------------------------------------------------------
const NAMED_ENTITIES: Record<string, string> = { lt: "<", gt: ">", amp: "&", quot: '"', apos: "'" };

export function decodeXmlEntities(s: string): string {
  if (s.indexOf("&") === -1) return s;
  return s.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z]+);/g, (m, ent: string) => {
    if (ent[0] === "#") {
      const code = ent[1] === "x" || ent[1] === "X" ? parseInt(ent.slice(2), 16) : parseInt(ent.slice(1), 10);
      if (!Number.isFinite(code) || code < 0 || code > 0x10ffff) return m;
      try {
        return String.fromCodePoint(code);
      } catch {
        return m;
      }
    }
    return NAMED_ENTITIES[ent] ?? m;
  });
}

const ESCAPE_RE = /[&<>"']/g;
const ESCAPE_MAP: Record<string, string> = { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" };

/** Escapes a plain string for safe use in HTML text or quoted attributes. */
export function encodeHtmlEntities(s: string): string {
  return s.replace(ESCAPE_RE, (c) => ESCAPE_MAP[c]);
}

/** Normalises raw XML character data into HTML-safe text. */
function safeText(raw: string): string {
  if (!/[&<>"']/.test(raw)) return raw;
  return encodeHtmlEntities(decodeXmlEntities(raw));
}

// ---------------------------------------------------------------------------
// parser
// ---------------------------------------------------------------------------
const LT = 60, // <
  GT = 62, // >
  MINUS = 45, // -
  SLASH = 47, // /
  BANG = 33, // !
  SQUOTE = 39, // '
  DQUOTE = 34, // "
  LBRACKET = 91; // [
const NAME_SPACER = "\r\n\t >/= ";
const VOID_TAGS = ["img", "br", "input", "meta", "link"];

/** Parses an XML string into a raw element list (tXml format). */
export function parseXml(xml: string): Array<RawElement | string> {
  let pos = 0;

  function parseChildren(): Array<RawElement | string> {
    const children: Array<RawElement | string> = [];
    while (pos < xml.length) {
      if (xml.charCodeAt(pos) === LT) {
        const next = xml.charCodeAt(pos + 1);
        if (next === SLASH) {
          // closing tag
          pos = xml.indexOf(">", pos);
          if (pos === -1) pos = xml.length;
          else pos += 1;
          return children;
        }
        if (next === BANG) {
          if (xml.charCodeAt(pos + 2) === MINUS) {
            // comment
            const end = xml.indexOf("-->", pos + 4);
            pos = end === -1 ? xml.length : end + 3;
          } else if (xml.charCodeAt(pos + 2) === LBRACKET && xml.startsWith("<![CDATA[", pos)) {
            const end = xml.indexOf("]]>", pos + 9);
            const content = xml.slice(pos + 9, end === -1 ? xml.length : end);
            if (content.length) children.push(encodeHtmlEntities(content));
            pos = end === -1 ? xml.length : end + 3;
          } else {
            // <!DOCTYPE ...> and friends
            pos += 2;
            while (pos < xml.length && xml.charCodeAt(pos) !== GT) pos++;
            pos++;
          }
          continue;
        }
        children.push(parseElement());
      } else {
        const text = parseText();
        if (text.trim().length > 0) children.push(safeText(text));
        pos++;
      }
    }
    return children;
  }

  function parseText(): string {
    const start = pos;
    pos = xml.indexOf("<", pos) - 1;
    if (pos === -2) pos = xml.length;
    return xml.slice(start, pos + 1);
  }

  function parseName(): string {
    const start = pos;
    while (pos < xml.length && NAME_SPACER.indexOf(xml[pos]) === -1) pos++;
    return xml.slice(start, pos);
  }

  function parseQuoted(): string {
    const quote = xml[pos];
    const start = ++pos;
    pos = xml.indexOf(quote, start);
    if (pos === -1) pos = xml.length;
    return xml.slice(start, pos);
  }

  function isNameStart(c: number): boolean {
    return (c > 64 && c < 91) || (c > 96 && c < 123) || c === 95; // A-Z a-z _
  }

  function parseElement(): RawElement {
    pos++; // skip <
    const el: RawElement = { tagName: parseName() };
    while (pos < xml.length && xml.charCodeAt(pos) !== GT) {
      const c = xml.charCodeAt(pos);
      if (isNameStart(c)) {
        const name = parseName();
        let code = xml.charCodeAt(pos);
        while (pos < xml.length && code !== SQUOTE && code !== DQUOTE && !isNameStart(code) && code !== GT) {
          pos++;
          code = xml.charCodeAt(pos);
        }
        el.attributes ??= {};
        let value: string | null;
        if (code === SQUOTE || code === DQUOTE) {
          value = safeText(parseQuoted());
        } else {
          value = null;
          pos--;
        }
        if (!FORBIDDEN_KEYS.has(name)) el.attributes[name] = value;
      }
      pos++;
    }
    if (xml.charCodeAt(pos - 1) !== SLASH) {
      if (VOID_TAGS.indexOf(el.tagName) === -1) {
        pos++;
        el.children = parseChildren();
      }
    } else {
      pos++;
    }
    return el;
  }

  return parseChildren();
}

// ---------------------------------------------------------------------------
// simplify
// ---------------------------------------------------------------------------
let orderCounter = 1;

/** Resets the document-order counter (tests only). */
export function resetXmlOrder(): void {
  orderCounter = 1;
}

/**
 * Converts the raw tXml element list into the simplified tree. Element order is
 * recorded in `attrs.order` (post-order, global across files exactly like v1,
 * because the renderer uses it as z-index).
 */
export function simplify(children: Array<RawElement | string> | undefined): any {
  const out: XmlNode = {};
  if (children === undefined) return out;
  if (children.length === 1 && typeof children[0] === "string") return children[0];
  for (const child of children) {
    if (typeof child !== "object") continue;
    const tag = child.tagName;
    const value = simplify(child.children || []);
    if (!FORBIDDEN_KEYS.has(tag)) {
      if (!Object.prototype.hasOwnProperty.call(out, tag)) out[tag] = [];
      out[tag].push(value);
    }
    if (typeof value === "object") {
      if (child.attributes) value.attrs = child.attributes;
      if (value.attrs === undefined) value.attrs = { order: orderCounter };
      else value.attrs.order = orderCounter;
    }
    orderCounter++;
  }
  for (const key of Object.keys(out)) {
    if (out[key].length === 1) out[key] = out[key][0];
  }
  return out;
}

/** Parses XML text and returns the simplified tree (the `<?xml?>` wrapper removed). */
export function parseXmlSimplified(xml: string): XmlNode {
  let tree = simplify(parseXml(xml));
  if (tree && typeof tree === "object" && tree["?xml"] !== undefined) tree = tree["?xml"];
  if (xml.indexOf("mc:AlternateContent") !== -1) resolveAlternateContent(tree);
  return tree;
}

/**
 * Markup-compatibility: replaces `mc:AlternateContent` blocks by the content of
 * their `mc:Fallback` (or first `mc:Choice`) inside ordinary elements, e.g. a
 * `p:blipFill` written by PowerPoint for Mac (issue #39). Shape trees keep the
 * block because the slide renderer handles it there as a group.
 */
export function resolveAlternateContent(node: any): void {
  if (!node || typeof node !== "object") return;
  if (Array.isArray(node)) {
    for (const n of node) resolveAlternateContent(n);
    return;
  }
  const alt = node["mc:AlternateContent"];
  const isShapeTree = node["p:nvGrpSpPr"] !== undefined || node["p:grpSpPr"] !== undefined;
  if (alt !== undefined && !isShapeTree) {
    for (const block of asArray(alt)) {
      const chosen = block?.["mc:Fallback"] ?? asArray(block?.["mc:Choice"])[0];
      if (!chosen || typeof chosen !== "object") continue;
      for (const key of Object.keys(chosen)) {
        if (key === "attrs" || FORBIDDEN_KEYS.has(key)) continue;
        if (node[key] === undefined) node[key] = chosen[key];
      }
    }
    delete node["mc:AlternateContent"];
  }
  for (const key of Object.keys(node)) {
    if (key !== "attrs") resolveAlternateContent(node[key]);
  }
}

// ---------------------------------------------------------------------------
// tree helpers (were inner functions of v1)
// ---------------------------------------------------------------------------

/** Walks `path` from `node`; returns `undefined` as soon as a step is missing. */
export function getTextByPathList(node: any, path: readonly (string | number)[]): any {
  if (!Array.isArray(path)) throw new TypeError("getTextByPathList: path must be an array");
  if (node === undefined || node === null) return undefined;
  let cur = node;
  for (const key of path) {
    if (cur === undefined || cur === null) return undefined;
    cur = cur[key];
    if (cur === undefined) return undefined;
  }
  return cur;
}

export function getTextByPathStr(node: any, pathStr: string): any {
  return getTextByPathList(node, pathStr.trim().split(/\s+/));
}

/**
 * Creates the path inside `node` (if missing) and sets the leaf to `value`
 * when it does not exist yet. Replaces v1's `Object.prototype.set` hack.
 */
export function setTextByPathList(node: any, path: readonly string[], value: any): void {
  if (!Array.isArray(path)) throw new TypeError("setTextByPathList: path must be an array");
  if (node === undefined || node === null) return;
  let obj = node;
  for (let i = 0; i < path.length; i++) {
    const p = path[i];
    if (FORBIDDEN_KEYS.has(p)) return;
    if (obj[p] === undefined) obj[p] = i === path.length - 1 ? value : {};
    obj = obj[p];
    if (obj === null || typeof obj !== "object") return;
  }
}

/** Calls `fn` for one node or for every node of an array and concatenates the results. */
export function eachElement(node: any, fn: (n: any, i: number) => string): string | undefined {
  if (node === undefined) return undefined;
  if (Array.isArray(node)) {
    let result = "";
    for (let i = 0; i < node.length; i++) result += fn(node[i], i);
    return result;
  }
  return fn(node, 0);
}

/** Normalises "one or many" simplified children into an array. */
export function asArray<T = any>(node: T | T[] | undefined | null): T[] {
  if (node === undefined || node === null) return [];
  return Array.isArray(node) ? node : [node];
}
