/**
 * Small, lossless XML DOM used for editing package parts. Unlike the
 * renderer's simplified tree it keeps element order, whitespace, namespaces,
 * comments and processing instructions, so serialising an unmodified part gives
 * back equivalent XML. Text and attribute values are stored *raw* (still
 * entity-encoded); use {@link XElement.text} / {@link XElement.attr} for decoded
 * values.
 */
import { decodeXmlEntities } from "./xml";

export type XNode = XElement | XText;

export class XText {
  readonly type = "text" as const;
  /** raw (encoded) character data, or CDATA content when `cdata` is true */
  constructor(
    public raw: string,
    public cdata = false,
  ) {}
  get value(): string {
    return this.cdata ? this.raw : decodeXmlEntities(this.raw);
  }
}

export class XElement {
  readonly type = "element" as const;
  name: string;
  /** ordered attributes with raw (encoded) values */
  attrs: Array<[string, string]>;
  children: XNode[];
  parent: XElement | null = null;

  constructor(name: string, attrs: Record<string, string | number> | Array<[string, string]> = [], children: XNode[] = []) {
    this.name = name;
    this.attrs = Array.isArray(attrs) ? attrs : Object.entries(attrs).map(([k, v]) => [k, escapeXml(String(v))]);
    this.children = [];
    for (const c of children) this.append(c);
  }

  /** Decoded attribute value. */
  attr(name: string): string | undefined {
    const a = this.attrs.find((x) => x[0] === name);
    return a ? decodeXmlEntities(a[1]) : undefined;
  }

  setAttr(name: string, value: string | number | undefined | null): this {
    const i = this.attrs.findIndex((x) => x[0] === name);
    if (value === undefined || value === null) {
      if (i !== -1) this.attrs.splice(i, 1);
      return this;
    }
    const raw = escapeXml(String(value));
    if (i === -1) this.attrs.push([name, raw]);
    else this.attrs[i][1] = raw;
    return this;
  }

  get elements(): XElement[] {
    return this.children.filter((c): c is XElement => c.type === "element");
  }

  /** First child element with the given qualified name. */
  child(name: string): XElement | undefined {
    return this.elements.find((c) => c.name === name);
  }

  childrenNamed(name: string): XElement[] {
    return this.elements.filter((c) => c.name === name);
  }

  /** Follows a path of child names. */
  path(...names: string[]): XElement | undefined {
    return names.reduce<XElement | undefined>((cur, n) => cur?.child(n), this);
  }

  /** Depth-first search for descendants matching `pred`. */
  findAll(pred: (e: XElement) => boolean, out: XElement[] = []): XElement[] {
    for (const c of this.elements) {
      if (pred(c)) out.push(c);
      c.findAll(pred, out);
    }
    return out;
  }

  find(pred: (e: XElement) => boolean): XElement | undefined {
    for (const c of this.elements) {
      if (pred(c)) return c;
      const r = c.find(pred);
      if (r) return r;
    }
    return undefined;
  }

  /** Returns the child, creating it (appended, or at `index`) when missing. */
  ensure(name: string, index?: number): XElement {
    let c = this.child(name);
    if (!c) {
      c = new XElement(name);
      if (index === undefined) this.append(c);
      else this.insert(c, index);
    }
    return c;
  }

  append(node: XNode): this {
    if (node.type === "element") {
      node.parent?.remove(node);
      node.parent = this;
    }
    this.children.push(node);
    return this;
  }

  insert(node: XNode, index: number): this {
    if (node.type === "element") {
      node.parent?.remove(node);
      node.parent = this;
    }
    this.children.splice(Math.max(0, Math.min(index, this.children.length)), 0, node);
    return this;
  }

  /** Inserts `node` right after the first existing child named in `after` (or first if none). */
  insertAfterAny(node: XElement, after: string[]): this {
    let idx = -1;
    this.children.forEach((c, i) => {
      if (c.type === "element" && after.includes(c.name)) idx = i;
    });
    return this.insert(node, idx + 1);
  }

  remove(node: XNode): this {
    const i = this.children.indexOf(node);
    if (i !== -1) this.children.splice(i, 1);
    if (node.type === "element" && node.parent === this) node.parent = null;
    return this;
  }

  /** Removes this element from its parent. */
  detach(): this {
    this.parent?.remove(this);
    return this;
  }

  /** Decoded concatenated text of all descendants. */
  get text(): string {
    let s = "";
    for (const c of this.children) s += c.type === "text" ? c.value : c.text;
    return s;
  }

  /** Replaces all children with one text node. */
  setText(value: string): this {
    for (const c of this.children) if (c.type === "element") c.parent = null;
    this.children = value === "" ? [] : [new XText(escapeXml(value))];
    return this;
  }

  /** Deep copy. */
  clone(): XElement {
    const e = new XElement(this.name, this.attrs.map(([k, v]) => [k, v] as [string, string]));
    for (const c of this.children) e.append(c.type === "text" ? new XText(c.raw, c.cdata) : c.clone());
    return e;
  }
}

export class XDocument {
  /** text before the root element (XML declaration, comments) */
  prolog: string;
  root: XElement;
  constructor(root: XElement, prolog = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n') {
    this.root = root;
    this.prolog = prolog;
  }
  toString(): string {
    return this.prolog + serialize(this.root);
  }
}

export function escapeXml(s: string): string {
  return s.replace(/[&<>"']/g, (c) => (c === "&" ? "&amp;" : c === "<" ? "&lt;" : c === ">" ? "&gt;" : c === '"' ? "&quot;" : "&apos;"));
}

const NAME_END = /[\s/>=]/;

/** Parses an XML document. Throws on malformed input. */
export function parseXmlDocument(xml: string): XDocument {
  let pos = 0;
  const len = xml.length;
  if (xml.charCodeAt(0) === 0xfeff) pos = 1;
  const stack: XElement[] = [];
  let root: XElement | null = null;
  let prologEnd = -1;

  const fail = (msg: string): never => {
    throw new Error(`XML parse error at ${pos}: ${msg}`);
  };

  while (pos < len) {
    const lt = xml.indexOf("<", pos);
    if (lt === -1) {
      if (stack.length) stack[stack.length - 1].append(new XText(xml.slice(pos)));
      break;
    }
    if (lt > pos && stack.length) stack[stack.length - 1].append(new XText(xml.slice(pos, lt)));
    pos = lt;
    if (xml.startsWith("<?", pos)) {
      const end = xml.indexOf("?>", pos);
      if (end === -1) fail("unterminated processing instruction");
      pos = end + 2;
      continue;
    }
    if (xml.startsWith("<!--", pos)) {
      const end = xml.indexOf("-->", pos);
      if (end === -1) fail("unterminated comment");
      pos = end + 3;
      continue;
    }
    if (xml.startsWith("<![CDATA[", pos)) {
      const end = xml.indexOf("]]>", pos);
      if (end === -1) fail("unterminated CDATA");
      if (stack.length) stack[stack.length - 1].append(new XText(xml.slice(pos + 9, end), true));
      pos = end + 3;
      continue;
    }
    if (xml.startsWith("<!", pos)) {
      const end = xml.indexOf(">", pos);
      pos = end === -1 ? len : end + 1;
      continue;
    }
    if (xml.charCodeAt(pos + 1) === 47 /* / */) {
      const end = xml.indexOf(">", pos);
      if (end === -1) fail("unterminated end tag");
      const name = xml.slice(pos + 2, end).trim();
      const open = stack.pop();
      if (!open || open.name !== name) fail(`mismatched end tag </${name}>`);
      pos = end + 1;
      continue;
    }
    // start tag
    pos++;
    let p = pos;
    while (p < len && !NAME_END.test(xml[p])) p++;
    const name = xml.slice(pos, p);
    if (!name) fail("empty tag name");
    pos = p;
    const attrs: Array<[string, string]> = [];
    let selfClosing = false;
    for (;;) {
      while (pos < len && /\s/.test(xml[pos])) pos++;
      if (pos >= len) fail("unterminated start tag");
      const c = xml[pos];
      if (c === ">") {
        pos++;
        break;
      }
      if (c === "/" && xml[pos + 1] === ">") {
        selfClosing = true;
        pos += 2;
        break;
      }
      let q = pos;
      while (q < len && !NAME_END.test(xml[q])) q++;
      const an = xml.slice(pos, q);
      pos = q;
      while (pos < len && /\s/.test(xml[pos])) pos++;
      if (xml[pos] !== "=") fail(`attribute ${an} without value`);
      pos++;
      while (pos < len && /\s/.test(xml[pos])) pos++;
      const quote = xml[pos];
      if (quote !== '"' && quote !== "'") fail("unquoted attribute value");
      const end = xml.indexOf(quote, pos + 1);
      if (end === -1) fail("unterminated attribute value");
      let raw = xml.slice(pos + 1, end);
      if (quote === "'") raw = raw.replace(/"/g, "&quot;");
      attrs.push([an, raw]);
      pos = end + 1;
    }
    const el = new XElement(name, attrs);
    if (stack.length) stack[stack.length - 1].append(el);
    else if (!root) {
      root = el;
      prologEnd = lt;
    } else fail("multiple root elements");
    if (!selfClosing) stack.push(el);
  }
  if (!root) throw new Error("XML parse error: no root element");
  if (stack.length) throw new Error(`XML parse error: unclosed <${stack[stack.length - 1].name}>`);
  const prolog = xml.slice(xml.charCodeAt(0) === 0xfeff ? 1 : 0, prologEnd);
  return new XDocument(root, prolog);
}

/** Serialises an element (and its subtree) to XML. */
export function serialize(node: XNode): string {
  if (node.type === "text") return node.cdata ? `<![CDATA[${node.raw}]]>` : node.raw;
  let s = "<" + node.name;
  for (const [k, v] of node.attrs) s += ` ${k}="${v}"`;
  if (node.children.length === 0) return s + "/>";
  s += ">";
  for (const c of node.children) s += serialize(c);
  return s + "</" + node.name + ">";
}

/** Convenience element factory: `x("a:off", { x: 0, y: 0 })`. */
export function x(name: string, attrs: Record<string, string | number | undefined> = {}, ...children: Array<XNode | string>): XElement {
  const clean: Record<string, string | number> = {};
  for (const [k, v] of Object.entries(attrs)) if (v !== undefined) clean[k] = v;
  const el = new XElement(name, clean);
  for (const c of children) el.append(typeof c === "string" ? new XText(escapeXml(c)) : c);
  return el;
}
