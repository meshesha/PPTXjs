import { describe, expect, it } from "vitest";
import {
  asArray,
  decodeXmlEntities,
  eachElement,
  getTextByPathList,
  parseXml,
  parseXmlSimplified,
  setTextByPathList,
} from "../../src/core/xml";

describe("parseXmlSimplified", () => {
  it("builds the simplified tree with attrs and document order", () => {
    const t = parseXmlSimplified(`<?xml version="1.0"?><a:root x="1"><a:b v="1"/><a:c>text</a:c><a:b v="2"/></a:root>`);
    const root = t["a:root"];
    expect(root.attrs.x).toBe("1");
    expect(Array.isArray(root["a:b"])).toBe(true);
    expect(root["a:b"][1].attrs.v).toBe("2");
    expect(root["a:c"]).toBe("text");
    // post-order counter: second a:b comes after first
    expect(root["a:b"][1].attrs.order).toBeGreaterThan(root["a:b"][0].attrs.order);
  });

  it("escapes text and attributes so they are HTML-safe", () => {
    const t = parseXmlSimplified(`<r a="x' onmouseover='alert(1)"><t>&lt;img src=x onerror=alert(1)&gt;</t></r>`);
    expect(t.r.attrs.a).toBe("x&#39; onmouseover=&#39;alert(1)");
    expect(t.r.t).toBe("&lt;img src=x onerror=alert(1)&gt;");
    expect(t.r.t).not.toContain("<");
  });

  it("decodes numeric entities before re-encoding", () => {
    const t = parseXmlSimplified(`<r><t>A&#38;B &#x3C;</t></r>`);
    expect(t.r.t).toBe("A&amp;B &lt;");
    expect(decodeXmlEntities(t.r.t)).toBe("A&B <");
  });

  it("keeps CDATA as escaped text", () => {
    const t = parseXmlSimplified(`<r><t><![CDATA[<script>alert(1)</script>]]></t></r>`);
    expect(t.r.t).toBe("&lt;script&gt;alert(1)&lt;/script&gt;");
  });

  it("never pollutes Object.prototype (issues #22, #28, #30)", () => {
    parseXmlSimplified(`<r><__proto__ polluted="yes"/><constructor/><x __proto__="1" constructor="2"/></r>`);
    expect(({} as any).polluted).toBeUndefined();
    expect(Object.prototype.hasOwnProperty.call(Object.prototype, "set")).toBe(false);
    const obj: any = {};
    setTextByPathList(obj, ["__proto__", "polluted"], "yes");
    expect(({} as any).polluted).toBeUndefined();
  });

  it("resolves mc:AlternateContent inside elements (issue #39)", () => {
    const t = parseXmlSimplified(
      `<p:pic><mc:AlternateContent><mc:Choice Requires="ma"><p:blipFill><a:blip r:embed="rId2"/></p:blipFill></mc:Choice>` +
        `<mc:Fallback><p:blipFill><a:blip r:embed="rId3"/></p:blipFill></mc:Fallback></mc:AlternateContent></p:pic>`,
    );
    expect(t["p:pic"]["p:blipFill"]["a:blip"].attrs["r:embed"]).toBe("rId3");
    expect(t["p:pic"]["mc:AlternateContent"]).toBeUndefined();
  });

  it("keeps mc:AlternateContent in shape trees for the renderer", () => {
    const t = parseXmlSimplified(`<p:spTree><p:nvGrpSpPr/><mc:AlternateContent><mc:Fallback><p:sp/></mc:Fallback></mc:AlternateContent></p:spTree>`);
    expect(t["p:spTree"]["mc:AlternateContent"]).toBeDefined();
  });

  it("ignores comments and doctype", () => {
    const raw = parseXml(`<!DOCTYPE x><!-- c --><r>1</r>`);
    expect(raw).toHaveLength(1);
  });
});

describe("tree helpers", () => {
  it("getTextByPathList tolerates null and missing nodes (issue #44)", () => {
    expect(getTextByPathList(null, ["a"])).toBeUndefined();
    expect(getTextByPathList({ a: null }, ["a", "b"])).toBeUndefined();
    expect(getTextByPathList({ a: { b: 1 } }, ["a", "b"])).toBe(1);
    expect(() => getTextByPathList({}, "a" as any)).toThrow();
  });

  it("setTextByPathList creates missing parts only", () => {
    const o: any = { a: { b: 1 } };
    setTextByPathList(o, ["a", "b"], 2);
    setTextByPathList(o, ["a", "c", "d"], 3);
    expect(o).toEqual({ a: { b: 1, c: { d: 3 } } });
  });

  it("eachElement and asArray", () => {
    expect(eachElement([1, 2], (n) => String(n))).toBe("12");
    expect(eachElement(5, (n) => String(n))).toBe("5");
    expect(asArray(undefined)).toEqual([]);
    expect(asArray(1)).toEqual([1]);
  });
});
