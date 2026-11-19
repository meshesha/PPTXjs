/**
 * Custom geometry (`<a:custGeom>`) to SVG path conversion, including the
 * DrawingML shape-guide formula language (ECMA-376 Part 1, 20.1.9.11).
 *
 * Replaces v1's implementation, which crashed on shapes with several
 * `<a:path>` elements (issues #41, #43), dropped single `<a:lnTo>` commands,
 * supported only one `<a:arcTo>` and ignored `<a:quadBezTo>` and guides.
 */
import { asArray } from "../core/xml";

export interface SvgSubPath {
  d: string;
  /** false when the path has `fill="none"` */
  fill: boolean;
  /** false when the path has `stroke="0"` */
  stroke: boolean;
}

type Env = Record<string, number>;

const ANGLE = 60000; // OOXML angles are in 60000ths of a degree
const toRad = (a: number) => (a / ANGLE) * (Math.PI / 180);

/** Built-in guide values available to every shape (20.1.9.11). */
function builtins(w: number, h: number): Env {
  const ss = Math.min(w, h);
  const ls = Math.max(w, h);
  const env: Env = {
    w,
    h,
    l: 0,
    t: 0,
    r: w,
    b: h,
    hc: w / 2,
    vc: h / 2,
    wd2: w / 2,
    wd3: w / 3,
    wd4: w / 4,
    wd5: w / 5,
    wd6: w / 6,
    wd8: w / 8,
    wd10: w / 10,
    wd12: w / 12,
    wd32: w / 32,
    hd2: h / 2,
    hd3: h / 3,
    hd4: h / 4,
    hd5: h / 5,
    hd6: h / 6,
    hd8: h / 8,
    hd10: h / 10,
    hd12: h / 12,
    hd32: h / 32,
    ss,
    ls,
    ssd2: ss / 2,
    ssd4: ss / 4,
    ssd6: ss / 6,
    ssd8: ss / 8,
    ssd16: ss / 16,
    ssd32: ss / 32,
    cd2: 10800000,
    cd4: 5400000,
    cd8: 2700000,
    "3cd4": 16200000,
    "3cd8": 8100000,
    "5cd8": 13500000,
    "7cd8": 18900000,
  };
  return env;
}

function value(token: string, env: Env): number {
  if (token in env) return env[token];
  const n = Number(token);
  return Number.isFinite(n) ? n : 0;
}

/** Evaluates one guide formula such as `"*\/ w adj 100000"`. */
export function evalGuide(fmla: string, env: Env): number {
  const parts = fmla.trim().split(/\s+/);
  const op = parts[0];
  const [x, y, z] = parts.slice(1).map((t) => value(t, env));
  switch (op) {
    case "val":
      return x;
    case "*/":
      return z === 0 ? 0 : (x * y) / z;
    case "+-":
      return x + y - z;
    case "+/":
      return z === 0 ? 0 : (x + y) / z;
    case "?:":
      return x > 0 ? y : z;
    case "abs":
      return Math.abs(x);
    case "at2":
      return (Math.atan2(y, x) * 180 * ANGLE) / Math.PI;
    case "cat2":
      return x * Math.cos(Math.atan2(z, y));
    case "cos":
      return x * Math.cos(toRad(y));
    case "max":
      return Math.max(x, y);
    case "min":
      return Math.min(x, y);
    case "mod":
      return Math.sqrt(x * x + y * y + z * z);
    case "pin":
      return y < x ? x : y > z ? z : y;
    case "sat2":
      return x * Math.sin(Math.atan2(z, y));
    case "sin":
      return x * Math.sin(toRad(y));
    case "sqrt":
      return Math.sqrt(Math.max(0, x));
    case "tan":
      return x * Math.tan(toRad(y));
    default:
      return 0;
  }
}

/** Evaluates `<a:avLst>` and `<a:gdLst>` guides into `env` (in order). */
export function evalGuides(geom: any, env: Env): Env {
  for (const list of [geom?.["a:avLst"], geom?.["a:gdLst"]]) {
    for (const gd of asArray(list?.["a:gd"])) {
      const name = gd?.attrs?.name;
      const fmla = gd?.attrs?.fmla;
      if (typeof name === "string" && typeof fmla === "string") env[name] = evalGuide(fmla, env);
    }
  }
  return env;
}

const COMMANDS = ["a:moveTo", "a:lnTo", "a:cubicBezTo", "a:quadBezTo", "a:arcTo", "a:close"];

/**
 * Converts a `<a:custGeom>` node into SVG sub-paths scaled to `width` x `height`
 * pixels. Commands are re-ordered using the document-order counter kept by the
 * XML parser (`attrs.order`), because the simplified tree groups elements by tag.
 */
export function custGeomToSvgPaths(custGeom: any, width: number, height: number): SvgSubPath[] {
  const out: SvgSubPath[] = [];
  const paths = asArray(custGeom?.["a:pathLst"]?.["a:path"]);
  const fmt = (n: number) => (Number.isFinite(n) ? Math.round(n * 1000) / 1000 : 0);
  for (const path of paths) {
    if (!path || typeof path !== "object") continue;
    const attrs = path.attrs ?? {};
    // path coordinate space; falls back to the shape size (in EMU-less px space)
    const pw = Number(attrs.w) > 0 ? Number(attrs.w) : width;
    const ph = Number(attrs.h) > 0 ? Number(attrs.h) : height;
    const env = evalGuides(custGeom, builtins(pw, ph));
    const sx = pw === 0 ? 0 : width / pw;
    const sy = ph === 0 ? 0 : height / ph;
    const pt = (p: any): [number, number] => [value(String(p?.attrs?.x ?? 0), env), value(String(p?.attrs?.y ?? 0), env)];

    const cmds: Array<{ tag: string; node: any; order: number }> = [];
    for (const tag of COMMANDS) {
      for (const node of asArray(path[tag])) {
        cmds.push({ tag, node, order: Number(node?.attrs?.order ?? 0) });
      }
    }
    cmds.sort((a, b) => a.order - b.order);

    let d = "";
    let cur: [number, number] = [0, 0];
    let start: [number, number] = [0, 0];
    for (const { tag, node } of cmds) {
      switch (tag) {
        case "a:moveTo": {
          cur = start = pt(asArray(node["a:pt"])[0]);
          d += ` M${fmt(cur[0] * sx)},${fmt(cur[1] * sy)}`;
          break;
        }
        case "a:lnTo": {
          cur = pt(asArray(node["a:pt"])[0]);
          d += ` L${fmt(cur[0] * sx)},${fmt(cur[1] * sy)}`;
          break;
        }
        case "a:cubicBezTo": {
          const pts = asArray(node["a:pt"]).map(pt);
          if (pts.length < 3) break;
          d += ` C${pts.map((p) => `${fmt(p[0] * sx)},${fmt(p[1] * sy)}`).join(" ")}`;
          cur = pts[2];
          break;
        }
        case "a:quadBezTo": {
          const pts = asArray(node["a:pt"]).map(pt);
          if (pts.length < 2) break;
          d += ` Q${pts.map((p) => `${fmt(p[0] * sx)},${fmt(p[1] * sy)}`).join(" ")}`;
          cur = pts[1];
          break;
        }
        case "a:arcTo": {
          const a = node.attrs ?? {};
          const wR = value(String(a.wR ?? 0), env);
          const hR = value(String(a.hR ?? 0), env);
          const stAng = value(String(a.stAng ?? 0), env);
          const swAng = value(String(a.swAng ?? 0), env);
          if (wR === 0 || hR === 0) break;
          const st = toRad(stAng);
          const en = toRad(stAng + swAng);
          const cx = cur[0] - wR * Math.cos(st);
          const cy = cur[1] - hR * Math.sin(st);
          const end: [number, number] = [cx + wR * Math.cos(en), cy + hR * Math.sin(en)];
          const large = Math.abs(swAng) > 180 * ANGLE ? 1 : 0;
          const sweep = swAng > 0 ? 1 : 0;
          if (Math.abs(swAng) >= 360 * ANGLE) {
            // full ellipse: split into two halves (SVG cannot draw a closed arc)
            const mid: [number, number] = [cx + wR * Math.cos(st + Math.PI), cy + hR * Math.sin(st + Math.PI)];
            d += ` A${fmt(wR * sx)},${fmt(hR * sy)} 0 0,${sweep} ${fmt(mid[0] * sx)},${fmt(mid[1] * sy)}`;
          }
          d += ` A${fmt(wR * sx)},${fmt(hR * sy)} 0 ${large},${sweep} ${fmt(end[0] * sx)},${fmt(end[1] * sy)}`;
          cur = end;
          break;
        }
        case "a:close":
          d += " z";
          cur = start;
          break;
      }
    }
    if (d) out.push({ d: d.trim(), fill: attrs.fill !== "none", stroke: attrs.stroke !== "0" && attrs.stroke !== "false" });
  }
  return out;
}
