/**
 * SVG geometry -> ODF `draw:enhanced-path`.
 *
 * Office applications (PowerPoint in particular) need explicit geometry in
 * `draw:enhanced-path`; ODF's named shape types are not understood everywhere.
 * The ODP exporter therefore takes the geometry the HTML renderer draws (SVG)
 * and converts it: all commands become absolute M/L/C/Q/Z, elliptical arcs are
 * approximated with cubic Béziers (SVG 1.1 implementation notes F.6.5).
 */

export type PathCmd = { c: "M" | "L" | "C" | "Q" | "Z"; p: number[] };

export interface GeometryPart {
  cmds: PathCmd[];
  fill: boolean;
  stroke: boolean;
}

const NUM = /[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?/g;

/** Parses SVG path data into absolute M/L/C/Q/Z commands. */
export function parseSvgPath(d: string): PathCmd[] {
  const out: PathCmd[] = [];
  const re = /([MmLlHhVvCcSsQqTtAaZz])([^MmLlHhVvCcSsQqTtAaZz]*)/g;
  let x = 0,
    y = 0,
    sx = 0,
    sy = 0;
  let lastC: [number, number] | null = null; // last cubic control point (for S)
  let lastQ: [number, number] | null = null; // last quadratic control point (for T)
  let m: RegExpExecArray | null;
  while ((m = re.exec(d))) {
    const cmd = m[1];
    const nums = (m[2].match(NUM) ?? []).map(Number);
    const rel = cmd === cmd.toLowerCase() && cmd !== "z" && cmd !== "Z";
    const C = cmd.toUpperCase();
    let i = 0;
    const take = (n: number) => {
      const v = nums.slice(i, i + n);
      i += n;
      return v.length === n ? v : null;
    };
    const repeat = () => i < nums.length;
    if (C === "Z") {
      out.push({ c: "Z", p: [] });
      x = sx;
      y = sy;
      lastC = lastQ = null;
      continue;
    }
    let first = true;
    do {
      switch (C) {
        case "M": {
          const v = take(2);
          if (!v) break;
          x = rel ? x + v[0] : v[0];
          y = rel ? y + v[1] : v[1];
          // subsequent pairs after a moveto are implicit linetos
          out.push({ c: first ? "M" : "L", p: [x, y] });
          if (first) {
            sx = x;
            sy = y;
          }
          lastC = lastQ = null;
          break;
        }
        case "L": {
          const v = take(2);
          if (!v) break;
          x = rel ? x + v[0] : v[0];
          y = rel ? y + v[1] : v[1];
          out.push({ c: "L", p: [x, y] });
          lastC = lastQ = null;
          break;
        }
        case "H": {
          const v = take(1);
          if (!v) break;
          x = rel ? x + v[0] : v[0];
          out.push({ c: "L", p: [x, y] });
          lastC = lastQ = null;
          break;
        }
        case "V": {
          const v = take(1);
          if (!v) break;
          y = rel ? y + v[0] : v[0];
          out.push({ c: "L", p: [x, y] });
          lastC = lastQ = null;
          break;
        }
        case "C": {
          const v = take(6);
          if (!v) break;
          const p = rel ? [x + v[0], y + v[1], x + v[2], y + v[3], x + v[4], y + v[5]] : v;
          out.push({ c: "C", p });
          lastC = [p[2], p[3]];
          lastQ = null;
          x = p[4];
          y = p[5];
          break;
        }
        case "S": {
          const v = take(4);
          if (!v) break;
          const c1: [number, number] = lastC ? [2 * x - lastC[0], 2 * y - lastC[1]] : [x, y];
          const p = rel ? [x + v[0], y + v[1], x + v[2], y + v[3]] : v;
          out.push({ c: "C", p: [c1[0], c1[1], p[0], p[1], p[2], p[3]] });
          lastC = [p[0], p[1]];
          lastQ = null;
          x = p[2];
          y = p[3];
          break;
        }
        case "Q": {
          const v = take(4);
          if (!v) break;
          const p = rel ? [x + v[0], y + v[1], x + v[2], y + v[3]] : v;
          out.push({ c: "Q", p });
          lastQ = [p[0], p[1]];
          lastC = null;
          x = p[2];
          y = p[3];
          break;
        }
        case "T": {
          const v = take(2);
          if (!v) break;
          const c1: [number, number] = lastQ ? [2 * x - lastQ[0], 2 * y - lastQ[1]] : [x, y];
          const p = rel ? [x + v[0], y + v[1]] : v;
          out.push({ c: "Q", p: [c1[0], c1[1], p[0], p[1]] });
          lastQ = c1;
          lastC = null;
          x = p[0];
          y = p[1];
          break;
        }
        case "A": {
          const v = take(7);
          if (!v) break;
          const ex = rel ? x + v[5] : v[5];
          const ey = rel ? y + v[6] : v[6];
          for (const c of arcToCubic(x, y, v[0], v[1], v[2], v[3] !== 0, v[4] !== 0, ex, ey)) out.push({ c: "C", p: c });
          x = ex;
          y = ey;
          lastC = lastQ = null;
          break;
        }
      }
      first = false;
    } while (repeat() && i > 0);
  }
  return out;
}

/** Converts an SVG elliptical arc to cubic Bézier segments. */
export function arcToCubic(x1: number, y1: number, rx: number, ry: number, phiDeg: number, largeArc: boolean, sweep: boolean, x2: number, y2: number): number[][] {
  if (rx === 0 || ry === 0 || (x1 === x2 && y1 === y2)) return x1 === x2 && y1 === y2 ? [] : [[x1, y1, x2, y2, x2, y2]];
  rx = Math.abs(rx);
  ry = Math.abs(ry);
  const phi = (phiDeg * Math.PI) / 180;
  const cos = Math.cos(phi),
    sin = Math.sin(phi);
  const dx = (x1 - x2) / 2,
    dy = (y1 - y2) / 2;
  const x1p = cos * dx + sin * dy;
  const y1p = -sin * dx + cos * dy;
  const lambda = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry);
  if (lambda > 1) {
    rx *= Math.sqrt(lambda);
    ry *= Math.sqrt(lambda);
  }
  const num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p;
  const den = rx * rx * y1p * y1p + ry * ry * x1p * x1p;
  let coef = Math.sqrt(Math.max(0, num / den));
  if (largeArc === sweep) coef = -coef;
  const cxp = (coef * rx * y1p) / ry;
  const cyp = (-coef * ry * x1p) / rx;
  const cx = cos * cxp - sin * cyp + (x1 + x2) / 2;
  const cy = sin * cxp + cos * cyp + (y1 + y2) / 2;
  const ang = (ux: number, uy: number, vx: number, vy: number) => {
    const a = Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy);
    return a;
  };
  const t1 = ang(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry);
  let dt = ang((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry);
  if (!sweep && dt > 0) dt -= 2 * Math.PI;
  if (sweep && dt < 0) dt += 2 * Math.PI;
  const segs = Math.max(1, Math.ceil(Math.abs(dt) / (Math.PI / 2)));
  const delta = dt / segs;
  const k = (4 / 3) * Math.tan(delta / 4);
  const out: number[][] = [];
  let t = t1;
  const pt = (a: number) => [cx + rx * Math.cos(a) * cos - ry * Math.sin(a) * sin, cy + rx * Math.cos(a) * sin + ry * Math.sin(a) * cos];
  const dpt = (a: number) => [-rx * Math.sin(a) * cos - ry * Math.cos(a) * sin, -rx * Math.sin(a) * sin + ry * Math.cos(a) * cos];
  for (let s = 0; s < segs; s++) {
    const t2 = t + delta;
    const p1 = pt(t),
      p2 = pt(t2),
      d1 = dpt(t),
      d2 = dpt(t2);
    out.push([p1[0] + k * d1[0], p1[1] + k * d1[1], p2[0] - k * d2[0], p2[1] - k * d2[1], p2[0], p2[1]]);
    t = t2;
  }
  // snap the end point exactly
  const last = out[out.length - 1];
  last[4] = x2;
  last[5] = y2;
  return out;
}

const attr = (tag: string, name: string): string | undefined => {
  const m = new RegExp(`\\s${name}\\s*=\\s*(['"])(.*?)\\1`).exec(tag);
  return m ? m[2] : undefined;
};
const n = (v: string | undefined, d = 0) => {
  const f = parseFloat(v ?? "");
  return Number.isFinite(f) ? f : d;
};

/** Extracts the geometry of the first `<svg>` in renderer HTML. */
export function geometryFromSvgMarkup(html: string): GeometryPart[] {
  const start = html.indexOf("<svg");
  if (start === -1) return [];
  const end = html.indexOf("</svg>", start);
  const svg = html.slice(start, end === -1 ? undefined : end);
  // ignore definitions (markers, gradients, patterns)
  const body = svg.replace(/<defs[\s\S]*?<\/defs>/g, "").replace(/<marker[\s\S]*?<\/marker>/g, "");
  const parts: GeometryPart[] = [];
  for (const m of body.matchAll(/<(path|rect|ellipse|circle|polygon|polyline|line)\b[^>]*>/g)) {
    const tag = m[0];
    const fillAttr = attr(tag, "fill");
    const strokeAttr = attr(tag, "stroke");
    const fill = !(fillAttr === "none" || fillAttr === "transparent");
    const stroke = !(strokeAttr === undefined || strokeAttr === "" || strokeAttr === "none");
    let d = "";
    switch (m[1]) {
      case "path":
        d = attr(tag, "d") ?? "";
        break;
      case "rect": {
        const x = n(attr(tag, "x")),
          y = n(attr(tag, "y")),
          w = n(attr(tag, "width")),
          h = n(attr(tag, "height"));
        d = `M${x},${y} L${x + w},${y} L${x + w},${y + h} L${x},${y + h} Z`;
        break;
      }
      case "ellipse":
      case "circle": {
        const cx = n(attr(tag, "cx")),
          cy = n(attr(tag, "cy"));
        const rx = m[1] === "circle" ? n(attr(tag, "r")) : n(attr(tag, "rx"));
        const ry = m[1] === "circle" ? rx : n(attr(tag, "ry"));
        d = `M${cx - rx},${cy} A${rx},${ry} 0 1,1 ${cx + rx},${cy} A${rx},${ry} 0 1,1 ${cx - rx},${cy} Z`;
        break;
      }
      case "polygon":
      case "polyline": {
        const pts = (attr(tag, "points") ?? "").match(NUM)?.map(Number) ?? [];
        if (pts.length < 4) break;
        d = "M" + pts[0] + "," + pts[1];
        for (let i = 2; i + 1 < pts.length; i += 2) d += " L" + pts[i] + "," + pts[i + 1];
        if (m[1] === "polygon") d += " Z";
        break;
      }
      case "line":
        d = `M${n(attr(tag, "x1"))},${n(attr(tag, "y1"))} L${n(attr(tag, "x2"))},${n(attr(tag, "y2"))}`;
        break;
    }
    const cmds = parseSvgPath(d);
    if (cmds.length) parts.push({ cmds, fill: fill && m[1] !== "line" && m[1] !== "polyline", stroke });
  }
  return parts;
}

/** Integer coordinate space used for enhanced geometry (as written by LibreOffice). */
export const ENHANCED_SIZE = 21600;

/**
 * Serialises geometry as `draw:enhanced-path` (F = no fill, S = no stroke,
 * N = end of sub-path). Coordinates given in a `width` x `height` space are
 * mapped to integers in 0..21600 (decimals are rejected by some importers).
 */
export function toEnhancedPath(parts: GeometryPart[], width: number, height: number): string {
  const sx = ENHANCED_SIZE / Math.max(width, 1e-6);
  const sy = ENHANCED_SIZE / Math.max(height, 1e-6);
  return parts
    .map((part) => {
      let s = "";
      for (const c of part.cmds) {
        if (c.c === "Z") {
          s += "Z ";
          continue;
        }
        s += c.c + " " + c.p.map((v, i) => String(Math.round(v * (i % 2 === 0 ? sx : sy)))).join(" ") + " ";
      }
      if (!part.fill) s += "F ";
      if (!part.stroke) s += "S ";
      return s + "N";
    })
    .join(" ");
}

/** Bounding box of the geometry (for svg:viewBox). */
export function geometryBounds(parts: GeometryPart[]): { minX: number; minY: number; maxX: number; maxY: number } | null {
  let minX = Infinity,
    minY = Infinity,
    maxX = -Infinity,
    maxY = -Infinity;
  for (const p of parts)
    for (const c of p.cmds)
      for (let i = 0; i + 1 < c.p.length; i += 2) {
        minX = Math.min(minX, c.p[i]);
        maxX = Math.max(maxX, c.p[i]);
        minY = Math.min(minY, c.p[i + 1]);
        maxY = Math.max(maxY, c.p[i + 1]);
      }
  return Number.isFinite(minX) ? { minX, minY, maxX, maxY } : null;
}
