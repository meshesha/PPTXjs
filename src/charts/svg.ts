/**
 * Dependency-free SVG chart renderer (replaces d3 + nvd3 used by PPTXjs v1).
 * Produces a static, accessible `<svg>` string sized to the chart frame.
 */
import { encodeHtmlEntities } from "../core/xml";
import type { ChartGroup, ChartModel, ChartSeries } from "./model";

/** Office default palette, used when neither the series nor the theme define colours. */
export const DEFAULT_PALETTE = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47", "#264478", "#9E480E", "#636363", "#997300"];

export interface ChartRenderOptions {
  width: number;
  height: number;
  /** Colours for series without explicit fill (theme accents). */
  palette?: string[];
  fontFamily?: string;
  /** Unique id used for clip paths etc. */
  id?: string;
}

const esc = (s: string) => encodeHtmlEntities(s);
const r2 = (n: number) => Math.round(n * 100) / 100;

function niceStep(range: number, ticks: number): number {
  if (range <= 0 || !Number.isFinite(range)) return 1;
  const raw = range / ticks;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const norm = raw / mag;
  const nice = norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10;
  return nice * mag;
}

function niceScale(min: number, max: number, ticks = 5): { min: number; max: number; step: number } {
  if (min === max) {
    if (min === 0) return { min: 0, max: 1, step: 0.2 };
    min = Math.min(0, min);
    max = Math.max(0, max);
  }
  const step = niceStep(max - min, ticks);
  return { min: Math.floor(min / step) * step, max: Math.ceil(max / step) * step, step };
}

function fmtNum(n: number): string {
  if (Math.abs(n) >= 1e6) return r2(n / 1e6) + "M";
  if (Math.abs(n) >= 1e4) return r2(n / 1e3) + "K";
  return String(r2(n));
}

function seriesColor(s: ChartSeries, i: number, palette: string[]): string {
  return s.color || palette[i % palette.length];
}

/** Renders a chart model to an SVG string. */
export function renderChartSvg(model: ChartModel, opts: ChartRenderOptions): string {
  const W = Math.max(10, opts.width);
  const H = Math.max(10, opts.height);
  const palette = opts.palette && opts.palette.length ? opts.palette : DEFAULT_PALETTE;
  const font = opts.fontFamily ?? "Calibri, Arial, sans-serif";
  const fs = Math.max(9, Math.min(14, Math.round(Math.min(W, H) / 28)));
  let body = "";
  let top = 6;
  if (model.title) {
    body += `<text x="${r2(W / 2)}" y="${top + fs + 2}" text-anchor="middle" font-size="${fs + 4}" fill="#404040">${esc(model.title)}</text>`;
    top += fs + 12;
  }

  const first = model.groups[0];
  const isPie = first.kind === "pie" || first.kind === "doughnut";

  // legend entries
  const legendItems: Array<{ name: string; color: string }> = [];
  if (isPie) {
    const s = first.series[0];
    model.categories.forEach((c, i) => legendItems.push({ name: c, color: s.pointColors?.[i] || palette[i % palette.length] }));
  } else {
    let n = 0;
    for (const g of model.groups) for (const s of g.series) legendItems.push({ name: s.name, color: seriesColor(s, n++, palette) });
  }
  const plot = { x: 8, y: top, w: W - 16, h: H - top - 8 };
  let legend = "";
  if (model.legend && legendItems.length) {
    const lh = fs + 6;
    if (model.legend === "b" || model.legend === "t") {
      const widths = legendItems.map((it) => 14 + it.name.length * fs * 0.55 + 12);
      const total = widths.reduce((a, b) => a + b, 0);
      let x = Math.max(4, (W - total) / 2);
      const y = model.legend === "b" ? H - lh : top;
      legendItems.forEach((it, i) => {
        legend += `<rect x="${r2(x)}" y="${r2(y + 3)}" width="10" height="10" fill="${it.color}"/><text x="${r2(x + 14)}" y="${r2(y + fs + 1)}" font-size="${fs}" fill="#595959">${esc(it.name)}</text>`;
        x += widths[i];
      });
      if (model.legend === "b") plot.h -= lh + 4;
      else {
        plot.y += lh + 4;
        plot.h -= lh + 4;
      }
    } else {
      const lw = Math.min(W * 0.35, 20 + Math.max(...legendItems.map((it) => it.name.length)) * fs * 0.55);
      const x = model.legend === "l" ? 6 : W - lw;
      let y = plot.y + Math.max(0, (plot.h - legendItems.length * lh) / 2);
      for (const it of legendItems) {
        legend += `<rect x="${r2(x)}" y="${r2(y + 3)}" width="10" height="10" fill="${it.color}"/><text x="${r2(x + 14)}" y="${r2(y + fs + 1)}" font-size="${fs}" fill="#595959">${esc(it.name)}</text>`;
        y += lh;
      }
      plot.w -= lw + 6;
      if (model.legend === "l") plot.x += lw + 6;
    }
  }

  body += isPie ? renderPie(first, model, plot, palette, fs) : renderAxes(model, plot, palette, fs);
  body += legend;

  const label = model.title ? esc(model.title) : "Chart";
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" class="pptx-chart" width="${r2(W)}" height="${r2(H)}" viewBox="0 0 ${r2(W)} ${r2(H)}" ` +
    `role="img" aria-label="${label}" font-family="${esc(font)}">${body}</svg>`
  );
}

type Rect = { x: number; y: number; w: number; h: number };

function renderPie(g: ChartGroup, model: ChartModel, p: Rect, palette: string[], fs: number): string {
  const s = g.series[0];
  const vals = s.values.map((v) => (v && v > 0 ? v : 0));
  const total = vals.reduce((a, b) => a + b, 0);
  if (total <= 0) return "";
  const cx = p.x + p.w / 2,
    cy = p.y + p.h / 2;
  const R = Math.max(1, Math.min(p.w, p.h) / 2 - 4);
  const r = g.kind === "doughnut" ? (R * (g.holeSize ?? 50)) / 100 : 0;
  let a0 = -Math.PI / 2;
  let out = "";
  vals.forEach((v, i) => {
    if (v <= 0) return;
    const a1 = a0 + (v / total) * Math.PI * 2;
    const color = s.pointColors?.[i] || (g.varyColors ? palette[i % palette.length] : seriesColor(s, 0, palette));
    const large = a1 - a0 > Math.PI ? 1 : 0;
    const pt = (rad: number, a: number) => `${r2(cx + rad * Math.cos(a))},${r2(cy + rad * Math.sin(a))}`;
    let d: string;
    if (v === total) {
      d = `M${pt(R, 0)} A${r2(R)},${r2(R)} 0 1,1 ${pt(R, Math.PI)} A${r2(R)},${r2(R)} 0 1,1 ${pt(R, 0)} z`;
      if (r > 0) d += ` M${pt(r, 0)} A${r2(r)},${r2(r)} 0 1,0 ${pt(r, Math.PI)} A${r2(r)},${r2(r)} 0 1,0 ${pt(r, 0)} z`;
    } else if (r > 0) {
      d = `M${pt(R, a0)} A${r2(R)},${r2(R)} 0 ${large},1 ${pt(R, a1)} L${pt(r, a1)} A${r2(r)},${r2(r)} 0 ${large},0 ${pt(r, a0)} z`;
    } else {
      d = `M${r2(cx)},${r2(cy)} L${pt(R, a0)} A${r2(R)},${r2(R)} 0 ${large},1 ${pt(R, a1)} z`;
    }
    out += `<path d="${d}" fill="${color}" stroke="#fff" stroke-width="1" fill-rule="evenodd"><title>${esc(model.categories[i] ?? "")}: ${v}</title></path>`;
    if (model.showValues) {
      const am = (a0 + a1) / 2;
      const lr = r > 0 ? (R + r) / 2 : R * 0.65;
      out += `<text x="${r2(cx + lr * Math.cos(am))}" y="${r2(cy + lr * Math.sin(am) + fs / 3)}" text-anchor="middle" font-size="${fs}" fill="#fff">${esc(fmtNum(v))}</text>`;
    }
    a0 = a1;
  });
  return out;
}

function renderAxes(model: ChartModel, p: Rect, palette: string[], fs: number): string {
  const cats = model.categories;
  const nCat = Math.max(1, cats.length);
  const horizontal = model.groups.some((g) => g.kind === "bar" && g.barDir === "bar");
  const scatter = model.groups.every((g) => g.kind === "scatter");

  // value range (stacked groups sum per category)
  let vmin = 0,
    vmax = 0;
  let xmin = Infinity,
    xmax = -Infinity;
  for (const g of model.groups) {
    if (g.grouping === "percentStacked") {
      vmax = Math.max(vmax, 1);
      continue;
    }
    if (g.grouping === "stacked") {
      for (let i = 0; i < nCat; i++) {
        let pos = 0,
          neg = 0;
        for (const s of g.series) {
          const v = s.values[i] ?? 0;
          if (v >= 0) pos += v;
          else neg += v;
        }
        vmax = Math.max(vmax, pos);
        vmin = Math.min(vmin, neg);
      }
    } else {
      for (const s of g.series)
        for (const v of s.values)
          if (v !== null) {
            vmax = Math.max(vmax, v);
            vmin = Math.min(vmin, v);
          }
    }
    if (g.kind === "scatter")
      for (const s of g.series)
        for (const x of s.xValues ?? [])
          if (x !== null) {
            xmin = Math.min(xmin, x);
            xmax = Math.max(xmax, x);
          }
  }
  const pct = model.groups.some((g) => g.grouping === "percentStacked");
  const scale = niceScale(vmin, vmax);
  const yLabels: string[] = [];
  for (let v = scale.min; v <= scale.max + scale.step / 2; v += scale.step) yLabels.push(pct ? Math.round(v * 100) + "%" : fmtNum(v));
  const labelW = Math.max(...yLabels.map((l) => l.length)) * fs * 0.6 + 6;
  const catLabelW = horizontal ? Math.min(p.w * 0.3, Math.max(...cats.map((c) => c.length)) * fs * 0.55 + 6) : 0;

  const area: Rect = horizontal
    ? { x: p.x + catLabelW, y: p.y + 4, w: p.w - catLabelW - 8, h: p.h - fs - 10 }
    : { x: p.x + labelW, y: p.y + 4, w: p.w - labelW - 4, h: p.h - fs - 12 };
  if (area.w <= 0 || area.h <= 0) return "";

  const val = (v: number) => (v - scale.min) / (scale.max - scale.min || 1);
  const vPos = horizontal ? (v: number) => area.x + val(v) * area.w : (v: number) => area.y + area.h - val(v) * area.h;
  const xs = scatter ? niceScale(Number.isFinite(xmin) ? xmin : 0, Number.isFinite(xmax) ? xmax : 1) : null;

  let out = "";
  // grid + value labels
  let k = 0;
  for (let v = scale.min; v <= scale.max + scale.step / 2; v += scale.step, k++) {
    const pos = vPos(v);
    if (horizontal) {
      out += `<line x1="${r2(pos)}" y1="${r2(area.y)}" x2="${r2(pos)}" y2="${r2(area.y + area.h)}" stroke="#d9d9d9" stroke-width="1"/>`;
      out += `<text x="${r2(pos)}" y="${r2(area.y + area.h + fs + 4)}" text-anchor="middle" font-size="${fs}" fill="#595959">${esc(yLabels[k])}</text>`;
    } else {
      out += `<line x1="${r2(area.x)}" y1="${r2(pos)}" x2="${r2(area.x + area.w)}" y2="${r2(pos)}" stroke="#d9d9d9" stroke-width="1"/>`;
      out += `<text x="${r2(area.x - 4)}" y="${r2(pos + fs / 3)}" text-anchor="end" font-size="${fs}" fill="#595959">${esc(yLabels[k])}</text>`;
    }
  }
  // category labels
  const band = (horizontal ? area.h : area.w) / nCat;
  if (scatter && xs) {
    for (let v = xs.min; v <= xs.max + xs.step / 2; v += xs.step) {
      const x = area.x + ((v - xs.min) / (xs.max - xs.min || 1)) * area.w;
      out += `<text x="${r2(x)}" y="${r2(area.y + area.h + fs + 4)}" text-anchor="middle" font-size="${fs}" fill="#595959">${esc(fmtNum(v))}</text>`;
    }
  } else {
    const every = Math.max(1, Math.ceil((nCat * fs * 3) / (horizontal ? area.h * 3 : area.w)));
    cats.forEach((c, i) => {
      if (i % every) return;
      if (horizontal) {
        out += `<text x="${r2(area.x - 4)}" y="${r2(area.y + band * (i + 0.5) + fs / 3)}" text-anchor="end" font-size="${fs}" fill="#595959">${esc(c)}</text>`;
      } else {
        out += `<text x="${r2(area.x + band * (i + 0.5))}" y="${r2(area.y + area.h + fs + 4)}" text-anchor="middle" font-size="${fs}" fill="#595959">${esc(c)}</text>`;
      }
    });
  }
  // axis line
  const zero = vPos(Math.max(scale.min, Math.min(0, scale.max)));
  out += horizontal
    ? `<line x1="${r2(zero)}" y1="${r2(area.y)}" x2="${r2(zero)}" y2="${r2(area.y + area.h)}" stroke="#bfbfbf"/>`
    : `<line x1="${r2(area.x)}" y1="${r2(zero)}" x2="${r2(area.x + area.w)}" y2="${r2(zero)}" stroke="#bfbfbf"/>`;

  let colorIdx = 0;
  const barGroups = model.groups.filter((g) => g.kind === "bar");
  for (const g of model.groups) {
    const colors = g.series.map((s) => seriesColor(s, colorIdx++, palette));
    if (g.kind === "bar") {
      const clustered = g.grouping === "clustered" || g.grouping === "standard";
      const slots = clustered ? g.series.length * barGroups.length : barGroups.length;
      const gi = barGroups.indexOf(g);
      const inner = band * 0.7;
      const bw = inner / Math.max(1, slots);
      for (let i = 0; i < nCat; i++) {
        let pos = 0,
          neg = 0;
        const total = g.series.reduce((a, s) => a + Math.abs(s.values[i] ?? 0), 0) || 1;
        g.series.forEach((s, si) => {
          let v = s.values[i];
          if (v === null || v === undefined) return;
          if (g.grouping === "percentStacked") v = v / total;
          let from: number, to: number;
          if (clustered) {
            from = 0;
            to = v;
          } else if (v >= 0) {
            from = pos;
            to = pos += v;
          } else {
            from = neg;
            to = neg += v;
          }
          const slot = clustered ? gi * g.series.length + si : gi;
          const offset = band * i + (band - inner) / 2 + slot * bw;
          const a = vPos(from),
            b = vPos(to);
          const fill = s.pointColors?.[i] || (g.varyColors && g.series.length === 1 ? palette[i % palette.length] : colors[si]);
          const rect = horizontal
            ? { x: Math.min(a, b), y: area.y + offset, w: Math.abs(b - a), h: bw }
            : { x: area.x + offset, y: Math.min(a, b), w: bw, h: Math.abs(b - a) };
          out += `<rect x="${r2(rect.x)}" y="${r2(rect.y)}" width="${r2(Math.max(0, rect.w))}" height="${r2(Math.max(0, rect.h))}" fill="${fill}"><title>${esc(s.name)} - ${esc(cats[i] ?? "")}: ${s.values[i]}</title></rect>`;
          if (model.showValues) {
            const tx = horizontal ? rect.x + rect.w + 2 : rect.x + rect.w / 2;
            const ty = horizontal ? rect.y + rect.h / 2 + fs / 3 : rect.y - 2;
            out += `<text x="${r2(tx)}" y="${r2(ty)}" text-anchor="${horizontal ? "start" : "middle"}" font-size="${fs}" fill="#404040">${esc(fmtNum(s.values[i] as number))}</text>`;
          }
        });
      }
    } else if (g.kind === "line" || g.kind === "area" || g.kind === "radar") {
      const stacked = g.grouping === "stacked" || g.grouping === "percentStacked";
      const base = new Array(nCat).fill(0);
      g.series.forEach((s, si) => {
        const pts: Array<[number, number] | null> = [];
        const totals = cats.map((_, i) => g.series.reduce((a, ss) => a + Math.abs(ss.values[i] ?? 0), 0) || 1);
        for (let i = 0; i < nCat; i++) {
          let v = s.values[i];
          if (v === null || v === undefined) {
            pts.push(null);
            continue;
          }
          if (g.grouping === "percentStacked") v = v / totals[i];
          const y = stacked ? (base[i] += v) : v;
          pts.push([area.x + band * (i + 0.5), vPos(y)]);
        }
        const valid = pts.filter((q): q is [number, number] => q !== null);
        if (!valid.length) return;
        const line = valid.map((q, j) => `${j ? "L" : "M"}${r2(q[0])},${r2(q[1])}`).join(" ");
        if (g.kind === "area") {
          out += `<path d="${line} L${r2(valid[valid.length - 1][0])},${r2(zero)} L${r2(valid[0][0])},${r2(zero)} z" fill="${colors[si]}" fill-opacity="0.75"/>`;
        } else {
          out += `<path d="${line}" fill="none" stroke="${colors[si]}" stroke-width="2.25" stroke-linejoin="round"/>`;
          for (const q of valid) out += `<circle cx="${r2(q[0])}" cy="${r2(q[1])}" r="2.5" fill="${colors[si]}"/>`;
        }
      });
    } else if (g.kind === "scatter" && xs) {
      g.series.forEach((s, si) => {
        s.values.forEach((v, i) => {
          const x = s.xValues?.[i] ?? i + 1;
          if (v === null || x === null) return;
          const px = area.x + ((x - xs.min) / (xs.max - xs.min || 1)) * area.w;
          out += `<circle cx="${r2(px)}" cy="${r2(vPos(v))}" r="3.5" fill="${colors[si]}"><title>${esc(s.name)}: (${x}, ${v})</title></circle>`;
        });
      });
    }
  }
  return out;
}
