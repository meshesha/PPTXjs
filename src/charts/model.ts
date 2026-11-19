/**
 * Extracts a renderer-independent chart model from a DrawingML chart part
 * (`c:chartSpace`). Only cached values are used, so no spreadsheet engine is
 * needed.
 */
import { asArray, decodeXmlEntities, getTextByPathList } from "../core/xml";

export type ChartKind = "bar" | "line" | "area" | "pie" | "doughnut" | "scatter" | "radar";

export interface ChartSeries {
  name: string;
  values: Array<number | null>;
  /** x values for scatter charts */
  xValues?: Array<number | null>;
  /** CSS colour of the series (`#rrggbb`) if specified in the file */
  color?: string;
  /** Per data point colours (pie/doughnut or varied bars) */
  pointColors?: Array<string | undefined>;
}

export interface ChartGroup {
  kind: ChartKind;
  /** bar charts: vertical columns (`col`) or horizontal bars (`bar`) */
  barDir: "col" | "bar";
  grouping: "clustered" | "stacked" | "percentStacked" | "standard";
  /** per-point colours (pie charts default to true) */
  varyColors: boolean;
  series: ChartSeries[];
  /** doughnut hole size in percent */
  holeSize?: number;
}

export interface ChartModel {
  title?: string;
  categories: string[];
  groups: ChartGroup[];
  legend: false | "r" | "l" | "t" | "b" | "tr";
  /** Show data labels with values */
  showValues: boolean;
}

const KIND: Record<string, ChartKind> = {
  "c:barChart": "bar",
  "c:bar3DChart": "bar",
  "c:lineChart": "line",
  "c:line3DChart": "line",
  "c:areaChart": "area",
  "c:area3DChart": "area",
  "c:pieChart": "pie",
  "c:pie3DChart": "pie",
  "c:ofPieChart": "pie",
  "c:doughnutChart": "doughnut",
  "c:scatterChart": "scatter",
  "c:bubbleChart": "scatter",
  "c:radarChart": "radar",
};

const text = (v: unknown) => (typeof v === "string" ? decodeXmlEntities(v) : v === undefined || v === null ? "" : String(v));

function readCache(ref: any): { pts: Map<number, string>; count: number } {
  const cache =
    getTextByPathList(ref, ["c:numRef", "c:numCache"]) ??
    getTextByPathList(ref, ["c:strRef", "c:strCache"]) ??
    getTextByPathList(ref, ["c:numLit"]) ??
    getTextByPathList(ref, ["c:strLit"]) ??
    getTextByPathList(ref, ["c:multiLvlStrRef", "c:multiLvlStrCache", "c:lvl"]);
  const pts = new Map<number, string>();
  let count = 0;
  const lvl = Array.isArray(cache) ? cache[0] : cache;
  if (lvl) {
    count = parseInt(getTextByPathList(lvl, ["c:ptCount", "attrs", "val"])) || 0;
    for (const pt of asArray(lvl["c:pt"])) {
      const idx = parseInt(pt?.attrs?.idx);
      if (Number.isFinite(idx)) pts.set(idx, text(pt["c:v"]));
    }
    count = Math.max(count, ...[...pts.keys()].map((k) => k + 1), 0);
  }
  return { pts, count };
}

function toNumbers(ref: any): Array<number | null> {
  const { pts, count } = readCache(ref);
  const out: Array<number | null> = [];
  for (let i = 0; i < count; i++) {
    const v = pts.get(i);
    const n = v === undefined || v === "" ? NaN : parseFloat(v);
    out.push(Number.isFinite(n) ? n : null);
  }
  return out;
}

function toStrings(ref: any): string[] {
  const { pts, count } = readCache(ref);
  const out: string[] = [];
  for (let i = 0; i < count; i++) out.push(pts.get(i) ?? "");
  return out;
}

function richText(node: any): string {
  const paras = asArray(getTextByPathList(node, ["c:tx", "c:rich", "a:p"]));
  return paras
    .map((p: any) => asArray(p?.["a:r"]).map((r: any) => text(r?.["a:t"])).join(""))
    .join(" ")
    .trim();
}

/**
 * @param chartSpace parsed chart part
 * @param resolveFill converts a `c:spPr` node into a CSS colour (uses the theme)
 */
export function parseChart(chartSpace: any, resolveFill: (spPr: any) => string | undefined): ChartModel | null {
  const chart = getTextByPathList(chartSpace, ["c:chartSpace", "c:chart"]);
  const plotArea = chart?.["c:plotArea"];
  if (!plotArea) return null;

  const groups: ChartGroup[] = [];
  let categories: string[] = [];
  const sortedKeys = Object.keys(plotArea)
    .filter((k) => KIND[k])
    .sort((a, b) => Number(asArray(plotArea[a])[0]?.attrs?.order ?? 0) - Number(asArray(plotArea[b])[0]?.attrs?.order ?? 0));

  for (const key of sortedKeys) {
    for (const g of asArray(plotArea[key])) {
      const kind = KIND[key];
      const grouping = (getTextByPathList(g, ["c:grouping", "attrs", "val"]) ?? "clustered") as ChartGroup["grouping"];
      const varyAttr = getTextByPathList(g, ["c:varyColors", "attrs", "val"]);
      const group: ChartGroup = {
        kind,
        barDir: getTextByPathList(g, ["c:barDir", "attrs", "val"]) === "bar" ? "bar" : "col",
        grouping: kind === "bar" || kind === "line" || kind === "area" ? grouping : "standard",
        varyColors: varyAttr === undefined ? kind === "pie" || kind === "doughnut" : varyAttr === "1" || varyAttr === "true",
        series: [],
      };
      if (kind === "doughnut") group.holeSize = parseInt(getTextByPathList(g, ["c:holeSize", "attrs", "val"])) || 50;
      const series = asArray(g["c:ser"]).sort(
        (a: any, b: any) =>
          Number(getTextByPathList(a, ["c:order", "attrs", "val"]) ?? 0) - Number(getTextByPathList(b, ["c:order", "attrs", "val"]) ?? 0),
      );
      series.forEach((ser: any, i: number) => {
        const name =
          text(getTextByPathList(ser, ["c:tx", "c:strRef", "c:strCache", "c:pt", "c:v"])) ||
          text(getTextByPathList(ser, ["c:tx", "c:v"])) ||
          `Series ${i + 1}`;
        const s: ChartSeries = { name, values: [] };
        if (kind === "scatter") {
          s.xValues = toNumbers(ser["c:xVal"]);
          s.values = toNumbers(ser["c:yVal"]);
        } else {
          s.values = toNumbers(ser["c:val"]);
          if (categories.length === 0 && ser["c:cat"]) categories = toStrings(ser["c:cat"]);
        }
        s.color = resolveFill(ser["c:spPr"]);
        const dPts = asArray(ser["c:dPt"]);
        if (dPts.length) {
          s.pointColors = [];
          for (const dp of dPts) {
            const idx = parseInt(getTextByPathList(dp, ["c:idx", "attrs", "val"]));
            if (Number.isFinite(idx)) s.pointColors[idx] = resolveFill(dp["c:spPr"]);
          }
        }
        group.series.push(s);
      });
      if (group.series.length) groups.push(group);
    }
  }
  if (!groups.length) return null;
  if (categories.length === 0) {
    const n = Math.max(...groups.flatMap((g) => g.series.map((s) => s.values.length)));
    categories = Array.from({ length: n }, (_, i) => String(i + 1));
  }

  const autoDeleted = getTextByPathList(chart, ["c:autoTitleDeleted", "attrs", "val"]) === "1";
  let title: string | undefined = chart["c:title"] ? richText(chart["c:title"]) : undefined;
  if (chart["c:title"] && !title && !autoDeleted && groups.length === 1 && groups[0].series.length === 1) {
    title = groups[0].series[0].name;
  }

  const legendNode = chart["c:legend"];
  const legend = legendNode ? ((getTextByPathList(legendNode, ["c:legendPos", "attrs", "val"]) ?? "r") as ChartModel["legend"]) : false;
  const showValues = asArray(plotArea[sortedKeys[0]]).some(
    (g: any) => getTextByPathList(g, ["c:dLbls", "c:showVal", "attrs", "val"]) === "1",
  );
  return { title: title || undefined, categories, groups, legend, showValues };
}
