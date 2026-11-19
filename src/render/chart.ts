/**
 * Chart graphic frames: parses the chart part and renders it to inline SVG
 * (v1 queued the data and drew it later with d3/nvd3).
 */
import type { WarpObj } from "./context";
import { getPosition, getSize } from "./nodes";
import { getTextByPathList } from "../core/xml";
import { getSolidFill } from "./color";
import { parseChart } from "../charts/model";
import { renderChartSvg } from "../charts/svg";
import { slideFactor } from "./units";

function cssColor(hex: string | undefined): string | undefined {
  if (!hex || typeof hex !== "string") return undefined;
  const h = hex.replace(/^#/, "");
  if (/^[0-9a-fA-F]{6}$/.test(h)) return "#" + h;
  if (/^[0-9a-fA-F]{8}$/.test(h)) {
    const a = parseInt(h.slice(6), 16) / 255;
    return `rgba(${parseInt(h.slice(0, 2), 16)},${parseInt(h.slice(2, 4), 16)},${parseInt(h.slice(4, 6), 16)},${Math.round(a * 100) / 100})`;
  }
  return undefined;
}

/** Theme accent colours accent1..accent6 as CSS colours. */
export function themePalette(warpObj: WarpObj): string[] {
  const out: string[] = [];
  for (let i = 1; i <= 6; i++) {
    try {
      const c = cssColor(getSolidFill({ "a:schemeClr": { attrs: { val: "accent" + i } } }, undefined, undefined, warpObj));
      if (c) out.push(c);
    } catch {
      /* theme without colour scheme */
    }
  }
  return out;
}

export function genChart(node: any, warpObj: WarpObj): string {
  const ctx = warpObj.ctx;
  const id = `${ctx.idPrefix}-chart-${ctx.chartCounter++}`;
  const order = node?.attrs?.order;
  const xfrmNode = getTextByPathList(node, ["p:xfrm"]);
  const open =
    `<div id='${id}' class='block content chart' style='` +
    getPosition(xfrmNode, node, undefined, undefined) +
    getSize(xfrmNode, undefined, undefined) +
    ` z-index: ${order};'>`;

  const rid = getTextByPathList(node, ["a:graphic", "a:graphicData", "c:chart", "attrs", "r:id"]);
  const target = rid ? warpObj.slideResObj[rid]?.target : undefined;
  const content = target ? warpObj.zip.xmlFresh(target) : null;
  if (!content) {
    ctx.warn(`Chart part not found (${target ?? rid})`);
    return open + "</div>";
  }
  const model = parseChart(content, (spPr) => {
    const fill = spPr?.["a:solidFill"];
    if (!fill) return undefined;
    try {
      return cssColor(getSolidFill(fill, undefined, undefined, warpObj));
    } catch {
      return undefined;
    }
  });
  if (!model) return open + "</div>";
  const width = parseInt(getTextByPathList(xfrmNode, ["a:ext", "attrs", "cx"])) * slideFactor || 0;
  const height = parseInt(getTextByPathList(xfrmNode, ["a:ext", "attrs", "cy"])) * slideFactor || 0;
  const svg = renderChartSvg(model, { width, height, palette: themePalette(warpObj), id });
  return open + svg + "</div>";
}
