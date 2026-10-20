/**
 * A DrawingML chart part back as the ECharts option that writes it -- the
 * inverse of `writer/chart-part.ts`. Mirrors PHP `LastWord\Reader\ChartReader`:
 * read from the series' CACHES (what Word and LibreOffice draw), returned in
 * CANONICAL form so `read(write(doc))` equals `doc` for a chart already in it.
 *
 * Bar, line and pie only. Anything else -- radar, scatter, 3-D, a part that is
 * not XML -- returns null and the reader skips the drawing.
 */
import { at, el, els, parseXml, type XmlNode } from "./xml";
import { DEFAULT_PALETTE, chartParseNum } from "../helpers/chart";

/* eslint-disable @typescript-eslint/no-explicit-any */
type Any = any;

export const NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";

const KINDS: Record<string, "bar" | "line" | "pie"> = { barChart: "bar", lineChart: "line", pieChart: "pie" };

export function readChart(xml: string): { option: Record<string, Any>; numberFormat: string | null } | null {
  if (/<!DOCTYPE/i.test(xml)) return null;
  let root: XmlNode | null;
  try {
    root = parseXml(xml);
  } catch {
    return null;
  }
  if (!root) return null;

  const chart = el(root, "chart");
  const plotArea = el(chart, "plotArea");
  if (!chart || !plotArea) return null;

  // The first plot decides; a plot this package does not model ends the read.
  const plot = plotArea.children.find((n) => n.name.endsWith("Chart"));
  if (!plot || KINDS[plot.name] === undefined) return null;
  const kind = KINDS[plot.name]!;
  const horizontal = kind === "bar" && val(el(plot, "barDir")) === "bar";
  const grouping = val(el(plot, "grouping"));
  const stacked = kind === "bar" && (grouping === "stacked" || grouping === "percentStacked");

  const sers = els(plot, "ser");
  if (sers.length === 0) return null;
  const categories = points(sers[0]!, "cat");

  const series = sers.map((ser, idx) => readSeries(ser, idx, kind, stacked, categories));

  const option: Record<string, Any> = {};
  const title = titleText(el(chart, "title"));
  if (title !== "") option.title = { text: title };
  const legend = el(chart, "legend");
  if (legend && val(el(legend, "delete")) !== "1") option.legend = { show: true };
  const valAx = el(plotArea, "valAx");
  if (kind !== "pie") {
    const categoryAxis: Record<string, Any> = { type: "category", data: categories };
    const catTitle = titleText(el(el(plotArea, "catAx"), "title"));
    if (catTitle !== "") categoryAxis.name = catTitle;
    const valueAxis: Record<string, Any> = { type: "value" };
    const valTitle = titleText(el(valAx, "title"));
    if (valTitle !== "") valueAxis.name = valTitle;
    option.xAxis = horizontal ? valueAxis : categoryAxis;
    option.yAxis = horizontal ? categoryAxis : valueAxis;
  }
  option.series = series;

  // The value axis's own number format, else the cache's when it is not General.
  let numberFormat: string | null = null;
  const numFmt = el(valAx, "numFmt");
  if (numFmt && at(numFmt, "sourceLinked") !== "1" && (at(numFmt, "formatCode") ?? "") !== "") {
    numberFormat = at(numFmt, "formatCode")!;
  } else {
    const code = descendant(el(sers[0], "val"), "formatCode")?.text ?? "";
    if (code !== "" && code !== "General") numberFormat = code;
  }

  return { option, numberFormat };
}

function readSeries(ser: XmlNode, idx: number, kind: "bar" | "line" | "pie", stacked: boolean, categories: string[]): Record<string, Any> {
  const fallback = DEFAULT_PALETTE[idx % DEFAULT_PALETTE.length];
  const one: Record<string, Any> = { type: kind };

  const name = descendant(el(ser, "tx"), "v")?.text ?? "";
  if (name !== "" && name !== `Series ${idx + 1}`) one.name = name;
  if (stacked) one.stack = "total";
  if (kind === "line" && val(el(ser, "smooth")) === "1") one.smooth = true;
  const dLbls = el(ser, "dLbls");
  if (dLbls && val(el(dLbls, "showVal")) === "1") one.label = { show: true };
  if (kind !== "pie") {
    const spPr = el(ser, "spPr");
    const color = solidColor(kind === "line" ? el(spPr, "ln") : spPr);
    if (color !== null && color !== fallback) one.itemStyle = { color: "#" + color };
  }

  const pointColors = new Map<number, string>();
  for (const dPt of els(ser, "dPt")) {
    const i = val(el(dPt, "idx"));
    const holder = kind === "line" ? el(dPt, "marker") : dPt;
    const color = solidColor(el(holder, "spPr"));
    if (i !== null && /^\d+$/.test(i) && color !== null) pointColors.set(parseInt(i, 10), color);
  }

  one.data = points(ser, "val").map((text, i) => {
    const value = chartParseNum(text);
    if (kind === "pie") {
      const point: Record<string, Any> = { name: categories[i] ?? "", value };
      const slice = pointColors.get(i);
      if (slice !== undefined && slice !== DEFAULT_PALETTE[i % DEFAULT_PALETTE.length]) point.itemStyle = { color: "#" + slice };
      return point;
    }
    const color = pointColors.get(i);
    return color !== undefined ? { value, itemStyle: { color: "#" + color } } : value;
  });

  return one;
}

/** A series' cached points under `host` (cat or val), in index order; a missing index is "". */
function points(ser: XmlNode, host: string): string[] {
  const node = el(ser, host);
  if (!node) return [];
  const countText = at(descendant(node, "ptCount"), "val") ?? "";
  const count = Math.min(/^\d+$/.test(countText) ? parseInt(countText, 10) : 0, 100000);
  const out: string[] = new Array(count).fill("");
  for (const pt of descendants(node, "pt")) {
    const i = at(pt, "idx") ?? "";
    if (/^\d+$/.test(i) && parseInt(i, 10) < count) out[parseInt(i, 10)] = el(pt, "v")?.text ?? "";
  }
  return out;
}

/** The upper-case hex of a `<a:solidFill><a:srgbClr>` directly under `holder`. */
function solidColor(holder: XmlNode | undefined): string | null {
  const hex = (at(el(el(holder, "solidFill"), "srgbClr"), "val") ?? "").toUpperCase();
  return /^[0-9A-F]{6}$/.test(hex) ? hex : null;
}

/** A title's text: every `a:t` under its rich text, joined. */
function titleText(title: XmlNode | undefined): string {
  return title ? descendants(title, "t").map((t) => t.text).join("") : "";
}

function val(node: XmlNode | undefined): string | null {
  return node?.attrs["val"] ?? null;
}

function descendant(node: XmlNode | undefined, name: string): XmlNode | undefined {
  return node ? descendants(node, name)[0] : undefined;
}

function descendants(node: XmlNode, name: string): XmlNode[] {
  const out: XmlNode[] = [];
  for (const child of node.children) {
    if (child.name === name) out.push(child);
    out.push(...descendants(child, name));
  }
  return out;
}
