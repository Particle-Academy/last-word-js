/**
 * The `chart` block's model: an Apache ECharts `option`, the same field
 * dark-slide's chart element takes, so one prompt teaches charts in both
 * formats and a chart can move from a deck into a report unchanged.
 *
 * Mirrors PHP `LastWord\Helpers\Chart` line for line: the supported series
 * types, the default palette, the default aspect, how an option becomes the
 * writer's spec, and how a number is spelled. The PHP package is the reference.
 */

/* eslint-disable @typescript-eslint/no-explicit-any */
type Any = any;

/**
 * The series types a document chart can draw. dark-slide adds `scatter`, which
 * has no category axis and so no place in this workbook layout; it is a
 * validation error here rather than a silent placeholder.
 */
export const SUPPORTED_CHART_TYPES = ["bar", "line", "pie"] as const;

/** Office's own accent order. Upper-case hex, no '#'. */
export const DEFAULT_PALETTE = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "264478", "9E480E"] as const;

/** Height as a fraction of width when the block gives no heightPx. */
export const DEFAULT_ASPECT = 0.6;

export interface ChartSeriesSpec {
  name: string;
  values: number[];
  color: string;
  labels: boolean;
  smooth: boolean;
  /** point index → colour; a Map so iteration is in index order, as PHP's ksort'ed array. */
  pointColors: Map<number, string>;
}

export interface ChartSpec {
  kind: "bar" | "line" | "pie";
  horizontal: boolean;
  stacked: boolean;
  title: string;
  categories: string[];
  categoryTitle: string;
  valueTitle: string;
  legend: boolean;
  series: ChartSeriesSpec[];
}

function isObject(v: unknown): v is Record<string, Any> {
  return typeof v === "object" && v !== null && !Array.isArray(v);
}

function isNumber(v: unknown): v is number {
  return typeof v === "number" && Number.isFinite(v);
}

/** Translate an option into the writer's spec; null when nothing drawable can be derived. */
export function chartSpec(option: unknown): ChartSpec | null {
  if (!isObject(option)) return null;
  let raw: Any = option.series;
  if (isObject(raw)) raw = [raw];
  if (!Array.isArray(raw) || raw.length === 0) return null;

  let kind: ChartSpec["kind"] | null = null;
  for (const series of raw) {
    if (!isObject(series)) return null;
    const type = series.type ?? "bar";
    if (typeof type !== "string" || !(SUPPORTED_CHART_TYPES as readonly string[]).includes(type)) return null;
    kind ??= type as ChartSpec["kind"];
    if (type !== kind) return null;
  }
  if (kind === null) return null;
  if (kind === "pie" && raw.length !== 1) return null;

  const xAxis = axis(option.xAxis);
  const yAxis = axis(option.yAxis);
  const horizontal = kind === "bar" && yAxis.type === "category";
  const categoryAxis = horizontal ? yAxis : xAxis;
  const valueAxis = horizontal ? xAxis : yAxis;

  let categories: string[] = [];
  let source: Any = categoryAxis.data;
  if (!Array.isArray(source)) source = option.categories;
  if (Array.isArray(source)) {
    for (const item of source) {
      const text = chartLabel(item);
      if (text !== null) categories.push(text);
    }
  }

  const palette: string[] = [];
  for (const color of Array.isArray(option.color) ? option.color : []) {
    const h = chartHex(color);
    if (h !== null) palette.push(h);
  }

  if (kind === "pie") {
    // A pie's labels are its slice names, never an axis.
    categories = [];
  }

  const series: ChartSeriesSpec[] = [];
  for (let idx = 0; idx < raw.length; idx++) {
    const one = raw[idx];
    const data = one.data;
    if (!Array.isArray(data)) return null;
    const values: number[] = [];
    const pointColors = new Map<number, string>();
    for (let i = 0; i < data.length; i++) {
      const point = data[i];
      const value = pointValue(point);
      if (value === null) return null;
      values.push(value);
      const color = isObject(point) && isObject(point.itemStyle) ? chartHex(point.itemStyle.color) : null;
      if (color !== null) pointColors.set(i, color);
      if (kind === "pie" && idx === 0) {
        const name = chartLabel(isObject(point) ? point.name : null);
        categories[i] = name ?? `Slice ${i + 1}`;
      }
    }

    let colors = pointColors;
    if (kind === "pie") {
      // Every slice takes its own colour: a point's, the palette's, the default's.
      colors = new Map<number, string>();
      for (let i = 0; i < values.length; i++) {
        colors.set(i, pointColors.get(i) ?? (palette.length > 0 ? palette[i % palette.length]! : DEFAULT_PALETTE[i % DEFAULT_PALETTE.length]!));
      }
    }

    const own = isObject(one.itemStyle) ? chartHex(one.itemStyle.color) : null;
    const name = chartLabel(one.name);
    series.push({
      name: name !== null && name !== "" ? name : `Series ${idx + 1}`,
      values,
      color: own ?? (palette.length > 0 ? palette[idx % palette.length]! : DEFAULT_PALETTE[idx % DEFAULT_PALETTE.length]!),
      labels: isObject(one.label) && one.label.show === true,
      smooth: kind === "line" && one.smooth === true,
      pointColors: colors,
    });
  }

  let stacked = false;
  if (kind === "bar") {
    for (const one of raw) {
      if (typeof one.stack === "string" && one.stack !== "") stacked = true;
    }
  }

  const legend = option.legend;

  return {
    kind,
    horizontal,
    stacked,
    title: chartTitle(option.title),
    categories,
    categoryTitle: typeof categoryAxis.name === "string" ? categoryAxis.name : "",
    valueTitle: typeof valueAxis.name === "string" ? valueAxis.name : "",
    // PHP: is_array($legend) && show !== false -- so a JSON [] counts as present there too.
    legend: (isObject(legend) && legend.show !== false) || Array.isArray(legend),
    series,
  };
}

/** An axis object: the object itself, or the first of a list of them. */
function axis(value: unknown): Record<string, Any> {
  if (Array.isArray(value)) return isObject(value[0]) ? value[0] : {};
  return isObject(value) ? value : {};
}

/** `option.title` as text: `{text}`, a list of those, or a bare string. */
export function chartTitle(title: unknown): string {
  if (typeof title === "string") return title;
  const first = Array.isArray(title) ? title[0] : title;
  return isObject(first) && typeof first.text === "string" ? first.text : "";
}

/** A point's number: a bare number, or `{value: number}`. Null when it has none. */
export function pointValue(point: unknown): number | null {
  if (isNumber(point)) return point;
  if (isObject(point) && isNumber(point.value)) return point.value;
  return null;
}

/** `#RRGGBB` as upper-case hex without the '#', or null. */
export function chartHex(color: unknown): string | null {
  return typeof color === "string" && /^#[0-9A-Fa-f]{6}$/.test(color) ? color.slice(1).toUpperCase() : null;
}

/** A label as text: a string as it is, a number spelled by {@link chartNum}; null otherwise. */
export function chartLabel(value: unknown): string | null {
  if (typeof value === "string") return value;
  if (isNumber(value)) return chartNum(value);
  return null;
}

/**
 * A number as the chart caches and the workbook spell it: the SHORTEST decimal
 * that reads back as the same double, never in exponent form. The digits come
 * from `toExponential(p)` at the smallest precision that round-trips, then the
 * exponent is written out positionally -- the same bytes PHP and Python produce.
 */
export function chartNum(value: number): string {
  if (!Number.isFinite(value)) return "0";
  if (Number.isInteger(value) && Math.abs(value) < 1e15) return String(value === 0 ? 0 : value);
  let s = value.toExponential(16);
  for (let p = 0; p <= 16; p++) {
    const candidate = value.toExponential(p);
    if (Number(candidate) === value) {
      s = candidate;
      break;
    }
  }
  return positional(s);
}

/** `-1.2345e-7` → `-0.00000012345`. */
function positional(exp: string): string {
  const m = /^(-?)(\d)(?:\.(\d+))?e([+-]?\d+)$/.exec(exp);
  if (!m) return exp;
  const sign = m[1]!;
  const digits = (m[2]! + (m[3] ?? "")).replace(/0+$/, "");
  if (digits === "") return "0";
  const point = parseInt(m[4]!, 10) + 1;
  let out: string;
  if (point <= 0) out = "0." + "0".repeat(-point) + digits;
  else if (point >= digits.length) out = digits + "0".repeat(point - digits.length);
  else out = digits.slice(0, point) + "." + digits.slice(point);
  return sign + out;
}

/** A cached number back as the model holds it. Only a plain decimal counts; anything else is 0. */
export function chartParseNum(text: string): number {
  const t = text.trim();
  const v = /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(t) ? Number(t) : 0;
  return Number.isFinite(v) ? v + 0 : 0;
}

/** 0 → A, 25 → Z, 26 → AA. */
export function chartColumn(index: number): string {
  let out = "";
  let n = index + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}
