/**
 * Public document types — the LastWord JSON document model. Inputs are loose
 * agent JSON, so most fields are optional and extra keys are tolerated; the
 * Validator is the gate. Mirrors PHP `particle-academy/last-word`.
 */

export type BlockType =
  | "heading"
  | "paragraph"
  | "list"
  | "table"
  | "code"
  | "quote"
  | "image"
  | "chart"
  | "pageBreak"
  | "hr";

export type Align = "left" | "center" | "right" | "justify";
export type VAlign = "top" | "center" | "bottom";
export type BorderStyle = "single" | "double" | "dashed" | "dotted" | "none";
export type PageSize = "letter" | "legal" | "a4";
export type Orientation = "portrait" | "landscape";

/**
 * One border edge.
 *
 * `{ style: "none" }` REMOVES a border — a zero width is not it, and a white
 * one only hides it against a white page.
 */
export interface Border {
  style?: BorderStyle;
  /** Points. Defaults to 0.5 (a hairline). */
  width?: number;
  /** #RRGGBB. Defaults to the document's automatic colour. */
  color?: string;
}

/** Box edges. Anything omitted is left alone rather than reset. */
export interface BoxBorders {
  top?: Border;
  right?: Border;
  bottom?: Border;
  left?: Border;
}

/** Table edges — the box, plus the two inside directions. */
export interface TableBorders extends BoxBorders {
  insideH?: Border;
  insideV?: Border;
}

/** Box spacing in points. Anything omitted is left alone. */
export interface BoxSides {
  top?: number;
  right?: number;
  bottom?: number;
  left?: number;
}

/** Inline text span. Only `text` is required; all flags optional. */
export interface Run {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  code?: boolean;
  smallCaps?: boolean;
  link?: string;
  /** #RRGGBB */
  color?: string;
  /** #RRGGBB */
  highlight?: string;
  /** Font size in points. Half-points are exactly representable. */
  size?: number;
  /** Font family name. */
  font?: string;
  /** Tracking in points; may be negative. Zero is the default and emits nothing. */
  letterSpacing?: number;
  /**
   * A footnote. The run carrying one has `text: ""` and sits where the
   * reference mark goes; Word numbers the marks 1, 2, 3 in document order, so
   * the model carries no number. Not in a header or footer, and never inside
   * another footnote.
   */
  footnote?: Footnote;
  [key: string]: unknown;
}

/**
 * A footnote's content: the body's block vocabulary less `chart` and
 * `pageBreak`, at any depth.
 */
export interface Footnote {
  blocks: Block[];
}

/**
 * Properties every paragraph-shaped block accepts — `paragraph`, `heading` and
 * a list item alike. Lengths are POINTS; `lineHeight` is a multiple.
 */
export interface ParagraphProps {
  align?: Align;
  spaceBefore?: number;
  spaceAfter?: number;
  lineHeight?: number;
  indentLeft?: number;
  indentRight?: number;
  /** Keep this block on the same page as the one after it. */
  keepNext?: boolean;
  /** #RRGGBB band behind the paragraph. */
  shading?: string;
  borders?: BoxBorders;
}

export interface ListItem extends ParagraphProps {
  runs: Run[];
  children?: ListItem[];
  [key: string]: unknown;
}

export interface HeadingBlock extends ParagraphProps {
  type: "heading";
  level: 1 | 2 | 3 | 4 | 5 | 6;
  runs: Run[];
}

export interface ParagraphBlock extends ParagraphProps {
  type: "paragraph";
  runs: Run[];
}

export interface ListBlock {
  type: "list";
  ordered?: boolean;
  items: ListItem[];
}

export interface TableCell {
  blocks: Block[];
  /** #RRGGBB fill. */
  shading?: string;
  borders?: BoxBorders;
  /** Padding in points, overriding the table's `cellPadding`. */
  padding?: BoxSides;
  valign?: VAlign;
  /** Columns this cell covers. */
  colSpan?: number;
  /**
   * Rows this cell covers. Written HTML-style: the cell appears ONCE, and the
   * rows it covers list only their own remaining cells.
   */
  rowSpan?: number;
  [key: string]: unknown;
}

export interface TableRow {
  header?: boolean;
  cells: TableCell[];
  [key: string]: unknown;
}

export interface TableBlock {
  type: "table";
  rows: TableRow[];
  /**
   * Relative column weights — `[30, 40, 30]` and `[3, 4, 3]` are the same
   * table. Also fixes the layout, so Word honours them instead of re-fitting
   * columns to their content.
   */
  widths?: number[];
  /** Table width as a percentage of the text column. */
  width?: number;
  align?: "left" | "center" | "right";
  borders?: TableBorders;
  /** Padding in points applied to every cell that does not set its own. */
  cellPadding?: BoxSides;
}

export interface CodeBlock {
  type: "code";
  language?: string;
  text: string;
}

export interface QuoteBlock {
  type: "quote";
  blocks: Block[];
}

export interface ImageBlock {
  type: "image";
  /** data:image/png;base64,… or data:image/jpeg;base64,… */
  src: string;
  widthPx?: number;
  heightPx?: number;
  alt?: string;
}

/** One ECharts series, as a document chart reads it. */
export interface ChartSeries {
  /** bar (columns over a category xAxis, horizontal bars over a category yAxis), line or pie; bar when absent. */
  type?: "bar" | "line" | "pie";
  name?: string;
  /** Numbers, or `{value, itemStyle?}`; a pie's points are `{name, value, itemStyle?}`. */
  data: Array<number | { value: number; name?: string; itemStyle?: { color?: string } }>;
  itemStyle?: { color?: string };
  label?: { show?: boolean };
  smooth?: boolean;
  stack?: string;
  [key: string]: unknown;
}

export interface ChartAxis {
  type?: "category" | "value";
  data?: Array<string | number>;
  /** The axis title. */
  name?: string;
  [key: string]: unknown;
}

/**
 * The ECharts option a chart is drawn from -- the same field dark-slide's chart
 * element takes. Open to any other ECharts key; the writer reads the ones typed.
 */
export interface ChartOption {
  series: ChartSeries[] | ChartSeries;
  xAxis?: ChartAxis | ChartAxis[];
  yAxis?: ChartAxis | ChartAxis[];
  title?: { text?: string } | string;
  legend?: { show?: boolean };
  color?: string[];
  [key: string]: unknown;
}

/** A native Word chart, with its numbers in an embedded workbook. */
export interface ChartBlock {
  type: "chart";
  option: ChartOption;
  /** Px at 96 dpi; capped at the text column. */
  widthPx?: number;
  /** Px at 96 dpi; 0.6 of the width when absent. */
  heightPx?: number;
  /** Width as a percentage of the text column, when widthPx is absent. */
  width?: number;
  /** An Excel number format: "#,##0", "0.0%". */
  numberFormat?: string;
  alt?: string;
}

export interface PageBreakBlock {
  type: "pageBreak";
}

export interface HrBlock {
  type: "hr";
}

export type Block =
  | HeadingBlock
  | ParagraphBlock
  | ListBlock
  | TableBlock
  | CodeBlock
  | QuoteBlock
  | ImageBlock
  | ChartBlock
  | PageBreakBlock
  | HrBlock;

/**
 * Page margins in points, 72 (one inch) a side by default, and the header and
 * footer DISTANCES: points from the top edge of the page to the header and from
 * the bottom edge to the footer, 36 by default. `read()` omits any value equal
 * to its default; resolve them with `Schema.pageMargins()`.
 */
export interface PageMargins extends BoxSides {
  header?: number;
  footer?: number;
}

/** Every margin and both distances, defaults filled in: `Schema.pageMargins()`. */
export interface ResolvedPageMargins {
  top: number;
  right: number;
  bottom: number;
  left: number;
  header: number;
  footer: number;
}

/** Section geometry. Margins are POINTS. */
export interface Page {
  size?: PageSize;
  orientation?: Orientation;
  margins?: PageMargins;
}

export interface Doc {
  title?: string;
  blocks: Block[];
  page?: Page;
  /** Font every run inherits unless it names its own. */
  defaultFont?: string;
  /** Size in points every run inherits unless it names its own. */
  defaultSize?: number;
  [key: string]: unknown;
}

/** Structured validation error. Empty array = valid. */
export interface ValidationError {
  path: string;
  message: string;
}

export interface RepairResult {
  ok: boolean;
  schema: Doc;
  errors: ValidationError[];
}

export interface WriteResult {
  path: string;
  bytes: number;
  blocks: number;
}
