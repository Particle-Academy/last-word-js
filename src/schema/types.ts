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
  | "pageBreak"
  | "hr";

export type Align = "left" | "center" | "right" | "justify";

/** Inline text span. Only `text` is required; all flags optional. */
export interface Run {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  code?: boolean;
  link?: string;
  /** #RRGGBB */
  color?: string;
  /** #RRGGBB */
  highlight?: string;
  [key: string]: unknown;
}

export interface ListItem {
  runs: Run[];
  children?: ListItem[];
  [key: string]: unknown;
}

export interface HeadingBlock {
  type: "heading";
  level: 1 | 2 | 3 | 4 | 5 | 6;
  runs: Run[];
}

export interface ParagraphBlock {
  type: "paragraph";
  runs: Run[];
  align?: Align;
}

export interface ListBlock {
  type: "list";
  ordered?: boolean;
  items: ListItem[];
}

export interface TableCell {
  blocks: Block[];
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
  | PageBreakBlock
  | HrBlock;

export interface Doc {
  title?: string;
  blocks: Block[];
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
