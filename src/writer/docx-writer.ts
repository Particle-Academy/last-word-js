/**
 * DocxWriter — Doc JSON model → .docx bytes (OOXML / WordprocessingML).
 * Deterministic: no timestamps, fixed zip entry order, rel ids assigned in
 * traversal order. Mirrors PHP `Writer\DocxWriter`.
 */

import { Xml } from "../helpers/xml";
import { DocxTemplate } from "./docx-template";
import { parseDataUrl, sniffImageSize } from "../helpers/image-size";
import type { Block, Doc, ListItem, Run } from "../schema/types";
import { Schema } from "../schema/schema";
import { zipSync, type ZipFile } from "../zip";
import { isPlainObject } from "../util";
import { DEFAULT_ASPECT, chartSpec, chartTitle } from "../helpers/chart";
import { chartRelsXml, chartXml, workbookParts } from "./chart-part";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

/** EMUs per pixel at 96 dpi. */
export const EMU_PER_PX = 9525;
/** Content width cap: 6.5in at 96 dpi. */
export const MAX_WIDTH_PX = 624;

const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const REL_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_NUMBERING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
const REL_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
const REL_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const REL_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
/** EMU per twip: the text column is measured in twips, a drawing in EMU. */
const EMU_PER_TWIP = 635;
/**
 * A chart's drawing id is offset from the images' so the two can never collide;
 * Word treats a repeated wp:docPr id as a damaged document. Shared with PHP.
 */
const CHART_DOCPR_BASE = 10000;
const REL_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_CORE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
const REL_SETTINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
const REL_FOOTNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";

/** The mark that opens a footnote's own text: Word's FootnoteReference-styled w:footnoteRef. Shared with PHP. */
const FOOTNOTE_MARK = `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>`;

/**
 * Block types a footnote's blocks refuse, at any depth (`Schema.FOOTNOTE_REFUSED_TYPES`):
 * validation refuses them, and unvalidated they are written as nothing.
 */
const FOOTNOTE_REFUSED_TYPES: readonly string[] = ["chart", "pageBreak", "portable"];

/** A run carrying a `footnote` key, as PHP's `array_key_exists('footnote', $run)`. */
function isFootnoteRun(run: unknown): boolean {
  return isPlainObject(run) && Object.prototype.hasOwnProperty.call(run, "footnote") && (run as Any).footnote !== undefined;
}
/** header / footer relationship types are `${REL_BASE}/header` and `${REL_BASE}/footer`. */
const REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const NUM_ID_BULLET = 1;
const NUM_ID_DECIMAL = 2;
const MAX_ILVL = 5;

/** Shaded fill for header cells. Shared with the PHP and Python engines. */
const HEADER_FILL = "E7E7E7";

/** Page sizes in twips (1/1440in), portrait. */
export const PAGE_SIZES: Record<string, [number, number]> = {
  letter: [12240, 15840],
  legal: [12240, 20160],
  a4: [11906, 16838],
};

/**
 * Cell margins every table gets unless it says otherwise, in POINTS —
 * 3 / 5.4 / 3 / 5.4, which is 60 / 108 / 60 / 108 twips. 108 is the value Word
 * itself uses for default side margins, which is why it reads as an odd number
 * rather than a round one.
 */
const DEFAULT_CELL_MARGINS_PT = { top: 3, left: 5.4, bottom: 3, right: 5.4 };

/** The border every table edge gets unless it says otherwise. */
const DEFAULT_BORDER = { style: "single", width: 0.5, color: undefined } as const;

const TABLE_EDGES = ["top", "left", "bottom", "right", "insideH", "insideV"] as const;
const BOX_EDGES = ["top", "left", "bottom", "right"] as const;

/** SDT tag prefixes used to round-trip block metadata that OOXML has no slot for. */
export const SDT_TAG_CODE = "lastword:code";
export const SDT_TAG_QUOTE = "lastword:quote";

interface MediaFile {
  name: string; // e.g. image1.png
  bytes: Uint8Array;
  ext: string;
}

interface Rel {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

interface RenderCtx {
  /** pStyle applied to plain paragraphs (used inside quote SDTs). */
  paragraphStyle?: string;
}

function normalizeHex(hex: string): string {
  return hex.replace(/^#/, "").toUpperCase();
}

// ── Units and property fragments ────────────────────────────────────────────
//
// WordprocessingML measures four different things in four different units, and
// getting one wrong produces a document that opens fine and is the wrong size.
// They are collected here, once, so the three engines can be compared line for
// line:
//
//   points → TWENTIETHS of a point (twips)  spacing, indents, margins
//   points → HALF-points                    font size
//   points → EIGHTHS of a point             border width
//   percent → FIFTIETHS of a percent        table width
//
// Suite: fancy-conformance `last-word/docx-constructs`.

const isNum = (v: unknown): v is number => typeof v === "number" && Number.isFinite(v);

/** Points → twips. */
function twips(pt: unknown): number | null {
  return isNum(pt) ? Math.round(pt * 20) : null;
}

/** Points → half-points (w:sz on a run). */
function halfPoints(pt: unknown): number | null {
  return isNum(pt) && pt > 0 ? Math.round(pt * 2) : null;
}

/** #RRGGBB → RRGGBB, upper-cased. Anything else is null. */
function hex(value: unknown): string | null {
  return typeof value === "string" && /^#[0-9a-fA-F]{6}$/.test(value) ? normalizeHex(value) : null;
}

/** A running header / footer, normalised. */
interface Band {
  left: string;
  center: string;
  right: string;
  size: number;
  color: string;
}

/** A band spec, normalised; null when it asks for nothing. Mirrors PHP DocxWriter::band(). */
function band(spec: unknown): Band | null {
  if (!isPlainObject(spec) && !Array.isArray(spec)) return null;
  const s = spec as Any;
  const slot = (k: string): string => (typeof s[k] === "string" ? s[k] : "");
  const left = slot("left");
  const center = slot("center");
  const right = slot("right");
  if (left === "" && center === "" && right === "") return null;
  const size = isNum(s.size) && s.size > 0 ? s.size : 9;
  return { left, center, right, size, color: hex(s.color) ?? "6B7280" };
}

/** PHP's trim(): only " \t\n\r\0\x0B", not JavaScript's wider whitespace. */
function phpTrim(s: string): string {
  return s.replace(/^[ \t\n\r\0\x0B]+|[ \t\n\r\0\x0B]+$/g, "");
}

/** A band slot key's Word reference type and the part number it is written as. */
const BAND_SLOTS = [
  ["default", 1],
  ["first", 2],
  ["even", 3],
] as const;

type BandKind = "header" | "footer";
type BandType = (typeof BAND_SLOTS)[number][0];

const capitalise = (s: string): string => s.charAt(0).toUpperCase() + s.slice(1);

/** `<w:shd>` — the one spelling used for runs, paragraphs and cells alike. */
function shadingXml(color: unknown): string {
  const fill = hex(color);
  return fill === null ? "" : `<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`;
}

/**
 * One border edge.
 *
 * `style: none` becomes `w:val="nil"` with no width and no colour, because nil
 * is the only way to REMOVE a border — a zero width is not it, and a white one
 * only hides it against a white page.
 */
function borderEdgeXml(tag: string, border: Any): string {
  const style = typeof border?.style === "string" ? border.style : "single";
  if (style === "none") return `<w:${tag} w:val="nil"/>`;
  const width = isNum(border?.width) ? border.width : 0.5;
  const sz = Math.max(2, Math.min(96, Math.round(width * 8)));
  const color = hex(border?.color) ?? "auto";
  return `<w:${tag} w:val="${style}" w:sz="${sz}" w:space="0" w:color="${color}"/>`;
}

/**
 * A border container (`w:pBdr`, `w:tblBorders`, `w:tcBorders`) in the edge
 * order its CT_ type declares. Absent edges are omitted, so a partial
 * `borders` stays partial.
 */
function bordersXml(wrapper: string, borders: Any, edges: readonly string[]): string {
  let inner = "";
  for (const edge of edges) {
    const spec = borders?.[edge];
    if (spec && typeof spec === "object") inner += borderEdgeXml(edge, spec);
  }
  return inner === "" ? "" : `<w:${wrapper}>${inner}</w:${wrapper}>`;
}

/**
 * A margin container (`w:tblCellMar`, `w:tcMar`). Sides are twips and a side
 * that was not given is not emitted.
 */
function marginsXml(wrapper: string, sides: Any): string {
  let inner = "";
  for (const side of BOX_EDGES) {
    const t = twips(sides?.[side]);
    if (t !== null) inner += `<w:${side} w:w="${t}" w:type="dxa"/>`;
  }
  return inner === "" ? "" : `<w:${wrapper}>${inner}</w:${wrapper}>`;
}

interface PageGeometry {
  w: number;
  h: number;
  top: number;
  right: number;
  bottom: number;
  left: number;
  orientation: "portrait" | "landscape";
  /** The header / footer distances (`w:pgMar` header / footer), in twips. */
  header: number;
  footer: number;
}

/** Page size, margins and the header / footer distances, in twips. */
export function pageGeometry(page: Any): PageGeometry {
  const size = typeof page?.size === "string" ? page.size.toLowerCase() : "letter";
  let [w, h] = PAGE_SIZES[size] ?? PAGE_SIZES.letter!;

  const orientation = page?.orientation === "landscape" ? "landscape" : "portrait";
  // Swapping the axes without w:orient gives a page that is the right shape
  // and prints portrait. Both are required.
  if (orientation === "landscape") [w, h] = [h, w];

  const side = (k: string): number => twips(page?.margins?.[k]) ?? 1440;
  // The header / footer distances are `margins.header` / `margins.footer`
  // (Word keeps them on w:pgMar too). Unsigned in OOXML (ST_TwipsMeasure): an
  // unvalidated negative is 0. Mirrors PHP `pageGeometry()`.
  const distance = (k: string): number => {
    const own = twips(page?.margins?.[k]);
    return own !== null ? Math.max(0, own) : 720;
  };

  return {
    w,
    h,
    top: side("top"),
    right: side("right"),
    bottom: side("bottom"),
    left: side("left"),
    orientation,
    header: distance("header"),
    footer: distance("footer"),
  };
}

interface Slot {
  cell: Any;
  span: number;
  vMerge: "restart" | "continue" | null;
  shading: unknown;
}

/**
 * Lay a table's authored rows onto a grid, resolving both merge directions.
 *
 * The author writes cells HTML-style: a `rowSpan` cell appears ONCE, and the
 * rows it covers list only their own remaining cells. OOXML has no such
 * shorthand — every row must carry a cell for every grid column, and a row that
 * is short is a malformed table Word repairs by shifting everything left. So
 * the covered rows get a synthesised `w:vMerge` continuation here.
 */
export function layoutRows(rows: Any[]): { laid: Slot[][]; colCount: number } {
  const pending = new Map<number, { rows: number; span: number; shading: unknown }>();
  const laid: Slot[][] = [];
  let colCount = 1;

  for (const row of rows) {
    const authored: Any[] = Array.isArray(row?.cells) ? row.cells.filter((c: Any) => c && typeof c === "object") : [];
    const line: Slot[] = [];
    let col = 0;
    let next = 0;

    for (;;) {
      const held = pending.get(col);
      if (held && held.rows > 0) {
        // A continuation carries the origin's shading and nothing else:
        // without the fill the merged block renders striped, and with the
        // origin's borders it would draw a rule straight through its own
        // middle.
        line.push({ cell: { blocks: [] }, span: held.span, vMerge: "continue", shading: held.shading });
        held.rows--;
        col += held.span;
        continue;
      }

      if (next < authored.length) {
        const cell = authored[next++];
        const span = Math.max(1, Math.trunc(Number(cell.colSpan) || 1));
        const rowSpan = Math.max(1, Math.trunc(Number(cell.rowSpan) || 1));
        line.push({ cell, span, vMerge: rowSpan > 1 ? "restart" : null, shading: cell.shading ?? null });
        if (rowSpan > 1) pending.set(col, { rows: rowSpan - 1, span, shading: cell.shading ?? null });
        col += span;
        continue;
      }

      // Nothing authored left — but a merge started further right still owes
      // this row a continuation.
      let ahead: number | null = null;
      for (const [at, held2] of pending) {
        if (at > col && held2.rows > 0 && (ahead === null || at < ahead)) ahead = at;
      }
      if (ahead === null) break;
      col = ahead;
    }

    colCount = Math.max(colCount, col);
    laid.push(line);
  }

  return { laid, colCount };
}

/** Header cells bold their runs — the shape PHP and Python have always used. */
function boldRuns(blocks: Any[]): Any[] {
  return blocks.map((block: Any) =>
    block && (block.type === "paragraph" || block.type === "heading")
      ? { ...block, runs: (block.runs ?? []).map((r: Any) => (r && typeof r === "object" ? { ...r, bold: true } : r)) }
      : block,
  );
}

/**
 * Split a width into `count` columns by relative weight, giving any rounding
 * remainder to the LAST column so the grid sums to the content width exactly.
 * Three engines rounding independently is how one language ends up with a
 * table a twip narrower than the other two.
 */
export function splitColumns(total: number, count: number, weights?: number[] | null): number[] {
  if (count < 1) return [];
  let w = weights && weights.length === count ? weights : new Array<number>(count).fill(1);
  let sum = w.reduce((a, b) => a + b, 0);
  if (!(sum > 0)) {
    w = new Array<number>(count).fill(1);
    sum = count;
  }

  const out: number[] = [];
  let used = 0;
  for (let i = 0; i < count - 1; i++) {
    const width = Math.round(total * (w[i]! / sum));
    out.push(width);
    used += width;
  }
  out.push(total - used);
  return out;
}

export class DocxWriter {
  private rels: Rel[] = [];
  private hyperlinkIds = new Map<string, string>();
  private media: MediaFile[] = [];
  private drawingId = 0;
  /** Chart parts queued for the archive, in chart order (chart N is index N-1). */
  private charts: { xml: string; rels: string; workbook: [string, string][] }[] = [];
  /** Twips between the page margins; set from `page` before any block renders. */
  private contentWidth = 9360;

  /**
   * The document's own running headers and footers, as finished parts, in
   * write order: `name` is `{kind}{n}`, numbered by SLOT -- 1 the default
   * (`page.header`), 2 the first page (`page.firstHeader`), 3 even pages
   * (`page.evenHeader`) -- written as `word/{name}.xml` with relationship id
   * `rId{Kind}{n}`. Empty when none were asked for, and then nothing about the
   * package changes. Mirrors PHP `DocxWriter::$bands` (0.16.0).
   */
  private bands: { name: string; kind: BandKind; xml: string }[] = [];

  /** `w:titlePg`: the document has a first-page band of its own. */
  private titlePg = false;

  /** `w:evenAndOddHeaders` (settings): the document has an even-page band of its own. */
  private evenAndOdd = false;

  /** The cached result an AUTHOR field gets: `meta.author`, or "". */
  private fieldAuthor = "";

  /** The cached result a TITLE field gets: the document `title`, or "". */
  private fieldTitle = "";

  /**
   * The document's footnotes: `w:footnote` content (without the wrapper),
   * numbered 1, 2, 3 in document order as their runs render (footnote N is
   * index N-1). The separator (-1) and continuation separator (0) Word requires
   * are written around them. Empty without a footnote, and then nothing about
   * the package changes.
   */
  private footnotes: string[] = [];

  /**
   * word/footnotes.xml's OWN relationships (its links and images), written as
   * `word/_rels/footnotes.xml.rels`, with the hyperlink ids they dedupe on.
   */
  private footnoteRels: Rel[] = [];
  private footnoteHyperlinkIds = new Map<string, string>();

  /** True while a footnote's blocks render: a footnote run inside one renders as nothing. */
  private inFootnote = false;

  /**
   * The next NUMBERED relationship id. Counted over the numbered rels only: the
   * band rels (rIdHeader1 / rIdFooter1) are named, and counting them shifted
   * every hyperlink and image id two places away from the PHP engine's.
   */
  private nextRelId(): string {
    return `rId${this.rels.filter((r) => /^rId\d+$/.test(r.id)).length + 1}`;
  }

  /**
   * One paragraph: left, TAB, centre, TAB, right, against a centre tab stop at
   * half the content width and a right one at its edge.
   */
  private bandXml(kind: BandKind, spec: Band): string {
    const rPr = `<w:rPr><w:color w:val="${spec.color}"/><w:sz w:val="${halfPoints(spec.size)}"/></w:rPr>`;
    const tag = kind === "header" ? "w:hdr" : "w:ftr";
    return (
      Xml.declaration() +
      `<${tag} xmlns:w="${NS_W}" xmlns:r="${NS_R}">` +
      `<w:p><w:pPr><w:tabs>` +
      `<w:tab w:val="center" w:pos="${Math.trunc(this.contentWidth / 2)}"/>` +
      `<w:tab w:val="right" w:pos="${this.contentWidth}"/>` +
      `</w:tabs><w:spacing w:before="0" w:after="0"/></w:pPr>` +
      this.bandRuns(spec.left, rPr) +
      `<w:r>${rPr}<w:tab/></w:r>` +
      this.bandRuns(spec.center, rPr) +
      `<w:r>${rPr}<w:tab/></w:r>` +
      this.bandRuns(spec.right, rPr) +
      `</w:p>` +
      `</${tag}>`
    );
  }

  /**
   * Text with its field tokens (`Schema.FIELD_TOKEN`) as live fields: `{page}`
   * PAGE, `{pages}` NUMPAGES, `{date:yyyy-MM-dd}` DATE with its picture, and the
   * rest. Anything else is literal text. COMPLEX fields, so every piece carries
   * the band's formatting (see PHP `DocxWriter::bandRuns()`).
   */
  private bandRuns(text: string, rPr: string): string {
    let out = "";
    for (const piece of Schema.splitFieldTokens(text)) {
      out +=
        typeof piece === "string"
          ? `<w:r>${rPr}<w:t xml:space="preserve">${Xml.text(piece)}</w:t></w:r>`
          : this.fieldRuns(piece.name, piece.instruction, `<w:r>${rPr}`);
    }
    return out;
  }

  /**
   * One field token as a complex field, every piece opened with `open` (`<w:r>`
   * and the run's properties): begin, the instruction between single spaces
   * (` DATE \@ "yyyy-MM-dd" `), separate, the cached result, end.
   *
   * The cached result is the value when the writer knows it and nothing when
   * it does not -- never a guess, and never the clock, which would make the same
   * document write different bytes: `1` for PAGE, NUMPAGES, SECTION and
   * SECTIONPAGES; the author and the title for AUTHOR and TITLE (empty without
   * one); and no result run at all for DATE, TIME and FILENAME. Mirrors PHP
   * `DocxWriter::fieldRuns()`.
   */
  private fieldRuns(name: string, instruction: string, open: string): string {
    const result =
      name === "page" || name === "pages" || name === "section" || name === "sectionpages"
        ? "1"
        : name === "author"
          ? this.fieldAuthor
          : name === "title"
            ? this.fieldTitle
            : "";
    const preserve = result !== phpTrim(result) ? ` xml:space="preserve"` : "";
    return (
      `${open}<w:fldChar w:fldCharType="begin"/></w:r>` +
      `${open}<w:instrText xml:space="preserve"> ${Xml.text(instruction)} </w:instrText></w:r>` +
      `${open}<w:fldChar w:fldCharType="separate"/></w:r>` +
      (result !== "" ? `${open}<w:t${preserve}>${Xml.text(result)}</w:t></w:r>` : "") +
      `${open}<w:fldChar w:fldCharType="end"/></w:r>`
    );
  }

  /**
   * An empty header or footer part: a first or even page with no band, when the
   * document's own default band would otherwise show there. Byte-identical to
   * PHP's `blankBandXml()` (no `r` namespace: it has nothing to reference).
   */
  private static blankBandXml(kind: BandKind): string {
    const tag = kind === "header" ? "w:hdr" : "w:ftr";
    return Xml.declaration() + `<${tag} xmlns:w="${NS_W}"><w:p/></${tag}>`;
  }

  /**
   * The document's own header and footer slots (PHP 0.16.0): parts by slot,
   * their relationships, `w:titlePg` / `w:evenAndOddHeaders`, and the section's
   * references -- default, first, even within a kind, header before footer.
   *
   * An empty first or even band (`{"blocks": []}`, or a band with no text)
   * blanks that page when the kind has a default band, and is nothing
   * otherwise. With a switch on, a kind with no part for those pages references
   * its default there, so a footer is never lost from page one by adding a
   * first-page header. A `{blocks}` band with content is a band this engine does
   * not write (PHP 0.14.0, last-word#6), and asks for nothing here, as it always
   * has for `header` / `footer`.
   */
  private buildBands(pageSpec: Any): string {
    this.bands = [];
    const own: Record<BandKind, Partial<Record<BandType, string>>> = { header: {}, footer: {} };
    for (const kind of ["header", "footer"] as const) {
      for (const [type, n] of BAND_SLOTS) {
        const key = type === "default" ? kind : `${type}${capitalise(kind)}`;
        const raw = pageSpec[key];
        const spec = band(raw);
        let xml: string | null = spec !== null ? this.bandXml(kind, spec) : null;
        const blocks = isPlainObject(raw) ? (raw as Any).blocks : undefined;
        const asksForNothing = (isPlainObject(raw) || Array.isArray(raw)) && !(Array.isArray(blocks) && blocks.length > 0);
        if (xml === null && type !== "default" && asksForNothing && own[kind].default !== undefined) {
          xml = DocxWriter.blankBandXml(kind);
        }
        if (xml !== null) {
          const name = `${kind}${n}`;
          const id = `rId${capitalise(name)}`;
          this.bands.push({ name, kind, xml });
          this.rels.push({ id, type: `${REL_BASE}/${kind}`, target: `${name}.xml`, external: false });
          own[kind][type] = id;
        }
      }
    }

    this.titlePg = own.header.first !== undefined || own.footer.first !== undefined;
    this.evenAndOdd = own.header.even !== undefined || own.footer.even !== undefined;

    let references = "";
    for (const kind of ["header", "footer"] as const) {
      const slots = { ...own[kind] };
      if (this.titlePg && slots.first === undefined && slots.default !== undefined) slots.first = slots.default;
      if (this.evenAndOdd && slots.even === undefined && slots.default !== undefined) slots.even = slots.default;
      for (const [type] of BAND_SLOTS) {
        if (slots[type] !== undefined) references += `<w:${kind}Reference w:type="${type}" r:id="${slots[type]}"/>`;
      }
    }
    return references;
  }

  /** settings.xml is written only when something needs it: even/odd headers, or footnotes. */
  private hasSettings(): boolean {
    return this.evenAndOdd || this.footnotes.length > 0;
  }

  /**
   * @param template a house template whose styles + theme this document renders
   *   onto (last-word#3). Undefined is the built-in look, byte-for-byte as before.
   */
  constructor(private readonly template?: DocxTemplate) {}

  toBytes(doc: Doc): Uint8Array {
    this.rels = [
      { id: "rId1", type: REL_STYLES, target: "styles.xml", external: false },
      { id: "rId2", type: REL_NUMBERING, target: "numbering.xml", external: false },
    ];
    if (this.template?.hasTheme() === true) {
      // The theme travels WITH the styles. A style naming a theme colour or font
      // ("accent1", "majorHAnsi") resolves against whatever theme is in the
      // package, so styles without their theme give the template's structure in
      // the default's colours.
      this.rels.push({ id: "rId3", type: REL_THEME, target: "theme/theme1.xml", external: false });
    }
    this.hyperlinkIds = new Map();
    this.media = [];
    this.drawingId = 0;
    this.charts = [];
    this.footnotes = [];
    this.footnoteRels = [];
    this.footnoteHyperlinkIds = new Map();
    this.inFootnote = false;

    // Table grids are laid out against the section's content width, so it has
    // to be known before any block renders. Carrying 9360 as a literal — which
    // all three engines did — silently gives a document with narrowed margins
    // a table that no longer matches its own page.
    const page = pageGeometry((doc as Any).page);
    this.contentWidth = page.w - page.left - page.right;

    // meta.author (PHP 0.16.0) is dc:creator, beside the title in the core
    // part; a document with neither writes no core part, as before. Both are
    // also what AUTHOR / TITLE fields in a band show until Word updates them.
    const meta = (doc as Any).meta;
    const author = isPlainObject(meta) && typeof (meta as Any).author === "string" && (meta as Any).author !== "" ? ((meta as Any).author as string) : null;
    const hasTitle = doc.title !== undefined;
    const hasCore = hasTitle || author !== null;
    this.fieldAuthor = author ?? "";
    this.fieldTitle = typeof doc.title === "string" ? doc.title : "";

    // Running headers / footers: built before any block, so their rels sit after
    // styles/numbering/theme and before every hyperlink and image -- the order
    // the PHP engine writes them in.
    const pageSpec = isPlainObject((doc as Any).page) ? (doc as Any).page : {};
    const references = this.buildBands(pageSpec);

    const body = this.renderBlocks(doc.blocks ?? [], {});
    const documentXml =
      Xml.declaration() +
      `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}" xmlns:wp="${NS_WP}" xmlns:a="${NS_A}" xmlns:pic="${NS_PIC}">` +
      `<w:body>${body}${this.sectPr(page, references)}</w:body></w:document>`;

    const files: ZipFile[] = [
      { name: "[Content_Types].xml", data: encode(this.contentTypesXml(hasCore)) },
      { name: "_rels/.rels", data: encode(this.packageRelsXml(hasCore)) },
    ];
    if (hasCore) {
      files.push({ name: "docProps/core.xml", data: encode(this.coreXml(hasTitle ? String(doc.title) : null, author)) });
    }
    files.push(
      { name: "word/document.xml", data: encode(documentXml) },
      { name: "word/styles.xml", data: encode(this.stylesXml(doc)) },
      { name: "word/numbering.xml", data: encode(this.numberingXml()) },
    );
    if (this.hasSettings()) {
      files.push({ name: "word/settings.xml", data: encode(this.settingsXml()) });
    }
    if (this.template?.hasTheme() === true) {
      files.push({ name: "word/theme/theme1.xml", data: encode(this.template.theme()!) });
    }
    for (const b of this.bands) {
      files.push({ name: `word/${b.name}.xml`, data: encode(b.xml) });
    }
    if (this.footnotes.length > 0) {
      files.push({ name: "word/footnotes.xml", data: encode(this.footnotesXml()) });
      if (this.footnoteRels.length > 0) {
        files.push({ name: "word/_rels/footnotes.xml.rels", data: encode(this.relsXml(this.footnoteRels)) });
      }
    }
    files.push({ name: "word/_rels/document.xml.rels", data: encode(this.documentRelsXml()) });
    for (const m of this.media) {
      files.push({ name: `word/media/${m.name}`, data: m.bytes });
    }
    this.charts.forEach((chart, i) => {
      const n = i + 1;
      files.push(
        { name: `word/charts/chart${n}.xml`, data: encode(chart.xml) },
        { name: `word/charts/_rels/chart${n}.xml.rels`, data: encode(chart.rels) },
        {
          name: `word/embeddings/Microsoft_Excel_Worksheet${n}.xlsx`,
          data: zipSync(chart.workbook.map(([name, xml]) => ({ name, data: encode(xml) }))),
        },
      );
    });
    return zipSync(files);
  }

  // ── Blocks ────────────────────────────────────────────────────────────────

  private renderBlocks(blocks: Block[], ctx: RenderCtx): string {
    let out = "";
    let prevWasTable = false;
    for (const block of blocks) {
      const type = (block as Any)?.type;
      // OOXML MERGES adjacent tables into one — pad with an empty paragraph.
      // Not cosmetic: without it a stat band followed by a callout becomes a
      // single two-row table in Word, with the first table's column grid
      // imposed on the second. PHP and Python have always done this; the port
      // did not, and its own reference document has adjacent tables.
      if (this.inFootnote && FOOTNOTE_REFUSED_TYPES.includes(type)) {
        continue; // validation refuses these in a footnote; unvalidated, they are nothing
      }
      if (type === "table" && prevWasTable) out += `<w:p/>`;
      out += this.renderBlock(block as Any, ctx);
      prevWasTable = type === "table";
    }
    return out;
  }

  private renderBlock(block: Any, ctx: RenderCtx): string {
    switch (block.type) {
      case "heading":
        // A heading is a paragraph and takes the same properties. Without
        // that, a section label that needed spacing or alignment had to be a
        // bold paragraph impersonating a heading — and so appeared in no
        // navigation pane and no table of contents.
        return this.paragraph(this.pPr(block, `Heading${clampLevel(block.level)}`), this.renderRuns(block.runs ?? []));
      case "paragraph":
        return this.paragraph(this.pPr(block, ctx.paragraphStyle ?? null), this.renderRuns(block.runs ?? []));
      case "list":
        return this.renderList(block.items ?? [], block.ordered === true, 0);
      case "table":
        return this.renderTable(block);
      case "code":
        return this.renderCode(block);
      case "quote":
        return this.renderQuote(block);
      case "image":
        return this.renderImage(block);
      case "chart":
        return this.renderChart(block);
      case "pageBreak":
        return `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
      case "hr":
        return (
          `<w:p><w:pPr><w:pBdr>` +
          `<w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/>` +
          `</w:pBdr></w:pPr></w:p>`
        );
      default:
        return "";
    }
  }

  private paragraph(pPr: string, runs: string): string {
    return `<w:p>${pPr}${runs}</w:p>`;
  }

  /**
   * Paragraph properties, in CT_PPr order:
   * pStyle, keepNext, numPr, pBdr, shd, spacing, ind, jc, outlineLvl.
   */
  private pPr(block: Any, style: string | null, numPr: string | null = null): string {
    let inner = "";
    if (style) inner += `<w:pStyle w:val="${Xml.attr(style)}"/>`;
    if (block?.keepNext === true) inner += `<w:keepNext/>`;
    if (numPr) inner += numPr;

    if (block?.borders && typeof block.borders === "object") {
      inner += bordersXml("pBdr", block.borders, BOX_EDGES);
    }
    inner += shadingXml(block?.shading);

    // before, after, line and lineRule all live on ONE w:spacing element;
    // emitting two would be invalid.
    let spacing = "";
    const before = twips(block?.spaceBefore);
    if (before !== null) spacing += ` w:before="${before}"`;
    const after = twips(block?.spaceAfter);
    if (after !== null) spacing += ` w:after="${after}"`;
    if (isNum(block?.lineHeight) && block.lineHeight > 0) {
      spacing += ` w:line="${Math.round(block.lineHeight * 240)}" w:lineRule="auto"`;
    }
    if (spacing !== "") inner += `<w:spacing${spacing}/>`;

    let ind = "";
    const left = twips(block?.indentLeft);
    if (left !== null) ind += ` w:left="${left}"`;
    const right = twips(block?.indentRight);
    if (right !== null) ind += ` w:right="${right}"`;
    if (ind !== "") inner += `<w:ind${ind}/>`;

    const align = block?.align;
    if (typeof align === "string" && align !== "left") {
      const jc = align === "justify" ? "both" : align === "center" || align === "right" ? align : null;
      if (jc) inner += `<w:jc w:val="${jc}"/>`;
    }

    return inner === "" ? "" : `<w:pPr>${inner}</w:pPr>`;
  }

  // ── Runs ──────────────────────────────────────────────────────────────────

  private renderRuns(runs: Run[]): string {
    let out = "";
    for (const run of runs) {
      if (isFootnoteRun(run)) {
        // The mark only: a footnote run carries no text, so no link either.
        // Unvalidated text typed on the mark is written before it rather than
        // lost, which is what repair does.
        if (typeof run.text === "string" && run.text !== "") {
          const { footnote: _note, ...plain } = run as Any;
          out += this.renderRuns([plain as Run]);
        }
        out += this.renderFootnoteRun(run);
        continue;
      }
      if (typeof run?.text !== "string") continue;
      if (run.link) {
        const rid = this.hyperlinkRel(run.link);
        out += `<w:hyperlink r:id="${rid}" w:history="1">${this.renderRun(run, true)}</w:hyperlink>`;
      } else {
        out += this.renderRun(run, false);
      }
    }
    return out;
  }

  private renderRun(run: Run, linked: boolean): string {
    // rPr children in CT_RPr schema order — it is an xsd:sequence, so this is
    // the schema's order and not a preference:
    // rStyle, rFonts, b, i, smallCaps, strike, color, spacing, sz, szCs, u, shd
    let rPr = "";
    let font = typeof run.font === "string" && run.font !== "" ? run.font : null;
    if (run.code) {
      // `code` wins over an explicit font: it is the more specific request.
      rPr += `<w:rStyle w:val="InlineCode"/>`;
      font = "Consolas";
    } else if (linked) {
      rPr += `<w:rStyle w:val="Hyperlink"/>`;
    }
    if (font !== null) {
      // Three attributes, not one: with only w:ascii, Word picks its own face
      // for anything it classes as high-ANSI or complex-script and one run
      // renders in two fonts.
      const f = Xml.attr(font);
      rPr += `<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:cs="${f}"/>`;
    }
    if (run.bold) rPr += `<w:b/>`;
    if (run.italic) rPr += `<w:i/>`;
    if (run.smallCaps) rPr += `<w:smallCaps/>`;
    if (run.strike) rPr += `<w:strike/>`;
    const color = hex(run.color);
    if (color !== null) rPr += `<w:color w:val="${color}"/>`;
    // Tracking of zero is already the default, so it is absent rather than
    // w:val="0" — otherwise every untracked run would differ from the same run
    // written before this feature existed. Negative is legal, and is how a
    // large display size gets tightened.
    if (isNum(run.letterSpacing) && run.letterSpacing !== 0) {
      rPr += `<w:spacing w:val="${twips(run.letterSpacing)}"/>`;
    }
    const size = halfPoints(run.size);
    // szCs is not decoration: omit it and a complex-script run silently keeps
    // the default size.
    if (size !== null) rPr += `<w:sz w:val="${size}"/><w:szCs w:val="${size}"/>`;
    if (run.underline) rPr += `<w:u w:val="single"/>`;
    // Exact-hex highlight via run shading — w:highlight only takes named
    // colors; the reader maps both back to `highlight`.
    rPr += shadingXml(run.highlight);
    const pr = rPr === "" ? "" : `<w:rPr>${rPr}</w:rPr>`;

    // Newlines inside a run become soft line breaks.
    const parts = run.text.split("\n");
    let content = "";
    parts.forEach((part, i) => {
      if (i > 0) content += `<w:br/>`;
      if (part !== "") content += `<w:t xml:space="preserve">${Xml.text(part)}</w:t>`;
    });
    if (content === "") content = `<w:t xml:space="preserve"></w:t>`;
    return `<w:r>${pr}${content}</w:r>`;
  }

  // ── Footnotes ─────────────────────────────────────────────────────────────

  /**
   * A footnote run: its blocks rendered into the next footnote, and the
   * reference run that marks it in the text. Nothing inside another footnote,
   * and nothing for a footnote with no blocks to render (validation refuses
   * both; this engine renders no runs in a header or footer). Mirrors PHP
   * `renderFootnoteRun()`.
   *
   * The blocks render with the footnote part's OWN relationships swapped in,
   * so a link or image in a footnote registers in
   * `word/_rels/footnotes.xml.rels`. Media share the document's counter, so
   * names and drawing ids stay unique across the package.
   */
  private renderFootnoteRun(run: Any): string {
    const note = run.footnote;
    const blocks: Block[] = isPlainObject(note) && Array.isArray((note as Any).blocks) ? (note as Any).blocks : [];
    if (this.inFootnote || blocks.length === 0) return "";

    const [rels, hyperlinkIds] = [this.rels, this.hyperlinkIds];
    [this.rels, this.hyperlinkIds] = [this.footnoteRels, this.footnoteHyperlinkIds];
    this.inFootnote = true;
    let content: string;
    try {
      content = this.renderBlocks(blocks, { paragraphStyle: "FootnoteText" });
    } finally {
      [this.footnoteRels, this.footnoteHyperlinkIds] = [this.rels, this.hyperlinkIds];
      [this.rels, this.hyperlinkIds] = [rels, hyperlinkIds];
      this.inFootnote = false;
    }
    if (content === "") return "";

    // The mark opens the first paragraph, after its properties, followed by
    // the space Word puts between the number and the text. A footnote that
    // opens with a table or a quote gets a paragraph of its own for the mark,
    // and one that ends with a table a closing paragraph: a footnote is a
    // block container, and must end in a paragraph as a cell must.
    const mark = FOOTNOTE_MARK + `<w:r><w:t xml:space="preserve"> </w:t></w:r>`;
    if (content.startsWith("<w:p>")) {
      let at = "<w:p>".length;
      if (content.slice(at, at + "<w:pPr>".length) === "<w:pPr>") {
        at = content.indexOf("</w:pPr>") + "</w:pPr>".length;
      }
      content = content.slice(0, at) + mark + content.slice(at);
    } else {
      content = `<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>${FOOTNOTE_MARK}</w:p>` + content;
    }
    if (!content.endsWith("</w:p>")) {
      content += `<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr></w:p>`;
    }

    this.footnotes.push(content);
    const id = this.footnotes.length;
    return `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
  }

  /**
   * word/footnotes.xml: the separator (-1) and continuation separator (0) Word
   * requires before any footnote of the document's own -- without them Word
   * reports the file as damaged -- then the document's, from 1.
   */
  private footnotesXml(): string {
    const separator = (type: string, id: number, tag: string): string =>
      `<w:footnote w:type="${type}" w:id="${id}">` +
      `<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:${tag}/></w:r></w:p>` +
      `</w:footnote>`;

    let xml =
      Xml.declaration() +
      `<w:footnotes xmlns:w="${NS_W}" xmlns:r="${NS_R}" xmlns:wp="${NS_WP}" xmlns:a="${NS_A}" xmlns:pic="${NS_PIC}">` +
      separator("separator", -1, "separator") +
      separator("continuationSeparator", 0, "continuationSeparator");
    this.footnotes.forEach((content, i) => {
      xml += `<w:footnote w:id="${i + 1}">${content}</w:footnote>`;
    });
    return xml + `</w:footnotes>`;
  }

  /**
   * The settings part, written only when something needs it: in this engine,
   * even/odd headers or footnotes. CT_Settings is a sequence: evenAndOddHeaders
   * precedes footnotePr, which names the two separator footnotes, as Word itself
   * writes it.
   */
  private settingsXml(): string {
    return (
      Xml.declaration() +
      `<w:settings xmlns:w="${NS_W}">` +
      (this.evenAndOdd ? `<w:evenAndOddHeaders/>` : "") +
      (this.footnotes.length > 0 ? `<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>` : "") +
      `</w:settings>`
    );
  }

  private hyperlinkRel(url: string): string {
    const existing = this.hyperlinkIds.get(url);
    if (existing) return existing;
    const id = this.nextRelId();
    this.rels.push({ id, type: REL_HYPERLINK, target: url, external: true });
    this.hyperlinkIds.set(url, id);
    return id;
  }

  // ── Lists ─────────────────────────────────────────────────────────────────

  private renderList(items: ListItem[], ordered: boolean, depth: number): string {
    const numId = ordered ? NUM_ID_DECIMAL : NUM_ID_BULLET;
    const ilvl = Math.min(depth, MAX_ILVL);
    let out = "";
    for (const item of items) {
      const numPr = `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`;
      out += this.paragraph(this.pPr(item, null, numPr), this.renderRuns(item.runs ?? []));
      if (item.children && item.children.length > 0) {
        out += this.renderList(item.children, ordered, depth + 1);
      }
    }
    return out;
  }

  // ── Tables ────────────────────────────────────────────────────────────────

  private renderTable(block: Any): string {
    const rows: Any[] = Array.isArray(block.rows) ? block.rows.filter((r: Any) => r && typeof r === "object") : [];
    if (rows.length === 0) return "";

    const { laid, colCount } = layoutRows(rows);

    // A table narrower than the text column narrows its grid too: emitting
    // w:tblW pct while leaving the grid at full width hands Word two
    // contradictory answers.
    let tableWidth = this.contentWidth;
    let tblW = `<w:tblW w:w="0" w:type="auto"/>`;
    if (isNum(block.width) && block.width > 0) {
      const pct = Math.min(100, block.width);
      tableWidth = Math.round(this.contentWidth * (pct / 100));
      tblW = `<w:tblW w:w="${Math.round(pct * 50)}" w:type="pct"/>`;
    }

    const weights =
      Array.isArray(block.widths) && block.widths.length === colCount
        ? block.widths.map((w: Any) => (isNum(w) ? w : 0))
        : null;
    const grid = splitColumns(tableWidth, colCount, weights);

    let tblPr = tblW;
    if (block.align === "center" || block.align === "right") tblPr += `<w:jc w:val="${block.align}"/>`;
    const borders =
      block.borders && typeof block.borders === "object"
        ? block.borders
        : Object.fromEntries(TABLE_EDGES.map((e) => [e, DEFAULT_BORDER]));
    tblPr += bordersXml("tblBorders", borders, TABLE_EDGES);
    // Without w:tblLayout fixed, Word re-fits columns to their content and the
    // requested proportions are advisory.
    if (weights !== null) tblPr += `<w:tblLayout w:type="fixed"/>`;
    tblPr += marginsXml(
      "tblCellMar",
      block.cellPadding && typeof block.cellPadding === "object" ? block.cellPadding : DEFAULT_CELL_MARGINS_PT,
    );

    let out = `<w:tbl><w:tblPr>${tblPr}</w:tblPr><w:tblGrid>`;
    for (const w of grid) out += `<w:gridCol w:w="${w}"/>`;
    out += `</w:tblGrid>`;

    laid.forEach((line, r) => {
      const header = rows[r]?.header === true;
      out += `<w:tr>`;
      if (header) out += `<w:trPr><w:tblHeader/></w:trPr>`;
      let col = 0;
      for (const slot of line) {
        let width = 0;
        for (let i = col; i < Math.min(col + slot.span, grid.length); i++) width += grid[i]!;
        col += slot.span;
        out += `<w:tc>${this.tcPr(slot, width, header)}</w:tc>`;
      }
      out += `</w:tr>`;
    });

    return out + `</w:tbl>`;
  }

  /**
   * Cell properties, in CT_TcPr order:
   * tcW, gridSpan, vMerge, tcBorders, shd, tcMar, vAlign — followed by the
   * cell's content, which every cell must end with a w:p of.
   */
  private tcPr(slot: Slot, width: number, header: boolean): string {
    const cell = slot.cell;
    const continuation = slot.vMerge === "continue";

    let tcPr = `<w:tcW w:w="${width}" w:type="dxa"/>`;
    if (slot.span > 1) tcPr += `<w:gridSpan w:val="${slot.span}"/>`;
    if (slot.vMerge === "restart") tcPr += `<w:vMerge w:val="restart"/>`;
    else if (continuation) tcPr += `<w:vMerge/>`;

    if (!continuation && cell.borders && typeof cell.borders === "object") {
      tcPr += bordersXml("tcBorders", cell.borders, BOX_EDGES);
    }

    let shading = slot.shading ?? null;
    if (shading === null && header && !continuation) shading = `#${HEADER_FILL}`;
    tcPr += shadingXml(shading);

    if (!continuation && cell.padding && typeof cell.padding === "object") {
      tcPr += marginsXml("tcMar", cell.padding);
    }
    if (!continuation && (cell.valign === "top" || cell.valign === "center" || cell.valign === "bottom")) {
      tcPr += `<w:vAlign w:val="${cell.valign}"/>`;
    }

    let inner = this.renderBlocks(
      header && !continuation ? boldRuns(cell.blocks ?? []) : (cell.blocks ?? []),
      {},
    );
    // A table cell must end with a paragraph.
    if (inner === "" || !inner.endsWith("</w:p>")) inner += `<w:p/>`;

    return `<w:tcPr>${tcPr}</w:tcPr>${inner}`;
  }

  // ── Code / quote (SDT-wrapped for lossless round-trip) ──────────────────

  private renderCode(block: Any): string {
    const language = typeof block.language === "string" && block.language !== "" ? block.language : null;
    const tag = language ? `${SDT_TAG_CODE}:${language}` : SDT_TAG_CODE;
    const lines = String(block.text ?? "").split("\n");
    let body = "";
    for (const line of lines) {
      const runs =
        line === "" ? "" : `<w:r><w:t xml:space="preserve">${Xml.text(line)}</w:t></w:r>`;
      body += this.paragraph(this.pPr(null, "CodeBlock"), runs);
    }
    return (
      `<w:sdt><w:sdtPr><w:alias w:val="Code"/><w:tag w:val="${Xml.attr(tag)}"/></w:sdtPr>` +
      `<w:sdtContent>${body}</w:sdtContent></w:sdt>`
    );
  }

  private renderQuote(block: Any): string {
    const body = this.renderBlocks(block.blocks ?? [], { paragraphStyle: "Quote" });
    return (
      `<w:sdt><w:sdtPr><w:alias w:val="Quote"/><w:tag w:val="${SDT_TAG_QUOTE}"/></w:sdtPr>` +
      `<w:sdtContent>${body === "" ? "<w:p/>" : body}</w:sdtContent></w:sdt>`
    );
  }

  // ── Images ────────────────────────────────────────────────────────────────

  private renderImage(block: Any): string {
    const decoded = parseDataUrl(String(block.src ?? ""));
    if (!decoded) {
      throw new Error("last-word: image `src` must be a base64 data URL (data:image/png;base64,… or data:image/jpeg;base64,…)");
    }

    const ext = extForMime(decoded.mime);
    const name = `image${this.media.length + 1}.${ext}`;
    this.media.push({ name, bytes: decoded.bytes, ext });
    const rid = this.nextRelId();
    this.rels.push({ id: rid, type: REL_IMAGE, target: `media/${name}`, external: false });

    const { width, height } = resolveImageSize(block, decoded.bytes);
    const cx = Math.max(1, Math.round(width * EMU_PER_PX));
    const cy = Math.max(1, Math.round(height * EMU_PER_PX));

    const id = ++this.drawingId;
    const alt = typeof block.alt === "string" ? block.alt : "";
    const descr = alt === "" ? "" : ` descr="${Xml.attr(alt)}"`;

    return (
      `<w:p><w:r><w:drawing>` +
      `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
      `<wp:extent cx="${cx}" cy="${cy}"/>` +
      `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
      `<wp:docPr id="${id}" name="Picture ${id}"${descr}/>` +
      `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
      `<a:graphic><a:graphicData uri="${NS_PIC}">` +
      `<pic:pic>` +
      `<pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"${descr}/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
      `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
    );
  }

  // ── Charts ────────────────────────────────────────────────────────────────

  /**
   * A native chart, inline, in its own paragraph. An option nothing can be drawn
   * from (validation stops those before here) degrades to its title as text.
   */
  private renderChart(block: Any): string {
    const spec = chartSpec(block.option);
    if (spec === null) {
      const label = chartTitle(isPlainObject(block.option) ? (block.option as Any).title : null);
      return `<w:p><w:r><w:t xml:space="preserve">${Xml.text("[chart" + (label !== "" ? `: ${label}` : "") + "]")}</w:t></w:r></w:p>`;
    }
    const numberFormat = typeof block.numberFormat === "string" && block.numberFormat !== "" ? block.numberFormat : null;

    const n = this.charts.length + 1;
    this.charts.push({ xml: chartXml(spec, numberFormat), rels: chartRelsXml(n), workbook: workbookParts(spec, numberFormat) });
    const rid = this.nextRelId();
    this.rels.push({ id: rid, type: REL_CHART, target: `charts/chart${n}.xml`, external: false });

    const [cx, cy] = this.chartExtents(block);
    const alt = typeof block.alt === "string" ? block.alt : "";
    const descr = alt === "" ? "" : ` descr="${Xml.attr(alt)}"`;

    return (
      `<w:p><w:r><w:drawing>` +
      `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
      `<wp:extent cx="${cx}" cy="${cy}"/>` +
      `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
      `<wp:docPr id="${CHART_DOCPR_BASE + n}" name="Chart ${n}"${descr}/>` +
      `<wp:cNvGraphicFramePr/>` +
      `<a:graphic><a:graphicData uri="${NS_C}">` +
      `<c:chart xmlns:c="${NS_C}" r:id="${rid}"/>` +
      `</a:graphicData></a:graphic>` +
      `</wp:inline>` +
      `</w:drawing></w:r></w:p>`
    );
  }

  /**
   * A chart's extents in EMU. Width: widthPx, else `width` percent of the text
   * column, else the whole column -- never wider than the column. Height:
   * heightPx (scaled with the width when the width was capped), else
   * DEFAULT_ASPECT of the width. Mirrors PHP `chartExtents()`.
   */
  private chartExtents(block: Any): [number, number] {
    const column = this.contentWidth * EMU_PER_TWIP;
    const num = (v: unknown): number | null => (typeof v === "number" && Number.isFinite(v) ? v : null);
    const widthPx = num(block.widthPx);
    const heightPx = num(block.heightPx);
    const percent = num(block.width);

    let cx = column;
    if (widthPx !== null && widthPx > 0) cx = Math.round(widthPx * EMU_PER_PX);
    else if (percent !== null && percent > 0) cx = Math.round((column * Math.min(100, percent)) / 100);
    let cy: number | null = heightPx !== null && heightPx > 0 ? Math.round(heightPx * EMU_PER_PX) : null;

    if (cx > column) {
      if (cy !== null) cy = Math.round((cy * column) / cx);
      cx = column;
    }
    cx = Math.max(1, cx);
    return [cx, Math.max(1, cy ?? Math.round(cx * DEFAULT_ASPECT))];
  }

  // ── Parts ─────────────────────────────────────────────────────────────────

  private sectPr(page: PageGeometry, references: string): string {
    const orient = page.orientation === "landscape" ? ` w:orient="landscape"` : "";
    // References FIRST: the schema orders them before w:pgSz. CT_SectPr is a
    // sequence, and w:titlePg comes after w:pgMar.
    return (
      `<w:sectPr>${references}<w:pgSz w:w="${page.w}" w:h="${page.h}"${orient}/>` +
      `<w:pgMar w:top="${page.top}" w:right="${page.right}" w:bottom="${page.bottom}" w:left="${page.left}"` +
      ` w:header="${page.header}" w:footer="${page.footer}" w:gutter="0"/>` +
      (this.titlePg ? `<w:titlePg/>` : "") +
      `</w:sectPr>`
    );
  }

  private contentTypesXml(hasCore: boolean): string {
    const exts = [...new Set(this.media.map((m) => m.ext))].sort();
    let defaults =
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>`;
    for (const ext of exts) {
      defaults += `<Default Extension="${ext}" ContentType="${mimeForExt(ext)}"/>`;
    }
    if (this.charts.length > 0) {
      defaults += `<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>`;
    }
    let overrides =
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
      `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>`;
    if (this.hasSettings()) {
      overrides += `<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>`;
    }
    if (this.template?.hasTheme() === true) {
      // A part present but undeclared makes the package invalid, and Word reports
      // that as "unreadable content" without naming the part.
      overrides += `<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>`;
    }
    for (const b of this.bands) {
      overrides += `<Override PartName="/word/${b.name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${b.kind}+xml"/>`;
    }
    if (this.footnotes.length > 0) {
      overrides += `<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>`;
    }
    this.charts.forEach((_, i) => {
      overrides += `<Override PartName="/word/charts/chart${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`;
    });
    if (hasCore) {
      overrides += `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`;
    }
    return (
      Xml.declaration() +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults}${overrides}</Types>`
    );
  }

  private packageRelsXml(hasCore: boolean): string {
    let rels = `<Relationship Id="rId1" Type="${REL_DOCUMENT}" Target="word/document.xml"/>`;
    if (hasCore) {
      rels += `<Relationship Id="rId2" Type="${REL_CORE}" Target="docProps/core.xml"/>`;
    }
    return (
      Xml.declaration() +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`
    );
  }

  /**
   * docProps/core.xml: dc:title, and (PHP 0.16.0) dc:creator for
   * `meta.author`, title first. Each only when there is one, so a title-only
   * document's part is byte-for-byte what it always was. Deterministic: no
   * dcterms created/modified timestamps.
   */
  private coreXml(title: string | null, author: string | null): string {
    return (
      Xml.declaration() +
      `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
      `xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ` +
      `xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
      (title !== null ? `<dc:title>${Xml.text(title)}</dc:title>` : "") +
      (author !== null ? `<dc:creator>${Xml.text(author)}</dc:creator>` : "") +
      `</cp:coreProperties>`
    );
  }

  /**
   * The document's relationships in the PHP engine's order: styles, numbering
   * and theme; then settings and footnotes, when written; then the bands, and
   * the hyperlinks, images and charts in the order they were made.
   */
  private documentRelsXml(): string {
    const fixed = [REL_STYLES, REL_NUMBERING, REL_THEME];
    const parts: Rel[] = [];
    if (this.hasSettings()) {
      parts.push({ id: "rIdSettings", type: REL_SETTINGS, target: "settings.xml", external: false });
    }
    if (this.footnotes.length > 0) {
      parts.push({ id: "rIdFootnotes", type: REL_FOOTNOTES, target: "footnotes.xml", external: false });
    }
    return this.relsXml([
      ...this.rels.filter((r) => fixed.includes(r.type)),
      ...parts,
      ...this.rels.filter((r) => !fixed.includes(r.type)),
    ]);
  }

  /** A relationships part: the document's, or a footnotes part's own. */
  private relsXml(rels: Rel[]): string {
    let out = Xml.declaration() + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
    for (const rel of rels) {
      const mode = rel.external ? ` TargetMode="External"` : "";
      out += `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${Xml.attr(rel.target)}"${mode}/>`;
    }
    return out + `</Relationships>`;
  }

  /**
   * `word/styles.xml`.
   *
   * With a template (last-word#3) the template's OWN part is emitted, and only
   * the definitions this writer emits that the template does not declare are
   * appended. Two Word behaviours shape this, and both fail SILENTLY:
   *
   *  - a `w:pStyle` naming an undefined style renders UNSTYLED rather than
   *    erroring, so a template with no `CodeBlock` would quietly produce
   *    unformatted code;
   *  - a duplicate definition is legal and the LATER one wins, so appending a
   *    style the template already has would override the house style for exactly
   *    the subset we happen to share.
   */
  private stylesXml(doc: Doc): string {
    const definitions = this.styleDefinitions(doc);

    if (this.template !== undefined) {
      let missing = "";
      for (const [styleId, xml] of Object.entries(definitions)) {
        if (!this.template.defines(styleId)) missing += xml;
      }
      return this.template.stylesWith(missing);
    }

    return (
      Xml.declaration() +
      `<w:styles xmlns:w="${NS_W}">` +
      this.docDefaults(doc) +
      Object.values(definitions).join("") +
      `</w:styles>`
    );
  }

  private docDefaults(doc: Doc): string {
    const font = typeof (doc as Any).defaultFont === "string" && (doc as Any).defaultFont !== ""
      ? Xml.attr((doc as Any).defaultFont)
      : "Calibri";
    const size = halfPoints((doc as Any).defaultSize) ?? 22;

    return (
      `<w:docDefaults><w:rPrDefault><w:rPr>` +
      `<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}" w:cs="${font}"/>` +
      `<w:sz w:val="${size}"/><w:szCs w:val="${size}"/>` +
      `</w:rPr></w:rPrDefault>` +
      `<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>` +
      `</w:docDefaults>`
    );
  }

  /**
   * Every style this writer can emit, keyed by `w:styleId`.
   *
   * Insertion order is the emission order and is unchanged from when this was
   * one concatenation, so the no-template output is byte-identical — which the
   * determinism suite and the PHP parity suite both check.
   */
  private styleDefinitions(doc: Doc): Record<string, string> {
    const headingSizes = [40, 32, 28, 26, 24, 22];
    const styles: Record<string, string> = {};

    styles["Normal"] =
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>`;

    for (let level = 1; level <= 6; level++) {
      const sz = headingSizes[level - 1]!;
      styles[`Heading${level}`] =
        `<w:style w:type="paragraph" w:styleId="Heading${level}"><w:name w:val="heading ${level}"/>` +
        `<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>` +
        `<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="${level - 1}"/></w:pPr>` +
        `<w:rPr><w:b/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr></w:style>`;
    }

    styles["Quote"] =
      `<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/>` +
      `<w:basedOn w:val="Normal"/><w:qFormat/>` +
      `<w:pPr><w:ind w:left="720"/></w:pPr>` +
      `<w:rPr><w:i/><w:color w:val="595959"/></w:rPr></w:style>`;

    styles["CodeBlock"] =
      `<w:style w:type="paragraph" w:styleId="CodeBlock"><w:name w:val="Code Block"/>` +
      `<w:basedOn w:val="Normal"/>` +
      `<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>` +
      `<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/></w:pPr>` +
      `<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>` +
      `<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>`;

    styles["InlineCode"] =
      `<w:style w:type="character" w:styleId="InlineCode"><w:name w:val="Inline Code"/>` +
      `<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>` +
      `<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/></w:rPr></w:style>`;

    styles["Hyperlink"] =
      `<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>` +
      `<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>`;

    // Only with a footnote, so every other document's styles stay byte-identical.
    // Word's own names ("footnote text", "footnote reference"): a template that
    // defines them keeps its own, as for every other style. Shared with PHP.
    if (this.footnotes.length > 0) {
      styles["FootnoteText"] =
        `<w:style w:type="paragraph" w:styleId="FootnoteText">` +
        `<w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>` +
        `<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>` +
        `<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>` +
        `</w:style>`;
      styles["FootnoteReference"] =
        `<w:style w:type="character" w:styleId="FootnoteReference">` +
        `<w:name w:val="footnote reference"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>` +
        `<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>` +
        `</w:style>`;
    }

    // The `LastWordTable` style used to live here, carrying the borders, the
    // cell margins and the first-row bold that PHP and Python wrote inline.
    // A named style cannot vary per table instance, so per-table borders
    // forced everything inline and the style has nothing left to say. Its
    // 60/108 cell margins became the shared default rather than being dropped.

    return styles;
  }

  private numberingXml(): string {
    let out = Xml.declaration() + `<w:numbering xmlns:w="${NS_W}">`;
    // abstractNum 0: bullet, 6 indent levels
    out += `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>`;
    for (let i = 0; i <= MAX_ILVL; i++) {
      out +=
        `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/>` +
        `<w:lvlText w:val="•"/><w:lvlJc w:val="left"/>` +
        `<w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr></w:lvl>`;
    }
    out += `</w:abstractNum>`;
    // abstractNum 1: decimal, 6 indent levels
    out += `<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>`;
    for (let i = 0; i <= MAX_ILVL; i++) {
      out +=
        `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="decimal"/>` +
        `<w:lvlText w:val="%${i + 1}."/><w:lvlJc w:val="left"/>` +
        `<w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr></w:lvl>`;
    }
    out += `</w:abstractNum>`;
    out += `<w:num w:numId="${NUM_ID_BULLET}"><w:abstractNumId w:val="0"/></w:num>`;
    out += `<w:num w:numId="${NUM_ID_DECIMAL}"><w:abstractNumId w:val="1"/></w:num>`;
    return out + `</w:numbering>`;
  }
}

// ── helpers ─────────────────────────────────────────────────────────────────

const encoder = new TextEncoder();

function encode(xml: string): Uint8Array {
  return encoder.encode(xml);
}

function clampLevel(level: Any): number {
  const n = Number.isFinite(Number(level)) ? Math.trunc(Number(level)) : 1;
  return Math.min(6, Math.max(1, n));
}

function extForMime(mime: string): string {
  switch (mime) {
    case "image/png":
      return "png";
    case "image/jpeg":
    case "image/jpg":
      return "jpeg";
    case "image/gif":
      return "gif";
    default:
      return mime.split("/")[1] ?? "bin";
  }
}

function mimeForExt(ext: string): string {
  switch (ext) {
    case "png":
      return "image/png";
    case "jpeg":
      return "image/jpeg";
    case "gif":
      return "image/gif";
    default:
      return "application/octet-stream";
  }
}

/**
 * Resolve the rendered pixel size of an image block: explicit px win, missing
 * dimensions are derived from the intrinsic (sniffed) size keeping aspect,
 * and everything is capped at 6.5in width.
 */
export function resolveImageSize(
  block: { widthPx?: number; heightPx?: number },
  bytes: Uint8Array,
): { width: number; height: number } {
  const intrinsic = sniffImageSize(bytes);
  let width = typeof block.widthPx === "number" && block.widthPx > 0 ? block.widthPx : undefined;
  let height = typeof block.heightPx === "number" && block.heightPx > 0 ? block.heightPx : undefined;

  if (width === undefined && height === undefined) {
    width = intrinsic?.width ?? 300;
    height = intrinsic?.height ?? 200;
  } else if (width === undefined) {
    const aspect = intrinsic ? intrinsic.width / intrinsic.height : 1.5;
    width = height! * aspect;
  } else if (height === undefined) {
    const aspect = intrinsic ? intrinsic.width / intrinsic.height : 1.5;
    height = width / aspect;
  }

  if (width! > MAX_WIDTH_PX) {
    height = (height! * MAX_WIDTH_PX) / width!;
    width = MAX_WIDTH_PX;
  }

  return { width: Math.max(1, Math.round(width!)), height: Math.max(1, Math.round(height!)) };
}
