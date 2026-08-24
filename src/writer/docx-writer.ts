/**
 * DocxWriter — Doc JSON model → .docx bytes (OOXML / WordprocessingML).
 * Deterministic: no timestamps, fixed zip entry order, rel ids assigned in
 * traversal order. Mirrors PHP `Writer\DocxWriter`.
 */

import { Xml } from "../helpers/xml";
import { parseDataUrl, sniffImageSize } from "../helpers/image-size";
import type { Block, Doc, ListItem, Run } from "../schema/types";
import { zipSync, type ZipFile } from "../zip";

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
const REL_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const REL_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_CORE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

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
}

/** Page size and margins in twips. */
export function pageGeometry(page: Any): PageGeometry {
  const size = typeof page?.size === "string" ? page.size.toLowerCase() : "letter";
  let [w, h] = PAGE_SIZES[size] ?? PAGE_SIZES.letter!;

  const orientation = page?.orientation === "landscape" ? "landscape" : "portrait";
  // Swapping the axes without w:orient gives a page that is the right shape
  // and prints portrait. Both are required.
  if (orientation === "landscape") [w, h] = [h, w];

  const side = (k: string): number => twips(page?.margins?.[k]) ?? 1440;

  return { w, h, top: side("top"), right: side("right"), bottom: side("bottom"), left: side("left"), orientation };
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
  /** Twips between the page margins; set from `page` before any block renders. */
  private contentWidth = 9360;

  toBytes(doc: Doc): Uint8Array {
    this.rels = [
      { id: "rId1", type: REL_STYLES, target: "styles.xml", external: false },
      { id: "rId2", type: REL_NUMBERING, target: "numbering.xml", external: false },
    ];
    this.hyperlinkIds = new Map();
    this.media = [];
    this.drawingId = 0;

    // Table grids are laid out against the section's content width, so it has
    // to be known before any block renders. Carrying 9360 as a literal — which
    // all three engines did — silently gives a document with narrowed margins
    // a table that no longer matches its own page.
    const page = pageGeometry((doc as Any).page);
    this.contentWidth = page.w - page.left - page.right;

    const body = this.renderBlocks(doc.blocks ?? [], {});
    const documentXml =
      Xml.declaration() +
      `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}" xmlns:wp="${NS_WP}" xmlns:a="${NS_A}" xmlns:pic="${NS_PIC}">` +
      `<w:body>${body}${this.sectPr(page)}</w:body></w:document>`;

    const files: ZipFile[] = [
      { name: "[Content_Types].xml", data: encode(this.contentTypesXml(doc)) },
      { name: "_rels/.rels", data: encode(this.packageRelsXml(doc)) },
    ];
    if (doc.title !== undefined) {
      files.push({ name: "docProps/core.xml", data: encode(this.coreXml(String(doc.title))) });
    }
    files.push(
      { name: "word/document.xml", data: encode(documentXml) },
      { name: "word/styles.xml", data: encode(this.stylesXml(doc)) },
      { name: "word/numbering.xml", data: encode(this.numberingXml()) },
      { name: "word/_rels/document.xml.rels", data: encode(this.documentRelsXml()) },
    );
    for (const m of this.media) {
      files.push({ name: `word/media/${m.name}`, data: m.bytes });
    }
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

  private hyperlinkRel(url: string): string {
    const existing = this.hyperlinkIds.get(url);
    if (existing) return existing;
    const id = `rId${this.rels.length + 1}`;
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
    const rid = `rId${this.rels.length + 1}`;
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

  // ── Parts ─────────────────────────────────────────────────────────────────

  private sectPr(page: PageGeometry): string {
    const orient = page.orientation === "landscape" ? ` w:orient="landscape"` : "";
    return (
      `<w:sectPr><w:pgSz w:w="${page.w}" w:h="${page.h}"${orient}/>` +
      `<w:pgMar w:top="${page.top}" w:right="${page.right}" w:bottom="${page.bottom}" w:left="${page.left}"` +
      ` w:header="720" w:footer="720" w:gutter="0"/>` +
      `</w:sectPr>`
    );
  }

  private contentTypesXml(doc: Doc): string {
    const exts = [...new Set(this.media.map((m) => m.ext))].sort();
    let defaults =
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>`;
    for (const ext of exts) {
      defaults += `<Default Extension="${ext}" ContentType="${mimeForExt(ext)}"/>`;
    }
    let overrides =
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
      `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>`;
    if (doc.title !== undefined) {
      overrides += `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`;
    }
    return (
      Xml.declaration() +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults}${overrides}</Types>`
    );
  }

  private packageRelsXml(doc: Doc): string {
    let rels = `<Relationship Id="rId1" Type="${REL_DOCUMENT}" Target="word/document.xml"/>`;
    if (doc.title !== undefined) {
      rels += `<Relationship Id="rId2" Type="${REL_CORE}" Target="docProps/core.xml"/>`;
    }
    return (
      Xml.declaration() +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`
    );
  }

  private coreXml(title: string): string {
    return (
      Xml.declaration() +
      `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
      `xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ` +
      `xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
      `<dc:title>${Xml.text(title)}</dc:title>` +
      `</cp:coreProperties>`
    );
  }

  private documentRelsXml(): string {
    let out = Xml.declaration() + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
    for (const rel of this.rels) {
      const mode = rel.external ? ` TargetMode="External"` : "";
      out += `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${Xml.attr(rel.target)}"${mode}/>`;
    }
    return out + `</Relationships>`;
  }

  private stylesXml(doc: Doc): string {
    const headingSizes = [40, 32, 28, 26, 24, 22];
    const font = typeof (doc as Any).defaultFont === "string" && (doc as Any).defaultFont !== ""
      ? Xml.attr((doc as Any).defaultFont)
      : "Calibri";
    const size = halfPoints((doc as Any).defaultSize) ?? 22;
    let styles =
      `<w:docDefaults><w:rPrDefault><w:rPr>` +
      `<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}" w:cs="${font}"/>` +
      `<w:sz w:val="${size}"/><w:szCs w:val="${size}"/>` +
      `</w:rPr></w:rPrDefault>` +
      `<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>` +
      `</w:docDefaults>` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>`;

    for (let level = 1; level <= 6; level++) {
      const sz = headingSizes[level - 1]!;
      styles +=
        `<w:style w:type="paragraph" w:styleId="Heading${level}"><w:name w:val="heading ${level}"/>` +
        `<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>` +
        `<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="${level - 1}"/></w:pPr>` +
        `<w:rPr><w:b/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr></w:style>`;
    }

    styles +=
      `<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/>` +
      `<w:basedOn w:val="Normal"/><w:qFormat/>` +
      `<w:pPr><w:ind w:left="720"/></w:pPr>` +
      `<w:rPr><w:i/><w:color w:val="595959"/></w:rPr></w:style>` +
      `<w:style w:type="paragraph" w:styleId="CodeBlock"><w:name w:val="Code Block"/>` +
      `<w:basedOn w:val="Normal"/>` +
      `<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>` +
      `<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/></w:pPr>` +
      `<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>` +
      `<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>` +
      `<w:style w:type="character" w:styleId="InlineCode"><w:name w:val="Inline Code"/>` +
      `<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>` +
      `<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/></w:rPr></w:style>` +
      `<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>` +
      `<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>`;
    // The `LastWordTable` style used to live here, carrying the borders, the
    // cell margins and the first-row bold that PHP and Python wrote inline.
    // A named style cannot vary per table instance, so per-table borders
    // forced everything inline and the style has nothing left to say. Its
    // 60/108 cell margins became the shared default rather than being dropped.

    return Xml.declaration() + `<w:styles xmlns:w="${NS_W}">${styles}</w:styles>`;
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
