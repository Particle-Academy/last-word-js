/**
 * DocxReader — .docx bytes → Doc JSON model. Handles this package's own
 * writer output losslessly (round-trip), the PHP mirror's output (same
 * metadata slots since 0.2.0, plus the legacy `LastWordCode_{lang}` bookmark
 * fallback for PHP ≤0.1.x files) and tolerates Word-authored files:
 * headings via pStyle Heading1-9 OR outlineLvl, numPr lists with ilvl nesting,
 * hyperlinks via rels, images via blip r:embed, page breaks, bottom-border-only
 * paragraphs → hr. Unknown constructs degrade to plain paragraphs, never throw.
 * Mirrors PHP `Reader\DocxReader`.
 */

import { base64Encode } from "../helpers/image-size";
import type { Block, Doc, ListItem, Run } from "../schema/types";
import { unzipSync } from "../zip";
import { at, el, els, parseXml, type XmlNode } from "./xml";
import { EMU_PER_PX, SDT_TAG_CODE, SDT_TAG_QUOTE, splitColumns, PAGE_SIZES } from "../writer/docx-writer";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

const DECODER = new TextDecoder();

interface RelInfo {
  target: string;
  external: boolean;
}

/** Named w:highlight values → hex (reader tolerance for Word-authored files). */
const HIGHLIGHT_NAMES: Record<string, string> = {
  yellow: "#FFFF00",
  green: "#00FF00",
  cyan: "#00FFFF",
  magenta: "#FF00FF",
  blue: "#0000FF",
  red: "#FF0000",
  darkBlue: "#00008B",
  darkCyan: "#008B8B",
  darkGreen: "#006400",
  darkMagenta: "#8B008B",
  darkRed: "#8B0000",
  darkYellow: "#808000",
  darkGray: "#A9A9A9",
  lightGray: "#D3D3D3",
  black: "#000000",
  white: "#FFFFFF",
};

const MONO_FONTS = ["consolas", "courier new", "courier", "menlo", "monaco", "source code pro"];

const ORDERED_NUM_FMTS = [
  "decimal",
  "decimalZero",
  "upperRoman",
  "lowerRoman",
  "upperLetter",
  "lowerLetter",
];

interface FlatListItem {
  numId: string;
  ilvl: number;
  runs: Run[];
}

interface WalkCtx {
  insideQuote?: boolean;
}

export class DocxReader {
  private parts: Record<string, Uint8Array> = {};
  private rels: Record<string, RelInfo> = {};
  private numOrdered: Record<string, boolean> = {};

  read(bytes: Uint8Array): Doc {
    this.parts = unzipSync(bytes);
    this.rels = this.loadRels("word/_rels/document.xml.rels");
    this.numOrdered = this.loadNumbering();

    const docXml = this.partText("word/document.xml");
    const root = docXml ? parseXml(docXml) : null;
    const body = el(root, "body");

    const doc: Doc = { blocks: body ? this.walkBody(body.children, {}) : [] };

    const title = this.readTitle();
    if (title !== null) doc.title = title;

    // Section geometry and document defaults are surfaced ONLY when they
    // differ from what the writer produces unasked, for the same reason table
    // options are: a Letter portrait page at one-inch margins is what every
    // document written before `page` existed contains.
    const page = this.readPage(el(body, "sectPr"));
    if (page !== null) doc.page = page as Any;
    const defaults = this.readDocDefaults();
    if (defaults.font !== null) (doc as Any).defaultFont = defaults.font;
    if (defaults.size !== null) (doc as Any).defaultSize = defaults.size;

    return doc;
  }

  private readPage(sectPr: XmlNode | undefined): Record<string, unknown> | null {
    if (!sectPr) return null;
    const out: Record<string, unknown> = {};

    const pgSz = el(sectPr, "pgSz");
    let w = numAttr(pgSz, "w") ?? 12240;
    let h = numAttr(pgSz, "h") ?? 15840;
    if (at(pgSz, "orient") === "landscape") {
      out.orientation = "landscape";
      [w, h] = [h, w];
    }
    for (const [name, dims] of Object.entries(PAGE_SIZES)) {
      if (dims[0] === w && dims[1] === h && name !== "letter") out.size = name;
    }

    const pgMar = el(sectPr, "pgMar");
    const margins: Record<string, number> = {};
    for (const side of BOX_EDGES) {
      const value = numAttr(pgMar, side);
      if (value !== null && value !== 1440) margins[side] = points(value);
    }
    if (Object.keys(margins).length > 0) out.margins = margins;

    return Object.keys(out).length > 0 ? out : null;
  }

  private readDocDefaults(): { font: string | null; size: number | null } {
    const xml = this.partText("word/styles.xml");
    if (!xml) return { font: null, size: null };
    const rPr = el(el(el(parseXml(xml), "docDefaults"), "rPrDefault"), "rPr");

    const ascii = at(el(rPr, "rFonts"), "ascii");
    const sz = numAttr(el(rPr, "sz"), "val");

    return {
      font: ascii !== undefined && ascii !== "" && ascii !== "Calibri" ? ascii : null,
      size: sz !== null && sz !== 22 ? sz / 2 : null,
    };
  }

  // ── Parts / metadata ─────────────────────────────────────────────────────

  private partText(name: string): string | null {
    const part = this.parts[name];
    return part ? DECODER.decode(part) : null;
  }

  private readTitle(): string | null {
    const xml = this.partText("docProps/core.xml");
    if (!xml) return null;
    const root = parseXml(xml);
    const title = el(root, "title");
    return title ? title.text : null;
  }

  private loadRels(name: string): Record<string, RelInfo> {
    const xml = this.partText(name);
    if (!xml) return {};
    const root = parseXml(xml);
    const out: Record<string, RelInfo> = {};
    for (const rel of els(root, "Relationship")) {
      const id = at(rel, "Id");
      const target = at(rel, "Target");
      if (!id || !target) continue;
      out[id] = { target, external: at(rel, "TargetMode") === "External" };
    }
    return out;
  }

  private loadNumbering(): Record<string, boolean> {
    const xml = this.partText("word/numbering.xml");
    if (!xml) return {};
    const root = parseXml(xml);
    const abstractOrdered: Record<string, boolean> = {};
    for (const abs of els(root, "abstractNum")) {
      const id = at(abs, "abstractNumId") ?? "";
      const lvl0 = els(abs, "lvl").find((l) => at(l, "ilvl") === "0") ?? el(abs, "lvl");
      const fmt = at(el(lvl0, "numFmt"), "val") ?? "bullet";
      abstractOrdered[id] = ORDERED_NUM_FMTS.includes(fmt);
    }
    const out: Record<string, boolean> = {};
    for (const num of els(root, "num")) {
      const numId = at(num, "numId") ?? "";
      const absId = at(el(num, "abstractNumId"), "val") ?? "";
      out[numId] = abstractOrdered[absId] ?? false;
    }
    return out;
  }

  // ── Body walking ─────────────────────────────────────────────────────────

  private walkBody(nodes: XmlNode[], ctx: WalkCtx): Block[] {
    const blocks: Block[] = [];
    let listBuf: FlatListItem[] = [];
    let codeBuf: string[] | null = null;
    let codeLang: string | null = null;
    let quoteBuf: Block[] | null = null;

    const flushList = (): void => {
      if (listBuf.length > 0) {
        blocks.push(...this.buildLists(listBuf));
        listBuf = [];
      }
    };
    const flushCode = (): void => {
      if (codeBuf !== null) {
        const block: Block = { type: "code", text: codeBuf.join("\n") };
        if (codeLang !== null) (block as Any).language = codeLang;
        blocks.push(orderCodeKeys(block));
        codeBuf = null;
        codeLang = null;
      }
    };
    const flushQuote = (): void => {
      if (quoteBuf !== null) {
        blocks.push({ type: "quote", blocks: quoteBuf });
        quoteBuf = null;
      }
    };
    const flushAll = (): void => {
      flushList();
      flushCode();
      flushQuote();
    };

    for (const node of nodes) {
      switch (node.name) {
        case "p": {
          const kind = this.classifyParagraph(node);
          if (kind.kind === "listItem") {
            flushCode();
            flushQuote();
            listBuf.push(kind.item);
          } else if (kind.kind === "codeLine") {
            flushList();
            flushQuote();
            if (codeBuf === null) {
              codeBuf = [];
              codeLang = kind.language ?? null;
            }
            codeBuf.push(kind.text);
          } else if (kind.kind === "quoteParagraph" && !ctx.insideQuote) {
            flushList();
            flushCode();
            if (quoteBuf === null) quoteBuf = [];
            quoteBuf.push(...kind.blocks);
          } else {
            flushAll();
            blocks.push(...kind.blocks);
          }
          break;
        }
        case "tbl":
          flushAll();
          // Drop the pad the writer puts BETWEEN adjacent tables — the one
          // that stops Word merging them. Deliberately narrow: only an empty
          // paragraph sandwiched between two tables goes, so a blank line an
          // author actually wrote still survives the read.
          if (blocks.length >= 2 && isEmptyParagraph(blocks[blocks.length - 1]) &&
              (blocks[blocks.length - 2] as Any)?.type === "table") {
            blocks.pop();
          }
          blocks.push(this.parseTable(node, ctx));
          break;
        case "sdt": {
          flushAll();
          blocks.push(...this.parseSdt(node, ctx));
          break;
        }
        case "sectPr":
        case "bookmarkStart":
        case "bookmarkEnd":
        case "proofErr":
          break;
        default:
          // Unknown container: recurse looking for readable content (degrade, never throw).
          if (node.children.length > 0) {
            const inner = this.walkBody(node.children, ctx);
            if (inner.length > 0) {
              flushAll();
              blocks.push(...inner);
            }
          }
          break;
      }
    }
    flushAll();
    return blocks;
  }

  private parseSdt(sdt: XmlNode, ctx: WalkCtx): Block[] {
    const tag = at(el(el(sdt, "sdtPr"), "tag"), "val") ?? "";
    const content = el(sdt, "sdtContent");
    if (!content) return [];

    if (tag === SDT_TAG_CODE || tag.startsWith(SDT_TAG_CODE + ":")) {
      const lines = els(content, "p").map((p) => this.plainText(p));
      const block: Block = { type: "code", text: lines.join("\n") };
      if (tag.length > SDT_TAG_CODE.length + 1) {
        (block as Any).language = tag.slice(SDT_TAG_CODE.length + 1);
      }
      return [orderCodeKeys(block)];
    }
    if (tag === SDT_TAG_QUOTE) {
      return [{ type: "quote", blocks: this.walkBody(content.children, { ...ctx, insideQuote: true }) }];
    }
    // Unknown SDT: read its content transparently.
    return this.walkBody(content.children, ctx);
  }

  // ── Paragraph classification ─────────────────────────────────────────────

  private classifyParagraph(
    p: XmlNode,
  ):
    | { kind: "listItem"; item: FlatListItem }
    | { kind: "codeLine"; text: string; language?: string }
    | { kind: "quoteParagraph"; blocks: Block[] }
    | { kind: "blocks"; blocks: Block[] } {
    const pPr = el(p, "pPr");
    const styleId = at(el(pPr, "pStyle"), "val") ?? "";

    // List item?
    const numPr = el(pPr, "numPr");
    if (numPr) {
      const numId = at(el(numPr, "numId"), "val") ?? "";
      const ilvl = parseInt(at(el(numPr, "ilvl"), "val") ?? "0", 10) || 0;
      return { kind: "listItem", item: { numId, ilvl, runs: this.parseRuns(p, undefined) } };
    }

    // Code line? (SDT-less tolerance path — bare style, or PHP last-word
    // ≤0.1.x output which stashed the language in an invisible
    // `LastWordCode_{lang}` bookmark on the first code paragraph.)
    if (/^(CodeBlock|SourceCode|HTMLPreformatted|Code)$/i.test(styleId)) {
      const language = this.legacyBookmarkLanguage(p);
      return language !== undefined
        ? { kind: "codeLine", text: this.plainText(p), language }
        : { kind: "codeLine", text: this.plainText(p) };
    }

    const blocks: Block[] = [];

    // Images (drawings may sit alongside text runs).
    const drawings = findAll(p, "drawing");
    const runs = this.parseRuns(p, undefined);

    // Heading?
    const headingMatch = /^Heading([1-9])$/i.exec(styleId);
    const outlineLvl = at(el(pPr, "outlineLvl"), "val");
    if (headingMatch || outlineLvl !== undefined) {
      const raw = headingMatch ? parseInt(headingMatch[1]!, 10) : parseInt(outlineLvl ?? "0", 10) + 1;
      const level = Math.min(6, Math.max(1, raw)) as 1 | 2 | 3 | 4 | 5 | 6;
      blocks.push({ type: "heading", level, runs, ...paragraphPropsFrom(pPr) });
      for (const d of drawings) blocks.push(...this.parseDrawing(d));
      return { kind: "blocks", blocks };
    }

    // Page break only?
    const hasPageBreak = findAll(p, "br").some((br) => at(br, "type") === "page");
    const text = runs.map((r) => r.text).join("");
    if (hasPageBreak && text === "" && drawings.length === 0) {
      return { kind: "blocks", blocks: [{ type: "pageBreak" }] };
    }

    // Horizontal rule? (bottom-border-only empty paragraph)
    const pBdr = el(pPr, "pBdr");
    if (pBdr && el(pBdr, "bottom") && text === "" && drawings.length === 0) {
      return { kind: "blocks", blocks: [{ type: "hr" }] };
    }

    // Plain (or quote-styled) paragraph.
    if (text !== "" || drawings.length === 0) {
      const para: Block = { type: "paragraph", runs, ...paragraphPropsFrom(pPr) };
      blocks.push(para);
    }
    for (const d of drawings) blocks.push(...this.parseDrawing(d));

    if (/^(Quote|IntenseQuote|BlockQuote|Blockquote)$/i.test(styleId)) {
      return { kind: "quoteParagraph", blocks };
    }
    return { kind: "blocks", blocks };
  }

  /**
   * Back-compat: the code language from a PHP last-word ≤0.1.x
   * `LastWordCode_{lang}` bookmark on this paragraph, if present. The
   * canonical carrier is the `lastword:code:{lang}` sdt tag.
   */
  private legacyBookmarkLanguage(p: XmlNode): string | undefined {
    for (const child of p.children) {
      if (child.name !== "bookmarkStart") continue;
      const name = at(child, "name") ?? "";
      if (name.startsWith("LastWordCode_") && name.length > "LastWordCode_".length) {
        return name.slice("LastWordCode_".length);
      }
    }
    return undefined;
  }

  /** Concatenated visible text of a paragraph (tabs and soft breaks included). */
  private plainText(p: XmlNode): string {
    return this.parseRuns(p, undefined)
      .map((r) => r.text)
      .join("");
  }

  // ── Runs ─────────────────────────────────────────────────────────────────

  private parseRuns(container: XmlNode, link: string | undefined): Run[] {
    const runs: Run[] = [];
    for (const child of container.children) {
      switch (child.name) {
        case "r":
          runs.push(...this.runFrom(child, link));
          break;
        case "hyperlink": {
          const rid = at(child, "id");
          const rel = rid ? this.rels[rid] : undefined;
          const anchor = at(child, "anchor");
          const url = rel?.external ? rel.target : anchor ? `#${anchor}` : undefined;
          runs.push(...this.parseRuns(child, url ?? link));
          break;
        }
        case "sdt": {
          const content = el(child, "sdtContent");
          if (content) runs.push(...this.parseRuns(content, link));
          break;
        }
        case "ins":
        case "smartTag":
          runs.push(...this.parseRuns(child, link));
          break;
        default:
          break;
      }
    }
    return mergeRuns(runs);
  }

  private runFrom(r: XmlNode, link: string | undefined): Run[] {
    const rPr = el(r, "rPr");
    let text = "";
    for (const child of r.children) {
      if (child.name === "t") text += child.text;
      else if (child.name === "br" && at(child, "type") !== "page") text += "\n";
      else if (child.name === "cr") text += "\n";
      else if (child.name === "tab") text += "\t";
    }
    if (text === "") return [];

    const run: Run = { text };
    if (onFlag(el(rPr, "b"))) run.bold = true;
    if (onFlag(el(rPr, "i"))) run.italic = true;
    if (onFlag(el(rPr, "strike"))) run.strike = true;

    const u = el(rPr, "u");
    if (u && (at(u, "val") ?? "single") !== "none") run.underline = true;

    if (onFlag(el(rPr, "smallCaps"))) run.smallCaps = true;

    const rStyle = at(el(rPr, "rStyle"), "val") ?? "";
    const ascii = at(el(rPr, "rFonts"), "ascii");
    const asciiFont = (ascii ?? "").toLowerCase();
    if (/^InlineCode$/i.test(rStyle) || MONO_FONTS.includes(asciiFont)) run.code = true;
    // A `code` run's Consolas came FROM `code`, so surfacing it as `font` too
    // would hand back a bigger model than went in — and the next write would
    // then differ from the one just read.
    else if (ascii !== undefined && ascii !== "") run.font = ascii;

    const sz = numAttr(el(rPr, "sz"), "val");
    if (sz !== null) run.size = sz / 2;

    // Tracking of zero is the writer's "absent", so a zero here would be a
    // property nobody asked for.
    const tracking = numAttr(el(rPr, "spacing"), "val");
    if (tracking !== null && tracking !== 0) run.letterSpacing = tracking / 20;

    if (link) run.link = link;

    const color = at(el(rPr, "color"), "val");
    if (color && color !== "auto") run.color = `#${color.toUpperCase()}`;

    const shdFill = at(el(rPr, "shd"), "fill");
    if (shdFill && shdFill !== "auto") {
      run.highlight = `#${shdFill.toUpperCase()}`;
    } else {
      const named = at(el(rPr, "highlight"), "val");
      if (named && named !== "none" && HIGHLIGHT_NAMES[named]) run.highlight = HIGHLIGHT_NAMES[named];
    }

    return [run];
  }

  // ── Lists ────────────────────────────────────────────────────────────────

  private buildLists(flat: FlatListItem[]): Block[] {
    const blocks: Block[] = [];
    let i = 0;
    while (i < flat.length) {
      const numId = flat[i]!.numId;
      const group: FlatListItem[] = [];
      while (i < flat.length && flat[i]!.numId === numId) {
        group.push(flat[i]!);
        i++;
      }
      const ordered = this.numOrdered[numId] ?? false;
      const items = buildTree(group);
      const block: Block = ordered
        ? ({ type: "list", ordered: true, items } as Block)
        : ({ type: "list", items } as Block);
      blocks.push(block);
    }
    return blocks;
  }

  // ── Tables ───────────────────────────────────────────────────────────────

  /**
   * A table back into the model.
   *
   * Two things make this the hardest read in the package. First, the writer
   * emits borders, cell margins and a `w:tcW` for every cell whether or not
   * the model asked — so anything it would have produced anyway is NOT
   * surfaced, or every document written before these options existed would
   * come back carrying options nobody set. Second, the file contains the
   * `w:vMerge` continuation cells the writer synthesised, and they have to be
   * dropped and turned back into a `rowSpan` on the cell above.
   */
  private parseTable(tbl: XmlNode, ctx: WalkCtx): Block {
    const tblPr = el(tbl, "tblPr");
    const grid = els(el(tbl, "tblGrid"), "gridCol").map((g) => numAttr(g, "w") ?? 0);
    const table: Record<string, unknown> = { type: "table" };

    const tblW = el(tblPr, "tblW");
    if (at(tblW, "type") === "pct") {
      const w = numAttr(tblW, "w");
      if (w !== null) table.width = Math.round((w / 50) * 100) / 100;
    }

    const jc = at(el(tblPr, "jc"), "val");
    if (jc === "center" || jc === "right") table.align = jc;

    const borders = bordersFrom(el(tblPr, "tblBorders"), TABLE_EDGES);
    if (borders && !isDefaultTableBorders(borders)) table.borders = borders;

    const padding = sidesFrom(el(tblPr, "tblCellMar"));
    if (padding && !isDefaultCellMargins(padding)) table.cellPadding = padding;

    // Weights are only surfaced when the grid is NOT what an equal split
    // would have produced — compared against the split the writer computes,
    // not tested for exact equality, so a three-column table whose width does
    // not divide by three is still recognised as equal.
    const total = grid.reduce((a, b) => a + b, 0);
    if (grid.length > 0 && total > 0 && !sameGrid(grid, splitColumns(total, grid.length))) {
      table.widths = grid.map((w) => Math.round((w / total) * 10000) / 100);
    }

    // Pass 1: read every emitted cell, keeping its grid column and merge state.
    interface Slot {
      col: number;
      span: number;
      vMerge: string | null;
      cell: Record<string, unknown> | null;
    }
    const laid: Slot[][] = els(tbl, "tr").map((tr) => {
      const header = el(el(tr, "trPr"), "tblHeader") !== undefined;
      const slots: Slot[] = [];
      let col = 0;
      for (const tc of els(tr, "tc")) {
        const tcPr = el(tc, "tcPr");
        const span = numAttr(el(tcPr, "gridSpan"), "val") ?? 1;
        const vMergeEl = el(tcPr, "vMerge");
        const vMerge = vMergeEl ? (at(vMergeEl, "val") ?? "continue") : null;

        let cell: Record<string, unknown> | null = null;
        if (vMerge !== "continue") {
          cell = { blocks: this.walkBody(tc.children, ctx) };

          const fill = at(el(tcPr, "shd"), "fill");
          // A header row's grey came FROM `header`, so it is attributable and
          // not surfaced. Any other fill is the author's.
          if (fill !== undefined && fill !== "auto" && !(header && fill.toUpperCase() === HEADER_FILL)) {
            cell.shading = `#${fill.toUpperCase()}`;
          }
          const cellBorders = bordersFrom(el(tcPr, "tcBorders"), BOX_EDGES);
          if (cellBorders) cell.borders = cellBorders;
          const cellPadding = sidesFrom(el(tcPr, "tcMar"));
          if (cellPadding) cell.padding = cellPadding;
          const valign = at(el(tcPr, "vAlign"), "val");
          if (valign === "top" || valign === "center" || valign === "bottom") cell.valign = valign;
          if (span > 1) cell.colSpan = span;
        }

        slots.push({ col, span, vMerge, cell });
        col += span;
      }
      return slots;
    });

    // Pass 2: fold each run of continuations back into the cell that started it.
    laid.forEach((line, r) => {
      for (const slot of line) {
        if (slot.vMerge !== "restart" || !slot.cell) continue;
        let covered = 1;
        for (let below = r + 1; below < laid.length; below++) {
          const match = laid[below]!.find((s) => s.col === slot.col && s.vMerge === "continue");
          if (!match) break;
          covered++;
        }
        if (covered > 1) slot.cell.rowSpan = covered;
      }
    });

    table.rows = laid.map((line, r) => {
      const header = el(el(els(tbl, "tr")[r]!, "trPr"), "tblHeader") !== undefined;
      const cells = line.filter((s) => s.cell !== null).map((s) => s.cell!);
      return header ? { header: true, cells } : { cells };
    });

    return table as unknown as Block;
  }

  // ── Images ───────────────────────────────────────────────────────────────

  private parseDrawing(drawing: XmlNode): Block[] {
    const frame = el(drawing, "inline") ?? el(drawing, "anchor");
    if (!frame) return [];

    const blip = findFirst(frame, "blip");
    const rid = at(blip, "embed") ?? at(blip, "link");
    const rel = rid ? this.rels[rid] : undefined;
    if (!rel) return [];

    const target = rel.target.replace(/^\.\//, "");
    const partName = target.startsWith("/") ? target.slice(1) : `word/${target}`;
    const media = this.parts[partName];
    if (!media) return [];

    const ext = (partName.split(".").pop() ?? "").toLowerCase();
    const mime = ext === "png" ? "image/png" : ext === "jpg" || ext === "jpeg" ? "image/jpeg" : ext === "gif" ? "image/gif" : "image/png";

    const block: Any = { type: "image", src: `data:${mime};base64,${base64Encode(media)}` };

    const extent = el(frame, "extent");
    const cx = parseInt(at(extent, "cx") ?? "0", 10);
    const cy = parseInt(at(extent, "cy") ?? "0", 10);
    if (cx > 0) block.widthPx = Math.max(1, Math.round(cx / EMU_PER_PX));
    if (cy > 0) block.heightPx = Math.max(1, Math.round(cy / EMU_PER_PX));

    const descr = at(el(frame, "docPr"), "descr");
    if (descr) block.alt = descr;

    return [block as Block];
  }
}

// ── helpers ─────────────────────────────────────────────────────────────────

/**
 * Paragraph properties back into the model — the same set `paragraph`,
 * `heading` and a list item all accept.
 */
function paragraphPropsFrom(pPr: XmlNode | undefined): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  if (!pPr) return out;

  const jc = at(el(pPr, "jc"), "val");
  if (jc === "center" || jc === "right") out.align = jc;
  else if (jc === "both" || jc === "distribute" || jc === "justify") out.align = "justify";

  const spacing = el(pPr, "spacing");
  const before = numAttr(spacing, "before");
  if (before !== null) out.spaceBefore = points(before);
  const after = numAttr(spacing, "after");
  if (after !== null) out.spaceAfter = points(after);
  const line = numAttr(spacing, "line");
  if (line !== null && at(spacing, "lineRule") === "auto") {
    out.lineHeight = Math.round((line / 240) * 1000) / 1000;
  }

  const ind = el(pPr, "ind");
  const left = numAttr(ind, "left");
  if (left !== null) out.indentLeft = points(left);
  const right = numAttr(ind, "right");
  if (right !== null) out.indentRight = points(right);

  if (onFlag(el(pPr, "keepNext"))) out.keepNext = true;

  const fill = at(el(pPr, "shd"), "fill");
  if (fill !== undefined && fill !== "auto") out.shading = `#${fill.toUpperCase()}`;

  const borders = bordersFrom(el(pPr, "pBdr"), BOX_EDGES);
  if (borders) out.borders = borders;

  return out;
}

/** A paragraph carrying no runs and no properties — the adjacent-table pad. */
function isEmptyParagraph(block: Block | undefined): boolean {
  const b = block as Any;
  return (
    b?.type === "paragraph" &&
    Array.isArray(b.runs) &&
    b.runs.length === 0 &&
    Object.keys(b).length === 2
  );
}

/** A numeric attribute, or null when absent or unparseable. */
function numAttr(node: XmlNode | undefined, name: string): number | null {
  const raw = at(node, name);
  if (raw === undefined) return null;
  const n = Number(raw);
  return Number.isFinite(n) ? n : null;
}

/** Twips → points, kept exact for the halves the writer can emit. */
function points(twips: number): number {
  return Math.round((twips / 20) * 100) / 100;
}

const BOX_EDGES = ["top", "right", "bottom", "left"] as const;
const TABLE_EDGES = [...BOX_EDGES, "insideH", "insideV"] as const;

/** The header-cell grey, shared with the writer and both sibling engines. */
const HEADER_FILL = "E7E7E7";

/**
 * Is this exactly what the writer emits for a table that asked for nothing?
 *
 * Not a tolerance and not a heuristic: the writer's defaults are a fixed set,
 * and anything differing from them by one edge or one twip is the author's and
 * must survive the read.
 */
function isDefaultTableBorders(borders: Record<string, unknown>): boolean {
  const edges = Object.keys(borders);
  if (edges.length !== TABLE_EDGES.length) return false;
  return TABLE_EDGES.every((edge) => {
    const b = borders[edge] as Record<string, unknown> | undefined;
    // The default edge is single / 0.5pt / auto, which reads back as width
    // alone — style and colour are both the omitted default.
    return b !== undefined && Object.keys(b).length === 1 && b.width === 0.5;
  });
}

function isDefaultCellMargins(sides: Record<string, number>): boolean {
  return (
    Object.keys(sides).length === 4 &&
    sides.top === 3 && sides.bottom === 3 && sides.left === 5.4 && sides.right === 5.4
  );
}

function sameGrid(a: number[], b: number[]): boolean {
  return a.length === b.length && a.every((v, i) => v === b[i]);
}

/**
 * One border edge back into the model.
 *
 * Defaults are NOT surfaced: `style` only when it is not `single`, `color`
 * only when it is not `auto`. Otherwise reading a document written from a
 * model returns a bigger model than went in, and writing that model back
 * produces a different file.
 */
function borderFrom(node: XmlNode | undefined): Record<string, unknown> | null {
  if (!node) return null;
  const val = at(node, "val") ?? "single";
  if (val === "nil" || val === "none") return { style: "none" };

  const out: Record<string, unknown> = {};
  if (val !== "single") out.style = val;
  const sz = numAttr(node, "sz");
  if (sz !== null) out.width = Math.round((sz / 8) * 1000) / 1000;
  const color = at(node, "color");
  if (color !== undefined && color !== "auto") out.color = `#${color.toUpperCase()}`;
  return out;
}

function bordersFrom(container: XmlNode | undefined, edges: readonly string[]): Record<string, unknown> | null {
  if (!container) return null;
  const out: Record<string, unknown> = {};
  for (const edge of edges) {
    const border = borderFrom(el(container, edge));
    if (border) out[edge] = border;
  }
  return Object.keys(out).length > 0 ? out : null;
}

function sidesFrom(container: XmlNode | undefined): Record<string, number> | null {
  if (!container) return null;
  const out: Record<string, number> = {};
  for (const side of BOX_EDGES) {
    const w = numAttr(el(container, side), "w");
    if (w !== null) out[side] = points(w);
  }
  return Object.keys(out).length > 0 ? out : null;
}

/** True when a toggle property element is present and not explicitly off. */
function onFlag(node: XmlNode | undefined): boolean {
  if (!node) return false;
  const val = node.attrs["val"];
  return val === undefined || !["0", "false", "none", "off"].includes(val);
}

/** Depth-first search for all descendants with the given local name. */
function findAll(node: XmlNode, name: string): XmlNode[] {
  const out: XmlNode[] = [];
  for (const child of node.children) {
    if (child.name === name) out.push(child);
    out.push(...findAll(child, name));
  }
  return out;
}

function findFirst(node: XmlNode | undefined, name: string): XmlNode | undefined {
  if (!node) return undefined;
  for (const child of node.children) {
    if (child.name === name) return child;
    const nested = findFirst(child, name);
    if (nested) return nested;
  }
  return undefined;
}

/** Merge adjacent runs whose properties are identical (writer normalization). */
export function mergeRuns(runs: Run[]): Run[] {
  const out: Run[] = [];
  for (const run of runs) {
    if (run.text === "") continue;
    const prev = out[out.length - 1];
    if (prev && sameProps(prev, run)) {
      prev.text += run.text;
    } else {
      out.push({ ...run });
    }
  }
  return out;
}

function sameProps(a: Run, b: Run): boolean {
  // Every property a run can carry. A key missing from this list makes two
  // differently-formatted runs merge into one, silently taking the first
  // one's formatting — so it grows whenever the run model does.
  const keys = [
    "bold", "italic", "underline", "strike", "code", "link", "color", "highlight",
    "smallCaps", "size", "font", "letterSpacing",
  ] as const;
  return keys.every((k) => (a[k] ?? undefined) === (b[k] ?? undefined));
}

/** Build a nested ListItem tree from flat (ilvl-tagged) paragraphs. */
function buildTree(flat: FlatListItem[]): ListItem[] {
  const root: ListItem[] = [];
  const lastAtDepth: (ListItem | undefined)[] = [];

  for (const entry of flat) {
    const item: ListItem = { runs: entry.runs };
    let depth = Math.max(0, entry.ilvl);

    // Clamp orphan depths to the nearest existing parent.
    while (depth > 0 && lastAtDepth[depth - 1] === undefined) depth--;

    if (depth === 0) {
      root.push(item);
    } else {
      const parent = lastAtDepth[depth - 1]!;
      if (!parent.children) parent.children = [];
      parent.children.push(item);
    }
    lastAtDepth[depth] = item;
    lastAtDepth.length = depth + 1;
  }
  return root;
}

/** Keep `language` before `text` for stable JSON output (cosmetic only). */
function orderCodeKeys(block: Block): Block {
  const b = block as Any;
  if (b.language === undefined) return block;
  return { type: "code", language: b.language, text: b.text } as Block;
}
