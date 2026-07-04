/**
 * DocxReader — .docx bytes → Doc JSON model. Handles this package's own
 * writer output losslessly (round-trip) and tolerates Word-authored files:
 * headings via pStyle Heading1-9 OR outlineLvl, numPr lists with ilvl nesting,
 * hyperlinks via rels, images via blip r:embed, page breaks, bottom-border-only
 * paragraphs → hr. Unknown constructs degrade to plain paragraphs, never throw.
 * Mirrors PHP `Reader\DocxReader`.
 */

import { base64Encode } from "../helpers/image-size";
import type { Block, Doc, ListItem, Run } from "../schema/types";
import { unzipSync } from "../zip";
import { at, el, els, parseXml, type XmlNode } from "./xml";
import { EMU_PER_PX, SDT_TAG_CODE, SDT_TAG_QUOTE } from "../writer/docx-writer";

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
    return doc;
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
    let quoteBuf: Block[] | null = null;

    const flushList = (): void => {
      if (listBuf.length > 0) {
        blocks.push(...this.buildLists(listBuf));
        listBuf = [];
      }
    };
    const flushCode = (): void => {
      if (codeBuf !== null) {
        blocks.push({ type: "code", text: codeBuf.join("\n") });
        codeBuf = null;
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
            if (codeBuf === null) codeBuf = [];
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
    | { kind: "codeLine"; text: string }
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

    // Code line? (SDT-less tolerance path)
    if (/^(CodeBlock|SourceCode|HTMLPreformatted|Code)$/i.test(styleId)) {
      return { kind: "codeLine", text: this.plainText(p) };
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
      blocks.push({ type: "heading", level, runs });
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
      const para: Block = { type: "paragraph", runs };
      const jc = at(el(pPr, "jc"), "val");
      if (jc === "center" || jc === "right") (para as Any).align = jc;
      else if (jc === "both" || jc === "distribute" || jc === "justify") (para as Any).align = "justify";
      blocks.push(para);
    }
    for (const d of drawings) blocks.push(...this.parseDrawing(d));

    if (/^(Quote|IntenseQuote|BlockQuote|Blockquote)$/i.test(styleId)) {
      return { kind: "quoteParagraph", blocks };
    }
    return { kind: "blocks", blocks };
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

    const rStyle = at(el(rPr, "rStyle"), "val") ?? "";
    const asciiFont = (at(el(rPr, "rFonts"), "ascii") ?? "").toLowerCase();
    if (/^InlineCode$/i.test(rStyle) || MONO_FONTS.includes(asciiFont)) run.code = true;

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

  private parseTable(tbl: XmlNode, ctx: WalkCtx): Block {
    const rows = els(tbl, "tr").map((tr) => {
      const header = el(el(tr, "trPr"), "tblHeader") !== undefined;
      const cells = els(tr, "tc").map((tc) => ({ blocks: this.walkBody(tc.children, ctx) }));
      return header ? { header: true, cells } : { cells };
    });
    return { type: "table", rows } as Block;
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
  const keys = ["bold", "italic", "underline", "strike", "code", "link", "color", "highlight"] as const;
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
