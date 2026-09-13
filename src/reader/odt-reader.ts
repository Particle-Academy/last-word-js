/**
 * OpenDocument Text → the same document shape `DocxReader` returns. Mirrors PHP
 * `Reader\OdtReader`.
 *
 * ## What comes through
 *
 * - headings (`text:h`, with their outline level) and paragraphs
 * - bold, italic, underline and strike from AUTOMATIC styles, which is where ODF
 *   keeps direct formatting; a named style's formatting is the style's, as it is
 *   for the `.docx` and `.doc` readers
 * - hyperlinks (`text:a`)
 * - lists with nesting, numbered or bulleted by their list style
 * - tables, including header rows and merged cells (`colSpan` / `rowSpan`)
 * - spaces, tabs and line breaks written as `text:s`, `text:tab`,
 *   `text:line-break`
 * - page breaks set on a paragraph's automatic style
 * - the title from `meta.xml`
 *
 * ## What does not
 *
 * Images and frames, footnotes and endnotes, comments, tracked deletions, fonts,
 * sizes and colours, sections' layout and page geometry.
 *
 * ## Hostile input
 *
 * A part carrying a DOCTYPE is refused before it is parsed. A part larger than
 * 64 MB uncompressed is refused rather than inflated, and inflating stops at that
 * size whatever the entry declares. Repeated rows and columns are capped, each
 * repeat at 1,000 and all of them together at 100,000 cells, and element nesting
 * is walked with a depth limit.
 */

import type { Block, Doc } from "../schema/types";
import { unzipEntries } from "../zip/zip-reader";
import { isElement, parseDom, textContent, type DomElement } from "./odt/xml-dom";
import { Structure, clamp, phpInt, phpTrim, type ListEntry } from "./structure";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

const MAX_PART_BYTES = 64 * 1024 * 1024;
const MAX_REPEAT = 1000;
const MAX_DEPTH = 256;
const MAX_REPEATED_CELLS = 100_000;

const DECODER = new TextDecoder();

interface RawRun {
  text: string;
  collapsible: boolean;
  [flag: string]: unknown;
}

export class OdtReader {
  private automaticStyles = new Map<string, DomElement>();
  private listStyles = new Map<string, DomElement>();
  private repeatedCells = 0;

  read(bytes: Uint8Array): Doc {
    const parts = this.parts(bytes);
    const content = parse(parts["content.xml"], "content.xml", true)!;

    this.automaticStyles = new Map();
    this.listStyles = new Map();
    this.repeatedCells = 0;
    this.indexStyles(content, true);
    const styles = parse(parts["styles.xml"], "styles.xml", false);
    if (styles !== null) this.indexStyles(styles, false);

    const body = descendant(content, "text");
    const doc: Any = {};

    const meta = parse(parts["meta.xml"], "meta.xml", false);
    const title = meta !== null ? descendant(meta, "title") : null;
    if (title !== null && phpTrim(textContent(title)) !== "") doc.title = phpTrim(textContent(title));

    doc.blocks = body !== null ? this.blocks(body, 0) : [];
    return doc as Doc;
  }

  // ─── Archive ──────────────────────────────────────────────────────────

  private parts(bytes: Uint8Array): Record<string, string> {
    let entries: Record<string, Uint8Array>;
    try {
      entries = unzipEntries(bytes, ["content.xml", "styles.xml", "meta.xml"], MAX_PART_BYTES);
    } catch (e) {
      const message = (e as Error).message;
      const large = /^zip entry (\S+) is too large/.exec(message);
      if (large) throw new Error(`ODT part ${large[1]} is too large to read.`);
      throw new Error("Could not open the ODT archive.");
    }
    if (entries["content.xml"] === undefined) throw new Error("ODT archive has no content.xml.");

    const parts: Record<string, string> = {};
    for (const [name, data] of Object.entries(entries)) parts[name] = DECODER.decode(data);
    return parts;
  }

  // ─── Styles ───────────────────────────────────────────────────────────

  private indexStyles(root: DomElement, isContent: boolean): void {
    for (const [el, parent] of elements(root, 0)) {
      const name = el.attrs.get("style:name") ?? "";
      if (name === "") continue;
      if (el.local === "list-style") {
        if (!this.listStyles.has(name)) this.listStyles.set(name, el);
      } else if (el.local === "style" && isContent && parent.local === "automatic-styles") {
        this.automaticStyles.set(name, el);
      }
    }
  }

  /** Direct formatting from an automatic style and its automatic parents. */
  private flags(styleName: string): Record<string, boolean> {
    const flags: Record<string, boolean> = {};
    const seen = new Set<string>();
    let name = styleName;
    while (name !== "" && this.automaticStyles.has(name) && !seen.has(name)) {
      seen.add(name);
      const style = this.automaticStyles.get(name)!;
      const props = child(style, "text-properties");
      if (props !== null) {
        const weight = props.attrs.get("fo:font-weight") ?? "";
        if (weight !== "" && flags.bold === undefined) {
          flags.bold = weight === "bold" || (isNumeric(weight) && Math.trunc(Number(weight)) >= 600);
        }
        const italic = props.attrs.get("fo:font-style") ?? "";
        if (italic !== "" && flags.italic === undefined) flags.italic = italic === "italic" || italic === "oblique";
        const underline = props.attrs.get("style:text-underline-style") ?? "";
        if (underline !== "" && flags.underline === undefined) flags.underline = underline !== "none";
        const strike = props.attrs.get("style:text-line-through-style") ?? "";
        if (strike !== "" && flags.strike === undefined) flags.strike = strike !== "none";
      }
      name = style.attrs.get("style:parent-style-name") ?? "";
    }

    const out: Record<string, boolean> = {};
    for (const [key, on] of Object.entries(flags)) if (on) out[key] = true;
    return out;
  }

  private breaks(styleName: string, attribute: string): boolean {
    const style = this.automaticStyles.get(styleName);
    const props = style !== undefined ? child(style, "paragraph-properties") : null;
    return props !== null && props.attrs.get(attribute) === "page";
  }

  private listIsOrdered(listStyleName: string): boolean {
    const style = this.listStyles.get(listStyleName);
    if (style === undefined) return false;
    for (const level of style.children) {
      if (isElement(level) && level.attrs.get("text:level") === "1") return level.local === "list-level-style-number";
    }
    return false;
  }

  // ─── Blocks ───────────────────────────────────────────────────────────

  private blocks(container: DomElement, depth: number): Block[] {
    const blocks: Block[] = [];
    for (const node of container.children) {
      if (isElement(node) && depth <= MAX_DEPTH) blocks.push(...this.block(node, depth + 1));
    }
    return blocks;
  }

  private block(el: DomElement, depth: number): Block[] {
    switch (el.local) {
      case "h":
      case "p":
        return this.paragraph(el);
      case "list": {
        const entries: ListEntry[] = [];
        this.listEntries(el, 0, this.listIsOrdered(el.attrs.get("text:style-name") ?? ""), entries, depth);
        return Structure.lists(entries);
      }
      case "table":
        return [this.table(el, depth)];
      case "section":
      case "index-body":
      case "soft-page-break":
        return this.blocks(el, depth);
      default:
        return [];
    }
  }

  private paragraph(el: DomElement): Block[] {
    const style = el.attrs.get("text:style-name") ?? "";
    const runs = this.runs(el, this.flags(style), null);
    const out: Block[] = [];

    if (this.breaks(style, "fo:break-before")) out.push({ type: "pageBreak" } as Block);
    if (phpTrim(Structure.text(runs)) !== "") {
      if (el.local === "h") {
        const level = phpInt(el.attrs.get("text:outline-level") || "1");
        out.push({ type: "heading", level: clamp(level, 1, 6), runs } as Block);
      } else {
        out.push({ type: "paragraph", runs } as Block);
      }
    }
    if (this.breaks(style, "fo:break-after")) out.push({ type: "pageBreak" } as Block);
    return out;
  }

  private listEntries(list: DomElement, level: number, ordered: boolean, entries: ListEntry[], depth: number): void {
    if (depth > MAX_DEPTH) return;
    for (const item of list.children) {
      if (!isElement(item) || (item.local !== "list-item" && item.local !== "list-header")) continue;
      for (const node of item.children) {
        if (!isElement(node)) continue;
        if (node.local === "p" || node.local === "h") {
          const runs = this.runs(node, this.flags(node.attrs.get("text:style-name") ?? ""), null);
          if (phpTrim(Structure.text(runs)) !== "") entries.push({ ilvl: Math.min(8, level), ordered, runs });
        } else if (node.local === "list") {
          this.listEntries(node, level + 1, ordered, entries, depth + 1);
        }
      }
    }
  }

  private table(table: DomElement, depth: number): Block {
    const rows: Any[] = [];
    this.rows(table, false, rows, depth);
    return { type: "table", rows } as Block;
  }

  private rows(container: DomElement, header: boolean, rows: Any[], depth: number): void {
    for (const node of container.children) {
      if (!isElement(node)) continue;
      if (node.local === "table-header-rows") {
        this.rows(node, true, rows, depth);
      } else if (node.local === "table-rows" || node.local === "table-row-group") {
        this.rows(node, header, rows, depth);
      } else if (node.local === "table-row") {
        const cells: Any[] = [];
        for (const cell of node.children) {
          if (!isElement(cell) || cell.local !== "table-cell") continue; // covered cells belong to the cell that spans them
          const out: Any = { blocks: this.blocks(cell, depth + 1) };
          const colSpan = phpInt(cell.attrs.get("table:number-columns-spanned"));
          const rowSpan = phpInt(cell.attrs.get("table:number-rows-spanned"));
          if (colSpan > 1) out.colSpan = Math.min(colSpan, MAX_REPEAT);
          if (rowSpan > 1) out.rowSpan = Math.min(rowSpan, MAX_REPEAT);
          const repeat = Math.max(1, Math.min(phpInt(cell.attrs.get("table:number-columns-repeated")), MAX_REPEAT));
          cells.push(out);
          for (let i = 1; i < repeat && this.repeatedCells < MAX_REPEATED_CELLS; i++) {
            cells.push(out);
            this.repeatedCells++;
          }
        }
        const row = header ? { header: true, cells } : { cells };
        // A repeated row is usually spreadsheet-style filler; only a row with
        // content is worth repeating, and never unboundedly.
        const hasContent = cells.some((c) => c.blocks.length > 0);
        const repeat = hasContent ? Math.max(1, Math.min(phpInt(node.attrs.get("table:number-rows-repeated")), MAX_REPEAT)) : 1;
        rows.push(row);
        for (let i = 1; i < repeat && this.repeatedCells < MAX_REPEATED_CELLS; i++) {
          rows.push(row);
          this.repeatedCells += Math.max(1, cells.length);
        }
      }
    }
  }

  // ─── Inline ───────────────────────────────────────────────────────────

  private runs(el: DomElement, flags: Record<string, boolean>, link: string | null): Any[] {
    const raw: RawRun[] = [];
    this.collect(el, flags, link, raw, 0);

    // ODF collapses white space in text nodes and ignores it at the edges of a
    // paragraph; spaces written as text:s are real.
    if (raw.length > 0) {
      if (raw[0]!.collapsible) raw[0]!.text = raw[0]!.text.replace(/^ +/, "");
      const last = raw[raw.length - 1]!;
      if (last.collapsible) last.text = last.text.replace(/ +$/, "");
    }

    return Structure.mergeRuns(
      raw.map((r) => {
        const { collapsible: _collapsible, ...run } = r;
        return run;
      }),
    );
  }

  private collect(node: DomElement, flags: Record<string, boolean>, link: string | null, out: RawRun[], depth: number): void {
    if (depth > MAX_DEPTH) return;
    for (const child of node.children) {
      if (typeof child === "string") {
        let text = child.replace(/[ \t\r\n]+/g, " ");
        // A collapsed space right after another one is a single space.
        const last = out[out.length - 1];
        if (last !== undefined && last.text.endsWith(" ") && text.startsWith(" ") && last.collapsible) {
          text = text.replace(/^ +/, "");
        }
        push(out, text, flags, link, true);
        continue;
      }

      switch (child.local) {
        case "span":
          this.collect(child, { ...flags, ...this.flags(child.attrs.get("text:style-name") ?? "") }, link, out, depth + 1);
          break;
        case "a": {
          const href = child.attrs.get("xlink:href") ?? "";
          this.collect(child, flags, href !== "" ? href : link, out, depth + 1);
          break;
        }
        case "s": {
          const count = Math.max(1, Math.min(phpInt(child.attrs.get("text:c") || "1"), MAX_REPEAT));
          push(out, " ".repeat(count), flags, link, false);
          break;
        }
        case "tab":
          push(out, "\t", flags, link, false);
          break;
        case "line-break":
          push(out, "\n", flags, link, false);
          break;
        case "note":
        case "annotation":
        case "annotation-end":
        case "bookmark":
        case "bookmark-start":
        case "bookmark-end":
        case "reference-mark":
        case "change":
        case "change-start":
        case "change-end":
        case "frame":
        case "soft-page-break":
          break;
        default:
          // Fields and other inline wrappers display their text content.
          this.collect(child, flags, link, out, depth + 1);
      }
    }
  }
}

function push(out: RawRun[], text: string, flags: Record<string, boolean>, link: string | null, collapsible: boolean): void {
  if (text === "") return;
  const run: RawRun = { text, ...flags, collapsible };
  if (link !== null) run.link = link;
  out.push(run);
}

function parse(xml: string | undefined, name: string, required: boolean): DomElement | null {
  if (xml === undefined) return null;
  if (/<!DOCTYPE/i.test(xml)) {
    throw new Error(`ODT part ${name} carries a DOCTYPE, which an ODT never does; refusing to parse it.`);
  }
  try {
    return parseDom(xml);
  } catch {
    if (required) throw new Error(`Could not parse ${name}.`);
    return null;
  }
}

/** Every element with its parent, depth-first, depth-limited. */
function elements(node: DomElement, depth: number, out: [DomElement, DomElement][] = []): [DomElement, DomElement][] {
  if (depth > MAX_DEPTH) return out;
  for (const c of node.children) {
    if (isElement(c)) {
      out.push([c, node]);
      elements(c, depth + 1, out);
    }
  }
  return out;
}

function child(parent: DomElement, local: string): DomElement | null {
  for (const c of parent.children) if (isElement(c) && c.local === local) return c;
  return null;
}

/** The first descendant with a local name, breadth-first, as PHP walks it. */
function descendant(root: DomElement, local: string): DomElement | null {
  const queue: DomElement[] = [root];
  let visited = 0;
  for (let head = 0; head < queue.length && visited++ < 1_000_000; head++) {
    for (const c of queue[head]!.children) {
      if (isElement(c)) {
        if (c.local === local) return c;
        queue.push(c);
      }
    }
  }
  return null;
}

/** PHP `is_numeric` for the attribute values ODF writes. */
function isNumeric(s: string): boolean {
  return /^[ \t\n\r\v\f]*[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?[ \t\n\r\v\f]*$/.test(s);
}
