/**
 * Word 97-2003 binary `.doc` → the same document shape `DocxReader` returns.
 * Mirrors PHP `Reader\DocReader`.
 *
 * ## What comes through
 *
 * - paragraphs and their text, from every piece of a fast-saved file
 * - headings, by the style's built-in identifier (Heading 1-9), so a localised
 *   style name ("Überschrift 1") is still a heading
 * - bold, italic, underline and strike applied directly to text
 * - hyperlinks, from `HYPERLINK` fields; other fields keep their displayed
 *   result and drop their instructions
 * - bulleted and numbered lists with nesting
 * - tables, with header rows and a paragraph per cell paragraph
 * - page breaks
 *
 * ## What does not
 *
 * Formatting inherited from styles (only direct formatting is read), fonts,
 * sizes and colours, images and embedded objects, text boxes, headers, footers,
 * footnotes, endnotes and comments, merged cells (each cell is read as written),
 * and the document title. Nested tables are flattened into their outer cell.
 *
 * ## Refused
 *
 * A compound file that is not a Word document (`.xls`, `.ppt`, `.msg`), a Word 6
 * or 95 file, and an encrypted file each raise `UnsupportedFormatException`. A
 * damaged container raises a plain `Error` naming what was wrong.
 */

import { UnsupportedFormatException } from "../exceptions";
import type { Block, Doc } from "../schema/types";
import { CompoundFile } from "./doc/compound-file";
import { WordBinary, type ParagraphProps } from "./doc/word-binary";
import { Structure, phpTrim, type ListEntry } from "./structure";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

/** ilfo 2047 means "explicitly not in a list" in Word 2003+ files. */
const ILFO_NOT_A_LIST = 2047;

interface PendingRun {
  text: string;
  flags: Record<string, boolean>;
  link: string | null;
}

interface Field {
  instruction: string;
  inResult: boolean;
  link: string | null;
}

export class DocReader {
  private blocks: Block[] = [];
  private runs: PendingRun[] = [];
  private listEntries: ListEntry[] = [];
  private tableRows: Any[] | null = null;
  private rowCells: Any[] = [];
  private cellBlocks: Block[] = [];
  private fields: Field[] = [];
  private word!: WordBinary;

  read(bytes: Uint8Array): Doc {
    const file = CompoundFile.fromBytes(bytes);
    if (!file.hasStream("WordDocument")) throw notWord(file);

    this.word = new WordBinary(file);
    for (const character of this.word.characters()) this.consume(character.char, character.fc);
    // A document whose last paragraph lacks a mark still has that paragraph.
    if (this.runs.length > 0) this.endParagraph(this.word.paragraphAt(0), null);
    this.flushLists();
    this.flushTable();

    return { blocks: this.blocks };
  }

  private consume(char: string, fc: number): void {
    switch (char) {
      case "\x13": // field begin
        this.fields.push({ instruction: "", inResult: false, link: null });
        return;
      case "\x14": { // field separator: instruction done, result follows
        const top = this.fields[this.fields.length - 1];
        if (top !== undefined) {
          top.inResult = true;
          top.link = hyperlink(top.instruction);
        }
        return;
      }
      case "\x15": // field end
        this.fields.pop();
        return;
    }

    // Inside a field's instruction nothing is displayed.
    const top = this.fields[this.fields.length - 1];
    if (top !== undefined && !top.inResult) {
      top.instruction += char;
      return;
    }

    switch (char) {
      case "\r":
        this.endParagraph(this.word.paragraphAt(fc), null);
        return;
      case "\x07":
        this.endParagraph(this.word.paragraphAt(fc), "cell");
        return;
      case "\x0C": // page or section break
        this.endParagraph(this.word.paragraphAt(fc), null);
        this.flushLists();
        this.flushTable();
        this.blocks.push({ type: "pageBreak" } as Block);
        return;
      case "\x0B": // line break inside a paragraph
        this.append("\n", fc);
        return;
      case "\x1E": // non-breaking hyphen
        this.append("-", fc);
        return;
      case "\x1F": // optional hyphen
      case "\x01": // picture or embedded object anchor
      case "\x02": // automatic footnote reference
      case "\x03": // footnote separator
      case "\x04": // footnote continuation
      case "\x05": // annotation reference
      case "\x08": // drawn object anchor
      case "": // second half of a surrogate pair
        return;
    }

    this.append(char, fc);
  }

  private append(text: string, fc: number): void {
    let link: string | null = null;
    for (let i = this.fields.length - 1; i >= 0 && link === null; i--) link = this.fields[i]!.link;
    const flags = this.word.charactersAt(fc);

    // One entry per formatting change, not per character.
    const last = this.runs[this.runs.length - 1];
    if (last !== undefined && sameFlags(last.flags, flags) && last.link === link) {
      last.text += text;
      return;
    }
    this.runs.push({ text, flags, link });
  }

  private endParagraph(props: ParagraphProps, mark: "cell" | null): void {
    const runs = Structure.mergeRuns(
      this.runs.map((r) => ({ text: r.text, ...r.flags, ...(r.link !== null ? { link: r.link } : {}) })),
    );
    this.runs = [];
    const hasText = phpTrim(Structure.text(runs)) !== "";

    if (props.inTable) {
      this.flushLists();
      this.tableRows ??= [];

      if (mark === "cell" && props.rowEnd) {
        this.tableRows.push(props.header ? { header: true, cells: this.rowCells } : { cells: this.rowCells });
        this.rowCells = [];
        this.cellBlocks = [];
        return;
      }

      if (hasText) this.cellBlocks.push({ type: "paragraph", runs } as Block);
      if (mark === "cell") {
        this.rowCells.push({ blocks: this.cellBlocks });
        this.cellBlocks = [];
      }
      return;
    }

    this.flushTable();
    if (!hasText) return;

    if (props.ilfo > 0 && props.ilfo !== ILFO_NOT_A_LIST) {
      // Orderedness is the list's, read at its top level, as DocxReader reads a
      // numbering definition.
      this.listEntries.push({ ilvl: props.ilvl, ordered: this.word.listIsOrdered(props.ilfo, 0), runs });
      return;
    }

    this.flushLists();
    const level = this.word.headingLevel(props.istd);
    this.blocks.push(
      (level !== null ? { type: "heading", level: Math.min(6, level), runs } : { type: "paragraph", runs }) as Block,
    );
  }

  private flushLists(): void {
    if (this.listEntries.length > 0) {
      this.blocks.push(...Structure.lists(this.listEntries));
      this.listEntries = [];
    }
  }

  private flushTable(): void {
    if (this.tableRows === null) return;
    if (this.rowCells.length > 0) this.tableRows.push({ cells: this.rowCells });
    if (this.tableRows.length > 0) this.blocks.push({ type: "table", rows: this.tableRows } as Block);
    this.tableRows = null;
    this.rowCells = [];
    this.cellBlocks = [];
  }
}

function sameFlags(a: Record<string, boolean>, b: Record<string, boolean>): boolean {
  const ka = Object.keys(a);
  const kb = Object.keys(b);
  return ka.length === kb.length && ka.every((k, i) => k === kb[i] && a[k] === b[k]);
}

/** The target of a `HYPERLINK` field instruction, or null for any other field. */
export function hyperlink(instruction: string): string | null {
  const m = /^[ \t\n\v\f\r]*HYPERLINK(?![A-Za-z0-9_])([\s\S]*)$/i.exec(instruction);
  if (!m) return null;
  const rest = m[1]!;
  let target: string | null = /^[ \t\n\v\f\r]*"([^"]*)"/.exec(rest)?.[1] ?? null;
  const anchor = /\\l[ \t\n\v\f\r]+"([^"]*)"/i.exec(rest);
  if (anchor) target = (target ?? "") + "#" + anchor[1];
  return target !== null && target !== "" ? target : null;
}

function notWord(file: CompoundFile): UnsupportedFormatException {
  const names = file.streamNames();
  let format = "cfb";
  let what = "a compound file that is not a Word document";
  if (names.includes("Workbook") || names.includes("Book")) {
    format = "xls";
    what = "an Excel 97-2003 workbook (.xls)";
  } else if (names.includes("PowerPoint Document")) {
    format = "ppt";
    what = "a PowerPoint 97-2003 presentation (.ppt)";
  } else if (names.includes("__properties_version1.0")) {
    format = "msg";
    what = "an Outlook message (.msg)";
  }
  return new UnsupportedFormatException(
    format,
    `This is ${what}, not a Word document. last-word reads .docx, .doc, .odt and .rtf.`,
  );
}
