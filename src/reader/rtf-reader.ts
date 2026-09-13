/**
 * RTF → the same document shape the other readers return. Mirrors PHP
 * `Reader\RtfReader`.
 *
 * A tokenizer over control words, control symbols, groups and text, with the
 * state RTF actually scopes to groups: character formatting, paragraph
 * properties, the current destination and the Unicode fallback count.
 *
 * ## What comes through
 *
 * - paragraphs, and headings: a paragraph whose style is named "heading N" or
 *   that carries an outline level
 * - bold, italic, underline and strike applied directly. RTF repeats a style's
 *   formatting inline, so what a paragraph or character style sets (a heading's
 *   bold) is subtracted again, and a hyperlink's underline is not reported:
 *   LibreOffice writes the link style's underline inline without naming the
 *   style, so it cannot be told from one the author applied
 * - hyperlinks from `HYPERLINK` fields; other fields keep their displayed result
 * - lists with nesting, numbered or bulleted from the list table
 * - tables, with header rows where the file marks one (`\trhdr`)
 * - Unicode (`\uN` with the `\ucN` fallback skipped, surrogate pairs joined) and
 *   `\'hh` bytes in the document's code page (`\ansicpg`, 1252 when absent)
 * - line breaks, tabs, page breaks, and the title from `{\info{\title}}`
 *
 * ## What does not
 *
 * Images and objects, footnotes, annotations, headers and footers, fonts, sizes
 * and colours, merged cells, Word 6/95-style `\pn` numbering (those items read as
 * paragraphs), a font's own `\fcharset`, and header rows in a file that does not
 * mark them with `\trhdr` (LibreOffice does not). Nested tables are flattened
 * into their outer cell.
 *
 * ## Hostile input
 *
 * Group nesting is capped; `\bin` data is skipped by its declared length, capped
 * by what remains; numeric parameters are bounded; every read is in range.
 */

import type { Block, Doc } from "../schema/types";
import { DEFAULT_CODE_PAGE, REPLACEMENT, decodeCodePage, fromCodePoint } from "./code-page";
import { hyperlink } from "./doc-reader";
import { Structure, phpTrim, type ListEntry } from "./structure";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

const MAX_DEPTH = 10_000;

/** Groups whose whole content is metadata or not body text. */
const SKIP_DESTINATIONS = new Set([
  "fonttbl", "colortbl", "pict", "header", "headerl", "headerr", "headerf", "footer", "footerl",
  "footerr", "footerf", "footnote", "annotation", "themedata", "colorschememapping", "latentstyles",
  "datastore", "generator", "rsidtbl", "listtext", "pntext", "pn", "object", "shp", "nonshppict",
  "xe", "tc", "txe", "filetbl", "revtbl", "protusertbl", "docvar", "userprops", "mmathPr",
  "nesttableprops", "fldtype", "author", "operator", "keywords", "comment", "doccomm", "subject",
  "company", "category", "manager", "hlinkbase", "creatim", "revtim", "printim", "buptim",
]);

const UNDERLINES = new Set([
  "uld", "uldash", "uldashd", "uldashdd", "uldb", "ulhwave", "ulldash", "ulth", "ulthd", "ulthdash",
  "ulthdashd", "ulthdashdd", "ulthldash", "ululdbwave", "ulw", "ulwave",
]);

/** Group-opening words that change what the group's text is. */
const DESTINATIONS = new Map([
  ["stylesheet", "stylesheet"],
  ["listtable", "listtable"],
  ["listoverridetable", "listoverridetable"],
  ["info", "info"],
  ["fldinst", "fldinst"],
  ["fldrslt", "body"],
]);

const NFC_BULLET = 23;
const NFC_NONE = 255;

const SYMBOLS: Record<string, string> = {
  emdash: String.fromCharCode(0x2014),
  endash: String.fromCharCode(0x2013),
  bullet: String.fromCharCode(0x2022),
  lquote: String.fromCharCode(0x2018),
  rquote: String.fromCharCode(0x2019),
  ldblquote: String.fromCharCode(0x201c),
  rdblquote: String.fromCharCode(0x201d),
  emspace: " ",
  enspace: " ",
  qmspace: " ",
  tab: "\t",
  line: "\n",
};

const BACKSLASH = 0x5c;

interface State {
  dest: string;
  uc: number;
  skip: number;
  b: boolean;
  i: boolean;
  ul: boolean;
  strike: boolean;
  intbl: boolean;
  style: number;
  ls: number;
  ilvl: number;
  outline: number | null;
  cs: number | null;
  link: string | null;
  opensField: boolean;
  ignorable: boolean;
  first: boolean;
  nfc?: number;
  overrideList?: number;
}

interface PendingRun {
  text: string;
  flags: { bold: boolean; italic: boolean; underline: boolean; strike: boolean };
  cs: number | null;
  link: string | null;
}

interface Style {
  name: string;
  flags: Record<string, boolean>;
  outline: number | null;
}

interface StyleEntry {
  key: string;
  name: string;
  b: boolean;
  i: boolean;
  ul: boolean;
  strike: boolean;
  outline: number | null;
}

function initialState(): State {
  return {
    dest: "body", uc: 1, skip: 0,
    b: false, i: false, ul: false, strike: false,
    intbl: false, style: 0, ls: 0, ilvl: 0, outline: null,
    cs: null, link: null, opensField: false, ignorable: false, first: true,
  };
}

function isAlpha(c: number): boolean {
  return (c >= 0x41 && c <= 0x5a) || (c >= 0x61 && c <= 0x7a);
}

function isDigit(c: number): boolean {
  return c >= 0x30 && c <= 0x39;
}

function isHex(c: number): boolean {
  return isDigit(c) || (c >= 0x41 && c <= 0x46) || (c >= 0x61 && c <= 0x66);
}

/** Ends a run of plain text: group and control delimiters, line ends, and 8-bit bytes. */
function isSpecial(c: number): boolean {
  return c === 0x7b || c === 0x7d || c === BACKSLASH || c === 0x0d || c === 0x0a || c >= 0x80;
}

export class RtfReader {
  private src: Uint8Array = new Uint8Array(0);
  private pos = 0;
  private state: State = initialState();
  private stack: State[] = [];
  private runs: PendingRun[] = [];
  private blocks: Block[] = [];
  private listEntries: ListEntry[] = [];
  private tableRows: Any[] | null = null;
  private rowCells: Any[] = [];
  private cellBlocks: Block[] = [];
  private rowHeader = false;
  private fields: { instruction: string }[] = [];
  private styles = new Map<string, Style>();
  private styleEntry: StyleEntry | null = null;
  private lists = new Map<number, number[]>();
  private listLevels: number[] = [];
  private overrides = new Map<number, number>();
  private title = "";
  private codepage = DEFAULT_CODE_PAGE;
  private highSurrogate: number | null = null;
  private pendingBytes: number[] = [];

  read(bytes: Uint8Array): Doc {
    const bom = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf;
    this.src = bom ? bytes.subarray(3) : bytes;
    this.pos = 0;
    this.state = initialState();
    this.stack = [];
    this.runs = [];
    this.blocks = [];
    this.listEntries = [];
    this.rowCells = [];
    this.cellBlocks = [];
    this.fields = [];
    this.styles = new Map();
    this.lists = new Map();
    this.listLevels = [];
    this.overrides = new Map();
    this.tableRows = null;
    this.styleEntry = null;
    this.rowHeader = false;
    this.title = "";
    this.pendingBytes = [];
    this.codepage = DEFAULT_CODE_PAGE;
    this.highSurrogate = null;

    this.parse();
    this.flushBytes();
    if (this.runs.length > 0) this.endParagraph(null);
    this.flushLists();
    this.flushTable();

    const doc: Any = {};
    if (phpTrim(this.title) !== "") doc.title = phpTrim(this.title);
    doc.blocks = this.blocks;
    return doc as Doc;
  }

  private parse(): void {
    const src = this.src;
    const length = src.length;
    while (this.pos < length) {
      const c = src[this.pos]!;
      const isByte = c >= 0x80 || (c === BACKSLASH && src[this.pos + 1] === 0x27);
      if (!isByte) this.flushBytes();

      if (c === 0x7b) {
        if (this.stack.length >= MAX_DEPTH) throw new Error("RTF nests groups too deeply to read.");
        this.stack.push(this.state);
        this.state = { ...this.state, opensField: false, ignorable: false, first: true, skip: 0 };
        this.pos++;
        continue;
      }
      if (c === 0x7d) {
        this.closeGroup();
        this.pos++;
        continue;
      }
      if (c === BACKSLASH) {
        this.controlWord();
        continue;
      }
      if (c === 0x0d || c === 0x0a) {
        this.pos++;
        continue;
      }
      this.state.first = false;
      if (isByte) {
        this.byte(c);
        this.pos++;
      } else if (this.state.skip > 0) {
        this.text(String.fromCharCode(c));
        this.pos++;
      } else {
        // A run of plain text in one step rather than a call per letter.
        let end = this.pos;
        while (end < length && !isSpecial(src[end]!)) end++;
        this.text(latin1(src.subarray(this.pos, end)));
        this.pos = end;
      }
    }
  }

  private closeGroup(): void {
    const closing = this.state;
    const parent = this.stack.pop();
    if (parent === undefined) return; // an unbalanced brace ends nothing
    this.state = parent;

    if (closing.opensField) this.fields.pop();
    if (closing.dest === "style" && this.styleEntry !== null) {
      const e = this.styleEntry;
      const flags: Record<string, boolean> = {};
      if (e.b) flags.bold = true;
      if (e.i) flags.italic = true;
      if (e.ul) flags.underline = true;
      if (e.strike) flags.strike = true;
      this.styles.set(e.key, { name: phpTrim(e.name.replace(/;+$/, "")), flags, outline: e.outline });
      this.styleEntry = null;
    }
    if (closing.dest === "listlevel" && this.state.dest === "list") {
      this.listLevels.push(closing.nfc ?? NFC_BULLET);
    }
  }

  private controlWord(): void {
    const src = this.src;
    const length = src.length;
    const next = this.pos + 1 < length ? src[this.pos + 1]! : -1;

    // Control symbols.
    if (next < 0 || !isAlpha(next)) {
      this.pos += 2;
      switch (next) {
        case BACKSLASH:
        case 0x7b:
        case 0x7d:
          this.state.first = false;
          this.text(String.fromCharCode(next));
          break;
        case 0x27: { // \'hh
          const h1 = src[this.pos];
          const h2 = src[this.pos + 1];
          this.pos += 2;
          if (h1 !== undefined && h2 !== undefined && isHex(h1) && isHex(h2)) {
            this.state.first = false;
            this.byte(parseInt(String.fromCharCode(h1, h2), 16));
          }
          break;
        }
        case 0x2a: // \* an ignorable destination
          this.state.ignorable = true;
          break;
        case 0x7e: // \~
          this.text(String.fromCharCode(0xa0));
          break;
        case 0x5f: // \_
          this.text("-");
          break;
        case 0x2d: // \- optional hyphen
          break;
        case 0x0d:
        case 0x0a:
          this.word("par", null);
          break;
      }
      return;
    }

    let end = this.pos + 1;
    while (end < length && end - (this.pos + 1) < 32 && isAlpha(src[end]!)) end++;
    const word = latin1(src.subarray(this.pos + 1, end));
    let after = end;
    let param: number | null = null;
    let digitsStart = after;
    if (src[after] === 0x2d && after + 1 < length && isDigit(src[after + 1]!)) digitsStart = after + 1;
    if (digitsStart < length && isDigit(src[digitsStart]!)) {
      let digitsEnd = digitsStart;
      while (digitsEnd < length && digitsEnd - digitsStart < 10 && isDigit(src[digitsEnd]!)) digitsEnd++;
      param = Number(latin1(src.subarray(after, digitsEnd)));
      after = digitsEnd;
    }
    if (after < length && src[after] === 0x20) after++;
    this.pos = after;

    if (word === "bin") {
      this.pos = Math.min(length, this.pos + Math.max(0, param ?? 0));
      return;
    }

    this.word(word, param);
  }

  private word(word: string, param: number | null): void {
    const state = this.state;
    const first = state.first;
    state.first = false;

    // Destinations, decided by the first word of a group.
    if (state.ignorable) {
      state.ignorable = false;
      const known =
        ["fldinst", "listtable", "listoverridetable"].includes(word) ||
        (state.dest === "stylesheet" && ["s", "cs", "ds", "ts"].includes(word));
      if (!known) {
        state.dest = "skip";
        return;
      }
    }
    if (state.dest === "skip") return;
    if (first || SKIP_DESTINATIONS.has(word)) {
      if (SKIP_DESTINATIONS.has(word)) {
        state.dest = "skip";
        return;
      }
      const destination = DESTINATIONS.get(word);
      if (destination !== undefined) {
        state.dest = destination;
        const field = this.fields[this.fields.length - 1];
        if (word === "fldrslt" && field !== undefined) {
          state.link = hyperlink(field.instruction) ?? state.link;
        }
        return;
      }
    }

    switch (state.dest) {
      case "stylesheet":
        if (first && ["s", "cs", "ds", "ts"].includes(word)) {
          state.dest = "style";
          const key = word === "s" ? "s" + (param ?? 0) : word === "cs" ? "c" + (param ?? 0) : "";
          this.styleEntry = { key, name: "", b: false, i: false, ul: false, strike: false, outline: null };
        }
        return;
      case "style": {
        const e = this.styleEntry;
        if (e !== null) {
          if (word === "b") e.b = param !== 0;
          else if (word === "i") e.i = param !== 0;
          else if (word === "strike" || word === "striked") e.strike = param !== 0;
          else if (isUnderlineOn(word)) e.ul = param !== 0;
          else if (word === "ulnone") e.ul = false;
          else if (word === "outlinelevel") e.outline = param;
        }
        return;
      }
      case "listtable":
        if (word === "list") {
          state.dest = "list";
          this.listLevels = [];
        }
        return;
      case "list":
        if (word === "listlevel") state.dest = "listlevel";
        else if (word === "listid" && param !== null) this.lists.set(param, this.listLevels);
        return;
      case "listlevel":
        if (word === "levelnfc" || word === "levelnfcn") state.nfc = param ?? 0;
        else if (word === "leveltext" || word === "levelnumbers") state.dest = "listlevelskip";
        return;
      case "listoverridetable":
        if (word === "listoverride") state.dest = "listoverride";
        return;
      case "listoverride":
        if (word === "listid") state.overrideList = param ?? 0;
        else if (word === "ls" && param !== null) this.overrides.set(param, state.overrideList ?? 0);
        return;
      case "info":
        if (word === "title") state.dest = "title";
        return;
      case "fldinst":
      case "title":
      case "listlevelskip":
        return;
    }

    // Body.
    if (word === "u" && param !== null) {
      this.unicode(param < 0 ? param + 65536 : param);
      state.skip = state.uc;
      return;
    }
    if (state.skip > 0) {
      state.skip--;
      return;
    }

    switch (word) {
      case "uc": state.uc = Math.max(0, Math.min(10, param ?? 0)); break;
      case "ansicpg": this.codepage = param ?? 0; break;
      case "field":
        this.fields.push({ instruction: "" });
        state.opensField = true;
        break;
      case "par":
      case "sect":
        this.endParagraph(null);
        break;
      case "cell":
      case "nestcell":
        this.endParagraph(word);
        break;
      case "row": this.endRow(); break;
      case "trowd": this.rowHeader = false; break;
      case "trhdr": this.rowHeader = true; break;
      case "page":
        this.endParagraph(null);
        this.flushLists();
        this.flushTable();
        this.blocks.push({ type: "pageBreak" } as Block);
        break;
      case "pard":
        state.intbl = false;
        state.style = 0;
        state.ls = 0;
        state.ilvl = 0;
        state.outline = null;
        break;
      case "plain":
        state.b = state.i = state.ul = state.strike = false;
        state.cs = null;
        break;
      case "cs": state.cs = param ?? 0; break;
      case "intbl": state.intbl = true; break;
      case "s": state.style = param ?? 0; break;
      case "ls": state.ls = param ?? 0; break;
      case "ilvl": state.ilvl = Math.max(0, Math.min(8, param ?? 0)); break;
      case "outlinelevel": state.outline = param; break;
      case "b": state.b = param !== 0; break;
      case "i": state.i = param !== 0; break;
      case "strike":
      case "striked":
        state.strike = param !== 0;
        break;
      case "ulnone": state.ul = false; break;
      default:
        if (isUnderlineOn(word)) state.ul = param !== 0;
        else if (Object.prototype.hasOwnProperty.call(SYMBOLS, word)) this.text(SYMBOLS[word]!);
    }
  }

  private unicode(unit: number): void {
    if (unit >= 0xdc00 && unit <= 0xdfff && this.highSurrogate !== null) {
      const codePoint = 0x10000 + ((this.highSurrogate - 0xd800) << 10) + (unit - 0xdc00);
      this.highSurrogate = null;
      this.text(fromCodePoint(codePoint));
      return;
    }
    if (unit >= 0xd800 && unit <= 0xdbff) {
      if (this.highSurrogate !== null) this.emit(REPLACEMENT);
      this.highSurrogate = unit;
      return;
    }
    this.text(fromCodePoint(unit)); // a lone low surrogate decodes as U+FFFD
  }

  private byte(byte: number): void {
    if (this.state.skip > 0) {
      this.state.skip--;
      return;
    }
    this.pendingBytes.push(byte);
  }

  private flushBytes(): void {
    if (this.pendingBytes.length === 0) return;
    const bytes = this.pendingBytes;
    this.pendingBytes = [];
    this.emit(decodeCodePage(bytes, this.codepage));
  }

  private text(text: string): void {
    if (this.state.skip > 0) {
      this.state.skip--;
      return;
    }
    this.emit(text);
  }

  private emit(text: string): void {
    if (this.highSurrogate !== null) {
      // Half a pair followed by anything but its other half.
      this.highSurrogate = null;
      this.emit(REPLACEMENT);
    }
    const state = this.state;

    switch (state.dest) {
      case "body": {
        const flags = { bold: state.b, italic: state.i, underline: state.ul, strike: state.strike };
        const last = this.runs[this.runs.length - 1];
        if (
          last !== undefined &&
          last.flags.bold === flags.bold && last.flags.italic === flags.italic &&
          last.flags.underline === flags.underline && last.flags.strike === flags.strike &&
          last.cs === state.cs && last.link === state.link
        ) {
          last.text += text;
          break;
        }
        this.runs.push({ text, flags, cs: state.cs, link: state.link });
        break;
      }
      case "fldinst": {
        const field = this.fields[this.fields.length - 1];
        if (field !== undefined) field.instruction += text;
        break;
      }
      case "style":
        if (this.styleEntry !== null) this.styleEntry.name += text;
        break;
      case "title":
        this.title += text;
        break;
    }
  }

  // ─── Blocks ───────────────────────────────────────────────────────────

  private endParagraph(mark: "cell" | "nestcell" | null): void {
    const state = this.state;
    const style = this.styles.get("s" + state.style) ?? { name: "", flags: {}, outline: null };

    const runs = Structure.mergeRuns(
      this.runs.map((r) => {
        const run: Any = { text: r.text };
        const characterStyle = r.cs !== null ? (this.styles.get("c" + r.cs)?.flags ?? {}) : {};
        for (const [flag, on] of Object.entries(r.flags)) {
          // A flag a style sets is the style's, not the text's; a link's
          // underline is the link's.
          if (on && !style.flags[flag] && !characterStyle[flag] && !(flag === "underline" && r.link !== null)) {
            run[flag] = true;
          }
        }
        if (r.link !== null) run.link = r.link;
        return run;
      }),
    );
    this.runs = [];
    const hasText = phpTrim(Structure.text(runs)) !== "";

    if (state.intbl || mark !== null) {
      this.flushLists();
      this.tableRows ??= [];
      if (hasText) this.cellBlocks.push({ type: "paragraph", runs } as Block);
      if (mark === "cell") {
        this.rowCells.push({ blocks: this.cellBlocks });
        this.cellBlocks = [];
      }
      return;
    }

    this.flushTable();
    if (!hasText) return;

    if (state.ls > 0) {
      this.listEntries.push({ ilvl: state.ilvl, ordered: this.listIsOrdered(state.ls), runs });
      return;
    }
    this.flushLists();

    const outline = state.outline ?? style.outline;
    let level: number | null = null;
    if (outline !== null && outline >= 0 && outline < 9) {
      level = outline + 1;
    } else {
      const m = /^heading[ \t\n\v\f\r]*([1-9])\n?$/i.exec(style.name);
      if (m) level = Number(m[1]);
    }

    this.blocks.push(
      (level !== null ? { type: "heading", level: Math.min(6, level), runs } : { type: "paragraph", runs }) as Block,
    );
  }

  private endRow(): void {
    this.tableRows ??= [];
    if (this.rowCells.length > 0) {
      this.tableRows.push(this.rowHeader ? { header: true, cells: this.rowCells } : { cells: this.rowCells });
    }
    this.rowCells = [];
    this.cellBlocks = [];
  }

  private flushLists(): void {
    if (this.listEntries.length > 0) {
      this.blocks.push(...Structure.lists(this.listEntries));
      this.listEntries = [];
    }
  }

  private flushTable(): void {
    if (this.tableRows === null) return;
    if (this.rowCells.length > 0) this.endRow();
    if (this.tableRows.length > 0) this.blocks.push({ type: "table", rows: this.tableRows } as Block);
    this.tableRows = null;
  }

  private listIsOrdered(ls: number): boolean {
    const levels = this.lists.get(this.overrides.get(ls) ?? -1) ?? [];
    const nfc = levels[0] ?? NFC_BULLET;
    return nfc !== NFC_BULLET && nfc !== NFC_NONE;
  }
}

function isUnderlineOn(word: string): boolean {
  // \ul and its styled variants. \ulc sets the underline COLOUR and \ulnone
  // turns underlining off, so neither is one.
  return word === "ul" || UNDERLINES.has(word);
}

/** Bytes below 0x80 as the ASCII text they are. */
function latin1(bytes: Uint8Array): string {
  let out = "";
  for (let i = 0; i < bytes.length; i += 8192) {
    out += String.fromCharCode(...bytes.subarray(i, i + 8192));
  }
  return out;
}
