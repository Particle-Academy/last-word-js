/**
 * The parts of a Word 97-2003 binary document (MS-DOC) a reader needs. Mirrors
 * PHP `Reader\Doc\WordBinary`.
 *
 * Deliberately not a Word object model. `DocReader` turns this into the same
 * document shape the other readers return; this class only answers "what is at
 * character N".
 *
 * ## What is read
 *
 * - **FIB**: the Word 97+ layout only. Word 6 and 95 files use an older FIB and
 *   are refused by name, as are encrypted files.
 * - **Piece table** (`Clx`): the main document text, including fast-saved files
 *   whose text is split and reordered across pieces, and both piece encodings
 *   (8-bit compressed and UTF-16).
 * - **Style sheet** (`STSH`): each style's built-in identifier, which is how a
 *   heading is recognised independent of the style's localised name.
 * - **Paragraph formatting** (PAPX FKPs): style, list id and level, table
 *   membership and row ends, header rows.
 * - **Character formatting** (CHPX FKPs): bold, italic, underline, strike.
 * - **Lists** (`PlfLfo`, `PlfLst`): whether each list level is numbered or a
 *   bullet.
 *
 * ## What is not
 *
 * Headers, footers, footnotes, comments and text boxes; formatting inherited
 * from a style (only direct formatting is read); images and embedded objects;
 * fonts, sizes and colours.
 *
 * ## Hostile input
 *
 * Every offset read from the file is bounds-checked. A truncated structure is
 * either skipped (formatting, which the text can live without) or refused (the
 * piece table, without which there is no text). Counts are capped by the bytes
 * available, never trusted to size a loop on their own.
 */

import { UnsupportedFormatException } from "../../exceptions";
import { decodeCodePage, fromCodePoint } from "../code-page";
import { CompoundFile, u16, u32 } from "./compound-file";

const NFIB_WORD97 = 0x00c1;
const STI_HEADING_FIRST = 1;
const STI_HEADING_LAST = 9;
const NFC_BULLET = 0x17;
const NFC_NONE = 0xff;

export interface ParagraphProps {
  istd: number;
  ilfo: number;
  ilvl: number;
  inTable: boolean;
  rowEnd: boolean;
  header: boolean;
}

interface Piece {
  cpStart: number;
  cpEnd: number;
  fc: number;
  compressed: boolean;
}

interface FkpRun<T> {
  fcStart: number;
  fcEnd: number;
  props: T;
}

export class WordBinary {
  readonly document: Uint8Array;
  readonly table: Uint8Array;
  private pieces: Piece[] = [];
  private textLengthValue = 0;
  private styleIds: number[] = [];
  private styleNames: string[] = [];
  private paragraphRuns: FkpRun<Partial<ParagraphProps>>[] = [];
  private characterRuns: FkpRun<Record<string, boolean>>[] = [];
  private listIds = new Map<number, number>();
  private listFormats = new Map<number, number[]>();
  private fibFlags = 0;
  private fcLcb: [number, number][] = [];

  constructor(file: CompoundFile) {
    const document = file.stream("WordDocument");
    if (document === null) throw new Error("The compound file has no WordDocument stream.");
    this.document = document;

    this.readFib();

    const tableName = (this.fibFlags & 0x0200) !== 0 ? "1Table" : "0Table";
    const table = file.stream(tableName);
    if (table === null) throw new Error(`The Word document has no ${tableName} stream.`);
    this.table = table;

    this.readPieces();
    this.readStyles();
    this.paragraphRuns = this.readFkps(13, true) as FkpRun<Partial<ParagraphProps>>[];
    this.characterRuns = this.readFkps(12, false) as FkpRun<Record<string, boolean>>[];
    this.readLists();
  }

  /** Number of characters in the main document. */
  textLength(): number {
    return this.textLengthValue;
  }

  /**
   * The main document text, one entry per character position. A surrogate pair
   * occupies two positions: its code point is returned at the first and an
   * empty string at the second. Half a pair on its own is U+FFFD.
   */
  *characters(): Generator<{ char: string; fc: number }> {
    const d = this.document;
    for (const piece of this.pieces) {
      const cpEnd = Math.min(piece.cpEnd, this.textLengthValue);
      let pairedLow = false;
      for (let cp = piece.cpStart; cp < cpEnd; cp++) {
        if (piece.compressed) {
          const fc = piece.fc + (cp - piece.cpStart);
          if (fc >= d.length) return;
          yield { char: decodeCodePage([d[fc]!], 1252), fc };
          continue;
        }

        const fc = piece.fc + 2 * (cp - piece.cpStart);
        if (fc + 2 > d.length) return;
        const unit = u16(d, fc);

        if (pairedLow) {
          pairedLow = false;
          yield { char: "", fc };
          continue;
        }
        if (unit >= 0xd800 && unit <= 0xdbff && cp + 1 < cpEnd) {
          const low = u16(d, fc + 2);
          if (low >= 0xdc00 && low <= 0xdfff) {
            pairedLow = true;
            yield { char: fromCodePoint(0x10000 + ((unit - 0xd800) << 10) + (low - 0xdc00)), fc };
            continue;
          }
        }
        yield { char: fromCodePoint(unit), fc };
      }
    }
  }

  paragraphAt(fc: number): ParagraphProps {
    const run = find(this.paragraphRuns, fc);
    return { istd: 0, ilfo: 0, ilvl: 0, inTable: false, rowEnd: false, header: false, ...(run?.props ?? {}) };
  }

  /** The character formatting flags set at a file position. */
  charactersAt(fc: number): Record<string, boolean> {
    return find(this.characterRuns, fc)?.props ?? {};
  }

  /** The heading level a paragraph style gives, or null when it is not a heading style. */
  headingLevel(istd: number): number | null {
    const sti = this.styleIds[istd];
    if (sti !== undefined && sti >= STI_HEADING_FIRST && sti <= STI_HEADING_LAST) return sti;
    // A user style named "Heading N" (as some converters write) is a heading too.
    const m = /^heading[ \t\n\v\f\r]*([1-9])\n?$/i.exec(phpTrimName(this.styleNames[istd] ?? ""));
    return m ? Number(m[1]) : null;
  }

  /** Whether list `ilfo` at level `ilvl` is numbered (true) or bulleted (false). */
  listIsOrdered(ilfo: number, ilvl: number): boolean {
    const lsid = this.listIds.get(ilfo);
    const formats = lsid !== undefined ? (this.listFormats.get(lsid) ?? []) : [];
    const nfc = formats[ilvl] ?? formats[0] ?? NFC_BULLET;
    return nfc !== NFC_BULLET && nfc !== NFC_NONE;
  }

  // ─── FIB ──────────────────────────────────────────────────────────────

  private readFib(): void {
    const d = this.document;
    if (d.length < 34 || u16(d, 0) !== 0xa5ec) {
      throw new Error("The WordDocument stream does not start with a Word File Information Block.");
    }
    if (u16(d, 2) < NFIB_WORD97) {
      throw new UnsupportedFormatException(
        "doc",
        "This is a Word 6 or Word 95 .doc, which predates the Word 97 binary format last-word reads. " +
          "Re-save it as .docx and read it again.",
      );
    }

    this.fibFlags = u16(d, 10);
    if ((this.fibFlags & 0x0100) !== 0) {
      throw new UnsupportedFormatException(
        "doc",
        "This .doc is password-protected (encrypted), so its text cannot be read. " +
          "Remove the password, save it as .docx and read it again.",
      );
    }

    let offset = 32;
    const csw = u16(d, offset);
    offset += 2 + csw * 2;
    const cslw = u16(d, offset);
    const rgLw = offset + 2;
    offset = rgLw + cslw * 4;
    if (cslw < 4 || offset + 2 > d.length) {
      throw new Error("The Word File Information Block is truncated.");
    }
    this.textLengthValue = u32(d, rgLw + 12); // ccpText

    const count = u16(d, offset);
    offset += 2;
    for (let i = 0; i < count && offset + (i + 1) * 8 <= d.length; i++) {
      this.fcLcb[i] = [u32(d, offset + i * 8), u32(d, offset + i * 8 + 4)];
    }
  }

  /** A structure from the table stream by its FibRgFcLcb97 index, or empty when absent or out of range. */
  private tableStructure(index: number): Uint8Array {
    const [fc, lcb] = this.fcLcb[index] ?? [0, 0];
    if (lcb === 0 || fc + lcb > this.table.length) return new Uint8Array(0);
    return this.table.subarray(fc, fc + lcb);
  }

  // ─── Text ─────────────────────────────────────────────────────────────

  private readPieces(): void {
    const clx = this.tableStructure(33);
    if (clx.length === 0) throw new Error("The Word document has no piece table, so it has no readable text.");

    let offset = 0;
    // Prc records (clxt 0x01) come first and hold property modifiers. Each is at
    // most 0x3FA2 bytes of properties, so every step moves forward.
    while (offset < clx.length && clx[offset] === 0x01) {
      const cbGrpprl = u16(clx, offset + 1);
      if (cbGrpprl > 0x3fa2) throw new Error("The Word piece table is malformed.");
      offset += 3 + cbGrpprl;
    }
    if (offset >= clx.length || clx[offset] !== 0x02) throw new Error("The Word piece table is malformed.");

    const lcb = u32(clx, offset + 1);
    const plc = clx.subarray(offset + 5, offset + 5 + lcb);
    if (plc.length !== lcb || lcb < 4) throw new Error("The Word piece table is truncated.");

    // Pieces cover the text in order, each starting where the last ended. A
    // table that overlaps or goes backwards would let one byte range be read
    // over and over, so reading stops at the first piece out of order.
    const count = Math.floor((lcb - 4) / 12);
    let expected = 0;
    for (let i = 0; i < count && expected < this.textLengthValue; i++) {
      const cpStart = u32(plc, i * 4);
      const cpEnd = u32(plc, (i + 1) * 4);
      const raw = u32(plc, (count + 1) * 4 + i * 8 + 2);
      const compressed = (raw & 0x40000000) !== 0;
      const fc = raw & 0x3fffffff;
      if (cpStart !== expected || cpEnd <= cpStart) break;
      expected = cpEnd;
      this.pieces.push({ cpStart, cpEnd, fc: compressed ? Math.floor(fc / 2) : fc, compressed });
    }

    if (this.pieces.length === 0) throw new Error("The Word piece table holds no text.");
  }

  // ─── Styles ───────────────────────────────────────────────────────────

  private readStyles(): void {
    const stsh = this.tableStructure(1);
    if (stsh.length < 4) return;

    const cbStshi = u16(stsh, 0);
    const count = u16(stsh, 2);
    const cbBase = u16(stsh, 4);
    let offset = 2 + cbStshi;

    for (let istd = 0; istd < count && offset + 2 <= stsh.length; istd++) {
      const cbStd = u16(stsh, offset);
      const std = stsh.subarray(offset + 2, offset + 2 + cbStd);
      offset += 2 + cbStd;
      if (cbStd === 0 || std.length < 2) {
        this.styleIds[istd] = -1;
        this.styleNames[istd] = "";
        continue;
      }

      this.styleIds[istd] = u16(std, 0) & 0x0fff;

      // The name follows the fixed-size base: a character count, then UTF-16.
      const length = u16(std, cbBase);
      this.styleNames[istd] =
        length > 0 && cbBase + 2 + length * 2 <= std.length ? utf16(std.subarray(cbBase + 2, cbBase + 2 + length * 2)) : "";
    }
  }

  // ─── Formatting ───────────────────────────────────────────────────────

  /** Read a PlcBte and every FKP it points at. */
  private readFkps(index: number, paragraphs: boolean): FkpRun<Record<string, unknown>>[] {
    const plc = this.tableStructure(index);
    if (plc.length < 8) return [];

    const count = Math.floor((plc.length - 4) / 8);
    const runs: FkpRun<Record<string, unknown>>[] = [];
    const seenPages = new Set<number>();
    for (let i = 0; i < count; i++) {
      const pn = u32(plc, (count + 1) * 4 + i * 4) & 0x3fffff;
      if (seenPages.has(pn)) continue;
      seenPages.add(pn);

      const page = this.document.subarray(pn * 512, pn * 512 + 512);
      if (page.length !== 512) continue;
      const crun = page[511]!;
      if (crun === 0 || 4 * (crun + 1) > 511) continue;

      for (let j = 0; j < crun; j++) {
        runs.push({
          fcStart: u32(page, j * 4),
          fcEnd: u32(page, (j + 1) * 4),
          props: paragraphs ? papx(page, crun, j) : chpx(page, crun, j),
        });
      }
    }

    // A stable sort, as PHP's usort is since 8.0.
    return runs.sort((a, b) => a.fcStart - b.fcStart);
  }

  // ─── Lists ────────────────────────────────────────────────────────────

  private readLists(): void {
    const lst = this.tableStructure(73);
    const lfo = this.tableStructure(74);
    if (lst.length < 2 || lfo.length < 4) return;

    // PlfLfo: a count, then 16-byte LFO records whose first field is the lsid.
    const lfoCount = Math.min(u32(lfo, 0), Math.floor((lfo.length - 4) / 16));
    for (let i = 0; i < lfoCount; i++) this.listIds.set(i + 1, u32(lfo, 4 + i * 16));

    // PlfLst: a count and 28-byte LSTF records; the LVLs for every list follow
    // the PlfLst immediately in the table stream, in the same order.
    const [fcLst, lcbLst] = this.fcLcb[73]!;
    const lstCount = Math.min(s16(lst, 0), Math.floor((lst.length - 2) / 28));
    let levelsAt = fcLst + lcbLst;
    const t = this.table;

    for (let i = 0; i < lstCount; i++) {
      const record = 2 + i * 28;
      const lsid = u32(lst, record);
      const simple = (lst[record + 26]! & 0x01) !== 0;
      const levels = simple ? 1 : 9;

      const formats: number[] = [];
      for (let level = 0; level < levels; level++) {
        if (levelsAt + 28 > t.length) return;
        formats.push(t[levelsAt + 4]!);
        const cbChpx = t[levelsAt + 24]!;
        const cbPapx = t[levelsAt + 25]!;
        levelsAt += 28 + cbPapx + cbChpx;
        if (levelsAt + 2 > t.length) return;
        levelsAt += 2 + 2 * u16(t, levelsAt);
      }
      this.listFormats.set(lsid, formats);
    }
  }
}

function papx(page: Uint8Array, crun: number, j: number): Record<string, unknown> {
  // rgbx: one 13-byte BxPap (a 1-byte offset and a 12-byte PHE) per run.
  const bxAt = 4 * (crun + 1) + 13 * j;
  if (bxAt >= 511) return {};
  const at = page[bxAt]! * 2;
  if (at === 0 || at >= 511) return {};

  const cb = page[at]!;
  let size: number;
  let start: number;
  if (cb === 0) {
    size = 2 * (page[at + 1] ?? 0);
    start = at + 2;
  } else {
    size = 2 * cb - 1;
    start = at + 1;
  }
  const grpprl = page.subarray(start, start + Math.max(0, Math.min(size, 511 - start)));
  if (grpprl.length < 2) return {};

  const props: Record<string, unknown> = { istd: u16(grpprl, 0) };
  for (const [sprm, operand] of sprms(grpprl.subarray(2))) {
    switch (sprm) {
      case 0x460b: props.ilfo = s16(operand, 0); break; // sprmPIlfo
      case 0x260a: props.ilvl = Math.min(8, operand[0] ?? 0); break; // sprmPIlvl
      case 0x2416: props.inTable = (operand[0] ?? 0) !== 0; break; // sprmPFInTable
      case 0x2417: props.rowEnd = (operand[0] ?? 0) !== 0; break; // sprmPFTtp
      case 0x3404: props.header = (operand[0] ?? 0) !== 0; break; // sprmTTableHeader
    }
  }
  return props;
}

function chpx(page: Uint8Array, crun: number, j: number): Record<string, boolean> {
  const rgbAt = 4 * (crun + 1) + j;
  if (rgbAt >= 511) return {};
  const at = page[rgbAt]! * 2;
  if (at === 0 || at >= 511) return {};
  const cb = page[at]!;
  const grpprl = page.subarray(at + 1, at + 1 + Math.max(0, Math.min(cb, 511 - at - 1)));

  const props: Record<string, boolean> = {};
  for (const [sprm, operand] of sprms(grpprl)) {
    const value = operand[0] ?? 0;
    // Toggle operands: 0 off, 1 on, 0x80 "as the style", 0x81 "opposite of the
    // style". Only direct formatting is read, so 0x81 counts as on.
    const on = value === 1 || value === 0x81;
    switch (sprm) {
      case 0x0835: props.bold = on; break; // sprmCFBold
      case 0x0836: props.italic = on; break; // sprmCFItalic
      case 0x0837: props.strike = on; break; // sprmCFStrike
      case 0x2a3e: props.underline = value !== 0; break; // sprmCKul: any underline kind
    }
  }

  // PHP array_filter: only the flags that are on, in the order they were set.
  const out: Record<string, boolean> = {};
  for (const [key, value] of Object.entries(props)) if (value) out[key] = true;
  return out;
}

/** Walk a grpprl, yielding each property's id and operand bytes. */
function* sprms(grpprl: Uint8Array): Generator<[number, Uint8Array]> {
  let offset = 0;
  const length = grpprl.length;
  while (offset + 2 <= length) {
    const sprm = u16(grpprl, offset);
    offset += 2;

    let size: number;
    switch (sprm >> 13) {
      case 0: case 1: size = 1; break;
      case 2: case 4: case 5: size = 2; break;
      case 3: size = 4; break;
      case 7: size = 3; break;
      default: size = variableOperandSize(sprm, grpprl, offset);
    }
    if (size < 0 || offset + size > length) return;

    yield [sprm, grpprl.subarray(offset, offset + size)];
    offset += size;
  }
}

function variableOperandSize(sprm: number, grpprl: Uint8Array, offset: number): number {
  if (offset >= grpprl.length) return -1;
  // sprmTDefTable and sprmTDefTable10: a 2-byte count of the rest, plus one.
  if (sprm === 0xd608 || sprm === 0xd606) {
    return offset + 2 <= grpprl.length ? 2 + u16(grpprl, offset) - 1 : -1;
  }
  // sprmPChgTabs with the 255 escape: deleted tabs (4 bytes each), then added
  // tabs (3 bytes each), each list prefixed by its count.
  if (sprm === 0xc615 && grpprl[offset] === 255) {
    const deleted = grpprl[offset + 1] ?? 0;
    const added = grpprl[offset + 2 + 4 * deleted] ?? 0;
    return 2 + 4 * deleted + 1 + 3 * added;
  }
  return 1 + grpprl[offset]!;
}

/** The run containing a file position, by binary search. */
function find<T>(runs: FkpRun<T>[], fc: number): FkpRun<T> | null {
  let low = 0;
  let high = runs.length - 1;
  while (low <= high) {
    const mid = (low + high) >> 1;
    const run = runs[mid]!;
    if (fc < run.fcStart) high = mid - 1;
    else if (fc >= run.fcEnd) low = mid + 1;
    else return run;
  }
  return null;
}

function utf16(raw: Uint8Array): string {
  let out = "";
  for (let i = 0; i + 1 < raw.length; i += 2) out += fromCodePoint(u16(raw, i));
  return out;
}

function phpTrimName(s: string): string {
  return s.replace(/^[ \t\n\r\0\x0B]+|[ \t\n\r\0\x0B]+$/g, "");
}

function s16(b: Uint8Array, o: number): number {
  const v = u16(b, o);
  return v >= 0x8000 ? v - 0x10000 : v;
}
