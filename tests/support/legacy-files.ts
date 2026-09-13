/**
 * Hand-built legacy files for the reader's guard tests. Mirrors PHP
 * `tests/Support/LegacyFiles.php`, byte for byte.
 *
 * The converted fixtures in `test/fixtures/formats/` prove the readers get a
 * real document right. They cannot prove the readers survive a damaged or
 * hostile one, because no converter writes those. These builders do, so each
 * test can break exactly one structure and nothing else.
 */

import { zipSync } from "../../src";

export const ENDOFCHAIN = 0xfffffffe;
export const FREESECT = 0xffffffff;
export const FATSECT = 0xfffffffd;
export const NOSTREAM = 0xffffffff;

const encoder = new TextEncoder();

/** File offset of sector N in a version-3 compound file. */
export function sectorOffset(sector: number): number {
  return 512 + sector * 512;
}

/** File offset of directory entry N in a file built by `cfb()` (the directory is sector 1). */
export function entryOffset(entry: number): number {
  return sectorOffset(1) + entry * 128;
}

class Bytes {
  private parts: number[] = [];
  u8(v: number): this { this.parts.push(v & 0xff); return this; }
  u16(v: number): this { return this.u8(v).u8(v >>> 8); }
  u32(v: number): this { return this.u16(v & 0xffff).u16(v >>> 16); }
  raw(b: ArrayLike<number>): this { for (let i = 0; i < b.length; i++) this.parts.push(b[i]!); return this; }
  zeros(n: number): this { for (let i = 0; i < n; i++) this.parts.push(0); return this; }
  pad(to: number): this { while (this.parts.length < to) this.parts.push(0); return this; }
  get length(): number { return this.parts.length; }
  done(): Uint8Array { return Uint8Array.from(this.parts); }
}

export function ascii(s: string): Uint8Array {
  return Uint8Array.from(s, (c) => c.charCodeAt(0));
}

/** Overwrite bytes at an offset, returning a copy. */
export function patch(bytes: Uint8Array, offset: number, replacement: Uint8Array): Uint8Array {
  const out = bytes.slice();
  out.set(replacement, offset);
  return out;
}

export function le32(...values: number[]): Uint8Array {
  const b = new Bytes();
  for (const v of values) b.u32(v);
  return b.done();
}

function entry(name: string, type: number, left: number, right: number, child: number, start: number, size: number): Uint8Array {
  const b = new Bytes();
  for (const c of name) b.u16(c.charCodeAt(0)); // ASCII names only, which is all the tests use
  b.pad(64);
  return b
    .u16(name === "" ? 0 : name.length * 2 + 2)
    .u8(type).u8(1)
    .u32(left).u32(right).u32(child)
    .zeros(16 + 4 + 16)
    .u32(start).u32(size).u32(0)
    .done();
}

/**
 * A version-3 compound file with up to three streams under the root. Sector 0
 * is the FAT, sector 1 the directory, and the streams follow, each padded to at
 * least 4096 bytes so it lives in regular sectors rather than the mini stream.
 */
export function cfb(streams: Record<string, Uint8Array>): Uint8Array {
  const names = Object.keys(streams);
  const fat = [FATSECT, ENDOFCHAIN];
  const data = new Bytes();
  const entries = [entry("Root Entry", 5, NOSTREAM, NOSTREAM, names.length === 0 ? NOSTREAM : 1, ENDOFCHAIN, 0)];

  names.forEach((name, i) => {
    const content = streams[name]!;
    const size = Math.max(4096, content.length);
    const sectors = Math.ceil(size / 512);
    const start = fat.length;
    for (let s = 0; s < sectors; s++) fat.push(s === sectors - 1 ? ENDOFCHAIN : start + s + 1);
    const at = data.length;
    data.raw(content).pad(at + sectors * 512);
    const right = i + 1 < names.length ? i + 2 : NOSTREAM;
    entries.push(entry(name, 2, NOSTREAM, right, NOSTREAM, start, size));
  });

  const directory = new Bytes();
  for (const e of entries) directory.raw(e);
  while (directory.length < 512) directory.raw(entry("", 0, NOSTREAM, NOSTREAM, NOSTREAM, 0, 0));

  const fatSector = new Bytes();
  for (let i = 0; i < 128; i++) fatSector.u32(fat[i] ?? FREESECT);

  const header = new Bytes()
    .raw([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]).zeros(16)
    .u16(0x003e).u16(3).u16(0xfffe).u16(9).u16(6)
    .zeros(6)
    .u32(0) // directory sectors (always 0 in version 3)
    .u32(1) // FAT sectors
    .u32(1) // first directory sector
    .u32(0) // transaction signature
    .u32(4096) // mini stream cutoff
    .u32(ENDOFCHAIN) // first mini FAT sector
    .u32(0) // mini FAT sectors
    .u32(ENDOFCHAIN) // first DIFAT sector
    .u32(0) // DIFAT sectors
    .u32(0); // DIFAT[0]: the FAT is sector 0
  for (let i = 0; i < 108; i++) header.u32(FREESECT);

  return new Bytes().raw(header.done()).raw(fatSector.done()).raw(directory.done().subarray(0, 512)).raw(data.done()).done();
}

export interface WordOptions {
  nFib?: number;
  flags?: number;
  ccpText?: number;
  clx?: Uint8Array;
  unicode?: boolean;
}

/**
 * A Word 97 binary document: the `WordDocument` and `0Table` streams. The FIB
 * carries only what the reader looks at; the text sits at offset 1024 of the
 * WordDocument stream, in one piece unless `clx` says otherwise.
 */
export function word(text: Uint8Array, options: WordOptions = {}): Record<string, Uint8Array> {
  const unicode = options.unicode ?? false;
  const characters = unicode ? Math.floor(text.length / 2) : text.length;
  const clx = options.clx ?? piecesClx([[0, characters, 1024, !unicode]]);

  const fib = new Bytes()
    .u16(0xa5ec).u16(options.nFib ?? 0x00c1).zeros(6).u16(options.flags ?? 0).zeros(20)
    .u16(14).zeros(28) // csw, fibRgW
    .u16(22); // cslw
  const rgLw = new Bytes().zeros(12).u32(options.ccpText ?? characters).zeros(88 - 16);
  fib.raw(rgLw.done());
  const rgFcLcb = new Bytes();
  for (let i = 0; i < 93; i++) {
    if (i === 33) rgFcLcb.u32(0).u32(clx.length);
    else rgFcLcb.u32(0).u32(0);
  }
  fib.u16(93).raw(rgFcLcb.done());

  const document = new Bytes().raw(fib.done()).pad(1024).raw(text).done();
  return { WordDocument: document, "0Table": clx };
}

/** A Clx holding one piece table: [cpStart, cpEnd, byte offset, compressed] per piece. */
export function piecesClx(pieces: [number, number, number, boolean][], prc: Uint8Array = new Uint8Array(0)): Uint8Array {
  const cps = new Bytes();
  const pcds = new Bytes();
  for (const [cpStart, , offset, compressed] of pieces) {
    cps.u32(cpStart);
    pcds.u16(0).u32(compressed ? ((offset * 2) | 0x40000000) >>> 0 : offset).u16(0);
  }
  cps.u32(pieces.length === 0 ? 0 : pieces[pieces.length - 1]![1]);
  const plc = new Bytes().raw(cps.done()).raw(pcds.done()).done();
  return new Bytes().raw(prc).u8(0x02).u32(plc.length).raw(plc).done();
}

/** An ODT whose content.xml is the given body markup, or the given whole part. */
export function odt(body: string, contentXml?: string): Uint8Array {
  const xml =
    contentXml ??
    '<?xml version="1.0" encoding="UTF-8"?>' +
      "<office:document-content" +
      ' xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"' +
      ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"' +
      ' xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">' +
      "<office:body><office:text>" + body + "</office:text></office:body>" +
      "</office:document-content>";
  return zipSync([
    { name: "mimetype", data: encoder.encode("application/vnd.oasis.opendocument.text") },
    { name: "content.xml", data: encoder.encode(xml) },
  ]);
}

/** A zip holding the given entries, for zips that are not documents. */
export function zip(entries: Record<string, string>): Uint8Array {
  return zipSync(Object.entries(entries).map(([name, content]) => ({ name, data: encoder.encode(content) })));
}
