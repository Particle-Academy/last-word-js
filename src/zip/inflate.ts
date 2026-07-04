/**
 * Raw DEFLATE (RFC 1951) decompression — `inflateRaw`.
 *
 * Office documents (.xlsx/.pptx) authored by Excel/PowerPoint store their parts
 * with DEFLATE compression, so the reader must inflate them. Our own writer uses
 * the STORE method (no compression), so this is only exercised when reading
 * externally-authored files. Decompress-only, zero-dependency, isomorphic.
 *
 * Algorithm adapted from the public-domain "tinf" inflate (Jørgen Ibsen).
 */

class Tree {
  // number of codes with a given bit length
  readonly table = new Uint16Array(16);
  // symbols sorted by code
  readonly trans = new Uint16Array(288);
}

const LENGTH_BITS = new Uint8Array([
  0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5,
  5, 5, 0,
]);
const LENGTH_BASE = new Uint16Array([
  3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67,
  83, 99, 115, 131, 163, 195, 227, 258,
]);
const DIST_BITS = new Uint8Array([
  0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11,
  11, 12, 12, 13, 13,
]);
const DIST_BASE = new Uint16Array([
  1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769,
  1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577,
]);
// order in which code-length-code lengths are stored
const CLC_INDEX = new Uint8Array([
  16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15,
]);

function buildFixedTrees(lt: Tree, dt: Tree): void {
  let i: number;
  for (i = 0; i < 7; i++) lt.table[i] = 0;
  lt.table[7] = 24;
  lt.table[8] = 152;
  lt.table[9] = 112;
  for (i = 0; i < 24; i++) lt.trans[i] = 256 + i;
  for (i = 0; i < 144; i++) lt.trans[24 + i] = i;
  for (i = 0; i < 8; i++) lt.trans[24 + 144 + i] = 280 + i;
  for (i = 0; i < 112; i++) lt.trans[24 + 144 + 8 + i] = 144 + i;
  for (i = 0; i < 5; i++) dt.table[i] = 0;
  dt.table[5] = 32;
  for (i = 0; i < 32; i++) dt.trans[i] = i;
}

const OFFS = new Uint16Array(16);

function buildTree(t: Tree, lengths: Uint8Array, off: number, num: number): void {
  let i: number;
  let sum = 0;
  for (i = 0; i < 16; i++) t.table[i] = 0;
  for (i = 0; i < num; i++) t.table[lengths[off + i]!]!++;
  t.table[0] = 0;
  for (i = 0; i < 16; i++) {
    OFFS[i] = sum;
    sum += t.table[i]!;
  }
  for (i = 0; i < num; i++) {
    const len = lengths[off + i]!;
    if (len) t.trans[OFFS[len]!++] = i;
  }
}

class Reader {
  source: Uint8Array;
  index = 0;
  tag = 0;
  bitcount = 0;
  constructor(source: Uint8Array) {
    this.source = source;
  }
}

function getBit(d: Reader): number {
  if (d.bitcount-- === 0) {
    d.tag = d.source[d.index++]!;
    d.bitcount = 7;
  }
  const bit = d.tag & 1;
  d.tag >>>= 1;
  return bit;
}

function readBits(d: Reader, num: number, base: number): number {
  if (!num) return base;
  let val = 0;
  for (let i = 0; i < num; i++) val |= getBit(d) << i;
  return val + base;
}

function decodeSymbol(d: Reader, t: Tree): number {
  let sum = 0;
  let cur = 0;
  let len = 0;
  do {
    cur = 2 * cur + getBit(d);
    len++;
    sum += t.table[len]!;
    cur -= t.table[len]!;
  } while (cur >= 0);
  return t.trans[sum + cur]!;
}

function decodeTrees(d: Reader, lt: Tree, dt: Tree): void {
  const lengths = new Uint8Array(288 + 32);
  const hlit = readBits(d, 5, 257);
  const hdist = readBits(d, 5, 1);
  const hclen = readBits(d, 4, 4);
  let i: number;
  for (i = 0; i < 19; i++) lengths[i] = 0;
  for (i = 0; i < hclen; i++) {
    const clen = readBits(d, 3, 0);
    lengths[CLC_INDEX[i]!] = clen;
  }
  const codeTree = new Tree();
  buildTree(codeTree, lengths, 0, 19);
  for (let num = 0; num < hlit + hdist; ) {
    const sym = decodeSymbol(d, codeTree);
    switch (sym) {
      case 16: {
        const prev = lengths[num - 1]!;
        for (let length = readBits(d, 2, 3); length; length--) lengths[num++] = prev;
        break;
      }
      case 17:
        for (let length = readBits(d, 3, 3); length; length--) lengths[num++] = 0;
        break;
      case 18:
        for (let length = readBits(d, 7, 11); length; length--) lengths[num++] = 0;
        break;
      default:
        lengths[num++] = sym;
        break;
    }
  }
  buildTree(lt, lengths, 0, hlit);
  buildTree(dt, lengths, hlit, hdist);
}

class Out {
  buf = new Uint8Array(1024);
  len = 0;
  push(b: number): void {
    if (this.len >= this.buf.length) {
      const next = new Uint8Array(this.buf.length * 2);
      next.set(this.buf);
      this.buf = next;
    }
    this.buf[this.len++] = b;
  }
}

function inflateBlockData(d: Reader, out: Out, lt: Tree, dt: Tree): void {
  for (;;) {
    const sym = decodeSymbol(d, lt);
    if (sym === 256) return;
    if (sym < 256) {
      out.push(sym);
    } else {
      const lengthSym = sym - 257;
      const length = readBits(d, LENGTH_BITS[lengthSym]!, LENGTH_BASE[lengthSym]!);
      const distSym = decodeSymbol(d, dt);
      const dist = readBits(d, DIST_BITS[distSym]!, DIST_BASE[distSym]!);
      const offs = out.len - dist;
      for (let i = 0; i < length; i++) out.push(out.buf[offs + i]!);
    }
  }
}

function inflateUncompressedBlock(d: Reader, out: Out): void {
  // discard remaining bits in the current (partial) byte
  d.bitcount = 0;
  let length = d.source[d.index + 1]! * 256 + d.source[d.index]!;
  d.index += 4; // skip LEN + NLEN
  for (; length; length--) out.push(d.source[d.index++]!);
}

export function inflateRaw(source: Uint8Array): Uint8Array {
  const d = new Reader(source);
  const out = new Out();
  const lt = new Tree();
  const dt = new Tree();
  let bfinal: number;
  do {
    bfinal = getBit(d);
    const btype = readBits(d, 2, 0);
    if (btype === 0) {
      inflateUncompressedBlock(d, out);
    } else if (btype === 1) {
      buildFixedTrees(lt, dt);
      inflateBlockData(d, out, lt, dt);
    } else if (btype === 2) {
      decodeTrees(d, lt, dt);
      inflateBlockData(d, out, lt, dt);
    } else {
      throw new Error("dark-slide: invalid DEFLATE block type");
    }
  } while (!bfinal);
  return out.buf.subarray(0, out.len);
}
