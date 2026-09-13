/**
 * A read-only Compound File Binary (MS-CFB) container: the "OLE2" file a Word
 * 97-2003 `.doc` is stored in, along with `.xls`, `.ppt`, `.msg` and others.
 * Mirrors PHP `Reader\Doc\CompoundFile`.
 *
 * It exposes the streams that sit directly under the root storage, which is all
 * a `.doc` needs (`WordDocument`, `0Table` / `1Table`). Nested storages such as
 * `ObjectPool` hold embedded objects, and an embedded Word document carries its
 * own `WordDocument` stream there; walking only the root's children is what
 * keeps that one from being mistaken for the document itself.
 *
 * ## Hostile input
 *
 * - every read is bounds-checked against the file, and a short read is an error,
 *   never a silently truncated stream;
 * - a sector chain is followed at most once per sector, so a FAT loop fails
 *   instead of spinning forever;
 * - a stream's declared size is capped by what its chain can actually hold, and
 *   a stream over 256 MB is refused outright;
 * - the allocation table may not list more sectors than the file holds;
 * - the directory tree walk tracks visited entries.
 *
 * Nothing here is recovered from: a damaged container is refused with a message
 * naming what was wrong.
 */

export const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

const FREESECT = 0xffffffff;
const ENDOFCHAIN = 0xfffffffe;
const NOSTREAM = 0xffffffff;
const MAX_STREAM_BYTES = 256 * 1024 * 1024;

export class CompoundFile {
  private fat: number[] = [];
  private miniFat: number[] = [];
  private sectorSize = 512;
  private miniSectorSize = 64;
  private miniStreamCutoff = 4096;
  private miniStream: Uint8Array = new Uint8Array(0);
  private streams = new Map<string, { start: number; size: number }>();

  private constructor(private readonly bytes: Uint8Array) {}

  static fromBytes(bytes: Uint8Array): CompoundFile {
    const file = new CompoundFile(bytes);
    file.parse();
    return file;
  }

  /** The names of the streams directly under the root storage. */
  streamNames(): string[] {
    return [...this.streams.keys()];
  }

  hasStream(name: string): boolean {
    return this.streams.has(name);
  }

  /** The whole stream, or null when the root storage has no stream by that name. */
  stream(name: string): Uint8Array | null {
    const entry = this.streams.get(name);
    if (entry === undefined) return null;
    return entry.size < this.miniStreamCutoff
      ? this.readChain(this.miniStream, this.miniFat, entry.start, this.miniSectorSize, entry.size, "mini stream", 0)
      : this.readChain(this.bytes, this.fat, entry.start, this.sectorSize, entry.size, `stream ${name}`, this.sectorSize);
  }

  private parse(): void {
    const b = this.bytes;
    if (b.length < 512 || !SIGNATURE.every((v, i) => b[i] === v)) {
      throw new Error("Not a Compound File Binary document: the header is missing or truncated.");
    }
    if (u16(b, 28) !== 0xfffe) {
      throw new Error("Compound File Binary header has an invalid byte-order mark.");
    }

    const major = u16(b, 26);
    const sectorShift = u16(b, 30);
    const miniShift = u16(b, 32);
    if (!((major === 3 && sectorShift === 9) || (major === 4 && sectorShift === 12)) || miniShift !== 6) {
      throw new Error("Compound File Binary header declares an unsupported sector size.");
    }
    this.sectorSize = 1 << sectorShift;
    this.miniSectorSize = 1 << miniShift;
    this.miniStreamCutoff = u32(b, 56);
    if (this.miniStreamCutoff !== 4096) {
      throw new Error("Compound File Binary header declares an invalid mini stream cutoff.");
    }

    this.fat = this.readFat();

    const directory = this.readChain(b, this.fat, u32(b, 48), this.sectorSize, null, "directory", this.sectorSize);
    const entries: Uint8Array[] = [];
    for (let at = 0; at < directory.length; at += 128) entries.push(directory.subarray(at, at + 128));
    if (entries.length === 0 || entries[0]!.length < 128) {
      throw new Error("Compound File Binary directory is empty.");
    }

    const root = entries[0]!;
    if (root[66] !== 5) {
      throw new Error("Compound File Binary directory does not start with the root storage.");
    }

    const miniFatStart = u32(b, 60);
    this.miniFat =
      miniFatStart === ENDOFCHAIN || miniFatStart === FREESECT
        ? []
        : u32s(this.readChain(b, this.fat, miniFatStart, this.sectorSize, null, "mini FAT", this.sectorSize));

    const rootStart = u32(root, 116);
    const rootSize = streamSize(root, major);
    this.miniStream =
      rootSize === 0 || rootStart === ENDOFCHAIN
        ? new Uint8Array(0)
        : this.readChain(b, this.fat, rootStart, this.sectorSize, rootSize, "mini stream container", this.sectorSize);

    // The root's children are a red-black tree threaded through left/right
    // sibling ids. Its shape does not matter to a reader; visiting every node
    // does, and a crafted file can make the ids a cycle.
    const pending = [u32(root, 76)];
    const visited = new Set<number>();
    while (pending.length > 0) {
      const id = pending.pop()!;
      if (id === NOSTREAM || visited.has(id)) continue;
      const entry = entries[id];
      if (entry === undefined || entry.length < 128) {
        throw new Error("Compound File Binary directory points at an entry that does not exist.");
      }
      visited.add(id);
      pending.push(u32(entry, 68), u32(entry, 72));

      if (entry[66] !== 2) continue; // a storage (or unused slot); its contents are not the document
      const nameLength = Math.min(64, u16(entry, 64));
      const name = utf16(entry.subarray(0, Math.max(0, nameLength - 2)));
      this.streams.set(name, { start: u32(entry, 116), size: streamSize(entry, major) });
    }
  }

  private readFat(): number[] {
    const b = this.bytes;
    const fatSectors: number[] = [];
    for (let i = 0; i < 109; i++) {
      const sector = u32(b, 76 + i * 4);
      if (sector !== FREESECT) fatSectors.push(sector);
    }

    let difat = u32(b, 68);
    const perDifat = this.sectorSize / 4 - 1;
    const seen = new Set<number>();
    while (difat !== ENDOFCHAIN && difat !== FREESECT) {
      if (seen.has(difat)) throw new Error("Compound File Binary DIFAT chain loops.");
      seen.add(difat);
      const sector = this.sector(b, difat, this.sectorSize, this.sectorSize, "DIFAT");
      for (let i = 0; i < perDifat; i++) {
        const entry = u32(sector, i * 4);
        if (entry !== FREESECT) fatSectors.push(entry);
      }
      difat = u32(sector, perDifat * 4);
    }

    if (fatSectors.length > Math.floor(b.length / this.sectorSize) + 1) {
      throw new Error("Compound File Binary allocation table lists more sectors than the file holds.");
    }

    const fat: number[] = [];
    for (const sector of fatSectors) {
      for (const next of u32s(this.sector(b, sector, this.sectorSize, this.sectorSize, "FAT"))) fat.push(next);
    }
    return fat;
  }

  /** Follow a chain and concatenate its sectors. `size` null means the whole chain. */
  private readChain(
    source: Uint8Array,
    table: number[],
    start: number,
    unit: number,
    size: number | null,
    what: string,
    headerOffset: number,
  ): Uint8Array {
    if (size === 0) return new Uint8Array(0);

    const parts: Uint8Array[] = [];
    let length = 0;
    const seen = new Set<number>();
    let sector = start;
    const limit = Math.floor(source.length / unit) + 1;

    while (sector !== ENDOFCHAIN) {
      if (sector === FREESECT || sector >= 0xfffffffa) {
        throw new Error(`Compound File Binary ${what} chain is broken.`);
      }
      if (seen.has(sector) || seen.size > limit) {
        throw new Error(`Compound File Binary ${what} chain loops.`);
      }
      seen.add(sector);
      const data = this.sector(source, sector, unit, headerOffset, what);
      parts.push(data);
      length += data.length;
      if (size !== null && length >= size) return concat(parts, size);
      const next = table[sector];
      if (next === undefined) {
        throw new Error(`Compound File Binary ${what} chain runs past the allocation table.`);
      }
      sector = next;
    }

    if (size !== null && length < size) {
      throw new Error(`Compound File Binary ${what} is shorter than its declared size.`);
    }
    return concat(parts, length);
  }

  private sector(source: Uint8Array, index: number, unit: number, headerOffset: number, what: string): Uint8Array {
    const offset = headerOffset + index * unit;
    if (offset < 0 || offset + unit > source.length) {
      throw new Error(`Compound File Binary ${what} points outside the file.`);
    }
    return source.subarray(offset, offset + unit);
  }
}

function streamSize(entry: Uint8Array, major: number): number {
  const low = u32(entry, 120);
  // Version 3 files may leave garbage in the high half; only version 4 can
  // hold a stream over 4GB, and nothing here reads one that large.
  const high = major === 4 ? u32(entry, 124) : 0;
  if (high !== 0 || low > MAX_STREAM_BYTES) {
    throw new Error("Compound File Binary stream is too large to read.");
  }
  return low;
}

function concat(parts: Uint8Array[], size: number): Uint8Array {
  const out = new Uint8Array(size);
  let at = 0;
  for (const part of parts) {
    if (at >= size) break;
    const take = part.subarray(0, Math.min(part.length, size - at));
    out.set(take, at);
    at += take.length;
  }
  return out;
}

function utf16(raw: Uint8Array): string {
  let out = "";
  for (let i = 0; i + 1 < raw.length; i += 2) {
    const unit = raw[i]! | (raw[i + 1]! << 8);
    out += unit >= 0xd800 && unit <= 0xdfff ? String.fromCharCode(0xfffd) : String.fromCharCode(unit);
  }
  return out;
}

function u32s(bytes: Uint8Array): number[] {
  const out: number[] = [];
  for (let i = 0; i + 3 < bytes.length; i += 4) out.push(u32(bytes, i));
  return out;
}

export function u16(b: Uint8Array, o: number): number {
  return o >= 0 && o + 2 <= b.length ? b[o]! | (b[o + 1]! << 8) : 0;
}

export function u32(b: Uint8Array, o: number): number {
  return o >= 0 && o + 4 <= b.length ? (b[o]! | (b[o + 1]! << 8) | (b[o + 2]! << 16) | (b[o + 3]! << 24)) >>> 0 : 0;
}
