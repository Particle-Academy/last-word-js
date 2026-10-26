import { crc32 } from "./crc32";
import { inflateRaw } from "./inflate";
import { ArchiveRefused } from "../exceptions";
export interface ArchiveLimits {
  maxArchiveBytes?: number;
  maxEntries?: number;
  maxPartBytes?: number;
  maxExpandedBytes?: number;
  maxCompressionRatio?: number;
}
export const DEFAULT_ARCHIVE_LIMITS: Readonly<Required<ArchiveLimits>> = Object.freeze({
  maxArchiveBytes: 67108864, maxEntries: 16384, maxPartBytes: 67108864,
  maxExpandedBytes: 268435456, maxCompressionRatio: 1000,
});
type Part = { name: string; method: number; crc: number; size: number; compressed: number; start: number };
const corrupt = (): never => { throw new ArchiveRefused("archive_corrupt"); };

/** @internal Bounded, immutable reader source. No extraction or network. */
export class SourceArchive {
  private constructor(private readonly original: Uint8Array, private readonly parts: Map<string, Uint8Array>) {}

  static open(input: Uint8Array, options: ArchiveLimits = {}): SourceArchive {
    for (const [key, value] of Object.entries(options)) {
      if (!Object.prototype.hasOwnProperty.call(DEFAULT_ARCHIVE_LIMITS, key) || !Number.isSafeInteger(value) || value < 1) throw new ArchiveRefused("archive_limits");
    }
    const limits = { ...DEFAULT_ARCHIVE_LIMITS, ...options };
    if (input.length > limits.maxArchiveBytes) throw new ArchiveRefused("archive_limit");
    const bytes = input.slice();
    const records = index(bytes, limits);
    const parts = new Map<string, Uint8Array>();
    for (const record of records) {
      const compressed = bytes.subarray(record.start, record.start + record.compressed);
      let payload: Uint8Array;
      try {
        // A forged declared size cannot induce an unbounded inflater allocation.
        payload = record.method === 0 ? compressed.slice() : inflateRaw(compressed, record.size + 1);
      } catch { throw new ArchiveRefused("archive_corrupt"); }
      if (payload.length !== record.size || crc32(payload) !== record.crc) corrupt();
      parts.set(record.name, payload);
    }
    return new SourceArchive(bytes, parts);
  }

  names(): string[] { return [...this.parts.keys()]; }
  read(name: string): Uint8Array | null { return this.parts.get(name)?.slice() ?? null; }
  toBytes(): Uint8Array { return this.original.slice(); }
}

function index(bytes: Uint8Array, limits: Required<ArchiveLimits>): Part[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const bounds = (at: number, size: number, end = bytes.length): void => {
    if (at < 0 || size < 0 || at + size > end || at + size > bytes.length) corrupt();
  };
  const u16 = (at: number): number => { bounds(at, 2); return view.getUint16(at, true); };
  const u32 = (at: number): number => { bounds(at, 4); return view.getUint32(at, true); };
  let end = -1;
  for (let at = bytes.length - 22; at >= Math.max(0, bytes.length - 65557); at--) {
    if (u32(at) === 0x06054b50 && at + 22 + u16(at + 20) === bytes.length && u32(at + 16) + u32(at + 12) === at) {
      end = at; break;
    }
  }
  if (end === -1) corrupt();
  const count = u16(end + 10), directory = u32(end + 16);
  if (u16(end + 4) !== 0 || u16(end + 6) !== 0 || u16(end + 8) !== count || count === 65535) throw new ArchiveRefused("archive_unsupported");
  if (count > limits.maxEntries) throw new ArchiveRefused("archive_limit");
  if (end - directory < count * 46) corrupt();
  const records: Part[] = [], seen = new Set<string>(), regions: [number, number][] = [];
  let at = directory, expanded = 0;
  const extra = (start: number, length: number): void => {
    const finish = start + length;
    while (start < finish) {
      bounds(start, 4, finish);
      const id = u16(start), size = u16(start + 2);
      bounds(start + 4, size, finish);
      if (id === 1 || id === 0x9901) throw new ArchiveRefused("archive_unsupported");
      start += 4 + size;
    }
  };
  for (let i = 0; i < count; i++) {
    bounds(at, 46, end);
    if (u32(at) !== 0x02014b50) corrupt();
    const flags = u16(at + 8), method = u16(at + 10), crc = u32(at + 16);
    const compressed = u32(at + 20), size = u32(at + 24), local = u32(at + 42);
    const nameLength = u16(at + 28), extraLength = u16(at + 30), commentLength = u16(at + 32);
    if ((flags & ~0x080e) !== 0 || ![0, 8].includes(method) || (method === 0 && (flags & 6) !== 0)
      || u16(at + 34) !== 0 || [compressed, size, local].includes(0xffffffff) || u16(at + 6) > 20) throw new ArchiveRefused("archive_unsupported");
    bounds(at + 46, nameLength + extraLength + commentLength, end);
    const name = partName(bytes.subarray(at + 46, at + 46 + nameLength), flags);
    if (seen.has(name)) throw new ArchiveRefused("archive_duplicate");
    seen.add(name); expanded += size;
    if (size > limits.maxPartBytes || expanded > limits.maxExpandedBytes || size > Math.max(1, compressed) * limits.maxCompressionRatio) throw new ArchiveRefused("archive_limit");
    if (method === 0 && compressed !== size) corrupt();
    extra(at + 46 + nameLength, extraLength);
    bounds(local, 30, directory);
    if (u32(local) !== 0x04034b50 || u16(local + 4) !== u16(at + 6) || u16(local + 6) !== flags || u16(local + 8) !== method) corrupt();
    const localName = u16(local + 26), localExtra = u16(local + 28);
    const start = local + 30 + localName + localExtra;
    bounds(local + 30, localName + localExtra + compressed, directory);
    const rawName = bytes.subarray(local + 30, local + 30 + localName);
    const centralName = bytes.subarray(at + 46, at + 46 + nameLength);
    if (rawName.length !== centralName.length || !rawName.every((byte, pos) => byte === centralName[pos])) corrupt();
    extra(local + 30 + localName, localExtra);
    for (const [offset, expected] of [[14, crc], [18, compressed], [22, size]]) {
      const value = u32(local + offset!);
      if (value !== expected && !((flags & 8) !== 0 && value === 0)) corrupt();
    }
    let finish = start + compressed;
    if ((flags & 8) !== 0) {
      let found = false;
      for (const prefix of [0, 4]) {
        if (finish + prefix + 12 <= directory && (prefix === 0 || u32(finish) === 0x08074b50)
          && u32(finish + prefix) === crc && u32(finish + prefix + 4) === compressed && u32(finish + prefix + 8) === size) {
          finish += prefix + 12; found = true; break;
        }
      }
      if (!found) corrupt();
    }
    records.push({ name, method, crc, size, compressed, start }); regions.push([local, finish]);
    at += 46 + nameLength + extraLength + commentLength;
  }
  if (at !== end) corrupt();
  regions.sort((a, b) => a[0] - b[0]);
  let previous = 0;
  for (const [start, finish] of regions) { if (start < previous) corrupt(); previous = finish; }
  return records;
}

function partName(bytes: Uint8Array, flags: number): string {
  let name: string;
  try { name = new TextDecoder("utf-8", { fatal: true, ignoreBOM: true }).decode(bytes); }
  catch { throw new ArchiveRefused("archive_path"); }
  if (name === "" || name.startsWith("/") || name.includes(String.fromCharCode(92)) || name.includes(":")) throw new ArchiveRefused("archive_path");
  for (const byte of bytes) {
    if (byte < 32 || byte === 127) throw new ArchiveRefused("archive_path");
    if (byte > 127 && (flags & 0x0800) === 0) throw new ArchiveRefused("archive_unsupported");
  }
  const parts = (name.endsWith("/") ? name.slice(0, -1) : name).split("/");
  if (parts.some(part => part === "" || part === "." || part === "..")) throw new ArchiveRefused("archive_path");
  return name;
}
