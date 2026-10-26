import { describe, expect, it } from "vitest";
import { deflateRawSync } from "node:zlib";
import { crc32 } from "../src/zip/crc32";
import { inflateRaw } from "../src/zip/inflate";
import { SourceArchive } from "../src/zip/source-archive";
import { ArchiveRefused } from "../src/exceptions";
const enc = new TextEncoder();
type Entry = { name: string; data: Uint8Array; method?: number; flags?: number; size?: number; crc?: number; descriptor?: boolean };
const text = (name: string, value: string, options: Partial<Entry> = {}): Entry => ({ name, data: enc.encode(value), ...options });
function fixture(entries: Entry[], comment = new Uint8Array()): Uint8Array {
  const locals: Uint8Array[] = [], central: Uint8Array[] = [];
  let offset = 0;
  for (const entry of entries) {
    const name = enc.encode(entry.name), method = entry.method ?? 0;
    const flags = entry.flags ?? (entry.descriptor ? 0x0808 : 0x0800);
    const data = method === 8 ? new Uint8Array(deflateRawSync(entry.data)) : entry.data;
    const crc = entry.crc ?? crc32(entry.data), size = entry.size ?? entry.data.length;
    const local = new Uint8Array(30 + name.length + data.length + (entry.descriptor ? 16 : 0));
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true); lv.setUint16(4, 20, true); lv.setUint16(6, flags, true); lv.setUint16(8, method, true);
    lv.setUint16(12, 33, true); lv.setUint32(14, entry.descriptor ? 0 : crc, true);
    lv.setUint32(18, entry.descriptor ? 0 : data.length, true); lv.setUint32(22, entry.descriptor ? 0 : size, true);
    lv.setUint16(26, name.length, true); local.set(name, 30); local.set(data, 30 + name.length);
    if (entry.descriptor) {
      const at = 30 + name.length + data.length;
      lv.setUint32(at, 0x08074b50, true); lv.setUint32(at + 4, crc, true);
      lv.setUint32(at + 8, data.length, true); lv.setUint32(at + 12, size, true);
    }
    const cd = new Uint8Array(46 + name.length), cv = new DataView(cd.buffer);
    cv.setUint32(0, 0x02014b50, true); cv.setUint16(4, 20, true); cv.setUint16(6, 20, true);
    cv.setUint16(8, flags, true); cv.setUint16(10, method, true); cv.setUint16(14, 33, true);
    cv.setUint32(16, crc, true); cv.setUint32(20, data.length, true); cv.setUint32(24, size, true);
    cv.setUint16(28, name.length, true); cv.setUint32(42, offset, true); cd.set(name, 46);
    locals.push(local); central.push(cd); offset += local.length;
  }
  const cdSize = central.reduce((size, part) => size + part.length, 0);
  const out = new Uint8Array(offset + cdSize + 22 + comment.length);
  let pos = 0;
  for (const part of [...locals, ...central]) { out.set(part, pos); pos += part.length; }
  const end = new DataView(out.buffer, pos);
  end.setUint32(0, 0x06054b50, true); end.setUint16(8, entries.length, true); end.setUint16(10, entries.length, true);
  end.setUint32(12, cdSize, true); end.setUint32(16, offset, true); end.setUint16(20, comment.length, true);
  out.set(comment, pos + 22); return out;
}
function refused(action: () => unknown, reason: string): void {
  let caught: unknown;
  try { action(); } catch (error) { caught = error; }
  expect(caught).toBeInstanceOf(ArchiveRefused);
  expect(caught).toMatchObject({ reason });
}

describe("preserved source archive", () => {
  it("enforces the output bound inside the inflater", () => {
    const compressed = new Uint8Array(deflateRawSync(enc.encode("a".repeat(10000))));
    expect(() => inflateRaw(compressed, 20)).toThrow(/limit/i);
  });
  it("refuses truncated DEFLATE without reading synthetic zero bits", () => {
    expect(() => inflateRaw(new Uint8Array(), 100)).toThrow(/truncated/i);
    const compressed = new Uint8Array(deflateRawSync(enc.encode("a".repeat(10000))));
    expect(() => inflateRaw(compressed.slice(0, -1), 10001)).toThrow();
  });
  it("checks the stored-block length complement", () => {
    expect(() => inflateRaw(new Uint8Array([1, 1, 0, 1, 0, 65]), 10)).toThrow(/length/i);
  });
  it("retains original bytes, central order and opaque payloads with defensive copies", () => {
    const opaque = new Uint8Array([0, 255, 1, 13, 10]);
    const comment = new Uint8Array([...enc.encode("original comment PK"), 5, 6, ...new Uint8Array(22)]);
    const bytes = fixture([{ name: "unknown/payload.bin", data: opaque }, text("[Content_Types].xml", "<Types/>"), text("xl/worksheets/custom.xml", '<x value="&#65;">🦊</x>')], comment);
    const saved = bytes.slice(), source = SourceArchive.open(bytes);
    bytes[0] = 0; source.read("unknown/payload.bin")![0] = 55; source.toBytes()[0] = 0;
    expect(source.toBytes()).toEqual(saved);
    expect(source.names()).toEqual(["unknown/payload.bin", "[Content_Types].xml", "xl/worksheets/custom.xml"]);
    expect(source.read("unknown/payload.bin")).toEqual(opaque);
    expect(source.read("absent.xml")).toBeNull();
  });
  it.each([false, true])("reads DEFLATE and descriptor=%s with exact bytes", descriptor => {
    const payload = "Spreadsheet 🦊 ".repeat(120), bytes = fixture([text("opaque.xml", payload, { method: 8, descriptor })]);
    const source = SourceArchive.open(bytes);
    expect(source.read("opaque.xml")).toEqual(enc.encode(payload)); expect(source.toBytes()).toEqual(bytes);
  });
  it("refuses duplicate members", () => refused(() => SourceArchive.open(fixture([text("same.xml", "first"), text("same.xml", "second")])), "archive_duplicate"));
  it.each(["", "../escape", "/absolute", "a/../b", "a\\b", "C:/part.xml", "a//b", "./a", "a/./b", "nul\0part"])("refuses unsafe name %j", name =>
    refused(() => SourceArchive.open(fixture([text(name, "x")])), "archive_path"));
  it.each([{ flags: 0x0801 }, { method: 99 }])("refuses unsupported features %j", options =>
    refused(() => SourceArchive.open(fixture([text("x.xml", "x", options)])), "archive_unsupported"));
  it.each([{ crc: 123 }, { size: 99 }])("checks CRC and size %j", options =>
    refused(() => SourceArchive.open(fixture([text("x.xml", "abc", options)])), "archive_corrupt"));
  it("refuses local/central name disagreement", () => {
    const bytes = fixture([text("item.xml", "abc")]); bytes[30] = 88;
    refused(() => SourceArchive.open(bytes), "archive_corrupt");
  });
  it("refuses truncation", () => refused(() => SourceArchive.open(fixture([text("x.xml", "abc")]).slice(0, -4)), "archive_corrupt"));
  it.each([{ maxArchiveBytes: 10 }, { maxEntries: 1 }, { maxPartBytes: 3 }, { maxExpandedBytes: 7 }])("enforces limits %j", limits =>
    refused(() => SourceArchive.open(fixture([text("a.xml", "abcd"), text("b.xml", "efgh")]), limits), "archive_limit"));
  it("checks ratio before inflation", () => refused(() => SourceArchive.open(fixture([text("large.xml", "a".repeat(10000), { method: 8 })]), { maxCompressionRatio: 2 }), "archive_limit"));
  it("bounds inflation when both headers understate output", () => refused(() => SourceArchive.open(fixture([text("large.xml", "a".repeat(10000), { method: 8, size: 20 })])), "archive_corrupt"));
});

