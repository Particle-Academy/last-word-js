/**
 * Minimal ZIP reader. Parses the central directory and returns a name → bytes
 * map. Supports STORE (method 0) and DEFLATE (method 8, via `inflateRaw`) so it
 * reads both our own output and Excel/PowerPoint-authored files. Zip64 is not
 * handled (OOXML parts are well under 4 GB).
 */
import { inflateRaw } from "./inflate";

const decoder = new TextDecoder();

export function unzipSync(data: Uint8Array): Record<string, Uint8Array> {
  const dv = new DataView(data.buffer, data.byteOffset, data.byteLength);

  // Locate the End Of Central Directory record (scan backward; it may carry a comment).
  let eocd = -1;
  for (let i = data.length - 22; i >= 0; i--) {
    if (dv.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error("dark-slide: not a zip archive (no EOCD record)");

  const count = dv.getUint16(eocd + 10, true);
  let cd = dv.getUint32(eocd + 16, true);
  const result: Record<string, Uint8Array> = {};

  for (let n = 0; n < count; n++) {
    if (dv.getUint32(cd, true) !== 0x02014b50) break;
    const method = dv.getUint16(cd + 10, true);
    const compSize = dv.getUint32(cd + 20, true);
    const nameLen = dv.getUint16(cd + 28, true);
    const extraLen = dv.getUint16(cd + 30, true);
    const commentLen = dv.getUint16(cd + 32, true);
    const localOffset = dv.getUint32(cd + 42, true);
    const name = decoder.decode(data.subarray(cd + 46, cd + 46 + nameLen));

    const lhNameLen = dv.getUint16(localOffset + 26, true);
    const lhExtraLen = dv.getUint16(localOffset + 28, true);
    const dataStart = localOffset + 30 + lhNameLen + lhExtraLen;
    const raw = data.subarray(dataStart, dataStart + compSize);

    let content: Uint8Array;
    if (method === 0) content = raw.slice();
    else if (method === 8) content = inflateRaw(raw);
    else throw new Error(`dark-slide: unsupported zip method ${method} for "${name}"`);

    result[name] = content;
    cd += 46 + nameLen + extraLen + commentLen;
  }

  return result;
}

/**
 * Read only the named entries, each refused if it declares or inflates to more
 * than `maxBytes`. For untrusted uploads: `unzipSync` inflates every entry
 * without a limit, which is fine for our own output and not for a stranger's.
 * Throws on any structure that points outside the archive.
 */
export function unzipEntries(data: Uint8Array, names: string[], maxBytes: number): Record<string, Uint8Array> {
  const dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const u16 = (o: number): number => {
    if (o < 0 || o + 2 > data.length) throw new Error("zip structure points outside the archive");
    return dv.getUint16(o, true);
  };
  const u32 = (o: number): number => {
    if (o < 0 || o + 4 > data.length) throw new Error("zip structure points outside the archive");
    return dv.getUint32(o, true);
  };

  let eocd = -1;
  for (let i = data.length - 22; i >= 0 && i >= data.length - 22 - 0xffff; i--) {
    if (dv.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error("not a zip archive (no EOCD record)");

  const count = u16(eocd + 10);
  let cd = u32(eocd + 16);
  const wanted = new Set(names);
  const result: Record<string, Uint8Array> = {};

  for (let n = 0; n < count; n++) {
    if (u32(cd) !== 0x02014b50) break;
    const method = u16(cd + 10);
    const compSize = u32(cd + 20);
    const size = u32(cd + 24);
    const nameLen = u16(cd + 28);
    const extraLen = u16(cd + 30);
    const commentLen = u16(cd + 32);
    const localOffset = u32(cd + 42);
    if (cd + 46 + nameLen > data.length) throw new Error("zip structure points outside the archive");
    const name = decoder.decode(data.subarray(cd + 46, cd + 46 + nameLen));
    cd += 46 + nameLen + extraLen + commentLen;

    if (!wanted.has(name) || result[name] !== undefined) continue;
    if (size > maxBytes) throw new Error(`zip entry ${name} is too large to read`);

    const dataStart = localOffset + 30 + u16(localOffset + 26) + u16(localOffset + 28);
    if (dataStart + compSize > data.length) throw new Error("zip structure points outside the archive");
    const raw = data.subarray(dataStart, dataStart + compSize);

    if (method === 0) result[name] = raw.slice(0, Math.min(raw.length, maxBytes));
    else if (method === 8) {
      try {
        result[name] = inflateRaw(raw, maxBytes);
      } catch (e) {
        if (/larger than allowed/.test((e as Error).message)) throw new Error(`zip entry ${name} is too large to read`);
        throw e;
      }
    }
    else throw new Error(`unsupported zip method ${method} for "${name}"`);
  }

  return result;
}
