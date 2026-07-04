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
