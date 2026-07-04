/**
 * Minimal ZIP writer using the STORE method (no compression). Valid OOXML
 * containers don't require compression, so this stays tiny, deterministic, and
 * isomorphic. Filenames are UTF-8 (general-purpose bit 11 set).
 */
import { crc32 } from "./crc32";

export interface ZipFile {
  name: string;
  data: Uint8Array;
}

const encoder = new TextEncoder();

// Fixed DOS timestamp (1980-01-01 00:00:00) for deterministic, reproducible output.
const DOS_TIME = 0;
const DOS_DATE = 0x0021;

export function zipSync(files: ZipFile[]): Uint8Array {
  const localParts: Uint8Array[] = [];
  const centralParts: Uint8Array[] = [];
  let offset = 0;

  for (const f of files) {
    const nameBytes = encoder.encode(f.name);
    const crc = crc32(f.data);
    const size = f.data.length;

    const lh = new Uint8Array(30 + nameBytes.length);
    const ldv = new DataView(lh.buffer);
    ldv.setUint32(0, 0x04034b50, true);
    ldv.setUint16(4, 20, true); // version needed to extract
    ldv.setUint16(6, 0x0800, true); // flags: UTF-8 filename
    ldv.setUint16(8, 0, true); // method: STORE
    ldv.setUint16(10, DOS_TIME, true);
    ldv.setUint16(12, DOS_DATE, true);
    ldv.setUint32(14, crc, true);
    ldv.setUint32(18, size, true); // compressed size
    ldv.setUint32(22, size, true); // uncompressed size
    ldv.setUint16(26, nameBytes.length, true);
    ldv.setUint16(28, 0, true); // extra length
    lh.set(nameBytes, 30);
    localParts.push(lh, f.data);

    const cd = new Uint8Array(46 + nameBytes.length);
    const cdv = new DataView(cd.buffer);
    cdv.setUint32(0, 0x02014b50, true);
    cdv.setUint16(4, 20, true); // version made by
    cdv.setUint16(6, 20, true); // version needed
    cdv.setUint16(8, 0x0800, true);
    cdv.setUint16(10, 0, true); // method
    cdv.setUint16(12, DOS_TIME, true);
    cdv.setUint16(14, DOS_DATE, true);
    cdv.setUint32(16, crc, true);
    cdv.setUint32(20, size, true);
    cdv.setUint32(24, size, true);
    cdv.setUint16(28, nameBytes.length, true);
    cdv.setUint16(30, 0, true); // extra length
    cdv.setUint16(32, 0, true); // comment length
    cdv.setUint16(34, 0, true); // disk number start
    cdv.setUint16(36, 0, true); // internal attrs
    cdv.setUint32(38, 0, true); // external attrs
    cdv.setUint32(42, offset, true); // local header offset
    cd.set(nameBytes, 46);
    centralParts.push(cd);

    offset += lh.length + f.data.length;
  }

  const localSize = offset;
  const centralSize = centralParts.reduce((s, c) => s + c.length, 0);

  const eocd = new Uint8Array(22);
  const edv = new DataView(eocd.buffer);
  edv.setUint32(0, 0x06054b50, true);
  edv.setUint16(4, 0, true); // this disk
  edv.setUint16(6, 0, true); // disk with central dir
  edv.setUint16(8, files.length, true);
  edv.setUint16(10, files.length, true);
  edv.setUint32(12, centralSize, true);
  edv.setUint32(16, localSize, true); // central dir offset
  edv.setUint16(20, 0, true); // comment length

  const out = new Uint8Array(localSize + centralSize + 22);
  let p = 0;
  for (const part of localParts) {
    out.set(part, p);
    p += part.length;
  }
  for (const part of centralParts) {
    out.set(part, p);
    p += part.length;
  }
  out.set(eocd, p);
  return out;
}
