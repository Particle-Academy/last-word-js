/**
 * Intrinsic image dimension sniffing (PNG IHDR / JPEG SOF) + data-URL decode.
 * Zero-dependency, isomorphic — used by the writer to size `wp:extent` when
 * `widthPx` / `heightPx` are omitted.
 */

export interface ImageSize {
  width: number;
  height: number;
}

export interface DecodedDataUrl {
  mime: string;
  bytes: Uint8Array;
}

function base64Decode(b64: string): Uint8Array {
  if (typeof Buffer !== "undefined") {
    const buf = Buffer.from(b64, "base64");
    return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
  }
  // eslint-disable-next-line no-undef
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

export function base64Encode(bytes: Uint8Array): string {
  if (typeof Buffer !== "undefined") {
    return Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength).toString("base64");
  }
  let bin = "";
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]!);
  // eslint-disable-next-line no-undef
  return btoa(bin);
}

/** Decode a `data:image/…;base64,…` URL. Returns null for anything else. */
export function parseDataUrl(src: string): DecodedDataUrl | null {
  const m = /^data:([^;,]+);base64,([\s\S]*)$/.exec(src);
  if (!m) return null;
  try {
    return { mime: m[1]!.toLowerCase(), bytes: base64Decode(m[2]!.replace(/\s+/g, "")) };
  } catch {
    return null;
  }
}

/** PNG IHDR width/height, or null when the bytes are not a PNG. */
export function pngSize(bytes: Uint8Array): ImageSize | null {
  if (bytes.length < 24) return null;
  const sig = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
  for (let i = 0; i < 8; i++) {
    if (bytes[i] !== sig[i]) return null;
  }
  // Bytes 12-15 must be "IHDR"; width/height are big-endian at 16/20.
  if (bytes[12] !== 0x49 || bytes[13] !== 0x48 || bytes[14] !== 0x44 || bytes[15] !== 0x52) return null;
  const width = (bytes[16]! << 24) | (bytes[17]! << 16) | (bytes[18]! << 8) | bytes[19]!;
  const height = (bytes[20]! << 24) | (bytes[21]! << 16) | (bytes[22]! << 8) | bytes[23]!;
  if (width <= 0 || height <= 0) return null;
  return { width, height };
}

/**
 * JPEG width/height from the first SOF frame header (SOF0-SOF3, SOF5-SOF7,
 * SOF9-SOF11, SOF13-SOF15), or null when the bytes are not a JPEG.
 */
export function jpegSize(bytes: Uint8Array): ImageSize | null {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return null;
  let i = 2;
  while (i + 9 < bytes.length) {
    if (bytes[i] !== 0xff) {
      i++;
      continue;
    }
    const marker = bytes[i + 1]!;
    // Standalone markers without a length segment.
    if (marker === 0xd8 || (marker >= 0xd0 && marker <= 0xd7) || marker === 0x01 || marker === 0xff) {
      i += 2;
      continue;
    }
    const length = (bytes[i + 2]! << 8) | bytes[i + 3]!;
    const isSof =
      marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
    if (isSof) {
      const height = (bytes[i + 5]! << 8) | bytes[i + 6]!;
      const width = (bytes[i + 7]! << 8) | bytes[i + 8]!;
      if (width <= 0 || height <= 0) return null;
      return { width, height };
    }
    if (marker === 0xd9 || marker === 0xda) break; // EOI / start of scan
    i += 2 + length;
  }
  return null;
}

/** Sniff intrinsic pixel dimensions from PNG or JPEG bytes. */
export function sniffImageSize(bytes: Uint8Array): ImageSize | null {
  return pngSize(bytes) ?? jpegSize(bytes);
}
