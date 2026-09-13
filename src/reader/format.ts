/**
 * What are these bytes? Mirrors PHP `Reader\Format`.
 *
 * Sniffs CONTENT, never a file name. A name is a claim by whoever chose it; the
 * signature is what the bytes actually are, and the two disagree often enough
 * that trusting the name is how a `.docx` that is really a `.doc` gets reported
 * as a corrupt archive.
 */

export const Format = {
  DOCX: "docx",
  ODT: "odt",
  DOC: "doc",
  RTF: "rtf",
  XLSX: "xlsx",
  PPTX: "pptx",
  ODS: "ods",
  ODP: "odp",
  UNKNOWN: "unknown",
} as const;

/** OLE2 / Compound File Binary: Word 97-2003 `.doc`, and `.xls`, `.ppt`, `.msg`. */
const OLE2 = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

export function detectFormat(bytes: Uint8Array): string {
  if (startsWith(bytes, OLE2, 0)) return Format.DOC;
  if (looksLikeRtf(bytes)) return Format.RTF;
  if (startsWith(bytes, [0x50, 0x4b, 0x03, 0x04], 0)) return zipFlavour(bytes);
  return Format.UNKNOWN;
}

/** An RTF opens `{\rtf`, possibly behind a BOM or leading whitespace. */
function looksLikeRtf(bytes: Uint8Array): boolean {
  let head = bytes.subarray(0, 16);
  if (startsWith(head, [0xef, 0xbb, 0xbf], 0)) head = head.subarray(3);
  let i = 0;
  // PHP's ltrim() default set: space, tab, LF, CR, NUL, vertical tab.
  while (i < head.length && [0x20, 0x09, 0x0a, 0x0d, 0x00, 0x0b].includes(head[i]!)) i++;
  return startsWith(head, [0x7b, 0x5c, 0x72, 0x74, 0x66], i); // {\rtf
}

/** Both docx and odt are zips, so the signature alone is not an answer. */
function zipFlavour(bytes: Uint8Array): string {
  const head = bytes.subarray(0, 256);
  if (contains(head, "application/vnd.oasis.opendocument.text")) return Format.ODT;
  if (contains(head, "application/vnd.oasis.opendocument.spreadsheet")) return Format.ODS;
  if (contains(head, "application/vnd.oasis.opendocument.presentation")) return Format.ODP;
  if (contains(bytes, "word/document.xml")) return Format.DOCX;
  if (contains(bytes, "content.xml")) return Format.ODT;
  if (contains(bytes, "xl/workbook.xml")) return Format.XLSX;
  if (contains(bytes, "ppt/presentation.xml")) return Format.PPTX;
  return Format.UNKNOWN;
}

function startsWith(bytes: Uint8Array, prefix: number[], at: number): boolean {
  if (bytes.length - at < prefix.length) return false;
  return prefix.every((b, i) => bytes[at + i] === b);
}

/** Whether the bytes contain an ASCII needle. */
export function contains(bytes: Uint8Array, needle: string): boolean {
  const first = needle.charCodeAt(0);
  const last = bytes.length - needle.length;
  for (let i = bytes.indexOf(first); i >= 0 && i <= last; i = bytes.indexOf(first, i + 1)) {
    let j = 1;
    while (j < needle.length && bytes[i + j] === needle.charCodeAt(j)) j++;
    if (j === needle.length) return true;
  }
  return false;
}
