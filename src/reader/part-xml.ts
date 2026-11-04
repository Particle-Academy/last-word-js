/**
 * The one way an XML part of an uploaded file gets parsed. Mirrors PHP
 * `Archive\PartXml`.
 *
 * A part carrying a DOCTYPE is refused, because no Office or ODF part has one and
 * an internal subset is the entry point for entity expansion. A refusal looks
 * exactly like malformed XML, so each caller keeps its own handling of a part it
 * cannot read.
 *
 * The check has two halves, and the second is the one that decides:
 *
 * - `prologDeclaresDoctype()` looks at the text before anything is parsed. It reads
 *   only the prolog: searching the whole part for the word would refuse legitimate
 *   text.
 * - Each parser (`parseXml` in `xml.ts`, `parseDom` in `odt/xml-dom.ts`) starts here
 *   and refuses a DOCTYPE it meets itself.
 *
 * This port parses text, so decode a part through `decodePart`: the check then sees
 * what the part actually says, whatever encoding it uses.
 */

const UTF8 = new TextDecoder();

/** The encoding a part is written in, or null when this port cannot decode it. */
function partEncoding(b: Uint8Array): string | null {
  if (b[0] === 0xff && b[1] === 0xfe) return b[2] === 0 && b[3] === 0 ? null : "utf-16le";
  if (b[0] === 0xfe && b[1] === 0xff) return "utf-16be";
  if (b[0] === 0 && b[1] === 0) return null;
  if (b[0] === 0x3c && b[1] === 0) return b[2] === 0 ? null : "utf-16le";
  if (b[0] === 0 && b[1] === 0x3c) return "utf-16be";
  const start = b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf ? 3 : 0;
  const head = String.fromCharCode(...b.subarray(start, start + 200));
  const label = /^<\?xml\s[^>]*?encoding\s*=\s*["']([A-Za-z0-9._-]+)["']/.exec(head)?.[1]?.toLowerCase();
  // Bytes that got this far are not UTF-16, whatever the declaration claims.
  if (label === undefined || label === "utf-8" || label.startsWith("utf-16") || label === "utf16") return "utf-8";
  try {
    new TextDecoder(label);
    return label;
  } catch {
    return null;
  }
}

function decodeOrNull(bytes: Uint8Array): string | null {
  const encoding = partEncoding(bytes);
  if (encoding === null) return null;
  return encoding === "utf-8" ? UTF8.decode(bytes) : new TextDecoder(encoding).decode(bytes);
}

/** A part's text, or "" when this port cannot decode it, which parses as nothing. */
export function decodePart(bytes: Uint8Array): string {
  return decodeOrNull(bytes) ?? "";
}

/** Whether a part is refused, as opposed to merely malformed or fine. */
export function partIsRefused(bytes: Uint8Array): boolean {
  const text = decodeOrNull(bytes);
  return text === null || declaresDoctype(text);
}

/** A DOCTYPE at this offset. Names are case-sensitive in XML; the check is not, on purpose. */
function doctypeAt(xml: string, at: number): boolean {
  return xml.slice(at, at + 9).toUpperCase() === "<!DOCTYPE";
}

/** Whether a DOCTYPE comes before the root element. */
export function prologDeclaresDoctype(source: string): boolean {
  const xml = source.slice(0, 4).includes("\0") ? source.split("\0").join("") : source;
  let at = 0;
  while (at < xml.length && (xml[at] === "﻿" || xml[at] === "�" || xml[at] === "ÿ" || xml[at] === "þ")) at++;
  for (;;) {
    while (at < xml.length && " \t\r\n".includes(xml[at]!)) at++;
    const close = xml.startsWith("<?", at) ? "?>" : xml.startsWith("<!--", at) ? "-->" : null;
    if (close === null) return doctypeAt(xml, at);
    const end = xml.indexOf(close, at);
    if (end < 0) return false;
    at = end + close.length;
  }
}

/**
 * Whether a part declares a DOCTYPE, for a part that is carried or named in an error
 * rather than read. It walks the part as the parsers do, skipping comments, CDATA and
 * processing instructions, so a part that merely mentions the word is not one.
 */
export function declaresDoctype(xml: string): boolean {
  if (prologDeclaresDoctype(xml)) return true;
  let at = 0;
  for (;;) {
    const open = xml.indexOf("<!", at);
    if (open < 0) return false;
    if (xml.startsWith("<!--", open)) {
      const end = xml.indexOf("-->", open);
      if (end < 0) return false;
      at = end + 3;
    } else if (xml.startsWith("<![CDATA[", open)) {
      const end = xml.indexOf("]]>", open);
      if (end < 0) return false;
      at = end + 3;
    } else if (doctypeAt(xml, open)) {
      return true;
    } else {
      at = open + 2;
    }
  }
}
