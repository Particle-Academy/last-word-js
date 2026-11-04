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
 * This port parses text, not bytes, so a part arrives already decoded as UTF-8.
 */

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

/** Whether markup outside any element is a DOCTYPE: what a parser asks of the `<!` it has met. */
export function isDoctype(xml: string, at: number): boolean {
  return doctypeAt(xml, at);
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
