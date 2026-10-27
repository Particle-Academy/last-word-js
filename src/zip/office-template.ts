/**
 * Office and LibreOffice TEMPLATE packages: `.dotx` / `.ott` for this package.
 * Mirrors PHP `LastWord\Archive\OfficeTemplate` (same family table, reason
 * codes and byte layout of a converted archive).
 *
 * A template is its document twin with a different declared type, nothing more:
 *
 * - OOXML (ECMA-376 Part 2 content types): the main part's Override in
 *   `[Content_Types].xml` reads `...wordprocessingml.template.main+xml` instead
 *   of `...document.main+xml`.
 * - ODF (1.3 Part 2 section 3.3): the first, STORED `mimetype` entry reads
 *   `...opendocument.text-template`, and that value SHALL equal the
 *   `manifest:media-type` of the manifest's root (`/`) entry. Rewriting the
 *   `mimetype` alone would produce a package that contradicts itself.
 *
 * So reading one is classifying it, and converting one is rewriting those
 * declarations while every other member is copied as its original raw zip
 * record: same compressed bytes, same CRC, same timestamps.
 *
 * Declarations are edited as latin1 strings (one char per byte) so the bytes
 * around the edited value survive exactly, whatever their encoding.
 *
 * Macro-enabled templates (`.dotm`, `.potm`, `.xltm`) are refused, never read
 * around. Every family is listed so a template declaring ANOTHER family is
 * recognised and refused rather than read as whatever the sniffer guessed.
 *
 * @internal
 */
import { OfficeTemplateRefused } from "../exceptions";
import { isElement, parseDom, type DomElement } from "../reader/odt/xml-dom";
import { crc32 } from "./crc32";
import { SourceArchive, type ArchiveEntry } from "./source-archive";

export type SourceKind = "document" | "template";

export const DOCUMENT: SourceKind = "document";
export const TEMPLATE: SourceKind = "template";

/** The family this package reads and converts. */
const FAMILY = "word";
/** Whether this package reads (and so converts) the family's ODF form. */
const ODF = true;

const VBA = "application/vnd.ms-office.vbaProject";

interface Family {
  part: string;
  document: string;
  template: string;
  macroDocument: string;
  macroTemplate: string;
  odfDocument: string;
  odfTemplate: string;
  odfBody: string;
}

const FAMILIES: Record<string, Family> = {
  word: {
    part: "/word/document.xml",
    document: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    template: "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
    macroDocument: "application/vnd.ms-word.document.macroEnabled.main+xml",
    macroTemplate: "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
    odfDocument: "application/vnd.oasis.opendocument.text",
    odfTemplate: "application/vnd.oasis.opendocument.text-template",
    odfBody: "text",
  },
  presentation: {
    part: "/ppt/presentation.xml",
    document: "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
    template: "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
    macroDocument: "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml",
    macroTemplate: "application/vnd.ms-powerpoint.template.macroEnabled.main+xml",
    odfDocument: "application/vnd.oasis.opendocument.presentation",
    odfTemplate: "application/vnd.oasis.opendocument.presentation-template",
    odfBody: "presentation",
  },
  spreadsheet: {
    part: "/xl/workbook.xml",
    document: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    template: "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
    macroDocument: "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
    macroTemplate: "application/vnd.ms-excel.template.macroEnabled.main+xml",
    odfDocument: "application/vnd.oasis.opendocument.spreadsheet",
    odfTemplate: "application/vnd.oasis.opendocument.spreadsheet-template",
    odfBody: "spreadsheet",
  },
};

const OWN = FAMILIES[FAMILY]!;
const UTF8 = new TextDecoder("utf-8");

/**
 * `document` or `template`, for a package this package's reader is about to
 * read. A package that declares nothing a template would is a `document`,
 * exactly as before: only templates are new.
 *
 * @throws OfficeTemplateRefused `macro_enabled` or `family_mismatch`
 */
export function sourceKind(zip: SourceArchive): SourceKind {
  const types = zip.read("[Content_Types].xml");
  if (types !== null) return ooxmlKind(zip, types);
  const mimetype = zip.read("mimetype");
  if (mimetype !== null) return odfKind(zip, latin1(mimetype));
  return DOCUMENT;
}

/**
 * Turn a template into its document twin, or a document into its template.
 * `to` null/undefined flips whichever it is.
 *
 * @throws OfficeTemplateRefused
 * @throws ArchiveRefused through the archive guard
 */
export function convert(bytes: Uint8Array, to: SourceKind | null = null): Uint8Array {
  if (to !== null && to !== DOCUMENT && to !== TEMPLATE) throw new OfficeTemplateRefused("bad_target");
  const zip = SourceArchive.open(bytes);
  const kind = sourceKind(zip);
  const target: SourceKind = to ?? (kind === TEMPLATE ? DOCUMENT : TEMPLATE);
  if (target === kind) throw new OfficeTemplateRefused(`already_${kind}`);

  const replacements = new Map<string, Uint8Array>();
  const types = zip.read("[Content_Types].xml");
  const mimetypeBytes = zip.read("mimetype");
  if (types !== null) {
    const declared = declarations(types);
    if (declared === null) throw new OfficeTemplateRefused("not_convertible");
    if (hasVba(zip, declared) || typesOf(declared, OWN.part).includes(OWN.macroDocument)) {
      throw new OfficeTemplateRefused("macro_enabled");
    }
    const [from, value] = target === TEMPLATE ? [OWN.document, OWN.template] : [OWN.template, OWN.document];
    replacements.set("[Content_Types].xml", fromLatin1(rewriteOverride(latin1(types), OWN.part, from, value)));
  } else if (ODF && mimetypeBytes !== null) {
    // First in the central directory AND first in the file, stored, no extra
    // field: ODF 1.3 Part 2 section 3.3.
    const entries = zip.entries();
    const first = entries[0];
    const byLocal = [...entries].sort((a, b) => a.local - b.local);
    if (first === undefined || first.name !== "mimetype" || byLocal[0]!.name !== "mimetype"
      || first.method !== 0 || first.localExtra !== 0) {
      throw new OfficeTemplateRefused("odf_mimetype");
    }
    const [from, value] = target === TEMPLATE ? [OWN.odfDocument, OWN.odfTemplate] : [OWN.odfTemplate, OWN.odfDocument];
    if (latin1(mimetypeBytes) !== from) throw new OfficeTemplateRefused("not_convertible");
    replacements.set("mimetype", fromLatin1(value));
    const manifest = zip.read("META-INF/manifest.xml");
    if (manifest !== null) {
      const rewritten = rewriteManifestRoot(latin1(manifest), from, value);
      if (rewritten !== null) replacements.set("META-INF/manifest.xml", fromLatin1(rewritten));
    }
  } else {
    throw new OfficeTemplateRefused("not_convertible");
  }

  const out = rebuild(zip, replacements);
  if (sourceKind(SourceArchive.open(out)) !== target) {
    throw new Error("Template conversion did not produce the requested kind.");
  }
  return out;
}

// ─── Classification ────────────────────────────────────────────────────

function ooxmlKind(zip: SourceArchive, types: Uint8Array): SourceKind {
  const declared = declarations(types);
  // Unparsable content types were never looked at before; the reader's own
  // handling of the package is unchanged.
  if (declared === null) return DOCUMENT;
  let kind = DOCUMENT;
  for (const [part, type] of declared) {
    for (const [name, family] of Object.entries(FAMILIES)) {
      if (type !== family.template && type !== family.macroTemplate) continue;
      // A template type on a Default, on another part, or of another family.
      if (name !== FAMILY || part !== asciiLower(OWN.part)) throw new OfficeTemplateRefused("family_mismatch");
      if (type === family.macroTemplate) throw new OfficeTemplateRefused("macro_enabled");
      kind = TEMPLATE;
    }
  }
  if (kind === TEMPLATE && hasVba(zip, declared)) throw new OfficeTemplateRefused("macro_enabled");
  return kind;
}

function odfKind(zip: SourceArchive, mimetype: string): SourceKind {
  if (mimetype !== OWN.odfTemplate) {
    if (mimetype.startsWith("application/vnd.oasis.opendocument.") && mimetype.endsWith("-template")) {
      throw new OfficeTemplateRefused("family_mismatch");
    }
    return DOCUMENT;
  }
  const content = dom(zip.read("content.xml"));
  const body = content !== null ? firstByLocalName(content, "body") : null;
  const first = body?.children.find(isElement) ?? null;
  if (first === null || first.local !== OWN.odfBody) throw new OfficeTemplateRefused("family_mismatch");
  const root = manifestRootType(dom(zip.read("META-INF/manifest.xml")));
  if (root !== null && root !== OWN.odfTemplate && root !== OWN.odfDocument) {
    throw new OfficeTemplateRefused("family_mismatch");
  }
  return TEMPLATE;
}

/** Every content-type declaration as [ASCII-lowercased part name or `*.ext`, type], by local name. */
function declarations(types: Uint8Array | string): [string, string][] | null {
  const root = dom(types);
  if (root === null) return null;
  const out: [string, string][] = [];
  for (const el of elements(root)) {
    if (el.local === "Override") out.push([asciiLower(el.attrs.get("PartName") ?? ""), el.attrs.get("ContentType") ?? ""]);
    else if (el.local === "Default") out.push([`*.${asciiLower(el.attrs.get("Extension") ?? "")}`, el.attrs.get("ContentType") ?? ""]);
  }
  return out;
}

function typesOf(declared: [string, string][], part: string): string[] {
  return declared.filter(([name]) => name === asciiLower(part)).map(([, type]) => type);
}

function hasVba(zip: SourceArchive, declared: [string, string][]): boolean {
  if (declared.some(([, type]) => asciiLower(type) === asciiLower(VBA))) return true;
  return zip.names().some((name) => /(^|\/)vbaProject\.bin$/i.test(name));
}

function manifestRootType(root: DomElement | null): string | null {
  if (root === null) return null;
  for (const el of elements(root)) {
    if (el.local === "file-entry" && attr(el, "full-path") === "/") return attr(el, "media-type");
  }
  return null;
}

// ─── Rewriting ─────────────────────────────────────────────────────────

/**
 * Replace one attribute VALUE in place. A text scan finds the candidate; the
 * parser then has to agree on the result, which is what stops a decoy in a
 * comment or an entity-encoded name from steering the edit.
 */
function rewriteOverride(xml: string, part: string, from: string, to: string): string {
  const hits: number[] = [];
  for (const tag of xml.matchAll(/<(?:[A-Za-z_][\w.\-]*:)?Override\b[^>]*>/g)) {
    const name = /\sPartName\s*=\s*(["'])(.*?)\1/s.exec(tag[0]);
    if (name === null || asciiLower(name[2]!) !== asciiLower(part)) continue;
    const type = /\sContentType\s*=\s*(["'])(.*?)\1/s.exec(tag[0]);
    if (type !== null && type[2] === from) hits.push(tag.index! + type.index + type[0].indexOf(type[1]!) + 1);
  }
  if (hits.length !== 1) throw new OfficeTemplateRefused("not_convertible");
  const out = xml.slice(0, hits[0]) + to + xml.slice(hits[0]! + from.length);
  const declared = declarations(out);
  const after = declared === null ? [] : typesOf(declared, part);
  if (after.length !== 1 || after[0] !== to) throw new OfficeTemplateRefused("not_convertible");
  return out;
}

/** The manifest with its root entry's media-type moved; null when it has no root entry. */
function rewriteManifestRoot(xml: string, from: string, to: string): string | null {
  const current = manifestRootType(dom(xml));
  if (current === null) return null;
  if (current !== from) throw new OfficeTemplateRefused("not_convertible");
  const hits: number[] = [];
  for (const tag of xml.matchAll(/<(?:[A-Za-z_][\w.\-]*:)?file-entry\b[^>]*>/g)) {
    if (!/\s(?:[A-Za-z_][\w.\-]*:)?full-path\s*=\s*(["'])\/\1/.test(tag[0])) continue;
    const type = /\s(?:[A-Za-z_][\w.\-]*:)?media-type\s*=\s*(["'])(.*?)\1/s.exec(tag[0]);
    if (type !== null && type[2] === from) hits.push(tag.index! + type.index + type[0].indexOf(type[1]!) + 1);
  }
  if (hits.length !== 1) throw new OfficeTemplateRefused("not_convertible");
  const out = xml.slice(0, hits[0]) + to + xml.slice(hits[0]! + from.length);
  if (manifestRootType(dom(out)) !== to) throw new OfficeTemplateRefused("not_convertible");
  return out;
}

/**
 * Re-assemble the archive. Untouched members are their original local records,
 * verbatim; a replaced member is written STORED with its original timestamp and
 * no extra field (which is what ODF requires of `mimetype`). Local order and
 * central order are both kept.
 */
function rebuild(zip: SourceArchive, replacements: Map<string, Uint8Array>): Uint8Array {
  // A real copy: when the caller handed in a Node Buffer, `toBytes()` returns a
  // Buffer and Buffer#slice is a VIEW, so `.buffer`-relative writes would land
  // in the wrong bytes.
  const original = new Uint8Array(zip.toBytes());
  const view = new DataView(original.buffer, original.byteOffset, original.byteLength);
  const u16 = (at: number): number => view.getUint16(at, true);
  const u32 = (at: number): number => view.getUint32(at, true);
  const entries = zip.entries();
  const byLocal = [...entries].sort((a, b) => a.local - b.local);
  const nameBytes = (entry: ArchiveEntry): Uint8Array => original.subarray(entry.central + 46, entry.central + 46 + u16(entry.central + 28));

  const chunks: Uint8Array[] = [];
  const offsets = new Map<string, number>();
  let length = 0;
  const push = (chunk: Uint8Array): void => { chunks.push(chunk); length += chunk.length; };

  for (const entry of byLocal) {
    offsets.set(entry.name, length);
    const data = replacements.get(entry.name);
    if (data === undefined) {
      push(original.subarray(entry.local, entry.finish));
      continue;
    }
    const name = nameBytes(entry);
    const header = new Uint8Array(30 + name.length + data.length);
    const hv = new DataView(header.buffer);
    hv.setUint32(0, 0x04034b50, true); hv.setUint16(4, 10, true); hv.setUint16(6, entry.flags & 0x0800, true);
    hv.setUint16(8, 0, true); hv.setUint16(10, u16(entry.local + 10), true); hv.setUint16(12, u16(entry.local + 12), true);
    hv.setUint32(14, crc32(data) >>> 0, true); hv.setUint32(18, data.length, true); hv.setUint32(22, data.length, true);
    hv.setUint16(26, name.length, true); hv.setUint16(28, 0, true);
    header.set(name, 30); header.set(data, 30 + name.length);
    push(header);
  }
  const localsLength = length;

  for (const entry of entries) {
    const data = replacements.get(entry.name);
    if (data === undefined) {
      const record = new Uint8Array(original.subarray(entry.central, entry.central + entry.centralLength));
      new DataView(record.buffer, record.byteOffset, record.byteLength).setUint32(42, offsets.get(entry.name)!, true);
      push(record);
      continue;
    }
    const name = nameBytes(entry);
    const at = entry.central;
    const record = new Uint8Array(46 + name.length);
    const rv = new DataView(record.buffer);
    rv.setUint32(0, 0x02014b50, true); rv.setUint16(4, u16(at + 4), true); rv.setUint16(6, 10, true);
    rv.setUint16(8, entry.flags & 0x0800, true); rv.setUint16(10, 0, true);
    rv.setUint16(12, u16(at + 12), true); rv.setUint16(14, u16(at + 14), true);
    rv.setUint32(16, crc32(data) >>> 0, true); rv.setUint32(20, data.length, true); rv.setUint32(24, data.length, true);
    rv.setUint16(28, name.length, true); rv.setUint16(36, u16(at + 36), true); rv.setUint32(38, u32(at + 38), true);
    rv.setUint32(42, offsets.get(entry.name)!, true);
    record.set(name, 46);
    push(record);
  }

  const comment = zip.comment();
  const end = new Uint8Array(22 + comment.length);
  const ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054b50, true); ev.setUint16(8, entries.length, true); ev.setUint16(10, entries.length, true);
  ev.setUint32(12, length - localsLength, true); ev.setUint32(16, localsLength, true); ev.setUint16(20, comment.length, true);
  end.set(comment, 22);
  push(end);

  const out = new Uint8Array(length);
  let pos = 0;
  for (const chunk of chunks) { out.set(chunk, pos); pos += chunk.length; }
  return out;
}

// ─── Helpers ───────────────────────────────────────────────────────────

/** Strict parse; null when the part is absent, empty, carries a DOCTYPE or does not parse. */
function dom(xml: Uint8Array | string | null): DomElement | null {
  if (xml === null) return null;
  const text = typeof xml === "string" ? UTF8.decode(fromLatin1(xml)) : UTF8.decode(xml);
  if (text === "" || /<!DOCTYPE/i.test(text)) return null;
  try {
    return parseDom(text);
  } catch {
    return null;
  }
}

function* elements(root: DomElement): Generator<DomElement> {
  const stack: DomElement[] = [root];
  while (stack.length > 0) {
    const el = stack.pop()!;
    if (el.local !== "#document") yield el;
    for (let i = el.children.length - 1; i >= 0; i--) {
      const child = el.children[i];
      if (isElement(child)) stack.push(child);
    }
  }
}

function firstByLocalName(root: DomElement, local: string): DomElement | null {
  for (const el of elements(root)) if (el.local === local) return el;
  return null;
}

/** An attribute by local name, whatever its prefix. */
function attr(el: DomElement, local: string): string | null {
  for (const [name, value] of el.attrs) {
    if (name === local || name.endsWith(`:${local}`)) return value;
  }
  return null;
}

function asciiLower(s: string): string {
  return s.replace(/[A-Z]/g, (c) => String.fromCharCode(c.charCodeAt(0) + 32));
}

function latin1(bytes: Uint8Array): string {
  let out = "";
  for (let i = 0; i < bytes.length; i += 0x8000) out += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return out;
}

function fromLatin1(s: string): Uint8Array {
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i) & 0xff;
  return out;
}
