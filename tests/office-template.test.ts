import { describe, expect, it } from "vitest";
import { deflateRawSync, inflateRawSync, crc32 as zlibCrc32 } from "node:zlib";
import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync, mkdtempSync, readFileSync, writeFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Agent } from "../src/agent";
import { ArchiveRefused, OfficeTemplateRefused, UnsupportedFormatException, convertTemplate } from "../src/index";
import { OdtReader } from "../src/reader/odt-reader";
import { run } from "../src/cli-run";

/*
 * Official template formats: Word .dotx and LibreOffice .ott. Mirrors PHP
 * tests/Unit/OfficeTemplateTest.php.
 *
 * A template IS its document twin with a different declared type: the main
 * part's content type in [Content_Types].xml (OOXML), or the `mimetype` entry
 * plus the manifest's root entry (ODF 1.3 Part 2 section 3.3 requires the two to
 * be equal). Every fixture is built and every archive is read by THIS file with
 * node:zlib, independently of the production zip code and converter, so a
 * converter bug cannot also forge its own evidence.
 */

const DOC = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const TEMPLATE = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml";
const MACRO_TEMPLATE = "application/vnd.ms-word.template.macroEnabledTemplate.main+xml";
const MACRO_DOC = "application/vnd.ms-word.document.macroEnabled.main+xml";
const PPT_TEMPLATE = "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml";
const ODT = "application/vnd.oasis.opendocument.text";
const OTT = "application/vnd.oasis.opendocument.text-template";
const RED_PNG =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEElEQVR42mP4z8AARAwQCgAf7gP9Y167WwAAAABJRU5ErkJggg==";

const enc = new TextEncoder();
const dec = new TextDecoder();

type Entries = Map<string, Uint8Array>;

function doc(): unknown {
  return {
    title: "Letterhead",
    blocks: [
      { type: "heading", level: 1, runs: [{ text: "Dear reader" }] },
      { type: "paragraph", runs: [{ text: "Body text, " }, { text: "bold", bold: true }] },
      { type: "image", src: RED_PNG, widthPx: 2, heightPx: 2 },
    ],
  };
}

/** name => bytes in archive order, plus each entry's compression method. Test-side reader. */
function unzip(bytes: Uint8Array): { entries: Entries; methods: Map<string, number> } {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let end = bytes.length - 22;
  while (end >= 0 && view.getUint32(end, true) !== 0x06054b50) end--;
  const count = view.getUint16(end + 10, true);
  let at = view.getUint32(end + 16, true);
  const entries: Entries = new Map();
  const methods = new Map<string, number>();
  for (let i = 0; i < count; i++) {
    const method = view.getUint16(at + 10, true);
    const compressed = view.getUint32(at + 20, true);
    const nameLength = view.getUint16(at + 28, true);
    const extra = view.getUint16(at + 30, true);
    const comment = view.getUint16(at + 32, true);
    const local = view.getUint32(at + 42, true);
    const name = dec.decode(bytes.subarray(at + 46, at + 46 + nameLength));
    const start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
    const raw = bytes.subarray(start, start + compressed);
    entries.set(name, method === 8 ? new Uint8Array(inflateRawSync(raw)) : raw.slice());
    methods.set(name, method);
    at += 46 + nameLength + extra + comment;
  }
  return { entries, methods };
}

/** Re-zip in order; everything DEFLATEd except `mimetype`, STORED as ODF requires. Test-side writer. */
function zip(entries: Entries): Uint8Array {
  const locals: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;
  for (const [entryName, data] of entries) {
    const name = enc.encode(entryName);
    const method = entryName === "mimetype" ? 0 : 8;
    const body = method === 8 ? new Uint8Array(deflateRawSync(data)) : data;
    const crc = zlibCrc32(data) >>> 0;
    const local = new Uint8Array(30 + name.length + body.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true); lv.setUint16(4, 20, true); lv.setUint16(6, 0x0800, true);
    lv.setUint16(8, method, true); lv.setUint16(12, 33, true); lv.setUint32(14, crc, true);
    lv.setUint32(18, body.length, true); lv.setUint32(22, data.length, true); lv.setUint16(26, name.length, true);
    local.set(name, 30); local.set(body, 30 + name.length);
    const cd = new Uint8Array(46 + name.length);
    const cv = new DataView(cd.buffer);
    cv.setUint32(0, 0x02014b50, true); cv.setUint16(4, 20, true); cv.setUint16(6, 20, true);
    cv.setUint16(8, 0x0800, true); cv.setUint16(10, method, true); cv.setUint16(14, 33, true);
    cv.setUint32(16, crc, true); cv.setUint32(20, body.length, true); cv.setUint32(24, data.length, true);
    cv.setUint16(28, name.length, true); cv.setUint32(42, offset, true); cd.set(name, 46);
    locals.push(local); central.push(cd); offset += local.length;
  }
  const cdSize = central.reduce((n, part) => n + part.length, 0);
  const out = new Uint8Array(offset + cdSize + 22);
  let pos = 0;
  for (const part of [...locals, ...central]) { out.set(part, pos); pos += part.length; }
  const ev = new DataView(out.buffer, pos);
  ev.setUint32(0, 0x06054b50, true); ev.setUint16(8, entries.size, true); ev.setUint16(10, entries.size, true);
  ev.setUint32(12, cdSize, true); ev.setUint32(16, offset, true);
  return out;
}

const text = (entries: Entries, name: string): string => dec.decode(entries.get(name)!);

function docx(): Uint8Array {
  // The writer STOREs; re-zipped here so untouched parts are DEFLATEd and the
  // "copied as records" assertion below has a compression method to keep.
  return zip(unzip(Agent.toBytes(doc())).entries);
}

function retyped(contentType: string, extra: Record<string, Uint8Array | string> = {}): Uint8Array {
  const entries = unzip(docx()).entries;
  entries.set("[Content_Types].xml", enc.encode(text(entries, "[Content_Types].xml").replace(DOC, contentType)));
  for (const [name, data] of Object.entries(extra)) entries.set(name, typeof data === "string" ? enc.encode(data) : data);
  return zip(entries);
}

const dotx = (): Uint8Array => retyped(TEMPLATE);

function odf(mimetype: string, body = "text", manifestType: string = mimetype): Uint8Array {
  return zip(new Map<string, Uint8Array>([
    ["mimetype", enc.encode(mimetype)],
    ["content.xml", enc.encode(
      '<?xml version="1.0" encoding="UTF-8"?>'
      + '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"'
      + ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">'
      + `<office:body><office:${body}><text:h text:outline-level="1">A Heading</text:h><text:p>Hello from ODF</text:p></office:${body}></office:body>`
      + "</office:document-content>")],
    ["meta.xml", enc.encode('<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/"><office:meta><dc:title>Memo</dc:title></office:meta></office:document-meta>')],
    ["Pictures/opaque.bin", new Uint8Array([0, 1, 2, 98, 105, 110, 255])],
    ["META-INF/manifest.xml", enc.encode(
      '<?xml version="1.0" encoding="UTF-8"?>\n'
      + '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">\n'
      + ` <manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="${manifestType}"/>\n`
      + ' <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>\n'
      + "</manifest:manifest>")],
  ]));
}

function refused(action: () => unknown, reason: string): void {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(OfficeTemplateRefused);
    expect((error as OfficeTemplateRefused).reason).toBe(reason);
    return;
  }
  throw new Error(`Expected a ${reason} refusal; the input was accepted.`);
}

function archiveRefused(action: () => unknown, reason: string): void {
  try {
    action();
  } catch (error) {
    expect(error).toBeInstanceOf(ArchiveRefused);
    expect((error as ArchiveRefused).reason).toBe(reason);
    return;
  }
  throw new Error(`Expected archive refusal ${reason}; the input was accepted.`);
}

const same = (a: Uint8Array | undefined, b: Uint8Array | undefined): boolean =>
  a !== undefined && b !== undefined && Buffer.from(a).equals(Buffer.from(b));

function expectSameEntries(a: Entries, b: Entries): void {
  expect([...a.keys()]).toEqual([...b.keys()]);
  for (const [name, data] of a) expect(same(data, b.get(name)), `${name} changed`).toBe(true);
}

describe("reading templates", () => {
  it("reads a Word .dotx and marks it as a template", () => {
    const document = Agent.read(docx()) as Record<string, unknown>;
    const template = Agent.read(dotx()) as Record<string, unknown>;
    expect(template.sourceKind).toBe("template");
    expect(Object.keys(template).at(-1)).toBe("sourceKind");
    delete template.sourceKind;
    expect(template).toEqual(document);
  });

  it("leaves a .docx read exactly as it was: no sourceKind key", () => {
    expect(Agent.read(docx())).not.toHaveProperty("sourceKind");
  });

  it("reads a LibreOffice .ott and marks it as a template", () => {
    const document = Agent.read(odf(ODT)) as Record<string, unknown>;
    const template = Agent.read(odf(OTT)) as Record<string, unknown>;
    expect(document).not.toHaveProperty("sourceKind");
    expect(template.sourceKind).toBe("template");
    delete template.sourceKind;
    expect(template).toEqual(document);
  });

  it("refuses a macro-enabled .dotm by name rather than reading around the VBA", () => {
    refused(() => Agent.read(retyped(MACRO_TEMPLATE)), "macro_enabled");
  });

  it("refuses a .dotx that carries a VBA project anyway", () => {
    refused(() => Agent.read(retyped(TEMPLATE, { "word/vbaProject.bin": "not really" })), "macro_enabled");
  });

  it("still reads a macro-enabled DOCUMENT (.docm) as before: only templates changed", () => {
    expect(Agent.read(retyped(MACRO_DOC))).not.toHaveProperty("sourceKind");
  });
});

describe("hostile templates", () => {
  it("refuses a template whose content type claims another family", () => {
    refused(() => Agent.read(retyped(PPT_TEMPLATE)), "family_mismatch");

    let entries = unzip(docx()).entries;
    entries.set("[Content_Types].xml", enc.encode(text(entries, "[Content_Types].xml").replace("</Types>",
      `<Override PartName="/ppt/presentation.xml" ContentType="${PPT_TEMPLATE}"/></Types>`)));
    refused(() => Agent.read(zip(entries)), "family_mismatch");

    entries = unzip(docx()).entries;
    entries.set("[Content_Types].xml", enc.encode(text(entries, "[Content_Types].xml").replace(
      'ContentType="application/xml"', `ContentType="${TEMPLATE}"`)));
    refused(() => Agent.read(zip(entries)), "family_mismatch");
  });

  it("refuses an ODF text template whose body is a spreadsheet", () => {
    refused(() => new OdtReader().read(odf(OTT, "spreadsheet")), "family_mismatch");
  });

  it("refuses an ODF template of another family, by name or by refusal", () => {
    refused(() => new OdtReader().read(odf("application/vnd.oasis.opendocument.presentation-template")), "family_mismatch");
    expect(() => Agent.read(odf("application/vnd.oasis.opendocument.spreadsheet-template", "spreadsheet")))
      .toThrow(UnsupportedFormatException);
  });

  it("refuses an ODF template whose manifest names another family", () => {
    refused(() => Agent.read(odf(OTT, "text", "application/vnd.oasis.opendocument.spreadsheet-template")), "family_mismatch");
  });

  it("refuses a template zip bomb through the archive guard, reading and converting", () => {
    const bomb = retyped(TEMPLATE, { "word/media/bomb.bin": new Uint8Array(65 * 1024 * 1024) });
    archiveRefused(() => Agent.read(bomb), "archive_limit");
    archiveRefused(() => convertTemplate(bomb), "archive_limit");
  });

  it("refuses a template with a traversing entry name, reading and converting", () => {
    const evil = retyped(TEMPLATE, { "../evil.xml": "<x/>" });
    archiveRefused(() => Agent.read(evil), "archive_path");
    archiveRefused(() => convertTemplate(evil), "archive_path");
  });
});

describe("converting templates", () => {
  it("converts a .dotx to a .docx changing only the content-type value", () => {
    const input = dotx();
    const output = Agent.convertTemplate(input, "document");
    const before = unzip(input).entries;
    const after = unzip(output).entries;
    expect([...after.keys()]).toEqual([...before.keys()]);
    for (const [name, data] of before) {
      if (name === "[Content_Types].xml") expect(text(after, name)).toBe(dec.decode(data).replace(TEMPLATE, DOC));
      else expect(same(after.get(name), data), `${name} changed`).toBe(true);
    }
    expect(Agent.read(output)).not.toHaveProperty("sourceKind");
  });

  it("round-trips docx to dotx and back with every part byte-identical", () => {
    const original = docx();
    const template = convertTemplate(original, "template");
    const back = convertTemplate(template, "document");
    expect((Agent.read(template) as Record<string, unknown>).sourceKind).toBe("template");
    expectSameEntries(unzip(back).entries, unzip(original).entries);

    // Untouched parts are copied as RECORDS, not re-encoded: same compression.
    const methods = unzip(original).methods;
    const converted = unzip(template).methods;
    methods.delete("[Content_Types].xml");
    converted.delete("[Content_Types].xml");
    expect(converted).toEqual(methods);
    expect([...methods.values()].every((m) => m === 8)).toBe(true);
  });

  it("converts a Node Buffer as it converts a Uint8Array", () => {
    // Buffer.prototype.slice is a VIEW, not a copy: a converter that patches a
    // "copied" record through `.buffer` writes into the wrong bytes. Found by
    // the LibreOffice oracle, which reads its files with readFileSync.
    const padded = Buffer.concat([Buffer.alloc(13, 0x41), Buffer.from(dotx())]).subarray(13);
    const fromBuffer = convertTemplate(padded, "document");
    expect(same(fromBuffer, convertTemplate(dotx(), "document"))).toBe(true);
    expect(Agent.read(fromBuffer)).not.toHaveProperty("sourceKind");
  });

  it("flips direction when no target is given", () => {
    expect((Agent.read(convertTemplate(docx())) as Record<string, unknown>).sourceKind).toBe("template");
    expect(Agent.read(convertTemplate(dotx()))).not.toHaveProperty("sourceKind");
  });

  it("converts .odt to .ott keeping mimetype first, STORED and without an extra field", () => {
    const input = odf(ODT);
    const output = convertTemplate(input, "template");
    const before = unzip(input).entries;
    const { entries: after, methods } = unzip(output);
    expect([...after.keys()]).toEqual([...before.keys()]);
    expect([...after.keys()][0]).toBe("mimetype");
    expect(text(after, "mimetype")).toBe(OTT);
    expect(methods.get("mimetype")).toBe(0);
    const view = new DataView(output.buffer, output.byteOffset, output.byteLength);
    expect(view.getUint32(0, true)).toBe(0x04034b50);
    expect(view.getUint16(28, true)).toBe(0);
    expect(dec.decode(output.subarray(30, 38))).toBe("mimetype");
    expect(dec.decode(output.subarray(38, 38 + OTT.length))).toBe(OTT);

    expect(text(after, "META-INF/manifest.xml")).toBe(text(before, "META-INF/manifest.xml").replace(
      `manifest:full-path="/" manifest:version="1.3" manifest:media-type="${ODT}"`,
      `manifest:full-path="/" manifest:version="1.3" manifest:media-type="${OTT}"`));
    for (const [name, data] of before) {
      if (name !== "mimetype" && name !== "META-INF/manifest.xml") expect(same(after.get(name), data), `${name} changed`).toBe(true);
    }
    expect((Agent.read(output) as Record<string, unknown>).sourceKind).toBe("template");
    expectSameEntries(unzip(convertTemplate(output, "document")).entries, before);
  });

  it("refuses conversions it must not make", () => {
    refused(() => convertTemplate(dotx(), "template"), "already_template");
    refused(() => convertTemplate(docx(), "document"), "already_document");
    refused(() => convertTemplate(retyped(MACRO_TEMPLATE), "document"), "macro_enabled");
    refused(() => convertTemplate(retyped(MACRO_DOC), "template"), "macro_enabled");
    refused(() => convertTemplate(docx(), "spreadsheet" as never), "bad_target");
    refused(() => convertTemplate(retyped(PPT_TEMPLATE)), "family_mismatch");
  });

  it("refuses an ODF package whose mimetype is not stored first", () => {
    const entries = unzip(odf(ODT)).entries;
    const mimetype = entries.get("mimetype")!;
    entries.delete("mimetype");
    // Re-added last; the test writer then deflates nothing special about it.
    const reordered = new Map([...entries, ["mimetype", mimetype] as const]);
    refused(() => convertTemplate(zip(reordered)), "odf_mimetype");
  });

  it("refuses a content-type declaration it cannot rewrite unambiguously", () => {
    const entries = unzip(dotx()).entries;
    entries.set("[Content_Types].xml", enc.encode(text(entries, "[Content_Types].xml").replace(
      '<Override PartName="/word/document.xml"',
      `<!-- <Override PartName="/word/document.xml" ContentType="${TEMPLATE}"/> --><Override PartName="/word/document.xml"`)));
    refused(() => convertTemplate(zip(entries)), "not_convertible");
  });

  it("names a spreadsheet instead of converting it", () => {
    const xlsx = zip(new Map([["[Content_Types].xml", enc.encode("<Types/>")], ["xl/workbook.xml", enc.encode("<workbook/>")]]));
    expect(() => convertTemplate(xlsx)).toThrow(UnsupportedFormatException);
  });
});

describe("template:convert CLI", () => {
  function cli(args: string[]): [number, string, string] {
    let out = "";
    let err = "";
    const code = run(args, { stdout: (s) => { out += s; }, stderr: (s) => { err += s; } });
    return [code, out, err];
  }

  it("converts through the template:convert CLI", () => {
    const dir = mkdtempSync(join(tmpdir(), "lwt-cli-"));
    try {
      writeFileSync(join(dir, "in.dotx"), dotx());
      let [code, out, err] = cli(["template:convert", join(dir, "in.dotx"), join(dir, "out.docx")]);
      expect(code).toBe(0);
      expect(out).toContain("document");
      expect(Agent.read(readFileSync(join(dir, "out.docx")))).not.toHaveProperty("sourceKind");

      [code, , err] = cli(["template:convert", join(dir, "out.docx"), join(dir, "out.docx"), "--to=template"]);
      expect(code).toBe(1);
      expect(err).toContain("exists");

      [code] = cli(["template:convert", join(dir, "out.docx"), join(dir, "back.dotx"), "--to=template"]);
      expect(code).toBe(0);
      expect((Agent.read(readFileSync(join(dir, "back.dotx"))) as Record<string, unknown>).sourceKind).toBe("template");

      writeFileSync(join(dir, "macro.dotm"), retyped(MACRO_TEMPLATE));
      [code, , err] = cli(["template:convert", join(dir, "macro.dotm"), join(dir, "macro.docx")]);
      expect(code).toBe(1);
      expect(err).toContain("macro_enabled");
      expect(existsSync(join(dir, "macro.docx"))).toBe(false);

      [code, , err] = cli(["template:convert"]);
      expect(code).toBe(2);
      expect(err).toContain("Usage");
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  it("ships the CLI as a bin with a shebang, built by tsup", () => {
    const pkg = JSON.parse(readFileSync(join(__dirname, "../package.json"), "utf8"));
    expect(pkg.bin).toEqual({ "last-word": "dist/cli.js" });
    expect(readFileSync(join(__dirname, "../src/cli.ts"), "utf8").startsWith("#!/usr/bin/env node\n")).toBe(true);
    expect(readFileSync(join(__dirname, "../tsup.config.ts"), "utf8")).toContain("src/cli.ts");
  });

  // The built bin, end to end, wherever `npm run build` has run first.
  it.runIf(existsSync(join(__dirname, "../dist/cli.js")))("runs the built bin", () => {
    const dir = mkdtempSync(join(tmpdir(), "lwt-bin-"));
    try {
      writeFileSync(join(dir, "in.dotx"), dotx());
      const result = spawnSync(process.execPath, [join(__dirname, "../dist/cli.js"), "template:convert", join(dir, "in.dotx"), join(dir, "out.docx")], { encoding: "utf8" });
      expect(result.status, result.stderr).toBe(0);
      expect(Agent.read(readFileSync(join(dir, "out.docx")))).not.toHaveProperty("sourceKind");
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });
});

// ─── LibreOffice as an oracle (skipped where it is not installed) ─────────

function soffice(): string | null {
  const candidates = [process.env.LIBREOFFICE_BIN, "C:/Program Files/LibreOffice/program/soffice.com",
    "/usr/bin/soffice", "/usr/local/bin/soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"];
  return candidates.find((c): c is string => typeof c === "string" && c !== "" && existsSync(c)) ?? null;
}

/** Headless convert; the profile lives in a SHORT temp path (long ones crash LO on Windows). */
function libreOffice(bin: string, input: string, format: string, outDir: string): string {
  const profile = mkdtempSync(join(tmpdir(), "lwlo"));
  const url = "file:///" + profile.split(String.fromCharCode(92)).join("/").replace(/^\/+/, "");
  const result = spawnSync(bin, [`-env:UserInstallation=${url}`, "--headless", "--norestore",
    "--convert-to", format, "--outdir", outDir, input], { encoding: "utf8", timeout: 120_000 });
  const base = input.split(/[/\\]/).pop()!.replace(/\.[^.]+$/, "");
  const out = join(outDir, `${base}.${format.split(":")[0]}`);
  expect(result.status, result.stderr).toBe(0);
  expect(existsSync(out), `LibreOffice produced no ${format}`).toBe(true);
  return out;
}

const bin = soffice();

describe.skipIf(bin === null)("LibreOffice oracle", () => {
  it("reads what LibreOffice writes as .dotx and .ott, and LibreOffice reads our conversions", () => {
    const dir = mkdtempSync(join(tmpdir(), "lwo"));
    try {
      writeFileSync(join(dir, "src.docx"), Agent.toBytes(doc()));
      for (const [template, document] of [["dotx", "docx"], ["ott", "odt"]] as const) {
        const loTemplate = libreOffice(bin!, join(dir, "src.docx"), template, dir);
        const read = Agent.read(readFileSync(loTemplate)) as Record<string, unknown>;
        expect(read.sourceKind, `LibreOffice .${template} not marked`).toBe("template");
        expect(JSON.stringify(read)).toContain("Dear reader");

        mkdirSync(join(dir, template));
        const loBack = libreOffice(bin!, loTemplate, document, join(dir, template));
        expect(Agent.read(readFileSync(loBack))).not.toHaveProperty("sourceKind");

        writeFileSync(join(dir, `ours.${document}`), convertTemplate(readFileSync(loTemplate), "document"));
        mkdirSync(join(dir, `ours-${template}`));
        const txt = libreOffice(bin!, join(dir, `ours.${document}`), "txt:Text", join(dir, `ours-${template}`));
        expect(readFileSync(txt, "utf8")).toContain("Dear reader");
      }
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  }, 180_000);
});
