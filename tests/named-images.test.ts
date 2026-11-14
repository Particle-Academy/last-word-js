import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { Agent, ImageRefused, SchemaException, unzipSync } from "../src";

/**
 * Named images (CC-fancy-0039): the HOST registers images with the writer --
 * `images: {logo: <bytes or data URL>, ...}`, with an optional size and alt each --
 * and the model places one with {"type": "image", "src": "image:logo"}, resolved at
 * write time, in the body and in header and footer parts alike, so a letterhead can
 * carry the logo without the model ever holding its bytes.
 *
 * test/fixtures/named-images.json is the PHP package's tests/fixtures/named-images.json,
 * byte for byte: the same images, documents, guide text, schema text and refusal
 * codes, and the sha256 of every part each document writes, taken from the PHP
 * engine (the reference). Asserting the same hashes here proves the two writers
 * agree part for part without PHP installed; parity.test.ts diffs against a live
 * PHP as well where one is on PATH.
 */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

const FIXTURE_SHA256 = "fdeebc4e1fca2c53d89a54bf4c98fbe95d1656cabc354d2a03803b44428679c6";
const FIXTURE_PATH = join(__dirname, "..", "test", "fixtures", "named-images.json");
const fixture: Any = JSON.parse(readFileSync(FIXTURE_PATH, "utf8"));
const options = (): Any => ({ images: fixture.images });
const CASES = ["body", "letterhead", "letterheadFirstPage", "footnoteAndFooter"];

const sha256 = (bytes: Uint8Array | string): string => createHash("sha256").update(bytes).digest("hex");
const dec = new TextDecoder();
const text = (parts: Record<string, Uint8Array>, name: string): string => dec.decode(parts[name]!);
const fromBase64 = (dataUrl: string): Uint8Array => new Uint8Array(Buffer.from(dataUrl.slice(dataUrl.indexOf(",") + 1), "base64"));
const PARAGRAPH = { blocks: [{ type: "paragraph", runs: [{ text: "x" }] }] };

function refusal(opts: Any): string | null {
  try {
    Agent.toBytes(PARAGRAPH, opts);
  } catch (e) {
    if (e instanceof ImageRefused) return e.reason;
    throw e;
  }
  return null;
}

describe("named images", () => {
  it("reads the shared fixture, and it is the shared one", () => {
    expect(sha256(readFileSync(FIXTURE_PATH, "utf8").replace(/\r\n/g, "\n"))).toBe(FIXTURE_SHA256);
  });

  it.each(CASES)("writes every part of %s exactly as the PHP engine does (the shared hashes)", (name) => {
    const c = fixture.cases[name];
    const parts = unzipSync(Agent.toBytes(c.doc, options()));
    expect(Object.keys(parts).sort()).toEqual(c.partNames);
    for (const [part, hash] of Object.entries(c.parts as Record<string, string>)) {
      expect(sha256(parts[part]!), `${name}: ${part}`).toBe(hash);
    }
  });

  it("resolves a name to exactly what writing its data URL by hand writes", () => {
    const doc: Any = structuredClone(fixture.cases.letterhead.doc);
    const byHand: Any = structuredClone(doc);
    byHand.page.header.blocks[0] = { type: "image", src: fixture.images.logo.data, widthPx: 150, heightPx: 50, alt: "Acme Ltd" };
    byHand.blocks[5] = { type: "image", src: fixture.images.signature.data, widthPx: 120, alt: "Signature of J. Smith" };

    expect(Agent.toBytes(doc, options())).toEqual(Agent.toBytes(byHand));
    // And the document handed in is not modified.
    expect(doc.page.header.blocks[0].src).toBe("image:logo");
  });

  it("puts the logo in the header part, through the header part's own relationships", () => {
    const parts = unzipSync(Agent.toBytes(fixture.cases.letterhead.doc, options()));
    const header = text(parts, "word/header1.xml");

    expect(header).toContain("<w:drawing>");
    expect(header).toContain('<wp:extent cx="1428750" cy="476250"/>');
    expect(header).toContain('descr="Acme Ltd"');
    expect(header).toContain('r:embed="rId1"');
    expect(header).toContain("hello@example.com");
    expect(header).toContain('<w:tag w:val="lastword:band"/>');
    // The body renders before the bands, as PHP's does: the signature is image1.
    expect(text(parts, "word/_rels/header1.xml.rels")).toContain('Target="media/image2.png"');
    expect(parts["word/media/image2.png"]).toEqual(fromBase64(fixture.images.logo.data));
    expect(text(parts, "word/footer1.xml")).toContain(" PAGE ");
    expect(text(parts, "word/_rels/document.xml.rels")).toContain('Target="media/image1.png"');
  });

  it("lets a block set its own size and alt over the registered ones", () => {
    const xml = text(unzipSync(Agent.toBytes(fixture.cases.body.doc, options())), "word/document.xml");

    expect(xml).toContain('<wp:extent cx="609600" cy="457200"/>');
    expect(xml).toContain('<wp:extent cx="2857500" cy="952500"/>');
    expect(xml).toContain('descr="Acme Ltd, large"');
    expect(xml).toContain('descr="The same seal, as a data URL"');
  });

  it.each(["doc", "bandDoc"])("refuses an unknown name with image_unknown: %s", (key) => {
    const doc = fixture.unknown[key];
    expect(Agent.validate(doc, options()).map((e) => e.code)).toEqual([fixture.unknown.code]);
    try {
      Agent.toBytes(doc, options());
      expect.unreachable("an unknown image name was written");
    } catch (e) {
      expect(e).toBeInstanceOf(SchemaException);
      const errors = (e as SchemaException).errors;
      expect(errors.map((x) => x.code)).toEqual(["image_unknown"]);
      expect(errors[0]!.message).toContain("image:logo");
    }
  });

  it("refuses an image name when no images are registered at all", () => {
    const errors = Agent.validate({ blocks: [{ type: "image", src: "image:logo" }] });
    expect(errors).toHaveLength(1);
    expect(errors[0]!.code).toBe("image_unknown");
    expect(errors[0]!.path).toBe("/blocks/0/src");
  });

  it("validates and repairs a document that uses names, leaving the names in it", () => {
    const doc: Any = structuredClone(fixture.cases.letterhead.doc);
    doc.blocks.push({ type: "heading", level: 9, runs: [{ text: "Clamped" }] });

    const result: Any = Agent.validateAndRepair(doc, options());
    expect(result.ok).toBe(true);
    expect(result.schema.page.header.blocks[0].src).toBe("image:logo");

    const unknown = Agent.validateAndRepair(fixture.unknown.doc, options());
    expect(unknown.ok).toBe(false);
    expect(unknown.errors.map((e) => e.code)).toContain("image_unknown");
  });

  it.each((fixture.refusals as Any[]).map((r) => [r.title, r]))("refuses a bad registration with its stable reason: %s", (_title, row: Any) => {
    const opts: Any = { images: row.images };
    if (row.imageLimits !== undefined) opts.imageLimits = row.imageLimits;
    expect(refusal(opts)).toBe(row.reason);
  });

  it.each((fixture.accepted as Any[]).map((r) => [r.title, r]))("takes a registration at its limits: %s", (_title, row: Any) => {
    const opts: Any = { images: row.images };
    if (row.imageLimits !== undefined) opts.imageLimits = row.imageLimits;
    expect(refusal(opts)).toBeNull();
  });

  it("takes raw PNG and JPEG bytes as well as data URLs", () => {
    const doc = { blocks: [{ type: "image", src: "image:logo" }, { type: "image", src: "image:seal" }] };
    const raw = {
      images: {
        logo: { data: fromBase64(fixture.images.logo.data), widthPx: 150, heightPx: 50, alt: "Acme Ltd" },
        seal: fromBase64(fixture.images.seal).buffer.slice(0) as ArrayBuffer,
      },
    };
    expect(Agent.toBytes(doc, raw)).toEqual(Agent.toBytes(doc, options()));
  });

  it("has secure default limits: 5 MiB, 4096 px a side, 32 images", () => {
    const png = (w: number, h: number, extra = 0): Uint8Array => {
      const out = new Uint8Array(29 + extra);
      out.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0x0d, 0x49, 0x48, 0x44, 0x52]);
      new DataView(out.buffer).setUint32(16, w);
      new DataView(out.buffer).setUint32(20, h);
      out.set([8, 2, 0, 0, 0], 24);
      return out;
    };
    const fine = png(4096, 4096);
    const many: Record<string, Uint8Array> = {};
    for (let i = 0; i < 33; i++) many[`i${i}`] = fine;

    expect(refusal({ images: { wide: png(4097, 10) } })).toBe("image_dimensions_too_large");
    expect(refusal({ images: { fine } })).toBeNull();
    expect(refusal({ images: { big: png(4096, 4096, 5 * 1024 * 1024) } })).toBe("image_too_large");
    expect(refusal({ images: many })).toBe("image_count_exceeded");
    expect(refusal({ images: { wide: png(4097, 10) }, imageLimits: { maxWidthPx: 5000 } })).toBeNull();
    // An IHDR claiming 2^31 px or more is too large, as PHP reads it, not unreadable.
    expect(refusal({ images: { huge: png(0x80000000, 10) } })).toBe("image_dimensions_too_large");
  });

  it("tells the model which names exist, in the guide, and says nothing new without them", () => {
    expect(Agent.guide(options())).toBe(Agent.guide() + fixture.guideSection);
    expect(Agent.guide({})).toBe(Agent.guide());
    expect(Agent.guide({ images: {} })).toBe(Agent.guide());
  });

  it("lists the names in the schema, and leaves the schema alone without them", () => {
    const src: Any = ((Agent.jsonSchema(options()) as Any).definitions.block.properties.src);
    expect(src.description).toBe(fixture.schemaSrcDescription);
    const re = new RegExp(src.pattern);
    expect(re.test("image:logo")).toBe(true);
    expect(re.test("image:other")).toBe(false);
    expect(re.test("data:image/png;base64,AAAA")).toBe(true);
    expect(Agent.jsonSchema({})).toEqual(Agent.jsonSchema());
    expect((Agent.jsonSchema() as Any).definitions.block.properties.src).toEqual({ type: "string" });
  });

  it.each(CASES)("writes %s with every part typed and every relationship resolving", (name) => {
    const parts = unzipSync(Agent.toBytes(fixture.cases[name].doc, options()));
    const types = text(parts, "[Content_Types].xml");
    const defaults = new Set([...types.matchAll(/<Default Extension="([^"]+)"/g)].map((m) => m[1]!.toLowerCase()));
    const overrides = new Set([...types.matchAll(/<Override PartName="\/([^"]+)"/g)].map((m) => m[1]!));
    for (const part of Object.keys(parts)) {
      const ext = part.slice(part.lastIndexOf(".") + 1).toLowerCase();
      expect(overrides.has(part) || defaults.has(ext), `${part} has no content type`).toBe(true);
      if (!part.endsWith(".rels")) continue;
      const base = part.startsWith("_rels/") ? "" : part.slice(0, part.indexOf("_rels/"));
      for (const m of text(parts, part).matchAll(/<Relationship [^>]*Target="([^"]+)"(?! TargetMode="External")/g)) {
        expect(parts, `${part} -> ${m[1]}`).toHaveProperty([base + m[1]!]);
      }
    }
    for (const part of Object.keys(parts)) {
      if (!part.endsWith(".xml")) continue;
      for (const m of text(parts, part).matchAll(/r:embed="([^"]+)"/g)) {
        const slash = part.lastIndexOf("/");
        const rels = `${part.slice(0, slash)}/_rels/${part.slice(slash + 1)}.rels`;
        expect(text(parts, rels), `${part} embeds ${m[1]}`).toContain(`Id="${m[1]}"`);
      }
    }
  });
});

// ─── LibreOffice as an oracle (skipped where it is not installed) ─────────

function soffice(): string | null {
  const candidates = [process.env.LIBREOFFICE_BIN, "C:/Program Files/LibreOffice/program/soffice.com",
    "/usr/bin/soffice", "/usr/local/bin/soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"];
  return candidates.find((c): c is string => typeof c === "string" && c !== "" && existsSync(c)) ?? null;
}

const bin = soffice();

describe.skipIf(bin === null)("LibreOffice opens a named-image letterhead written by this engine", () => {
  it("has the logo in the header, the page field in the footer and the signature in the body", () => {
    const dir = mkdtempSync(join(tmpdir(), "lwni"));
    const profile = mkdtempSync(join(tmpdir(), "lwlo"));
    try {
      writeFileSync(join(dir, "letter.docx"), Agent.toBytes(fixture.cases.letterhead.doc, options()));
      const url = "file:///" + profile.split(String.fromCharCode(92)).join("/").replace(/^\/+/, "");
      const result = spawnSync(bin!, [`-env:UserInstallation=${url}`, "--headless", "--norestore",
        "--convert-to", "odt", "--outdir", dir, join(dir, "letter.docx")], { encoding: "utf8", timeout: 120_000 });
      expect(result.status, result.stderr).toBe(0);

      const odt = unzipSync(new Uint8Array(readFileSync(join(dir, "letter.odt"))));
      const styles = text(odt, "styles.xml");
      const content = text(odt, "content.xml");
      expect(styles).toMatch(/<style:header>[\s\S]*?<draw:image [\s\S]*?<\/style:header>/);
      expect(styles).toMatch(/<style:header>[\s\S]*?Acme Ltd[\s\S]*?<\/style:header>/);
      expect(styles).toMatch(/<style:footer>[\s\S]*?<text:page-number[\s\S]*?<\/style:footer>/);
      expect(content).toContain("<draw:image ");
      expect(content).toContain("Dear Ms Rivera,");
      const logo = sha256(fromBase64(fixture.images.logo.data));
      expect(Object.keys(odt).filter((p) => p.startsWith("Pictures/")).map((p) => sha256(odt[p]!))).toContain(logo);
    } finally {
      rmSync(dir, { recursive: true, force: true });
      rmSync(profile, { recursive: true, force: true });
    }
  }, 180_000);
});
