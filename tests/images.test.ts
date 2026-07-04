import { describe, it, expect } from "vitest";
import { Agent, jpegSize, parseDataUrl, pngSize, resolveImageSize, sniffImageSize, unzipSync } from "../src";
import canonical from "../test/fixtures/canonical.json";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

const RED_PNG = (canonical.blocks.find((b: Any) => b.type === "image") as Any).src as string;

/** Minimal synthetic JPEG: SOI + SOF0 frame (3x2 px, 3 components) + EOI. */
const TINY_JPEG = new Uint8Array([
  0xff, 0xd8, // SOI
  0xff, 0xc0, 0x00, 0x11, 0x08, 0x00, 0x02, 0x00, 0x03, 0x03, // SOF0: h=2 w=3
  0x01, 0x22, 0x00, 0x02, 0x11, 0x01, 0x03, 0x11, 0x01,
  0xff, 0xd9, // EOI
]);

describe("image dimension sniffing (vector 7)", () => {
  it("reads PNG IHDR dimensions", () => {
    const decoded = parseDataUrl(RED_PNG)!;
    expect(decoded.mime).toBe("image/png");
    expect(pngSize(decoded.bytes)).toEqual({ width: 2, height: 2 });
    expect(sniffImageSize(decoded.bytes)).toEqual({ width: 2, height: 2 });
  });

  it("reads JPEG SOF0 dimensions", () => {
    expect(jpegSize(TINY_JPEG)).toEqual({ width: 3, height: 2 });
    expect(sniffImageSize(TINY_JPEG)).toEqual({ width: 3, height: 2 });
  });

  it("returns null for non-image bytes", () => {
    expect(sniffImageSize(new TextEncoder().encode("not an image"))).toBeNull();
  });

  it("uses sniffed dimensions when widthPx/heightPx are omitted", () => {
    const decoded = parseDataUrl(RED_PNG)!;
    expect(resolveImageSize({}, decoded.bytes)).toEqual({ width: 2, height: 2 });
  });

  it("derives the missing dimension from the intrinsic aspect ratio", () => {
    const decoded = parseDataUrl(RED_PNG)!;
    expect(resolveImageSize({ widthPx: 100 }, decoded.bytes)).toEqual({ width: 100, height: 100 });
  });

  it("caps at 6.5in width keeping aspect", () => {
    const decoded = parseDataUrl(RED_PNG)!;
    expect(resolveImageSize({ widthPx: 1248, heightPx: 624 }, decoded.bytes)).toEqual({
      width: 624,
      height: 312,
    });
  });

  it("drives wp:extent through the writer when px are absent (full path)", () => {
    const doc: Any = { blocks: [{ type: "image", src: RED_PNG }] };
    const out: Any = Agent.read(Agent.toBytes(doc));
    expect(out.blocks[0].widthPx).toBe(2);
    expect(out.blocks[0].heightPx).toBe(2);
    expect(out.blocks[0].src).toBe(RED_PNG);
  });

  it("stores decoded media bytes and rels in the archive", () => {
    const doc: Any = {
      blocks: [{ type: "image", src: RED_PNG, widthPx: 10, heightPx: 10, alt: "red" }],
    };
    const parts = unzipSync(Agent.toBytes(doc));
    expect(parts["word/media/image1.png"]).toBeDefined();
    expect(pngSize(parts["word/media/image1.png"]!)).toEqual({ width: 2, height: 2 });
    const rels = new TextDecoder().decode(parts["word/_rels/document.xml.rels"]!);
    expect(rels).toContain('Target="media/image1.png"');
  });

  it("round-trips alt text via wp:docPr descr", () => {
    const doc: Any = {
      blocks: [{ type: "image", src: RED_PNG, widthPx: 4, heightPx: 4, alt: "a red square" }],
    };
    const out: Any = Agent.read(Agent.toBytes(doc));
    expect(out.blocks[0].alt).toBe("a red square");
  });
});
