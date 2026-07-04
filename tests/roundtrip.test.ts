import { describe, it, expect } from "vitest";
import { Agent, unzipSync } from "../src";
import canonical from "../test/fixtures/canonical.json";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

const DECODER = new TextDecoder();

describe("round-trip (toBytes → read)", () => {
  it("produces a valid minimal docx container", () => {
    const bytes = Agent.toBytes(canonical);
    const parts = unzipSync(bytes);
    expect(Object.keys(parts)).toEqual(
      expect.arrayContaining([
        "[Content_Types].xml",
        "_rels/.rels",
        "docProps/core.xml",
        "word/document.xml",
        "word/styles.xml",
        "word/numbering.xml",
        "word/_rels/document.xml.rels",
        "word/media/image1.png",
      ]),
    );
    const contentTypes = DECODER.decode(parts["[Content_Types].xml"]!);
    expect(contentTypes).toContain("wordprocessingml.document.main+xml");
    expect(contentTypes).toContain('Extension="png"');
  });

  it("recovers the canonical doc semantically (deep-equal)", () => {
    const out = Agent.read(Agent.toBytes(canonical));
    expect(out).toEqual(canonical);
  });

  it("recovers the title from docProps/core.xml", () => {
    const out = Agent.read(Agent.toBytes(canonical));
    expect(out.title).toBe("LastWord Canonical");
  });

  it("round-trips heading levels 1-6", () => {
    const doc = {
      blocks: [1, 2, 3, 4, 5, 6].map((level) => ({
        type: "heading",
        level,
        runs: [{ text: `H${level}` }],
      })),
    };
    const out = Agent.read(Agent.toBytes(doc));
    expect(out.blocks).toEqual(doc.blocks);
  });

  it("round-trips a 3-deep nested list with ilvl nesting", () => {
    const doc = {
      blocks: [
        {
          type: "list",
          items: [
            {
              runs: [{ text: "a" }],
              children: [{ runs: [{ text: "b" }], children: [{ runs: [{ text: "c" }] }] }],
            },
          ],
        },
      ],
    };
    const out: Any = Agent.read(Agent.toBytes(doc));
    expect(out.blocks).toEqual(doc.blocks);
  });

  it("round-trips paragraph alignment (center/right/justify)", () => {
    const doc = {
      blocks: [
        { type: "paragraph", align: "center", runs: [{ text: "c" }] },
        { type: "paragraph", align: "right", runs: [{ text: "r" }] },
        { type: "paragraph", align: "justify", runs: [{ text: "j" }] },
        { type: "paragraph", runs: [{ text: "default" }] },
      ],
    };
    const out = Agent.read(Agent.toBytes(doc));
    expect(out.blocks).toEqual(doc.blocks);
  });

  it("normalizes adjacent same-styled runs into one (run-merge normalization)", () => {
    const doc = {
      blocks: [
        {
          type: "paragraph",
          runs: [{ text: "Hello " }, { text: "world" }, { text: "!", bold: true }],
        },
      ],
    };
    const out: Any = Agent.read(Agent.toBytes(doc));
    expect(out.blocks[0].runs).toEqual([{ text: "Hello world" }, { text: "!", bold: true }]);
  });

  it("emits deterministic bytes (vector 8)", () => {
    const a = Agent.toBytes(canonical);
    const b = Agent.toBytes(canonical);
    expect(a.length).toBe(b.length);
    expect(Buffer.from(a).equals(Buffer.from(b))).toBe(true);
  });
});
