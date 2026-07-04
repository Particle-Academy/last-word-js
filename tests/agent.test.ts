import { describe, it, expect } from "vitest";
import { existsSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Agent } from "../src";
import canonical from "../test/fixtures/canonical.json";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

describe("describe (vector 6)", () => {
  it("includes title, per-type counts, and word count", () => {
    const summary = Agent.describe(canonical);
    expect(summary).toContain("Doc: LastWord Canonical");
    expect(summary).toContain("Blocks: 13");
    expect(summary).toContain("3 heading");
    expect(summary).toContain("2 paragraph");
    expect(summary).toContain("2 list");
    expect(summary).toContain("1 table");
    expect(summary).toContain("1 code");
    expect(summary).toContain("1 quote");
    expect(summary).toContain("1 image");
    expect(summary).toContain("1 pageBreak");
    expect(summary).toContain("1 hr");
    expect(summary).toMatch(/Words: \d+/);
  });

  it("counts words across headings, paragraphs, lists, tables, and quotes", () => {
    const summary = Agent.describe({
      title: "T",
      blocks: [
        { type: "heading", level: 1, runs: [{ text: "one two" }] },
        { type: "quote", blocks: [{ type: "paragraph", runs: [{ text: "three" }] }] },
        { type: "list", items: [{ runs: [{ text: "four" }], children: [{ runs: [{ text: "five" }] }] }] },
        { type: "code", text: "not counted at all" },
      ],
    });
    expect(summary).toContain("Words: 5");
  });

  it("defaults the title to Untitled", () => {
    expect(Agent.describe({ blocks: [] })).toContain("Doc: Untitled");
  });
});

describe("write (Node only)", () => {
  it("writes a .docx to disk and reports path/bytes/blocks", async () => {
    const path = join(tmpdir(), `last-word-test-${process.pid}.docx`);
    try {
      const result = await Agent.write(canonical, path);
      expect(result.path).toBe(path);
      expect(result.blocks).toBe(13);
      expect(result.bytes).toBeGreaterThan(0);
      expect(existsSync(path)).toBe(true);

      const onDisk = readFileSync(path);
      expect(onDisk.length).toBe(result.bytes);
      const doc: Any = Agent.read(new Uint8Array(onDisk));
      expect(doc.title).toBe("LastWord Canonical");
    } finally {
      rmSync(path, { force: true });
    }
  });
});

describe("read aliases", () => {
  it("fromBytes is an alias for read and accepts ArrayBuffer", () => {
    const bytes = Agent.toBytes(canonical);
    const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    expect(Agent.fromBytes(buffer as ArrayBuffer)).toEqual(Agent.read(bytes));
  });
});
