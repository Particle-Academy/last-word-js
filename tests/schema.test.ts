import { describe, it, expect } from "vitest";
import { Agent, SchemaException, Schema, Validator } from "../src";
import canonical from "../test/fixtures/canonical.json";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

describe("validate (vector 4)", () => {
  it("accepts the canonical doc", () => {
    expect(Agent.validate(canonical)).toEqual([]);
  });

  it("errors on missing blocks", () => {
    const errors = Agent.validate({ title: "No blocks" });
    expect(errors.length).toBeGreaterThan(0);
    expect(errors[0]).toEqual({ path: "/blocks", message: expect.stringContaining("blocks") });
  });

  it("errors on heading level 9", () => {
    const errors = Agent.validate({ blocks: [{ type: "heading", level: 9, runs: [{ text: "x" }] }] });
    expect(errors.some((e) => e.path === "/blocks/0/level")).toBe(true);
  });

  it("errors on unknown block types and bad runs", () => {
    const errors = Agent.validate({
      blocks: [{ type: "banana" }, { type: "paragraph", runs: [{ text: 42 }] }],
    });
    expect(errors.some((e) => e.path === "/blocks/0/type")).toBe(true);
    expect(errors.some((e) => e.path === "/blocks/1/runs/0/text")).toBe(true);
  });

  it("errors on non-hex colors and non-data-URL images", () => {
    const errors = Agent.validate({
      blocks: [
        { type: "paragraph", runs: [{ text: "x", color: "red" }] },
        { type: "image", src: "https://example.com/a.png" },
      ],
    });
    expect(errors.some((e) => e.path === "/blocks/0/runs/0/color")).toBe(true);
    expect(errors.some((e) => e.path === "/blocks/1/src")).toBe(true);
  });
});

describe("validateAndRepair (vector 4)", () => {
  it("clamps heading level 9 to 6 and coerces string runs", () => {
    const result: Any = Agent.validateAndRepair({
      blocks: [{ type: "heading", level: 9, runs: "Big heading" }],
    });
    expect(result.ok).toBe(true);
    expect(result.schema.blocks[0]).toEqual({
      type: "heading",
      level: 6,
      runs: [{ text: "Big heading" }],
    });
  });

  it("coerces bare string blocks to paragraphs", () => {
    const result: Any = Agent.validateAndRepair({ blocks: ["just text"] });
    expect(result.ok).toBe(true);
    expect(result.schema.blocks[0]).toEqual({ type: "paragraph", runs: [{ text: "just text" }] });
  });

  it("defaults missing blocks to []", () => {
    const result: Any = Agent.validateAndRepair({ title: "Empty" });
    expect(result.ok).toBe(true);
    expect(result.schema.blocks).toEqual([]);
  });

  it("drops unknown block types with the error retained", () => {
    const result: Any = Agent.validateAndRepair({
      blocks: [{ type: "paragraph", runs: [{ text: "keep" }] }, { type: "banana", stuff: 1 }],
    });
    expect(result.ok).toBe(true);
    expect(result.schema.blocks).toHaveLength(1);
    expect(result.errors.some((e: Any) => e.path === "/blocks/1/type" && /banana/.test(e.message))).toBe(true);
  });

  it("returns the doc untouched when already valid", () => {
    const result = Agent.validateAndRepair(canonical);
    expect(result.ok).toBe(true);
    expect(result.errors).toEqual([]);
    expect(result.schema).toBe(canonical);
  });

  it("never mutates its input", () => {
    const input: Any = { blocks: [{ type: "heading", level: 9, runs: "x" }] };
    Agent.validateAndRepair(input);
    expect(input.blocks[0].level).toBe(9);
    expect(input.blocks[0].runs).toBe("x");
  });
});

describe("toBytes gating", () => {
  it("throws SchemaException with structured errors for invalid docs", () => {
    try {
      Agent.toBytes({ nope: true });
      expect.unreachable("should have thrown");
    } catch (e) {
      expect(e).toBeInstanceOf(SchemaException);
      expect((e as SchemaException).errors[0]!.path).toBe("/blocks");
    }
  });
});

describe("jsonSchema / version", () => {
  it("exports a JSON Schema for LLM tool registration", () => {
    const schema: Any = Agent.jsonSchema();
    expect(schema.title).toBe("LastWord Doc");
    expect(schema.required).toEqual(["blocks"]);
    expect(schema.definitions.block.properties.type.enum).toEqual([...Schema.BLOCK_TYPES]);
  });

  it("reports its version", () => {
    expect(Agent.version()).toBe("0.1.0");
  });

  it("exposes the Validator service directly", () => {
    expect(new Validator().validate(canonical)).toEqual([]);
  });
});
