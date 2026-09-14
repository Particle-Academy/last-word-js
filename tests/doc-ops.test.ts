import { describe, it, expect } from "vitest";
import canonical from "../test/fixtures/canonical.json";
import { Agent, DocDiff, DocOpSchema, type DocOp } from "../src";
import { type Any, EDITS, deepFreeze, lwOpsDoc, p, randomBlockEdits } from "./doc-ops-fixtures";

/*
 * Agent.diff / Agent.reduce / Agent.opSchema (last-word#2), ported case for case
 * from PHP `tests/Unit/DocOpsTest.php` (last-word 0.6.3).
 *
 * What a version history built on these needs, pinned:
 *
 * 1. Round trip: reduce(a, diff(a, b)) equals b, both ways.
 * 2. Small edits stay small: rewording one paragraph is one block-level op,
 *    at whatever depth it sits — asserted as the exact ops, with their paths.
 * 3. A save without a change records nothing: diff(d, read(toBytes(d))) is [].
 *
 * That the ops are the SAME ops PHP emits is `doc-ops-parity.test.ts`.
 */

const SHAPE_KEYS = ["op", "path", "index", "from", "to", "key"];

describe("Agent.diff / Agent.reduce", () => {
  describe("reproduces the target exactly, with one op at the right path", () => {
    for (const [name, [edit, expected]] of Object.entries(EDITS)) {
      it(name, () => {
        const a = deepFreeze(lwOpsDoc());
        const b = deepFreeze(edit(lwOpsDoc()));

        const ops = Agent.diff(a, b);

        // The op, path and position, in PHP's key order; the payload is checked by the round trip.
        const shape = ops.map((op) => Object.fromEntries(Object.entries(op).filter(([key]) => SHAPE_KEYS.includes(key))));
        expect(shape).toStrictEqual(expected);
        expect(shape.map((op) => Object.keys(op))).toEqual(expected.map((op) => Object.keys(op)));
        expect(DocDiff.same(Agent.reduce(a, ops), b)).toBe(true);

        const reverse = Agent.diff(b, a);
        expect(DocDiff.same(Agent.reduce(b, reverse), a)).toBe(true);
      });
    }
  });

  it("records nothing for a save without a change", () => {
    const doc = lwOpsDoc();
    const docBack = Agent.read(Agent.toBytes(doc));
    expect(Agent.diff(doc, docBack)).toEqual([]);
    // This document's header row reads back bold, so the [] above comes from
    // `equivalent`, not from the two being the same JSON.
    expect(DocDiff.same(docBack, doc)).toBe(false);

    // The canonical fixture reads back as the same JSON, in PHP (checked on its
    // own fixture and on this one) and here, so this line pins the literal-same
    // path, and the two documents around it pin normalisation. (PHP's test said
    // otherwise until 0.6.3, when this port checked.)
    const canonicalBack = Agent.read(Agent.toBytes(canonical));
    expect(Agent.diff(canonical, canonicalBack)).toEqual([]);
    expect(DocDiff.same(canonicalBack, canonical)).toBe(true);

    // Runs the reader will merge, and a header row it will read back bold, are
    // not changes to the file.
    const normalised = {
      blocks: [
        { type: "paragraph", runs: [{ text: "Split " }, { text: "run" }] },
        { type: "table", rows: [{ header: true, cells: [{ blocks: [p("Head")] }] }] },
      ],
    };
    const readBack = Agent.read(Agent.toBytes(normalised));
    expect(DocDiff.same(readBack, normalised), "the fixture must actually exercise a normalisation").toBe(false);
    expect(Agent.diff(normalised, readBack)).toEqual([]);
    expect(Agent.equivalent(normalised, readBack)).toBe(true);
  });

  it("keeps the round trip over a seeded run of random edits, without replacing the document", () => {
    const pairs = randomBlockEdits(20260915, 60);
    expect(pairs).toHaveLength(60);

    pairs.forEach(([a, b], run) => {
      const ops = Agent.diff(a, b);

      expect(DocDiff.same(Agent.reduce(a, ops), b), `run ${run}`).toBe(true);
      expect(
        ops.map((op) => op.op),
        `run ${run} fell back`,
      ).not.toContain("doc.replace");
    });
  });

  it("skips an op whose path or index does not resolve, and never modifies its input", () => {
    const d = deepFreeze(lwOpsDoc());

    expect(Agent.reduce(d, { op: "blocks.remove", path: "/blocks", index: 99 })).toStrictEqual(d);
    expect(Agent.reduce(d, { op: "blocks.replace", path: "/nowhere/blocks", index: 0, block: p("x") })).toStrictEqual(d);
    // A path must end in the kind of list the op edits.
    expect(Agent.reduce(d, { op: "rows.remove", path: "/blocks", index: 0 })).toStrictEqual(d);
    expect(Agent.reduce(d, { op: "no.such", path: "/blocks" } as Any)).toStrictEqual(d);
    expect(d).toStrictEqual(lwOpsDoc());

    // Inserting into a children list that does not exist yet creates it.
    const withChild: Any = Agent.reduce(d, {
      op: "items.insert",
      path: "/blocks/3/items/1/children",
      index: 0,
      item: { runs: [{ text: "Retail" }] },
    });
    expect(withChild.blocks[3].items[1].children).toStrictEqual([{ runs: [{ text: "Retail" }] }]);

    // One op or a list.
    const one = Agent.reduce(d, { op: "doc.set", key: "title", value: null });
    expect(one).not.toHaveProperty("title");
    expect(Agent.reduce(d, [{ op: "doc.set", key: "title", value: "X" }]).title).toBe("X");
  });

  it("publishes one schema variant per op, and diff only emits those", () => {
    const schema = Agent.opSchema() as { oneOf: Any[] };

    expect(schema).toStrictEqual(DocOpSchema.jsonSchema());
    expect(schema.oneOf.map((v) => v.properties.op.const)).toStrictEqual([...DocOpSchema.TYPES]);

    // The `DocOp` union and TYPES are held to each other in `doc-op-schema.ts`, where `tsc` sees them.
    const names: DocOp["op"][] = [...DocOpSchema.TYPES];
    expect(new Set(names).size).toBe(18);
  });

  it("aligns lists by content, breaking ties toward deleting first", () => {
    expect(DocDiff.hunks(["a", "b", "c"], ["a", "x", "b", "c"])).toStrictEqual([[1, 0, 1, 1]]);
    expect(DocDiff.hunks(["a", "b", "c"], ["a", "c"])).toStrictEqual([[1, 1, 0, 1]]);
    expect(DocDiff.hunks(["a", "b"], ["a", "z"])).toStrictEqual([[1, 1, 1, 1]]);
  });

  // PHP 0.6.1. PHP's case is two different invalid UTF-8 bytes, which both
  // encoded to "" and compared equal; JS strings cannot hold bytes, so the cases
  // here are the values this port's canonical form would otherwise collapse.
  it("refuses to compare values JSON cannot hold, instead of calling them the same", () => {
    // JSON.stringify writes every one of these as null.
    expect(() => DocDiff.same(Number.NaN, Number.POSITIVE_INFINITY)).toThrow(TypeError);
    expect(() => DocDiff.same({ width: Number.NaN }, { width: null })).toThrow(TypeError);
    // Text that is not valid Unicode, the JS form of PHP's invalid UTF-8.
    expect(() => DocDiff.same("\uD800", "\uDC01")).toThrow(TypeError);
    expect(() => DocDiff.same(p("fine"), p("\uDBFF broken"))).toThrow(TypeError);
    expect(() => Agent.diff(lwOpsDoc(), { ...lwOpsDoc(), blocks: [p("\uD83D")] })).toThrow(TypeError);
    // A surrogate PAIR is valid Unicode.
    expect(DocDiff.same(p("\u{1F600}"), p("\u{1F600}"))).toBe(true);

    // PHP's depth argument: 4096 nested arrays encode, 4097 throw, and an empty
    // one counts. A cycle ends up there too rather than overflowing the stack.
    const nest = (n: number, leaf: unknown): unknown => (n === 0 ? leaf : nest(n - 1, [leaf]));
    expect(DocDiff.same(nest(4096, 1), nest(4096, 1))).toBe(true);
    expect(() => DocDiff.same(nest(4097, 1), 1)).toThrow(TypeError);
    expect(() => DocDiff.same(nest(4096, []), 1)).toThrow(TypeError);
    const cycle: Any = { blocks: [] };
    cycle.blocks.push(cycle);
    expect(() => DocDiff.same(cycle, cycle)).toThrow(TypeError);

    // undefined is JSON's absent, not a value: nothing to refuse.
    expect(DocDiff.same({ a: 1, b: undefined }, { a: 1 })).toBe(true);
  });

  // PHP 0.6.2.
  it('skips an op whose position or key is not one, instead of casting it to 0 or "1"', () => {
    const d = deepFreeze(lwOpsDoc());

    // `(int) "abc"` is 0: these edited the FIRST block.
    expect(Agent.reduce(d, { op: "blocks.remove", path: "/blocks", index: "abc" } as Any)).toStrictEqual(d);
    expect(Agent.reduce(d, { op: "blocks.replace", path: "/blocks", index: "x", block: p("x") } as Any)).toStrictEqual(d);
    expect(Agent.reduce(d, { op: "blocks.move", path: "/blocks", from: "first", to: 3 } as Any)).toStrictEqual(d);
    expect(Agent.reduce(d, { op: "blocks.insert", path: "/blocks", index: true, block: p("x") } as Any)).toStrictEqual(d);

    // `(string) true` is "1": this set a top-level key "1".
    expect(Agent.reduce(d, { op: "doc.set", key: true, value: "x" } as Any)).toStrictEqual(d);

    // An op name or path that is not a string is no op, not "Array".
    expect(Agent.reduce(d, { op: ["blocks.remove"], path: "/blocks", index: 0 } as Any)).toStrictEqual(d);
    expect(Agent.reduce(d, { op: "blocks.remove", path: ["/blocks"], index: 0 } as Any)).toStrictEqual(d);

    // Digit strings and ints still work.
    expect((Agent.reduce(d, { op: "blocks.remove", path: "/blocks", index: "1" } as Any).blocks as Any[]).length).toBe(d.blocks.length - 1);
  });

  // PHP 0.6.3. PHP stores the key "5" as the int 5 and emitted `key: 5`, which
  // the reducer refuses, so the diff fell back to doc.replace. A JS object key
  // is always a string, but 0.6.0 re-created PHP's int to match it.
  it("keeps a small diff for a numeric top-level key, which PHP stores as an int", () => {
    // Unknown top-level keys are not written, so another difference (the title)
    // is what keeps these two documents from being the same file.
    const a = deepFreeze({ title: "A", blocks: [p("x")], "5": "five" });
    const b = deepFreeze({ title: "B", blocks: [p("x")], "5": "FIVE" });

    const ops = Agent.diff(a, b);

    expect(ops).toStrictEqual([
      { op: "doc.set", key: "5", value: "FIVE" },
      { op: "doc.set", key: "title", value: "B" },
    ]);
    expect(DocDiff.same(Agent.reduce(a, ops), b)).toBe(true);
  });

  // PHP 0.6.3: this insert turned the top-level blocks list into a map.
  it("skips an op whose path names a list by a key instead of a position", () => {
    const d = deepFreeze(lwOpsDoc());

    expect(Agent.reduce(d, { op: "blocks.insert", path: "/blocks/blocks", index: 0, block: p("x") })).toStrictEqual(d);
  });

  // PHP 0.6.3.
  it("refuses an empty doc.set key in the op schema, as the reducer does", () => {
    const variant = (Agent.opSchema() as { oneOf: Any[] }).oneOf.find((v) => v.properties.op.const === "doc.set");

    expect(variant.properties.key.minLength).toBe(1);
  });
});
