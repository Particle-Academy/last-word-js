import { describe, it, expect, beforeAll } from "vitest";
import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import canonical from "../test/fixtures/canonical.json";
import { Agent, DocDiff, DocOpSchema } from "../src";
import { type Any, EDITS, lwOpsDoc, p, randomBlockEdits, seeded } from "./doc-ops-fixtures";

// Cross-engine OPS parity: for the same two documents, the PHP last-word's
// Agent::diff (scripts/php-diff.php) and this port's Agent.diff must return the
// SAME ops in the SAME order, and replaying them must give the same document. A
// version history written by one runtime is replayed by the other, so "both
// round-trip" is not enough: the stored op lists have to be interchangeable.
//
// PHP has one array type for lists and maps, so its JSON writes an empty map as
// `[]` and a map keyed 0..n-1 as a list. normalize() folds both sides into that
// model before comparing; nothing else is relaxed.
//
// Skips when `php` isn't on PATH LOCALLY; in CI a missing php THROWS instead,
// because a skip there is a green build with zero cross-engine coverage.

const PHP_SCRIPT = join(__dirname, "..", "scripts", "php-diff.php");

/** `php` may only resolve through the shell (Herd shims etc.). */
function php(args: string[], opts: Parameters<typeof execFileSync>[2] = {}): Buffer {
  return execFileSync("php", args, { shell: true, maxBuffer: 256 * 1024 * 1024, ...opts }) as Buffer;
}

function phpAvailable(): boolean {
  try {
    php(["--version"], { stdio: "ignore" });
    return true;
  } catch {
    return false;
  }
}

function normalize(value: Any): Any {
  if (Array.isArray(value)) return value.map(normalize);
  if (typeof value === "object" && value !== null) {
    const keys = Object.keys(value);
    if (keys.every((key, i) => key === String(i))) return keys.map((key) => normalize(value[key]));
    const out: Record<string, unknown> = {};
    for (const key of keys.sort()) out[key] = normalize(value[key]);
    return out;
  }
  return value;
}

type Case = { name: string; a?: Any; b?: Any; ops?: Any; hunks?: [string[], string[]] };

const cases: Case[] = [];

// Every edit PHP's DocOpsTest diffs, both ways.
for (const [name, [edit]] of Object.entries(EDITS)) {
  cases.push({ name: `edit: ${name}`, a: lwOpsDoc(), b: edit(lwOpsDoc()) });
  cases.push({ name: `edit reversed: ${name}`, a: edit(lwOpsDoc()), b: lwOpsDoc() });
}

// PHP's seeded top-level edits (this port's generator), both ways.
randomBlockEdits(20260915, 60).forEach(([a, b], run) => {
  cases.push({ name: `random blocks ${run}`, a, b });
  cases.push({ name: `random blocks reversed ${run}`, a: b, b: a });
});

// Seeded edits at every depth, on documents built from a tiny vocabulary so
// items repeat and the alignment has ties to break and moves to find: this is
// where two correct implementations can still disagree on WHICH item moved.
const VOCAB = ["red", "green", "blue"];

function randomDoc(rand: (min: number, max: number) => number): Any {
  const word = () => VOCAB[rand(0, 2)]!;
  const item = (depth: number): Any => {
    const out: Any = { runs: [{ text: word() }] };
    if (depth < 2 && rand(0, 2) === 0) out.children = Array.from({ length: rand(1, 3) }, () => item(depth + 1));
    return out;
  };
  const cell = (): Any => ({ blocks: Array.from({ length: rand(1, 2) }, () => p(word())) });
  const block = (): Any => {
    switch (rand(0, 5)) {
      case 0:
        return { type: "list", ...(rand(0, 1) ? { ordered: true } : {}), items: Array.from({ length: rand(1, 4) }, () => item(0)) };
      case 1:
        return { type: "table", rows: Array.from({ length: rand(1, 3) }, (_, r) => ({ ...(r === 0 && rand(0, 1) ? { header: true } : {}), cells: Array.from({ length: rand(1, 3) }, cell) })) };
      case 2:
        return { type: "quote", blocks: Array.from({ length: rand(1, 4) }, () => p(word())) };
      case 3:
        return { type: "heading", level: rand(1, 3), runs: [{ text: word() }] };
      default:
        return p(word());
    }
  };
  return { title: "Random", blocks: Array.from({ length: rand(3, 8) }, block) };
}

/** Insert, remove, move or reword one entry of `list`, keeping at least one. */
function editList(rand: (min: number, max: number) => number, list: Any[], make: () => Any, reword: (x: Any) => Any): void {
  const count = list.length;
  switch (rand(0, 3)) {
    case 0:
      list.splice(rand(0, count), 0, make());
      break;
    case 1:
      if (count > 1) list.splice(rand(0, count - 1), 1);
      break;
    case 2: {
      const [moved] = list.splice(rand(0, count - 1), 1);
      list.splice(rand(0, count - 1), 0, moved);
      break;
    }
    default: {
      const at = rand(0, count - 1);
      list[at] = reword(list[at]);
    }
  }
}

function randomNestedEdit(rand: (min: number, max: number) => number, doc: Any): void {
  const word = () => VOCAB[rand(0, 2)]!;
  const blocks: Any[] = doc.blocks;
  const ofType = (type: string) => blocks.filter((b) => b.type === type);
  const pick = <T>(list: T[]): T | undefined => (list.length === 0 ? undefined : list[rand(0, list.length - 1)]);

  switch (rand(0, 5)) {
    case 0:
      editList(rand, blocks, () => p(word()), () => p(`${word()} ${rand(1, 9)}`));
      return;
    case 1: {
      const list = pick(ofType("list"));
      if (!list) return;
      let items: Any[] = list.items;
      // Sometimes descend into an item's existing children.
      const parent = pick(items.filter((i) => Array.isArray(i.children)));
      if (parent && rand(0, 1)) items = parent.children;
      editList(rand, items, () => ({ runs: [{ text: word() }] }), (i) => ({ ...i, runs: [{ text: `${word()} ${rand(1, 9)}` }] }));
      return;
    }
    case 2: {
      const table = pick(ofType("table"));
      if (!table) return;
      const row = pick(table.rows as Any[])!;
      switch (rand(0, 2)) {
        case 0:
          editList(rand, table.rows, () => ({ cells: [{ blocks: [p(word())] }] }), (r) => ({ ...r, cells: [...r.cells, { blocks: [p(word())] }] }));
          return;
        case 1:
          editList(rand, row.cells, () => ({ blocks: [p(word())] }), () => ({ blocks: [p(`${word()} ${rand(1, 9)}`)] }));
          return;
        default: {
          const cell = pick(row.cells as Any[])!;
          editList(rand, cell.blocks, () => p(word()), () => p(`${word()} ${rand(1, 9)}`));
          return;
        }
      }
    }
    case 3: {
      const quote = pick(ofType("quote"));
      if (!quote) return;
      editList(rand, quote.blocks, () => p(word()), () => p(`${word()} ${rand(1, 9)}`));
      return;
    }
    case 4:
      switch (rand(0, 3)) {
        case 0:
          doc.title = `T${rand(1, 9)}`;
          return;
        case 1:
          if (doc.page) delete doc.page;
          else doc.page = { size: "a4", orientation: rand(0, 1) ? "landscape" : "portrait" };
          return;
        case 2:
          doc.defaultFont = ["Georgia", "Arial"][rand(0, 1)];
          return;
        default:
          doc.defaultSize = rand(10, 12);
          return;
      }
    default: {
      // A container's own properties: the whole block is replaced.
      const target = pick([...ofType("list"), ...ofType("table")]);
      if (!target) return;
      if (target.type === "list") target.ordered = !target.ordered;
      else target.width = rand(50, 100);
    }
  }
}

{
  const rand = seeded(7);
  for (let run = 0; run < 80; run++) {
    const a = randomDoc(rand);
    const b = structuredClone(a);
    for (let k = 0, n = rand(1, 4); k < n; k++) randomNestedEdit(rand, b);
    cases.push({ name: `random nested ${run}`, a, b });
    cases.push({ name: `random nested reversed ${run}`, a: b, b: a });
  }
}

// One entry moved in each kind of nested list, from a three-word vocabulary so
// entries repeat: ties decide which entry the diff calls moved.
{
  const rand = seeded(13);
  const word = () => VOCAB[rand(0, 2)]!;
  const move = (list: Any[]) => {
    const [moved] = list.splice(rand(0, list.length - 1), 1);
    list.splice(rand(0, list.length), 0, moved);
  };
  const row = () => ({ cells: Array.from({ length: rand(1, 2) }, () => ({ blocks: [p(word())] })) });
  const lists: [string, () => Any, (doc: Any) => Any[]][] = [
    ["rows", () => ({ type: "table", rows: Array.from({ length: 6 }, row) }), (doc) => doc.blocks[1].rows],
    ["cells", () => ({ type: "table", rows: [{ cells: Array.from({ length: 6 }, () => ({ blocks: [p(word())] })) }] }), (doc) => doc.blocks[1].rows[0].cells],
    ["cell blocks", () => ({ type: "table", rows: [{ cells: [{ blocks: Array.from({ length: 6 }, () => p(word())) }] }] }), (doc) => doc.blocks[1].rows[0].cells[0].blocks],
    ["items", () => ({ type: "list", items: Array.from({ length: 6 }, () => ({ runs: [{ text: word() }] })) }), (doc) => doc.blocks[1].items],
    ["children", () => ({ type: "list", items: [{ runs: [{ text: "parent" }], children: Array.from({ length: 6 }, () => ({ runs: [{ text: word() }] })) }] }), (doc) => doc.blocks[1].items[0].children],
    ["quote blocks", () => ({ type: "quote", blocks: Array.from({ length: 6 }, () => p(word())) }), (doc) => doc.blocks[1].blocks],
  ];
  for (const [name, container, listIn] of lists) {
    for (let run = 0; run < 12; run++) {
      const a = { blocks: [p("before"), container(), p("after")] };
      const b = structuredClone(a);
      move(listIn(b));
      cases.push({ name: `random moved ${name} ${run}`, a, b });
    }
  }
}

// The other branches of the algorithm.
const quote = (...texts: string[]) => ({ type: "quote", blocks: texts.map(p) });
const saved = (doc: Any) => Agent.read(Agent.toBytes(doc));
const splitRuns = {
  blocks: [
    { type: "paragraph", runs: [{ text: "Split " }, { text: "run" }] },
    { type: "table", rows: [{ header: true, cells: [{ blocks: [p("Head")] }] }] },
  ],
};
const many = (prefix: string, n: number) => ({ blocks: Array.from({ length: n }, (_, i) => p(`${prefix} ${i}`)) });

cases.push(
  { name: "a save of the ops document is no change", a: lwOpsDoc(), b: saved(lwOpsDoc()) },
  { name: "a save of the canonical fixture is no change", a: canonical, b: saved(canonical) },
  { name: "split runs and a header row are no change", a: splitRuns, b: saved(splitRuns) },
  { name: "identical documents", a: lwOpsDoc(), b: lwOpsDoc() },
  {
    name: "a quote with half its paragraphs reworded is edited inside",
    a: { blocks: [quote("one", "two", "three", "four")] },
    b: { blocks: [quote("one", "2", "three", "4")] },
  },
  {
    name: "a quote with more than half reworded is replaced",
    a: { blocks: [quote("one", "two", "three", "four")] },
    b: { blocks: [quote("1", "2", "three", "4")] },
  },
  {
    name: "a single-paragraph quote reworded is one inner op",
    a: { blocks: [quote("one")] },
    b: { blocks: [quote("uno")] },
  },
  {
    name: "a list item gaining its first children is replaced",
    a: { blocks: [{ type: "list", items: [{ runs: [{ text: "a" }] }, { runs: [{ text: "b" }] }] }] },
    b: { blocks: [{ type: "list", items: [{ runs: [{ text: "a" }], children: [{ runs: [{ text: "a1" }] }] }, { runs: [{ text: "b" }] }] }] },
  },
  {
    name: "a row losing its header flag is replaced",
    a: { blocks: [{ type: "table", rows: [{ header: true, cells: [{ blocks: [p("h")] }] }, { cells: [{ blocks: [p("x")] }] }] }] },
    b: { blocks: [{ type: "table", rows: [{ cells: [{ blocks: [p("h")] }] }, { cells: [{ blocks: [p("x")] }] }] }] },
  },
  {
    name: "a shaded cell is replaced",
    a: { blocks: [{ type: "table", rows: [{ cells: [{ blocks: [p("a")] }, { blocks: [p("b")] }] }] }] },
    b: { blocks: [{ type: "table", rows: [{ cells: [{ blocks: [p("a")] }, { blocks: [p("b")], shading: "#EEEEEE" }] }] }] },
  },
  {
    name: "identical blocks moved among duplicates",
    a: { blocks: [p("x"), p("y"), p("x"), p("z"), p("x"), p("y")] },
    b: { blocks: [p("y"), p("x"), p("x"), p("z"), p("y"), p("x")] },
  },
  {
    name: "a block moved and another reworded",
    a: { blocks: [p("a"), p("b"), p("c"), p("d"), p("e")] },
    b: { blocks: [p("e"), p("a"), p("B"), p("c"), p("d")] },
  },
  { name: "every block removed", a: lwOpsDoc(), b: { title: "Q3 review", blocks: [] } },
  { name: "every block added", a: { title: "Q3 review", blocks: [] }, b: lwOpsDoc() },
  {
    name: "page and fonts removed",
    a: { title: "T", page: { size: "a4" }, defaultFont: "Georgia", defaultSize: 11, blocks: [p("a")] },
    b: { title: "U", blocks: [p("a")] },
  },
  {
    // PHP stores the key "5" as the int 5. Through 0.6.2 diff() emitted
    // `"key": 5`, which the reducer refuses as not a string, so the replay check
    // failed and the diff fell back to `doc.replace`. 0.6.3 emits "5".
    name: "a numeric top-level key changed is a doc.set with a string key",
    a: { title: "x", "5": "y", blocks: [p("a")] },
    b: { title: "z", "5": "w", blocks: [p("a")] },
  },
  // Past ALIGN_LIMIT the changed middle is one hunk: 500 replaces and a removal.
  { name: "a list too long to align", a: many("old", 501), b: many("new", 500) },
);

// The alignment on its own: short lines from a two-letter alphabet, so nearly
// every step is a tie and the delete-first rule decides the hunks.
{
  const rand = seeded(11);
  for (let run = 0; run < 60; run++) {
    const line = () => Array.from({ length: rand(0, 7) }, () => (rand(0, 1) ? "p" : "q"));
    cases.push({ name: `hunks ${run}`, hunks: [line(), line()] });
  }
  cases.push({
    name: "hunks past ALIGN_LIMIT",
    hunks: [["s", ...Array.from({ length: 501 }, (_, i) => `a${i}`), "e"], ["s", ...Array.from({ length: 500 }, (_, i) => `b${i}`), "e"]],
  });
}

// The reducer on its own, with ops no diff would emit.
const odd = { ...lwOpsDoc(), "a/b": { blocks: [p("slash")] }, "m~n": {}, x: { blocks: { a: 1 } } };
cases.push(
  { name: "reduce: insert clamps and casts", a: lwOpsDoc(), ops: [
    { op: "blocks.insert", path: "/blocks", index: 99, block: p("end") },
    { op: "blocks.insert", path: "/blocks", index: -5, block: p("start") },
    { op: "blocks.insert", path: "/blocks", index: "2", block: p("string index") },
    { op: "blocks.insert", path: "/blocks", index: 1.7, block: p("float index") },
    { op: "blocks.insert", path: "/blocks", index: null, block: p("null index") },
    { op: "blocks.insert", path: "/blocks", block: p("no index") },
    { op: "blocks.insert", path: "/blocks", index: 0, block: null },
    { op: "blocks.insert", path: "/blocks", index: 0 },
  ] },
  { name: "reduce: remove, replace and move ranges", a: lwOpsDoc(), ops: [
    { op: "blocks.remove", path: "/blocks", index: "1" },
    { op: "blocks.remove", path: "/blocks", index: -1 },
    { op: "blocks.remove", path: "/blocks" },
    { op: "blocks.replace", path: "/blocks", index: 0 },
    { op: "blocks.replace", path: "/blocks", index: 7, block: p("past the end") },
    { op: "blocks.replace", path: "/blocks", index: 6, block: p("last") },
    { op: "blocks.move", path: "/blocks", from: 0, to: 99 },
    { op: "blocks.move", path: "/blocks", from: 5, to: -3 },
    { op: "blocks.move", path: "/blocks", from: 2 },
    { op: "blocks.move", path: "/blocks", from: 9, to: 0 },
  ] },
  { name: "reduce: nested lists, rows and cells", a: lwOpsDoc(), ops: [
    { op: "items.move", path: "/blocks/3/items", from: 1, to: 0 },
    { op: "items.remove", path: "/blocks/3/items/1/children", index: 0 },
    { op: "items.insert", path: "/blocks/3/items/0/children", index: 0, item: { runs: [{ text: "new" }] } },
    { op: "rows.move", path: "/blocks/4/rows", from: 1, to: 0 },
    { op: "cells.replace", path: "/blocks/4/rows/0/cells", index: 1, cell: { blocks: [p("cell")] } },
    { op: "cells.remove", path: "/blocks/4/rows/1/cells", index: 0 },
    { op: "blocks.insert", path: "/blocks/4/rows/0/cells/0/blocks", index: 1, block: p("second") },
    { op: "blocks.move", path: "/blocks/05/blocks", from: 1, to: 0 },
    { op: "items.insert", path: "/blocks/0/items", index: 0, item: { runs: [{ text: "on a heading" }] } },
  ] },
  { name: "reduce: paths that do not resolve", a: lwOpsDoc(), ops: [
    { op: "blocks.remove", path: "", index: 0 },
    { op: "blocks.remove", path: "blocks", index: 0 },
    { op: "blocks.remove", path: "/", index: 0 },
    { op: "blocks.remove", path: "/blocks/99/blocks", index: 0 },
    { op: "blocks.remove", path: "/blocks/-1/blocks", index: 0 },
    { op: "blocks.remove", path: "/blocks/length/blocks", index: 0 },
    { op: "items.remove", path: "/blocks/3/items/0/children/0", index: 0 },
    { op: "cells.remove", path: "/blocks/4/rows", index: 0 },
    { op: "rows.insert", path: "/title/rows", index: 0, row: { cells: [] } },
    { op: "blocks.remove", path: 5, index: 0 },
  ] },
  { name: "reduce: pointer escapes, maps and a list reached by name", a: odd, ops: [
    { op: "blocks.insert", path: "/a~1b/blocks", index: 1, block: p("escaped slash") },
    { op: "blocks.insert", path: "/m~0n/blocks", index: 0, block: p("into an empty object") },
    { op: "blocks.remove", path: "/x/blocks", index: 0 },
  ] },
  // PHP 0.6.3: a non-empty list is reached by position only; an empty one still takes a name.
  { name: "reduce: a list reached by name", a: { ...odd, e: [], obj: { "0": { blocks: [p("zero")] } } }, ops: [
    { op: "blocks.insert", path: "/blocks/blocks", index: 0, block: p("list reached by name") },
    { op: "blocks.remove", path: "/blocks/x/blocks", index: 0 },
    { op: "items.insert", path: "/blocks/3/items/children", index: 0, item: { runs: [{ text: "items by name" }] } },
    { op: "blocks.insert", path: "/blocks/4/rows/1/cells/first/blocks", index: 0, block: p("cells by name") },
    { op: "blocks.insert", path: "/obj/blocks", index: 0, block: p("a map keyed 0 is a list") },
    { op: "blocks.insert", path: "/obj/0/blocks", index: 0, block: p("but a position reaches it") },
    { op: "blocks.insert", path: "/e/blocks", index: 0, block: p("an empty list takes a name") },
    { op: "blocks.insert", path: "/blocks/2/runs/0/blocks", index: 0, block: p("a map inside a list") },
  ] },
  { name: "reduce: doc ops and names that are not ops", a: lwOpsDoc(), ops: [
    { op: "doc.set", key: "", value: 1 },
    { op: "doc.set", key: "blocks", value: [] },
    { op: "doc.set", key: "title", value: false },
    { op: "doc.set", key: "defaultSize", value: 0 },
    { op: "doc.set", key: 5, value: "numeric" },
    { op: "doc.set", key: "page", value: { size: "legal" } },
    { op: "doc.set", key: "page" },
    { op: "doc.replace", doc: "not a document" },
    { op: "blocks" },
    { op: "blocks.", path: "/blocks" },
    { op: "items.splice", path: "/blocks/3/items", index: 0 },
    { op: 5, path: "/blocks" },
    { path: "/blocks", index: 0 },
    { op: "doc.set", key: "title", value: "kept" },
  ] },
  // PHP 0.6.2: a position is an int or a digit string, and op, path and key are strings.
  { name: "reduce: positions that are not positions", a: lwOpsDoc(), ops: [
    { op: "blocks.remove", path: "/blocks", index: "abc" },
    { op: "blocks.remove", path: "/blocks", index: " 1" },
    { op: "blocks.remove", path: "/blocks", index: "-1" },
    { op: "blocks.remove", path: "/blocks", index: "1e0" },
    { op: "blocks.remove", path: "/blocks", index: true },
    { op: "blocks.remove", path: "/blocks", index: [0] },
    { op: "blocks.remove", path: "/blocks", index: {} },
    { op: "blocks.replace", path: "/blocks", index: "x", block: p("x") },
    { op: "blocks.replace", path: "/blocks", index: 2.5, block: p("x") },
    { op: "blocks.insert", path: "/blocks", index: true, block: p("x") },
    { op: "blocks.insert", path: "/blocks", index: "", block: p("x") },
    { op: "blocks.move", path: "/blocks", from: "first", to: 3 },
    { op: "blocks.move", path: "/blocks", from: 1, to: "last" },
    { op: "blocks.move", path: "/blocks", from: 1, to: null },
    { op: "doc.set", key: true, value: "x" },
    { op: "doc.set", key: ["title"], value: "x" },
    { op: ["blocks.remove"], path: "/blocks", index: 0 },
    { op: "blocks.remove", path: ["/blocks"], index: 0 },
    { op: "blocks.remove", path: "/blocks", index: "03" },
    { op: "blocks.move", path: "/blocks", from: "0", to: "2" },
    { op: "blocks.insert", path: "/blocks", index: -2, block: p("negative int clamps") },
    { op: "blocks.insert", path: "/blocks", index: "99999999999999999999", block: p("huge digits clamp") },
  ] },
  { name: "reduce: doc.replace", a: lwOpsDoc(), ops: [
    { op: "doc.replace", doc: { blocks: [p("replaced")] } },
    { op: "blocks.insert", path: "/blocks", index: 1, block: p("after") },
  ] },
  { name: "reduce: one op, not a list", a: lwOpsDoc(), ops: { op: "blocks.remove", path: "/blocks", index: 0 } },
  { name: "reduce: an empty op list", a: lwOpsDoc(), ops: [] },
);

const HAS_PHP = phpAvailable();

if (process.env.CI && !HAS_PHP) {
  throw new Error(
    "php is not on PATH. This suite is the cross-engine ops parity guarantee; " +
      "skipping it in CI would report success with no coverage. Install PHP, " +
      "or set LAST_WORD_PHP_SRC and ensure `php` resolves.",
  );
}

describe.skipIf(!HAS_PHP)("cross-engine ops parity (PHP vs TS)", () => {
  let results: Any[];

  beforeAll(() => {
    const dir = mkdtempSync(join(tmpdir(), "last-word-ops-parity-"));
    const file = join(dir, "cases.json");
    writeFileSync(file, JSON.stringify(cases.map(({ a, b, ops, hunks }) => (hunks ? { hunks } : ops !== undefined ? { a, ops } : { a, b }))));
    results = JSON.parse(php([PHP_SCRIPT, file]).toString("utf8"));
  }, 300_000);

  it("answers every case", () => {
    expect(results).toHaveLength(cases.length);
  });

  cases.forEach((testCase, index) => {
    it(testCase.name, () => {
      const phpResult = results[index];
      expect(phpResult.error, "PHP threw").toBeUndefined();

      if (testCase.hunks !== undefined) {
        expect(DocDiff.hunks(...testCase.hunks)).toEqual(phpResult.hunks);
        return;
      }

      const tsOps = testCase.ops ?? JSON.parse(JSON.stringify(Agent.diff(testCase.a, testCase.b)));

      // The same ops, in the same order, with the same fields in the same order.
      expect(normalize(tsOps)).toEqual(normalize(phpResult.ops));
      if (testCase.ops === undefined) {
        expect(tsOps.map((op: Any) => Object.keys(op))).toEqual(phpResult.ops.map((op: Any) => Object.keys(op)));
      }

      // And replaying them gives the same document.
      const tsReduced = JSON.parse(JSON.stringify(Agent.reduce(testCase.a, tsOps)));
      expect(normalize(tsReduced)).toEqual(normalize(phpResult.reduced));

      if (testCase.b !== undefined) {
        // No ops means the two write the same file, not that they are the same JSON.
        expect(tsOps.length > 0 ? DocDiff.same(tsReduced, testCase.b) : Agent.equivalent(tsReduced, testCase.b)).toBe(true);
      }
    });
  });

  // Parity between two engines that both fell back would pass too. The seeded
  // cases exist to compare ALIGNMENTS, so make sure they still do.
  it("compares alignments on the seeded cases, not fallbacks", () => {
    const ops = cases
      .map((testCase, index) => [testCase.name, results[index].ops ?? []] as const)
      .filter(([name]) => name.startsWith("random"))
      .flatMap(([, list]) => list.map((op: Any) => op.op as string));

    for (const name of DocOpSchema.TYPES.filter((type) => type !== "doc.replace")) {
      expect(ops, name).toContain(name);
    }
    expect(ops).not.toContain("doc.replace");

    const byName = (name: string) => results[cases.findIndex((c) => c.name === name)].ops.map((op: Any) => op.op);
    expect(byName("a quote with half its paragraphs reworded is edited inside")).toEqual(["blocks.replace", "blocks.replace"]);
    expect(byName("a quote with more than half reworded is replaced")).toEqual(["blocks.replace"]);
    expect(byName("a list too long to align")).toHaveLength(501);
    expect(byName("a save of the ops document is no change")).toEqual([]);
    expect(results[cases.findIndex((c) => c.name === "a numeric top-level key changed is a doc.set with a string key")].ops).toEqual([
      { op: "doc.set", key: "5", value: "w" },
      { op: "doc.set", key: "title", value: "z" },
    ]);
  });

  it("publishes the same op schema, byte for byte", () => {
    expect(JSON.stringify(Agent.opSchema())).toBe(php([PHP_SCRIPT, "--op-schema"]).toString("utf8"));
    expect(DocOpSchema.TYPES.length).toBe(18);
  });
});

if (!HAS_PHP) {
  // eslint-disable-next-line no-console
  console.warn("[doc-ops-parity] php not found on PATH — cross-engine ops parity tests skipped.");
}
