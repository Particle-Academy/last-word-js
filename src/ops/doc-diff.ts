import { Agent } from "../agent";
import { DocReducer } from "./doc-reducer";
import { type Any, compareStrings, entriesOf, get, isArr, valuesOf, withoutKey } from "./php";
import type { DocOp, DocOpKind } from "./types";

/** PHP `json_encode`'s depth argument in `canon()`: more nested arrays and objects than this throw. */
const CANON_DEPTH = 4096;

/** Block type => the list inside it that is diffed, and that list's kind. */
const CONTAINERS: Record<string, readonly [string, DocOpKind]> = {
  quote: ["blocks", "blocks"],
  list: ["items", "items"],
  table: ["rows", "rows"],
};

/**
 * The op list that turns one Last Word document into another. Mirrors PHP
 * `LastWord\Ops\DocDiff` as of 0.6.3: the same algorithm, so the same inputs give
 * the same ops in the same order in both runtimes (pinned against the PHP
 * package by `tests/doc-ops-parity.test.ts`).
 *
 * ## Two guarantees, and where each applies
 *
 * 1. **Same file, no ops.** When `a` and `b` write the same document —
 *    compared as `read(toBytes(...))`, so runs the reader merges, a header row's
 *    bold, an empty paragraph the writer drops are not changes — the diff is
 *    `[]`. That is what makes `diff(d, read(toBytes(d)))` `[]`: saving without
 *    a change records nothing.
 * 2. **Otherwise, exact.** `reduce(a, diff(a, b))` equals `b`, key order aside.
 *    The ops are computed on the documents as given and VERIFIED by replaying
 *    them through `DocReducer`; if they do not reproduce `b`, the diff is one
 *    `doc.replace`.
 *
 * ## Small edits stay small
 *
 * Every list — the top-level blocks, a quote's blocks, a list's items and their
 * children, a table's rows, a row's cells, a cell's blocks — is ALIGNED by
 * content (a longest common subsequence), so rewording one paragraph is one
 * `blocks.replace`, moving one is one `blocks.move`, and rewording a paragraph
 * inside a table cell is one `blocks.replace` at that cell's path rather than a
 * new table. A changed container whose own properties are unchanged is diffed
 * inside; one whose properties changed, or whose inner diff would be more than
 * half its length in ops, is replaced whole.
 *
 * ## Determinism
 *
 * The same inputs give the same ops, in the same order, in the PHP, Node and
 * Python ports: alignment ties break toward deleting first, and the alignment is
 * skipped past `ALIGN_LIMIT` cells of work.
 */
export const DocDiff = {
  /** Above this many LCS cells (after trimming the common prefix and suffix), a list is not aligned. */
  ALIGN_LIMIT: 250_000,

  diff(a: Any, b: Any): DocOp[] {
    if (DocDiff.same(a, b) || DocDiff.equivalent(a, b)) {
      return [];
    }

    const ops: Any[] = [];
    const keys = [...new Set([...entriesOf(a), ...entriesOf(b)].map(([key]) => key))].sort(compareStrings);

    for (const key of keys) {
      if (key === "blocks") {
        continue;
      }
      if (!DocDiff.same(get(a, key), get(b, key))) {
        // Always a string. PHP stores "5" as the int 5 and, before 0.6.3, emitted
        // `"key": 5`, which the reducer refuses: the replay check failed and the
        // whole diff fell back to doc.replace. PHP casts it now.
        ops.push({ op: "doc.set", key, value: get(b, key) });
      }
    }

    for (const op of listOps("blocks", "/blocks", listOf(a, "blocks"), listOf(b, "blocks"))) {
      ops.push(op);
    }

    if (DocDiff.same(DocReducer.applyAll(a, ops), b)) {
      return ops;
    }

    return [{ op: "doc.replace", doc: b }];
  },

  /** Whether two documents write the same file: both are written and read back. */
  equivalent(a: Any, b: Any): boolean {
    return DocDiff.same(Agent.read(Agent.toBytes(a)), Agent.read(Agent.toBytes(b)));
  },

  /**
   * Structural equality with map key order ignored and list order kept.
   *
   * It follows PHP's canonical form, where lists and maps are one array type:
   * `{}` equals `[]`, and an object keyed exactly "0".."n-1" equals the list of
   * its values.
   *
   * Two divergences from the PHP `same()` cannot be removed, because the
   * distinctions do not exist in JS values:
   *
   * - PHP encodes the float `1.0` as `1.0` and the int `1` as `1`
   *   (`JSON_PRESERVE_ZERO_FRACTION`), so it calls them different; JS has one
   *   number, so here they are the same (and so are `0` and `-0`). Parsed JSON
   *   gives PHP a float only for a literal written with a fraction or exponent,
   *   so `{"width": 80}` against `{"width": 80.0}` is a change in PHP and none
   *   here: PHP emits a `blocks.replace` (or, if the two write the same file,
   *   nothing) where this port emits nothing.
   * - PHP sorts a map's keys as strings before encoding, so a map keyed
   *   0..n-1 out of order becomes a list when n <= 10 ("10" sorts before "2")
   *   and stays a map from 11 keys up. A JS object has no such order to compare,
   *   so here it is always the list. No valid document carries one.
   *
   * ## A value JSON cannot hold throws
   *
   * As PHP 0.6.1 does (`JsonException`), rather than encoding it to something
   * another value also encodes to. PHP's cases are invalid UTF-8, NAN and INF;
   * here it is a `TypeError` for:
   *
   * - a number that is not finite (`NaN`, `Infinity`), which `JSON.stringify`
   *   writes as `null`, so it would equal `null` and every other one;
   * - a string with a lone surrogate — the JS form of text that is not valid
   *   Unicode, which UTF-8 (and therefore PHP, and the writer) cannot carry.
   *   `JSON.stringify` would keep two of them apart, but PHP could never be
   *   handed either, so this port refuses them the way PHP refuses bad bytes;
   * - a function, symbol or bigint;
   * - more than 4096 nested arrays and objects (PHP's depth limit), which is
   *   also where a cyclic value ends up.
   *
   * `undefined` is not refused: an object key holding it is absent and a list
   * entry holding it is `null`, as JSON has them.
   */
  same(a: unknown, b: unknown): boolean {
    return canon(a) === canon(b);
  },

  /**
   * Hunks of a longest-common-subsequence alignment:
   * [start in a, deleted, inserted, start in b].
   *
   * Ties break toward deleting first. Past `ALIGN_LIMIT` the changed middle is
   * one hunk.
   */
  hunks(a: readonly string[], b: readonly string[]): [number, number, number, number][] {
    const n = a.length;
    const m = b.length;
    let prefix = 0;

    while (prefix < n && prefix < m && a[prefix] === b[prefix]) {
      prefix++;
    }

    let suffix = 0;

    while (suffix < n - prefix && suffix < m - prefix && a[n - 1 - suffix] === b[m - 1 - suffix]) {
      suffix++;
    }

    const midA = a.slice(prefix, n - suffix);
    const midB = b.slice(prefix, m - suffix);
    const rows = midA.length;
    const cols = midB.length;

    if (rows === 0 && cols === 0) {
      return [];
    }

    if (rows * cols > DocDiff.ALIGN_LIMIT) {
      return [[prefix, rows, cols, prefix]];
    }

    // lengths[i][j] = LCS of midA[i..] and midB[j..]
    const lengths: Int32Array[] = [];
    for (let i = 0; i <= rows; i++) {
      lengths.push(new Int32Array(cols + 1));
    }

    for (let i = rows - 1; i >= 0; i--) {
      for (let j = cols - 1; j >= 0; j--) {
        lengths[i]![j] =
          midA[i] === midB[j] ? lengths[i + 1]![j + 1]! + 1 : Math.max(lengths[i + 1]![j]!, lengths[i]![j + 1]!);
      }
    }

    const hunks: [number, number, number, number][] = [];
    let open: [number, number, number, number] | null = null;
    let i = 0;
    let j = 0;

    while (i < rows || j < cols) {
      if (i < rows && j < cols && midA[i] === midB[j]) {
        if (open !== null) {
          hunks.push(open);
          open = null;
        }
        i++;
        j++;
        continue;
      }

      open ??= [prefix + i, 0, 0, prefix + j];

      if (j >= cols || (i < rows && lengths[i + 1]![j]! >= lengths[i]![j + 1]!)) {
        open[1]++;
        i++;
      } else {
        open[2]++;
        j++;
      }
    }

    if (open !== null) {
      hunks.push(open);
    }

    return hunks;
  },
};

/**
 * Ops that turn list `from` into list `to`, the list at `path`.
 *
 * 1. Pairs: identical items the alignment keeps; items changed in place
 *    (the first of each run of deletes paired with the first of the inserts
 *    beside it); and an item deleted in one place and inserted identical in
 *    another, which is a move.
 * 2. Changed-in-place items are edited first, at their old index.
 * 3. Unpaired old items are removed, last first.
 * 4. Walking the target in order, each position is filled by an insert or a
 *    move.
 */
function listOps(kind: DocOpKind, path: string, from: readonly Any[], to: readonly Any[]): Any[] {
  const hashFrom = from.map(canon);
  const hashTo = to.map(canon);
  const n = from.length;
  const m = to.length;

  /** target index => source index */
  const source = new Map<number, number>();
  const changed: [number, number][] = [];
  let deleted: number[] = [];
  const inserted: number[] = [];

  let i = 0;
  let j = 0;

  for (const [start, dels, ins, targetStart] of DocDiff.hunks(hashFrom, hashTo)) {
    while (i < start) {
      source.set(j++, i++);
    }

    const paired = Math.min(dels, ins);

    for (let k = 0; k < paired; k++) {
      changed.push([i + k, targetStart + k]);
      source.set(targetStart + k, i + k);
    }
    for (let k = paired; k < dels; k++) {
      deleted.push(i + k);
    }
    for (let k = paired; k < ins; k++) {
      inserted.push(targetStart + k);
    }

    i += dels;
    j = targetStart + ins;
  }

  while (i < n) {
    source.set(j++, i++);
  }

  // An item removed here and inserted identical there is a move: each leftover
  // insert, in order, takes the first leftover delete with the same content.
  for (const target of inserted) {
    const y = deleted.findIndex((old) => hashFrom[old] === hashTo[target]);
    if (y !== -1) {
      source.set(target, deleted[y]!);
      deleted = [...deleted.slice(0, y), ...deleted.slice(y + 1)];
    }
  }

  const [valueKey] = DocReducer.KINDS[kind];
  const ops: Any[] = [];

  for (const [old, target] of changed) {
    for (const op of itemOps(kind, path, old, from[old], to[target])) {
      ops.push(op);
    }
  }

  const removed = deleted.slice().sort((x, y) => y - x);

  for (const old of removed) {
    ops.push({ op: `${kind}.remove`, path, index: old });
  }

  // The working order: surviving source indices, in source order.
  const work: number[] = [];
  for (let s = 0; s < n; s++) {
    if (!removed.includes(s)) work.push(s);
  }

  for (let t = 0; t < m; t++) {
    const want = source.get(t) ?? null;

    if (want === null) {
      ops.push({ op: `${kind}.insert`, path, index: t, [valueKey]: to[t] });
      work.splice(t, 0, -1 - t);
      continue;
    }

    if ((work[t] ?? null) === want) {
      continue;
    }

    const at = work.indexOf(want);
    // PHP's array_search gives false for a miss, which (int) reads as 0. A
    // wanted index is always in the working order, so this never happens.
    ops.push({ op: `${kind}.move`, path, from: at === -1 ? false : at, to: t });
    work.splice(Math.max(at, 0), 1);
    work.splice(t, 0, want);
  }

  return ops;
}

/**
 * Ops for one item changed in place: an edit inside it when it is a container
 * whose own properties did not change, otherwise a replace.
 */
function itemOps(kind: DocOpKind, path: string, index: number, old: Any, next: Any): Any[] {
  const [valueKey] = DocReducer.KINDS[kind];
  const replace = [{ op: `${kind}.replace`, path, index, [valueKey]: next }];

  if (!isArr(old) || !isArr(next)) {
    return replace;
  }

  let child: readonly [string, DocOpKind] | null = null;

  switch (kind) {
    case "blocks": {
      const type = get(next, "type");
      child =
        identical(get(old, "type"), type) && typeof type === "string" && Object.prototype.hasOwnProperty.call(CONTAINERS, type)
          ? CONTAINERS[type]!
          : null;
      break;
    }
    case "items":
      child = ["children", "items"];
      break;
    case "rows":
      child = ["cells", "cells"];
      break;
    case "cells":
      child = ["blocks", "blocks"];
      break;
  }

  if (child === null) {
    return replace;
  }

  const [key, childKind] = child;

  if (!isArr(get(old, key)) || !isArr(get(next, key))) {
    return replace;
  }

  if (!DocDiff.same(withoutKey(old, key), withoutKey(next, key))) {
    return replace;
  }

  const inner = listOps(childKind, `${path}/${index}/${key}`, valuesOf(old[key]), valuesOf(next[key]));

  return inner.length <= Math.max(1, Math.floor(valuesOf(next[key]).length / 2)) ? inner : replace;
}

/**
 * PHP `===` for a block `type`. Only a string can name a container; any other
 * value makes the block a non-container, which is where PHP ends too (a valid
 * document cannot reach here with one: `diff` writes both documents first).
 */
function identical(x: unknown, y: unknown): boolean {
  return !isArr(x) && !isArr(y) && x === y;
}

function listOf(node: Any, key: string): Any[] {
  return isArr(get(node, key)) ? valuesOf(node[key]) : [];
}

/**
 * PHP's canonical JSON, for equality only: map keys sorted, list order kept, and
 * an array keyed exactly 0..n-1 (every empty one included) written as a list.
 * The key order and escaping differ from PHP's bytes; which values compare
 * equal does not (see `same()` for the two exceptions, and for what throws).
 *
 * One stack frame per level (loops, not `map`), so PHP's full depth of 4096
 * fits in Node's default stack.
 */
function canon(value: unknown, depth = 0): string {
  if (typeof value === "object" && value !== null) {
    if (depth >= CANON_DEPTH) {
      throw new TypeError(`DocDiff.same(): more than ${CANON_DEPTH} nested arrays and objects (or a cycle); JSON cannot hold it.`);
    }

    const entries = entriesOf(value);
    let out = "";

    if (Array.isArray(value) || entries.every(([key], i) => key === String(i))) {
      for (let i = 0; i < entries.length; i++) {
        out += (i === 0 ? "" : ",") + canon(entries[i]![1], depth + 1);
      }
      return "[" + out + "]";
    }

    entries.sort((x, y) => compareStrings(x[0], y[0]));
    for (let i = 0; i < entries.length; i++) {
      out += (i === 0 ? "" : ",") + canonString(entries[i]![0]) + ":" + canon(entries[i]![1], depth + 1);
    }
    return "{" + out + "}";
  }

  switch (typeof value) {
    case "string":
      return canonString(value);
    case "number":
      if (!Number.isFinite(value)) {
        throw new TypeError(`DocDiff.same(): ${value} is not a finite number; JSON cannot hold it.`);
      }
      return JSON.stringify(value);
    case "boolean":
      return value ? "true" : "false";
    case "undefined":
      return "null";
    default:
      if (value === null) {
        return "null";
      }
      throw new TypeError(`DocDiff.same(): a ${typeof value} is not a JSON value.`);
  }
}

const LONE_SURROGATE = /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/;

function canonString(value: string): string {
  if (LONE_SURROGATE.test(value)) {
    throw new TypeError("DocDiff.same(): a string with a lone surrogate is not valid Unicode; JSON cannot hold it.");
  }
  return JSON.stringify(value);
}
