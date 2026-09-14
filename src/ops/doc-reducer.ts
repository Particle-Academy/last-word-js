import { type Any, get, has, isArr, isDigits, isList, phpInt, valuesOf, withKey, withoutKey } from "./php";
import type { DocOpKind } from "./types";

/**
 * Apply `DocOpSchema` ops to a Last Word document, returning a new document.
 * Mirrors PHP `LastWord\Ops\DocReducer` as of 0.6.2, op for op (pinned against
 * the PHP package by `tests/doc-ops-parity.test.ts`).
 *
 * Pure: the input is never modified. An op whose `path` does not reach a list of
 * the kind it edits, or whose index is out of range, is skipped, so a replayed
 * history degrades rather than throws. Unchanged parts of the input are shared
 * with the result rather than copied, as in any immutable reducer: nothing here
 * mutates either, but a caller that mutates a result in place reaches into the
 * input too.
 *
 * ## Paths
 *
 * A list op names the LIST it edits with a JSON Pointer, and the item by index
 * in it:
 *
 * | op | path ends in | e.g. |
 * |---|---|---|
 * | `blocks.*` | `blocks` | `/blocks`, `/blocks/4/blocks` (a quote), `/blocks/2/rows/1/cells/0/blocks` |
 * | `items.*` | `items` or `children` | `/blocks/3/items`, `/blocks/3/items/0/children` |
 * | `rows.*` | `rows` | `/blocks/2/rows` |
 * | `cells.*` | `cells` | `/blocks/2/rows/1/cells` |
 *
 * Each kind has `insert {path, index, <value>}`, `remove {path, index}`,
 * `move {path, from, to}` and `replace {path, index, <value>}`, where the value
 * key is `block`, `item`, `row` or `cell`. Inserting into a `children` list that
 * is not there yet creates it.
 *
 * A position (`index`, `from`, `to`) is an integer or a string of digits; an op
 * carrying anything else is skipped rather than read as 0. An `op`, `path` or
 * `doc.set` `key` that is not a string skips the op too.
 */
export const DocReducer = {
  /** List kind => [the value key its ops carry, the path tokens a list of that kind ends in]. */
  KINDS: {
    blocks: ["block", ["blocks"]],
    items: ["item", ["items", "children"]],
    rows: ["row", ["rows"]],
    cells: ["cell", ["cells"]],
  } as const satisfies Record<DocOpKind, readonly [string, readonly string[]]>,

  applyAll(doc: Any, ops: readonly Any[]): Any {
    for (const op of ops) {
      doc = DocReducer.apply(doc, op);
    }
    return doc;
  },

  apply(doc: Any, op: Any): Any {
    // Strings only: PHP casting an array to string gives "Array" and a warning.
    const name = typeof get(op, "op") === "string" ? (op.op as string) : "";

    if (name === "doc.replace") {
      return isArr(get(op, "doc")) ? op.doc : doc;
    }

    if (name === "doc.set") {
      // A string key only: PHP's `(string) true` is "1", which set a key "1".
      const key = typeof get(op, "key") === "string" ? (op.key as string) : "";

      if (key === "" || key === "blocks") {
        return doc;
      }

      if (get(op, "value") === null) {
        return withoutKey(doc, key);
      }

      return withKey(doc, key, op.value);
    }

    const dot = name.indexOf(".");
    const kind = dot === -1 ? name : name.slice(0, dot);
    const action = dot === -1 ? "" : name.slice(dot + 1);

    if (!Object.prototype.hasOwnProperty.call(DocReducer.KINDS, kind) || !ACTIONS.includes(action)) {
      return doc;
    }

    const [valueKey, ends] = DocReducer.KINDS[kind as DocOpKind];
    const tokens = typeof get(op, "path") === "string" ? DocReducer.tokens(op.path) : null;

    if (tokens === null || tokens.length === 0 || !(ends as readonly string[]).includes(tokens[tokens.length - 1]!)) {
      return doc;
    }

    return edit(doc, tokens, (list) => editList(list, action, op, valueKey));
  },

  /** RFC 6901 JSON Pointer tokens, or null for a pointer that is not one. */
  tokens(pointer: string): string[] | null {
    if (pointer === "" || pointer[0] !== "/") {
      return null;
    }

    return pointer
      .slice(1)
      .split("/")
      .map((token) => token.split("~1").join("/").split("~0").join("~"));
  },
};

const ACTIONS = ["insert", "remove", "move", "replace"];

/**
 * Walk to the list at `tokens` and replace it with `change(list)`. A null
 * result (or an unreachable path) leaves the document as it was.
 *
 * As in PHP, a numeric token indexes a list, and a list reached by a key that is
 * not an index becomes a map: an insert at `/blocks/blocks` turns the blocks
 * list into an object holding the old blocks under "0", "1", … and the new list
 * under "blocks".
 */
function edit(node: Any, tokens: readonly string[], change: (list: Any[] | null) => Any[] | null): Any {
  const [token, ...rest] = tokens as [string, ...string[]];
  const key = isList(node) && isDigits(token) ? String(phpInt(token)) : token;

  if (rest.length === 0) {
    const current = get(node, key);

    if (current !== null && (!isArr(current) || !isList(current))) {
      return node;
    }

    const changed = change(current === null ? null : valuesOf(current));

    return changed !== null ? withKey(node, key, changed) : node;
  }

  const next = get(node, key);

  if (!isArr(next)) {
    return node;
  }

  const edited = edit(next, rest, change);

  return edited === next ? node : withKey(node, key, edited);
}

function editList(list: Any[] | null, action: string, op: Any, valueKey: string): Any[] | null {
  if (list === null && action !== "insert") {
    return null;
  }

  const out = list === null ? [] : list.slice();
  const count = out.length;

  switch (action) {
    case "insert": {
      if (!has(op, valueKey)) {
        return null;
      }
      // A missing index appends; one that is present but not a position skips.
      const index = has(op, "index") ? position(op.index) : count;
      if (index === null) {
        return null;
      }
      out.splice(Math.max(0, Math.min(count, index)), 0, op[valueKey]);

      return out;
    }

    case "remove": {
      const index = position(get(op, "index")) ?? -1;
      if (index < 0 || index >= count) {
        return null;
      }
      out.splice(index, 1);

      return out;
    }

    case "replace": {
      const index = position(get(op, "index")) ?? -1;
      if (index < 0 || index >= count || !has(op, valueKey)) {
        return null;
      }
      out[index] = op[valueKey];

      return out;
    }

    case "move": {
      const from = position(get(op, "from")) ?? -1;
      // A missing `to` stays put; one that is present but not a position skips.
      const to = has(op, "to") ? position(op.to) : from;
      if (from < 0 || from >= count || to === null) {
        return null;
      }
      const [moved] = out.splice(from, 1);
      out.splice(Math.max(0, Math.min(out.length, to)), 0, moved);

      return out;
    }
  }

  return null;
}

/**
 * PHP `DocReducer::position()`: an int, or a string of digits. Anything else is
 * null, so an op carrying one is skipped — PHP's `(int) "abc"` is 0, which
 * edited the first item before 0.6.2.
 *
 * PHP's `is_int` sees the JSON literal `2.0` as a float and skips it; a JS
 * number cannot say how it was written, so `2.0` is the integer 2 here. A JSON
 * integer past 2^63 is a float to PHP and is skipped by both.
 */
function position(value: unknown): number | null {
  if (typeof value === "number") {
    return Number.isInteger(value) && Math.abs(value) < 2 ** 63 ? value : null;
  }

  return typeof value === "string" && isDigits(value) ? phpInt(value) : null;
}
