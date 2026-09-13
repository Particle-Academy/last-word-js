/**
 * The shape decisions every legacy reader makes identically, and the PHP string
 * semantics they are written against. Mirrors PHP `Reader\Structure`.
 *
 * `mergeRuns` and `lists` are the PHP engine's rules, unchanged: neighbouring
 * runs with equal formatting merge, and a flat sequence of list paragraphs
 * becomes the nested list model, a change of orderedness at the top level
 * starting a new list.
 *
 * The helpers below them exist because the three engines must return the same
 * document for the same bytes, and JavaScript's `trim()` and `\s` are not PHP's:
 * both treat U+00A0 as white space, so a paragraph holding only a no-break space
 * would vanish here and survive in PHP.
 */

import type { Block, ListItem, Run } from "../schema/types";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

export interface ListEntry {
  ilvl: number;
  ordered: boolean;
  runs: Run[];
}

/** Run keys that, when equal, let two neighbouring runs become one. */
const FORMAT_KEYS = ["bold", "italic", "underline", "strike", "code", "link"] as const;

export const Structure = {
  mergeRuns(runs: Any[]): Run[] {
    const out: Any[] = [];
    for (const raw of runs) {
      if ((raw.text ?? "") === "") continue;
      const run = normalise(raw);
      const last = out[out.length - 1];
      if (last !== undefined && sameFormat(last, run)) {
        last.text += run.text;
        continue;
      }
      out.push(run);
    }
    return out;
  },

  text(runs: Any[]): string {
    return runs.map((r) => String(r.text ?? "")).join("");
  },

  lists(entries: ListEntry[]): Block[] {
    const blocks: Block[] = [];
    let pending: ListEntry[] = [];
    for (const entry of entries) {
      if (pending.length > 0 && entry.ilvl === 0 && pending[0]!.ordered !== entry.ordered) {
        blocks.push(assembleList(pending));
        pending = [];
      }
      pending.push(entry);
    }
    if (pending.length > 0) blocks.push(assembleList(pending));
    return blocks;
  },
};

function assembleList(entries: ListEntry[]): Block {
  const block: Any = { type: "list" };
  if (entries[0]!.ordered) block.ordered = true;
  block.items = [];

  const stack: ListItem[][] = [block.items];
  for (const entry of entries) {
    const depth = Math.min(entry.ilvl, stack.length);
    while (stack.length - 1 > depth) stack.pop();
    let parent = stack[stack.length - 1]!;
    if (depth > stack.length - 1 && parent.length > 0) {
      const last = parent[parent.length - 1]!;
      if (!last.children) last.children = [];
      stack.push(last.children);
      parent = last.children;
    }
    parent.push({ runs: entry.runs });
  }

  return block;
}

function normalise(run: Any): Any {
  const out: Any = { text: String(run.text) };
  for (const key of FORMAT_KEYS) {
    if (key === "link") {
      if (typeof run.link === "string" && run.link !== "") out.link = run.link;
    } else if (run[key]) {
      out[key] = true;
    }
  }
  return out;
}

function sameFormat(a: Any, b: Any): boolean {
  const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
  keys.delete("text");
  for (const key of keys) {
    if (a[key] !== b[key]) return false;
  }
  // PHP compares the arrays with ===, which also compares key order.
  return Object.keys(a).join() === Object.keys(b).join();
}

/** PHP `trim()` with its default character set: space, tab, LF, CR, NUL, vertical tab. */
export function phpTrim(s: string): string {
  return s.replace(/^[ \t\n\r\0\x0B]+|[ \t\n\r\0\x0B]+$/g, "");
}

/** PHP `(int)` of a string: leading white space, an optional sign, digits; anything else is 0. */
export function phpInt(s: string | undefined | null): number {
  const m = /^[ \t\n\r\v\f]*([+-]?\d+)/.exec(s ?? "");
  if (!m) return 0;
  const n = Number(m[1]);
  return Number.isFinite(n) ? n : 0;
}

/** PHP `min(max(...))` clamp, kept in one place so the ports cannot disagree on order. */
export function clamp(value: number, low: number, high: number): number {
  return Math.max(low, Math.min(high, value));
}
