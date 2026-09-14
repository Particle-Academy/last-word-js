/**
 * Fixtures shared by `doc-ops.test.ts` and `doc-ops-parity.test.ts`: the PHP
 * `DocOpsTest` document and edits, ported value for value.
 */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
export type Any = any;

export function p(text: string): Any {
  return { type: "paragraph", runs: [{ text }] };
}

/** PHP `lwOpsDoc()`: every container kind, one level of each nesting. */
export function lwOpsDoc(): Any {
  return {
    title: "Q3 review",
    blocks: [
      { type: "heading", level: 1, runs: [{ text: "Q3 review" }] },
      p("Revenue grew in every region."),
      p("Costs held flat."),
      {
        type: "list",
        items: [
          { runs: [{ text: "North" }], children: [{ runs: [{ text: "Enterprise" }] }] },
          { runs: [{ text: "South" }] },
        ],
      },
      {
        type: "table",
        rows: [
          { header: true, cells: [{ blocks: [p("Region")] }, { blocks: [p("Revenue")] }] },
          { cells: [{ blocks: [p("North")] }, { blocks: [p("1,250,000")] }] },
        ],
      },
      { type: "quote", blocks: [p("Best quarter yet."), p("— the CFO")] },
      { type: "hr" },
      p("Next steps follow."),
    ],
  };
}

export type Shape = { op: string; path?: string; index?: number; from?: number; to?: number; key?: string };

/** PHP dataset 'doc edits': each edit, and the op / path / position it must give. */
export const EDITS: Record<string, [(d: Any) => Any, Shape[]]> = {
  "a paragraph reworded": [
    (d) => {
      d.blocks[2] = p("Costs fell 3%.");
      return d;
    },
    [{ op: "blocks.replace", path: "/blocks", index: 2 }],
  ],
  "a paragraph inserted": [
    (d) => {
      d.blocks.splice(2, 0, p("Margins widened."));
      return d;
    },
    [{ op: "blocks.insert", path: "/blocks", index: 2 }],
  ],
  "a paragraph removed": [
    (d) => {
      d.blocks.splice(1, 1);
      return d;
    },
    [{ op: "blocks.remove", path: "/blocks", index: 1 }],
  ],
  "a paragraph moved": [
    (d) => {
      const [moved] = d.blocks.splice(7, 1);
      d.blocks.splice(1, 0, moved);
      return d;
    },
    [{ op: "blocks.move", path: "/blocks", from: 7, to: 1 }],
  ],
  "the title": [
    (d) => {
      d.title = "Q3 review (final)";
      return d;
    },
    [{ op: "doc.set", key: "title" }],
  ],
  "page settings added": [
    (d) => {
      d.page = { size: "a4", orientation: "landscape" };
      return d;
    },
    [{ op: "doc.set", key: "page" }],
  ],
  "a list item reworded": [
    (d) => {
      d.blocks[3].items[1].runs = [{ text: "South and West" }];
      return d;
    },
    [{ op: "items.replace", path: "/blocks/3/items", index: 1 }],
  ],
  "a nested list item added": [
    (d) => {
      d.blocks[3].items[0].children.push({ runs: [{ text: "Mid-market" }] });
      return d;
    },
    [{ op: "items.insert", path: "/blocks/3/items/0/children", index: 1 }],
  ],
  "a table cell reworded": [
    (d) => {
      d.blocks[4].rows[1].cells[1].blocks = [p("1,300,000")];
      return d;
    },
    [{ op: "blocks.replace", path: "/blocks/4/rows/1/cells/1/blocks", index: 0 }],
  ],
  "a table row added": [
    (d) => {
      d.blocks[4].rows.push({ cells: [{ blocks: [p("South")] }, { blocks: [p("980,400")] }] });
      return d;
    },
    [{ op: "rows.insert", path: "/blocks/4/rows", index: 2 }],
  ],
  "a quoted paragraph reworded": [
    (d) => {
      d.blocks[5].blocks[1] = p("— our CFO");
      return d;
    },
    [{ op: "blocks.replace", path: "/blocks/5/blocks", index: 1 }],
  ],
  "a block changed type": [
    (d) => {
      d.blocks[2] = { type: "heading", level: 2, runs: [{ text: "Costs held flat." }] };
      return d;
    },
    [{ op: "blocks.replace", path: "/blocks", index: 2 }],
  ],
  "a table re-styled": [
    (d) => {
      d.blocks[4].width = 80;
      return d;
    },
    [{ op: "blocks.replace", path: "/blocks", index: 4 }],
  ],
};

/**
 * A fixed-seed PRNG (mulberry32) returning an integer in [min, max], like
 * `mt_rand(min, max)`. Not PHP's Mersenne Twister, so the runs are not PHP's
 * runs: the parity suite sends each generated pair to PHP rather than asking it
 * to regenerate them.
 */
export function seeded(seed: number): (min: number, max: number) => number {
  let state = seed >>> 0;
  return (min, max) => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    const unit = ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    return min + Math.floor(unit * (max - min + 1));
  };
}

const WORDS = ["alpha", "beta", "gamma", "delta", "epsilon", "zeta"];

/**
 * PHP's "keeps the round trip over a seeded run of random edits": 1-4 edits per
 * run to the top-level blocks — an insert, a removal, a reword, a move — or a
 * new title. Returns [a, b] pairs.
 */
export function randomBlockEdits(seed: number, runs: number): [Any, Any][] {
  const rand = seeded(seed);
  const pairs: [Any, Any][] = [];

  for (let run = 0; run < runs; run++) {
    const a = lwOpsDoc();
    const b = lwOpsDoc();

    for (let k = 0, edits = rand(1, 4); k < edits; k++) {
      const blocks: Any[] = b.blocks;
      const count = blocks.length;

      switch (rand(0, 4)) {
        case 0:
          blocks.splice(rand(0, count), 0, p(`${WORDS[rand(0, 5)]} ${rand(1, 99)}`));
          break;
        case 1:
          if (count > 1) blocks.splice(rand(0, count - 1), 1);
          break;
        case 2:
          blocks[rand(0, count - 1)] = p(WORDS[rand(0, 5)]!);
          break;
        case 3: {
          const [moved] = blocks.splice(rand(0, count - 1), 1);
          blocks.splice(rand(0, count - 1), 0, moved);
          break;
        }
        default:
          b.title = `T${rand(1, 9)}`;
      }
    }

    pairs.push([a, b]);
  }

  return pairs;
}

/** Freeze a value all the way down, so a reducer or diff that mutates its input throws. */
export function deepFreeze<T>(value: T): T {
  if (typeof value === "object" && value !== null && !Object.isFrozen(value)) {
    Object.freeze(value);
    for (const key of Object.keys(value)) deepFreeze((value as Any)[key]);
  }
  return value;
}
