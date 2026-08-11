import { describe, it, expect, beforeAll } from "vitest";
import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { Agent, unzipSync } from "../src";

/**
 * Cross-engine WRITER parity — the guarantee this pair was missing.
 *
 * `last-word` and `last-word-js` are a matched PHP + Node pair, and until now
 * only one direction was checked: `cross-read.test.ts` proves the Node reader
 * restores a frozen docx the PHP writer produced. Nothing compared the two
 * WRITERS. Both siblings in this family — `holy-sheet` and `dark-slide` — diff
 * PHP output against Node's for the same input; this one did not, so the two
 * engines could drift apart on everything the frozen fixture happens not to
 * contain, and nothing would say so.
 *
 * PHP is the reference: it shipped first, and the pair's contract is "same
 * document either backend."
 *
 * ## Parts, not the container
 *
 * The .docx files can NEVER match byte-for-byte and it would be wrong to try.
 * PHP writes through `ZipArchive` (DEFLATE, real mtimes); this port writes
 * STORE with a fixed 1980-01-01 DOS date. So the comparison unzips both and
 * diffs each part — which is the real contract anyway, since a reader sees
 * parts and never the compression.
 */

const PHP_SCRIPT = join(__dirname, "..", "scripts", "php-tobytes.php");

/** `php` may only resolve through the shell (Herd shims etc.). */
function php(args: string[], opts: Parameters<typeof execFileSync>[2] = {}): Buffer {
  return execFileSync("php", args, { shell: true, ...opts }) as Buffer;
}

function phpAvailable(): boolean {
  try {
    php(["--version"], { stdio: "ignore" });
    return true;
  } catch {
    return false;
  }
}

const META = { creator: "Parity", created: "2024-01-01T00:00:00Z", modified: "2024-01-01T00:00:00Z" };

const DOCS: Record<string, unknown> = {
  /** The floor: if this diverges, nothing below is worth reading. */
  minimal: {
    title: "Minimal",
    metadata: META,
    blocks: [{ type: "paragraph", runs: [{ text: "Hello." }] }],
  },

  /** Every inline mark, together, since they compose into one run list. */
  inlineMarks: {
    title: "Marks",
    metadata: META,
    blocks: [
      { type: "heading", level: 1, runs: [{ text: "Heading" }] },
      {
        type: "paragraph",
        runs: [
          { text: "plain " },
          { text: "bold", bold: true },
          { text: " " },
          { text: "italic", italic: true },
          { text: " " },
          { text: "under", underline: true },
          { text: " " },
          { text: "struck", strike: true },
          { text: " " },
          { text: "code", code: true },
          { text: " " },
          { text: "link", link: "https://particle.academy" },
          { text: " " },
          { text: "colored", color: "#C0392B" },
          { text: " " },
          { text: "high", highlight: "#FFF3A0" },
        ],
      },
    ],
  },

  /** Structure: lists, a table with a header row, and a block quote. */
  structure: {
    title: "Structure",
    metadata: META,
    blocks: [
      { type: "list", ordered: false, items: [{ runs: [{ text: "one" }] }, { runs: [{ text: "two" }] }] },
      { type: "list", ordered: true, items: [{ runs: [{ text: "first" }] }, { runs: [{ text: "second" }] }] },
      { type: "quote", blocks: [{ type: "paragraph", runs: [{ text: "a quotation" }] }] },
      {
        type: "table",
        rows: [
          {
            header: true,
            cells: [
              { blocks: [{ type: "paragraph", runs: [{ text: "Name", bold: true }] }] },
              { blocks: [{ type: "paragraph", runs: [{ text: "Qty", bold: true }] }] },
            ],
          },
          {
            cells: [
              { blocks: [{ type: "paragraph", runs: [{ text: "Widget" }] }] },
              { blocks: [{ type: "paragraph", runs: [{ text: "3" }] }] },
            ],
          },
        ],
      },
    ],
  },

  /**
   * Non-ASCII, because this is where two engines that index strings
   * differently come apart. PHP slices by BYTE and this port by UTF-16 code
   * unit; they agree today by accident rather than by design, and nothing was
   * checking that for the WRITER.
   */
  unicode: {
    title: "Unicode",
    metadata: META,
    blocks: [
      { type: "heading", level: 2, runs: [{ text: "日本語の見出し" }] },
      {
        type: "paragraph",
        runs: [
          { text: "café naïve " },
          { text: "強調", bold: true },
          { text: " emoji 🎉 " },
          { text: "á combining" },
        ],
      },
      { type: "code", language: "php", text: '<?php $x = "café"; // 日本語\n$emoji = "🎉";' },
    ],
  },

  /**
   * The empty-vs-absent case the conformance policy calls for: a zero, an
   * empty string and an empty list in optional positions, which is exactly
   * what a serializer with `omitempty` semantics drops on the floor.
   */
  emptyValues: {
    title: "",
    metadata: META,
    blocks: [
      { type: "paragraph", runs: [{ text: "" }] },
      { type: "paragraph", runs: [{ text: "0" }] },
    ],
  },
};

/**
 * Parts where the two writers are KNOWN to differ, with why.
 *
 * Every one of these was found by running this suite for the first time. None
 * is a rendering difference — Word opens both files identically — but the
 * engines are demonstrably not emitting the same document, and that is worth
 * being able to see rather than assume away.
 *
 * **Deliberately not "fixed" here.** Reconciling them means changing what one
 * engine writes, which is consumer-visible output, and in two of the three
 * cases the PORT is arguably the more correct one — so "make TS match PHP",
 * the obvious move given PHP is the reference, would remove the better
 * behaviour. That is a decision for the owner of the pair, not a side effect of
 * writing the suite that found it.
 *
 * The list ratchets both ways: a NEW divergent part fails, and an entry that
 * stops being true fails too, so it cannot rot into a permanent excuse.
 */
const KNOWN_DIVERGENT_PARTS: Record<string, string> = {
  /**
   * THE SERIOUS ONE. This is the document body — not boilerplate — so the two
   * engines are emitting materially different documents for the same input.
   * Three distinct causes, and they drift in BOTH directions, which rules out
   * "one engine is simply ahead":
   *
   *   hyperlink  the port emits w:history="1"; PHP does not
   *   list       PHP emits <w:pStyle w:val="ListParagraph"/>; the port does not
   *   table      the port references <w:tblStyle w:val="LastWordTable"/>; PHP
   *              inlines <w:tblBorders> instead
   *
   * The table difference is the one a reader can SEE: borders coming from a
   * named style versus borders written into the table properties. Combined with
   * the styles.xml divergence below, a table written by one backend and a table
   * written by the other are not the same object.
   *
   * Recorded rather than fixed because choosing between a named style and inline
   * borders changes rendered output for every existing consumer, and that is the
   * pair owner's call. It is the headline finding of this suite's first run.
   */
  "word/document.xml":
    "Lists, tables and hyperlinks are built differently by the two engines — see " +
    "the block comment above. User-visible for tables.",

  "[Content_Types].xml":
    "PHP declares Default extensions for png/jpeg unconditionally; the port only " +
    "declares what the document actually contains. Both are valid OOXML — Word " +
    "tolerates an unused Default — so this is a question of which is tidier, not " +
    "which works.",

  "word/styles.xml":
    'The port emits w:eastAsia="Calibri" alongside w:ascii/w:hAnsi/w:cs in the ' +
    "default run properties; PHP omits it. The port is arguably more correct: " +
    "without an eastAsia font Word picks its own for CJK runs, which is exactly " +
    "the text the `unicode` case above exercises.",

  "word/numbering.xml":
    '<w:multiLevelType w:val="hybridMultilevel"/> is present in the port and absent ' +
    "in PHP. Word itself writes it, so again the port looks closer to what a real " +
    "document contains; PHP's lists still render, they are just less explicitly typed.",
};

const HAS_PHP = phpAvailable();

/**
 * In CI a missing PHP is a FAILURE, not a skip.
 *
 * `skipIf(!HAS_PHP)` on its own is how the other two suites in this family
 * reported success with zero cross-engine coverage for months — the workflow
 * installed Node only, every parity test skipped, and the build went green. A
 * green build asserting nothing is worse than a red one, because nobody
 * investigates green.
 */
if (process.env.CI && !HAS_PHP) {
  throw new Error(
    "php is not on PATH. This suite is the cross-engine parity guarantee; " +
      "skipping it in CI would report success with no coverage. Install PHP, " +
      "or set LAST_WORD_PHP_SRC and ensure `php` resolves.",
  );
}

describe.skipIf(!HAS_PHP)("cross-engine writer parity (PHP vs TS)", () => {
  let dir: string;

  beforeAll(() => {
    dir = mkdtempSync(join(tmpdir(), "last-word-parity-"));
  });

  const writeBoth = (name: string, doc: unknown) => {
    const docFile = join(dir, `${name}.json`);
    const phpOut = join(dir, `${name}.php.docx`);
    writeFileSync(docFile, JSON.stringify(doc));
    php([PHP_SCRIPT, docFile, phpOut]);

    return {
      phpParts: unzipSync(new Uint8Array(readFileSync(phpOut))),
      tsParts: unzipSync(Agent.toBytes(doc)),
    };
  };

  for (const [name, doc] of Object.entries(DOCS)) {
    it(`emits the same OOXML parts: ${name}`, () => {
      const { phpParts, tsParts } = writeBoth(name, doc);

      expect(Object.keys(tsParts).sort(), "the two engines wrote different part sets").toEqual(
        Object.keys(phpParts).sort(),
      );

      const dec = new TextDecoder();
      const differing: string[] = [];

      for (const part of Object.keys(phpParts)) {
        const a = dec.decode(phpParts[part]!);
        const b = dec.decode(tsParts[part]!);
        if (a === b) continue;

        differing.push(part);
        expect(
          KNOWN_DIVERGENT_PARTS,
          `NEW divergence in ${part} for "${name}" — the engines have drifted apart somewhere ` +
            `that was previously identical. Fix it, or add it to KNOWN_DIVERGENT_PARTS with a reason.`,
        ).toHaveProperty(part);
      }

      // The floor, and the reason `minimal` exists: whatever else diverges, a
      // plain paragraph must be identical. If this ever fails, the engines have
      // stopped agreeing about the simplest possible document and nothing else
      // in this file is worth reading.
      if (name === "minimal") {
        expect(
          differing,
          "the engines disagree on the BODY of a single plain paragraph — the boilerplate " +
            "parts diverge everywhere (see KNOWN_DIVERGENT_PARTS), but this one must not",
        ).not.toContain("word/document.xml");
      }
    });
  }

  it("the known-divergence list is accurate and only shrinks", () => {
    // A stale entry is worse than none: it would keep excusing a part that has
    // since been reconciled, and quietly re-open the hole if it regressed.
    const { phpParts, tsParts } = writeBoth("ledger", DOCS.structure);
    const dec = new TextDecoder();

    const actuallyDiffering = Object.keys(phpParts).filter(
      (part) => dec.decode(phpParts[part]!) !== dec.decode(tsParts[part]!),
    );

    expect(actuallyDiffering.sort(), "KNOWN_DIVERGENT_PARTS no longer matches reality").toEqual(
      Object.keys(KNOWN_DIVERGENT_PARTS).sort(),
    );
  });

  it("compared something — the loop is not empty", () => {
    // The guard the other suites lacked. If DOCS emptied, or unzip returned
    // nothing, every assertion above would vanish and this file would still
    // report success.
    const { phpParts } = writeBoth("guard", DOCS.minimal);

    expect(Object.keys(DOCS).length).toBeGreaterThan(3);
    expect(Object.keys(phpParts).length).toBeGreaterThan(3);
    expect(Object.keys(phpParts)).toContain("word/document.xml");
  });

  it("each engine reads the other's file to the same document", () => {
    // Parts matching is the strict check; this is the one that matters to a
    // consumer, and it catches a shared writer bug that a part diff cannot —
    // if both engines emit the same WRONG bytes, they still agree here, but a
    // reader disagreeing with the writer shows up.
    const { tsParts } = writeBoth("roundtrip", DOCS.inlineMarks);
    void tsParts;

    const phpOut = join(dir, "roundtrip.php.docx");
    const fromPhp = Agent.read(new Uint8Array(readFileSync(phpOut)));
    const fromTs = Agent.read(Agent.toBytes(DOCS.inlineMarks));

    expect(fromTs).toEqual(fromPhp);
  });
});

if (!HAS_PHP) {
  // eslint-disable-next-line no-console
  console.warn("[parity] php not found on PATH — cross-engine writer parity skipped (fails in CI).");
}
