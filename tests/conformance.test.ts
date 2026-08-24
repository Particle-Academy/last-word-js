import { describe, it, expect } from "vitest";
import { loadSuite, runTable, formatSummary, deepEquals } from "@particle-academy/fancy-conformance";
import type { ConformanceCase } from "@particle-academy/fancy-conformance";

import { Agent, unzipSync } from "../src";
import { parseXml, el, els, at, type XmlNode } from "../src/reader/xml";

/**
 * The shared construct table — `last-word/docx-constructs` in
 * `fancy-conformance`.
 *
 * The rows are NOT transcribed here. This package, its PHP twin and its Python
 * twin all assert the same file, so a mapping that drifts in one engine fails
 * in that engine rather than quietly becoming that engine's behaviour. Adding a
 * construct means adding a row there, once.
 *
 * What lives here is only the six extractors: how to get from `toBytes(doc)` to
 * the value a row compares. They are deliberately thin — a normaliser clever
 * enough to paper over a difference is a normaliser that stops the suite
 * finding one.
 */

const SUITE = "last-word/docx-constructs";

/**
 * An ordered normalisation of one property container.
 *
 * ORDER IS THE POINT: CT_RPr, CT_PPr, CT_TcPr, CT_TblPr and CT_SectPr are all
 * xsd:sequence, so a map keyed by element name would let two engines emit
 * different XML and still agree here. Attribute order is not pinned, because
 * attributes are unordered in XML.
 */
type Props = Array<[string, true | Record<string, string> | Props]>;

function normProps(node: XmlNode | undefined): Props {
  if (!node) return [];
  return node.children.map((child): [string, true | Record<string, string> | Props] => {
    if (child.children.length > 0) return [child.name, normProps(child)];
    return [child.name, Object.keys(child.attrs).length > 0 ? child.attrs : true];
  });
}

/** Every element with this local name, in document order. */
function collect(root: XmlNode | null, name: string): XmlNode[] {
  const found: XmlNode[] = [];
  const walk = (node: XmlNode): void => {
    if (node.name === name) found.push(node);
    for (const child of node.children) walk(child);
  };
  if (root) walk(root);
  return found;
}

function documentXml(doc: unknown): XmlNode | null {
  const parts = unzipSync(Agent.toBytes(doc));
  const part = parts["word/document.xml"];
  if (!part) throw new Error("no word/document.xml in the archive");
  return parseXml(new TextDecoder().decode(part));
}

function runText(run: XmlNode): string {
  return collect(run, "t").map((t) => t.text).join("");
}

const EXTRACTORS: Record<string, (doc: unknown) => unknown> = {
  runProps: (doc) =>
    collect(documentXml(doc), "r").map((r) => ({ text: runText(r), rPr: normProps(el(r, "rPr")) })),

  paragraphProps: (doc) => collect(documentXml(doc), "p").map((p) => normProps(el(p, "pPr"))),

  tableProps: (doc) =>
    collect(documentXml(doc), "tbl").map((tbl) => ({
      tblPr: normProps(el(tbl, "tblPr")),
      grid: els(el(tbl, "tblGrid"), "gridCol").map((g) => at(g, "w")),
    })),

  cellProps: (doc) => collect(documentXml(doc), "tc").map((tc) => normProps(el(tc, "tcPr"))),

  sectionProps: (doc) => normProps(el(el(documentXml(doc), "body"), "sectPr")),

  readBack: (doc) => Agent.read(Agent.toBytes(doc)),

  // The comparator is the suite's own, so "equal" means the same thing here as
  // it does for every other row.
  roundTripFixpoint: (doc) => ({ fixpoint: deepEquals(Agent.read(Agent.toBytes(doc)), doc) }),
};

function extract(c: ConformanceCase): unknown {
  const fn = EXTRACTORS[c.fn ?? ""];
  if (!fn) throw new Error(`no extractor for fn "${c.fn}"`);
  return fn((c.input as { doc: unknown }).doc);
}

describe("fancy-conformance: last-word/docx-constructs", () => {
  it("runs every row in the shared table", () => {
    const summary = runTable(SUITE, extract, { language: "node" });
    // The whole summary, not just a count: a failure names the row and prints
    // both sides, which is the difference between "the suite is red" and
    // knowing which construct moved.
    expect(summary.ok, `\n${formatSummary(summary)}\n`).toBe(true);
  });

  it("compared something — the loop is not empty", () => {
    // The guard `parity.test.ts` documents the need for. Without it, an empty
    // or unloadable table reports success over zero assertions, which is worse
    // than a red build because nobody investigates green.
    const { cases } = loadSuite(SUITE);
    expect(cases.length).toBeGreaterThan(40);

    const fns = new Set(cases.map((c) => c.fn));
    for (const fn of fns) {
      expect(Object.keys(EXTRACTORS), `row declares fn "${fn}" with no extractor`).toContain(fn);
    }
  });

  it("the extractors discriminate — a wrong document fails its own row", () => {
    // The control. Every extractor above is a projection, and a projection
    // that returns a constant passes every row it is given. This proves the
    // one the acceptance case leans on can return the other answer, and that a
    // property genuinely reaches the XML rather than the row being vacuous.
    const shaded = EXTRACTORS.cellProps!({
      blocks: [{ type: "table", rows: [{ cells: [{ blocks: [], shading: "#123456" }] }] }],
    });
    const plain = EXTRACTORS.cellProps!({
      blocks: [{ type: "table", rows: [{ cells: [{ blocks: [] }] }] }],
    });
    expect(shaded).not.toEqual(plain);
  });
});
