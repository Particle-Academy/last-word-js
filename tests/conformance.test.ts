import { describe, it, expect } from "vitest";
import { loadSuite, runTable, formatSummary, deepEquals } from "@particle-academy/fancy-conformance";
import type { CaseResult, ConformanceCase, RunSummary } from "@particle-academy/fancy-conformance";

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
 * What lives here is only the extractors: how to get from `toBytes(doc)` to
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

/** Every part of a written document, by name, as text. */
function writtenParts(doc: unknown): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [name, bytes] of Object.entries(unzipSync(Agent.toBytes(doc)))) out[name] = new TextDecoder().decode(bytes);
  return out;
}

function partRoot(xml: string | undefined): XmlNode | null {
  return xml === undefined ? null : parseXml(xml);
}

/** The `footnotesPart` projection, as the manifest defines it. */
function footnotesPart(doc: unknown): unknown {
  const parts = writtenParts(doc);

  let contentType: string | null = null;
  for (const o of collect(partRoot(parts["[Content_Types].xml"]), "Override")) {
    if (at(o, "PartName") === "/word/footnotes.xml") contentType = at(o, "ContentType") ?? null;
  }
  let relationship: string | null = null;
  for (const rel of collect(partRoot(parts["word/_rels/document.xml.rels"]), "Relationship")) {
    if ((at(rel, "Type") ?? "").endsWith("/footnotes")) relationship = at(rel, "Target") ?? null;
  }
  const settings = partRoot(parts["word/settings.xml"]);
  const footnotePr = settings ? el(settings, "footnotePr") : undefined;

  const root = partRoot(parts["word/footnotes.xml"]);
  const footnotes = (root?.children ?? [])
    .filter((node) => node.name === "footnote")
    .map((node) => ({
      id: at(node, "id") ?? "",
      type: at(node, "type") ?? null,
      text: runText(node),
      content: normProps(node),
    }));

  return {
    contentType,
    relationship,
    footnotePr: footnotePr === undefined ? null : normProps(footnotePr),
    references: collect(documentXml(doc), "footnoteReference").map((r) => at(r, "id") ?? ""),
    footnotes,
  };
}

/**
 * A row's whole input (fancy-conformance 0.41.0): `{doc, template?, options?}`.
 * `template` is ITSELF a document model, written with toBytes() and passed as
 * the template write option; `options` are further write options, as given.
 * This engine has no `templateParts` / `templateFirstPageOnly` (PHP 0.13.0 /
 * 0.16.0), so it ignores those as it ignores any unknown option -- the rows
 * that need them are in NODE_SKIPS below, each with its reason.
 */
type RowInput = { doc: unknown; template?: unknown; options?: Record<string, unknown> };

function writeInput(input: RowInput): Uint8Array {
  const options: Record<string, unknown> = { ...(input.options ?? {}) };
  if (input.template !== undefined && input.template !== null) options.template = Agent.toBytes(input.template as never);
  return Agent.toBytes(input.doc as never, options as never);
}

function partsOf(bytes: Uint8Array): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [name, data] of Object.entries(unzipSync(bytes))) out[name] = new TextDecoder().decode(data);
  return out;
}

/** The first child element with this local name. */
function child(node: XmlNode | null | undefined, name: string): XmlNode | undefined {
  return node?.children.find((c) => c.name === name);
}

/** Header/footer relationship id => target, from the document's rels. */
function bandTargets(parts: Record<string, string>): Record<string, string> {
  const out: Record<string, string> = {};
  for (const rel of collect(partRoot(parts["word/_rels/document.xml.rels"]), "Relationship")) {
    if (/\/(header|footer)$/.test(at(rel, "Type") ?? "")) out[at(rel, "Id") ?? ""] = at(rel, "Target") ?? "";
  }
  return out;
}

function bodySectPr(parts: Record<string, string>): XmlNode | undefined {
  return child(child(partRoot(parts["word/document.xml"]), "body"), "sectPr");
}

const isBandReference = (node: XmlNode): boolean => node.name === "headerReference" || node.name === "footerReference";

/** The `bandParts` projection, as the manifest defines it. */
function bandParts(input: RowInput): unknown {
  const parts = partsOf(writeInput(input));
  const sectPr = bodySectPr(parts);
  const targets = bandTargets(parts);
  const references = (sectPr?.children ?? []).filter(isBandReference).map((node) => {
    const target = targets[at(node, "id") ?? ""] ?? null;
    const xml = target === null ? undefined : parts[`word/${target}`];
    return {
      kind: node.name.slice(0, -"Reference".length),
      type: at(node, "type") ?? "",
      target,
      text: xml === undefined ? null : collect(partRoot(xml), "t").map((t) => t.text).join(""),
    };
  });
  return {
    titlePg: child(sectPr, "titlePg") !== undefined,
    evenAndOddHeaders: child(partRoot(parts["word/settings.xml"]), "evenAndOddHeaders") !== undefined,
    references,
    parts: Object.keys(parts).filter((name) => /^word\/(header|footer)[^/]*\.xml$/.test(name)).sort(),
  };
}

/** Every complex field in a part, in document order, as {instruction, result}. */
function complexFields(root: XmlNode | null): { instruction: string; result: string }[] {
  const out: { instruction: string; result: string }[] = [];
  let state: "instruction" | "result" | null = null;
  let instruction = "";
  let result = "";
  const walk = (node: XmlNode): void => {
    for (const el of node.children) {
      if (el.name === "fldChar") {
        const type = at(el, "fldCharType");
        if (type === "begin") [state, instruction, result] = ["instruction", "", ""];
        else if (type === "separate") state = "result";
        else if (type === "end" && state !== null) {
          out.push({ instruction, result });
          state = null;
        }
      } else if (el.name === "instrText" && state === "instruction") {
        instruction += el.text;
      } else if (el.name === "t" && state === "result") {
        result += el.text;
      }
      walk(el);
    }
  };
  if (root) walk(root);
  return out;
}

/** The `bandFields` projection: each referenced part once, in reference order. */
function bandFields(input: RowInput): unknown {
  const parts = partsOf(writeInput(input));
  const targets = bandTargets(parts);
  const seen = new Set<string>();
  const out: { target: string; fields: unknown }[] = [];
  for (const node of (bodySectPr(parts)?.children ?? []).filter(isBandReference)) {
    const target = targets[at(node, "id") ?? ""];
    if (target === undefined || seen.has(target)) continue;
    seen.add(target);
    out.push({ target, fields: complexFields(partRoot(parts[`word/${target}`])) });
  }
  return out;
}

/** The `coreProps` projection: null without docProps/core.xml. */
function coreProps(input: RowInput): unknown {
  const parts = partsOf(writeInput(input));
  const xml = parts["docProps/core.xml"];
  if (xml === undefined) return null;
  let contentType: string | null = null;
  for (const o of collect(partRoot(parts["[Content_Types].xml"]), "Override")) {
    if (at(o, "PartName") === "/docProps/core.xml") contentType = at(o, "ContentType") ?? null;
  }
  let relationship: string | null = null;
  for (const rel of collect(partRoot(parts["_rels/.rels"]), "Relationship")) {
    if ((at(rel, "Type") ?? "").endsWith("/core-properties")) relationship = at(rel, "Target") ?? null;
  }
  // The qualified names, which this package's parser strips to local names:
  // every start tag of the raw part in order, the root's dropped (core
  // properties are flat, so the rest are its children), so a wrong prefix is a
  // failure rather than normalised away.
  const qualified = [...xml.matchAll(/<([A-Za-z_][\w.:-]*)[\s/>]/g)].map((m) => m[1]!).slice(1);
  const children = partRoot(xml)?.children ?? [];
  expect(qualified.length, "core properties are expected to be flat").toBe(children.length);
  const elements = children.map((el, i) => [qualified[i]!, el.text]);
  return { contentType, relationship, elements };
}

/** Functions that take the row's WHOLE input (fancy-conformance 0.41.0). */
const INPUT_EXTRACTORS: Record<string, (input: RowInput) => unknown> = {
  bandParts,
  bandFields,
  coreProps,
  readBack: (input) => Agent.read(writeInput(input)),
  // The comparator is the suite's own, so "equal" means the same thing here as
  // it does for every other row.
  roundTripFixpoint: (input) => ({ fixpoint: deepEquals(Agent.read(writeInput(input)), input.doc) }),
};

/**
 * Rows this engine cannot pass yet, by id, each with WHY. The table has no
 * `skip.node` for them (the reference shipped them for all three engines), so
 * the skip is declared here and applied to the summary below -- printed by
 * formatSummary with its reason, never folded into a pass. Each one must still
 * FAIL: a row that starts passing fails the build until it is taken off this
 * list, so the list can only shrink.
 */
const TEMPLATE_PARTS =
  "needs templateParts (PHP 0.13.0, last-word#5): this engine takes a template's styles and theme only, never its " +
  "header/footer section, so templateFirstPageOnly (PHP 0.16.0) has nothing to act on";
const BLOCKS_BAND =
  "writes a {blocks} header/footer (PHP 0.14.0, last-word#6), which this engine does not write; the one-line " +
  "form of the same contract is covered by the other rows";
const NODE_SKIPS: Record<string, string> = {
  "0054-template-letterhead-every-page": TEMPLATE_PARTS,
  "0055-template-first-page-only-false": TEMPLATE_PARTS,
  "0056-template-first-page-only": TEMPLATE_PARTS,
  "0057-template-first-page-only-later-pages": TEMPLATE_PARTS,
  "0058-template-first-page-only-uses-its-first-part": TEMPLATE_PARTS,
  "0059-template-first-page-only-document-first-wins": TEMPLATE_PARTS,
  "0060-read-back-template-first-page-only": TEMPLATE_PARTS,
  "0061-first-header-parts": BLOCKS_BAND,
  "0064-read-back-first-and-even": BLOCKS_BAND,
  "0065-fixpoint-first-and-even-bands": BLOCKS_BAND,
  "0069-field-tokens-in-a-blocks-band": BLOCKS_BAND,
};

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

  footnotesPart: (doc) => footnotesPart(doc),
};

function extract(c: ConformanceCase): unknown {
  const byInput = INPUT_EXTRACTORS[c.fn ?? ""];
  if (byInput) return byInput(c.input as RowInput);
  const fn = EXTRACTORS[c.fn ?? ""];
  if (!fn) throw new Error(`no extractor for fn "${c.fn}"`);
  return fn((c.input as { doc: unknown }).doc);
}

/**
 * The summary with NODE_SKIPS applied: a listed row that failed becomes a skip
 * carrying its reason, and the counts and verdict are recomputed. A listed row
 * that PASSED is left as it is and reported by the caller, so a skip never
 * outlives the gap it records.
 */
function applySkips(summary: RunSummary): { summary: RunSummary; nowPassing: string[] } {
  const nowPassing: string[] = [];
  const results: CaseResult[] = summary.results.map((r) => {
    const reason = NODE_SKIPS[r.id];
    if (reason === undefined) return r;
    if (r.status === "pass") {
      nowPassing.push(r.id);
      return r;
    }
    return r.status === "fail" ? { id: r.id, title: r.title, status: "skip", reason: `node: ${reason}` } : r;
  });
  const count = (s: CaseResult["status"]): number => results.filter((r) => r.status === s).length;
  const failed = count("fail");
  return {
    summary: { ...summary, results, passed: count("pass"), failed, skipped: count("skip"), ok: failed === 0 },
    nowPassing,
  };
}

/**
 * The last-word half of `doc-model/page-refusal`: the same messages as PHP.
 * Messages only: Node's paths are JSON Pointers, and the suite says so.
 */
describe("fancy-conformance: doc-model/page-refusal", () => {
  it("satisfies the last-word half of every row", () => {
    const { cases } = loadSuite("doc-model/page-refusal");
    const failed: string[] = [];
    for (const row of cases) {
      const doc = (row.input as { doc: unknown }).doc;
      const want = (row.expected as { lastWord: { messages?: string[]; repairOk: boolean } }).lastWord;
      const got: { repairOk: boolean; messages?: string[] } = { repairOk: Agent.validateAndRepair(doc).ok };
      if (want.messages !== undefined) got.messages = Agent.validate(doc).map((e) => e.message);
      if (!deepEquals(got, want)) failed.push(`${row.id} expected ${JSON.stringify(want)} got ${JSON.stringify(got)}`);
    }
    expect(failed).toEqual([]);
    expect(cases.length).toBeGreaterThanOrEqual(18);
  });
});

describe("fancy-conformance: last-word/docx-constructs", () => {
  it("runs every row in the shared table", () => {
    const { summary, nowPassing } = applySkips(runTable(SUITE, extract, { language: "node" }));
    expect(nowPassing, "these rows PASS now: take them off NODE_SKIPS").toEqual([]);
    // Every skip in the log, by name and reason, on a green run too: "11
    // skipped" folded into a pass is indistinguishable from full coverage.
    // eslint-disable-next-line no-console
    if (summary.skipped > 0) console.log(formatSummary(summary));
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
      expect([...Object.keys(EXTRACTORS), ...Object.keys(INPUT_EXTRACTORS)], `row declares fn "${fn}" with no extractor`).toContain(fn);
    }
  });

  it("every skip names a row that exists, says why, and is the gap it claims", () => {
    const { cases } = loadSuite(SUITE);
    const byId = new Map(cases.map((c) => [c.id, c]));
    for (const [id, reason] of Object.entries(NODE_SKIPS)) {
      const row = byId.get(id);
      expect(row, `NODE_SKIPS names ${id}, which is not in the table`).toBeDefined();
      expect(reason.trim(), `${id}: a skip without a reason`).not.toBe("");
      // A template skip must be a row that carries a template; a blocks-band
      // skip a row whose document has a {blocks} band. Nothing else may hide here.
      const input = row!.input as RowInput;
      if (reason === TEMPLATE_PARTS) expect(input.template, `${id} is skipped for templateParts but has no template`).toBeDefined();
      else expect(JSON.stringify((input.doc as { page?: unknown }).page ?? {}), `${id} is skipped for a {blocks} band but has none`).toContain('"blocks"');
    }
    // 0.41.0's rows, 26 of them: the skips are a minority, and the rest run.
    const added = cases.filter((c) => c.id >= "0054" && c.id < "0080");
    expect(added.length).toBe(26);
    expect(added.filter((c) => !(c.id in NODE_SKIPS)).length).toBe(15);
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
