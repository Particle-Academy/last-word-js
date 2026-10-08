import { describe, expect, it } from "vitest";
import { Agent, TemplateException, unzipSync } from "../src";
import { generatedTemplate, templateWithoutStyles, TEMPLATE_FONT, TEMPLATE_HEADING_COLOR, TEMPLATE_ACCENT1 } from "./support/generated-template";

/**
 * Rendering onto a house template — `{ template }`.
 *
 * The Node twin of PHP's `TemplateBindingTest` (last-word#3). `toBytes()` took no
 * options at all, so every document came out in the built-in look and an
 * automation producing customer-facing collateral still needed a human to
 * re-apply the house style.
 *
 * **Bind by style name**: the document model already uses Word's own style ids,
 * so a template defining `Normal`, `Heading1..n`, `Quote` and `Hyperlink` binds by
 * carrying its own `word/styles.xml` and `word/theme/theme1.xml`.
 *
 * The theme travels WITH the styles on purpose. A style that says "the major
 * heading font, accent 1" resolves against whatever theme is in the package, so
 * taking styles alone would give the template's structure in the default's
 * colours — a wrong answer that looks deliberate.
 *
 * ONE DIFFERENCE FROM THE PHP ENGINE, and it is pre-existing rather than
 * introduced here: this engine does not emit `<w:pStyle w:val="ListParagraph"/>`
 * on list paragraphs and defines no `ListParagraph` or `Title` style, where PHP
 * does both (recorded in `parity.test.ts`'s KNOWN_DIVERGENT_PARTS). So a
 * template's List Paragraph and Title styles bind in PHP and have nothing to bind
 * to here. Reported upstream; reconciling it changes rendered output for existing
 * consumers in whichever engine moves, which is the pair owner's call.
 */
type Any = Record<string, unknown>;

function templateDoc(): Any {
  return {
    title: "Strawman Business Case",
    blocks: [
      { type: "heading", level: 1, runs: [{ text: "Executive Summary" }] },
      // An inline `code` run is what emits the InlineCode character style.
      {
        type: "paragraph",
        runs: [{ text: "The case rests on " }, { text: "three", code: true }, { text: " things." }],
      },
      // A heading level the generated template deliberately omits.
      { type: "heading", level: 3, runs: [{ text: "A level the template omits" }] },
      { type: "list", items: [{ runs: [{ text: "One" }] }, { runs: [{ text: "Two" }] }] },
      { type: "quote", blocks: [{ type: "paragraph", runs: [{ text: "Their words, not ours." }] }] },
      { type: "code", language: "php", text: "echo 'hello';" },
    ],
  };
}

const DECODER = new TextDecoder();

function partsOf(bytes: Uint8Array): Record<string, string> {
  const raw = unzipSync(bytes);
  const out: Record<string, string> = {};
  for (const [name, data] of Object.entries(raw)) out[name] = DECODER.decode(data);
  return out;
}

const W3C_FONT = TEMPLATE_FONT;

describe("rendering onto a house template", () => {
  it("writes the built-in look when no template is given", () => {
    // The regression guard. This feature is additive and must not move a byte for
    // the callers who do not use it.
    const parts = partsOf(Agent.toBytes(templateDoc()));

    expect(parts["word/styles.xml"]).toContain("Calibri");
    expect(parts).not.toHaveProperty("word/theme/theme1.xml");
  });

  it("renders onto the template styles instead of its own", () => {
    const parts = partsOf(Agent.toBytes(templateDoc(), { template: generatedTemplate() }));

    expect(parts["word/styles.xml"]).toContain(W3C_FONT);
    expect(parts["word/styles.xml"]).toContain(TEMPLATE_HEADING_COLOR);
    expect(parts["word/styles.xml"]).not.toContain("Calibri");

    // And its own styles survive untouched, rather than being filtered to the
    // ones this writer happens to recognise.
    expect(parts["word/styles.xml"]).toContain("HouseNote");
  });

  it("carries the template theme, declared in the content types and rels", () => {
    const parts = partsOf(Agent.toBytes(templateDoc(), { template: generatedTemplate() }));

    expect(parts).toHaveProperty("word/theme/theme1.xml");
    expect(parts["word/theme/theme1.xml"]).toContain(TEMPLATE_ACCENT1);

    // A part that is present but undeclared makes the package invalid, and Word
    // reports that as "unreadable content" rather than naming the part.
    expect(parts["[Content_Types].xml"]).toContain(
      '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
    );
    expect(parts["word/_rels/document.xml.rels"]).toContain("theme/theme1.xml");
  });

  it("supplies definitions for the styles it emits and the template lacks", () => {
    const parts = partsOf(Agent.toBytes(templateDoc(), { template: generatedTemplate() }));
    const styles = parts["word/styles.xml"]!;

    // Anchored to the template first: every assertion below is also true of the
    // built-in styles.xml, so without this the test would pass whether or not a
    // template was applied.
    expect(styles).toContain(W3C_FONT);

    expect(styles).toContain('w:styleId="CodeBlock"');
    expect(styles).toContain('w:styleId="InlineCode"');
    expect(styles).toContain('w:styleId="Heading3"');

    // And it must NOT re-define what the template already has, or the duplicate
    // wins by document order and silently overrides the house style.
    expect(styles.match(/w:styleId="Heading1"/g)).toHaveLength(1);
    expect(styles.match(/w:styleId="Normal"/g)).toHaveLength(1);
    expect(styles.match(/w:styleId="Quote"/g)).toHaveLength(1);
    expect(styles.match(/w:styleId="Hyperlink"/g)).toHaveLength(1);
  });

  it("keeps its own numbering so lists still resolve", () => {
    const parts = partsOf(Agent.toBytes(templateDoc(), { template: generatedTemplate() }));

    expect(parts).toHaveProperty("word/numbering.xml");
    expect(parts["word/document.xml"]).toContain("<w:numPr>");
  });

  it("is deterministic with a template, as without one", () => {
    const template = generatedTemplate();

    expect(Agent.toBytes(templateDoc(), { template })).toEqual(
      Agent.toBytes(templateDoc(), { template }),
    );
  });

  it("works with a template that ships no theme", () => {
    const parts = partsOf(Agent.toBytes(templateDoc(), { template: generatedTemplate(false) }));

    expect(parts["word/styles.xml"]).toContain(W3C_FONT);
    expect(parts).not.toHaveProperty("word/theme/theme1.xml");
    expect(parts["[Content_Types].xml"]).not.toContain("theme1.xml");
  });

  it("refuses a template it cannot use rather than falling back", () => {
    // Falling back to the built-in look would reproduce the exact complaint this
    // feature answers. A host can also use this to validate a customer-supplied
    // template at upload.
    expect(() => Agent.toBytes(templateDoc(), { template: new TextEncoder().encode("not-a-package") }))
      .toThrow(TemplateException);

    expect(() => Agent.toBytes(templateDoc(), { template: templateWithoutStyles() }))
      .toThrow(/no word\/styles\.xml/);
  });

  it("refuses a path through toBytes, which is universal and has no filesystem", () => {
    // Saying so beats a confusing zip error. `write()` accepts a path.
    expect(() => Agent.toBytes(templateDoc(), { template: "/templates/house.dotx" }))
      .toThrow(/needs the template as bytes/);
  });

  it("treats an empty template as absent", () => {
    // A host passing through a blank form field gets the built-in look rather
    // than an exception.
    const parts = partsOf(Agent.toBytes(templateDoc(), { template: new Uint8Array() }));

    expect(parts["word/styles.xml"]).toContain("Calibri");
  });
});
