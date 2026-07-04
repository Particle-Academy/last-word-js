import { readFileSync } from "node:fs";
import { describe, it, expect } from "vitest";
import { Agent, zipSync } from "../src";
import phpCanonical from "../test/fixtures/php-canonical.json";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

/**
 * Cross-language parity (last-word-js#1): the frozen fixture
 * `php-canonical.docx` was written by the PHP mirror
 * (particle-academy/last-word) from `php-canonical.json` — this suite
 * asserts the Node reader restores it, metadata slots included (title from
 * docProps/core.xml, code language from the `lastword:code:{lang}` sdt tag).
 */

const phpDocx = () =>
  new Uint8Array(readFileSync(new URL("../test/fixtures/php-canonical.docx", import.meta.url)));

describe("cross-read: PHP-written docx (frozen fixture)", () => {
  it("restores the title from docProps/core.xml", () => {
    const doc = Agent.read(phpDocx());
    expect(doc.title).toBe((phpCanonical as Any).title);
    expect(doc.title).toBe("Last Word Canonical");
  });

  it("restores every code block language from the sdt tag", () => {
    const doc: Any = Agent.read(phpDocx());
    const languages = doc.blocks
      .filter((b: Any) => b.type === "code")
      .map((b: Any) => b.language);
    const expected = (phpCanonical as Any).blocks
      .filter((b: Any) => b.type === "code")
      .map((b: Any) => b.language);
    expect(languages).toEqual(expected);
    expect(languages).toEqual(["typescript"]);
  });

  it("restores the exact block-type sequence", () => {
    const doc: Any = Agent.read(phpDocx());
    expect(doc.blocks.map((b: Any) => b.type)).toEqual(
      (phpCanonical as Any).blocks.map((b: Any) => b.type),
    );
  });

  it("recovers the full PHP canonical doc semantically (deep-equal)", () => {
    // The reader's run-merge normalization is the suite's semantic
    // normalizer; the PHP canonical has no mergeable adjacent runs, so the
    // read result must deep-equal the fixture exactly.
    expect(Agent.read(phpDocx())).toEqual(phpCanonical);
  });
});

describe("cross-read: PHP ≤0.1.x legacy code-language bookmark", () => {
  it("reads LastWordCode_{lang} bookmarks as the code language (back-compat)", () => {
    // Exactly the shape the PHP 0.1.0 writer emitted: CodeBlock-styled
    // paragraphs with an invisible bookmark on the first line, no sdt.
    const encoder = new TextEncoder();
    const documentXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:body>` +
      `<w:p><w:pPr><w:pStyle w:val="CodeBlock"/></w:pPr>` +
      `<w:bookmarkStart w:id="1" w:name="LastWordCode_php"/><w:bookmarkEnd w:id="1"/>` +
      `<w:r><w:t xml:space="preserve">echo "hi";</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:pStyle w:val="CodeBlock"/></w:pPr>` +
      `<w:r><w:t xml:space="preserve">exit(0);</w:t></w:r></w:p>` +
      `</w:body></w:document>`;
    const bytes = zipSync([{ name: "word/document.xml", data: encoder.encode(documentXml) }]);

    const doc: Any = Agent.read(bytes);
    expect(doc.blocks).toEqual([
      { type: "code", language: "php", text: 'echo "hi";\nexit(0);' },
    ]);
  });
});
