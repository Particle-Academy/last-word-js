import { describe, it, expect } from "vitest";
import { Agent, unzipSync } from "../src";

/**
 * Can this package produce a document someone would be PROUD to send?
 *
 * The Node half of `last-word`'s `PremiumDocumentTest`, on the same document.
 * The suites beside it prove each feature works ON ITS OWN and that the bytes
 * are stable. Neither asks whether you can turn all of it on at once and get a
 * document with flair rather than a wall of Calibri.
 *
 * A docx is a pile of parts that reference each other by id — a numbered list is
 * a `<w:numPr>` pointing at a definition in `numbering.xml`, a heading is a
 * `<w:pStyle>` pointing at `styles.xml` — and every one of those links can be
 * written wrong in a way that still opens.
 *
 * **A dropped feature is invisible.** Word shows no error; the document is
 * merely plain. Nobody files a bug against a report that looks boring — they
 * conclude the library is boring. So these assert the ARTIFACT: unzip, and look
 * for the feature in the XML.
 */

const dec = new TextDecoder();

const cell = (text: string) => ({ blocks: [{ type: "paragraph", runs: [{ text }] }] });

/** Everything a rich document uses, in one file. */
const premiumDoc = () => ({
  title: "Q3 Revenue Review",
  defaultFont: "Inter",
  defaultSize: 11,
  page: {
    size: "a4",
    orientation: "landscape",
    // POINTS, per the schema's own `boxSides` description — not twips.
    // 72pt is an inch (1440 twips); 54pt is three quarters (1080).
    margins: { top: 72, bottom: 72, left: 54, right: 54 },
  },
  blocks: [
    { type: "heading", level: 1, runs: [{ text: "Q3 Revenue Review" }] },
    { type: "heading", level: 2, runs: [{ text: "By region" }] },

    // Every run flag the schema has, in ONE paragraph. They land in a single
    // `<w:rPr>` whose child order OOXML fixes, so this is where a mis-ordered
    // or overwritten property shows up.
    {
      type: "paragraph",
      align: "justify",
      runs: [
        { text: "Bold ", bold: true },
        { text: "italic ", italic: true },
        { text: "underlined ", underline: true },
        { text: "struck ", strike: true },
        { text: "small caps ", smallCaps: true },
        { text: "red ", color: "#C00000" },
        { text: "large ", size: 18, font: "Playfair Display" },
        { text: "highlighted", highlight: "#FFFF00" },
      ],
    },

    {
      type: "table",
      widths: [2, 1, 1],
      rows: [
        { header: true, cells: [cell("Region"), cell("Revenue"), cell("Growth")] },
        { cells: [cell("North"), cell("$1,250,000.50"), cell("18.4%")] },
        { cells: [cell("EMEA"), cell("$2,100,000.00"), cell("31.1%")] },
      ],
    },

    {
      type: "list",
      ordered: true,
      items: [
        { runs: [{ text: "EMEA led on growth" }] },
        { runs: [{ text: "North led on absolute revenue" }] },
      ],
    },
    { type: "list", ordered: false, items: [{ runs: [{ text: "A bulleted note" }] }] },

    {
      type: "quote",
      blocks: [{ type: "paragraph", runs: [{ text: "The quarter turned on EMEA." }] }],
    },
    { type: "code", text: "const growth = 0.311;", language: "javascript" },
    { type: "hr" },
    { type: "pageBreak" },
    { type: "paragraph", runs: [{ text: "Appendix" }] },
  ],
});

/** Read one part out of a written document. */
const partOf = (doc: unknown, name: string): string => {
  // Fail loudly on an invalid fixture rather than asserting against a document
  // the writer never got to see.
  expect(Agent.validate(doc as never)).toEqual([]);

  const parts = unzipSync(Agent.toBytes(doc as never));
  const part = parts[name];
  // A missing part and an empty one are different failures, and the difference
  // is the whole diagnosis: absent means the writer never emitted it.
  expect(part, `document has no ${name}`).toBeDefined();
  return dec.decode(part!);
};

describe("a rich document keeps EVERY formatting feature, together", () => {
  it("keeps all eight run properties in one paragraph", () => {
    // The composition check. Every one of these is a child of the same
    // `<w:rPr>`, whose order OOXML fixes — so this is where a property written
    // in the wrong slot, or overwritten by the next one, shows up. The
    // per-feature tests each write one run and cannot see it.
    const body = partOf(premiumDoc(), "word/document.xml");

    const expected: Record<string, string> = {
      bold: "<w:b/>",
      italic: "<w:i/>",
      underline: "<w:u ",
      strike: "<w:strike/>",
      "small caps": "<w:smallCaps/>",
      "red text": "C00000",
      "custom font": "Playfair Display",
      "larger size": 'w:val="36"',
    };

    const missing = Object.entries(expected)
      .filter(([, needle]) => !body.includes(needle))
      .map(([label]) => label);

    expect(
      missing,
      `these run properties were accepted and never reached document.xml: ${missing.join(", ")}`,
    ).toEqual([]);
  });

  it("shades a highlight rather than using the 16-colour w:highlight", () => {
    // Pinned deliberately, because it looks like a gap and is a decision.
    // `<w:highlight>` takes SIXTEEN NAMED COLOURS and nothing else, so a schema
    // accepting `#RRGGBB` cannot render through it without rejecting most values
    // or snapping them to the nearest of sixteen. `<w:shd>` takes any hex.
    const body = partOf(premiumDoc(), "word/document.xml");

    expect(body, "the highlight colour never reached the document").toContain("FFFF00");
    expect(
      body,
      "if this now fails, highlight moved to w:highlight — check what happens to a colour outside the 16",
    ).not.toContain("<w:highlight");
  });

  it("makes headings real Word styles, so the navigation pane works", () => {
    // A "heading" that is only a large bold paragraph produces a document with
    // no outline: no navigation pane, no automatic table of contents, and
    // nothing for a screen reader to jump between. It looks identical.
    const body = partOf(premiumDoc(), "word/document.xml");
    const styles = partOf(premiumDoc(), "word/styles.xml");

    expect(body, "h1 is not a styled heading").toContain('<w:pStyle w:val="Heading1"/>');
    expect(body, "h2 is not a styled heading").toContain('<w:pStyle w:val="Heading2"/>');
    expect(styles, "Heading1 is referenced but never defined").toContain('w:styleId="Heading1"');
    expect(styles, "headings carry no outline level").toContain("<w:outlineLvl");
  });

  it("numbers an ordered list through numbering.xml, not literal text", () => {
    // A list faked with typed numbers renumbers wrong the moment anyone inserts
    // an item, and is the most common way a generated document betrays itself.
    const body = partOf(premiumDoc(), "word/document.xml");
    const numbering = partOf(premiumDoc(), "word/numbering.xml");

    expect(body, "list items carry no numbering reference").toContain("<w:numPr>");
    expect(numbering, "no decimal numbering format defined").toContain("decimal");
    expect(numbering, "no bullet numbering format defined").toContain("bullet");

    // The reference has to RESOLVE. A numId pointing at nothing is the classic
    // silent break: Word drops the list formatting entirely.
    const numId = /<w:numId w:val="(\d+)"\/>/.exec(body)?.[1];
    expect(numId, "no numId on any list paragraph").toBeDefined();
    expect(
      numbering,
      `list references numId ${numId}, which numbering.xml never defines`,
    ).toContain(`<w:num w:numId="${numId}"`);
  });

  it("repeats the table header on every page it breaks across", () => {
    // `<w:tblHeader/>` is a docx-only capability — PPTX has no pagination at
    // all — and it is the difference between a long table that stays readable
    // and one whose column labels vanish after page one.
    const body = partOf(premiumDoc(), "word/document.xml");

    expect(body, "the header row does not repeat").toContain("<w:trPr><w:tblHeader/></w:trPr>");
  });

  it("fixes the table layout so the requested column widths hold", () => {
    // Without `w:tblLayout fixed`, Word refits columns to their content and the
    // widths become advisory — a 2:1:1 table silently renders even.
    const body = partOf(premiumDoc(), "word/document.xml");

    expect(body, "column widths are advisory").toContain('<w:tblLayout w:type="fixed"/>');
    expect(body, "the table has no borders").toContain("<w:tblBorders>");
  });

  it("sets the page up as A4 landscape with the margins it was given", () => {
    const body = partOf(premiumDoc(), "word/document.xml");

    expect(body, "orientation ignored").toContain('w:orient="landscape"');
    // A4 landscape is 16838 x 11906 twips. Landscape means the two are SWAPPED,
    // not merely flagged — a document flagged landscape at portrait dimensions
    // prints wrong.
    expect(body, "page width is not A4 landscape").toContain('w:w="16838"');
    expect(body, "margins ignored").toContain('w:left="1080"');
  });

  it("puts the default font in docDefaults, where it governs the whole document", () => {
    // Set anywhere else, it applies to the runs the writer happened to touch and
    // nothing else — so a document looks right until someone types in it.
    const styles = partOf(premiumDoc(), "word/styles.xml");

    expect(styles, "no docDefaults block").toContain("<w:docDefaults>");
    expect(styles, "the default font never reached the style table").toContain("Inter");
  });
});

describe("the guard against a feature that is accepted and dropped", () => {
  it("proves the assertions can FAIL — a plain document has none of it", () => {
    // Without this, every assertion above could be passing on boilerplate that
    // appears in any document, and the suite would be green for a file with no
    // formatting whatsoever. This is the control.
    const plain = { blocks: [{ type: "paragraph", runs: [{ text: "just text" }] }] };

    const body = partOf(plain, "word/document.xml");

    expect(body, "a plain document somehow contains the display font").not.toContain(
      "Playfair Display",
    );
    expect(body, "a plain document somehow repeats a table header").not.toContain("<w:tblHeader/>");
    expect(body, "a plain document somehow contains a list").not.toContain("<w:numPr>");
    expect(body, "a plain document is somehow landscape").not.toContain("landscape");
  });
});
