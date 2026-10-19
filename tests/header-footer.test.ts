import { describe, expect, it } from "vitest";
import { Agent, unzipSync } from "../src";

/**
 * Running headers and footers with live page numbers. Mirrors the PHP engine's
 * HeaderFooterTest; byte parity with PHP is in parity.test.ts ("bands").
 */
type Any = Record<string, unknown>;

const doc = (page: Any): Any => ({ title: "Bands", page, blocks: [{ type: "paragraph", runs: [{ text: "Body" }] }] });

function parts(d: Any): Record<string, string> {
  const files = unzipSync(Agent.toBytes(d as never));
  const out: Record<string, string> = {};
  for (const [name, bytes] of Object.entries(files)) out[name] = new TextDecoder().decode(bytes);
  return out;
}

describe("running header and footer", () => {
  it("writes a footer with live page fields and wires it into the section", () => {
    const p = parts(doc({ footer: { left: "Q3 Board Memo | Confidential", right: "Page {page} of {pages}" } }));

    expect(p["word/footer1.xml"]).toContain("Q3 Board Memo | Confidential");
    expect(p["word/footer1.xml"]).toContain('<w:instrText xml:space="preserve"> PAGE </w:instrText>');
    expect(p["word/footer1.xml"]).toContain('<w:instrText xml:space="preserve"> NUMPAGES </w:instrText>');
    expect(p["word/document.xml"]).toContain('<w:footerReference w:type="default" r:id="rIdFooter1"/>');
    expect(p["word/_rels/document.xml.rels"]).toContain('Target="footer1.xml"');
    expect(p["[Content_Types].xml"]).toContain('PartName="/word/footer1.xml"');
  });

  it("adds nothing when no band is asked for", () => {
    const p = parts(doc({ margins: { top: 54 } }));

    expect(p["word/footer1.xml"]).toBeUndefined();
    expect(p["word/document.xml"]).not.toContain("Reference");
  });

  it("reads its own bands back as it wrote them", () => {
    const page = {
      header: { right: "Board of Directors" },
      footer: { left: "Q3 Board Memo", center: "Confidential", right: "Page {page} of {pages}" },
    };

    expect((Agent.read(Agent.toBytes(doc(page) as never)) as Any).page).toEqual(page);
  });
});
