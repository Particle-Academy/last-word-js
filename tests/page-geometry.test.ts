import { describe, expect, it } from "vitest";
import { Agent, DocxWriter, Schema, unzipSync } from "../src";

/**
 * The page model a paged viewer needs (MOIC, 2026-11-01): every margin, and
 * the header and footer DISTANCES -- `w:pgMar`'s `header` / `footer`, how far
 * the header sits from the top edge and the footer from the bottom edge.
 * Mirrors PHP last-word 0.15.0 `tests/Unit/PageGeometryTest.php`.
 *
 * The model keeps its rule that a value equal to the writer's default is not
 * surfaced (a side at 72pt, a distance at 36pt), so every document read before
 * this release reads the same. `Schema.pageMargins()` resolves a document's
 * page to all six numbers, defaults filled in, and is what a viewer should
 * call rather than re-deriving the defaults.
 */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

function pgDoc(margins: Record<string, unknown>): Any {
  return { page: { margins }, blocks: [{ type: "paragraph", runs: [{ text: "Body" }] }] };
}

function documentXml(doc: Any): string {
  return new TextDecoder().decode(unzipSync(Agent.toBytes(doc))["word/document.xml"]!);
}

describe("page geometry: header and footer distances", () => {
  it("publishes the header and footer distances under page.margins", () => {
    const schema = Agent.jsonSchema() as Any;
    const ref: string = schema.properties.page.properties.margins.$ref;
    const margins = schema.definitions[ref.slice("#/definitions/".length)];

    expect(Object.keys(margins.properties)).toEqual(["top", "right", "bottom", "left", "header", "footer"]);
    expect(margins.properties.header.minimum).toBe(0);
    expect(margins.properties.footer.minimum).toBe(0);
  });

  it("writes the distances into w:pgMar", () => {
    expect(documentXml(pgDoc({ bottom: 54, header: 18, footer: 27 }))).toContain(
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1080" w:left="1440" w:header="360" w:footer="540" w:gutter="0"/>',
    );
  });

  it("reads a non-default bottom margin and both distances back", () => {
    const doc = pgDoc({ bottom: 54, header: 18, footer: 27 });

    expect(Agent.read(Agent.toBytes(doc)).page).toStrictEqual({ margins: { bottom: 54, header: 18, footer: 27 } });
  });

  it("reads every side and both distances when none is the default", () => {
    const margins = { top: 90, left: 63, bottom: 81, right: 45, header: 22.5, footer: 40 };

    expect((Agent.read(Agent.toBytes(pgDoc(margins))) as Any).page.margins).toEqual(margins);
  });

  it("omits defaults exactly as before, so existing documents read unchanged", () => {
    const read = Agent.read(Agent.toBytes(pgDoc({ top: 72, bottom: 72, header: 36, footer: 36 })));

    expect(read).not.toHaveProperty("page");
  });

  it("resolves every margin and both distances for a viewer, defaults included", () => {
    expect(Schema.pageMargins({ blocks: [] })).toStrictEqual({ top: 72, right: 72, bottom: 72, left: 72, header: 36, footer: 36 });
    expect(Schema.pageMargins(Agent.read(Agent.toBytes(pgDoc({ bottom: 54, footer: 27 }))))).toStrictEqual({
      top: 72,
      right: 72,
      bottom: 54,
      left: 72,
      header: 36,
      footer: 27,
    });
    expect(Schema.PAGE_MARGIN_DEFAULT).toBe(72);
    expect(Schema.PAGE_BAND_DISTANCE_DEFAULT).toBe(36);
  });

  for (const value of [-1, "18pt"]) {
    it(`refuses a distance that is not a non-negative number: ${JSON.stringify(value)}`, () => {
      expect(Agent.validate(pgDoc({ header: value })).map((e) => e.path)).toContain("/page/margins/header");
    });
  }

  it("lets the document own a distance (no template section is taken in this engine)", () => {
    // Without a document value the 720-twip default stands; with one, the
    // document's value is written, field by field, as for the four sides.
    expect(documentXml(pgDoc({ footer: 12 }))).toContain('w:header="720" w:footer="240"');
  });

  it("writes an unvalidated negative distance as 0, never a negative twip count", () => {
    // ST_TwipsMeasure is unsigned. Validation refuses it; the writer clamps.
    const xml = new TextDecoder().decode(unzipSync(new DocxWriter().toBytes(pgDoc({ header: -5 })))["word/document.xml"]!);
    expect(xml).toContain('w:header="0" w:footer="720"');
  });
});
