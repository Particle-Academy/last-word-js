import { unzipSync } from "../zip";
import { TemplateException } from "../exceptions";

const DECODER = new TextDecoder();

/**
 * A `.dotx` / `.docx` opened for the parts that carry its LOOK.
 *
 * The Node twin of PHP's `LastWord\Writer\DocxTemplate` (last-word#3). Every
 * document came out in the built-in look, so an automation producing a
 * structurally correct `.docx` still needed a human to re-apply the house style.
 *
 * ## What is taken, and why only this
 *
 * A house template's look lives almost entirely in two parts:
 *
 *   - `word/styles.xml`        the definitions — fonts, sizes, colours, spacing,
 *                              and `w:docDefaults` for everything unstyled
 *   - `word/theme/theme1.xml`  the colour and font scheme those definitions
 *                              reference by name (`majorHAnsi`, `accent1`, …)
 *
 * Taking the styles WITHOUT the theme is the trap worth naming: a style saying
 * "the major heading font, accent 1" resolves against whatever theme ships in
 * the package, so the document would come out in the template's *structure* and
 * the default's *colours* — a wrong answer that looks deliberate.
 *
 * ## What is deliberately NOT taken
 *
 *   - **`word/numbering.xml`.** `document.xml` references `w:numId`s this
 *     package's numbering part defines; a template's would repoint every list.
 *   - **`w:sectPr`** — page size, margins, headers, footers. It lives in
 *     `document.xml`, which the writer owns.
 *   - **`word/settings.xml`** — mostly `w:rsid` revision junk, which would make
 *     output depend on a template's editing history.
 *
 * The styles part is passed through as bytes. Word's own serialisation is
 * already correct, and re-emitting it from a parser would be a chance to be
 * wrong for no gain. The only thing read out of it is WHICH style ids it
 * defines, so the writer can supply its own for the ones the template lacks.
 *
 * Universal: it uses this package's own inflate, so it works in a browser.
 */
export class DocxTemplate {
  private constructor(
    private readonly stylesXml: string,
    private readonly themeXml: string | null,
    private readonly ids: ReadonlySet<string>,
  ) {}

  /**
   * Open a template from raw bytes.
   *
   * Refuses rather than degrading. Falling back to the built-in look reproduces
   * the exact complaint this feature answers — a document that silently comes
   * out in the wrong style, with nothing said.
   */
  static open(bytes: Uint8Array): DocxTemplate {
    let parts: Record<string, Uint8Array>;
    try {
      parts = unzipSync(bytes);
    } catch (cause) {
      throw new TemplateException(
        "The template is not a readable .dotx/.docx package (it did not open as a zip).",
        cause,
      );
    }

    const styles = parts["word/styles.xml"];
    if (styles === undefined || styles.length === 0) {
      throw new TemplateException(
        "The template has no word/styles.xml, so it defines no styles to bind to.",
      );
    }

    const stylesXml = DECODER.decode(styles);
    const theme = parts["word/theme/theme1.xml"];

    return new DocxTemplate(
      stylesXml,
      theme === undefined || theme.length === 0 ? null : DECODER.decode(theme),
      idsIn(stylesXml),
    );
  }

  /** The template's `word/theme/theme1.xml`, or null when it ships none. */
  theme(): string | null {
    return this.themeXml;
  }

  hasTheme(): boolean {
    return this.themeXml !== null;
  }

  /** Whether the template defines a paragraph/character style with this id. */
  defines(styleId: string): boolean {
    return this.ids.has(styleId);
  }

  /**
   * The template's styles with `styleXml` inserted before `</w:styles>`.
   *
   * A string splice rather than a parse-and-reserialise: re-emitting a template
   * through a DOM reorders attributes and rewrites namespace prefixes, and the
   * result still has to be a part Word accepts — so the fewer bytes of someone
   * else's file we touch, the better.
   */
  stylesWith(styleXml: string): string {
    if (styleXml === "") return this.stylesXml;

    const close = this.stylesXml.lastIndexOf("</w:styles>");
    if (close === -1) {
      throw new TemplateException(
        "The template's word/styles.xml has no closing </w:styles> element.",
      );
    }

    return this.stylesXml.slice(0, close) + styleXml + this.stylesXml.slice(close);
  }
}

/**
 * `w:styleId` values declared in the part.
 *
 * A regex on purpose: the question is only "is this id already taken", the input
 * is a part Word wrote, and a wrong answer costs a duplicate definition rather
 * than a wrong document. Mirrors the PHP engine exactly.
 */
function idsIn(stylesXml: string): Set<string> {
  const ids = new Set<string>();
  const pattern = /<w:style\b[^>]*\sw:styleId="([^"]*)"/g;

  for (const match of stylesXml.matchAll(pattern)) {
    ids.add(decodeEntities(match[1]!));
  }

  return ids;
}

function decodeEntities(value: string): string {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}
