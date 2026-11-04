/**
 * DocxReader — .docx bytes → Doc JSON model. Handles this package's own
 * writer output losslessly (round-trip), the PHP mirror's output (same
 * metadata slots since 0.2.0, plus the legacy `LastWordCode_{lang}` bookmark
 * fallback for PHP ≤0.1.x files) and tolerates Word-authored files:
 * headings via pStyle Heading1-9 OR outlineLvl, numPr lists with ilvl nesting,
 * hyperlinks via rels, images via blip r:embed, page breaks, bottom-border-only
 * paragraphs → hr. Unknown constructs degrade to plain paragraphs, never throw.
 * Mirrors PHP `Reader\DocxReader`.
 */

import { base64Encode } from "../helpers/image-size";
import type { Block, Doc, ListItem, Run } from "../schema/types";
import { SourceArchive } from "../zip/source-archive";
import { TEMPLATE, sourceKind } from "../zip/office-template";
import { decodePart } from "./part-xml";
import { at, el, els, parseXml, type XmlNode } from "./xml";
import { Schema } from "../schema/schema";
import { EMU_PER_PX, SDT_TAG_CODE, SDT_TAG_QUOTE, splitColumns, PAGE_SIZES } from "../writer/docx-writer";
import { NS_C, readChart } from "./chart-reader";
import { DEFAULT_ASPECT } from "../helpers/chart";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;


interface RelInfo {
  target: string;
  external: boolean;
}

/** Named w:highlight values → hex (reader tolerance for Word-authored files). */
const HIGHLIGHT_NAMES: Record<string, string> = {
  yellow: "#FFFF00",
  green: "#00FF00",
  cyan: "#00FFFF",
  magenta: "#FF00FF",
  blue: "#0000FF",
  red: "#FF0000",
  darkBlue: "#00008B",
  darkCyan: "#008B8B",
  darkGreen: "#006400",
  darkMagenta: "#8B008B",
  darkRed: "#8B0000",
  darkYellow: "#808000",
  darkGray: "#A9A9A9",
  lightGray: "#D3D3D3",
  black: "#000000",
  white: "#FFFFFF",
};

const MONO_FONTS = ["consolas", "courier new", "courier", "menlo", "monaco", "source code pro"];

const ORDERED_NUM_FMTS = [
  "decimal",
  "decimalZero",
  "upperRoman",
  "lowerRoman",
  "upperLetter",
  "lowerLetter",
];

interface FlatListItem {
  numId: string;
  ilvl: number;
  runs: Run[];
}

interface WalkCtx {
  insideQuote?: boolean;
  /** The body, or a wrapper it reads through: a section can end here. */
  topLevel?: boolean;
}

/** What ending a section does to the blocks: see `planSections()`. */
interface SectionPlan {
  break: Record<string, unknown> | null;
  width: number;
  dropped: number;
}

/** Marks where a section ended, among the blocks of a walk, until the walk is finished. */
const SECTION_END = "#section-end";

export class DocxReader {
  private parts: Record<string, Uint8Array> = {};
  private rels: Record<string, RelInfo> = {};
  private numOrdered: Record<string, boolean> = {};
  /** The section's text column in twips. */
  private contentWidth = 9360;

  /**
   * The document's footnotes, by `w:id`, from the part its `footnotes`
   * relationship names: the `w:footnote` elements of its own (separators and
   * continuation notices are not footnotes).
   */
  private footnoteNodes = new Map<string, XmlNode>();

  /** The footnotes part's OWN relationships, which resolve its links and images. */
  private footnoteRels: Record<string, RelInfo> = {};

  /** A footnote's blocks once read, by id: a footnote referenced twice is read once. */
  private footnoteCache = new Map<string, Block[]>();

  /** True while a footnote's blocks are read: a reference inside one is ignored. */
  private inFootnote = false;

  private sectionPlan: SectionPlan[] = [];

  private sectionCursor = 0;

  /** The band part a section's reference types name, inherited from the section before. */
  private sectionTargets: Record<string, Record<string, string>> = { header: {}, footer: {} };

  read(bytes: Uint8Array): Doc {
    const source = SourceArchive.open(bytes);
    const kind = sourceKind(source);
    this.parts = Object.create(null) as Record<string, Uint8Array>;
    for (const name of source.names()) this.parts[name] = source.read(name)!;
    this.rels = this.loadRels("word/_rels/document.xml.rels");
    this.inFootnote = false;
    this.footnoteCache = new Map();
    this.loadFootnotes();
    this.evenAndOddHeaders = this.readEvenAndOdd();
    this.numOrdered = this.loadNumbering();

    const docXml = this.partText("word/document.xml");
    const root = docXml ? parseXml(docXml) : null;
    const body = el(root, "body");

    // Top-level `page` is the first section, and every later one a break holding what
    // differs from the page before it. Read before any block, so a chart that spans
    // the column -- the writer's default -- reads back without a size.
    const { sections, nested } = body ? sectionsOf(body) : { sections: [undefined], nested: 0 };
    this.sectionPlan = [];
    this.sectionCursor = 0;
    this.sectionTargets = { header: {}, footer: {} };
    this.contentWidth = contentWidthOf(sections[0]);
    const page = this.readPage(sections[0]);
    this.planSections(sections, page ?? {});
    const unrepresented: { where: string; kind: "layout-not-preserved"; count: number }[] = [];

    const doc: Doc = { blocks: body ? finishSections(this.walkBody(body.children, { topLevel: true }), unrepresented) : [] };

    const title = this.readTitle();
    if (title !== null) doc.title = title;
    // meta.author (PHP 0.16.0): dc:creator, beside the title it shares a part with.
    const author = this.readCoreElement("creator");
    if (author !== null && author !== "") doc.meta = { author };

    // Section geometry and document defaults are surfaced ONLY when they
    // differ from what the writer produces unasked, for the same reason table
    // options are: a Letter portrait page at one-inch margins is what every
    // document written before `page` existed contains.
    if (page !== null) doc.page = page as Any;
    const defaults = this.readDocDefaults();
    if (defaults.font !== null) (doc as Any).defaultFont = defaults.font;
    if (defaults.size !== null) (doc as Any).defaultSize = defaults.size;
    // Never document text: locations and counts only, like the exchange report.
    const dropped = nested + droppedByFirstSection(sections[0]);
    if (dropped > 0) unrepresented.push({ where: "/page", kind: "layout-not-preserved", count: dropped });
    if (unrepresented.length > 0) doc.unrepresented = unrepresented;
    // A .dotx reads exactly like its .docx twin, and says so.
    if (kind === TEMPLATE) (doc as Any).sourceKind = TEMPLATE;

    return doc;
  }

  /**
   * What ending each section does to the blocks. `w:type` says how a section STARTS, so
   * the break between sections k-1 and k is decided by section k's type; absent is
   * nextPage. A continuous section that does not break leaves the model's page as it
   * was. Mirrors PHP `DocxReader::planSections()`.
   */
  private planSections(sections: (XmlNode | undefined)[], first: Record<string, unknown>): void {
    let current = first;
    for (const sectPr of sections.slice(1)) {
      let next = this.readPage(sectPr) ?? {};
      const knownPaper = paperKnown(sectPr);
      const change = pageChange(current, next, knownPaper);
      const type = at(el(sectPr, "type"), "val");
      const breaks = type !== "continuous" || change.orientation !== undefined || change.size !== undefined;
      let dropped = 0;
      if (!breaks) {
        dropped += 1;
      } else {
        dropped += type === undefined || type === "nextPage" || type === "continuous" ? 0 : 1;
        dropped += columns(sectPr) > 1 ? 1 : 0;
        if (!knownPaper) {
          next = { ...next };
          delete next.size;
          if (current.size !== undefined) next.size = current.size;
        }
        current = next;
      }
      dropped += restarts(sectPr, true) ? 1 : 0;
      dropped += knownPaper ? 0 : 1;
      this.sectionPlan.push({ break: breaks ? change : null, width: contentWidthOf(sectPr), dropped });
    }
  }

  /** The section the top-level walk has just passed the end of, as a marker block. */
  private endSection(): Block {
    const plan = this.sectionPlan[this.sectionCursor++];
    if (plan !== undefined) this.contentWidth = plan.width;
    return { type: SECTION_END, plan } as unknown as Block;
  }

  private readPage(sectPr: XmlNode | undefined): Record<string, unknown> | null {
    if (!sectPr) return null;
    const out: Record<string, unknown> = {};

    const pgSz = el(sectPr, "pgSz");
    let w = numAttr(pgSz, "w") ?? 12240;
    let h = numAttr(pgSz, "h") ?? 15840;
    if (at(pgSz, "orient") === "landscape") {
      out.orientation = "landscape";
      [w, h] = [h, w];
    }
    for (const [name, dims] of Object.entries(PAGE_SIZES)) {
      if (dims[0] === w && dims[1] === h && name !== "letter") out.size = name;
    }

    const pgMar = el(sectPr, "pgMar");
    const margins: Record<string, number> = {};
    // Each value only when it differs from what the writer puts there unasked
    // -- 1440 twips a side, 720 to the header and the footer -- so a document
    // read before the distances were read reads the same. Schema.pageMargins()
    // fills the defaults back in for a viewer.
    for (const side of [...BOX_EDGES, "header", "footer"] as const) {
      const value = numAttr(pgMar, side);
      const fallback = side === "header" || side === "footer" ? 720 : 1440;
      if (value !== null && value !== fallback) margins[side] = points(value);
    }
    if (Object.keys(margins).length > 0) out.margins = margins;

    Object.assign(out, this.readBands(sectPr));

    return Object.keys(out).length > 0 ? out : null;
  }

  /**
   * The section's headers and footers, by model key, in `Schema.PAGE_BANDS`
   * order: `header` / `footer` from the DEFAULT references, and (PHP 0.16.0)
   * `firstHeader` / `firstFooter` from the FIRST ones while `w:titlePg` is on,
   * `evenHeader` / `evenFooter` from the EVEN ones while settings turn on
   * `w:evenAndOddHeaders`. A switch that is off makes its references nothing,
   * as in Word.
   *
   * A variant that names the same part as the default is not repeated (the
   * writer references the default for page one when only one kind has a
   * first-page band). One with no part, or an empty part, while the default
   * has content, is `{"blocks": []}`: that page shows no header in Word, and
   * the model says so rather than letting a write-back put the default there.
   * Mirrors PHP `DocxReader::parseBands()`.
   */
  private readBands(sectPr: XmlNode): Record<string, unknown> {
    const titlePgEl = el(sectPr, "titlePg");
    const titlePg = titlePgEl !== undefined && toggleOn(at(titlePgEl, "val"));
    // A reference a section leaves out is inherited from the section before.
    for (const kind of ["header", "footer"] as const) {
      for (const type of ["default", "first", "even"]) {
        const own = this.bandTarget(sectPr, kind, type);
        if (own !== null) this.sectionTargets[kind]![type] = own;
      }
    }
    const found: Record<string, Record<string, unknown>> = {};
    for (const kind of ["header", "footer"] as const) {
      const def = this.sectionTargets[kind]!.default ?? null;
      const band = def !== null ? this.readBandPart(def) : null;
      if (band !== null) found[kind] = band;
      for (const [type, on] of [
        ["first", titlePg],
        ["even", this.evenAndOddHeaders],
      ] as const) {
        if (!on) continue;
        const key = `${type}${kind === "header" ? "Header" : "Footer"}`;
        const target = this.sectionTargets[kind]![type] ?? null;
        if (target !== null && target === def) continue;
        let variant: Record<string, unknown> | null = target !== null ? this.readBandPart(target) : null;
        if (variant === null && band !== null) variant = { blocks: [] };
        if (variant !== null) found[key] = variant;
      }
    }

    const out: Record<string, unknown> = {};
    for (const key of Schema.PAGE_BANDS) {
      if (found[key] !== undefined) out[key] = found[key];
    }
    return out;
  }

  /**
   * The part name (`word/header1.xml`) a section's header or footer reference
   * of this type names, or null. The last reference of a type wins, as it
   * always has for the default; a reference with no type is the default.
   */
  private bandTarget(sectPr: XmlNode, kind: "header" | "footer", type: string): string | null {
    let partName: string | null = null;
    for (const node of sectPr.children) {
      if (node.name === `${kind}Reference` && (at(node, "type") ?? "default") === type) {
        const target = this.rels[at(node, "id") ?? ""]?.target;
        partName = typeof target === "string" ? "word/" + target.replace(/\\/g, "/").replace(/^\/+/, "") : null;
      }
    }
    return partName;
  }

  /** `w:evenAndOddHeaders` in word/settings.xml: even pages have their own header and footer. */
  private evenAndOddHeaders = false;

  /** Whether the settings part the document's rels name turns on even/odd headers. A DOCTYPE leaves it off. */
  private readEvenAndOdd(): boolean {
    const relsXml = this.partText("word/_rels/document.xml.rels");
    const target = relsXml !== null ? relTargetOfType(relsXml, "settings") : null;
    if (target === null) return false;
    const name = target.startsWith("/") ? target.replace(/^\/+/, "") : "word/" + target.split("\\").join("/").replace(/^\/+/, "");
    const xml = this.partText(name);
    if (xml === null) return false;
    const node = el(parseXml(xml), "evenAndOddHeaders");
    return node !== undefined && toggleOn(at(node, "val"));
  }

  /**
   * One header or footer part as a one-line band: left / center / right text
   * with its field tokens read back (`{page}`, `{date:yyyy-MM-dd}`, ...). Null
   * when the part is missing or holds no text. Mirrors the slot half of PHP
   * DocxReader::parseBandPart(): tab-separated segments become
   * left/center/right; a field that is a token becomes it and its cached result
   * is skipped, any other field keeps its cached result as text.
   */
  private readBandPart(partName: string): Record<string, unknown> | null {
    const xml = this.partText(partName);
    const root = xml ? parseXml(xml) : null;
    const p = el(root, "p");
    if (!p) return null;

    const segments: string[] = [""];
    let firstRun: XmlNode | null = null;
    let field: string | false | null = null; // inside a complex field: its instruction so far
    const append = (s: string): void => {
      segments[segments.length - 1] += s;
    };
    const walk = (node: XmlNode): void => {
      for (const child of node.children) {
        const name = child.name;
        if (name === "fldSimple") {
          const token = fieldToken(at(child, "instr") ?? "");
          if (token !== "") append(token);
          else walk(child); // not a token: its cached result is the text
        } else if (name === "r") {
          firstRun ??= child;
          walk(child);
        } else if (name === "fldChar") {
          const type = at(child, "fldCharType");
          if (type === "begin") {
            field = "";
          } else if (type === "separate" && typeof field === "string") {
            const token = fieldToken(field);
            append(token);
            // A token's cached result ("1", the author) is skipped; any other
            // field's is its text, as in PHP's blocks band.
            field = token !== "" ? false : null;
          } else if (type === "end") {
            field = null;
          }
        } else if (name === "instrText" && typeof field === "string") {
          field += child.text;
        } else if (field === null && name === "t") {
          append(child.text);
        } else if (field === null && name === "tab") {
          segments.push("");
        } else if (name === "hyperlink" || name === "smartTag") {
          walk(child);
        }
      }
    };
    walk(p);

    let slots: Record<string, string>;
    if (segments.length === 1) {
      const jc = at(el(el(p, "pPr"), "jc"), "val") ?? "";
      const slot = jc === "center" ? "center" : jc === "right" || jc === "end" ? "right" : "left";
      slots = { [slot]: segments[0] ?? "" };
    } else if (segments.length === 2) {
      slots = { left: segments[0] ?? "", right: segments[1] ?? "" };
    } else {
      slots = { left: segments[0] ?? "", center: segments[1] ?? "", right: segments.slice(2).join(" ") };
    }

    const band: Record<string, unknown> = {};
    for (const slot of ["left", "center", "right"]) {
      if ((slots[slot] ?? "") !== "") band[slot] = slots[slot];
    }
    if (Object.keys(band).length === 0) return null;

    const rPr = el(firstRun ?? undefined, "rPr");
    const sz = Number(at(el(rPr, "sz"), "val"));
    if (Number.isInteger(sz) && at(el(rPr, "sz"), "val") !== undefined && sz !== 18) band.size = sz / 2;
    const color = (at(el(rPr, "color"), "val") ?? "").toUpperCase();
    if (/^[0-9A-F]{6}$/.test(color) && color !== "6B7280") band.color = "#" + color;

    return band;
  }

  private readDocDefaults(): { font: string | null; size: number | null } {
    const xml = this.partText("word/styles.xml");
    if (!xml) return { font: null, size: null };
    const rPr = el(el(el(parseXml(xml), "docDefaults"), "rPrDefault"), "rPr");

    const ascii = at(el(rPr, "rFonts"), "ascii");
    const sz = numAttr(el(rPr, "sz"), "val");

    return {
      font: ascii !== undefined && ascii !== "" && ascii !== "Calibri" ? ascii : null,
      size: sz !== null && sz !== 22 ? sz / 2 : null,
    };
  }

  // ── Parts / metadata ─────────────────────────────────────────────────────

  private partText(name: string): string | null {
    const part = this.parts[name];
    return part ? decodePart(part) : null;
  }

  private readTitle(): string | null {
    return this.readCoreElement("title");
  }

  /**
   * One Dublin Core element's text from docProps/core.xml (`title`,
   * `creator`), or null when the part is missing, declares a DOCTYPE, or has
   * no such element. (An empty title still reads as "", as it always has here.)
   */
  private readCoreElement(localName: string): string | null {
    const xml = this.partText("docProps/core.xml");
    if (!xml) return null;
    const node = el(parseXml(xml), localName);
    return node ? node.text : null;
  }

  private loadRels(name: string): Record<string, RelInfo> {
    const xml = this.partText(name);
    if (!xml) return {};
    const root = parseXml(xml);
    const out: Record<string, RelInfo> = {};
    for (const rel of els(root, "Relationship")) {
      const id = at(rel, "Id");
      const target = at(rel, "Target");
      if (!id || !target) continue;
      out[id] = { target, external: at(rel, "TargetMode") === "External" };
    }
    return out;
  }

  /**
   * The footnotes part the document's `footnotes` relationship names, and its
   * own relationships. Sized by the archive's own limits, as document.xml is; a
   * part declaring a DOCTYPE is not read, and its references then read as
   * nothing. Mirrors PHP `DocxReader::loadFootnotes()`.
   */
  private loadFootnotes(): void {
    this.footnoteNodes = new Map();
    this.footnoteRels = {};
    const relsXml = this.partText("word/_rels/document.xml.rels");
    const target = relsXml !== null ? relTargetOfType(relsXml, "footnotes") : null;
    if (target === null) return;
    const name = target.startsWith("/")
      ? target.replace(/^\/+/, "")
      : "word/" + target.split("\\").join("/").replace(/^\/+/, "");
    const xml = this.partText(name);
    if (xml === null) return;
    const root = parseXml(xml);
    if (!root) return;
    for (const node of root.children) {
      const type = at(node, "type");
      if (node.name === "footnote" && (type === undefined || type === "normal")) {
        const id = at(node, "id");
        if (id !== undefined && !this.footnoteNodes.has(id)) this.footnoteNodes.set(id, node);
      }
    }

    const relsName = `word/_rels/${name.split("/").pop()}.rels`;
    const partRels = this.partText(relsName);
    this.footnoteRels = partRels !== null ? this.loadRels(relsName) : {};
  }

  /**
   * The model run for a `w:footnoteReference`: `{text: "", footnote: {blocks}}`,
   * its blocks read by the body's own block reader against the footnotes
   * part's relationships. Null -- nothing in the model -- for a reference inside
   * another footnote (Word allows none), for an id with no footnote, and for a
   * footnote with no content. (A header or footer is read as a band here,
   * which never reaches this.)
   */
  private footnoteRun(id: string): Run | null {
    const node = this.footnoteNodes.get(id);
    if (this.inFootnote || node === undefined) return null;
    if (!this.footnoteCache.has(id)) {
      const rels = this.rels;
      this.rels = this.footnoteRels;
      this.inFootnote = true;
      try {
        // Without the content-free paragraphs, as PHP reads every container:
        // the writer puts the mark in a paragraph of its own before a note that
        // opens with a table, and closes one that ends with a table with an
        // empty one, and neither is the author's.
        this.footnoteCache.set(id, this.walkBody(node.children, {}).filter((block) => !isEmptyParagraph(block)));
      } finally {
        this.rels = rels;
        this.inFootnote = false;
      }
    }
    const blocks = this.footnoteCache.get(id)!;
    // A copy each time: a footnote referenced twice must not be one object
    // shared between two runs of the model.
    return blocks.length === 0 ? null : { text: "", footnote: { blocks: JSON.parse(JSON.stringify(blocks)) as Block[] } };
  }

  private loadNumbering(): Record<string, boolean> {
    const xml = this.partText("word/numbering.xml");
    if (!xml) return {};
    const root = parseXml(xml);
    const abstractOrdered: Record<string, boolean> = {};
    for (const abs of els(root, "abstractNum")) {
      const id = at(abs, "abstractNumId") ?? "";
      const lvl0 = els(abs, "lvl").find((l) => at(l, "ilvl") === "0") ?? el(abs, "lvl");
      const fmt = at(el(lvl0, "numFmt"), "val") ?? "bullet";
      abstractOrdered[id] = ORDERED_NUM_FMTS.includes(fmt);
    }
    const out: Record<string, boolean> = {};
    for (const num of els(root, "num")) {
      const numId = at(num, "numId") ?? "";
      const absId = at(el(num, "abstractNumId"), "val") ?? "";
      out[numId] = abstractOrdered[absId] ?? false;
    }
    return out;
  }

  // ── Body walking ─────────────────────────────────────────────────────────

  private walkBody(nodes: XmlNode[], ctx: WalkCtx): Block[] {
    const blocks: Block[] = [];
    let listBuf: FlatListItem[] = [];
    let codeBuf: string[] | null = null;
    let codeLang: string | null = null;
    let quoteBuf: Block[] | null = null;

    const flushList = (): void => {
      if (listBuf.length > 0) {
        blocks.push(...this.buildLists(listBuf));
        listBuf = [];
      }
    };
    const flushCode = (): void => {
      if (codeBuf !== null) {
        const block: Block = { type: "code", text: codeBuf.join("\n") };
        if (codeLang !== null) (block as Any).language = codeLang;
        blocks.push(orderCodeKeys(block));
        codeBuf = null;
        codeLang = null;
      }
    };
    const flushQuote = (): void => {
      if (quoteBuf !== null) {
        blocks.push({ type: "quote", blocks: quoteBuf });
        quoteBuf = null;
      }
    };
    const flushAll = (): void => {
      flushList();
      flushCode();
      flushQuote();
    };

    for (const node of nodes) {
      switch (node.name) {
        case "p": {
          const kind = this.classifyParagraph(node);
          const endsSection = ctx.topLevel === true && el(el(node, "pPr"), "sectPr") !== undefined;
          if (endsSection && kind.kind === "blocks" && kind.blocks.every(holdsNothing)) {
            // The paragraph that only carries the section break is not the author's.
            flushAll();
            blocks.push(this.endSection());
            break;
          }
          if (kind.kind === "listItem") {
            flushCode();
            flushQuote();
            listBuf.push(kind.item);
          } else if (kind.kind === "codeLine") {
            flushList();
            flushQuote();
            if (codeBuf === null) {
              codeBuf = [];
              codeLang = kind.language ?? null;
            }
            codeBuf.push(kind.text);
          } else if (kind.kind === "quoteParagraph" && !ctx.insideQuote) {
            flushList();
            flushCode();
            if (quoteBuf === null) quoteBuf = [];
            quoteBuf.push(...kind.blocks);
          } else {
            flushAll();
            blocks.push(...kind.blocks);
          }
          if (endsSection) {
            flushAll();
            blocks.push(this.endSection());
          }
          break;
        }
        case "tbl":
          flushAll();
          // Drop the pad the writer puts BETWEEN adjacent tables — the one
          // that stops Word merging them. Deliberately narrow: only an empty
          // paragraph sandwiched between two tables goes, so a blank line an
          // author actually wrote still survives the read.
          if (blocks.length >= 2 && isEmptyParagraph(blocks[blocks.length - 1]) &&
              (blocks[blocks.length - 2] as Any)?.type === "table") {
            blocks.pop();
          }
          blocks.push(this.parseTable(node, { ...ctx, topLevel: false }));
          break;
        case "sdt": {
          flushAll();
          blocks.push(...this.parseSdt(node, ctx));
          break;
        }
        case "sectPr":
        case "bookmarkStart":
        case "bookmarkEnd":
        case "proofErr":
          break;
        default:
          // Unknown container: recurse looking for readable content (degrade, never throw).
          if (node.children.length > 0) {
            const inner = this.walkBody(node.children, ctx);
            if (inner.length > 0) {
              flushAll();
              blocks.push(...inner);
            }
          }
          break;
      }
    }
    flushAll();
    return blocks;
  }

  private parseSdt(sdt: XmlNode, ctx: WalkCtx): Block[] {
    const tag = at(el(el(sdt, "sdtPr"), "tag"), "val") ?? "";
    const content = el(sdt, "sdtContent");
    if (!content) return [];

    if (tag === SDT_TAG_CODE || tag.startsWith(SDT_TAG_CODE + ":")) {
      const lines = els(content, "p").map((p) => this.plainText(p));
      const block: Block = { type: "code", text: lines.join("\n") };
      if (tag.length > SDT_TAG_CODE.length + 1) {
        (block as Any).language = tag.slice(SDT_TAG_CODE.length + 1);
      }
      return [orderCodeKeys(block)];
    }
    if (tag === SDT_TAG_QUOTE) {
      return [{ type: "quote", blocks: this.walkBody(content.children, { ...ctx, insideQuote: true, topLevel: false }) }];
    }
    // Unknown SDT: read its content transparently.
    return this.walkBody(content.children, ctx);
  }

  // ── Paragraph classification ─────────────────────────────────────────────

  private classifyParagraph(
    p: XmlNode,
  ):
    | { kind: "listItem"; item: FlatListItem }
    | { kind: "codeLine"; text: string; language?: string }
    | { kind: "quoteParagraph"; blocks: Block[] }
    | { kind: "blocks"; blocks: Block[] } {
    const pPr = el(p, "pPr");
    const styleId = at(el(pPr, "pStyle"), "val") ?? "";

    // List item?
    const numPr = el(pPr, "numPr");
    if (numPr) {
      const numId = at(el(numPr, "numId"), "val") ?? "";
      const ilvl = parseInt(at(el(numPr, "ilvl"), "val") ?? "0", 10) || 0;
      return { kind: "listItem", item: { numId, ilvl, runs: this.parseRuns(p, undefined) } };
    }

    // Code line? (SDT-less tolerance path — bare style, or PHP last-word
    // ≤0.1.x output which stashed the language in an invisible
    // `LastWordCode_{lang}` bookmark on the first code paragraph.)
    if (/^(CodeBlock|SourceCode|HTMLPreformatted|Code)$/i.test(styleId)) {
      const language = this.legacyBookmarkLanguage(p);
      return language !== undefined
        ? { kind: "codeLine", text: this.plainText(p), language }
        : { kind: "codeLine", text: this.plainText(p) };
    }

    const blocks: Block[] = [];

    // Images (drawings may sit alongside text runs).
    const drawings = findAll(p, "drawing");
    const runs = this.parseRuns(p, undefined);

    // Heading?
    const headingMatch = /^Heading([1-9])$/i.exec(styleId);
    const outlineLvl = at(el(pPr, "outlineLvl"), "val");
    if (headingMatch || outlineLvl !== undefined) {
      const raw = headingMatch ? parseInt(headingMatch[1]!, 10) : parseInt(outlineLvl ?? "0", 10) + 1;
      const level = Math.min(6, Math.max(1, raw)) as 1 | 2 | 3 | 4 | 5 | 6;
      blocks.push({ type: "heading", level, runs, ...paragraphPropsFrom(pPr) });
      for (const d of drawings) blocks.push(...this.parseDrawing(d));
      return { kind: "blocks", blocks };
    }

    // Page break only?
    const hasPageBreak = findAll(p, "br").some((br) => at(br, "type") === "page");
    const text = runs.map((r) => r.text).join("");
    if (hasPageBreak && text === "" && drawings.length === 0) {
      return { kind: "blocks", blocks: [{ type: "pageBreak" }] };
    }

    // Horizontal rule? (bottom-border-only empty paragraph)
    const pBdr = el(pPr, "pBdr");
    if (pBdr && el(pBdr, "bottom") && text === "" && drawings.length === 0) {
      return { kind: "blocks", blocks: [{ type: "hr" }] };
    }

    // Plain (or quote-styled) paragraph.
    if (text !== "" || drawings.length === 0) {
      const para: Block = { type: "paragraph", runs, ...paragraphPropsFrom(pPr) };
      blocks.push(para);
    }
    for (const d of drawings) blocks.push(...this.parseDrawing(d));

    if (/^(Quote|IntenseQuote|BlockQuote|Blockquote)$/i.test(styleId)) {
      return { kind: "quoteParagraph", blocks };
    }
    return { kind: "blocks", blocks };
  }

  /**
   * Back-compat: the code language from a PHP last-word ≤0.1.x
   * `LastWordCode_{lang}` bookmark on this paragraph, if present. The
   * canonical carrier is the `lastword:code:{lang}` sdt tag.
   */
  private legacyBookmarkLanguage(p: XmlNode): string | undefined {
    for (const child of p.children) {
      if (child.name !== "bookmarkStart") continue;
      const name = at(child, "name") ?? "";
      if (name.startsWith("LastWordCode_") && name.length > "LastWordCode_".length) {
        return name.slice("LastWordCode_".length);
      }
    }
    return undefined;
  }

  /** Concatenated visible text of a paragraph (tabs and soft breaks included). */
  private plainText(p: XmlNode): string {
    return this.parseRuns(p, undefined)
      .map((r) => r.text)
      .join("");
  }

  // ── Runs ─────────────────────────────────────────────────────────────────

  /**
   * The runs of a paragraph-like container. `state` is the paragraph's: it
   * carries "the footnote's own mark was just read" from one w:r to the next,
   * so the space after a footnote's number is stripped once.
   */
  private parseRuns(container: XmlNode, link: string | undefined, state: RunState = { afterMark: false }): Run[] {
    const runs: Run[] = [];
    for (const child of container.children) {
      switch (child.name) {
        case "r":
          // A list: one w:r holding text AND a footnote reference (another
          // producer's spelling) is two runs in the model.
          runs.push(...this.runFrom(child, link, state));
          break;
        case "hyperlink": {
          const rid = at(child, "id");
          const rel = rid ? this.rels[rid] : undefined;
          const anchor = at(child, "anchor");
          const url = rel?.external ? rel.target : anchor ? `#${anchor}` : undefined;
          runs.push(...this.parseRuns(child, url ?? link, state));
          break;
        }
        case "sdt": {
          const content = el(child, "sdtContent");
          if (content) runs.push(...this.parseRuns(content, link, state));
          break;
        }
        case "ins":
        case "smartTag":
          runs.push(...this.parseRuns(child, link, state));
          break;
        default:
          break;
      }
    }
    return mergeRuns(runs);
  }

  /**
   * One w:r as model runs: usually one, none when it carries no text, and a
   * footnote run where it holds a `w:footnoteReference` (with the text either
   * side of it as runs of their own). Mirrors PHP `parseRun()`.
   */
  private runFrom(r: XmlNode, link: string | undefined, state: RunState = { afterMark: false }): Run[] {
    const out: Run[] = [];
    let text = "";
    for (const child of r.children) {
      if (child.name === "t") {
        let piece = child.text;
        // The space Word puts between a footnote's own mark and its text is
        // the mark's, not the text's.
        if (state.afterMark && piece !== "") {
          if (piece[0] === " ") piece = piece.slice(1);
          state.afterMark = false;
        }
        text += piece;
      } else if (child.name === "footnoteRef") {
        // A footnote's own number, inside footnotes.xml: Word draws it.
        state.afterMark = true;
      } else if (child.name === "footnoteReference") {
        const note = this.footnoteRun(at(child, "id") ?? "");
        if (note !== null) {
          if (text !== "") {
            out.push(this.runWith(text, r, link));
            text = "";
          }
          out.push(note);
        }
      } else if (child.name === "br" && at(child, "type") !== "page") text += "\n";
      else if (child.name === "cr") text += "\n";
      else if (child.name === "tab") {
        if (state.afterMark) {
          state.afterMark = false; // a tab after the mark is the mark's too
          continue;
        }
        text += "\t";
      }
    }
    if (text !== "") out.push(this.runWith(text, r, link));
    return out;
  }

  /** A model run from text and its w:r's run properties. */
  private runWith(text: string, r: XmlNode, link: string | undefined): Run {
    const rPr = el(r, "rPr");
    const run: Run = { text };
    if (onFlag(el(rPr, "b"))) run.bold = true;
    if (onFlag(el(rPr, "i"))) run.italic = true;
    if (onFlag(el(rPr, "strike"))) run.strike = true;

    const u = el(rPr, "u");
    if (u && (at(u, "val") ?? "single") !== "none") run.underline = true;

    if (onFlag(el(rPr, "smallCaps"))) run.smallCaps = true;

    const rStyle = at(el(rPr, "rStyle"), "val") ?? "";
    const ascii = at(el(rPr, "rFonts"), "ascii");
    const asciiFont = (ascii ?? "").toLowerCase();
    if (/^InlineCode$/i.test(rStyle) || MONO_FONTS.includes(asciiFont)) run.code = true;
    // A `code` run's Consolas came FROM `code`, so surfacing it as `font` too
    // would hand back a bigger model than went in — and the next write would
    // then differ from the one just read.
    else if (ascii !== undefined && ascii !== "") run.font = ascii;

    const sz = numAttr(el(rPr, "sz"), "val");
    if (sz !== null) run.size = sz / 2;

    // Tracking of zero is the writer's "absent", so a zero here would be a
    // property nobody asked for.
    const tracking = numAttr(el(rPr, "spacing"), "val");
    if (tracking !== null && tracking !== 0) run.letterSpacing = tracking / 20;

    if (link) run.link = link;

    const color = at(el(rPr, "color"), "val");
    if (color && color !== "auto") run.color = `#${color.toUpperCase()}`;

    const shdFill = at(el(rPr, "shd"), "fill");
    if (shdFill && shdFill !== "auto") {
      run.highlight = `#${shdFill.toUpperCase()}`;
    } else {
      const named = at(el(rPr, "highlight"), "val");
      if (named && named !== "none" && HIGHLIGHT_NAMES[named]) run.highlight = HIGHLIGHT_NAMES[named];
    }

    return run;
  }

  // ── Lists ────────────────────────────────────────────────────────────────

  private buildLists(flat: FlatListItem[]): Block[] {
    const blocks: Block[] = [];
    let i = 0;
    while (i < flat.length) {
      const numId = flat[i]!.numId;
      const group: FlatListItem[] = [];
      while (i < flat.length && flat[i]!.numId === numId) {
        group.push(flat[i]!);
        i++;
      }
      const ordered = this.numOrdered[numId] ?? false;
      const items = buildTree(group);
      const block: Block = ordered
        ? ({ type: "list", ordered: true, items } as Block)
        : ({ type: "list", items } as Block);
      blocks.push(block);
    }
    return blocks;
  }

  // ── Tables ───────────────────────────────────────────────────────────────

  /**
   * A table back into the model.
   *
   * Two things make this the hardest read in the package. First, the writer
   * emits borders, cell margins and a `w:tcW` for every cell whether or not
   * the model asked — so anything it would have produced anyway is NOT
   * surfaced, or every document written before these options existed would
   * come back carrying options nobody set. Second, the file contains the
   * `w:vMerge` continuation cells the writer synthesised, and they have to be
   * dropped and turned back into a `rowSpan` on the cell above.
   */
  private parseTable(tbl: XmlNode, ctx: WalkCtx): Block {
    const tblPr = el(tbl, "tblPr");
    const grid = els(el(tbl, "tblGrid"), "gridCol").map((g) => numAttr(g, "w") ?? 0);
    const table: Record<string, unknown> = { type: "table" };

    const tblW = el(tblPr, "tblW");
    if (at(tblW, "type") === "pct") {
      const w = numAttr(tblW, "w");
      if (w !== null) table.width = Math.round((w / 50) * 100) / 100;
    }

    const jc = at(el(tblPr, "jc"), "val");
    if (jc === "center" || jc === "right") table.align = jc;

    const borders = bordersFrom(el(tblPr, "tblBorders"), TABLE_EDGES);
    if (borders && !isDefaultTableBorders(borders)) table.borders = borders;

    const padding = sidesFrom(el(tblPr, "tblCellMar"));
    if (padding && !isDefaultCellMargins(padding)) table.cellPadding = padding;

    // Weights are only surfaced when the grid is NOT what an equal split
    // would have produced — compared against the split the writer computes,
    // not tested for exact equality, so a three-column table whose width does
    // not divide by three is still recognised as equal.
    const total = grid.reduce((a, b) => a + b, 0);
    if (grid.length > 0 && total > 0 && !sameGrid(grid, splitColumns(total, grid.length))) {
      table.widths = grid.map((w) => Math.round((w / total) * 10000) / 100);
    }

    // Pass 1: read every emitted cell, keeping its grid column and merge state.
    interface Slot {
      col: number;
      span: number;
      vMerge: string | null;
      cell: Record<string, unknown> | null;
    }
    const laid: Slot[][] = els(tbl, "tr").map((tr) => {
      const header = el(el(tr, "trPr"), "tblHeader") !== undefined;
      const slots: Slot[] = [];
      let col = 0;
      for (const tc of els(tr, "tc")) {
        const tcPr = el(tc, "tcPr");
        const span = Math.max(1, Math.min(Schema.MAX_COL_SPAN, Math.trunc(numAttr(el(tcPr, "gridSpan"), "val") ?? 1)));
        const vMergeEl = el(tcPr, "vMerge");
        const vMerge = vMergeEl ? (at(vMergeEl, "val") ?? "continue") : null;

        let cell: Record<string, unknown> | null = null;
        if (vMerge !== "continue") {
          cell = { blocks: this.walkBody(tc.children, ctx) };

          const fill = at(el(tcPr, "shd"), "fill");
          // A header row's grey came FROM `header`, so it is attributable and
          // not surfaced. Any other fill is the author's.
          if (fill !== undefined && fill !== "auto" && !(header && fill.toUpperCase() === HEADER_FILL)) {
            cell.shading = `#${fill.toUpperCase()}`;
          }
          const cellBorders = bordersFrom(el(tcPr, "tcBorders"), BOX_EDGES);
          if (cellBorders) cell.borders = cellBorders;
          const cellPadding = sidesFrom(el(tcPr, "tcMar"));
          if (cellPadding) cell.padding = cellPadding;
          const valign = at(el(tcPr, "vAlign"), "val");
          if (valign === "top" || valign === "center" || valign === "bottom") cell.valign = valign;
          if (span > 1) cell.colSpan = span;
        }

        slots.push({ col, span, vMerge, cell });
        col += span;
      }
      return slots;
    });

    // Pass 2: fold each run of continuations back into the cell that started it.
    laid.forEach((line, r) => {
      for (const slot of line) {
        if (slot.vMerge !== "restart" || !slot.cell) continue;
        let covered = 1;
        for (let below = r + 1; below < laid.length; below++) {
          const match = laid[below]!.find((s) => s.col === slot.col && s.vMerge === "continue");
          if (!match) break;
          covered++;
        }
        if (covered > 1) slot.cell.rowSpan = covered;
      }
    });

    table.rows = laid.map((line, r) => {
      const header = el(el(els(tbl, "tr")[r]!, "trPr"), "tblHeader") !== undefined;
      const cells = line.filter((s) => s.cell !== null).map((s) => s.cell!);
      return header ? { header: true, cells } : { cells };
    });

    return table as unknown as Block;
  }

  // ── Images ───────────────────────────────────────────────────────────────

  private parseDrawing(drawing: XmlNode): Block[] {
    const frame = el(drawing, "inline") ?? el(drawing, "anchor");
    if (!frame) return [];

    const blip = findFirst(frame, "blip");
    if (!blip) return this.parseChartDrawing(frame);
    const rid = at(blip, "embed") ?? at(blip, "link");
    const rel = rid ? this.rels[rid] : undefined;
    if (!rel) return [];

    const target = rel.target.replace(/^\.\//, "");
    const partName = target.startsWith("/") ? target.slice(1) : `word/${target}`;
    const media = this.parts[partName];
    if (!media) return [];

    const ext = (partName.split(".").pop() ?? "").toLowerCase();
    const mime = ext === "png" ? "image/png" : ext === "jpg" || ext === "jpeg" ? "image/jpeg" : ext === "gif" ? "image/gif" : "image/png";

    const block: Any = { type: "image", src: `data:${mime};base64,${base64Encode(media)}` };

    const extent = el(frame, "extent");
    const cx = parseInt(at(extent, "cx") ?? "0", 10);
    const cy = parseInt(at(extent, "cy") ?? "0", 10);
    if (cx > 0) block.widthPx = Math.max(1, Math.round(cx / EMU_PER_PX));
    if (cy > 0) block.heightPx = Math.max(1, Math.round(cy / EMU_PER_PX));

    const descr = at(el(frame, "docPr"), "descr");
    if (descr) block.alt = descr;

    return [block as Block];
  }

  /**
   * A drawing holding a chart, back as a chart block; [] for any drawing that is
   * not a chart this package can model -- skipped, the read carries on.
   */
  private parseChartDrawing(frame: XmlNode): Block[] {
    const graphicData = findFirst(frame, "graphicData");
    if (!graphicData || at(graphicData, "uri") !== NS_C) return [];
    const rid = at(el(graphicData, "chart"), "id");
    const rel = rid ? this.rels[rid] : undefined;
    if (!rel) return [];
    const target = rel.target.replace(/^\.\//, "");
    const partName = target.startsWith("/") ? target.slice(1) : `word/${target}`;
    const xml = this.partText(partName);
    const read = xml !== null ? readChart(xml) : null;
    if (read === null) return [];

    const block: Any = { type: "chart", option: read.option };

    // The size, unless it is exactly what the writer gives a chart that names
    // none: the text column, DEFAULT_ASPECT as tall.
    const extent = el(frame, "extent");
    const cx = parseInt(at(extent, "cx") ?? "0", 10) || 0;
    const cy = parseInt(at(extent, "cy") ?? "0", 10) || 0;
    const column = this.contentWidth * 635;
    if (cx > 0 && cy > 0 && !(cx === column && cy === Math.round(column * DEFAULT_ASPECT))) {
      block.widthPx = Math.round(cx / EMU_PER_PX);
      block.heightPx = Math.round(cy / EMU_PER_PX);
    }
    if (read.numberFormat !== null) block.numberFormat = read.numberFormat;
    const descr = at(el(frame, "docPr"), "descr");
    if (descr) block.alt = descr;

    return [block as Block];
  }
}

/** Twips between the margins: page width less left and right. Mirrors PHP `contentWidthOf()`. */
/**
 * The body's sections in order, each by the `w:sectPr` describing it, the body's own
 * last; and how many sit inside a code or quote block, which are not read. Through the
 * same wrappers `walkBody()` reads through. PHP `DocxReader::sections()`.
 */
function sectionsOf(body: XmlNode): { sections: (XmlNode | undefined)[]; nested: number } {
  const sections: (XmlNode | undefined)[] = [];
  let nested = 0;
  const walk = (nodes: XmlNode[]): void => {
    for (const node of nodes) {
      switch (node.name) {
        case "p": {
          const sectPr = el(el(node, "pPr"), "sectPr");
          if (sectPr) sections.push(sectPr);
          break;
        }
        case "tbl":
        case "sectPr":
        case "bookmarkStart":
        case "bookmarkEnd":
        case "proofErr":
          break;
        case "sdt": {
          const tag = at(el(el(node, "sdtPr"), "tag"), "val") ?? "";
          const content = el(node, "sdtContent");
          if (!content) break;
          if (tag === SDT_TAG_CODE || tag.startsWith(SDT_TAG_CODE + ":") || tag === SDT_TAG_QUOTE) nested += sectionsIn(content);
          else walk(content.children);
          break;
        }
        default:
          walk(node.children);
      }
    }
  };
  walk(body.children);
  sections.push(el(body, "sectPr"));
  return { sections, nested };
}

/** Section breaks at any block level under a node. A walk: this nesting is attacker-supplied. */
function sectionsIn(root: XmlNode): number {
  let count = 0;
  const stack: XmlNode[] = [root];
  while (stack.length > 0) {
    for (const node of stack.pop()!.children) {
      if (node.name === "p") {
        if (el(el(node, "pPr"), "sectPr")) count++;
      } else if (node.name === "customXml") {
        stack.push(node);
      } else if (node.name === "sdt") {
        const content = el(node, "sdtContent");
        if (content) stack.push(content);
      }
    }
  }
  return count;
}

/**
 * Turns the walk's section markers into breaks, each report at the block it belongs to.
 * The paragraph the writer closes an empty last section with is not the author's either.
 */
function finishSections(walked: Block[], unrepresented: { where: string; kind: "layout-not-preserved"; count: number }[]): Block[] {
  const out: Block[] = [];
  walked.forEach((block, i) => {
    if ((block as Any).type !== SECTION_END) {
      if (isEmptyParagraph(block) && i === walked.length - 1 && (walked[i - 1] as Any)?.type === SECTION_END) return;
      out.push(block);
      return;
    }
    const plan = (block as Any).plan as SectionPlan | undefined;
    if (plan === undefined) return;
    if (plan.dropped > 0) unrepresented.push({ where: `/blocks/${out.length}`, kind: "layout-not-preserved", count: plan.dropped });
    if (plan.break !== null) out.push({ type: "pageBreak", page: plan.break } as Block);
  });
  return out;
}

/**
 * What a section changes from the page before it. PHP `pageChange()`: a side is compared
 * with defaults filled in; a band a section no longer shows is `{"blocks": []}`; the
 * first-page bands are never inherited, so never compared.
 */
function pageChange(from: Record<string, unknown>, to: Record<string, unknown>, knownPaper: boolean): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  const orientation = (p: Record<string, unknown>): unknown => p.orientation ?? "portrait";
  if (orientation(from) !== orientation(to)) out.orientation = orientation(to);
  const size = (p: Record<string, unknown>): unknown => p.size ?? "letter";
  if (knownPaper && size(from) !== size(to)) out.size = size(to);
  const was = (from.margins ?? {}) as Record<string, number>;
  const now = (to.margins ?? {}) as Record<string, number>;
  const margins: Record<string, number> = {};
  // In PHP's order, which the writer reads back the same: top, left, bottom, right.
  for (const side of ["top", "left", "bottom", "right", "header", "footer"]) {
    const fallback = side === "header" || side === "footer" ? 36 : 72;
    const value = now[side] ?? fallback;
    if ((was[side] ?? fallback) !== value) margins[side] = value;
  }
  if (Object.keys(margins).length > 0) out.margins = margins;
  for (const key of Schema.PAGE_BANDS) {
    const previous = key === "firstHeader" || key === "firstFooter" ? undefined : (from[key] as Any);
    const band = to[key];
    if (band !== undefined && JSON.stringify(band) !== JSON.stringify(previous)) out[key] = band;
    else if (band === undefined && previous !== undefined && !(Array.isArray(previous.blocks) && previous.blocks.length === 0)) out[key] = { blocks: [] };
  }
  return out;
}

/** What the first section asks for that the model cannot hold. */
function droppedByFirstSection(sectPr: XmlNode | undefined): number {
  return (columns(sectPr) > 1 ? 1 : 0) + (restarts(sectPr, false) ? 1 : 0) + (paperKnown(sectPr) ? 0 : 1);
}

/** `w:cols w:num`, 1 when absent. */
function columns(sectPr: XmlNode | undefined): number {
  return numAttr(el(sectPr, "cols"), "num") ?? 1;
}

/** `w:pgNumType w:start`: a first section starting at 1 changes nothing; any later start is a restart. */
function restarts(sectPr: XmlNode | undefined, later: boolean): boolean {
  const start = numAttr(el(sectPr, "pgNumType"), "start");
  return start !== null && (later || start !== 1);
}

/** Whether the section's paper is one the model names. No `w:pgSz` is the default, so it is. */
function paperKnown(sectPr: XmlNode | undefined): boolean {
  const pgSz = el(sectPr, "pgSz");
  if (!pgSz) return true;
  let w = numAttr(pgSz, "w") ?? 12240;
  let h = numAttr(pgSz, "h") ?? 15840;
  if (at(pgSz, "orient") === "landscape") [w, h] = [h, w];
  return Object.values(PAGE_SIZES).some((dims) => dims[0] === w && dims[1] === h);
}

function contentWidthOf(sectPr: XmlNode | undefined): number {
  if (!sectPr) return 9360;
  const num = (v: string | undefined, d: number): number => (v !== undefined && /^-?\d+(\.\d+)?$/.test(v.trim()) ? Math.trunc(Number(v)) : d);
  const w = num(at(el(sectPr, "pgSz"), "w"), 12240);
  const pgMar = el(sectPr, "pgMar");
  const left = num(at(pgMar, "left"), 1440);
  const right = num(at(pgMar, "right"), 1440);
  return Math.max(1, w - left - right);
}

// ── helpers ─────────────────────────────────────────────────────────────────

/**
 * Paragraph properties back into the model — the same set `paragraph`,
 * `heading` and a list item all accept.
 */
function paragraphPropsFrom(pPr: XmlNode | undefined): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  if (!pPr) return out;

  const jc = at(el(pPr, "jc"), "val");
  if (jc === "center" || jc === "right") out.align = jc;
  else if (jc === "both" || jc === "distribute" || jc === "justify") out.align = "justify";

  const spacing = el(pPr, "spacing");
  const before = numAttr(spacing, "before");
  if (before !== null) out.spaceBefore = points(before);
  const after = numAttr(spacing, "after");
  if (after !== null) out.spaceAfter = points(after);
  const line = numAttr(spacing, "line");
  if (line !== null && at(spacing, "lineRule") === "auto") {
    out.lineHeight = Math.round((line / 240) * 1000) / 1000;
  }

  const ind = el(pPr, "ind");
  const left = numAttr(ind, "left");
  if (left !== null) out.indentLeft = points(left);
  const right = numAttr(ind, "right");
  if (right !== null) out.indentRight = points(right);

  if (onFlag(el(pPr, "keepNext"))) out.keepNext = true;

  const fill = at(el(pPr, "shd"), "fill");
  if (fill !== undefined && fill !== "auto") out.shading = `#${fill.toUpperCase()}`;

  const borders = bordersFrom(el(pPr, "pBdr"), BOX_EDGES);
  if (borders) out.borders = borders;

  return out;
}

/** A paragraph carrying no runs and no properties — the adjacent-table pad. */
/** A paragraph with no runs, whatever its spacing says: nothing an author could lose. */
function holdsNothing(block: Block): boolean {
  const b = block as Any;
  return b?.type === "paragraph" && Array.isArray(b.runs) && b.runs.length === 0;
}

function isEmptyParagraph(block: Block | undefined): boolean {
  const b = block as Any;
  return (
    b?.type === "paragraph" &&
    Array.isArray(b.runs) &&
    b.runs.length === 0 &&
    Object.keys(b).length === 2
  );
}

/** A numeric attribute, or null when absent or unparseable. */
/**
 * A field instruction as a band token (`Schema.fieldTokenOf()`): PAGE {page},
 * NUMPAGES {pages}, SECTIONPAGES {sectionpages}, `DATE \@ "d MMMM yyyy"`
 * {date:d MMMM yyyy} ... and '' for any other field. Mirrors PHP.
 */
function fieldToken(instr: string): string {
  return Schema.fieldTokenOf(instr) ?? "";
}

/** A WordprocessingML on/off value: on unless it says 0, false or off. */
function toggleOn(val: string | undefined): boolean {
  return val === undefined || !["0", "false", "none", "off"].includes(val.toLowerCase());
}

function numAttr(node: XmlNode | undefined, name: string): number | null {
  const raw = at(node, name);
  if (raw === undefined) return null;
  const n = Number(raw);
  return Number.isFinite(n) ? n : null;
}

/** Twips → points, kept exact for the halves the writer can emit. */
function points(twips: number): number {
  return Math.round((twips / 20) * 100) / 100;
}

const BOX_EDGES = ["top", "right", "bottom", "left"] as const;
const TABLE_EDGES = [...BOX_EDGES, "insideH", "insideV"] as const;

/** The header-cell grey, shared with the writer and both sibling engines. */
const HEADER_FILL = "E7E7E7";

/**
 * Is this exactly what the writer emits for a table that asked for nothing?
 *
 * Not a tolerance and not a heuristic: the writer's defaults are a fixed set,
 * and anything differing from them by one edge or one twip is the author's and
 * must survive the read.
 */
function isDefaultTableBorders(borders: Record<string, unknown>): boolean {
  const edges = Object.keys(borders);
  if (edges.length !== TABLE_EDGES.length) return false;
  return TABLE_EDGES.every((edge) => {
    const b = borders[edge] as Record<string, unknown> | undefined;
    // The default edge is single / 0.5pt / auto, which reads back as width
    // alone — style and colour are both the omitted default.
    return b !== undefined && Object.keys(b).length === 1 && b.width === 0.5;
  });
}

function isDefaultCellMargins(sides: Record<string, number>): boolean {
  return (
    Object.keys(sides).length === 4 &&
    sides.top === 3 && sides.bottom === 3 && sides.left === 5.4 && sides.right === 5.4
  );
}

function sameGrid(a: number[], b: number[]): boolean {
  return a.length === b.length && a.every((v, i) => v === b[i]);
}

/**
 * One border edge back into the model.
 *
 * Defaults are NOT surfaced: `style` only when it is not `single`, `color`
 * only when it is not `auto`. Otherwise reading a document written from a
 * model returns a bigger model than went in, and writing that model back
 * produces a different file.
 */
function borderFrom(node: XmlNode | undefined): Record<string, unknown> | null {
  if (!node) return null;
  const val = at(node, "val") ?? "single";
  if (val === "nil" || val === "none") return { style: "none" };

  const out: Record<string, unknown> = {};
  if (val !== "single") out.style = val;
  const sz = numAttr(node, "sz");
  if (sz !== null) out.width = Math.round((sz / 8) * 1000) / 1000;
  const color = at(node, "color");
  if (color !== undefined && color !== "auto") out.color = `#${color.toUpperCase()}`;
  return out;
}

function bordersFrom(container: XmlNode | undefined, edges: readonly string[]): Record<string, unknown> | null {
  if (!container) return null;
  const out: Record<string, unknown> = {};
  for (const edge of edges) {
    const border = borderFrom(el(container, edge));
    if (border) out[edge] = border;
  }
  return Object.keys(out).length > 0 ? out : null;
}

function sidesFrom(container: XmlNode | undefined): Record<string, number> | null {
  if (!container) return null;
  const out: Record<string, number> = {};
  for (const side of BOX_EDGES) {
    const w = numAttr(el(container, side), "w");
    if (w !== null) out[side] = points(w);
  }
  return Object.keys(out).length > 0 ? out : null;
}

/** True when a toggle property element is present and not explicitly off. */
function onFlag(node: XmlNode | undefined): boolean {
  if (!node) return false;
  const val = node.attrs["val"];
  return val === undefined || !["0", "false", "none", "off"].includes(val);
}

/** Depth-first search for all descendants with the given local name. */
function findAll(node: XmlNode, name: string): XmlNode[] {
  const out: XmlNode[] = [];
  for (const child of node.children) {
    if (child.name === name) out.push(child);
    out.push(...findAll(child, name));
  }
  return out;
}

function findFirst(node: XmlNode | undefined, name: string): XmlNode | undefined {
  if (!node) return undefined;
  for (const child of node.children) {
    if (child.name === name) return child;
    const nested = findFirst(child, name);
    if (nested) return nested;
  }
  return undefined;
}

/** Merge adjacent runs whose properties are identical (writer normalization). */
export function mergeRuns(runs: Run[]): Run[] {
  const out: Run[] = [];
  for (const run of runs) {
    // A footnote run's text is always "", and it is never merged: two
    // identical footnotes in a row are two footnotes.
    const isNote = isNoteRun(run);
    if (run.text === "" && !isNote) continue;
    const prev = out[out.length - 1];
    if (prev && !isNote && !isNoteRun(prev) && sameProps(prev, run)) {
      prev.text += run.text;
    } else {
      out.push({ ...run });
    }
  }
  return out;
}

function isNoteRun(run: Run): boolean {
  return Object.prototype.hasOwnProperty.call(run, "footnote") && run.footnote !== undefined;
}

/** What one paragraph's runs carry from one w:r to the next. */
interface RunState {
  /** A footnote's own mark was just read: the space or tab after it is the mark's. */
  afterMark: boolean;
}

/** The target of the first non-external relationship whose type ends `/${type}`, or null. Mirrors PHP. */
function relTargetOfType(relsXml: string, type: string): string | null {
  const root = parseXml(relsXml);
  for (const rel of root?.children ?? []) {
    if (
      rel.name === "Relationship" &&
      (at(rel, "Type") ?? "").endsWith(`/${type}`) &&
      (at(rel, "TargetMode") ?? "").toLowerCase() !== "external" &&
      (at(rel, "Target") ?? "") !== ""
    ) {
      return at(rel, "Target")!;
    }
  }
  return null;
}

function sameProps(a: Run, b: Run): boolean {
  // Every property a run can carry. A key missing from this list makes two
  // differently-formatted runs merge into one, silently taking the first
  // one's formatting — so it grows whenever the run model does.
  const keys = [
    "bold", "italic", "underline", "strike", "code", "link", "color", "highlight",
    "smallCaps", "size", "font", "letterSpacing",
  ] as const;
  return keys.every((k) => (a[k] ?? undefined) === (b[k] ?? undefined));
}

/** Build a nested ListItem tree from flat (ilvl-tagged) paragraphs. */
function buildTree(flat: FlatListItem[]): ListItem[] {
  const root: ListItem[] = [];
  const lastAtDepth: (ListItem | undefined)[] = [];

  for (const entry of flat) {
    const item: ListItem = { runs: entry.runs };
    let depth = Math.max(0, entry.ilvl);

    // Clamp orphan depths to the nearest existing parent.
    while (depth > 0 && lastAtDepth[depth - 1] === undefined) depth--;

    if (depth === 0) {
      root.push(item);
    } else {
      const parent = lastAtDepth[depth - 1]!;
      if (!parent.children) parent.children = [];
      parent.children.push(item);
    }
    lastAtDepth[depth] = item;
    lastAtDepth.length = depth + 1;
  }
  return root;
}

/** Keep `language` before `text` for stable JSON output (cosmetic only). */
function orderCodeKeys(block: Block): Block {
  const b = block as Any;
  if (b.language === undefined) return block;
  return { type: "code", language: b.language, text: b.text } as Block;
}
