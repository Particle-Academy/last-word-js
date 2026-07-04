/**
 * DocxWriter — Doc JSON model → .docx bytes (OOXML / WordprocessingML).
 * Deterministic: no timestamps, fixed zip entry order, rel ids assigned in
 * traversal order. Mirrors PHP `Writer\DocxWriter`.
 */

import { Xml } from "../helpers/xml";
import { parseDataUrl, sniffImageSize } from "../helpers/image-size";
import type { Block, Doc, ListItem, Run } from "../schema/types";
import { zipSync, type ZipFile } from "../zip";

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

/** EMUs per pixel at 96 dpi. */
export const EMU_PER_PX = 9525;
/** Content width cap: 6.5in at 96 dpi. */
export const MAX_WIDTH_PX = 624;

const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const REL_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
const REL_NUMBERING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
const REL_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const REL_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const REL_CORE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

const NUM_ID_BULLET = 1;
const NUM_ID_DECIMAL = 2;
const MAX_ILVL = 5;

/** SDT tag prefixes used to round-trip block metadata that OOXML has no slot for. */
export const SDT_TAG_CODE = "lastword:code";
export const SDT_TAG_QUOTE = "lastword:quote";

interface MediaFile {
  name: string; // e.g. image1.png
  bytes: Uint8Array;
  ext: string;
}

interface Rel {
  id: string;
  type: string;
  target: string;
  external: boolean;
}

interface RenderCtx {
  /** pStyle applied to plain paragraphs (used inside quote SDTs). */
  paragraphStyle?: string;
}

function normalizeHex(hex: string): string {
  return hex.replace(/^#/, "").toUpperCase();
}

export class DocxWriter {
  private rels: Rel[] = [];
  private hyperlinkIds = new Map<string, string>();
  private media: MediaFile[] = [];
  private drawingId = 0;

  toBytes(doc: Doc): Uint8Array {
    this.rels = [
      { id: "rId1", type: REL_STYLES, target: "styles.xml", external: false },
      { id: "rId2", type: REL_NUMBERING, target: "numbering.xml", external: false },
    ];
    this.hyperlinkIds = new Map();
    this.media = [];
    this.drawingId = 0;

    const body = this.renderBlocks(doc.blocks ?? [], {});
    const documentXml =
      Xml.declaration() +
      `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}" xmlns:wp="${NS_WP}" xmlns:a="${NS_A}" xmlns:pic="${NS_PIC}">` +
      `<w:body>${body}${this.sectPr()}</w:body></w:document>`;

    const files: ZipFile[] = [
      { name: "[Content_Types].xml", data: encode(this.contentTypesXml(doc)) },
      { name: "_rels/.rels", data: encode(this.packageRelsXml(doc)) },
    ];
    if (doc.title !== undefined) {
      files.push({ name: "docProps/core.xml", data: encode(this.coreXml(String(doc.title))) });
    }
    files.push(
      { name: "word/document.xml", data: encode(documentXml) },
      { name: "word/styles.xml", data: encode(this.stylesXml()) },
      { name: "word/numbering.xml", data: encode(this.numberingXml()) },
      { name: "word/_rels/document.xml.rels", data: encode(this.documentRelsXml()) },
    );
    for (const m of this.media) {
      files.push({ name: `word/media/${m.name}`, data: m.bytes });
    }
    return zipSync(files);
  }

  // ── Blocks ────────────────────────────────────────────────────────────────

  private renderBlocks(blocks: Block[], ctx: RenderCtx): string {
    let out = "";
    for (const block of blocks) {
      out += this.renderBlock(block as Any, ctx);
    }
    return out;
  }

  private renderBlock(block: Any, ctx: RenderCtx): string {
    switch (block.type) {
      case "heading":
        return this.paragraph(this.pPr(`Heading${clampLevel(block.level)}`), this.renderRuns(block.runs ?? []));
      case "paragraph": {
        const jc = block.align && block.align !== "left" ? (block.align === "justify" ? "both" : block.align) : null;
        return this.paragraph(this.pPr(ctx.paragraphStyle ?? null, null, jc), this.renderRuns(block.runs ?? []));
      }
      case "list":
        return this.renderList(block.items ?? [], block.ordered === true, 0);
      case "table":
        return this.renderTable(block);
      case "code":
        return this.renderCode(block);
      case "quote":
        return this.renderQuote(block);
      case "image":
        return this.renderImage(block);
      case "pageBreak":
        return `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
      case "hr":
        return (
          `<w:p><w:pPr><w:pBdr>` +
          `<w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/>` +
          `</w:pBdr></w:pPr></w:p>`
        );
      default:
        return "";
    }
  }

  private paragraph(pPr: string, runs: string): string {
    return `<w:p>${pPr}${runs}</w:p>`;
  }

  private pPr(style: string | null, numPr: string | null = null, jc: string | null = null): string {
    let inner = "";
    if (style) inner += `<w:pStyle w:val="${Xml.attr(style)}"/>`;
    if (numPr) inner += numPr;
    if (jc) inner += `<w:jc w:val="${jc}"/>`;
    return inner === "" ? "" : `<w:pPr>${inner}</w:pPr>`;
  }

  // ── Runs ──────────────────────────────────────────────────────────────────

  private renderRuns(runs: Run[]): string {
    let out = "";
    for (const run of runs) {
      if (typeof run?.text !== "string") continue;
      if (run.link) {
        const rid = this.hyperlinkRel(run.link);
        out += `<w:hyperlink r:id="${rid}" w:history="1">${this.renderRun(run, true)}</w:hyperlink>`;
      } else {
        out += this.renderRun(run, false);
      }
    }
    return out;
  }

  private renderRun(run: Run, linked: boolean): string {
    let rPr = "";
    if (run.code) rPr += `<w:rStyle w:val="InlineCode"/>`;
    else if (linked) rPr += `<w:rStyle w:val="Hyperlink"/>`;
    if (run.code) rPr += `<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>`;
    if (run.bold) rPr += `<w:b/>`;
    if (run.italic) rPr += `<w:i/>`;
    if (run.strike) rPr += `<w:strike/>`;
    if (run.color) rPr += `<w:color w:val="${normalizeHex(run.color)}"/>`;
    if (run.underline) rPr += `<w:u w:val="single"/>`;
    if (run.highlight) {
      rPr += `<w:shd w:val="clear" w:color="auto" w:fill="${normalizeHex(run.highlight)}"/>`;
    }
    const pr = rPr === "" ? "" : `<w:rPr>${rPr}</w:rPr>`;

    // Newlines inside a run become soft line breaks.
    const parts = run.text.split("\n");
    let content = "";
    parts.forEach((part, i) => {
      if (i > 0) content += `<w:br/>`;
      if (part !== "") content += `<w:t xml:space="preserve">${Xml.text(part)}</w:t>`;
    });
    if (content === "") content = `<w:t xml:space="preserve"></w:t>`;
    return `<w:r>${pr}${content}</w:r>`;
  }

  private hyperlinkRel(url: string): string {
    const existing = this.hyperlinkIds.get(url);
    if (existing) return existing;
    const id = `rId${this.rels.length + 1}`;
    this.rels.push({ id, type: REL_HYPERLINK, target: url, external: true });
    this.hyperlinkIds.set(url, id);
    return id;
  }

  // ── Lists ─────────────────────────────────────────────────────────────────

  private renderList(items: ListItem[], ordered: boolean, depth: number): string {
    const numId = ordered ? NUM_ID_DECIMAL : NUM_ID_BULLET;
    const ilvl = Math.min(depth, MAX_ILVL);
    let out = "";
    for (const item of items) {
      const numPr = `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`;
      out += this.paragraph(this.pPr(null, numPr), this.renderRuns(item.runs ?? []));
      if (item.children && item.children.length > 0) {
        out += this.renderList(item.children, ordered, depth + 1);
      }
    }
    return out;
  }

  // ── Tables ────────────────────────────────────────────────────────────────

  private renderTable(block: Any): string {
    const rows: Any[] = block.rows ?? [];
    const cols = Math.max(1, ...rows.map((r: Any) => (Array.isArray(r?.cells) ? r.cells.length : 0)));
    const colWidth = Math.floor(9360 / cols);

    let out = `<w:tbl><w:tblPr><w:tblStyle w:val="LastWordTable"/><w:tblW w:w="0" w:type="auto"/>`;
    out += `<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>`;
    out += `</w:tblPr><w:tblGrid>`;
    for (let c = 0; c < cols; c++) out += `<w:gridCol w:w="${colWidth}"/>`;
    out += `</w:tblGrid>`;

    for (const row of rows) {
      const header = row?.header === true;
      out += `<w:tr>`;
      if (header) out += `<w:trPr><w:tblHeader/></w:trPr>`;
      const cells: Any[] = Array.isArray(row?.cells) ? row.cells : [];
      for (const cell of cells) {
        out += `<w:tc><w:tcPr><w:tcW w:w="${colWidth}" w:type="dxa"/>`;
        if (header) out += `<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>`;
        out += `</w:tcPr>`;
        let inner = this.renderBlocks(cell?.blocks ?? [], {});
        // A table cell must end with a paragraph.
        if (inner === "" || !inner.endsWith("</w:p>")) inner += `<w:p/>`;
        out += inner + `</w:tc>`;
      }
      out += `</w:tr>`;
    }
    out += `</w:tbl>`;
    return out;
  }

  // ── Code / quote (SDT-wrapped for lossless round-trip) ──────────────────

  private renderCode(block: Any): string {
    const language = typeof block.language === "string" && block.language !== "" ? block.language : null;
    const tag = language ? `${SDT_TAG_CODE}:${language}` : SDT_TAG_CODE;
    const lines = String(block.text ?? "").split("\n");
    let body = "";
    for (const line of lines) {
      const runs =
        line === "" ? "" : `<w:r><w:t xml:space="preserve">${Xml.text(line)}</w:t></w:r>`;
      body += this.paragraph(this.pPr("CodeBlock"), runs);
    }
    return (
      `<w:sdt><w:sdtPr><w:alias w:val="Code"/><w:tag w:val="${Xml.attr(tag)}"/></w:sdtPr>` +
      `<w:sdtContent>${body}</w:sdtContent></w:sdt>`
    );
  }

  private renderQuote(block: Any): string {
    const body = this.renderBlocks(block.blocks ?? [], { paragraphStyle: "Quote" });
    return (
      `<w:sdt><w:sdtPr><w:alias w:val="Quote"/><w:tag w:val="${SDT_TAG_QUOTE}"/></w:sdtPr>` +
      `<w:sdtContent>${body === "" ? "<w:p/>" : body}</w:sdtContent></w:sdt>`
    );
  }

  // ── Images ────────────────────────────────────────────────────────────────

  private renderImage(block: Any): string {
    const decoded = parseDataUrl(String(block.src ?? ""));
    if (!decoded) {
      throw new Error("last-word: image `src` must be a base64 data URL (data:image/png;base64,… or data:image/jpeg;base64,…)");
    }

    const ext = extForMime(decoded.mime);
    const name = `image${this.media.length + 1}.${ext}`;
    this.media.push({ name, bytes: decoded.bytes, ext });
    const rid = `rId${this.rels.length + 1}`;
    this.rels.push({ id: rid, type: REL_IMAGE, target: `media/${name}`, external: false });

    const { width, height } = resolveImageSize(block, decoded.bytes);
    const cx = Math.max(1, Math.round(width * EMU_PER_PX));
    const cy = Math.max(1, Math.round(height * EMU_PER_PX));

    const id = ++this.drawingId;
    const alt = typeof block.alt === "string" ? block.alt : "";
    const descr = alt === "" ? "" : ` descr="${Xml.attr(alt)}"`;

    return (
      `<w:p><w:r><w:drawing>` +
      `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
      `<wp:extent cx="${cx}" cy="${cy}"/>` +
      `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
      `<wp:docPr id="${id}" name="Picture ${id}"${descr}/>` +
      `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
      `<a:graphic><a:graphicData uri="${NS_PIC}">` +
      `<pic:pic>` +
      `<pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"${descr}/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
      `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
    );
  }

  // ── Parts ─────────────────────────────────────────────────────────────────

  private sectPr(): string {
    return (
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>` +
      `</w:sectPr>`
    );
  }

  private contentTypesXml(doc: Doc): string {
    const exts = [...new Set(this.media.map((m) => m.ext))].sort();
    let defaults =
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>`;
    for (const ext of exts) {
      defaults += `<Default Extension="${ext}" ContentType="${mimeForExt(ext)}"/>`;
    }
    let overrides =
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
      `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>`;
    if (doc.title !== undefined) {
      overrides += `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`;
    }
    return (
      Xml.declaration() +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${defaults}${overrides}</Types>`
    );
  }

  private packageRelsXml(doc: Doc): string {
    let rels = `<Relationship Id="rId1" Type="${REL_DOCUMENT}" Target="word/document.xml"/>`;
    if (doc.title !== undefined) {
      rels += `<Relationship Id="rId2" Type="${REL_CORE}" Target="docProps/core.xml"/>`;
    }
    return (
      Xml.declaration() +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`
    );
  }

  private coreXml(title: string): string {
    return (
      Xml.declaration() +
      `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
      `xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ` +
      `xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
      `<dc:title>${Xml.text(title)}</dc:title>` +
      `</cp:coreProperties>`
    );
  }

  private documentRelsXml(): string {
    let out = Xml.declaration() + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
    for (const rel of this.rels) {
      const mode = rel.external ? ` TargetMode="External"` : "";
      out += `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${Xml.attr(rel.target)}"${mode}/>`;
    }
    return out + `</Relationships>`;
  }

  private stylesXml(): string {
    const headingSizes = [40, 32, 28, 26, 24, 22];
    let styles =
      `<w:docDefaults><w:rPrDefault><w:rPr>` +
      `<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/>` +
      `<w:sz w:val="22"/><w:szCs w:val="22"/>` +
      `</w:rPr></w:rPrDefault>` +
      `<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>` +
      `</w:docDefaults>` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>`;

    for (let level = 1; level <= 6; level++) {
      const sz = headingSizes[level - 1]!;
      styles +=
        `<w:style w:type="paragraph" w:styleId="Heading${level}"><w:name w:val="heading ${level}"/>` +
        `<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>` +
        `<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="${level - 1}"/></w:pPr>` +
        `<w:rPr><w:b/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr></w:style>`;
    }

    styles +=
      `<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/>` +
      `<w:basedOn w:val="Normal"/><w:qFormat/>` +
      `<w:pPr><w:ind w:left="720"/></w:pPr>` +
      `<w:rPr><w:i/><w:color w:val="595959"/></w:rPr></w:style>` +
      `<w:style w:type="paragraph" w:styleId="CodeBlock"><w:name w:val="Code Block"/>` +
      `<w:basedOn w:val="Normal"/>` +
      `<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>` +
      `<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/></w:pPr>` +
      `<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>` +
      `<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:style>` +
      `<w:style w:type="character" w:styleId="InlineCode"><w:name w:val="Inline Code"/>` +
      `<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>` +
      `<w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/></w:rPr></w:style>` +
      `<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>` +
      `<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>` +
      `<w:style w:type="table" w:styleId="LastWordTable"><w:name w:val="LastWord Table"/>` +
      `<w:tblPr><w:tblBorders>` +
      `<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
      `<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
      `<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
      `<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
      `<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
      `<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>` +
      `</w:tblBorders>` +
      `<w:tblCellMar>` +
      `<w:top w:w="60" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>` +
      `<w:bottom w:w="60" w:type="dxa"/><w:right w:w="108" w:type="dxa"/>` +
      `</w:tblCellMar></w:tblPr>` +
      `<w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr>` +
      `<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/></w:tcPr></w:tblStylePr>` +
      `</w:style>`;

    return Xml.declaration() + `<w:styles xmlns:w="${NS_W}">${styles}</w:styles>`;
  }

  private numberingXml(): string {
    let out = Xml.declaration() + `<w:numbering xmlns:w="${NS_W}">`;
    // abstractNum 0: bullet, 6 indent levels
    out += `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>`;
    for (let i = 0; i <= MAX_ILVL; i++) {
      out +=
        `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/>` +
        `<w:lvlText w:val="•"/><w:lvlJc w:val="left"/>` +
        `<w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr></w:lvl>`;
    }
    out += `</w:abstractNum>`;
    // abstractNum 1: decimal, 6 indent levels
    out += `<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>`;
    for (let i = 0; i <= MAX_ILVL; i++) {
      out +=
        `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="decimal"/>` +
        `<w:lvlText w:val="%${i + 1}."/><w:lvlJc w:val="left"/>` +
        `<w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr></w:lvl>`;
    }
    out += `</w:abstractNum>`;
    out += `<w:num w:numId="${NUM_ID_BULLET}"><w:abstractNumId w:val="0"/></w:num>`;
    out += `<w:num w:numId="${NUM_ID_DECIMAL}"><w:abstractNumId w:val="1"/></w:num>`;
    return out + `</w:numbering>`;
  }
}

// ── helpers ─────────────────────────────────────────────────────────────────

const encoder = new TextEncoder();

function encode(xml: string): Uint8Array {
  return encoder.encode(xml);
}

function clampLevel(level: Any): number {
  const n = Number.isFinite(Number(level)) ? Math.trunc(Number(level)) : 1;
  return Math.min(6, Math.max(1, n));
}

function extForMime(mime: string): string {
  switch (mime) {
    case "image/png":
      return "png";
    case "image/jpeg":
    case "image/jpg":
      return "jpeg";
    case "image/gif":
      return "gif";
    default:
      return mime.split("/")[1] ?? "bin";
  }
}

function mimeForExt(ext: string): string {
  switch (ext) {
    case "png":
      return "image/png";
    case "jpeg":
      return "image/jpeg";
    case "gif":
      return "image/gif";
    default:
      return "application/octet-stream";
  }
}

/**
 * Resolve the rendered pixel size of an image block: explicit px win, missing
 * dimensions are derived from the intrinsic (sniffed) size keeping aspect,
 * and everything is capped at 6.5in width.
 */
export function resolveImageSize(
  block: { widthPx?: number; heightPx?: number },
  bytes: Uint8Array,
): { width: number; height: number } {
  const intrinsic = sniffImageSize(bytes);
  let width = typeof block.widthPx === "number" && block.widthPx > 0 ? block.widthPx : undefined;
  let height = typeof block.heightPx === "number" && block.heightPx > 0 ? block.heightPx : undefined;

  if (width === undefined && height === undefined) {
    width = intrinsic?.width ?? 300;
    height = intrinsic?.height ?? 200;
  } else if (width === undefined) {
    const aspect = intrinsic ? intrinsic.width / intrinsic.height : 1.5;
    width = height! * aspect;
  } else if (height === undefined) {
    const aspect = intrinsic ? intrinsic.width / intrinsic.height : 1.5;
    height = width / aspect;
  }

  if (width! > MAX_WIDTH_PX) {
    height = (height! * MAX_WIDTH_PX) / width!;
    width = MAX_WIDTH_PX;
  }

  return { width: Math.max(1, Math.round(width!)), height: Math.max(1, Math.round(height!)) };
}
