import { zipSync, type ZipFile } from "../../src/zip";

/**
 * A minimal, valid `.dotx` built at test time.
 *
 * Generated rather than committed, matching the PHP engine's
 * `tests/Support/GeneratedTemplate` byte for byte in its *content* — a real house
 * template is someone's licensed property, and a binary fixture nobody can read
 * is a fixture nobody can reason about.
 *
 * Every value is deliberately unlike the writer's own defaults, so a test can
 * tell "the template was applied" from "the built-in look happens to match".
 *
 * Deliberately absent: `CodeBlock`, `InlineCode` and `Heading3`..`Heading6`, so a
 * test can show the writer supplies what the template lacks rather than emitting
 * a document that references an undefined style.
 */
export const TEMPLATE_FONT = "Garamond";
export const TEMPLATE_HEADING_COLOR = "1F3864";
export const TEMPLATE_ACCENT1 = "C00000";

const ENCODER = new TextEncoder();
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";

export function generatedTemplate(withTheme = true): Uint8Array {
  const files: ZipFile[] = [
    { name: "[Content_Types].xml", data: ENCODER.encode(contentTypes(withTheme)) },
    { name: "_rels/.rels", data: ENCODER.encode(topRels()) },
    { name: "word/document.xml", data: ENCODER.encode(documentXml()) },
    { name: "word/styles.xml", data: ENCODER.encode(stylesXml()) },
  ];
  if (withTheme) {
    files.push({ name: "word/theme/theme1.xml", data: ENCODER.encode(themeXml()) });
  }
  return zipSync(files);
}

/** A package that opens as a zip but defines no styles. */
export function templateWithoutStyles(): Uint8Array {
  return zipSync([
    { name: "[Content_Types].xml", data: ENCODER.encode(contentTypes(false)) },
    { name: "word/document.xml", data: ENCODER.encode(documentXml()) },
  ]);
}

function stylesXml(): string {
  const font = TEMPLATE_FONT;
  const color = TEMPLATE_HEADING_COLOR;

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:styles xmlns:w="${NS_W}">` +
    `<w:docDefaults><w:rPrDefault><w:rPr>` +
    `<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}" w:cs="${font}"/>` +
    `<w:sz w:val="24"/><w:szCs w:val="24"/>` +
    `</w:rPr></w:rPrDefault></w:docDefaults>` +
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>` +
    `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}"/><w:sz w:val="24"/></w:rPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/>` +
    `<w:qFormat/><w:rPr><w:b/><w:sz w:val="72"/></w:rPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>` +
    `<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>` +
    `<w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr>` +
    `<w:rPr><w:b/><w:color w:val="${color}"/><w:sz w:val="44"/></w:rPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/>` +
    `<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>` +
    `<w:pPr><w:keepNext/><w:outlineLvl w:val="1"/></w:pPr>` +
    `<w:rPr><w:b/><w:color w:val="${color}"/><w:sz w:val="32"/></w:rPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/>` +
    `<w:qFormat/><w:pPr><w:ind w:left="1440"/></w:pPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/>` +
    `<w:basedOn w:val="Normal"/><w:qFormat/></w:style>` +
    `<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/>` +
    `<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>` +
    // A style the writer never emits. Proves the template's own definitions are
    // carried through rather than filtered to the ones we recognise.
    `<w:style w:type="paragraph" w:styleId="HouseNote"><w:name w:val="House Note"/>` +
    `<w:basedOn w:val="Normal"/><w:rPr><w:i/><w:color w:val="${TEMPLATE_ACCENT1}"/></w:rPr></w:style>` +
    `</w:styles>`
  );
}

function themeXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<a:theme xmlns:a="${NS_A}" name="House"><a:themeElements>` +
    `<a:clrScheme name="House">` +
    `<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>` +
    `<a:dk2><a:srgbClr val="${TEMPLATE_HEADING_COLOR}"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>` +
    `<a:accent1><a:srgbClr val="${TEMPLATE_ACCENT1}"/></a:accent1>` +
    `<a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>` +
    `<a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>` +
    `<a:accent6><a:srgbClr val="70AD47"/></a:accent6>` +
    `<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>` +
    `</a:clrScheme>` +
    `<a:fontScheme name="House">` +
    `<a:majorFont><a:latin typeface="${TEMPLATE_FONT}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>` +
    `<a:minorFont><a:latin typeface="${TEMPLATE_FONT}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>` +
    `</a:fontScheme>` +
    `<a:fmtScheme name="House">` +
    `<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>` +
    `<a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>` +
    `<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>` +
    `<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>` +
    `</a:fmtScheme>` +
    `</a:themeElements></a:theme>`
  );
}

function contentTypes(withTheme: boolean): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    (withTheme
      ? '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
      : "") +
    "</Types>"
  );
}

function topRels(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    "</Relationships>"
  );
}

function documentXml(): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${NS_W}"><w:body><w:p/></w:body></w:document>`
  );
}
