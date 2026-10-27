export { Agent, VERSION, convertTemplate, type WriteOptions } from "./agent";
export { SchemaException, UnsupportedFormatException, TemplateException, ArchiveRefused, OfficeTemplateRefused } from "./exceptions";
export type { SourceKind } from "./zip/office-template";
export type { ArchiveLimits } from "./zip/source-archive";
export * from "./schema/types";

// Lower-level building blocks (advanced use / parity with PHP services).
export { Validator } from "./schema/validator";
export { Repairer } from "./schema/repairer";
export { Schema } from "./schema/schema";
export { DocxWriter, resolveImageSize, EMU_PER_PX, MAX_WIDTH_PX } from "./writer/docx-writer";
export { DocxReader, mergeRuns } from "./reader/docx-reader";
export { DocReader } from "./reader/doc-reader";
export { OdtReader } from "./reader/odt-reader";
export { RtfReader } from "./reader/rtf-reader";
export { Format, detectFormat } from "./reader/format";
export { toMarkdown } from "./markdown/to-markdown";
export { fromMarkdown, parseInline } from "./markdown/from-markdown";
export { pngSize, jpegSize, sniffImageSize, parseDataUrl } from "./helpers/image-size";
export {
  SUPPORTED_CHART_TYPES,
  DEFAULT_PALETTE,
  DEFAULT_ASPECT,
  chartSpec,
  chartNum,
  chartLabel,
  chartParseNum,
  type ChartSpec,
  type ChartSeriesSpec,
} from "./helpers/chart";
export { zipSync, unzipSync, type ZipFile } from "./zip";

// Document versions as ops (Agent.diff / Agent.reduce / Agent.opSchema).
export { DocDiff } from "./ops/doc-diff";
export { DocReducer } from "./ops/doc-reducer";
export { DocOpSchema } from "./ops/doc-op-schema";
export type * from "./ops/types";
