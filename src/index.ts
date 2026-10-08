export { Agent, VERSION, type WriteOptions } from "./agent";
export { SchemaException, UnsupportedFormatException, TemplateException } from "./exceptions";
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
export { zipSync, unzipSync, type ZipFile } from "./zip";

// Document versions as ops (Agent.diff / Agent.reduce / Agent.opSchema).
export { DocDiff } from "./ops/doc-diff";
export { DocReducer } from "./ops/doc-reducer";
export { DocOpSchema } from "./ops/doc-op-schema";
export type * from "./ops/types";
