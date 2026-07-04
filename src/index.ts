export { Agent, VERSION } from "./agent";
export { SchemaException } from "./exceptions";
export * from "./schema/types";

// Lower-level building blocks (advanced use / parity with PHP services).
export { Validator } from "./schema/validator";
export { Repairer } from "./schema/repairer";
export { Schema } from "./schema/schema";
export { DocxWriter, resolveImageSize, EMU_PER_PX, MAX_WIDTH_PX } from "./writer/docx-writer";
export { DocxReader, mergeRuns } from "./reader/docx-reader";
export { toMarkdown } from "./markdown/to-markdown";
export { fromMarkdown, parseInline } from "./markdown/from-markdown";
export { pngSize, jpegSize, sniffImageSize, parseDataUrl } from "./helpers/image-size";
export { zipSync, unzipSync, type ZipFile } from "./zip";
