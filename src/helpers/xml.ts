/**
 * XML escaping for the PPTX writer. Mirrors PHP `Helpers\Xml`:
 *  - text(): ENT_XML1 | ENT_COMPAT  — escapes & < > " (NOT apostrophe)
 *  - attr(): ENT_XML1 | ENT_QUOTES  — escapes & < > " '
 *  - declaration(): trailing newline included (matches PHP byte output)
 */
export const Xml = {
  text(s: string): string {
    return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  },

  attr(s: string): string {
    return s
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  },

  declaration(standalone = true): string {
    const sa = standalone ? ' standalone="yes"' : "";
    return `<?xml version="1.0" encoding="UTF-8"${sa}?>\n`;
  },
};
