# Changelog

## 0.2.0

Cross-language metadata parity with the PHP mirror
(particle-academy/last-word) — last-word-js#1.

- The two metadata slots that didn't cross languages now do, both
  directions: **title** (docProps/core.xml `dc:title`) and **code block
  `language`** (`lastword:code:{lang}` w:sdt tag — canonical on both
  sides now).
- Reader: back-compat fallback for the PHP ≤0.1.x legacy slot — an
  invisible `LastWordCode_{lang}` bookmark on the first code paragraph.
- New frozen cross-read vector: `test/fixtures/php-canonical.docx`
  (written by the PHP engine) + its JSON, asserted semantically
  deep-equal on read.

## 0.1.0

Initial release — the docx sibling of holy-sheet (xlsx) and dark-slide (pptx).

- JSON document model: heading / paragraph / list (nested) / table / code /
  quote / image / pageBreak / hr blocks with styled runs (bold, italic,
  underline, strike, inline code, link, color, highlight).
- `Agent` façade: `validate`, `validateAndRepair`, `toBytes`, `write`,
  `read`/`fromBytes`, `toMarkdown`, `fromMarkdown`, `describe`, `jsonSchema`,
  `version` — mirrors the PHP `particle-academy/last-word` surface.
- DOCX writer: deterministic OOXML output (styles, numbering, hyperlink +
  image rels, media parts, EMU extents with PNG IHDR / JPEG SOF sniffing).
- DOCX reader: round-trips its own output and tolerates Word-authored files
  (outlineLvl headings, named highlights, unknown numIds, unknown elements
  degrade to paragraphs).
- Markdown bridges: hand-rolled GFM emitter + parser (no external markdown
  dependency) for the react-fancy Editor round-trip.
