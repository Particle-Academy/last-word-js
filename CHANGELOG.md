# Changelog

All notable changes to `@particle-academy/last-word` are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- **A rich-layout surface, so a business one-pager is expressible.** The model
  was far narrower than the XML this engine already emitted: font size, font
  family, small caps, letter spacing, per-cell shading, borders, padding,
  vertical alignment and both merge directions were produced from hardcoded
  blocks or from `styles.xml` and were **unreachable from the model**. An agent
  could emit `size`, `colSpan` or `shading`, the validator returned no errors,
  and every one of them was silently dropped.

  | where | new keys |
  |---|---|
  | run | `size` (points, half-points exact) · `font` · `smallCaps` · `letterSpacing` (points, may be negative) |
  | paragraph, **heading** and list item | `spaceBefore` · `spaceAfter` · `lineHeight` · `indentLeft` · `indentRight` · `keepNext` · `shading` · `borders` · `align` (on headings too) |
  | table | `widths` (relative column weights) · `width` (% of the text column) · `align` · `borders` (incl. `insideH` / `insideV`) · `cellPadding` |
  | cell | `shading` · `borders` · `padding` · `valign` · `colSpan` · `rowSpan` |
  | document | `page` (`size`, `orientation`, `margins`) · `defaultFont` · `defaultSize` |

  Every key is validated, every key round-trips through the reader, and every
  key appears in `jsonSchema()` so an agent registering the tool is told it
  exists.

- **A heading is now a paragraph.** It takes the same properties, so a section
  label that needs spacing or alignment no longer has to be a bold paragraph
  impersonating a heading — and therefore appearing in no navigation pane and
  no table of contents.

- **Both merge directions**, written HTML-style: a `rowSpan` cell appears ONCE
  and the rows it covers list only their own remaining cells. The writer
  synthesises the `w:vMerge` continuations OOXML requires and the reader folds
  them back, so the model that comes out is the model that went in.

- **The table grid is computed from the section.** All three engines carried
  `9360` twips as a literal, so a document that narrowed its margins got a
  table that no longer matched its own page — too narrow, and silently so.

- **`last-word/docx-constructs` in `fancy-conformance`** — 44 shared rows
  pinning which construct emits which XML, in which order, and what the reader
  gives back. The rows are not transcribed into this repo: all three engines
  assert the same file, so a mapping that drifts in one fails there rather than
  quietly becoming that engine's behaviour.

### Fixed

- **Adjacent tables no longer merge into one in Word.** OOXML merges two
  `<w:tbl>` elements that touch, imposing the first table's column grid on the
  second. A stat band followed by a callout became a single two-row table.
- **A run's properties can no longer be lost to run-merging.** Adjacent runs
  are merged when their formatting matches, and the comparison listed only the
  properties that existed when it was written — so two differently-sized runs
  merged into one and took the first one's size.

- **The `LastWordTable` style is no longer emitted.** Nothing references it
  now that table properties are inline. If you were overriding it in a template,
  override the table's `borders` and `cellPadding` instead.

### Changed

- **Table properties are now written inline, not taken from a named style.**
  A named table style cannot vary per table instance, so per-table borders
  forced this. Also reconciled with it: header cells are one grey in all three
  engines, and header bold is one mechanism.

- **Document defaults name an East Asian font.** Without it Word picks its own
  face for CJK runs, which is exactly the text a mixed-script document
  contains.

  **What you must do: nothing.** No existing key changed meaning, nothing was
  removed and nothing was renamed. A document written before this release
  produces the same page. The visible differences are confined to tables, are
  small, and are listed above so a pixel comparison against an old build is not
  a surprise.

### Notes

- **A `header: true` row is not a round-trip fixpoint, and now says so.** The
  writer bolds the row's runs and the reader honestly reports the bold it
  finds, so the model that comes out is not the model that went in. The
  alternatives were to stop bolding header rows (changing every existing
  consumer's output) or to have the reader strip bold from header rows
  (discarding bold an author really asked for). Pinned as case `0042` rather
  than left as a surprise.

- **Release order: `fancy-conformance` first.** The shared table is version
  `0.7.0`, which is not on a registry yet, so this package's dev dependency
  cannot resolve and its lockfile cannot be regenerated until that release
  lands. Nothing about the runtime surface depends on it — only the test that
  asserts the shared rows.

- **What DOCX cannot do, so nobody chases it:** table corners are always
  square (there is no border radius in WordprocessingML), a background cannot
  bleed past the page margin without an anchored drawing, and naming a font is
  not shipping one — a reader without it substitutes. None of the three is
  worked around here; a layout that needs them needs a different format.

### Added

- **Cross-engine WRITER parity against the PHP `last-word`** — the guarantee
  this pair was missing. `cross-read.test.ts` already proved the Node *reader*
  restores a frozen docx PHP wrote; nothing compared the two *writers*. Both
  siblings in this family (`holy-sheet`, `dark-slide`) diff PHP output against
  Node's; this one did not, so the engines could drift apart on anything the
  frozen fixture happened not to contain.

  Five documents — minimal, every inline mark, structure (lists / table /
  quote), non-ASCII (CJK, emoji, combining marks), and empty-vs-zero values —
  written by both engines and diffed part by part. Containers are never
  byte-compared: PHP writes DEFLATE with real mtimes, this port writes STORE
  with a fixed 1980 date, so the files can never match and it would be wrong to
  try.

  **PHP is REQUIRED in CI**, and the suite throws rather than skipping when it
  is absent. A skip is a green build with zero parity coverage, which is how the
  other two suites in this family reported success over nothing for months.

### Notes

- **The pair is NOT at parity, and the first run of the new suite proved it.**
  Four parts differ, and they drift in BOTH directions — so this is not "the
  port is behind":

  | part | difference |
  |---|---|
  | `word/document.xml` | hyperlinks: port emits `w:history="1"`, PHP does not · lists: PHP emits `<w:pStyle w:val="ListParagraph"/>`, port does not · tables: port references `<w:tblStyle w:val="LastWordTable"/>`, PHP inlines `<w:tblBorders>` |
  | `word/styles.xml` | port emits `w:eastAsia="Calibri"` |
  | `word/numbering.xml` | port emits `<w:multiLevelType w:val="hybridMultilevel"/>` |
  | `[Content_Types].xml` | PHP declares png/jpeg defaults unconditionally |

  **The table difference is user-visible**: borders from a named style versus
  borders written into the table properties. A table written by one backend is
  not the same object as one written by the other.

  Recorded in `KNOWN_DIVERGENT_PARTS` rather than fixed, because reconciling
  them changes rendered output for every existing consumer, and in several cases
  the port is arguably the more correct side — so "make the port match PHP", the
  obvious move given PHP is the reference, would remove the better behaviour.
  That is the pair owner's decision, not a side effect of writing the suite that
  found it.

  The list ratchets both ways: a new divergent part fails, and an entry that
  stops being true fails too. `word/document.xml` for the *minimal* document is
  asserted identical unconditionally — whatever else drifts, the engines must
  agree on a single plain paragraph.

## 0.3.0 — 2026-08-07

### Changed

- **BREAKING — Node 18 is no longer supported.** `engines.node` moves from `>=18` to `>=22`.

  **What you must do:** on Node 22 or newer, nothing. Note npm only *warns* on an `engines` mismatch while **pnpm fails the install**, so this surfaces differently depending on your package manager. Node 18 is end-of-life and 20 is maintenance-only.

### Why

These are the kit 0.5 platform floors, applied across every package at once so a consumer never has to resolve a mix. **No API changed, nothing was removed, nothing was renamed** — only what the package requires.

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
