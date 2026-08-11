# Changelog

All notable changes to `@particle-academy/last-word` are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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
