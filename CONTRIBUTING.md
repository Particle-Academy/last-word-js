# Contributing to @particle-academy/last-word

Notes for anyone, human or agent, working ON this package's code. How to USE the
package is in [AGENTS.md](./AGENTS.md); what it is, is in [README.md](./README.md).
`CLAUDE.md` is a symlink to this file, so Claude Code reads these notes when it
works in this repo.

## Build and test

```bash
npm install
npm run build     # tsup
npm run lint      # tsc --noEmit
npm test          # vitest run
```

Tests are vitest, under `tests/`, with fixtures in `tests/fixtures` and
`test/fixtures`.

- **The agent guide is the same bytes in all three engines.**
  `agent-guide/professional-documents.md` ships in the package and is embedded
  through `src/agent-guide.json`; `tests/agent-guide.test.ts` pins its SHA-256
  over LF-normalised content and checks the embedded copy equals the published
  one. The canonical copy is the PHP package's: edit all three copies, then all
  three hash constants. `tests/agent-guide-covers-the-api.test.ts` fails if the
  guide stops naming part of the API.
- **Parity with the PHP package.** `tests/parity.test.ts` runs
  `scripts/php-tobytes.php` through `php` (resolved via the shell) and diffs
  each part of the unzipped output, since PHP writes DEFLATE with real mtimes
  and this port writes STORE with a fixed date. It needs `php` on the path.

## Changelog

Update `CHANGELOG.md` (`[Unreleased]`) in the same commit as the change.
