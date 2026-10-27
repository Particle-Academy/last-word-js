/**
 * `npx last-word` (the `bin`). One command today, mirroring PHP `bin/last-word`:
 *
 *     last-word template:convert <input> <output> [--to=document|template] [--force]
 *
 * Exit codes: 0 converted, 1 refused (the reason is printed), 2 usage.
 *
 * Node-only, and never imported by the isomorphic index.
 *
 * @internal
 */
import { closeSync, existsSync, fstatSync, openSync, readSync, realpathSync, renameSync, statSync, unlinkSync, writeFileSync } from "node:fs";
import { randomBytes } from "node:crypto";
import { Agent } from "./agent";
import { ArchiveRefused, OfficeTemplateRefused, UnsupportedFormatException } from "./exceptions";
import { DEFAULT_ARCHIVE_LIMITS, SourceArchive } from "./zip/source-archive";
import { sourceKind, type SourceKind } from "./zip/office-template";

export interface CliIo {
  stdout: (text: string) => void;
  stderr: (text: string) => void;
}

const USAGE = "Usage: last-word template:convert <input> <output> [--to=document|template] [--force]\n"
  + "  Converts a .dotx/.ott template to its .docx/.odt twin, or back.\n"
  + "  Without --to it flips whichever the input is.\n";

/** A regular file, bounded before allocation (the sync twin of `readFileBounded`). */
function readFileBoundedSync(path: string): Uint8Array {
  if (!existsSync(path) || !statSync(path).isFile()) throw new ArchiveRefused("archive_path");
  const fd = openSync(path, "r");
  const limit = DEFAULT_ARCHIVE_LIMITS.maxArchiveBytes;
  try {
    if (!fstatSync(fd).isFile()) throw new ArchiveRefused("archive_path");
    if (fstatSync(fd).size > limit) throw new ArchiveRefused("archive_limit");
    const chunks: Uint8Array[] = [];
    let total = 0;
    for (;;) {
      const chunk = new Uint8Array(Math.min(65536, limit + 1 - total));
      const count = readSync(fd, chunk, 0, chunk.length, null);
      if (count === 0) break;
      if (total + count > limit) throw new ArchiveRefused("archive_limit");
      chunks.push(chunk.subarray(0, count));
      total += count;
    }
    const bytes = new Uint8Array(total);
    let at = 0;
    for (const chunk of chunks) { bytes.set(chunk, at); at += chunk.length; }
    return bytes;
  } finally {
    closeSync(fd);
  }
}

export function run(args: string[], io: CliIo): number {
  const [command, ...rest] = args;
  const positional: string[] = [];
  let to: string | null = null;
  let force = false;
  for (const arg of rest) {
    if (arg.startsWith("--to=")) to = arg.slice(5);
    else if (arg === "--force") force = true;
    else if (arg.startsWith("--")) {
      io.stderr(`Unknown option ${arg}\n${USAGE}`);
      return 2;
    } else positional.push(arg);
  }
  if (command !== "template:convert" || positional.length !== 2) {
    io.stderr(USAGE);
    return 2;
  }
  const [input, output] = positional as [string, string];

  if (existsSync(output) && !force) {
    io.stderr(`Refused: ${output} already exists (pass --force to replace it).\n`);
    return 1;
  }
  if (existsSync(input) && existsSync(output) && realpathSync(input) === realpathSync(output)) {
    io.stderr("Refused: the output is the input.\n");
    return 1;
  }

  let converted: Uint8Array;
  try {
    converted = Agent.convertTemplate(readFileBoundedSync(input), to as SourceKind | null);
  } catch (error) {
    if (error instanceof OfficeTemplateRefused || error instanceof ArchiveRefused) {
      io.stderr(`Refused (${error.reason}): ${error.message}\n`);
      return 1;
    }
    if (error instanceof UnsupportedFormatException) {
      io.stderr(`Refused (${error.format}): ${error.message}\n`);
      return 1;
    }
    throw error;
  }

  // Written beside the target and renamed, so a failure leaves no half file.
  const temp = `${output}.tmp-${randomBytes(4).toString("hex")}`;
  try {
    writeFileSync(temp, converted);
    renameSync(temp, output);
  } catch {
    try { unlinkSync(temp); } catch { /* already gone */ }
    io.stderr(`Could not write ${output}.\n`);
    return 1;
  }
  io.stdout(`Wrote ${output} (${sourceKind(SourceArchive.open(converted))}).\n`);
  return 0;
}
