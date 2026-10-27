#!/usr/bin/env node
/** The `last-word` bin. All behaviour lives in `cli-run.ts`, which the tests import. */
import { run } from "./cli-run";

process.exitCode = run(process.argv.slice(2), {
  stdout: (text) => process.stdout.write(text),
  stderr: (text) => process.stderr.write(text),
});
