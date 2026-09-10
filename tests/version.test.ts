import { describe, it, expect } from "vitest";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { VERSION } from "../src";

/**
 * `VERSION` must not be able to lie.
 *
 * It is exported from every sibling in this family, and on ALL of them it
 * misreported: each Node constant was stale against its own `package.json`,
 * each PHP constant stale against its own CHANGELOG. This one said
 * "0.2.0" while the package shipped as something else entirely.
 *
 * None of that was carelessness. It is the predictable result of a number
 * living in two files with nothing comparing them — the same failure the
 * envelope's `kit.json` rule exists to stop.
 *
 * The comment above the constant used to call it a "feature-parity baseline
 * with PHP, bumped independently on npm", which reads as a decision and was
 * really a description of the drift: nothing ever bumped it, and nothing could
 * tell the two cases apart. `dark-slide-py` already had this assertion and was
 * the only engine in the family to catch itself, on its release preflight.
 *
 * It costs one assertion and removes the whole class.
 */
describe("the package reports the version it ships as", () => {
  const pkg = JSON.parse(readFileSync(join(__dirname, "..", "package.json"), "utf8")) as {
    version: string;
  };

  it("matches package.json", () => {
    expect(
      VERSION,
      "VERSION and package.json disagree. Fix the constant, do not relax this test.",
    ).toBe(pkg.version);
  });

  it("is a semver triple", () => {
    expect(VERSION).toMatch(/^\d+\.\d+\.\d+$/);
  });
});
