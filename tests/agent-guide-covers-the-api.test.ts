import { describe, expect, it } from "vitest";
import { Agent } from "../src";

/**
 * The agent guide names every public Agent method, or the method is on the
 * allowlist below with the reason it is left out. A method the guide never
 * names is a method no model reading the guide will use, so this fails the
 * build when a release adds one and the guide is not brought along.
 *
 * "Names" means written as code: `Agent::x`, `Agent.x` or a backticked `x`, so
 * the English word "read" in a sentence does not count as documenting read().
 * The PHP and Python packages carry the same test over their own surfaces.
 */
const ALLOWLIST: Record<string, string> = {
  guide: "the guide itself: a model reading it already has it",
  version: "package metadata; it has no bearing on building or reading a document",
};

const methods = (): string[] =>
  Object.entries(Agent)
    .filter(([, v]) => typeof v === "function")
    .map(([k]) => k);

const escape = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
const names = (guide: string, method: string): boolean =>
  new RegExp("(Agent::|Agent\\.|`)" + escape(method) + "\\b").test(guide);

describe("agent guide covers the API", () => {
  it("names every public Agent method, or allowlists it with a reason", () => {
    const guide = Agent.guide();
    const missing = methods().filter((m) => !(m in ALLOWLIST) && !names(guide, m));
    expect(missing, `agent-guide/professional-documents.md never names: ${missing.join(", ")}`).toEqual([]);
  });

  it("allowlists only methods that exist, each with a reason, none the guide already names", () => {
    const guide = Agent.guide();
    const all = methods();
    for (const [method, reason] of Object.entries(ALLOWLIST)) {
      expect(all).toContain(method);
      expect(reason.trim()).not.toBe("");
      expect(names(guide, method), `${method} is named in the guide; drop it from the allowlist`).toBe(false);
    }
  });
});
