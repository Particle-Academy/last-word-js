import { readFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { Agent } from "../src";

/**
 * The agent guide ships, and is the same bytes in all three engines. Each
 * engine pins this hash over LF-normalised content; edit all three copies,
 * then all three constants.
 */
const AGENT_GUIDE_SHA256 = "6d70d774cf630c0c40813d70ddd00191cc8a0b57a2cd3c603cae54f64a9501fa";
const lf = (s: string): string => s.replace(/\r\n/g, "\n");

describe("agent guide", () => {
  it("returns the guide, and it is the shared one", () => {
    const guide = Agent.guide();
    expect(createHash("sha256").update(lf(guide), "utf8").digest("hex")).toBe(AGENT_GUIDE_SHA256);
    expect(guide.startsWith("# Agent guide:")).toBe(true);
  });

  it("embeds exactly the markdown it publishes in agent-guide/", () => {
    const published = readFileSync(join(__dirname, "..", "agent-guide", "professional-documents.md"), "utf8");
    expect(lf(Agent.guide())).toBe(lf(published));
  });
});
