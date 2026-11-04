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
const AGENT_GUIDE_SHA256 = "5f897d2a73063379598a31aad83aea7935c84ffc4a23026d8ffaf3e5034aa31a";
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
