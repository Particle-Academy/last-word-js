import { describe, expect, it } from "vitest";
import { chartParseNum } from "../src/helpers/chart";

describe("chartParseNum", () => {
  it("reads a very long run of digits in linear time", () => {
    const started = performance.now();
    expect(chartParseNum(`${"9".repeat(100_000)}x`)).toBe(0);
    expect(performance.now() - started).toBeLessThan(1000);
  });

  it("reads every shape of decimal, and anything else as 0", () => {
    for (const [text, want] of [
      ["5", 5],
      ["5.", 5],
      ["5.25", 5.25],
      [".5", 0.5],
      ["-5.5", -5.5],
      ["+5", 5],
      ["1e2", 100],
      ["5.e2", 500],
      [".", 0],
      ["0x10", 0],
      ["1.2.3", 0],
    ] as const) {
      expect(chartParseNum(text), text).toBe(want);
    }
  });
});
