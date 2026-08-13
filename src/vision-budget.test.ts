import { describe, expect, it } from "vitest";
import { VisionBudget } from "./vision-budget.js";

describe("VisionBudget", () => {
  it("0 turns vision spend OFF (it is the kill switch, not 'unlimited')", () => {
    // Regression, and the reason this class changed: 0 used to mean UNLIMITED, so an operator who set
    // maxVisionPerMinute to 0 to switch vision off got uncapped spend on the one key that costs money.
    const b = new VisionBudget(0);
    expect(b.enabled).toBe(false);
    for (let i = 0; i < 100; i++) {
      expect(b.tryConsume("c", 1000 + i)).toBe(false);
    }
  });

  it("a negative cap is off too, not a wrap-around to unlimited", () => {
    const b = new VisionBudget(-1);
    expect(b.enabled).toBe(false);
    expect(b.tryConsume("c", 1000)).toBe(false);
  });

  it("a large cap is how you ask for 'effectively unlimited'", () => {
    const b = new VisionBudget(10_000);
    expect(b.enabled).toBe(true);
    for (let i = 0; i < 100; i++) {
      expect(b.tryConsume("c", 1000 + i)).toBe(true);
    }
  });

  it("caps to maxPerMinute within the sliding window", () => {
    const b = new VisionBudget(2);
    expect(b.tryConsume("c", 1000)).toBe(true);
    expect(b.tryConsume("c", 1100)).toBe(true);
    expect(b.tryConsume("c", 1200)).toBe(false); // 3rd within the minute
  });

  it("frees up as the window slides past 60s", () => {
    const b = new VisionBudget(1);
    expect(b.tryConsume("c", 0)).toBe(true);
    expect(b.tryConsume("c", 30_000)).toBe(false); // still within 60s
    expect(b.tryConsume("c", 61_000)).toBe(true); // first hit aged out
  });

  it("tracks calls independently", () => {
    const b = new VisionBudget(1);
    expect(b.tryConsume("a", 0)).toBe(true);
    expect(b.tryConsume("b", 0)).toBe(true); // different call, own window
    expect(b.tryConsume("a", 0)).toBe(false);
  });

  it("refund returns the most recent hit (failed vision call does not count)", () => {
    const b = new VisionBudget(1);
    expect(b.tryConsume("c", 0)).toBe(true);
    b.refund("c"); // the send failed — the spend never happened
    expect(b.tryConsume("c", 0)).toBe(true);
    expect(b.tryConsume("c", 0)).toBe(false);
    b.refund("unknown"); // no-op for an untracked call
  });

  it("release clears a call's window", () => {
    const b = new VisionBudget(1);
    expect(b.tryConsume("c", 0)).toBe(true);
    expect(b.tryConsume("c", 0)).toBe(false);
    b.release("c");
    expect(b.tryConsume("c", 0)).toBe(true);
  });
});
