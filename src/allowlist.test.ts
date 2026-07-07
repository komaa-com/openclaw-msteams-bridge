import { describe, expect, it } from "vitest";
import {
  describeInboundRejection,
  isAllowlistedCaller,
  isInboundCallAllowed,
  normalizePhoneNumber,
} from "./allowlist.js";

describe("normalizePhoneNumber", () => {
  it("strips non-digits", () => {
    expect(normalizePhoneNumber("+1 (415) 555-0100")).toBe("14155550100");
    expect(normalizePhoneNumber(undefined)).toBe("");
  });
});

describe("isAllowlistedCaller", () => {
  it("matches a Teams aadId exactly (case-insensitive)", () => {
    const aad = "AB12CD34-0000-0000-0000-0000DEAD";
    expect(isAllowlistedCaller(aad, [aad.toLowerCase()])).toBe(true);
    expect(isAllowlistedCaller(aad, ["someone-else"])).toBe(false);
  });
  it("matches phone numbers by digits only (exact digit sequence)", () => {
    // Same digits, different formatting → match.
    expect(isAllowlistedCaller("+1 (415) 555-0100", ["1-415-555-0100"])).toBe(true);
    // Differing digits (country code present on one side only) → no fuzzy match.
    expect(isAllowlistedCaller("+1-415-555-0100", ["(415) 555 0100"])).toBe(false);
  });
  it("rejects empty caller or empty allowlist", () => {
    expect(isAllowlistedCaller("", ["x"])).toBe(false);
    expect(isAllowlistedCaller("x", [])).toBe(false);
    expect(isAllowlistedCaller("x", undefined)).toBe(false);
  });
});

describe("isInboundCallAllowed", () => {
  it("'open' allows anyone", () => {
    expect(isInboundCallAllowed("open", undefined, "anyone")).toBe(true);
  });
  it("'allowlist'/'pairing' gate on the allowlist", () => {
    expect(isInboundCallAllowed("allowlist", ["caller-1"], "caller-1")).toBe(true);
    expect(isInboundCallAllowed("allowlist", ["caller-1"], "caller-2")).toBe(false);
    expect(isInboundCallAllowed("pairing", ["caller-1"], "caller-1")).toBe(true);
  });
  it("'disabled' or unset rejects (defensive)", () => {
    expect(isInboundCallAllowed("disabled", ["caller-1"], "caller-1")).toBe(false);
    expect(isInboundCallAllowed(undefined, ["caller-1"], "caller-1")).toBe(false);
  });
});

describe("describeInboundRejection", () => {
  it("names the rejected caller and the policy", () => {
    const msg = describeInboundRejection("allowlist", "aad-123");
    expect(msg).toContain('policy "allowlist"');
    expect(msg).toContain('caller "aad-123"');
    expect(msg).toContain("allowFrom");
  });
  it("is honest about 'pairing' being a plain allowlist and says how to fix it", () => {
    const msg = describeInboundRejection("pairing", "aad-123");
    expect(msg).toContain('policy "pairing"');
    expect(msg).toContain('caller "aad-123"');
    expect(msg).toContain("no pairing codes");
    expect(msg).toContain("add the caller's AAD object id to allowFrom");
    expect(msg).toContain("gateway's pairing flow");
  });
  it("handles a missing caller id and an unset policy", () => {
    const msg = describeInboundRejection(undefined, "");
    expect(msg).toContain('policy "disabled"');
    expect(msg).toContain("no caller id");
  });
});
