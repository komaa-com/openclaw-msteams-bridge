import { describe, expect, it } from "vitest";
import { resolvePluginConfig } from "./plugin-config.js";

describe("resolvePluginConfig", () => {
  it("passes a string sharedSecret through unchanged", () => {
    const resolved = resolvePluginConfig({ sharedSecret: "s3cret" });
    expect(resolved.media.sharedSecret).toBe("s3cret");
  });

  it("fails closed on a non-string sharedSecret (object)", () => {
    // Regression: the manifest allows a secret-input reference object. If the host ever
    // passes it through UNRESOLVED, String({...}) would yield "[object Object]" -- a
    // non-empty, guessable secret. It must coerce to "" so index.ts refuses to start.
    const resolved = resolvePluginConfig({ sharedSecret: { env: "MSTEAMS_SHARED_SECRET" } });
    expect(resolved.media.sharedSecret).toBe("");
  });

  it("fails closed on a non-string sharedSecret (number)", () => {
    const resolved = resolvePluginConfig({ sharedSecret: 12345 });
    expect(resolved.media.sharedSecret).toBe("");
  });

  it("fails closed when sharedSecret is missing entirely", () => {
    expect(resolvePluginConfig({}).media.sharedSecret).toBe("");
    expect(resolvePluginConfig(undefined).media.sharedSecret).toBe("");
  });
});
