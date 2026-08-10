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

describe("flat messages-lane keys", () => {
  it("accepts gatewayReplyUrl at the root, as the config reference documents it", () => {
    // The docs put it at the root beside messagesPort/messagesPath; only managedBot.gatewayReplyUrl was
    // wired, and with additionalProperties:false a config that followed the docs failed validation.
    const cfg = resolvePluginConfig({
      enabled: true,
      secret: "s3cret",
      gatewayReplyUrl: "https://self-hosted.example.com/api/chat/reply",
    });
    expect(cfg.managedChat.gatewayReplyUrl).toBe("https://self-hosted.example.com/api/chat/reply");
  });

  it("the flat key wins over the compatibility block", () => {
    const cfg = resolvePluginConfig({
      enabled: true,
      secret: "s3cret",
      managedBot: { gatewayReplyUrl: "https://old.example.com/x" },
      gatewayReplyUrl: "https://new.example.com/y",
    });
    expect(cfg.managedChat.gatewayReplyUrl).toBe("https://new.example.com/y");
  });

  it("the messages lane defaults to loopback, like calling", () => {
    // Undefined was handed to server.listen(port, undefined), which binds every interface - so a
    // config naming no bind address got calling on loopback and messages on the LAN.
    const cfg = resolvePluginConfig({ enabled: true, secret: "s3cret" });
    expect(cfg.managedChat.bindAddress).toBe("127.0.0.1");
  });
});
