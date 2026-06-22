// Smoke test: the runtime constructs against a faked plugin API and the entry config resolves.
// Live media WS / realtime bridging is exercised in integration, not here.
import { describe, expect, it, vi } from "vitest";
import { MsteamsVoiceRuntime } from "./msteams-runtime.js";
import { resolvePluginConfig } from "./plugin-config.js";

function fakeApi() {
  const store = new Map<string, unknown>();
  const syncKeyedStore = {
    lookup: (k: string) => store.get(k),
    register: (k: string, v: unknown) => store.set(k, v),
    delete: (k: string) => store.delete(k),
    entries: () => [...store.entries()].map(([key, value]) => ({ key, value })),
  };
  const logger = { info: vi.fn(), warn: vi.fn(), error: vi.fn(), debug: vi.fn() };
  return {
    pluginConfig: {
      enabled: true,
      port: 0,
      path: "/voice/msteams/stream",
      sharedSecret: "s3cret",
      realtime: { provider: "openai" },
    },
    config: {},
    runtime: {
      state: { openSyncKeyedStore: () => syncKeyedStore },
      logging: { getChildLogger: () => logger },
      agent: {},
    },
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  } as any;
}

describe("MsteamsVoiceRuntime", () => {
  it("resolves config and constructs without throwing", () => {
    const api = fakeApi();
    const cfg = resolvePluginConfig(api.pluginConfig);
    expect(cfg.enabled).toBe(true);
    expect(cfg.media.port).toBe(0);
    expect(cfg.media.sharedSecret).toBe("s3cret");
    expect(cfg.voice.realtime.toolPolicy).toBe("none");
    expect(() => new MsteamsVoiceRuntime(api, cfg)).not.toThrow();
  });

  it("treats enabled:false as disabled", () => {
    const cfg = resolvePluginConfig({ enabled: false });
    expect(cfg.enabled).toBe(false);
  });

  it("defaults a missing config to sane values", () => {
    const cfg = resolvePluginConfig(undefined);
    expect(cfg.enabled).toBe(true);
    expect(cfg.media.path).toBe("/voice/msteams/stream");
    expect(cfg.limits.maxConcurrentCalls).toBe(4);
    expect(cfg.voice.responseTimeoutMs).toBe(30000);
  });
});
