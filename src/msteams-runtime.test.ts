// Smoke test: the runtime constructs against a faked plugin API and the entry config resolves.
// Live media WS / realtime bridging is exercised in integration, not here.
import { describe, expect, it, vi } from "vitest";
import { MsteamsVoiceRuntime } from "./msteams-runtime.js";
import { resolvePluginConfig } from "./plugin-config.js";

vi.mock("openclaw/plugin-sdk/ssrf-runtime", () => ({ fetchWithSsrFGuard: vi.fn() }));
import { fetchWithSsrFGuard } from "openclaw/plugin-sdk/ssrf-runtime";

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

function apiWithOutbound() {
  const api = fakeApi();
  api.pluginConfig.outbound = {
    enabled: true,
    workerBaseUrl: "http://127.0.0.1:9440",
    tenantId: "tenant-1",
    defaultMode: "notify",
  };
  return api;
}

describe("MsteamsVoiceRuntime.placeCall (outbound)", () => {
  it("posts a signed place-call and registers an outbound record", async () => {
    vi.mocked(fetchWithSsrFGuard).mockResolvedValue({
      response: { ok: true, json: async () => ({ callId: "wc-1" }), text: async () => "" },
      release: async () => {},
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any);
    const api = apiWithOutbound();
    const rt = new MsteamsVoiceRuntime(api, resolvePluginConfig(api.pluginConfig));

    const res = await rt.placeCall("user:abc-123", { message: "Your report is ready" });

    expect(res.callId).toBe("wc-1");
    expect(rt.getCallStatus("wc-1")?.state).toBe("initiated"); // not answered until WS attaches
    expect(vi.mocked(fetchWithSsrFGuard)).toHaveBeenCalledOnce();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const arg = vi.mocked(fetchWithSsrFGuard).mock.calls[0][0] as any;
    expect(String(arg.url)).toContain("/api/calls");
    expect(arg.init.headers["x-openclawteamsbridge-signature"]).toMatch(/^[0-9a-f]{64}$/);
    expect(JSON.parse(arg.init.body)).toEqual({ userObjectId: "abc-123", tenantId: "tenant-1" });
  });

  it("throws when outbound is disabled", async () => {
    const api = fakeApi(); // no outbound config
    const rt = new MsteamsVoiceRuntime(api, resolvePluginConfig(api.pluginConfig));
    await expect(rt.placeCall("user:x")).rejects.toThrow(/outbound calling is disabled/);
  });
});
