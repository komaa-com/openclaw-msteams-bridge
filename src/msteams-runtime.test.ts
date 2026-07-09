// Smoke test: the runtime constructs against a faked plugin API and the entry config resolves.
// Live media WS / realtime bridging is exercised in integration, not here.
import { describe, expect, it, vi } from "vitest";
import { MsteamsVoiceRuntime } from "./msteams-runtime.js";
import { resolvePluginConfig } from "./plugin-config.js";

vi.mock("openclaw/plugin-sdk/ssrf-runtime", () => ({ fetchWithSsrFGuard: vi.fn() }));
import { fetchWithSsrFGuard } from "openclaw/plugin-sdk/ssrf-runtime";

// No realtime provider resolves (no credentials) — lets us exercise the mode:"realtime" startup
// warning deterministically without depending on ambient provider env vars.
vi.mock("openclaw/plugin-sdk/realtime-voice", () => ({
  resolveConfiguredRealtimeVoiceProvider: vi.fn(() => ({})),
  consultRealtimeVoiceAgent: vi.fn(),
  resolveRealtimeVoiceAgentConsultToolsAllow: vi.fn(() => []),
}));

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

  it("maps the voice mode that selects realtime vs streaming", () => {
    expect(resolvePluginConfig({ mode: "streaming" }).voice.mode).toBe("streaming");
    expect(resolvePluginConfig({ mode: "realtime" }).voice.mode).toBe("realtime");
    // Unset → runtime decides (realtime if a provider resolves, else streaming).
    expect(resolvePluginConfig({}).voice.mode).toBeUndefined();
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

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function fakeVideoFrame(callId: string): any {
  return {
    callId,
    source: "camera",
    dataBase64: "AAAA",
    mime: "image/jpeg",
    width: 2,
    height: 2,
    ts: 1,
  };
}

describe("MsteamsVoiceRuntime teardown (H7 reaper + vision leak)", () => {
  function runtimeWithLiveCall() {
    const api = fakeApi();
    const cfg = resolvePluginConfig(api.pluginConfig);
    cfg.limits.maxDurationMs = 1000; // enable the over-duration reaper
    const rt = new MsteamsVoiceRuntime(api, cfg);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const inner = rt as any;
    const lifecycle = inner.lifecycle;
    const calls: Map<string, { close: (r?: string) => void }> = inner.calls;
    const vision = inner.vision;
    const closedReasons: Array<string | undefined> = [];

    lifecycle.initiate({ callId: "c1", providerCallId: "c1", direction: "inbound", from: "a", to: "" });
    lifecycle.answer("c1");
    calls.set("c1", { close: (r?: string) => closedReasons.push(r) });
    vision.store(fakeVideoFrame("c1"));
    expect(vision.getLatest("c1")).toBeDefined();
    return { rt, inner, lifecycle, calls, vision, closedReasons };
  }

  it("reaping an over-duration call closes its bridge + frees frames (no zombie / no gate bypass)", () => {
    const { lifecycle, calls, vision, closedReasons } = runtimeWithLiveCall();
    lifecycle.getRecord("c1").answeredAt = 0; // force past maxDurationMs
    lifecycle.reapStale();

    expect(closedReasons.length).toBe(1); // the media/realtime bridge was torn down
    expect(closedReasons[0]).toBe("timeout"); // reason passed → Teams worker session closed too
    expect(calls.has("c1")).toBe(false); // dropped from the active registry
    expect(vision.getLatest("c1")).toBeUndefined(); // per-call frames released (leak fixed)
    expect(lifecycle.activeCount()).toBe(0); // gate accounting stays consistent
  });

  it("a caller hangup also releases the call's retained vision frames (leak fix)", () => {
    const { inner, calls, vision, closedReasons } = runtimeWithLiveCall();
    inner.onSessionEnd({ callId: "c1", reason: "hangup" });

    expect(closedReasons).toEqual([undefined]); // caller hangup: session already closing → no reason
    expect(calls.has("c1")).toBe(false);
    expect(vision.getLatest("c1")).toBeUndefined();
  });
});

describe("MsteamsVoiceRuntime.start (realtime provider warning)", () => {
  it("warns loudly when mode:'realtime' is set but no provider resolves", async () => {
    const api = fakeApi();
    const logger = api.runtime.logging.getChildLogger();
    api.pluginConfig.mode = "realtime";
    api.pluginConfig.realtime = { provider: "openai" }; // no credentials → does not resolve
    api.pluginConfig.port = 0; // OS-assigned; avoids collisions
    const rt = new MsteamsVoiceRuntime(api, resolvePluginConfig(api.pluginConfig));
    await rt.start();
    try {
      expect(logger.warn).toHaveBeenCalledWith(
        expect.stringContaining('mode is "realtime" but no realtime voice provider resolved'),
      );
    } finally {
      await rt.stop();
    }
  });
});

describe("MsteamsVoiceRuntime.onSessionStart (late outbound answer)", () => {
  it("denies a late media attach for an outbound call whose answer-timeout already fired", () => {
    const api = apiWithOutbound();
    const rt = new MsteamsVoiceRuntime(api, resolvePluginConfig(api.pluginConfig));
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const inner = rt as any;
    inner.mode = "streaming"; // skip the realtime-provider guard so we reach the late-answer branch

    inner.lifecycle.initiate({
      callId: "wc-1",
      providerCallId: "wc-1",
      direction: "outbound",
      from: "",
      to: "user:callee",
      message: "hi",
    });
    inner.pendingOutbound.set("wc-1", { to: "user:callee", message: "hi", mode: "notify" });
    inner.finalizeUnansweredOutbound("wc-1"); // answer window elapsed → pending gone, record terminal

    const closed: string[] = [];
    const session = {
      callId: "wc-1",
      threadId: "t",
      caller: { aadId: "callee" },
      send: () => true,
      close: (r: string) => closed.push(r),
    };
    inner.onSessionStart(session);

    expect(closed).toEqual(["answer-timeout"]); // denied, not mis-routed to inbound
    expect(inner.calls.has("wc-1")).toBe(false); // no call handle was created
  });
});
