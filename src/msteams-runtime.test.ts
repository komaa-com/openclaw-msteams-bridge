// Smoke test: the runtime constructs against a faked plugin API and the entry config resolves.
// Live media WS / realtime bridging is exercised in integration, not here.
import { describe, expect, it, vi } from "vitest";
import { MsteamsBridgeRuntime } from "./msteams-runtime.js";
import { resolvePluginConfig } from "./plugin-config.js";
import { MAX_VISION_PER_MINUTE_DEFAULT } from "./vision-budget.js";

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
      path: "/msteams/calling",
      secret: "s3cret",
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

describe("MsteamsBridgeRuntime", () => {
  it("resolves config and constructs without throwing", () => {
    const api = fakeApi();
    const cfg = resolvePluginConfig(api.pluginConfig);
    expect(cfg.enabled).toBe(true);
    expect(cfg.media.port).toBe(0);
    expect(cfg.media.sharedSecret).toBe("s3cret");
    expect(cfg.voice.realtime.toolPolicy).toBe("none");
    expect(() => new MsteamsBridgeRuntime(api, cfg)).not.toThrow();
  });

  // The global post_chat_message tool resolves through postableCalls. That map was written only by the
  // REALTIME path, so in streaming mode - the mode that needs the global tool, because it has no tool
  // dispatch of its own - the tool was offered and answered "there is no active Teams call" every time.
  // Registration moved to trackManagedCall, which runs for both modes.
  it("registers a postable call in STREAMING mode, not only realtime", () => {
    const api = fakeApi();
    api.pluginConfig = {
      ...api.pluginConfig,
      mode: "streaming",
      secret: "s3cret",
      messagesPort: 0,
      managedBot: { gatewayReplyUrl: "https://gw.example.com/api/chat/reply" },
    };
    const cfg = resolvePluginConfig(api.pluginConfig);
    const runtime = new MsteamsBridgeRuntime(api, cfg) as unknown as {
      trackManagedCall: (s: unknown, add: boolean) => void;
      resolvePostableCall: () => { post?: unknown; error?: string };
    };
    const session = { callId: "call-1", threadId: "19:meeting@thread.v2", tenantId: "tenant-1" };

    expect(runtime.resolvePostableCall().error).toBeTruthy();   // nothing live yet
    runtime.trackManagedCall(session, true);
    expect(runtime.resolvePostableCall().post).toBeTypeOf("function");
    runtime.trackManagedCall(session, false);
    expect(runtime.resolvePostableCall().error).toBeTruthy();   // and cleaned up on teardown
  });

  it("does not offer the tool for a 1:1 call, which has no chat thread", () => {
    const api = fakeApi();
    api.pluginConfig = { ...api.pluginConfig, secret: "s3cret", managedBot: { gatewayReplyUrl: "https://gw/x" } };
    const cfg = resolvePluginConfig(api.pluginConfig);
    const runtime = new MsteamsBridgeRuntime(api, cfg) as unknown as {
      trackManagedCall: (s: unknown, add: boolean) => void;
      resolvePostableCall: () => { post?: unknown; error?: string };
    };
    // The worker falls back to threadId = callId on a 1:1 call; Teams cannot address that.
    runtime.trackManagedCall({ callId: "c2", threadId: "c2", tenantId: "tenant-1" }, true);
    expect(runtime.resolvePostableCall().error).toBeTruthy();
  });

  it("treats enabled:false as disabled", () => {
    const cfg = resolvePluginConfig({ enabled: false });
    expect(cfg.enabled).toBe(false);
  });

  it("defaults a missing config to sane values", () => {
    const cfg = resolvePluginConfig(undefined);
    expect(cfg.enabled).toBe(true);
    expect(cfg.media.path).toBe("/msteams/calling");
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

describe("MsteamsBridgeRuntime.placeCall (outbound)", () => {
  it("posts a signed place-call and registers an outbound record", async () => {
    vi.mocked(fetchWithSsrFGuard).mockResolvedValue({
      response: { ok: true, json: async () => ({ callId: "wc-1" }), text: async () => "" },
      release: async () => {},
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any);
    const api = apiWithOutbound();
    const rt = new MsteamsBridgeRuntime(api, resolvePluginConfig(api.pluginConfig));

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
    const rt = new MsteamsBridgeRuntime(api, resolvePluginConfig(api.pluginConfig));
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

describe("MsteamsBridgeRuntime teardown (H7 reaper + vision leak)", () => {
  function runtimeWithLiveCall() {
    const api = fakeApi();
    const cfg = resolvePluginConfig(api.pluginConfig);
    cfg.limits.maxDurationMs = 1000; // enable the over-duration reaper
    const rt = new MsteamsBridgeRuntime(api, cfg);
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

describe("MsteamsBridgeRuntime.start (realtime provider warning)", () => {
  it("warns loudly when mode:'realtime' is set but no provider resolves", async () => {
    const api = fakeApi();
    const logger = api.runtime.logging.getChildLogger();
    api.pluginConfig.mode = "realtime";
    api.pluginConfig.realtime = { provider: "openai" }; // no credentials → does not resolve
    api.pluginConfig.port = 0; // OS-assigned; avoids collisions
    // Both lanes now come up from the single `secret`, so the messages listener needs an OS-assigned
    // port too. It used to stay off here because the fixture set the calling-only `sharedSecret`; with
    // that key removed there is no half-on configuration, and a fixed 9444 collides with any real
    // gateway running on the same machine.
    api.pluginConfig.messagesPort = 0;
    const rt = new MsteamsBridgeRuntime(api, resolvePluginConfig(api.pluginConfig));
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

/**
 * The vision spend cap has to survive the trip from plugin config to the object that enforces it.
 * `VisionBudget`'s own unit tests prove 0 means OFF, but they pass just as happily if the runtime
 * never hands 0 over - and coercing it away is a one-character mistake (`||` for `??`) that reads as
 * a harmless default. These assertions go through the real constructor and read the budget the call
 * path actually uses, so they fail the moment that wiring is dropped or coerced.
 */
describe("MsteamsBridgeRuntime vision spend cap (config reaches the budget intact)", () => {
  function budgetFor(pluginConfig: Record<string, unknown>) {
    const api = fakeApi();
    api.pluginConfig = { ...api.pluginConfig, ...pluginConfig };
    const rt = new MsteamsBridgeRuntime(api, resolvePluginConfig(api.pluginConfig));
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const inner = rt as any;
    return { rt, inner, budget: inner.visionBudget, vision: inner.vision };
  }

  it("carries a configured 0 through as OFF, not as the default cap", () => {
    // The footgun in full: `|| MAX_VISION_PER_MINUTE_DEFAULT` here would turn the operator's kill
    // switch into 30 paid vision calls a minute. `??` is what keeps 0 alive.
    const { budget } = budgetFor({ maxVisionPerMinute: 0 });
    expect(budget.enabled).toBe(false);
    expect(budget.tryConsume("c1", 1000)).toBe(false);
  });

  it("carries a configured cap through exactly", () => {
    const { budget } = budgetFor({ maxVisionPerMinute: 2 });
    expect(budget.enabled).toBe(true);
    expect(budget.tryConsume("c1", 1000)).toBe(true);
    expect(budget.tryConsume("c1", 1000)).toBe(true);
    expect(budget.tryConsume("c1", 1000)).toBe(false); // the third is over the configured cap
  });

  it("falls back to the shared default when the key is absent", () => {
    const { budget } = budgetFor({});
    for (let i = 0; i < MAX_VISION_PER_MINUTE_DEFAULT; i++) {
      expect(budget.tryConsume("c1", 1000)).toBe(true);
    }
    expect(budget.tryConsume("c1", 1000)).toBe(false); // and no more than the default
  });

  it("gives the store the SAME budget instance, so look_at_screen and ambient share one window", () => {
    // Two VisionBudget objects over one call would mean two independent caps and double the spend.
    // Fails if `this.vision.setBudget(this.visionBudget)` is dropped: the store would lazily build
    // its own.
    const { budget, vision } = budgetFor({ maxVisionPerMinute: 1 });
    expect(vision.budget()).toBe(budget);
    expect(budget.tryConsume("c1", 1000)).toBe(true);
    expect(vision.budget().tryConsume("c1", 1000)).toBe(false); // the one slot was already spent
  });

  it("warns at startup when continuous vision is on but the cap switches all spend off", async () => {
    // The two keys cancel each other out and every resulting skip is debug-level, so without this the
    // operator sees a healthy startup and a bot that never once looks at anything.
    const api = fakeApi();
    const logger = api.runtime.logging.getChildLogger();
    api.pluginConfig = {
      ...api.pluginConfig,
      ambientVision: true,
      maxVisionPerMinute: 0,
      port: 0,
      messagesPort: 0,
    };
    const rt = new MsteamsBridgeRuntime(api, resolvePluginConfig(api.pluginConfig));
    await rt.start();
    try {
      expect(logger.warn).toHaveBeenCalledWith(
        expect.stringContaining("0 is the kill switch"),
      );
    } finally {
      await rt.stop();
    }
  });

  it("stays quiet when the cap is a real number", async () => {
    const api = fakeApi();
    const logger = api.runtime.logging.getChildLogger();
    api.pluginConfig = {
      ...api.pluginConfig,
      ambientVision: true,
      maxVisionPerMinute: 5,
      port: 0,
      messagesPort: 0,
    };
    const rt = new MsteamsBridgeRuntime(api, resolvePluginConfig(api.pluginConfig));
    await rt.start();
    try {
      expect(logger.warn).not.toHaveBeenCalledWith(
        expect.stringContaining("0 is the kill switch"),
      );
    } finally {
      await rt.stop();
    }
  });
});

describe("MsteamsBridgeRuntime.onSessionStart (late outbound answer)", () => {
  it("denies a late media attach for an outbound call whose answer-timeout already fired", () => {
    const api = apiWithOutbound();
    const rt = new MsteamsBridgeRuntime(api, resolvePluginConfig(api.pluginConfig));
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

/**
 * A worker is free to re-send session.start for a call that is already live (a reconnect, a retried
 * frame). CallLifecycle.initiate RETURNS the existing record rather than throwing, so before the
 * guard the duplicate fell all the way through to `calls.set(callId, createCall(...))`:
 *
 *   - the replacement handle re-seeded `recordingActive` from the session frame, discarding whatever
 *     `recording.status` had already set — session.start silently overwriting an explicit recording
 *     status, which is the race this repo just fixed at the media-stream layer; and
 *   - the previous handle was dropped with no close(), leaking a provider socket per duplicate.
 */
describe("MsteamsBridgeRuntime.onSessionStart (duplicate session.start)", () => {
  function liveCallRuntime() {
    const api = fakeApi();
    // Accept the inbound caller, so the run reaches the duplicate guard rather than policy refusal.
    api.pluginConfig = { ...api.pluginConfig, inboundPolicy: "open" };
    const rt = new MsteamsBridgeRuntime(api, resolvePluginConfig(api.pluginConfig));
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const inner = rt as any;
    inner.mode = "streaming"; // no realtime provider is configured in fakeApi
    // Stand in for the per-call handle so the assertion is about identity and teardown, not about
    // booting a real STT/TTS bridge. The guard under test is in onSessionStart itself.
    const handles: Array<{ closed: string[] }> = [];
    inner.createCall = () => {
      const handle = { closed: [] as string[], close: (r?: string) => handle.closed.push(r ?? "") };
      handles.push(handle);
      return handle;
    };
    return { api, rt, inner, handles };
  }

  function inboundSession(recordingStatus: "active" | "inactive") {
    return {
      callId: "c-dup",
      threadId: "thread-dup",
      tenantId: "tenant-1",
      caller: { aadId: "aad-dup", displayName: "Caller" },
      recordingStatus,
      send: () => true,
      close: vi.fn(),
    };
  }

  it("ignores a second session.start and keeps the original call handle", () => {
    const { inner, handles } = liveCallRuntime();

    inner.onSessionStart(inboundSession("active"));
    expect(handles).toHaveLength(1);
    const first = inner.calls.get("c-dup");
    expect(first).toBe(handles[0]);

    // The duplicate arrives, reporting the STALE setup-time snapshot.
    inner.onSessionStart(inboundSession("inactive"));

    expect(handles).toHaveLength(1); // no second handle was built...
    expect(inner.calls.get("c-dup")).toBe(first); // ...and the live one was not replaced
    expect(inner.calls.size).toBe(1);
    expect(handles[0]!.closed).toEqual([]); // nor orphaned: nothing leaked, nothing torn down
  });

  it("says so at warn, so a worker stuck in a session.start loop is visible", () => {
    const { api, inner } = liveCallRuntime();
    const logger = api.runtime.logging.getChildLogger();
    inner.onSessionStart(inboundSession("active"));
    logger.warn.mockClear();
    inner.onSessionStart(inboundSession("active"));
    expect(logger.warn).toHaveBeenCalledWith(expect.stringContaining("duplicate session.start"));
  });
});

describe("MsteamsBridgeRuntime.finalizeUnansweredOutbound (H7a cancel-by-callId)", () => {
  it("sends a signed DELETE /api/calls/{callId} to cancel the still-ringing outbound", () => {
    vi.mocked(fetchWithSsrFGuard).mockClear();
    vi.mocked(fetchWithSsrFGuard).mockResolvedValue({
      response: { ok: true, text: async () => "" },
      release: async () => {},
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any);
    const api = apiWithOutbound();
    const rt = new MsteamsBridgeRuntime(api, resolvePluginConfig(api.pluginConfig));
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const inner = rt as any;
    inner.mode = "streaming";
    inner.lifecycle.initiate({
      callId: "wc-42",
      providerCallId: "wc-42",
      direction: "outbound",
      from: "",
      to: "user:callee",
      message: "hi",
    });
    inner.pendingOutbound.set("wc-42", { to: "user:callee", message: "hi", mode: "notify" });

    inner.finalizeUnansweredOutbound("wc-42");

    // Best-effort cancel is fired synchronously (up to its first await) when finalizing.
    expect(vi.mocked(fetchWithSsrFGuard)).toHaveBeenCalledTimes(1);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const arg = vi.mocked(fetchWithSsrFGuard).mock.calls[0][0] as any;
    expect(arg.init.method).toBe("DELETE");
    expect(String(arg.url)).toContain("/api/calls/wc-42");
    expect(arg.init.headers["x-openclawteamsbridge-signature"]).toMatch(/^[0-9a-f]{64}$/);
    expect(arg.init.headers["x-openclawteamsbridge-timestamp"]).toMatch(/^\d+$/);
  });
});
