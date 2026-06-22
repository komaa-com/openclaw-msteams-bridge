import { describe, expect, it } from "vitest";
import {
  CallLifecycle,
  MaxConcurrentCallsError,
  type LifecycleRuntime,
  type SyncKeyedStore,
} from "./call-lifecycle.js";
import type { CallRecord } from "./types.js";

function fakeStore(seed: CallRecord[] = []): SyncKeyedStore<CallRecord> {
  const m = new Map<string, CallRecord>(seed.map((r) => [r.callId, r]));
  return {
    get: (k) => m.get(k),
    set: (k, v) => void m.set(k, v),
    delete: (k) => void m.delete(k),
    keys: () => [...m.keys()],
  };
}

function harness(opts?: Partial<ConstructorParameters<typeof CallLifecycle>[1]>, seed: CallRecord[] = []) {
  let t = 1000;
  const store = fakeStore(seed);
  const rt: LifecycleRuntime = {
    openSyncKeyedStore: <T>() => store as unknown as SyncKeyedStore<T>,
    log: { info: () => {}, warn: () => {}, error: () => {} },
    now: () => t,
    setInterval: () => 0 as unknown as ReturnType<typeof setInterval>,
    clearInterval: () => {},
  };
  const lc = new CallLifecycle(rt, {
    maxConcurrentCalls: 5,
    maxDurationMs: 300_000,
    staleCallReaperMs: 120_000,
    ...opts,
  });
  return { lc, store, advance: (ms: number) => (t += ms) };
}

const init = { callId: "c1", providerCallId: "p1", direction: "inbound" as const, from: "a", to: "b" };

describe("CallLifecycle", () => {
  it("initiate registers + persists; getStatus/resolveByProviderId/activeCount work", () => {
    const { lc, store } = harness();
    const rec = lc.initiate(init);
    expect(rec.state).toBe("initiated");
    expect(lc.getStatus("c1")).toEqual({ state: "initiated", isTerminal: false });
    expect(lc.resolveByProviderId("p1")?.callId).toBe("c1");
    expect(lc.activeCount()).toBe(1);
    expect(store.get("c1")?.state).toBe("initiated");
  });

  it("initiate is idempotent on the same callId", () => {
    const { lc } = harness();
    const a = lc.initiate(init);
    const b = lc.initiate(init);
    expect(a).toBe(b);
    expect(lc.activeCount()).toBe(1);
  });

  it("enforces maxConcurrentCalls", () => {
    const { lc } = harness({ maxConcurrentCalls: 1 });
    lc.initiate(init);
    expect(() => lc.initiate({ ...init, callId: "c2", providerCallId: "p2" })).toThrow(
      MaxConcurrentCallsError,
    );
  });

  it("answer sets answeredAt and moves to active", () => {
    const { lc } = harness();
    lc.initiate(init);
    lc.answer("c1");
    expect(lc.getStatus("c1")).toEqual({ state: "active", isTerminal: false });
    expect(lc.getRecord("c1")?.answeredAt).toBeDefined();
  });

  it("end is terminal + idempotent; record stays in the store, leaves the active registry", () => {
    const { lc } = harness();
    lc.initiate(init);
    lc.end("c1", "hangup-user");
    expect(lc.getStatus("c1")).toEqual({ state: "completed", isTerminal: true });
    expect(lc.activeCount()).toBe(0);
    expect(lc.resolveByProviderId("p1")).toBeUndefined();
    lc.end("c1", "error"); // no-op
    expect(lc.getRecord("c1")?.endReason).toBe("hangup-user");
  });

  it("end with error -> failed", () => {
    const { lc } = harness();
    lc.initiate(init);
    lc.end("c1", "error");
    expect(lc.getStatus("c1")?.state).toBe("failed");
  });

  it("admitEvent dedupes by event id", () => {
    const { lc } = harness();
    lc.initiate(init);
    expect(lc.admitEvent("c1", "e1")).toBe(true);
    expect(lc.admitEvent("c1", "e1")).toBe(false);
    expect(lc.admitEvent("c1", "e2")).toBe(true);
  });

  it("appendTranscript records + caps growth", () => {
    const { lc } = harness();
    lc.initiate(init);
    for (let i = 0; i < 250; i++) lc.appendTranscript("c1", { role: "caller", text: `t${i}`, at: i });
    expect(lc.getRecord("c1")!.transcript.length).toBe(200);
    expect(lc.getRecord("c1")!.transcript.at(-1)?.text).toBe("t249");
  });

  it("reapStale ends unanswered calls past staleCallReaperMs (no-answer)", () => {
    const { lc, advance } = harness({ staleCallReaperMs: 1000 });
    lc.initiate(init);
    advance(1500);
    lc.reapStale();
    expect(lc.getStatus("c1")).toEqual({ state: "completed", isTerminal: true });
    expect(lc.getRecord("c1")?.endReason).toBe("no-answer");
  });

  it("reapStale ends answered calls past maxDurationMs (timeout)", () => {
    const { lc, advance } = harness({ maxDurationMs: 1000 });
    lc.initiate(init);
    lc.answer("c1");
    advance(1500);
    lc.reapStale();
    expect(lc.getRecord("c1")?.endReason).toBe("timeout");
  });

  it("rehydrates non-terminal records from the store on construction", () => {
    const seed: CallRecord[] = [
      {
        callId: "c1", providerCallId: "p1", direction: "inbound", state: "active",
        from: "a", to: "b", startedAt: 1, answeredAt: 2, transcript: [], processedEventIds: [],
      },
      {
        callId: "c2", providerCallId: "p2", direction: "inbound", state: "completed",
        from: "a", to: "b", startedAt: 1, transcript: [], processedEventIds: [],
      },
    ];
    const { lc } = harness({}, seed);
    expect(lc.activeCount()).toBe(1); // only the non-terminal one
    expect(lc.resolveByProviderId("p1")?.callId).toBe("c1");
  });
});
