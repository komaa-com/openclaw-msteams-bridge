// Call-lifecycle coordinator — the ONLY substantial hand-written glue (DESIGN.md).
//
// Replaces the subset of voice-call's CallManager a Teams-only realtime plugin needs:
//  - in-memory active-call registry + callId <-> providerCallId map
//  - state machine + transitions (initiate/answer/end)
//  - record persistence via api.runtime.state.openSyncKeyedStore (NOT a vendored store)
//  - getStatus, stale-call reaping, max-concurrency, event dedupe
//
// 🚧 SCAFFOLD: method bodies are stubs with TODOs. ~400–600 LOC when filled in.

import {
  TERMINAL_STATES,
  type CallEndReason,
  type CallRecord,
  type CallState,
} from "./types.js";

// Minimal shapes of the api.runtime bits we use (kept local so this file documents its own surface;
// swap for the real openclaw types when wiring index.ts).
export interface SyncKeyedStore<T> {
  get(key: string): T | undefined;
  set(key: string, value: T): void;
  delete(key: string): void;
  keys(): string[];
}
export interface LifecycleRuntime {
  openSyncKeyedStore<T>(name: string): SyncKeyedStore<T>;
  log: { info: (m: string) => void; warn: (m: string) => void; error: (m: string) => void };
  now: () => number;
}

export interface CallLifecycleOptions {
  maxConcurrentCalls: number;
  maxDurationMs: number;
  staleCallReaperMs: number;
}

const ALLOWED_TRANSITIONS: Record<CallState, CallState[]> = {
  initiated: ["ringing", "answered", "failed", "completed"],
  ringing: ["answered", "failed", "completed"],
  answered: ["active", "completed", "failed"],
  active: ["completed", "failed"],
  completed: [],
  failed: [],
};

export class CallLifecycle {
  private readonly calls = new Map<string, CallRecord>();
  private readonly byProviderId = new Map<string, string>();
  private readonly store: SyncKeyedStore<CallRecord>;
  private reaper?: ReturnType<typeof setInterval>;

  constructor(
    private readonly rt: LifecycleRuntime,
    private readonly opts: CallLifecycleOptions,
  ) {
    this.store = rt.openSyncKeyedStore<CallRecord>("msteams-voice/calls");
    // TODO: rehydrate non-terminal records from the store on startup if desired.
  }

  /** Start the stale-call reaper. */
  start(): void {
    // TODO: setInterval(() => this.reapStale(), CHECK_INTERVAL); store handle in this.reaper.
  }

  stop(): void {
    if (this.reaper) clearInterval(this.reaper);
    this.reaper = undefined;
  }

  /** Register a placed/received call. Returns the new CallRecord. */
  initiate(_params: {
    callId: string;
    providerCallId: string;
    direction: CallRecord["direction"];
    from: string;
    to: string;
    message?: string;
  }): CallRecord {
    // TODO: enforce maxConcurrentCalls; create CallRecord(state:"initiated"); register in maps;
    //       persist via this.persist(record).
    throw new Error("CallLifecycle.initiate: not implemented (scaffold)");
  }

  /** Mark a call answered (callee picked up / recording active). */
  answer(_callId: string): void {
    // TODO: transition -> "answered" (then "active"), set answeredAt, persist.
    throw new Error("CallLifecycle.answer: not implemented (scaffold)");
  }

  /** End a call (caller hangup, agent hangup, timeout, error). Idempotent. */
  end(_callId: string, _reason: CallEndReason): void {
    // TODO: transition -> terminal, set endedAt/endReason, persist, then drop from in-memory maps.
    throw new Error("CallLifecycle.end: not implemented (scaffold)");
  }

  getStatus(callId: string): { state: CallState; isTerminal: boolean } | undefined {
    const rec = this.calls.get(callId) ?? this.store.get(callId);
    if (!rec) return undefined;
    return { state: rec.state, isTerminal: TERMINAL_STATES.has(rec.state) };
  }

  resolveByProviderId(providerCallId: string): CallRecord | undefined {
    const id = this.byProviderId.get(providerCallId);
    return id ? this.calls.get(id) : undefined;
  }

  /** Returns true if this event id is new (and records it); false if already processed (dedupe). */
  admitEvent(_callId: string, _eventId: string): boolean {
    // TODO: check/append record.processedEventIds; persist.
    throw new Error("CallLifecycle.admitEvent: not implemented (scaffold)");
  }

  appendTranscript(_callId: string, _entry: CallRecord["transcript"][number]): void {
    // TODO: push + persist (chunk if oversized).
    throw new Error("CallLifecycle.appendTranscript: not implemented (scaffold)");
  }

  /** End calls that never answered within staleCallReaperMs / exceeded maxDurationMs. */
  reapStale(): void {
    // TODO: iterate this.calls; end() stale/over-duration ones.
  }

  // --- internals ---
  private transition(rec: CallRecord, next: CallState): void {
    if (!ALLOWED_TRANSITIONS[rec.state].includes(next)) {
      this.rt.log.warn(`msteams-voice: illegal transition ${rec.state} -> ${next} (${rec.callId})`);
      return;
    }
    rec.state = next;
  }

  private persist(rec: CallRecord): void {
    this.store.set(rec.callId, rec);
  }
}
