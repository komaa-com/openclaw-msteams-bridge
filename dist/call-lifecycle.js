// Call-lifecycle coordinator — the ONLY substantial hand-written glue (DESIGN.md).
//
// Replaces the subset of voice-call's CallManager a Teams-only realtime plugin needs:
//  - in-memory active-call registry
//  - state machine + transitions (initiate/answer/end)
//  - a keyed record store (in-memory; call state is ephemeral — a gateway restart drops live calls)
//  - getStatus, stale-call reaping, max-concurrency
//
// Decoupled from openclaw by design: the store/log/clock come in via LifecycleRuntime, so this is
// unit-testable with a fake store + clock, and the runtime supplies an in-memory Map-backed store.
import { TERMINAL_STATES, } from "./types.js";
/** Cap retained transcript entries per call so a long meeting can't grow a record unbounded. */
const MAX_TRANSCRIPT_ENTRIES = 200;
const REAPER_CHECK_INTERVAL_MS = 15_000;
const STORE_NAME = "msteams-voice/calls";
const ALLOWED_TRANSITIONS = {
    initiated: ["ringing", "answered", "active", "failed", "completed"],
    ringing: ["answered", "active", "failed", "completed"],
    answered: ["active", "completed", "failed"],
    active: ["completed", "failed"],
    completed: [],
    failed: [],
};
export class MaxConcurrentCallsError extends Error {
    constructor(limit) {
        super(`msteams-voice: max concurrent calls reached (${limit})`);
        this.name = "MaxConcurrentCallsError";
    }
}
export class CallLifecycle {
    rt;
    opts;
    calls = new Map();
    store;
    reaper;
    constructor(rt, opts) {
        this.rt = rt;
        this.opts = opts;
        this.store = rt.openSyncKeyedStore(STORE_NAME);
        this.rehydrate();
    }
    /** Load non-terminal records persisted before a restart back into the active registry. */
    rehydrate() {
        for (const key of this.store.keys()) {
            const rec = this.store.get(key);
            if (!rec || TERMINAL_STATES.has(rec.state))
                continue;
            this.calls.set(rec.callId, rec);
        }
    }
    /** Start the stale-call reaper (no-op if neither timeout is configured). */
    start() {
        if (this.reaper)
            return;
        if (this.opts.staleCallReaperMs <= 0 && this.opts.maxDurationMs <= 0)
            return;
        const set = this.rt.setInterval ?? ((fn, ms) => setInterval(fn, ms));
        this.reaper = set(() => this.reapStale(), REAPER_CHECK_INTERVAL_MS);
        // Don't keep the process alive just for the reaper.
        this.reaper?.unref?.();
    }
    stop() {
        if (this.reaper)
            (this.rt.clearInterval ?? clearInterval)(this.reaper);
        this.reaper = undefined;
    }
    /** Register a placed/received call. Throws MaxConcurrentCallsError if over the limit. */
    initiate(params) {
        if (this.calls.has(params.callId)) {
            return this.calls.get(params.callId);
        }
        if (this.calls.size >= this.opts.maxConcurrentCalls) {
            throw new MaxConcurrentCallsError(this.opts.maxConcurrentCalls);
        }
        const rec = {
            callId: params.callId,
            providerCallId: params.providerCallId,
            direction: params.direction,
            state: "initiated",
            from: params.from,
            to: params.to,
            startedAt: this.rt.now(),
            transcript: [],
            message: params.message,
        };
        this.calls.set(rec.callId, rec);
        this.persist(rec);
        return rec;
    }
    /** Mark a call answered (callee picked up / Teams recording active). */
    answer(callId) {
        const rec = this.calls.get(callId);
        if (!rec || TERMINAL_STATES.has(rec.state))
            return;
        if (rec.answeredAt === undefined)
            rec.answeredAt = this.rt.now();
        this.transition(rec, "answered");
        this.transition(rec, "active");
        this.persist(rec);
    }
    /** End a call. Idempotent — a second call after a terminal state is a no-op. */
    end(callId, reason) {
        const rec = this.calls.get(callId) ?? this.store.get(callId);
        if (!rec || TERMINAL_STATES.has(rec.state))
            return;
        this.transition(rec, reason === "error" ? "failed" : "completed");
        rec.endedAt = this.rt.now();
        rec.endReason = reason;
        this.persist(rec);
        this.calls.delete(rec.callId);
    }
    getStatus(callId) {
        const rec = this.calls.get(callId) ?? this.store.get(callId);
        if (!rec)
            return undefined;
        return { state: rec.state, isTerminal: TERMINAL_STATES.has(rec.state) };
    }
    getRecord(callId) {
        return this.calls.get(callId) ?? this.store.get(callId);
    }
    activeCount() {
        return this.calls.size;
    }
    appendTranscript(callId, entry) {
        const rec = this.calls.get(callId);
        if (!rec)
            return;
        rec.transcript.push(entry);
        if (rec.transcript.length > MAX_TRANSCRIPT_ENTRIES) {
            rec.transcript.splice(0, rec.transcript.length - MAX_TRANSCRIPT_ENTRIES);
        }
        this.persist(rec);
    }
    /** End calls that never answered (staleCallReaperMs) or exceeded maxDurationMs. */
    reapStale() {
        const now = this.rt.now();
        for (const rec of [...this.calls.values()]) {
            if (TERMINAL_STATES.has(rec.state))
                continue;
            const unanswered = rec.answeredAt === undefined &&
                this.opts.staleCallReaperMs > 0 &&
                now - rec.startedAt > this.opts.staleCallReaperMs;
            const overDuration = rec.answeredAt !== undefined &&
                this.opts.maxDurationMs > 0 &&
                now - rec.answeredAt > this.opts.maxDurationMs;
            if (unanswered) {
                this.rt.log.info(`msteams-voice: reaping unanswered call ${rec.callId}`);
                this.end(rec.callId, "no-answer");
                // Signal the owner to tear down the live bridge too (see onReap): end() alone forgets the
                // record but leaves the media/realtime sockets open.
                this.opts.onReap?.(rec.callId, "no-answer");
            }
            else if (overDuration) {
                this.rt.log.info(`msteams-voice: reaping over-duration call ${rec.callId}`);
                this.end(rec.callId, "timeout");
                this.opts.onReap?.(rec.callId, "timeout");
            }
        }
    }
    // --- internals ---
    transition(rec, next) {
        if (rec.state === next)
            return;
        if (!ALLOWED_TRANSITIONS[rec.state].includes(next)) {
            this.rt.log.warn(`msteams-voice: illegal transition ${rec.state} -> ${next} (${rec.callId})`);
            return;
        }
        rec.state = next;
    }
    persist(rec) {
        this.store.set(rec.callId, rec);
    }
}
