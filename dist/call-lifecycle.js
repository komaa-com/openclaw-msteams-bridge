import { TERMINAL_STATES, } from "./types.js";
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
    rehydrate() {
        for (const key of this.store.keys()) {
            const rec = this.store.get(key);
            if (!rec || TERMINAL_STATES.has(rec.state))
                continue;
            this.calls.set(rec.callId, rec);
        }
    }
    start() {
        if (this.reaper)
            return;
        if (this.opts.staleCallReaperMs <= 0 && this.opts.maxDurationMs <= 0)
            return;
        const set = this.rt.setInterval ?? ((fn, ms) => setInterval(fn, ms));
        this.reaper = set(() => this.reapStale(), REAPER_CHECK_INTERVAL_MS);
        this.reaper?.unref?.();
    }
    stop() {
        if (this.reaper)
            (this.rt.clearInterval ?? clearInterval)(this.reaper);
        this.reaper = undefined;
    }
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
                this.opts.onReap?.(rec.callId, "no-answer");
            }
            else if (overDuration) {
                this.rt.log.info(`msteams-voice: reaping over-duration call ${rec.callId}`);
                this.end(rec.callId, "timeout");
                this.opts.onReap?.(rec.callId, "timeout");
            }
        }
    }
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
