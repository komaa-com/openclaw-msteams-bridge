import { VisionBudget } from "./vision-budget.js";
/**
 * Scene-change keyframes retained per call for retroactive vision ("what did the earlier slide
 * say?"). The worker only emits changed frames, so every stored frame IS a keyframe; the ring keeps
 * memory bounded (~16 JPEG frames ≈ 1–2 MB per call).
 */
const HISTORY_MAX_FRAMES = 16;
/**
 * Per-call inbound video frames (latest per source) plus the per-call vision spend cap. The provider
 * owns the recording gate (recording-active is shared call state), so this only stores/serves frames
 * that have already passed the gate and tracks the budget.
 */
export class MsteamsVisionStore {
    maxPerMinute;
    frames = new Map();
    history = new Map();
    budgetInstance = null;
    /** @param maxPerMinute lazy read of `msteams.maxVisionPerMinute` (config may be wired after construction). */
    constructor(maxPerMinute) {
        this.maxPerMinute = maxPerMinute;
    }
    /** Retain the latest frame per source for a call. Caller must have passed the recording gate. */
    store(frame) {
        const pair = this.frames.get(frame.callId) ?? {};
        pair[frame.source] = {
            source: frame.source,
            dataBase64: frame.dataBase64,
            mime: frame.mime,
            width: frame.width,
            height: frame.height,
            ts: frame.ts,
            participantId: frame.participantId,
            participantName: frame.participantName,
        };
        this.frames.set(frame.callId, pair);
        // Keyframe ring for retroactive vision: oldest dropped past the cap.
        const ring = this.history.get(frame.callId) ?? [];
        ring.push(pair[frame.source]);
        if (ring.length > HISTORY_MAX_FRAMES) {
            ring.shift();
        }
        this.history.set(frame.callId, ring);
    }
    /** The most recent scene-change keyframes for a call, oldest first (up to `limit`). */
    getHistory(callId, limit = HISTORY_MAX_FRAMES) {
        const ring = this.history.get(callId) ?? [];
        return ring.slice(-Math.max(1, limit));
    }
    /** Latest frame for a call; with no source, prefers screen-share over camera. */
    getLatest(callId, source) {
        const pair = this.frames.get(callId);
        if (!pair) {
            return undefined;
        }
        return source ? pair[source] : (pair.screenshare ?? pair.camera);
    }
    /** Shared per-call vision spend cap, lazily built from config on first use. */
    budget() {
        this.budgetInstance ??= new VisionBudget(this.maxPerMinute());
        return this.budgetInstance;
    }
    /** Adopt a budget supplied by the realtime runtime, so both paths share one cap. */
    setBudget(budget) {
        this.budgetInstance = budget;
    }
    /** Drop a call's frames (latest + history) and release its budget slot. */
    release(callId) {
        this.frames.delete(callId);
        this.history.delete(callId);
        this.budgetInstance?.release(callId);
    }
}
