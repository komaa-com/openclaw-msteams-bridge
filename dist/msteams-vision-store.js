import { VisionBudget } from "./vision-budget.js";
const HISTORY_MAX_FRAMES = 16;
export class MsteamsVisionStore {
    maxPerMinute;
    frames = new Map();
    history = new Map();
    budgetInstance = null;
    constructor(maxPerMinute) {
        this.maxPerMinute = maxPerMinute;
    }
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
        const ring = this.history.get(frame.callId) ?? [];
        ring.push(pair[frame.source]);
        if (ring.length > HISTORY_MAX_FRAMES) {
            ring.shift();
        }
        this.history.set(frame.callId, ring);
    }
    getHistory(callId, limit = HISTORY_MAX_FRAMES) {
        const ring = this.history.get(callId) ?? [];
        return ring.slice(-Math.max(1, limit));
    }
    getLatest(callId, source) {
        const pair = this.frames.get(callId);
        if (!pair) {
            return undefined;
        }
        return source ? pair[source] : (pair.screenshare ?? pair.camera);
    }
    budget() {
        this.budgetInstance ??= new VisionBudget(this.maxPerMinute());
        return this.budgetInstance;
    }
    setBudget(budget) {
        this.budgetInstance = budget;
    }
    release(callId) {
        this.frames.delete(callId);
        this.history.delete(callId);
        this.budgetInstance?.release(callId);
    }
}
