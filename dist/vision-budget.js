export const MAX_VISION_PER_MINUTE_DEFAULT = 30;
export class VisionBudget {
    maxPerMinute;
    hitsByCall = new Map();
    constructor(maxPerMinute) {
        this.maxPerMinute = maxPerMinute;
    }
    get enabled() {
        return this.maxPerMinute > 0;
    }
    tryConsume(callId, nowMs) {
        if (this.maxPerMinute <= 0) {
            return false;
        }
        const recent = (this.hitsByCall.get(callId) ?? []).filter((t) => nowMs - t < 60_000);
        if (recent.length >= this.maxPerMinute) {
            this.hitsByCall.set(callId, recent);
            return false;
        }
        recent.push(nowMs);
        this.hitsByCall.set(callId, recent);
        return true;
    }
    refund(callId) {
        this.hitsByCall.get(callId)?.pop();
    }
    release(callId) {
        this.hitsByCall.delete(callId);
    }
}
