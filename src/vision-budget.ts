/**
 * Per-call vision spend cap. Vision-model calls (look_at_screen, the ambient frame push, the
 * per-turn frame attach on the streaming path) are the dominant cost of "continuous perception", so
 * bound them with a simple sliding 60-second window per call.
 *
 * `maxPerMinute <= 0` means OFF: no vision spend at all.
 *
 * That is a deliberate flip of what this class used to do, and it is the only reading of the number
 * that is safe. It previously treated 0 as UNLIMITED, so an operator who set `maxVisionPerMinute: 0`
 * to switch vision off got uncapped spend instead - the exact opposite of the intent, on the one key
 * that costs money. There is no "unlimited" setting any more; a live call always has a ceiling. Set a
 * large number if you want one that never bites in practice.
 *
 * The cap is not the feature switch for CONTINUOUS vision either - that is the separate
 * `ambientVision` flag, which is off by default. This class caps every paid look, requested or not.
 *
 * Pure-ish (the caller injects `nowMs`) so it is unit-testable.
 */
/**
 * Default cap when `msteams.maxVisionPerMinute` is not set. One place, so the runtime, the store and
 * the docs cannot drift - and so the fallback is applied with `??` (a configured 0 must survive) in
 * exactly one expression per reader.
 */
export const MAX_VISION_PER_MINUTE_DEFAULT = 30;

export class VisionBudget {
  private readonly hitsByCall = new Map<string, number[]>();

  constructor(private readonly maxPerMinute: number) {}

  /** Whether any vision spend is permitted at all. False when the cap is 0 (or negative). */
  get enabled(): boolean {
    return this.maxPerMinute > 0;
  }

  /** True (and records a hit) if under budget for this call; false if the caller should skip the vision call. */
  tryConsume(callId: string, nowMs: number): boolean {
    if (this.maxPerMinute <= 0) {
      return false; // vision spend is off
    }
    const recent = (this.hitsByCall.get(callId) ?? []).filter((t) => nowMs - t < 60_000);
    if (recent.length >= this.maxPerMinute) {
      this.hitsByCall.set(callId, recent); // keep the trimmed window
      return false;
    }
    recent.push(nowMs);
    this.hitsByCall.set(callId, recent);
    return true;
  }

  /**
   * Return the most recent hit for a call — the vision call it paid for never actually happened
   * (e.g. the frame push threw before reaching the model), so the spend should not count.
   */
  refund(callId: string): void {
    this.hitsByCall.get(callId)?.pop();
  }

  /** Drop a call's window when it ends. */
  release(callId: string): void {
    this.hitsByCall.delete(callId);
  }
}
