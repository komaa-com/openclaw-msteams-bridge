// Shared vision→agent helpers used by BOTH voice paths.
//
// Two concerns:
//  1. Turning stored video frames into the `images` input the agent consult accepts
//     (`consultRealtimeVoiceAgent({ images })` → `RunEmbeddedAgentParams["images"]`).
//  2. Ambient push de-gating: the realtime bridge's `sendImage` is not in PUBLISHED openclaw, so when
//     it's absent we fall back to queueing the frame as a consult image for the next agent turn — the
//     agent still sees it on stock openclaw (no longer build-gated).

import type { MsteamsVideoFrame } from "./msteams-video-frame.js";
import type { VisionBudget } from "./vision-budget.js";

/** Agent consult image input (matches `consultRealtimeVoiceAgent({ images })`). */
export interface ConsultImage {
  type: "image";
  data: string;
  mimeType: string;
}

/** Ambient image payload for the (private) realtime-voice `sendImage`. */
export interface BridgeImagePush {
  dataBase64: string;
  mime: string;
  text: string;
}

export function frameToConsultImage(frame: { dataBase64: string; mime: string }): ConsultImage {
  return { type: "image", data: frame.dataBase64, mimeType: frame.mime };
}

/**
 * Gather the latest screen-share + camera frames as consult images, honoring the per-call vision
 * budget. Used by the streaming path to give the agent "look at what's shared" awareness on each turn.
 */
export function collectLatestFrameImages(opts: {
  getLatestFrame?: (source?: "camera" | "screenshare") => MsteamsVideoFrame | undefined;
  visionBudget?: VisionBudget;
  callId: string;
  now?: () => number;
}): ConsultImage[] {
  const { getLatestFrame, visionBudget, callId } = opts;
  if (!getLatestFrame) return [];
  const now = opts.now ?? (() => Date.now());
  const images: ConsultImage[] = [];
  for (const source of ["screenshare", "camera"] as const) {
    const frame = getLatestFrame(source);
    if (!frame) continue;
    if (visionBudget && !visionBudget.tryConsume(callId, now())) break;
    images.push(frameToConsultImage(frame));
  }
  return images;
}

/**
 * Push an ambient image into the realtime bridge when it implements `sendImage` (the `next` openclaw
 * build); otherwise queue it as a consult image for the next agent turn so the agent still sees it on
 * stock published openclaw. Returns which path was taken. Throws only if a present `sendImage` throws
 * (the caller refunds the budget and retries) — the queue path never throws.
 */
/**
 * Make consult images actually reach the model on a PUBLISHED openclaw host.
 *
 * `consultRealtimeVoiceAgent` on every published host through 2026.7.1 has NO `images` parameter - the
 * plugin passed one anyway, JavaScript silently discarded the unknown property, and every look ran as a
 * text-only consult. The model then answered, truthfully, "I'm not receiving the shared-screen image" -
 * which read exactly like a media bug and cost a full day of chasing the media path, every leg of which
 * turned out to be working.
 *
 * The host's underlying `runEmbeddedAgent` DOES accept `images` (RunEmbeddedAgentParams). So instead of
 * bypassing the consult wrapper - which owns session resolution, prompt assembly and lifecycle - hand it
 * an agentRuntime whose `runEmbeddedAgent` injects the images at the layer that understands them. If a
 * future host's consult starts forwarding images itself, the guard (`p.images` already set) makes this a
 * no-op rather than a double-attach.
 */
export function withConsultImages<R extends { runEmbeddedAgent: (p: never) => unknown }>(
  agentRuntime: R,
  images: ConsultImage[] | undefined,
): R {
  if (!images?.length) return agentRuntime;
  return {
    ...agentRuntime,
    runEmbeddedAgent: (p: { images?: unknown }) =>
      (agentRuntime.runEmbeddedAgent as (q: unknown) => unknown)(
        p && typeof p === "object" && !p.images ? { ...p, images } : p,
      ),
  } as R;
}

export function pushOrQueueBridgeImage(
  bridge: unknown,
  image: BridgeImagePush,
  queue: ConsultImage[],
): "pushed" | "queued" {
  const fn = (bridge as { sendImage?: (image: BridgeImagePush) => void }).sendImage;
  if (typeof fn === "function") {
    fn.call(bridge, image);
    return "pushed";
  }
  queue.push({ type: "image", data: image.dataBase64, mimeType: image.mime });
  return "queued";
}
