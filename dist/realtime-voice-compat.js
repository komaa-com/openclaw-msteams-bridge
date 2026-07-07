// Private compat for two realtime-voice members our CVI provider uses that are NOT yet in the
// PUBLISHED openclaw typings: `sendImage` on the realtime bridge session, and `mediaPaths` on the
// agent-consult result. Centralized here so call sites stay clean and there is ONE documented place
// for the SDK gap.
//
// Runtime behaviour is optional/graceful: if the installed realtime provider doesn't implement
// sendImage, the ambient-vision push silently no-ops until the SDK addition ships in published
// openclaw (then this file can be deleted and the members used directly).
/** Push an ambient image into the realtime bridge if it supports sendImage; no-op otherwise. */
export function sendBridgeImage(bridge, image) {
    bridge.sendImage?.(image);
}
/** Read `mediaPaths` off a consult result if present; empty array otherwise. */
export function consultMediaPaths(result) {
    return result.mediaPaths ?? [];
}
