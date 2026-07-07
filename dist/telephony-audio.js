// Voice Call plugin module implements telephony audio behavior.
export { convertPcmToMulaw8k, resamplePcmTo8k } from "openclaw/plugin-sdk/realtime-voice";
/**
 * Chunk audio buffer into 20ms frames for streaming (8kHz mono mu-law).
 */
export function chunkAudio(audio, chunkSize = 160) {
    return (function* () {
        for (let i = 0; i < audio.length; i += chunkSize) {
            yield audio.subarray(i, Math.min(i + chunkSize, audio.length));
        }
    })();
}
