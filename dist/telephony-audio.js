export { convertPcmToMulaw8k, resamplePcmTo8k } from "openclaw/plugin-sdk/realtime-voice";
export function chunkAudio(audio, chunkSize = 160) {
    return (function* () {
        for (let i = 0; i < audio.length; i += chunkSize) {
            yield audio.subarray(i, Math.min(i + chunkSize, audio.length));
        }
    })();
}
