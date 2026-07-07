/**
 * Microsoft Teams TTS adapter.
 *
 * The Teams bridge consumes raw PCM 16 kHz, 16-bit mono LE audio, whereas the
 * shared telephony TTS path (`telephony-tts.ts`) emits 8 kHz mu-law for Twilio
 * Media Streams. To keep upstream files untouched, the msteams-specific
 * behavior — synthesize raw PCM, then resample to 16 kHz — lives here. It
 * reuses the upstream host TTS runtime (`TelephonyTtsRuntime.textToSpeechTelephony`)
 * and the SDK resampler instead of duplicating synthesis logic.
 */
import { resamplePcm } from "openclaw/plugin-sdk/realtime-voice";
import { deepMergeDefined } from "./deep-merge.js";
import { MSTEAMS_PCM_SAMPLE_RATE_HZ } from "./msteams-media-stream.js";
/** Teams wire format: PCM 16 kHz, 16-bit mono, little-endian. */
export const MSTEAMS_TTS_SAMPLE_RATE_HZ = MSTEAMS_PCM_SAMPLE_RATE_HZ;
export function createMsteamsTtsProvider(params) {
    const { coreConfig, ttsOverride, runtime, logger } = params;
    const mergedConfig = applyTtsOverride(coreConfig, ttsOverride);
    const synthesizePcm16kWithTiming = async (text) => {
        const result = await runtime.textToSpeechTelephony({ text, cfg: mergedConfig });
        if (!result.success || !result.audioBuffer || !result.sampleRate) {
            throw new Error(result.error ?? "msteams TTS synthesis failed");
        }
        if (result.fallbackFrom && result.provider && result.fallbackFrom !== result.provider) {
            const attemptedChain = result.attemptedProviders && result.attemptedProviders.length > 0
                ? result.attemptedProviders.join(" -> ")
                : `${result.fallbackFrom} -> ${result.provider}`;
            logger?.warn?.(`[voice-call] msteams TTS fallback used from=${result.fallbackFrom} to=${result.provider} attempts=${attemptedChain}`);
        }
        // Alignment is wall-clock seconds, so it stays valid across the resample below.
        const pcm16k = result.sampleRate === MSTEAMS_TTS_SAMPLE_RATE_HZ
            ? result.audioBuffer
            : resamplePcm(result.audioBuffer, result.sampleRate, MSTEAMS_TTS_SAMPLE_RATE_HZ);
        return { pcm16k, alignment: result.alignment };
    };
    return {
        synthesizePcm16k: async (text) => (await synthesizePcm16kWithTiming(text)).pcm16k,
        synthesizePcm16kWithTiming,
    };
}
/** Layer the voice-call `tts` override on top of the core `messages.tts` config. */
function applyTtsOverride(coreConfig, override) {
    if (!override) {
        return coreConfig;
    }
    const base = coreConfig.messages?.tts;
    // Merge the override onto the host messages.tts. (The standalone plugin's config is already
    // validated by openclaw.plugin.json's configSchema at load, so no extra schema re-parse here.)
    const merged = (base ? deepMergeDefined(base, override) : override);
    return {
        ...coreConfig,
        messages: {
            ...coreConfig.messages,
            tts: merged,
        },
    };
}
