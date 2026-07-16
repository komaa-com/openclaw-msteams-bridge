import { resamplePcm } from "openclaw/plugin-sdk/realtime-voice";
import { deepMergeDefined } from "./deep-merge.js";
import { MSTEAMS_PCM_SAMPLE_RATE_HZ } from "./msteams-media-stream.js";
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
            logger?.warn?.(`[msteams-voice] TTS fallback used from=${result.fallbackFrom} to=${result.provider} attempts=${attemptedChain}`);
        }
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
function applyTtsOverride(coreConfig, override) {
    if (!override) {
        return coreConfig;
    }
    const base = coreConfig.messages?.tts;
    const merged = (base ? deepMergeDefined(base, override) : override);
    return {
        ...coreConfig,
        messages: {
            ...coreConfig.messages,
            tts: merged,
        },
    };
}
