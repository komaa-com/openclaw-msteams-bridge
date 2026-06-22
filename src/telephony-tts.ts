// Type of the host TTS runtime the Teams TTS adapter consumes. This is exactly the shape of
// `api.runtime.tts` (its `textToSpeechTelephony` method) — so the plugin entry passes `api.runtime.tts`
// straight in. (Type only; no synthesis logic is vendored — synthesis lives in the host runtime.)

import type { TtsDirectiveOverrides } from "openclaw/plugin-sdk/speech";
import type { CoreConfig } from "./core-bridge.js";

export type TelephonyTtsRuntime = {
  textToSpeechTelephony: (params: {
    text: string;
    cfg: CoreConfig;
    prefsPath?: string;
    overrides?: TtsDirectiveOverrides;
  }) => Promise<{
    success: boolean;
    audioBuffer?: Buffer;
    sampleRate?: number;
    provider?: string;
    fallbackFrom?: string;
    attemptedProviders?: string[];
    error?: string;
    /** Per-character timing (e.g. ElevenLabs with-timestamps); wall-clock seconds. */
    alignment?: { characters: string[]; startTimesSeconds: number[] };
  }>;
};
