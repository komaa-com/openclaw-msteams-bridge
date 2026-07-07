// Type of the host TTS runtime the Teams TTS adapter consumes. This is exactly the shape of
// `api.runtime.tts` (its `textToSpeechTelephony` method) — so the plugin entry passes `api.runtime.tts`
// straight in. (Type only; no synthesis logic is vendored — synthesis lives in the host runtime.)
export {};
