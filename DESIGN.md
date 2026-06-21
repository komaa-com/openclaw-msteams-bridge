# Design — self-contained Teams voice plugin on the public SDK

Goal: ship the Teams CVI voice agent as **one plugin** depending only on published `openclaw` —
no `@alaamh/voice-call` fork, no vendored `CallManager`. Validated by two spikes (talk-voice fit +
`api.runtime` surface).

## What comes for free (verified, file:line in openclaw `next`)

| Need | Source | Evidence |
|---|---|---|
| Realtime audio bridge (audio in/out, barge-in, vision push) | `openclaw/plugin-sdk/realtime-voice` → `createRealtimeVoiceBridgeSession`, `RealtimeVoiceAudioSink` | realtime-voice.ts:168-171 |
| Agent run behind the voice consult | `consultRealtimeVoiceAgent` (same SDK) | realtime-voice.ts:123 |
| Agent invocation (general) | `api.runtime.agent.runEmbeddedAgent` | runtime/types-core.ts:237 |
| TTS (incl. telephony) | `api.runtime.tts.textToSpeechTelephony` | types-core.ts:314 |
| STT (file) + streaming transcription provider | `api.runtime.mediaUnderstanding.transcribeAudioFile` + `RealtimeTranscriptionProviderPlugin` | types-core.ts:323 |
| Media utils | `api.runtime.media.*` | types-core.ts:303-310 |
| Config (own manifest config) | `api.pluginConfig` + `api.runtime.config.*` | types-core.ts:182-198 |
| **State / persistence** (call records) | `api.runtime.state.openSyncKeyedStore` | types-core.ts:376 |
| Logging | `api.runtime.logging.getChildLogger` | types-core.ts:369 |

> Note: `voice-call`'s own `telephony-tts.ts` already just calls `runtime.textToSpeechTelephony` — so
> TTS is a genuine passthrough, not something to reimplement.

## What we write ourselves — the ONLY substantial glue

A Teams **call-lifecycle** coordinator (`src/call-lifecycle.ts`, ~400–600 LOC) — the subset of
`voice-call`'s `CallManager` that a Teams-only realtime plugin actually needs:

- Call **state machine** (initiated → ringing → answered → active → terminal)
- **Active-call registry** (`callId` ↔ `providerCallId` map) + event **dedupe**
- **Record persistence** via `api.runtime.state.openSyncKeyedStore` (reuse the Zod `CallRecord` shape)
- **Lifecycle**: `initiate / answer / end / getStatus / reapStale`
- Per-call **timers** (max duration, idle) + **max-concurrency**

We DON'T need (so we don't write): multi-carrier (Twilio/Telnyx/Plivo) webhook normalization,
the webhook media plane, DTMF/TTS playback routing, or crash-restore — none apply to Teams realtime.

## Cost comparison
- Fork the whole voice-call extension: ~21k LOC carried.
- Two-package fork + vendor: ~1.2k–4.6k LOC vendored.
- **This plugin: your existing CVI code + ~500 LOC lifecycle.** Fully independent, single install.

## Open verification before "done"
- Confirm the exact `definePluginEntry` shape for a non-channel runtime plugin (entry import path).
- Confirm `runEmbeddedAgent` / `consultRealtimeVoiceAgent` cover the consult behaviors the CVI uses
  (streaming, model selection, media paths).
