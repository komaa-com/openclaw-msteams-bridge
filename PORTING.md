# Porting guide — bring the CVI code in

Source of the CVI files: **`alaamh/openclaw-voice-call`** repo, package
`packages/voice-call-msteams/src/` (already import-rewritten + tested there). That repo stays as the
fallback; here we re-home the same files onto the **public SDK + api.runtime** (no `@alaamh/voice-call`).

## 1. Copy these CVI files into `src/` (they're our IP, ship as-is)
`msteams-media-stream.ts`, `msteams-realtime.ts`, `msteams-realtime-tools.ts`, `msteams-tts.ts`,
`msteams-tts-playback.ts`, `msteams-vision-store.ts`, `msteams-video-frame.ts`, `group-call-gate.ts`,
`verbal-interrupt.ts`, `viseme-estimate.ts`, `expression.ts`, `vision-budget.ts`,
`meeting-minutes-docx.ts`, `realtime-fast-context.ts`, `realtime-voice-compat.ts`
(+ their `*.test.ts`).

## 2. Rewrite imports (from `@alaamh/voice-call/sdk` → SDK / api.runtime / local)
| Was (in the fork) | Now |
|---|---|
| `createRealtimeVoiceBridgeSession`, `consultRealtimeVoiceAgent`, `resolveRealtimeVoiceAgentConsultTools`, RealtimeVoice* types, pcm helpers | `openclaw/plugin-sdk/realtime-voice` (already the source in the fork too) |
| `generateVoiceResponse` (response-generator) | `api.runtime.agent.runEmbeddedAgent` (wrap in a tiny local helper) |
| `TelephonyTtsRuntime` / TTS synth | `api.runtime.tts.textToSpeechTelephony` |
| `buildRealtimeVoiceInstructions` | port the small helper (or inline) — not in api.runtime |
| `resolveVoiceResponseModel` | `api.runtime.agent.defaults` / model resolution helper (small) |
| `CallManager`, `manager/*`, call config/lifecycle | **`src/call-lifecycle.ts`** (write — see DESIGN) |
| `VoiceCallConfig` / `resolveVoiceCallEffectiveConfig` | local `src/config.ts` resolving `api.pluginConfig` against the manifest schema (small) |
| `isInboundCallAllowed` (allowlist) | port the small `allowlist.ts` (pure) |
| `chunkAudio` / `readArgText` / `deepMergeDefined` | inline the few tiny utils |

## 3. Wire the entry (`src/index.ts`)
- `definePluginEntry({ id: "msteams-voice", registerFull })`.
- In `registerFull(api)`: read `api.pluginConfig`; start the Teams media WS server
  (`msteams-media-stream`); on each call, create a `CallLifecycle` entry and a realtime bridge via
  `createRealtimeVoiceBridgeSession` (audioSink → Teams WS); route the consult tool to
  `consultRealtimeVoiceAgent` / `api.runtime.agent`.

## 4. Verify
- `pnpm build` clean (deps: only published `openclaw` + ws/zod/docx/jszip).
- `pnpm test` — port the unit tests (vision-push tests stay `it.skip` only if the bridge lacks
  `sendImage`; here `sendImage` works because we target our own openclaw build / the SDK that has it).
- Smoke: `provider`/runtime resolves and a Teams `session.start` drives a realtime turn.
