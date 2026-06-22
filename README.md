# @alaamh/msteams-voice

**Self-contained Microsoft Teams voice agent (CVI) for OpenClaw.**

An AI assistant that joins Microsoft Teams calls as a real participant — **realtime speech-to-speech**
*or* a **streaming STT→agent→TTS** pipeline — with continuous vision, "speak-only-when-addressed"
gating, outbound call-backs with voicemail, avatar lip-sync, and meeting recap.

It is **one plugin** depending only on the **published `openclaw` plugin-sdk** + `api.runtime` — no
`@alaamh/voice-call` fork, no vendored `CallManager`/config. Install like any third-party plugin:

```
openclaw plugins install clawhub:@alaamh/msteams-voice
cd extensions/msteams-voice && pnpm install && pnpm build
```

> The original two-package fork approach is kept at `alaamh/openclaw-voice-call` as a fallback.

## Two modes

| | `realtime` | `streaming` |
|---|---|---|
| How it talks | speech-to-speech realtime model (e.g. OpenAI Realtime) | your openclaw-configured **STT → agent/model → TTS** |
| Needs a realtime provider | **yes** (`realtime.provider` + key) | **no** — uses openclaw's transcription + TTS + agent |
| Latency | lowest | higher (per-turn) |
| Continuous vision push | yes | on-demand (see vision note) |
| Use it when | you have a realtime voice model | you want any STT/TTS/model, or lower cost |

**Mode selection:** set `mode` to `"realtime"` or `"streaming"`. If omitted, the runtime auto-selects:
**realtime** when a realtime provider resolves, otherwise **streaming**. Both modes honor the inbound
allowlist, outbound call-backs, recording gate, and `sessionScope` agent memory.

## Configuration

Config lives under `plugins.entries."msteams-voice".config` in your OpenClaw config. `sharedSecret`
**must match** the AzureBot/Teams worker that connects to this plugin's media WebSocket.

### Realtime mode (speech-to-speech)
```jsonc
{
  "plugins": { "entries": { "msteams-voice": { "config": {
    "enabled": true,
    "mode": "realtime",
    "port": 9442,
    "path": "/voice/msteams/stream",
    "sharedSecret": "<same secret as the Teams worker>",
    "requireRecordingStatus": true,
    "inboundPolicy": "allowlist",
    "allowFrom": ["<caller AAD object id or phone number>"],
    "inboundGreeting": "Hello, this is the assistant.",
    "maxConcurrentCalls": 4,
    "maxDurationSeconds": 3600,
    "groupCall": { "requireAddress": true, "wakePhrases": ["assistant"], "followUpWindowMs": 8000 },
    "maxVisionPerMinute": 30,
    "meetingRecap": true,
    "bilingual": true,
    "realtime": {
      "provider": "openai",
      "providers": { "openai": { "apiKey": "<key>", "model": "gpt-realtime" } },
      "instructions": "You are a helpful Teams meeting assistant.",
      "toolPolicy": "safe-read-only",
      "suppressInputDuringPlayback": true,
      "echoSuppressionWindowMs": 250,
      "echoBargeInRms": 0.02
    }
  } } } }
}
```

### Streaming mode (STT → agent → TTS, no realtime model)
```jsonc
{
  "plugins": { "entries": { "msteams-voice": { "config": {
    "enabled": true,
    "mode": "streaming",
    "port": 9442,
    "path": "/voice/msteams/stream",
    "sharedSecret": "<same secret as the Teams worker>",
    "requireRecordingStatus": true,
    "inboundPolicy": "allowlist",
    "allowFrom": ["<caller id>"],
    "inboundGreeting": "Hello, this is the assistant.",
    "maxConcurrentCalls": 4,
    "groupCall": { "requireAddress": true, "wakePhrases": ["assistant"] },
    "meetingRecap": true
  } } } }
}
```
In **streaming** mode the **STT (transcription), TTS, and agent/model come from your openclaw
configuration** — there is no realtime provider or key to set here. The `realtime.*` block is ignored
except the echo-guard knobs (`suppressInputDuringPlayback`, `echoSuppressionWindowMs`, `echoBargeInRms`),
which apply to playback in both modes.

### Outbound call-backs (optional, either mode)
```jsonc
"outbound": {
  "enabled": true,
  "workerBaseUrl": "https://<your-teams-worker>",
  "tenantId": "<aad-tenant-id>",
  "answerTimeoutMs": 120000,
  "defaultMode": "notify"        // "notify" delivers a message then ends; "conversation" opens a turn
}
```
Then call `runtime.placeCall(userObjectId, { message, mode })` (no-answer/declined → voicemail/no-answer).

### Key reference
| Key | Applies | Meaning |
|---|---|---|
| `enabled` | both | master on/off |
| `mode` | both | `"realtime"` \| `"streaming"` (auto if omitted) |
| `port` / `bindAddress` / `path` | both | media WebSocket server the Teams worker connects to |
| `sharedSecret` | both | HMAC secret — **must match the worker** (secret input) |
| `requireRecordingStatus` | both | only engage once Teams reports recording active |
| `inboundPolicy` | both | `disabled` \| `allowlist` \| `pairing` \| `open` — **enforced** on inbound |
| `allowFrom` | both | allowlisted caller ids (Teams aadId or phone digits) |
| `inboundGreeting` | both | opening line |
| `sessionScope` | both | `per-phone` \| `per-call` \| `per-thread` agent memory scope |
| `maxConcurrentCalls` / `maxDurationSeconds` / `staleCallReaperSeconds` | both | capacity + reaper |
| `groupCall.{requireAddress,wakePhrases,followUpWindowMs}` | both | speak-only-when-addressed gating |
| `maxVisionPerMinute` | both | vision spend cap |
| `meetingRecap` / `bilingual` | both | post-call minutes / Arabic-English |
| `realtime.{provider,providers,instructions,toolPolicy,…}` | realtime | the realtime voice provider + behavior |
| `outbound.{…}` | both | outbound call-backs / voicemail |

## Architecture (why it's small)
The hard parts come from OpenClaw (see [DESIGN.md](DESIGN.md)):
- **Realtime audio bridge** → `openclaw/plugin-sdk/realtime-voice` (`createRealtimeVoiceBridgeSession`,
  `consultRealtimeVoiceAgent`, `resolveConfiguredRealtimeVoiceProvider`).
- **Agent / TTS / STT / media / state / config / logging** → `api.runtime`.
- **We own** only `src/call-lifecycle.ts` (~500 LOC) + thin adapters + the Teams CVI logic.
- Entry registers a host-managed service (`api.registerService({ id, start, stop })`).
