# @komaa/msteams-voice

[![npm](https://img.shields.io/npm/v/@komaa/msteams-voice.svg)](https://www.npmjs.com/package/@komaa/msteams-voice)
[![docs](https://img.shields.io/badge/docs-komaa.com-2563eb.svg)](https://docs.komaa.com/)
[![license](https://img.shields.io/badge/license-MIT-green.svg)](./LICENSE)

**A self-contained Microsoft Teams voice agent (CVI) for [OpenClaw](https://openclaw.ai).** An AI
assistant that joins Teams calls as a real participant — **realtime speech-to-speech** *or* a
**streaming STT → agent → TTS** pipeline — with continuous vision, "speak-only-when-addressed" gating,
outbound call-backs with voicemail, avatar lip-sync, and meeting recap.

> 📖 **Full documentation: [docs.komaa.com](https://docs.komaa.com/)** — setup walkthroughs, the Teams
> worker, configuration reference, and troubleshooting. This README is the quick start.

It's **one plugin** depending only on the published `openclaw` plugin-sdk + `api.runtime` — no fork, no
vendored runtime.

## Features

- 🎙️ **Realtime speech-to-speech** (e.g. OpenAI Realtime) **or** streaming **STT → agent → TTS** (any provider)
- 👁️ **Continuous vision** — the agent can "look at" screen-share / camera frames, with a per-minute budget
- 🙋 **Group-call gating** — only answers when addressed by a wake phrase (silent otherwise)
- 📞 **Outbound call-backs + voicemail** — place a call, deliver a message, or open a conversation
- 📝 **Meeting recap** — a `.docx` of minutes with per-speaker attribution
- 🌐 **Bilingual** (Arabic / English) · ⌨️ **DTMF** · 🔇 **barge-in / echo guard** · 🔐 **HMAC-signed media bridge + caller allowlist**

## Requirements

- An **OpenClaw** install (host ≥ `2026.6.9`).
- A **Microsoft Teams worker** (Azure Bot) that bridges the call audio to this plugin's media WebSocket
  — see [docs.komaa.com](https://docs.komaa.com/).
- For **realtime** mode: a realtime voice provider + key. For **streaming** mode: your
  openclaw-configured STT/TTS/agent (no realtime key needed).

## Install

```bash
openclaw plugins install clawhub:@komaa/msteams-voice
cd extensions/msteams-voice && pnpm install && pnpm build
```

## Two modes

| | `realtime` | `streaming` |
|---|---|---|
| How it talks | speech-to-speech realtime model | your openclaw **STT → agent/model → TTS** |
| Needs a realtime provider | **yes** (`realtime.provider` + key) | **no** |
| Latency | lowest | higher (per-turn) |
| Vision | continuous push (live) | attached to each agent turn |
| Use it when | you have a realtime voice model | any STT/TTS/model, or lower cost |

**Mode selection:** set `mode` to `"realtime"` or `"streaming"`. If omitted, the runtime auto-selects
**realtime** when a realtime provider resolves, else **streaming**. Both modes honor the inbound
allowlist, outbound call-backs, recording gate, and `sessionScope` agent memory.

## Configuration

Config lives under `plugins.entries."msteams-voice".config` in your OpenClaw config. `sharedSecret`
**must match** the Teams worker that connects to this plugin's media WebSocket.

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
    "maxVisionPerMinute": 30,
    "meetingRecap": true,
    "stt": {
      "provider": "<your-stt-provider>",
      "providers": { "<your-stt-provider>": { "apiKey": "<key>" } }
    }
  } } } }
}
```

In **streaming** mode, **TTS and the agent/model come from your openclaw configuration**. STT uses a
live transcription session — selected by `stt.provider`/`stt.providers` if set, else your
openclaw-configured transcription provider; if none resolves it falls back to VAD-segmented file
transcription. No realtime provider/key needed. The `realtime.*` block is ignored except the
echo-guard knobs (`suppressInputDuringPlayback`, `echoSuppressionWindowMs`, `echoBargeInRms`), which
apply in both modes. Group-call gating, DTMF, and vision work in streaming mode too.

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

`placeCall(userObjectId, { message, mode })` is implemented on the runtime (no-answer/declined →
voicemail/no-answer); the `outbound` block enables it. Triggering it currently requires a host call
into the runtime — a built-in agent tool / endpoint is a small follow-up.

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
| `sessionScope` | both | `per-phone` \| `per-call` \| `per-thread` agent-memory scope |
| `maxConcurrentCalls` / `maxDurationSeconds` / `staleCallReaperSeconds` | both | capacity + reaper |
| `groupCall.{requireAddress,wakePhrases,followUpWindowMs}` | both | speak-only-when-addressed gating |
| `maxVisionPerMinute` | both | vision spend cap |
| `meetingRecap` / `bilingual` | both | post-call minutes / Arabic-English |
| `realtime.{provider,providers,instructions,toolPolicy,…}` | realtime (echo knobs: both) | realtime voice provider + behavior; provider key is a secret input |
| `stt.{provider,providers}` | streaming | live transcription provider (else openclaw STT / file fallback); provider key is a secret input |
| `outbound.{enabled,workerBaseUrl,tenantId,answerTimeoutMs,defaultMode}` | both | outbound call-backs / voicemail |

## Architecture

The hard parts come from OpenClaw — this plugin is intentionally small:

- **Realtime audio bridge** → `openclaw/plugin-sdk/realtime-voice` (`createRealtimeVoiceBridgeSession`, `consultRealtimeVoiceAgent`, `resolveConfiguredRealtimeVoiceProvider`).
- **Agent / TTS / STT / media / state / config / logging** → `api.runtime`.
- **Owned code:** `src/call-lifecycle.ts` (~500 LOC) + thin adapters + the Teams CVI logic.
- The entry registers a host-managed service (`api.registerService({ id, start, stop })`).

## Links

- 📖 **Docs:** [docs.komaa.com](https://docs.komaa.com/)
- 💻 **Source:** [github.com/komaa-com/openclaw-msteams-voice](https://github.com/komaa-com/openclaw-msteams-voice)
- 📦 **npm:** [@komaa/msteams-voice](https://www.npmjs.com/package/@komaa/msteams-voice)

---

<p align="center"><sub>Built by <a href="https://komaa.com">Komaa.com</a> · MIT licensed</sub></p>
