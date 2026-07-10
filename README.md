# Microsoft Teams Bridge for OpenClaw

[![CI](https://github.com/komaa-com/openclaw-msteams-voice/actions/workflows/dist-sync.yml/badge.svg)](https://github.com/komaa-com/openclaw-msteams-voice/actions/workflows/dist-sync.yml)
[![npm version](https://img.shields.io/npm/v/@komaa/msteams-voice.svg)](https://www.npmjs.com/package/@komaa/msteams-voice)
[![downloads](https://img.shields.io/npm/dm/@komaa/msteams-voice.svg)](https://www.npmjs.com/package/@komaa/msteams-voice)
[![docs](https://img.shields.io/badge/docs-komaa--com.github.io-2563eb.svg)](https://komaa-com.github.io/openclaw-msteams-voice/)
[![MIT License](https://img.shields.io/badge/license-MIT-blue.svg)](./LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-strict-3178C6.svg)](https://www.typescriptlang.org/)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](./CONTRIBUTING.md)

**`@komaa/msteams-voice`** is a Microsoft Teams voice and video agent (CVI) for
[OpenClaw](https://openclaw.ai). It turns an
ordinary Teams call into a true two-way video conversation: the agent sees what you show it, talks
back in real time, and appears in the call as an animated, lip-synced avatar.

It is one plugin, depending only on the published `openclaw` plugin-sdk + `api.runtime`. No fork, no
vendored runtime, no trusted-plugin privileges required.

## Three pillars

| Pillar | The agent... | In the call |
|---|---|---|
| **Perception** | sees you | Reads inbound camera and screen-share (VBSS) frames. It can `look_at_screen` on demand, auto-attaches a keyframe each turn in streaming mode, and keeps a continuous ambient view in realtime mode. In meetings, frames are attributed to the participant who sent them, and vision spend is capped per call. |
| **Dialogue** | talks with you | Holds a natural spoken conversation, either realtime speech-to-speech or a streaming STT to agent to TTS pipeline. It supports barge-in and deterministic verbal interrupts, a "speak only when addressed" gate for group calls, DTMF/IVR entry, bilingual English and Arabic, and greeting attendees by name from the roster. |
| **Rendering** | is seen by you | Appears as a lip-synced avatar tile. The plugin emits expression cues (happy, sad, surprised), viseme lip-sync driven by speech marks, and `show_to_caller` image overlays; the hosted StandIn bridge draws them into the video the caller sees. |

## Capabilities

- **Two dialogue modes** - realtime speech-to-speech (OpenAI or Azure OpenAI), or a streaming
  STT to agent to TTS pipeline that works with any provider.
- **Vision** - reads camera and screen-share frames, keeps a continuous ambient view, retains a
  retroactive keyframe history, and stays inside a per-call budget cap.
- **Group and meeting etiquette** - stays silent until addressed by a wake phrase, then answers
  through a short follow-up window; 1:1 calls always answer; every frame and utterance is attributed
  to the right speaker.
- **Outbound call-backs** - places a call, speaks the result, and hangs up, with a voicemail
  fallback when no one answers.
- **Meeting recap and minutes** - end-of-call summary of key points, decisions, and action items,
  plus an on-demand `.docx` of minutes with per-person attribution.
- **Avatar driver cues** - expression changes, viseme lip-sync, and picture-in-picture image
  sharing, rendered by the StandIn bridge.
- **Chat governance** - an "Ask about this" message action, voice-message transcription, an
  audit-log mirror, and outbound DLP redaction.
- **Secure transport** - a replay-proof HMAC handshake, a caller allowlist that is closed by
  default, and a recording-status gate that holds media until recording is active.

## Getting started

This plugin adds **voice and video (CVI)** on top of OpenClaw's Microsoft Teams **chat** channel, so
set those up first:

1. **Install OpenClaw** using the official docs at
   [docs.openclaw.ai](https://docs.openclaw.ai).
2. **Add Microsoft Teams as a channel** (bot app + credentials) following the
   [OpenClaw Teams channel docs](https://docs.openclaw.ai/channels/msteams).
3. **Use the StandIn sandbox** ([standin.komaa.com/sandbox](https://standin.komaa.com/sandbox), free,
   no Teams bot needed), the hosted media bridge that joins the call and connects to this plugin. Add
   your own Teams bot later at [standin.komaa.com](https://standin.komaa.com) for inbound calls.
4. **Add this plugin.** The one-line installer detects your OpenClaw install and walks you through the
   config (mode, shared secret, provider key), then prints the next steps:

   ```bash
   curl -fsSL https://standin.komaa.com/install.sh | bash
   ```

   Prefer to do it by hand? See [Install](#install) and [Configuration](#configuration) below.

## Requirements

- An **OpenClaw** install (host `>= 2026.6.10`).
- **StandIn** to bridge the call: start free in the [sandbox](https://standin.komaa.com/sandbox) (no
  Teams bot), or add your own bot at [standin.komaa.com](https://standin.komaa.com) for inbound calls.
  It is the hosted media bridge that joins the Teams call and connects to this plugin's WebSocket.
- For **realtime** mode: a realtime voice provider + key (OpenAI or Azure OpenAI). For **streaming**
  mode: your OpenClaw-configured STT / TTS / agent (no realtime key needed).

## Install

One-line installer (detects OpenClaw, installs and configures the plugin):

```bash
curl -fsSL https://standin.komaa.com/install.sh | bash
```

Or install the plugin manually, then restart the gateway:

```bash
openclaw plugins install npm:@komaa/msteams-voice
openclaw gateway restart
```

Also on [ClawHub](https://clawhub.ai): `openclaw plugins install clawhub:@komaa/msteams-voice`
(OpenClaw falls back to npm automatically if the ClawHub fetch fails). The package ships prebuilt
(v0.1.10+): no build step either way.

## Two modes

| | `realtime` | `streaming` |
|---|---|---|
| How it talks | speech-to-speech model | your OpenClaw STT to agent to TTS |
| Needs a realtime key | yes | no |
| Latency | lowest | higher (per turn) |
| Vision | continuous push | attached per turn |

Set `mode` to `"realtime"` or `"streaming"`. If omitted, the runtime auto-selects realtime when a
realtime provider resolves, else streaming.

## Configuration

Config lives under `plugins.entries."msteams-voice".config`. `sharedSecret` must match the value set
in your StandIn dashboard. Set `bindAddress` to `0.0.0.0` so the hosted bridge can reach it.

### Realtime (OpenAI)

```jsonc
{
  "plugins": {
    "entries": {
      "msteams-voice": {
        "config": {
          "enabled": true,
          "mode": "realtime",
          "bindAddress": "0.0.0.0",
          "port": 9442,
          "path": "/voice/msteams/stream",
          "sharedSecret": "<same secret as in StandIn>",
          "requireRecordingStatus": true,
          "inboundPolicy": "allowlist",
          "allowFrom": ["<caller AAD object id>"],
          "realtime": {
            "provider": "openai",
            "providers": {
              "openai": { "apiKey": "<key>", "model": "gpt-realtime" }
            }
          }
        }
      }
    }
  }
}
```

### Realtime (Azure OpenAI)

Azure is the `openai` provider plus `azureEndpoint` and `azureDeployment`:

```jsonc
"realtime": {
  "provider": "openai",
  "providers": {
    "openai": {
      "apiKey": "<azure-key>",
      "azureEndpoint": "https://<resource>.cognitiveservices.azure.com",
      "azureDeployment": "gpt-realtime"
    }
  }
}
```

### Streaming (no realtime key)

```jsonc
"mode": "streaming",
"stt": {
  "provider": "<your-stt-provider>",
  "providers": { "<your-stt-provider>": { "apiKey": "<key>" } }
}
```

In streaming mode the TTS and agent come from your OpenClaw configuration. STT uses `stt.provider`
if set, else your configured transcription provider, else a VAD-segmented file fallback. Group-call
gating, DTMF, and vision all work in streaming mode too.

### Outbound call-backs (optional)

```jsonc
"outbound": {
  "enabled": true,
  "workerBaseUrl": "https://<your-standin-endpoint>",
  "tenantId": "<aad-tenant-id>",
  "answerTimeoutMs": 120000,
  "defaultMode": "notify"
}
```

## Key reference

Full reference in the [Configuration Reference](https://komaa-com.github.io/openclaw-msteams-voice/configuration-reference/). Common keys:

**Core**

| Key | Description |
|---|---|
| `enabled` | master on/off |
| `mode` | `realtime` or `streaming` (auto if omitted) |
| `port` | WebSocket port (default `9442`) |
| `bindAddress` | bind address; use `0.0.0.0` for the hosted bridge |
| `path` | WebSocket path (default `/voice/msteams/stream`) |
| `sharedSecret` | HMAC secret; must match StandIn |
| `requireRecordingStatus` | engage only once recording is active |
| `inboundPolicy` | `disabled`, `allowlist`, `pairing`, `open`. `pairing` currently behaves exactly like `allowlist` (the plugin issues no pairing codes or approvals for calls; callers must be in `allowFrom`) |
| `allowFrom` | allowlisted caller ids |
| `inboundGreeting` | opening line |
| `sessionScope` | `per-phone`, `per-call`, `per-thread` |
| `maxConcurrentCalls` | concurrent-call cap |
| `maxDurationSeconds` | max answered-call duration |
| `maxVisionPerMinute` | vision spend cap |
| `meetingRecap` | post end-of-call minutes |
| `bilingual` | Arabic / English |

**Group call**

| Key | Description |
|---|---|
| `groupCall.requireAddress` | answer only when addressed |
| `groupCall.wakePhrases` | wake words |
| `groupCall.followUpWindowMs` | follow-up window (ms) |

**Realtime**

| Key | Description |
|---|---|
| `realtime.provider` | `openai` |
| `realtime.providers.openai.apiKey` | provider key (secret) |
| `realtime.providers.openai.model` | e.g. `gpt-realtime` |
| `realtime.providers.openai.azureEndpoint` | Azure OpenAI endpoint |
| `realtime.providers.openai.azureDeployment` | Azure deployment name |
| `realtime.instructions` | system instructions |
| `realtime.toolPolicy` | `safe-read-only`, `owner`, `none` |
| `realtime.suppressInputDuringPlayback` | echo guard (both modes) |
| `realtime.echoSuppressionWindowMs` | echo guard window (ms) |
| `realtime.echoBargeInRms` | barge-in RMS threshold |

**Streaming**

| Key | Description |
|---|---|
| `stt.provider` | transcription provider |
| `stt.providers.<id>.apiKey` | STT key (secret) |

**Outbound**

| Key | Description |
|---|---|
| `outbound.enabled` | enable call-backs |
| `outbound.workerBaseUrl` | StandIn outbound API URL |
| `outbound.tenantId` | your AAD tenant id |
| `outbound.answerTimeoutMs` | no-answer timeout |
| `outbound.defaultMode` | `notify` or `conversation` |

## Links

- Plugin docs: [komaa-com.github.io/openclaw-msteams-voice](https://komaa-com.github.io/openclaw-msteams-voice/)
- StandIn (hosted service) docs: [docs.komaa.com](https://docs.komaa.com/)
- Source: [github.com/komaa-com/openclaw-msteams-voice](https://github.com/komaa-com/openclaw-msteams-voice)
- npm: [@komaa/msteams-voice](https://www.npmjs.com/package/@komaa/msteams-voice)

---

<p align="center"><sub>Built by <a href="https://komaa.com">Komaa.com</a> - MIT licensed</sub></p>
