# @komaa/msteams-voice

[![npm](https://img.shields.io/npm/v/@komaa/msteams-voice.svg)](https://www.npmjs.com/package/@komaa/msteams-voice)
[![docs](https://img.shields.io/badge/docs-komaa.com-2563eb.svg)](https://docs.komaa.com/)
[![license](https://img.shields.io/badge/license-MIT-green.svg)](./LICENSE)

A self-contained Microsoft Teams voice agent (CVI) for [OpenClaw](https://openclaw.ai). An AI
assistant that joins Teams calls as a real participant, in **realtime speech-to-speech** or a
**streaming STT to agent to TTS** pipeline, with continuous vision, speak-only-when-addressed
gating, outbound call-backs with voicemail, avatar lip-sync, and meeting recap.

> Full documentation: **[docs.komaa.com](https://docs.komaa.com/)** (setup, StandIn bridge,
> configuration reference, troubleshooting). This README is the quick start.

It is one plugin, depending only on the published `openclaw` plugin-sdk + `api.runtime`. No fork, no
vendored runtime, no trusted-plugin privileges required.

## Features

- **Two dialogue modes** - realtime speech-to-speech (OpenAI / Azure) or streaming STT to agent to TTS
- **Continuous vision** - the agent can look at screen-share / camera frames, budget-capped
- **Group-call gating** - answers only when addressed by a wake phrase
- **Outbound call-backs + voicemail** - place a call, deliver a message, or open a conversation
- **Meeting recap** - a `.docx` of minutes with per-speaker attribution
- **Bilingual** (Arabic / English), **DTMF**, **barge-in / echo guard**
- **Secure transport** - HMAC-signed media bridge + caller allowlist

## Requirements

- An **OpenClaw** install (host `>= 2026.6.9`).
- A **StandIn** subscription ([standin.komaa.com](https://standin.komaa.com), free tier), the hosted
  media bridge that joins the Teams call and connects to this plugin's WebSocket.
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

Full reference at [docs.komaa.com](https://docs.komaa.com/openclaw/configuration). Common keys:

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
| `inboundPolicy` | `disabled`, `allowlist`, `pairing`, `open` |
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

- Docs: [docs.komaa.com](https://docs.komaa.com/)
- Source: [github.com/komaa-com/openclaw-msteams-voice](https://github.com/komaa-com/openclaw-msteams-voice)
- npm: [@komaa/msteams-voice](https://www.npmjs.com/package/@komaa/msteams-voice)

---

<p align="center"><sub>Built by <a href="https://komaa.com">Komaa.com</a> - MIT licensed</sub></p>
