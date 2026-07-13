# Microsoft Teams Bridge for OpenClaw

[![CI](https://github.com/komaa-com/openclaw-msteams-bridge/actions/workflows/dist-sync.yml/badge.svg)](https://github.com/komaa-com/openclaw-msteams-bridge/actions/workflows/dist-sync.yml)
[![npm version](https://img.shields.io/npm/v/@komaa/openclaw-msteams-bridge.svg)](https://www.npmjs.com/package/@komaa/openclaw-msteams-bridge)
[![downloads](https://img.shields.io/npm/dm/@komaa/openclaw-msteams-bridge.svg)](https://www.npmjs.com/package/@komaa/openclaw-msteams-bridge)
[![docs](https://img.shields.io/badge/docs-komaa--com.github.io-2563eb.svg)](https://komaa-com.github.io/openclaw-msteams-bridge/)
[![MIT License](https://img.shields.io/badge/license-MIT-blue.svg)](./LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-strict-3178C6.svg)](https://www.typescriptlang.org/)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](./CONTRIBUTING.md)

**`@komaa/openclaw-msteams-bridge`** is a Microsoft Teams voice and video agent (CVI) for
[OpenClaw](https://openclaw.ai). It turns an
ordinary Teams call into a true two-way video conversation: the agent sees what you show it, talks
back in real time, and appears in the call as an animated, lip-synced avatar.

It is one plugin, depending only on the published `openclaw` plugin-sdk + `api.runtime`. No fork, no
vendored runtime, no trusted-plugin privileges required.

## Three pillars

| Pillar | The agent | What it does in the call |
|:--|:--|:--|
| **Perception** | sees you | Reads camera and screen-share (VBSS) frames: `look_at_screen` on demand, a keyframe per turn in streaming mode, an ambient view in realtime mode. Frames are attributed per participant and vision spend is capped per call. |
| **Dialogue** | talks with you | Holds a spoken conversation - realtime speech-to-speech or a streaming STT to agent to TTS pipeline. Barge-in, verbal interrupts, a "speak only when addressed" gate, DTMF entry, English and Arabic, and greeting attendees by name. |
| **Rendering** | is seen by you | Appears as a lip-synced avatar tile. Emits expression cues, viseme lip-sync, and `show_to_caller` overlays, which the hosted StandIn bridge draws into the video the caller sees. |

Details of each capability follow below.

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

Do these in order. Steps 1 and 2 are prerequisites (see [Getting started](#getting-started) above);
this plugin is step 3.

1. **OpenClaw is installed and running** (host `>= 2026.6.10`).
2. **Microsoft Teams is added as a channel** (bot app + credentials).
3. **Install this plugin**, then restart the gateway so it loads:

   ```bash
   openclaw plugins install npm:@komaa/openclaw-msteams-bridge
   openclaw gateway restart
   ```

4. **Configure it** under `plugins.entries."msteams-voice".config` - at minimum set `sharedSecret`
   (to match StandIn), `inboundPolicy`, and your provider key. See [Configuration](#configuration)
   and [Security](#security) below. The server refuses to start until `sharedSecret` is set.
5. **Connect StandIn** to the plugin's WebSocket (start in the
   [sandbox](https://standin.komaa.com/sandbox)) and place a test call.

Prefer a guided setup? The one-line installer detects your OpenClaw install and walks you through
steps 3 and 4 (mode, shared secret, provider key), applying the secure defaults for you:

```bash
curl -fsSL https://standin.komaa.com/install.sh | bash
```

Also on [ClawHub](https://clawhub.ai): `openclaw plugins install clawhub:@komaa/openclaw-msteams-bridge`
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

## Security

The plugin ships with **secure defaults**, and the recommendation is simple: keep them. Each option
below states its safe default and why it is safe. You only relax a default when your deployment
genuinely needs it, and only in the narrow way described.

| Option | Safe default | Why it is safe | When to change it |
|:--|:--|:--|:--|
| `sharedSecret` | none (fails closed) | The media WebSocket authenticates every connection with a replay-proof HMAC handshake keyed on this secret. With no secret the server refuses to start, so a misconfig can never expose an unauthenticated port. A non-string value coerces to empty and also fails closed. | Always set it, to a strong random value that matches your StandIn dashboard. Prefer an OpenClaw secret reference over a literal in config. |
| `inboundPolicy` | unset = deny all | Inbound calls are rejected until you name a policy, so the agent never answers an unknown caller by default. | Set `allowlist` and list trusted callers in `allowFrom` (by AAD object id or phone number). Reserve `open` for throwaway sandbox testing only. |
| `requireRecordingStatus` | `true` | Media is held until Teams reports recording is active, so the agent never sees or hears the call before participants have the recording indicator. This keeps you on the right side of Teams' notice expectations. | Leave it on. Only disable for a controlled test where no real participants are present. |
| `bindAddress` | `127.0.0.1` (loopback) | The WebSocket listens on localhost only, so it is unreachable from other hosts by default. | Widen to `0.0.0.0` only when the StandIn bridge runs on a different host, and only on a trusted or VPN-only interface behind your firewall. The HMAC handshake still guards it, but do not expose the port to the open internet. |
| `realtime.toolPolicy` | `none` | The voice model cannot invoke any agent tools, so a caller cannot drive tools by voice unless you opt in. | Use `safe-read-only` to allow read-only tools. Reserve `owner` (full tool access) for calls you have restricted to trusted owners via `inboundPolicy`. |
| Installer | applies the above | The one-line installer configures these secure defaults for you rather than leaving them blank. | If your policy forbids piping a script to a shell, download and read `install.sh` first, then run it, or follow the manual [Install](#install) steps. |

In short: set a strong `sharedSecret`, keep `inboundPolicy` restrictive with an explicit `allowFrom`,
leave `requireRecordingStatus` on, keep `bindAddress` as tight as your topology allows, and only
raise `toolPolicy` for callers you trust.

## Key reference

Full reference in the [Configuration Reference](https://komaa-com.github.io/openclaw-msteams-bridge/configuration-reference/). Common keys:

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

- Plugin docs: [komaa-com.github.io/openclaw-msteams-bridge](https://komaa-com.github.io/openclaw-msteams-bridge/)
- StandIn (hosted service) docs: [docs.komaa.com](https://docs.komaa.com/)
- Source: [github.com/komaa-com/openclaw-msteams-bridge](https://github.com/komaa-com/openclaw-msteams-bridge)
- npm: [@komaa/openclaw-msteams-bridge](https://www.npmjs.com/package/@komaa/openclaw-msteams-bridge)

---

<p align="center"><sub>Built by <a href="https://komaa.com">Komaa.com</a> - MIT licensed</sub></p>
