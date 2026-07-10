---
title: "Configuration Reference"
description: "Every configuration option of the plugin, with types, defaults, and meaning."
---

All options live under `plugins.entries."msteams-voice".config`. The schema is
`additionalProperties: false`, so an unknown key is rejected. Defaults below come from the config
resolver; secret-valued keys accept either a literal string or an OpenClaw secret reference.

:::caution[Two settings gate every call]
Inbound calls are **denied unless `inboundPolicy` is set** (use `allowlist` + `allowFrom`, or
`open` for sandbox testing), and the hosted bridge cannot reach the plugin until `bindAddress`
is opened up from its local-only default. If your first call never connects or is rejected,
check these two first. See [Troubleshooting](/openclaw-msteams-bridge/troubleshooting/).
:::

## Core

| Key | Type | Default | Meaning |
|---|---|---|---|
| `enabled` | bool | `true` | Master on/off. |
| `port` | int (1-65535) | `9442` | WebSocket server port. |
| `bindAddress` | string | `127.0.0.1` | Bind address. Use `0.0.0.0` so the hosted StandIn bridge can connect. |
| `path` | string | `/voice/msteams/stream` | WebSocket route; StandIn connects to `{path}/{callId}`. |
| `sharedSecret` | string \| secret-ref | - | HMAC secret; **must match StandIn**. Fails closed - a non-string coerces to empty and rejects all handshakes. |
| `requireRecordingStatus` | bool | `true` | Hold media processing until Teams reports recording is active. |
| `inboundPolicy` | enum | unset (deny all) | `disabled` \| `allowlist` \| `pairing` \| `open`. **Unset or `disabled` rejects every inbound call** - you must set a policy to receive calls. `pairing` currently behaves like `allowlist`. |
| `allowFrom` | string[] | `[]` | Allowlisted callers, matched by AAD object id (case-insensitive) or phone number (digits only). Empty + `allowlist` = deny all. |
| `inboundGreeting` | string | - | Opening line the agent speaks on answer. |
| `mode` | enum | auto | `realtime` \| `streaming`. Auto-selects realtime if a realtime provider resolves. |
| `sessionScope` | enum | - | Conversation continuity: `per-phone` \| `per-call` \| `per-thread`. |
| `maxConcurrentCalls` | int | `4` | Concurrent active-call cap. |
| `maxDurationSeconds` | int | `0` (unlimited) | Hard cap on a single answered call's duration. |
| `staleCallReaperSeconds` | int | `120` | Tear down calls that stop being serviced after this long. |
| `maxVisionPerMinute` | int | - | Per-call vision spend cap. |
| `meetingRecap` | bool | - | Post an end-of-call recap / minutes. |
| `bilingual` | bool | - | Enable English/Arabic handling. |

## Group call

| Key | Type | Default | Meaning |
|---|---|---|---|
| `groupCall.requireAddress` | bool | - | In meetings (2+ humans), speak only when addressed. |
| `groupCall.wakePhrases` | string[] | - | Wake words that address the agent. |
| `groupCall.followUpWindowMs` | int | - | After being addressed, keep listening for follow-ups for this long. |

1:1 calls always answer regardless of these settings.

## Realtime

| Key | Type | Default | Meaning |
|---|---|---|---|
| `realtime.provider` | enum | `openai` | The realtime provider. |
| `realtime.providers.<id>.apiKey` | secret | - | Provider API key. |
| `realtime.providers.<id>.model` | string | - | Model, e.g. `gpt-realtime`. |
| `realtime.providers.<id>.azureEndpoint` | string | - | Azure OpenAI endpoint (selects Azure). |
| `realtime.providers.<id>.azureDeployment` | string | - | Azure deployment name. |
| `realtime.instructions` | string | - | System instructions for the voice agent. |
| `realtime.toolPolicy` | enum | `none` | Which agent tools the voice model may call: `safe-read-only` \| `owner` \| `none`. |
| `realtime.suppressInputDuringPlayback` | bool | - | Echo guard: ignore input while the agent is speaking. |
| `realtime.echoSuppressionWindowMs` | int | - | Echo-guard window. |
| `realtime.echoBargeInRms` | int | - | RMS threshold above which caller speech counts as barge-in. |

## Streaming

| Key | Type | Default | Meaning |
|---|---|---|---|
| `stt.provider` | string | - | Transcription provider id (streaming mode). |
| `stt.providers.<id>.apiKey` | secret | - | STT provider key. |

In streaming mode, TTS and the agent come from your OpenClaw configuration. If `stt.provider` is unset,
the plugin uses your configured transcription provider, then a VAD-segmented file fallback.

## Outbound

| Key | Type | Default | Meaning |
|---|---|---|---|
| `outbound.enabled` | bool | - | Enable outbound call-backs. |
| `outbound.workerBaseUrl` | string | - | StandIn outbound API base URL. |
| `outbound.tenantId` | string | - | Your AAD tenant id for outbound. |
| `outbound.answerTimeoutMs` | int | `120000` | No-answer timeout before voicemail/cancel. |
| `outbound.defaultMode` | enum | - | `notify` (speak and hang up) \| `conversation`. |

See [Outbound Calls](/openclaw-msteams-bridge/outbound-calls/).

## Secret-valued keys

These accept a literal string or an OpenClaw secret reference: `sharedSecret`,
`realtime.providers.*.apiKey`, `stt.providers.*.apiKey`. Prefer secret references in production.

## Full example

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
          "inboundGreeting": "Hi, you've reached the assistant. How can I help?",
          "sessionScope": "per-thread",
          "maxConcurrentCalls": 4,
          "maxVisionPerMinute": 30,
          "meetingRecap": true,
          "groupCall": {
            "requireAddress": true,
            "wakePhrases": ["assistant", "hey team"],
            "followUpWindowMs": 8000
          },
          "realtime": {
            "provider": "openai",
            "providers": {
              "openai": { "apiKey": "<key>", "model": "gpt-realtime" }
            },
            "toolPolicy": "safe-read-only",
            "suppressInputDuringPlayback": true
          }
        }
      }
    }
  }
}
```

### Azure OpenAI realtime

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

### Streaming

```jsonc
"mode": "streaming",
"stt": {
  "provider": "<your-stt-provider>",
  "providers": { "<your-stt-provider>": { "apiKey": "<key>" } }
}
```
