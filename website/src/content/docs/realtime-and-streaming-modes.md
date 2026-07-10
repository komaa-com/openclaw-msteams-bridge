---
title: "Realtime and Streaming Modes"
description: "Pick between realtime speech-to-speech and streaming STT-agent-TTS, and configure providers."
---

The plugin holds a spoken conversation in one of two modes. Set `mode` explicitly, or leave it out and
the runtime auto-selects **realtime** when a realtime provider resolves, otherwise **streaming**.

| | `realtime` | `streaming` |
|---|---|---|
| How it talks | speech-to-speech model | your OpenClaw STT → agent → TTS |
| Needs a realtime key | yes | no |
| Latency | lowest | higher (per turn) |
| Vision | continuous push | attached per turn |

Both modes support barge-in, verbal interrupts, echo suppression, the group-call gate, DTMF, and
bilingual EN/AR.

## Realtime

Realtime uses a single speech-to-speech model for the whole turn - the caller's audio goes to the
provider and the model's audio comes straight back, giving the lowest latency and a continuous ambient
vision view.

### OpenAI

```jsonc
"mode": "realtime",
"realtime": {
  "provider": "openai",
  "providers": { "openai": { "apiKey": "<key>", "model": "gpt-realtime" } },
  "instructions": "You are a concise, friendly voice assistant.",
  "toolPolicy": "safe-read-only"
}
```

### Azure OpenAI

Azure is the same `openai` provider with an endpoint and deployment; setting `azureEndpoint` selects
Azure:

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

### Realtime knobs

- `instructions` - the system prompt for the voice agent.
- `toolPolicy` - `safe-read-only`, `owner`, or `none`. Controls which agent tools the voice session may
  call. See [Features](/openclaw-msteams-voice/features/) for the tool set.
- `suppressInputDuringPlayback`, `echoSuppressionWindowMs`, `echoBargeInRms` - the echo guard and
  barge-in sensitivity. Raise `echoBargeInRms` if the agent interrupts itself; lower it if barge-in
  feels unresponsive.

## Streaming

Streaming runs a classic STT → agent → TTS pipeline, so it works with **any** provider you have
configured in OpenClaw and needs no realtime key. Vision is attached per turn rather than pushed
continuously.

```jsonc
"mode": "streaming",
"stt": {
  "provider": "<your-stt-provider>",
  "providers": { "<your-stt-provider>": { "apiKey": "<key>" } }
}
```

- **TTS and the agent** come from your OpenClaw configuration.
- **STT** uses `stt.provider` if set, else your configured transcription provider, else a
  VAD-segmented file fallback.
- Group-call gating, DTMF, and vision all work in streaming mode too.

## Choosing

- Want the most natural, lowest-latency conversation and have an OpenAI/Azure realtime key → **realtime**.
- Want to reuse your existing STT/TTS stack or avoid a realtime key → **streaming**.
- Not sure → leave `mode` unset and it will use realtime if a provider resolves.
