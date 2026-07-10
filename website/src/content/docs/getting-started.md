---
title: "Getting Started"
description: "Install the plugin, connect to the StandIn sandbox, and make your first Teams voice call."
---

This walks you from nothing to a working Teams voice call with your OpenClaw agent.

## Prerequisites

- **An OpenClaw install**, host `>= 2026.6.10`. Follow [docs.openclaw.ai](https://docs.openclaw.ai).
- **Microsoft Teams set up as an OpenClaw channel** (the chat channel). This plugin adds voice/video
  on top of it. See the [OpenClaw Teams channel docs](https://docs.openclaw.ai/channels/msteams).
- **A StandIn connection.** The [sandbox](https://standin.komaa.com/sandbox) is free and needs no
  Teams bot of your own - perfect for a first call. See [Connecting to StandIn](/openclaw-msteams-voice/connecting-to-standin/).
- **For realtime mode**, a realtime voice provider key (OpenAI or Azure OpenAI). For streaming mode,
  your OpenClaw-configured STT/TTS/agent is enough.

## 1. Install the plugin

The one-line installer detects your OpenClaw install and walks you through the config (mode, shared
secret, provider key), then prints next steps:

```bash
curl -fsSL https://standin.komaa.com/install.sh | bash
```

Prefer to do it by hand?

```bash
openclaw plugins install npm:@komaa/msteams-voice
openclaw gateway restart
```

It is also on [ClawHub](https://clawhub.ai): `openclaw plugins install clawhub:@komaa/msteams-voice`
(OpenClaw falls back to npm automatically). The package ships prebuilt (v0.1.10+) - no build step
either way.

## 2. Try it on the StandIn sandbox

The sandbox is the fastest path to a first call and needs no Azure/Teams bot of your own:

1. Open [standin.komaa.com/sandbox](https://standin.komaa.com/sandbox) and follow it to generate a
   Teams meeting link. It gives you a **shared secret** for the session.
2. Put that secret in your plugin config as `sharedSecret` (below).
3. Join the meeting yourself; the shared StandIn bot joins and connects to your plugin.

The sandbox is time-limited (about 5 minutes/day per session) - enough to see it end to end. When the
limit is reached, the agent speaks a short goodbye and the call ends gracefully.

## 3. Minimal configuration

Config lives under `plugins.entries."msteams-voice".config`. A minimal realtime setup:

```jsonc
{
  "plugins": {
    "entries": {
      "msteams-voice": {
        "config": {
          "enabled": true,
          "mode": "realtime",
          "bindAddress": "0.0.0.0",          // so the hosted bridge can reach the plugin
          "port": 9442,
          "sharedSecret": "<the secret from StandIn>",
          "realtime": {
            "provider": "openai",
            "providers": { "openai": { "apiKey": "<key>", "model": "gpt-realtime" } }
          }
        }
      }
    }
  }
}
```

- `bindAddress: "0.0.0.0"` lets the hosted bridge connect (the default `127.0.0.1` only accepts local
  connections).
- `sharedSecret` must match the value StandIn uses, or the handshake is rejected.
- Leave `mode` out to auto-select: realtime if a provider resolves, otherwise streaming.

See the full [Configuration Reference](/openclaw-msteams-voice/configuration-reference/) for every option, and
[Realtime & Streaming Modes](/openclaw-msteams-voice/realtime-and-streaming-modes/) for provider setup.

## 4. Restart and call

Restart the gateway so the config takes effect:

```bash
openclaw gateway restart
```

Then place (or join) the Teams call. You should hear the agent, see its avatar, and - if you turn on
your camera or share your screen - it can see you too.

## Next steps

- Add **your own Teams bot** for real inbound calls: [Connecting to StandIn](/openclaw-msteams-voice/connecting-to-standin/).
- Lock down who can call you: `inboundPolicy` + `allowFrom` in the [Configuration Reference](/openclaw-msteams-voice/configuration-reference/).
- Turn on **meeting recap**, **vision budget**, **group-call etiquette**: [Features](/openclaw-msteams-voice/features/).
- Something not working? [Troubleshooting](/openclaw-msteams-voice/troubleshooting/).
