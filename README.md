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
- **Outbound call-backs** - hand over a long task and the agent rings you back and speaks the result
  when it is done, or answer a Teams chat with a phone call. An unanswered call is finalized after a
  no-answer timeout.
- **Meeting recap and minutes** - end-of-call summary of key points, decisions, and action items,
  plus an on-demand `.docx` of minutes with per-person attribution.
- **Avatar driver cues** - expression changes, viseme lip-sync, and picture-in-picture image
  sharing, rendered by the StandIn bridge.
- **Chat governance** - an "Ask about this" message action, voice-message transcription, an
  audit-log mirror, and outbound DLP redaction.
- **Secure transport** - a replay-proof HMAC handshake, a caller allowlist that is closed by
  default, and a recording-status gate that holds media until recording is active.

## Getting started

There are two ways to connect this plugin to Teams. Pick one - they differ in who owns the Teams bot.

### StandIn Managed Bot (recommended)

StandIn provides the Teams bot. You install **StandIn** from the Teams Store, connect this agent in
the StandIn portal, and paste **one secret** here. No Azure bot registration, no App ID, no client
secret, no endpoint configuration.

That one secret covers **both lanes** of the connection: calls arrive on the calling WebSocket and
Teams messages on the messages endpoint. They are two lanes of a single StandIn binding, not two
products, which is why there is a single value to paste and no enable flag to remember.

```jsonc
{
  "plugins": {
    "entries": {
      "msteams-bridge": {
        "config": {
          "enabled": true,
          // The connection secret from the StandIn portal. It turns on BOTH lanes:
          // calling (ws://:9442/msteams/calling) and messages (http://:9444/msteams/messages).
          "secret": "paste-the-value-from-the-StandIn-portal"
        }
      }
    }
  }
}
```

Then expose both endpoints to the internet (a public host, or a tunnel) and register them on your
connection in the StandIn portal as the **Agent calling URL** and **Agent messages URL**.

Bind address: **both lanes default to loopback** (`127.0.0.1`), matching the documented posture of a
tunnel that terminates TLS publicly and proxies in. The chat listener used to default to all
interfaces while calling defaulted to loopback, so a config that named no bind address at all put the
messages port on your LAN and the docs described the two as sharing one. If StandIn reaches this host
directly, set `bindAddress` to that interface (or `0.0.0.0` behind a firewall) - the HMAC keeps
unauthenticated callers out, but an open port is still an open port. `messagesBindAddress` overrides
the messages lane alone.

One agent instance serves **one** StandIn connection: the chat secret is a single value, scoped to
one tenant binding. Serving several tenants means running several instances, each with its own
secret. Never share one secret across tenants.

### Bring your own Azure bot (advanced)

You own the Microsoft Entra app, the client secret, and the Azure Bot resource. Choose this when you
need the bot to live entirely inside your own tenant. This plugin then adds **voice and video (CVI)**
on top of OpenClaw's Microsoft Teams **chat** channel, so set those up first:

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

Do these in order.

**Step 2 is BRING-YOUR-OWN-BOT only.** On the StandIn Managed Bot path there is no Teams channel to
create and no bot credentials to hold - StandIn owns the bot. If you followed the Managed Bot
quickstart above, skip straight from step 1 to step 3 and set `secret`.

1. **OpenClaw is installed and running** (host `>= 2026.6.10`).
2. **(BYO only) Microsoft Teams is added as a channel** (bot app + credentials).
3. **Install this plugin**, then restart the gateway so it loads:

   ```bash
   openclaw plugins install npm:@komaa/openclaw-msteams-bridge
   openclaw gateway restart
   ```

4. **Configure it** under `plugins.entries."msteams-bridge".config`. The minimum that actually WORKS -
   a secret alone leaves `inboundPolicy` unset, which denies every inbound call, so the bridge
   connects and then answers nothing:

   ```jsonc
   "msteams-bridge": {
     "config": {
       "enabled": true,
       "secret": "PASTE_THE_STANDIN_CONNECTION_SECRET",   // covers calling AND messages
       "bindAddress": "127.0.0.1",
       "inboundPolicy": "open"                            // or "allowlist" + allowFrom: [...]
     }
   }
   ```

   Set `secret` (the value the StandIn portal shows you). You also need your provider key. See
   [Configuration](#configuration) and [Security](#security) below. Nothing starts until a secret is
   set somewhere.
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

Config lives under `plugins.entries."msteams-bridge".config`. `secret` must match the value set
in your StandIn dashboard.

`bindAddress` defaults to **loopback for both lanes**, because the documented posture is a tunnel that
terminates TLS publicly and proxies to `127.0.0.1` - nothing is exposed on your LAN. Widen it to
`0.0.0.0` only when StandIn reaches this host directly, and only on a trusted interface.

You do **not** need OpenClaw's `voice-call` plugin. That one is telephony (Twilio, `fromNumber`,
webhook URLs); this one is self-contained. If your config nests Teams settings inside it as
`plugins.entries."voice-call".config.msteams`, lift that block up to
`plugins.entries."msteams-bridge".config` and drop the `voice-call` entry.

### Two models, two places

The **voice** model is configured here (`realtime.providers.<provider>.model`) and runs the spoken
conversation. The **agent** model - used by `openclaw_agent_consult` and `openclaw_agent_task` for
lookups and background work - is **inherited from your OpenClaw config**
(`agents.defaults.model.primary`). There is nothing to set here for it.

> **`FailoverError: Unknown model: openai/gpt-5.5`** logged mid-call, while the voice side works fine?
> On **v0.2.x and earlier** the agent model fell back to OpenClaw's compiled-in default pair rather
> than your configured agent, so a host on any other provider failed every consult - and only after
> the caller had already been greeted, since the voice lane has its own provider block. Nothing in
> your config names `gpt-5.5`, which is what makes it hard to find. **Fixed in v0.3.1.**

`responseModel` is an optional override for when the voice lane should consult a different model than
your default agent. Order: `responseModel` → `agents.defaults.model.primary` → OpenClaw's built-in
default (only with no agent configured at all).

### Realtime (OpenAI)

```jsonc
{
  "plugins": {
    "entries": {
      "msteams-bridge": {
        "config": {
          "enabled": true,
          "mode": "realtime",
          // Loopback + a tunnel; 0.0.0.0 only when StandIn reaches this host directly.
          "bindAddress": "127.0.0.1",
          "callingPort": 9442,
          "messagesPort": 9444,
          "path": "/msteams/calling",
          // ONE connection secret from the StandIn portal, covering calling AND messages.
          "secret": "<the connection secret from StandIn>",
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

This one block turns on **two** things:

- **Call me back when done.** `openclaw_agent_task` takes `deliverVia`. The default `"message"` sends the
  result as a Teams chat message; `"call"` rings the caller back and speaks it once the work finishes.
- **Chat-to-call.** The `call_me_with_the_answer` tool answers a Teams chat with a phone call. It rings
  the person it is chatting with - always, and only. There is no target parameter, deliberately: the
  agent reads untrusted text all day, and a tool that took a user id would turn any of it into "ring
  this person". Needs the messages lane as well.

Both are offered ONLY when `enabled`, `workerBaseUrl`, `tenantId` and `secret` are all set. Miss any one
and the plugin does not expose the capability at all - the background task quietly delivers by message
and the chat tool refuses and says why. An agent that promises "I'll call you back" and then cannot is
worse than one that answers in the chat.

Your Azure bot also needs **`Calls.Initiate.All`**, admin-consented, on top of the join/media
permissions. Without it every attempt fails at Graph no matter how this is configured.

> Needs **v0.4.0+**. Earlier versions accepted `deliverVia: "call"` and told the caller "I'll call you
> back", then never placed a call - the delivery was routed to a tool that does not exist in OpenClaw.
> Nothing errored, so there was no way to tell misconfiguration from breakage.

## StandIn Managed Bot: what differs from BYO

On a managed connection StandIn owns the Teams bot, so this plugin has **no customer Bot Framework
credentials**. Everything that would normally be sent "as your bot" goes through StandIn's signed
gateway instead, and two behaviours follow from that:

- **Meeting recap, minutes and background-task results are delivered as TEXT.** The gateway hop carries
  text and cards, not files, so the Word document is not attached; the message says so. BYO deployments
  post through your own bot and still get the `.docx`.
- **`post_chat_message` needs a meeting call.** A 1:1 call has no Teams conversation of its own, so
  there is nothing to post into and the tool is not offered for those calls.

**Call outcomes.** StandIn POSTs the real terminal state of a call you placed to
`{calling path}/outcome/{callId}`, signed with the same HMAC recipe as the WebSocket upgrade. The route
is served whenever the calling lane is - nothing to configure - and it is what makes a declined or busy
call finalize immediately instead of waiting out `outbound.answerTimeoutMs`. See
[Outbound calls](https://komaa-com.github.io/openclaw-msteams-bridge/outbound-calls/).

## Security

The plugin ships with **secure defaults**, and the recommendation is simple: keep them. Each option
below states its safe default and why it is safe. You only relax a default when your deployment
genuinely needs it, and only in the narrow way described.

| Option | Safe default | Why it is safe | When to change it |
|:--|:--|:--|:--|
| `secret` | none (fails closed) | Both lanes authenticate every connection with a replay-proof HMAC handshake keyed on this secret. With no secret the server refuses to start, so a misconfig can never expose an unauthenticated port. A non-string value coerces to empty and also fails closed. | Always set it, to a strong random value that matches your StandIn dashboard. Prefer an OpenClaw secret reference over a literal in config. |
| `inboundPolicy` | unset = deny all | Inbound calls are rejected until you name a policy, so the agent never answers an unknown caller by default. | Set `allowlist` and list trusted callers in `allowFrom` (by AAD object id or phone number). Reserve `open` for throwaway sandbox testing only. |
| `requireRecordingStatus` | `true` | Media is held until Teams reports recording is active, so the agent never sees or hears the call before participants have the recording indicator. This keeps you on the right side of Teams' notice expectations. | Leave it on. Only disable for a controlled test where no real participants are present. |
| `bindAddress` | `127.0.0.1` (loopback) | The WebSocket listens on localhost only, so it is unreachable from other hosts by default. | Widen to `0.0.0.0` only when the StandIn bridge runs on a different host, and only on a trusted or VPN-only interface behind your firewall. The HMAC handshake still guards it, but do not expose the port to the open internet. |
| `realtime.toolPolicy` | `none` | The voice model cannot invoke any agent tools, so a caller cannot drive tools by voice unless you opt in. | Use `safe-read-only` to allow read-only tools. Reserve `owner` (full tool access) for calls you have restricted to trusted owners via `inboundPolicy`. |
| Installer | applies the above | The one-line installer configures these secure defaults for you rather than leaving them blank. | If your policy forbids piping a script to a shell, download and read `install.sh` first, then run it, or follow the manual [Install](#install) steps. |

In short: set a strong `secret`, keep `inboundPolicy` restrictive with an explicit `allowFrom`,
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
| `path` | Calling WebSocket path (default `/msteams/calling`; was `/voice/msteams/stream` before the rename - see [Upgrading](#upgrading-from-msteams-voice)) |
| `secret` | HMAC secret for BOTH lanes; must match StandIn |
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
| `responseModel` | *optional* - overrides the agent model for consult/task; omit to inherit `agents.defaults.model.primary` |
| `outbound.enabled` + `workerBaseUrl` + `tenantId` | enables call-back delivery (`deliverVia: "call"`) and chat-to-call (`call_me_with_the_answer`); also needs `Calls.Initiate.All` on the bot |
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
