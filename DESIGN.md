# Design

This is the architecture note for contributors. It explains what `@komaa/openclaw-msteams-bridge` is, how it
is put together, and how it talks to the outside world. For usage, start with the
[README](./README.md) and the [docs site](https://komaa-com.github.io/openclaw-msteams-bridge/);
the site's [Architecture page](https://komaa-com.github.io/openclaw-msteams-bridge/architecture/) is
the diagrammed version of this note.

## What it is

`@komaa/openclaw-msteams-bridge` is an **OpenClaw channel plugin** that adds Microsoft Teams **CVI** (customer
voice + video interaction) to an OpenClaw agent. It layers real-time voice and video on top of
OpenClaw's Teams chat channel: the agent hears the caller, sees their camera and screen-share, talks
back with low latency, and appears in the call as a lip-synced avatar.

It is a single plugin that depends only on the published `openclaw` plugin SDK and `api.runtime`.
There is no fork, no vendored runtime, and no trusted-plugin privileges.

## The moving parts

There are two processes at run time:

1. **This plugin**, running inside your OpenClaw gateway. It hosts a local WebSocket **server** and
   owns the conversation: speech, vision, tools, avatar cues, and the call state machine.
2. **The StandIn media bridge** - a hosted service that actually joins the Teams call. It carries the
   media and, for each call, **dials into this plugin's WebSocket server**. This plugin never dials
   out for media; it only accepts authenticated inbound connections from StandIn.

```
Teams call  <-->  StandIn media bridge  ==WebSocket==>  @komaa/openclaw-msteams-bridge  <-->  OpenClaw agent
                    (hosted, joins the call)             (this plugin, WS server)      + voice provider
```

Because the plugin is a server, the same build works whether StandIn runs in the hosted sandbox or
against your own paired Teams bot - only the shared secret and the identity differ. See
[Connecting to StandIn](https://komaa-com.github.io/openclaw-msteams-bridge/connecting-to-standin/).

## The wire

One WebSocket per active call, opened by the bridge to `/{path}/{callId}` and authenticated with a
replay-proof HMAC handshake (timestamp + signature headers, signed over `"{timestamp}.{callId}"`).
After the handshake the two sides exchange a small JSON message protocol - `session.start`,
`audio.frame`, `video.frame`, `assistant.say`, `assistant.cancel`, `expression`, `speech.marks`,
`display.image`, and so on. Audio is PCM 16 kHz / 16-bit / mono, base64-framed, both directions. The
full contract is the [Wire Protocol](https://komaa-com.github.io/openclaw-msteams-bridge/wire-protocol/)
page. The protocol is deliberately transport-only: this plugin knows nothing about how the bridge
produces Teams media, and the bridge knows nothing about the agent.

## Call lifecycle

The lifecycle coordinator (`src/call-lifecycle.ts`) runs a small state machine per call:

```
initiate -> ringing -> answered -> active -> terminal
```

- An inbound connection + `session.start` moves a call to **active**; an active-call registry
  deduplicates and enforces `maxConcurrentCalls`.
- **Reconnects** are new authenticated handshakes with a fresh timestamp; a stale-call reaper
  (`staleCallReaperSeconds`) tears down calls that stop being serviced.
- On `session.end`, an abrupt socket close, or a `maxDurationSeconds` deadline, teardown runs once and
  releases the session, timers, and provider connections.
- At a limit cutoff (sandbox time cap, free daily budget, or a paid max-minutes governor) the bridge
  sends `assistant.say`; the agent speaks that line, then the call ends gracefully.

## Dialogue: two modes

The runtime (`src/msteams-runtime.ts`) picks a mode per the config, auto-selecting **realtime** when a
realtime provider resolves, otherwise **streaming**:

- **Realtime** (`src/msteams-realtime.ts`) - a speech-to-speech model (OpenAI or Azure OpenAI). Caller
  audio streams to the provider; model audio streams back, chunked into wire frames. Vision is pushed
  continuously.
- **Streaming** (`src/msteams-streaming.ts`, `msteams-tts*.ts`, `telephony-*.ts`) - STT to the OpenClaw
  agent to TTS. Any provider works; vision is attached per turn.

Both modes share barge-in and a deterministic verbal-interrupt path (`verbal-interrupt.ts`), echo
suppression, the group-call gate, DTMF, and bilingual EN/AR handling.

## Perception (vision)

Inbound camera and screen-share (VBSS) frames arrive as `video.frame` messages and land in a
per-source latest-frame buffer with a scene-change keyframe history (`vision-store`,
`msteams-video-frame.ts`). Spend is capped per call (`vision-budget.ts`, `maxVisionPerMinute`). The
agent reaches vision through tools (`look_at_screen`) and, in realtime mode, an ambient view.

## Rendering (avatar cues)

The plugin does not draw anything itself - it emits **cues** the bridge renders into the caller-facing
video: `expression` (emotion), `speech.marks` (viseme lip-sync, `viseme-estimate.ts` / `expression.ts`),
and `display.image` (picture-in-picture / fullscreen image sharing).

## Group / meeting etiquette

`group-call-gate.ts` keeps the agent quiet in meetings until it is addressed by a wake phrase, then
opens a short follow-up window; 1:1 calls always answer. Frames and utterances are attributed to the
speaker from the roster so the agent can greet and reference people by name.

## Outbound call-backs

When `outbound.enabled`, the agent can place a call through the StandIn outbound API (an HMAC-signed
REST call), speak a result or hold a conversation, and hang up - with a no-answer / voicemail fallback
and a cancel-ringing path so a late pickup does not strand the callee. See
[Outbound Calls](https://komaa-com.github.io/openclaw-msteams-bridge/outbound-calls/).

## Module map

| Area | Files |
|---|---|
| Entry / plugin manifest | `src/index.ts`, `openclaw.plugin.json` |
| Config resolution | `src/config.ts`, `src/plugin-config.ts` |
| Transport + handshake | `src/msteams-media-stream.ts`, `src/protocol.gen.ts` |
| Runtime / mode select | `src/msteams-runtime.ts` |
| Call state machine | `src/call-lifecycle.ts` |
| Realtime dialogue | `src/msteams-realtime.ts`, `src/msteams-realtime-tools.ts` |
| Streaming dialogue | `src/msteams-streaming.ts`, `src/msteams-tts*.ts`, `src/telephony-*.ts` |
| Vision | `src/vision-store.ts`, `src/vision-budget.ts`, `src/vision-consult.ts`, `src/msteams-video-frame.ts` |
| Group gate | `src/group-call-gate.ts` |
| Avatar cues | `src/expression.ts`, `src/viseme-estimate.ts` |
| Meeting minutes | `src/meeting-minutes-docx.ts` |
| Allowlist / security | `src/allowlist.ts` |

## Non-goals

- The plugin never touches Teams media internals - that is entirely the bridge's job, reached only
  through the documented wire protocol.
- No trusted-plugin privileges and no runtime fork: everything runs on the public OpenClaw plugin SDK.

## Design principles

- **Fail closed.** A missing shared secret, an empty allowlist, or an inactive recording gate denies,
  it does not fall open.
- **Transport-only coupling.** The plugin and the bridge share a JSON message contract and nothing
  else, so either side can evolve independently.
- **Bounded by default.** Connection caps, per-call vision budgets, concurrency limits, and reapers
  keep a single agent from being overrun.
