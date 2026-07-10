---
title: "Architecture"
description: "How the plugin is designed: system overview, call lifecycle, the two dialogue pipelines, module map, and the trust model."
---

This page is the contributor-level design of `@komaa/msteams-bridge`. It covers what runs where,
how a call moves through the system, and which module owns what. For the wire-level details see the
[Wire Protocol](/openclaw-msteams-bridge/wire-protocol/); for option-by-option settings see the
[Configuration Reference](/openclaw-msteams-bridge/configuration-reference/).

## System overview

Two processes cooperate at run time. The plugin runs inside your OpenClaw gateway and owns the
conversation. The hosted StandIn media bridge joins the Teams call and carries the media. They meet
on one HMAC-authenticated WebSocket per call, with the plugin as the server and StandIn as the
client.

```
                 (hosted service)                        (your machine / your gateway)
 ┌────────────┐   ┌─────────────────────┐  HMAC WebSocket  ┌──────────────────────────────┐
 │ Teams call │◄─►│ StandIn media bridge │═══════(dials in)═►│ @komaa/msteams-bridge          │
 └────────────┘   │  joins the meeting,  │  audio/video/    │  WS server + call brain       │
                  │  carries the media   │  events, JSON    │  inside the OpenClaw gateway  │
                  └─────────┬───────────┘                   └───────┬───────────┬──────────┘
                            ▲                                       │           │
                            │ REST: place / cancel call             │           │
                            │ (HMAC-signed, SSRF-guarded)           ▼           ▼
                  ┌─────────┴───────────┐                 ┌───────────────┐ ┌────────────────────┐
                  │ StandIn outbound API │                 │ realtime model │ │ OpenClaw agent      │
                  └─────────────────────┘                 │ (OpenAI/Azure) │ │ + STT/TTS providers │
                                                          └───────────────┘ └────────────────────┘
```

Key consequences of this shape:

- **The plugin never handles Teams media itself.** It receives PCM audio and JPEG frames on the
  WebSocket and sends PCM audio and avatar cues back. Everything Teams-specific happens on the
  StandIn side of the wire.
- **The plugin never dials out for media.** Media connections are always inbound from StandIn and
  always authenticated. The only outbound requests the plugin makes are the HMAC-signed REST calls
  that place or cancel an outbound Teams call.
- **The agent stays where it is.** Tools, memory, and the actual "brain" are your OpenClaw agent;
  the plugin routes the conversation to it.

## Call lifecycle

The call state machine lives in `src/call-lifecycle.ts` and every call, inbound or outbound, walks
the same states:

```
                 place call (outbound only)
                          │
                          ▼
   ┌──────────┐    ┌──────────┐    ┌──────────┐    ┌──────────┐
   │ initiate │───►│ ringing  │───►│ answered │───►│  active  │
   └──────────┘    └────┬─────┘    └──────────┘    └────┬─────┘
                        │ answer timeout                │ session.end / socket close /
                        ▼                               │ reaper / duration cap / policy
                  no-answer / voicemail                 ▼
                  (+ cancel ringing)              ┌──────────┐
                                                  │ terminal │  (teardown runs exactly once)
                                                  └──────────┘
```

How the wire events map onto it:

1. StandIn connects to `{path}/{callId}` and passes the HMAC handshake.
2. `session.start` arrives (within the 10 s pre-start timeout). For inbound calls the caller is
   checked against the inbound policy; a rejected caller closes the socket with `not-allowed`. For
   outbound calls the session is matched to the pending placed call, and a late answer after the
   answer timeout is declined.
3. With `requireRecordingStatus: true` the **recording gate** holds media processing until Teams
   reports recording `active` (`recording.status`).
4. Audio and video frames flow; the dialogue pipeline (below) runs; avatar cues go back.
5. The call ends by exactly one of: `session.end`, an abrupt socket close, the stale-call reaper
   (`staleCallReaperSeconds`), the duration cap (`maxDurationSeconds`), or a StandIn tier cutoff
   (which first delivers an `assistant.say` goodbye). Teardown is idempotent.

## The two dialogue pipelines

Both pipelines consume the same inbound 16 kHz PCM and produce the same outbound 16 kHz PCM; they
differ in what happens in between. `mode` picks one, or the runtime auto-selects realtime when a
realtime provider resolves.

```
 REALTIME (speech-to-speech)                    STREAMING (STT → agent → TTS)

 caller audio 16 kHz PCM                        caller audio 16 kHz PCM
      │ resample to 24 kHz                           │ VAD / segmentation
      ▼                                              ▼
 realtime model (OpenAI/Azure)                  STT provider → text
      │ model audio 24 kHz                           │
      │ + tool calls                                 ▼
      ▼ resample to 16 kHz                      OpenClaw agent → reply text
 outbound audio.frame 16 kHz                         │
      + speech.marks (visemes)                       ▼
      + expression cues                         TTS provider → audio
                                                     │ resample to 16 kHz
 vision: ambient push of the                         ▼
 latest changed frame                           outbound audio.frame 16 kHz
                                                     + speech.marks + expression

 lowest latency, continuous vision              works with any configured STT/TTS,
 needs a realtime provider key                  vision attached per turn
```

Shared by both: barge-in (`assistant.cancel`), verbal interrupts (EN/AR), the echo guard, the
group-call gate, DTMF, and bilingual handling. See
[Realtime and Streaming Modes](/openclaw-msteams-bridge/realtime-and-streaming-modes/).

## The three pillars and their wire messages

| Pillar | What it covers | Inbound messages | Outbound messages |
|---|---|---|---|
| **Perception** | hearing and seeing the call | `audio.frame`, `video.frame`, `participants`, `dtmf`, `recording.status` | - |
| **Dialogue** | the conversation itself | `session.start`, `session.end`, `assistant.say`, `ping` | `audio.frame`, `assistant.cancel`, `pong` |
| **Rendering** | what the caller sees | - | `expression`, `speech.marks`, `display.image` |

## Module map

Contributor-level map of `src/` (no internal bridge knowledge required to work on any of these):

| Area | Modules | Responsibility |
|---|---|---|
| Entry + manifest | `index.ts`, `openclaw.plugin.json` | Plugin registration, config schema, fail-closed secret check. |
| Config resolution | `config.ts`, `plugin-config.ts` | Types + resolver from raw config to runtime settings and defaults. |
| Transport + auth | `msteams-media-stream.ts`, `protocol.gen.ts` | The WebSocket server, HMAC handshake, replay guard, connection caps, message validation. |
| Runtime + mode select | `msteams-runtime.ts` | Session routing, inbound policy enforcement, realtime/streaming selection, outbound placement + answer timeout. |
| Call state machine | `call-lifecycle.ts` | initiate/ringing/answered/active/terminal, concurrency cap, exactly-once teardown. |
| Realtime | `msteams-realtime.ts`, `msteams-realtime-tools.ts` | The speech-to-speech session, tool exposure and policy. |
| Streaming | `msteams-streaming.ts`, `msteams-tts*.ts`, `telephony-*.ts` | STT segmentation, agent turns, TTS synthesis. |
| Vision | `vision-store.ts`, `vision-budget.ts`, `vision-consult.ts`, `msteams-video-frame.ts` | Frame ingest, latest-frame + keyframe history, per-call budget, look_at_screen. |
| Group etiquette | `group-call-gate.ts` | Speak-when-addressed, wake phrases, follow-up window. |
| Avatar cues | `expression.ts`, `viseme-estimate.ts` | Emotion inference and viseme timelines for lip-sync. |
| Minutes | `meeting-minutes-docx.ts` | Recap and `.docx` minutes with attribution. |
| Interrupts + echo | `verbal-interrupt.ts` | Deterministic stop phrases, echo/barge-in interplay. |

## Trust and security model

The plugin assumes the network between StandIn and itself is hostile and applies these layers:

| Layer | Mechanism | What it protects against |
|---|---|---|
| Handshake auth | HMAC-SHA256 over `{timestampMs}.{callId}`, constant-time compare | Unauthenticated peers connecting at all |
| Replay guard | 60 s timestamp window + single-use `(callId, ts, sig)` tuples | Captured handshakes being replayed |
| Fail-closed secret | Empty or malformed `sharedSecret` refuses every handshake | Accidentally running open |
| Inbound policy | `inboundPolicy` + `allowFrom`; **unset denies all callers** | Unknown callers reaching the agent |
| Recording gate | `requireRecordingStatus` holds media until recording is `active` | Processing an un-notified conversation |
| Outbound SSRF guard | Place-call requests pin validated public destinations | The gateway being steered at internal targets |
| Resource bounds | 64 total / 8 per-IP connections, 10 s pre-start, 2 MB inbound frames, 1 MB outbound buffer, concurrency + duration caps, stale reaper | Resource-exhaustion and wedged calls |

:::note
The shared secret authenticates **both** directions of business: the inbound media handshake and
the outbound place-call signature. Rotating it means updating StandIn and the plugin together.
:::

## Design principles

- **One wire, additive messages.** New message types and fields must degrade gracefully; older and
  newer peers interoperate. Unknown inbound types are ignored.
- **Fail closed.** No secret, no service; policy unset, no callers.
- **Exactly-once teardown.** Every exit path funnels through the same teardown; a second trigger is
  a no-op.
- **The bridge is a black box.** The plugin documents and depends only on the wire contract, never
  on how the StandIn media bridge is built.
