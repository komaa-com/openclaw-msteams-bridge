---
title: "Features"
description: "A full tour of dialogue, vision, avatar cues, meetings, tools, telephony, and reliability features."
---

A tour of what the plugin can do on a Teams call, and the config that governs each. Everything here
is implemented in this repo; the hosted StandIn media bridge handles the Teams media so these
features "just work" once you are connected.

## Dialogue

- **Realtime speech-to-speech** (`mode: "realtime"`) - OpenAI or Azure OpenAI Realtime; low latency,
  full-duplex feel. The default when a realtime provider resolves.
- **Streaming STT → agent → TTS** (`mode: "streaming"`) - half-duplex; uses your OpenClaw-configured
  STT/TTS/agent, so it works without a realtime provider.
- **Barge-in** - the caller can interrupt the agent mid-reply; playback is flushed
  (`assistant.cancel`) and the current turn is cancelled immediately. Tune the trigger with
  `realtime.echoBargeInRms`.
- **Verbal interrupts (EN/AR)** - deterministic "stop" / "توقف" phrase detection that cuts playback
  even when voice-activity detection misses it.
- **Echo suppression** - the agent does not hear (and answer) itself:
  `realtime.suppressInputDuringPlayback` ignores input while the agent speaks, and
  `realtime.echoSuppressionWindowMs` extends the guard window past playback.
- **Recording gate** - with `requireRecordingStatus: true` (default), no media is processed until
  Teams reports recording `active`.
- **Greeting** - `inboundGreeting` opens the call; in meetings the agent can greet attendees by name
  from the roster.
- **DTMF / IVR** - keypad presses are surfaced to the agent so it can run "press 1 to…" flows.
- **Bilingual EN/AR** (`bilingual: true`) - the agent detects and mirrors the caller's language and
  translates on request.

## Perception (vision)

- **Camera + screen-share (VBSS)** - the plugin ingests inbound `video.frame`s from both sources,
  attributed per participant.
- **`look_at_screen`** - the agent looks at the current frame (`live`) or at the recent
  scene-change keyframe history (`history`) to answer about something shown earlier.
- **Ambient view (realtime)** - the latest changed frame per source is pushed to the model
  periodically, so it stays visually aware between explicit looks; in streaming mode a keyframe is
  attached per turn instead.
- **Per-call vision budget** - `maxVisionPerMinute` caps total vision spend (explicit looks +
  ambient pushes); over budget, ambient pushes back off first.

## Rendering (avatar)

The plugin emits cues that the StandIn bridge draws into the caller-facing video:

- **`expression`** - emotion cues (`neutral`, `happy`, `sad`, `surprised`, `thinking`) inferred from
  the reply text; a thinking face shows while a tool runs.
- **`speech.marks`** - a viseme timeline (Azure viseme ids 0-21) drives lip-sync in time with the
  agent's audio.
- **`display.image`** - `show_to_caller` renders an image into the call, fullscreen or
  picture-in-picture, with an optional caption.

## Group / meeting etiquette

- **Speak only when addressed** - in a call with 2+ humans, the agent stays silent unless someone
  addresses it with a wake phrase (`groupCall.requireAddress`, `groupCall.wakePhrases`); a follow-up
  window (`groupCall.followUpWindowMs`) then lets the exchange continue without repeating the name.
- **1:1 always answers** - the gate only applies to meetings.
- **Per-speaker attribution** - utterances and frames are tied to the participant who produced them
  (used in minutes and vision answers).

## Agent tools

Exposed to the realtime voice agent, governed by `realtime.toolPolicy`
(`safe-read-only` | `owner` | `none`, default `none`):

| Tool | What it does |
|---|---|
| `look_at_screen` | Look at the shared screen/camera (live or history) and answer. |
| `show_to_caller` | Generate or fetch an image and show it on the bot's tile. |
| `post_meeting_minutes` | Summarize the call and post minutes to the Teams chat. |
| background task | Run longer work in the background and report back (optionally via an outbound call-back). |

## Meetings & productivity

- **End-of-call recap** (`meetingRecap: true`) - post minutes (key points, decisions, action items)
  to the Teams chat when the call ends.
- **On-demand minutes** - `post_meeting_minutes` (or asking the agent to summarize) posts minutes
  mid-call.
- **`.docx` minutes** - generated with per-person attribution.

## Outbound call-backs

Place a call, speak a result (`notify`) or hold a conversation (`conversation`), and hang up - with
a no-answer/voicemail fallback and cancel-ringing so the callee's phone stops ringing when the
plugin gives up. See [Outbound Calls](/openclaw-msteams-bridge/outbound-calls/).

## Sessions

- **`sessionScope`** - conversation continuity: `per-call` (fresh each call), `per-thread` (keyed by
  the Teams thread), or `per-phone` (keyed by the caller).

## Reliability and security

- **Replay-proof HMAC handshake** on every connection - constant-time compare, 60 s window,
  single-use tuples. See [Wire Protocol](/openclaw-msteams-bridge/wire-protocol/).
- **Fail-closed secret** - no shared secret (or a malformed one) means the server refuses to accept
  any handshake at all.
- **Inbound policy, closed by default** - `inboundPolicy` + `allowFrom`; with the policy unset
  (the default) every inbound call is denied, and `allowlist` with an empty list denies all too.
  Callers match by AAD object id or phone number.
- **Recording-status gate** - hold media until recording is active (`requireRecordingStatus`).
- **Reconnect** - a dropped connection re-authenticates with a fresh handshake.
- **Stale-call reaper** - `staleCallReaperSeconds` (default 120) tears down calls that stop being
  serviced.
- **Concurrency cap** - `maxConcurrentCalls` (default 4).
- **Duration cap** - `maxDurationSeconds` (default 0 = unlimited) bounds a single call's wall-clock
  time.
- **DoS guards** - 64 total / 8 per-IP connection caps, a 10 s pre-start timeout, a 2 MB frame cap,
  and a 1 MB outbound backpressure bound.
- **Graceful cutoff** - a StandIn tier limit triggers a spoken `assistant.say` goodbye before the
  call ends, instead of an abrupt drop.
