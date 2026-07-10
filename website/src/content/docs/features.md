---
title: "Features"
description: "Vision, meeting etiquette, avatar cues, recap and minutes, DTMF, bilingual EN/AR, and reliability guards."
---

A tour of what the plugin can do in a call, and the config that governs each.

## Dialogue

- **Two modes** - realtime speech-to-speech or streaming STT→agent→TTS. See
  [Realtime & Streaming Modes](/openclaw-msteams-voice/realtime-and-streaming-modes/).
- **Barge-in** - the caller can interrupt; the agent stops speaking (an `assistant.cancel` cancels the
  current TTS turn). Tune with `realtime.echoBargeInRms`.
- **Verbal interrupts** - deterministic interrupt phrases, EN/AR.
- **Echo suppression** - the agent does not hear itself; `suppressInputDuringPlayback` +
  `echoSuppressionWindowMs`.
- **DTMF / IVR** - keypad digits are surfaced to the agent.
- **Bilingual EN/AR** - `bilingual: true`.
- **Greeting** - `inboundGreeting` opens the call; the agent can greet meeting attendees by name from
  the roster.

## Perception (vision)

- Reads inbound **camera** and **screen-share (VBSS)** frames.
- Keeps a **continuous ambient view** in realtime mode; **attaches a keyframe per turn** in streaming
  mode.
- Retains a **scene-change keyframe history** so the agent can look back.
- Stays inside a **per-call budget** - `maxVisionPerMinute`.
- The agent reaches vision through the `look_at_screen` tool (live or from history).

## Rendering (avatar)

The plugin emits cues the StandIn bridge draws into the caller-facing video:

- **`expression`** - emotion cues (happy, sad, surprised, …).
- **`speech.marks`** - viseme lip-sync driven by speech marks.
- **`display.image`** - share an image as a picture-in-picture overlay or fullscreen (`show_to_caller`).

## Meeting etiquette

- **Wake-phrase gate** - in meetings (2+ humans), the agent stays quiet until addressed
  (`groupCall.requireAddress`, `groupCall.wakePhrases`), then answers through a short follow-up window
  (`groupCall.followUpWindowMs`).
- **1:1 always answers.**
- **Per-speaker attribution** - frames and utterances are tied to the participant who produced them.

## Agent tools

Exposed to the voice agent (governed by `realtime.toolPolicy`):

- `look_at_screen` - inspect the current screen-share/camera or recent history.
- `show_to_caller` - display an image to the caller.
- `post_meeting_minutes` - post minutes for the call.
- A background-task tool for longer work that reports back.

## Meeting recap and minutes

- **Recap** - an end-of-call summary of key points, decisions, and action items (`meetingRecap`).
- **`.docx` minutes** - on-demand minutes with per-person attribution.

## Outbound call-backs

Place a call, speak a result or hold a conversation, and hang up - with a no-answer/voicemail fallback
and cancel-ringing. See [Outbound Calls](/openclaw-msteams-voice/outbound-calls/).

## Reliability and security

- **Replay-proof HMAC handshake** on every connection (see [Wire Protocol](/openclaw-msteams-voice/wire-protocol/)).
- **Allowlist, closed by default** - `inboundPolicy` + `allowFrom`.
- **Recording-status gate** - hold media until recording is active (`requireRecordingStatus`).
- **Reconnect** - a dropped connection re-authenticates with a fresh handshake.
- **Stale-call reaper** - `staleCallReaperSeconds`.
- **Concurrency cap** - `maxConcurrentCalls`.
- **Duration cap** - `maxDurationSeconds`.
- **Backpressure / frame-size bounds** to keep a single agent from being overrun.
- **Graceful cutoff** - a tier limit triggers a spoken `assistant.say` goodbye before teardown.
