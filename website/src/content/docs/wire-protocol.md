---
title: "Wire Protocol"
description: "The WebSocket contract the StandIn media bridge speaks with the plugin: handshake, HMAC auth, and every message type."
---

This is the contract between the StandIn media bridge and the plugin: one WebSocket per active call,
an HMAC-authenticated upgrade, then a small JSON message protocol. It is documented here so you can
understand call behavior, debug a connection, or build a compatible integration. It describes only the
client-facing wire - nothing about how the bridge produces Teams media.

## Connection

- The plugin hosts a WebSocket **server**; the bridge is the **client** and opens one connection per
  call to `{path}/{callId}` (default path `/voice/msteams/stream`), with the call id in the URL.
- Audio is **PCM 16 kHz, 16-bit, mono**, base64-encoded in frame payloads, both directions.

## Handshake (HMAC)

On the WebSocket upgrade the bridge sends two headers:

| Header | Value |
|---|---|
| `x-openclawteamsbridge-timestamp` | Unix epoch milliseconds |
| `x-openclawteamsbridge-signature` | `HMAC-SHA256(sharedSecret, "{timestamp}.{callId}")`, hex |

Verification:

- The signature is computed over the exact string `"{timestamp}.{callId}"` and compared
  constant-time. Incoming signatures are normalized (`trim` + lowercase) before comparison.
- A **replay window** (default 60,000 ms) rejects stale timestamps, and each `(callId, timestamp,
  signature)` tuple is single-use, so a captured handshake cannot be replayed.
- If `sharedSecret` is empty/unset, the handshake fails closed.

### Guardrails

The server bounds resource use: up to 64 total connections and 8 per IP, a 10-second window to send
`session.start` after connecting, a 2 MB max inbound frame, and a 1 MB outbound backpressure cap.

## Inbound messages (bridge → plugin)

Each is a JSON object with a `type` field. Validated on receipt.

| `type` | Fields | Meaning |
|---|---|---|
| `session.start` | `callId`, `threadId`, `caller{aadId, displayName, tenantId}`, `recordingStatus?`, `direction?` | Call started; begins the session. |
| `session.end` | `reason` | Call ended. |
| `recording.status` | `status` (`active`\|`inactive`\|`unknown`) | Teams recording state; gates media when `requireRecordingStatus`. |
| `audio.frame` | `seq`, `timestampMs`, `payloadBase64`, `speakerName?` | Inbound caller audio. |
| `video.frame` | `source` (`camera`\|`screenshare`), `ts`, `width`, `height`, `mime`, `dataBase64`, `participantId?`, `participantName?` | Inbound video/screen-share frame. |
| `participants` | `count` | Current human participant count (drives the group gate). |
| `dtmf` | `digit` (`0-9`, `*`, `#`) | Keypad entry. |
| `ping` | `ts` | Liveness ping; the plugin replies `pong`. |
| `assistant.say` | `text` | Speak this line (used for the cutoff goodbye). |

## Outbound messages (plugin → bridge)

| `type` | Fields | Meaning |
|---|---|---|
| `audio.frame` | `seq`, `timestampMs`, `payloadBase64`, `speakerName?` | Agent audio to play into the call. |
| `assistant.cancel` | `turnId` | Cancel the current TTS turn (barge-in). |
| `expression` | `emotion` | Avatar face cue (e.g. happy, sad, surprised). |
| `speech.marks` | `ts`, `marks[]` | Viseme timeline for lip-sync. |
| `display.image` | `dataBase64`, `mime`, `durationMs?`, `mode?` (`fullscreen`\|`overlay`), `ts?`, `caption?` | Show an image in the call (picture-in-picture or fullscreen). |
| `pong` | `ts` | Reply to `ping`. |

## Lifecycle in brief

1. Bridge connects to `{path}/{callId}` and passes the HMAC handshake.
2. Bridge sends `session.start`; the plugin creates the call session.
3. Audio (and video) frames flow both ways; the plugin emits avatar cues as needed.
4. On a limit cutoff the bridge sends `assistant.say`; the agent speaks it.
5. `session.end` (or a socket close, or a duration cap) tears the session down once.

## Notes for integrators

- Keep the audio format exact (16 kHz/16-bit/mono, base64) - resample on your side if needed.
- Respect the replay window and single-use handshake; re-signing with a fresh timestamp is how a
  reconnect authenticates.
- The plugin is strict about message shape; send well-formed JSON with the fields above.
