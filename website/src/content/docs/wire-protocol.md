---
title: "Wire Protocol"
description: "The WebSocket contract with the StandIn media bridge: HMAC handshake, audio format, and every message type."
---

This is the contract the **StandIn media bridge** speaks with the plugin over the WebSocket. It is
useful if you are debugging a connection, writing tests, or building a compatible integration. It
describes only the client-facing wire - nothing about how the bridge produces Teams media.

The design goal is **forward compatibility**: messages are camelCase JSON, additive, and tolerant -
unknown fields are ignored and unknown message types degrade gracefully, so older and newer peers
interoperate.

## The upgrade

The plugin hosts a WebSocket **server**; the bridge is the **client**. For each Teams call, StandIn
opens **one WebSocket** to:

```
ws://<bindAddress>:<port>{path}/{callId}
```

with the default path `/voice/msteams/stream`, so for example
`ws://host:9442/voice/msteams/stream/abc123`. `{callId}` in the URL is authenticated by the HMAC
headers and later cross-checked against the `callId` in the `session.start` body.

### The two HMAC headers

On the upgrade request, StandIn sends:

| Header | Value |
|---|---|
| `x-openclawteamsbridge-timestamp` | the signing timestamp, in **milliseconds** (Unix epoch) |
| `x-openclawteamsbridge-signature` | the signature (hex) |

The signature is:

```
HMAC-SHA256(sharedSecret, "{timestampMs}.{callId}")   # hex
```

Example:

```
x-openclawteamsbridge-timestamp: 1720598400000
x-openclawteamsbridge-signature: 9f8c...   (hex digest of "1720598400000.abc123")
```

Verification order: header presence, timestamp parse + window, constant-time signature compare
(incoming signatures are normalized with `trim` + lowercase first), then the single-use replay
check. On success the WebSocket is accepted; any failure rejects the upgrade.

- **Window:** the timestamp must be within the replay window of the server clock - default
  **60,000 ms** (60 s).
- **Replay guard:** each accepted `(callId, timestamp, signature)` tuple is **single-use**; a
  captured handshake cannot be replayed. A reconnect authenticates by re-signing with a fresh
  timestamp.
- **Fail closed:** if `sharedSecret` is empty or unset, the server refuses to start; a non-string
  secret value coerces to empty and rejects every handshake.

### Connection guards

After auth, the server bounds resource use:

| Guard | Default |
|---|---|
| Total concurrent connections | 64 |
| Connections per IP | 8 |
| Time to send `session.start` after connecting | 10 s (then the socket is reaped) |
| Max inbound frame size | 2 MB |
| Outbound backpressure cap | 1 MB buffered, then the connection is considered stalled |

## Audio format

All audio on the wire is **PCM, 16 kHz, 16-bit signed, mono**, base64-encoded in the
`payloadBase64` field, in both directions. (Realtime models typically run at 24 kHz internally; the
plugin resamples on both sides.)

## Inbound messages (bridge → plugin)

All frames are JSON text with a `type` discriminator. Inbound messages are schema-validated on
receipt. Fields are camelCase.

### `session.start`

Opens the call. Must be the first message, within the pre-start timeout.

```json
{
  "type": "session.start",
  "callId": "abc123",
  "threadId": "19:meeting_xyz@thread.v2",
  "caller": { "aadId": "00000000-...", "displayName": "Ada Lovelace", "tenantId": "..." },
  "recordingStatus": "active",
  "direction": "inbound"
}
```

| Field | Type | Notes |
|---|---|---|
| `callId` | string (required) | Must match the URL `callId`. |
| `threadId` | string (required) | Teams chat/thread id. |
| `caller` | object | `aadId`, `displayName`, `tenantId` - best-effort. Drives the allowlist check and greeting. |
| `recordingStatus` | string | `active` \| `inactive` \| `unknown`. |
| `direction` | string | `inbound` \| `outbound` (outbound = a call the plugin placed). Any other value the bridge sends (for example a meeting join) is normalized to `inbound`. |

### `session.end`

```json
{ "type": "session.end", "reason": "hangup" }
```

`reason` is a free-form string. The plugin also tears the session down on an abrupt socket close.

### `recording.status`

```json
{ "type": "recording.status", "status": "active" }
```

`status` (required): `active` \| `inactive` \| `unknown`. With `requireRecordingStatus: true`
(default), media processing is gated until `active`.

### `audio.frame`

```json
{ "type": "audio.frame", "seq": 42, "timestampMs": 840, "payloadBase64": "…", "speakerName": "Ada" }
```

| Field | Type | Notes |
|---|---|---|
| `seq` | int | Monotonic frame sequence number (per direction); the receiver may use it to detect drops. |
| `timestampMs` | int | Playout timestamp in ms. |
| `payloadBase64` | string (required) | PCM 16 kHz / 16-bit / mono, base64. |
| `speakerName` | string | Optional - unmixed-audio speaker attribution (used for minutes). |

### `video.frame`

```json
{ "type": "video.frame", "source": "screenshare", "ts": 1234, "width": 1280,
  "height": 720, "mime": "image/jpeg", "dataBase64": "…",
  "participantId": "…", "participantName": "Ada" }
```

| Field | Type | Notes |
|---|---|---|
| `source` | string (required) | `camera` \| `screenshare`. |
| `ts` | int | Frame timestamp (a new `ts` = a new scene). |
| `width`, `height` | int | Pixel dimensions. |
| `mime` | string | Typically `image/jpeg`. |
| `dataBase64` | string (required) | The encoded image. |
| `participantId`, `participantName` | string | Optional attribution. |

### `participants`

```json
{ "type": "participants", "count": 3 }
```

Current human participant count; drives the group-call gate (2+ humans = meeting etiquette).

### `dtmf`

```json
{ "type": "dtmf", "digit": "1" }
```

`digit` (required): `0`-`9`, `*`, or `#`. In-band keypad tones surfaced to the agent for IVR-style
flows.

### `ping`

```json
{ "type": "ping", "ts": 1720598400000 }
```

The plugin replies with a `pong` echoing `ts`.

### `assistant.say`

```json
{ "type": "assistant.say", "text": "We're at time - thanks for calling, goodbye!" }
```

`text` (required). StandIn asks the agent to speak this line in its own voice - for example a brief
goodbye right before a tier-limit cutoff (sandbox/free daily cap, or a subscription max-minutes
governor). The agent speaks it, then StandIn ends the call gracefully.

## Outbound messages (plugin → bridge)

### `audio.frame`

Agent audio to play into the call. Same shape as inbound `audio.frame` (`seq`, `timestampMs`,
`payloadBase64`).

### `assistant.cancel`

```json
{ "type": "assistant.cancel", "turnId": 7 }
```

Barge-in: tell the bridge to flush playback for `turnId` so the caller's interruption takes effect
immediately.

### `expression`

```json
{ "type": "expression", "emotion": "happy" }
```

Avatar emotion cue for the bot's video tile (for example `neutral`, `happy`, `sad`, `surprised`,
`thinking`). Cosmetic and best-effort.

### `speech.marks`

```json
{ "type": "speech.marks", "ts": 840, "marks": [ { "tMs": 0, "visemeId": 12 }, { "tMs": 60, "visemeId": 3 } ] }
```

Viseme timeline for lip-sync. Each mark is `{tMs, visemeId}`: `tMs` is milliseconds from utterance
start, `visemeId` is an Azure viseme id (0-21).

### `display.image`

```json
{ "type": "display.image", "dataBase64": "…", "mime": "image/png", "ts": 0,
  "durationMs": 5000, "mode": "fullscreen", "caption": "Here's the chart" }
```

`show_to_caller`: render an image into the call video.

| Field | Type | Notes |
|---|---|---|
| `dataBase64` | string (required) | The encoded image. |
| `mime` | string | Image MIME type. |
| `durationMs` | int, optional | How long to show it. |
| `mode` | string, optional | `fullscreen` \| `overlay` (picture-in-picture). |
| `ts` | int, optional | Timestamp. |
| `caption` | string, optional | Caption drawn with the image. |

### `pong`

```json
{ "type": "pong", "ts": 1720598400000 }
```

Keepalive reply echoing the inbound `ping` timestamp.

## Lifecycle in brief

1. The bridge connects to `{path}/{callId}` and passes the HMAC handshake.
2. The bridge sends `session.start`; the plugin checks the inbound policy and creates the call
   session (a rejected caller closes the socket with reason `not-allowed`).
3. Audio (and video) frames flow both ways; the plugin emits avatar cues as needed.
4. On a tier-limit cutoff the bridge sends `assistant.say`; the agent speaks it.
5. `session.end` (or a socket close, a stale-call reaper, or a duration cap) tears the session down
   exactly once.

## Notes for integrators

- Keep the audio format exact (16 kHz / 16-bit / mono, base64) - resample on your side if needed.
- Respect the replay window and single-use handshake; re-signing with a fresh timestamp is how a
  reconnect authenticates.
- Send well-formed JSON with the fields above; inbound messages are validated and malformed frames
  are dropped.
