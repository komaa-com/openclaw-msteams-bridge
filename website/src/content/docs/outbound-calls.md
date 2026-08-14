---
title: "Outbound Calls"
description: "Let the agent place Teams calls: configuration, no-answer handling, and cancel-ringing."
---

Outbound lets the agent **place** a Teams call (a call-back), speak a result or hold a conversation,
and hang up. It is optional and off unless configured.

## Enable it

```jsonc
"outbound": {
  "enabled": true,
  "workerBaseUrl": "https://<your-standin-endpoint>",
  "tenantId": "<aad-tenant-id>",
  "answerTimeoutMs": 120000,
  "defaultMode": "notify"
}
```

| Key | Meaning |
|---|---|
| `outbound.enabled` | Turn outbound on. |
| `outbound.workerBaseUrl` | The StandIn outbound API base URL. |
| `outbound.tenantId` | Your AAD tenant id for the callee. |
| `outbound.answerTimeoutMs` | How long to wait for an answer before treating it as no-answer (default `120000`). |
| `outbound.defaultMode` | `notify` (speak the message and hang up) or `conversation` (stay and talk). |

## How a call is placed

The agent triggers an outbound call through its realtime tools. Under the hood the plugin makes an
**HMAC-signed** request to the StandIn outbound API:

- Signature headers `x-standin-timestamp` / `x-standin-signature` (the legacy
  `x-openclawteamsbridge-*` names are still accepted), signed over
  `"{timestamp}.{userObjectId}"` with your `secret`.
- The request identifies the callee (`userObjectId`) and `tenantId`; StandIn returns a `callId`.
- Requests are SSRF-guarded.

Once the callee answers, the same per-call WebSocket session begins and the conversation runs exactly
like an inbound call.

## No answer and cancel

- The StandIn worker reports the **real** terminal state as soon as it knows it, so a declined or busy
  call finalizes immediately instead of waiting out `answerTimeoutMs`. See the outcome callback below.
- If no outcome arrives, the attempt is finalized as **no-answer** when `answerTimeoutMs`
  expires. That timer is the fallback, not the primary path.
- The plugin also best-effort **cancels the ringing** call so a late pickup does not strand the callee
  in a dead call.
- A **late answer** (after the timeout) is declined cleanly.

### The outcome callback

StandIn POSTs the terminal state of a call you placed to the calling lane:

```
POST {calling path}/outcome/{callId}
X-StandIn-Timestamp: <ms epoch>
X-StandIn-Signature: HMAC-SHA256(secret, "{ts}.{callId}")

{"outcome": "declined"}
```

- **Path**: under the calling prefix (default `/msteams/calling/outcome/{callId}`), because a tunnel
  usually forwards only that prefix. No extra port or route to open.
- **Auth**: the same HMAC contract as the WebSocket upgrade, over the **callId** - so a signature
  captured for one call cannot be replayed against another, and the timestamp must be inside the
  replay window.
- **Body**: `outcome` is one of `answered`, `no-answer`, `declined`, `busy`, `failed`. `answered` is a
  no-op: the media socket attaches and the call proceeds normally.
- **Responses**: `200 {"ok": true}`; `401` on a bad signature or a stale timestamp; `400` on a
  malformed body; `404` on any other path.

Nothing to configure - the route is served whenever the calling lane is. Before this existed the
plugin learned only "no answer", and only after its own timeout, so a declined call and an unanswered
one were indistinguishable.

## Modes

- **`notify`** - the agent delivers a message and hangs up. Good for reminders and alerts.
- **`conversation`** - the agent stays on the line for a back-and-forth.

## Tips

- Outbound needs the same `secret` as inbound - it signs the place-call request.
- Set `tenantId` to the callee's tenant.
- Keep `answerTimeoutMs` realistic (people take a few rings); too short gives up before they pick up.
