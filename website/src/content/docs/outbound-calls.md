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

- Signature headers `x-openclawteamsbridge-timestamp` / `x-openclawteamsbridge-signature`, signed over
  `"{timestamp}.{userObjectId}"` with your `sharedSecret`.
- The request identifies the callee (`userObjectId`) and `tenantId`; StandIn returns a `callId`.
- Requests are SSRF-guarded.

Once the callee answers, the same per-call WebSocket session begins and the conversation runs exactly
like an inbound call.

## No answer, voicemail, and cancel

- If no one answers within `answerTimeoutMs`, the attempt is finalized as **no-answer / voicemail**.
- The plugin also best-effort **cancels the ringing** call so a late pickup does not strand the callee
  in a dead call.
- A **late answer** (after the timeout) is declined cleanly.

## Modes

- **`notify`** - the agent delivers a message and hangs up. Good for reminders and alerts.
- **`conversation`** - the agent stays on the line for a back-and-forth.

## Tips

- Outbound needs the same `sharedSecret` as inbound - it signs the place-call request.
- Set `tenantId` to the callee's tenant.
- Keep `answerTimeoutMs` realistic (people take a few rings); too short causes premature voicemail.
