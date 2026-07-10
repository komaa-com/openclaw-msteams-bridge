---
title: "Troubleshooting"
description: "Fixes for handshake failures, missing audio, gated media, allowlist blocks, and provider errors."
---

Common issues and what to check.

## The bridge can't connect / handshake rejected

- **Shared secret mismatch.** `sharedSecret` must exactly match the value StandIn uses. Re-copy it
  from the sandbox session or the dashboard.
- **Bind address.** The default `127.0.0.1` only accepts local connections. Set `bindAddress` to
  `0.0.0.0` so the hosted bridge can reach the plugin.
- **Port not reachable.** Confirm `port` (default `9442`) is open to StandIn and not taken by another
  process.
- **Clock skew.** The handshake enforces a replay window (60 s). If the host clock is far off, the
  timestamp is rejected - sync the clock (NTP).

## Config changes don't take effect

Restart the gateway after editing config:

```bash
openclaw gateway restart
```

Also confirm the config is under `plugins.entries."msteams-voice".config` and that keys match the
schema exactly - unknown keys are rejected (`additionalProperties: false`).

## The agent answers but stays silent / no audio

- **Recording gate.** With `requireRecordingStatus: true`, the agent holds media until Teams reports
  recording active. Start recording, or set it to `false` for testing.
- **No realtime provider resolved.** If you intended realtime but no provider key resolves, the runtime
  falls back to streaming. Check `realtime.provider` and the provider `apiKey`.

## Calls are declined

- **Allowlist is closed by default.** With `inboundPolicy: "allowlist"` and an empty `allowFrom`, every
  caller is denied. Add the caller's AAD object id to `allowFrom`, or use a different policy.
- `pairing` behaves like `allowlist` today - callers still must be in `allowFrom`.

## The agent interrupts itself / barge-in feels off

- Tune the echo guard: raise `realtime.echoBargeInRms` if it interrupts itself; lower it if barge-in is
  unresponsive. `suppressInputDuringPlayback` and `echoSuppressionWindowMs` also help.

## The call ends early with a goodbye

That is the tier limit. The **sandbox** is about 5 minutes/day per session and the **free** tier is
5 minutes/day; when the limit is hit the agent speaks a goodbye and ends. Move to a subscription for
uncapped calls. See [Connecting to StandIn](/openclaw-msteams-voice/connecting-to-standin/).

## Outbound never connects

- `outbound.enabled` must be true and `outbound.workerBaseUrl` set.
- The place-call request is signed with `sharedSecret`; a mismatch fails it.
- Check `tenantId` is the callee's tenant and `answerTimeoutMs` is not too short.

## Still stuck?

Open an issue on [GitHub](https://github.com/komaa-com/openclaw-msteams-voice/issues) with your
(secret-redacted) config, the mode, and gateway logs around the failed call.
