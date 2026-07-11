---
title: "Troubleshooting"
description: "Common problems and fixes: handshake rejections, unreachable plugin, silent calls, allowlist blocks, provider errors."
---

Common problems and how to fix them. The gateway log is your first stop - the plugin logs the
handshake result, session lifecycle, and provider errors for every call.

## The bridge can't connect / handshake rejected

**Symptom:** StandIn reports it cannot reach or authenticate with your agent; no `session.start`
ever appears in the gateway log.

**Causes & fixes:**

- **Secret mismatch** (most common) - `sharedSecret` in your plugin config does not equal the value
  StandIn holds. They must be byte-for-byte identical. Re-copy it from the StandIn sandbox page or
  dashboard.
- **Bind address** - the default `bindAddress` is `127.0.0.1`, which only accepts local
  connections. Set `bindAddress: "0.0.0.0"` so the hosted bridge can reach the plugin.
- **Port not reachable** - confirm `port` (default `9442`) is open to StandIn (firewall/NAT) and
  not taken by another process.
- **Clock skew** - the handshake enforces a 60 s replay window on the signed timestamp. If the host
  clock is far off, every handshake is rejected. Sync time (NTP).
- **No secret at all** - the plugin fails closed: with an empty or non-string `sharedSecret`, it
  rejects every handshake. Set a real secret.

## The plugin isn't loading

**Symptom:** the gateway starts but the voice endpoint is not listening; no plugin banner in the
log.

- Confirm the plugin is installed: `openclaw plugins list` should show `msteams-voice`.
  Reinstall with `openclaw plugins install npm:@komaa/openclaw-msteams-bridge` if not.
- Confirm `enabled` is not set to `false` under `plugins.entries."msteams-voice".config`.
- Restart the gateway after any install or config change: `openclaw gateway restart`.

## Config changes don't take effect

Restart the gateway after editing config:

```bash
openclaw gateway restart
```

Also confirm the config is under `plugins.entries."msteams-voice".config` and that keys match the
schema exactly - unknown keys are rejected (`additionalProperties: false`), so a typo makes the
whole config invalid rather than being silently ignored.

## The agent answers but stays silent / no audio

- **Recording gate** - with `requireRecordingStatus: true` (default), nothing is processed until
  Teams reports recording `active`. If the meeting is not being recorded, the agent stays silent.
  Start recording, or (for testing only) set `requireRecordingStatus: false`.
- **No realtime provider resolved** - if you intended realtime but no provider key resolves, the
  runtime falls back to streaming, which needs your OpenClaw STT/TTS configured. Check
  `realtime.provider` and the provider `apiKey`.
- **Realtime provider errors** - an invalid key, missing model access, or a wrong
  `azureEndpoint`/`azureDeployment` shows up as a provider connect error in the log. Verify the key
  and model/deployment name.
- **Group gate** - in a meeting (2+ people) the agent only speaks when **addressed** by a wake
  phrase (`groupCall.wakePhrases`). Say its name, or disable `groupCall.requireAddress`.

## Calls are declined (`not-allowed` in the log)

The inbound gate is **deny-by-default**:

- With `inboundPolicy` **unset** (the default) or `disabled`, **every** inbound call is denied -
  including your first sandbox call. Set a policy to receive calls.
- With `inboundPolicy: "allowlist"` and an **empty** `allowFrom`, every caller is denied too.
- Callers are matched by **AAD object id** (case-insensitive) or **phone number** (digits only) -
  add the caller to `allowFrom`.
- `pairing` currently behaves exactly like `allowlist` - callers still must be in `allowFrom`.
- `open` accepts any caller (use for sandbox testing only).

The gateway log's rejection line names the policy and the caller id it saw, so you can paste that
id straight into `allowFrom`.

## The agent interrupts itself / barge-in feels off

- If the agent keeps cutting itself off (hearing its own voice as a barge-in), raise
  `realtime.echoBargeInRms` and/or enable `realtime.suppressInputDuringPlayback`.
- If real barge-in feels unresponsive, lower `realtime.echoBargeInRms` or shorten
  `realtime.echoSuppressionWindowMs`.

## The call drops itself mid-conversation

- **Stale-call reaper** - a call that stops being serviced is torn down after
  `staleCallReaperSeconds` (default 120). If long silent periods are expected, raise it.
- **Duration cap** - if `maxDurationSeconds` is set, the call is closed once it exceeds that
  wall-clock budget. Raise it or leave it unset (unlimited).
- **Concurrency cap** - beyond `maxConcurrentCalls` (default 4), additional calls are declined.

## The call ends with a goodbye after a few minutes

That is the **StandIn tier cutoff**, not a bug. The **sandbox** and **free** tiers are daily-capped
(about 5 minutes/day); a **subscription** may have a max-minutes governor. StandIn sends an
`assistant.say` goodbye that the agent speaks, then the call ends gracefully. For longer calls,
move to a subscription tier - see
[Connecting to StandIn](/openclaw-msteams-bridge/connecting-to-standin/).

## Outbound never connects

- `outbound.enabled` must be `true` and `outbound.workerBaseUrl` set.
- The place-call request is signed with `sharedSecret`; a mismatch fails it.
- Check `outbound.tenantId` is the callee's AAD tenant.
- **No answer:** after `outbound.answerTimeoutMs` (default 120,000 ms) the plugin treats the call
  as unanswered, delivers the voicemail-style fallback if configured, and cancels the ringing call
  so the callee's Teams stops ringing. A late answer after that point is declined by design.

## Where to look for logs

Everything the plugin does is logged through the OpenClaw gateway log: handshake accept/reject
reasons, `session.start`/`session.end`, recording-status changes, provider connects and errors, and
teardown reasons. Watch it live while you place a test call.

## Still stuck?

Open an issue on [GitHub](https://github.com/komaa-com/openclaw-msteams-bridge/issues) with the mode
you used (`realtime`/`streaming`), the gateway log around the failed call, and your
(secret-redacted) config. Hosted-service questions (account, pairing, dashboard) belong at
[docs.komaa.com](https://docs.komaa.com).
