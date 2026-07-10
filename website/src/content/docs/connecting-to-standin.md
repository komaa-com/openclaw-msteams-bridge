---
title: "Connecting to StandIn"
description: "The connection model, the sandbox / free / subscription tiers, and where the shared secret comes from."
---

StandIn is the hosted media bridge that joins the Microsoft Teams call and connects it to your agent.
This page explains the connection model and the three tiers you can connect through.

## The connection model

Your plugin runs a **WebSocket server** inside the OpenClaw gateway. StandIn is the **client**: for
each call it opens an authenticated WebSocket to your plugin and streams audio/video over it. You do
not dial out and you never handle Teams media yourself.

```
Teams call <-> StandIn media bridge ==WebSocket (StandIn dials in)==> @komaa/msteams-voice
```

What you configure to make this work:

| Setting | Why |
|---|---|
| `bindAddress` | Set to `0.0.0.0` so the hosted bridge can reach the plugin. Default `127.0.0.1` is local-only. |
| `port` | The port the plugin listens on (default `9442`). Make it reachable by StandIn. |
| `path` | The WebSocket route (default `/voice/msteams/stream`). |
| `sharedSecret` | The HMAC secret. It **must match** the value StandIn uses, or the handshake is rejected. |

Every connection is authenticated with a replay-proof HMAC handshake - see the [Wire Protocol](/openclaw-msteams-voice/wire-protocol/).
This is the same regardless of which tier you use; only the identity and limits differ.

## The three tiers

Pick a tier by where you are in your journey. From the plugin's side the mechanics are identical - you
always get a shared secret and StandIn dials in.

### Sandbox - try it in minutes

- **What it is:** a shared StandIn bot that joins a Teams meeting you generate. **No Azure/Teams bot
  of your own required.**
- **Limits:** time-limited (about 5 minutes/day per session).
- **Use it for:** your first call and quick experiments.
- **Start:** [standin.komaa.com/sandbox](https://standin.komaa.com/sandbox) - it walks you through
  generating a meeting link and gives you the shared secret to paste into `sharedSecret`.

### Free - develop with your own bot

- **What it is:** the developer tier. You **bring your own Microsoft Teams bot** (an Azure Bot) and
  pair it in the StandIn dashboard. Pairing issues the shared secret for that identity.
- **Limits:** daily-capped (5 minutes/day), on its own slot.
- **Use it for:** building and testing against a real inbound number/identity you control.
- **Start:** [standin.komaa.com](https://standin.komaa.com) - create an account and pair your bot.

### Subscription - production

- **What it is:** your own Teams bot with no daily cap, managed in the StandIn dashboard.
- **Use it for:** real, always-on deployments.
- **Start:** [standin.komaa.com](https://standin.komaa.com).

> For account, dashboard, and bot-pairing specifics (creating the Azure Bot, entering credentials,
> retrieving the secret), follow the StandIn docs at [docs.komaa.com](https://docs.komaa.com). Those
> steps live on the StandIn side; this plugin only needs the resulting `sharedSecret`.

## Pairing your own Teams bot

When you move past the sandbox, you register your Teams bot with OpenClaw's Teams **chat** channel and
pair it with StandIn for **voice**. The Teams bot credentials are supplied to OpenClaw through the
channel environment variables:

| Env var | Meaning |
|---|---|
| `MSTEAMS_APP_ID` | Your Teams bot (Azure Bot) app id |
| `MSTEAMS_APP_PASSWORD` | The bot app secret |
| `MSTEAMS_TENANT_ID` | Your Microsoft Entra (AAD) tenant id |

Pairing in the StandIn dashboard links that bot identity to a shared secret; put the secret in
`sharedSecret`. From then on, inbound calls to your bot are bridged to your plugin.

## Restricting who can reach the agent

Inbound calls are gated by policy - closed by default:

- `inboundPolicy`: `disabled` | `allowlist` | `pairing` | `open`.
- `allowFrom`: the list of allowed caller AAD object ids.

> Note: `pairing` currently behaves exactly like `allowlist` - the plugin issues no per-call pairing
> codes; callers must be present in `allowFrom`.

See [Configuration Reference](/openclaw-msteams-voice/configuration-reference/) for details.

## Cutoff and the spoken goodbye

When a tier limit is reached mid-call (sandbox time cap, free daily budget, or a paid max-minutes
governor), StandIn sends an `assistant.say` message with a short goodbye line. The agent speaks it,
then the call ends gracefully rather than cutting off abruptly.
