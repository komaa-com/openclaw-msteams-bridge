---
title: "Contributing"
description: "How to set up a dev environment and contribute to the plugin."
---

Contributions are welcome. The repo's [CONTRIBUTING.md](https://github.com/komaa-com/openclaw-msteams-bridge/blob/main/CONTRIBUTING.md)
is the authoritative guide; this page is the quick orientation.

## Setup

```bash
git clone https://github.com/komaa-com/openclaw-msteams-bridge
cd openclaw-msteams-bridge
npm ci
npm run build       # tsc -> dist/
npm run typecheck
npm test            # vitest
```

The package is TypeScript under `src/`, compiled to `dist/`, and published **prebuilt** - commit
source under `src/`, never hand-edit `dist/`.

## Where things live

| Area | Files |
|---|---|
| Entry / manifest | `src/index.ts`, `openclaw.plugin.json` |
| Config | `src/config.ts`, `src/plugin-config.ts` |
| Transport + handshake | `src/msteams-media-stream.ts`, `src/protocol.gen.ts` |
| Runtime / mode select | `src/msteams-runtime.ts` |
| Call state machine | `src/call-lifecycle.ts` |
| Realtime | `src/msteams-realtime.ts`, `src/msteams-realtime-tools.ts` |
| Streaming | `src/msteams-streaming.ts`, `src/msteams-tts*.ts`, `src/telephony-*.ts` |
| Vision | `src/vision-*.ts`, `src/msteams-video-frame.ts` |
| Group gate | `src/group-call-gate.ts` |
| Avatar cues | `src/expression.ts`, `src/viseme-estimate.ts` |

See [DESIGN.md](https://github.com/komaa-com/openclaw-msteams-bridge/blob/main/DESIGN.md) for the
architecture.

## Conventions

- Branch from `main` (`feat/…`, `fix/…`, `docs/…`, `ci/…`); keep PRs focused.
- Keep the config surface in sync across `config.ts`, `plugin-config.ts`, and the `configSchema`.
- `build`, `typecheck`, and `test` must pass. A real call through the sandbox is the best verification.

## Documentation leak policy

The StandIn media bridge is a hosted service and its internals are not public. In any doc, comment, or
example, describe only the client side and the wire protocol; refer to the counterpart as "the StandIn
media bridge" and never describe how it produces Teams media or any internal component behind it.
