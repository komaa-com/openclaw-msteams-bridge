# Contributing

Thanks for your interest in improving `@komaa/openclaw-msteams-bridge`. This guide covers local setup, the
conventions we follow, and how releases work.

## Prerequisites

- Node.js (a current LTS) and npm.
- For running a real call end to end: an OpenClaw install (`>= 2026.6.10`) and a StandIn connection
  (the [sandbox](https://standin.komaa.com/sandbox) is enough).

## Local setup

```bash
git clone https://github.com/komaa-com/openclaw-msteams-bridge
cd openclaw-msteams-bridge
npm ci
npm run build       # tsc -> dist/
npm run typecheck   # type-only pass
npm test            # vitest
```

The package is authored in TypeScript under `src/` and compiled to `dist/`. The published package
ships **prebuilt**: `files` is `["dist", "openclaw.plugin.json"]`, and `prepublishOnly` runs the
build, so consumers install without a build step. Always commit source changes under `src/`; do not
hand-edit `dist/` (it is regenerated).

## Working on it

- Point a local OpenClaw gateway at your working copy and connect it to the StandIn sandbox to
  exercise a real call. See [Getting Started](https://komaa-com.github.io/openclaw-msteams-bridge/getting-started/).
- Keep the config surface in sync across `src/config.ts`, `src/plugin-config.ts`, and the
  `configSchema` in `openclaw.plugin.json` - these three must agree, and the schema is
  `additionalProperties: false`, so a new option must be added to the schema or it is rejected.
- Add or update `vitest` tests alongside behavior changes.

## Branches and pull requests

- Branch from `main`; use a short prefixed name, e.g. `feat/…`, `fix/…`, `docs/…`, `ci/…`.
- Keep PRs focused. Describe the change and how you verified it (a real call is the gold standard).
- `npm run build`, `npm run typecheck`, and `npm test` must pass before review.

## Releases

Publishing to npm is automated (`.github/workflows/publish.yml`): bump the version, tag it / cut a
GitHub Release, and CI runs `npm publish --access public --provenance` (supply-chain attestation).
Keep the `version` in `package.json` and any version references in the docs consistent.

## Documentation and the leak policy

The StandIn media bridge is a hosted service; **its internal implementation is not public**. When you
write docs, comments, or examples, describe only the client side and the wire contract. Refer to the
counterpart as "the StandIn media bridge" and never document how it produces Teams media, what it runs
on, or any internal component, source file, or codegen behind it. If you need bridge behavior to
explain something, describe it through the observable wire protocol only.
