# @alaamh/msteams-voice

**Self-contained Microsoft Teams CVI realtime-voice plugin for OpenClaw.**

An AI assistant that joins Microsoft Teams calls as a real participant — speech-to-speech, continuous
vision, "speak-only-when-addressed" gating, outbound call-backs with voicemail, avatar lip-sync,
meeting recap.

Unlike the earlier two-package approach (`@alaamh/voice-call` fork + provider), this is **one plugin**
that depends only on the **published `openclaw` plugin-sdk** and `api.runtime` — **no fork**, no
`CallManager`/config vendoring. Install like any third-party plugin:

```
openclaw plugins install clawhub:@alaamh/msteams-voice
cd extensions/msteams-voice && pnpm install && pnpm build
```

> 🚧 **Status: SCAFFOLD.** Structure + manifests + the call-lifecycle skeleton are in place. The CVI
> code (media-stream, realtime bridge wiring, vision, gate, tools, recap) is ported in next — see
> [PORTING.md](PORTING.md). The original implementation lives in `alaamh/openclaw-voice-call` (kept as
> the fallback).

## Why this is small
The hard parts come from OpenClaw itself (verified — see [DESIGN.md](DESIGN.md)):
- **Realtime audio bridge** (audio in/out, barge-in, vision) → `openclaw/plugin-sdk/realtime-voice`
  (`createRealtimeVoiceBridgeSession`, `consultRealtimeVoiceAgent`).
- **Agent / TTS / STT / media / config / state / logging** → `api.runtime`.
- **We write only** a ~500-LOC Teams **call-lifecycle** coordinator (`src/call-lifecycle.ts`).

## Layout
```
.
├─ openclaw.plugin.json     # plugin id "msteams-voice" + Teams-only configSchema
├─ DESIGN.md                # architecture, api.runtime map, write-it-yourself list
├─ PORTING.md               # which CVI files to bring from alaamh/openclaw-voice-call
└─ src/
   ├─ index.ts              # plugin entry (definePluginEntry) — wires runtime + bridge
   ├─ runtime-bridge.ts     # createRealtimeVoiceBridgeSession + api.runtime wiring (stub)
   ├─ call-lifecycle.ts     # the ONLY substantial hand-written glue (skeleton)
   ├─ types.ts              # CallState / CallRecord
   └─ (CVI files dropped in here: msteams-media-stream, msteams-realtime, vision, gate, tools, …)
```
