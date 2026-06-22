// @alaamh/msteams-voice — plugin entry.
//
// 🚧 SCAFFOLD (wired in Phase 3). The exact `definePluginEntry` contract (DefinePluginEntryOptions)
// + runtime wiring are Phase-3 work — kept as notes here so Phase 2 stays a clean, building checkpoint.
//
// Phase 3 plan (in the entry's register/registerFull):
//   1. const cfg = api.pluginConfig;  if (!cfg?.enabled) return;
//   2. const lifecycle = new CallLifecycle(
//        { openSyncKeyedStore: (n) => api.runtime.state.openSyncKeyedStore(n),
//          log: api.runtime.logging.getChildLogger({ plugin: "msteams-voice" }),
//          now: () => Date.now() },
//        { maxConcurrentCalls, maxDurationMs, staleCallReaperMs });
//      lifecycle.start();
//   3. start the Teams media WS server (msteams-media-stream) on cfg.port/path with cfg.sharedSecret
//   4. on session.start -> lifecycle.initiate(...) + bindTeamsCallToRealtime({ session, realtimeProvider, api })
//      on session.end   -> lifecycle.end(...)
//   5. register teardown (stop WS server + lifecycle.stop()).
//
// Realtime bridge + TTS (Phase 2) are ported and tested: ./msteams-realtime.ts, ./msteams-tts.ts,
// ./msteams-tts-playback.ts (consume api.runtime.tts via ./telephony-tts.ts).

export {};
