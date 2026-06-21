// @alaamh/msteams-voice — plugin entry.
//
// 🚧 SCAFFOLD. Confirm the exact plugin-entry contract for a non-channel runtime plugin (DESIGN.md
// open item) — `definePluginEntry` is re-exported from openclaw/plugin-sdk/talk-voice.
//
// On startup this wires the self-contained Teams voice runtime:
//   config (api.pluginConfig) -> CallLifecycle (api.runtime.state) -> Teams media WS server
//   -> per call: createRealtimeVoiceBridgeSession + consult via api.runtime.agent.

import { definePluginEntry, type OpenClawPluginApi } from "openclaw/plugin-sdk/talk-voice";
// import { CallLifecycle } from "./call-lifecycle.js";
// import { bindTeamsCallToRealtime } from "./runtime-bridge.js";

export default definePluginEntry({
  id: "msteams-voice",
  registerFull(api: OpenClawPluginApi) {
    // TODO:
    //  1. const cfg = api.pluginConfig;  if (!cfg?.enabled) return;
    //  2. const lifecycle = new CallLifecycle(
    //       { openSyncKeyedStore: (n) => api.runtime.state.openSyncKeyedStore(n),
    //         log: api.runtime.logging.getChildLogger({ plugin: "msteams-voice" }),
    //         now: () => Date.now() },
    //       { maxConcurrentCalls, maxDurationMs, staleCallReaperMs });
    //     lifecycle.start();
    //  3. start the Teams media WS server (msteams-media-stream) on cfg.port/path with cfg.sharedSecret
    //  4. on session.start  -> lifecycle.initiate(...) + bindTeamsCallToRealtime({ session, realtimeProvider, api })
    //     on session.end    -> lifecycle.end(...)
    //  5. register teardown (stop WS server + lifecycle.stop()).
    void api;
  },
});
