// @alaamh/msteams-voice — plugin entry (self-contained Teams CVI voice).
//
// Registers an on-startup runtime that owns the call lifecycle. The Teams media WS server + realtime
// bridge wiring is the remaining Phase-3 step (TODO below) — kept out so this stays a green checkpoint.

import { definePluginEntry } from "openclaw/plugin-sdk/core";
import { CallLifecycle, type LifecycleRuntime, type SyncKeyedStore } from "./call-lifecycle.js";

const STORE_MAX_ENTRIES = 2000;

interface MsteamsVoicePluginConfig {
  enabled?: boolean;
  maxConcurrentCalls?: number;
  maxDurationSeconds?: number;
  staleCallReaperSeconds?: number;
}

export default definePluginEntry({
  id: "msteams-voice",
  name: "Microsoft Teams Voice (CVI)",
  description: "Self-contained Microsoft Teams realtime voice agent (CVI) for OpenClaw.",
  register(api) {
    // api.pluginConfig is the plugin's own validated config slice; accessed defensively so the entry
    // compiles regardless of the exact api typing surfaced by the published SDK.
    const cfg = (api as { pluginConfig?: MsteamsVoicePluginConfig }).pluginConfig;
    if (cfg?.enabled === false) return;

    const logger = api.runtime.logging.getChildLogger({ plugin: "msteams-voice" });

    const rt: LifecycleRuntime = {
      // Adapt api.runtime.state's sync keyed store (register/lookup/delete/entries) to our
      // get/set/delete/keys surface. (plugin-state-store.types.ts)
      openSyncKeyedStore: <T>(name: string): SyncKeyedStore<T> => {
        const s = api.runtime.state.openSyncKeyedStore<T>({
          namespace: name,
          maxEntries: STORE_MAX_ENTRIES,
        });
        return {
          get: (k) => s.lookup(k),
          set: (k, v) => s.register(k, v),
          delete: (k) => {
            s.delete(k);
          },
          keys: () => s.entries().map((e) => e.key),
        };
      },
      log: {
        info: (m) => logger.info(m),
        warn: (m) => logger.warn(m),
        error: (m) => logger.error(m),
      },
      now: () => Date.now(),
    };

    const lifecycle = new CallLifecycle(rt, {
      maxConcurrentCalls: cfg?.maxConcurrentCalls ?? 5,
      maxDurationMs: (cfg?.maxDurationSeconds ?? 0) * 1000,
      staleCallReaperMs: (cfg?.staleCallReaperSeconds ?? 0) * 1000,
    });
    lifecycle.start();
    logger.info("msteams-voice: call lifecycle started");

    // TODO (Phase 3 final — provider orchestration):
    //  - start the Teams media WS server (./msteams-media-stream) on cfg.port/path with cfg.sharedSecret
    //  - resolve the realtime voice provider via resolveConfiguredRealtimeVoiceProvider
    //    (openclaw/plugin-sdk/realtime-voice)
    //  - on session.start  -> lifecycle.initiate(...) + bind ./msteams-realtime (createRealtimeVoiceBridgeSession),
    //                         route consult to consultRealtimeVoiceAgent / api.runtime.agent
    //  - on session.end    -> lifecycle.end(...)
    //  - register teardown: stop the WS server + lifecycle.stop()
  },
});
