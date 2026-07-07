// @komaa/msteams-voice — plugin entry (self-contained Teams CVI voice).
//
// Registers a host-managed background service so the runtime's lifecycle (start on boot, stop on
// shutdown/reload) is wired by OpenClaw — this is the teardown hook. On start the service brings up
// the MsteamsVoiceRuntime (Teams media WS server, CallLifecycle, per-call bridge); on stop it tears
// it all down (closes calls, stops the lifecycle reaper, closes the WS server).
import { definePluginEntry } from "openclaw/plugin-sdk/core";
import { MsteamsVoiceRuntime } from "./msteams-runtime.js";
import { resolvePluginConfig } from "./plugin-config.js";
export default definePluginEntry({
    id: "msteams-voice",
    name: "Microsoft Teams Voice (CVI)",
    description: "Self-contained Microsoft Teams realtime voice agent (CVI) for OpenClaw.",
    register(api) {
        const cfg = resolvePluginConfig(api.pluginConfig);
        if (!cfg.enabled)
            return;
        const logger = api.runtime.logging.getChildLogger({ plugin: "msteams-voice" });
        if (!cfg.media.sharedSecret) {
            logger.warn("msteams-voice: sharedSecret is not configured — media server not started");
            return;
        }
        let runtime;
        api.registerService({
            id: "msteams-voice",
            start: async () => {
                runtime = new MsteamsVoiceRuntime(api, cfg);
                await runtime.start();
            },
            // Teardown: host calls stop() on shutdown/reload → close calls, stop reaper, close WS server.
            stop: async () => {
                await runtime?.stop();
                runtime = undefined;
            },
        });
    },
});
