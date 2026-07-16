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
            stop: async () => {
                await runtime?.stop();
                runtime = undefined;
            },
        });
    },
});
