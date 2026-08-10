import { definePluginEntry } from "openclaw/plugin-sdk/core";
import { MsteamsVoiceRuntime } from "./msteams-runtime.js";
import { resolvePluginConfig } from "./plugin-config.js";
export default definePluginEntry({
    id: "msteams-call",
    name: "Teams Call by StandIn",
    description: "Microsoft Teams calls and chat for your OpenClaw agent, through StandIn.",
    register(api) {
        const cfg = resolvePluginConfig(api.pluginConfig);
        if (!cfg.enabled)
            return;
        const logger = api.runtime.logging.getChildLogger({ plugin: "msteams-call" });
        if (!cfg.media.sharedSecret && !cfg.managedChat.enabled) {
            logger.warn("msteams-call: no secret configured - set `secret` to the value StandIn shows you (it covers " +
                "both calling and messages). Nothing started.");
            return;
        }
        if (!cfg.media.sharedSecret) {
            logger.warn("msteams-call: no calling secret - the messages lane starts, calls will not be answered");
        }
        let runtime;
        api.registerService({
            id: "msteams-call",
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
