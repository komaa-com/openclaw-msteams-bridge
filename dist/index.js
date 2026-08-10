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
        if (cfg.managedChat.enabled) {
            api.registerTool({
                name: "post_chat_message",
                label: "Post to the Teams chat",
                description: "Post a text message into the Teams chat for the call in progress. Use when the caller asks " +
                    'you to "send that to the chat", "post it", or "message me the link" - anything they want to ' +
                    "keep after the call ends.",
                promptSnippet: "post_chat_message: put text in the Teams chat of the call in progress",
                parameters: {
                    type: "object",
                    properties: { text: { type: "string", description: "The message to post. Markdown is supported." } },
                    required: ["text"],
                },
                async execute(_toolCallId, params) {
                    const text = String(params?.text ?? "").trim();
                    if (!text) {
                        return { content: [{ type: "text", text: "There was nothing to post." }], isError: true };
                    }
                    const target = runtime?.resolvePostableCall();
                    if (!target || "error" in target) {
                        const reason = target?.error ?? "The Teams bridge is not running.";
                        return { content: [{ type: "text", text: reason }], isError: true };
                    }
                    const ok = await target.post(text);
                    return {
                        content: [{ type: "text", text: ok ? "Posted to the Teams chat." : "Could not post to the Teams chat." }],
                        isError: !ok,
                    };
                },
            });
        }
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
