import { definePluginEntry } from "openclaw/plugin-sdk/core";
import { MsteamsBridgeRuntime } from "./msteams-runtime.js";
import { resolvePluginConfig } from "./plugin-config.js";
export default definePluginEntry({
    id: "msteams-bridge",
    name: "Teams Bridge by StandIn",
    description: "Microsoft Teams calls and chat for your OpenClaw agent, through StandIn.",
    register(api) {
        const cfg = resolvePluginConfig(api.pluginConfig);
        if (!cfg.enabled)
            return;
        const logger = api.runtime.logging.getChildLogger({ plugin: "msteams-bridge" });
        if (!cfg.media.sharedSecret && !cfg.managedChat.enabled) {
            logger.warn("msteams-bridge: no secret configured - set `secret` to the value StandIn shows you (it covers " +
                "both calling and messages). Nothing started.");
            return;
        }
        if (!cfg.media.sharedSecret) {
            logger.warn("msteams-bridge: no calling secret - the messages lane starts, calls will not be answered");
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
        if (cfg.managedChat.enabled) {
            api.registerTool({
                name: "call_me_with_the_answer",
                label: "Call back with the answer",
                description: "Ring the person you are chatting with on Microsoft Teams and speak an answer to them, instead " +
                    'of replying in the chat. Use when they ask to be CALLED - "call me with the answer", "ring me ' +
                    'when you know", "tell me by phone". The message is spoken aloud to someone who may not ' +
                    "remember asking, so make it a complete, self-contained sentence or two - restate the topic and " +
                    "give the answer. Do not use this for a normal reply; just answer in the chat for those.",
                promptSnippet: "call_me_with_the_answer: phone the Teams chat sender and speak an answer, when they asked to be called",
                parameters: {
                    type: "object",
                    properties: {
                        message: {
                            type: "string",
                            description: "What to say when they answer. Plain spoken language, no markdown. Self-contained: " +
                                'restate the topic and give the answer, e.g. "About the Dubai time you asked for - it is 5:41 PM."',
                        },
                    },
                    required: ["message"],
                },
                async execute(_toolCallId, params) {
                    const message = String(params?.message ?? "").trim();
                    if (!message) {
                        return {
                            content: [{ type: "text", text: "There was nothing to say, so I did not call." }],
                            isError: true,
                        };
                    }
                    const target = runtime?.resolveChatCallbackTarget();
                    if (!target || "error" in target) {
                        const reason = target?.error ?? "The Teams bridge is not running.";
                        return { content: [{ type: "text", text: reason }], isError: true };
                    }
                    try {
                        const placed = await runtime.placeCall(target.to, {
                            message,
                            mode: "notify",
                            fallback: target.fallback,
                        });
                        return {
                            content: [
                                {
                                    type: "text",
                                    text: `Calling ${target.displayName ?? "them"} now to deliver that (call ${placed.callId}).`,
                                },
                            ],
                        };
                    }
                    catch (err) {
                        const why = err instanceof Error ? err.message : String(err);
                        return {
                            content: [{ type: "text", text: `I could not place the call: ${why}` }],
                            isError: true,
                        };
                    }
                },
            });
        }
        api.registerService({
            id: "msteams-bridge",
            start: async () => {
                runtime = new MsteamsBridgeRuntime(api, cfg);
                await runtime.start();
            },
            stop: async () => {
                await runtime?.stop();
                runtime = undefined;
            },
        });
    },
});
