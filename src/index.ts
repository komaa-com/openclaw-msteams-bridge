// @komaa/openclaw-msteams-bridge — plugin entry (self-contained Teams CVI voice).
//
// Registers a host-managed background service so the runtime's lifecycle (start on boot, stop on
// shutdown/reload) is wired by OpenClaw — this is the teardown hook. On start the service brings up
// the MsteamsVoiceRuntime (Teams media WS server, CallLifecycle, per-call bridge); on stop it tears
// it all down (closes calls, stops the lifecycle reaper, closes the WS server).

import { definePluginEntry } from "openclaw/plugin-sdk/core";
import { MsteamsVoiceRuntime } from "./msteams-runtime.js";
import { resolvePluginConfig } from "./plugin-config.js";
import { MSTEAMS_POST_CHAT_TOOL_NAME } from "./msteams-realtime-tools.js";

export default definePluginEntry({
  id: "msteams-call",
  name: "Teams Call by StandIn",
  description: "Microsoft Teams calls and chat for your OpenClaw agent, through StandIn.",
  register(api) {
    const cfg = resolvePluginConfig((api as { pluginConfig?: unknown }).pluginConfig);
    if (!cfg.enabled) return;

    const logger = api.runtime.logging.getChildLogger({ plugin: "msteams-call" });
    // Either lane is reason enough to start. This used to return unless the CALLING secret was set,
    // so a Managed Bot config - which sets one secret and expects both lanes - configured chat
    // correctly and then started nothing at all, silently. The runtime brings up whichever lanes are
    // configured; refusing here is only for a plugin that is enabled with no secret anywhere.
    if (!cfg.media.sharedSecret && !cfg.managedChat.enabled) {
      logger.warn(
        "msteams-call: no secret configured - set `secret` to the value StandIn shows you (it covers " +
          "both calling and messages). Nothing started.",
      );
      return;
    }
    if (!cfg.media.sharedSecret) {
      logger.warn("msteams-call: no calling secret - the messages lane starts, calls will not be answered");
    }

    let runtime: MsteamsVoiceRuntime | undefined;

    // The in-call "post to the Teams chat" tool, for the STREAMING path.
    //
    // Realtime gets this injected per call, with exact context. Streaming has no tool-dispatch
    // surface - STT -> consult -> TTS - so the only way its delegated agent turn can reach a tool is
    // an OpenClaw tool, and those are global. The handler context carries no session id, so the tool
    // asks the runtime which call is postable and REFUSES when that is ambiguous rather than posting
    // into the wrong conversation. Registered only when the messages lane is configured, so a
    // voice-only deployment never sees a tool it cannot honour.
    if (cfg.managedChat.enabled) {
      api.registerTool({
        name: "post_chat_message",
        label: "Post to the Teams chat",
        description:
          "Post a text message into the Teams chat for the call in progress. Use when the caller asks " +
          'you to "send that to the chat", "post it", or "message me the link" - anything they want to ' +
          "keep after the call ends.",
        promptSnippet: "post_chat_message: put text in the Teams chat of the call in progress",
        parameters: {
          type: "object",
          properties: { text: { type: "string", description: "The message to post. Markdown is supported." } },
          required: ["text"],
        } as never,
        async execute(_toolCallId: string, params: unknown) {
          const text = String((params as { text?: unknown })?.text ?? "").trim();
          if (!text) {
            return { content: [{ type: "text", text: "There was nothing to post." }], isError: true } as never;
          }
          const target = runtime?.resolvePostableCall();
          if (!target || "error" in target) {
            const reason = target?.error ?? "The Teams bridge is not running.";
            return { content: [{ type: "text", text: reason }], isError: true } as never;
          }
          const ok = await target.post(text);
          return {
            content: [{ type: "text", text: ok ? "Posted to the Teams chat." : "Could not post to the Teams chat." }],
            isError: !ok,
          } as never;
        },
      } as never);
    }

    // Chat-to-call: ask in a Teams chat, get the answer as a phone call.
    //
    // Registered on the same messages-lane condition as post_chat_message, because the only target it will
    // ever ring is the authenticated sender of a chat we just answered - no target parameter, by design.
    // The agent reads untrusted text all day; a tool that accepted an arbitrary AAD id would turn any of
    // that into "call this person". Here the worst a prompt injection achieves is calling the person who
    // is already in the conversation.
    if (cfg.managedChat.enabled) {
      api.registerTool({
        name: "call_me_with_the_answer",
        label: "Call back with the answer",
        description:
          "Ring the person you are chatting with on Microsoft Teams and speak an answer to them, instead " +
          'of replying in the chat. Use when they ask to be CALLED - "call me with the answer", "ring me ' +
          'when you know", "tell me by phone". The message is spoken aloud to someone who may not ' +
          "remember asking, so make it a complete, self-contained sentence or two - restate the topic and " +
          "give the answer. Do not use this for a normal reply; just answer in the chat for those.",
        promptSnippet:
          "call_me_with_the_answer: phone the Teams chat sender and speak an answer, when they asked to be called",
        parameters: {
          type: "object",
          properties: {
            message: {
              type: "string",
              description:
                "What to say when they answer. Plain spoken language, no markdown. Self-contained: " +
                'restate the topic and give the answer, e.g. "About the Dubai time you asked for - it is 5:41 PM."',
            },
          },
          required: ["message"],
        } as never,
        async execute(_toolCallId: string, params: unknown) {
          const message = String((params as { message?: unknown })?.message ?? "").trim();
          if (!message) {
            // Placing a call with nothing to say is worse than not calling: they answer to silence.
            return {
              content: [{ type: "text", text: "There was nothing to say, so I did not call." }],
              isError: true,
            } as never;
          }
          const target = runtime?.resolveChatCallbackTarget();
          if (!target || "error" in target) {
            const reason = target?.error ?? "The Teams bridge is not running.";
            return { content: [{ type: "text", text: reason }], isError: true } as never;
          }
          try {
            const placed = await runtime!.placeCall(target.to, { message, mode: "notify" });
            return {
              content: [
                {
                  type: "text",
                  text: `Calling ${target.displayName ?? "them"} now to deliver that (call ${placed.callId}).`,
                },
              ],
            } as never;
          } catch (err) {
            // Surface the real reason - placeCall's errors name the missing config precisely, and a vague
            // "could not call" here is what sends someone hunting through the wrong settings.
            const why = err instanceof Error ? err.message : String(err);
            return {
              content: [{ type: "text", text: `I could not place the call: ${why}` }],
              isError: true,
            } as never;
          }
        },
      } as never);
    }

    api.registerService({
      id: "msteams-call",
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
