// @komaa/openclaw-msteams-bridge — plugin entry (self-contained Teams CVI voice).
//
// Registers a host-managed background service so the runtime's lifecycle (start on boot, stop on
// shutdown/reload) is wired by OpenClaw — this is the teardown hook. On start the service brings up
// the MsteamsVoiceRuntime (Teams media WS server, CallLifecycle, per-call bridge); on stop it tears
// it all down (closes calls, stops the lifecycle reaper, closes the WS server).

import { definePluginEntry } from "openclaw/plugin-sdk/core";
import { MsteamsVoiceRuntime } from "./msteams-runtime.js";
import { resolvePluginConfig } from "./plugin-config.js";

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
