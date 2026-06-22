// @alaamh/msteams-voice — plugin entry (self-contained Teams CVI voice).
//
// On startup, resolves the plugin config and (if enabled + a shared secret is set) starts the
// MsteamsVoiceRuntime: the Teams media WS server, the CallLifecycle, and the per-call realtime bridge.

import { definePluginEntry } from "openclaw/plugin-sdk/core";
import { MsteamsVoiceRuntime } from "./msteams-runtime.js";
import { resolvePluginConfig } from "./plugin-config.js";

export default definePluginEntry({
  id: "msteams-voice",
  name: "Microsoft Teams Voice (CVI)",
  description: "Self-contained Microsoft Teams realtime voice agent (CVI) for OpenClaw.",
  register(api) {
    const cfg = resolvePluginConfig((api as { pluginConfig?: unknown }).pluginConfig);
    if (!cfg.enabled) return;

    const logger = api.runtime.logging.getChildLogger({ plugin: "msteams-voice" });
    if (!cfg.media.sharedSecret) {
      logger.warn("msteams-voice: sharedSecret is not configured — media server not started");
      return;
    }

    const runtime = new MsteamsVoiceRuntime(api, cfg);
    void runtime.start().catch((err) => {
      logger.error(
        `msteams-voice: failed to start — ${err instanceof Error ? err.message : String(err)}`,
      );
    });

    // TODO(teardown): no on-dispose hook is wired yet — the media WS + reaper live for the process
    // lifetime (the reaper interval is unref'd). Call runtime.stop() from a teardown hook once the
    // plugin API exposes one.
  },
});
