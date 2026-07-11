// Narrow host runtime/config contracts, derived from the PUBLIC plugin API only.

import type { OpenClawPluginApi } from "openclaw/plugin-sdk/core";
import type { VoiceCallTtsConfig } from "./config.js";

/** Core config subset read by the TTS adapter. */
export type CoreConfig = {
  session?: {
    store?: string;
  };
  messages?: {
    tts?: VoiceCallTtsConfig;
  };
  [key: string]: unknown;
};

/** Agent runtime API subset (api.runtime.agent), exposed through the public plugin SDK. */
export type CoreAgentDeps = OpenClawPluginApi["runtime"]["agent"];
