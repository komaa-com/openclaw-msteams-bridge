// Voice Call plugin module implements response model behavior.
import type { OpenClawConfig } from "openclaw/plugin-sdk/config-contracts";
import type { VoiceCallConfig } from "./config.js";
import type { CoreAgentDeps } from "./core-bridge.js";

// Resolves the model used for text response generation.

/**
 * Resolve provider/model fields for the consult/task lanes.
 *
 * Order matters, and the middle step is the one that was missing. `agentRuntime.defaults` is NOT the
 * user's agent - it is typed `{ model: typeof DEFAULT_MODEL, provider: typeof DEFAULT_PROVIDER }`, two
 * constants compiled into openclaw ("openai" / "gpt-5.5"). Falling straight to it meant an install that
 * had configured `agents.defaults.model.primary` to its own provider still asked for openai/gpt-5.5,
 * which is not a registered model, so every consult died with `FailoverError: Unknown model` and the
 * call lane logged a lookup failure mid-conversation. The agent the user configured is the sane default;
 * the compiled-in pair is only a last resort for an install that has configured nothing at all.
 */
export function resolveVoiceResponseModel(params: {
  voiceConfig: VoiceCallConfig;
  agentRuntime: CoreAgentDeps;
  /** Host config, used to honour the configured agent before the compiled-in default. */
  cfg?: OpenClawConfig;
}): {
  modelRef: string;
  provider: string;
  model: string;
} {
  // `model` is `string | { primary?: string }` - both spellings are valid config.
  const configured = params.cfg?.agents?.defaults?.model;
  const configuredPrimary = (typeof configured === "string" ? configured : configured?.primary)?.trim();
  const modelRef =
    params.voiceConfig.responseModel ??
    (configuredPrimary || `${params.agentRuntime.defaults.provider}/${params.agentRuntime.defaults.model}`);
  const slashIndex = modelRef.indexOf("/");

  return {
    modelRef,
    provider:
      slashIndex === -1 ? params.agentRuntime.defaults.provider : modelRef.slice(0, slashIndex),
    model: slashIndex === -1 ? modelRef : modelRef.slice(slashIndex + 1),
  };
}
