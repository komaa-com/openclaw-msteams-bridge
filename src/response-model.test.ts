import type { OpenClawConfig } from "openclaw/plugin-sdk/config-contracts";
import { describe, expect, it } from "vitest";
import type { VoiceCallConfig } from "./config.js";
import type { CoreAgentDeps } from "./core-bridge.js";
import { resolveVoiceResponseModel } from "./response-model.js";

/**
 * Which model the consult/task lanes ask for.
 *
 * The case that mattered: an install had configured `agents.defaults.model.primary` to its own
 * provider and set no `responseModel`, so this resolver fell straight through to openclaw's compiled-in
 * DEFAULT_PROVIDER/DEFAULT_MODEL pair ("openai"/"gpt-5.5"). That model is not registered on such an
 * install, so every consult failed with `FailoverError: Unknown model: openai/gpt-5.5` mid-call - and
 * only mid-call, because the realtime leg has its own provider block and connected fine.
 */
describe("resolveVoiceResponseModel", () => {
  // openclaw's compiled-in constants, deliberately NOT the caller's provider.
  const agentRuntime = { defaults: { provider: "openai", model: "gpt-5.5" } } as CoreAgentDeps;
  const voice = (responseModel?: string) => ({ responseModel }) as VoiceCallConfig;
  const cfgWith = (model: unknown) =>
    ({ agents: { defaults: { model } } }) as unknown as OpenClawConfig;

  it("prefers an explicit responseModel over everything", () => {
    const r = resolveVoiceResponseModel({
      voiceConfig: voice("microsoft-foundry/gpt-5.6-sol"),
      agentRuntime,
      cfg: cfgWith({ primary: "anthropic/claude-opus-5" }),
    });
    expect(r).toMatchObject({ provider: "microsoft-foundry", model: "gpt-5.6-sol" });
  });

  it("falls back to the CONFIGURED agent, not the compiled-in default", () => {
    // The regression. Without cfg this returned openai/gpt-5.5 and the lane threw.
    const r = resolveVoiceResponseModel({
      voiceConfig: voice(),
      agentRuntime,
      cfg: cfgWith({ primary: "microsoft-foundry/gpt-5.6-sol" }),
    });
    expect(r.modelRef).toBe("microsoft-foundry/gpt-5.6-sol");
    expect(r).toMatchObject({ provider: "microsoft-foundry", model: "gpt-5.6-sol" });
  });

  it("accepts the bare-string spelling of agents.defaults.model", () => {
    // AgentModelConfig is `string | { primary?: string }`; both are valid config and a user who wrote
    // the short form must not silently get someone else's model.
    const r = resolveVoiceResponseModel({
      voiceConfig: voice(),
      agentRuntime,
      cfg: cfgWith("microsoft-foundry/gpt-5.6-sol"),
    });
    expect(r).toMatchObject({ provider: "microsoft-foundry", model: "gpt-5.6-sol" });
  });

  it("uses the compiled-in default only when nothing is configured", () => {
    // A stock install that configured nothing keeps its previous behaviour.
    for (const cfg of [undefined, cfgWith(undefined), cfgWith({}), cfgWith({ primary: "   " })]) {
      const r = resolveVoiceResponseModel({ voiceConfig: voice(), agentRuntime, cfg });
      expect(r).toMatchObject({ provider: "openai", model: "gpt-5.5" });
    }
  });

  it("treats a slash-less ref as a model on the default provider", () => {
    const r = resolveVoiceResponseModel({ voiceConfig: voice("gpt-5.6-sol"), agentRuntime });
    expect(r).toMatchObject({ provider: "openai", model: "gpt-5.6-sol" });
  });
});
