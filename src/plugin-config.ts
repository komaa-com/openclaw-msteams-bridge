// Resolve the raw plugin config (api.pluginConfig, validated against openclaw.plugin.json's
// configSchema) into (a) the runtime/media settings and (b) the `MsteamsVoiceConfig` the CVI bridge
// reads. Boundary adapter — tolerant casts on untyped raw input.

import type { MsteamsVoiceConfig } from "./config.js";

export interface ResolvedPluginConfig {
  enabled: boolean;
  media: { port: number; bindAddress?: string; path: string; sharedSecret: string };
  outbound?: {
    enabled?: boolean;
    workerBaseUrl?: string;
    tenantId?: string;
    answerTimeoutMs?: number;
    defaultMode?: "notify" | "conversation";
  };
  limits: { maxConcurrentCalls: number; maxDurationMs: number; staleCallReaperMs: number };
  voice: MsteamsVoiceConfig;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Raw = Record<string, any>;

export function resolvePluginConfig(rawInput: unknown): ResolvedPluginConfig {
  const c: Raw = (rawInput as Raw) ?? {};
  const r: Raw = c.realtime ?? {};
  return {
    enabled: c.enabled !== false,
    media: {
      port: Number(c.port ?? 9442),
      bindAddress: c.bindAddress,
      path: String(c.path ?? "/voice/msteams/stream"),
      sharedSecret: String(c.sharedSecret ?? ""),
    },
    outbound: c.outbound,
    limits: {
      maxConcurrentCalls: Number(c.maxConcurrentCalls ?? 4),
      maxDurationMs: Number(c.maxDurationSeconds ?? 0) * 1000,
      staleCallReaperMs: Number(c.staleCallReaperSeconds ?? 120) * 1000,
    },
    voice: {
      agentId: c.agentId,
      sessionScope: c.sessionScope,
      responseModel: c.responseModel,
      responseTimeoutMs: Number(c.responseTimeoutMs ?? 30000),
      inboundPolicy: c.inboundPolicy,
      allowFrom: c.allowFrom,
      inboundGreeting: c.inboundGreeting,
      mode: c.mode,
      realtime: {
        provider: r.provider,
        providers: r.providers,
        instructions: r.instructions,
        toolPolicy: r.toolPolicy ?? "none",
        consultPolicy: r.consultPolicy,
        consultThinkingLevel: r.consultThinkingLevel,
        consultFastMode: r.consultFastMode,
        suppressInputDuringPlayback: r.suppressInputDuringPlayback,
        echoSuppressionWindowMs: r.echoSuppressionWindowMs,
        echoBargeInRms: r.echoBargeInRms,
        tools: r.tools,
        // Default fast-context off; shape comes from the SDK type.
        fastContext: (r.fastContext ?? {
          enabled: false,
          timeoutMs: 800,
          maxResults: 3,
          sources: ["memory", "sessions"],
          fallbackToConsult: false,
        }) as MsteamsVoiceConfig["realtime"]["fastContext"],
      },
      stt: c.stt,
      // Manifest exposes these flat (own plugin namespace); build the nested `msteams` object the
      // CVI bridge reads (config.msteams.*). A top-level `msteams` key would be rejected by the
      // manifest's additionalProperties:false.
      msteams: {
        requireRecordingStatus: c.requireRecordingStatus,
        groupCall: c.groupCall,
        maxVisionPerMinute: c.maxVisionPerMinute,
        meetingRecap: c.meetingRecap,
        bilingual: c.bilingual,
      },
      tts: c.tts,
    },
  };
}
