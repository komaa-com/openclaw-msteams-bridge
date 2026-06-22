// Local resolved config for the self-contained Teams voice plugin.
// Hand-written (NOT the 1,078-line voice-call config) — only the fields the CVI realtime/TTS code
// reads. Populated from `api.pluginConfig` (validated by openclaw.plugin.json's configSchema).

import type {
  RealtimeVoiceAgentConsultToolPolicy,
  RealtimeVoiceFastContextConfig,
  RealtimeVoiceTool,
} from "openclaw/plugin-sdk/realtime-voice";

/** Host TTS config (messages.tts) — opaque to this plugin; merged + passed to the TTS runtime. */
export type VoiceCallTtsConfig = Record<string, unknown>;

export interface VoiceCallConfig {
  /** Agent id the consult/recap runs against (defaults to "main"). */
  agentId?: string;
  sessionScope?: "per-phone" | "per-call" | "per-thread";
  responseModel?: string;
  /** Resolved config always carries these (defaults applied at load). */
  responseTimeoutMs: number;
  inboundPolicy?: "disabled" | "allowlist" | "pairing" | "open";
  allowFrom?: string[];
  inboundGreeting?: string;
  /**
   * Voice path: "realtime" speech-to-speech model (default when a realtime provider is configured)
   * or "streaming" STT→agent→TTS (for non-realtime providers). Defaults to "realtime" if a realtime
   * provider resolves, else "streaming".
   */
  mode?: "realtime" | "streaming";
  realtime: {
    provider?: string;
    providers?: Record<string, Record<string, unknown>>;
    instructions?: string;
    toolPolicy: RealtimeVoiceAgentConsultToolPolicy;
    consultPolicy?: "auto" | "substantive" | "always";
    consultThinkingLevel?:
      | "off"
      | "minimal"
      | "low"
      | "medium"
      | "high"
      | "xhigh"
      | "adaptive"
      | "max";
    consultFastMode?: boolean;
    suppressInputDuringPlayback?: boolean;
    echoSuppressionWindowMs?: number;
    echoBargeInRms?: number;
    tools?: RealtimeVoiceTool[];
    fastContext: RealtimeVoiceFastContextConfig;
  };
  msteams?: {
    requireRecordingStatus?: boolean;
    groupCall?: {
      requireAddress?: boolean;
      wakePhrases?: string[];
      followUpWindowMs?: number;
    };
    maxVisionPerMinute?: number;
    meetingRecap?: boolean;
    bilingual?: boolean;
  };
  tts?: VoiceCallTtsConfig;
}

/** Alias — the CVI files were authored against `VoiceCallConfig`. */
export type MsteamsVoiceConfig = VoiceCallConfig;
