// Local resolved config for the self-contained Teams bridge plugin.
// Hand-written config, only the fields the CVI realtime/TTS code
// reads. Populated from `api.pluginConfig` (validated by openclaw.plugin.json's configSchema).

import type {
  RealtimeVoiceAgentConsultToolPolicy,
  RealtimeVoiceFastContextConfig,
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
    fastContext: RealtimeVoiceFastContextConfig;
  };
  /**
   * Streaming-mode STT provider selection. Auto-selects from the host's configured realtime
   * transcription providers when `provider` is omitted; if none resolve, the streaming path falls
   * back to file-based STT (`api.runtime.mediaUnderstanding.transcribeAudioFile`).
   */
  stt?: {
    provider?: string;
    providers?: Record<string, Record<string, unknown>>;
  };
  msteams?: {
    requireRecordingStatus?: boolean;
    groupCall?: {
      requireAddress?: boolean;
      wakePhrases?: string[];
      followUpWindowMs?: number;
    };
    /**
     * Sliding 60-second cap on PAID vision calls per Teams call, shared by `look_at_screen`, the
     * ambient push and the streaming per-turn attach. `0` turns vision spend OFF; there is no
     * "unlimited" value (see {@link ../vision-budget.js VisionBudget}). Default 30.
     */
    maxVisionPerMinute?: number;
    /**
     * Continuous (unrequested) vision: keep pushing the newest changed frame at the model between
     * turns. OFF by default - it spends a vision call per scene change for the whole call, so it is
     * opt-in. `look_at_screen` is unaffected; the agent can always ask for a look.
     */
    ambientVision?: boolean;
    meetingRecap?: boolean;
    bilingual?: boolean;
  };
  tts?: VoiceCallTtsConfig;
}

/** Alias — the CVI files were authored against `VoiceCallConfig`. */
export type MsteamsBridgeConfig = VoiceCallConfig;
