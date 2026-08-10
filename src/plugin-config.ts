// Resolve the raw plugin config (api.pluginConfig, validated against openclaw.plugin.json's
// configSchema) into (a) the runtime/media settings and (b) the `MsteamsVoiceConfig` the CVI bridge
// reads. Boundary adapter — tolerant casts on untyped raw input.

import type { MsteamsVoiceConfig } from "./config.js";
import { resolveManagedChatConfig, type ManagedChatConfig } from "./managed-chat.js";

export interface ResolvedPluginConfig {
  enabled: boolean;
  media: { port: number; bindAddress?: string; path: string; sharedSecret: string };
  /** StandIn managed chat lane (protocol/chat-schema.yaml). Disabled by default; BYO voice is untouched. */
  /** StandIn Managed Bot connection mode. Chat is one lane of it, hence `managedBot`, not
   * `managedChat` - the voice lane of the same connection is configured by the fields above. */
  managedChat: ManagedChatConfig;
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

/** A config value only counts when it is a non-empty STRING - see the sharedSecret note below. */
const str = (v: unknown): string => (typeof v === "string" ? v : "");
/** The compatibility blocks are objects or nothing; anything else is ignored rather than spread. */
const asObject = (v: unknown): Raw | undefined =>
  v && typeof v === "object" && !Array.isArray(v) ? (v as Raw) : undefined;

export function resolvePluginConfig(rawInput: unknown): ResolvedPluginConfig {
  const c: Raw = (rawInput as Raw) ?? {};
  const r: Raw = c.realtime ?? {};
  return {
    enabled: c.enabled !== false,
    media: {
      port: Number(c.callingPort ?? c.port ?? 9442),
      bindAddress: c.bindAddress,
      path: String(c.path ?? "/msteams/calling"),
      // ONE secret, both lanes (owner decision, matching Hermes): `secret` is THE connection secret
      // StandIn issues per binding, and the user pastes it once. `sharedSecret` / `messagesSecret`
      // remain as per-lane OVERRIDES for deployments that insist on separate keys, and win when set.
      //
      // Only accept a STRING. The manifest allows an object (secret-input reference) form; if the host
      // ever passes an UNRESOLVED object (e.g. an env descriptor whose var is unset), String({}) would yield the
      // literal "[object Object]" — a non-empty, guessable secret that the fail-closed check in index.ts would
      // accept. Coerce a non-string to "" so it fails CLOSED (server refuses to start) instead.
      sharedSecret: str(c.sharedSecret) || str(c.secret),
    },
    // The messages lane, configured with FLAT keys beside the calling ones - the two lanes of one
    // connection, not a nested sub-product. `managedBot` stays accepted as the compatibility block.
    //
    // The older `managedChat` alias is GONE from both here and the manifest schema. Keeping it in one
    // and not the other was the worst of both: configSchema sets additionalProperties:false, so a
    // config using the alias would have failed validation before this resolver ever saw it - dead code
    // behind a closed door. Removing it outright is safe here because the plugin id itself changed in
    // this release (msteams-voice -> msteams-call), so every existing config has to be edited anyway;
    // renaming one more key while you are already in the file is not an extra migration.
    //
    // Note the asymmetry that is deliberate: `secret` fills BOTH lanes, so pasting the one value the
    // portal gives you turns on calling AND messages. That is the whole point of one secret. A
    // deployment that wants calling only sets `sharedSecret` instead, which fills nothing else.
    managedChat: resolveManagedChatConfig({
      // The messages lane binds the SAME address as calling unless told otherwise - one machine, one
      // interface, and it is what Hermes does. Without this, someone setting bindAddress 127.0.0.1 for
      // the documented tunnel posture (TLS terminated publicly, proxied to loopback) got calling on
      // loopback and messages on EVERY interface, quietly exposing it on the LAN.
      ...(str(c.bindAddress) ? { bindAddress: str(c.bindAddress) } : {}),
      ...(asObject(c.managedBot) ?? {}),
      // Flat keys win over the compatibility block.
      ...(str(c.messagesSecret) || str(c.secret)
        ? { chatSecret: str(c.messagesSecret) || str(c.secret) }
        : {}),
      ...(c.messagesPort !== undefined ? { port: Number(c.messagesPort) } : {}),
      ...(str(c.messagesPath) ? { path: str(c.messagesPath) } : {}),
    }),
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
