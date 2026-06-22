// MsteamsVoiceRuntime — self-contained orchestration that replaces voice-call's MsteamsProvider +
// CallManager for the standalone plugin. Owns the Teams media WebSocket, drives CallLifecycle, and
// bridges each call to the realtime voice model via createMsteamsRealtimeCall.
//
// Scope: REALTIME INBOUND speech-to-speech path. Deferred (notes): outbound call-backs (worker
// place-call), the streaming STT→agent→TTS path, and getCallStatus as a VoiceCallProvider — none are
// needed for a Teams realtime assistant and would re-introduce the heavier provider surface.

import type { OpenClawConfig } from "openclaw/plugin-sdk/config-contracts";
import type { OpenClawPluginApi } from "openclaw/plugin-sdk/core";
import { resolveConfiguredRealtimeVoiceProvider } from "openclaw/plugin-sdk/realtime-voice";
import { isInboundCallAllowed } from "./allowlist.js";
import { CallLifecycle, type SyncKeyedStore } from "./call-lifecycle.js";
import { resolveGroupCallGateConfig } from "./group-call-gate.js";
import {
  MsteamsMediaStream,
  type MsteamsLogger,
  type MsteamsSession,
} from "./msteams-media-stream.js";
import {
  createMsteamsRealtimeCall,
  type MsteamsRealtimeCall,
  type MsteamsRealtimeDeps,
} from "./msteams-realtime.js";
import { MsteamsVisionStore } from "./msteams-vision-store.js";
import { VisionBudget } from "./vision-budget.js";
import type { ResolvedPluginConfig } from "./plugin-config.js";

export class MsteamsVoiceRuntime {
  private readonly lifecycle: CallLifecycle;
  private readonly media: MsteamsMediaStream;
  private readonly vision: MsteamsVisionStore;
  private readonly visionBudget: VisionBudget;
  private readonly calls = new Map<string, MsteamsRealtimeCall>();
  private readonly log: MsteamsLogger;
  private realtime?: { provider?: unknown; providerConfig?: unknown };

  constructor(
    private readonly api: OpenClawPluginApi,
    private readonly cfg: ResolvedPluginConfig,
  ) {
    const logger = api.runtime.logging.getChildLogger({ plugin: "msteams-voice" });
    this.log = {
      info: (m) => logger.info(m),
      warn: (m) => logger.warn(m),
      error: (m) => logger.error(m),
      debug: (m) => logger.debug?.(m),
    };
    this.visionBudget = new VisionBudget(this.cfg.voice.msteams?.maxVisionPerMinute ?? 30);
    this.vision = new MsteamsVisionStore(() => this.cfg.voice.msteams?.maxVisionPerMinute ?? 30);
    this.vision.setBudget(this.visionBudget);

    this.lifecycle = new CallLifecycle(
      {
        // Adapt api.runtime.state's sync keyed store (namespace/lookup/register/delete/entries) to
        // our minimal get/set/delete/keys surface. (plugin-state-store.types.ts)
        openSyncKeyedStore: <T>(name: string): SyncKeyedStore<T> => {
          const s = api.runtime.state.openSyncKeyedStore<T>({
            namespace: name,
            maxEntries: 2000,
          });
          return {
            get: (k) => s.lookup(k),
            set: (k, v) => s.register(k, v),
            delete: (k) => {
              s.delete(k);
            },
            keys: () => s.entries().map((e) => e.key),
          };
        },
        log: this.log,
        now: () => Date.now(),
      },
      {
        maxConcurrentCalls: cfg.limits.maxConcurrentCalls,
        maxDurationMs: cfg.limits.maxDurationMs,
        staleCallReaperMs: cfg.limits.staleCallReaperMs,
      },
    );

    this.media = new MsteamsMediaStream({
      port: cfg.media.port,
      bindAddress: cfg.media.bindAddress,
      path: cfg.media.path,
      sharedSecret: cfg.media.sharedSecret,
      logger: this.log,
      onSessionStart: (s) => this.onSessionStart(s),
      onSessionEnd: (i) => this.onSessionEnd(i),
      onAudioFrame: (i) => this.calls.get(i.callId)?.pushAudio(i.payload),
      onVideoFrame: (i) => {
        this.vision.store({ ...i, callId: i.callId });
        this.calls.get(i.callId)?.notifyInboundFrame();
      },
      onRecordingStatus: (i) => this.calls.get(i.callId)?.setRecordingActive(i.status === "active"),
      onDtmf: (i) => this.calls.get(i.callId)?.notifyDtmf(i.digit),
      onParticipants: (i) => this.calls.get(i.callId)?.setHumanCount(i.count),
    });
  }

  async start(): Promise<void> {
    this.realtime = resolveConfiguredRealtimeVoiceProvider({
      configuredProviderId: this.cfg.voice.realtime.provider,
      providerConfigs: this.cfg.voice.realtime.providers,
      cfg: this.api.config as unknown as OpenClawConfig,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
    } as any) as { provider?: unknown; providerConfig?: unknown };
    this.lifecycle.start();
    await this.media.start();
    this.log.info("[msteams-voice] started");
  }

  async stop(): Promise<void> {
    for (const call of this.calls.values()) call.close("shutdown");
    this.calls.clear();
    this.lifecycle.stop();
    await this.media.stop();
  }

  private onSessionStart(session: MsteamsSession): void {
    const provider = this.realtime?.provider;
    if (!provider) {
      this.log.error("[msteams-voice] no realtime voice provider resolved — rejecting call");
      session.close("realtime-unavailable");
      return;
    }
    const from = session.caller?.aadId ?? "";
    if (!isInboundCallAllowed(this.cfg.voice.inboundPolicy, this.cfg.voice.allowFrom, from)) {
      this.log.warn(
        `[msteams-voice] inbound call rejected by policy "${this.cfg.voice.inboundPolicy ?? "disabled"}"`,
      );
      session.close("not-allowed");
      return;
    }
    try {
      this.lifecycle.initiate({
        callId: session.callId,
        providerCallId: session.callId,
        direction: "inbound",
        from,
        to: "",
      });
    } catch (err) {
      this.log.warn(`[msteams-voice] cannot accept call: ${String(err)}`);
      session.close("busy");
      return;
    }
    // Inbound realtime is active as soon as the bridge connects; mark answered so the
    // unanswered-call reaper doesn't kill it (maxDuration still applies).
    this.lifecycle.answer(session.callId);

    const deps: MsteamsRealtimeDeps = {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      provider: provider as any,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      providerConfig: this.realtime?.providerConfig as any,
      cfg: this.api.config as unknown as OpenClawConfig,
      instructions: this.cfg.voice.realtime.instructions,
      greetingInstructions: this.cfg.voice.inboundGreeting,
      inboundPolicy: this.cfg.voice.inboundPolicy,
      allowFrom: this.cfg.voice.allowFrom,
      requireRecordingStatus: this.cfg.voice.msteams?.requireRecordingStatus,
      tools: this.cfg.voice.realtime.tools,
      toolPolicy: this.cfg.voice.realtime.toolPolicy,
      suppressInputDuringPlayback: this.cfg.voice.realtime.suppressInputDuringPlayback,
      echoSuppressionWindowMs: this.cfg.voice.realtime.echoSuppressionWindowMs,
      echoBargeInRms: this.cfg.voice.realtime.echoBargeInRms,
      groupCallGate: resolveGroupCallGateConfig(this.cfg.voice.msteams?.groupCall),
      visionBudget: this.visionBudget,
      getLatestFrame: (src) => this.vision.getLatest(session.callId, src),
      getFrameHistory: (limit) => this.vision.getHistory(session.callId, limit),
      agentRuntime: this.api.runtime.agent,
      voiceConfig: this.cfg.voice,
      logger: this.log,
    };

    this.calls.set(session.callId, createMsteamsRealtimeCall({ session, deps }));
  }

  private onSessionEnd(info: { callId: string; reason: string }): void {
    this.calls.get(info.callId)?.close();
    this.calls.delete(info.callId);
    this.lifecycle.end(info.callId, "hangup-user");
  }
}
