// MsteamsVoiceRuntime — self-contained orchestration that replaces voice-call's MsteamsProvider +
// CallManager for the standalone plugin. Owns the Teams media WebSocket, drives CallLifecycle, and
// bridges each call to the realtime voice model via createMsteamsRealtimeCall.
//
// Scope: REALTIME INBOUND speech-to-speech path. Deferred (notes): outbound call-backs (worker
// place-call), the streaming STT→agent→TTS path, and getCallStatus as a VoiceCallProvider — none are
// needed for a Teams realtime assistant and would re-introduce the heavier provider surface.

import type { OpenClawConfig } from "openclaw/plugin-sdk/config-contracts";
import type { OpenClawPluginApi } from "openclaw/plugin-sdk/core";
import {
  consultRealtimeVoiceAgent,
  resolveConfiguredRealtimeVoiceProvider,
  resolveRealtimeVoiceAgentConsultToolsAllow,
} from "openclaw/plugin-sdk/realtime-voice";
import {
  getRealtimeTranscriptionProvider,
  listRealtimeTranscriptionProviders,
  type RealtimeTranscriptionProviderConfig,
  type RealtimeTranscriptionProviderPlugin,
} from "openclaw/plugin-sdk/realtime-transcription";
import { resolveConfiguredCapabilityProvider } from "openclaw/plugin-sdk/provider-selection-runtime";
import { fetchWithSsrFGuard } from "openclaw/plugin-sdk/ssrf-runtime";
import { createHmac } from "node:crypto";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import { isInboundCallAllowed } from "./allowlist.js";
import { CallLifecycle, type SyncKeyedStore } from "./call-lifecycle.js";
import type { CoreConfig } from "./core-bridge.js";
import { resolveGroupCallGateConfig } from "./group-call-gate.js";
import { collectLatestFrameImages } from "./vision-consult.js";
import {
  MSTEAMS_PCM_SAMPLE_RATE_HZ,
  MsteamsMediaStream,
  type MsteamsLogger,
  type MsteamsSession,
} from "./msteams-media-stream.js";
import {
  createMsteamsRealtimeCall,
  type MsteamsRealtimeCall,
  type MsteamsRealtimeDeps,
} from "./msteams-realtime.js";
import { createMsteamsStreamingCall, type MsteamsStreamingDeps } from "./msteams-streaming.js";
import { createMsteamsTtsProvider, type MsteamsTtsProvider } from "./msteams-tts.js";
import { MsteamsVisionStore } from "./msteams-vision-store.js";
import { resolveVoiceResponseModel } from "./response-model.js";
import { VisionBudget } from "./vision-budget.js";
import type { ResolvedPluginConfig } from "./plugin-config.js";

/** Default no-answer guard for a placed outbound call (overridable via outbound.answerTimeoutMs). */
const OUTBOUND_ANSWER_TIMEOUT_DEFAULT_MS = 120_000;

export type PlaceCallMode = "notify" | "conversation";

export class MsteamsVoiceRuntime {
  private readonly lifecycle: CallLifecycle;
  private readonly media: MsteamsMediaStream;
  private readonly vision: MsteamsVisionStore;
  private readonly visionBudget: VisionBudget;
  private readonly calls = new Map<string, MsteamsRealtimeCall>();
  private readonly log: MsteamsLogger;
  private realtime?: { provider?: unknown; providerConfig?: unknown };
  /** Resolved streaming STT provider (mode:"streaming"); undefined → file-based STT fallback. */
  private transcription?: {
    provider: RealtimeTranscriptionProviderPlugin;
    providerConfig: RealtimeTranscriptionProviderConfig;
  };
  /** Selected voice path; finalized in start() once the realtime provider is resolved. */
  private mode: "realtime" | "streaming" = "realtime";
  /** Lazily-built TTS provider for the streaming path (api.runtime.tts). */
  private ttsProvider?: MsteamsTtsProvider;
  /** Monotonic suffix for streaming STT temp-file names. */
  private sttSeq = 0;
  /** Calls we placed via the worker, awaiting their media WS session.start to attach. */
  private readonly pendingOutbound = new Map<
    string,
    { to: string; message?: string; mode: PlaceCallMode }
  >();
  private readonly pendingOutboundTimers = new Map<string, ReturnType<typeof setTimeout>>();

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
    // "realtime" when a realtime provider resolved (or explicitly configured), else "streaming".
    this.mode =
      this.cfg.voice.mode ?? (this.realtime?.provider ? "realtime" : "streaming");
    if (this.mode === "streaming") this.resolveTranscriptionProvider();
    this.lifecycle.start();
    await this.media.start();
    this.log.info(`[msteams-voice] started (mode=${this.mode})`);
  }

  async stop(): Promise<void> {
    for (const t of this.pendingOutboundTimers.values()) clearTimeout(t);
    this.pendingOutboundTimers.clear();
    this.pendingOutbound.clear();
    for (const call of this.calls.values()) call.close("shutdown");
    this.calls.clear();
    this.lifecycle.stop();
    await this.media.stop();
  }

  /**
   * Place an outbound Teams call to a user (AAD object id, optionally "user:"-prefixed) via the
   * worker. `mode` "notify" instructs the model to deliver `message` and end; "conversation" starts a
   * full realtime conversation. A no-answer timer finalizes the call if it never connects back
   * (declined/offline → effectively voicemail/no-answer). Returns the worker call id.
   */
  async placeCall(
    to: string,
    opts?: { message?: string; mode?: PlaceCallMode },
  ): Promise<{ callId: string }> {
    const ob = this.cfg.outbound;
    if (!ob?.enabled)
      throw new Error("msteams-voice: outbound calling is disabled (set outbound.enabled)");
    if (!ob.workerBaseUrl)
      throw new Error("msteams-voice: outbound.workerBaseUrl is not configured");
    if (!ob.tenantId) throw new Error("msteams-voice: outbound.tenantId is not configured");
    if (!this.cfg.media.sharedSecret)
      throw new Error("msteams-voice: sharedSecret is not configured");
    const userObjectId = to.replace(/^user:/i, "").trim();
    if (!userObjectId) throw new Error("msteams-voice: target userObjectId (to) is required");
    if (this.lifecycle.activeCount() >= this.cfg.limits.maxConcurrentCalls)
      throw new Error("msteams-voice: max concurrent calls reached; not placing outbound call");

    // HMAC over `${timestampMs}.${userObjectId}` — same scheme as the media WS handshake.
    const timestampMs = Date.now();
    const signature = createHmac("sha256", this.cfg.media.sharedSecret)
      .update(`${timestampMs}.${userObjectId}`)
      .digest("hex");
    const url = `${ob.workerBaseUrl.replace(/\/+$/, "")}/api/calls`;
    const { response, release } = await fetchWithSsrFGuard({
      url,
      init: {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-openclawteamsbridge-timestamp": String(timestampMs),
          "x-openclawteamsbridge-signature": signature,
        },
        body: JSON.stringify({ userObjectId, tenantId: ob.tenantId }),
      },
      // The worker is operator-configured trusted infra, often on loopback → permit private network.
      policy: { allowedHostnames: [new URL(url).hostname], allowPrivateNetwork: true },
    });
    let workerCallId: string | undefined;
    try {
      if (!response.ok) {
        const text = await response.text().catch(() => "");
        throw new Error(
          `msteams-voice: worker returned ${response.status}${text ? ` — ${text.slice(0, 200)}` : ""}`,
        );
      }
      const payload = (await response.json().catch(() => ({}))) as { callId?: string };
      workerCallId = payload.callId;
    } finally {
      await release();
    }
    if (!workerCallId)
      throw new Error("msteams-voice: worker response did not include a callId");

    const mode: PlaceCallMode = opts?.mode ?? ob.defaultMode ?? "notify";
    this.lifecycle.initiate({
      callId: workerCallId,
      providerCallId: workerCallId,
      direction: "outbound",
      from: "",
      to,
      message: opts?.message,
    });
    this.pendingOutbound.set(workerCallId, { to, message: opts?.message, mode });
    const timer = setTimeout(
      () => this.finalizeUnansweredOutbound(workerCallId as string),
      ob.answerTimeoutMs ?? OUTBOUND_ANSWER_TIMEOUT_DEFAULT_MS,
    );
    timer.unref?.();
    this.pendingOutboundTimers.set(workerCallId, timer);
    this.log.info(
      `[msteams-voice] outbound call placed callId=${workerCallId} -> ${userObjectId} (${mode})`,
    );
    return { callId: workerCallId };
  }

  private finalizeUnansweredOutbound(callId: string): void {
    if (!this.pendingOutbound.has(callId)) return;
    this.pendingOutbound.delete(callId);
    this.clearOutboundTimer(callId);
    this.log.warn(
      `[msteams-voice] outbound call ${callId} not answered within timeout; finalizing (no-answer/voicemail)`,
    );
    this.lifecycle.end(callId, "no-answer");
  }

  private clearOutboundTimer(callId: string): void {
    const t = this.pendingOutboundTimers.get(callId);
    if (t) clearTimeout(t);
    this.pendingOutboundTimers.delete(callId);
  }

  /**
   * Query the current status of a call (state + whether it has reached a terminal state).
   * Returns undefined for an unknown call id. Note: OpenClaw exposes no provider-status registration
   * hook for non-channel plugins, so this is surfaced as a runtime method (callable by an embedding
   * host or a future admin/tool surface) rather than a registered VoiceCallProvider.getCallStatus.
   */
  getCallStatus(callId: string): ReturnType<CallLifecycle["getStatus"]> {
    return this.lifecycle.getStatus(callId);
  }

  private onSessionStart(session: MsteamsSession): void {
    // Realtime mode requires a resolved provider; streaming mode does not.
    if (this.mode === "realtime" && !this.realtime?.provider) {
      this.log.error("[msteams-voice] no realtime voice provider resolved — rejecting call");
      session.close("realtime-unavailable");
      return;
    }
    // Outbound: a call we placed via the worker has connected back (media WS attached).
    const pending = this.pendingOutbound.get(session.callId);
    if (pending) {
      this.pendingOutbound.delete(session.callId);
      this.clearOutboundTimer(session.callId);
      this.lifecycle.answer(session.callId);
      const greeting =
        pending.mode === "notify" && pending.message
          ? `Deliver this message to the person, then say goodbye and end the call: "${pending.message}"`
          : (pending.message ?? this.cfg.voice.inboundGreeting);
      this.calls.set(session.callId, this.createCall(session, greeting));
      return;
    }

    // Inbound: enforce caller policy before accepting.
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
    // Inbound is active as soon as the bridge connects; mark answered so the
    // unanswered-call reaper doesn't kill it (maxDuration still applies).
    this.lifecycle.answer(session.callId);
    this.calls.set(session.callId, this.createCall(session, this.cfg.voice.inboundGreeting));
  }

  /** Build the call handle for the selected voice path (realtime speech-to-speech vs streaming). */
  private createCall(session: MsteamsSession, greeting?: string): MsteamsRealtimeCall {
    if (this.mode === "streaming") {
      return createMsteamsStreamingCall({ session, deps: this.buildStreamingDeps(session, greeting) });
    }
    return createMsteamsRealtimeCall({
      session,
      deps: this.buildDeps(session, this.realtime?.provider, greeting),
    });
  }

  private getTtsProvider(): MsteamsTtsProvider {
    if (!this.ttsProvider) {
      this.ttsProvider = createMsteamsTtsProvider({
        coreConfig: this.api.config as unknown as CoreConfig,
        ttsOverride: this.cfg.voice.tts,
        runtime: this.api.runtime.tts,
        logger: { warn: (m) => this.log.warn(m) },
      });
    }
    return this.ttsProvider;
  }

  /** Consult session key, honoring sessionScope (mirrors the realtime path). */
  private streamingSessionKey(session: MsteamsSession): string {
    const scope = this.cfg.voice.sessionScope;
    if (scope === "per-call") return `msteams:${session.callId}`;
    if (scope === "per-thread") return `msteams:${session.threadId || session.callId}`;
    return `msteams:${session.caller?.aadId || session.callId}`;
  }

  /**
   * Resolve a streaming STT provider (mode:"streaming"). Auto-selects from the host's configured
   * realtime transcription providers when `stt.provider` is omitted; leaves `this.transcription`
   * undefined (→ file-based STT fallback) when none resolve.
   */
  private resolveTranscriptionProvider(): void {
    const cfg = this.api.config as unknown as OpenClawConfig;
    const res = resolveConfiguredCapabilityProvider({
      configuredProviderId: this.cfg.voice.stt?.provider,
      providerConfigs: this.cfg.voice.stt?.providers,
      cfg,
      cfgForResolve: cfg,
      getConfiguredProvider: (id) => getRealtimeTranscriptionProvider(id, cfg),
      listProviders: () => listRealtimeTranscriptionProviders(cfg),
      resolveProviderConfig: ({ provider, cfg: c, rawConfig }) =>
        provider.resolveConfig?.({ cfg: c, rawConfig }) ?? rawConfig,
      isProviderConfigured: ({ provider, cfg: c, providerConfig }) =>
        provider.isConfigured({ cfg: c, providerConfig }),
    });
    if (res.ok) {
      this.transcription = { provider: res.provider, providerConfig: res.providerConfig };
      this.log.info(`[msteams-voice] streaming STT provider: ${res.provider.id}`);
    } else {
      this.log.info(
        `[msteams-voice] no streaming STT provider resolved (${res.code}); using file-based STT fallback`,
      );
    }
  }

  private buildStreamingDeps(session: MsteamsSession, greeting?: string): MsteamsStreamingDeps {
    const cfg = this.api.config as unknown as OpenClawConfig;
    const agentRuntime = this.api.runtime.agent;
    const transcription = this.transcription;
    return {
      // Preferred: a live streaming STT session (lower latency) when a provider resolved.
      ...(transcription
        ? {
            createTranscriptionSession: (callbacks) =>
              transcription.provider.createSession({
                cfg,
                providerConfig: transcription.providerConfig,
                ...callbacks,
              }),
          }
        : {}),
      // Fallback: VAD-segmented utterance → temp WAV → file STT (used when no provider resolved).
      transcribe: async (pcm16k: Buffer): Promise<string> => {
        const wav = pcmToWav(pcm16k, MSTEAMS_PCM_SAMPLE_RATE_HZ);
        const tmp = path.join(os.tmpdir(), `msteams-voice-${session.callId}-${this.sttSeq++}.wav`);
        await fs.writeFile(tmp, wav);
        try {
          const res = await this.api.runtime.mediaUnderstanding.transcribeAudioFile({
            filePath: tmp,
            cfg,
          });
          return res.text ?? "";
        } finally {
          await fs.unlink(tmp).catch(() => {});
        }
      },
      consult: async ({ question, transcript, images }) => {
        const { provider, model } = resolveVoiceResponseModel({
          voiceConfig: this.cfg.voice,
          agentRuntime,
        });
        const thinkLevel =
          this.cfg.voice.realtime.consultThinkingLevel ??
          agentRuntime.resolveThinkingDefault({ cfg, provider, model });
        const result = await consultRealtimeVoiceAgent({
          cfg,
          agentRuntime,
          logger: { warn: (m) => this.log.warn(m) },
          agentId: this.cfg.voice.agentId ?? "main",
          sessionKey: this.streamingSessionKey(session),
          messageProvider: "voice",
          lane: "voice",
          runIdPrefix: `msteams-stream-${session.callId}`,
          args: { question },
          ...(images && images.length ? { images } : {}),
          transcript,
          surface: "a Microsoft Teams call",
          userLabel: "Caller",
          assistantLabel: "Agent",
          questionSourceLabel: "caller",
          provider,
          model,
          thinkLevel,
          fastMode: this.cfg.voice.realtime.consultFastMode,
          toolsAllow: resolveRealtimeVoiceAgentConsultToolsAllow(this.cfg.voice.realtime.toolPolicy),
        });
        return { text: result.text };
      },
      ttsProvider: this.getTtsProvider(),
      greetingInstruction: greeting,
      suppressInputDuringPlayback: this.cfg.voice.realtime.suppressInputDuringPlayback,
      echoBargeInRms: this.cfg.voice.realtime.echoBargeInRms,
      requireRecordingStatus: this.cfg.voice.msteams?.requireRecordingStatus,
      groupCallGate: resolveGroupCallGateConfig(this.cfg.voice.msteams?.groupCall),
      getVisionImages: () =>
        collectLatestFrameImages({
          getLatestFrame: (s) => this.vision.getLatest(session.callId, s),
          visionBudget: this.visionBudget,
          callId: session.callId,
        }),
      appendTranscript: (e) => this.lifecycle.appendTranscript(session.callId, e),
      logger: this.log,
    };
  }

  private buildDeps(
    session: MsteamsSession,
    provider: unknown,
    greetingInstructions?: string,
  ): MsteamsRealtimeDeps {
    return {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      provider: provider as any,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      providerConfig: this.realtime?.providerConfig as any,
      cfg: this.api.config as unknown as OpenClawConfig,
      instructions: this.cfg.voice.realtime.instructions,
      greetingInstructions,
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
  }

  private onSessionEnd(info: { callId: string; reason: string }): void {
    this.pendingOutbound.delete(info.callId);
    this.clearOutboundTimer(info.callId);
    this.calls.get(info.callId)?.close();
    this.calls.delete(info.callId);
    this.lifecycle.end(info.callId, "hangup-user");
  }
}

/** Wrap raw PCM (16-bit mono LE) in a minimal WAV container so file-based STT can read it. */
function pcmToWav(pcm: Buffer, sampleRate: number): Buffer {
  const numChannels = 1;
  const bitsPerSample = 16;
  const byteRate = (sampleRate * numChannels * bitsPerSample) / 8;
  const blockAlign = (numChannels * bitsPerSample) / 8;
  const header = Buffer.alloc(44);
  header.write("RIFF", 0);
  header.writeUInt32LE(36 + pcm.length, 4);
  header.write("WAVE", 8);
  header.write("fmt ", 12);
  header.writeUInt32LE(16, 16);
  header.writeUInt16LE(1, 20); // PCM
  header.writeUInt16LE(numChannels, 22);
  header.writeUInt32LE(sampleRate, 24);
  header.writeUInt32LE(byteRate, 28);
  header.writeUInt16LE(blockAlign, 32);
  header.writeUInt16LE(bitsPerSample, 34);
  header.write("data", 36);
  header.writeUInt32LE(pcm.length, 40);
  return Buffer.concat([header, pcm]);
}
