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
import { fetchWithSsrFGuard } from "openclaw/plugin-sdk/ssrf-runtime";
import { createHmac } from "node:crypto";
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
    this.lifecycle.start();
    await this.media.start();
    this.log.info("[msteams-voice] started");
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
    const provider = this.realtime?.provider;
    if (!provider) {
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
      this.calls.set(
        session.callId,
        createMsteamsRealtimeCall({ session, deps: this.buildDeps(session, provider, greeting) }),
      );
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
    // Inbound realtime is active as soon as the bridge connects; mark answered so the
    // unanswered-call reaper doesn't kill it (maxDuration still applies).
    this.lifecycle.answer(session.callId);
    this.calls.set(
      session.callId,
      createMsteamsRealtimeCall({
        session,
        deps: this.buildDeps(session, provider, this.cfg.voice.inboundGreeting),
      }),
    );
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
