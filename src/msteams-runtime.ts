// MsteamsVoiceRuntime, self-contained orchestration for the standalone plugin: a full msteams
// provider and call-manager. Owns the Teams media WebSocket, drives CallLifecycle, and
// bridges each call to the realtime voice model via createMsteamsRealtimeCall.
//
// Scope: realtime speech-to-speech (inbound + outbound call-backs via worker place-call) and the
// streaming STT→agent→TTS path. Still out of scope: getCallStatus as a VoiceCallProvider — not
// needed for a Teams realtime assistant and would re-introduce the heavier provider surface.

import type { OpenClawConfig } from "openclaw/plugin-sdk/config-contracts";
import type { OpenClawPluginApi } from "openclaw/plugin-sdk/core";
import { MSTEAMS_POST_CHAT_TOOL_NAME } from "./msteams-realtime-tools.js";
import {
  fetchAttachmentImages,
  ManagedChatServer,
  postManagedMessage,
  type ManagedInbound,
} from "./managed-chat.js";
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
import type { CallEndReason } from "./types.js";
import { createHash } from "node:crypto";
import { createHmac } from "node:crypto";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import { describeInboundRejection, isInboundCallAllowed } from "./allowlist.js";
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

/** The worker's outcome vocabulary, mapped onto this plugin's own end reasons.
 *
 * They are deliberately different vocabularies: CallEndReason is a closed set the lifecycle reasons
 * about, while the worker reports what Graph said. Declined and busy are both "the callee did not take
 * it", which is no-answer here; a failure is an error. Anything the map does not know falls back to
 * no-answer rather than widening the type, so a word a future worker adds cannot become an unknown
 * lifecycle state. The distinction the caller actually hears comes from the agent's own wording, not
 * from this enum. */
const WORKER_OUTCOME_TO_END_REASON: Record<string, CallEndReason> = {
  "no-answer": "no-answer",
  declined: "no-answer",
  busy: "no-answer",
  failed: "error",
};

/** Default no-answer guard for a placed outbound call (overridable via outbound.answerTimeoutMs). */
const OUTBOUND_ANSWER_TIMEOUT_DEFAULT_MS = 120_000;

export type PlaceCallMode = "notify" | "conversation";

export class MsteamsVoiceRuntime {
  private readonly lifecycle: CallLifecycle;
  private readonly media: MsteamsMediaStream;
  private readonly vision: MsteamsVisionStore;
  private readonly visionBudget: VisionBudget;
  private readonly calls = new Map<string, MsteamsRealtimeCall>();

  /**
   * Live calls that have a postable Teams conversation, keyed by call id.
   *
   * The realtime path gets its poster injected with exact per-call context. STREAMING has no
   * tool-dispatch surface at all - it is STT -> consult -> TTS - so its only route to a tool is an
   * OpenClaw tool the delegated agent turn can call. Those are registered globally and the handler
   * context carries no session or run id, so a tool cannot tell WHICH call invoked it. Hence this
   * registry, and hence the tool refusing when more than one entry is live rather than guessing: one
   * OpenClaw instance serves one binding, so a wrong guess is not a cross-tenant leak, but it would
   * still put a caller's message in someone else's conversation.
   */
  private readonly postableCalls = new Map<string, { conversationId: string; post: (text: string) => Promise<boolean> }>();
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
  /** StandIn managed chat endpoint (protocol/chat-schema.yaml); undefined unless configured. */
  private managedChat?: ManagedChatServer;

  constructor(
    private readonly api: OpenClawPluginApi,
    private readonly cfg: ResolvedPluginConfig,
  ) {
    const logger = api.runtime.logging.getChildLogger({ plugin: "msteams-call" });
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
        // In-memory keyed store for call records. api.runtime.state.openSyncKeyedStore is gated to
        // trusted (bundled/official) plugins, which a third-party npm/ClawHub install is not. Call
        // records are ephemeral — a gateway restart drops live media calls anyway — so an in-process
        // Map is sufficient and keeps the plugin installable as an untrusted plugin.
        openSyncKeyedStore: <T>(_name: string): SyncKeyedStore<T> => {
          const m = new Map<string, T>();
          return {
            get: (k) => m.get(k),
            set: (k, v) => {
              m.set(k, v);
            },
            delete: (k) => {
              m.delete(k);
            },
            keys: () => [...m.keys()],
          };
        },
        log: this.log,
        now: () => Date.now(),
      },
      {
        maxConcurrentCalls: cfg.limits.maxConcurrentCalls,
        maxDurationMs: cfg.limits.maxDurationMs,
        staleCallReaperMs: cfg.limits.staleCallReaperMs,
        // The reaper only ends the lifecycle record; run the SAME runtime teardown as a user hangup so
        // the reaped call's media + realtime sockets actually close (H7: no zombie, no maxConcurrentCalls
        // bypass). Pass the reason so the Teams worker session is closed too (the call is still live,
        // unlike a caller-driven hangup where the session is already closing).
        onReap: (callId, reason) => this.disposeCall(callId, reason),
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
      onCallOutcome: (i) => this.onCallOutcome(i.callId, i.outcome),
      onDtmf: (i) => this.calls.get(i.callId)?.notifyDtmf(i.digit),
      onParticipants: (i) => this.calls.get(i.callId)?.setHumanCount(i.count),
      // H4: the worker asks the agent to speak a line (e.g. a goodbye) in its own realtime voice.
      onAssistantSay: (i) => this.calls.get(i.callId)?.say(i.text),
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
    // Fail-LOUD at startup (mirror Hermes cli.py's fail-fast): if the operator explicitly set
    // mode:"realtime" but no provider resolved, start() would otherwise log a healthy
    // "started (mode=realtime)" and then reject EVERY inbound call with "realtime-unavailable" (the
    // #1 silent-first-call class). Name the configured provider so the missing key is obvious.
    if (this.mode === "realtime" && !this.realtime?.provider) {
      const providerId = this.cfg.voice.realtime.provider;
      this.log.warn(
        `[msteams-call] mode is "realtime" but no realtime voice provider resolved` +
          (providerId
            ? ` (configured provider "${providerId}" has no usable credentials)`
            : ` (no realtime provider configured)`) +
          `. Every inbound call will be rejected with "realtime-unavailable". Set the provider's API` +
          ` key, or set mode:"streaming".`,
      );
    }
    if (this.mode === "streaming") this.resolveTranscriptionProvider();
    if (this.cfg.managedChat.configuredWithoutSecret) {
      // Same fail-LOUD posture as the realtime-provider check above - the operator set
      // managedBot.enabled and would otherwise get a silently dead chat surface.
      this.log.warn(
        "[msteams-call] managedBot.enabled is set but no secret resolved - the messages lane is OFF. " +
          "Set `secret` to the value StandIn shows you; it covers calling and messages both.",
      );
    }

    // TRANSACTIONAL startup. Each resource that comes up is registered for rollback, so a later
    // failure - most often the messages port already being held by a previous runtime that has not
    // finished exiting - leaves nothing behind. Without this, a throw here left the lifecycle reaper
    // running and the media server bound, and the host's next start attempt collided with the
    // resources the failed one still owned.
    const rollback: Array<() => Promise<void> | void> = [];
    try {
      this.lifecycle.start();
      rollback.push(() => this.lifecycle.stop());

      // Calling is optional: a messages-only connection is a valid shape, and demanding a calling
      // secret here is what made the Managed Bot quickstart start nothing.
      if (this.cfg.media.sharedSecret) {
        await this.media.start();
        rollback.push(() => this.media.stop());
      } else {
        this.log.warn("[msteams-call] no calling secret - messages lane only, calls will not be answered");
      }

      if (this.cfg.managedChat.enabled) {
        const chat = new ManagedChatServer(this.cfg.managedChat, {
          respond: (message) => this.respondToManagedChat(message),
          log: this.log,
        });
        await chat.start();
        this.managedChat = chat;
        rollback.push(async () => {
          await chat.stop();
          this.managedChat = undefined;
        });
      }
    } catch (err) {
      for (const undo of rollback.reverse()) {
        try {
          await undo();
        } catch {
          // A rollback step that itself fails must not mask the ORIGINAL startup error, which is the
          // one that says what actually went wrong.
        }
      }
      throw err;
    }
    this.log.info(`[msteams-call] started (mode=${this.mode})`);
  }

  async stop(): Promise<void> {
    for (const t of this.pendingOutboundTimers.values()) clearTimeout(t);
    this.pendingOutboundTimers.clear();
    this.pendingOutbound.clear();
    for (const call of this.calls.values()) call.close("shutdown");
    this.calls.clear();
    this.lifecycle.stop();
    await this.media.stop();
    await this.managedChat?.stop();
    this.managedChat = undefined;
  }

  /**
   * One agent turn for a managed-chat message (4.8): the same embedded consult the voice paths use,
   * on a per-conversation session key so a Teams chat keeps its context across messages. The consult
   * runs isolated from voice calls; attachments arrive by reference (the message text names them -
   * fetching them into the turn is a follow-up).
   */
  private async respondToManagedChat(message: ManagedInbound): Promise<string> {
    const cfg = this.api.config as unknown as OpenClawConfig;
    const attachmentNote = (message.attachments ?? [])
      .map((a) =>
        a.relayable === false
          ? `[attachment not relayed: ${a.name ?? a.kind}]`
          : `[attachment ${a.kind}: ${a.name ?? "unnamed"} at ${a.url}]`,
      )
      .join("\n");
    // Card submits arrive with EMPTY text and the payload in cardAction (protocol v1 additive field):
    // fold the payload into the question so a button press is a meaningful turn, not "I didn't catch
    // that". The agent authored the card, so its own field names give the payload meaning.
    const cardActionNote = message.cardAction
      ? `[card button pressed - submit payload: ${JSON.stringify(message.cardAction)}]`
      : "";
    const question = [message.text, cardActionNote, attachmentNote].filter(Boolean).join("\n");
    // 4.7's agent-side leg: fetch relayable images from their gateway-signed URLs into the consult, so
    // the agent SEES a pasted screenshot instead of reading a URL it cannot open (best-effort; the
    // text names every attachment either way).
    // Pin fetches to the gateway we reply through. The attachment URL is gateway-signed, but that
    // signature is verified BY the gateway - it proves nothing on this side, and the URL arrives
    // inside a message. The origin pin is what actually bounds where this fetch can go.
    const images = await fetchAttachmentImages(message.attachments, {
      gatewayOrigin: this.cfg.managedChat.gatewayReplyUrl,
    });
    if (images.length) {
      // The images param is ADDITIVE - openclaw hosts up to 2026.6.10 type the consult
      // without it, and a host that ignores it answers from the attachment NOTE alone. The log line
      // is the tell when verifying a host upgrade.
      this.log.info(`[msteams-chat] attaching ${images.length} image(s) to the consult (ignored by hosts without consult image support)`);
    }
    const result = await consultRealtimeVoiceAgent({
      cfg,
      agentRuntime: this.api.runtime.agent,
      ...(images.length ? { images } : {}),
      logger: { warn: (m: string) => this.log.warn(m) },
      ...(this.cfg.voice.agentId ? { agentId: this.cfg.voice.agentId } : {}),
      sessionKey: `msteams-chat:${message.tenantId}:${message.conversationId}`,
      messageProvider: "msteams",
      lane: "chat",
      runIdPrefix: "msteams-chat",
      args: { question },
      transcript: [],
      surface: "Microsoft Teams chat (StandIn managed)",
      userLabel: message.sender.displayName ?? "User",
      assistantLabel: "Agent",
      // The consult's default framing is a VOICE sidebar ("answer briefly, you are being
      // read aloud"); this is a persistent text chat - markdown renders, brevity is not a constraint,
      // and D12 self-disclosure applies.
      extraSystemPrompt:
        "You are answering in a Microsoft Teams text chat as the user's AI teammate (StandIn). " +
        "Write normal chat messages: Teams-flavored markdown is fine, match the length the question " +
        "deserves, and there is no text-to-speech constraint. If asked, be clear that you are an AI " +
        "assistant. Attachments are described in the message text; images may be attached directly.",
      fallbackText: "",
    });
    return result.text;
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
      throw new Error("msteams-call: outbound calling is disabled (set outbound.enabled)");
    if (!ob.workerBaseUrl)
      throw new Error("msteams-call: outbound.workerBaseUrl is not configured");
    if (!ob.tenantId) throw new Error("msteams-call: outbound.tenantId is not configured");
    if (!this.cfg.media.sharedSecret)
      throw new Error("msteams-call: secret is not configured");
    const userObjectId = to.replace(/^user:/i, "").trim();
    if (!userObjectId) throw new Error("msteams-call: target userObjectId (to) is required");
    if (this.lifecycle.activeCount() >= this.cfg.limits.maxConcurrentCalls)
      throw new Error("msteams-call: max concurrent calls reached; not placing outbound call");

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
          // Both header pairs during the X-StandIn-* transition: old workers verify
          // only the legacy names, new ones prefer the StandIn names.
          "x-standin-timestamp": String(timestampMs),
          "x-standin-signature": signature,
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
          `msteams-call: worker returned ${response.status}${text ? ` — ${text.slice(0, 200)}` : ""}`,
        );
      }
      const payload = (await response.json().catch(() => ({}))) as { callId?: string };
      workerCallId = payload.callId;
    } finally {
      await release();
    }
    if (!workerCallId)
      throw new Error("msteams-call: worker response did not include a callId");

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
      `[msteams-call] outbound call placed callId=${workerCallId} -> ${userObjectId} (${mode})`,
    );
    return { callId: workerCallId };
  }

  /** The worker told us how a call we placed really ended.
   *
   * Without this route the plugin only ever learned "no answer", and only after its own timeout - so a
   * DECLINED call and an unanswered one were indistinguishable, and both took the full answer-timeout
   * to resolve. Hermes has consumed this signal for a while; the two agents therefore behaved
   * differently for the same call, which is the divergence rather than either one being wrong.
   *
   * "answered" is not terminal: the media socket attaches and onSessionStart owns it from there. */
  private onCallOutcome(callId: string, outcome: string): void {
    if (outcome === "answered") return;
    if (!this.pendingOutbound.has(callId)) return;   // already finalized, or never ours
    this.pendingOutbound.delete(callId);
    this.clearOutboundTimer(callId);
    this.log.info(`[msteams-call] outbound call ${callId} ended as ${outcome} (worker outcome)`);
    // The worker already knows the call is over, so unlike the timeout path there is nothing to cancel.
    this.lifecycle.end(callId, WORKER_OUTCOME_TO_END_REASON[outcome] ?? "no-answer");
  }

  private finalizeUnansweredOutbound(callId: string): void {
    if (!this.pendingOutbound.has(callId)) return;
    this.pendingOutbound.delete(callId);
    this.clearOutboundTimer(callId);
    this.log.warn(
      `[msteams-call] outbound call ${callId} not answered within timeout; finalizing (no-answer/voicemail)`,
    );
    // H7a: the call may still be RINGING on the Teams worker side, and an unanswered outbound never
    // got a threadId (so DELETE /Calls?threadId can't reach it). Best-effort cancel it by callId via
    // the worker's cancel-by-callId endpoint so the callee stops ringing. Fire-and-forget: a late
    // answer is still denied in onSessionStart, and a cancel failure must never break finalization.
    void this.cancelRingingOutbound(callId);
    this.lifecycle.end(callId, "no-answer");
  }

  /**
   * H7a: cancel a ringing/active outbound call by its worker callId (DELETE /api/calls/{callId}),
   * signed exactly like {@link placeCall} (HMAC-SHA256 over `${timestampMs}.${callId}`). Best-effort:
   * logs and swallows every failure (never throws) — the outbound is already being finalized locally.
   */
  private async cancelRingingOutbound(callId: string): Promise<void> {
    const ob = this.cfg.outbound;
    const workerBaseUrl = ob?.workerBaseUrl;
    const sharedSecret = this.cfg.media.sharedSecret;
    if (!workerBaseUrl || !sharedSecret) return; // outbound cancel not configured — nothing to do
    try {
      const timestampMs = Date.now();
      const signature = createHmac("sha256", sharedSecret)
        .update(`${timestampMs}.${callId}`)
        .digest("hex");
      const url = `${workerBaseUrl.replace(/\/+$/, "")}/api/calls/${encodeURIComponent(callId)}`;
      const { response, release } = await fetchWithSsrFGuard({
        url,
        init: {
          method: "DELETE",
          headers: {
            "x-standin-timestamp": String(timestampMs),
            "x-standin-signature": signature,
            "x-openclawteamsbridge-timestamp": String(timestampMs),
            "x-openclawteamsbridge-signature": signature,
          },
        },
        // Same trust posture as placeCall: operator-configured worker, often on loopback.
        policy: { allowedHostnames: [new URL(url).hostname], allowPrivateNetwork: true },
      });
      try {
        if (!response.ok) {
          this.log.warn(
            `[msteams-call] cancel-by-callId ${callId} returned ${response.status}`,
          );
        } else {
          this.log.info(`[msteams-call] cancelled ringing outbound ${callId}`);
        }
      } finally {
        await release();
      }
    } catch (err) {
      this.log.warn(
        `[msteams-call] cancel-by-callId ${callId} failed: ${(err as Error).message}`,
      );
    }
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
    // Make this call resolvable by the registered OpenClaw tool for the duration of the call.
    this.trackManagedCall(session, true);
    // Realtime mode requires a resolved provider; streaming mode does not.
    if (this.mode === "realtime" && !this.realtime?.provider) {
      this.log.error("[msteams-call] no realtime voice provider resolved — rejecting call");
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

    // A late answer to an outbound call whose answer-timeout already fired: the pending entry is
    // gone and the lifecycle record is a terminal outbound. Without this guard it falls through to
    // the inbound branch and is evaluated against inbound policy using the CALLEE's id — the wrong
    // reject/greeting. Deny the late media attach instead of mis-routing it.
    const prior = this.lifecycle.getStatus(session.callId);
    if (prior?.isTerminal) {
      const rec = this.lifecycle.getRecord(session.callId);
      this.log.warn(
        `[msteams-call] ignoring late media attach for ${session.callId} — call already finalized` +
          ` (${rec?.endReason ?? "ended"}); closing`,
      );
      session.close(rec?.direction === "outbound" ? "answer-timeout" : "already-ended");
      return;
    }

    // Inbound: enforce caller policy before accepting.
    const from = session.caller?.aadId ?? "";
    if (!isInboundCallAllowed(this.cfg.voice.inboundPolicy, this.cfg.voice.allowFrom, from)) {
      this.log.warn(
        `[msteams-call] ${describeInboundRejection(this.cfg.voice.inboundPolicy, from)}`,
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
      this.log.warn(`[msteams-call] cannot accept call: ${String(err)}`);
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

  /**
   * The in-call "post this to the Teams chat" sender, or undefined when this call cannot have one.
   *
   * post_meeting_minutes delivers through the HOST's message tool, which needs the customer's own
   * Teams channel - a managed customer has none, so on that tier posting to chat could only fail.
   * This uses the messages lane the plugin is already serving, over the gateway that already carries
   * its chat replies.
   */
  /**
   * sessionKey -> the live managed call it belongs to.
   *
   * An OpenClaw tool is registered once and global, but "post to the chat" only means something
   * inside a particular call. The consult carries a sessionKey (both the streaming and realtime paths
   * build it the same way), and the tool context hands that back - so this is what turns a global tool
   * into a per-call one without inventing a second context mechanism.
   */
  private readonly managedCallBySession = new Map<string, { tenantId: string; conversationId: string }>();

  /** callId -> the session key it registered under, so teardown removes the right entry. */
  private readonly sessionKeyByCall = new Map<string, string>();

  /** Drop a call from the registry when it ends. */
  private forgetManagedCall(callId: string): void {
    const key = this.sessionKeyByCall.get(callId);
    if (!key) return;
    this.sessionKeyByCall.delete(callId);
    this.managedCallBySession.delete(key);
  }

  /** Look up the poster for a consult running under this session key. Used by the registered tool. */
  chatPosterForSession(sessionKey: string): ((text: string) => Promise<boolean>) | undefined {
    const call = this.managedCallBySession.get(sessionKey);
    const chat = this.cfg.managedChat;
    if (!call || !chat.enabled || !chat.chatSecret) return undefined;
    return async (text: string) =>
      postManagedMessage({
        chatSecret: chat.chatSecret,
        gatewayReplyUrl: chat.gatewayReplyUrl,
        tenantId: call.tenantId,
        conversationId: call.conversationId,
        text,
        idempotencyKey: `sess-${sessionKey}-${createHash("sha256").update(text).digest("hex").slice(0, 12)}`,
      });
  }

  /** Register/forget a managed call against the STREAMING consult's session key.
   *
   * Streaming only, on purpose: the realtime path dispatches post_chat_message itself through its own
   * onToolCall and already has the poster injected, so it needs nothing here. Streaming has no tool
   * dispatch at all - its agent turn IS an OpenClaw consult - which is why that path has to go through
   * a registered OpenClaw tool and the consult tool policy instead. */
  private trackManagedCall(session: MsteamsSession, add: boolean): void {
    const tenantId = session.tenantId;
    // Only a real Teams conversation can be posted into. A 1:1 call has none of its own - the worker
    // sends threadId=callId as a fallback, which Teams cannot resolve - so those calls simply do not
    // get the tool rather than getting one that fails.
    if (!tenantId || !session.threadId?.startsWith("19:")) return;
    const key = this.streamingSessionKey(session);
    if (add) {
      this.managedCallBySession.set(key, { tenantId, conversationId: session.threadId });
      // ALSO register the poster the global tool resolves through. It used to be registered only by
      // buildChatPoster, which runs on the realtime path alone - so in STREAMING mode the tool was
      // offered and could never work: postableCalls was empty, and every call came back "there is no
      // active Teams call with a chat to post into". Streaming is the mode that NEEDS the global tool,
      // because it has no tool dispatch of its own. Registering here covers both modes, from the one
      // place that already runs for both.
      const poster = this.buildChatPoster(session);
      if (poster) this.postableCalls.set(session.callId, { conversationId: session.threadId, post: poster });
      // Remember the key BY CALL ID: sessionScope can make the key per-thread or per-AAD, so it is not
      // derivable from a call id alone at teardown, and a stale entry would let a later consult in the
      // same scope post into a call that has ended.
      this.sessionKeyByCall.set(session.callId, key);
    } else {
      this.managedCallBySession.delete(key);
      this.postableCalls.delete(session.callId);
    }
  }

  private buildChatPoster(session: MsteamsSession): ((text: string) => Promise<boolean>) | undefined {
    const chat = this.cfg.managedChat;
    const tenantId = session.tenantId;
    if (!chat.enabled || !chat.chatSecret || !tenantId || !session.threadId) return undefined;
    // A 1:1 call has no Teams conversation of its own - the worker falls back to the call id, which
    // the gateway cannot address (it 404s rather than acknowledging). Only a real conversation
    // ("19:...") can be posted into, so anything else is not postable and the tool is not offered.
    if (!session.threadId.startsWith("19:")) return undefined;
    const poster = async (text: string) =>
      postManagedMessage({
        chatSecret: chat.chatSecret,
        gatewayReplyUrl: chat.gatewayReplyUrl,
        tenantId,
        conversationId: session.threadId,
        text,
        // Scoped to the call AND the content, so a retried tool call does not double-post while two
        // genuinely different posts in one call both go out.
        idempotencyKey: `call-${session.callId}-${createHash("sha256").update(text).digest("hex").slice(0, 12)}`,
      });
    // Registration into postableCalls lives in trackManagedCall, which runs for BOTH modes. Doing it
    // here meant the global tool only ever saw realtime calls.
    return poster;
  }

  /** The one live call a global tool may post into, or a reason it cannot decide. */
  resolvePostableCall(): { post: (text: string) => Promise<boolean> } | { error: string } {
    const entries = [...this.postableCalls.values()];
    if (entries.length === 1) return entries[0];
    if (entries.length === 0) {
      return {
        error:
          "There is no active Teams call with a chat to post into. This works during a meeting call " +
          "on a StandIn managed connection; a 1:1 call has no chat thread of its own.",
      };
    }
    return {
      error:
        `There are ${entries.length} calls in progress, so I cannot tell which chat you mean. ` +
        "Ask me again when only one call is active.",
    };
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
      this.log.info(`[msteams-call] streaming STT provider: ${res.provider.id}`);
    } else {
      this.log.info(
        `[msteams-call] no streaming STT provider resolved (${res.code}); using file-based STT fallback`,
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
        const tmp = path.join(os.tmpdir(), `msteams-call-${session.callId}-${this.sttSeq++}.wav`);
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
          cfg,
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
          // The consult tool policy decides what this turn may call. Append the in-call chat tool so
          // STREAMING mode can post to the Teams chat - the realtime path has its own dispatch for
          // this, streaming has none, and the policy is the mechanism that exists for exactly this.
          // The tool factory returns null unless this session is a live managed call, so allowing the
          // name here costs nothing on a BYO or 1:1 call: it simply is not there to call.
          toolsAllow: [
            ...(resolveRealtimeVoiceAgentConsultToolsAllow(this.cfg.voice.realtime.toolPolicy) ?? []),
            MSTEAMS_POST_CHAT_TOOL_NAME,
          ],
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
      // MANAGED only. Wired when the messages lane is configured AND the session carries the tenant
      // this call belongs to - both are required to address a post, and neither exists on BYO/free.
      // Undefined here means the tool is not offered at all, which is better than offering it and
      // failing: the model would otherwise promise the caller something that cannot happen.
      postChatMessage: this.buildChatPoster(session),
      inboundPolicy: this.cfg.voice.inboundPolicy,
      allowFrom: this.cfg.voice.allowFrom,
      requireRecordingStatus: this.cfg.voice.msteams?.requireRecordingStatus,
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
    this.forgetManagedCall(info.callId);
    // Caller-driven hangup: the Teams worker session is already closing, so tear down locally
    // (close() with no reason) and end the lifecycle record.
    this.disposeCall(info.callId);
    this.lifecycle.end(info.callId, "hangup-user");
  }

  /**
   * Drop every per-call resource: the media/realtime bridge, the outbound bookkeeping, and the
   * retained vision frames. Shared by a caller hangup ({@link onSessionEnd}) and the reaper's onReap
   * hook so a reaped call is torn down exactly like a hangup instead of leaking a zombie socket.
   * `closeReason` is forwarded to the bridge's close(): pass a reason (reaper) to ALSO close the
   * Teams worker session; omit it (caller hangup) when the session is already closing.
   */
  private disposeCall(callId: string, closeReason?: string): void {
    this.pendingOutbound.delete(callId);
    this.clearOutboundTimer(callId);
    this.calls.get(callId)?.close(closeReason);
    this.calls.delete(callId);
    this.postableCalls.delete(callId);
    // Release the per-call vision frames (latest + keyframe history, ~1-2 MB/call). These were never
    // released outside tests, leaking for the process lifetime on every completed call.
    this.vision.release(callId);
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
