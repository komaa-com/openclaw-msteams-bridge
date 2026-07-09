// MsteamsVoiceRuntime — self-contained orchestration that replaces voice-call's MsteamsProvider +
// CallManager for the standalone plugin. Owns the Teams media WebSocket, drives CallLifecycle, and
// bridges each call to the realtime voice model via createMsteamsRealtimeCall.
//
// Scope: realtime speech-to-speech (inbound + outbound call-backs via worker place-call) and the
// streaming STT→agent→TTS path. Still out of scope: getCallStatus as a VoiceCallProvider — not
// needed for a Teams realtime assistant and would re-introduce the heavier provider surface.
import { consultRealtimeVoiceAgent, resolveConfiguredRealtimeVoiceProvider, resolveRealtimeVoiceAgentConsultToolsAllow, } from "openclaw/plugin-sdk/realtime-voice";
import { getRealtimeTranscriptionProvider, listRealtimeTranscriptionProviders, } from "openclaw/plugin-sdk/realtime-transcription";
import { resolveConfiguredCapabilityProvider } from "openclaw/plugin-sdk/provider-selection-runtime";
import { fetchWithSsrFGuard } from "openclaw/plugin-sdk/ssrf-runtime";
import { createHmac } from "node:crypto";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import { describeInboundRejection, isInboundCallAllowed } from "./allowlist.js";
import { CallLifecycle } from "./call-lifecycle.js";
import { resolveGroupCallGateConfig } from "./group-call-gate.js";
import { collectLatestFrameImages } from "./vision-consult.js";
import { MSTEAMS_PCM_SAMPLE_RATE_HZ, MsteamsMediaStream, } from "./msteams-media-stream.js";
import { createMsteamsRealtimeCall, } from "./msteams-realtime.js";
import { createMsteamsStreamingCall } from "./msteams-streaming.js";
import { createMsteamsTtsProvider } from "./msteams-tts.js";
import { MsteamsVisionStore } from "./msteams-vision-store.js";
import { resolveVoiceResponseModel } from "./response-model.js";
import { VisionBudget } from "./vision-budget.js";
/** Default no-answer guard for a placed outbound call (overridable via outbound.answerTimeoutMs). */
const OUTBOUND_ANSWER_TIMEOUT_DEFAULT_MS = 120_000;
export class MsteamsVoiceRuntime {
    api;
    cfg;
    lifecycle;
    media;
    vision;
    visionBudget;
    calls = new Map();
    log;
    realtime;
    /** Resolved streaming STT provider (mode:"streaming"); undefined → file-based STT fallback. */
    transcription;
    /** Selected voice path; finalized in start() once the realtime provider is resolved. */
    mode = "realtime";
    /** Lazily-built TTS provider for the streaming path (api.runtime.tts). */
    ttsProvider;
    /** Monotonic suffix for streaming STT temp-file names. */
    sttSeq = 0;
    /** Calls we placed via the worker, awaiting their media WS session.start to attach. */
    pendingOutbound = new Map();
    pendingOutboundTimers = new Map();
    constructor(api, cfg) {
        this.api = api;
        this.cfg = cfg;
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
        this.lifecycle = new CallLifecycle({
            // In-memory keyed store for call records. api.runtime.state.openSyncKeyedStore is gated to
            // trusted (bundled/official) plugins, which a third-party npm/ClawHub install is not. Call
            // records are ephemeral — a gateway restart drops live media calls anyway — so an in-process
            // Map is sufficient and keeps the plugin installable as an untrusted plugin.
            openSyncKeyedStore: (_name) => {
                const m = new Map();
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
        }, {
            maxConcurrentCalls: cfg.limits.maxConcurrentCalls,
            maxDurationMs: cfg.limits.maxDurationMs,
            staleCallReaperMs: cfg.limits.staleCallReaperMs,
            // The reaper only ends the lifecycle record; run the SAME runtime teardown as a user hangup so
            // the reaped call's media + realtime sockets actually close (H7: no zombie, no maxConcurrentCalls
            // bypass). Pass the reason so the Teams worker session is closed too (the call is still live,
            // unlike a caller-driven hangup where the session is already closing).
            onReap: (callId, reason) => this.disposeCall(callId, reason),
        });
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
    async start() {
        this.realtime = resolveConfiguredRealtimeVoiceProvider({
            configuredProviderId: this.cfg.voice.realtime.provider,
            providerConfigs: this.cfg.voice.realtime.providers,
            cfg: this.api.config,
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
        });
        // "realtime" when a realtime provider resolved (or explicitly configured), else "streaming".
        this.mode =
            this.cfg.voice.mode ?? (this.realtime?.provider ? "realtime" : "streaming");
        if (this.mode === "streaming")
            this.resolveTranscriptionProvider();
        this.lifecycle.start();
        await this.media.start();
        this.log.info(`[msteams-voice] started (mode=${this.mode})`);
    }
    async stop() {
        for (const t of this.pendingOutboundTimers.values())
            clearTimeout(t);
        this.pendingOutboundTimers.clear();
        this.pendingOutbound.clear();
        for (const call of this.calls.values())
            call.close("shutdown");
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
    async placeCall(to, opts) {
        const ob = this.cfg.outbound;
        if (!ob?.enabled)
            throw new Error("msteams-voice: outbound calling is disabled (set outbound.enabled)");
        if (!ob.workerBaseUrl)
            throw new Error("msteams-voice: outbound.workerBaseUrl is not configured");
        if (!ob.tenantId)
            throw new Error("msteams-voice: outbound.tenantId is not configured");
        if (!this.cfg.media.sharedSecret)
            throw new Error("msteams-voice: sharedSecret is not configured");
        const userObjectId = to.replace(/^user:/i, "").trim();
        if (!userObjectId)
            throw new Error("msteams-voice: target userObjectId (to) is required");
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
        let workerCallId;
        try {
            if (!response.ok) {
                const text = await response.text().catch(() => "");
                throw new Error(`msteams-voice: worker returned ${response.status}${text ? ` — ${text.slice(0, 200)}` : ""}`);
            }
            const payload = (await response.json().catch(() => ({})));
            workerCallId = payload.callId;
        }
        finally {
            await release();
        }
        if (!workerCallId)
            throw new Error("msteams-voice: worker response did not include a callId");
        const mode = opts?.mode ?? ob.defaultMode ?? "notify";
        this.lifecycle.initiate({
            callId: workerCallId,
            providerCallId: workerCallId,
            direction: "outbound",
            from: "",
            to,
            message: opts?.message,
        });
        this.pendingOutbound.set(workerCallId, { to, message: opts?.message, mode });
        const timer = setTimeout(() => this.finalizeUnansweredOutbound(workerCallId), ob.answerTimeoutMs ?? OUTBOUND_ANSWER_TIMEOUT_DEFAULT_MS);
        timer.unref?.();
        this.pendingOutboundTimers.set(workerCallId, timer);
        this.log.info(`[msteams-voice] outbound call placed callId=${workerCallId} -> ${userObjectId} (${mode})`);
        return { callId: workerCallId };
    }
    finalizeUnansweredOutbound(callId) {
        if (!this.pendingOutbound.has(callId))
            return;
        this.pendingOutbound.delete(callId);
        this.clearOutboundTimer(callId);
        this.log.warn(`[msteams-voice] outbound call ${callId} not answered within timeout; finalizing (no-answer/voicemail)`);
        this.lifecycle.end(callId, "no-answer");
    }
    clearOutboundTimer(callId) {
        const t = this.pendingOutboundTimers.get(callId);
        if (t)
            clearTimeout(t);
        this.pendingOutboundTimers.delete(callId);
    }
    /**
     * Query the current status of a call (state + whether it has reached a terminal state).
     * Returns undefined for an unknown call id. Note: OpenClaw exposes no provider-status registration
     * hook for non-channel plugins, so this is surfaced as a runtime method (callable by an embedding
     * host or a future admin/tool surface) rather than a registered VoiceCallProvider.getCallStatus.
     */
    getCallStatus(callId) {
        return this.lifecycle.getStatus(callId);
    }
    onSessionStart(session) {
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
            const greeting = pending.mode === "notify" && pending.message
                ? `Deliver this message to the person, then say goodbye and end the call: "${pending.message}"`
                : (pending.message ?? this.cfg.voice.inboundGreeting);
            this.calls.set(session.callId, this.createCall(session, greeting));
            return;
        }
        // Inbound: enforce caller policy before accepting.
        const from = session.caller?.aadId ?? "";
        if (!isInboundCallAllowed(this.cfg.voice.inboundPolicy, this.cfg.voice.allowFrom, from)) {
            this.log.warn(`[msteams-voice] ${describeInboundRejection(this.cfg.voice.inboundPolicy, from)}`);
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
        }
        catch (err) {
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
    createCall(session, greeting) {
        if (this.mode === "streaming") {
            return createMsteamsStreamingCall({ session, deps: this.buildStreamingDeps(session, greeting) });
        }
        return createMsteamsRealtimeCall({
            session,
            deps: this.buildDeps(session, this.realtime?.provider, greeting),
        });
    }
    getTtsProvider() {
        if (!this.ttsProvider) {
            this.ttsProvider = createMsteamsTtsProvider({
                coreConfig: this.api.config,
                ttsOverride: this.cfg.voice.tts,
                runtime: this.api.runtime.tts,
                logger: { warn: (m) => this.log.warn(m) },
            });
        }
        return this.ttsProvider;
    }
    /** Consult session key, honoring sessionScope (mirrors the realtime path). */
    streamingSessionKey(session) {
        const scope = this.cfg.voice.sessionScope;
        if (scope === "per-call")
            return `msteams:${session.callId}`;
        if (scope === "per-thread")
            return `msteams:${session.threadId || session.callId}`;
        return `msteams:${session.caller?.aadId || session.callId}`;
    }
    /**
     * Resolve a streaming STT provider (mode:"streaming"). Auto-selects from the host's configured
     * realtime transcription providers when `stt.provider` is omitted; leaves `this.transcription`
     * undefined (→ file-based STT fallback) when none resolve.
     */
    resolveTranscriptionProvider() {
        const cfg = this.api.config;
        const res = resolveConfiguredCapabilityProvider({
            configuredProviderId: this.cfg.voice.stt?.provider,
            providerConfigs: this.cfg.voice.stt?.providers,
            cfg,
            cfgForResolve: cfg,
            getConfiguredProvider: (id) => getRealtimeTranscriptionProvider(id, cfg),
            listProviders: () => listRealtimeTranscriptionProviders(cfg),
            resolveProviderConfig: ({ provider, cfg: c, rawConfig }) => provider.resolveConfig?.({ cfg: c, rawConfig }) ?? rawConfig,
            isProviderConfigured: ({ provider, cfg: c, providerConfig }) => provider.isConfigured({ cfg: c, providerConfig }),
        });
        if (res.ok) {
            this.transcription = { provider: res.provider, providerConfig: res.providerConfig };
            this.log.info(`[msteams-voice] streaming STT provider: ${res.provider.id}`);
        }
        else {
            this.log.info(`[msteams-voice] no streaming STT provider resolved (${res.code}); using file-based STT fallback`);
        }
    }
    buildStreamingDeps(session, greeting) {
        const cfg = this.api.config;
        const agentRuntime = this.api.runtime.agent;
        const transcription = this.transcription;
        return {
            // Preferred: a live streaming STT session (lower latency) when a provider resolved.
            ...(transcription
                ? {
                    createTranscriptionSession: (callbacks) => transcription.provider.createSession({
                        cfg,
                        providerConfig: transcription.providerConfig,
                        ...callbacks,
                    }),
                }
                : {}),
            // Fallback: VAD-segmented utterance → temp WAV → file STT (used when no provider resolved).
            transcribe: async (pcm16k) => {
                const wav = pcmToWav(pcm16k, MSTEAMS_PCM_SAMPLE_RATE_HZ);
                const tmp = path.join(os.tmpdir(), `msteams-voice-${session.callId}-${this.sttSeq++}.wav`);
                await fs.writeFile(tmp, wav);
                try {
                    const res = await this.api.runtime.mediaUnderstanding.transcribeAudioFile({
                        filePath: tmp,
                        cfg,
                    });
                    return res.text ?? "";
                }
                finally {
                    await fs.unlink(tmp).catch(() => { });
                }
            },
            consult: async ({ question, transcript, images }) => {
                const { provider, model } = resolveVoiceResponseModel({
                    voiceConfig: this.cfg.voice,
                    agentRuntime,
                });
                const thinkLevel = this.cfg.voice.realtime.consultThinkingLevel ??
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
            getVisionImages: () => collectLatestFrameImages({
                getLatestFrame: (s) => this.vision.getLatest(session.callId, s),
                visionBudget: this.visionBudget,
                callId: session.callId,
            }),
            appendTranscript: (e) => this.lifecycle.appendTranscript(session.callId, e),
            logger: this.log,
        };
    }
    buildDeps(session, provider, greetingInstructions) {
        return {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            provider: provider,
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            providerConfig: this.realtime?.providerConfig,
            cfg: this.api.config,
            instructions: this.cfg.voice.realtime.instructions,
            greetingInstructions,
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
    onSessionEnd(info) {
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
    disposeCall(callId, closeReason) {
        this.pendingOutbound.delete(callId);
        this.clearOutboundTimer(callId);
        this.calls.get(callId)?.close(closeReason);
        this.calls.delete(callId);
        // Release the per-call vision frames (latest + keyframe history, ~1-2 MB/call). These were never
        // released outside tests, leaking for the process lifetime on every completed call.
        this.vision.release(callId);
    }
}
/** Wrap raw PCM (16-bit mono LE) in a minimal WAV container so file-based STT can read it. */
function pcmToWav(pcm, sampleRate) {
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
