import { MSTEAMS_POST_CHAT_TOOL_NAME } from "./msteams-realtime-tools.js";
import { fetchAttachmentImages, ManagedChatServer, postManagedMessage, } from "./managed-chat.js";
import { consultRealtimeVoiceAgent, resolveConfiguredRealtimeVoiceProvider, resolveRealtimeVoiceAgentConsultToolsAllow, } from "openclaw/plugin-sdk/realtime-voice";
import { getRealtimeTranscriptionProvider, listRealtimeTranscriptionProviders, } from "openclaw/plugin-sdk/realtime-transcription";
import { resolveConfiguredCapabilityProvider } from "openclaw/plugin-sdk/provider-selection-runtime";
import { fetchWithSsrFGuard } from "openclaw/plugin-sdk/ssrf-runtime";
import { createHash } from "node:crypto";
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
const WORKER_OUTCOME_TO_END_REASON = {
    "no-answer": "no-answer",
    declined: "no-answer",
    busy: "no-answer",
    failed: "error",
};
const OUTBOUND_ANSWER_TIMEOUT_DEFAULT_MS = 120_000;
export class MsteamsVoiceRuntime {
    api;
    cfg;
    lifecycle;
    media;
    vision;
    visionBudget;
    calls = new Map();
    postableCalls = new Map();
    log;
    realtime;
    transcription;
    mode = "realtime";
    ttsProvider;
    sttSeq = 0;
    pendingOutbound = new Map();
    pendingOutboundTimers = new Map();
    managedChat;
    constructor(api, cfg) {
        this.api = api;
        this.cfg = cfg;
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
        this.lifecycle = new CallLifecycle({
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
            onCallOutcome: (i) => this.onCallOutcome(i.callId, i.outcome),
            onDtmf: (i) => this.calls.get(i.callId)?.notifyDtmf(i.digit),
            onParticipants: (i) => this.calls.get(i.callId)?.setHumanCount(i.count),
            onAssistantSay: (i) => this.calls.get(i.callId)?.say(i.text),
        });
    }
    async start() {
        this.realtime = resolveConfiguredRealtimeVoiceProvider({
            configuredProviderId: this.cfg.voice.realtime.provider,
            providerConfigs: this.cfg.voice.realtime.providers,
            cfg: this.api.config,
        });
        this.mode =
            this.cfg.voice.mode ?? (this.realtime?.provider ? "realtime" : "streaming");
        if (this.mode === "realtime" && !this.realtime?.provider) {
            const providerId = this.cfg.voice.realtime.provider;
            this.log.warn(`[msteams-call] mode is "realtime" but no realtime voice provider resolved` +
                (providerId
                    ? ` (configured provider "${providerId}" has no usable credentials)`
                    : ` (no realtime provider configured)`) +
                `. Every inbound call will be rejected with "realtime-unavailable". Set the provider's API` +
                ` key, or set mode:"streaming".`);
        }
        if (this.mode === "streaming")
            this.resolveTranscriptionProvider();
        if (this.cfg.managedChat.configuredWithoutSecret) {
            this.log.warn("[msteams-call] managedBot.enabled is set but no secret resolved - the messages lane is OFF. " +
                "Set `secret` to the value StandIn shows you; it covers calling and messages both.");
        }
        const rollback = [];
        try {
            this.lifecycle.start();
            rollback.push(() => this.lifecycle.stop());
            if (this.cfg.media.sharedSecret) {
                await this.media.start();
                rollback.push(() => this.media.stop());
            }
            else {
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
        }
        catch (err) {
            for (const undo of rollback.reverse()) {
                try {
                    await undo();
                }
                catch {
                }
            }
            throw err;
        }
        this.log.info(`[msteams-call] started (mode=${this.mode})`);
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
        await this.managedChat?.stop();
        this.managedChat = undefined;
    }
    async respondToManagedChat(message) {
        const cfg = this.api.config;
        const attachmentNote = (message.attachments ?? [])
            .map((a) => a.relayable === false
            ? `[attachment not relayed: ${a.name ?? a.kind}]`
            : `[attachment ${a.kind}: ${a.name ?? "unnamed"} at ${a.url}]`)
            .join("\n");
        const cardActionNote = message.cardAction
            ? `[card button pressed - submit payload: ${JSON.stringify(message.cardAction)}]`
            : "";
        const question = [message.text, cardActionNote, attachmentNote].filter(Boolean).join("\n");
        const images = await fetchAttachmentImages(message.attachments, {
            gatewayOrigin: this.cfg.managedChat.gatewayReplyUrl,
        });
        if (images.length) {
            this.log.info(`[msteams-chat] attaching ${images.length} image(s) to the consult (ignored by hosts without consult image support)`);
        }
        const result = await consultRealtimeVoiceAgent({
            cfg,
            agentRuntime: this.api.runtime.agent,
            ...(images.length ? { images } : {}),
            logger: { warn: (m) => this.log.warn(m) },
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
            extraSystemPrompt: "You are answering in a Microsoft Teams text chat as the user's AI teammate (StandIn). " +
                "Write normal chat messages: Teams-flavored markdown is fine, match the length the question " +
                "deserves, and there is no text-to-speech constraint. If asked, be clear that you are an AI " +
                "assistant. Attachments are described in the message text; images may be attached directly.",
            fallbackText: "",
        });
        return result.text;
    }
    async placeCall(to, opts) {
        const ob = this.cfg.outbound;
        if (!ob?.enabled)
            throw new Error("msteams-call: outbound calling is disabled (set outbound.enabled)");
        if (!ob.workerBaseUrl)
            throw new Error("msteams-call: outbound.workerBaseUrl is not configured");
        if (!ob.tenantId)
            throw new Error("msteams-call: outbound.tenantId is not configured");
        if (!this.cfg.media.sharedSecret)
            throw new Error("msteams-call: secret is not configured");
        const userObjectId = to.replace(/^user:/i, "").trim();
        if (!userObjectId)
            throw new Error("msteams-call: target userObjectId (to) is required");
        if (this.lifecycle.activeCount() >= this.cfg.limits.maxConcurrentCalls)
            throw new Error("msteams-call: max concurrent calls reached; not placing outbound call");
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
                    "x-standin-timestamp": String(timestampMs),
                    "x-standin-signature": signature,
                    "x-openclawteamsbridge-timestamp": String(timestampMs),
                    "x-openclawteamsbridge-signature": signature,
                },
                body: JSON.stringify({ userObjectId, tenantId: ob.tenantId }),
            },
            policy: { allowedHostnames: [new URL(url).hostname], allowPrivateNetwork: true },
        });
        let workerCallId;
        try {
            if (!response.ok) {
                const text = await response.text().catch(() => "");
                throw new Error(`msteams-call: worker returned ${response.status}${text ? ` — ${text.slice(0, 200)}` : ""}`);
            }
            const payload = (await response.json().catch(() => ({})));
            workerCallId = payload.callId;
        }
        finally {
            await release();
        }
        if (!workerCallId)
            throw new Error("msteams-call: worker response did not include a callId");
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
        this.log.info(`[msteams-call] outbound call placed callId=${workerCallId} -> ${userObjectId} (${mode})`);
        return { callId: workerCallId };
    }
    onCallOutcome(callId, outcome) {
        if (outcome === "answered")
            return;
        if (!this.pendingOutbound.has(callId))
            return;
        this.pendingOutbound.delete(callId);
        this.clearOutboundTimer(callId);
        this.log.info(`[msteams-call] outbound call ${callId} ended as ${outcome} (worker outcome)`);
        this.lifecycle.end(callId, WORKER_OUTCOME_TO_END_REASON[outcome] ?? "no-answer");
    }
    finalizeUnansweredOutbound(callId) {
        if (!this.pendingOutbound.has(callId))
            return;
        this.pendingOutbound.delete(callId);
        this.clearOutboundTimer(callId);
        this.log.warn(`[msteams-call] outbound call ${callId} not answered within timeout; finalizing (no-answer/voicemail)`);
        void this.cancelRingingOutbound(callId);
        this.lifecycle.end(callId, "no-answer");
    }
    async cancelRingingOutbound(callId) {
        const ob = this.cfg.outbound;
        const workerBaseUrl = ob?.workerBaseUrl;
        const sharedSecret = this.cfg.media.sharedSecret;
        if (!workerBaseUrl || !sharedSecret)
            return;
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
                policy: { allowedHostnames: [new URL(url).hostname], allowPrivateNetwork: true },
            });
            try {
                if (!response.ok) {
                    this.log.warn(`[msteams-call] cancel-by-callId ${callId} returned ${response.status}`);
                }
                else {
                    this.log.info(`[msteams-call] cancelled ringing outbound ${callId}`);
                }
            }
            finally {
                await release();
            }
        }
        catch (err) {
            this.log.warn(`[msteams-call] cancel-by-callId ${callId} failed: ${err.message}`);
        }
    }
    clearOutboundTimer(callId) {
        const t = this.pendingOutboundTimers.get(callId);
        if (t)
            clearTimeout(t);
        this.pendingOutboundTimers.delete(callId);
    }
    getCallStatus(callId) {
        return this.lifecycle.getStatus(callId);
    }
    onSessionStart(session) {
        this.trackManagedCall(session, true);
        if (this.mode === "realtime" && !this.realtime?.provider) {
            this.log.error("[msteams-call] no realtime voice provider resolved — rejecting call");
            session.close("realtime-unavailable");
            return;
        }
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
        const prior = this.lifecycle.getStatus(session.callId);
        if (prior?.isTerminal) {
            const rec = this.lifecycle.getRecord(session.callId);
            this.log.warn(`[msteams-call] ignoring late media attach for ${session.callId} — call already finalized` +
                ` (${rec?.endReason ?? "ended"}); closing`);
            session.close(rec?.direction === "outbound" ? "answer-timeout" : "already-ended");
            return;
        }
        const from = session.caller?.aadId ?? "";
        if (!isInboundCallAllowed(this.cfg.voice.inboundPolicy, this.cfg.voice.allowFrom, from)) {
            this.log.warn(`[msteams-call] ${describeInboundRejection(this.cfg.voice.inboundPolicy, from)}`);
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
            this.log.warn(`[msteams-call] cannot accept call: ${String(err)}`);
            session.close("busy");
            return;
        }
        this.lifecycle.answer(session.callId);
        this.calls.set(session.callId, this.createCall(session, this.cfg.voice.inboundGreeting));
    }
    createCall(session, greeting) {
        if (this.mode === "streaming") {
            return createMsteamsStreamingCall({ session, deps: this.buildStreamingDeps(session, greeting) });
        }
        return createMsteamsRealtimeCall({
            session,
            deps: this.buildDeps(session, this.realtime?.provider, greeting),
        });
    }
    managedCallBySession = new Map();
    sessionKeyByCall = new Map();
    forgetManagedCall(callId) {
        const key = this.sessionKeyByCall.get(callId);
        if (!key)
            return;
        this.sessionKeyByCall.delete(callId);
        this.managedCallBySession.delete(key);
    }
    chatPosterForSession(sessionKey) {
        const call = this.managedCallBySession.get(sessionKey);
        const chat = this.cfg.managedChat;
        if (!call || !chat.enabled || !chat.chatSecret)
            return undefined;
        return async (text) => postManagedMessage({
            chatSecret: chat.chatSecret,
            gatewayReplyUrl: chat.gatewayReplyUrl,
            tenantId: call.tenantId,
            conversationId: call.conversationId,
            text,
            idempotencyKey: `sess-${sessionKey}-${createHash("sha256").update(text).digest("hex").slice(0, 12)}`,
        });
    }
    trackManagedCall(session, add) {
        const tenantId = session.tenantId;
        if (!tenantId || !session.threadId?.startsWith("19:"))
            return;
        const key = this.streamingSessionKey(session);
        if (add) {
            this.managedCallBySession.set(key, { tenantId, conversationId: session.threadId });
            const poster = this.buildChatPoster(session);
            if (poster)
                this.postableCalls.set(session.callId, { conversationId: session.threadId, post: poster });
            this.sessionKeyByCall.set(session.callId, key);
        }
        else {
            this.managedCallBySession.delete(key);
            this.postableCalls.delete(session.callId);
        }
    }
    buildChatPoster(session) {
        const chat = this.cfg.managedChat;
        const tenantId = session.tenantId;
        if (!chat.enabled || !chat.chatSecret || !tenantId || !session.threadId)
            return undefined;
        if (!session.threadId.startsWith("19:"))
            return undefined;
        const poster = async (text) => postManagedMessage({
            chatSecret: chat.chatSecret,
            gatewayReplyUrl: chat.gatewayReplyUrl,
            tenantId,
            conversationId: session.threadId,
            text,
            idempotencyKey: `call-${session.callId}-${createHash("sha256").update(text).digest("hex").slice(0, 12)}`,
        });
        return poster;
    }
    resolvePostableCall() {
        const entries = [...this.postableCalls.values()];
        if (entries.length === 1)
            return entries[0];
        if (entries.length === 0) {
            return {
                error: "There is no active Teams call with a chat to post into. This works during a meeting call " +
                    "on a StandIn managed connection; a 1:1 call has no chat thread of its own.",
            };
        }
        return {
            error: `There are ${entries.length} calls in progress, so I cannot tell which chat you mean. ` +
                "Ask me again when only one call is active.",
        };
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
    streamingSessionKey(session) {
        const scope = this.cfg.voice.sessionScope;
        if (scope === "per-call")
            return `msteams:${session.callId}`;
        if (scope === "per-thread")
            return `msteams:${session.threadId || session.callId}`;
        return `msteams:${session.caller?.aadId || session.callId}`;
    }
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
            this.log.info(`[msteams-call] streaming STT provider: ${res.provider.id}`);
        }
        else {
            this.log.info(`[msteams-call] no streaming STT provider resolved (${res.code}); using file-based STT fallback`);
        }
    }
    buildStreamingDeps(session, greeting) {
        const cfg = this.api.config;
        const agentRuntime = this.api.runtime.agent;
        const transcription = this.transcription;
        return {
            ...(transcription
                ? {
                    createTranscriptionSession: (callbacks) => transcription.provider.createSession({
                        cfg,
                        providerConfig: transcription.providerConfig,
                        ...callbacks,
                    }),
                }
                : {}),
            transcribe: async (pcm16k) => {
                const wav = pcmToWav(pcm16k, MSTEAMS_PCM_SAMPLE_RATE_HZ);
                const tmp = path.join(os.tmpdir(), `msteams-call-${session.callId}-${this.sttSeq++}.wav`);
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
            provider: provider,
            providerConfig: this.realtime?.providerConfig,
            cfg: this.api.config,
            instructions: this.cfg.voice.realtime.instructions,
            greetingInstructions,
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
    onSessionEnd(info) {
        this.forgetManagedCall(info.callId);
        this.disposeCall(info.callId);
        this.lifecycle.end(info.callId, "hangup-user");
    }
    disposeCall(callId, closeReason) {
        this.pendingOutbound.delete(callId);
        this.clearOutboundTimer(callId);
        this.calls.get(callId)?.close(closeReason);
        this.calls.delete(callId);
        this.postableCalls.delete(callId);
        this.vision.release(callId);
    }
}
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
    header.writeUInt16LE(1, 20);
    header.writeUInt16LE(numChannels, 22);
    header.writeUInt32LE(sampleRate, 24);
    header.writeUInt32LE(byteRate, 28);
    header.writeUInt16LE(blockAlign, 32);
    header.writeUInt16LE(bitsPerSample, 34);
    header.write("data", 36);
    header.writeUInt32LE(pcm.length, 40);
    return Buffer.concat([header, pcm]);
}
