import { readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { setTimeout as sleep } from "node:timers/promises";
import { buildRealtimeVoiceAgentConsultWorkingResponse, consultRealtimeVoiceAgent, createRealtimeVoiceBridgeSession, REALTIME_VOICE_AGENT_CONSULT_TOOL_NAME, REALTIME_VOICE_AUDIO_FORMAT_PCM16_24KHZ, resamplePcm, resolveRealtimeVoiceAgentConsultToolsAllow, } from "openclaw/plugin-sdk/realtime-voice";
import { fetchWithSsrFGuard } from "openclaw/plugin-sdk/ssrf-runtime";
import { inferEmotion } from "./expression.js";
import { consultMediaPaths } from "./realtime-voice-compat.js";
import { isAddressed } from "./group-call-gate.js";
import { pushOrQueueBridgeImage } from "./vision-consult.js";
import { buildMinutesDocx } from "./meeting-minutes-docx.js";
import { MSTEAMS_PCM_SAMPLE_RATE_HZ, } from "./msteams-media-stream.js";
import { MSTEAMS_AGENT_TASK_TOOL, MSTEAMS_AGENT_TASK_TOOL_NAME, MSTEAMS_ASYNC_TASK_ACK, MSTEAMS_ASYNC_TASK_ACK_CALL, MSTEAMS_ASYNC_TASK_NO_TARGET, MSTEAMS_LOOK_BUDGETED, MSTEAMS_LOOK_NO_FRAME, MSTEAMS_LOOK_TOOL, MSTEAMS_LOOK_TOOL_NAME, MSTEAMS_MINUTES_TOOL, MSTEAMS_MINUTES_TOOL_NAME, MSTEAMS_REALTIME_CONSULT_SYSTEM_PROMPT, MSTEAMS_REALTIME_LOOK_SYSTEM_PROMPT, MSTEAMS_REALTIME_SHOW_SYSTEM_PROMPT, MSTEAMS_RECORDING_BLOCKED, MSTEAMS_SHOW_TOOL, MSTEAMS_SHOW_TOOL_NAME, } from "./msteams-realtime-tools.js";
import { describeMsteamsVideoFrameOwner } from "./msteams-video-frame.js";
import { resolveRealtimeFastContextConsult } from "./realtime-fast-context.js";
import { resolveVoiceResponseModel } from "./response-model.js";
import { readArgText } from "./utils.js";
import { isVerbalInterrupt } from "./verbal-interrupt.js";
const MSTEAMS_SAMPLE_RATE_HZ = MSTEAMS_PCM_SAMPLE_RATE_HZ;
const REALTIME_SAMPLE_RATE_HZ = 24_000;
const MAX_TRANSCRIPT_ENTRIES = 40;
const MAX_TRANSCRIPT_ENTRY_CHARS = 1_000;
const NOTIFY_AUDIO_QUIET_MS = 1000;
export const ECHO_SUPPRESSION_WINDOW_MS = 600;
export const ECHO_BARGE_IN_RMS = 0.04;
export function pcm16Rms(pcm) {
    const samples = Math.floor(pcm.length / 2);
    if (samples === 0) {
        return 0;
    }
    let sum = 0;
    for (let i = 0; i < samples; i += 1) {
        const sample = pcm.readInt16LE(i * 2) / 32768;
        sum += sample * sample;
    }
    return Math.sqrt(sum / samples);
}
export function parseMinutesSections(text) {
    const sections = [];
    let current;
    for (const rawLine of text.split(/\r?\n/)) {
        const line = rawLine.trim();
        if (!line) {
            continue;
        }
        const headingMatch = /^#{1,6}\s+(.*\S)\s*$/.exec(line) ?? /^\*\*(.+?)\*\*:?\s*$/.exec(line);
        if (headingMatch?.[1]) {
            current = { heading: headingMatch[1].replace(/:$/, "").trim(), items: [] };
            sections.push(current);
            continue;
        }
        const bulletMatch = /^(?:[-*•]|\d+[.)])\s+(.*\S)\s*$/.exec(line);
        const item = bulletMatch?.[1] ?? line;
        if (!current) {
            current = { heading: "Summary", items: [] };
            sections.push(current);
        }
        current.items.push(item);
    }
    return sections;
}
export function shouldSuppressEcho(pcm16k, playbackActiveUntil, opts) {
    if (opts?.suppressInputDuringPlayback === false) {
        return false;
    }
    const inPlaybackWindow = Date.now() <
        playbackActiveUntil + (opts?.echoSuppressionWindowMs ?? ECHO_SUPPRESSION_WINDOW_MS);
    if (!inPlaybackWindow) {
        return false;
    }
    if (opts?.allowBargeIn === false) {
        return true;
    }
    return pcm16Rms(pcm16k) < (opts?.echoBargeInRms ?? ECHO_BARGE_IN_RMS);
}
export function toTileCaption(text) {
    const trimmed = text?.replace(/\s+/g, " ").trim();
    if (!trimmed) {
        return undefined;
    }
    return trimmed.length > 140 ? `${trimmed.slice(0, 139)}…` : trimmed;
}
const REALTIME_VISION_PUSH_INTERVAL_MS = 6000;
const MSTEAMS_MAX_DISPLAY_IMAGE_BYTES = 4_000_000;
const MSTEAMS_LOOK_HISTORY_FRAMES = 6;
const DISPLAY_SLIDESHOW_MS = 4_000;
const DISPLAY_SLIDESHOW_OVERLAP_MS = 500;
const MSTEAMS_SHOW_TIMEOUT_MS = 90_000;
function mimeForImageExtension(pathOrUrl) {
    const ext = (pathOrUrl.split(/[?#]/)[0] ?? "").toLowerCase().split(".").pop() ?? "";
    switch (ext) {
        case "png":
            return "image/png";
        case "jpg":
        case "jpeg":
            return "image/jpeg";
        case "gif":
            return "image/gif";
        case "webp":
            return "image/webp";
        default:
            return null;
    }
}
function withRosterInstruction(instructions, callerName) {
    const name = callerName?.trim();
    if (!name) {
        return instructions;
    }
    const clause = [
        `CALLER IDENTITY: You are speaking with ${name}. Greet them by their first name once, warmly`,
        "and briefly, then continue naturally — do not repeat their name every turn. In a group call,",
        'each caller turn is prefixed with the speaker\'s name (e.g. "Sara: ..."); use those names to',
        "address people directly when it helps, but never read the prefix aloud as part of your reply.",
    ].join(" ");
    return instructions ? `${instructions}\n\n${clause}` : clause;
}
function withBilingualInstruction(instructions, bilingual) {
    if (!bilingual) {
        return instructions;
    }
    const clause = [
        "BILINGUAL (Arabic / English): Detect the language the caller is speaking — Arabic or English —",
        "and always reply in that same language, matching dialect and register. If the caller switches",
        "language mid-call, switch with them. When asked to translate, translate accurately between",
        "Arabic and English and read out only the translation.",
    ].join(" ");
    return instructions ? `${instructions}\n\n${clause}` : clause;
}
function withGroupGateInstruction(instructions, gate) {
    const phrases = gate?.wakePhrases?.filter((p) => p.trim().length > 0) ?? [];
    if (!gate?.requireAddress || phrases.length === 0) {
        return instructions;
    }
    const names = phrases.map((p) => `"${p}"`).join(", ");
    const clause = [
        "GROUP-CALL ETIQUETTE: If more than one person is on this call (a group meeting), do NOT reply",
        `unless someone addresses you by name (${names}) or clearly directs a question to you. When you`,
        "are not addressed, stay silent and just listen — do not narrate, acknowledge, or interject.",
        "Once addressed, you may continue a short back-and-forth until the topic moves on. In a",
        "one-on-one call (only you and one person), respond normally to everything.",
    ].join(" ");
    return instructions ? `${instructions}\n\n${clause}` : clause;
}
export function createMsteamsRealtimeCall(params) {
    const { session, deps } = params;
    const { logger } = deps;
    const callId = session.callId;
    const sessionScopeId = deps.voiceConfig?.sessionScope === "per-call"
        ? callId
        : deps.voiceConfig?.sessionScope === "per-thread"
            ? session.threadId?.trim() || session.caller.aadId || callId
            : session.caller.aadId || callId;
    let outboundSeq = 0;
    let outboundTimestampMs = 0;
    let turnId = 0;
    let closed = false;
    let deliveryComplete = false;
    const callStartedAt = Date.now();
    let playbackEndAt = 0;
    let recordingActive = session.recordingStatus === "active";
    const requireRecordingStatus = deps.requireRecordingStatus !== false;
    const recordingGateBlocks = () => requireRecordingStatus && !recordingActive;
    const transcript = [];
    let lastLookData;
    let lastLookText;
    const lastPushedFrameData = {};
    let visionPushTimer;
    const pendingAmbientImages = [];
    let lastSentExpression;
    let thinking = false;
    let humanCount = 1;
    let lastAddressedAt;
    const groupGateActive = !!deps.groupCallGate &&
        deps.groupCallGate.requireAddress &&
        deps.groupCallGate.wakePhrases.some((p) => p.trim().length > 0);
    let currentSpeakerName;
    let greetingTriggered = false;
    let callerTurnStarted = false;
    let lastUserEntrySpeaker;
    function recordTranscript(role, text, speaker) {
        if (recordingGateBlocks()) {
            return;
        }
        const trimmed = text.trim();
        if (!trimmed) {
            return;
        }
        const last = transcript.at(-1);
        const sameSpeaker = role !== "user" || speaker === lastUserEntrySpeaker;
        if (last &&
            last.role === role &&
            sameSpeaker &&
            last.text.length + trimmed.length < MAX_TRANSCRIPT_ENTRY_CHARS) {
            last.text = `${last.text} ${trimmed}`.trim();
        }
        else {
            transcript.push({ role, text: trimmed });
        }
        if (role === "user") {
            lastUserEntrySpeaker = speaker;
        }
        if (transcript.length > MAX_TRANSCRIPT_ENTRIES) {
            transcript.splice(0, transcript.length - MAX_TRANSCRIPT_ENTRIES);
        }
    }
    const consultToolPolicy = deps.toolPolicy ?? deps.voiceConfig?.realtime.toolPolicy ?? "none";
    const consultAgentId = deps.voiceConfig?.agentId ?? "main";
    const consultSessionKey = `agent:${consultAgentId}:subagent:msteams:${sessionScopeId}`;
    const asyncTasksEnabled = Boolean(deps.agentRuntime && deps.voiceConfig && deps.cfg) && consultToolPolicy === "owner";
    const visionEnabled = Boolean(deps.agentRuntime && deps.voiceConfig && deps.cfg && deps.getLatestFrame) &&
        consultToolPolicy !== "none";
    const showEnabled = Boolean(deps.agentRuntime && deps.voiceConfig && deps.cfg) && consultToolPolicy !== "none";
    const bridgeTools = [
        ...(deps.tools ?? []),
        ...(asyncTasksEnabled ? [MSTEAMS_AGENT_TASK_TOOL] : []),
        ...(visionEnabled ? [MSTEAMS_LOOK_TOOL] : []),
        ...(showEnabled ? [MSTEAMS_SHOW_TOOL] : []),
        ...(asyncTasksEnabled && session.caller.aadId ? [MSTEAMS_MINUTES_TOOL] : []),
    ];
    const runMsteamsConsult = (opts) => {
        const { provider: agentProvider, model } = resolveVoiceResponseModel({
            voiceConfig: opts.voiceConfig,
            agentRuntime: opts.agentRuntime,
        });
        const thinkLevel = opts.voiceConfig.realtime.consultThinkingLevel ??
            opts.agentRuntime.resolveThinkingDefault({ cfg: opts.cfg, provider: agentProvider, model });
        const ambientImages = pendingAmbientImages.splice(0, pendingAmbientImages.length);
        const mergedImages = [...(opts.images ?? []), ...ambientImages];
        return consultRealtimeVoiceAgent({
            cfg: opts.cfg,
            agentRuntime: opts.agentRuntime,
            logger: { warn: (message) => logger?.warn(message) },
            agentId: opts.agentId,
            sessionKey: opts.sessionKey,
            messageProvider: "voice",
            lane: "voice",
            runIdPrefix: opts.runIdPrefix,
            args: opts.args,
            ...(mergedImages.length ? { images: mergedImages } : {}),
            transcript: [...transcript],
            surface: opts.surface,
            userLabel: "Caller",
            assistantLabel: "Agent",
            questionSourceLabel: "caller",
            provider: agentProvider,
            model,
            thinkLevel,
            fastMode: opts.fastMode ?? opts.voiceConfig.realtime.consultFastMode,
            ...(opts.timeoutMs !== undefined ? { timeoutMs: opts.timeoutMs } : {}),
            ...(opts.trustLocalMedia ? { trustLocalMedia: true } : {}),
            ...(opts.deliveryContext ? { deliveryContext: opts.deliveryContext } : {}),
            toolsAllow: resolveRealtimeVoiceAgentConsultToolsAllow(opts.toolPolicy),
            extraSystemPrompt: opts.extraSystemPrompt,
        });
    };
    const realtime = createRealtimeVoiceBridgeSession({
        provider: deps.provider,
        cfg: deps.cfg,
        providerConfig: deps.providerConfig,
        audioFormat: REALTIME_VOICE_AUDIO_FORMAT_PCM16_24KHZ,
        instructions: withGroupGateInstruction(withBilingualInstruction(withRosterInstruction(deps.instructions, session.caller.displayName ?? undefined), deps.voiceConfig?.msteams?.bilingual), deps.groupCallGate),
        initialGreetingInstructions: deps.greetingInstructions,
        triggerGreetingOnReady: Boolean(deps.greetingInstructions) && !deps.greetingOnRecordingActive,
        autoRespondToAudio: true,
        interruptResponseOnInputAudio: true,
        tools: bridgeTools,
        audioSink: {
            isOpen: () => !closed,
            sendAudio: (pcm24k) => {
                if (closed || pcm24k.length === 0) {
                    return;
                }
                if (humanCount >= 2 && groupGateActive) {
                    const g = deps.groupCallGate;
                    const now = Date.now();
                    if (lastAddressedAt === undefined || now - lastAddressedAt > g.followUpWindowMs) {
                        return;
                    }
                    lastAddressedAt = now;
                }
                const pcm16k = resamplePcm(pcm24k, REALTIME_SAMPLE_RATE_HZ, MSTEAMS_SAMPLE_RATE_HZ);
                playbackEndAt = Math.max(playbackEndAt, Date.now()) + pcm16k.length / 32;
                session.send({
                    type: "audio.frame",
                    seq: outboundSeq,
                    timestampMs: outboundTimestampMs,
                    payloadBase64: pcm16k.toString("base64"),
                });
                outboundSeq += 1;
                outboundTimestampMs += Math.round((pcm16k.length / 2 / MSTEAMS_SAMPLE_RATE_HZ) * 1000);
            },
            clearAudio: () => {
                turnId += 1;
                session.send({ type: "assistant.cancel", turnId });
                playbackEndAt = Date.now();
            },
        },
        onTranscript: (role, text, isFinal) => {
            if (role === "user" && text.trim().length > 0) {
                callerTurnStarted = true;
            }
            if (role === "user" &&
                groupGateActive &&
                text.trim().length > 0 &&
                isAddressed(text, deps.groupCallGate.wakePhrases)) {
                lastAddressedAt = Date.now();
            }
            if (role === "assistant" && !closed && !thinking) {
                const emotion = inferEmotion(text);
                if (emotion !== lastSentExpression) {
                    lastSentExpression = emotion;
                    logger?.debug?.(`MsteamsRealtime: expression cue '${emotion}' for ${callId}`);
                    try {
                        session.send({ type: "expression", emotion });
                    }
                    catch {
                    }
                }
            }
            if (isFinal) {
                recordTranscript(role, role === "user" && currentSpeakerName ? `${currentSpeakerName}: ${text}` : text, currentSpeakerName);
                if (role === "user" &&
                    Date.now() < playbackEndAt &&
                    isVerbalInterrupt(text, deps.groupCallGate?.wakePhrases)) {
                    logger?.debug?.(`MsteamsRealtime: verbal interrupt on ${callId} — flushing playback`);
                    turnId += 1;
                    try {
                        session.send({ type: "assistant.cancel", turnId });
                    }
                    catch {
                    }
                    playbackEndAt = Date.now();
                }
                if (role === "assistant" && !deliveryComplete && deps.onDeliveryComplete) {
                    deliveryComplete = true;
                    watchAudioDrainThenSignal();
                }
            }
        },
        onToolCall: (event, rtSession) => {
            if (event.name === MSTEAMS_AGENT_TASK_TOOL_NAME) {
                handleAsyncTask(event, rtSession);
                return;
            }
            const handler = event.name === MSTEAMS_MINUTES_TOOL_NAME
                ? handleMinutes
                : event.name === MSTEAMS_LOOK_TOOL_NAME
                    ? handleLook
                    : event.name === MSTEAMS_SHOW_TOOL_NAME
                        ? handleShow
                        : event.name === REALTIME_VOICE_AGENT_CONSULT_TOOL_NAME
                            ? handleConsult
                            : undefined;
            if (!handler) {
                logger?.warn(`MsteamsRealtime: no handler for tool '${event.name}' on ${callId}`);
                rtSession.submitToolResult(event.callId, {
                    text: `The tool "${event.name}" is not available on this call.`,
                });
                return;
            }
            setThinking(true);
            void handler(event, rtSession).finally(() => setThinking(false));
        },
        onError: (error) => {
            logger?.warn(`MsteamsRealtime: bridge error — ${error.message}`);
        },
        onClose: () => {
            closeCall("realtime-closed");
        },
    });
    function pushLatestFrameToModel() {
        if (closed || recordingGateBlocks() || !deps.getLatestFrame) {
            return;
        }
        for (const source of ["screenshare", "camera"]) {
            const frame = deps.getLatestFrame(source);
            if (!frame || frame.dataBase64 === lastPushedFrameData[source]) {
                continue;
            }
            if (deps.visionBudget && !deps.visionBudget.tryConsume(callId, Date.now())) {
                break;
            }
            const label = source === "camera" ? "camera" : "screen-share";
            logger?.debug?.(`MsteamsRealtime: ambient vision push (${label}) for ${callId}`);
            try {
                const owner = describeMsteamsVideoFrameOwner(frame);
                pushOrQueueBridgeImage(realtime, {
                    dataBase64: frame.dataBase64,
                    mime: frame.mime,
                    text: owner ? `Live ${label} — ${owner}.` : `Live ${label} of the call.`,
                }, pendingAmbientImages);
                lastPushedFrameData[source] = frame.dataBase64;
            }
            catch (err) {
                deps.visionBudget?.refund(callId);
                logger?.debug?.(`MsteamsRealtime: vision push failed for ${callId} — ${err instanceof Error ? err.message : String(err)}`);
            }
        }
    }
    function watchAudioDrainThenSignal() {
        const check = () => {
            if (closed) {
                return;
            }
            const dueAt = playbackEndAt + NOTIFY_AUDIO_QUIET_MS;
            if (Date.now() >= dueAt) {
                deps.onDeliveryComplete?.();
                return;
            }
            const t = setTimeout(check, Math.max(50, dueAt - Date.now()));
            t.unref?.();
        };
        const t = setTimeout(check, NOTIFY_AUDIO_QUIET_MS);
        t.unref?.();
    }
    function setThinking(on) {
        if (on === thinking || closed) {
            return;
        }
        thinking = on;
        try {
            if (on) {
                session.send({ type: "expression", emotion: "thinking" });
                lastSentExpression = "thinking";
            }
            else if (lastSentExpression === "thinking") {
                session.send({ type: "expression", emotion: "neutral" });
                lastSentExpression = "neutral";
            }
        }
        catch {
        }
    }
    if (deps.getLatestFrame) {
        visionPushTimer = setInterval(pushLatestFrameToModel, REALTIME_VISION_PUSH_INTERVAL_MS);
        visionPushTimer.unref?.();
    }
    function withConsultGuards(opts) {
        return async (event, rtSession) => {
            if (recordingGateBlocks()) {
                logger?.debug?.(`MsteamsRealtime: ${opts.label} refused for ${callId} — recording not active`);
                rtSession.submitToolResult(event.callId, MSTEAMS_RECORDING_BLOCKED);
                return;
            }
            const { agentRuntime, voiceConfig, cfg } = deps;
            if (!agentRuntime ||
                !voiceConfig ||
                !cfg ||
                (opts.requireFrameSource && !deps.getLatestFrame)) {
                rtSession.submitToolResult(event.callId, { text: opts.unavailableText });
                return;
            }
            try {
                await opts.handler({
                    event,
                    rtSession,
                    consult: {
                        agentRuntime,
                        voiceConfig,
                        cfg,
                        agentId: consultAgentId,
                        sessionKey: consultSessionKey,
                        toolPolicy: deps.toolPolicy ?? voiceConfig.realtime.toolPolicy,
                    },
                    sendWorkingFiller: () => {
                        if (rtSession.bridge?.supportsToolResultContinuation) {
                            rtSession.submitToolResult(event.callId, buildRealtimeVoiceAgentConsultWorkingResponse("caller"), { willContinue: true });
                        }
                    },
                });
            }
            catch (err) {
                logger?.warn(`MsteamsRealtime: ${opts.label} failed for ${callId} — ${err instanceof Error ? err.message : String(err)}`);
                rtSession.submitToolResult(event.callId, { text: opts.errorText });
            }
        };
    }
    const handleConsult = withConsultGuards({
        label: "consult",
        unavailableText: "The assistant agent is not available right now.",
        errorText: "Sorry, I ran into a problem while working on that.",
        handler: async ({ event, rtSession, consult, sendWorkingFiller }) => {
            const fastContext = await resolveRealtimeFastContextConsult({
                cfg: consult.cfg,
                agentId: consult.agentId,
                sessionKey: consult.sessionKey,
                config: consult.voiceConfig.realtime.fastContext,
                args: event.args,
                logger: { debug: (message) => logger?.debug?.(message) },
            });
            if (fastContext.handled) {
                rtSession.submitToolResult(event.callId, fastContext.result);
                return;
            }
            sendWorkingFiller();
            const result = await runMsteamsConsult({
                ...consult,
                runIdPrefix: `voice-realtime-consult:${callId}`,
                args: event.args,
                surface: "a live Microsoft Teams call",
                extraSystemPrompt: MSTEAMS_REALTIME_CONSULT_SYSTEM_PROMPT,
                timeoutMs: consult.voiceConfig.responseTimeoutMs,
            });
            rtSession.submitToolResult(event.callId, result);
        },
    });
    const handleMinutes = withConsultGuards({
        label: "minutes",
        unavailableText: "I can't post minutes from this call right now.",
        errorText: "Sorry, I had trouble posting the minutes.",
        handler: async ({ event, rtSession }) => {
            rtSession.submitToolResult(event.callId, {
                text: "Minutes are being written and posted to the Teams chat now.",
            });
            await runMeetingRecap();
        },
    });
    const handleLook = withConsultGuards({
        label: "look",
        unavailableText: "The assistant can't look at video right now.",
        errorText: "Sorry, I had trouble seeing that.",
        requireFrameSource: true,
        handler: async ({ event, rtSession, consult, sendWorkingFiller }) => {
            const sourceArg = readArgText(event.args, "source");
            const source = sourceArg === "camera" || sourceArg === "screenshare" ? sourceArg : undefined;
            const historyScope = readArgText(event.args, "scope") === "history";
            const historyFrames = historyScope
                ? (deps.getFrameHistory?.(MSTEAMS_LOOK_HISTORY_FRAMES) ?? [])
                : [];
            const frame = historyScope ? historyFrames.at(-1) : deps.getLatestFrame?.(source);
            if (!frame) {
                rtSession.submitToolResult(event.callId, MSTEAMS_LOOK_NO_FRAME);
                return;
            }
            if (!historyScope && lastLookData === frame.dataBase64 && lastLookText) {
                logger?.debug?.(`MsteamsRealtime: look cache hit for ${callId} (unchanged frame)`);
                rtSession.submitToolResult(event.callId, { text: lastLookText });
                return;
            }
            if (deps.visionBudget && !deps.visionBudget.tryConsume(callId, Date.now())) {
                logger?.debug?.(`MsteamsRealtime: look over vision budget for ${callId}`);
                rtSession.submitToolResult(event.callId, MSTEAMS_LOOK_BUDGETED);
                return;
            }
            sendWorkingFiller();
            const lookFrames = historyScope ? historyFrames : [frame];
            const result = await runMsteamsConsult({
                ...consult,
                runIdPrefix: `voice-realtime-look:${callId}`,
                args: event.args,
                images: lookFrames.map((f) => ({
                    type: "image",
                    data: f.dataBase64,
                    mimeType: f.mime,
                })),
                surface: historyScope
                    ? `a live Microsoft Teams call — the attached images are scene-change keyframes from earlier in the call, oldest first: ${lookFrames
                        .map((f, i) => {
                        const owner = describeMsteamsVideoFrameOwner(f);
                        const age = Math.max(0, Math.round((Date.now() - f.ts) / 1000));
                        return `image ${i + 1} (~${age}s ago${owner ? `, ${owner}` : ""})`;
                    })
                        .join("; ")}`
                    : (() => {
                        const owner = describeMsteamsVideoFrameOwner(frame);
                        return owner
                            ? `a live Microsoft Teams call — the attached image is ${owner}`
                            : "a live Microsoft Teams call (a participant is sharing video)";
                    })(),
                extraSystemPrompt: MSTEAMS_REALTIME_LOOK_SYSTEM_PROMPT,
                timeoutMs: consult.voiceConfig.responseTimeoutMs,
            });
            if (!historyScope) {
                lastLookData = frame.dataBase64;
                lastLookText = result.text;
            }
            rtSession.submitToolResult(event.callId, result);
        },
    });
    async function loadDisplayImage(pathOrUrl) {
        if (/^https?:\/\//i.test(pathOrUrl)) {
            try {
                const { response, release } = await fetchWithSsrFGuard({
                    url: pathOrUrl,
                    init: { method: "GET" },
                    policy: {},
                    timeoutMs: 15_000,
                });
                try {
                    if (!response.ok) {
                        return null;
                    }
                    const contentType = (response.headers.get("content-type") ?? "").split(";")[0]?.trim() ?? "";
                    const mime = contentType.startsWith("image/")
                        ? contentType
                        : mimeForImageExtension(pathOrUrl);
                    if (!mime) {
                        return null;
                    }
                    const bytes = Buffer.from(await response.arrayBuffer());
                    if (bytes.length === 0 || bytes.length > MSTEAMS_MAX_DISPLAY_IMAGE_BYTES) {
                        return null;
                    }
                    return { bytes, mime };
                }
                finally {
                    void release?.();
                }
            }
            catch {
                return null;
            }
        }
        const mime = mimeForImageExtension(pathOrUrl);
        if (!mime) {
            return null;
        }
        try {
            const bytes = await readFile(pathOrUrl);
            if (bytes.length === 0 || bytes.length > MSTEAMS_MAX_DISPLAY_IMAGE_BYTES) {
                return null;
            }
            return { bytes, mime };
        }
        catch {
            return null;
        }
    }
    async function forwardDisplayImages(mediaPaths, caption) {
        const images = [];
        for (const pathOrUrl of mediaPaths) {
            const img = await loadDisplayImage(pathOrUrl);
            if (img) {
                images.push(img);
            }
            else {
                logger?.debug?.(`MsteamsRealtime: skipped non-displayable media ${pathOrUrl} for ${callId}`);
            }
        }
        const [first, ...rest] = images;
        if (!first) {
            return 0;
        }
        const sequence = rest.length > 0;
        const sendOne = (img, isLast) => {
            try {
                logger?.debug?.(`MsteamsRealtime: display.image (${img.mime}, ${img.bytes.length}B${caption ? ", captioned" : ""}) for ${callId}`);
                session.send({
                    type: "display.image",
                    dataBase64: img.bytes.toString("base64"),
                    mime: img.mime,
                    mode: "overlay",
                    ...(sequence && !isLast
                        ? { durationMs: DISPLAY_SLIDESHOW_MS + DISPLAY_SLIDESHOW_OVERLAP_MS }
                        : {}),
                    ...(caption ? { caption } : {}),
                });
            }
            catch (err) {
                logger?.debug?.(`MsteamsRealtime: display.image send failed for ${callId} — ${err instanceof Error ? err.message : String(err)}`);
            }
        };
        sendOne(first, !sequence);
        if (sequence) {
            void (async () => {
                for (const [idx, img] of rest.entries()) {
                    await sleep(DISPLAY_SLIDESHOW_MS);
                    if (closed) {
                        return;
                    }
                    sendOne(img, idx === rest.length - 1);
                }
            })();
        }
        return images.length;
    }
    const handleShow = withConsultGuards({
        label: "show",
        unavailableText: "I can't show images on this call.",
        errorText: "Sorry, I had trouble showing that.",
        handler: async ({ event, rtSession, consult, sendWorkingFiller }) => {
            sendWorkingFiller();
            const showRequest = event.args &&
                typeof event.args === "object" &&
                typeof event.args.request === "string"
                ? event.args.request
                : undefined;
            const result = await runMsteamsConsult({
                ...consult,
                runIdPrefix: `voice-realtime-show:${callId}`,
                args: showRequest ? { question: showRequest } : event.args,
                surface: "a live Microsoft Teams video call — show the caller an image on the bot's video tile",
                extraSystemPrompt: MSTEAMS_REALTIME_SHOW_SYSTEM_PROMPT,
                toolPolicy: "owner",
                timeoutMs: Math.max(consult.voiceConfig.responseTimeoutMs, MSTEAMS_SHOW_TIMEOUT_MS),
                trustLocalMedia: true,
            });
            const resultMediaPaths = consultMediaPaths(result);
            const shown = await forwardDisplayImages(resultMediaPaths, toTileCaption(result.text));
            rtSession.submitToolResult(event.callId, {
                text: shown > 0
                    ? result.text || "I'm showing it on your screen now."
                    : result.text || "Sorry, I couldn't produce an image to show.",
            });
        },
    });
    function handleAsyncTask(event, rtSession) {
        if (recordingGateBlocks()) {
            logger?.debug?.(`MsteamsRealtime: task refused for ${callId} — recording not active`);
            rtSession.submitToolResult(event.callId, MSTEAMS_RECORDING_BLOCKED);
            return;
        }
        if (!session.caller.aadId) {
            logger?.warn(`MsteamsRealtime: task refused for ${callId} — no caller.aadId delivery target`);
            rtSession.submitToolResult(event.callId, MSTEAMS_ASYNC_TASK_NO_TARGET);
            return;
        }
        const task = readArgText(event.args, "task") ?? readArgText(event.args, "question");
        const deliverVia = readArgText(event.args, "deliverVia") === "call" ? "call" : "message";
        if (!task) {
            logger?.warn(`MsteamsRealtime: task refused for ${callId} — no task text in tool args`);
            rtSession.submitToolResult(event.callId, {
                text: "I didn't catch what the task is — please tell me again what you'd like me to do.",
            });
            return;
        }
        rtSession.submitToolResult(event.callId, deliverVia === "call" ? MSTEAMS_ASYNC_TASK_ACK_CALL : MSTEAMS_ASYNC_TASK_ACK);
        void runAsyncTask(task, deliverVia);
    }
    async function runAsyncTask(task, deliverVia) {
        const { agentRuntime, voiceConfig, cfg } = deps;
        if (!agentRuntime || !voiceConfig || !cfg) {
            return;
        }
        const aadId = session.caller.aadId ?? undefined;
        const deliveryTarget = aadId ? `user:${aadId}` : undefined;
        const deliveryInstruction = !deliveryTarget
            ? "This task was delegated from a Microsoft Teams voice call and runs in the background; deliver the final result to the caller when complete."
            : deliverVia === "call"
                ? `This task was delegated from a live Microsoft Teams voice call and now runs in the background; the caller is no longer on the line. FIRST actually complete the task and determine the final answer. THEN deliver it by invoking the voice_call tool exactly once: action "initiate_call", to "${deliveryTarget}", mode "notify". CRITICAL: the caller hears ONLY your "message" and has NO memory of what they asked — so "message" must be a COMPLETE, STANDALONE spoken result that both restates the topic AND gives the answer in one breath. Good: "Here's the Dubai time you asked for — it's 5:41 PM." Bad (do NOT do this): a greeting, a question, "I'm calling about your request", "let me check", an empty/placeholder value, or a bare answer with no context. If you genuinely could not determine the answer, set "message" to a clear one-sentence explanation of what went wrong instead. Place the call exactly once.`
                : `This task was delegated from a live Microsoft Teams voice call and now runs in the background; the caller is no longer waiting on the line. Complete the task, then deliver the final result to the caller by calling the message tool exactly once with action "send", channel "msteams", target "${deliveryTarget}". Keep the delivered message concise.`;
        try {
            await runMsteamsConsult({
                agentRuntime,
                voiceConfig,
                cfg,
                agentId: consultAgentId,
                sessionKey: consultSessionKey,
                runIdPrefix: `voice-realtime-task:${callId}`,
                args: { question: task },
                surface: "a Microsoft Teams voice call (background task)",
                extraSystemPrompt: `${MSTEAMS_REALTIME_CONSULT_SYSTEM_PROMPT} ${deliveryInstruction}`,
                toolPolicy: consultToolPolicy,
                fastMode: false,
            });
            logger?.debug?.(`MsteamsRealtime: background task complete for ${callId}`);
        }
        catch (err) {
            logger?.warn(`MsteamsRealtime: background task failed for ${callId} — ${err instanceof Error ? err.message : String(err)}`);
        }
    }
    async function runMeetingRecap() {
        const { agentRuntime, voiceConfig, cfg } = deps;
        if (!agentRuntime || !voiceConfig || !cfg) {
            return;
        }
        const aadId = session.caller.aadId;
        const durationMin = Math.max(1, Math.round((Date.now() - callStartedAt) / 60_000));
        const callerName = session.caller.displayName ?? "the caller";
        const lines = transcript
            .map((t) => `${t.role === "assistant" ? "Assistant" : "Caller side"}: ${t.text}`)
            .join("\n");
        const isGroupRecap = humanCount >= 2 && Boolean(session.threadId?.trim());
        const recapTarget = isGroupRecap ? `conversation:${session.threadId.trim()}` : `user:${aadId}`;
        const deliveryContext = { channel: "msteams", to: recapTarget };
        try {
            logger?.info(`MsteamsRealtime: posting meeting recap for ${callId}`);
            const summary = await runMsteamsConsult({
                agentRuntime,
                voiceConfig,
                cfg,
                agentId: consultAgentId,
                sessionKey: consultSessionKey,
                runIdPrefix: `voice-realtime-recap:${callId}`,
                args: {
                    question: `Write concise meeting minutes from this Microsoft Teams call transcript.\n` +
                        `Call: with ${callerName}, ~${durationMin} min, ${humanCount} human participant(s).\n` +
                        `Transcript (most recent ${transcript.length} turns; "Caller side" may include multiple ` +
                        `people. Some caller turns begin with "Name:" — that is real speaker attribution from ` +
                        `unmixed audio; attribute statements and action items ONLY via those prefixes, and ` +
                        `never guess attribution for unprefixed turns):\n${lines}`,
                },
                surface: "a Microsoft Teams call that just ended (meeting recap)",
                extraSystemPrompt: `${MSTEAMS_REALTIME_CONSULT_SYSTEM_PROMPT} Return ONLY the minutes as markdown with these ` +
                    `headed sections, omitting empty ones: "### Key points", "### Decisions", "### Action ` +
                    `items"; under each, one "- " bullet per item. Keep it brief and factual; no invented ` +
                    `attribution. Do NOT write any files, do NOT call the message tool, and do NOT add any ` +
                    `preamble or closing — your reply IS the minutes body.`,
                toolPolicy: consultToolPolicy,
                deliveryContext,
                fastMode: false,
            });
            const summaryText = summary.text?.trim() ?? "";
            const subtitle = `Call with ${callerName} — ~${durationMin} min, ${humanCount} human participant(s).`;
            let docxPath;
            try {
                const buffer = await buildMinutesDocx({
                    title: "Meeting minutes",
                    subtitle,
                    sections: parseMinutesSections(summaryText),
                    transcript: transcript.map((t) => ({ role: t.role, text: t.text })),
                });
                let outDir = tmpdir();
                try {
                    const workspaceDir = agentRuntime.resolveAgentWorkspaceDir(cfg, consultAgentId);
                    if (workspaceDir) {
                        await agentRuntime.ensureAgentWorkspace({ dir: workspaceDir });
                        outDir = workspaceDir;
                    }
                }
                catch {
                }
                docxPath = join(outDir, `meeting-minutes-${callId}.docx`);
                await writeFile(docxPath, buffer);
            }
            catch (err) {
                logger?.warn(`MsteamsRealtime: minutes docx build failed for ${callId}, sending text-only — ${err instanceof Error ? err.message : String(err)}`);
                docxPath = undefined;
            }
            const bodyForSend = summaryText || "Meeting minutes are attached.";
            const mediaInstruction = docxPath
                ? `Attach the local file at this absolute path as the message tool's media parameter on the ` +
                    `SAME send: ${docxPath}. If the attachment fails, send the text-only message. `
                : "";
            await runMsteamsConsult({
                agentRuntime,
                voiceConfig,
                cfg,
                agentId: consultAgentId,
                sessionKey: consultSessionKey,
                runIdPrefix: `voice-realtime-recap-send:${callId}`,
                args: {
                    question: `Deliver these meeting minutes verbatim. Do not rewrite, summarize, or add to them.\n\n` +
                        bodyForSend,
                },
                surface: "delivering meeting minutes from a Microsoft Teams call that just ended",
                extraSystemPrompt: `${MSTEAMS_REALTIME_CONSULT_SYSTEM_PROMPT} Call the message tool exactly once with action ` +
                    `"send", channel "msteams", target "${recapTarget}", and the provided minutes as the text. ` +
                    mediaInstruction +
                    `Do NOT send it to any other conversation; if that exact target cannot be reached, do not ` +
                    `send. Do NOT author or edit the minutes content — send it as given.`,
                toolPolicy: consultToolPolicy,
                deliveryContext,
                trustLocalMedia: Boolean(docxPath),
                fastMode: false,
            });
        }
        catch (err) {
            logger?.warn(`MsteamsRealtime: meeting recap failed for ${callId} — ${err instanceof Error ? err.message : String(err)}`);
        }
    }
    function closeCall(reason) {
        if (closed) {
            return;
        }
        closed = true;
        if (deps.voiceConfig?.msteams?.meetingRecap === true &&
            !deps.onDeliveryComplete &&
            recordingActive &&
            session.caller.aadId &&
            transcript.length >= 4) {
            void runMeetingRecap();
        }
        if (visionPushTimer) {
            clearInterval(visionPushTimer);
            visionPushTimer = undefined;
        }
        try {
            realtime.close();
        }
        catch {
        }
        if (reason !== undefined) {
            try {
                session.close(reason);
            }
            catch {
            }
        }
    }
    void realtime.connect().catch((err) => {
        logger?.error(`MsteamsRealtime: connect failed — ${err instanceof Error ? err.message : String(err)}`);
        closeCall("realtime-unavailable");
    });
    return {
        pushAudio: (pcm16k) => {
            if (closed || pcm16k.length === 0) {
                return;
            }
            if (recordingGateBlocks()) {
                return;
            }
            if (shouldSuppressEcho(pcm16k, playbackEndAt, { ...deps, allowBargeIn: callerTurnStarted })) {
                return;
            }
            const pcm24k = resamplePcm(pcm16k, MSTEAMS_SAMPLE_RATE_HZ, REALTIME_SAMPLE_RATE_HZ);
            realtime.sendAudio(pcm24k);
        },
        notifyInboundFrame: () => {
            pushLatestFrameToModel();
        },
        setHumanCount: (count) => {
            humanCount = count;
        },
        notifyDtmf: (digit) => {
            if (recordingGateBlocks()) {
                return;
            }
            const text = `[The caller pressed the "${digit}" key on their phone keypad.]`;
            recordTranscript("user", text);
            realtime.sendUserMessage(text);
        },
        setCurrentSpeaker: (name) => {
            currentSpeakerName = name;
        },
        setRecordingActive: (active) => {
            recordingActive = active;
            if (active &&
                deps.greetingOnRecordingActive &&
                deps.greetingInstructions &&
                !greetingTriggered) {
                greetingTriggered = true;
                try {
                    realtime.triggerGreeting(deps.greetingInstructions);
                }
                catch (err) {
                    logger?.warn(`MsteamsRealtime: deferred greeting failed for ${callId} — ${err instanceof Error ? err.message : String(err)}`);
                }
            }
        },
        say: (text) => {
            if (closed || !text || !text.trim())
                return;
            try {
                realtime.triggerGreeting(text);
            }
            catch (err) {
                logger?.warn(`MsteamsRealtime: assistant.say failed for ${callId} — ${err instanceof Error ? err.message : String(err)}`);
            }
        },
        close: (reason) => {
            closeCall(reason);
        },
    };
}
