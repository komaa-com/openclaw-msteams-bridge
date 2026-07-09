/**
 * Streaming STT→agent→TTS voice path for non-realtime providers.
 *
 * Parallel to {@link createMsteamsRealtimeCall}: instead of a realtime speech-to-speech model, this
 * drives a deterministic turn loop — inbound caller PCM is segmented into utterances by a simple
 * energy VAD, transcribed (STT), answered by the OpenClaw agent, synthesized (TTS) and streamed back
 * as paced 20 ms frames. Barge-in cancels in-flight playback; a half-duplex echo guard ignores our
 * own audio echoing off the caller's device while we speak.
 *
 * STT, agent, and TTS are INJECTED (`transcribe` / `consult` / `ttsProvider`) so this module is
 * provider-agnostic and unit-testable with fakes. The runtime wires the concrete implementations
 * (STT via `api.runtime.mediaUnderstanding.transcribeAudioFile`, agent via `consultRealtimeVoiceAgent`,
 * TTS via `api.runtime.tts`).
 */
import { MSTEAMS_PCM_SAMPLE_RATE_HZ, } from "./msteams-media-stream.js";
import { playTtsToCall } from "./msteams-tts-playback.js";
import { isAddressed } from "./group-call-gate.js";
/** Default energy VAD / echo-guard tuning. */
const DEFAULT_VAD_ENERGY_THRESHOLD = 0.02;
const DEFAULT_VAD_SILENCE_MS = 700;
const DEFAULT_MIN_UTTERANCE_MS = 300;
const DEFAULT_ECHO_BARGE_IN_RMS = 0.15;
/** Cap a single utterance so a continuously-noisy leg can't grow an unbounded buffer (~30 s). */
const MAX_UTTERANCE_MS = 30_000;
function frameRms(pcm) {
    const n = Math.floor(pcm.length / 2);
    if (n === 0)
        return 0;
    let sumSq = 0;
    for (let i = 0; i < n; i++) {
        const s = pcm.readInt16LE(i * 2) / 32768;
        sumSq += s * s;
    }
    return Math.sqrt(sumSq / n);
}
function frameMs(pcm) {
    return (pcm.length / 2 / MSTEAMS_PCM_SAMPLE_RATE_HZ) * 1000;
}
/**
 * Create a streaming (non-realtime) bridge for one Teams call. Returns the same {@link
 * MsteamsRealtimeCall} surface the runtime drives, so inbound media callbacks and the per-call map
 * work uniformly across both paths.
 */
export function createMsteamsStreamingCall(params) {
    const { session, deps } = params;
    const log = deps.logger;
    const now = deps.now ?? (() => Date.now());
    const energyThreshold = deps.vadEnergyThreshold ?? DEFAULT_VAD_ENERGY_THRESHOLD;
    const silenceMsLimit = deps.vadSilenceMs ?? DEFAULT_VAD_SILENCE_MS;
    const minUtteranceMs = deps.minUtteranceMs ?? DEFAULT_MIN_UTTERANCE_MS;
    const bargeInRms = deps.echoBargeInRms ?? DEFAULT_ECHO_BARGE_IN_RMS;
    const echoGuard = deps.suppressInputDuringPlayback !== false;
    const requireRecording = deps.requireRecordingStatus !== false;
    const transcript = [];
    const playback = {
        providerCallId: session.callId,
        session,
        ttsAbort: null,
        turnId: 0,
        outboundSeq: 0,
        outboundTimestampMs: 0,
        lastOutboundFrameAt: 0,
    };
    let closed = false;
    let speaking = false;
    let processing = false;
    let greeted = false;
    let recordingActive = false;
    let currentSpeaker;
    let humanCount = 1;
    let lastAddressedAt;
    // Group-call gate is active only in a meeting with a usable wake phrase (mirrors the realtime path).
    const gateActive = !!deps.groupCallGate &&
        deps.groupCallGate.requireAddress &&
        deps.groupCallGate.wakePhrases.some((p) => p.trim().length > 0);
    /** In a meeting, is this caller turn for the bot (wake phrase, or within the follow-up window)? */
    function addressed(text) {
        if (!gateActive || humanCount < 2)
            return true; // 1:1 (or no gate) is never gated
        const gate = deps.groupCallGate;
        if (isAddressed(text, gate.wakePhrases)) {
            lastAddressedAt = now();
            return true;
        }
        return lastAddressedAt !== undefined && now() - lastAddressedAt <= gate.followUpWindowMs;
    }
    // VAD accumulation.
    let speechBuffers = [];
    let bufferedMs = 0;
    let inSpeech = false;
    let trailingSilenceMs = 0;
    function resetUtterance() {
        speechBuffers = [];
        bufferedMs = 0;
        inSpeech = false;
        trailingSilenceMs = 0;
    }
    function cancelPlayback() {
        if (!speaking)
            return;
        playback.ttsAbort?.abort();
        playback.ttsAbort = null;
        speaking = false;
        // Tell the worker to flush its playout queue.
        session.send({ type: "assistant.cancel", turnId: playback.turnId });
    }
    async function speak(text) {
        if (closed || !text.trim())
            return;
        speaking = true;
        try {
            await playTtsToCall({ ttsProvider: deps.ttsProvider, logger: log }, playback, text);
        }
        catch (err) {
            // Socket closed mid-playback (caller hung up) or synthesis failed — surface, don't crash.
            log?.warn(`[msteams-voice] streaming playback failed for ${session.callId}: ${String(err)}`);
        }
        finally {
            if (playback.ttsAbort === null)
                speaking = false;
        }
    }
    /** Run one turn: caller question → agent → spoken reply. */
    async function runTurn(question, opts) {
        if (closed || processing)
            return;
        const attributed = currentSpeaker ? `${currentSpeaker}: ${question}` : question;
        // Group-call gate: record the turn (for context / minutes) but don't answer unless addressed.
        if (opts?.gated && !addressed(question)) {
            transcript.push({ role: "user", text: attributed });
            deps.appendTranscript?.({ role: "caller", text: attributed, at: now() });
            log?.debug?.(`[msteams-voice] streaming: caller turn not addressed to the bot — not answering (${session.callId})`);
            return;
        }
        processing = true;
        try {
            transcript.push({ role: "user", text: attributed });
            deps.appendTranscript?.({ role: "caller", text: attributed, at: now() });
            const images = deps.getVisionImages?.();
            const { text } = await deps.consult({
                question,
                transcript: [...transcript],
                ...(images && images.length ? { images } : {}),
            });
            if (closed)
                return;
            const reply = text.trim();
            if (!reply)
                return;
            transcript.push({ role: "assistant", text: reply });
            deps.appendTranscript?.({ role: "bot", text: reply, at: now() });
            await speak(reply);
        }
        catch (err) {
            log?.warn(`[msteams-voice] streaming turn failed for ${session.callId}: ${String(err)}`);
        }
        finally {
            processing = false;
        }
    }
    async function endUtterance() {
        const transcribe = deps.transcribe;
        const pcm = Buffer.concat(speechBuffers);
        const lengthMs = bufferedMs;
        resetUtterance();
        if (!transcribe)
            return; // session mode — no file fallback configured
        if (lengthMs < minUtteranceMs || pcm.length === 0)
            return;
        if (processing || speaking)
            return; // ignore overlapping speech while busy
        let text = "";
        try {
            text = (await transcribe(pcm)).trim();
        }
        catch (err) {
            log?.warn(`[msteams-voice] streaming STT failed for ${session.callId}: ${String(err)}`);
            return;
        }
        if (text)
            await runTurn(text, { gated: true });
    }
    function maybeGreet() {
        if (greeted || closed || !deps.greetingInstruction)
            return;
        greeted = true;
        void runTurn(deps.greetingInstruction);
    }
    // Preferred STT: a live streaming session. Final transcripts drive turns; partials are unused (we
    // do energy-based barge-in below). When absent we fall back to the energy-VAD + `transcribe` path.
    const sttSession = deps.createTranscriptionSession?.({
        onTranscript: (t) => {
            const text = t.trim();
            if (text)
                void runTurn(text, { gated: true });
        },
        onError: (e) => log?.warn(`[msteams-voice] streaming STT session error for ${session.callId}: ${e.message}`),
    });
    if (sttSession) {
        void sttSession
            .connect()
            .catch((e) => log?.warn(`[msteams-voice] streaming STT connect failed for ${session.callId}: ${String(e)}`));
    }
    return {
        pushAudio: (pcm16k) => {
            if (closed)
                return;
            if (requireRecording && !recordingActive)
                return; // Media Access gate
            // When recording isn't required there is no recording.status event to greet on — open on the
            // first inbound audio instead (fires once).
            if (!requireRecording)
                maybeGreet();
            const rms = frameRms(pcm16k);
            if (speaking) {
                // Echo guard: while we speak, only audio loud enough to be a genuine barge-in counts.
                if (echoGuard && rms < bargeInRms)
                    return;
                // Barge-in: caller interrupted → stop playback and start capturing their utterance.
                cancelPlayback();
            }
            // While the agent/STT is processing, don't feed audio (we ignore overlapping turns anyway).
            if (processing)
                return;
            if (sttSession) {
                // Live STT: stream the frame; the final transcript arrives via the onTranscript callback.
                if (sttSession.isConnected())
                    sttSession.sendAudio(pcm16k);
                return;
            }
            // Fallback: energy-VAD utterance segmentation → file STT.
            const ms = frameMs(pcm16k);
            if (rms >= energyThreshold) {
                speechBuffers.push(pcm16k);
                bufferedMs += ms;
                inSpeech = true;
                trailingSilenceMs = 0;
                if (bufferedMs >= MAX_UTTERANCE_MS)
                    void endUtterance();
            }
            else if (inSpeech) {
                // Trailing silence after speech — keep it (natural tail), end the utterance once long enough.
                speechBuffers.push(pcm16k);
                bufferedMs += ms;
                trailingSilenceMs += ms;
                if (trailingSilenceMs >= silenceMsLimit)
                    void endUtterance();
            }
        },
        // Frames are pulled at turn time via deps.getVisionImages (attached to the consult), so there's
        // nothing to push on each inbound frame here.
        notifyInboundFrame: () => { },
        setHumanCount: (count) => {
            humanCount = count;
        },
        notifyDtmf: (digit) => {
            // Surface a DTMF key as a caller turn so the agent can run simple IVR flows.
            if (closed || processing || speaking)
                return;
            void runTurn(`The caller pressed the key "${digit}".`);
        },
        setCurrentSpeaker: (name) => {
            currentSpeaker = name;
        },
        setRecordingActive: (active) => {
            recordingActive = active;
            if (active)
                maybeGreet();
        },
        say: (text) => {
            // H4: speak the worker-provided line (e.g. a goodbye before a limit cutoff) through the existing
            // TTS path. `speak` already no-ops when closed or when the text is blank.
            void speak(text);
        },
        close: (reason) => {
            if (closed)
                return;
            closed = true;
            cancelPlayback();
            resetUtterance();
            sttSession?.close();
            if (reason)
                session.close(reason);
        },
    };
}
