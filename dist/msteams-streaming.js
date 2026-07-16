import { MSTEAMS_PCM_SAMPLE_RATE_HZ, } from "./msteams-media-stream.js";
import { playTtsToCall } from "./msteams-tts-playback.js";
import { isAddressed } from "./group-call-gate.js";
const DEFAULT_VAD_ENERGY_THRESHOLD = 0.02;
const DEFAULT_VAD_SILENCE_MS = 700;
const DEFAULT_MIN_UTTERANCE_MS = 300;
const DEFAULT_ECHO_BARGE_IN_RMS = 0.15;
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
    const gateActive = !!deps.groupCallGate &&
        deps.groupCallGate.requireAddress &&
        deps.groupCallGate.wakePhrases.some((p) => p.trim().length > 0);
    function addressed(text) {
        if (!gateActive || humanCount < 2)
            return true;
        const gate = deps.groupCallGate;
        if (isAddressed(text, gate.wakePhrases)) {
            lastAddressedAt = now();
            return true;
        }
        return lastAddressedAt !== undefined && now() - lastAddressedAt <= gate.followUpWindowMs;
    }
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
            log?.warn(`[msteams-voice] streaming playback failed for ${session.callId}: ${String(err)}`);
        }
        finally {
            if (playback.ttsAbort === null)
                speaking = false;
        }
    }
    async function runTurn(question, opts) {
        if (closed || processing)
            return;
        const attributed = currentSpeaker ? `${currentSpeaker}: ${question}` : question;
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
            return;
        if (lengthMs < minUtteranceMs || pcm.length === 0)
            return;
        if (processing || speaking)
            return;
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
                return;
            if (!requireRecording)
                maybeGreet();
            const rms = frameRms(pcm16k);
            if (speaking) {
                if (echoGuard && rms < bargeInRms)
                    return;
                cancelPlayback();
            }
            if (processing)
                return;
            if (sttSession) {
                if (sttSession.isConnected())
                    sttSession.sendAudio(pcm16k);
                return;
            }
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
                speechBuffers.push(pcm16k);
                bufferedMs += ms;
                trailingSilenceMs += ms;
                if (trailingSilenceMs >= silenceMsLimit)
                    void endUtterance();
            }
        },
        notifyInboundFrame: () => { },
        setHumanCount: (count) => {
            humanCount = count;
        },
        notifyDtmf: (digit) => {
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
