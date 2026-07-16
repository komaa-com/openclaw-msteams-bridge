import { setTimeout as sleep } from "node:timers/promises";
import { inferEmotion } from "./expression.js";
import { MSTEAMS_PCM_SAMPLE_RATE_HZ, } from "./msteams-media-stream.js";
import { chunkAudio } from "./telephony-audio.js";
import { estimateVisemes, visemesFromAlignment } from "./viseme-estimate.js";
const MSTEAMS_SAMPLE_RATE_HZ = MSTEAMS_PCM_SAMPLE_RATE_HZ;
const FRAME_DURATION_MS = 20;
const BYTES_PER_SAMPLE = 2;
const FRAME_BYTES = (MSTEAMS_SAMPLE_RATE_HZ / 1000) * FRAME_DURATION_MS * BYTES_PER_SAMPLE;
export async function playTtsToCall(deps, state, text) {
    state.ttsAbort?.abort();
    const abort = new AbortController();
    state.ttsAbort = abort;
    state.turnId += 1;
    try {
        const emotion = inferEmotion(text);
        deps.logger?.debug?.(`msteams-voice: expression cue '${emotion}' for ${state.providerCallId}`);
        state.session.send({ type: "expression", emotion });
    }
    catch {
    }
    const synthesis = deps.ttsProvider.synthesizePcm16kWithTiming
        ? await deps.ttsProvider.synthesizePcm16kWithTiming(text)
        : { pcm16k: await deps.ttsProvider.synthesizePcm16k(text) };
    const pcm16k = synthesis.pcm16k;
    if (abort.signal.aborted) {
        return;
    }
    if (pcm16k.length === 0) {
        throw new Error("playTts: TTS produced no audio");
    }
    try {
        const alignment = synthesis.alignment;
        let marks = alignment
            ? visemesFromAlignment(alignment.characters, alignment.startTimesSeconds)
            : [];
        if (marks.length === 0) {
            const durationMs = (pcm16k.length / BYTES_PER_SAMPLE / MSTEAMS_SAMPLE_RATE_HZ) * 1000;
            marks = estimateVisemes(text, durationMs);
        }
        if (marks.length > 0) {
            deps.logger?.debug?.(`msteams-voice: speech.marks ${marks.length} visemes (${alignment ? "aligned" : "estimated"}) for ${state.providerCallId}`);
            state.session.send({ type: "speech.marks", ts: 0, marks });
        }
    }
    catch {
    }
    await streamPcmFrames(deps, state, pcm16k, abort.signal);
    if (state.ttsAbort === abort) {
        state.ttsAbort = null;
    }
}
async function streamPcmFrames(deps, state, pcm, signal) {
    let nextFrameDueAt = Date.now() + FRAME_DURATION_MS;
    for (const frame of chunkAudio(pcm, FRAME_BYTES)) {
        if (signal.aborted) {
            return;
        }
        const delivered = state.session.send({
            type: "audio.frame",
            seq: state.outboundSeq,
            timestampMs: state.outboundTimestampMs,
            payloadBase64: frame.toString("base64"),
        });
        if (!delivered) {
            deps.logger?.warn(`msteams-voice: audio.frame dropped for ${state.providerCallId}, Teams socket closed; aborting playback`);
            throw new Error(`msteams audio send failed for ${state.providerCallId}: session socket closed`);
        }
        state.outboundSeq += 1;
        state.outboundTimestampMs += FRAME_DURATION_MS;
        state.lastOutboundFrameAt = Date.now();
        const waitMs = nextFrameDueAt - Date.now();
        if (waitMs > 0) {
            await sleep(waitMs);
        }
        nextFrameDueAt += FRAME_DURATION_MS;
    }
}
