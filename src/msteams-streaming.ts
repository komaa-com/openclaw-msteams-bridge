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

import type { RealtimeVoiceAgentConsultTranscriptEntry } from "openclaw/plugin-sdk/realtime-voice";
import type {
  RealtimeTranscriptionSession,
  RealtimeTranscriptionSessionCallbacks,
} from "openclaw/plugin-sdk/realtime-transcription";
import {
  MSTEAMS_PCM_SAMPLE_RATE_HZ,
  type MsteamsLogger,
  type MsteamsSession,
} from "./msteams-media-stream.js";
import type { MsteamsRealtimeCall } from "./msteams-realtime.js";
import type { MsteamsTtsProvider } from "./msteams-tts.js";
import { playTtsToCall, type TtsPlaybackTarget } from "./msteams-tts-playback.js";
import { type GroupCallGateConfig, isAddressed } from "./group-call-gate.js";
import type { ConsultImage } from "./vision-consult.js";

/** Default energy VAD / echo-guard tuning. */
const DEFAULT_VAD_ENERGY_THRESHOLD = 0.02;
const DEFAULT_VAD_SILENCE_MS = 700;
const DEFAULT_MIN_UTTERANCE_MS = 300;
const DEFAULT_ECHO_BARGE_IN_RMS = 0.15;
/** Cap a single utterance so a continuously-noisy leg can't grow an unbounded buffer (~30 s). */
const MAX_UTTERANCE_MS = 30_000;

export interface MsteamsStreamingDeps {
  /**
   * Fallback STT: transcribe one VAD-segmented utterance (PCM 16 kHz, 16-bit mono LE) → text. Used
   * when {@link createTranscriptionSession} is not supplied (no streaming transcription provider
   * resolved). Optional, but at least one of `transcribe` / `createTranscriptionSession` is required.
   */
  transcribe?: (pcm16k: Buffer) => Promise<string>;
  /**
   * Preferred STT: live streaming transcription session (lower latency). When supplied, inbound PCM
   * is streamed to the session and a turn fires on each FINAL transcript; the energy-VAD +
   * {@link transcribe} path is bypassed. The runtime wires this to the configured realtime
   * transcription provider's `createSession`.
   */
  createTranscriptionSession?: (
    callbacks: RealtimeTranscriptionSessionCallbacks,
  ) => RealtimeTranscriptionSession;
  /** Run the agent for a caller question (with prior transcript + any shared frames) → reply text. */
  consult: (params: {
    question: string;
    transcript: RealtimeVoiceAgentConsultTranscriptEntry[];
    images?: ConsultImage[];
  }) => Promise<{ text: string }>;
  ttsProvider: MsteamsTtsProvider;
  /**
   * Group-call gating: in a meeting (>= 2 humans) with a usable wake phrase, only answer a caller
   * turn that addresses the bot (or is within the follow-up window). 1:1 calls are never gated.
   */
  groupCallGate?: GroupCallGateConfig;
  /**
   * Latest shared frame(s) as consult images, attached to each agent turn so the streaming agent has
   * the same "look at what's shared" awareness as the realtime path. Runtime supplies (vision store +
   * budget); omitted → audio-only.
   */
  getVisionImages?: () => ConsultImage[];
  /**
   * Optional opening line the agent speaks first (model speaks first). Phrased as an instruction
   * (e.g. "Greet the caller and ask how you can help") — run through `consult` to produce real words.
   */
  greetingInstruction?: string;
  /** Half-duplex echo guard (default true): while speaking, ignore inbound below the barge-in gate. */
  suppressInputDuringPlayback?: boolean;
  /** Barge-in RMS gate (0..1) while speaking; default {@link DEFAULT_ECHO_BARGE_IN_RMS}. */
  echoBargeInRms?: number;
  /** VAD speech-energy gate (0..1); default {@link DEFAULT_VAD_ENERGY_THRESHOLD}. */
  vadEnergyThreshold?: number;
  /** End-of-utterance trailing silence (ms); default {@link DEFAULT_VAD_SILENCE_MS}. */
  vadSilenceMs?: number;
  /** Minimum utterance length (ms) worth transcribing; default {@link DEFAULT_MIN_UTTERANCE_MS}. */
  minUtteranceMs?: number;
  /** Require active Teams recording before processing audio (Media Access API); default true. */
  requireRecordingStatus?: boolean;
  /** Persist transcript turns (e.g. into CallLifecycle). */
  appendTranscript?: (entry: { role: "caller" | "bot"; text: string; at: number }) => void;
  logger?: MsteamsLogger;
  /** Clock injection for tests. */
  now?: () => number;
}

function frameRms(pcm: Buffer): number {
  const n = Math.floor(pcm.length / 2);
  if (n === 0) return 0;
  let sumSq = 0;
  for (let i = 0; i < n; i++) {
    const s = pcm.readInt16LE(i * 2) / 32768;
    sumSq += s * s;
  }
  return Math.sqrt(sumSq / n);
}

function frameMs(pcm: Buffer): number {
  return (pcm.length / 2 / MSTEAMS_PCM_SAMPLE_RATE_HZ) * 1000;
}

/**
 * Create a streaming (non-realtime) bridge for one Teams call. Returns the same {@link
 * MsteamsRealtimeCall} surface the runtime drives, so inbound media callbacks and the per-call map
 * work uniformly across both paths.
 */
export function createMsteamsStreamingCall(params: {
  session: MsteamsSession;
  deps: MsteamsStreamingDeps;
}): MsteamsRealtimeCall {
  const { session, deps } = params;
  const log = deps.logger;
  const now = deps.now ?? (() => Date.now());

  const energyThreshold = deps.vadEnergyThreshold ?? DEFAULT_VAD_ENERGY_THRESHOLD;
  const silenceMsLimit = deps.vadSilenceMs ?? DEFAULT_VAD_SILENCE_MS;
  const minUtteranceMs = deps.minUtteranceMs ?? DEFAULT_MIN_UTTERANCE_MS;
  const bargeInRms = deps.echoBargeInRms ?? DEFAULT_ECHO_BARGE_IN_RMS;
  const echoGuard = deps.suppressInputDuringPlayback !== false;
  const requireRecording = deps.requireRecordingStatus !== false;

  const transcript: RealtimeVoiceAgentConsultTranscriptEntry[] = [];
  const playback: TtsPlaybackTarget = {
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
  let currentSpeaker: string | undefined;
  let humanCount = 1;
  let lastAddressedAt: number | undefined;
  // Group-call gate is active only in a meeting with a usable wake phrase (mirrors the realtime path).
  const gateActive =
    !!deps.groupCallGate &&
    deps.groupCallGate.requireAddress &&
    deps.groupCallGate.wakePhrases.some((p) => p.trim().length > 0);

  /** In a meeting, is this caller turn for the bot (wake phrase, or within the follow-up window)? */
  function addressed(text: string): boolean {
    if (!gateActive || humanCount < 2) return true; // 1:1 (or no gate) is never gated
    const gate = deps.groupCallGate!;
    if (isAddressed(text, gate.wakePhrases)) {
      lastAddressedAt = now();
      return true;
    }
    return lastAddressedAt !== undefined && now() - lastAddressedAt <= gate.followUpWindowMs;
  }

  // VAD accumulation.
  let speechBuffers: Buffer[] = [];
  let bufferedMs = 0;
  let inSpeech = false;
  let trailingSilenceMs = 0;

  function resetUtterance(): void {
    speechBuffers = [];
    bufferedMs = 0;
    inSpeech = false;
    trailingSilenceMs = 0;
  }

  function cancelPlayback(): void {
    if (!speaking) return;
    playback.ttsAbort?.abort();
    playback.ttsAbort = null;
    speaking = false;
    // Tell the worker to flush its playout queue.
    session.send({ type: "assistant.cancel", turnId: playback.turnId });
  }

  async function speak(text: string): Promise<void> {
    if (closed || !text.trim()) return;
    speaking = true;
    try {
      await playTtsToCall({ ttsProvider: deps.ttsProvider, logger: log }, playback, text);
    } catch (err) {
      // Socket closed mid-playback (caller hung up) or synthesis failed — surface, don't crash.
      log?.warn(`[msteams-voice] streaming playback failed for ${session.callId}: ${String(err)}`);
    } finally {
      if (playback.ttsAbort === null) speaking = false;
    }
  }

  /** Run one turn: caller question → agent → spoken reply. */
  async function runTurn(question: string, opts?: { gated?: boolean }): Promise<void> {
    if (closed || processing) return;
    const attributed = currentSpeaker ? `${currentSpeaker}: ${question}` : question;
    // Group-call gate: record the turn (for context / minutes) but don't answer unless addressed.
    if (opts?.gated && !addressed(question)) {
      transcript.push({ role: "user", text: attributed });
      deps.appendTranscript?.({ role: "caller", text: attributed, at: now() });
      log?.debug?.(
        `[msteams-voice] streaming: caller turn not addressed to the bot — not answering (${session.callId})`,
      );
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
      if (closed) return;
      const reply = text.trim();
      if (!reply) return;
      transcript.push({ role: "assistant", text: reply });
      deps.appendTranscript?.({ role: "bot", text: reply, at: now() });
      await speak(reply);
    } catch (err) {
      log?.warn(`[msteams-voice] streaming turn failed for ${session.callId}: ${String(err)}`);
    } finally {
      processing = false;
    }
  }

  async function endUtterance(): Promise<void> {
    const transcribe = deps.transcribe;
    const pcm = Buffer.concat(speechBuffers);
    const lengthMs = bufferedMs;
    resetUtterance();
    if (!transcribe) return; // session mode — no file fallback configured
    if (lengthMs < minUtteranceMs || pcm.length === 0) return;
    if (processing || speaking) return; // ignore overlapping speech while busy
    let text = "";
    try {
      text = (await transcribe(pcm)).trim();
    } catch (err) {
      log?.warn(`[msteams-voice] streaming STT failed for ${session.callId}: ${String(err)}`);
      return;
    }
    if (text) await runTurn(text, { gated: true });
  }

  function maybeGreet(): void {
    if (greeted || closed || !deps.greetingInstruction) return;
    greeted = true;
    void runTurn(deps.greetingInstruction);
  }

  // Preferred STT: a live streaming session. Final transcripts drive turns; partials are unused (we
  // do energy-based barge-in below). When absent we fall back to the energy-VAD + `transcribe` path.
  const sttSession: RealtimeTranscriptionSession | undefined = deps.createTranscriptionSession?.({
    onTranscript: (t: string) => {
      const text = t.trim();
      if (text) void runTurn(text, { gated: true });
    },
    onError: (e: Error) =>
      log?.warn(`[msteams-voice] streaming STT session error for ${session.callId}: ${e.message}`),
  });
  if (sttSession) {
    void sttSession
      .connect()
      .catch((e) =>
        log?.warn(`[msteams-voice] streaming STT connect failed for ${session.callId}: ${String(e)}`),
      );
  }

  return {
    pushAudio: (pcm16k: Buffer) => {
      if (closed) return;
      if (requireRecording && !recordingActive) return; // Media Access gate
      // When recording isn't required there is no recording.status event to greet on — open on the
      // first inbound audio instead (fires once).
      if (!requireRecording) maybeGreet();
      const rms = frameRms(pcm16k);

      if (speaking) {
        // Echo guard: while we speak, only audio loud enough to be a genuine barge-in counts.
        if (echoGuard && rms < bargeInRms) return;
        // Barge-in: caller interrupted → stop playback and start capturing their utterance.
        cancelPlayback();
      }
      // While the agent/STT is processing, don't feed audio (we ignore overlapping turns anyway).
      if (processing) return;

      if (sttSession) {
        // Live STT: stream the frame; the final transcript arrives via the onTranscript callback.
        if (sttSession.isConnected()) sttSession.sendAudio(pcm16k);
        return;
      }

      // Fallback: energy-VAD utterance segmentation → file STT.
      const ms = frameMs(pcm16k);
      if (rms >= energyThreshold) {
        speechBuffers.push(pcm16k);
        bufferedMs += ms;
        inSpeech = true;
        trailingSilenceMs = 0;
        if (bufferedMs >= MAX_UTTERANCE_MS) void endUtterance();
      } else if (inSpeech) {
        // Trailing silence after speech — keep it (natural tail), end the utterance once long enough.
        speechBuffers.push(pcm16k);
        bufferedMs += ms;
        trailingSilenceMs += ms;
        if (trailingSilenceMs >= silenceMsLimit) void endUtterance();
      }
    },

    // Frames are pulled at turn time via deps.getVisionImages (attached to the consult), so there's
    // nothing to push on each inbound frame here.
    notifyInboundFrame: () => {},

    setHumanCount: (count: number) => {
      humanCount = count;
    },

    notifyDtmf: (digit: string) => {
      // Surface a DTMF key as a caller turn so the agent can run simple IVR flows.
      if (closed || processing || speaking) return;
      void runTurn(`The caller pressed the key "${digit}".`);
    },

    setCurrentSpeaker: (name: string | undefined) => {
      currentSpeaker = name;
    },

    setRecordingActive: (active: boolean) => {
      recordingActive = active;
      if (active) maybeGreet();
    },

    close: (reason?: string) => {
      if (closed) return;
      closed = true;
      cancelPlayback();
      resetUtterance();
      sttSession?.close();
      if (reason) session.close(reason);
    },
  };
}
