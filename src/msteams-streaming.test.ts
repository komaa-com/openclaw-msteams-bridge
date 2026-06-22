import { describe, expect, it, vi } from "vitest";
import type { MsteamsSession } from "./msteams-media-stream.js";
import { createMsteamsStreamingCall, type MsteamsStreamingDeps } from "./msteams-streaming.js";
import type { MsteamsTtsProvider } from "./msteams-tts.js";

/** 20 ms PCM 16 kHz mono frame (320 samples) at a given amplitude. amplitude 0 ⇒ silence. */
function frame(amplitude: number): Buffer {
  const samples = 320;
  const buf = Buffer.alloc(samples * 2);
  for (let i = 0; i < samples; i++) {
    // Alternate sign so it's a real waveform, not DC.
    buf.writeInt16LE(i % 2 === 0 ? amplitude : -amplitude, i * 2);
  }
  return buf;
}
const LOUD = () => frame(8000); // rms ≈ 0.24
const SILENT = () => frame(0);

function fakeSession(): MsteamsSession & { sent: unknown[] } {
  const sent: unknown[] = [];
  return {
    callId: "call-1",
    threadId: "thread-1",
    caller: { aadId: "caller-1", displayName: "Caller One" },
    sent,
    send: (m: unknown) => {
      sent.push(m);
      return true;
    },
    close: vi.fn(),
  } as MsteamsSession & { sent: unknown[] };
}

function fakeTts(frames = 1): MsteamsTtsProvider {
  return {
    synthesizePcm16k: vi.fn(async () => Buffer.alloc(frames * 640)),
  };
}

function sentTypes(session: { sent: unknown[] }): string[] {
  return session.sent.map((m) => (m as { type?: string }).type ?? "");
}

const baseDeps = (over: Partial<MsteamsStreamingDeps>): MsteamsStreamingDeps => ({
  transcribe: vi.fn(async () => "hello there"),
  consult: vi.fn(async () => ({ text: "Hi, how can I help?" })),
  ttsProvider: fakeTts(1),
  // Tiny VAD windows so a couple of frames complete an utterance.
  vadSilenceMs: 40,
  minUtteranceMs: 20,
  requireRecordingStatus: false,
  ...over,
});

describe("createMsteamsStreamingCall", () => {
  it("runs one full turn: utterance → STT → agent → TTS playback", async () => {
    const session = fakeSession();
    const transcribe = vi.fn(async () => "what's the weather");
    const consult = vi.fn(async () => ({ text: "It is sunny." }));
    const appendTranscript = vi.fn();
    const call = createMsteamsStreamingCall({
      session,
      deps: baseDeps({ transcribe, consult, appendTranscript }),
    });

    // Speech (2 frames) then trailing silence (>= 40 ms) to close the utterance.
    call.pushAudio(LOUD());
    call.pushAudio(LOUD());
    call.pushAudio(SILENT());
    call.pushAudio(SILENT());
    call.pushAudio(SILENT());

    await vi.waitFor(() => expect(consult).toHaveBeenCalledTimes(1));
    // STT got the captured utterance; agent got the transcribed question.
    expect(transcribe).toHaveBeenCalledTimes(1);
    expect(consult.mock.calls[0]![0].question).toBe("what's the weather");
    // Reply was synthesized + streamed back as audio frames.
    await vi.waitFor(() => expect(sentTypes(session)).toContain("audio.frame"));
    // Both turns recorded.
    expect(appendTranscript).toHaveBeenCalledWith(
      expect.objectContaining({ role: "caller" }),
    );
    expect(appendTranscript).toHaveBeenCalledWith(expect.objectContaining({ role: "bot" }));
  });

  it("barge-in: loud caller audio during playback cancels the in-flight reply", async () => {
    const session = fakeSession();
    const call = createMsteamsStreamingCall({
      session,
      // Many frames ⇒ playback lasts long enough to interrupt.
      deps: baseDeps({ ttsProvider: fakeTts(40) }),
    });

    call.pushAudio(LOUD());
    call.pushAudio(LOUD());
    call.pushAudio(SILENT());
    call.pushAudio(SILENT());

    // Wait until playback has started (an audio.frame went out).
    await vi.waitFor(() => expect(sentTypes(session)).toContain("audio.frame"));

    // Caller interrupts loudly while we're speaking.
    call.pushAudio(LOUD());

    await vi.waitFor(() => expect(sentTypes(session)).toContain("assistant.cancel"));
  });

  it("greets first when a greeting instruction is configured", async () => {
    const session = fakeSession();
    const consult = vi.fn(async () => ({ text: "Hello! How can I help?" }));
    const call = createMsteamsStreamingCall({
      session,
      deps: baseDeps({ greetingInstruction: "Greet the caller", consult }),
    });
    // No recording required → greeting fires on first inbound audio.
    call.pushAudio(SILENT());
    await vi.waitFor(() => expect(consult).toHaveBeenCalledTimes(1));
    expect(consult.mock.calls[0]![0].question).toBe("Greet the caller");
  });

  it("requireRecordingStatus gates processing until recording is active", async () => {
    const session = fakeSession();
    const transcribe = vi.fn(async () => "hi");
    const call = createMsteamsStreamingCall({
      session,
      deps: baseDeps({ transcribe, requireRecordingStatus: true }),
    });
    // Audio before recording is active is dropped.
    call.pushAudio(LOUD());
    call.pushAudio(SILENT());
    call.pushAudio(SILENT());
    await new Promise((r) => setTimeout(r, 30));
    expect(transcribe).not.toHaveBeenCalled();

    // Once recording is active, audio is processed.
    call.setRecordingActive(true);
    call.pushAudio(LOUD());
    call.pushAudio(LOUD());
    call.pushAudio(SILENT());
    call.pushAudio(SILENT());
    await vi.waitFor(() => expect(transcribe).toHaveBeenCalledTimes(1));
  });
});
