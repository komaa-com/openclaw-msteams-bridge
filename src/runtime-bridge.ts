// Wiring a Teams call's WebSocket to a realtime voice bridge using ONLY the public SDK + api.runtime.
// 🚧 SCAFFOLD — illustrative; completed when the CVI files (msteams-media-stream, msteams-realtime)
// are ported in (see PORTING.md). The point: NO voice-call internals are needed here.

import { createRealtimeVoiceBridgeSession } from "openclaw/plugin-sdk/realtime-voice";

// Placeholders — replace with the real types when porting.
type ResolvedRealtimeProvider = unknown;
type TeamsSession = { send: (msg: unknown) => void; close: (reason?: string) => void };

export function bindTeamsCallToRealtime(params: {
  /** Teams WS session from msteams-media-stream. */
  session: TeamsSession;
  /** Resolved realtime voice provider (e.g. openai realtime). */
  realtimeProvider: ResolvedRealtimeProvider;
  // api: OpenClawPluginApi;  // for api.runtime.agent / tts / state
}): void {
  // 1) audioSink: realtime model audio OUT -> Teams WS
  const audioSink = {
    sendAudio: (pcm: Buffer) =>
      params.session.send({ type: "audio.frame", payloadBase64: pcm.toString("base64") }),
    clearAudio: () => params.session.send({ type: "assistant.cancel" }),
  };

  // 2) public SDK does the realtime plumbing:
  //    const bridge = createRealtimeVoiceBridgeSession({ provider, audioSink, onTranscript, onToolCall, ... });
  // 3) inbound Teams audio -> bridge.sendAudio(pcm); barge-in -> bridge.handleBargeIn(); vision -> sendImage
  // 4) consult tool -> consultRealtimeVoiceAgent(...) / api.runtime.agent.runEmbeddedAgent(...)

  void createRealtimeVoiceBridgeSession; // referenced; real wiring is TODO
  void audioSink;
  throw new Error("bindTeamsCallToRealtime: not implemented (scaffold)");
}
