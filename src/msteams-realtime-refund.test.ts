import { describe, expect, it, vi } from "vitest";

// End-to-end refund test: inject a bridge SESSION whose `sendImage` throws (a configured/`next` build
// HAS sendImage; here we force it to fail) so the ambient-push flow runs through the real
// pushLatestFrameToModel → pushOrQueueBridgeImage(present sendImage → throws) → catch → budget refund.
const sendImage = vi.hoisted(() =>
  vi.fn(() => {
    throw new Error("provider rejected the image");
  }),
);
vi.mock("openclaw/plugin-sdk/realtime-voice", async (importOriginal) => {
  const actual = await importOriginal<typeof import("openclaw/plugin-sdk/realtime-voice")>();
  return {
    ...actual,
    // Ignore the provider/wrapper and return a controllable session with a throwing sendImage.
    createRealtimeVoiceBridgeSession: () => ({
      connect: async () => {},
      close: () => {},
      sendAudio: () => {},
      sendImage,
      sendUserMessage: () => {},
      triggerGreeting: () => {},
      handleBargeIn: () => {},
      setMediaTimestamp: () => {},
      acknowledgeMark: () => {},
      submitToolResult: () => {},
      isConnected: () => true,
      supportsToolResultContinuation: true,
    }),
  };
});

import type { OpenClawConfig } from "openclaw/plugin-sdk/config-contracts";
import type { RealtimeVoiceProviderPlugin } from "openclaw/plugin-sdk/realtime-voice";
import type { CoreAgentDeps } from "./core-bridge.js";
import type { VoiceCallConfig } from "./config.js";
import type { MsteamsSession } from "./msteams-media-stream.js";
import { createMsteamsRealtimeCall } from "./msteams-realtime.js";
import { VisionBudget } from "./vision-budget.js";

function fakeSession(): MsteamsSession {
  return {
    callId: "call-1",
    threadId: "t1",
    caller: { aadId: "a1", displayName: "Caller" },
    recordingStatus: "active",
    send: () => true,
    close: () => {},
  } as unknown as MsteamsSession;
}

describe("ambient vision push: budget refund when a present sendImage throws", () => {
  it("refunds the per-call vision budget so a failed push doesn't consume the cap", () => {
    const frame = {
      source: "screenshare" as const,
      dataBase64: "F1",
      mime: "image/jpeg",
      width: 1,
      height: 1,
      ts: 0,
    };
    const budget = new VisionBudget(1);
    const refund = vi.spyOn(budget, "refund");

    const call = createMsteamsRealtimeCall({
      session: fakeSession(),
      deps: {
        provider: {
          id: "openai",
          label: "m",
          isConfigured: () => true,
          createBridge: () => ({}),
        } as unknown as RealtimeVoiceProviderPlugin,
        providerConfig: {},
        toolPolicy: "safe-read-only",
        agentRuntime: { resolveThinkingDefault: () => "high" } as unknown as CoreAgentDeps,
        voiceConfig: {
          realtime: {},
          agentId: "main",
          responseTimeoutMs: 5000,
        } as unknown as VoiceCallConfig,
        cfg: {} as unknown as OpenClawConfig,
        requireRecordingStatus: false,
        getLatestFrame: (s) => (s === "screenshare" ? frame : undefined),
        visionBudget: budget,
      },
    });

    // Ambient push: the present sendImage throws → the catch refunds the budget.
    call.notifyInboundFrame();

    expect(sendImage).toHaveBeenCalledTimes(1);
    expect(refund).toHaveBeenCalledWith("call-1");
    // The single budget slot was refunded — still available for a real look_at_screen.
    expect(budget.tryConsume("call-1", Date.now())).toBe(true);

    call.close();
  });
});
