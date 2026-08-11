import { describe, expect, it, vi } from "vitest";
import { type ConsultImage, withConsultImages } from "./vision-consult.js";

/**
 * Consult images must survive a host whose consult wrapper does not know the `images` param.
 *
 * The regression: every published openclaw through 2026.7.1 types `consultRealtimeVoiceAgent` WITHOUT
 * `images`. The plugin passed one anyway; JavaScript silently discarded the unknown property; every
 * look_at_screen ran text-only, and the vision model truthfully answered "I'm not receiving the
 * shared-screen image". By then the media path had been debugged leg by leg for a day - subscribes,
 * gates, bridge, all working - because a silently dropped parameter produces exactly the same symptom
 * as a broken camera.
 *
 * withConsultImages fixes it at the layer that DOES understand images (runEmbeddedAgent), so these
 * tests simulate the host: a consult wrapper that forwards everything EXCEPT images.
 */
describe("withConsultImages", () => {
  const IMG: ConsultImage[] = [{ type: "image", data: "base64bytes", mimeType: "image/jpeg" }];

  /** A published host: builds its own runEmbeddedAgent params and never forwards images. */
  function publishedHostConsult(agentRuntime: {
    runEmbeddedAgent: (p: Record<string, unknown>) => unknown;
  }) {
    return agentRuntime.runEmbeddedAgent({ sessionId: "s1", prompt: "look at this" });
  }

  it("injects the images at the runEmbeddedAgent layer, through a host that drops the param", () => {
    const run = vi.fn().mockReturnValue("ok");
    const wrapped = withConsultImages({ runEmbeddedAgent: run } as never, IMG);
    publishedHostConsult(wrapped as never);
    expect(run).toHaveBeenCalledWith({ sessionId: "s1", prompt: "look at this", images: IMG });
  });

  it("does not double-attach on a future host that forwards images itself", () => {
    const run = vi.fn().mockReturnValue("ok");
    const wrapped = withConsultImages({ runEmbeddedAgent: run } as never, IMG);
    const hostImages = [{ type: "image", data: "host-supplied", mimeType: "image/png" }];
    (wrapped as { runEmbeddedAgent: (p: unknown) => unknown }).runEmbeddedAgent({
      sessionId: "s1",
      images: hostImages,
    });
    // The host's own images win untouched - no merge, no overwrite.
    expect(run).toHaveBeenCalledWith({ sessionId: "s1", images: hostImages });
  });

  it("is the identity when there is nothing to attach", () => {
    const runtime = { runEmbeddedAgent: vi.fn() };
    expect(withConsultImages(runtime as never, [])).toBe(runtime);
    expect(withConsultImages(runtime as never, undefined)).toBe(runtime);
  });

  it("leaves every other runtime member reachable on the wrapped object", () => {
    // The consult wrapper uses agentRuntime for more than the run (session store, timeouts, thinking
    // defaults). Wrapping must not hide any of that.
    const runtime = {
      runEmbeddedAgent: vi.fn(),
      resolveThinkingDefault: vi.fn().mockReturnValue("high"),
      defaults: { provider: "openai", model: "gpt-5.5" },
    };
    const wrapped = withConsultImages(runtime as never, IMG) as unknown as typeof runtime;
    expect(wrapped.resolveThinkingDefault({})).toBe("high");
    expect(wrapped.defaults).toEqual({ provider: "openai", model: "gpt-5.5" });
  });
});
