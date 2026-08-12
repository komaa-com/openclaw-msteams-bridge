import { describe, expect, it, vi } from "vitest";
import { fetchAttachmentAudio } from "./managed-chat.js";

/**
 * Fetching inbound voice messages.
 *
 * Same posture as the image fetcher and for the same reasons: the attachment URL is gateway-SIGNED,
 * but that signature is verified BY the gateway - it proves nothing on this side, and the URL arrives
 * inside a message someone else wrote. Origin pinning is what actually bounds where this fetch can go.
 */
describe("fetchAttachmentAudio", () => {
  const GATEWAY = "https://teams.standin.komaa.com/api/chat/reply";
  const clip = (over: Record<string, unknown> = {}) => ({
    kind: "audio",
    name: "voice.ogg",
    url: "https://teams.standin.komaa.com/att/1",
    ...over,
  });
  const ok = (body: string, mime = "audio/ogg") =>
    vi.fn(async () => new Response(body, { headers: { "content-type": mime } })) as unknown as typeof fetch;

  it("fetches an audio attachment from the pinned gateway", async () => {
    const got = await fetchAttachmentAudio([clip()] as never, {
      fetchFn: ok("OGGDATA"),
      gatewayOrigin: GATEWAY,
    });
    expect(got).toHaveLength(1);
    expect(got[0]!.bytes.toString()).toBe("OGGDATA");
    expect(got[0]!.name).toBe("voice.ogg");
  });

  it("refuses a URL that is not on the gateway", async () => {
    // The whole SSRF guard: a message can name any URL it likes.
    const fetchFn = ok("EVIL");
    const got = await fetchAttachmentAudio([clip({ url: "https://evil.example/x" })] as never, {
      fetchFn,
      gatewayOrigin: GATEWAY,
    });
    expect(got).toHaveLength(0);
    expect(fetchFn).not.toHaveBeenCalled(); // refused BEFORE the request, not after
  });

  it("ignores non-audio attachments and unrelayable ones", async () => {
    const got = await fetchAttachmentAudio(
      [clip({ kind: "image" }), clip({ relayable: false }), clip({ url: undefined })] as never,
      { fetchFn: ok("X"), gatewayOrigin: GATEWAY },
    );
    expect(got).toHaveLength(0);
  });

  it("refuses a body whose content-type is not audio, before reading it", async () => {
    const got = await fetchAttachmentAudio([clip()] as never, {
      fetchFn: ok("<html>", "text/html"),
      gatewayOrigin: GATEWAY,
    });
    expect(got).toHaveLength(0);
  });

  it("caps the number of clips so one message cannot multiply into unbounded STT calls", async () => {
    const many = Array.from({ length: 6 }, (_, i) => clip({ url: `https://teams.standin.komaa.com/att/${i}` }));
    const got = await fetchAttachmentAudio(many as never, {
      fetchFn: ok("A"),
      gatewayOrigin: GATEWAY,
      maxClips: 2,
    });
    expect(got).toHaveLength(2);
  });

  it("refuses a clip that declares itself over the byte cap", async () => {
    const fetchFn = vi.fn(
      async () =>
        new Response("A".repeat(50), {
          headers: { "content-type": "audio/ogg", "content-length": "999999999" },
        }),
    ) as unknown as typeof fetch;
    const got = await fetchAttachmentAudio([clip()] as never, {
      fetchFn,
      gatewayOrigin: GATEWAY,
      maxBytes: 1024,
    });
    expect(got).toHaveLength(0);
  });
});
