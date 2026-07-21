import crypto from "node:crypto";
import { afterEach, describe, expect, it, vi } from "vitest";
import { WebSocket } from "ws";
import { MsteamsMediaStream, type MsteamsSession } from "./msteams-media-stream.js";

const SECRET = "test-shared-secret";
const PATH = "/voice/msteams/stream";

function signHmac(secret: string, ts: number, callId: string): string {
  return crypto.createHmac("sha256", secret).update(`${ts}.${callId}`).digest("hex");
}

/** Pick a port unlikely to collide. Range 31000-39999. */
function randomPort(): number {
  return 31000 + Math.floor(Math.random() * 9000);
}

async function startServer(opts: {
  port: number;
  maxConnections?: number;
  maxConnectionsPerIp?: number;
  preStartTimeoutMs?: number;
  hmacWindowMs?: number;
  onSessionStart?: (s: MsteamsSession) => void;
  onSessionEnd?: (info: { callId: string; reason: string }) => void;
  onAudioFrame?: (info: {
    callId: string;
    seq: number;
    timestampMs: number;
    payload: Buffer;
  }) => void;
  onRecordingStatus?: (info: { callId: string; status: string }) => void;
  onVideoFrame?: (info: {
    callId: string;
    source: "camera" | "screenshare";
    ts: number;
    width: number;
    height: number;
    mime: string;
    dataBase64: string;
  }) => void;
  onDtmf?: (info: { callId: string; digit: string }) => void;
  onAssistantSay?: (info: { callId: string; text: string }) => void;
}): Promise<MsteamsMediaStream> {
  const server = new MsteamsMediaStream({
    port: opts.port,
    path: PATH,
    sharedSecret: SECRET,
    maxConnections: opts.maxConnections,
    maxConnectionsPerIp: opts.maxConnectionsPerIp,
    preStartTimeoutMs: opts.preStartTimeoutMs,
    hmacWindowMs: opts.hmacWindowMs,
    onSessionStart: opts.onSessionStart,
    onSessionEnd: opts.onSessionEnd,
    onAudioFrame: opts.onAudioFrame,
    onRecordingStatus: opts.onRecordingStatus,
    onVideoFrame: opts.onVideoFrame,
    onDtmf: opts.onDtmf,
    onAssistantSay: opts.onAssistantSay,
  });
  await server.start();
  return server;
}

/** Open an authenticated WS connection for a callId (legacy header names -
 * keeps the pre-rename compatibility path under test). */
function openAuthed(port: number, callId: string): WebSocket {
  const ts = Date.now();
  return new WebSocket(`ws://127.0.0.1:${port}${PATH}/${callId}`, {
    headers: {
      "x-openclawteamsbridge-timestamp": String(ts),
      "x-openclawteamsbridge-signature": signHmac(SECRET, ts, callId),
    },
  });
}

/** Same, with the current X-StandIn-* header names. */
function openAuthedStandIn(port: number, callId: string): WebSocket {
  const ts = Date.now();
  return new WebSocket(`ws://127.0.0.1:${port}${PATH}/${callId}`, {
    headers: {
      "x-standin-timestamp": String(ts),
      "x-standin-signature": signHmac(SECRET, ts, callId),
    },
  });
}

async function waitFor(predicate: () => boolean, timeoutMs = 1000): Promise<void> {
  const start = Date.now();
  while (!predicate()) {
    if (Date.now() - start > timeoutMs) {
      throw new Error("waitFor timed out");
    }
    await new Promise<void>((resolve) => {
      setTimeout(resolve, 10);
    });
  }
}

describe("MsteamsMediaStream", () => {
  let server: MsteamsMediaStream | undefined;

  afterEach(async () => {
    if (server) {
      await server.stop();
      server = undefined;
    }
  });

  it("accepts a connection with valid HMAC + parses session.start", async () => {
    const port = randomPort();
    let receivedSession: MsteamsSession | undefined;
    server = await startServer({
      port,
      onSessionStart: (s) => {
        receivedSession = s;
      },
    });

    const callId = "call-abc";
    const ts = Date.now();
    const sig = signHmac(SECRET, ts, callId);

    const ws = new WebSocket(`ws://127.0.0.1:${port}${PATH}/${callId}`, {
      headers: {
        "x-openclawteamsbridge-timestamp": String(ts),
        "x-openclawteamsbridge-signature": sig,
      },
    });

    await new Promise<void>((resolve, reject) => {
      ws.once("open", () => resolve());
      ws.once("error", reject);
    });

    ws.send(
      JSON.stringify({
        type: "session.start",
        callId,
        threadId: "thread-xyz",
        caller: { aadId: "aad-1", displayName: "Alice", tenantId: "tenant-1" },
      }),
    );

    await waitFor(() => receivedSession !== undefined);

    expect(receivedSession?.callId).toBe(callId);
    expect(receivedSession?.threadId).toBe("thread-xyz");
    expect(receivedSession?.caller.displayName).toBe("Alice");
    expect(receivedSession?.caller.aadId).toBe("aad-1");
    expect(server.sessionCount).toBe(1);

    ws.close();
  });

  it("accepts a connection with the X-StandIn-* header names", async () => {
    const port = randomPort();
    let receivedSession: MsteamsSession | undefined;
    server = await startServer({
      port,
      onSessionStart: (s) => {
        receivedSession = s;
      },
    });

    const callId = "call-standin-headers";
    const ws = openAuthedStandIn(port, callId);
    await new Promise<void>((resolve, reject) => {
      ws.once("open", () => resolve());
      ws.once("error", reject);
    });
    ws.send(
      JSON.stringify({
        type: "session.start",
        callId,
        threadId: "thread-si",
        caller: { aadId: "aad-1", displayName: "Alice", tenantId: "tenant-1" },
      }),
    );
    await waitFor(() => receivedSession !== undefined);
    expect(receivedSession?.callId).toBe(callId);

    ws.close();
  });

  it('normalizes unknown session.start direction (e.g. "join" from meeting joins) to inbound', async () => {
    // The hosted bridge sends direction:"join" when the bot joins a meeting; the protocol
    // enum only has inbound|outbound. Rejecting the message killed the whole session.
    const port = randomPort();
    let receivedSession: MsteamsSession | undefined;
    server = await startServer({
      port,
      onSessionStart: (s) => {
        receivedSession = s;
      },
    });

    const callId = "call-join";
    const ts = Date.now();
    const ws = new WebSocket(`ws://127.0.0.1:${port}${PATH}/${callId}`, {
      headers: {
        "x-openclawteamsbridge-timestamp": String(ts),
        "x-openclawteamsbridge-signature": signHmac(SECRET, ts, callId),
      },
    });
    await new Promise<void>((resolve, reject) => {
      ws.once("open", () => resolve());
      ws.once("error", reject);
    });

    ws.send(
      JSON.stringify({
        type: "session.start",
        callId,
        threadId: "thread-meet",
        direction: "join",
        caller: { aadId: "aad-2" },
      }),
    );

    await waitFor(() => receivedSession !== undefined);
    expect(receivedSession?.direction).toBe("inbound");

    ws.close();
  });

  it("session.send signals delivery: true while open, false once the socket has closed", async () => {
    // streamPcmFrames relies on this to abort playback when a caller hangs up mid-frame, instead of
    // advancing seq/timestamps and reporting audio as delivered on a dead socket.
    const port = randomPort();
    let session: MsteamsSession | undefined;
    server = await startServer({
      port,
      onSessionStart: (s) => {
        session = s;
      },
    });

    const callId = "call-send-status";
    const ws = openAuthed(port, callId);
    await new Promise<void>((resolve, reject) => {
      ws.once("open", () => resolve());
      ws.once("error", reject);
    });
    ws.send(
      JSON.stringify({
        type: "session.start",
        callId,
        threadId: "thread-1",
        caller: { aadId: "aad-1", displayName: "Alice", tenantId: "tenant-1" },
      }),
    );
    await waitFor(() => session !== undefined);

    // Open socket → the frame is delivered.
    expect(
      session?.send({ type: "audio.frame", seq: 0, timestampMs: 0, payloadBase64: "AA==" }),
    ).toBe(true);

    // Closed socket → the send is dropped and reported as not delivered.
    ws.close();
    await waitFor(() => server?.sessionCount === 0);
    expect(
      session?.send({ type: "audio.frame", seq: 1, timestampMs: 20, payloadBase64: "AA==" }),
    ).toBe(false);
  });

  it("rejects a replayed upgrade handshake (verified HMAC tuple is single-use)", async () => {
    const port = randomPort();
    server = await startServer({ port });

    const callId = "call-replay";
    const ts = Date.now();
    const headers = {
      "x-openclawteamsbridge-timestamp": String(ts),
      "x-openclawteamsbridge-signature": signHmac(SECRET, ts, callId),
    };

    // First connection with the signed tuple succeeds...
    const ws1 = new WebSocket(`ws://127.0.0.1:${port}${PATH}/${callId}`, { headers });
    await new Promise<void>((resolve, reject) => {
      ws1.once("open", () => resolve());
      ws1.once("error", reject);
    });
    ws1.close();
    await waitFor(() => server?.sessionCount === 0);

    // ...replaying the SAME tuple (a captured handshake) is rejected even though the
    // timestamp is still inside the HMAC window and no live session holds the callId.
    const ws2 = new WebSocket(`ws://127.0.0.1:${port}${PATH}/${callId}`, { headers });
    const outcome = await new Promise<"unexpected-response" | "error" | "open">((resolve) => {
      ws2.once("unexpected-response", (_req, res) => {
        expect(res.statusCode).toBe(401);
        resolve("unexpected-response");
      });
      ws2.once("error", () => resolve("error"));
      ws2.once("open", () => resolve("open"));
    });
    expect(outcome).not.toBe("open");

    // A fresh timestamp (legitimate reconnect) is still accepted.
    const ts2 = ts + 1;
    const ws3 = new WebSocket(`ws://127.0.0.1:${port}${PATH}/${callId}`, {
      headers: {
        "x-openclawteamsbridge-timestamp": String(ts2),
        "x-openclawteamsbridge-signature": signHmac(SECRET, ts2, callId),
      },
    });
    await new Promise<void>((resolve, reject) => {
      ws3.once("open", () => resolve());
      ws3.once("error", reject);
    });
    ws3.close();
  });

  it("still rejects a replay at exactly ts + hmacWindowMs (prune must not race the timestamp check)", async () => {
    // At now === ts + windowMs the timestamp check still accepts the handshake
    // (Math.abs(now - ts) > windowMs is false), so the replay record — whose expiry
    // is exactly ts + windowMs — must survive the prune at that same instant. A
    // `<=` prune would delete it one message too early and let a captured handshake
    // replay through. Only Date is faked; sockets and timers stay real.
    const windowMs = 5000;
    const t0 = Date.now();
    vi.useFakeTimers({ toFake: ["Date"] });
    try {
      vi.setSystemTime(t0);
      const port = randomPort();
      server = await startServer({ port, hmacWindowMs: windowMs });

      const callId = "call-replay-boundary";
      const headers = {
        "x-openclawteamsbridge-timestamp": String(t0),
        "x-openclawteamsbridge-signature": signHmac(SECRET, t0, callId),
      };

      const ws1 = new WebSocket(`ws://127.0.0.1:${port}${PATH}/${callId}`, { headers });
      await new Promise<void>((resolve, reject) => {
        ws1.once("open", () => resolve());
        ws1.once("error", reject);
      });
      ws1.close();
      await waitFor(() => server?.sessionCount === 0);

      // Jump to the exact edge of the HMAC window and replay the SAME tuple.
      vi.setSystemTime(t0 + windowMs);
      const ws2 = new WebSocket(`ws://127.0.0.1:${port}${PATH}/${callId}`, { headers });
      const outcome = await new Promise<"unexpected-response" | "error" | "open">((resolve) => {
        ws2.once("unexpected-response", (_req, res) => {
          expect(res.statusCode).toBe(401);
          resolve("unexpected-response");
        });
        ws2.once("error", () => resolve("error"));
        ws2.once("open", () => resolve("open"));
      });
      expect(outcome).not.toBe("open");
    } finally {
      vi.useRealTimers();
    }
  });

  it("rejects upgrade with a bad HMAC signature", async () => {
    const port = randomPort();
    server = await startServer({ port });

    const callId = "call-bad-sig";
    const ws = new WebSocket(`ws://127.0.0.1:${port}${PATH}/${callId}`, {
      headers: {
        "x-openclawteamsbridge-timestamp": String(Date.now()),
        "x-openclawteamsbridge-signature": "deadbeef",
      },
    });

    const outcome = await new Promise<"unexpected-response" | "error" | "open">((resolve) => {
      ws.once("unexpected-response", (_req, res) => {
        expect(res.statusCode).toBe(401);
        resolve("unexpected-response");
      });
      ws.once("error", () => resolve("error"));
      ws.once("open", () => resolve("open"));
    });

    expect(outcome).not.toBe("open");
    expect(server.sessionCount).toBe(0);
  });

  it("rejects upgrade when timestamp is far outside the HMAC window", async () => {
    const port = randomPort();
    server = await startServer({ port });

    const callId = "call-stale-ts";
    const staleTs = Date.now() - 5 * 60_000; // 5 minutes old
    const sig = signHmac(SECRET, staleTs, callId);

    const ws = new WebSocket(`ws://127.0.0.1:${port}${PATH}/${callId}`, {
      headers: {
        "x-openclawteamsbridge-timestamp": String(staleTs),
        "x-openclawteamsbridge-signature": sig,
      },
    });

    const outcome = await new Promise<"unexpected-response" | "error" | "open">((resolve) => {
      ws.once("unexpected-response", () => resolve("unexpected-response"));
      ws.once("error", () => resolve("error"));
      ws.once("open", () => resolve("open"));
    });

    expect(outcome).not.toBe("open");
  });

  it("rejects upgrade missing the callId in the path", async () => {
    const port = randomPort();
    server = await startServer({ port });

    const ts = Date.now();
    const sig = signHmac(SECRET, ts, "");
    const ws = new WebSocket(`ws://127.0.0.1:${port}${PATH}`, {
      headers: {
        "x-openclawteamsbridge-timestamp": String(ts),
        "x-openclawteamsbridge-signature": sig,
      },
    });

    const outcome = await new Promise<"unexpected-response" | "error" | "open">((resolve) => {
      ws.once("unexpected-response", () => resolve("unexpected-response"));
      ws.once("error", () => resolve("error"));
      ws.once("open", () => resolve("open"));
    });

    expect(outcome).not.toBe("open");
  });

  it("404s a path that only prefix-matches (exact segment, not startsWith)", async () => {
    const port = randomPort();
    server = await startServer({ port });

    // "/voice/msteams/streamX/..." must be a 404 (wrong endpoint), not fall through
    // to the HMAC check and read as a confusing 401.
    const callId = "call-prefix";
    const ts = Date.now();
    const ws = new WebSocket(`ws://127.0.0.1:${port}${PATH}X/${callId}`, {
      headers: {
        "x-standin-timestamp": String(ts),
        "x-standin-signature": signHmac(SECRET, ts, callId),
      },
    });
    const status = await new Promise<number>((resolve, reject) => {
      ws.once("unexpected-response", (_req, res) => resolve(res.statusCode ?? 0));
      ws.once("open", () => reject(new Error("connection should have been rejected")));
      ws.once("error", () => resolve(0)); // some ws versions surface only 'error'
    });
    if (status !== 0) {
      expect(status).toBe(404);
    }
  });

  it("survives a malformed absolute-form request-target (new URL would throw)", async () => {
    const port = randomPort();
    server = await startServer({ port });

    // A scanner's request line like "GET http://[ HTTP/1.1" reaches the upgrade
    // handler with request.url = "http://[", which new URL() throws on. The server
    // must answer 400 and keep running, not crash the whole gateway process.
    const net = await import("node:net");
    const reply = await new Promise<string>((resolve) => {
      const sock = net.connect(port, "127.0.0.1", () => {
        sock.write(
          "GET http://[ HTTP/1.1\r\n" +
            "Host: 127.0.0.1\r\n" +
            "Connection: Upgrade\r\n" +
            "Upgrade: websocket\r\n" +
            "Sec-WebSocket-Key: dGhlIHNhbXBsZSBub25jZQ==\r\n" +
            "Sec-WebSocket-Version: 13\r\n\r\n",
        );
      });
      let buf = "";
      sock.on("data", (d) => {
        buf += d.toString();
      });
      sock.on("close", () => resolve(buf));
      sock.on("error", () => resolve(buf));
      setTimeout(() => {
        sock.destroy();
        resolve(buf);
      }, 1500);
    });
    expect(reply).toContain("400");

    // The server must still accept a valid handshake afterwards.
    const callId = "call-after-malformed";
    const ws = openAuthedStandIn(port, callId);
    await new Promise<void>((resolve, reject) => {
      ws.once("open", () => resolve());
      ws.once("error", reject);
      ws.once("unexpected-response", () => reject(new Error("rejected")));
    });
    ws.close();
  });

  it("survives a peer that starts an upgrade and immediately destroys the socket", async () => {
    const port = randomPort();
    server = await startServer({ port });

    // Scanner behavior: send the upgrade request then slam the connection shut. The
    // reject write then races a dead socket; without an 'error' listener that is an
    // unhandled event that would take the process down.
    const net = await import("node:net");
    await new Promise<void>((resolve) => {
      const sock = net.connect(port, "127.0.0.1", () => {
        sock.write(
          `GET ${PATH}/call-rst HTTP/1.1\r\n` +
            "Host: 127.0.0.1\r\n" +
            "Connection: Upgrade\r\n" +
            "Upgrade: websocket\r\n" +
            "Sec-WebSocket-Key: dGhlIHNhbXBsZSBub25jZQ==\r\n" +
            "Sec-WebSocket-Version: 13\r\n\r\n",
          () => {
            sock.destroy(); // no HMAC headers -> server writes a 401 into a dead socket
            resolve();
          },
        );
      });
      sock.on("error", () => resolve());
    });
    // Give the reject write a beat to race the destroyed socket, then prove liveness.
    await new Promise<void>((resolve) => {
      setTimeout(resolve, 100);
    });

    const callId = "call-after-rst";
    const ws = openAuthedStandIn(port, callId);
    await new Promise<void>((resolve, reject) => {
      ws.once("open", () => resolve());
      ws.once("error", reject);
      ws.once("unexpected-response", () => reject(new Error("rejected")));
    });
    ws.close();
  });

  it("decodes audio.frame and emits via onAudioFrame", async () => {
    const port = randomPort();
    const received: Array<{ callId: string; seq: number; payload: Buffer }> = [];
    server = await startServer({
      port,
      onAudioFrame: (info) => {
        received.push({ callId: info.callId, seq: info.seq, payload: info.payload });
      },
    });

    const callId = "call-audio";
    const ts = Date.now();
    const sig = signHmac(SECRET, ts, callId);
    const ws = new WebSocket(`ws://127.0.0.1:${port}${PATH}/${callId}`, {
      headers: {
        "x-openclawteamsbridge-timestamp": String(ts),
        "x-openclawteamsbridge-signature": sig,
      },
    });
    await new Promise<void>((resolve, reject) => {
      ws.once("open", () => resolve());
      ws.once("error", reject);
    });

    const rawAudio = Buffer.from([0x01, 0x02, 0x03, 0x04]);
    ws.send(
      JSON.stringify({
        type: "audio.frame",
        seq: 42,
        timestampMs: Date.now(),
        payloadBase64: rawAudio.toString("base64"),
      }),
    );

    await waitFor(() => received.length > 0);
    expect(received[0]?.callId).toBe(callId);
    expect(received[0]?.seq).toBe(42);
    expect(received[0]?.payload.equals(rawAudio)).toBe(true);

    ws.close();
  });

  it("decodes assistant.say and emits via onAssistantSay (H4)", async () => {
    const port = randomPort();
    const said: Array<{ callId: string; text: string }> = [];
    server = await startServer({
      port,
      onAssistantSay: (info) => {
        said.push({ callId: info.callId, text: info.text });
      },
    });

    const callId = "call-say";
    const ts = Date.now();
    const sig = signHmac(SECRET, ts, callId);
    const ws = new WebSocket(`ws://127.0.0.1:${port}${PATH}/${callId}`, {
      headers: {
        "x-openclawteamsbridge-timestamp": String(ts),
        "x-openclawteamsbridge-signature": sig,
      },
    });
    await new Promise<void>((resolve, reject) => {
      ws.once("open", () => resolve());
      ws.once("error", reject);
    });

    ws.send(JSON.stringify({ type: "assistant.say", text: "Goodbye for now." }));

    await waitFor(() => said.length > 0);
    expect(said[0]?.callId).toBe(callId);
    expect(said[0]?.text).toBe("Goodbye for now.");

    ws.close();
  });

  it("session.end triggers onSessionEnd and closes the socket", async () => {
    const port = randomPort();
    let endInfo: { callId: string; reason: string } | undefined;
    server = await startServer({
      port,
      onSessionEnd: (info) => {
        endInfo = info;
      },
    });

    const callId = "call-end";
    const ts = Date.now();
    const sig = signHmac(SECRET, ts, callId);
    const ws = new WebSocket(`ws://127.0.0.1:${port}${PATH}/${callId}`, {
      headers: {
        "x-openclawteamsbridge-timestamp": String(ts),
        "x-openclawteamsbridge-signature": sig,
      },
    });
    await new Promise<void>((resolve, reject) => {
      ws.once("open", () => resolve());
      ws.once("error", reject);
    });

    ws.send(JSON.stringify({ type: "session.end", reason: "call-ended" }));

    await waitFor(() => endInfo !== undefined);
    expect(endInfo?.callId).toBe(callId);
    expect(endInfo?.reason).toBe("call-ended");
  });

  it("fires onSessionEnd when a started session's socket closes abruptly", async () => {
    const port = randomPort();
    const ends: Array<{ callId: string; reason: string }> = [];
    let started = false;
    server = await startServer({
      port,
      onSessionStart: () => {
        started = true;
      },
      onSessionEnd: (info) => ends.push(info),
    });

    const callId = "call-drop";
    const ws = openAuthed(port, callId);
    await new Promise<void>((resolve, reject) => {
      ws.once("open", () => resolve());
      ws.once("error", reject);
    });
    ws.send(
      JSON.stringify({ type: "session.start", callId, threadId: "t", caller: { aadId: "a" } }),
    );
    await waitFor(() => started);

    ws.close(); // abrupt close — no session.end frame

    await waitFor(() => ends.length > 0);
    expect(ends).toEqual([{ callId, reason: "socket-closed" }]);
  });

  it("does not double-fire onSessionEnd when the socket closes after session.end", async () => {
    const port = randomPort();
    const ends: Array<{ callId: string; reason: string }> = [];
    server = await startServer({ port, onSessionEnd: (info) => ends.push(info) });

    const callId = "call-end-once";
    const ws = openAuthed(port, callId);
    await new Promise<void>((resolve, reject) => {
      ws.once("open", () => resolve());
      ws.once("error", reject);
    });
    ws.send(
      JSON.stringify({ type: "session.start", callId, threadId: "t", caller: { aadId: "a" } }),
    );
    ws.send(JSON.stringify({ type: "session.end", reason: "call-ended" }));

    await waitFor(() => ends.length > 0);
    // The server closes the socket after session.end; let the close event run.
    await new Promise<void>((resolve) => {
      setTimeout(resolve, 100);
    });
    expect(ends).toEqual([{ callId, reason: "call-ended" }]);
  });

  it("does not fire onSessionEnd when the socket closes before session.start", async () => {
    const port = randomPort();
    const ends: Array<{ callId: string; reason: string }> = [];
    server = await startServer({ port, onSessionEnd: (info) => ends.push(info) });

    const ws = openAuthed(port, "call-prestart");
    await new Promise<void>((resolve, reject) => {
      ws.once("open", () => resolve());
      ws.once("error", reject);
    });
    ws.close(); // close before any session.start

    await new Promise<void>((resolve) => {
      setTimeout(resolve, 100);
    });
    expect(ends).toHaveLength(0);
  });

  it("normalizes blank caller ids to null in session.start (B11)", async () => {
    // An empty-string aadId would survive downstream `aadId ?? fallback` checks and collapse all
    // such callers into ONE session key (cross-caller memory bleed) / one delivery target.
    const port = randomPort();
    let session: MsteamsSession | undefined;
    server = await startServer({
      port,
      onSessionStart: (s) => {
        session = s;
      },
    });

    const callId = "call-blank-aad";
    const ws = openAuthed(port, callId);
    await new Promise<void>((resolve, reject) => {
      ws.once("open", () => resolve());
      ws.once("error", reject);
    });
    ws.send(
      JSON.stringify({
        type: "session.start",
        callId,
        threadId: "t",
        caller: { aadId: "", displayName: "  ", tenantId: "tenant-1" },
      }),
    );
    await waitFor(() => session !== undefined);

    expect(session?.caller.aadId).toBeNull();
    expect(session?.caller.displayName).toBeNull();
    expect(session?.caller.tenantId).toBe("tenant-1");

    ws.close();
  });

  it("fires onSessionEnd exactly once when the host closes a started session (session.close)", async () => {
    // A host-initiated close (e.g. realtime connect failure calling session.close) destroys the
    // connection meta before the ws close event runs, so the close handler alone would never
    // deliver onSessionEnd — leaking host call state. closeSession must deliver it itself. (B1)
    const port = randomPort();
    const ends: Array<{ callId: string; reason: string }> = [];
    let session: MsteamsSession | undefined;
    server = await startServer({
      port,
      onSessionStart: (s) => {
        session = s;
      },
      onSessionEnd: (info) => ends.push(info),
    });

    const callId = "call-host-close";
    const ws = openAuthed(port, callId);
    await new Promise<void>((resolve, reject) => {
      ws.once("open", () => resolve());
      ws.once("error", reject);
    });
    ws.send(
      JSON.stringify({ type: "session.start", callId, threadId: "t", caller: { aadId: "a" } }),
    );
    await waitFor(() => session !== undefined);

    session?.close("realtime-unavailable");

    await waitFor(() => ends.length > 0);
    // Let the ws close event run too — it must not double-deliver.
    await new Promise<void>((resolve) => {
      setTimeout(resolve, 100);
    });
    expect(ends).toEqual([{ callId, reason: "realtime-unavailable" }]);
  });

  it("drops the connection when an inbound frame exceeds the payload cap", async () => {
    const port = randomPort();
    let frames = 0;
    server = await startServer({
      port,
      onAudioFrame: () => {
        frames += 1;
      },
    });

    const callId = "call-oversize";
    const ts = Date.now();
    const sig = signHmac(SECRET, ts, callId);
    const ws = new WebSocket(`ws://127.0.0.1:${port}${PATH}/${callId}`, {
      headers: {
        "x-openclawteamsbridge-timestamp": String(ts),
        "x-openclawteamsbridge-signature": sig,
      },
    });
    await new Promise<void>((resolve, reject) => {
      ws.once("open", () => resolve());
      ws.once("error", reject);
    });

    // ~3 MB base64 payload — over the 2 MB inbound cap (sized for video.frame). ws closes
    // oversized frames with code 1009 (message too big) before they reach JSON parsing.
    const closeCode = await new Promise<number>((resolve) => {
      ws.once("close", (code) => resolve(code));
      ws.send(
        JSON.stringify({
          type: "audio.frame",
          seq: 0,
          timestampMs: Date.now(),
          payloadBase64: "A".repeat(3 * 1024 * 1024),
        }),
      );
    });

    expect(closeCode).toBe(1009);
    expect(frames).toBe(0);
  });

  it("parses video.frame and emits via onVideoFrame", async () => {
    const port = randomPort();
    const received: Array<{ source: string; width: number; height: number; dataBase64: string }> =
      [];
    server = await startServer({ port, onVideoFrame: (info) => received.push(info) });

    const callId = "call-video";
    const ws = openAuthed(port, callId);
    await new Promise<void>((resolve, reject) => {
      ws.once("open", () => resolve());
      ws.once("error", reject);
    });
    ws.send(
      JSON.stringify({
        type: "video.frame",
        source: "screenshare",
        ts: 1719,
        width: 1280,
        height: 720,
        mime: "image/jpeg",
        dataBase64: "AQID",
      }),
    );

    await waitFor(() => received.length > 0);
    expect(received[0]).toMatchObject({
      source: "screenshare",
      width: 1280,
      height: 720,
      dataBase64: "AQID",
    });
  });

  it("rejects connections beyond maxConnections", async () => {
    const port = randomPort();
    server = await startServer({ port, maxConnections: 1 });

    const ws1 = openAuthed(port, "call-cap-1");
    await new Promise<void>((resolve, reject) => {
      ws1.once("open", () => resolve());
      ws1.once("error", reject);
    });
    expect(server.sessionCount).toBe(1);

    const ws2 = openAuthed(port, "call-cap-2");
    const outcome = await new Promise<"unexpected-response" | "error" | "open">((resolve) => {
      ws2.once("unexpected-response", (_req, res) => {
        expect(res.statusCode).toBe(503);
        resolve("unexpected-response");
      });
      ws2.once("error", () => resolve("error"));
      ws2.once("open", () => resolve("open"));
    });
    expect(outcome).not.toBe("open");
    expect(server.sessionCount).toBe(1);
    ws1.close();
  });

  it("rejects connections beyond maxConnectionsPerIp", async () => {
    const port = randomPort();
    server = await startServer({ port, maxConnectionsPerIp: 1 });

    const ws1 = openAuthed(port, "call-ip-1");
    await new Promise<void>((resolve, reject) => {
      ws1.once("open", () => resolve());
      ws1.once("error", reject);
    });

    const ws2 = openAuthed(port, "call-ip-2");
    const outcome = await new Promise<"unexpected-response" | "error" | "open">((resolve) => {
      ws2.once("unexpected-response", (_req, res) => {
        expect(res.statusCode).toBe(503);
        resolve("unexpected-response");
      });
      ws2.once("error", () => resolve("error"));
      ws2.once("open", () => resolve("open"));
    });
    expect(outcome).not.toBe("open");
    ws1.close();
  });

  it("rejects session.start whose callId does not match the authenticated path", async () => {
    const port = randomPort();
    let started = false;
    server = await startServer({
      port,
      onSessionStart: () => {
        started = true;
      },
    });

    const ws = openAuthed(port, "call-auth");
    await new Promise<void>((resolve, reject) => {
      ws.once("open", () => resolve());
      ws.once("error", reject);
    });

    const closeCode = await new Promise<number>((resolve) => {
      ws.once("close", (code) => resolve(code));
      ws.send(
        JSON.stringify({
          type: "session.start",
          callId: "call-spoofed",
          threadId: "thread-1",
          caller: { aadId: "aad-1" },
        }),
      );
    });

    expect(started).toBe(false);
    expect(closeCode).toBeGreaterThan(0);
    expect(server.sessionCount).toBe(0);
  });

  it("closes a connection that never sends session.start", async () => {
    const port = randomPort();
    server = await startServer({ port, preStartTimeoutMs: 120 });

    const ws = openAuthed(port, "call-idle");
    await new Promise<void>((resolve, reject) => {
      ws.once("open", () => resolve());
      ws.once("error", reject);
    });
    expect(server.sessionCount).toBe(1);

    await new Promise<void>((resolve) => {
      ws.once("close", () => resolve());
    });
    expect(server.sessionCount).toBe(0);
  });

  it("surfaces recording status from session.start and recording.status messages", async () => {
    const port = randomPort();
    let startStatus: string | undefined;
    const statuses: string[] = [];
    server = await startServer({
      port,
      onSessionStart: (s) => {
        startStatus = s.recordingStatus;
      },
      onRecordingStatus: (info) => {
        statuses.push(info.status);
      },
    });

    const callId = "call-rec";
    const ws = openAuthed(port, callId);
    await new Promise<void>((resolve, reject) => {
      ws.once("open", () => resolve());
      ws.once("error", reject);
    });

    ws.send(
      JSON.stringify({
        type: "session.start",
        callId,
        threadId: "thread-rec",
        caller: { aadId: "aad-1" },
        recordingStatus: "inactive",
      }),
    );
    await waitFor(() => startStatus !== undefined);
    expect(startStatus).toBe("inactive");

    ws.send(JSON.stringify({ type: "recording.status", status: "active" }));
    await waitFor(() => statuses.length > 0);
    expect(statuses).toEqual(["active"]);

    ws.close();
  });

  it("authenticates a signature that is upper-case hex (HMAC normalization, matches Hermes)", async () => {
    // Hermes verify_upgrade does .strip().lower() on the signature; a worker that hex-encodes upper-
    // case must still authenticate. Previously the raw header was compared and this failed.
    const port = randomPort();
    let started = false;
    server = await startServer({ port, onSessionStart: () => (started = true) });

    const callId = "call-upper-sig";
    const ts = Date.now();
    const ws = new WebSocket(`ws://127.0.0.1:${port}${PATH}/${callId}`, {
      headers: {
        "x-openclawteamsbridge-timestamp": String(ts),
        "x-openclawteamsbridge-signature": signHmac(SECRET, ts, callId).toUpperCase(),
      },
    });
    const outcome = await new Promise<"open" | "error" | "unexpected-response">((resolve) => {
      ws.once("open", () => resolve("open"));
      ws.once("error", () => resolve("error"));
      ws.once("unexpected-response", () => resolve("unexpected-response"));
    });
    expect(outcome).toBe("open"); // upper-case hex now authenticates
    ws.close();
  });

  it("drops an outbound frame when the send buffer is backed up (egress backpressure)", () => {
    // sendTo is fire-and-forget; a stalled worker must not let ws.bufferedAmount grow unbounded.
    const media = new MsteamsMediaStream({ port: 0, path: PATH, sharedSecret: SECRET });
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const inner = media as any;
    const fakeWs = { readyState: WebSocket.OPEN, bufferedAmount: 0, send: vi.fn() };
    inner.sessions.set("c1", fakeWs);

    expect(inner.sendTo("c1", { type: "audio.frame" })).toBe(true);
    expect(fakeWs.send).toHaveBeenCalledTimes(1);

    fakeWs.bufferedAmount = 2 * 1024 * 1024; // over the 1 MB cap
    expect(inner.sendTo("c1", { type: "audio.frame" })).toBe(false); // dropped
    expect(fakeWs.send).toHaveBeenCalledTimes(1); // not sent again
  });

  it("accepts only valid DTMF digits (0-9, *, #) and drops anything else", async () => {
    const port = randomPort();
    const digits: string[] = [];
    server = await startServer({
      port,
      onDtmf: (info) => {
        digits.push(info.digit);
      },
    });

    const callId = "call-dtmf";
    const ws = openAuthed(port, callId);
    await new Promise<void>((resolve, reject) => {
      ws.once("open", () => resolve());
      ws.once("error", reject);
    });
    ws.send(
      JSON.stringify({
        type: "session.start",
        callId,
        threadId: "thread-dtmf",
        caller: { aadId: "aad-1" },
      }),
    );

    // Invalid digits first: wrong character, multi-char, empty. None may reach onDtmf.
    for (const digit of ["a", "12", "", "+"]) {
      ws.send(JSON.stringify({ type: "dtmf", digit }));
    }
    // Then the full valid alphabet.
    const valid = ["0", "5", "9", "*", "#"];
    for (const digit of valid) {
      ws.send(JSON.stringify({ type: "dtmf", digit }));
    }

    await waitFor(() => digits.length >= valid.length);
    expect(digits).toEqual(valid);

    ws.close();
  });
});
