import crypto from "node:crypto";
import { afterEach, describe, expect, it, vi } from "vitest";
import { WebSocket } from "ws";
import { MsteamsMediaStream, type MsteamsSession } from "./msteams-media-stream.js";

const SECRET = "test-shared-secret";
const PATH = "/msteams/calling";

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

/** Open an authenticated WS connection for a callId, using the only accepted header
 * pair: X-StandIn-Timestamp / X-StandIn-Signature. */
function openAuthed(port: number, callId: string): WebSocket {
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

  it("accepts a connection with valid HMAC (X-StandIn-*) + parses session.start", async () => {
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
        "x-standin-timestamp": String(ts),
        "x-standin-signature": sig,
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

  it("accepts a session.start with an EMPTY threadId (a 1:1 call has no meeting thread)", async () => {
    // Regression: the schema carried a min-length check on threadId, so the moment the media bridge
    // stopped smuggling the call id into that field (a 1:1 call genuinely has no thread) every 1:1
    // session.start was rejected as invalid, no session ever started, and the caller heard the
    // worker's local echo. Empty is the honest wire value and MUST be accepted.
    const port = randomPort();
    let receivedSession: MsteamsSession | undefined;
    server = await startServer({
      port,
      onSessionStart: (s) => {
        receivedSession = s;
      },
    });

    const callId = "call-one-to-one";
    const ws = openAuthed(port, callId);
    await new Promise<void>((resolve, reject) => {
      ws.once("open", () => resolve());
      ws.once("error", reject);
    });

    ws.send(
      JSON.stringify({
        type: "session.start",
        callId,
        threadId: "",
        caller: { aadId: "aad-1", displayName: "Alice", tenantId: "tenant-1" },
      }),
    );

    await waitFor(() => receivedSession !== undefined);
    expect(receivedSession?.callId).toBe(callId);
    expect(receivedSession?.threadId).toBe("");
    expect(server.sessionCount).toBe(1);

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
        "x-standin-timestamp": String(ts),
        "x-standin-signature": signHmac(SECRET, ts, callId),
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
      "x-standin-timestamp": String(ts),
      "x-standin-signature": signHmac(SECRET, ts, callId),
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
        "x-standin-timestamp": String(ts2),
        "x-standin-signature": signHmac(SECRET, ts2, callId),
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
        "x-standin-timestamp": String(t0),
        "x-standin-signature": signHmac(SECRET, t0, callId),
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
        "x-standin-timestamp": String(Date.now()),
        "x-standin-signature": "deadbeef",
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
        "x-standin-timestamp": String(staleTs),
        "x-standin-signature": sig,
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
        "x-standin-timestamp": String(ts),
        "x-standin-signature": sig,
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

    // "/msteams/callingX/..." must be a 404 (wrong endpoint), not fall through
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
    const ws = openAuthed(port, callId);
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

    // A peer that starts an upgrade then drops the connection mid-handshake: the
    // server must absorb it quietly and keep serving (a liveness guard; the raw
    // socket also carries its own error handler so a stray error stays contained).
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
    const ws = openAuthed(port, callId);
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
        "x-standin-timestamp": String(ts),
        "x-standin-signature": sig,
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
        "x-standin-timestamp": String(ts),
        "x-standin-signature": sig,
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
        "x-standin-timestamp": String(ts),
        "x-standin-signature": sig,
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
        "x-standin-timestamp": String(ts),
        "x-standin-signature": sig,
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

  // THE RECORDING-STATUS RACE. recording.status is free to arrive before session.start, and
  // session.start carries only a setup-time SNAPSHOT of the field. Letting the snapshot win closed
  // the Media Access gate for the whole call — audio, vision, consults and DTMF all refused silently,
  // with no way back, because nothing re-sends a status that never changed.
  it("an explicit recording.status before session.start is not overwritten by the session.start seed", async () => {
    const port = randomPort();
    let seeded: string | undefined;
    let seenAtStart = false;
    // Order matters as much as the value: the replay has to land AFTER the call handle exists.
    const events: string[] = [];
    server = await startServer({
      port,
      onSessionStart: (s) => {
        seeded = s.recordingStatus;
        seenAtStart = true;
        events.push("session.start");
      },
      onRecordingStatus: (info) => {
        events.push(`recording.status=${info.status}:started=${seenAtStart}`);
      },
    });

    const callId = "call-race";
    const ws = openAuthed(port, callId);
    await new Promise<void>((resolve, reject) => {
      ws.once("open", () => resolve());
      ws.once("error", reject);
    });

    // The live value arrives FIRST.
    ws.send(JSON.stringify({ type: "recording.status", status: "active" }));
    await waitFor(() => events.length > 0);

    // ...and session.start then reports the stale snapshot.
    ws.send(
      JSON.stringify({
        type: "session.start",
        callId,
        threadId: "thread-race",
        caller: { aadId: "aad-race" },
        recordingStatus: "inactive",
      }),
    );
    await waitFor(() => seeded !== undefined);

    // The seed handed to the call is the LIVE value, not the snapshot.
    expect(seeded).toBe("active");
    // And the pre-start status is replayed once the call handle exists, so the per-call work that
    // only runs on a status CHANGE (deferred outbound greeting, ambient-vision flush) still runs.
    expect(events).toEqual([
      "recording.status=active:started=false", // the original delivery, dropped on the floor downstream
      "session.start",
      "recording.status=active:started=true", // the replay the call handle can actually receive
    ]);

    ws.close();
  });

  // The latch is a latch, not a one-way switch: a worker that stopped recording before session.start
  // must not have the gate opened by a stale "active" snapshot either.
  it("a pre-start recording.status of inactive also wins over an active session.start seed", async () => {
    const port = randomPort();
    let seeded: string | undefined = "unset";
    server = await startServer({
      port,
      onSessionStart: (s) => {
        seeded = s.recordingStatus;
      },
    });

    const callId = "call-race-off";
    const ws = openAuthed(port, callId);
    await new Promise<void>((resolve, reject) => {
      ws.once("open", () => resolve());
      ws.once("error", reject);
    });

    ws.send(JSON.stringify({ type: "recording.status", status: "inactive" }));
    ws.send(
      JSON.stringify({
        type: "session.start",
        callId,
        threadId: "thread-race-off",
        caller: { aadId: "aad-race-off" },
        recordingStatus: "active",
      }),
    );
    await waitFor(() => seeded !== "unset");
    expect(seeded).toBe("inactive");

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
        "x-standin-timestamp": String(ts),
        "x-standin-signature": signHmac(SECRET, ts, callId).toUpperCase(),
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

describe("worker call-outcome route", () => {
  // Hermes has answered POST {base}/outcome/{callId} for a while; OpenClaw did not, so the worker's
  // signal 404'd and this plugin sat on its own answer-timeout - a DECLINED call and an unanswered one
  // were indistinguishable here and distinct on the other agent.
  const OUTCOME_PATH = (callId: string) => `${PATH}/outcome/${callId}`;

  function sign(callId: string, ts: number, secret = SECRET): string {
    return crypto.createHmac("sha256", secret).update(`${ts}.${callId}`).digest("hex");
  }

  async function post(port: number, callId: string, body: unknown, headers: Record<string, string>) {
    const res = await fetch(`http://127.0.0.1:${port}${OUTCOME_PATH(callId)}`, {
      method: "POST",
      headers: { "content-type": "application/json", ...headers },
      body: JSON.stringify(body),
    });
    return res.status;
  }

  it("delivers a signed outcome and refuses an unsigned or mis-signed one", async () => {
    const seen: { callId: string; outcome: string }[] = [];
    const port = randomPort();
    const server = new MsteamsMediaStream({
      port,
      path: PATH,
      sharedSecret: SECRET,
      onCallOutcome: (i) => seen.push(i),
    });
    await server.start();
    try {
      const ts = Date.now();
      expect(await post(port, "call-1", { outcome: "declined" }, {
        "x-standin-timestamp": String(ts),
        "x-standin-signature": sign("call-1", ts),
      })).toBe(200);
      expect(seen).toEqual([{ callId: "call-1", outcome: "declined" }]);

      expect(await post(port, "call-2", { outcome: "busy" }, {})).toBe(401);
      expect(await post(port, "call-3", { outcome: "busy" }, {
        "x-standin-timestamp": String(ts),
        "x-standin-signature": sign("call-3", ts, "wrong-secret"),
      })).toBe(401);
      // The signature covers the callId, so one captured for another call must not be reusable here.
      expect(await post(port, "call-4", { outcome: "busy" }, {
        "x-standin-timestamp": String(ts),
        "x-standin-signature": sign("call-1", ts),
      })).toBe(401);

      // v2 covers the BODY, which v1 does not - so a tampered outcome word must be refused, and a
      // correct v2 must still pass.
      const canonical = (body: string, callId: string) =>
        `POST\n${OUTCOME_PATH(callId)}\n${crypto.createHash("sha256").update(body).digest("hex")}`;
      const signV2 = (body: string, callId: string, ts: number) =>
        crypto.createHmac("sha256", SECRET).update(`${ts}.${canonical(body, callId)}`).digest("hex");

      const ts2 = Date.now();
      const good = JSON.stringify({ outcome: "busy" });
      const res = await fetch(`http://127.0.0.1:${port}${OUTCOME_PATH("call-5")}`, {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-standin-timestamp": String(ts2),
          "x-standin-signature": sign("call-5", ts2),
          "x-standin-signature-v2": signV2(good, "call-5", ts2),
        },
        body: good,
      });
      expect(res.status).toBe(200);
      expect(seen).toHaveLength(2);

      // Same v1 signature (it only covers the callId), body swapped: v1 alone cannot tell, v2 can.
      const tampered = await fetch(`http://127.0.0.1:${port}${OUTCOME_PATH("call-6")}`, {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-standin-timestamp": String(ts2),
          "x-standin-signature": sign("call-6", ts2),
          "x-standin-signature-v2": signV2(JSON.stringify({ outcome: "busy" }), "call-6", ts2),
        },
        body: JSON.stringify({ outcome: "failed" }),
      });
      expect(tampered.status).toBe(401);

      expect(seen).toHaveLength(2);
    } finally {
      await server.stop();
    }
  });
});
