/**
 * MsteamsMediaStream
 *
 * WebSocket server that accepts connections from the StandIn media bridge
 * bridge worker and relays Microsoft Teams call audio in both directions. One
 * connection per Teams call, keyed by callId in the URL path.
 *
 * Responsibilities:
 * - HTTP upgrade with HMAC-SHA256 verification of timestamp + callId, plus a
 *   replay window on the timestamp.
 * - Session lifecycle messages (session.start / session.end) parsed and emitted
 *   via callbacks for the host to wire into its session machinery.
 * - Inbound audio frames surfaced via `onAudioFrame` for the host to forward to
 *   the realtime-transcription provider.
 * - Outbound `send` and `close` exposed on the SessionStart callback so the host
 *   can push synthesized TTS audio and control messages back to the worker.
 */

import crypto from "node:crypto";
import http from "node:http";
import type { Duplex } from "node:stream";
import { safeEqualSecret } from "openclaw/plugin-sdk/security-runtime";
import { type RawData, WebSocket, WebSocketServer } from "ws";
import { type InboundMessage, InboundMessageSchema, type MsteamsRecordingStatus } from "./protocol.gen.js";

// The wire protocol (message schemas + shared constants) is generated from the
// the shared wire-protocol schema, mirrored in src/protocol.gen.ts. This file
// keeps only the transport: HMAC/replay/connection guards and message routing.
export { MSTEAMS_PCM_SAMPLE_RATE_HZ } from "./protocol.gen.js";
export type { MsteamsRecordingStatus } from "./protocol.gen.js";

export interface MsteamsLogger {
  info(message: string): void;
  warn(message: string): void;
  error(message: string): void;
  debug?(message: string): void;
}

export interface MsteamsSession {
  callId: string;
  threadId: string;
  caller: {
    aadId?: string | null;
    displayName?: string | null;
    tenantId?: string | null;
  };
  /** Teams recording status reported at answer time (if the worker set it). */
  recordingStatus?: MsteamsRecordingStatus;
  /** "inbound" (default) or "outbound" (the bot placed this call). */
  direction?: "inbound" | "outbound";
  /**
   * Push a JSON-serializable message to the worker. Returns false if the socket is closed/missing
   * (the message was dropped) so delivery-sensitive callers (audio frames) can observe the drop;
   * best-effort control frames can ignore the result.
   */
  send: (message: unknown) => boolean;
  /** Close the WebSocket gracefully with the given reason text. */
  close: (reason: string) => void;
}

export interface MsteamsMediaStreamConfig {
  port: number;
  /**
   * Interface address to bind the WebSocket server to. Defaults to the loopback
   * interface (127.0.0.1) so the bridge is not exposed on all interfaces.
   */
  bindAddress?: string;
  path: string;
  sharedSecret: string;
  /**
   * Reject upgrades whose timestamp is more than this many ms off from the
   * server clock. Mitigates replay. Default 60_000 (60 seconds).
   */
  hmacWindowMs?: number;
  /** Hard cap on total concurrent connections (pending + active). Default 64. */
  maxConnections?: number;
  /** Hard cap on concurrent connections per source IP. Default 8. */
  maxConnectionsPerIp?: number;
  /** Close a connection that has not sent session.start within this many ms. Default 10_000. */
  preStartTimeoutMs?: number;
  logger?: MsteamsLogger;
  onSessionStart?: (session: MsteamsSession) => void;
  /** Teams recording status changed mid-call (worker called Graph updateRecordingStatus). */
  onRecordingStatus?: (info: { callId: string; status: MsteamsRecordingStatus }) => void;
  onSessionEnd?: (info: { callId: string; reason: string }) => void;
  onAudioFrame?: (info: {
    callId: string;
    seq: number;
    timestampMs: number;
    payload: Buffer;
    /** Active speaker (unmixed-audio worker), for transcript attribution. */
    speakerName?: string;
  }) => void;
  /** A sampled inbound video frame (caller camera or screen-share) for the agent to "see". */
  onVideoFrame?: (info: {
    callId: string;
    source: "camera" | "screenshare";
    ts: number;
    width: number;
    height: number;
    mime: string;
    dataBase64: string;
    participantId?: string;
    participantName?: string;
  }) => void;
  /** Human participant count on the call changed (excludes the bot). count >= 2 ⇒ group/meeting. */
  onParticipants?: (info: { callId: string; count: number }) => void;
  /** A DTMF key the caller pressed ("0"-"9", "*", "#"). See #21. */
  onDtmf?: (info: { callId: string; digit: string }) => void;
  /**
   * H4: the worker asks the agent to speak `text` in its own realtime voice (e.g. a brief goodbye
   * right before a limit-cutoff teardown, instead of a bare mid-sentence hangup).
   */
  onAssistantSay?: (info: { callId: string; text: string }) => void;
}

const DEFAULT_HMAC_WINDOW_MS = 60_000;

/**
 * Bind to loopback by default. The Teams bridge worker typically runs on the
 * same host (or reaches OpenClaw over a private/VPN interface the operator names
 * explicitly via `bindAddress`); binding all interfaces would expose the audio
 * transport to untrusted networks.
 */
const DEFAULT_BIND_ADDRESS = "127.0.0.1";

/**
 * Hard cap on a single inbound WebSocket frame. Control + audio messages are tiny
 * (<1 KB), but a `video.frame` carries a base64 JPEG (the worker downscales and caps
 * the JPEG at ~1 MB, so base64 ≈ 1.4 MB). The cap is sized to admit one such frame
 * and still bound memory: the sender is the HMAC-authenticated worker, not an
 * arbitrary peer, so the looser bound only loosens the trusted path.
 */
const MAX_INBOUND_PAYLOAD_BYTES = 2 * 1024 * 1024;

/**
 * Connection guardrails, mirroring the Twilio media-stream path so a valid
 * shared secret (or a leaked/misbehaving worker) cannot open call sockets
 * unbounded and exhaust file descriptors or memory.
 */
const DEFAULT_MAX_CONNECTIONS = 64;
const DEFAULT_MAX_CONNECTIONS_PER_IP = 8;
const DEFAULT_PRE_START_TIMEOUT_MS = 10_000;
// Protocol-level WebSocket ping interval. A caller that dies UNCLEANLY (pod killed, network drop) never sends a
// TCP FIN, so `ws` fires no "close" event and the session — plus its maxConcurrentCalls slot — leaks until the
// far slower call reaper. Ping every client each interval; a client that missed the previous ping (isAlive still
// false) is terminated, which fires "close" → onSessionEnd → the concurrency slot frees. The .NET caller's
// ClientWebSocket auto-replies to protocol pings, so a healthy call always stays alive.
const HEARTBEAT_INTERVAL_MS = 30_000;

/** A ws-like socket for the heartbeat sweep (subset of `ws` used here; keeps the sweep unit-testable). */
export interface HeartbeatSocket {
  isAlive?: boolean;
  terminate(): void;
  ping(): void;
}

/** One heartbeat sweep over the connected clients: terminate any that missed the previous ping (isAlive still
 * false → dead caller), otherwise mark not-alive and ping (the pong handler flips it back). Exported for tests. */
export function heartbeatSweep(clients: Iterable<HeartbeatSocket>): void {
  for (const ws of clients) {
    if (ws.isAlive === false) {
      ws.terminate(); // fires "close" → onSessionEnd → the concurrency slot frees
      continue;
    }
    ws.isAlive = false;
    try {
      ws.ping();
    } catch {
      // best-effort; a socket erroring on ping is torn down on the next sweep
    }
  }
}

/**
 * Backpressure bound on a single call's OUTBOUND send buffer. Audio egress (sendTo) is
 * fire-and-forget; if the worker stalls, ws.bufferedAmount grows unbounded and leaks memory. When
 * the buffer is already this deep we drop the frame (return false) instead of queuing more — audio
 * is real-time, so a dropped stale frame beats ever-growing latency debt.
 */
const MAX_OUTBOUND_BUFFER_BYTES = 1 * 1024 * 1024;

/** Per-connection bookkeeping for caps + pre-start idle reaping. */
interface ConnectionMeta {
  ip: string;
  started: boolean;
  /** Set once onSessionEnd has fired, so socket close does not double-deliver it. */
  ended: boolean;
  preStartTimer: ReturnType<typeof setTimeout>;
}

export class MsteamsMediaStream {
  private readonly config: MsteamsMediaStreamConfig;
  private readonly hmacWindowMs: number;
  /** Verified upgrade tuples already used once (replay guard); value = expiry epoch ms. */
  private readonly seenUpgrades = new Map<string, number>();
  private readonly maxConnections: number;
  private readonly maxConnectionsPerIp: number;
  private readonly preStartTimeoutMs: number;
  private readonly sessions = new Map<string, WebSocket>();
  private readonly connectionMeta = new Map<string, ConnectionMeta>();
  private readonly connectionsByIp = new Map<string, number>();
  private server?: http.Server;
  private wss?: WebSocketServer;
  private heartbeat?: ReturnType<typeof setInterval>;

  constructor(config: MsteamsMediaStreamConfig) {
    this.config = config;
    this.hmacWindowMs = config.hmacWindowMs ?? DEFAULT_HMAC_WINDOW_MS;
    this.maxConnections = config.maxConnections ?? DEFAULT_MAX_CONNECTIONS;
    this.maxConnectionsPerIp = config.maxConnectionsPerIp ?? DEFAULT_MAX_CONNECTIONS_PER_IP;
    this.preStartTimeoutMs = config.preStartTimeoutMs ?? DEFAULT_PRE_START_TIMEOUT_MS;
  }

  async start(): Promise<void> {
    if (this.server) {
      throw new Error("MsteamsMediaStream is already started");
    }

    const server = http.createServer();
    const wss = new WebSocketServer({ noServer: true, maxPayload: MAX_INBOUND_PAYLOAD_BYTES });

    server.on("upgrade", (request, socket, head) => {
      // Give the raw socket an error handler for the window before the WebSocket
      // exists (a peer may connect then drop mid-handshake), so a stray socket
      // error stays contained rather than surfacing as an unhandled 'error' event.
      socket.on("error", () => {
        socket.destroy();
      });
      // Backstop: nothing thrown while vetting an (unauthenticated) upgrade may
      // escape into the event loop - a malformed request must never crash the host.
      try {
        this.handleUpgrade(request, socket, head, wss);
      } catch (error) {
        this.config.logger?.warn(
          `MsteamsMediaStream: rejecting upgrade — unparseable request (${String(error)})`,
        );
        this.rejectUpgrade(socket, 400, "Bad Request");
      }
    });

    await new Promise<void>((resolve, reject) => {
      const onError = (err: Error) => {
        server.off("listening", onListening);
        reject(err);
      };
      const onListening = () => {
        server.off("error", onError);
        resolve();
      };
      server.once("error", onError);
      server.once("listening", onListening);
      server.listen(this.config.port, this.config.bindAddress ?? DEFAULT_BIND_ADDRESS);
    });

    this.server = server;
    this.wss = wss;

    // Heartbeat: terminate dead sockets so their maxConcurrentCalls slot frees promptly (the canonical `ws`
    // pattern). Without it, a caller that dies uncleanly leaks a session until the call reaper's maxDuration.
    const heartbeat = setInterval(() => {
      heartbeatSweep(wss.clients as unknown as Iterable<HeartbeatSocket>);
    }, HEARTBEAT_INTERVAL_MS);
    (heartbeat as { unref?: () => void }).unref?.(); // don't keep the process alive just for the heartbeat
    this.heartbeat = heartbeat;

    this.config.logger?.info(
      `MsteamsMediaStream listening host=${this.config.bindAddress ?? DEFAULT_BIND_ADDRESS} port=${this.config.port} path=${this.config.path}`,
    );
  }

  async stop(): Promise<void> {
    if (!this.server) {
      return;
    }

    if (this.heartbeat) {
      clearInterval(this.heartbeat);
      this.heartbeat = undefined;
    }

    for (const ws of this.sessions.values()) {
      try {
        ws.close(1001, "shutdown");
      } catch {
        // best-effort close
      }
    }
    for (const meta of this.connectionMeta.values()) {
      clearTimeout(meta.preStartTimer);
    }
    this.connectionMeta.clear();
    this.connectionsByIp.clear();
    this.sessions.clear();

    this.wss?.close();
    await new Promise<void>((resolve) => {
      // server is non-null because we early-returned above; cast is for ts narrowing
      const s = this.server as http.Server;
      s.close(() => resolve());
    });

    this.server = undefined;
    this.wss = undefined;
    this.config.logger?.info("MsteamsMediaStream stopped");
  }

  /** Number of currently open sessions. Exposed for tests + telemetry. */
  get sessionCount(): number {
    return this.sessions.size;
  }

  private handleUpgrade(
    request: http.IncomingMessage,
    socket: Duplex,
    head: Buffer,
    wss: WebSocketServer,
  ): void {
    // request.url is attacker-controlled and may be an invalid absolute-form
    // request-target (e.g. "GET http://[ HTTP/1.1" from a scanner): new URL()
    // throws on those, so parse defensively instead of crashing the gateway.
    let url: URL;
    try {
      url = new URL(request.url ?? "", "http://localhost");
    } catch {
      this.rejectUpgrade(socket, 400, "Bad Request");
      return;
    }

    // Exact segment match, not a loose prefix: "/voice/msteams/streamX" must 404
    // here rather than fall through and confusingly 401 at the HMAC check.
    if (url.pathname !== this.config.path && !url.pathname.startsWith(this.config.path + "/")) {
      this.rejectUpgrade(socket, 404, "Not Found");
      return;
    }

    const callId = url.pathname.slice(this.config.path.length).replace(/^\//, "");
    if (!callId) {
      this.rejectUpgrade(socket, 400, "Bad Request (missing callId)");
      return;
    }

    // Prefer the X-StandIn-* names; fall back to the legacy X-OpenClawTeamsBridge-*
    // pair so pre-rename StandIn deployments keep connecting.
    const timestamp = request.headers["x-standin-timestamp"] ?? request.headers["x-openclawteamsbridge-timestamp"];
    const signature = request.headers["x-standin-signature"] ?? request.headers["x-openclawteamsbridge-signature"];
    if (typeof timestamp !== "string" || typeof signature !== "string") {
      this.config.logger?.warn(
        `MsteamsMediaStream: rejecting upgrade for ${callId} — missing HMAC headers`,
      );
      this.rejectUpgrade(socket, 401, "Unauthorized");
      return;
    }

    const ts = Number(timestamp);
    if (!Number.isFinite(ts) || Math.abs(Date.now() - ts) > this.hmacWindowMs) {
      this.config.logger?.warn(
        `MsteamsMediaStream: rejecting upgrade for ${callId} — timestamp out of window`,
      );
      this.rejectUpgrade(socket, 401, "Unauthorized");
      return;
    }

    // Normalize the incoming signature the same way Hermes's hmac_auth.verify_upgrade does
    // (.strip().lower()) so a worker that hex-encodes upper-case or pads whitespace still
    // authenticates. `expected` is already lower-case hex; the compare stays constant-time.
    const sig = signature.trim().toLowerCase();
    const expected = crypto
      .createHmac("sha256", this.config.sharedSecret)
      .update(`${ts}.${callId}`)
      .digest("hex");
    if (!safeEqualSecret(sig, expected)) {
      this.config.logger?.warn(
        `MsteamsMediaStream: rejecting upgrade for ${callId} — bad signature`,
      );
      this.rejectUpgrade(socket, 401, "Unauthorized");
      return;
    }

    // Replay guard: a verified (callId, timestamp, signature) tuple is single-use. A captured
    // handshake replayed within the HMAC window would otherwise open a ghost session once the live
    // call ends (the duplicate-callId check only protects while it's connected). A legitimate
    // reconnect always carries a fresh timestamp, so this rejects only true replays. Only verified
    // tuples are recorded (an attacker without the secret cannot grow the map); entries expire with
    // the timestamp window.
    const now = Date.now();
    for (const [key, expiry] of this.seenUpgrades) {
      // Strict <: at exactly ts + windowMs the timestamp check above still accepts the handshake
      // (Math.abs(now - ts) > windowMs is false), so the replay record must survive that instant.
      if (expiry < now) {
        this.seenUpgrades.delete(key);
      }
    }
    const replayKey = `${callId}.${ts}.${sig}`;
    if (this.seenUpgrades.has(replayKey)) {
      this.config.logger?.warn(
        `MsteamsMediaStream: rejecting upgrade for ${callId} — replayed handshake`,
      );
      this.rejectUpgrade(socket, 401, "Unauthorized");
      return;
    }
    // Expire the replay record when the TIMESTAMP stops being valid (ts + window), not at now +
    // window. A future-dated handshake (worker clock skew) is signature-valid until ts + window; a
    // record swept at now + window would leave it replayable in between.
    this.seenUpgrades.set(replayKey, ts + this.hmacWindowMs);

    const ip = normalizeIp(request.socket.remoteAddress);
    // Bound total + per-IP concurrent sockets before accepting the upgrade.
    if (this.sessions.size >= this.maxConnections) {
      this.config.logger?.warn(
        `MsteamsMediaStream: rejecting ${callId} — max connections (${this.maxConnections}) reached`,
      );
      this.rejectUpgrade(socket, 503, "Too Many Connections");
      return;
    }
    if ((this.connectionsByIp.get(ip) ?? 0) >= this.maxConnectionsPerIp) {
      this.config.logger?.warn(
        `MsteamsMediaStream: rejecting ${callId} — per-IP cap (${this.maxConnectionsPerIp}) reached for ${ip}`,
      );
      this.rejectUpgrade(socket, 503, "Too Many Connections");
      return;
    }

    wss.handleUpgrade(request, socket, head, (ws) => {
      this.attachSession(callId, ws, ip);
    });
  }

  private attachSession(callId: string, ws: WebSocket, ip: string): void {
    if (this.sessions.has(callId)) {
      // Same callId already connected — close the new one to avoid clobbering.
      try {
        ws.close(1008, "duplicate-callId");
      } catch {
        // ignore
      }
      this.config.logger?.warn(`MsteamsMediaStream: rejected duplicate connection for ${callId}`);
      return;
    }

    this.sessions.set(callId, ws);
    this.connectionsByIp.set(ip, (this.connectionsByIp.get(ip) ?? 0) + 1);
    // Reap sockets that authenticate but never send session.start (idle hold).
    const preStartTimer = setTimeout(() => {
      this.config.logger?.warn(
        `MsteamsMediaStream: closing ${callId} — no session.start within ${this.preStartTimeoutMs}ms`,
      );
      this.closeSession(callId, "pre-start-timeout");
    }, this.preStartTimeoutMs);
    if (typeof preStartTimer.unref === "function") {
      preStartTimer.unref();
    }
    this.connectionMeta.set(callId, { ip, started: false, ended: false, preStartTimer });
    this.config.logger?.info(`MsteamsMediaStream: connection open ${callId}`);

    // Heartbeat liveness: mark alive on connect + on every protocol pong. The interval in start() pings each
    // client and terminates any that didn't pong since the last sweep (dead caller), freeing its concurrency slot.
    const live = ws as WebSocket & { isAlive?: boolean };
    live.isAlive = true;
    ws.on("pong", () => {
      live.isAlive = true;
    });

    ws.on("message", (data) => this.handleMessage(callId, data));
    ws.on("close", () => {
      // An abrupt socket close (worker crash, network loss, hangup without a
      // session.end frame) must still tear down provider + manager state for a
      // session that already started — otherwise the call record leaks until the
      // stale-call reaper. The `ended` guard avoids double-delivery when the
      // close follows an explicit session.end.
      const meta = this.connectionMeta.get(callId);
      if (meta?.started && !meta.ended) {
        meta.ended = true;
        this.config.onSessionEnd?.({ callId, reason: "socket-closed" });
      }
      this.cleanupConnection(callId);
      this.config.logger?.info(`MsteamsMediaStream: connection closed ${callId}`);
    });
    ws.on("error", (err) => {
      const message = err instanceof Error ? err.message : String(err);
      this.config.logger?.warn(`MsteamsMediaStream: ws error ${callId} — ${message}`);
    });
  }

  /** Release per-connection tracking (timer, per-IP count, session entry). */
  private cleanupConnection(callId: string): void {
    const meta = this.connectionMeta.get(callId);
    if (meta) {
      clearTimeout(meta.preStartTimer);
      const remaining = (this.connectionsByIp.get(meta.ip) ?? 1) - 1;
      if (remaining > 0) {
        this.connectionsByIp.set(meta.ip, remaining);
      } else {
        this.connectionsByIp.delete(meta.ip);
      }
      this.connectionMeta.delete(callId);
    }
    this.sessions.delete(callId);
  }

  private handleMessage(callId: string, data: RawData): void {
    const text = rawDataToString(data);
    if (text === null) {
      return;
    }

    let parsed: InboundMessage;
    try {
      const raw = JSON.parse(text) as Record<string, unknown> | null;
      // The hosted bridge sends direction:"join" for meeting joins; the protocol enum only
      // has inbound|outbound and receivers default to inbound when absent - normalize
      // unknown values instead of rejecting the whole session.start.
      if (
        raw !== null &&
        typeof raw === "object" &&
        raw.type === "session.start" &&
        raw.direction !== undefined &&
        raw.direction !== "inbound" &&
        raw.direction !== "outbound"
      ) {
        raw.direction = "inbound";
      }
      parsed = InboundMessageSchema.parse(raw);
    } catch (err) {
      this.config.logger?.warn(
        `MsteamsMediaStream: invalid message from ${callId}: ${(err as Error).message}`,
      );
      return;
    }

    // Pre-GA (openclaw): one case handler throwing must not crash the per-call WS message pump.
    try {
    switch (parsed.type) {
      case "session.start": {
        // The callId is authenticated via HMAC in the URL path; a session.start
        // body claiming a different callId must be rejected, otherwise the call
        // record and the send/close closures would key off different ids.
        if (parsed.callId !== callId) {
          this.config.logger?.warn(
            `MsteamsMediaStream: session.start callId mismatch (authenticated=${callId} payload=${parsed.callId}); closing`,
          );
          this.closeSession(callId, "callid-mismatch");
          return;
        }
        const meta = this.connectionMeta.get(callId);
        if (meta) {
          meta.started = true;
          clearTimeout(meta.preStartTimer);
        }
        this.config.onSessionStart?.({
          callId,
          threadId: parsed.threadId,
          // Blank ids become null at the boundary: an empty-string aadId would survive every
          // downstream `aadId ?? fallback` and collapse all such callers into one session key
          // (cross-caller memory bleed) or one delivery target.
          caller: {
            aadId: blankToNull(parsed.caller.aadId),
            displayName: blankToNull(parsed.caller.displayName),
            tenantId: blankToNull(parsed.caller.tenantId),
          },
          recordingStatus: parsed.recordingStatus,
          direction: parsed.direction,
          send: (message) => this.sendTo(callId, message),
          close: (reason) => this.closeSession(callId, reason),
        });
        break;
      }
      case "recording.status": {
        this.config.onRecordingStatus?.({ callId, status: parsed.status });
        break;
      }
      case "session.end": {
        const meta = this.connectionMeta.get(callId);
        if (meta) {
          meta.ended = true;
        }
        this.config.onSessionEnd?.({ callId, reason: parsed.reason });
        this.closeSession(callId, parsed.reason);
        break;
      }
      case "audio.frame": {
        this.config.onAudioFrame?.({
          callId,
          seq: parsed.seq,
          timestampMs: parsed.timestampMs,
          payload: Buffer.from(parsed.payloadBase64, "base64"),
          speakerName: parsed.speakerName,
        });
        break;
      }
      case "video.frame": {
        this.config.onVideoFrame?.({
          callId,
          source: parsed.source,
          ts: parsed.ts,
          width: parsed.width,
          height: parsed.height,
          mime: parsed.mime,
          dataBase64: parsed.dataBase64,
          participantId: parsed.participantId,
          participantName: parsed.participantName,
        });
        break;
      }
      case "participants": {
        this.config.onParticipants?.({ callId, count: parsed.count });
        break;
      }
      case "dtmf": {
        this.config.onDtmf?.({ callId, digit: parsed.digit });
        break;
      }
      case "assistant.say": {
        this.config.onAssistantSay?.({ callId, text: parsed.text });
        break;
      }
      case "ping": {
        this.sendTo(callId, { type: "pong", ts: parsed.ts });
        break;
      }
    }
    } catch (err) {
      this.config.logger?.warn(
        `MsteamsMediaStream: error handling ${parsed.type} for ${callId}: ${(err as Error).message}`,
      );
    }
  }

  // Returns whether the message was actually sent. A closed/missing socket drops it (false) rather
  // than throwing, so best-effort control frames stay no-ops while delivery-sensitive callers (audio
  // frames) can observe the drop and abort.
  private sendTo(callId: string, message: unknown): boolean {
    const ws = this.sessions.get(callId);
    if (ws && ws.readyState === WebSocket.OPEN) {
      // Backpressure: shed ONLY realtime frames (audio.frame, and display.frame =
      // avatar-relay video) when the outbound buffer is large. Control/one-shot
      // frames (assistant.cancel, session.end, expression, display.image, pong,
      // speech.marks) must NOT be dropped — dropping a cancel or an image the model
      // was told succeeded desyncs the call. Mirrors the standalone bridges'
      // droppable/undroppable split (BRIDGE-11).
      const type = (message as { type?: unknown } | null)?.type;
      const droppable = type === "audio.frame" || type === "display.frame";
      if (droppable && ws.bufferedAmount > MAX_OUTBOUND_BUFFER_BYTES) {
        this.config.logger?.warn(
          `MsteamsMediaStream: dropping ${String(type)} for ${callId} — send buffer backpressure (${ws.bufferedAmount} bytes)`,
        );
        return false;
      }
      ws.send(JSON.stringify(message));
      return true;
    }
    return false;
  }

  private closeSession(callId: string, reason: string): void {
    const ws = this.sessions.get(callId);
    if (!ws) {
      return;
    }
    // A host-initiated close must still deliver onSessionEnd for a started session:
    // cleanupConnection below deletes the connection meta, so by the time the ws
    // close event fires there is no meta and the close handler skips delivery —
    // which leaked host call state on paths like a realtime connect-failure close
    // (the call stayed "in-progress" forever). The `ended` guard keeps a
    // session.end-driven close single-delivery.
    const meta = this.connectionMeta.get(callId);
    if (meta?.started && !meta.ended) {
      meta.ended = true;
      this.config.onSessionEnd?.({ callId, reason });
    }
    try {
      ws.close(1000, reason);
    } catch {
      // ignore
    }
    this.cleanupConnection(callId);
  }

  private rejectUpgrade(socket: Duplex, code: number, reason: string): void {
    // The peer may already have torn the socket down (scanner RST): writing to a
    // destroyed stream raises an async error, so only write while it is alive.
    if (!socket.destroyed) {
      socket.write(`HTTP/1.1 ${code} ${reason}\r\nConnection: close\r\n\r\n`);
    }
    socket.destroy();
  }
}

/** Blank/whitespace-only strings become null so downstream `?? fallback` checks fire. */
function blankToNull(value: string | null | undefined): string | null | undefined {
  return typeof value === "string" && value.trim() === "" ? null : value;
}

function normalizeIp(raw: string | undefined): string {
  if (!raw) {
    return "unknown";
  }
  // Collapse IPv4-mapped IPv6 (::ffff:1.2.3.4 -> 1.2.3.4) for stable per-IP keys.
  return raw.startsWith("::ffff:") ? raw.slice("::ffff:".length) : raw;
}

function rawDataToString(data: RawData): string | null {
  if (typeof data === "string") {
    return data;
  }
  if (Buffer.isBuffer(data)) {
    return data.toString("utf8");
  }
  if (Array.isArray(data)) {
    return Buffer.concat(data).toString("utf8");
  }
  // ArrayBuffer fallback
  if (data instanceof ArrayBuffer) {
    return Buffer.from(data).toString("utf8");
  }
  return null;
}
