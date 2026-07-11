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
import { safeEqualSecret } from "openclaw/plugin-sdk/security-runtime";
import { WebSocket, WebSocketServer } from "ws";
import { InboundMessageSchema } from "./protocol.gen.js";
// The wire protocol (message schemas + shared constants) is generated from the
// the shared wire-protocol schema, mirrored in src/protocol.gen.ts. This file
// keeps only the transport: HMAC/replay/connection guards and message routing.
export { MSTEAMS_PCM_SAMPLE_RATE_HZ } from "./protocol.gen.js";
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
/**
 * Backpressure bound on a single call's OUTBOUND send buffer. Audio egress (sendTo) is
 * fire-and-forget; if the worker stalls, ws.bufferedAmount grows unbounded and leaks memory. When
 * the buffer is already this deep we drop the frame (return false) instead of queuing more — audio
 * is real-time, so a dropped stale frame beats ever-growing latency debt.
 */
const MAX_OUTBOUND_BUFFER_BYTES = 1 * 1024 * 1024;
export class MsteamsMediaStream {
    config;
    hmacWindowMs;
    /** Verified upgrade tuples already used once (replay guard); value = expiry epoch ms. */
    seenUpgrades = new Map();
    maxConnections;
    maxConnectionsPerIp;
    preStartTimeoutMs;
    sessions = new Map();
    connectionMeta = new Map();
    connectionsByIp = new Map();
    server;
    wss;
    constructor(config) {
        this.config = config;
        this.hmacWindowMs = config.hmacWindowMs ?? DEFAULT_HMAC_WINDOW_MS;
        this.maxConnections = config.maxConnections ?? DEFAULT_MAX_CONNECTIONS;
        this.maxConnectionsPerIp = config.maxConnectionsPerIp ?? DEFAULT_MAX_CONNECTIONS_PER_IP;
        this.preStartTimeoutMs = config.preStartTimeoutMs ?? DEFAULT_PRE_START_TIMEOUT_MS;
    }
    async start() {
        if (this.server) {
            throw new Error("MsteamsMediaStream is already started");
        }
        const server = http.createServer();
        const wss = new WebSocketServer({ noServer: true, maxPayload: MAX_INBOUND_PAYLOAD_BYTES });
        server.on("upgrade", (request, socket, head) => {
            this.handleUpgrade(request, socket, head, wss);
        });
        await new Promise((resolve, reject) => {
            const onError = (err) => {
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
        this.config.logger?.info(`MsteamsMediaStream listening host=${this.config.bindAddress ?? DEFAULT_BIND_ADDRESS} port=${this.config.port} path=${this.config.path}`);
    }
    async stop() {
        if (!this.server) {
            return;
        }
        for (const ws of this.sessions.values()) {
            try {
                ws.close(1001, "shutdown");
            }
            catch {
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
        await new Promise((resolve) => {
            // server is non-null because we early-returned above; cast is for ts narrowing
            const s = this.server;
            s.close(() => resolve());
        });
        this.server = undefined;
        this.wss = undefined;
        this.config.logger?.info("MsteamsMediaStream stopped");
    }
    /** Number of currently open sessions. Exposed for tests + telemetry. */
    get sessionCount() {
        return this.sessions.size;
    }
    handleUpgrade(request, socket, head, wss) {
        const url = new URL(request.url ?? "", "http://localhost");
        if (!url.pathname.startsWith(this.config.path)) {
            this.rejectUpgrade(socket, 404, "Not Found");
            return;
        }
        const callId = url.pathname.slice(this.config.path.length).replace(/^\//, "");
        if (!callId) {
            this.rejectUpgrade(socket, 400, "Bad Request (missing callId)");
            return;
        }
        const timestamp = request.headers["x-openclawteamsbridge-timestamp"];
        const signature = request.headers["x-openclawteamsbridge-signature"];
        if (typeof timestamp !== "string" || typeof signature !== "string") {
            this.config.logger?.warn(`MsteamsMediaStream: rejecting upgrade for ${callId} — missing HMAC headers`);
            this.rejectUpgrade(socket, 401, "Unauthorized");
            return;
        }
        const ts = Number(timestamp);
        if (!Number.isFinite(ts) || Math.abs(Date.now() - ts) > this.hmacWindowMs) {
            this.config.logger?.warn(`MsteamsMediaStream: rejecting upgrade for ${callId} — timestamp out of window`);
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
            this.config.logger?.warn(`MsteamsMediaStream: rejecting upgrade for ${callId} — bad signature`);
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
            this.config.logger?.warn(`MsteamsMediaStream: rejecting upgrade for ${callId} — replayed handshake`);
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
            this.config.logger?.warn(`MsteamsMediaStream: rejecting ${callId} — max connections (${this.maxConnections}) reached`);
            this.rejectUpgrade(socket, 503, "Too Many Connections");
            return;
        }
        if ((this.connectionsByIp.get(ip) ?? 0) >= this.maxConnectionsPerIp) {
            this.config.logger?.warn(`MsteamsMediaStream: rejecting ${callId} — per-IP cap (${this.maxConnectionsPerIp}) reached for ${ip}`);
            this.rejectUpgrade(socket, 503, "Too Many Connections");
            return;
        }
        wss.handleUpgrade(request, socket, head, (ws) => {
            this.attachSession(callId, ws, ip);
        });
    }
    attachSession(callId, ws, ip) {
        if (this.sessions.has(callId)) {
            // Same callId already connected — close the new one to avoid clobbering.
            try {
                ws.close(1008, "duplicate-callId");
            }
            catch {
                // ignore
            }
            this.config.logger?.warn(`MsteamsMediaStream: rejected duplicate connection for ${callId}`);
            return;
        }
        this.sessions.set(callId, ws);
        this.connectionsByIp.set(ip, (this.connectionsByIp.get(ip) ?? 0) + 1);
        // Reap sockets that authenticate but never send session.start (idle hold).
        const preStartTimer = setTimeout(() => {
            this.config.logger?.warn(`MsteamsMediaStream: closing ${callId} — no session.start within ${this.preStartTimeoutMs}ms`);
            this.closeSession(callId, "pre-start-timeout");
        }, this.preStartTimeoutMs);
        if (typeof preStartTimer.unref === "function") {
            preStartTimer.unref();
        }
        this.connectionMeta.set(callId, { ip, started: false, ended: false, preStartTimer });
        this.config.logger?.info(`MsteamsMediaStream: connection open ${callId}`);
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
    cleanupConnection(callId) {
        const meta = this.connectionMeta.get(callId);
        if (meta) {
            clearTimeout(meta.preStartTimer);
            const remaining = (this.connectionsByIp.get(meta.ip) ?? 1) - 1;
            if (remaining > 0) {
                this.connectionsByIp.set(meta.ip, remaining);
            }
            else {
                this.connectionsByIp.delete(meta.ip);
            }
            this.connectionMeta.delete(callId);
        }
        this.sessions.delete(callId);
    }
    handleMessage(callId, data) {
        const text = rawDataToString(data);
        if (text === null) {
            return;
        }
        let parsed;
        try {
            const raw = JSON.parse(text);
            // The hosted bridge sends direction:"join" for meeting joins; the protocol enum only
            // has inbound|outbound and receivers default to inbound when absent - normalize
            // unknown values instead of rejecting the whole session.start.
            if (raw !== null &&
                typeof raw === "object" &&
                raw.type === "session.start" &&
                raw.direction !== undefined &&
                raw.direction !== "inbound" &&
                raw.direction !== "outbound") {
                raw.direction = "inbound";
            }
            parsed = InboundMessageSchema.parse(raw);
        }
        catch (err) {
            this.config.logger?.warn(`MsteamsMediaStream: invalid message from ${callId}: ${err.message}`);
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
                        this.config.logger?.warn(`MsteamsMediaStream: session.start callId mismatch (authenticated=${callId} payload=${parsed.callId}); closing`);
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
        }
        catch (err) {
            this.config.logger?.warn(`MsteamsMediaStream: error handling ${parsed.type} for ${callId}: ${err.message}`);
        }
    }
    // Returns whether the message was actually sent. A closed/missing socket drops it (false) rather
    // than throwing, so best-effort control frames stay no-ops while delivery-sensitive callers (audio
    // frames) can observe the drop and abort.
    sendTo(callId, message) {
        const ws = this.sessions.get(callId);
        if (ws && ws.readyState === WebSocket.OPEN) {
            // Backpressure: a slow/stalled worker must not let the outbound buffer grow without bound.
            if (ws.bufferedAmount > MAX_OUTBOUND_BUFFER_BYTES) {
                this.config.logger?.warn(`MsteamsMediaStream: dropping frame for ${callId} — send buffer backpressure (${ws.bufferedAmount} bytes)`);
                return false;
            }
            ws.send(JSON.stringify(message));
            return true;
        }
        return false;
    }
    closeSession(callId, reason) {
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
        }
        catch {
            // ignore
        }
        this.cleanupConnection(callId);
    }
    rejectUpgrade(socket, code, reason) {
        socket.write(`HTTP/1.1 ${code} ${reason}\r\nConnection: close\r\n\r\n`);
        socket.destroy();
    }
}
/** Blank/whitespace-only strings become null so downstream `?? fallback` checks fire. */
function blankToNull(value) {
    return typeof value === "string" && value.trim() === "" ? null : value;
}
function normalizeIp(raw) {
    if (!raw) {
        return "unknown";
    }
    // Collapse IPv4-mapped IPv6 (::ffff:1.2.3.4 -> 1.2.3.4) for stable per-IP keys.
    return raw.startsWith("::ffff:") ? raw.slice("::ffff:".length) : raw;
}
function rawDataToString(data) {
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
