import crypto from "node:crypto";
import http from "node:http";
import { safeEqualSecret } from "openclaw/plugin-sdk/security-runtime";
import { WebSocket, WebSocketServer } from "ws";
import { InboundMessageSchema } from "./protocol.gen.js";
export { MSTEAMS_PCM_SAMPLE_RATE_HZ } from "./protocol.gen.js";
const DEFAULT_HMAC_WINDOW_MS = 60_000;
const DEFAULT_BIND_ADDRESS = "127.0.0.1";
const MAX_INBOUND_PAYLOAD_BYTES = 2 * 1024 * 1024;
const DEFAULT_MAX_CONNECTIONS = 64;
const DEFAULT_MAX_CONNECTIONS_PER_IP = 8;
const DEFAULT_PRE_START_TIMEOUT_MS = 10_000;
const MAX_OUTBOUND_BUFFER_BYTES = 1 * 1024 * 1024;
export class MsteamsMediaStream {
    config;
    hmacWindowMs;
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
            const s = this.server;
            s.close(() => resolve());
        });
        this.server = undefined;
        this.wss = undefined;
        this.config.logger?.info("MsteamsMediaStream stopped");
    }
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
        const timestamp = request.headers["x-standin-timestamp"] ?? request.headers["x-openclawteamsbridge-timestamp"];
        const signature = request.headers["x-standin-signature"] ?? request.headers["x-openclawteamsbridge-signature"];
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
        const now = Date.now();
        for (const [key, expiry] of this.seenUpgrades) {
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
        this.seenUpgrades.set(replayKey, ts + this.hmacWindowMs);
        const ip = normalizeIp(request.socket.remoteAddress);
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
            try {
                ws.close(1008, "duplicate-callId");
            }
            catch {
            }
            this.config.logger?.warn(`MsteamsMediaStream: rejected duplicate connection for ${callId}`);
            return;
        }
        this.sessions.set(callId, ws);
        this.connectionsByIp.set(ip, (this.connectionsByIp.get(ip) ?? 0) + 1);
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
        try {
            switch (parsed.type) {
                case "session.start": {
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
    sendTo(callId, message) {
        const ws = this.sessions.get(callId);
        if (ws && ws.readyState === WebSocket.OPEN) {
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
        const meta = this.connectionMeta.get(callId);
        if (meta?.started && !meta.ended) {
            meta.ended = true;
            this.config.onSessionEnd?.({ callId, reason });
        }
        try {
            ws.close(1000, reason);
        }
        catch {
        }
        this.cleanupConnection(callId);
    }
    rejectUpgrade(socket, code, reason) {
        socket.write(`HTTP/1.1 ${code} ${reason}\r\nConnection: close\r\n\r\n`);
        socket.destroy();
    }
}
function blankToNull(value) {
    return typeof value === "string" && value.trim() === "" ? null : value;
}
function normalizeIp(raw) {
    if (!raw) {
        return "unknown";
    }
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
    if (data instanceof ArrayBuffer) {
        return Buffer.from(data).toString("utf8");
    }
    return null;
}
