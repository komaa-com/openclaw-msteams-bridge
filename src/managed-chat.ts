// StandIn MANAGED chat mode (MANAGED-BOT-TIER.md 4.8, protocol/chat-schema.yaml).
//
// On the managed tier the customer does NOT own a Teams bot: StandIn's gateway terminates Bot
// Framework and speaks the normalized chat protocol to this agent instead. This module is that
// endpoint: an HTTP server accepting InboundMessage (bridge-HMAC-signed with the binding's CHAT
// key), answering 200 immediately (the gateway's durable relay handles retry/ordering), consulting
// the agent async, and POSTing the reply back to the gateway's /api/chat/reply signed with the SAME
// key. The agent never holds a Bot Framework credential — that is the whole point (D5).
//
// The voice WebSocket (msteams-media-stream) is UNCHANGED by managed mode; only chat gains a lane.

import { createHmac, timingSafeEqual } from "node:crypto";
import http from "node:http";

/** Accepted clock skew for the bridge HMAC, both directions (chat-schema.yaml REPLAY_WINDOW_MS). */
export const REPLAY_WINDOW_MS = 300_000;

/** Value carried in every message's schemaVersion (chat-schema.yaml SCHEMA_VERSION). */
export const SCHEMA_VERSION = 1;

export interface ManagedChatConfig {
  enabled: boolean;
  port: number;
  bindAddress?: string;
  path: string;
  /** The binding's CHAT key (separate, domain-derived from the voice key — section 9). */
  chatSecret: string;
  /** The gateway's reply endpoint, e.g. https://teams.standin.komaa.com/api/chat/reply */
  gatewayReplyUrl: string;
}

/** The inbound fields this agent actually consumes (a SUBSET of chat-schema.yaml InboundMessage —
 * subsets are legal; unknown fields are ignored by contract). */
export interface ManagedInbound {
  tenantId: string;
  conversationId: string;
  activityId: string;
  scope: string;
  text: string;
  sender: { aadObjectId?: string; displayName?: string; isGuest?: boolean; isLinkedOwner?: boolean };
  attachments?: Array<{ kind: string; name?: string; url?: string; relayable?: boolean }>;
  locale?: string;
}

// ── bridge HMAC (identical construction to @standin/bridge-hmac; KAT-pinned in the tests) ─────────

export function signBridge(
  secret: string,
  body: string,
  nowMs = Date.now(),
): { timestamp: string; signature: string } {
  const timestamp = String(nowMs);
  return { timestamp, signature: computeBridgeSignature(secret, timestamp, body) };
}

export function computeBridgeSignature(secret: string, timestamp: string, body: string): string {
  return createHmac("sha256", secret).update(`${timestamp}.${body}`).digest("hex");
}

export function verifyBridge(
  secret: string,
  timestamp: string | undefined,
  body: string,
  signature: string | undefined,
  nowMs = Date.now(),
): boolean {
  if (!secret || !timestamp || !signature) return false;
  const ts = Number(timestamp);
  if (!Number.isFinite(ts)) return false;
  if (Math.abs(nowMs - ts) > REPLAY_WINDOW_MS) return false;
  const expected = Buffer.from(computeBridgeSignature(secret, timestamp, body), "utf8");
  const provided = Buffer.from(signature.toLowerCase(), "utf8");
  return expected.length === provided.length && timingSafeEqual(expected, provided);
}

// ── inbound parsing (tolerant of unknown fields, strict on the routing keys) ──────────────────────

export function parseInbound(body: string): { ok: true; message: ManagedInbound } | { ok: false; error: string } {
  let raw: unknown;
  try {
    raw = JSON.parse(body);
  } catch {
    return { ok: false, error: "malformed json" };
  }
  if (typeof raw !== "object" || raw === null) return { ok: false, error: "body must be an object" };
  const m = raw as Record<string, unknown>;
  for (const key of ["tenantId", "conversationId", "activityId"] as const) {
    if (typeof m[key] !== "string" || (m[key] as string).length === 0) {
      return { ok: false, error: `${key} is required` };
    }
  }
  const sender = (typeof m.sender === "object" && m.sender !== null ? m.sender : {}) as ManagedInbound["sender"];
  return {
    ok: true,
    message: {
      tenantId: m.tenantId as string,
      conversationId: m.conversationId as string,
      activityId: m.activityId as string,
      scope: typeof m.scope === "string" ? m.scope : "personal",
      text: typeof m.text === "string" ? m.text : "",
      sender,
      attachments: Array.isArray(m.attachments) ? (m.attachments as ManagedInbound["attachments"]) : undefined,
      locale: typeof m.locale === "string" ? m.locale : undefined,
    },
  };
}

/** At-least-once dedupe: the gateway REDELIVERS on retry, and activityId is the idempotency key the
 * schema tells agents to honor. Bounded LRU so the window cannot grow without limit. */
export class SeenActivities {
  private readonly seen = new Map<string, true>();

  constructor(private readonly capacity = 2048) {}

  /** True the FIRST time an activity id is offered; false on a redelivery. */
  markFirst(activityId: string): boolean {
    if (this.seen.has(activityId)) return false;
    this.seen.set(activityId, true);
    if (this.seen.size > this.capacity) {
      const oldest = this.seen.keys().next().value;
      if (oldest !== undefined) this.seen.delete(oldest);
    }
    return true;
  }
}

/** Build the reply the gateway expects. tenantId/conversationId echo the inbound EXACTLY — the
 * gateway rejects a tenant mismatch (the cross-tenant guard, the load-bearing check of the relay). */
export function buildReply(
  inbound: Pick<ManagedInbound, "tenantId" | "conversationId" | "activityId">,
  text: string,
  kind: "message" | "typing" | "error" = "message",
): Record<string, unknown> {
  return {
    schemaVersion: SCHEMA_VERSION,
    tenantId: inbound.tenantId,
    conversationId: inbound.conversationId,
    replyToId: inbound.activityId,
    kind,
    ...(kind === "typing" ? {} : { text }),
    idempotencyKey: `${inbound.activityId}:${kind}`,
  };
}

// ── the server ────────────────────────────────────────────────────────────────────────────────────

/** An image fetched from the gateway's signed URL, ready for the agent consult. */
export interface FetchedImage {
  type: "image";
  data: string;
  mimeType: string;
}

/**
 * Fetch RELAYABLE IMAGE attachments from their gateway-signed URLs into consult images (4.7: "the
 * agent fetches within the reference TTL"). Images only - the consult accepts images, and files are
 * still named in the turn text. Size-capped and best-effort per attachment: one bad fetch drops that
 * image, never the turn.
 */
export async function fetchAttachmentImages(
  attachments: ManagedInbound["attachments"],
  opts?: { fetchFn?: typeof fetch; maxBytes?: number },
): Promise<FetchedImage[]> {
  const fetchFn = opts?.fetchFn ?? fetch;
  const maxBytes = opts?.maxBytes ?? 4 * 1024 * 1024;
  const images: FetchedImage[] = [];
  for (const a of attachments ?? []) {
    if (a.kind !== "image" || a.relayable === false || !a.url) continue;
    try {
      const res = await fetchFn(a.url);
      if (!res.ok) continue;
      const buf = Buffer.from(await res.arrayBuffer());
      if (buf.length === 0 || buf.length > maxBytes) continue;
      const mime = (a as { contentType?: string }).contentType
        ?? res.headers.get("content-type")
        ?? "image/png";
      images.push({ type: "image", data: buf.toString("base64"), mimeType: mime });
    } catch {
      // Best-effort: the turn still runs; the text names the attachment either way.
    }
  }
  return images;
}

export interface ManagedChatDeps {
  /** Run one agent turn for an inbound message; returns the reply text. */
  respond: (message: ManagedInbound) => Promise<string>;
  log: { info: (m: string) => void; warn: (m: string) => void; error: (m: string) => void };
  /** Injectable for tests. */
  fetchFn?: typeof fetch;
  nowMs?: () => number;
}

export class ManagedChatServer {
  private server?: http.Server;
  private readonly seen = new SeenActivities();
  /** Per-conversation processing chains (review P0-4): the schema promises per-conversation ORDERING,
   * and independent tasks per message let replies overtake each other. Each conversation's turns run
   * strictly sequentially; different conversations still run concurrently. Entries are cleaned when
   * their chain drains so idle conversations cost nothing. */
  private readonly chains = new Map<string, Promise<void>>();

  constructor(
    private readonly cfg: ManagedChatConfig,
    private readonly deps: ManagedChatDeps,
  ) {}

  async start(): Promise<void> {
    const server = http.createServer((req, res) => {
      void this.handle(req, res);
    });
    this.server = server;
    await new Promise<void>((resolve, reject) => {
      server.once("error", reject);
      server.listen(this.cfg.port, this.cfg.bindAddress, () => resolve());
    });
    this.deps.log.info(
      `msteams managed chat: listening on ${this.cfg.bindAddress ?? "0.0.0.0"}:${this.cfg.port}${this.cfg.path}`,
    );
  }

  async stop(): Promise<void> {
    const server = this.server;
    this.server = undefined;
    if (server) await new Promise<void>((resolve) => server.close(() => resolve()));
  }

  private async handle(req: http.IncomingMessage, res: http.ServerResponse): Promise<void> {
    if (req.method !== "POST" || (req.url ?? "").split("?")[0] !== this.cfg.path) {
      res.writeHead(404).end();
      return;
    }
    // Bounded read BEFORE any auth work (review P0-4): an unauthenticated peer must not make us buffer
    // an arbitrary body. 1 MB comfortably fits any relay payload (attachments travel by REFERENCE).
    const maxBody = 1024 * 1024;
    const declared = Number(req.headers["content-length"]);
    if (Number.isFinite(declared) && declared > maxBody) {
      res.writeHead(413).end();
      return;
    }
    const chunks: Buffer[] = [];
    let received = 0;
    for await (const chunk of req) {
      received += (chunk as Buffer).length;
      if (received > maxBody) {
        res.writeHead(413).end();
        req.destroy();
        return;
      }
      chunks.push(chunk as Buffer);
    }
    const body = Buffer.concat(chunks).toString("utf8");

    const ts = header(req, "x-standin-timestamp");
    const sig = header(req, "x-standin-signature");
    if (!verifyBridge(this.cfg.chatSecret, ts, body, sig, this.deps.nowMs?.() ?? Date.now())) {
      res.writeHead(401, { "content-type": "application/json" }).end(JSON.stringify({ error: "unauthorized" }));
      return;
    }

    const parsed = parseInbound(body);
    if (!parsed.ok) {
      res.writeHead(400, { "content-type": "application/json" }).end(JSON.stringify({ error: parsed.error }));
      return;
    }

    // ACK FIRST (the gateway's relay window is short; agent latency is not) — then process async.
    // A redelivered activity ACKs and does nothing: the first delivery's turn is already running.
    const fresh = this.seen.markFirst(parsed.message.activityId);
    res.writeHead(200, { "content-type": "application/json" }).end(JSON.stringify({ ok: true }));
    if (!fresh) return;

    this.enqueueTurn(parsed.message);
  }

  /** Chain the turn behind the conversation's previous one (ordering); see `chains`. */
  private enqueueTurn(message: ManagedInbound): void {
    const key = `${message.tenantId}:${message.conversationId}`;
    const prev = this.chains.get(key) ?? Promise.resolve();
    const next = prev.then(() => this.processAsync(message)).catch(() => undefined);
    this.chains.set(key, next);
    void next.finally(() => {
      if (this.chains.get(key) === next) this.chains.delete(key);
    });
  }

  /** Re-review P2: serialization makes a HUNG turn wedge its whole conversation forever (every later
   * message chains behind it, and the chain entry never drains). Every turn is therefore bounded: a
   * turn that exceeds the budget fails like any other error, the user hears about it, and the chain
   * moves on. Generous, because agent turns legitimately run long. */
  static readonly TURN_TIMEOUT_MS = 5 * 60 * 1000;

  private async processAsync(message: ManagedInbound): Promise<void> {
    // Typing while the agent thinks — best-effort, ephemeral by design.
    await this.postReply(buildReply(message, "", "typing")).catch(() => undefined);
    try {
      const text = await withTimeout(this.deps.respond(message), ManagedChatServer.TURN_TIMEOUT_MS, "agent turn");
      if (text.trim().length > 0) {
        await this.postReply(buildReply(message, text));
      }
    } catch (err) {
      this.deps.log.error(`msteams managed chat: agent turn failed: ${String(err)}`);
      await this.postReply(
        buildReply(message, "Something went wrong answering that — please try again.", "error"),
      ).catch(() => undefined);
    }
  }

  private async postReply(reply: Record<string, unknown>): Promise<void> {
    const body = JSON.stringify(reply);
    const { timestamp, signature } = signBridge(this.cfg.chatSecret, body, this.deps.nowMs?.() ?? Date.now());
    const fetchFn = this.deps.fetchFn ?? fetch;
    const res = await fetchFn(this.cfg.gatewayReplyUrl, {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-standin-timestamp": timestamp,
        "x-standin-signature": signature,
      },
      body,
      // Re-review P2: an unbounded reply POST wedges the conversation chain exactly like a hung turn.
      signal: AbortSignal.timeout(30_000),
    });
    if (!res.ok) {
      this.deps.log.warn(`msteams managed chat: gateway reply -> HTTP ${res.status}`);
    }
  }
}

/** Reject after ms; the underlying promise keeps running (we cannot cancel the agent) but the chain
 * stops waiting on it — a wedged conversation beats a cancelled-mid-tool-use agent turn either way. */
function withTimeout<T>(promise: Promise<T>, ms: number, what: string): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error(`${what} timed out after ${ms}ms`)), ms);
    promise.then(
      (v) => { clearTimeout(timer); resolve(v); },
      (e) => { clearTimeout(timer); reject(e); },
    );
  });
}

function header(req: http.IncomingMessage, name: string): string | undefined {
  const v = req.headers[name];
  return Array.isArray(v) ? v[0] : v;
}

/** Resolve the managed-chat block of the plugin config; disabled unless BOTH secrets of the lane are
 * present (an enabled endpoint without a verifiable key would accept nothing anyway — fail closed). */
export function resolveManagedChatConfig(raw: unknown): ManagedChatConfig {
  const c = (raw ?? {}) as Record<string, unknown>;
  const chatSecret = typeof c.chatSecret === "string" ? c.chatSecret : "";
  return {
    enabled: c.enabled === true && chatSecret.length > 0,
    port: Number(c.port ?? 9444),
    bindAddress: typeof c.bindAddress === "string" ? c.bindAddress : undefined,
    path: typeof c.path === "string" ? c.path : "/managed/chat",
    chatSecret,
    gatewayReplyUrl:
      typeof c.gatewayReplyUrl === "string" && c.gatewayReplyUrl.length > 0
        ? c.gatewayReplyUrl
        : "https://teams.standin.komaa.com/api/chat/reply",
  };
}
