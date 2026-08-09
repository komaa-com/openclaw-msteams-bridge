// StandIn managed chat mode (MANAGED-BOT-TIER.md 4.8): the agent side of the normalized chat relay.
// The HMAC KAT is the SAME vector pinned in @standin/bridge-hmac (TS), the gateway's LinkTokensTests
// (C#), and the media-bridge callers — four independent implementations, one set of bytes.

import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterEach, describe, expect, it } from "vitest";
import {
  buildReply,
  fetchAttachmentImages,
  computeBridgeSignature,
  ManagedChatServer,
  parseInbound,
  REPLAY_WINDOW_MS,
  resolveManagedChatConfig,
  SeenActivities,
  signBridge,
  verifyBridge,
  type ManagedInbound,
} from "./managed-chat.js";

const KAT_SECRET = "test-secret";
const KAT_TS = "1700000000000";
const KAT_BODY = "hello";
const KAT_SIG = "1ea836ba1a9714e5a5824a9026b2b40567ee9e5e2ddd0d1cb598da3b42afce38";
const KAT_NOW = 1700000000000;

describe("bridge HMAC", () => {
  it("matches the cross-repo KAT", () => {
    expect(computeBridgeSignature(KAT_SECRET, KAT_TS, KAT_BODY)).toBe(KAT_SIG);
    const signed = signBridge(KAT_SECRET, KAT_BODY, KAT_NOW);
    expect(signed).toEqual({ timestamp: KAT_TS, signature: KAT_SIG });
  });

  it("verifies inside the replay window and refuses outside it", () => {
    expect(verifyBridge(KAT_SECRET, KAT_TS, KAT_BODY, KAT_SIG, KAT_NOW)).toBe(true);
    expect(verifyBridge(KAT_SECRET, KAT_TS, KAT_BODY, KAT_SIG, KAT_NOW + REPLAY_WINDOW_MS - 1)).toBe(true);
    expect(verifyBridge(KAT_SECRET, KAT_TS, KAT_BODY, KAT_SIG, KAT_NOW + REPLAY_WINDOW_MS + 1)).toBe(false);
    expect(verifyBridge(KAT_SECRET, KAT_TS, KAT_BODY, KAT_SIG, KAT_NOW - REPLAY_WINDOW_MS - 1)).toBe(false);
  });

  it("refuses tampering, wrong keys, and absent headers", () => {
    expect(verifyBridge(KAT_SECRET, KAT_TS, KAT_BODY + "x", KAT_SIG, KAT_NOW)).toBe(false);
    expect(verifyBridge("other-secret", KAT_TS, KAT_BODY, KAT_SIG, KAT_NOW)).toBe(false);
    expect(verifyBridge(KAT_SECRET, undefined, KAT_BODY, KAT_SIG, KAT_NOW)).toBe(false);
    expect(verifyBridge(KAT_SECRET, KAT_TS, KAT_BODY, undefined, KAT_NOW)).toBe(false);
    expect(verifyBridge("", KAT_TS, KAT_BODY, KAT_SIG, KAT_NOW)).toBe(false); // no secret = fail closed
    expect(verifyBridge(KAT_SECRET, "not-a-number", KAT_BODY, KAT_SIG, KAT_NOW)).toBe(false);
  });

  it("accepts an UPPERCASE hex signature (hex case is not part of the contract)", () => {
    expect(verifyBridge(KAT_SECRET, KAT_TS, KAT_BODY, KAT_SIG.toUpperCase(), KAT_NOW)).toBe(true);
  });
});

describe("inbound parsing", () => {
  const valid = JSON.stringify({
    schemaVersion: 1,
    tenantId: "t1",
    bindingId: "b1",
    conversationId: "c1",
    activityId: "a1",
    scope: "personal",
    sender: { displayName: "Alaa", isLinkedOwner: true },
    text: "hello agent",
    someFutureField: { ignored: true },
  });

  it("accepts a valid message and ignores unknown fields (additive evolution)", () => {
    const r = parseInbound(valid);
    expect(r.ok).toBe(true);
    if (r.ok) {
      expect(r.message.tenantId).toBe("t1");
      expect(r.message.text).toBe("hello agent");
      expect(r.message.sender.displayName).toBe("Alaa");
    }
  });

  it("requires the routing keys and rejects malformed bodies", () => {
    expect(parseInbound("not json").ok).toBe(false);
    expect(parseInbound("[]").ok).toBe(false);
    for (const missing of ["tenantId", "conversationId", "activityId"]) {
      const m = JSON.parse(valid) as Record<string, unknown>;
      delete m[missing];
      const r = parseInbound(JSON.stringify(m));
      expect(r.ok).toBe(false);
      if (!r.ok) expect(r.error).toContain(missing);
    }
  });
});

describe("redelivery dedupe", () => {
  it("is first-time-true, redelivery-false, and bounded", () => {
    const seen = new SeenActivities(2);
    expect(seen.markFirst("a")).toBe(true);
    expect(seen.markFirst("a")).toBe(false);
    expect(seen.markFirst("b")).toBe(true);
    expect(seen.markFirst("c")).toBe(true); // evicts "a"
    expect(seen.markFirst("a")).toBe(true); // aged out of the window - acceptable at-least-once behavior
  });
});

describe("reply building", () => {
  const inbound = { tenantId: "t1", conversationId: "c1", activityId: "a1" };

  it("echoes tenant + conversation exactly (the gateway's cross-tenant guard depends on it)", () => {
    const reply = buildReply(inbound, "answer");
    expect(reply.tenantId).toBe("t1");
    expect(reply.conversationId).toBe("c1");
    expect(reply.replyToId).toBe("a1");
    expect(reply.kind).toBe("message");
    expect(reply.text).toBe("answer");
    expect(reply.idempotencyKey).toBe("a1:message");
  });

  it("typing carries no text and its own idempotency key", () => {
    const typing = buildReply(inbound, "ignored", "typing");
    expect(typing.kind).toBe("typing");
    expect("text" in typing).toBe(false);
    expect(typing.idempotencyKey).toBe("a1:typing");
  });
});

describe("config resolution", () => {
  it("fails closed without a string chatSecret, exactly like the voice sharedSecret", () => {
    expect(resolveManagedChatConfig({ enabled: true, chatSecret: "k" }).enabled).toBe(true);
    expect(resolveManagedChatConfig({ enabled: true }).enabled).toBe(false);
    expect(resolveManagedChatConfig({ enabled: true, chatSecret: { env: "UNSET" } }).enabled).toBe(false);
    expect(resolveManagedChatConfig(undefined).enabled).toBe(false);
  });

  it("carries sane defaults", () => {
    const cfg = resolveManagedChatConfig({ enabled: true, chatSecret: "k" });
    expect(cfg.port).toBe(9444);
    expect(cfg.path).toBe("/managed/chat");
    expect(cfg.gatewayReplyUrl).toContain("/api/chat/reply");
  });
});

describe("schema drift (protocol/chat-schema.yaml is the source of truth)", () => {
  const here = dirname(fileURLToPath(import.meta.url));
  const schema = readFileSync(join(here, "..", "protocol", "chat-schema.yaml"), "utf8");

  it("every wire name this module reads or writes exists in the schema copy", () => {
    // A SUBSET is legal (unknown fields are ignored by contract); a name the schema does not know is
    // a typo that silently drops data. Line-based on purpose - no YAML dependency.
    const consumed = [
      "schemaVersion", "tenantId", "bindingId", "conversationId", "activityId", "scope", "sender",
      "text", "attachments", "locale", "replyToId", "kind", "idempotencyKey",
      "aadObjectId", "displayName", "isGuest", "isLinkedOwner", "name", "url", "relayable",
    ];
    for (const field of consumed) {
      expect(schema, `field '${field}' is not in chat-schema.yaml`).toMatch(
        new RegExp(`name: ${field}$`, "m"),
      );
    }
  });

  it("the schema constants this module hardcodes match the schema copy", () => {
    expect(schema).toContain("name: SCHEMA_VERSION\n    value: 1");
    expect(schema).toContain(`value: ${REPLAY_WINDOW_MS}`);
  });
});

describe("attachment image fetch (4.7 agent-side leg)", () => {
  const img = (over: Record<string, unknown> = {}) => ({
    kind: "image", name: "shot.png", url: "https://gw.test/api/chat/attachment/a1?e=1&s=x",
    contentType: "image/png", relayable: true, ...over,
  });
  const fakeFetch = (bytes: number, status = 200) =>
    (async () => new Response(new Uint8Array(bytes), { status, headers: { "content-type": "image/png" } })) as unknown as typeof fetch;

  it("fetches relayable images into base64 consult images", async () => {
    const images = await fetchAttachmentImages([img()], { fetchFn: fakeFetch(16) });
    expect(images).toHaveLength(1);
    expect(images[0].mimeType).toBe("image/png");
    expect(Buffer.from(images[0].data, "base64")).toHaveLength(16);
  });

  it("skips files, unrelayable, missing urls, failures, and oversize - never throws", async () => {
    expect(await fetchAttachmentImages([img({ kind: "file" })], { fetchFn: fakeFetch(16) })).toHaveLength(0);
    expect(await fetchAttachmentImages([img({ relayable: false })], { fetchFn: fakeFetch(16) })).toHaveLength(0);
    expect(await fetchAttachmentImages([img({ url: undefined })], { fetchFn: fakeFetch(16) })).toHaveLength(0);
    expect(await fetchAttachmentImages([img()], { fetchFn: fakeFetch(16, 404) })).toHaveLength(0);
    expect(await fetchAttachmentImages([img()], { fetchFn: fakeFetch(64, 200), maxBytes: 32 })).toHaveLength(0);
    expect(
      await fetchAttachmentImages([img()], { fetchFn: (async () => { throw new Error("net"); }) as unknown as typeof fetch }),
    ).toHaveLength(0);
    expect(await fetchAttachmentImages(undefined, { fetchFn: fakeFetch(16) })).toHaveLength(0);
  });
});

describe("the server end to end", () => {
  let server: ManagedChatServer | undefined;
  afterEach(async () => {
    await server?.stop();
    server = undefined;
  });

  async function startServer(opts?: { respond?: (m: ManagedInbound) => Promise<string> }) {
    const replies: Array<{ url: string; body: Record<string, unknown>; ts: string; sig: string }> = [];
    let resolveReplyDone: (() => void) | undefined;
    const replyDone = new Promise<void>((r) => { resolveReplyDone = r; });
    const port = 19_444 + Math.floor(Math.random() * 1000);
    const cfg = {
      enabled: true, port, bindAddress: "127.0.0.1", path: "/managed/chat",
      chatSecret: KAT_SECRET, gatewayReplyUrl: "https://gateway.test/api/chat/reply",
    };
    server = new ManagedChatServer(cfg, {
      respond: opts?.respond ?? (async () => "the answer"),
      log: { info: () => {}, warn: () => {}, error: () => {} },
      fetchFn: (async (url: string | URL | Request, init?: RequestInit) => {
        const headers = init?.headers as Record<string, string>;
        replies.push({
          url: String(url),
          body: JSON.parse(String(init?.body)) as Record<string, unknown>,
          ts: headers["x-standin-timestamp"],
          sig: headers["x-standin-signature"],
        });
        if (replies.length >= 2) resolveReplyDone?.(); // typing + message
        return new Response("{}", { status: 200 });
      }) as typeof fetch,
    });
    await server.start();
    return { port, replies, replyDone };
  }

  function post(port: number, body: string, sign = true, path = "/managed/chat") {
    const { timestamp, signature } = signBridge(KAT_SECRET, body);
    return fetch(`http://127.0.0.1:${port}${path}`, {
      method: "POST",
      headers: {
        "content-type": "application/json",
        ...(sign ? { "x-standin-timestamp": timestamp, "x-standin-signature": signature } : {}),
      },
      body,
    });
  }

  const inbound = JSON.stringify({
    tenantId: "t1", conversationId: "c1", activityId: "a-e2e", scope: "personal",
    sender: { displayName: "Alaa" }, text: "hi",
  });

  it("ACKs a signed message, then posts typing + the reply back with the SAME chat key", async () => {
    const { port, replies, replyDone } = await startServer();
    const res = await post(port, inbound);
    expect(res.status).toBe(200);
    await replyDone;

    expect(replies.map((r) => r.body.kind)).toEqual(["typing", "message"]);
    const reply = replies[1];
    expect(reply.url).toBe("https://gateway.test/api/chat/reply");
    expect(reply.body.tenantId).toBe("t1");
    expect(reply.body.text).toBe("the answer");
    // The reply is signed with the same chat key, verifiable by the gateway's construction.
    expect(verifyBridge(KAT_SECRET, reply.ts, JSON.stringify(reply.body), reply.sig, Number(reply.ts))).toBe(true);
  });

  it("rejects unsigned and mis-signed requests without consulting the agent", async () => {
    let consulted = 0;
    const { port } = await startServer({ respond: async () => { consulted++; return "x"; } });
    expect((await post(port, inbound, false)).status).toBe(401);
    const bad = await fetch(`http://127.0.0.1:${port}/managed/chat`, {
      method: "POST",
      headers: { "x-standin-timestamp": String(Date.now()), "x-standin-signature": "deadbeef" },
      body: inbound,
    });
    expect(bad.status).toBe(401);
    expect(consulted).toBe(0);
  });

  it("a redelivered activity ACKs but does not run a second agent turn", async () => {
    let consulted = 0;
    const { port } = await startServer({ respond: async () => { consulted++; return "x"; } });
    expect((await post(port, inbound)).status).toBe(200);
    expect((await post(port, inbound)).status).toBe(200);
    await new Promise((r) => setTimeout(r, 50));
    expect(consulted).toBe(1);
  });

  it("404s other paths", async () => {
    const { port } = await startServer();
    expect((await post(port, inbound, true, "/other")).status).toBe(404);
  });
});
