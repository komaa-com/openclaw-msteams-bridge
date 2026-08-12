import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { isFollowUpWindowOpen, shouldRespondToGroupTurn } from "./group-call-gate.js";

/**
 * Four of the five gaps this file guards had the same shape: the capability was fully built, and
 * nothing reached it.
 *
 *   - shouldRespondToGroupTurn had zero production callers; the behaviour was reimplemented inline in
 *     two files, so the unit tests were exercising code the product did not run.
 *   - greetingOnRecordingActive was honoured by the realtime session and set only by a test, so every
 *     call-back spoke into a still-ringing phone.
 *   - display.image's mode was hardcoded to "overlay", making the worker's fullscreen path
 *     unreachable.
 *   - openclaw_agent_consult was dispatched but never advertised, so the model was never told it could
 *     ask the agent anything.
 *
 * A behavioural test cannot catch "nobody calls this" - the code under test passes in isolation, which
 * is exactly why it stayed broken. So these assertions are on the WIRING, in the same spirit as the
 * media bridge's PathBSourceRelayTests. They are cheap and they fail the moment a call site is dropped.
 */

const SRC = new URL(".", import.meta.url).pathname;
const read = (f: string) => readFileSync(join(SRC, f), "utf8");
/** Strip line comments so prose naming a symbol cannot satisfy a wiring assertion. */
const code = (f: string) =>
  read(f)
    .split("\n")
    .filter((l) => !l.trimStart().startsWith("//") && !l.trimStart().startsWith("*"))
    .join("\n");

describe("the group-call gate is the one used in production", () => {
  it("is called by the streaming path rather than reimplemented", () => {
    expect(code("msteams-streaming.ts")).toContain("shouldRespondToGroupTurn({");
  });

  it("shares its follow-up window with the realtime audio egress", () => {
    // Realtime has no transcript at the egress, so it can only ask "is the window open" - but it must
    // ask the SAME question, from the same code, or the two drift.
    expect(code("msteams-realtime.ts")).toContain("isFollowUpWindowOpen({");
  });

  it("still decides correctly - 1:1 always answers, a group needs the wake phrase", () => {
    const config = { requireAddress: true, wakePhrases: ["assistant"], followUpWindowMs: 8000 };
    const now = 1_000_000;

    expect(
      shouldRespondToGroupTurn({ transcript: "what time is it", isGroup: false, config, lastAddressedAt: undefined, now }).respond,
    ).toBe(true);

    expect(
      shouldRespondToGroupTurn({ transcript: "what time is it", isGroup: true, config, lastAddressedAt: undefined, now }).respond,
    ).toBe(false);

    const hailed = shouldRespondToGroupTurn({ transcript: "assistant, what time is it", isGroup: true, config, lastAddressedAt: undefined, now });
    expect(hailed).toMatchObject({ respond: true, addressed: true });
  });

  it("keeps the follow-up window open only for its duration", () => {
    expect(isFollowUpWindowOpen({ lastAddressedAt: 1000, followUpWindowMs: 8000, now: 9000 })).toBe(true);
    expect(isFollowUpWindowOpen({ lastAddressedAt: 1000, followUpWindowMs: 8000, now: 9001 })).toBe(false);
    // A window of 0 is "no follow-up", not "always open".
    expect(isFollowUpWindowOpen({ lastAddressedAt: 1000, followUpWindowMs: 0, now: 1000 })).toBe(false);
    expect(isFollowUpWindowOpen({ lastAddressedAt: undefined, followUpWindowMs: 8000, now: 1000 })).toBe(false);
  });
});

describe("capabilities that exist are actually reachable", () => {
  it("offers the agent consult to the model", () => {
    // Dispatching a tool you never advertise is a menu nobody can order from.
    expect(code("msteams-realtime.ts")).toContain("REALTIME_VOICE_AGENT_CONSULT_TOOL]");
  });

  it("lets the model choose fullscreen instead of hardcoding the overlay", () => {
    const rt = code("msteams-realtime.ts");
    expect(rt).toContain("mode: displayMode");
    expect(rt).not.toContain('mode: "overlay",');
    // and the model has to be able to ASK for it
    expect(code("msteams-realtime-tools.ts")).toContain('"fullscreen"');
  });

  it("defers the outbound greeting until the callee answers", () => {
    const rt = code("msteams-runtime.ts");
    // The outbound branch passes the defer flag; buildDeps forwards it to the session.
    expect(rt).toContain("this.createCall(session, greeting, true)");
    expect(rt).toContain("greetingOnRecordingActive,");
  });

  it("keys chat sessions through sessionScope instead of a hardcoded string", () => {
    const rt = code("msteams-runtime.ts");
    expect(rt).toContain("this.chatSessionKey(message)");
    expect(rt).toContain("private chatSessionKey(");
  });
});

describe("speaker attribution survives the runtime boundary", () => {
  it("forwards the speaker name from audio frames instead of dropping it", () => {
    // The dead link: onAudioFrame forwarded ONLY i.payload, so setCurrentSpeaker - implemented on both
    // call paths, and used by both to prefix the transcript "Sara: ..." - had zero callers. The whole
    // mechanism was built and waiting on one dropped field.
    const rt = code("msteams-runtime.ts");
    expect(rt).toContain("call.setCurrentSpeaker(speaker)");
    expect(rt).toContain("i.speakerName");
  });

  it("only pushes the speaker when it changes", () => {
    // Audio arrives ~50x/second; the speaker changes at conversational pace. Pushing every frame would
    // be 50 redundant writes a second for one meaningful transition.
    expect(code("msteams-runtime.ts")).toContain("lastSpeakerByCall");
  });

  it("releases the speaker map with the call", () => {
    const rt = code("msteams-runtime.ts");
    expect(rt).toContain("this.lastSpeakerByCall.delete(callId)");
  });
});
