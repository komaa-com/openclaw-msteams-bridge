/**
 * Human description of whose camera/screen a frame shows — e.g. `"Alice's shared screen"` or
 * `"Bob's camera"` — for attributing an attached image to a person. Returns undefined when the
 * participant is unknown (1:1, guest/anonymous, or an older worker). Shared by the streaming
 * attach (provider) and the realtime `look_at_screen` surface so the wording stays consistent.
 */
export function describeMsteamsVideoFrameOwner(frame) {
    if (!frame.participantName) {
        return undefined;
    }
    const kind = frame.source === "screenshare" ? "shared screen" : "camera";
    return `${frame.participantName}'s ${kind}`;
}
