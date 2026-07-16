export function describeMsteamsVideoFrameOwner(frame) {
    if (!frame.participantName) {
        return undefined;
    }
    const kind = frame.source === "screenshare" ? "shared screen" : "camera";
    return `${frame.participantName}'s ${kind}`;
}
