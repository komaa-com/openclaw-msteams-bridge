import { describeMsteamsVideoFrameOwner } from "./msteams-video-frame.js";
export function frameToConsultImage(frame) {
    return { type: "image", data: frame.dataBase64, mimeType: frame.mime };
}
export function collectLatestFrameImages(opts) {
    const { ambientVision, getLatestFrame, visionBudget, callId } = opts;
    if (!ambientVision || !getLatestFrame)
        return { images: [], owners: [] };
    const now = opts.now ?? (() => Date.now());
    const images = [];
    const owners = [];
    for (const source of ["screenshare", "camera"]) {
        const frame = getLatestFrame(source);
        if (!frame)
            continue;
        if (visionBudget && !visionBudget.tryConsume(callId, now()))
            break;
        images.push(frameToConsultImage(frame));
        owners.push(describeMsteamsVideoFrameOwner(frame) ?? (source === "screenshare" ? "a shared screen" : "a camera"));
    }
    return { images, owners };
}
export function withConsultImages(agentRuntime, images) {
    if (!images?.length)
        return agentRuntime;
    return {
        ...agentRuntime,
        runEmbeddedAgent: (p) => agentRuntime.runEmbeddedAgent(p && typeof p === "object" && !p.images ? { ...p, images } : p),
    };
}
export const MAX_QUEUED_AMBIENT_IMAGES = 6;
export function pushOrQueueBridgeImage(bridge, image, queue) {
    const fn = bridge.sendImage;
    if (typeof fn === "function") {
        fn.call(bridge, image);
        return "pushed";
    }
    queue.push({ type: "image", data: image.dataBase64, mimeType: image.mime });
    if (queue.length > MAX_QUEUED_AMBIENT_IMAGES) {
        queue.splice(0, queue.length - MAX_QUEUED_AMBIENT_IMAGES);
    }
    return "queued";
}
