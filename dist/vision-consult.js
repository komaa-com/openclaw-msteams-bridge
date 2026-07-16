export function frameToConsultImage(frame) {
    return { type: "image", data: frame.dataBase64, mimeType: frame.mime };
}
export function collectLatestFrameImages(opts) {
    const { getLatestFrame, visionBudget, callId } = opts;
    if (!getLatestFrame)
        return [];
    const now = opts.now ?? (() => Date.now());
    const images = [];
    for (const source of ["screenshare", "camera"]) {
        const frame = getLatestFrame(source);
        if (!frame)
            continue;
        if (visionBudget && !visionBudget.tryConsume(callId, now()))
            break;
        images.push(frameToConsultImage(frame));
    }
    return images;
}
export function pushOrQueueBridgeImage(bridge, image, queue) {
    const fn = bridge.sendImage;
    if (typeof fn === "function") {
        fn.call(bridge, image);
        return "pushed";
    }
    queue.push({ type: "image", data: image.dataBase64, mimeType: image.mime });
    return "queued";
}
