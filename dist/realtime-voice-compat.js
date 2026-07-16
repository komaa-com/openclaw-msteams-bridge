export function sendBridgeImage(bridge, image) {
    bridge.sendImage?.(image);
}
export function consultMediaPaths(result) {
    return result.mediaPaths ?? [];
}
