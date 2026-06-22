import { describe, expect, it, vi } from "vitest";
import type { MsteamsVideoFrame } from "./msteams-video-frame.js";
import { VisionBudget } from "./vision-budget.js";
import {
  type ConsultImage,
  collectLatestFrameImages,
  frameToConsultImage,
  pushOrQueueBridgeImage,
} from "./vision-consult.js";

function frame(source: "camera" | "screenshare", data: string): MsteamsVideoFrame {
  return { source, dataBase64: data, mime: "image/jpeg", width: 100, height: 100, ts: 0 };
}

describe("frameToConsultImage", () => {
  it("maps a frame to the consult image input shape", () => {
    expect(frameToConsultImage({ dataBase64: "AAA", mime: "image/png" })).toEqual({
      type: "image",
      data: "AAA",
      mimeType: "image/png",
    });
  });
});

describe("collectLatestFrameImages", () => {
  it("gathers screen-share + camera and honors the vision budget", () => {
    const getLatestFrame = (s?: "camera" | "screenshare") =>
      s === "camera" ? frame("camera", "CAM") : frame("screenshare", "SCREEN");
    const images = collectLatestFrameImages({ getLatestFrame, callId: "c1" });
    // screen-share first, then camera (matches the realtime push order).
    expect(images.map((i) => i.data)).toEqual(["SCREEN", "CAM"]);
  });

  it("stops at the budget cap", () => {
    const budget = new VisionBudget(1); // only one frame per minute
    const getLatestFrame = (s?: "camera" | "screenshare") =>
      s === "camera" ? frame("camera", "CAM") : frame("screenshare", "SCREEN");
    const images = collectLatestFrameImages({
      getLatestFrame,
      visionBudget: budget,
      callId: "c1",
      now: () => 0,
    });
    expect(images.map((i) => i.data)).toEqual(["SCREEN"]); // camera dropped by budget
  });

  it("returns nothing when there is no frame source", () => {
    expect(collectLatestFrameImages({ callId: "c1" })).toEqual([]);
  });
});

describe("pushOrQueueBridgeImage", () => {
  const img = { dataBase64: "IMG", mime: "image/jpeg", text: "live screen-share" };

  it("(a) pushes via bridge.sendImage when the bridge supports it", () => {
    const sendImage = vi.fn();
    const queue: ConsultImage[] = [];
    const result = pushOrQueueBridgeImage({ sendImage }, img, queue);
    expect(result).toBe("pushed");
    expect(sendImage).toHaveBeenCalledWith(img);
    expect(queue).toHaveLength(0); // not queued when pushed live
  });

  it("(b) queues a consult image when the bridge lacks sendImage (stock openclaw)", () => {
    const queue: ConsultImage[] = [];
    const result = pushOrQueueBridgeImage({}, img, queue);
    expect(result).toBe("queued");
    expect(queue).toEqual([{ type: "image", data: "IMG", mimeType: "image/jpeg" }]);
  });

  it("propagates a throwing sendImage so the caller can refund the budget", () => {
    const sendImage = vi.fn(() => {
      throw new Error("provider rejected the image");
    });
    const queue: ConsultImage[] = [];
    expect(() => pushOrQueueBridgeImage({ sendImage }, img, queue)).toThrow();
    expect(queue).toHaveLength(0);
  });
});
