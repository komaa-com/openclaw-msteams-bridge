import { describe, expect, it, vi } from "vitest";
import type { MsteamsVideoFrame } from "./msteams-video-frame.js";
import { VisionBudget } from "./vision-budget.js";
import {
  type ConsultImage,
  collectLatestFrameImages,
  MAX_QUEUED_AMBIENT_IMAGES,
  pushOrQueueBridgeImage,
  type ConsultImage,
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
    const { images, owners } = collectLatestFrameImages({ getLatestFrame, callId: "c1" });
    // screen-share first, then camera (matches the realtime push order).
    expect(images.map((i) => i.data)).toEqual(["SCREEN", "CAM"]);
    // Owners run parallel to images so the caller can say whose screen it attached. Without a
    // participantName the label degrades to the source kind rather than vanishing.
    expect(owners).toHaveLength(2);
    expect(owners[0]).toMatch(/shared screen/);
    expect(owners[1]).toMatch(/camera/);
  });

  it("stops at the budget cap", () => {
    const budget = new VisionBudget(1); // only one frame per minute
    const getLatestFrame = (s?: "camera" | "screenshare") =>
      s === "camera" ? frame("camera", "CAM") : frame("screenshare", "SCREEN");
    const { images, owners } = collectLatestFrameImages({
      getLatestFrame,
      visionBudget: budget,
      callId: "c1",
      now: () => 0,
    });
    expect(images.map((i) => i.data)).toEqual(["SCREEN"]); // camera dropped by budget
    expect(owners).toHaveLength(1); // labels stay in step with the images they describe
  });

  it("returns nothing when there is no frame source", () => {
    expect(collectLatestFrameImages({ callId: "c1" })).toEqual({ images: [], owners: [] });
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

describe("ambient queue is bounded", () => {
  const img = (n: number) => ({ dataBase64: `IMG${n}`, mime: "image/jpeg", text: `f${n}` });

  it("keeps only the newest frames when the host cannot push live", () => {
    // The queue only drains when a consult runs. On a realtime call where the caller shares a changing
    // screen and never asks a question, nothing drains it - and at the default budget (30 frames/min,
    // 50-200 KB of base64 each) an hour-long meeting retained hundreds of megabytes PER CALL.
    const queue: ConsultImage[] = [];
    const hostWithoutSendImage = {};
    for (let n = 0; n < MAX_QUEUED_AMBIENT_IMAGES + 4; n++) {
      expect(pushOrQueueBridgeImage(hostWithoutSendImage, img(n), queue)).toBe("queued");
    }
    expect(queue).toHaveLength(MAX_QUEUED_AMBIENT_IMAGES);
    // Oldest dropped: ambient context is about what is on screen NOW.
    expect(queue.at(-1)!.data).toBe(`IMG${MAX_QUEUED_AMBIENT_IMAGES + 3}`);
    expect(queue.some((i) => i.data === "IMG0")).toBe(false);
  });

  it("does not queue at all when the host can push live", () => {
    const queue: ConsultImage[] = [];
    const pushed: unknown[] = [];
    const hostWithSendImage = { sendImage: (i: unknown) => pushed.push(i) };
    expect(pushOrQueueBridgeImage(hostWithSendImage, img(1), queue)).toBe("pushed");
    expect(queue).toHaveLength(0);
    expect(pushed).toHaveLength(1);
  });
});
