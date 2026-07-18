import { describe, expect, it, vi } from "vitest";
import { heartbeatSweep, type HeartbeatSocket } from "./msteams-media-stream.js";

function fakeSocket(isAlive: boolean | undefined): HeartbeatSocket & {
  terminated: boolean;
  pinged: number;
} {
  return {
    isAlive,
    terminated: false,
    pinged: 0,
    terminate() {
      this.terminated = true;
    },
    ping() {
      this.pinged += 1;
    },
  };
}

describe("heartbeatSweep", () => {
  it("terminates a socket that missed the previous ping (isAlive=false) — frees its concurrency slot", () => {
    const dead = fakeSocket(false);
    heartbeatSweep([dead]);
    expect(dead.terminated).toBe(true);
    expect(dead.pinged).toBe(0); // not pinged; it is torn down
  });

  it("pings a live socket and marks it not-alive until its pong flips it back", () => {
    const alive = fakeSocket(true);
    heartbeatSweep([alive]);
    expect(alive.terminated).toBe(false);
    expect(alive.pinged).toBe(1);
    expect(alive.isAlive).toBe(false); // next sweep terminates it unless a pong arrives first
  });

  it("a socket that pongs between sweeps (isAlive back to true) survives", () => {
    const s = fakeSocket(true);
    heartbeatSweep([s]); // isAlive -> false, ping sent
    s.isAlive = true; // simulate the pong handler
    heartbeatSweep([s]);
    expect(s.terminated).toBe(false);
    expect(s.pinged).toBe(2);
  });

  it("a ping that throws does not abort the sweep of the other clients", () => {
    const bad = fakeSocket(true);
    bad.ping = () => {
      throw new Error("socket closed");
    };
    const good = fakeSocket(true);
    expect(() => heartbeatSweep([bad, good])).not.toThrow();
    expect(good.pinged).toBe(1);
  });
});
