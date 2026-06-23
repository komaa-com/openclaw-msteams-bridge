// Core call types for the self-contained Teams voice plugin.
// (CallRecord mirrors the shape voice-call persisted; trimmed to what a Teams realtime plugin needs.)

export type CallState =
  | "initiated"
  | "ringing"
  | "answered"
  | "active"
  | "completed"
  | "failed";

export type CallDirection = "inbound" | "outbound";

export type CallEndReason =
  | "completed"
  | "hangup-user"
  | "hangup-agent"
  | "no-answer"
  | "timeout"
  | "error";

export interface TranscriptEntry {
  role: "caller" | "bot";
  text: string;
  at: number;
}

export interface CallRecord {
  /** Internal id we mint. */
  callId: string;
  /** Teams/Graph call id (the worker/provider call id). */
  providerCallId: string;
  direction: CallDirection;
  state: CallState;
  from: string;
  to: string;
  startedAt: number;
  answeredAt?: number;
  endedAt?: number;
  endReason?: CallEndReason;
  transcript: TranscriptEntry[];
  /** Optional notify message delivered on answer (outbound call-backs). */
  message?: string;
}

export const TERMINAL_STATES: ReadonlySet<CallState> = new Set(["completed", "failed"]);
