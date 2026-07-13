// GENERATED from the shared wire-protocol schema - do not hand-edit.
// The wire protocol the StandIn media bridge speaks with this plugin.
/**
 * Wire protocol between the StandIn media bridge and this OpenClaw msteams voice plugin.
 * One WebSocket per active Teams call. JSON-encoded discriminated union on the 'type' field. Zod
 * schemas validate the inbound (worker -> plugin) direction; the outbound (plugin -> worker)
 * message shapes are exported as types.
 */
import { z } from "zod";

/**
 * The Teams bridge wire format is PCM 16 kHz, 16-bit, mono in both directions. Single source of
 * truth for the sample rate shared by the provider, realtime bridge, and TTS adapter.
 */
export const MSTEAMS_PCM_SAMPLE_RATE_HZ = 16_000;

/**
 * Microsoft Teams recording status for the call. The bot must have called Graph
 * updateRecordingStatus before media-derived data may be persisted; receivers gate
 * persistence/processing on 'active'.
 */
export const RecordingStatusSchema = z.enum(["active", "inactive", "unknown"]);
export type MsteamsRecordingStatus = z.infer<typeof RecordingStatusSchema>;

/**
 * 'inbound' (caller dialed the bot) or 'outbound' (the bot placed this call, e.g. via
 * /api/calls/place). Receivers default to inbound when absent.
 */
export const CallDirectionSchema = z.enum(["inbound", "outbound"]);
export type CallDirection = z.infer<typeof CallDirectionSchema>;

/**
 * Which inbound video stream a sampled frame came from.
 */
export const VideoFrameSourceSchema = z.enum(["camera", "screenshare"]);
export type VideoFrameSource = z.infer<typeof VideoFrameSourceSchema>;

// ---- Inbound messages (worker -> plugin), validated with Zod ----------------

/**
 * Sent once after WebSocket open. Carries call metadata + caller identity. recordingStatus is the
 * Teams recording state at answer time; omitted when not yet known, in which case the worker
 * reports it later via recording.status.
 */
export const SessionStartSchema = z.object({
  type: z.literal("session.start"),
  /**
   * Graph call id; must match the HMAC-authenticated callId in the WebSocket URL path.
   */
  callId: z.string().min(1),
  /**
   * Teams chat thread id for the call.
   */
  threadId: z.string().min(1),
  /**
   * Caller identity (best-effort; see CallerInfo).
   */
  caller: z.object({
    aadId: z.string().nullable().optional(),
    displayName: z.string().nullable().optional(),
    tenantId: z.string().nullable().optional(),
  }),
  /**
   * Teams recording status at answer time; omitted (null) when not yet known.
   */
  recordingStatus: RecordingStatusSchema.optional(),
  /**
   * Lets the plugin correlate calls it initiated. Defaults to inbound when absent.
   */
  direction: CallDirectionSchema.optional(),
});

/**
 * Sent right before close. Reason is opaque text (e.g. 'call-ended', 'transport-failure').
 */
export const SessionEndSchema = z.object({
  type: z.literal("session.end"),
  /**
   * Opaque close reason text.
   */
  reason: z.string(),
});

/**
 * Worker -> plugin. Reports the Microsoft Teams recording status for the call. The plugin gates
 * persistence/processing of media-derived data on 'active' (Media Access API: the bot must have
 * called Graph updateRecordingStatus before media-derived data may be persisted).
 */
export const RecordingStatusMessageSchema = z.object({
  type: z.literal("recording.status"),
  /**
   * Current Teams recording status.
   */
  status: RecordingStatusSchema,
});

/**
 * Bidirectional. Sender increments seq; receiver may use it to detect drops. payloadBase64 is the
 * base64-encoded PCM16K (16kHz, 16-bit, mono, little-endian) audio bytes, usually one 20ms frame =
 * 640 bytes raw = ~856 chars base64.
 */
export const AudioFrameSchema = z.object({
  type: z.literal("audio.frame"),
  /**
   * Monotonic frame sequence number (per direction).
   */
  seq: z.number().int().nonnegative(),
  /**
   * Frame timestamp in ms on the sender's audio timeline.
   */
  timestampMs: z.number().int().nonnegative(),
  /**
   * Base64-encoded PCM16K bytes.
   */
  payloadBase64: z.string(),
  /**
   * Active speaker's display name when unmixed audio is on (additive; absent on the mixed path), so
   * transcripts can carry real per-person attribution. Worker -> plugin only.
   */
  speakerName: z.string().optional(),
});

/**
 * Worker -> plugin. A sampled inbound video frame (caller camera or screen-share) as a base64 JPEG.
 * Emitted sparsely, best-effort (dropped if the socket is busy). The plugin buffers the latest
 * frame per source and feeds it to a vision model on demand.
 */
export const VideoFrameSchema = z.object({
  type: z.literal("video.frame"),
  /**
   * 'camera' | 'screenshare'.
   */
  source: VideoFrameSourceSchema,
  /**
   * Capture timestamp in ms.
   */
  ts: z.number().int().nonnegative(),
  /**
   * Frame width in pixels (already downscaled worker-side).
   */
  width: z.number().int().positive(),
  /**
   * Frame height in pixels.
   */
  height: z.number().int().positive(),
  /**
   * Image MIME type; the worker sends 'image/jpeg'.
   */
  mime: z.string().min(1),
  /**
   * Base64-encoded image bytes.
   */
  dataBase64: z.string().min(1),
  /**
   * Who this frame belongs to, so a vision model can reason about 'who is saying what' in a group
   * call. For 'camera' the currently-subscribed speaker; for 'screenshare' the sharer. Best-effort:
   * absent for anonymous/guest participants or when identity is unavailable.
   */
  participantId: z.string().min(1).optional(),
  /**
   * Display name matching participantId. Best-effort.
   */
  participantName: z.string().min(1).optional(),
});

/**
 * Worker -> plugin. The number of human participants on the call (excludes the bot), sent at join
 * and whenever the roster changes. Lets the plugin tell a 1:1 call (count <= 1) from a
 * group/meeting (count >= 2) so it can stay quiet until addressed in groups.
 */
export const ParticipantsSchema = z.object({
  type: z.literal("participants"),
  /**
   * Human participant count, excluding the bot.
   */
  count: z.number().int().nonnegative(),
});

/**
 * Worker -> plugin. A DTMF key the caller pressed during the call (in-band tone detected by the
 * media platform). The plugin surfaces it to the agent so it can drive IVR-style flows ('press 1
 * to...').
 */
export const DtmfSchema = z.object({
  type: z.literal("dtmf"),
  /**
   * A single character: '0'-'9', '*', or '#'.
   */
  digit: z.string().regex(/^[0-9*#]$/),
});

/**
 * Worker -> plugin heartbeat (every 30s; keeps proxy/NAT idle timeouts from half-opening the
 * socket). The plugin replies with pong.
 */
export const PingSchema = z.object({
  type: z.literal("ping"),
  /**
   * Sender's unix epoch ms.
   */
  ts: z.number().int().nonnegative(),
});

/**
 * Worker -> plugin (H4). Ask the agent to speak the given text in its OWN realtime voice before the
 * worker tears the call down - e.g. a brief goodbye when a time/minute limit (free trial, daily
 * budget, or the paid max-minutes governor) cuts the call off, instead of a bare mid-sentence
 * hangup. The plugin injects the text into the realtime session and triggers a spoken response; in
 * streaming/TTS mode it speaks via the existing TTS path, else it no-ops. Additive/best-effort: an
 * older plugin ignores it.
 */
export const AssistantSaySchema = z.object({
  type: z.literal("assistant.say"),
  /**
   * The exact utterance the agent should speak (a short goodbye).
   */
  text: z.string().min(1),
});

/**
 * Every message the worker can send, discriminated on 'type'.
 */
export const InboundMessageSchema = z.discriminatedUnion("type", [
  SessionStartSchema,
  SessionEndSchema,
  RecordingStatusMessageSchema,
  AudioFrameSchema,
  VideoFrameSchema,
  ParticipantsSchema,
  DtmfSchema,
  PingSchema,
  AssistantSaySchema,
]);

export type InboundMessage = z.infer<typeof InboundMessageSchema>;

// ---- Outbound messages (plugin -> worker), plain types ----------------------

/**
 * One viseme mark: tMs milliseconds from utterance start, visemeId an Azure viseme id (0-21).
 */
export interface SpeechMark {
  tMs: number;
  visemeId: number;
}

/**
 * Bidirectional. Sender increments seq; receiver may use it to detect drops. payloadBase64 is the
 * base64-encoded PCM16K (16kHz, 16-bit, mono, little-endian) audio bytes, usually one 20ms frame =
 * 640 bytes raw = ~856 chars base64.
 */
export interface AudioFrameMessage {
  type: "audio.frame";
  /**
   * Monotonic frame sequence number (per direction).
   */
  seq: number;
  /**
   * Frame timestamp in ms on the sender's audio timeline.
   */
  timestampMs: number;
  /**
   * Base64-encoded PCM16K bytes.
   */
  payloadBase64: string;
  /**
   * Active speaker's display name when unmixed audio is on (additive; absent on the mixed path), so
   * transcripts can carry real per-person attribution. Worker -> plugin only.
   */
  speakerName?: string;
}

/**
 * Plugin -> worker. Cancel the current TTS turn (barge-in). The worker flushes queued playback for
 * audio received before this cancel.
 */
export interface AssistantCancelMessage {
  type: "assistant.cancel";
  /**
   * The assistant turn being cancelled (sender-scoped counter).
   */
  turnId: number;
}

/**
 * Plugin -> worker (CVI Phase 6b). Hint the emotion the avatar should express on its video tile.
 * Additive and best-effort: an older worker ignores it; unknown values fall back to neutral.
 * Affects only the rendered face, never audio.
 */
export interface ExpressionMessage {
  type: "expression";
  /**
   * Emotion tag; see Emotion enum for well-known values.
   */
  emotion: string;
}

/**
 * Plugin -> worker (CVI Phase 5). Viseme timing for one TTS utterance. The worker maps tMs onto the
 * same baseTicks timeline the audio uses and shapes the avatar mouth per viseme, blended with the
 * RMS openness. Additive/best-effort: an older worker ignores it, and the avatar falls back to
 * RMS-only lip-sync.
 */
export interface SpeechMarksMessage {
  type: "speech.marks";
  /**
   * Reserved timeline anchor; senders currently send 0.
   */
  ts: number;
  /**
   * Viseme marks, ascending tMs.
   */
  marks: SpeechMark[];
}

/**
 * Plugin -> worker (CVI Phase 8, assistant visual sharing). Show the caller an image on the bot's
 * outbound video tile for a few seconds, then return to the avatar. Additive/best-effort: an older
 * worker ignores it; the avatar keeps rendering.
 */
export interface DisplayImageMessage {
  type: "display.image";
  /**
   * Image bytes (JPEG/PNG per mime), base64-encoded.
   */
  dataBase64: string;
  /**
   * 'image/jpeg' or 'image/png'.
   */
  mime: string;
  /**
   * How long to show the image; the worker applies its default when omitted/null or <= 0.
   */
  durationMs?: number | null;
  /**
   * 'fullscreen' (replace the tile; worker default) or 'overlay' (PiP inset).
   */
  mode?: string | null;
  /**
   * Reserved timeline anchor. The TS sender omits it, the Python sender sends 0; the worker
   * defaults it to 0.
   */
  ts?: number;
  /**
   * Optional short label drawn along the bottom of the image.
   */
  caption?: string | null;
}

/**
 * Plugin -> worker (EXPERIMENTAL, avatar video relay). One frame of a continuous bot-tile video
 * stream, JPEG-compressed. The worker decodes it, scales it to its tile, and shows the LATEST
 * frame; if no new frame arrives within its stall window (~750 ms) it reverts to the rendered
 * avatar. No lifecycle handshake: the first frames start the takeover, silence ends it.
 * Latest-wins: the worker keeps no queue, and senders MUST drop (not buffer) frames under
 * backpressure, like hot-path audio. Additive/best-effort: an older worker ignores it and keeps
 * rendering its avatar.
 */
export interface DisplayFrameMessage {
  type: "display.frame";
  /**
   * Monotonic frame sequence number; the worker drops frames older than the newest seen.
   */
  seq: number;
  /**
   * Capture timestamp in ms on the SENDER's media timeline - the same timeline the sender stamps on
   * its outbound audio.frame messages. Used for A/V skew measurement and the worker's max-fps
   * ceiling, not for scheduling.
   */
  ts: number;
  /**
   * Frame encoding; senders send 'image/jpeg'. PNG is legal to decode but too large at rate.
   */
  mime: string;
  /**
   * Base64 image bytes. Senders should target <= 40 KB raw per frame; the worker rejects frames
   * above 256 KB base64.
   */
  dataBase64: string;
  /**
   * Source pixel width (informational; the worker scales to its tile regardless).
   */
  width?: number | null;
  /**
   * Source pixel height (informational).
   */
  height?: number | null;
}

/**
 * Plugin -> worker heartbeat response, echoing the ping timestamp.
 */
export interface PongMessage {
  type: "pong";
  /**
   * Echoed ping timestamp (unix epoch ms).
   */
  ts: number;
}

/**
 * Every message a plugin may send to the worker.
 */
export type OutboundMessage =
  | AudioFrameMessage
  | AssistantCancelMessage
  | ExpressionMessage
  | SpeechMarksMessage
  | DisplayImageMessage
  | DisplayFrameMessage
  | PongMessage;
