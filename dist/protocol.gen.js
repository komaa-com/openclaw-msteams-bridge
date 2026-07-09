// GENERATED from protocol/schema.yaml - do not hand-edit.
// Regenerate with: python3 protocol/generate.py (in the OpenClawBridge repo).
/**
 * Wire protocol between the OpenClawBridge media worker and OpenClaw's voice-call msteams provider.
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
/**
 * 'inbound' (caller dialed the bot) or 'outbound' (the bot placed this call, e.g. via
 * /api/calls/place). Receivers default to inbound when absent.
 */
export const CallDirectionSchema = z.enum(["inbound", "outbound"]);
/**
 * Which inbound video stream a sampled frame came from.
 */
export const VideoFrameSourceSchema = z.enum(["camera", "screenshare"]);
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
