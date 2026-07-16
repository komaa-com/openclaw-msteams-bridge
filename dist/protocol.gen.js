import { z } from "zod";
export const MSTEAMS_PCM_SAMPLE_RATE_HZ = 16_000;
export const RecordingStatusSchema = z.enum(["active", "inactive", "unknown"]);
export const CallDirectionSchema = z.enum(["inbound", "outbound"]);
export const VideoFrameSourceSchema = z.enum(["camera", "screenshare"]);
export const SessionStartSchema = z.object({
    type: z.literal("session.start"),
    callId: z.string().min(1),
    threadId: z.string().min(1),
    caller: z.object({
        aadId: z.string().nullable().optional(),
        displayName: z.string().nullable().optional(),
        tenantId: z.string().nullable().optional(),
    }),
    recordingStatus: RecordingStatusSchema.optional(),
    direction: CallDirectionSchema.optional(),
});
export const SessionEndSchema = z.object({
    type: z.literal("session.end"),
    reason: z.string(),
});
export const RecordingStatusMessageSchema = z.object({
    type: z.literal("recording.status"),
    status: RecordingStatusSchema,
});
export const AudioFrameSchema = z.object({
    type: z.literal("audio.frame"),
    seq: z.number().int().nonnegative(),
    timestampMs: z.number().int().nonnegative(),
    payloadBase64: z.string(),
    speakerName: z.string().optional(),
});
export const VideoFrameSchema = z.object({
    type: z.literal("video.frame"),
    source: VideoFrameSourceSchema,
    ts: z.number().int().nonnegative(),
    width: z.number().int().positive(),
    height: z.number().int().positive(),
    mime: z.string().min(1),
    dataBase64: z.string().min(1),
    participantId: z.string().min(1).optional(),
    participantName: z.string().min(1).optional(),
});
export const ParticipantsSchema = z.object({
    type: z.literal("participants"),
    count: z.number().int().nonnegative(),
});
export const DtmfSchema = z.object({
    type: z.literal("dtmf"),
    digit: z.string().regex(/^[0-9*#]$/),
});
export const PingSchema = z.object({
    type: z.literal("ping"),
    ts: z.number().int().nonnegative(),
});
export const AssistantSaySchema = z.object({
    type: z.literal("assistant.say"),
    text: z.string().min(1),
});
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
