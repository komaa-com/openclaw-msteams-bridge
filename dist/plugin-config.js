import { resolveManagedChatConfig } from "./managed-chat.js";
const str = (v) => (typeof v === "string" ? v : "");
const asObject = (v) => v && typeof v === "object" && !Array.isArray(v) ? v : undefined;
export function resolvePluginConfig(rawInput) {
    const c = rawInput ?? {};
    const r = c.realtime ?? {};
    return {
        enabled: c.enabled !== false,
        media: {
            port: Number(c.callingPort ?? c.port ?? 9442),
            bindAddress: c.bindAddress,
            path: String(c.path ?? "/msteams/calling"),
            sharedSecret: str(c.secret),
        },
        managedChat: resolveManagedChatConfig({
            ...(str(c.bindAddress) ? { bindAddress: str(c.bindAddress) } : {}),
            ...(asObject(c.managedBot) ?? {}),
            ...(str(c.messagesBindAddress) ? { bindAddress: str(c.messagesBindAddress) } : {}),
            ...(str(c.secret) ? { chatSecret: str(c.secret) } : {}),
            ...(c.messagesPort !== undefined ? { port: Number(c.messagesPort) } : {}),
            ...(str(c.messagesPath) ? { path: str(c.messagesPath) } : {}),
            ...(str(c.gatewayReplyUrl) ? { gatewayReplyUrl: str(c.gatewayReplyUrl) } : {}),
        }),
        outbound: c.outbound,
        limits: {
            maxConcurrentCalls: Number(c.maxConcurrentCalls ?? 4),
            maxDurationMs: Number(c.maxDurationSeconds ?? 0) * 1000,
            staleCallReaperMs: Number(c.staleCallReaperSeconds ?? 120) * 1000,
        },
        voice: {
            agentId: c.agentId,
            sessionScope: c.sessionScope,
            responseModel: c.responseModel,
            responseTimeoutMs: Number(c.responseTimeoutMs ?? 30000),
            inboundPolicy: c.inboundPolicy,
            allowFrom: c.allowFrom,
            inboundGreeting: c.inboundGreeting,
            mode: c.mode,
            realtime: {
                provider: r.provider,
                providers: r.providers,
                instructions: r.instructions,
                toolPolicy: r.toolPolicy ?? "none",
                consultPolicy: r.consultPolicy,
                consultThinkingLevel: r.consultThinkingLevel,
                consultFastMode: r.consultFastMode,
                suppressInputDuringPlayback: r.suppressInputDuringPlayback,
                echoSuppressionWindowMs: r.echoSuppressionWindowMs,
                echoBargeInRms: r.echoBargeInRms,
                fastContext: (r.fastContext ?? {
                    enabled: false,
                    timeoutMs: 800,
                    maxResults: 3,
                    sources: ["memory", "sessions"],
                    fallbackToConsult: false,
                }),
            },
            stt: c.stt,
            msteams: {
                requireRecordingStatus: c.requireRecordingStatus,
                groupCall: c.groupCall,
                maxVisionPerMinute: c.maxVisionPerMinute,
                meetingRecap: c.meetingRecap,
                bilingual: c.bilingual,
            },
            tts: c.tts,
        },
    };
}
