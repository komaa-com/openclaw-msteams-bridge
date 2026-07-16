import { resolveRealtimeVoiceFastContextConsult, } from "openclaw/plugin-sdk/realtime-voice";
export async function resolveRealtimeFastContextConsult(params) {
    return await resolveRealtimeVoiceFastContextConsult({
        ...params,
        labels: {
            audienceLabel: "caller",
            contextName: "OpenClaw memory or session context",
        },
    });
}
