import { resolveRealtimeVoiceFastContextConsult, } from "openclaw/plugin-sdk/realtime-voice";
// Voice-call labels for the SDK realtime fast-context resolver.
/** Resolve fast-context consult data using caller-oriented labels. */
export async function resolveRealtimeFastContextConsult(params) {
    return await resolveRealtimeVoiceFastContextConsult({
        ...params,
        labels: {
            audienceLabel: "caller",
            contextName: "OpenClaw memory or session context",
        },
    });
}
