export function resolveVoiceResponseModel(params) {
    const configured = params.cfg?.agents?.defaults?.model;
    const configuredPrimary = (typeof configured === "string" ? configured : configured?.primary)?.trim();
    const modelRef = params.voiceConfig.responseModel ??
        (configuredPrimary || `${params.agentRuntime.defaults.provider}/${params.agentRuntime.defaults.model}`);
    const slashIndex = modelRef.indexOf("/");
    return {
        modelRef,
        provider: slashIndex === -1 ? params.agentRuntime.defaults.provider : modelRef.slice(0, slashIndex),
        model: slashIndex === -1 ? modelRef : modelRef.slice(slashIndex + 1),
    };
}
