export const GROUP_CALL_GATE_DEFAULTS = {
    requireAddress: true,
    wakePhrases: ["assistant"],
    followUpWindowMs: 12_000,
};
export function resolveGroupCallGateConfig(raw) {
    return {
        requireAddress: raw?.requireAddress ?? GROUP_CALL_GATE_DEFAULTS.requireAddress,
        wakePhrases: raw?.wakePhrases ?? GROUP_CALL_GATE_DEFAULTS.wakePhrases,
        followUpWindowMs: raw?.followUpWindowMs ?? GROUP_CALL_GATE_DEFAULTS.followUpWindowMs,
    };
}
export function isAddressed(transcript, wakePhrases) {
    const text = transcript.toLowerCase();
    for (const phrase of wakePhrases) {
        const needle = phrase.trim().toLowerCase();
        if (!needle) {
            continue;
        }
        let from = 0;
        for (;;) {
            const at = text.indexOf(needle, from);
            if (at < 0) {
                break;
            }
            const before = at === 0 ? "" : text[at - 1];
            const after = at + needle.length >= text.length ? "" : text[at + needle.length];
            if (!isWordChar(before) && !isWordChar(after)) {
                return true;
            }
            from = at + needle.length;
        }
    }
    return false;
}
function isWordChar(ch) {
    return ch.length > 0 && /[\p{L}\p{N}]/u.test(ch);
}
export function shouldRespondToGroupTurn(params) {
    const { transcript, isGroup, config, lastAddressedAt, now } = params;
    const addressed = isAddressed(transcript, config.wakePhrases);
    const gateActive = isGroup && config.requireAddress && config.wakePhrases.some((p) => p.trim().length > 0);
    if (!gateActive) {
        return { respond: true, addressed, gated: false };
    }
    if (addressed) {
        return { respond: true, addressed: true, gated: true };
    }
    const inFollowUp = config.followUpWindowMs > 0 &&
        lastAddressedAt !== undefined &&
        now - lastAddressedAt <= config.followUpWindowMs;
    return { respond: inFollowUp, addressed: false, gated: true };
}
