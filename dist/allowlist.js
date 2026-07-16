export function normalizePhoneNumber(input) {
    if (!input) {
        return "";
    }
    return input.replace(/\D/g, "");
}
export function isAllowlistedCaller(from, allowFrom) {
    const raw = from?.trim();
    if (!raw) {
        return false;
    }
    const idFrom = raw.toLowerCase();
    const normalizedFrom = normalizePhoneNumber(raw);
    return (allowFrom ?? []).some((entry) => {
        const trimmed = entry.trim();
        if (!trimmed) {
            return false;
        }
        if (trimmed.toLowerCase() === idFrom) {
            return true;
        }
        const normalizedAllow = normalizePhoneNumber(trimmed);
        return normalizedAllow !== "" && normalizedFrom !== "" && normalizedAllow === normalizedFrom;
    });
}
export function isInboundCallAllowed(inboundPolicy, allowFrom, from) {
    switch (inboundPolicy) {
        case "open":
            return true;
        case "allowlist":
        case "pairing":
            return isAllowlistedCaller(from, allowFrom);
        default:
            return false;
    }
}
export function describeInboundRejection(inboundPolicy, from) {
    const policy = inboundPolicy ?? "disabled";
    const caller = from?.trim() ? `caller "${from.trim()}"` : "caller with no caller id";
    if (policy === "disabled" || policy === "open") {
        return `inbound call rejected by policy "${policy}": inbound calling is ${policy === "disabled" ? "disabled - set inboundPolicy to \"allowlist\" and add callers to allowFrom to accept calls" : "open"} (${caller})`;
    }
    const pairingHint = policy === "pairing"
        ? ' Note: "pairing" currently enforces a plain allowlist (this plugin issues no pairing codes or approvals for calls) - add the caller\'s AAD object id to allowFrom, or approve them via your gateway\'s pairing flow if your host supports one.'
        : " Add the caller's AAD object id (or phone number) to allowFrom to accept them.";
    return `inbound call rejected by policy "${policy}": ${caller} is not in allowFrom.${pairingHint}`;
}
