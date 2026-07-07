// Caller allowlist helpers for provider-normalized caller ids (phone or Teams aadId).
/** Normalize a phone number to digits only. */
export function normalizePhoneNumber(input) {
    if (!input) {
        return "";
    }
    return input.replace(/\D/g, "");
}
/**
 * Return true when the caller matches an allowlist entry — by phone number
 * (digits only) OR by exact caller id, case-insensitive. The id match lets
 * providers whose caller id is not a phone number (e.g. the Microsoft Teams
 * `aadId`, an AAD object id) use the allowlist without being phone-normalized.
 * `from` may be a raw caller id or an already-normalized phone string.
 */
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
        // Exact caller-id match (e.g. a Teams AAD object id), case-insensitive.
        if (trimmed.toLowerCase() === idFrom) {
            return true;
        }
        // Phone-number match (digits only).
        const normalizedAllow = normalizePhoneNumber(trimmed);
        return normalizedAllow !== "" && normalizedFrom !== "" && normalizedAllow === normalizedFrom;
    });
}
/**
 * Inbound-policy decision. `from` is the caller id — a phone number, or a Teams
 * `aadId` for msteams. `isAllowlistedCaller` matches either form.
 * "disabled"/unset → reject (defensive).
 */
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
/**
 * Actionable log line for a caller rejected by the inbound policy.
 *
 * Honesty note: this plugin enforces "pairing" as a plain allowlist - it issues
 * no pairing codes, expirations, or approval prompts for calls - so the fix for
 * a rejected caller is the same as under "allowlist": put them in `allowFrom`
 * (or use the gateway host's own pairing flow, where one exists, to do so).
 */
export function describeInboundRejection(inboundPolicy, from) {
    const policy = inboundPolicy ?? "disabled";
    const caller = from?.trim() ? `caller "${from.trim()}"` : "caller with no caller id";
    if (policy === "disabled" || policy === "open") {
        // "open" never rejects; kept for exhaustiveness if callers reuse this.
        return `inbound call rejected by policy "${policy}": inbound calling is ${policy === "disabled" ? "disabled - set inboundPolicy to \"allowlist\" and add callers to allowFrom to accept calls" : "open"} (${caller})`;
    }
    const pairingHint = policy === "pairing"
        ? ' Note: "pairing" currently enforces a plain allowlist (this plugin issues no pairing codes or approvals for calls) - add the caller\'s AAD object id to allowFrom, or approve them via your gateway\'s pairing flow if your host supports one.'
        : " Add the caller's AAD object id (or phone number) to allowFrom to accept them.";
    return `inbound call rejected by policy "${policy}": ${caller} is not in allowFrom.${pairingHint}`;
}
